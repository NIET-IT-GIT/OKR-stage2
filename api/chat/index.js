const https = require("https");
const { getContainer, CORS } = require("../shared/cosmos");

module.exports = async function (context, req) {
  if (req.method === "OPTIONS") { context.res = { status: 204, headers: CORS }; return; }

  const { question, userId, systemPrompt } = req.body || {};
  if (!question) {
    context.res = { status: 400, headers: CORS, body: { error: "missing question" } };
    return;
  }

  const apiKey = process.env.CLAUDE_API_KEY;
  if (!apiKey) {
    context.res = { status: 500, headers: CORS, body: { error: "CLAUDE_API_KEY not configured" } };
    return;
  }

  try {
    // Verify caller is admin
    const { resources: usersAll } = await getContainer("users").items.readAll().fetchAll();
    const caller = usersAll.find(u => u.id === userId);
    if (!caller || caller.role !== "admin") {
      context.res = { status: 403, headers: CORS, body: { error: "forbidden" } };
      return;
    }

    // Fetch all data in parallel
    const [deptsRes, memberDataRes, weeklySubs, monthlyReports] = await Promise.all([
      getContainer("depts").items.readAll().fetchAll(),
      getContainer("memberData").items.readAll().fetchAll(),
      getContainer("weeklySubs").items.readAll().fetchAll(),
      getContainer("monthlyReports").items.readAll().fetchAll(),
    ]);

    const depts = deptsRes.resources;
    const memberDataArr = memberDataRes.resources;
    const subs = weeklySubs.resources; // okrSubmissions stored here
    const reports = monthlyReports.resources;

    const members = usersAll.filter(u => u.role === "member" || u.role === "manager");
    const today = new Date().toISOString().slice(0, 10);
    const monthKey = today.slice(0, 7);

    // Build context summary
    const memberSummary = members.map(u => {
      const kd = memberDataArr.find(d => d.id === u.id) || { krs: [] };
      const mySubs = subs.filter(s => s.memberId === u.id && s.answer !== null);
      const thisMonthSubs = mySubs.filter(s => s.periodKey && s.periodKey.startsWith(monthKey));
      const pendingSubs = subs.filter(s => s.memberId === u.id && s.answer === null);
      const deptName = depts.find(d => d.id === u.deptId)?.name || "—";

      // Completion rate: mean of per-KR rates
      const krRates = kd.krs
        .filter(kr => kr.type !== "tracker")
        .map(kr => {
          const krSubs = mySubs.filter(s => s.krId === kr.id);
          if (!krSubs.length) return null;
          const yesSubs = krSubs.filter(s => s.answer === "yes");
          return (yesSubs.length / krSubs.length) * 100;
        })
        .filter(r => r !== null);
      const rate = krRates.length ? krRates.reduce((a, b) => a + b, 0) / krRates.length : null;

      return `${u.name} (${deptName}, ${u.role}): ` +
        (rate !== null ? `${rate.toFixed(1)}% target met rate` : "no submissions yet") +
        `, ${thisMonthSubs.length} check-ins this month` +
        (pendingSubs.length ? `, ${pendingSubs.length} pending` : "");
    }).join("\n");

    const deptSummary = depts.map(d => {
      const deptMembers = members.filter(u => u.deptId === d.id);
      const rates = deptMembers.map(u => {
        const kd = memberDataArr.find(md => md.id === u.id) || { krs: [] };
        const mySubs = subs.filter(s => s.memberId === u.id && s.answer !== null);
        const krRates = kd.krs
          .filter(kr => kr.type !== "tracker")
          .map(kr => {
            const krSubs = mySubs.filter(s => s.krId === kr.id);
            if (!krSubs.length) return null;
            return (krSubs.filter(s => s.answer === "yes").length / krSubs.length) * 100;
          })
          .filter(r => r !== null);
        return krRates.length ? krRates.reduce((a, b) => a + b, 0) / krRates.length : null;
      }).filter(r => r !== null);
      const avgRate = rates.length ? rates.reduce((a, b) => a + b, 0) / rates.length : null;
      return `${d.name}: ${avgRate !== null ? avgRate.toFixed(1) + "%" : "no data"} (${deptMembers.length} members)`;
    }).join("\n");

    const latestReport = reports.sort((a, b) => (b.publishedDate || "").localeCompare(a.publishedDate || ""))[0];
    const reportSummary = latestReport
      ? `Latest report (${latestReport.month || latestReport.publishedDate}): company rate ${latestReport.data?.companyRate?.toFixed(1)}%`
      : "No published reports yet.";

    const contextText = [
      `[Today: ${today}, FY month: ${monthKey}]`,
      `\nDepartment summary:\n${deptSummary}`,
      `\nMember details:\n${memberSummary}`,
      `\nReport: ${reportSummary}`,
    ].join("\n");

    const resolvedPrompt = systemPrompt || `You are an OKR analytics assistant for NIET (National Institute for Excellence in Teaching).
Answer questions concisely based on the provided OKR data.
Use specific numbers and names.
Reply in the same language the user asks in (Chinese or English).
Do not make up data — only use what's provided.`;

    const answer = await callClaude(apiKey, resolvedPrompt, contextText, question);
    context.res = { status: 200, headers: { ...CORS, "Content-Type": "application/json" }, body: { answer } };

  } catch (err) {
    context.log.error("Chat error:", err);
    context.res = { status: 500, headers: CORS, body: { error: err.message } };
  }
};

function callClaude(apiKey, system, contextText, question) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: "claude-opus-4-5",
      max_tokens: 1024,
      system,
      messages: [{ role: "user", content: `${contextText}\n\n问题：${question}` }],
    });

    const options = {
      hostname: "api.anthropic.com",
      path: "/v1/messages",
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
        "Content-Length": Buffer.byteLength(body),
      },
    };

    const req = https.request(options, res => {
      let data = "";
      res.on("data", chunk => { data += chunk; });
      res.on("end", () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.error) return reject(new Error(parsed.error.message));
          resolve(parsed.content?.[0]?.text || "No response.");
        } catch (e) { reject(e); }
      });
    });
    req.on("error", reject);
    req.write(body);
    req.end();
  });
}
