const https = require("https");

module.exports = async (req, res) => {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const { question, systemPrompt, contextData } = req.body || {};
  if (!question) return res.status(400).json({ error: "missing question" });

  const apiKey = process.env.CLAUDE_API_KEY;
  if (!apiKey) return res.status(500).json({ error: "CLAUDE_API_KEY not configured" });

  try {
    const answer = await callClaude(apiKey, systemPrompt, contextData, question);
    res.status(200).json({ answer });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
};

function callClaude(apiKey, systemPrompt, contextData, question) {
  const system = systemPrompt || "You are an OKR analytics assistant. Answer questions based on the provided data. Be concise and use specific numbers. Reply in the same language as the question.";
  const userMessage = contextData ? `${contextData}\n\n问题：${question}` : question;

  return new Promise((resolve, reject) => {
    const body = JSON.stringify({
      model: "claude-opus-4-5",
      max_tokens: 1024,
      system,
      messages: [{ role: "user", content: userMessage }],
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

    const req = https.request(options, r => {
      let data = "";
      r.on("data", chunk => { data += chunk; });
      r.on("end", () => {
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
