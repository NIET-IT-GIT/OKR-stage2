import nodemailer from "nodemailer";

const CORS = {
  "Access-Control-Allow-Origin": "https://okr.nietgroup.com.au",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

const ALLOWED_DOMAINS = ["niet.edu.au", "charltonbrown.edu.au", "educare.edu.au", "rhodes.edu.au"];

export default async function handler(req, res) {
  Object.entries(CORS).forEach(([k, v]) => res.setHeader(k, v));
  res.setHeader("Vary", "Origin");
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method not allowed" }); return; }

  const { to, name, projects = [] } = req.body || {};

  const domain = (typeof to === "string" ? to : "").split("@")[1]?.toLowerCase();
  if (!domain || !ALLOWED_DOMAINS.includes(domain)) {
    res.status(403).json({ error: "Recipient domain not allowed" }); return;
  }
  if (!to || !name || !projects.length) {
    res.status(400).json({ error: "Missing required fields" }); return;
  }

  const smtpUser = process.env.SMTP_USER;
  const smtpPass = process.env.SMTP_PASS;
  const appUrl   = process.env.APP_URL || "https://okr.nietgroup.com.au";

  if (!smtpUser || !smtpPass) {
    console.warn("SMTP_USER / SMTP_PASS not configured — email skipped");
    res.status(200).json({ ok: true, skipped: true }); return;
  }

  const transporter = nodemailer.createTransport({
    host: "smtp.office365.com",
    port: 587,
    secure: false,
    auth: { user: smtpUser, pass: smtpPass },
    tls: { ciphers: "SSLv3" },
  });

  const now = new Date();
  const fmtDate = d => {
    if (!d || d === "TBD") return "TBD";
    const dt = new Date(d);
    return isNaN(dt) ? d : dt.toLocaleDateString("en-AU", { day: "2-digit", month: "short", year: "numeric" });
  };
  const isOverdue = p => {
    if (!p.due || p.due === "TBD") return false;
    const dt = new Date(p.due);
    return !isNaN(dt) && dt < now;
  };

  const rows = projects.map(p => {
    const overdue = isOverdue(p);
    const progressColor = p.progress >= 70 ? "#28CD41" : p.progress >= 35 ? "#FF9500" : "#FF3B30";
    const rowBg = overdue ? "#fff7ed" : "#ffffff";
    return `
    <tr style="background:${rowBg}">
      <td style="padding:10px 14px;border-bottom:1px solid #e5e7eb;font-size:14px;font-weight:600;color:#1d1d1f">
        ${p.name}${overdue ? ' <span style="font-size:11px;font-weight:700;color:#b45309;background:#fef3c7;border:1px solid #fcd34d;border-radius:6px;padding:1px 6px;margin-left:6px">OVERDUE</span>' : ""}
      </td>
      <td style="padding:10px 14px;border-bottom:1px solid #e5e7eb;text-align:center">
        <span style="font-size:13px;font-weight:700;color:${progressColor};font-family:monospace">${p.progress}%</span>
        <div style="margin-top:4px;height:4px;background:#e5e7eb;border-radius:2px;width:80px;display:inline-block;vertical-align:middle;margin-left:8px">
          <div style="height:4px;width:${p.progress}%;background:${progressColor};border-radius:2px"></div>
        </div>
      </td>
      <td style="padding:10px 14px;border-bottom:1px solid #e5e7eb;font-size:13px;color:${overdue ? "#b45309" : "#6e6e73"};text-align:center;white-space:nowrap">${fmtDate(p.due)}</td>
      <td style="padding:10px 14px;border-bottom:1px solid #e5e7eb;font-size:12px;color:#6e6e73;text-align:center">${p.updatedDate || "<em style=\"color:#a1a1aa\">Never</em>"}</td>
    </tr>`;
  }).join("");

  const overdueCount = projects.filter(isOverdue).length;

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f5f5f7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif">
  <div style="max-width:600px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08)">
    <div style="background:linear-gradient(135deg,#0071E3,#6b47dc);padding:28px 32px">
      <div style="color:#fff;font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;opacity:0.8;margin-bottom:6px">NIET Group OKRs System</div>
      <div style="color:#fff;font-size:22px;font-weight:700">Project Status Update Required</div>
    </div>
    <div style="padding:28px 32px">
      <p style="margin:0 0 16px;font-size:15px;color:#1d1d1f">Hi ${name},</p>
      <p style="margin:0 0 20px;font-size:14px;color:#6e6e73;line-height:1.6">
        Please log in to the portal and update the status of your project${projects.length !== 1 ? "s" : ""} listed below.
        ${overdueCount > 0 ? `<strong style="color:#b45309">${overdueCount} project${overdueCount !== 1 ? "s are" : " is"} overdue.</strong>` : ""}
      </p>
      <table style="width:100%;border-collapse:collapse;margin-bottom:24px;border:1px solid #e5e7eb;border-radius:8px;overflow:hidden">
        <thead>
          <tr style="background:#f5f5f7">
            <th style="padding:9px 14px;text-align:left;font-size:11px;font-weight:700;color:#6e6e73;letter-spacing:0.06em;text-transform:uppercase">Project</th>
            <th style="padding:9px 14px;text-align:center;font-size:11px;font-weight:700;color:#6e6e73;letter-spacing:0.06em;text-transform:uppercase">Progress</th>
            <th style="padding:9px 14px;text-align:center;font-size:11px;font-weight:700;color:#6e6e73;letter-spacing:0.06em;text-transform:uppercase">Due Date</th>
            <th style="padding:9px 14px;text-align:center;font-size:11px;font-weight:700;color:#6e6e73;letter-spacing:0.06em;text-transform:uppercase">Last Updated</th>
          </tr>
        </thead>
        <tbody>${rows}</tbody>
      </table>
      <a href="${appUrl}/manager/projects"
         style="display:inline-block;padding:12px 28px;background:#0071e3;color:#fff;text-decoration:none;border-radius:8px;font-size:15px;font-weight:600;letter-spacing:-0.01em">
        Update My Projects →
      </a>
      <p style="margin:20px 0 0;font-size:12px;color:#a1a1aa">
        You are receiving this because you have active projects in the NIET Group OKRs system.<br/>
        Please do not reply to this email.
      </p>
    </div>
  </div>
</body></html>`;

  const plain = [
    `Hi ${name},`,
    "",
    `Please update the status of your active project${projects.length !== 1 ? "s" : ""}:`,
    "",
    ...projects.map(p => `• ${p.name} — ${p.progress}% complete | Due: ${fmtDate(p.due)}${isOverdue(p) ? " [OVERDUE]" : ""} | Last updated: ${p.updatedDate || "Never"}`),
    "",
    `Update your projects at: ${appUrl}/manager/projects`,
    "",
    "You are receiving this because you have active projects in the NIET Group OKRs system.",
    "Please do not reply to this email.",
  ].join("\n");

  try {
    await transporter.sendMail({
      from: `"NIET Group OKRs" <${smtpUser}>`,
      to,
      subject: `Action Required: Please update your project status`,
      text: plain,
      html,
      headers: {
        "List-Unsubscribe": `<mailto:${smtpUser}?subject=Unsubscribe>`,
        "X-Mailer": "NIET Group OKR System",
      },
    });
    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("Project reminder email error:", err.message);
    res.status(500).json({ error: err.message });
  }
}
