import nodemailer from "nodemailer";

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
};

export default async function handler(req, res) {
  Object.entries(CORS).forEach(([k, v]) => res.setHeader(k, v));
  if (req.method === "OPTIONS") { res.status(204).end(); return; }
  if (req.method !== "POST") { res.status(405).json({ error: "Method not allowed" }); return; }

  const { to, name, period, periodKey, krs, template = {} } = req.body || {};
  if (!to || !krs?.length) { res.status(400).json({ error: "Missing required fields" }); return; }

  const smtpUser = process.env.SMTP_USER;
  const smtpPass = process.env.SMTP_PASS;
  const appUrl   = process.env.APP_URL || "https://okr.nietgroup.com.au";

  if (!smtpUser || !smtpPass) {
    console.warn("SMTP_USER / SMTP_PASS not configured — email skipped");
    res.status(200).json({ ok: true, skipped: true });
    return;
  }

  const transporter = nodemailer.createTransport({
    host: "smtp.office365.com",
    port: 587,
    secure: false,
    auth: { user: smtpUser, pass: smtpPass },
    tls: { rejectUnauthorized: false },
  });

  const PERIOD_LABELS = { daily: "Daily", weekly: "Weekly", monthly: "Monthly", quarterly: "Quarterly", annual: "Annual" };
  const periodLabel = PERIOD_LABELS[period] || period || "Monthly";
  const periodLower = periodLabel.toLowerCase();

  const resolveTmpl = (tpl, fallback) => {
    const v = template[tpl] || fallback;
    return v
      .replace(/\{periodLabel\}/g, periodLabel)
      .replace(/\{periodLower\}/g, periodLower)
      .replace(/\{periodKey\}/g,   periodKey || "");
  };

  const fromName  = template.fromName || "NIET Group OKR";
  const subject   = resolveTmpl("subject", `Action Required: ${periodLabel} KPI Check-in — ${periodKey || ""}`);
  const bodyText  = resolveTmpl("body", `Here are your ${periodLower} KPI targets for <strong>${periodKey || ""}</strong>.\nPlease log in to the portal and mark whether you have met each target.`);
  const ctaText   = template.ctaText  || "Submit My Check-in →";
  const footerText = (template.footer || "You are receiving this because you have KPI targets in the NIET Group OKR system.\nPlease do not reply to this email.").replace(/\n/g, "<br/>");

  const krRows = krs.map(kr => `
    <tr>
      <td style="padding:8px 14px;border-bottom:1px solid #e5e7eb;font-size:14px">${kr.label || "—"}</td>
      <td style="padding:8px 14px;border-bottom:1px solid #e5e7eb;text-align:right;font-size:14px;font-family:monospace">
        ${kr.target != null ? kr.target : "—"}${kr.unit ? " " + kr.unit : ""}
      </td>
    </tr>`).join("");

  const html = `<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="margin:0;padding:0;background:#f5f5f7;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif">
  <div style="max-width:560px;margin:32px auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08)">
    <div style="background-color:#0071e3;background:linear-gradient(135deg,#0071e3,#6b47dc);padding:28px 32px">
      <div style="color:#fff;font-size:11px;font-weight:700;letter-spacing:0.1em;text-transform:uppercase;opacity:0.8;margin-bottom:6px">NIET Group OKR System</div>
      <div style="color:#fff;font-size:22px;font-weight:700">${periodLabel} KPI Check-in</div>
    </div>
    <div style="padding:28px 32px">
      <p style="margin:0 0 18px;font-size:15px;color:#1d1d1f">Hi ${name || "there"},</p>
      <p style="margin:0 0 18px;font-size:14px;color:#6e6e73;line-height:1.6">${bodyText.replace(/\n/g, "<br/>")}</p>
      <table style="width:100%;border-collapse:collapse;margin:0 0 24px">
        <thead>
          <tr style="background:#f5f5f7">
            <th style="padding:8px 14px;text-align:left;font-size:11px;font-weight:700;color:#6e6e73;letter-spacing:0.06em;text-transform:uppercase">Key Result</th>
            <th style="padding:8px 14px;text-align:right;font-size:11px;font-weight:700;color:#6e6e73;letter-spacing:0.06em;text-transform:uppercase">Target</th>
          </tr>
        </thead>
        <tbody>${krRows}</tbody>
      </table>
      <a href="${appUrl}/member/checkin"
         style="display:inline-block;padding:12px 28px;background:#0071e3;color:#fff;text-decoration:none;border-radius:8px;font-size:15px;font-weight:600;letter-spacing:-0.01em">
        ${ctaText}
      </a>
      <p style="margin:20px 0 0;font-size:12px;color:#a1a1aa">${footerText}</p>
    </div>
  </div>
</body></html>`;

  try {
    await transporter.sendMail({
      from: `"${fromName}" <${smtpUser}>`,
      to,
      subject,
      html,
    });
    res.status(200).json({ ok: true });
  } catch (err) {
    console.error("Email send error:", err.message, err.code, err.response);
    res.status(500).json({ error: err.message, code: err.code, response: err.response });
  }
}
