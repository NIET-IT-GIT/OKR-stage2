import { requireUser, sendAuthError } from "./_lib/auth.js";

// Reads HSP/PSP trainer attendance from the Charlton Brown attendance portal for the Trainer Attendance page.
// The portal's API key stays on the server (HSP_ATTENDANCE_API_KEY); the browser only ever talks to this endpoint.
const UPSTREAM = "https://attendance.charltonbrown.edu.au/hsp-psp/api/kpi-sync/attendance-by-class";
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_DAYS = 400; // the portal allows about 13 months

function validDate(s) {
  if (!DATE_RE.test(s || "")) return null;
  const d = new Date(s + "T00:00:00Z");
  return isNaN(d) || d.toISOString().slice(0, 10) !== s ? null : d;
}

export default async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "https://okr.nietgroup.com.au");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  res.setHeader("Vary", "Origin");
  res.setHeader("Cache-Control", "no-store");
  if (req.method === "OPTIONS") return res.status(204).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  try { await requireUser(req, { roles: ["admin"] }); } catch (err) { return sendAuthError(res, err); }

  const from = String(req.query?.from || ""), to = String(req.query?.to || "");
  const f = validDate(from), t = validDate(to);
  if (!f || !t) return res.status(400).json({ error: "Choose a start and end date." });
  if (t < f) return res.status(400).json({ error: "The end date is before the start date." });
  if ((t - f) / 86400000 > MAX_DAYS) return res.status(400).json({ error: "Choose a period of 13 months or less." });

  const key = process.env.HSP_ATTENDANCE_API_KEY;
  if (!key) { console.error("[attendance] HSP_ATTENDANCE_API_KEY not configured"); return res.status(500).json({ error: "The attendance connection is not set up yet. Ask IT to add the API key." }); }

  let upstream;
  try {
    upstream = await fetch(`${UPSTREAM}?from=${from}&to=${to}`, {
      headers: { Authorization: `Bearer ${key}`, Accept: "application/json" },
      signal: AbortSignal.timeout(25000),
    });
  } catch (err) {
    const timedOut = err?.name === "TimeoutError" || err?.name === "AbortError";
    console.error("[attendance] request failed:", err?.name, err?.message);
    return res.status(timedOut ? 504 : 502).json({ error: timedOut ? "The attendance portal took too long to answer. Try again." : "Couldn't reach the attendance portal. Try again shortly." });
  }

  if (upstream.status === 401 || upstream.status === 403) {
    console.error("[attendance] portal rejected the API key:", upstream.status);
    return res.status(502).json({ error: "The attendance portal rejected our API key (missing, wrong or cancelled). Ask IT to check it." });
  }
  if (upstream.status === 400) return res.status(400).json({ error: "The attendance portal didn't accept that period. Check the dates and try again." });
  if (!upstream.ok) {
    console.error("[attendance] portal error:", upstream.status);
    return res.status(502).json({ error: "The attendance portal had an error. Try again in a minute." });
  }

  let data;
  try { data = await upstream.json(); } catch { return res.status(502).json({ error: "The attendance portal sent a reply we couldn't read." }); }
  if (!data || !Array.isArray(data.trainers) || !Array.isArray(data.classes)) {
    return res.status(502).json({ error: "The attendance portal sent an unexpected reply." });
  }
  return res.status(200).json({ from: data.from || from, to: data.to || to, generatedAt: data.generatedAt || null, trainers: data.trainers, classes: data.classes });
}
