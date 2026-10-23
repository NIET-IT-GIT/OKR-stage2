import { requireUser, sendAuthError } from "./_lib/auth.js";

// Reads trainer attendance from the Charlton Brown attendance portals for the Trainer Attendance page and the Pilot.
// HSP/PSP and VET are separate portals with separate API keys; both keys stay on the server and the browser
// only ever talks to this endpoint.
const PORTALS = {
  hsp: { name: "HSP/PSP", url: "https://attendance.charltonbrown.edu.au/hsp-psp/api/kpi-sync/attendance-by-class", env: "HSP_ATTENDANCE_API_KEY" },
  vet: { name: "VET", url: "https://attendance.charltonbrown.edu.au/api/kpi-sync/attendance-by-class", env: "VET_ATTENDANCE_API_KEY" },
};
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const MAX_DAYS = 400; // the portals allow about 13 months

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

  const portalId = String(req.query?.portal || "hsp").toLowerCase();
  const portal = PORTALS[portalId];
  if (!portal) return res.status(400).json({ error: "Unknown attendance portal." });

  const from = String(req.query?.from || ""), to = String(req.query?.to || "");
  const f = validDate(from), t = validDate(to);
  if (!f || !t) return res.status(400).json({ error: "Choose a start and end date." });
  if (t < f) return res.status(400).json({ error: "The end date is before the start date." });
  if ((t - f) / 86400000 > MAX_DAYS) return res.status(400).json({ error: "Choose a period of 13 months or less." });

  const key = process.env[portal.env];
  if (!key) { console.error(`[attendance] ${portal.env} not configured`); return res.status(500).json({ error: `The ${portal.name} attendance connection is not set up yet. Ask IT to add its API key.` }); }

  let upstream;
  try {
    upstream = await fetch(`${portal.url}?from=${from}&to=${to}`, {
      headers: { Authorization: `Bearer ${key}`, Accept: "application/json" },
      signal: AbortSignal.timeout(25000),
    });
  } catch (err) {
    const timedOut = err?.name === "TimeoutError" || err?.name === "AbortError";
    console.error(`[attendance] ${portalId} request failed:`, err?.name, err?.message);
    return res.status(timedOut ? 504 : 502).json({ error: timedOut ? `The ${portal.name} attendance portal took too long to answer. Try again.` : `Couldn't reach the ${portal.name} attendance portal. Try again shortly.` });
  }

  if (upstream.status === 401 || upstream.status === 403) {
    console.error(`[attendance] ${portalId} portal rejected the API key:`, upstream.status);
    return res.status(502).json({ error: `The ${portal.name} attendance portal rejected our API key (missing, wrong or cancelled). Ask IT to check it.` });
  }
  if (upstream.status === 400) return res.status(400).json({ error: `The ${portal.name} attendance portal didn't accept that period. Check the dates and try again.` });
  if (!upstream.ok) {
    console.error(`[attendance] ${portalId} portal error:`, upstream.status);
    return res.status(502).json({ error: `The ${portal.name} attendance portal had an error. Try again in a minute.` });
  }

  let data;
  try { data = await upstream.json(); } catch { return res.status(502).json({ error: `The ${portal.name} attendance portal sent a reply we couldn't read.` }); }
  if (!data || !Array.isArray(data.trainers) || !Array.isArray(data.classes)) {
    return res.status(502).json({ error: `The ${portal.name} attendance portal sent an unexpected reply.` });
  }
  const out = { portal: portalId, from: data.from || from, to: data.to || to, generatedAt: data.generatedAt || null, trainers: data.trainers, classes: data.classes };
  // VET also sends timetable-based figures (not yet in its documentation); passed on so the page can show them as a preview
  if (portalId === "vet" && data.scheduled && Array.isArray(data.scheduled.trainers)) out.scheduled = { trainers: data.scheduled.trainers };
  return res.status(200).json(out);
}
