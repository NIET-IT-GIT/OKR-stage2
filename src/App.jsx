import { useState, useEffect, useCallback, useRef, Fragment } from "react";
import { useMsal } from "@azure/msal-react";
import { EventType } from "@azure/msal-browser";
import { loginRequest } from "./authConfig";
import { createClient } from "@supabase/supabase-js";

const T = {
  bg: "#F5F5F7", bgSoft: "#FAFAFA", surface: "#FFFFFF", surfaceHover: "#F5F5F7",
  raised: "#F2F2F2", border: "rgba(0,0,0,0.08)", borderFocus: "#0071E3",
  text: "#1D1D1F", textSoft: "#3A3A3C", textMuted: "#6E6E73", textDim: "#AEAEB2",
  brand: "#0071E3", brandSoft: "#0077ED", brandDim: "rgba(0,113,227,0.07)",
  brandBorder: "rgba(0,113,227,0.18)",
  ok: "#28CD41", okDim: "rgba(52,199,89,0.08)", okBorder: "rgba(52,199,89,0.22)",
  warn: "#FF9F0A", warnDim: "rgba(255,159,10,0.08)", warnBorder: "rgba(255,159,10,0.22)",
  bad: "#FF3B30", badDim: "rgba(255,59,48,0.08)", badBorder: "rgba(255,59,48,0.22)",
  orange: "#FF9500", purple: "#BF5AF2",
  glass: "rgba(255,255,255,0.72)",
  glassBorder: "rgba(255,255,255,0.5)",
  shadow: "0 1px 3px rgba(0,0,0,0.06), 0 8px 24px rgba(0,0,0,0.05)",
  shadowSm: "0 1px 2px rgba(0,0,0,0.05), 0 2px 8px rgba(0,0,0,0.04)",
};
const F = { body: "-apple-system,'SF Pro Text',BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif", mono: "'SF Mono','Fira Code','Cascadia Code',monospace" };
const TP = 66.7;

const STATUS_THEME = {
  green:  { bg: T.okDim,   border: T.okBorder,   color: T.ok,   tag: "On Track" },
  yellow: { bg: T.warnDim, border: T.warnBorder,  color: T.warn, tag: "At Risk"  },
  red:    { bg: T.badDim,  border: T.badBorder,   color: T.bad,  tag: "Behind"   },
};
const APPROVAL = {
  pending:  { bg: T.warnDim, border: T.warnBorder, color: T.warn, label: "Pending Review" },
  approved: { bg: T.okDim,   border: T.okBorder,   color: T.ok,   label: "Approved"       },
  rejected: { bg: T.badDim,  border: T.badBorder,  color: T.bad,  label: "Rejected"       },
};

/* ─── MOCK DATA ─── */
const INIT_USERS = [
  { id: "sysadmin", name: "System Admin",    email: "__admin__",                           role: "admin",   av: "SA", title: "System Administrator" },
  { id: "admin1",   name: "Troy Yue",        email: "troy.yue@niet.edu.au",                role: "admin",   av: "TY", title: "CEO" },
  { id: "mgr1",     name: "Florence Fan",    email: "florence.fan@niet.edu.au",            role: "manager", av: "FF", title: "Head of Admissions",  deptId: "admissions", teamIds: ["domestic","international"] },
  { id: "mgr2",     name: "Olivia An",       email: "olivia.an@niet.edu.au",               role: "manager", av: "OA", title: "Head of Marketing",   deptId: "marketing",  teamIds: ["digital"] },
  { id: "mem1",     name: "Grace He",        email: "gracie.he@charltonbrown.edu.au",      role: "member",  av: "GH", title: "Admissions Officer",  teamId: "domestic",      deptId: "admissions", mgrId: "mgr1" },
  { id: "mem2",     name: "Mary Joy Caraig", email: "mary.joy.caraig@charltonbrown.edu.au",role: "member",  av: "MJ", title: "Admissions Officer",  teamId: "domestic",      deptId: "admissions", mgrId: "mgr1" },
  { id: "mem3",     name: "Tom Walker",      email: "tom@niet.edu.au",                     role: "member",  av: "TW", title: "Junior Officer",       teamId: "domestic",      deptId: "admissions", mgrId: "mgr1" },
  { id: "mem4",     name: "Amy Zhang",       email: "amy@niet.edu.au",                     role: "member",  av: "AZ", title: "Intl Officer",         teamId: "international", deptId: "admissions", mgrId: "mgr1" },
  { id: "mem5",     name: "Emma Wilson",     email: "emma@charltonbrown.edu.au",           role: "member",  av: "EW", title: "Digital Specialist",   teamId: "digital",       deptId: "marketing",  mgrId: "mgr2" },
  { id: "sam",      name: "Samuel Zhong",    email: "samuel.zhong@niet.edu.au",            role: "admin",   av: "SZ", title: "IT Administrator" },
];

const INIT_DEPTS = [
  { id: "admissions", name: "Admissions", head: "Florence Fan", college: "NIET",
    obj: "Drive enrolment targets for FY26 Q1",
    krs: [
      { id: "AKR1", label: "New domestic enrolments",     target: 120, actual: 95 },
      { id: "AKR2", label: "New international enrolments", target: 80,  actual: 62 },
      { id: "AKR3", label: "Conversion rate (%)",          target: 45,  actual: 38 },
    ],
    teams: [
      { id: "domestic", name: "Domestic Team", lead: "Florence Fan", obj: "Hit domestic enrolment KPIs",
        krs: [{ id: "DTK1", label: "Domestic enrolments closed", target: 60, actual: 52 }, { id: "DTK2", label: "Weekly follow-ups", target: 25, actual: 22 }],
        members: ["mem1","mem2","mem3"] },
      { id: "international", name: "International Team", lead: "Amy Zhang", obj: "Hit intl enrolment KPIs",
        krs: [{ id: "ITK1", label: "International enrolments", target: 40, actual: 31 }, { id: "ITK2", label: "Agent meetings/mo", target: 12, actual: 10 }],
        members: ["mem4"] },
    ] },
  { id: "marketing", name: "Marketing", head: "Olivia An", college: "NIET",
    obj: "Build brand awareness & lead generation",
    krs: [
      { id: "MKR1", label: "Qualified leads",       target: 500, actual: 420 },
      { id: "MKR2", label: "Social engagement (%)", target: 5,   actual: 4.2 },
    ],
    teams: [
      { id: "digital", name: "Digital Marketing", lead: "Nick Egan", obj: "Drive online leads",
        krs: [{ id: "DMK1", label: "Monthly digital leads", target: 200, actual: 175 }],
        members: ["mem5"] },
    ] },
  { id: "services", name: "Student Services", head: "Mark Thompson", college: "Rhodes",
    obj: "Maximise satisfaction & retention",
    krs: [{ id: "SKR1", label: "Student NPS", target: 80, actual: 74 }, { id: "SKR2", label: "Ticket resolution <24h (%)", target: 90, actual: 82 }],
    teams: [] },
  { id: "it", name: "IT & Systems", head: "Samuel Zhong", college: "NIET Group",
    obj: "Deliver OKR platform & system excellence",
    krs: [{ id: "IKR1", label: "OKR system delivery (%)", target: 100, actual: 65 }, { id: "IKR2", label: "System uptime (%)", target: 99.5, actual: 99.2 }],
    teams: [] },
];

const INIT_MEMBER_DATA = {
  mem1: { krs: [{ id: "GH1", label: "Monthly enrolments",      target: 8,   actual: 7   }, { id: "GH2", label: "Callback rate (%)",     target: 100, actual: 95  }, { id: "GH3", label: "Lead response <2h (%)", target: 95, actual: 88 }] },
  mem2: { krs: [{ id: "MJ1", label: "Monthly enrolments",      target: 8,   actual: 9   }, { id: "MJ2", label: "Callback rate (%)",     target: 100, actual: 100 }, { id: "MJ3", label: "Lead response <2h (%)", target: 95, actual: 97 }] },
  mem3: { krs: [{ id: "TW1", label: "Monthly enrolments",      target: 6,   actual: 3   }, { id: "TW2", label: "Callback rate (%)",     target: 100, actual: 72  }, { id: "TW3", label: "Lead response <2h (%)", target: 95, actual: 65 }] },
  mem4: { krs: [{ id: "AZ1", label: "Monthly intl enrolments", target: 5,   actual: 4   }, { id: "AZ2", label: "Visa success rate (%)", target: 90,  actual: 88  }] },
  mem5: { krs: [{ id: "EW1", label: "Campaign leads/mo",       target: 50,  actual: 48  }, { id: "EW2", label: "Ad spend ROI (%)",      target: 300, actual: 280 }] },
};

const INIT_WEEKLY_SUBS = [
  { id: "ws1", memberId: "mem1", week: "Wk 15 · Apr 13-19", items: "Closed 2 enrolments, 18 follow-up calls, updated CRM for 12 leads, attended open day",    date: "2026-04-14", approval: "approved", mgrNote: "Good work, keep push on callback rate." },
  { id: "ws2", memberId: "mem1", week: "Wk 14 · Apr 6-12",  items: "Closed 1 enrolment, 15 follow-ups, prepared open day materials",                           date: "2026-04-07", approval: "approved", mgrNote: "" },
  { id: "ws3", memberId: "mem2", week: "Wk 15 · Apr 13-19", items: "Closed 3 enrolments, launched referral campaign, 22 follow-ups, trained new agent partner", date: "2026-04-14", approval: "approved", mgrNote: "Excellent — top performer this week." },
  { id: "ws4", memberId: "mem3", week: "Wk 15 · Apr 13-19", items: "8 follow-up calls, attended product training, updated 5 lead records",                     date: "2026-04-15", approval: "pending",  mgrNote: "" },
  { id: "ws5", memberId: "mem4", week: "Wk 15 · Apr 13-19", items: "Processed 3 visa applications, met with 2 agents, follow-up with 10 prospective students",  date: "2026-04-14", approval: "pending",  mgrNote: "" },
  { id: "ws6", memberId: "mem5", week: "Wk 15 · Apr 13-19", items: "Launched 2 Google Ads campaigns, published 3 blog posts, social content for Instagram",     date: "2026-04-14", approval: "pending",  mgrNote: "" },
];

const INIT_MGR_SPRINTS = [
  { id: "ms1", mgrId: "mgr1", week: "Wk 15 · Apr 13-19", summary: "Domestic team: 5 enrolments closed (target pace). International pipeline growing but visa delays affecting 3 applicants. Open day was a success — 40 attendees, 12 qualified leads. Tom Walker underperforming; scheduled 1-on-1 for Monday.", date: "2026-04-14", status: "submitted" },
  { id: "ms2", mgrId: "mgr1", week: "Wk 14 · Apr 6-12",  summary: "Domestic: 3 enrolments, callback rate improving. Started new referral program with Mary leading. International: agent meeting in Southport went well.", date: "2026-04-07", status: "submitted" },
];

const INIT_PROJECTS = [
  { id: "p1", mgrId: "mgr1", name: "2026 Open Day Campaign",     status: "active", due: "2026-05-30", progress: 45 },
  { id: "p2", mgrId: "mgr1", name: "Agent Partnership Expansion", status: "active", due: "2026-06-15", progress: 20 },
];

const INIT_MONTHLY_REPORTS = [
  { id: "mr1", month: "March 2026", publishedDate: "2026-04-01", publishedBy: "admin1", data: {
    companyRate: 72.8,
    deptRanks: [
      { name: "Marketing",        rate: 81.2, status: "green"  },
      { name: "Admissions",       rate: 76.5, status: "yellow" },
      { name: "Student Services", rate: 71.3, status: "yellow" },
      { name: "IT & Systems",     rate: 62.1, status: "yellow" },
    ],
    topPerformers: ["Mary Joy Caraig — 100%", "Grace He — 93.3%", "Emma Wilson — 92.7%"],
    redFlags: ["Tom Walker — 46.7% (action plan required)", "IT System Delivery — 65% (behind schedule)"],
  }},
];

/* ─── HELPERS ─── */
function krCompletion(kr) {
  const op = kr.operator || ">=";
  let actual, target;
  if (kr.monthlyTargets) {
    const key = currentFYMonthKey();
    target = Number(kr.monthlyTargets[key]) || 0;
    actual = Number((kr.monthlyActuals || {})[key]) || 0;
  } else {
    actual = Number(kr.actual) || 0;
    target = Number(kr.target) || 0;
  }
  if (target === 0) return 100;
  switch (op) {
    case ">=": return Math.min((actual / target) * 100, 100);
    case ">":  return actual > target ? 100 : Math.min((actual / target) * 100, 100);
    case "<=": return actual <= target ? 100 : Math.min((target / actual) * 100, 100);
    case "<":  return actual < target ? 100 : Math.min((target / actual) * 100, 100);
    case "=":  return actual === target ? 100 : 0;
    default:   return Math.min((actual / target) * 100, 100);
  }
}
function calcRate(krs) {
  const scored = (krs || []).filter(kr => kr.type !== "tracker");
  if (!scored.length) return 0;
  return scored.reduce((sum, kr) => sum + krCompletion(kr), 0) / scored.length;
}
function calcMemberRate(memberId, memberKrs, okrSubs) {
  const now = Date.now();
  const memberSubs = (okrSubs || []).filter(s => s.memberId === memberId);
  const scores = [];
  for (const kr of (memberKrs || [])) {
    if (kr.type === "tracker") continue;
    const krSubs = memberSubs.filter(s => s.krId === kr.id);
    if (!krSubs.length) continue; // no check-in sent → excluded
    const answered = krSubs.filter(s => s.answer !== null);
    if (answered.length) { scores.push(krCompletion(kr)); continue; } // answered → use real completion
    const latest = krSubs.filter(s => s.sentAt).sort((a, b) => b.sentAt.localeCompare(a.sentAt))[0];
    if (latest && (now - new Date(latest.sentAt).getTime()) >= 86400000) scores.push(0); // overdue → 0%
    // else < 24h grace period → excluded
  }
  if (!scores.length) return 0;
  return scores.reduce((a, b) => a + b, 0) / scores.length;
}
function getStatus(r) { return r >= TP ? "green" : r >= 60 ? "yellow" : "red"; }
function fmt(v) { return typeof v === "number" ? (v % 1 ? v.toFixed(1) : v.toLocaleString()) : v; }
function currentFYQuarter() {
  const d = new Date();
  const m = d.getMonth() + 1;
  const y = d.getFullYear();
  const fy = m >= 7 ? y + 1 : y;
  const q = m >= 7 && m <= 9 ? 1 : m >= 10 ? 2 : m <= 3 ? 3 : 4;
  return `FY${String(fy).slice(2)} Q${q}`;
}
function currentFYHalf() {
  const d = new Date();
  const m = d.getMonth() + 1;
  const y = d.getFullYear();
  const fy = m >= 7 ? y + 1 : y;
  const h = m >= 7 ? 1 : 2;
  return `FY${String(fy).slice(2)} H${h}`;
}
function currentPeriodKey(period) {
  const d = new Date();
  if (period === "daily") return d.toISOString().slice(0, 10);
  if (period === "weekly") return currentFYWeek();
  if (period === "monthly") return currentFYMonthKey();
  if (period === "quarterly") return currentFYQuarter();
  if (period === "biannual") return currentFYHalf();
  if (period === "annual") return String(d.getFullYear());
  return d.toISOString().slice(0, 10);
}
function periodDisplayLabel(period, key) {
  if (!key) return "—";
  if (period === "monthly") { const [y, m] = key.split("-"); return new Date(+y, +m - 1).toLocaleString("en", { month: "short", year: "numeric" }); }
  if (period === "annual") return `FY ${key}`;
  return key;
}
function calcSubmissionRate(okrSubs, memberId, monthKey) {
  const relevant = okrSubs.filter(s => s.memberId === memberId && s.answer !== null && (s.periodKey || "").slice(0, 7) === monthKey);
  if (!relevant.length) return null;
  return (relevant.filter(s => s.answer === "yes").length / relevant.length) * 100;
}
function currentWeekLabel() {
  const d = new Date();
  const w = Math.ceil(((d - new Date(d.getFullYear(), 0, 1)) / 86400000 + 1) / 7);
  const mon = new Date(d); mon.setDate(d.getDate() - d.getDay() + 1);
  const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
  return `Wk ${w} · ${mon.toLocaleDateString("en-AU", { month: "short", day: "numeric" })}-${sun.toLocaleDateString("en-AU", { month: "short", day: "numeric" })}`;
}
function currentMonth() {
  return new Date().toLocaleDateString("en-AU", { month: "long", year: "numeric" });
}
function getFYWeeks() {
  const now = new Date();
  const m = now.getMonth() + 1;
  const y = now.getFullYear();
  const fy = m >= 7 ? y : y - 1;
  const fyStart = new Date(fy, 6, 1);
  const day = fyStart.getDay();
  fyStart.setDate(fyStart.getDate() + (day === 0 ? 1 : day === 1 ? 0 : 8 - day));
  const weeks = [];
  for (let i = 0; i < 52; i++) {
    const mon = new Date(fyStart); mon.setDate(fyStart.getDate() + i * 7);
    const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
    weeks.push(`Wk ${i + 1} · ${mon.toLocaleDateString("en-AU", { month: "short", day: "numeric" })}-${sun.toLocaleDateString("en-AU", { month: "short", day: "numeric" })}`);
  }
  return weeks;
}
function currentFYWeek() {
  const weeks = getFYWeeks();
  const now = new Date();
  const m = now.getMonth() + 1;
  const y = now.getFullYear();
  const fy = m >= 7 ? y : y - 1;
  const fyStart = new Date(fy, 6, 1);
  const day = fyStart.getDay();
  fyStart.setDate(fyStart.getDate() + (day === 0 ? 1 : day === 1 ? 0 : 8 - day));
  const idx = Math.floor((now - fyStart) / (7 * 86400000));
  return weeks[Math.max(0, Math.min(idx, 51))];
}
function getFYMonths() {
  const now = new Date();
  const m = now.getMonth() + 1;
  const y = now.getFullYear();
  // June is treated as the transition month: show upcoming FY Jul-Jun
  const fy = m >= 6 ? y : y - 1;
  const months = [];
  for (let i = 7; i <= 12; i++) months.push({ key: `${fy}-${String(i).padStart(2,"0")}`, label: new Date(fy, i-1).toLocaleDateString("en-AU",{month:"short",year:"numeric"}) });
  for (let i = 1; i <= 6; i++) months.push({ key: `${fy+1}-${String(i).padStart(2,"0")}`, label: new Date(fy+1, i-1).toLocaleDateString("en-AU",{month:"short",year:"numeric"}) });
  return months;
}
function currentFYMonthKey() {
  const n = new Date();
  const key = `${n.getFullYear()}-${String(n.getMonth()+1).padStart(2,"0")}`;
  const fyKeys = getFYMonths().map(m => m.key);
  // If current calendar month is outside the active FY (e.g. June transition), use first FY month
  return fyKeys.includes(key) ? key : fyKeys[0];
}
function prevPeriodKey(period) {
  const now = new Date();
  if (period === "daily") return new Date(now.getTime() - 86400000).toISOString().slice(0, 10);
  if (period === "weekly") {
    const weeks = getFYWeeks();
    const m = now.getMonth() + 1, y = now.getFullYear();
    const fy = m >= 7 ? y : y - 1;
    const fyStart = new Date(fy, 6, 1);
    const d = fyStart.getDay();
    fyStart.setDate(fyStart.getDate() + (d === 0 ? 1 : d === 1 ? 0 : 8 - d));
    const idx = Math.floor((now - fyStart) / (7 * 86400000));
    return weeks[Math.max(0, idx - 1)];
  }
  if (period === "monthly") {
    const d = new Date(now.getFullYear(), now.getMonth() - 1, 1);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  }
  if (period === "quarterly") {
    const m = now.getMonth() + 1, y = now.getFullYear();
    const fy = m >= 7 ? y + 1 : y;
    const q = m >= 7 && m <= 9 ? 1 : m >= 10 ? 2 : m <= 3 ? 3 : 4;
    return q === 1 ? `FY${String(fy - 1).slice(2)} Q4` : `FY${String(fy).slice(2)} Q${q - 1}`;
  }
  if (period === "biannual") {
    const m = now.getMonth() + 1, y = now.getFullYear();
    const fy = m >= 7 ? y + 1 : y;
    const h = m >= 7 ? 1 : 2;
    return h === 1 ? `FY${String(fy - 1).slice(2)} H2` : `FY${String(fy).slice(2)} H1`;
  }
  if (period === "annual") return String(now.getFullYear() - 1);
  return currentPeriodKey(period);
}
function periodDateRange(period, periodKey) {
  if (!periodKey) return "";
  const fmt = d => d.toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" });
  if (period === "daily") return fmt(new Date(periodKey + "T00:00:00"));
  if (period === "weekly") {
    const wm = periodKey.match(/Wk\s*(\d+)/); if (!wm) return periodKey;
    const wk = parseInt(wm[1]);
    const now2 = new Date(), mo = now2.getMonth() + 1, yr = now2.getFullYear();
    const fy = mo >= 7 ? yr : yr - 1;
    const fyS = new Date(fy, 6, 1); const dw = fyS.getDay();
    fyS.setDate(fyS.getDate() + (dw === 0 ? 1 : dw === 1 ? 0 : 8 - dw));
    const mon = new Date(fyS); mon.setDate(fyS.getDate() + (wk - 1) * 7);
    const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
    return `${fmt(mon)} – ${fmt(sun)}`;
  }
  if (period === "monthly") {
    const [y, mo] = periodKey.split("-").map(Number);
    return `${fmt(new Date(y, mo - 1, 1))} – ${fmt(new Date(y, mo, 0))}`;
  }
  if (period === "quarterly") {
    const r = periodKey.match(/FY(\d+)\s*Q(\d)/); if (!r) return periodKey;
    const fy = 2000 + +r[1], q = +r[2];
    const ranges = [[new Date(fy-1,6,1),new Date(fy-1,8,30)],[new Date(fy-1,9,1),new Date(fy-1,11,31)],[new Date(fy,0,1),new Date(fy,2,31)],[new Date(fy,3,1),new Date(fy,5,30)]];
    return `${fmt(ranges[q-1][0])} – ${fmt(ranges[q-1][1])}`;
  }
  if (period === "biannual") {
    const r = periodKey.match(/FY(\d+)\s*H(\d)/); if (!r) return periodKey;
    const fy = 2000 + +r[1], h = +r[2];
    return h === 1 ? `${fmt(new Date(fy-1,6,1))} – ${fmt(new Date(fy-1,11,31))}` : `${fmt(new Date(fy,0,1))} – ${fmt(new Date(fy,5,30))}`;
  }
  if (period === "annual") { const y = +periodKey; return `${fmt(new Date(y,0,1))} – ${fmt(new Date(y,11,31))}`; }
  return periodKey;
}
function makeAv(name) {
  return (name || "?").trim().split(/\s+/).map(w => w[0]).join("").slice(0, 2).toUpperCase();
}

/* ─── DB API (Supabase) ─── */
const supabase = createClient(
  import.meta.env.VITE_SUPABASE_URL,
  import.meta.env.VITE_SUPABASE_ANON_KEY
);

async function dbGet() {
  const { data, error } = await supabase.from("app_data").select("collection, id, doc");
  if (error) throw new Error(error.message);
  const result = { users: [], depts: [], memberData: [], weeklySubs: [], mgrSprints: [], projects: [], monthlyReports: [], okrSubmissions: [], settings: [] };
  for (const row of data) {
    if (result[row.collection]) result[row.collection].push(row.doc);
  }
  return result;
}

async function dbUpsert(collection, item) {
  const { error } = await supabase
    .from("app_data")
    .upsert({ collection, id: item.id, doc: item }, { onConflict: "collection,id" });
  if (error) throw new Error(`dbUpsert(${collection}/${item.id}): ${error.message}`);
}

async function dbDelete(collection, id) {
  const { error } = await supabase.from("app_data").delete().eq("collection", collection).eq("id", id);
  if (error) throw new Error(error.message);
}

async function dbSeed(data) {
  const rows = [];
  for (const [collection, items] of Object.entries(data)) {
    for (const item of items) rows.push({ collection, id: item.id, doc: item });
  }
  const { error } = await supabase.from("app_data").upsert(rows);
  if (error) throw new Error(error.message);
}

// Diffs prev vs next state and syncs only changed items to Supabase.
async function syncChanges(prev, next) {
  const tasks = [];
  if (JSON.stringify(prev.settings) !== JSON.stringify(next.settings)) {
    tasks.push(dbUpsert("settings", next.settings));
  }
  for (const col of ["users", "depts", "weeklySubs", "mgrSprints", "projects", "monthlyReports", "okrSubmissions"]) {
    const prevArr = prev[col] || [];
    const nextArr = next[col] || [];
    for (const item of nextArr) {
      const old = prevArr.find(x => x.id === item.id);
      if (!old || JSON.stringify(old) !== JSON.stringify(item)) tasks.push(dbUpsert(col, item));
    }
    for (const item of prevArr) {
      if (!nextArr.find(x => x.id === item.id)) tasks.push(dbDelete(col, item.id));
    }
  }
  // memberData is keyed by memberId in state
  const prevMd = prev.memberData || {};
  const nextMd = next.memberData || {};
  for (const [id, data] of Object.entries(nextMd)) {
    const old = prevMd[id];
    if (!old || JSON.stringify(old) !== JSON.stringify(data)) tasks.push(dbUpsert("memberData", { id, ...data }));
  }
  for (const id of Object.keys(prevMd)) {
    if (!nextMd[id]) tasks.push(dbDelete("memberData", id));
  }
  await Promise.all(tasks);
}

/* ─── UI PRIMITIVES ─── */
function Tag({ type = "green", label, small }) {
  const s = STATUS_THEME[type] || APPROVAL[type] || STATUS_THEME.green;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 4,
      background: s.bg, border: `1px solid ${s.border}`, borderRadius: 20,
      padding: small ? "2px 8px" : "3px 10px",
      fontSize: small ? 9 : 10, fontWeight: 600, color: s.color, letterSpacing: "0.02em", whiteSpace: "nowrap",
    }}>
      <span style={{ width: 4, height: 4, borderRadius: "50%", background: s.color, flexShrink: 0 }} />
      {label || s.tag || s.label}
    </span>
  );
}

function RoleTag({ role }) {
  const cfg = {
    admin:   { color: T.brand,  bg: T.brandDim,  border: T.brandBorder, label: "Admin"   },
    manager: { color: T.orange, bg: T.warnDim,   border: T.warnBorder,  label: "Manager" },
    member:  { color: T.ok,     bg: T.okDim,     border: T.okBorder,    label: "Member"  },
  }[role] || { color: T.textMuted, bg: T.raised, border: T.border, label: role };
  return (
    <span style={{ display: "inline-flex", alignItems: "center", background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 20, padding: "2px 9px", fontSize: 11, fontWeight: 600, color: cfg.color, whiteSpace: "nowrap", letterSpacing: "0.02em" }}>
      {cfg.label}
    </span>
  );
}

function Bar({ value, status, h = 6 }) {
  const c = STATUS_THEME[status]?.color || T.brand;
  return (
    <div style={{ flex: 1, height: h, background: "rgba(0,0,0,0.06)", borderRadius: h, overflow: "hidden", position: "relative" }}>
      <div style={{ width: `${Math.min(value, 100)}%`, height: "100%", background: c, borderRadius: h, transition: "width 0.6s cubic-bezier(0.4,0,0.2,1)" }} />
      <div style={{ position: "absolute", left: `${TP}%`, top: 0, bottom: 0, width: 1, background: T.textDim, opacity: 0.5 }} />
    </div>
  );
}

function Metric({ label, value, sub, status }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 14, padding: "18px 20px", flex: 1, minWidth: 130, boxShadow: T.shadowSm }}>
      <div style={{ fontSize: 12, color: T.textMuted, fontWeight: 500, letterSpacing: "0.04em", textTransform: "uppercase", marginBottom: 10 }}>{label}</div>
      <div style={{ fontSize: 28, fontWeight: 700, color: status ? STATUS_THEME[status]?.color : T.text, fontFamily: F.mono, letterSpacing: "-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: T.textDim, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

function Btn({ children, primary, danger, small, disabled, onClick, style: sx }) {
  return (
    <button onClick={disabled ? undefined : onClick} style={{
      background: primary ? T.brand : danger ? T.bad : T.surface,
      color: primary || danger ? "#fff" : T.textSoft,
      border: primary || danger ? "none" : `1px solid ${T.border}`,
      borderRadius: 9, padding: small ? "6px 13px" : "9px 18px",
      fontSize: small ? 11 : 12, fontWeight: 500, cursor: disabled ? "not-allowed" : "pointer",
      opacity: disabled ? 0.35 : 1, transition: "all 0.15s", fontFamily: F.body,
      boxShadow: primary ? "0 1px 4px rgba(0,113,227,0.22)" : danger ? "0 1px 4px rgba(255,59,48,0.18)" : T.shadowSm,
      letterSpacing: "-0.01em", ...sx,
    }}>{children}</button>
  );
}

function Input({ value, onChange, placeholder, type, style: sx, ...props }) {
  return (
    <input type={type || "text"} value={value} onChange={onChange} placeholder={placeholder} {...props}
      style={{
        background: T.surface, border: `1px solid ${T.border}`, borderRadius: 9,
        padding: "9px 13px", color: T.text, fontSize: 15, fontFamily: F.body, outline: "none",
        transition: "border-color 0.15s, box-shadow 0.15s", boxSizing: "border-box",
        boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)", ...sx,
      }}
      onFocus={e => { e.target.style.borderColor = T.brand; e.target.style.boxShadow = `0 0 0 3px rgba(0,113,227,0.12), inset 0 1px 2px rgba(0,0,0,0.04)`; }}
      onBlur={e => { e.target.style.borderColor = T.border; e.target.style.boxShadow = "inset 0 1px 2px rgba(0,0,0,0.04)"; }}
    />
  );
}

function Select({ value, onChange, children, style: sx }) {
  return (
    <select value={value} onChange={onChange}
      style={{
        background: T.surface, border: `1px solid ${T.border}`, borderRadius: 9,
        padding: "9px 13px", color: value ? T.text : T.textMuted,
        fontSize: 15, fontFamily: F.body, outline: "none", cursor: "pointer",
        boxSizing: "border-box", boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)", ...sx,
      }}
      onFocus={e => e.target.style.borderColor = T.brand}
      onBlur={e => e.target.style.borderColor = T.border}
    >{children}</select>
  );
}

function TextArea({ value, onChange, placeholder, rows = 4 }) {
  return (
    <textarea value={value} onChange={onChange} placeholder={placeholder} rows={rows}
      style={{
        width: "100%", boxSizing: "border-box", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 9,
        padding: "10px 13px", color: T.text, fontSize: 15, fontFamily: F.body, outline: "none", resize: "vertical",
        transition: "border-color 0.15s, box-shadow 0.15s", boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)",
      }}
      onFocus={e => { e.target.style.borderColor = T.brand; e.target.style.boxShadow = `0 0 0 3px rgba(0,113,227,0.12), inset 0 1px 2px rgba(0,0,0,0.04)`; }}
      onBlur={e => { e.target.style.borderColor = T.border; e.target.style.boxShadow = "inset 0 1px 2px rgba(0,0,0,0.04)"; }}
    />
  );
}

function Avatar({ letters, size = 32, color }) {
  return (
    <div style={{
      width: size, height: size, borderRadius: "50%", flexShrink: 0,
      background: color || `linear-gradient(145deg, ${T.brandSoft}, ${T.brand})`,
      display: "flex", alignItems: "center", justifyContent: "center",
      fontSize: size * 0.33, fontWeight: 600, color: "#fff", letterSpacing: "0.01em",
      boxShadow: "0 1px 3px rgba(0,0,0,0.12)",
    }}>{letters}</div>
  );
}

function Card({ children, style: sx, onClick }) {
  return (
    <div onClick={onClick} style={{
      background: T.surface, border: `1px solid ${T.border}`, borderRadius: 14,
      boxShadow: T.shadowSm, ...sx,
      cursor: onClick ? "pointer" : "default", transition: "box-shadow 0.18s, border-color 0.18s",
    }}
      onMouseEnter={onClick ? e => { e.currentTarget.style.boxShadow = T.shadow; e.currentTarget.style.borderColor = "rgba(0,0,0,0.12)"; } : undefined}
      onMouseLeave={onClick ? e => { e.currentTarget.style.boxShadow = T.shadowSm; e.currentTarget.style.borderColor = T.border; } : undefined}
    >{children}</div>
  );
}

function SectionLabel({ children, style: sx }) {
  return <div style={{ fontSize: 12, fontWeight: 600, color: T.textMuted, letterSpacing: "0.05em", textTransform: "uppercase", marginBottom: 10, ...sx }}>{children}</div>;
}

function EmptyState({ text }) {
  return <div style={{ padding: "48px 20px", textAlign: "center", color: T.textDim, fontSize: 15, fontWeight: 400 }}>{text}</div>;
}

function CountBadge({ count, color }) {
  if (!count) return null;
  return <span style={{ background: color || T.bad, color: "#fff", borderRadius: 20, padding: "1px 6px", fontSize: 11, fontWeight: 700, marginLeft: 6, letterSpacing: "0.02em" }}>{count}</span>;
}

function Side({ items, active, onSelect, user, onLogout, pendingCounts, subItems, subItemsFor, activeSubItem, onSelectSubItem }) {
  return (
    <div style={{
      width: 240, background: T.glass, borderRight: `1px solid ${T.border}`,
      backdropFilter: "saturate(180%) blur(20px)", WebkitBackdropFilter: "saturate(180%) blur(20px)",
      display: "flex", flexDirection: "column", height: "100vh", flexShrink: 0,
    }}>
      <div style={{ padding: "22px 16px 14px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
          <div style={{ width: 40, height: 40, borderRadius: 8, background: `linear-gradient(145deg, ${T.brand}, #A78BFA)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, color: "#fff", boxShadow: "0 2px 8px rgba(0,113,227,0.28)" }}>NIET</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, color: T.text, letterSpacing: "-0.01em" }}>NIET Group OKR's</div>
            <div style={{ fontSize: 11, color: T.textDim, letterSpacing: "0.06em", textTransform: "uppercase" }}>NIET GROUP</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 11px", background: "rgba(0,0,0,0.03)", borderRadius: 11, border: `1px solid ${T.border}` }}>
          <Avatar letters={user.av} size={32} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 14, fontWeight: 600, color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", letterSpacing: "-0.01em" }}>{user.name}</div>
            <div style={{ fontSize: 12, color: T.textMuted }}>{user.title}</div>
          </div>
        </div>
      </div>
      <div style={{ flex: 1, padding: "10px 8px", display: "flex", flexDirection: "column", gap: 1, overflowY: "auto" }}>
        {items.map(item => (
          <div key={item.id}>
            <button onClick={() => onSelect(item.id)} style={{
              background: active === item.id ? T.brandDim : "transparent",
              border: active === item.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
              borderRadius: 9, padding: "9px 12px", cursor: "pointer",
              display: "flex", alignItems: "center", gap: 9,
              color: active === item.id ? T.brand : T.textMuted,
              fontSize: 14, fontWeight: active === item.id ? 600 : 400, textAlign: "left", width: "100%",
              transition: "all 0.12s", fontFamily: F.body, letterSpacing: "-0.01em",
            }}>
              <span style={{ fontSize: 15, width: 18, textAlign: "center", flexShrink: 0, opacity: active === item.id ? 1 : 0.6 }}>{item.icon}</span>
              <span style={{ flex: 1 }}>{item.label}</span>
              {item.id === subItemsFor && subItems && <span style={{ fontSize: 10, opacity: 0.6 }}>{active === item.id ? "▾" : "▸"}</span>}
              {pendingCounts?.[item.id] > 0 && <CountBadge count={pendingCounts[item.id]} />}
            </button>
            {active === item.id && item.id === subItemsFor && subItems && (
              <div style={{ paddingLeft: 10, marginTop: 2, display: "flex", flexDirection: "column", gap: 1 }}>
                {subItems.map(sub => (
                  <button key={sub.id} onClick={() => onSelectSubItem(sub.id)} style={{
                    background: activeSubItem === sub.id ? T.brandDim : "transparent",
                    border: activeSubItem === sub.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
                    borderRadius: 7, padding: "7px 10px", cursor: "pointer",
                    display: "flex", alignItems: "center", gap: 7,
                    color: activeSubItem === sub.id ? T.brand : T.textMuted,
                    fontSize: 13, fontWeight: activeSubItem === sub.id ? 600 : 400, textAlign: "left", width: "100%",
                    fontFamily: F.body, letterSpacing: "-0.01em",
                  }}>
                    <span style={{ fontSize: 12, width: 16, textAlign: "center", flexShrink: 0, opacity: 0.5 }}>{sub.icon || "·"}</span>
                    <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{sub.label}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        ))}
      </div>
      <div style={{ padding: "10px 8px 14px", borderTop: `1px solid ${T.border}` }}>
        <button onClick={onLogout} style={{ background: "none", border: "none", borderRadius: 9, padding: "9px 12px", cursor: "pointer", display: "flex", alignItems: "center", gap: 9, color: T.textMuted, fontSize: 14, width: "100%", fontFamily: F.body, letterSpacing: "-0.01em" }}>
          <span style={{ fontSize: 15, width: 18, textAlign: "center", opacity: 0.6 }}>↩</span> Sign Out
        </button>
      </div>
    </div>
  );
}

function Header({ title, sub, right }) {
  return (
    <div style={{ padding: "24px 32px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "flex-start", justifyContent: "space-between", background: T.glass, backdropFilter: "saturate(180%) blur(20px)", WebkitBackdropFilter: "saturate(180%) blur(20px)", position: "sticky", top: 0, zIndex: 10 }}>
      <div>
        <h1 style={{ margin: 0, fontSize: 26, fontWeight: 700, color: T.text, letterSpacing: "-0.03em" }}>{title}</h1>
        {sub && <p style={{ margin: "3px 0 0", fontSize: 14, color: T.textMuted, fontWeight: 400 }}>{sub}</p>}
      </div>
      {right}
    </div>
  );
}

function Pane({ children }) {
  return <div style={{ padding: "28px 32px", display: "flex", flexDirection: "column", gap: 20 }}>{children}</div>;
}

/* ─────────────────────────────────────────────────────────────
   LOADING SCREEN
   ───────────────────────────────────────────────────────────── */
function LoadingScreen({ error }) {
  return (
    <div style={{ display: "flex", height: "100vh", alignItems: "center", justifyContent: "center", background: T.bg, fontFamily: F.body, flexDirection: "column", gap: 14 }}>
      {error ? (
        <>
          <span style={{ fontSize: 26, color: T.bad }}>⚠</span>
          <div style={{ fontSize: 15, color: T.bad, textAlign: "center", maxWidth: 320 }}>{error}</div>
        </>
      ) : (
        <>
          <span style={{ fontSize: 26, color: T.brand, animation: "spin 1s linear infinite" }}>◌</span>
          <div style={{ fontSize: 15, color: T.textMuted }}>Loading data…</div>
          <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
        </>
      )}
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   LOGIN PAGE
   ───────────────────────────────────────────────────────────── */
function LoginPage({ onLogin, users, msalErr, onDismissErr }) {
  const { instance } = useMsal();
  const [show, setShow] = useState(false);
  const [msLoading, setMsLoading] = useState(false);
  const [err, setErr] = useState("");

  useEffect(() => { setTimeout(() => setShow(true), 80); }, []);

  const handleMicrosoftLogin = async () => {
    setErr("");
    setMsLoading(true);
    try {
      await instance.loginRedirect(loginRequest);
    } catch (e) {
      setMsLoading(false);
      if (e.errorCode !== "user_cancelled") setErr("Sign-in failed. Please try again.");
    }
  };

  return (
    <div style={{ minHeight: "100vh", display: "flex", fontFamily: F.body, position: "relative", overflow: "hidden" }}>
      {/* Background image with dark overlay */}
      <div style={{ position: "absolute", inset: 0, backgroundImage: "url('/bg_lp.png')", backgroundSize: "cover", backgroundPosition: "center", zIndex: 0 }} />
      <div style={{ position: "absolute", inset: 0, background: "rgba(10,6,20,0.72)", zIndex: 1 }} />

      {msalErr && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.80)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 200, fontFamily: F.body }}>
          <div style={{ background: T.surface, border: `1px solid ${T.badBorder}`, borderRadius: 14, padding: "36px 32px", maxWidth: 420, width: "90%", boxShadow: "0 24px 60px rgba(0,0,0,0.5)" }}>
            <div style={{ fontSize: 17, fontWeight: 800, color: T.bad, marginBottom: 10 }}>Access Not Granted</div>
            <p style={{ margin: "0 0 8px", fontSize: 15, color: T.textSoft, lineHeight: 1.65 }}>
              The account <span style={{ color: T.text, fontWeight: 600 }}>{msalErr}</span> is not registered in this system.
            </p>
            <p style={{ margin: "0 0 24px", fontSize: 15, color: T.textSoft, lineHeight: 1.65 }}>
              Please contact your <strong style={{ color: T.text }}>team manager</strong> or the <strong style={{ color: T.text }}>IT team</strong> to request access.
            </p>
            <Btn primary onClick={onDismissErr}>OK</Btn>
          </div>
        </div>
      )}
      {/* Hero */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 72px", position: "relative", zIndex: 2, opacity: show ? 1 : 0, transform: show ? "none" : "translateX(-20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 44 }}>
          <div style={{ width: 60, height: 60, borderRadius: 13, background: `linear-gradient(135deg, ${T.brand}, #A78BFA)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 26, fontWeight: 900, color: "#fff" }}>NIET</div>
          <div>
            <div style={{ fontSize: 21, fontWeight: 900, color: "#fff" }}>NIET Group OKR's</div>
            <div style={{ fontSize: 12, color: "rgba(255,255,255,0.55)", letterSpacing: "0.14em" }}>NIET · CHARLTON BROWN · RHODES · EDUCARE</div>
          </div>
        </div>
        <h1 style={{ margin: "0 0 14px", fontSize: 44, fontWeight: 900, lineHeight: 1.08, color: "#fff", letterSpacing: "-0.03em", maxWidth: 460 }}>
          Align goals.<br /><span style={{ color: "#A78BFA" }}>Track everyone.</span><br />Drive results.
        </h1>
        <p style={{ margin: 0, fontSize: 16, color: "rgba(255,255,255,0.65)", lineHeight: 1.7, maxWidth: 400 }}>
          Monthly KPI reporting, weekly outcome tracking, real-time leaderboards — full transparency from company goals down to every team member.
        </p>
        <div style={{ marginTop: 48, display: "flex", gap: 36 }}>
          {[{ n: "Monthly", l: "KPI Reports" }, { n: "Weekly", l: "Submissions" }, { n: "Real-time", l: "Rankings" }, { n: "100%", l: "Transparent" }].map((x, i) => (
            <div key={i}>
              <div style={{ fontSize: 18, fontWeight: 900, color: "#A78BFA", fontFamily: F.mono }}>{x.n}</div>
              <div style={{ fontSize: 11, color: "rgba(255,255,255,0.5)", letterSpacing: "0.06em", marginTop: 2 }}>{x.l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Sign-in card */}
      <div style={{ width: 420, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 44px", position: "relative", zIndex: 2, opacity: show ? 1 : 0, transform: show ? "none" : "translateY(20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1) 0.15s" }}>
        <Card style={{ padding: "36px 30px" }}>
          <h2 style={{ margin: "0 0 4px", fontSize: 18, fontWeight: 800, color: T.text }}>Sign in</h2>
          <p style={{ margin: "0 0 24px", fontSize: 14, color: T.textMuted }}>Use your NIET Microsoft account to access your portal.</p>

          {err && (
            <div style={{ padding: "10px 14px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, fontSize: 14, color: T.bad, marginBottom: 16, lineHeight: 1.5 }}>{err}</div>
          )}

          {/* Microsoft button */}
          <button onClick={handleMicrosoftLogin} disabled={msLoading} style={{
            width: "100%", padding: "13px 16px", background: msLoading ? T.raised : "#fff",
            border: `1px solid ${T.border}`, borderRadius: 8, cursor: msLoading ? "not-allowed" : "pointer",
            display: "flex", alignItems: "center", justifyContent: "center", gap: 12,
            fontSize: 15, fontWeight: 700, color: "#1a1a1a", fontFamily: F.body, transition: "all 0.15s", opacity: msLoading ? 0.6 : 1,
          }}
            onMouseEnter={e => { if (!msLoading) e.currentTarget.style.boxShadow = "0 2px 12px rgba(0,0,0,0.18)"; }}
            onMouseLeave={e => { e.currentTarget.style.boxShadow = "none"; }}
          >
            <svg width="20" height="20" viewBox="0 0 21 21" xmlns="http://www.w3.org/2000/svg">
              <rect x="1"  y="1"  width="9" height="9" fill="#f25022"/>
              <rect x="11" y="1"  width="9" height="9" fill="#7fba00"/>
              <rect x="1"  y="11" width="9" height="9" fill="#00a4ef"/>
              <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>
            {msLoading ? "Signing in…" : "Sign in with Microsoft"}
          </button>

          <p style={{ margin: "18px 0 0", fontSize: 12, color: T.textDim, textAlign: "center", lineHeight: 1.6 }}>
            Your role and portal access are determined by your registered account.
          </p>
        </Card>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   USER MANAGEMENT PAGE
   ───────────────────────────────────────────────────────────── */
const BLANK_FORM = { name: "", email: "", role: "member", title: "", deptId: "", teamId: "", teamIds: [], mgrDeptIds: [] };

function UserMgmtPage({ users, depts, dispatch, currentUserId }) {
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState(BLANK_FORM);
  const [formErr, setFormErr] = useState("");
  const [editId, setEditId] = useState(null);
  const [editForm, setEditForm] = useState(BLANK_FORM);
  const [search, setSearch] = useState("");
  const [deptFilter, setDeptFilter] = useState("all");

  const teamsForDept = (deptId) => depts.find(d => d.id === deptId)?.teams || [];

  function handleAdd() {
    if (!form.name.trim() || !form.email.trim()) { setFormErr("Name and email are required."); return; }
    if (users.some(u => u.email.toLowerCase() === form.email.trim().toLowerCase())) { setFormErr("A user with this email already exists."); return; }
    const id = `usr_${Date.now().toString(36)}`;
    const newUser = {
      id, name: form.name.trim(), email: form.email.trim().toLowerCase(),
      role: form.role, av: makeAv(form.name), title: form.title.trim() || form.role,
      ...(form.deptId && { deptId: form.deptId }),
      ...((form.role === "member" || form.role === "manager") && form.teamId && { teamId: form.teamId }),
      ...(form.role === "manager" && form.teamIds?.length && { teamIds: form.teamIds }),
      ...(form.role === "manager" && form.mgrDeptIds?.length && { mgrDeptIds: form.mgrDeptIds }),
    };
    dispatch({ type: "ADD_USER", user: newUser });
    setForm(BLANK_FORM); setFormErr(""); setShowAdd(false);
  }

  function startEdit(u) {
    setEditId(u.id);
    setEditForm({ name: u.name, email: u.email, role: u.role, title: u.title || "", deptId: u.deptId || "", teamId: u.teamId || "", teamIds: u.teamIds || [], mgrDeptIds: u.mgrDeptIds || [], secondTeamId: u.secondTeamId || "" });
  }

  function saveEdit() {
    dispatch({ type: "UPDATE_USER", userId: editId, updates: {
      name: editForm.name.trim(), email: editForm.email.trim().toLowerCase(),
      role: editForm.role, av: makeAv(editForm.name), title: editForm.title.trim() || editForm.role,
      deptId: editForm.deptId || undefined,
      teamId: (editForm.role === "member" || editForm.role === "manager") ? (editForm.teamId || undefined) : undefined,
      teamIds: editForm.role === "manager" ? (editForm.teamIds?.length ? editForm.teamIds : undefined) : undefined,
      mgrDeptIds: editForm.role === "manager" ? (editForm.mgrDeptIds?.length ? editForm.mgrDeptIds : undefined) : undefined,
      secondTeamId: editForm.role === "member" ? (editForm.secondTeamId || undefined) : undefined,
    }});
    setEditId(null);
  }

  const roleColor = { admin: T.brand, manager: T.orange, member: T.ok };

  const roleCounts = users.reduce((a, u) => { a[u.role] = (a[u.role] || 0) + 1; return a; }, {});
  const filteredUsers = users.filter(u => {
    const matchesSearch = !search.trim() || u.name.toLowerCase().includes(search.toLowerCase()) || u.email.toLowerCase().includes(search.toLowerCase());
    const matchesDept = deptFilter === "all" || (deptFilter === "__none__" ? !u.deptId : u.deptId === deptFilter);
    return matchesSearch && matchesDept;
  });

  return (<>
    <Header title="User Management" sub="Add users, set roles, and control portal access"
      right={<Btn primary onClick={() => { setShowAdd(p => !p); setFormErr(""); setForm(BLANK_FORM); }}>{showAdd ? "Cancel" : "+ Add User"}</Btn>} />
    <Pane>
      {/* Stats */}
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <Metric label="Total Users"  value={users.length} />
        <Metric label="Admins"       value={roleCounts.admin   || 0} />
        <Metric label="Managers"     value={roleCounts.manager || 0} />
        <Metric label="Members"      value={roleCounts.member  || 0} />
      </div>

      {/* Add user form */}
      {showAdd && (
        <Card style={{ padding: 22 }}>
          <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 16, color: T.text }}>New User</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Full Name *</div>
              <Input value={form.name} onChange={e => setForm(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Jane Smith" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Email Address *</div>
              <Input value={form.email} onChange={e => setForm(p => ({ ...p, email: e.target.value }))} placeholder="jane@niet.edu.au" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Role *</div>
              <Select value={form.role} onChange={e => setForm(p => ({ ...p, role: e.target.value, deptId: "", teamId: "", teamIds: [] }))} style={{ width: "100%" }}>
                <option value="admin">Admin</option>
                <option value="manager">Manager</option>
                <option value="member">Member</option>
              </Select>
            </div>
            <div>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Job Title</div>
              <Input value={form.title} onChange={e => setForm(p => ({ ...p, title: e.target.value }))} placeholder="e.g. Admissions Officer" style={{ width: "100%" }} />
            </div>
            {form.role !== "admin" && (
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Department</div>
                <Select value={form.deptId} onChange={e => setForm(p => ({ ...p, deptId: e.target.value, teamId: "", teamIds: [] }))} style={{ width: "100%" }}>
                  <option value="">— Select department —</option>
                  {depts.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
                </Select>
              </div>
            )}
            {form.role !== "admin" && form.deptId && teamsForDept(form.deptId).length > 0 && (
              <div>
                <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Team</div>
                <Select value={form.teamId} onChange={e => setForm(p => ({ ...p, teamId: e.target.value }))} style={{ width: "100%" }}>
                  <option value="">— Select team —</option>
                  {teamsForDept(form.deptId).map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                </Select>
              </div>
            )}
          </div>
          {form.role === "manager" && form.deptId && depts.filter(d => d.id !== form.deptId).length > 0 && (
            <div style={{ marginBottom: 12 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 6 }}>Additional Departments (optional)</div>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 12 }}>
                {depts.filter(d => d.id !== form.deptId).map(d => (
                  <label key={d.id} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 13, cursor: "pointer", color: T.text }}>
                    <input type="checkbox" checked={(form.mgrDeptIds || []).includes(d.id)}
                      onChange={e => setForm(p => ({ ...p, mgrDeptIds: e.target.checked ? [...(p.mgrDeptIds || []), d.id] : (p.mgrDeptIds || []).filter(id => id !== d.id) }))} />
                    {d.name}
                  </label>
                ))}
              </div>
              <div style={{ fontSize: 11, color: T.textDim, marginTop: 5 }}>Manager can view and approve submissions from these departments (full dept access).</div>
            </div>
          )}
          {formErr && <div style={{ padding: "8px 12px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 6, fontSize: 13, color: T.bad, marginBottom: 12 }}>{formErr}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <Btn small onClick={() => { setShowAdd(false); setForm(BLANK_FORM); setFormErr(""); }}>Cancel</Btn>
            <Btn primary small onClick={handleAdd}>Create User</Btn>
          </div>
        </Card>
      )}

      {/* Search + Department filter */}
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ position: "relative", flex: 1, maxWidth: 320 }}>
            <span style={{ position: "absolute", left: 12, top: "50%", transform: "translateY(-50%)", color: T.textDim, fontSize: 15, pointerEvents: "none" }}>⌕</span>
            <Input
              value={search}
              onChange={e => setSearch(e.target.value)}
              placeholder="Search by name or email…"
              style={{ width: "100%", paddingLeft: 34 }}
            />
          </div>
          <span style={{ fontSize: 13, color: T.textMuted }}>
            {filteredUsers.length} of {users.length} user{users.length !== 1 ? "s" : ""}
          </span>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <Btn small primary={deptFilter === "all"} onClick={() => setDeptFilter("all")}>All</Btn>
          {depts.map(d => (
            <Btn key={d.id} small primary={deptFilter === d.id} onClick={() => setDeptFilter(d.id)}>
              {d.name} <span style={{ opacity: 0.7, fontWeight: 400 }}>({users.filter(u => u.deptId === d.id).length})</span>
            </Btn>
          ))}
          <Btn small primary={deptFilter === "__none__"} onClick={() => setDeptFilter("__none__")}>
            No Dept <span style={{ opacity: 0.7, fontWeight: 400 }}>({users.filter(u => !u.deptId).length})</span>
          </Btn>
        </div>
      </div>

      <Card style={{ overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "32px 1fr 160px 80px 120px 130px 90px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
          <span></span><span>Name / Email</span><span>Title</span><span>Role</span><span>Department</span><span>Team</span><span style={{ textAlign: "right" }}>Actions</span>
        </div>
        {filteredUsers.length === 0 && <EmptyState text={`No users match "${search}".`} />}
        {filteredUsers.map((u, i) => {
          const dept = depts.find(d => d.id === u.deptId);
          const team = dept?.teams.find(t => t.id === u.teamId);
          const managerTeams = u.teamIds?.length ? dept?.teams.filter(t => u.teamIds.includes(t.id)) : [];
          const isSystem = u.id === "sysadmin";
          const isSelf = u.id === currentUserId;

          if (editId === u.id) {
            const editTeams = teamsForDept(editForm.deptId);
            const lbl = { fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 4 };
            return (
              <div key={u.id} style={{ background: T.brandDim, borderBottom: `1px solid ${T.border}`, padding: "14px 18px" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                  <div><div style={lbl}>Full Name</div><Input value={editForm.name} onChange={e => setEditForm(p => ({ ...p, name: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={lbl}>Email</div><Input value={editForm.email} onChange={e => setEditForm(p => ({ ...p, email: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={lbl}>Job Title</div><Input value={editForm.title} onChange={e => setEditForm(p => ({ ...p, title: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={lbl}>Role</div>
                    <Select value={editForm.role} onChange={e => setEditForm(p => ({ ...p, role: e.target.value, deptId: "", teamId: "", teamIds: [] }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }}>
                      <option value="admin">Admin</option>
                      <option value="manager">Manager</option>
                      <option value="member">Member</option>
                    </Select>
                  </div>
                  {editForm.role !== "admin" && (
                    <div><div style={lbl}>Department</div>
                      <Select value={editForm.deptId} onChange={e => setEditForm(p => ({ ...p, deptId: e.target.value, teamId: "", teamIds: [] }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }}>
                        <option value="">— Select department —</option>
                        {depts.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
                      </Select>
                    </div>
                  )}
                  {(editForm.role === "member" || editForm.role === "manager") && editTeams.length > 0 && (
                    <div>
                      <div style={lbl}>{editForm.role === "manager" ? "My Team (KPI tracking)" : "Team"}</div>
                      <Select value={editForm.teamId} onChange={e => setEditForm(p => ({ ...p, teamId: e.target.value, secondTeamId: p.secondTeamId === e.target.value ? "" : p.secondTeamId }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }}>
                        <option value="">— Select team —</option>
                        {editTeams.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                      </Select>
                      {editForm.role === "manager" && <div style={{ fontSize: 11, color: T.textDim, marginTop: 4 }}>This team's KRs are synced to this manager's own KPI list for personal tracking.</div>}
                    </div>
                  )}
                  {editForm.role === "member" && editTeams.filter(t => t.id !== editForm.teamId).length > 0 && (
                    <div><div style={lbl}>Second Team (optional)</div>
                      <Select value={editForm.secondTeamId || ""} onChange={e => setEditForm(p => ({ ...p, secondTeamId: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }}>
                        <option value="">— None —</option>
                        {editTeams.filter(t => t.id !== editForm.teamId).map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                      </Select>
                    </div>
                  )}
                </div>
                {editForm.role === "manager" && editTeams.length > 0 && (
                  <div style={{ marginBottom: 12 }}>
                    <div style={lbl}>Teams (manager oversees)</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                      {editTeams.map(t => (
                        <label key={t.id} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 13, cursor: "pointer", color: T.text }}>
                          <input type="checkbox" checked={editForm.teamIds.includes(t.id)} onChange={e => setEditForm(p => ({ ...p, teamIds: e.target.checked ? [...p.teamIds, t.id] : p.teamIds.filter(id => id !== t.id) }))} />
                          {t.name}
                        </label>
                      ))}
                    </div>
                    <div style={{ fontSize: 11, color: T.textDim, marginTop: 6 }}>Teams this manager can view and approve member submissions for.</div>
                  </div>
                )}
                {editForm.role === "manager" && editForm.deptId && depts.filter(d => d.id !== editForm.deptId).length > 0 && (
                  <div style={{ marginBottom: 12 }}>
                    <div style={lbl}>Additional Departments</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                      {depts.filter(d => d.id !== editForm.deptId).map(d => (
                        <label key={d.id} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 13, cursor: "pointer", color: T.text }}>
                          <input type="checkbox" checked={(editForm.mgrDeptIds || []).includes(d.id)}
                            onChange={e => setEditForm(p => ({ ...p, mgrDeptIds: e.target.checked ? [...(p.mgrDeptIds || []), d.id] : (p.mgrDeptIds || []).filter(id => id !== d.id) }))} />
                          {d.name}
                        </label>
                      ))}
                    </div>
                    <div style={{ fontSize: 11, color: T.textDim, marginTop: 6 }}>Manager can view and approve submissions from these departments (full dept access, no team filter).</div>
                  </div>
                )}
                <div style={{ display: "flex", gap: 8 }}>
                  <Btn small onClick={() => setEditId(null)}>Cancel</Btn>
                  <Btn primary small onClick={saveEdit}>Save</Btn>
                </div>
              </div>
            );
          }

          return (
            <div key={u.id} style={{ display: "grid", gridTemplateColumns: "32px 1fr 160px 80px 120px 130px 90px", padding: "10px 18px", gap: 10, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
              <Avatar letters={u.av} size={26} />
              <div>
                <div style={{ fontWeight: 600, color: T.text }}>{u.name}{isSelf && <span style={{ fontSize: 11, color: T.brand, marginLeft: 6 }}>you</span>}</div>
                <div style={{ fontSize: 12, color: T.textMuted }}>{isSystem ? "System login only" : u.email}</div>
              </div>
              <span style={{ fontSize: 13, color: T.textSoft }}>{u.title}</span>
              <RoleTag role={u.role} />
              <span style={{ fontSize: 13, color: T.textMuted }}>{dept?.name || "—"}{u.mgrDeptIds?.length > 0 && <span style={{ fontSize: 11, color: T.brand, marginLeft: 5 }}>+{u.mgrDeptIds.length} dept{u.mgrDeptIds.length > 1 ? "s" : ""}</span>}</span>
              <span style={{ fontSize: 12, color: T.textMuted }}>
                {u.role === "member" && team ? (() => { const st = u.secondTeamId ? dept?.teams.find(t => t.id === u.secondTeamId) : null; return st ? `${team.name} / ${st.name}` : team.name; })() : u.role === "manager" && managerTeams.length ? managerTeams.map(t => t.name).join(", ") : "—"}
              </span>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                {!isSystem && (
                  <button onClick={() => startEdit(u)} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                )}
                {!isSystem && !isSelf && (
                  <button onClick={() => { if (window.confirm(`Delete user "${u.name}"? This cannot be undone.`)) dispatch({ type: "REMOVE_USER", userId: u.id }); }} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>✕</button>
                )}
              </div>
            </div>
          );
        })}
      </Card>
    </Pane>
  </>);
}

/* ─────────────────────────────────────────────────────────────
   DEPARTMENT MANAGEMENT PAGE
   ───────────────────────────────────────────────────────────── */
const BLANK_DEPT = { name: "", obj: "", head: "", college: "" };

function DeptMgmtPage({ depts, users, memberData, okrSubmissions, dispatch }) {
  const [showAdd, setShowAdd] = useState(false);
  const [addForm, setAddForm] = useState(BLANK_DEPT);
  const [addErr, setAddErr] = useState("");
  const [editId, setEditId] = useState(null);
  const [editForm, setEditForm] = useState(BLANK_DEPT);
  const [confirmDel, setConfirmDel] = useState(null);
  const [selDept, setSelDept] = useState(null);
  const [showAddTeam, setShowAddTeam] = useState(null);
  const [addTeamForm, setAddTeamForm] = useState({ name: "", lead: "", obj: "" });
  const [editTeam, setEditTeam] = useState(null);
  const [editTeamForm, setEditTeamForm] = useState({ name: "", lead: "", obj: "", members: [] });
  const [confirmDelTeam, setConfirmDelTeam] = useState(null);

  function handleAdd() {
    if (!addForm.name.trim()) { setAddErr("Department name is required."); return; }
    if (depts.some(d => d.name.toLowerCase() === addForm.name.trim().toLowerCase())) { setAddErr("A department with this name already exists."); return; }
    const id = `dept_${Date.now().toString(36)}`;
    dispatch({ type: "ADD_DEPT", dept: { id, name: addForm.name.trim(), obj: addForm.obj.trim(), head: addForm.head.trim(), college: addForm.college.trim(), krs: [], teams: [] } });
    setAddForm(BLANK_DEPT); setAddErr(""); setShowAdd(false);
  }

  function startEdit(d) {
    setEditId(d.id);
    setEditForm({ name: d.name, obj: d.obj || "", head: d.head || "", college: d.college || "" });
  }

  function saveEdit() {
    if (!editForm.name.trim()) return;
    dispatch({ type: "UPDATE_DEPT", deptId: editId, updates: { name: editForm.name.trim(), obj: editForm.obj.trim(), head: editForm.head.trim(), college: editForm.college.trim() } });
    setEditId(null);
  }

  const labelStyle = { fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 };

  return (<>
    <Header title="Departments" sub="Manage department structure and descriptions"
      right={<Btn primary onClick={() => { setShowAdd(p => !p); setAddErr(""); setAddForm(BLANK_DEPT); }}>{showAdd ? "Cancel" : "+ Add Department"}</Btn>} />
    <Pane>
      <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
        <Metric label="Total Departments" value={depts.length} />
        <Metric label="Total Teams" value={depts.reduce((a, d) => a + d.teams.length, 0)} />
      </div>

      {showAdd && (
        <Card style={{ padding: 22 }}>
          <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 16 }}>New Department</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <div>
              <div style={labelStyle}>Department Name *</div>
              <Input value={addForm.name} onChange={e => setAddForm(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Finance" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={labelStyle}>Department Head</div>
              <Input value={addForm.head} onChange={e => setAddForm(p => ({ ...p, head: e.target.value }))} placeholder="e.g. Jane Smith" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={labelStyle}>College / Group</div>
              <Input value={addForm.college} onChange={e => setAddForm(p => ({ ...p, college: e.target.value }))} placeholder="e.g. NIET" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={labelStyle}>Description / Objective</div>
              <Input value={addForm.obj} onChange={e => setAddForm(p => ({ ...p, obj: e.target.value }))} placeholder="e.g. Drive enrolment targets" style={{ width: "100%" }} />
            </div>
          </div>
          {addErr && <div style={{ padding: "8px 12px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 6, fontSize: 13, color: T.bad, marginBottom: 12 }}>{addErr}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <Btn small onClick={() => { setShowAdd(false); setAddForm(BLANK_DEPT); setAddErr(""); }}>Cancel</Btn>
            <Btn primary small onClick={handleAdd}>Create Department</Btn>
          </div>
        </Card>
      )}

      <Card style={{ overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 200px 80px 80px 110px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
          <span>Department / Description</span><span>Head · College</span><span>Teams</span><span>Completion</span><span style={{ textAlign: "right" }}>Actions</span>
        </div>
        {depts.length === 0 && <EmptyState text="No departments yet. Add one above." />}
        {depts.map((d, i) => {
          const r = calcRate(d.krs); const s = getStatus(r);
          if (editId === d.id) {
            return (
              <div key={d.id} style={{ background: T.brandDim, borderBottom: `1px solid ${T.border}`, padding: "14px 18px" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                  <div><div style={labelStyle}>Name *</div><Input value={editForm.name} onChange={e => setEditForm(p => ({ ...p, name: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={labelStyle}>Description / Objective</div><Input value={editForm.obj} onChange={e => setEditForm(p => ({ ...p, obj: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={labelStyle}>Department Head</div><Input value={editForm.head} onChange={e => setEditForm(p => ({ ...p, head: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={labelStyle}>College / Group</div><Input value={editForm.college} onChange={e => setEditForm(p => ({ ...p, college: e.target.value }))} style={{ fontSize: 13, padding: "7px 10px", width: "100%" }} /></div>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <Btn small onClick={() => setEditId(null)}>Cancel</Btn>
                  <Btn primary small onClick={saveEdit}>Save</Btn>
                </div>
              </div>
            );
          }
          const isSelected = selDept === d.id;
          return (
            <div key={d.id}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 200px 80px 80px 110px", padding: "12px 18px", gap: 10, alignItems: "center", background: isSelected ? T.brandDim : i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <div style={{ cursor: "pointer" }} onClick={() => { setSelDept(isSelected ? null : d.id); }}>
                  <div style={{ fontWeight: 700, color: isSelected ? T.brand : T.text }}>{d.name}</div>
                  {d.obj && <div style={{ fontSize: 12, color: T.textMuted, marginTop: 2 }}>{d.obj}</div>}
                </div>
                <div>
                  <div style={{ fontSize: 13, color: T.textSoft }}>{d.head || "—"}</div>
                  <div style={{ fontSize: 12, color: T.textMuted }}>{d.college || ""}</div>
                </div>
                <span style={{ fontSize: 13, color: T.textSoft }}>{d.teams.length} team{d.teams.length !== 1 ? "s" : ""}</span>
                <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 14, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</span>
                <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                  <button onClick={() => startEdit(d)} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                  {confirmDel === d.id ? (<>
                    <button onClick={() => { dispatch({ type: "REMOVE_DEPT", deptId: d.id }); setConfirmDel(null); if (selDept === d.id) setSelDept(null); }} style={{ background: T.bad, border: "none", borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: "#fff", fontSize: 12, fontWeight: 700, fontFamily: F.body }}>Confirm</button>
                    <button onClick={() => setConfirmDel(null)} style={{ background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.textSoft, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>✕</button>
                  </>) : (
                    <button onClick={() => setConfirmDel(d.id)} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>✕</button>
                  )}
                </div>
              </div>
              {isSelected && (() => {
                const r2 = calcRate(d.krs); const s2 = getStatus(r2);
                const deptMembers = users
                  .filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id)
                  .map(u => { const kd = memberData[u.id] || { krs: [] }; const mr = calcMemberRate(u.id, kd.krs, okrSubmissions); return { ...u, rate: mr, status: getStatus(mr) }; })
                  .sort((a, b) => b.rate - a.rate);
                return (
                  <div style={{ background: T.bgSoft, borderBottom: `1px solid ${T.border}`, padding: "16px 18px" }}>
                    <div style={{ display: "flex", gap: 12, marginBottom: 14 }}>
                      <Metric label="Completion" value={`${r2.toFixed(1)}%`} status={s2} />
                      <Metric label="Teams"   value={d.teams.length} />
                      <Metric label="Members" value={deptMembers.length} />
                    </div>

                    {d.krs.length > 0 && (
                      <Card style={{ overflow: "hidden", marginBottom: 14 }}>
                        <div style={{ padding: "9px 16px", borderBottom: `1px solid ${T.border}`, fontSize: 12, fontWeight: 700, color: T.textMuted, letterSpacing: "0.05em" }}>DEPARTMENT KEY RESULTS</div>
                        {d.krs.map((kr, ki) => { const cr = krCompletion(kr); const cs = getStatus(cr);
                          return (<div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 70px 70px 55px 130px 65px", padding: "9px 16px", gap: 8, alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                            <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span><span>{kr.label}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700 }}>{fmt(kr.actual)}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span>
                            <Bar value={cr} status={cs} h={5} />
                            <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={cs} small /></div>
                          </div>);
                        })}
                      </Card>
                    )}

                    <Card style={{ overflow: "hidden", marginBottom: 14 }}>
                      <div style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 120px 55px 140px 70px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                        <span>#</span><span></span><span>Name</span><span>Title</span><span style={{ textAlign: "right" }}>Rate</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                      </div>
                      {deptMembers.length === 0
                        ? <div style={{ padding: "14px 18px", fontSize: 14, color: T.textMuted }}>No members assigned to this department.</div>
                        : deptMembers.map((m, mi) => (
                          <div key={m.id} style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 120px 55px 140px 70px", padding: "9px 18px", gap: 10, alignItems: "center", background: mi % 2 ? T.raised : "transparent", borderBottom: mi < deptMembers.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 14 }}>
                            <span style={{ fontFamily: F.mono, fontWeight: 800, color: mi === 0 ? T.ok : mi === deptMembers.length - 1 && deptMembers.length > 2 ? T.bad : T.textMuted }}>#{mi + 1}</span>
                            <Avatar letters={m.av} size={26} />
                            <div><span style={{ fontWeight: 600 }}>{m.name}</span></div>
                            <span style={{ fontSize: 12, color: T.textMuted }}>{m.title || "—"}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[m.status].color }}>{m.rate.toFixed(1)}%</span>
                            <Bar value={m.rate} status={m.status} h={4} />
                            <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={m.status} small /></div>
                          </div>
                        ))
                      }
                    </Card>

                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                      <SectionLabel style={{ margin: 0 }}>Teams</SectionLabel>
                      <Btn small primary onClick={() => { setShowAddTeam(showAddTeam === d.id ? null : d.id); setAddTeamForm({ name: "", lead: "", obj: "" }); setEditTeam(null); }}>
                        {showAddTeam === d.id ? "Cancel" : "+ Add Team"}
                      </Btn>
                    </div>

                    {showAddTeam === d.id && (
                      <Card style={{ padding: 16, marginBottom: 10, borderLeft: `3px solid ${T.brand}` }}>
                        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 10 }}>New Team</div>
                        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                          <div><div style={labelStyle}>Team Name *</div><Input value={addTeamForm.name} onChange={e => setAddTeamForm(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Domestic Team" style={{ width: "100%" }} /></div>
                          <div><div style={labelStyle}>Team Lead</div><Input value={addTeamForm.lead} onChange={e => setAddTeamForm(p => ({ ...p, lead: e.target.value }))} placeholder="e.g. Jane Smith" style={{ width: "100%" }} /></div>
                          <div><div style={labelStyle}>Objective</div><Input value={addTeamForm.obj} onChange={e => setAddTeamForm(p => ({ ...p, obj: e.target.value }))} placeholder="e.g. Hit domestic KPIs" style={{ width: "100%" }} /></div>
                        </div>
                        <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                          <Btn small onClick={() => { setShowAddTeam(null); setAddTeamForm({ name: "", lead: "", obj: "" }); }}>Cancel</Btn>
                          <Btn primary small disabled={!addTeamForm.name.trim()} onClick={() => {
                            dispatch({ type: "ADD_TEAM", deptId: d.id, team: { id: `t${Date.now().toString(36)}`, name: addTeamForm.name.trim(), lead: addTeamForm.lead.trim(), obj: addTeamForm.obj.trim(), krs: [], members: [] } });
                            setShowAddTeam(null); setAddTeamForm({ name: "", lead: "", obj: "" });
                          }}>Create Team</Btn>
                        </div>
                      </Card>
                    )}

                    {d.teams.length === 0 && showAddTeam !== d.id && (
                      <div style={{ fontSize: 13, color: T.textMuted, padding: "8px 0 12px" }}>No teams yet. Add one above.</div>
                    )}

                    {d.teams.map(t => { const tr = calcRate(t.krs); const ts = getStatus(tr);
                      const isEditingTeam = editTeam?.deptId === d.id && editTeam?.teamId === t.id;
                      return (
                        <Card key={t.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                          {isEditingTeam ? (
                            <div style={{ padding: 16 }}>
                              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 14 }}>
                                <div><div style={labelStyle}>Team Name *</div><Input value={editTeamForm.name} onChange={e => setEditTeamForm(p => ({ ...p, name: e.target.value }))} style={{ width: "100%" }} /></div>
                                <div><div style={labelStyle}>Team Lead</div><Input value={editTeamForm.lead} onChange={e => setEditTeamForm(p => ({ ...p, lead: e.target.value }))} style={{ width: "100%" }} /></div>
                                <div><div style={labelStyle}>Objective</div><Input value={editTeamForm.obj} onChange={e => setEditTeamForm(p => ({ ...p, obj: e.target.value }))} style={{ width: "100%" }} /></div>
                              </div>
                              {(() => {
                                const deptMembers = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id);
                                return deptMembers.length > 0 && (
                                  <div style={{ marginBottom: 14 }}>
                                    <div style={labelStyle}>Team Members</div>
                                    <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                                      {deptMembers.map(u => {
                                        const checked = editTeamForm.members.includes(u.id);
                                        return (
                                          <label key={u.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 13, cursor: "pointer", padding: "5px 10px", borderRadius: 8, border: `1px solid ${checked ? T.brandBorder : T.border}`, background: checked ? T.brandDim : T.raised, color: checked ? T.brand : T.text }}>
                                            <input type="checkbox" checked={checked} onChange={e => setEditTeamForm(p => ({ ...p, members: e.target.checked ? [...p.members, u.id] : p.members.filter(id => id !== u.id) }))} style={{ accentColor: T.brand }} />
                                            <Avatar letters={u.av} size={18} />
                                            <span style={{ fontWeight: checked ? 700 : 400 }}>{u.name}</span>
                                          </label>
                                        );
                                      })}
                                    </div>
                                  </div>
                                );
                              })()}
                              <div style={{ display: "flex", gap: 8 }}>
                                <Btn small onClick={() => setEditTeam(null)}>Cancel</Btn>
                                <Btn primary small disabled={!editTeamForm.name.trim()} onClick={() => {
                                  dispatch({ type: "UPDATE_TEAM", deptId: d.id, teamId: t.id, updates: { name: editTeamForm.name.trim(), lead: editTeamForm.lead.trim(), obj: editTeamForm.obj.trim(), members: editTeamForm.members } });
                                  setEditTeam(null);
                                }}>Save</Btn>
                              </div>
                            </div>
                          ) : (
                            <div style={{ padding: "10px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                              <div>
                                <span style={{ fontSize: 15, fontWeight: 700 }}>{t.name}</span>
                                {t.lead && <span style={{ fontSize: 12, color: T.textMuted, marginLeft: 10 }}>Lead: {t.lead}</span>}
                                {t.obj && <div style={{ fontSize: 12, color: T.textMuted, marginTop: 2 }}>{t.obj}</div>}
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                                <span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[ts].color }}>{tr.toFixed(1)}%</span>
                                <Tag type={ts} small />
                                <button onClick={() => { setEditTeam({ deptId: d.id, teamId: t.id }); setEditTeamForm({ name: t.name, lead: t.lead || "", obj: t.obj || "", members: t.members || [] }); setShowAddTeam(null); }} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                                {confirmDelTeam?.deptId === d.id && confirmDelTeam?.teamId === t.id ? (<>
                                  <button onClick={() => { dispatch({ type: "REMOVE_TEAM", deptId: d.id, teamId: t.id }); setConfirmDelTeam(null); }} style={{ background: T.bad, border: "none", borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: "#fff", fontSize: 12, fontWeight: 700, fontFamily: F.body }}>Confirm</button>
                                  <button onClick={() => setConfirmDelTeam(null)} style={{ background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.textSoft, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>✕</button>
                                </>) : (
                                  <button onClick={() => setConfirmDelTeam({ deptId: d.id, teamId: t.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>✕</button>
                                )}
                              </div>
                            </div>
                          )}
                        </Card>
                      );
                    })}
                  </div>
                );
              })()}
            </div>
          );
        })}
      </Card>
    </Pane>
  </>);
}

/* ─────────────────────────────────────────────────────────────
   ADMIN PORTAL
   ───────────────────────────────────────────────────────────── */
function AdminPortal({ user, onLogout, state, dispatch }) {
  const [page, setPageRaw] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'admin' ? (p[2] || 'overview') : 'overview';
  });
  const [selDept, setSelDept] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'admin' && p[2] === 'departments' ? (p[3] || null) : null;
  });
  const setPage = useCallback(p => { window.history.pushState(null, '', `/admin/${p}`); setPageRaw(p); }, []);
  useEffect(() => {
    if (window.location.pathname.split('/')[1] !== 'admin') {
      window.history.replaceState(null, '', `/admin/overview`);
    }
    const onPop = () => {
      const p = window.location.pathname.split('/');
      setPageRaw(p[1] === 'admin' ? (p[2] || 'overview') : 'overview');
      setSelDept(p[1] === 'admin' && p[2] === 'departments' ? (p[3] || null) : null);
    };
    window.addEventListener('popstate', onPop);
    return () => window.removeEventListener('popstate', onPop);
  }, []);
  useEffect(() => {
    if (page === 'departments') {
      const target = selDept ? `/admin/departments/${selDept}` : '/admin/departments';
      if (window.location.pathname !== target) window.history.replaceState(null, '', target);
    }
  }, [selDept, page]);
  const [selTeam, setSelTeam] = useState(null);
  const [newKr, setNewKr] = useState({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false });
  const [addTarget, setAddTarget] = useState(null);
  const [showGenReport, setShowGenReport] = useState(false);
  const [genPeriod, setGenPeriod] = useState({ label: "", from: "", to: "" });
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ name: "", progress: 0, status: "active", log: "", due: "" });
  const [subFilter, setSubFilter] = useState("all");
  const [colWidths, setColWidths] = useState({ id: 50, label: 220, operator: 72, period: 90, target: 90, actual: 80, unit: 100, dataSource: 200 });
  const colWidthsRef = useRef({ id: 50, label: 220, operator: 72, period: 90, target: 90, actual: 80, unit: 100, dataSource: 200 });
  const dragColRef = useRef(null);
  const [hiddenCols, setHiddenCols] = useState(new Set());
  const [customColWidthOverride, setCustomColWidthOverride] = useState(null);
  const [addingCol, setAddingCol] = useState(false);
  const [newColName, setNewColName] = useState("");
  const [overviewView, setOverviewView] = useState("monthly");
  const [syncPrompt, setSyncPrompt] = useState(null);
  const syncTimerRef = useRef(null);
  const [dirtySync, setDirtySync] = useState(null);
  const [editReportId, setEditReportId] = useState(null);
  const [editReportForm, setEditReportForm] = useState({ month: "", notes: "" });
  const [reportPeriodView, setReportPeriodView] = useState("all");
  const [lbSearch, setLbSearch] = useState("");
  const [lbDeptFilter, setLbDeptFilter] = useState("all");
  const [lbPeriod, setLbPeriod] = useState("all");
  const [lbExpandedMember, setLbExpandedMember] = useState(null);
  const [confirmDeleteKr, setConfirmDeleteKr] = useState(null);
  const [syncNote, setSyncNote] = useState(null);
  const syncNoteTimer = useRef(null);
  const [subSearch, setSubSearch] = useState("");
  const [subDeptFilter, setSubDeptFilter] = useState("all");
  const [expandedMonthlyKr, setExpandedMonthlyKr] = useState(null);
  const [expandedPersonalMember, setExpandedPersonalMember] = useState(null);
  const [addPersonalKr, setAddPersonalKr] = useState(null);
  const [subPeriod, setSubPeriod] = useState("monthly");
  const [sendingCheckin, setSendingCheckin] = useState(false);
  const [checkinResult, setCheckinResult] = useState(null);
  const [checkinScope, setCheckinScope] = useState({ deptId: "", teamId: "", userId: "" });
  const [checkinPreview, setCheckinPreview] = useState(null);
  const [rejectOkr, setRejectOkr] = useState(null);
  const [tmplPeriod, setTmplPeriod] = useState("default");
  const [testEmailState, setTestEmailState] = useState({ status: "idle", msg: "" });
  const [testEmailTo, setTestEmailTo] = useState(user?.email || "");

  const { depts, memberData, mgrSprints, monthlyReports, projects, weeklySubs, okrSubmissions = [], users, settings } = state;
  const colOrder = settings?.colOrder || ["id", "label", "operator", "period", "target", "actual", "unit", "dataSource"];
  const navItems = [
    { id: "overview",         icon: "◎", label: "Company Overview"  },
    { id: "departments",      icon: "⬛", label: "Departments"       },
    { id: "submissions",      icon: "✉", label: "OKR Submissions"   },
    { id: "reports",          icon: "⊞", label: "OKR Reports"       },
    { id: "projects",         icon: "⚡", label: "Projects"          },
    { id: "leaderboard",      icon: "▲", label: "Leaderboard"       },
    { id: "users",            icon: "⊹", label: "User Management"   },
    { id: "email-templates",  icon: "✦", label: "Email Templates"   },
  ];
  const deptSubItems = [
    { id: "__all__", label: "All Departments", icon: "⊕" },
    ...depts.map(d => ({ id: d.id, label: d.name })),
  ];

  const OV_TYPES = { weekly: ["daily","weekly"], monthly: ["daily","weekly","monthly"], annual: ["daily","weekly","monthly","quarterly","biannual","annual"] };
  const ovTypes = OV_TYPES[overviewView] || OV_TYPES.monthly;
  const ovSubs = okrSubmissions.filter(s => {
    if (!ovTypes.includes(s.period) || !s.sentAt) return false;
    const d = new Date(s.sentAt), now = new Date();
    if (overviewView === "weekly") {
      const dow = now.getDay();
      const mon = new Date(now); mon.setHours(0,0,0,0); mon.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1));
      const sun = new Date(mon); sun.setDate(mon.getDate() + 7);
      return d >= mon && d < sun;
    }
    if (overviewView === "monthly") return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
    const m = now.getMonth() + 1, y = now.getFullYear();
    const fyS = new Date(m >= 7 ? y : y - 1, 6, 1);
    const fyE = new Date(m >= 7 ? y + 1 : y, 5, 30, 23, 59, 59);
    return d >= fyS && d <= fyE;
  });
  const filtKrs = (krs) => krs.filter(kr => ovTypes.includes(kr.period || "monthly"));
  const deptRanks = depts.map(d => {
    const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id);
    const now = Date.now();
    const rates = members.map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const krs = filtKrs(kd.krs);
      // Include member only if they have at least one answered or overdue-unanswered submission
      const hasEligible = krs.some(kr => {
        const krSubs = ovSubs.filter(s => s.memberId === u.id && s.krId === kr.id);
        if (!krSubs.length) return false;
        if (krSubs.some(s => s.answer !== null)) return true;
        const latest = krSubs.filter(s => s.sentAt).sort((a, b) => b.sentAt.localeCompare(a.sentAt))[0];
        return latest && (now - new Date(latest.sentAt).getTime()) >= 86400000;
      });
      if (!hasEligible) return null;
      return calcMemberRate(u.id, krs, ovSubs);
    }).filter(r => r !== null);
    const rate = rates.length ? rates.reduce((a, b) => a + b, 0) / rates.length : 0;
    return { ...d, rate, status: getStatus(rate) };
  }).sort((a, b) => b.rate - a.rate);
  const compRate = deptRanks.length ? deptRanks.reduce((a, d) => a + d.rate, 0) / deptRanks.length : 0;
  const rptMonthKey = currentFYMonthKey();
  const rptSubs = okrSubmissions.filter(s => s.answer !== null && (s.periodKey || "").slice(0, 7) === rptMonthKey);
  const rptSubRate = rptSubs.length > 0 ? Math.round((rptSubs.filter(s => s.answer === "yes").length / rptSubs.length) * 1000) / 10 : 0;
  const rptDeptRanks = depts.map(d => {
    const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id);
    const rates = members.map(u => {
      const kd = memberData[u.id] || { krs: [] };
      if (!kd.krs.some(kr => rptSubs.some(s => s.memberId === u.id && s.krId === kr.id))) return null;
      return calcMemberRate(u.id, kd.krs, rptSubs);
    }).filter(r => r !== null);
    const rate = rates.length ? rates.reduce((a, b) => a + b, 0) / rates.length : 0;
    return { name: d.name, rate, status: getStatus(rate) };
  }).sort((a, b) => b.rate - a.rate);
  const rptCompRate = rptDeptRanks.length ? rptDeptRanks.reduce((a, d) => a + d.rate, 0) / rptDeptRanks.length : 0;
  const allMembers = users
    .filter(u => u.role === "member" || u.role === "manager")
    .map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const r = calcMemberRate(u.id, kd.krs, ovSubs);
      const hasData = kd.krs.some(kr => ovSubs.some(s => s.memberId === u.id && s.krId === kr.id));
      const dept = depts.find(d => d.id === u.deptId);
      const deptName = dept?.name || "—";
      const primaryTeam = dept?.teams.find(t => t.id === u.teamId);
      const secondTeam = u.secondTeamId ? dept?.teams.find(t => t.id === u.secondTeamId) : null;
      const teamName = primaryTeam ? (secondTeam ? `${primaryTeam.name} / ${secondTeam.name}` : primaryTeam.name) : "—";
      return { ...u, deptName, teamName, rate: r, hasData, status: getStatus(r) };
    })
    .sort((a, b) => b.rate - a.rate);

  function triggerSyncPrompt(deptId, teamId) {
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    const dept = depts.find(d => d.id === deptId);
    const team = dept?.teams.find(t => t.id === teamId);
    if (!team) return;
    setDirtySync({ deptId, teamId });
    syncTimerRef.current = setTimeout(() => setSyncPrompt({ deptId, teamId, teamName: team.name }), 1500);
  }

  function doSync(deptId, teamId) {
    const dept = depts.find(d => d.id === deptId);
    const team = dept?.teams.find(t => t.id === teamId);
    if (!team) return;
    const count = new Set([...(team.members || []), ...users.filter(u => u.teamId === teamId).map(u => u.id)]).size;
    dispatch({ type: "SYNC_TEAM_KRS_TO_MEMBERS", deptId, teamId });
    setDirtySync(null);
    if (syncNoteTimer.current) clearTimeout(syncNoteTimer.current);
    setSyncNote({ teamName: team.name, count });
    syncNoteTimer.current = setTimeout(() => setSyncNote(null), 3500);
  }

  function doDeptSync(deptId) {
    const dept = depts.find(d => d.id === deptId);
    if (!dept) return;
    const count = users.filter(u => u.deptId === deptId && (u.role === "member" || u.role === "manager")).length;
    dispatch({ type: "SYNC_DEPT_KRS_TO_MEMBERS", deptId });
    if (syncNoteTimer.current) clearTimeout(syncNoteTimer.current);
    setSyncNote({ teamName: `${dept.name} (all members)`, count });
    syncNoteTimer.current = setTimeout(() => setSyncNote(null), 3500);
  }

  function addKr(deptId, teamId) {
    if (!newKr.label) return;
    if (newKr.krType !== "tracker" && !newKr.useMonthlyTargets && Number(newKr.target) <= 0) return;
    const newId = `N${Date.now().toString(36).slice(-4).toUpperCase()}`;
    const baseKr = { id: newId, label: newKr.label, unit: newKr.unit.trim(), dataSource: newKr.dataSource.trim(), operator: newKr.operator || ">=", period: newKr.period || "monthly" };
    const kr = newKr.krType === "tracker"
      ? { ...baseKr, type: "tracker", target: 0, actual: 0 }
      : newKr.useMonthlyTargets
        ? { ...baseKr, monthlyTargets: Object.fromEntries(getFYMonths().map(m => [m.key, 0])), monthlyActuals: {}, ...(Number(newKr.dreamTarget) > 0 && { annualTarget: Number(newKr.dreamTarget) }) }
        : { ...baseKr, target: Number(newKr.target), actual: 0 };
    dispatch({ type: "ADD_KR", deptId, teamId, kr });
    if (teamId) triggerSyncPrompt(deptId, teamId);
    if (newKr.useMonthlyTargets) setExpandedMonthlyKr(newId);
    setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "" }); setAddTarget(null);
  }
  function startResize(key, e) {
    e.preventDefault();
    const startX = e.clientX;
    const startW = colWidthsRef.current[key] ?? 100;
    function onMove(me) {
      const w = Math.max(50, startW + me.clientX - startX);
      colWidthsRef.current = { ...colWidthsRef.current, [key]: w };
      setColWidths(prev => ({ ...prev, [key]: w }));
    }
    function onUp() { document.removeEventListener("mousemove", onMove); document.removeEventListener("mouseup", onUp); }
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }
  function startResizeCustom(colId, deptId, e) {
    e.preventDefault();
    const startX = e.clientX;
    const dept = depts.find(d => d.id === deptId);
    const startW = (dept?.customCols || []).find(c => c.id === colId)?.width ?? 150;
    function onMove(me) {
      setCustomColWidthOverride({ colId, width: Math.max(60, startW + me.clientX - startX) });
    }
    function onUp(me) {
      const finalW = Math.max(60, startW + me.clientX - startX);
      setCustomColWidthOverride(null);
      const current = depts.find(d => d.id === deptId)?.customCols || [];
      dispatch({ type: "SET_DEPT_CUSTOM_COLS", deptId, customCols: current.map(c => c.id === colId ? { ...c, width: finalW } : c) });
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup", onUp);
    }
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  }
  function addCustomCol() {
    if (!newColName.trim() || !selDept) return;
    const current = depts.find(d => d.id === selDept)?.customCols || [];
    dispatch({ type: "SET_DEPT_CUSTOM_COLS", deptId: selDept, customCols: [...current, { id: `col_${Date.now()}`, name: newColName.trim(), width: 150 }] });
    setNewColName(""); setAddingCol(false);
  }

  function resolveScopePool(scope) {
    let pool = users.filter(u => u.role === "member" || u.role === "manager");
    if (scope.userId) return pool.filter(u => u.id === scope.userId);
    if (scope.teamId) {
      const team = depts.flatMap(d => d.teams).find(t => t.id === scope.teamId);
      const set = new Set([...(team?.members || []), ...pool.filter(u => u.teamId === scope.teamId || u.secondTeamId === scope.teamId).map(u => u.id)]);
      return pool.filter(u => set.has(u.id));
    }
    if (scope.deptId) return pool.filter(u => u.deptId === scope.deptId);
    return pool;
  }

  function scopeLabel(scope) {
    if (scope.userId) { const u = users.find(u => u.id === scope.userId); return u?.name || "Unknown"; }
    if (scope.teamId) {
      const d = depts.find(d => d.teams.some(t => t.id === scope.teamId));
      const t = d?.teams.find(t => t.id === scope.teamId);
      return d && t ? `${d.name} › ${t.name}` : "Unknown Team";
    }
    if (scope.deptId) { const d = depts.find(d => d.id === scope.deptId); return d?.name || "Unknown Dept"; }
    return "All Departments";
  }

  function previewCheckin(period, scope = {}) {
    const periodKey = prevPeriodKey(period);
    const dateRange = periodDateRange(period, periodKey);
    const existing = new Set(okrSubmissions.filter(s => s.period === period && s.periodKey === periodKey).map(s => `${s.memberId}:${s.krId}`));
    const userPool = resolveScopePool(scope);
    const recipients = [];
    for (const u of userPool) {
      const dept = depts.find(d => d.id === u.deptId);
      if (!dept) continue;
      const krList = [];
      dept.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
      dept.teams.forEach(t => { if (t.members?.includes(u.id) || u.teamId === t.id || u.secondTeamId === t.id) t.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr)); });
      (memberData[u.id]?.krs || []).filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
      if (!krList.length) continue;
      const uniqueKrs = [...new Map(krList.map(kr => [kr.id, kr])).values()];
      const freshKrs = uniqueKrs.filter(kr => !existing.has(`${u.id}:${kr.id}`));
      if (!freshKrs.length) continue;
      recipients.push({ user: u, dept, krs: freshKrs });
    }
    setCheckinPreview({ period, scope, periodKey, dateRange, recipients });
  }

  async function sendCheckin(period, scope = {}) {
    setSendingCheckin(true);
    const periodKey = prevPeriodKey(period);
    const dateRange = periodDateRange(period, periodKey);
    const existing = new Set(okrSubmissions.filter(s => s.period === period && s.periodKey === periodKey).map(s => `${s.memberId}:${s.krId}`));
    const newSubs = [];
    const emailPromises = [];
    let ctr = Date.now();
    const userPool = resolveScopePool(scope);
    for (const u of userPool) {
      const dept = depts.find(d => d.id === u.deptId);
      if (!dept) continue;
      const krList = [];
      dept.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
      dept.teams.forEach(t => { if (t.members?.includes(u.id) || u.teamId === t.id || u.secondTeamId === t.id) t.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr)); });
      (memberData[u.id]?.krs || []).filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
      if (!krList.length) continue;
      const uniqueKrs = [...new Map(krList.map(kr => [kr.id, kr])).values()];
      const freshKrs = uniqueKrs.filter(kr => !existing.has(`${u.id}:${kr.id}`));
      const monthKey = period === "monthly" ? periodKey
        : period === "weekly" ? (() => { const d = new Date(Date.now() - 7 * 86400000); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; })()
        : currentFYMonthKey();
      const resolveTarget = kr => kr.monthlyTargets ? (kr.monthlyTargets[monthKey] ?? kr.target ?? 0) : (kr.target ?? 0);
      freshKrs.forEach(kr => { newSubs.push({ id: `os_${(ctr++).toString(36)}`, memberId: u.id, memberName: u.name, deptId: u.deptId, krId: kr.id, krLabel: kr.label, krTarget: resolveTarget(kr), krUnit: kr.unit || "", krOperator: kr.operator || ">=", krType: kr.type || "", period, periodKey, dateRange, sentAt: new Date().toISOString(), answeredAt: null, answer: null, approval: "pending", approvedBy: null }); });
      if (freshKrs.length && u.email) {
        const emailTemplates = settings?.emailTemplates || {};
        const template = { ...emailTemplates.default, ...(emailTemplates[period] || {}) };
        const krsForEmail = freshKrs.map(kr => ({ ...kr, target: resolveTarget(kr) }));
        emailPromises.push(
          fetch("/api/send-email", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ to: u.email, name: u.name, period, periodKey, dateRange, krs: krsForEmail, template }) })
            .then(res => res.ok ? null : { name: u.name, reason: `HTTP ${res.status}` })
            .catch(err => ({ name: u.name, reason: err.message || "Network error" }))
        );
      }
    }
    if (newSubs.length) dispatch({ type: "CREATE_OKR_SUBMISSIONS", submissions: newSubs });
    const memberCount = new Set(newSubs.map(s => s.memberId)).size;
    const emailOutcomes = await Promise.all(emailPromises);
    const emailFailures = emailOutcomes.filter(Boolean);
    setCheckinResult({ count: newSubs.length, memberCount, period, scope, emailFailures });
    setSendingCheckin(false);
    setTimeout(() => setCheckinResult(null), 8000);
  }

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      {checkinPreview && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div style={{ background: T.surface, borderRadius: 14, boxShadow: "0 8px 40px rgba(0,0,0,0.22)", width: "100%", maxWidth: 580, maxHeight: "80vh", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${T.border}` }}>
              <div style={{ fontWeight: 800, fontSize: 17 }}>Confirm: Send {checkinPreview.period.charAt(0).toUpperCase() + checkinPreview.period.slice(1)} Check-in</div>
              <div style={{ fontSize: 12, color: T.textDim, marginTop: 5 }}>
                <span>Period: <strong style={{ color: T.text }}>{checkinPreview.dateRange || checkinPreview.periodKey}</strong></span>
                <span style={{ margin: "0 6px", color: T.border }}>·</span>
                <span>To: <strong style={{ color: T.text }}>{scopeLabel(checkinPreview.scope || {})}</strong></span>
                <span style={{ margin: "0 6px", color: T.border }}>·</span>
                <span>{checkinPreview.recipients.length} recipient{checkinPreview.recipients.length !== 1 ? "s" : ""}</span>
                <span style={{ margin: "0 6px", color: T.border }}>·</span>
                <span>{checkinPreview.recipients.reduce((n, r) => n + r.krs.length, 0)} KRs total</span>
              </div>
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: "14px 24px" }}>
              {checkinPreview.recipients.length === 0 ? (
                <div style={{ padding: "28px 0", color: T.textDim, fontSize: 13, textAlign: "center" }}>
                  No new check-ins to send — all KRs for this period already have submissions.
                </div>
              ) : checkinPreview.recipients.map(({ user: u, dept, krs }) => (
                <div key={u.id} style={{ marginBottom: 10, padding: "10px 14px", borderRadius: 8, background: T.raised, border: `1px solid ${T.border}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 7, flexWrap: "wrap" }}>
                    <Avatar letters={u.av || "?"} size={26} />
                    <span style={{ fontWeight: 700, fontSize: 13 }}>{u.name}</span>
                    {u.role === "manager" && <span style={{ fontSize: 10, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Manager</span>}
                    <span style={{ fontSize: 11, color: T.textMuted, background: T.surface, borderRadius: 5, padding: "1px 7px", border: `1px solid ${T.border}` }}>{dept.name}</span>
                    {u.email && <span style={{ fontSize: 11, color: T.textDim }}>{u.email}</span>}
                  </div>
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                    {krs.map(kr => (
                      <span key={kr.id} style={{ fontSize: 11, background: kr.type === "tracker" ? "#ede9fe" : T.surface, color: kr.type === "tracker" ? "#7c3aed" : T.text, border: `1px solid ${kr.type === "tracker" ? "#c4b5fd" : T.border}`, borderRadius: 6, padding: "2px 8px" }}>
                        {kr.label}
                      </span>
                    ))}
                  </div>
                </div>
              ))}
            </div>
            <div style={{ padding: "14px 24px", borderTop: `1px solid ${T.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn onClick={() => setCheckinPreview(null)}>Cancel</Btn>
              <Btn primary disabled={checkinPreview.recipients.length === 0}
                onClick={() => { const { period, scope } = checkinPreview; setCheckinPreview(null); sendCheckin(period, scope); }}>
                📨 Confirm &amp; Send to {checkinPreview.recipients.length} recipient{checkinPreview.recipients.length !== 1 ? "s" : ""}
              </Btn>
            </div>
          </div>
        </div>
      )}
      <Side items={navItems} active={page} onSelect={p => { setPage(p); setSelDept(null); }} user={user} onLogout={onLogout}
        subItems={deptSubItems} subItemsFor="departments" activeSubItem={selDept || "__all__"} onSelectSubItem={id => { setPage("departments"); setSelDept(id === "__all__" ? null : id); setSelTeam(null); setAddTarget(null); }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "users" && <UserMgmtPage users={users} depts={depts} dispatch={dispatch} currentUserId={user.id} />}

        {page === "overview" && (<>
          <Header title="Company Overview" sub={(() => {
            const now = new Date();
            if (overviewView === "weekly") {
              const dow = now.getDay();
              const mon = new Date(now); mon.setHours(0,0,0,0); mon.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1));
              const sun = new Date(mon); sun.setDate(mon.getDate() + 6);
              const fmt = d => d.toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" });
              return `${fmt(mon)} – ${fmt(sun)} · All departments`;
            }
            if (overviewView === "monthly") return `${now.toLocaleDateString("en-AU", { month: "long", year: "numeric" })} · All departments`;
            const m = now.getMonth() + 1, y = now.getFullYear();
            return `FY${String(m >= 7 ? y + 1 : y).slice(2)} · All departments`;
          })()}
            right={<div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 12, color: T.textMuted, fontFamily: F.mono }}>Time: {TP}%</span><Tag type={getStatus(compRate)} /></div>} />
          <Pane>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {[["weekly","Weekly Overview"],["monthly","Monthly Overview"],["annual","Annual Overview"]].map(([v,label]) => (
                <Btn key={v} small primary={overviewView === v} onClick={() => setOverviewView(v)}>{label}</Btn>
              ))}
            </div>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label={overviewView === "weekly" ? "Weekly Completion" : overviewView === "annual" ? "Annual Completion" : "Monthly Completion"} value={`${compRate.toFixed(1)}%`} status={getStatus(compRate)} sub={`Target pace: ${TP}%`} />
              <Metric label="Departments" value={depts.length} />
              <Metric label="Teams" value={depts.reduce((a, d) => a + d.teams.length, 0)} />
              <Metric label="Staff Tracked" value={allMembers.length} />
            </div>
            <div>
              <SectionLabel>Department Rankings</SectionLabel>
              {ovSubs.length === 0
                ? <Card style={{ padding: "18px 20px", color: T.textMuted, textAlign: "center", fontSize: 14 }}>No check-in submissions for this period yet — send check-ins to see department rankings.</Card>
                : deptRanks.map((d, i) => (
                <Card key={d.id} onClick={() => { setSelDept(d.id); setPage("departments"); }} style={{ padding: "16px 20px", marginBottom: 8 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "36px 1fr 60px 180px 80px", alignItems: "center", gap: 14 }}>
                    <span style={{ fontSize: 18, fontWeight: 900, fontFamily: F.mono, color: i === 0 ? T.ok : i === deptRanks.length - 1 ? T.bad : T.textMuted }}>#{i + 1}</span>
                    <div><div style={{ fontSize: 16, fontWeight: 700 }}>{d.name}</div><div style={{ fontSize: 12, color: T.textMuted }}>{d.college} · {d.head} · {d.teams.length} teams</div></div>
                    <span style={{ textAlign: "right", fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[d.status].color }}>{d.rate.toFixed(1)}%</span>
                    <Bar value={d.rate} status={d.status} h={7} />
                    <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={d.status} /></div>
                  </div>
                </Card>
              ))}
            </div>
          </Pane>
        </>)}

        {page === "departments" && (!selDept ? (
          <DeptMgmtPage depts={depts} users={users} memberData={memberData} okrSubmissions={okrSubmissions} dispatch={dispatch} />
        ) : (() => {
          const dept = depts.find(d => d.id === selDept);
          if (!dept) return null;
          const COLS_DEF = [
            { key: "id",         label: "ID" },
            { key: "label",      label: "Key Result" },
            { key: "operator",   label: "Op" },
            { key: "period",     label: "Period" },
            { key: "target",     label: "Target" },
            { key: "actual",     label: "Actual" },
            { key: "unit",       label: "Unit" },
            { key: "dataSource", label: "Data Source" },
          ];
          const opSelect = (val, onChange) => (
            <select value={val} onChange={onChange} style={{ width: "100%", padding: "5px 4px", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontFamily: F.mono }}>
              <option value=">=">&gt;=</option><option value=">">&gt;</option><option value="<=">&lt;=</option><option value="<">&lt;</option><option value="=">=</option>
            </select>
          );
          const customCols = dept.customCols || [];
          const getCustomColWidth = col => customColWidthOverride?.colId === col.id ? customColWidthOverride.width : (col.width ?? 150);
          const orderedDef = colOrder.map(k => COLS_DEF.find(c => c.key === k)).filter(Boolean);
          const visibleBuiltIn = orderedDef.filter(c => !hiddenCols.has(c.key));
          const COL = [...visibleBuiltIn.map(c => `${colWidths[c.key]}px`), ...customCols.map(c => `${getCustomColWidth(c)}px`), "34px"].join(" ");
          const rszHandle = onMd => (
            <div onMouseDown={onMd} title="Drag to resize" style={{ width: 6, flexShrink: 0, alignSelf: "stretch", cursor: "col-resize", display: "flex", alignItems: "center", justifyContent: "center" }}>
              <div style={{ width: 2, height: "40%", background: T.border, borderRadius: 1 }} />
            </div>
          );

          const renderEditor = (krs, deptId, teamId, sectionPeriod) => {
            const onTeamChange = (krId, field, value) => { dispatch({ type: "UPDATE_KR", deptId, teamId, krId, field, value }); if (teamId) triggerSyncPrompt(deptId, teamId); };
            return (
              <Card style={{ overflow: "auto" }}>
                <div style={{ minWidth: "max-content" }}>
                  <div style={{ display: "grid", gridTemplateColumns: COL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase", alignItems: "center" }}>
                    {visibleBuiltIn.map(({ key, label }) => (
                      <div key={key} draggable
                        onDragStart={e => { dragColRef.current = key; e.dataTransfer.effectAllowed = "move"; }}
                        onDragOver={e => { e.preventDefault(); e.dataTransfer.dropEffect = "move"; }}
                        onDrop={e => { e.preventDefault(); const from = dragColRef.current; if (!from || from === key) return; const next = [...colOrder]; const fi = next.indexOf(from); const ti = next.indexOf(key); if (fi >= 0 && ti >= 0) { next.splice(fi, 1); next.splice(ti, 0, from); dispatch({ type: "SET_SETTINGS", updates: { colOrder: next } }); } dragColRef.current = null; }}
                        onDragEnd={() => { dragColRef.current = null; }}
                        style={{ display: "flex", alignItems: "center", minWidth: 0, gap: 2, cursor: "grab", userSelect: "none" }} title="Drag to reorder">
                        <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}</span>
                        <button onClick={() => setHiddenCols(prev => new Set([...prev, key]))} title="Hide" style={{ background: "none", border: "none", cursor: "pointer", color: T.textDim, fontSize: 9, padding: 0, lineHeight: 1, opacity: 0.6 }}>✕</button>
                        {rszHandle(e => startResize(key, e))}
                      </div>
                    ))}
                    {customCols.map(col => (
                      <div key={col.id} style={{ display: "flex", alignItems: "center", minWidth: 0, gap: 3 }}>
                        <input value={col.name} onChange={e => dispatch({ type: "SET_DEPT_CUSTOM_COLS", deptId: dept.id, customCols: customCols.map(c => c.id === col.id ? { ...c, name: e.target.value } : c) })}
                          style={{ flex: 1, minWidth: 0, background: "transparent", border: "none", outline: "none", fontFamily: "inherit", fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase", padding: 0, cursor: "text" }} />
                        <button onClick={() => dispatch({ type: "SET_DEPT_CUSTOM_COLS", deptId: dept.id, customCols: customCols.filter(c => c.id !== col.id) })} style={{ background: "none", border: "none", cursor: "pointer", color: T.textDim, fontSize: 9, padding: 0, lineHeight: 1 }}>✕</button>
                        {rszHandle(e => startResizeCustom(col.id, dept.id, e))}
                      </div>
                    ))}
                    <span />
                  </div>
                  {krs.map((kr, i) => {
                    const isMonthly = !!kr.monthlyTargets;
                    const curKey = currentFYMonthKey();
                    const curTarget = isMonthly ? (kr.monthlyTargets[curKey] || 0) : null;
                    const curActual = isMonthly ? ((kr.monthlyActuals || {})[curKey] || 0) : null;
                    const fyMs = isMonthly ? getFYMonths() : [];
                    const annSumTarget = fyMs.reduce((s, {key}) => s + (kr.monthlyTargets[key] || 0), 0);
                    const annActual = fyMs.reduce((s, {key}) => s + ((kr.monthlyActuals || {})[key] || 0), 0);
                    const annDream = isMonthly ? (kr.annualTarget || 0) : 0;
                    const annVsSum = annSumTarget > 0 ? Math.min((annActual / annSumTarget) * 100, 100) : 0;
                    const annVsDream = annDream > 0 ? Math.min((annActual / annDream) * 100, 100) : 0;
                    const annSt = (annDream > 0 ? annVsDream : annVsSum) >= 80 ? "green" : (annDream > 0 ? annVsDream : annVsSum) >= 50 ? "yellow" : "red";
                    return (
                      <Fragment key={kr.id}>
                      <div style={{ display: "grid", gridTemplateColumns: COL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                        {visibleBuiltIn.map(({ key }) => {
                          if (key === "id") return <span key="id" style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>;
                          if (key === "label") return <div key="label"><span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>{kr.type === "tracker" && <><span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", marginTop: 2, display: "inline-block" }}>Tracker · does not affect rate</span><label style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4, cursor: "pointer", userSelect: "none" }}><input type="checkbox" checked={kr.showInOverview !== false} onChange={e => onTeamChange(kr.id, "showInOverview", e.target.checked)} style={{ accentColor: "#7c3aed" }} /><span style={{ fontSize: 10, color: "#7c3aed" }}>Show in portals' OKR Overview</span></label></>}{isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", marginTop: 2, display: "inline-block" }}>Monthly Breakdown</span>}</div>;
                          if (key === "operator") return <span key="operator">{opSelect(kr.operator || ">=", e => onTeamChange(kr.id, "operator", e.target.value))}</span>;
                          if (key === "period") return <span key="period"><select value={kr.period || "monthly"} onChange={e => onTeamChange(kr.id, "period", e.target.value)} style={{ width: "100%", padding: "5px 4px", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontFamily: F.body }}><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="biannual">Biannual</option><option value="annual">Annual</option></select></span>;
                          if (key === "target") return kr.type === "tracker" ? <span key="target" style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : isMonthly ? <span key="target" style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: T.brand }}>{fmt(curTarget)} <span style={{ color: T.textDim }}>this mo.</span></span> : <Input key="target" value={kr.target} onChange={e => onTeamChange(kr.id, "target", Number(e.target.value) || 0)} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />;
                          if (key === "actual") return isMonthly ? <span key="actual" style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(curActual)}</span> : <span key="actual" style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.actual)}</span>;
                          if (key === "unit") return <Input key="unit" value={kr.unit || ""} onChange={e => onTeamChange(kr.id, "unit", e.target.value)} placeholder="e.g. %, students" style={{ padding: "5px 8px", fontSize: 13 }} />;
                          if (key === "dataSource") return <Input key="dataSource" value={kr.dataSource || ""} onChange={e => onTeamChange(kr.id, "dataSource", e.target.value)} placeholder="e.g. CRM, Manual" style={{ padding: "5px 8px", fontSize: 13 }} />;
                          return null;
                        })}
                        {customCols.map(col => <Input key={col.id} value={(kr.extras || {})[col.id] || ""} onChange={e => onTeamChange(kr.id, "extras", { ...(kr.extras || {}), [col.id]: e.target.value })} placeholder="—" style={{ padding: "5px 8px", fontSize: 13 }} />)}
                        <div style={{ display: "flex", gap: 4 }}>
                          {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="Edit monthly targets" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 12, fontWeight: 700 }}>📅</button>}
                          <button onClick={() => { dispatch({ type: "REMOVE_KR", deptId, teamId, krId: kr.id }); if (teamId) triggerSyncPrompt(deptId, teamId); }} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 8px", cursor: "pointer", color: T.bad, fontSize: 12, fontWeight: 700 }}>✕</button>
                        </div>
                      </div>
                      {isMonthly && expandedMonthlyKr === kr.id && (
                        <div style={{ padding: "14px 16px 18px", background: T.brandDim, borderBottom: `1px solid ${T.border}` }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: T.brand, marginBottom: 12 }}>KPI Breakdown — {kr.label}{kr.unit ? ` (${kr.unit})` : ""}</div>
                          <div style={{ overflowX: "auto" }}>
                            <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12, minWidth: 960 }}>
                              <thead><tr style={{ background: T.surface }}>
                                <th style={{ textAlign: "left", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, width: 110 }}></th>
                                {fyMs.map(({ key, label }) => { const isCur = key === curKey; return <th key={key} style={{ textAlign: "center", padding: "6px 3px", borderBottom: `2px solid ${isCur ? T.brand : T.border}`, fontSize: 11, fontWeight: isCur ? 700 : 400, color: isCur ? T.brand : T.textDim, minWidth: 62, background: isCur ? T.brandDim : T.surface }}>{label.split(" ")[0]}{isCur ? " ●" : ""}</th>; })}
                                <th style={{ textAlign: "right", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, minWidth: 80, background: T.surface }}>FY Total</th>
                                <th style={{ textAlign: "center", padding: "6px 10px", borderBottom: `2px solid ${T.okBorder}`, fontSize: 11, fontWeight: 700, color: T.ok, minWidth: 100, background: T.okDim }}>Dream Target</th>
                              </tr></thead>
                              <tbody>
                                <tr>
                                  <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Perf. Target</td>
                                  {fyMs.map(({ key }) => { const isCur = key === curKey; const t = kr.monthlyTargets[key] || 0; return <td key={key} style={{ padding: "3px 3px", background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}` }}><Input value={t} onChange={e => dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "target", value: Number(e.target.value) || 0 })} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} /></td>; })}
                                  <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annSumTarget)}</td>
                                  <td style={{ padding: "5px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, borderLeft: `1px solid ${T.okBorder}` }}>
                                    <Input value={annDream} onChange={e => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "annualTarget", value: Number(e.target.value) || 0 })} style={{ padding: "4px 6px", fontSize: 13, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box", fontWeight: 700, background: T.surface, border: `1px solid ${T.okBorder}`, borderRadius: 4 }} />
                                    {kr.unit && <div style={{ fontSize: 10, color: T.ok, textAlign: "center", marginTop: 2 }}>{kr.unit}</div>}
                                  </td>
                                </tr>
                                <tr>
                                  <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Actual</td>
                                  {fyMs.map(({ key }) => { const isCur = key === curKey; const a = (kr.monthlyActuals || {})[key] || 0; return <td key={key} style={{ padding: "3px 3px", background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}` }}><Input value={a} onChange={e => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "actual", value: Number(e.target.value) || 0 }); if (teamId) triggerSyncPrompt(deptId, teamId); }} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} /></td>; })}
                                  <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[annSt].color, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annActual)}</td>
                                  <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, textAlign: "center", color: T.textDim, fontSize: 12, borderLeft: `1px solid ${T.okBorder}` }}>—</td>
                                </tr>
                                <tr>
                                  <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 11, color: T.textDim, background: T.surface }}>Achievement</td>
                                  {fyMs.map(({ key }) => { const t = kr.monthlyTargets[key] || 0; const a = (kr.monthlyActuals || {})[key] || 0; const pct = t > 0 ? Math.min((a / t) * 100, 100) : null; const isCur = key === curKey; return <td key={key} style={{ padding: "4px 4px", textAlign: "center", background: isCur ? T.brandDim : "transparent" }}>{pct !== null ? <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[getStatus(pct)].color }}>{pct.toFixed(0)}%</span> : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}</td>; })}
                                  <td style={{ padding: "4px 10px", textAlign: "right", background: T.surface }}>{annSumTarget > 0 ? <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[annSt].color }}>{annVsSum.toFixed(0)}%</span><div style={{ fontSize: 10, color: T.textDim }}>vs. sum</div></> : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}</td>
                                  <td style={{ padding: "4px 8px", background: T.okDim, textAlign: "center", borderLeft: `1px solid ${T.okBorder}` }}>{annDream > 0 ? <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[getStatus(annVsDream)].color }}>{annVsDream.toFixed(0)}%</span><div style={{ fontSize: 10, color: T.ok }}>vs. dream</div></> : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}</td>
                                </tr>
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}
                      </Fragment>
                    );
                  })}
                  {addTarget === (teamId || `dept-${deptId}`) ? (
                    <div>
                      <div style={{ display: "grid", gridTemplateColumns: COL, padding: "9px 16px", gap: 8, alignItems: "center", background: T.brandDim }}>
                        {visibleBuiltIn.map(({ key }) => {
                          if (key === "id") return <span key="id" style={{ fontSize: 12, color: T.brand }}>NEW</span>;
                          if (key === "label") return <Input key="label" value={newKr.label} onChange={e => setNewKr(p => ({ ...p, label: e.target.value }))} placeholder="KR description *" style={{ padding: "5px 8px", fontSize: 14 }} />;
                          if (key === "operator") return <span key="operator">{opSelect(newKr.operator, e => setNewKr(p => ({ ...p, operator: e.target.value })))}</span>;
                          if (key === "period") return <span key="period"><select value={newKr.period} onChange={e => setNewKr(p => ({ ...p, period: e.target.value }))} style={{ width: "100%", padding: "5px 4px", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontFamily: F.body }}><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="biannual">Biannual</option><option value="annual">Annual</option></select></span>;
                          if (key === "target") return newKr.useMonthlyTargets ? <span key="target" style={{ fontSize: 11, color: T.brand, textAlign: "right" }}>Set per month ↓</span> : newKr.krType === "tracker" ? <span key="target" style={{ fontSize: 11, color: T.textMuted, textAlign: "right", fontStyle: "italic" }}>N/A</span> : <Input key="target" value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Target *" style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />;
                          if (key === "actual") return <span key="actual" />;
                          if (key === "unit") return <Input key="unit" value={newKr.unit} onChange={e => setNewKr(p => ({ ...p, unit: e.target.value }))} placeholder="Unit" style={{ padding: "5px 8px", fontSize: 13 }} />;
                          if (key === "dataSource") return <Input key="dataSource" value={newKr.dataSource} onChange={e => setNewKr(p => ({ ...p, dataSource: e.target.value }))} placeholder="Data source" style={{ padding: "5px 8px", fontSize: 13 }} />;
                          return null;
                        })}
                        {customCols.map(col => <span key={col.id} />)}
                        <button onClick={() => addKr(deptId, teamId)} style={{ background: T.brand, border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", color: "#fff", fontSize: 12, fontWeight: 700 }}>✓</button>
                      </div>
                      {sectionPeriod === "monthly" && newKr.krType !== "tracker" && (
                        <div style={{ padding: "8px 16px", background: T.brandDim, borderTop: `1px solid ${T.brandBorder}` }}>
                          <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, cursor: "pointer", color: T.brand, fontWeight: 600 }}>
                            <input type="checkbox" checked={newKr.useMonthlyTargets} onChange={e => setNewKr(p => ({ ...p, useMonthlyTargets: e.target.checked, target: e.target.checked ? "" : p.target }))} style={{ accentColor: T.brand }} />
                            Monthly Breakdown — set a different target for each month (Jul–Jun)
                          </label>
                          {newKr.useMonthlyTargets && (
                            <div style={{ display: "flex", alignItems: "flex-start", gap: 16, marginTop: 8 }}>
                              <div>
                                <div style={{ fontSize: 11, fontWeight: 700, color: T.ok, marginBottom: 4 }}>Dream Target <span style={{ fontWeight: 400, color: T.textMuted }}>(optional annual ceiling)</span></div>
                                <Input value={newKr.dreamTarget || ""} onChange={e => setNewKr(p => ({ ...p, dreamTarget: e.target.value }))} placeholder="e.g. 4997300" style={{ padding: "4px 8px", fontSize: 13, fontFamily: F.mono, textAlign: "right", width: 160 }} />
                                {newKr.unit && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>{newKr.unit}</div>}
                              </div>
                              <div style={{ fontSize: 11, color: T.brand, lineHeight: 1.6, paddingTop: 18 }}>Monthly targets open automatically after adding. When Dream Target is set, annual progress tracks against it instead of the sum of monthly targets.</div>
                            </div>
                          )}
                        </div>
                      )}
                      <div style={{ padding: "8px 16px", background: T.brandDim, borderTop: `1px solid ${T.brandBorder}` }}>
                        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, cursor: "pointer", color: "#7c3aed", fontWeight: 600 }}>
                          <input type="checkbox" checked={newKr.krType === "tracker"} onChange={e => setNewKr(p => ({ ...p, krType: e.target.checked ? "tracker" : "", useMonthlyTargets: false }))} style={{ accentColor: "#7c3aed" }} />
                          Tracker — record values only, does not affect completion rate
                        </label>
                      </div>
                    </div>
                  ) : (
                    <div style={{ padding: "10px 16px" }}>
                      <button onClick={() => { setAddTarget(teamId || `dept-${deptId}`); setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: sectionPeriod, useMonthlyTargets: false, krType: "" }); }} style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: 6, padding: "8px 14px", cursor: "pointer", color: T.brand, fontSize: 13, fontWeight: 600, width: "100%", fontFamily: F.body }}>+ Add Key Result</button>
                    </div>
                  )}
                  {hiddenCols.size > 0 && (
                    <div style={{ padding: "6px 16px", borderTop: `1px solid ${T.border}`, display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                      <span style={{ fontSize: 11, color: T.textMuted, fontWeight: 600 }}>Hidden:</span>
                      {[...hiddenCols].map(key => { const c = COLS_DEF.find(c => c.key === key); return c ? <button key={key} onClick={() => setHiddenCols(prev => { const n = new Set(prev); n.delete(key); return n; })} style={{ fontSize: 11, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "2px 8px", cursor: "pointer" }}>{c.label} +</button> : null; })}
                    </div>
                  )}
                </div>
              </Card>
            );
          };

          const renderSection = (sectionPeriod, sectionLabel, accentColor) => {
            const filterKrs = krs => krs.filter(kr => (kr.period || "monthly") === sectionPeriod);
            const deptKrs = filterKrs(dept.krs);
            return (
              <div key={sectionPeriod} style={{ marginBottom: 36 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, paddingBottom: 10, borderBottom: `2px solid ${accentColor}` }}>
                  <div style={{ width: 4, height: 22, background: accentColor, borderRadius: 2, flexShrink: 0 }} />
                  <div style={{ fontSize: 18, fontWeight: 800, color: T.text, letterSpacing: "-0.02em" }}>{sectionLabel}</div>
                  <div style={{ flex: 1 }} />
                  {addingCol ? (<>
                    <input autoFocus value={newColName} onChange={e => setNewColName(e.target.value)}
                      onKeyDown={e => { if (e.key === "Enter") addCustomCol(); if (e.key === "Escape") { setAddingCol(false); setNewColName(""); } }}
                      placeholder="Column name" style={{ padding: "5px 10px", fontSize: 13, border: `1px solid ${T.borderFocus}`, borderRadius: 6, background: T.surface, fontFamily: F.body, color: T.text, outline: "none" }} />
                    <Btn primary small onClick={addCustomCol}>Add</Btn>
                    <Btn small onClick={() => { setAddingCol(false); setNewColName(""); }}>Cancel</Btn>
                  </>) : <Btn small onClick={() => setAddingCol(true)}>+ Add Column</Btn>}
                </div>
                <div style={{ marginBottom: 14 }}>
                  <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", flex: 1 }}>{dept.name} — Dept OKRs</div>
                    {deptKrs.length > 0 && <Btn primary small onClick={() => doDeptSync(dept.id)}>⟳ Sync to All Members</Btn>}
                  </div>
                  {deptKrs.length === 0 && !addTarget && <div style={{ fontSize: 13, color: T.textMuted, padding: "6px 0 10px" }}>No {sectionLabel} for this department yet.</div>}
                  {renderEditor(deptKrs, dept.id, null, sectionPeriod)}
                </div>
                {dept.teams.length > 0 && (<>
                  <div style={{ fontSize: 12, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>Team KRs</div>
                  <div style={{ display: "flex", gap: 6, marginBottom: 10, flexWrap: "wrap", alignItems: "center" }}>
                    {dept.teams.map(t => <Btn key={t.id} primary={selTeam === t.id} small onClick={() => setSelTeam(p => p === t.id ? null : t.id)}>{t.name}</Btn>)}
                  </div>
                  {selTeam && (() => {
                    const team = dept.teams.find(t => t.id === selTeam);
                    if (!team) return null;
                    const teamKrs = filterKrs(team.krs);
                    return (<>
                      <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
                        <div style={{ fontSize: 13, color: T.textMuted, flex: 1 }}>Objective: {team.obj}{team.lead ? ` · Lead: ${team.lead}` : ""}</div>
                        {dirtySync?.deptId === dept.id && dirtySync?.teamId === team.id && <div style={{ fontSize: 12, color: T.warn, background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 6, padding: "3px 10px", fontWeight: 600 }}>⚠ Unsynced changes</div>}
                        <Btn primary small onClick={() => doSync(dept.id, team.id)}>⟳ Sync to Team Members</Btn>
                      </div>
                      {teamKrs.length === 0 && !addTarget && <div style={{ fontSize: 13, color: T.textMuted, padding: "6px 0" }}>No {sectionLabel} for {team.name} yet.</div>}
                      {renderEditor(teamKrs, dept.id, team.id, sectionPeriod)}
                    </>);
                  })()}
                </>)}
              </div>
            );
          };

          const renderPersonalSection = () => {
            const deptMembers = users.filter(u => u.deptId === dept.id && (u.role === "member" || u.role === "manager"))
              .map(u => ({ ...u, md: memberData[u.id] || { krs: [] } }));
            return (
              <div style={{ marginBottom: 36 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, paddingBottom: 10, borderBottom: `2px solid ${T.ok}` }}>
                  <div style={{ width: 4, height: 22, background: T.ok, borderRadius: 2, flexShrink: 0 }} />
                  <div style={{ fontSize: 18, fontWeight: 800, color: T.text, letterSpacing: "-0.02em" }}>Personal OKR</div>
                  <div style={{ fontSize: 12, color: T.textMuted, marginLeft: 4 }}>Individual targets per team member</div>
                </div>
                {deptMembers.length === 0
                  ? <EmptyState text="No members in this department yet. Add members via User Management." />
                  : deptMembers.map(member => {
                      const personalKrs = member.md.krs || [];
                      const isExpanded = expandedPersonalMember === member.id;
                      const rate = calcRate(personalKrs); const st = getStatus(rate);
                      return (
                        <Card key={member.id} style={{ marginBottom: 8, overflow: "hidden" }}>
                          <div onClick={() => setExpandedPersonalMember(p => p === member.id ? null : member.id)}
                            style={{ display: "flex", alignItems: "center", padding: "12px 16px", gap: 12, cursor: "pointer", userSelect: "none" }}>
                            <Avatar letters={member.av || member.name?.slice(0,2)} size={34} />
                            <div style={{ flex: 1 }}>
                              <div style={{ fontWeight: 700, fontSize: 14 }}>{member.name}</div>
                              <div style={{ fontSize: 12, color: T.textMuted }}>{personalKrs.length} personal KR{personalKrs.length !== 1 ? "s" : ""}{member.teamId ? ` · ${dept.teams.find(t => t.id === member.teamId)?.name || ""}` : ""}</div>
                            </div>
                            {personalKrs.length > 0 && <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 14, color: STATUS_THEME[st].color }}>{rate.toFixed(0)}%</span><div style={{ width: 80 }}><Bar value={rate} status={st} h={4} /></div><Tag type={st} small /></>}
                            <span style={{ fontSize: 11, color: T.textMuted }}>{isExpanded ? "▲" : "▼"}</span>
                          </div>
                          {isExpanded && (
                            <div style={{ borderTop: `1px solid ${T.border}`, padding: "14px 16px" }}>
                              {personalKrs.length > 0 && (
                                <div style={{ marginBottom: 10 }}>
                                  <div style={{ display: "grid", gridTemplateColumns: "1fr 90px 90px 80px 28px", gap: 8, padding: "5px 0 6px", fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", borderBottom: `1px solid ${T.border}` }}>
                                    <span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Unit</span><span />
                                  </div>
                                  {personalKrs.map((kr, i) => {
                                    const pct = krCompletion(kr); const st2 = getStatus(pct);
                                    return (
                                      <div key={kr.id} style={{ display: "grid", gridTemplateColumns: "1fr 90px 90px 80px 28px", gap: 8, padding: "6px 0", alignItems: "center", borderBottom: `1px solid ${T.border}`, background: i % 2 ? T.raised : "transparent", fontSize: 13 }}>
                                        <div><span>{kr.label}</span><span style={{ marginLeft: 8, fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[st2].color }}>{pct.toFixed(0)}%</span></div>
                                        <Input value={kr.target || 0} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "target", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "3px 6px", fontFamily: F.mono, fontSize: 13 }} />
                                        <Input value={kr.actual || 0} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "3px 6px", fontFamily: F.mono, fontSize: 13 }} />
                                        <span style={{ fontSize: 12, color: T.textMuted }}>{kr.unit || "—"}</span>
                                        <button onClick={() => dispatch({ type: "REMOVE_MEMBER_KR", memberId: member.id, krId: kr.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 4, padding: "2px 6px", cursor: "pointer", color: T.bad, fontSize: 11 }}>✕</button>
                                      </div>
                                    );
                                  })}
                                </div>
                              )}
                              {addPersonalKr?.memberId === member.id ? (
                                <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 10px", background: T.brandDim, borderRadius: 7 }}>
                                  <Input value={addPersonalKr.label} onChange={e => setAddPersonalKr(p => ({...p, label: e.target.value}))} placeholder="KR description *" style={{ flex: 1, padding: "4px 8px", fontSize: 13 }} />
                                  <Input value={addPersonalKr.target} onChange={e => setAddPersonalKr(p => ({...p, target: e.target.value}))} placeholder="Target" style={{ width: 80, textAlign: "right", padding: "4px 6px", fontFamily: F.mono, fontSize: 13 }} />
                                  <Input value={addPersonalKr.unit} onChange={e => setAddPersonalKr(p => ({...p, unit: e.target.value}))} placeholder="Unit" style={{ width: 70, padding: "4px 6px", fontSize: 13 }} />
                                  <button onClick={() => { if (!addPersonalKr.label.trim()) return; dispatch({ type: "ADD_MEMBER_KR", memberId: member.id, kr: { id: `P${Date.now().toString(36).slice(-4).toUpperCase()}`, label: addPersonalKr.label.trim(), target: Number(addPersonalKr.target) || 0, actual: 0, unit: addPersonalKr.unit.trim(), operator: ">=", period: "monthly" } }); setAddPersonalKr(null); }} style={{ background: T.brand, border: "none", borderRadius: 5, padding: "4px 10px", cursor: "pointer", color: "#fff", fontSize: 12, fontWeight: 700 }}>✓</button>
                                  <button onClick={() => setAddPersonalKr(null)} style={{ background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 8px", cursor: "pointer", color: T.text, fontSize: 12 }}>✕</button>
                                </div>
                              ) : (
                                <button onClick={() => setAddPersonalKr({ memberId: member.id, label: "", target: "", unit: "" })}
                                  style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: 6, padding: "7px 14px", cursor: "pointer", color: T.brand, fontSize: 13, fontWeight: 600, width: "100%", fontFamily: F.body }}>+ Add Personal KR</button>
                              )}
                            </div>
                          )}
                        </Card>
                      );
                    })}
              </div>
            );
          };

          const sections = [
            { period: "daily",     label: "Daily OKR",      color: T.warn    },
            { period: "weekly",    label: "Weekly OKR",     color: T.brand   },
            { period: "monthly",   label: "Monthly OKR",    color: "#6B7280" },
            { period: "quarterly", label: "Quarterly OKR",  color: "#F97316" },
            { period: "biannual",  label: "Bi-Annual OKR",  color: "#06B6D4" },
            { period: "annual",    label: "Annual OKR",     color: "#A78BFA" },
          ];

          return (<>
            <Header title={dept.name} sub={[dept.college, dept.obj].filter(Boolean).join(" · ")}
              right={<div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                {syncNote && <div style={{ fontSize: 12, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "3px 10px", fontWeight: 600 }}>✓ Synced {syncNote.teamName} → {syncNote.count} members</div>}
                <Btn small onClick={() => setSelDept(null)}>← All Departments</Btn>
              </div>}
            />
            <Pane>
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 24, paddingBottom: 14, borderBottom: `1px solid ${T.border}` }}>
                {depts.map(d => <Btn key={d.id} primary={d.id === selDept} small onClick={() => { setSelDept(d.id); setSelTeam(null); setAddTarget(null); }}>{d.name}</Btn>)}
              </div>
              {sections.map(({ period, label, color }) => renderSection(period, label, color))}
              {renderPersonalSection()}
            </Pane>
          </>);
        })())}

        {page === "submissions" && (() => {
          const PERIOD_TABS = [
            { id: "daily", label: "Daily", color: T.warn },
            { id: "weekly", label: "Weekly", color: T.brand },
            { id: "monthly", label: "Monthly", color: "#A78BFA" },
            { id: "quarterly", label: "Quarterly", color: "#F97316" },
            { id: "biannual",  label: "Bi-Annual",  color: "#06B6D4" },
            { id: "annual", label: "Annual", color: T.ok },
          ];
          const periodSubs = okrSubmissions.filter(s => s.period === subPeriod);
          const totalPending = okrSubmissions.filter(s => s.answer !== null && s.approval === "pending").length;
          const q = subSearch.trim().toLowerCase();
          const filtered = periodSubs.filter(s => {
            const mem = users.find(u => u.id === s.memberId);
            if (subDeptFilter !== "all" && mem?.deptId !== subDeptFilter) return false;
            if (q && !mem?.name?.toLowerCase().includes(q) && !(s.krLabel || "").toLowerCase().includes(q)) return false;
            if (subFilter === "unanswered") return s.answer === null;
            if (subFilter === "yes") return s.answer === "yes";
            if (subFilter === "no") return s.answer === "no";
            if (subFilter === "pending") return s.answer !== null && s.approval === "pending";
            if (subFilter === "approved") return s.approval === "approved";
            return true;
          }).sort((a, b) => (b.sentAt || "").localeCompare(a.sentAt || ""));
          return (<>
            <Header title="OKR Check-in Submissions" sub="Staff respond to emailed yes/no KPI check-ins — managers approve in their portal"
              right={totalPending > 0 ? <div style={{ fontSize: 12, color: T.warn, background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 6, padding: "3px 10px", fontWeight: 600 }}>{totalPending} awaiting approval</div> : null} />
            <Pane>
              {/* Period tabs */}
              <div style={{ display: "flex", gap: 0, marginBottom: 20, borderBottom: `2px solid ${T.border}` }}>
                {PERIOD_TABS.map(p => {
                  const cnt = okrSubmissions.filter(s => s.period === p.id && s.answer !== null && s.approval === "pending").length;
                  return (
                    <button key={p.id} onClick={() => setSubPeriod(p.id)} style={{ padding: "8px 22px", fontSize: 13, fontWeight: 600, fontFamily: F.body, cursor: "pointer", background: "none", border: "none", borderBottom: subPeriod === p.id ? `3px solid ${p.color}` : "3px solid transparent", color: subPeriod === p.id ? p.color : T.textMuted, display: "flex", alignItems: "center", gap: 6 }}>
                      {p.label}{cnt > 0 ? <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 6px", fontSize: 10, fontWeight: 700 }}>{cnt}</span> : null}
                    </button>
                  );
                })}
              </div>
              {/* Metrics + Send button */}
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 12, marginBottom: 18 }}>
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap" }}>
                  <Metric label="Total" value={periodSubs.length} />
                  <Metric label="Unanswered" value={periodSubs.filter(s => s.answer === null).length} status="yellow" />
                  <Metric label="Yes" value={periodSubs.filter(s => s.answer === "yes").length} status="green" />
                  <Metric label="No" value={periodSubs.filter(s => s.answer === "no").length} status="red" />
                  <Metric label="Approved" value={periodSubs.filter(s => s.approval === "approved").length} status="green" />
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                  {/* Dept */}
                  <select value={checkinScope.deptId}
                    onChange={e => setCheckinScope({ deptId: e.target.value, teamId: "", userId: "" })}
                    style={{ padding: "6px 10px", fontSize: 13, fontFamily: F.body, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none", maxWidth: 180 }}>
                    <option value="">All Departments</option>
                    {depts.slice().sort((a,b) => a.name.localeCompare(b.name)).map(d => (
                      <option key={d.id} value={d.id}>{d.name}</option>
                    ))}
                  </select>
                  {/* Team — only if dept selected */}
                  {checkinScope.deptId && (() => {
                    const scopeTeams = depts.find(d => d.id === checkinScope.deptId)?.teams || [];
                    return scopeTeams.length > 0 ? (
                      <select value={checkinScope.teamId}
                        onChange={e => setCheckinScope(p => ({ ...p, teamId: e.target.value, userId: "" }))}
                        style={{ padding: "6px 10px", fontSize: 13, fontFamily: F.body, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none", maxWidth: 160 }}>
                        <option value="">All Teams</option>
                        {scopeTeams.slice().sort((a,b) => a.name.localeCompare(b.name)).map(t => (
                          <option key={t.id} value={t.id}>{t.name}</option>
                        ))}
                      </select>
                    ) : null;
                  })()}
                  {/* Member */}
                  {(() => {
                    const pool = resolveScopePool({ deptId: checkinScope.deptId, teamId: checkinScope.teamId });
                    return (
                      <select value={checkinScope.userId}
                        onChange={e => setCheckinScope(p => ({ ...p, userId: e.target.value }))}
                        style={{ padding: "6px 10px", fontSize: 13, fontFamily: F.body, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none", maxWidth: 180 }}>
                        <option value="">All Members</option>
                        {pool.slice().sort((a,b) => a.name.localeCompare(b.name)).map(u => (
                          <option key={u.id} value={u.id}>{u.name} {u.deptId !== checkinScope.deptId ? `(${depts.find(d=>d.id===u.deptId)?.name||""})` : ""}</option>
                        ))}
                      </select>
                    );
                  })()}
                  <Btn primary onClick={() => previewCheckin(subPeriod, checkinScope)} disabled={sendingCheckin}>
                    {sendingCheckin ? "Sending…" : `📨 Send ${subPeriod.charAt(0).toUpperCase() + subPeriod.slice(1)} Check-in`}
                  </Btn>
                </div>
              </div>
              {checkinResult && (
                <div style={{ marginBottom: 14, padding: "10px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600,
                  background: checkinResult.emailFailures?.length > 0 ? T.warnDim : checkinResult.count > 0 ? T.okDim : T.warnDim,
                  border: `1px solid ${checkinResult.emailFailures?.length > 0 ? T.warnBorder : checkinResult.count > 0 ? T.okBorder : T.warnBorder}`,
                  color: checkinResult.emailFailures?.length > 0 ? T.warn : checkinResult.count > 0 ? T.ok : T.warn }}>
                  {checkinResult.count > 0
                    ? <div>
                        <div>{`✓ Created ${checkinResult.count} check-in${checkinResult.count !== 1 ? "s" : ""} for ${checkinResult.memberCount} member${checkinResult.memberCount !== 1 ? "s" : ""} — ${scopeLabel(checkinResult.scope || {})}`}</div>
                        {checkinResult.emailFailures?.length > 0 && (
                          <div style={{ marginTop: 5, fontWeight: 500 }}>
                            {`⚠ ${checkinResult.emailFailures.length} email${checkinResult.emailFailures.length !== 1 ? "s" : ""} failed to send: ${checkinResult.emailFailures.map(f => f.name).join(", ")}`}
                          </div>
                        )}
                      </div>
                    : `⚠ No submissions created — no KRs found for this period. Check that KRs are configured with the "${checkinResult.period}" period.`}
                </div>
              )}
              {/* Filter bar */}
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 14 }}>
                <div style={{ display: "flex", gap: 6 }}>
                  {[["all","All"],["unanswered","Unanswered"],["yes","Yes"],["no","No"],["pending","Needs Approval"],["approved","Approved"]].map(([f,l]) => (
                    <Btn key={f} small primary={subFilter === f} onClick={() => setSubFilter(f)}>{l}</Btn>
                  ))}
                </div>
                <div style={{ width: 1, height: 22, background: T.border, flexShrink: 0 }} />
                <div style={{ position: "relative", flex: "0 0 200px" }}>
                  <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: T.textDim, fontSize: 14, pointerEvents: "none" }}>⌕</span>
                  <input value={subSearch} onChange={e => setSubSearch(e.target.value)} placeholder="Search name or KR..."
                    style={{ width: "100%", boxSizing: "border-box", paddingLeft: 28, paddingRight: 10, paddingTop: 6, paddingBottom: 6, fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, fontFamily: F.body, outline: "none" }} />
                </div>
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                  <Btn small primary={subDeptFilter === "all"} onClick={() => setSubDeptFilter("all")}>All Depts</Btn>
                  {depts.map(d => <Btn key={d.id} small primary={subDeptFilter === d.id} onClick={() => setSubDeptFilter(d.id)}>{d.name}</Btn>)}
                </div>
              </div>
              {(() => {
                const pendingMgrSubs = periodSubs.filter(s => {
                  const m = users.find(u => u.id === s.memberId);
                  return m?.role === "manager" && s.answer !== null && s.approval === "pending";
                });
                return pendingMgrSubs.length > 0 ? (
                  <div style={{ marginBottom: 14, padding: "10px 14px", borderRadius: 8, fontSize: 13, fontWeight: 600, background: "#ede9fe", border: "1px solid #c4b5fd", color: "#6d28d9", display: "flex", alignItems: "center", gap: 8 }}>
                    <span>⚑</span>
                    <span>{pendingMgrSubs.length} manager submission{pendingMgrSubs.length !== 1 ? "s" : ""} awaiting your approval — these are not visible to any Manager and can only be approved here</span>
                  </div>
                ) : null;
              })()}
              {filtered.length === 0 && <EmptyState text={periodSubs.length === 0 ? `No ${subPeriod} check-ins sent yet. Click "Send ${subPeriod.charAt(0).toUpperCase()+subPeriod.slice(1)} Check-in" to generate and email them.` : "No submissions match your filter."} />}
              {filtered.map(s => {
                const mem = users.find(u => u.id === s.memberId);
                const dept = depts.find(d => d.id === s.deptId);
                const accentColor = s.approval === "approved" ? T.ok : s.answer === "yes" ? T.ok : s.answer === "no" ? T.bad : s.answer === null ? T.warn : T.border;
                return (
                  <Card key={s.id} style={{ padding: "14px 18px", borderLeft: `3px solid ${accentColor}`, marginBottom: 6 }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                          <Avatar letters={mem?.av || "?"} size={26} />
                          <span style={{ fontWeight: 700, fontSize: 14 }}>{mem?.name || s.memberName || "Unknown"}</span>
                          {mem?.role === "manager" && <span style={{ fontSize: 11, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 6, padding: "1px 6px", fontWeight: 700 }}>Manager</span>}
                          {dept && <span style={{ fontSize: 11, color: T.textMuted, background: T.raised, borderRadius: 6, padding: "1px 6px" }}>{dept.name}</span>}
                          {s.krType === "tracker" && <span style={{ fontSize: 11, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 6, padding: "1px 6px", fontWeight: 700 }}>Tracker · does not affect rate</span>}
                        </div>
                        <div style={{ fontSize: 14, color: T.text, marginBottom: 4, fontWeight: 600 }}>{s.krLabel}</div>
                        <div style={{ display: "flex", alignItems: "baseline", gap: 6, marginBottom: 3 }}>
                          <span style={{ fontSize: 22, fontWeight: 800, fontFamily: F.mono, color: T.text, lineHeight: 1 }}>{s.krTarget ?? "—"}</span>
                          {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                          <span style={{ fontSize: 11, color: T.textDim }}>target</span>
                          {s.answer === "no" && s.actualValue != null && <><span style={{ fontSize: 11, color: T.textDim, marginLeft: 8 }}>·</span><span style={{ fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: T.bad, lineHeight: 1, marginLeft: 8 }}>{s.actualValue}</span><span style={{ fontSize: 11, color: T.bad }}>actual</span></>}
                          {s.answer === "yes" && <><span style={{ fontSize: 11, color: T.textDim, marginLeft: 8 }}>·</span><span style={{ fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: T.ok, lineHeight: 1, marginLeft: 8 }}>{s.krTarget}</span><span style={{ fontSize: 11, color: T.ok }}>actual</span></>}
                        </div>
                        <div style={{ fontSize: 12, color: T.textMuted }}>
                          {periodDisplayLabel(s.period, s.periodKey)} · Sent: {s.sentAt?.slice(0,10) || "—"}
                        </div>
                        {s.answer === "no" && s.reason && <div style={{ fontSize: 12, color: T.bad, marginTop: 3, fontStyle: "italic" }}>Reason: {s.reason}</div>}
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                        {s.answer === null
                          ? <span style={{ fontSize: 12, color: T.textMuted, background: T.raised, borderRadius: 6, padding: "3px 8px" }}>{s.krType === "tracker" ? "Awaiting record" : "Awaiting answer"}</span>
                          : s.krType === "tracker"
                            ? <span style={{ fontSize: 12, fontWeight: 700, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 6, padding: "3px 8px" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                            : <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad, background: s.answer === "yes" ? T.okDim : T.badDim, border: `1px solid ${s.answer === "yes" ? T.okBorder : T.badBorder}`, borderRadius: 6, padding: "3px 8px" }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>}
                        {s.answer !== null && s.approval === "pending"
                          ? <div style={{ display: "flex", gap: 6 }}>
                              <Btn danger small onClick={() => setRejectOkr({ id: s.id, actual: "" })}>Reject</Btn>
                              <Btn primary small onClick={() => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id })}>Approve</Btn>
                            </div>
                          : s.approval !== "pending" && <Tag type={s.approval === "approved" ? "approved" : "rejected"} label={s.approval === "approved" ? "Approved" : "Rejected"} small />}
                        <button onClick={() => dispatch({ type: "REMOVE_OKR_SUBMISSION", id: s.id })} title="Delete" style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 14, lineHeight: 1, padding: "2px 4px" }}>✕</button>
                      </div>
                    </div>
                    {rejectOkr?.id === s.id && (
                      <div style={{ marginTop: 10, padding: "10px 12px", background: T.badDim, borderRadius: 7, border: `1px solid ${T.badBorder}` }}>
                        <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 6 }}>Enter actual value for rejection</div>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                          <Input value={rejectOkr.actual} onChange={e => setRejectOkr(p => ({ ...p, actual: e.target.value }))} placeholder="Actual value" style={{ width: 120, textAlign: "right", fontFamily: F.mono }} />
                          {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                          <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                        </div>
                        <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                          <Btn small onClick={() => setRejectOkr(null)}>Cancel</Btn>
                          <Btn danger small onClick={() => { dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "rejected", approvedBy: user.id, actualValue: Number(rejectOkr.actual) || 0 }); setRejectOkr(null); }}>Confirm Reject</Btn>
                        </div>
                      </div>
                    )}
                  </Card>
                );
              })}
            </Pane>
          </>);
        })()}

        {page === "reports" && (<>
          <Header title="OKR Reports" sub="Published reports visible to ALL teams across the company"
            right={<div style={{ display: "flex", gap: 8 }}>
              <Btn onClick={() => { setShowGenReport(v => !v); setGenPeriod({ label: "", from: "", to: "" }); }}>{showGenReport ? "Cancel" : "Generate for Period"}</Btn>
              <Btn primary onClick={() => {
                if (state.monthlyReports.some(r => r.month === currentMonth())) {
                  if (!window.confirm(`A report for ${currentMonth()} already exists. Publish another?`)) return;
                }
                const rptMembers = users.filter(u => u.role === "member" || u.role === "manager").map(u => {
                  const kd = memberData[u.id] || { krs: [] };
                  const hasData = kd.krs.some(kr => rptSubs.some(s => s.memberId === u.id && s.krId === kr.id));
                  const rate = hasData ? calcMemberRate(u.id, kd.krs, rptSubs) : 0;
                  return { ...u, rate, hasData, status: getStatus(rate) };
                }).sort((a, b) => b.rate - a.rate);
                const report = { id: `mr${Date.now()}`, month: currentMonth(), publishedDate: new Date().toISOString().slice(0, 10), publishedBy: user.id,
                  reportType: "monthly",
                  notes: "",
                  submissionRate: rptSubRate,
                  data: { companyRate: rptCompRate, deptRanks: rptDeptRanks,
                    topPerformers: rptMembers.filter(m => m.hasData).slice(0, 3).map(m => `${m.name} — ${m.rate.toFixed(1)}%`),
                    redFlags: rptMembers.filter(m => m.hasData && m.status === "red").map(m => `${m.name} — ${m.rate.toFixed(1)}% (action required)`),
                  },
                };
                dispatch({ type: "PUBLISH_REPORT", report });
              }}>Publish {currentMonth()} Report</Btn>
            </div>} />
          <Pane>
            {showGenReport && (
              <Card style={{ padding: 20, borderLeft: `3px solid ${T.brand}` }}>
                <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 14 }}>Generate Monthly Report for Specific Period</div>
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
                  <div>
                    <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 4 }}>Period Label</div>
                    <Input value={genPeriod.label} onChange={e => setGenPeriod(p => ({ ...p, label: e.target.value }))} placeholder="e.g. Q1 FY2026, Week 22" style={{ width: 200 }} />
                  </div>
                  <div>
                    <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 4 }}>From</div>
                    <Input type="date" value={genPeriod.from} onChange={e => setGenPeriod(p => ({ ...p, from: e.target.value }))} style={{ width: 150 }} />
                  </div>
                  <div>
                    <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 4 }}>To</div>
                    <Input type="date" value={genPeriod.to} onChange={e => setGenPeriod(p => ({ ...p, to: e.target.value }))} style={{ width: 150 }} />
                  </div>
                  <Btn primary disabled={!genPeriod.label} onClick={() => {
                    const label = genPeriod.label.trim();
                    if (state.monthlyReports.some(r => r.month === label)) {
                      if (!window.confirm(`A report labelled "${label}" already exists. Publish another?`)) return;
                    }
                    const allAnswered = (state.okrSubmissions || []).filter(s => {
                      if (s.answer === null) return false;
                      if (genPeriod.from && s.sentAt && s.sentAt < genPeriod.from) return false;
                      if (genPeriod.to && s.sentAt && s.sentAt > genPeriod.to + "T23:59:59") return false;
                      return true;
                    });
                    const gSubRate = allAnswered.length > 0 ? Math.round((allAnswered.filter(s => s.answer === "yes").length / allAnswered.length) * 1000) / 10 : 0;
                    const gDeptRanks = depts.map(d => {
                      const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id);
                      const rates = members.map(u => {
                        const kd = memberData[u.id] || { krs: [] };
                        if (!kd.krs.some(kr => allAnswered.some(s => s.memberId === u.id && s.krId === kr.id))) return null;
                        return calcMemberRate(u.id, kd.krs, allAnswered);
                      }).filter(r => r !== null);
                      const rate = rates.length ? rates.reduce((a, b) => a + b, 0) / rates.length : 0;
                      return { name: d.name, rate, status: getStatus(rate) };
                    }).sort((a, b) => b.rate - a.rate);
                    const gCompRate = gDeptRanks.length ? gDeptRanks.reduce((a, d) => a + d.rate, 0) / gDeptRanks.length : 0;
                    const gMembers = users.filter(u => u.role === "member" || u.role === "manager").map(u => {
                      const kd = memberData[u.id] || { krs: [] };
                      const hasData = kd.krs.some(kr => allAnswered.some(s => s.memberId === u.id && s.krId === kr.id));
                      const rate = hasData ? calcMemberRate(u.id, kd.krs, allAnswered) : 0;
                      return { ...u, rate, hasData, status: getStatus(rate) };
                    }).sort((a, b) => b.rate - a.rate);
                    const report = {
                      id: `mr${Date.now()}`,
                      month: label,
                      reportType: "monthly",
                      publishedDate: new Date().toISOString().slice(0, 10),
                      publishedBy: user.id,
                      periodFrom: genPeriod.from || null,
                      periodTo: genPeriod.to || null,
                      notes: "",
                      submissionRate: gSubRate,
                      data: {
                        companyRate: gCompRate,
                        deptRanks: gDeptRanks,
                        topPerformers: gMembers.filter(m => m.hasData).slice(0, 3).map(m => `${m.name} — ${m.rate.toFixed(1)}%`),
                        redFlags: gMembers.filter(m => m.hasData && m.status === "red").map(m => `${m.name} — ${m.rate.toFixed(1)}% (action required)`),
                      },
                    };
                    dispatch({ type: "PUBLISH_REPORT", report });
                    setShowGenReport(false);
                    setGenPeriod({ label: "", from: "", to: "" });
                  }}>Generate & Publish</Btn>
                </div>
              </Card>
            )}
            <Card style={{ padding: "14px 18px", background: T.brandDim, border: `1px solid ${T.brandBorder}`, marginBottom: 4 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: T.brand, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.07em" }}>Current Data Preview — what will be published</div>
              <div style={{ display: "flex", gap: 20, flexWrap: "wrap", fontSize: 13 }}>
                <div><span style={{ color: T.textMuted }}>Submission rate ({currentMonth()}): </span><strong style={{ color: STATUS_THEME[getStatus(rptSubRate)].color }}>{rptSubRate}%</strong><span style={{ color: T.textMuted, fontSize: 11, marginLeft: 6 }}>({rptSubs.length} answered)</span></div>
                <div><span style={{ color: T.textMuted }}>Top performers: </span>{allMembers.slice(0, 3).map(m => m.name).join(", ") || "—"}</div>
                <div><span style={{ color: T.textMuted }}>Needs attention: </span>{allMembers.filter(m => m.status === "red").length > 0 ? allMembers.filter(m => m.status === "red").map(m => m.name).join(", ") : "None"}</div>
              </div>
            </Card>
            {(() => {
              const visibleReports = state.monthlyReports;
              if (visibleReports.length === 0) return <EmptyState text="No OKR reports published yet." />;
              return visibleReports.map(r => (
                <Card key={r.id} style={{ overflow: "hidden" }}>
                  <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ fontSize: 18, fontWeight: 800 }}>{r.month}</div>
                        <span style={{ fontSize: 10, fontWeight: 700, background: r.reportType === "weekly" ? T.brandDim : T.okDim, border: `1px solid ${r.reportType === "weekly" ? T.brandBorder : T.okBorder}`, color: r.reportType === "weekly" ? T.brand : T.ok, borderRadius: 10, padding: "1px 7px", textTransform: "uppercase", letterSpacing: "0.06em" }}>{r.reportType || "monthly"}</span>
                      </div>
                      <div style={{ fontSize: 12, color: T.textMuted }}>Published: {r.publishedDate}{r.periodFrom && r.periodTo ? ` · ${r.periodFrom} → ${r.periodTo}` : ""} · Visible to all teams</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
                      <Btn small onClick={() => { setEditReportId(r.id); setEditReportForm({ month: r.month, notes: r.notes || "" }); }}>Edit</Btn>
                      <button onClick={() => { if (window.confirm(`Delete report "${r.month}"? This cannot be undone.`)) dispatch({ type: "REMOVE_REPORT", reportId: r.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete report">✕</button>
                    </div>
                  </div>
                  {editReportId === r.id ? (
                    <div style={{ padding: "16px 20px" }}>
                      <div style={{ marginBottom: 12 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Period Label</div>
                        <Input value={editReportForm.month} onChange={e => setEditReportForm(f => ({ ...f, month: e.target.value }))} style={{ width: "100%", marginBottom: 10 }} />
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Notes</div>
                        <TextArea value={editReportForm.notes} onChange={e => setEditReportForm(f => ({ ...f, notes: e.target.value }))} placeholder="Add notes about this report period..." rows={3} />
                      </div>
                      <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                        <Btn small onClick={() => setEditReportId(null)}>Cancel</Btn>
                        <Btn primary small onClick={() => { dispatch({ type: "EDIT_REPORT", reportId: r.id, updates: { month: editReportForm.month, notes: editReportForm.notes } }); setEditReportId(null); }}>Save</Btn>
                      </div>
                    </div>
                  ) : (
                    <div style={{ padding: "16px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
                      <div>
                        <SectionLabel>Department Rankings</SectionLabel>
                        {r.data.deptRanks.map((d, i) => (
                          <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 0", fontSize: 14 }}>
                            <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                            <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                            <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{d.rate}%</span>
                            <Tag type={d.status} small />
                          </div>
                        ))}
                        {r.submissionRate != null && (
                          <div style={{ marginTop: 12 }}>
                            <span style={{ fontSize: 12, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "2px 8px", color: T.brand }}>Submission rate: {r.submissionRate}%</span>
                          </div>
                        )}
                        {r.notes && <div style={{ marginTop: 12, padding: "8px 12px", background: T.raised, borderRadius: 7, fontSize: 13, color: T.textSoft, lineHeight: 1.6 }}><strong>Notes:</strong> {r.notes}</div>}
                      </div>
                      <div>
                        <SectionLabel>Top Performers</SectionLabel>
                        {r.data.topPerformers.map((p, i) => <div key={i} style={{ padding: "5px 0", fontSize: 14, color: T.ok, display: "flex", alignItems: "center", gap: 6 }}><span>★</span> {p}</div>)}
                        {r.data.redFlags.length > 0 && (<><SectionLabel>Action Required</SectionLabel>{r.data.redFlags.map((f, i) => <div key={i} style={{ padding: "5px 0", fontSize: 14, color: T.bad, display: "flex", alignItems: "center", gap: 6 }}><span>⚠</span> {f}</div>)}</>)}
                      </div>
                    </div>
                  )}
                </Card>
              ));
            })()}
          </Pane>
        </>)}

        {page === "projects" && (<>
          <Header title="Manager Projects" sub="Projects grouped by department" />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Total Projects" value={projects.length} />
              <Metric label="Active"         value={projects.filter(p => p.status === "active").length}   status="yellow" />
              <Metric label="Completed"      value={projects.filter(p => p.status !== "active").length}   status="green"  />
            </div>
            {projects.length === 0 && <EmptyState text="No projects submitted by managers yet." />}
            {projects.length > 0 && (() => {
              return depts.map(dept => {
                const deptManagers = users.filter(u => u.role === "manager" && u.deptId === dept.id);
                const deptProjects = projects.filter(p => deptManagers.some(m => m.id === p.mgrId));
                if (deptProjects.length === 0) return null;
                return (
                  <div key={dept.id} style={{ marginBottom: 24 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, paddingBottom: 8, borderBottom: `2px solid ${T.border}` }}>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{dept.name}</div>
                      <span style={{ marginLeft: "auto", fontSize: 12, color: T.textMuted }}>{deptProjects.length} project{deptProjects.length !== 1 ? "s" : ""}</span>
                    </div>
                    {deptProjects.map(p => {
                      const mgr = deptManagers.find(m => m.id === p.mgrId);
                      const isActive = p.status === "active";
                      const ps = p.progress >= 70 ? "green" : p.progress >= 35 ? "yellow" : "red";
                      const isEditing = editProjId === p.id;
                      return (
                        <Card key={p.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                          <div style={{ padding: "12px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                            <div>
                              <div style={{ fontSize: 15, fontWeight: 700 }}>{p.name}</div>
                              <div style={{ fontSize: 12, color: T.textMuted }}>{mgr ? `${mgr.name} · ` : ""}Due: {p.due}{p.updatedDate ? ` · Last updated: ${p.updatedDate}` : ""}</div>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <Tag type={isActive ? "pending" : "approved"} label={isActive ? "ACTIVE" : "COMPLETED"} small />
                              <Btn small onClick={() => {
                                if (isEditing) { setEditProjId(null); return; }
                                setEditProjId(p.id);
                                setEditProjForm({ name: p.name, progress: p.progress, status: p.status, log: p.log || "", due: p.due || "" });
                              }}>{isEditing ? "Cancel" : "Edit"}</Btn>
                              <button onClick={() => { if (window.confirm(`Delete project "${p.name}"? This cannot be undone.`)) dispatch({ type: "REMOVE_PROJECT", projectId: p.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete project">✕</button>
                            </div>
                          </div>
                          {!isEditing && (
                            <div style={{ padding: "12px 18px" }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: p.log ? 10 : 0 }}>
                                <Bar value={p.progress} status={ps} h={6} />
                                <span style={{ fontFamily: F.mono, fontSize: 14, fontWeight: 700, color: STATUS_THEME[ps].color, whiteSpace: "nowrap" }}>{p.progress}%</span>
                              </div>
                              {p.log && <p style={{ margin: 0, fontSize: 13, color: T.textSoft, lineHeight: 1.6, padding: "8px 12px", background: T.raised, borderRadius: 6 }}>{p.log}</p>}
                            </div>
                          )}
                          {isEditing && (
                            <div style={{ padding: "14px 18px" }}>
                              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Name</div>
                                  <Input value={editProjForm.name} onChange={e => setEditProjForm(f => ({ ...f, name: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Completion %</div>
                                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                    <input type="range" min={0} max={100} value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Number(e.target.value) }))} style={{ flex: 1 }} />
                                    <Input value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 50, textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                                  </div>
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                                  <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 14, fontFamily: F.body }}>
                                    <option value="active">Active</option>
                                    <option value="completed">Completed</option>
                                  </select>
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                                  <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                                </div>
                              </div>
                              <div style={{ marginBottom: 12 }}>
                                <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Log / Notes</div>
                                <TextArea value={editProjForm.log} onChange={e => setEditProjForm(f => ({ ...f, log: e.target.value }))} placeholder="Notes, updates, observations..." rows={3} />
                              </div>
                              <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                                <Btn small onClick={() => setEditProjId(null)}>Cancel</Btn>
                                <Btn primary small disabled={!editProjForm.name.trim()} onClick={() => {
                                  dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { name: editProjForm.name.trim(), progress: editProjForm.progress, status: editProjForm.status, log: editProjForm.log, due: editProjForm.due || p.due, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                                  setEditProjId(null);
                                }}>Save</Btn>
                              </div>
                            </div>
                          )}
                        </Card>
                      );
                    })}
                  </div>
                );
              });
            })()}
          </Pane>
        </>)}

        {page === "leaderboard" && (<>
          <Header title="Company Leaderboard" sub={`All staff ranked by OKR completion · ${currentFYQuarter()}`} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Total"    value={allMembers.length} />
              <Metric label="On Track" value={allMembers.filter(m => m.status === "green").length}  status="green"  />
              <Metric label="At Risk"  value={allMembers.filter(m => m.status === "yellow").length} status="yellow" />
              <Metric label="Behind"   value={allMembers.filter(m => m.status === "red").length}    status="red"    />
            </div>
            <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                <div style={{ position: "relative", flex: "0 0 220px" }}>
                  <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", color: T.textDim, fontSize: 14, pointerEvents: "none" }}>⌕</span>
                  <input
                    value={lbSearch}
                    onChange={e => setLbSearch(e.target.value)}
                    placeholder="Search name..."
                    style={{ width: "100%", boxSizing: "border-box", paddingLeft: 28, paddingRight: 10, paddingTop: 7, paddingBottom: 7, fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, fontFamily: F.body, outline: "none" }}
                  />
                </div>
                <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                  <Btn small primary={lbDeptFilter === "all"} onClick={() => setLbDeptFilter("all")}>All Depts</Btn>
                  {depts.map(d => (
                    <Btn key={d.id} small primary={lbDeptFilter === d.id} onClick={() => setLbDeptFilter(d.id)}>{d.name}</Btn>
                  ))}
                </div>
              </div>
              <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                {[{ key: "all", label: "All OKRs" }, { key: "daily", label: "Daily" }, { key: "weekly", label: "Weekly" }, { key: "monthly", label: "Monthly" }, { key: "quarterly", label: "Quarterly" }, { key: "biannual", label: "Bi-Annual" }, { key: "annual", label: "Annual" }].map(({ key, label }) => (
                  <Btn key={key} small primary={lbPeriod === key} onClick={() => { setLbPeriod(key); setLbExpandedMember(null); }}>{label}</Btn>
                ))}
              </div>
            </div>
            {(() => {
              const q = lbSearch.trim().toLowerCase();
              const periodMembers = lbPeriod === "all" ? allMembers : allMembers.map(m => {
                const periodKrs = (memberData[m.id]?.krs || []).filter(kr => (kr.period || "monthly") === lbPeriod);
                const periodSubs = okrSubmissions.filter(s => s.period === lbPeriod);
                const rate = calcMemberRate(m.id, periodKrs, periodSubs);
                return { ...m, rate, status: getStatus(rate) };
              }).sort((a, b) => b.rate - a.rate);
              const filtered = periodMembers.filter(m =>
                (lbDeptFilter === "all" || m.deptId === lbDeptFilter) &&
                (!q || m.name.toLowerCase().includes(q) || m.title?.toLowerCase().includes(q))
              );
              if (filtered.length === 0) return <EmptyState text="No members match your search." />;
              const COL = "50px 32px 1fr 120px 110px 55px 150px 70px 56px";
              return (
                <Card style={{ overflow: "hidden" }}>
                  <div style={{ display: "grid", gridTemplateColumns: COL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                    <span>Rank</span><span></span><span>Name</span><span>Department</span><span>Team</span><span style={{ textAlign: "right" }}>Rate</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span><span></span>
                  </div>
                  {filtered.map(m => {
                    const globalRank = periodMembers.indexOf(m) + 1;
                    const isTop = globalRank === 1;
                    const isExpanded = lbExpandedMember === m.id;
                    const kd = memberData[m.id];
                    const krs = kd?.krs || [];
                    return (
                      <div key={m.id}>
                        <div style={{ display: "grid", gridTemplateColumns: COL, padding: "10px 16px", gap: 8, alignItems: "center", background: isExpanded ? T.brandDim : isTop ? T.okDim : m.status === "red" ? T.badDim : globalRank % 2 ? T.raised : "transparent", borderBottom: isExpanded ? "none" : `1px solid ${T.border}`, borderLeft: isTop ? `3px solid ${T.ok}` : m.status === "red" ? `3px solid ${T.bad}` : isExpanded ? `3px solid ${T.brand}` : "3px solid transparent", fontSize: 14 }}>
                          <span style={{ fontFamily: F.mono, fontWeight: 900, color: isTop ? T.ok : m.status === "red" ? T.bad : T.textMuted }}>#{globalRank}</span>
                          <Avatar letters={m.av} size={24} />
                          <div><span style={{ fontWeight: 600 }}>{m.name}</span><span style={{ color: T.textDim, marginLeft: 6, fontSize: 12 }}>{m.title}</span></div>
                          <span style={{ fontSize: 12, color: T.textMuted }}>{m.deptName}</span>
                          <span style={{ fontSize: 12, color: T.textMuted }}>{m.teamName}</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[m.status].color }}>{m.rate.toFixed(1)}%</span>
                          <Bar value={m.rate} status={m.status} h={5} />
                          <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={m.status} small /></div>
                          <div style={{ display: "flex", justifyContent: "flex-end" }}>
                            <button onClick={() => setLbExpandedMember(isExpanded ? null : m.id)}
                              style={{ fontSize: 11, fontWeight: 700, fontFamily: F.body, padding: "3px 8px", borderRadius: 6, cursor: "pointer", border: `1px solid ${isExpanded ? T.brand : T.border}`, background: isExpanded ? T.brandDim : T.surface, color: isExpanded ? T.brand : T.textSoft }}>
                              {isExpanded ? "Close" : "Edit"}
                            </button>
                          </div>
                        </div>
                        {isExpanded && (
                          <div style={{ borderBottom: `1px solid ${T.border}`, borderLeft: `3px solid ${T.brand}`, background: T.bg, padding: "14px 20px" }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: T.brand, marginBottom: 12 }}>Editing KPIs — {m.name}</div>
                            {krs.length === 0
                              ? <div style={{ fontSize: 13, color: T.textMuted }}>No OKRs synced yet — assign this member to a team in User Management, then click ⟳ Sync to Team Members in the Departments tab.</div>
                              : [{ key: "daily", label: "Daily OKRs" }, { key: "weekly", label: "Weekly OKRs" }, { key: "monthly", label: "Monthly OKRs" }, { key: "quarterly", label: "Quarterly OKRs" }, { key: "biannual", label: "Bi-Annual OKRs" }, { key: "annual", label: "Annual OKRs" }].map(({ key, label }) => {
                                const group = krs.filter(kr => (kr.period || "monthly") === key);
                                if (group.length === 0) return null;
                                return (
                                  <div key={key} style={{ marginBottom: 14 }}>
                                    <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase", marginBottom: 6 }}>{label} ({group.length})</div>
                                    <div style={{ display: "grid", gridTemplateColumns: "50px 1fr 90px 110px 55px 140px 32px", gap: 8, padding: "6px 10px", fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.06em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}` }}>
                                      <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span></span>
                                    </div>
                                    {group.map((kr, ki) => {
                                      const pct = krCompletion(kr); const st = getStatus(pct);
                                      return (
                                        <div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 90px 110px 55px 140px 32px", gap: 8, padding: "8px 10px", alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 13 }}>
                                          <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim }}>{kr.id}</span>
                                          <div>
                                            <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }} title={kr.label}>{kr.label}</span>
                                            {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                                          </div>
                                          {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 11, color: "#7c3aed" }}>N/A</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)}{kr.unit ? ` ${kr.unit}` : ""}</span>}
                                          {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.actual)}</span> : <Input value={kr.actual} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: m.id, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "4px 8px", fontSize: 13, fontFamily: F.mono }} />}
                                          {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 11, color: "#7c3aed" }}>—</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[st].color }}>{pct.toFixed(0)}%</span>}
                                          {kr.type === "tracker" ? <span /> : <Bar value={pct} status={st} h={5} />}
                                          <button onClick={() => setConfirmDeleteKr({ memberId: m.id, memberName: m.name, krId: kr.id, krLabel: kr.label })}
                                            title="Delete this OKR"
                                            style={{ background: "none", border: `1px solid ${T.badBorder || T.bad}`, borderRadius: 5, padding: "3px 6px", cursor: "pointer", color: T.bad, fontSize: 12, lineHeight: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                                        </div>
                                      );
                                    })}
                                  </div>
                                );
                              })
                            }
                          </div>
                        )}
                      </div>
                    );
                  })}
                </Card>
              );
            })()}
          </Pane>
        </>)}

        {confirmDeleteKr && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16, padding: "28px 32px", width: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.22)" }}>
              <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 8, color: T.text }}>Delete OKR?</div>
              <p style={{ fontSize: 14, color: T.textSoft, lineHeight: 1.6, margin: "0 0 6px" }}>
                You are about to delete this OKR for <strong style={{ color: T.text }}>{confirmDeleteKr.memberName}</strong>:
              </p>
              <div style={{ background: T.badDim, border: `1px solid ${T.badBorder || T.bad}`, borderRadius: 8, padding: "10px 14px", fontSize: 14, fontWeight: 600, color: T.bad, margin: "0 0 20px" }}>
                {confirmDeleteKr.krLabel}
              </div>
              <p style={{ fontSize: 13, color: T.textMuted, margin: "0 0 20px", lineHeight: 1.5 }}>This only removes the OKR from their personal KPI list. It does not affect department or team KRs. This action cannot be undone.</p>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <Btn onClick={() => setConfirmDeleteKr(null)}>Cancel</Btn>
                <Btn danger onClick={() => { dispatch({ type: "REMOVE_MEMBER_KR", memberId: confirmDeleteKr.memberId, krId: confirmDeleteKr.krId }); setConfirmDeleteKr(null); }}>Delete OKR</Btn>
              </div>
            </div>
          </div>
        )}

        {page === "email-templates" && (() => {
          const TMPL_PERIODS = [
            { id: "default",   label: "Default (all periods)" },
            { id: "daily",     label: "Daily"     },
            { id: "weekly",    label: "Weekly"    },
            { id: "monthly",   label: "Monthly"   },
            { id: "quarterly", label: "Quarterly" },
            { id: "biannual",  label: "Bi-Annual" },
            { id: "annual",    label: "Annual"    },
          ];
          const PERIOD_LABELS = { daily: "Daily", weekly: "Weekly", monthly: "Monthly", quarterly: "Quarterly", biannual: "Bi-Annual", annual: "Annual", default: "Default" };
          const DEFAULT_TMPL = {
            fromName: "NIET Group OKR",
            subject:  "Action Required: {periodLabel} KPI Check-in — {periodKey}",
            body:     "Here are your {periodLower} KPI targets for <strong>{periodKey}</strong>.\nPlease log in to the portal and mark whether you have met each target.",
            ctaText:  "Submit My Check-in →",
            footer:   "You are receiving this because you have KPI targets in the NIET Group OKR system.\nPlease do not reply to this email.",
          };
          const saved = settings?.emailTemplates || {};
          const effective = { ...DEFAULT_TMPL, ...(saved.default || {}), ...(tmplPeriod !== "default" ? (saved[tmplPeriod] || {}) : {}) };
          const localDraft = tmplPeriod === "default" ? { ...DEFAULT_TMPL, ...(saved.default || {}) } : { ...effective };
          const setField = (field, value) => {
            const updated = { ...saved, [tmplPeriod]: { ...(saved[tmplPeriod] || {}), [field]: value } };
            dispatch({ type: "SET_SETTINGS", updates: { emailTemplates: updated } });
          };
          const previewPeriodKey = { daily: "Mon 14 Jul 2026", weekly: "FY26 W02", monthly: "July 2026", quarterly: "FY26 Q1", biannual: "FY26 H1", annual: "FY 2026", default: "July 2026" }[tmplPeriod] || "July 2026";
          const previewPeriodLabel = PERIOD_LABELS[tmplPeriod === "default" ? "monthly" : tmplPeriod] || "Monthly";
          const previewSubject = localDraft.subject.replace(/\{periodLabel\}/g, previewPeriodLabel).replace(/\{periodKey\}/g, previewPeriodKey);
          const previewBody = localDraft.body.replace(/\{periodLower\}/g, previewPeriodLabel.toLowerCase()).replace(/\{periodKey\}/g, `<strong>${previewPeriodKey}</strong>`).replace(/\{periodLabel\}/g, previewPeriodLabel);
          const Field = ({ label, value, onChange, hint, multiline }) => (
            <div style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>{label}</div>
              {multiline
                ? <textarea value={value} onChange={e => onChange(e.target.value)} rows={3} style={{ width: "100%", boxSizing: "border-box", padding: "8px 12px", fontSize: 14, fontFamily: F.body, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, resize: "vertical", outline: "none", lineHeight: 1.5 }} />
                : <input value={value} onChange={e => onChange(e.target.value)} style={{ width: "100%", boxSizing: "border-box", padding: "8px 12px", fontSize: 14, fontFamily: F.body, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none" }} />}
              {hint && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 4 }}>{hint}</div>}
            </div>
          );
          return (<>
            <Header title="Email Templates" sub="Customise the check-in notification emails sent to staff" />
            <Pane>
              <div style={{ display: "flex", gap: 4, marginBottom: 20, paddingBottom: 14, borderBottom: `1px solid ${T.border}` }}>
                {TMPL_PERIODS.map(p => <Btn key={p.id} primary={tmplPeriod === p.id} small onClick={() => setTmplPeriod(p.id)}>{p.label}</Btn>)}
              </div>
              {tmplPeriod !== "default" && (
                <div style={{ display: "flex", alignItems: "center", gap: 10, fontSize: 12, color: T.textMuted, background: T.raised, borderRadius: 8, padding: "8px 14px", marginBottom: 18, border: `1px solid ${T.border}` }}>
                  <span style={{ flex: 1 }}>
                    {saved[tmplPeriod]
                      ? `✓ Custom override active for ${PERIOD_LABELS[tmplPeriod]} check-ins.`
                      : `No override set — inheriting the Default template. Edit any field below to create a ${PERIOD_LABELS[tmplPeriod]}-specific override.`}
                  </span>
                  {saved[tmplPeriod] && (
                    <button onClick={() => {
                      const upd = { ...saved };
                      delete upd[tmplPeriod];
                      dispatch({ type: "SET_SETTINGS", updates: { emailTemplates: upd } });
                    }} style={{ background: "none", border: `1px solid ${T.badBorder || T.bad}`, color: T.bad, fontSize: 12, cursor: "pointer", fontWeight: 600, padding: "3px 10px", borderRadius: 6 }}>Remove override</button>
                  )}
                </div>
              )}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 32, alignItems: "start" }}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14, color: T.text }}>Template Fields</div>
                  <Field label="Sender Name" value={localDraft.fromName} onChange={v => setField("fromName", v)} hint='Displayed in the From field, e.g. "NIET Group OKR"' />
                  <Field label="Subject Line" value={localDraft.subject} onChange={v => setField("subject", v)} hint="Placeholders: {periodLabel} · {periodKey}" />
                  <Field label="Email Body" value={localDraft.body} onChange={v => setField("body", v)} hint="HTML tags OK. Placeholders: {periodLower} · {periodLabel} · {periodKey}" multiline />
                  <Field label="CTA Button Text" value={localDraft.ctaText} onChange={v => setField("ctaText", v)} />
                  <Field label="Footer Text" value={localDraft.footer} onChange={v => setField("footer", v)} hint="Use \n for line breaks." multiline />
                  {tmplPeriod === "default" && <div style={{ fontSize: 12, color: T.textMuted, padding: "10px 0 4px" }}>Changes to the Default template are saved automatically and apply to all periods that don't have an override.</div>}
                </div>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14, color: T.text }}>Live Preview</div>
                  <div style={{ border: `1px solid ${T.border}`, borderRadius: 10, overflow: "hidden" }}>
                    <div style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, padding: "8px 14px", fontSize: 12, color: T.textMuted, display: "flex", gap: 12, flexWrap: "wrap" }}>
                      <span><span style={{ fontWeight: 600, color: T.text }}>From: </span>{localDraft.fromName} &lt;noreply@niet.edu.au&gt;</span>
                      <span><span style={{ fontWeight: 600, color: T.text }}>Subject: </span>{previewSubject}</span>
                    </div>
                    <div style={{ padding: 12, background: "#f0f0f5", maxHeight: 520, overflowY: "auto" }}>
                      <div style={{ maxWidth: 440, margin: "0 auto", background: "#fff", borderRadius: 10, overflow: "hidden", boxShadow: "0 2px 10px rgba(0,0,0,0.1)", fontFamily: "-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif", fontSize: 14 }}>
                        <div style={{ background: "linear-gradient(135deg,#0071e3,#6b47dc)", padding: "20px 24px" }}>
                          <div style={{ color: "#fff", fontSize: 9, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", opacity: 0.8, marginBottom: 4 }}>NIET Group OKR System</div>
                          <div style={{ color: "#fff", fontSize: 17, fontWeight: 700 }}>{previewPeriodLabel} KPI Check-in</div>
                        </div>
                        <div style={{ padding: "20px 24px" }}>
                          <p style={{ margin: "0 0 14px", fontSize: 14, color: "#1d1d1f" }}>Hi <strong>John Smith</strong>,</p>
                          <p style={{ margin: "0 0 14px", fontSize: 13, color: "#6e6e73", lineHeight: 1.6 }} dangerouslySetInnerHTML={{ __html: previewBody.replace(/\n/g, "<br/>") }} />
                          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 18 }}>
                            <thead>
                              <tr style={{ background: "#f5f5f7" }}>
                                <th style={{ padding: "6px 10px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#6e6e73", textTransform: "uppercase", letterSpacing: "0.06em" }}>Key Result</th>
                                <th style={{ padding: "6px 10px", textAlign: "right", fontSize: 10, fontWeight: 700, color: "#6e6e73", textTransform: "uppercase", letterSpacing: "0.06em" }}>Target</th>
                              </tr>
                            </thead>
                            <tbody>
                              {[{ label: "Complete 3 coaching sessions", target: "3", unit: "sessions" }, { label: "Customer satisfaction score", target: "≥ 90", unit: "%" }].map((kr, i) => (
                                <tr key={i}><td style={{ padding: "7px 10px", borderBottom: "1px solid #e5e7eb", fontSize: 13 }}>{kr.label}</td><td style={{ padding: "7px 10px", borderBottom: "1px solid #e5e7eb", textAlign: "right", fontFamily: "monospace", fontSize: 13 }}>{kr.target} {kr.unit}</td></tr>
                              ))}
                            </tbody>
                          </table>
                          <div style={{ display: "inline-block", padding: "10px 20px", background: "#0071e3", color: "#fff", borderRadius: 8, fontSize: 13, fontWeight: 600 }}>{localDraft.ctaText}</div>
                          <p style={{ margin: "14px 0 0", fontSize: 11, color: "#a1a1aa", lineHeight: 1.5 }} dangerouslySetInnerHTML={{ __html: localDraft.footer.replace(/\n/g, "<br/>") }} />
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
              <div style={{ marginTop: 28, paddingTop: 20, borderTop: `1px solid ${T.border}` }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 12, color: T.text }}>Send Test Email</div>
                <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                  <input
                    value={testEmailTo}
                    onChange={e => { setTestEmailTo(e.target.value); setTestEmailState({ status: "idle", msg: "" }); }}
                    placeholder="recipient@example.com"
                    type="email"
                    style={{ flex: "1 1 220px", minWidth: 200, padding: "8px 12px", fontSize: 14, fontFamily: F.body, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none" }}
                  />
                  <Btn primary disabled={testEmailState.status === "sending" || !testEmailTo} onClick={async () => {
                    if (!testEmailTo) return;
                    setTestEmailState({ status: "sending", msg: "" });
                    const period = tmplPeriod === "default" ? "monthly" : tmplPeriod;
                    try {
                      const res = await fetch("/api/send-email", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({
                          to: testEmailTo,
                          name: "Test User",
                          period,
                          periodKey: previewPeriodKey,
                          krs: [
                            { label: "Complete 3 coaching sessions", target: "3", unit: "sessions" },
                            { label: "Customer satisfaction score", target: "90", unit: "%" },
                          ],
                          template: localDraft,
                        }),
                      });
                      const data = await res.json();
                      if (data.ok) {
                        setTestEmailState({ status: "sent", msg: data.skipped ? "SMTP not configured — email skipped (check SMTP_USER/SMTP_PASS env vars)" : `Sent to ${testEmailTo}` });
                      } else {
                        setTestEmailState({ status: "error", msg: data.error || "Send failed" });
                      }
                    } catch (err) {
                      setTestEmailState({ status: "error", msg: err.message });
                    }
                  }}>
                    {testEmailState.status === "sending" ? "Sending…" : "Send Test Email"}
                  </Btn>
                  {testEmailState.status === "sent" && <span style={{ fontSize: 13, color: T.ok }}>✓ {testEmailState.msg}</span>}
                  {testEmailState.status === "error" && <span style={{ fontSize: 13, color: T.bad }}>✗ {testEmailState.msg}</span>}
                </div>
              </div>
            </Pane>
          </>);
        })()}

        {syncPrompt && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div style={{ background: T.surface, borderRadius: 16, padding: "28px 32px", width: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}>
              <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 10 }}>Sync KPIs to Team Members?</div>
              <p style={{ fontSize: 14, color: T.textSoft, marginBottom: 24, lineHeight: 1.6, margin: "0 0 24px" }}>
                You updated KPIs for <strong>{syncPrompt.teamName}</strong>. Sync these changes to all team members' personal KPI lists?
              </p>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <Btn onClick={() => setSyncPrompt(null)}>Skip</Btn>
                <Btn primary onClick={() => { doSync(syncPrompt.deptId, syncPrompt.teamId); setSyncPrompt(null); }}>Yes, Sync</Btn>
              </div>
            </div>
          </div>
        )}
        {syncNote && (
          <div style={{ position: "fixed", bottom: 28, right: 28, zIndex: 1100, background: T.ok, color: "#fff", borderRadius: 12, padding: "14px 22px", boxShadow: "0 6px 28px rgba(0,0,0,0.22)", display: "flex", alignItems: "center", gap: 12, fontSize: 14, fontWeight: 600, minWidth: 280 }}>
            <span style={{ fontSize: 20 }}>✓</span>
            <div>
              <div>KPIs synced to {syncNote.count} member{syncNote.count !== 1 ? "s" : ""}</div>
              <div style={{ fontSize: 12, fontWeight: 400, opacity: 0.85, marginTop: 2 }}>{syncNote.teamName}</div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   MANAGER PORTAL
   ───────────────────────────────────────────────────────────── */
function ManagerPortal({ user, onLogout, state, dispatch }) {
  const [page, setPageRaw] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'manager' ? (p[2] || 'dashboard') : 'dashboard';
  });
  const setPage = useCallback(p => { window.history.pushState(null, '', `/manager/${p}`); setPageRaw(p); }, []);
  useEffect(() => {
    if (window.location.pathname.split('/')[1] !== 'manager') {
      window.history.replaceState(null, '', `/manager/dashboard`);
    }
    const onPop = () => {
      const p = window.location.pathname.split('/');
      setPageRaw(p[1] === 'manager' ? (p[2] || 'dashboard') : 'dashboard');
    };
    window.addEventListener('popstate', onPop);
    return () => window.removeEventListener('popstate', onPop);
  }, []);
  const [newProj, setNewProj] = useState({ name: "", due: "" });
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ progress: 0, status: "active", log: "", due: "" });
  const [syncPrompt, setSyncPrompt] = useState(null);
  const syncTimerRef = useRef(null);
  const [syncNote, setSyncNote] = useState(null);
  const syncNoteTimer = useRef(null);
  const [mgrDirtySync, setMgrDirtySync] = useState(null);
  const [expandedMonthlyKr, setExpandedMonthlyKr] = useState(null);
  const [mgrSelTeam, setMgrSelTeam] = useState(null);
  const [okrPeriod, setOkrPeriod] = useState("all");
  const [noReason, setNoReason] = useState(null);
  const [rejectOkr, setRejectOkr] = useState(null);
  const [trackerInput, setTrackerInput] = useState({});

  const { depts, memberData, okrSubmissions: allOkrSubs = [], projects, monthlyReports, users } = state;
  const dept = depts.find(d => d.id === user.deptId);
  const overseeFilter = u => !user.teamIds?.length || (u.teamId && user.teamIds.includes(u.teamId)) || (u.secondTeamId && user.teamIds.includes(u.secondTeamId));
  const extraDeptIds = user.mgrDeptIds || [];
  const myMembers = users.filter(u => (u.role === "member" || u.role === "manager") && (
    (u.deptId === user.deptId && overseeFilter(u)) || extraDeptIds.includes(u.deptId)
  ));
  const myTeamMemberIds = users.filter(u => u.role === "member" && (
    (u.deptId === user.deptId && overseeFilter(u)) || extraDeptIds.includes(u.deptId)
  )).map(u => u.id);
  const myOkrSubs = allOkrSubs.filter(s => s.memberId === user.id);
  const myPendingCheckins = myOkrSubs.filter(s => s.answer === null);
  const myOkrSubsForApproval = allOkrSubs.filter(s => myTeamMemberIds.includes(s.memberId));
  const pendingOkrSubs = myOkrSubsForApproval.filter(s => s.answer !== null && s.approval === "pending");
  const myProjects = projects.filter(p => p.mgrId === user.id);

  function mgrTriggerSync(deptId, teamId) {
    if (syncTimerRef.current) clearTimeout(syncTimerRef.current);
    const d = depts.find(x => x.id === deptId);
    const t = d?.teams.find(x => x.id === teamId);
    if (!t) return;
    setMgrDirtySync({ deptId, teamId });
    syncTimerRef.current = setTimeout(() => setSyncPrompt({ deptId, teamId, teamName: t.name }), 1500);
  }

  function mgrDoSync(deptId, teamId) {
    const d = depts.find(x => x.id === deptId);
    const t = d?.teams.find(x => x.id === teamId);
    if (!t) return;
    const count = new Set([...(t.members || []), ...users.filter(u => u.teamId === teamId || u.secondTeamId === teamId).map(u => u.id)]).size;
    dispatch({ type: "SYNC_TEAM_KRS_TO_MEMBERS", deptId, teamId });
    setMgrDirtySync(null);
    if (syncNoteTimer.current) clearTimeout(syncNoteTimer.current);
    setSyncNote({ teamName: t.name, count });
    syncNoteTimer.current = setTimeout(() => setSyncNote(null), 3500);
  }

  function mgrDoDeptSync(deptId) {
    const d = depts.find(x => x.id === deptId);
    if (!d) return;
    const count = users.filter(u => u.deptId === deptId && (u.role === "member" || u.role === "manager")).length;
    dispatch({ type: "SYNC_DEPT_KRS_TO_MEMBERS", deptId });
    if (syncNoteTimer.current) clearTimeout(syncNoteTimer.current);
    setSyncNote({ teamName: `${d.name} (all members)`, count });
    syncNoteTimer.current = setTimeout(() => setSyncNote(null), 3500);
  }

  const navItems = [
    { id: "dashboard",    icon: "⧉", label: "Team Dashboard"       },
    { id: "okr-overview", icon: "◎", label: "OKR Overview"         },
    { id: "checkin",      icon: "✓", label: "OKR Check-in"         },
    { id: "approvals",    icon: "✓", label: "Approve Submissions"   },
    { id: "projects",     icon: "⚡", label: "Projects"             },
    { id: "members",      icon: "✎", label: "Edit Member KPIs"     },
    { id: "reports",      icon: "⊞", label: "OKR Reports"          },
  ];

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} pendingCounts={{ approvals: pendingOkrSubs.length, checkin: myPendingCheckins.length }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "okr-overview" && (() => {
          if (!dept) return (<><Header title="OKR Overview" sub="Your department's key results by period" /><Pane><EmptyState text="No department assigned to your account." /></Pane></>);
          const PERIODS = [{ id: "all", label: "All" }, { id: "daily", label: "Daily" }, { id: "weekly", label: "Weekly" }, { id: "monthly", label: "Monthly" }, { id: "quarterly", label: "Quarterly" }, { id: "biannual", label: "Bi-Annual" }, { id: "annual", label: "Annual" }];
          const filterP = krs => {
            const byPeriod = okrPeriod === "all" ? krs : krs.filter(kr => (kr.period || "monthly") === okrPeriod);
            return byPeriod.filter(kr => kr.type !== "tracker" || kr.showInOverview !== false);
          };
          const pLabel = okrPeriod === "all" ? "All Periods" : periodDateRange(okrPeriod, okrPeriod === "weekly" ? prevPeriodKey(okrPeriod) : currentPeriodKey(okrPeriod));
          const PChip = () => <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 7px", fontWeight: 700, marginLeft: 8, textTransform: "none", letterSpacing: "normal", verticalAlign: "middle" }}>{pLabel}</span>;
          const KCOL = "50px 1fr 100px 110px 150px 55px 130px 65px";
          const renderRows = (krs, deptId, teamId) => krs.map((kr, i) => {
            const pct = krCompletion(kr); const st = getStatus(pct);
            const isMonthly = !!kr.monthlyTargets;
            const curKey = currentFYMonthKey();
            const curTarget = isMonthly ? (kr.monthlyTargets[curKey] || 0) : null;
            const curActual = isMonthly ? ((kr.monthlyActuals || {})[curKey] || 0) : null;
            const fyMs = isMonthly ? getFYMonths() : [];
            const annSumTarget = fyMs.reduce((s, {key}) => s + (kr.monthlyTargets[key] || 0), 0);
            const annActual = fyMs.reduce((s, {key}) => s + ((kr.monthlyActuals || {})[key] || 0), 0);
            const annDream = isMonthly ? (kr.annualTarget || 0) : 0;
            const annVsSum = annSumTarget > 0 ? Math.min((annActual / annSumTarget) * 100, 100) : 0;
            const annVsDream = annDream > 0 ? Math.min((annActual / annDream) * 100, 100) : 0;
            const annSt = (annDream > 0 ? annVsDream : annVsSum) >= 80 ? "green" : (annDream > 0 ? annVsDream : annVsSum) >= 50 ? "yellow" : "red";
            return (
              <Fragment key={kr.id}>
              <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                <div>
                  <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                  {kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                  {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                  {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Monthly Breakdown</span>}
                  {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                </div>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? `${kr.operator||">="} ${fmt(curTarget)} this mo.` : `${kr.operator || ">="} ${fmt(kr.target)}${kr.unit ? ` ${kr.unit}` : ""}`}</span>}
                {kr.type === "tracker"
                  ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textDim }}>—</span>
                  : isMonthly
                  ? <Input value={curActual} onChange={e => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: curKey, field: "actual", value: Number(e.target.value) || 0 }); }} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                  : <Input value={kr.actual} onChange={e => { dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 }); if (teamId) mgrTriggerSync(deptId, teamId); }} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />}
                <span style={{ fontSize: 12, color: T.textMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.dataSource || "—"}</span>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 14, fontWeight: 700, color: "#7c3aed" }}>{fmt(isMonthly ? curActual : kr.actual)}{kr.unit ? <span style={{ fontSize: 11, fontWeight: 400 }}> {kr.unit}</span> : ""}</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[st].color }}>{pct.toFixed(0)}%</span>}
                {kr.type === "tracker" ? <span /> : <Bar value={pct} status={st} h={5} />}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 4 }}>
                  {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="View all months" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "2px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 11, fontWeight: 700 }}>📅</button>}
                  {kr.type === "tracker" ? null : <Tag type={st} small />}
                </div>
              </div>
              {isMonthly && expandedMonthlyKr === kr.id && (
                <div style={{ padding: "14px 16px 18px", background: T.brandDim, borderBottom: `1px solid ${T.border}` }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: T.brand, marginBottom: 12 }}>KPI Breakdown — {kr.label}{kr.unit ? ` (${kr.unit})` : ""}</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12, minWidth: 960 }}>
                      <thead>
                        <tr style={{ background: T.surface }}>
                          <th style={{ textAlign: "left", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, width: 110 }}></th>
                          {fyMs.map(({ key, label }) => {
                            const isCur = key === curKey;
                            return (
                              <th key={key} style={{ textAlign: "center", padding: "6px 3px", borderBottom: `2px solid ${isCur ? T.brand : T.border}`, fontSize: 11, fontWeight: isCur ? 700 : 400, color: isCur ? T.brand : T.textDim, minWidth: 62, background: isCur ? T.brandDim : T.surface }}>
                                {label.split(" ")[0]}{isCur ? " ●" : ""}
                              </th>
                            );
                          })}
                          <th style={{ textAlign: "right", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, minWidth: 80, background: T.surface }}>FY Total</th>
                          <th style={{ textAlign: "center", padding: "6px 10px", borderBottom: `2px solid ${T.okBorder}`, fontSize: 11, fontWeight: 700, color: T.ok, minWidth: 100, background: T.okDim }}>Dream Target</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Perf. Target</td>
                          {fyMs.map(({ key }) => {
                            const isCur = key === curKey;
                            const t = kr.monthlyTargets[key] || 0;
                            return (
                              <td key={key} style={{ padding: "4px 6px", textAlign: "right", fontFamily: F.mono, fontSize: 12, background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}`, color: T.textMuted }}>{fmt(t)}</td>
                            );
                          })}
                          <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annSumTarget)}</td>
                          <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, borderLeft: `1px solid ${T.okBorder}`, textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: T.ok }}>{annDream > 0 ? fmt(annDream) : <span style={{ color: T.textDim, fontWeight: 400 }}>—</span>}{kr.unit && annDream > 0 && <div style={{ fontSize: 10, color: T.ok, fontWeight: 400 }}>{kr.unit}</div>}</td>
                        </tr>
                        <tr>
                          <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Actual</td>
                          {fyMs.map(({ key }) => {
                            const isCur = key === curKey;
                            const a = (kr.monthlyActuals || {})[key] || 0;
                            return (
                              <td key={key} style={{ padding: "3px 3px", background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}` }}>
                                <Input value={a} onChange={e => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "actual", value: Number(e.target.value) || 0 }); if (teamId) mgrTriggerSync(deptId, teamId); }} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} />
                              </td>
                            );
                          })}
                          <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[annSt].color, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annActual)}</td>
                          <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, textAlign: "center", color: T.textDim, fontSize: 12, borderLeft: `1px solid ${T.okBorder}` }}>—</td>
                        </tr>
                        <tr>
                          <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 11, color: T.textDim, background: T.surface }}>Achievement</td>
                          {fyMs.map(({ key }) => {
                            const t = kr.monthlyTargets[key] || 0;
                            const a = (kr.monthlyActuals || {})[key] || 0;
                            const pct = t > 0 ? Math.min((a / t) * 100, 100) : null;
                            const isCur = key === curKey;
                            return (
                              <td key={key} style={{ padding: "4px 4px", textAlign: "center", background: isCur ? T.brandDim : "transparent" }}>
                                {pct !== null
                                  ? <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[getStatus(pct)].color }}>{pct.toFixed(0)}%</span>
                                  : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}
                              </td>
                            );
                          })}
                          <td style={{ padding: "4px 10px", textAlign: "right", background: T.surface }}>
                            {annSumTarget > 0 ? <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[annSt].color }}>{annVsSum.toFixed(0)}%</span><div style={{ fontSize: 10, color: T.textDim }}>vs. sum</div></> : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}
                          </td>
                          <td style={{ padding: "4px 8px", background: T.okDim, textAlign: "center", borderLeft: `1px solid ${T.okBorder}` }}>
                            {annDream > 0
                              ? <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[getStatus(annVsDream)].color }}>{annVsDream.toFixed(0)}%</span><div style={{ fontSize: 10, color: T.ok }}>vs. dream</div></>
                              : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
              </Fragment>
            );
          });
          const renderSection = (krs, deptId, teamId) => {
            if (krs.length === 0) return <div style={{ fontSize: 13, color: T.textMuted, padding: "10px 0" }}>No {okrPeriod} KRs for this section yet.</div>;
            return (
              <Card style={{ overflow: "hidden" }}>
                <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                  <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Data Source</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                </div>
                {renderRows(krs, deptId, teamId)}
              </Card>
            );
          };
          const dKrs = filterP(dept.krs);
          const deptRate = calcRate(dKrs); const deptStatus = getStatus(deptRate);
          const teamStats = dept.teams.map(t => { const tKrs = filterP(t.krs); return { ...t, krs: tKrs, rate: calcRate(tKrs), status: getStatus(calcRate(tKrs)) }; });
          const totalKrs = dKrs.length + teamStats.reduce((s, t) => s + t.krs.length, 0);
          return (<>
            <Header title="OKR Overview" sub={`${dept.name} · ${okrPeriod.charAt(0).toUpperCase() + okrPeriod.slice(1)} key results`} right={<Tag type={deptStatus} />} />
            <Pane>
              <div style={{ display: "flex", gap: 4, marginBottom: 20, paddingBottom: 12, borderBottom: `1px solid ${T.border}` }}>
                {PERIODS.map(p => <Btn key={p.id} primary={okrPeriod === p.id} small onClick={() => setOkrPeriod(p.id)}>{p.label}</Btn>)}
              </div>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                <Metric label="Dept Completion" value={`${deptRate.toFixed(1)}%`} status={deptStatus} sub={`Target: ${TP}%`} />
                <Metric label="KRs this period" value={totalKrs} />
                <Metric label="Teams" value={dept.teams.length} />
              </div>
              <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontSize: 14, fontWeight: 700, flex: 1 }}>{dept.name} — Department KRs<PChip /></div>
                {dKrs.length > 0 && <Btn primary small onClick={() => mgrDoDeptSync(dept.id)}>⟳ Sync to All Members</Btn>}
              </div>
              {renderSection(dKrs, dept.id, null)}
              {dept.teams.length > 0 && (<>
                <SectionLabel>Team Key Results</SectionLabel>
                <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
                  {dept.teams.map(t => <Btn key={t.id} primary={mgrSelTeam === t.id} small onClick={() => setMgrSelTeam(p => p === t.id ? null : t.id)}>{t.name}</Btn>)}
                  {mgrDirtySync?.deptId === dept.id && mgrSelTeam && mgrDirtySync?.teamId === mgrSelTeam && (
                    <div style={{ fontSize: 12, color: T.warn, background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 6, padding: "2px 9px", fontWeight: 600 }}>⚠ Unsynced changes</div>
                  )}
                </div>
                {mgrSelTeam && (() => {
                  const team = dept.teams.find(t => t.id === mgrSelTeam);
                  if (!team) return null;
                  return (<>
                    <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 10 }}>
                      {team.obj && <div style={{ fontSize: 13, color: T.textMuted, flex: 1 }}>Objective: {team.obj}{team.lead ? ` · Lead: ${team.lead}` : ""}</div>}
                      <Btn primary small onClick={() => mgrDoSync(dept.id, team.id)}>⟳ Sync to Team Members</Btn>
                    </div>
                    {renderSection(filterP(team.krs), dept.id, team.id)}
                  </>);
                })()}
              </>)}
              {teamStats.filter(t => t.krs.length > 0).length > 0 && (<>
                <SectionLabel>Team Overview<PChip /></SectionLabel>
                {teamStats.filter(t => t.krs.length > 0).map(t => (
                  <Card key={t.id} style={{ padding: "14px 16px", marginBottom: 10 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                      <span style={{ fontSize: 15, fontWeight: 700, flex: 1 }}>{t.name}<PChip /></span>
                      {t.lead && <span style={{ fontSize: 12, color: T.textMuted }}>Lead: {t.lead}</span>}
                      <span style={{ fontSize: 14, fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[t.status].color }}>{t.rate.toFixed(1)}%</span>
                      <div style={{ width: 100, flexShrink: 0 }}><Bar value={t.rate} status={t.status} h={5} /></div>
                      <Tag type={t.status} small />
                    </div>
                    {t.krs.map(kr => {
                      const pct = krCompletion(kr); const st = getStatus(pct);
                      const trackerVal = kr.type === "tracker"
                        ? (allOkrSubs.some(s => s.krId === kr.id && s.answer !== null) || (kr.actual != null && kr.actual !== 0)
                            ? `${fmt(kr.actual)}${kr.unit ? ` ${kr.unit}` : ""}` : null)
                        : null;
                      return (
                      <div key={kr.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0 5px 12px", borderTop: `1px solid ${T.border}`, fontSize: 13 }}>
                        <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim, width: 50, flexShrink: 0 }}>{kr.id}</span>
                        <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.label}</span>
                        {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Tracker</span>}
                        {kr.type !== "tracker" && kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                        {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                        {kr.type === "tracker"
                          ? <span style={{ fontSize: 12, fontFamily: F.mono, color: trackerVal ? "#7c3aed" : T.textDim, fontWeight: 700, textAlign: "right", flexShrink: 0 }}>{trackerVal ?? "N/A"}</span>
                          : <span style={{ fontSize: 12, fontFamily: F.mono, color: STATUS_THEME[st].color, fontWeight: 700, width: 40, textAlign: "right" }}>{pct.toFixed(0)}%</span>}
                        {kr.type === "tracker" ? <span style={{ width: 100, flexShrink: 0 }} /> : <div style={{ width: 100, flexShrink: 0 }}><Bar value={pct} status={st} h={5} /></div>}
                        {kr.type !== "tracker" && <Tag type={st} small />}
                      </div>
                    ); })}
                  </Card>
                ))}
              </>)}
            </Pane>
          </>);
        })()}

        {page === "dashboard" && (<>
          <Header title={`${dept?.name || "Team"} Dashboard`} sub={`${dept?.college} · Manager view`} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Dept Completion"   value={`${calcRate(dept?.krs || []).toFixed(1)}%`} status={getStatus(calcRate(dept?.krs || []))} sub={`Time: ${TP}%`} />
              <Metric label="My Members"        value={myMembers.length} />
              <Metric label="Pending Approvals" value={pendingOkrSubs.length} status={pendingOkrSubs.length > 0 ? "yellow" : undefined} />
            </div>
            <SectionLabel>My Team Members</SectionLabel>
            {myMembers.map(m => {
              const kd = memberData[m.id]; if (!kd) return null;
              const r = calcMemberRate(m.id, kd.krs, allOkrSubs); const s = getStatus(r);
              return (
                <Card key={m.id} style={{ padding: "14px 18px", marginBottom: 8 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "36px 1fr 55px 150px 70px", alignItems: "center", gap: 12 }}>
                    <Avatar letters={m.av} size={30} />
                    <div><div style={{ fontSize: 15, fontWeight: 700 }}>{m.name}</div><div style={{ fontSize: 12, color: T.textMuted }}>{m.title}</div></div>
                    <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</span>
                    <Bar value={r} status={s} h={6} />
                    <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={s} small /></div>
                  </div>
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "checkin" && (() => {
          const PERIOD_ORDER = ["daily", "weekly", "monthly", "quarterly", "biannual", "annual"];
          const grouped = PERIOD_ORDER.map(p => ({ period: p, pending: myOkrSubs.filter(s => s.period === p && s.answer === null), answered: myOkrSubs.filter(s => s.period === p && s.answer !== null).sort((a,b) => (b.answeredAt||"").localeCompare(a.answeredAt||"")) })).filter(g => g.pending.length + g.answered.length > 0);
          const PERIOD_COLORS = { daily: T.warn, weekly: T.brand, monthly: "#A78BFA", quarterly: "#F97316", biannual: "#06B6D4", annual: T.ok };
          const currentMonthKey = currentFYMonthKey();
          const subRate = calcSubmissionRate(myOkrSubs, user.id, currentMonthKey);
          return (<>
            <Header title="OKR Check-in" sub="Answer your KPI check-ins sent by the system"
              right={subRate !== null ? <div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 12, color: T.textMuted }}>This month:</span><span style={{ fontWeight: 700, fontSize: 15, color: STATUS_THEME[getStatus(subRate)].color, fontFamily: F.mono }}>{subRate.toFixed(0)}%</span></div> : null} />
            <Pane>
              {myPendingCheckins.length > 0 && (
                <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 8, fontSize: 13, color: T.warn, fontWeight: 600, marginBottom: 16 }}>
                  {myPendingCheckins.length} pending check-in{myPendingCheckins.length !== 1 ? "s" : ""} — please respond below
                </div>
              )}
              {grouped.length === 0 && <EmptyState text="No check-ins yet. Admin will send them when due." />}
              {grouped.map(({ period, pending, answered }) => (
                <div key={period} style={{ marginBottom: 28 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${PERIOD_COLORS[period]}` }}>
                    <div style={{ width: 4, height: 18, background: PERIOD_COLORS[period], borderRadius: 2 }} />
                    <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>{period.charAt(0).toUpperCase() + period.slice(1)} Check-ins</span>
                    {pending.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 7px", fontSize: 11, fontWeight: 700 }}>{pending.length} pending</span>}
                  </div>
                  {pending.map(s => (
                    <Card key={s.id} style={{ padding: "14px 18px", marginBottom: 8, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : noReason?.id === s.id ? T.bad : T.warn}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                            <span style={{ fontSize: 15, fontWeight: 700 }}>{s.krLabel}</span>
                            {s.krType === "tracker" && <span style={{ fontSize: 10, fontWeight: 700, background: "#ede9fe", color: "#6d28d9", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Tracker · does not affect rate</span>}
                          </div>
                          <div style={{ fontSize: 12, color: T.textMuted }}>
                            {s.krType !== "tracker" && <span>Target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                            {s.krType === "tracker" && s.krUnit && <span>Unit: {s.krUnit}</span>}
                            <span style={{ display: "block", marginTop: 3, fontSize: 14, fontWeight: 600, color: T.text }}>Review period: {s.dateRange || (s.period === "weekly" ? s.periodKey : periodDateRange(s.period, s.periodKey))}</span>
                          </div>
                        </div>
                        {s.krType === "tracker" ? (
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            <Input value={trackerInput[s.id] || ""} onChange={e => setTrackerInput(p => ({ ...p, [s.id]: e.target.value }))} placeholder="Enter value" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                            {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                            <Btn primary small onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "submitted", actualValue: Number(trackerInput[s.id]) || 0 }); setTrackerInput(p => ({ ...p, [s.id]: "" })); }} disabled={!trackerInput[s.id]}>Record</Btn>
                          </div>
                        ) : noReason?.id !== s.id ? (
                          <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                            <button onClick={() => setNoReason({ id: s.id, reason: "", actual: "" })}
                              style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.bad, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✗ No
                            </button>
                            <button onClick={() => dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "yes", actualValue: s.krTarget })}
                              style={{ background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.ok, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✓ Yes
                            </button>
                          </div>
                        ) : null}
                      </div>
                      {s.krType !== "tracker" && noReason?.id === s.id && (
                        <div style={{ marginTop: 12, padding: "12px 14px", background: T.badDim, borderRadius: 8, border: `1px solid ${T.badBorder}` }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: T.bad, marginBottom: 8 }}>Why was this OKR not met?</div>
                          <TextArea value={noReason.reason} onChange={e => setNoReason(p => ({ ...p, reason: e.target.value }))} placeholder="Briefly explain why this target was not reached..." rows={2} />
                          <div style={{ marginTop: 10 }}>
                            <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 5 }}>Your actual value</div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <Input value={noReason.actual} onChange={e => setNoReason(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                              {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                              <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                            </div>
                          </div>
                          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 10 }}>
                            <Btn small onClick={() => setNoReason(null)}>Cancel</Btn>
                            <Btn danger small onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "no", reason: noReason.reason.trim() || null, actualValue: Number(noReason.actual) || 0 }); setNoReason(null); }}>Submit No</Btn>
                          </div>
                        </div>
                      )}
                    </Card>
                  ))}
                  {answered.length > 0 && (
                    <div style={{ marginTop: pending.length ? 10 : 0 }}>
                      <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>Answered</div>
                      {answered.slice(0, 10).map(s => (
                        <Card key={s.id} style={{ padding: "10px 14px", marginBottom: 4, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : s.answer === "yes" ? T.ok : T.bad}`, opacity: s.approval === "approved" ? 0.7 : 1 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div>
                              <span style={{ fontSize: 13, fontWeight: 600 }}>{s.krLabel}</span>
                              <span style={{ fontSize: 11, color: T.textMuted, marginLeft: 8 }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              {s.krType === "tracker"
                                ? <span style={{ fontSize: 12, fontWeight: 700, color: "#6d28d9" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                                : <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>}
                              <Tag type={s.approval === "approved" ? "approved" : s.approval === "rejected" ? "rejected" : "pending"} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                            </div>
                          </div>
                          {s.krType !== "tracker" && s.answer === "no" && s.reason && <div style={{ fontSize: 12, color: T.textSoft, marginTop: 5, paddingTop: 5, borderTop: `1px solid ${T.border}` }}>Note: {s.reason}</div>}
                        </Card>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </Pane>
          </>);
        })()}

        {page === "approvals" && (() => {
          const totalPending = pendingOkrSubs.length;
          return (<>
            <Header title="Approve Member Submissions" sub={`${totalPending} pending review`} />
            <Pane>
              {/* OKR Check-in submissions */}
              {myOkrSubsForApproval.length > 0 && (
                <div style={{ marginBottom: 28 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: T.text, marginBottom: 10, display: "flex", alignItems: "center", gap: 8 }}>
                    OKR Check-ins
                    {pendingOkrSubs.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 7px", fontSize: 11, fontWeight: 700 }}>{pendingOkrSubs.length} pending</span>}
                  </div>
                  {myOkrSubsForApproval.filter(s => s.answer !== null).sort((a,b) => { const o={pending:0,approved:1,rejected:2}; return o[a.approval]-o[b.approval]||(b.answeredAt||"").localeCompare(a.answeredAt||""); }).map(s => {
                    const mem = users.find(u => u.id === s.memberId);
                    return (
                      <Card key={s.id} style={{ padding: "12px 16px", marginBottom: 6, borderLeft: `3px solid ${s.approval === "approved" ? T.ok : s.approval === "rejected" ? T.bad : T.warn}` }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                              <Avatar letters={mem?.av || "?"} size={24} />
                              <span style={{ fontWeight: 700, fontSize: 13 }}>{mem?.name || s.memberName}</span>
                              <span style={{ fontSize: 11, color: T.textMuted, background: T.raised, borderRadius: 5, padding: "1px 5px" }}>{s.period}</span>
                              <span style={{ fontSize: 11, color: T.textMuted }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                            </div>
                            <div style={{ fontSize: 13, fontWeight: 600 }}>{s.krLabel}</div>
                            <div style={{ fontSize: 11, color: T.textMuted }}>
                              {s.krType !== "tracker" && <>{`Target: ${s.krTarget}${s.krUnit ? ` ${s.krUnit}` : ""}`}</>}
                              {s.krType === "tracker" && s.actualValue != null && <span style={{ color: "#6d28d9", fontWeight: 700 }}>Recorded: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                              {s.krType !== "tracker" && s.answer === "no" && s.actualValue != null && <span style={{ color: T.bad, marginLeft: 8, fontWeight: 700 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                              {s.krType !== "tracker" && s.answer === "yes" && <span style={{ color: T.ok, marginLeft: 8 }}>Actual: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                              <span style={{ marginLeft: 8 }}>· Answered: {s.answeredAt?.slice(0,10) || "—"}</span>
                            </div>
                            {s.krType !== "tracker" && s.answer === "no" && s.reason && <div style={{ fontSize: 11, color: T.bad, marginTop: 2, fontStyle: "italic" }}>Reason: {s.reason}</div>}
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            {s.krType === "tracker"
                              ? <span style={{ fontSize: 11, fontWeight: 700, background: "#ede9fe", color: "#6d28d9", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 6px" }}>Tracker · does not affect rate</span>
                              : <span style={{ fontSize: 13, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>}
                            {s.approval === "pending"
                              ? <div style={{ display: "flex", gap: 6 }}>
                                  <Btn danger small onClick={() => setRejectOkr({ id: s.id, actual: "" })}>Reject</Btn>
                                  <Btn primary small onClick={() => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id })}>Approve</Btn>
                                </div>
                              : <Tag type={s.approval === "approved" ? "approved" : "rejected"} label={s.approval === "approved" ? "Approved" : "Rejected"} small />}
                          </div>
                        </div>
                        {rejectOkr?.id === s.id && (
                          <div style={{ marginTop: 10, padding: "10px 12px", background: T.badDim, borderRadius: 7, border: `1px solid ${T.badBorder}` }}>
                            <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 6 }}>Enter actual value for rejection</div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                              <Input value={rejectOkr.actual} onChange={e => setRejectOkr(p => ({ ...p, actual: e.target.value }))} placeholder="Actual value" style={{ width: 120, textAlign: "right", fontFamily: F.mono }} />
                              {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                              <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                            </div>
                            <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                              <Btn small onClick={() => setRejectOkr(null)}>Cancel</Btn>
                              <Btn danger small onClick={() => { dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "rejected", approvedBy: user.id, actualValue: Number(rejectOkr.actual) || 0 }); setRejectOkr(null); }}>Confirm Reject</Btn>
                            </div>
                          </div>
                        )}
                      </Card>
                    );
                  })}
                  {myOkrSubsForApproval.filter(s => s.answer === null).length > 0 && (
                    <div style={{ fontSize: 12, color: T.textMuted, padding: "6px 10px", background: T.raised, borderRadius: 6 }}>
                      {myOkrSubsForApproval.filter(s => s.answer === null).length} check-in(s) awaiting staff response
                    </div>
                  )}
                </div>
              )}
              {myOkrSubsForApproval.filter(s => s.answer !== null).length === 0 && (
                <EmptyState text="No submissions to review yet." />
              )}
            </Pane>
          </>);
        })()}

        {page === "projects" && (<>
          <Header title="Projects" sub="Create and track team projects" />
          <Pane>
            <Card style={{ padding: 20 }}>
              <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 14 }}>Create New Project</div>
              <div style={{ display: "flex", gap: 12 }}>
                <Input value={newProj.name} onChange={e => setNewProj(p => ({ ...p, name: e.target.value }))} placeholder="Project name..." style={{ flex: 1 }} />
                <Input type="date" value={newProj.due} onChange={e => setNewProj(p => ({ ...p, due: e.target.value }))} style={{ width: 160 }} />
                <Btn primary onClick={() => { if (!newProj.name) return; dispatch({ type: "ADD_PROJECT", project: { id: `p${Date.now()}`, mgrId: user.id, name: newProj.name, status: "active", due: newProj.due || "TBD", progress: 0 } }); setNewProj({ name: "", due: "" }); }}>Create</Btn>
              </div>
            </Card>
            {myProjects.map(p => {
              const ps = p.progress >= 70 ? "green" : p.progress >= 35 ? "yellow" : "red";
              const isEditing = editProjId === p.id;
              return (
                <Card key={p.id} style={{ overflow: "hidden" }}>
                  <div style={{ padding: "14px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontSize: 16, fontWeight: 700 }}>{p.name}</div>
                      <div style={{ fontSize: 12, color: T.textMuted }}>Due: {p.due}{p.updatedDate ? ` · Last updated: ${p.updatedDate}` : ""}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <Tag type={p.status === "active" ? "pending" : "approved"} label={p.status === "active" ? "ACTIVE" : "COMPLETED"} small />
                      <Btn small onClick={() => {
                        if (isEditing) { setEditProjId(null); return; }
                        setEditProjId(p.id);
                        setEditProjForm({ progress: p.progress, status: p.status, log: p.log || "", due: p.due || "" });
                      }}>{isEditing ? "Cancel" : "Edit"}</Btn>
                    </div>
                  </div>
                  {!isEditing && (
                    <div style={{ padding: "12px 18px" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: p.log ? 10 : 0 }}>
                        <Bar value={p.progress} status={ps} h={6} />
                        <span style={{ fontFamily: F.mono, fontSize: 15, fontWeight: 800, color: STATUS_THEME[ps].color, whiteSpace: "nowrap" }}>{p.progress}%</span>
                      </div>
                      {p.log && <p style={{ margin: 0, fontSize: 14, color: T.textSoft, lineHeight: 1.6, padding: "10px 14px", background: T.raised, borderRadius: 7 }}>{p.log}</p>}
                    </div>
                  )}
                  {isEditing && (
                    <div style={{ padding: "14px 18px" }}>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Completion %</div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <input type="range" min={0} max={100} value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Number(e.target.value) }))} style={{ flex: 1 }} />
                            <Input value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 55, textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                            <span style={{ fontSize: 13, color: T.textMuted }}>%</span>
                          </div>
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                          <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 14, fontFamily: F.body }}>
                            <option value="active">Active</option>
                            <option value="completed">Completed</option>
                          </select>
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                          <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                        </div>
                      </div>
                      <div style={{ marginBottom: 12 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Log / Notes</div>
                        <TextArea value={editProjForm.log} onChange={e => setEditProjForm(f => ({ ...f, log: e.target.value }))} placeholder="Update on progress, blockers, milestones reached..." rows={3} />
                      </div>
                      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                        <Btn small onClick={() => setEditProjId(null)}>Cancel</Btn>
                        <Btn primary small onClick={() => {
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { progress: editProjForm.progress, status: editProjForm.status, log: editProjForm.log, due: editProjForm.due || p.due, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          setEditProjId(null);
                        }}>Save</Btn>
                      </div>
                    </div>
                  )}
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "members" && (<>
          <Header title="Edit Member KPIs" sub="Review and adjust KPI actuals submitted by your team" />
          <Pane>
            {myMembers.filter(m => m.role === "member").map(m => {
              const kd = memberData[m.id];
              const krs = kd?.krs || [];
              const r = calcMemberRate(m.id, krs, allOkrSubs); const s = getStatus(r);
              return (
                <Card key={m.id} style={{ overflow: "hidden" }}>
                  <div style={{ padding: "14px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}><Avatar letters={m.av} size={30} /><div><div style={{ fontSize: 15, fontWeight: 700 }}>{m.name}</div><div style={{ fontSize: 12, color: T.textMuted }}>{m.title}</div></div></div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}><span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</span><Tag type={s} /></div>
                  </div>
                  {krs.length === 0
                    ? <div style={{ padding: "14px 18px", fontSize: 13, color: T.textMuted }}>No KPI data for this member yet. Sync team KPIs to populate.</div>
                    : krs.map((kr, ki) => {
                      const cr = krCompletion(kr); const cs = getStatus(cr);
                      return (
                        <div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 80px 100px 55px 130px", padding: "9px 18px", gap: 8, alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: ki < krs.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 14 }}>
                          <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                          <span>{kr.label}</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)}</span>
                          <Input value={kr.actual} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: m.id, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                          <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span>
                          <Bar value={cr} status={cs} h={5} />
                        </div>
                      );
                    })
                  }
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "reports" && (<>
          <Header title="OKR Reports" sub="Published company-wide reports — visible to all teams" />
          <Pane>
            {(() => {
              const visibleReports = monthlyReports;
              if (visibleReports.length === 0) return <EmptyState text="No OKR reports published yet." />;
              return visibleReports.map(r => (
                <Card key={r.id} style={{ overflow: "hidden" }}>
                  <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ fontSize: 18, fontWeight: 800 }}>{r.month}</div>
                        <span style={{ fontSize: 10, fontWeight: 700, background: r.reportType === "weekly" ? T.brandDim : T.okDim, border: `1px solid ${r.reportType === "weekly" ? T.brandBorder : T.okBorder}`, color: r.reportType === "weekly" ? T.brand : T.ok, borderRadius: 10, padding: "1px 7px", textTransform: "uppercase", letterSpacing: "0.06em" }}>{r.reportType || "monthly"}</span>
                      </div>
                      <div style={{ fontSize: 12, color: T.textMuted }}>Published: {r.publishedDate}{r.periodFrom && r.periodTo ? ` · ${r.periodFrom} → ${r.periodTo}` : ""}</div>
                    </div>
                    <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
                  </div>
                  <div style={{ padding: "14px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                    <div>
                      <SectionLabel>Department Rankings</SectionLabel>
                      {r.data.deptRanks.map((d, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 14 }}>
                          <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                          <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                          <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{d.rate}%</span>
                          <Tag type={d.status} small />
                        </div>
                      ))}
                      {r.submissionRate != null && (
                        <div style={{ marginTop: 10 }}>
                          <span style={{ fontSize: 12, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "2px 8px", color: T.brand }}>Submission rate: {r.submissionRate}%</span>
                        </div>
                      )}
                      {r.notes && <div style={{ marginTop: 10, padding: "8px 12px", background: T.raised, borderRadius: 7, fontSize: 13, color: T.textSoft, lineHeight: 1.6 }}><strong>Notes:</strong> {r.notes}</div>}
                    </div>
                    <div>
                      <SectionLabel>Top Performers</SectionLabel>
                      {r.data.topPerformers.map((p, i) => <div key={i} style={{ padding: "5px 0", fontSize: 14, color: T.ok, display: "flex", alignItems: "center", gap: 6 }}><span>★</span> {p}</div>)}
                      {r.data.redFlags.length > 0 && (<><SectionLabel>Action Required</SectionLabel>{r.data.redFlags.map((f, i) => <div key={i} style={{ padding: "5px 0", fontSize: 14, color: T.bad, display: "flex", alignItems: "center", gap: 6 }}><span>⚠</span> {f}</div>)}</>)}
                    </div>
                  </div>
                </Card>
              ));
            })()}
          </Pane>
        </>)}

        {syncPrompt && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div style={{ background: T.surface, borderRadius: 16, padding: "28px 32px", width: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}>
              <div style={{ fontSize: 17, fontWeight: 700, marginBottom: 10 }}>Sync KPIs to Team Members?</div>
              <p style={{ fontSize: 14, color: T.textSoft, marginBottom: 24, lineHeight: 1.6, margin: "0 0 24px" }}>
                You updated KPIs for <strong>{syncPrompt.teamName}</strong>. Sync these changes to all team members' personal KPI lists?
              </p>
              <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
                <Btn onClick={() => setSyncPrompt(null)}>Skip</Btn>
                <Btn primary onClick={() => { mgrDoSync(syncPrompt.deptId, syncPrompt.teamId); setSyncPrompt(null); }}>Yes, Sync</Btn>
              </div>
            </div>
          </div>
        )}
        {syncNote && (
          <div style={{ position: "fixed", bottom: 24, right: 24, background: T.ok, color: "#fff", borderRadius: 12, padding: "12px 20px", zIndex: 1100, fontSize: 14, fontWeight: 700, boxShadow: "0 4px 18px rgba(0,0,0,0.18)" }}>
            <div>✓ Synced</div>
            <div style={{ fontSize: 12, fontWeight: 400, opacity: 0.85, marginTop: 2 }}>{syncNote.teamName} · {syncNote.count} member{syncNote.count !== 1 ? "s" : ""}</div>
          </div>
        )}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   MEMBER PORTAL
   ───────────────────────────────────────────────────────────── */
function MemberPortal({ user, onLogout, state, dispatch }) {
  const [page, setPageRaw] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'member' ? (p[2] || 'mykpis') : 'mykpis';
  });
  const setPage = useCallback(p => { window.history.pushState(null, '', `/member/${p}`); setPageRaw(p); }, []);
  useEffect(() => {
    if (window.location.pathname.split('/')[1] !== 'member') {
      window.history.replaceState(null, '', `/member/mykpis`);
    }
    const onPop = () => {
      const p = window.location.pathname.split('/');
      setPageRaw(p[1] === 'member' ? (p[2] || 'mykpis') : 'mykpis');
    };
    window.addEventListener('popstate', onPop);
    return () => window.removeEventListener('popstate', onPop);
  }, []);
  const [myKpiPeriod, setMyKpiPeriod] = useState("all");
  const [okrPeriod, setOkrPeriod] = useState("all");
  const [expandedMonthlyKr, setExpandedMonthlyKr] = useState(null);
  const [noReason, setNoReason] = useState(null);
  const [trackerInput, setTrackerInput] = useState({});
  const [expandedKrHistory, setExpandedKrHistory] = useState(null);
  const [histPeriod, setHistPeriod] = useState("all");

  const { memberData, monthlyReports, depts } = state;
  const kd = memberData[user.id] || { krs: [] };
  const myDept = depts.find(d => d.id === user.deptId);
  const myTeam = myDept?.teams.find(t => t.id === user.teamId);
  const mySecondTeam = user.secondTeamId ? myDept?.teams.find(t => t.id === user.secondTeamId) : null;
  const myOkrSubs = (state.okrSubmissions || []).filter(s => s.memberId === user.id);
  const myPendingCheckins = myOkrSubs.filter(s => s.answer === null);
  const rate = calcMemberRate(user.id, kd.krs, state.okrSubmissions || []); const st = getStatus(rate);
  const pendingCount = myOkrSubs.filter(s => s.answer !== null && s.approval === "pending").length;
  const navItems = [
    { id: "mykpis",       icon: "◎", label: "My OKRs"          },
    { id: "checkin",      icon: "✓", label: "OKR Check-in"     },
    { id: "okr-overview", icon: "⬛", label: "OKR Overview"     },
    { id: "history",      icon: "⊞", label: "My History"        },
    { id: "reports",      icon: "⊠", label: "OKR Reports"      },
  ];

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} pendingCounts={{ checkin: myPendingCheckins.length }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "mykpis" && (<>
          <Header title="My OKRs" sub={`${user.title} · ${currentFYQuarter()}`} right={<Tag type={st} />} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="My Completion"  value={`${rate.toFixed(1)}%`} status={st} sub={`Time: ${TP}%`} />
              <Metric label="KRs Tracked"    value={kd.krs.length} />
              <Metric label="Check-ins" value={myPendingCheckins.length === 0 ? "All Done" : `${myPendingCheckins.length} Pending`} status={myPendingCheckins.length === 0 ? "green" : "yellow"} />
              <Metric label="Pending Review" value={pendingCount} status={pendingCount > 0 ? "yellow" : undefined} />
            </div>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {["all", "weekly", "monthly", "quarterly", "biannual", "annual"].map(p => (
                <Btn key={p} small primary={myKpiPeriod === p} onClick={() => setMyKpiPeriod(p)}>
                  {p.charAt(0).toUpperCase() + p.slice(1)}
                </Btn>
              ))}
            </div>
            {kd.krs.length === 0 && (
              <div style={{ padding: "28px 20px", textAlign: "center", color: T.textMuted, background: T.raised, borderRadius: 10, border: `1px dashed ${T.border}` }}>
                <div style={{ fontSize: 32, marginBottom: 10 }}>◎</div>
                <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6, color: T.text }}>No KPIs assigned yet</div>
                {!user.teamId
                  ? <div style={{ fontSize: 13 }}>You haven't been assigned to a team. Ask your admin to assign you to a team so your KPIs can be set up.</div>
                  : <div style={{ fontSize: 13 }}>Your manager hasn't synced KRs to your profile yet. Once they do, your KPIs will appear here.</div>}
              </div>
            )}
            {(myKpiPeriod === "all" ? kd.krs : kd.krs.filter(kr => (kr.period || "monthly") === myKpiPeriod)).map(kr => {
              const r = krCompletion(kr); const s = getStatus(r);
              const isMonthly = !!kr.monthlyTargets;
              const curKey = currentFYMonthKey();
              const curTarget = isMonthly ? (kr.monthlyTargets[curKey] || 0) : null;
              const curActual = isMonthly ? ((kr.monthlyActuals || {})[curKey] || 0) : null;
              const fyMs = isMonthly ? getFYMonths() : [];
              const annSumTarget = fyMs.reduce((s, {key}) => s + (kr.monthlyTargets[key] || 0), 0);
              const annActual = fyMs.reduce((s, {key}) => s + ((kr.monthlyActuals || {})[key] || 0), 0);
              const annDream = isMonthly ? (kr.annualTarget || 0) : 0;
              const annVsSum = annSumTarget > 0 ? Math.min((annActual / annSumTarget) * 100, 100) : 0;
              const annVsDream = annDream > 0 ? Math.min((annActual / annDream) * 100, 100) : 0;
              const annSt = (annDream > 0 ? annVsDream : annVsSum) >= 80 ? "green" : (annDream > 0 ? annVsDream : annVsSum) >= 50 ? "yellow" : "red";
              return (
                <Card key={kr.id} style={{ padding: "18px 20px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: T.textDim, fontFamily: F.mono, marginBottom: 3 }}>{kr.id}</div>
                      <div style={{ fontSize: 16, fontWeight: 700 }}>{kr.label}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 10, padding: "1px 6px", whiteSpace: "nowrap" }}>Tracker · does not affect rate</span>}
                      {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 10, padding: "1px 6px", whiteSpace: "nowrap" }}>Monthly Breakdown</span>}
                      <span style={{ fontSize: 10, color: T.textDim, background: T.raised, padding: "1px 6px", borderRadius: 10, border: `1px solid ${T.border}`, whiteSpace: "nowrap" }}>{kr.period || "monthly"}</span>
                      {kr.type !== "tracker" && <Tag type={s} />}
                    </div>
                  </div>
                  {isMonthly ? (
                    <>
                      <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 10 }}>This month's progress</div>
                      <div style={{ display: "flex", alignItems: "flex-end", gap: 20, marginBottom: 14 }}>
                        <div><div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{fmt(curActual)}</div><div style={{ fontSize: 12, color: T.textMuted }}>{kr.operator || ">="} {fmt(curTarget)} target{kr.unit ? ` (${kr.unit})` : ""}</div></div>
                        <div style={{ flex: 1 }}><Bar value={r} status={s} h={10} /></div>
                        <div style={{ fontSize: 26, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</div>
                      </div>
                      <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} style={{ background: "none", border: `1px solid ${T.brandBorder}`, borderRadius: 7, padding: "5px 12px", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 600, fontFamily: F.body }}>
                        {expandedMonthlyKr === kr.id ? "▲ Hide" : "▼ Annual & Monthly Detail"}
                      </button>
                      {expandedMonthlyKr === kr.id && (
                        <div style={{ marginTop: 14, overflowX: "auto" }}>
                          <table style={{ borderCollapse: "collapse", width: "100%", fontSize: 12, minWidth: 900 }}>
                            <thead>
                              <tr style={{ background: T.raised }}>
                                <th style={{ textAlign: "left", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, width: 90 }}></th>
                                {fyMs.map(({ key, label }) => {
                                  const isCur = key === curKey;
                                  return (
                                    <th key={key} style={{ textAlign: "center", padding: "6px 3px", borderBottom: `2px solid ${isCur ? T.brand : T.border}`, fontSize: 11, fontWeight: isCur ? 700 : 400, color: isCur ? T.brand : T.textDim, minWidth: 56, background: isCur ? T.brandDim : T.raised }}>
                                      {label.split(" ")[0]}{isCur ? " ●" : ""}
                                    </th>
                                  );
                                })}
                                <th style={{ textAlign: "right", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, minWidth: 72, background: T.raised }}>FY Total</th>
                                <th style={{ textAlign: "center", padding: "6px 8px", borderBottom: `2px solid ${T.okBorder}`, fontSize: 11, fontWeight: 700, color: T.ok, minWidth: 88, background: T.okDim }}>Dream Target</th>
                              </tr>
                            </thead>
                            <tbody>
                              <tr>
                                <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 11, color: T.textMuted, background: T.raised, borderBottom: `1px solid ${T.border}` }}>Perf. Target</td>
                                {fyMs.map(({ key }) => {
                                  const isCur = key === curKey;
                                  const t = kr.monthlyTargets[key] || 0;
                                  return (
                                    <td key={key} style={{ padding: "4px 6px", textAlign: "right", fontFamily: F.mono, fontSize: 12, background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}`, color: T.textMuted }}>{fmt(t)}</td>
                                  );
                                })}
                                <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 12, background: T.raised, borderBottom: `1px solid ${T.border}` }}>{fmt(annSumTarget)}</td>
                                <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, borderLeft: `1px solid ${T.okBorder}`, textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: T.ok }}>{annDream > 0 ? fmt(annDream) : <span style={{ color: T.textDim, fontWeight: 400 }}>—</span>}</td>
                              </tr>
                              <tr>
                                <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 11, color: T.text, background: T.raised, borderBottom: `1px solid ${T.border}` }}>Actual</td>
                                {fyMs.map(({ key }) => {
                                  const isCur = key === curKey;
                                  const t = kr.monthlyTargets[key] || 0;
                                  const a = (kr.monthlyActuals || {})[key] || 0;
                                  const pct = t > 0 ? Math.min((a / t) * 100, 100) : null;
                                  return (
                                    <td key={key} style={{ padding: "4px 6px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 12, background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}`, color: pct !== null ? STATUS_THEME[getStatus(pct)].color : T.text }}>{fmt(a)}</td>
                                  );
                                })}
                                <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[annSt].color, background: T.raised, borderBottom: `1px solid ${T.border}` }}>{fmt(annActual)}</td>
                                <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, textAlign: "center", color: T.textDim, fontSize: 11, borderLeft: `1px solid ${T.okBorder}` }}>—</td>
                              </tr>
                              <tr>
                                <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 11, color: T.textDim, background: T.raised }}>Achievement</td>
                                {fyMs.map(({ key }) => {
                                  const t = kr.monthlyTargets[key] || 0;
                                  const a = (kr.monthlyActuals || {})[key] || 0;
                                  const pct = t > 0 ? Math.min((a / t) * 100, 100) : null;
                                  const isCur = key === curKey;
                                  return (
                                    <td key={key} style={{ padding: "4px 4px", textAlign: "center", background: isCur ? T.brandDim : "transparent" }}>
                                      {pct !== null
                                        ? <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[getStatus(pct)].color }}>{pct.toFixed(0)}%</span>
                                        : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}
                                    </td>
                                  );
                                })}
                                <td style={{ padding: "4px 10px", textAlign: "right", background: T.raised }}>
                                  {annSumTarget > 0 ? <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[annSt].color }}>{annVsSum.toFixed(0)}%</span><div style={{ fontSize: 10, color: T.textDim }}>vs. sum</div></> : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}
                                </td>
                                <td style={{ padding: "4px 8px", background: T.okDim, textAlign: "center", borderLeft: `1px solid ${T.okBorder}` }}>
                                  {annDream > 0
                                    ? <><span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[getStatus(annVsDream)].color }}>{annVsDream.toFixed(0)}%</span><div style={{ fontSize: 10, color: T.ok }}>vs. dream</div></>
                                    : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}
                                </td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      )}
                    </>
                  ) : kr.type === "tracker" ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
                      <div>
                        <div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: "#7c3aed" }}>
                          {fmt(kr.actual)}{kr.unit ? <span style={{ fontSize: 16, fontWeight: 600, color: T.textMuted, marginLeft: 6 }}>{kr.unit}</span> : null}
                        </div>
                        <div style={{ fontSize: 12, color: T.textMuted }}>Last recorded value</div>
                      </div>
                      <div style={{ background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "6px 14px" }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "#7c3aed" }}>N/A</div>
                        <div style={{ fontSize: 11, color: "#7c3aed", opacity: 0.8 }}>does not affect rate</div>
                      </div>
                    </div>
                  ) : (
                    <div style={{ display: "flex", alignItems: "flex-end", gap: 20 }}>
                      <div><div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{fmt(kr.actual)}</div><div style={{ fontSize: 12, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)} target{kr.unit ? ` (${kr.unit})` : ""}</div></div>
                      <div style={{ flex: 1 }}><Bar value={r} status={s} h={10} /></div>
                      <div>
                        <div style={{ fontSize: 26, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</div>
                        {kr.actual > kr.target && <div style={{ fontSize: 10, color: T.ok, fontWeight: 600 }}>↑ exceeded</div>}
                      </div>
                    </div>
                  )}
                  {(() => {
                    const hist = myOkrSubs.filter(s => s.krId === kr.id).sort((a, b) => (b.sentAt || "").localeCompare(a.sentAt || ""));
                    if (!hist.length) return null;
                    const isOpen = expandedKrHistory === kr.id;
                    return (
                      <div style={{ marginTop: 14, borderTop: `1px solid ${T.border}`, paddingTop: 12 }}>
                        <button onClick={() => setExpandedKrHistory(p => p === kr.id ? null : kr.id)} style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 7, padding: "4px 12px", cursor: "pointer", color: T.textDim, fontSize: 12, fontWeight: 600, fontFamily: F.body }}>
                          {isOpen ? "▲ Hide Check-in History" : `▼ Check-in History (${hist.length})`}
                        </button>
                        {isOpen && (
                          <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 5 }}>
                            {hist.map(s => {
                              const leftCol = s.approval === "approved" ? T.ok : s.approval === "rejected" ? T.bad : s.answer !== null ? T.warn : T.border;
                              const ansCol = s.answer === "yes" ? T.ok : s.answer === "no" ? T.bad : s.answer === "submitted" ? "#7c3aed" : T.textDim;
                              const ansLabel = s.answer === "yes" ? "✓ Yes" : s.answer === "no" ? "✗ No" : s.answer === "submitted" ? "Recorded" : "Not answered";
                              return (
                                <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 12px", borderRadius: 7, background: T.raised, borderLeft: `3px solid ${leftCol}`, fontSize: 13, flexWrap: "wrap" }}>
                                  <span style={{ fontWeight: 600, flex: 1, minWidth: 120 }}>{s.dateRange || s.periodKey}</span>
                                  <span style={{ fontSize: 11, color: T.textMuted }}>{s.period}</span>
                                  <span style={{ fontWeight: 700, color: ansCol, minWidth: 90, textAlign: "right" }}>{ansLabel}</span>
                                  {(s.answer === "no" || s.answer === "submitted") && s.actualValue != null && <span style={{ fontFamily: F.mono, fontSize: 12, color: s.answer === "submitted" ? "#7c3aed" : T.textMuted }}>{s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                  <Tag type={s.approval} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                                </div>
                              );
                            })}
                          </div>
                        )}
                      </div>
                    );
                  })()}
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "okr-overview" && (() => {
          if (!myDept) return (<><Header title="OKR Overview" sub="" /><Pane><EmptyState text="No department assigned." /></Pane></>);
          const PERIODS = [{ id: "all", label: "All" }, { id: "daily", label: "Daily" }, { id: "weekly", label: "Weekly" }, { id: "monthly", label: "Monthly" }, { id: "quarterly", label: "Quarterly" }, { id: "biannual", label: "Bi-Annual" }, { id: "annual", label: "Annual" }];
          const filterP = krs => {
            const byPeriod = okrPeriod === "all" ? krs : krs.filter(kr => (kr.period || "monthly") === okrPeriod);
            return byPeriod.filter(kr => kr.type !== "tracker" || kr.showInOverview !== false);
          };
          const pLabel = okrPeriod === "all" ? "All Periods" : periodDateRange(okrPeriod, okrPeriod === "weekly" ? prevPeriodKey(okrPeriod) : currentPeriodKey(okrPeriod));
          const PChip = () => <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 7px", fontWeight: 700, marginLeft: 8, textTransform: "none", letterSpacing: "normal", verticalAlign: "middle" }}>{pLabel}</span>;
          const KCOL = "50px 1fr 100px 110px 55px 130px 65px";
          const renderKrRows = (krs) => krs.map((kr, i) => {
            const pct = krCompletion(kr); const s = getStatus(pct);
            const isMonthly = !!kr.monthlyTargets;
            const curKey = currentFYMonthKey();
            const curTarget = isMonthly ? (kr.monthlyTargets[curKey] || 0) : null;
            const curActual = isMonthly ? ((kr.monthlyActuals || {})[curKey] || 0) : null;
            const fyMs = isMonthly ? getFYMonths() : [];
            const annSumTarget = fyMs.reduce((s, {key}) => s + (kr.monthlyTargets[key] || 0), 0);
            const annActual = fyMs.reduce((s, {key}) => s + ((kr.monthlyActuals || {})[key] || 0), 0);
            const annDream = isMonthly ? (kr.annualTarget || 0) : 0;
            const annVsSum = annSumTarget > 0 ? Math.min((annActual / annSumTarget) * 100, 100) : 0;
            const annVsDream = annDream > 0 ? Math.min((annActual / annDream) * 100, 100) : 0;
            const annSt = (annDream > 0 ? annVsDream : annVsSum) >= 80 ? "green" : (annDream > 0 ? annVsDream : annVsSum) >= 50 ? "yellow" : "red";
            return (
              <Fragment key={kr.id}>
              <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                <div>
                  <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                  {kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                  {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                  {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Monthly Breakdown</span>}
                  {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                </div>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? `${kr.operator||">="} ${fmt(curTarget)}` : `${kr.operator || ">="} ${fmt(kr.target)}${kr.unit ? ` ${kr.unit}` : ""}`}</span>}
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textDim }}>—</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? fmt(curActual) : fmt(kr.actual)}</span>}
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 14, fontWeight: 700, color: "#7c3aed" }}>{fmt(isMonthly ? curActual : kr.actual)}{kr.unit ? <span style={{ fontSize: 11, fontWeight: 400 }}> {kr.unit}</span> : ""}</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[s].color }}>{pct.toFixed(0)}%</span>}
                {kr.type === "tracker" ? <span /> : <Bar value={pct} status={s} h={5} />}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 4 }}>
                  {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="View all months" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "2px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 11, fontWeight: 700 }}>📅</button>}
                  {kr.type === "tracker" ? null : <Tag type={s} small />}
                </div>
              </div>
              {isMonthly && expandedMonthlyKr === kr.id && (
                <div style={{ padding: "14px 16px 16px", background: T.brandDim, borderBottom: `1px solid ${T.border}` }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: T.brand, marginBottom: 12 }}>KPI Breakdown — {kr.label}</div>
                  <div style={{ background: T.surface, borderRadius: 10, padding: "14px 16px", marginBottom: 14, border: `1px solid ${T.border}` }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10 }}>Annual Summary</div>
                    <div style={{ display: "grid", gridTemplateColumns: annDream > 0 ? "1fr 1fr 1fr" : "1fr 1fr", gap: 16, marginBottom: 12 }}>
                      {annDream > 0 && <div><div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Dream Target</div><div style={{ fontSize: 22, fontWeight: 800, fontFamily: F.mono }}>{fmt(annDream)}</div>{kr.unit && <div style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</div>}</div>}
                      <div><div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Monthly Sum Target</div><div style={{ fontSize: 22, fontWeight: 800, fontFamily: F.mono }}>{fmt(annSumTarget)}</div>{kr.unit && <div style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</div>}</div>
                      <div><div style={{ fontSize: 11, color: T.textDim, marginBottom: 4 }}>Annual Actual</div><div style={{ fontSize: 22, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[annSt].color }}>{fmt(annActual)}</div>{kr.unit && <div style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</div>}</div>
                    </div>
                    <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 5 }}>{annDream > 0 ? `${annVsDream.toFixed(1)}% of dream target · ${annVsSum.toFixed(1)}% of monthly sum` : annSumTarget > 0 ? `${annVsSum.toFixed(1)}% of monthly sum target` : "No targets set yet"}</div>
                    {(annDream > 0 || annSumTarget > 0) && <Bar value={annDream > 0 ? annVsDream : annVsSum} status={annSt} h={8} />}
                  </div>
                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8 }}>Monthly Breakdown</div>
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(6, 1fr)", gap: 8 }}>
                    {fyMs.map(({ key, label }) => {
                      const isCur = key === curKey;
                      const t = kr.monthlyTargets[key] || 0;
                      const a = (kr.monthlyActuals || {})[key] || 0;
                      const mpct = t > 0 ? Math.min((a / t) * 100, 100) : 0;
                      const mst = mpct >= 80 ? "green" : mpct >= 50 ? "yellow" : "red";
                      return (
                        <div key={key} style={{ background: T.surface, borderRadius: 8, padding: "8px 10px", border: `2px solid ${isCur ? T.brand : T.border}` }}>
                          <div style={{ fontSize: 11, fontWeight: isCur ? 700 : 400, color: isCur ? T.brand : T.textMuted, marginBottom: 4 }}>{label}{isCur ? " ●" : ""}</div>
                          <div style={{ fontSize: 12, fontFamily: F.mono, color: T.text }}>{fmt(a)}</div>
                          <div style={{ fontSize: 11, color: T.textMuted }}>{kr.operator||">="} {fmt(t)} target</div>
                          {t > 0 && <><Bar value={mpct} status={mst} h={4} /><div style={{ fontSize: 11, fontWeight: 700, color: STATUS_THEME[mst].color, marginTop: 2 }}>{mpct.toFixed(0)}%</div></>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}
              </Fragment>
            );
          });
          const renderGroup = krs => {
            const filtered = filterP(krs);
            if (filtered.length === 0) return <div style={{ fontSize: 13, color: T.textMuted, padding: "8px 0" }}>No {okrPeriod} KRs yet.</div>;
            return (
              <Card style={{ overflow: "hidden" }}>
                <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                  <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                </div>
                {renderKrRows(filtered)}
              </Card>
            );
          };
          const dKrs = filterP(myDept.krs);
          const deptRate = calcRate(dKrs); const deptStatus = getStatus(deptRate);
          const allTeamStats = myDept.teams.map(t => { const tKrs = filterP(t.krs); return { ...t, krs: tKrs, rate: calcRate(tKrs), status: getStatus(calcRate(tKrs)) }; }).filter(t => t.krs.length > 0);
          const totalKrs = dKrs.length + allTeamStats.reduce((s, t) => s + t.krs.length, 0);
          return (<>
            <Header title="OKR Overview" sub={`${myDept.name} · ${okrPeriod.charAt(0).toUpperCase() + okrPeriod.slice(1)} key results`} right={<Tag type={deptStatus} />} />
            <Pane>
              <div style={{ display: "flex", gap: 4, marginBottom: 20, paddingBottom: 12, borderBottom: `1px solid ${T.border}` }}>
                {PERIODS.map(p => <Btn key={p.id} primary={okrPeriod === p.id} small onClick={() => setOkrPeriod(p.id)}>{p.label}</Btn>)}
              </div>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                <Metric label="Dept Completion" value={`${deptRate.toFixed(1)}%`} status={deptStatus} sub={`Target: ${TP}%`} />
                <Metric label="KRs this period" value={totalKrs} />
                {myTeam && <Metric label={mySecondTeam ? "My Teams" : "My Team"} value={mySecondTeam ? `${myTeam.name} / ${mySecondTeam.name}` : myTeam.name} />}
              </div>
              <SectionLabel>Department Key Results<PChip /></SectionLabel>
              {renderGroup(myDept.krs)}
              {myTeam && (<>
                <SectionLabel>My Team — {myTeam.name}<PChip /></SectionLabel>
                {myTeam.obj && <div style={{ fontSize: 13, color: T.textMuted, marginBottom: 10 }}>Objective: {myTeam.obj}{myTeam.lead ? ` · Lead: ${myTeam.lead}` : ""}</div>}
                {renderGroup(myTeam.krs)}
              </>)}
              {mySecondTeam && (<>
                <SectionLabel>My Team — {mySecondTeam.name}<PChip /></SectionLabel>
                {mySecondTeam.obj && <div style={{ fontSize: 13, color: T.textMuted, marginBottom: 10 }}>Objective: {mySecondTeam.obj}{mySecondTeam.lead ? ` · Lead: ${mySecondTeam.lead}` : ""}</div>}
                {renderGroup(mySecondTeam.krs)}
              </>)}
              {allTeamStats.length > 0 && (<>
                <SectionLabel>All Teams Overview<PChip /></SectionLabel>
                {allTeamStats.map(t => (
                  <Card key={t.id} style={{ padding: "14px 16px", marginBottom: 10 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                      <span style={{ fontSize: 15, fontWeight: 700, flex: 1 }}>{t.name}<PChip /></span>
                      <span style={{ fontSize: 14, fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[t.status].color }}>{t.rate.toFixed(1)}%</span>
                      <div style={{ width: 100, flexShrink: 0 }}><Bar value={t.rate} status={t.status} h={5} /></div>
                      <Tag type={t.status} small />
                    </div>
                    {t.krs.map(kr => {
                      const pct = krCompletion(kr); const st = getStatus(pct);
                      const trackerVal = kr.type === "tracker"
                        ? ((state.okrSubmissions || []).some(s => s.krId === kr.id && s.answer !== null) || (kr.actual != null && kr.actual !== 0)
                            ? `${fmt(kr.actual)}${kr.unit ? ` ${kr.unit}` : ""}` : null)
                        : null;
                      return (
                      <div key={kr.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0 5px 12px", borderTop: `1px solid ${T.border}`, fontSize: 13 }}>
                        <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim, width: 50, flexShrink: 0 }}>{kr.id}</span>
                        <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.label}</span>
                        {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Tracker</span>}
                        {kr.type !== "tracker" && kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                        {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                        {kr.type === "tracker"
                          ? <span style={{ fontSize: 12, fontFamily: F.mono, color: trackerVal ? "#7c3aed" : T.textDim, fontWeight: 700, textAlign: "right", flexShrink: 0 }}>{trackerVal ?? "N/A"}</span>
                          : <span style={{ fontSize: 12, fontFamily: F.mono, color: STATUS_THEME[st].color, fontWeight: 700, width: 40, textAlign: "right" }}>{pct.toFixed(0)}%</span>}
                        {kr.type === "tracker" ? <span style={{ width: 100, flexShrink: 0 }} /> : <div style={{ width: 100, flexShrink: 0 }}><Bar value={pct} status={st} h={5} /></div>}
                        {kr.type !== "tracker" && <Tag type={st} small />}
                      </div>
                    ); })}
                  </Card>
                ))}
              </>)}
            </Pane>
          </>);
        })()}


        {page === "checkin" && (() => {
          const PERIOD_ORDER = ["daily", "weekly", "monthly", "quarterly", "biannual", "annual"];
          const grouped = PERIOD_ORDER.map(p => ({ period: p, pending: myOkrSubs.filter(s => s.period === p && s.answer === null), answered: myOkrSubs.filter(s => s.period === p && s.answer !== null).sort((a,b) => (b.answeredAt||"").localeCompare(a.answeredAt||"")) })).filter(g => g.pending.length + g.answered.length > 0);
          const PERIOD_COLORS = { daily: T.warn, weekly: T.brand, monthly: "#A78BFA", quarterly: "#F97316", biannual: "#06B6D4", annual: T.ok };
          const currentMonthKey = currentFYMonthKey();
          const subRate = calcSubmissionRate(myOkrSubs, user.id, currentMonthKey);
          return (<>
            <Header title="OKR Check-in" sub="Answer your KPI check-ins sent by the system"
              right={subRate !== null ? <div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 12, color: T.textMuted }}>This month:</span><span style={{ fontWeight: 700, fontSize: 15, color: STATUS_THEME[getStatus(subRate)].color, fontFamily: F.mono }}>{subRate.toFixed(0)}%</span></div> : null} />
            <Pane>
              {myPendingCheckins.length > 0 && (
                <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 8, fontSize: 13, color: T.warn, fontWeight: 600, marginBottom: 16 }}>
                  {myPendingCheckins.length} pending check-in{myPendingCheckins.length !== 1 ? "s" : ""} — please respond below
                </div>
              )}
              {grouped.length === 0 && <EmptyState text="No check-ins yet. Your manager will send them when due." />}
              {grouped.map(({ period, pending, answered }) => (
                <div key={period} style={{ marginBottom: 28 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${PERIOD_COLORS[period]}` }}>
                    <div style={{ width: 4, height: 18, background: PERIOD_COLORS[period], borderRadius: 2 }} />
                    <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>{period.charAt(0).toUpperCase() + period.slice(1)} Check-ins</span>
                    {pending.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 7px", fontSize: 11, fontWeight: 700 }}>{pending.length} pending</span>}
                  </div>
                  {pending.map(s => (
                    <Card key={s.id} style={{ padding: "14px 18px", marginBottom: 8, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : noReason?.id === s.id ? T.bad : T.warn}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                            <span style={{ fontSize: 15, fontWeight: 700 }}>{s.krLabel}</span>
                            {s.krType === "tracker" && <span style={{ fontSize: 10, fontWeight: 700, background: "#ede9fe", color: "#6d28d9", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Tracker · does not affect rate</span>}
                          </div>
                          <div style={{ fontSize: 12, color: T.textMuted }}>
                            {s.krType !== "tracker" && <span>Target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                            {s.krType === "tracker" && s.krUnit && <span>Unit: {s.krUnit}</span>}
                            <span style={{ display: "block", marginTop: 3, fontSize: 14, fontWeight: 600, color: T.text }}>Review period: {s.dateRange || (s.period === "weekly" ? s.periodKey : periodDateRange(s.period, s.periodKey))}</span>
                          </div>
                        </div>
                        {s.krType === "tracker" ? (
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            <Input value={trackerInput[s.id] || ""} onChange={e => setTrackerInput(p => ({ ...p, [s.id]: e.target.value }))} placeholder="Enter value" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                            {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                            <Btn primary small onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "submitted", actualValue: Number(trackerInput[s.id]) || 0 }); setTrackerInput(p => ({ ...p, [s.id]: "" })); }} disabled={!trackerInput[s.id]}>Record</Btn>
                          </div>
                        ) : noReason?.id !== s.id ? (
                          <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                            <button onClick={() => setNoReason({ id: s.id, reason: "", actual: "" })}
                              style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.bad, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✗ No
                            </button>
                            <button onClick={() => dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "yes", actualValue: s.krTarget })}
                              style={{ background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.ok, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✓ Yes
                            </button>
                          </div>
                        ) : null}
                      </div>
                      {s.krType !== "tracker" && noReason?.id === s.id && (
                        <div style={{ marginTop: 12, padding: "12px 14px", background: T.badDim, borderRadius: 8, border: `1px solid ${T.badBorder}` }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: T.bad, marginBottom: 8 }}>Why was this OKR not met?</div>
                          <TextArea value={noReason.reason} onChange={e => setNoReason(p => ({ ...p, reason: e.target.value }))} placeholder="Briefly explain why this target was not reached..." rows={2} />
                          <div style={{ marginTop: 10 }}>
                            <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 5 }}>Your actual value</div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <Input value={noReason.actual} onChange={e => setNoReason(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                              {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                              <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                            </div>
                          </div>
                          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 10 }}>
                            <Btn small onClick={() => setNoReason(null)}>Cancel</Btn>
                            <Btn danger small onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "no", reason: noReason.reason.trim() || null, actualValue: Number(noReason.actual) || 0 }); setNoReason(null); }}>Submit No</Btn>
                          </div>
                        </div>
                      )}
                    </Card>
                  ))}
                  {answered.length > 0 && (
                    <div style={{ marginTop: pending.length ? 10 : 0 }}>
                      <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>Answered</div>
                      {answered.slice(0, 10).map(s => (
                        <Card key={s.id} style={{ padding: "10px 14px", marginBottom: 4, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : s.answer === "yes" ? T.ok : T.bad}`, opacity: s.approval === "approved" ? 0.7 : 1 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div>
                              <span style={{ fontSize: 13, fontWeight: 600 }}>{s.krLabel}</span>
                              <span style={{ fontSize: 11, color: T.textMuted, marginLeft: 8 }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              {s.krType === "tracker"
                                ? <span style={{ fontSize: 12, fontWeight: 700, color: "#6d28d9" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                                : <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>}
                              <Tag type={s.approval === "approved" ? "approved" : s.approval === "rejected" ? "rejected" : "pending"} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                            </div>
                          </div>
                          {s.krType !== "tracker" && s.answer === "no" && s.reason && <div style={{ fontSize: 12, color: T.textSoft, marginTop: 5, paddingTop: 5, borderTop: `1px solid ${T.border}` }}>Note: {s.reason}</div>}
                        </Card>
                      ))}
                    </div>
                  )}
                </div>
              ))}
            </Pane>
          </>);
        })()}

        {page === "history" && (<>
          <Header title="My OKR Check-in History" sub="All check-in submissions and their approval status" />
          <Pane>
            {(() => {
              const periods = [...new Set(myOkrSubs.map(s => s.period))].filter(Boolean);
              const filtered = (histPeriod === "all" ? myOkrSubs : myOkrSubs.filter(s => s.period === histPeriod))
                .slice().sort((a, b) => (b.sentAt || "").localeCompare(a.sentAt || ""));
              return (<>
                {periods.length > 0 && (
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 16 }}>
                    {["all", ...periods].map(p => (
                      <button key={p} onClick={() => setHistPeriod(p)} style={{ padding: "4px 14px", borderRadius: 20, fontSize: 12, fontWeight: 600, cursor: "pointer", background: histPeriod === p ? T.brand : T.raised, color: histPeriod === p ? "#fff" : T.textDim, border: `1px solid ${histPeriod === p ? T.brand : T.border}` }}>
                        {p === "all" ? "All Periods" : p.charAt(0).toUpperCase() + p.slice(1)}
                      </button>
                    ))}
                  </div>
                )}
                {filtered.length === 0 && <EmptyState text="No check-in submissions yet." />}
                {filtered.map(s => {
                  const leftCol = s.approval === "approved" ? T.ok : s.approval === "rejected" ? T.bad : s.answer !== null ? T.warn : T.border;
                  const ansCol = s.answer === "yes" ? T.ok : s.answer === "no" ? T.bad : s.answer === "submitted" ? "#7c3aed" : T.textDim;
                  const ansLabel = s.answer === "yes" ? "✓ Met target" : s.answer === "no" ? "✗ Missed target" : s.answer === "submitted" ? "Recorded" : "Not answered yet";
                  return (
                    <Card key={s.id} style={{ padding: "14px 18px", marginBottom: 8, borderLeft: `3px solid ${leftCol}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontWeight: 700, fontSize: 14, marginBottom: 4 }}>{s.krLabel}</div>
                          <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                            <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 6px", fontWeight: 700 }}>{s.period}</span>
                            <span style={{ fontSize: 12, color: T.textMuted }}>{s.dateRange || s.periodKey}</span>
                            {s.krType === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Tracker</span>}
                          </div>
                          <div style={{ marginTop: 8, fontSize: 13, fontWeight: 700, color: ansCol }}>{ansLabel}</div>
                          {s.answer === "no" && s.actualValue != null && (
                            <div style={{ fontSize: 12, color: T.textMuted, marginTop: 3 }}>
                              Actual: <span style={{ fontFamily: F.mono, fontWeight: 700, color: T.bad }}>{s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                              <span style={{ margin: "0 6px" }}>·</span>
                              Target: <span style={{ fontFamily: F.mono }}>{s.krTarget != null ? s.krTarget : "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                            </div>
                          )}
                          {s.answer === "yes" && s.krTarget != null && (
                            <div style={{ fontSize: 12, color: T.textMuted, marginTop: 3 }}>
                              Target: <span style={{ fontFamily: F.mono }}>{s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                            </div>
                          )}
                          {s.answer === "submitted" && s.actualValue != null && (
                            <div style={{ fontSize: 12, color: "#7c3aed", marginTop: 3 }}>
                              Recorded: <span style={{ fontFamily: F.mono, fontWeight: 700 }}>{s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                            </div>
                          )}
                        </div>
                        <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: 5, flexShrink: 0 }}>
                          <Tag type={s.approval} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                          <span style={{ fontSize: 11, color: T.textDim }}>Sent {s.sentAt?.slice(0, 10) || "—"}</span>
                          {s.answeredAt && <span style={{ fontSize: 11, color: T.textDim }}>Answered {s.answeredAt.slice(0, 10)}</span>}
                        </div>
                      </div>
                    </Card>
                  );
                })}
              </>);
            })()}
          </Pane>
        </>)}

        {page === "reports" && (<>
          <Header title="OKR Reports" sub="Company-wide reports — published at end of each period" />
          <Pane>
            {(() => {
              const visibleReports = monthlyReports;
              if (visibleReports.length === 0) return <EmptyState text="No OKR reports published yet." />;
              return visibleReports.map(r => (
                <Card key={r.id} style={{ overflow: "hidden" }}>
                  <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between" }}>
                    <div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <div style={{ fontSize: 18, fontWeight: 800 }}>{r.month}</div>
                        <span style={{ fontSize: 10, fontWeight: 700, background: r.reportType === "weekly" ? T.brandDim : T.okDim, border: `1px solid ${r.reportType === "weekly" ? T.brandBorder : T.okBorder}`, color: r.reportType === "weekly" ? T.brand : T.ok, borderRadius: 10, padding: "1px 7px", textTransform: "uppercase", letterSpacing: "0.06em" }}>{r.reportType || "monthly"}</span>
                      </div>
                      <div style={{ fontSize: 12, color: T.textMuted }}>Published: {r.publishedDate} · Visible to everyone</div>
                    </div>
                    <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
                  </div>
                  <div style={{ padding: "14px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                    <div>
                      <SectionLabel>Department Rankings</SectionLabel>
                      {r.data.deptRanks.map((d, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 14 }}>
                          <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                          <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                          <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{d.rate}%</span>
                        </div>
                      ))}
                      {r.submissionRate != null && (
                        <div style={{ marginTop: 10 }}>
                          <span style={{ fontSize: 12, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "2px 8px", color: T.brand }}>Submission rate: {r.submissionRate}%</span>
                        </div>
                      )}
                      {r.notes && <div style={{ marginTop: 10, padding: "8px 12px", background: T.raised, borderRadius: 7, fontSize: 13, color: T.textSoft, lineHeight: 1.6 }}><strong>Notes:</strong> {r.notes}</div>}
                    </div>
                    <div>
                      <SectionLabel>Top Performers</SectionLabel>
                      {r.data.topPerformers.map((p, i) => <div key={i} style={{ padding: "4px 0", fontSize: 14, color: T.ok }}>★ {p}</div>)}
                      {r.data.redFlags?.length > 0 && (<><div style={{ marginTop: 10 }} /><SectionLabel>Needs Improvement</SectionLabel>{r.data.redFlags.map((f, i) => <div key={i} style={{ padding: "4px 0", fontSize: 14, color: T.bad }}>⚠ {f}</div>)}</>)}
                    </div>
                  </div>
                </Card>
              ));
            })()}
          </Pane>
        </>)}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   STATE MANAGEMENT
   ───────────────────────────────────────────────────────────── */
function appReducer(state, action) {
  switch (action.type) {
    case "UPDATE_KR": return { ...state, depts: state.depts.map(d => {
      if (d.id !== action.deptId) return d;
      if (!action.teamId) return { ...d, krs: d.krs.map(kr => kr.id === action.krId ? { ...kr, [action.field]: action.value } : kr) };
      return { ...d, teams: d.teams.map(t => t.id !== action.teamId ? t : { ...t, krs: t.krs.map(kr => kr.id === action.krId ? { ...kr, [action.field]: action.value } : kr) }) };
    })};
    case "UPDATE_KR_MONTHLY": {
      const updateMonthlyKr = kr => {
        if (kr.id !== action.krId) return kr;
        if (action.field === "target") return { ...kr, monthlyTargets: { ...kr.monthlyTargets, [action.monthKey]: action.value } };
        if (action.field === "actual") return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [action.monthKey]: action.value } };
        return kr;
      };
      return { ...state, depts: state.depts.map(d => {
        if (d.id !== action.deptId) return d;
        if (!action.teamId) return { ...d, krs: d.krs.map(updateMonthlyKr) };
        return { ...d, teams: d.teams.map(t => t.id !== action.teamId ? t : { ...t, krs: t.krs.map(updateMonthlyKr) }) };
      })};
    }
    case "ADD_KR": return { ...state, depts: state.depts.map(d => {
      if (d.id !== action.deptId) return d;
      if (!action.teamId) return { ...d, krs: [...d.krs, action.kr] };
      return { ...d, teams: d.teams.map(t => t.id !== action.teamId ? t : { ...t, krs: [...t.krs, action.kr] }) };
    })};
    case "REMOVE_KR": return { ...state, depts: state.depts.map(d => {
      if (d.id !== action.deptId) return d;
      if (!action.teamId) return { ...d, krs: d.krs.filter(kr => kr.id !== action.krId) };
      return { ...d, teams: d.teams.map(t => t.id !== action.teamId ? t : { ...t, krs: t.krs.filter(kr => kr.id !== action.krId) }) };
    })};
    case "SYNC_DEPT_KRS_TO_MEMBERS": {
      const dept = state.depts.find(d => d.id === action.deptId);
      if (!dept) return state;
      const memberIds = state.users.filter(u => u.deptId === action.deptId && (u.role === "member" || u.role === "manager")).map(u => u.id);
      const krsToSync = dept.krs;
      const newMemberData = { ...state.memberData };
      for (const memberId of memberIds) {
        const md = newMemberData[memberId] || { krs: [] };
        const existing = md.krs || [];
        const updated = existing.map(kr => { const dk = krsToSync.find(t => t.id === kr.id); return dk ? { ...kr, ...dk } : kr; });
        const added = krsToSync.filter(kr => !existing.some(e => e.id === kr.id));
        newMemberData[memberId] = { ...md, krs: [...updated, ...added] };
      }
      return { ...state, memberData: newMemberData };
    }
    case "SYNC_TEAM_KRS_TO_MEMBERS": {
      const dept = state.depts.find(d => d.id === action.deptId);
      const team = dept?.teams.find(t => t.id === action.teamId);
      if (!team) return state;
      // Merge members from team.members array AND users with matching teamId/secondTeamId
      const memberIds = [...new Set([
        ...(team.members || []),
        ...state.users.filter(u => u.teamId === action.teamId || u.secondTeamId === action.teamId).map(u => u.id),
      ])];
      const krsToSync = team.krs;
      const newMemberData = { ...state.memberData };
      for (const memberId of memberIds) {
        const md = newMemberData[memberId] || { krs: [] };
        const existing = md.krs || [];
        const updated = existing.map(kr => { const tk = krsToSync.find(t => t.id === kr.id); return tk ? { ...kr, ...tk } : kr; });
        const added = krsToSync.filter(kr => !existing.some(e => e.id === kr.id));
        newMemberData[memberId] = { ...md, krs: [...updated, ...added] };
      }
      return { ...state, memberData: newMemberData };
    }
    case "EDIT_REPORT":
      return { ...state, monthlyReports: state.monthlyReports.map(r => r.id === action.reportId ? { ...r, ...action.updates } : r) };
    case "SET_DEPT_CUSTOM_COLS": return { ...state, depts: state.depts.map(d => d.id === action.deptId ? { ...d, customCols: action.customCols } : d) };
    case "UPDATE_MEMBER_KR": {
      const md = state.memberData[action.memberId];
      if (!md) return state;
      return { ...state, memberData: { ...state.memberData, [action.memberId]: { ...md, krs: (md.krs || []).map(kr => kr.id === action.krId ? { ...kr, [action.field]: action.value } : kr) } } };
    }
    case "ADD_MEMBER_KR": {
      const md = state.memberData[action.memberId] || { krs: [] };
      return { ...state, memberData: { ...state.memberData, [action.memberId]: { ...md, krs: [...(md.krs || []), action.kr] } } };
    }
    case "REMOVE_MEMBER_KR": {
      const md = state.memberData[action.memberId];
      if (!md) return state;
      return { ...state, memberData: { ...state.memberData, [action.memberId]: { ...md, krs: (md.krs || []).filter(kr => kr.id !== action.krId) } } };
    }
    case "CREATE_OKR_SUBMISSIONS": {
      const seen = new Set((state.okrSubmissions || []).map(s => `${s.memberId}:${s.krId}:${s.periodKey}`));
      const fresh = action.submissions.filter(s => !seen.has(`${s.memberId}:${s.krId}:${s.periodKey}`));
      return { ...state, okrSubmissions: [...(state.okrSubmissions || []), ...fresh] };
    }
    case "ANSWER_OKR_SUBMISSION":
      return { ...state, okrSubmissions: (state.okrSubmissions || []).map(s => s.id === action.id ? { ...s, answer: action.answer, answeredAt: new Date().toISOString(), reason: action.reason || null, actualValue: action.actualValue ?? null } : s) };
    case "APPROVE_OKR_SUBMISSION": {
      const newSubs = (state.okrSubmissions || []).map(s => s.id === action.id ? { ...s, approval: action.status, approvedBy: action.approvedBy } : s);
      const sub = (state.okrSubmissions || []).find(s => s.id === action.id);
      if (!sub) return { ...state, okrSubmissions: newSubs };
      const actualToWrite = action.status === "approved"
        ? (sub.answer === "yes" ? sub.krTarget : (sub.actualValue ?? sub.krTarget))
        : (action.status === "rejected" && action.actualValue != null ? action.actualValue : null);
      if (actualToWrite === null) return { ...state, okrSubmissions: newSubs };
      const md = state.memberData[sub.memberId];
      if (!md) return { ...state, okrSubmissions: newSubs };
      // Update this member's personal KR
      const updatedMemberKrs = (md.krs || []).map(kr => {
        if (kr.id !== sub.krId) return kr;
        if (kr.monthlyTargets) {
          const mk = (sub.periodKey || "").slice(0, 7);
          return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: actualToWrite } };
        }
        return { ...kr, actual: actualToWrite };
      });
      const newMemberData = { ...state.memberData, [sub.memberId]: { ...md, krs: updatedMemberKrs } };
      // Average actual across all members with approved submissions for this KR
      const approvedMemberIds = [...new Set(newSubs.filter(s => s.krId === sub.krId && s.approval === "approved" && s.periodKey === sub.periodKey && s.deptId === sub.deptId).map(s => s.memberId))];
      const isMonthly = !!(newMemberData[sub.memberId]?.krs?.find(k => k.id === sub.krId)?.monthlyTargets);
      const mk = (sub.periodKey || "").slice(0, 7);
      const memberVals = approvedMemberIds.map(mId => {
        const kr = (newMemberData[mId]?.krs || []).find(k => k.id === sub.krId);
        if (!kr) return null;
        return isMonthly ? ((kr.monthlyActuals || {})[mk] ?? null) : (kr.actual ?? null);
      }).filter(v => v !== null);
      const teamActual = memberVals.length > 0
        ? Math.round(memberVals.reduce((a, b) => a + b, 0) / memberVals.length * 100) / 100
        : actualToWrite;
      const updateDeptKr = kr => {
        if (kr.id !== sub.krId) return kr;
        if (kr.monthlyTargets) { return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: teamActual } }; }
        return { ...kr, actual: teamActual };
      };
      const newDepts = state.depts.map(dept => {
        if (dept.id !== sub.deptId) return dept;
        return { ...dept, krs: dept.krs.map(updateDeptKr), teams: dept.teams.map(t => ({ ...t, krs: (t.krs || []).map(updateDeptKr) })) };
      });
      return { ...state, okrSubmissions: newSubs, memberData: newMemberData, depts: newDepts };
    }
    case "REMOVE_OKR_SUBMISSION":
      return { ...state, okrSubmissions: (state.okrSubmissions || []).filter(s => s.id !== action.id) };
    case "ADD_MGR_SPRINT":    return { ...state, mgrSprints: [action.sprint, ...state.mgrSprints] };
    case "REMOVE_MGR_SPRINT": return { ...state, mgrSprints: state.mgrSprints.filter(s => s.id !== action.sprintId) };
    case "ADD_PROJECT":     return { ...state, projects: [...state.projects, action.project] };
    case "UPDATE_PROJECT":  return { ...state, projects: state.projects.map(p => p.id === action.projectId ? { ...p, ...action.updates } : p) };
    case "REMOVE_PROJECT":  return { ...state, projects: state.projects.filter(p => p.id !== action.projectId) };
    case "PUBLISH_REPORT":  return { ...state, monthlyReports: [action.report, ...state.monthlyReports] };
    case "REMOVE_REPORT":   return { ...state, monthlyReports: state.monthlyReports.filter(r => r.id !== action.reportId) };

    case "ADD_USER": {
      const u = action.user;
      const newDepts = (u.role === "member" || u.role === "manager") && u.teamId
        ? state.depts.map(d => d.id !== u.deptId ? d : { ...d, teams: d.teams.map(t => t.id !== u.teamId ? t : { ...t, members: [...t.members, u.id] }) })
        : state.depts;
      const newMemberData = (u.role === "member" || u.role === "manager")
        ? { ...state.memberData, [u.id]: { krs: [] } }
        : state.memberData;
      return { ...state, users: [...state.users, u], depts: newDepts, memberData: newMemberData };
    }

    case "UPDATE_USER": {
      const prev = state.users.find(u => u.id === action.userId);
      const updated = { ...prev, ...action.updates };
      const uid = action.userId;
      const isTracked = role => role === "member" || role === "manager";
      const oldTeam = isTracked(prev?.role) ? prev.teamId : null;
      const oldDept = isTracked(prev?.role) ? prev.deptId : null;
      const newTeam = isTracked(updated.role) ? updated.teamId : null;
      const newDept = isTracked(updated.role) ? updated.deptId : null;
      const teamChanged = oldTeam !== newTeam || oldDept !== newDept;

      const oldSecondTeam = prev?.role === "member" ? (prev.secondTeamId || null) : null;
      const newSecondTeam = updated.role === "member" ? (updated.secondTeamId || null) : null;
      const secondTeamChanged = oldSecondTeam !== newSecondTeam;

      let newDepts = state.depts;
      let newMemberData = state.memberData;

      if (teamChanged) {
        // Remove from old team
        if (oldTeam && oldDept) {
          newDepts = newDepts.map(d => d.id !== oldDept ? d : {
            ...d, teams: d.teams.map(t => t.id !== oldTeam ? t : { ...t, members: t.members.filter(id => id !== uid) })
          });
        }
        // Add to new team
        if (newTeam && newDept) {
          newDepts = newDepts.map(d => d.id !== newDept ? d : {
            ...d, teams: d.teams.map(t => t.id !== newTeam ? t : { ...t, members: t.members.includes(uid) ? t.members : [...t.members, uid] })
          });
          if (!newMemberData[uid]) newMemberData = { ...newMemberData, [uid]: { krs: [] } };
        }
      }

      if (secondTeamChanged) {
        const deptId = newDept || oldDept;
        // Remove from old second team
        if (oldSecondTeam && deptId) {
          newDepts = newDepts.map(d => d.id !== deptId ? d : {
            ...d, teams: d.teams.map(t => t.id !== oldSecondTeam ? t : { ...t, members: t.members.filter(id => id !== uid) })
          });
        }
        // Add to new second team
        if (newSecondTeam && newDept) {
          newDepts = newDepts.map(d => d.id !== newDept ? d : {
            ...d, teams: d.teams.map(t => t.id !== newSecondTeam ? t : { ...t, members: t.members.includes(uid) ? t.members : [...t.members, uid] })
          });
        }
      }

      return { ...state, users: state.users.map(u => u.id === uid ? updated : u), depts: newDepts, memberData: newMemberData };
    }

    case "REMOVE_USER": {
      const { [action.userId]: _removed, ...remainingMemberData } = state.memberData;
      return {
        ...state,
        users: state.users.filter(u => u.id !== action.userId),
        memberData: remainingMemberData,
        okrSubmissions: (state.okrSubmissions || []).filter(s => s.memberId !== action.userId),
      };
    }

    case "SET_SETTINGS":
      return { ...state, settings: { ...state.settings, ...action.updates } };

    case "RENAME_DEPT":
      return { ...state, depts: state.depts.map(d => d.id === action.deptId ? { ...d, name: action.name } : d) };

    case "ADD_DEPT":
      return { ...state, depts: [...state.depts, action.dept] };

    case "UPDATE_DEPT":
      return { ...state, depts: state.depts.map(d => d.id === action.deptId ? { ...d, ...action.updates } : d) };

    case "REMOVE_DEPT":
      return { ...state, depts: state.depts.filter(d => d.id !== action.deptId) };

    case "ADD_TEAM":
      return { ...state, depts: state.depts.map(d => d.id === action.deptId ? { ...d, teams: [...d.teams, action.team] } : d) };

    case "UPDATE_TEAM":
      return { ...state, depts: state.depts.map(d => d.id === action.deptId ? { ...d, teams: d.teams.map(t => t.id === action.teamId ? { ...t, ...action.updates } : t) } : d) };

    case "REMOVE_TEAM":
      return {
        ...state,
        depts: state.depts.map(d => d.id === action.deptId ? { ...d, teams: d.teams.filter(t => t.id !== action.teamId) } : d),
        users: state.users.map(u => {
          if (u.teamId !== action.teamId && u.secondTeamId !== action.teamId && !u.teamIds?.includes(action.teamId)) return u;
          const updated = { ...u };
          if (u.teamId === action.teamId) delete updated.teamId;
          if (u.secondTeamId === action.teamId) delete updated.secondTeamId;
          if (u.teamIds?.includes(action.teamId)) updated.teamIds = u.teamIds.filter(id => id !== action.teamId);
          return updated;
        }),
      };

    default: return state;
  }
}

/* ─────────────────────────────────────────────────────────────
   ROOT APP
   ───────────────────────────────────────────────────────────── */
export default function App({ redirectAccount = null }) {
  const [user, setUser] = useState(null);
  const [msalErr, setMsalErr] = useState("");
  const [dbReady, setDbReady] = useState(false);
  const [dbError, setDbError] = useState("");
  const [syncErr, setSyncErr] = useState("");
  const syncErrTimer = useRef(null);
  const { instance, accounts } = useMsal();
  const [state, rawDispatch] = useState({
    users: INIT_USERS,
    depts: INIT_DEPTS,
    memberData: INIT_MEMBER_DATA,
    weeklySubs: INIT_WEEKLY_SUBS,
    okrSubmissions: [],
    mgrSprints: INIT_MGR_SPRINTS,
    projects: INIT_PROJECTS,
    monthlyReports: INIT_MONTHLY_REPORTS,
    settings: { id: "settings", colOrder: ["id", "label", "operator", "period", "target", "actual", "unit", "dataSource"] },
  });

  // Dispatch updates local state immediately (optimistic), then syncs to Cosmos DB in background.
  const dispatch = useCallback((action) => {
    rawDispatch(prev => {
      const next = appReducer(prev, action);
      syncChanges(prev, next).catch(err => {
        console.error("[DB sync error]", err.message);
        setSyncErr("Changes could not be saved — check your connection and try again.");
        if (syncErrTimer.current) clearTimeout(syncErrTimer.current);
        syncErrTimer.current = setTimeout(() => setSyncErr(""), 6000);
      });
      return next;
    });
  }, []); // eslint-disable-line

  // On mount: load all data from Supabase. Seed the DB with initial data if it is empty.
  useEffect(() => {
    dbGet()
      .then(data => {
        if (data.users?.length) {
          rawDispatch(() => ({
            users: data.users,
            depts: data.depts,
            memberData: Object.fromEntries((data.memberData || []).map(m => [m.id, { krs: m.krs || [] }])),
            weeklySubs: data.weeklySubs || [],
            okrSubmissions: data.okrSubmissions || [],
            mgrSprints: data.mgrSprints || [],
            projects: data.projects || [],
            monthlyReports: data.monthlyReports || [],
            settings: data.settings?.[0] || { id: "settings", colOrder: ["id", "label", "operator", "period", "target", "actual", "unit", "dataSource"] },
          }));
        } else {
          // First run — seed the database with the built-in initial data.
          dbSeed({
            users: INIT_USERS,
            depts: INIT_DEPTS,
            memberData: Object.entries(INIT_MEMBER_DATA).map(([id, d]) => ({ id, ...d })),
            weeklySubs: INIT_WEEKLY_SUBS,
            mgrSprints: INIT_MGR_SPRINTS,
            projects: INIT_PROJECTS,
            monthlyReports: INIT_MONTHLY_REPORTS,
          }).catch(console.error);
        }
        setDbReady(true);
      })
      .catch(err => {
        console.error("DB load failed, falling back to in-memory data:", err);
        setDbError("Could not connect to database. Running in offline mode — changes will not be saved.");
        setDbReady(true);
      });
  }, []); // eslint-disable-line

  const usersRef = useRef(state.users);
  useEffect(() => { usersRef.current = state.users; }, [state.users]);

  const pendingEmailRef = useRef(null);

  const routeByEmail = useCallback((email) => {
    if (!email) return;
    pendingEmailRef.current = email;
    const lc = email.toLowerCase();
    const matched = usersRef.current.find(u => u.email.toLowerCase() === lc);
    if (matched) {
      setMsalErr("");
      setUser(matched);
      pendingEmailRef.current = null;
    } else {
      setMsalErr(lc);
    }
  }, []);

  // Route from the redirect result resolved in main.jsx before render.
  // This is the primary path after loginRedirect completes.
  useEffect(() => {
    if (redirectAccount?.username) routeByEmail(redirectAccount.username);
  }, []); // eslint-disable-line

  // Route from cached accounts — covers page refresh while already signed in.
  useEffect(() => {
    if (user || accounts.length === 0) return;
    routeByEmail(accounts[0].username);
  }, [accounts]); // eslint-disable-line

  // Retry routing once Supabase data is loaded — handles users added via admin portal
  // who aren't in the hardcoded INIT_USERS seed used before DB is ready.
  useEffect(() => {
    if (dbReady && pendingEmailRef.current && !user) routeByEmail(pendingEmailRef.current);
  }, [dbReady]); // eslint-disable-line

  if (!dbReady) return <LoadingScreen error={dbError || null} />;

  if (!user) return <LoginPage onLogin={setUser} users={state.users} msalErr={msalErr} onDismissErr={() => { setMsalErr(""); try { instance.clearCache(); } catch (_) {} }} />;
  const logout = () => {
    setUser(null);
    setMsalErr("");
    try { instance.clearCache(); } catch (_) {}
  };

  // Always derive the active user from state.users so admin edits (role, title, dept, etc.) are reflected immediately.
  const activeUser = state.users.find(u => u.id === user.id) ?? user;

  const offlineBanner = dbError ? (
    <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: T.badDim, borderTop: `1px solid ${T.badBorder}`, color: T.bad, fontSize: 13, padding: "8px 20px", textAlign: "center", zIndex: 9999 }}>
      ⚠ {dbError}
    </div>
  ) : null;
  const syncErrToast = syncErr ? (
    <div style={{ position: "fixed", bottom: dbError ? 40 : 24, right: 24, background: T.bad, color: "#fff", borderRadius: 12, padding: "12px 18px", zIndex: 9999, fontSize: 13, fontWeight: 600, maxWidth: 340, boxShadow: "0 4px 18px rgba(0,0,0,0.22)", display: "flex", alignItems: "center", gap: 10 }}>
      <span style={{ fontSize: 18 }}>⚠</span>
      <span>{syncErr}</span>
    </div>
  ) : null;

  if (activeUser.role === "admin")   return <>{offlineBanner}{syncErrToast}<AdminPortal   user={activeUser} onLogout={logout} state={state} dispatch={dispatch} /></>;
  if (activeUser.role === "manager") return <>{offlineBanner}{syncErrToast}<ManagerPortal user={activeUser} onLogout={logout} state={state} dispatch={dispatch} /></>;
  return <>{offlineBanner}{syncErrToast}<MemberPortal user={activeUser} onLogout={logout} state={state} dispatch={dispatch} /></>;
}
