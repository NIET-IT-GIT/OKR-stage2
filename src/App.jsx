import { useState, useEffect, useCallback, useRef, Fragment, Component, createContext, useContext } from "react";
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
  shadow: "0 2px 6px rgba(0,0,0,0.08), 0 12px 32px rgba(0,0,0,0.09)",
  shadowSm: "0 1px 3px rgba(0,0,0,0.07), 0 4px 12px rgba(0,0,0,0.06)",
};
const F = { body: "-apple-system,'SF Pro Text',BlinkMacSystemFont,'Segoe UI',system-ui,sans-serif", mono: "'SF Mono','Fira Code','Cascadia Code',monospace" };
const MobileContext = createContext({ isMobile: false, drawerOpen: false, setDrawerOpen: () => {} });
function useIsMobile(bp = 768) {
  const [v, setV] = useState(() => typeof window !== "undefined" && window.innerWidth < bp);
  useEffect(() => {
    const h = () => setV(p => { const n = window.innerWidth < bp; return p === n ? p : n; });
    window.addEventListener("resize", h);
    return () => window.removeEventListener("resize", h);
  }, [bp]);
  return v;
}
const TP = 66.7;

const STATUS_THEME = {
  green:  { bg: T.okDim,    border: T.okBorder,    color: T.ok,        tag: "On Track" },
  yellow: { bg: T.warnDim,  border: T.warnBorder,  color: T.warn,      tag: "At Risk"  },
  red:    { bg: T.badDim,   border: T.badBorder,   color: T.bad,       tag: "Behind"   },
  blue:   { bg: T.brandDim, border: T.brandBorder, color: T.brand,     tag: "Blue"     },
  none:   { bg: T.raised,   border: T.border,      color: T.textMuted, tag: "N/A"      },
};
const APPROVAL = {
  pending:  { bg: T.warnDim,  border: T.warnBorder,  color: T.warn,  label: "Pending Review"   },
  approved: { bg: T.okDim,    border: T.okBorder,    color: T.ok,    label: "Approved"          },
  rejected: { bg: T.badDim,   border: T.badBorder,   color: T.bad,   label: "Rejected"          },
  review:   { bg: T.brandDim, border: T.brandBorder, color: T.brand, label: "Pending Approval"  },
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
  let actual, target, hasData;
  if (kr.monthlyTargets) {
    const key = currentFYMonthKey();
    target = Number(kr.monthlyTargets[key]) || 0;
    const rawActual = (kr.monthlyActuals || {})[key];
    hasData = rawActual != null;
    actual = hasData ? (Number(rawActual) || 0) : 0;
  } else {
    hasData = kr.actual != null;
    actual = hasData ? (Number(kr.actual) || 0) : 0;
    target = Number(kr.target) || 0;
  }
  if (!hasData) return 0;
  if (target === 0) {
    if (op === ">=" || op === ">") return 100;
    return actual <= 0 ? 100 : Math.min((1 / (actual + 1)) * 100, 99);
  }
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
  const scored = (krs || []).filter(kr => kr.type !== "tracker" && kr.type !== "project_profit");
  if (!scored.length) return 0;
  return scored.reduce((sum, kr) => sum + krCompletion(kr), 0) / scored.length;
}
function meetsTarget(actual, operator, target) {
  const a = Number(actual);
  const t = Number(target);
  if (isNaN(a)) return null;
  switch (operator || ">=") {
    case ">=": return a >= t;
    case "<=": return a <= t;
    case "=":  return a === t;
    case ">":  return a > t;
    case "<":  return a < t;
    default:   return a >= t;
  }
}
// Calculates completion from a submission snapshot, not live KR data.
// krTarget is already the resolved per-period value stored at send time.
function submissionCompletion(sub) {
  if (!sub || sub.krType === "tracker") return null;
  if (sub.krType === "progress") {
    const target = sub.krTarget != null ? Number(sub.krTarget) : 0;
    if (target === 0) return 100;
    return Math.min((Number(sub.actualValue || 0) / target) * 100, 100);
  }
  const op = sub.krOperator || ">=";
  const target = sub.krTarget != null ? Number(sub.krTarget) : 0;
  // answer=yes with no actualValue → member confirmed they met target → treat actual as target (100%)
  const raw = sub.actualValue != null ? sub.actualValue : (sub.answer === "yes" ? target : null);
  if (raw == null) return 0;
  const actual = Number(raw) || 0;
  if (target === 0) {
    if (op === ">=" || op === ">") return 100;
    return actual <= 0 ? 100 : Math.min((1 / (actual + 1)) * 100, 99);
  }
  switch (op) {
    case ">=": return Math.min((actual / target) * 100, 100);
    case ">":  return actual > target ? 100 : Math.min((actual / target) * 100, 100);
    case "<=": return actual <= target ? 100 : Math.min((target / actual) * 100, 100);
    case "<":  return actual < target ? 100 : Math.min((target / actual) * 100, 100);
    case "=":  return actual === target ? 100 : 0;
    default:   return Math.min((actual / target) * 100, 100);
  }
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
    if (answered.length) {
      // Use most recent answered submission snapshot; prefer approved over pending
      const approved = answered.filter(s => s.approval === "approved").sort((a, b) => (b.answeredAt || "").localeCompare(a.answeredAt || ""));
      const best = approved.length ? approved[0] : answered.slice().sort((a, b) => (b.answeredAt || "").localeCompare(a.answeredAt || ""))[0];
      scores.push(submissionCompletion(best) ?? 0);
      continue;
    }
    const latest = krSubs.filter(s => s.sentAt).sort((a, b) => b.sentAt.localeCompare(a.sentAt))[0];
    if (latest && (now - new Date(latest.sentAt).getTime()) >= 86400000) scores.push(0); // overdue → 0%
    // else < 24h grace period → excluded
  }
  if (!scores.length) return null;
  return scores.reduce((a, b) => a + b, 0) / scores.length;
}
function getStatus(r) { if (r == null) return "none"; return r >= TP ? "green" : r >= 60 ? "yellow" : "red"; }
function memberHasRateKrs(krs) { return (krs || []).some(kr => kr.type !== "tracker" && kr.type !== "project_profit"); }
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
  const relevant = okrSubs.filter(s => s.memberId === memberId && s.answer !== null && !s.managerFilled && (s.periodKey || "").slice(0, 7) === monthKey);
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
  const result = { users: [], depts: [], memberData: [], weeklySubs: [], mgrSprints: [], projects: [], monthlyReports: [], okrSubmissions: [], emailLogs: [], settings: [] };
  const PAGE = 1000;
  let offset = 0;
  while (true) {
    const { data, error } = await supabase.from("app_data").select("collection, id, doc").in("collection", Object.keys(result)).range(offset, offset + PAGE - 1);
    if (error) throw new Error(error.message);
    if (!data?.length) break;
    for (const row of data) {
      if (result[row.collection]) result[row.collection].push(row.doc);
    }
    if (data.length < PAGE) break;
    offset += PAGE;
  }
  console.log(`[DB] loaded ${Object.values(result).flat().length} rows`);
  return result;
}

async function dbUpsert(collection, item) {
  const { data: updated, error: updErr } = await supabase
    .from("app_data")
    .update({ doc: item })
    .eq("collection", collection)
    .eq("id", item.id)
    .select("id");
  if (updErr) throw new Error(`dbUpsert update (${collection}/${item.id}): ${updErr.message}`);
  if (updated?.length) { console.log(`[DB] updated ${collection}/${item.id}`); return; }
  const { data: inserted, error: insErr } = await supabase
    .from("app_data")
    .insert({ collection, id: item.id, doc: item })
    .select("id");
  if (insErr) throw new Error(`dbUpsert insert (${collection}/${item.id}): ${insErr.message}`);
  if (!inserted?.length) throw new Error(`dbUpsert insert (${collection}/${item.id}): silent failure — 0 rows written, check Supabase table permissions`);
  console.log(`[DB] inserted ${collection}/${item.id}`);
}

async function dbDelete(collection, id) {
  const { error } = await supabase.from("app_data").delete().eq("collection", collection).eq("id", id);
  if (error) throw new Error(error.message);
}

async function dbGetEnrolments() {
  const result = { enrolment_records: [], enrolment_batches: [] };
  const PAGE = 1000;
  let offset = 0;
  while (true) {
    const { data, error } = await supabase.from("app_data").select("collection, id, doc")
      .in("collection", ["enrolment_records", "enrolment_batches"])
      .range(offset, offset + PAGE - 1);
    if (error) throw new Error(error.message);
    if (!data?.length) break;
    for (const row of data) { if (result[row.collection]) result[row.collection].push(row.doc); }
    if (data.length < PAGE) break;
    offset += PAGE;
  }
  return result;
}

async function dbGetCoeData() {
  const result = { coe_records: [], coe_batches: [] };
  const PAGE = 1000;
  let offset = 0;
  while (true) {
    const { data, error } = await supabase.from("app_data").select("collection, id, doc")
      .in("collection", ["coe_records", "coe_batches"])
      .range(offset, offset + PAGE - 1);
    if (error) throw new Error(error.message);
    if (!data?.length) break;
    for (const row of data) { if (result[row.collection]) result[row.collection].push(row.doc); }
    if (data.length < PAGE) break;
    offset += PAGE;
  }
  return result;
}

async function dbBulkInsert(collection, items) {
  const CHUNK = 200;
  for (let i = 0; i < items.length; i += CHUNK) {
    const rows = items.slice(i, i + CHUNK).map(item => ({ collection, id: item.id, doc: item }));
    const { error } = await supabase.from("app_data").insert(rows);
    if (error) throw new Error(`dbBulkInsert(${collection}): ${error.message}`);
  }
}

async function dbDeleteCollection(collection) {
  const { error } = await supabase.from("app_data").delete().eq("collection", collection);
  if (error) throw new Error(`dbDeleteCollection(${collection}): ${error.message}`);
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
  for (const col of ["users", "depts", "weeklySubs", "mgrSprints", "projects", "monthlyReports", "okrSubmissions", "emailLogs"]) {
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
      padding: small ? "2px 9px" : "3px 11px",
      fontSize: small ? 10 : 11, fontWeight: 600, color: s.color, letterSpacing: "0.02em", whiteSpace: "nowrap",
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
    <span style={{ display: "inline-flex", alignItems: "center", background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 20, padding: "2px 10px", fontSize: 12, fontWeight: 600, color: cfg.color, whiteSpace: "nowrap", letterSpacing: "0.02em" }}>
      {cfg.label}
    </span>
  );
}

function Bar({ value, status, h = 8 }) {
  const c = STATUS_THEME[status]?.color || T.brand;
  const barBg = !status ? `linear-gradient(90deg, ${T.brand}, #A78BFA)` : c;
  return (
    <div style={{ flex: 1, height: h, background: "rgba(0,0,0,0.06)", borderRadius: h, overflow: "hidden", position: "relative" }}>
      <div style={{ width: `${Math.min(value, 100)}%`, height: "100%", background: barBg, borderRadius: h, transition: "width 0.6s cubic-bezier(0.4,0,0.2,1)" }} />
      <div style={{ position: "absolute", left: `${TP}%`, top: 0, bottom: 0, width: 1, background: T.textDim, opacity: 0.5 }} />
    </div>
  );
}

function Metric({ label, value, sub, status }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16, padding: "20px 24px", flex: 1, minWidth: 130, boxShadow: T.shadowSm, borderTop: status && STATUS_THEME[status]?.color ? `3px solid ${STATUS_THEME[status].color}` : undefined }}>
      <div style={{ fontSize: 13, color: T.textMuted, fontWeight: 500, letterSpacing: "0.04em", textTransform: "uppercase", marginBottom: 10 }}>{label}</div>
      <div style={{ fontSize: 30, fontWeight: 700, color: status ? STATUS_THEME[status]?.color : T.text, fontFamily: F.mono, letterSpacing: "-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize: 13, color: T.textDim, marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

function Btn({ children, primary, danger, small, disabled, onClick, style: sx }) {
  return (
    <button onClick={disabled ? undefined : onClick} style={{
      background: primary ? T.brand : danger ? T.bad : T.surface,
      color: primary || danger ? "#fff" : T.textSoft,
      border: primary || danger ? "none" : `1px solid ${T.border}`,
      borderRadius: 10, padding: small ? "7px 15px" : "10px 20px",
      fontSize: small ? 12 : 13, fontWeight: 500, cursor: disabled ? "not-allowed" : "pointer",
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
        background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
        padding: "10px 14px", color: T.text, fontSize: 15, fontFamily: F.body, outline: "none",
        transition: "border-color 0.15s, box-shadow 0.15s", boxSizing: "border-box",
        boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)", ...sx,
      }}
      onFocus={e => { e.target.style.borderColor = T.brand; e.target.style.boxShadow = `0 0 0 3px rgba(0,113,227,0.12), inset 0 1px 2px rgba(0,0,0,0.04)`; }}
      onBlur={e => { e.target.style.borderColor = T.border; e.target.style.boxShadow = "inset 0 1px 2px rgba(0,0,0,0.04)"; }}
    />
  );
}

function NumInput({ value, onChange, style: sx, ...props }) {
  const toStr = v => (v == null) ? "" : String(v);
  const [draft, setDraft] = useState(() => toStr(value));
  const syncRef = useRef(parseFloat(toStr(value)));
  useEffect(() => {
    const n = parseFloat(toStr(value));
    if ((isNaN(n) && !isNaN(syncRef.current)) || (!isNaN(n) && n !== syncRef.current)) {
      syncRef.current = n;
      setDraft(toStr(value));
    }
  }, [value]);
  return (
    <input value={draft}
      onChange={e => {
        const raw = e.target.value;
        setDraft(raw);
        const n = parseFloat(raw);
        if (!isNaN(n)) { syncRef.current = n; onChange(n); }
        else if (raw === "" || raw === "-") { syncRef.current = 0; onChange(0); }
      }}
      style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 14px", color: T.text, fontSize: 15, fontFamily: F.body, outline: "none", transition: "border-color 0.15s, box-shadow 0.15s", boxSizing: "border-box", boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)", ...sx }}
      onFocus={e => { e.target.style.borderColor = T.brand; e.target.style.boxShadow = `0 0 0 3px rgba(0,113,227,0.12), inset 0 1px 2px rgba(0,0,0,0.04)`; }}
      onBlur={e => { e.target.style.borderColor = T.border; e.target.style.boxShadow = "inset 0 1px 2px rgba(0,0,0,0.04)"; }}
      {...props} />
  );
}
function Select({ value, onChange, children, style: sx }) {
  return (
    <select value={value} onChange={onChange}
      style={{
        background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
        padding: "10px 14px", color: value ? T.text : T.textMuted,
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
        width: "100%", boxSizing: "border-box", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10,
        padding: "11px 14px", color: T.text, fontSize: 15, fontFamily: F.body, outline: "none", resize: "vertical",
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
      background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16,
      boxShadow: T.shadowSm, ...sx,
      cursor: onClick ? "pointer" : "default", transition: "box-shadow 0.18s, border-color 0.18s",
    }}
      onMouseEnter={onClick ? e => { e.currentTarget.style.boxShadow = T.shadow; e.currentTarget.style.borderColor = "rgba(0,0,0,0.12)"; } : undefined}
      onMouseLeave={onClick ? e => { e.currentTarget.style.boxShadow = T.shadowSm; e.currentTarget.style.borderColor = T.border; } : undefined}
    >{children}</div>
  );
}

function SectionLabel({ children, style: sx }) {
  return <div style={{ fontSize: 13, fontWeight: 600, color: T.textMuted, letterSpacing: "0.05em", textTransform: "uppercase", marginBottom: 10, borderLeft: `3px solid ${T.brand}`, paddingLeft: 8, ...sx }}>{children}</div>;
}

function EmptyState({ text }) {
  return <div style={{ padding: "52px 24px", textAlign: "center", color: T.textDim, fontSize: 16, fontWeight: 400 }}>{text}</div>;
}

function CountBadge({ count, color }) {
  if (!count) return null;
  return <span style={{ background: color || T.bad, color: "#fff", borderRadius: 20, padding: "1px 7px", fontSize: 12, fontWeight: 700, marginLeft: 6, letterSpacing: "0.02em" }}>{count}</span>;
}

function Side({ items, active, onSelect, user, onLogout, pendingCounts, subItems, subItemsFor, activeSubItem, onSelectSubItem }) {
  const { isMobile, drawerOpen, setDrawerOpen } = useContext(MobileContext);
  return (
    <>
      {isMobile && drawerOpen && (
        <div onClick={() => setDrawerOpen(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.4)", zIndex: 299 }} />
      )}
      <div style={{
        width: 252, background: T.glass, borderRight: `1px solid ${T.border}`,
        backdropFilter: "saturate(180%) blur(20px)", WebkitBackdropFilter: "saturate(180%) blur(20px)",
        display: "flex", flexDirection: "column", height: "100vh", flexShrink: 0,
        ...(isMobile ? { position: "fixed", top: 0, left: 0, zIndex: 300, transform: drawerOpen ? "none" : "translateX(-100%)", transition: "transform 0.25s ease", boxShadow: drawerOpen ? "4px 0 24px rgba(0,0,0,0.18)" : "none" } : {}),
      }}>
      <div style={{ padding: "22px 16px 14px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
          <div style={{ width: 40, height: 40, borderRadius: 8, background: `linear-gradient(145deg, ${T.brand}, #A78BFA)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, fontWeight: 700, color: "#fff", boxShadow: "0 2px 8px rgba(0,113,227,0.28)" }}>NIET</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 700, color: T.text, letterSpacing: "-0.01em" }}>NIET Group OKRs</div>
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
            <button onClick={() => { onSelect(item.id); if (isMobile) setDrawerOpen(false); }} style={{
              background: active === item.id ? "rgba(0,113,227,0.12)" : "transparent",
              borderTop: active === item.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
              borderRight: active === item.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
              borderBottom: active === item.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
              borderLeft: active === item.id ? `3px solid ${T.brand}` : "1px solid transparent",
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
                  <button key={sub.id} onClick={() => { onSelectSubItem(sub.id); if (isMobile) setDrawerOpen(false); }} style={{
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
    </>
  );
}

function Header({ title, sub, right }) {
  const { isMobile, setDrawerOpen } = useContext(MobileContext);
  return (
    <div style={{ padding: isMobile ? "14px 16px 12px" : "26px 36px 22px", borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", background: T.glass, backdropFilter: "saturate(180%) blur(20px)", WebkitBackdropFilter: "saturate(180%) blur(20px)", position: "sticky", top: 0, zIndex: 10 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0, flex: 1 }}>
        {isMobile && (
          <button onClick={() => setDrawerOpen(true)} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: T.textSoft, padding: "2px 6px", lineHeight: 1, flexShrink: 0, fontFamily: F.body }}>☰</button>
        )}
        <div style={{ minWidth: 0 }}>
          <h1 style={{ margin: 0, fontSize: isMobile ? 20 : 28, fontWeight: 700, color: T.text, letterSpacing: "-0.03em", whiteSpace: isMobile ? "nowrap" : "normal", overflow: "hidden", textOverflow: "ellipsis" }}>{title}</h1>
          {sub && <p style={{ margin: "3px 0 0", fontSize: isMobile ? 12 : 15, color: T.textMuted, fontWeight: 400, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{sub}</p>}
        </div>
      </div>
      {right && <div style={{ flexShrink: 0, marginLeft: 12 }}>{right}</div>}
    </div>
  );
}

function Pane({ children }) {
  const { isMobile } = useContext(MobileContext);
  return <div style={{ padding: isMobile ? "20px 16px" : "32px 36px", display: "flex", flexDirection: "column", gap: isMobile ? 16 : 24 }}>{children}</div>;
}

class FinErrorBoundary extends Component {
  constructor(props) { super(props); this.state = { error: null }; }
  static getDerivedStateFromError(err) { return { error: err }; }
  render() {
    if (this.state.error) {
      return (
        <div style={{ padding: 32, color: T.bad, fontFamily: F.body }}>
          <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 8 }}>Financial Performance failed to render</div>
          <pre style={{ fontSize: 12, whiteSpace: "pre-wrap", color: T.textMuted }}>{String(this.state.error)}</pre>
        </div>
      );
    }
    return this.props.children;
  }
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
  const isMobile = useIsMobile();

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
    <div style={{ minHeight: "100vh", display: "flex", flexDirection: isMobile ? "column" : "row", fontFamily: F.body, position: "relative", overflow: "hidden" }}>
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
      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: isMobile ? "flex-end" : "center", padding: isMobile ? "52px 28px 32px" : "60px 72px", position: "relative", zIndex: 2, opacity: show ? 1 : 0, transform: show ? "none" : "translateX(-20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: isMobile ? 24 : 44 }}>
          <div style={{ width: isMobile ? 48 : 60, height: isMobile ? 48 : 60, borderRadius: 13, background: `linear-gradient(135deg, ${T.brand}, #A78BFA)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: isMobile ? 20 : 26, fontWeight: 900, color: "#fff" }}>NIET</div>
          <div>
            <div style={{ fontSize: isMobile ? 17 : 21, fontWeight: 900, color: "#fff" }}>NIET Group OKRs</div>
            <div style={{ fontSize: 11, color: "rgba(255,255,255,0.55)", letterSpacing: "0.14em" }}>NIET · CHARLTON BROWN · RHODES · EDUCARE</div>
          </div>
        </div>
        <h1 style={{ margin: "0 0 14px", fontSize: isMobile ? 30 : 44, fontWeight: 900, lineHeight: 1.08, color: "#fff", letterSpacing: "-0.03em", maxWidth: 460 }}>
          Align goals.<br /><span style={{ color: "#A78BFA" }}>Track everyone.</span><br />Drive results.
        </h1>
        {!isMobile && (
          <p style={{ margin: 0, fontSize: 16, color: "rgba(255,255,255,0.65)", lineHeight: 1.7, maxWidth: 400 }}>
            Monthly KPI reporting, weekly outcome tracking, real-time leaderboards — full transparency from company goals down to every team member.
          </p>
        )}
        {!isMobile && (
          <div style={{ marginTop: 48, display: "flex", gap: 36 }}>
            {[{ n: "Monthly", l: "KPI Reports" }, { n: "Weekly", l: "Submissions" }, { n: "Real-time", l: "Rankings" }, { n: "100%", l: "Transparent" }].map((x, i) => (
              <div key={i}>
                <div style={{ fontSize: 18, fontWeight: 900, color: "#A78BFA", fontFamily: F.mono }}>{x.n}</div>
                <div style={{ fontSize: 11, color: "rgba(255,255,255,0.5)", letterSpacing: "0.06em", marginTop: 2 }}>{x.l}</div>
              </div>
            ))}
          </div>
        )}
      </div>

      {/* Sign-in card */}
      <div style={{ width: isMobile ? "100%" : 420, display: "flex", flexDirection: "column", justifyContent: isMobile ? "flex-start" : "center", padding: isMobile ? "0 20px 48px" : "60px 44px", position: "relative", zIndex: 2, opacity: show ? 1 : 0, transform: show ? "none" : "translateY(20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1) 0.15s" }}>
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
const BLANK_FORM = { name: "", email: "", role: "member", title: "", deptId: "", teamId: "", teamIds: [], mgrDeptIds: [], canApprovePeers: false, designatedApproverId: "" };

function UserMgmtPage({ users, depts, dispatch, currentUserId, onImpersonate }) {
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
      ...(form.role === "manager" && form.canApprovePeers && { canApprovePeers: true }),
      ...(form.designatedApproverId && { designatedApproverId: form.designatedApproverId }),
    };
    dispatch({ type: "ADD_USER", user: newUser });
    setForm(BLANK_FORM); setFormErr(""); setShowAdd(false);
  }

  function startEdit(u) {
    setEditId(u.id);
    setEditForm({ name: u.name, email: u.email, role: u.role, title: u.title || "", deptId: u.deptId || "", teamId: u.teamId || "", teamIds: u.teamIds || [], mgrDeptIds: u.mgrDeptIds || [], secondTeamId: u.secondTeamId || "", canApprovePeers: !!u.canApprovePeers, designatedApproverId: u.designatedApproverId || "" });
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
      canApprovePeers: editForm.role === "manager" ? (editForm.canApprovePeers || undefined) : undefined,
      designatedApproverId: editForm.designatedApproverId || undefined,
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
          {form.role === "manager" && (
            <div style={{ marginBottom: 12 }}>
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, cursor: "pointer", color: T.text }}>
                <input type="checkbox" checked={!!form.canApprovePeers}
                  onChange={e => setForm(p => ({ ...p, canApprovePeers: e.target.checked }))} />
                Can approve same-department managers' submissions
              </label>
              <div style={{ fontSize: 11, color: T.textDim, marginTop: 5 }}>When enabled, this manager can view and approve pending check-ins from other managers in the same department.</div>
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
        <div style={{ overflowX: "auto" }}><div style={{ minWidth: 760 }}>
        <div style={{ display: "grid", gridTemplateColumns: "32px 1fr 160px 80px 120px 130px 140px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
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
                {editForm.role === "manager" && (
                  <div style={{ marginBottom: 12 }}>
                    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, cursor: "pointer", color: T.text }}>
                      <input type="checkbox" checked={!!editForm.canApprovePeers}
                        onChange={e => setEditForm(p => ({ ...p, canApprovePeers: e.target.checked }))} />
                      Can approve same-department managers' submissions
                    </label>
                    <div style={{ fontSize: 11, color: T.textDim, marginTop: 6 }}>When enabled, this manager can view and approve pending check-ins from other managers in the same department.</div>
                  </div>
                )}
                <div style={{ marginBottom: 12 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 6 }}>Designated Approver</div>
                  <select value={editForm.designatedApproverId || ""} onChange={e => setEditForm(p => ({ ...p, designatedApproverId: e.target.value }))}
                    style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "8px 12px", color: T.text, fontSize: 13, fontFamily: F.body, outline: "none", width: "100%", maxWidth: 300 }}>
                    <option value="">— None (default team manager) —</option>
                    {users.filter(u => u.id !== editId).map(u => (
                      <option key={u.id} value={u.id}>{u.name} ({u.role})</option>
                    ))}
                  </select>
                  <div style={{ fontSize: 11, color: T.textDim, marginTop: 6 }}>When set, this user's OKR submissions will appear in the designated approver's Approvals tab.</div>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <Btn small onClick={() => setEditId(null)}>Cancel</Btn>
                  <Btn primary small onClick={saveEdit}>Save</Btn>
                </div>
              </div>
            );
          }

          return (
            <div key={u.id} style={{ display: "grid", gridTemplateColumns: "32px 1fr 160px 80px 120px 130px 140px", padding: "10px 18px", gap: 10, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
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
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end", flexWrap: "wrap" }}>
                {!isSystem && u.role === "manager" && (
                  <button onClick={() => dispatch({ type: "UPDATE_USER", userId: u.id, updates: { financeAccess: !u.financeAccess } })} style={{ background: u.financeAccess ? "#d1fae5" : T.raised, border: `1px solid ${u.financeAccess ? "#6ee7b7" : T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: u.financeAccess ? "#065f46" : T.textMuted, fontSize: 12, fontWeight: 700, fontFamily: F.body }} title="Toggle Financial Performance access">$</button>
                )}
                {!isSystem && (u.role === "member" || u.role === "manager") && (
                  <button onClick={() => dispatch({ type: "UPDATE_USER", userId: u.id, updates: { excludeFromRate: !u.excludeFromRate } })} style={{ background: u.excludeFromRate ? T.badDim : T.raised, border: `1px solid ${u.excludeFromRate ? T.badBorder : T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: u.excludeFromRate ? T.bad : T.textMuted, fontSize: 12, fontWeight: 700, fontFamily: F.body }} title={u.excludeFromRate ? "Include in dept rate" : "Exclude from dept rate"}>%</button>
                )}
                {!isSystem && u.role !== "admin" && (
                  <button onClick={() => dispatch({ type: "UPDATE_USER", userId: u.id, updates: { projectAccess: !u.projectAccess } })} style={{ background: u.projectAccess ? T.brandDim : T.raised, border: `1px solid ${u.projectAccess ? T.brandBorder : T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: u.projectAccess ? T.brand : T.textMuted, fontSize: 12, fontWeight: 700, fontFamily: F.body }} title="Toggle Project access">◫</button>
                )}
                {!isSystem && u.role !== "admin" && (
                  <button onClick={() => onImpersonate(u)} style={{ background: "#fff3e0", border: "1px solid #ffb74d", borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: "#e65100", fontSize: 12, fontWeight: 700, fontFamily: F.body }} title={`View portal as ${u.name}`}>👁</button>
                )}
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
        </div></div>
      </Card>
    </Pane>
  </>);
}

/* ─────────────────────────────────────────────────────────────
   DEPARTMENT MANAGEMENT PAGE
   ───────────────────────────────────────────────────────────── */
const BLANK_DEPT = { name: "", obj: "", head: "", college: "" };

function DeptMgmtPage({ depts, users, memberData, okrSubmissions, dispatch, onViewKrs }) {
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

  // Month-windowed submission pool — mirrors Company Overview "Monthly" logic
  const _dmNow = new Date();
  const _dmTypes = ["daily", "weekly", "monthly"];
  const dmSubs = okrSubmissions.filter(s => {
    if (!_dmTypes.includes(s.period) || !s.sentAt) return false;
    const d = new Date(s.sentAt);
    return d.getFullYear() === _dmNow.getFullYear() && d.getMonth() === _dmNow.getMonth();
  });
  const dmFiltKrs = krs => krs.filter(kr => _dmTypes.includes(kr.period || "monthly"));
  const _dmRateForMembers = members => {
    const nowMs = _dmNow.getTime();
    const rates = members.map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const krs = dmFiltKrs(kd.krs);
      const hasEligible = krs.some(kr => {
        const krSubs = dmSubs.filter(s => s.memberId === u.id && s.krId === kr.id);
        if (!krSubs.length) return false;
        if (krSubs.some(s => s.answer !== null)) return true;
        const latest = krSubs.filter(s => s.sentAt).sort((a, b) => b.sentAt.localeCompare(a.sentAt))[0];
        return latest && (nowMs - new Date(latest.sentAt).getTime()) >= 86400000;
      });
      if (!hasEligible) return null;
      return calcMemberRate(u.id, krs, dmSubs);
    }).filter(r => r !== null);
    return rates.length ? rates.reduce((a, b) => a + b, 0) / rates.length : 0;
  };
  const calcDeptRate = deptId => {
    const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === deptId);
    return _dmRateForMembers(members);
  };
  const calcTeamRate = (deptId, team) => {
    const memberSet = new Set([...(team.members || []), ...users.filter(u => u.teamId === team.id || u.secondTeamId === team.id).map(u => u.id)]);
    const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === deptId && memberSet.has(u.id));
    return _dmRateForMembers(members);
  };

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
    <Header title="OKR Management" sub="Manage department structure, teams, and key results"
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
        <div style={{ overflowX: "auto" }}><div style={{ minWidth: 650 }}>
        <div style={{ display: "grid", gridTemplateColumns: `1fr 200px 80px${onViewKrs ? " 110px" : ""} 160px`, padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
          <span>Department / Description</span><span>Head · College</span><span>Teams</span>{onViewKrs && <span>Key Results</span>}<span style={{ textAlign: "right" }}>Actions</span>
        </div>
        {depts.length === 0 && <EmptyState text="No departments yet. Add one above." />}
        {depts.map((d, i) => {
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
              <div style={{ display: "grid", gridTemplateColumns: `1fr 200px 80px${onViewKrs ? " 110px" : ""} 160px`, padding: "12px 18px", gap: 10, alignItems: "center", background: isSelected ? T.brandDim : i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <div style={{ cursor: "pointer" }} onClick={() => { setSelDept(isSelected ? null : d.id); }}>
                  <div style={{ fontWeight: 700, color: isSelected ? T.brand : T.text }}>{d.name}</div>
                  {d.obj && <div style={{ fontSize: 12, color: T.textMuted, marginTop: 2 }}>{d.obj}</div>}
                </div>
                <div>
                  <div style={{ fontSize: 13, color: T.textSoft }}>{d.head || "—"}</div>
                  <div style={{ fontSize: 12, color: T.textMuted }}>{d.college || ""}</div>
                </div>
                <span style={{ fontSize: 13, color: T.textSoft }}>{d.teams.length} team{d.teams.length !== 1 ? "s" : ""}</span>
                {onViewKrs && <button onClick={() => onViewKrs(d.id)} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, fontFamily: F.body }}>KR Editor →</button>}
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
                const r2 = calcDeptRate(d.id); const s2 = getStatus(r2);
                const deptMembers = users
                  .filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id)
                  .map(u => { const kd = memberData[u.id] || { krs: [] }; const mr = memberHasRateKrs(kd.krs) ? calcMemberRate(u.id, dmFiltKrs(kd.krs), dmSubs) : null; return { ...u, rate: mr, status: getStatus(mr) }; })
                  .sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1));
                return (
                  <div style={{ background: T.bgSoft, borderBottom: `1px solid ${T.border}`, padding: "16px 18px" }}>
                    <div style={{ display: "flex", gap: 12, marginBottom: 14 }}>
                      <Metric label="Completion" value={`${r2.toFixed(1)}%`} status={s2} />
                      <Metric label="Teams"   value={d.teams.length} />
                      <Metric label="Members" value={deptMembers.length} />
                    </div>

                    {d.krs.length > 0 && (
                      <Card style={{ overflow: "hidden", marginBottom: 14 }}>
                        <div style={{ overflowX: "auto" }}><div style={{ minWidth: 540 }}>
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
                        </div></div>
                      </Card>
                    )}

                    <Card style={{ overflow: "hidden", marginBottom: 14 }}>
                      <div style={{ overflowX: "auto" }}><div style={{ minWidth: 560 }}>
                      <div style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 120px 55px 140px 70px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                        <span>#</span><span></span><span>Name</span><span>Title</span><span style={{ textAlign: "right" }}>Rate</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                      </div>
                      {deptMembers.length === 0
                        ? <div style={{ padding: "14px 18px", fontSize: 14, color: T.textMuted }}>No members assigned to this department.</div>
                        : deptMembers.map((m, mi) => (
                          <div key={m.id} style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 120px 55px 140px 70px", padding: "9px 18px", gap: 10, alignItems: "center", background: mi % 2 ? T.raised : "transparent", borderBottom: mi < deptMembers.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 14 }}>
                            <span style={{ fontFamily: F.mono, fontWeight: 800, color: m.status === "green" ? T.ok : m.status === "red" ? T.bad : T.textMuted }}>#{mi + 1}</span>
                            <Avatar letters={m.av} size={26} />
                            <div><span style={{ fontWeight: 600 }}>{m.name}</span></div>
                            <span style={{ fontSize: 12, color: T.textMuted }}>{m.title || "—"}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[m.status].color }}>{m.rate != null ? `${m.rate.toFixed(1)}%` : "N/A"}</span>
                            <Bar value={m.rate ?? 0} status={m.status} h={4} />
                            <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={m.status} small /></div>
                          </div>
                        ))
                      }
                      </div></div>
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

                    {d.teams.map(t => { const tr = calcTeamRate(d.id, t); const ts = getStatus(tr);
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
        </div></div>
      </Card>
    </Pane>
  </>);
}

/* ─────────────────────────────────────────────────────────────
   FINANCIAL PERFORMANCE PAGE  (shared by Admin + Manager portals)
   ───────────────────────────────────────────────────────────── */
function FinancialPerformancePage({ state, dispatch }) {
  const [finTab, setFinTab] = useState("revenue");
  const [revMonth, setRevMonth] = useState(() => { const m = new Date().getMonth(); return m >= 6 ? m - 6 : m + 6; });
  const [revEditMode, setRevEditMode] = useState(false);
  const [revDraft, setRevDraft] = useState(null);
  const [npEditMode, setNpEditMode] = useState(false);
  const [npDraft, setNpDraft] = useState(null);
  const [expEditMode, setExpEditMode] = useState(false);
  const [expDraft, setExpDraft] = useState(null);

  const REV_DIVS = ["NIET", "CB", "Rhodes", "Educare"];
  const FY_MONTHS = ["Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"];
  const DIV_COLORS = { NIET: "#0071e3", CB: "#7c3aed", Rhodes: "#f97316", Educare: "#06b6d4" };
  const fmtMoney = v => v >= 1_000_000 ? `$${(v/1_000_000).toFixed(2)}M` : v >= 1_000 ? `$${(v/1_000).toFixed(1)}K` : `$${Math.round(v).toLocaleString()}`;
  const nowFYMonth = (() => { const m = new Date().getMonth(); return m >= 6 ? m - 6 : m + 6; })();
  const mkDefault = (pt, dt) => ({ pt, dt, divisions: Object.fromEntries(REV_DIVS.map(d => [d, Array(12).fill(0)])) });

  const derivedNpDivisions = (() => {
    const r = state.settings?.revenue ?? mkDefault(0, 0);
    const e = state.settings?.expense ?? mkDefault(0, 0);
    return Object.fromEntries(REV_DIVS.map(d => [d, Array(12).fill(0).map((_, i) => (r.divisions[d]?.[i] || 0) - (e.divisions[d]?.[i] || 0))]));
  })();

  const renderModule = (cfgKey, title, gradId, editMode, setEditMode, moduleDraft, setModuleDraft, defaultPt, defaultDt, noTargets, accentColor, derivedDivisions = null) => {
    const cfg = state.settings?.[cfgKey] ?? mkDefault(defaultPt, defaultDt);
    const draft = moduleDraft ?? cfg;
    const dispDivs = derivedDivisions || cfg.divisions;
    const monthlyGroup = FY_MONTHS.map((_, i) => REV_DIVS.reduce((s, d) => s + (dispDivs[d]?.[i] || 0), 0));
    const cumulative = monthlyGroup.map((_, i) => monthlyGroup.slice(0, i + 1).reduce((a, b) => a + b, 0));
    const selCum = cumulative[revMonth] || 0;
    const thisMonthTotal = monthlyGroup[revMonth] || 0;
    const ptPct = cfg.pt > 0 ? selCum / cfg.pt : 0;
    const dtPct = cfg.dt > 0 ? selCum / cfg.dt : 0;
    const divCums = REV_DIVS.map(d => (dispDivs[d] || Array(12).fill(0)).slice(0, revMonth + 1).reduce((a, b) => a + b, 0));
    const divAnnuals = REV_DIVS.map(d => (dispDivs[d] || Array(12).fill(0)).reduce((a, b) => a + b, 0));
    const CPad = { t: 28, r: 40, b: 38, l: 72 };
    const CW = 720, CH = 230;
    const PW = CW - CPad.l - CPad.r, PH = CH - CPad.t - CPad.b;
    const maxY = (noTargets
      ? Math.max(...cumulative.slice(0, revMonth + 1), ...monthlyGroup, 100) * 1.1
      : Math.max(cfg.dt * 1.08, ...cumulative.slice(0, revMonth + 1), 100) * 1.05);
    const xAt = i => CPad.l + (i / 11) * PW;
    const yAt = v => CPad.t + PH - Math.min(v / maxY, 1) * PH;
    const plotData = cumulative.slice(0, revMonth + 1);
    const areaD = plotData.length > 0 ? [`M ${xAt(0)} ${yAt(plotData[0])}`, ...plotData.slice(1).map((v, i) => `L ${xAt(i+1)} ${yAt(v)}`), `L ${xAt(plotData.length-1)} ${CPad.t+PH}`, `L ${xAt(0)} ${CPad.t+PH}`, 'Z'].join(' ') : '';
    const lineD = plotData.length > 0 ? [`M ${xAt(0)} ${yAt(plotData[0])}`, ...plotData.slice(1).map((v, i) => `L ${xAt(i+1)} ${yAt(v)}`)].join(' ') : '';
    const yTicks = [0, 0.25, 0.5, 0.75, 1].map(t => ({ v: maxY * t, y: yAt(maxY * t) }));

    return (
      <div key={cfgKey} style={{ marginBottom: 20, paddingTop: 18, borderTop: `4px solid ${accentColor}` }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: accentColor }}>{title} · Jul – {FY_MONTHS[revMonth]} FY2027</div>
          <Btn small onClick={() => {
            if (editMode) {
              if (moduleDraft) dispatch({ type: "SET_SETTINGS", updates: { [cfgKey]: moduleDraft } });
              setEditMode(false); setModuleDraft(null);
            } else {
              setModuleDraft(JSON.parse(JSON.stringify(cfg)));
              setEditMode(true);
            }
          }}>{editMode ? "✓ Save" : derivedDivisions ? "✎ Set Targets" : "✎ Edit Data"}</Btn>
        </div>

        {editMode && (
          <Card style={{ padding: 16, marginBottom: 14 }}>
            <div style={{ fontWeight: 700, fontSize: 13, marginBottom: 14, color: T.text }}>{title} Configuration</div>
            {!noTargets && (
            <div style={{ display: "flex", gap: 20, marginBottom: 18, flexWrap: "wrap" }}>
              {[["Annual " + title + " Performance Target (PT)", "pt"], ["Annual " + title + " Dream Target (DT)", "dt"]].map(([lbl, key]) => (
                <div key={key}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 5 }}>{lbl}</div>
                  <div style={{ display: "flex", alignItems: "center", gap: 5 }}>
                    <span style={{ fontSize: 14, color: T.textMuted, fontFamily: F.mono }}>$</span>
                    <Input value={draft[key] || ""} onChange={e => setModuleDraft(p => ({ ...p, [key]: Number(String(e.target.value).replace(/,/g,"")) || 0 }))} placeholder="0" style={{ width: 140, textAlign: "right", fontFamily: F.mono }} />
                  </div>
                  <div style={{ fontSize: 11, color: T.textDim, marginTop: 3 }}>{fmtMoney(draft[key] || 0)}</div>
                </div>
              ))}
            </div>
            )}
            {derivedDivisions
              ? <div style={{ padding: "10px 14px", background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 7, fontSize: 13, color: T.brand, lineHeight: 1.5 }}>Monthly values are auto-calculated as Income − Expenses. Only the annual targets above need to be set manually.</div>
              : (<>
                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 8 }}>Monthly {title} by Division ($)</div>
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ borderCollapse: "collapse", fontSize: 12, minWidth: 780 }}>
                      <thead>
                        <tr>
                          <td style={{ padding: "4px 10px 4px 4px", fontWeight: 700, color: T.textMuted, minWidth: 85 }}>Division</td>
                          {FY_MONTHS.map((m, mi) => <td key={m} style={{ padding: "4px 3px", fontWeight: 700, color: mi <= nowFYMonth ? T.text : T.textDim, textAlign: "center", minWidth: 65, fontSize: 11 }}>{m}</td>)}
                          <td style={{ padding: "4px 6px 4px 12px", fontWeight: 700, color: T.text, textAlign: "right", minWidth: 90, fontSize: 11, borderLeft: `2px solid ${T.border}`, whiteSpace: "nowrap" }}>Annual Total</td>
                        </tr>
                      </thead>
                      <tbody>
                        {REV_DIVS.map((div, di) => (
                          <tr key={div} style={{ background: di % 2 ? T.raised : "transparent" }}>
                            <td style={{ padding: "3px 10px 3px 4px", fontWeight: 700, color: DIV_COLORS[div], fontSize: 12 }}>{div}</td>
                            {FY_MONTHS.map((_, mi) => (
                              <td key={mi} style={{ padding: "2px 3px" }}>
                                <Input value={draft.divisions?.[div]?.[mi] || ""} placeholder="0"
                                  onChange={e => {
                                    const val = Number(String(e.target.value).replace(/,/g,"")) || 0;
                                    setModuleDraft(p => ({ ...p, divisions: { ...p.divisions, [div]: (p.divisions?.[div] || Array(12).fill(0)).map((v, j) => j === mi ? val : v) } }));
                                  }}
                                  style={{ width: 62, textAlign: "right", fontFamily: F.mono, fontSize: 11, padding: "3px 5px" }} />
                              </td>
                            ))}
                            {(() => { const annTot = (draft.divisions?.[div] || Array(12).fill(0)).reduce((a, b) => a + b, 0); return <td style={{ padding: "3px 6px 3px 12px", fontFamily: F.mono, fontSize: 12, fontWeight: 700, color: annTot > 0 ? T.brand : T.textDim, textAlign: "right", borderLeft: `2px solid ${T.border}`, whiteSpace: "nowrap" }}>{annTot > 0 ? fmtMoney(annTot) : "—"}</td>; })()}
                          </tr>
                        ))}
                        <tr style={{ borderTop: `2px solid ${T.border}` }}>
                          <td style={{ padding: "4px 10px 4px 4px", fontWeight: 700, color: T.textMuted, fontSize: 11 }}>Group Total</td>
                          {FY_MONTHS.map((_, mi) => {
                            const tot = REV_DIVS.reduce((s, d) => s + (draft.divisions?.[d]?.[mi] || 0), 0);
                            return <td key={mi} style={{ padding: "4px 3px", textAlign: "right", fontFamily: F.mono, fontSize: 11, fontWeight: 700, color: tot > 0 ? T.text : T.textDim }}>{tot > 0 ? fmtMoney(tot) : "—"}</td>;
                          })}
                          {(() => { const grandTot = REV_DIVS.reduce((s, d) => s + (draft.divisions?.[d] || Array(12).fill(0)).reduce((a, b) => a + b, 0), 0); return <td style={{ padding: "4px 6px 4px 12px", textAlign: "right", fontFamily: F.mono, fontSize: 12, fontWeight: 900, color: grandTot > 0 ? T.brand : T.textDim, borderLeft: `2px solid ${T.border}`, whiteSpace: "nowrap" }}>{grandTot > 0 ? fmtMoney(grandTot) : "—"}</td>; })()}
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </>)
            }
          </Card>
        )}

        <div style={{ display: "grid", gridTemplateColumns: noTargets ? "1fr 1fr" : "1fr 1fr 1fr", gap: 12, marginBottom: 14 }}>
          {noTargets && (
            <Card style={{ padding: "16px 20px" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>{title} — {FY_MONTHS[revMonth]} Only</div>
              <div style={{ fontSize: 26, fontWeight: 900, fontFamily: F.mono, color: T.text, lineHeight: 1.1 }}>{fmtMoney(thisMonthTotal)}</div>
              <div style={{ fontSize: 11, color: T.textDim, marginTop: 4 }}>Single month · {FY_MONTHS[revMonth]} FY2027</div>
            </Card>
          )}
          <Card style={{ padding: "16px 20px" }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>Cumulative {title}</div>
            <div style={{ fontSize: 26, fontWeight: 900, fontFamily: F.mono, color: T.text, lineHeight: 1.1 }}>{fmtMoney(selCum)}</div>
            <div style={{ fontSize: 11, color: T.textDim, marginTop: 4 }}>Jul – {FY_MONTHS[revMonth]} FY2027</div>
          </Card>
          {!noTargets && [["vs Performance Target (PT)", ptPct, cfg.pt, "#F59E0B"], ["vs Dream Target (DT)", dtPct, cfg.dt, "#10B981"]].map(([lbl, pct, target, lineColor]) => {
            const st = pct >= 1 ? "green" : pct >= 0.7 ? "yellow" : "red";
            return (
              <Card key={lbl} style={{ padding: "16px 20px" }}>
                <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 6 }}>{lbl}</div>
                <div style={{ fontSize: 26, fontWeight: 900, fontFamily: F.mono, color: STATUS_THEME[st].color, lineHeight: 1.1 }}>{(pct * 100).toFixed(1)}%</div>
                <div style={{ fontSize: 11, color: T.textDim, marginTop: 4, marginBottom: 8 }}>of {fmtMoney(target)}</div>
                <div style={{ height: 8, background: T.raised, borderRadius: 4, overflow: "hidden", position: "relative" }}>
                  <div style={{ position: "absolute", inset: 0, width: `${Math.min(pct * 100, 100)}%`, background: lineColor, borderRadius: 4, transition: "width 0.5s ease" }} />
                </div>
              </Card>
            );
          })}
        </div>

        <Card style={{ padding: "16px 20px", marginBottom: 14 }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: T.text, marginBottom: 4 }}>Cumulative {title} Trend</div>
          <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 12 }}>{noTargets ? "Rolling cumulative from July" : "Rolling cumulative from July — PT and DT shown as reference lines"}</div>
          <svg viewBox={`0 0 ${CW} ${CH}`} width="100%" style={{ display: "block", overflow: "visible" }}>
            <defs>
              <linearGradient id={gradId} x1="0" y1="0" x2="0" y2="1">
                <stop offset="0%" stopColor={accentColor} stopOpacity="0.3" />
                <stop offset="100%" stopColor={accentColor} stopOpacity="0.03" />
              </linearGradient>
            </defs>
            {yTicks.map(({ v, y }, i) => (
              <g key={i}>
                <line x1={CPad.l} y1={y} x2={CW - CPad.r} y2={y} stroke={T.border} strokeWidth="1" strokeDasharray={i === 0 ? "none" : "3 4"} opacity="0.7" />
                <text x={CPad.l - 6} y={y + 4} textAnchor="end" fontSize="10" fill={T.textMuted} fontFamily="monospace">{fmtMoney(v)}</text>
              </g>
            ))}
            {areaD && <path d={areaD} fill={`url(#${gradId})`} />}
            {lineD && <path d={lineD} fill="none" stroke={accentColor} strokeWidth="2.5" strokeLinejoin="round" strokeLinecap="round" />}
            {!noTargets && cfg.pt > 0 && yAt(cfg.pt) >= CPad.t && (
              <g>
                <line x1={CPad.l} y1={yAt(cfg.pt)} x2={CW - CPad.r} y2={yAt(cfg.pt)} stroke="#F59E0B" strokeWidth="1.5" strokeDasharray="7 4" />
                <rect x={CW - CPad.r + 2} y={yAt(cfg.pt) - 8} width={28} height={16} rx="3" fill="#FEF3C7" />
                <text x={CW - CPad.r + 16} y={yAt(cfg.pt) + 4} textAnchor="middle" fontSize="10" fill="#B45309" fontWeight="700" fontFamily="sans-serif">PT</text>
              </g>
            )}
            {!noTargets && cfg.dt > 0 && yAt(cfg.dt) >= CPad.t && (
              <g>
                <line x1={CPad.l} y1={yAt(cfg.dt)} x2={CW - CPad.r} y2={yAt(cfg.dt)} stroke="#10B981" strokeWidth="1.5" strokeDasharray="7 4" />
                <rect x={CW - CPad.r + 2} y={yAt(cfg.dt) - 8} width={28} height={16} rx="3" fill="#D1FAE5" />
                <text x={CW - CPad.r + 16} y={yAt(cfg.dt) + 4} textAnchor="middle" fontSize="10" fill="#065F46" fontWeight="700" fontFamily="sans-serif">DT</text>
              </g>
            )}
            <line x1={xAt(revMonth)} y1={CPad.t} x2={xAt(revMonth)} y2={CPad.t + PH} stroke={accentColor} strokeWidth="1" strokeDasharray="4 3" opacity="0.4" />
            {plotData.map((v, i) => (
              <circle key={i} cx={xAt(i)} cy={yAt(v)} r={i === revMonth ? 5.5 : 3} fill={i === revMonth ? accentColor : T.surface} stroke={accentColor} strokeWidth={i === revMonth ? 0 : 1.5} />
            ))}
            {plotData.length > 0 && (() => {
              const lx = xAt(revMonth), ly = yAt(plotData[revMonth]);
              const lgtxt = fmtMoney(plotData[revMonth]);
              const boxW = lgtxt.length * 7.5 + 12;
              const boxX = Math.min(Math.max(lx - boxW / 2, CPad.l), CW - CPad.r - boxW);
              return (
                <g>
                  <rect x={boxX} y={ly - 26} width={boxW} height={18} rx="4" fill={accentColor} />
                  <text x={boxX + boxW / 2} y={ly - 13} textAnchor="middle" fontSize="10" fill="#fff" fontWeight="700" fontFamily="monospace">{lgtxt}</text>
                </g>
              );
            })()}
            {FY_MONTHS.map((m, i) => (
              <text key={m} x={xAt(i)} y={CH - CPad.b + 16} textAnchor="middle" fontSize="11" fill={i === revMonth ? accentColor : i < revMonth ? T.text : T.textDim} fontWeight={i === revMonth ? 700 : 400} fontFamily="sans-serif" opacity={i > nowFYMonth ? 0.45 : 1}>{m}</text>
            ))}
          </svg>
          <div style={{ display: "flex", gap: 18, marginTop: 6, fontSize: 12, flexWrap: "wrap" }}>
            {(noTargets
              ? [{ color: T.brand, dash: false, lgtxt: `Cumulative ${title}` }]
              : [
                  { color: accentColor, dash: false, lgtxt: `Cumulative ${title}` },
                  { color: "#F59E0B", dash: true, lgtxt: `PT (${fmtMoney(cfg.pt)})` },
                  { color: "#10B981", dash: true, lgtxt: `DT (${fmtMoney(cfg.dt)})` },
                ]
            ).map(({ color, dash, lgtxt }) => (
              <div key={lgtxt} style={{ display: "flex", alignItems: "center", gap: 6 }}>
                <svg width="24" height="12"><line x1="0" y1="6" x2="24" y2="6" stroke={color} strokeWidth={dash ? 1.5 : 2.5} strokeDasharray={dash ? "5 3" : "none"} /></svg>
                <span style={{ color: T.textMuted }}>{lgtxt}</span>
              </div>
            ))}
          </div>
        </Card>

        <Card style={{ padding: "16px 20px" }}>
          <div style={{ fontSize: 13, fontWeight: 700, color: T.text, marginBottom: 12 }}>Division Contribution · Jul – {FY_MONTHS[revMonth]}</div>
          {selCum > 0 ? (
            <div style={{ height: 20, borderRadius: 6, overflow: "hidden", display: "flex", marginBottom: 16 }}>
              {REV_DIVS.map((div, i) => {
                const pct = selCum > 0 ? (divCums[i] / selCum) * 100 : 0;
                return pct > 0 ? <div key={div} title={`${div}: ${fmtMoney(divCums[i])} (${pct.toFixed(1)}%)`} style={{ width: `${pct}%`, background: DIV_COLORS[div], transition: "width 0.4s" }} /> : null;
              })}
            </div>
          ) : (
            <div style={{ fontSize: 13, color: T.textMuted, marginBottom: 16, fontStyle: "italic" }}>No data entered yet — click "✎ Edit Data" to add monthly figures.</div>
          )}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10 }}>
            {REV_DIVS.map((div, i) => {
              const pct = selCum > 0 ? (divCums[i] / selCum) * 100 : 0;
              return (
                <div key={div} style={{ background: T.raised, borderRadius: 8, padding: "12px 14px", borderLeft: `3px solid ${DIV_COLORS[div]}` }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: DIV_COLORS[div], marginBottom: 4 }}>{div}</div>
                  <div style={{ fontSize: 10, color: T.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 }}>Cumulative to {FY_MONTHS[revMonth]}</div>
                  <div style={{ fontSize: 19, fontWeight: 900, fontFamily: F.mono, color: T.text, lineHeight: 1.1 }}>{fmtMoney(divCums[i])}</div>
                  <div style={{ fontSize: 11, color: T.textMuted, marginTop: 3 }}>{pct > 0 ? `${pct.toFixed(1)}% of group` : "No data"}</div>
                  {divAnnuals[i] > 0 && <div style={{ marginTop: 8, paddingTop: 8, borderTop: `1px solid ${T.border}` }}><div style={{ fontSize: 10, color: T.textMuted, fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>Annual Total</div><div style={{ fontSize: 14, fontWeight: 800, fontFamily: F.mono, color: T.textMuted, marginTop: 1 }}>{fmtMoney(divAnnuals[i])}</div></div>}
                </div>
              );
            })}
          </div>
        </Card>
      </div>
    );
  };

  const FIN_MODULES = [
    { key: "revenue",   label: "Income",     accent: "#0071e3", tab: "revenue"   },
    { key: "expense",   label: "Expenses",   accent: "#f59e0b", tab: "expense"   },
    { key: "netProfit", label: "Net Profit", accent: "#10B981", tab: "netProfit" },
  ];

  return (
    <div style={{ marginTop: 8 }}>
      {/* ── Summary row: all three at a glance ── */}
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 22 }}>
        {FIN_MODULES.map(({ key, label, accent, tab: t }) => {
          const scfg = state.settings?.[key] ?? mkDefault(0, 0);
          const sdivs = key === "netProfit" ? derivedNpDivisions : scfg.divisions;
          const smg = FY_MONTHS.map((_, i) => REV_DIVS.reduce((s, d) => s + (sdivs[d]?.[i] || 0), 0));
          const scum = smg.map((_, i) => smg.slice(0, i + 1).reduce((a, b) => a + b, 0));
          const sval = scum[revMonth] || 0;
          const isActive = finTab === t;
          return (
            <div key={key} onClick={() => setFinTab(t)} style={{ background: T.surface, border: `1.5px solid ${isActive ? accent : T.border}`, borderTop: `4px solid ${accent}`, borderRadius: 10, padding: "16px 18px", cursor: "pointer", boxShadow: isActive ? `0 0 0 3px ${accent}22` : T.shadowSm, userSelect: "none" }}>
              <div style={{ fontSize: 11, fontWeight: 700, color: accent, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 6 }}>{label}</div>
              <div style={{ fontSize: 22, fontWeight: 900, fontFamily: F.mono, color: T.text, lineHeight: 1.1 }}>{fmtMoney(sval)}</div>
              <div style={{ fontSize: 11, color: T.textMuted, marginTop: 4 }}>Jul – {FY_MONTHS[revMonth]} cumulative</div>
              {isActive && <div style={{ marginTop: 8, fontSize: 11, fontWeight: 700, color: accent }}>↓ Details below</div>}
            </div>
          );
        })}
      </div>

      {/* ── Controls row: inner tab switcher + month selector ── */}
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", flexWrap: "wrap", gap: 14, marginBottom: 22 }}>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 7 }}>Metric</div>
          <div style={{ display: "flex", gap: 0, background: T.raised, borderRadius: 10, padding: 4, boxShadow: "inset 0 1px 3px rgba(0,0,0,0.08)" }}>
            {FIN_MODULES.map(({ tab: t, label, accent }) => (
              <button key={t} onClick={() => setFinTab(t)} style={{ padding: "8px 22px", fontWeight: finTab === t ? 700 : 400, background: finTab === t ? accent : "transparent", color: finTab === t ? "#fff" : T.textMuted, border: "none", borderRadius: 7, cursor: "pointer", fontSize: 13, fontFamily: F.body, transition: "background 0.15s, color 0.15s" }}>{label}</button>
            ))}
          </div>
        </div>
        <div>
          <div style={{ fontSize: 10, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 7 }}>View through month</div>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {FY_MONTHS.map((m, i) => {
              const isFuture = i > nowFYMonth;
              const isSel = i === revMonth;
              const isPast = i < revMonth;
              return (
                <button key={m} onClick={() => setRevMonth(i)} style={{ padding: "5px 13px", fontSize: 12, fontWeight: isSel ? 700 : 400, background: isSel ? T.brand : isPast ? T.brandDim : T.raised, color: isSel ? "#fff" : isPast ? T.brand : T.textMuted, border: `1px solid ${isSel ? T.brand : isPast ? T.brandBorder : T.border}`, borderRadius: 6, cursor: "pointer", fontFamily: F.body, opacity: isFuture && !isSel ? 0.5 : 1, display: "flex", alignItems: "center", gap: 3 }}>
                  {m}{i === nowFYMonth && <span style={{ width: 5, height: 5, borderRadius: "50%", background: isSel ? "rgba(255,255,255,0.7)" : T.brand, flexShrink: 0 }} />}
                </button>
              );
            })}
          </div>
        </div>
      </div>

      {/* ── Detail panel for selected module ── */}
      {finTab === "revenue"   && renderModule("revenue",   "Income",     "revAreaGrad", revEditMode, setRevEditMode, revDraft, setRevDraft, 5000000, 7000000, false, "#0071e3")}
      {finTab === "netProfit" && renderModule("netProfit", "Net Profit", "npAreaGrad",  npEditMode,  setNpEditMode,  npDraft,  setNpDraft,  2000000, 3000000, false, "#10B981", derivedNpDivisions)}
      {finTab === "expense"   && renderModule("expense",   "Expenses",   "expAreaGrad", expEditMode, setExpEditMode, expDraft, setExpDraft,  0,       0,       true,  "#f59e0b")}
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   MARKDOWN RENDERER (for AI chat)
   ───────────────────────────────────────────────────────────── */
function inlineFmt(text, key) {
  const parts = text.split(/(\*\*[^*]+\*\*|\*[^*]+\*|`[^`]+`)/);
  return parts.map((p, i) => {
    if (p.startsWith("**") && p.endsWith("**")) return <strong key={i}>{p.slice(2, -2)}</strong>;
    if (p.startsWith("*") && p.endsWith("*")) return <em key={i}>{p.slice(1, -1)}</em>;
    if (p.startsWith("`") && p.endsWith("`")) return <code key={i} style={{ fontFamily: F.mono, fontSize: 12, background: "rgba(0,0,0,0.06)", padding: "1px 4px", borderRadius: 3 }}>{p.slice(1, -1)}</code>;
    return p;
  });
}
function MdMsg({ text }) {
  const lines = text.split("\n");
  const out = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i];
    const trim = line.trim();
    // table
    if (trim.startsWith("|") && i + 1 < lines.length && lines[i + 1].trim().startsWith("|---")) {
      const headers = trim.split("|").filter(c => c.trim()).map(c => c.trim());
      i += 2;
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) {
        rows.push(lines[i].trim().split("|").filter(c => c.trim()).map(c => c.trim()));
        i++;
      }
      out.push(
        <div key={i} style={{ overflowX: "auto", margin: "10px 0" }}>
          <table style={{ borderCollapse: "collapse", fontSize: 13, width: "100%" }}>
            <thead>
              <tr>{headers.map((h, j) => <th key={j} style={{ padding: "6px 12px", borderBottom: `2px solid ${T.border}`, textAlign: "left", fontWeight: 700, color: T.textMuted, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em", whiteSpace: "nowrap" }}>{inlineFmt(h, j)}</th>)}</tr>
            </thead>
            <tbody>
              {rows.map((row, ri) => (
                <tr key={ri} style={{ background: ri % 2 ? "rgba(0,0,0,0.02)" : "transparent" }}>
                  {row.map((cell, ci) => <td key={ci} style={{ padding: "6px 12px", borderBottom: `1px solid ${T.border}`, fontSize: 13 }}>{inlineFmt(cell, ci)}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
      continue;
    }
    // hr
    if (trim === "---" || trim === "***" || trim === "___") { out.push(<hr key={i} style={{ border: "none", borderTop: `1px solid ${T.border}`, margin: "10px 0" }} />); i++; continue; }
    // headings
    if (trim.startsWith("### ")) { out.push(<div key={i} style={{ fontWeight: 700, fontSize: 13, color: T.brand, margin: "12px 0 4px", textTransform: "uppercase", letterSpacing: "0.05em" }}>{inlineFmt(trim.slice(4), i)}</div>); i++; continue; }
    if (trim.startsWith("## ")) { out.push(<div key={i} style={{ fontWeight: 800, fontSize: 15, margin: "14px 0 6px" }}>{inlineFmt(trim.slice(3), i)}</div>); i++; continue; }
    if (trim.startsWith("# ")) { out.push(<div key={i} style={{ fontWeight: 800, fontSize: 17, margin: "14px 0 8px" }}>{inlineFmt(trim.slice(2), i)}</div>); i++; continue; }
    // bullet list — collect consecutive items
    if (trim.startsWith("- ") || trim.startsWith("* ") || trim.startsWith("• ")) {
      const items = [];
      while (i < lines.length && (lines[i].trim().startsWith("- ") || lines[i].trim().startsWith("* ") || lines[i].trim().startsWith("• "))) {
        items.push(lines[i].trim().replace(/^[-*•] /, ""));
        i++;
      }
      out.push(<ul key={i} style={{ margin: "6px 0", paddingLeft: 20 }}>{items.map((it, j) => <li key={j} style={{ marginBottom: 3, lineHeight: 1.55 }}>{inlineFmt(it, j)}</li>)}</ul>);
      continue;
    }
    // numbered list
    if (/^\d+\.\s/.test(trim)) {
      const items = [];
      while (i < lines.length && /^\d+\.\s/.test(lines[i].trim())) {
        items.push(lines[i].trim().replace(/^\d+\.\s/, ""));
        i++;
      }
      out.push(<ol key={i} style={{ margin: "6px 0", paddingLeft: 20 }}>{items.map((it, j) => <li key={j} style={{ marginBottom: 3, lineHeight: 1.55 }}>{inlineFmt(it, j)}</li>)}</ol>);
      continue;
    }
    // blank line
    if (!trim) { out.push(<div key={i} style={{ height: 6 }} />); i++; continue; }
    // paragraph
    out.push(<div key={i} style={{ margin: "3px 0", lineHeight: 1.65 }}>{inlineFmt(trim, i)}</div>);
    i++;
  }
  return <div>{out}</div>;
}

/* ─────────────────────────────────────────────────────────────
   ACTION REVIEW CARD  (NIET Pilot bulk approve / reject)
   ───────────────────────────────────────────────────────────── */
function ActionReviewCard({ action, submissions, onConfirm, onCancel }) {
  const isApprove = action.type === "approve";
  const accentColor = isApprove ? T.ok : T.bad;
  const accentDim   = isApprove ? T.okDim  : T.badDim;
  const accentBdr   = isApprove ? T.okBorder : T.badBorder;

  // Pre-deselect "no"-answer submissions when approving (likely need review)
  const [skipped, setSkipped] = useState(() =>
    new Set(isApprove ? submissions.filter(s => s.answer === "no").map(s => s.id) : [])
  );
  const [done, setDone]       = useState(false);
  const [doneCount, setDoneCount] = useState(0);
  const [cancelled, setCancelled] = useState(false);

  const toAct  = submissions.filter(s => !skipped.has(s.id));
  const noCount = submissions.filter(s => s.answer === "no").length;

  if (done) return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 16px", borderRadius: 10, background: accentDim, border: `1px solid ${accentBdr}`, fontSize: 14, fontWeight: 600, color: accentColor }}>
      {isApprove ? "✓" : "✕"} {isApprove ? "Approved" : "Rejected"} {doneCount} submission{doneCount !== 1 ? "s" : ""}
    </div>
  );

  if (cancelled) return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 8, padding: "10px 16px", borderRadius: 10, background: T.raised, border: `1px solid ${T.border}`, fontSize: 14, color: T.textMuted }}>
      ✕ Action cancelled
    </div>
  );

  if (submissions.length === 0) return (
    <div style={{ padding: "14px 18px", borderRadius: 12, background: T.surface, border: `1px solid ${T.border}`, fontSize: 14, color: T.textMuted }}>
      No pending submissions found matching that criteria.
    </div>
  );

  return (
    <div style={{ borderRadius: 14, background: T.surface, border: `1px solid ${T.border}`, overflow: "hidden", maxWidth: "100%", boxShadow: T.shadowSm }}>
      {/* Header */}
      <div style={{ padding: "14px 18px", borderBottom: `1px solid ${T.border}`, background: accentDim, display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{ width: 28, height: 28, borderRadius: "50%", background: accentColor, color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 13, fontWeight: 700, flexShrink: 0 }}>{isApprove ? "✓" : "✕"}</div>
        <div>
          <div style={{ fontSize: 14, fontWeight: 700, color: T.text }}>{action.message}</div>
          <div style={{ fontSize: 12, color: T.textMuted, marginTop: 1 }}>{submissions.length} submission{submissions.length !== 1 ? "s" : ""} pending — review before confirming</div>
        </div>
      </div>

      {/* Warning: "no" answers when approving */}
      {isApprove && noCount > 0 && (
        <div style={{ padding: "9px 18px", background: T.warnDim, borderBottom: `1px solid ${T.warnBorder}`, fontSize: 13, color: T.warn, display: "flex", gap: 7, alignItems: "center" }}>
          <span>⚠️</span>
          <span>{noCount} submission{noCount !== 1 ? "s" : ""} answered "no" — pre-deselected. Check before including.</span>
        </div>
      )}

      {/* Table */}
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ background: T.raised }}>
              {["", "Member", "KR", "Answer", "Actual", "Target", "Period"].map(h => (
                <th key={h} style={{ padding: "8px 12px", textAlign: h === "Actual" || h === "Target" ? "right" : "left", fontWeight: 700, fontSize: 11, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", whiteSpace: "nowrap" }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {submissions.map((s, i) => {
              const isNo       = s.answer === "no";
              const isTracker  = s.krType === "tracker";
              const isProgress = s.krType === "progress";
              const isSkipped  = skipped.has(s.id);
              const answerColor = isTracker ? "#7c3aed" : isProgress ? T.brand : isNo ? T.bad : T.ok;
              return (
                <tr key={s.id} style={{ background: isSkipped ? "transparent" : (i % 2 ? T.raised : "transparent"), opacity: isSkipped ? 0.4 : 1, borderBottom: `1px solid ${T.border}` }}>
                  <td style={{ padding: "10px 12px", textAlign: "center", width: 36 }}>
                    <input type="checkbox" checked={!isSkipped} style={{ cursor: "pointer", accentColor }}
                      onChange={e => setSkipped(prev => { const n = new Set(prev); e.target.checked ? n.delete(s.id) : n.add(s.id); return n; })} />
                  </td>
                  <td style={{ padding: "10px 12px" }}>
                    <div style={{ fontWeight: 600, fontSize: 13 }}>{s.memberName}</div>
                    <div style={{ fontSize: 11, color: T.textMuted }}>{s.deptName}</div>
                  </td>
                  <td style={{ padding: "10px 12px", color: T.textSoft, maxWidth: 200 }}>
                    <span style={{ display: "block", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.krLabel}</span>
                    {isTracker && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", borderRadius: 6, padding: "1px 5px" }}>Tracker</span>}
                    {isProgress && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "1px 5px" }}>Progress</span>}
                  </td>
                  <td style={{ padding: "10px 12px", textAlign: "left" }}>
                    <span style={{ fontWeight: 700, color: answerColor, fontSize: 13 }}>
                      {(isTracker || isProgress) ? "recorded" : s.answer}
                      {isApprove && isNo && !isSkipped && " ⚠️"}
                    </span>
                  </td>
                  <td style={{ padding: "10px 12px", textAlign: "right", fontFamily: F.mono, fontSize: 13 }}>
                    {s.actualValue != null ? `${s.actualValue}${s.krUnit ? " " + s.krUnit : ""}` : "—"}
                  </td>
                  <td style={{ padding: "10px 12px", textAlign: "right", fontFamily: F.mono, fontSize: 12, color: T.textMuted, whiteSpace: "nowrap" }}>
                    {(!isTracker && !isProgress) && s.krTarget != null ? `${s.krOperator || ">="} ${s.krTarget}${s.krUnit ? " " + s.krUnit : ""}` : isProgress && s.krTarget != null ? `target: ${s.krTarget}${s.krUnit ? " " + s.krUnit : ""}` : "—"}
                  </td>
                  <td style={{ padding: "10px 12px", fontSize: 12, color: T.textMuted, whiteSpace: "nowrap" }}>
                    {s.dateRange || s.periodKey || "—"}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Footer */}
      <div style={{ padding: "12px 18px", borderTop: `1px solid ${T.border}`, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, flexWrap: "wrap", background: T.raised }}>
        <div style={{ fontSize: 12, color: T.textMuted }}>
          {toAct.length} of {submissions.length} selected{skipped.size > 0 ? ` · ${skipped.size} skipped` : ""}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button onClick={() => { setCancelled(true); onCancel(); }}
            style={{ padding: "7px 16px", borderRadius: 8, border: `1px solid ${T.border}`, background: "none", cursor: "pointer", fontSize: 13, color: T.textMuted, fontFamily: F.body }}>
            Cancel
          </button>
          <button
            disabled={toAct.length === 0}
            onClick={() => { const n = toAct.length; onConfirm(toAct.map(s => s.id)); setDoneCount(n); setDone(true); }}
            style={{ padding: "7px 20px", borderRadius: 8, border: "none", background: toAct.length > 0 ? accentColor : T.raised, cursor: toAct.length > 0 ? "pointer" : "default", fontSize: 13, fontWeight: 700, color: toAct.length > 0 ? "#fff" : T.textDim, fontFamily: F.body }}>
            {isApprove ? "✓ Approve" : "✕ Reject"} {toAct.length} submission{toAct.length !== 1 ? "s" : ""}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   ADMIN PORTAL
   ───────────────────────────────────────────────────────────── */
// ── Enrolment helpers ────────────────────────────────────────────────────────
function enrSortWeeksDesc(weeks) {
  return [...weeks].sort((a, b) => {
    const p = w => { const m = w.match(/(\d{4})[_\-]?W(\d+)/i); return m ? parseInt(m[1]) * 1000 + parseInt(m[2]) : 0; };
    return p(b) - p(a);
  });
}
// Week convention: Week 1 = first Monday of the calendar year; Mon–Fri only.
function weekToDateRange(weekKey, short = false) {
  const m = String(weekKey).match(/^(\d{4})-W(\d{1,2})$/i);
  if (!m) return weekKey;
  const year = parseInt(m[1]), week = parseInt(m[2]);
  const jan1Day = new Date(Date.UTC(year, 0, 1)).getUTCDay();
  const daysToMon = (8 - jan1Day) % 7;
  const firstMon = new Date(Date.UTC(year, 0, 1 + daysToMon));
  const mon = new Date(firstMon); mon.setUTCDate(firstMon.getUTCDate() + (week - 1) * 7);
  const fri = new Date(mon); fri.setUTCDate(mon.getUTCDate() + 4);
  const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  const fmt = d => `${d.getUTCDate()} ${MONTHS[d.getUTCMonth()]}`;
  if (short) return `${fmt(mon)} – ${fmt(fri)}`;
  return mon.getUTCFullYear() === fri.getUTCFullYear()
    ? `${fmt(mon)} – ${fmt(fri)} ${fri.getUTCFullYear()}`
    : `${fmt(mon)} ${mon.getUTCFullYear()} – ${fmt(fri)} ${fri.getUTCFullYear()}`;
}
function enrParseMarketerSheet(rows, fileName) {
  const result = { week: "", records: [], error: null, totalEnrolments: 0, marketers: [], rtos: [] };
  // Find week from "Week Number" row (search first 10 rows)
  for (let i = 0; i < Math.min(10, rows.length); i++) {
    const cell = String(rows[i]?.[0] ?? "").trim().toLowerCase();
    if (cell.includes("week number") || cell.includes("weeknumber")) {
      const wv = String(rows[i]?.[1] ?? "").trim();
      if (wv.startsWith("(")) { result.error = `Select a single week in the Excel slicer before uploading (current value: "${wv}").`; return result; }
      result.week = wv;
      break;
    }
  }
  // Fallback: try to extract week from filename
  if (!result.week) {
    const fm = String(fileName || "").match(/(\d{4})[_\-\s]?[Ww](\d{1,2})/);
    if (fm) result.week = `${fm[1]}-W${fm[2].padStart(2, "0")}`;
  }
  if (!result.week) { result.error = "Could not find 'Week Number' in the Marketer sheet. Check the file format."; return result; }
  // Find header row: col A = "Marketer" (case-insensitive)
  let hIdx = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i]?.[0] ?? "").trim().toLowerCase() === "marketer") { hIdx = i; break; }
  }
  if (hIdx < 0) { result.error = "Could not find 'Marketer' header row. Check the sheet format."; return result; }
  // Extract RTO columns — skip first (Marketer) and any "Grand Total" column
  const hRow = rows[hIdx] || [];
  const rtoCols = [];
  for (let c = 1; c < hRow.length; c++) {
    const h = String(hRow[c] ?? "").trim();
    if (h && h.toLowerCase() !== "grand total") rtoCols.push({ name: h, idx: c });
  }
  if (rtoCols.length === 0) { result.error = "No RTO columns found in the Marketer header row."; return result; }
  result.rtos = rtoCols.map(r => r.name);
  // Parse data rows
  for (let i = hIdx + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const m = String(row[0] ?? "").trim();
    if (!m || m.toLowerCase() === "grand total") continue;
    if (!result.marketers.includes(m)) result.marketers.push(m);
    for (const { name: rto, idx } of rtoCols) {
      const n = Number(row[idx]) || 0;
      if (n > 0) { result.records.push({ marketerName: m, rto, count: n }); result.totalEnrolments += n; }
    }
  }
  if (result.records.length === 0) result.error = "No enrolment data found. The sheet may be empty or fully filtered out.";
  return result;
}
async function coeParseSheetsFromFile(file) {
  const COE_TARGETS = [
    { names: ["1. NIET CoE"], rto: "NIET", type: "CoE" },
    { names: ["1. NIET Non-CoE", "1. NIET Non-COE"], rto: "NIET", type: "Non-CoE" },
    { names: ["2. CB CoE"], rto: "CB", type: "CoE" },
    { names: ["2. CB Non-CoE", "2. CB Non-COE"], rto: "CB", type: "Non-CoE" },
    { names: ["3. Rhodes Accepted & Paid", "3. Rhodes_Accepted & Paid"], rto: "Rhodes", type: "Accepted & Paid" },
  ];
  const result = { week: null, sheetsFound: [], sheetSummary: [], records: [], totalRecords: 0, fileName: file.name, fileSize: file.size, error: null };
  const { default: readXlsxFile } = await import("read-excel-file/browser");
  const formatDate = v => { if (!v) return ""; if (v instanceof Date) return v.toLocaleDateString("en-AU"); return String(v).trim(); };
  for (const target of COE_TARGETS) {
    let rows = null;
    for (const name of target.names) {
      try { const res = await readXlsxFile(file, { sheets: [name] }); rows = res[0]?.data || []; break; } catch {}
    }
    if (!rows) { result.sheetSummary.push({ rto: target.rto, type: target.type, found: false, count: 0 }); continue; }
    // Find week (search first 6 rows for "Week Number" label in col A)
    let week = null;
    for (let i = 0; i < Math.min(6, rows.length); i++) {
      if (String(rows[i]?.[0] ?? "").trim().toLowerCase().includes("week number")) {
        const wv = String(rows[i]?.[1] ?? "").trim();
        if (wv) { week = wv; break; }
      }
    }
    if (week && !result.week) result.week = week;
    // Find header row: col A = "Period"
    let hIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i]?.[0] ?? "").trim().toLowerCase() === "period") { hIdx = i; break; }
    }
    if (hIdx < 0) { result.sheetSummary.push({ rto: target.rto, type: target.type, found: true, count: 0, note: "No Period header" }); continue; }
    // Parse data rows with forward-fill on Period
    let lastPeriod = "";
    let count = 0;
    for (let i = hIdx + 1; i < rows.length; i++) {
      const row = rows[i] || [];
      const pCell = String(row[0] ?? "").trim();
      if (pCell.toLowerCase() === "grand total" || pCell.toLowerCase() === "total") break;
      if (pCell) lastPeriod = pCell;
      const studentId = String(row[1] ?? "").trim();
      if (!studentId) continue;
      result.records.push({ rto: target.rto, type: target.type, period: lastPeriod, studentId, courseName: String(row[2] ?? "").trim(), intakeDate: formatDate(row[3]), agent: String(row[4] ?? "").trim(), marketer: String(row[5] ?? "").trim(), date: formatDate(row[6]), createdBy: String(row[7] ?? "").trim(), onshoreOffshore: String(row[8] ?? "").trim(), pathway: String(row[9] ?? "").trim() });
      count++;
    }
    result.sheetsFound.push(`${target.rto} ${target.type}`);
    result.sheetSummary.push({ rto: target.rto, type: target.type, found: true, count });
    result.totalRecords += count;
  }
  if (!result.week) {
    const fm = String(file.name || "").match(/(\d{4})[_\-\s]?[Ww](\d{1,2})/);
    if (fm) result.week = `${fm[1]}-W${fm[2].padStart(2, "0")}`;
  }
  if (!result.week) result.error = "Could not determine week number. Check the file has a 'Week Number' cell.";
  else if (result.sheetsFound.length === 0) result.error = "No COE sheets found. Expected: '1. NIET CoE', '1. NIET Non-CoE', '2. CB CoE', '2. CB Non-CoE', '3. Rhodes Accepted & Paid'.";
  return result;
}
async function educareParseSheetsFromFile(file) {
  const EDUCARE_TARGETS = [
    { names: ["Educare BNE CoE"], rto: "Educare BNE", type: "CoE" },
    { names: ["Educare GC CoE"], rto: "Educare GC", type: "CoE" },
    { names: ["Educare ONLINE_Accepted & Paid", "Educare ONLINE Accepted & Paid"], rto: "Educare ONLINE", type: "Accepted & Paid" },
    { names: ["Educare GC_Non-CoE", "Educare GC_Non-COE", "Educare GC Non-CoE"], rto: "Educare GC", type: "Non-CoE" },
    { names: ["Educare BNE_Non-CoE", "Educare BNE_Non-COE", "Educare BNE Non-CoE"], rto: "Educare BNE", type: "Non-CoE" },
    { names: ["Educare Dom Accepted & Paid", "Educare DOM Accepted & Paid"], rto: "Educare Dom", type: "Accepted & Paid" },
  ];
  const result = { week: null, sheetsFound: [], sheetSummary: [], records: [], totalRecords: 0, fileName: file.name, fileSize: file.size, error: null };
  const { default: readXlsxFile } = await import("read-excel-file/browser");
  const formatDate = v => { if (!v) return ""; if (v instanceof Date) return v.toLocaleDateString("en-AU"); return String(v).trim(); };
  for (const target of EDUCARE_TARGETS) {
    let rows = null;
    for (const name of target.names) {
      try { const res = await readXlsxFile(file, { sheets: [name] }); rows = res[0]?.data || []; break; } catch {}
    }
    if (!rows) { result.sheetSummary.push({ rto: target.rto, type: target.type, found: false, count: 0 }); continue; }
    let week = null;
    for (let i = 0; i < Math.min(6, rows.length); i++) {
      if (String(rows[i]?.[0] ?? "").trim().toLowerCase().includes("week number")) {
        const wv = String(rows[i]?.[1] ?? "").trim();
        if (wv) { week = wv; break; }
      }
    }
    if (week && !result.week) result.week = week;
    let hIdx = -1;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i]?.[0] ?? "").trim().toLowerCase() === "period") { hIdx = i; break; }
    }
    if (hIdx < 0) { result.sheetSummary.push({ rto: target.rto, type: target.type, found: true, count: 0, note: "No Period header" }); continue; }
    let lastPeriod = "";
    let count = 0;
    for (let i = hIdx + 1; i < rows.length; i++) {
      const row = rows[i] || [];
      const pCell = String(row[0] ?? "").trim();
      if (pCell.toLowerCase() === "grand total" || pCell.toLowerCase() === "total") break;
      if (pCell) lastPeriod = pCell;
      const studentId = String(row[1] ?? "").trim();
      if (!studentId) continue;
      result.records.push({ rto: target.rto, type: target.type, period: lastPeriod, studentId, courseName: String(row[2] ?? "").trim(), intakeDate: formatDate(row[3]), agent: String(row[4] ?? "").trim(), marketer: String(row[5] ?? "").trim(), date: formatDate(row[6]), createdBy: String(row[7] ?? "").trim(), onshoreOffshore: String(row[8] ?? "").trim(), pathway: String(row[9] ?? "").trim() });
      count++;
    }
    result.sheetsFound.push(`${target.rto} ${target.type}`);
    result.sheetSummary.push({ rto: target.rto, type: target.type, found: true, count });
    result.totalRecords += count;
  }
  if (!result.week) {
    const fm = String(file.name || "").match(/(\d{4})[_\-\s]?[Ww](\d{1,2})/);
    if (fm) result.week = `${fm[1]}-W${fm[2].padStart(2, "0")}`;
  }
  if (!result.week) result.error = "Could not determine week number. Check the file has a 'Week Number' cell.";
  else if (result.sheetsFound.length === 0) result.error = "No Educare sheets found. Expected: 'Educare BNE CoE', 'Educare GC CoE', 'Educare ONLINE_Accepted & Paid', 'Educare GC_Non-CoE', 'Educare BNE_Non-CoE', 'Educare Dom Accepted & Paid'.";
  return result;
}
// ─────────────────────────────────────────────────────────────────────────────

function AdminPortal({ user, onLogout, state, dispatch, onImpersonate }) {
  const [page, setPageRaw] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'admin' ? (p[2] || 'ai-chat') : 'ai-chat';
  });
  const [selDept, setSelDept] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'admin' && p[2] === 'departments' ? (p[3] || null) : null;
  });
  const setPage = useCallback(p => { window.history.pushState(null, '', `/admin/${p}`); setPageRaw(p); }, []);
  const isMobile = useIsMobile();
  const [drawerOpen, setDrawerOpen] = useState(false);
  useEffect(() => {
    if (window.location.pathname.split('/')[1] !== 'admin') {
      window.history.replaceState(null, '', `/admin/ai-chat`);
    }
    const onPop = () => {
      const p = window.location.pathname.split('/');
      setPageRaw(p[1] === 'admin' ? (p[2] || 'ai-chat') : 'ai-chat');
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
  const [newKr, setNewKr] = useState({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false });
  const [addTarget, setAddTarget] = useState(null);
  const [showGenReport, setShowGenReport] = useState(false);
  const [genPeriod, setGenPeriod] = useState({ label: "", from: "", to: "" });
  const [projSearch, setProjSearch] = useState("");
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ name: "", status: "active", startDate: "", due: "", income: "", margin: "", contributeRate: "" });
  const [progressEdits, setProgressEdits] = useState({});
  const [logDrafts, setLogDrafts] = useState({});
  const [subFilter, setSubFilter] = useState("all");
  const [enrTab, setEnrTab] = useState("overview");
  const [enrRecords, setEnrRecords] = useState([]);
  const [enrBatches, setEnrBatches] = useState([]);
  const [enrLoaded, setEnrLoaded] = useState(false);
  const [enrLoading, setEnrLoading] = useState(false);
  const [enrFilterWeek, setEnrFilterWeek] = useState("all");
  const [enrFilterMarketer, setEnrFilterMarketer] = useState("all");
  const [enrFilterRto, setEnrFilterRto] = useState("all");
  const [enrParsed, setEnrParsed] = useState(null);
  const [enrImporting, setEnrImporting] = useState(false);
  const [enrError, setEnrError] = useState(null);
  const [enrChartTooltip, setEnrChartTooltip] = useState(null);
  useEffect(() => {
    if (!enrLoaded && !enrLoading) {
      setEnrLoading(true);
      dbGetEnrolments().then(r => {
        setEnrRecords(r.enrolment_records);
        setEnrBatches(r.enrolment_batches);
        setEnrError(null);
        setEnrLoaded(true);
        setEnrLoading(false);
      }).catch(e => { setEnrError(e.message); setEnrLoading(false); setEnrLoaded(true); });
    }
  }, [enrLoaded, enrLoading]);
  const [coeTab, setCoeTab] = useState("overview");
  const [coeRecords, setCoeRecords] = useState([]);
  const [coeBatches, setCoeBatches] = useState([]);
  const [coeLoaded, setCoeLoaded] = useState(false);
  const [coeLoading, setCoeLoading] = useState(false);
  const [coeFilterWeek, setCoeFilterWeek] = useState("all");
  const [coeFilterRto, setCoeFilterRto] = useState("all");
  const [coeFilterType, setCoeFilterType] = useState("all");
  const [coeParsed, setCoeParsed] = useState(null);
  const [coeImporting, setCoeImporting] = useState(false);
  const [coeError, setCoeError] = useState(null);
  const [educareParsed, setEducareParsed] = useState(null);
  const [educareImporting, setEducareImporting] = useState(false);
  useEffect(() => {
    if (!coeLoaded && !coeLoading) {
      setCoeLoading(true);
      dbGetCoeData().then(r => {
        setCoeRecords(r.coe_records);
        setCoeBatches(r.coe_batches);
        setCoeError(null);
        setCoeLoaded(true);
        setCoeLoading(false);
      }).catch(e => { setCoeError(e.message); setCoeLoading(false); setCoeLoaded(true); });
    }
  }, [coeLoaded, coeLoading]);
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
  const [lbEditKr, setLbEditKr] = useState(null);
  const [lbAddMember, setLbAddMember] = useState(null);
  const [lbKrForm, setLbKrForm] = useState({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false });
  const [confirmDeleteKr, setConfirmDeleteKr] = useState(null);
  const [syncNote, setSyncNote] = useState(null);
  const syncNoteTimer = useRef(null);
  const [subSearch, setSubSearch] = useState("");
  const [subDeptFilter, setSubDeptFilter] = useState("all");
  const [expandedMonthlyKr, setExpandedMonthlyKr] = useState(null);
  const [expandedPersonalMember, setExpandedPersonalMember] = useState(null);
  const [addPersonalKr, setAddPersonalKr] = useState(null);
  const [subPeriod, setSubPeriod] = useState("monthly");
  const [checkinPeriods, setCheckinPeriods] = useState(["monthly"]);
  const [sendingCheckin, setSendingCheckin] = useState(false);
  const [checkinResult, setCheckinResult] = useState(null);
  const [checkinScope, setCheckinScope] = useState({ deptId: "", teamId: "", userId: "" });
  const [checkinPreview, setCheckinPreview] = useState(null);
  const [showCheckinDialog, setShowCheckinDialog] = useState(false);
  const [rejectOkr, setRejectOkr] = useState(null);
  const [editingSub, setEditingSub] = useState(null);
  const [editingApproved, setEditingApproved] = useState(null);
  const [tmplPeriod, setTmplPeriod] = useState("default");
  const [testEmailState, setTestEmailState] = useState({ status: "idle", msg: "" });
  const [testEmailTo, setTestEmailTo] = useState(user?.email || "");
  const [adminOkrPeriod, setAdminOkrPeriod] = useState("all");
  const [expandedLog, setExpandedLog] = useState(null);
  const [resendingEmail, setResendingEmail] = useState(null); // `${logId}:${email}`
  const [chatHistory, setChatHistory] = useState([]);
  const [chatInput, setChatInput] = useState("");
  const [chatLoading, setChatLoading] = useState(false);
  const [chatPromptOpen, setChatPromptOpen] = useState(false);
  const chatEndRef = useRef(null);
  const chatHistoryRef = useRef([]);
  const pilotPrevPageRef = useRef(null);
  const [pilotSessions, setPilotSessions] = useState(() => {
    try { return JSON.parse(localStorage.getItem("niet_pilot_sessions") || "[]"); } catch { return []; }
  });

  const DEFAULT_CHAT_PROMPT = `You are NIET Pilot, the OKR analytics assistant for NIET (National Institute for Excellence in Teaching) admin portal.

Your role is to help school administrators quickly understand staff OKR performance, identify who needs follow-up, and prepare summary insights for meetings.

DATA RULES:
- OKR data covers the current calendar month's check-in submissions only.
- Completion rates compare actual values against targets using operator logic (not simple yes/no counts).
- Status thresholds: green ≥ 66.7%, yellow ≥ 60%, red < 60%.
- A member shown as "no data this month" has not yet received or answered a check-in — do not treat them as 0% performers.
- Tracker KRs record numerical values only and do not affect completion rates.
- Yes/no KRs may include a reported actual value (e.g. total sales revenue). When a yes/no KR actual value and a tracker KR cover a similar metric, the yes/no KR actual is the confirmed reported figure and takes precedence over the tracker value.
- Progress KRs record a cumulative running total toward a target and contribute proportionally to completion rate (actual ÷ target × 100%). They are used for project-oriented goals like annual profit targets where multiple check-ins occur through the year.
- Project Profit KRs are annual KRs assigned to managers. Their actual value is auto-calculated from the sum of (income × margin %) of projects the manager owns that were completed in the KR's target year. No check-ins are generated for this type. They do not contribute to the monthly completion rate shown in OKR stats. Treat them as a separate, always-current progress indicator for the manager's revenue contribution.
- Financial data (Income, Net Profit, Expenses) is cumulative from July to the current FY month, broken down by division (NIET, CB, Rhodes, Educare). PT = Performance Target, DT = Dream Target.
- Project data includes all manager-owned projects: name, department, responsible manager, status (active/pending approval/completed), progress (0–100%), start date, due date, optional income ($), optional profit margin (%), and computed profit ($). Proactively flag projects that are overdue or have low progress close to their due date.

RESPONSE STYLE:
- Lead with the direct answer, then supporting detail.
- Use a markdown table when comparing 3 or more items across the same dimensions (e.g., dept rankings).
- Use a short bullet list for enumerations (e.g., list of members needing follow-up).
- For a single fact or number, one sentence is enough.
- Always name specific people and departments — never give vague summaries.
- Proactively flag actionable items: red-status members, pending unanswered check-ins, departments below target.
- Do not repeat the question. Do not explain the data structure unless asked.
- Reply in the same language as the user's question (Chinese or English).
- Never fabricate data. If the information is not in the provided context, say so clearly.

PREDICTIVE ANALYSIS:
You can answer forward-looking and predictive questions (e.g. "how will the company perform next month?", "which department is at risk of missing targets?"). When doing so:
- Base predictions only on trends visible in the provided data: month-by-month financial figures, tracker KR trajectories (use the period keys to determine direction), weekly enrolment trends, current OKR completion rates, and project progress vs. due dates.
- Always state the specific data you are reasoning from (e.g. "Based on the past 3 months of net profit figures…").
- Express uncertainty clearly — use language like "on current trajectory", "if the trend continues", "this is an estimate based on available data".
- Flag the key risks or assumptions that could change the outcome.
- Never invent trend data that is not present in the context. If insufficient historical data exists for a reliable prediction, say so and explain what additional data would help.

MARKETING & SALES ANALYSIS:
When answering questions about marketing performance, sales, or student acquisition (e.g. "who is the top marketer?", "how are we tracking on enrolments?", "what does the pipeline look like?"):
- Consider both Applications data (WEEKLY ENROLMENTS, by marketer and RTO) and COE data (WEEKLY COE RECORDS, by marketer and RTO) — they cover different stages of the student pipeline.
- Applications = initial enrolment stage. COE (Confirmation of Enrolment) = a later confirmation stage. Both datasets track marketer and RTO.
- Proactively cross-reference both datasets when answering marketer performance questions — a marketer's Applications count and their COE count together give a fuller picture of their pipeline.
- COE data covers two separate report sources: NIET/CB/Rhodes and Educare. When answering across all RTOs, include both.
- If only one dataset has been imported, answer from what's available and note the other is not yet loaded.

ACTIONS:
When the user asks to approve or reject OKR submissions (e.g. "approve all pending IT submissions", "reject Sarah's check-in"), call the propose_bulk_action tool with appropriate filter criteria. Never describe or confirm the action in text — always use the tool. The frontend will show the admin a full submission review card with all details before any action is executed.`;

  function archivePilotSession(history) {
    const msgs = history.filter(m => m.role === "user" || m.role === "ai");
    if (msgs.length === 0) return;
    const firstUser = msgs.find(m => m.role === "user");
    if (!firstUser) return;
    const session = {
      id: Date.now(),
      date: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }),
      preview: firstUser.text.slice(0, 80),
      messages: msgs,
    };
    setPilotSessions(prev => {
      if (prev[0] && prev[0].preview === session.preview && prev[0].messages.length === msgs.length) return prev;
      const next = [session, ...prev].slice(0, 5);
      localStorage.setItem("niet_pilot_sessions", JSON.stringify(next));
      return next;
    });
  }

  function buildChatContext() {
    const { depts, memberData, okrSubmissions = [], monthlyReports = [], users, projects = [] } = state;
    const now = new Date();
    const getCompletedYear = p => {
      if (p.completedYear) return p.completedYear;
      const parts = (p.updatedDate || "").split("/");
      return parts.length >= 3 ? parseInt(parts[2].split(",")[0].trim()) : null;
    };
    const today = now.toISOString().slice(0, 10);
    const monthLabel = now.toLocaleString("en-AU", { month: "long", year: "numeric" });
    // Match Company Overview monthly view: submissions sent in current calendar month, period in daily/weekly/monthly
    const monthlyTypes = ["daily", "weekly", "monthly"];
    const monthSubs = okrSubmissions.filter(s => {
      if (!monthlyTypes.includes(s.period) || !s.sentAt) return false;
      const d = new Date(s.sentAt);
      return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
    });
    const nowMs = Date.now();
    const members = users.filter(u => u.role === "member" || u.role === "manager");

    // Per-member stats using calcMemberRate — identical to Company Overview deptRanks logic
    const memberStats = members.map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const krs = kd.krs.filter(kr => monthlyTypes.includes(kr.period || "monthly"));
      const dept = depts.find(d => d.id === u.deptId)?.name || "—";
      // Eligible = has at least one answered or overdue-unanswered submission this month (non-tracker only)
      const hasEligible = krs.filter(kr => kr.type !== "tracker").some(kr => {
        const krSubs = monthSubs.filter(s => s.memberId === u.id && s.krId === kr.id);
        if (!krSubs.length) return false;
        if (krSubs.some(s => s.answer !== null)) return true;
        const latest = krSubs.filter(s => s.sentAt).sort((a, b) => b.sentAt.localeCompare(a.sentAt))[0];
        return latest && (nowMs - new Date(latest.sentAt).getTime()) >= 86400000;
      });
      const rate = hasEligible ? calcMemberRate(u.id, krs, monthSubs) : null;
      const status = rate !== null ? getStatus(rate) : "no data";
      const answered = monthSubs.filter(s => s.memberId === u.id && s.answer !== null).length;
      const pending = monthSubs.filter(s => s.memberId === u.id && s.answer === null).length;
      const trackerLines = kd.krs.filter(kr => kr.type === "tracker").map(kr => {
        const krSubs = okrSubmissions.filter(s => s.memberId === u.id && s.krId === kr.id)
          .sort((a, b) => (b.sentAt || "").localeCompare(a.sentAt || "")).slice(0, 3);
        if (!krSubs.length) return null;
        const entries = krSubs.map(s => `${s.actualValue ?? 0}${kr.unit ? " " + kr.unit : ""} (${s.dateRange || s.periodKey})`).join(", ");
        return `    Tracker — ${kr.label}: ${entries}`;
      }).filter(Boolean);
      const krLines = kd.krs.filter(kr => kr.type !== "tracker" && kr.type !== "manager-fill" && kr.type !== "project_profit").map(kr => {
        const latest = okrSubmissions.filter(s => s.memberId === u.id && s.krId === kr.id && s.answer !== null)
          .sort((a, b) => (b.answeredAt || b.sentAt || "").localeCompare(a.answeredAt || a.sentAt || ""))[0];
        if (!latest || latest.actualValue == null) return null;
        return `    KR — ${kr.label}: ${latest.answer === "yes" ? "✓ Yes" : "✗ No"} · actual ${latest.actualValue}${kr.unit ? " " + kr.unit : ""} (${latest.dateRange || latest.periodKey})`;
      }).filter(Boolean);
      const ppLines = kd.krs.filter(kr => kr.type === "project_profit").map(kr => {
        const yearProjects = projects.filter(p => p.mgrId === u.id && p.status === "completed" && getCompletedYear(p) === kr.krYear);
        const actual = yearProjects.reduce((s, p) => s + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0);
        const pct = kr.target > 0 ? Math.min(Math.round(actual / kr.target * 100), 100) : 0;
        const missing = yearProjects.filter(p => p.income == null || p.margin == null).length;
        const hasPartialRate = yearProjects.some(p => p.contributeRate != null && p.contributeRate < 100);
        return `    Proj Profit KR — ${kr.label}: $${actual.toLocaleString()} of $${(kr.target || 0).toLocaleString()} (${pct}%)${missing > 0 ? ` [⚠ ${missing} project(s) missing income/margin]` : ""}${hasPartialRate ? " [⚡ partial rates applied]" : ""} · Year ${kr.krYear || "?"}`;
      });
      return { id: u.id, name: u.name, dept, role: u.role, rate, status, answered, pending, hasEligible, trackerLines, krLines, ppLines };
    });

    // Dept rates = average of eligible member rates (same as Company Overview deptRanks)
    const deptStats = depts.map(d => {
      const dm = memberStats.filter(m => members.find(u => u.id === m.id)?.deptId === d.id);
      const eligible = dm.filter(m => m.hasEligible);
      const rate = eligible.length ? eligible.reduce((s, m) => s + m.rate, 0) / eligible.length : null;
      return { name: d.name, rate, status: rate !== null ? getStatus(rate) : "no data", total: dm.length, eligible: eligible.length };
    }).sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1));

    // Company rate = average of depts with data (same as Company Overview compRate)
    const deptsWithData = deptStats.filter(d => d.rate !== null);
    const compRate = deptsWithData.length ? deptsWithData.reduce((s, d) => s + d.rate, 0) / deptsWithData.length : null;

    const deptSection = deptStats.map(d =>
      `  ${d.name}: ${d.rate !== null ? d.rate.toFixed(1) + "% [" + d.status + "]" : "no data"} (${d.eligible}/${d.total} members with submissions)`
    ).join("\n");

    const memberSection = memberStats.map(m => {
      let line = `  ${m.name} (${m.dept}, ${m.role}): ${m.rate !== null ? m.rate.toFixed(1) + "% [" + m.status + "]" : "no data this month"}`;
      if (m.answered || m.pending) line += ` — ${m.answered} answered, ${m.pending} awaiting`;
      else line += ` — no check-ins sent this month`;
      if (m.trackerLines.length) line += "\n" + m.trackerLines.join("\n");
      if (m.krLines.length) line += "\n" + m.krLines.join("\n");
      if (m.ppLines && m.ppLines.length) line += "\n" + m.ppLines.join("\n");
      return line;
    }).join("\n");

    const latestReport = [...monthlyReports].sort((a, b) => (b.publishedDate || "").localeCompare(a.publishedDate || ""))[0];
    const reportSection = latestReport
      ? [`LATEST PUBLISHED REPORT (${latestReport.month || latestReport.publishedDate}): company rate ${Number(latestReport.data?.companyRate).toFixed(1)}%`,
         ...(latestReport.data?.deptRanks?.length ? ["  Dept breakdown: " + latestReport.data.deptRanks.map(d => `${d.name} ${Number(d.rate).toFixed(1)}%`).join(", ")] : [])
        ].join("\n")
      : "";

    // Financial Performance — same data as Financial Performance page
    const FY_MONTHS = ["Jul","Aug","Sep","Oct","Nov","Dec","Jan","Feb","Mar","Apr","May","Jun"];
    const REV_DIVS = ["NIET","CB","Rhodes","Educare"];
    const nowFYMonth = (() => { const m = now.getMonth(); return m >= 6 ? m - 6 : m + 6; })();
    const fyLabel = `Jul–${FY_MONTHS[nowFYMonth]} FY${String(now.getMonth() >= 6 ? now.getFullYear() + 1 : now.getFullYear()).slice(2)}`;
    const fmtMoney = v => v >= 1_000_000 ? `$${(v/1_000_000).toFixed(2)}M` : v >= 1_000 ? `$${(v/1_000).toFixed(1)}K` : `$${Math.round(v)}`;
    const mkDefault = () => ({ pt: 0, dt: 0, divisions: Object.fromEntries(REV_DIVS.map(d => [d, Array(12).fill(0)])) });
    const finModules = [
      { key: "revenue",   label: "Income",     hasTargets: true  },
      { key: "expense",   label: "Expenses",   hasTargets: false },
      { key: "netProfit", label: "Net Profit", hasTargets: true  },
    ];
    const _aiRevCfg = state.settings?.revenue ?? mkDefault();
    const _aiExpCfg = state.settings?.expense ?? mkDefault();
    const _aiNpDivs = Object.fromEntries(REV_DIVS.map(d => [d, Array(12).fill(0).map((_, i) => (_aiRevCfg.divisions[d]?.[i] || 0) - (_aiExpCfg.divisions[d]?.[i] || 0))]));
    const finSection = finModules.map(({ key, label, hasTargets }) => {
      const cfg = state.settings?.[key] ?? mkDefault();
      const divs = key === "netProfit" ? _aiNpDivs : cfg.divisions;
      const monthly = FY_MONTHS.map((_, i) => REV_DIVS.reduce((s, d) => s + (divs[d]?.[i] || 0), 0));
      const cumulative = monthly.map((_, i) => monthly.slice(0, i + 1).reduce((a, b) => a + b, 0));
      const cum = cumulative[nowFYMonth] || 0;
      const ptPct = hasTargets && cfg.pt > 0 ? (cum / cfg.pt * 100).toFixed(1) + "%" : null;
      const dtPct = hasTargets && cfg.dt > 0 ? (cum / cfg.dt * 100).toFixed(1) + "%" : null;
      const divLines = REV_DIVS.map(div => {
        const divCum = (divs[div] || Array(12).fill(0)).slice(0, nowFYMonth + 1).reduce((a, b) => a + b, 0);
        const pct = cum > 0 ? (divCum / cum * 100).toFixed(1) + "%" : "0%";
        return `    ${div}: ${fmtMoney(divCum)} (${pct} of group)`;
      });
      const monthlyLine = FY_MONTHS.slice(0, nowFYMonth + 1).map((m, i) => `${m} ${fmtMoney(monthly[i])}`).join(", ");
      let lines = [`  ${label}: cumulative ${fmtMoney(cum)} (${fyLabel})`];
      if (hasTargets) {
        if (cfg.pt > 0) lines.push(`    vs PT (${fmtMoney(cfg.pt)}): ${ptPct}`);
        if (cfg.dt > 0) lines.push(`    vs DT (${fmtMoney(cfg.dt)}): ${dtPct}`);
      }
      if (cum > 0) {
        lines.push(`    Division breakdown:\n${divLines.join("\n")}`);
        lines.push(`    Monthly: ${monthlyLine}`);
      } else {
        lines.push(`    No data entered yet`);
      }
      return lines.join("\n");
    }).join("\n");

    // Pending submissions awaiting admin approval — grouped by dept for AI awareness
    const allPending = okrSubmissions.filter(s => s.answer !== null && s.approval === "pending");
    let pendingSection;
    if (allPending.length === 0) {
      pendingSection = "PENDING SUBMISSIONS: None currently awaiting approval.";
    } else {
      const byDept = depts.map(d => {
        const dp = allPending.filter(s => s.deptId === d.id);
        if (!dp.length) return null;
        const byMonth = {};
        dp.forEach(s => { const mk = (s.sentAt || "").slice(0, 7) || "unknown"; byMonth[mk] = (byMonth[mk] || 0) + 1; });
        const breakdown = Object.entries(byMonth).sort().map(([k, n]) => `${n} in ${k}`).join(", ");
        return `  ${d.name}: ${dp.length} pending (${breakdown})`;
      }).filter(Boolean);
      pendingSection = `PENDING SUBMISSIONS (answered by member, awaiting admin approval):\n${byDept.join("\n")}\n  Total: ${allPending.length}`;
    }

    // Projects — grouped by status, dept derived from mgrId
    const activeProjects = projects.filter(p => p.status === "active");
    const pendingApprovalProjects = projects.filter(p => p.status === "pending approval");
    const completedProjects = projects.filter(p => p.status === "completed");
    const fmtProject = p => {
      const mgr = users.find(u => u.id === p.mgrId);
      const dept = mgr ? (depts.find(d => d.id === mgr.deptId)?.name || "—") : "—";
      const mgrName = mgr?.name || "—";
      const overdue = p.due && p.due !== "TBD" && new Date(p.due) < now && p.status === "active" ? " [OVERDUE]" : "";
      let line = `  "${p.name}" | ${dept} | Mgr: ${mgrName} | Progress: ${p.progress}%${p.startDate ? ` | Start: ${p.startDate}` : ""} | Due: ${p.due || "TBD"}${overdue}`;
      if (p.income != null) line += ` | Income: ${fmtMoney(p.income)}`;
      if (p.income != null && p.margin != null) line += ` | Profit: ${fmtMoney(Math.round(p.income * p.margin / 100))} (Margin: ${p.margin}%)${p.contributeRate != null && p.contributeRate < 100 ? ` [Owner's KR share: ${p.contributeRate}%]` : ""}`;
      const logEntries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []);
      if (logEntries.length) line += `\n    Latest Log: ${logEntries[0].text.slice(0, 100)}${logEntries[0].text.length > 100 ? "…" : ""}`;
      return line;
    };
    const projectSection = [
      `PROJECTS: Total ${projects.length} (${activeProjects.length} active, ${pendingApprovalProjects.length} pending approval, ${completedProjects.length} completed)`,
      activeProjects.length ? `Active (${activeProjects.length}):\n${activeProjects.map(fmtProject).join("\n")}` : "Active: none",
      pendingApprovalProjects.length ? `Pending Approval (${pendingApprovalProjects.length}) — awaiting System Admin sign-off:\n${pendingApprovalProjects.map(fmtProject).join("\n")}` : "",
      completedProjects.length ? `Completed (${completedProjects.length}):\n${completedProjects.map(fmtProject).join("\n")}` : "",
    ].filter(Boolean).join("\n");

    // Weekly Enrolment data
    let enrolmentSection;
    if (!enrLoaded) {
      enrolmentSection = "WEEKLY ENROLMENTS: Data not yet loaded this session.";
    } else if (enrError && enrRecords.length === 0) {
      enrolmentSection = `WEEKLY ENROLMENTS: Failed to load — ${enrError}`;
    } else if (enrRecords.length === 0) {
      enrolmentSection = "WEEKLY ENROLMENTS: No enrolment data imported yet.";
    } else {
      const enrWeeks = enrSortWeeksDesc([...new Set(enrRecords.map(r => r.week))]);
      const enrMarketers = [...new Set(enrRecords.map(r => r.marketerName))].sort();
      const enrRtos = [...new Set(enrRecords.map(r => r.rto))].sort();
      const weekLines = enrWeeks.map(w => {
        const wRecs = enrRecords.filter(r => r.week === w);
        const wTotal = wRecs.reduce((s, r) => s + r.count, 0);
        const mLines = enrMarketers.map(m => {
          const mRecs = wRecs.filter(r => r.marketerName === m);
          if (!mRecs.length) return null;
          const mTotal = mRecs.reduce((s, r) => s + r.count, 0);
          const rtoParts = mRecs.map(r => `${r.rto}:${r.count}`).join(", ");
          return `    ${m}: ${rtoParts} (total:${mTotal})`;
        }).filter(Boolean);
        return `  ${w} — ${wTotal} enrolments:\n${mLines.join("\n")}`;
      });
      enrolmentSection = [
        `WEEKLY ENROLMENTS: ${enrWeeks.length} weeks | Marketers: ${enrMarketers.join(", ")} | RTOs: ${enrRtos.join(", ")}`,
        weekLines.join("\n"),
      ].join("\n");
    }

    // COE records section
    let coeSection;
    if (!coeLoaded) {
      coeSection = "COE RECORDS: Data not yet loaded this session.";
    } else if (coeError && coeRecords.length === 0) {
      coeSection = `COE RECORDS: Failed to load — ${coeError}`;
    } else if (coeRecords.length === 0) {
      coeSection = "COE RECORDS: No COE data imported yet.";
    } else {
      const NIET_CB_COMBOS = [{ rto: "NIET", type: "CoE" }, { rto: "NIET", type: "Non-CoE" }, { rto: "CB", type: "CoE" }, { rto: "CB", type: "Non-CoE" }, { rto: "Rhodes", type: "Accepted & Paid" }];
      const EDUCARE_COMBOS = [{ rto: "Educare BNE", type: "CoE" }, { rto: "Educare GC", type: "CoE" }, { rto: "Educare ONLINE", type: "Accepted & Paid" }, { rto: "Educare GC", type: "Non-CoE" }, { rto: "Educare BNE", type: "Non-CoE" }, { rto: "Educare Dom", type: "Accepted & Paid" }];
      const coeWeeks = enrSortWeeksDesc([...new Set(coeRecords.map(r => r.week))]);
      const coeByWeek = coeWeeks.map(w => {
        const wRecs = coeRecords.filter(r => r.week === w);
        const mkLines = (combos, prefix) => combos.map(c => { const n = wRecs.filter(r => r.rto === c.rto && r.type === c.type).length; return n > 0 ? `    [${prefix}] ${c.rto} ${c.type}: ${n}` : null; }).filter(Boolean);
        const lines = [...mkLines(NIET_CB_COMBOS, "NIET/CB/Rhodes"), ...mkLines(EDUCARE_COMBOS, "Educare")];
        return `  ${w} — ${wRecs.length} total:\n${lines.join("\n")}`;
      });
      coeSection = [`COE RECORDS: ${coeWeeks.length} week(s) | Total: ${coeRecords.length} records (NIET/CB/Rhodes + Educare combined)`, coeByWeek.join("\n")].join("\n");
    }

    return [
      `[Today: ${today} | Month: ${monthLabel} | Company OKR completion: ${compRate !== null ? compRate.toFixed(1) + "%" : "no data"} | Target: ${TP}%]`,
      `\nDEPARTMENT COMPLETION (current month, same logic as Company Overview):\n${deptSection}`,
      `\nMEMBER DETAILS (current month):\n${memberSection}`,
      reportSection ? `\n${reportSection}` : "",
      `\nFINANCIAL PERFORMANCE (${fyLabel}):\n${finSection}`,
      `\n${pendingSection}`,
      `\n${projectSection}`,
      `\n${enrolmentSection}`,
      `\n${coeSection}`,
    ].join("\n");
  }

  function applyActionFilters(filters) {
    // Start from all pending: member has answered, admin hasn't reviewed yet
    let subs = okrSubmissions.filter(s => s.answer !== null && s.approval === "pending");
    if (filters?.deptName) {
      const lc = filters.deptName.toLowerCase();
      const matched = depts.find(d => d.name.toLowerCase().includes(lc));
      if (matched) subs = subs.filter(s => s.deptId === matched.id);
    }
    if (filters?.memberName) {
      const lc = filters.memberName.toLowerCase();
      subs = subs.filter(s => (s.memberName || "").toLowerCase().includes(lc));
    }
    if (!filters?.allPeriods) {
      if (filters?.periodKey) {
        subs = subs.filter(s => (s.periodKey || "").startsWith(filters.periodKey));
      } else {
        // Default: submissions sent in current calendar month
        const nowD = new Date();
        subs = subs.filter(s => {
          if (!s.sentAt) return false;
          const d = new Date(s.sentAt);
          return d.getFullYear() === nowD.getFullYear() && d.getMonth() === nowD.getMonth();
        });
      }
    }
    // Enrich with deptName for display
    return subs.map(s => ({ ...s, deptName: depts.find(d => d.id === s.deptId)?.name || "—" }));
  }

  async function sendChat(question) {
    const q = (question || chatInput).trim();
    if (!q || chatLoading) return;
    setChatInput("");
    setChatHistory(h => [...h, { role: "user", text: q }]);
    setChatLoading(true);
    const customPrompt = settings?.aiChatPrompt || DEFAULT_CHAT_PROMPT;
    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ question: q, systemPrompt: customPrompt, contextData: buildChatContext(), lang: "en" }),
      });
      const text = await res.text();
      let data;
      try { data = JSON.parse(text); } catch { throw new Error(text.slice(0, 200)); }
      if (data.action) {
        // Tool use response — apply filters in frontend, show review card
        const pendingSubs = applyActionFilters(data.action.filters);
        setChatHistory(h => [...h, { role: "action", action: data.action, submissions: pendingSubs }]);
      } else {
        setChatHistory(h => [...h, { role: "ai", text: data.answer || data.error || "No response." }]);
      }
    } catch (err) {
      setChatHistory(h => [...h, { role: "ai", text: `Error: ${err.message}` }]);
    }
    setChatLoading(false);
  }

  useEffect(() => { chatEndRef.current?.scrollIntoView({ behavior: "smooth" }); }, [chatHistory, chatLoading]);

  // Keep ref in sync so page-change effect can read latest chatHistory without stale closure
  useEffect(() => { chatHistoryRef.current = chatHistory; }, [chatHistory]);

  // Auto-save draft on every message — survives browser refresh
  useEffect(() => {
    const msgs = chatHistory.filter(m => m.role === "user" || m.role === "ai");
    if (msgs.length > 0) localStorage.setItem("niet_pilot_draft", JSON.stringify(msgs));
    else localStorage.removeItem("niet_pilot_draft");
  }, [chatHistory]);

  // Restore draft on mount (recover from refresh)
  useEffect(() => {
    try {
      const draft = JSON.parse(localStorage.getItem("niet_pilot_draft") || "null");
      if (Array.isArray(draft) && draft.length > 0) setChatHistory(draft);
    } catch {}
  }, []);

  // Archive session when navigating away from Pilot page
  useEffect(() => {
    if (pilotPrevPageRef.current === "ai-chat" && page !== "ai-chat") {
      archivePilotSession(chatHistoryRef.current);
    }
    pilotPrevPageRef.current = page;
  }, [page]);

  async function resendEmail(log, recipient) {
    const key = `${log.id}:${recipient.email}`;
    setResendingEmail(key);
    const member = users.find(u => u.email === recipient.email);
    const emailTemplates = settings?.emailTemplates || {};
    try {
      let payload;
      const periods = log.periods || [log.period];
      const template = periods.length === 1
        ? { ...emailTemplates.default, ...(emailTemplates[periods[0]] || {}) }
        : { ...emailTemplates.default };
      if (periods.length > 1) {
        const sections = periods.map(p => {
          const pk = log.periodKeys?.[p] || log.periodKey;
          const dr = log.dateRanges?.[p] || log.dateRange;
          const subs = (okrSubmissions || []).filter(s => s.memberId === member?.id && s.period === p && s.periodKey === pk);
          const krs = subs.map(s => ({ id: s.krId, label: s.krLabel, target: s.krTarget, unit: s.krUnit, type: s.krType, operator: s.krOperator, isMonthly: s.krIsMonthly }));
          return { period: p, periodKey: pk, dateRange: dr, krs };
        }).filter(sec => sec.krs.length > 0);
        payload = { to: recipient.email, name: recipient.name, sections, template };
      } else {
        const p = periods[0];
        const subs = (okrSubmissions || []).filter(s => s.memberId === member?.id && s.period === p && s.periodKey === log.periodKey);
        const krs = subs.map(s => ({ id: s.krId, label: s.krLabel, target: s.krTarget, unit: s.krUnit, type: s.krType, operator: s.krOperator, isMonthly: s.krIsMonthly }));
        payload = { to: recipient.email, name: recipient.name, period: p, periodKey: log.periodKey, dateRange: log.dateRange, krs, template };
      }
      const res = await fetch("/api/send-email", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload) });
      const data = await res.json();
      dispatch({ type: "UPDATE_EMAIL_LOG_RECIPIENT", logId: log.id, email: recipient.email, success: data.ok, reason: data.ok ? null : (data.error || `HTTP ${res.status}`) });
    } catch (err) {
      dispatch({ type: "UPDATE_EMAIL_LOG_RECIPIENT", logId: log.id, email: recipient.email, success: false, reason: err.message || "Network error" });
    }
    setResendingEmail(null);
  }
  const [adminSelDept, setAdminSelDept] = useState(null);
  const [ovExpandedDept, setOvExpandedDept] = useState(null);
  const [logPopup, setLogPopup] = useState(null);
  useEffect(() => { if (!logPopup) return; const h = e => { if (e.key === "Escape") setLogPopup(null); }; window.addEventListener("keydown", h); return () => window.removeEventListener("keydown", h); }, [logPopup]);

  const { depts, memberData, mgrSprints, monthlyReports, projects, weeklySubs, okrSubmissions = [], emailLogs = [], users, settings } = state;
  const colOrder = settings?.colOrder || ["id", "label", "operator", "period", "target", "actual", "unit", "dataSource"];
  const navItems = [
    { id: "ai-chat",          icon: "⬡", label: "NIET Pilot"        },
    { id: "overview",         icon: "⬡", label: "Company Overview"  },
    { id: "okr-mgmt",         icon: "⬡", label: "OKR Management"    },
    { id: "submissions",      icon: "⬡", label: "OKR Submissions"   },
    { id: "reports",          icon: "⬡", label: "OKR Reports"       },
    { id: "projects",         icon: "⬡", label: "Projects"          },
    { id: "admissions",       icon: "⬡", label: "Applications"       },
    { id: "coe",              icon: "⬡", label: "COE"                },
    { id: "leaderboard",      icon: "⬡", label: "Leaderboard"       },
    { id: "users",            icon: "⬡", label: "User Management"   },
    { id: "email-templates",  icon: "⬡", label: "Email Templates"   },
  ];
  const deptSubItems = [
    { id: "__all__",   label: "All Departments", icon: "⬡" },
    { id: "__setup__", label: "Set Up OKRs",      icon: "⬡" },
  ];

  const OV_TYPES = { weekly: ["daily","weekly"], monthly: ["daily","weekly","monthly"], annual: ["daily","weekly","monthly","quarterly","biannual","annual"] };
  const ovTypes = OV_TYPES[overviewView] || OV_TYPES.monthly;
  const ovSubs = okrSubmissions.filter(s => {
    if (!ovTypes.includes(s.period) || !s.sentAt) return false;
    const d = new Date(s.sentAt), now = new Date();
    if (overviewView === "weekly") {
      const dow = now.getDay();
      const mon = new Date(now); mon.setHours(0,0,0,0); mon.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1) - 7);
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
    const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id && !u.excludeFromRate);
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
    const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id && !u.excludeFromRate);
    const rates = members.map(u => {
      const kd = memberData[u.id] || { krs: [] };
      if (!kd.krs.some(kr => rptSubs.some(s => s.memberId === u.id && s.krId === kr.id))) return null;
      return calcMemberRate(u.id, kd.krs, rptSubs);
    }).filter(r => r !== null);
    const hasData = rates.length > 0;
    const rate = hasData ? Math.round(rates.reduce((a, b) => a + b, 0) / rates.length * 10) / 10 : 0;
    return { name: d.name, rate, hasData, status: getStatus(rate) };
  }).sort((a, b) => b.rate - a.rate);
  const rptActiveDepts = rptDeptRanks.filter(d => d.hasData);
  const rptCompRate = rptActiveDepts.length ? Math.round(rptActiveDepts.reduce((a, d) => a + d.rate, 0) / rptActiveDepts.length * 10) / 10 : 0;
  const rptMembers = users
    .filter(u => u.role === "member" || u.role === "manager")
    .map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const hasData = kd.krs.some(kr => rptSubs.some(s => s.memberId === u.id && s.krId === kr.id));
      const rate = hasData ? calcMemberRate(u.id, kd.krs, rptSubs) : 0;
      return { ...u, rate, hasData, status: getStatus(rate) };
    }).sort((a, b) => b.rate - a.rate);
  const allMembers = users
    .filter(u => u.role === "member" || u.role === "manager")
    .map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const r = memberHasRateKrs(kd.krs) ? calcMemberRate(u.id, kd.krs, ovSubs) : null;
      const hasData = kd.krs.some(kr => ovSubs.some(s => s.memberId === u.id && s.krId === kr.id));
      const dept = depts.find(d => d.id === u.deptId);
      const deptName = dept?.name || "—";
      const primaryTeam = dept?.teams.find(t => t.id === u.teamId);
      const secondTeam = u.secondTeamId ? dept?.teams.find(t => t.id === u.secondTeamId) : null;
      const teamName = primaryTeam ? (secondTeam ? `${primaryTeam.name} / ${secondTeam.name}` : primaryTeam.name) : "—";
      return { ...u, deptName, teamName, rate: r, hasData, status: getStatus(r) };
    })
    .sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1));

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
    if (newKr.krType !== "tracker" && newKr.krType !== "progress" && !newKr.useMonthlyTargets && newKr.target === "") return;
    if (newKr.krType === "progress" && newKr.target === "") return;
    const newId = `N${Date.now().toString(36).slice(-4).toUpperCase()}`;
    const baseKr = { id: newId, label: newKr.label, unit: newKr.unit.trim(), dataSource: newKr.dataSource.trim(), operator: newKr.operator || ">=", period: newKr.period || "monthly" };
    const kr = newKr.krType === "tracker"
      ? { ...baseKr, type: "tracker", target: 0, actual: null, disallowZero: !!newKr.disallowZero }
      : newKr.krType === "progress"
      ? { ...baseKr, type: "progress", target: Number(newKr.target), actual: null }
      : newKr.krType === "manager-fill"
      ? { ...baseKr, type: "manager-fill", target: Number(newKr.target), actual: null }
      : newKr.krType === "project_profit"
      ? { ...baseKr, type: "project_profit", period: "annual", target: Number(newKr.target), krYear: Number(newKr.krYear) || new Date().getFullYear(), actual: null }
      : newKr.useMonthlyTargets
        ? { ...baseKr, monthlyTargets: Object.fromEntries(getFYMonths().map(m => [m.key, 0])), monthlyActuals: {}, ...(Number(newKr.dreamTarget) > 0 && { annualTarget: Number(newKr.dreamTarget) }) }
        : { ...baseKr, target: Number(newKr.target), actual: null };
    dispatch({ type: "ADD_KR", deptId, teamId, kr });
    if (teamId) triggerSyncPrompt(deptId, teamId);
    if (newKr.useMonthlyTargets) setExpandedMonthlyKr(newId);
    setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false }); setAddTarget(null);
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

  function previewCheckin(periods, scope = {}) {
    const periodDataList = periods.map(period => {
      const periodKey = prevPeriodKey(period);
      const dateRange = periodDateRange(period, periodKey);
      const existing = new Set(okrSubmissions.filter(s => s.period === period && s.periodKey === periodKey).map(s => `${s.memberId}:${s.krId}`));
      return { period, periodKey, dateRange, existing };
    });
    const userPool = resolveScopePool(scope);
    const recipients = [];
    for (const u of userPool) {
      const dept = depts.find(d => d.id === u.deptId);
      if (!dept) continue;
      const userSections = [];
      for (const { period, periodKey, dateRange, existing } of periodDataList) {
        const krList = [];
        (memberData[u.id]?.krs || []).filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
        dept.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
        dept.teams.forEach(t => { if (t.members?.includes(u.id) || u.teamId === t.id || u.secondTeamId === t.id) t.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr)); });
        const uniqueKrs = [...new Map(krList.map(kr => [kr.id, kr])).values()];
        const freshKrs = uniqueKrs.filter(kr => !existing.has(`${u.id}:${kr.id}`));
        if (freshKrs.length) userSections.push({ period, periodKey, dateRange, krs: freshKrs });
      }
      if (!userSections.length) continue;
      recipients.push({ user: u, dept, sections: userSections });
    }
    setCheckinPreview({ periods, scope, recipients });
  }

  async function sendCheckin(periods, scope = {}) {
    setSendingCheckin(true);
    let ctr = Date.now();
    const newSubs = [];
    const emailPromises = [];
    const userPool = resolveScopePool(scope);
    const periodDataList = periods.map(period => {
      const periodKey = prevPeriodKey(period);
      const dateRange = periodDateRange(period, periodKey);
      const existing = new Set(okrSubmissions.filter(s => s.period === period && s.periodKey === periodKey).map(s => `${s.memberId}:${s.krId}`));
      return { period, periodKey, dateRange, existing };
    });
    for (const u of userPool) {
      const dept = depts.find(d => d.id === u.deptId);
      if (!dept) continue;
      const emailSections = [];
      for (const { period, periodKey, dateRange, existing } of periodDataList) {
        const krList = [];
        (memberData[u.id]?.krs || []).filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
        dept.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr));
        dept.teams.forEach(t => { if (t.members?.includes(u.id) || u.teamId === t.id || u.secondTeamId === t.id) t.krs.filter(kr => (kr.period || "monthly") === period).forEach(kr => krList.push(kr)); });
        if (!krList.length) continue;
        const uniqueKrs = [...new Map(krList.map(kr => [kr.id, kr])).values()];
        const freshKrs = uniqueKrs.filter(kr => !existing.has(`${u.id}:${kr.id}`) && kr.type !== "manager-fill" && kr.type !== "project_profit");
        if (!freshKrs.length) continue;
        const monthKey = period === "monthly" ? periodKey
          : period === "weekly" ? (() => { const d = new Date(Date.now() - 7 * 86400000); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`; })()
          : currentFYMonthKey();
        const resolveTarget = kr => kr.monthlyTargets ? (kr.monthlyTargets[monthKey] ?? kr.target ?? 0) : (kr.target ?? 0);
        freshKrs.forEach(kr => { newSubs.push({ id: `os_${(ctr++).toString(36)}`, memberId: u.id, memberName: u.name, deptId: u.deptId, krId: kr.id, krLabel: kr.label, krTarget: resolveTarget(kr), krUnit: kr.unit || "", krOperator: kr.operator || ">=", krType: kr.type || "", krIsMonthly: !!(kr.monthlyTargets), krDisallowZero: !!(kr.disallowZero), period, periodKey, dateRange, sentAt: new Date().toISOString(), answeredAt: null, answer: null, approval: "pending", approvedBy: null }); });
        const krsForEmail = freshKrs.map(kr => ({ ...kr, target: resolveTarget(kr), isMonthly: !!(kr.monthlyTargets) }));
        emailSections.push({ period, periodKey, dateRange, krs: krsForEmail });
      }
      if (emailSections.length && u.email) {
        const emailTemplates = settings?.emailTemplates || {};
        const template = periods.length === 1
          ? { ...emailTemplates.default, ...(emailTemplates[periods[0]] || {}) }
          : { ...emailTemplates.default };
        const sendingPeriodKeys = periodDataList.map(pd => pd.periodKey);
        const userOverdue = (okrSubmissions || []).filter(s => s.memberId === u.id && s.answer === null && !sendingPeriodKeys.includes(s.periodKey))
          .map(s => ({ krLabel: s.krLabel, period: s.period, dateRange: s.dateRange, periodKey: s.periodKey }));
        const totalKrCount = emailSections.reduce((n, sec) => n + sec.krs.length, 0);
        emailPromises.push(
          fetch("/api/send-email", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ to: u.email, name: u.name, sections: emailSections, template, overdueSubs: userOverdue }) })
            .then(res => res.ok
              ? { name: u.name, email: u.email, krCount: totalKrCount, success: true }
              : { name: u.name, email: u.email, krCount: totalKrCount, success: false, reason: `HTTP ${res.status}` })
            .catch(err => ({ name: u.name, email: u.email, krCount: totalKrCount, success: false, reason: err.message || "Network error" }))
        );
      }
    }
    if (newSubs.length) dispatch({ type: "CREATE_OKR_SUBMISSIONS", submissions: newSubs });
    const memberCount = new Set(newSubs.map(s => s.memberId)).size;
    const emailOutcomes = await Promise.all(emailPromises);
    const emailFailures = emailOutcomes.filter(o => !o.success);
    if (emailOutcomes.length > 0 || newSubs.length > 0) {
      const PERIOD_LABELS_L = { daily: "Daily", weekly: "Weekly", monthly: "Monthly", quarterly: "Quarterly", biannual: "Bi-Annual", annual: "Annual" };
      const periodKeys = Object.fromEntries(periodDataList.map(pd => [pd.period, pd.periodKey]));
      const dateRanges = Object.fromEntries(periodDataList.map(pd => [pd.period, pd.dateRange]));
      dispatch({ type: "LOG_EMAIL_SEND", log: {
        id: `el_${Date.now().toString(36)}`,
        sentAt: new Date().toISOString(),
        periods,
        period: periods.length === 1 ? periods[0] : periods.map(p => PERIOD_LABELS_L[p] || p).join("+"),
        periodKeys,
        periodKey: periods.length === 1 ? periodDataList[0].periodKey : periodDataList.map(pd => pd.periodKey).join("+"),
        dateRanges,
        dateRange: periods.length === 1 ? periodDataList[0].dateRange : periodDataList.map(pd => pd.dateRange).join(" + "),
        scope,
        submissionsCreated: newSubs.length,
        recipientCount: emailOutcomes.length,
        recipients: emailOutcomes,
        failureCount: emailFailures.length,
      }});
    }
    setCheckinResult({ count: newSubs.length, memberCount, periods, scope, emailFailures });
    setSendingCheckin(false);
    setTimeout(() => setCheckinResult(null), 8000);
  }

  return (
    <MobileContext.Provider value={{ isMobile, drawerOpen, setDrawerOpen }}>
    <div style={{ display: "flex", minHeight: "100dvh", fontFamily: F.body, background: T.bg, color: T.text }}>
      {logPopup && <div onClick={() => setLogPopup(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}><div onClick={e => e.stopPropagation()} style={{ background: T.surface, borderRadius: 12, border: `1px solid ${T.border}`, width: "100%", maxWidth: 640, maxHeight: "75vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}><div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexShrink: 0 }}><div><div style={{ fontSize: 15, fontWeight: 700 }}>{logPopup.projName}</div>{logPopup.date && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>{logPopup.date}</div>}</div><button onClick={() => setLogPopup(null)} style={{ background: "none", border: "none", cursor: "pointer", color: T.textMuted, fontSize: 18, lineHeight: 1, padding: "0 2px", marginLeft: 12 }}>✕</button></div><div style={{ padding: "16px 20px", overflowY: "auto", fontSize: 14, lineHeight: 1.65, color: T.text, whiteSpace: "pre-wrap" }}>{logPopup.text}</div></div></div>}
      {checkinPreview && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div style={{ background: T.surface, borderRadius: 14, boxShadow: "0 8px 40px rgba(0,0,0,0.22)", width: "100%", maxWidth: 620, maxHeight: "80vh", display: "flex", flexDirection: "column" }}>
            <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${T.border}` }}>
              <div style={{ fontWeight: 800, fontSize: 17 }}>Confirm: Send {(checkinPreview.periods || [checkinPreview.period]).map(p => p.charAt(0).toUpperCase() + p.slice(1)).join(" & ")} Check-In</div>
              <div style={{ fontSize: 12, color: T.textDim, marginTop: 5 }}>
                <span>To: <strong style={{ color: T.text }}>{scopeLabel(checkinPreview.scope || {})}</strong></span>
                <span style={{ margin: "0 6px", color: T.border }}>·</span>
                <span>{checkinPreview.recipients.length} recipient{checkinPreview.recipients.length !== 1 ? "s" : ""}</span>
                <span style={{ margin: "0 6px", color: T.border }}>·</span>
                <span>{checkinPreview.recipients.reduce((n, r) => n + (r.sections ? r.sections.reduce((m, s) => m + s.krs.length, 0) : (r.krs?.length || 0)), 0)} KRs total</span>
              </div>
            </div>
            <div style={{ flex: 1, overflowY: "auto", padding: "14px 24px" }}>
              {checkinPreview.recipients.length === 0 ? (
                <div style={{ padding: "28px 0", color: T.textDim, fontSize: 13, textAlign: "center" }}>
                  No new check-ins to send — all KRs for selected periods already have submissions.
                </div>
              ) : checkinPreview.recipients.map(({ user: u, dept, sections, krs }) => (
                <div key={u.id} style={{ marginBottom: 10, padding: "10px 14px", borderRadius: 8, background: T.raised, border: `1px solid ${T.border}` }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 7, flexWrap: "wrap" }}>
                    <Avatar letters={u.av || "?"} size={26} />
                    <span style={{ fontWeight: 700, fontSize: 13 }}>{u.name}</span>
                    {u.role === "manager" && <span style={{ fontSize: 10, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Manager</span>}
                    <span style={{ fontSize: 11, color: T.textMuted, background: T.surface, borderRadius: 5, padding: "1px 7px", border: `1px solid ${T.border}` }}>{dept.name}</span>
                    {u.email && <span style={{ fontSize: 11, color: T.textDim }}>{u.email}</span>}
                  </div>
                  {(sections || [{ period: checkinPreview.period, periodKey: checkinPreview.periodKey, dateRange: checkinPreview.dateRange, krs: krs || [] }]).map(sec => (
                    <div key={sec.period} style={{ marginBottom: 5 }}>
                      {(checkinPreview.periods || []).length > 1 && (
                        <div style={{ fontSize: 10, fontWeight: 700, color: T.brand, textTransform: "uppercase", letterSpacing: "0.06em", marginBottom: 3 }}>
                          {sec.period.charAt(0).toUpperCase() + sec.period.slice(1)} · {sec.dateRange || sec.periodKey}
                        </div>
                      )}
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                        {sec.krs.map(kr => (
                          <span key={kr.id} style={{ fontSize: 11, background: kr.type === "tracker" ? "#ede9fe" : T.surface, color: kr.type === "tracker" ? "#7c3aed" : T.text, border: `1px solid ${kr.type === "tracker" ? "#c4b5fd" : T.border}`, borderRadius: 6, padding: "2px 8px" }}>
                            {kr.label}
                          </span>
                        ))}
                      </div>
                    </div>
                  ))}
                </div>
              ))}
            </div>
            <div style={{ padding: "14px 24px", borderTop: `1px solid ${T.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn onClick={() => setCheckinPreview(null)}>Cancel</Btn>
              <Btn primary disabled={checkinPreview.recipients.length === 0}
                onClick={() => { const { periods, period, scope } = checkinPreview; setCheckinPreview(null); sendCheckin(periods || [period], scope); }}>
                📨 Confirm &amp; Send to {checkinPreview.recipients.length} recipient{checkinPreview.recipients.length !== 1 ? "s" : ""}
              </Btn>
            </div>
          </div>
        </div>
      )}
      {showCheckinDialog && (
        <div onClick={() => setShowCheckinDialog(false)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
          <div onClick={e => e.stopPropagation()} style={{ background: T.surface, borderRadius: 14, boxShadow: "0 8px 40px rgba(0,0,0,0.22)", width: "100%", maxWidth: 500 }}>
            <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${T.border}` }}>
              <div style={{ fontWeight: 800, fontSize: 17 }}>📨 Send Check-In</div>
              <div style={{ fontSize: 12, color: T.textDim, marginTop: 4 }}>Select periods and recipients, then preview before sending.</div>
            </div>
            <div style={{ padding: "20px 24px" }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Periods</div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 20 }}>
                {["daily","weekly","monthly","quarterly","biannual","annual"].map(p => (
                  <label key={p} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 13, fontWeight: 600, cursor: "pointer", userSelect: "none", padding: "5px 11px", borderRadius: 7, border: `1px solid ${checkinPeriods.includes(p) ? T.brandBorder : T.border}`, background: checkinPeriods.includes(p) ? T.brandDim : T.raised, color: checkinPeriods.includes(p) ? T.brand : T.textDim }}>
                    <input type="checkbox" checked={checkinPeriods.includes(p)} onChange={e => setCheckinPeriods(prev => e.target.checked ? [...prev, p] : prev.filter(x => x !== p))} style={{ accentColor: T.brand, margin: 0 }} />
                    {p.charAt(0).toUpperCase() + p.slice(1)}
                  </label>
                ))}
              </div>
              <div style={{ fontSize: 12, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 10 }}>Recipients</div>
              <div style={{ display: "flex", gap: 8, flexDirection: "column" }}>
                <select value={checkinScope.deptId} onChange={e => setCheckinScope({ deptId: e.target.value, teamId: "", userId: "" })}
                  style={{ padding: "8px 12px", fontSize: 13, fontFamily: F.body, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none", width: "100%" }}>
                  <option value="">All Departments</option>
                  {depts.slice().sort((a,b) => a.name.localeCompare(b.name)).map(d => (
                    <option key={d.id} value={d.id}>{d.name}</option>
                  ))}
                </select>
                {checkinScope.deptId && (() => {
                  const scopeTeams = depts.find(d => d.id === checkinScope.deptId)?.teams || [];
                  return scopeTeams.length > 0 ? (
                    <select value={checkinScope.teamId} onChange={e => setCheckinScope(p => ({ ...p, teamId: e.target.value, userId: "" }))}
                      style={{ padding: "8px 12px", fontSize: 13, fontFamily: F.body, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none", width: "100%" }}>
                      <option value="">All Teams</option>
                      {scopeTeams.slice().sort((a,b) => a.name.localeCompare(b.name)).map(t => (
                        <option key={t.id} value={t.id}>{t.name}</option>
                      ))}
                    </select>
                  ) : null;
                })()}
                <select value={checkinScope.userId} onChange={e => setCheckinScope(p => ({ ...p, userId: e.target.value }))}
                  style={{ padding: "8px 12px", fontSize: 13, fontFamily: F.body, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, color: T.text, outline: "none", width: "100%" }}>
                  <option value="">All Members</option>
                  {resolveScopePool({ deptId: checkinScope.deptId, teamId: checkinScope.teamId }).slice().sort((a,b) => a.name.localeCompare(b.name)).map(u => (
                    <option key={u.id} value={u.id}>{u.name}{u.deptId !== checkinScope.deptId ? ` (${depts.find(d=>d.id===u.deptId)?.name||""})` : ""}</option>
                  ))}
                </select>
              </div>
            </div>
            <div style={{ padding: "14px 24px", borderTop: `1px solid ${T.border}`, display: "flex", justifyContent: "flex-end", gap: 10 }}>
              <Btn onClick={() => setShowCheckinDialog(false)}>Cancel</Btn>
              <Btn primary disabled={!checkinPeriods.length} onClick={() => { setShowCheckinDialog(false); previewCheckin(checkinPeriods, checkinScope); }}>
                Preview &amp; Send
              </Btn>
            </div>
          </div>
        </div>
      )}
      <Side items={navItems} active={page} onSelect={p => { setPage(p); setSelDept(null); setAdminSelDept(null); }} user={user} onLogout={onLogout}
        subItems={deptSubItems} subItemsFor="okr-mgmt" activeSubItem={adminSelDept ? "__setup__" : "__all__"} onSelectSubItem={id => { setPage("okr-mgmt"); if (id === "__setup__") { setAdminSelDept(depts[0]?.id || null); } else { setAdminSelDept(null); } setSelTeam(null); setAddTarget(null); setExpandedMonthlyKr(null); }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "users" && <UserMgmtPage users={users} depts={depts} dispatch={dispatch} currentUserId={user.id} onImpersonate={onImpersonate} />}

        {page === "overview" && (<>
          <Header title="Company Overview" sub={(() => {
            const now = new Date();
            if (overviewView === "weekly") {
              const dow = now.getDay();
              const mon = new Date(now); mon.setHours(0,0,0,0); mon.setDate(now.getDate() - (dow === 0 ? 6 : dow - 1) - 7);
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
              {[["weekly","Weekly Overview"],["monthly","Monthly Overview"],["annual","Annual Overview"],["financial","Financial Performance"]].map(([v,label]) => (
                <Btn key={v} small primary={overviewView === v} onClick={() => setOverviewView(v)}>{label}</Btn>
              ))}
            </div>
            {overviewView !== "financial" && (<>
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
                : deptRanks.map((d, i) => {
                  const isExpanded = ovExpandedDept === d.id;
                  const deptMembers = isExpanded ? users
                    .filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id)
                    .map(u => { const excl = !!u.excludeFromRate; const kd = memberData[u.id] || { krs: [] }; const mr = (!excl && memberHasRateKrs(kd.krs)) ? calcMemberRate(u.id, filtKrs(kd.krs), ovSubs) : null; return { ...u, rate: mr, status: getStatus(mr), excluded: excl }; })
                    .sort((a, b) => { if (a.excluded !== b.excluded) return a.excluded ? 1 : -1; return (b.rate ?? -1) - (a.rate ?? -1); }) : [];
                  return (
                    <Card key={d.id} style={{ marginBottom: 8, overflow: "hidden" }}>
                      <div style={{ overflowX: "auto" }}>
                      <div onClick={() => setOvExpandedDept(p => p === d.id ? null : d.id)} style={{ padding: "16px 20px", cursor: "pointer", display: "grid", gridTemplateColumns: "36px 1fr 60px 180px 80px 24px", alignItems: "center", gap: 14, minWidth: 480 }}>
                        <span style={{ fontSize: 18, fontWeight: 900, fontFamily: F.mono, color: i === 0 ? T.ok : i === deptRanks.length - 1 ? T.bad : T.textMuted }}>#{i + 1}</span>
                        <div><div style={{ fontSize: 16, fontWeight: 700 }}>{d.name}</div><div style={{ fontSize: 12, color: T.textMuted }}>{d.college} · {d.head} · {d.teams.length} teams</div></div>
                        <span style={{ textAlign: "right", fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[d.status].color }}>{d.rate.toFixed(1)}%</span>
                        <Bar value={d.rate} status={d.status} h={7} />
                        <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={d.status} /></div>
                        <span style={{ fontSize: 11, color: T.textMuted, textAlign: "right" }}>{isExpanded ? "▲" : "▼"}</span>
                      </div>
                      </div>
                      {isExpanded && (
                        <div style={{ borderTop: `1px solid ${T.border}`, padding: "14px 20px", background: T.bgSoft }}>
                          {deptMembers.length === 0
                            ? <div style={{ fontSize: 13, color: T.textMuted }}>No members in this department yet.</div>
                            : deptMembers.map((m, mi) => (
                              <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "8px 0", borderBottom: mi < deptMembers.length - 1 ? `1px solid ${T.border}` : "none", opacity: m.excluded ? 0.55 : 1 }}>
                                <span style={{ fontSize: 13, fontWeight: 400, color: T.textMuted, fontFamily: F.mono, width: 20, textAlign: "right", flexShrink: 0 }}>#{mi + 1}</span>
                                <Avatar letters={m.av || m.name?.slice(0, 2)} size={30} />
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ fontWeight: 600, fontSize: 14, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m.name}</div>
                                  <div style={{ fontSize: 12, color: T.textMuted }}>{m.title || ""}</div>
                                </div>
                                {m.excluded
                                  ? <span style={{ fontSize: 11, color: T.textDim, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "2px 8px", flexShrink: 0 }}>Excluded from rate</span>
                                  : m.rate !== null ? (<>
                                    <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 14, color: STATUS_THEME[m.status].color, flexShrink: 0 }}>{m.rate.toFixed(1)}%</span>
                                    <div style={{ width: 100, flexShrink: 0 }}><Bar value={m.rate} status={m.status} h={5} /></div>
                                    <Tag type={m.status} small />
                                  </>) : <span style={{ fontSize: 12, color: T.textDim }}>No data</span>
                                }
                                <button onClick={() => dispatch({ type: "UPDATE_USER", userId: m.id, updates: { excludeFromRate: !m.excludeFromRate } })} title={m.excluded ? "Include in dept rate" : "Exclude from dept rate"} style={{ background: m.excluded ? T.raised : T.badDim, border: `1px solid ${m.excluded ? T.border : T.badBorder}`, borderRadius: 5, padding: "2px 7px", cursor: "pointer", color: m.excluded ? T.textMuted : T.bad, fontSize: 11, fontWeight: 700, flexShrink: 0, fontFamily: F.body }}>{m.excluded ? "Include" : "Exclude"}</button>
                              </div>
                            ))
                          }
                        </div>
                      )}
                    </Card>
                  );
                })}
            </div>
            </>)}

            {overviewView === "financial" && <FinancialPerformancePage state={state} dispatch={dispatch} />}

          </Pane>
        </>)}

        {page === "okr-mgmt" && (() => {
          const PERIODS = [{ id: "all", label: "All" }, { id: "daily", label: "Daily" }, { id: "weekly", label: "Weekly" }, { id: "monthly", label: "Monthly" }, { id: "quarterly", label: "Quarterly" }, { id: "biannual", label: "Bi-Annual" }, { id: "annual", label: "Annual" }];
          const filterP = krs => {
            const byPeriod = adminOkrPeriod === "all" ? krs : krs.filter(kr => (kr.period || "monthly") === adminOkrPeriod);
            return byPeriod.filter(kr => kr.type !== "tracker" || kr.showInOverview !== false);
          };
          const pLabel = adminOkrPeriod === "all" ? "All Periods" : periodDateRange(adminOkrPeriod, adminOkrPeriod === "weekly" ? prevPeriodKey(adminOkrPeriod) : currentPeriodKey(adminOkrPeriod));
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
            const _pkAdmin = adminOkrPeriod === "weekly" ? prevPeriodKey(adminOkrPeriod) : currentPeriodKey(adminOkrPeriod);
            const hasSub = adminOkrPeriod === "all" ? (isMonthly ? Object.values(kr.monthlyActuals || {}).some(v => v != null) : kr.actual != null) : okrSubmissions.some(s => s.krId === kr.id && s.period === (kr.period || "monthly") && s.periodKey === _pkAdmin && s.answer !== null);
            return (
              <Fragment key={kr.id}>
              <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                <div>
                  <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                  {kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                  {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                  {kr.type === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Progress · affects rate proportionally</span>}
                  {kr.type === "manager-fill" && <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Mgr Fill · manager assesses</span>}
                  {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Monthly Breakdown</span>}
                  {adminOkrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                </div>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? `${kr.operator||">="} ${fmt(curTarget)} this mo.` : kr.type === "progress" ? fmt(kr.target) : `${kr.operator || ">="} ${fmt(kr.target)}${kr.unit ? ` ${kr.unit}` : ""}`}</span>}
                {kr.type === "tracker"
                  ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textDim }}>—</span>
                  : kr.type === "manager-fill"
                  ? <span style={{ textAlign: "right", fontFamily: F.mono, color: "#d97706", fontSize: 12 }}>via mgr</span>
                  : isMonthly
                  ? <NumInput value={curActual} onChange={n => dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: curKey, field: "actual", value: n })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                  : <NumInput value={kr.actual} onChange={n => { dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "actual", value: n }); if (teamId) triggerSyncPrompt(deptId, teamId); }} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />}
                <span style={{ fontSize: 12, color: T.textMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.dataSource || "—"}</span>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 14, fontWeight: 700, color: "#7c3aed" }}>{fmt(isMonthly ? curActual : kr.actual)}{kr.unit ? <span style={{ fontSize: 11, fontWeight: 400 }}> {kr.unit}</span> : ""}</span> : hasSub ? <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[st].color }}>{pct.toFixed(0)}%</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: T.textDim }}>N/A</span>}
                {kr.type === "tracker" ? <span /> : hasSub ? <Bar value={pct} status={st} h={5} /> : <span />}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 4 }}>
                  {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="View all months" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "2px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 11, fontWeight: 700 }}>📅</button>}
                  {kr.type === "tracker" ? null : hasSub ? <Tag type={st} small /> : null}
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
                            return <th key={key} style={{ textAlign: "center", padding: "6px 3px", borderBottom: `2px solid ${isCur ? T.brand : T.border}`, fontSize: 11, fontWeight: isCur ? 700 : 400, color: isCur ? T.brand : T.textDim, minWidth: 62, background: isCur ? T.brandDim : T.surface }}>{label.split(" ")[0]}{isCur ? " ●" : ""}</th>;
                          })}
                          <th style={{ textAlign: "right", padding: "6px 10px", borderBottom: `2px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, minWidth: 80, background: T.surface }}>FY Total</th>
                          <th style={{ textAlign: "center", padding: "6px 10px", borderBottom: `2px solid ${T.okBorder}`, fontSize: 11, fontWeight: 700, color: T.ok, minWidth: 100, background: T.okDim }}>Dream Target</th>
                        </tr>
                      </thead>
                      <tbody>
                        <tr>
                          <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Perf. Target</td>
                          {fyMs.map(({ key }) => { const isCur = key === curKey; const t = kr.monthlyTargets[key] || 0; return <td key={key} style={{ padding: "4px 6px", textAlign: "right", fontFamily: F.mono, fontSize: 12, background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}`, color: T.textMuted }}>{fmt(t)}</td>; })}
                          <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annSumTarget)}</td>
                          <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, borderLeft: `1px solid ${T.okBorder}`, textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: T.ok }}>{annDream > 0 ? fmt(annDream) : <span style={{ color: T.textDim, fontWeight: 400 }}>—</span>}{kr.unit && annDream > 0 && <div style={{ fontSize: 10, color: T.ok, fontWeight: 400 }}>{kr.unit}</div>}</td>
                        </tr>
                        <tr>
                          <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Actual</td>
                          {fyMs.map(({ key }) => { const isCur = key === curKey; const a = (kr.monthlyActuals || {})[key] || 0; return <td key={key} style={{ padding: "3px 3px", background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}` }}><NumInput value={a} onChange={n => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "actual", value: n }); if (teamId) triggerSyncPrompt(deptId, teamId); }} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} /></td>; })}
                          <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, color: STATUS_THEME[annSt].color, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annActual)}</td>
                          <td style={{ padding: "4px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, textAlign: "center", color: T.textDim, fontSize: 12, borderLeft: `1px solid ${T.okBorder}` }}>—</td>
                        </tr>
                        <tr>
                          <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 11, color: T.textDim, background: T.surface }}>Achievement</td>
                          {fyMs.map(({ key }) => { const t = kr.monthlyTargets[key] || 0; const a = (kr.monthlyActuals || {})[key] || 0; const mp = t > 0 ? Math.min((a / t) * 100, 100) : null; const isCur = key === curKey; return <td key={key} style={{ padding: "4px 4px", textAlign: "center", background: isCur ? T.brandDim : "transparent" }}>{mp !== null ? <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[getStatus(mp)].color }}>{mp.toFixed(0)}%</span> : <span style={{ color: T.textDim, fontSize: 11 }}>—</span>}</td>; })}
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
          });
          const renderSection = (krs, deptId, teamId) => {
            if (krs.length === 0) return <div style={{ fontSize: 13, color: T.textMuted, padding: "10px 0" }}>No {adminOkrPeriod === "all" ? "" : adminOkrPeriod + " "}KRs for this section yet.</div>;
            return (
              <Card style={{ overflow: "hidden" }}>
                <div style={{ overflowX: "auto" }}><div style={{ minWidth: 760 }}>
                <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                  <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Performance Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Data Source</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                </div>
                {renderRows(krs, deptId, teamId)}
                </div></div>
              </Card>
            );
          };
          const DeptNav = () => (
            <div style={{ display: "flex", gap: 4, marginBottom: 20, paddingBottom: 12, borderBottom: `1px solid ${T.border}`, flexWrap: "wrap", alignItems: "center" }}>
              <span style={{ fontSize: 12, color: T.textMuted, marginRight: 4 }}>Department:</span>
              <Btn small primary={!adminSelDept} onClick={() => { setAdminSelDept(null); setExpandedMonthlyKr(null); }}>All</Btn>
              {depts.map(dd => <Btn key={dd.id} primary={adminSelDept === dd.id} small onClick={() => { setAdminSelDept(dd.id); setExpandedMonthlyKr(null); }}>{dd.name}</Btn>)}
            </div>
          );
          const selDeptObj = adminSelDept ? depts.find(d => d.id === adminSelDept) : null;
          if (selDeptObj) {
            const d = selDeptObj;
            const filterSetupP = krs => adminOkrPeriod === "all" ? (krs || []) : (krs || []).filter(kr => (kr.period || "monthly") === adminOkrPeriod);
            const KCOL_S = "50px 1fr 100px 90px 130px 80px 1fr 92px";
            const renderSetupRows = (krs, deptId, teamId) => krs.map((kr, i) => {
              if (addTarget === `edit-${kr.id}`) {
                const isTracker = newKr.krType === "tracker";
                const isMgrFill = newKr.krType === "manager-fill";
                const isProjectProfit = newKr.krType === "project_profit";
                const isStandard = !isTracker && newKr.krType !== "progress" && !isMgrFill && !isProjectProfit;
                const sel = { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 12px", color: T.text, fontSize: 14, fontFamily: F.body, outline: "none" };
                const saveEdit = () => {
                  if (!newKr.label) return;
                  if (newKr.krType !== "tracker" && newKr.krType !== "progress" && newKr.krType !== "project_profit" && !newKr.useMonthlyTargets && newKr.target === "") return;
                  if ((newKr.krType === "progress" || newKr.krType === "project_profit") && newKr.target === "") return;
                  const base = { id: kr.id, label: newKr.label, unit: newKr.unit.trim(), dataSource: newKr.dataSource.trim(), operator: newKr.operator || ">=", period: newKr.period || "monthly" };
                  let updated;
                  if (newKr.krType === "tracker") updated = { ...base, type: "tracker", target: 0, actual: kr.actual, disallowZero: !!newKr.disallowZero };
                  else if (newKr.krType === "progress") updated = { ...base, type: "progress", target: Number(newKr.target), actual: kr.actual };
                  else if (newKr.krType === "manager-fill") updated = { ...base, type: "manager-fill", target: Number(newKr.target), actual: kr.actual };
                  else if (newKr.krType === "project_profit") updated = { ...base, type: "project_profit", period: "annual", target: Number(newKr.target), krYear: Number(newKr.krYear) || new Date().getFullYear(), actual: null };
                  else if (newKr.useMonthlyTargets) updated = { ...base, monthlyTargets: newKr.monthlyTargets, monthlyActuals: kr.monthlyActuals || {}, ...(Number(newKr.dreamTarget) > 0 && { annualTarget: Number(newKr.dreamTarget) }) };
                  else updated = { ...base, target: Number(newKr.target), actual: kr.actual };
                  dispatch({ type: "REPLACE_KR", deptId, teamId, krId: kr.id, kr: updated });
                  if (teamId) triggerSyncPrompt(deptId, teamId);
                  setAddTarget(null);
                  setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", krYear: "", monthlyTargets: {}, disallowZero: false });
                };
                return (
                  <div key={kr.id} style={{ padding: "14px 16px", background: T.warnDim, borderTop: `2px solid ${T.warn}`, borderBottom: `1px solid ${T.border}` }}>
                    <div style={{ fontSize: 12, fontWeight: 700, color: T.warn, marginBottom: 10, letterSpacing: "0.04em", textTransform: "uppercase" }}>Edit Key Result — {kr.id}</div>
                    <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
                      <Input value={newKr.label} onChange={e => setNewKr(p => ({ ...p, label: e.target.value }))} placeholder="Key result name *" style={{ flex: 2, minWidth: 200 }} />
                      <select value={newKr.krType} onChange={e => setNewKr(p => ({ ...p, krType: e.target.value, useMonthlyTargets: false }))} style={{ ...sel, flex: 1, minWidth: 160 }}>
                        <option value="">Standard (Yes / No)</option>
                        <option value="tracker">Tracker (number only)</option>
                        <option value="progress">Progress (cumulative)</option>
                        <option value="manager-fill">Manager Fill</option>
                        <option value="project_profit">Project Profit (auto)</option>
                      </select>
                      {!isProjectProfit && <select value={newKr.period || "monthly"} onChange={e => setNewKr(p => ({ ...p, period: e.target.value }))} style={{ ...sel, minWidth: 120 }}>
                        <option value="daily">Daily</option>
                        <option value="weekly">Weekly</option>
                        <option value="monthly">Monthly</option>
                        <option value="quarterly">Quarterly</option>
                        <option value="biannual">Bi-Annual</option>
                        <option value="annual">Annual</option>
                      </select>}
                    </div>
                    {isProjectProfit ? (
                      <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap", alignItems: "center" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>$</span>
                          <Input value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Annual profit target *" style={{ width: 180, fontFamily: F.mono }} />
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>Year:</span>
                          <Input value={newKr.krYear || ""} onChange={e => setNewKr(p => ({ ...p, krYear: e.target.value }))} placeholder={String(new Date().getFullYear())} style={{ width: 90, fontFamily: F.mono }} />
                        </div>
                      </div>
                    ) : (
                    <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap", alignItems: "center" }}>
                      {!isTracker && !newKr.useMonthlyTargets && (<>
                        <select value={newKr.operator || ">="} onChange={e => setNewKr(p => ({ ...p, operator: e.target.value }))} style={{ ...sel, width: 80 }}>
                          <option value=">=">≥</option>
                          <option value="<=">≤</option>
                          <option value="=">=</option>
                        </select>
                        <Input value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Target *" style={{ width: 110 }} />
                      </>)}
                      <Input value={newKr.unit} onChange={e => setNewKr(p => ({ ...p, unit: e.target.value }))} placeholder="Unit (e.g. $, Days)" style={{ width: 150 }} />
                      <Input value={newKr.dataSource} onChange={e => setNewKr(p => ({ ...p, dataSource: e.target.value }))} placeholder="Data source" style={{ flex: 1, minWidth: 140 }} />
                      {isStandard && (
                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                          <input type="checkbox" checked={!!newKr.useMonthlyTargets} onChange={e => setNewKr(p => ({ ...p, useMonthlyTargets: e.target.checked }))} style={{ accentColor: T.brand }} />
                          Monthly targets
                        </label>
                      )}
                      {isTracker && (
                        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                          <input type="checkbox" checked={!!newKr.disallowZero} onChange={e => setNewKr(p => ({ ...p, disallowZero: e.target.checked }))} style={{ accentColor: T.bad }} />
                          Block zero submission
                        </label>
                      )}
                    </div>
                    )}
                    {isStandard && newKr.useMonthlyTargets && (
                      <div style={{ marginBottom: 8 }}>
                        <Input value={newKr.dreamTarget} onChange={e => setNewKr(p => ({ ...p, dreamTarget: e.target.value }))} placeholder="Dream / annual target (optional)" style={{ width: 260, marginBottom: 10 }} />
                        <div style={{ border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden", maxWidth: 360 }}>
                          <div style={{ display: "grid", gridTemplateColumns: "120px 1fr", padding: "5px 10px", fontSize: 10, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em", background: T.raised, borderBottom: `1px solid ${T.border}` }}>
                            <span>Month</span><span>Target</span>
                          </div>
                          {getFYMonths().map((mo, mi) => (
                            <div key={mo.key} style={{ display: "grid", gridTemplateColumns: "120px 1fr", gap: 8, padding: "4px 10px", alignItems: "center", background: mi % 2 ? T.raised : "transparent", borderBottom: mi < getFYMonths().length - 1 ? `1px solid ${T.border}` : "none" }}>
                              <span style={{ fontSize: 12, color: T.textMuted }}>{mo.label}</span>
                              <NumInput value={newKr.monthlyTargets[mo.key] ?? 0} onChange={n => setNewKr(p => ({ ...p, monthlyTargets: { ...p.monthlyTargets, [mo.key]: n } }))} style={{ padding: "3px 8px", fontSize: 13, fontFamily: F.mono, width: 100 }} />
                            </div>
                          ))}
                        </div>
                      </div>
                    )}
                    {isMgrFill && (
                      <div style={{ fontSize: 12, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "7px 12px", marginBottom: 8 }}>
                        Manager Fill — member will not receive a check-in; manager assesses yes/no + value in their portal.
                      </div>
                    )}
                    {isProjectProfit && (
                      <div style={{ fontSize: 12, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "7px 12px", marginBottom: 8 }}>
                        Project Profit — actual is auto-calculated from this user's completed projects in the target year. No check-ins required.
                      </div>
                    )}
                    <div style={{ display: "flex", gap: 8 }}>
                      <Btn primary small onClick={saveEdit}>✓ Save Changes</Btn>
                      <Btn small onClick={() => { setAddTarget(null); setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", krYear: "", monthlyTargets: {}, disallowZero: false }); }}>Cancel</Btn>
                    </div>
                  </div>
                );
              }
              const isMonthly = !!kr.monthlyTargets;
              const targetDisplay = kr.type === "tracker" ? "—" : kr.type === "project_profit" ? `$${(kr.target || 0).toLocaleString()} (${kr.krYear || "?"})` : isMonthly ? "Monthly breakdown" : `${kr.operator || ">="} ${fmt(kr.target)}`;
              const typeLabel = kr.type === "tracker" ? "Tracker" : kr.type === "progress" ? "Progress" : kr.type === "manager-fill" ? "Mgr Fill" : kr.type === "project_profit" ? "Proj Profit" : "Standard";
              const typeStyle = kr.type === "tracker"
                ? { color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd" }
                : kr.type === "progress"
                ? { color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}` }
                : kr.type === "manager-fill"
                ? { color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a" }
                : kr.type === "project_profit"
                ? { color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}` }
                : { color: T.textMuted, background: T.raised, border: `1px solid ${T.border}` };
              return (
                <div key={kr.id} style={{ display: "grid", gridTemplateColumns: KCOL_S, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                  <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim }}>{kr.id}</span>
                  <div>
                    <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                    {isMonthly && <div style={{ marginTop: 2 }}><span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px" }}>Monthly</span></div>}
                  </div>
                  <div><span style={{ fontSize: 10, fontWeight: 600, borderRadius: 8, padding: "2px 7px", ...typeStyle }}>{typeLabel}</span></div>
                  <span style={{ fontSize: 12, color: T.textMuted }}>{kr.period ? kr.period.charAt(0).toUpperCase() + kr.period.slice(1) : "Monthly"}</span>
                  <span style={{ fontFamily: F.mono, fontSize: 12, color: kr.type === "tracker" ? T.textDim : T.textMuted }}>{targetDisplay}</span>
                  <span style={{ fontSize: 12, color: T.textMuted }}>{kr.unit || "—"}</span>
                  <span style={{ fontSize: 12, color: T.textMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.dataSource || "—"}</span>
                  <div style={{ display: "flex", justifyContent: "flex-end", gap: 4 }}>
                    <button onClick={() => {
                      const krType = kr.type === "tracker" ? "tracker" : kr.type === "progress" ? "progress" : kr.type === "manager-fill" ? "manager-fill" : kr.type === "project_profit" ? "project_profit" : "";
                      const isMonthlyKr = !!kr.monthlyTargets;
                      setAddTarget(`edit-${kr.id}`);
                      setNewKr({ label: kr.label, target: isMonthlyKr ? "" : String(kr.target ?? ""), dreamTarget: String(kr.annualTarget || ""), unit: kr.unit || "", dataSource: kr.dataSource || "", operator: kr.operator || ">=", period: kr.period || "monthly", useMonthlyTargets: isMonthlyKr && krType === "", krType, krYear: String(kr.krYear || ""), monthlyTargets: Object.fromEntries(getFYMonths().map(m => [m.key, kr.monthlyTargets?.[m.key] ?? 0])), disallowZero: !!kr.disallowZero });
                    }} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", cursor: "pointer", color: T.brand, fontSize: 13 }}>✏</button>
                    <button onClick={() => dispatch({ type: "REMOVE_KR", deptId, teamId, krId: kr.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 6, padding: "2px 9px", cursor: "pointer", color: T.bad, fontSize: 15, fontWeight: 700, lineHeight: 1 }}>×</button>
                  </div>
                </div>
              );
            });
            const renderAddRow = (deptId, teamId) => {
              const targetKey = teamId || `dept-${deptId}`;
              if (addTarget !== targetKey) return (
                <div style={{ padding: "10px 16px" }}>
                  <Btn small onClick={() => { setAddTarget(targetKey); setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: adminOkrPeriod !== "all" ? adminOkrPeriod : "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, disallowZero: false }); }}>+ Add Key Result</Btn>
                </div>
              );
              const isTracker = newKr.krType === "tracker";
              const isMgrFill = newKr.krType === "manager-fill";
              const isProjectProfit = newKr.krType === "project_profit";
              const isStandard = !isTracker && newKr.krType !== "progress" && !isMgrFill && !isProjectProfit;
              const sel = { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 12px", color: T.text, fontSize: 14, fontFamily: F.body, outline: "none" };
              return (
                <div style={{ padding: "14px 16px", background: T.brandDim, borderTop: `2px solid ${T.brand}` }}>
                  <div style={{ fontSize: 12, fontWeight: 700, color: T.brand, marginBottom: 10, letterSpacing: "0.04em", textTransform: "uppercase" }}>New Key Result</div>
                  <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
                    <Input value={newKr.label} onChange={e => setNewKr(p => ({ ...p, label: e.target.value }))} placeholder="Key result name *" style={{ flex: 2, minWidth: 200 }} />
                    <select value={newKr.krType} onChange={e => setNewKr(p => ({ ...p, krType: e.target.value, useMonthlyTargets: false }))} style={{ ...sel, flex: 1, minWidth: 160 }}>
                      <option value="">Standard (Yes / No)</option>
                      <option value="tracker">Tracker (number only)</option>
                      <option value="progress">Progress (cumulative)</option>
                      <option value="manager-fill">Manager Fill</option>
                      <option value="project_profit">Project Profit (auto)</option>
                    </select>
                    {!isProjectProfit && <select value={newKr.period || "monthly"} onChange={e => setNewKr(p => ({ ...p, period: e.target.value }))} style={{ ...sel, minWidth: 120 }}>
                      <option value="daily">Daily</option>
                      <option value="weekly">Weekly</option>
                      <option value="monthly">Monthly</option>
                      <option value="quarterly">Quarterly</option>
                      <option value="biannual">Bi-Annual</option>
                      <option value="annual">Annual</option>
                    </select>}
                  </div>
                  {isProjectProfit ? (
                    <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>$</span>
                        <Input value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Annual profit target *" style={{ width: 180, fontFamily: F.mono }} />
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                        <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>Year:</span>
                        <Input value={newKr.krYear || ""} onChange={e => setNewKr(p => ({ ...p, krYear: e.target.value }))} placeholder={String(new Date().getFullYear())} style={{ width: 90, fontFamily: F.mono }} />
                      </div>
                    </div>
                  ) : (
                  <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap", alignItems: "center" }}>
                    {!isTracker && !newKr.useMonthlyTargets && (<>
                      <select value={newKr.operator || ">="} onChange={e => setNewKr(p => ({ ...p, operator: e.target.value }))} style={{ ...sel, width: 80 }}>
                        <option value=">=">≥</option>
                        <option value="<=">≤</option>
                        <option value="=">=</option>
                      </select>
                      <Input value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Target *" style={{ width: 110 }} />
                    </>)}
                    <Input value={newKr.unit} onChange={e => setNewKr(p => ({ ...p, unit: e.target.value }))} placeholder="Unit (e.g. $, Days)" style={{ width: 150 }} />
                    <Input value={newKr.dataSource} onChange={e => setNewKr(p => ({ ...p, dataSource: e.target.value }))} placeholder="Data source" style={{ flex: 1, minWidth: 140 }} />
                    {isStandard && (
                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                        <input type="checkbox" checked={!!newKr.useMonthlyTargets} onChange={e => setNewKr(p => ({ ...p, useMonthlyTargets: e.target.checked }))} style={{ accentColor: T.brand }} />
                        Monthly targets
                      </label>
                    )}
                    {isTracker && (
                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                        <input type="checkbox" checked={!!newKr.disallowZero} onChange={e => setNewKr(p => ({ ...p, disallowZero: e.target.checked }))} style={{ accentColor: T.bad }} />
                        Block zero submission
                      </label>
                    )}
                  </div>
                  )}
                  {isStandard && newKr.useMonthlyTargets && (
                    <div style={{ marginBottom: 8 }}>
                      <Input value={newKr.dreamTarget} onChange={e => setNewKr(p => ({ ...p, dreamTarget: e.target.value }))} placeholder="Dream / annual target (optional)" style={{ width: 260 }} />
                    </div>
                  )}
                  {isMgrFill && (
                    <div style={{ fontSize: 12, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "7px 12px", marginBottom: 8 }}>
                      Manager Fill — member will not receive a check-in; manager assesses yes/no + value in their portal.
                    </div>
                  )}
                  {isProjectProfit && (
                    <div style={{ fontSize: 12, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "7px 12px", marginBottom: 8 }}>
                      Project Profit — actual is auto-calculated from this user's completed projects in the target year. No check-ins required.
                    </div>
                  )}
                  <div style={{ display: "flex", gap: 8 }}>
                    <Btn primary small onClick={() => addKr(deptId, teamId)}>✓ Add Key Result</Btn>
                    <Btn small onClick={() => { setAddTarget(null); setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", krYear: "", monthlyTargets: {}, disallowZero: false }); }}>Cancel</Btn>
                  </div>
                </div>
              );
            };
            const renderSetupSection = (krs, deptId, teamId) => (
              <Card style={{ overflow: "hidden", marginBottom: 0 }}>
                <div style={{ overflowX: "auto" }}><div style={{ minWidth: 742 }}>
                {krs.length > 0 && (<>
                  <div style={{ display: "grid", gridTemplateColumns: KCOL_S, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                    <span>ID</span><span>Key Result</span><span>Type</span><span>Period</span><span>Target</span><span>Unit</span><span>Data Source</span><span />
                  </div>
                  {renderSetupRows(krs, deptId, teamId)}
                </>)}
                {krs.length === 0 && addTarget !== (teamId || `dept-${deptId}`) && (
                  <div style={{ fontSize: 13, color: T.textMuted, padding: "12px 16px" }}>No KRs yet — add your first key result below.</div>
                )}
                {renderAddRow(deptId, teamId)}
                </div></div>
              </Card>
            );
            const totalKrs = (d.krs || []).length + (d.teams || []).reduce((s, t) => s + (t.krs || []).length, 0);
            return (<>
              <Header title="Set Up OKRs" sub={`${d.name} · configure key results`} />
              <Pane>
                <div style={{ display: "flex", gap: 4, marginBottom: 12, flexWrap: "wrap" }}>
                  {PERIODS.map(p => <Btn key={p.id} primary={adminOkrPeriod === p.id} small onClick={() => setAdminOkrPeriod(p.id)}>{p.label}</Btn>)}
                </div>
                <DeptNav />
                <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                  <Metric label="Total KRs" value={totalKrs} />
                  <Metric label="Teams" value={(d.teams || []).length} />
                </div>
                <div style={{ display: "flex", alignItems: "center", marginBottom: 8 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, flex: 1 }}>{d.name} — Department KRs<PChip /></div>
                  {(d.krs || []).length > 0 && <Btn primary small onClick={() => doDeptSync(d.id)}>⟳ Sync to All Members</Btn>}
                </div>
                {renderSetupSection(filterSetupP(d.krs || []), d.id, null)}
                {(d.teams || []).length > 0 && (<>
                  <SectionLabel>Team Key Results<PChip /></SectionLabel>
                  {(d.teams || []).map(t => {
                    const tKrs = filterSetupP(t.krs || []);
                    return (
                      <div key={t.id} style={{ marginBottom: 16 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8, padding: "0 2px" }}>
                          <span style={{ fontSize: 14, fontWeight: 700, flex: 1 }}>{t.name}{t.lead && <span style={{ fontSize: 12, color: T.textMuted, fontWeight: 400, marginLeft: 8 }}>Lead: {t.lead}</span>}</span>
                          {(t.krs || []).length > 0 && <Btn small onClick={() => doSync(d.id, t.id)}>⟳ Sync</Btn>}
                        </div>
                        {renderSetupSection(tKrs, d.id, t.id)}
                      </div>
                    );
                  })}
                </>)}
              </Pane>
            </>);
          }
          return <DeptMgmtPage depts={depts} users={users} memberData={memberData} okrSubmissions={okrSubmissions} dispatch={dispatch} onViewKrs={id => { setAdminSelDept(id); setExpandedMonthlyKr(null); }} />;
        })()}

        {false && false && (() => {
          const dept = depts.find(d => d.id === selDept);
          if (!dept) return null;
          const COLS_DEF = [
            { key: "id",         label: "ID" },
            { key: "label",      label: "Key Result" },
            { key: "operator",   label: "Op" },
            { key: "period",     label: "Period" },
            { key: "target",     label: "Performance Target" },
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
                          if (key === "label") return <div key="label"><span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>{kr.type === "tracker" && <><span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", marginTop: 2, display: "inline-block" }}>Tracker · does not affect rate</span><label style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4, cursor: "pointer", userSelect: "none" }}><input type="checkbox" checked={kr.showInOverview !== false} onChange={e => onTeamChange(kr.id, "showInOverview", e.target.checked)} style={{ accentColor: "#7c3aed" }} /><span style={{ fontSize: 10, color: "#7c3aed" }}>Show in portals' OKR Overview</span></label></>}{isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", marginTop: 2, display: "inline-block" }}>Monthly Breakdown</span>}{kr.type !== "tracker" && <label style={{ display: "flex", alignItems: "center", gap: 4, marginTop: 4, cursor: "pointer", userSelect: "none" }}><input type="checkbox" checked={!!kr.autoApprove} onChange={e => onTeamChange(kr.id, "autoApprove", e.target.checked)} style={{ accentColor: T.ok }} /><span style={{ fontSize: 10, color: kr.autoApprove ? T.ok : T.textDim }}>Auto-approve ✓ Yes</span></label>}</div>;
                          if (key === "operator") return <span key="operator">{opSelect(kr.operator || ">=", e => onTeamChange(kr.id, "operator", e.target.value))}</span>;
                          if (key === "period") return <span key="period"><select value={kr.period || "monthly"} onChange={e => onTeamChange(kr.id, "period", e.target.value)} style={{ width: "100%", padding: "5px 4px", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontFamily: F.body }}><option value="daily">Daily</option><option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="biannual">Biannual</option><option value="annual">Annual</option></select></span>;
                          if (key === "target") return kr.type === "tracker" ? <span key="target" style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : isMonthly ? <span key="target" style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: T.brand }}>{fmt(curTarget)} <span style={{ color: T.textDim }}>this mo.</span></span> : <NumInput key="target" value={kr.target} onChange={n => onTeamChange(kr.id, "target", n)} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />;
                          if (key === "actual") return isMonthly ? <NumInput key="actual" value={curActual} onChange={n => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: curKey, field: "actual", value: n }); if (teamId) triggerSyncPrompt(deptId, teamId); }} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} /> : <span key="actual" style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.actual)}</span>;
                          if (key === "unit") return <Input key="unit" value={kr.unit || ""} onChange={e => onTeamChange(kr.id, "unit", e.target.value)} placeholder="e.g. %, students" style={{ padding: "5px 8px", fontSize: 13 }} />;
                          if (key === "dataSource") return <Input key="dataSource" value={kr.dataSource || ""} onChange={e => onTeamChange(kr.id, "dataSource", e.target.value)} placeholder="e.g. CRM, Manual" style={{ padding: "5px 8px", fontSize: 13 }} />;
                          return null;
                        })}
                        {customCols.map(col => <Input key={col.id} value={(kr.extras || {})[col.id] || ""} onChange={e => onTeamChange(kr.id, "extras", { ...(kr.extras || {}), [col.id]: e.target.value })} placeholder="—" style={{ padding: "5px 8px", fontSize: 13 }} />)}
                        <div style={{ display: "flex", gap: 4 }}>
                          {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="Monthly breakdown — all 12 months" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 12, fontWeight: 700 }}>📅</button>}
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
                                  {fyMs.map(({ key }) => { const isCur = key === curKey; const t = kr.monthlyTargets[key] || 0; return <td key={key} style={{ padding: "3px 3px", background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}` }}><NumInput value={t} onChange={n => dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "target", value: n })} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} /></td>; })}
                                  <td style={{ padding: "4px 10px", textAlign: "right", fontFamily: F.mono, fontWeight: 700, fontSize: 13, background: T.surface, borderBottom: `1px solid ${T.border}` }}>{fmt(annSumTarget)}</td>
                                  <td style={{ padding: "5px 8px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, borderLeft: `1px solid ${T.okBorder}` }}>
                                    <NumInput value={annDream} onChange={n => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "annualTarget", value: n })} style={{ padding: "4px 6px", fontSize: 13, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box", fontWeight: 700, background: T.surface, border: `1px solid ${T.okBorder}`, borderRadius: 4 }} />
                                    {kr.unit && <div style={{ fontSize: 10, color: T.ok, textAlign: "center", marginTop: 2 }}>{kr.unit}</div>}
                                  </td>
                                </tr>
                                <tr>
                                  <td style={{ padding: "4px 10px", fontWeight: 700, fontSize: 12, color: T.text, background: T.surface, borderBottom: `1px solid ${T.border}` }}>Actual</td>
                                  {fyMs.map(({ key }) => { const isCur = key === curKey; const a = (kr.monthlyActuals || {})[key] || 0; return <td key={key} style={{ padding: "3px 3px", background: isCur ? T.brandDim : "transparent", borderBottom: `1px solid ${T.border}` }}><NumInput value={a} onChange={n => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "actual", value: n }); if (teamId) triggerSyncPrompt(deptId, teamId); }} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} /></td>; })}
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
                      {sectionPeriod === "monthly" && newKr.krType !== "tracker" && newKr.krType !== "progress" && (
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
                      <div style={{ padding: "8px 16px", background: T.brandDim, borderTop: `1px solid ${T.brandBorder}`, display: "flex", gap: 24, flexWrap: "wrap" }}>
                        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, cursor: "pointer", color: "#7c3aed", fontWeight: 600 }}>
                          <input type="checkbox" checked={newKr.krType === "tracker"} onChange={e => setNewKr(p => ({ ...p, krType: e.target.checked ? "tracker" : "", useMonthlyTargets: false, disallowZero: false }))} style={{ accentColor: "#7c3aed" }} />
                          Tracker — record values only, does not affect completion rate
                        </label>
                        {newKr.krType === "tracker" && (
                          <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, cursor: "pointer", color: T.bad, fontWeight: 500 }}>
                            <input type="checkbox" checked={!!newKr.disallowZero} onChange={e => setNewKr(p => ({ ...p, disallowZero: e.target.checked }))} style={{ accentColor: T.bad }} />
                            Block zero submission
                          </label>
                        )}
                        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, cursor: "pointer", color: "#0771e3", fontWeight: 600 }}>
                          <input type="checkbox" checked={newKr.krType === "progress"} onChange={e => setNewKr(p => ({ ...p, krType: e.target.checked ? "progress" : "", useMonthlyTargets: false }))} style={{ accentColor: "#0771e3" }} />
                          Progress — records cumulative progress toward target; affects rate proportionally
                        </label>
                        <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12, cursor: "pointer", color: "#d97706", fontWeight: 600 }}>
                          <input type="checkbox" checked={newKr.krType === "manager-fill"} onChange={e => setNewKr(p => ({ ...p, krType: e.target.checked ? "manager-fill" : "", useMonthlyTargets: false }))} style={{ accentColor: "#d97706" }} />
                          Manager Fill — manager enters actual value; member does not receive a check-in
                        </label>
                      </div>
                    </div>
                  ) : (
                    <div style={{ padding: "10px 16px" }}>
                      <button onClick={() => { setAddTarget(teamId || `dept-${deptId}`); setNewKr({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: sectionPeriod, useMonthlyTargets: false, krType: "", monthlyTargets: {}, disallowZero: false }); }} style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: 6, padding: "8px 14px", cursor: "pointer", color: T.brand, fontSize: 13, fontWeight: 600, width: "100%", fontFamily: F.body }}>+ Add Key Result</button>
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
                                  <div style={{ overflowX: "auto" }}><div style={{ minWidth: 510 }}>
                                  <div style={{ display: "grid", gridTemplateColumns: "1fr 52px 90px 90px 60px 90px 28px", gap: 8, padding: "5px 0 6px", fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", borderBottom: `1px solid ${T.border}` }}>
                                    <span>Key Result</span><span>Op</span><span style={{ textAlign: "right" }}>Performance Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Unit</span><span>Period</span><span />
                                  </div>
                                  {personalKrs.map((kr, i) => {
                                    const isTracker = kr.type === "tracker";
                                    const isMonthly = !!(kr.monthlyTargets);
                                    const mk = currentFYMonthKey();
                                    const mTgt = isMonthly ? (Number(kr.monthlyTargets[mk]) || 0) : null;
                                    const mAct = isMonthly ? ((kr.monthlyActuals || {})[mk] ?? null) : null;
                                    const pct = isTracker ? 0 : krCompletion(kr); const st2 = getStatus(pct);
                                    return (
                                      <Fragment key={kr.id}>
                                      <div style={{ display: "grid", gridTemplateColumns: "1fr 52px 90px 90px 60px 90px 28px", gap: 8, padding: "6px 0", alignItems: "center", borderBottom: isMonthly && !isTracker ? "none" : `1px solid ${T.border}`, background: i % 2 ? T.raised : "transparent", fontSize: 13 }}>
                                        <div>
                                          <span>{kr.label}</span>
                                          {isTracker ? <span style={{ marginLeft: 8, fontSize: 10, fontWeight: 700, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 6px" }}>Tracker</span> : <span style={{ marginLeft: 8, fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[st2].color }}>{pct.toFixed(0)}%</span>}
                                          {isMonthly && !isTracker && <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 700, color: "#0369a1", background: "#e0f2fe", border: "1px solid #7dd3fc", borderRadius: 8, padding: "1px 5px" }}>Monthly</span>}
                                        </div>
                                        <select value={kr.operator || ">="} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "operator", value: e.target.value })} style={{ padding: "3px 4px", fontSize: 12, borderRadius: 4, border: `1px solid ${T.border}`, background: T.card, color: T.text, fontFamily: F.mono, cursor: "pointer" }}>
                                          <option value=">=">≥</option><option value="<=">≤</option><option value="=">=</option><option value=">">{">"}</option><option value="<">{"<"}</option>
                                        </select>
                                        {isTracker ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span>
                                          : isMonthly ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: T.brand }}>{fmt(mTgt)} <span style={{ color: T.textDim }}>this mo.</span></span>
                                          : <NumInput value={kr.target || 0} onChange={n => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "target", value: n })} style={{ textAlign: "right", padding: "3px 6px", fontFamily: F.mono, fontSize: 13 }} />}
                                        {isTracker ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted, fontSize: 12 }}>{fmt(kr.actual)}</span>
                                          : isMonthly ? <NumInput value={mAct} onChange={n => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "monthlyActuals", value: { ...(kr.monthlyActuals || {}), [mk]: n } })} style={{ textAlign: "right", padding: "3px 6px", fontFamily: F.mono, fontSize: 13 }} />
                                          : <NumInput value={kr.actual || 0} onChange={n => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "actual", value: n })} style={{ textAlign: "right", padding: "3px 6px", fontFamily: F.mono, fontSize: 13 }} />}
                                        <span style={{ fontSize: 12, color: T.textMuted }}>{kr.unit || "—"}</span>
                                        <select value={kr.period || "monthly"} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: member.id, krId: kr.id, field: "period", value: e.target.value })} style={{ padding: "3px 4px", fontSize: 12, borderRadius: 4, border: `1px solid ${T.border}`, background: T.card, color: T.text, cursor: "pointer" }}>
                                          <option value="weekly">Weekly</option><option value="monthly">Monthly</option><option value="quarterly">Quarterly</option><option value="biannual">Bi-Annual</option><option value="annual">Annual</option>
                                        </select>
                                        <button onClick={() => dispatch({ type: "REMOVE_MEMBER_KR", memberId: member.id, krId: kr.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 4, padding: "2px 6px", cursor: "pointer", color: T.bad, fontSize: 11 }}>✕</button>
                                      </div>
                                      {isMonthly && !isTracker && (
                                        <div style={{ padding: "4px 10px 10px 16px", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}` }}>
                                          <div style={{ display: "grid", gridTemplateColumns: "76px 100px 100px 56px 1fr", gap: 6, padding: "4px 6px", fontSize: 10, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em", borderBottom: `1px solid ${T.border}`, marginBottom: 2 }}>
                                            <span>Month</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span />
                                          </div>
                                          {getFYMonths().map(({ key, label }) => {
                                            const mt = Number(kr.monthlyTargets[key]) || 0;
                                            const ma = (kr.monthlyActuals || {})[key];
                                            const mp = ma != null ? (mt === 0 ? 100 : Math.min((ma / mt) * 100, 100)) : null;
                                            const ms = mp != null ? getStatus(mp) : null;
                                            const isCurr = key === mk;
                                            return (
                                              <div key={key} style={{ display: "grid", gridTemplateColumns: "76px 100px 100px 56px 1fr", gap: 6, padding: "2px 6px", alignItems: "center", fontSize: 12, borderRadius: 4, background: isCurr ? T.brandDim : "transparent" }}>
                                                <span style={{ fontWeight: isCurr ? 700 : 400, color: isCurr ? T.brand : T.textMuted }}>{label}</span>
                                                <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(mt)}{kr.unit ? ` ${kr.unit}` : ""}</span>
                                                <span style={{ textAlign: "right", fontFamily: F.mono, color: ma != null ? T.text : T.textDim }}>{ma != null ? fmt(ma) : "—"}</span>
                                                <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: ms ? STATUS_THEME[ms].color : T.textDim }}>{mp != null ? mp.toFixed(0) + "%" : "—"}</span>
                                                <Bar value={mp ?? 0} status={ms || "behind"} h={3} />
                                              </div>
                                            );
                                          })}
                                        </div>
                                      )}
                                      </Fragment>
                                    );
                                  })}
                                  </div></div>
                                </div>
                              )}
                              {addPersonalKr?.memberId === member.id ? (
                                <div style={{ display: "flex", gap: 8, alignItems: "center", padding: "8px 10px", background: T.brandDim, borderRadius: 7 }}>
                                  <Input value={addPersonalKr.label} onChange={e => setAddPersonalKr(p => ({...p, label: e.target.value}))} placeholder="KR description *" style={{ flex: 1, padding: "4px 8px", fontSize: 13 }} />
                                  <select value={addPersonalKr.operator} onChange={e => setAddPersonalKr(p => ({...p, operator: e.target.value}))} style={{ padding: "4px 5px", fontSize: 13, borderRadius: 5, border: `1px solid ${T.border}`, background: T.card, color: T.text, fontFamily: F.mono, cursor: "pointer" }}>
                                    <option value=">=">≥</option><option value="<=">≤</option><option value="=">=</option><option value=">">{">"}</option><option value="<">{"<"}</option>
                                  </select>
                                  <Input value={addPersonalKr.target} onChange={e => setAddPersonalKr(p => ({...p, target: e.target.value}))} placeholder="Target" style={{ width: 80, textAlign: "right", padding: "4px 6px", fontFamily: F.mono, fontSize: 13 }} />
                                  <Input value={addPersonalKr.unit} onChange={e => setAddPersonalKr(p => ({...p, unit: e.target.value}))} placeholder="Unit" style={{ width: 70, padding: "4px 6px", fontSize: 13 }} />
                                  <select value={addPersonalKr.period} onChange={e => setAddPersonalKr(p => ({...p, period: e.target.value}))} style={{ padding: "4px 6px", fontSize: 13, borderRadius: 5, border: `1px solid ${T.border}`, background: T.card, color: T.text, cursor: "pointer" }}>
                                    <option value="weekly">Weekly</option>
                                    <option value="monthly">Monthly</option>
                                    <option value="quarterly">Quarterly</option>
                                    <option value="biannual">Bi-Annual</option>
                                    <option value="annual">Annual</option>
                                  </select>
                                  <button onClick={() => { if (!addPersonalKr.label.trim()) return; dispatch({ type: "ADD_MEMBER_KR", memberId: member.id, kr: { id: `P${Date.now().toString(36).slice(-4).toUpperCase()}`, label: addPersonalKr.label.trim(), target: Number(addPersonalKr.target) || 0, actual: null, unit: addPersonalKr.unit.trim(), operator: addPersonalKr.operator, period: addPersonalKr.period } }); setAddPersonalKr(null); }} style={{ background: T.brand, border: "none", borderRadius: 5, padding: "4px 10px", cursor: "pointer", color: "#fff", fontSize: 12, fontWeight: 700 }}>✓</button>
                                  <button onClick={() => setAddPersonalKr(null)} style={{ background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 8px", cursor: "pointer", color: T.text, fontSize: 12 }}>✕</button>
                                </div>
                              ) : (
                                <button onClick={() => setAddPersonalKr({ memberId: member.id, label: "", target: "", unit: "", operator: ">=", period: "monthly" })}
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
        })()}

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
            <Header title="OKR Check-In Submissions" sub="Staff respond to emailed yes/no KPI check-ins — managers approve in their portal"
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
                <Btn primary onClick={() => setShowCheckinDialog(true)} disabled={sendingCheckin}>
                  {sendingCheckin ? "Sending…" : "📨 Send Check-In"}
                </Btn>
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
                    : `⚠ No submissions created — no KRs found for selected periods. Check that KRs are configured with the correct period.`}
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
                    <span>{pendingMgrSubs.length} manager submission{pendingMgrSubs.length !== 1 ? "s" : ""} awaiting your approval — visible only to admins and same-department managers with "Can approve peers" enabled</span>
                  </div>
                ) : null;
              })()}
              {filtered.length === 0 && <EmptyState text={periodSubs.length === 0 ? `No ${subPeriod} check-ins sent yet. Click "Send ${subPeriod.charAt(0).toUpperCase()+subPeriod.slice(1)} Check-In" to generate and email them.` : "No submissions match your filter."} />}
              {(() => {
                const order = [];
                const groups = {};
                filtered.forEach(s => {
                  if (!groups[s.memberId]) { groups[s.memberId] = []; order.push(s.memberId); }
                  groups[s.memberId].push(s);
                });
                return order.map(memberId => {
                  const subs = groups[memberId];
                  const mem = users.find(u => u.id === memberId);
                  const dept = depts.find(d => d.id === subs[0]?.deptId);
                  const pendingCount = subs.filter(s => s.answer !== null && s.approval === "pending").length;
                  return (
                    <Card key={memberId} style={{ marginBottom: 10, overflow: "hidden" }}>
                      {/* Person header */}
                      <div style={{ padding: "11px 18px", background: T.raised, borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                        <Avatar letters={mem?.av || "?"} size={30} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" }}>
                            <span style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{mem?.name || subs[0]?.memberName || "Unknown"}</span>
                            {mem?.role === "manager" && <span style={{ fontSize: 10, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Manager</span>}
                            {dept && <span style={{ fontSize: 11, color: T.textMuted, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 5, padding: "1px 6px" }}>{dept.name}</span>}
                          </div>
                          <div style={{ fontSize: 11, color: T.textDim, marginTop: 1 }}>{subs.length} KR{subs.length !== 1 ? "s" : ""}</div>
                        </div>
                        {pendingCount > 0 && (
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            <span style={{ background: T.warnDim, color: T.warn, border: `1px solid ${T.warnBorder}`, borderRadius: 6, padding: "2px 9px", fontSize: 11, fontWeight: 700 }}>{pendingCount} pending</span>
                            <Btn primary small onClick={() => subs.filter(s => s.answer !== null && s.approval === "pending").forEach(s => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id }))}>Approve All</Btn>
                          </div>
                        )}
                      </div>
                      {/* KR rows */}
                      {subs.map((s, idx) => {
                        const accentColor = s.approval === "approved" ? T.ok : s.answer === "yes" ? T.ok : s.answer === "no" ? T.bad : s.answer === null ? T.warn : T.border;
                        const isLast = idx === subs.length - 1 && rejectOkr?.id !== s.id && editingSub?.id !== s.id && editingApproved?.id !== s.id;
                        return (
                          <div key={s.id} style={{ borderBottom: isLast ? "none" : `1px solid ${T.border}` }}>
                            <div style={{ padding: "11px 18px 11px 21px", borderLeft: `3px solid ${accentColor}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 3 }}>
                                  <span style={{ fontSize: 14, fontWeight: 600, color: T.text }}>{s.krLabel}</span>
                                  {s.krType === "tracker" && <span style={{ fontSize: 10, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Tracker · no rate impact</span>}
                                  {s.krType === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Progress · affects rate</span>}
                                </div>
                                <div style={{ display: "flex", alignItems: "baseline", gap: 6, marginBottom: 3 }}>
                                  {s.krType !== "tracker" && s.krType !== "progress" && <span style={{ fontSize: 13, color: T.textMuted, fontFamily: F.mono, fontWeight: 600, alignSelf: "center" }}>{s.krOperator || ">="}</span>}
                                  <span style={{ fontSize: 20, fontWeight: 800, fontFamily: F.mono, color: T.text, lineHeight: 1 }}>{s.krTarget ?? "—"}</span>
                                  {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                                  <span style={{ fontSize: 11, color: T.textDim }}>{s.krType === "progress" ? "target" : "performance target"}</span>
                                  {s.answer === "no" && s.actualValue != null && <><span style={{ fontSize: 11, color: T.textDim, marginLeft: 8 }}>·</span><span style={{ fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: T.bad, lineHeight: 1, marginLeft: 8 }}>{s.actualValue}</span><span style={{ fontSize: 11, color: T.bad }}>actual</span></>}
                                  {s.answer === "yes" && <><span style={{ fontSize: 11, color: T.textDim, marginLeft: 8 }}>·</span><span style={{ fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: T.ok, lineHeight: 1, marginLeft: 8 }}>{s.actualValue ?? s.krTarget}</span><span style={{ fontSize: 11, color: T.ok }}>actual</span>{s.actualValue == null && <span style={{ fontSize: 10, color: T.textDim, marginLeft: 3 }}>(assumed)</span>}</>}
                                  {s.krType === "progress" && s.actualValue != null && <><span style={{ fontSize: 11, color: T.textDim, marginLeft: 8 }}>·</span><span style={{ fontSize: 18, fontWeight: 800, fontFamily: F.mono, color: T.brand, lineHeight: 1, marginLeft: 8 }}>{s.actualValue}</span><span style={{ fontSize: 11, color: T.brand }}>recorded</span></>}
                                </div>
                                <div style={{ fontSize: 12, color: T.textMuted }}>{periodDisplayLabel(s.period, s.periodKey)} · Sent: {s.sentAt?.slice(0,10) || "—"}</div>
                                {s.answer === "no" && s.reason && <div style={{ fontSize: 12, color: T.bad, marginTop: 3, fontStyle: "italic" }}>Reason: {s.reason}</div>}
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                                {s.answer === null
                                  ? <span style={{ fontSize: 12, color: T.textMuted, background: T.raised, borderRadius: 6, padding: "3px 8px" }}>{(s.krType === "tracker" || s.krType === "progress") ? "Awaiting record" : "Awaiting answer"}</span>
                                  : s.krType === "tracker"
                                    ? <span style={{ fontSize: 12, fontWeight: 700, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 6, padding: "3px 8px" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                                    : s.krType === "progress"
                                    ? <span style={{ fontSize: 12, fontWeight: 700, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "3px 8px" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}{s.krTarget ? ` (${Math.min(Math.round((Number(s.actualValue || 0) / Number(s.krTarget)) * 100), 100)}%)` : ""}</span>
                                    : <div style={{ textAlign: "right" }}>
                                        <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad, background: s.answer === "yes" ? T.okDim : T.badDim, border: `1px solid ${s.answer === "yes" ? T.okBorder : T.badBorder}`, borderRadius: 6, padding: "3px 8px" }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>
                                        {s.actualValue != null && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</div>}
                                      </div>}
                                {s.answer !== null && s.approval === "pending"
                                  ? <div style={{ display: "flex", gap: 6 }}>
                                      <Btn small onClick={() => { setEditingSub({ id: s.id, answer: s.answer, actual: s.actualValue != null ? String(s.actualValue) : "" }); setRejectOkr(null); }}>✎ Edit</Btn>
                                      <Btn danger small onClick={() => { setRejectOkr({ id: s.id, actual: "" }); setEditingSub(null); }}>Reject</Btn>
                                      <Btn primary small onClick={() => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id })}>Approve</Btn>
                                    </div>
                                  : s.approval !== "pending" && <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                      <Tag type={s.approval === "approved" ? "approved" : "rejected"} label={s.approval === "approved" ? "Approved" : "Rejected"} small />
                                      {s.approvedBy && s.approvedBy !== "auto" && <span style={{ fontSize: 11, color: T.textMuted }}>by {users?.find(u => u.id === s.approvedBy)?.name || "Admin"}</span>}
                                      <Btn small onClick={() => { setEditingApproved({ id: s.id, actual: s.actualValue != null ? String(s.actualValue) : "", answer: s.answer }); setEditingSub(null); setRejectOkr(null); }}>✎</Btn>
                                    </div>}
                                <button onClick={() => dispatch({ type: "REMOVE_OKR_SUBMISSION", id: s.id })} title="Delete" style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 14, lineHeight: 1, padding: "2px 4px" }}>✕</button>
                              </div>
                            </div>
                            {rejectOkr?.id === s.id && (
                              <div style={{ margin: "0 18px 10px 21px", padding: "10px 12px", background: T.badDim, borderRadius: 7, border: `1px solid ${T.badBorder}` }}>
                                <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 6 }}>Enter actual value for rejection</div>
                                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                                  <Input value={rejectOkr.actual} onChange={e => setRejectOkr(p => ({ ...p, actual: e.target.value }))} placeholder="Actual value" style={{ width: 120, textAlign: "right", fontFamily: F.mono }} />
                                  {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                  {s.krType === "progress" ? <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span> : <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                </div>
                                <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                  <Btn small onClick={() => setRejectOkr(null)}>Cancel</Btn>
                                  <Btn danger small onClick={() => { dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "rejected", approvedBy: user.id, actualValue: Number(rejectOkr.actual) || 0 }); setRejectOkr(null); }}>Confirm Reject</Btn>
                                </div>
                              </div>
                            )}
                            {editingSub?.id === s.id && (
                              <div style={{ margin: "0 18px 10px 21px", padding: "12px 14px", background: T.raised, borderRadius: 8, border: `1px solid ${T.border}` }}>
                                <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 10 }}>Edit Submission</div>
                                {s.krType !== "tracker" && s.krType !== "progress" ? (<>
                                  <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                                    <button onClick={() => setEditingSub(p => ({ ...p, answer: "yes", actual: String(s.krTarget ?? "") }))}
                                      style={{ background: editingSub.answer === "yes" ? T.okDim : T.surface, border: `1px solid ${editingSub.answer === "yes" ? T.okBorder : T.border}`, borderRadius: 7, padding: "7px 18px", cursor: "pointer", color: editingSub.answer === "yes" ? T.ok : T.textMuted, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>✓ Yes</button>
                                    <button onClick={() => setEditingSub(p => ({ ...p, answer: "no" }))}
                                      style={{ background: editingSub.answer === "no" ? T.badDim : T.surface, border: `1px solid ${editingSub.answer === "no" ? T.badBorder : T.border}`, borderRadius: 7, padding: "7px 18px", cursor: "pointer", color: editingSub.answer === "no" ? T.bad : T.textMuted, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>✗ No</button>
                                  </div>
                                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                                    <span style={{ fontSize: 12, color: T.textMuted }}>Actual value:</span>
                                    <Input value={editingSub.actual} onChange={e => setEditingSub(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                                    {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                    <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                                  </div>
                                </>) : (
                                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                                    <span style={{ fontSize: 12, color: T.textMuted }}>Recorded value:</span>
                                    <Input value={editingSub.actual} onChange={e => setEditingSub(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                                    {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                    {s.krType === "progress" && <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                  </div>
                                )}
                                {s.krType !== "tracker" && s.krType !== "progress" && editingSub.actual !== "" && editingSub.answer && ((editingSub.answer === "yes" && meetsTarget(editingSub.actual, s.krOperator, s.krTarget) === false) || (editingSub.answer === "no" && meetsTarget(editingSub.actual, s.krOperator, s.krTarget) === true)) && (
                                  <div style={{ fontSize: 12, color: T.bad, fontWeight: 600, marginBottom: 8 }}>⚠ Actual ({editingSub.actual}{s.krUnit ? " " + s.krUnit : ""}) {meetsTarget(editingSub.actual, s.krOperator, s.krTarget) ? "meets" : "doesn't meet"} target ({s.krOperator || ">="} {s.krTarget}{s.krUnit ? " " + s.krUnit : ""}) — answer should be {meetsTarget(editingSub.actual, s.krOperator, s.krTarget) ? "Yes" : "No"}</div>
                                )}
                                <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                  <Btn small onClick={() => setEditingSub(null)}>Cancel</Btn>
                                  <Btn primary small disabled={s.krType !== "tracker" && s.krType !== "progress" && editingSub.actual !== "" && editingSub.answer && ((editingSub.answer === "yes" && meetsTarget(editingSub.actual, s.krOperator, s.krTarget) === false) || (editingSub.answer === "no" && meetsTarget(editingSub.actual, s.krOperator, s.krTarget) === true))} onClick={() => { const newAnswer = (s.krType === "tracker" || s.krType === "progress") ? "submitted" : editingSub.answer; const newActual = Number(editingSub.actual) || 0; dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: newAnswer, actualValue: newActual }); setEditingSub(null); }}>Save Changes</Btn>
                                </div>
                              </div>
                            )}
                            {editingApproved?.id === s.id && (
                              <div style={{ margin: "0 18px 10px 21px", padding: "10px 12px", background: T.raised, borderRadius: 7, border: `1px solid ${T.border}` }}>
                                <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 8 }}>Edit Submission — {s.approval === "approved" ? "Approved" : "Rejected"}</div>
                                {s.krType !== "tracker" && s.krType !== "progress" && (
                                  <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                                    <button onClick={() => setEditingApproved(p => ({ ...p, answer: "yes" }))} style={{ background: editingApproved.answer === "yes" ? T.okDim : T.surface, border: `1px solid ${editingApproved.answer === "yes" ? T.okBorder : T.border}`, borderRadius: 7, padding: "5px 16px", cursor: "pointer", color: editingApproved.answer === "yes" ? T.ok : T.textMuted, fontSize: 13, fontWeight: 700, fontFamily: F.body }}>✓ Yes</button>
                                    <button onClick={() => setEditingApproved(p => ({ ...p, answer: "no" }))} style={{ background: editingApproved.answer === "no" ? T.badDim : T.surface, border: `1px solid ${editingApproved.answer === "no" ? T.badBorder : T.border}`, borderRadius: 7, padding: "5px 16px", cursor: "pointer", color: editingApproved.answer === "no" ? T.bad : T.textMuted, fontSize: 13, fontWeight: 700, fontFamily: F.body }}>✗ No</button>
                                  </div>
                                )}
                                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                                  <span style={{ fontSize: 12, color: T.textMuted }}>Actual value:</span>
                                  <Input value={editingApproved.actual} onChange={e => setEditingApproved(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                                  {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                  {s.krType !== "tracker" && s.krType !== "progress" && <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                  {s.krType === "progress" && <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                </div>
                                {s.krType !== "tracker" && s.krType !== "progress" && editingApproved.actual !== "" && editingApproved.answer && ((editingApproved.answer === "yes" && meetsTarget(editingApproved.actual, s.krOperator, s.krTarget) === false) || (editingApproved.answer === "no" && meetsTarget(editingApproved.actual, s.krOperator, s.krTarget) === true)) && (
                                  <div style={{ fontSize: 12, color: T.bad, fontWeight: 600, marginBottom: 8 }}>⚠ Actual ({editingApproved.actual}{s.krUnit ? " " + s.krUnit : ""}) {meetsTarget(editingApproved.actual, s.krOperator, s.krTarget) ? "meets" : "doesn't meet"} target ({s.krOperator || ">="} {s.krTarget}{s.krUnit ? " " + s.krUnit : ""}) — answer should be {meetsTarget(editingApproved.actual, s.krOperator, s.krTarget) ? "Yes" : "No"}</div>
                                )}
                                <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                  <Btn small onClick={() => setEditingApproved(null)}>Cancel</Btn>
                                  <Btn primary small disabled={s.krType !== "tracker" && s.krType !== "progress" && editingApproved.actual !== "" && editingApproved.answer && ((editingApproved.answer === "yes" && meetsTarget(editingApproved.actual, s.krOperator, s.krTarget) === false) || (editingApproved.answer === "no" && meetsTarget(editingApproved.actual, s.krOperator, s.krTarget) === true))} onClick={() => { dispatch({ type: "EDIT_APPROVED_SUBMISSION", id: s.id, actualValue: Number(editingApproved.actual) || 0, answer: editingApproved.answer }); setEditingApproved(null); }}>Save</Btn>
                                </div>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </Card>
                  );
                });
              })()}
              {/* Email Send Log */}
              {emailLogs.length > 0 && (
                <div style={{ marginTop: 32 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 14, paddingBottom: 10, borderBottom: `2px solid ${T.border}` }}>
                    <div style={{ fontSize: 16, fontWeight: 800, color: T.text }}>Email Send Log</div>
                    <div style={{ fontSize: 12, color: T.textMuted }}>{emailLogs.length} send event{emailLogs.length !== 1 ? "s" : ""} (last 100)</div>
                  </div>
                  <div style={{ overflowX: "auto" }}><div style={{ minWidth: 640 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "160px 80px 1fr 1fr 70px 70px 60px", gap: 8, padding: "5px 0 6px", fontSize: 11, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.07em", borderBottom: `1px solid ${T.border}` }}>
                    <span>Date / Time</span><span>Period</span><span>Date Range</span><span>Scope</span><span style={{ textAlign: "right" }}>Emails</span><span style={{ textAlign: "right" }}>Created</span><span style={{ textAlign: "right" }}>Fails</span>
                  </div>
                  {emailLogs.map((log, i) => {
                    const isOpen = expandedLog === log.id;
                    const PERIOD_LABELS = { daily: "Daily", weekly: "Weekly", monthly: "Monthly", quarterly: "Quarterly", biannual: "Bi-Annual", annual: "Annual" };
                    return (
                      <div key={log.id} style={{ borderBottom: `1px solid ${T.border}` }}>
                        <div onClick={() => setExpandedLog(isOpen ? null : log.id)} style={{ display: "grid", gridTemplateColumns: "160px 80px 1fr 1fr 70px 70px 60px", gap: 8, padding: "8px 0", alignItems: "center", fontSize: 13, cursor: "pointer", background: i % 2 ? T.raised : "transparent", userSelect: "none" }}>
                          <span style={{ fontFamily: F.mono, fontSize: 12 }}>{new Date(log.sentAt).toLocaleString("en-AU", { day: "2-digit", month: "short", hour: "2-digit", minute: "2-digit" })}</span>
                          <span style={{ fontWeight: 600 }}>{log.periods && log.periods.length > 1 ? log.periods.map(p => PERIOD_LABELS[p] || p).join(" & ") : (PERIOD_LABELS[log.period] || log.period)}</span>
                          <span style={{ color: T.textMuted, fontSize: 12 }}>{log.dateRange || log.periodKey}</span>
                          <span style={{ color: T.textMuted, fontSize: 12 }}>{scopeLabel(log.scope || {})}</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono }}>{log.recipientCount}</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono }}>{log.submissionsCreated}</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: log.failureCount > 0 ? T.bad : T.ok }}>{log.failureCount > 0 ? log.failureCount : "—"}</span>
                        </div>
                        {isOpen && (
                          <div style={{ padding: "6px 14px 12px", background: T.surface, borderRadius: 6, margin: "0 0 4px" }}>
                            {(log.recipients || []).length === 0
                              ? <div style={{ fontSize: 12, color: T.textMuted }}>No emails were sent (no valid email addresses in scope).</div>
                              : (log.recipients || []).map((r, j) => {
                                const rKey = `${log.id}:${r.email}`;
                                const isSending = resendingEmail === rKey;
                                return (
                                  <div key={j} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 13, borderBottom: `1px solid ${T.border}` }}>
                                    <span style={{ width: 14, textAlign: "center", color: r.success ? T.ok : T.bad, fontWeight: 700 }}>{r.success ? "✓" : "✗"}</span>
                                    <span style={{ flex: 1 }}>{r.name}</span>
                                    <span style={{ color: T.textMuted, fontSize: 12 }}>{r.email}</span>
                                    <span style={{ color: T.textMuted, fontSize: 12 }}>{r.krCount} KR{r.krCount !== 1 ? "s" : ""}</span>
                                    {!r.success && <span style={{ color: T.bad, fontSize: 12, flex: 1 }}>{r.reason}</span>}
                                    {!r.success && (
                                      <button disabled={!!resendingEmail} onClick={() => resendEmail(log, r)}
                                        style={{ background: isSending ? T.raised : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "2px 10px", cursor: resendingEmail ? "not-allowed" : "pointer", color: T.brand, fontSize: 12, fontWeight: 700, opacity: resendingEmail && !isSending ? 0.5 : 1 }}>
                                        {isSending ? "Sending…" : "Resend"}
                                      </button>
                                    )}
                                  </div>
                                );
                              })
                            }
                          </div>
                        )}
                      </div>
                    );
                  })}
                  </div></div>
                </div>
              )}
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
                      if (genPeriod.from && s.sentAt && s.sentAt.slice(0, 10) < genPeriod.from) return false;
                      if (genPeriod.to && s.sentAt && s.sentAt.slice(0, 10) > genPeriod.to) return false;
                      return true;
                    });
                    const gSubRate = allAnswered.length > 0 ? Math.round((allAnswered.filter(s => s.answer === "yes").length / allAnswered.length) * 1000) / 10 : 0;
                    const gDeptRanks = depts.map(d => {
                      const members = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === d.id && !u.excludeFromRate);
                      const rates = members.map(u => {
                        const kd = memberData[u.id] || { krs: [] };
                        if (!kd.krs.some(kr => allAnswered.some(s => s.memberId === u.id && s.krId === kr.id))) return null;
                        return calcMemberRate(u.id, kd.krs, allAnswered);
                      }).filter(r => r !== null);
                      const hasData = rates.length > 0;
                      const rate = hasData ? Math.round(rates.reduce((a, b) => a + b, 0) / rates.length * 10) / 10 : 0;
                      return { name: d.name, rate, hasData, status: getStatus(rate) };
                    }).sort((a, b) => b.rate - a.rate);
                    const gActiveDepts = gDeptRanks.filter(d => d.hasData);
                    const gCompRate = gActiveDepts.length ? Math.round(gActiveDepts.reduce((a, d) => a + d.rate, 0) / gActiveDepts.length * 10) / 10 : 0;
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
                <div><span style={{ color: T.textMuted }}>Target met rate ({currentMonth()}): </span><strong style={{ color: STATUS_THEME[getStatus(rptSubRate)].color }}>{rptSubRate}%</strong><span style={{ color: T.textMuted, fontSize: 11, marginLeft: 6 }}>({rptSubs.length} answered)</span></div>
                <div><span style={{ color: T.textMuted }}>Top performers: </span>{rptMembers.filter(m => m.hasData).slice(0, 3).map(m => m.name).join(", ") || "—"}</div>
                <div><span style={{ color: T.textMuted }}>Needs attention: </span>{rptMembers.filter(m => m.hasData && m.status === "red").map(m => m.name).join(", ") || "None"}</div>
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
                      <Tag type={getStatus(r.data.companyRate)} label={`Company: ${Number(r.data.companyRate).toFixed(1)}%`} />
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
                            <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{Number(d.rate).toFixed(1)}%</span>
                            <Tag type={d.status} small />
                          </div>
                        ))}
                        {r.submissionRate != null && (
                          <div style={{ marginTop: 12 }}>
                            <span style={{ fontSize: 12, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "2px 8px", color: T.brand }}>Target met rate: {r.submissionRate}%</span>
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
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 6 }}>
              <Metric label="Total Projects"    value={projects.length} />
              <Metric label="Active"          value={projects.filter(p => p.status === "active").length}            status="yellow" />
              <Metric label="Pending Approval" value={projects.filter(p => p.status === "pending approval").length} status="blue"   />
              <Metric label="Completed"       value={projects.filter(p => p.status === "completed").length}         status="green"  />
              {projects.length > 0 && <Metric label="Avg Progress" value={`${Math.round(projects.reduce((a, p) => a + p.progress, 0) / projects.length)}%`} />}
              {(() => { const ti = projects.filter(p => p.status !== "completed").reduce((a, p) => a + (p.income || 0), 0); return ti > 0 ? <Metric label="Estimated Income" value={`$${ti.toLocaleString()}`} status="blue" /> : null; })()}
              {(() => { const tp = projects.filter(p => p.status !== "completed").reduce((a, p) => a + (p.income != null && p.margin != null ? Math.round(p.income * p.margin / 100) : 0), 0); return tp > 0 ? <Metric label="Estimated Profit" value={`$${tp.toLocaleString()}`} status="blue" /> : null; })()}
              {(() => { const ti = projects.filter(p => p.status === "completed").reduce((a, p) => a + (p.income || 0), 0); return ti > 0 ? <Metric label="Completed Income" value={`$${ti.toLocaleString()}`} status="green" /> : null; })()}
              {(() => { const tp = projects.filter(p => p.status === "completed").reduce((a, p) => a + (p.income != null && p.margin != null ? Math.round(p.income * p.margin / 100) : 0), 0); return tp > 0 ? <Metric label="Completed Profit" value={`$${tp.toLocaleString()}`} status="green" /> : null; })()}
            </div>
            <div style={{ marginBottom: 16 }}>
              <Input
                value={projSearch}
                onChange={e => setProjSearch(e.target.value)}
                placeholder="Search by user name..."
                style={{ width: 260, padding: "7px 12px", fontSize: 14 }}
              />
              {projSearch && <button onClick={() => setProjSearch("")} style={{ marginLeft: 8, background: "none", border: "none", cursor: "pointer", color: T.textMuted, fontSize: 13 }}>✕ Clear</button>}
            </div>
            {projects.length === 0 && <EmptyState text="No projects yet." />}
            {projects.length > 0 && (() => {
              const ownerDept = p => users.find(u => u.id === p.mgrId)?.deptId || null;
              const searchLower = projSearch.trim().toLowerCase();
              const groups = [...depts.map(d => ({ id: d.id, name: d.name })), { id: null, name: "Other" }];
              const groupElements = groups.map(group => {
                const deptProjects = (group.id
                  ? projects.filter(p => ownerDept(p) === group.id)
                  : projects.filter(p => !ownerDept(p))
                ).filter(p => !searchLower || (users.find(u => u.id === p.mgrId)?.name || "").toLowerCase().includes(searchLower));
                if (deptProjects.length === 0) return null;
                return (
                  <div key={group.id ?? "__other"} style={{ marginBottom: 24 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, paddingBottom: 8, borderBottom: `2px solid ${T.border}` }}>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{group.name}</div>
                      <span style={{ marginLeft: "auto", fontSize: 12, color: T.textMuted }}>{deptProjects.length} project{deptProjects.length !== 1 ? "s" : ""}</span>
                    </div>
                    {deptProjects.map(p => {
                      const mgr = users.find(u => u.id === p.mgrId);
                      const draftProg = progressEdits[p.id] ?? p.progress;
                      const ps = draftProg >= 70 ? "green" : draftProg >= 35 ? "yellow" : "red";
                      const progChanged = progressEdits[p.id] !== undefined;
                      const isDetailsOpen = editProjId === p.id;
                      return (
                        <Card key={p.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                          <div style={{ padding: "12px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                            <div>
                              <div style={{ fontSize: 15, fontWeight: 700 }}>{p.name}</div>
                              <div style={{ fontSize: 12, color: T.textMuted }}>{mgr ? `${mgr.name} · ` : ""}Due: {p.due}{p.updatedDate ? ` · Updated: ${p.updatedDate}` : ""}</div>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <Tag type={p.status === "active" ? "pending" : p.status === "pending approval" ? "review" : "approved"} label={p.status === "active" ? "ACTIVE" : p.status === "pending approval" ? "PENDING APPROVAL" : "COMPLETED"} small />
                              <button onClick={() => { if (window.confirm(`Delete project "${p.name}"? This cannot be undone.`)) dispatch({ type: "REMOVE_PROJECT", projectId: p.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete project">✕</button>
                            </div>
                          </div>
                          <div style={{ padding: "10px 18px", display: "flex", alignItems: "center", gap: 10, borderBottom: `1px solid ${T.border}` }}>
                            <Bar value={draftProg} status={ps} h={6} />
                            <Input value={draftProg} onChange={e => setProgressEdits(d => ({ ...d, [p.id]: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 52, textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                            <span style={{ fontSize: 13, color: T.textMuted }}>%</span>
                            {p.income != null && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Income: ${p.income.toLocaleString()}</span>}
                            {p.income != null && p.margin != null && <span style={{ fontSize: 11, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Profit: ${Math.round(p.income * p.margin / 100).toLocaleString()} ({p.margin}%)</span>}
                            <Btn primary small disabled={!progChanged} onClick={() => {
                              dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { progress: draftProg, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                              setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                            }}>Save</Btn>
                          </div>
                          {!isDetailsOpen && (() => { const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []); const latest = entries[0]; if (!latest) return null; const truncated = latest.text.length > 160; const preview = truncated ? latest.text.slice(0, 160) + "…" : latest.text; return <div style={{ padding: "8px 18px 4px", fontSize: 13, color: T.textSoft, lineHeight: 1.5 }}>{latest.date && <span style={{ fontSize: 11, color: T.textMuted, marginRight: 6 }}>{latest.date}</span>}{preview}{truncated && <button onClick={e => { e.stopPropagation(); setLogPopup({ text: latest.text, date: latest.date, projName: p.name }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, padding: "0 0 0 4px", fontFamily: F.body }}>Read more →</button>}</div>; })()}
                          <div style={{ padding: "8px 18px" }}>
                            <button onClick={() => {
                              if (isDetailsOpen) { setEditProjId(null); return; }
                              setEditProjId(p.id);
                              setEditProjForm({ name: p.name, status: p.status, startDate: p.startDate || "", due: p.due || "", income: p.income != null ? String(p.income) : "", margin: p.margin != null ? String(p.margin) : "", contributeRate: p.contributeRate != null ? String(p.contributeRate) : "" });
                            }} style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 10px", cursor: "pointer", color: T.textMuted, fontSize: 12, fontFamily: F.body }}>
                              {isDetailsOpen ? "▼ Edit Details" : "▸ Edit Details"}
                            </button>
                          </div>
                          {isDetailsOpen && (
                            <div style={{ padding: "14px 18px", borderTop: `1px solid ${T.border}` }}>
                              <div style={{ display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Name</div>
                                  <Input value={editProjForm.name} onChange={e => setEditProjForm(f => ({ ...f, name: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                                  <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 14, fontFamily: F.body }}>
                                    <option value="active">Active</option>
                                    <option value="pending approval">Pending Approval</option>
                                    <option value="completed">Completed</option>
                                  </select>
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Start Date</div>
                                  <Input type="date" value={editProjForm.startDate} onChange={e => setEditProjForm(f => ({ ...f, startDate: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                                  <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                                </div>
                              </div>
                              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Income ($)</div>
                                  <Input type="number" value={editProjForm.income} onChange={e => setEditProjForm(f => ({ ...f, income: e.target.value }))} placeholder="e.g. 250000" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Profit Margin (%)</div>
                                  <Input type="number" value={editProjForm.margin} onChange={e => setEditProjForm(f => ({ ...f, margin: e.target.value }))} placeholder="e.g. 30" min="0" max="100" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Contribution Rate (%)</div>
                                  <Input type="number" value={editProjForm.contributeRate} onChange={e => setEditProjForm(f => ({ ...f, contributeRate: e.target.value }))} placeholder="100 (default)" min="1" max="100" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                                  <div style={{ fontSize: 10, color: T.textMuted, marginTop: 3 }}>Leave blank if sole owner</div>
                                </div>
                                <div>
                                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Profit (Auto)</div>
                                  {(() => {
                                    const hasData = editProjForm.income !== "" && editProjForm.margin !== "";
                                    const fullProfit = hasData ? Math.round(Number(editProjForm.income) * Number(editProjForm.margin) / 100) : null;
                                    const rate = editProjForm.contributeRate !== "" ? Math.min(100, Math.max(1, Number(editProjForm.contributeRate))) : 100;
                                    const myProfit = fullProfit != null ? Math.round(fullProfit * rate / 100) : null;
                                    return (
                                      <div style={{ padding: "7px 10px", fontSize: 14, fontFamily: F.mono, color: hasData ? T.ok : T.textMuted, fontWeight: 700 }}>
                                        {fullProfit != null ? `$${fullProfit.toLocaleString()}` : "—"}
                                        {fullProfit != null && rate < 100 && <div style={{ fontSize: 11, color: T.brand, fontWeight: 600, marginTop: 2 }}>Your KR: ${myProfit.toLocaleString()} ({rate}%)</div>}
                                      </div>
                                    );
                                  })()}
                                </div>
                              </div>
                              <div style={{ marginBottom: 16 }}>
                                <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>Project Logbook</div>
                                <TextArea value={logDrafts[p.id] || ""} onChange={e => setLogDrafts(d => ({ ...d, [p.id]: e.target.value }))} placeholder="Add a log entry..." rows={2} />
                                <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 6, marginBottom: 10 }}>
                                  <Btn primary small disabled={!logDrafts[p.id]?.trim()} onClick={() => {
                                    const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []);
                                    dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { log: [{ text: logDrafts[p.id].trim(), date: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) }, ...entries] } });
                                    setLogDrafts(d => { const n = { ...d }; delete n[p.id]; return n; });
                                  }}>Add Entry</Btn>
                                </div>
                                {(() => { const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []); if (!entries.length) return <div style={{ fontSize: 12, color: T.textMuted }}>No log entries yet.</div>; return entries.map((e, i) => <div key={i} style={{ padding: "8px 10px", marginBottom: 6, background: T.bg, borderRadius: 6, border: `1px solid ${T.border}` }}>{e.date && <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 3 }}>{e.date}</div>}<div style={{ fontSize: 13, color: T.text, lineHeight: 1.5, whiteSpace: "pre-wrap", maxHeight: 200, overflowY: "auto" }}>{e.text}</div></div>); })()}
                              </div>
                              <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                                <Btn small onClick={() => setEditProjId(null)}>Cancel</Btn>
                                <Btn primary small disabled={!editProjForm.name.trim()} onClick={() => {
                                  const becomingCompleted = editProjForm.status === "completed" && p.status !== "completed";
                                  const revertingFromCompleted = editProjForm.status !== "completed" && p.status === "completed";
                                  const cr = editProjForm.contributeRate !== "" ? Math.min(100, Math.max(1, Number(editProjForm.contributeRate))) : null;
                                  dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { name: editProjForm.name.trim(), status: editProjForm.status, startDate: editProjForm.startDate || p.startDate || "", due: editProjForm.due || p.due, income: editProjForm.income !== "" ? Number(editProjForm.income) : null, margin: editProjForm.margin !== "" ? Math.min(100, Math.max(0, Number(editProjForm.margin))) : null, contributeRate: cr, ...(becomingCompleted && { completedYear: new Date().getFullYear() }), ...(revertingFromCompleted && { completedYear: null }), updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                                  setEditProjId(null);
                                }}>Save Details</Btn>
                              </div>
                            </div>
                          )}
                        </Card>
                      );
                    })}
                  </div>
                );
              });
              const pendingApprovalProjects = projects.filter(p => p.status === "pending approval" && (!searchLower || (users.find(u => u.id === p.mgrId)?.name || "").toLowerCase().includes(searchLower)));
              const hasAny = pendingApprovalProjects.length > 0 || groupElements.some(Boolean);
              if (!hasAny) return <EmptyState text="No managers match your search." />;
              return (
                <>
                  {pendingApprovalProjects.length > 0 && (
                    <div style={{ marginBottom: 24 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10, paddingBottom: 8, borderBottom: `2px solid ${T.brandBorder}` }}>
                        <div style={{ fontSize: 15, fontWeight: 700, color: T.brand }}>Pending Approval</div>
                        <span style={{ marginLeft: "auto", fontSize: 12, color: T.textMuted }}>{pendingApprovalProjects.length} project{pendingApprovalProjects.length !== 1 ? "s" : ""}</span>
                      </div>
                      {pendingApprovalProjects.map(p => {
                        const mgr = users.find(u => u.id === p.mgrId);
                        const draftProg = progressEdits[p.id] ?? p.progress;
                        const ps = draftProg >= 70 ? "green" : draftProg >= 35 ? "yellow" : "red";
                        return (
                          <Card key={p.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                            <div style={{ padding: "12px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                              <div>
                                <div style={{ fontSize: 15, fontWeight: 700 }}>{p.name}</div>
                                <div style={{ fontSize: 12, color: T.textMuted }}>{mgr ? `${mgr.name} · ` : ""}{p.startDate ? `Start: ${p.startDate} · ` : ""}Due: {p.due}{p.updatedDate ? ` · Updated: ${p.updatedDate}` : ""}</div>
                              </div>
                              <Tag type="review" label="PENDING APPROVAL" small />
                            </div>
                            <div style={{ padding: "10px 18px", display: "flex", alignItems: "center", gap: 10, borderBottom: `1px solid ${T.border}` }}>
                              <Bar value={draftProg} status={ps} h={6} />
                              <span style={{ fontSize: 13, color: T.textMuted, fontFamily: F.mono }}>{draftProg}%</span>
                              {p.income != null && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Income: ${p.income.toLocaleString()}</span>}
                            {p.income != null && p.margin != null && <span style={{ fontSize: 11, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Profit: ${Math.round(p.income * p.margin / 100).toLocaleString()} ({p.margin}%)</span>}
                            </div>
                            <div style={{ padding: "8px 18px", display: "flex", justifyContent: "flex-end", gap: 8 }}>
                              <Btn small onClick={() => { if (window.confirm(`Reject "${p.name}"? This will revert the project to Active.`)) dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: "active", completedYear: null, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } }); }}>Reject</Btn>
                              <Btn primary small onClick={() => { if (window.confirm(`Approve "${p.name}" as completed?`)) dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: "completed", completedYear: new Date().getFullYear(), updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } }); }}>Approve</Btn>
                            </div>
                          </Card>
                        );
                      })}
                    </div>
                  )}
                  {groupElements}
                </>
              );
            })()}
          </Pane>
        </>)}

        {page === "admissions" && (<>
          <Header title="Weekly Applications Dashboard" sub="Marketer application tracking by RTO" />
          <Pane>
            {enrError && <div style={{ padding: "10px 14px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, fontSize: 13, color: T.bad, marginBottom: 16, lineHeight: 1.5 }}>{enrError}</div>}
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 20 }}>
              {[["overview","Overview"],["imports","Imports"],["dashboard","Dashboard"],["data","Data"]].map(([v, label]) => (
                <Btn key={v} small primary={enrTab === v} onClick={() => setEnrTab(v)}>{label}</Btn>
              ))}
            </div>
            {enrLoading && <div style={{ padding: 32, textAlign: "center", color: T.textMuted, fontSize: 14 }}>Loading application data…</div>}
            {!enrLoading && enrTab === "overview" && (() => {
              const weeks = enrSortWeeksDesc([...new Set(enrRecords.map(r => r.week))]);
              const selWeek = enrFilterWeek !== "all" && weeks.includes(enrFilterWeek) ? enrFilterWeek : (weeks[0] || null);
              const weekRecs = selWeek ? enrRecords.filter(r => r.week === selWeek) : [];
              const marketers = [...new Set(weekRecs.map(r => r.marketerName))].sort();
              const rtos = [...new Set(weekRecs.map(r => r.rto))].sort();
              const total = weekRecs.reduce((s, r) => s + r.count, 0);
              const matrix = {};
              weekRecs.forEach(r => { if (!matrix[r.marketerName]) matrix[r.marketerName] = {}; matrix[r.marketerName][r.rto] = (matrix[r.marketerName][r.rto] || 0) + r.count; });
              const prevWeekIdx = selWeek ? weeks.indexOf(selWeek) + 1 : -1;
              const prevWeek = prevWeekIdx > 0 && prevWeekIdx < weeks.length ? weeks[prevWeekIdx] : null;
              const prevTotal = prevWeek ? enrRecords.filter(r => r.week === prevWeek).reduce((s, r) => s + r.count, 0) : null;
              const diff = prevTotal !== null ? total - prevTotal : null;
              const selCss = { padding: "7px 12px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 13, fontFamily: F.body };
              const thCss = { padding: "8px 14px", textAlign: "left", borderBottom: `2px solid ${T.border}`, color: T.textMuted, fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", whiteSpace: "nowrap", background: T.raised };
              const tdCss = { padding: "8px 14px", borderBottom: `1px solid ${T.border}`, fontSize: 13 };
              return (<>
                <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 16 }}>
                  <span style={{ fontSize: 13, color: T.textMuted }}>Week:</span>
                  {weeks.length === 0 ? <span style={{ fontSize: 13, color: T.textMuted }}>No data yet</span> : (
                    <select value={selWeek || ""} onChange={e => setEnrFilterWeek(e.target.value)} style={selCss}>
                      {weeks.map(w => <option key={w} value={w}>{w} · {weekToDateRange(w, true)}</option>)}
                    </select>
                  )}
                  {selWeek && <span style={{ fontSize: 12, color: T.textMuted }}>{weekToDateRange(selWeek)}</span>}
                  {diff !== null && <span style={{ fontSize: 12, padding: "3px 8px", borderRadius: 5, background: diff >= 0 ? T.okDim : T.badDim, color: diff >= 0 ? T.ok : T.bad, border: `1px solid ${diff >= 0 ? T.okBorder : T.badBorder}`, fontFamily: F.mono }}>{diff >= 0 ? "▲" : "▼"} {Math.abs(diff)} vs {prevWeek} · {weekToDateRange(prevWeek, true)}</span>}
                </div>
                {weeks.length === 0 ? <EmptyState text="No application data yet. Go to Imports to upload a file." /> : (<>
                  <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                    <Metric label="Total Applications" value={total} status={total > 0 ? "green" : undefined} />
                    <Metric label="Marketers" value={marketers.length} />
                    <Metric label="RTOs" value={rtos.length} />
                  </div>
                  {weeks.length >= 2 && (() => {
                    const LINE_COLORS = ["#4C8BF5","#E8542A","#2DB66F","#F5A623","#9B59B6","#1ABC9C","#E74C3C","#F39C12"];
                    const chartWeeks = [...weeks].reverse();
                    const allMarketers = [...new Set(enrRecords.map(r => r.marketerName))];
                    const mTotals = allMarketers.map(m => [m, enrRecords.filter(r => r.marketerName === m).reduce((s, r) => s + r.count, 0)]);
                    mTotals.sort((a, b) => b[1] - a[1]);
                    const topMarketers = mTotals.slice(0, 8).map(x => x[0]);
                    const weeklyTotals = {};
                    topMarketers.forEach(m => { weeklyTotals[m] = {}; });
                    enrRecords.forEach(r => { if (weeklyTotals[r.marketerName] !== undefined) weeklyTotals[r.marketerName][r.week] = (weeklyTotals[r.marketerName][r.week] || 0) + r.count; });
                    const maxVal = Math.max(...topMarketers.flatMap(m => chartWeeks.map(w => weeklyTotals[m][w] || 0)));
                    const yMax = Math.ceil(maxVal / 5) * 5 || 5;
                    const PAD_L = 36, PAD_R = 16, PAD_T = 16, PAD_B = 68;
                    const SVG_W = 700, SVG_H = 260;
                    const cW = SVG_W - PAD_L - PAD_R, cH = SVG_H - PAD_T - PAD_B;
                    const xPos = i => PAD_L + (i / (chartWeeks.length - 1)) * cW;
                    const yPos = v => PAD_T + cH - (v / yMax) * cH;
                    const buildPath = m => {
                      let d = "", prevHad = false;
                      chartWeeks.forEach((w, i) => {
                        const v = weeklyTotals[m][w];
                        if (v !== undefined) {
                          const x = xPos(i).toFixed(1), y = yPos(v).toFixed(1);
                          d += prevHad ? ` L ${x} ${y}` : `M ${x} ${y}`;
                          prevHad = true;
                        } else { prevHad = false; }
                      });
                      return d;
                    };
                    const yTicks = [0,1,2,3,4,5].map(i => Math.round((yMax / 5) * i));
                    return (<>
                      <SectionLabel style={{ marginTop: 4, marginBottom: 8 }}>Application Trend</SectionLabel>
                      <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, padding: "12px 8px 4px", marginBottom: 20 }}>
                        <svg viewBox={`0 0 ${SVG_W} ${SVG_H}`} width="100%" style={{ display: "block" }}>
                          {yTicks.map(v => {
                            const y = yPos(v);
                            return (
                              <g key={v}>
                                <line x1={PAD_L} y1={y} x2={PAD_L + cW} y2={y} stroke={T.border} strokeDasharray={v === 0 ? "0" : "3,3"} strokeWidth="1" />
                                <text x={PAD_L - 6} y={y + 4} textAnchor="end" fontSize="10" fill={T.textMuted}>{v}</text>
                              </g>
                            );
                          })}
                          <line x1={PAD_L} y1={PAD_T + cH} x2={PAD_L + cW} y2={PAD_T + cH} stroke={T.border} strokeWidth="1" />
                          {chartWeeks.map((w, i) => (
                            <text key={w} x={xPos(i).toFixed(1)} y={PAD_T + cH + 10} textAnchor="end" fontSize="10" fill={T.textMuted}
                              transform={`rotate(-40, ${xPos(i).toFixed(1)}, ${PAD_T + cH + 10})`}>{w}</text>
                          ))}
                          {topMarketers.map((m, mi) => {
                            const color = LINE_COLORS[mi];
                            return (
                              <g key={m}>
                                <path d={buildPath(m)} stroke={color} strokeWidth="2.2" fill="none" strokeLinejoin="round" />
                                {chartWeeks.map((w, i) => {
                                  const v = weeklyTotals[m][w];
                                  if (v === undefined) return null;
                                  return (
                                    <circle key={w} cx={xPos(i).toFixed(1)} cy={yPos(v).toFixed(1)} r="4.5" fill={color} stroke={T.surface} strokeWidth="1.5"
                                      style={{ cursor: "pointer" }}
                                      onMouseEnter={() => setEnrChartTooltip({ marketer: m, week: w, count: v, x: xPos(i), y: yPos(v) })}
                                      onMouseLeave={() => setEnrChartTooltip(null)} />
                                  );
                                })}
                              </g>
                            );
                          })}
                          {enrChartTooltip && (() => {
                            const { marketer, week, count, x, y } = enrChartTooltip;
                            const tipW = 152, tipH = 46;
                            const tx = x > SVG_W / 2 ? x - tipW - 10 : x + 12;
                            const ty = Math.max(PAD_T, Math.min(y - tipH / 2, PAD_T + cH - tipH));
                            return (
                              <g style={{ pointerEvents: "none" }}>
                                <rect x={tx} y={ty} width={tipW} height={tipH} rx="5" fill={T.raised} stroke={T.border} strokeWidth="1" />
                                <text x={tx + 9} y={ty + 17} fontSize="11" fontWeight="700" fill={T.text}>{marketer}</text>
                                <text x={tx + 9} y={ty + 33} fontSize="11" fill={T.textMuted}>{week} · {count} applications</text>
                              </g>
                            );
                          })()}
                        </svg>
                        <div style={{ display: "flex", gap: 12, flexWrap: "wrap", padding: "0 8px 8px", marginTop: 2 }}>
                          {topMarketers.map((m, i) => (
                            <div key={m} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                              <div style={{ width: 14, height: 3, background: LINE_COLORS[i], borderRadius: 2, flexShrink: 0 }} />
                              <span style={{ fontSize: 11, color: T.textMuted }}>{m}</span>
                            </div>
                          ))}
                          {allMarketers.length > 8 && <span style={{ fontSize: 11, color: T.textMuted, fontStyle: "italic" }}>· Showing top 8 marketers</span>}
                        </div>
                      </div>
                    </>);
                  })()}
                  {weekRecs.length === 0 ? <EmptyState text="No records for this week." /> : (
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                        <thead>
                          <tr>
                            <th style={thCss}>Marketer</th>
                            {rtos.map(rto => <th key={rto} style={{ ...thCss, textAlign: "right", color: T.brand }}>{rto}</th>)}
                            <th style={{ ...thCss, textAlign: "right" }}>Total</th>
                          </tr>
                        </thead>
                        <tbody>
                          {marketers.map(m => {
                            const rowTotal = rtos.reduce((s, rto) => s + (matrix[m]?.[rto] || 0), 0);
                            return (
                              <tr key={m}>
                                <td style={{ ...tdCss, fontWeight: 600 }}>{m}</td>
                                {rtos.map(rto => <td key={rto} style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, color: matrix[m]?.[rto] ? T.text : T.textMuted }}>{matrix[m]?.[rto] || "—"}</td>)}
                                <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: T.brand }}>{rowTotal}</td>
                              </tr>
                            );
                          })}
                          <tr style={{ background: T.raised }}>
                            <td style={{ ...tdCss, fontWeight: 700, borderBottom: "none" }}>Grand Total</td>
                            {rtos.map(rto => { const n = weekRecs.filter(r => r.rto === rto).reduce((s, r) => s + r.count, 0); return <td key={rto} style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700, borderBottom: "none" }}>{n}</td>; })}
                            <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: T.brand, borderBottom: "none" }}>{total}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  )}
                  {weeks.length > 1 && (<>
                    <SectionLabel style={{ marginTop: 28 }}>All Weeks</SectionLabel>
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                        <thead><tr>{["Week","Total","Marketers","Top Marketer"].map(h => <th key={h} style={thCss}>{h}</th>)}</tr></thead>
                        <tbody>
                          {weeks.map(w => {
                            const wR = enrRecords.filter(r => r.week === w);
                            const wTot = wR.reduce((s, r) => s + r.count, 0);
                            const mMap = {}; wR.forEach(r => { mMap[r.marketerName] = (mMap[r.marketerName] || 0) + r.count; });
                            const topM = Object.entries(mMap).sort((a, b) => b[1] - a[1])[0];
                            return (
                              <tr key={w} style={{ cursor: "pointer", background: w === selWeek ? T.brandDim : "transparent" }} onClick={() => setEnrFilterWeek(w)}>
                                <td style={{ ...tdCss, fontWeight: w === selWeek ? 700 : 400, color: w === selWeek ? T.brand : T.text }}>{w}</td>
                                <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700 }}>{wTot}</td>
                                <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono }}>{new Set(wR.map(r => r.marketerName)).size}</td>
                                <td style={tdCss}>{topM ? `${topM[0]} (${topM[1]})` : "—"}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </>)}
                </>)}
              </>);
            })()}
            {!enrLoading && enrTab === "imports" && (() => {
              const handleFile = async file => {
                if (!file) return;
                if (file.size > 10 * 1024 * 1024) { setEnrError("File exceeds 10 MB."); return; }
                setEnrError(null);
                try {
                  const { default: readXlsxFile } = await import("read-excel-file/browser");
                  let result;
                  try { result = await readXlsxFile(file, { sheets: ["Marketer"] }); }
                  catch (e) {
                    if (e.message?.toLowerCase().includes("not found") || e.constructor?.name === "SheetNotFoundError") throw new Error('Sheet "Marketer" not found. Check the file has a worksheet named exactly "Marketer".');
                    throw e;
                  }
                  const rows = result[0]?.data || [];
                  const parsed = enrParseMarketerSheet(rows, file.name);
                  if (parsed.error) { setEnrError(parsed.error); return; }
                  const existingWeekRtos = new Set(enrRecords.filter(r => r.week === parsed.week).map(r => r.rto));
                  const overlappingRtos = parsed.rtos.filter(rto => existingWeekRtos.has(rto));
                  setEnrParsed({ ...parsed, fileName: file.name, fileSize: file.size, weekAlreadyHasData: existingWeekRtos.size > 0, overlappingRtos, existingRtos: [...existingWeekRtos] });
                } catch (e) { setEnrError(`Parse error: ${e.message}`); }
              };
              const doImport = async () => {
                if (!enrParsed) return;
                setEnrImporting(true);
                try {
                  const now = new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
                  const batchId = `enr${Date.now()}`;
                  const batch = { id: batchId, fileName: enrParsed.fileName, fileSize: enrParsed.fileSize, week: enrParsed.week, totalEnrolments: enrParsed.totalEnrolments, importedAt: now };
                  const records = enrParsed.records.map((r, i) => ({ ...r, id: `${batchId}_${i}`, week: enrParsed.week, batchId, importedAt: now }));
                  await dbUpsert("enrolment_batches", batch);
                  await dbBulkInsert("enrolment_records", records);
                  setEnrRecords(prev => [...prev, ...records]);
                  setEnrBatches(prev => [...prev, batch]);
                  setEnrParsed(null);
                  setEnrError(null);
                } catch (e) { setEnrError(`Import failed: ${e.message}`); }
                finally { setEnrImporting(false); }
              };
              const thCss = { padding: "6px 12px", textAlign: "left", borderBottom: `1px solid ${T.border}`, color: T.textMuted, fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", whiteSpace: "nowrap" };
              const tdCss = { padding: "7px 12px", borderBottom: `1px solid ${T.border}`, fontSize: 13 };
              return (<>
                {!enrParsed && (
                  <Card style={{ padding: 32, textAlign: "center", border: `2px dashed ${T.border}`, cursor: "pointer" }}
                    onDragOver={e => e.preventDefault()}
                    onDrop={e => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }}>
                    <div style={{ fontSize: 30, marginBottom: 8 }}>⬆</div>
                    <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 4 }}>Drop file here or browse</div>
                    <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 16 }}>.xlsx · Marketer worksheet · max 10 MB</div>
                    <input type="file" accept=".xlsx,.xls" id="enr-file-input" style={{ display: "none" }} onChange={e => { handleFile(e.target.files[0]); e.target.value = ""; }} />
                    <Btn primary onClick={() => document.getElementById("enr-file-input").click()}>Browse files</Btn>
                  </Card>
                )}
                {enrParsed && (<>
                  <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 14 }}>
                    <Metric label="Week" value={enrParsed.week} />
                    <Metric label="Total Applications" value={enrParsed.totalEnrolments} status="green" />
                    <Metric label="Marketers" value={enrParsed.marketers.length} />
                    <Metric label="RTOs" value={enrParsed.rtos.length} />
                  </div>
                  {enrParsed.weekAlreadyHasData && enrParsed.overlappingRtos.length === 0 && (
                    <div style={{ padding: "10px 14px", background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 7, fontSize: 13, color: T.brand, marginBottom: 14, lineHeight: 1.5 }}>
                      ℹ Week {enrParsed.week} already has data for: {enrParsed.existingRtos.sort().join(", ")}. This file adds new RTOs ({enrParsed.rtos.join(", ")}) — safe to combine.
                    </div>
                  )}
                  {enrParsed.weekAlreadyHasData && enrParsed.overlappingRtos.length > 0 && (
                    <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 7, fontSize: 13, color: T.warn, marginBottom: 14, lineHeight: 1.5 }}>
                      ⚠ Week {enrParsed.week} already has data for: {enrParsed.overlappingRtos.join(", ")}. Importing again will double-count these RTOs. Cancel if this is a re-upload of the same file.
                    </div>
                  )}
                  <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 10 }}>{enrParsed.fileName} · {(enrParsed.fileSize / 1024).toFixed(1)} KB</div>
                  <SectionLabel>Preview</SectionLabel>
                  <div style={{ overflowX: "auto", marginBottom: 16 }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                      <thead>
                        <tr>
                          <th style={thCss}>Marketer</th>
                          {enrParsed.rtos.map(rto => <th key={rto} style={{ ...thCss, textAlign: "right" }}>{rto}</th>)}
                          <th style={{ ...thCss, textAlign: "right" }}>Total</th>
                        </tr>
                      </thead>
                      <tbody>
                        {enrParsed.marketers.map(m => {
                          const mRecs = enrParsed.records.filter(r => r.marketerName === m);
                          const mTotal = mRecs.reduce((s, r) => s + r.count, 0);
                          return (
                            <tr key={m}>
                              <td style={{ ...tdCss, fontWeight: 600 }}>{m}</td>
                              {enrParsed.rtos.map(rto => { const n = mRecs.find(r => r.rto === rto)?.count; return <td key={rto} style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, color: n ? T.text : T.textMuted }}>{n || "—"}</td>; })}
                              <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: T.brand }}>{mTotal}</td>
                            </tr>
                          );
                        })}
                        <tr style={{ background: T.raised }}>
                          <td style={{ ...tdCss, fontWeight: 700, borderBottom: "none" }}>Total</td>
                          {enrParsed.rtos.map(rto => { const n = enrParsed.records.filter(r => r.rto === rto).reduce((s, r) => s + r.count, 0); return <td key={rto} style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700, borderBottom: "none" }}>{n || "—"}</td>; })}
                          <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: T.brand, borderBottom: "none" }}>{enrParsed.totalEnrolments}</td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <Btn onClick={() => { setEnrParsed(null); setEnrError(null); }}>Cancel</Btn>
                    <Btn primary disabled={enrImporting} onClick={doImport}>{enrImporting ? "Importing…" : `Import ${enrParsed.totalEnrolments} applications`}</Btn>
                  </div>
                </>)}
                {enrBatches.length > 0 && (<>
                  <SectionLabel style={{ marginTop: 28 }}>Import History</SectionLabel>
                  {[...enrBatches].sort((a, b) => String(b.id).localeCompare(String(a.id))).map(b => (
                    <Card key={b.id} style={{ padding: "10px 18px", marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{b.fileName}</div>
                        <div style={{ fontSize: 12, color: T.textMuted }}>{b.importedAt} · Week {b.week}</div>
                      </div>
                      <div style={{ fontSize: 13, fontFamily: F.mono, color: T.ok, fontWeight: 700 }}>{b.totalEnrolments} applications</div>
                    </Card>
                  ))}
                </>)}
              </>);
            })()}
            {!enrLoading && enrTab === "dashboard" && (() => {
              const weeks = enrSortWeeksDesc([...new Set(enrRecords.map(r => r.week))]);
              const selW = enrFilterWeek;
              const filtered = selW === "all" ? enrRecords : enrRecords.filter(r => r.week === selW);
              const mMap = {}; filtered.forEach(r => { mMap[r.marketerName] = (mMap[r.marketerName] || 0) + r.count; });
              const byMarketer = Object.entries(mMap).sort((a, b) => b[1] - a[1]);
              const maxM = byMarketer[0]?.[1] || 1;
              const rMap = {}; filtered.forEach(r => { rMap[r.rto] = (rMap[r.rto] || 0) + r.count; });
              const byRto = Object.entries(rMap).sort((a, b) => b[1] - a[1]);
              const maxR = byRto[0]?.[1] || 1;
              const selCss = { padding: "7px 12px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 13, fontFamily: F.body };
              const barRow = (label, n, max, color) => (
                <div key={label} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 8 }}>
                  <div style={{ width: 150, fontSize: 13, textAlign: "right", color: T.text, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{label}</div>
                  <div style={{ flex: 1, height: 20, background: T.raised, borderRadius: 4, overflow: "hidden", border: `1px solid ${T.border}` }}>
                    <div style={{ width: `${(n / max) * 100}%`, height: "100%", background: color, borderRadius: 4, minWidth: n > 0 ? 4 : 0 }} />
                  </div>
                  <div style={{ width: 36, fontSize: 13, fontFamily: F.mono, textAlign: "right", fontWeight: 700, flexShrink: 0 }}>{n}</div>
                </div>
              );
              return (<>
                <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 20 }}>
                  <span style={{ fontSize: 13, color: T.textMuted }}>Filter:</span>
                  <select value={selW} onChange={e => setEnrFilterWeek(e.target.value)} style={selCss}>
                    <option value="all">All Weeks</option>
                    {weeks.map(w => <option key={w} value={w}>{w} · {weekToDateRange(w, true)}</option>)}
                  </select>
                  {selW !== "all" && <span style={{ fontSize: 12, color: T.textMuted }}>{weekToDateRange(selW)}</span>}
                </div>
                {filtered.length === 0 ? <EmptyState text={enrRecords.length === 0 ? "No data yet. Import a file to get started." : "No data for this week."} /> : (<>
                  <SectionLabel>By Marketer</SectionLabel>
                  <div style={{ marginBottom: 24 }}>{byMarketer.map(([name, n]) => barRow(name, n, maxM, T.brand))}</div>
                  <SectionLabel>By RTO</SectionLabel>
                  <div style={{ marginBottom: 8 }}>{byRto.map(([name, n]) => barRow(name, n, maxR, T.ok))}</div>
                </>)}
              </>);
            })()}
            {!enrLoading && enrTab === "data" && (() => {
              const weeks = enrSortWeeksDesc([...new Set(enrRecords.map(r => r.week))]);
              const marketers = [...new Set(enrRecords.map(r => r.marketerName))].sort();
              const rtos = [...new Set(enrRecords.map(r => r.rto))].sort();
              const filtered = enrRecords.filter(r => {
                if (enrFilterWeek !== "all" && r.week !== enrFilterWeek) return false;
                if (enrFilterMarketer !== "all" && r.marketerName !== enrFilterMarketer) return false;
                if (enrFilterRto !== "all" && r.rto !== enrFilterRto) return false;
                return true;
              }).sort((a, b) => b.week.localeCompare(a.week) || a.marketerName.localeCompare(b.marketerName) || a.rto.localeCompare(b.rto));
              const exportCSV = () => {
                const hdr = ["Week","Marketer","RTO","Count"];
                const rows = filtered.map(r => [r.week, r.marketerName, r.rto, r.count]);
                const csv = [hdr, ...rows].map(r => r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
                const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(new Blob([csv], { type: "text/csv" })), download: "applications.csv" });
                a.click(); URL.revokeObjectURL(a.href);
              };
              const selCss = { padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 13, fontFamily: F.body };
              const thCss = { padding: "8px 12px", textAlign: "left", borderBottom: `2px solid ${T.border}`, color: T.textMuted, fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", whiteSpace: "nowrap" };
              return (<>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12, alignItems: "center" }}>
                  <select value={enrFilterWeek} onChange={e => setEnrFilterWeek(e.target.value)} style={selCss}>
                    <option value="all">All weeks</option>
                    {weeks.map(w => <option key={w} value={w}>{w} · {weekToDateRange(w, true)}</option>)}
                  </select>
                  <select value={enrFilterMarketer} onChange={e => setEnrFilterMarketer(e.target.value)} style={selCss}>
                    <option value="all">All marketers</option>
                    {marketers.map(m => <option key={m} value={m}>{m}</option>)}
                  </select>
                  <select value={enrFilterRto} onChange={e => setEnrFilterRto(e.target.value)} style={selCss}>
                    <option value="all">All RTOs</option>
                    {rtos.map(r => <option key={r} value={r}>{r}</option>)}
                  </select>
                  <Btn small onClick={exportCSV}>Export CSV</Btn>
                </div>
                <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 8 }}>Showing {filtered.length} of {enrRecords.length} records</div>
                {filtered.length === 0 ? <EmptyState text={enrRecords.length === 0 ? "No records yet. Import a file to get started." : "No records match the current filters."} /> : (
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                      <thead><tr>{["Week","Marketer","RTO","Count"].map(h => <th key={h} style={thCss}>{h}</th>)}</tr></thead>
                      <tbody>
                        {filtered.map((r, i) => (
                          <tr key={r.id || i} style={{ borderBottom: `1px solid ${T.border}` }}>
                            <td style={{ padding: "7px 12px", fontFamily: F.mono, fontSize: 12 }}>{r.week}</td>
                            <td style={{ padding: "7px 12px" }}>{r.marketerName}</td>
                            <td style={{ padding: "7px 12px" }}>{r.rto}</td>
                            <td style={{ padding: "7px 12px", fontFamily: F.mono, fontWeight: 700, textAlign: "right" }}>{r.count}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
                <div style={{ marginTop: 32, paddingTop: 16, borderTop: `1px solid ${T.border}` }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>Danger Zone</div>
                  <Btn danger small onClick={async () => {
                    if (!window.confirm("Delete ALL application records and import history? This cannot be undone.")) return;
                    if (!window.confirm("Are you sure? All application data will be permanently deleted.")) return;
                    try {
                      await dbDeleteCollection("enrolment_records");
                      await dbDeleteCollection("enrolment_batches");
                      setEnrRecords([]); setEnrBatches([]); setEnrError(null);
                    } catch (e) { setEnrError(`Clear failed: ${e.message}`); }
                  }}>Clear all application data</Btn>
                </div>
              </>);
            })()}
          </Pane>
        </>)}

        {page === "coe" && (() => {
          const thCss = { padding: "8px 12px", textAlign: "left", borderBottom: `2px solid ${T.border}`, color: T.textMuted, fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: "0.05em", whiteSpace: "nowrap", background: T.raised };
          const tdCss = { padding: "8px 12px", borderBottom: `1px solid ${T.border}`, fontSize: 13 };
          const selCss = { padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 13, fontFamily: F.body };
          return (<>
            <Header title="COE Dashboard" sub="Weekly COE & Non-CoE applicant records by RTO" />
            <Pane>
              {coeError && <div style={{ padding: "10px 14px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, fontSize: 13, color: T.bad, marginBottom: 16, lineHeight: 1.5 }}>{coeError}</div>}
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 20 }}>
                {[["overview","Overview"],["marketer","Marketer"],["imports","Imports"],["data","Data"]].map(([v, label]) => (
                  <Btn key={v} small primary={coeTab === v} onClick={() => setCoeTab(v)}>{label}</Btn>
                ))}
              </div>
              {coeLoading && <div style={{ padding: 32, textAlign: "center", color: T.textMuted, fontSize: 14 }}>Loading COE data…</div>}

              {!coeLoading && coeTab === "overview" && (() => {
                const NIET_ROWS = [{ rto: "NIET", type: "CoE" }, { rto: "NIET", type: "Non-CoE" }, { rto: "CB", type: "CoE" }, { rto: "CB", type: "Non-CoE" }, { rto: "Rhodes", type: "Accepted & Paid" }];
                const EDUCARE_ROWS = [{ rto: "Educare BNE", type: "CoE" }, { rto: "Educare GC", type: "CoE" }, { rto: "Educare ONLINE", type: "Accepted & Paid" }, { rto: "Educare GC", type: "Non-CoE" }, { rto: "Educare BNE", type: "Non-CoE" }, { rto: "Educare Dom", type: "Accepted & Paid" }];
                const weeks = enrSortWeeksDesc([...new Set(coeRecords.map(r => r.week))]);
                const selWeek = coeFilterWeek !== "all" && weeks.includes(coeFilterWeek) ? coeFilterWeek : (weeks[0] || null);
                const weekRecs = selWeek ? coeRecords.filter(r => r.week === selWeek) : [];
                const prevWeekIdx = selWeek ? weeks.indexOf(selWeek) + 1 : -1;
                const prevWeek = prevWeekIdx > 0 && prevWeekIdx < weeks.length ? weeks[prevWeekIdx] : null;
                const prevTotal = prevWeek ? coeRecords.filter(r => r.week === prevWeek).length : null;
                const diff = prevTotal !== null ? weekRecs.length - prevTotal : null;
                const nietCbRhodesTotal = weekRecs.filter(r => !r.rto.startsWith("Educare")).length;
                const educareTotal = weekRecs.filter(r => r.rto.startsWith("Educare")).length;
                const renderSectionTable = (rows, label) => {
                  const sectionRecs = rows.map(c => ({ ...c, n: weekRecs.filter(r => r.rto === c.rto && r.type === c.type).length }));
                  const subtotal = sectionRecs.reduce((s, c) => s + c.n, 0);
                  return (<>
                    <SectionLabel>{label}</SectionLabel>
                    <div style={{ overflowX: "auto", marginBottom: 24 }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden", maxWidth: 480 }}>
                        <thead><tr><th style={thCss}>RTO</th><th style={thCss}>Type</th><th style={{ ...thCss, textAlign: "right" }}>Records</th></tr></thead>
                        <tbody>
                          {sectionRecs.map(c => (
                            <tr key={`${c.rto}-${c.type}`}>
                              <td style={{ ...tdCss, fontWeight: 600 }}>{c.rto}</td>
                              <td style={tdCss}>{c.type}</td>
                              <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, color: c.n ? T.text : T.textMuted }}>{c.n || "—"}</td>
                            </tr>
                          ))}
                          <tr style={{ background: T.raised }}>
                            <td colSpan={2} style={{ ...tdCss, fontWeight: 700, borderBottom: "none" }}>Subtotal</td>
                            <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: T.brand, borderBottom: "none" }}>{subtotal}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </>);
                };
                return (<>
                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginBottom: 16 }}>
                    <span style={{ fontSize: 13, color: T.textMuted }}>Week:</span>
                    {weeks.length === 0 ? <span style={{ fontSize: 13, color: T.textMuted }}>No data yet</span> : (
                      <select value={selWeek || ""} onChange={e => setCoeFilterWeek(e.target.value)} style={selCss}>
                        {weeks.map(w => <option key={w} value={w}>{w} · {weekToDateRange(w, true)}</option>)}
                      </select>
                    )}
                    {selWeek && <span style={{ fontSize: 12, color: T.textMuted }}>{weekToDateRange(selWeek)}</span>}
                    {diff !== null && <span style={{ fontSize: 12, padding: "3px 8px", borderRadius: 5, background: diff >= 0 ? T.okDim : T.badDim, color: diff >= 0 ? T.ok : T.bad, border: `1px solid ${diff >= 0 ? T.okBorder : T.badBorder}`, fontFamily: F.mono }}>{diff >= 0 ? "▲" : "▼"} {Math.abs(diff)} vs {prevWeek} · {weekToDateRange(prevWeek, true)}</span>}
                  </div>
                  {weeks.length === 0 ? <EmptyState text="No COE data yet. Go to Imports to upload a file." /> : (<>
                    <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                      <Metric label="Total Records" value={weekRecs.length} status={weekRecs.length > 0 ? "green" : undefined} />
                      <Metric label="NIET / CB / Rhodes" value={nietCbRhodesTotal} />
                      <Metric label="Educare" value={educareTotal} />
                    </div>
                    {renderSectionTable(NIET_ROWS, "NIET / CB / Rhodes")}
                    {renderSectionTable(EDUCARE_ROWS, "Educare")}
                    {weekRecs.length > 0 && (<>
                      <SectionLabel>Student Records</SectionLabel>
                      <div style={{ overflowX: "auto" }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                          <thead><tr>
                            {["Student ID","RTO","Type","Course Name","Agent","Marketer","Created By","Onshore/Offshore"].map(h => <th key={h} style={thCss}>{h}</th>)}
                          </tr></thead>
                          <tbody>
                            {weekRecs.sort((a, b) => a.rto.localeCompare(b.rto) || a.type.localeCompare(b.type) || a.studentId.localeCompare(b.studentId)).map((r, i) => (
                              <tr key={r.id || i} style={{ borderBottom: `1px solid ${T.border}` }}>
                                <td style={{ padding: "7px 12px", fontFamily: F.mono, fontSize: 12 }}>{r.studentId}</td>
                                <td style={{ padding: "7px 12px", fontWeight: 600 }}>{r.rto}</td>
                                <td style={{ padding: "7px 12px" }}>{r.type}</td>
                                <td style={{ padding: "7px 12px" }}>{r.courseName}</td>
                                <td style={{ padding: "7px 12px" }}>{r.agent}</td>
                                <td style={{ padding: "7px 12px" }}>{r.marketer}</td>
                                <td style={{ padding: "7px 12px" }}>{r.createdBy}</td>
                                <td style={{ padding: "7px 12px" }}>{r.onshoreOffshore}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </>)}
                  </>)}
                </>);
              })()}

              {!coeLoading && coeTab === "marketer" && (() => {
                const weeks = enrSortWeeksDesc([...new Set(coeRecords.map(r => r.week))]);
                const selW = coeFilterWeek;
                const filtered = (selW === "all" ? coeRecords : coeRecords.filter(r => r.week === selW)).filter(r => r.marketer);
                const COE_COLS = [
                  { rto: "NIET", type: "CoE", label: "NIET CoE" },
                  { rto: "NIET", type: "Non-CoE", label: "NIET Non-CoE" },
                  { rto: "CB", type: "CoE", label: "CB CoE" },
                  { rto: "CB", type: "Non-CoE", label: "CB Non-CoE" },
                  { rto: "Rhodes", type: "Accepted & Paid", label: "Rhodes" },
                  { rto: "Educare BNE", type: "CoE", label: "EDU BNE CoE" },
                  { rto: "Educare GC", type: "CoE", label: "EDU GC CoE" },
                  { rto: "Educare ONLINE", type: "Accepted & Paid", label: "EDU ONLINE" },
                  { rto: "Educare GC", type: "Non-CoE", label: "EDU GC Non-CoE" },
                  { rto: "Educare BNE", type: "Non-CoE", label: "EDU BNE Non-CoE" },
                  { rto: "Educare Dom", type: "Accepted & Paid", label: "EDU Dom" },
                ];
                const mTotals = {};
                filtered.forEach(r => { mTotals[r.marketer] = (mTotals[r.marketer] || 0) + 1; });
                const marketers = Object.entries(mTotals).sort((a, b) => b[1] - a[1]).map(([m]) => m);
                const maxTotal = mTotals[marketers[0]] || 1;
                const matrix = {};
                filtered.forEach(r => { if (!matrix[r.marketer]) matrix[r.marketer] = {}; const k = `${r.rto} ${r.type}`; matrix[r.marketer][k] = (matrix[r.marketer][k] || 0) + 1; });
                const colTotals = {};
                COE_COLS.forEach(c => { colTotals[`${c.rto} ${c.type}`] = filtered.filter(r => r.rto === c.rto && r.type === c.type).length; });
                return (<>
                  <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 20 }}>
                    <span style={{ fontSize: 13, color: T.textMuted }}>Week:</span>
                    <select value={selW} onChange={e => setCoeFilterWeek(e.target.value)} style={selCss}>
                      <option value="all">All Weeks</option>
                      {weeks.map(w => <option key={w} value={w}>{w} · {weekToDateRange(w, true)}</option>)}
                    </select>
                    {selW !== "all" && <span style={{ fontSize: 12, color: T.textMuted }}>{weekToDateRange(selW)}</span>}
                  </div>
                  {filtered.length === 0 ? <EmptyState text={coeRecords.length === 0 ? "No data yet. Import a file to get started." : "No data for this week."} /> : (<>
                    <SectionLabel>Marketer Breakdown</SectionLabel>
                    <div style={{ overflowX: "auto", marginBottom: 28 }}>
                      <table style={{ borderCollapse: "collapse", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                        <thead>
                          <tr>
                            <th style={{ ...thCss, minWidth: 130 }}>Marketer</th>
                            {COE_COLS.map(c => <th key={c.label} style={{ ...thCss, textAlign: "right", minWidth: 90 }}>{c.label}</th>)}
                            <th style={{ ...thCss, textAlign: "right", minWidth: 70 }}>Total</th>
                          </tr>
                        </thead>
                        <tbody>
                          {marketers.map(m => {
                            const row = matrix[m] || {};
                            const total = mTotals[m] || 0;
                            return (<tr key={m}>
                              <td style={{ ...tdCss, fontWeight: 600 }}>{m}</td>
                              {COE_COLS.map(c => { const n = row[`${c.rto} ${c.type}`] || 0; return <td key={c.label} style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, color: n ? T.text : T.textMuted }}>{n || "—"}</td>; })}
                              <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: T.brand }}>{total}</td>
                            </tr>);
                          })}
                          <tr style={{ background: T.raised }}>
                            <td style={{ ...tdCss, fontWeight: 700, borderBottom: "none" }}>Total</td>
                            {COE_COLS.map(c => { const n = colTotals[`${c.rto} ${c.type}`] || 0; return <td key={c.label} style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 700, borderBottom: "none" }}>{n || "—"}</td>; })}
                            <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: T.brand, borderBottom: "none" }}>{filtered.length}</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <SectionLabel>By Marketer</SectionLabel>
                    <div style={{ marginBottom: 8 }}>
                      {marketers.map(m => {
                        const n = mTotals[m] || 0;
                        return (<div key={m} style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 8 }}>
                          <div style={{ width: 140, fontSize: 13, textAlign: "right", color: T.text, flexShrink: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{m}</div>
                          <div style={{ flex: 1, height: 20, background: T.raised, borderRadius: 4, overflow: "hidden", border: `1px solid ${T.border}` }}>
                            <div style={{ width: `${(n / maxTotal) * 100}%`, height: "100%", background: T.brand, borderRadius: 4, minWidth: n > 0 ? 4 : 0 }} />
                          </div>
                          <div style={{ width: 36, fontSize: 13, fontFamily: F.mono, textAlign: "right", fontWeight: 700, flexShrink: 0 }}>{n}</div>
                        </div>);
                      })}
                    </div>
                  </>)}
                </>);
              })()}

              {!coeLoading && coeTab === "imports" && (() => {
                const handleNietFile = async file => {
                  if (!file) return;
                  if (file.size > 10 * 1024 * 1024) { setCoeError("File exceeds 10 MB."); return; }
                  setCoeError(null);
                  try {
                    const parsed = await coeParseSheetsFromFile(file);
                    if (parsed.error && parsed.sheetsFound.length === 0) { setCoeError(parsed.error); return; }
                    const existingWeekRecs = coeRecords.filter(r => r.week === parsed.week);
                    const existingTypes = [...new Set(existingWeekRecs.map(r => `${r.rto} ${r.type}`))];
                    const overlapping = parsed.sheetsFound.filter(s => existingTypes.includes(s));
                    setCoeParsed({ ...parsed, weekAlreadyHasData: existingWeekRecs.length > 0, overlapping, existingTypes });
                  } catch (e) { setCoeError(`Parse error: ${e.message}`); }
                };
                const doNietImport = async () => {
                  if (!coeParsed) return;
                  setCoeImporting(true);
                  try {
                    const now = new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
                    const batchId = `coe${Date.now()}`;
                    const batch = { id: batchId, fileName: coeParsed.fileName, fileSize: coeParsed.fileSize, week: coeParsed.week, totalRecords: coeParsed.totalRecords, sheetsFound: coeParsed.sheetsFound, importedAt: now };
                    const records = coeParsed.records.map((r, i) => ({ ...r, id: `${batchId}_${i}`, week: coeParsed.week, batchId, importedAt: now }));
                    await dbUpsert("coe_batches", batch);
                    await dbBulkInsert("coe_records", records);
                    setCoeRecords(prev => [...prev, ...records]);
                    setCoeBatches(prev => [...prev, batch]);
                    setCoeParsed(null);
                    setCoeError(null);
                  } catch (e) { setCoeError(`Import failed: ${e.message}`); }
                  finally { setCoeImporting(false); }
                };
                const handleEducareFile = async file => {
                  if (!file) return;
                  if (file.size > 10 * 1024 * 1024) { setCoeError("File exceeds 10 MB."); return; }
                  setCoeError(null);
                  try {
                    const parsed = await educareParseSheetsFromFile(file);
                    if (parsed.error && parsed.sheetsFound.length === 0) { setCoeError(parsed.error); return; }
                    const existingWeekRecs = coeRecords.filter(r => r.week === parsed.week);
                    const existingTypes = [...new Set(existingWeekRecs.map(r => `${r.rto} ${r.type}`))];
                    const overlapping = parsed.sheetsFound.filter(s => existingTypes.includes(s));
                    setEducareParsed({ ...parsed, weekAlreadyHasData: existingWeekRecs.length > 0, overlapping, existingTypes });
                  } catch (e) { setCoeError(`Parse error: ${e.message}`); }
                };
                const doEducareImport = async () => {
                  if (!educareParsed) return;
                  setEducareImporting(true);
                  try {
                    const now = new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
                    const batchId = `edu${Date.now()}`;
                    const batch = { id: batchId, fileName: educareParsed.fileName, fileSize: educareParsed.fileSize, week: educareParsed.week, totalRecords: educareParsed.totalRecords, sheetsFound: educareParsed.sheetsFound, importedAt: now };
                    const records = educareParsed.records.map((r, i) => ({ ...r, id: `${batchId}_${i}`, week: educareParsed.week, batchId, importedAt: now }));
                    await dbUpsert("coe_batches", batch);
                    await dbBulkInsert("coe_records", records);
                    setCoeRecords(prev => [...prev, ...records]);
                    setCoeBatches(prev => [...prev, batch]);
                    setEducareParsed(null);
                    setCoeError(null);
                  } catch (e) { setCoeError(`Import failed: ${e.message}`); }
                  finally { setEducareImporting(false); }
                };
                const renderImportCard = (label, fileInputId, parsed, importing, parsedSetter, handleFile, doImport) => (
                  <div style={{ flex: 1, minWidth: 280 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 10 }}>{label}</div>
                    {!parsed && (
                      <Card style={{ padding: 24, textAlign: "center", border: `2px dashed ${T.border}`, cursor: "pointer" }}
                        onDragOver={e => e.preventDefault()}
                        onDrop={e => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); }}>
                        <div style={{ fontSize: 28, marginBottom: 6 }}>⬆</div>
                        <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 3 }}>Drop file here or browse</div>
                        <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 14 }}>.xlsx · max 10 MB</div>
                        <input type="file" accept=".xlsx,.xls" id={fileInputId} style={{ display: "none" }} onChange={e => { handleFile(e.target.files[0]); e.target.value = ""; }} />
                        <Btn primary small onClick={() => document.getElementById(fileInputId).click()}>Browse files</Btn>
                      </Card>
                    )}
                    {parsed && (<>
                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 10 }}>
                        <Metric label="Week" value={parsed.week || "?"} />
                        <Metric label="Records" value={parsed.totalRecords} status={parsed.totalRecords > 0 ? "green" : undefined} />
                        <Metric label="Sheets" value={parsed.sheetsFound.length} />
                      </div>
                      {parsed.weekAlreadyHasData && parsed.overlapping.length > 0 && (
                        <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 7, fontSize: 13, color: T.warn, marginBottom: 10, lineHeight: 1.5 }}>
                          ⚠ Week {parsed.week} already has data for: {parsed.overlapping.join(", ")}. Importing again will duplicate records.
                        </div>
                      )}
                      {parsed.weekAlreadyHasData && parsed.overlapping.length === 0 && (
                        <div style={{ padding: "10px 14px", background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 7, fontSize: 13, color: T.brand, marginBottom: 10, lineHeight: 1.5 }}>
                          ℹ Week {parsed.week} already has data for: {parsed.existingTypes.join(", ")}. This adds new sheets — safe to combine.
                        </div>
                      )}
                      {parsed.error && (
                        <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 7, fontSize: 13, color: T.warn, marginBottom: 10, lineHeight: 1.5 }}>
                          ⚠ {parsed.error}
                        </div>
                      )}
                      <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 8 }}>{parsed.fileName} · {(parsed.fileSize / 1024).toFixed(1)} KB</div>
                      <div style={{ overflowX: "auto", marginBottom: 12 }}>
                        <table style={{ width: "100%", borderCollapse: "collapse", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                          <thead><tr><th style={thCss}>RTO</th><th style={thCss}>Type</th><th style={{ ...thCss, textAlign: "right" }}>Found</th></tr></thead>
                          <tbody>
                            {parsed.sheetSummary.map(s => (
                              <tr key={`${s.rto}-${s.type}`}>
                                <td style={{ ...tdCss, fontWeight: 600 }}>{s.rto}</td>
                                <td style={tdCss}>{s.type}</td>
                                <td style={{ ...tdCss, textAlign: "right", fontFamily: F.mono, color: s.found && s.count > 0 ? T.text : T.textMuted }}>{s.found ? (s.count || (s.note ? `0 (${s.note})` : "—")) : <span style={{ color: T.textDim, fontStyle: "italic" }}>Not found</span>}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div style={{ display: "flex", gap: 8 }}>
                        <Btn onClick={() => { parsedSetter(null); setCoeError(null); }}>Cancel</Btn>
                        <Btn primary disabled={importing || parsed.totalRecords === 0} onClick={doImport}>{importing ? "Importing…" : `Import ${parsed.totalRecords} records`}</Btn>
                      </div>
                    </>)}
                  </div>
                );
                return (<>
                  <div style={{ display: "flex", gap: 20, flexWrap: "wrap", marginBottom: 28 }}>
                    {renderImportCard("NIET / CB / Rhodes Report", "coe-file-input", coeParsed, coeImporting, setCoeParsed, handleNietFile, doNietImport)}
                    {renderImportCard("Educare Report", "educare-file-input", educareParsed, educareImporting, setEducareParsed, handleEducareFile, doEducareImport)}
                  </div>
                  {coeBatches.length > 0 && (<>
                    <SectionLabel>Import History</SectionLabel>
                    {[...coeBatches].sort((a, b) => String(b.id).localeCompare(String(a.id))).map(b => (
                      <Card key={b.id} style={{ padding: "10px 18px", marginBottom: 8, display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                        <div>
                          <div style={{ fontSize: 13, fontWeight: 700 }}>{b.fileName}</div>
                          <div style={{ fontSize: 12, color: T.textMuted }}>{b.importedAt} · Week {b.week} · {b.sheetsFound?.join(", ")}</div>
                        </div>
                        <div style={{ fontSize: 13, fontFamily: F.mono, color: T.ok, fontWeight: 700 }}>{b.totalRecords} records</div>
                      </Card>
                    ))}
                  </>)}
                </>);
              })()}

              {!coeLoading && coeTab === "data" && (() => {
                const weeks = enrSortWeeksDesc([...new Set(coeRecords.map(r => r.week))]);
                const rtos = [...new Set(coeRecords.map(r => r.rto))].sort();
                const types = [...new Set(coeRecords.map(r => r.type))].sort();
                const filtered = coeRecords.filter(r => {
                  if (coeFilterWeek !== "all" && r.week !== coeFilterWeek) return false;
                  if (coeFilterRto !== "all" && r.rto !== coeFilterRto) return false;
                  if (coeFilterType !== "all" && r.type !== coeFilterType) return false;
                  return true;
                }).sort((a, b) => b.week.localeCompare(a.week) || a.rto.localeCompare(b.rto) || a.type.localeCompare(b.type));
                const exportCSV = () => {
                  const hdr = ["Student ID","RTO","Type","Week","Period","Course Name","Intake Date","Agent","Marketer","Created By","Onshore/Offshore","Pathway"];
                  const rows = filtered.map(r => [r.studentId, r.rto, r.type, r.week, r.period, r.courseName, r.intakeDate, r.agent, r.marketer, r.createdBy, r.onshoreOffshore, r.pathway]);
                  const csv = [hdr, ...rows].map(r => r.map(v => `"${String(v ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
                  const a = Object.assign(document.createElement("a"), { href: URL.createObjectURL(new Blob([csv], { type: "text/csv" })), download: "coe_records.csv" });
                  a.click(); URL.revokeObjectURL(a.href);
                };
                const colHdr = ["Student ID","RTO","Type","Week","Course Name","Intake Date","Agent","Marketer","Created By","Onshore/Offshore","Pathway"];
                return (<>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12, alignItems: "center" }}>
                    <select value={coeFilterWeek} onChange={e => setCoeFilterWeek(e.target.value)} style={selCss}>
                      <option value="all">All weeks</option>
                      {weeks.map(w => <option key={w} value={w}>{w} · {weekToDateRange(w, true)}</option>)}
                    </select>
                    <select value={coeFilterRto} onChange={e => setCoeFilterRto(e.target.value)} style={selCss}>
                      <option value="all">All RTOs</option>
                      {rtos.map(r => <option key={r} value={r}>{r}</option>)}
                    </select>
                    <select value={coeFilterType} onChange={e => setCoeFilterType(e.target.value)} style={selCss}>
                      <option value="all">All types</option>
                      {types.map(t => <option key={t} value={t}>{t}</option>)}
                    </select>
                    <Btn small onClick={exportCSV}>Export CSV</Btn>
                  </div>
                  <div style={{ fontSize: 12, color: T.textMuted, marginBottom: 8 }}>Showing {filtered.length} of {coeRecords.length} records</div>
                  {filtered.length === 0 ? <EmptyState text={coeRecords.length === 0 ? "No records yet. Import a file to get started." : "No records match the current filters."} /> : (
                    <div style={{ overflowX: "auto" }}>
                      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden" }}>
                        <thead><tr>{colHdr.map(h => <th key={h} style={thCss}>{h}</th>)}</tr></thead>
                        <tbody>
                          {filtered.map((r, i) => (
                            <tr key={r.id || i} style={{ borderBottom: `1px solid ${T.border}` }}>
                              <td style={{ padding: "7px 12px", fontFamily: F.mono, fontSize: 12 }}>{r.studentId}</td>
                              <td style={{ padding: "7px 12px", fontWeight: 600 }}>{r.rto}</td>
                              <td style={{ padding: "7px 12px" }}>{r.type}</td>
                              <td style={{ padding: "7px 12px", fontFamily: F.mono, fontSize: 12 }}>{r.week}</td>
                              <td style={{ padding: "7px 12px" }}>{r.courseName}</td>
                              <td style={{ padding: "7px 12px", fontFamily: F.mono, fontSize: 12 }}>{r.intakeDate}</td>
                              <td style={{ padding: "7px 12px" }}>{r.agent}</td>
                              <td style={{ padding: "7px 12px" }}>{r.marketer}</td>
                              <td style={{ padding: "7px 12px" }}>{r.createdBy}</td>
                              <td style={{ padding: "7px 12px" }}>{r.onshoreOffshore}</td>
                              <td style={{ padding: "7px 12px" }}>{r.pathway}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                  <div style={{ marginTop: 32, paddingTop: 16, borderTop: `1px solid ${T.border}` }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>Danger Zone</div>
                    <Btn danger small onClick={async () => {
                      if (!window.confirm("Delete ALL COE records and import history? This cannot be undone.")) return;
                      if (!window.confirm("Are you sure? All COE data will be permanently deleted.")) return;
                      try {
                        await dbDeleteCollection("coe_records");
                        await dbDeleteCollection("coe_batches");
                        setCoeRecords([]); setCoeBatches([]); setCoeError(null);
                      } catch (e) { setCoeError(`Clear failed: ${e.message}`); }
                    }}>Clear all COE data</Btn>
                  </div>
                </>);
              })()}
            </Pane>
          </>);
        })()}

        {page === "leaderboard" && (<>
          <Header title="Company Leaderboard" sub={`All staff ranked by OKR completion · ${currentFYQuarter()}`} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Total"    value={allMembers.length} />
              <Metric label="On Track" value={allMembers.filter(m => m.status === "green").length}  status="green"  />
              <Metric label="At Risk"  value={allMembers.filter(m => m.status === "yellow").length} status="yellow" />
              <Metric label="Behind"   value={allMembers.filter(m => m.status === "red").length}    status="red"    />
              {allMembers.filter(m => m.status === "none").length > 0 && (
                <Metric label="No OKRs" value={allMembers.filter(m => m.status === "none").length} status="none" />
              )}
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
                const rate = memberHasRateKrs(periodKrs) ? calcMemberRate(m.id, periodKrs, periodSubs) : null;
                return { ...m, rate, status: getStatus(rate) };
              }).sort((a, b) => (b.rate ?? -1) - (a.rate ?? -1));
              const filtered = periodMembers.filter(m =>
                (lbDeptFilter === "all" || m.deptId === lbDeptFilter) &&
                (!q || m.name.toLowerCase().includes(q) || m.title?.toLowerCase().includes(q))
              );
              if (filtered.length === 0) return <EmptyState text="No members match your search." />;
              const COL = "50px 32px 1fr 120px 110px 55px 150px 70px 56px";
              return (
                <Card style={{ overflow: "hidden" }}>
                  <div style={{ overflowX: "auto" }}><div style={{ minWidth: 743 }}>
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
                          <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[m.status].color }}>{m.rate != null ? `${m.rate.toFixed(1)}%` : "N/A"}</span>
                          <Bar value={m.rate ?? 0} status={m.status} h={5} />
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
                            <div style={{ fontSize: 13, fontWeight: 700, color: T.brand, marginBottom: 8 }}>Editing KPIs — {m.name}</div>
                            <div style={{ fontSize: 12, color: "#92400e", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 6, padding: "7px 12px", marginBottom: 12 }}>⚠ Changes here are personal overrides for this member only. They may be overwritten if dept/team KRs are re-synced from the Departments tab.</div>
                            {krs.length === 0
                              ? <div style={{ fontSize: 13, color: T.textMuted }}>No OKRs synced yet — assign this member to a team in User Management, then click ⟳ Sync to Team Members in the Departments tab.</div>
                              : [{ key: "daily", label: "Daily OKRs" }, { key: "weekly", label: "Weekly OKRs" }, { key: "monthly", label: "Monthly OKRs" }, { key: "quarterly", label: "Quarterly OKRs" }, { key: "biannual", label: "Bi-Annual OKRs" }, { key: "annual", label: "Annual OKRs" }].map(({ key, label }) => {
                                const group = krs.filter(kr => (kr.period || "monthly") === key);
                                if (group.length === 0) return null;
                                return (
                                  <div key={key} style={{ marginBottom: 14 }}>
                                    <div style={{ fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase", marginBottom: 6 }}>{label} ({group.length})</div>
                                    <div style={{ overflowX: "auto" }}><div style={{ minWidth: 600 }}>
                                    <div style={{ display: "grid", gridTemplateColumns: "50px 1fr 90px 110px 55px 130px 28px 30px", gap: 8, padding: "6px 10px", fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.06em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}` }}>
                                      <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Performance Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span></span><span></span>
                                    </div>
                                    {group.map((kr, ki) => {
                                      if (lbEditKr?.memberId === m.id && lbEditKr?.krId === kr.id) {
                                        const isTracker = lbKrForm.krType === "tracker";
                                        const isMgrFill = lbKrForm.krType === "manager-fill";
                                        const isProjProfit = lbKrForm.krType === "project_profit";
                                        const isStandard = !isTracker && lbKrForm.krType !== "progress" && !isMgrFill && !isProjProfit;
                                        const sel = { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 12px", color: T.text, fontSize: 14, fontFamily: F.body, outline: "none" };
                                        const saveLbEdit = () => {
                                          if (!lbKrForm.label) return;
                                          if (!isProjProfit && lbKrForm.krType !== "tracker" && lbKrForm.krType !== "progress" && !lbKrForm.useMonthlyTargets && lbKrForm.target === "") return;
                                          if (lbKrForm.krType === "progress" && lbKrForm.target === "") return;
                                          if (isProjProfit && (lbKrForm.target === "" || !lbKrForm.krYear)) return;
                                          const base = { id: kr.id, label: lbKrForm.label, unit: lbKrForm.unit.trim(), dataSource: lbKrForm.dataSource.trim(), operator: lbKrForm.operator || ">=", period: lbKrForm.period || "monthly" };
                                          let updated;
                                          if (lbKrForm.krType === "tracker") updated = { ...base, type: "tracker", target: 0, actual: kr.actual, disallowZero: !!lbKrForm.disallowZero };
                                          else if (lbKrForm.krType === "progress") updated = { ...base, type: "progress", target: Number(lbKrForm.target), actual: kr.actual };
                                          else if (lbKrForm.krType === "manager-fill") updated = { ...base, type: "manager-fill", target: Number(lbKrForm.target), actual: kr.actual };
                                          else if (isProjProfit) updated = { id: kr.id, label: lbKrForm.label, type: "project_profit", period: "annual", target: Number(lbKrForm.target), krYear: Number(lbKrForm.krYear) || new Date().getFullYear(), actual: null };
                                          else if (lbKrForm.useMonthlyTargets) updated = { ...base, monthlyTargets: lbKrForm.monthlyTargets, monthlyActuals: kr.monthlyActuals || {}, ...(Number(lbKrForm.dreamTarget) > 0 && { annualTarget: Number(lbKrForm.dreamTarget) }) };
                                          else updated = { ...base, target: Number(lbKrForm.target), actual: kr.actual };
                                          dispatch({ type: "REPLACE_MEMBER_KR", memberId: m.id, krId: kr.id, kr: updated });
                                          setLbEditKr(null);
                                          setLbKrForm({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false });
                                        };
                                        return (
                                          <div key={kr.id} style={{ padding: "14px 16px", background: T.warnDim, borderTop: `2px solid ${T.warn}`, borderBottom: `1px solid ${T.border}` }}>
                                            <div style={{ fontSize: 12, fontWeight: 700, color: T.warn, marginBottom: 10, letterSpacing: "0.04em", textTransform: "uppercase" }}>Edit Key Result — {kr.id}</div>
                                            <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
                                              <Input value={lbKrForm.label} onChange={e => setLbKrForm(p => ({ ...p, label: e.target.value }))} placeholder="Key result name *" style={{ flex: 2, minWidth: 200 }} />
                                              <select value={lbKrForm.krType} onChange={e => setLbKrForm(p => ({ ...p, krType: e.target.value, useMonthlyTargets: false }))} style={{ ...sel, flex: 1, minWidth: 160 }}>
                                                <option value="">Standard (Yes / No)</option>
                                                <option value="tracker">Tracker (number only)</option>
                                                <option value="progress">Progress (cumulative)</option>
                                                <option value="manager-fill">Manager Fill</option>
                                                <option value="project_profit">Project Profit (auto)</option>
                                              </select>
                                              {!isProjProfit && <select value={lbKrForm.period || "monthly"} onChange={e => setLbKrForm(p => ({ ...p, period: e.target.value }))} style={{ ...sel, minWidth: 120 }}>
                                                <option value="daily">Daily</option>
                                                <option value="weekly">Weekly</option>
                                                <option value="monthly">Monthly</option>
                                                <option value="quarterly">Quarterly</option>
                                                <option value="biannual">Bi-Annual</option>
                                                <option value="annual">Annual</option>
                                              </select>}
                                            </div>
                                            <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap", alignItems: "center" }}>
                                              {!isTracker && !isProjProfit && !lbKrForm.useMonthlyTargets && (<>
                                                <select value={lbKrForm.operator || ">="} onChange={e => setLbKrForm(p => ({ ...p, operator: e.target.value }))} style={{ ...sel, width: 80 }}>
                                                  <option value=">=">≥</option>
                                                  <option value="<=">≤</option>
                                                  <option value="=">=</option>
                                                </select>
                                                <Input value={lbKrForm.target} onChange={e => setLbKrForm(p => ({ ...p, target: e.target.value }))} placeholder="Target *" style={{ width: 110 }} />
                                              </>)}
                                              {isProjProfit && (<>
                                                <span style={{ fontSize: 13, color: T.textMuted }}>$</span>
                                                <Input value={lbKrForm.target} onChange={e => setLbKrForm(p => ({ ...p, target: e.target.value }))} placeholder="Annual profit target *" style={{ width: 160 }} />
                                                <Input value={lbKrForm.krYear || ""} onChange={e => setLbKrForm(p => ({ ...p, krYear: e.target.value }))} placeholder="Year (e.g. 2026) *" style={{ width: 140 }} />
                                              </>)}
                                              {!isProjProfit && <Input value={lbKrForm.unit} onChange={e => setLbKrForm(p => ({ ...p, unit: e.target.value }))} placeholder="Unit (e.g. $, Days)" style={{ width: 150 }} />}
                                              {!isProjProfit && <Input value={lbKrForm.dataSource} onChange={e => setLbKrForm(p => ({ ...p, dataSource: e.target.value }))} placeholder="Data source" style={{ flex: 1, minWidth: 140 }} />}
                                              {isStandard && (
                                                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                                                  <input type="checkbox" checked={!!lbKrForm.useMonthlyTargets} onChange={e => setLbKrForm(p => ({ ...p, useMonthlyTargets: e.target.checked }))} style={{ accentColor: T.brand }} />
                                                  Monthly targets
                                                </label>
                                              )}
                                              {isTracker && (
                                                <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                                                  <input type="checkbox" checked={!!lbKrForm.disallowZero} onChange={e => setLbKrForm(p => ({ ...p, disallowZero: e.target.checked }))} style={{ accentColor: T.bad }} />
                                                  Block zero submission
                                                </label>
                                              )}
                                            </div>
                                            {isStandard && lbKrForm.useMonthlyTargets && (
                                              <div style={{ marginBottom: 8 }}>
                                                <Input value={lbKrForm.dreamTarget} onChange={e => setLbKrForm(p => ({ ...p, dreamTarget: e.target.value }))} placeholder="Dream / annual target (optional)" style={{ width: 260, marginBottom: 10 }} />
                                                <div style={{ border: `1px solid ${T.border}`, borderRadius: 8, overflow: "hidden", maxWidth: 360 }}>
                                                  <div style={{ display: "grid", gridTemplateColumns: "120px 1fr", padding: "5px 10px", fontSize: 10, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em", background: T.raised, borderBottom: `1px solid ${T.border}` }}>
                                                    <span>Month</span><span>Target</span>
                                                  </div>
                                                  {getFYMonths().map((mo, mi) => (
                                                    <div key={mo.key} style={{ display: "grid", gridTemplateColumns: "120px 1fr", gap: 8, padding: "4px 10px", alignItems: "center", background: mi % 2 ? T.raised : "transparent", borderBottom: mi < getFYMonths().length - 1 ? `1px solid ${T.border}` : "none" }}>
                                                      <span style={{ fontSize: 12, color: T.textMuted }}>{mo.label}</span>
                                                      <NumInput value={lbKrForm.monthlyTargets[mo.key] ?? 0} onChange={n => setLbKrForm(p => ({ ...p, monthlyTargets: { ...p.monthlyTargets, [mo.key]: n } }))} style={{ padding: "3px 8px", fontSize: 13, fontFamily: F.mono, width: 100 }} />
                                                    </div>
                                                  ))}
                                                </div>
                                              </div>
                                            )}
                                            {isMgrFill && (
                                              <div style={{ fontSize: 12, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "7px 12px", marginBottom: 8 }}>
                                                Manager Fill — member will not receive a check-in; manager assesses yes/no + value in their portal.
                                              </div>
                                            )}
                                            <div style={{ display: "flex", gap: 8 }}>
                                              <Btn primary small onClick={saveLbEdit}>✓ Save Changes</Btn>
                                              <Btn small onClick={() => { setLbEditKr(null); setLbKrForm({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false }); }}>Cancel</Btn>
                                            </div>
                                          </div>
                                        );
                                      }
                                      const isMonthly = !!(kr.monthlyTargets);
                                      const mk = currentFYMonthKey();
                                      const mTgt = isMonthly ? (Number(kr.monthlyTargets[mk]) || 0) : null;
                                      const mAct = isMonthly ? ((kr.monthlyActuals || {})[mk] ?? null) : null;
                                      const lbPPAct = kr.type === "project_profit" ? (() => { const lbGetCY = p => { if (p.completedYear) return p.completedYear; const pts = (p.updatedDate || "").split("/"); return pts.length >= 3 ? parseInt(pts[2].split(",")[0].trim()) : null; }; return (state.projects || []).filter(p => p.mgrId === m.id && p.status === "completed" && lbGetCY(p) === kr.krYear).reduce((s, p) => s + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0); })() : null;
                                      const pct = kr.type === "project_profit" ? (kr.target > 0 ? Math.min(Math.round(lbPPAct / kr.target * 100), 100) : 0) : krCompletion(kr);
                                      const st = getStatus(pct);
                                      return (
                                        <div key={kr.id} style={{ borderBottom: `1px solid ${T.border}` }}>
                                          <div style={{ display: "grid", gridTemplateColumns: "50px 1fr 90px 110px 55px 130px 28px 30px", gap: 8, padding: "8px 10px", alignItems: "center", background: ki % 2 ? T.raised : "transparent", fontSize: 13 }}>
                                            <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim }}>{kr.id}</span>
                                            <div>
                                              <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }} title={kr.label}>{kr.label}</span>
                                              {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                                              {kr.type === "tracker" && kr.disallowZero && <span style={{ fontSize: 10, color: T.bad, background: T.badDim, border: `1px solid ${T.badBorder || T.bad}`, borderRadius: 8, padding: "1px 5px", display: "inline-block", marginLeft: 3 }}>No zeros</span>}
                                              {kr.type === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Progress</span>}
                                              {kr.type === "manager-fill" && <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Mgr Fill</span>}
                                              {kr.type === "project_profit" && <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Proj Profit · auto</span>}
                                              {isMonthly && kr.type !== "tracker" && kr.type !== "progress" && <span style={{ fontSize: 10, fontWeight: 700, color: "#0369a1", background: "#e0f2fe", border: "1px solid #7dd3fc", borderRadius: 8, padding: "1px 5px", display: "inline-block", marginLeft: 4 }}>Monthly</span>}
                                            </div>
                                            {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 11, color: "#7c3aed" }}>N/A</span>
                                              : kr.type === "project_profit" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted, fontSize: 12 }}>${(kr.target || 0).toLocaleString()} ({kr.krYear || "?"})</span>
                                              : isMonthly ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted, fontSize: 12 }}>{kr.operator || ">="} {fmt(mTgt)}{kr.unit ? ` ${kr.unit}` : ""}</span>
                                              : kr.type === "progress" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.target)}{kr.unit ? ` ${kr.unit}` : ""}</span>
                                              : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)}{kr.unit ? ` ${kr.unit}` : ""}</span>}
                                            {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.actual)}</span>
                                              : kr.type === "project_profit" ? (() => { const lbGetCY = p => { if (p.completedYear) return p.completedYear; const pts = (p.updatedDate || "").split("/"); return pts.length >= 3 ? parseInt(pts[2].split(",")[0].trim()) : null; }; const lbAct = (state.projects || []).filter(p => p.mgrId === m.id && p.status === "completed" && lbGetCY(p) === kr.krYear).reduce((s, p) => s + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0); return <span style={{ textAlign: "right", fontFamily: F.mono, color: T.ok, fontWeight: 700 }}>${lbAct.toLocaleString()}</span>; })()
                                              : isMonthly ? <NumInput value={mAct} onChange={n => dispatch({ type: "UPDATE_MEMBER_KR", memberId: m.id, krId: kr.id, field: "monthlyActuals", value: { ...(kr.monthlyActuals || {}), [mk]: n } })} style={{ textAlign: "right", padding: "4px 8px", fontSize: 13, fontFamily: F.mono }} />
                                              : <NumInput value={kr.actual} onChange={n => dispatch({ type: "UPDATE_MEMBER_KR", memberId: m.id, krId: kr.id, field: "actual", value: n })} style={{ textAlign: "right", padding: "4px 8px", fontSize: 13, fontFamily: F.mono }} />}
                                            {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 11, color: "#7c3aed" }}>—</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[st].color }}>{pct.toFixed(0)}%</span>}
                                            {kr.type === "tracker" ? <span /> : <Bar value={pct} status={st} h={5} />}
                                            <button onClick={() => {
                                              const krType = kr.type === "tracker" ? "tracker" : kr.type === "progress" ? "progress" : kr.type === "manager-fill" ? "manager-fill" : kr.type === "project_profit" ? "project_profit" : "";
                                              const isMonthlyKr = !!kr.monthlyTargets;
                                              setLbAddMember(null);
                                              setLbEditKr({ memberId: m.id, krId: kr.id });
                                              setLbKrForm({ label: kr.label, target: isMonthlyKr ? "" : String(kr.target ?? ""), dreamTarget: String(kr.annualTarget || ""), unit: kr.unit || "", dataSource: kr.dataSource || "", operator: kr.operator || ">=", period: kr.period || "monthly", useMonthlyTargets: isMonthlyKr && krType === "", krType, monthlyTargets: Object.fromEntries(getFYMonths().map(mo => [mo.key, kr.monthlyTargets?.[mo.key] ?? 0])), krYear: String(kr.krYear || ""), disallowZero: !!kr.disallowZero });
                                            }} title="Edit this KR" style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 5px", cursor: "pointer", color: T.brand, fontSize: 12, lineHeight: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>✏</button>
                                            <button onClick={() => setConfirmDeleteKr({ memberId: m.id, memberName: m.name, krId: kr.id, krLabel: kr.label })}
                                              title="Delete this OKR"
                                              style={{ background: "none", border: `1px solid ${T.badBorder || T.bad}`, borderRadius: 5, padding: "3px 6px", cursor: "pointer", color: T.bad, fontSize: 12, lineHeight: 1, display: "flex", alignItems: "center", justifyContent: "center" }}>✕</button>
                                          </div>
                                          {isMonthly && kr.type !== "tracker" && (
                                            <div style={{ padding: "4px 10px 10px 60px", background: ki % 2 ? T.raised : "transparent" }}>
                                              <div style={{ display: "grid", gridTemplateColumns: "76px 100px 100px 56px 1fr", gap: 6, padding: "4px 6px", fontSize: 10, fontWeight: 700, color: T.textDim, textTransform: "uppercase", letterSpacing: "0.06em", borderBottom: `1px solid ${T.border}`, marginBottom: 2 }}>
                                                <span>Month</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span />
                                              </div>
                                              {getFYMonths().map(({ key, label }) => {
                                                const mt = Number(kr.monthlyTargets[key]) || 0;
                                                const ma = (kr.monthlyActuals || {})[key];
                                                const mp = ma != null ? (() => {
                                                  const op = kr.operator || ">=";
                                                  if (mt === 0) return (op === ">=" || op === ">") ? 100 : (ma <= 0 ? 100 : Math.min((1/(ma+1))*100, 99));
                                                  switch (op) {
                                                    case ">=": return Math.min((ma/mt)*100, 100);
                                                    case ">":  return ma > mt ? 100 : Math.min((ma/mt)*100, 100);
                                                    case "<=": return ma <= mt ? 100 : Math.min((mt/ma)*100, 100);
                                                    case "<":  return ma < mt ? 100 : Math.min((mt/ma)*100, 100);
                                                    case "=":  return ma === mt ? 100 : 0;
                                                    default:   return Math.min((ma/mt)*100, 100);
                                                  }
                                                })() : null;
                                                const ms = mp != null ? getStatus(mp) : null;
                                                const isCurr = key === mk;
                                                return (
                                                  <div key={key} style={{ display: "grid", gridTemplateColumns: "76px 100px 100px 56px 1fr", gap: 6, padding: "2px 6px", alignItems: "center", fontSize: 12, borderRadius: 4, background: isCurr ? T.brandDim : "transparent" }}>
                                                    <span style={{ fontWeight: isCurr ? 700 : 400, color: isCurr ? T.brand : T.textMuted }}>{label}</span>
                                                    <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(mt)}{kr.unit ? ` ${kr.unit}` : ""}</span>
                                                    <span style={{ textAlign: "right", fontFamily: F.mono, color: ma != null ? T.text : T.textDim }}>{ma != null ? fmt(ma) : "—"}</span>
                                                    <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: ms ? STATUS_THEME[ms].color : T.textDim }}>{mp != null ? mp.toFixed(0) + "%" : "—"}</span>
                                                    <Bar value={mp ?? 0} status={ms || "behind"} h={3} />
                                                  </div>
                                                );
                                              })}
                                            </div>
                                          )}
                                        </div>
                                      );
                                    })}
                                    </div></div>
                                  </div>
                                );
                              })
                            }
                            {lbAddMember === m.id ? (() => {
                              const isTracker = lbKrForm.krType === "tracker";
                              const isMgrFill = lbKrForm.krType === "manager-fill";
                              const isProjProfit = lbKrForm.krType === "project_profit";
                              const isStandard = !isTracker && lbKrForm.krType !== "progress" && !isMgrFill && !isProjProfit;
                              const sel = { background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "10px 12px", color: T.text, fontSize: 14, fontFamily: F.body, outline: "none" };
                              const addLbKr = () => {
                                if (!lbKrForm.label) return;
                                if (!isProjProfit && lbKrForm.krType !== "tracker" && lbKrForm.krType !== "progress" && !lbKrForm.useMonthlyTargets && lbKrForm.target === "") return;
                                if (lbKrForm.krType === "progress" && lbKrForm.target === "") return;
                                if (isProjProfit && (lbKrForm.target === "" || !lbKrForm.krYear)) return;
                                const base = { id: `mkr_${Date.now().toString(36)}`, label: lbKrForm.label, unit: lbKrForm.unit.trim(), dataSource: lbKrForm.dataSource.trim(), operator: lbKrForm.operator || ">=", period: lbKrForm.period || "monthly" };
                                let kr;
                                if (lbKrForm.krType === "tracker") kr = { ...base, type: "tracker", target: 0, actual: null, disallowZero: !!lbKrForm.disallowZero };
                                else if (lbKrForm.krType === "progress") kr = { ...base, type: "progress", target: Number(lbKrForm.target), actual: null };
                                else if (lbKrForm.krType === "manager-fill") kr = { ...base, type: "manager-fill", target: Number(lbKrForm.target), actual: null };
                                else if (isProjProfit) kr = { id: `mkr_${Date.now().toString(36)}`, label: lbKrForm.label, type: "project_profit", period: "annual", target: Number(lbKrForm.target), krYear: Number(lbKrForm.krYear) || new Date().getFullYear(), actual: null };
                                else if (lbKrForm.useMonthlyTargets) kr = { ...base, monthlyTargets: Object.fromEntries(getFYMonths().map(mo => [mo.key, 0])), monthlyActuals: {}, ...(Number(lbKrForm.dreamTarget) > 0 && { annualTarget: Number(lbKrForm.dreamTarget) }) };
                                else kr = { ...base, target: Number(lbKrForm.target), actual: null };
                                dispatch({ type: "ADD_MEMBER_KR", memberId: m.id, kr });
                                setLbAddMember(null);
                                setLbKrForm({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false });
                              };
                              return (
                                <div style={{ padding: "14px 16px", background: T.brandDim, borderTop: `2px solid ${T.brand}`, marginTop: 8 }}>
                                  <div style={{ fontSize: 12, fontWeight: 700, color: T.brand, marginBottom: 10, letterSpacing: "0.04em", textTransform: "uppercase" }}>New Key Result</div>
                                  <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap" }}>
                                    <Input value={lbKrForm.label} onChange={e => setLbKrForm(p => ({ ...p, label: e.target.value }))} placeholder="Key result name *" style={{ flex: 2, minWidth: 200 }} />
                                    <select value={lbKrForm.krType} onChange={e => setLbKrForm(p => ({ ...p, krType: e.target.value, useMonthlyTargets: false }))} style={{ ...sel, flex: 1, minWidth: 160 }}>
                                      <option value="">Standard (Yes / No)</option>
                                      <option value="tracker">Tracker (number only)</option>
                                      <option value="progress">Progress (cumulative)</option>
                                      <option value="manager-fill">Manager Fill</option>
                                      <option value="project_profit">Project Profit (auto)</option>
                                    </select>
                                    {!isProjProfit && <select value={lbKrForm.period || "monthly"} onChange={e => setLbKrForm(p => ({ ...p, period: e.target.value }))} style={{ ...sel, minWidth: 120 }}>
                                      <option value="daily">Daily</option>
                                      <option value="weekly">Weekly</option>
                                      <option value="monthly">Monthly</option>
                                      <option value="quarterly">Quarterly</option>
                                      <option value="biannual">Bi-Annual</option>
                                      <option value="annual">Annual</option>
                                    </select>}
                                  </div>
                                  <div style={{ display: "flex", gap: 8, marginBottom: 8, flexWrap: "wrap", alignItems: "center" }}>
                                    {!isTracker && !isProjProfit && !lbKrForm.useMonthlyTargets && (<>
                                      <select value={lbKrForm.operator || ">="} onChange={e => setLbKrForm(p => ({ ...p, operator: e.target.value }))} style={{ ...sel, width: 80 }}>
                                        <option value=">=">≥</option>
                                        <option value="<=">≤</option>
                                        <option value="=">=</option>
                                      </select>
                                      <Input value={lbKrForm.target} onChange={e => setLbKrForm(p => ({ ...p, target: e.target.value }))} placeholder="Target *" style={{ width: 110 }} />
                                    </>)}
                                    {isProjProfit && (<>
                                      <span style={{ fontSize: 13, color: T.textMuted }}>$</span>
                                      <Input value={lbKrForm.target} onChange={e => setLbKrForm(p => ({ ...p, target: e.target.value }))} placeholder="Annual profit target *" style={{ width: 160 }} />
                                      <Input value={lbKrForm.krYear || ""} onChange={e => setLbKrForm(p => ({ ...p, krYear: e.target.value }))} placeholder="Year (e.g. 2026) *" style={{ width: 140 }} />
                                    </>)}
                                    {!isProjProfit && <Input value={lbKrForm.unit} onChange={e => setLbKrForm(p => ({ ...p, unit: e.target.value }))} placeholder="Unit (e.g. $, Days)" style={{ width: 150 }} />}
                                    {!isProjProfit && <Input value={lbKrForm.dataSource} onChange={e => setLbKrForm(p => ({ ...p, dataSource: e.target.value }))} placeholder="Data source" style={{ flex: 1, minWidth: 140 }} />}
                                    {isStandard && (
                                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                                        <input type="checkbox" checked={!!lbKrForm.useMonthlyTargets} onChange={e => setLbKrForm(p => ({ ...p, useMonthlyTargets: e.target.checked }))} style={{ accentColor: T.brand }} />
                                        Monthly targets
                                      </label>
                                    )}
                                    {isTracker && (
                                      <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textMuted, cursor: "pointer", whiteSpace: "nowrap" }}>
                                        <input type="checkbox" checked={!!lbKrForm.disallowZero} onChange={e => setLbKrForm(p => ({ ...p, disallowZero: e.target.checked }))} style={{ accentColor: T.bad }} />
                                        Block zero submission
                                      </label>
                                    )}
                                  </div>
                                  {isStandard && lbKrForm.useMonthlyTargets && (
                                    <div style={{ marginBottom: 8 }}>
                                      <Input value={lbKrForm.dreamTarget} onChange={e => setLbKrForm(p => ({ ...p, dreamTarget: e.target.value }))} placeholder="Dream / annual target (optional)" style={{ width: 260 }} />
                                    </div>
                                  )}
                                  {isMgrFill && (
                                    <div style={{ fontSize: 12, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "7px 12px", marginBottom: 8 }}>
                                      Manager Fill — member will not receive a check-in; manager assesses yes/no + value in their portal.
                                    </div>
                                  )}
                                  <div style={{ display: "flex", gap: 8 }}>
                                    <Btn primary small onClick={addLbKr}>✓ Add Key Result</Btn>
                                    <Btn small onClick={() => { setLbAddMember(null); setLbKrForm({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false }); }}>Cancel</Btn>
                                  </div>
                                </div>
                              );
                            })() : (
                              <div style={{ marginTop: 10 }}>
                                <Btn small onClick={() => { setLbEditKr(null); setLbAddMember(m.id); setLbKrForm({ label: "", target: "", dreamTarget: "", unit: "", dataSource: "", operator: ">=", period: "monthly", useMonthlyTargets: false, krType: "", monthlyTargets: {}, krYear: "", disallowZero: false }); }}>+ Add Key Result</Btn>
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    );
                  })}
                  </div></div>
                </Card>
              );
            })()}
          </Pane>
        </>)}

        {confirmDeleteKr && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
            <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 16, padding: "28px 32px", width: "100%", maxWidth: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.22)" }}>
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
            fromName: "NIET Group OKRs",
            subject:  "Action Required: {periodLabel} KPI Check-In — {periodKey}",
            body:     "Here are your {periodLower} KPI targets for <strong>{periodKey}</strong>.\nPlease log in to the portal and mark whether you have met each target.",
            ctaText:  "Submit My Check-In →",
            footer:   "You are receiving this because you have KPI targets in the NIET Group OKRs system.\nPlease do not reply to this email.",
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
                  <Field label="Sender Name" value={localDraft.fromName} onChange={v => setField("fromName", v)} hint='Displayed in the From field, e.g. "NIET Group OKRs"' />
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
                          <div style={{ color: "#fff", fontSize: 9, fontWeight: 700, letterSpacing: "0.1em", textTransform: "uppercase", opacity: 0.8, marginBottom: 4 }}>NIET Group OKRs System</div>
                          <div style={{ color: "#fff", fontSize: 17, fontWeight: 700 }}>{previewPeriodLabel} KPI Check-In</div>
                        </div>
                        <div style={{ padding: "20px 24px" }}>
                          <p style={{ margin: "0 0 14px", fontSize: 14, color: "#1d1d1f" }}>Hi <strong>John Smith</strong>,</p>
                          <p style={{ margin: "0 0 14px", fontSize: 13, color: "#6e6e73", lineHeight: 1.6 }} dangerouslySetInnerHTML={{ __html: previewBody.replace(/\n/g, "<br/>") }} />
                          <table style={{ width: "100%", borderCollapse: "collapse", marginBottom: 18 }}>
                            <thead>
                              <tr style={{ background: "#f5f5f7" }}>
                                <th style={{ padding: "6px 10px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#6e6e73", textTransform: "uppercase", letterSpacing: "0.06em" }}>Key Result</th>
                                <th style={{ padding: "6px 10px", textAlign: "right", fontSize: 10, fontWeight: 700, color: "#6e6e73", textTransform: "uppercase", letterSpacing: "0.06em" }}>Performance Target</th>
                                <th style={{ padding: "6px 10px", textAlign: "left", fontSize: 10, fontWeight: 700, color: "#6e6e73", textTransform: "uppercase", letterSpacing: "0.06em" }}>Unit</th>
                              </tr>
                            </thead>
                            <tbody>
                              {[{ label: "Complete 3 coaching sessions", target: "3", unit: "sessions" }, { label: "Customer satisfaction score", target: "≥ 90", unit: "%" }].map((kr, i) => (
                                <tr key={i}><td style={{ padding: "7px 10px", borderBottom: "1px solid #e5e7eb", fontSize: 13 }}>{kr.label}</td><td style={{ padding: "7px 10px", borderBottom: "1px solid #e5e7eb", textAlign: "right", fontFamily: "monospace", fontSize: 13 }}>{kr.target}</td><td style={{ padding: "7px 10px", borderBottom: "1px solid #e5e7eb", fontSize: 12, color: "#6e6e73" }}>{kr.unit}</td></tr>
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
            <div style={{ background: T.surface, borderRadius: 16, padding: "28px 32px", width: "100%", maxWidth: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}>
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

        {page === "ai-chat" && (() => {
          const SUGGESTIONS = [
            { icon: "🏢", text: "Give me a company OKR review for this month." },
            { icon: "📊", text: "Give me a department-by-department OKR review." },
            { icon: "🔴", text: "How many members are in red status this month?" },
            { icon: "📋", text: "Which members haven't submitted their check-in this month?" },
            { icon: "💰", text: "How is the company tracking financially this FY?" },
            { icon: "📈", text: "Which division contributes the most to income?" },
          ];
          const NP_AVATAR = (
            <div style={{ width: 32, height: 32, borderRadius: "50%", background: "linear-gradient(135deg,#0071E3,#6B47DC)", color: "#fff", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, flexShrink: 0, letterSpacing: "0.02em", boxShadow: "0 2px 8px rgba(0,113,227,0.35)" }}>NP</div>
          );
          return (
            <div style={{ display: "flex", flexDirection: "column", height: "100%", maxWidth: 860, margin: "0 auto" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: isMobile ? "14px 16px 12px" : "18px 32px 14px", borderBottom: `1px solid ${T.border}`, flexShrink: 0 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                  {isMobile && (
                    <button onClick={() => setDrawerOpen(true)} style={{ background: "none", border: "none", fontSize: 20, cursor: "pointer", color: T.textSoft, padding: "2px 6px", lineHeight: 1, flexShrink: 0, fontFamily: F.body }}>☰</button>
                  )}
                  <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg,#0071E3,#6B47DC)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 12, fontWeight: 800, letterSpacing: "0.02em", boxShadow: "0 2px 10px rgba(0,113,227,0.3)" }}>NP</div>
                  <div>
                    <div style={{ fontSize: 17, fontWeight: 800, letterSpacing: "-0.01em" }}>NIET Pilot</div>
                    <div style={{ fontSize: 12, color: T.textMuted, marginTop: 1 }}>AI-powered OKR analytics</div>
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                  {chatHistory.length > 0 && <button onClick={() => { archivePilotSession(chatHistory); localStorage.removeItem("niet_pilot_draft"); setChatHistory([]); }} style={{ background: "none", border: `1px solid ${T.border}`, cursor: "pointer", fontSize: 12, color: T.textMuted, fontFamily: F.body, padding: "5px 12px", borderRadius: 8 }}>Clear chat</button>}
                  <button onClick={() => setChatPromptOpen(o => !o)} style={{ background: chatPromptOpen ? T.brandDim : "none", border: `1px solid ${chatPromptOpen ? T.brandBorder : T.border}`, cursor: "pointer", fontSize: 12, color: chatPromptOpen ? T.brand : T.textMuted, fontFamily: F.body, padding: "5px 12px", borderRadius: 8 }}>⚙ System Prompt</button>
                </div>
              </div>
              {chatPromptOpen && (() => {
                const saved = settings?.aiChatPrompt || DEFAULT_CHAT_PROMPT;
                return (
                  <div style={{ margin: isMobile ? "12px 16px 0" : "12px 32px 0", padding: "14px 18px", background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 12 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: T.brand, marginBottom: 8, textTransform: "uppercase", letterSpacing: "0.07em" }}>System Prompt</div>
                    <textarea key="sysPrompt" defaultValue={saved} rows={5} id="chatSysPromptTA"
                      style={{ width: "100%", padding: "9px 12px", fontSize: 13, fontFamily: F.mono, background: T.surface, border: `1px solid ${T.brandBorder}`, borderRadius: 8, color: T.text, outline: "none", resize: "vertical", lineHeight: 1.55, boxSizing: "border-box" }}
                      onFocus={e => e.target.style.borderColor = T.borderFocus} onBlur={e => e.target.style.borderColor = T.brandBorder} />
                    <div style={{ display: "flex", gap: 8, marginTop: 10, justifyContent: "flex-end", alignItems: "center" }}>
                      <button onClick={() => { document.getElementById("chatSysPromptTA").value = DEFAULT_CHAT_PROMPT; }} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 12, color: T.textDim, fontFamily: F.body, padding: 0 }}>Reset to default</button>
                      <Btn small primary onClick={() => { const val = document.getElementById("chatSysPromptTA").value.trim(); if (val) { dispatch({ type: "SET_SETTINGS", updates: { aiChatPrompt: val } }); setChatPromptOpen(false); } }}>Save</Btn>
                    </div>
                  </div>
                );
              })()}
              <div style={{ flex: 1, overflowY: "auto", padding: isMobile ? "0 16px" : "0 32px" }}>
                {chatHistory.length === 0 && (
                  <div style={{ paddingTop: 40, paddingBottom: 24 }}>
                    <div style={{ textAlign: "center", marginBottom: 40 }}>
                      <div style={{ width: 64, height: 64, borderRadius: 18, background: "linear-gradient(135deg,#0071E3,#6B47DC)", display: "flex", alignItems: "center", justifyContent: "center", color: "#fff", fontSize: 22, fontWeight: 800, margin: "0 auto 16px", boxShadow: "0 8px 28px rgba(0,113,227,0.3)" }}>NP</div>
                      <div style={{ fontSize: 24, fontWeight: 800, letterSpacing: "-0.02em", marginBottom: 8 }}>NIET Pilot</div>
                      <div style={{ fontSize: 14, color: T.textMuted, maxWidth: 380, margin: "0 auto", lineHeight: 1.6 }}>Ask anything about your organisation’s OKR performance — completions, trends, member progress, and more.</div>
                    </div>
                    {pilotSessions.length > 0 && (
                      <div style={{ marginBottom: 32 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 10 }}>Recent Conversations</div>
                        {pilotSessions.map(s => (
                          <button key={s.id} onClick={() => setChatHistory(s.messages)}
                            style={{ width: "100%", display: "flex", justifyContent: "space-between", alignItems: "center", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "11px 16px", marginBottom: 8, cursor: "pointer", fontFamily: F.body, textAlign: "left" }}
                            onMouseEnter={e => { e.currentTarget.style.borderColor = T.brandBorder; e.currentTarget.style.boxShadow = `0 0 0 3px ${T.brandDim}`; }}
                            onMouseLeave={e => { e.currentTarget.style.borderColor = T.border; e.currentTarget.style.boxShadow = "none"; }}>
                            <div style={{ minWidth: 0, flex: 1 }}>
                              <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 3 }}>{s.date}</div>
                              <div style={{ fontSize: 13, color: T.textSoft, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{s.preview}{s.preview.length >= 80 ? "…" : ""}</div>
                            </div>
                            <span style={{ fontSize: 16, color: T.textMuted, marginLeft: 12, flexShrink: 0 }}>→</span>
                          </button>
                        ))}
                      </div>
                    )}
                    <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 14 }}>Try asking</div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                      {SUGGESTIONS.map(s => (
                        <button key={s.text} onClick={() => sendChat(s.text)}
                          style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 12, padding: "13px 16px", textAlign: "left", cursor: "pointer", fontFamily: F.body, display: "flex", alignItems: "flex-start", gap: 10 }}
                          onMouseEnter={e => { e.currentTarget.style.borderColor = T.brandBorder; e.currentTarget.style.boxShadow = `0 0 0 3px ${T.brandDim}`; }}
                          onMouseLeave={e => { e.currentTarget.style.borderColor = T.border; e.currentTarget.style.boxShadow = "none"; }}>
                          <span style={{ fontSize: 18, flexShrink: 0, lineHeight: 1.3 }}>{s.icon}</span>
                          <span style={{ fontSize: 13, color: T.textSoft, lineHeight: 1.5 }}>{s.text}</span>
                        </button>
                      ))}
                    </div>
                  </div>
                )}
                {chatHistory.length > 0 && <div style={{ height: 20 }} />}
                {chatHistory.map((msg, i) => {
                  if (msg.role === "action") {
                    return (
                      <div key={i} style={{ display: "flex", alignItems: "flex-start", marginBottom: 16, gap: 10 }}>
                        {NP_AVATAR}
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <ActionReviewCard
                            action={msg.action}
                            submissions={msg.submissions}
                            onConfirm={ids => {
                              ids.forEach(id => {
                                const sub = okrSubmissions.find(s => s.id === id);
                                dispatch({
                                  type: "APPROVE_OKR_SUBMISSION",
                                  id,
                                  status: msg.action.type === "approve" ? "approved" : "rejected",
                                  approvedBy: user.id,
                                  ...(msg.action.type === "reject" ? { actualValue: sub?.actualValue } : {}),
                                });
                              });
                            }}
                            onCancel={() => {}}
                          />
                        </div>
                      </div>
                    );
                  }
                  return (
                    <div key={i} style={{ display: "flex", justifyContent: msg.role === "user" ? "flex-end" : "flex-start", marginBottom: 16, alignItems: "flex-start", gap: 10 }}>
                      {msg.role === "ai" && NP_AVATAR}
                      <div style={{
                        maxWidth: "74%", padding: "12px 18px",
                        borderRadius: msg.role === "user" ? "20px 20px 5px 20px" : "5px 20px 20px 20px",
                        background: msg.role === "user" ? "linear-gradient(135deg,#0071E3,#0077ED)" : T.surface,
                        color: msg.role === "user" ? "#fff" : T.text,
                        fontSize: 14, lineHeight: 1.65,
                        border: msg.role === "ai" ? `1px solid ${T.border}` : "none",
                        boxShadow: msg.role === "user" ? "0 2px 12px rgba(0,113,227,0.25)" : "0 1px 4px rgba(0,0,0,0.06)",
                      }}>{msg.role === "ai" ? <MdMsg text={msg.text} /> : msg.text}</div>
                      {msg.role === "user" && <div style={{ width: 32, height: 32, borderRadius: "50%", background: T.raised, border: `1px solid ${T.border}`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 12, fontWeight: 700, flexShrink: 0, color: T.textMuted }}>{(user.name || "U").slice(0,1).toUpperCase()}</div>}
                    </div>
                  );
                })}
                {chatLoading && (
                  <div style={{ display: "flex", alignItems: "flex-start", marginBottom: 16, gap: 10 }}>
                    {NP_AVATAR}
                    <div style={{ padding: "14px 18px", borderRadius: "5px 20px 20px 20px", background: T.surface, border: `1px solid ${T.border}`, display: "flex", gap: 5, alignItems: "center", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
                      {[0,1,2].map(n => <div key={n} style={{ width: 7, height: 7, borderRadius: "50%", background: T.brand, opacity: 0.35 + n * 0.3 }} />)}
                    </div>
                  </div>
                )}
                <div ref={chatEndRef} />
              </div>
              <div style={{ padding: isMobile ? "12px 16px 24px" : "14px 32px 24px", flexShrink: 0 }}>
                <div style={{ display: "flex", gap: 10, alignItems: "flex-end", background: T.surface, border: `1.5px solid ${T.border}`, borderRadius: 16, padding: "8px 8px 8px 16px", boxShadow: "0 2px 12px rgba(0,0,0,0.06)" }}
                  onFocusCapture={e => { e.currentTarget.style.borderColor = T.borderFocus; e.currentTarget.style.boxShadow = `0 0 0 3px ${T.brandDim}`; }}
                  onBlurCapture={e => { e.currentTarget.style.borderColor = T.border; e.currentTarget.style.boxShadow = "0 2px 12px rgba(0,0,0,0.06)"; }}>
                  <textarea value={chatInput} onChange={e => setChatInput(e.target.value)}
                    onKeyDown={e => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); sendChat(); } }}
                    placeholder="Ask about OKR performance, members, departments…"
                    rows={1}
                    style={{ flex: 1, padding: "6px 0", fontSize: 14, fontFamily: F.body, background: "transparent", border: "none", color: T.text, outline: "none", resize: "none", lineHeight: 1.55 }} />
                  <button onClick={() => sendChat()} disabled={!chatInput.trim() || chatLoading}
                    style={{ width: 36, height: 36, borderRadius: 10, background: chatInput.trim() && !chatLoading ? "linear-gradient(135deg,#0071E3,#6B47DC)" : T.raised, border: "none", cursor: chatInput.trim() && !chatLoading ? "pointer" : "default", color: chatInput.trim() && !chatLoading ? "#fff" : T.textDim, fontSize: 16, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>↑</button>
                </div>
                <div style={{ marginTop: 8, fontSize: 11, color: T.textDim, textAlign: "center" }}>Answers based on live OKR data · Admin only</div>
              </div>
            </div>
          );
        })()}

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
    </MobileContext.Provider>
  );
}

/* ─────────────────────────────────────────────────────────────
   MANAGER PORTAL
   ───────────────────────────────────────────────────────────── */
function ManagerPortal({ user, onLogout, state, dispatch, onReload }) {
  const [page, setPageRaw] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'manager' ? (p[2] || 'dashboard') : 'dashboard';
  });
  const setPage = useCallback(p => { window.history.pushState(null, '', `/manager/${p}`); setPageRaw(p); }, []);
  const isMobile = useIsMobile();
  const [drawerOpen, setDrawerOpen] = useState(false);
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
  useEffect(() => { if (page === "checkin") onReload(); }, [page]); // eslint-disable-line
  const [newProj, setNewProj] = useState({ name: "", startDate: "", due: "" });
  const [showNewProj, setShowNewProj] = useState(false);
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ status: "active", startDate: "", due: "", income: "", margin: "", contributeRate: "" });
  const [progressEdits, setProgressEdits] = useState({});
  const [logDrafts, setLogDrafts] = useState({});
  const [syncPrompt, setSyncPrompt] = useState(null);
  const syncTimerRef = useRef(null);
  const [syncNote, setSyncNote] = useState(null);
  const syncNoteTimer = useRef(null);
  const [mgrDirtySync, setMgrDirtySync] = useState(null);
  const [expandedMonthlyKr, setExpandedMonthlyKr] = useState(null);
  const [mgrSelTeam, setMgrSelTeam] = useState(null);
  const [okrPeriod, setOkrPeriod] = useState("all");
  const [noReason, setNoReason] = useState(null);
  const [yesConfirm, setYesConfirm] = useState(null);
  const [rejectOkr, setRejectOkr] = useState(null);
  const [editingSub, setEditingSub] = useState(null);
  const [editingApproved, setEditingApproved] = useState(null);
  const [mgrKpiPeriodKeys, setMgrKpiPeriodKeys] = useState({});
  const [mgrAssess, setMgrAssess] = useState({});
  const [trackerInput, setTrackerInput] = useState({});
  const [syncing, setSyncing] = useState(false);
  const [logPopup, setLogPopup] = useState(null);
  useEffect(() => { if (!logPopup) return; const h = e => { if (e.key === "Escape") setLogPopup(null); }; window.addEventListener("keydown", h); return () => window.removeEventListener("keydown", h); }, [logPopup]);
  const handleSync = useCallback(async () => { setSyncing(true); await onReload(); setSyncing(false); }, [onReload]);

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
  const peerManagerIds = user.canApprovePeers
    ? users.filter(u => u.role === "manager" && u.deptId === user.deptId && u.id !== user.id).map(u => u.id)
    : [];
  const myOkrSubs = allOkrSubs.filter(s => s.memberId === user.id);
  const myPendingCheckins = myOkrSubs.filter(s => s.answer === null);
  const designatedMemberIds = users.filter(u => u.designatedApproverId === user.id).map(u => u.id);
  const myOkrSubsForApproval = allOkrSubs.filter(s => myTeamMemberIds.includes(s.memberId) || peerManagerIds.includes(s.memberId) || designatedMemberIds.includes(s.memberId));
  const pendingOkrSubs = myOkrSubsForApproval.filter(s => s.answer !== null && s.approval === "pending");
  const myProjects = projects.filter(p => user.deptId ? users.find(u => u.id === p.mgrId)?.deptId === user.deptId : p.mgrId === user.id);
  const dmSubs = allOkrSubs.filter(s => s.answer !== null);
  const myExcludedCount = myMembers.filter(u => u.excludeFromRate).length;
  const myMemberRates = myMembers.filter(u => !u.excludeFromRate).map(u => {
    const kd = memberData[u.id] || { krs: [] };
    if (!memberHasRateKrs(kd.krs)) return null;
    return calcMemberRate(u.id, kd.krs, dmSubs);
  }).filter(r => r !== null);
  const myDeptRate = myMemberRates.length ? myMemberRates.reduce((a, b) => a + b, 0) / myMemberRates.length : 0;

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
    { id: "dashboard",    icon: "⬡", label: "Team Dashboard"       },
    { id: "okr-overview", icon: "⬡", label: "OKR Overview"         },
    { id: "checkin",      icon: "⬡", label: "OKR Check-In"         },
    { id: "approvals",    icon: "⬡", label: "Approve Submissions"   },
    { id: "projects",     icon: "⬡", label: "Projects"             },
    { id: "members",      icon: "⬡", label: "Edit Member KPIs"     },
    { id: "reports",      icon: "⬡", label: "OKR Reports"          },
    ...(user.financeAccess ? [{ id: "financial", icon: "⬡", label: "Financial Performance" }] : []),
  ];

  return (
    <MobileContext.Provider value={{ isMobile, drawerOpen, setDrawerOpen }}>
    <div style={{ display: "flex", minHeight: "100dvh", fontFamily: F.body, background: T.bg, color: T.text }}>
      {logPopup && <div onClick={() => setLogPopup(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}><div onClick={e => e.stopPropagation()} style={{ background: T.surface, borderRadius: 12, border: `1px solid ${T.border}`, width: "100%", maxWidth: 640, maxHeight: "75vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}><div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexShrink: 0 }}><div><div style={{ fontSize: 15, fontWeight: 700 }}>{logPopup.projName}</div>{logPopup.date && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>{logPopup.date}</div>}</div><button onClick={() => setLogPopup(null)} style={{ background: "none", border: "none", cursor: "pointer", color: T.textMuted, fontSize: 18, lineHeight: 1, padding: "0 2px", marginLeft: 12 }}>✕</button></div><div style={{ padding: "16px 20px", overflowY: "auto", fontSize: 14, lineHeight: 1.65, color: T.text, whiteSpace: "pre-wrap" }}>{logPopup.text}</div></div></div>}
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
            const _pkMgr = okrPeriod === "weekly" ? prevPeriodKey(okrPeriod) : currentPeriodKey(okrPeriod);
            const hasSub = okrPeriod === "all" ? (isMonthly ? Object.values(kr.monthlyActuals || {}).some(v => v != null) : kr.actual != null) : allOkrSubs.some(s => s.krId === kr.id && s.period === (kr.period || "monthly") && s.periodKey === _pkMgr && s.answer !== null);
            return (
              <Fragment key={kr.id}>
              <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                <div>
                  <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                  {kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                  {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                  {kr.type === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Progress · affects rate proportionally</span>}
                  {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Monthly Breakdown</span>}
                  {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                </div>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? `${kr.operator||">="} ${fmt(curTarget)} this mo.` : kr.type === "progress" ? fmt(kr.target) : `${kr.operator || ">="} ${fmt(kr.target)}${kr.unit ? ` ${kr.unit}` : ""}`}</span>}
                {kr.type === "tracker"
                  ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textDim }}>—</span>
                  : isMonthly
                  ? <NumInput value={curActual} onChange={n => dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: curKey, field: "actual", value: n })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                  : <NumInput value={kr.actual} onChange={n => { dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "actual", value: n }); if (teamId) mgrTriggerSync(deptId, teamId); }} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />}
                <span style={{ fontSize: 12, color: T.textMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.dataSource || "—"}</span>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 14, fontWeight: 700, color: "#7c3aed" }}>{fmt(isMonthly ? curActual : kr.actual)}{kr.unit ? <span style={{ fontSize: 11, fontWeight: 400 }}> {kr.unit}</span> : ""}</span> : hasSub ? <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[st].color }}>{pct.toFixed(0)}%</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: T.textDim }}>N/A</span>}
                {kr.type === "tracker" ? <span /> : hasSub ? <Bar value={pct} status={st} h={5} /> : <span />}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 4 }}>
                  {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="View all months" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "2px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 11, fontWeight: 700 }}>📅</button>}
                  {kr.type === "tracker" ? null : hasSub ? <Tag type={st} small /> : null}
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
                                <NumInput value={a} onChange={n => { dispatch({ type: "UPDATE_KR_MONTHLY", deptId, teamId, krId: kr.id, monthKey: key, field: "actual", value: n }); if (teamId) mgrTriggerSync(deptId, teamId); }} style={{ padding: "3px 5px", fontSize: 12, fontFamily: F.mono, textAlign: "right", width: "100%", boxSizing: "border-box" }} />
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
                <div style={{ overflowX: "auto" }}><div style={{ minWidth: 760 }}>
                <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                  <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Performance Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Data Source</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                </div>
                {renderRows(krs, deptId, teamId)}
                </div></div>
              </Card>
            );
          };
          const dKrs = filterP(dept.krs);
          const deptStatus = getStatus(myDeptRate);
          const filtMKrs = krs => okrPeriod === "all" ? krs : krs.filter(kr => (kr.period || "monthly") === okrPeriod);
          const teamStats = dept.teams.map(t => {
            const tKrs = filterP(t.krs);
            const teamMembers = myMembers.filter(u => !u.excludeFromRate && (u.teamId === t.id || u.secondTeamId === t.id));
            const memberRates = teamMembers.map(u => {
              const kd = memberData[u.id];
              if (!kd || !memberHasRateKrs(filtMKrs(kd.krs))) return null;
              return calcMemberRate(u.id, filtMKrs(kd.krs), allOkrSubs);
            }).filter(r => r !== null);
            const rate = memberRates.length ? memberRates.reduce((a, b) => a + b, 0) / memberRates.length : null;
            return { ...t, krs: tKrs, rate, status: getStatus(rate) };
          });
          const totalKrs = dKrs.length + teamStats.reduce((s, t) => s + t.krs.length, 0);
          return (<>
            <Header title="OKR Overview" sub={`${dept.name} · ${okrPeriod.charAt(0).toUpperCase() + okrPeriod.slice(1)} key results`} right={<Tag type={deptStatus} />} />
            <Pane>
              <div style={{ display: "flex", gap: 4, marginBottom: 20, paddingBottom: 12, borderBottom: `1px solid ${T.border}` }}>
                {PERIODS.map(p => <Btn key={p.id} primary={okrPeriod === p.id} small onClick={() => setOkrPeriod(p.id)}>{p.label}</Btn>)}
              </div>
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 20 }}>
                <Metric label="Dept Completion" value={`${myDeptRate.toFixed(1)}%`} status={deptStatus} sub={myExcludedCount > 0 ? `${myExcludedCount} excluded · Target: ${TP}%` : `Target: ${TP}%`} />
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
                      <span style={{ fontSize: 14, fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[t.status].color }}>{t.rate != null ? `${t.rate.toFixed(1)}%` : "N/A"}</span>
                      <div style={{ width: 100, flexShrink: 0 }}><Bar value={t.rate ?? 0} status={t.status} h={5} /></div>
                      <Tag type={t.status} small />
                    </div>
                    {t.krs.map(kr => {
                      const pct = krCompletion(kr); const st = getStatus(pct);
                      const trackerVal = (kr.type === "tracker" && kr.actual != null && kr.actual !== 0)
                        ? `${fmt(kr.actual)}${kr.unit ? ` ${kr.unit}` : ""}` : null;
                      const hasSub = kr.type !== "tracker" && (!!kr.monthlyTargets
                        ? Object.values(kr.monthlyActuals || {}).some(v => v != null && v !== 0)
                        : (kr.actual != null && kr.actual !== 0));
                      return (
                      <div key={kr.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0 5px 12px", borderTop: `1px solid ${T.border}`, fontSize: 13 }}>
                        <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim, width: 50, flexShrink: 0 }}>{kr.id}</span>
                        <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.label}</span>
                        {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Tracker</span>}
                        {kr.type === "manager-fill" && <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Mgr Fill</span>}
                        {kr.type === "project_profit" && <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Proj Profit</span>}
                        {kr.type !== "tracker" && kr.type !== "manager-fill" && kr.type !== "project_profit" && kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                        {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                        {kr.type === "tracker"
                          ? <span style={{ fontSize: 12, fontFamily: F.mono, color: trackerVal ? "#7c3aed" : T.textDim, fontWeight: 700, textAlign: "right", flexShrink: 0 }}>{trackerVal ?? "N/A"}</span>
                          : hasSub ? <span style={{ fontSize: 12, fontFamily: F.mono, color: STATUS_THEME[st].color, fontWeight: 700, width: 40, textAlign: "right" }}>{pct.toFixed(0)}%</span>
                          : <span style={{ fontSize: 12, fontFamily: F.mono, color: T.textDim, fontWeight: 700, width: 40, textAlign: "right" }}>N/A</span>}
                        {kr.type === "tracker" ? <span style={{ width: 100, flexShrink: 0 }} /> : hasSub ? <div style={{ width: 100, flexShrink: 0 }}><Bar value={pct} status={st} h={5} /></div> : <span style={{ width: 100, flexShrink: 0 }} />}
                        {kr.type !== "tracker" && hasSub && <Tag type={st} small />}
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
              <Metric label="Dept Completion"   value={`${myDeptRate.toFixed(1)}%`} status={getStatus(myDeptRate)} sub={myExcludedCount > 0 ? `${myExcludedCount} excluded · Time: ${TP}%` : `Time: ${TP}%`} />
              <Metric label="My Members"        value={myMembers.length} />
              <Metric label="Pending Approvals" value={pendingOkrSubs.length} status={pendingOkrSubs.length > 0 ? "yellow" : undefined} />
            </div>
            <SectionLabel>My Team Members</SectionLabel>
            {myMembers.map(m => {
              const kd = memberData[m.id]; if (!kd) return null;
              const hasRateKrs = memberHasRateKrs(kd.krs);
              const r = hasRateKrs ? calcMemberRate(m.id, kd.krs, allOkrSubs) : null;
              const s = getStatus(r);
              const memberProjCount = projects.filter(p => p.mgrId === m.id).length;
              return (
                <Card key={m.id} style={{ marginBottom: 8, overflow: "hidden" }}>
                  <div style={{ overflowX: "auto" }}><div style={{ minWidth: 420, padding: "14px 18px" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "36px 1fr 55px 150px 70px", alignItems: "center", gap: 12 }}>
                    <Avatar letters={m.av} size={30} />
                    <div>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{m.name}</div>
                      <div style={{ display: "flex", gap: 6, alignItems: "center", marginTop: 2 }}>
                        <span style={{ fontSize: 12, color: T.textMuted }}>{m.title}</span>
                        {memberProjCount > 0 && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 6px", fontWeight: 700 }}>◫ {memberProjCount} project{memberProjCount > 1 ? "s" : ""}</span>}
                      </div>
                    </div>
                    <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[s].color }}>{r != null ? `${r.toFixed(1)}%` : "N/A"}</span>
                    <Bar value={r ?? 0} status={s} h={6} />
                    <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={s} small /></div>
                  </div>
                  </div></div>
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
            <Header title="OKR Check-In" sub="Answer your KPI check-ins sent by the system"
              right={<div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                {subRate !== null && <><span style={{ fontSize: 12, color: T.textMuted }}>This month:</span><span style={{ fontWeight: 700, fontSize: 15, color: STATUS_THEME[getStatus(subRate)].color, fontFamily: F.mono }}>{subRate.toFixed(0)}%</span></>}
                <Btn small onClick={handleSync} disabled={syncing}>{syncing ? "Syncing…" : "⟳ Sync"}</Btn>
              </div>} />
            <Pane>
              {myPendingCheckins.length > 0 && (
                <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 8, fontSize: 13, color: T.warn, fontWeight: 600, marginBottom: 16 }}>
                  {myPendingCheckins.length} pending check-in{myPendingCheckins.length !== 1 ? "s" : ""} — please respond below
                </div>
              )}
              {grouped.length === 0 && <EmptyState text="No check-ins yet. Admin will send them when due." />}
              {grouped.map(({ period, pending, answered }) => {
                const byPK = {};
                [...pending, ...answered].forEach(s => {
                  if (!byPK[s.periodKey]) byPK[s.periodKey] = { pk: s.periodKey, dr: s.dateRange || "", p: [], a: [] };
                  if (s.answer === null) byPK[s.periodKey].p.push(s); else byPK[s.periodKey].a.push(s);
                });
                const pkGroups = Object.values(byPK).sort((a, b) =>
                  a.p.length > 0 && b.p.length === 0 ? -1 : a.p.length === 0 && b.p.length > 0 ? 1 : b.pk.localeCompare(a.pk));
                return (
                <div key={period} style={{ marginBottom: 28 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${PERIOD_COLORS[period]}` }}>
                    <div style={{ width: 4, height: 18, background: PERIOD_COLORS[period], borderRadius: 2 }} />
                    <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>{period.charAt(0).toUpperCase() + period.slice(1)} Check-Ins</span>
                    {pending.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 7px", fontSize: 11, fontWeight: 700 }}>{pending.length} pending</span>}
                  </div>
                  {pkGroups.map(({ pk, dr, p: pkPending, a: pkAnswered }) => (
                    <div key={pk} style={{ marginBottom: 14 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 12px", marginBottom: 8, borderRadius: 6, background: pkPending.length > 0 ? T.warnDim : T.raised, border: `1px solid ${pkPending.length > 0 ? T.warnBorder : T.border}` }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: pkPending.length > 0 ? T.warn : T.textMuted }}>{dr || pk}</span>
                        {pkPending.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 6px", fontSize: 10, fontWeight: 700 }}>{pkPending.length} pending</span>}
                        {pkPending.length === 0 && <span style={{ fontSize: 11, color: T.ok }}>✓ All answered</span>}
                      </div>
                  {pkPending.map(s => (
                    <Card key={s.id} style={{ padding: "14px 18px", marginBottom: 8, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : s.krType === "progress" ? T.brand : noReason?.id === s.id ? T.bad : yesConfirm?.id === s.id ? T.ok : T.warn}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                            <span style={{ fontSize: 15, fontWeight: 700 }}>{s.krLabel}</span>
                            {s.krUnit && <span style={{ fontSize: 10, fontWeight: 700, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 6px" }}>{s.krUnit}</span>}
                            {s.krIsMonthly && s.krType !== "tracker" && s.krType !== "progress" && <span style={{ fontSize: 10, fontWeight: 700, background: "#e0f2fe", color: "#0369a1", border: "1px solid #7dd3fc", borderRadius: 5, padding: "1px 6px" }}>Monthly</span>}
                            {s.krType === "tracker" && <span style={{ fontSize: 10, fontWeight: 700, background: "#ede9fe", color: "#6d28d9", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Tracker · does not affect rate</span>}
                            {s.krType === "progress" && <span style={{ fontSize: 10, fontWeight: 700, background: T.brandDim, color: T.brand, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Progress · affects rate proportionally</span>}
                          </div>
                          <div style={{ fontSize: 12, color: T.textMuted }}>
                            {(s.krType !== "tracker" && s.krType !== "progress") && <span>Performance Target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                            {s.krType === "progress" && <span>Target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                            {s.krType === "tracker" && s.krUnit && <span>Unit: {s.krUnit}</span>}
                          </div>
                        </div>
                        {(s.krType === "tracker" || s.krType === "progress") ? (
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            <Input value={trackerInput[s.id] || ""} onChange={e => setTrackerInput(p => ({ ...p, [s.id]: e.target.value }))} placeholder="Enter value" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                            {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                            <Btn primary small onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "submitted", actualValue: Number(trackerInput[s.id]) || 0 }); setTrackerInput(p => ({ ...p, [s.id]: "" })); }} disabled={!trackerInput[s.id] || (s.krType === "tracker" && s.krDisallowZero && Number(trackerInput[s.id]) === 0)}>{s.krType === "progress" ? "Record Progress" : "Record"}</Btn>
                            {s.krType === "tracker" && s.krDisallowZero && trackerInput[s.id] === "0" && <span style={{ fontSize: 11, color: T.bad, whiteSpace: "nowrap" }}>Cannot be 0</span>}
                          </div>
                        ) : (noReason?.id !== s.id && yesConfirm?.id !== s.id) ? (
                          <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                            <button onClick={() => { setNoReason({ id: s.id, reason: "", actual: "" }); setYesConfirm(null); }}
                              style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.bad, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✗ No
                            </button>
                            <button onClick={() => { setYesConfirm({ id: s.id, actual: String(s.krTarget ?? "") }); setNoReason(null); }}
                              style={{ background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.ok, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✓ Yes
                            </button>
                          </div>
                        ) : null}
                      </div>
                      {s.krType !== "tracker" && s.krType !== "progress" && yesConfirm?.id === s.id && (
                        <div style={{ marginTop: 12, padding: "12px 14px", background: T.okDim, borderRadius: 8, border: `1px solid ${T.okBorder}` }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: T.ok, marginBottom: 8 }}>Enter your actual value</div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                            <Input value={yesConfirm.actual} onChange={e => setYesConfirm(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} autoFocus />
                            {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                            <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                          </div>
                          {yesConfirm.actual !== "" && meetsTarget(yesConfirm.actual, s.krOperator, s.krTarget) === false && (
                            <div style={{ fontSize: 12, color: T.bad, fontWeight: 600, marginBottom: 8 }}>⚠ Actual ({yesConfirm.actual}{s.krUnit ? " " + s.krUnit : ""}) doesn't meet target ({s.krOperator || ">="} {s.krTarget}{s.krUnit ? " " + s.krUnit : ""}) — answer should be No</div>
                          )}
                          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                            <Btn small onClick={() => setYesConfirm(null)}>Cancel</Btn>
                            <Btn primary small disabled={yesConfirm.actual !== "" && meetsTarget(yesConfirm.actual, s.krOperator, s.krTarget) === false} onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "yes", actualValue: Number(yesConfirm.actual) || 0 }); setYesConfirm(null); }}>✓ Submit Yes</Btn>
                          </div>
                        </div>
                      )}
                      {s.krType !== "tracker" && s.krType !== "progress" && noReason?.id === s.id && (
                        <div style={{ marginTop: 12, padding: "12px 14px", background: T.badDim, borderRadius: 8, border: `1px solid ${T.badBorder}` }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: T.bad, marginBottom: 8 }}>Why was this OKR not met?</div>
                          <TextArea value={noReason.reason} onChange={e => setNoReason(p => ({ ...p, reason: e.target.value }))} placeholder="Briefly explain why this target was not reached..." rows={2} />
                          <div style={{ marginTop: 10 }}>
                            <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 5 }}>Your actual value</div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <Input value={noReason.actual} onChange={e => setNoReason(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                              {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                              <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                            </div>
                            {noReason.actual !== "" && meetsTarget(noReason.actual, s.krOperator, s.krTarget) === true && (
                              <div style={{ fontSize: 12, color: T.bad, fontWeight: 600, marginTop: 6 }}>⚠ Actual ({noReason.actual}{s.krUnit ? " " + s.krUnit : ""}) meets target ({s.krOperator || ">="} {s.krTarget}{s.krUnit ? " " + s.krUnit : ""}) — answer should be Yes</div>
                            )}
                          </div>
                          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 10 }}>
                            <Btn small onClick={() => setNoReason(null)}>Cancel</Btn>
                            <Btn danger small disabled={noReason.actual !== "" && meetsTarget(noReason.actual, s.krOperator, s.krTarget) === true} onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "no", reason: noReason.reason.trim() || null, actualValue: Number(noReason.actual) || 0 }); setNoReason(null); }}>Submit No</Btn>
                          </div>
                        </div>
                      )}
                    </Card>
                  ))}
                  {pkAnswered.length > 0 && (
                    <div style={{ marginTop: pkPending.length > 0 ? 6 : 0 }}>
                      {pkAnswered.slice(0, 10).map(s => (
                        <Card key={s.id} style={{ padding: "10px 14px", marginBottom: 4, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : s.krType === "progress" ? T.brand : s.answer === "yes" ? T.ok : T.bad}`, opacity: s.approval === "approved" ? 0.7 : 1 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div>
                              <span style={{ fontSize: 13, fontWeight: 600 }}>{s.krLabel}</span>
                              <span style={{ fontSize: 11, color: T.textMuted, marginLeft: 8 }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              {s.krType === "tracker"
                                ? <span style={{ fontSize: 12, fontWeight: 700, color: "#6d28d9" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                                : s.krType === "progress"
                                ? <span style={{ fontSize: 12, fontWeight: 700, color: T.brand }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}{s.krTarget ? ` (${Math.min(Math.round((Number(s.actualValue || 0) / Number(s.krTarget)) * 100), 100)}%)` : ""}</span>
                                : <div style={{ textAlign: "right" }}>
                                    <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>
                                    {s.actualValue != null && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</div>}
                                  </div>}
                              <Tag type={s.approval === "approved" ? "approved" : s.approval === "rejected" ? "rejected" : "pending"} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                              {s.approvedBy === "auto" && <span style={{ fontSize: 9, fontWeight: 700, background: "rgba(16,185,129,0.1)", border: "1px solid rgba(16,185,129,0.25)", color: "#10B981", borderRadius: 8, padding: "1px 6px", letterSpacing: "0.05em" }}>AUTO</span>}
                              {s.approvedBy && s.approvedBy !== "auto" && <span style={{ fontSize: 10, color: T.textMuted }}>by {users?.find(u => u.id === s.approvedBy)?.name || "Admin"}</span>}
                            </div>
                          </div>
                          {s.krType !== "tracker" && s.answer === "no" && s.reason && <div style={{ fontSize: 12, color: T.textSoft, marginTop: 5, paddingTop: 5, borderTop: `1px solid ${T.border}` }}>Note: {s.reason}</div>}
                        </Card>
                      ))}
                    </div>
                  )}
                    </div>
                  ))}
                </div>
                );
              })}
            </Pane>
          </>);
        })()}

        {page === "approvals" && (() => {
          const totalPending = pendingOkrSubs.length;
          return (<>
            <Header title="Approve Member Submissions" sub={`${totalPending} pending review`} />
            <Pane>
              {/* OKR Check-In submissions */}
              {myOkrSubsForApproval.length > 0 && (
                <div style={{ marginBottom: 28 }}>
                  <div style={{ fontSize: 14, fontWeight: 700, color: T.text, marginBottom: 10, display: "flex", alignItems: "center", gap: 8 }}>
                    OKR Check-Ins
                    {pendingOkrSubs.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 7px", fontSize: 11, fontWeight: 700 }}>{pendingOkrSubs.length} pending</span>}
                  </div>
                  {(() => {
                    const answeredSubs = myOkrSubsForApproval.filter(s => s.answer !== null && !s.managerFilled).sort((a,b) => { const o={pending:0,approved:1,rejected:2}; return o[a.approval]-o[b.approval]||(b.answeredAt||"").localeCompare(a.answeredAt||""); });
                    const order = [];
                    const groups = {};
                    answeredSubs.forEach(s => {
                      if (!groups[s.memberId]) { groups[s.memberId] = []; order.push(s.memberId); }
                      groups[s.memberId].push(s);
                    });
                    return order.map(memberId => {
                      const subs = groups[memberId];
                      const mem = users.find(u => u.id === memberId);
                      const pendingCount = subs.filter(s => s.approval === "pending").length;
                      return (
                        <Card key={memberId} style={{ marginBottom: 10, overflow: "hidden" }}>
                          {/* Person header */}
                          <div style={{ padding: "11px 16px", background: T.raised, borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                            <Avatar letters={mem?.av || "?"} size={28} />
                            <div style={{ flex: 1, minWidth: 0 }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                <span style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{mem?.name || subs[0]?.memberName || "Unknown"}</span>
                                {designatedMemberIds.includes(memberId) && !myTeamMemberIds.includes(memberId) && !peerManagerIds.includes(memberId) && (
                                  <span style={{ fontSize: 9, fontWeight: 700, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 6px", letterSpacing: "0.05em" }}>DESIGNATED</span>
                                )}
                              </div>
                              <div style={{ fontSize: 11, color: T.textDim, marginTop: 1 }}>{subs.length} KR{subs.length !== 1 ? "s" : ""}</div>
                            </div>
                            {pendingCount > 0 && (
                              <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                                <span style={{ background: T.warnDim, color: T.warn, border: `1px solid ${T.warnBorder}`, borderRadius: 6, padding: "2px 9px", fontSize: 11, fontWeight: 700 }}>{pendingCount} pending</span>
                                <Btn primary small onClick={() => subs.filter(s => s.approval === "pending").forEach(s => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id }))}>Approve All</Btn>
                              </div>
                            )}
                          </div>
                          {/* KR rows */}
                          {subs.map((s, idx) => {
                            const accentColor = s.approval === "approved" ? T.ok : s.approval === "rejected" ? T.bad : T.warn;
                            const isLast = idx === subs.length - 1 && rejectOkr?.id !== s.id && editingSub?.id !== s.id && editingApproved?.id !== s.id;
                            return (
                              <div key={s.id} style={{ borderBottom: isLast ? "none" : `1px solid ${T.border}` }}>
                                <div style={{ padding: "10px 16px 10px 19px", borderLeft: `3px solid ${accentColor}`, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                                  <div style={{ flex: 1, minWidth: 0 }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 2 }}>
                                      <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>{s.krLabel}</span>
                                      {s.krType === "tracker" && <span style={{ fontSize: 10, color: "#6d28d9", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Tracker · no rate impact</span>}
                                      {s.krType === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Progress · affects rate</span>}
                                    </div>
                                    <div style={{ fontSize: 11, color: T.textMuted }}>
                                      <span style={{ fontSize: 11, color: T.textMuted, background: T.surface, border: `1px solid ${T.border}`, borderRadius: 4, padding: "0px 5px", marginRight: 6 }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                                      {s.krType !== "tracker" && s.krType !== "progress" && <>{`Performance Target: ${s.krOperator || ">="} ${s.krTarget}${s.krUnit ? ` ${s.krUnit}` : ""}`}</>}
                                      {s.krType === "progress" && <span>Target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                      {s.krType === "tracker" && s.actualValue != null && <span style={{ color: "#6d28d9", fontWeight: 700 }}>Recorded: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                      {s.krType === "progress" && s.actualValue != null && <span style={{ color: T.brand, marginLeft: 8, fontWeight: 700 }}>Recorded: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}{s.krTarget ? ` (${Math.min(Math.round((Number(s.actualValue) / Number(s.krTarget)) * 100), 100)}%)` : ""}</span>}
                                      {s.krType !== "tracker" && s.krType !== "progress" && s.answer === "no" && s.actualValue != null && <span style={{ color: T.bad, marginLeft: 8, fontWeight: 700 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                      {s.krType !== "tracker" && s.krType !== "progress" && s.answer === "yes" && <span style={{ color: T.ok, marginLeft: 8 }}>Actual: {s.actualValue ?? s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                      <span style={{ marginLeft: 8 }}>· Answered: {s.answeredAt?.slice(0,10) || "—"}</span>
                                    </div>
                                    {s.krType !== "tracker" && s.krType !== "progress" && s.answer === "no" && s.reason && <div style={{ fontSize: 11, color: T.bad, marginTop: 2, fontStyle: "italic" }}>Reason: {s.reason}</div>}
                                  </div>
                                  <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                                    {s.krType === "tracker"
                                      ? <span style={{ fontSize: 11, fontWeight: 700, background: "#ede9fe", color: "#6d28d9", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 6px" }}>Tracker</span>
                                      : s.krType === "progress"
                                      ? <span style={{ fontSize: 11, fontWeight: 700, background: T.brandDim, color: T.brand, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 6px" }}>Progress</span>
                                      : <div style={{ textAlign: "right" }}>
                                    <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>
                                    {s.actualValue != null && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</div>}
                                  </div>}
                                    {s.approval === "pending"
                                      ? <div style={{ display: "flex", gap: 6 }}>
                                          <Btn small onClick={() => { setEditingSub({ id: s.id, answer: s.answer, actual: s.actualValue != null ? String(s.actualValue) : "" }); setRejectOkr(null); }}>✎ Edit</Btn>
                                          <Btn danger small onClick={() => { setRejectOkr({ id: s.id, actual: "" }); setEditingSub(null); }}>Reject</Btn>
                                          <Btn primary small onClick={() => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id })}>Approve</Btn>
                                        </div>
                                      : <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                          <Tag type={s.approval === "approved" ? "approved" : "rejected"} label={s.approval === "approved" ? "Approved" : "Rejected"} small />
                                          {s.approvedBy && s.approvedBy !== "auto" && <span style={{ fontSize: 11, color: T.textMuted }}>by {users?.find(u => u.id === s.approvedBy)?.name || "Admin"}</span>}
                                          <Btn small onClick={() => { setEditingApproved({ id: s.id, actual: s.actualValue != null ? String(s.actualValue) : "", answer: s.answer }); setEditingSub(null); setRejectOkr(null); }}>✎</Btn>
                                        </div>}
                                  </div>
                                </div>
                                {rejectOkr?.id === s.id && (
                                  <div style={{ margin: "0 16px 10px 19px", padding: "10px 12px", background: T.badDim, borderRadius: 7, border: `1px solid ${T.badBorder}` }}>
                                    <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 6 }}>Enter actual value for rejection</div>
                                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                                      <Input value={rejectOkr.actual} onChange={e => setRejectOkr(p => ({ ...p, actual: e.target.value }))} placeholder="Actual value" style={{ width: 120, textAlign: "right", fontFamily: F.mono }} />
                                      {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                      {s.krType === "progress" ? <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span> : <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                    </div>
                                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                      <Btn small onClick={() => setRejectOkr(null)}>Cancel</Btn>
                                      <Btn danger small onClick={() => { dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "rejected", approvedBy: user.id, actualValue: Number(rejectOkr.actual) || 0 }); setRejectOkr(null); }}>Confirm Reject</Btn>
                                    </div>
                                  </div>
                                )}
                                {editingSub?.id === s.id && (
                                  <div style={{ margin: "0 16px 10px 19px", padding: "12px 14px", background: T.raised, borderRadius: 8, border: `1px solid ${T.border}` }}>
                                    <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 10 }}>Edit Submission</div>
                                    {s.krType !== "tracker" && s.krType !== "progress" ? (<>
                                      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                                        <button onClick={() => setEditingSub(p => ({ ...p, answer: "yes", actual: String(s.krTarget ?? "") }))} style={{ background: editingSub.answer === "yes" ? T.okDim : T.surface, border: `1px solid ${editingSub.answer === "yes" ? T.okBorder : T.border}`, borderRadius: 7, padding: "7px 18px", cursor: "pointer", color: editingSub.answer === "yes" ? T.ok : T.textMuted, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>✓ Yes</button>
                                        <button onClick={() => setEditingSub(p => ({ ...p, answer: "no" }))} style={{ background: editingSub.answer === "no" ? T.badDim : T.surface, border: `1px solid ${editingSub.answer === "no" ? T.badBorder : T.border}`, borderRadius: 7, padding: "7px 18px", cursor: "pointer", color: editingSub.answer === "no" ? T.bad : T.textMuted, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>✗ No</button>
                                      </div>
                                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                                        <span style={{ fontSize: 12, color: T.textMuted }}>Actual value:</span>
                                        <Input value={editingSub.actual} onChange={e => setEditingSub(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                                        {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                        <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                                      </div>
                                    </>) : (
                                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
                                        <span style={{ fontSize: 12, color: T.textMuted }}>{s.krType === "progress" ? "Recorded value:" : "Recorded value:"}</span>
                                        <Input value={editingSub.actual} onChange={e => setEditingSub(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                                        {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                        {s.krType === "progress" && <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                      </div>
                                    )}
                                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                      <Btn small onClick={() => setEditingSub(null)}>Cancel</Btn>
                                      <Btn primary small onClick={() => { const newAnswer = (s.krType === "tracker" || s.krType === "progress") ? "submitted" : editingSub.answer; const newActual = Number(editingSub.actual) || 0; dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: newAnswer, actualValue: newActual }); setEditingSub(null); }} disabled={s.krType === "tracker" && s.krDisallowZero && Number(editingSub?.actual) === 0}>Save Changes</Btn>
                                    </div>
                                  </div>
                                )}
                                {editingApproved?.id === s.id && (
                                  <div style={{ margin: "0 16px 10px 19px", padding: "10px 12px", background: T.raised, borderRadius: 7, border: `1px solid ${T.border}` }}>
                                    <div style={{ fontSize: 12, fontWeight: 700, color: T.text, marginBottom: 8 }}>Edit Submission — {s.approval === "approved" ? "Approved" : "Rejected"}</div>
                                    {s.krType !== "tracker" && s.krType !== "progress" && (
                                      <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
                                        <button onClick={() => setEditingApproved(p => ({ ...p, answer: "yes" }))} style={{ background: editingApproved.answer === "yes" ? T.okDim : T.surface, border: `1px solid ${editingApproved.answer === "yes" ? T.okBorder : T.border}`, borderRadius: 7, padding: "5px 16px", cursor: "pointer", color: editingApproved.answer === "yes" ? T.ok : T.textMuted, fontSize: 13, fontWeight: 700, fontFamily: F.body }}>✓ Yes</button>
                                        <button onClick={() => setEditingApproved(p => ({ ...p, answer: "no" }))} style={{ background: editingApproved.answer === "no" ? T.badDim : T.surface, border: `1px solid ${editingApproved.answer === "no" ? T.badBorder : T.border}`, borderRadius: 7, padding: "5px 16px", cursor: "pointer", color: editingApproved.answer === "no" ? T.bad : T.textMuted, fontSize: 13, fontWeight: 700, fontFamily: F.body }}>✗ No</button>
                                      </div>
                                    )}
                                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                                      <span style={{ fontSize: 12, color: T.textMuted }}>Actual value:</span>
                                      <Input value={editingApproved.actual} onChange={e => setEditingApproved(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                                      {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                      {s.krType !== "tracker" && s.krType !== "progress" && <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                      {s.krType === "progress" && <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                    </div>
                                    <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                      <Btn small onClick={() => setEditingApproved(null)}>Cancel</Btn>
                                      <Btn primary small onClick={() => { dispatch({ type: "EDIT_APPROVED_SUBMISSION", id: s.id, actualValue: Number(editingApproved.actual) || 0, answer: editingApproved.answer }); setEditingApproved(null); }}>Save</Btn>
                                    </div>
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </Card>
                      );
                    });
                  })()}
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
            {myProjects.length > 0 && (
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 6 }}>
                <Metric label="Total" value={myProjects.length} />
                <Metric label="Active" value={myProjects.filter(p => p.status === "active").length} status="yellow" />
                <Metric label="Pending Approval" value={myProjects.filter(p => p.status === "pending approval").length} status="blue" />
                <Metric label="Completed" value={myProjects.filter(p => p.status === "completed").length} status="green" />
                <Metric label="Avg Progress" value={`${Math.round(myProjects.reduce((a, p) => a + p.progress, 0) / myProjects.length)}%`} />
                {(() => { const ti = myProjects.filter(p => p.status !== "completed").reduce((a, p) => a + (p.income || 0), 0); return ti > 0 ? <Metric label="Estimated Income" value={`$${ti.toLocaleString()}`} status="blue" /> : null; })()}
                {(() => { const tp = myProjects.filter(p => p.status !== "completed").reduce((a, p) => a + (p.income != null && p.margin != null ? Math.round(p.income * p.margin / 100) : 0), 0); return tp > 0 ? <Metric label="Estimated Profit" value={`$${tp.toLocaleString()}`} status="blue" /> : null; })()}
                {(() => { const ti = myProjects.filter(p => p.status === "completed").reduce((a, p) => a + (p.income || 0), 0); return ti > 0 ? <Metric label="Completed Income" value={`$${ti.toLocaleString()}`} status="green" /> : null; })()}
                {(() => { const tp = myProjects.filter(p => p.status === "completed").reduce((a, p) => a + (p.income != null && p.margin != null ? Math.round(p.income * p.margin / 100) : 0), 0); return tp > 0 ? <Metric label="Completed Profit" value={`$${tp.toLocaleString()}`} status="green" /> : null; })()}
              </div>
            )}
            {user.projectAccess && (
              <div style={{ marginBottom: 16 }}>
                {!showNewProj ? (
                  <Btn primary onClick={() => setShowNewProj(true)}>+ New Project</Btn>
                ) : (
                  <Card style={{ padding: 16 }}>
                    <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>New Project</div>
                    <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
                      <div style={{ flex: 1, minWidth: 180 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Name</div>
                        <Input value={newProj.name} onChange={e => setNewProj(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Term 2 NAPLAN Prep" style={{ width: "100%" }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Start Date</div>
                        <Input type="date" value={newProj.startDate} onChange={e => setNewProj(p => ({ ...p, startDate: e.target.value }))} style={{ width: 160 }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                        <Input type="date" value={newProj.due} onChange={e => setNewProj(p => ({ ...p, due: e.target.value }))} style={{ width: 160 }} />
                      </div>
                      <Btn primary disabled={!newProj.name.trim() || !newProj.startDate} onClick={() => { dispatch({ type: "ADD_PROJECT", project: { id: `p${Date.now()}`, mgrId: user.id, name: newProj.name.trim(), status: "active", startDate: newProj.startDate, due: newProj.due || "TBD", progress: 0, income: null, margin: null } }); setNewProj({ name: "", startDate: "", due: "" }); setShowNewProj(false); }}>Create</Btn>
                      <Btn onClick={() => { setNewProj({ name: "", startDate: "", due: "" }); setShowNewProj(false); }}>Cancel</Btn>
                    </div>
                  </Card>
                )}
              </div>
            )}
            {myProjects.length === 0 && <EmptyState text={user.projectAccess ? "No projects yet. Click '+ New Project' to get started." : "No projects assigned."} />}
            {myProjects.map(p => {
              const draftProg = progressEdits[p.id] ?? p.progress;
              const ps = draftProg >= 70 ? "green" : draftProg >= 35 ? "yellow" : "red";
              const progChanged = progressEdits[p.id] !== undefined;
              const isDetailsOpen = editProjId === p.id;
              return (
                <Card key={p.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                  <div style={{ padding: "12px 18px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{p.name}</div>
                      <div style={{ fontSize: 12, color: T.textMuted }}>{p.mgrId !== user.id && (() => { const owner = users.find(u => u.id === p.mgrId); return owner ? <span>Owner: {owner.name} · </span> : null; })()}{p.startDate ? `Start: ${p.startDate} · ` : ""}Due: {p.due}{p.updatedDate ? ` · Updated: ${p.updatedDate}` : ""}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <Tag type={p.status === "active" ? "pending" : p.status === "pending approval" ? "review" : "approved"} label={p.status === "active" ? "ACTIVE" : p.status === "pending approval" ? "PENDING APPROVAL" : "COMPLETED"} small />
                      {user.projectAccess && p.mgrId === user.id && <button onClick={() => { if (window.confirm(`Delete project "${p.name}"? This cannot be undone.`)) dispatch({ type: "REMOVE_PROJECT", projectId: p.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete project">✕</button>}
                    </div>
                  </div>
                  <div style={{ padding: "10px 18px", display: "flex", alignItems: "center", gap: 10, borderTop: `1px solid ${T.border}` }}>
                    <Bar value={draftProg} status={ps} h={6} />
                    {p.mgrId === user.id ? <>
                      <Input value={draftProg} onChange={e => setProgressEdits(d => ({ ...d, [p.id]: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 52, textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                      <span style={{ fontSize: 13, color: T.textMuted }}>%</span>
                      {p.income != null && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Income: ${p.income.toLocaleString()}</span>}
                      {p.income != null && p.margin != null && <span style={{ fontSize: 11, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Profit: ${Math.round(p.income * p.margin / 100).toLocaleString()} ({p.margin}%)</span>}
                      <Btn primary small disabled={!progChanged} onClick={() => {
                        if (draftProg === 100 && p.status === "active") {
                          if (!window.confirm("Progress is now 100%. Mark as completed?")) {
                            setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                            return;
                          }
                          const updDate = new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { progress: draftProg, updatedDate: updDate } });
                          setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                          const submit = window.confirm("Submit for System Admin approval?");
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: submit ? "pending approval" : "active" } });
                        } else {
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { progress: draftProg, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                        }
                      }}>Save</Btn>
                    </> : <>
                      <span style={{ fontSize: 13, color: T.textMuted, fontFamily: F.mono }}>{draftProg}%</span>
                      {p.income != null && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Income: ${p.income.toLocaleString()}</span>}
                      {p.income != null && p.margin != null && <span style={{ fontSize: 11, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Profit: ${Math.round(p.income * p.margin / 100).toLocaleString()} ({p.margin}%)</span>}
                    </>}
                  </div>
                  {!isDetailsOpen && (() => { const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []); const latest = entries[0]; if (!latest) return null; const truncated = latest.text.length > 160; const preview = truncated ? latest.text.slice(0, 160) + "…" : latest.text; return <div style={{ padding: "6px 18px 4px", fontSize: 13, color: T.textSoft, lineHeight: 1.5 }}>{latest.date && <span style={{ fontSize: 11, color: T.textMuted, marginRight: 6 }}>{latest.date}</span>}{preview}{truncated && <button onClick={e => { e.stopPropagation(); setLogPopup({ text: latest.text, date: latest.date, projName: p.name }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, padding: "0 0 0 4px", fontFamily: F.body }}>Read more →</button>}</div>; })()}
                  {user.projectAccess && p.mgrId === user.id && (
                    <div style={{ padding: "6px 18px 10px" }}>
                      <button onClick={() => {
                        if (isDetailsOpen) { setEditProjId(null); return; }
                        setEditProjId(p.id);
                        setEditProjForm({ status: p.status, startDate: p.startDate || "", due: p.due || "", income: p.income != null ? String(p.income) : "", margin: p.margin != null ? String(p.margin) : "", contributeRate: p.contributeRate != null ? String(p.contributeRate) : "" });
                      }} style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 10px", cursor: "pointer", color: T.textMuted, fontSize: 12, fontFamily: F.body }}>
                        {isDetailsOpen ? "▼ Edit Details" : "▸ Edit Details"}
                      </button>
                    </div>
                  )}
                  {isDetailsOpen && (
                    <div style={{ padding: "14px 18px", borderTop: `1px solid ${T.border}` }}>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                          <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 14, fontFamily: F.body }}>
                            <option value="active">Active</option>
                            <option value="pending approval">Pending Approval</option>
                            <option value="completed">Completed</option>
                          </select>
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Start Date</div>
                          <Input type="date" value={editProjForm.startDate} onChange={e => setEditProjForm(f => ({ ...f, startDate: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                          <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                        </div>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Income ($)</div>
                          <Input type="number" value={editProjForm.income} onChange={e => setEditProjForm(f => ({ ...f, income: e.target.value }))} placeholder="e.g. 250000" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Profit Margin (%)</div>
                          <Input type="number" value={editProjForm.margin} onChange={e => setEditProjForm(f => ({ ...f, margin: e.target.value }))} placeholder="e.g. 30" min="0" max="100" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Contribution Rate (%)</div>
                          <Input type="number" value={editProjForm.contributeRate} onChange={e => setEditProjForm(f => ({ ...f, contributeRate: e.target.value }))} placeholder="100 (default)" min="1" max="100" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                          <div style={{ fontSize: 10, color: T.textMuted, marginTop: 3 }}>Leave blank if sole owner</div>
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Profit (Auto)</div>
                          {(() => {
                            const hasData = editProjForm.income !== "" && editProjForm.margin !== "";
                            const fullProfit = hasData ? Math.round(Number(editProjForm.income) * Number(editProjForm.margin) / 100) : null;
                            const rate = editProjForm.contributeRate !== "" ? Math.min(100, Math.max(1, Number(editProjForm.contributeRate))) : 100;
                            const myProfit = fullProfit != null ? Math.round(fullProfit * rate / 100) : null;
                            return (
                              <div style={{ padding: "7px 10px", fontSize: 14, fontFamily: F.mono, color: hasData ? T.ok : T.textMuted, fontWeight: 700 }}>
                                {fullProfit != null ? `$${fullProfit.toLocaleString()}` : "—"}
                                {fullProfit != null && rate < 100 && <div style={{ fontSize: 11, color: T.brand, fontWeight: 600, marginTop: 2 }}>Your KR: ${myProfit.toLocaleString()} ({rate}%)</div>}
                              </div>
                            );
                          })()}
                        </div>
                      </div>
                      <div style={{ marginBottom: 16 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>Project Logbook</div>
                        <TextArea value={logDrafts[p.id] || ""} onChange={e => setLogDrafts(d => ({ ...d, [p.id]: e.target.value }))} placeholder="Add a log entry..." rows={2} />
                        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 6, marginBottom: 10 }}>
                          <Btn primary small disabled={!logDrafts[p.id]?.trim()} onClick={() => {
                            const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []);
                            dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { log: [{ text: logDrafts[p.id].trim(), date: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) }, ...entries] } });
                            setLogDrafts(d => { const n = { ...d }; delete n[p.id]; return n; });
                          }}>Add Entry</Btn>
                        </div>
                        {(() => { const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []); if (!entries.length) return <div style={{ fontSize: 12, color: T.textMuted }}>No log entries yet.</div>; return entries.map((e, i) => <div key={i} style={{ padding: "8px 10px", marginBottom: 6, background: T.bg, borderRadius: 6, border: `1px solid ${T.border}` }}>{e.date && <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 3 }}>{e.date}</div>}<div style={{ fontSize: 13, color: T.text, lineHeight: 1.5, whiteSpace: "pre-wrap", maxHeight: 200, overflowY: "auto" }}>{e.text}</div></div>); })()}
                      </div>
                      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                        <Btn small onClick={() => setEditProjId(null)}>Cancel</Btn>
                        <Btn primary small onClick={() => {
                          const cr = editProjForm.contributeRate !== "" ? Math.min(100, Math.max(1, Number(editProjForm.contributeRate))) : null;
                          if (editProjForm.status === "completed" && p.status !== "completed") {
                            const submit = window.confirm("Submit for System Admin approval?");
                            dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: submit ? "pending approval" : "active", startDate: editProjForm.startDate || p.startDate || "", due: editProjForm.due || p.due, income: editProjForm.income !== "" ? Number(editProjForm.income) : null, margin: editProjForm.margin !== "" ? Math.min(100, Math.max(0, Number(editProjForm.margin))) : null, contributeRate: cr, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          } else {
                            dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: editProjForm.status, startDate: editProjForm.startDate || p.startDate || "", due: editProjForm.due || p.due, income: editProjForm.income !== "" ? Number(editProjForm.income) : null, margin: editProjForm.margin !== "" ? Math.min(100, Math.max(0, Number(editProjForm.margin))) : null, contributeRate: cr, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          }
                          setEditProjId(null);
                        }}>Save Details</Btn>
                      </div>
                    </div>
                  )}
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "members" && (() => {
          const PERIOD_ORDER = ["daily","weekly","monthly","quarterly","biannual","annual"];
          const getPK = p => mgrKpiPeriodKeys[p] || currentPeriodKey(p);
          const isCurPK = p => getPK(p) === currentPeriodKey(p);
          const periodLabel = p => p.charAt(0).toUpperCase() + p.slice(1);
          const KCOL = "50px 1fr 100px 110px 60px 130px";
          return (<>
            <Header title="Edit Member KPIs" sub="Enter or override KPI actuals for your team — values are saved as approved submissions" />
            <Pane>
              {myMembers.filter(m => m.role === "member").map(m => {
                const kd = memberData[m.id];
                const krs = kd?.krs || [];
                const hasRateKrs = memberHasRateKrs(krs);
                const r = hasRateKrs ? calcMemberRate(m.id, krs, allOkrSubs) : null;
                const s = getStatus(r);
                const groups = {};
                krs.filter(kr => kr.type !== "manager-fill" && kr.type !== "project_profit").forEach(kr => { const p = kr.period || "monthly"; if (!groups[p]) groups[p] = []; groups[p].push(kr); });
                const mgrFillKrs = krs.filter(kr => kr.type === "manager-fill");
                const ppKrsTeam = krs.filter(kr => kr.type === "project_profit");
                const sortedPeriods = Object.keys(groups).sort((a, b) => PERIOD_ORDER.indexOf(a) - PERIOD_ORDER.indexOf(b));
                return (
                  <Card key={m.id} style={{ marginBottom: 16, overflow: "hidden" }}>
                    <div style={{ padding: "14px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}><Avatar letters={m.av} size={30} /><div><div style={{ fontSize: 15, fontWeight: 700 }}>{m.name}</div><div style={{ fontSize: 12, color: T.textMuted }}>{m.title}</div></div></div>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}><span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[s].color }}>{r != null ? `${r.toFixed(1)}%` : "N/A"}</span><Tag type={s} /></div>
                    </div>
                    {krs.length === 0
                      ? <div style={{ padding: "14px 18px", fontSize: 13, color: T.textMuted }}>No KPI data for this member yet. Sync team KPIs to populate.</div>
                      : sortedPeriods.map(period => {
                          const periodKrs = groups[period];
                          const pk = getPK(period);
                          return (
                            <div key={period}>
                              <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 18px", background: T.raised, borderBottom: `1px solid ${T.border}` }}>
                                <span style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, textTransform: "uppercase", letterSpacing: "0.07em", flex: 1 }}>{periodLabel(period)} KRs</span>
                                <span style={{ fontSize: 11, color: T.brand, fontFamily: F.mono, fontWeight: 600 }}>{pk}</span>
                                <Btn small primary={!isCurPK(period)} onClick={() => setMgrKpiPeriodKeys(prev => ({ ...prev, [period]: prevPeriodKey(period) }))}>← Prev</Btn>
                                <Btn small primary={isCurPK(period)} onClick={() => setMgrKpiPeriodKeys(prev => ({ ...prev, [period]: currentPeriodKey(period) }))}>Current</Btn>
                              </div>
                              <div style={{ overflowX: "auto" }}><div style={{ minWidth: 560 }}>
                              <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "5px 18px", gap: 8, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase", borderBottom: `1px solid ${T.border}` }}>
                                <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span>
                              </div>
                              {periodKrs.map((kr, ki) => {
                                if (kr.monthlyTargets) {
                                  const mk = currentFYMonthKey();
                                  const mTgt = kr.monthlyTargets[mk] || 0;
                                  const mAct = (kr.monthlyActuals || {})[mk];
                                  const cr = krCompletion(kr); const cs = getStatus(cr);
                                  return (
                                    <div key={kr.id} style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 18px", gap: 8, alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                                      <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                                      <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
                                        <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.label}</span>
                                        <span style={{ fontSize: 10, flexShrink: 0, color: "#0369a1", background: "#e0f2fe", border: "1px solid #7dd3fc", borderRadius: 8, padding: "1px 5px" }}>Monthly</span>
                                      </div>
                                      <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(mTgt)}</span>
                                      <NumInput value={mAct} placeholder="—" onChange={n => dispatch({ type: "UPDATE_MEMBER_KR", memberId: m.id, krId: kr.id, field: "monthlyActuals", value: { ...(kr.monthlyActuals || {}), [mk]: n } })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                                      <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span>
                                      <Bar value={cr} status={cs} h={5} />
                                    </div>
                                  );
                                }
                                const sub = allOkrSubs.find(s2 => s2.memberId === m.id && s2.krId === kr.id && s2.periodKey === pk && s2.answer !== null);
                                const av = sub != null ? sub.actualValue : null;
                                const cr = av !== null ? krCompletion({ ...kr, actual: av }) : null;
                                const cs = getStatus(cr);
                                return (
                                  <div key={kr.id} style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 18px", gap: 8, alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                                    <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                                    <div style={{ display: "flex", alignItems: "center", gap: 6, minWidth: 0 }}>
                                      <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.label}</span>
                                      {sub && sub.approval === "approved" && <span style={{ fontSize: 10, flexShrink: 0, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px" }}>Approved</span>}
                                      {sub && sub.approval === "pending" && <span style={{ fontSize: 10, flexShrink: 0, color: T.warn, background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 8, padding: "1px 5px" }}>Submitted</span>}
                                    </div>
                                    <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)}</span>
                                    <NumInput value={av} placeholder="—" onChange={n => dispatch({ type: "MANAGER_SUBMIT_KR", memberId: m.id, memberName: m.name, deptId: m.deptId, kr, period, periodKey: pk, actualValue: n, approvedBy: user.id, newId: `ms_${m.id}_${kr.id}_${Date.now().toString(36)}` })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                                    {cr !== null ? <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span> : <span style={{ textAlign: "right", fontSize: 12, color: T.textDim }}>—</span>}
                                    {cr !== null ? <Bar value={cr} status={cs} h={5} /> : <span />}
                                  </div>
                                );
                              })}
                              </div></div>
                            </div>
                          );
                        })
                    }
                    {mgrFillKrs.length > 0 && (
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 18px", background: "#fffbeb", borderBottom: `1px solid #fde68a`, borderTop: `1px solid ${T.border}` }}>
                          <span style={{ fontSize: 11, fontWeight: 700, color: "#d97706", textTransform: "uppercase", letterSpacing: "0.07em", flex: 1 }}>Manager Assessments</span>
                          <span style={{ fontSize: 11, color: "#d97706" }}>Manager fills these — member does not receive check-ins</span>
                        </div>
                        {mgrFillKrs.map((kr, ki) => {
                          const p = kr.period || "monthly";
                          const pk = getPK(p);
                          const existingSub = allOkrSubs.find(s => s.memberId === m.id && s.krId === kr.id && s.periodKey === pk && s.managerFilled);
                          const aKey = `${m.id}:${kr.id}:${pk}`;
                          const aState = mgrAssess[aKey] || { answer: existingSub?.answer || null, value: existingSub?.actualValue != null ? String(existingSub.actualValue) : "" };
                          return (
                            <div key={kr.id} style={{ padding: "12px 18px", borderBottom: `1px solid ${T.border}`, background: ki % 2 ? T.raised : "transparent" }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                                <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim }}>{kr.id}</span>
                                <span style={{ fontWeight: 600, fontSize: 14, flex: 1 }}>{kr.label}</span>
                                <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 5px" }}>Mgr Fill</span>
                                <span style={{ fontSize: 11, color: T.textMuted, fontFamily: F.mono }}>{kr.operator || ">="} {fmt(kr.target)}{kr.unit ? ` ${kr.unit}` : ""}</span>
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
                                <button
                                  onClick={() => setMgrAssess(p2 => ({ ...p2, [aKey]: { ...aState, answer: aState.answer === "yes" ? null : "yes" } }))}
                                  style={{ padding: "5px 14px", borderRadius: 6, border: `1px solid ${aState.answer === "yes" ? "#16a34a" : T.border}`, background: aState.answer === "yes" ? "#dcfce7" : T.surface, color: aState.answer === "yes" ? "#16a34a" : T.textMuted, fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                                  ✓ Yes
                                </button>
                                <button
                                  onClick={() => setMgrAssess(p2 => ({ ...p2, [aKey]: { ...aState, answer: aState.answer === "no" ? null : "no" } }))}
                                  style={{ padding: "5px 14px", borderRadius: 6, border: `1px solid ${aState.answer === "no" ? "#dc2626" : T.border}`, background: aState.answer === "no" ? "#fee2e2" : T.surface, color: aState.answer === "no" ? "#dc2626" : T.textMuted, fontWeight: 700, fontSize: 13, cursor: "pointer" }}>
                                  ✗ No
                                </button>
                                <Input value={aState.value} onChange={e => setMgrAssess(p2 => ({ ...p2, [aKey]: { ...aState, value: e.target.value } }))} placeholder={`Actual value${kr.unit ? ` (${kr.unit})` : ""}`} style={{ padding: "5px 8px", fontSize: 13, fontFamily: F.mono, width: 140 }} />
                                <Btn small primary disabled={!aState.answer} onClick={() => {
                                  if (!aState.answer) return;
                                  dispatch({ type: "MANAGER_ASSESS_KR", memberId: m.id, memberName: m.name, deptId: m.deptId, kr, period: p, periodKey: pk, answer: aState.answer, actualValue: aState.value !== "" ? Number(aState.value) : null, approvedBy: user.id, newId: `ma_${m.id}_${kr.id}_${Date.now().toString(36)}` });
                                }}>Save</Btn>
                                {existingSub && <span style={{ fontSize: 12, color: T.textMuted }}>Last: <b style={{ color: existingSub.answer === "yes" ? T.ok : T.bad }}>{existingSub.answer}</b>{existingSub.actualValue != null ? ` · ${fmt(existingSub.actualValue)}${kr.unit ? ` ${kr.unit}` : ""}` : ""}</span>}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                    {ppKrsTeam.length > 0 && (
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "8px 18px", background: T.okDim, borderBottom: `1px solid ${T.okBorder}`, borderTop: `1px solid ${T.border}` }}>
                          <span style={{ fontSize: 11, fontWeight: 700, color: T.ok, textTransform: "uppercase", letterSpacing: "0.07em", flex: 1 }}>Project Profit KRs</span>
                          <span style={{ fontSize: 11, color: T.ok }}>auto-tracked from completed projects — read only</span>
                        </div>
                        {ppKrsTeam.map((kr, ki) => {
                          const ppTGetCY = p => { if (p.completedYear) return p.completedYear; const pts = (p.updatedDate || "").split("/"); return pts.length >= 3 ? parseInt(pts[2].split(",")[0].trim()) : null; };
                          const ppTAct = (state.projects || []).filter(p => p.mgrId === m.id && p.status === "completed" && ppTGetCY(p) === kr.krYear).reduce((s, p) => s + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0);
                          const ppTPct = kr.target > 0 ? Math.min(Math.round(ppTAct / kr.target * 100), 100) : 0;
                          const ppTSt = getStatus(ppTPct);
                          const ppTMissing = (state.projects || []).filter(p => p.mgrId === m.id && p.status === "completed" && ppTGetCY(p) === kr.krYear && (p.income == null || p.margin == null)).length;
                          return (
                            <div key={kr.id} style={{ padding: "12px 18px", borderBottom: ki < ppKrsTeam.length - 1 ? `1px solid ${T.border}` : "none", background: ki % 2 ? T.raised : "transparent" }}>
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                                <div style={{ flex: 1, minWidth: 0 }}>
                                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                                    <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim }}>{kr.id}</span>
                                    <span style={{ fontWeight: 600, fontSize: 14, flex: 1 }}>{kr.label}</span>
                                    <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px" }}>Proj Profit</span>
                                    <span style={{ fontSize: 11, color: T.textMuted }}>Year {kr.krYear || "?"}</span>
                                  </div>
                                  <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 5 }}>Target: ${(kr.target || 0).toLocaleString()} profit</div>
                                  <Bar value={ppTPct} status={ppTSt} h={5} />
                                  {ppTMissing > 0 && <div style={{ fontSize: 11, color: T.warn, marginTop: 4 }}>⚠ {ppTMissing} project{ppTMissing !== 1 ? "s" : ""} missing income/margin — not counted</div>}
                                </div>
                                <div style={{ textAlign: "right", flexShrink: 0 }}>
                                  <div style={{ fontSize: 20, fontWeight: 800, color: STATUS_THEME[ppTSt].color, fontFamily: F.mono }}>${ppTAct.toLocaleString()}</div>
                                  <div style={{ fontSize: 11, color: T.textMuted }}>of ${(kr.target || 0).toLocaleString()}</div>
                                  <div style={{ fontSize: 13, fontWeight: 700, color: STATUS_THEME[ppTSt].color }}>{ppTPct}%</div>
                                </div>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    )}
                  </Card>
                );
              })}
            </Pane>
          </>);
        })()}

        {page === "financial" && user.financeAccess && (<>
          <Header title="Financial Performance" sub="Income, Net Profit and Expenses tracking — FY2027" />
          <Pane>
            <FinErrorBoundary>
              <FinancialPerformancePage state={state} dispatch={dispatch} />
            </FinErrorBoundary>
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
                    <Tag type={getStatus(r.data.companyRate)} label={`Company: ${Number(r.data.companyRate).toFixed(1)}%`} />
                  </div>
                  <div style={{ padding: "14px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                    <div>
                      <SectionLabel>Department Rankings</SectionLabel>
                      {r.data.deptRanks.map((d, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 14 }}>
                          <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                          <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                          <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{Number(d.rate).toFixed(1)}%</span>
                          <Tag type={d.status} small />
                        </div>
                      ))}
                      {r.submissionRate != null && (
                        <div style={{ marginTop: 10 }}>
                          <span style={{ fontSize: 12, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "2px 8px", color: T.brand }}>Target met rate: {r.submissionRate}%</span>
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
            <div style={{ background: T.surface, borderRadius: 16, padding: "28px 32px", width: "100%", maxWidth: 420, boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}>
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
    </MobileContext.Provider>
  );
}

/* ─────────────────────────────────────────────────────────────
   MEMBER PORTAL
   ───────────────────────────────────────────────────────────── */
function MemberPortal({ user, onLogout, state, dispatch, onReload }) {
  const [page, setPageRaw] = useState(() => {
    const p = window.location.pathname.split('/');
    return p[1] === 'member' ? (p[2] || 'mykpis') : 'mykpis';
  });
  const setPage = useCallback(p => { window.history.pushState(null, '', `/member/${p}`); setPageRaw(p); }, []);
  const isMobile = useIsMobile();
  const [drawerOpen, setDrawerOpen] = useState(false);
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
  useEffect(() => { if (page === "checkin") onReload(); }, [page]); // eslint-disable-line
  const [myKpiPeriod, setMyKpiPeriod] = useState("all");
  const [okrPeriod, setOkrPeriod] = useState("all");
  const [expandedMonthlyKr, setExpandedMonthlyKr] = useState(null);
  const [noReason, setNoReason] = useState(null);
  const [yesConfirm, setYesConfirm] = useState(null);
  const [trackerInput, setTrackerInput] = useState({});
  const [expandedKrHistory, setExpandedKrHistory] = useState(null);
  const [histPeriod, setHistPeriod] = useState("all");
  const [syncing, setSyncing] = useState(false);
  const handleSync = useCallback(async () => { setSyncing(true); await onReload(); setSyncing(false); }, [onReload]);
  const [newProj, setNewProj] = useState({ name: "", startDate: "", due: "" });
  const [showNewProj, setShowNewProj] = useState(false);
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ status: "active", startDate: "", due: "", income: "", margin: "", contributeRate: "" });
  const [progressEdits, setProgressEdits] = useState({});
  const [logDrafts, setLogDrafts] = useState({});
  const [logPopup, setLogPopup] = useState(null);
  useEffect(() => { if (!logPopup) return; const h = e => { if (e.key === "Escape") setLogPopup(null); }; window.addEventListener("keydown", h); return () => window.removeEventListener("keydown", h); }, [logPopup]);
  const [designatedRejectOkr, setDesignatedRejectOkr] = useState(null);

  const { memberData, monthlyReports, depts, projects = [], users } = state;
  const kd = memberData[user.id] || { krs: [] };
  const myDept = depts.find(d => d.id === user.deptId);
  const myTeam = myDept?.teams.find(t => t.id === user.teamId);
  const mySecondTeam = user.secondTeamId ? myDept?.teams.find(t => t.id === user.secondTeamId) : null;
  const myOkrSubs = (state.okrSubmissions || []).filter(s => s.memberId === user.id);
  const myPendingCheckins = myOkrSubs.filter(s => s.answer === null);
  const hasRateKrs = memberHasRateKrs(kd.krs);
  const rate = hasRateKrs ? calcMemberRate(user.id, kd.krs, state.okrSubmissions || []) : null;
  const st = getStatus(rate);
  const pendingCount = myOkrSubs.filter(s => s.answer !== null && s.approval === "pending").length;
  const myOwnProjects = projects.filter(p => p.mgrId === user.id);
  const myProjects = projects.filter(p => user.deptId ? users.find(u => u.id === p.mgrId)?.deptId === user.deptId : p.mgrId === user.id);
  const designatedApproveeIds = users.filter(u => u.designatedApproverId === user.id).map(u => u.id);
  const designatedApproveeSubs = (state.okrSubmissions || []).filter(s => designatedApproveeIds.includes(s.memberId) && s.answer !== null && !s.managerFilled);
  const designatedPendingCount = designatedApproveeSubs.filter(s => s.approval === "pending").length;

  const navItems = [
    { id: "mykpis",       icon: "⬡", label: "My OKRs"          },
    { id: "checkin",      icon: "⬡", label: "OKR Check-In"     },
    { id: "okr-overview", icon: "⬡", label: "OKR Overview"     },
    { id: "history",      icon: "⬡", label: "My History"        },
    { id: "reports",      icon: "⬡", label: "OKR Reports"      },
    ...(designatedApproveeIds.length > 0 ? [{ id: "approvals", icon: "⬡", label: "Approvals" }] : []),
    ...(user.projectAccess ? [{ id: "projects", icon: "⬡", label: "Projects" }] : []),
  ];

  return (
    <MobileContext.Provider value={{ isMobile, drawerOpen, setDrawerOpen }}>
    <div style={{ display: "flex", minHeight: "100dvh", fontFamily: F.body, background: T.bg, color: T.text }}>
      {logPopup && <div onClick={() => setLogPopup(null)} style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.45)", zIndex: 9999, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}><div onClick={e => e.stopPropagation()} style={{ background: T.surface, borderRadius: 12, border: `1px solid ${T.border}`, width: "100%", maxWidth: 640, maxHeight: "75vh", display: "flex", flexDirection: "column", boxShadow: "0 8px 40px rgba(0,0,0,0.18)" }}><div style={{ padding: "14px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexShrink: 0 }}><div><div style={{ fontSize: 15, fontWeight: 700 }}>{logPopup.projName}</div>{logPopup.date && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>{logPopup.date}</div>}</div><button onClick={() => setLogPopup(null)} style={{ background: "none", border: "none", cursor: "pointer", color: T.textMuted, fontSize: 18, lineHeight: 1, padding: "0 2px", marginLeft: 12 }}>✕</button></div><div style={{ padding: "16px 20px", overflowY: "auto", fontSize: 14, lineHeight: 1.65, color: T.text, whiteSpace: "pre-wrap" }}>{logPopup.text}</div></div></div>}
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} pendingCounts={{ checkin: myPendingCheckins.length, approvals: designatedPendingCount }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "mykpis" && (<>
          <Header title="My OKRs" sub={`${user.title} · ${currentFYQuarter()}`} right={<Tag type={st} />} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="My Completion" value={(hasRateKrs && rate !== null) ? `${rate.toFixed(1)}%` : "N/A"} status={(hasRateKrs && rate !== null) ? st : undefined} sub={(hasRateKrs && rate !== null) ? `Time: ${TP}%` : undefined} />
              <Metric label="KRs Tracked"    value={kd.krs.length} />
              <Metric label="Check-Ins" value={myPendingCheckins.length === 0 ? "All Done" : `${myPendingCheckins.length} Pending`} status={myPendingCheckins.length === 0 ? "green" : "yellow"} />
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
                {user.projectAccess && myOwnProjects.length > 0 ? (<>
                  <div style={{ fontSize: 32, marginBottom: 10 }}>◫</div>
                  <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6, color: T.text }}>No OKRs — contribution tracked via Projects</div>
                  <div style={{ fontSize: 13, marginBottom: 16 }}>You have {myOwnProjects.length} project{myOwnProjects.length > 1 ? "s" : ""}. Your contribution is tracked through the Projects section.</div>
                  <div style={{ display: "flex", gap: 12, justifyContent: "center", flexWrap: "wrap", marginBottom: 16 }}>
                    <Metric label="Total Projects" value={myOwnProjects.length} />
                    <Metric label="Active" value={myOwnProjects.filter(p => p.status === "active").length} />
                    <Metric label="Avg Progress" value={`${Math.round(myOwnProjects.reduce((a, p) => a + p.progress, 0) / myOwnProjects.length)}%`} />
                  </div>
                  <Btn primary onClick={() => setPage("projects")}>Go to Projects →</Btn>
                </>) : user.projectAccess ? (<>
                  <div style={{ fontSize: 32, marginBottom: 10 }}>◫</div>
                  <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6, color: T.text }}>No OKRs assigned</div>
                  <div style={{ fontSize: 13, marginBottom: 16 }}>Track your work through Projects.</div>
                  <Btn primary onClick={() => setPage("projects")}>Go to Projects →</Btn>
                </>) : (<>
                  <div style={{ fontSize: 32, marginBottom: 10 }}>◎</div>
                  <div style={{ fontSize: 15, fontWeight: 700, marginBottom: 6, color: T.text }}>No KPIs assigned yet</div>
                  {!user.teamId
                    ? <div style={{ fontSize: 13 }}>You haven't been assigned to a team. Ask your admin to assign you to a team so your KPIs can be set up.</div>
                    : <div style={{ fontSize: 13 }}>Your manager hasn't synced KRs to your profile yet. Once they do, your KPIs will appear here.</div>}
                </>)}
              </div>
            )}
            {(myKpiPeriod === "all" ? kd.krs : kd.krs.filter(kr => (kr.period || "monthly") === myKpiPeriod)).map(kr => {
              // For standard KRs where kr.actual is null/0 (e.g. KR was synced after its check-ins were approved),
              // fall back to the most recent approved submission's actualValue so the display is meaningful.
              const isStdKr = !kr.monthlyTargets && kr.type !== "tracker" && kr.type !== "progress" && kr.type !== "manager-fill" && kr.type !== "project_profit";
              const effectiveKr = (isStdKr && (kr.actual == null || kr.actual === 0))
                ? (() => { const lastApproved = myOkrSubs.filter(sub => sub.krId === kr.id && sub.approval === "approved" && sub.answer === "yes" && sub.actualValue).sort((a, b) => (b.answeredAt || b.sentAt || "").localeCompare(a.answeredAt || a.sentAt || ""))[0]; return lastApproved ? { ...kr, actual: lastApproved.actualValue } : kr; })()
                : kr;
              const r = krCompletion(effectiveKr); const s = getStatus(r);
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
                      {kr.type === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 10, padding: "1px 6px", whiteSpace: "nowrap" }}>Progress · affects rate proportionally</span>}
                      {kr.type === "manager-fill" && <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 10, padding: "1px 6px", whiteSpace: "nowrap" }}>Mgr Fill · assessed by manager</span>}
                      {kr.type === "project_profit" && <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 10, padding: "1px 6px", whiteSpace: "nowrap" }}>Project Profit · auto-tracked</span>}
                      {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 10, padding: "1px 6px", whiteSpace: "nowrap" }}>Monthly Breakdown</span>}
                      <span style={{ fontSize: 10, color: T.textDim, background: T.raised, padding: "1px 6px", borderRadius: 10, border: `1px solid ${T.border}`, whiteSpace: "nowrap" }}>{kr.period || "monthly"}</span>
                      {kr.type !== "tracker" && kr.type !== "manager-fill" && kr.type !== "project_profit" && <Tag type={s} />}
                      {kr.type === "manager-fill" && (() => { const mfSub = myOkrSubs.filter(s2 => s2.krId === kr.id && s2.managerFilled).sort((a,b)=>(b.answeredAt||"").localeCompare(a.answeredAt||""))[0]; return mfSub ? <Tag type={mfSub.answer === "yes" ? "green" : "red"} /> : <Tag type="none" />; })()}
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
                  ) : kr.type === "manager-fill" ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
                      {(() => {
                        const mfSub = myOkrSubs.filter(s2 => s2.krId === kr.id && s2.managerFilled).sort((a,b)=>(b.answeredAt||"").localeCompare(a.answeredAt||""))[0];
                        if (!mfSub) return (
                          <div style={{ background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "12px 18px" }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: "#d97706" }}>Awaiting manager assessment</div>
                            <div style={{ fontSize: 11, color: "#d97706", opacity: 0.8, marginTop: 3 }}>Your manager will fill in this KR</div>
                          </div>
                        );
                        const isYes = mfSub.answer === "yes";
                        return (<>
                          <div>
                            <div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: isYes ? T.ok : T.bad }}>{isYes ? "Yes" : "No"}</div>
                            <div style={{ fontSize: 12, color: T.textMuted }}>Manager assessment</div>
                          </div>
                          {mfSub.actualValue != null && (
                            <div>
                              <div style={{ fontSize: 28, fontWeight: 800, fontFamily: F.mono, color: isYes ? T.ok : T.bad }}>{fmt(mfSub.actualValue)}{kr.unit ? <span style={{ fontSize: 14, fontWeight: 400, marginLeft: 5 }}>{kr.unit}</span> : null}</div>
                              <div style={{ fontSize: 12, color: T.textMuted }}>{kr.operator || ">="} {fmt(kr.target)} target</div>
                            </div>
                          )}
                          <div style={{ background: isYes ? "#dcfce7" : "#fee2e2", border: `1px solid ${isYes ? "#86efac" : "#fca5a5"}`, borderRadius: 8, padding: "6px 14px" }}>
                            <div style={{ fontSize: 13, fontWeight: 700, color: isYes ? "#16a34a" : "#dc2626" }}>{isYes ? "Met" : "Not Met"}</div>
                            <div style={{ fontSize: 11, color: isYes ? "#16a34a" : "#dc2626", opacity: 0.8 }}>assessed by manager</div>
                          </div>
                        </>);
                      })()}
                    </div>
                  ) : kr.type === "tracker" ? (
                    <div style={{ display: "flex", alignItems: "center", gap: 24 }}>
                      <div>
                        {(() => {
                          const latestApproved = myOkrSubs.filter(s => s.krId === kr.id && s.approval === "approved" && s.actualValue != null).sort((a, b) => (b.answeredAt || b.sentAt || "").localeCompare(a.answeredAt || a.sentAt || ""))[0];
                          const displayVal = latestApproved != null ? latestApproved.actualValue : kr.actual;
                          return (<>
                            <div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: "#7c3aed" }}>
                              {fmt(displayVal)}{kr.unit ? <span style={{ fontSize: 16, fontWeight: 600, color: T.textMuted, marginLeft: 6 }}>{kr.unit}</span> : null}
                            </div>
                            <div style={{ fontSize: 12, color: T.textMuted }}>Last recorded value</div>
                          </>);
                        })()}
                      </div>
                      <div style={{ background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "6px 14px" }}>
                        <div style={{ fontSize: 13, fontWeight: 700, color: "#7c3aed" }}>N/A</div>
                        <div style={{ fontSize: 11, color: "#7c3aed", opacity: 0.8 }}>does not affect rate</div>
                      </div>
                    </div>
                  ) : kr.type === "project_profit" ? (
                    (() => {
                      const kpiGetCY = p => { if (p.completedYear) return p.completedYear; const pts = (p.updatedDate || "").split("/"); return pts.length >= 3 ? parseInt(pts[2].split(",")[0].trim()) : null; };
                      const kpiYearProjects = projects.filter(p => p.mgrId === user.id && p.status === "completed" && kpiGetCY(p) === kr.krYear);
                      const kpiAct = kpiYearProjects.reduce((acc, p) => acc + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0);
                      const kpiPct = kr.target > 0 ? Math.min(Math.round(kpiAct / kr.target * 100), 100) : 0;
                      const kpiSt = getStatus(kpiPct);
                      const kpiMissing = kpiYearProjects.filter(p => p.income == null || p.margin == null).length;
                      const kpiHasPartial = kpiYearProjects.some(p => p.contributeRate != null && p.contributeRate < 100);
                      return (
                        <div style={{ display: "flex", alignItems: "flex-end", gap: 20 }}>
                          <div>
                            <div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: STATUS_THEME[kpiSt].color }}>${kpiAct.toLocaleString()}</div>
                            <div style={{ fontSize: 12, color: T.textMuted }}>of ${(kr.target || 0).toLocaleString()} profit · Year {kr.krYear || "?"}</div>
                            {kpiMissing > 0 && <div style={{ fontSize: 11, color: T.warn, marginTop: 3 }}>⚠ {kpiMissing} project{kpiMissing !== 1 ? "s" : ""} missing income/margin</div>}
                            {kpiHasPartial && <div style={{ fontSize: 11, color: T.brand, marginTop: 2 }}>⚡ Partial contribution rates applied</div>}
                          </div>
                          <div style={{ flex: 1 }}><Bar value={kpiPct} status={kpiSt} h={10} /></div>
                          <div>
                            <div style={{ fontSize: 26, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[kpiSt].color }}>{kpiPct}%</div>
                            <Tag type={kpiSt} />
                          </div>
                        </div>
                      );
                    })()
                  ) : (
                    <div style={{ display: "flex", alignItems: "flex-end", gap: 20 }}>
                      <div><div style={{ fontSize: 34, fontWeight: 900, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{fmt(effectiveKr.actual)}</div><div style={{ fontSize: 12, color: T.textMuted }}>{kr.type === "progress" ? "" : `${kr.operator || ">="} `}{fmt(kr.target)} target{kr.unit ? ` (${kr.unit})` : ""}</div></div>
                      <div style={{ flex: 1 }}><Bar value={r} status={s} h={10} /></div>
                      <div>
                        <div style={{ fontSize: 26, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</div>
                        {effectiveKr.actual > kr.target && <div style={{ fontSize: 10, color: T.ok, fontWeight: 600 }}>↑ exceeded</div>}
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
                          {isOpen ? `▲ Hide ${kr.type === "manager-fill" ? "Assessment" : "Check-In"} History` : `▼ ${kr.type === "manager-fill" ? "Assessment" : "Check-In"} History (${hist.length})`}
                        </button>
                        {isOpen && (
                          <div style={{ marginTop: 10, display: "flex", flexDirection: "column", gap: 5 }}>
                            {hist.map(s => {
                              const leftCol = s.approval === "approved" ? T.ok : s.approval === "rejected" ? T.bad : s.answer !== null ? T.warn : T.border;
                              const ansCol = s.answer === "yes" ? T.ok : s.answer === "no" ? T.bad : s.answer === "submitted" ? (s.krType === "progress" ? T.brand : "#7c3aed") : T.textDim;
                              const ansLabel = s.answer === "yes" ? "✓ Yes" : s.answer === "no" ? "✗ No" : s.answer === "submitted" ? "Recorded" : "Not answered";
                              return (
                                <div key={s.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "7px 12px", borderRadius: 7, background: T.raised, borderLeft: `3px solid ${leftCol}`, fontSize: 13, flexWrap: "wrap" }}>
                                  <span style={{ fontWeight: 600, flex: 1, minWidth: 120 }}>{s.dateRange || s.periodKey}</span>
                                  <span style={{ fontSize: 11, color: T.textMuted }}>{s.period}</span>
                                  <span style={{ fontWeight: 700, color: ansCol, minWidth: 90, textAlign: "right" }}>{ansLabel}</span>
                                  {(s.answer === "no" || s.answer === "submitted") && s.actualValue != null && <span style={{ fontFamily: F.mono, fontSize: 12, color: s.answer === "submitted" ? "#7c3aed" : T.textMuted }}>{s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                  <Tag type={s.approval} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                                  {s.managerFilled && <span style={{ fontSize: 9, fontWeight: 700, background: "#fef3c7", border: "1px solid #fde68a", color: "#d97706", borderRadius: 8, padding: "1px 6px", letterSpacing: "0.05em" }}>MGR</span>}
                                  {!s.managerFilled && s.approvedBy === "auto" && <span style={{ fontSize: 9, fontWeight: 700, background: "rgba(16,185,129,0.1)", border: "1px solid rgba(16,185,129,0.25)", color: "#10B981", borderRadius: 8, padding: "1px 6px", letterSpacing: "0.05em" }}>AUTO</span>}
                                  {!s.managerFilled && s.approvedBy && s.approvedBy !== "auto" && <span style={{ fontSize: 10, color: T.textMuted }}>by {users?.find(u => u.id === s.approvedBy)?.name || "Admin"}</span>}
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
            const _pkMem = okrPeriod === "weekly" ? prevPeriodKey(okrPeriod) : currentPeriodKey(okrPeriod);
            const hasSub = okrPeriod === "all" ? (isMonthly ? Object.values(kr.monthlyActuals || {}).some(v => v != null) : kr.actual != null) : (state.okrSubmissions || []).some(s => s.krId === kr.id && s.period === (kr.period || "monthly") && s.periodKey === _pkMem && s.answer !== null);
            const { ppActual, ppHasPartial } = kr.type === "project_profit" ? (() => { const getCompletedYear = p => { if (p.completedYear) return p.completedYear; const parts = (p.updatedDate || "").split("/"); return parts.length >= 3 ? parseInt(parts[2].split(",")[0].trim()) : null; }; const ppYrProj = projects.filter(p => p.mgrId === user.id && p.status === "completed" && getCompletedYear(p) === kr.krYear); return { ppActual: ppYrProj.reduce((sum, p) => sum + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0), ppHasPartial: ppYrProj.some(p => p.contributeRate != null && p.contributeRate < 100) }; })() : { ppActual: null, ppHasPartial: false };
            const ppPct = kr.type === "project_profit" && kr.target > 0 ? Math.min(Math.round(ppActual / kr.target * 100), 100) : null;
            const ppSt = ppPct != null ? getStatus(ppPct) : null;
            return (
              <Fragment key={kr.id}>
              <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 14 }}>
                <span style={{ fontFamily: F.mono, fontSize: 12, color: T.textDim }}>{kr.id}</span>
                <div>
                  <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                  {kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                  {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Tracker · does not affect rate</span>}
                  {kr.type === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Progress · affects rate proportionally</span>}
                  {kr.type === "project_profit" && <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Project Profit · auto-tracked</span>}
                  {kr.type === "project_profit" && ppHasPartial && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>⚡ Partial rates</span>}
                  {isMonthly && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>Monthly Breakdown</span>}
                  {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", display: "inline-block" }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                </div>
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 12, color: "#7c3aed" }}>N/A</span> : kr.type === "project_profit" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>${(kr.target || 0).toLocaleString()} ({kr.krYear})</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? `${kr.operator||">="} ${fmt(curTarget)}` : kr.type === "progress" ? fmt(kr.target) : `${kr.operator || ">="} ${fmt(kr.target)}${kr.unit ? ` ${kr.unit}` : ""}`}</span>}
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textDim }}>—</span> : kr.type === "project_profit" ? <span style={{ textAlign: "right", fontFamily: F.mono, color: T.ok, fontWeight: 700 }}>${(ppActual || 0).toLocaleString()}</span> : <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{isMonthly ? fmt(curActual) : fmt(kr.actual)}</span>}
                {kr.type === "tracker" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontSize: 14, fontWeight: 700, color: "#7c3aed" }}>{fmt(isMonthly ? curActual : kr.actual)}{kr.unit ? <span style={{ fontSize: 11, fontWeight: 400 }}> {kr.unit}</span> : ""}</span> : kr.type === "project_profit" ? <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[ppSt || "none"].color }}>{ppPct ?? 0}%</span> : hasSub ? <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[s].color }}>{pct.toFixed(0)}%</span> : <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: T.textDim }}>N/A</span>}
                {kr.type === "tracker" ? <span /> : kr.type === "project_profit" ? <Bar value={ppPct || 0} status={ppSt || "none"} h={5} /> : hasSub ? <Bar value={pct} status={s} h={5} /> : <span />}
                <div style={{ display: "flex", justifyContent: "flex-end", gap: 4 }}>
                  {isMonthly && <button onClick={() => setExpandedMonthlyKr(p => p === kr.id ? null : kr.id)} title="View all months" style={{ background: expandedMonthlyKr === kr.id ? T.brand : T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "2px 7px", cursor: "pointer", color: expandedMonthlyKr === kr.id ? "#fff" : T.brand, fontSize: 11, fontWeight: 700 }}>📅</button>}
                  {kr.type === "tracker" ? null : kr.type === "project_profit" ? <Tag type={ppSt || "none"} small /> : hasSub ? <Tag type={s} small /> : null}
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
                <div style={{ overflowX: "auto" }}><div style={{ minWidth: 610 }}>
                <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 11, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                  <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Performance Target</span><span style={{ textAlign: "right" }}>Actual</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                </div>
                {renderKrRows(filtered)}
                </div></div>
              </Card>
            );
          };
          const dKrs = filterP(myDept.krs);
          const _deptSubs = (state.okrSubmissions || []).filter(s => s.answer !== null);
          const _deptRates = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === user.deptId && !u.excludeFromRate).map(u => { const kd2 = memberData[u.id] || { krs: [] }; if (!memberHasRateKrs(kd2.krs)) return null; return calcMemberRate(u.id, kd2.krs, _deptSubs); }).filter(r => r !== null);
          const deptRate = _deptRates.length ? _deptRates.reduce((a, b) => a + b, 0) / _deptRates.length : 0;
          const deptStatus = getStatus(deptRate);
          const filtMKrs = krs => okrPeriod === "all" ? krs : krs.filter(kr => (kr.period || "monthly") === okrPeriod);
          const allTeamStats = myDept.teams.map(t => {
            const tKrs = filterP(t.krs);
            const teamMembers = users.filter(u => (u.role === "member" || u.role === "manager") && u.deptId === user.deptId && !u.excludeFromRate && (u.teamId === t.id || u.secondTeamId === t.id));
            const memberRates = teamMembers.map(u => {
              const kd = memberData[u.id];
              if (!kd || !memberHasRateKrs(filtMKrs(kd.krs))) return null;
              return calcMemberRate(u.id, filtMKrs(kd.krs), _deptSubs);
            }).filter(r => r !== null);
            const rate = memberRates.length ? memberRates.reduce((a, b) => a + b, 0) / memberRates.length : null;
            return { ...t, krs: tKrs, rate, status: getStatus(rate) };
          }).filter(t => t.krs.length > 0);
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
                      <span style={{ fontSize: 14, fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[t.status].color }}>{t.rate != null ? `${t.rate.toFixed(1)}%` : "N/A"}</span>
                      <div style={{ width: 100, flexShrink: 0 }}><Bar value={t.rate ?? 0} status={t.status} h={5} /></div>
                      <Tag type={t.status} small />
                    </div>
                    {t.krs.map(kr => {
                      const pct = krCompletion(kr); const st = getStatus(pct);
                      const trackerVal = (kr.type === "tracker" && kr.actual != null && kr.actual !== 0)
                        ? `${fmt(kr.actual)}${kr.unit ? ` ${kr.unit}` : ""}` : null;
                      const hasSub = kr.type !== "tracker" && (!!kr.monthlyTargets
                        ? Object.values(kr.monthlyActuals || {}).some(v => v != null && v !== 0)
                        : (kr.actual != null && kr.actual !== 0));
                      return (
                      <div key={kr.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0 5px 12px", borderTop: `1px solid ${T.border}`, fontSize: 13 }}>
                        <span style={{ fontFamily: F.mono, fontSize: 11, color: T.textDim, width: 50, flexShrink: 0 }}>{kr.id}</span>
                        <span style={{ flex: 1, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.label}</span>
                        {kr.type === "tracker" && <span style={{ fontSize: 10, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Tracker</span>}
                        {kr.type === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Progress</span>}
                        {kr.type === "manager-fill" && <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Mgr Fill</span>}
                        {kr.type === "project_profit" && <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Proj Profit</span>}
                        {kr.type !== "tracker" && kr.type !== "progress" && kr.type !== "manager-fill" && kr.type !== "project_profit" && kr.unit && <span style={{ fontSize: 11, color: T.textMuted }}>{kr.unit}</span>}
                        {okrPeriod === "all" && kr.period && <span style={{ fontSize: 10, color: T.textMuted, background: T.raised, border: `1px solid ${T.border}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>{kr.period.charAt(0).toUpperCase() + kr.period.slice(1)}</span>}
                        {kr.type === "tracker"
                          ? <span style={{ fontSize: 12, fontFamily: F.mono, color: trackerVal ? "#7c3aed" : T.textDim, fontWeight: 700, textAlign: "right", flexShrink: 0 }}>{trackerVal ?? "N/A"}</span>
                          : hasSub ? <span style={{ fontSize: 12, fontFamily: F.mono, color: STATUS_THEME[st].color, fontWeight: 700, width: 40, textAlign: "right" }}>{pct.toFixed(0)}%</span>
                          : <span style={{ fontSize: 12, fontFamily: F.mono, color: T.textDim, fontWeight: 700, width: 40, textAlign: "right" }}>N/A</span>}
                        {kr.type === "tracker" ? <span style={{ width: 100, flexShrink: 0 }} /> : hasSub ? <div style={{ width: 100, flexShrink: 0 }}><Bar value={pct} status={st} h={5} /></div> : <span style={{ width: 100, flexShrink: 0 }} />}
                        {kr.type !== "tracker" && hasSub && <Tag type={st} small />}
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
          const grouped = PERIOD_ORDER.map(p => ({ period: p, pending: myOkrSubs.filter(s => s.period === p && s.answer === null && !s.managerFilled), answered: myOkrSubs.filter(s => s.period === p && s.answer !== null && !s.managerFilled).sort((a,b) => (b.answeredAt||"").localeCompare(a.answeredAt||"")) })).filter(g => g.pending.length + g.answered.length > 0);
          const PERIOD_COLORS = { daily: T.warn, weekly: T.brand, monthly: "#A78BFA", quarterly: "#F97316", biannual: "#06B6D4", annual: T.ok };
          const currentMonthKey = currentFYMonthKey();
          const subRate = calcSubmissionRate(myOkrSubs, user.id, currentMonthKey);
          return (<>
            <Header title="OKR Check-In" sub="Answer your KPI check-ins sent by the system"
              right={<div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                {subRate !== null && <><span style={{ fontSize: 12, color: T.textMuted }}>This month:</span><span style={{ fontWeight: 700, fontSize: 15, color: STATUS_THEME[getStatus(subRate)].color, fontFamily: F.mono }}>{subRate.toFixed(0)}%</span></>}
                <Btn small onClick={handleSync} disabled={syncing}>{syncing ? "Syncing…" : "⟳ Sync"}</Btn>
              </div>} />
            <Pane>
              {myPendingCheckins.length > 0 && (
                <div style={{ padding: "10px 14px", background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 8, fontSize: 13, color: T.warn, fontWeight: 600, marginBottom: 16 }}>
                  {myPendingCheckins.length} pending check-in{myPendingCheckins.length !== 1 ? "s" : ""} — please respond below
                </div>
              )}
              {(() => {
                const ppKrs = (memberData[user.id]?.krs || []).filter(kr => kr.type === "project_profit");
                if (!ppKrs.length) return null;
                const getCompletedYear = p => {
                  if (p.completedYear) return p.completedYear;
                  const parts = (p.updatedDate || "").split("/");
                  return parts.length >= 3 ? parseInt(parts[2].split(",")[0].trim()) : null;
                };
                return (
                  <div style={{ marginBottom: 24 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${T.okBorder}` }}>
                      <div style={{ width: 4, height: 18, background: T.ok, borderRadius: 2 }} />
                      <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>Project Profit KRs</span>
                      <span style={{ fontSize: 11, color: T.textMuted }}>auto-tracked from completed projects</span>
                    </div>
                    {ppKrs.map(kr => {
                      const yearProjects = projects.filter(p => p.mgrId === user.id && p.status === "completed" && getCompletedYear(p) === kr.krYear);
                      const actual = yearProjects.reduce((s, p) => s + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0);
                      const pct = kr.target > 0 ? Math.min(Math.round(actual / kr.target * 100), 100) : 0;
                      const st = getStatus(pct);
                      const missingMargin = yearProjects.filter(p => p.income == null || p.margin == null).length;
                      const hasPartialRate = yearProjects.some(p => p.contributeRate != null && p.contributeRate < 100);
                      return (
                        <Card key={kr.id} style={{ padding: "14px 18px", marginBottom: 8, borderLeft: `3px solid ${T.ok}` }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 10 }}>
                            <div style={{ flex: 1 }}>
                              <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                                <span style={{ fontSize: 15, fontWeight: 700 }}>{kr.label}</span>
                                <span style={{ fontSize: 10, fontWeight: 700, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Project Profit · auto</span>
                              </div>
                              <div style={{ fontSize: 12, color: T.textMuted }}>Target: ${(kr.target || 0).toLocaleString()} · {kr.krYear}</div>
                            </div>
                            <div style={{ textAlign: "right", flexShrink: 0 }}>
                              <div style={{ fontSize: 20, fontWeight: 700, fontFamily: F.mono, color: STATUS_THEME[st].color }}>{pct}%</div>
                              <div style={{ fontSize: 12, color: T.textMuted, fontFamily: F.mono }}>${actual.toLocaleString()} / ${(kr.target || 0).toLocaleString()}</div>
                            </div>
                          </div>
                          <Bar value={pct} status={st} h={6} />
                          <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
                            <span style={{ fontSize: 11, color: T.textMuted }}>{yearProjects.length} completed project{yearProjects.length !== 1 ? "s" : ""} in {kr.krYear}</span>
                            {missingMargin > 0 && <span style={{ fontSize: 11, color: T.warn, background: T.warnDim, border: `1px solid ${T.warnBorder}`, borderRadius: 5, padding: "1px 6px" }}>⚠ {missingMargin} project{missingMargin !== 1 ? "s" : ""} missing income/margin — not counted</span>}
                            {hasPartialRate && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 6px" }}>⚡ Partial rates applied</span>}
                          </div>
                        </Card>
                      );
                    })}
                  </div>
                );
              })()}
              {grouped.length === 0 && <EmptyState text="No check-ins yet. Admin will send them when due." />}
              {grouped.map(({ period, pending, answered }) => {
                const byPK = {};
                [...pending, ...answered].forEach(s => {
                  if (!byPK[s.periodKey]) byPK[s.periodKey] = { pk: s.periodKey, dr: s.dateRange || "", p: [], a: [] };
                  if (s.answer === null) byPK[s.periodKey].p.push(s); else byPK[s.periodKey].a.push(s);
                });
                const pkGroups = Object.values(byPK).sort((a, b) =>
                  a.p.length > 0 && b.p.length === 0 ? -1 : a.p.length === 0 && b.p.length > 0 ? 1 : b.pk.localeCompare(a.pk));
                return (
                <div key={period} style={{ marginBottom: 28 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${PERIOD_COLORS[period]}` }}>
                    <div style={{ width: 4, height: 18, background: PERIOD_COLORS[period], borderRadius: 2 }} />
                    <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>{period.charAt(0).toUpperCase() + period.slice(1)} Check-Ins</span>
                    {pending.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 7px", fontSize: 11, fontWeight: 700 }}>{pending.length} pending</span>}
                  </div>
                  {pkGroups.map(({ pk, dr, p: pkPending, a: pkAnswered }) => (
                    <div key={pk} style={{ marginBottom: 14 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "6px 12px", marginBottom: 8, borderRadius: 6, background: pkPending.length > 0 ? T.warnDim : T.raised, border: `1px solid ${pkPending.length > 0 ? T.warnBorder : T.border}` }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: pkPending.length > 0 ? T.warn : T.textMuted }}>{dr || pk}</span>
                        {pkPending.length > 0 && <span style={{ background: T.warn, color: "#fff", borderRadius: 8, padding: "1px 6px", fontSize: 10, fontWeight: 700 }}>{pkPending.length} pending</span>}
                        {pkPending.length === 0 && <span style={{ fontSize: 11, color: T.ok }}>✓ All answered</span>}
                      </div>
                  {pkPending.map(s => (
                    <Card key={s.id} style={{ padding: "14px 18px", marginBottom: 8, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : s.krType === "progress" ? T.brand : noReason?.id === s.id ? T.bad : yesConfirm?.id === s.id ? T.ok : T.warn}` }}>
                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 12 }}>
                        <div style={{ flex: 1 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 3 }}>
                            <span style={{ fontSize: 15, fontWeight: 700 }}>{s.krLabel}</span>
                            {s.krUnit && <span style={{ fontSize: 10, fontWeight: 700, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 6px" }}>{s.krUnit}</span>}
                            {s.krIsMonthly && s.krType !== "tracker" && s.krType !== "progress" && <span style={{ fontSize: 10, fontWeight: 700, background: "#e0f2fe", color: "#0369a1", border: "1px solid #7dd3fc", borderRadius: 5, padding: "1px 6px" }}>Monthly</span>}
                            {s.krType === "tracker" && <span style={{ fontSize: 10, fontWeight: 700, background: "#ede9fe", color: "#6d28d9", border: "1px solid #c4b5fd", borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Tracker · does not affect rate</span>}
                            {s.krType === "progress" && <span style={{ fontSize: 10, fontWeight: 700, background: T.brandDim, color: T.brand, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 6px", textTransform: "uppercase", letterSpacing: ".05em" }}>Progress · affects rate proportionally</span>}
                          </div>
                          <div style={{ fontSize: 12, color: T.textMuted }}>
                            {(s.krType !== "tracker" && s.krType !== "progress") && <span>Performance Target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                            {s.krType === "progress" && <span>Target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                            {s.krType === "tracker" && s.krUnit && <span>Unit: {s.krUnit}</span>}
                          </div>
                        </div>
                        {(s.krType === "tracker" || s.krType === "progress") ? (
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            <Input value={trackerInput[s.id] || ""} onChange={e => setTrackerInput(p => ({ ...p, [s.id]: e.target.value }))} placeholder="Enter value" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                            {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                            <Btn primary small onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "submitted", actualValue: Number(trackerInput[s.id]) || 0 }); setTrackerInput(p => ({ ...p, [s.id]: "" })); }} disabled={!trackerInput[s.id] || (s.krType === "tracker" && s.krDisallowZero && Number(trackerInput[s.id]) === 0)}>{s.krType === "progress" ? "Record Progress" : "Record"}</Btn>
                            {s.krType === "tracker" && s.krDisallowZero && trackerInput[s.id] === "0" && <span style={{ fontSize: 11, color: T.bad, whiteSpace: "nowrap" }}>Cannot be 0</span>}
                          </div>
                        ) : (noReason?.id !== s.id && yesConfirm?.id !== s.id) ? (
                          <div style={{ display: "flex", gap: 8, flexShrink: 0 }}>
                            <button onClick={() => { setNoReason({ id: s.id, reason: "", actual: "" }); setYesConfirm(null); }}
                              style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.bad, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✗ No
                            </button>
                            <button onClick={() => { setYesConfirm({ id: s.id, actual: String(s.krTarget ?? "") }); setNoReason(null); }}
                              style={{ background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 7, padding: "8px 18px", cursor: "pointer", color: T.ok, fontSize: 14, fontWeight: 700, fontFamily: F.body }}>
                              ✓ Yes
                            </button>
                          </div>
                        ) : null}
                      </div>
                      {s.krType !== "tracker" && s.krType !== "progress" && yesConfirm?.id === s.id && (
                        <div style={{ marginTop: 12, padding: "12px 14px", background: T.okDim, borderRadius: 8, border: `1px solid ${T.okBorder}` }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: T.ok, marginBottom: 8 }}>Enter your actual value</div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                            <Input value={yesConfirm.actual} onChange={e => setYesConfirm(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} autoFocus />
                            {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                            <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                          </div>
                          {yesConfirm.actual !== "" && meetsTarget(yesConfirm.actual, s.krOperator, s.krTarget) === false && (
                            <div style={{ fontSize: 12, color: T.bad, fontWeight: 600, marginBottom: 8 }}>⚠ Actual ({yesConfirm.actual}{s.krUnit ? " " + s.krUnit : ""}) doesn't meet target ({s.krOperator || ">="} {s.krTarget}{s.krUnit ? " " + s.krUnit : ""}) — answer should be No</div>
                          )}
                          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                            <Btn small onClick={() => setYesConfirm(null)}>Cancel</Btn>
                            <Btn primary small disabled={yesConfirm.actual !== "" && meetsTarget(yesConfirm.actual, s.krOperator, s.krTarget) === false} onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "yes", actualValue: Number(yesConfirm.actual) || 0 }); setYesConfirm(null); }}>✓ Submit Yes</Btn>
                          </div>
                        </div>
                      )}
                      {s.krType !== "tracker" && s.krType !== "progress" && noReason?.id === s.id && (
                        <div style={{ marginTop: 12, padding: "12px 14px", background: T.badDim, borderRadius: 8, border: `1px solid ${T.badBorder}` }}>
                          <div style={{ fontSize: 13, fontWeight: 700, color: T.bad, marginBottom: 8 }}>Why was this OKR not met?</div>
                          <TextArea value={noReason.reason} onChange={e => setNoReason(p => ({ ...p, reason: e.target.value }))} placeholder="Briefly explain why this target was not reached..." rows={2} />
                          <div style={{ marginTop: 10 }}>
                            <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 5 }}>Your actual value</div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              <Input value={noReason.actual} onChange={e => setNoReason(p => ({ ...p, actual: e.target.value }))} placeholder="0" style={{ width: 110, textAlign: "right", fontFamily: F.mono }} />
                              {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted, fontWeight: 600 }}>{s.krUnit}</span>}
                              <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>
                            </div>
                            {noReason.actual !== "" && meetsTarget(noReason.actual, s.krOperator, s.krTarget) === true && (
                              <div style={{ fontSize: 12, color: T.bad, fontWeight: 600, marginTop: 6 }}>⚠ Actual ({noReason.actual}{s.krUnit ? " " + s.krUnit : ""}) meets target ({s.krOperator || ">="} {s.krTarget}{s.krUnit ? " " + s.krUnit : ""}) — answer should be Yes</div>
                            )}
                          </div>
                          <div style={{ display: "flex", gap: 8, justifyContent: "flex-end", marginTop: 10 }}>
                            <Btn small onClick={() => setNoReason(null)}>Cancel</Btn>
                            <Btn danger small disabled={noReason.actual !== "" && meetsTarget(noReason.actual, s.krOperator, s.krTarget) === true} onClick={() => { dispatch({ type: "ANSWER_OKR_SUBMISSION", id: s.id, answer: "no", reason: noReason.reason.trim() || null, actualValue: Number(noReason.actual) || 0 }); setNoReason(null); }}>Submit No</Btn>
                          </div>
                        </div>
                      )}
                    </Card>
                  ))}
                  {pkAnswered.length > 0 && (
                    <div style={{ marginTop: pkPending.length > 0 ? 6 : 0 }}>
                      {pkAnswered.slice(0, 10).map(s => (
                        <Card key={s.id} style={{ padding: "10px 14px", marginBottom: 4, borderLeft: `3px solid ${s.krType === "tracker" ? "#7c3aed" : s.krType === "progress" ? T.brand : s.answer === "yes" ? T.ok : T.bad}`, opacity: s.approval === "approved" ? 0.7 : 1 }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div>
                              <span style={{ fontSize: 13, fontWeight: 600 }}>{s.krLabel}</span>
                              <span style={{ fontSize: 11, color: T.textMuted, marginLeft: 8 }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                            </div>
                            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                              {s.krType === "tracker"
                                ? <span style={{ fontSize: 12, fontWeight: 700, color: "#6d28d9" }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                                : s.krType === "progress"
                                ? <span style={{ fontSize: 12, fontWeight: 700, color: T.brand }}>Recorded: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}{s.krTarget ? ` (${Math.min(Math.round((Number(s.actualValue || 0) / Number(s.krTarget)) * 100), 100)}%)` : ""}</span>
                                : <div style={{ textAlign: "right" }}>
                                    <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>
                                    {s.actualValue != null && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</div>}
                                  </div>}
                              <Tag type={s.approval === "approved" ? "approved" : s.approval === "rejected" ? "rejected" : "pending"} label={s.approval === "approved" ? "Approved" : s.approval === "rejected" ? "Rejected" : "Pending"} small />
                              {s.approvedBy === "auto" && <span style={{ fontSize: 9, fontWeight: 700, background: "rgba(16,185,129,0.1)", border: "1px solid rgba(16,185,129,0.25)", color: "#10B981", borderRadius: 8, padding: "1px 6px", letterSpacing: "0.05em" }}>AUTO</span>}
                              {s.approvedBy && s.approvedBy !== "auto" && <span style={{ fontSize: 10, color: T.textMuted }}>by {users?.find(u => u.id === s.approvedBy)?.name || "Admin"}</span>}
                            </div>
                          </div>
                          {s.krType !== "tracker" && s.answer === "no" && s.reason && <div style={{ fontSize: 12, color: T.textSoft, marginTop: 5, paddingTop: 5, borderTop: `1px solid ${T.border}` }}>Note: {s.reason}</div>}
                        </Card>
                      ))}
                    </div>
                  )}
                    </div>
                  ))}
                </div>
                );
              })}
            </Pane>
          </>);
        })()}

        {page === "approvals" && (<>
          <Header title="Pending Approvals" sub={`${designatedPendingCount} pending review`} />
          <Pane>
            {(() => {
              const sorted = designatedApproveeSubs.slice().sort((a, b) => {
                const o = { pending: 0, approved: 1, rejected: 2 };
                return o[a.approval] - o[b.approval] || (b.answeredAt || "").localeCompare(a.answeredAt || "");
              });
              const order = [];
              const groups = {};
              sorted.forEach(s => {
                if (!groups[s.memberId]) { groups[s.memberId] = []; order.push(s.memberId); }
                groups[s.memberId].push(s);
              });
              if (order.length === 0) return <EmptyState text="No submissions to review yet." />;
              return (<>
                {order.map(memberId => {
                  const subs = groups[memberId];
                  const mem = users.find(u => u.id === memberId);
                  const pCount = subs.filter(s => s.approval === "pending").length;
                  return (
                    <Card key={memberId} style={{ marginBottom: 10, overflow: "hidden" }}>
                      <div style={{ padding: "11px 16px", background: T.raised, borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                        <Avatar letters={mem?.av || "?"} size={28} />
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <span style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{mem?.name || subs[0]?.memberName || "Unknown"}</span>
                          <div style={{ fontSize: 11, color: T.textDim, marginTop: 1 }}>{subs.length} KR{subs.length !== 1 ? "s" : ""}</div>
                        </div>
                        {pCount > 0 && (
                          <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                            <span style={{ background: T.warnDim, color: T.warn, border: `1px solid ${T.warnBorder}`, borderRadius: 6, padding: "2px 9px", fontSize: 11, fontWeight: 700 }}>{pCount} pending</span>
                            <Btn primary small onClick={() => subs.filter(s => s.approval === "pending").forEach(s => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id }))}>Approve All</Btn>
                          </div>
                        )}
                      </div>
                      {subs.map((s, idx) => {
                        const accentColor = s.approval === "approved" ? T.ok : s.approval === "rejected" ? T.bad : T.warn;
                        const isLast = idx === subs.length - 1 && designatedRejectOkr?.id !== s.id;
                        return (
                          <div key={s.id} style={{ borderBottom: isLast ? "none" : `1px solid ${T.border}` }}>
                            <div style={{ padding: "10px 16px 10px 19px", borderLeft: `3px solid ${accentColor}`, display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 2 }}>
                                  <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>{s.krLabel}</span>
                                  {s.krType === "tracker" && <span style={{ fontSize: 10, fontWeight: 700, color: "#7c3aed", background: "#ede9fe", border: "1px solid #c4b5fd", borderRadius: 8, padding: "1px 5px" }}>Tracker</span>}
                                  {s.krType === "progress" && <span style={{ fontSize: 10, fontWeight: 700, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "1px 5px" }}>Progress</span>}
                                </div>
                                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                                  <span style={{ fontSize: 11, color: T.textMuted }}>{periodDisplayLabel(s.period, s.periodKey)}</span>
                                  {s.krType !== "tracker" && s.krTarget != null && <span style={{ fontSize: 11, color: T.textMuted }}>Target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                  {s.actualValue != null && <span style={{ fontSize: 11, fontFamily: F.mono, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.textMuted }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>}
                                  {s.answeredAt && <span style={{ fontSize: 11, color: T.textDim }}>Answered {s.answeredAt.slice(0, 10)}</span>}
                                </div>
                                {s.answer === "no" && s.reason && <div style={{ fontSize: 11, color: T.textSoft, marginTop: 3 }}>Note: {s.reason}</div>}
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
                                {s.krType === "tracker" || s.krType === "progress"
                                  ? <span style={{ fontSize: 12, fontWeight: 700, color: s.krType === "tracker" ? "#6d28d9" : T.brand }}>{s.krType === "tracker" ? "Recorded" : "Progress"}: {s.actualValue ?? "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                                  : <div style={{ textAlign: "right" }}>
                                    <span style={{ fontSize: 12, fontWeight: 700, color: s.answer === "yes" ? T.ok : T.bad }}>{s.answer === "yes" ? "✓ Yes" : "✗ No"}</span>
                                    {s.actualValue != null && <div style={{ fontSize: 11, color: T.textMuted, marginTop: 2 }}>Actual: {s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</div>}
                                  </div>}
                                {s.approval === "pending"
                                  ? <div style={{ display: "flex", gap: 6 }}>
                                      <Btn danger small onClick={() => setDesignatedRejectOkr({ id: s.id, actual: "" })}>Reject</Btn>
                                      <Btn primary small onClick={() => dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "approved", approvedBy: user.id })}>Approve</Btn>
                                    </div>
                                  : <Tag type={s.approval === "approved" ? "approved" : "rejected"} label={s.approval === "approved" ? "Approved" : "Rejected"} small />}
                              </div>
                            </div>
                            {designatedRejectOkr?.id === s.id && (
                              <div style={{ margin: "0 16px 10px 19px", padding: "10px 12px", background: T.badDim, borderRadius: 7, border: `1px solid ${T.badBorder}` }}>
                                <div style={{ fontSize: 12, fontWeight: 700, color: T.bad, marginBottom: 6 }}>Enter actual value for rejection</div>
                                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                                  <Input value={designatedRejectOkr.actual} onChange={e => setDesignatedRejectOkr(p => ({ ...p, actual: e.target.value }))} placeholder="Actual value" style={{ width: 120, textAlign: "right", fontFamily: F.mono }} />
                                  {s.krUnit && <span style={{ fontSize: 13, color: T.textMuted }}>{s.krUnit}</span>}
                                  {s.krType === "progress" ? <span style={{ fontSize: 12, color: T.textMuted }}>(target: {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span> : <span style={{ fontSize: 12, color: T.textMuted }}>(performance target: {s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""})</span>}
                                </div>
                                <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                                  <Btn small onClick={() => setDesignatedRejectOkr(null)}>Cancel</Btn>
                                  <Btn danger small onClick={() => { dispatch({ type: "APPROVE_OKR_SUBMISSION", id: s.id, status: "rejected", approvedBy: user.id, actualValue: Number(designatedRejectOkr.actual) || 0 }); setDesignatedRejectOkr(null); }}>Confirm Reject</Btn>
                                </div>
                              </div>
                            )}
                          </div>
                        );
                      })}
                    </Card>
                  );
                })}
                {designatedApproveeSubs.filter(s => s.answer === null).length > 0 && (
                  <div style={{ fontSize: 13, color: T.textMuted, padding: "8px 4px" }}>
                    {designatedApproveeSubs.filter(s => s.answer === null).length} check-in{designatedApproveeSubs.filter(s => s.answer === null).length !== 1 ? "s" : ""} awaiting staff response
                  </div>
                )}
              </>);
            })()}
            {(() => {
              const getCompletedYear = p => {
                if (p.completedYear) return p.completedYear;
                const parts = (p.updatedDate || "").split("/");
                return parts.length >= 3 ? parseInt(parts[2].split(",")[0].trim()) : null;
              };
              const ppMembers = designatedApproveeIds
                .map(memberId => {
                  const ppKrs = (memberData[memberId]?.krs || []).filter(kr => kr.type === "project_profit");
                  if (!ppKrs.length) return null;
                  return { memberId, mem: users.find(u => u.id === memberId), ppKrs };
                })
                .filter(Boolean);
              if (!ppMembers.length) return null;
              return (
                <div style={{ marginTop: 20 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10, paddingBottom: 6, borderBottom: `2px solid ${T.okBorder}` }}>
                    <div style={{ width: 4, height: 18, background: T.ok, borderRadius: 2 }} />
                    <span style={{ fontSize: 15, fontWeight: 700, color: T.text }}>Project Profit KRs</span>
                    <span style={{ fontSize: 11, color: T.textMuted }}>auto-tracked from completed projects</span>
                  </div>
                  {ppMembers.map(({ memberId, mem, ppKrs }) => (
                    <Card key={memberId} style={{ marginBottom: 10, overflow: "hidden" }}>
                      <div style={{ padding: "11px 16px", background: T.raised, borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                        <Avatar letters={mem?.av || "?"} size={28} />
                        <span style={{ fontWeight: 700, fontSize: 14, color: T.text }}>{mem?.name || "Unknown"}</span>
                        <span style={{ fontSize: 11, color: T.textDim, marginLeft: "auto" }}>{ppKrs.length} KR{ppKrs.length !== 1 ? "s" : ""}</span>
                      </div>
                      {ppKrs.map((kr, idx) => {
                        const yearProjects = projects.filter(p => p.mgrId === memberId && p.status === "completed" && getCompletedYear(p) === kr.krYear);
                        const actual = yearProjects.reduce((s, p) => s + (p.income != null && p.margin != null ? Math.round(p.income * p.margin * (p.contributeRate ?? 100) / 10000) : 0), 0);
                        const pct = kr.target > 0 ? Math.min(Math.round(actual / kr.target * 100), 100) : 0;
                        const st = getStatus(pct);
                        const missingMargin = yearProjects.filter(p => p.income == null || p.margin == null).length;
                        const hasPartialRate = yearProjects.some(p => p.contributeRate != null && p.contributeRate < 100);
                        return (
                          <div key={kr.id} style={{ padding: "12px 16px 12px 19px", borderLeft: `3px solid ${T.ok}`, borderBottom: idx < ppKrs.length - 1 ? `1px solid ${T.border}` : "none" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 4 }}>
                                  <span style={{ fontSize: 13, fontWeight: 600, color: T.text }}>{kr.label}</span>
                                  <span style={{ fontSize: 10, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 8, padding: "1px 5px", flexShrink: 0 }}>Proj Profit</span>
                                </div>
                                <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 6 }}>Target: ${(kr.target || 0).toLocaleString()} profit · Year {kr.krYear || "?"}</div>
                                <Bar value={pct} status={st} h={5} />
                                {missingMargin > 0 && <div style={{ marginTop: 4, fontSize: 11, color: T.warn }}>⚠ {missingMargin} project{missingMargin !== 1 ? "s" : ""} missing income/margin — not counted</div>}
                                {hasPartialRate && <div style={{ marginTop: 3, fontSize: 11, color: T.brand }}>⚡ Partial rates applied</div>}
                              </div>
                              <div style={{ textAlign: "right", flexShrink: 0 }}>
                                <div style={{ fontSize: 16, fontWeight: 800, color: st === "met" ? T.ok : st === "near" ? T.warn : T.bad, fontFamily: F.mono }}>${actual.toLocaleString()}</div>
                                <div style={{ fontSize: 11, color: T.textMuted }}>of ${(kr.target || 0).toLocaleString()}</div>
                                <div style={{ fontSize: 12, fontWeight: 700, color: st === "met" ? T.ok : st === "near" ? T.warn : T.bad }}>{pct}%</div>
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </Card>
                  ))}
                </div>
              );
            })()}
          </Pane>
        </>)}

        {page === "history" && (<>
          <Header title="My OKR Check-In History" sub="All check-in submissions and their approval status" />
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
                  const ansCol = s.answer === "yes" ? T.ok : s.answer === "no" ? T.bad : s.answer === "submitted" ? (s.krType === "progress" ? T.brand : "#7c3aed") : T.textDim;
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
                            {s.krType === "progress" && <span style={{ fontSize: 10, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Progress</span>}
                            {s.managerFilled && <span style={{ fontSize: 10, color: "#d97706", background: "#fef3c7", border: "1px solid #fde68a", borderRadius: 5, padding: "1px 5px", fontWeight: 700 }}>Mgr Assessed</span>}
                          </div>
                          <div style={{ marginTop: 8, fontSize: 13, fontWeight: 700, color: ansCol }}>{s.managerFilled ? (s.answer === "yes" ? "✓ Manager assessed: Met" : s.answer === "no" ? "✗ Manager assessed: Not met" : ansLabel) : ansLabel}</div>
                          {s.answer === "no" && s.actualValue != null && (
                            <div style={{ fontSize: 12, color: T.textMuted, marginTop: 3 }}>
                              Actual: <span style={{ fontFamily: F.mono, fontWeight: 700, color: T.bad }}>{s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                              <span style={{ margin: "0 6px" }}>·</span>
                              Performance Target: <span style={{ fontFamily: F.mono }}>{s.krOperator || ">="} {s.krTarget != null ? s.krTarget : "—"}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
                            </div>
                          )}
                          {s.answer === "yes" && s.krTarget != null && (
                            <div style={{ fontSize: 12, color: T.textMuted, marginTop: 3 }}>
                              {s.actualValue != null && <><span style={{ fontFamily: F.mono, fontWeight: 700, color: T.ok }}>{s.actualValue}{s.krUnit ? ` ${s.krUnit}` : ""}</span><span style={{ margin: "0 6px" }}>·</span></>}
                              Performance Target: <span style={{ fontFamily: F.mono }}>{s.krOperator || ">="} {s.krTarget}{s.krUnit ? ` ${s.krUnit}` : ""}</span>
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
                          {s.approvedBy === "auto" && <span style={{ fontSize: 9, fontWeight: 700, background: "rgba(16,185,129,0.1)", border: "1px solid rgba(16,185,129,0.25)", color: "#10B981", borderRadius: 8, padding: "1px 6px", letterSpacing: "0.05em" }}>AUTO</span>}
                          {s.approvedBy && s.approvedBy !== "auto" && <span style={{ fontSize: 10, color: T.textMuted }}>by {users?.find(u => u.id === s.approvedBy)?.name || "Admin"}</span>}
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

        {page === "projects" && (<>
          <Header title="Projects" sub="Create and track your projects" />
          <Pane>
            {myProjects.length > 0 && (
              <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 6 }}>
                <Metric label="Total" value={myProjects.length} />
                <Metric label="Active" value={myProjects.filter(p => p.status === "active").length} status="yellow" />
                <Metric label="Pending Approval" value={myProjects.filter(p => p.status === "pending approval").length} status="blue" />
                <Metric label="Completed" value={myProjects.filter(p => p.status === "completed").length} status="green" />
                <Metric label="Avg Progress" value={`${Math.round(myProjects.reduce((a, p) => a + p.progress, 0) / myProjects.length)}%`} />
                {(() => { const ti = myProjects.filter(p => p.status !== "completed").reduce((a, p) => a + (p.income || 0), 0); return ti > 0 ? <Metric label="Estimated Income" value={`$${ti.toLocaleString()}`} status="blue" /> : null; })()}
                {(() => { const tp = myProjects.filter(p => p.status !== "completed").reduce((a, p) => a + (p.income != null && p.margin != null ? Math.round(p.income * p.margin / 100) : 0), 0); return tp > 0 ? <Metric label="Estimated Profit" value={`$${tp.toLocaleString()}`} status="blue" /> : null; })()}
                {(() => { const ti = myProjects.filter(p => p.status === "completed").reduce((a, p) => a + (p.income || 0), 0); return ti > 0 ? <Metric label="Completed Income" value={`$${ti.toLocaleString()}`} status="green" /> : null; })()}
                {(() => { const tp = myProjects.filter(p => p.status === "completed").reduce((a, p) => a + (p.income != null && p.margin != null ? Math.round(p.income * p.margin / 100) : 0), 0); return tp > 0 ? <Metric label="Completed Profit" value={`$${tp.toLocaleString()}`} status="green" /> : null; })()}
              </div>
            )}
            {user.projectAccess && (
              <div style={{ marginBottom: 16 }}>
                {!showNewProj ? (
                  <Btn primary onClick={() => setShowNewProj(true)}>+ New Project</Btn>
                ) : (
                  <Card style={{ padding: 16 }}>
                    <div style={{ fontSize: 14, fontWeight: 700, marginBottom: 12 }}>New Project</div>
                    <div style={{ display: "flex", gap: 12, alignItems: "flex-end", flexWrap: "wrap" }}>
                      <div style={{ flex: 1, minWidth: 180 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Name</div>
                        <Input value={newProj.name} onChange={e => setNewProj(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Term 2 NAPLAN Prep" style={{ width: "100%" }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Start Date</div>
                        <Input type="date" value={newProj.startDate} onChange={e => setNewProj(p => ({ ...p, startDate: e.target.value }))} style={{ width: 160 }} />
                      </div>
                      <div>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                        <Input type="date" value={newProj.due} onChange={e => setNewProj(p => ({ ...p, due: e.target.value }))} style={{ width: 160 }} />
                      </div>
                      <Btn primary disabled={!newProj.name.trim() || !newProj.startDate} onClick={() => { dispatch({ type: "ADD_PROJECT", project: { id: `p${Date.now()}`, mgrId: user.id, name: newProj.name.trim(), status: "active", startDate: newProj.startDate, due: newProj.due || "TBD", progress: 0, income: null, margin: null } }); setNewProj({ name: "", startDate: "", due: "" }); setShowNewProj(false); }}>Create</Btn>
                      <Btn onClick={() => { setNewProj({ name: "", startDate: "", due: "" }); setShowNewProj(false); }}>Cancel</Btn>
                    </div>
                  </Card>
                )}
              </div>
            )}
            {myProjects.length === 0 && <EmptyState text={user.projectAccess ? "No projects yet. Click '+ New Project' to get started." : "No projects assigned."} />}
            {myProjects.map(p => {
              const draftProg = progressEdits[p.id] ?? p.progress;
              const ps = draftProg >= 70 ? "green" : draftProg >= 35 ? "yellow" : "red";
              const progChanged = progressEdits[p.id] !== undefined;
              const isDetailsOpen = editProjId === p.id;
              return (
                <Card key={p.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                  <div style={{ padding: "12px 18px", display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div>
                      <div style={{ fontSize: 15, fontWeight: 700 }}>{p.name}</div>
                      <div style={{ fontSize: 12, color: T.textMuted }}>{p.mgrId !== user.id && (() => { const owner = users.find(u => u.id === p.mgrId); return owner ? <span>Owner: {owner.name} · </span> : null; })()}{p.startDate ? `Start: ${p.startDate} · ` : ""}Due: {p.due}{p.updatedDate ? ` · Updated: ${p.updatedDate}` : ""}</div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <Tag type={p.status === "active" ? "pending" : p.status === "pending approval" ? "review" : "approved"} label={p.status === "active" ? "ACTIVE" : p.status === "pending approval" ? "PENDING APPROVAL" : "COMPLETED"} small />
                      {user.projectAccess && p.mgrId === user.id && <button onClick={() => { if (window.confirm(`Delete project "${p.name}"? This cannot be undone.`)) dispatch({ type: "REMOVE_PROJECT", projectId: p.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete project">✕</button>}
                    </div>
                  </div>
                  <div style={{ padding: "10px 18px", display: "flex", alignItems: "center", gap: 10, borderTop: `1px solid ${T.border}` }}>
                    <Bar value={draftProg} status={ps} h={6} />
                    {p.mgrId === user.id ? <>
                      <Input value={draftProg} onChange={e => setProgressEdits(d => ({ ...d, [p.id]: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 52, textAlign: "right", padding: "5px 8px", fontSize: 14, fontFamily: F.mono }} />
                      <span style={{ fontSize: 13, color: T.textMuted }}>%</span>
                      {p.income != null && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Income: ${p.income.toLocaleString()}</span>}
                      {p.income != null && p.margin != null && <span style={{ fontSize: 11, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Profit: ${Math.round(p.income * p.margin / 100).toLocaleString()} ({p.margin}%)</span>}
                      <Btn primary small disabled={!progChanged} onClick={() => {
                        if (draftProg === 100 && p.status === "active") {
                          if (!window.confirm("Progress is now 100%. Mark as completed?")) {
                            setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                            return;
                          }
                          const updDate = new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { progress: draftProg, updatedDate: updDate } });
                          setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                          const submit = window.confirm("Submit for System Admin approval?");
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: submit ? "pending approval" : "active" } });
                        } else {
                          dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { progress: draftProg, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          setProgressEdits(d => { const n = { ...d }; delete n[p.id]; return n; });
                        }
                      }}>Save</Btn>
                    </> : <>
                      <span style={{ fontSize: 13, color: T.textMuted, fontFamily: F.mono }}>{draftProg}%</span>
                      {p.income != null && <span style={{ fontSize: 11, color: T.brand, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Income: ${p.income.toLocaleString()}</span>}
                      {p.income != null && p.margin != null && <span style={{ fontSize: 11, color: T.ok, background: T.okDim, border: `1px solid ${T.okBorder}`, borderRadius: 6, padding: "2px 8px", fontFamily: F.mono, fontWeight: 700, whiteSpace: "nowrap" }}>Profit: ${Math.round(p.income * p.margin / 100).toLocaleString()} ({p.margin}%)</span>}
                    </>}
                  </div>
                  {!isDetailsOpen && (() => { const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []); const latest = entries[0]; if (!latest) return null; const truncated = latest.text.length > 160; const preview = truncated ? latest.text.slice(0, 160) + "…" : latest.text; return <div style={{ padding: "6px 18px 4px", fontSize: 13, color: T.textSoft, lineHeight: 1.5 }}>{latest.date && <span style={{ fontSize: 11, color: T.textMuted, marginRight: 6 }}>{latest.date}</span>}{preview}{truncated && <button onClick={e => { e.stopPropagation(); setLogPopup({ text: latest.text, date: latest.date, projName: p.name }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.brand, fontSize: 12, fontWeight: 700, padding: "0 0 0 4px", fontFamily: F.body }}>Read more →</button>}</div>; })()}
                  {user.projectAccess && p.mgrId === user.id && (
                    <div style={{ padding: "6px 18px 10px" }}>
                      <button onClick={() => {
                        if (isDetailsOpen) { setEditProjId(null); return; }
                        setEditProjId(p.id);
                        setEditProjForm({ status: p.status, startDate: p.startDate || "", due: p.due || "", income: p.income != null ? String(p.income) : "", margin: p.margin != null ? String(p.margin) : "", contributeRate: p.contributeRate != null ? String(p.contributeRate) : "" });
                      }} style={{ background: "none", border: `1px solid ${T.border}`, borderRadius: 5, padding: "4px 10px", cursor: "pointer", color: T.textMuted, fontSize: 12, fontFamily: F.body }}>
                        {isDetailsOpen ? "▼ Edit Details" : "▸ Edit Details"}
                      </button>
                    </div>
                  )}
                  {isDetailsOpen && (
                    <div style={{ padding: "14px 18px", borderTop: `1px solid ${T.border}` }}>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                          <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 14, fontFamily: F.body }}>
                            <option value="active">Active</option>
                            <option value="pending approval">Pending Approval</option>
                            <option value="completed">Completed</option>
                          </select>
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Start Date</div>
                          <Input type="date" value={editProjForm.startDate} onChange={e => setEditProjForm(f => ({ ...f, startDate: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                          <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 14 }} />
                        </div>
                      </div>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Income ($)</div>
                          <Input type="number" value={editProjForm.income} onChange={e => setEditProjForm(f => ({ ...f, income: e.target.value }))} placeholder="e.g. 250000" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Profit Margin (%)</div>
                          <Input type="number" value={editProjForm.margin} onChange={e => setEditProjForm(f => ({ ...f, margin: e.target.value }))} placeholder="e.g. 30" min="0" max="100" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Contribution Rate (%)</div>
                          <Input type="number" value={editProjForm.contributeRate} onChange={e => setEditProjForm(f => ({ ...f, contributeRate: e.target.value }))} placeholder="100 (default)" min="1" max="100" style={{ width: "100%", padding: "7px 10px", fontSize: 14, fontFamily: F.mono }} />
                          <div style={{ fontSize: 10, color: T.textMuted, marginTop: 3 }}>Leave blank if sole owner</div>
                        </div>
                        <div>
                          <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Profit (Auto)</div>
                          {(() => {
                            const hasData = editProjForm.income !== "" && editProjForm.margin !== "";
                            const fullProfit = hasData ? Math.round(Number(editProjForm.income) * Number(editProjForm.margin) / 100) : null;
                            const rate = editProjForm.contributeRate !== "" ? Math.min(100, Math.max(1, Number(editProjForm.contributeRate))) : 100;
                            const myProfit = fullProfit != null ? Math.round(fullProfit * rate / 100) : null;
                            return (
                              <div style={{ padding: "7px 10px", fontSize: 14, fontFamily: F.mono, color: hasData ? T.ok : T.textMuted, fontWeight: 700 }}>
                                {fullProfit != null ? `$${fullProfit.toLocaleString()}` : "—"}
                                {fullProfit != null && rate < 100 && <div style={{ fontSize: 11, color: T.brand, fontWeight: 600, marginTop: 2 }}>Your KR: ${myProfit.toLocaleString()} ({rate}%)</div>}
                              </div>
                            );
                          })()}
                        </div>
                      </div>
                      <div style={{ marginBottom: 16 }}>
                        <div style={{ fontSize: 11, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>Project Logbook</div>
                        <TextArea value={logDrafts[p.id] || ""} onChange={e => setLogDrafts(d => ({ ...d, [p.id]: e.target.value }))} placeholder="Add a log entry..." rows={2} />
                        <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 6, marginBottom: 10 }}>
                          <Btn primary small disabled={!logDrafts[p.id]?.trim()} onClick={() => {
                            const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []);
                            dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { log: [{ text: logDrafts[p.id].trim(), date: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) }, ...entries] } });
                            setLogDrafts(d => { const n = { ...d }; delete n[p.id]; return n; });
                          }}>Add Entry</Btn>
                        </div>
                        {(() => { const entries = Array.isArray(p.log) ? p.log : (p.log ? [{ text: p.log, date: "" }] : []); if (!entries.length) return <div style={{ fontSize: 12, color: T.textMuted }}>No log entries yet.</div>; return entries.map((e, i) => <div key={i} style={{ padding: "8px 10px", marginBottom: 6, background: T.bg, borderRadius: 6, border: `1px solid ${T.border}` }}>{e.date && <div style={{ fontSize: 11, color: T.textMuted, marginBottom: 3 }}>{e.date}</div>}<div style={{ fontSize: 13, color: T.text, lineHeight: 1.5, whiteSpace: "pre-wrap", maxHeight: 200, overflowY: "auto" }}>{e.text}</div></div>); })()}
                      </div>
                      <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
                        <Btn small onClick={() => setEditProjId(null)}>Cancel</Btn>
                        <Btn primary small onClick={() => {
                          const cr = editProjForm.contributeRate !== "" ? Math.min(100, Math.max(1, Number(editProjForm.contributeRate))) : null;
                          if (editProjForm.status === "completed" && p.status !== "completed") {
                            const submit = window.confirm("Submit for System Admin approval?");
                            dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: submit ? "pending approval" : "active", startDate: editProjForm.startDate || p.startDate || "", due: editProjForm.due || p.due, income: editProjForm.income !== "" ? Number(editProjForm.income) : null, margin: editProjForm.margin !== "" ? Math.min(100, Math.max(0, Number(editProjForm.margin))) : null, contributeRate: cr, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          } else {
                            dispatch({ type: "UPDATE_PROJECT", projectId: p.id, updates: { status: editProjForm.status, startDate: editProjForm.startDate || p.startDate || "", due: editProjForm.due || p.due, income: editProjForm.income !== "" ? Number(editProjForm.income) : null, margin: editProjForm.margin !== "" ? Math.min(100, Math.max(0, Number(editProjForm.margin))) : null, contributeRate: cr, updatedDate: new Date().toLocaleString("en-AU", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" }) } });
                          }
                          setEditProjId(null);
                        }}>Save Details</Btn>
                      </div>
                    </div>
                  )}
                </Card>
              );
            })}
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
                    <Tag type={getStatus(r.data.companyRate)} label={`Company: ${Number(r.data.companyRate).toFixed(1)}%`} />
                  </div>
                  <div style={{ padding: "14px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                    <div>
                      <SectionLabel>Department Rankings</SectionLabel>
                      {r.data.deptRanks.map((d, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 14 }}>
                          <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                          <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                          <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{Number(d.rate).toFixed(1)}%</span>
                        </div>
                      ))}
                      {r.submissionRate != null && (
                        <div style={{ marginTop: 10 }}>
                          <span style={{ fontSize: 12, background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 8, padding: "2px 8px", color: T.brand }}>Target met rate: {r.submissionRate}%</span>
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
    </MobileContext.Provider>
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
    case "REPLACE_KR": return { ...state, depts: state.depts.map(d => {
      if (d.id !== action.deptId) return d;
      if (!action.teamId) return { ...d, krs: d.krs.map(kr => kr.id !== action.krId ? kr : action.kr) };
      return { ...d, teams: d.teams.map(t => t.id !== action.teamId ? t : { ...t, krs: t.krs.map(kr => kr.id !== action.krId ? kr : action.kr) }) };
    })};
    case "SYNC_DEPT_KRS_TO_MEMBERS": {
      const dept = state.depts.find(d => d.id === action.deptId);
      if (!dept) return state;
      const memberIds = state.users.filter(u => u.deptId === action.deptId && (u.role === "member" || u.role === "manager")).map(u => u.id);
      const krsToSync = dept.krs || [];
      const newMemberData = { ...state.memberData };
      for (const memberId of memberIds) {
        const md = newMemberData[memberId] || { krs: [] };
        const existing = md.krs || [];
        // Update metadata only. Reset actual: 0 → null (initial default, not a real submission).
        // Non-zero actuals from real approved check-ins are preserved.
        const updated = existing.map(kr => { const dk = krsToSync.find(t => t.id === kr.id); if (!dk) return kr; const { actual: _a, monthlyActuals: _m, ...meta } = dk; const merged = { ...kr, ...meta, actual: (kr.actual === 0 || kr.actual == null) ? null : kr.actual }; if (!meta.monthlyTargets) { delete merged.monthlyTargets; delete merged.monthlyActuals; } return merged; });
        // New KRs start with null actual so krCompletion correctly treats them as not yet measured
        const added = krsToSync.filter(kr => !existing.some(e => e.id === kr.id)).map(({ actual: _a, monthlyActuals: _m, ...meta }) => ({ ...meta, actual: null, ...(meta.monthlyTargets ? { monthlyActuals: {} } : {}) }));
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
      const krsToSync = team.krs || [];
      const newMemberData = { ...state.memberData };
      for (const memberId of memberIds) {
        const md = newMemberData[memberId] || { krs: [] };
        const existing = md.krs || [];
        // Update metadata only. Reset actual: 0 → null (initial default, not a real submission).
        // Non-zero actuals from real approved check-ins are preserved.
        const updated = existing.map(kr => { const tk = krsToSync.find(t => t.id === kr.id); if (!tk) return kr; const { actual: _a, monthlyActuals: _m, ...meta } = tk; const merged = { ...kr, ...meta, actual: (kr.actual === 0 || kr.actual == null) ? null : kr.actual }; if (!meta.monthlyTargets) { delete merged.monthlyTargets; delete merged.monthlyActuals; } return merged; });
        // New KRs start with null actual so krCompletion correctly treats them as not yet measured
        const added = krsToSync.filter(kr => !existing.some(e => e.id === kr.id)).map(({ actual: _a, monthlyActuals: _m, ...meta }) => ({ ...meta, actual: null, ...(meta.monthlyTargets ? { monthlyActuals: {} } : {}) }));
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
    case "REPLACE_MEMBER_KR": {
      const md = state.memberData[action.memberId];
      if (!md) return state;
      return { ...state, memberData: { ...state.memberData, [action.memberId]: { ...md, krs: (md.krs || []).map(kr => kr.id !== action.krId ? kr : action.kr) } } };
    }
    case "MANAGER_SUBMIT_KR": {
      const { memberId, memberName, deptId, kr, period, periodKey, actualValue, approvedBy, newId } = action;
      const existingIdx = state.okrSubmissions.findIndex(s => s.memberId === memberId && s.krId === kr.id && s.periodKey === periodKey);
      const krTarget = kr.monthlyTargets ? (kr.monthlyTargets[periodKey] || 0) : (Number(kr.target) || 0);
      const now = new Date().toISOString();
      const baseSub = existingIdx >= 0 ? state.okrSubmissions[existingIdx] : null;
      const newSub = {
        id: baseSub ? baseSub.id : newId,
        memberId, memberName, deptId,
        krId: kr.id, krLabel: kr.label,
        krTarget, krUnit: kr.unit || "", krOperator: kr.operator || ">=",
        krType: kr.type || "", krIsMonthly: !!(kr.monthlyTargets),
        period, periodKey, dateRange: "",
        sentAt: baseSub ? baseSub.sentAt : now,
        answeredAt: now,
        answer: "yes",
        actualValue,
        approval: "approved",
        approvedBy,
        reason: null,
      };
      const newSubs = existingIdx >= 0
        ? state.okrSubmissions.map((s, i) => i === existingIdx ? newSub : s)
        : [...state.okrSubmissions, newSub];
      const md2 = state.memberData[memberId];
      const newMemberData = md2 ? {
        ...state.memberData,
        [memberId]: { ...md2, krs: (md2.krs || []).map(k => {
          if (k.id !== kr.id) return k;
          if (k.monthlyTargets) return { ...k, monthlyActuals: { ...(k.monthlyActuals || {}), [periodKey]: actualValue } };
          return { ...k, actual: actualValue };
        })}
      } : state.memberData;
      const allApproved = newSubs.filter(s => s.krId === kr.id && s.periodKey === periodKey && s.deptId === deptId && s.approval === "approved" && s.actualValue != null);
      const avgActual = allApproved.length ? allApproved.reduce((sum, s) => sum + (Number(s.actualValue) || 0), 0) / allApproved.length : actualValue;
      const newDepts = state.depts.map(d => {
        if (d.id !== deptId) return d;
        const dKrs2 = (d.krs || []).map(k => {
          if (k.id !== kr.id) return k;
          if (k.monthlyTargets) return { ...k, monthlyActuals: { ...(k.monthlyActuals || {}), [periodKey]: avgActual } };
          return { ...k, actual: avgActual };
        });
        const dTeams2 = (d.teams || []).map(t => ({ ...t, krs: (t.krs || []).map(k => {
          if (k.id !== kr.id) return k;
          if (k.monthlyTargets) return { ...k, monthlyActuals: { ...(k.monthlyActuals || {}), [periodKey]: avgActual } };
          return { ...k, actual: avgActual };
        })}));
        return { ...d, krs: dKrs2, teams: dTeams2 };
      });
      return { ...state, okrSubmissions: newSubs, memberData: newMemberData, depts: newDepts };
    }
    case "MANAGER_ASSESS_KR": {
      const { memberId, memberName, deptId, kr, period, periodKey, answer, actualValue, approvedBy, newId } = action;
      const existingIdx = state.okrSubmissions.findIndex(s => s.memberId === memberId && s.krId === kr.id && s.periodKey === periodKey);
      const now = new Date().toISOString();
      const baseSub = existingIdx >= 0 ? state.okrSubmissions[existingIdx] : null;
      const resolvedActual = actualValue != null ? actualValue : (answer === "yes" ? (Number(kr.target) || 0) : 0);
      const newSub = {
        id: baseSub ? baseSub.id : newId,
        memberId, memberName, deptId,
        krId: kr.id, krLabel: kr.label,
        krTarget: Number(kr.target) || 0, krUnit: kr.unit || "", krOperator: kr.operator || ">=",
        krType: "manager-fill", krIsMonthly: false,
        period, periodKey, dateRange: "",
        sentAt: baseSub ? baseSub.sentAt : now,
        answeredAt: now,
        answer,
        actualValue: resolvedActual,
        approval: "approved",
        approvedBy,
        reason: null,
        managerFilled: true,
      };
      const newSubs = existingIdx >= 0
        ? state.okrSubmissions.map((s, i) => i === existingIdx ? newSub : s)
        : [...state.okrSubmissions, newSub];
      const md2 = state.memberData[memberId];
      const newMemberData = md2 ? {
        ...state.memberData,
        [memberId]: { ...md2, krs: (md2.krs || []).map(k => k.id !== kr.id ? k : { ...k, actual: resolvedActual }) }
      } : state.memberData;
      return { ...state, okrSubmissions: newSubs, memberData: newMemberData };
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
    case "LOG_EMAIL_SEND": {
      const logs = [action.log, ...(state.emailLogs || [])].slice(0, 100);
      return { ...state, emailLogs: logs };
    }
    case "UPDATE_EMAIL_LOG_RECIPIENT": {
      return { ...state, emailLogs: (state.emailLogs || []).map(log => {
        if (log.id !== action.logId) return log;
        const newRecipients = (log.recipients || []).map(r =>
          r.email !== action.email ? r : { ...r, success: action.success, reason: action.reason || null }
        );
        const newFailCount = newRecipients.filter(r => !r.success).length;
        return { ...log, recipients: newRecipients, failureCount: newFailCount };
      })};
    }
    case "CREATE_OKR_SUBMISSIONS": {
      const seen = new Set((state.okrSubmissions || []).map(s => `${s.memberId}:${s.krId}:${s.periodKey}`));
      const fresh = action.submissions.filter(s => !seen.has(`${s.memberId}:${s.krId}:${s.periodKey}`));
      return { ...state, okrSubmissions: [...(state.okrSubmissions || []), ...fresh] };
    }
    case "ANSWER_OKR_SUBMISSION": {
      const answerSub = (state.okrSubmissions || []).find(s => s.id === action.id);
      const answered = { ...answerSub, answer: action.answer, answeredAt: new Date().toISOString(), reason: action.reason || null, actualValue: action.actualValue ?? null };
      if (answerSub?.krType === "tracker") {
        const trackerSub = { ...answered, approval: "approved", approvedBy: "auto" };
        const newTrackerSubs = (state.okrSubmissions || []).map(s => s.id === action.id ? trackerSub : s);
        const trackerActual = trackerSub.actualValue;
        if (trackerActual == null) return { ...state, okrSubmissions: newTrackerSubs };
        const trackerMd = state.memberData[answerSub.memberId];
        if (!trackerMd) return { ...state, okrSubmissions: newTrackerSubs };
        const trackerMk = (answerSub.periodKey || "").slice(0, 7);
        const trackerMemberKrs = (trackerMd.krs || []).map(k => {
          if (k.id !== answerSub.krId) return k;
          if (k.monthlyTargets) return { ...k, monthlyActuals: { ...(k.monthlyActuals || {}), [trackerMk]: trackerActual } };
          return { ...k, actual: trackerActual };
        });
        return { ...state, okrSubmissions: newTrackerSubs, memberData: { ...state.memberData, [answerSub.memberId]: { ...trackerMd, krs: trackerMemberKrs } } };
      }
      if (!answerSub || action.answer !== "yes") {
        return { ...state, okrSubmissions: (state.okrSubmissions || []).map(s => s.id === action.id ? answered : s) };
      }
      const aDept = state.depts.find(d => d.id === answerSub.deptId);
      const aKr = aDept ? (aDept.krs.find(k => k.id === answerSub.krId) || aDept.teams.flatMap(t => t.krs || []).find(k => k.id === answerSub.krId)) : null;
      if (!aKr?.autoApprove) {
        return { ...state, okrSubmissions: (state.okrSubmissions || []).map(s => s.id === action.id ? answered : s) };
      }
      const autoSub = { ...answered, approval: "approved", approvedBy: "auto" };
      const newAutoSubs = (state.okrSubmissions || []).map(s => s.id === action.id ? autoSub : s);
      const autoActual = autoSub.actualValue ?? autoSub.krTarget;
      if (autoActual == null) return { ...state, okrSubmissions: newAutoSubs };
      const autoMd = state.memberData[answerSub.memberId];
      if (!autoMd) return { ...state, okrSubmissions: newAutoSubs };
      const autoMk = (answerSub.periodKey || "").slice(0, 7);
      const autoMemberKrs = (autoMd.krs || []).map(k => {
        if (k.id !== answerSub.krId) return k;
        if (k.monthlyTargets) return { ...k, monthlyActuals: { ...(k.monthlyActuals || {}), [autoMk]: autoActual } };
        return { ...k, actual: autoActual };
      });
      const autoMemberData = { ...state.memberData, [answerSub.memberId]: { ...autoMd, krs: autoMemberKrs } };
      const autoApprovedIds = [...new Set(newAutoSubs.filter(s => s.krId === answerSub.krId && s.approval === "approved" && s.periodKey === answerSub.periodKey && s.deptId === answerSub.deptId).map(s => s.memberId))];
      const autoIsMonthly = !!(autoMemberData[answerSub.memberId]?.krs?.find(k => k.id === answerSub.krId)?.monthlyTargets);
      const autoMemberVals = autoApprovedIds.map(mId => {
        const k = (autoMemberData[mId]?.krs || []).find(k => k.id === answerSub.krId);
        if (!k) return null;
        return autoIsMonthly ? ((k.monthlyActuals || {})[autoMk] ?? null) : (k.actual ?? null);
      }).filter(v => v !== null);
      const autoTeamActual = autoMemberVals.length > 0 ? Math.round(autoMemberVals.reduce((a, b) => a + b, 0) / autoMemberVals.length * 100) / 100 : autoActual;
      const autoUpdateKr = k => {
        if (k.id !== answerSub.krId) return k;
        if (k.monthlyTargets) return { ...k, monthlyActuals: { ...(k.monthlyActuals || {}), [autoMk]: autoTeamActual } };
        return { ...k, actual: autoTeamActual };
      };
      const autoDepts = state.depts.map(d => {
        if (d.id !== answerSub.deptId) return d;
        return { ...d, krs: d.krs.map(autoUpdateKr), teams: d.teams.map(t => ({ ...t, krs: (t.krs || []).map(autoUpdateKr) })) };
      });
      return { ...state, okrSubmissions: newAutoSubs, memberData: autoMemberData, depts: autoDepts };
    }
    case "APPROVE_OKR_SUBMISSION": {
      const newSubs = (state.okrSubmissions || []).map(s => s.id === action.id ? { ...s, approval: action.status, approvedBy: action.approvedBy } : s);
      const sub = (state.okrSubmissions || []).find(s => s.id === action.id);
      if (!sub) return { ...state, okrSubmissions: newSubs };
      const actualToWrite = action.status === "approved"
        ? (sub.actualValue ?? sub.krTarget)
        : (action.status === "rejected" && action.actualValue != null ? action.actualValue : null);
      if (actualToWrite === null) return { ...state, okrSubmissions: newSubs };
      const md = state.memberData[sub.memberId];
      if (!md) return { ...state, okrSubmissions: newSubs };
      // Update this member's personal KR
      const existingKrs = md.krs || [];
      const krExistsInMember = existingKrs.some(k => k.id === sub.krId);
      let updatedMemberKrs = existingKrs.map(kr => {
        if (kr.id !== sub.krId) return kr;
        if (kr.monthlyTargets) {
          const mk = (sub.periodKey || "").slice(0, 7);
          return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: actualToWrite } };
        }
        return { ...kr, actual: actualToWrite };
      });
      // If the KR wasn't in memberData yet (synced after the check-in was created), add it
      // so the approval is not silently dropped
      if (!krExistsInMember) {
        const deptKr = state.depts.flatMap(d => [...(d.krs || []), ...(d.teams || []).flatMap(t => t.krs || [])]).find(k => k.id === sub.krId);
        if (deptKr) {
          const { actual: _a, monthlyActuals: _m, ...meta } = deptKr;
          const newKr = meta.monthlyTargets
            ? { ...meta, monthlyActuals: { [(sub.periodKey || "").slice(0, 7)]: actualToWrite } }
            : { ...meta, actual: actualToWrite };
          updatedMemberKrs = [...updatedMemberKrs, newKr];
        }
      }
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
    case "EDIT_APPROVED_SUBMISSION": {
      const sub = (state.okrSubmissions || []).find(s => s.id === action.id);
      if (!sub) return state;
      const newActual = action.actualValue;
      const newSubs = (state.okrSubmissions || []).map(s => s.id === action.id ? { ...s, actualValue: newActual, ...(action.answer !== undefined ? { answer: action.answer } : {}) } : s);
      if (sub.approval !== "approved") return { ...state, okrSubmissions: newSubs };
      const md = state.memberData[sub.memberId];
      if (!md) return { ...state, okrSubmissions: newSubs };
      const mk = (sub.periodKey || "").slice(0, 7);
      const updatedMemberKrs = (md.krs || []).map(kr => {
        if (kr.id !== sub.krId) return kr;
        if (kr.monthlyTargets) return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: newActual } };
        return { ...kr, actual: newActual };
      });
      const newMemberData = { ...state.memberData, [sub.memberId]: { ...md, krs: updatedMemberKrs } };
      const approvedMemberIds = [...new Set(newSubs.filter(s => s.krId === sub.krId && s.approval === "approved" && s.periodKey === sub.periodKey && s.deptId === sub.deptId).map(s => s.memberId))];
      const isMonthly = !!(newMemberData[sub.memberId]?.krs?.find(k => k.id === sub.krId)?.monthlyTargets);
      const memberVals = approvedMemberIds.map(mId => {
        const kr = (newMemberData[mId]?.krs || []).find(k => k.id === sub.krId);
        if (!kr) return null;
        return isMonthly ? ((kr.monthlyActuals || {})[mk] ?? null) : (kr.actual ?? null);
      }).filter(v => v !== null);
      const teamActual = memberVals.length > 0
        ? Math.round(memberVals.reduce((a, b) => a + b, 0) / memberVals.length * 100) / 100
        : newActual;
      const updateDeptKr = kr => {
        if (kr.id !== sub.krId) return kr;
        if (kr.monthlyTargets) return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: teamActual } };
        return { ...kr, actual: teamActual };
      };
      const newDepts = state.depts.map(dept => {
        if (dept.id !== sub.deptId) return dept;
        return { ...dept, krs: dept.krs.map(updateDeptKr), teams: dept.teams.map(t => ({ ...t, krs: (t.krs || []).map(updateDeptKr) })) };
      });
      return { ...state, okrSubmissions: newSubs, memberData: newMemberData, depts: newDepts };
    }
    case "REMOVE_OKR_SUBMISSION": {
      const sub = (state.okrSubmissions || []).find(s => s.id === action.id);
      const remaining = (state.okrSubmissions || []).filter(s => s.id !== action.id);
      if (!sub || sub.approval !== "approved") return { ...state, okrSubmissions: remaining };
      const isMonthly = !!(state.memberData[sub.memberId]?.krs?.find(k => k.id === sub.krId)?.monthlyTargets);
      const mk = (sub.periodKey || "").slice(0, 7);
      // Revert member's personal KR to last remaining approved value, or null (no submission)
      const memberOtherApproved = remaining.filter(s => s.memberId === sub.memberId && s.krId === sub.krId && s.approval === "approved");
      const memberNewActual = memberOtherApproved.length > 0
        ? (memberOtherApproved[memberOtherApproved.length - 1].actualValue ?? memberOtherApproved[memberOtherApproved.length - 1].krTarget ?? null)
        : null;
      const md = state.memberData[sub.memberId];
      let newMemberData = state.memberData;
      if (md) {
        const updatedMemberKrs = (md.krs || []).map(kr => {
          if (kr.id !== sub.krId) return kr;
          if (kr.monthlyTargets) return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: memberNewActual } };
          return { ...kr, actual: memberNewActual };
        });
        newMemberData = { ...state.memberData, [sub.memberId]: { ...md, krs: updatedMemberKrs } };
      }
      // Recalculate team/dept KR from remaining approved submissions across all members
      const approvedMemberIds = [...new Set(remaining.filter(s =>
        s.krId === sub.krId && s.approval === "approved" && s.periodKey === sub.periodKey && s.deptId === sub.deptId
      ).map(s => s.memberId))];
      const teamVals = approvedMemberIds.map(mId => {
        const kr = (newMemberData[mId]?.krs || []).find(k => k.id === sub.krId);
        if (!kr) return null;
        return isMonthly ? ((kr.monthlyActuals || {})[mk] ?? null) : (kr.actual ?? null);
      }).filter(v => v !== null);
      const newTeamActual = teamVals.length > 0
        ? Math.round(teamVals.reduce((a, b) => a + b, 0) / teamVals.length * 100) / 100
        : null;
      const updateDeptKr = kr => {
        if (kr.id !== sub.krId) return kr;
        if (kr.monthlyTargets) return { ...kr, monthlyActuals: { ...(kr.monthlyActuals || {}), [mk]: newTeamActual } };
        return { ...kr, actual: newTeamActual };
      };
      const newDepts = state.depts.map(dept => {
        if (dept.id !== sub.deptId) return dept;
        return { ...dept, krs: dept.krs.map(updateDeptKr), teams: dept.teams.map(t => ({ ...t, krs: (t.krs || []).map(updateDeptKr) })) };
      });
      return { ...state, okrSubmissions: remaining, memberData: newMemberData, depts: newDepts };
    }
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
  const [originalAdmin, setOriginalAdmin] = useState(null);
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
    emailLogs: [],
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

  // Shared data processor: resets actual: 0 → null on every load (0 is the legacy init default,
  // not a real submitted value; approved submissions always write back a non-null actualValue).
  // Used by both the mount effect and reloadState.
  const applyLoadedData = useCallback((data) => {
    const okrSubs = data.okrSubmissions || [];
    const fixActuals = krs => (krs || []).map(kr =>
      kr.actual === 0 ? { ...kr, actual: null } : kr
    );
    const fixedDepts = (data.depts || []).map(d => ({
      ...d,
      krs: fixActuals(d.krs),
      teams: (d.teams || []).map(t => ({ ...t, krs: fixActuals(t.krs) })),
    }));
    const fixedMemberData = Object.fromEntries(
      (data.memberData || []).map(m => [m.id, { krs: fixActuals(m.krs || []) }])
    );
    rawDispatch(() => ({
      users: data.users,
      depts: fixedDepts,
      memberData: fixedMemberData,
      weeklySubs: data.weeklySubs || [],
      okrSubmissions: okrSubs,
      emailLogs: data.emailLogs || [],
      mgrSprints: data.mgrSprints || [],
      projects: data.projects || [],
      monthlyReports: data.monthlyReports || [],
      settings: data.settings?.[0] || { id: "settings", colOrder: ["id", "label", "operator", "period", "target", "actual", "unit", "dataSource"] },
    }));
  }, []); // eslint-disable-line

  const reloadState = useCallback(async () => {
    try {
      const data = await dbGet();
      if (data.users?.length) applyLoadedData(data);
    } catch (err) {
      console.error("Reload failed:", err);
    }
  }, [applyLoadedData]);

  // On mount: load all data from Supabase. Seed the DB with initial data if it is empty.
  useEffect(() => {
    dbGet()
      .then(data => {
        if (data.users?.length) {
          applyLoadedData(data);
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
    setOriginalAdmin(null);
    setMsalErr("");
    try { instance.clearCache(); } catch (_) {}
  };
  const onImpersonate = (targetUser) => { setOriginalAdmin(user); setUser(targetUser); };
  const onExitImpersonate = () => { setUser(originalAdmin); setOriginalAdmin(null); };

  // Always derive the active user from state.users so admin edits (role, title, dept, etc.) are reflected immediately.
  // Email fallback handles the case where INIT_USERS matched by email before the DB loaded, but the DB user has a different ID.
  const activeUser = state.users.find(u => u.id === user.id)
    || state.users.find(u => u.email && user.email && u.email.toLowerCase() === user.email.toLowerCase())
    || user;

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
  const impersonationBanner = originalAdmin ? (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, background: "#c85a00", color: "#fff", fontSize: 13, fontWeight: 600, padding: "9px 20px", display: "flex", alignItems: "center", justifyContent: "space-between", zIndex: 10000, boxShadow: "0 2px 8px rgba(0,0,0,0.25)" }}>
      <span>👁 Viewing as <strong>{activeUser.name}</strong> ({activeUser.role}) — all actions apply to this user's account</span>
      <button onClick={onExitImpersonate} style={{ background: "rgba(255,255,255,0.2)", border: "1px solid rgba(255,255,255,0.5)", borderRadius: 6, color: "#fff", fontSize: 12, fontWeight: 700, padding: "4px 14px", cursor: "pointer", fontFamily: "inherit" }}>
        ✕ Exit &amp; Return to Admin
      </button>
    </div>
  ) : null;

  if (activeUser.role === "admin")   return <>{offlineBanner}{syncErrToast}<AdminPortal   user={activeUser} onLogout={logout} state={state} dispatch={dispatch} onReload={reloadState} onImpersonate={onImpersonate} /></>;
  if (activeUser.role === "manager") return <>{offlineBanner}{syncErrToast}{impersonationBanner}<ManagerPortal user={activeUser} onLogout={originalAdmin ? onExitImpersonate : logout} state={state} dispatch={dispatch} onReload={reloadState} /></>;
  return <>{offlineBanner}{syncErrToast}{impersonationBanner}<MemberPortal user={activeUser} onLogout={originalAdmin ? onExitImpersonate : logout} state={state} dispatch={dispatch} onReload={reloadState} /></>;
}
