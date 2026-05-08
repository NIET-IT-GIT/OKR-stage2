import { useState, useEffect, useCallback, useRef } from "react";
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
function calcRate(krs) {
  if (!krs?.length) return 0;
  return krs.reduce((sum, kr) => sum + Math.min((kr.actual / kr.target) * 100, 100), 0) / krs.length;
}
function getStatus(r) { return r >= TP ? "green" : r >= 60 ? "yellow" : "red"; }
function fmt(v) { return typeof v === "number" ? (v % 1 ? v.toFixed(1) : v.toLocaleString()) : v; }
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
  const result = { users: [], depts: [], memberData: [], weeklySubs: [], mgrSprints: [], projects: [], monthlyReports: [] };
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
  for (const col of ["users", "depts", "weeklySubs", "mgrSprints", "projects", "monthlyReports"]) {
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
    <span style={{ display: "inline-flex", alignItems: "center", background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 20, padding: "2px 9px", fontSize: 9, fontWeight: 600, color: cfg.color, whiteSpace: "nowrap", letterSpacing: "0.02em" }}>
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
      <div style={{ fontSize: 10, color: T.textMuted, fontWeight: 500, letterSpacing: "0.04em", textTransform: "uppercase", marginBottom: 10 }}>{label}</div>
      <div style={{ fontSize: 24, fontWeight: 700, color: status ? STATUS_THEME[status]?.color : T.text, fontFamily: F.mono, letterSpacing: "-0.02em" }}>{value}</div>
      {sub && <div style={{ fontSize: 10, color: T.textDim, marginTop: 4 }}>{sub}</div>}
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
        padding: "9px 13px", color: T.text, fontSize: 13, fontFamily: F.body, outline: "none",
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
        fontSize: 13, fontFamily: F.body, outline: "none", cursor: "pointer",
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
        padding: "10px 13px", color: T.text, fontSize: 13, fontFamily: F.body, outline: "none", resize: "vertical",
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
  return <div style={{ fontSize: 10, fontWeight: 600, color: T.textMuted, letterSpacing: "0.05em", textTransform: "uppercase", marginBottom: 10, ...sx }}>{children}</div>;
}

function EmptyState({ text }) {
  return <div style={{ padding: "48px 20px", textAlign: "center", color: T.textDim, fontSize: 13, fontWeight: 400 }}>{text}</div>;
}

function CountBadge({ count, color }) {
  if (!count) return null;
  return <span style={{ background: color || T.bad, color: "#fff", borderRadius: 20, padding: "1px 6px", fontSize: 9, fontWeight: 700, marginLeft: 6, letterSpacing: "0.02em" }}>{count}</span>;
}

function Side({ items, active, onSelect, user, onLogout, pendingCounts }) {
  return (
    <div style={{
      width: 240, background: T.glass, borderRight: `1px solid ${T.border}`,
      backdropFilter: "saturate(180%) blur(20px)", WebkitBackdropFilter: "saturate(180%) blur(20px)",
      display: "flex", flexDirection: "column", height: "100vh", flexShrink: 0,
    }}>
      <div style={{ padding: "22px 16px 14px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 18 }}>
          <div style={{ width: 30, height: 30, borderRadius: 8, background: `linear-gradient(145deg, ${T.brand}, #0052A3)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 700, color: "#fff", boxShadow: "0 2px 8px rgba(0,113,227,0.28)" }}>O</div>
          <div>
            <div style={{ fontSize: 13, fontWeight: 700, color: T.text, letterSpacing: "-0.01em" }}>OKR Tracker</div>
            <div style={{ fontSize: 9, color: T.textDim, letterSpacing: "0.06em", textTransform: "uppercase" }}>NIET GROUP</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 11px", background: "rgba(0,0,0,0.03)", borderRadius: 11, border: `1px solid ${T.border}` }}>
          <Avatar letters={user.av} size={32} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 12, fontWeight: 600, color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", letterSpacing: "-0.01em" }}>{user.name}</div>
            <div style={{ fontSize: 10, color: T.textMuted }}>{user.title}</div>
          </div>
        </div>
      </div>
      <div style={{ flex: 1, padding: "10px 8px", display: "flex", flexDirection: "column", gap: 1, overflowY: "auto" }}>
        {items.map(item => (
          <button key={item.id} onClick={() => onSelect(item.id)} style={{
            background: active === item.id ? T.brandDim : "transparent",
            border: active === item.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
            borderRadius: 9, padding: "9px 12px", cursor: "pointer",
            display: "flex", alignItems: "center", gap: 9,
            color: active === item.id ? T.brand : T.textMuted,
            fontSize: 12, fontWeight: active === item.id ? 600 : 400, textAlign: "left", width: "100%",
            transition: "all 0.12s", fontFamily: F.body, letterSpacing: "-0.01em",
          }}>
            <span style={{ fontSize: 13, width: 18, textAlign: "center", flexShrink: 0, opacity: active === item.id ? 1 : 0.6 }}>{item.icon}</span>
            <span style={{ flex: 1 }}>{item.label}</span>
            {pendingCounts?.[item.id] > 0 && <CountBadge count={pendingCounts[item.id]} />}
          </button>
        ))}
      </div>
      <div style={{ padding: "10px 8px 14px", borderTop: `1px solid ${T.border}` }}>
        <button onClick={onLogout} style={{ background: "none", border: "none", borderRadius: 9, padding: "9px 12px", cursor: "pointer", display: "flex", alignItems: "center", gap: 9, color: T.textMuted, fontSize: 12, width: "100%", fontFamily: F.body, letterSpacing: "-0.01em" }}>
          <span style={{ fontSize: 13, width: 18, textAlign: "center", opacity: 0.6 }}>↩</span> Sign Out
        </button>
      </div>
    </div>
  );
}

function Header({ title, sub, right }) {
  return (
    <div style={{ padding: "24px 32px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "flex-start", justifyContent: "space-between", background: T.glass, backdropFilter: "saturate(180%) blur(20px)", WebkitBackdropFilter: "saturate(180%) blur(20px)", position: "sticky", top: 0, zIndex: 10 }}>
      <div>
        <h1 style={{ margin: 0, fontSize: 22, fontWeight: 700, color: T.text, letterSpacing: "-0.03em" }}>{title}</h1>
        {sub && <p style={{ margin: "3px 0 0", fontSize: 12, color: T.textMuted, fontWeight: 400 }}>{sub}</p>}
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
          <span style={{ fontSize: 22, color: T.bad }}>⚠</span>
          <div style={{ fontSize: 13, color: T.bad, textAlign: "center", maxWidth: 320 }}>{error}</div>
        </>
      ) : (
        <>
          <span style={{ fontSize: 22, color: T.brand, animation: "spin 1s linear infinite" }}>◌</span>
          <div style={{ fontSize: 13, color: T.textMuted }}>Loading data…</div>
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
  const [showAdminLogin, setShowAdminLogin] = useState(false);
  const [adminUser, setAdminUser] = useState("");
  const [adminPass, setAdminPass] = useState("");
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

  const handleAdminLogin = () => {
    setErr("");
    if (adminUser === "admin" && adminPass === "Ntr1#qez66") {
      const u = users.find(u => u.id === "sysadmin");
      onLogin(u);
    } else {
      setErr("Invalid credentials.");
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
            <p style={{ margin: "0 0 8px", fontSize: 13, color: T.textSoft, lineHeight: 1.65 }}>
              The account <span style={{ color: T.text, fontWeight: 600 }}>{msalErr}</span> is not registered in this system.
            </p>
            <p style={{ margin: "0 0 24px", fontSize: 13, color: T.textSoft, lineHeight: 1.65 }}>
              Please contact your <strong style={{ color: T.text }}>team manager</strong> or the <strong style={{ color: T.text }}>IT team</strong> to request access.
            </p>
            <Btn primary onClick={onDismissErr}>OK</Btn>
          </div>
        </div>
      )}
      {/* Hero */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 72px", position: "relative", zIndex: 2, opacity: show ? 1 : 0, transform: show ? "none" : "translateX(-20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 44 }}>
          <div style={{ width: 46, height: 46, borderRadius: 13, background: `linear-gradient(135deg, ${T.brand}, #7c5bf5)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, fontWeight: 900, color: "#fff" }}>O</div>
          <div>
            <div style={{ fontSize: 21, fontWeight: 900, color: "#fff" }}>NIET OKR Performance Tracker</div>
            <div style={{ fontSize: 10, color: "rgba(255,255,255,0.55)", letterSpacing: "0.14em" }}>NIET · CHARLTON BROWN · RHODES · EDUCARE</div>
          </div>
        </div>
        <h1 style={{ margin: "0 0 14px", fontSize: 44, fontWeight: 900, lineHeight: 1.08, color: "#fff", letterSpacing: "-0.03em", maxWidth: 460 }}>
          Align goals.<br /><span style={{ color: "#A78BFA" }}>Track everyone.</span><br />Drive results.
        </h1>
        <p style={{ margin: 0, fontSize: 14, color: "rgba(255,255,255,0.65)", lineHeight: 1.7, maxWidth: 400 }}>
          Monthly KPI reporting, weekly outcome tracking, real-time leaderboards — full transparency from company goals down to every team member.
        </p>
        <div style={{ marginTop: 48, display: "flex", gap: 36 }}>
          {[{ n: "Monthly", l: "KPI Reports" }, { n: "Weekly", l: "Submissions" }, { n: "Real-time", l: "Rankings" }, { n: "100%", l: "Transparent" }].map((x, i) => (
            <div key={i}>
              <div style={{ fontSize: 18, fontWeight: 900, color: "#A78BFA", fontFamily: F.mono }}>{x.n}</div>
              <div style={{ fontSize: 9, color: "rgba(255,255,255,0.5)", letterSpacing: "0.06em", marginTop: 2 }}>{x.l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Sign-in card */}
      <div style={{ width: 420, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 44px", position: "relative", zIndex: 2, opacity: show ? 1 : 0, transform: show ? "none" : "translateY(20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1) 0.15s" }}>
        <Card style={{ padding: "36px 30px" }}>
          <h2 style={{ margin: "0 0 4px", fontSize: 18, fontWeight: 800, color: T.text }}>Sign in</h2>
          <p style={{ margin: "0 0 24px", fontSize: 12, color: T.textMuted }}>Use your NIET Microsoft account to access your portal.</p>

          {err && (
            <div style={{ padding: "10px 14px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, fontSize: 12, color: T.bad, marginBottom: 16, lineHeight: 1.5 }}>{err}</div>
          )}

          {/* Microsoft button */}
          <button onClick={handleMicrosoftLogin} disabled={msLoading} style={{
            width: "100%", padding: "13px 16px", background: msLoading ? T.raised : "#fff",
            border: `1px solid ${T.border}`, borderRadius: 8, cursor: msLoading ? "not-allowed" : "pointer",
            display: "flex", alignItems: "center", justifyContent: "center", gap: 12,
            fontSize: 13, fontWeight: 700, color: "#1a1a1a", fontFamily: F.body, transition: "all 0.15s", opacity: msLoading ? 0.6 : 1,
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

          {/* Divider */}
          <div style={{ display: "flex", alignItems: "center", gap: 10, margin: "20px 0" }}>
            <div style={{ flex: 1, height: 1, background: T.border }} />
            <span style={{ fontSize: 10, color: T.textDim, fontWeight: 600, letterSpacing: "0.05em" }}>OR</span>
            <div style={{ flex: 1, height: 1, background: T.border }} />
          </div>

          {/* Admin credentials toggle */}
          <button onClick={() => { setShowAdminLogin(p => !p); setErr(""); }} style={{
            width: "100%", padding: "10px 14px", background: "transparent",
            border: `1px solid ${T.border}`, borderRadius: 8, cursor: "pointer",
            display: "flex", alignItems: "center", justifyContent: "space-between",
            fontSize: 12, fontWeight: 600, color: T.textSoft, fontFamily: F.body, transition: "all 0.12s",
          }}
            onMouseEnter={e => e.currentTarget.style.borderColor = T.brandBorder}
            onMouseLeave={e => e.currentTarget.style.borderColor = T.border}
          >
            <span>Sign in with admin credentials</span>
            <span style={{ fontSize: 10, color: T.textDim, transition: "transform 0.2s", display: "inline-block", transform: showAdminLogin ? "rotate(180deg)" : "none" }}>▼</span>
          </button>

          {showAdminLogin && (
            <div style={{ marginTop: 12, display: "flex", flexDirection: "column", gap: 10 }}>
              <Input
                value={adminUser} onChange={e => { setAdminUser(e.target.value); setErr(""); }}
                placeholder="Username" style={{ width: "100%" }}
                onKeyDown={e => e.key === "Enter" && handleAdminLogin()}
              />
              <Input
                type="password" value={adminPass} onChange={e => { setAdminPass(e.target.value); setErr(""); }}
                placeholder="Password" style={{ width: "100%" }}
                onKeyDown={e => e.key === "Enter" && handleAdminLogin()}
              />
              <button onClick={handleAdminLogin} style={{
                width: "100%", padding: "11px", background: `linear-gradient(135deg, ${T.brand}, #5e6bf7)`,
                border: "none", borderRadius: 8, color: "#fff", fontSize: 13, fontWeight: 800,
                cursor: "pointer", fontFamily: F.body,
              }}>Sign In</button>
            </div>
          )}

          <p style={{ margin: "18px 0 0", fontSize: 10, color: T.textDim, textAlign: "center", lineHeight: 1.6 }}>
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
const BLANK_FORM = { name: "", email: "", role: "member", title: "", deptId: "", teamId: "", teamIds: [] };

function UserMgmtPage({ users, depts, dispatch, currentUserId }) {
  const [showAdd, setShowAdd] = useState(false);
  const [form, setForm] = useState(BLANK_FORM);
  const [formErr, setFormErr] = useState("");
  const [editId, setEditId] = useState(null);
  const [editForm, setEditForm] = useState(BLANK_FORM);

  const teamsForDept = (deptId) => depts.find(d => d.id === deptId)?.teams || [];

  function handleAdd() {
    if (!form.name.trim() || !form.email.trim()) { setFormErr("Name and email are required."); return; }
    if (users.some(u => u.email.toLowerCase() === form.email.trim().toLowerCase())) { setFormErr("A user with this email already exists."); return; }
    const id = `usr_${Date.now().toString(36)}`;
    const newUser = {
      id, name: form.name.trim(), email: form.email.trim().toLowerCase(),
      role: form.role, av: makeAv(form.name), title: form.title.trim() || form.role,
      ...(form.deptId && { deptId: form.deptId }),
      ...(form.role === "member" && form.teamId && { teamId: form.teamId }),
      ...(form.role === "manager" && form.teamIds?.length && { teamIds: form.teamIds }),
    };
    dispatch({ type: "ADD_USER", user: newUser });
    setForm(BLANK_FORM); setFormErr(""); setShowAdd(false);
  }

  function startEdit(u) {
    setEditId(u.id);
    setEditForm({ name: u.name, email: u.email, role: u.role, title: u.title || "", deptId: u.deptId || "", teamId: u.teamId || "", teamIds: u.teamIds || [] });
  }

  function saveEdit() {
    dispatch({ type: "UPDATE_USER", userId: editId, updates: {
      name: editForm.name.trim(), email: editForm.email.trim().toLowerCase(),
      role: editForm.role, av: makeAv(editForm.name), title: editForm.title.trim() || editForm.role,
      deptId: editForm.deptId || undefined,
      teamId: editForm.role === "member" ? (editForm.teamId || undefined) : undefined,
      teamIds: editForm.role === "manager" ? (editForm.teamIds?.length ? editForm.teamIds : undefined) : undefined,
    }});
    setEditId(null);
  }

  const roleColor = { admin: T.brand, manager: T.orange, member: T.ok };

  const roleCounts = users.reduce((a, u) => { a[u.role] = (a[u.role] || 0) + 1; return a; }, {});

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
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 16, color: T.text }}>New User</div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, marginBottom: 12 }}>
            <div>
              <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Full Name *</div>
              <Input value={form.name} onChange={e => setForm(p => ({ ...p, name: e.target.value }))} placeholder="e.g. Jane Smith" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Email Address *</div>
              <Input value={form.email} onChange={e => setForm(p => ({ ...p, email: e.target.value }))} placeholder="jane@niet.edu.au" style={{ width: "100%" }} />
            </div>
            <div>
              <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Role *</div>
              <Select value={form.role} onChange={e => setForm(p => ({ ...p, role: e.target.value, deptId: "", teamId: "", teamIds: [] }))} style={{ width: "100%" }}>
                <option value="admin">Admin</option>
                <option value="manager">Manager</option>
                <option value="member">Member</option>
              </Select>
            </div>
            <div>
              <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Job Title</div>
              <Input value={form.title} onChange={e => setForm(p => ({ ...p, title: e.target.value }))} placeholder="e.g. Admissions Officer" style={{ width: "100%" }} />
            </div>
            {form.role !== "admin" && (
              <div>
                <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Department</div>
                <Select value={form.deptId} onChange={e => setForm(p => ({ ...p, deptId: e.target.value, teamId: "", teamIds: [] }))} style={{ width: "100%" }}>
                  <option value="">— Select department —</option>
                  {depts.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
                </Select>
              </div>
            )}
          </div>
          {formErr && <div style={{ padding: "8px 12px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 6, fontSize: 11, color: T.bad, marginBottom: 12 }}>{formErr}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <Btn small onClick={() => { setShowAdd(false); setForm(BLANK_FORM); setFormErr(""); }}>Cancel</Btn>
            <Btn primary small onClick={handleAdd}>Create User</Btn>
          </div>
        </Card>
      )}

      {/* User table */}
      <Card style={{ overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "32px 1fr 180px 80px 160px 90px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
          <span></span><span>Name / Email</span><span>Title</span><span>Role</span><span>Department</span><span style={{ textAlign: "right" }}>Actions</span>
        </div>
        {users.map((u, i) => {
          const dept = depts.find(d => d.id === u.deptId);
          const team = dept?.teams.find(t => t.id === u.teamId || u.teamIds?.includes(t.id));
          const isSystem = u.id === "sysadmin";
          const isSelf = u.id === currentUserId;

          if (editId === u.id) {
            const editTeams = teamsForDept(editForm.deptId);
            return (
              <div key={u.id} style={{ background: T.brandDim, borderBottom: `1px solid ${T.border}`, padding: "12px 18px" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                  <Input value={editForm.name}   onChange={e => setEditForm(p => ({ ...p, name:  e.target.value }))} placeholder="Name"  style={{ fontSize: 11, padding: "7px 10px" }} />
                  <Input value={editForm.email}  onChange={e => setEditForm(p => ({ ...p, email: e.target.value }))} placeholder="Email" style={{ fontSize: 11, padding: "7px 10px" }} />
                  <Input value={editForm.title}  onChange={e => setEditForm(p => ({ ...p, title: e.target.value }))} placeholder="Title" style={{ fontSize: 11, padding: "7px 10px" }} />
                  <Select value={editForm.role} onChange={e => setEditForm(p => ({ ...p, role: e.target.value, deptId: "", teamId: "", teamIds: [] }))} style={{ fontSize: 11, padding: "7px 10px" }}>
                    <option value="admin">Admin</option>
                    <option value="manager">Manager</option>
                    <option value="member">Member</option>
                  </Select>
                  {editForm.role !== "admin" && (
                    <Select value={editForm.deptId} onChange={e => setEditForm(p => ({ ...p, deptId: e.target.value, teamId: "", teamIds: [] }))} style={{ fontSize: 11, padding: "7px 10px" }}>
                      <option value="">— Department —</option>
                      {depts.map(d => <option key={d.id} value={d.id}>{d.name}</option>)}
                    </Select>
                  )}
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <Btn small onClick={() => setEditId(null)}>Cancel</Btn>
                  <Btn primary small onClick={saveEdit}>Save</Btn>
                </div>
              </div>
            );
          }

          return (
            <div key={u.id} style={{ display: "grid", gridTemplateColumns: "32px 1fr 180px 80px 160px 90px", padding: "10px 18px", gap: 10, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
              <Avatar letters={u.av} size={26} />
              <div>
                <div style={{ fontWeight: 600, color: T.text }}>{u.name}{isSelf && <span style={{ fontSize: 9, color: T.brand, marginLeft: 6 }}>you</span>}</div>
                <div style={{ fontSize: 10, color: T.textMuted }}>{isSystem ? "System login only" : u.email}</div>
              </div>
              <span style={{ fontSize: 11, color: T.textSoft }}>{u.title}</span>
              <RoleTag role={u.role} />
              <span style={{ fontSize: 11, color: T.textMuted }}>{dept?.name || "—"}</span>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                {!isSystem && (
                  <button onClick={() => startEdit(u)} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                )}
                {!isSystem && !isSelf && (
                  <button onClick={() => { if (window.confirm(`Delete user "${u.name}"? This cannot be undone.`)) dispatch({ type: "REMOVE_USER", userId: u.id }); }} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>✕</button>
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

function DeptMgmtPage({ depts, users, memberData, dispatch }) {
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
  const [editTeamForm, setEditTeamForm] = useState({ name: "", lead: "", obj: "" });
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

  const labelStyle = { fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 };

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
          <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 16 }}>New Department</div>
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
          {addErr && <div style={{ padding: "8px 12px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 6, fontSize: 11, color: T.bad, marginBottom: 12 }}>{addErr}</div>}
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <Btn small onClick={() => { setShowAdd(false); setAddForm(BLANK_DEPT); setAddErr(""); }}>Cancel</Btn>
            <Btn primary small onClick={handleAdd}>Create Department</Btn>
          </div>
        </Card>
      )}

      <Card style={{ overflow: "hidden" }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 200px 80px 80px 110px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
          <span>Department / Description</span><span>Head · College</span><span>Teams</span><span>Completion</span><span style={{ textAlign: "right" }}>Actions</span>
        </div>
        {depts.length === 0 && <EmptyState text="No departments yet. Add one above." />}
        {depts.map((d, i) => {
          const r = calcRate(d.krs); const s = getStatus(r);
          if (editId === d.id) {
            return (
              <div key={d.id} style={{ background: T.brandDim, borderBottom: `1px solid ${T.border}`, padding: "14px 18px" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                  <div><div style={labelStyle}>Name *</div><Input value={editForm.name} onChange={e => setEditForm(p => ({ ...p, name: e.target.value }))} style={{ fontSize: 11, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={labelStyle}>Description / Objective</div><Input value={editForm.obj} onChange={e => setEditForm(p => ({ ...p, obj: e.target.value }))} style={{ fontSize: 11, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={labelStyle}>Department Head</div><Input value={editForm.head} onChange={e => setEditForm(p => ({ ...p, head: e.target.value }))} style={{ fontSize: 11, padding: "7px 10px", width: "100%" }} /></div>
                  <div><div style={labelStyle}>College / Group</div><Input value={editForm.college} onChange={e => setEditForm(p => ({ ...p, college: e.target.value }))} style={{ fontSize: 11, padding: "7px 10px", width: "100%" }} /></div>
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
              <div style={{ display: "grid", gridTemplateColumns: "1fr 200px 80px 80px 110px", padding: "12px 18px", gap: 10, alignItems: "center", background: isSelected ? T.brandDim : i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                <div style={{ cursor: "pointer" }} onClick={() => { setSelDept(isSelected ? null : d.id); }}>
                  <div style={{ fontWeight: 700, color: isSelected ? T.brand : T.text }}>{d.name}</div>
                  {d.obj && <div style={{ fontSize: 10, color: T.textMuted, marginTop: 2 }}>{d.obj}</div>}
                </div>
                <div>
                  <div style={{ fontSize: 11, color: T.textSoft }}>{d.head || "—"}</div>
                  <div style={{ fontSize: 10, color: T.textMuted }}>{d.college || ""}</div>
                </div>
                <span style={{ fontSize: 11, color: T.textSoft }}>{d.teams.length} team{d.teams.length !== 1 ? "s" : ""}</span>
                <span style={{ fontFamily: F.mono, fontWeight: 700, fontSize: 12, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</span>
                <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                  <button onClick={() => startEdit(d)} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                  {confirmDel === d.id ? (<>
                    <button onClick={() => { dispatch({ type: "REMOVE_DEPT", deptId: d.id }); setConfirmDel(null); if (selDept === d.id) setSelDept(null); }} style={{ background: T.bad, border: "none", borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: "#fff", fontSize: 10, fontWeight: 700, fontFamily: F.body }}>Confirm</button>
                    <button onClick={() => setConfirmDel(null)} style={{ background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.textSoft, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>✕</button>
                  </>) : (
                    <button onClick={() => setConfirmDel(d.id)} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>✕</button>
                  )}
                </div>
              </div>
              {isSelected && (() => {
                const r2 = calcRate(d.krs); const s2 = getStatus(r2);
                const deptMembers = users
                  .filter(u => u.role === "member" && u.deptId === d.id)
                  .map(u => { const kd = memberData[u.id] || { krs: [] }; const mr = calcRate(kd.krs); return { ...u, rate: mr, status: getStatus(mr) }; })
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
                        <div style={{ padding: "9px 16px", borderBottom: `1px solid ${T.border}`, fontSize: 10, fontWeight: 700, color: T.textMuted, letterSpacing: "0.05em" }}>DEPARTMENT KEY RESULTS</div>
                        {d.krs.map((kr, ki) => { const cr = Math.min((kr.actual / kr.target) * 100, 100); const cs = getStatus(cr);
                          return (<div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 70px 70px 55px 130px 65px", padding: "9px 16px", gap: 8, alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                            <span style={{ fontFamily: F.mono, fontSize: 10, color: T.textDim }}>{kr.id}</span><span>{kr.label}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.target)}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700 }}>{fmt(kr.actual)}</span>
                            <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span>
                            <Bar value={cr} status={cs} h={5} />
                            <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={cs} small /></div>
                          </div>);
                        })}
                      </Card>
                    )}

                    <Card style={{ overflow: "hidden", marginBottom: 14 }}>
                      <div style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 120px 55px 140px 70px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                        <span>#</span><span></span><span>Name</span><span>Title</span><span style={{ textAlign: "right" }}>Rate</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                      </div>
                      {deptMembers.length === 0
                        ? <div style={{ padding: "14px 18px", fontSize: 12, color: T.textMuted }}>No members assigned to this department.</div>
                        : deptMembers.map((m, mi) => (
                          <div key={m.id} style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 120px 55px 140px 70px", padding: "9px 18px", gap: 10, alignItems: "center", background: mi % 2 ? T.raised : "transparent", borderBottom: mi < deptMembers.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 12 }}>
                            <span style={{ fontFamily: F.mono, fontWeight: 800, color: mi === 0 ? T.ok : mi === deptMembers.length - 1 && deptMembers.length > 2 ? T.bad : T.textMuted }}>#{mi + 1}</span>
                            <Avatar letters={m.av} size={26} />
                            <div><span style={{ fontWeight: 600 }}>{m.name}</span></div>
                            <span style={{ fontSize: 10, color: T.textMuted }}>{m.title || "—"}</span>
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
                        <div style={{ fontSize: 12, fontWeight: 700, marginBottom: 10 }}>New Team</div>
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
                      <div style={{ fontSize: 11, color: T.textMuted, padding: "8px 0 12px" }}>No teams yet. Add one above.</div>
                    )}

                    {d.teams.map(t => { const tr = calcRate(t.krs); const ts = getStatus(tr);
                      const isEditingTeam = editTeam?.deptId === d.id && editTeam?.teamId === t.id;
                      return (
                        <Card key={t.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                          {isEditingTeam ? (
                            <div style={{ padding: 14 }}>
                              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                                <div><div style={labelStyle}>Team Name *</div><Input value={editTeamForm.name} onChange={e => setEditTeamForm(p => ({ ...p, name: e.target.value }))} style={{ width: "100%" }} /></div>
                                <div><div style={labelStyle}>Team Lead</div><Input value={editTeamForm.lead} onChange={e => setEditTeamForm(p => ({ ...p, lead: e.target.value }))} style={{ width: "100%" }} /></div>
                                <div><div style={labelStyle}>Objective</div><Input value={editTeamForm.obj} onChange={e => setEditTeamForm(p => ({ ...p, obj: e.target.value }))} style={{ width: "100%" }} /></div>
                              </div>
                              <div style={{ display: "flex", gap: 8 }}>
                                <Btn small onClick={() => setEditTeam(null)}>Cancel</Btn>
                                <Btn primary small disabled={!editTeamForm.name.trim()} onClick={() => {
                                  dispatch({ type: "UPDATE_TEAM", deptId: d.id, teamId: t.id, updates: { name: editTeamForm.name.trim(), lead: editTeamForm.lead.trim(), obj: editTeamForm.obj.trim() } });
                                  setEditTeam(null);
                                }}>Save</Btn>
                              </div>
                            </div>
                          ) : (
                            <div style={{ padding: "10px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                              <div>
                                <span style={{ fontSize: 13, fontWeight: 700 }}>{t.name}</span>
                                {t.lead && <span style={{ fontSize: 10, color: T.textMuted, marginLeft: 10 }}>Lead: {t.lead}</span>}
                                {t.obj && <div style={{ fontSize: 10, color: T.textMuted, marginTop: 2 }}>{t.obj}</div>}
                              </div>
                              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                                <span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[ts].color }}>{tr.toFixed(1)}%</span>
                                <Tag type={ts} small />
                                <button onClick={() => { setEditTeam({ deptId: d.id, teamId: t.id }); setEditTeamForm({ name: t.name, lead: t.lead || "", obj: t.obj || "" }); setShowAddTeam(null); }} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                                {confirmDelTeam?.deptId === d.id && confirmDelTeam?.teamId === t.id ? (<>
                                  <button onClick={() => { dispatch({ type: "REMOVE_TEAM", deptId: d.id, teamId: t.id }); setConfirmDelTeam(null); }} style={{ background: T.bad, border: "none", borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: "#fff", fontSize: 10, fontWeight: 700, fontFamily: F.body }}>Confirm</button>
                                  <button onClick={() => setConfirmDelTeam(null)} style={{ background: T.raised, border: `1px solid ${T.border}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.textSoft, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>✕</button>
                                </>) : (
                                  <button onClick={() => setConfirmDelTeam({ deptId: d.id, teamId: t.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>✕</button>
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
  const [page, setPage] = useState("overview");
  const [selDept, setSelDept] = useState(null);
  const [selTeam, setSelTeam] = useState(null);
  const [newKr, setNewKr] = useState({ label: "", target: "", unit: "", dataSource: "" });
  const [addTarget, setAddTarget] = useState(null);
  const [showGenReport, setShowGenReport] = useState(false);
  const [genPeriod, setGenPeriod] = useState({ label: "", from: "", to: "" });
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ name: "", progress: 0, status: "active", log: "", due: "" });
  const [subFilter, setSubFilter] = useState("all");

  const { depts, memberData, mgrSprints, monthlyReports, projects, weeklySubs, users } = state;
  const navItems = [
    { id: "overview",     icon: "◎", label: "Company Overview"    },
    { id: "departments",  icon: "⬛", label: "Departments"         },
    { id: "setup",        icon: "⚙", label: "OKR / KPI Setup"     },
    { id: "submissions",  icon: "✉", label: "Weekly Submissions"  },
    { id: "reports",      icon: "⊞", label: "Monthly Reports"     },
    { id: "projects",     icon: "⚡", label: "Projects"            },
    { id: "leaderboard",  icon: "▲", label: "Leaderboard"         },
    { id: "users",        icon: "⊹", label: "User Management"     },
  ];

  const deptRanks = depts.map(d => ({ ...d, rate: calcRate(d.krs), status: getStatus(calcRate(d.krs)) })).sort((a, b) => b.rate - a.rate);
  const compRate = deptRanks.length ? deptRanks.reduce((a, d) => a + d.rate, 0) / deptRanks.length : 0;
  const allMembers = users
    .filter(u => u.role === "member")
    .map(u => {
      const kd = memberData[u.id] || { krs: [] };
      const r = calcRate(kd.krs);
      const deptName = depts.find(d => d.id === u.deptId)?.name || "—";
      return { ...u, deptName, rate: r, status: getStatus(r) };
    })
    .sort((a, b) => b.rate - a.rate);

  function addKr(deptId, teamId) {
    if (!newKr.label || !newKr.target) return;
    const kr = { id: `N${Date.now().toString(36).slice(-4).toUpperCase()}`, label: newKr.label, target: Number(newKr.target), actual: 0, unit: newKr.unit.trim(), dataSource: newKr.dataSource.trim() };
    dispatch({ type: "ADD_KR", deptId, teamId, kr });
    setNewKr({ label: "", target: "", unit: "", dataSource: "" }); setAddTarget(null);
  }

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "users" && <UserMgmtPage users={users} depts={depts} dispatch={dispatch} currentUserId={user.id} />}

        {page === "overview" && (<>
          <Header title="Company Overview" sub="FY26 Q1 · All colleges · All departments"
            right={<div style={{ display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 10, color: T.textMuted, fontFamily: F.mono }}>Time: {TP}%</span><Tag type={getStatus(compRate)} /></div>} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Company Completion" value={`${compRate.toFixed(1)}%`} status={getStatus(compRate)} sub={`Target pace: ${TP}%`} />
              <Metric label="Departments" value={depts.length} />
              <Metric label="Teams" value={depts.reduce((a, d) => a + d.teams.length, 0)} />
              <Metric label="Staff Tracked" value={allMembers.length} />
            </div>
            <div>
              <SectionLabel>Department Rankings</SectionLabel>
              {deptRanks.map((d, i) => (
                <Card key={d.id} onClick={() => { setSelDept(d.id); setPage("departments"); }} style={{ padding: "16px 20px", marginBottom: 8 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "36px 1fr 60px 180px 80px", alignItems: "center", gap: 14 }}>
                    <span style={{ fontSize: 18, fontWeight: 900, fontFamily: F.mono, color: i === 0 ? T.ok : i === deptRanks.length - 1 ? T.bad : T.textMuted }}>#{i + 1}</span>
                    <div><div style={{ fontSize: 14, fontWeight: 700 }}>{d.name}</div><div style={{ fontSize: 10, color: T.textMuted }}>{d.college} · {d.head} · {d.teams.length} teams</div></div>
                    <span style={{ textAlign: "right", fontSize: 16, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[d.status].color }}>{d.rate.toFixed(1)}%</span>
                    <Bar value={d.rate} status={d.status} h={7} />
                    <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={d.status} /></div>
                  </div>
                </Card>
              ))}
            </div>
          </Pane>
        </>)}

        {page === "departments" && <DeptMgmtPage depts={depts} users={users} memberData={memberData} dispatch={dispatch} />}

        {page === "setup" && (<>
          <Header title="OKR / KPI Setup" sub="Add, edit, or remove key results for each department and team" />
          <Pane>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>{depts.map(d => <Btn key={d.id} primary={selDept === d.id} small onClick={() => { setSelDept(d.id); setSelTeam(null); setAddTarget(null); }}>{d.name}</Btn>)}</div>
            {selDept && (() => {
              const dept = depts.find(d => d.id === selDept); if (!dept) return null;
              const COL = "50px 160px 90px 80px 90px 500px 50px";
              const renderEditor = (krs, deptId, teamId) => (
                <Card style={{ overflow: "hidden" }}>
                  <div style={{ display: "grid", gridTemplateColumns: COL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                    <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Unit</span><span>Data Source</span><span></span>
                  </div>
                  {krs.map((kr, i) => (
                    <div key={kr.id} style={{ display: "grid", gridTemplateColumns: COL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                      <span style={{ fontFamily: F.mono, fontSize: 10, color: T.textDim }}>{kr.id}</span>
                      <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                      <Input value={kr.target} onChange={e => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "target", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                      <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.actual)}</span>
                      <Input value={kr.unit || ""} onChange={e => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "unit", value: e.target.value })} placeholder="e.g. %, students" style={{ padding: "5px 8px", fontSize: 11 }} />
                      <Input value={kr.dataSource || ""} onChange={e => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "dataSource", value: e.target.value })} placeholder="e.g. CRM, Manual" style={{ padding: "5px 8px", fontSize: 11 }} />
                      <button onClick={() => dispatch({ type: "REMOVE_KR", deptId, teamId, krId: kr.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 8px", cursor: "pointer", color: T.bad, fontSize: 10, fontWeight: 700 }}>✕</button>
                    </div>
                  ))}
                  {addTarget === (teamId || `dept-${deptId}`) ? (
                    <div style={{ display: "grid", gridTemplateColumns: COL, padding: "9px 16px", gap: 8, alignItems: "center", background: T.brandDim }}>
                      <span style={{ fontSize: 10, color: T.brand }}>NEW</span>
                      <Input value={newKr.label} onChange={e => setNewKr(p => ({ ...p, label: e.target.value }))} placeholder="KR description" style={{ padding: "5px 8px", fontSize: 12 }} />
                      <Input value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Target" style={{ textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                      <span />
                      <Input value={newKr.unit} onChange={e => setNewKr(p => ({ ...p, unit: e.target.value }))} placeholder="Unit" style={{ padding: "5px 8px", fontSize: 11 }} />
                      <Input value={newKr.dataSource} onChange={e => setNewKr(p => ({ ...p, dataSource: e.target.value }))} placeholder="Data source" style={{ padding: "5px 8px", fontSize: 11 }} />
                      <button onClick={() => addKr(deptId, teamId)} style={{ background: T.brand, border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", color: "#fff", fontSize: 10, fontWeight: 700 }}>✓</button>
                    </div>
                  ) : (
                    <div style={{ padding: "10px 16px" }}>
                      <button onClick={() => { setAddTarget(teamId || `dept-${deptId}`); setNewKr({ label: "", target: "", unit: "", dataSource: "" }); }} style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: 6, padding: "8px 14px", cursor: "pointer", color: T.brand, fontSize: 11, fontWeight: 600, width: "100%", fontFamily: F.body }}>+ Add Key Result</button>
                    </div>
                  )}
                </Card>
              );
              return (<>
                <div><div style={{ fontSize: 14, fontWeight: 800, marginBottom: 4 }}>{dept.name} — Department KRs</div><div style={{ fontSize: 11, color: T.textMuted, marginBottom: 12 }}>Objective: {dept.obj}</div>{renderEditor(dept.krs, dept.id, null)}</div>
                {dept.teams.length > 0 && (<div>
                  <SectionLabel>Team KRs</SectionLabel>
                  <div style={{ display: "flex", gap: 6, marginBottom: 12 }}>{dept.teams.map(t => <Btn key={t.id} primary={selTeam === t.id} small onClick={() => { setSelTeam(t.id); setAddTarget(null); }}>{t.name}</Btn>)}</div>
                  {selTeam && (() => { const team = dept.teams.find(t => t.id === selTeam); return team ? (<><div style={{ fontSize: 11, color: T.textMuted, marginBottom: 10 }}>Objective: {team.obj} · Lead: {team.lead}</div>{renderEditor(team.krs, dept.id, team.id)}</>) : null; })()}
                </div>)}
              </>);
            })()}
          </Pane>
        </>)}

        {page === "submissions" && (<>
          <Header title="Weekly Staff Submissions" sub="All weekly work outcome submissions — managers approve in their portal" />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Total"    value={weeklySubs.length} />
              <Metric label="Pending"  value={weeklySubs.filter(s => s.approval === "pending").length}  status="yellow" />
              <Metric label="Approved" value={weeklySubs.filter(s => s.approval === "approved").length} status="green"  />
              <Metric label="Rejected" value={weeklySubs.filter(s => s.approval === "rejected").length} status="red"    />
            </div>
            <div style={{ display: "flex", gap: 6 }}>
              {["all", "pending", "approved", "rejected"].map(f => (
                <Btn key={f} small primary={subFilter === f} onClick={() => setSubFilter(f)}>
                  {f.charAt(0).toUpperCase() + f.slice(1)}
                </Btn>
              ))}
            </div>
            {weeklySubs.length === 0 && <EmptyState text="No weekly submissions yet." />}
            {(() => {
              const filtered = weeklySubs
                .filter(s => subFilter === "all" || s.approval === subFilter)
                .sort((a, b) => b.date.localeCompare(a.date));
              if (filtered.length === 0) return <EmptyState text={`No ${subFilter} submissions.`} />;
              return filtered.map(s => {
                const mem = users.find(u => u.id === s.memberId);
                const dept = depts.find(d => d.id === mem?.deptId);
                const mgr = users.find(u => u.role === "manager" && u.deptId === mem?.deptId);
                return (
                  <Card key={s.id} style={{ padding: "16px 20px", borderLeft: s.approval === "pending" ? `3px solid ${T.warn}` : s.approval === "approved" ? `3px solid ${T.ok}` : `3px solid ${T.bad}` }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <Avatar letters={mem?.av || "?"} size={30} />
                        <div>
                          <div style={{ fontSize: 13, fontWeight: 700 }}>{mem?.name || "Unknown"}</div>
                          <div style={{ fontSize: 10, color: T.textMuted }}>{mem?.title || "—"}{dept ? ` · ${dept.name}` : ""}{mgr ? ` · Manager: ${mgr.name}` : ""}</div>
                        </div>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                        <div style={{ textAlign: "right" }}>
                          <div style={{ fontSize: 12, fontWeight: 600 }}>{s.week}</div>
                          <div style={{ fontSize: 10, color: T.textMuted }}>{s.date}</div>
                        </div>
                        <Tag type={s.approval} label={APPROVAL[s.approval]?.label || s.approval} />
                      </div>
                    </div>
                    <p style={{ margin: 0, fontSize: 12, color: T.textSoft, lineHeight: 1.6, padding: "10px 14px", background: T.raised, borderRadius: 7 }}>{s.items}</p>
                    {s.mgrNote && <div style={{ marginTop: 8, fontSize: 11, color: T.textMuted, fontStyle: "italic" }}>Manager note: {s.mgrNote}</div>}
                  </Card>
                );
              });
            })()}
          </Pane>
        </>)}

        {page === "reports" && (<>
          <Header title="Monthly KPI Reports" sub="Published reports visible to ALL teams across the company"
            right={<div style={{ display: "flex", gap: 8 }}>
              <Btn onClick={() => { setShowGenReport(v => !v); setGenPeriod({ label: "", from: "", to: "" }); }}>{showGenReport ? "Cancel" : "Generate for Period"}</Btn>
              <Btn primary onClick={() => {
                const report = { id: `mr${Date.now()}`, month: currentMonth(), publishedDate: new Date().toISOString().slice(0, 10), publishedBy: user.id,
                  data: { companyRate: Math.round(compRate * 10) / 10, deptRanks: deptRanks.map(d => ({ name: d.name, rate: Math.round(d.rate * 10) / 10, status: d.status })),
                    topPerformers: allMembers.slice(0, 3).map(m => `${m.name} — ${m.rate.toFixed(1)}%`),
                    redFlags: allMembers.filter(m => m.status === "red").map(m => `${m.name} — ${m.rate.toFixed(1)}% (action required)`),
                  },
                };
                dispatch({ type: "PUBLISH_REPORT", report });
              }}>Publish {currentMonth()} Report</Btn>
            </div>} />
          <Pane>
            {showGenReport && (
              <Card style={{ padding: 20, borderLeft: `3px solid ${T.brand}` }}>
                <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14 }}>Generate Report for Specific Period</div>
                <div style={{ display: "flex", gap: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
                  <div>
                    <div style={{ fontSize: 10, color: T.textMuted, marginBottom: 4 }}>Period Label</div>
                    <Input value={genPeriod.label} onChange={e => setGenPeriod(p => ({ ...p, label: e.target.value }))} placeholder="e.g. Q1 FY2026, March 2026" style={{ width: 200 }} />
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: T.textMuted, marginBottom: 4 }}>From</div>
                    <Input type="date" value={genPeriod.from} onChange={e => setGenPeriod(p => ({ ...p, from: e.target.value }))} style={{ width: 150 }} />
                  </div>
                  <div>
                    <div style={{ fontSize: 10, color: T.textMuted, marginBottom: 4 }}>To</div>
                    <Input type="date" value={genPeriod.to} onChange={e => setGenPeriod(p => ({ ...p, to: e.target.value }))} style={{ width: 150 }} />
                  </div>
                  <Btn primary disabled={!genPeriod.label} onClick={() => {
                    const label = genPeriod.label.trim();
                    const range = genPeriod.from && genPeriod.to ? ` (${genPeriod.from} → ${genPeriod.to})` : "";
                    const report = {
                      id: `mr${Date.now()}`,
                      month: label,
                      publishedDate: new Date().toISOString().slice(0, 10),
                      publishedBy: user.id,
                      periodFrom: genPeriod.from || null,
                      periodTo: genPeriod.to || null,
                      data: {
                        companyRate: Math.round(compRate * 10) / 10,
                        deptRanks: deptRanks.map(d => ({ name: d.name, rate: Math.round(d.rate * 10) / 10, status: d.status })),
                        topPerformers: allMembers.slice(0, 3).map(m => `${m.name} — ${m.rate.toFixed(1)}%`),
                        redFlags: allMembers.filter(m => m.status === "red").map(m => `${m.name} — ${m.rate.toFixed(1)}% (action required)`),
                      },
                    };
                    dispatch({ type: "PUBLISH_REPORT", report });
                    setShowGenReport(false);
                    setGenPeriod({ label: "", from: "", to: "" });
                  }}>Generate & Publish</Btn>
                </div>
              </Card>
            )}
            {state.monthlyReports.length === 0 && <EmptyState text="No monthly reports published yet." />}
            {state.monthlyReports.map(r => (
              <Card key={r.id} style={{ overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <div style={{ fontSize: 16, fontWeight: 800 }}>{r.month}</div>
                    <div style={{ fontSize: 10, color: T.textMuted }}>
                      Published: {r.publishedDate}{r.periodFrom && r.periodTo ? ` · Period: ${r.periodFrom} → ${r.periodTo}` : ""} · Visible to all teams
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
                    <button onClick={() => { if (window.confirm(`Delete report "${r.month}"? This cannot be undone.`)) dispatch({ type: "REMOVE_REPORT", reportId: r.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete report">✕</button>
                  </div>
                </div>
                <div style={{ padding: "16px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
                  <div>
                    <SectionLabel>Department Rankings</SectionLabel>
                    {r.data.deptRanks.map((d, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "6px 0", fontSize: 12 }}>
                        <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                        <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                        <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{d.rate}%</span>
                        <Tag type={d.status} small />
                      </div>
                    ))}
                  </div>
                  <div>
                    <SectionLabel>Top Performers</SectionLabel>
                    {r.data.topPerformers.map((p, i) => <div key={i} style={{ padding: "5px 0", fontSize: 12, color: T.ok, display: "flex", alignItems: "center", gap: 6 }}><span>★</span> {p}</div>)}
                    {r.data.redFlags.length > 0 && (<><SectionLabel>Action Required</SectionLabel>{r.data.redFlags.map((f, i) => <div key={i} style={{ padding: "5px 0", fontSize: 12, color: T.bad, display: "flex", alignItems: "center", gap: 6 }}><span>⚠</span> {f}</div>)}</>)}
                  </div>
                </div>
              </Card>
            ))}
          </Pane>
        </>)}

        {page === "projects" && (<>
          <Header title="Manager Projects" sub="All projects submitted by department managers" />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Total Projects" value={projects.length} />
              <Metric label="Active"         value={projects.filter(p => p.status === "active").length}   status="yellow" />
              <Metric label="Completed"      value={projects.filter(p => p.status !== "active").length}   status="green"  />
            </div>
            {projects.length === 0 && <EmptyState text="No projects submitted by managers yet." />}
            {projects.length > 0 && (() => {
              const managers = users.filter(u => u.role === "manager");
              return managers.map(mgr => {
                const mgrProjects = projects.filter(p => p.mgrId === mgr.id);
                if (mgrProjects.length === 0) return null;
                const dept = depts.find(d => d.id === mgr.deptId);
                return (
                  <div key={mgr.id} style={{ marginBottom: 20 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                      <Avatar letters={mgr.av} size={28} />
                      <div>
                        <div style={{ fontSize: 13, fontWeight: 700 }}>{mgr.name}</div>
                        <div style={{ fontSize: 10, color: T.textMuted }}>{mgr.title}{dept ? ` · ${dept.name}` : ""}</div>
                      </div>
                      <span style={{ marginLeft: "auto", fontSize: 10, color: T.textMuted }}>{mgrProjects.length} project{mgrProjects.length !== 1 ? "s" : ""}</span>
                    </div>
                    {mgrProjects.map(p => {
                      const isActive = p.status === "active";
                      const ps = p.progress >= 70 ? "green" : p.progress >= 35 ? "yellow" : "red";
                      const isEditing = editProjId === p.id;
                      return (
                        <Card key={p.id} style={{ overflow: "hidden", marginBottom: 8 }}>
                          <div style={{ padding: "12px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                            <div>
                              <div style={{ fontSize: 13, fontWeight: 700 }}>{p.name}</div>
                              <div style={{ fontSize: 10, color: T.textMuted }}>Due: {p.due}{p.updatedDate ? ` · Last updated: ${p.updatedDate}` : ""}</div>
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
                                <span style={{ fontFamily: F.mono, fontSize: 12, fontWeight: 700, color: STATUS_THEME[ps].color, whiteSpace: "nowrap" }}>{p.progress}%</span>
                              </div>
                              {p.log && <p style={{ margin: 0, fontSize: 11, color: T.textSoft, lineHeight: 1.6, padding: "8px 12px", background: T.raised, borderRadius: 6 }}>{p.log}</p>}
                            </div>
                          )}
                          {isEditing && (
                            <div style={{ padding: "14px 18px" }}>
                              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                                <div>
                                  <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Name</div>
                                  <Input value={editProjForm.name} onChange={e => setEditProjForm(f => ({ ...f, name: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 12 }} />
                                </div>
                                <div>
                                  <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Completion %</div>
                                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                                    <input type="range" min={0} max={100} value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Number(e.target.value) }))} style={{ flex: 1 }} />
                                    <Input value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 50, textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                                  </div>
                                </div>
                                <div>
                                  <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                                  <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 12, fontFamily: F.body }}>
                                    <option value="active">Active</option>
                                    <option value="completed">Completed</option>
                                  </select>
                                </div>
                                <div>
                                  <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                                  <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 12 }} />
                                </div>
                              </div>
                              <div style={{ marginBottom: 12 }}>
                                <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Log / Notes</div>
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
          <Header title="Company Leaderboard" sub="All staff ranked by KPI completion · FY26 Q1" />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Total"    value={allMembers.length} />
              <Metric label="On Track" value={allMembers.filter(m => m.status === "green").length}  status="green"  />
              <Metric label="At Risk"  value={allMembers.filter(m => m.status === "yellow").length} status="yellow" />
              <Metric label="Behind"   value={allMembers.filter(m => m.status === "red").length}    status="red"    />
            </div>
            <Card style={{ overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "36px 32px 1fr 140px 55px 160px 70px", padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                <span>#</span><span></span><span>Name</span><span>Department</span><span style={{ textAlign: "right" }}>Rate</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
              </div>
              {allMembers.map((m, i) => (
                <div key={m.id} style={{ display: "grid", gridTemplateColumns: "36px 32px 1fr 140px 55px 160px 70px", padding: "10px 16px", gap: 8, alignItems: "center", background: i === 0 ? T.okDim : m.status === "red" ? T.badDim : i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, borderLeft: i === 0 ? `3px solid ${T.ok}` : m.status === "red" ? `3px solid ${T.bad}` : "3px solid transparent", fontSize: 12 }}>
                  <span style={{ fontFamily: F.mono, fontWeight: 900, color: i === 0 ? T.ok : m.status === "red" ? T.bad : T.textMuted }}>#{i + 1}</span>
                  <Avatar letters={m.av} size={24} />
                  <div><span style={{ fontWeight: 600 }}>{m.name}</span><span style={{ color: T.textDim, marginLeft: 6, fontSize: 10 }}>{m.title}</span></div>
                  <span style={{ fontSize: 10, color: T.textMuted }}>{m.deptName}</span>
                  <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[m.status].color }}>{m.rate.toFixed(1)}%</span>
                  <Bar value={m.rate} status={m.status} h={5} />
                  <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={m.status} small /></div>
                </div>
              ))}
            </Card>
          </Pane>
        </>)}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   MANAGER PORTAL
   ───────────────────────────────────────────────────────────── */
function ManagerPortal({ user, onLogout, state, dispatch }) {
  const [page, setPage] = useState("dashboard");
  const [newProj, setNewProj] = useState({ name: "", due: "" });
  const [editProjId, setEditProjId] = useState(null);
  const [editProjForm, setEditProjForm] = useState({ progress: 0, status: "active", log: "", due: "" });

  const { depts, memberData, weeklySubs, projects, monthlyReports, users } = state;
  const dept = depts.find(d => d.id === user.deptId);
  const myMembers = users.filter(u => u.role === "member" && u.deptId === user.deptId);
  const myTeamMemberIds = myMembers.map(u => u.id);
  const pendingSubs = weeklySubs.filter(s => myTeamMemberIds.includes(s.memberId) && s.approval === "pending");
  const myProjects = projects.filter(p => p.mgrId === user.id);

  const navItems = [
    { id: "dashboard",  icon: "⧉", label: "Team Dashboard"     },
    { id: "dept-kpis",  icon: "◎", label: "Department KPIs"    },
    { id: "approvals",  icon: "✓", label: "Approve Submissions" },
    { id: "projects",   icon: "⚡", label: "Projects"           },
    { id: "members",    icon: "✎", label: "Edit Member KPIs"   },
    { id: "reports",    icon: "⊞", label: "Monthly Reports"    },
  ];

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} pendingCounts={{ approvals: pendingSubs.length }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "dept-kpis" && (<>
          <Header title={`${dept?.name || "Department"} KPIs`} sub="Update actual values for your department's key results" />
          <Pane>
            {!dept && <EmptyState text="No department assigned to your account." />}
            {dept && (() => {
              const KCOL = "50px 1fr 100px 110px 150px 55px 130px 65px";
              const renderKrTable = (krs, deptId, teamId) => {
                if (krs.length === 0) return <div style={{ fontSize: 11, color: T.textMuted, padding: "10px 0" }}>No key results set up yet.</div>;
                return (
                  <Card style={{ overflow: "hidden", marginBottom: 14 }}>
                    <div style={{ display: "grid", gridTemplateColumns: KCOL, padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                      <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span>Data Source</span><span style={{ textAlign: "right" }}>%</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
                    </div>
                    {krs.map((kr, i) => {
                      const pct = kr.target > 0 ? Math.min((kr.actual / kr.target) * 100, 100) : 0;
                      const st = getStatus(pct);
                      return (
                        <div key={kr.id} style={{ display: "grid", gridTemplateColumns: KCOL, padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                          <span style={{ fontFamily: F.mono, fontSize: 10, color: T.textDim }}>{kr.id}</span>
                          <div>
                            <span title={kr.label} style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", display: "block" }}>{kr.label}</span>
                            {kr.unit && <span style={{ fontSize: 9, color: T.textMuted }}>{kr.unit}</span>}
                          </div>
                          <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.target)}{kr.unit ? ` ${kr.unit}` : ""}</span>
                          <Input value={kr.actual} onChange={e => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                          <span style={{ fontSize: 10, color: T.textMuted, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{kr.dataSource || "—"}</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[st].color }}>{pct.toFixed(0)}%</span>
                          <Bar value={pct} status={st} h={5} />
                          <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={st} small /></div>
                        </div>
                      );
                    })}
                  </Card>
                );
              };
              const deptRate = calcRate(dept.krs); const deptStatus = getStatus(deptRate);
              const myTeams = dept.teams.filter(t => user.teamIds?.includes(t.id));
              return (<>
                <div style={{ display: "flex", gap: 14, flexWrap: "wrap", marginBottom: 4 }}>
                  <Metric label="Dept Completion" value={`${deptRate.toFixed(1)}%`} status={deptStatus} />
                  <Metric label="Dept KRs" value={dept.krs.length} />
                  <Metric label="My Teams" value={myTeams.length} />
                </div>
                <SectionLabel>Department Key Results</SectionLabel>
                {renderKrTable(dept.krs, dept.id, null)}
                {myTeams.length > 0 && (<>
                  <SectionLabel>My Team Key Results</SectionLabel>
                  {myTeams.map(t => {
                    const tr = calcRate(t.krs); const ts = getStatus(tr);
                    return (
                      <div key={t.id} style={{ marginBottom: 16 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                          <span style={{ fontSize: 13, fontWeight: 700 }}>{t.name}</span>
                          {t.lead && <span style={{ fontSize: 10, color: T.textMuted }}>Lead: {t.lead}</span>}
                          <span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[ts].color, marginLeft: "auto" }}>{tr.toFixed(1)}%</span>
                          <Tag type={ts} small />
                        </div>
                        {renderKrTable(t.krs, dept.id, t.id)}
                      </div>
                    );
                  })}
                </>)}
              </>);
            })()}
          </Pane>
        </>)}

        {page === "dashboard" && (<>
          <Header title={`${dept?.name || "Team"} Dashboard`} sub={`${dept?.college} · Manager view`} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Dept Completion"   value={`${calcRate(dept?.krs || []).toFixed(1)}%`} status={getStatus(calcRate(dept?.krs || []))} sub={`Time: ${TP}%`} />
              <Metric label="My Members"        value={myMembers.length} />
              <Metric label="Pending Approvals" value={pendingSubs.length} status={pendingSubs.length > 0 ? "yellow" : undefined} />
            </div>
            <SectionLabel>My Team Members</SectionLabel>
            {myMembers.map(m => {
              const kd = memberData[m.id]; if (!kd) return null;
              const r = calcRate(kd.krs); const s = getStatus(r);
              const lastSub = weeklySubs.filter(x => x.memberId === m.id).sort((a, b) => b.date.localeCompare(a.date))[0];
              return (
                <Card key={m.id} style={{ padding: "14px 18px", marginBottom: 8 }}>
                  <div style={{ display: "grid", gridTemplateColumns: "36px 1fr 120px 55px 150px 70px", alignItems: "center", gap: 12 }}>
                    <Avatar letters={m.av} size={30} />
                    <div><div style={{ fontSize: 13, fontWeight: 700 }}>{m.name}</div><div style={{ fontSize: 10, color: T.textMuted }}>{m.title} · Last: {lastSub ? lastSub.week : "No submission"}</div></div>
                    <div style={{ display: "flex", justifyContent: "flex-end" }}>{lastSub && <Tag type={lastSub.approval} small label={APPROVAL[lastSub.approval].label} />}</div>
                    <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</span>
                    <Bar value={r} status={s} h={6} />
                    <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={s} small /></div>
                  </div>
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "approvals" && (<>
          <Header title="Approve Member Submissions" sub={`${pendingSubs.length} pending review`} />
          <Pane>
            {pendingSubs.length === 0 && <EmptyState text="All member submissions have been reviewed." />}
            {weeklySubs.filter(s => myTeamMemberIds.includes(s.memberId)).sort((a, b) => { const o = { pending: 0, approved: 1, rejected: 2 }; return o[a.approval] - o[b.approval] || b.date.localeCompare(a.date); }).map(sub => {
              const mem = users.find(u => u.id === sub.memberId);
              return (
                <Card key={sub.id} style={{ padding: "16px 20px", borderLeft: sub.approval === "pending" ? `3px solid ${T.warn}` : sub.approval === "rejected" ? `3px solid ${T.bad}` : `3px solid ${T.ok}` }}>
                  <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <Avatar letters={mem?.av || "?"} size={28} />
                      <div><div style={{ fontSize: 13, fontWeight: 700 }}>{mem?.name}</div><div style={{ fontSize: 10, color: T.textMuted }}>{mem?.title}</div></div>
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                      <div style={{ textAlign: "right" }}><div style={{ fontSize: 12, fontWeight: 600 }}>{sub.week}</div><div style={{ fontSize: 10, color: T.textMuted }}>{sub.date}</div></div>
                      <button onClick={() => { if (window.confirm(`Delete submission by ${mem?.name || "member"} for ${sub.week}?`)) dispatch({ type: "REMOVE_WEEKLY_SUB", subId: sub.id }); }} style={{ background: "none", border: "none", cursor: "pointer", color: T.bad, fontSize: 15, lineHeight: 1, padding: "2px 4px", borderRadius: 4 }} title="Delete submission">✕</button>
                    </div>
                  </div>
                  <p style={{ margin: "0 0 12px", fontSize: 12, color: T.textSoft, lineHeight: 1.6, padding: "10px 14px", background: T.raised, borderRadius: 7 }}>{sub.items}</p>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <Tag type={sub.approval} label={APPROVAL[sub.approval].label} />
                    {sub.approval === "pending" && (
                      <div style={{ display: "flex", gap: 8 }}>
                        <Btn danger small onClick={() => dispatch({ type: "APPROVE_SUB", subId: sub.id, status: "rejected" })}>Reject</Btn>
                        <Btn primary small onClick={() => dispatch({ type: "APPROVE_SUB", subId: sub.id, status: "approved" })}>Approve</Btn>
                      </div>
                    )}
                    {sub.approval !== "pending" && sub.mgrNote && <span style={{ fontSize: 11, color: T.textMuted, fontStyle: "italic" }}>Note: {sub.mgrNote}</span>}
                  </div>
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "projects" && (<>
          <Header title="Projects" sub="Create and track team projects" />
          <Pane>
            <Card style={{ padding: 20 }}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14 }}>Create New Project</div>
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
                      <div style={{ fontSize: 14, fontWeight: 700 }}>{p.name}</div>
                      <div style={{ fontSize: 10, color: T.textMuted }}>Due: {p.due}{p.updatedDate ? ` · Last updated: ${p.updatedDate}` : ""}</div>
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
                        <span style={{ fontFamily: F.mono, fontSize: 13, fontWeight: 800, color: STATUS_THEME[ps].color, whiteSpace: "nowrap" }}>{p.progress}%</span>
                      </div>
                      {p.log && <p style={{ margin: 0, fontSize: 12, color: T.textSoft, lineHeight: 1.6, padding: "10px 14px", background: T.raised, borderRadius: 7 }}>{p.log}</p>}
                    </div>
                  )}
                  {isEditing && (
                    <div style={{ padding: "14px 18px" }}>
                      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12, marginBottom: 12 }}>
                        <div>
                          <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Completion %</div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                            <input type="range" min={0} max={100} value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Number(e.target.value) }))} style={{ flex: 1 }} />
                            <Input value={editProjForm.progress} onChange={e => setEditProjForm(f => ({ ...f, progress: Math.min(100, Math.max(0, Number(e.target.value) || 0)) }))} style={{ width: 55, textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                            <span style={{ fontSize: 11, color: T.textMuted }}>%</span>
                          </div>
                        </div>
                        <div>
                          <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Status</div>
                          <select value={editProjForm.status} onChange={e => setEditProjForm(f => ({ ...f, status: e.target.value }))} style={{ width: "100%", padding: "7px 10px", background: T.surface, border: `1px solid ${T.border}`, borderRadius: 6, color: T.text, fontSize: 12, fontFamily: F.body }}>
                            <option value="active">Active</option>
                            <option value="completed">Completed</option>
                          </select>
                        </div>
                        <div>
                          <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Due Date</div>
                          <Input type="date" value={editProjForm.due} onChange={e => setEditProjForm(f => ({ ...f, due: e.target.value }))} style={{ width: "100%", padding: "7px 10px", fontSize: 12 }} />
                        </div>
                      </div>
                      <div style={{ marginBottom: 12 }}>
                        <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Project Log / Notes</div>
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
            {myMembers.map(m => {
              const kd = memberData[m.id]; if (!kd) return null;
              const r = calcRate(kd.krs); const s = getStatus(r);
              return (
                <Card key={m.id} style={{ overflow: "hidden" }}>
                  <div style={{ padding: "14px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}><Avatar letters={m.av} size={30} /><div><div style={{ fontSize: 13, fontWeight: 700 }}>{m.name}</div><div style={{ fontSize: 10, color: T.textMuted }}>{m.title}</div></div></div>
                    <div style={{ display: "flex", alignItems: "center", gap: 10 }}><span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</span><Tag type={s} /></div>
                  </div>
                  {kd.krs.map((kr, ki) => {
                    const cr = Math.min((kr.actual / kr.target) * 100, 100); const cs = getStatus(cr);
                    return (
                      <div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 80px 100px 55px 130px", padding: "9px 18px", gap: 8, alignItems: "center", background: ki % 2 ? T.raised : "transparent", borderBottom: ki < kd.krs.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 12 }}>
                        <span style={{ fontFamily: F.mono, fontSize: 10, color: T.textDim }}>{kr.id}</span>
                        <span>{kr.label}</span>
                        <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.target)}</span>
                        <Input value={kr.actual} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: m.id, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                        <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span>
                        <Bar value={cr} status={cs} h={5} />
                      </div>
                    );
                  })}
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "reports" && (<>
          <Header title="Monthly KPI Reports" sub="Published company-wide reports — visible to all teams" />
          <Pane>
            {monthlyReports.length === 0 && <EmptyState text="No monthly reports published yet." />}
            {monthlyReports.map(r => (
              <Card key={r.id} style={{ overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between" }}>
                  <div><div style={{ fontSize: 16, fontWeight: 800 }}>{r.month}</div><div style={{ fontSize: 10, color: T.textMuted }}>Published: {r.publishedDate}</div></div>
                  <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
                </div>
                <div style={{ padding: "14px 20px" }}>
                  {r.data.deptRanks.map((d, i) => (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 12 }}>
                      <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                      <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                      <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{d.rate}%</span>
                      <Tag type={d.status} small />
                    </div>
                  ))}
                </div>
              </Card>
            ))}
          </Pane>
        </>)}
      </div>
    </div>
  );
}

/* ─────────────────────────────────────────────────────────────
   MEMBER PORTAL
   ───────────────────────────────────────────────────────────── */
function MemberPortal({ user, onLogout, state, dispatch }) {
  const [page, setPage] = useState("mykpis");
  const [newOut, setNewOut] = useState({ week: currentWeekLabel(), items: "" });

  const { memberData, weeklySubs, monthlyReports } = state;
  const kd = memberData[user.id] || { krs: [] };
  const mySubs = weeklySubs.filter(s => s.memberId === user.id).sort((a, b) => b.date.localeCompare(a.date));
  const rate = calcRate(kd.krs); const st = getStatus(rate);
  const pendingCount = mySubs.filter(s => s.approval === "pending").length;
  const thisWeekSub = mySubs.find(s => s.week === currentWeekLabel());

  const navItems = [
    { id: "mykpis",  icon: "◎", label: "My KPIs"          },
    { id: "submit",  icon: "✎", label: "Weekly Submission" },
    { id: "history", icon: "⊞", label: "My History"        },
    { id: "reports", icon: "⊠", label: "Monthly Reports"  },
  ];

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} pendingCounts={{ submit: thisWeekSub ? 0 : 1 }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "mykpis" && (<>
          <Header title="My KPIs" sub={`${user.title} · FY26 Q1`} right={<Tag type={st} />} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="My Completion"  value={`${rate.toFixed(1)}%`} status={st} sub={`Time: ${TP}%`} />
              <Metric label="KRs Tracked"    value={kd.krs.length} />
              <Metric label="This Week"      value={thisWeekSub ? "Submitted" : "Due"} status={thisWeekSub ? "green" : "red"} />
              <Metric label="Pending Review" value={pendingCount} status={pendingCount > 0 ? "yellow" : undefined} />
            </div>
            {kd.krs.map(kr => {
              const r = Math.min((kr.actual / kr.target) * 100, 100); const s = getStatus(r);
              return (
                <Card key={kr.id} style={{ padding: "18px 20px" }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
                    <div><div style={{ fontSize: 10, fontWeight: 700, color: T.textDim, fontFamily: F.mono, marginBottom: 3 }}>{kr.id}</div><div style={{ fontSize: 14, fontWeight: 700 }}>{kr.label}</div></div>
                    <Tag type={s} />
                  </div>
                  <div style={{ display: "flex", alignItems: "flex-end", gap: 20 }}>
                    <div><div style={{ fontSize: 30, fontWeight: 900, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{fmt(kr.actual)}</div><div style={{ fontSize: 10, color: T.textMuted }}>of {fmt(kr.target)} target</div></div>
                    <div style={{ flex: 1 }}><Bar value={r} status={s} h={10} /></div>
                    <div style={{ fontSize: 22, fontWeight: 800, fontFamily: F.mono, color: STATUS_THEME[s].color }}>{r.toFixed(1)}%</div>
                  </div>
                </Card>
              );
            })}
          </Pane>
        </>)}

        {page === "submit" && (<>
          <Header title="Weekly Submission" sub="Submit your work outcomes — due every week"
            right={thisWeekSub ? <Tag type="approved" label="This week: Submitted" /> : <Tag type="rejected" label="This week: Not yet submitted" />} />
          <Pane>
            <Card style={{ padding: 20 }}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14 }}>Update KPI Actuals</div>
              {kd.krs.map((kr, i) => {
                const r = Math.min((kr.actual / kr.target) * 100, 100); const s = getStatus(r);
                return (
                  <div key={kr.id} style={{ display: "grid", gridTemplateColumns: "1fr 70px 100px 50px 130px", padding: "9px 0", gap: 10, alignItems: "center", borderBottom: i < kd.krs.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 12 }}>
                    <div><div style={{ fontWeight: 600 }}>{kr.label}</div><div style={{ fontSize: 10, color: T.textMuted }}>Target: {fmt(kr.target)}</div></div>
                    <span style={{ fontSize: 10, color: T.textMuted, textAlign: "right" }}>Actual:</span>
                    <Input value={kr.actual} onChange={e => dispatch({ type: "UPDATE_MEMBER_KR", memberId: user.id, krId: kr.id, field: "actual", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "7px 10px", fontSize: 13, fontFamily: F.mono }} />
                    <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[s].color }}>{r.toFixed(0)}%</span>
                    <Bar value={r} status={s} h={5} />
                  </div>
                );
              })}
            </Card>
            <Card style={{ padding: 20 }}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14 }}>Work Outcome Summary</div>
              <Input value={newOut.week} onChange={e => setNewOut(p => ({ ...p, week: e.target.value }))} style={{ width: 220, marginBottom: 10 }} />
              <TextArea value={newOut.items} onChange={e => setNewOut(p => ({ ...p, items: e.target.value }))} placeholder="What did you accomplish this week? List your key tasks, wins, and any blockers..." rows={5} />
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 14 }}>
                <span style={{ fontSize: 10, color: T.textMuted }}>Your submission will be sent to your manager for approval</span>
                <Btn primary disabled={!newOut.items} onClick={() => {
                  dispatch({ type: "ADD_WEEKLY_SUB", sub: { id: `ws${Date.now()}`, memberId: user.id, week: newOut.week, items: newOut.items, date: new Date().toISOString().slice(0, 10), approval: "pending", mgrNote: "" } });
                  setNewOut({ week: currentWeekLabel(), items: "" });
                }}>Submit for Approval</Btn>
              </div>
            </Card>
          </Pane>
        </>)}

        {page === "history" && (<>
          <Header title="My Submission History" sub="All weekly submissions and their approval status" />
          <Pane>
            {mySubs.length === 0 && <EmptyState text="No submissions yet — go to Weekly Submission to log your first one." />}
            {mySubs.map(s => (
              <Card key={s.id} style={{ padding: "16px 20px", borderLeft: `3px solid ${APPROVAL[s.approval].color}` }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <span style={{ fontSize: 13, fontWeight: 700 }}>{s.week}</span>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}><span style={{ fontSize: 10, color: T.textMuted }}>{s.date}</span><Tag type={s.approval} label={APPROVAL[s.approval].label} small /></div>
                </div>
                <p style={{ margin: 0, fontSize: 12, color: T.textSoft, lineHeight: 1.6 }}>{s.items}</p>
                {s.mgrNote && <div style={{ marginTop: 8, padding: "8px 12px", background: T.raised, borderRadius: 6, fontSize: 11, color: T.textMuted }}><strong style={{ color: T.textSoft }}>Manager Note:</strong> {s.mgrNote}</div>}
              </Card>
            ))}
          </Pane>
        </>)}

        {page === "reports" && (<>
          <Header title="Monthly KPI Reports" sub="Company-wide reports — published at end of each month" />
          <Pane>
            {monthlyReports.length === 0 && <EmptyState text="No monthly reports published yet." />}
            {monthlyReports.map(r => (
              <Card key={r.id} style={{ overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between" }}>
                  <div><div style={{ fontSize: 16, fontWeight: 800 }}>{r.month}</div><div style={{ fontSize: 10, color: T.textMuted }}>Published: {r.publishedDate} · Visible to everyone</div></div>
                  <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
                </div>
                <div style={{ padding: "14px 20px", display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                  <div>
                    <SectionLabel>Department Rankings</SectionLabel>
                    {r.data.deptRanks.map((d, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 10, padding: "5px 0", fontSize: 12 }}>
                        <span style={{ fontFamily: F.mono, fontWeight: 800, color: i === 0 ? T.ok : T.textMuted, width: 22 }}>#{i + 1}</span>
                        <span style={{ flex: 1, fontWeight: 600 }}>{d.name}</span>
                        <span style={{ fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[d.status].color }}>{d.rate}%</span>
                      </div>
                    ))}
                  </div>
                  <div>
                    <SectionLabel>Top Performers</SectionLabel>
                    {r.data.topPerformers.map((p, i) => <div key={i} style={{ padding: "4px 0", fontSize: 12, color: T.ok }}>★ {p}</div>)}
                    {r.data.redFlags?.length > 0 && (<><div style={{ marginTop: 10 }} /><SectionLabel>Needs Improvement</SectionLabel>{r.data.redFlags.map((f, i) => <div key={i} style={{ padding: "4px 0", fontSize: 12, color: T.bad }}>⚠ {f}</div>)}</>)}
                  </div>
                </div>
              </Card>
            ))}
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
    case "UPDATE_MEMBER_KR": return { ...state, memberData: { ...state.memberData, [action.memberId]: { ...state.memberData[action.memberId], krs: state.memberData[action.memberId].krs.map(kr => kr.id === action.krId ? { ...kr, [action.field]: action.value } : kr) } } };
    case "ADD_WEEKLY_SUB":  return { ...state, weeklySubs:  [action.sub,    ...state.weeklySubs]  };
    case "APPROVE_SUB":     return { ...state, weeklySubs: state.weeklySubs.map(s => s.id === action.subId ? { ...s, approval: action.status } : s) };
    case "REMOVE_WEEKLY_SUB": return { ...state, weeklySubs: state.weeklySubs.filter(s => s.id !== action.subId) };
    case "ADD_MGR_SPRINT":    return { ...state, mgrSprints: [action.sprint, ...state.mgrSprints] };
    case "REMOVE_MGR_SPRINT": return { ...state, mgrSprints: state.mgrSprints.filter(s => s.id !== action.sprintId) };
    case "ADD_PROJECT":     return { ...state, projects: [...state.projects, action.project] };
    case "UPDATE_PROJECT":  return { ...state, projects: state.projects.map(p => p.id === action.projectId ? { ...p, ...action.updates } : p) };
    case "REMOVE_PROJECT":  return { ...state, projects: state.projects.filter(p => p.id !== action.projectId) };
    case "PUBLISH_REPORT":  return { ...state, monthlyReports: [action.report, ...state.monthlyReports] };
    case "REMOVE_REPORT":   return { ...state, monthlyReports: state.monthlyReports.filter(r => r.id !== action.reportId) };

    case "ADD_USER": {
      const u = action.user;
      const newDepts = u.role === "member" && u.teamId
        ? state.depts.map(d => d.id !== u.deptId ? d : { ...d, teams: d.teams.map(t => t.id !== u.teamId ? t : { ...t, members: [...t.members, u.id] }) })
        : state.depts;
      const newMemberData = u.role === "member"
        ? { ...state.memberData, [u.id]: { krs: [] } }
        : state.memberData;
      return { ...state, users: [...state.users, u], depts: newDepts, memberData: newMemberData };
    }

    case "UPDATE_USER": {
      const prev = state.users.find(u => u.id === action.userId);
      const updated = { ...prev, ...action.updates };
      const uid = action.userId;
      const oldTeam = prev?.role === "member" ? prev.teamId : null;
      const oldDept = prev?.role === "member" ? prev.deptId : null;
      const newTeam = updated.role === "member" ? updated.teamId : null;
      const newDept = updated.role === "member" ? updated.deptId : null;
      const teamChanged = oldTeam !== newTeam || oldDept !== newDept;

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

      return { ...state, users: state.users.map(u => u.id === uid ? updated : u), depts: newDepts, memberData: newMemberData };
    }

    case "REMOVE_USER":
      return { ...state, users: state.users.filter(u => u.id !== action.userId) };

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
      return { ...state, depts: state.depts.map(d => d.id === action.deptId ? { ...d, teams: d.teams.filter(t => t.id !== action.teamId) } : d) };

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
  const { instance, accounts } = useMsal();
  const [state, rawDispatch] = useState({
    users: INIT_USERS,
    depts: INIT_DEPTS,
    memberData: INIT_MEMBER_DATA,
    weeklySubs: INIT_WEEKLY_SUBS,
    mgrSprints: INIT_MGR_SPRINTS,
    projects: INIT_PROJECTS,
    monthlyReports: INIT_MONTHLY_REPORTS,
  });

  // Dispatch updates local state immediately (optimistic), then syncs to Cosmos DB in background.
  const dispatch = useCallback((action) => {
    rawDispatch(prev => {
      const next = appReducer(prev, action);
      syncChanges(prev, next).catch(err => console.error("[DB sync error]", err.message));
      return next;
    });
  }, []);

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
            mgrSprints: data.mgrSprints || [],
            projects: data.projects || [],
            monthlyReports: data.monthlyReports || [],
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
    <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: T.badDim, borderTop: `1px solid ${T.badBorder}`, color: T.bad, fontSize: 11, padding: "8px 20px", textAlign: "center", zIndex: 9999 }}>
      ⚠ {dbError}
    </div>
  ) : null;

  if (activeUser.role === "admin")   return <>{offlineBanner}<AdminPortal   user={activeUser} onLogout={logout} state={state} dispatch={dispatch} /></>;
  if (activeUser.role === "manager") return <>{offlineBanner}<ManagerPortal user={activeUser} onLogout={logout} state={state} dispatch={dispatch} /></>;
  return <>{offlineBanner}<MemberPortal user={activeUser} onLogout={logout} state={state} dispatch={dispatch} /></>;
}
