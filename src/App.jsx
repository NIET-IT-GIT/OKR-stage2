import { useState, useEffect, useCallback, useRef } from "react";
import { useMsal } from "@azure/msal-react";
import { InteractionStatus, EventType } from "@azure/msal-browser";
import { loginRequest } from "./authConfig";

const T = {
  bg: "#090b10", bgSoft: "#0e1118", surface: "#141720", surfaceHover: "#1a1e2b",
  raised: "#1c2030", border: "#222840", borderFocus: "#3d4870",
  text: "#e4e7f1", textSoft: "#9da3be", textMuted: "#636a88", textDim: "#3b4160",
  brand: "#4b7cf3", brandSoft: "#3661d4", brandDim: "rgba(75,124,243,0.07)",
  brandBorder: "rgba(75,124,243,0.22)",
  ok: "#2dd47a", okDim: "rgba(45,212,122,0.07)", okBorder: "rgba(45,212,122,0.22)",
  warn: "#f0b030", warnDim: "rgba(240,176,48,0.07)", warnBorder: "rgba(240,176,48,0.22)",
  bad: "#f04848", badDim: "rgba(240,72,72,0.07)", badBorder: "rgba(240,72,72,0.22)",
  orange: "#e0823a", purple: "#9b6cf7",
};
const F = { body: "'Outfit', 'DM Sans', system-ui, sans-serif", mono: "'JetBrains Mono', 'Fira Code', monospace" };
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

/* ─── UI PRIMITIVES ─── */
function Tag({ type = "green", label, small }) {
  const s = STATUS_THEME[type] || APPROVAL[type] || STATUS_THEME.green;
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", gap: 5,
      background: s.bg, border: `1px solid ${s.border}`, borderRadius: 5,
      padding: small ? "2px 7px" : "3px 10px",
      fontSize: small ? 9 : 10, fontWeight: 700, color: s.color, letterSpacing: "0.03em", whiteSpace: "nowrap",
    }}>
      <span style={{ width: 5, height: 5, borderRadius: "50%", background: s.color }} />
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
    <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: cfg.bg, border: `1px solid ${cfg.border}`, borderRadius: 5, padding: "2px 8px", fontSize: 9, fontWeight: 700, color: cfg.color, whiteSpace: "nowrap" }}>
      {cfg.label}
    </span>
  );
}

function Bar({ value, status, h = 6 }) {
  const c = STATUS_THEME[status]?.color || T.brand;
  return (
    <div style={{ flex: 1, height: h, background: T.raised, borderRadius: h, overflow: "hidden", position: "relative" }}>
      <div style={{ width: `${Math.min(value, 100)}%`, height: "100%", background: c, borderRadius: h, transition: "width 0.5s ease" }} />
      <div style={{ position: "absolute", left: `${TP}%`, top: 0, bottom: 0, width: 1.5, background: T.textDim, opacity: 0.6 }} />
    </div>
  );
}

function Metric({ label, value, sub, status }) {
  return (
    <div style={{ background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, padding: "16px 18px", flex: 1, minWidth: 130 }}>
      <div style={{ fontSize: 10, color: T.textMuted, fontWeight: 600, letterSpacing: "0.05em", textTransform: "uppercase", marginBottom: 8 }}>{label}</div>
      <div style={{ fontSize: 22, fontWeight: 800, color: status ? STATUS_THEME[status]?.color : T.text, fontFamily: F.mono }}>{value}</div>
      {sub && <div style={{ fontSize: 10, color: T.textMuted, marginTop: 3 }}>{sub}</div>}
    </div>
  );
}

function Btn({ children, primary, danger, small, disabled, onClick, style: sx }) {
  return (
    <button onClick={disabled ? undefined : onClick} style={{
      background: primary ? T.brand : danger ? T.bad : "transparent",
      color: primary || danger ? "#fff" : T.textSoft,
      border: primary || danger ? "none" : `1px solid ${T.border}`,
      borderRadius: 7, padding: small ? "6px 12px" : "9px 18px",
      fontSize: small ? 11 : 12, fontWeight: 700, cursor: disabled ? "not-allowed" : "pointer",
      opacity: disabled ? 0.4 : 1, transition: "all 0.15s", fontFamily: F.body, ...sx,
    }}>{children}</button>
  );
}

function Input({ value, onChange, placeholder, type, style: sx, ...props }) {
  return (
    <input type={type || "text"} value={value} onChange={onChange} placeholder={placeholder} {...props}
      style={{
        background: T.bgSoft, border: `1px solid ${T.border}`, borderRadius: 7,
        padding: "10px 14px", color: T.text, fontSize: 12, fontFamily: F.body, outline: "none",
        transition: "border-color 0.2s", boxSizing: "border-box", ...sx,
      }}
      onFocus={e => e.target.style.borderColor = T.brand}
      onBlur={e => e.target.style.borderColor = T.border}
    />
  );
}

function Select({ value, onChange, children, style: sx }) {
  return (
    <select value={value} onChange={onChange}
      style={{
        background: T.bgSoft, border: `1px solid ${T.border}`, borderRadius: 7,
        padding: "10px 14px", color: value ? T.text : T.textMuted,
        fontSize: 12, fontFamily: F.body, outline: "none", cursor: "pointer",
        boxSizing: "border-box", ...sx,
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
        width: "100%", boxSizing: "border-box", background: T.bgSoft, border: `1px solid ${T.border}`, borderRadius: 7,
        padding: "10px 14px", color: T.text, fontSize: 12, fontFamily: F.body, outline: "none", resize: "vertical",
        transition: "border-color 0.2s",
      }}
      onFocus={e => e.target.style.borderColor = T.brand}
      onBlur={e => e.target.style.borderColor = T.border}
    />
  );
}

function Avatar({ letters, size = 32, color }) {
  return (
    <div style={{
      width: size, height: size, borderRadius: "50%", flexShrink: 0,
      background: color || `linear-gradient(135deg, ${T.brandSoft}, ${T.brand})`,
      display: "flex", alignItems: "center", justifyContent: "center",
      fontSize: size * 0.34, fontWeight: 800, color: "#fff",
    }}>{letters}</div>
  );
}

function Card({ children, style: sx, onClick }) {
  return (
    <div onClick={onClick} style={{
      background: T.surface, border: `1px solid ${T.border}`, borderRadius: 10, ...sx,
      cursor: onClick ? "pointer" : "default", transition: "border-color 0.15s",
    }}
      onMouseEnter={onClick ? e => e.currentTarget.style.borderColor = T.brand : undefined}
      onMouseLeave={onClick ? e => e.currentTarget.style.borderColor = T.border : undefined}
    >{children}</div>
  );
}

function SectionLabel({ children }) {
  return <div style={{ fontSize: 10, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 10 }}>{children}</div>;
}

function EmptyState({ text }) {
  return <div style={{ padding: "40px 20px", textAlign: "center", color: T.textMuted, fontSize: 13 }}>{text}</div>;
}

function CountBadge({ count, color }) {
  if (!count) return null;
  return <span style={{ background: color || T.warn, color: "#fff", borderRadius: 10, padding: "1px 7px", fontSize: 9, fontWeight: 800, marginLeft: 6 }}>{count}</span>;
}

function Side({ items, active, onSelect, user, onLogout, pendingCounts }) {
  return (
    <div style={{ width: 250, background: T.bgSoft, borderRight: `1px solid ${T.border}`, display: "flex", flexDirection: "column", height: "100vh", flexShrink: 0 }}>
      <div style={{ padding: "20px 16px 14px", borderBottom: `1px solid ${T.border}` }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 16 }}>
          <div style={{ width: 32, height: 32, borderRadius: 9, background: `linear-gradient(135deg, ${T.brand}, #7c5bf5)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 14, fontWeight: 900, color: "#fff" }}>O</div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 800, color: T.text }}>OKR Tracker</div>
            <div style={{ fontSize: 9, color: T.textMuted, letterSpacing: "0.08em" }}>NIET GROUP</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 12px", background: T.surface, borderRadius: 8, border: `1px solid ${T.border}` }}>
          <Avatar letters={user.av} size={34} />
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: T.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{user.name}</div>
            <div style={{ fontSize: 9, color: T.textMuted }}>{user.title}</div>
          </div>
        </div>
      </div>
      <div style={{ flex: 1, padding: "10px 8px", display: "flex", flexDirection: "column", gap: 2, overflowY: "auto" }}>
        {items.map(item => (
          <button key={item.id} onClick={() => onSelect(item.id)} style={{
            background: active === item.id ? T.brandDim : "transparent",
            border: active === item.id ? `1px solid ${T.brandBorder}` : "1px solid transparent",
            borderRadius: 8, padding: "10px 14px", cursor: "pointer",
            display: "flex", alignItems: "center", gap: 10,
            color: active === item.id ? T.brand : T.textSoft,
            fontSize: 12, fontWeight: active === item.id ? 700 : 500, textAlign: "left", width: "100%",
            transition: "all 0.12s", fontFamily: F.body,
          }}>
            <span style={{ fontSize: 14, width: 20, textAlign: "center", flexShrink: 0 }}>{item.icon}</span>
            <span style={{ flex: 1 }}>{item.label}</span>
            {pendingCounts?.[item.id] > 0 && <CountBadge count={pendingCounts[item.id]} />}
          </button>
        ))}
      </div>
      <div style={{ padding: "10px 8px", borderTop: `1px solid ${T.border}` }}>
        <button onClick={onLogout} style={{ background: "none", border: "none", borderRadius: 8, padding: "10px 14px", cursor: "pointer", display: "flex", alignItems: "center", gap: 10, color: T.textMuted, fontSize: 12, width: "100%", fontFamily: F.body }}>
          <span>↩</span> Sign Out
        </button>
      </div>
    </div>
  );
}

function Header({ title, sub, right }) {
  return (
    <div style={{ padding: "20px 28px 16px", borderBottom: `1px solid ${T.border}`, display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
      <div>
        <h1 style={{ margin: 0, fontSize: 20, fontWeight: 800, color: T.text }}>{title}</h1>
        {sub && <p style={{ margin: "3px 0 0", fontSize: 12, color: T.textMuted }}>{sub}</p>}
      </div>
      {right}
    </div>
  );
}

function Pane({ children }) {
  return <div style={{ padding: 28, display: "flex", flexDirection: "column", gap: 22 }}>{children}</div>;
}

/* ─────────────────────────────────────────────────────────────
   LOGIN PAGE
   ───────────────────────────────────────────────────────────── */
function LoginPage({ onLogin, users, inProgress, msalErr }) {
  const { instance } = useMsal();
  const [show, setShow] = useState(false);
  const [msLoading, setMsLoading] = useState(false);
  const [showAdminLogin, setShowAdminLogin] = useState(false);
  const [adminUser, setAdminUser] = useState("");
  const [adminPass, setAdminPass] = useState("");
  const [err, setErr] = useState("");

  useEffect(() => { setTimeout(() => setShow(true), 80); }, []);

  // Show a clean loading screen while MSAL is processing the redirect response
  // so the landing page never flashes after coming back from Microsoft.
  if (inProgress === InteractionStatus.HandleRedirect || inProgress === InteractionStatus.Login) {
    return (
      <div style={{ minHeight: "100vh", background: T.bg, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: F.body }}>
        <div style={{ textAlign: "center" }}>
          <div style={{ width: 40, height: 40, borderRadius: "50%", border: `3px solid ${T.border}`, borderTopColor: T.brand, animation: "spin 0.8s linear infinite", margin: "0 auto 16px" }} />
          <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
          <div style={{ color: T.textMuted, fontSize: 13 }}>Signing you in…</div>
        </div>
      </div>
    );
  }

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
    <div style={{ minHeight: "100vh", background: T.bg, display: "flex", fontFamily: F.body, position: "relative", overflow: "hidden" }}>
      <div style={{ position: "absolute", inset: 0, opacity: 0.03, backgroundImage: `radial-gradient(${T.brand} 1px, transparent 1px)`, backgroundSize: "30px 30px" }} />
      <div style={{ position: "absolute", top: "-25%", right: "-8%", width: 650, height: 650, background: `radial-gradient(circle, ${T.brand}12, transparent 65%)`, borderRadius: "50%" }} />
      <div style={{ position: "absolute", bottom: "-20%", left: "-10%", width: 500, height: 500, background: `radial-gradient(circle, ${T.purple}15, transparent 65%)`, borderRadius: "50%" }} />

      {/* Hero */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 72px", position: "relative", zIndex: 1, opacity: show ? 1 : 0, transform: show ? "none" : "translateX(-20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1)" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 44 }}>
          <div style={{ width: 46, height: 46, borderRadius: 13, background: `linear-gradient(135deg, ${T.brand}, #7c5bf5)`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22, fontWeight: 900, color: "#fff" }}>O</div>
          <div>
            <div style={{ fontSize: 21, fontWeight: 900, color: T.text }}>NIET OKR Performance Tracker</div>
            <div style={{ fontSize: 10, color: T.textMuted, letterSpacing: "0.14em" }}>NIET · CHARLTON BROWN · RHODES · EDUCARE</div>
          </div>
        </div>
        <h1 style={{ margin: "0 0 14px", fontSize: 44, fontWeight: 900, lineHeight: 1.08, color: T.text, letterSpacing: "-0.03em", maxWidth: 460 }}>
          Align goals.<br /><span style={{ color: T.brand }}>Track everyone.</span><br />Drive results.
        </h1>
        <p style={{ margin: 0, fontSize: 14, color: T.textSoft, lineHeight: 1.7, maxWidth: 400 }}>
          Monthly KPI reporting, weekly outcome tracking, real-time leaderboards — full transparency from company goals down to every team member.
        </p>
        <div style={{ marginTop: 48, display: "flex", gap: 36 }}>
          {[{ n: "Monthly", l: "KPI Reports" }, { n: "Weekly", l: "Submissions" }, { n: "Real-time", l: "Rankings" }, { n: "100%", l: "Transparent" }].map((x, i) => (
            <div key={i}>
              <div style={{ fontSize: 18, fontWeight: 900, color: T.brand, fontFamily: F.mono }}>{x.n}</div>
              <div style={{ fontSize: 9, color: T.textMuted, letterSpacing: "0.06em", marginTop: 2 }}>{x.l}</div>
            </div>
          ))}
        </div>
      </div>

      {/* Sign-in card */}
      <div style={{ width: 420, display: "flex", flexDirection: "column", justifyContent: "center", padding: "60px 44px", position: "relative", zIndex: 1, opacity: show ? 1 : 0, transform: show ? "none" : "translateY(20px)", transition: "all 0.7s cubic-bezier(0.16,1,0.3,1) 0.15s" }}>
        <Card style={{ padding: "36px 30px" }}>
          <h2 style={{ margin: "0 0 4px", fontSize: 18, fontWeight: 800, color: T.text }}>Sign in</h2>
          <p style={{ margin: "0 0 24px", fontSize: 12, color: T.textMuted }}>Use your NIET Microsoft account to access your portal.</p>

          {(err || msalErr) && (
            <div style={{ padding: "10px 14px", background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 7, fontSize: 12, color: T.bad, marginBottom: 16, lineHeight: 1.5 }}>{err || msalErr}</div>
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
            {form.role === "member" && form.deptId && (
              <div>
                <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 5 }}>Team</div>
                <Select value={form.teamId} onChange={e => setForm(p => ({ ...p, teamId: e.target.value }))} style={{ width: "100%" }}>
                  <option value="">— Select team —</option>
                  {teamsForDept(form.deptId).map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                </Select>
              </div>
            )}
            {form.role === "manager" && form.deptId && teamsForDept(form.deptId).length > 0 && (
              <div style={{ gridColumn: "1 / -1" }}>
                <div style={{ fontSize: 9, fontWeight: 700, color: T.textMuted, letterSpacing: "0.06em", textTransform: "uppercase", marginBottom: 8 }}>Manages Teams</div>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {teamsForDept(form.deptId).map(t => (
                    <label key={t.id} style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: T.textSoft, cursor: "pointer", padding: "6px 10px", background: form.teamIds.includes(t.id) ? T.brandDim : T.raised, border: `1px solid ${form.teamIds.includes(t.id) ? T.brandBorder : T.border}`, borderRadius: 6 }}>
                      <input type="checkbox" checked={form.teamIds.includes(t.id)} onChange={e => setForm(p => ({ ...p, teamIds: e.target.checked ? [...p.teamIds, t.id] : p.teamIds.filter(id => id !== t.id) }))} style={{ accentColor: T.brand }} />
                      {t.name}
                    </label>
                  ))}
                </div>
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
        <div style={{ display: "grid", gridTemplateColumns: "32px 1fr 180px 80px 120px 120px 90px", padding: "7px 18px", gap: 10, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
          <span></span><span>Name / Email</span><span>Title</span><span>Role</span><span>Department</span><span>Team</span><span style={{ textAlign: "right" }}>Actions</span>
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
                  {editForm.role === "member" && editForm.deptId && (
                    <Select value={editForm.teamId} onChange={e => setEditForm(p => ({ ...p, teamId: e.target.value }))} style={{ fontSize: 11, padding: "7px 10px" }}>
                      <option value="">— Team —</option>
                      {editTeams.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
                    </Select>
                  )}
                </div>
                {editForm.role === "manager" && editForm.deptId && editTeams.length > 0 && (
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                    {editTeams.map(t => (
                      <label key={t.id} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 11, color: T.textSoft, cursor: "pointer", padding: "5px 9px", background: editForm.teamIds.includes(t.id) ? T.brandDim : T.raised, border: `1px solid ${editForm.teamIds.includes(t.id) ? T.brandBorder : T.border}`, borderRadius: 5 }}>
                        <input type="checkbox" checked={editForm.teamIds.includes(t.id)} onChange={e => setEditForm(p => ({ ...p, teamIds: e.target.checked ? [...p.teamIds, t.id] : p.teamIds.filter(id => id !== t.id) }))} style={{ accentColor: T.brand }} />
                        {t.name}
                      </label>
                    ))}
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
            <div key={u.id} style={{ display: "grid", gridTemplateColumns: "32px 1fr 180px 80px 120px 120px 90px", padding: "10px 18px", gap: 10, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
              <Avatar letters={u.av} size={26} />
              <div>
                <div style={{ fontWeight: 600, color: T.text }}>{u.name}{isSelf && <span style={{ fontSize: 9, color: T.brand, marginLeft: 6 }}>you</span>}</div>
                <div style={{ fontSize: 10, color: T.textMuted }}>{isSystem ? "System login only" : u.email}</div>
              </div>
              <span style={{ fontSize: 11, color: T.textSoft }}>{u.title}</span>
              <RoleTag role={u.role} />
              <span style={{ fontSize: 11, color: T.textMuted }}>{dept?.name || "—"}</span>
              <span style={{ fontSize: 11, color: T.textMuted }}>{team?.name || (u.teamIds?.length ? `${u.teamIds.length} teams` : "—")}</span>
              <div style={{ display: "flex", gap: 6, justifyContent: "flex-end" }}>
                {!isSystem && (
                  <button onClick={() => startEdit(u)} style={{ background: T.brandDim, border: `1px solid ${T.brandBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.brand, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>Edit</button>
                )}
                {!isSystem && !isSelf && (
                  <button onClick={() => dispatch({ type: "REMOVE_USER", userId: u.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 9px", cursor: "pointer", color: T.bad, fontSize: 10, fontWeight: 700, fontFamily: F.body }}>✕</button>
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
   ADMIN PORTAL
   ───────────────────────────────────────────────────────────── */
function AdminPortal({ user, onLogout, state, dispatch }) {
  const [page, setPage] = useState("overview");
  const [selDept, setSelDept] = useState(null);
  const [selTeam, setSelTeam] = useState(null);
  const [newKr, setNewKr] = useState({ label: "", target: "" });
  const [addTarget, setAddTarget] = useState(null);

  const { depts, memberData, mgrSprints, monthlyReports, users } = state;
  const navItems = [
    { id: "overview",    icon: "◎", label: "Company Overview"  },
    { id: "departments", icon: "⬛", label: "Departments"       },
    { id: "setup",       icon: "⚙", label: "OKR / KPI Setup"   },
    { id: "reports",     icon: "⊞", label: "Monthly Reports"   },
    { id: "sprints",     icon: "↻", label: "Manager Sprints"   },
    { id: "leaderboard", icon: "▲", label: "Leaderboard"       },
    { id: "users",       icon: "⊹", label: "User Management"   },
  ];

  const deptRanks = depts.map(d => ({ ...d, rate: calcRate(d.krs), status: getStatus(calcRate(d.krs)) })).sort((a, b) => b.rate - a.rate);
  const compRate = deptRanks.length ? deptRanks.reduce((a, d) => a + d.rate, 0) / deptRanks.length : 0;
  const allMembers = [];
  depts.forEach(d => d.teams.forEach(t => t.members.forEach(mId => {
    const u = users.find(x => x.id === mId); const kd = memberData[mId];
    if (u && kd) { const r = calcRate(kd.krs); allMembers.push({ ...u, dept: d.name, team: t.name, rate: r, status: getStatus(r) }); }
  })));
  allMembers.sort((a, b) => b.rate - a.rate);

  function addKr(deptId, teamId) {
    if (!newKr.label || !newKr.target) return;
    const kr = { id: `N${Date.now().toString(36).slice(-4).toUpperCase()}`, label: newKr.label, target: Number(newKr.target), actual: 0 };
    dispatch({ type: "ADD_KR", deptId, teamId, kr });
    setNewKr({ label: "", target: "" }); setAddTarget(null);
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

        {page === "departments" && (<>
          <Header title="Department Detail" sub="View KRs, teams, and member performance" />
          <Pane>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
              {depts.map(d => <Btn key={d.id} primary={selDept === d.id} small onClick={() => { setSelDept(d.id); setSelTeam(null); }}>{d.name}</Btn>)}
            </div>
            {selDept && (() => {
              const dept = depts.find(d => d.id === selDept); if (!dept) return null;
              const r = calcRate(dept.krs); const s = getStatus(r);
              return (<>
                <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between" }}>
                  <div><h2 style={{ margin: 0, fontSize: 18, fontWeight: 800 }}>{dept.name}</h2><p style={{ margin: "2px 0", fontSize: 12, color: T.textMuted }}>{dept.obj} · Head: {dept.head}</p></div>
                  <Tag type={s} />
                </div>
                <div style={{ display: "flex", gap: 12 }}><Metric label="Completion" value={`${r.toFixed(1)}%`} status={s} /><Metric label="Teams" value={dept.teams.length} /></div>
                <Card style={{ overflow: "hidden" }}>
                  <div style={{ padding: "10px 16px", borderBottom: `1px solid ${T.border}`, fontSize: 10, fontWeight: 700, color: T.textMuted, letterSpacing: "0.05em" }}>DEPARTMENT KEY RESULTS</div>
                  {dept.krs.map((kr, i) => { const cr = Math.min((kr.actual / kr.target) * 100, 100); const cs = getStatus(cr);
                    return (<div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 80px 80px 60px 150px 70px", padding: "10px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                      <span style={{ fontFamily: F.mono, fontSize: 10, color: T.textDim }}>{kr.id}</span><span>{kr.label}</span>
                      <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.target)}</span>
                      <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700 }}>{fmt(kr.actual)}</span>
                      <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[cs].color }}>{cr.toFixed(0)}%</span>
                      <Bar value={cr} status={cs} h={5} />
                      <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={cs} small /></div>
                    </div>);
                  })}
                </Card>
                {dept.teams.map(t => { const tr = calcRate(t.krs); const ts = getStatus(tr);
                  const tMembers = t.members.map(mId => { const u = users.find(x => x.id === mId); const kd = memberData[mId]; if (!u || !kd) return null; const mr = calcRate(kd.krs); return { ...u, rate: mr, status: getStatus(mr) }; }).filter(Boolean).sort((a, b) => b.rate - a.rate);
                  return (
                    <Card key={t.id} style={{ overflow: "hidden" }}>
                      <div style={{ padding: "14px 18px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <div><span style={{ fontSize: 13, fontWeight: 700 }}>{t.name}</span><span style={{ fontSize: 10, color: T.textMuted, marginLeft: 10 }}>Lead: {t.lead}</span></div>
                        <div style={{ display: "flex", alignItems: "center", gap: 10 }}><span style={{ fontFamily: F.mono, fontWeight: 800, color: STATUS_THEME[ts].color }}>{tr.toFixed(1)}%</span><Tag type={ts} small /></div>
                      </div>
                      {tMembers.map((m, mi) => (
                        <div key={m.id} style={{ display: "grid", gridTemplateColumns: "28px 36px 1fr 100px 55px 140px 70px", padding: "9px 18px", gap: 10, alignItems: "center", background: mi % 2 ? T.raised : "transparent", borderBottom: mi < tMembers.length - 1 ? `1px solid ${T.border}` : "none", fontSize: 12 }}>
                          <span style={{ fontFamily: F.mono, fontWeight: 800, color: mi === 0 ? T.ok : mi === tMembers.length - 1 && tMembers.length > 2 ? T.bad : T.textMuted }}>#{mi + 1}</span>
                          <Avatar letters={m.av} size={26} />
                          <div><span style={{ fontWeight: 600 }}>{m.name}</span><span style={{ color: T.textDim, marginLeft: 6, fontSize: 10 }}>{m.title}</span></div>
                          <span style={{ fontSize: 10, color: T.textMuted, textAlign: "right" }}>{memberData[m.id]?.krs.length || 0} KRs</span>
                          <span style={{ textAlign: "right", fontFamily: F.mono, fontWeight: 700, color: STATUS_THEME[m.status].color }}>{m.rate.toFixed(1)}%</span>
                          <Bar value={m.rate} status={m.status} h={4} />
                          <div style={{ display: "flex", justifyContent: "flex-end" }}><Tag type={m.status} small /></div>
                        </div>
                      ))}
                    </Card>
                  );
                })}
              </>);
            })()}
          </Pane>
        </>)}

        {page === "setup" && (<>
          <Header title="OKR / KPI Setup" sub="Add, edit, or remove key results for each department and team" />
          <Pane>
            <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>{depts.map(d => <Btn key={d.id} primary={selDept === d.id} small onClick={() => { setSelDept(d.id); setSelTeam(null); setAddTarget(null); }}>{d.name}</Btn>)}</div>
            {selDept && (() => {
              const dept = depts.find(d => d.id === selDept); if (!dept) return null;
              const renderEditor = (krs, deptId, teamId) => (
                <Card style={{ overflow: "hidden" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "50px 1fr 100px 100px 50px", padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                    <span>ID</span><span>Key Result</span><span style={{ textAlign: "right" }}>Target</span><span style={{ textAlign: "right" }}>Actual</span><span></span>
                  </div>
                  {krs.map((kr, i) => (
                    <div key={kr.id} style={{ display: "grid", gridTemplateColumns: "50px 1fr 100px 100px 50px", padding: "9px 16px", gap: 8, alignItems: "center", background: i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, fontSize: 12 }}>
                      <span style={{ fontFamily: F.mono, fontSize: 10, color: T.textDim }}>{kr.id}</span>
                      <span>{kr.label}</span>
                      <Input value={kr.target} onChange={e => dispatch({ type: "UPDATE_KR", deptId, teamId, krId: kr.id, field: "target", value: Number(e.target.value) || 0 })} style={{ textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                      <span style={{ textAlign: "right", fontFamily: F.mono, color: T.textMuted }}>{fmt(kr.actual)}</span>
                      <button onClick={() => dispatch({ type: "REMOVE_KR", deptId, teamId, krId: kr.id })} style={{ background: T.badDim, border: `1px solid ${T.badBorder}`, borderRadius: 5, padding: "3px 8px", cursor: "pointer", color: T.bad, fontSize: 10, fontWeight: 700 }}>✕</button>
                    </div>
                  ))}
                  {addTarget === (teamId || `dept-${deptId}`) ? (
                    <div style={{ display: "grid", gridTemplateColumns: "50px 1fr 100px 100px 50px", padding: "9px 16px", gap: 8, alignItems: "center", background: T.brandDim }}>
                      <span style={{ fontSize: 10, color: T.brand }}>NEW</span>
                      <Input value={newKr.label} onChange={e => setNewKr(p => ({ ...p, label: e.target.value }))} placeholder="KR description" style={{ padding: "5px 8px", fontSize: 12 }} />
                      <Input value={newKr.target} onChange={e => setNewKr(p => ({ ...p, target: e.target.value }))} placeholder="Target" style={{ textAlign: "right", padding: "5px 8px", fontSize: 12, fontFamily: F.mono }} />
                      <span />
                      <button onClick={() => addKr(deptId, teamId)} style={{ background: T.brand, border: "none", borderRadius: 5, padding: "4px 8px", cursor: "pointer", color: "#fff", fontSize: 10, fontWeight: 700 }}>✓</button>
                    </div>
                  ) : (
                    <div style={{ padding: "10px 16px" }}>
                      <button onClick={() => { setAddTarget(teamId || `dept-${deptId}`); setNewKr({ label: "", target: "" }); }} style={{ background: "none", border: `1px dashed ${T.border}`, borderRadius: 6, padding: "8px 14px", cursor: "pointer", color: T.brand, fontSize: 11, fontWeight: 600, width: "100%", fontFamily: F.body }}>+ Add Key Result</button>
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

        {page === "reports" && (<>
          <Header title="Monthly KPI Reports" sub="Published reports visible to ALL teams across the company"
            right={<Btn primary onClick={() => {
              const report = { id: `mr${Date.now()}`, month: currentMonth(), publishedDate: new Date().toISOString().slice(0, 10), publishedBy: user.id,
                data: { companyRate: Math.round(compRate * 10) / 10, deptRanks: deptRanks.map(d => ({ name: d.name, rate: Math.round(d.rate * 10) / 10, status: d.status })),
                  topPerformers: allMembers.slice(0, 3).map(m => `${m.name} — ${m.rate.toFixed(1)}%`),
                  redFlags: allMembers.filter(m => m.status === "red").map(m => `${m.name} — ${m.rate.toFixed(1)}% (action required)`),
                },
              };
              dispatch({ type: "PUBLISH_REPORT", report });
            }}>Publish {currentMonth()} Report</Btn>} />
          <Pane>
            {state.monthlyReports.length === 0 && <EmptyState text="No monthly reports published yet." />}
            {state.monthlyReports.map(r => (
              <Card key={r.id} style={{ overflow: "hidden" }}>
                <div style={{ padding: "16px 20px", borderBottom: `1px solid ${T.border}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div><div style={{ fontSize: 16, fontWeight: 800 }}>{r.month}</div><div style={{ fontSize: 10, color: T.textMuted }}>Published: {r.publishedDate} · Visible to all teams</div></div>
                  <Tag type={getStatus(r.data.companyRate)} label={`Company: ${r.data.companyRate}%`} />
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

        {page === "sprints" && (<>
          <Header title="Manager Weekly Sprints" sub="Sprint reports submitted by team managers" />
          <Pane>
            {state.mgrSprints.length === 0 && <EmptyState text="No manager sprint reports yet." />}
            {state.mgrSprints.map(s => { const mgr = users.find(u => u.id === s.mgrId); return (
              <Card key={s.id} style={{ padding: "16px 20px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <Avatar letters={mgr?.av || "?"} size={28} />
                    <div><div style={{ fontSize: 13, fontWeight: 700 }}>{mgr?.name}</div><div style={{ fontSize: 10, color: T.textMuted }}>{mgr?.title}</div></div>
                  </div>
                  <div style={{ textAlign: "right" }}><div style={{ fontSize: 12, fontWeight: 700 }}>{s.week}</div><div style={{ fontSize: 10, color: T.textMuted }}>{s.date}</div></div>
                </div>
                <p style={{ margin: 0, fontSize: 12, color: T.textSoft, lineHeight: 1.7 }}>{s.summary}</p>
              </Card>
            ); })}
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
              <div style={{ display: "grid", gridTemplateColumns: "36px 32px 1fr 100px 90px 55px 150px 70px", padding: "7px 16px", gap: 8, borderBottom: `1px solid ${T.border}`, fontSize: 9, fontWeight: 700, color: T.textDim, letterSpacing: "0.07em", textTransform: "uppercase" }}>
                <span>#</span><span></span><span>Name</span><span>Dept</span><span>Team</span><span style={{ textAlign: "right" }}>Rate</span><span>Progress</span><span style={{ textAlign: "right" }}>Status</span>
              </div>
              {allMembers.map((m, i) => (
                <div key={m.id} style={{ display: "grid", gridTemplateColumns: "36px 32px 1fr 100px 90px 55px 150px 70px", padding: "10px 16px", gap: 8, alignItems: "center", background: i === 0 ? T.okDim : m.status === "red" ? T.badDim : i % 2 ? T.raised : "transparent", borderBottom: `1px solid ${T.border}`, borderLeft: i === 0 ? `3px solid ${T.ok}` : m.status === "red" ? `3px solid ${T.bad}` : "3px solid transparent", fontSize: 12 }}>
                  <span style={{ fontFamily: F.mono, fontWeight: 900, color: i === 0 ? T.ok : m.status === "red" ? T.bad : T.textMuted }}>#{i + 1}</span>
                  <Avatar letters={m.av} size={24} />
                  <div><span style={{ fontWeight: 600 }}>{m.name}</span><span style={{ color: T.textDim, marginLeft: 6, fontSize: 10 }}>{m.title}</span></div>
                  <span style={{ fontSize: 10, color: T.textMuted }}>{m.dept}</span>
                  <span style={{ fontSize: 10, color: T.textMuted }}>{m.team}</span>
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
  const [newSprint, setNewSprint] = useState({ week: currentWeekLabel(), summary: "" });
  const [newProj, setNewProj] = useState({ name: "", due: "" });

  const { depts, memberData, weeklySubs, mgrSprints, projects, monthlyReports, users } = state;
  const dept = depts.find(d => d.id === user.deptId);
  const myTeamMemberIds = dept?.teams.filter(t => user.teamIds?.includes(t.id)).flatMap(t => t.members) || [];
  const myMembers = users.filter(u => myTeamMemberIds.includes(u.id));
  const pendingSubs = weeklySubs.filter(s => myTeamMemberIds.includes(s.memberId) && s.approval === "pending");
  const mySprints = mgrSprints.filter(s => s.mgrId === user.id);
  const myProjects = projects.filter(p => p.mgrId === user.id);

  const navItems = [
    { id: "dashboard", icon: "⧉", label: "Team Dashboard"     },
    { id: "approvals", icon: "✓", label: "Approve Submissions" },
    { id: "sprints",   icon: "↻", label: "Weekly Sprint"       },
    { id: "projects",  icon: "⚡", label: "Projects"           },
    { id: "members",   icon: "✎", label: "Edit Member KPIs"   },
    { id: "reports",   icon: "⊞", label: "Monthly Reports"    },
  ];

  return (
    <div style={{ display: "flex", height: "100vh", fontFamily: F.body, background: T.bg, color: T.text }}>
      <Side items={navItems} active={page} onSelect={setPage} user={user} onLogout={onLogout} pendingCounts={{ approvals: pendingSubs.length }} />
      <div style={{ flex: 1, overflow: "auto" }}>

        {page === "dashboard" && (<>
          <Header title={`${dept?.name || "Team"} Dashboard`} sub={`${dept?.college} · Manager view`} />
          <Pane>
            <div style={{ display: "flex", gap: 14, flexWrap: "wrap" }}>
              <Metric label="Dept Completion"   value={`${calcRate(dept?.krs || []).toFixed(1)}%`} status={getStatus(calcRate(dept?.krs || []))} sub={`Time: ${TP}%`} />
              <Metric label="My Members"        value={myMembers.length} />
              <Metric label="Pending Approvals" value={pendingSubs.length} status={pendingSubs.length > 0 ? "yellow" : undefined} />
              <Metric label="Sprints Submitted" value={mySprints.length} />
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
                    <div style={{ textAlign: "right" }}><div style={{ fontSize: 12, fontWeight: 600 }}>{sub.week}</div><div style={{ fontSize: 10, color: T.textMuted }}>{sub.date}</div></div>
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

        {page === "sprints" && (<>
          <Header title="Weekly Sprint Report" sub="Submitted to company admin" />
          <Pane>
            <Card style={{ padding: 20 }}>
              <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 14 }}>New Sprint Report → Admin</div>
              <Input value={newSprint.week} onChange={e => setNewSprint(p => ({ ...p, week: e.target.value }))} style={{ width: 220, marginBottom: 10 }} />
              <TextArea value={newSprint.summary} onChange={e => setNewSprint(p => ({ ...p, summary: e.target.value }))} placeholder="Week summary: accomplishments, blockers, team performance notes, escalations..." rows={5} />
              <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 12 }}>
                <Btn primary disabled={!newSprint.summary} onClick={() => {
                  dispatch({ type: "ADD_MGR_SPRINT", sprint: { id: `ms${Date.now()}`, mgrId: user.id, week: newSprint.week, summary: newSprint.summary, date: new Date().toISOString().slice(0, 10), status: "submitted" } });
                  setNewSprint({ week: currentWeekLabel(), summary: "" });
                }}>Submit to Admin</Btn>
              </div>
            </Card>
            <SectionLabel>Previous Sprints</SectionLabel>
            {mySprints.map(s => (
              <Card key={s.id} style={{ padding: "14px 18px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
                  <span style={{ fontSize: 13, fontWeight: 700 }}>{s.week}</span>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}><span style={{ fontSize: 10, color: T.textMuted }}>{s.date}</span><Tag type="green" label="Submitted" small /></div>
                </div>
                <p style={{ margin: 0, fontSize: 12, color: T.textSoft, lineHeight: 1.6 }}>{s.summary}</p>
              </Card>
            ))}
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
            {myProjects.map(p => (
              <Card key={p.id} style={{ padding: "16px 18px" }}>
                <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 10 }}>
                  <div><div style={{ fontSize: 14, fontWeight: 700 }}>{p.name}</div><div style={{ fontSize: 10, color: T.textMuted }}>Due: {p.due}</div></div>
                  <Tag type={p.status === "active" ? "pending" : "approved"} label={p.status.toUpperCase()} small />
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <Bar value={p.progress} status={p.progress >= 60 ? "green" : p.progress >= 30 ? "yellow" : "red"} h={6} />
                  <span style={{ fontFamily: F.mono, fontSize: 12, fontWeight: 700, color: T.textSoft }}>{p.progress}%</span>
                </div>
              </Card>
            ))}
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
    case "APPROVE_SUB":     return { ...state, weeklySubs:  state.weeklySubs.map(s => s.id === action.subId ? { ...s, approval: action.status } : s) };
    case "ADD_MGR_SPRINT":  return { ...state, mgrSprints:  [action.sprint, ...state.mgrSprints]  };
    case "ADD_PROJECT":     return { ...state, projects:    [...state.projects, action.project]    };
    case "PUBLISH_REPORT":  return { ...state, monthlyReports: [action.report, ...state.monthlyReports] };

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
      const updated = { ...state.users.find(u => u.id === action.userId), ...action.updates };
      return { ...state, users: state.users.map(u => u.id === action.userId ? updated : u) };
    }

    case "REMOVE_USER":
      return { ...state, users: state.users.filter(u => u.id !== action.userId) };

    default: return state;
  }
}

/* ─────────────────────────────────────────────────────────────
   ROOT APP
   ───────────────────────────────────────────────────────────── */
export default function App() {
  const [user, setUser] = useState(null);
  const [msalErr, setMsalErr] = useState("");
  const { instance, accounts, inProgress } = useMsal();
  const [state, rawDispatch] = useState({
    users: INIT_USERS,
    depts: INIT_DEPTS,
    memberData: INIT_MEMBER_DATA,
    weeklySubs: INIT_WEEKLY_SUBS,
    mgrSprints: INIT_MGR_SPRINTS,
    projects: INIT_PROJECTS,
    monthlyReports: INIT_MONTHLY_REPORTS,
  });
  const dispatch = useCallback((action) => rawDispatch(prev => appReducer(prev, action)), []);

  const usersRef = useRef(state.users);
  useEffect(() => { usersRef.current = state.users; }, [state.users]);

  const routeByEmail = useCallback((email) => {
    if (!email) return;
    const lc = email.toLowerCase();
    const matched = usersRef.current.find(u => u.email.toLowerCase() === lc);
    if (matched) {
      setMsalErr("");
      setUser(matched);
    } else {
      setMsalErr(`No account found for ${lc}. Contact your administrator.`);
    }
  }, []);

  // Event-callback routing: fires immediately when MsalProvider processes the redirect
  useEffect(() => {
    const id = instance.addEventCallback((event) => {
      // LOGIN_SUCCESS payload IS the AccountInfo object (not a wrapper around it)
      if (event.eventType === EventType.LOGIN_SUCCESS && event.payload?.username) {
        routeByEmail(event.payload.username);
      } else if (event.eventType === EventType.ACQUIRE_TOKEN_SUCCESS && event.payload?.account?.username) {
        // ACQUIRE_TOKEN_SUCCESS payload is the full AuthenticationResult (.account exists)
        routeByEmail(event.payload.account.username);
      } else if (event.eventType === EventType.HANDLE_REDIRECT_END) {
        const all = instance.getAllAccounts();
        if (all.length > 0) routeByEmail(all[0].username);
      }
    });
    return () => { if (id) instance.removeEventCallback(id); };
  }, [instance, routeByEmail]);

  // Reactive accounts watch: fires whenever MsalProvider's accounts state updates
  useEffect(() => {
    if (user || inProgress !== InteractionStatus.None || accounts.length === 0) return;
    routeByEmail(accounts[0].username);
  }, [accounts, inProgress, user]); // eslint-disable-line

  // Mount / inProgress-settled check: routes user if account already in cache
  useEffect(() => {
    if (user || inProgress !== InteractionStatus.None) return;
    const all = instance.getAllAccounts();
    if (all.length > 0) routeByEmail(all[0].username);
  }, [inProgress]); // eslint-disable-line

  if (!user) return <LoginPage onLogin={setUser} users={state.users} inProgress={inProgress} msalErr={msalErr} />;
  const logout = () => {
    setUser(null);
    setMsalErr("");
    try { instance.clearCache(); } catch (_) {}
  };

  if (user.role === "admin")   return <AdminPortal   user={user} onLogout={logout} state={state} dispatch={dispatch} />;
  if (user.role === "manager") return <ManagerPortal user={user} onLogout={logout} state={state} dispatch={dispatch} />;
  return <MemberPortal user={user} onLogout={logout} state={state} dispatch={dispatch} />;
}
