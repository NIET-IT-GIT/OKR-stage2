import { createRemoteJWKSet, jwtVerify } from "jose";
import { createClient } from "@supabase/supabase-js";

const TENANT = process.env.AZURE_TENANT_ID || process.env.VITE_AZURE_TENANT_ID || "";
const CLIENT_ID = process.env.AZURE_CLIENT_ID || process.env.VITE_AZURE_CLIENT_ID || "";
const SUPABASE_URL = process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL || "";
const SUPABASE_KEY = process.env.SUPABASE_ANON_KEY || process.env.VITE_SUPABASE_ANON_KEY || "";

export class AuthError extends Error {
  constructor(status, message) { super(message); this.status = status; }
}

// The tenant may be configured as a GUID or a domain, so read the real issuer from Microsoft's metadata.
let oidcPromise = null;
function getOidc() {
  if (!oidcPromise) {
    oidcPromise = fetch(`https://login.microsoftonline.com/${encodeURIComponent(TENANT)}/v2.0/.well-known/openid-configuration`)
      .then(r => { if (!r.ok) throw new Error(`OpenID metadata HTTP ${r.status}`); return r.json(); })
      .then(cfg => ({ issuer: cfg.issuer, jwks: createRemoteJWKSet(new URL(cfg.jwks_uri)) }))
      .catch(err => { oidcPromise = null; throw err; });
  }
  return oidcPromise;
}

// Verifies the caller's Microsoft ID token and returns their record from the users collection.
export async function requireUser(req, { roles } = {}) {
  if (!TENANT || TENANT === "common" || !CLIENT_ID || !SUPABASE_URL || !SUPABASE_KEY) {
    console.error("[auth] AZURE_TENANT_ID / AZURE_CLIENT_ID / SUPABASE_URL / SUPABASE_ANON_KEY not configured");
    throw new AuthError(500, "Sign-in checking is not configured on the server.");
  }
  const m = /^Bearer\s+(.+)$/i.exec(req.headers.authorization || "");
  if (!m) throw new AuthError(401, "Sign in required.");

  let payload;
  try {
    const { issuer, jwks } = await getOidc();
    ({ payload } = await jwtVerify(m[1], jwks, { issuer, audience: CLIENT_ID }));
  } catch (err) {
    console.warn("[auth] token rejected:", err.code || err.message);
    throw new AuthError(401, "Your sign-in has expired. Refresh the page and try again.");
  }

  const email = String(payload.preferred_username || payload.email || "").trim().toLowerCase();
  if (!email) throw new AuthError(401, "Your sign-in token has no email address.");

  const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, { auth: { persistSession: false } });
  const { data, error } = await supabase.from("app_data").select("doc").eq("collection", "users");
  if (error) {
    console.error("[auth] user lookup failed:", error.message);
    throw new AuthError(500, "Could not check your account. Try again shortly.");
  }
  const user = (data || []).map(r => r.doc).find(u => String(u.email || "").trim().toLowerCase() === email);
  if (!user || user.active === false) throw new AuthError(403, "Your account is not registered in this system.");
  if (roles && !roles.includes(user.role)) throw new AuthError(403, "You don't have permission to do this.");
  return user;
}

export function sendAuthError(res, err) {
  if (err instanceof AuthError) return res.status(err.status).json({ error: err.message });
  console.error("[auth] unexpected error:", err);
  return res.status(500).json({ error: "Something went wrong. Try again shortly." });
}

export function esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
