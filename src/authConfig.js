import { PublicClientApplication } from "@azure/msal-browser";

export const msalConfig = {
  auth: {
    clientId: import.meta.env.VITE_AZURE_CLIENT_ID || "YOUR_CLIENT_ID",
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_AZURE_TENANT_ID || "common"}`,
    redirectUri: window.location.origin,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: true,
  },
};

export const loginRequest = {
  scopes: ["User.Read"],
  prompt: "select_account",
};

export const msalInstance = new PublicClientApplication(msalConfig);

// Our API verifies the caller's Microsoft ID token, so every /api request carries it.
async function getIdToken() {
  const account = msalInstance.getActiveAccount() || msalInstance.getAllAccounts()[0];
  if (!account) throw new Error("You are signed out. Refresh the page and sign in again.");
  const request = { scopes: loginRequest.scopes, account };
  let result = await msalInstance.acquireTokenSilent(request);
  if (!result.idTokenClaims?.exp || result.idTokenClaims.exp * 1000 < Date.now() + 60000) {
    result = await msalInstance.acquireTokenSilent({ ...request, forceRefresh: true });
  }
  return result.idToken;
}

export async function apiFetch(path, init = {}) {
  const idToken = await getIdToken();
  return fetch(path, { ...init, headers: { ...(init.headers || {}), Authorization: `Bearer ${idToken}` } });
}
