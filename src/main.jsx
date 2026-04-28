import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App.jsx";

const msalInstance = new PublicClientApplication(msalConfig);

// MSAL v3+ requires initialize() before any auth operations.
// We also process handleRedirectPromise here so the redirect-based
// auth code in the URL is consumed before React renders.
(async () => {
  await msalInstance.initialize();
  // handleRedirectPromise throws no_token_request_cache_error on normal
  // loads when there is no pending redirect — ignore it and always render.
  await msalInstance.handleRedirectPromise().catch(() => null);
  createRoot(document.getElementById("root")).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>
  );
})().catch(console.error);
