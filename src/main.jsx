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
msalInstance.initialize().then(() => {
  return msalInstance.handleRedirectPromise();
}).then(() => {
  createRoot(document.getElementById("root")).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>
  );
}).catch(console.error);
