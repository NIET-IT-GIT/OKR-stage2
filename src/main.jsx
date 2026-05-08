import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App.jsx";

const msalInstance = new PublicClientApplication(msalConfig);

msalInstance.initialize().then(async () => {
  // Process any pending redirect before rendering so we have the account
  // synchronously available as a prop — no event-callback timing issues.
  const redirectResult = await msalInstance.handleRedirectPromise().catch(() => null);

  createRoot(document.getElementById("root")).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App redirectAccount={redirectResult?.account ?? null} />
      </MsalProvider>
    </StrictMode>
  );
}).catch(console.error);
