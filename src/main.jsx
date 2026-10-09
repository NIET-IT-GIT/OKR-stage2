import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import { msalInstance } from "./authConfig";
import App from "./App.jsx";

msalInstance.initialize().then(async () => {
  // Process any pending redirect before rendering so we have the account
  // synchronously available as a prop — no event-callback timing issues.
  const redirectResult = await msalInstance.handleRedirectPromise().catch(() => null);
  if (redirectResult?.account) msalInstance.setActiveAccount(redirectResult.account);

  createRoot(document.getElementById("root")).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App redirectAccount={redirectResult?.account ?? null} />
      </MsalProvider>
    </StrictMode>
  );
}).catch(console.error);
