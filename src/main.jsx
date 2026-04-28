import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";
import App from "./App.jsx";

const msalInstance = new PublicClientApplication(msalConfig);

// initialize() is required for msal-browser v3+ before any auth operation.
// Do NOT call handleRedirectPromise() here — MsalProvider calls it internally
// and fires the LOGIN_SUCCESS / HANDLE_REDIRECT_END events that update
// useMsal()'s accounts state. Calling it early consumes the auth code
// before the provider can process it, so accounts never updates.
msalInstance.initialize().then(() => {
  createRoot(document.getElementById("root")).render(
    <StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </StrictMode>
  );
}).catch(console.error);
