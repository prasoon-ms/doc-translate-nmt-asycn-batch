import React from "react";
import ReactDOM from "react-dom/client";
import { MsalProvider } from "@azure/msal-react";
import { pca } from "./auth";
import { App } from "./App";

pca.initialize()
  .then(() => pca.handleRedirectPromise())
  .then(() => {
    const accounts = pca.getAllAccounts();
    if (accounts.length > 0) pca.setActiveAccount(accounts[0]);
    ReactDOM.createRoot(document.getElementById("root")!).render(
      <React.StrictMode>
        <MsalProvider instance={pca}>
          <App />
        </MsalProvider>
      </React.StrictMode>
    );
  });
