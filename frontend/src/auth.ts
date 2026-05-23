import { PublicClientApplication, type AccountInfo } from "@azure/msal-browser";

export const API_SCOPE = `api://${import.meta.env.VITE_API_CLIENT_ID}/Translate.Submit`;

export const pca = new PublicClientApplication({
  auth: {
    clientId: import.meta.env.VITE_SPA_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID}`,
    redirectUri: window.location.origin,
  },
  cache: { cacheLocation: "sessionStorage" },
});

async function acquireToken(): Promise<string> {
  const account: AccountInfo | undefined = pca.getAllAccounts()[0];
  if (!account) throw new Error("Not signed in");
  const result = await pca.acquireTokenSilent({ scopes: [API_SCOPE], account });
  return result.accessToken;
}

export async function apiFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const token = await acquireToken();
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  return fetch(path, { ...init, headers });
}
