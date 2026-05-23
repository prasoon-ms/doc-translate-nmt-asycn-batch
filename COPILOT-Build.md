# 📄 React Document Translator App

A React-based web application that translates documents (PDF, DOCX, PPTX, etc.) using the Azure AI Translator **Document Translation API (GA `2024-05-01`)** in Neural Machine Translation (NMT) mode. Supports glossary integration and secure authentication via Microsoft Entra ID with managed identity. Translated documents are stored in Azure Blob Storage while preserving original formatting.

---

## 🚀 Features

- Translate documents in bulk using Azure Translator’s batch API (NMT only)
- Supports PDF, DOCX, PPTX, XLSX, TXT, HTML, and more
- Upload and apply custom glossaries (TSV/XLIFF)
- Monitor translation job status asynchronously
- Secure authentication using Microsoft Entra ID + managed identity
- Deployable to Azure Static Web Apps or Azure App Service

---

## 🧱 Architecture Overview

```plaintext
User (React SPA)
     |
     v
Azure Static Web Apps (React Frontend)
     |
     v
Backend API (Node.js / Python - Azure Functions or App Service)
     |
     v
Azure Translator Document Translation API (api-version=2024-05-01)
     |
     v
Azure Blob Storage
  - source/ (input documents)
  - target/ (translated documents)
  - glossaries/ (TSV/XLIFF files)
```

> Note: The React frontend should not call Azure Translator directly. Use a backend API (Azure Functions or App Service) that authenticates to Azure Translator and Blob Storage using a **managed identity** (system-assigned or user-assigned). Do not use SAS tokens or account keys.

---

## 🔐 Authentication (Microsoft Entra ID + Managed Identity)

Two independent identities are used:

- **End-user identity** — The React SPA signs the user in with MSAL.js and acquires an access token for the backend API's scope.
- **Backend workload identity** — A **managed identity** on the Azure Functions / App Service backend calls Azure Translator and Blob Storage.

**No SAS tokens, account keys, or client secrets are used anywhere.**

### App registrations (when *not* using SWA built-in auth)

1. **SPA app** in Microsoft Entra ID — SPA platform, redirect URI = SWA origin, consents to the API scope below.
2. **API app** in Microsoft Entra ID — exposes a scope, e.g. `api://<api-client-id>/Translate.Submit`.

> **Simpler alternative:** Deploy the React app to **Azure Static Web Apps** with the **linked Azure Functions** backend and use SWA's built-in authentication. The user principal is forwarded to the backend in the `x-ms-client-principal` header — you can skip MSAL.js and both app registrations entirely. Pick **one** model and stick with it; do not mix.

### RBAC for the backend managed identity

Enable a managed identity on the backend (system-assigned for simplicity, user-assigned for portability) and assign:

- **Cognitive Services User** on the Translator resource (data-plane access via AAD tokens).
- **Storage Blob Data Contributor** on the Storage account (upload source files, read results).

### RBAC for the Translator resource's own managed identity

The Translator service must read inputs and write outputs in your Storage account. Enable a **system-assigned managed identity on the Translator resource** and assign it **Storage Blob Data Contributor** on the Storage account. With this in place, batch payloads use plain blob/container URLs — **no SAS**.

### Frontend sign-in (MSAL.js path)

```tsx
import { PublicClientApplication } from "@azure/msal-browser";
import { MsalProvider, useMsal } from "@azure/msal-react";

const pca = new PublicClientApplication({
  auth: {
    clientId: import.meta.env.VITE_SPA_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${import.meta.env.VITE_TENANT_ID}`,
    redirectUri: window.location.origin,
  },
});

const API_SCOPE = `api://${import.meta.env.VITE_API_CLIENT_ID}/Translate.Submit`;

export function Root({ children }) {
  return <MsalProvider instance={pca}>{children}</MsalProvider>;
}

export function SignInButton() {
  const { instance } = useMsal();
  return (
    <button onClick={() => instance.loginPopup({ scopes: [API_SCOPE] })}>
      Sign in
    </button>
  );
}

// Attach the token to every API call:
export async function apiFetch(path: string, init: RequestInit = {}) {
  const { instance, accounts } = (window as any).__msal; // wire from useMsal in real code
  const result = await instance.acquireTokenSilent({ scopes: [API_SCOPE], account: accounts[0] });
  return fetch(path, {
    ...init,
    headers: { ...init.headers, Authorization: `Bearer ${result.accessToken}` },
  });
}
```

### Backend: validate the SPA token (MSAL.js path)

The backend **must** validate the incoming bearer token before calling Azure on the user's behalf. Verify `aud`, `iss`, `tid`, and the expected `scp` claim (`Translate.Submit`).

```ts
import jwksClient from "jwks-rsa";
import jwt from "jsonwebtoken";

const TENANT_ID = process.env.AZURE_TENANT_ID!;
const API_CLIENT_ID = process.env.API_CLIENT_ID!;
const REQUIRED_SCOPE = "Translate.Submit";

const jwks = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`,
});

function getKey(header: jwt.JwtHeader, cb: jwt.SigningKeyCallback) {
  jwks.getSigningKey(header.kid!, (err, key) => cb(err, key?.getPublicKey()));
}

export function requireUser(req, res, next) {
  const auth = req.headers.authorization ?? "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : null;
  if (!token) return res.status(401).end();

  jwt.verify(
    token,
    getKey,
    {
      audience: `api://${API_CLIENT_ID}`,
      issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
      algorithms: ["RS256"],
    },
    (err, payload: any) => {
      if (err) return res.status(401).end();
      const scopes = (payload.scp ?? "").split(" ");
      if (!scopes.includes(REQUIRED_SCOPE)) return res.status(403).end();
      req.user = payload;
      next();
    }
  );
}
```

### Backend: SWA built-in auth path (simpler)

If you chose SWA built-in auth, the user principal arrives base64-encoded in `x-ms-client-principal`. No token verification code is needed — SWA already validated it:

```ts
export function getUser(req) {
  const header = req.headers["x-ms-client-principal"];
  if (!header) return null;
  return JSON.parse(Buffer.from(header as string, "base64").toString("utf8"));
}
```

### Backend: call Azure with managed identity

```ts
import { DefaultAzureCredential, ManagedIdentityCredential } from "@azure/identity";

// In Azure, DefaultAzureCredential uses the managed identity automatically.
// For user-assigned MI, pass the clientId explicitly to avoid env-var ambiguity.
const credential = process.env.USER_ASSIGNED_MI_CLIENT_ID
  ? new ManagedIdentityCredential({ clientId: process.env.USER_ASSIGNED_MI_CLIENT_ID })
  : new DefaultAzureCredential();

// Document Translation uses the resource's regional custom subdomain endpoint
// (e.g. https://<your-translator>.cognitiveservices.azure.com). With AAD bearer
// tokens the only required header is Authorization. Ocp-Apim-Subscription-Region
// and subscription keys are NOT used in this path.
const token = await credential.getToken("https://cognitiveservices.azure.com/.default");
const headers = { Authorization: `Bearer ${token!.token}` };
```

### Backend: blob storage with managed identity

```ts
import { BlobServiceClient } from "@azure/storage-blob";

const blobService = new BlobServiceClient(
  `https://${process.env.STORAGE_ACCOUNT_NAME}.blob.core.windows.net`,
  credential
);

export async function uploadSourceBlob(jobId: string, file: { name: string; buffer: Buffer; contentType: string }) {
  const container = blobService.getContainerClient(`source-${jobId}`);
  await container.createIfNotExists();
  await container.getBlockBlobClient(file.name).uploadData(file.buffer, {
    blobHTTPHeaders: { blobContentType: file.contentType },
  });
  return container.url; // plain URL, no SAS
}
```

### CORS

- **SWA-linked Functions:** the SPA and API share an origin (`/api/*` is reverse-proxied). **No CORS configuration needed.** This is the recommended setup.
- **Separate App Service / Functions backend:** add the SWA origin to the backend's CORS allow-list via `az functionapp cors add` / `az webapp cors add`. Do **not** use `*` when the API requires auth.

---

## 📁 Supported Document Formats

| Format Type | Extensions |
|-------------|------------|
| Word        | `.docx`    |
| PowerPoint  | `.pptx`    |
| Excel       | `.xlsx`    |
| PDF         | `.pdf`     |
| Text        | `.txt`, `.html`, `.md` |
| Images (Preview) | `.jpeg`, `.png`, `.bmp`, `.webp` |

---

## 📚 Glossary Integration

Create a glossary file (TSV or XLIFF):

```tsv
Bank    Banque
Card    Carte
Crane   Grue
```

Upload to Blob Storage (using the managed identity from the backend) and reference it in the translation request. Because the **Translator resource's own managed identity** holds *Storage Blob Data Contributor* on the storage account, blob URLs are submitted **without SAS**:

```json
"glossaries": [
  {
    "glossaryUrl": "https://<storage>.blob.core.windows.net/glossaries/en-fr.tsv",
    "format": "tsv"
  }
]
```

---

## 🔄 Translation Pipeline

The React app posts file uploads to the backend's API contract (below). The backend uploads each file to Blob Storage with its managed identity, then submits a batch job to:

```
POST https://<your-translator>.cognitiveservices.azure.com/translator/document/batches?api-version=2024-05-01
Authorization: Bearer <AAD token for cognitiveservices.azure.com/.default>
Content-Type: application/json
```

The Translator service accesses the blobs via its **own managed identity** — URLs contain no SAS:

```json
{
  "inputs": [
    {
      "source": {
        "sourceUrl": "https://<storage>.blob.core.windows.net/source-<jobId>/",
        "language": "en",
        "storageSource": "AzureBlob"
      },
      "targets": [
        {
          "targetUrl": "https://<storage>.blob.core.windows.net/target-<jobId>/",
          "language": "fr",
          "category": "general",
          "glossaries": [
            {
              "glossaryUrl": "https://<storage>.blob.core.windows.net/glossaries/en-fr.tsv",
              "format": "tsv"
            }
          ]
        }
      ]
    }
  ]
}
```

A successful submission returns HTTP `202 Accepted` with the job URL in the `Operation-Location` response header; the last URL segment is the job id.

### Backend HTTP contract (frontend ↔ backend)

All endpoints require a valid bearer token (MSAL.js path) or a SWA-authenticated user (SWA path).

| Method | Path                  | Request                                                                                       | Response                                              |
|--------|-----------------------|-----------------------------------------------------------------------------------------------|-------------------------------------------------------|
| `POST` | `/api/translate`      | `multipart/form-data`: `documents` (1..n files), `sourceLanguage`, `targetLanguage`, optional `glossaryId` | `202` `{ "jobId": "<guid>", "status": "NotStarted" }` |
| `GET`  | `/api/jobs`           | —                                                                                             | `200` `{ "jobs": [{ jobId, status, createdAt }] }`    |
| `GET`  | `/api/jobs/{jobId}`   | —                                                                                             | `200` `{ jobId, status, summary, documents: [...] }`  |
| `DELETE` | `/api/jobs/{jobId}` | —                                                                                             | `204`                                                 |

**Upload limits enforced server-side:**
- Max 50 files per request, max 40 MB per file, max 250 MB per request.
- Allowed extensions match the [Supported Document Formats](#-supported-document-formats) table; reject everything else with `415`.

### Example React upload component

```tsx
import { useState } from "react";
import { apiFetch } from "./auth";

export function UploadForm() {
  const [error, setError] = useState<string | null>(null);

  async function handleSubmit(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();
    setError(null);
    const form = new FormData(e.currentTarget);
    const res = await apiFetch("/api/translate", { method: "POST", body: form });
    if (!res.ok) {
      setError(`Submit failed (${res.status})`);
      return;
    }
    const job = await res.json();
    console.log("Job id:", job.jobId);
  }

  return (
    <form onSubmit={handleSubmit}>
      <input
        type="file"
        name="documents"
        multiple
        required
        accept=".docx,.pptx,.xlsx,.pdf,.txt,.html,.md"
      />
      <select name="sourceLanguage" defaultValue="en">
        <option value="en">English</option>
      </select>
      <select name="targetLanguage" defaultValue="fr">
        <option value="fr">French</option>
        <option value="es">Spanish</option>
        <option value="de">German</option>
      </select>
      <button type="submit">Translate</button>
      {error && <p role="alert">{error}</p>}
    </form>
  );
}
```

---

## 📊 Monitor Translation Jobs

The React UI polls the backend, which proxies to the Translator endpoints (api-version=2024-05-01):

- `GET /translator/document/batches` – list all jobs
- `GET /translator/document/batches/{id}` – job status
- `GET /translator/document/batches/{id}/documents` – document statuses
- `DELETE /translator/document/batches/{id}` – cancel job

Example polling hook:

```tsx
import { useEffect, useState } from "react";

export function useJobStatus(jobId: string) {
  const [status, setStatus] = useState<string>("pending");
  useEffect(() => {
    const t = setInterval(async () => {
      const res = await fetch(`/api/jobs/${jobId}`);
      const data = await res.json();
      setStatus(data.status);
      if (["Succeeded", "Failed", "Cancelled"].includes(data.status)) clearInterval(t);
    }, 3000);
    return () => clearInterval(t);
  }, [jobId]);
  return status;
}
```

---

## ☁️ Deployment to Azure

**Frontend (React):**
1. Build the app: `npm run build`
2. Deploy the `dist/` folder (Vite) to **Azure Static Web Apps**.
3. Configure `staticwebapp.config.json` with route rules, auth provider, and (for SWA-linked Functions) the backend is automatically exposed at `/api/*` — no CORS required.

**Backend (API):**
1. Deploy to **Azure Functions** (linked via SWA — recommended) or standalone **Azure App Service**.
2. Enable a **managed identity** on the resource and assign RBAC roles:
   - *Cognitive Services User* on the Translator resource
   - *Storage Blob Data Contributor* on the Storage account
3. Also enable a **system-assigned managed identity on the Translator resource** and grant it *Storage Blob Data Contributor* on the Storage account (required so Translator can read/write blobs without SAS).
4. Configure environment variables (no secrets):
   - `TRANSLATOR_ENDPOINT` — e.g. `https://<your-translator>.cognitiveservices.azure.com`
   - `TRANSLATOR_API_VERSION` — `2024-05-01`
   - `STORAGE_ACCOUNT_NAME`
   - `AZURE_TENANT_ID`, `API_CLIENT_ID` — *only* if validating MSAL tokens (skip if using SWA built-in auth)
   - `USER_ASSIGNED_MI_CLIENT_ID` — *only* if using a user-assigned managed identity

---

## 📦 Requirements

**Frontend** (React 18+, Vite):

```bash
npm create vite@latest doc-translate -- --template react-ts
cd doc-translate
npm install react react-dom @azure/msal-browser @azure/msal-react
```

**Backend** (Node.js 18+ — uses built-in `fetch`):

```bash
npm install @azure/identity @azure/storage-blob express multer jsonwebtoken jwks-rsa
```

> Skip `jsonwebtoken` + `jwks-rsa` if you use Static Web Apps built-in auth (read `x-ms-client-principal` instead).

---

## 📌 Notes

- Uses Azure AI Translator **Document Translation API `2024-05-01`** (GA), NMT only.
- AAD bearer tokens are used for data-plane Translator calls — do **not** send `Ocp-Apim-Subscription-Key` or `Ocp-Apim-Subscription-Region` on this path.
- Glossary terms are case-sensitive by default.
- Never embed Azure secrets in the React bundle — always proxy through the backend.
- **Do not use SAS tokens or storage account keys.** All Azure access is via managed identity + Azure RBAC.
- Always validate the incoming user token on the backend (or rely on Static Web Apps built-in auth); never trust the caller because they reached `/api/*`.
