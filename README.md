# Document Translator (React + Express)

Implementation of [COPILOT-Build.md](./COPILOT-Build.md): React SPA + Express backend that runs Azure AI Translator **Document Translation API `2024-05-01`** in NMT mode. Auth is MSAL.js (SPA) + AAD bearer token (backend). Azure access is via **managed identity only** — no SAS, no account keys.

```
frontend/   Vite + React + TypeScript + MSAL.js
backend/    Express + TypeScript + @azure/identity + @azure/storage-blob
```

---

## Prerequisites

- Node.js **18+** (uses built-in `fetch`)
- An Azure AI Translator resource (with a custom subdomain enabled)
- An Azure Storage account
- Azure CLI (`az`) signed in to the target tenant: `az login --tenant <TENANT_ID>`
- For local dev, `DefaultAzureCredential` uses your `az` login identity

## RBAC (required before the backend can call Azure)

| Identity                                         | Scope                 | Role                              |
|--------------------------------------------------|-----------------------|-----------------------------------|
| Backend managed identity (or your dev user)      | Translator resource   | Cognitive Services User           |
| Backend managed identity (or your dev user)      | Storage account       | Storage Blob Data Contributor     |
| **Translator resource's system-assigned MI**     | Storage account       | Storage Blob Data Contributor     |

The last row is what lets Translator read inputs / write outputs **without SAS**.

---

## 1. Create the two Microsoft Entra app registrations

You need two app registrations:

1. **API app** — represents the backend; exposes the `Translate.Submit` scope; issues **v2** access tokens.
2. **SPA app** — the React frontend; signs the user in and requests the API scope.

The script below creates both, configures the SPA redirect URI, exposes the scope on the API, pre-authorizes the SPA, and sets `requestedAccessTokenVersion = 2`. Run it once in **PowerShell** from any folder.

```powershell
# --- inputs ---
$tenantId        = (az account show --query tenantId -o tsv)
$apiDisplayName  = "doc-translate-api"
$spaDisplayName  = "doc-translate-spa"
$spaRedirectUri  = "http://localhost:5173"
$scopeName       = "Translate.Submit"

# --- 1a. API app ---
$apiAppId    = az ad app create --display-name $apiDisplayName --sign-in-audience AzureADMyOrg --query appId -o tsv
$apiObjectId = az ad app show --id $apiAppId --query id -o tsv

# Service principal is required so SPA can request tokens for this resource
az ad sp create --id $apiAppId | Out-Null

# Set Application ID URI = api://<apiAppId>
az ad app update --id $apiAppId --identifier-uris "api://$apiAppId" | Out-Null

# Expose scope Translate.Submit AND opt into v2 tokens (via Graph PATCH)
$scopeId = [guid]::NewGuid().ToString()
$apiBody = @{
  api = @{
    requestedAccessTokenVersion = 2
    oauth2PermissionScopes = @(
      @{
        id                      = $scopeId
        adminConsentDescription = "Allow the app to submit document translation jobs"
        adminConsentDisplayName = "Submit translation jobs"
        userConsentDescription  = "Allow the app to submit document translation jobs on your behalf"
        userConsentDisplayName  = "Submit translation jobs"
        value                   = $scopeName
        type                    = "User"
        isEnabled               = $true
      }
    )
  }
} | ConvertTo-Json -Depth 6
$tmp = New-TemporaryFile; Set-Content $tmp $apiBody -Encoding utf8
az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$apiObjectId" `
  --headers "Content-Type=application/json" `
  --body "@$tmp" | Out-Null
Remove-Item $tmp

# --- 1b. SPA app ---
$spaAppId    = az ad app create --display-name $spaDisplayName --sign-in-audience AzureADMyOrg --query appId -o tsv
$spaObjectId = az ad app show --id $spaAppId --query id -o tsv
az ad sp create --id $spaAppId | Out-Null

# Register as SPA platform (NOT Web) and add the localhost redirect
$spaBody = @{
  spa = @{ redirectUris = @($spaRedirectUri) }
  web = @{ redirectUris = @() }
} | ConvertTo-Json -Depth 4
$tmp = New-TemporaryFile; Set-Content $tmp $spaBody -Encoding utf8
az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$spaObjectId" `
  --headers "Content-Type=application/json" `
  --body "@$tmp" | Out-Null
Remove-Item $tmp

# Grant the SPA the Translate.Submit delegated permission on the API
az ad app permission add --id $spaAppId `
  --api $apiAppId `
  --api-permissions "$scopeId=Scope" | Out-Null

# Pre-authorize the SPA on the API (skips the consent prompt)
$preauthBody = @{
  api = @{
    preAuthorizedApplications = @(
      @{ appId = $spaAppId; delegatedPermissionIds = @($scopeId) }
    )
  }
} | ConvertTo-Json -Depth 6
$tmp = New-TemporaryFile; Set-Content $tmp $preauthBody -Encoding utf8
az rest --method PATCH `
  --uri "https://graph.microsoft.com/v1.0/applications/$apiObjectId" `
  --headers "Content-Type=application/json" `
  --body "@$tmp" | Out-Null
Remove-Item $tmp

# --- output the values you need in .env files ---
Write-Host ""
Write-Host "TENANT_ID      = $tenantId"
Write-Host "API_CLIENT_ID  = $apiAppId"
Write-Host "SPA_CLIENT_ID  = $spaAppId"
Write-Host "API_SCOPE      = api://$apiAppId/$scopeName"
```

### Verify the API app is configured for v2 tokens

```powershell
az ad app show --id $apiAppId --query "api.requestedAccessTokenVersion"   # must print 2
az ad app show --id $apiAppId --query "api.oauth2PermissionScopes[].value" # must include "Translate.Submit"
```

If `requestedAccessTokenVersion` is `null` or `1`, the backend (which validates the v2 issuer `https://login.microsoftonline.com/<tenant>/v2.0`) will reject every token with `401 invalid_token`.

### Add additional redirect URIs later (e.g., for App Service / SWA)

```powershell
$spaBody = @{ spa = @{ redirectUris = @(
  "http://localhost:5173",
  "https://<your-app>.azurewebsites.net"
) } } | ConvertTo-Json -Depth 4
$tmp = New-TemporaryFile; Set-Content $tmp $spaBody -Encoding utf8
az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/$spaObjectId" `
  --headers "Content-Type=application/json" --body "@$tmp"
Remove-Item $tmp
```

---

## 2. Configure environment variables

```powershell
# Backend
cd backend
Copy-Item .env.example .env
```

Edit `backend/.env`:

```
PORT=8080
AZURE_TENANT_ID=<TENANT_ID>
API_CLIENT_ID=<API_CLIENT_ID>
REQUIRED_SCOPE=Translate.Submit
TRANSLATOR_ENDPOINT=https://<your-translator-name>.cognitiveservices.azure.com
TRANSLATOR_API_VERSION=2024-05-01
STORAGE_ACCOUNT_NAME=<your-storage-account-name>
# USER_ASSIGNED_MI_CLIENT_ID=   # only set when running on Azure with a user-assigned MI
```

```powershell
# Frontend
cd ..\frontend
Copy-Item .env.example .env
```

Edit `frontend/.env`:

```
VITE_TENANT_ID=<TENANT_ID>
VITE_SPA_CLIENT_ID=<SPA_CLIENT_ID>
VITE_API_CLIENT_ID=<API_CLIENT_ID>
VITE_API_PROXY=http://localhost:8080
```

---

## 3. Install and run

```powershell
cd backend  ; npm install
cd ..\frontend ; npm install
```

```powershell
# Terminal 1
cd backend
npm run dev      # http://localhost:8080

# Terminal 2
cd frontend
npm run dev      # http://localhost:5173 (proxies /api to backend)
```

Sign in with a user from your tenant, pick source/target languages (full GA list is fetched from `/api/languages`), upload documents, watch jobs in the list, expand a `Succeeded` job and click **Download** for each translated file.

## Type-check / build

```powershell
cd backend  ; npm run typecheck ; npm run build
cd frontend ; npm run typecheck ; npm run build
```

---

## Backend HTTP contract

All endpoints under `/api/jobs` and `/api/translate` require `Authorization: Bearer <MSAL access token>` with the `Translate.Submit` scope. `/api/languages` is public (it just proxies the Translator public language list and caches it).

| Method   | Path                                              | Body / params                                                                  | Response |
|----------|---------------------------------------------------|--------------------------------------------------------------------------------|----------|
| `GET`    | `/api/languages`                                  | —                                                                              | `200 { languages: [{ code, name, nativeName, dir }] }` |
| `POST`   | `/api/translate`                                  | `multipart/form-data`: `documents` (1..50), `sourceLanguage`, `targetLanguage` | `202 { jobId, localJobId, status }` |
| `GET`    | `/api/jobs`                                       | —                                                                              | `200 { jobs: [{ jobId, status, createdAt, summary }] }` |
| `GET`    | `/api/jobs/:id`                                   | —                                                                              | `200 { jobId, summary, documents: [{ id, status, to, fileName, progress, downloadUrl }] }` |
| `GET`    | `/api/jobs/:id/documents/:docId/download`         | —                                                                              | `200` translated file stream (`Content-Disposition: attachment`) |
| `DELETE` | `/api/jobs/:id`                                   | —                                                                              | `204` |

Limits enforced server-side: 50 files / 40 MB per file / 250 MB per request, extensions limited to `.docx .pptx .xlsx .pdf .txt .html .md`.

