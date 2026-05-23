import { DefaultAzureCredential, ManagedIdentityCredential, type TokenCredential } from "@azure/identity";
import { BlobServiceClient } from "@azure/storage-blob";

function requireEnv(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required env var: ${name}`);
  return v;
}

export const STORAGE_ACCOUNT_NAME = requireEnv("STORAGE_ACCOUNT_NAME");
export const TRANSLATOR_ENDPOINT = requireEnv("TRANSLATOR_ENDPOINT").replace(/\/$/, "");
export const TRANSLATOR_API_VERSION = process.env.TRANSLATOR_API_VERSION ?? "2024-05-01";

export const credential: TokenCredential = process.env.USER_ASSIGNED_MI_CLIENT_ID
  ? new ManagedIdentityCredential({ clientId: process.env.USER_ASSIGNED_MI_CLIENT_ID })
  : new DefaultAzureCredential();

export const blobService = new BlobServiceClient(
  `https://${STORAGE_ACCOUNT_NAME}.blob.core.windows.net`,
  credential
);

export async function uploadSourceFile(
  jobId: string,
  file: { originalname: string; buffer: Buffer; mimetype: string }
): Promise<string> {
  const container = blobService.getContainerClient(`source-${jobId}`);
  await container.createIfNotExists();
  await container
    .getBlockBlobClient(file.originalname)
    .uploadData(file.buffer, { blobHTTPHeaders: { blobContentType: file.mimetype } });
  return container.url; // no SAS — Translator MI reads via RBAC
}

export async function ensureTargetContainer(jobId: string): Promise<string> {
  const container = blobService.getContainerClient(`target-${jobId}`);
  await container.createIfNotExists();
  return container.url;
}

async function getTranslatorToken(): Promise<string> {
  const token = await credential.getToken("https://cognitiveservices.azure.com/.default");
  if (!token) throw new Error("Failed to acquire Translator token");
  return token.token;
}

async function translatorFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const token = await getTranslatorToken();
  const url = `${TRANSLATOR_ENDPOINT}${path}${path.includes("?") ? "&" : "?"}api-version=${TRANSLATOR_API_VERSION}`;
  const headers = new Headers(init.headers);
  headers.set("Authorization", `Bearer ${token}`);
  if (init.body && !headers.has("Content-Type")) headers.set("Content-Type", "application/json");
  return fetch(url, { ...init, headers });
}

export interface SubmitJobInput {
  sourceUrl: string;
  targetUrl: string;
  sourceLanguage: string;
  targetLanguage: string;
}

export async function submitBatchJob(input: SubmitJobInput): Promise<string> {
  const body = {
    inputs: [
      {
        source: {
          sourceUrl: input.sourceUrl,
          language: input.sourceLanguage,
          storageSource: "AzureBlob",
        },
        targets: [
          {
            targetUrl: input.targetUrl,
            language: input.targetLanguage,
            category: "general",
          },
        ],
      },
    ],
  };

  const res = await translatorFetch("/translator/document/batches", {
    method: "POST",
    body: JSON.stringify(body),
  });
  if (res.status !== 202) {
    throw new Error(`Translator submit failed (${res.status}): ${await res.text()}`);
  }
  const opLoc = res.headers.get("operation-location");
  if (!opLoc) throw new Error("Missing Operation-Location header from Translator");
  const id = opLoc.split("/").pop()?.split("?")[0];
  if (!id) throw new Error(`Cannot parse job id from ${opLoc}`);
  return id;
}

export async function listBatchJobs(): Promise<unknown> {
  const res = await translatorFetch("/translator/document/batches");
  if (!res.ok) throw new Error(`Translator list failed (${res.status}): ${await res.text()}`);
  return res.json();
}

export async function getBatchJob(id: string): Promise<unknown> {
  const res = await translatorFetch(`/translator/document/batches/${id}`);
  if (!res.ok) throw new Error(`Translator status failed (${res.status}): ${await res.text()}`);
  return res.json();
}

export async function getBatchJobDocuments(id: string): Promise<unknown> {
  const res = await translatorFetch(`/translator/document/batches/${id}/documents`);
  if (!res.ok) throw new Error(`Translator documents failed (${res.status}): ${await res.text()}`);
  return res.json();
}

/** Download a translated blob via managed identity (no SAS). Given the Translator-reported `path`
 *  (a full blob URL like https://acct.blob.core.windows.net/target-<jobId>/<name>), this returns
 *  a readable stream plus content metadata for the HTTP response. */
export async function downloadTargetBlob(blobUrl: string): Promise<{
  stream: NodeJS.ReadableStream;
  contentType?: string;
  contentLength?: number;
  fileName: string;
}> {
  const u = new URL(blobUrl);
  const expectedHost = `${STORAGE_ACCOUNT_NAME}.blob.core.windows.net`.toLowerCase();
  if (u.host.toLowerCase() !== expectedHost) {
    throw new Error(`Refusing to download blob from unexpected host: ${u.host}`);
  }
  // pathname: /<container>/<blob...>
  const parts = u.pathname.replace(/^\/+/, "").split("/");
  const container = parts.shift();
  const blobName = parts.join("/");
  if (!container || !blobName) throw new Error(`Cannot parse container/blob from ${blobUrl}`);
  const client = blobService.getContainerClient(container).getBlobClient(decodeURIComponent(blobName));
  const dl = await client.download();
  if (!dl.readableStreamBody) throw new Error("Blob download returned no stream");
  return {
    stream: dl.readableStreamBody,
    contentType: dl.contentType,
    contentLength: dl.contentLength,
    fileName: decodeURIComponent(blobName.split("/").pop() ?? "download"),
  };
}

export async function cancelBatchJob(id: string): Promise<void> {
  const res = await translatorFetch(`/translator/document/batches/${id}`, { method: "DELETE" });
  if (res.status !== 200 && res.status !== 202 && res.status !== 204) {
    throw new Error(`Translator cancel failed (${res.status}): ${await res.text()}`);
  }
}
