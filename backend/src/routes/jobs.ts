import { Router } from "express";
import { requireUser } from "../auth.js";
import {
  cancelBatchJob,
  downloadTargetBlob,
  getBatchJob,
  getBatchJobDocuments,
  listBatchJobs,
} from "../azure.js";

export const jobsRouter: Router = Router();

jobsRouter.use(requireUser);

type Doc = {
  id?: string;
  status?: string;
  path?: string;
  sourcePath?: string;
  to?: string;
  characterCharged?: number;
  progress?: number;
  error?: unknown;
};

jobsRouter.get("/jobs", async (_req, res) => {
  try {
    const data = (await listBatchJobs()) as { value?: Array<Record<string, unknown>> };
    const jobs = (data.value ?? []).map((j) => ({
      jobId: j.id,
      status: j.status,
      createdAt: j.createdDateTimeUtc,
      summary: j.summary,
    }));
    res.json({ jobs });
  } catch (err) {
    res.status(502).json({ error: "list_failed", message: (err as Error).message });
  }
});

jobsRouter.get("/jobs/:id", async (req, res) => {
  try {
    const [summary, documentsRaw] = await Promise.all([
      getBatchJob(req.params.id),
      getBatchJobDocuments(req.params.id),
    ]);
    const docs = ((documentsRaw as { value?: Doc[] }).value ?? []).map((d) => ({
      id: d.id,
      status: d.status,
      to: d.to,
      progress: d.progress,
      characterCharged: d.characterCharged,
      fileName: d.path ? decodeURIComponent(d.path.split("/").pop()?.split("?")[0] ?? "") : undefined,
      downloadUrl:
        d.status === "Succeeded" && d.id
          ? `/api/jobs/${encodeURIComponent(req.params.id)}/documents/${encodeURIComponent(d.id)}/download`
          : undefined,
      error: d.error,
    }));
    res.json({ jobId: req.params.id, summary, documents: docs });
  } catch (err) {
    res.status(502).json({ error: "status_failed", message: (err as Error).message });
  }
});

jobsRouter.get("/jobs/:id/documents/:docId/download", async (req, res) => {
  try {
    const documentsRaw = (await getBatchJobDocuments(req.params.id)) as { value?: Doc[] };
    const doc = (documentsRaw.value ?? []).find((d) => d.id === req.params.docId);
    if (!doc) {
      res.status(404).json({ error: "document_not_found" });
      return;
    }
    if (doc.status !== "Succeeded" || !doc.path) {
      res.status(409).json({ error: "not_ready", status: doc.status });
      return;
    }
    const { stream, contentType, contentLength, fileName } = await downloadTargetBlob(doc.path);
    if (contentType) res.setHeader("Content-Type", contentType);
    if (contentLength) res.setHeader("Content-Length", String(contentLength));
    res.setHeader("Content-Disposition", `attachment; filename="${fileName.replace(/"/g, "")}"`);
    stream.on("error", (e) => {
      console.error("blob stream error", e);
      if (!res.headersSent) res.status(502);
      res.end();
    });
    stream.pipe(res);
  } catch (err) {
    console.error("download failed", err);
    if (!res.headersSent) res.status(502).json({ error: "download_failed", message: (err as Error).message });
  }
});

jobsRouter.delete("/jobs/:id", async (req, res) => {
  try {
    await cancelBatchJob(req.params.id);
    res.status(204).end();
  } catch (err) {
    res.status(502).json({ error: "cancel_failed", message: (err as Error).message });
  }
});
