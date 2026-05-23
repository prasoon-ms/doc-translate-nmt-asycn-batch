import { Router } from "express";
import multer from "multer";
import { randomUUID } from "node:crypto";
import { requireUser } from "../auth.js";
import { ensureTargetContainer, submitBatchJob, uploadSourceFile } from "../azure.js";

const ALLOWED_EXT = new Set([".docx", ".pptx", ".xlsx", ".pdf", ".txt", ".html", ".md"]);
const MAX_FILE_BYTES = 40 * 1024 * 1024; // 40 MB
const MAX_REQUEST_BYTES = 250 * 1024 * 1024; // 250 MB
const MAX_FILES = 50;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: MAX_FILE_BYTES, files: MAX_FILES, fieldSize: MAX_REQUEST_BYTES },
  fileFilter: (_req, file, cb) => {
    const ext = file.originalname.slice(file.originalname.lastIndexOf(".")).toLowerCase();
    if (!ALLOWED_EXT.has(ext)) {
      cb(new Error(`Unsupported file extension: ${ext}`));
      return;
    }
    cb(null, true);
  },
});

export const translateRouter: Router = Router();

translateRouter.post("/translate", requireUser, upload.array("documents", MAX_FILES), async (req, res) => {
  try {
    const files = (req.files as Express.Multer.File[]) ?? [];
    if (files.length === 0) {
      res.status(400).json({ error: "no_files" });
      return;
    }
    const sourceLanguage = String(req.body.sourceLanguage ?? "").trim();
    const targetLanguage = String(req.body.targetLanguage ?? "").trim();
    if (!sourceLanguage || !targetLanguage) {
      res.status(400).json({ error: "missing_language" });
      return;
    }

    const jobId = randomUUID();
    let sourceUrl = "";
    for (const file of files) {
      sourceUrl = await uploadSourceFile(jobId, file);
    }
    const targetUrl = await ensureTargetContainer(jobId);

    const translatorJobId = await submitBatchJob({
      sourceUrl,
      targetUrl,
      sourceLanguage,
      targetLanguage,
    });

    res.status(202).json({
      jobId: translatorJobId,
      localJobId: jobId,
      status: "NotStarted",
    });
  } catch (err) {
    const message = (err as Error).message;
    if (message.includes("Unsupported file extension")) {
      res.status(415).json({ error: message });
      return;
    }
    console.error("translate failed", err);
    res.status(500).json({ error: "translate_failed", message });
  }
});
