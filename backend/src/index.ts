import "dotenv/config";
import express from "express";
import { translateRouter } from "./routes/translate.js";
import { jobsRouter } from "./routes/jobs.js";
import { languagesRouter } from "./routes/languages.js";

const app = express();
app.use(express.json());

app.get("/healthz", (_req, res) => {
  res.json({ ok: true });
});

app.use("/api", languagesRouter);
app.use("/api", translateRouter);
app.use("/api", jobsRouter);

const port = Number(process.env.PORT ?? 8080);
app.listen(port, () => {
  console.log(`Backend listening on http://localhost:${port}`);
});
