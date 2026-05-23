import { useState } from "react";
import { apiFetch } from "../auth";
import { LanguageSelect } from "./LanguageSelect";

export function UploadForm() {
  const [error, setError] = useState<string | null>(null);
  const [jobId, setJobId] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function handleSubmit(e: React.FormEvent<HTMLFormElement>) {
    e.preventDefault();
    setError(null);
    setJobId(null);
    setBusy(true);
    try {
      const form = new FormData(e.currentTarget);
      const res = await apiFetch("/api/translate", { method: "POST", body: form });
      if (!res.ok) {
        const text = await res.text();
        setError(`Submit failed (${res.status}): ${text}`);
        return;
      }
      const job = await res.json();
      setJobId(job.jobId);
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <section>
      <h2>New translation</h2>
      <form onSubmit={handleSubmit}>
        <p>
          <label>
            Documents:{" "}
            <input
              type="file"
              name="documents"
              multiple
              required
              accept=".docx,.pptx,.xlsx,.pdf,.txt,.html,.md"
            />
          </label>
        </p>
        <p>
          <LanguageSelect name="sourceLanguage" label="Source language" defaultValue="en" />
        </p>
        <p>
          <LanguageSelect name="targetLanguage" label="Target language" defaultValue="fr" />
        </p>
        <button type="submit" disabled={busy}>
          {busy ? "Submitting…" : "Translate"}
        </button>
      </form>
      {jobId && <p>Submitted. Job id: <code>{jobId}</code></p>}
      {error && <p role="alert" style={{ color: "crimson" }}>{error}</p>}
    </section>
  );
}
