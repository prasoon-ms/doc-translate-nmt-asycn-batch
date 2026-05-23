import { useEffect, useState } from "react";
import { apiFetch } from "../auth";

type Job = { jobId: string; status: string; createdAt?: string };
type JobDoc = {
  id: string;
  status: string;
  to?: string;
  fileName?: string;
  progress?: number;
  characterCharged?: number;
  downloadUrl?: string;
  error?: unknown;
};

function JobRow({ job }: { job: Job }) {
  const [open, setOpen] = useState(false);
  const [docs, setDocs] = useState<JobDoc[] | null>(null);
  const [err, setErr] = useState<string | null>(null);
  const [downloading, setDownloading] = useState<string | null>(null);

  async function loadDocs() {
    setErr(null);
    try {
      const r = await apiFetch(`/api/jobs/${encodeURIComponent(job.jobId)}`);
      if (!r.ok) {
        setErr(`Load failed (${r.status})`);
        return;
      }
      const data = await r.json();
      setDocs(data.documents ?? []);
    } catch (e) {
      setErr((e as Error).message);
    }
  }

  useEffect(() => {
    if (!open) return;
    loadDocs();
    // refresh while job is in progress
    const inProgress = job.status !== "Succeeded" && job.status !== "Failed" && job.status !== "Cancelled";
    if (!inProgress) return;
    const t = setInterval(loadDocs, 5000);
    return () => clearInterval(t);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open, job.status, job.jobId]);

  async function download(doc: JobDoc) {
    if (!doc.downloadUrl) return;
    setDownloading(doc.id);
    try {
      const r = await apiFetch(doc.downloadUrl);
      if (!r.ok) {
        setErr(`Download failed (${r.status})`);
        return;
      }
      const blob = await r.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = doc.fileName ?? "translated";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (e) {
      setErr((e as Error).message);
    } finally {
      setDownloading(null);
    }
  }

  return (
    <li style={{ marginBottom: "0.75rem" }}>
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        style={{ marginRight: "0.5rem" }}
        aria-expanded={open}
      >
        {open ? "▼" : "▶"}
      </button>
      <code>{job.jobId}</code> — {job.status}
      {job.createdAt ? ` (${new Date(job.createdAt).toLocaleString()})` : ""}
      {open && (
        <div style={{ marginLeft: "1.5rem", marginTop: "0.5rem" }}>
          {err && <p role="alert" style={{ color: "crimson" }}>{err}</p>}
          {!docs && !err && <p>Loading documents…</p>}
          {docs && docs.length === 0 && <p>No documents.</p>}
          {docs && docs.length > 0 && (
            <table style={{ borderCollapse: "collapse" }}>
              <thead>
                <tr>
                  <th style={{ textAlign: "left", paddingRight: "1rem" }}>File</th>
                  <th style={{ textAlign: "left", paddingRight: "1rem" }}>Target</th>
                  <th style={{ textAlign: "left", paddingRight: "1rem" }}>Status</th>
                  <th style={{ textAlign: "left", paddingRight: "1rem" }}>Progress</th>
                  <th style={{ textAlign: "left" }}>Download</th>
                </tr>
              </thead>
              <tbody>
                {docs.map((d) => (
                  <tr key={d.id}>
                    <td style={{ paddingRight: "1rem" }}>{d.fileName ?? d.id}</td>
                    <td style={{ paddingRight: "1rem" }}>{d.to ?? "—"}</td>
                    <td style={{ paddingRight: "1rem" }}>{d.status}</td>
                    <td style={{ paddingRight: "1rem" }}>
                      {typeof d.progress === "number" ? `${Math.round(d.progress * 100)}%` : "—"}
                    </td>
                    <td>
                      {d.downloadUrl ? (
                        <button
                          type="button"
                          onClick={() => download(d)}
                          disabled={downloading === d.id}
                        >
                          {downloading === d.id ? "Downloading…" : "Download"}
                        </button>
                      ) : (
                        <span style={{ color: "#888" }}>—</span>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      )}
    </li>
  );
}

export function JobList() {
  const [jobs, setJobs] = useState<Job[]>([]);
  const [error, setError] = useState<string | null>(null);

  async function refresh() {
    setError(null);
    try {
      const res = await apiFetch("/api/jobs");
      if (!res.ok) {
        setError(`Load failed (${res.status})`);
        return;
      }
      const data = await res.json();
      setJobs(data.jobs ?? []);
    } catch (err) {
      setError((err as Error).message);
    }
  }

  useEffect(() => {
    refresh();
    const t = setInterval(refresh, 5000);
    return () => clearInterval(t);
  }, []);

  return (
    <section>
      <h2>Jobs</h2>
      <button onClick={refresh}>Refresh</button>
      {error && <p role="alert" style={{ color: "crimson" }}>{error}</p>}
      {jobs.length === 0 ? (
        <p>No jobs yet.</p>
      ) : (
        <ul style={{ listStyle: "none", padding: 0 }}>
          {jobs.map((j) => (
            <JobRow key={j.jobId} job={j} />
          ))}
        </ul>
      )}
    </section>
  );
}
