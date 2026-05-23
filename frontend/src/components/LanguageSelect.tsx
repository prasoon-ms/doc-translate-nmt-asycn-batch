import { useEffect, useState } from "react";

export type Language = { code: string; name: string; nativeName: string; dir: string };

let cachePromise: Promise<Language[]> | null = null;

export function useLanguages(): { languages: Language[]; error: string | null; loading: boolean } {
  const [languages, setLanguages] = useState<Language[]>([]);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!cachePromise) {
      cachePromise = fetch("/api/languages")
        .then(async (r) => {
          if (!r.ok) throw new Error(`languages ${r.status}`);
          const data = (await r.json()) as { languages: Language[] };
          return data.languages;
        })
        .catch((e) => {
          cachePromise = null;
          throw e;
        });
    }
    cachePromise
      .then((langs) => {
        setLanguages(langs);
        setLoading(false);
      })
      .catch((e: Error) => {
        setError(e.message);
        setLoading(false);
      });
  }, []);

  return { languages, error, loading };
}

type Props = {
  name: string;
  label: string;
  defaultValue?: string;
};

export function LanguageSelect({ name, label, defaultValue }: Props) {
  const { languages, error, loading } = useLanguages();
  return (
    <label>
      {label}:{" "}
      <select name={name} defaultValue={defaultValue} disabled={loading || !!error} required>
        {loading && <option value="">Loading…</option>}
        {error && <option value="">Failed to load</option>}
        {!loading &&
          !error &&
          languages.map((l) => (
            <option key={l.code} value={l.code}>
              {l.name} ({l.code})
            </option>
          ))}
      </select>
    </label>
  );
}
