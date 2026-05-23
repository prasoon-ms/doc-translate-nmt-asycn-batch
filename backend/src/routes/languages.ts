import { Router } from "express";

export const languagesRouter: Router = Router();

type LanguageEntry = { name: string; nativeName: string; dir: string };
type LanguagesResponse = { translation: Record<string, LanguageEntry> };

let cache: { at: number; data: LanguagesResponse } | null = null;
const TTL_MS = 24 * 60 * 60 * 1000;

languagesRouter.get("/languages", async (_req, res) => {
  try {
    if (!cache || Date.now() - cache.at > TTL_MS) {
      const r = await fetch(
        "https://api.cognitive.microsofttranslator.com/languages?api-version=3.0&scope=translation",
        { headers: { "Accept-Language": "en" } }
      );
      if (!r.ok) throw new Error(`languages fetch failed (${r.status})`);
      cache = { at: Date.now(), data: (await r.json()) as LanguagesResponse };
    }
    const list = Object.entries(cache.data.translation)
      .map(([code, v]) => ({ code, name: v.name, nativeName: v.nativeName, dir: v.dir }))
      .sort((a, b) => a.name.localeCompare(b.name));
    res.setHeader("Cache-Control", "public, max-age=3600");
    res.json({ languages: list });
  } catch (err) {
    res.status(502).json({ error: "languages_failed", message: (err as Error).message });
  }
});
