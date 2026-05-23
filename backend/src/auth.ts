import type { NextFunction, Request, Response } from "express";
import jwt, { type JwtHeader, type SigningKeyCallback } from "jsonwebtoken";
import jwksClient from "jwks-rsa";

const TENANT_ID = requireEnv("AZURE_TENANT_ID");
const API_CLIENT_ID = requireEnv("API_CLIENT_ID");
const REQUIRED_SCOPE = process.env.REQUIRED_SCOPE ?? "Translate.Submit";

function requireEnv(name: string): string {
  const v = process.env[name];
  if (!v) throw new Error(`Missing required env var: ${name}`);
  return v;
}

const jwks = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`,
  cache: true,
  rateLimit: true,
});

function getKey(header: JwtHeader, cb: SigningKeyCallback) {
  jwks.getSigningKey(header.kid!, (err, key) => {
    if (err || !key) return cb(err ?? new Error("Signing key not found"));
    cb(null, key.getPublicKey());
  });
}

export interface AuthenticatedRequest extends Request {
  user?: { oid?: string; preferred_username?: string; scp?: string };
}

export function requireUser(req: AuthenticatedRequest, res: Response, next: NextFunction) {
  const auth = req.headers.authorization ?? "";
  const token = auth.startsWith("Bearer ") ? auth.slice(7) : null;
  if (!token) {
    res.status(401).json({ error: "missing_token" });
    return;
  }

  jwt.verify(
    token,
    getKey,
    {
      audience: [`api://${API_CLIENT_ID}`, API_CLIENT_ID],
      issuer: `https://login.microsoftonline.com/${TENANT_ID}/v2.0`,
      algorithms: ["RS256"],
    },
    (err, payload) => {
      if (err || !payload || typeof payload === "string") {
        console.warn("JWT verify failed:", err?.message ?? "no payload");
        res.status(401).json({ error: "invalid_token", message: err?.message });
        return;
      }
      const scopes = (payload.scp ?? "").toString().split(" ");
      if (!scopes.includes(REQUIRED_SCOPE)) {
        console.warn(`Token missing scope ${REQUIRED_SCOPE}; got:`, payload.scp);
        res.status(403).json({ error: "insufficient_scope", got: payload.scp });
        return;
      }
      req.user = payload as AuthenticatedRequest["user"];
      next();
    }
  );
}
