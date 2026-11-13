import { createHash, timingSafeEqual } from "node:crypto";

const REALM = 'Basic realm="Crewlist", charset="UTF-8"';

function sha256(value) {
  return createHash("sha256").update(value, "utf8").digest();
}

function sameSecret(left, right) {
  return timingSafeEqual(sha256(left), sha256(right));
}

function configured(env) {
  const user = env?.BASIC_AUTH_USER;
  const password = env?.BASIC_AUTH_PASSWORD;
  return typeof user === "string" && user.length > 0 && typeof password === "string" && password.length > 0;
}

function decodeToken(token) {
  if (!/^[A-Za-z0-9+/]+={0,2}$/.test(token) || token.length % 4 !== 0) return null;
  const bytes = Buffer.from(token, "base64");
  if (bytes.toString("base64") !== token) return null;
  try {
    return new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch {
    return null;
  }
}

function parseBasic(header) {
  if (typeof header !== "string") return null;
  const match = /^Basic[ \t]+([A-Za-z0-9+/]+={0,2})$/i.exec(header);
  if (!match) return null;
  const decoded = decodeToken(match[1]);
  if (decoded == null) return null;
  const separator = decoded.indexOf(":");
  if (separator < 0) return null;
  return {
    user: decoded.slice(0, separator),
    password: decoded.slice(separator + 1),
  };
}

function privateHeaders(extra) {
  return {
    "Cache-Control": "private, no-store",
    "CDN-Cache-Control": "no-store",
    "Vercel-CDN-Cache-Control": "no-store",
    Vary: "Authorization",
    "X-Content-Type-Options": "nosniff",
    ...extra,
  };
}

export function unavailable() {
  return {
    ok: false,
    status: 503,
    body: "No disponible",
    headers: privateHeaders({
      "Content-Type": "text/plain; charset=utf-8",
    }),
  };
}

export function authorize(authorizationHeader, env = process.env) {
  if (!configured(env)) return unavailable();

  const parsed = parseBasic(authorizationHeader);
  const userMatches = sameSecret(parsed?.user ?? "", env.BASIC_AUTH_USER);
  const passwordMatches = sameSecret(parsed?.password ?? "", env.BASIC_AUTH_PASSWORD);
  if (parsed && userMatches && passwordMatches) return { ok: true };

  return {
    ok: false,
    status: 401,
    body: "Autenticación requerida",
    headers: privateHeaders({
      "Content-Type": "text/plain; charset=utf-8",
      "WWW-Authenticate": REALM,
    }),
  };
}
