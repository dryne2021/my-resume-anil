import { createHmac, timingSafeEqual } from "node:crypto";
import { getCookie } from "@tanstack/react-start/server";

// Server-only. No database: the login is a signed cookie checked on every request.
export const SESSION_COOKIE = "dryne_session";
export const SESSION_MAX_AGE = 60 * 60 * 24 * 30; // 30 days

function secret() {
  const s = process.env.SESSION_SECRET || process.env.OWNER_PASSWORD;
  if (!s) throw new Error("SESSION_SECRET / OWNER_PASSWORD not configured");
  return s;
}

function sign(value: string) {
  return createHmac("sha256", secret()).update(value).digest("base64url");
}

export function safeEqual(a: string, b: string) {
  const x = Buffer.from(a);
  const y = Buffer.from(b);
  return x.length === y.length && timingSafeEqual(x, y);
}

export function createSessionToken() {
  const expires = String(Date.now() + SESSION_MAX_AGE * 1000);
  return `${expires}.${sign(expires)}`;
}

export function isSessionValid(token: string | undefined) {
  if (!token) return false;
  const [expires, sig] = token.split(".");
  if (!expires || !sig) return false;
  if (!safeEqual(sig, sign(expires))) return false;
  return Number(expires) > Date.now();
}

export function hasValidSession() {
  return isSessionValid(getCookie(SESSION_COOKIE));
}
