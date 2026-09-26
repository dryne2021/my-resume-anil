import { createMiddleware, createServerFn } from "@tanstack/react-start";
import { deleteCookie, setCookie } from "@tanstack/react-start/server";
import {
  SESSION_COOKIE,
  SESSION_MAX_AGE,
  createSessionToken,
  hasValidSession,
  safeEqual,
} from "./session.server";

/** Blocks a server function unless the owner is logged in. */
export const requireOwner = createMiddleware({ type: "function" }).server(async ({ next }) => {
  if (!hasValidSession()) throw new Error("Unauthorized: please sign in again.");
  return next();
});

export const getSession = createServerFn({ method: "GET" }).handler(async () => {
  return { loggedIn: hasValidSession() };
});

export const login = createServerFn({ method: "POST" })
  .inputValidator((input: { email: string; password: string }) => {
    if (!input || typeof input.email !== "string" || typeof input.password !== "string") {
      throw new Error("Email and password required");
    }
    return { email: input.email.trim().toLowerCase(), password: input.password };
  })
  .handler(async ({ data }) => {
    const ownerEmail = (process.env.OWNER_EMAIL ?? "").trim().toLowerCase();
    const ownerPassword = process.env.OWNER_PASSWORD ?? "";
    if (!ownerEmail || !ownerPassword) {
      throw new Error("OWNER_EMAIL / OWNER_PASSWORD not configured on the server");
    }
    const ok = safeEqual(data.email, ownerEmail) && safeEqual(data.password, ownerPassword);
    if (!ok) return { ok: false as const };

    setCookie(SESSION_COOKIE, createSessionToken(), {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      sameSite: "lax",
      path: "/",
      maxAge: SESSION_MAX_AGE,
    });
    return { ok: true as const };
  });

export const logout = createServerFn({ method: "POST" }).handler(async () => {
  deleteCookie(SESSION_COOKIE, { path: "/" });
  return { ok: true };
});
