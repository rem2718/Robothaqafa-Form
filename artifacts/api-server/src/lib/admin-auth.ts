import { createHmac, timingSafeEqual } from "node:crypto";
import type { Request, Response } from "express";

const COOKIE_NAME = "randomized_admin_session";

function configuredPassword() {
  const password = process.env.ADMIN_PASSWORD;
  if (!password) throw new Error("ADMIN_PASSWORD must be configured.");
  return password;
}

function sessionToken() {
  return createHmac(
    "sha256",
    process.env.SESSION_SECRET ?? configuredPassword(),
  )
    .update("randomized-forms-admin")
    .digest("hex");
}

function cookieValue(req: Request) {
  const cookies = req.headers.cookie?.split(";") ?? [];
  const entry = cookies.find((cookie) =>
    cookie.trim().startsWith(`${COOKIE_NAME}=`),
  );
  return entry?.trim().slice(COOKIE_NAME.length + 1);
}

export function isAdminAuthenticated(req: Request) {
  const actual = cookieValue(req);
  const expected = sessionToken();
  if (!actual || actual.length !== expected.length) return false;
  return timingSafeEqual(Buffer.from(actual), Buffer.from(expected));
}

export function verifyAdminPassword(password: string) {
  const expected = configuredPassword();
  if (password.length !== expected.length) return false;
  return timingSafeEqual(Buffer.from(password), Buffer.from(expected));
}

export function setAdminSession(res: Response) {
  const secure =
    process.env.COOKIE_SECURE === "true" ||
    (process.env.COOKIE_SECURE !== "false" &&
      process.env.NODE_ENV === "production")
      ? "; Secure"
      : "";
  res.setHeader(
    "Set-Cookie",
    `${COOKIE_NAME}=${sessionToken()}; Path=/; HttpOnly; SameSite=Lax${secure}`,
  );
}

export function clearAdminSession(res: Response) {
  res.setHeader(
    "Set-Cookie",
    `${COOKIE_NAME}=; Path=/; HttpOnly; Max-Age=0; SameSite=Lax`,
  );
}

export function requireAdmin(req: Request, res: Response) {
  if (isAdminAuthenticated(req)) return true;
  res.status(401).json({ error: "Admin session required" });
  return false;
}