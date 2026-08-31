import crypto from "crypto";
import { cookies } from "next/headers";

export const ADMIN_COOKIE_NAME = "marketing-dashboard-admin";
const SESSION_DURATION_SECONDS = 60 * 60 * 12;

function getSessionSecret() {
  return process.env.ADMIN_SESSION_SECRET || "";
}

function signSessionValue(value) {
  return crypto
    .createHmac("sha256", getSessionSecret())
    .update(value)
    .digest("base64url");
}

function safeEqual(left, right) {
  const leftBuffer = Buffer.from(String(left));
  const rightBuffer = Buffer.from(String(right));
  return leftBuffer.length === rightBuffer.length
    && crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

export function isAdminAuthConfigured() {
  return Boolean(process.env.ADMIN_PASSWORD && getSessionSecret());
}

export function verifyAdminPassword(password) {
  if (!isAdminAuthConfigured()) return false;
  return safeEqual(password, process.env.ADMIN_PASSWORD);
}

export function createAdminSessionToken() {
  const expiresAt = String(Math.floor(Date.now() / 1000) + SESSION_DURATION_SECONDS);
  return `${expiresAt}.${signSessionValue(expiresAt)}`;
}

export function verifyAdminSessionToken(token) {
  if (!isAdminAuthConfigured() || !token) return false;
  const [expiresAt, signature, ...rest] = String(token).split(".");
  if (!expiresAt || !signature || rest.length > 0) return false;
  if (!/^\d+$/.test(expiresAt) || Number(expiresAt) <= Math.floor(Date.now() / 1000)) return false;
  return safeEqual(signature, signSessionValue(expiresAt));
}

export async function isAdminAuthenticated() {
  const cookieStore = await cookies();
  return verifyAdminSessionToken(cookieStore.get(ADMIN_COOKIE_NAME)?.value);
}

export const adminSessionCookieOptions = {
  httpOnly: true,
  sameSite: "strict",
  secure: process.env.NODE_ENV === "production",
  path: "/",
  maxAge: SESSION_DURATION_SECONDS
};
