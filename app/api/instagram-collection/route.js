import { createHmac, timingSafeEqual } from "crypto";
import { promises as fs } from "fs";
import path from "path";
import { kv } from "@vercel/kv";

export const dynamic = "force-dynamic";
export const revalidate = 0;

const TOKEN_TTL_SECONDS = 5 * 60;
const CACHE_TTL_SECONDS = 30 * 24 * 60 * 60;
const MAX_POSTS = 12;
const localCachePath = path.join(process.cwd(), "data", "instagramPosts.json");
const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type",
  "Cache-Control": "no-store, max-age=0, must-revalidate"
};

function hasSharedStorage() {
  return Boolean(process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN);
}

function cacheKey(branch) {
  const branchKey = Buffer.from(branch, "utf8").toString("base64url");
  return `instagram-posts:${branchKey}`;
}

function getSecret() {
  return process.env.INSTAGRAM_COLLECTOR_SECRET || "";
}

function encodeToken(payload) {
  const secret = getSecret();
  if (!secret) throw new Error("Instagram collector secret is not configured.");

  const encodedPayload = Buffer.from(JSON.stringify(payload), "utf8").toString("base64url");
  const signature = createHmac("sha256", secret).update(encodedPayload).digest("base64url");
  return `${encodedPayload}.${signature}`;
}

function decodeToken(token) {
  const secret = getSecret();
  if (!secret) throw new Error("Instagram collector secret is not configured.");

  const [encodedPayload, providedSignature] = String(token || "").split(".");
  if (!encodedPayload || !providedSignature) throw new Error("Invalid collector token.");

  const expectedSignature = createHmac("sha256", secret).update(encodedPayload).digest("base64url");
  const expectedBuffer = Buffer.from(expectedSignature);
  const providedBuffer = Buffer.from(providedSignature);
  if (
    expectedBuffer.length !== providedBuffer.length ||
    !timingSafeEqual(expectedBuffer, providedBuffer)
  ) {
    throw new Error("Invalid collector token.");
  }

  const payload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8"));
  if (!payload.branch || !payload.username || !payload.exp || Date.now() > payload.exp) {
    throw new Error("Collector token expired or is incomplete.");
  }
  return payload;
}

function getInstagramUsername(rawUrl) {
  try {
    const parsed = new URL(String(rawUrl || "").trim());
    if (!["instagram.com", "www.instagram.com"].includes(parsed.hostname.toLowerCase())) return "";
    const username = parsed.pathname.split("/").filter(Boolean)[0] || "";
    return /^[a-zA-Z0-9._]{1,30}$/.test(username) ? username : "";
  } catch {
    return "";
  }
}

function cleanText(value, maxLength) {
  return String(value || "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

function cleanHttpUrl(value, allowedHosts) {
  try {
    const parsed = new URL(String(value || ""));
    if (parsed.protocol !== "https:") return "";
    const hostname = parsed.hostname.toLowerCase();
    if (!allowedHosts.some((host) => hostname === host || hostname.endsWith(`.${host}`))) return "";
    return parsed.toString();
  } catch {
    return "";
  }
}

function sanitizePosts(posts) {
  if (!Array.isArray(posts)) return [];

  const seenUrls = new Set();
  return posts.slice(0, MAX_POSTS * 2).reduce((result, post) => {
    if (result.length >= MAX_POSTS) return result;

    const url = cleanHttpUrl(post?.url, ["instagram.com"]);
    if (!url || seenUrls.has(url)) return result;

    const thumbnailUrl = cleanHttpUrl(post?.thumbnailUrl, ["cdninstagram.com", "fbcdn.net"]);
    seenUrls.add(url);
    result.push({
      url,
      thumbnailUrl,
      caption: cleanText(post?.caption, 500),
      publishedAt: cleanText(post?.publishedAt, 40),
      type: ["reel", "carousel", "image"].includes(post?.type) ? post.type : "image"
    });
    return result;
  }, []);
}

async function readLocalCache() {
  try {
    return JSON.parse(await fs.readFile(localCachePath, "utf8"));
  } catch {
    return {};
  }
}

async function readCollection(branch) {
  if (hasSharedStorage()) {
    const collection = await kv.get(cacheKey(branch));
    return typeof collection === "string" ? JSON.parse(collection) : collection;
  }
  if (process.env.VERCEL) return null;
  const cache = await readLocalCache();
  return cache[branch] || null;
}

async function writeCollection(branch, collection) {
  if (hasSharedStorage()) {
    await kv.set(cacheKey(branch), JSON.stringify(collection), { ex: CACHE_TTL_SECONDS });
    return;
  }
  if (process.env.VERCEL) throw new Error("Shared storage is not configured.");

  const cache = await readLocalCache();
  cache[branch] = collection;
  await fs.mkdir(path.dirname(localCachePath), { recursive: true });
  await fs.writeFile(localCachePath, JSON.stringify(cache, null, 2), "utf8");
}

export function OPTIONS() {
  return new Response(null, { status: 204, headers: corsHeaders });
}

export async function GET(request) {
  const branch = cleanText(new URL(request.url).searchParams.get("branch"), 80);
  if (!branch) {
    return Response.json({ error: "지점명이 필요합니다." }, { status: 400, headers: corsHeaders });
  }

  try {
    const collection = await readCollection(branch);
    return Response.json(
      { ok: true, collection: collection || null },
      { headers: corsHeaders }
    );
  } catch (error) {
    console.error("Failed to read Instagram collection.", error);
    return Response.json({ error: "인스타그램 수집 데이터를 불러오지 못했습니다." }, { status: 500, headers: corsHeaders });
  }
}

export async function POST(request) {
  try {
    const body = await request.json();

    if (body.action === "start") {
      const branch = cleanText(body.branch, 80);
      const username = getInstagramUsername(body.instagramUrl);
      if (!branch || !username) {
        return Response.json({ error: "올바른 지점명과 인스타그램 주소가 필요합니다." }, { status: 400, headers: corsHeaders });
      }

      const expiresAt = Date.now() + TOKEN_TTL_SECONDS * 1000;
      const token = encodeToken({ branch, username, exp: expiresAt });
      const origin = new URL(request.url).origin;
      const fragment = new URLSearchParams({
        etoos247_collect: token,
        etoos247_origin: origin
      });

      return Response.json({
        ok: true,
        expiresAt: new Date(expiresAt).toISOString(),
        collectUrl: `https://www.instagram.com/${username}/#${fragment.toString()}`
      }, { headers: corsHeaders });
    }

    if (body.action === "complete") {
      const tokenPayload = decodeToken(body.token);
      const posts = sanitizePosts(body.posts);
      if (!posts.length) {
        return Response.json({ error: "수집된 공개 게시물이 없습니다." }, { status: 422, headers: corsHeaders });
      }

      const collection = {
        branch: tokenPayload.branch,
        username: tokenPayload.username,
        posts,
        collectedAt: new Date().toISOString()
      };
      await writeCollection(tokenPayload.branch, collection);
      return Response.json({ ok: true, collection }, { headers: corsHeaders });
    }

    return Response.json({ error: "지원하지 않는 요청입니다." }, { status: 400, headers: corsHeaders });
  } catch (error) {
    console.error("Instagram collection request failed.", error);
    const errorMessage = String(error?.message || "");
    const isConfigurationError = errorMessage.includes("not configured");
    const isTokenError = /token|expired/i.test(errorMessage);
    return Response.json(
      { error: isConfigurationError ? "Instagram 수집 환경 변수가 설정되지 않았습니다." : isTokenError ? "수집 인증 정보가 만료되었거나 올바르지 않습니다." : "인스타그램 수집 요청을 처리하지 못했습니다." },
      { status: isConfigurationError ? 503 : isTokenError ? 401 : 500, headers: corsHeaders }
    );
  }
}
