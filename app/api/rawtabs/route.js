import { promises as fs } from "fs";
import path from "path";
import { kv } from "@vercel/kv";

export const dynamic = "force-dynamic";
export const revalidate = 0;

const dataPath = path.join(process.cwd(), "data", "rawTabs.json");
const sharedStateKey = "branch-activation-dashboard-state";
let writeQueue = Promise.resolve();
const noStoreHeaders = {
  "Cache-Control": "no-store, max-age=0, must-revalidate"
};

function getStorageMode() {
  if (process.env.KV_REST_API_URL && process.env.KV_REST_API_TOKEN) {
    return "shared-kv";
  }

  if (process.env.VERCEL) {
    return "browser-fallback";
  }

  return "local-file";
}

async function readLocalFile() {
  const file = await fs.readFile(dataPath, "utf8");
  return JSON.parse(file);
}

async function persistLocalFile(payload) {
  const tempPath = `${dataPath}.tmp`;
  const serialized = JSON.stringify(payload, null, 2);

  await fs.mkdir(path.dirname(dataPath), { recursive: true });

  let lastError;
  for (let attempt = 0; attempt < 3; attempt += 1) {
    try {
      await fs.writeFile(tempPath, serialized, "utf8");
      await fs.rename(tempPath, dataPath);
      return;
    } catch (error) {
      lastError = error;
      await fs.rm(tempPath, { force: true }).catch(() => {});
      await new Promise((resolve) => setTimeout(resolve, 120 * (attempt + 1)));
    }
  }

  throw lastError;
}

async function readRawTabs() {
  const storageMode = getStorageMode();

  if (storageMode === "shared-kv") {
    const payload = await kv.get(sharedStateKey);
    if (!payload) return null;

    if (typeof payload === "string") {
      return JSON.parse(payload);
    }

    return payload;
  }

  if (storageMode === "local-file") {
    return readLocalFile();
  }

  return null;
}

async function persistRawTabs(payload) {
  const storageMode = getStorageMode();
  const nextPayload = {
    ...payload,
    updatedAt: new Date().toISOString()
  };

  if (storageMode === "shared-kv") {
    await kv.set(sharedStateKey, JSON.stringify(nextPayload));
    return nextPayload;
  }

  if (storageMode === "local-file") {
    await persistLocalFile(nextPayload);
    return nextPayload;
  }

  throw new Error("Shared storage is not configured.");
}

export async function GET() {
  const storageMode = getStorageMode();

  try {
    const data = await readRawTabs();

    if (data) {
      return Response.json({ ...data, storageMode }, { headers: noStoreHeaders });
    }

    if (storageMode === "browser-fallback") {
      return Response.json(
        { error: "Shared storage is not configured.", storageMode },
        { status: 503, headers: noStoreHeaders }
      );
    }

    return Response.json(
      { error: "Failed to read raw tabs data.", storageMode },
      { status: 500, headers: noStoreHeaders }
    );
  } catch (error) {
    return Response.json(
      { error: "Failed to read raw tabs data.", storageMode },
      { status: 500, headers: noStoreHeaders }
    );
  }
}

export async function POST(request) {
  const storageMode = getStorageMode();

  try {
    const payload = await request.json();
    writeQueue = writeQueue
      .catch(() => {})
      .then(() => persistRawTabs(payload));

    const savedPayload = await writeQueue;
    return Response.json({ ok: true, storageMode, updatedAt: savedPayload.updatedAt }, { headers: noStoreHeaders });
  } catch (error) {
    console.error("Failed to save raw tabs data.", error);
    return Response.json(
      { error: "Failed to save raw tabs data.", storageMode },
      { status: storageMode === "browser-fallback" ? 503 : 500, headers: noStoreHeaders }
    );
  }
}
