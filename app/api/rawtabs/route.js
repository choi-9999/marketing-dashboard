import { promises as fs } from "fs";
import path from "path";
import { kv } from "@vercel/kv";

const dataPath = path.join(process.cwd(), "data", "rawTabs.json");
const sharedStateKey = "branch-activation-dashboard-state";
let writeQueue = Promise.resolve();

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
    return payload || null;
  }

  if (storageMode === "local-file") {
    return readLocalFile();
  }

  return null;
}

async function persistRawTabs(payload) {
  const storageMode = getStorageMode();

  if (storageMode === "shared-kv") {
    await kv.set(sharedStateKey, payload);
    return;
  }

  if (storageMode === "local-file") {
    await persistLocalFile(payload);
    return;
  }

  throw new Error("Shared storage is not configured.");
}

export async function GET() {
  const storageMode = getStorageMode();

  try {
    const data = await readRawTabs();

    if (data) {
      return Response.json({ ...data, storageMode });
    }

    if (storageMode === "browser-fallback") {
      return Response.json(
        { error: "Shared storage is not configured.", storageMode },
        { status: 503 }
      );
    }

    return Response.json(
      { error: "Failed to read raw tabs data.", storageMode },
      { status: 500 }
    );
  } catch (error) {
    return Response.json(
      { error: "Failed to read raw tabs data.", storageMode },
      { status: 500 }
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

    await writeQueue;
    return Response.json({ ok: true, storageMode });
  } catch (error) {
    console.error("Failed to save raw tabs data.", error);
    return Response.json(
      { error: "Failed to save raw tabs data.", storageMode },
      { status: storageMode === "browser-fallback" ? 503 : 500 }
    );
  }
}
