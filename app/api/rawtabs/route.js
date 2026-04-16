import { promises as fs } from "fs";
import path from "path";

const dataPath = path.join(process.cwd(), "data", "rawTabs.json");
let writeQueue = Promise.resolve();

async function readRawTabs() {
  const file = await fs.readFile(dataPath, "utf8");
  return JSON.parse(file);
}

async function persistRawTabs(payload) {
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

export async function GET() {
  try {
    const data = await readRawTabs();
    return Response.json(data);
  } catch (error) {
    return Response.json(
      { error: "Failed to read raw tabs data." },
      { status: 500 }
    );
  }
}

export async function POST(request) {
  try {
    const payload = await request.json();
    writeQueue = writeQueue
      .catch(() => {})
      .then(() => persistRawTabs(payload));

    await writeQueue;
    return Response.json({ ok: true });
  } catch (error) {
    console.error("Failed to save raw tabs data.", error);
    return Response.json(
      { error: "Failed to save raw tabs data." },
      { status: 500 }
    );
  }
}
