import { promises as fs } from "fs";
import path from "path";
import { isAdminAuthenticated } from "../../lib/adminAuth";

export const dynamic = "force-dynamic";

const downloadableFiles = {
  "scholarship-essays-2025.xlsx": path.join(
    process.cwd(),
    "data",
    "downloads",
    "scholarship-essays-2025.xlsx"
  ),
  "scholarship-essays-2026.xlsx": path.join(
    process.cwd(),
    "data",
    "downloads",
    "scholarship-essays-2026.xlsx"
  )
};

const noStoreHeaders = {
  "Cache-Control": "private, no-store, max-age=0, must-revalidate"
};

export async function GET(_request, context) {
  if (!(await isAdminAuthenticated())) {
    return Response.json(
      { error: "관리자 로그인이 필요합니다." },
      { status: 401, headers: noStoreHeaders }
    );
  }

  const { file } = await context.params;
  const filePath = downloadableFiles[file];
  if (!filePath) {
    return Response.json(
      { error: "파일을 찾을 수 없습니다." },
      { status: 404, headers: noStoreHeaders }
    );
  }

  try {
    const contents = await fs.readFile(filePath);

    return new Response(contents, {
      headers: {
        ...noStoreHeaders,
        "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": `attachment; filename="${file}"`,
        "X-Content-Type-Options": "nosniff"
      }
    });
  } catch {
    return Response.json(
      { error: "파일을 찾을 수 없습니다." },
      { status: 404, headers: noStoreHeaders }
    );
  }
}
