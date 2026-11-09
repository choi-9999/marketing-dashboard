import scholarshipEssays2025 from "../../data/scholarshipEssays2025.json";
import scholarshipEssays2026 from "../../data/scholarshipEssays.json";
import { isAdminAuthenticated } from "../../lib/adminAuth";

export const dynamic = "force-dynamic";

const noStoreHeaders = {
  "Cache-Control": "no-store, max-age=0, must-revalidate"
};

const scholarshipEssaysByYear = {
  "2025": scholarshipEssays2025,
  "2026": scholarshipEssays2026
};

export async function POST(request) {
  if (!(await isAdminAuthenticated())) {
    return Response.json(
      { error: "관리자 로그인이 필요합니다." },
      { status: 401, headers: noStoreHeaders }
    );
  }

  const payload = await request.json().catch(() => null);
  const year = String(payload?.year || "").trim();
  const name = String(payload?.name || "").trim();
  const branch = String(payload?.branch || "").trim();

  if (!year || !name) {
    return Response.json(
      { error: "연도와 학생 이름이 필요합니다." },
      { status: 400, headers: noStoreHeaders }
    );
  }

  const yearEssays = scholarshipEssaysByYear[year] || [];
  let essay = yearEssays.find(
    (item) =>
      item.name === name &&
      branch &&
      (item.branch.includes(branch) || branch.includes(item.branch))
  );

  if (!essay) essay = yearEssays.find((item) => item.name === name);

  if (!essay) {
    return Response.json(
      { error: "등록된 합격 수기가 없습니다." },
      { status: 404, headers: noStoreHeaders }
    );
  }

  return Response.json({ essay: { ...essay, year } }, { headers: noStoreHeaders });
}
