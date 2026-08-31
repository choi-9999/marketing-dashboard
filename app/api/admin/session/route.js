import { NextResponse } from "next/server";
import {
  ADMIN_COOKIE_NAME,
  adminSessionCookieOptions,
  createAdminSessionToken,
  isAdminAuthConfigured,
  isAdminAuthenticated,
  verifyAdminPassword
} from "../../../lib/adminAuth";

export const dynamic = "force-dynamic";

const noStoreHeaders = {
  "Cache-Control": "no-store, max-age=0, must-revalidate"
};

export async function GET() {
  return NextResponse.json(
    {
      authenticated: await isAdminAuthenticated(),
      configured: isAdminAuthConfigured()
    },
    { headers: noStoreHeaders }
  );
}

export async function POST(request) {
  if (!isAdminAuthConfigured()) {
    return NextResponse.json(
      { error: "관리자 로그인이 설정되지 않았습니다." },
      { status: 503, headers: noStoreHeaders }
    );
  }

  const payload = await request.json().catch(() => null);
  if (!verifyAdminPassword(payload?.password || "")) {
    await new Promise((resolve) => setTimeout(resolve, 350));
    return NextResponse.json(
      { error: "비밀번호가 올바르지 않습니다." },
      { status: 401, headers: noStoreHeaders }
    );
  }

  const response = NextResponse.json({ authenticated: true }, { headers: noStoreHeaders });
  response.cookies.set(ADMIN_COOKIE_NAME, createAdminSessionToken(), adminSessionCookieOptions);
  return response;
}

export async function DELETE() {
  const response = NextResponse.json({ authenticated: false }, { headers: noStoreHeaders });
  response.cookies.set(ADMIN_COOKIE_NAME, "", {
    ...adminSessionCookieOptions,
    maxAge: 0
  });
  return response;
}
