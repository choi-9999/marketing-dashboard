import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const query = searchParams.get("query");

  if (!query || !query.trim()) {
    return NextResponse.json({ success: false, error: "query is required" }, { status: 400 });
  }

  const clientId =
    process.env.NAVER_MAPS_CLIENT_ID ||
    process.env.NAVER_MAP_CLIENT_ID ||
    process.env.NEXT_PUBLIC_NAVER_MAP_CLIENT_ID ||
    process.env.NEXT_PUBLIC_NAVER_CLIENT_ID;
  const clientSecret =
    process.env.NAVER_MAPS_CLIENT_SECRET ||
    process.env.NAVER_MAP_CLIENT_SECRET ||
    process.env.NAVER_CLOUD_CLIENT_SECRET;

  if (!clientId || !clientSecret) {
    return NextResponse.json(
      {
        success: false,
        error: "NAVER map geocoding credentials are missing",
        requiredEnv: ["NEXT_PUBLIC_NAVER_MAP_CLIENT_ID", "NAVER_MAPS_CLIENT_SECRET"]
      },
      { status: 500 }
    );
  }

  try {
    const url = new URL("https://maps.apigw.ntruss.com/map-geocode/v2/geocode");
    url.searchParams.set("query", query.trim());

    const response = await fetch(url, {
      headers: {
        "X-NCP-APIGW-API-KEY-ID": clientId,
        "X-NCP-APIGW-API-KEY": clientSecret
      },
      cache: "no-store"
    });

    const data = await response.json().catch(() => ({}));

    if (!response.ok) {
      return NextResponse.json(
        {
          success: false,
          error: data?.errorMessage || data?.message || `NAVER geocode failed with ${response.status}`,
          status: response.status
        },
        { status: response.status }
      );
    }

    const firstAddress = data?.addresses?.[0];
    const lat = Number(firstAddress?.y);
    const lng = Number(firstAddress?.x);

    if (!Number.isFinite(lat) || !Number.isFinite(lng)) {
      return NextResponse.json(
        { success: false, error: "No geocode result", rawStatus: data?.status },
        { status: 404 }
      );
    }

    return NextResponse.json({
      success: true,
      coords: [lat, lng],
      address: firstAddress?.roadAddress || firstAddress?.jibunAddress || query.trim()
    });
  } catch (error) {
    return NextResponse.json(
      { success: false, error: error?.message || "NAVER geocode request failed" },
      { status: 500 }
    );
  }
}
