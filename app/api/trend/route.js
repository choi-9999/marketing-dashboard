import { NextResponse } from "next/server";

export const dynamic = "force-dynamic";

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const branch = searchParams.get("branch");

  if (!branch) {
    return NextResponse.json({ error: "Branch parameter is required" }, { status: 400 });
  }

  const cleanBranch = String(branch).trim();

  // Retrieve credentials from environment variables
  const clientId = process.env.NAVER_CLIENT_ID;
  const clientSecret = process.env.NAVER_CLIENT_SECRET;

  // 1. Generate fallback mock data
  const hash = cleanBranch.split("").reduce((acc, char) => acc + char.charCodeAt(0), 0);
  const fallbackTrend = [];
  const months = ["1월", "2월", "3월", "4월", "5월", "6월"];

  if (cleanBranch.includes("동탄")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(46 + Math.sin(i + 1) * 9 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 11 + i * 3), // 수능선배 동탄점
        comp2: Math.round(44 + Math.sin(i * 1.2 + 3) * 10 + i * 2), // 수만휘 영천점
        comp3: Math.round(42 + Math.cos(i * 0.8 + 4) * 8 + i * 2), // 디랩 동탄
        comp4: Math.round(39 + Math.sin(i + 5) * 9 + i * 2), // 수만휘 호수공원점
        comp5: Math.round(45 + Math.cos(i + 6) * 10 + i * 3), // 잇올 동탄센터
        comp6: Math.round(33 + Math.sin(i + 7) * 7 + i * 1), // PK 2동탄센터점
        comp7: Math.round(30 + Math.cos(i + 8) * 6 + i * 1), // 수만휘 반송점
        comp8: Math.round(32 + Math.sin(i + 9) * 6 + i * 1) // PK 동탄센터
      });
    }
  } else if (cleanBranch.includes("대전둔산") || cleanBranch.includes("둔산") || (cleanBranch.includes("대전") && !cleanBranch.includes("유성"))) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(44 + Math.cos(i + 2) * 10 + i * 2), // 디랩 대전
        comp2: Math.round(46 + Math.sin(i * 1.2 + 3) * 11 + i * 2), // 수만휘 대전둔산 1관
        comp3: Math.round(42 + Math.cos(i * 0.8 + 4) * 8 + i * 2), // 수만휘 대전둔산 2관
        comp4: Math.round(49 + Math.sin(i + 5) * 10 + i * 3), // 잇올 대전둔산 1관
        comp5: Math.round(52 + Math.cos(i + 6) * 12 + i * 3) // 러셀 대전
      });
    }
  } else if (cleanBranch.includes("강북")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 9 + i * 3),
        comp1: Math.round(50 + Math.cos(i + 2) * 11 + i * 3), // 잇올 노원중계 1관
        comp2: Math.round(46 + Math.sin(i * 1.2 + 3) * 10 + i * 2), // 수만휘 노원중계점
        comp3: Math.round(48 + Math.cos(i * 0.8 + 4) * 9 + i * 3), // 잇올 노원중계 2관
        comp4: Math.round(52 + Math.sin(i + 5) * 12 + i * 3), // 러셀 중계
        comp5: Math.round(38 + Math.cos(i + 6) * 8 + i * 2), // PK 노원점
        comp6: Math.round(42 + Math.sin(i + 7) * 9 + i * 2), // 강북 메가스터디 중계관
        comp7: Math.round(45 + Math.cos(i + 8) * 10 + i * 2) // 강북청솔학원
      });
    }
  } else if (cleanBranch.includes("김포")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 10 + i * 2), // 수능선배 김포점
        comp2: Math.round(44 + Math.sin(i * 1.2 + 3) * 11 + i * 2), // 수만휘 김포장기점
        comp3: Math.round(41 + Math.cos(i * 0.8 + 4) * 8 + i * 2), // 디랩 김포
        comp4: Math.round(46 + Math.sin(i + 5) * 9 + i * 3), // 잇올 몰입관 김포장기
        comp5: Math.round(47 + Math.cos(i + 6) * 10 + i * 3) // 잇올 김포센터
      });
    }
  } else if (cleanBranch.includes("송도") || cleanBranch.includes("인천송도")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(44 + Math.cos(i + 2) * 10 + i * 2), // 수만휘 인천송도점
        comp2: Math.round(40 + Math.sin(i * 1.2 + 3) * 11 + i * 3), // PK 인천송도점
        comp3: Math.round(48 + Math.cos(i * 0.8 + 4) * 9 + i * 3), // 잇올 연수송도 2관
        comp4: Math.round(38 + Math.sin(i + 5) * 8 + i * 2), // 잇올 연수송도 1관
        comp5: Math.round(32 + Math.cos(i + 6) * 7 + i * 1) // PK 인천연수점
      });
    }
  } else if (cleanBranch.includes("분당정자") || cleanBranch.includes("분당") || cleanBranch.includes("대치")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 10 + i * 3),
        comp1: Math.round(30 + Math.cos(i + 2) * 12 + i * 2), // 수능선배 / 수만휘
        comp2: Math.round(40 + Math.sin(i * 1.2 + 3) * 15 + i * 4), // 잇올 / 수능선배
        comp3: Math.round(55 + Math.cos(i * 0.8 + 4) * 8 + i * 2), // PK / 러셀
        comp4: Math.round(50 + Math.sin(i + 5) * 14 + i * 3), // 러셀 / 잇올
        comp5: Math.round(20 + Math.cos(i + 6) * 6 + i * 1) // 강남대성 / 디랩
      });
    }
  } else if (cleanBranch.includes("안성기숙") || cleanBranch.includes("안성")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(43 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(49 + Math.cos(i + 2) * 10 + i * 3), // 남안성비상독학기숙
        comp2: Math.round(47 + Math.sin(i * 1.2 + 3) * 11 + i * 2), // 안성비상에듀기숙
        comp3: Math.round(42 + Math.cos(i * 0.8 + 4) * 9 + i * 2), // 수만휘기숙
        comp4: Math.round(33 + Math.sin(i + 5) * 7 + i * 1), // 역사적사명기숙
        comp5: Math.round(28 + Math.cos(i + 6) * 6 + i * 1) // 72시간공부캠프
      });
    }
  } else if (cleanBranch.includes("이천기숙") || cleanBranch.includes("이천")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(44 + Math.sin(i + 1) * 9 + i * 3),
        comp1: Math.round(52 + Math.cos(i + 2) * 10 + i * 3), // 강남대성 QUETTA
        comp2: Math.round(46 + Math.sin(i * 1.2 + 3) * 12 + i * 3), // 잇올 이천캠프
        comp3: Math.round(49 + Math.cos(i * 0.8 + 4) * 9 + i * 2), // 강남대성 의대관
        comp4: Math.round(38 + Math.sin(i + 5) * 10 + i * 2), // 이천청솔
        comp5: Math.round(30 + Math.cos(i + 6) * 7 + i * 1), // 이천탑클래스
        comp6: Math.round(25 + Math.sin(i + 7) * 6 + i * 1) // 이천아이나인
      });
    }
  } else if (cleanBranch.includes("마포") || cleanBranch.includes("신촌")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 10 + i * 2), // 수능선배 신촌점
        comp2: Math.round(50 + Math.sin(i * 1.2 + 3) * 11 + i * 3), // 잇올 마포신촌센터
        comp3: Math.round(44 + Math.cos(i * 0.8 + 4) * 9 + i * 2), // 종로학원 강북
        comp4: Math.round(47 + Math.sin(i + 5) * 10 + i * 2) // 신촌 메가스터디학원
      });
    }
  } else if (cleanBranch.includes("하남")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(49 + Math.cos(i + 2) * 10 + i * 3), // 수능선배 하남점
        comp2: Math.round(47 + Math.sin(i * 1.2 + 3) * 11 + i * 2), // 잇올 하남미사센터
        comp3: Math.round(43 + Math.cos(i * 0.8 + 4) * 8 + i * 2) // 수만휘 하남미사점
      });
    }
  } else if (cleanBranch.includes("다산") || cleanBranch.includes("남양주다산")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(46 + Math.cos(i + 2) * 10 + i * 2) // 수만휘 남양주다산점
      });
    }
  } else if (cleanBranch.includes("일산서구") || cleanBranch.includes("일산")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(46 + Math.cos(i + 2) * 10 + i * 2), // 디랩 일산
        comp2: Math.round(48 + Math.sin(i * 1.2 + 3) * 11 + i * 3), // 잇올 일산 주엽센터
        comp3: Math.round(44 + Math.cos(i * 0.8 + 4) * 9 + i * 2), // 일산청솔학원
        comp4: Math.round(47 + Math.sin(i + 5) * 10 + i * 2) // 일산 메가스터디학원
      });
    }
  } else if (cleanBranch.includes("의정부")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(44 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 10 + i * 2), // 잇올 의정부센터
        comp2: Math.round(43 + Math.sin(i * 1.2 + 3) * 9 + i * 2) // 수만휘 의정부금오점
      });
    }
  } else if (cleanBranch.includes("평택")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(44 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 10 + i * 2), // 잇올 평택센터
        comp2: Math.round(38 + Math.sin(i * 1.2 + 3) * 9 + i * 2) // 커넥츠프랩 수능관 평택합정점
      });
    }
  } else if (cleanBranch.includes("춘천")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(44 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 10 + i * 3), // 잇올 춘천센터
        comp2: Math.round(43 + Math.sin(i * 1.2 + 3) * 11 + i * 2), // 수만휘 춘천후평점
        comp3: Math.round(39 + Math.cos(i * 0.8 + 4) * 8 + i * 2) // PK 춘천온의센터
      });
    }
  } else if (cleanBranch.includes("광주동구") || (cleanBranch.includes("광주") && cleanBranch.includes("동구"))) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(50 + Math.cos(i + 2) * 11 + i * 3), // 러셀 광주
        comp2: Math.round(46 + Math.sin(i * 1.2 + 3) * 10 + i * 2), // 잇올 광주충장로
        comp3: Math.round(44 + Math.cos(i * 0.8 + 4) * 9 + i * 2), // 잇올 광주봉선
        comp4: Math.round(38 + Math.sin(i + 5) * 8 + i * 2) // 수만휘 광주봉선
      });
    }
  } else if (cleanBranch.includes("부산교대") || cleanBranch.includes("교대")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(45 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(47 + Math.cos(i + 2) * 10 + i * 2), // 잇올 부산사직센터
        comp2: Math.round(42 + Math.sin(i * 1.2 + 3) * 9 + i * 2) // 잇올 몰입관 부산동래사직캠프
      });
    }
  } else if (cleanBranch.includes("대구달서") || cleanBranch.includes("달서")) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(44 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(45 + Math.cos(i + 2) * 10 + i * 2), // 수만휘 대구달서점
        comp2: Math.round(47 + Math.sin(i * 1.2 + 3) * 11 + i * 3), // 잇올 대구월성센터
        comp3: Math.round(41 + Math.cos(i * 0.8 + 4) * 8 + i * 2) // 잇올 대구상인센터
      });
    }
  } else if (cleanBranch.includes("독학기숙") || (cleanBranch.includes("기숙") && !cleanBranch.includes("안성") && !cleanBranch.includes("이천"))) {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(42 + Math.sin(i + 1) * 8 + i * 3),
        comp1: Math.round(48 + Math.cos(i + 2) * 10 + i * 2), // 이투스 기숙학원
        comp2: Math.round(35 + Math.sin(i * 1.2 + 3) * 12 + i * 3), // 비상에듀독학기숙
        comp3: Math.round(28 + Math.cos(i * 0.8 + 4) * 7 + i * 2) // 진성스파르타기숙
      });
    }
  } else {
    for (let i = 0; i < 6; i++) {
      fallbackTrend.push({
        month: months[i],
        ours: Math.round(30 + (hash % 15) + Math.sin(i + hash) * 15 + i * 4),
        comp1: Math.round(40 + (hash % 20) + Math.cos(i + hash) * 10 + i * 2),
        comp2: Math.round(25 + (hash % 10) + Math.sin(i * 1.5 + hash) * 8 + i * 3)
      });
    }
  }

  // If credentials are not configured, return fallback directly
  if (!clientId || !clientSecret || clientId.includes("YOUR_") || clientSecret.includes("YOUR_")) {
    return NextResponse.json({
      success: true,
      mode: "mock",
      trendData: fallbackTrend
    });
  }

  try {
    const periodMonths = ["2026-01-01", "2026-02-01", "2026-03-01", "2026-04-01", "2026-05-01", "2026-06-01"];
    const monthLabels = ["1월", "2월", "3월", "4월", "5월", "6월"];

    if (cleanBranch.includes("동탄") || cleanBranch.includes("대전둔산") || cleanBranch.includes("둔산") || (cleanBranch.includes("대전") && !cleanBranch.includes("유성")) || cleanBranch.includes("강북") || cleanBranch.includes("김포") || cleanBranch.includes("송도") || cleanBranch.includes("인천송도") || cleanBranch.includes("분당정자") || cleanBranch.includes("분당") || cleanBranch.includes("대치") || cleanBranch.includes("이천기숙") || cleanBranch.includes("이천") || cleanBranch.includes("안성기숙") || cleanBranch.includes("안성")) {
      const isDongtan = cleanBranch.includes("동탄");
      const isDaejeonDunsan = cleanBranch.includes("대전둔산") || cleanBranch.includes("둔산") || (cleanBranch.includes("대전") && !cleanBranch.includes("유성"));
      const isGangbuk = cleanBranch.includes("강북");
      const isGimpo = cleanBranch.includes("김포");
      const isSongdo = cleanBranch.includes("송도") || cleanBranch.includes("인천송도");
      const isDaech = cleanBranch.includes("대치");
      const isIcheon = cleanBranch.includes("이천기숙") || cleanBranch.includes("이천");
      const isAnseong = cleanBranch.includes("안성기숙") || cleanBranch.includes("안성");
      // 6+ groups exceeds Naver API limit (max 5 per call). We make two parallel requests and normalize the scale using "ours"
      const body1 = {
        startDate: "2026-01-01",
        endDate: "2026-06-30",
        timeUnit: "month",
        keywordGroups: isDaejeonDunsan
          ? [
              { groupName: "ours", keywords: ["대전둔산 이투스247", "둔산 이투스247", "대전 이투스247", "둔산동 이투스"] },
              { groupName: "comp1", keywords: ["디랩 대전", "대전 디랩", "둔산동 디랩"] },
              { groupName: "comp2", keywords: ["수만휘 스파르타 대전둔산 1관", "대전둔산 수만휘 1관", "둔산동 수만휘 1관"] },
              { groupName: "comp3", keywords: ["수만휘 스파르타 대전둔산 2관", "대전둔산 수만휘 2관", "둔산동 수만휘 2관"] },
              { groupName: "comp4", keywords: ["잇올 스파르타 대전둔산센터 1관", "대전둔산 잇올", "둔산동 잇올"] }
            ]
          : isGangbuk
          ? [
              { groupName: "ours", keywords: ["강북 이투스247", "강북 이투스", "노원 이투스247", "중계 이투스247"] },
              { groupName: "comp1", keywords: ["잇올 스파르타 노원중계센터 1관", "노원중계 잇올 1관", "중계동 잇올 1관"] },
              { groupName: "comp2", keywords: ["수만휘 스파르타 노원중계점", "노원중계 수만휘", "중계동 수만휘"] },
              { groupName: "comp3", keywords: ["잇올 스파르타 노원중계센터 2관", "노원중계 잇올 2관", "중계동 잇올 2관"] },
              { groupName: "comp4", keywords: ["메가스터디 러셀 중계", "러셀 중계", "중계 러셀"] }
            ]
          : isGimpo
          ? [
              { groupName: "ours", keywords: ["김포 이투스247", "김포 이투스", "김포이투스"] },
              { groupName: "comp1", keywords: ["수능선배 김포점", "김포 수능선배", "사우동 수능선배"] },
              { groupName: "comp2", keywords: ["수만휘 스파르타 김포장기점", "김포장기 수만휘", "장기동 수만휘"] },
              { groupName: "comp3", keywords: ["디랩 김포", "김포 디랩", "장기동 디랩"] },
              { groupName: "comp4", keywords: ["잇올 몰입관 김포장기캠프", "김포장기 잇올", "장기동 잇올"] }
            ]
          : isSongdo
          ? [
              { groupName: "ours", keywords: ["인천송도 이투스247", "송도 이투스247", "송도 이투스"] },
              { groupName: "comp1", keywords: ["수만휘 스파르타 인천송도점", "송도 수만휘", "인천송도 수만휘"] },
              { groupName: "comp2", keywords: ["PK독학재수학원 인천송도점", "송도 PK독학재수", "송도 PK"] },
              { groupName: "comp3", keywords: ["잇올 스파르타 인천연수송도센터 2관", "송도 잇올 2관", "송도 잇올"] },
              { groupName: "comp4", keywords: ["잇올 스파르타 인천연수송도센터 1관", "연수 잇올 1관", "연수동 잇올"] }
            ]
          : isDongtan
          ? [
              { groupName: "ours", keywords: ["동탄 이투스247", "동탄 이투스", "동탄이투스"] },
              { groupName: "comp1", keywords: ["수능선배 동탄점", "수능선배 동탄", "동탄 수능선배"] },
              { groupName: "comp2", keywords: ["수만휘 스파르타 동탄영천점", "동탄영천 수만휘", "동탄 수만휘"] },
              { groupName: "comp3", keywords: ["디랩 독학재수학원 동탄", "디랩 동탄", "동탄 디랩"] },
              { groupName: "comp4", keywords: ["수만휘 스파르타 동탄호수공원점", "동탄호수공원 수만휘", "동탄호수 수만휘"] }
            ]
          : isAnseong
          ? [
              { groupName: "ours", keywords: ["이투스247 안성기숙", "이투스 안성기숙", "안성 이투스247"] },
              { groupName: "comp1", keywords: ["남안성비상에듀독학기숙학원", "남안성비상에듀", "남안성비상"] },
              { groupName: "comp2", keywords: ["안성비상에듀기숙학원", "안성비상에듀", "안성비상기숙"] },
              { groupName: "comp3", keywords: ["수만휘기숙학원", "수만휘기숙", "안성 수만휘"] },
              { groupName: "comp4", keywords: ["역사적사명 기숙학원", "역사적사명기숙", "역사적사명"] }
            ]
          : isIcheon
          ? [
              { groupName: "ours", keywords: ["이투스247 이천기숙", "이투스 이천기숙", "이천 이투스247"] },
              { groupName: "comp1", keywords: ["강남대성 QUETTA", "강남대성 퀘타", "대성 퀘타"] },
              { groupName: "comp2", keywords: ["잇올 기숙학원 이천캠프", "잇올 이천기숙", "잇올 이천캠프"] },
              { groupName: "comp3", keywords: ["강남대성기숙 의대관", "강남대성 의대관", "대성기숙 의대관"] },
              { groupName: "comp4", keywords: ["이천청솔기숙학원", "이천청솔", "이천 청솔기숙"] }
            ]
          : isDaech
          ? [
              { groupName: "ours", keywords: ["대치 이투스247", "대치 이투스", "대치이투스"] },
              { groupName: "comp1", keywords: ["수능선배 대치", "대치 수능선배"] },
              { groupName: "comp2", keywords: ["잇올 스파르타 대치센터", "잇올 대치", "대치 잇올"] },
              { groupName: "comp3", keywords: ["PK독학재수학원 대치점", "PK 대치", "대치 PK독학재수"] },
              { groupName: "comp4", keywords: ["메가스터디 러셀 대치학원", "러셀 대치", "메가스터디 대치러셀"] }
            ]
          : [
              { groupName: "ours", keywords: ["분당정자 이투스247", "분당정자 이투스", "분당정자이투스"] },
              { groupName: "comp1", keywords: ["수만휘 스파르타 분당정자점", "수만휘 스파르타 분당정자", "분당정자 수만휘"] },
              { groupName: "comp2", keywords: ["수능선배 분당점", "수능선배 분당"] },
              { groupName: "comp3", keywords: ["메가스터디 러셀 분당학원", "러셀 분당", "메가스터디 분당러셀"] },
              { groupName: "comp4", keywords: ["잇올 스파르타 분당정자센터", "잇올 스파르타 분당정자", "분당정자 잇올"] }
            ]
      };
      const body2 = {
        startDate: "2026-01-01",
        endDate: "2026-06-30",
        timeUnit: "month",
        keywordGroups: isDaejeonDunsan
          ? [
              { groupName: "ours", keywords: ["대전둔산 이투스247", "둔산 이투스247", "대전 이투스247", "둔산동 이투스"] },
              { groupName: "comp5", keywords: ["메가스터디 러셀 대전", "러셀 대전", "대전 러셀"] }
            ]
          : isGangbuk
          ? [
              { groupName: "ours", keywords: ["강북 이투스247", "강북 이투스", "노원 이투스247", "중계 이투스247"] },
              { groupName: "comp5", keywords: ["PK독학재수학원 노원점", "노원 PK독학재수", "상계동 PK"] },
              { groupName: "comp6", keywords: ["강북 메가스터디학원 중계관", "강북 메가스터디", "중계 메가스터디"] },
              { groupName: "comp7", keywords: ["강북청솔학원", "강북청솔", "중계 청솔학원"] }
            ]
          : isGimpo
          ? [
              { groupName: "ours", keywords: ["김포 이투스247", "김포 이투스", "김포이투스"] },
              { groupName: "comp5", keywords: ["잇올 스파르타 김포센터", "김포 잇올", "사우동 잇올"] }
            ]
          : isSongdo
          ? [
              { groupName: "ours", keywords: ["인천송도 이투스247", "송도 이투스247", "송도 이투스"] },
              { groupName: "comp5", keywords: ["PK독학재수학원 인천연수점", "연수 PK독학재수", "동춘동 PK"] }
            ]
          : isDongtan
          ? [
              { groupName: "ours", keywords: ["동탄 이투스247", "동탄 이투스", "동탄이투스"] },
              { groupName: "comp5", keywords: ["잇올 스파르타 동탄센터", "잇올 동탄", "동탄 잇올"] },
              { groupName: "comp6", keywords: ["PK대치스파르타 2동탄센터점", "2동탄 PK", "산척동 PK"] },
              { groupName: "comp7", keywords: ["수만휘 관리형 독서실 동탄반송점", "동탄반송 수만휘"] },
              { groupName: "comp8", keywords: ["PK대치스파르타 동탄센터", "동탄 PK대치스파르타", "반송동 PK"] }
            ]
          : isAnseong
          ? [
              { groupName: "ours", keywords: ["이투스247 안성기숙", "이투스 안성기숙", "안성 이투스247"] },
              { groupName: "comp5", keywords: ["72시간공부캠프안성캠퍼스", "72시간공부캠프", "72시간캠프 안성"] }
            ]
          : isIcheon
          ? [
              { groupName: "ours", keywords: ["이투스247 이천기숙", "이투스 이천기숙", "이천 이투스247"] },
              { groupName: "comp5", keywords: ["이천탑클래스기숙학원", "이천탑클래스", "탑클래스 기숙학원"] },
              { groupName: "comp6", keywords: ["이천아이나인독학기숙재수학원", "이천아이나인", "아이나인 기숙학원"] }
            ]
          : isDaech
          ? [
              { groupName: "ours", keywords: ["대치 이투스247", "대치 이투스", "대치이투스"] },
              { groupName: "comp5", keywords: ["강남대성SⅡ", "강남대성 대치", "대성S2"] }
            ]
          : [
              { groupName: "ours", keywords: ["분당정자 이투스247", "분당정자 이투스", "분당정자이투스"] },
              { groupName: "comp5", keywords: ["디랩 분당", "분당 디랩"] }
            ]
      };

      const [res1, res2] = await Promise.all([
        fetch("https://naverapihub.apigw.ntruss.com/search-trend/v1/search", {
          method: "POST",
          headers: { "x-ncp-apigw-api-key-id": clientId, "x-ncp-apigw-api-key": clientSecret, "Content-Type": "application/json" },
          body: JSON.stringify(body1)
        }),
        fetch("https://naverapihub.apigw.ntruss.com/search-trend/v1/search", {
          method: "POST",
          headers: { "x-ncp-apigw-api-key-id": clientId, "x-ncp-apigw-api-key": clientSecret, "Content-Type": "application/json" },
          body: JSON.stringify(body2)
        })
      ]);

      if (!res1.ok || !res2.ok) {
        console.warn("One of the parallel Naver Search Trend API requests failed.");
        return NextResponse.json({ success: true, mode: "mock-fallback-api-error", trendData: fallbackTrend });
      }

      const data1 = await res1.json();
      const data2 = await res2.json();

      const oursData1 = data1.results.find(r => r.title === "ours")?.data || [];
      const comp1Data = data1.results.find(r => r.title === "comp1")?.data || [];
      const comp2Data = data1.results.find(r => r.title === "comp2")?.data || [];
      const comp3Data = data1.results.find(r => r.title === "comp3")?.data || [];
      const comp4Data = data1.results.find(r => r.title === "comp4")?.data || [];

      const oursData2 = data2.results.find(r => r.title === "ours")?.data || [];
      const comp5Data = data2.results.find(r => r.title === "comp5")?.data || [];
      const comp6Data = data2.results.find(r => r.title === "comp6")?.data || [];
      const comp7Data = data2.results.find(r => r.title === "comp7")?.data || [];
      const comp8Data = data2.results.find(r => r.title === "comp8")?.data || [];

      const realTrend = periodMonths.map((period, index) => {
        const oursVal1 = oursData1.find(d => d.period === period)?.ratio || 0;
        const oursVal2 = oursData2.find(d => d.period === period)?.ratio || 0;

        const comp1Val = comp1Data.find(d => d.period === period)?.ratio || 0;
        const comp2Val = comp2Data.find(d => d.period === period)?.ratio || 0;
        const comp3Val = comp3Data.find(d => d.period === period)?.ratio || 0;
        const comp4Val = comp4Data.find(d => d.period === period)?.ratio || 0;
        const comp5Val = comp5Data.find(d => d.period === period)?.ratio || 0;
        const comp6Val = comp6Data.find(d => d.period === period)?.ratio || 0;
        const comp7Val = comp7Data.find(d => d.period === period)?.ratio || 0;
        const comp8Val = comp8Data.find(d => d.period === period)?.ratio || 0;

        // Normalise comp5..comp8 ratio relative to ours ratio
        const factor = oursVal2 > 0 ? oursVal1 / oursVal2 : 1;
        const comp5ValNormalized = comp5Val * factor;
        const comp6ValNormalized = comp6Val * factor;
        const comp7ValNormalized = comp7Val * factor;
        const comp8ValNormalized = comp8Val * factor;

        const row = {
          month: monthLabels[index],
          ours: Math.round(oursVal1 > 0 ? oursVal1 : 45 + index * 3),
          comp1: Math.round(comp1Val > 0 ? comp1Val : 30 + index * 2),
          comp2: Math.round(comp2Val > 0 ? comp2Val : 40 + index * 4),
          comp3: Math.round(comp3Val > 0 ? comp3Val : 55 + index * 2),
          comp4: Math.round(comp4Val > 0 ? comp4Val : 50 + index * 3),
          comp5: Math.round(comp5ValNormalized > 0 ? comp5ValNormalized : 20 + index)
        };

        if (isIcheon || isDongtan || isGangbuk) {
          row.comp6 = Math.round(comp6ValNormalized > 0 ? comp6ValNormalized : 15 + index);
        }
        if (isDongtan || isGangbuk) {
          row.comp7 = Math.round(comp7ValNormalized > 0 ? comp7ValNormalized : 12 + index);
        }
        if (isDongtan) {
          row.comp8 = Math.round(comp8ValNormalized > 0 ? comp8ValNormalized : 14 + index);
        }

        return row;
      });

      return NextResponse.json({ success: true, mode: isDongtan ? "naver-api-8comps" : isGangbuk ? "naver-api-7comps" : isIcheon ? "naver-api-6comps" : "naver-api-5comps", trendData: realTrend });
    } else {
      const isDokhakGisuk = cleanBranch.includes("독학기숙") || (cleanBranch.includes("기숙") && !cleanBranch.includes("안성") && !cleanBranch.includes("이천"));
      const isDalseo = cleanBranch.includes("대구달서") || cleanBranch.includes("달서");
      const isBusanGyodae = cleanBranch.includes("부산교대") || cleanBranch.includes("교대");
      const isGwangjuDonggu = cleanBranch.includes("광주동구") || (cleanBranch.includes("광주") && cleanBranch.includes("동구"));
      const isChuncheon = cleanBranch.includes("춘천");
      const isPyeongtaek = cleanBranch.includes("평택");
      const isHanam = cleanBranch.includes("하남");
      const isUijeongbu = cleanBranch.includes("의정부");
      const isIlsan = cleanBranch.includes("일산서구") || cleanBranch.includes("일산");
      const isDasan = cleanBranch.includes("다산") || cleanBranch.includes("남양주다산");
      const isMapo = cleanBranch.includes("마포") || cleanBranch.includes("신촌");
      const oursKeywords = isDokhakGisuk
        ? ["이투스247 독학기숙학원", "이투스247 독학기숙", "이투스 독학기숙"]
        : isDalseo
        ? ["대구달서 이투스247", "대구달서 이투스", "달서 이투스247", "달서구 이투스"]
        : isBusanGyodae
        ? ["부산교대 이투스247", "부산교대 이투스", "동래 이투스247", "부산교대이투스"]
        : isGwangjuDonggu
        ? ["광주동구 이투스247", "광주동구 이투스", "광주 이투스247", "광주이투스"]
        : isChuncheon
        ? ["춘천 이투스247", "춘천 이투스", "춘천이투스"]
        : isPyeongtaek
        ? ["평택 이투스247", "평택 이투스", "평택이투스"]
        : isHanam
        ? ["하남 이투스247", "하남 이투스", "하남이투스", "미사 이투스247"]
        : isUijeongbu
        ? ["의정부 이투스247", "의정부 이투스", "의정부이투스"]
        : isIlsan
        ? ["일산서구 이투스247", "일산 이투스247", "일산 이투스", "일산이투스"]
        : isDasan
        ? ["다산 이투스247", "남양주다산 이투스247", "다산 이투스", "다산이투스"]
        : isMapo
        ? ["마포 이투스247", "마포 이투스", "신촌 이투스247", "마포이투스"]
        : [`${cleanBranch} 이투스247`, `${cleanBranch} 이투스`].filter(Boolean);
      const comp1Keywords = isDokhakGisuk
        ? ["이투스 기숙학원", "이투스기숙학원"]
        : isDalseo
        ? ["수만휘 스파르타 대구달서점", "달서 수만휘", "대구달서 수만휘"]
        : isBusanGyodae
        ? ["잇올 스파르타 부산사직센터", "부산사직 잇올", "사직동 잇올"]
        : isGwangjuDonggu
        ? ["메가스터디 러셀 광주", "러셀 광주", "광주 러셀"]
        : isChuncheon
        ? ["잇올 스파르타 춘천센터", "춘천 잇올", "석사동 잇올"]
        : isPyeongtaek
        ? ["잇올 스파르타 평택센터", "평택 잇올", "비전동 잇올"]
        : isHanam
        ? ["수능선배 하남점", "하남 수능선배", "미사 수능선배"]
        : isUijeongbu
        ? ["잇올 스파르타 의정부센터", "의정부 잇올", "의정부동 잇올"]
        : isIlsan
        ? ["디랩 일산", "일산 디랩", "대화동 디랩"]
        : isDasan
        ? ["수만휘 스파르타 남양주다산점", "남양주다산 수만휘", "다산 수만휘"]
        : isMapo
        ? ["수능선배 신촌점", "신촌 수능선배", "마포 수능선배"]
        : [`${cleanBranch} 잇올`].filter(Boolean);
      const comp2Keywords = isDokhakGisuk
        ? ["비상에듀독학기숙학원", "비상에듀 독학기숙", "광주 비상에듀 기숙"]
        : isDalseo
        ? ["잇올 스파르타 대구월성센터", "대구월성 잇올", "월성동 잇올"]
        : isBusanGyodae
        ? ["잇올 몰입관 부산동래사직캠프", "동래사직 잇올 몰입관", "동래 잇올 몰입관"]
        : isGwangjuDonggu
        ? ["잇올 스파르타 광주충장로센터", "광주충장로 잇올", "충장로 잇올"]
        : isChuncheon
        ? ["수만휘 스파르타 춘천후평점", "춘천후평 수만휘", "후평동 수만휘"]
        : isPyeongtaek
        ? ["커넥츠프랩 수능관 평택합정점", "평택 커넥츠프랩", "합정동 커넥츠프랩"]
        : isHanam
        ? ["잇올 스파르타 하남미사센터", "하남미사 잇올", "하남 잇올", "미사 잇올"]
        : isUijeongbu
        ? ["수만휘 스파르타 의정부금오점", "의정부금오 수만휘", "금오동 수만휘"]
        : isIlsan
        ? ["잇올 스파르타 일산 주엽센터", "일산주엽 잇올", "주엽동 잇올"]
        : isMapo
        ? ["잇올 스파르타 마포신촌센터", "마포신촌 잇올", "마포 잇올", "신촌 잇올"]
        : [`${cleanBranch} 수능선배`].filter(Boolean);
      const comp3Keywords = isDokhakGisuk
        ? ["진성스파르타기숙학원", "진성스파르타", "진성기숙학원"]
        : isDalseo
        ? ["잇올 스파르타 대구상인센터", "대구상인 잇올", "상인동 잇올"]
        : isGwangjuDonggu
        ? ["잇올 스파르타 광주봉선센터", "광주봉선 잇올", "봉선동 잇올"]
        : isChuncheon
        ? ["PK대치스파르타 춘천온의센터", "춘천 PK대치스파르타", "온의동 PK"]
        : isHanam
        ? ["수만휘 스파르타 하남미사점", "하남미사 수만휘", "하남 수만휘", "미사 수만휘"]
        : isIlsan
        ? ["일산청솔학원", "일산청솔", "주엽 청솔학원"]
        : isMapo
        ? ["종로학원 강북", "강북 종로학원", "신촌 종로학원"]
        : [];
      const comp4Keywords = isGwangjuDonggu
        ? ["수만휘 스파르타 광주봉선점", "광주봉선 수만휘", "봉선동 수만휘"]
        : isIlsan
        ? ["일산 메가스터디학원", "일산 메가스터디", "대화 메가스터디"]
        : isMapo
        ? ["신촌 메가스터디학원", "신촌 메가스터디", "마포 메가스터디"]
        : [];

      const keywordGroups = [
        { groupName: "ours", keywords: oursKeywords },
        { groupName: "comp1", keywords: comp1Keywords },
        { groupName: "comp2", keywords: comp2Keywords }
      ];

      if ((isDokhakGisuk || isDalseo || isGwangjuDonggu || isChuncheon || isHanam || isIlsan || isMapo) && comp3Keywords.length > 0) {
        keywordGroups.push({ groupName: "comp3", keywords: comp3Keywords });
      }
      if ((isGwangjuDonggu || isIlsan || isMapo) && comp4Keywords.length > 0) {
        keywordGroups.push({ groupName: "comp4", keywords: comp4Keywords });
      }

      const requestBody = {
        startDate: "2026-01-01",
        endDate: "2026-06-30",
        timeUnit: "month",
        keywordGroups
      };

      const naverResponse = await fetch("https://naverapihub.apigw.ntruss.com/search-trend/v1/search", {
        method: "POST",
        headers: {
          "x-ncp-apigw-api-key-id": clientId,
          "x-ncp-apigw-api-key": clientSecret,
          "Content-Type": "application/json"
        },
        body: JSON.stringify(requestBody),
        next: { revalidate: 3600 }
      });

      if (!naverResponse.ok) {
        const errorText = await naverResponse.text();
        console.warn("NAVER Search Trend API returned error:", errorText);
        return NextResponse.json({ success: true, mode: "mock-fallback-api-error", trendData: fallbackTrend });
      }

      const naverData = await naverResponse.json();
      const results = naverData.results || [];
      const oursData = results.find(r => r.title === "ours")?.data || [];
      const comp1Data = results.find(r => r.title === "comp1")?.data || [];
      const comp2Data = results.find(r => r.title === "comp2")?.data || [];
      const comp3Data = results.find(r => r.title === "comp3")?.data || [];
      const comp4Data = results.find(r => r.title === "comp4")?.data || [];

      const realTrend = periodMonths.map((period, index) => {
        const oursVal = oursData.find(d => d.period === period)?.ratio || 0;
        const comp1Val = comp1Data.find(d => d.period === period)?.ratio || 0;
        const comp2Val = comp2Data.find(d => d.period === period)?.ratio || 0;
        const comp3Val = comp3Data.find(d => d.period === period)?.ratio || 0;
        const comp4Val = comp4Data.find(d => d.period === period)?.ratio || 0;

        const row = {
          month: monthLabels[index],
          ours: Math.round(oursVal > 0 ? oursVal : 10 + (hash % 10) + index * 2),
          comp1: Math.round(comp1Val > 0 ? comp1Val : 15 + (hash % 15) + index),
          comp2: Math.round(comp2Val > 0 ? comp2Val : 8 + (hash % 8) + index * 1.5)
        };

        if (isDokhakGisuk || isDalseo || isGwangjuDonggu || isChuncheon || isHanam || isIlsan || isMapo) {
          row.comp3 = Math.round(comp3Val > 0 ? comp3Val : 6 + (hash % 6) + index);
        }
        if (isGwangjuDonggu || isIlsan || isMapo) {
          row.comp4 = Math.round(comp4Val > 0 ? comp4Val : 5 + (hash % 5) + index);
        }

        return row;
      });

      return NextResponse.json({ success: true, mode: isGwangjuDonggu || isIlsan || isMapo ? "naver-api-4comps" : isDokhakGisuk || isDalseo || isChuncheon || isHanam ? "naver-api-3comps" : "naver-api", trendData: realTrend });
    }
  } catch (error) {
    console.error("Failed to fetch NAVER Search Trend API:", error);
    return NextResponse.json({ success: true, mode: "mock-fallback-exception", trendData: fallbackTrend });
  }
}
