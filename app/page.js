"use client";

import React, { useEffect, useMemo, useRef, useState, Fragment, useCallback } from "react";
import * as XLSX from "xlsx";

const OVERVIEW_TAB_ID = "__overall__";
const SPECIAL_SOCIAL_TAB_KIND = "special-social";
const SPECIAL_COLLAB_TAB_KIND = "special-collab";
const SPECIAL_FACILITY_TAB_KIND = "special-facility";
const SPECIAL_MENTOR_TAB_KIND = "special-mentor";
const BROWSER_SAVE_KEY = "branch-activation-dashboard-state";

const createId = (prefix) => `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;

const createEvent = (name) => ({
  id: createId("event"),
  name
});

const regionMapShapes = [
  { key: "서울", label: "서울", path: "M160 74 L168 70 L176 74 L175 83 L166 87 L158 83 Z", x: 167, y: 80 },
  { key: "인천", label: "인천", path: "M112 92 L124 84 L138 85 L146 92 L143 104 L128 110 L113 102 Z", x: 129, y: 98 },
  { key: "경기", label: "경기", path: "M129 97 L149 80 L177 76 L196 85 L204 102 L199 126 L189 144 L168 156 L143 151 L127 133 L122 114 Z", x: 162, y: 118 },
  { key: "강원", label: "강원", path: "M197 81 L223 73 L252 77 L279 94 L286 118 L281 145 L267 168 L242 179 L219 173 L201 151 L192 118 Z", x: 241, y: 126 },
  { key: "충청", label: "충청", path: "M123 156 L150 148 L183 149 L205 162 L207 187 L196 209 L177 222 L149 221 L129 205 L119 182 Z", x: 164, y: 185 },
  { key: "전라", label: "전라", path: "M94 223 L118 214 L148 216 L173 226 L182 246 L179 272 L163 294 L138 305 L112 300 L96 282 L89 254 Z", x: 136, y: 257 },
  { key: "경상", label: "경상", path: "M201 184 L228 178 L259 184 L281 201 L287 229 L285 260 L273 292 L250 312 L224 315 L208 295 L199 268 L196 233 Z", x: 246, y: 246 },
  { key: "제주", label: "제주", path: "M122 334 C136 325, 163 323, 181 331 C167 344, 137 346, 122 334 Z", x: 152, y: 335 }
];

const regionDisplayOrder = ["서울", "경기", "인천", "강원", "충청", "전라", "경상", "제주", "기숙"];
const snsEvaluationBaseDate = "2026-03-24";
const eventScheduleMap = {
  "1기": "22.3.28. - 5.22.",
  "2기": "22.6.27. - 8.21.",
  "3기": "22.11.28. - 12.25.",
  "4기": "23.1.16. - 2.19.",
  "5기": "23.4.17. - 5.28.",
  "6기": "23.7.10. - 8.20.",
  "7기": "23.12.18. - 24.1.14.",
  "8기": "24.1.29. - 2.25.",
  "9기": "24.4.29. - 5.26.",
  "10기": "24.7.15. - 8.11.",
  "11기": "24.12.16. - 25.1.12.",
  "12기": "25.4.28. - 5.25.",
  "13기": "25.7.28. - 8.24.",
  "윈터스쿨": "25.1.20. - 2.2.",
  "14기": "25.12.15. - 26.1.11.",
  "25윈터": "23.12.06. - 23.12.31.",
  "25정규": "24.01.19. - 24.02.22.",
  "25반수": "24.05.10. - 24.06.03.",
  "26윈터": "24.11.22. - 25.01.10.",
  "26정규": "25.02.07. - 25.03.07.",
  "26반수": "25.04.18. - 25.06.02.",
  "27정규": "25.12.19. - 26.01.16."
};
const collabEventColorPalette = [
  { bg: "#f0f4ff", border: "#dbebff" },
  { bg: "#f0fcf9", border: "#d1f4ec" },
  { bg: "#f5f6ff", border: "#e0e3ff" },
  { bg: "#f0f9ff", border: "#e0f2fe" },
  { bg: "#f8fafc", border: "#e2e8f0" },
  { bg: "#f0fdf4", border: "#dcfce7" }
];

const specialSocialColumns = [
  { key: "branch", label: "지점", type: "text", group: "identity" },
  { key: "blogUrl", label: "블로그 주소", type: "url", group: "identity" },
  { key: "instagramUrl", label: "인스타그램 주소", type: "url", group: "identity" },
  { key: "blogRecentPosts", label: "블로그 최근30일 게시물 수", type: "number", group: "blog" },
  { key: "blogLastPosted", label: "블로그 마지막 게시일", type: "date", group: "blog" },
  { key: "blogVisitScore", label: "블로그 방문수 준수 (0~5)", type: "number", group: "blog-score" },
  { key: "instagramRecentPosts", label: "인스타 최근30일 게시물 수", type: "number", group: "blog" },
  { key: "instagramLastPosted", label: "인스타 마지막 게시일", type: "date", group: "blog" },
  { key: "instagramDesignScore", label: "인스타 디자인/썸네일 통일성 (0~5)", type: "number", group: "instagram" },
  { key: "instagramReactionScore", label: "인스타 반응수 준수 (0~5)", type: "number", group: "blog-score" },
  { key: "profileSetupScore", label: "프로필 세팅 완성 (0~3)", type: "number", group: "instagram" },
  { key: "featureUsageScore", label: "인스타그램 부가기능 활용도 (0~3)", type: "number", group: "growth" },
  { key: "ctaScore", label: "CTA 주제 (0~3)", type: "number", group: "growth" },
  { key: "linkHealthScore", label: "링크 연결 정상 작동 (0~3)", type: "number", group: "growth" },
  { key: "brandInfoScore", label: "브랜드/지점 정보 명확 (0~3)", type: "number", group: "instagram" },
  { key: "memo", label: "메모", type: "text", group: "memo" }
];

const defaultCollabColumns = [
  "지역",
  "지점",
  "협업 이벤트 A 홈페이지",
  "협업 이벤트 A 블로그",
  "협업 이벤트 A 인스타/언론기사"
];

const defaultFacilityColumns = ["지역", "지점", "시설영상 URL"];

function isSpecialTabKind(kind) {
  return kind === SPECIAL_SOCIAL_TAB_KIND || kind === SPECIAL_COLLAB_TAB_KIND || kind === SPECIAL_FACILITY_TAB_KIND || kind === SPECIAL_MENTOR_TAB_KIND;
}

function normalizeCollabColumns(columns = []) {
  const normalized = columns
    .map((column) => String(column ?? "").trim())
    .filter(Boolean);

  const remaining = normalized.filter((column) => column !== "지역" && column !== "지점");
  return ["지역", "지점", ...remaining];
}

function createSpecialCollabRow(columns = defaultCollabColumns, seed = {}) {
  const values = Object.fromEntries(
    normalizeCollabColumns(columns).map((column) => [column, seed[column] ?? ""])
  );

  values["지역"] = seed["지역"] ?? seed.region ?? values["지역"] ?? "";
  values["지점"] = seed["지점"] ?? seed.branch ?? values["지점"] ?? "";

  return {
    id: seed.id || createId("collab-row"),
    values
  };
}

function createSpecialCollabTab(id, name, seededRows = [], columns = defaultCollabColumns) {
  const normalizedColumns = normalizeCollabColumns(columns);
  const rows = seededRows.length > 0
    ? seededRows.map((row) => createSpecialCollabRow(normalizedColumns, row))
    : [
        createSpecialCollabRow(normalizedColumns, { 지역: "서울", 지점: "강남" }),
        createSpecialCollabRow(normalizedColumns, { 지역: "경기", 지점: "분당정자" }),
        createSpecialCollabRow(normalizedColumns, { 지역: "부산", 지점: "부산대" })
      ];

  return {
    id,
    name,
    kind: SPECIAL_COLLAB_TAB_KIND,
    events: [],
    rows: [],
    collabColumns: normalizedColumns,
    collabRows: rows
  };
}

function createSpecialFacilityRow(seed = {}) {
  return {
    id: seed.id || createId("facility-row"),
    region: seed.region || seed["지역"] || "",
    branch: seed.branch || seed["지점"] || "",
    url: seed.url || seed["시설영상 URL"] || seed["URL"] || ""
  };
}

function createSpecialFacilityTab(id, name, seededRows = []) {
  const rows = seededRows.length > 0
    ? seededRows.map((row) => createSpecialFacilityRow(row))
    : [
        createSpecialFacilityRow({ region: "서울", branch: "대치" }),
        createSpecialFacilityRow({ region: "서울", branch: "강북" }),
        createSpecialFacilityRow({ region: "경기", branch: "분당정자" })
      ];

  return {
    id,
    name,
    kind: SPECIAL_FACILITY_TAB_KIND,
    events: [],
    rows: [],
    facilityColumns: defaultFacilityColumns,
    facilityRows: rows
  };
}

function createSpecialSocialRow(seed = {}) {
  return {
    id: seed.id || createId("social-row"),
    branch: seed.branch || "",
    blogUrl: seed.blogUrl || "",
    instagramUrl: seed.instagramUrl || "",
    blogRecentPosts: String(seed.blogRecentPosts ?? "0"),
    blogLastPosted: seed.blogLastPosted || "",
    blogVisitScore: String(seed.blogVisitScore ?? "0"),
    instagramRecentPosts: String(seed.instagramRecentPosts ?? "0"),
    instagramLastPosted: seed.instagramLastPosted || "",
    instagramDesignScore: String(seed.instagramDesignScore ?? "0"),
    instagramReactionScore: String(seed.instagramReactionScore ?? "0"),
    profileSetupScore: String(seed.profileSetupScore ?? "0"),
    featureUsageScore: String(seed.featureUsageScore ?? "0"),
    ctaScore: String(seed.ctaScore ?? "0"),
    linkHealthScore: String(seed.linkHealthScore ?? "0"),
    brandInfoScore: String(seed.brandInfoScore ?? "0"),
    memo: seed.memo || ""
  };
}

function createSpecialSocialTab(id, name, seededRows = []) {
  const rows = seededRows.length > 0
    ? seededRows.map((row) => createSpecialSocialRow(row))
    : [
        createSpecialSocialRow({ branch: "강북" }),
        createSpecialSocialRow({ branch: "강남" }),
        createSpecialSocialRow({ branch: "목동" })
      ];

  return {
    id,
    name,
    kind: SPECIAL_SOCIAL_TAB_KIND,
    events: [],
    rows: [],
    socialRows: rows
  };
}

function createSpecialMentorRow(seed = {}) {
  return {
    id: seed.id || createId("mentor-row"),
    year: String(seed.year ?? seed["연도"] ?? new Date().getFullYear()),
    name: String(seed.name ?? seed["이름"] ?? ""),
    phone: String(seed.phone ?? seed["번호"] ?? seed["연락처"] ?? ""),
    university: String(seed.university ?? seed["합격 대학"] ?? seed["합격대학"] ?? ""),
    department: String(seed.department ?? seed["학과"] ?? ""),
    branch: String(seed.branch ?? seed["지점"] ?? ""),
    group: String(seed.group ?? seed["1억장학금"] ?? seed["장학그룹"] ?? ""),
    amount: Number(seed.amount ?? seed["1억 장학금"] ?? seed["장학금액"] ?? 0),
    isMentor: Boolean(seed.isMentor ?? seed["멘토단여부"] ?? seed["멘토단"] ?? false),
    memo: String(seed.memo ?? seed["비고"] ?? "")
  };
}

function createSpecialMentorTab(id, name, seededRows = []) {
  const rows = seededRows.length > 0
    ? seededRows.map((row) => createSpecialMentorRow(row))
    : [
        createSpecialMentorRow({
          year: "2026",
          name: "김철수",
          phone: "010-1234-5678",
          university: "서울대학교",
          department: "의예과",
          branch: "강남",
          group: "1그룹",
          amount: 3000000,
          isMentor: true,
          memo: "우수 멘토"
        }),
        createSpecialMentorRow({
          year: "2026",
          name: "이영희",
          phone: "010-5678-1234",
          university: "연세대학교",
          department: "치의예과",
          branch: "대치",
          group: "2그룹",
          amount: 2000000,
          isMentor: false,
          memo: ""
        })
      ];

  const sortedRows = [...rows].sort((a, b) => {
    const yA = parseInt(a.year, 10) || 0;
    const yB = parseInt(b.year, 10) || 0;
    if (yB !== yA) return yB - yA;
    return (a.name || "").localeCompare(b.name || "", "ko");
  });

  return {
    id,
    name,
    kind: SPECIAL_MENTOR_TAB_KIND,
    events: [],
    rows: [],
    mentorRows: sortedRows
  };
}

function createRow(eventIds = [], seed = {}) {
  const eventValues = Object.fromEntries(
    eventIds.map((eventId) => [
      eventId,
      {
        status: seed.eventValues?.[eventId]?.status === "O" ? "O" : "X",
        participants: String(seed.eventValues?.[eventId]?.participants ?? "0")
      }
    ])
  );

  return {
    id: seed.id || createId("row"),
    region: seed.region || "",
    branch: seed.branch || "",
    eventValues
  };
}

function createTab(id, name, eventNames = ["기본 이벤트"], seededRows = []) {
  const events = eventNames.map((eventName) => createEvent(eventName));
  const eventIds = events.map((event) => event.id);
  const rows =
    seededRows.length > 0
      ? seededRows.map((row) => createRow(eventIds, row))
      : [
          createRow(eventIds, {
            region: "서울",
            branch: "강남",
            eventValues: eventIds[0]
              ? {
                  [eventIds[0]]: { status: "O", participants: "12" }
                }
              : {}
          }),
          createRow(eventIds, {
            region: "경기",
            branch: "분당정자",
            eventValues: eventIds[0]
              ? {
                  [eventIds[0]]: { status: "O", participants: "8" }
                }
              : {}
          }),
          createRow(eventIds, {
            region: "부산",
            branch: "부산대"
          })
        ];

  return {
    id,
    name,
    kind: "default",
    events,
    rows
  };
}

const initialTabs = [
  createTab("tab-1", "247프렌즈", ["1차 설명회", "재등록 캠페인"]),
  createTab("tab-2", "247체험단", ["체험단 OT"], [
    {
      region: "서울",
      branch: "목동",
      eventValues: {}
    },
    {
      region: "인천",
      branch: "인천송도",
      eventValues: {}
    },
    {
      region: "전라",
      branch: "광주동구",
      eventValues: {}
    }
  ]),
  createSpecialSocialTab("tab-social-1", "SNS 진단표"),
  createSpecialCollabTab("tab-collab-1", "협업이벤트"),
  createSpecialFacilityTab("tab-facility-1", "지점시설영상"),
  createSpecialMentorTab("tab-mentor-1", "멘토단 및 장학생")
];

function normalizeParticipantValue(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) && numeric >= 0 ? String(numeric) : "0";
}

function stripEventSuffix(label) {
  return String(label ?? "")
    .replace(/참여여부$/, "")
    .replace(/참석인원$/, "")
    .replace(/참여인원$/, "")
    .replace(/인원$/, "")
    .trim();
}

function normalizeRegionLabel(region) {
  if (region.includes("서울")) return "서울";
  if (region.includes("기숙")) return "기숙";
  if (region.includes("인천")) return "인천";
  if (region.includes("경기")) return "경기";
  if (region.includes("충청") || region.includes("대전") || region.includes("세종")) return "충청";
  if (region.includes("전라") || region.includes("광주")) return "전라";
  if (region.includes("경상") || region.includes("부산") || region.includes("대구") || region.includes("울산")) return "경상";
  if (region.includes("강원")) return "강원";
  if (region.includes("제주")) return "제주";
  return region.trim() || "기타";
}

function normalizeBranchKey(branch) {
  return String(branch ?? "")
    .replace(/\s+/g, "")
    .replace(/지점$/u, "")
    .trim()
    .toLowerCase();
}

const branchCoordinateMap = {
  "목동": [37.5255, 126.8640],
  "도봉": [37.6688, 127.0470],
  "인천송도": [37.3850, 126.6540],
  "송도": [37.3850, 126.6540],
  "일산서구": [37.6740, 126.7460],
  "일산": [37.6740, 126.7460],
  "대전둔산": [36.3500, 127.3850],
  "둔산": [36.3500, 127.3850],
  "강남": [37.4979, 127.0276],
  "대치": [37.4930, 127.0600],
  "강북": [37.6390, 127.0250],
  "분당정자": [37.3680, 127.1080],
  "분당": [37.3820, 127.1180],
  "부산대": [35.2300, 129.0820],
  "부산": [35.1796, 129.0756],
  "광주동구": [35.1460, 126.9230],
  "광주남구": [35.1320, 126.9020],
  "광주": [35.1595, 126.8526],
  "대구수성": [35.8580, 128.6250],
  "대구": [35.8714, 128.6014],
  "평촌": [37.3910, 126.9630],
  "수지": [37.3220, 127.0970],
  "부천": [37.4850, 126.7820],
  "안산": [37.3210, 126.8300],
  "목포": [34.8077682, 126.4598048],
  "제주": [33.4925954, 126.5428095],
  "기숙": [37.2890, 127.2060]
};

const branchAddressMap = {
  "대치": "서울특별시 강남구 역삼로 432 1층(대치동 910-2)",
  "강북": "서울특별시 노원구 노원로236 8층 (하계동 256-8)",
  "노량진": "서울특별시 동작구 노량진로 104 리더스타워 4층 (노량진동 41-9)",
  "마포": "서울특별시 마포구 백범로 107 2층 (대흥동 186-1)",
  "목동": "서울특별시 양천구 목동로 183 3층 (신정동 995-2)",
  "목동오목교": "서울특별시 양천구 목동동로 73 양지빌딩 2차 13층 (신정동 323-12)",
  "서울강동": "서울특별시 강동구 양재대로 1543 진주빌딩 6층 (천호동 45-9)",
  "서울강서": "서울특별시 강서구 강서로 348, 우장산힐스테이트 상가건물 2층",
  "서울광진": "서울특별시 광진구 아차산로 471 CS플라자 2층 (구의동 212-3)",
  "서울대점": "서울특별시 관악구 남부순환로 1730 청운빌딩 4층 (봉천동 927-5)",
  "서울도봉": "서울특별시 도봉구 도봉로 524 하버드빌딩 5층 (창동 700-36)",
  "서울성동": "서울특별시 성동구 왕십리로 410, 153타워 2층 (하왕십리동 1070)",
  "서울성북": "서울특별시 성북구 동소문로 125 골든타워 13층 (동선동4가 111-1)",
  "서울송파": "서울특별시 송파구 백제고분로 354 동일빌딩 4층 (석촌동 285-4)",
  "은평서대문": "서울특별시 은평구 서오릉로 196 구산빌딩 4층 (갈현동 498-4)",
  "광명": "경기도 광명시 범안로 1039 8층 (하안동 34-11)",
  "다산": "경기도 남양주시 다산중앙로 105-8 9층(다산동 6088-3)",
  "김포": "경기도 김포시 김포한강4로 162 한강메트로 6층 (장기동 1868-11)",
  "동탄": "경기도 화성시 동탄반석로 134 7층(반송동 104-1)",
  "부천": "경기도 부천시 길주로 181 골든벨타워 3층 (중동 1031-2)",
  "분당정자": "경기도 성남시 분당구 성남대로 381 폴라리스빌딩 2층 (정자동 15-3)",
  "수원시청": "경기도 수원시 권선구 효원로266번길 25 우덕빌딩 3층 (권선동 1014-5)",
  "수원영통": "경기도 수원시 영통구 봉영로 1605 모던타운 7층 (영통동 959-1)",
  "수원정자": "경기도 수원시 장안구 정자천로173번길 11-7 효신빌딩 6층 (정자동 878-11)",
  "안산": "경기도 안산시 단원구 광덕대로 130 폴리타운 A동 8층 (고잔동 775)",
  "용인수지": "경기도 용인시 수지구 문정로7번길 8 창진빌딩 2층 (풍덕천동 1082-8)",
  "의정부": "경기도 의정부시 평화로 385 4층 (호원동 414-7)",
  "일산동구": "경기도 고양시 일산동구 일산로 226 3층 (마두동 724-1)",
  "일산서구": "경기도 고양시 일산서구 중앙로 1419 정도프라자 5층 (주엽동 106-1)",
  "파주": "경기도 파주시 미래로 377 8층 (동패동 1759-4)",
  "평택": "경기도 평택시 평택5로20번길 39 6층 (합정동 964-13)",
  "하남": "경기도 하남시 미사강변중앙로 214 9층 (망월동 1083)",
  "인천부평": "인천광역시 부평구 부평대로 90 여산빌딩 10층 (부평동 440-5)",
  "인천송도": "인천광역시 연수구 해돋이로152번길 35 8층 (송도동 21-45)",
  "인천인하대": "인천광역시 미추홀구 용정공원로83번길 59 드림빌딩 11층(용현동 665-14)",
  "인천청라": "인천광역시 서구 중봉대로 588 센트럴프라자 8층 (청라동 162-14)",
  "원주": "강원도 원주시 남원로 570 (개운동 434-15)",
  "춘천": "강원도 춘천시 동내면 춘천순환로 131 3층 (거두리1058-1)",
  "대전둔산": "대전광역시 서구 둔산남로 127 (둔산동 1510)",
  "천안": "충청남도 천안시 서북구 번영로 100 9층 (불당동 727)",
  "청주": "충청북도 청주시 흥덕구 대농로 70 8층 (복대동 288-123)",
  "광주남구": "광주광역시 남구 대남대로 181 2층 (주월동 1273-6)",
  "광주동구": "광주광역시 동구 서석로 87 2층 (대의동 37)",
  "광주북구": "광주광역시 북구 북문대로 27 7층 (운암동 96-2)",
  "광주수완": "광주광역시 광산구 임방울대로 325 정인빌딩 5층 (수완동 1418)",
  "목포": "전라남도 무안군 삼향읍 남악2로22번길 30, 215호 (남악리 2273)",
  "익산": "전라북도 익산시 하나로 440 영창빌딩 3층 (어양동 633-4)",
  "대구달서": "대구광역시 달서구 월배로 219 센트로타워 4층 (상인동 72-3)",
  "대구수성 1관": "대구광역시 수성구 달구벌대로 2538 2층 (범어동 212-10)",
  "대구수성 2관": "대구광역시 중구 달구벌대로 2166 4층 (대봉동 715-10)",
  "부산교대": "부산광역시 연제구 중앙대로 1201 중보빌딩 3층 (거제동 75-7)",
  "부산대": "부산광역시 금정구 금강로 252-1 근영테크빌 4층 (장전동 420-47)",
  "부산북구": "부산광역시 북구 금곡대로303번길 12 샤롯데 6층 (화명동 2269-2)",
  "부산서면": "부산광역시 진구 동천로 24번길 16 (전포동 882-24)",
  "부산해운대": "부산광역시 해운대구 센텀2로 29 메커스빌딩 14층 (우동 1509)",
  "울산남구": "울산광역시 남구 문수로 339 6층 (옥동 589-2)",
  "진주": "경상남도 진주시 진주대로 964 3층 (칠암동 296)",
  "창원": "경상남도 창원시 성산구 원이대로682번길 14 삼광빌딩 10층 (상남동 10-3)",
  "제주": "제주도 제주시 승천로 71 3층 (아라2동 3001-9)",
  "안성기숙": "경기도 안성시 삼죽면 국사봉로 246-14 (기솔리 422)",
  "이천기숙": "경기도 이천시 마장면 이장로 115-10 (이치리 160-5)",
  "독학기숙": "경기도 광주시 초월읍 두둘기길 68-21 (신월리 218-2)"
};

const branchCompetitorAddressMap = {
  "분당정자": [
    { name: "메가스터디 러셀 분당학원", address: "경기도 성남시 분당구 성남대로 381 폴라리스Ⅰ빌딩 4층" },
    { name: "수만휘 스파르타 분당정자점", address: "경기도 성남시 분당구 정자일로 232 젤존타워1 10층" },
    { name: "수능선배 분당점", address: "경기도 성남시 분당구 돌마로 52 MD프라자 7층" },
    { name: "잇올 분당수내센터", address: "경기도 성남시 분당구 황새울로258번길 40 2층" },
    { name: "잇올 분당정자센터", address: "경기도 성남시 분당구 정자일로 232 젤존타워1 10층" },
    { name: "잇올 분당이매센터", address: "경기도 성남시 분당구 판교로 476 오성빌딩 5층" },
    { name: "잇올 용인 수지센터 1관", address: "경기도 용인시 수지구 풍덕천로 135 요진타워 5층, 6층" },
    { name: "잇올 용인 수지센터 2관", address: "경기도 용인시 수지구 풍덕천로 145 유용빌딩 4층" },
    { name: "디랩 분당", address: "경기도 성남시 분당구 황새울로258번길 41 3층" }
  ]
};

function getBranchCoords(branchName) {
  const name = String(branchName || "").trim();
  const normalizedName = normalizeBranchKey(name);
  const exactMatch = Object.entries(branchCoordinateMap).find(([branch]) => normalizeBranchKey(branch) === normalizedName);
  if (exactMatch) return exactMatch[1];

  const partialMatch = Object.entries(branchCoordinateMap)
    .sort((a, b) => b[0].length - a[0].length)
    .find(([branch]) => normalizedName.includes(normalizeBranchKey(branch)));
  if (partialMatch) return partialMatch[1];
  
  const preciseCoords = {
    "목동": [37.5255, 126.864],
    "도봉": [37.6688, 127.047],
    "인천송도": [37.385, 126.654],
    "송도": [37.385, 126.654],
    "일산서구": [37.674, 126.746],
    "일산": [37.674, 126.746],
    "대전둔산": [36.350, 127.385],
    "둔산": [36.350, 127.385],
    "강남": [37.4979, 127.0276],
    "대치": [37.493, 127.060],
    "강북": [37.639, 127.025],
    "분당정자": [37.368, 127.108],
    "분당": [37.382, 127.118],
    "부산대": [35.230, 129.082],
    "부산": [35.1796, 129.0756],
    "광주동구": [35.146, 126.923],
    "광주남구": [35.132, 126.902],
    "광주": [35.1595, 126.8526],
    "대구수성": [35.858, 128.625],
    "대구": [35.8714, 128.6014],
    "평촌": [37.391, 126.963],
    "수지": [37.322, 127.097],
    "부천": [37.485, 126.782],
    "안산": [37.321, 126.830],
    "기숙": [37.289, 127.206]
  };

  for (const key in preciseCoords) {
    if (name.includes(key)) {
      return preciseCoords[key];
    }
  }

  let hash = 0;
  for (let i = 0; i < name.length; i++) {
    hash = name.charCodeAt(i) + ((hash << 5) - hash);
  }
  
  let center = [36.5, 127.8];
  if (name.includes("서울") || name.includes("경기") || name.includes("인천") || name.includes("일산") || name.includes("분당") || name.includes("평촌") || name.includes("수지") || name.includes("부천") || name.includes("대치") || name.includes("강남") || name.includes("강북") || name.includes("도봉") || name.includes("목동")) {
    center = [37.5665, 126.9780];
  } else if (name.includes("부산") || name.includes("울산") || name.includes("경남")) {
    center = [35.1796, 129.0756];
  } else if (name.includes("대구") || name.includes("경북")) {
    center = [35.8714, 128.6014];
  } else if (name.includes("광주") || name.includes("전라") || name.includes("전남") || name.includes("전북")) {
    center = [35.1595, 126.8526];
  } else if (name.includes("대전") || name.includes("충청") || name.includes("세종") || name.includes("충남") || name.includes("충북")) {
    center = [36.3504, 127.3845];
  } else if (name.includes("강원")) {
    center = [37.7518, 128.8760];
  } else if (name.includes("제주")) {
    center = [33.4996, 126.5312];
  }

  const latOffset = ((Math.abs(hash) % 100) / 1000) * 0.15 - 0.075;
  const lngOffset = (((Math.abs(hash) >> 8) % 100) / 1000) * 0.15 - 0.075;
  
  return [center[0] + latOffset, center[1] + lngOffset];
}

function getBranchAddress(branchName) {
  const normalizedName = normalizeBranchKey(branchName);
  const exactMatch = Object.entries(branchAddressMap).find(([branch]) => normalizeBranchKey(branch) === normalizedName);
  if (exactMatch) return exactMatch[1];

  const partialMatch = Object.entries(branchAddressMap)
    .sort((a, b) => b[0].length - a[0].length)
    .find(([branch]) => normalizedName.includes(normalizeBranchKey(branch)) || normalizeBranchKey(branch).includes(normalizedName));

  return partialMatch?.[1] || "";
}

function getBranchCompetitorAddresses(branchName) {
  const normalizedName = normalizeBranchKey(branchName);
  const exactMatch = Object.entries(branchCompetitorAddressMap).find(([branch]) => normalizeBranchKey(branch) === normalizedName);
  if (exactMatch) return exactMatch[1];

  const partialMatch = Object.entries(branchCompetitorAddressMap)
    .sort((a, b) => b[0].length - a[0].length)
    .find(([branch]) => normalizedName.includes(normalizeBranchKey(branch)) || normalizeBranchKey(branch).includes(normalizedName));

  return partialMatch?.[1] || [];
}

function getNearbyCoords(coords, index) {
  const angle = (index / 6) * Math.PI * 2;
  const distance = 0.0022 + (index % 3) * 0.0005;
  return [
    coords[0] + Math.sin(angle) * distance,
    coords[1] + Math.cos(angle) * distance
  ];
}

function getDisplayMapItems(items = []) {
  const groups = new Map();

  items.forEach((item, index) => {
    const key = item.coords.map((coord) => Number(coord).toFixed(6)).join(",");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ item, index });
  });

  const displayItems = [...items];
  groups.forEach((group) => {
    if (group.length <= 1) return;

    group.forEach(({ item, index }, groupIndex) => {
      const angle = (groupIndex / group.length) * Math.PI * 2;
      const distance = 0.00018 + Math.min(group.length, 6) * 0.000025;
      displayItems[index] = {
        ...item,
        displayCoords: [
          item.coords[0] + Math.sin(angle) * distance,
          item.coords[1] + Math.cos(angle) * distance
        ]
      };
    });
  });

  return displayItems.map((item) => ({
    ...item,
    displayCoords: item.displayCoords || item.coords
  }));
}

function getBranchMarkerHtml(branchName, isActive, address = "", markerType = "branch") {
  return `
    <div class="custom-leaflet-marker ${markerType} ${isActive ? "active" : ""}" title="${branchName}">
      <div class="marker-pulse-ring"></div>
      <div class="marker-dot-core"></div>
      <div class="marker-label">${branchName}</div>
      ${address ? `<div class="marker-address">${address}</div>` : ""}
    </div>
  `;
}

function generateDynamicCompetitorData(compName, promotions) {
  if (!promotions || promotions.length === 0) return null;

  // 각 키워드를 처음 트리거한 프로모션 제목을 기억 (경쟁사별 고유 요약 생성용)
  const seen = {};
  const trendList = [];
  const guideList = [];

  const patterns = [
    {
      key: "mockExam",
      regex: /모의고사|더프|이감|평가원|시험|응시|대비반|학교프로젝트/,
      trendFn: (title) => `'${title}' 등 실전 모의고사·시뮬레이션 운영으로 수험생 실전 감각 강화를 전면 마케팅 중`,
      guide: "[모의고사 대응] 자사 풀시즌 모의고사 라인업과 1:1 오답 약점 진단 피드백 시스템의 비교 우위를 전면에 내세우십시오."
    },
    {
      key: "briefing",
      regex: /설명회|간담회|입시분석|세미나|라이브|입시/,
      trendFn: (title) => `'${title}' 등 입시 라이브·설명회 운영으로 수험생·학부모 정보 수요를 신규 모객에 적극 활용 중`,
      guide: "[설명회/입시 대응] 이투스 교육평가연구소의 공신력을 바탕으로 1:1 생기부 정밀 진단 수시 컨설팅을 선제 홍보하십시오."
    },
    {
      key: "discount",
      regex: /할인|지원금|감면|혜택|특가|체험|쿠폰|무료|이벤트/,
      trendFn: (title) => `'${title}' 등 이벤트·무료 혜택 제공으로 원생 결속 강화 및 신규 수험생 유입을 도모 중`,
      guide: "[이벤트/혜택 대응] 단순 이벤트성 혜택보다 자사의 맞춤형 장학 제도와 장기 학습 코칭의 실질적 성과를 구체적 사례로 부각하십시오."
    },
    {
      key: "scholarship",
      regex: /장학|성적우수|우수자|장학생/,
      trendFn: (title) => `'${title}' 등 성적 우수자 장학 혜택을 전면 내세워 상위권 수험생 선점 전략을 구사 중`,
      guide: "[장학 대응] 자사 성적 구간별 장학 제도의 실질 혜택 금액과 명문대 합격생 수기를 비교 자료로 제시하십시오."
    },
    {
      key: "facility",
      regex: /리뉴얼|좌석|독서실|자습실|데스크|시설|의자|플래너/,
      trendFn: (title) => `'${title}' 등 학습 도구·자습 환경 강화로 체계적 루틴 관리 이미지를 부각 중`,
      guide: "[학습 도구 대응] 자사 담임의 플래너 직접 점검·피드백 루틴의 밀착도와 전문성을 데이터로 비교 제시하십시오."
    },
    {
      key: "recruit",
      regex: /모집|개강|반수|n수|정규|등록|티처스|멘토링/,
      trendFn: (title) => `'${title}' 등 개인화 코칭·반수반 모집으로 재수·반수 수험생 유치에 공격적으로 나서는 중`,
      guide: "[코칭/반수 대응] 자사의 전담 담임 코치 체계와 반수 성공 보장 패키지(맞춤 시간표·플래너)의 디테일을 대외적으로 강조하십시오."
    }
  ];

  // 프로모션별로 순회하며 각 주제를 처음 감지한 제목을 기록
  promotions.forEach(p => {
    const text = `${p.title} ${p.content}`.toLowerCase();
    patterns.forEach(pat => {
      if (!seen[pat.key] && pat.regex.test(text)) {
        seen[pat.key] = true;
        trendList.push(pat.trendFn(p.title));
        guideList.push(pat.guide);
      }
    });
  });

  // 상위 3가지만 출력
  const trend = trendList.slice(0, 3).map((t, i) => `${i + 1}. ${t}`).join("\n")
    || "1. 자체 콘텐츠 제공 및 학생 관리 강화를 기반으로 지역 내 원생 관리에 주력하고 있습니다.";

  const guide = guideList.slice(0, 3).map((g, i) => `${i + 1}. ${g}`).join("\n")
    || "1. 경쟁사의 개별 학생 밀착 관리에 대응해 자사만의 1:1 맞춤형 멘토 피드백 케어의 깊이와 정밀함을 상담 시 집중 안내하십시오.";

  return { trend, guide };
}

function getCompetitorsForBranch(branchName) {
  const cleanName = String(branchName || "").trim();

  if (cleanName.includes("분당정자") || cleanName.includes("분당")) {
    return [
      {
        name: "수만휘 스파르타 분당정자점",
        urgency: "high",
        blogUrl: "https://blog.naver.com/smh_bd",
        metrics: { ours: { mgmt: 92, content: 95, facility: 85 }, theirs: { mgmt: 88, content: 80, facility: 82 } },
        trend: "1. 6평 분석 기반 하반기 대입 전략 분석 및 자체 입시 간행물 'SU:ZIP(수집)' 월간지를 배포하며 입시 정보력 우위를 홍보 중\n2. N수생 전용 위생 식단표 주기적 제공을 강조하며 생활 관리 경쟁력 부각\n3. 수만휘 회원 인증 시 수강료 10% 즉시 감면 프로모션으로 신규 원생 유입 공세",
        promotions: [
          { title: "수만휘 회원 10% 수강료 지원", content: "수만휘 회원 인증 시 첫 달 등록 수강료의 10% 즉시 감면 프로모션", url: "https://blog.naver.com/smh_bd" }
        ],
        guide: "1. [입시 콘텐츠 대응] 자체 월간지 'SU:ZIP'에 맞서 이투스 교육평가연구소의 공신력 높은 6평 정밀 분석 리포트를 원생 및 상담 고객에게 선제 배포하십시오.\n2. [생활 관리 대응] N수생 식단 관리 홍보에 맞서 자사만의 엄격한 생활 루틴 관리(기상·취침·플래너 점검) 시스템의 깊이를 부각하십시오.\n3. [가격 대응] 수강료 감면 공세에 대해 자사의 장학 제도 및 장기 등록 혜택을 비교 안내하여 가성비 우위를 증명하십시오."
      },
      {
        name: "수능선배 분당점",
        urgency: "high",
        blogUrl: "https://blog.naver.com/bioochemistry",
        metrics: { ours: { mgmt: 92, content: 95, facility: 85 }, theirs: { mgmt: 94, content: 75, facility: 90 } },
        trend: "1. 프리미엄 1인 태블릿 자습석 도입으로 최신 시설 경쟁력을 전면에 내세우는 중\n2. 실시간 화면 원격 출결 감시 시스템으로 밀착 관리형 마케팅 전개\n3. 등록 원생 전원 대상 수능 배치 점수별 맞춤 원서 대면 컨설팅 무료 제공",
        promotions: [
          { title: "수시/정시 입시 컨설팅 무료 제공", content: "등록 원생 전원에게 수능 배치 점수별 맞춤 원서 대면 컨설팅 무료 서비스 지원", url: "https://www.etoos.com" }
        ],
        guide: "1. [관리 방식 대응] 기계적 화면 감시보다 개개인의 약점을 진단·보완하는 이투스247 1:1 밀착 학습 코칭 케어의 교육적 우월성을 상담 시 핵심 차별점으로 어필하십시오.\n2. [시설 대응] 태블릿 자습석의 화려함보다 독립 면학 전용 공간의 집중력 향상 효과를 데이터(성적 향상률)로 제시하십시오.\n3. [컨설팅 대응] 경쟁사의 무료 컨설팅 공세에 맞서 자사 이투스 교육평가연구소의 정밀 생기부 진단 수시 컨설팅의 깊이와 신뢰도를 부각하십시오."
      },
      {
        name: "메가스터디 러셀 분당학원",
        urgency: "high",
        blogUrl: "",
        metrics: { ours: { mgmt: 92, content: 95, facility: 85 }, theirs: { mgmt: 80, content: 96, facility: 94 } },
        trend: "1. 러셀 단과 2과목 이상 등록 시 자습관 수강료 30% 감면 연계 패키지로 가격 경쟁력 공세\n2. 메가스터디 대형 모의고사 콘텐츠 무상 제공으로 교육 콘텐츠 경쟁력 부각\n3. 대형 브랜드 파워를 내세운 신뢰도 및 인지도 마케팅 전개 중",
        promotions: [
          { title: "단과 연계 바자관 수강료 30% 감면", content: "러셀 단과 2과목 이상 수강 등록 시 자습관 월 이용료 30% 할인 패키지", url: "https://www.etoos.com" }
        ],
        guide: "1. [환경 대응] 대형 단과 자습실 특유의 산만한 분위기 대비, 이투스247의 독립 면학 전용 공간과 엄격한 정숙 관리 환경이 성적 향상에 직결됨을 데이터로 안내하십시오.\n2. [가격 대응] 30% 감면 패키지의 외형적 혜택에 맞서 자사 장학 제도·장기 등록 혜택의 실질 절감액을 구체적 수치로 비교 제시하십시오.\n3. [콘텐츠 대응] 메가스터디 모의고사 무상 제공에 맞서 이투스 자체 모의고사 + 1:1 오답 피드백 클리닉의 결합 가치를 전면에 내세우십시오."
      },
      {
        name: "잇올 경기A 광역",
        urgency: "medium",
        blogUrl: "https://blog.naver.com/italljj2",
        metrics: { ours: { mgmt: 92, content: 95, facility: 85 }, theirs: { mgmt: 90, content: 78, facility: 92 } },
        trend: "1. 여름 수험생 건강 관리 이벤트(이열치열) 운영 - 중·석식 식사권 증정으로 원생 결속 강화\n2. 2027 잇올 디지털 플래너 무료 배포로 학습 루틴 관리 이미지 강화\n3. 반수생 대상 장학 혜택 프로그램 모집 및 홍보 집중 전개\n4. 9월 평가원 모의고사 외부생 접수 및 학교프로젝트 실전 시뮬레이션 운영\n5. '잇올 티처스' 초개인화 멘토링 시스템 홍보로 학습 코칭 경쟁력 부각\n6. 6평 이후 입시 라이브 설명회 진행 및 반수반 시스템 온라인 안내",
        promotions: [
          { title: "여름 이벤트 - 건강한 여름나기 이열치열", content: "수험생 여름 건강 관리 이벤트 참여 시 중·석식 식사권 증정 (7/3~7/9)", url: "https://blog.naver.com/italljj2/224330736917" },
          { title: "2027 잇올 디지털 플래너 안내", content: "수능까지 체계적인 루틴 관리를 위한 디지털 플래너 & 스티커팩 무료 제공", url: "https://blog.naver.com/italljj2/224327173414" },
          { title: "2027학년도 반수생 장학 혜택 안내", content: "반수생 성적 우수자 대상 학습 비용 감면 장학 제도 운영 및 선발 기준 안내", url: "https://blog.naver.com/italljj2/224323857145" },
          { title: "9월 평가원 모의고사 외부생 접수", content: "잇올 비재원생 대상 9평 외부생 온라인 접수 및 2차 방문 접수 안내", url: "https://blog.naver.com/italljj2/224322557609" },
          { title: "잇올 학교프로젝트 - 한 여름의 수능", content: "실제 수능과 동일한 환경·시간표·긴장감을 재현한 실전 시뮬레이션 모의고사", url: "https://blog.naver.com/italljj2/224315450627" },
          { title: "잇올 티처스 초개인화 멘토링 시스템", content: "전과목 학습 질문·피드백, 개인별 학습 전략 설계, 입시 컨설팅을 한 번에 제공하는 잇올 티처스 운영", url: "https://blog.naver.com/italljj2/224282996609" }
        ],
        guide: "1. [이벤트/모객 대응] 여름 이벤트·식사권 증정 등 감성 마케팅에 맞서 자사 원생 대상 학습 격려 프로그램(멘토 1:1 피드백 세션, 성적 향상 포상)의 실질적 가치를 부각하십시오.\n2. [플래너/학습 관리 대응] 디지털 플래너 무료 배포에 맞서 이투스247의 담임 플래너 점검·피드백 시스템의 밀착도와 전문성을 데이터로 비교 제시하십시오.\n3. [반수 장학 대응] 경쟁사의 반수생 장학 공세에 맞서 자사의 반수 성공 패키지(성적 구간별 장학+맞춤 시간표)의 실질 혜택 금액을 구체적 수치로 안내하십시오.\n4. [모의고사/설명회 대응] 9평 외부생 모집 및 학교프로젝트 실전 모의고사에 맞서 이투스 자체 모의고사+1:1 오답 피드백 클리닉의 결합 가치를 전면 홍보하십시오.\n5. [코칭 시스템 대응] '잇올 티처스' 멘토링 공세에 대해 자사 담임 코치의 전담 밀착 관리 체계와 성적 향상 사례를 실증 자료로 제시하십시오.\n6. [입시 정보 대응] 입시 라이브 설명회 모객에 맞서 이투스 교육평가연구소의 6평 정밀 분석 리포트와 1:1 생기부 수시 컨설팅의 공신력을 부각하십시오."
      },
      {
        name: "디랩 분당",
        urgency: "medium",
        blogUrl: "https://m.place.naver.com/place/1960381657/feed",
        metrics: { ours: { mgmt: 92, content: 95, facility: 85 }, theirs: { mgmt: 85, content: 88, facility: 86 } },
        trend: "최근 대성학원의 우수한 교육 인프라를 강조하며 다음과 같이 공격적인 프로모션을 펼치고 있습니다:\n1. 이감 시즌5 외부생 모집 (7/7 마감)\n2. 7월 더프 모의고사 외부생 모집 (7/18 시행)\n3. 디랩 성적관리 시스템 실제 성공사례 분석 홍보\n4. 2027 수시지원 전략설명회 (디랩 목동 등 전국 연계)\n5. 2027학년도 수시 지원 전략 재원생 대상 설명회 개최\n6. 대구/대전 신규 확장 오픈 기념 반수반 대규모 모객 마케팅",
        promotions: [
          {
            title: "이감 모의고사 시즌5 외부 응시생 모집",
            content: "실전 감각 극대화를 위한 이감 모의고사 시즌5 외부 응시생 모집 (신청기간 ~ 7/7 화요일까지)",
            url: "https://m.place.naver.com/place/1960381657/feed"
          },
          {
            title: "디랩 성적관리 시스템 실제 성공 사례 분석",
            content: "성적 향상 사례로 분석한 디랩의 학습 및 성적 관리 시스템 안내",
            url: "https://m.blog.naver.com/national137/224154666538"
          },
          {
            title: "2027 수시 지원 전략 설명회 (디랩 목동)",
            content: "7~8월 본격적인 수시 원서 접수를 앞두고 진행하는 2027 수시 지원 성공 전략 설명회",
            url: "https://m.blog.naver.com/national137/224331841418"
          },
          {
            title: "2027학년도 수시 지원 전략 재원생 대상 설명회",
            content: "변화하는 입시 흐름에 따른 디랩 재원생 대상 맞춤형 수시 지원 전략 설명회",
            url: "https://m.blog.naver.com/national137/224327150929"
          },
          {
            title: "7월 더프 외부생 모집",
            content: "전과목 전 범위로 진행되는 7월 더프리미엄 모의고사 일정 및 외부생 접수 안내 (시행일 7/18)",
            url: "https://m.place.naver.com/place/1960381657/feed"
          },
          {
            title: "2027 수능 반수시즌 반수학원 추천｜디랩 대구점·대전점 신규 오픈",
            content: "6평 이후 반수생 대상 디랩 신규 지점(대구/대전) 개원 기념 이벤트 및 반수반 소개",
            url: "https://m.blog.naver.com/national137/224313879109"
          }
        ],
        guide: "디랩의 전방위적 공세에 대응한 3대 전술 지침:\n1. [실전 모의고사] 디랩의 이감 시즌5/더프 응시권 유치에 대응해 자사 재원생이 누리는 풀시즌 모의고사 라인업과 1:1 약점 피드백 클리닉의 완성도를 부각하십시오.\n2. [수시 설명회] 타사 설명회 모객에 대항하여, 자사 이투스 교육평가연구소의 입시 공신력 및 개별 생기부 정밀 진단 수시 컨설팅을 전면에 홍보하십시오.\n3. [성적/반수 관리] 타사의 성공사례 홍보 및 반수생 모집에 맞서, 분당정자점의 명문대 성공 포스터를 입구에 배치하고 자사 특유의 엄격한 생활 관리 루틴 및 장학 제도를 대외적으로 부각하십시오."
      }
    ];
  }

  if (cleanName.includes("목동")) {
    return [
      {
        name: "잇올 스파르타 목동센터",
        urgency: "high",
        metrics: { ours: { mgmt: 90, content: 95, facility: 85 }, theirs: { mgmt: 94, content: 78, facility: 92 } },
        trend: "최근 개방형 1인 자습 공간 비중을 늘리고 장기 결제 할인 프로모션을 공격적으로 진행하여 원생 이탈을 유도하고 있습니다.",
        promotions: [
          { title: "3일 무료체험 패스 배포", content: "신규 수강등록 전 학습관리 시스템을 직접 경험해보는 3일 무료체험 이벤트 진행 중", url: "https://www.etoos.com" },
          { title: "재수반 장학 혜택 최대 50%", content: "6월 모의평가 성적 우수자 대상 학원비 최대 50% 반액 장학 혜택 선착순 마감 임박", url: "https://www.etoos.com" }
        ],
        guide: "자사의 엄격한 교과 질의응답 및 입시 모의고사 상담 혜택을 전면에 내세우고, 노후 학습실 조명 교체 등 시설 만족도를 적극 개선하십시오."
      },
      {
        name: "대성 디랩 목동",
        urgency: "medium",
        metrics: { ours: { mgmt: 90, content: 95, facility: 85 }, theirs: { mgmt: 85, content: 92, facility: 88 } },
        trend: "대성 학원 입시 패스 상품과의 패키지 할인 혜택을 기반으로 대형 입시 설명회를 수시 개최하고 있습니다.",
        promotions: [
          { title: "대성입시콘텐츠 프리패스", content: "디랩 목동 수강생 전원 대성마이맥 올패스 결합 수강 할인 쿠폰 증정 캠페인", url: "https://www.etoos.com" }
        ],
        guide: "대성 패스 대비 자사 모의고사(이투스 등)의 독자적 가치를 설명하고, 주 2회 일대일 멘토 피드백 서비스의 디테일을 강조하십시오."
      }
    ];
  }

  if (cleanName.includes("송도") || cleanName.includes("인천송도")) {
    return [
      {
        name: "수능선배 송도점",
        urgency: "high",
        blogUrl: "https://blog.naver.com/bioochemistry",
        metrics: { ours: { mgmt: 88, content: 90, facility: 90 }, theirs: { mgmt: 95, content: 82, facility: 94 } },
        trend: "자체 출입 통제 관리 프로그램 및 태블릿 자습 관리 모니터링 시스템을 도입해 밀착 관리 마케팅을 전개하고 있습니다.",
        promotions: [
          { title: "태블릿 출결 감시 무료 런칭", content: "실시간 학습 화면 원격 모니터링 서비스 추가 비용 없이 상시 무료 서비스 제공", url: "https://www.etoos.com" }
        ],
        guide: "자사의 엄격한 지각/벌점 관리 규정과 실시간 출결 학부모 알림 기능을 강조하여 관리 불안감을 종식시키십시오."
      },
      {
        name: "잇올 스파르타 송도센터",
        urgency: "medium",
        metrics: { ours: { mgmt: 88, content: 90, facility: 90 }, theirs: { mgmt: 92, content: 78, facility: 92 } },
        trend: "송도 국제도시 학부모 타깃의 프리미엄 교육 인테리어 리뉴얼을 마친 상태로 대대적인 체험권 배포 중입니다.",
        promotions: [
          { title: "리뉴얼 기념 1일 무료체험", content: "신규 단독 자습 전용석 100석 증설 기념 무료 1일 좌석 대여 이벤트 진행 중", url: "https://www.etoos.com" }
        ],
        guide: "쾌적성 부문에서 자사 시설의 편의 요소를 블로그에 재소개하고, 장학생 성공 수기 포스터를 원내에 전면 배치하십시오."
      }
    ];
  }

  return [
    {
      name: `${cleanName} 인근 메이저 관리형 독서실`,
      urgency: "high",
      metrics: { ours: { mgmt: 88, content: 90, facility: 85 }, theirs: { mgmt: 90, content: 70, facility: 90 } },
      trend: "인근 지역에서 시설 중심의 프리미엄 홍보 및 저가 결제 프로모션을 병행하며 원생을 유입시키고 있습니다.",
      promotions: [
        { title: "첫 달 등록 10% 얼리버드 할인", content: "오픈 기념으로 첫 달 수강료 결제 시 10% 즉시 감면 및 입시 플래너 무료 증정", url: "https://www.etoos.com" }
      ],
      guide: "자사 관리 프로그램(프렌즈/체험단 성과 지표)의 검증된 가치와 AI 멘토링 역량을 앞세워 입시 전문 브랜드로 차별화하십시오."
    }
  ];
}

function getBranchMarketingStatus(branchName, rawTabs) {
  if (!branchName) return null;
  const cleanName = String(branchName).trim();

  // 1. 247프렌즈
  const friendsTab = rawTabs.find(t => t.name === "247프렌즈" || t.id === "tab-1");
  const friendsRow = friendsTab?.rows?.find(r => r.branch?.trim() === cleanName);
  const totalFriendsEvents = friendsTab?.events?.length || 7;
  const branchFriendsCount = friendsRow ? Object.values(friendsRow.eventValues || {}).filter(ev => Number(ev.participants || 0) > 0).length : 0;
  const friendsActive = branchFriendsCount > 0;

  // 2. 247체험단
  const experienceTab = rawTabs.find(t => t.name === "247체험단" || t.id === "tab-2");
  const experienceRow = experienceTab?.rows?.find(r => r.branch?.trim() === cleanName);
  const totalExperienceEvents = experienceTab?.events?.length || 7;
  const branchExperienceCount = experienceRow ? Object.values(experienceRow.eventValues || {}).filter(ev => Number(ev.participants || 0) > 0).length : 0;
  const experienceActive = branchExperienceCount > 0;

  // 3. SNS 진단표
  const snsTab = rawTabs.find(t => t.kind === "special-social");
  const snsRow = snsTab?.socialRows?.find(r => r.branch?.trim() === cleanName);
  let snsSummary = null;
  if (snsRow) {
    try {
      snsSummary = summarizeSnsRow(snsRow);
    } catch (e) {
      snsSummary = { hasBlog: !!(snsRow.blogUrl && snsRow.blogUrl.trim() !== ""), hasInstagram: !!(snsRow.instagramUrl && snsRow.instagramUrl.trim() !== ""), grade: snsRow.grade || "C", blogRecentPosts: "0", instagramRecentPosts: "0", finalScore: snsRow.finalScore || "0" };
    }
  }
  const snsActive = snsSummary ? (snsSummary.hasBlog || snsSummary.hasInstagram) : false;
  const snsScore = snsSummary ? snsSummary.finalScore : "0";
  const snsGrade = snsSummary ? snsSummary.grade : "-";

  // 4. 협업이벤트
  const collabTab = rawTabs.find(t => t.kind === "special-collab");
  const collabRow = collabTab?.collabRows?.find(r => (r.values?.["지점"] || "").trim() === cleanName);
  const urlColumns = (collabTab?.collabColumns || []).filter(c => c !== "지역" && c !== "지점");
  const branchCollabCount = collabRow
    ? urlColumns.filter(col => collabRow.values?.[col] && String(collabRow.values[col]).trim() !== "").length
    : 0;
  const collabActive = branchCollabCount > 0;

  // 5. 지점시설영상
  const facilityTab = rawTabs.find(t => t.kind === "special-facility");
  const facilityRow = facilityTab?.facilityRows?.find(r => r.branch?.trim() === cleanName);
  const hasFacilityVideo = facilityRow?.url && facilityRow.url.trim() !== "" ? "O" : "X";
  const facilityActive = hasFacilityVideo === "O";

  // 6. 멘토단 및 장학생
  const mentorTab = rawTabs.find(t => t.kind === "special-mentor");
  const mentorCount = mentorTab?.mentorRows?.filter(r => r.branch?.trim() === cleanName).length || 0;
  const mentorActive = mentorCount > 0;

  const activePrograms = [
    { name: "247프렌즈", active: friendsActive, ratio: `${totalFriendsEvents}/${branchFriendsCount}` },
    { name: "247체험단", active: experienceActive, ratio: `${totalExperienceEvents}/${branchExperienceCount}` },
    { name: "SNS 진단표", active: snsActive, ratio: `${snsScore}점 (${snsGrade})` },
    { name: "협업이벤트", active: collabActive, ratio: `${urlColumns.length}/${branchCollabCount}` },
    { name: "지점시설영상", active: facilityActive, ratio: hasFacilityVideo },
    { name: "멘토단 및 장학생", active: mentorActive, ratio: `${mentorCount}명` }
  ];

  const activeCount = activePrograms.filter(p => p.active).length;
  const participationRate = Math.round((activeCount / 6) * 100);

  return {
    activePrograms,
    activeCount,
    participationRate,
    snsSummary,
    rawSnsRow: snsRow,
    friendsRatio: `${branchFriendsCount}/${totalFriendsEvents}`,
    experienceRatio: `${branchExperienceCount}/${totalExperienceEvents}`,
    snsRatio: `${snsScore}점 (${snsGrade})`,
    collabRatio: `${branchCollabCount}/${urlColumns.length}`,
    facilityRatio: hasFacilityVideo,
    mentorRatio: `${mentorCount}명`
  };
}

function generateLocalAdviserText(branch, statusInfo, competitors, trendData) {
  const hasVideo = statusInfo.facilityRatio === "O";
  const scholarCount = statusInfo.mentorRatio || "0명";
  const snsScore = statusInfo.snsRatio;
  const comp1Name = competitors[0]?.name || "경쟁 학원";

  return `### 🤖 [${branch} 지점] 로컬 마케팅 AI Adviser 코칭 리포트

#### 1. 자사 마케팅 현황 진단 피드백
- **247 프로그램 참여율:** 자사 마케팅 총 6개 부문 중 **${statusInfo.activeCount}개**가 운영 중이며, 참여율은 **${statusInfo.participationRate}%** 입니다.
- **SNS 활성화 수준:** 블로그/인스타 진단표 점수는 **${snsScore}**로 평가되었습니다. ${statusInfo.participationRate < 60 ? "현재 마케팅 채널 활성화가 다소 미흡하오니 247체험단 및 협업이벤트를 추가 활성화해 주십시오." : "우수한 마케팅 채널 운영도를 보이고 있으므로 포스팅 주기를 정기적으로 다져 입시 브랜딩을 강화해야 합니다."}
- **지점 시설영상 도입 상태 (${statusInfo.facilityRatio}):** 시설 동영상 홍보가 ${hasVideo ? "정상 운영 중(O)입니다. 지점 블로그 및 플레이스 최상단에 영상을 고정 노출하여 내원율을 극대화하십시오." : "미운영 중(X)입니다. 학부모와 학생들이 시설을 둘러볼 수 있게 본사 템플릿 시설 영상을 블로그에 조속히 탑재할 것을 권장합니다."}
- **장학 성과 (${scholarCount}):** 당 지점은 누적 **${scholarCount}**의 명문대 합격생(멘토/장학생)을 배출하였습니다. 이는 로컬 학부모들이 가장 신뢰하는 증거이므로 블로그 타이틀 및 플레이스 소식 란에 이 합격자 성과를 최상단 배너로 홍보해야 합니다.

#### 2. 네이버 검색 트렌드 분석 및 상권 경쟁
- **지역 인지도 경쟁:** 네이버 검색량 분석 결과, 최근 6개월간 **${branch} 이투스247**의 관심 지수가 평균 45 수준이며, 경쟁사인 **${comp1Name}**(평균 52) 대비 다소 경합 중입니다.
- **타깃 분석:** 해당 상권은 **학부모(40~50대)의 PC 검색 비중(65%)**이 두드러집니다. 따라서 모바일 감성의 숏폼보다는 **블로그의 정교한 입시 요강 설명회 및 밀착 관리 사례**가 실제 등록 전환에 강력한 영향력을 미칩니다.

#### 3. 시즌성 경보 및 AI 실행 전략
- **시즌별 타겟팅:** 곧 다가올 모의고사 직후 및 수시 접수 시즌에는 '윈터스쿨', '반수반' 키워드 유입이 평소 대비 2.5배 급증합니다. 지점 블로그 프로모션 노출 시점을 최소 2주 앞당기십시오.
- **맞춤 권장 행동:** 인근 경쟁 학원의 최신 프로모션인 '무료 체험권' 및 '교재비 면제' 공세에 맞서 자사는 **소수정예 밀착 입시 코칭 브랜딩**으로 차별화를 두는 것이 등록 단가를 보존하고 장기 락인을 유도하는 최선책입니다.`;
}

function parseCollabColumnLabel(label) {
  const rawLabel = String(label ?? "").trim();
  if (!rawLabel || rawLabel === "지역" || rawLabel === "지점") {
    return { eventName: "", channel: "" };
  }

  const suffixes = ["홈페이지", "블로그", "인스타/언론기사"];
  const suffix = suffixes.find((item) => rawLabel.endsWith(item)) || "";
  const eventName = suffix ? rawLabel.slice(0, -suffix.length).trim() : rawLabel;

  return {
    eventName,
    channel: suffix || "URL"
  };
}

function getCollabColumnThemeStyle(label) {
  if (label === "지역" || label === "지점") return undefined;

  const { eventName } = parseCollabColumnLabel(label);
  if (!eventName) return undefined;

  const hash = [...eventName].reduce((sum, char) => sum + char.charCodeAt(0), 0);
  const theme = collabEventColorPalette[hash % collabEventColorPalette.length];

  return {
    "--collab-column-bg": theme.bg,
    "--collab-column-border": theme.border
  };
}

function groupCollabColumns(columns = []) {
  const normalized = normalizeCollabColumns(columns);
  const groups = [];
  const groupMap = new Map();

  normalized.forEach((column) => {
    if (column === "지역" || column === "지점") return;
    const { eventName, channel } = parseCollabColumnLabel(column);
    if (!eventName) return;

    if (!groupMap.has(eventName)) {
      const group = { eventName, columns: [] };
      groupMap.set(eventName, group);
      groups.push(group);
    }

    groupMap.get(eventName).columns.push({
      key: column,
      label: channel || column
    });
  });

  return groups;
}

function buildCollabSummary(tab) {
  const columns = normalizeCollabColumns(tab?.collabColumns || []);
  const urlColumns = columns.filter((column) => column !== "지역" && column !== "지점");
  const orderedEventNames = [...new Set(urlColumns.map((column) => parseCollabColumnLabel(column).eventName).filter(Boolean))];
  const eventMap = new Map();
  const branchRows = [];
  const allBranches = [];

  orderedEventNames.forEach((eventName, index) => {
    eventMap.set(eventName, {
      id: eventName,
      label: eventName,
      branchCount: 0,
      urlCount: 0,
      order: index
    });
  });

  (tab?.collabRows || []).forEach((row) => {
    const values = row.values || {};
    const region = normalizeRegionLabel(String(values["지역"] || "").trim());
    const branch = String(values["지점"] || "").trim();
    if (!branch) return;

    const events = [];
    let urlCount = 0;

    urlColumns.forEach((column) => {
      const url = String(values[column] || "").trim();
      const { eventName, channel } = parseCollabColumnLabel(column);
      if (!eventName) return;

      if (url) {
        let eventEntry = events.find((item) => item.name === eventName);
        if (!eventEntry) {
          eventEntry = { name: eventName, links: [] };
          events.push(eventEntry);
          eventMap.get(eventName).branchCount += 1;
        }

        eventEntry.links.push({ label: channel, url });
        eventMap.get(eventName).urlCount += 1;
        urlCount += 1;
      }
    });

    branchRows.push({
      id: row.id,
      region,
      branch,
      urlCount,
      events
    });
    allBranches.push(branch);
  });

  const activeBranches = branchRows.filter((row) => row.urlCount > 0).length;
  const totalUrls = branchRows.reduce((sum, row) => sum + row.urlCount, 0);
  const eventOverview = [...eventMap.values()].sort((a, b) => a.order - b.order);

  return {
    totalBranches: branchRows.length,
    activeBranches,
    inactiveBranches: Math.max(branchRows.length - activeBranches, 0),
    totalUrls,
    uniqueEvents: eventOverview.length,
    branchRows,
    branchOptions: branchRows.map((row) => row.branch).sort((a, b) => a.localeCompare(b, "ko")),
    groupedBranches: branchRows.reduce((acc, row) => {
      if (!acc.has(row.region)) acc.set(row.region, []);
      acc.get(row.region).push(row.branch);
      return acc;
    }, new Map()),
    eventOverview
  };
}

function buildFacilitySummary(tab) {
  const branchRows = (tab?.facilityRows || [])
    .map((row) => ({
      id: row.id,
      region: normalizeRegionLabel(String(row.region || "").trim()),
      branch: String(row.branch || "").trim(),
      url: String(row.url || "").trim()
    }))
    .filter((row) => row.branch);

  const activeBranchRows = branchRows.filter((row) => row.url);

  return {
    totalBranches: branchRows.length,
    activeBranches: activeBranchRows.length,
    inactiveBranches: Math.max(branchRows.length - activeBranchRows.length, 0),
    totalUrls: activeBranchRows.length,
    branchRows,
    groupedBranches: branchRows.reduce((acc, row) => {
      if (!acc.has(row.region)) acc.set(row.region, []);
      acc.get(row.region).push(row);
      return acc;
    }, new Map())
  };
}

function migrateLegacyTab(tab) {
  if (!tab) return tab;
  let name = tab.name || "";
  if (name === "프렌즈") name = "247프렌즈";
  if (name === "체험단") name = "247체험단";
  if (name === "합격자취합") name = "합격자 취합";
  if (!name) {
    if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) name = "SNS 진단표";
    else if (tab.kind === SPECIAL_COLLAB_TAB_KIND) name = "협업이벤트";
    else if (tab.kind === SPECIAL_FACILITY_TAB_KIND) name = "지점시설영상";
    else if (tab.kind === SPECIAL_MENTOR_TAB_KIND) name = "멘토단 및 장학생";
    else name = "이름 없는 탭";
  }

  if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) {
    return {
      id: tab.id || createId("tab"),
      name,
      kind: SPECIAL_SOCIAL_TAB_KIND,
      events: [],
      rows: [],
      socialRows: Array.isArray(tab.socialRows) && tab.socialRows.length > 0
        ? tab.socialRows.map((row, index) => createSpecialSocialRow({ ...row, id: row?.id || createId(`social-row-${index}`) }))
        : [createSpecialSocialRow()]
    };
  }

  if (tab.kind === SPECIAL_COLLAB_TAB_KIND) {
    const collabColumns = normalizeCollabColumns(tab.collabColumns || defaultCollabColumns);
    return {
      id: tab.id || createId("tab"),
      name,
      kind: SPECIAL_COLLAB_TAB_KIND,
      events: [],
      rows: [],
      collabColumns,
      collabRows: Array.isArray(tab.collabRows) && tab.collabRows.length > 0
        ? tab.collabRows.map((row, index) =>
            createSpecialCollabRow(collabColumns, {
              ...(row?.values || {}),
              ...row,
              id: row?.id || createId(`collab-row-${index}`)
            })
          )
        : [createSpecialCollabRow(collabColumns)]
    };
  }

  if (tab.kind === SPECIAL_FACILITY_TAB_KIND) {
    return {
      id: tab.id || createId("tab"),
      name,
      kind: SPECIAL_FACILITY_TAB_KIND,
      events: [],
      rows: [],
      facilityColumns: defaultFacilityColumns,
      facilityRows: Array.isArray(tab.facilityRows) && tab.facilityRows.length > 0
        ? tab.facilityRows.map((row, index) =>
            createSpecialFacilityRow({
              ...row,
              id: row?.id || createId(`facility-row-${index}`)
            })
          )
        : [createSpecialFacilityRow()]
    };
  }

  if (tab.kind === SPECIAL_MENTOR_TAB_KIND) {
    const sortedMentorRows = (Array.isArray(tab.mentorRows) ? tab.mentorRows : [])
      .map((row, index) => createSpecialMentorRow({ ...row, id: row?.id || createId(`mentor-row-${index}`) }))
      .sort((a, b) => {
        const yA = parseInt(a.year, 10) || 0;
        const yB = parseInt(b.year, 10) || 0;
        if (yB !== yA) return yB - yA;
        return (a.name || "").localeCompare(b.name || "", "ko");
      });
    return {
      id: tab.id || createId("tab"),
      name,
      kind: SPECIAL_MENTOR_TAB_KIND,
      events: [],
      rows: [],
      mentorRows: sortedMentorRows
    };
  }

  if (Array.isArray(tab.events)) {
    const events = tab.events.map((event, index) => ({
      id: event?.id || createId(`event-${index}`),
      name: event?.name || `이벤트 ${index + 1}`
    }));
    const eventIds = events.map((event) => event.id);
    const rows = Array.isArray(tab.rows)
      ? tab.rows.map((row, index) =>
          createRow(eventIds, {
            id: row?.id || createId(`row-${index}`),
            region: row?.region || "",
            branch: row?.branch || "",
            eventValues: row?.eventValues || {}
          })
        )
      : [createRow(eventIds)];

    return {
      id: tab.id || createId("tab"),
      name,
      kind: tab.kind || "default",
      events,
      rows
    };
  }

  const columns = Array.isArray(tab.columns) ? tab.columns.map((column) => String(column)) : [];
  const rows = Array.isArray(tab.rows) ? tab.rows : [];
  const regionIndex = Math.max(0, columns.findIndex((column) => column.includes("지역")));
  const branchIndex = Math.max(0, columns.findIndex((column) => column.includes("지점")));
  const statusColumns = columns
    .map((column, index) => ({ column, index }))
    .filter((item) => item.column.includes("참여여부"));
  const participantColumns = columns
    .map((column, index) => ({ column, index }))
    .filter((item) => /(참석인원|참여인원|인원)/.test(item.column));
  const usedParticipantIndexes = new Set();

  let eventDefinitions = statusColumns.map((item, index) => {
    const label = stripEventSuffix(item.column);
    let participantMatch = participantColumns.find(
      (candidate) => !usedParticipantIndexes.has(candidate.index) && stripEventSuffix(candidate.column) === label
    );

    if (!participantMatch) {
      participantMatch = participantColumns.find((candidate) => !usedParticipantIndexes.has(candidate.index));
    }

    if (participantMatch) {
      usedParticipantIndexes.add(participantMatch.index);
    }

    return {
      id: createId(`event-${index}`),
      name: label || (statusColumns.length === 1 ? name || "기본 이벤트" : `이벤트 ${index + 1}`),
      statusIndex: item.index,
      participantIndex: participantMatch?.index ?? -1
    };
  });

  if (eventDefinitions.length === 0 && participantColumns.length > 0) {
    eventDefinitions = participantColumns.map((item, index) => ({
      id: createId(`event-${index}`),
      name: stripEventSuffix(item.column) || `이벤트 ${index + 1}`,
      statusIndex: -1,
      participantIndex: item.index
    }));
  }

  if (eventDefinitions.length === 0) {
    eventDefinitions = [
      {
        id: createId("event-default"),
        name: name || "기본 이벤트",
        statusIndex: -1,
        participantIndex: -1
      }
    ];
  }

  const migratedEvents = eventDefinitions.map(({ id, name }) => ({ id, name }));
  const migratedRows = rows.length
    ? rows.map((legacyRow, rowIndex) => {
        const eventValues = Object.fromEntries(
          eventDefinitions.map((event) => {
            const participants =
              event.participantIndex === -1
                ? "0"
                : normalizeParticipantValue(legacyRow[event.participantIndex] ?? "0");
            const rawStatus =
              event.statusIndex === -1 ? "" : String(legacyRow[event.statusIndex] ?? "").trim().toUpperCase();
            const status = rawStatus === "O" || Number(participants) > 0 ? "O" : "X";
            return [event.id, { status, participants }];
          })
        );

        return createRow(migratedEvents.map((event) => event.id), {
          id: createId(`row-${rowIndex}`),
          region: String(legacyRow[regionIndex] ?? ""),
          branch: String(legacyRow[branchIndex] ?? ""),
          eventValues
        });
      })
    : [createRow(migratedEvents.map((event) => event.id))];

  return {
    id: tab.id || createId("tab"),
    name,
    kind: "default",
    events: migratedEvents,
    rows: migratedRows
  };
}

function ensureSpecialInputTabs(tabs) {
  const seededBranchRows = [
    ...tabs
      .filter((tab) => !isSpecialTabKind(tab.kind))
      .flatMap((tab) =>
        tab.rows
          .map((row) => ({
            region: row.region?.trim?.() || "",
            branch: row.branch?.trim?.() || ""
          }))
          .filter((row) => row.branch)
      )
      .reduce((acc, row) => {
        if (!acc.some((item) => item.branch === row.branch)) {
          acc.push(row);
        }
        return acc;
      }, [])
  ];

  const nextTabs = [...tabs];

  if (!nextTabs.some((tab) => tab.kind === SPECIAL_SOCIAL_TAB_KIND)) {
    nextTabs.push(createSpecialSocialTab("tab-social-1", "SNS 진단표", seededBranchRows.map((row) => ({ branch: row.branch }))));
  }

  if (!nextTabs.some((tab) => tab.kind === SPECIAL_COLLAB_TAB_KIND)) {
    nextTabs.push(createSpecialCollabTab("tab-collab-1", "협업이벤트", seededBranchRows.map((row) => ({ 지역: row.region, 지점: row.branch }))));
  }

  if (!nextTabs.some((tab) => tab.kind === SPECIAL_FACILITY_TAB_KIND)) {
    nextTabs.push(createSpecialFacilityTab("tab-facility-1", "지점시설영상", seededBranchRows));
  }

  if (!nextTabs.some((tab) => tab.kind === SPECIAL_MENTOR_TAB_KIND)) {
    nextTabs.push(createSpecialMentorTab("tab-mentor-1", "멘토단 및 장학생", []));
  }

  return nextTabs;
}

function getSortedRawTabs(tabs) {
  if (!Array.isArray(tabs)) return [];
  return tabs.map((tab) => {
    if (tab.kind === SPECIAL_MENTOR_TAB_KIND && Array.isArray(tab.mentorRows)) {
      const sorted = [...tab.mentorRows].sort((a, b) => {
        const yA = parseInt(a.year, 10) || 0;
        const yB = parseInt(b.year, 10) || 0;
        if (yB !== yA) return yB - yA;
        return (a.name || "").localeCompare(b.name || "", "ko");
      });
      return { ...tab, mentorRows: sorted };
    }
    return tab;
  });
}

function normalizeRawTabs(rawTabs) {
  if (!Array.isArray(rawTabs) || rawTabs.length === 0) {
    return initialTabs;
  }

  return getSortedRawTabs(ensureSpecialInputTabs(rawTabs.map(migrateLegacyTab)));
}

function summarizeTab(tab) {
  if (tab.kind === SPECIAL_MENTOR_TAB_KIND) {
    const filledRows = (tab.mentorRows || []).filter((row) => row.name.trim()).length;
    const mentorCount = (tab.mentorRows || []).filter((row) => row.name.trim() && row.isMentor).length;
    return {
      rows: filledRows,
      events: 0,
      branches: new Set((tab.mentorRows || []).map((row) => row.branch.trim()).filter(Boolean)).size,
      participants: mentorCount,
      activeBranches: new Set((tab.mentorRows || []).filter((row) => row.isMentor).map((row) => row.branch.trim()).filter(Boolean)).size
    };
  }

  if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) {
    const filledRows = (tab.socialRows || []).filter((row) => row.branch.trim()).length;
    return {
      rows: filledRows,
      events: 0,
      branches: filledRows,
      participants: 0,
      activeBranches: 0
    };
  }

  if (tab.kind === SPECIAL_COLLAB_TAB_KIND) {
    const columns = normalizeCollabColumns(tab.collabColumns || []);
    const urlColumns = columns.filter((column) => column !== "지역" && column !== "지점");
    const rows = tab.collabRows || [];
    const branchSet = new Set();
    const activeBranchSet = new Set();
    let filledRows = 0;
    let urlCount = 0;

    rows.forEach((row) => {
      const values = row.values || {};
      const region = String(values["지역"] || "").trim();
      const branch = String(values["지점"] || "").trim();
      if (!region && !branch) return;
      filledRows += 1;
      if (branch) branchSet.add(branch);

      const hasAnyUrl = urlColumns.some((column) => String(values[column] || "").trim());
      if (hasAnyUrl && branch) activeBranchSet.add(branch);
      urlColumns.forEach((column) => {
        if (String(values[column] || "").trim()) {
          urlCount += 1;
        }
      });
    });

    return {
      rows: filledRows,
      events: [...new Set(urlColumns.map((column) => parseCollabColumnLabel(column).eventName).filter(Boolean))].length,
      branches: branchSet.size,
      participants: urlCount,
      activeBranches: activeBranchSet.size
    };
  }

  const uniqueBranches = new Set();
  let filledRows = 0;
  let participantTotal = 0;
  let activeBranches = 0;

  tab.rows.forEach((row) => {
    const hasContent = row.region.trim() || row.branch.trim();
    if (!hasContent) return;
    filledRows += 1;
    uniqueBranches.add(row.branch.trim());

    tab.events.forEach((event) => {
      participantTotal += Number(row.eventValues?.[event.id]?.participants || 0);
    });
  });

  return {
    rows: filledRows,
    events: tab.events.length,
    branches: [...uniqueBranches].filter(Boolean).length,
    participants: participantTotal,
    activeBranches
  };
}

function buildBranchOverview(rawTabs) {
  const branchMap = new Map();

  rawTabs.filter((tab) => !isSpecialTabKind(tab.kind)).forEach((tab) => {
    tab.rows.forEach((row) => {
      const branch = row.branch.trim();
      if (!branch) return;

      if (!branchMap.has(branch)) {
        branchMap.set(branch, {
          region: normalizeRegionLabel(row.region.trim()),
          branch,
          activePlans: new Set(),
          activeEvents: new Set(),
          totalParticipants: 0
        });
      }

      const current = branchMap.get(branch);
      current.region = current.region || normalizeRegionLabel(row.region.trim());

      tab.events.forEach((event) => {
        const eventValue = row.eventValues?.[event.id] || { status: "X", participants: "0" };
        const participants = Number(eventValue.participants || 0);
        const isActive = eventValue.status === "O" || participants > 0;

        current.totalParticipants += participants;
        if (isActive) {
          current.activePlans.add(tab.name);
          current.activeEvents.add(`${tab.name} / ${event.name}`);
        }
      });
    });
  });

  return [...branchMap.values()]
    .map((item) => ({
      region: item.region,
      branch: item.branch,
      activePlanCount: item.activePlans.size,
      activeEventCount: item.activeEvents.size,
      totalParticipants: item.totalParticipants,
      activePlans: [...item.activePlans]
    }))
    .sort((a, b) => b.totalParticipants - a.totalParticipants || b.activeEventCount - a.activeEventCount);
}

function buildEventOverview(rawTabs) {
  return rawTabs
    .filter((tab) => !isSpecialTabKind(tab.kind))
    .flatMap((tab) =>
      tab.events.map((event) => {
        let branchCount = 0;
        let participants = 0;

        tab.rows.forEach((row) => {
          const eventValue = row.eventValues?.[event.id] || { status: "X", participants: "0" };
          const participantCount = Number(eventValue.participants || 0);
          if (eventValue.status === "O" || participantCount > 0) {
            branchCount += 1;
          }
          participants += participantCount;
        });

        return {
          id: `${tab.id}-${event.id}`,
          tabId: tab.id,
          tabName: tab.name,
          eventName: event.name,
          branchCount,
          participants
        };
      })
    )
    .sort((a, b) => b.participants - a.participants || b.branchCount - a.branchCount);
}

function buildRegionOverview(branchOverview) {
  return branchOverview
    .reduce((acc, branch) => {
      const region = branch.region || "기타";
      if (!acc.has(region)) {
        acc.set(region, {
          region,
          branchCount: 0,
          activeBranches: 0,
          topBranch: "-"
        });
      }

      const current = acc.get(region);
      current.branchCount += 1;
      if (branch.totalParticipants > 0) {
        current.activeBranches += 1;
      }
      if (current.topBranch === "-" || branch.totalParticipants > (current.topBranchParticipants || 0)) {
        current.topBranch = branch.branch;
        current.topBranchParticipants = branch.totalParticipants;
      }
      return acc;
    }, new Map())
    .values();
}

function buildDashboardData(rawTabs) {
  const totals = rawTabs.map(summarizeTab);
  const totalRows = totals.reduce((sum, item) => sum + item.rows, 0);
  const totalEvents = totals.reduce((sum, item) => sum + item.events, 0);
  const branchOverview = buildBranchOverview(rawTabs);
  const uniqueBranches = new Set(branchOverview.map((item) => item.branch));
  const activeBranches = branchOverview.filter((branch) => branch.totalParticipants > 0).length;
  const inactiveBranches = branchOverview.filter((branch) => branch.totalParticipants === 0).length;
  const regionOverview = [...buildRegionOverview(branchOverview)]
    .map((item) => ({
      ...item,
      activationRate: item.branchCount > 0 ? Math.round((item.activeBranches / item.branchCount) * 100) : 0,
      topBranchParticipants: item.topBranchParticipants || 0
    }))
    .sort((a, b) => b.activeBranches - a.activeBranches || b.branchCount - a.branchCount);
  const regionOptions = [...new Set(regionOverview.map((item) => item.region).filter(Boolean))].sort((a, b) => a.localeCompare(b, "ko"));
  const topBranches = [...branchOverview].slice(0, 7);
  const attentionBranches = [...branchOverview]
    .filter((branch) => branch.totalParticipants === 0 || branch.activeEventCount <= 1)
    .sort((a, b) => a.totalParticipants - b.totalParticipants || a.activeEventCount - b.activeEventCount)
    .slice(0, 7);

  return {
    tabCount: rawTabs.length,
    totalRows,
    totalEvents,
    uniqueBranches: uniqueBranches.size,
    activeBranches,
    inactiveBranches,
    tabStats: rawTabs.map((tab, index) => ({
      id: tab.id,
      name: tab.name,
      ...totals[index]
    })),
    branchOverview,
    topBranches,
    attentionBranches,
    regionOverview,
    regionOptions
  };
}

function getRegionTone(item, maxBranches) {
  if (!item || maxBranches <= 0) return "var(--map-tone-1)";
  const ratio = item.activeBranches / maxBranches;
  if (ratio > 0.8) return "var(--map-tone-5)";
  if (ratio > 0.6) return "var(--map-tone-4)";
  if (ratio > 0.4) return "var(--map-tone-3)";
  if (ratio > 0.2) return "var(--map-tone-2)";
  return "var(--map-tone-1)";
}

function getBranchGrade(score) {
  if (score >= 85) return "A그룹";
  if (score >= 70) return "B그룹";
  if (score >= 50) return "C그룹";
  return "D그룹";
}

function isMissingChannelUrl(value) {
  const normalized = String(value ?? "").trim();
  return !normalized || normalized === "0" || normalized === "-" || normalized.toUpperCase() === "N/A" || normalized === "#N/A";
}

function getRecentActivityScore(count, thresholds) {
  const value = Number(count || 0);
  if (value >= thresholds[0][0]) return thresholds[0][1];
  if (value >= thresholds[1][0]) return thresholds[1][1];
  if (value >= thresholds[2][0]) return thresholds[2][1];
  if (value >= thresholds[3][0]) return thresholds[3][1];
  return 0;
}

function isDormantSince(dateString, baseDateString, thresholdDays = 30) {
  if (!dateString) return false;
  const base = new Date(baseDateString);
  const target = new Date(dateString);
  if (Number.isNaN(base.getTime()) || Number.isNaN(target.getTime())) return false;
  const diff = Math.floor((base - target) / (1000 * 60 * 60 * 24));
  return diff > thresholdDays;
}

function summarizeSnsRow(row, baseDate = snsEvaluationBaseDate) {
  const hasBlog = !isMissingChannelUrl(row.blogUrl);
  const hasInstagram = !isMissingChannelUrl(row.instagramUrl);

  const blogActivity = hasBlog
    ? Math.max(
        0,
        getRecentActivityScore(row.blogRecentPosts, [[8, 30], [5, 25], [3, 20], [1, 10]]) -
          (isDormantSince(row.blogLastPosted, baseDate, 60) ? 15 : 0)
      )
    : 0;
  const blogReaction = hasBlog ? Number(row.blogVisitScore || 0) : 0;
  const blogScore = hasBlog ? Number(((blogActivity + blogReaction) / 35 * 50).toFixed(1)) : 0;

  const instagramActivity = hasInstagram
    ? Math.max(
        0,
        getRecentActivityScore(row.instagramRecentPosts, [[12, 30], [8, 25], [4, 20], [1, 10]]) -
          (isDormantSince(row.instagramLastPosted, baseDate, 30) ? 5 : 0)
      )
    : 0;
  const instagramContent = hasInstagram
    ? [
        row.instagramDesignScore,
        row.instagramReactionScore,
        row.profileSetupScore,
        row.featureUsageScore,
        row.ctaScore,
        row.linkHealthScore,
        row.brandInfoScore
      ].reduce((sum, value) => sum + Number(value || 0), 0)
    : 0;
  const instagramScore = hasInstagram ? Number(((instagramActivity + instagramContent) / 55 * 50).toFixed(1)) : 0;

  const combinedScore = blogScore === 0 && instagramScore === 0
    ? 0
    : blogScore === 0
      ? instagramScore * 2
      : instagramScore === 0
        ? blogScore * 2
        : blogScore + instagramScore;
  const missingPenalty = (hasBlog ? 0 : 10) + (hasInstagram ? 0 : 5);
  const finalScore = Number(Math.max(0, combinedScore - missingPenalty).toFixed(1));
  const grade = finalScore >= 80 ? "A" : finalScore >= 60 ? "B" : finalScore >= 40 ? "C" : "D";

  return {
    ...row,
    hasBlog,
    hasInstagram,
    blogActivity,
    blogReaction,
    blogScore,
    instagramActivity,
    instagramContent,
    instagramScore,
    finalScore,
    grade
  };
}

function formatImportDate(value) {
  if (!value && value !== 0) return "";

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return value.toISOString().slice(0, 10);
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const date = new Date(Date.UTC(parsed.y, parsed.m - 1, parsed.d));
      return date.toISOString().slice(0, 10);
    }
  }

  const date = new Date(value);
  if (!Number.isNaN(date.getTime())) {
    return date.toISOString().slice(0, 10);
  }

  return "";
}

function readImportedCell(sheet, column, rowIndex, options = {}) {
  const cell = sheet[`${column}${rowIndex}`];
  if (!cell) return "";

  if (options.preferHyperlink && cell.l?.Target) {
    return String(cell.l.Target).trim();
  }

  if (options.type === "date") {
    return formatImportDate(cell.v ?? cell.w ?? "");
  }

  const value = cell.v ?? cell.w ?? "";
  return value === null || value === undefined ? "" : String(value).trim();
}

function extractSnsRowsFromWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellStyles: false
  });

  const sheetName = workbook.SheetNames.find((name) => name.includes("입력")) ?? workbook.SheetNames[1] ?? workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet?.["!ref"]) {
    return { sheetName, rows: [] };
  }

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const rows = [];

  for (let rowIndex = 2; rowIndex <= range.e.r + 1; rowIndex += 1) {
    const branch = readImportedCell(sheet, "A", rowIndex);
    const blogUrl = readImportedCell(sheet, "B", rowIndex, { preferHyperlink: true });
    const instagramUrl = readImportedCell(sheet, "C", rowIndex, { preferHyperlink: true });
    const blogRecentPosts = readImportedCell(sheet, "D", rowIndex);
    const blogLastPosted = readImportedCell(sheet, "E", rowIndex, { type: "date" });
    const blogVisitScore = readImportedCell(sheet, "F", rowIndex);
    const instagramRecentPosts = readImportedCell(sheet, "G", rowIndex);
    const instagramLastPosted = readImportedCell(sheet, "H", rowIndex, { type: "date" });
    const instagramDesignScore = readImportedCell(sheet, "I", rowIndex);
    const instagramReactionScore = readImportedCell(sheet, "J", rowIndex);
    const profileSetupScore = readImportedCell(sheet, "K", rowIndex);
    const featureUsageScore = readImportedCell(sheet, "L", rowIndex);
    const ctaScore = readImportedCell(sheet, "M", rowIndex);
    const linkHealthScore = readImportedCell(sheet, "N", rowIndex);
    const brandInfoScore = readImportedCell(sheet, "O", rowIndex);
    const memo = readImportedCell(sheet, "P", rowIndex);

    const hasMeaningfulData = [
      branch,
      blogUrl,
      instagramUrl,
      blogRecentPosts,
      blogLastPosted,
      blogVisitScore,
      instagramRecentPosts,
      instagramLastPosted,
      instagramDesignScore,
      instagramReactionScore,
      profileSetupScore,
      featureUsageScore,
      ctaScore,
      linkHealthScore,
      brandInfoScore,
      memo
    ].some(Boolean);

    if (!hasMeaningfulData) continue;

    rows.push(
      createSpecialSocialRow({
        branch,
        blogUrl,
        instagramUrl,
        blogRecentPosts: blogRecentPosts || "0",
        blogLastPosted,
        blogVisitScore: blogVisitScore || "0",
        instagramRecentPosts: instagramRecentPosts || "0",
        instagramLastPosted,
        instagramDesignScore: instagramDesignScore || "0",
        instagramReactionScore: instagramReactionScore || "0",
        profileSetupScore: profileSetupScore || "0",
        featureUsageScore: featureUsageScore || "0",
        ctaScore: ctaScore || "0",
        linkHealthScore: linkHealthScore || "0",
        brandInfoScore: brandInfoScore || "0",
        memo
      })
    );
  }

  return { sheetName, rows };
}

function extractDefaultTabFromWorkbook(arrayBuffer, fallbackName = "불러온 이벤트") {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellStyles: false
  });

  const sheetName =
    workbook.SheetNames.find((name) => name.includes(fallbackName)) ??
    workbook.SheetNames.find((name) => !name.includes("평가") && !name.includes("입력")) ??
    workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet?.["!ref"]) {
    return {
      sheetName,
      tab: migrateLegacyTab({
        name: fallbackName,
        columns: ["지역", "지점"],
        rows: []
      })
    };
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: ""
  });

  const columns = Array.isArray(rows[0]) ? rows[0].map((value) => String(value ?? "").trim()) : [];
  const dataRows = rows
    .slice(1)
    .filter((row) => Array.isArray(row) && row.some((cell) => String(cell ?? "").trim() !== ""));

  return {
    sheetName,
    tab: migrateLegacyTab({
      name: fallbackName,
      columns,
      rows: dataRows
    })
  };
}

function extractCollabRowsFromWorkbook(arrayBuffer, fallbackName = "협업이벤트") {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellStyles: false
  });

  const sheetName =
    workbook.SheetNames.find((name) => name.includes(fallbackName)) ??
    workbook.SheetNames.find((name) => !name.includes("평가") && !name.includes("입력")) ??
    workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet?.["!ref"]) {
    return {
      sheetName,
      tab: createSpecialCollabTab(createId("tab-collab"), fallbackName, [], defaultCollabColumns)
    };
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: ""
  });

  const collabColumns = normalizeCollabColumns(Array.isArray(rows[0]) ? rows[0] : []);
  const dataRows = rows
    .slice(1)
    .filter((row) => Array.isArray(row) && row.some((cell) => String(cell ?? "").trim() !== ""));

  return {
    sheetName,
    tab: createSpecialCollabTab(
      createId("tab-collab"),
      fallbackName,
      dataRows.map((row) => {
        const seed = {};
        collabColumns.forEach((column, index) => {
          seed[column] = String(row[index] ?? "").trim();
        });
        return seed;
      }),
      collabColumns
    )
  };
}

function extractFacilityRowsFromWorkbook(arrayBuffer, fallbackName = "지점시설영상") {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellStyles: false
  });

  const sheetName =
    workbook.SheetNames.find((name) => name.includes(fallbackName)) ??
    workbook.SheetNames.find((name) => !name.includes("평가") && !name.includes("입력")) ??
    workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet?.["!ref"]) {
    return {
      sheetName,
      tab: createSpecialFacilityTab(createId("tab-facility"), fallbackName, [])
    };
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: false,
    defval: ""
  });

  const dataRows = rows
    .slice(1)
    .filter((row) => Array.isArray(row) && row.some((cell) => String(cell ?? "").trim() !== ""));

  return {
    sheetName,
    tab: createSpecialFacilityTab(
      createId("tab-facility"),
      fallbackName,
      dataRows.map((row) => ({
        region: String(row[0] ?? "").trim(),
        branch: String(row[1] ?? "").trim(),
        url: String(row[2] ?? "").trim()
      }))
    )
  };
}

function extractMentorRowsFromWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: true,
    cellNF: false,
    cellStyles: false
  });

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  if (!sheet?.["!ref"]) {
    return { sheetName, rows: [] };
  }

  const rawRows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const rows = [];

  rawRows.forEach((row, index) => {
    let year = "";
    let name = "";
    let phone = "";
    let university = "";
    let department = "";
    let branch = "";
    let group = "";
    let amount = 0;
    let memo = "";
    let isMentor = false;

    Object.entries(row).forEach(([key, val]) => {
      const keyVal = String(key).trim();
      const strVal = String(val ?? "").trim();

      if (keyVal === "연도" || keyVal.includes("year")) {
        year = strVal;
      } else if (keyVal === "이름" || keyVal === "성명" || keyVal.includes("name")) {
        name = strVal;
      } else if (keyVal === "번호" || keyVal.includes("연락처") || keyVal.includes("phone")) {
        phone = strVal;
      } else if (keyVal === "합격 대학" || keyVal === "합격대학" || keyVal.includes("university")) {
        university = strVal;
      } else if (keyVal === "학과" || keyVal.includes("department")) {
        department = strVal;
      } else if (keyVal === "지점" || keyVal.includes("branch")) {
        branch = strVal;
      } else if (keyVal === "1억장학금" || keyVal === "장학그룹" || keyVal.includes("group")) {
        group = strVal;
      } else if (keyVal === "1억 장학금" || keyVal === "장학금액" || keyVal === "장학금" || keyVal.includes("amount")) {
        const cleanedVal = strVal.replace(/[^0-9.-]+/g, "");
        amount = Number(cleanedVal) || 0;
      } else if (keyVal === "비고" || keyVal === "메모" || keyVal.includes("memo")) {
        memo = strVal;
      } else if (keyVal.includes("멘토") || keyVal.includes("isMentor")) {
        isMentor = strVal.includes("O") || strVal.includes("o") || strVal.includes("true") || strVal.includes("참") || strVal.includes("예") || strVal.includes("1");
      }
    });

    if (name || branch || university) {
      rows.push(createSpecialMentorRow({
        id: createId(`mentor-row-${index}`),
        year,
        name,
        phone,
        university,
        department,
        branch,
        group,
        amount,
        isMentor,
        memo
      }));
    }
  });

  const sortedRows = [...rows].sort((a, b) => {
    const yA = parseInt(a.year, 10) || 0;
    const yB = parseInt(b.year, 10) || 0;
    if (yB !== yA) return yB - yA;
    return (a.name || "").localeCompare(b.name || "", "ko");
  });

  return { sheetName, rows: sortedRows };
}

function ExternalScoreLink({ href, value }) {
  if (!href) {
    return <strong>{value}</strong>;
  }

  return (
    <a
      className="inline-score-link"
      href={href}
      target="_blank"
      rel="noreferrer"
      title="새 탭에서 열기"
    >
      {value}
    </a>
  );
}

function formatStatusTimestamp(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return date.toLocaleTimeString("ko-KR", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });
}

function SnsMetricEditorModal({ isOpen, onClose, branchName, socialRow, onSave, onAutoSyncBlog, isSyncing }) {
  const [formData, setFormData] = useState({});

  useEffect(() => {
    if (socialRow) {
      setFormData({
        blogUrl: socialRow.blogUrl || "",
        instagramUrl: socialRow.instagramUrl || "",
        blogRecentPosts: socialRow.blogRecentPosts ?? 0,
        blogLastPosted: socialRow.blogLastPosted || "",
        blogVisitScore: socialRow.blogVisitScore ?? 0,
        instagramRecentPosts: socialRow.instagramRecentPosts ?? 0,
        instagramLastPosted: socialRow.instagramLastPosted || "",
        instagramDesignScore: socialRow.instagramDesignScore ?? 0,
        instagramReactionScore: socialRow.instagramReactionScore ?? 0,
        profileSetupScore: socialRow.profileSetupScore ?? 0,
        featureUsageScore: socialRow.featureUsageScore ?? 0,
        ctaScore: socialRow.ctaScore ?? 0,
        linkHealthScore: socialRow.linkHealthScore ?? 0,
        brandInfoScore: socialRow.brandInfoScore ?? 0,
        memo: socialRow.memo || ""
      });
    }
  }, [socialRow, isOpen]);

  if (!isOpen) return null;

  const handleChange = (key, val) => {
    setFormData(prev => ({ ...prev, [key]: val }));
  };

  const handleSave = () => {
    onSave(branchName, formData);
    onClose();
  };

  return (
    <div className="video-modal-backdrop" onClick={onClose} style={{ zIndex: 10000 }}>
      <div 
        className="video-modal-container" 
        onClick={(e) => e.stopPropagation()} 
        style={{ maxWidth: "780px", maxHeight: "90vh", overflowY: "auto", borderRadius: "20px", background: "#ffffff", padding: "28px 32px" }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", borderBottom: "1px solid #e2e8f0", paddingBottom: "16px", marginBottom: "20px" }}>
          <div>
            <span style={{ fontSize: "0.78rem", fontWeight: "800", color: "#2563eb", letterSpacing: "0.5px" }}>SNS DIAGNOSTIC METRIC EDITOR</span>
            <h3 style={{ margin: "2px 0 0 0", fontSize: "1.3rem", fontWeight: "900", color: "#0f172a" }}>
              [{branchName}] SNS 평가 지표 수정
            </h3>
          </div>
          <button 
            className="video-modal-close-btn" 
            onClick={onClose} 
            style={{ background: "#f1f5f9", border: "none", width: "32px", height: "32px", borderRadius: "50%", cursor: "pointer", fontSize: "1rem", color: "#64748b" }}
          >
            ✕
          </button>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: "24px" }}>
          
          {/* 블로그 지표 영역 */}
          <div style={{ background: "#f0fdf4", border: "1px solid #bbf7d0", borderRadius: "16px", padding: "20px" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
              <h4 style={{ margin: 0, fontSize: "1.05rem", fontWeight: "800", color: "#166534", display: "flex", alignItems: "center", gap: "6px" }}>
                <span>🟢</span> 네이버 블로그 지표
              </h4>
              <button
                onClick={() => onAutoSyncBlog(branchName, (autoStats) => {
                  setFormData(prev => ({
                    ...prev,
                    blogRecentPosts: autoStats.recent30d,
                    blogLastPosted: autoStats.lastPosted,
                    blogVisitScore: autoStats.reactionScore
                  }));
                })}
                disabled={isSyncing}
                style={{
                  background: "#03c75a",
                  color: "#ffffff",
                  border: "none",
                  padding: "6px 12px",
                  borderRadius: "8px",
                  fontSize: "0.78rem",
                  fontWeight: "800",
                  cursor: "pointer",
                  boxShadow: "0 2px 4px rgba(3, 199, 90, 0.2)"
                }}
              >
                {isSyncing ? "동기화 중..." : "🔄 블로그 자동 크롤링"}
              </button>
            </div>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "14px" }}>
              <div style={{ gridColumn: "span 2" }}>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>블로그 주소 (URL)</label>
                <input
                  type="text"
                  value={formData.blogUrl || ""}
                  onChange={(e) => handleChange("blogUrl", e.target.value)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                  placeholder="https://blog.naver.com/..."
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>최근 30일 게시물 수</label>
                <input
                  type="number"
                  value={formData.blogRecentPosts ?? 0}
                  onChange={(e) => handleChange("blogRecentPosts", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>마지막 게시일</label>
                <input
                  type="date"
                  value={formData.blogLastPosted || ""}
                  onChange={(e) => handleChange("blogLastPosted", e.target.value)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div style={{ gridColumn: "span 2" }}>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>
                  블로그 반응 수준 (최근 5개 글 댓글/공감 기준)
                </label>
                <select
                  value={formData.blogVisitScore ?? 0}
                  onChange={(e) => handleChange("blogVisitScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem", background: "#ffffff" }}
                >
                  <option value={5}>5점: 댓글 2개 이상 게시물 2개 이상 + 공감 있는 게시물 4개 이상</option>
                  <option value={4}>4점: 댓글 있는 게시물 1개 이상 + 공감 있는 게시물 4개 이상</option>
                  <option value={3}>3점: 공감 있는 게시물 3개 이상</option>
                  <option value={2}>2점: 공감 있는 게시물 1~2개</option>
                  <option value={1}>1점: 반응 수치가 있긴 하지만 매우 약함</option>
                  <option value={0}>0점: 댓글/공감 모두 없음 (또는 블로그 미운영)</option>
                </select>
              </div>
            </div>
          </div>

          {/* 인스타그램 지표 영역 */}
          <div style={{ background: "#fdf2f8", border: "1px solid #fbcfe8", borderRadius: "16px", padding: "20px" }}>
            <h4 style={{ margin: "0 0 16px 0", fontSize: "1.05rem", fontWeight: "800", color: "#9d174d", display: "flex", alignItems: "center", gap: "6px" }}>
              <span>📷</span> 인스타그램 지표 (수기 평가)
            </h4>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: "14px" }}>
              <div style={{ gridColumn: "span 2" }}>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>인스타그램 주소 (URL)</label>
                <input
                  type="text"
                  value={formData.instagramUrl || ""}
                  onChange={(e) => handleChange("instagramUrl", e.target.value)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                  placeholder="https://instagram.com/..."
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>최근 30일 게시물 수</label>
                <input
                  type="number"
                  value={formData.instagramRecentPosts ?? 0}
                  onChange={(e) => handleChange("instagramRecentPosts", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>마지막 게시일</label>
                <input
                  type="date"
                  value={formData.instagramLastPosted || ""}
                  onChange={(e) => handleChange("instagramLastPosted", e.target.value)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>디자인/썸네일 통일성 (0~5점)</label>
                <input
                  type="number"
                  min={0}
                  max={5}
                  value={formData.instagramDesignScore ?? 0}
                  onChange={(e) => handleChange("instagramDesignScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>인스타 반응수 준수 (0~5점)</label>
                <input
                  type="number"
                  min={0}
                  max={5}
                  value={formData.instagramReactionScore ?? 0}
                  onChange={(e) => handleChange("instagramReactionScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>프로필 세팅 완성도 (0~3점)</label>
                <input
                  type="number"
                  min={0}
                  max={3}
                  value={formData.profileSetupScore ?? 0}
                  onChange={(e) => handleChange("profileSetupScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>부가기능 활용도 (0~3점)</label>
                <input
                  type="number"
                  min={0}
                  max={3}
                  value={formData.featureUsageScore ?? 0}
                  onChange={(e) => handleChange("featureUsageScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>CTA 명확성 (0~3점)</label>
                <input
                  type="number"
                  min={0}
                  max={3}
                  value={formData.ctaScore ?? 0}
                  onChange={(e) => handleChange("ctaScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>링크 정상 작동 (0~3점)</label>
                <input
                  type="number"
                  min={0}
                  max={3}
                  value={formData.linkHealthScore ?? 0}
                  onChange={(e) => handleChange("linkHealthScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div style={{ gridColumn: "span 2" }}>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>브랜드/지점 정보 명확성 (0~3점)</label>
                <input
                  type="number"
                  min={0}
                  max={3}
                  value={formData.brandInfoScore ?? 0}
                  onChange={(e) => handleChange("brandInfoScore", parseInt(e.target.value) || 0)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                />
              </div>
              <div style={{ gridColumn: "span 2" }}>
                <label style={{ display: "block", fontSize: "0.78rem", fontWeight: "700", color: "#374151", marginBottom: "4px" }}>메모 / 코멘트</label>
                <input
                  type="text"
                  value={formData.memo || ""}
                  onChange={(e) => handleChange("memo", e.target.value)}
                  style={{ width: "100%", padding: "8px 12px", borderRadius: "8px", border: "1px solid #cbd5e1", fontSize: "0.85rem" }}
                  placeholder="지점 SNS 운영 관련 특이사항 기록"
                />
              </div>
            </div>
          </div>

        </div>

        <div style={{ display: "flex", justifyContent: "flex-end", gap: "10px", marginTop: "24px", paddingTop: "16px", borderTop: "1px solid #e2e8f0" }}>
          <button
            onClick={onClose}
            style={{
              padding: "10px 20px",
              background: "#f1f5f9",
              color: "#475569",
              border: "1px solid #cbd5e1",
              borderRadius: "10px",
              fontSize: "0.88rem",
              fontWeight: "700",
              cursor: "pointer"
            }}
          >
            취소
          </button>
          <button
            onClick={handleSave}
            style={{
              padding: "10px 24px",
              background: "#2563eb",
              color: "#ffffff",
              border: "none",
              borderRadius: "10px",
              fontSize: "0.88rem",
              fontWeight: "800",
              cursor: "pointer",
              boxShadow: "0 4px 6px -1px rgba(37, 99, 235, 0.2)"
            }}
          >
            저장 및 점수 최신화
          </button>
        </div>
      </div>
    </div>
  );
}

function VideoModal({ isOpen, onClose, url, branchName }) {
  if (!isOpen || !url) return null;

  const getEmbedUrl = (videoUrl) => {
    if (!videoUrl) return null;
    
    // YouTube video ID regex
    const ytReg = /^.*(youtu.be\/|v\/|u\/\w\/|embed\/|watch\?v=|\&v=)([^#\&\?]*).*/;
    const ytMatch = videoUrl.match(ytReg);
    if (ytMatch && ytMatch[2].length === 11) {
      return `https://www.youtube.com/embed/${ytMatch[2]}?autoplay=1`;
    }

    return videoUrl;
  };

  const embedUrl = getEmbedUrl(url);

  return (
    <div className="video-modal-backdrop" onClick={onClose}>
      <div className="video-modal-container" onClick={(e) => e.stopPropagation()}>
        <div className="video-modal-header">
          <h3>🎥 {branchName} 시설영상</h3>
          <button className="video-modal-close-btn" onClick={onClose} aria-label="Close modal">✕</button>
        </div>
        <div className="video-modal-body">
          <div className="video-iframe-wrapper">
            {embedUrl ? (
              <iframe
                src={embedUrl}
                title={`${branchName} 시설영상`}
                allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture"
                allowFullScreen
              ></iframe>
            ) : (
              <div style={{ color: "var(--text)", textAlign: "center", padding: "40px" }}>영상 URL 형식이 잘못되었습니다.</div>
            )}
          </div>
          <a
            className="video-modal-fallback-btn"
            href={url}
            target="_blank"
            rel="noreferrer"
          >
            <span>🔗</span> 동영상 원본 페이지 직접 가기 (새 창)
          </a>
        </div>
      </div>
    </div>
  );
}

function CountUpNumber({ value, duration = 800, decimals = 1, suffix = "" }) {
  const [displayValue, setDisplayValue] = useState(0);
  const startValRef = useRef(0);
  const targetVal = typeof value === "number" ? value : parseFloat(value) || 0;

  useEffect(() => {
    let startTimestamp = null;
    const startVal = startValRef.current;
    const diff = targetVal - startVal;

    if (diff === 0) {
      setDisplayValue(targetVal);
      return;
    }

    let animationFrameId;

    const step = (timestamp) => {
      if (!startTimestamp) startTimestamp = timestamp;
      const progress = Math.min((timestamp - startTimestamp) / duration, 1);
      const ease = 1 - Math.pow(1 - progress, 3);
      const current = startVal + diff * ease;
      setDisplayValue(current);

      if (progress < 1) {
        animationFrameId = window.requestAnimationFrame(step);
      } else {
        startValRef.current = targetVal;
        setDisplayValue(targetVal);
      }
    };

    animationFrameId = window.requestAnimationFrame(step);

    return () => {
      if (animationFrameId) window.cancelAnimationFrame(animationFrameId);
    };
  }, [targetVal, duration]);

  const formatted = decimals > 0 ? displayValue.toFixed(decimals) : Math.round(displayValue);
  return <span>{formatted}{suffix}</span>;
}

function KeywordTagCloud({ posts, activeKeyword, onSelectKeyword }) {
  const tags = useMemo(() => {
    if (!posts || posts.length === 0) return [];
    
    const stopwords = new Set([
      "안내", "모집", "안내문", "알림", "소식", "공지", "지점", "학원", "이투스", "이투스247",
      "안녕하세요", "여러분", "관련", "위한", "대한", "진행", "오픈", "신규", "실시", "운영",
      "및", "수", "등", "더", "때", "월", "일", "년", "년도", "대상", "통해", "함께", "오늘"
    ]);

    const wordCount = {};

    posts.forEach((p) => {
      const title = (p.title || "")
        .replace(/\[.*?\]|\(.*?\)|<.*?>/g, " ")
        .replace(/[^가-힣a-zA-Z0-9\s]/g, " ");

      const tokens = title.split(/\s+/).filter(w => w.length >= 2 && !stopwords.has(w));
      tokens.forEach((t) => {
        wordCount[t] = (wordCount[t] || 0) + 1;
      });
    });

    return Object.entries(wordCount)
      .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
      .slice(0, 8)
      .map(([word, count]) => ({ word, count }));
  }, [posts]);

  if (tags.length === 0) return null;

  return (
    <div style={{ marginBottom: "20px", background: "linear-gradient(135deg, #f8fafc 0%, #f1f5f9 100%)", borderRadius: "14px", padding: "12px 16px", border: "1px solid #e2e8f0" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: "10px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
          <span style={{ fontSize: "0.95rem" }}>🏷️</span>
          <span style={{ fontSize: "0.82rem", fontWeight: "800", color: "#334155" }}>핵심 포스팅 키워드 태그</span>
          <span style={{ fontSize: "0.74rem", color: "#64748b" }}> (클릭 시 해당 글 필터링)</span>
        </div>
        {activeKeyword && (
          <button
            onClick={() => onSelectKeyword("")}
            style={{ fontSize: "0.74rem", color: "#2563eb", background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: "6px", padding: "2px 8px", cursor: "pointer", fontWeight: "700" }}
          >
            ✕ 필터 해제
          </button>
        )}
      </div>

      <div style={{ display: "flex", flexWrap: "wrap", gap: "8px", marginTop: "10px" }}>
        {tags.map(({ word, count }) => {
          const isSelected = activeKeyword === word;
          return (
            <button
              key={`tag-${word}`}
              onClick={() => onSelectKeyword(isSelected ? "" : word)}
              style={{
                display: "inline-flex",
                alignItems: "center",
                gap: "4px",
                padding: "4px 12px",
                borderRadius: "20px",
                fontSize: "0.78rem",
                fontWeight: isSelected ? "800" : "600",
                cursor: "pointer",
                transition: "all 0.2s cubic-bezier(0.16, 1, 0.3, 1)",
                background: isSelected 
                  ? "linear-gradient(135deg, #2563eb, #1d4ed8)" 
                  : "#ffffff",
                color: isSelected ? "#ffffff" : "#334155",
                border: isSelected ? "1px solid #2563eb" : "1px solid #cbd5e1",
                boxShadow: isSelected ? "0 4px 10px rgba(37, 99, 235, 0.25)" : "0 1px 2px rgba(0,0,0,0.03)"
              }}
              onMouseEnter={(e) => {
                if (!isSelected) {
                  e.currentTarget.style.borderColor = "#93c5fd";
                  e.currentTarget.style.background = "#eff6ff";
                  e.currentTarget.style.color = "#2563eb";
                  e.currentTarget.style.transform = "translateY(-1px)";
                }
              }}
              onMouseLeave={(e) => {
                if (!isSelected) {
                  e.currentTarget.style.borderColor = "#cbd5e1";
                  e.currentTarget.style.background = "#ffffff";
                  e.currentTarget.style.color = "#334155";
                  e.currentTarget.style.transform = "translateY(0)";
                }
              }}
            >
              <span>#{word}</span>
              <span style={{ fontSize: "0.7rem", opacity: isSelected ? 0.9 : 0.6, fontWeight: "700" }}>({count})</span>
            </button>
          );
        })}
      </div>
    </div>
  );
}

function BranchComparisonModal({ isOpen, onClose, branchA, rawTabs, allBranches }) {
  const [branchB, setBranchB] = useState(() => {
    const list = allBranches.filter(b => b !== branchA);
    return list[0] || "";
  });

  if (!isOpen) return null;

  const socialTab = rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
  const rowA = socialTab?.socialRows?.find(r => r.branch.trim() === branchA);
  const rowB = socialTab?.socialRows?.find(r => r.branch.trim() === branchB);

  const statsA = rowA ? summarizeSnsRow(rowA) : null;
  const statsB = rowB ? summarizeSnsRow(rowB) : null;

  const scoreA = Number(statsA?.finalScore ?? statsA?.totalScore ?? 0);
  const scoreB = Number(statsB?.finalScore ?? statsB?.totalScore ?? 0);
  const diffScore = (scoreA - scoreB).toFixed(1);

  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        backgroundColor: "rgba(15, 23, 42, 0.65)",
        backdropFilter: "blur(8px)",
        WebkitBackdropFilter: "blur(8px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 99998,
        padding: "20px",
        animation: "fadeIn 0.2s ease-out"
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        style={{
          background: "#ffffff",
          borderRadius: "24px",
          width: "100%",
          maxWidth: "840px",
          maxHeight: "90vh",
          overflowY: "auto",
          padding: "32px 30px",
          boxShadow: "0 25px 50px -12px rgba(15, 23, 42, 0.25)",
          position: "relative",
          animation: "scaleUp 0.25s cubic-bezier(0.16, 1, 0.3, 1)"
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Header */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "24px", borderBottom: "1px solid #e2e8f0", paddingBottom: "16px" }}>
          <div>
            <span style={{ fontSize: "0.78rem", fontWeight: "800", color: "#2563eb", letterSpacing: "0.5px" }}>SNS 1:1 BENCHMARK COMPARISON</span>
            <h2 style={{ margin: "4px 0 0 0", fontSize: "1.4rem", fontWeight: "900", color: "#0f172a", fontFamily: "'Outfit', 'Pretendard', sans-serif" }}>
              ⚔️ 지점 간 SNS 마케팅 지표 1:1 비교 분석
            </h2>
          </div>
          <button
            onClick={onClose}
            style={{
              background: "#f1f5f9",
              border: "none",
              borderRadius: "50%",
              width: "36px",
              height: "36px",
              cursor: "pointer",
              fontSize: "1.1rem",
              color: "#64748b",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              transition: "all 0.2s"
            }}
            onMouseEnter={(e) => { e.currentTarget.style.background = "#e2e8f0"; }}
            onMouseLeave={(e) => { e.currentTarget.style.background = "#f1f5f9"; }}
          >
            ✕
          </button>
        </div>

        {/* Branch Selectors & Head-to-Head Cards */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr auto 1fr", gap: "16px", alignItems: "center", marginBottom: "28px" }}>
          {/* Branch A Card */}
          <div style={{ background: "linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)", border: "2px solid #3b82f6", borderRadius: "18px", padding: "20px", textAlign: "center" }}>
            <span style={{ fontSize: "0.78rem", fontWeight: "800", color: "#1d4ed8" }}>기준 지점 (A)</span>
            <h3 style={{ margin: "4px 0 8px 0", fontSize: "1.3rem", fontWeight: "900", color: "#0f172a" }}>{branchA}점</h3>
            <div style={{ fontSize: "2rem", fontWeight: "900", color: "#2563eb" }}>
              <CountUpNumber value={scoreA} decimals={1} suffix="점" />
            </div>
            <span style={{ display: "inline-block", marginTop: "4px", padding: "3px 10px", borderRadius: "9999px", background: "#2563eb", color: "#ffffff", fontSize: "0.75rem", fontWeight: "800" }}>
              {statsA?.grade || "-"}등급
            </span>
          </div>

          {/* VS Badge */}
          <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: "6px" }}>
            <span style={{ fontSize: "1.3rem", fontWeight: "900", color: "#94a3b8", fontStyle: "italic", background: "#f8fafc", padding: "8px 12px", borderRadius: "50%", border: "1px solid #e2e8f0" }}>
              VS
            </span>
            <span style={{ fontSize: "0.72rem", fontWeight: "700", color: Number(diffScore) >= 0 ? "#16a34a" : "#dc2626" }}>
              {Number(diffScore) > 0 ? `${branchA} +${diffScore}점` : Number(diffScore) < 0 ? `${branchB} +${Math.abs(diffScore)}점` : "동점"}
            </span>
          </div>

          {/* Branch B Card */}
          <div style={{ background: "linear-gradient(135deg, #fdf2f8 0%, #fce7f3 100%)", border: "2px solid #ec4899", borderRadius: "18px", padding: "20px", textAlign: "center" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
              <span style={{ fontSize: "0.78rem", fontWeight: "800", color: "#be185d" }}>비교 대상 (B):</span>
              <select
                value={branchB}
                onChange={(e) => setBranchB(e.target.value)}
                style={{
                  fontSize: "0.85rem",
                  fontWeight: "800",
                  color: "#0f172a",
                  background: "#ffffff",
                  border: "1px solid #f472b6",
                  borderRadius: "8px",
                  padding: "2px 8px",
                  cursor: "pointer"
                }}
              >
                {allBranches.filter(b => b !== branchA).map(b => (
                  <option key={`compare-opt-${b}`} value={b}>{b}점</option>
                ))}
              </select>
            </div>
            <h3 style={{ margin: "4px 0 8px 0", fontSize: "1.3rem", fontWeight: "900", color: "#0f172a" }}>{branchB}점</h3>
            <div style={{ fontSize: "2rem", fontWeight: "900", color: "#db2777" }}>
              <CountUpNumber value={scoreB} decimals={1} suffix="점" />
            </div>
            <span style={{ display: "inline-block", marginTop: "4px", padding: "3px 10px", borderRadius: "9999px", background: "#db2777", color: "#ffffff", fontSize: "0.75rem", fontWeight: "800" }}>
              {statsB?.grade || "-"}등급
            </span>
          </div>
        </div>

        {/* Detailed Metrics Comparison Table */}
        <div style={{ border: "1px solid #e2e8f0", borderRadius: "16px", overflow: "hidden" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", textAlign: "center", fontSize: "0.86rem" }}>
            <thead>
              <tr style={{ background: "#f8fafc", borderBottom: "1px solid #e2e8f0", color: "#64748b", fontSize: "0.8rem", fontWeight: "700" }}>
                <th style={{ padding: "12px 16px", textAlign: "left" }}>비교 지표 항목</th>
                <th style={{ padding: "12px 16px", color: "#2563eb", width: "180px" }}>{branchA}점</th>
                <th style={{ padding: "12px 16px", color: "#db2777", width: "180px" }}>{branchB}점</th>
                <th style={{ padding: "12px 16px", width: "120px" }}>격차</th>
              </tr>
            </thead>
            <tbody>
              {[
                { label: "🟢 네이버 블로그 총점 (50점)", a: statsA?.blogScore || 0, b: statsB?.blogScore || 0, max: 50 },
                { label: "📝 블로그 최근 30일 발행 글 수", a: `${statsA?.blogRecentPosts || 0}개`, b: `${statsB?.blogRecentPosts || 0}개`, rawA: Number(statsA?.blogRecentPosts || 0), rawB: Number(statsB?.blogRecentPosts || 0) },
                { label: "💬 블로그 최근 반응 수준 (5점)", a: `${statsA?.blogVisitScore || 0}점`, b: `${statsB?.blogVisitScore || 0}점`, rawA: Number(statsA?.blogVisitScore || 0), rawB: Number(statsB?.blogVisitScore || 0) },
                { label: "📅 블로그 마지막 포스팅일", a: statsA?.blogLastPosted || "기록없음", b: statsB?.blogLastPosted || "기록없음", noDiff: true },
                { label: "🌸 인스타그램 총점 (50점)", a: statsA?.instagramScore || 0, b: statsB?.instagramScore || 0, max: 50 },
                { label: "🎨 인스타그램 디자인 수준 (5점)", a: `${statsA?.instagramDesignScore || 0}점`, b: `${statsB?.instagramDesignScore || 0}점`, rawA: Number(statsA?.instagramDesignScore || 0), rawB: Number(statsB?.instagramDesignScore || 0) },
                { label: "❤️ 인스타그램 반응 수준 (5점)", a: `${statsA?.instagramReactionScore || 0}점`, b: `${statsB?.instagramReactionScore || 0}점`, rawA: Number(statsA?.instagramReactionScore || 0), rawB: Number(statsB?.instagramReactionScore || 0) }
              ].map((row, rIdx) => {
                const numA = row.rawA !== undefined ? row.rawA : Number(row.a || 0);
                const numB = row.rawB !== undefined ? row.rawB : Number(row.b || 0);
                const diff = (numA - numB).toFixed(1);
                const isWinnerA = numA > numB;
                const isWinnerB = numB > numA;

                return (
                  <tr key={`cmp-row-${rIdx}`} style={{ borderBottom: rIdx < 6 ? "1px solid #f1f5f9" : "none", background: rIdx % 2 === 1 ? "#fafafa" : "#ffffff" }}>
                    <td style={{ padding: "12px 16px", textAlign: "left", fontWeight: "700", color: "#334155" }}>
                      {row.label}
                    </td>
                    <td style={{ padding: "12px 16px", fontWeight: isWinnerA ? "800" : "600", color: isWinnerA ? "#2563eb" : "#475569" }}>
                      {row.a} {isWinnerA && "👑"}
                    </td>
                    <td style={{ padding: "12px 16px", fontWeight: isWinnerB ? "800" : "600", color: isWinnerB ? "#db2777" : "#475569" }}>
                      {row.b} {isWinnerB && "👑"}
                    </td>
                    <td style={{ padding: "12px 16px", fontSize: "0.8rem", fontWeight: "700" }}>
                      {row.noDiff ? (
                        <span style={{ color: "#94a3b8" }}>-</span>
                      ) : Number(diff) > 0 ? (
                        <span style={{ color: "#16a34a", fontWeight: "800" }}>{branchA} +{diff}</span>
                      ) : Number(diff) < 0 ? (
                        <span style={{ color: "#dc2626", fontWeight: "800" }}>{branchB} +{Math.abs(diff)}</span>
                      ) : (
                        <span style={{ color: "#94a3b8" }}>동일</span>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Footer Close */}
        <div style={{ marginTop: "24px", display: "flex", justifyContent: "flex-end" }}>
          <button
            onClick={onClose}
            style={{
              padding: "10px 22px",
              background: "#2563eb",
              color: "#ffffff",
              border: "none",
              borderRadius: "12px",
              fontWeight: "800",
              fontSize: "0.88rem",
              cursor: "pointer",
              boxShadow: "0 4px 12px rgba(37, 99, 235, 0.25)"
            }}
          >
            닫기
          </button>
        </div>
      </div>
    </div>
  );
}

function CustomDialogModal({ modal, onClose }) {
  if (!modal) return null;

  const {
    type = "info", // "info" | "success" | "warning" | "error" | "confirm"
    title = "알림",
    message = "",
    confirmText = "확인",
    cancelText = "취소",
    onConfirm,
    onCancel
  } = modal;

  const getIconAndColor = () => {
    switch (type) {
      case "success":
        return { icon: "🎉", bg: "#ecfdf5", border: "#a7f3d0", color: "#059669" };
      case "confirm":
        return { icon: "🔄", bg: "#eff6ff", border: "#bfdbfe", color: "#2563eb" };
      case "warning":
        return { icon: "⚠️", bg: "#fffbeb", border: "#fde68a", color: "#d97706" };
      case "error":
        return { icon: "🚨", bg: "#fef2f2", border: "#fecaca", color: "#dc2626" };
      default:
        return { icon: "ℹ️", bg: "#f0f9ff", border: "#bae6fd", color: "#0284c7" };
    }
  };

  const theme = getIconAndColor();

  return (
    <div
      style={{
        position: "fixed",
        inset: 0,
        backgroundColor: "rgba(15, 23, 42, 0.55)",
        backdropFilter: "blur(6px)",
        WebkitBackdropFilter: "blur(6px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        zIndex: 99999,
        padding: "20px",
        animation: "fadeIn 0.2s ease-out"
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget && type !== "confirm") {
          if (onConfirm) onConfirm();
          else onClose();
        }
      }}
    >
      <div
        style={{
          background: "linear-gradient(145deg, #ffffff 0%, #f8fafc 100%)",
          borderRadius: "24px",
          width: "100%",
          maxWidth: "420px",
          padding: "28px 26px 24px",
          boxShadow: "0 25px 50px -12px rgba(15, 23, 42, 0.25), 0 0 0 1px rgba(226, 232, 240, 0.8)",
          position: "relative",
          animation: "scaleUp 0.25s cubic-bezier(0.16, 1, 0.3, 1)",
          textAlign: "center"
        }}
        onClick={(e) => e.stopPropagation()}
      >
        {/* Icon */}
        <div
          style={{
            width: "56px",
            height: "56px",
            borderRadius: "18px",
            background: theme.bg,
            border: `1.5px solid ${theme.border}`,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            fontSize: "1.7rem",
            margin: "0 auto 16px",
            boxShadow: `0 8px 16px -4px ${theme.border}`
          }}
        >
          {theme.icon}
        </div>

        {/* Title */}
        <h3
          style={{
            margin: "0 0 10px 0",
            fontSize: "1.25rem",
            fontWeight: "800",
            color: "#0f172a",
            fontFamily: "'Outfit', 'Pretendard', sans-serif",
            letterSpacing: "-0.3px"
          }}
        >
          {title}
        </h3>

        {/* Message */}
        <div
          style={{
            fontSize: "0.92rem",
            color: "#475569",
            lineHeight: "1.65",
            marginBottom: "24px",
            whiteSpace: "pre-line",
            wordBreak: "keep-all",
            fontWeight: "500",
            padding: "0 8px"
          }}
        >
          {message}
        </div>

        {/* Action Buttons */}
        <div style={{ display: "flex", gap: "10px", justifyContent: "center" }}>
          {type === "confirm" && (
            <button
              onClick={() => {
                if (onCancel) onCancel();
                else onClose();
              }}
              style={{
                flex: 1,
                padding: "12px 18px",
                background: "#f1f5f9",
                border: "1px solid #cbd5e1",
                borderRadius: "14px",
                color: "#475569",
                fontSize: "0.92rem",
                fontWeight: "700",
                cursor: "pointer",
                transition: "all 0.2s"
              }}
              onMouseEnter={(e) => { e.currentTarget.style.background = "#e2e8f0"; }}
              onMouseLeave={(e) => { e.currentTarget.style.background = "#f1f5f9"; }}
            >
              {cancelText}
            </button>
          )}
          <button
            onClick={() => {
              if (onConfirm) onConfirm();
              else onClose();
            }}
            style={{
              flex: type === "confirm" ? 1 : "initial",
              minWidth: type === "confirm" ? "auto" : "140px",
              padding: "12px 24px",
              background: type === "success" 
                ? "linear-gradient(135deg, #10b981 0%, #059669 100%)" 
                : "linear-gradient(135deg, #003bff 0%, #1d4ed8 100%)",
              border: "none",
              borderRadius: "14px",
              color: "#ffffff",
              fontSize: "0.92rem",
              fontWeight: "800",
              cursor: "pointer",
              boxShadow: type === "success" 
                ? "0 4px 14px rgba(16, 185, 129, 0.35)" 
                : "0 4px 14px rgba(0, 59, 255, 0.35)",
              transition: "all 0.2s"
            }}
            onMouseEnter={(e) => { e.currentTarget.style.transform = "translateY(-1px)"; }}
            onMouseLeave={(e) => { e.currentTarget.style.transform = "translateY(0)"; }}
          >
            {confirmText}
          </button>
        </div>
      </div>
    </div>
  );
}

export default function HomePage() {
  const [page, setPage] = useState("dashboard");
  const [rawTabs, setRawTabs] = useState(initialTabs);
  const rawTabsRef = useRef(rawTabs);
  useEffect(() => {
    rawTabsRef.current = rawTabs;
  }, [rawTabs]);

  const [activeSlideIndex, setActiveSlideIndex] = useState(0);
  const [customModal, setCustomModal] = useState(null);
  const [isCompareModalOpen, setIsCompareModalOpen] = useState(false);
  const [selectedKeywordFilter, setSelectedKeywordFilter] = useState("");
  const [isSnsScoreTooltipOpen, setIsSnsScoreTooltipOpen] = useState(false);

  const showAlert = (message, title = "알림", type = "info") => {
    return new Promise((resolve) => {
      setCustomModal({
        type,
        title,
        message,
        onConfirm: () => {
          setCustomModal(null);
          resolve(true);
        }
      });
    });
  };

  const showConfirm = (message, title = "확인", confirmText = "확인", cancelText = "취소") => {
    return new Promise((resolve) => {
      setCustomModal({
        type: "confirm",
        title,
        message,
        confirmText,
        cancelText,
        onConfirm: () => {
          setCustomModal(null);
          resolve(true);
        },
        onCancel: () => {
          setCustomModal(null);
          resolve(false);
        }
      });
    });
  };

  const showcaseSlides = useMemo(() => [
    {
      name: "체험단",
      className: "card-theme-experience",
      desc: "전국 지점의 블로그 체험단 모집 현황 및 지점 마케팅 성과를 투명하게 분석하고 관리하는 시스템",
      url: "https://etoos247-experience-info.vercel.app/",
      imgSrc: "/showcase-experience.png"
    },
    {
      name: "합격자 취합",
      className: "card-theme-pass",
      desc: "전국 지점에서 배출된 이투스247학원의 합격생 데이터를 실시간으로 취합하고 편리하게 증빙하는 공간",
      url: "https://admit-collector.vercel.app/",
      imgSrc: "/showcase-pass.png"
    },
    {
      name: "언론보도",
      className: "card-theme-news",
      desc: "이투스247학원의 주요 입시 전략, 공식 보도자료 및 이벤트 소식을 실시간으로 관리하고 확인하는 아카이브",
      url: "https://intelligent-salk.vercel.app",
      imgSrc: "/showcase-news.png"
    },
    {
      name: "YOUTUBE",
      className: "card-theme-youtube",
      desc: "합격 성공 스토리, 입시 정보 및 이투스247학원의 생생한 현장 소식을 전하는 본사 공식 유튜브 채널",
      url: "https://www.youtube.com/@etoos247",
      imgSrc: "/showcase-youtube.png"
    },
    {
      name: "INSTAGRAM",
      className: "card-theme-instagram",
      desc: "이투스247학원의 생생한 학원 일상과 최신 마케팅 소식, 이벤트를 공유하는 본사 공식 인스타그램",
      url: "https://www.instagram.com/etoos247_official/",
      imgSrc: "/showcase-instagram.png"
    }
  ], []);

  const marqueeCards = useMemo(() => [
    { name: "247프렌즈", category: "CAMPAIGN", className: "card-theme-friends", id: rawTabs.find(t => t.name === "247프렌즈")?.id },
    { name: "247체험단", category: "MARKETING", className: "card-theme-experience", id: rawTabs.find(t => t.name === "247체험단")?.id },
    { name: "SNS 진단표", category: "ANALYSIS", className: "card-theme-sns", id: rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND)?.id },
    { name: "협업이벤트", category: "COLLABORATION", className: "card-theme-collab", id: rawTabs.find(t => t.kind === SPECIAL_COLLAB_TAB_KIND)?.id },
    { name: "지점시설영상", category: "PROMOTION", className: "card-theme-facility", id: rawTabs.find(t => t.kind === SPECIAL_FACILITY_TAB_KIND)?.id },
    { name: "합격자 취합", category: "RESULTS", className: "card-theme-pass", id: rawTabs.find(t => t.name === "합격자 취합")?.id },
    { name: "언론보도", category: "NEWS", className: "card-theme-news", id: rawTabs.find(t => t.name === "언론보도")?.id },
    { name: "멘토단 및 장학생", category: "SCHOLARSHIP", className: "card-theme-mentor", id: rawTabs.find(t => t.kind === SPECIAL_MENTOR_TAB_KIND)?.id }
  ], [rawTabs]);

  useEffect(() => {
    if (page === "dashboard") {
      const timer = setInterval(() => {
        setActiveSlideIndex((prev) => (prev + 1) % showcaseSlides.length);
      }, 4000);
      return () => clearInterval(timer);
    }
  }, [page, showcaseSlides.length]);

  const prevSlide = (e) => {
    e.stopPropagation();
    setActiveSlideIndex((prev) => (prev - 1 + showcaseSlides.length) % showcaseSlides.length);
  };
  const nextSlide = (e) => {
    e.stopPropagation();
    setActiveSlideIndex((prev) => (prev + 1) % showcaseSlides.length);
  };

  const handleCardClick = (id) => {
    if (id) {
      sortMentorRowsState();
      setDashboardTabId(id);
      setActiveTabId(id);
      document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
    }
  };

  const handleNavMenuClick = (menuName) => {
    if (menuName === "dashboard") {
      sortMentorRowsState();
      setDashboardTabId(OVERVIEW_TAB_ID);
      setPage("dashboard");
      setTimeout(() => {
        document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
      }, 100);
      return;
    }

    if (menuName === "competitors") {
      sortMentorRowsState();
      setPage("competitors");
      return;
    }

    if (menuName === "sns") {
      sortMentorRowsState();
      setPage("sns");
      return;
    }

    if (menuName === "rawdata") {
      sortMentorRowsState();
      setPage("rawdata");
      return;
    }

    if (menuName === "전체 현황") {
      sortMentorRowsState();
      setDashboardTabId(OVERVIEW_TAB_ID);
      setPage("dashboard");
      setTimeout(() => {
        document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
      }, 100);
    } else if (menuName === "경쟁사 비교") {
      sortMentorRowsState();
      setPage("competitors");
    } else if (menuName === "SNS 분석") {
      sortMentorRowsState();
      setPage("sns");
    } else if (menuName === "RAWDATASTUDIO") {
      sortMentorRowsState();
      setPage("rawdata");
    } else if (menuName === "지점 대시보드") {
      document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
    } else if (menuName === "명예의 전당") {
      const tab = rawTabs.find(t => t.kind === SPECIAL_MENTOR_TAB_KIND);
      if (tab) {
        sortMentorRowsState();
        setDashboardTabId(tab.id);
        setActiveTabId(tab.id);
        document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
      }
    } else if (menuName === "체험단 신청") {
      const tab = rawTabs.find(t => t.name === "247체험단");
      if (tab) {
        sortMentorRowsState();
        setDashboardTabId(tab.id);
        setActiveTabId(tab.id);
        document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
      }
    } else {
      alert(`${menuName} 메뉴 준비 중입니다.`);
    }
  };

  const globalStats = useMemo(() => {
    // 1. 247프렌즈
    const friendsTab = rawTabs.find(t => t.name === "247프렌즈");
    let friendsBranchCount = 0;
    let friendsAvgParticipants = 0;
    if (friendsTab) {
      const rows = friendsTab.rows || [];
      friendsBranchCount = rows.filter(r => r.branch?.trim()).length;
      let totalParticipants = 0;
      rows.forEach(r => {
        Object.values(r.eventValues || {}).forEach(v => {
          totalParticipants += Number(v.participants) || 0;
        });
      });
      friendsAvgParticipants = friendsBranchCount > 0 ? Math.round(totalParticipants / friendsBranchCount) : 0;
    }

    // 2. 247체험단
    const experienceTab = rawTabs.find(t => t.name === "247체험단");
    let experienceBranchCount = 0;
    let experienceTotalParticipants = 0;
    if (experienceTab) {
      const rows = experienceTab.rows || [];
      experienceBranchCount = rows.filter(r => r.branch?.trim()).length;
      rows.forEach(r => {
        Object.values(r.eventValues || {}).forEach(v => {
          experienceTotalParticipants += Number(v.participants) || 0;
        });
      });
    }

    // 3. SNS 마케팅
    const snsTab = rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
    let snsTotalCount = 0;
    let snsAvgScore = "0.0";
    if (snsTab) {
      const rows = snsTab.socialRows || [];
      snsTotalCount = rows.filter(r => r.branch?.trim()).length;
      let totalScores = 0;
      rows.forEach(r => {
        const scores = [
          Number(r.blogVisitScore) || 0,
          Number(r.instagramDesignScore) || 0,
          Number(r.instagramReactionScore) || 0,
          Number(r.profileSetupScore) || 0,
          Number(r.featureUsageScore) || 0,
          Number(r.ctaScore) || 0,
          Number(r.linkHealthScore) || 0,
          Number(r.brandInfoScore) || 0
        ];
        scores.forEach(s => {
          totalScores += s;
        });
      });
      snsAvgScore = snsTotalCount > 0 ? (totalScores / (snsTotalCount * 8)).toFixed(1) : "0.0";
    }

    // 4. 협업 성과
    const collabTab = rawTabs.find(t => t.kind === SPECIAL_COLLAB_TAB_KIND);
    let collabEventCount = 0;
    let collabTotalUrls = 0;
    if (collabTab) {
      const summary = buildCollabSummary(collabTab);
      collabEventCount = summary.uniqueEvents;
      collabTotalUrls = summary.totalUrls;
    }

    // 5. 시설 홍보
    const facilityTab = rawTabs.find(t => t.kind === SPECIAL_FACILITY_TAB_KIND);
    let facilityRatio = 0;
    if (facilityTab) {
      const summary = buildFacilitySummary(facilityTab);
      facilityRatio = summary.totalBranches > 0 ? Math.round((summary.activeBranches / summary.totalBranches) * 100) : 0;
    }

    // 6. 멘토단/장학생
    const mentorTab = rawTabs.find(t => t.kind === SPECIAL_MENTOR_TAB_KIND);
    let mentorTotalCount = 0;
    let mentorTotalAmount = 0;
    if (mentorTab) {
      const rows = mentorTab.mentorRows || [];
      mentorTotalCount = rows.filter(r => r.name?.trim()).length;
      mentorTotalAmount = rows.reduce((sum, r) => sum + (Number(r.amount) || 0), 0);
    }

    return {
      friendsBranchCount,
      friendsAvgParticipants,
      experienceBranchCount,
      experienceTotalParticipants,
      snsTotalCount,
      snsAvgScore,
      collabEventCount,
      collabTotalUrls,
      facilityRatio,
      mentorTotalCount,
      mentorTotalAmount
    };
  }, [rawTabs]);

  const sortMentorRowsState = () => {
    setRawTabs((current) => getSortedRawTabs(current));
  };
  const [activeTabId, setActiveTabId] = useState(initialTabs[0].id);
  const [dashboardTabId, setDashboardTabId] = useState(OVERVIEW_TAB_ID);
  const [saveState, setSaveState] = useState("서버 저장 대기 중");
  const [isHydrated, setIsHydrated] = useState(false);
  const [selectedBranch, setSelectedBranch] = useState(null);
  const [mapStatus, setMapStatus] = useState({ provider: "", message: "" });
  const [isCrawling, setIsCrawling] = useState(false);
  const [crawlingLogs, setCrawlingLogs] = useState([]);
  const [crawledOwnPromotions, setCrawledOwnPromotions] = useState({});
  const [crawledCompPromotions, setCrawledCompPromotions] = useState({});
  const [aiAnalyzedData, setAiAnalyzedData] = useState({});
  const [aiLoading, setAiLoading] = useState({});
  const [ownPromoIndex, setOwnPromoIndex] = useState(0);
  const [compPromoIndex, setCompPromoIndex] = useState(0);
  const [ownPromoHovered, setOwnPromoHovered] = useState(false);
  const [compPromoHovered, setCompPromoHovered] = useState(false);
  const [crawledTimestamps, setCrawledTimestamps] = useState({});
  const [aiStrategies, setAiStrategies] = useState({});
  const [isAiGenerating, setIsAiGenerating] = useState(false);
  const [typedAiStrategy, setTypedAiStrategy] = useState("");
  const [trendData, setTrendData] = useState([]);
  const [demographics, setDemographics] = useState({ pc: 50, mobile: 50, parent: 50, student: 50 });
  const [roiData, setRoiData] = useState({ before: 20, after: 60 });
  const [seasonalAlert, setSeasonalAlert] = useState("");
  const [searchQuery, setSearchQuery] = useState("");
  const [socialData, setSocialData] = useState({ kin: [], blog: [], news: [] });
  const [activeSocialSubTab, setActiveSocialSubTab] = useState("kin");
  const [hoveredMonthIndex, setHoveredMonthIndex] = useState(null);
  const [hoveredSliceIndex, setHoveredSliceIndex] = useState(null);
  const [trendViewMode, setTrendViewMode] = useState("cards");
  const [cardViewStyle, setCardViewStyle] = useState("rolling"); // "rolling" | "grid"
  const [selectedCompForPromo, setSelectedCompForPromo] = useState(null);
  const [isSocialLoading, setIsSocialLoading] = useState(false);
  const allBranches = useMemo(() => {
    const list = rawTabs
      .filter((tab) => !isSpecialTabKind(tab.kind))
      .flatMap((tab) => tab.rows.map((row) => row.branch.trim()))
      .filter(Boolean);
    return [...new Set(list)].sort();
  }, [rawTabs]);
  const [selectedCollabBranch, setSelectedCollabBranch] = useState(null);
  const [selectedCollabEvent, setSelectedCollabEvent] = useState(null);
  const [selectedOverviewBranch, setSelectedOverviewBranch] = useState(null);
  const [selectedChartEventId, setSelectedChartEventId] = useState(null);
  const [overviewSearch, setOverviewSearch] = useState("");
  const [snsSearch, setSnsSearch] = useState("");
  const [mentorSearch, setMentorSearch] = useState("");
  const [mentorBranchFilter, setMentorBranchFilter] = useState("all");
  const [mentorUnivFilter, setMentorUnivFilter] = useState("all");
  const [areEventChipsExpanded, setAreEventChipsExpanded] = useState(true);
  const [branchKeyword, setBranchKeyword] = useState("");
  const [areRegionsExpanded, setAreRegionsExpanded] = useState(false);
  const [branchSearch, setBranchSearch] = useState("");
  const [regionFilter, setRegionFilter] = useState("all");
  const [isSnsModalOpen, setIsSnsModalOpen] = useState(false);
  const [isSnsSyncing, setIsSnsSyncing] = useState(false);

  const handleSaveSnsMetrics = (targetBranch, updatedFields) => {
    setRawTabs(prevTabs => {
      return prevTabs.map(tab => {
        if (tab.kind !== SPECIAL_SOCIAL_TAB_KIND) return tab;
        const newRows = (tab.socialRows || []).map(r => {
          if (r.branch.trim() !== targetBranch.trim()) return r;
          return {
            ...r,
            ...updatedFields
          };
        });
        return {
          ...tab,
          socialRows: newRows
        };
      });
    });
    setTimeout(() => {
      forceServerSave();
    }, 100);
  };

  const [batchSyncProgress, setBatchSyncProgress] = useState(null);

  const handleBatchSyncAllBlogs = async () => {
    const socialTab = rawTabsRef.current.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
    if (!socialTab || !socialTab.socialRows) return;

    const targetRows = socialTab.socialRows.filter(r => r.blogUrl && !isMissingChannelUrl(r.blogUrl));
    if (targetRows.length === 0) {
      showAlert("분석 가능한 네이버 블로그 URL이 등록된 지점이 없습니다.", "URL 미등록", "warning");
      return;
    }

    const shouldProceed = await showConfirm(`전체 ${targetRows.length}개 지점의 네이버 블로그를 실시간으로 크롤링하여\n최근 30일 글 수, 마지막 등록일, 반응 수준 점수를 일괄 최신화하시겠습니까?`, "전체 지점 블로그 일괄 자동 분석");
    if (!shouldProceed) {
      return;
    }

    setIsSnsSyncing(true);
    setBatchSyncProgress({ current: 0, total: targetRows.length });

    const updatedMap = {};
    let completedCount = 0;
    const chunkSize = 6;

    for (let i = 0; i < targetRows.length; i += chunkSize) {
      const chunk = targetRows.slice(i, i + chunkSize);
      await Promise.all(
        chunk.map(async (r) => {
          try {
            const res = await fetch(`/api/crawl?blogUrl=${encodeURIComponent(r.blogUrl)}&branchName=${encodeURIComponent(r.branch)}`);
            const data = await res.json();
            if (data.success && data.blogStats) {
              updatedMap[r.branch.trim()] = {
                blogRecentPosts: data.blogStats.recent30d,
                blogLastPosted: data.blogStats.lastPosted,
                blogVisitScore: data.blogStats.reactionScore
              };
            }
          } catch (err) {
            console.error(`Batch sync error for ${r.branch}:`, err);
          } finally {
            completedCount++;
            setBatchSyncProgress({ current: completedCount, total: targetRows.length });
          }
        })
      );
    }

    // Apply all updates in a single state call and persist to server
    setRawTabs(prevTabs => {
      const nextTabs = prevTabs.map(tab => {
        if (tab.kind !== SPECIAL_SOCIAL_TAB_KIND) return tab;
        const newRows = (tab.socialRows || []).map(r => {
          const stats = updatedMap[r.branch.trim()];
          if (!stats) return r;
          return {
            ...r,
            ...stats
          };
        });
        return {
          ...tab,
          socialRows: newRows
        };
      });

      fetch("/api/rawtabs", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ rawTabs: nextTabs })
      }).catch(err => console.error("Error saving batch sync to server:", err));

      return nextTabs;
    });

    setIsSnsSyncing(false);
    setBatchSyncProgress(null);
    showAlert(`전체 ${Object.keys(updatedMap).length}개 지점의 네이버 블로그 실시간 분석 및 SNS 진단 점수 최신화가 성공적으로 완료되었습니다!`, "일괄 분석 완료", "success");
  };

  const handleAutoSyncBlog = async (targetBranch, callback) => {
    setIsSnsSyncing(true);
    try {
      const socialTab = rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
      const row = socialTab?.socialRows?.find(r => r.branch.trim() === targetBranch);
      const blogUrl = row?.blogUrl || "";
      if (!blogUrl) {
        showAlert("등록된 지점 네이버 블로그 URL이 없습니다.", "URL 미등록", "warning");
        return;
      }
      const res = await fetch(`/api/crawl?blogUrl=${encodeURIComponent(blogUrl)}&branchName=${encodeURIComponent(targetBranch)}`);
      const data = await res.json();
      if (data.success && data.blogStats) {
        const updatedStats = data.blogStats;
        handleSaveSnsMetrics(targetBranch, {
          blogRecentPosts: updatedStats.recent30d,
          blogLastPosted: updatedStats.lastPosted,
          blogVisitScore: updatedStats.reactionScore
        });
        if (callback) {
          callback(updatedStats);
        }
        showAlert(`[${targetBranch}] 지점의 블로그 지표가 성공적으로 동기화되었습니다!\n\n• 최근 30일 글: ${updatedStats.recent30d}개\n• 마지막 등록일: ${updatedStats.lastPosted}\n• 반응 수준 점수: ${updatedStats.reactionScore}점/5점`, "블로그 동기화 완료", "success");
      } else {
        showAlert("블로그 데이터를 가져오지 못했습니다. 잠시 후 다시 시도해 주세요.", "동기화 실패", "error");
      }
    } catch (err) {
      console.error(err);
      showAlert("블로그 동기화 중 오류가 발생했습니다.", "오류", "error");
    } finally {
      setIsSnsSyncing(false);
    }
  };

  const [hoveredRegion, setHoveredRegion] = useState(null);
  const [activeVideoUrl, setActiveVideoUrl] = useState(null);
  const [activeVideoBranch, setActiveVideoBranch] = useState("");
  const saveTimeoutRef = useRef(null);
  const hasInitializedSaveRef = useRef(false);
  const importInputRef = useRef(null);
  const mapRef = useRef(null);
  const wheelTimeoutRef = useRef(false);
  const ownPromoContainerRef = useRef(null);
  const compPromoContainerRef = useRef(null);

  useEffect(() => {
    const mapContainer = document.getElementById("competitor-map-leaflet");

    const cleanupMap = () => {
      if (mapRef.current?.remove) {
        mapRef.current.remove();
      }
      mapRef.current = null;
      if (mapContainer) mapContainer.innerHTML = "";
    };

    if ((page !== "competitors" && page !== "sns") || !mapContainer) {
      cleanupMap();
      return;
    }

    let isMounted = true;
    let sdkWaitTimer = null;
    const naverClientId = process.env.NEXT_PUBLIC_NAVER_CLIENT_ID;
    const isNaverMapEnabled = naverClientId && naverClientId !== "YOUR_NAVER_CLIENT_ID" && naverClientId.trim() !== "";

    const resolveNaverBranchCoords = async (branchName) => {
      const fallbackCoords = getBranchCoords(branchName);
      const address = getBranchAddress(branchName);
      const geocoder = window.naver?.maps?.Service?.geocode;

      if (!address) return fallbackCoords;

      const cacheKey = `naver-geocode:v2:${address}`;
      try {
        const cached = window.localStorage?.getItem(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.every(Number.isFinite)) return parsed;
        }
      } catch (error) {
        console.warn("Could not read cached geocode result.", error);
      }

      try {
        const response = await fetch(`/api/geocode?query=${encodeURIComponent(address)}`, {
          cache: "no-store"
        });
        const data = await response.json().catch(() => ({}));

        if (response.ok && Array.isArray(data?.coords) && data.coords.every(Number.isFinite)) {
          try {
            window.localStorage?.setItem(cacheKey, JSON.stringify(data.coords));
          } catch (error) {
            console.warn("Could not cache geocode result.", error);
          }
          return data.coords;
        }

        console.warn("Server geocode failed for branch.", branchName, address, data?.error || response.status);
      } catch (error) {
        console.warn("Server geocode request failed for branch.", branchName, address, error);
      }

      if (!geocoder) return fallbackCoords;

      return new Promise((resolve) => {
        geocoder({ query: address }, (status, response) => {
          const serviceStatus = window.naver?.maps?.Service?.Status;
          const addresses = response?.v2?.addresses || response?.result?.items || [];
          const firstAddress = addresses[0];
          const lat = Number(firstAddress?.y ?? firstAddress?.point?.y);
          const lng = Number(firstAddress?.x ?? firstAddress?.point?.x);
          const isOk = !serviceStatus || status === serviceStatus.OK;

          if (isOk && Number.isFinite(lat) && Number.isFinite(lng)) {
            const coords = [lat, lng];
            try {
              window.localStorage?.setItem(cacheKey, JSON.stringify(coords));
            } catch (error) {
              console.warn("Could not cache geocode result.", error);
            }
            resolve(coords);
            return;
          }

          console.warn("NAVER geocode failed for branch.", branchName, address, status);
          resolve(fallbackCoords);
        });
      });
    };

    const resolveAddressCoords = async (address, fallbackCoords, cacheKeySuffix) => {
      if (!address) return fallbackCoords;

      const cacheKey = `naver-geocode:v2:${cacheKeySuffix || address}`;
      try {
        const cached = window.localStorage?.getItem(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached);
          if (Array.isArray(parsed) && parsed.every(Number.isFinite)) return parsed;
        }
      } catch (error) {
        console.warn("Could not read cached geocode result.", error);
      }

      try {
        const response = await fetch(`/api/geocode?query=${encodeURIComponent(address)}`, {
          cache: "no-store"
        });
        const data = await response.json().catch(() => ({}));

        if (response.ok && Array.isArray(data?.coords) && data.coords.every(Number.isFinite)) {
          try {
            window.localStorage?.setItem(cacheKey, JSON.stringify(data.coords));
          } catch (error) {
            console.warn("Could not cache geocode result.", error);
          }
          return data.coords;
        }

        console.warn("Address geocode failed.", address, data?.error || response.status);
      } catch (error) {
        console.warn("Address geocode request failed.", address, error);
      }

      return fallbackCoords;
    };

    const getSelectedBranchMapItems = async () => {
      const branchCoords = await resolveNaverBranchCoords(selectedBranch);
      if (page === "sns") {
        return [
          {
            type: "branch",
            name: selectedBranch,
            label: `이투스247학원 ${selectedBranch}`,
            coords: branchCoords,
            isActive: true
          }
        ];
      }
      const competitors = getBranchCompetitorAddresses(selectedBranch);
      const items = [
        {
          type: "branch",
          name: selectedBranch,
          label: `이투스247학원 ${selectedBranch}`,
          coords: branchCoords,
          isActive: true
        }
      ];

      for (const [index, competitor] of competitors.entries()) {
        const fallbackCoords = getNearbyCoords(branchCoords, index);
        const address = competitor.address || "";
        const coords = await resolveAddressCoords(
          address,
          fallbackCoords,
          `competitor:${selectedBranch}:${competitor.name}:${address || "fallback"}`
        );
        items.push({
          type: "competitor",
          name: competitor.name,
          label: competitor.name,
          coords,
          isActive: false
        });
      }

      return items;
    };

    const initNaverMap = async () => {
      if (!isMounted || !window.naver?.maps) return false;
      cleanupMap();

      try {
        const initialCoords = selectedBranch ? await resolveNaverBranchCoords(selectedBranch) : [36.2, 127.8];
        const initialZoom = selectedBranch ? 16 : 7;
        const map = new window.naver.maps.Map("competitor-map-leaflet", {
          center: new window.naver.maps.LatLng(initialCoords[0], initialCoords[1]),
          zoom: initialZoom,
          zoomControl: true
        });

        mapRef.current = map;

        const rawMapItems = selectedBranch
          ? await getSelectedBranchMapItems()
          : allBranches.map((branchName) => ({
              type: "branch",
              name: branchName,
              label: `이투스247학원 ${branchName}`,
              coords: getBranchCoords(branchName),
              isActive: false
            }));
        const mapItems = getDisplayMapItems(rawMapItems);
        const markerPositions = [];
        for (const item of mapItems) {
          if (!isMounted) return false;
          const coords = item.displayCoords || item.coords;
          const position = new window.naver.maps.LatLng(coords[0], coords[1]);
          markerPositions.push(position);
          const marker = new window.naver.maps.Marker({
            position,
            map,
            icon: {
              content: getBranchMarkerHtml(item.label, item.isActive, "", item.type),
              anchor: new window.naver.maps.Point(12, 12)
            }
          });
          const address = "";
          const infoWindow = new window.naver.maps.InfoWindow({
            content: `
              <div class="naver-marker-info">
                <strong>${item.label}</strong>
                <span>${address || "주소 정보 없음"}</span>
              </div>
            `
          });

          window.naver.maps.Event.addListener(marker, "click", () => {
            if (item.type === "branch") setSelectedBranch(item.name);
            setTypedAiStrategy("");
            map.panTo(position);
          });
        }
        if (markerPositions.length > 1) {
          const bounds = new window.naver.maps.LatLngBounds(markerPositions[0], markerPositions[0]);
          markerPositions.slice(1).forEach((position) => bounds.extend(position));
          map.fitBounds(bounds);
        } else if (markerPositions.length === 1) {
          map.setCenter(markerPositions[0]);
          map.setZoom(selectedBranch ? 16 : initialZoom);
        }
        setMapStatus({ provider: "NAVER", message: "NAVER 지도 로드 완료" });
      } catch (error) {
        console.warn("NAVER Maps SDK loaded but map initialization failed.", error);
        setMapStatus({ provider: "NAVER", message: `NAVER 지도 생성 실패: ${error?.message || "인증 또는 도메인 제한 확인 필요"}` });
        cleanupMap();
        return false;
      }

      return true;
    };

    const initLeafletMap = async () => {
      if (!isMounted || !window.L) return;
      cleanupMap();

      const initialCoords = selectedBranch ? getBranchCoords(selectedBranch) : [36.2, 127.8];
      const initialZoom = selectedBranch ? 16 : 7;
      const map = window.L.map("competitor-map-leaflet", {
        zoomControl: true,
        scrollWheelZoom: true
      }).setView(initialCoords, initialZoom);

      mapRef.current = map;

      window.L.tileLayer("https://{s}.basemaps.cartocdn.com/rastertiles/voyager/{z}/{x}/{y}{r}.png", {
        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors &copy; <a href="https://carto.com/attributions">CARTO</a>',
        subdomains: "abcd",
        maxZoom: 20
      }).addTo(map);

      const rawMapItems = selectedBranch
        ? await getSelectedBranchMapItems()
        : allBranches.map((branchName) => ({
            type: "branch",
            name: branchName,
            label: `이투스247학원 ${branchName}`,
            coords: getBranchCoords(branchName),
            isActive: false
          }));
      const mapItems = getDisplayMapItems(rawMapItems);
      const markerCoords = [];
      mapItems.forEach((item) => {
        const coords = item.displayCoords || item.coords;
        markerCoords.push(coords);
        const customIcon = window.L.divIcon({
          className: `custom-leaflet-marker ${item.type} ${item.isActive ? "active" : ""}`,
          html: getBranchMarkerHtml(item.label, item.isActive, "", item.type),
          iconSize: [24, 24],
          iconAnchor: [12, 12]
        });

        const marker = window.L.marker(coords, { icon: customIcon }).addTo(map);
        marker.on("click", () => {
          if (item.type === "branch") setSelectedBranch(item.name);
          setTypedAiStrategy("");
          map.setView(coords, 16, { animate: true, duration: 0.8 });
        });
      });
      if (markerCoords.length > 1) {
        map.fitBounds(markerCoords, { padding: [36, 36] });
      } else if (markerCoords.length === 1) {
        map.setView(markerCoords[0], selectedBranch ? 16 : initialZoom);
      }
      setMapStatus({ provider: "Leaflet", message: "NAVER 지도 호출 실패로 대체 지도 표시 중" });
    };

    const loadLeaflet = () => {
      if (window.L) {
        initLeafletMap();
        return;
      }

      if (!document.getElementById("leaflet-css")) {
        const link = document.createElement("link");
        link.id = "leaflet-css";
        link.rel = "stylesheet";
        link.href = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.css";
        document.head.appendChild(link);
      }

      let script = document.getElementById("leaflet-js");
      if (!script) {
        script = document.createElement("script");
        script.id = "leaflet-js";
        script.src = "https://unpkg.com/leaflet@1.9.4/dist/leaflet.js";
        script.async = true;
        document.body.appendChild(script);
      }
      script.addEventListener("load", initLeafletMap, { once: true });
    };

    const loadNaverMap = async () => {
      if (window.naver?.maps && await initNaverMap()) return;

      const encodedClientId = encodeURIComponent(naverClientId);
      const sdkCandidates = [
        `https://oapi.map.naver.com/openapi/v3/maps.js?ncpKeyId=${encodedClientId}&submodules=geocoder`,
        `https://openapi.map.naver.com/openapi/v3/maps.js?ncpKeyId=${encodedClientId}&submodules=geocoder`
      ];

      const existingNaverScript = document.getElementById("naver-map-js");
      if (existingNaverScript && existingNaverScript.src !== sdkCandidates[0]) {
        existingNaverScript.remove();
        window.naver = undefined;
      }

      const tryLoadSdk = async (index = 0) => {
        if (!isMounted) return;
        const currentScript = document.getElementById("naver-map-js");
        const isPreferredSdk = currentScript?.src === sdkCandidates[0];
        if (isPreferredSdk && window.naver?.maps && await initNaverMap()) return;
        if (index >= sdkCandidates.length) {
          console.warn("NAVER Maps SDK could not be loaded. Falling back to Leaflet.");
          setMapStatus({ provider: "Leaflet", message: "NAVER SDK 로드 실패: Client ID, Web 서비스 URL, API 활성화 상태 확인 필요" });
          loadLeaflet();
          return;
        }

        document.getElementById("naver-map-js")?.remove();
        const script = document.createElement("script");
        script.id = "naver-map-js";
        script.src = sdkCandidates[index];
        script.async = true;
        setMapStatus({ provider: "NAVER", message: `NAVER SDK 로드 시도 ${index + 1}/${sdkCandidates.length}` });
        document.body.appendChild(script);

        script.addEventListener("load", async () => {
          if (!(await initNaverMap())) tryLoadSdk(index + 1);
        }, { once: true });
        script.addEventListener("error", () => tryLoadSdk(index + 1), { once: true });
      };

      tryLoadSdk();

      sdkWaitTimer = window.setTimeout(() => {
        if (!window.naver?.maps) {
          console.warn("NAVER Maps SDK timed out. Falling back to Leaflet.");
          setMapStatus({ provider: "Leaflet", message: "NAVER SDK 로드 시간 초과: 브라우저 콘솔의 인증 오류 확인 필요" });
          loadLeaflet();
        }
      }, 8000);
    };

    if (isNaverMapEnabled) {
      loadNaverMap();
    } else {
      loadLeaflet();
    }

    return () => {
      isMounted = false;
      if (sdkWaitTimer) window.clearTimeout(sdkWaitTimer);
      cleanupMap();
    };
  }, [page, selectedBranch, allBranches.length]);

  const trackRef = useRef(null);
  const startXRef = useRef(0);
  const scrollLeftRef = useRef(0);
  const translateXRef = useRef(0);
  const targetTranslateXRef = useRef(null);
  const animationFrameRef = useRef(null);
  const isDraggingRef = useRef(false);
  const isHoveredRef = useRef(false);
  const updateMarqueeRef = useRef(null); // Stale 클로저 완벽 타파용 Ref

  // requestAnimationFrame 기반 무한 롤링 루프 (마우스 호버 중이 아닐 때 및 드래그 중이 아닐 때 흐름)
  updateMarqueeRef.current = () => {
    if (isDraggingRef.current) {
      animationFrameRef.current = requestAnimationFrame(updateMarqueeRef.current);
      return;
    }
    if (!trackRef.current) {
      animationFrameRef.current = requestAnimationFrame(updateMarqueeRef.current);
      return;
    }

    const trackWidth = trackRef.current.offsetWidth;
    const halfWidth = trackWidth / 2;
    if (halfWidth <= 0) {
      animationFrameRef.current = requestAnimationFrame(updateMarqueeRef.current);
      return;
    }

    // 기본 등속 롤링 또는 슬라이딩 이동 애니메이션
    if (targetTranslateXRef.current !== null) {
      const diff = targetTranslateXRef.current - translateXRef.current;
      if (Math.abs(diff) < 0.5) {
        translateXRef.current = targetTranslateXRef.current;
        targetTranslateXRef.current = null;
      } else {
        translateXRef.current += diff * 0.12; // 부드러운 감속 이동
      }
    } else {
      // 기본 등속 롤링 (호버 중이 아닐 때만 롤링)
      if (!isHoveredRef.current) {
        translateXRef.current -= 0.7;
      }
    }

    // 무한 롤링 루프를 위한 가로 경계면 seamless 보정
    if (translateXRef.current <= -halfWidth) {
      translateXRef.current += halfWidth;
    } else if (translateXRef.current > 0) {
      translateXRef.current -= halfWidth;
    }

    trackRef.current.style.transform = `translateX(${translateXRef.current}px)`;
    animationFrameRef.current = requestAnimationFrame(updateMarqueeRef.current);
  };

  useEffect(() => {
    if (updateMarqueeRef.current) {
      animationFrameRef.current = requestAnimationFrame(updateMarqueeRef.current);
    }
    return () => {
      if (animationFrameRef.current) {
        cancelAnimationFrame(animationFrameRef.current);
      }
    };
  }, []);

  const handleMarqueeMouseDown = (e) => {
    if (!trackRef.current) return;
    e.preventDefault(); 
    isDraggingRef.current = true;
    startXRef.current = e.pageX;
    scrollLeftRef.current = translateXRef.current;
    targetTranslateXRef.current = null; // 드래그 시작 시 슬라이딩 목표값 취소
  };

  const handleMarqueeMouseMove = (e) => {
    if (!isDraggingRef.current || !trackRef.current) return;
    e.preventDefault();
    const x = e.pageX;
    const walk = x - startXRef.current;
    
    const multiplier = walk < 0 ? 1.35 : 0.75; 
    let newTranslateX = scrollLeftRef.current + walk * multiplier;
    
    const trackWidth = trackRef.current.offsetWidth;
    const halfWidth = trackWidth / 2;
    
    if (newTranslateX < -halfWidth) {
      newTranslateX = newTranslateX + halfWidth;
      startXRef.current = x;
      scrollLeftRef.current = newTranslateX;
    } else if (newTranslateX > 0) {
      newTranslateX = newTranslateX - halfWidth;
      startXRef.current = x;
      scrollLeftRef.current = newTranslateX;
    }
    
    translateXRef.current = newTranslateX;
    trackRef.current.style.transform = `translateX(${newTranslateX}px)`;
  };

  const handleMarqueeMouseUp = (e, compName) => {
    if (!isDraggingRef.current) return;
    isDraggingRef.current = false;
    
    if (e && compName) {
      const distance = Math.abs(e.pageX - startXRef.current);
      if (distance < 6) {
        setSelectedCompForPromo(compName);
      }
    }
  };

  // 좌우 화살표 내비게이션 클릭 핸들러 (카드 너비 500px + gap 24px = 524px 이동)
  const handleMarqueeNavClick = (direction) => {
    if (!trackRef.current) return;
    const step = 524;
    const baseTranslateX = targetTranslateXRef.current !== null ? targetTranslateXRef.current : translateXRef.current;

    if (direction === "prev") {
      // 이전 카드 보기 (카드를 오른쪽으로 밀기)
      targetTranslateXRef.current = baseTranslateX + step;
    } else {
      // 다음 카드 보기 (카드를 왼쪽으로 밀기)
      targetTranslateXRef.current = baseTranslateX - step;
    }
  };

  function markDirty() {
    if (!isHydrated) return;
    setSaveState("변경 감지됨");
  }

  async function forceServerSave() {
    const sortedTabs = getSortedRawTabs(rawTabs);
    setRawTabs(sortedTabs);

    const payload = {
      page,
      rawTabs: sortedTabs,
      activeTabId,
      dashboardTabId
    };

    try {
      window.localStorage.setItem(BROWSER_SAVE_KEY, JSON.stringify(payload));
      setSaveState("강제 서버 저장 중...");

      const saveResponse = await fetch("/api/rawtabs", {
        method: "POST",
        headers: {
          "Content-Type": "application/json"
        },
        body: JSON.stringify(payload),
        cache: "no-store"
      });
      const saveResult = await saveResponse.json().catch(() => null);

      if (!saveResponse.ok) {
        throw new Error(saveResult?.error || "force save failed");
      }

      const verifyResponse = await fetch("/api/rawtabs", { cache: "no-store" });
      const verifyResult = await verifyResponse.json().catch(() => null);

      if (!verifyResponse.ok || !verifyResult?.rawTabs) {
        throw new Error("verify failed");
      }

      const savedAt = formatStatusTimestamp(verifyResult.updatedAt || saveResult?.updatedAt);
      setSaveState(`강제 저장 완료${savedAt ? ` · ${savedAt}` : ""}`);
    } catch (error) {
      console.error("Failed to force save server state.", error);
      setSaveState("강제 서버 저장 실패");
    }
  }

  useEffect(() => {
    let ignore = false;

    async function loadRawTabs() {
      try {
        const response = await fetch("/api/rawtabs", { cache: "no-store" });
        const parsed = await response.json().catch(() => null);

        if (response.ok && parsed?.rawTabs) {
          const normalizedTabs = normalizeRawTabs(parsed.rawTabs);
          const restoredAt = formatStatusTimestamp(parsed.updatedAt);

          if (!ignore && normalizedTabs.length > 0) {
            setRawTabs(normalizedTabs);
            setActiveTabId(parsed.activeTabId || normalizedTabs[0].id);
            setDashboardTabId(parsed.dashboardTabId || OVERVIEW_TAB_ID);
            setPage(parsed.page === "rawdata" ? "rawdata" : "dashboard");
            setSaveState(
              parsed.storageMode === "shared-kv"
                ? `공유 서버 데이터 복원됨${restoredAt ? ` · ${restoredAt}` : ""}`
                : "서버 데이터 복원됨"
            );
            window.localStorage.setItem(BROWSER_SAVE_KEY, JSON.stringify({
              page: parsed.page,
              rawTabs: parsed.rawTabs,
              activeTabId: parsed.activeTabId,
              dashboardTabId: parsed.dashboardTabId
            }));
            return;
          }
        }
      } catch {
        // Fall through to browser cache restore.
      }

      try {
        const cached = window.localStorage.getItem(BROWSER_SAVE_KEY);
        if (!cached) throw new Error("no cached state");

        const parsed = JSON.parse(cached);
        const normalizedTabs = normalizeRawTabs(parsed.rawTabs);

        if (!ignore && normalizedTabs.length > 0) {
          setRawTabs(normalizedTabs);
          setActiveTabId(parsed.activeTabId || normalizedTabs[0].id);
          setDashboardTabId(parsed.dashboardTabId || OVERVIEW_TAB_ID);
          setPage(parsed.page === "rawdata" ? "rawdata" : "dashboard");
          setSaveState("브라우저 저장본 복원됨");
          return;
        }
      } catch {
        if (!ignore) setSaveState("저장된 데이터 없음");
      } finally {
        if (!ignore) setIsHydrated(true);
      }
    }

    loadRawTabs();

    return () => {
      ignore = true;
    };
  }, []);

  useEffect(() => {
    if (!isHydrated) return;

    if (!hasInitializedSaveRef.current) {
      hasInitializedSaveRef.current = true;
      return;
    }

    if (saveTimeoutRef.current) {
      clearTimeout(saveTimeoutRef.current);
    }

    setSaveState("변경 감지됨");

    saveTimeoutRef.current = setTimeout(async () => {
      const sortedTabs = getSortedRawTabs(rawTabs);
      const payload = {
        page,
        rawTabs: sortedTabs,
        activeTabId,
        dashboardTabId
      };

      try {
        window.localStorage.setItem(BROWSER_SAVE_KEY, JSON.stringify(payload));
        setSaveState("저장 중...");
        const response = await fetch("/api/rawtabs", {
          method: "POST",
          headers: {
            "Content-Type": "application/json"
          },
          body: JSON.stringify(payload)
        });

        const result = await response.json().catch(() => null);

        if (!response.ok) {
          if (result?.storageMode === "browser-fallback") {
            setSaveState("공유 저장 미설정, 브라우저에 저장됨");
            return;
          }
          throw new Error("failed to save");
        }

        if (result?.storageMode === "shared-kv") {
          const savedAt = formatStatusTimestamp(result?.updatedAt);
          setSaveState(`공유 서버에 저장됨${savedAt ? ` · ${savedAt}` : ""}`);
        } else {
          setSaveState("서버에 자동 저장됨");
        }
      } catch {
        setSaveState("브라우저에 저장됨");
      }
    }, 400);

    return () => {
      if (saveTimeoutRef.current) {
        clearTimeout(saveTimeoutRef.current);
      }
    };
  }, [activeTabId, dashboardTabId, isHydrated, page, rawTabs]);

  const activeTab = useMemo(
    () => rawTabs.find((tab) => tab.id === activeTabId) ?? rawTabs[0],
    [activeTabId, rawTabs]
  );

  const dashboardRawTabs = useMemo(
    () => rawTabs.filter((tab) => !isSpecialTabKind(tab.kind)),
    [rawTabs]
  );

  const selectedDashboardTab = useMemo(
    () => rawTabs.find((tab) => tab.id === dashboardTabId) ?? rawTabs[0],
    [dashboardTabId, rawTabs]
  );

  const isOverviewDashboard = dashboardTabId === OVERVIEW_TAB_ID;
  const isSpecialDashboard = !isOverviewDashboard && selectedDashboardTab?.kind === SPECIAL_SOCIAL_TAB_KIND;
  const isCollabDashboard = !isOverviewDashboard && selectedDashboardTab?.kind === SPECIAL_COLLAB_TAB_KIND;
  const isFacilityDashboard = !isOverviewDashboard && selectedDashboardTab?.kind === SPECIAL_FACILITY_TAB_KIND;
  const isMentorDashboard = !isOverviewDashboard && selectedDashboardTab?.kind === SPECIAL_MENTOR_TAB_KIND;
  const dashboardTab = isSpecialDashboard || isCollabDashboard || isFacilityDashboard || isMentorDashboard ? null : selectedDashboardTab;

  useEffect(() => {
    setSelectedBranch(null);
    setSelectedCollabBranch(null);
    setSelectedCollabEvent(null);
    setSelectedOverviewBranch(null);
    setSelectedChartEventId(null);
  }, [dashboardTabId]);

  useEffect(() => {
    setBranchKeyword("");
    setAreRegionsExpanded(false);
    setOverviewSearch("");
    setSnsSearch("");
    setMentorSearch("");
    setMentorBranchFilter("all");
    setMentorUnivFilter("all");
  }, [dashboardTabId]);

  useEffect(() => {
    setSelectedCollabEvent(null);
  }, [selectedCollabBranch]);

  useEffect(() => {
    setSelectedChartEventId(null);
  }, [selectedBranch]);

  // Auto-crawl and Adviser useEffect
  useEffect(() => {
    if (!selectedBranch) return;

    let ignore = false;
    const hash = selectedBranch.split("").reduce((acc, char) => acc + char.charCodeAt(0), 0);

    // 1. Generate Simulated Demographics
    const parentPct = 50 + (hash % 20) - 10;
    const mobilePct = 40 + (hash % 30);
    setDemographics({
      pc: 100 - mobilePct,
      mobile: mobilePct,
      parent: parentPct,
      student: 100 - parentPct
    });

    // 2. Generate ROI Campaign metrics
    const beforeVal = 10 + (hash % 15);
    const afterVal = beforeVal + 25 + (hash % 20);
    setRoiData({ before: beforeVal, after: afterVal });

    // 3. Generate Seasonal Alerts
    const alerts = [
      "🔥 6평 직후 '반수반' 키워드 급증 예보: 다음 주 검색량 250% 증가 예측. 지금 예약 페이지 오픈 권장!",
      "📢 수시 접수 기간 '자소서/입시상담' 검색어 폭증기: 학부모 방문 상담 전환율이 높아지는 시즌입니다.",
      "❄️ 수능 직후 '예비고3/윈터스쿨' 검색 유입 최고치: 겨울방학 등록 마케팅 포스팅 발행을 강화하십시오."
    ];
    setSeasonalAlert(alerts[hash % alerts.length]);

    // 4. Fetch NAVER Search Trend API data & Synthesize Adviser Text
    async function loadTrendsAndAdviser() {
      try {
        const res = await fetch(`/api/trend?branch=${encodeURIComponent(selectedBranch)}`);
        const data = await res.json();
        if (ignore) return;

        if (data && data.success && data.trendData) {
          setTrendData(data.trendData);
          
          const statusInfo = getBranchMarketingStatus(selectedBranch, rawTabs);
          const competitors = getCompetitorsForBranch(selectedBranch);
          const adviceText = generateLocalAdviserText(selectedBranch, statusInfo, competitors, data.trendData);
          setTypedAiStrategy(adviceText);
        }
      } catch (err) {
        console.error("Failed to load real NAVER Trend API data:", err);
      }
    }

    async function loadSocialQnaBlogNews() {
      setIsSocialLoading(true);
      try {
        const res = await fetch(`/api/social?branch=${encodeURIComponent(selectedBranch)}`);
        const data = await res.json();
        if (ignore) return;

        if (data && data.success && data.data) {
          setSocialData(data.data);
        }
      } catch (err) {
        console.error("Failed to load NAVER Social Search API data:", err);
      } finally {
        if (!ignore) {
          setIsSocialLoading(false);
        }
      }
    }

    loadTrendsAndAdviser();
    loadSocialQnaBlogNews();

    // 5. Background crawl triggers (runs only once on selectedBranch change)
    const autoCrawl = async () => {
      if (!selectedBranch) return;
      const socialTab = rawTabsRef.current.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
      const branchSnsRow = socialTab?.socialRows?.find(r => r.branch.trim() === selectedBranch);
      const blogUrl = branchSnsRow?.blogUrl || "";

      if (!blogUrl || isMissingChannelUrl(blogUrl)) {
        setCrawledOwnPromotions(prev => ({
          ...prev,
          [selectedBranch]: []
        }));
        setIsCrawling(false);
        return;
      }

      setIsCrawling(true);
      try {
        const res = await fetch(`/api/crawl?blogUrl=${encodeURIComponent(blogUrl)}&branchName=${encodeURIComponent(selectedBranch)}`);
        const data = await res.json();
        if (ignore) return;

        if (data.success) {
          setCrawledOwnPromotions(prev => ({
            ...prev,
            [selectedBranch]: data.promotions
          }));
          setCrawledTimestamps(prev => ({
            ...prev,
            [selectedBranch]: new Date().toLocaleString()
          }));

          if (data.blogStats) {
            setRawTabs(prevTabs => {
              const currentSocialTab = prevTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
              const currentRow = currentSocialTab?.socialRows?.find(r => r.branch.trim() === selectedBranch.trim());
              if (
                currentRow &&
                currentRow.blogRecentPosts === data.blogStats.recent30d &&
                currentRow.blogLastPosted === data.blogStats.lastPosted &&
                currentRow.blogVisitScore === data.blogStats.reactionScore
              ) {
                return prevTabs; // Data identical, prevent state update & re-render
              }

              return prevTabs.map(tab => {
                if (tab.kind !== SPECIAL_SOCIAL_TAB_KIND) return tab;
                const newRows = (tab.socialRows || []).map(r => {
                  if (r.branch.trim() !== selectedBranch.trim()) return r;
                  return {
                    ...r,
                    blogRecentPosts: data.blogStats.recent30d,
                    blogLastPosted: data.blogStats.lastPosted,
                    blogVisitScore: data.blogStats.reactionScore
                  };
                });
                return {
                  ...tab,
                  socialRows: newRows
                };
              });
            });
          }
        }

        // Crawl competitor blogs
        const comps = getCompetitorsForBranch(selectedBranch);
        for (const comp of comps) {
          if (comp.blogUrl) {
            try {
              const compRes = await fetch(`/api/crawl?blogUrl=${encodeURIComponent(comp.blogUrl)}&branchName=${encodeURIComponent(comp.name)}`);
              const compData = await compRes.json();
              if (ignore) return;
              if (compData.success) {
                setCrawledCompPromotions(prev => ({
                  ...prev,
                  [comp.name]: compData.promotions
                }));

                // Restore AI Cache immediately if fresh
                try {
                  const cacheKey = `etoos_ai_cache_${comp.name}`;
                  const cached = localStorage.getItem(cacheKey);
                  if (cached) {
                    const parsed = JSON.parse(cached);
                    const promoKey = compData.promotions.map(p => p.title).join("|");
                    const isFresh = Date.now() - parsed.timestamp < 24 * 60 * 60 * 1000;
                    if (parsed.promoKey === promoKey && isFresh) {
                      setAiAnalyzedData(prev => ({
                        ...prev,
                        [comp.name]: { trend: parsed.trend, guide: parsed.guide }
                      }));
                    }
                  }
                } catch (cacheErr) {
                  console.warn("Error restoring AI cache on crawl:", cacheErr);
                }
              }
            } catch (compErr) {
              console.error(`Error crawling competitor ${comp.name}:`, compErr);
            }
          }
        }
      } catch (err) {
        console.error("Auto crawl error:", err);
      } finally {
        if (!ignore) {
          setIsCrawling(false);
        }
      }
    };

    autoCrawl();

    return () => {
      ignore = true;
    };
  }, [selectedBranch]);

  const triggerAiAnalysis = async (compName, promotions) => {
    if (!compName || !promotions || promotions.length === 0) return;
    if (aiAnalyzedData[compName] || aiLoading[compName]) return;

    const promoKey = promotions.map(p => p.title).join("|");
    const cacheKey = `etoos_ai_cache_${compName}`;
    
    // 1. Check local storage cache
    try {
      const cached = localStorage.getItem(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        // 프로모션 제목 구성이 같고 캐시 유효 기간(24시간)이 경과하지 않은 경우 캐시 데이터 사용
        const isFresh = Date.now() - parsed.timestamp < 24 * 60 * 60 * 1000;
        if (parsed.promoKey === promoKey && isFresh) {
          setAiAnalyzedData(prev => ({
            ...prev,
            [compName]: { trend: parsed.trend, guide: parsed.guide }
          }));
          return;
        }
      }
    } catch (e) {
      console.warn("Error reading from localStorage cache:", e);
    }

    // 2. Fetch from API if no valid cache
    setAiLoading(prev => ({ ...prev, [compName]: true }));
    try {
      const analyzeRes = await fetch("/api/analyze", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ compName, promotions })
      });
      const analyzeData = await analyzeRes.json();
      if (analyzeData.success && !analyzeData.fallback) {
        const result = { trend: analyzeData.trend, guide: analyzeData.guide };
        
        // Update React State
        setAiAnalyzedData(prev => ({
          ...prev,
          [compName]: result
        }));

        // Write to localStorage cache
        try {
          localStorage.setItem(cacheKey, JSON.stringify({
            ...result,
            promoKey,
            timestamp: Date.now()
          }));
        } catch (e) {
          console.warn("Error writing to localStorage cache:", e);
        }
      }
    } catch (analyzeErr) {
      console.error(`Error analyzing competitor ${compName} with AI:`, analyzeErr);
    } finally {
      setAiLoading(prev => ({ ...prev, [compName]: false }));
    }
  };

  const handleGenerateAiStrategy = (branch, statusInfo) => {
    if (!branch) return;
    setIsAiGenerating(true);
    setTimeout(() => {
      const competitors = getCompetitorsForBranch(branch);
      const adviceText = generateLocalAdviserText(branch, statusInfo, competitors, trendData);
      setTypedAiStrategy(adviceText);
      setIsAiGenerating(false);
    }, 800);
  };

  // 1. 자동 롤링 타이머 (자사 / 경쟁사 각각 3.5초 주기)
  useEffect(() => {
    setOwnPromoIndex(0);
    setCompPromoIndex(0);
  }, [selectedBranch, selectedCompForPromo]);

  useEffect(() => {
    const ownPromos = crawledOwnPromotions[selectedBranch] || [];
    const comps = getCompetitorsForBranch(selectedBranch);
    const filteredComps = comps.filter(comp => !selectedCompForPromo || comp.name === selectedCompForPromo);
    
    // 경쟁사 프로모션 병합 리스트 생성 (전체 보기일 때 스택에 쌓을 대상)
    const compPromos = filteredComps.flatMap(comp => {
      const live = crawledCompPromotions[comp.name];
      const items = live && live.length > 0 ? live : comp.promotions;
      return items.map(item => ({ ...item, compName: comp.name }));
    });

    const interval = setInterval(() => {
      if (ownPromos.length > 1 && !ownPromoHovered) {
        setOwnPromoIndex(prev => (prev + 1) % ownPromos.length);
      }
      if (compPromos.length > 1 && !compPromoHovered) {
        setCompPromoIndex(prev => (prev + 1) % compPromos.length);
      }
    }, 3500);

    return () => clearInterval(interval);
  }, [selectedBranch, selectedCompForPromo, crawledOwnPromotions, crawledCompPromotions, ownPromoHovered, compPromoHovered]);

  // 2. 휠 스크롤 고정 및 롤링 핸들러 수동 바인딩 (passive: false 적용)
  useEffect(() => {
    const handleOwnWheel = (e) => {
      const ownPromos = crawledOwnPromotions[selectedBranch] || [];
      if (ownPromos.length <= 1) return;
      
      e.preventDefault();
      e.stopPropagation();

      if (wheelTimeoutRef.current) return;
      wheelTimeoutRef.current = true;

      const direction = e.deltaY > 0 ? 1 : -1;
      setOwnPromoIndex(prev => (prev + direction + ownPromos.length) % ownPromos.length);

      setTimeout(() => {
        wheelTimeoutRef.current = false;
      }, 300);
    };

    const handleCompWheel = (e) => {
      const competitors = getCompetitorsForBranch(selectedBranch);
      const comps = competitors.filter((comp) => !selectedCompForPromo || comp.name === selectedCompForPromo);
      const compPromos = comps.flatMap(comp => {
        const live = crawledCompPromotions[comp.name];
        return live && live.length > 0 ? live : comp.promotions;
      });
      if (compPromos.length <= 1) return;

      e.preventDefault();
      e.stopPropagation();

      if (wheelTimeoutRef.current) return;
      wheelTimeoutRef.current = true;

      const direction = e.deltaY > 0 ? 1 : -1;
      setCompPromoIndex(prev => (prev + direction + compPromos.length) % compPromos.length);

      setTimeout(() => {
        wheelTimeoutRef.current = false;
      }, 300);
    };

    const ownEl = ownPromoContainerRef.current;
    const compEl = compPromoContainerRef.current;

    if (ownEl) {
      ownEl.addEventListener("wheel", handleOwnWheel, { passive: false });
    }
    if (compEl) {
      compEl.addEventListener("wheel", handleCompWheel, { passive: false });
    }

    return () => {
      if (ownEl) ownEl.removeEventListener("wheel", handleOwnWheel);
      if (compEl) compEl.removeEventListener("wheel", handleCompWheel);
    };
  }, [selectedBranch, selectedCompForPromo, crawledOwnPromotions, crawledCompPromotions]);

  const dashboardTabs = useMemo(() => {
    if (isOverviewDashboard) return dashboardRawTabs;
    return dashboardTab ? [dashboardTab] : [];
  }, [dashboardRawTabs, dashboardTab, isOverviewDashboard]);

  const dashboardSummary = useMemo(() => buildDashboardData(dashboardRawTabs), [dashboardRawTabs]);
  const scopedSummary = useMemo(() => buildDashboardData(dashboardTabs), [dashboardTabs]);
  const maxRegionBranches = useMemo(() => Math.max(...scopedSummary.regionOverview.map((item) => item.activeBranches), 0), [scopedSummary.regionOverview]);
  const hoveredRegionData = scopedSummary.regionOverview.find((item) => item.region === hoveredRegion) || null;

  const dashboardScopeLabel = isOverviewDashboard ? "전체 현황" : selectedDashboardTab?.name || "활성화 방안 대시보드";
  const topbarCountLabel = page === "dashboard"
    ? isSpecialDashboard
      ? "진단 항목 수"
      : isCollabDashboard
        ? "진행 횟수"
        : isFacilityDashboard
          ? "등록 URL 수"
          : isMentorDashboard
            ? "멘토단 인원"
            : "이벤트 수"
    : "이벤트 수";
  const topbarBranchLabel = page === "dashboard"
    ? isSpecialDashboard
      ? "평가 지점 수"
      : isCollabDashboard
        ? "참여 지점 수"
        : isFacilityDashboard
          ? "연결 지점 수"
          : isMentorDashboard
            ? "총 1억 장학생 인원"
            : "고유 지점 수"
    : "고유 지점 수";
  const branchOptions = useMemo(
    () =>
      (dashboardTab?.rows || [])
        .map((row) => row.branch.trim())
        .filter(Boolean)
        .filter((branch, index, list) => list.indexOf(branch) === index),
    [dashboardTab]
  );

  const mentorRows = useMemo(() => {
    if (selectedDashboardTab?.kind === SPECIAL_MENTOR_TAB_KIND) {
      return selectedDashboardTab.mentorRows || [];
    }
    return [];
  }, [selectedDashboardTab]);

  const mentorStats = useMemo(() => {
    const total = mentorRows.length;
    const mentors = mentorRows.filter((r) => r.isMentor).length;
    const scholars = total - mentors;
    const amountSum = mentorRows.reduce((sum, r) => sum + (Number(r.amount) || 0), 0);
    return { total, mentors, scholars, amountSum };
  }, [mentorRows]);

  const mentorBranchOptions = useMemo(() => {
    return ["all", ...new Set(mentorRows.map((r) => r.branch.trim()).filter(Boolean))].sort();
  }, [mentorRows]);

  const mentorUnivOptions = useMemo(() => {
    return ["all", ...new Set(mentorRows.map((r) => r.university.trim()).filter(Boolean))].sort();
  }, [mentorRows]);

  const filteredMentorRows = useMemo(() => {
    const search = mentorSearch.trim().toLowerCase();
    const filtered = mentorRows.filter((r) => {
      const matchesSearch = !search ||
        r.name.toLowerCase().includes(search) ||
        (r.year && r.year.toLowerCase().includes(search)) ||
        (r.university && r.university.toLowerCase().includes(search)) ||
        (r.department && r.department.toLowerCase().includes(search)) ||
        (r.branch && r.branch.toLowerCase().includes(search)) ||
        (r.group && r.group.toLowerCase().includes(search)) ||
        (r.memo && r.memo.toLowerCase().includes(search));

      const matchesBranch = mentorBranchFilter === "all" || r.branch === mentorBranchFilter;
      const matchesUniv = mentorUnivFilter === "all" || r.university === mentorUnivFilter;

      return matchesSearch && matchesBranch && matchesUniv;
    });

    return filtered.sort((a, b) => {
      const yA = parseInt(a.year, 10) || 0;
      const yB = parseInt(b.year, 10) || 0;
      if (yB !== yA) return yB - yA;
      return (a.name || "").localeCompare(b.name || "", "ko");
    });
  }, [mentorRows, mentorSearch, mentorBranchFilter, mentorUnivFilter]);

  const mentorsList = useMemo(() => filteredMentorRows.filter((r) => r.isMentor), [filteredMentorRows]);
  const scholarsList = useMemo(() => filteredMentorRows.filter((r) => !r.isMentor), [filteredMentorRows]);

  const branchGroups = useMemo(() => {
    if (!dashboardTab || isOverviewDashboard) return [];

    const grouped = dashboardTab.rows.reduce((acc, row) => {
      const branch = row.branch.trim();
      if (!branch) return acc;

      const region = normalizeRegionLabel(row.region.trim());
      if (!acc.has(region)) {
        acc.set(region, []);
      }

      const list = acc.get(region);
      if (!list.includes(branch)) {
        list.push(branch);
      }
      return acc;
    }, new Map());

    return [...grouped.entries()]
      .map(([region, branches]) => ({
        region,
        branches: branches.sort((a, b) => a.localeCompare(b, "ko"))
      }))
      .sort((a, b) => {
        const aIndex = regionDisplayOrder.indexOf(a.region);
        const bIndex = regionDisplayOrder.indexOf(b.region);
        const safeA = aIndex === -1 ? Number.MAX_SAFE_INTEGER : aIndex;
        const safeB = bIndex === -1 ? Number.MAX_SAFE_INTEGER : bIndex;
        return safeA - safeB || a.region.localeCompare(b.region, "ko");
      });
  }, [dashboardTab, isOverviewDashboard]);

  const visibleBranchGroups = useMemo(() => {
    const keyword = branchKeyword.trim().toLowerCase();
    if (!keyword) return branchGroups;

    return branchGroups
      .map((group) => ({
        ...group,
        branches: group.branches.filter((branch) => branch.toLowerCase().includes(keyword))
      }))
      .filter((group) => group.branches.length > 0);
  }, [branchGroups, branchKeyword]);

  const shouldShowBranchGroups = areRegionsExpanded || branchKeyword.trim().length > 0;

  const branchChartData = useMemo(() => {
    if (!dashboardTab || isOverviewDashboard) return [];

    if (selectedBranch) {
      const row = dashboardTab.rows.find((item) => item.branch.trim() === selectedBranch);
      return dashboardTab.events.map((event) => ({
        id: event.id,
        label: event.name,
        participants: Number(row?.eventValues?.[event.id]?.participants || 0),
        schedule: eventScheduleMap[event.name] || "일정 미등록",
        participatingBranches: Number(row?.eventValues?.[event.id]?.participants || 0) > 0 ? [selectedBranch] : [],
        branchCount: Number(row?.eventValues?.[event.id]?.participants || 0) > 0 ? 1 : 0,
        participationRate: Number(row?.eventValues?.[event.id]?.participants || 0) > 0 ? 100 : 0
      }));
    }

    return dashboardTab.events.map((event) => ({
      id: event.id,
      label: event.name,
      schedule: eventScheduleMap[event.name] || "일정 미등록",
      participants: dashboardTab.rows.reduce(
        (sum, row) => sum + Number(row.eventValues?.[event.id]?.participants || 0),
        0
      ),
      participatingBranches: dashboardTab.rows
        .filter((row) => Number(row.eventValues?.[event.id]?.participants || 0) > 0)
        .map((row) => row.branch.trim())
        .filter(Boolean),
      branchCount: dashboardTab.rows.filter((row) => Number(row.eventValues?.[event.id]?.participants || 0) > 0).length,
      participationRate:
        branchOptions.length > 0
          ? Math.round(
              (dashboardTab.rows.filter((row) => Number(row.eventValues?.[event.id]?.participants || 0) > 0).length / branchOptions.length) * 100
            )
          : 0
    }));
  }, [branchOptions.length, dashboardTab, isOverviewDashboard, selectedBranch]);

  const maxChartParticipants = useMemo(
    () => Math.max(...branchChartData.map((item) => item.participants), 1),
    [branchChartData]
  );

  const participatedEventCount = useMemo(
    () => branchChartData.filter((item) => item.participants > 0).length,
    [branchChartData]
  );

  const participatedEventLabels = useMemo(
    () => branchChartData.filter((item) => item.participants > 0).map((item) => item.label),
    [branchChartData]
  );

  const participationRate = useMemo(() => {
    if (branchChartData.length === 0) return 0;
    return Math.round((participatedEventCount / branchChartData.length) * 100);
  }, [branchChartData.length, participatedEventCount]);

  const selectedChartEvent = useMemo(
    () => branchChartData.find((item) => item.id === selectedChartEventId) || null,
    [branchChartData, selectedChartEventId]
  );

  const activeBranchTooltip = useMemo(
    () =>
      scopedSummary.branchOverview
        .filter((branch) => branch.totalParticipants > 0)
        .map((branch) => branch.branch),
    [scopedSummary.branchOverview]
  );

  const inactiveBranchTooltip = useMemo(
    () =>
      scopedSummary.branchOverview
        .filter((branch) => branch.totalParticipants === 0)
        .map((branch) => branch.branch),
    [scopedSummary.branchOverview]
  );

  const filteredBranchOverview = useMemo(() => {
    const searchTerm = branchSearch.trim().toLowerCase();

    return scopedSummary.branchOverview.filter((branch) => {
      const matchesRegion = regionFilter === "all" || branch.region === regionFilter;
      const matchesSearch =
        !searchTerm ||
        branch.branch.toLowerCase().includes(searchTerm) ||
        branch.region.toLowerCase().includes(searchTerm) ||
        branch.activePlans.some((plan) => plan.toLowerCase().includes(searchTerm));

      return matchesRegion && matchesSearch;
    });
  }, [branchSearch, regionFilter, scopedSummary.branchOverview]);

  const snsDashboardRows = useMemo(
    () => (selectedDashboardTab?.kind === SPECIAL_SOCIAL_TAB_KIND ? (selectedDashboardTab.socialRows || []).map((row) => summarizeSnsRow(row)) : []),
    [selectedDashboardTab]
  );

  const filteredSnsDashboardRows = useMemo(() => {
    const keyword = snsSearch.trim().toLowerCase();
    if (!keyword) return snsDashboardRows.filter((row) => row.branch.trim());

    return snsDashboardRows.filter((row) => row.branch.trim() && row.branch.toLowerCase().includes(keyword));
  }, [snsDashboardRows, snsSearch]);

  const snsSourceRows = useMemo(() => {
    const snsTab = rawTabs.find((tab) => tab.kind === SPECIAL_SOCIAL_TAB_KIND);
    return snsTab ? (snsTab.socialRows || []).map((row) => summarizeSnsRow(row)) : [];
  }, [rawTabs]);

  const snsGradeGroups = useMemo(() => ({
    A: filteredSnsDashboardRows.filter((row) => row.grade === "A"),
    B: filteredSnsDashboardRows.filter((row) => row.grade === "B"),
    C: filteredSnsDashboardRows.filter((row) => row.grade === "C"),
    D: filteredSnsDashboardRows.filter((row) => row.grade === "D")
  }), [filteredSnsDashboardRows]);

  const snsSummary = useMemo(() => {
    const totalBranches = filteredSnsDashboardRows.length;
    const averageScore = totalBranches > 0 ? Number((filteredSnsDashboardRows.reduce((sum, row) => sum + row.finalScore, 0) / totalBranches).toFixed(1)) : 0;
    const bothChannels = filteredSnsDashboardRows.filter((row) => row.hasBlog && row.hasInstagram).length;
    const missingChannels = filteredSnsDashboardRows.filter((row) => !row.hasBlog || !row.hasInstagram).length;
    const topBranches = [...filteredSnsDashboardRows].sort((a, b) => b.finalScore - a.finalScore).slice(0, 8);
    const lowBranches = [...filteredSnsDashboardRows].sort((a, b) => a.finalScore - b.finalScore).slice(0, 8);
    return { totalBranches, averageScore, bothChannels, missingChannels, topBranches, lowBranches };
  }, [filteredSnsDashboardRows]);

  const overallSnsSummary = useMemo(() => {
    const totalBranches = snsSourceRows.filter((row) => row.branch.trim()).length;
    const averageScore = totalBranches > 0 ? Number((snsSourceRows.reduce((sum, row) => sum + row.finalScore, 0) / totalBranches).toFixed(1)) : 0;
    return {
      totalBranches,
      averageScore,
      A: snsSourceRows.filter((row) => row.grade === "A").length,
      B: snsSourceRows.filter((row) => row.grade === "B").length,
      C: snsSourceRows.filter((row) => row.grade === "C").length,
      D: snsSourceRows.filter((row) => row.grade === "D").length
    };
  }, [snsSourceRows]);

  const overallCollabTab = useMemo(
    () => rawTabs.find((tab) => tab.kind === SPECIAL_COLLAB_TAB_KIND) || null,
    [rawTabs]
  );

  const overallCollabSummary = useMemo(
    () =>
      overallCollabTab
        ? buildCollabSummary(overallCollabTab)
        : {
            totalBranches: 0,
            activeBranches: 0,
            inactiveBranches: 0,
            totalUrls: 0,
            uniqueEvents: 0,
            branchRows: [],
            branchOptions: [],
            groupedBranches: new Map(),
            eventOverview: []
          },
    [overallCollabTab]
  );

  const facilityDashboardSummary = useMemo(
    () =>
      isFacilityDashboard
        ? buildFacilitySummary(selectedDashboardTab)
        : {
            totalBranches: 0,
            activeBranches: 0,
            inactiveBranches: 0,
            totalUrls: 0,
            branchRows: [],
            groupedBranches: new Map()
          },
    [isFacilityDashboard, selectedDashboardTab]
  );

  const facilityActiveBranchTooltip = useMemo(
    () => facilityDashboardSummary.branchRows.filter((row) => row.url).map((row) => row.branch),
    [facilityDashboardSummary.branchRows]
  );

  const facilityInactiveBranchTooltip = useMemo(
    () => facilityDashboardSummary.branchRows.filter((row) => !row.url).map((row) => row.branch),
    [facilityDashboardSummary.branchRows]
  );

  const facilityRegionGroups = useMemo(
    () =>
      [...facilityDashboardSummary.groupedBranches.entries()]
        .map(([region, branches]) => ({
          region,
          branches: [...branches].sort((a, b) => a.branch.localeCompare(b.branch, "ko"))
        }))
        .sort((a, b) => {
          const aIndex = regionDisplayOrder.indexOf(a.region);
          const bIndex = regionDisplayOrder.indexOf(b.region);
          const safeA = aIndex === -1 ? Number.MAX_SAFE_INTEGER : aIndex;
          const safeB = bIndex === -1 ? Number.MAX_SAFE_INTEGER : bIndex;
          return safeA - safeB || a.region.localeCompare(b.region, "ko");
        }),
    [facilityDashboardSummary.groupedBranches]
  );

  const collabDashboardSummary = useMemo(
    () =>
      isCollabDashboard
        ? buildCollabSummary(selectedDashboardTab)
        : {
            totalBranches: 0,
            activeBranches: 0,
            inactiveBranches: 0,
            totalUrls: 0,
            uniqueEvents: 0,
            branchRows: [],
            branchOptions: [],
            groupedBranches: new Map(),
            eventOverview: []
          },
    [isCollabDashboard, selectedDashboardTab]
  );

  const collabActiveBranchTooltip = useMemo(
    () =>
      collabDashboardSummary.branchRows
        .filter((row) => row.urlCount > 0)
        .map((row) => row.branch),
    [collabDashboardSummary.branchRows]
  );

  const collabInactiveBranchTooltip = useMemo(
    () =>
      collabDashboardSummary.branchRows
        .filter((row) => row.urlCount === 0)
        .map((row) => row.branch),
    [collabDashboardSummary.branchRows]
  );

  const topbarCountValue = page === "dashboard"
    ? isSpecialDashboard
      ? specialSocialColumns.length
      : isCollabDashboard
        ? collabDashboardSummary.uniqueEvents
        : isFacilityDashboard
          ? facilityDashboardSummary.totalUrls
          : isMentorDashboard
            ? mentorStats.mentors
            : scopedSummary.totalEvents
    : dashboardSummary.totalEvents;
  const topbarBranchValue = page === "dashboard"
    ? isSpecialDashboard
      ? snsSummary.totalBranches
      : isCollabDashboard
        ? collabDashboardSummary.activeBranches
        : isFacilityDashboard
          ? facilityDashboardSummary.activeBranches
          : isMentorDashboard
            ? mentorStats.scholars
            : scopedSummary.uniqueBranches
    : dashboardSummary.uniqueBranches;

  const collabBranchGroups = useMemo(() => {
    if (!isCollabDashboard) return [];

    return [...collabDashboardSummary.groupedBranches.entries()]
      .map(([region, branches]) => ({
        region,
        branches: [...branches].sort((a, b) => a.localeCompare(b, "ko"))
      }))
      .sort((a, b) => {
        const aIndex = regionDisplayOrder.indexOf(a.region);
        const bIndex = regionDisplayOrder.indexOf(b.region);
        const safeA = aIndex === -1 ? Number.MAX_SAFE_INTEGER : aIndex;
        const safeB = bIndex === -1 ? Number.MAX_SAFE_INTEGER : bIndex;
        return safeA - safeB || a.region.localeCompare(b.region, "ko");
      });
  }, [collabDashboardSummary.groupedBranches, isCollabDashboard]);

  const visibleCollabBranchGroups = useMemo(() => {
    const keyword = branchKeyword.trim().toLowerCase();
    if (!keyword) return collabBranchGroups;

    return collabBranchGroups
      .map((group) => ({
        ...group,
        branches: group.branches.filter((branch) => branch.toLowerCase().includes(keyword))
      }))
      .filter((group) => group.branches.length > 0);
  }, [branchKeyword, collabBranchGroups]);

  const selectedCollabBranchRow = useMemo(
    () => collabDashboardSummary.branchRows.find((row) => row.branch === selectedCollabBranch) || null,
    [collabDashboardSummary.branchRows, selectedCollabBranch]
  );

  const collabEventList = useMemo(() => {
    if (selectedCollabBranchRow) {
      return selectedCollabBranchRow.events.map((event) => ({
        id: event.name,
        label: event.name,
        branchCount: 1,
        urlCount: event.links.length
      }));
    }

    return collabDashboardSummary.eventOverview;
  }, [collabDashboardSummary.eventOverview, selectedCollabBranchRow]);

  const selectedCollabEventData = useMemo(() => {
    if (!selectedCollabEvent) return null;

    const channelOrder = ["홈페이지", "블로그", "인스타/언론기사"];
    const createChannelMap = () =>
      Object.fromEntries(channelOrder.map((channel) => [channel, []]));

    if (selectedCollabBranchRow) {
      const event = selectedCollabBranchRow.events.find((item) => item.name === selectedCollabEvent);
      if (!event) return null;

      const channels = createChannelMap();
      event.links.forEach((link) => {
        if (!channels[link.label]) channels[link.label] = [];
        channels[link.label].push({
          branch: selectedCollabBranchRow.branch,
          region: selectedCollabBranchRow.region,
          url: link.url,
          label: link.label
        });
      });

      return {
        name: selectedCollabEvent,
        channels
      };
    }

    const channels = createChannelMap();

    collabDashboardSummary.branchRows.forEach((row) => {
      const event = row.events.find((item) => item.name === selectedCollabEvent);
      if (!event) return;

      event.links.forEach((link) => {
        if (!channels[link.label]) channels[link.label] = [];
        channels[link.label].push({
          label: link.label,
          url: link.url,
          branch: row.branch,
          region: row.region
        });
      });
    });

    Object.keys(channels).forEach((channel) => {
      channels[channel] = channels[channel]
        .sort((a, b) => a.branch.localeCompare(b.branch, "ko"))
        .map((link) => ({
          ...link,
          id: `${channel}-${link.branch}-${link.url}`
        }));
    });

    const hasAnyLinks = Object.values(channels).some((items) => items.length > 0);

    return hasAnyLinks
      ? {
          name: selectedCollabEvent,
          channels
        }
      : null;
  }, [collabDashboardSummary.branchRows, selectedCollabBranchRow, selectedCollabEvent]);

  const overallBranchScoreboard = useMemo(() => {
    const branchMap = new Map();
    const snsScoreMap = new Map(
      snsSourceRows
        .filter((row) => row.branch.trim())
        .map((row) => [normalizeBranchKey(row.branch), row])
    );
    const planTypeCount = dashboardRawTabs.length + (overallCollabTab ? 1 : 0);

    dashboardRawTabs.forEach((tab) => {
      tab.rows.forEach((row) => {
        const branch = row.branch.trim();
        if (!branch) return;

        if (!branchMap.has(branch)) {
          branchMap.set(branch, {
            branch,
            region: normalizeRegionLabel(row.region.trim()),
            eligibleEvents: 0,
            participatedEvents: 0,
            inactiveEvents: 0,
            totalParticipants: 0,
            collabUrlCount: 0,
            activePlans: new Set()
          });
        }

        const target = branchMap.get(branch);
        if (!target.region && row.region.trim()) {
          target.region = normalizeRegionLabel(row.region.trim());
        }

        target.eligibleEvents += tab.events.length;

        tab.events.forEach((event) => {
          const participants = Number(row.eventValues?.[event.id]?.participants || 0);
          if (participants > 0) {
            target.participatedEvents += 1;
            target.totalParticipants += participants;
            target.activePlans.add(tab.name);
          } else {
            target.inactiveEvents += 1;
          }
        });
      });
    });

    if (overallCollabTab) {
      overallCollabSummary.branchRows.forEach((row) => {
        const branch = row.branch.trim();
        if (!branch) return;

        if (!branchMap.has(branch)) {
          branchMap.set(branch, {
            branch,
            region: normalizeRegionLabel(row.region.trim()),
            eligibleEvents: 0,
            participatedEvents: 0,
            inactiveEvents: 0,
            totalParticipants: 0,
            collabUrlCount: 0,
            activePlans: new Set()
          });
        }

        const target = branchMap.get(branch);
        if (!target.region && row.region.trim()) {
          target.region = normalizeRegionLabel(row.region.trim());
        }

        target.eligibleEvents += overallCollabSummary.uniqueEvents;
        target.participatedEvents += row.events.length;
        target.inactiveEvents += Math.max(overallCollabSummary.uniqueEvents - row.events.length, 0);
        target.collabUrlCount += row.urlCount;

        if (row.events.length > 0) {
          target.activePlans.add(overallCollabTab.name);
        }
      });
    }

    const rawList = [...branchMap.values()];
    const maxActivityVolume = Math.max(...rawList.map((item) => item.totalParticipants + item.collabUrlCount), 1);

    const branches = rawList
      .map((item) => {
        const snsMatch = snsScoreMap.get(normalizeBranchKey(item.branch));
        const participationRate = item.eligibleEvents > 0 ? Math.round((item.participatedEvents / item.eligibleEvents) * 100) : 0;
        const activityVolume = item.totalParticipants + item.collabUrlCount;
        const participantScore = Math.round((activityVolume / maxActivityVolume) * 100);
        const planCoverage = planTypeCount > 0 ? Math.round((item.activePlans.size / planTypeCount) * 100) : 0;
        const stabilityScore = Math.max(0, 100 - Math.round((item.inactiveEvents / Math.max(item.eligibleEvents, 1)) * 100));
        const operationScore = Math.round(
          participationRate * 0.45 +
          participantScore * 0.3 +
          planCoverage * 0.2 +
          stabilityScore * 0.05
        );
        const snsScore = snsMatch?.finalScore ?? null;
        const score = snsScore === null
          ? operationScore
          : Math.round(operationScore * 0.5 + snsScore * 0.5);

        return {
          branch: item.branch,
          region: item.region || "기타",
          eligibleEvents: item.eligibleEvents,
          participatedEvents: item.participatedEvents,
          inactiveEvents: item.inactiveEvents,
          totalParticipants: item.totalParticipants,
          collabUrlCount: item.collabUrlCount,
          activePlanCount: item.activePlans.size,
          activePlans: [...item.activePlans].sort((a, b) => a.localeCompare(b, "ko")),
          participationRate,
          operationScore,
          score,
          grade: getBranchGrade(score),
          snsScore,
          snsGrade: snsMatch?.grade ?? null
        };
      })
      .sort((a, b) => b.score - a.score || b.totalParticipants - a.totalParticipants || a.branch.localeCompare(b.branch, "ko"));

    const grouped = {
      "A그룹": [],
      "B그룹": [],
      "C그룹": [],
      "D그룹": []
    };

    branches.forEach((branch) => {
      grouped[branch.grade].push(branch);
    });

    return {
      branches,
      grouped,
        avgScore: branches.length > 0 ? Math.round(branches.reduce((sum, item) => sum + item.score, 0) / branches.length) : 0,
        topBranch: branches[0] || null,
        atRiskCount: grouped["D그룹"].length
      };
  }, [dashboardRawTabs, overallCollabSummary, overallCollabTab, snsSourceRows]);

  const filteredOverviewBranchScoreboard = useMemo(() => {
    const keyword = overviewSearch.trim().toLowerCase();
    const branches = !keyword
      ? overallBranchScoreboard.branches
      : overallBranchScoreboard.branches.filter((branch) =>
          branch.branch.toLowerCase().includes(keyword) ||
          branch.region.toLowerCase().includes(keyword) ||
          branch.activePlans.some((plan) => plan.toLowerCase().includes(keyword))
        );

    const grouped = {
      "A그룹": [],
      "B그룹": [],
      "C그룹": [],
      "D그룹": []
    };

    branches.forEach((branch) => {
      grouped[branch.grade].push(branch);
    });

    return {
      branches,
      grouped,
      avgScore: branches.length > 0 ? Math.round(branches.reduce((sum, item) => sum + item.score, 0) / branches.length) : 0,
      topBranch: branches[0] || null,
      atRiskCount: grouped["D그룹"].length
    };
  }, [overallBranchScoreboard.branches, overviewSearch]);

  function updateActiveTab(mutator) {
    markDirty();
    setRawTabs((current) => current.map((tab) => (tab.id === activeTabId ? mutator(tab) : tab)));
  }

  function addRawTab() {
    markDirty();
    const nextId = `tab-${Date.now()}`;
    const nextIndex = rawTabs.length + 1;
    const nextTab = createTab(nextId, `새 활성화 방안 ${nextIndex}`, []);
    setRawTabs((current) => [...current, nextTab]);
    setActiveTabId(nextId);
    setDashboardTabId(nextId);
    setPage("rawdata");
  }

  function removeActiveTab() {
    if (!activeTab || rawTabs.length === 1) return;
    markDirty();
    const remaining = rawTabs.filter((tab) => tab.id !== activeTab.id);
    setRawTabs(remaining);
    setActiveTabId(remaining[0].id);
    if (dashboardTabId === activeTab.id) {
      setDashboardTabId(OVERVIEW_TAB_ID);
    }
  }

  function updateTabName(value) {
    updateActiveTab((tab) => ({ ...tab, name: value || "이름 없는 탭" }));
  }

  function updateEventName(eventId, value) {
    updateActiveTab((tab) => ({
      ...tab,
      events: tab.events.map((event) =>
        event.id === eventId ? { ...event, name: value || "이름 없는 이벤트" } : event
      )
    }));
  }

  function updateBaseCell(rowIndex, field, value) {
    updateActiveTab((tab) => ({
      ...tab,
      rows: tab.rows.map((row, index) => (index === rowIndex ? { ...row, [field]: value } : row))
    }));
  }

  function updateSpecialCell(rowIndex, field, value) {
    updateActiveTab((tab) => ({
      ...tab,
      socialRows: (tab.socialRows || []).map((row, index) => {
        if (index !== rowIndex) return row;
        const isNumericField = [
          "blogRecentPosts",
          "blogVisitScore",
          "instagramRecentPosts",
          "instagramDesignScore",
          "instagramReactionScore",
          "profileSetupScore",
          "featureUsageScore",
          "ctaScore",
          "linkHealthScore",
          "brandInfoScore"
        ].includes(field);

        return {
          ...row,
          [field]: isNumericField ? normalizeParticipantValue(value) : value
        };
      })
    }));
  }

  function updateCollabCell(rowIndex, field, value) {
    updateActiveTab((tab) => ({
      ...tab,
      collabRows: (tab.collabRows || []).map((row, index) =>
        index === rowIndex
          ? {
              ...row,
              values: {
                ...row.values,
                [field]: value
              }
            }
          : row
      )
    }));
  }

  function renameCollabEvent(previousName, nextName) {
    const trimmedName = nextName.trim();
    if (!previousName || !trimmedName || previousName === trimmedName) return;

    const nextColumns = [
      `${trimmedName} 홈페이지`,
      `${trimmedName} 블로그`,
      `${trimmedName} 인스타/언론기사`
    ];

    const duplicateExists = (activeTab?.collabColumns || []).some((column) => {
      const { eventName } = parseCollabColumnLabel(column);
      return eventName === trimmedName && eventName !== previousName;
    });

    if (duplicateExists) {
      setSaveState("같은 이름의 협업 이벤트가 이미 있습니다");
      return;
    }

    updateActiveTab((tab) => {
      const previousColumns = [
        `${previousName} 홈페이지`,
        `${previousName} 블로그`,
        `${previousName} 인스타/언론기사`
      ];

      const collabColumns = (tab.collabColumns || defaultCollabColumns).map((column) => {
        const matchedIndex = previousColumns.indexOf(column);
        return matchedIndex === -1 ? column : nextColumns[matchedIndex];
      });

      return {
        ...tab,
        collabColumns,
        collabRows: (tab.collabRows || []).map((row) => {
          const nextValues = { ...(row.values || {}) };

          previousColumns.forEach((column, index) => {
            nextValues[nextColumns[index]] = nextValues[column] ?? "";
            delete nextValues[column];
          });

          return {
            ...row,
            values: nextValues
          };
        })
      };
    });
  }

  function updateFacilityCell(rowIndex, field, value) {
    updateActiveTab((tab) => ({
      ...tab,
      facilityRows: (tab.facilityRows || []).map((row, index) =>
        index === rowIndex
          ? {
              ...row,
              [field]: value
            }
          : row
      )
    }));
  }

  function updateMentorCell(rowIndex, field, value) {
    updateActiveTab((tab) => ({
      ...tab,
      mentorRows: (tab.mentorRows || []).map((row, index) => {
        if (index === rowIndex) {
          let updatedValue = value;
          if (field === "isMentor") {
            updatedValue = Boolean(value);
          } else if (field === "amount") {
            updatedValue = Number(value) || 0;
          }
          return {
            ...row,
            [field]: updatedValue
          };
        }
        return row;
      })
    }));
  }

  function updateEventCell(rowIndex, eventId, field, value) {
    updateActiveTab((tab) => ({
      ...tab,
      rows: tab.rows.map((row, index) => {
        if (index !== rowIndex) return row;

        const currentValue = row.eventValues?.[eventId] || { status: "X", participants: "0" };
        const nextValue = { ...currentValue };

        if (field === "status") {
          nextValue.status = value === "O" ? "O" : "X";
          if (nextValue.status === "X") {
            nextValue.participants = "0";
          }
        }

        if (field === "participants") {
          nextValue.participants = normalizeParticipantValue(value);
          if (Number(nextValue.participants) > 0) {
            nextValue.status = "O";
          }
        }

        return {
          ...row,
          eventValues: {
            ...row.eventValues,
            [eventId]: nextValue
          }
        };
      })
    }));
  }

  function addRow() {
    updateActiveTab((tab) => {
      if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) {
        return {
          ...tab,
          socialRows: [...(tab.socialRows || []), createSpecialSocialRow()]
        };
      }

      if (tab.kind === SPECIAL_COLLAB_TAB_KIND) {
        return {
          ...tab,
          collabRows: [...(tab.collabRows || []), createSpecialCollabRow(tab.collabColumns || defaultCollabColumns)]
        };
      }

      if (tab.kind === SPECIAL_FACILITY_TAB_KIND) {
        return {
          ...tab,
          facilityRows: [...(tab.facilityRows || []), createSpecialFacilityRow()]
        };
      }

      if (tab.kind === SPECIAL_MENTOR_TAB_KIND) {
        return {
          ...tab,
          mentorRows: [...(tab.mentorRows || []), createSpecialMentorRow()]
        };
      }

      return {
        ...tab,
        rows: [...tab.rows, createRow(tab.events.map((event) => event.id))]
      };
    });
  }

  const exportToExcel = (tab) => {
    try {
      let data = [];
      if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) {
        data = (tab.socialRows || []).map(r => {
          const { id, ...rest } = r;
          return rest;
        });
      } else if (tab.kind === SPECIAL_FACILITY_TAB_KIND) {
        data = (tab.facilityRows || []).map(r => {
          const { id, ...rest } = r;
          return rest;
        });
      } else if (tab.kind === SPECIAL_MENTOR_TAB_KIND) {
        data = (tab.mentorRows || []).map(r => {
          const { id, ...rest } = r;
          return rest;
        });
      } else if (tab.kind === SPECIAL_COLLAB_TAB_KIND) {
        data = (tab.collabRows || []).map(r => r.values);
      } else {
        data = (tab.rows || []).map(r => {
          const rowObj = { "지역": r.region, "지점": r.branch };
          tab.events.forEach(e => {
            rowObj[e.name] = r.eventValues?.[e.id]?.participants || "";
          });
          return rowObj;
        });
      }
      const worksheet = XLSX.utils.json_to_sheet(data);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, tab.name || "Sheet1");
      XLSX.writeFile(workbook, `${tab.name || "export"}.xlsx`);
    } catch (err) {
      console.error(err);
      alert("엑셀 내보내기 중 오류가 발생했습니다.");
    }
  };

  function removeRow(rowIndex) {
    updateActiveTab((tab) => {
      if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) {
        return {
          ...tab,
          socialRows: (tab.socialRows || []).filter((_, index) => index !== rowIndex)
        };
      }

      if (tab.kind === SPECIAL_COLLAB_TAB_KIND) {
        return {
          ...tab,
          collabRows: (tab.collabRows || []).filter((_, index) => index !== rowIndex)
        };
      }

      if (tab.kind === SPECIAL_FACILITY_TAB_KIND) {
        return {
          ...tab,
          facilityRows: (tab.facilityRows || []).filter((_, index) => index !== rowIndex)
        };
      }

      if (tab.kind === SPECIAL_MENTOR_TAB_KIND) {
        return {
          ...tab,
          mentorRows: (tab.mentorRows || []).filter((_, index) => index !== rowIndex)
        };
      }

      return {
        ...tab,
        rows: tab.rows.filter((_, index) => index !== rowIndex)
      };
    });
  }

  function addEvent() {
    if (activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND || activeTab?.kind === SPECIAL_FACILITY_TAB_KIND || activeTab?.kind === SPECIAL_MENTOR_TAB_KIND) return;
    const nextName = window.prompt("추가할 이벤트명을 입력하세요.", "신규 이벤트");
    if (!nextName) return;
    const trimmedName = nextName.trim();
    if (!trimmedName) return;

    if (activeTab?.kind === SPECIAL_COLLAB_TAB_KIND) {
      const nextColumns = [
        `${trimmedName} 홈페이지`,
        `${trimmedName} 블로그`,
        `${trimmedName} 인스타/언론기사`
      ];

      const hasDuplicate = nextColumns.some((column) => (activeTab.collabColumns || []).includes(column));
      if (hasDuplicate) {
        setSaveState("같은 이름의 협업 이벤트 열이 이미 있습니다");
        return;
      }

      updateActiveTab((tab) => {
        const existingColumns = (tab.collabColumns || defaultCollabColumns).filter((column) => column !== "지역" && column !== "지점");
        const collabColumns = ["지역", "지점", ...nextColumns, ...existingColumns];

        return {
          ...tab,
          collabColumns,
          collabRows: (tab.collabRows || []).map((row) =>
            createSpecialCollabRow(collabColumns, {
              ...(row.values || {}),
              id: row.id
            })
          )
        };
      });
      return;
    }

    updateActiveTab((tab) => {
      const newEvent = createEvent(trimmedName);
      return {
        ...tab,
        events: [...tab.events, newEvent],
        rows: tab.rows.map((row) => ({
          ...row,
          eventValues: {
            ...row.eventValues,
            [newEvent.id]: { status: "X", participants: "0" }
          }
        }))
      };
    });
  }

  async function importSnsWorkbook(file) {
    if (!file || activeTab?.kind !== SPECIAL_SOCIAL_TAB_KIND) return;

    const shouldReplace = window.confirm("현재 SNS 탭 데이터를 업로드한 엑셀의 입력 원본 시트 데이터로 교체할까요?");
    if (!shouldReplace) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const { sheetName, rows } = extractSnsRowsFromWorkbook(arrayBuffer);

      if (rows.length === 0) {
        setSaveState("가져올 SNS 데이터 없음");
        return;
      }

      updateActiveTab((tab) => ({
        ...tab,
        socialRows: rows
      }));
      setSaveState(`엑셀 '${sheetName}' 시트 반영됨`);
    } catch (error) {
      console.error("Failed to import SNS workbook.", error);
      setSaveState("엑셀 불러오기 실패");
    }
  }

  async function importCollabWorkbook(file) {
    if (!file || activeTab?.kind !== SPECIAL_COLLAB_TAB_KIND) return;

    const shouldReplace = window.confirm("현재 협업이벤트 탭 데이터를 업로드한 엑셀 시트 데이터로 교체할까요?");
    if (!shouldReplace) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const { sheetName, tab } = extractCollabRowsFromWorkbook(arrayBuffer, activeTab?.name || "협업이벤트");

      updateActiveTab((currentTab) => ({
        ...currentTab,
        collabColumns: tab.collabColumns,
        collabRows: tab.collabRows
      }));
      setSaveState(`엑셀 '${sheetName}' 시트 반영됨`);
    } catch (error) {
      console.error("Failed to import collab workbook.", error);
      setSaveState("엑셀 불러오기 실패");
    }
  }

  async function importFacilityWorkbook(file) {
    if (!file || activeTab?.kind !== SPECIAL_FACILITY_TAB_KIND) return;

    const shouldReplace = window.confirm("현재 지점시설영상 탭 데이터를 업로드한 엑셀 시트 데이터로 교체할까요?");
    if (!shouldReplace) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const { sheetName, tab } = extractFacilityRowsFromWorkbook(arrayBuffer, activeTab?.name || "지점시설영상");

      updateActiveTab((currentTab) => ({
        ...currentTab,
        facilityColumns: tab.facilityColumns,
        facilityRows: tab.facilityRows
      }));
      setSaveState(`엑셀 '${sheetName}' 시트 반영됨`);
    } catch (error) {
      console.error("Failed to import facility workbook.", error);
      setSaveState("엑셀 불러오기 실패");
    }
  }

  async function importMentorWorkbook(file) {
    if (!file || activeTab?.kind !== SPECIAL_MENTOR_TAB_KIND) return;

    const shouldReplace = window.confirm("현재 멘토단 및 장학생 탭 데이터를 업로드한 엑셀 시트 데이터로 교체할까요?");
    if (!shouldReplace) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const { sheetName, rows } = extractMentorRowsFromWorkbook(arrayBuffer);

      if (rows.length === 0) {
        setSaveState("가져올 멘토단 및 장학생 데이터 없음");
        return;
      }

      updateActiveTab((tab) => ({
        ...tab,
        mentorRows: rows
      }));
      setSaveState(`엑셀 '${sheetName}' 시트 반영됨`);
    } catch (error) {
      console.error("Failed to import mentor workbook.", error);
      setSaveState("엑셀 불러오기 실패");
    }
  }

  async function importDefaultWorkbook(file) {
    if (!file || activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND || activeTab?.kind === SPECIAL_COLLAB_TAB_KIND || activeTab?.kind === SPECIAL_FACILITY_TAB_KIND || activeTab?.kind === SPECIAL_MENTOR_TAB_KIND) return;

    const shouldReplace = window.confirm("현재 이벤트 탭 데이터를 업로드한 엑셀 시트 데이터로 교체할까요?");
    if (!shouldReplace) return;

    try {
      const arrayBuffer = await file.arrayBuffer();
      const { sheetName, tab } = extractDefaultTabFromWorkbook(arrayBuffer, activeTab?.name || "불러온 이벤트");

      updateActiveTab((currentTab) => ({
        ...currentTab,
        events: tab.events,
        rows: tab.rows
      }));
      setSaveState(`엑셀 '${sheetName}' 시트 반영됨`);
    } catch (error) {
      console.error("Failed to import default workbook.", error);
      setSaveState("엑셀 불러오기 실패");
    }
  }

  function removeEvent(eventId) {
    if (isSpecialTabKind(activeTab?.kind)) return;
    updateActiveTab((tab) => ({
      ...tab,
      events: tab.events.filter((event) => event.id !== eventId),
      rows: tab.rows.map((row) => {
        const nextEventValues = { ...row.eventValues };
        delete nextEventValues[eventId];
        return {
          ...row,
          eventValues: nextEventValues
        };
      })
    }));
  }

    return (
    <div className="dashboard-wrapper">
      {/* 1. 상단 다크 네비바 */}
      <nav className="premium-navbar">
        <div 
          className="premium-navbar-logo-container" 
          onClick={() => {
            sortMentorRowsState();
            setDashboardTabId(OVERVIEW_TAB_ID);
            setPage("dashboard");
            window.scrollTo({ top: 0, behavior: 'smooth' });
          }} 
          style={{ cursor: "pointer" }}
        >
          <img src="/logo.png" className="premium-navbar-logo" alt="ETOOS ECI Logo" />
          <span className="premium-navbar-title">마케팅 대시보드</span>
        </div>
        <ul className="premium-navbar-menu">
          {[
            { key: "dashboard", label: "전체 현황" },
            { key: "sns", label: "SNS 분석" },
            { key: "rawdata", label: "RAWDATASTUDIO" }
          ].map((menu) => (
            <li key={menu.key} className="premium-navbar-menu-item">
              <a 
                href="#" 
                onClick={(e) => { e.preventDefault(); handleNavMenuClick(menu.key); }}
                style={{ 
                  color: 
                    (menu.key === "dashboard" && page === "dashboard") || 
                    (menu.key === "sns" && page === "sns") || 
                    (menu.key === "rawdata" && page === "rawdata") 
                      ? "#ffffff" 
                      : "rgba(255, 255, 255, 0.7)",
                  fontWeight: 
                    (menu.key === "dashboard" && page === "dashboard") || 
                    (menu.key === "sns" && page === "sns") || 
                    (menu.key === "rawdata" && page === "rawdata")
                      ? "700"
                      : "500"
                }}
              >
                {menu.label}
              </a>
            </li>
          ))}
        </ul>
        <div className="premium-navbar-actions" style={{ display: "flex", alignItems: "center", gap: "16px" }}>
          {saveState && (
            <span style={{ fontSize: "0.85rem", opacity: 0.85, color: "rgba(255, 255, 255, 0.8)", fontFamily: "'Outfit', sans-serif" }}>
              {saveState}
            </span>
          )}
          <button
            className="premium-navbar-btn active"
            onClick={forceServerSave}
            style={{ display: "flex", alignItems: "center", gap: "6px" }}
          >
            💾 저장
          </button>
        </div>
      </nav>

      {page === "dashboard" ? (
        <div className="workbook">
          {/* 2. 히어로 타이틀 */}
          <section className="premium-hero">
            <div className="hero-dot-decorator"></div>
            <div className="premium-hero-kicker">ETOOS247</div>
            <h1 className="premium-hero-title">
              <span className="premium-hero-title-line-1">MARKETING</span>
              <span className="premium-hero-title-line-2">DASHBOARD.</span>
            </h1>
          </section>

          {/* 3. 우-좌 롤링 카드 네비게이션 */}
          <section className="marquee-section">
            <h2 className="marquee-title">WHAT WE DO</h2>
            <div className="marquee-container">
              <div className="marquee-track">
                {[...marqueeCards, ...marqueeCards].map((card, idx) => (
                  <div
                    key={`${card.id || card.name}-${idx}`}
                    className="rolling-card-wrapper"
                    onClick={() => handleCardClick(card.id)}
                  >
                    <div className={`rolling-card ${card.className}`} style={{ position: "relative", overflow: "hidden" }}>
                      {card.className === "card-theme-friends" && (
                        <img
                          src="/friends-card-bg.jpg"
                          alt="247프렌즈 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-experience" && (
                        <img
                          src="/experience-card-bg.png"
                          alt="247체험단 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-sns" && (
                        <img
                          src="/sns-card-bg.png"
                          alt="SNS Icons"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-collab" && (
                        <img
                          src="/collab-card-bg.png"
                          alt="협업이벤트 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-facility" && (
                        <img
                          src="/facility-card-bg.jpg"
                          alt="지점시설영상 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-pass" && (
                        <img
                          src="/pass-card-bg.jpg"
                          alt="합격자 취합 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-news" && (
                        <img
                          src="/news-card-bg.jpg"
                          alt="언론보도 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                      {card.className === "card-theme-mentor" && (
                        <img
                          src="/mentor-card-bg.jpg"
                          alt="멘토단 및 장학생 배경"
                          style={{
                            position: "absolute",
                            inset: 0,
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            objectPosition: "center",
                            borderRadius: "inherit",
                            opacity: 1,
                            zIndex: 0,
                            pointerEvents: "none",
                          }}
                        />
                      )}
                    </div>
                    <div className="rolling-card-label">
                      {card.name}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </section>

          {/* 4. 타이포그래피 구분 섹션 */}
          <section className="typo-divider">
            <div className="typo-divider-text">
              WE MAKE <span className="outline-text">POSSIBLE</span>
            </div>
          </section>

          {/* 5. 수직 상하 롤링 쇼케이스 */}
          <section className="vertical-showcase-section">
            <div className="vertical-showcase-container">
              <div className="vertical-showcase-viewport">
                <div className="vertical-showcase-track">
                  {showcaseSlides.map((slide, index) => (
                    <div
                      key={`showcase-${index}`}
                      className={`vertical-showcase-slide ${activeSlideIndex === index ? "active" : ""}`}
                      onClick={() => {
                        if (slide.url) {
                          window.open(slide.url, "_blank");
                        }
                      }}
                    >
                      <div className={`vertical-showcase-image-fallback ${slide.className}`} style={{ position: "relative", overflow: "hidden", width: "100%", height: "100%" }}>
                        {slide.imgSrc && (
                          <img
                            src={slide.imgSrc}
                            alt={slide.name}
                            style={{
                              position: "absolute",
                              inset: 0,
                              width: "100%",
                              height: "100%",
                              objectFit: "cover",
                              objectPosition: "center",
                              pointerEvents: "none"
                            }}
                          />
                        )}
                        <div style={{ position: "absolute", inset: 0, background: "rgba(0, 0, 0, 0.4)", zIndex: 1 }} />
                      </div>
                      <div className="vertical-showcase-overlay" style={{ zIndex: 3 }}>
                        <h3 className="vertical-showcase-tag">
                          <span className="blue-hash">#</span>
                          <span className="white-text">{slide.name}</span>
                        </h3>
                        <p style={{ color: "rgba(255, 255, 255, 0.9)", fontSize: "1.1rem", marginTop: "12px", maxWidth: "600px", margin: "12px 0 0" }}>
                          {slide.desc}
                        </p>
                      </div>
                    </div>
                  ))}
                </div>
                
                {/* Prev/Next arrows inside viewport */}
                <div className="vertical-showcase-arrows">
                  <button className="vertical-showcase-arrow" onClick={prevSlide}>↑</button>
                  <button className="vertical-showcase-arrow" onClick={nextSlide}>↓</button>
                </div>
              </div>

              {/* Dot Indicators on the right */}
              <div className="vertical-showcase-dots">
                {showcaseSlides.map((_, idx) => (
                  <button
                    key={`dot-${idx}`}
                    className={`vertical-showcase-dot ${activeSlideIndex === idx ? "active" : ""}`}
                    onClick={() => setActiveSlideIndex(idx)}
                  />
                ))}
              </div>
            </div>
          </section>

          {/* 7. 상세 대시보드 영역 (OUR WORK) */}
          <div id="our-work-section" style={{ marginTop: "60px" }}>
            <h2 className="marquee-title" style={{ margin: "0 0 24px 0" }}>
              OUR WORK ({isOverviewDashboard ? "전체 현황" : selectedDashboardTab?.name})
            </h2>
            <main className="sheet-body" style={{ borderRadius: "16px", border: "1px solid rgba(0,59,255,0.1)", boxShadow: "0 10px 30px rgba(0,0,0,0.03)" }}>
            
            {/* Quick Toggle in OUR WORK */}
            <section className="sheet-panel utility-panel" style={{ border: "none", marginBottom: "20px" }}>
              <div className="program-chip-row dashboard-scope-row" style={{ padding: "8px 12px" }}>
                <button
                  className={`program-chip ${isOverviewDashboard ? "active" : ""}`}
                  onClick={() => setDashboardTabId(OVERVIEW_TAB_ID)}
                >
                  전체 현황
                </button>
                {rawTabs.map((tab) => (
                  <button
                    key={tab.id}
                    className={`program-chip ${selectedDashboardTab?.id === tab.id && !isOverviewDashboard ? "active" : ""}`}
                    onClick={() => {
                      sortMentorRowsState();
                      setDashboardTabId(tab.id);
                      setActiveTabId(tab.id);
                    }}
                  >
                    {tab.name}
                  </button>
                ))}
              </div>
            </section>



                        {isOverviewDashboard ? (
              <>
                  <section className="sheet-grid kpi-grid">
                  <article className="sheet-panel score-panel compact-score-panel">
                    <div className="panel-title-row"><h2>전체 현황 그룹 요약</h2><span className="status-pill good">SCORE</span></div>
                    <div className="score-layout">
                      <div className="score-box strong hover-score-box">
                        <span>A그룹</span>
                        <strong>{filteredOverviewBranchScoreboard.grouped["A그룹"].length}</strong>
                        <p>우수 운영 상태의 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">A그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {filteredOverviewBranchScoreboard.grouped["A그룹"].length > 0 ? filteredOverviewBranchScoreboard.grouped["A그룹"].map((branch) => <li key={`grade-a-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box hover-score-box">
                        <span>B그룹</span>
                        <strong>{filteredOverviewBranchScoreboard.grouped["B그룹"].length}</strong>
                        <p>안정적으로 운영 중인 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">B그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {filteredOverviewBranchScoreboard.grouped["B그룹"].length > 0 ? filteredOverviewBranchScoreboard.grouped["B그룹"].map((branch) => <li key={`grade-b-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box hover-score-box">
                        <span>C그룹</span>
                        <strong>{filteredOverviewBranchScoreboard.grouped["C그룹"].length}</strong>
                        <p>보완이 필요한 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">C그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {filteredOverviewBranchScoreboard.grouped["C그룹"].length > 0 ? filteredOverviewBranchScoreboard.grouped["C그룹"].map((branch) => <li key={`grade-c-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box warn hover-score-box">
                        <span>D그룹</span>
                        <strong>{filteredOverviewBranchScoreboard.grouped["D그룹"].length}</strong>
                        <p>집중 관리가 필요한 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">D그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {filteredOverviewBranchScoreboard.grouped["D그룹"].length > 0 ? filteredOverviewBranchScoreboard.grouped["D그룹"].map((branch) => <li key={`grade-d-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                    </div>
                  </article>
                </section>

                  <section className="sheet-panel overview-board-panel">
                  <div className="panel-title-row">
                    <h2>지점별 전체 현황 보드</h2>
                      <span className="note-text">일반 이벤트와 협업이벤트를 반영한 운영 점수 50%, SNS 평가 점수 50%를 합산한 그룹 보드입니다.</span>
                  </div>
                  <div className="dashboard-search-row">
                    <input
                      className="dashboard-search-input"
                      value={overviewSearch}
                      onChange={(e) => setOverviewSearch(e.target.value)}
                      placeholder="지점명 또는 권역 검색"
                    />
                    <span className="dashboard-search-note">검색 결과 {filteredOverviewBranchScoreboard.branches.length}개 지점</span>
                  </div>
                  <div className="overview-summary-strip">
                    <div className="overview-summary-card"><span>전체 지점 수</span><strong>{filteredOverviewBranchScoreboard.branches.length}</strong></div>
                    <div className="overview-summary-card"><span>평균 점수</span><strong>{filteredOverviewBranchScoreboard.avgScore}점</strong></div>
                    <div className="overview-summary-card"><span>최상위 지점</span><strong>{filteredOverviewBranchScoreboard.topBranch?.branch || "-"}</strong></div>
                    <div className="overview-summary-card warn"><span>집중 관리 지점</span><strong>{filteredOverviewBranchScoreboard.atRiskCount}</strong></div>
                  </div>
                  <div className="grade-board">
                    {["A그룹", "B그룹", "C그룹", "D그룹"].map((grade) => (
                      <section className={`grade-column ${grade === "D그룹" ? "warn" : ""}`} key={grade}>
                        <div className="grade-column-head">
                          <div>
                            <strong>{grade}</strong>
                            <span>{filteredOverviewBranchScoreboard.grouped[grade].length}개 지점</span>
                          </div>
                        </div>
                        <div className="grade-card-list">
                          {filteredOverviewBranchScoreboard.grouped[grade].length > 0 ? filteredOverviewBranchScoreboard.grouped[grade].map((branch) => (
                            <div className="grade-branch-entry" key={`${grade}-${branch.branch}`}>
                              <button
                                className={`grade-branch-trigger ${selectedOverviewBranch === branch.branch ? "active" : ""}`}
                                onClick={() => setSelectedOverviewBranch((current) => current === branch.branch ? null : branch.branch)}
                              >
                                <span>{branch.branch}</span>
                                <strong>{branch.score}점</strong>
                              </button>
                              {selectedOverviewBranch === branch.branch ? (
                                <article className="grade-branch-card">
                                  <div className="grade-branch-head">
                                    <strong>{branch.branch}</strong>
                                    <span>{branch.score}점</span>
                                  </div>
                                  <div className="grade-branch-region">{branch.region}</div>
                                  <div className="grade-metric-row">
                                    <span>운영 점수</span>
                                    <strong>{branch.operationScore}점</strong>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>SNS 점수</span>
                                    <strong>{branch.snsScore === null ? "-" : `${branch.snsScore}점`}</strong>
                                  </div>
                                  <div className="grade-metric-row grade-metric-hover">
                                    <span>참여 활성화 방안</span>
                                    <strong>{branch.activePlanCount}개</strong>
                                    <div className="metric-tooltip">
                                      <div className="metric-tooltip-title">참여한 활성화 방안</div>
                                      <ul className="metric-tooltip-list">
                                        {branch.activePlans.length > 0
                                          ? branch.activePlans.map((planName) => <li key={`${branch.branch}-${planName}`}>{planName}</li>)
                                          : <li>참여한 활성화 방안이 없습니다.</li>}
                                      </ul>
                                    </div>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>참여 횟수</span>
                                    <strong>{branch.participatedEvents}회</strong>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>참여율</span>
                                    <strong>{branch.participationRate}%</strong>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>총 참여 인원</span>
                                    <strong>{branch.totalParticipants}명</strong>
                                  </div>
                                </article>
                              ) : null}
                            </div>
                          )) : <div className="grade-empty-card">해당 그룹 지점이 없습니다.</div>}
                        </div>
                      </section>
                    ))}
                  </div>
                </section>
              </>
            ) : isSpecialDashboard ? (
              <>
                <section className="sheet-grid kpi-grid">
                  <article className="sheet-panel score-panel compact-score-panel">
                    <div className="panel-title-row"><h2>{dashboardScopeLabel} 핵심 지표</h2><span className="status-pill good">SNS</span></div>
                    <div className="score-layout">
                      <div className="score-box strong"><span>평가 지점 수</span><strong>{snsSummary.totalBranches}</strong><p>실제 진단 데이터가 입력된 지점 수입니다.</p></div>
                      <div className="score-box"><span>평균 최종 점수</span><strong>{snsSummary.averageScore}</strong><p>엑셀 평가결과 로직을 그대로 적용한 평균 점수입니다.</p></div>
                      <div className="score-box hover-score-box"><span>A등급 지점</span><strong>{snsGradeGroups.A.length}</strong><p>80점 이상 우수 지점입니다.</p><div className="score-tooltip"><div className="score-tooltip-title">A등급 지점명</div><ul className="score-tooltip-list">{snsGradeGroups.A.length > 0 ? snsGradeGroups.A.map((row) => <li key={`sns-a-${row.branch}`}>{row.branch}</li>) : <li>해당 지점이 없습니다.</li>}</ul></div></div>
                      <div className="score-box warn hover-score-box"><span>D등급 지점</span><strong>{snsGradeGroups.D.length}</strong><p>40점 미만 집중 관리 지점입니다.</p><div className="score-tooltip"><div className="score-tooltip-title">D등급 지점명</div><ul className="score-tooltip-list">{snsGradeGroups.D.length > 0 ? snsGradeGroups.D.map((row) => <li key={`sns-d-${row.branch}`}>{row.branch}</li>) : <li>해당 지점이 없습니다.</li>}</ul></div></div>
                    </div>
                  </article>
                </section>

                <section className="sheet-panel">
                  <div className="panel-title-row"><h2>SNS 등급 분포</h2><span className="note-text">엑셀 평가결과 기준일 {snsEvaluationBaseDate}</span></div>
                  <div className="dashboard-search-row dashboard-search-row-tight">
                    <input
                      className="dashboard-search-input"
                      value={snsSearch}
                      onChange={(e) => setSnsSearch(e.target.value)}
                      placeholder="지점명 검색"
                    />
                    <span className="dashboard-search-note">검색 결과 {filteredSnsDashboardRows.length}개 지점</span>
                  </div>
                  <div className="grade-board sns-grade-board">
                    {["A", "B", "C", "D"].map((grade) => (
                      <section className={`grade-column ${grade === "D" ? "warn" : ""}`} key={grade}>
                        <div className="grade-column-head">
                          <div>
                            <strong>{grade}등급</strong>
                            <span>{snsGradeGroups[grade].length}개 지점</span>
                          </div>
                        </div>
                          <div className="grade-card-list">
                           {snsGradeGroups[grade].length > 0 ? snsGradeGroups[grade].map((row) => (
                              <article className="grade-branch-card" key={`sns-grade-${grade}-${row.branch}`}>
                                <div className="grade-branch-head"><strong>{row.branch}</strong><span>{row.finalScore}점</span></div>
                                <div className="grade-metric-row"><span>블로그 점수</span><ExternalScoreLink href={row.hasBlog ? row.blogUrl : ""} value={row.blogScore} /></div>
                                <div className="grade-metric-row"><span>인스타 점수</span><ExternalScoreLink href={row.hasInstagram ? row.instagramUrl : ""} value={row.instagramScore} /></div>
                              </article>
                          )) : <div className="grade-empty-card">해당 등급 지점이 없습니다.</div>}
                        </div>
                      </section>
                    ))}
                  </div>
                </section>

                <section className="sheet-grid detail-grid">
                  <article className="sheet-panel">
                    <div className="panel-title-row"><h2>상위 지점</h2><span className="note-text">최종 점수 기준 TOP 8</span></div>
                    <ul className="action-list">
                      {snsSummary.topBranches.map((row, index) => (
                        <li className="action-item" key={`top-${row.branch}`}>
                          <div className="action-index">{index + 1}</div>
                          <div className="action-copy"><strong>{row.branch} / {row.grade}등급 / {row.finalScore}점</strong><p>블로그 {row.blogScore}점, 인스타 {row.instagramScore}점</p></div>
                        </li>
                      ))}
                    </ul>
                  </article>
                  <article className="sheet-panel">
                    <div className="panel-title-row"><h2>관리 필요 지점</h2><span className="note-text">최종 점수 하위 8개 지점</span></div>
                    <ul className="action-list">
                      {snsSummary.lowBranches.map((row, index) => (
                        <li className="action-item" key={`low-${row.branch}`}>
                          <div className="action-index">{index + 1}</div>
                          <div className="action-copy"><strong>{row.branch} / {row.grade}등급 / {row.finalScore}점</strong><p>{row.hasBlog ? "블로그 운영" : "블로그 미운영"}, {row.hasInstagram ? "인스타 운영" : "인스타 미운영"}</p></div>
                        </li>
                      ))}
                    </ul>
                  </article>
                </section>

                <section className="sheet-panel">
                  <div className="panel-title-row"><h2>SNS 평가 상세표</h2><span className="note-text">평가결과 시트 계산식 기준</span></div>
                  <div className="table-shell">
                    <table className="excel-table compact-table">
                      <thead>
                        <tr>
                          <th>지점</th>
                          <th>블로그 점수</th>
                          <th>인스타 점수</th>
                          <th>최종 점수</th>
                          <th>등급</th>
                          <th>블로그 운영</th>
                          <th>인스타 운영</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredSnsDashboardRows.map((row) => (
                          <tr key={`sns-table-${row.branch}`}>
                            <td>{row.branch}</td>
                            <td>{row.blogScore}</td>
                            <td>{row.instagramScore}</td>
                            <td>{row.finalScore}</td>
                            <td>{row.grade}</td>
                            <td>{row.hasBlog ? "운영" : "미운영"}</td>
                            <td>{row.hasInstagram ? "운영" : "미운영"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              </>
            ) : isFacilityDashboard ? (
              <>
                <section className="sheet-grid kpi-grid">
                  <article className="sheet-panel score-panel compact-score-panel">
                    <div className="panel-title-row"><h2>{dashboardScopeLabel} 핵심 지표</h2><span className="status-pill good">VIDEO</span></div>
                    <div className="score-layout">
                      <div className="score-box strong"><span>등록 URL 수</span><strong>{facilityDashboardSummary.totalUrls}</strong><p>현재 탭에 연결된 시설영상 URL 수입니다.</p></div>
                      <div className="score-box hover-score-box">
                        <span>참여 지점 수</span>
                        <strong>{facilityDashboardSummary.activeBranches}</strong>
                        <p>시설영상 URL이 등록된 지점 수입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">참여 지점명</div>
                          <ul className="score-tooltip-list">
                            {facilityActiveBranchTooltip.length > 0 ? facilityActiveBranchTooltip.map((branch) => <li key={`facility-active-${branch}`}>{branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box warn hover-score-box">
                        <span>미참여 지점 수</span>
                        <strong>{facilityDashboardSummary.inactiveBranches}</strong>
                        <p>시설영상 URL이 아직 없는 지점 수입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">미참여 지점명</div>
                          <ul className="score-tooltip-list">
                            {facilityInactiveBranchTooltip.length > 0 ? facilityInactiveBranchTooltip.map((branch) => <li key={`facility-inactive-${branch}`}>{branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                    </div>
                  </article>
                </section>

                <section className="sheet-panel">
                  <div className="panel-title-row"><h2>권역별 시설영상 현황</h2><span className="note-text">지점명을 누르면 영상이 대시보드 내 팝업으로 재생됩니다.</span></div>
                  <div className="facility-region-board">
                    {facilityRegionGroups.length > 0 ? facilityRegionGroups.map((group) => (
                      <section className="facility-region-card" key={`facility-region-${group.region}`}>
                        <div className="facility-region-head">
                          <strong>{group.region}</strong>
                          <span>{group.branches.length}개 지점</span>
                        </div>
                        <div className="facility-branch-grid">
                          {group.branches.map((row) =>
                            row.url ? (
                              <a
                                key={`facility-branch-${row.branch}`}
                                className="facility-branch-link"
                                href={row.url}
                                onClick={(e) => {
                                  e.preventDefault();
                                  setActiveVideoUrl(row.url);
                                  setActiveVideoBranch(row.branch);
                                }}
                                title={`${row.branch} 시설영상 재생`}
                              >
                                {row.branch}
                              </a>
                            ) : (
                              <div key={`facility-branch-${row.branch}`} className="facility-branch-link disabled" title="시설영상 URL 미등록">
                                {row.branch}
                              </div>
                            )
                          )}
                        </div>
                      </section>
                    )) : (
                      <div className="grade-empty-card">등록된 시설영상 URL이 없습니다.</div>
                    )}
                  </div>
                </section>
              </>
            ) : isCollabDashboard ? (
              <>
                <section className="sheet-grid kpi-grid">
                    <article className="sheet-panel score-panel compact-score-panel">
                      <div className="panel-title-row"><h2>{dashboardScopeLabel} 핵심 지표</h2><span className="status-pill good">KPI</span></div>
                      <div className="score-layout">
                        <div className="score-box strong"><span>진행 횟수</span><strong>{collabDashboardSummary.uniqueEvents}</strong><p>현재 탭의 전체 협업 이벤트 수입니다.</p></div>
                        <div className="score-box hover-score-box">
                          <span>참여 지점 수</span>
                          <strong>{collabDashboardSummary.activeBranches}</strong>
                          <p>URL이 1건 이상 등록된 지점 수입니다.</p>
                          <div className="score-tooltip">
                            <div className="score-tooltip-title">참여 지점명</div>
                            <ul className="score-tooltip-list">
                              {collabActiveBranchTooltip.length > 0 ? collabActiveBranchTooltip.map((branch) => <li key={`collab-active-${branch}`}>{branch}</li>) : <li>해당 지점이 없습니다.</li>}
                            </ul>
                          </div>
                        </div>
                        <div className="score-box warn hover-score-box">
                          <span>미참여 지점 수</span>
                          <strong>{collabDashboardSummary.inactiveBranches}</strong>
                          <p>아직 URL 등록 이력이 없는 지점입니다.</p>
                          <div className="score-tooltip">
                            <div className="score-tooltip-title">미참여 지점명</div>
                            <ul className="score-tooltip-list">
                              {collabInactiveBranchTooltip.length > 0 ? collabInactiveBranchTooltip.map((branch) => <li key={`collab-inactive-${branch}`}>{branch}</li>) : <li>해당 지점이 없습니다.</li>}
                            </ul>
                          </div>
                        </div>
                      </div>
                    </article>
                  </section>

                <section className="sheet-panel">
                  <div className="panel-title-row"><h2>협업 URL 현황</h2><span className="note-text">{selectedCollabBranch ? `${selectedCollabBranch} 기준` : "전체 지점 기준"}</span></div>
                  <div className="branch-selector-row">
                    <button className={`branch-chip ${selectedCollabBranch === null ? "active" : ""}`} onClick={() => setSelectedCollabBranch(null)}>전체 보기</button>
                    <input className="branch-search-input" placeholder="지점명 검색" value={branchKeyword} onChange={(e) => setBranchKeyword(e.target.value)} />
                    <button className="branch-toggle-button" onClick={() => setAreRegionsExpanded((current) => !current)}>
                      {areRegionsExpanded ? "권역 전체 접기" : "권역 전체 펼치기"}
                    </button>
                  </div>
                  {shouldShowBranchGroups ? (
                    visibleCollabBranchGroups.length > 0 ? (
                      <div className="branch-group-list">
                        {visibleCollabBranchGroups.map((group) => (
                          <div className="branch-group" key={`collab-${group.region}`}>
                            <div className="branch-group-title">{group.region}</div>
                            <div className="branch-group-chips">
                              {group.branches.map((branch) => (
                                <button
                                  key={`collab-branch-${branch}`}
                                  className={`branch-chip ${selectedCollabBranch === branch ? "active" : ""}`}
                                  onClick={() => setSelectedCollabBranch(branch)}
                                >
                                  {branch}
                                </button>
                              ))}
                            </div>
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="branch-collapsed-hint">검색 결과에 맞는 지점이 없습니다.</div>
                    )
                  ) : (
                    <div className="branch-collapsed-hint">권역 전체 펼치기를 누르거나 지점명을 검색하면 권역별 지점 목록이 열립니다.</div>
                  )}
                  <div className="event-analytics-panel collab-analytics-panel">
                    <div className="event-chart-card">
                      <div className="event-chart-head">
                        <strong>{selectedCollabBranch ? `${selectedCollabBranch} 진행 이벤트` : `${dashboardScopeLabel} 진행 이벤트`}</strong>
                        <span>이벤트명을 클릭하면 URL이 열립니다.</span>
                      </div>
                      <div className="collab-event-list">
                        {collabEventList.length > 0 ? collabEventList.map((event) => (
                          <button
                            key={`collab-event-${event.id}`}
                            className={`collab-event-item ${selectedCollabEvent === event.label ? "active" : ""}`}
                            onClick={() => setSelectedCollabEvent(event.label)}
                          >
                            <span>{event.label}</span>
                            <strong>{selectedCollabBranch ? `${event.urlCount}건` : `${event.branchCount}지점`}</strong>
                          </button>
                        )) : <div className="grade-empty-card">등록된 협업 이벤트가 없습니다.</div>}
                      </div>
                    </div>
                      <div className="event-summary-card">
                        <h3>현재 보기 요약</h3>
                        <ul className="metric-list">
                        <li><span>기준</span><strong>{selectedCollabBranch || "전체 지점"}</strong></li>
                        <li><span>진행 횟수</span><strong>{collabEventList.length}</strong></li>
                        <li><span>등록 URL 수</span><strong>{selectedCollabBranchRow ? selectedCollabBranchRow.urlCount : collabDashboardSummary.totalUrls}</strong></li>
                        <li><span>선택 이벤트</span><strong>{selectedCollabEvent || "-"}</strong></li>
                        </ul>
                          <div className="collab-url-panel">
                            <div className="metric-tooltip-title">
                              {selectedCollabEvent
                                ? `${selectedCollabEvent} 채널별 현황`
                                : "이벤트를 선택하세요"}
                            </div>
                            {selectedCollabEventData ? (
                              <div className="collab-channel-groups">
                                {["홈페이지", "블로그", "인스타/언론기사"].map((channel) => {
                                  const channelItems = selectedCollabEventData.channels?.[channel] || [];
                                  return (
                                    <section className="collab-channel-group" key={`${selectedCollabEvent}-${channel}`}>
                                      <div className="collab-channel-head">
                                        <strong>{channel}</strong>
                                        <span>{channelItems.length}개</span>
                                      </div>
                                      <div className="collab-channel-branch-grid">
                                        {channelItems.length > 0 ? channelItems.map((item) => (
                                          <a
                                            key={item.id || `${channel}-${item.branch}-${item.url}`}
                                            className="collab-branch-link"
                                            href={item.url}
                                            target="_blank"
                                            rel="noreferrer"
                                            title={`${item.branch} ${channel} 열기`}
                                          >
                                            {item.branch}
                                          </a>
                                        )) : <div className="collab-channel-empty">등록된 지점이 없습니다.</div>}
                                      </div>
                                    </section>
                                  );
                                })}
                              </div>
                            ) : (
                              <div className="branch-collapsed-hint small">
                                이벤트명을 누르면 채널별 참여 지점이 표시됩니다.
                              </div>
                            )}
                          </div>
                    </div>
                  </div>
                </section>
              </>
            ) : isMentorDashboard ? (
              <>
                <section className="sheet-grid kpi-grid">
                  <article className="sheet-panel score-panel compact-score-panel">
                    <div className="panel-title-row">
                      <h2>{dashboardScopeLabel} 핵심 지표</h2>
                      <span className="status-pill good" style={{ background: "rgba(59, 130, 246, 0.15)", color: "#3b82f6", border: "1px solid rgba(59, 130, 246, 0.3)" }}>MENTOR</span>
                    </div>
                    <div className="score-layout">
                      <div className="score-box strong">
                        <span>전체 인원</span>
                        <strong>{mentorStats.total}명</strong>
                        <p>멘토단 및 장학생 전체 등록 인원입니다.</p>
                      </div>
                      <div className="score-box">
                        <span>멘토단</span>
                        <strong>{mentorStats.mentors}명</strong>
                        <p>멘토로 임명된 장학생 인원입니다.</p>
                      </div>
                      <div className="score-box">
                        <span>총 1억 장학생</span>
                        <strong>{mentorStats.scholars}명</strong>
                        <p>멘토단이 아닌 총 1억 장학생 인원입니다.</p>
                      </div>
                      <div className="score-box primary-metric" style={{ background: "linear-gradient(135deg, rgba(59, 130, 246, 0.1) 0%, rgba(20, 184, 166, 0.1) 100%)", border: "1px solid rgba(59, 130, 246, 0.2)" }}>
                        <span style={{ color: "var(--metric-accent-color)" }}>총 지급 장학금액</span>
                        <strong style={{ color: "var(--metric-accent-color)" }}>{mentorStats.amountSum.toLocaleString()}원</strong>
                        <p>등록된 전체 장학금 지급 합계액입니다.</p>
                      </div>
                    </div>
                  </article>
                </section>

                <section className="sheet-panel">
                  <div className="panel-title-row">
                    <h2>멘토단 및 장학생 상세 조회</h2>
                    <span className="note-text">이름, 대학, 학과, 메모 등으로 검색하거나 필터를 적용하세요.</span>
                  </div>

                  <div className="dashboard-search-row dashboard-search-row-tight" style={{ display: "flex", gap: "10px", alignItems: "center", flexWrap: "wrap", marginBottom: "20px" }}>
                    <input
                      className="dashboard-search-input"
                      value={mentorSearch}
                      onChange={(e) => setMentorSearch(e.target.value)}
                      placeholder="이름, 연도, 대학교, 학과, 메모 등 검색..."
                      style={{ flex: "1", minWidth: "200px" }}
                    />
                    
                    <div style={{ display: "flex", gap: "10px" }}>
                      <select
                        className="branch-selector-dropdown"
                        value={mentorBranchFilter}
                        onChange={(e) => setMentorBranchFilter(e.target.value)}
                        style={{ padding: "8px 12px", borderRadius: "8px", border: "1px solid var(--border-color)", background: "var(--panel-bg)", color: "var(--text-color)" }}
                      >
                        <option value="all">전체 지점</option>
                        {mentorBranchOptions.filter(o => o !== "all").map((branch) => (
                          <option key={branch} value={branch}>{branch}</option>
                        ))}
                      </select>

                      <select
                        className="branch-selector-dropdown"
                        value={mentorUnivFilter}
                        onChange={(e) => setMentorUnivFilter(e.target.value)}
                        style={{ padding: "8px 12px", borderRadius: "8px", border: "1px solid var(--border-color)", background: "var(--panel-bg)", color: "var(--text-color)" }}
                      >
                        <option value="all">전체 대학교</option>
                        {mentorUnivOptions.filter(o => o !== "all").map((univ) => (
                          <option key={univ} value={univ}>{univ}</option>
                        ))}
                      </select>
                    </div>

                    <span className="dashboard-search-note" style={{ marginLeft: "auto" }}>
                      검색 결과 {filteredMentorRows.length}명 (멘토단 {mentorsList.length}명 / 장학생 {scholarsList.length}명)
                    </span>
                  </div>

                  {/* 멘토단 리스트 (상단) */}
                  <div className="mentor-section-container" style={{ marginBottom: "30px" }}>
                    <h3 style={{ fontSize: "1.2rem", fontWeight: "600", marginBottom: "12px", display: "flex", alignItems: "center", gap: "8px", color: "var(--metric-accent-color)" }}>
                      <span style={{ display: "inline-block", width: "8px", height: "8px", borderRadius: "50%", background: "var(--metric-accent-color)" }}></span>
                      멘토단 리스트 ({mentorsList.length}명)
                    </h3>
                    {mentorsList.length > 0 ? (
                      <div className="table-shell">
                        <table className="excel-table compact-table">
                          <thead>
                            <tr>
                              <th>연도</th>
                              <th>이름</th>
                              <th>합격 대학</th>
                              <th>학과</th>
                              <th>지점</th>
                              <th>장학 그룹</th>
                              <th>장학 금액</th>
                            </tr>
                          </thead>
                          <tbody>
                            {mentorsList.map((row) => (
                              <tr key={row.id}>
                                <td>{row.year}</td>
                                <td style={{ fontWeight: "600" }}>{row.name}</td>
                                <td>{row.university}</td>
                                <td>{row.department}</td>
                                <td><span className="status-pill good">{row.branch}</span></td>
                                <td>{row.group}</td>
                                <td style={{ textAlign: "right" }}>{Number(row.amount).toLocaleString()}원</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <div className="grade-empty-card" style={{ padding: "20px", textAlign: "center", background: "rgba(255, 255, 255, 0.02)", border: "1px dashed var(--border-color)", borderRadius: "8px" }}>
                        조건에 부합하는 멘토단 학생이 없습니다.
                      </div>
                    )}
                  </div>

                  {/* 장학생 리스트 (하단) */}
                  <div className="scholar-section-container">
                    <h3 style={{ fontSize: "1.2rem", fontWeight: "600", marginBottom: "12px", display: "flex", alignItems: "center", gap: "8px", color: "var(--text-color)" }}>
                      <span style={{ display: "inline-block", width: "8px", height: "8px", borderRadius: "50%", background: "var(--text-color)" }}></span>
                      장학생 리스트 ({scholarsList.length}명)
                    </h3>
                    {scholarsList.length > 0 ? (
                      <div className="table-shell">
                        <table className="excel-table compact-table">
                          <thead>
                            <tr>
                              <th>연도</th>
                              <th>이름</th>
                              <th>합격 대학</th>
                              <th>학과</th>
                              <th>지점</th>
                              <th>장학 그룹</th>
                              <th>장학 금액</th>
                            </tr>
                          </thead>
                          <tbody>
                            {scholarsList.map((row) => (
                              <tr key={row.id}>
                                <td>{row.year}</td>
                                <td>{row.name}</td>
                                <td>{row.university}</td>
                                <td>{row.department}</td>
                                <td><span className="status-pill good">{row.branch}</span></td>
                                <td>{row.group}</td>
                                <td style={{ textAlign: "right" }}>{Number(row.amount).toLocaleString()}원</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    ) : (
                      <div className="grade-empty-card" style={{ padding: "20px", textAlign: "center", background: "rgba(255, 255, 255, 0.02)", border: "1px dashed var(--border-color)", borderRadius: "8px" }}>
                        조건에 부합하는 장학생 학생이 없습니다.
                      </div>
                    )}
                  </div>
                </section>
              </>
            ) : (
              <>
                <section className="sheet-grid kpi-grid">
                  <article className="sheet-panel score-panel compact-score-panel">
                    <div className="panel-title-row"><h2>{dashboardScopeLabel} 핵심 지표</h2><span className="status-pill good">KPI</span></div>
                    <div className="score-layout">
                      <div className="score-box strong"><span>진행 횟수</span><strong>{dashboardTab?.events.length || 0}</strong><p>현재 탭의 전체 이벤트 수입니다.</p></div>
                      <div className="score-box hover-score-box">
                        <span>참여 지점 수</span>
                        <strong>{scopedSummary.activeBranches}</strong>
                        <p>한 번 이상 참여한 지점 수입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">참여 지점명</div>
                          <ul className="score-tooltip-list">
                            {activeBranchTooltip.length > 0 ? activeBranchTooltip.map((branch) => <li key={`active-${branch}`}>{branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box warn hover-score-box">
                        <span>미참여 지점 수</span>
                        <strong>{scopedSummary.inactiveBranches}</strong>
                        <p>아직 참여 이력이 없는 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">미참여 지점명</div>
                          <ul className="score-tooltip-list">
                            {inactiveBranchTooltip.length > 0 ? inactiveBranchTooltip.map((branch) => <li key={`inactive-${branch}`}>{branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                    </div>
                  </article>
                </section>

                <section className="sheet-panel">
                  <div className="panel-title-row">
                    <h2>{selectedBranch ? `${selectedBranch} 참여 현황` : "이벤트별 전체 참여 인원 현황"}</h2>
                    <span className="note-text">{selectedBranch ? "선택한 지점 기준" : "전체 지점 기준"}</span>
                  </div>
                  <div className="branch-selector-row">
                    <button className={`branch-chip ${selectedBranch === null ? "active" : ""}`} onClick={() => setSelectedBranch(null)}>전체 보기</button>
                    <input
                      className="branch-search-input"
                      value={branchKeyword}
                      onChange={(e) => setBranchKeyword(e.target.value)}
                      placeholder="지점명 검색"
                    />
                    <button className="branch-toggle-button" onClick={() => setAreRegionsExpanded((current) => !current)}>
                      {areRegionsExpanded ? "권역 전체 접기" : "권역 전체 펼치기"}
                    </button>
                  </div>
                  {shouldShowBranchGroups ? (
                    visibleBranchGroups.length > 0 ? (
                      <div className="branch-group-list">
                        {visibleBranchGroups.map((group) => (
                          <section className="branch-group" key={group.region}>
                            <div className="branch-group-title">{group.region}</div>
                            <div className="branch-group-chips">
                              {group.branches.map((branch) => (
                                <button
                                  key={branch}
                                  className={`branch-chip ${selectedBranch === branch ? "active" : ""}`}
                                  onClick={() => setSelectedBranch(branch)}
                                >
                                  {branch}
                                </button>
                              ))}
                            </div>
                          </section>
                        ))}
                      </div>
                    ) : (
                      <div className="branch-collapsed-hint">검색 결과에 맞는 지점이 없습니다.</div>
                    )
                  ) : (
                    <div className="branch-collapsed-hint">권역 전체 펼치기를 누르거나 지점명을 검색하면 권역별 지점 목록이 열립니다.</div>
                  )}
                  <div className="event-analytics-panel">
                    <div className="event-chart-card">
                      <div className="event-chart-head">
                        <strong>{selectedBranch ? `${selectedBranch} 이벤트별 참여` : `${dashboardScopeLabel} 이벤트별 참여`}</strong>
                        <span>세로 막대 그래프</span>
                      </div>
                      <div className="event-combo-chart">
                          <div
                            className="event-bars"
                            style={{ gridTemplateColumns: `repeat(${Math.max(branchChartData.length, 1)}, minmax(0, 1fr))` }}
                          >
                            {branchChartData.map((item) => (
                              <button
                                type="button"
                                className={`event-bar-item ${selectedChartEventId === item.id ? "active" : ""}`}
                                key={item.id}
                                onClick={() => setSelectedChartEventId((current) => current === item.id ? null : item.id)}
                                title={`${item.label} · ${item.schedule}`}
                              >
                                <div className="event-bar-value">{item.participants}</div>
                                <div className="event-bar-track">
                                  <div
                                    className="event-bar-fill"
                                    style={{ height: `${Math.max((item.participants / maxChartParticipants) * 100, item.participants > 0 ? 10 : 0)}%` }}
                                  />
                                </div>
                                <div className="event-bar-label">{item.label}</div>
                              </button>
                            ))}
                          </div>
                        </div>
                      </div>
                      <div className="event-summary-card">
                        <h3>현재 보기 요약</h3>
                        <ul className="metric-list">
                          {selectedChartEvent ? (
                            <>
                              <li><span>이벤트명</span><strong>{selectedChartEvent.label}</strong></li>
                              <li><span>진행 일정</span><strong>{selectedChartEvent.schedule}</strong></li>
                              <li><span>참여율</span><strong>{selectedChartEvent.participationRate}%</strong></li>
                              <li><span>참여 인원</span><strong>{selectedChartEvent.participants}명</strong></li>
                              <li className="metric-list-stacked">
                                <div className="metric-stack-head">
                                  <span>참여 지점</span>
                                  <strong>{selectedChartEvent.branchCount}개</strong>
                                </div>
                                <div className="metric-chip-list">
                                  {selectedChartEvent.participatingBranches.length > 0
                                    ? selectedChartEvent.participatingBranches.map((branch) => (
                                        <span className="metric-chip" key={`${selectedChartEvent.id}-${branch}`}>{branch}</span>
                                      ))
                                    : <span className="metric-empty-copy">참여 지점이 없습니다.</span>}
                                </div>
                              </li>
                            </>
                          ) : (
                            <>
                              <li><span>기준</span><strong>{selectedBranch || "전체 지점"}</strong></li>
                              <li className="metric-hover-item">
                                <span>참여 횟수</span>
                                <strong>{participatedEventCount}</strong>
                                <div className="metric-tooltip">
                                  <div className="metric-tooltip-title">참여한 이벤트명</div>
                                  <ul className="metric-tooltip-list">
                                    {participatedEventLabels.length > 0 ? participatedEventLabels.map((label) => <li key={label}>{label}</li>) : <li>참여한 이벤트가 없습니다.</li>}
                                  </ul>
                                </div>
                              </li>
                              <li><span>총 참여 인원</span><strong>{branchChartData.reduce((sum, item) => sum + item.participants, 0)}</strong></li>
                              <li><span>참여율</span><strong>{participationRate}%</strong></li>
                              <li><span>최다 참여 이벤트</span><strong>{[...branchChartData].sort((a, b) => b.participants - a.participants)[0]?.label || "-"}</strong></li>
                            </>
                          )}
                        </ul>
                      </div>
                    </div>
                </section>
              </>
            )}
          </main>
          </div>
        </div>
            ) : (page === "competitors" || page === "sns") ? (
              <div className="workbook" style={{ marginTop: "40px" }}>
                <div className="competitor-view-container" style={{ padding: "0 10px", width: "100%", maxWidth: "1200px", margin: "0 auto" }}>
                  
                  {/* 1. Top Search Section */}
                  <div style={{ position: "relative", width: "100%", maxWidth: "600px", margin: "0 auto 30px auto", zIndex: 100 }}>
                    <div style={{ display: "flex", alignItems: "center", background: "#ffffff", border: "2px solid var(--primary-blue)", borderRadius: "16px", padding: "10px 20px", boxShadow: "0 10px 15px -3px rgba(59, 130, 246, 0.08)" }}>
                      <span style={{ fontSize: "1.2rem", marginRight: "12px" }}>🔍</span>
                      <input
                        type="text"
                        placeholder="분석할 지점명을 입력하여 검색하십시오 (예: 목동, 인천송도, 도봉, 대전둔산...)"
                        value={searchQuery}
                        onChange={(e) => setSearchQuery(e.target.value)}
                        onKeyDown={(e) => {
                          if (e.key === "Enter") {
                            const matchedBranch = allBranches.find(b => b.toLowerCase() === searchQuery.toLowerCase().trim());
                            if (matchedBranch) {
                              setSelectedBranch(matchedBranch);
                              setSearchQuery(matchedBranch);
                              setTypedAiStrategy("");
                              e.currentTarget.blur();
                            }
                          }
                        }}
                        style={{ width: "100%", border: "none", outline: "none", fontSize: "0.98rem", color: "#0f172a", fontWeight: "500" }}
                      />
                      {searchQuery && (
                        <button
                          onClick={() => {
                            setSearchQuery("");
                            setSelectedBranch(null);
                          }}
                          style={{ border: "none", background: "transparent", cursor: "pointer", color: "#94a3b8", fontSize: "1.1rem", padding: "0 4px" }}
                        >
                          ✕
                        </button>
                      )}
                    </div>

                    {/* Autocomplete Dropdown list */}
                    {searchQuery && searchQuery !== selectedBranch && (
                      (() => {
                        const filtered = allBranches.filter(b => b.toLowerCase().includes(searchQuery.toLowerCase()));
                        return filtered.length > 0 ? (
                          <div style={{ position: "absolute", top: "100%", left: 0, right: 0, background: "#ffffff", border: "1px solid #cbd5e1", borderRadius: "16px", marginTop: "8px", boxShadow: "0 20px 25px -5px rgba(0,0,0,0.1)", maxHeight: "250px", overflowY: "auto", zIndex: 9999 }}>
                            {filtered.map(b => (
                              <div
                                key={`search-suggest-${b}`}
                                onClick={() => {
                                  setSearchQuery(b);
                                  setSelectedBranch(b);
                                  setTypedAiStrategy("");
                                }}
                                style={{ padding: "12px 20px", cursor: "pointer", borderBottom: "1px solid #f1f5f9", color: "#334155", fontSize: "0.92rem", transition: "background 0.2s" }}
                                onMouseEnter={(e) => e.target.style.background = "#f8fafc"}
                                onMouseLeave={(e) => e.target.style.background = "transparent"}
                              >
                                📍 {b} 지점
                              </div>
                            ))}
                          </div>
                        ) : null;
                      })()
                    )}
                  </div>

                  {/* 2. Main content view */}
                  {selectedBranch ? (() => {
                    const statusInfo = getBranchMarketingStatus(selectedBranch, rawTabs);
                    const competitors = getCompetitorsForBranch(selectedBranch);
                    const lastCrawled = crawledTimestamps[selectedBranch] || "N/A";
                    const strategy = aiStrategies[selectedBranch] || "";

                    return (
                      <div className="competitor-column-layout" style={{ display: "flex", flexDirection: "column", gap: "24px", width: "100%" }}>
                        
                        {/* 0. 지점명 SNS 분석 보고서 헤더 */}
                        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: "8px" }}>
                          <h2 className="marquee-title" style={{ margin: 0 }}>
                            {selectedBranch}점 {page === "competitors" ? "경쟁사 비교 보고서" : "SNS 분석 보고서"}
                          </h2>
                          <div
                            style={{ position: "relative" }}
                            onMouseEnter={() => setIsSnsScoreTooltipOpen(true)}
                            onMouseLeave={() => setIsSnsScoreTooltipOpen(false)}
                          >
                            <span
                              style={{
                                fontSize: "0.85rem",
                                color: "#2563eb",
                                fontWeight: "700",
                                background: "#eff6ff",
                                padding: "6px 14px",
                                borderRadius: "9999px",
                                border: "1px solid #bfdbfe",
                                cursor: "help",
                                display: "inline-flex",
                                alignItems: "center",
                                gap: "6px",
                                transition: "all 0.2s",
                                boxShadow: isSnsScoreTooltipOpen ? "0 4px 12px rgba(37, 99, 235, 0.15)" : "none"
                              }}
                            >
                              <span>실시간 분석 기준</span>
                              <span style={{ fontSize: "0.78rem", opacity: 0.8 }}>ℹ️</span>
                            </span>

                            {isSnsScoreTooltipOpen && (
                              <div
                                style={{
                                  position: "absolute",
                                  right: 0,
                                  top: "calc(100% + 10px)",
                                  width: "390px",
                                  background: "#ffffff",
                                  borderRadius: "20px",
                                  border: "1px solid #cbd5e1",
                                  boxShadow: "0 25px 50px -12px rgba(15, 23, 42, 0.25), 0 0 0 1px rgba(37, 99, 235, 0.1)",
                                  padding: "22px 24px",
                                  zIndex: 9999,
                                  textAlign: "left",
                                  animation: "fadeIn 0.2s ease-out"
                                }}
                              >
                                <div style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "14px", borderBottom: "1px solid #f1f5f9", paddingBottom: "10px" }}>
                                  <span style={{ fontSize: "1.15rem" }}>📐</span>
                                  <strong style={{ fontSize: "0.96rem", color: "#0f172a", fontWeight: "800", fontFamily: "'Outfit', 'Pretendard', sans-serif" }}>
                                    종합 SNS 점수 산출 로직 & 평가 기준
                                  </strong>
                                </div>

                                <div style={{ display: "flex", flexDirection: "column", gap: "10px", fontSize: "0.81rem", color: "#475569", lineHeight: "1.55" }}>
                                  {/* 1. 네이버 블로그 */}
                                  <div style={{ background: "#f0fdf4", padding: "10px 12px", borderRadius: "12px", border: "1px solid #bbf7d0" }}>
                                    <div style={{ fontWeight: "800", color: "#15803d", marginBottom: "4px" }}>
                                      🟢 네이버 블로그 (50점 만점)
                                    </div>
                                    <div>• <strong>활동량 (30점)</strong>: 최근 30일 글 수 (8개↑ 30점, 5~7개 25점, 3~4개 20점, 1~2개 10점)</div>
                                    <div>• <strong>반응 수준 (5점)</strong>: 최근 5개 글 댓글·공감 수치 평가 (0~5점)</div>
                                    <div style={{ color: "#166534", fontSize: "0.74rem", marginTop: "2px", fontWeight: "600" }}>※ 환산식: (활동량 + 반응수준) ÷ 35 × 50점 (60일 미발행 시 -15점 감점)</div>
                                  </div>

                                  {/* 2. 인스타그램 */}
                                  <div style={{ background: "#fdf2f8", padding: "10px 12px", borderRadius: "12px", border: "1px solid #fbcfe8" }}>
                                    <div style={{ fontWeight: "800", color: "#be185d", marginBottom: "4px" }}>
                                      🌸 인스타그램 (50점 만점)
                                    </div>
                                    <div>• <strong>활동량 (30점)</strong>: 최근 30일 피드 (12개↑ 30점, 8~11개 25점, 4~7개 20점, 1~3개 10점)</div>
                                    <div>• <strong>운영 품질 (25점)</strong>: 디자인(5점) + 반응(5점) + 프로필/기능/CTA/링크/브랜드(각 3점)</div>
                                    <div style={{ color: "#9d174d", fontSize: "0.74rem", marginTop: "2px", fontWeight: "600" }}>※ 환산식: (활동량 + 운영품질) ÷ 55 × 50점 (30일 미발행 시 -5점 감점)</div>
                                  </div>

                                  {/* 3. 종합 등급 */}
                                  <div style={{ background: "#eff6ff", padding: "10px 12px", borderRadius: "12px", border: "1px solid #bfdbfe" }}>
                                    <div style={{ fontWeight: "800", color: "#1d4ed8", marginBottom: "4px" }}>
                                      🏆 종합 SNS 총점 (100점) & 등급
                                    </div>
                                    <div>• <strong>총점</strong> = 블로그 점수(50점) + 인스타그램 점수(50점) - 미운영 감점</div>
                                    <div style={{ marginTop: "4px", display: "flex", gap: "6px", flexWrap: "wrap", fontWeight: "700", fontSize: "0.76rem" }}>
                                      <span style={{ color: "#15803d", background: "#dcfce7", padding: "1px 6px", borderRadius: "4px" }}>A등급 (80점↑)</span>
                                      <span style={{ color: "#2563eb", background: "#dbeafe", padding: "1px 6px", borderRadius: "4px" }}>B등급 (60~79.9)</span>
                                      <span style={{ color: "#d97706", background: "#fef3c7", padding: "1px 6px", borderRadius: "4px" }}>C등급 (40~59.9)</span>
                                      <span style={{ color: "#dc2626", background: "#fee2e2", padding: "1px 6px", borderRadius: "4px" }}>D등급 (40점 미만)</span>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            )}
                          </div>
                        </div>

                        {/* 0-1. SNS 분석 탭 전용 종합 KPI 스코어보드 */}
                        {page === "sns" && (() => {
                          const socialTab = rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
                          const rawRow = socialTab?.socialRows?.find(r => r.branch.trim() === selectedBranch);
                          const row = rawRow ? summarizeSnsRow(rawRow) : null;
                          if (!row) return null;

                          return (
                            <div style={{ background: "linear-gradient(135deg, #ffffff 0%, #f0fdf4 100%)", border: "1px solid #bbf7d0", borderRadius: "20px", padding: "24px 28px", boxShadow: "0 10px 25px -5px rgba(34, 197, 94, 0.08)", marginBottom: "8px" }}>
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px", flexWrap: "wrap", gap: "12px" }}>
                                <div>
                                  <span style={{ fontSize: "0.78rem", fontWeight: "800", color: "#15803d", letterSpacing: "0.5px" }}>SNS DIAGNOSTIC KPI & SCORE</span>
                                  <h3 style={{ margin: "2px 0 0 0", fontSize: "1.35rem", fontWeight: "900", color: "#0f172a" }}>
                                    {selectedBranch} SNS 종합 진단 점수 및 운영 상태
                                  </h3>
                                </div>
                                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                                  <button
                                    onClick={handleBatchSyncAllBlogs}
                                    disabled={isSnsSyncing}
                                    style={{
                                      display: "inline-flex",
                                      alignItems: "center",
                                      gap: "6px",
                                      background: "#03c75a",
                                      color: "#ffffff",
                                      border: "none",
                                      padding: "8px 16px",
                                      borderRadius: "10px",
                                      fontSize: "0.82rem",
                                      fontWeight: "800",
                                      cursor: "pointer",
                                      boxShadow: "0 4px 6px -1px rgba(3, 199, 90, 0.2)",
                                      transition: "all 0.2s"
                                    }}
                                  >
                                    {isSnsSyncing ? (batchSyncProgress ? `🔄 분석 중 (${batchSyncProgress.current}/${batchSyncProgress.total})` : "동기화 중...") : "🔄 전체 지점 블로그 자동 분석"}
                                  </button>
                                  <button
                                    onClick={() => setIsCompareModalOpen(true)}
                                    style={{
                                      display: "inline-flex",
                                      alignItems: "center",
                                      gap: "6px",
                                      background: "#7c3aed",
                                      color: "#ffffff",
                                      border: "none",
                                      padding: "8px 16px",
                                      borderRadius: "10px",
                                      fontSize: "0.82rem",
                                      fontWeight: "800",
                                      cursor: "pointer",
                                      boxShadow: "0 4px 6px -1px rgba(124, 58, 237, 0.2)",
                                      transition: "all 0.2s"
                                    }}
                                    onMouseEnter={(e) => { e.currentTarget.style.transform = "translateY(-1px)"; }}
                                    onMouseLeave={(e) => { e.currentTarget.style.transform = "translateY(0)"; }}
                                  >
                                    ⚔️ 지점 1:1 비교
                                  </button>
                                  <button
                                    onClick={() => setIsSnsModalOpen(true)}
                                    style={{
                                      display: "inline-flex",
                                      alignItems: "center",
                                      gap: "6px",
                                      background: "#2563eb",
                                      color: "#ffffff",
                                      border: "none",
                                      padding: "8px 16px",
                                      borderRadius: "10px",
                                      fontSize: "0.82rem",
                                      fontWeight: "800",
                                      cursor: "pointer",
                                      boxShadow: "0 4px 6px -1px rgba(37, 99, 235, 0.2)",
                                      transition: "all 0.2s"
                                    }}
                                  >
                                    📝 SNS 지표 편집
                                  </button>
                                </div>
                              </div>

                              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))", gap: "16px" }}>
                                {/* Card 1: 종합 점수 & 등급 */}
                                <div style={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "14px", padding: "18px 20px" }}>
                                  <div style={{ fontSize: "0.78rem", color: "#64748b", fontWeight: "700", marginBottom: "6px" }}>종합 SNS 점수 / 등급</div>
                                  <div style={{ display: "flex", alignItems: "baseline", gap: "8px" }}>
                                    <span style={{ fontSize: "1.8rem", fontWeight: "900", color: "#0f172a" }}>{row.finalScore}</span>
                                    <span style={{ fontSize: "0.9rem", color: "#64748b", fontWeight: "700" }}>/ 100점</span>
                                    <span style={{
                                      marginLeft: "auto",
                                      padding: "4px 10px",
                                      borderRadius: "9999px",
                                      fontSize: "0.8rem",
                                      fontWeight: "900",
                                      background: row.grade === "A" ? "#dcfce7" : row.grade === "B" ? "#e0e7ff" : row.grade === "C" ? "#fef3c7" : "#fee2e2",
                                      color: row.grade === "A" ? "#16a34a" : row.grade === "B" ? "#4338ca" : row.grade === "C" ? "#d97706" : "#dc2626"
                                    }}>
                                      {row.grade}등급 ({row.grade === "A" ? "우수" : row.grade === "B" ? "안정" : row.grade === "C" ? "보통" : "집중관리"})
                                    </span>
                                  </div>
                                </div>

                                {/* Card 2: 블로그 점수 */}
                                <div
                                  onClick={() => {
                                    if (row.blogUrl && !isMissingChannelUrl(row.blogUrl)) {
                                      window.open(row.blogUrl, "_blank");
                                    } else {
                                      showAlert("등록된 지점 네이버 블로그 URL이 없습니다.\n상단의 '📝 SNS 지표 편집' 버튼을 통해 등록하실 수 있습니다.", "채널 안내", "warning");
                                    }
                                  }}
                                  style={{
                                    background: "#ffffff",
                                    border: "1px solid #e2e8f0",
                                    borderRadius: "14px",
                                    padding: "18px 20px",
                                    cursor: "pointer",
                                    transition: "all 0.2s cubic-bezier(0.16, 1, 0.3, 1)",
                                    boxShadow: "0 2px 4px rgba(0,0,0,0.02)"
                                  }}
                                  onMouseEnter={(e) => {
                                    e.currentTarget.style.transform = "translateY(-3px)";
                                    e.currentTarget.style.boxShadow = "0 8px 16px -4px rgba(3, 199, 90, 0.15)";
                                    e.currentTarget.style.borderColor = "#86efac";
                                  }}
                                  onMouseLeave={(e) => {
                                    e.currentTarget.style.transform = "translateY(0)";
                                    e.currentTarget.style.boxShadow = "0 2px 4px rgba(0,0,0,0.02)";
                                    e.currentTarget.style.borderColor = "#e2e8f0";
                                  }}
                                >
                                  <div style={{ fontSize: "0.78rem", color: "#03c75a", fontWeight: "800", marginBottom: "6px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                    <span style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                                      네이버 블로그 <span style={{ fontSize: "0.75rem", opacity: 0.8 }}>↗</span>
                                    </span>
                                    <span>최근 30일 <CountUpNumber value={row.blogRecentPosts || 0} decimals={0} />개</span>
                                  </div>
                                  <div style={{ display: "flex", alignItems: "baseline", gap: "8px" }}>
                                    <span style={{ fontSize: "1.8rem", fontWeight: "900", color: "#0f172a" }}><CountUpNumber value={row.blogScore} decimals={1} /></span>
                                    <span style={{ fontSize: "0.9rem", color: "#64748b", fontWeight: "700" }}>/ 50점</span>
                                  </div>
                                  <div style={{ fontSize: "0.75rem", color: "#64748b", marginTop: "4px" }}>
                                    반응 수준: <strong><CountUpNumber value={row.blogVisitScore || 0} decimals={0} />점/5점</strong> | 최근: {row.blogLastPosted && !row.blogLastPosted.startsWith("1899") ? row.blogLastPosted : "기록없음"}
                                  </div>
                                </div>

                                {/* Card 3: 인스타그램 점수 */}
                                <div
                                  onClick={() => {
                                    if (row.instagramUrl && !isMissingChannelUrl(row.instagramUrl)) {
                                      window.open(row.instagramUrl, "_blank");
                                    } else {
                                      showAlert("등록된 지점 인스타그램 URL이 없습니다.\n상단의 '📝 SNS 지표 편집' 버튼을 통해 등록하실 수 있습니다.", "채널 안내", "warning");
                                    }
                                  }}
                                  style={{
                                    background: "#ffffff",
                                    border: "1px solid #e2e8f0",
                                    borderRadius: "14px",
                                    padding: "18px 20px",
                                    cursor: "pointer",
                                    transition: "all 0.2s cubic-bezier(0.16, 1, 0.3, 1)",
                                    boxShadow: "0 2px 4px rgba(0,0,0,0.02)"
                                  }}
                                  onMouseEnter={(e) => {
                                    e.currentTarget.style.transform = "translateY(-3px)";
                                    e.currentTarget.style.boxShadow = "0 8px 16px -4px rgba(225, 48, 108, 0.15)";
                                    e.currentTarget.style.borderColor = "#f472b6";
                                  }}
                                  onMouseLeave={(e) => {
                                    e.currentTarget.style.transform = "translateY(0)";
                                    e.currentTarget.style.boxShadow = "0 2px 4px rgba(0,0,0,0.02)";
                                    e.currentTarget.style.borderColor = "#e2e8f0";
                                  }}
                                >
                                  <div style={{ fontSize: "0.78rem", color: "#e1306c", fontWeight: "800", marginBottom: "6px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                    <span style={{ display: "flex", alignItems: "center", gap: "4px" }}>
                                      인스타그램 <span style={{ fontSize: "0.75rem", opacity: 0.8 }}>↗</span>
                                    </span>
                                    <span>최근 30일 <CountUpNumber value={row.instagramRecentPosts || 0} decimals={0} />개</span>
                                  </div>
                                  <div style={{ display: "flex", alignItems: "baseline", gap: "8px" }}>
                                    <span style={{ fontSize: "1.8rem", fontWeight: "900", color: "#0f172a" }}><CountUpNumber value={row.instagramScore} decimals={1} /></span>
                                    <span style={{ fontSize: "0.9rem", color: "#64748b", fontWeight: "700" }}>/ 50점</span>
                                  </div>
                                  <div style={{ fontSize: "0.75rem", color: "#64748b", marginTop: "4px" }}>
                                    디자인: {row.instagramDesignScore || 0}점 | 반응: {row.instagramReactionScore || 0}점 | 최근: {row.instagramLastPosted && !row.instagramLastPosted.startsWith("1899") ? row.instagramLastPosted : "기록없음"}
                                  </div>
                                </div>
                              </div>
                            </div>
                          );
                        })()}
                        
                                                {/* 1. 지점 로컬 지도 위치 */}
                        {page === "sns" && (
                          <div className="sheet-panel" style={{ background: "transparent", border: "none", borderRadius: "0", padding: "20px 0", boxShadow: "none" }}>
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "24px", padding: 0, background: "transparent", border: "none" }}>
                              <div>
                                <h2 style={{ margin: 0, fontSize: "1.4rem", fontWeight: "900", color: "#2563eb", fontFamily: "'Outfit', 'Inter', sans-serif", letterSpacing: "-0.5px" }}>
                                  /LOCAL MAP POSITION
                                </h2>
                                <div style={{ fontSize: "0.78rem", color: "#64748b", marginTop: "4px", fontWeight: "500" }}>
                                  {page === "sns" ? "선택 지점의 지도상 상세 위치" : "선택 지점과 인근 경쟁 학원의 지도 위치를 한눈에 비교"}
                                </div>
                              </div>
                            </div>
                            <div style={{ height: "320px", width: "100%", borderRadius: "14px", overflow: "hidden", border: "1px solid #cbd5e1", position: "relative" }}>
                              <div id="competitor-map-leaflet" style={{ width: "100%", height: "100%" }} />
                            </div>
                          </div>
                        )}

                        {/* 2. 지점 네이버 블로그 최신 포스팅 전용 섹션 (SNS 분석 탭 전용) */}
                        {page === "sns" && (() => {
                          const ownPromos = crawledOwnPromotions[selectedBranch] || [];
                          const socialTab = rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
                          const branchSnsRow = socialTab?.socialRows?.find(r => r.branch.trim() === selectedBranch);
                          const blogUrl = branchSnsRow?.blogUrl || "";

                          return (
                            <div className="sheet-panel" style={{ background: "transparent", border: "none", borderRadius: "0", padding: "20px 0", boxShadow: "none" }}>
                              {/* 박스 바깥 상단 헤더 (대표 파란색 폰트) */}
                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "20px", flexWrap: "wrap", gap: "12px" }}>
                                <div>
                                  <h2 style={{ margin: 0, fontSize: "1.4rem", fontWeight: "900", color: "#2563eb", fontFamily: "'Outfit', 'Inter', sans-serif", letterSpacing: "-0.5px" }}>
                                    /LATEST BLOG POSTS
                                  </h2>
                                  <div style={{ fontSize: "0.78rem", color: "#64748b", marginTop: "4px", fontWeight: "500" }}>
                                    {selectedBranch} 이투스247 공식 블로그의 실시간 최신 등록 글 목록 ({ownPromos.length}개 수집됨)
                                  </div>
                                </div>
                                
                                <div style={{ display: "flex", alignItems: "center", gap: "10px" }}>
                                  {isCrawling && (
                                    <span style={{ fontSize: "0.78rem", color: "#2563eb", fontWeight: "700", animation: "pulse 1.5s infinite" }}>
                                      실시간 수집 중...
                                    </span>
                                  )}
                                  {blogUrl && !isMissingChannelUrl(blogUrl) && (
                                    <a
                                      href={blogUrl}
                                      target="_blank"
                                      rel="noopener noreferrer"
                                      style={{
                                        display: "inline-flex",
                                        alignItems: "center",
                                        gap: "4px",
                                        background: "#f1f5f9",
                                        color: "#334155",
                                        padding: "6px 14px",
                                        borderRadius: "10px",
                                        fontSize: "0.8rem",
                                        fontWeight: "700",
                                        textDecoration: "none",
                                        border: "1px solid #cbd5e1",
                                        transition: "all 0.2s"
                                      }}
                                      onMouseEnter={(e) => { e.currentTarget.style.background = "#e2e8f0"; }}
                                      onMouseLeave={(e) => { e.currentTarget.style.background = "#f1f5f9"; }}
                                    >
                                      블로그 바로가기 ↗
                                    </a>
                                  )}
                                </div>
                              </div>

                              {/* 포스팅 카드 목록 */}
                              {ownPromos.length > 0 ? (
                                <>
                                  <KeywordTagCloud
                                    posts={ownPromos}
                                    activeKeyword={selectedKeywordFilter}
                                    onSelectKeyword={setSelectedKeywordFilter}
                                  />
                                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(320px, 1fr))", gap: "16px" }}>
                                    {ownPromos
                                      .filter(p => !selectedKeywordFilter || (p.title + " " + p.content).includes(selectedKeywordFilter))
                                      .slice(0, 9)
                                      .map((post, pIdx) => (
                                    <div
                                      key={`latest-post-${pIdx}`}
                                      onClick={() => window.open(post.url, "_blank")}
                                      style={{
                                        background: "linear-gradient(135deg, #ffffff 0%, #f8fafc 100%)",
                                        border: "1px solid #e2e8f0",
                                        borderRadius: "14px",
                                        padding: "18px 20px",
                                        cursor: "pointer",
                                        display: "flex",
                                        flexDirection: "column",
                                        justifyContent: "space-between",
                                        minHeight: "130px",
                                        transition: "all 0.25s cubic-bezier(0.16, 1, 0.3, 1)",
                                        boxShadow: "0 2px 4px rgba(0,0,0,0.02)"
                                      }}
                                      onMouseEnter={(e) => {
                                        e.currentTarget.style.transform = "translateY(-3px)";
                                        e.currentTarget.style.boxShadow = "0 10px 20px -5px rgba(59, 130, 246, 0.12)";
                                        e.currentTarget.style.borderColor = "#93c5fd";
                                      }}
                                      onMouseLeave={(e) => {
                                        e.currentTarget.style.transform = "translateY(0)";
                                        e.currentTarget.style.boxShadow = "0 2px 4px rgba(0,0,0,0.02)";
                                        e.currentTarget.style.borderColor = "#e2e8f0";
                                      }}
                                    >
                                      <div>
                                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "8px" }}>
                                          <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#ffffff", background: "#03c75a", padding: "2px 8px", borderRadius: "6px", letterSpacing: "0.2px" }}>
                                            NAVER BLOG
                                          </span>
                                          <span style={{ fontSize: "0.76rem", color: "#64748b", fontWeight: "700", display: "flex", alignItems: "center", gap: "4px" }}>
                                            <span>📅</span>
                                            <span>{post.pubDate || "2026-08-24"}</span>
                                          </span>
                                        </div>
                                        <h4 style={{ margin: "0 0 8px 0", fontSize: "0.95rem", fontWeight: "800", color: "#0f172a", lineHeight: "1.45", wordBreak: "keep-all" }}>
                                          {post.title}
                                        </h4>
                                        {post.content && (
                                          <p style={{ margin: 0, fontSize: "0.82rem", color: "#64748b", lineHeight: "1.5", wordBreak: "keep-all", display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" }}>
                                            {post.content}
                                          </p>
                                        )}
                                      </div>
                                      <div style={{ marginTop: "14px", display: "flex", justifyContent: "flex-end", alignItems: "center" }}>
                                        <span style={{ fontSize: "0.75rem", color: "#2563eb", fontWeight: "700", display: "flex", alignItems: "center", gap: "2px" }}>
                                          원문 보기 ➔
                                        </span>
                                      </div>
                                    </div>
                                    ))}
                                  </div>
                                </>
                              ) : (
                                <div style={{ textAlign: "center", padding: "40px 20px", background: "#f8fafc", borderRadius: "14px", border: "1px dashed #cbd5e1" }}>
                                  <span style={{ fontSize: "1.8rem", display: "block", marginBottom: "8px" }}>📭</span>
                                  <div style={{ fontSize: "0.92rem", fontWeight: "700", color: "#334155", marginBottom: "4px" }}>
                                    {isMissingChannelUrl(blogUrl) ? "등록된 공식 네이버 블로그 URL이 없습니다." : "수집된 최신 블로그 포스팅이 없습니다."}
                                  </div>
                                  <p style={{ margin: 0, fontSize: "0.8rem", color: "#64748b" }}>
                                    {isMissingChannelUrl(blogUrl) 
                                      ? "상단의 '📝 SNS 지표 편집' 버튼을 눌러 블로그 주소를 등록하시면 실시간 포스팅이 자동으로 연동됩니다." 
                                      : "블로그 주소는 등록되어 있으나 최근 공개 글이 없거나 크롤링 중입니다."}
                                  </p>
                                </div>
                              )}
                            </div>
                          );
                        })()}

                        {/* 2. 자사 마케팅 현황 요약 */}
                        {page === "competitors" && (
                          <div className="marketing-summary-card" style={{ background: "transparent", border: "none", borderRadius: "0", padding: "12px 0", boxShadow: "none" }}>
                          
                          {/* Section Header styled like /WHAT WE DO */}
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "20px" }}>
                            <h3 style={{ margin: 0, fontSize: "1.4rem", fontWeight: "900", color: "#2563eb", fontFamily: "'Outfit', 'Inter', sans-serif", letterSpacing: "-0.5px" }}>
                              /MARKETING STATUS
                            </h3>
                            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                              <span style={{ fontSize: "0.82rem", color: "#64748b", fontWeight: "600" }}>{selectedBranch} 자사 요약</span>
                              <span style={{ background: "#2563eb", color: "#ffffff", padding: "6px 16px", borderRadius: "9999px", fontSize: "0.8rem", fontWeight: "700", boxShadow: "0 4px 6px -1px rgba(37, 99, 235, 0.2)" }}>
                                {statusInfo.participationRate}% 작동 ({statusInfo.activeCount}/6 개 활성)
                              </span>
                            </div>
                          </div>

                          {/* 3x2 collapsed-border grid layout matching the uploaded image */}
                          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: "16px" }}>
                            
                            {/* Cell 1: 247프렌즈 */}
                            <div style={{ padding: "24px", background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", position: "relative", minHeight: "150px", display: "flex", flexDirection: "column", justifyContent: "space-between", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.02)" }}>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                                  <div>
                                    <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>PLANNING</span>
                                    <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>247프렌즈</h4>
                                  </div>
                                  <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: statusInfo.activePrograms[0].active ? "#dcfce7" : "#fee2e2", color: statusInfo.activePrograms[0].active ? "#15803d" : "#b91c1c" }}>
                                    {statusInfo.activePrograms[0].active ? "ON" : "OFF"}
                                  </span>
                                </div>
                                <div style={{ marginTop: "16px" }}>
                                  <span style={{ fontSize: "1.5rem", fontWeight: "800", color: "#1e293b" }}>{statusInfo.friendsRatio}</span>
                                  <span style={{ fontSize: "0.75rem", color: "#64748b", marginLeft: "4px" }}>(진행/총)</span>
                                </div>
                              </div>
                              <div style={{ position: "absolute", bottom: "12px", right: "12px", opacity: 0.9 }}>
                                <svg width="60" height="60" viewBox="0 0 60 60" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <path d="M12 25 L48 25" stroke="#2563eb" strokeWidth="1.5" strokeLinecap="round"/>
                                  <path d="M12 25 C20 12, 40 12, 48 25" stroke="#2563eb" strokeWidth="1.5" strokeDasharray="2 2"/>
                                  <rect x="9" y="22" width="6" height="6" fill="#ffffff" stroke="#2563eb" strokeWidth="1.5"/>
                                  <rect x="45" y="22" width="6" height="6" fill="#ffffff" stroke="#2563eb" strokeWidth="1.5"/>
                                  <circle cx="30" cy="14" r="3" fill="#facc15" stroke="#2563eb" strokeWidth="1.5"/>
                                  <path d="M30 17 L30 35" stroke="#2563eb" strokeWidth="1.5"/>
                                  <path d="M26 39 L30 35 L34 39 L30 45 Z" fill="#ffffff" stroke="#2563eb" strokeWidth="1.5" strokeLinejoin="round"/>
                                </svg>
                              </div>
                            </div>

                            {/* Cell 2: 247체험단 */}
                            <div style={{ padding: "24px", background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", position: "relative", minHeight: "150px", display: "flex", flexDirection: "column", justifyContent: "space-between", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.02)" }}>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                                  <div>
                                    <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>BRANDING</span>
                                    <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>247체험단</h4>
                                  </div>
                                  <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: statusInfo.activePrograms[1].active ? "#dcfce7" : "#fee2e2", color: statusInfo.activePrograms[1].active ? "#15803d" : "#b91c1c" }}>
                                    {statusInfo.activePrograms[1].active ? "ON" : "OFF"}
                                  </span>
                                </div>
                                <div style={{ marginTop: "16px" }}>
                                  <span style={{ fontSize: "1.5rem", fontWeight: "800", color: "#1e293b" }}>{statusInfo.experienceRatio}</span>
                                  <span style={{ fontSize: "0.75rem", color: "#64748b", marginLeft: "4px" }}>(진행/총)</span>
                                </div>
                              </div>
                              <div style={{ position: "absolute", bottom: "12px", right: "12px", opacity: 0.9 }}>
                                <svg width="60" height="60" viewBox="0 0 60 60" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <circle cx="24" cy="30" r="12" fill="#67e8f9" fillOpacity="0.6" stroke="#2563eb" strokeWidth="1.5"/>
                                  <circle cx="36" cy="24" r="12" stroke="#2563eb" strokeWidth="1.5"/>
                                  <circle cx="36" cy="36" r="12" stroke="#2563eb" strokeWidth="1.5"/>
                                </svg>
                              </div>
                            </div>

                            {/* Cell 3: 협업이벤트 */}
                            <div style={{ padding: "24px", background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", position: "relative", minHeight: "150px", display: "flex", flexDirection: "column", justifyContent: "space-between", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.02)" }}>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                                  <div>
                                    <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>UI/UX DESIGN</span>
                                    <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>협업이벤트</h4>
                                  </div>
                                  <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: statusInfo.activePrograms[3].active ? "#dcfce7" : "#fee2e2", color: statusInfo.activePrograms[3].active ? "#15803d" : "#b91c1c" }}>
                                    {statusInfo.activePrograms[3].active ? "ON" : "OFF"}
                                  </span>
                                </div>
                                <div style={{ marginTop: "16px" }}>
                                  <span style={{ fontSize: "1.5rem", fontWeight: "800", color: "#1e293b" }}>{statusInfo.collabRatio}</span>
                                  <span style={{ fontSize: "0.75rem", color: "#64748b", marginLeft: "4px" }}>(진행/총)</span>
                                </div>
                              </div>
                              <div style={{ position: "absolute", bottom: "12px", right: "12px", opacity: 0.9 }}>
                                <svg width="60" height="60" viewBox="0 0 60 60" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <rect x="15" y="15" width="30" height="30" rx="4" fill="#ffffff" stroke="#2563eb" strokeWidth="1.5"/>
                                  <rect x="19" y="19" width="22" height="4" fill="#67e8f9" fillOpacity="0.5" stroke="#2563eb" strokeWidth="1"/>
                                  <rect x="19" y="27" width="10" height="14" fill="#ffffff" stroke="#2563eb" strokeWidth="1"/>
                                  <rect x="33" y="27" width="8" height="6" fill="#facc15" fillOpacity="0.8" stroke="#2563eb" strokeWidth="1"/>
                                  <rect x="33" y="37" width="8" height="4" fill="#ffffff" stroke="#2563eb" strokeWidth="1"/>
                                </svg>
                              </div>
                            </div>

                            {/* Cell 4: SNS진단표 */}
                            <div style={{ padding: "24px", background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", position: "relative", minHeight: "150px", display: "flex", flexDirection: "column", justifyContent: "space-between", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.02)" }}>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                                  <div>
                                    <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>MARKETING</span>
                                    <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>SNS진단표</h4>
                                  </div>
                                  <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: statusInfo.activePrograms[2].active ? "#dcfce7" : "#fee2e2", color: statusInfo.activePrograms[2].active ? "#15803d" : "#b91c1c" }}>
                                    {statusInfo.activePrograms[2].active ? "ON" : "OFF"}
                                  </span>
                                </div>
                                <div style={{ marginTop: "16px" }}>
                                  <span style={{ fontSize: "1.5rem", fontWeight: "800", color: "#4f46e5" }}>{statusInfo.snsRatio}</span>
                                </div>
                              </div>
                              <div style={{ position: "absolute", bottom: "12px", right: "12px", opacity: 0.9 }}>
                                <svg width="60" height="60" viewBox="0 0 60 60" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <circle cx="30" cy="30" r="16" stroke="#2563eb" strokeWidth="1.5" strokeDasharray="32 12"/>
                                  <line x1="24" y1="36" x2="24" y2="28" stroke="#facc15" strokeWidth="3" strokeLinecap="round"/>
                                  <line x1="30" y1="36" x2="30" y2="22" stroke="#2563eb" strokeWidth="3" strokeLinecap="round"/>
                                  <line x1="36" y1="36" x2="36" y2="26" stroke="#67e8f9" strokeWidth="3" strokeLinecap="round"/>
                                  <circle cx="44" cy="20" r="2" fill="#2563eb"/>
                                  <circle cx="16" cy="40" r="2" fill="#2563eb"/>
                                </svg>
                              </div>
                            </div>

                            {/* Cell 5: 지점 시설영상 */}
                            <div style={{ padding: "24px", background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", position: "relative", minHeight: "150px", display: "flex", flexDirection: "column", justifyContent: "space-between", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.02)" }}>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                                  <div>
                                    <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>PUBLISH</span>
                                    <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>지점 시설영상</h4>
                                  </div>
                                  <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: statusInfo.activePrograms[4].active ? "#dcfce7" : "#fee2e2", color: statusInfo.activePrograms[4].active ? "#15803d" : "#b91c1c" }}>
                                    {statusInfo.activePrograms[4].active ? "ON" : "OFF"}
                                  </span>
                                </div>
                                <div style={{ marginTop: "16px" }}>
                                  <span style={{ fontSize: "1.4rem", fontWeight: "800", color: "#1e293b" }}>{statusInfo.facilityRatio === "O" ? "상영중 (O)" : "미상영 (X)"}</span>
                                </div>
                              </div>
                              <div style={{ position: "absolute", bottom: "12px", right: "12px", opacity: 0.9 }}>
                                <svg width="60" height="60" viewBox="0 0 60 60" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <rect x="15" y="18" width="30" height="24" rx="3" fill="#ffffff" stroke="#2563eb" strokeWidth="1.5"/>
                                  <rect x="15" y="18" width="30" height="5" fill="#facc15" stroke="#2563eb" strokeWidth="1"/>
                                  <rect x="19" y="27" width="6" height="6" fill="#ffffff" stroke="#2563eb" strokeWidth="1"/>
                                  <rect x="29" y="27" width="12" height="6" fill="#ffffff" stroke="#2563eb" strokeWidth="1"/>
                                  <path d="M22 36 L25 39 L27 34 L31 38 L33 36 L29 32 Z" fill="#67e8f9" stroke="#2563eb" strokeWidth="1" strokeLinejoin="round"/>
                                </svg>
                              </div>
                            </div>

                            {/* Cell 6: 멘토단 및 장학생 */}
                            <div style={{ padding: "24px", background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", position: "relative", minHeight: "150px", display: "flex", flexDirection: "column", justifyContent: "space-between", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.02)" }}>
                              <div>
                                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                                  <div>
                                    <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>DEVELOP</span>
                                    <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>멘토단 및 장학생</h4>
                                  </div>
                                  <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: statusInfo.activePrograms[5].active ? "#dcfce7" : "#fee2e2", color: statusInfo.activePrograms[5].active ? "#15803d" : "#b91c1c" }}>
                                    {statusInfo.activePrograms[5].active ? "ON" : "OFF"}
                                  </span>
                                </div>
                                <div style={{ marginTop: "16px" }}>
                                  <span style={{ fontSize: "1.4rem", fontWeight: "800", color: "#1e293b" }}>{statusInfo.mentorRatio} 배출</span>
                                </div>
                              </div>
                              <div style={{ position: "absolute", bottom: "12px", right: "12px", opacity: 0.9 }}>
                                <svg width="60" height="60" viewBox="0 0 60 60" fill="none" xmlns="http://www.w3.org/2000/svg">
                                  <circle cx="40" cy="40" r="6" stroke="#2563eb" strokeWidth="1.5"/>
                                  <path d="M40 31 L40 33 M40 47 L40 49 M31 40 L33 40 M47 40 L49 40" stroke="#2563eb" strokeWidth="1.5" strokeLinecap="round"/>
                                  <line x1="16" y1="20" x2="36" y2="20" stroke="#2563eb" strokeWidth="1.5" strokeLinecap="round"/>
                                  <line x1="16" y1="28" x2="36" y2="28" stroke="#2563eb" strokeWidth="1.5" strokeLinecap="round"/>
                                  <circle cx="22" cy="20" r="3" fill="#facc15" stroke="#2563eb" strokeWidth="1.5"/>
                                  <circle cx="30" cy="28" r="3" fill="#67e8f9" stroke="#2563eb" strokeWidth="1.5"/>
                                </svg>
                              </div>
                            </div>

                          </div>
                        </div>
                        )}

                        {/* 3. 로컬 키워드 관심도 분석 (네이버 API 연동) */}
                        {page === "competitors" && (
                          <div className="marketing-summary-card" style={{ background: "transparent", border: "none", borderRadius: "0", padding: "20px 0", boxShadow: "none", position: "relative" }}>
                          
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "24px" }}>
                            <div>
                              <h3 style={{ margin: 0, fontSize: "1.4rem", fontWeight: "900", color: "#2563eb", fontFamily: "'Outfit', 'Inter', sans-serif", letterSpacing: "-0.5px" }}>
                                /LOCAL KEYWORD TREND
                              </h3>
                              <div style={{ fontSize: "0.78rem", color: "#64748b", marginTop: "4px", fontWeight: "500" }}>
                                최근 6개월간 지역 포털 통합검색 브랜드 키워드 트렌드 & 관심 점유율 진단
                              </div>
                            </div>
                            {/* Inject 3D Card Flip CSS & Marquee Styles */}
                            <style dangerouslySetInnerHTML={{ __html: `
                              .flip-card {
                                background-color: transparent;
                                perspective: 1000px;
                                width: 500px;
                                height: 600px;
                                flex-shrink: 0;
                              }
                              .flip-card-inner {
                                position: relative;
                                width: 100%;
                                height: 100%;
                                text-align: left;
                                transition: transform 0.6s cubic-bezier(0.4, 0, 0.2, 1);
                                transform-style: preserve-3d;
                              }
                              .flip-card:hover .flip-card-inner {
                                transform: rotateY(180deg);
                              }
                              .flip-card-front, .flip-card-back {
                                position: absolute;
                                width: 100%;
                                height: 100%;
                                -webkit-backface-visibility: hidden;
                                backface-visibility: hidden;
                                border-radius: 20px;
                                overflow: hidden;
                              }
                              .flip-card-front {
                                background-color: #ffffff;
                                border: 1px solid #dbeaf5;
                                color: #0f172a;
                              }
                              .flip-card-back {
                                background-color: #0f172a;
                                color: #ffffff;
                                transform: rotateY(180deg);
                                border: 1px solid rgba(255, 255, 255, 0.15);
                              }
                              .marquee-container {
                                overflow: hidden;
                                width: 100%;
                                position: relative;
                                padding: 16px 0;
                                cursor: grab;
                                user-select: none;
                                -webkit-user-select: none;
                                transform-style: preserve-3d;
                              }
                              .marquee-container:active {
                                cursor: grabbing;
                              }
                              .marquee-track {
                                display: flex;
                                gap: 24px;
                                width: max-content;
                              }
                              .flip-card, .flip-card-inner, .flip-card-front, .flip-card-back {
                                user-select: none;
                                -webkit-user-select: none;
                                -webkit-user-drag: none;
                              }
                              .flip-card-list {
                                width: 100%;
                                height: 680px;
                              }
                              .flip-card-list .flip-card-back {
                                overflow-y: auto;
                              }
                              @keyframes marquee-roll {
                                0% {
                                  transform: translateX(0);
                                }
                                100% {
                                  transform: translateX(-50%);
                                }
                              }
                            `}} />

                            <div style={{ display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "8px" }}>
                              {/* Segmented Control Toggle bar */}
                              <div style={{ display: "flex", gap: "8px", background: "transparent", padding: 0, borderRadius: "999px" }}>
                                <button
                                  onClick={() => setTrendViewMode("cards")}
                                  style={{
                                    border: trendViewMode === "cards" ? "1px solid #2563eb" : "1px solid #dbe7f3",
                                    background: trendViewMode === "cards" ? "#2563eb" : "#ffffff",
                                    color: trendViewMode === "cards" ? "#ffffff" : "#2563eb",
                                    fontWeight: "800",
                                    fontSize: "0.78rem",
                                    padding: "7px 14px",
                                    borderRadius: "999px",
                                    cursor: "pointer",
                                    transition: "all 0.2s",
                                    boxShadow: trendViewMode === "cards" ? "0 5px 12px rgba(37, 99, 235, 0.28)" : "0 1px 2px rgba(15, 23, 42, 0.06)"
                                  }}
                                >
                                  🔲 경쟁사별 카드
                                </button>
                                <button
                                  onClick={() => setTrendViewMode("all")}
                                  style={{
                                    border: trendViewMode === "all" ? "1px solid #2563eb" : "1px solid #dbe7f3",
                                    background: trendViewMode === "all" ? "#2563eb" : "#ffffff",
                                    color: trendViewMode === "all" ? "#ffffff" : "#2563eb",
                                    fontWeight: "800",
                                    fontSize: "0.78rem",
                                    padding: "7px 14px",
                                    borderRadius: "999px",
                                    cursor: "pointer",
                                    transition: "all 0.2s",
                                    boxShadow: trendViewMode === "all" ? "0 5px 12px rgba(37, 99, 235, 0.28)" : "0 1px 2px rgba(15, 23, 42, 0.06)"
                                  }}
                                >
                                  📊 한눈에 비교
                                </button>
                              </div>

                              {trendViewMode === "cards" && (
                                <div style={{ display: "flex", alignItems: "center", gap: "8px", paddingRight: "2px" }}>
                                  <span style={{ fontSize: "0.7rem", color: "#94a3b8", fontWeight: "800", letterSpacing: "0.02em" }}>
                                    보기 방식
                                  </span>
                                  <div style={{ display: "flex", gap: "6px", background: "transparent", padding: 0, borderRadius: "999px" }}>
                                  <button
                                    onClick={() => setCardViewStyle("rolling")}
                                    style={{
                                      border: cardViewStyle === "rolling" ? "1px solid #2563eb" : "1px solid #dbe7f3",
                                      background: cardViewStyle === "rolling" ? "#2563eb" : "#ffffff",
                                      color: cardViewStyle === "rolling" ? "#ffffff" : "#2563eb",
                                      fontWeight: "800",
                                      fontSize: "0.72rem",
                                      padding: "5px 11px",
                                      borderRadius: "999px",
                                      cursor: "pointer",
                                      transition: "all 0.2s",
                                      boxShadow: cardViewStyle === "rolling" ? "0 4px 10px rgba(37, 99, 235, 0.24)" : "0 1px 2px rgba(15, 23, 42, 0.05)"
                                    }}
                                    title="자동 롤링 캐러셀 보기"
                                  >
                                    🔄 롤링 뷰
                                  </button>
                                  <button
                                    onClick={() => setCardViewStyle("grid")}
                                    style={{
                                      border: cardViewStyle === "grid" ? "1px solid #2563eb" : "1px solid #dbe7f3",
                                      background: cardViewStyle === "grid" ? "#2563eb" : "#ffffff",
                                      color: cardViewStyle === "grid" ? "#ffffff" : "#2563eb",
                                      fontWeight: "800",
                                      fontSize: "0.72rem",
                                      padding: "5px 11px",
                                      borderRadius: "999px",
                                      cursor: "pointer",
                                      transition: "all 0.2s",
                                      boxShadow: cardViewStyle === "grid" ? "0 4px 10px rgba(37, 99, 235, 0.24)" : "0 1px 2px rgba(15, 23, 42, 0.05)"
                                    }}
                                    title="나열해서 전체 보기"
                                  >
                                    📋 리스트 뷰
                                  </button>
                                  </div>
                                </div>
                              )}

                            </div>
                          </div>

                          {/* Data calculation block for SOV and Targets */}
                          {(() => {
                            if (trendData.length === 0) return null;

                            // 브랜드 색상 매핑 함수
                            const getBrandColor = (name) => {
                              const clean = String(name || "").toLowerCase();
                              if (clean.includes("이투스")) return "#ff6b00"; // 이투스 주황색
                              if (clean.includes("잇올")) return "#ee1c25"; // 잇올 빨간색
                              if (clean.includes("러셀") || clean.includes("메가스터디")) return "#2f80ed"; // 러셀 파란색
                              if (clean.includes("수능선배")) return "#5856d6"; // 수능선배 보라색
                              if (clean.includes("수만휘")) return "#0f52ba"; // 수만휘 남색
                              if (clean.includes("디랩")) return "#00c7b1"; // 디랩 민트색
                              return null;
                            };

                            const defaultColors = ["#f43f5e", "#d97706", "#10b981", "#0ea5e9", "#a855f7"];

                            // 1. SOV 계산
                            const oursSum = trendData.reduce((sum, d) => sum + d.ours, 0);
                            const compSums = competitors.map((_, i) => trendData.reduce((sum, d) => sum + (d[`comp${i + 1}`] || 0), 0));
                            const totalSum = oursSum + compSums.reduce((sum, s) => sum + s, 0) || 1;

                            const oursSOV = Math.round((oursSum / totalSum) * 100);
                            const compSOVs = competitors.map((_, i) => Math.round((compSums[i] / totalSum) * 100));
                            const remainingSOV = 100 - oursSOV;

                            let normalizedCompSOVs = compSOVs;
                            if (compSOVs.length > 0) {
                              const currentSum = compSOVs.reduce((sum, s) => sum + s, 0) || 1;
                              normalizedCompSOVs = compSOVs.map(s => Math.round((s / currentSum) * remainingSOV));
                              const normalizedSum = normalizedCompSOVs.reduce((sum, s) => sum + s, 0);
                              const diff = remainingSOV - normalizedSum;
                              normalizedCompSOVs[normalizedCompSOVs.length - 1] += diff; // Adjust last element
                            }

                            // 전월 대비 SOV 변동 추이 계산
                            let sovDiff = 0;
                            if (trendData.length >= 2) {
                              const lastMonth = trendData[trendData.length - 1];
                              const prevMonth = trendData[trendData.length - 2];
                              const lastTotal = lastMonth.ours + competitors.reduce((sum, _, idx) => sum + (lastMonth[`comp${idx + 1}`] || 0), 0) || 1;
                              const prevTotal = prevMonth.ours + competitors.reduce((sum, _, idx) => sum + (prevMonth[`comp${idx + 1}`] || 0), 0) || 1;
                              const lastOursSOV = (lastMonth.ours / lastTotal) * 100;
                              const prevOursSOV = (prevMonth.ours / prevTotal) * 100;
                              sovDiff = Number((lastOursSOV - prevOursSOV).toFixed(1));
                            }

                            // 2. 연령대별 타깃 점유율 계산 (학생 vs 학부모)
                            const studentWeights = [1.1, 1.15, 1.2, 1.05, 1.0]; // Sumanhwi, SuneungSeonbae, Russell, Itall, DLab
                            const parentWeights = [0.9, 0.85, 0.8, 0.95, 1.0];

                            let rawStudentOurs = oursSOV * 0.75;
                            let rawStudentComps = competitors.map((_, i) => (normalizedCompSOVs[i] || 0) * (studentWeights[i % studentWeights.length]));
                            const studentSum = rawStudentOurs + rawStudentComps.reduce((sum, s) => sum + s, 0) || 1;
                            const sOurs = Math.max(5, Math.round((rawStudentOurs / studentSum) * 100));
                            let sComps = competitors.map((_, i) => Math.max(3, Math.round((rawStudentComps[i] / studentSum) * 100)));
                            const sCompsSum = sComps.reduce((sum, s) => sum + s, 0);
                            const remainingStudent = 100 - sOurs;
                            let normalizedSComps = sComps.map(s => Math.round((s / (sCompsSum || 1)) * remainingStudent));
                            if (normalizedSComps.length > 0) {
                              const sum = normalizedSComps.reduce((sum, s) => sum + s, 0);
                              normalizedSComps[normalizedSComps.length - 1] += (remainingStudent - sum);
                            }

                            let rawParentOurs = oursSOV * 1.25;
                            let rawParentComps = competitors.map((_, i) => (normalizedCompSOVs[i] || 0) * (parentWeights[i % parentWeights.length]));
                            const parentSum = rawParentOurs + rawParentComps.reduce((sum, s) => sum + s, 0) || 1;
                            const pOurs = Math.max(5, Math.round((rawParentOurs / parentSum) * 100));
                            let pComps = competitors.map((_, i) => Math.max(3, Math.round((rawParentComps[i] / parentSum) * 100)));
                            const pCompsSum = pComps.reduce((sum, s) => sum + s, 0);
                            const remainingParent = 100 - pOurs;
                            let normalizedPComps = pComps.map(s => Math.round((s / (pCompsSum || 1)) * remainingParent));
                            if (normalizedPComps.length > 0) {
                              const sum = normalizedPComps.reduce((sum, s) => sum + s, 0);
                              normalizedPComps[normalizedPComps.length - 1] += (remainingParent - sum);
                            }

                            // 4. Spike Detector 감지 로직
                            let spikes = [];
                            competitors.forEach((comp, i) => {
                              const key = `comp${i + 1}`;
                              for (let idx = 1; idx < trendData.length; idx++) {
                                const prev = trendData[idx - 1];
                                const curr = trendData[idx];
                                const currVal = curr[key] || 0;
                                const prevVal = prev[key] || 0;
                                if (currVal >= prevVal * 2 && currVal >= 15) {
                                  spikes.push({
                                    name: comp.name,
                                    month: curr.month,
                                    increase: Math.round((currVal / (prevVal || 1)) * 100) - 100
                                  });
                                }
                              }
                            });

                            // 특정 월의 브랜드 수치 및 순위 반환 함수
                            const getSortedRanksForMonth = (monthData) => {
                              const list = [
                                { name: "자사 (이투스247)", val: monthData.ours, color: "#ff6b00" }
                              ];
                              competitors.forEach((comp, idx) => {
                                const key = `comp${idx + 1}`;
                                const compColor = getBrandColor(comp.name) || defaultColors[idx % defaultColors.length];
                                list.push({ name: comp.name, val: monthData[key] || 0, color: compColor });
                              });
                              return list.sort((a, b) => b.val - a.val);
                            };

                                                        // Card rendering helper definition
                            const renderCard = (comp, idx, listLength) => {

                                    const compKey = `comp${(idx % competitors.length) + 1}`;
                                    const compColor = getBrandColor(comp.name) || defaultColors[(idx % competitors.length) % defaultColors.length];

                                    // 실시간 크롤링된 프로모션을 기반으로 동적 동향 및 대응지침 생성
                                    const livePromos = crawledCompPromotions[comp.name];
                                    const dynamicData = generateDynamicCompetitorData(comp.name, livePromos);

                                    const aiResult = aiAnalyzedData[comp.name];
                                    const isLoading = aiLoading[comp.name];

                                    let trendToDisplay = comp.trend;
                                    let guideToDisplay = comp.guide;

                                    if (isLoading) {
                                      trendToDisplay = "AI 분석 요약 진행 중...\n(실시간 정보 추출을 위해 약 2초간 대기)";
                                      guideToDisplay = "AI 전술 지침 실시간 로딩 중...";
                                    } else if (aiResult) {
                                      trendToDisplay = aiResult.trend;
                                      guideToDisplay = aiResult.guide;
                                    } else if (dynamicData) {
                                      trendToDisplay = dynamicData.trend;
                                      guideToDisplay = dynamicData.guide;
                                    }

                                    // Calculate 1:1 comparative SOV
                                    const oursSum11 = trendData.reduce((sum, d) => sum + d.ours, 0);
                                    const theirsSum11 = trendData.reduce((sum, d) => sum + (d[compKey] || 0), 0);
                                    const total11 = oursSum11 + theirsSum11 || 1;
                                    const oursSOV11 = Math.round((oursSum11 / total11) * 100);
                                    const theirsSOV11 = 100 - oursSOV11;

                                    // Monthly 1:1 trend data for mini line chart
                                    const oneToOneTrendData = trendData.map(d => {
                                      const oVal = d.ours;
                                      const tVal = d[compKey] || 0;
                                      const tot = oVal + tVal || 1;
                                      return {
                                        month: d.month,
                                        oursRatio: Math.round((oVal / tot) * 100),
                                        theirsRatio: Math.round((tVal / tot) * 100)
                                      };
                                    });

                                    let logoUrl = "";
                                    if (comp.name.includes("수만휘")) logoUrl = "/sumanhwi-logo.png";
                                    else if (comp.name.includes("잇올")) logoUrl = "/itall-logo.png";
                                    else if (comp.name.includes("디랩")) logoUrl = "/dlab-logo.png";
                                    else if (comp.name.includes("수능선배")) logoUrl = "/suneung-logo.png";
                                    else if (comp.name.includes("러셀")) logoUrl = "/russell-logo.png";

                                    const bgGradient = `linear-gradient(135deg, ${compColor}dd, ${compColor}99)`;

                                    return (
                                      <div 
                                        key={`${comp.name}-${idx}`} 
                                        className={`flip-card ${cardViewStyle === "grid" ? "flip-card-list" : ""}`}
                                        onMouseUp={(e) => handleMarqueeMouseUp(e, comp.name)}
                                        onMouseEnter={() => triggerAiAnalysis(comp.name, livePromos)}
                                        style={{ cursor: "pointer" }}
                                      >
                                        <div className="flip-card-inner">
                                          {/* Card Front */}
                                          <div className="flip-card-front" style={{
                                            position: "absolute",
                                            top: 0,
                                            left: 0,
                                            width: "100%",
                                            height: "100%",
                                            background: logoUrl 
                                              ? `linear-gradient(rgba(15, 23, 42, 0.55), rgba(15, 23, 42, 0.55)), url('${logoUrl}') no-repeat center center / cover` 
                                              : bgGradient,
                                            display: "flex",
                                            alignItems: "center",
                                            justifyContent: "center",
                                            color: "#ffffff",
                                            padding: "36px",
                                            textAlign: "center",
                                            overflow: "hidden",
                                            boxShadow: "inset 0 0 80px rgba(0,0,0,0.2)",
                                            backdropFilter: logoUrl ? "blur(3px)" : "none",
                                            borderRadius: "24px"
                                          }}>
                                            {!logoUrl && (
                                              <svg style={{ position: "absolute", opacity: 0.15, width: "150%", height: "150%", top: "-25%", left: "-25%", pointerEvents: "none" }} viewBox="0 0 100 100">
                                                <circle cx="50" cy="50" r="40" fill="none" stroke="#ffffff" strokeWidth="8" strokeDasharray="5,15" />
                                                <path d="M 0 50 Q 25 25, 50 50 T 100 50" fill="none" stroke="#ffffff" strokeWidth="4" />
                                              </svg>
                                            )}
                                            <div style={{ zIndex: 2, display: "flex", flexDirection: "column", alignItems: "center", width: "100%", padding: "0 16px" }}>
                                              <div style={{ fontSize: "2.0rem", fontWeight: "900", fontFamily: "'Outfit', 'Inter', sans-serif", letterSpacing: "-1px", textShadow: "0 4px 12px rgba(0,0,0,0.5)", textTransform: "uppercase", textAlign: "center", wordBreak: "keep-all", lineHeight: "1.25" }}>
                                                / {comp.name}
                                              </div>
                                              <span style={{ fontSize: "0.85rem", fontWeight: "700", letterSpacing: "2px", opacity: 0.85, textTransform: "uppercase", marginTop: "4px" }}>
                                                COMPETITOR
                                              </span>
                                            </div>
                                            <div style={{ position: "absolute", bottom: "24px", left: 0, right: 0, fontSize: "0.72rem", color: "rgba(255,255,255,0.7)", fontWeight: "600", letterSpacing: "1px" }}>
                                              🔍 마우스를 올려 상세 정보 카드 보기
                                            </div>
                                          </div>

                                          {/* Card Back */}
                                          <div className="flip-card-back" style={{ position: "absolute", top: 0, left: 0, width: "100%", height: "100%", padding: "34px", display: "flex", flexDirection: "column", justifyContent: "space-between", overflowY: "auto" }}>
                                            <div>
                                              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "16px" }}>
                                                <div>
                                                  <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#67e8f9", letterSpacing: "1px" }}>1:1 ANALYSIS</span>
                                                  <h4 style={{ margin: "2px 0 0 0", fontSize: "1.2rem", fontWeight: "900", color: "#ffffff", wordBreak: "keep-all", lineHeight: "1.2" }}>
                                                    {comp.name}
                                                  </h4>
                                                </div>
                                                <span style={{ padding: "3px 10px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: comp.urgency === "high" ? "#fee2e2" : "#fffbeb", color: comp.urgency === "high" ? "#b91c1c" : "#b45309" }}>
                                                  위협: {comp.urgency === "high" ? "상" : "중"}
                                                </span>
                                              </div>

                                              {/* 1:1 Trend mini chart */}
                                              <div style={{ background: "rgba(255,255,255,0.04)", borderRadius: "12px", padding: "12px 12px 6px 12px", marginBottom: "16px", border: "1px solid rgba(255,255,255,0.06)" }}>
                                                <div style={{ fontSize: "0.72rem", color: "#94a3b8", fontWeight: "700", marginBottom: "8px" }}>📈 자사 vs 경쟁사 검색 관심도 추이 (비율%)</div>
                                                <svg width="100%" height="65" viewBox="0 0 320 65" style={{ overflow: "visible" }}>
                                                  {/* Lines */}
                                                  {(() => {
                                                    const oPoints = oneToOneTrendData.map((d, idx) => `${20 + idx * 52},${55 - d.oursRatio * 0.45}`).join(" ");
                                                    const tPoints = oneToOneTrendData.map((d, idx) => `${20 + idx * 52},${55 - d.theirsRatio * 0.45}`).join(" ");
                                                    return (
                                                      <>
                                                        <polyline fill="none" stroke="#ff6b00" strokeWidth="2.5" points={oPoints} />
                                                        <polyline fill="none" stroke={compColor} strokeWidth="2" strokeDasharray="3,3" points={tPoints} />
                                                      </>
                                                    );
                                                  })()}
                                                  {/* Dots & Labels */}
                                                  {oneToOneTrendData.map((d, idx) => (
                                                    <g key={idx}>
                                                      <circle cx={20 + idx * 52} cy={55 - d.oursRatio * 0.45} r="3" fill="#ff6b00" />
                                                      <circle cx={20 + idx * 52} cy={55 - d.theirsRatio * 0.45} r="2.5" fill={compColor} />
                                                      <text x={20 + idx * 52} y="64" fill="#64748b" fontSize="7" textAnchor="middle" fontWeight="bold">{d.month}</text>
                                                    </g>
                                                  ))}
                                                </svg>
                                              </div>

                                              {/* 1:1 SOV Gauge bar */}
                                              <div style={{ marginBottom: "18px" }}>
                                                <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.72rem", color: "#94a3b8", fontWeight: "700", marginBottom: "6px" }}>
                                                  <span>📊 1:1 실시간 검색 점유율</span>
                                                  <span>자사 {oursSOV11}% : {theirsSOV11}%</span>
                                                </div>
                                                <div style={{ height: "8px", display: "flex", borderRadius: "4px", overflow: "hidden", background: "rgba(255,255,255,0.1)" }}>
                                                  <div style={{ width: `${oursSOV11}%`, background: "#ff6b00" }} />
                                                  <div style={{ width: `${theirsSOV11}%`, background: compColor }} />
                                                </div>
                                              </div>

                                              {/* 동향 요약 및 활성 프로모션 */}
                                              <div style={{ borderTop: "1px solid rgba(255,255,255,0.08)", paddingTop: "16px", marginBottom: "22px", fontSize: "0.94rem" }}>
                                                <span style={{ fontSize: "0.92rem", fontWeight: "800", color: "#38bdf8", display: "block", marginBottom: "8px" }}>📋 동향 요약</span>
                                                <p style={{ margin: 0, color: "#f8fafc", lineHeight: "1.65", wordBreak: "keep-all", whiteSpace: "pre-line" }}>
                                                  {trendToDisplay ? String(trendToDisplay).replace(/\s*(2\.|3\.)/g, '\n$1') : ""}
                                                </p>
                                              </div>
                                            </div>

                                            {/* AI Action Plan */}
                                            <div style={{ borderTop: "1px solid rgba(255,255,255,0.1)", paddingTop: "16px" }}>
                                              <span style={{ fontSize: "0.92rem", fontWeight: "800", color: "#34d399", display: "block", marginBottom: "8px" }}>💡 AI 전술적 대응 지침 (Action Plan)</span>
                                              <p style={{ margin: 0, fontSize: "0.94rem", color: "#f8fafc", lineHeight: "1.65", wordBreak: "keep-all", whiteSpace: "pre-line" }}>
                                                {guideToDisplay ? String(guideToDisplay).replace(/\s*(2\.|3\.)/g, '\n$1') : ""}
                                              </p>
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    );
                                  };

                            return trendViewMode === "all" ? (
                              <>
                                {/* 6개월 추이 선 그래프 */}
                                <div style={{ width: "100%", background: "#f8fafc", padding: "20px", borderRadius: "16px", border: "1px solid #f1f5f9", position: "relative", marginBottom: "24px" }}>
                                  <h4 style={{ margin: "0 0 16px 0", fontSize: "0.85rem", color: "#334155", fontWeight: "700" }}>📈 브랜드 검색 관심도 추이 (6개월간 흐름)</h4>
                                  
                                  <div style={{ position: "relative", width: "100%" }}>
                                    <svg width="100%" height="240" viewBox="0 0 500 240" style={{ overflow: "visible" }}>
                                      {/* Horizontal gridlines */}
                                      <line x1="40" y1="20" x2="460" y2="20" stroke="#e2e8f0" strokeDasharray="3,3" />
                                      <line x1="40" y1="80" x2="460" y2="80" stroke="#e2e8f0" strokeDasharray="3,3" />
                                      <line x1="40" y1="140" x2="460" y2="140" stroke="#e2e8f0" strokeDasharray="3,3" />
                                      <line x1="40" y1="200" x2="460" y2="200" stroke="#e2e8f0" strokeDasharray="3,3" />

                                      {/* Vertical hover guide line */}
                                      {hoveredMonthIndex !== null && (
                                        <line x1={40 + hoveredMonthIndex * 84} y1="15" x2={40 + hoveredMonthIndex * 84} y2="205" stroke="#94a3b8" strokeWidth="1.5" strokeDasharray="4,4" />
                                      )}

                                      {/* Lines for Ours (Indigo -> Brand Orange) */}
                                      {(() => {
                                        const points = trendData.map((d, i) => `${40 + i * 84},${200 - d.ours * 1.7}`).join(" ");
                                        return <polyline fill="none" stroke="#ff6b00" strokeWidth="3" points={points} />;
                                      })()}

                                      {/* Lines for Competitors (Dynamic Brand Colors) */}
                                      {competitors.map((comp, i) => {
                                        const key = `comp${i + 1}`;
                                        const color = getBrandColor(comp.name) || defaultColors[i % defaultColors.length];
                                        const dashArrays = ["3,3", "5,2", "1,1", "4,4", "none"];
                                        const dash = dashArrays[i % dashArrays.length];
                                        const points = trendData.map((d, idx) => `${40 + idx * 84},${200 - (d[key] || 0) * 1.7}`).join(" ");
                                        return <polyline key={key} fill="none" stroke={color} strokeWidth="2.5" strokeDasharray={dash} points={points} />;
                                      })}

                                      {/* Bullet Dots */}
                                      {trendData.map((d, i) => (
                                        <g key={i}>
                                          <circle cx={40 + i * 84} cy={200 - d.ours * 1.7} r={hoveredMonthIndex === i ? "6" : "4"} fill="#ff6b00" />
                                          {competitors.map((comp, idx) => {
                                            const key = `comp${idx + 1}`;
                                            const color = getBrandColor(comp.name) || defaultColors[idx % defaultColors.length];
                                            return <circle key={key} cx={40 + i * 84} cy={200 - (d[key] || 0) * 1.7} r={hoveredMonthIndex === i ? "5" : "3.5"} fill={color} />;
                                          })}
                                          <text x={40 + i * 84} y="220" fill="#64748b" fontSize="9" textAnchor="middle" fontWeight="bold">
                                            {d.month}
                                          </text>
                                        </g>
                                      ))}

                                      {/* Transparent hover hit-zones */}
                                      {trendData.map((d, i) => (
                                        <rect
                                          key={`hit-zone-${i}`}
                                          x={40 + i * 84 - 42}
                                          y="10"
                                          width="84"
                                          height="220"
                                          fill="transparent"
                                          style={{ cursor: "pointer" }}
                                          onMouseEnter={() => setHoveredMonthIndex(i)}
                                          onMouseLeave={() => setHoveredMonthIndex(null)}
                                        />
                                      ))}
                                    </svg>

                                    {/* Line Chart Tooltip */}
                                    {hoveredMonthIndex !== null && (() => {
                                      const sortedRanks = getSortedRanksForMonth(trendData[hoveredMonthIndex]);
                                      return (
                                        <div style={{
                                          position: "absolute",
                                          top: "10px",
                                          left: hoveredMonthIndex < 3 ? `${40 + hoveredMonthIndex * 84 + 20}px` : "auto",
                                          right: hoveredMonthIndex >= 3 ? `${500 - (40 + hoveredMonthIndex * 84) + 20}px` : "auto",
                                          background: "rgba(15, 23, 42, 0.95)",
                                          color: "#ffffff",
                                          padding: "12px 14px",
                                          borderRadius: "12px",
                                          boxShadow: "0 10px 15px -3px rgba(0,0,0,0.3)",
                                          fontSize: "0.78rem",
                                          zIndex: 1000,
                                          width: "210px",
                                          pointerEvents: "none",
                                          backdropFilter: "blur(4px)",
                                          border: "1px solid rgba(255,255,255,0.1)"
                                        }}>
                                          <div style={{ fontWeight: "800", borderBottom: "1px solid rgba(255,255,255,0.15)", paddingBottom: "6px", marginBottom: "6px", color: "#67e8f9" }}>
                                            📅 {trendData[hoveredMonthIndex].month} 검색관심도 순위
                                          </div>
                                          {sortedRanks.map((item, idx) => (
                                            <div key={idx} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: "5px" }}>
                                              <span style={{ display: "flex", alignItems: "center", gap: "6px", color: "rgba(255,255,255,0.9)" }}>
                                                <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#facc15" }}>{idx + 1}위</span>
                                                <span style={{ width: "6px", height: "6px", borderRadius: "50%", background: item.color }} />
                                                <span style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", maxWidth: "90px" }}>{item.name.replace(/스파르타|학원|센터|점/g, "").trim()}</span>
                                              </span>
                                              <strong style={{ color: "#ffffff" }}>{item.val}</strong>
                                            </div>
                                          ))}
                                        </div>
                                      );
                                    })()}
                                  </div>

                                  {/* Chart Legend */}
                                  <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: "16px", marginTop: "18px", fontSize: "0.78rem" }}>
                                    <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                                      <span style={{ display: "inline-block", width: "12px", height: "4px", borderRadius: "2px", background: "#ff6b00" }} />
                                      <span style={{ fontWeight: "700", color: "#0f172a" }}>자사 (이투스247)</span>
                                    </div>
                                    {competitors.map((comp, i) => {
                                      const color = getBrandColor(comp.name) || defaultColors[i % defaultColors.length];
                                      return (
                                        <div key={comp.name} style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                                          <span style={{ display: "inline-block", width: "12px", height: "3px", borderRadius: "1.5px", background: color }} />
                                          <span style={{ color: "#475569" }}>{comp.name}</span>
                                        </div>
                                      );
                                    })}
                                  </div>
                                </div>

                                {/* 1. SOV 및 2. 연령대별 비교 및 3. 시즌별 마케팅 골든타임 통합 그리드 */}
                                <div style={{ border: "1px solid #dbeaf5", borderRadius: "20px", overflow: "hidden", background: "#ffffff", display: "grid", gridTemplateColumns: "1fr 1.5fr", boxShadow: "0 10px 15px -3px rgba(0,0,0,0.02)", marginBottom: "24px" }}>
                                  
                                  {/* Cell 1: Donut Chart */}
                                  <div style={{ padding: "24px", borderRight: "1px solid #dbeaf5", borderBottom: "1px solid #dbeaf5", position: "relative", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                                    <div>
                                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "16px" }}>
                                        <div>
                                          <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>SOV ANALYSIS</span>
                                          <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>실시간 검색 점유율</h4>
                                        </div>
                                        <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: "#dcfce7", color: "#15803d" }}>
                                          ON
                                        </span>
                                      </div>

                                      <div style={{ display: "flex", flexDirection: "column", alignItems: "center", position: "relative", marginTop: "12px" }}>
                                        <div style={{ position: "relative", width: "180px", height: "180px", display: "flex", alignItems: "center", justifyContent: "center", overflow: "visible" }}>
                                          <svg width="180" height="180" viewBox="0 0 36 36" style={{ transform: "rotate(-90deg)", width: "100%", height: "100%", overflow: "visible" }}>
                                            <circle cx="18" cy="18" r="15.915" fill="none" stroke="#f1f5f9" strokeWidth="4" />
                                            <circle
                                              cx="18"
                                              cy="18"
                                              r="15.915"
                                              fill="none"
                                              stroke="#ff6b00"
                                              strokeWidth="4"
                                              strokeDasharray={`${oursSOV} ${100 - oursSOV}`}
                                              strokeDashoffset="0"
                                              strokeLinecap="round"
                                              style={{ cursor: "pointer", transition: "filter 0.25s, transform 0.25s, opacity 0.25s", transformOrigin: "18px 18px", transform: hoveredSliceIndex === 0 ? "scale(1.07)" : "scale(1)", filter: hoveredSliceIndex === 0 ? "drop-shadow(0 0 2.5px #ff6b00aa)" : "none", opacity: hoveredSliceIndex !== null && hoveredSliceIndex !== 0 ? 0.35 : 1 }}
                                              onMouseEnter={() => setHoveredSliceIndex(0)}
                                              onMouseLeave={() => setHoveredSliceIndex(null)}
                                            />
                                            {(() => {
                                              let currentOffset = -oursSOV;
                                              return normalizedCompSOVs.map((sov, idx) => {
                                                const strokeDash = `${sov} ${100 - sov}`;
                                                const offset = currentOffset;
                                                currentOffset -= sov;
                                                const color = getBrandColor(competitors[idx].name) || defaultColors[idx % defaultColors.length];
                                                const sIndex = idx + 1;
                                                return (
                                                  <circle
                                                    key={`donut-slice-${idx}`}
                                                    cx="18"
                                                    cy="18"
                                                    r="15.915"
                                                    fill="none"
                                                    stroke={color}
                                                    strokeWidth="4"
                                                    strokeDasharray={strokeDash}
                                                    strokeDashoffset={offset}
                                                    strokeLinecap="round"
                                                    style={{ cursor: "pointer", transition: "filter 0.25s, transform 0.25s, opacity 0.25s", transformOrigin: "18px 18px", transform: hoveredSliceIndex === sIndex ? "scale(1.07)" : "scale(1)", filter: hoveredSliceIndex === sIndex ? `drop-shadow(0 0 2.5px ${color}aa)` : "none", opacity: hoveredSliceIndex !== null && hoveredSliceIndex !== sIndex ? 0.35 : 1 }}
                                                    onMouseEnter={() => setHoveredSliceIndex(sIndex)}
                                                    onMouseLeave={() => setHoveredSliceIndex(null)}
                                                  />
                                                );
                                              });
                                            })()}
                                          </svg>
                                          <div style={{ position: "absolute", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", textAlign: "center", width: "110px", pointerEvents: "none" }}>
                                            {hoveredSliceIndex === null ? (
                                              <>
                                                <span style={{ fontSize: "0.7rem", color: "#64748b", fontWeight: "600" }}>이투스247</span>
                                                <span style={{ fontSize: "1.2rem", fontWeight: "900", color: "#0f172a" }}>{oursSOV}%</span>
                                              </>
                                            ) : hoveredSliceIndex === 0 ? (
                                              <>
                                                <span style={{ fontSize: "0.68rem", color: "#ff6b00", fontWeight: "800" }}>이투스247</span>
                                                <span style={{ fontSize: "1.2rem", fontWeight: "900", color: "#ff6b00" }}>{oursSOV}%</span>
                                              </>
                                            ) : (() => {
                                              const idx = hoveredSliceIndex - 1;
                                              const comp = competitors[idx];
                                              const color = getBrandColor(comp.name) || defaultColors[idx % defaultColors.length];
                                              const sov = normalizedCompSOVs[idx] || 0;
                                              return (
                                                <>
                                                  <span style={{ fontSize: "0.68rem", color: color, fontWeight: "800", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap", width: "70px" }}>
                                                    {comp.name.replace(/스파르타|학원|센터|점/g, "").trim()}
                                                  </span>
                                                  <span style={{ fontSize: "1.2rem", fontWeight: "900", color: color }}>{sov}%</span>
                                                </>
                                              );
                                            })()}
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>

                                  {/* Cell 2: Target Demographics */}
                                  <div style={{ padding: "24px", borderBottom: "1px solid #dbeaf5", position: "relative", display: "flex", flexDirection: "column", justifyContent: "space-between" }}>
                                    <div>
                                      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "16px" }}>
                                        <div>
                                          <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>DEMOGRAPHICS</span>
                                          <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>소비자 관심도 분석</h4>
                                        </div>
                                        <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: "#dcfce7", color: "#15803d" }}>
                                          ON
                                        </span>
                                      </div>

                                      <div style={{ display: "flex", flexDirection: "column", gap: "14px" }}>
                                        {/* 점유율 증감 수치 */}
                                        <div style={{ display: "flex", flexDirection: "column", gap: "5px", fontSize: "0.82rem" }}>
                                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                            <span style={{ color: "#334155", fontWeight: "600", display: "flex", alignItems: "center", gap: "6px" }}>
                                              <span style={{ width: "8px", height: "8px", borderRadius: "50%", background: "#ff6b00" }} />
                                              자사 점유율
                                            </span>
                                            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                                              <strong style={{ fontSize: "0.9rem", color: "#0f172a" }}>{oursSOV}%</strong>
                                              <span style={{ fontSize: "0.72rem", fontWeight: "700", padding: "2px 6px", borderRadius: "4px", background: sovDiff >= 0 ? "#ecfdf5" : "#fef2f2", color: sovDiff >= 0 ? "#10b981" : "#ef4444" }}>
                                                {sovDiff >= 0 ? `▲ ${sovDiff}%` : `▼ ${Math.abs(sovDiff)}%`}
                                              </span>
                                            </div>
                                          </div>
                                          
                                          {competitors.map((comp, idx) => {
                                            const color = getBrandColor(comp.name) || defaultColors[idx % defaultColors.length];
                                            const sov = normalizedCompSOVs[idx] || 0;
                                            return (
                                              <div key={comp.name} style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                                                <span style={{ color: "#475569", display: "flex", alignItems: "center", gap: "6px" }}>
<span style={{ width: "8px", height: "8px", borderRadius: "50%", background: color }} />
                                                  {comp.name}
                                                </span>
                                                <strong style={{ color: "#334155" }}>{sov}%</strong>
                                              </div>
                                            );
                                          })}
                                        </div>

                                        {/* 연령대별 비교 */}
                                        <div style={{ borderTop: "1px solid #f1f5f9", paddingTop: "12px", display: "flex", flexDirection: "column", gap: "10px" }}>
                                          <div>
                                            <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.75rem", color: "#475569", marginBottom: "4px", fontWeight: "600" }}>
                                              <span>👥 학생층 선호 점유율 (10-20대 초)</span>
                                              <span style={{ color: "#2563eb" }}>
                                                자사 {sOurs}% / {competitors.map((c, idx) => `${c.name.split(" ")[0]} ${normalizedSComps[idx]}%`).join(" / ")}
                                              </span>
                                            </div>
                                            <div style={{ height: "8px", display: "flex", borderRadius: "4px", overflow: "hidden" }}>
                                              <div style={{ width: `${sOurs}%`, background: "#ff6b00" }} />
                                              {competitors.map((comp, idx) => {
                                                const color = getBrandColor(comp.name) || defaultColors[idx % defaultColors.length];
                                                const w = normalizedSComps[idx] || 0;
                                                return <div key={`student-bar-${idx}`} style={{ width: `${w}%`, background: color }} />;
                                              })}
                                            </div>
                                          </div>

                                          <div>
                                            <div style={{ display: "flex", justifyContent: "space-between", fontSize: "0.75rem", color: "#475569", marginBottom: "4px", fontWeight: "600" }}>
                                              <span>👥 학부모층 신뢰 점유율 (40-50대)</span>
                                              <span style={{ color: "#2563eb" }}>
                                                자사 {pOurs}% / {competitors.map((c, idx) => `${c.name.split(" ")[0]} ${normalizedPComps[idx]}%`).join(" / ")}
                                              </span>
                                            </div>
                                            <div style={{ height: "8px", display: "flex", borderRadius: "4px", overflow: "hidden" }}>
                                              <div style={{ width: `${pOurs}%`, background: "#ff6b00" }} />
                                              {competitors.map((comp, idx) => {
                                                const color = getBrandColor(comp.name) || defaultColors[idx % defaultColors.length];
                                                const w = normalizedPComps[idx] || 0;
                                                return <div key={`parent-bar-${idx}`} style={{ width: `${w}%`, background: color }} />;
                                              })}
                                            </div>
                                          </div>
                                        </div>
                                      </div>
                                    </div>
                                  </div>

                                  {/* Cell 3: Season Timeline (span 2) */}
                                  <div style={{ gridColumn: "span 2", padding: "24px", position: "relative" }}>
                                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: "16px" }}>
                                      <div>
                                        <span style={{ fontSize: "0.72rem", fontWeight: "800", color: "#2563eb", letterSpacing: "1px" }}>MARKETING CALENDAR</span>
                                        <h4 style={{ margin: "2px 0 0 0", fontSize: "1.05rem", fontWeight: "700", color: "#0f172a" }}>시즌별 입시 마케팅 골든타임</h4>
                                      </div>
                                      <span style={{ padding: "2px 8px", borderRadius: "12px", fontSize: "0.72rem", fontWeight: "800", background: "#dcfce7", color: "#15803d" }}>
                                        ON
                                      </span>
                                    </div>

                                    <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: "12px", marginTop: "12px" }}>
                                      <div style={{ padding: "12px", borderRadius: "10px", background: "#f8fafc", border: "1px solid #e2e8f0" }}>
                                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                                          <span style={{ fontSize: "0.72rem", color: "#64748b", fontWeight: "700" }}>6월 반수반</span>
                                          <span style={{ fontSize: "0.6rem", background: "#e2e8f0", color: "#475569", padding: "1px 4px", borderRadius: "3px", fontWeight: "800" }}>종료</span>
                                        </div>
                                        <strong style={{ fontSize: "0.78rem", color: "#0f172a" }}>6 ~ 7월</strong>
                                      </div>
                                      
                                      <div style={{ padding: "12px", borderRadius: "10px", background: "#eff6ff", border: "1px solid #bfdbfe" }}>
                                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                                          <span style={{ fontSize: "0.72rem", color: "#2563eb", fontWeight: "700" }}>9월 모평반</span>
                                          <span style={{ fontSize: "0.6rem", background: "#3b82f6", color: "#ffffff", padding: "1px 4px", borderRadius: "3px", fontWeight: "800" }}>진행중</span>
                                        </div>
                                        <strong style={{ fontSize: "0.78rem", color: "#1e3a8a" }}>7 ~ 8월 (D-Day)</strong>
                                      </div>
                                      
                                      <div style={{ padding: "12px", borderRadius: "10px", background: "#fffbeb", border: "1px solid #fde68a" }}>
                                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                                          <span style={{ fontSize: "0.72rem", color: "#b45309", fontWeight: "700" }}>윈터스쿨</span>
                                          <span style={{ fontSize: "0.6rem", background: "#d97706", color: "#ffffff", padding: "1px 4px", borderRadius: "3px", fontWeight: "800" }}>D-85</span>
                                        </div>
                                        <strong style={{ fontSize: "0.78rem", color: "#78350f" }}>10 ~ 12월</strong>
                                      </div>

                                      <div style={{ padding: "12px", borderRadius: "10px", background: "#f8fafc", border: "1px solid #e2e8f0" }}>
                                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "4px" }}>
                                          <span style={{ fontSize: "0.72rem", color: "#64748b", fontWeight: "700" }}>조기선발반</span>
                                          <span style={{ fontSize: "0.6rem", background: "#e2e8f0", color: "#475569", padding: "1px 4px", borderRadius: "3px", fontWeight: "800" }}>대기</span>
                                        </div>
                                        <strong style={{ fontSize: "0.78rem", color: "#0f172a" }}>1 ~ 2월</strong>
                                      </div>
                                    </div>
                                  </div>

                                </div>

                                {/* 4. 경쟁사 마케팅 Spike 감지 경보 */}
                                {spikes.length > 0 && (
                                  <div style={{ padding: "16px", background: "#fff5f5", border: "1px solid #feb2b2", borderRadius: "16px", display: "flex", gap: "12px", alignItems: "start" }}>
                                    <span style={{ fontSize: "1.3rem" }}>⚡</span>
                                    <div>
                                      <h4 style={{ margin: "0 0 6px 0", fontSize: "0.85rem", color: "#9b2c2c", fontWeight: "800" }}>경쟁사 로컬 마케팅 급증 감지 (Spike Detector)</h4>
                                      <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
                                        {spikes.map((spike, idx) => (
                                          <p key={idx} style={{ margin: 0, fontSize: "0.78rem", color: "#742a2a", lineHeight: "1.5" }}>
                                            경쟁사 <strong>{spike.name}</strong>의 {spike.month} 검색량이 전월 대비 <strong>{spike.increase}%</strong> 급증한 스파이크를 보였습니다. 이는 로컬 대형 유료 광고 집행이나 지점 설명회 프로모션 전개의 영향입니다.
                                          </p>
                                        ))}
                                        <p style={{ margin: "6px 0 0 0", fontSize: "0.78rem", color: "#9b2c2c", fontWeight: "700" }}>
                                          💡 추천 대응안: 경쟁사들의 적극적인 홍보 공세에 맞추어 지점 블로그 및 SNS에 자사 지점의 강점인 '1:1 학습 밀착 케어 시스템' 및 '올해 주요 입결 실적'을 메인 팝업으로 걸고 포스팅 빈도를 상향하십시오.
                                        </p>
                                      </div>
                                    </div>
                                  </div>
                                )}
                              </>
                            ) : (
                                cardViewStyle === "rolling" ? (
                                <div 
                                  className="marquee-container" 
                                  style={{ marginBottom: "24px", overflow: "hidden", position: "relative" }}
                                  onMouseEnter={() => { isHoveredRef.current = true; }}
                                  onMouseLeave={() => { isHoveredRef.current = false; handleMarqueeMouseUp(null, null); }}
                                >
                                  <div 
                                    ref={trackRef}
                                    className="marquee-track"
                                    onMouseDown={handleMarqueeMouseDown}
                                    onMouseMove={handleMarqueeMouseMove}
                                    onMouseUp={(e) => handleMarqueeMouseUp(e, null)}
                                    onMouseLeave={(e) => handleMarqueeMouseUp(e, null)}
                                    style={{ 
                                      userSelect: "none", 
                                      WebkitUserDrag: "none"
                                    }}
                                  >
                                    {[...competitors, ...competitors].map((comp, idx) => renderCard(comp, idx, competitors.length))}
                                  </div>
                                </div>
                              ) : (
                                <div style={{ 
                                  display: "grid", 
                                  gridTemplateColumns: "repeat(auto-fill, minmax(360px, 1fr))", 
                                  gap: "32px", 
                                  alignItems: "start",
                                  marginBottom: "24px", 
                                  width: "100%" 
                                }}>
                                  {competitors.map((comp, idx) => renderCard(comp, idx, competitors.length))}
                                </div>
                              )
                            );
                          })()}

                        </div>
                        )}

                        {/* Card 4.5: 네이버 실시간 소셜 여론 모니터링 (지식iN, 블로그, 뉴스, 카페, 오르비) */}
                        {page === "competitors" && (
                          <div className="marketing-summary-card" style={{ background: "#ffffff", border: "1px solid #e2e8f0", borderRadius: "16px", padding: "24px", boxShadow: "0 4px 6px -1px rgba(0,0,0,0.05)" }}>
                          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: "16px" }}>
                            <div>
                              <h3 style={{ margin: "0 0 4px 0", fontSize: "1.1rem", color: "#0f172a", fontWeight: "700" }}>
                                💬 네이버 실시간 소셜 여론 모니터링
                              </h3>
                              <div style={{ fontSize: "0.78rem", color: "#64748b" }}>
                                지식iN Q&A, 네이버 블로그/뉴스, 입시 카페(수만휘), 오르비 커뮤니티에서 실시간 수집한 소셜 평판 및 지점 언급 동향
                              </div>
                            </div>
                            {isSocialLoading && (
                              <span style={{ fontSize: "0.75rem", color: "var(--primary-blue)", animation: "pulse 1.5s infinite" }}>
                                API 동기화 중...
                              </span>
                            )}
                          </div>

                          {/* Sub-tabs header */}
                          <div style={{ display: "flex", gap: "8px", borderBottom: "1px solid #e2e8f0", paddingBottom: "10px", marginBottom: "16px", flexWrap: "wrap" }}>
                            {[
                              { id: "kin", label: "🙋‍♂️ 지식iN 입시 문의" },
                              { id: "blog", label: "📝 블로그 평판 & 리뷰" },
                              { id: "news", label: "📰 교육/입시 뉴스" },
                              { id: "cafe", label: "🏫 입시·맘카페 동향" },
                              { id: "orbi", label: "🦉 오르비 게시글" }
                            ].map(tab => (
                              <button
                                key={tab.id}
                                onClick={() => setActiveSocialSubTab(tab.id)}
                                style={{
                                  padding: "6px 14px",
                                  borderRadius: "8px",
                                  fontSize: "0.82rem",
                                  fontWeight: "bold",
                                  cursor: "pointer",
                                  border: "none",
                                  background: activeSocialSubTab === tab.id ? "var(--primary-blue)" : "#f1f5f9",
                                  color: activeSocialSubTab === tab.id ? "#ffffff" : "#475569",
                                  transition: "all 0.2s",
                                  marginBottom: "4px"
                                }}
                              >
                                {tab.label}
                              </button>
                            ))}
                          </div>

                          {/* Feed List */}
                          <div style={{ display: "flex", flexDirection: "column", gap: "12px" }}>
                            {isSocialLoading ? (
                              <div style={{ display: "flex", justifyContent: "center", alignItems: "center", gap: "8px", padding: "40px 0" }}>
                                <div style={{ width: "8px", height: "8px", background: "var(--primary-blue)", borderRadius: "50%", animation: "ai-pulse 1.2s infinite ease-in-out both" }} />
                                <div style={{ width: "8px", height: "8px", background: "var(--primary-blue)", borderRadius: "50%", animation: "ai-pulse 1.2s infinite ease-in-out both", animationDelay: "0.2s" }} />
                                <div style={{ width: "8px", height: "8px", background: "var(--primary-blue)", borderRadius: "50%", animation: "ai-pulse 1.2s infinite ease-in-out both", animationDelay: "0.4s" }} />
                              </div>
                            ) : socialData[activeSocialSubTab] && socialData[activeSocialSubTab].length > 0 ? (
                              socialData[activeSocialSubTab].map((item, idx) => (
                                <div
                                  key={idx}
                                  style={{
                                    padding: "14px",
                                    border: "1px solid #f1f5f9",
                                    borderRadius: "10px",
                                    background: "#f8fafc",
                                    cursor: "pointer",
                                    transition: "all 0.2s"
                                  }}
                                  onClick={() => window.open(item.link, "_blank")}
                                  onMouseEnter={(e) => e.currentTarget.style.borderColor = "var(--primary-blue)"}
                                  onMouseLeave={(e) => e.currentTarget.style.borderColor = "#f1f5f9"}
                                >
                                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "start", marginBottom: "6px" }}>
                                    <h4
                                      style={{ margin: 0, fontSize: "0.88rem", color: "#1e293b", fontWeight: "700", lineHeight: "1.4" }}
                                      dangerouslySetInnerHTML={{ __html: item.title }}
                                    />
                                    {item.pubDate && (
                                      <span style={{ fontSize: "0.72rem", color: "#94a3b8", whiteSpace: "nowrap", marginLeft: "12px" }}>
                                        {item.pubDate}
                                      </span>
                                    )}
                                  </div>
                                  {item.description && (
                                    <p
                                      style={{ fontSize: "0.78rem", color: "#475569", margin: "6px 0 0 0", lineHeight: "1.5" }}
                                      dangerouslySetInnerHTML={{ __html: item.description }}
                                    />
                                  )}
                                </div>
                              ))
                            ) : (
                              <div style={{ textAlign: "center", padding: "30px 0", border: "1px dashed #cbd5e1", borderRadius: "10px", color: "#64748b", fontSize: "0.82rem" }}>
                                최근 30일 동안 수집된 관련 피드가 없습니다.
                              </div>
                            )}
                          </div>
                        </div>
                        )}

                        {/* 5. 로컬 마케팅 Adviser */}
                        {page === "competitors" && (
                          <div className="ai-consulting-card" style={{ background: "linear-gradient(135deg, #0f172a 0%, #1e1b4b 100%)", border: "none", borderRadius: "16px", padding: "24px", color: "#f8fafc", boxShadow: "0 10px 15px -3px rgba(0,0,0,0.1)" }}>
                          <div className="ai-consulting-header" style={{ borderBottom: "1px solid rgba(255,255,255,0.1)", paddingBottom: "12px", marginBottom: "16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                            <div className="ai-sparkle-title" style={{ display: "flex", alignItems: "center", gap: "8px", fontSize: "1.1rem", fontWeight: "700", color: "#38bdf8" }}>
                              <span>🤖</span>
                              <span>AI 로컬 마케팅 Adviser 의사결정 코칭</span>
                            </div>
                            <button
                              className="ai-generate-btn"
                              disabled={isAiGenerating}
                              onClick={() => handleGenerateAiStrategy(selectedBranch, statusInfo)}
                              style={{ padding: "6px 12px", background: "rgba(56, 189, 248, 0.15)", border: "1px solid #38bdf8", color: "#38bdf8", borderRadius: "8px", cursor: "pointer", fontSize: "0.78rem", fontWeight: "bold" }}
                            >
                              {isAiGenerating ? "재생성 중..." : "🤖 AI 코칭 재분석"}
                            </button>
                          </div>

                          <div className="ai-strategy-content-box">
                            {isAiGenerating ? (
                              <div className="ai-pulse-loader" style={{ display: "flex", alignItems: "center", gap: "10px", padding: "20px 0" }}>
                                <div className="ai-pulse-dot" style={{ background: "#38bdf8", width: "10px", height: "10px", borderRadius: "50%", animation: "ai-pulse 1s infinite" }} />
                                <span style={{ fontSize: "0.82rem", color: "#38bdf8" }}>상권 경쟁 및 마케팅 지표 분석 심층 가치 추출 중...</span>
                              </div>
                            ) : typedAiStrategy || strategy ? (
                              <div style={{ whiteSpace: "pre-wrap", fontFamily: "inherit", fontSize: "0.85rem", lineHeight: "1.6", color: "#cbd5e1" }}>
                                {(typedAiStrategy || strategy).split("\n").map((line, lIdx) => {
                                  if (line.startsWith("###")) {
                                    return <h3 key={lIdx} style={{ color: "#38bdf8", borderBottom: "1px solid rgba(255,255,255,0.1)", paddingBottom: "6px", margin: "16px 0 10px 0" }}>{line.replace("###", "").trim()}</h3>;
                                  }
                                  if (line.startsWith("####")) {
                                    return <h4 key={lIdx} style={{ color: "#818cf8", margin: "14px 0 8px 0", fontWeight: "700" }}>{line.replace("####", "").trim()}</h4>;
                                  }
                                  if (line.startsWith("-")) {
                                    const cleaned = line.replace("-", "").trim();
                                    return (
                                      <div key={lIdx} style={{ paddingLeft: "14px", position: "relative", marginBottom: "8px" }}>
                                        <span style={{ position: "absolute", left: 0, color: "#38bdf8" }}>•</span>
                                        {cleaned}
                                      </div>
                                    );
                                  }
                                  return <p key={lIdx} style={{ margin: "0 0 10px 0" }}>{line}</p>;
                                })}
                              </div>
                            ) : (
                              <div style={{ textAlign: "center", color: "#64748b", padding: "20px 0" }}>
                                분석 중 에러가 발생했거나 데이터가 없습니다.
                              </div>
                            )}
                          </div>
                        </div>
                        )}

                      </div>
                    );
                  })() : (
                    /* Chip guidelines when no branch selected */
                    <div style={{ padding: "80px 40px", textAlign: "center", background: "#ffffff", border: "2px dashed #cbd5e1", borderRadius: "20px", color: "#64748b" }}>
                      <span style={{ fontSize: "2.4rem", display: "block", marginBottom: "16px" }}>
                        {page === "sns" ? "📱" : "🔍"}
                      </span>
                      <h3 style={{ margin: "0 0 8px 0", color: "#0f172a", fontWeight: "700" }}>
                        {page === "sns" ? "SNS 분석을 진행할 이투스247 지점을 선택하여 주십시오" : "분석할 이투스247 지점을 선택하여 주십시오"}
                      </h3>
                      <p style={{ fontSize: "0.88rem", margin: "0 0 24px 0", color: "#64748b", lineHeight: "1.5" }}>
                        {page === "sns" ? (
                          <>
                            상단의 검색창에서 지점명을 입력하거나, 아래의 추천 주요 지점 단추를 선택하시면<br />
                            실시간 네이버 블로그 자동 채점, 인스타그램 운영 지표 및 상세 분석 리포트가 즉시 전개됩니다.
                          </>
                        ) : (
                          <>
                            상단의 검색창에서 지점명을 입력하거나, 아래의 추천 주요 지점 단추를 선택하시면<br />
                            실시간 경쟁사 마케팅 자동 크롤링과 로컬 키워드 분석 리포트가 즉시 전개됩니다.
                          </>
                        )}
                      </p>

                      {/* Suggested chips list */}
                      <div style={{ display: "flex", flexWrap: "wrap", justifyContent: "center", gap: "10px", maxWidth: "650px", margin: "0 auto" }}>
                        {(page === "sns" 
                          ? ["분당정자", "강북", "다산", "목동", "광명", "김포", "부산대", "대구달서", "인천송도", "대전둔산"]
                          : ["목동", "인천송도", "도봉", "대전둔산", "일산서구", "강남", "분당정자", "부산대", "대구수성", "광주동구"]
                        ).map(b => (
                          <button
                            key={`suggest-chip-${b}`}
                            onClick={() => {
                              setSearchQuery(b);
                              setSelectedBranch(b);
                              setTypedAiStrategy("");
                            }}
                            style={{ padding: "8px 16px", background: "#f1f5f9", border: "1px solid #cbd5e1", borderRadius: "20px", color: "#334155", fontSize: "0.85rem", cursor: "pointer", fontWeight: "600", transition: "all 0.2s" }}
                            onMouseEnter={(e) => { e.target.style.background = "var(--primary-blue)"; e.target.style.color = "#ffffff"; e.target.style.borderColor = "var(--primary-blue)"; }}
                            onMouseLeave={(e) => { e.target.style.background = "#f1f5f9"; e.target.style.color = "#334155"; e.target.style.borderColor = "#cbd5e1"; }}
                          >
                            {page === "sns" ? "📱 " : "📍 "}{b}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}

                </div>
              </div>
            ) : (
        <div className="workbook" style={{ marginTop: "40px" }}>
          <section className="sheet-panel premium-studio-panel">
            <div className="panel-title-row premium-studio-header">
              <h2>RAWDATA Studio</h2>
              <span className="note-text">
                {activeTab.kind === SPECIAL_SOCIAL_TAB_KIND
                  ? "SNS 채널 진단표 전용 입력 형식입니다."
                  : activeTab.kind === SPECIAL_COLLAB_TAB_KIND
                  ? "지점별 협업 URL 등록 전용 입력 형식입니다."
                  : activeTab.kind === SPECIAL_FACILITY_TAB_KIND
                  ? "지점 시설영상 URL 전용 입력 형식입니다."
                  : activeTab.kind === SPECIAL_MENTOR_TAB_KIND
                  ? "멘토단 및 장학생 인적사항 관리 전용 입력 형식입니다."
                  : "`지역`, `지점`은 고정이고 이벤트만 확장됩니다."}
              </span>
            </div>
            <div className="editor-toolbar">
              <div className="editor-name-block">
                <div className="editor-meta">탭 이름</div>
                <input
                  className="tab-name-input"
                  value={activeTab.name}
                  onChange={(e) => updateTabName(e.target.value)}
                />
              </div>
              {!isSpecialTabKind(activeTab.kind) ? (
                <div className="editor-actions">
                  <button className="reset-button" onClick={removeActiveTab} disabled={rawTabs.length === 1}>
                    탭 삭제
                  </button>
                  <button className="reset-button" onClick={addRawTab}>
                    새 탭 추가
                  </button>
                </div>
              ) : null}
            </div>

            <div className="utility-row">
              <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
                <span className="utility-badge navy">RAWDATA 편집기</span>
                <span className="utility-badge yellow">
                  {activeTab.kind === SPECIAL_SOCIAL_TAB_KIND
                    ? "지점별 SNS 계정 및 점수 진단"
                    : activeTab.kind === SPECIAL_COLLAB_TAB_KIND
                    ? "지점별 홈페이지/블로그/인스타 협업 링크"
                    : activeTab.kind === SPECIAL_FACILITY_TAB_KIND
                    ? "지점별 시설 동영상 URL"
                    : activeTab.kind === SPECIAL_MENTOR_TAB_KIND
                    ? "멘토단 여부 선택 및 장학 정보"
                    : "지점별 이벤트 참여여부 및 인원수"}
                </span>
              </div>
              <div className="editor-actions" style={{ marginLeft: "auto", display: "flex", gap: "10px" }}>
                <button
                  className="mini-button highlight"
                  onClick={() => importInputRef.current?.click()}
                  title="엑셀 파일을 가져와 현재 데이터를 교체합니다."
                >
                  엑셀 불러오기
                </button>
                {activeTab.kind !== SPECIAL_SOCIAL_TAB_KIND &&
                activeTab.kind !== SPECIAL_COLLAB_TAB_KIND &&
                activeTab.kind !== SPECIAL_FACILITY_TAB_KIND &&
                activeTab.kind !== SPECIAL_MENTOR_TAB_KIND ? (
                  <button
                    className="mini-button"
                    onClick={() => exportToExcel(activeTab)}
                    title="현재 데이터를 엑셀 파일로 다운로드합니다."
                  >
                    엑셀 내보내기
                  </button>
                ) : null}
                <button
                  className="mini-button"
                  onClick={addRow}
                >
                  {activeTab.kind === SPECIAL_SOCIAL_TAB_KIND
                    ? "+ 진단 행 추가"
                    : activeTab.kind === SPECIAL_COLLAB_TAB_KIND
                    ? "+ URL 행 추가"
                    : activeTab.kind === SPECIAL_FACILITY_TAB_KIND
                    ? "+ 영상 행 추가"
                    : activeTab.kind === SPECIAL_MENTOR_TAB_KIND
                    ? "+ 학생 추가"
                    : "+ 지점 행 추가"}
                </button>
                {activeTab.kind !== SPECIAL_SOCIAL_TAB_KIND &&
                activeTab.kind !== SPECIAL_FACILITY_TAB_KIND &&
                activeTab.kind !== SPECIAL_MENTOR_TAB_KIND ? (
                  <button className="reset-button" onClick={addEvent}>
                    + 이벤트 추가
                  </button>
                ) : null}
              </div>
              <input
                type="file"
                ref={importInputRef}
                style={{ display: "none" }}
                accept=".xlsx, .xls"
                onChange={async (e) => {
                  const file = e.target.files?.[0];
                  if (file) {
                    if (activeTab.kind === SPECIAL_SOCIAL_TAB_KIND) {
                      await importSocialWorkbook(file);
                    } else if (activeTab.kind === SPECIAL_COLLAB_TAB_KIND) {
                      await importCollabWorkbook(file);
                    } else if (activeTab.kind === SPECIAL_FACILITY_TAB_KIND) {
                      await importFacilityWorkbook(file);
                    } else if (activeTab.kind === SPECIAL_MENTOR_TAB_KIND) {
                      await importMentorWorkbook(file);
                    } else {
                      await importDefaultWorkbook(file);
                    }
                  }
                  e.target.value = "";
                }}
              />
            </div>
            <div className="program-chip-row">
              {rawTabs.map((tab) => (
                <button
                  key={tab.id}
                  className={`program-chip raw-tab-chip ${activeTab?.id === tab.id ? "active" : ""}`}
                  onClick={() => {
                    sortMentorRowsState();
                    setActiveTabId(tab.id);
                  }}
                >
                  {tab.name}
                </button>
              ))}
            </div>
          </section>

          {activeTab ? (
            <section className="sheet-panel" style={{ marginTop: "14px", border: "1px solid rgba(0, 59, 255, 0.15)", borderRadius: "16px", overflow: "hidden" }}>
              {activeTab.kind === SPECIAL_SOCIAL_TAB_KIND ? (
                <div className="table-shell special-input-shell">
                  <table className="excel-table special-input-table">
                    <thead>
                      <tr>
                        <th className="special-head special-identity" style={{ width: "130px" }}>지점</th>
                        {specialSocialColumns
                          .filter((c) => c.key !== "branch")
                          .map((col) => (
                            <th key={col.key} className="special-head special-sns-head" style={{ width: "120px" }}>
                              {col.label}
                            </th>
                          ))}
                        <th className="special-head special-memo" style={{ width: "80px" }}>행 삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(activeTab.socialRows || []).map((row, rowIndex) => (
                        <tr key={row.id}>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.branch ?? ""}
                              onChange={(e) => updateSocialCell(rowIndex, "branch", e.target.value)}
                              placeholder="지점"
                            />
                          </td>
                          {specialSocialColumns
                            .filter((c) => c.key !== "branch")
                            .map((col) => (
                              <td key={col.key} className="special-cell">
                                <input
                                  type={col.type === "number" ? "number" : col.type === "date" ? "date" : "text"}
                                  min={col.type === "number" ? "0" : undefined}
                                  max={col.key.includes("Score") && col.key.includes("Visit") ? "5" : col.key.includes("Score") && col.key.includes("Reaction") ? "5" : col.key.includes("Score") && col.key.includes("Design") ? "5" : col.key.includes("Score") ? "3" : undefined}
                                  value={row[col.key] ?? ""}
                                  onChange={(e) => updateSocialCell(rowIndex, col.key, e.target.value)}
                                  placeholder={col.label}
                                />
                              </td>
                            ))}
                          <td className="special-cell special-memo" style={{ textAlign: "center" }}>
                            <button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : activeTab.kind === SPECIAL_FACILITY_TAB_KIND ? (
                <div className="table-shell special-input-shell">
                  <table className="excel-table special-input-table">
                    <thead>
                      <tr>
                        <th className="special-head special-identity" style={{ width: "100px" }}>지역</th>
                        <th className="special-head special-identity" style={{ width: "130px" }}>지점</th>
                        <th className="special-head special-growth">시설영상 URL</th>
                        <th className="special-head special-memo" style={{ width: "80px" }}>행 삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(activeTab.facilityRows || []).map((row, rowIndex) => (
                        <tr key={row.id}>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.region ?? ""}
                              onChange={(e) => updateFacilityCell(rowIndex, "region", e.target.value)}
                              placeholder="지역"
                            />
                          </td>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.branch ?? ""}
                              onChange={(e) => updateFacilityCell(rowIndex, "branch", e.target.value)}
                              placeholder="지점"
                            />
                          </td>
                          <td className="special-cell special-growth">
                            <input
                              type="url"
                              value={row.url ?? ""}
                              onChange={(e) => updateFacilityCell(rowIndex, "url", e.target.value)}
                              placeholder="시설영상 URL"
                            />
                          </td>
                          <td className="special-cell special-memo" style={{ textAlign: "center" }}>
                            <button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : activeTab.kind === SPECIAL_MENTOR_TAB_KIND ? (
                <div className="table-shell special-input-shell">
                  <table className="excel-table special-input-table">
                    <thead>
                      <tr>
                        <th className="special-head special-identity" style={{ width: "80px", textAlign: "center" }}>멘토여부</th>
                        <th className="special-head special-identity" style={{ width: "90px" }}>연도</th>
                        <th className="special-head special-identity" style={{ width: "120px" }}>이름</th>
                        <th className="special-head special-growth" style={{ width: "160px" }}>합격 대학</th>
                        <th className="special-head special-growth" style={{ width: "140px" }}>학과</th>
                        <th className="special-head special-identity" style={{ width: "130px" }}>지점</th>
                        <th className="special-head special-blog" style={{ width: "110px" }}>장학 그룹</th>
                        <th className="special-head special-blog-score" style={{ width: "130px" }}>장학 금액</th>
                        <th className="special-head special-memo">비고</th>
                        <th className="special-head special-memo" style={{ width: "80px" }}>행 삭제</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(activeTab.mentorRows || []).map((row, rowIndex) => (
                        <tr key={row.id}>
                          <td className="special-cell special-identity" style={{ textAlign: "center" }}>
                            <input
                              type="checkbox"
                              checked={row.isMentor || false}
                              onChange={(e) => updateMentorCell(rowIndex, "isMentor", e.target.checked)}
                              style={{ width: "20px", height: "20px", cursor: "pointer" }}
                            />
                          </td>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.year ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "year", e.target.value)}
                              placeholder="연도"
                            />
                          </td>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.name ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "name", e.target.value)}
                              placeholder="이름"
                            />
                          </td>
                          <td className="special-cell special-growth">
                            <input
                              type="text"
                              value={row.university ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "university", e.target.value)}
                              placeholder="합격 대학"
                            />
                          </td>
                          <td className="special-cell special-growth">
                            <input
                              type="text"
                              value={row.department ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "department", e.target.value)}
                              placeholder="학과"
                            />
                          </td>
                          <td className="special-cell special-identity">
                            <select
                              value={row.branch ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "branch", e.target.value)}
                              style={{ width: "100%", height: "100%", border: "none", background: "transparent", color: "var(--text-color)", outline: "none" }}
                            >
                              <option value="" style={{ background: "var(--panel-bg)" }}>지점 선택</option>
                              {allBranches.map((branch) => (
                                <option key={branch} value={branch} style={{ background: "var(--panel-bg)" }}>{branch}</option>
                              ))}
                            </select>
                          </td>
                          <td className="special-cell special-blog">
                            <input
                              type="text"
                              value={row.group ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "group", e.target.value)}
                              placeholder="예: 1그룹"
                            />
                          </td>
                          <td className="special-cell special-blog-score">
                            <input
                              type="number"
                              min="0"
                              value={row.amount ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "amount", e.target.value)}
                              placeholder="장학 금액"
                            />
                          </td>
                          <td className="special-cell special-memo">
                            <input
                              type="text"
                              value={row.memo ?? ""}
                              onChange={(e) => updateMentorCell(rowIndex, "memo", e.target.value)}
                              placeholder="비고"
                            />
                          </td>
                          <td className="special-cell special-memo" style={{ textAlign: "center" }}>
                            <button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              ) : activeTab.kind === SPECIAL_COLLAB_TAB_KIND ? (
                <div className="table-shell special-input-shell">
                  {(() => {
                    const collabColumns = activeTab.collabColumns || defaultCollabColumns;
                    const collabEventGroups = groupCollabColumns(collabColumns);

                    return (
                      <table className="excel-table special-input-table">
                        <thead>
                          <tr>
                            <th className="special-head special-identity" rowSpan={2}>지역</th>
                            <th className="special-head special-identity" rowSpan={2}>지점</th>
                            {collabEventGroups.map((group) => (
                              <th
                                key={`group-${group.eventName}`}
                                className="special-head special-collab-group special-collab-group-head event-group-start"
                                style={getCollabColumnThemeStyle(group.columns[0]?.key)}
                                colSpan={group.columns.length}
                              >
                                <input
                                  className="collab-group-name-input"
                                  value={group.eventName}
                                  onChange={(e) => renameCollabEvent(group.eventName, e.target.value)}
                                />
                              </th>
                            ))}
                            <th className="special-head special-memo" rowSpan={2}>행 삭제</th>
                          </tr>
                          <tr>
                            {collabEventGroups.flatMap((group) =>
                              group.columns.map((column, colIndex) => (
                                <th
                                  key={`head-${column.key}`}
                                  className={`special-head special-collab-group special-collab-channel ${colIndex === 0 ? "event-group-start" : ""}`}
                                  style={getCollabColumnThemeStyle(column.key)}
                                >
                                  {column.label}
                                </th>
                              ))
                            )}
                          </tr>
                        </thead>
                        <tbody>
                          {(activeTab.collabRows || []).map((row, rowIndex) => (
                            <tr key={row.id}>
                              <td className="special-cell special-identity">
                                <input
                                  type="text"
                                  value={row.values?.["지역"] ?? ""}
                                  onChange={(e) => updateCollabCell(rowIndex, "지역", e.target.value)}
                                  placeholder="지역"
                                />
                              </td>
                              <td className="special-cell special-identity">
                                <input
                                  type="text"
                                  value={row.values?.["지점"] ?? ""}
                                  onChange={(e) => updateCollabCell(rowIndex, "지점", e.target.value)}
                                  placeholder="지점"
                                />
                              </td>
                              {collabEventGroups.flatMap((group) =>
                                group.columns.map((column, colIndex) => (
                                  <td
                                    key={`${row.id}-${column.key}`}
                                    className={`special-cell special-collab-group ${colIndex === 0 ? "event-group-start" : ""}`}
                                    style={getCollabColumnThemeStyle(column.key)}
                                  >
                                    <input
                                      type="url"
                                      value={row.values?.[column.key] ?? ""}
                                      onChange={(e) => updateCollabCell(rowIndex, column.key, e.target.value)}
                                      placeholder={`${group.eventName} ${column.label}`}
                                    />
                                  </td>
                                ))
                              )}
                              <td className="special-cell special-memo">
                                <button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button>
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    );
                  })()}
                </div>
              ) : (
                <div className="table-shell special-input-shell">
                  <table className="excel-table special-input-table">
                    <thead>
                      <tr>
                        <th className="special-head special-identity" rowSpan={2} style={{ width: "100px" }}>지역</th>
                        <th className="special-head special-identity" rowSpan={2} style={{ width: "130px" }}>지점</th>
                        {(activeTab.events || []).map((event) => (
                          <th key={event.id} colSpan={2} className="special-head" style={{ textAlign: "center" }}>
                            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: "6px" }}>
                              <input
                                className="event-name-input"
                                value={event.name}
                                onChange={(e) => updateEventName(event.id, e.target.value)}
                                style={{ width: "90px", textAlign: "center", background: "transparent", border: "none", color: "inherit", fontWeight: "bold", outline: "none" }}
                              />
                              <button
                                className="mini-button remove-event-btn"
                                onClick={() => removeEvent(event.id)}
                                style={{ padding: "2px 4px", fontSize: "0.8rem", cursor: "pointer", background: "transparent", color: "#ef4444", border: "none" }}
                                title="이벤트 삭제"
                              >
                                ✕
                              </button>
                            </div>
                          </th>
                        ))}
                        <th className="special-head special-memo" rowSpan={2} style={{ width: "80px" }}>행 삭제</th>
                      </tr>
                      <tr>
                        {(activeTab.events || []).map((event) => (
                          <Fragment key={`sub-${event.id}`}>
                            <th className="special-head" style={{ width: "90px", textAlign: "center" }}>참여여부</th>
                            <th className="special-head" style={{ width: "90px", textAlign: "center" }}>참여인원</th>
                          </Fragment>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(activeTab.rows || []).map((row, rowIndex) => (
                        <tr key={row.id}>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.region ?? ""}
                              onChange={(e) => updateBaseCell(rowIndex, "region", e.target.value)}
                              placeholder="지역"
                            />
                          </td>
                          <td className="special-cell special-identity">
                            <input
                              type="text"
                              value={row.branch ?? ""}
                              onChange={(e) => updateBaseCell(rowIndex, "branch", e.target.value)}
                              placeholder="지점"
                            />
                          </td>
                          {(activeTab.events || []).map((event) => {
                            const val = row.eventValues?.[event.id] || { status: "X", participants: "0" };
                            return (
                              <Fragment key={`${row.id}-${event.id}`}>
                                <td className="special-cell" style={{ textAlign: "center" }}>
                                  <select
                                    value={val.status === "O" ? "O" : "X"}
                                    onChange={(e) => updateEventCell(rowIndex, event.id, "status", e.target.value)}
                                    style={{ width: "100%", height: "100%", border: "none", background: "transparent", color: "var(--text-color)", outline: "none", textAlignLast: "center" }}
                                  >
                                    <option value="O">O</option>
                                    <option value="X">X</option>
                                  </select>
                                </td>
                                <td className="special-cell">
                                  <input
                                    type="number"
                                    min="0"
                                    value={val.participants ?? "0"}
                                    onChange={(e) => updateEventCell(rowIndex, event.id, "participants", e.target.value)}
                                    placeholder="0"
                                    style={{ textAlign: "right" }}
                                  />
                                </td>
                              </Fragment>
                            );
                          })}
                          <td className="special-cell special-memo" style={{ textAlign: "center" }}>
                            <button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          ) : null}
        </div>
      )}

      {/* 8. 최하단 ETOOS247 시그니처 푸터 */}
      <footer className="premium-footer">
        <div className="premium-footer-top">
          <h2 className="premium-footer-title">/LET'S CONNECT US.</h2>
          <div className="premium-footer-buttons">
            <a href="https://247.etoos.com/index.do" target="_blank" rel="noopener noreferrer" className="premium-footer-btn">CONTACT US</a>
          </div>
          <p className="premium-footer-copy">
            이투스ECI 주식회사 | 서울특별시 서초구 남부순환로 2547, 3층 (서초동 1354-3)<br />
            COPYRIGHT ⓒ ETOOS ECI Co.,Ltd. ALL RIGHTS RESERVED.
          </p>
        </div>
        <div className="premium-footer-signature-container">
          <div className="premium-footer-signature">ETOOS247</div>
          <div className="premium-footer-badges">
            <span className="premium-footer-badge top-btn" onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}>▲ TOP</span>
          </div>
        </div>
      {selectedBranch && (() => {
        const socialTab = rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND);
        const selectedRow = socialTab?.socialRows?.find(r => r.branch.trim() === selectedBranch.trim());
        return (
          <SnsMetricEditorModal
            isOpen={isSnsModalOpen}
            onClose={() => setIsSnsModalOpen(false)}
            branchName={selectedBranch}
            socialRow={selectedRow}
            onSave={handleSaveSnsMetrics}
            onAutoSyncBlog={handleAutoSyncBlog}
            isSyncing={isSnsSyncing}
          />
        );
      })()}
      <VideoModal
        isOpen={activeVideoUrl !== null}
        onClose={() => {
          setActiveVideoUrl(null);
          setActiveVideoBranch("");
        }}
        url={activeVideoUrl}
        branchName={activeVideoBranch}
      />
            <BranchComparisonModal
        isOpen={isCompareModalOpen}
        onClose={() => setIsCompareModalOpen(false)}
        branchA={selectedBranch}
        rawTabs={rawTabs}
        allBranches={allBranches}
      />
      <CustomDialogModal
        modal={customModal}
        onClose={() => setCustomModal(null)}
      />
      </footer>
    </div>
  );
}



