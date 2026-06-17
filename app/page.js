"use client";

import { useEffect, useMemo, useRef, useState } from "react";
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
  if (tab?.kind === SPECIAL_SOCIAL_TAB_KIND) {
    return {
      id: tab.id || createId("tab"),
      name: tab.name || "SNS 진단표",
      kind: SPECIAL_SOCIAL_TAB_KIND,
      events: [],
      rows: [],
      socialRows: Array.isArray(tab.socialRows) && tab.socialRows.length > 0
        ? tab.socialRows.map((row, index) => createSpecialSocialRow({ ...row, id: row?.id || createId(`social-row-${index}`) }))
        : [createSpecialSocialRow()]
    };
  }

  if (tab?.kind === SPECIAL_COLLAB_TAB_KIND) {
    const collabColumns = normalizeCollabColumns(tab.collabColumns || defaultCollabColumns);
    return {
      id: tab.id || createId("tab"),
      name: tab.name || "협업이벤트",
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

  if (tab?.kind === SPECIAL_FACILITY_TAB_KIND) {
    return {
      id: tab.id || createId("tab"),
      name: tab.name || "지점시설영상",
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

  if (tab?.kind === SPECIAL_MENTOR_TAB_KIND) {
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
      name: tab.name || "멘토단 및 장학생",
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
      name: tab.name || "이름 없는 탭",
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
      name: label || (statusColumns.length === 1 ? tab.name || "기본 이벤트" : `이벤트 ${index + 1}`),
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
        name: tab.name || "기본 이벤트",
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
    name: tab.name || "이름 없는 탭",
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

export default function HomePage() {
  const [page, setPage] = useState("dashboard");
  const [rawTabs, setRawTabs] = useState(initialTabs);

  const [activeSlideIndex, setActiveSlideIndex] = useState(0);

  const showcaseSlides = useMemo(() => [
    { name: "프렌즈", className: "card-theme-friends", desc: "지점 활성화를 위한 동반자, 247프렌즈 캠페인", id: rawTabs.find(t => t.name === "247프렌즈")?.id },
    { name: "체험단", className: "card-theme-experience", desc: "직접 체험하고 검증하는 247체험단 홍보", id: rawTabs.find(t => t.name === "247체험단")?.id },
    { name: "SNS 진단표", className: "card-theme-sns", desc: "블로그 및 인스타그램 지점 마케팅 정밀 진단", id: rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND)?.id },
    { name: "협업이벤트", className: "card-theme-collab", desc: "대형 교육 협력사와의 제휴 공동 프로모션", id: rawTabs.find(t => t.kind === SPECIAL_COLLAB_TAB_KIND)?.id },
    { name: "지점시설영상", className: "card-theme-facility", desc: "지점 시설 및 분위기를 전달하는 고화질 영상", id: rawTabs.find(t => t.kind === SPECIAL_FACILITY_TAB_KIND)?.id },
    { name: "멘토단 및 장학생", className: "card-theme-mentor", desc: "합격의 결실을 함께 나누는 명예로운 장학제도", id: rawTabs.find(t => t.kind === SPECIAL_MENTOR_TAB_KIND)?.id }
  ], [rawTabs]);

  const marqueeCards = useMemo(() => [
    { name: "247프렌즈", category: "CAMPAIGN", className: "card-theme-friends", id: rawTabs.find(t => t.name === "247프렌즈")?.id },
    { name: "247체험단", category: "MARKETING", className: "card-theme-experience", id: rawTabs.find(t => t.name === "247체험단")?.id },
    { name: "SNS 진단표", category: "ANALYSIS", className: "card-theme-sns", id: rawTabs.find(t => t.kind === SPECIAL_SOCIAL_TAB_KIND)?.id },
    { name: "협업이벤트", category: "COLLABORATION", className: "card-theme-collab", id: rawTabs.find(t => t.kind === SPECIAL_COLLAB_TAB_KIND)?.id },
    { name: "지점시설영상", category: "PROMOTION", className: "card-theme-facility", id: rawTabs.find(t => t.kind === SPECIAL_FACILITY_TAB_KIND)?.id },
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
    if (menuName === "종합 성과") {
      document.getElementById("stats-section")?.scrollIntoView({ behavior: "smooth" });
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
  const [hoveredRegion, setHoveredRegion] = useState(null);
  const saveTimeoutRef = useRef(null);
  const hasInitializedSaveRef = useRef(false);
  const importInputRef = useRef(null);

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

  const allBranches = useMemo(() => {
    const list = rawTabs
      .filter((tab) => !isSpecialTabKind(tab.kind))
      .flatMap((tab) => tab.rows.map((row) => row.branch.trim()))
      .filter(Boolean);
    return [...new Set(list)].sort();
  }, [rawTabs]);

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
          <span className="premium-navbar-title">이투스247 학원 | 마케팅 관리 시스템</span>
        </div>
        <ul className="premium-navbar-menu">
          {["공지사항", "종합 성과", "명예의 전당", "지점 대시보드", "체험단 신청"].map((menu) => (
            <li key={menu} className="premium-navbar-menu-item">
              <a href="#" onClick={(e) => { e.preventDefault(); handleNavMenuClick(menu); }}>{menu}</a>
            </li>
          ))}
        </ul>
        <div className="premium-navbar-actions">
          <button
            className={`premium-navbar-btn ${page === "dashboard" ? "active" : ""}`}
            onClick={() => {
              sortMentorRowsState();
              setPage("dashboard");
            }}
          >
            GUEST 지점 사용자
          </button>
          <button
            className={`premium-navbar-btn ${page === "rawdata" ? "active" : ""}`}
            onClick={() => {
              sortMentorRowsState();
              setPage("rawdata");
            }}
          >
            🔒 관리자 모드
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
            <h2 className="marquee-title">Welcome</h2>
            <div className="marquee-container">
              <div className="marquee-track">
                {[...marqueeCards, ...marqueeCards].map((card, idx) => (
                  <div
                    key={`${card.id || card.name}-${idx}`}
                    className={`rolling-card ${card.className}`}
                    onClick={() => handleCardClick(card.id)}
                  >
                    <div className="rolling-card-info">
                      <div className="rolling-card-category">{card.category}</div>
                      <div className="rolling-card-title">{card.name}</div>
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
                        if (slide.id) {
                          sortMentorRowsState();
                          setDashboardTabId(slide.id);
                          setActiveTabId(slide.id);
                          document.getElementById("our-work-section")?.scrollIntoView({ behavior: "smooth" });
                        }
                      }}
                    >
                      <div className={`vertical-showcase-image-fallback ${slide.className}`}>
                        <div style={{ opacity: 0.12, fontSize: "clamp(3rem, 10vw, 8rem)", fontWeight: 900, color: "#fff", textTransform: "uppercase" }}>
                          {slide.name}
                        </div>
                      </div>
                      <div className="vertical-showcase-overlay">
                        <h3 className="vertical-showcase-tag">
                          <span className="blue-hash">#</span>
                          <span className="white-text">{slide.name}</span>
                        </h3>
                        <p style={{ color: "rgba(255, 255, 255, 0.8)", fontSize: "1.1rem", marginTop: "12px", maxWidth: "600px", margin: "12px 0 0" }}>
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

          {/* 6. 통계 매핑 시각화 */}
          <section id="stats-section" className="stats-section">
            <div className="stats-header-row">
              <h2 className="stats-title">
                <span className="blue-slash">/</span>WHAT WE DO
              </h2>
            </div>
            <div className="stats-grid">
              <div className="stats-cell">
                <div className="stats-cell-label">247프렌즈</div>
                <div className="stats-cell-value">{globalStats.friendsBranchCount}개 지점</div>
                <div className="stats-cell-desc">프렌즈 참여 지점 수<br />(평균 {globalStats.friendsAvgParticipants}명 참석)</div>
              </div>
              <div className="stats-cell">
                <div className="stats-cell-label">247체험단</div>
                <div className="stats-cell-value">{globalStats.experienceTotalParticipants}명</div>
                <div className="stats-cell-desc">체험단 총 등록 인원 수<br />({globalStats.experienceBranchCount}개 지점 활성화)</div>
              </div>
              <div className="stats-cell">
                <div className="stats-cell-label">SNS 마케팅</div>
                <div className="stats-cell-value">{globalStats.snsAvgScore}점 / 5.0</div>
                <div className="stats-cell-desc">지점 SNS 평균 진단 점수<br />(블로그 및 인스타 채널 종합)</div>
              </div>
              <div className="stats-cell">
                <div className="stats-cell-label">협업 제휴</div>
                <div className="stats-cell-value">{globalStats.collabTotalUrls}개 URL</div>
                <div className="stats-cell-desc">협업 이벤트 URL 총 등록 수<br />(진행 제휴 이벤트 {globalStats.collabEventCount}종)</div>
              </div>
              <div className="stats-cell">
                <div className="stats-cell-label">시설 홍보</div>
                <div className="stats-cell-value">{globalStats.facilityRatio}%</div>
                <div className="stats-cell-desc">지점 시설영상 업로드 완료 비율<br />(연결 대상 지점 기준)</div>
              </div>
              <div className="stats-cell">
                <div className="stats-cell-label">멘토 및 장학</div>
                <div className="stats-cell-value">{(globalStats.mentorTotalAmount / 10000).toLocaleString()}만원</div>
                <div className="stats-cell-desc">선발 장학생 지급액 합계<br />(총 {globalStats.mentorTotalCount}명 선발 완료)</div>
              </div>
            </div>
          </section>

          {/* 7. 상세 대시보드 영역 (OUR WORK) */}
          <main id="our-work-section" className="sheet-body" style={{ marginTop: "40px", borderRadius: "16px", border: "1px solid rgba(0,59,255,0.1)", boxShadow: "0 10px 30px rgba(0,0,0,0.03)" }}>
            
            {/* Quick Toggle in OUR WORK */}
            <section className="sheet-panel utility-panel" style={{ border: "none", marginBottom: "20px" }}>
              <div className="program-chip-row dashboard-scope-row" style={{ padding: 0 }}>
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

            <h2 className="active-tab-details-title" style={{ margin: "20px 0 30px" }}>
              <span className="blue-slash">/</span>OUR WORK ({isOverviewDashboard ? "전체 현황" : selectedDashboardTab?.name})
            </h2>

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
                  <div className="panel-title-row"><h2>권역별 시설영상 현황</h2><span className="note-text">지점명을 누르면 시설영상 URL이 새 탭에서 열립니다.</span></div>
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
                                target="_blank"
                                rel="noreferrer"
                                title={`${row.branch} 시설영상 열기`}
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
              ) : (
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
            <a href="#" className="premium-footer-btn" onClick={(e) => { e.preventDefault(); alert("준비 중인 채널입니다."); }}>CONTACT US</a>
            <a href="#" className="premium-footer-btn" onClick={(e) => { e.preventDefault(); alert("준비 중인 채널입니다."); }}>BROCHURE</a>
          </div>
          <p className="premium-footer-copy">
            이투스ECI 주식회사 | 서울특별시 서초구 남부순환로 2547, 3층 (서초동 1354-3)<br />
            COPYRIGHT ⓒ ETOOS ECI Co.,Ltd. ALL RIGHTS RESERVED.
          </p>
        </div>
        <div className="premium-footer-signature-container">
          <div className="premium-footer-signature">ETOOS247</div>
          <div className="premium-footer-badges">
            <span className="premium-footer-badge yellow">이투스ECI | 서울특별시 서초구 남부순환로 2547, 3층</span>
            <span className="premium-footer-badge yellow">대표번호: 1599-2470</span>
            <span className="premium-footer-badge blue">WWW.ETOOS247.CO.KR</span>
            <span className="premium-footer-badge top-btn" onClick={() => window.scrollTo({ top: 0, behavior: 'smooth' })}>▲ TOP</span>
          </div>
        </div>
      </footer>
    </div>
  );
}













