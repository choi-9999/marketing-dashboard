"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";

const OVERVIEW_TAB_ID = "__overall__";
const SPECIAL_SOCIAL_TAB_KIND = "special-social";
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
  createSpecialSocialTab("tab-social-1", "SNS 진단표")
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

function ensureSpecialInputTab(tabs) {
  if (tabs.some((tab) => tab.kind === SPECIAL_SOCIAL_TAB_KIND)) {
    return tabs;
  }

  const seededBranches = [...new Set(
    tabs
      .filter((tab) => tab.kind !== SPECIAL_SOCIAL_TAB_KIND)
      .flatMap((tab) => tab.rows.map((row) => row.branch.trim()).filter(Boolean))
  )];

  return [
    ...tabs,
    createSpecialSocialTab("tab-social-1", "SNS 진단표", seededBranches.map((branch) => ({ branch })))
  ];
}

function normalizeRawTabs(rawTabs) {
  if (!Array.isArray(rawTabs) || rawTabs.length === 0) {
    return initialTabs;
  }

  return ensureSpecialInputTab(rawTabs.map(migrateLegacyTab));
}

function summarizeTab(tab) {
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

  rawTabs.filter((tab) => tab.kind !== SPECIAL_SOCIAL_TAB_KIND).forEach((tab) => {
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
    .filter((tab) => tab.kind !== SPECIAL_SOCIAL_TAB_KIND)
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

function isDormantSince(dateString, baseDateString) {
  if (!dateString) return false;
  const base = new Date(baseDateString);
  const target = new Date(dateString);
  if (Number.isNaN(base.getTime()) || Number.isNaN(target.getTime())) return false;
  const diff = Math.floor((base - target) / (1000 * 60 * 60 * 24));
  return diff > 30;
}

function summarizeSnsRow(row, baseDate = snsEvaluationBaseDate) {
  const hasBlog = !isMissingChannelUrl(row.blogUrl);
  const hasInstagram = !isMissingChannelUrl(row.instagramUrl);

  const blogActivity = hasBlog ? Math.max(0, getRecentActivityScore(row.blogRecentPosts, [[8, 30], [5, 25], [3, 20], [1, 10]]) - (isDormantSince(row.blogLastPosted, baseDate) ? 5 : 0)) : 0;
  const blogReaction = hasBlog ? Number(row.blogVisitScore || 0) : 0;
  const blogScore = hasBlog ? Number(((blogActivity + blogReaction) / 35 * 50).toFixed(1)) : 0;

  const instagramActivity = hasInstagram ? Math.max(0, getRecentActivityScore(row.instagramRecentPosts, [[12, 30], [8, 25], [4, 20], [1, 10]]) - (isDormantSince(row.instagramLastPosted, baseDate) ? 5 : 0)) : 0;
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
  const missingPenalty = (hasBlog ? 0 : 5) + (hasInstagram ? 0 : 5);
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
  const [activeTabId, setActiveTabId] = useState(initialTabs[0].id);
  const [dashboardTabId, setDashboardTabId] = useState(OVERVIEW_TAB_ID);
  const [saveState, setSaveState] = useState("서버 저장 대기 중");
  const [isHydrated, setIsHydrated] = useState(false);
  const [selectedBranch, setSelectedBranch] = useState(null);
  const [selectedOverviewBranch, setSelectedOverviewBranch] = useState(null);
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
    const payload = {
      page,
      rawTabs,
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
      const payload = {
        page,
        rawTabs,
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
    () => rawTabs.filter((tab) => tab.kind !== SPECIAL_SOCIAL_TAB_KIND),
    [rawTabs]
  );

  const selectedDashboardTab = useMemo(
    () => rawTabs.find((tab) => tab.id === dashboardTabId) ?? rawTabs[0],
    [dashboardTabId, rawTabs]
  );

  const isOverviewDashboard = dashboardTabId === OVERVIEW_TAB_ID;
  const isSpecialDashboard = !isOverviewDashboard && selectedDashboardTab?.kind === SPECIAL_SOCIAL_TAB_KIND;
  const dashboardTab = isSpecialDashboard ? null : selectedDashboardTab;

  useEffect(() => {
    setSelectedBranch(null);
    setSelectedOverviewBranch(null);
  }, [dashboardTabId]);

  useEffect(() => {
    setBranchKeyword("");
    setAreRegionsExpanded(false);
  }, [dashboardTabId]);

  const dashboardTabs = useMemo(() => {
    if (isOverviewDashboard) return dashboardRawTabs;
    return dashboardTab ? [dashboardTab] : [];
  }, [dashboardRawTabs, dashboardTab, isOverviewDashboard]);

  const dashboardSummary = useMemo(() => buildDashboardData(dashboardRawTabs), [dashboardRawTabs]);
  const scopedSummary = useMemo(() => buildDashboardData(dashboardTabs), [dashboardTabs]);
  const maxRegionBranches = useMemo(() => Math.max(...scopedSummary.regionOverview.map((item) => item.activeBranches), 0), [scopedSummary.regionOverview]);
  const hoveredRegionData = scopedSummary.regionOverview.find((item) => item.region === hoveredRegion) || null;

  const dashboardScopeLabel = isOverviewDashboard ? "전체 현황" : selectedDashboardTab?.name || "활성화 방안 대시보드";

  const branchOptions = useMemo(
    () =>
      (dashboardTab?.rows || [])
        .map((row) => row.branch.trim())
        .filter(Boolean)
        .filter((branch, index, list) => list.indexOf(branch) === index),
    [dashboardTab]
  );

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
        participants: Number(row?.eventValues?.[event.id]?.participants || 0)
      }));
    }

    return dashboardTab.events.map((event) => ({
      id: event.id,
      label: event.name,
      participants: dashboardTab.rows.reduce(
        (sum, row) => sum + Number(row.eventValues?.[event.id]?.participants || 0),
        0
      )
    }));
  }, [dashboardTab, isOverviewDashboard, selectedBranch]);

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

  const snsSourceRows = useMemo(() => {
    const snsTab = rawTabs.find((tab) => tab.kind === SPECIAL_SOCIAL_TAB_KIND);
    return snsTab ? (snsTab.socialRows || []).map((row) => summarizeSnsRow(row)) : [];
  }, [rawTabs]);

  const snsGradeGroups = useMemo(() => ({
    A: snsDashboardRows.filter((row) => row.grade === "A"),
    B: snsDashboardRows.filter((row) => row.grade === "B"),
    C: snsDashboardRows.filter((row) => row.grade === "C"),
    D: snsDashboardRows.filter((row) => row.grade === "D")
  }), [snsDashboardRows]);

  const snsSummary = useMemo(() => {
    const totalBranches = snsDashboardRows.filter((row) => row.branch.trim()).length;
    const averageScore = totalBranches > 0 ? Number((snsDashboardRows.reduce((sum, row) => sum + row.finalScore, 0) / totalBranches).toFixed(1)) : 0;
    const bothChannels = snsDashboardRows.filter((row) => row.hasBlog && row.hasInstagram).length;
    const missingChannels = snsDashboardRows.filter((row) => !row.hasBlog || !row.hasInstagram).length;
    const topBranches = [...snsDashboardRows].sort((a, b) => b.finalScore - a.finalScore).slice(0, 8);
    const lowBranches = [...snsDashboardRows].sort((a, b) => a.finalScore - b.finalScore).slice(0, 8);
    return { totalBranches, averageScore, bothChannels, missingChannels, topBranches, lowBranches };
  }, [snsDashboardRows]);

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

  const overallBranchScoreboard = useMemo(() => {
    const branchMap = new Map();
    const snsScoreMap = new Map(
      snsSourceRows
        .filter((row) => row.branch.trim())
        .map((row) => [normalizeBranchKey(row.branch), row])
    );

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

    const rawList = [...branchMap.values()];
    const maxParticipants = Math.max(...rawList.map((item) => item.totalParticipants), 1);

    const branches = rawList
      .map((item) => {
        const snsMatch = snsScoreMap.get(normalizeBranchKey(item.branch));
        const participationRate = item.eligibleEvents > 0 ? Math.round((item.participatedEvents / item.eligibleEvents) * 100) : 0;
        const participantScore = Math.round((item.totalParticipants / maxParticipants) * 100);
        const planCoverage = dashboardRawTabs.length > 0 ? Math.round((item.activePlans.size / dashboardRawTabs.length) * 100) : 0;
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
          activePlanCount: item.activePlans.size,
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
  }, [dashboardRawTabs, snsSourceRows]);

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

      return {
        ...tab,
        rows: [...tab.rows, createRow(tab.events.map((event) => event.id))]
      };
    });
  }

  function removeRow(rowIndex) {
    updateActiveTab((tab) => {
      if (tab.kind === SPECIAL_SOCIAL_TAB_KIND) {
        return {
          ...tab,
          socialRows: (tab.socialRows || []).filter((_, index) => index !== rowIndex)
        };
      }

      return {
        ...tab,
        rows: tab.rows.filter((_, index) => index !== rowIndex)
      };
    });
  }

  function addEvent() {
    if (activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND) return;
    const nextName = window.prompt("추가할 이벤트명을 입력하세요.", "신규 이벤트");
    if (!nextName) return;

    updateActiveTab((tab) => {
      const newEvent = createEvent(nextName.trim());
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

  async function importDefaultWorkbook(file) {
    if (!file || activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND) return;

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
    if (activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND) return;
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
    <div className="workbook">
      <header className="sheet-topbar">
        <div>
          <p className="sheet-kicker">ETOOS247 MARKETING DASHBOARD</p>
          <h1>지점 활성화 방안 대시보드</h1>
        </div>
        <div className="topbar-meta">
          <div className="meta-cell"><span>담당부서</span><strong>교육팀</strong></div>
          <div className="meta-cell"><span>대시보드 기준</span><strong>{page === "dashboard" ? dashboardScopeLabel : activeTab?.name || "-"}</strong></div>
          <div className="meta-cell"><span>{page === "dashboard" && isSpecialDashboard ? "진단 항목 수" : "이벤트 수"}</span><strong>{page === "dashboard" ? (isSpecialDashboard ? specialSocialColumns.length : scopedSummary.totalEvents) : dashboardSummary.totalEvents}</strong></div>
          <div className="meta-cell"><span>{page === "dashboard" && isSpecialDashboard ? "평가 지점 수" : "고유 지점 수"}</span><strong>{page === "dashboard" ? (isSpecialDashboard ? snsSummary.totalBranches : scopedSummary.uniqueBranches) : dashboardSummary.uniqueBranches}</strong></div>
          <div className="meta-cell highlight"><span>저장 상태</span><strong>{saveState}</strong></div>
        </div>
      </header>

      <div className="page-tabs">
        <button className={`page-tab ${page === "dashboard" ? "active" : ""}`} onClick={() => setPage("dashboard")}>Dashboard</button>
        <button className={`page-tab ${page === "rawdata" ? "active" : ""}`} onClick={() => setPage("rawdata")}>RAWDATA Studio</button>
      </div>

      <main className="sheet-body">
        {page === "dashboard" ? (
          <>
            <section className="sheet-panel utility-panel">
              <div className="program-chip-row dashboard-scope-row">
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
                        <strong>{overallBranchScoreboard.grouped["A그룹"].length}</strong>
                        <p>우수 운영 상태의 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">A그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {overallBranchScoreboard.grouped["A그룹"].length > 0 ? overallBranchScoreboard.grouped["A그룹"].map((branch) => <li key={`grade-a-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box hover-score-box">
                        <span>B그룹</span>
                        <strong>{overallBranchScoreboard.grouped["B그룹"].length}</strong>
                        <p>안정적으로 운영 중인 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">B그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {overallBranchScoreboard.grouped["B그룹"].length > 0 ? overallBranchScoreboard.grouped["B그룹"].map((branch) => <li key={`grade-b-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box hover-score-box">
                        <span>C그룹</span>
                        <strong>{overallBranchScoreboard.grouped["C그룹"].length}</strong>
                        <p>보완이 필요한 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">C그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {overallBranchScoreboard.grouped["C그룹"].length > 0 ? overallBranchScoreboard.grouped["C그룹"].map((branch) => <li key={`grade-c-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                      <div className="score-box warn hover-score-box">
                        <span>D그룹</span>
                        <strong>{overallBranchScoreboard.grouped["D그룹"].length}</strong>
                        <p>집중 관리가 필요한 지점입니다.</p>
                        <div className="score-tooltip">
                          <div className="score-tooltip-title">D그룹 지점명</div>
                          <ul className="score-tooltip-list">
                            {overallBranchScoreboard.grouped["D그룹"].length > 0 ? overallBranchScoreboard.grouped["D그룹"].map((branch) => <li key={`grade-d-${branch.branch}`}>{branch.branch}</li>) : <li>해당 지점이 없습니다.</li>}
                          </ul>
                        </div>
                      </div>
                    </div>
                  </article>
                </section>

                <section className="sheet-panel">
                  <div className="panel-title-row">
                    <h2>지점별 전체 현황 보드</h2>
                    <span className="note-text">운영 점수 50%와 SNS 평가 점수 50%를 합산한 그룹 보드입니다.</span>
                  </div>
                  <div className="overview-summary-strip">
                    <div className="overview-summary-card"><span>전체 지점 수</span><strong>{overallBranchScoreboard.branches.length}</strong></div>
                    <div className="overview-summary-card"><span>평균 점수</span><strong>{overallBranchScoreboard.avgScore}점</strong></div>
                    <div className="overview-summary-card"><span>최상위 지점</span><strong>{overallBranchScoreboard.topBranch?.branch || "-"}</strong></div>
                    <div className="overview-summary-card warn"><span>집중 관리 지점</span><strong>{overallBranchScoreboard.atRiskCount}</strong></div>
                  </div>
                  <div className="grade-board">
                    {["A그룹", "B그룹", "C그룹", "D그룹"].map((grade) => (
                      <section className={`grade-column ${grade === "D그룹" ? "warn" : ""}`} key={grade}>
                        <div className="grade-column-head">
                          <div>
                            <strong>{grade}</strong>
                            <span>{overallBranchScoreboard.grouped[grade].length}개 지점</span>
                          </div>
                        </div>
                        <div className="grade-card-list">
                          {overallBranchScoreboard.grouped[grade].length > 0 ? overallBranchScoreboard.grouped[grade].map((branch) => (
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
                                  <div className="grade-metric-row">
                                    <span>참여율</span>
                                    <strong>{branch.participationRate}%</strong>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>참여 횟수</span>
                                    <strong>{branch.participatedEvents}회</strong>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>총 참여 인원</span>
                                    <strong>{branch.totalParticipants}명</strong>
                                  </div>
                                  <div className="grade-metric-row">
                                    <span>참여 활성화 방안</span>
                                    <strong>{branch.activePlanCount}개</strong>
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
                          {snsGradeGroups[grade].length > 0 ? snsGradeGroups[grade].slice(0, 8).map((row) => (
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
                        {snsDashboardRows.filter((row) => row.branch.trim()).map((row) => (
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
                            <div className="event-bar-item" key={item.id}>
                              <div className="event-bar-value">{item.participants}</div>
                              <div className="event-bar-track">
                                <div
                                  className="event-bar-fill"
                                  style={{ height: `${Math.max((item.participants / maxChartParticipants) * 100, item.participants > 0 ? 10 : 0)}%` }}
                                />
                              </div>
                              <div className="event-bar-label">{item.label}</div>
                            </div>
                          ))}
                        </div>
                      </div>
                    </div>
                    <div className="event-summary-card">
                      <h3>현재 보기 요약</h3>
                      <ul className="metric-list">
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
                      </ul>
                    </div>
                  </div>
                </section>
              </>
            )}
          </>
        ) : (
          <>
            <section className="sheet-panel utility-panel">
              <div className="utility-row">
                <button className="reset-button" onClick={addRawTab}>+ RAWDATA 탭 추가</button>
                <button className="reset-button" onClick={removeActiveTab} disabled={rawTabs.length === 1}>현재 탭 삭제</button>
                <button className="reset-button" onClick={addRow}>{activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND ? "+ 진단 행 추가" : "+ 지점 행 추가"}</button>
                {activeTab?.kind !== SPECIAL_SOCIAL_TAB_KIND ? <button className="reset-button" onClick={addEvent}>+ 이벤트 추가</button> : null}
                <button className="reset-button" onClick={forceServerSave}>강제 서버 저장</button>
                <button className="reset-button" onClick={() => importInputRef.current?.click()}>엑셀 불러오기</button>
                <input
                  ref={importInputRef}
                  type="file"
                  accept=".xlsx,.xlsm,.xls"
                  style={{ display: "none" }}
                  onChange={async (e) => {
                    const file = e.target.files?.[0];
                    if (file) {
                      if (activeTab?.kind === SPECIAL_SOCIAL_TAB_KIND) {
                        await importSnsWorkbook(file);
                      } else {
                        await importDefaultWorkbook(file);
                      }
                    }
                    e.target.value = "";
                  }}
                />
              </div>
              <div className="program-chip-row">
                {rawTabs.map((tab) => <button key={tab.id} className={`program-chip raw-tab-chip ${activeTab?.id === tab.id ? "active" : ""}`} onClick={() => setActiveTabId(tab.id)}>{tab.name}</button>)}
              </div>
            </section>

            {activeTab ? (
              <section className="sheet-panel">
                <div className="panel-title-row"><h2>RAWDATA Studio</h2><span className="note-text">{activeTab.kind === SPECIAL_SOCIAL_TAB_KIND ? "SNS 채널 진단표 전용 입력 형식입니다." : "`지역`, `지점`은 고정이고 이벤트만 확장됩니다."}</span></div>
                <div className="editor-toolbar">
                  <div className="editor-name-block">
                    <div className="editor-meta">탭 이름</div>
                    <input className="tab-name-input" value={activeTab.name} onChange={(e) => updateTabName(e.target.value)} />
                  </div>
                  {activeTab.kind !== SPECIAL_SOCIAL_TAB_KIND ? (
                    <button className="mini-button" onClick={() => setAreEventChipsExpanded((current) => !current)}>
                      {areEventChipsExpanded ? "이벤트명 접기" : "이벤트명 펼치기"}
                    </button>
                  ) : null}
                </div>
                {activeTab.kind !== SPECIAL_SOCIAL_TAB_KIND ? (
                  <>
                    <div className={`event-chip-list ${areEventChipsExpanded ? "" : "collapsed"}`}>
                      {activeTab.events.length > 0 ? activeTab.events.map((event) => (
                        <div className="event-chip" key={event.id}>
                          <input
                            className="event-name-input"
                            value={event.name}
                            onChange={(e) => updateEventName(event.id, e.target.value)}
                            aria-label="이벤트명"
                          />
                          <button className="mini-button" onClick={() => removeEvent(event.id)}>삭제</button>
                        </div>
                      )) : <p className="empty-copy">아직 이벤트가 없습니다. `이벤트 추가` 버튼으로 시작할 수 있어요.</p>}
                    </div>
                    <div className="table-shell">
                      <table className="excel-table editor-table">
                        <thead>
                          <tr>
                            <th rowSpan={2}>지역</th>
                            <th rowSpan={2}>지점</th>
                            {activeTab.events.map((event) => (
                              <th key={event.id} colSpan={2}>
                                <div className="event-header-cell">
                                  <strong>{event.name}</strong>
                                  <span>참여여부 / 참석인원</span>
                                </div>
                              </th>
                            ))}
                            <th rowSpan={2}>행 삭제</th>
                          </tr>
                          <tr>
                            {activeTab.events.flatMap((event) => ([
                              <th key={`${event.id}-status`}>참여여부</th>,
                              <th key={`${event.id}-participants`}>참석인원</th>
                            ]))}
                          </tr>
                        </thead>
                        <tbody>
                          {activeTab.rows.map((row, rowIndex) => (
                            <tr key={row.id}>
                              <td><input value={row.region} onChange={(e) => updateBaseCell(rowIndex, "region", e.target.value)} /></td>
                              <td><input value={row.branch} onChange={(e) => updateBaseCell(rowIndex, "branch", e.target.value)} /></td>
                              {activeTab.events.flatMap((event) => {
                                const eventValue = row.eventValues?.[event.id] || { status: "X", participants: "0" };
                                return [
                                  <td key={`${row.id}-${event.id}-status`}>
                                    <select value={eventValue.status} onChange={(e) => updateEventCell(rowIndex, event.id, "status", e.target.value)}>
                                      <option value="O">O</option>
                                      <option value="X">X</option>
                                    </select>
                                  </td>,
                                  <td key={`${row.id}-${event.id}-participants`}>
                                    <input type="number" min="0" value={eventValue.participants} onChange={(e) => updateEventCell(rowIndex, event.id, "participants", e.target.value)} />
                                  </td>
                                ];
                              })}
                              <td><button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                ) : (
                  <div className="table-shell special-input-shell">
                    <table className="excel-table special-input-table">
                      <thead>
                        <tr>
                          {specialSocialColumns.map((column) => (
                            <th key={column.key} className={`special-head special-${column.group}`}>{column.label}</th>
                          ))}
                          <th className="special-head special-memo">행 삭제</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(activeTab.socialRows || []).map((row, rowIndex) => (
                          <tr key={row.id}>
                            {specialSocialColumns.map((column) => (
                              <td key={`${row.id}-${column.key}`} className={`special-cell special-${column.group}`}>
                                <input
                                  type={column.type === "number" ? "number" : column.type}
                                  min={column.type === "number" ? "0" : undefined}
                                  value={row[column.key] ?? ""}
                                  onChange={(e) => updateSpecialCell(rowIndex, column.key, e.target.value)}
                                  placeholder={column.label}
                                />
                              </td>
                            ))}
                            <td className="special-cell special-memo"><button className="mini-button" onClick={() => removeRow(rowIndex)}>삭제</button></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </section>
            ) : null}
          </>
        )}
      </main>
      <footer className="sheet-footer">Copyright ⓒ ETOOS ECI Co.,Ltd. All rights Reserved.</footer>
    </div>
  );
}













