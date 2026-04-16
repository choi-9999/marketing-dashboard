const programMeta = {
  all: { label: "전체 프로그램", key: null },
  friends: { label: "247프렌즈", key: "friends" },
  trial: { label: "247체험단", key: "trial" },
  exam: { label: "입시교육", key: "exam" },
  marketing: { label: "마케팅교육", key: "marketing" },
  press: { label: "언론보도", key: "press" },
  event: { label: "협업이벤트", key: "event" }
};

const workbookSeries = {
  friends: {
    label: "247프렌즈",
    series: [
      { label: "1기", count: 16 },
      { label: "2기", count: 17 },
      { label: "3기", count: 20 },
      { label: "4기", count: 22 },
      { label: "5기", count: 28 },
      { label: "6기", count: 28 },
      { label: "7기", count: 37 },
      { label: "8기", count: 33 },
      { label: "9기", count: 28 },
      { label: "10기", count: 25 },
      { label: "11기", count: 35 },
      { label: "12기", count: 26 },
      { label: "13기", count: 20 },
      { label: "윈터스쿨", count: 9 }
    ]
  },
  trial: {
    label: "247체험단",
    series: [
      { label: "1차", count: 32 },
      { label: "2차", count: 15 },
      { label: "3차", count: 22 },
      { label: "4차", count: 15 },
      { label: "5차", count: 18 },
      { label: "6차", count: 14 }
    ]
  },
  exam: {
    label: "입시교육",
    series: [
      { label: "23수시", count: 12 },
      { label: "23정시", count: 43 },
      { label: "24수시", count: 10 },
      { label: "24정시", count: 27 },
      { label: "25수시", count: 44 },
      { label: "25정시", count: 49 },
      { label: "26수시", count: 28 }
    ]
  },
  marketing: {
    label: "마케팅교육",
    series: [
      { label: "SNS1", count: 7 },
      { label: "SNS2", count: 4 },
      { label: "영상1", count: 13 },
      { label: "영상2", count: 7 }
    ]
  },
  press: {
    label: "언론보도",
    series: [
      { label: "2023", count: 28 },
      { label: "2024", count: 26 },
      { label: "2025", count: 9 }
    ]
  },
  event: {
    label: "협업이벤트",
    series: [
      { label: "사우어코코", count: 31 },
      { label: "양파링", count: 7 },
      { label: "오션믹스", count: 5 },
      { label: "린트", count: 4 },
      { label: "보노스프", count: 27 },
      { label: "스키틀즈", count: 24 },
      { label: "이클립스", count: 34 },
      { label: "퓨어포커스", count: 16 }
    ]
  }
};

const rawBranches = [
  ["서울","강남","o","o","o","x","x","o"],["","강북","x","x","o","x","o","o"],["","노량진","x","o","o","x","o","x"],["","마포","o","x","o","o","o","o"],["","목동","o","o","o","o","o","o"],["","목동오목교","o","o","o","x","x","o"],["","서울강동","o","x","o","x","o","o"],["","서울강서","o","o","o","x","o","x"],["","서울광진","o","o","o","x","o","o"],["","서울대점","x","x","o","o","o","o"],["","서울도봉","o","o","o","o","o","o"],["","서울성동","o","x","o","x","o","o"],["","서울성북","o","o","o","o","x","o"],["","서울송파","o","x","o","o","o","x"],["","은평서대문","o","o","o","x","x","o"],["경기","광명","o","x","o","x","o","o"],["","구리남양주","o","o","o","x","o","o"],["","김포","o","x","o","x","o","o"],["","동탄","o","o","o","x","x","o"],["","부천","x","o","o","x","o","x"],["","분당정자","o","o","o","x","x","o"],["","수원시청","o","o","o","x","x","x"],["","수원영통","o","o","o","x","x","o"],["","수원정자","o","o","o","x","o","o"],["","안산","o","o","o","x","x","o"],["","용인수지","x","x","o","x","x","o"],["","의정부","o","o","o","x","o","o"],["","일산동구","o","o","o","o","o","o"],["","일산서구","o","o","o","x","o","o"],["","일산화정","o","x","o","x","x","x"],["","파주","o","o","o","x","o","o"],["","평택","o","x","o","x","x","o"],["","하남","o","x","o","x","x","o"],["인천","인천부평","o","o","o","x","x","x"],["","인천송도","o","o","o","o","o","o"],["","인천청라","o","x","o","o","o","x"],["강원","원주","o","o","o","x","x","o"],["","춘천","o","o","o","x","o","o"],["충청","대전둔산","o","x","o","x","o","o"],["","천안","o","x","o","x","o","o"],["","청주","x","x","o","x","x","x"],["전라","광주남구","o","o","o","o","x","x"],["","광주동구","o","o","o","o","x","o"],["","광주북구","o","o","o","x","o","o"],["","광주수완","o","o","o","x","x","o"],["","목포","x","x","o","x","x","x"],["","익산","o","x","o","x","x","o"],["경상","대구달서","o","x","o","x","x","o"],["","대구수성 1관","o","x","o","x","x","x"],["","대구수성 2관","o","x","o","x","o","o"],["","부산교대","o","o","o","x","o","o"],["","부산대","o","o","o","x","o","o"],["","부산북구","o","x","o","x","o","o"],["","부산서면","o","x","o","x","o","o"],["","부산해운대","o","x","o","x","x","o"],["","울산남구","x","x","o","x","x","x"],["","진주","x","o","o","x","o","o"],["","창원","x","x","o","x","o","o"],["제주","제주","x","x","o","x","o","x"],["기숙","안성기숙","x","x","o","x","x","o"],["","이천기숙","o","o","o","x","o","o"],["","독학기숙","o","o","o","x","o","o"]
];

const actionTemplates = {
  friends: [
    ["프렌즈 저활성 지점 재모집", "프렌즈 미운영 지점에 1:1 우수사례 공유와 리마인드 운영"],
    ["프렌즈 후기 회수 강화", "활동 완료 지점에서 후기 확보 후 SNS/랜딩 재활용"],
    ["오프라인 설명회 연계", "정시·수시 교육과 프렌즈 모집 일정을 연결"],
    ["서울권 우수 전파", "서울 다운영 지점 사례를 경기·전라권에 이식"]
  ],
  trial: [
    ["체험단 추가 모집", "미운영 지점 대상 차기 차수 모집 배너와 문자 발송"],
    ["체험 후 전환 시나리오 정비", "체험 종료 후 상담 연결 스크립트 재정비"],
    ["우수 지점 확장", "운영 지점의 광고 세트를 비운영 지점으로 복제"],
    ["반응 점검", "권역별 체험단 모집 랜딩 이탈 구간 점검"]
  ],
  exam: [
    ["입시교육 재홍보", "설명회 일정과 녹화본을 미참여 지점에 재배포"],
    ["정시교육 후속 상담", "참여 지점의 학부모 상담 예약률 추적"],
    ["수시교육 패키지화", "수시/정시 교육을 묶은 지점별 제안서 구성"],
    ["온라인 전환 확대", "현장 참석 어려운 지점에 온라인형 재제공"]
  ],
  marketing: [
    ["마케팅교육 참여 확대", "영상/SNS 교육 미참여 지점 대상 재교육 공지"],
    ["직영점 운영 노하우 전파", "SNS 2차 직영점 사례를 지역 지점에 전달"],
    ["교육 자료 표준화", "촬영 가이드와 카드 템플릿 통합"],
    ["성과 리뷰", "교육 수강 후 게시물 발행 건수 추적"]
  ],
  press: [
    ["언론보도 소재 확보", "원장 인터뷰와 성적 향상 사례 취합"],
    ["권역별 미노출 지점 보강", "미운영 지점에 보도자료 작성 지원"],
    ["연간 기사 캘린더", "2026년 분기별 보도 아이템 선제 설계"],
    ["성과 아카이브 정리", "기존 보도 페이지와 랜딩 연결 개선"]
  ],
  event: [
    ["협업 이벤트 재배치", "브랜드별 반응을 기준으로 지점 재배분"],
    ["저활성 지점 샘플 발송", "미운영 지점에 체험 키트/이벤트 물량 우선 배정"],
    ["후기 콘텐츠 회수", "협업 상품 리뷰를 SNS/블로그로 재가공"],
    ["재계약 후보 발굴", "반응 우수 브랜드를 다음 분기 협업 후보로 선정"]
  ],
  all: [
    ["프로그램 믹스 보완", "운영 프로그램 수가 적은 지점에 우선 프로그램 제안"],
    ["권역별 저활성 관리", "운영 비율이 낮은 권역 중심으로 액션 플랜 점검"],
    ["우수 지점 사례 공유", "6개 프로그램 동시 운영 지점을 베스트 프랙티스로 정리"],
    ["미운영 지점 집중 점검", "O/X 기준 2개 이하 운영 지점을 우선 관리군으로 묶기"]
  ]
};

const branches = rawBranches.map((row) => ({
  region: row[0],
  branch: row[1],
  friends: row[2],
  trial: row[3],
  exam: row[4],
  marketing: row[5],
  press: row[6],
  event: row[7]
}));

let currentRegion = "all";
let currentProgram = "friends";
let currentStatus = "all";

const elements = {
  tabs: [...document.querySelectorAll(".sheet-tab")],
  chips: [...document.querySelectorAll(".program-chip")],
  regionFilter: document.getElementById("region-filter"),
  programFilter: document.getElementById("program-filter"),
  statusFilter: document.getElementById("status-filter"),
  resetButton: document.getElementById("reset-button"),
  currentViewLabel: document.getElementById("current-view-label"),
  summaryBadge: document.getElementById("summary-badge"),
  activeBranches: document.getElementById("active-branches"),
  totalBranches: document.getElementById("total-branches"),
  activeRate: document.getElementById("active-rate"),
  activeRateSub: document.getElementById("active-rate-sub"),
  activeBranchesSub: document.getElementById("active-branches-sub"),
  inactiveBranches: document.getElementById("inactive-branches"),
  summaryHeadRow: document.getElementById("summary-head-row"),
  summaryCountRow: document.getElementById("summary-count-row"),
  summaryShareRow: document.getElementById("summary-share-row"),
  trendNote: document.getElementById("trend-note"),
  trendCards: document.getElementById("trend-cards"),
  programBars: document.getElementById("program-bars"),
  heatCells: document.getElementById("heat-cells"),
  actionList: document.getElementById("action-list"),
  branchTable: document.getElementById("branch-table"),
  planTable: document.getElementById("plan-table")
};

function fillRegions(data) {
  let region = "";
  return data.map((item) => {
    if (item.region) region = item.region;
    return { ...item, region };
  });
}

const normalizedBranches = fillRegions(branches);
const regions = ["all", ...new Set(normalizedBranches.map((item) => item.region))];

function populateFilters() {
  elements.regionFilter.innerHTML = regions
    .map((region) => `<option value="${region}">${region === "all" ? "전체 권역" : region}</option>`)
    .join("");

  elements.programFilter.innerHTML = Object.entries(programMeta)
    .map(([value, meta]) => `<option value="${value}">${meta.label}</option>`)
    .join("");

  elements.regionFilter.value = currentRegion;
  elements.programFilter.value = currentProgram;
}

function isActiveForProgram(branch, program) {
  if (program === "all") {
    return ["friends", "trial", "exam", "marketing", "press", "event"].some((key) => branch[key] === "o");
  }
  return branch[programMeta[program].key] === "o";
}

function filterBranches() {
  return normalizedBranches.filter((branch) => {
    const regionMatch = currentRegion === "all" || branch.region === currentRegion;
    const statusMatch = currentStatus === "all"
      || (currentStatus === "active" && isActiveForProgram(branch, currentProgram))
      || (currentStatus === "inactive" && !isActiveForProgram(branch, currentProgram));
    return regionMatch && statusMatch;
  });
}

function getProgramCounts(data) {
  return ["friends", "trial", "exam", "marketing", "press", "event"].map((key) => ({
    key,
    label: programMeta[key].label,
    count: data.filter((item) => item[key] === "o").length
  }));
}

function renderSummary(data) {
  const total = data.length;
  const active = data.filter((item) => isActiveForProgram(item, currentProgram)).length;
  const inactive = total - active;
  const rate = total ? `${((active / total) * 100).toFixed(1)}%` : "0%";

  elements.activeBranches.textContent = active;
  elements.totalBranches.textContent = total;
  elements.activeRate.textContent = rate;
  elements.inactiveBranches.textContent = inactive;
  elements.activeBranchesSub.textContent = `${programMeta[currentProgram].label} 기준`;
  elements.activeRateSub.textContent = `${currentRegion === "all" ? "전체 권역" : currentRegion} 기준`;
  elements.currentViewLabel.textContent = `${programMeta[currentProgram].label} / ${currentRegion === "all" ? "전체 권역" : currentRegion}`;
  elements.summaryBadge.textContent = currentStatus === "all" ? "실데이터 기준" : currentStatus === "active" ? "운영중만 표시" : "미운영만 표시";
}

function renderCycleSummary(data) {
  const program = currentProgram === "all" ? "friends" : currentProgram;
  const series = workbookSeries[program].series;
  const totalBranches = data.length || 1;
  const lastFour = series.slice(-4);

  elements.summaryHeadRow.innerHTML = ["<th>구분</th>", ...lastFour.map((item) => `<th>${item.label}</th>`)].join("");
  elements.summaryCountRow.innerHTML = ["<th>참여 지점 수</th>", ...lastFour.map((item) => `<td class="cell-emphasis">${item.count}</td>`)].join("");
  elements.summaryShareRow.innerHTML = ["<th>참여 비율</th>", ...lastFour.map((item) => `<td class="cell-good">${((item.count / totalBranches) * 100).toFixed(1)}%</td>`)].join("");
}

function renderTrendCards() {
  const entries = currentProgram === "all"
    ? Object.entries(workbookSeries)
    : [[currentProgram, workbookSeries[currentProgram]]];

  elements.trendNote.textContent = currentProgram === "all"
    ? "프로그램별 실제 회차 참여 지점 수 비교"
    : `${programMeta[currentProgram].label} 시트의 실제 회차 참여 지점 수`;

  elements.trendCards.innerHTML = entries.map(([key, info], index) => {
    const max = Math.max(...info.series.map((item) => item.count), 1);
    return `
      <article class="trend-card ${index === 0 ? "featured" : ""}">
        <div class="trend-head">
          <span>${info.label}</span>
          <span>${info.series.reduce((sum, item) => sum + item.count, 0)}회 누적</span>
        </div>
        <div class="trend-body">
          <div class="sparkline-row">
            ${info.series.map((item) => `<div class="spark-bar" style="height:${(item.count / max) * 100}%" title="${item.label}: ${item.count}"></div>`).join("")}
          </div>
          <div class="spark-labels">
            ${info.series.map((item) => `<span title="${item.label}">${item.label}</span>`).join("")}
          </div>
        </div>
      </article>
    `;
  }).join("");
}

function renderProgramBars(data) {
  const counts = getProgramCounts(data);
  const max = Math.max(...counts.map((item) => item.count), 1);
  elements.programBars.innerHTML = counts.map((item) => `
    <div class="bar-row">
      <div class="bar-meta">
        <strong>${item.label}</strong>
        <span class="bar-value">${item.count}개</span>
      </div>
      <div class="bar-track">
        <div class="bar-fill" style="width:${(item.count / max) * 100}%"></div>
      </div>
    </div>
  `).join("");
}

function levelByRate(rate) {
  if (rate >= 0.8) return 4;
  if (rate >= 0.6) return 3;
  if (rate >= 0.4) return 2;
  return 1;
}

function renderHeat(data) {
  const regionMap = new Map();
  data.forEach((item) => {
    if (!regionMap.has(item.region)) regionMap.set(item.region, []);
    regionMap.get(item.region).push(item);
  });

  elements.heatCells.innerHTML = [...regionMap.entries()].map(([region, items]) => {
    const active = items.filter((item) => isActiveForProgram(item, currentProgram)).length;
    const rate = items.length ? active / items.length : 0;
    return `
      <div class="heat-cell level-${levelByRate(rate)}">
        <div class="heat-cell-head">
          <strong>${region}</strong>
          <span>${(rate * 100).toFixed(1)}%</span>
        </div>
        <p>운영 ${active} / 전체 ${items.length}</p>
      </div>
    `;
  }).join("");
}

function renderActions(data) {
  const items = actionTemplates[currentProgram] || actionTemplates.all;
  const inactive = data.filter((item) => !isActiveForProgram(item, currentProgram)).slice(0, 5).map((item) => item.branch);
  elements.actionList.innerHTML = items.map((item, index) => `
    <li class="action-item">
      <div class="action-index">${index + 1}</div>
      <div class="action-copy">
        <strong>${item[0]}</strong>
        <p>${item[1]}${inactive.length && index === 0 ? ` / 우선 검토: ${inactive.join(", ")}` : ""}</p>
      </div>
    </li>
  `).join("");
}

function markBadge(value) {
  const active = value === "o";
  return `<span class="badge-cell ${active ? "on" : "off"}">${active ? "O" : "X"}</span>`;
}

function renderBranches(data) {
  elements.branchTable.innerHTML = data.map((item) => `
    <tr>
      <td>${item.region}</td>
      <td><strong>${item.branch}</strong></td>
      <td>${markBadge(item.friends)}</td>
      <td>${markBadge(item.trial)}</td>
      <td>${markBadge(item.exam)}</td>
      <td>${markBadge(item.marketing)}</td>
      <td>${markBadge(item.press)}</td>
      <td>${markBadge(item.event)}</td>
    </tr>
  `).join("");
}

function renderPlans(data) {
  const inactiveCount = data.filter((item) => !isActiveForProgram(item, currentProgram)).length;
  const plans = [
    { area: "CRM", task: `${programMeta[currentProgram].label} 미운영 지점 재접촉`, owner: "김다은", due: "04/18", progress: Math.max(20, 100 - inactiveCount * 3) },
    { area: "콘텐츠", task: `${programMeta[currentProgram].label} 우수 사례 카드뉴스 제작`, owner: "박은호", due: "04/20", progress: 62 },
    { area: "이벤트", task: `${currentRegion === "all" ? "권역별" : currentRegion} 집중 운영안 정리`, owner: "이수민", due: "04/22", progress: 54 },
    { area: "현장지원", task: "저성과 지점 코칭 및 스크립트 공유", owner: "정예린", due: "04/24", progress: 41 },
    { area: "광고", task: `${programMeta[currentProgram].label} 집행 세트 재배분`, owner: "최민준", due: "04/17", progress: 86 }
  ];

  elements.planTable.innerHTML = plans.map((item) => `
    <tr>
      <td>${item.area}</td>
      <td>${item.task}</td>
      <td>${item.owner}</td>
      <td>${item.due}</td>
      <td class="progress-cell">
        <div class="progress-track">
          <div class="progress-fill" style="width:${item.progress}%">${item.progress}%</div>
        </div>
      </td>
    </tr>
  `).join("");
}

function syncControls() {
  elements.regionFilter.value = currentRegion;
  elements.programFilter.value = currentProgram;
  elements.statusFilter.value = currentStatus;

  elements.tabs.forEach((button) => button.classList.toggle("active", button.dataset.program === currentProgram || (currentProgram === "all" && button.dataset.program === "all")));
  elements.chips.forEach((button) => button.classList.toggle("active", button.dataset.program === currentProgram));
}

function render() {
  syncControls();
  const filtered = filterBranches();
  renderSummary(filtered);
  renderCycleSummary(filtered);
  renderTrendCards();
  renderProgramBars(filtered);
  renderHeat(filtered);
  renderActions(filtered);
  renderBranches(filtered);
  renderPlans(filtered);
}

function setProgram(program) {
  currentProgram = program;
  render();
}

elements.tabs.forEach((button) => {
  button.addEventListener("click", () => setProgram(button.dataset.program));
});

elements.chips.forEach((button) => {
  button.addEventListener("click", () => setProgram(button.dataset.program));
});

elements.regionFilter.addEventListener("change", (event) => {
  currentRegion = event.target.value;
  render();
});

elements.programFilter.addEventListener("change", (event) => {
  currentProgram = event.target.value;
  render();
});

elements.statusFilter.addEventListener("change", (event) => {
  currentStatus = event.target.value;
  render();
});

elements.resetButton.addEventListener("click", () => {
  currentRegion = "all";
  currentProgram = "friends";
  currentStatus = "all";
  render();
});

populateFilters();
render();
