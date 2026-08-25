(() => {
  const params = new URLSearchParams(window.location.hash.slice(1));
  const token = params.get("etoos247_collect");
  const dashboardOrigin = params.get("etoos247_origin");
  if (!token || !dashboardOrigin) return;

  function isAllowedDashboardOrigin(rawOrigin) {
    try {
      const parsed = new URL(rawOrigin);
      return parsed.protocol === "https:" && (
        parsed.hostname === "marketing-dashboard-sable.vercel.app" ||
        (parsed.hostname.endsWith(".vercel.app") && parsed.hostname.startsWith("marketing-dashboard"))
      );
    } catch {
      return false;
    }
  }

  if (!isAllowedDashboardOrigin(dashboardOrigin)) return;

  const badge = document.createElement("div");
  Object.assign(badge.style, {
    position: "fixed",
    top: "18px",
    left: "50%",
    transform: "translateX(-50%)",
    zIndex: "2147483647",
    padding: "12px 18px",
    borderRadius: "999px",
    background: "rgba(15, 23, 42, 0.94)",
    color: "#fff",
    font: "700 14px/1.4 -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif",
    boxShadow: "0 12px 30px rgba(0,0,0,.28)"
  });
  badge.textContent = "이투스247 대시보드: 최신 게시물을 불러오는 중…";
  document.documentElement.appendChild(badge);

  const normalizeUrl = (rawUrl) => {
    try {
      const parsed = new URL(rawUrl, window.location.origin);
      parsed.hash = "";
      parsed.search = "";
      return parsed.toString();
    } catch {
      return "";
    }
  };

  const extractDate = (altText) => {
    const englishDate = String(altText || "").match(/\b([A-Za-z]+ \d{1,2}, \d{4})\b/i);
    if (englishDate) {
      const date = new Date(englishDate[1]);
      if (!Number.isNaN(date.getTime())) return date.toISOString().slice(0, 10);
    }
    const koreanDate = String(altText || "").match(/(20\d{2})년\s*(\d{1,2})월\s*(\d{1,2})일/);
    if (koreanDate) {
      return `${koreanDate[1]}-${koreanDate[2].padStart(2, "0")}-${koreanDate[3].padStart(2, "0")}`;
    }
    return "";
  };

  const parseCompactCount = (rawValue) => {
    const normalized = String(rawValue || "").replace(/,/g, "").replace(/\s+/g, "").toLowerCase();
    const match = normalized.match(/([\d.]+)(천|만|k|m|b)?/i);
    if (!match) return null;
    const multipliers = { "천": 1000, "만": 10000, k: 1000, m: 1000000, b: 1000000000 };
    const number = Number(match[1]) * (multipliers[match[2]?.toLowerCase()] || 1);
    return Number.isFinite(number) ? Math.round(number) : null;
  };

  const sortByPublishedAt = (posts) => posts
    .map((post, sourceIndex) => ({ ...post, sourceIndex }))
    .sort((a, b) => {
      if (a.publishedAt && b.publishedAt && a.publishedAt !== b.publishedAt) {
        return b.publishedAt.localeCompare(a.publishedAt);
      }
      if (a.publishedAt && !b.publishedAt) return -1;
      if (!a.publishedAt && b.publishedAt) return 1;
      return a.sourceIndex - b.sourceIndex;
    })
    .map(({ sourceIndex, ...post }) => post);

  function collectPosts() {
    const links = Array.from(document.querySelectorAll('main a[href*="/p/"], main a[href*="/reel/"]'));
    const seen = new Set();
    return links.reduce((posts, link) => {
      if (posts.length >= 60) return posts;
      const url = normalizeUrl(link.href);
      if (!url || seen.has(url)) return posts;

      const image = link.querySelector("img");
      if (!image?.src) return posts;
      seen.add(url);

      const caption = (image.alt || link.getAttribute("aria-label") || "").replace(/\s+/g, " ").trim();
      const carouselIcon = link.querySelector('svg[aria-label*="Carousel"], svg[aria-label*="슬라이드"], svg[aria-label*="여러"]');
      posts.push({
        url,
        thumbnailUrl: image.currentSrc || image.src,
        caption,
        publishedAt: extractDate(caption),
        type: url.includes("/reel/") ? "reel" : carouselIcon ? "carousel" : "image"
      });
      return posts;
    }, []);
  }

  async function enrichPost(post) {
    try {
      const controller = new AbortController();
      const timeout = window.setTimeout(() => controller.abort(), 10000);
      const response = await fetch(post.url, {
        credentials: "include",
        cache: "no-store",
        signal: controller.signal
      });
      window.clearTimeout(timeout);
      if (!response.ok) return post;

      const html = await response.text();
      const documentCopy = new DOMParser().parseFromString(html, "text/html");
      const description = documentCopy.querySelector('meta[property="og:description"]')?.content || "";
      const publishedMeta = documentCopy.querySelector('meta[property="article:published_time"]')?.content || "";
      const englishEngagement = description.match(/([\d,.]+(?:\s?[KMB])?)\s+likes?,\s+([\d,.]+(?:\s?[KMB])?)\s+comments?/i);
      const koreanEngagement = description.match(/좋아요\s*([\d,.]+(?:\s?[천만])?)개?.*?댓글\s*([\d,.]+(?:\s?[천만])?)개?/i);
      const engagement = englishEngagement || koreanEngagement;

      return {
        ...post,
        publishedAt: publishedMeta ? publishedMeta.slice(0, 10) : extractDate(description) || post.publishedAt,
        likes: engagement ? parseCompactCount(engagement[1]) : null,
        comments: engagement ? parseCompactCount(engagement[2]) : null
      };
    } catch {
      return post;
    }
  }

  async function enrichLatestPosts(posts) {
    const targets = sortByPublishedAt(posts).slice(0, 12);
    const enriched = [];
    for (let index = 0; index < targets.length; index += 3) {
      badge.textContent = `최신 게시물 반응 수 확인 중… (${Math.min(index + 3, targets.length)}/${targets.length})`;
      enriched.push(...await Promise.all(targets.slice(index, index + 3).map(enrichPost)));
      await new Promise((resolve) => window.setTimeout(resolve, 300));
    }
    const enrichedByUrl = new Map(enriched.map((post) => [post.url, post]));
    return sortByPublishedAt(posts.map((post) => enrichedByUrl.get(post.url) || post));
  }

  async function sendCollection(posts, coverageComplete) {
    const response = await fetch(`${dashboardOrigin}/api/instagram-collection`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "complete", token, posts, coverageComplete })
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || "대시보드 저장 실패");
    return result;
  }

  async function runCollection() {
    let posts = [];
    const collectedByUrl = new Map();
    let previousCount = 0;
    let stableRounds = 0;

    for (let attempt = 0; attempt < 18; attempt += 1) {
      collectPosts().forEach((post) => collectedByUrl.set(post.url, post));
      posts = Array.from(collectedByUrl.values()).slice(0, 60);
      badge.textContent = `최근 30일 게시물 확인 중… (${posts.length}개 발견)`;
      stableRounds = posts.length === previousCount ? stableRounds + 1 : 0;
      previousCount = posts.length;

      if (posts.length >= 60 || (posts.length > 0 && stableRounds >= 3)) break;
      window.scrollTo({ top: document.body.scrollHeight, behavior: "smooth" });
      await new Promise((resolve) => window.setTimeout(resolve, 1100));
    }

    if (!posts.length) {
      badge.style.background = "rgba(217, 119, 6, 0.96)";
      badge.textContent = "공개 게시물을 찾지 못했습니다. 로그인 상태와 계정 공개 여부를 확인해 주세요.";
      return;
    }

    try {
      const enrichedPosts = await enrichLatestPosts(posts);
      const koreaToday = new Date(Date.now() + 9 * 60 * 60 * 1000).toISOString().slice(0, 10);
      const cutoffDate = new Date(`${koreaToday}T00:00:00.000Z`);
      cutoffDate.setUTCDate(cutoffDate.getUTCDate() - 29);
      const cutoff = cutoffDate.toISOString().slice(0, 10);
      const datedPosts = enrichedPosts.filter((post) => post.publishedAt);
      const oldestPublishedAt = datedPosts.at(-1)?.publishedAt || "";
      const isLoggedOut = Boolean(document.querySelector('a[href*="/accounts/login"], form[action*="/accounts/login"]')) ||
        Array.from(document.querySelectorAll("button")).some((button) => ["로그인", "log in"].includes(button.textContent.trim().toLowerCase()));
      const coverageComplete = !isLoggedOut && Boolean(oldestPublishedAt) && oldestPublishedAt < cutoff;
      const result = await sendCollection(enrichedPosts, coverageComplete);
      badge.style.background = "rgba(5, 150, 105, 0.96)";
      badge.textContent = coverageComplete
        ? `수집 완료: ${result.collection.recent30d}개(최근 30일), 최신 ${Math.min(enrichedPosts.length, 12)}개 반응 수 반영`
        : `게시물 ${enrichedPosts.length}개 수집 완료 · 로그인하면 30일 지표도 자동 최신화됩니다.`;
      window.setTimeout(() => badge.remove(), 6000);
    } catch (error) {
      badge.style.background = "rgba(220, 38, 38, 0.96)";
      badge.textContent = `수집 실패: ${error.message}`;
    }
  }

  runCollection();
})();
