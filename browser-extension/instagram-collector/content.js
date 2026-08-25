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
    const englishDate = String(altText || "").match(/\bon ([A-Za-z]+ \d{1,2}, \d{4})\b/i);
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

  function collectPosts() {
    const links = Array.from(document.querySelectorAll('a[href*="/p/"], a[href*="/reel/"]'));
    const seen = new Set();
    return links.reduce((posts, link) => {
      if (posts.length >= 12) return posts;
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

  async function sendCollection(posts) {
    const response = await fetch(`${dashboardOrigin}/api/instagram-collection`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "complete", token, posts })
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || "대시보드 저장 실패");
    return result;
  }

  let attempts = 0;
  const timer = window.setInterval(async () => {
    attempts += 1;
    const posts = collectPosts();
    if (posts.length >= 3 || (posts.length > 0 && attempts >= 8)) {
      window.clearInterval(timer);
      try {
        await sendCollection(posts);
        badge.style.background = "rgba(5, 150, 105, 0.96)";
        badge.textContent = `수집 완료: 최신 게시물 ${posts.length}개를 대시보드에 반영했습니다.`;
        window.setTimeout(() => badge.remove(), 5000);
      } catch (error) {
        badge.style.background = "rgba(220, 38, 38, 0.96)";
        badge.textContent = `수집 실패: ${error.message}`;
      }
    } else if (attempts >= 30) {
      window.clearInterval(timer);
      badge.style.background = "rgba(217, 119, 6, 0.96)";
      badge.textContent = "공개 게시물을 찾지 못했습니다. 로그인 상태와 계정 공개 여부를 확인해 주세요.";
    }
  }, 1000);
})();
