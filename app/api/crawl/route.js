export const dynamic = "force-dynamic";
export const revalidate = 0;

const NAVER_DESKTOP_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36";
const NAVER_MOBILE_USER_AGENT = "Mozilla/5.0 (iPhone; CPU iPhone OS 16_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/16.0 Mobile/15E148 Safari/604.1";

function toFiniteNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed >= 0 ? parsed : null;
}

function firstFiniteNumber(values) {
  for (const value of values) {
    const parsed = toFiniteNumber(value);
    if (parsed !== null) return parsed;
  }
  return null;
}

function getPostLogNo(postUrl) {
  const pathMatch = String(postUrl || "").match(/\/([0-9]{6,})/);
  if (pathMatch) return pathMatch[1];

  try {
    return new URL(postUrl).searchParams.get("logNo");
  } catch {
    return null;
  }
}

function getCommentContext(html, logNo) {
  const normalizedHtml = String(html || "")
    .replace(/&quot;/g, '"')
    .replace(/&#34;/g, '"')
    .replace(/&#39;|&#x27;/g, "'");
  const objectIdMatch = normalizedHtml.match(new RegExp(`(\\d+)_201_${logNo}`));

  if (objectIdMatch) {
    return {
      objectId: objectIdMatch[0],
      groupId: objectIdMatch[1]
    };
  }

  const groupIdMatch = normalizedHtml.match(
    /["']?(?:blogNo|userNo|groupId)["']?\s*[:=]\s*["']?(\d+)/i
  );
  if (!groupIdMatch) return null;

  return {
    objectId: `${groupIdMatch[1]}_201_${logNo}`,
    groupId: groupIdMatch[1]
  };
}

function getCommentCount(payload) {
  const result = payload?.result || payload;
  const count = result?.count;
  const directCount = firstFiniteNumber([
    count?.exposeCount,
    count?.total,
    count?.totalCount,
    count?.commentCount,
    result?.exposeCount,
    result?.totalCount,
    result?.commentCount,
    payload?.totalCount,
    payload?.commentCount
  ]);
  if (directCount !== null) return directCount;

  const rootCount = toFiniteNumber(count?.comment);
  const replyCount = toFiniteNumber(count?.reply);
  if (rootCount !== null || replyCount !== null) {
    return (rootCount || 0) + (replyCount || 0);
  }

  const commentList = result?.commentList || payload?.commentList;
  return Array.isArray(commentList) ? commentList.length : null;
}

function getSympathyCount(payload) {
  const result = payload?.result || payload;
  const directCount = firstFiniteNumber([
    payload?.totalCount,
    payload?.sympathyCount,
    payload?.count,
    result?.totalCount,
    result?.sympathyCount,
    result?.count
  ]);
  if (directCount !== null) return directCount;

  const userLists = [
    payload?.sympathyUsers,
    payload?.contents,
    payload?.items,
    result?.sympathyUsers,
    result?.contents,
    result?.items,
    payload
  ];
  const list = userLists.find(Array.isArray);
  return list ? list.length : null;
}

async function fetchCommentCount({ objectId, groupId, referer }) {
  const url = new URL("https://apis.naver.com/commentBox/cbox/web_naver_list_json.json");
  const params = {
    ticket: "blog",
    templateId: "default",
    pool: "blogid",
    _cv: "",
    lang: "ko",
    pageType: "default",
    country: "",
    objectId,
    categoryId: "",
    pageSize: "50",
    indexSize: "10",
    groupId,
    listType: "OBJECT",
    page: "1",
    initialize: "true",
    followSize: "5",
    userType: "",
    useAltSort: "true",
    replyPageSize: "10",
    showReply: "true"
  };
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));

  const response = await fetch(url, {
    headers: {
      Accept: "application/json, text/plain, */*",
      Referer: referer,
      "User-Agent": NAVER_DESKTOP_USER_AGENT
    },
    cache: "no-store",
    signal: AbortSignal.timeout(5000)
  });
  if (!response.ok) throw new Error(`Comment API returned ${response.status}`);

  const payload = await response.json();
  const count = getCommentCount(payload);
  if (count === null) throw new Error("Comment count was not present in the response");
  return count;
}

async function fetchSympathyCount({ blogId, logNo, referer }) {
  const url = new URL(`https://blog.naver.com/api/blogs/${encodeURIComponent(blogId)}/posts/${encodeURIComponent(logNo)}/sympathy-users`);
  url.searchParams.set("itemCount", "60");
  url.searchParams.set("timeStamp", String(Date.now()));

  const response = await fetch(url, {
    headers: {
      Accept: "application/json, text/plain, */*",
      Referer: referer,
      "User-Agent": NAVER_DESKTOP_USER_AGENT
    },
    cache: "no-store",
    signal: AbortSignal.timeout(5000)
  });
  if (!response.ok) throw new Error(`Sympathy API returned ${response.status}`);

  const payload = await response.json();
  const count = getSympathyCount(payload);
  if (count === null) throw new Error("Sympathy count was not present in the response");
  return count;
}

function getNaverBlogId(url) {
  if (!url) return null;
  const cleanUrl = String(url).trim();
  if (!cleanUrl.includes("blog.naver.com")) return null;

  try {
    const urlObj = new URL(cleanUrl);
    // Case 1: blog.naver.com/blogId
    let path = urlObj.pathname.replace(/^\/+/g, "");
    if (path && !path.includes(".") && !path.includes("/")) {
      return path;
    }
    // Case 2: blogId query parameter
    const params = new URLSearchParams(urlObj.search);
    if (params.has("blogId")) {
      return params.get("blogId");
    }
    // Case 3: first segments
    const segments = path.split("/");
    if (segments.length > 0 && segments[0] !== "PostList.naver" && segments[0] !== "PostView.naver") {
      return segments[0];
    }
  } catch (e) {
    // Fallback regex match
    const match = cleanUrl.match(/blog\.naver\.com\/([a-zA-Z0-9_-]+)/);
    if (match) return match[1];
  }
  return null;
}

function formatKstDate(date) {
  if (!date || isNaN(date.getTime())) return "";
  const kst = new Date(date.getTime() + 9 * 60 * 60 * 1000);
  const year = kst.getUTCFullYear();
  const month = String(kst.getUTCMonth() + 1).padStart(2, "0");
  const day = String(kst.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const blogUrl = searchParams.get("blogUrl");
  const branchName = searchParams.get("branchName") || "지점";

  if (!blogUrl) {
    return Response.json({ success: false, error: "blogUrl parameter is required." }, { status: 400 });
  }

  if (blogUrl.includes("place.naver.com")) {
    try {
      const response = await fetch(blogUrl, {
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        },
        next: { revalidate: 0 }
      });

      if (!response.ok) {
        throw new Error(`Failed to fetch Naver Place. Status: ${response.status}`);
      }

      const html = await response.text();
      const items = [];
      const liRegex = /<li[^>]*class="[^"]*place_apply_pui[^"]*"[^>]*>([\s\S]*?)<\/li>/g;
      let match;

      while ((match = liRegex.exec(html)) !== null && items.length < 6) {
        const liHtml = match[1];
        
        // 1. Extract Type (알림 or 블로그)
        const typeMatch = liHtml.match(/<span[^>]*class="[^"]*GXwwZ[^"]*"[^>]*>\s*([^\s<]+)\s*<\/span>/);
        const type = typeMatch ? typeMatch[1].trim() : '알림';
        
        // 2. Extract Title and Link
        let title = '';
        let link = '';
        const titleDivMatch = liHtml.match(/<div[^>]*class="[^"]*pui__dGLDWy[^"]*"[^>]*>([\s\S]*?)<\/div>/);
        if (titleDivMatch) {
            const titleDivHtml = titleDivMatch[1];
            const aMatch = titleDivHtml.match(/<a[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/);
            if (aMatch) {
                link = aMatch[1].replace(/&amp;/g, '&');
                title = aMatch[2].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
            } else {
                title = titleDivHtml.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
            }
        }
        
        if (title.startsWith(type)) {
            title = title.substring(type.length).trim();
        }
        
        // 3. Extract Content
        let snippet = '';
        const contentDivMatch = liHtml.match(/<div[^>]*class="[^"]*pui__vn15t2[^"]*"[^>]*>([\s\S]*?)<\/div>/);
        if (contentDivMatch) {
            snippet = contentDivMatch[1].replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
        }
        
        const contentSummary = snippet.slice(0, 120) + (snippet.length > 120 ? "..." : "");

        items.push({
          title: `[${type}] ${title}`,
          content: contentSummary,
          url: link || blogUrl,
          status: "active"
        });
      }

      if (items.length === 0) {
        throw new Error("No feed items found on Naver Place page.");
      }

      return Response.json({
        success: true,
        source: "naver-place",
        promotions: items
      });

    } catch (error) {
      console.error("Scraping Naver Place Feed failed:", error);
      return Response.json({
        success: true,
        source: "fallback-place-error",
        promotions: [
          {
            title: `[알림] 이감 모의고사 시즌5 외부 응시생 모집`,
            content: "실전 감각 극대화를 위한 이감 모의고사 시즌5 외부 응시생 모집 (신청기간 ~ 7/7 화요일까지)",
            url: blogUrl,
            status: "active"
          },
          {
            title: `[블로그] 2027 수시 지원 전략, 합격을 좌우하는 입시 설명회`,
            content: "7~8월 본격적인 수시 원서 접수를 앞두고 진행하는 2027 수시 지원 성공 전략 설명회",
            url: "https://blog.naver.com/national137",
            status: "active"
          }
        ]
      });
    }
  }

  const blogId = getNaverBlogId(blogUrl);
  if (!blogId) {
    return Response.json({
      success: true,
      blogId: null,
      source: "no-url",
      promotions: [],
      blogStats: {
        recent30d: 0,
        lastPosted: "",
        reactionScore: 0,
        reactionDetail: {
          evaluatedCount: 0,
          postsWith2PlusComments: 0,
          postsWithComments: 0,
          postsWithLikes: 0
        }
      }
    });
  }

  try {
    const rssUrl = `https://rss.blog.naver.com/${blogId}.xml`;
    const response = await fetch(rssUrl, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
      },
      next: { revalidate: 0 }
    });

    if (!response.ok) {
      throw new Error(`Failed to fetch Naver RSS. Status: ${response.status}`);
    }

    const xml = await response.text();
    const items = [];
    const itemMatches = xml.match(/<item>([\s\S]*?)<\/item>/g) || [];

    const now = new Date();
    let recent30d = 0;
    let lastPostedDate = "";
    const fullParsedPosts = [];

    for (let i = 0; i < itemMatches.length; i++) {
      const itemXml = itemMatches[i];
      const titleMatch = itemXml.match(/<title><!\[CDATA\[([\s\S]*?)\]\]><\/title>/) || itemXml.match(/<title>([\s\S]*?)<\/title>/);
      const linkMatch = itemXml.match(/<link><!\[CDATA\[([\s\S]*?)\]\]><\/link>/) || itemXml.match(/<link>([\s\S]*?)<\/link>/);
      const descMatch = itemXml.match(/<description><!\[CDATA\[([\s\S]*?)\]\]><\/description>/) || itemXml.match(/<description>([\s\S]*?)<\/description>/);
      const pubDateMatch = itemXml.match(/<pubDate><!\[CDATA\[([\s\S]*?)\]\]><\/pubDate>/) || itemXml.match(/<pubDate>([\s\S]*?)<\/pubDate>/);

      if (titleMatch && linkMatch) {
        const title = titleMatch[1].trim();
        const link = linkMatch[1].trim();
        let rawDesc = descMatch ? descMatch[1].trim() : "";
        rawDesc = rawDesc.replace(/<[^>]*>/g, "").replace(/\s+/g, " ").trim();
        const content = rawDesc.slice(0, 120) + (rawDesc.length > 120 ? "..." : "");
        let pubDateStr = "";

        if (pubDateMatch) {
          const d = new Date(pubDateMatch[1].trim());
          if (!isNaN(d.getTime())) {
            pubDateStr = formatKstDate(d);
            if (!lastPostedDate) lastPostedDate = pubDateStr;
            const diffDays = (now - d) / (1000 * 60 * 60 * 24);
            if (diffDays <= 30) recent30d++;
          }
        }

        const postObj = {
          title,
          content,
          url: link,
          pubDate: pubDateStr,
          status: "active"
        };

        fullParsedPosts.push(postObj);
        if (items.length < 9) {
          items.push(postObj);
        }
      }
    }

    if (items.length === 0) {
      throw new Error("No items found in RSS feed.");
    }

    // Inspect recent 5 posts through NAVER's comment and sympathy JSON requests.
    const targetPosts = fullParsedPosts.slice(0, 5);
    let c2Plus = 0;
    let c1Plus = 0;
    let likesCount = 0;

    const reactionDetails = await Promise.all(
      targetPosts.map(async (p, idx) => {
        const logNo = getPostLogNo(p.url);
        if (!logNo) {
          return { idx, comments: 0, likes: 0, commentsCollected: false, likesCollected: false, title: p.title };
        }

        try {
          const postUrl = `https://m.blog.naver.com/PostView.naver?blogId=${blogId}&logNo=${logNo}`;
          const pRes = await fetch(postUrl, {
            headers: {
              "User-Agent": NAVER_MOBILE_USER_AGENT
            },
            cache: "no-store",
            signal: AbortSignal.timeout(5000)
          });
          if (!pRes.ok) throw new Error(`Post page returned ${pRes.status}`);
          const html = await pRes.text();
          const commentContext = getCommentContext(html, logNo);
          const referer = `https://blog.naver.com/${blogId}/${logNo}`;

          const [commentResult, sympathyResult] = await Promise.allSettled([
            commentContext
              ? fetchCommentCount({ ...commentContext, referer })
              : Promise.reject(new Error("Comment objectId was not found")),
            fetchSympathyCount({ blogId, logNo, referer })
          ]);

          return {
            idx,
            logNo,
            comments: commentResult.status === "fulfilled" ? commentResult.value : 0,
            likes: sympathyResult.status === "fulfilled" ? sympathyResult.value : 0,
            commentsCollected: commentResult.status === "fulfilled",
            likesCollected: sympathyResult.status === "fulfilled",
            title: p.title
          };
        } catch {
          return { idx, logNo, comments: 0, likes: 0, commentsCollected: false, likesCollected: false, title: p.title };
        }
      })
    );

    reactionDetails.forEach((r) => {
      if (r.comments >= 2) c2Plus++;
      if (r.comments >= 1) c1Plus++;
      if (r.likes >= 1) likesCount++;
    });

    let reactionScore = 0;
    if (c2Plus >= 2 && likesCount >= 4) {
      reactionScore = 5;
    } else if (c1Plus >= 1 && likesCount >= 4) {
      reactionScore = 4;
    } else if (likesCount >= 3) {
      reactionScore = 3;
    } else if (likesCount >= 1) {
      reactionScore = 2;
    } else if (c1Plus > 0) {
      reactionScore = 1;
    }

    const commentsCollectedCount = reactionDetails.filter((item) => item.commentsCollected).length;
    const likesCollectedCount = reactionDetails.filter((item) => item.likesCollected).length;

    return Response.json({
      success: true,
      blogId,
      source: "naver-rss",
      promotions: items,
      blogStats: {
        recent30d,
        lastPosted: lastPostedDate,
        reactionScore,
        reactionDetail: {
          evaluatedCount: targetPosts.length,
          postsWith2PlusComments: c2Plus,
          postsWithComments: c1Plus,
          postsWithLikes: likesCount,
          commentsCollectedCount,
          likesCollectedCount,
          complete: commentsCollectedCount === targetPosts.length && likesCollectedCount === targetPosts.length,
          posts: reactionDetails
        }
      }
    });

  } catch (error) {
    console.error("Scraping Naver Blog RSS failed:", error);
    return Response.json({
      success: true,
      blogId,
      source: "error-empty",
      promotions: [],
      blogStats: {
        recent30d: 0,
        lastPosted: "",
        reactionScore: 0,
        reactionDetail: {
          evaluatedCount: 0,
          postsWith2PlusComments: 0,
          postsWithComments: 0,
          postsWithLikes: 0
        }
      }
    });
  }
}
