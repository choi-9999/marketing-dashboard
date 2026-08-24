import { NextResponse } from "next/server";
import https from "https";

export const dynamic = "force-dynamic";

function fetchHtmlLive(url) {
  const host = new URL(url).host;
  return new Promise((resolve, reject) => {
    https.get(url, {
      headers: {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection": "keep-alive",
        "Host": host,
        "sec-ch-ua": '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
        "sec-ch-ua-mobile": "?0",
        "sec-ch-ua-platform": '"Windows"',
        "Upgrade-Insecure-Requests": "1"
      }
    }, (res) => {
      let data = "";
      res.on("data", (chunk) => {
        data += chunk;
      });
      res.on("end", () => {
        resolve(data);
      });
    }).on("error", (err) => {
      reject(err);
    });
  });
}

export async function GET(request) {
  const { searchParams } = new URL(request.url);
  const branch = searchParams.get("branch");

  if (!branch) {
    return NextResponse.json({ error: "Branch parameter is required" }, { status: 400 });
  }

  const cleanBranch = String(branch).trim();
  const clientId = process.env.NAVER_CLIENT_ID;
  const clientSecret = process.env.NAVER_CLIENT_SECRET;

  // Fallback mock data in case the API keys are invalid or unconfigured
  const fallbackSocial = {
    kin: [
      { title: `<b>${cleanBranch}</b> 독학재수학원 잇올이랑 이투스247 비교해주세요`, description: `내년에 <b>${cleanBranch}</b> 지역에서 독재하려고 하는데 잇올 수능선배 이투스247 중에서 어디가 분위기 관리가 더 잘되나요? 급식이랑 좌석 종류도 궁금해요...`, link: "https://kin.naver.com", pubDate: "2026-06-25" },
      { title: `<b>${cleanBranch} 이투스247</b> 질문이요`, description: `제가 윈터스쿨로 <b>${cleanBranch} 이투스247</b>에 들어가려고 하는데요. 벌써 마감인가요? 대기 타야하는지 궁금해서 질문 올립니다.`, link: "https://kin.naver.com", pubDate: "2026-06-20" },
      { title: `재수학원 <b>${cleanBranch}</b> 근처 추천`, description: `인강 위주로 독재하면서 주 1회 상담받는 <b>${cleanBranch}</b> 주변 학원을 찾고 있습니다. 이투스247도 괜찮다는 말을 들었는데 실제 다녀보신 분 평이 어때요?`, link: "https://kin.naver.com", pubDate: "2026-06-18" }
    ],
    blog: [
      { title: `[합격 수기] <b>${cleanBranch} 이투스247</b>에서 독재하고 연세대 합격했어요!`, description: `안녕하세요! 오늘은 제가 작년에 <b>${cleanBranch} 이투스</b>학원을 다니며 멘토쌤들의 입시 코칭을 받아 합격한 후기를 솔직하게 남겨보려고 합니다...`, link: "https://section.blog.naver.com", postdate: "2026-06-22" },
      { title: `<b>${cleanBranch} 이투스247</b> 윈터스쿨 실제 시설 투어 & 솔직 후기`, description: `오늘 소개해드릴 곳은 자기주도학습 분위기가 확실하게 잡혀있는 <b>${cleanBranch} 이투스247</b>입니다. 면학 분위기와 개인별 학습 스케줄러 관리 덕에...`, link: "https://section.blog.naver.com", postdate: "2026-06-15" }
    ],
    news: [
      { title: "2027학년도 대입 수시 모집 가이드 핵심 공개... 이투스 ECI 주목", description: "이투스ECI에서 운영하는 독학재수 브랜드 이투스247학원이 전국 지점망을 통해 개인화 맞춤 입시 컨설팅과 정기 진단 평가 시스템을 도입해 큰 기대를 모으고 있다...", link: "https://news.naver.com", pubDate: "2026-06-28" },
      { title: "대입 수능 6월 모평 가채점 결과 분석과 반수 성공 전략 좌담회", description: "평가원 6월 모의평가 분석 결과 국수영 난이도가 조율되면서 반수반 수요가 증가하고 있으며, 독학재수 전문 학원들의 학습 코칭과 멘토단 상담이 핵심 합격 요인으로...", link: "https://news.naver.com", pubDate: "2026-06-26" }
    ],
    cafe: [
      { title: `[수만휘] <b>${cleanBranch} 이투스247</b> 수강평이나 장단점 있을까요?`, description: "이번에 독재 시작하려는데 수만휘 선배님들 중에서 여기 다녀보신 분 계신지 여쭙습니다. 친목 차단이나 시설 상태가 어떤지 궁금해요...", link: "https://cafe.naver.com/f-e", pubDate: "2026-06-27" },
      { title: `[수만휘] <b>${cleanBranch}</b> 주변 독재학원 친목 관리 수준`, description: "이투스247이랑 잇올 고민 중인데 분위기 엄격한 곳 선호합니다. 특히 남녀 분리나 인강 사이트 통제가 엄격하게 이루어지는 곳 추천 부탁드립니다...", link: "https://cafe.naver.com/f-e", pubDate: "2026-06-21" }
    ],
    orbi: [
      { title: `[오르비] <b>${cleanBranch} 이투스247</b> 질문받음`, link: "https://orbi.kr", pubDate: "2026-06-24" },
      { title: `[오르비] <b>${cleanBranch}</b> 독재생들 자습 분위기 질문`, link: "https://orbi.kr", pubDate: "2026-06-19" }
    ]
  };

  // 1. Orbi 실제 크롤링 (list 주소 fetch 후 파싱)
  let parsedOrbi = [];
  try {
    const orbiHtml = await fetchHtmlLive("https://orbi.kr/list");
    const blockRegex = /<p class="title">([\s\S]*?)<\/p>/g;
    let blockMatch;
    while ((blockMatch = blockRegex.exec(orbiHtml)) !== null) {
      const blockContent = blockMatch[1];
      const linkRegex = /<a href="\/(\d+)\/([^"]+)"[^>]*>([\s\S]*?)<\/a>/g;
      let linkMatch;
      while ((linkMatch = linkRegex.exec(blockContent)) !== null) {
        const id = linkMatch[1];
        const slug = linkMatch[2];
        const rawTitle = linkMatch[3].trim();
        const title = rawTitle.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ');
        
        if (title.length > 0) {
          const shortBranch = cleanBranch.substring(0, 2);
          const targetKeywords = [shortBranch, "이투스247", "독재", "재수학원", "수능", "수학", "인강", "공부", "국어", "실모", "학평", "채점"];
          if (targetKeywords.some(kw => title.includes(kw))) {
            parsedOrbi.push({
              title: `[오르비] ${title}`,
              link: `https://orbi.kr/${id}/${slug}`,
              pubDate: new Date().toISOString().split("T")[0]
            });
          }
        }
      }
    }
  } catch (err) {
    console.error("Orbi live crawling failed:", err);
  }

  // 만약 검색된 실시간 오르비 글이 없으면 데모용 Mock 오르비 데이터를 믹스인
  if (parsedOrbi.length === 0) {
    parsedOrbi = fallbackSocial.orbi;
  }

  // 2. Naver Cafe Search Live Web Crawling (API Key 에러 대안 및 백업망) - 데몬 저장 JSON 읽기
  let liveCafeArticles = [];
  try {
    const fs = require("fs");
    const path = require("path");
    const jsonPath = path.join(process.cwd(), "data", "live_cafe.json");
    if (fs.existsSync(jsonPath)) {
      const data = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
      liveCafeArticles = data[cleanBranch] || [];
    }
  } catch (err) {
    console.error("Failed to read live_cafe.json:", err);
  }

  // 실시간 수집된 카페 글이 있다면 그것을 fallbackSocial.cafe 에 병합
  if (liveCafeArticles.length > 0) {
    fallbackSocial.cafe = liveCafeArticles;
  }

  if (!clientId || !clientSecret || clientId.includes("YOUR_") || clientSecret.includes("YOUR_")) {
    return NextResponse.json({ success: true, mode: "mock", data: { ...fallbackSocial, cafe: fallbackSocial.cafe, orbi: parsedOrbi } }, {
      headers: {
        "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate",
        "Pragma": "no-cache",
        "Expires": "0"
      }
    });
  }

  try {
    const headers = {
      "x-ncp-apigw-api-key-id": clientId,
      "x-ncp-apigw-api-key": clientSecret
    };

    // Clean query keywords
    const kinQuery = encodeURIComponent(`"${cleanBranch} 이투스247"`);
    const kinBackupQuery = encodeURIComponent(`"${cleanBranch} 독학재수"`);
    const blogQuery = encodeURIComponent(`"${cleanBranch} 이투스247"`);
    const newsQuery = encodeURIComponent(`"이투스247" OR "독학재수학원"`);
    
    // 수만휘(f-e) 카페 글 수집을 위해 cafearticle API 호출
    const cafeQuery = encodeURIComponent(`"${cleanBranch} 이투스247" OR "${cleanBranch} 독재"`);

    const [kinRes, blogRes, newsRes, cafeRes] = await Promise.all([
      fetch(`https://naverapihub.apigw.ntruss.com/search/v1/kin?query=${kinQuery}&display=5&sort=sim`, { headers }).then(r => r.json()),
      fetch(`https://naverapihub.apigw.ntruss.com/search/v1/blog?query=${blogQuery}&display=5&sort=sim`, { headers }).then(r => r.json()),
      fetch(`https://naverapihub.apigw.ntruss.com/search/v1/news?query=${newsQuery}&display=5&sort=sim`, { headers }).then(r => r.json()),
      fetch(`https://naverapihub.apigw.ntruss.com/search/v1/cafearticle?query=${cafeQuery}&display=15&sort=sim`, { headers }).then(r => r.json())
    ]);

    let kinItems = kinRes.items || [];
    if (kinItems.length === 0) {
      const backupKin = await fetch(`https://naverapihub.apigw.ntruss.com/search/v1/kin?query=${kinBackupQuery}&display=5&sort=sim`, { headers }).then(r => r.json());
      kinItems = backupKin.items || [];
    }

    const formatItem = (item) => ({
      title: item.title || "",
      description: item.description || "",
      link: item.link || "#",
      pubDate: item.pubDate || item.postdate || ""
    });

    // 입시 및 맘카페 글 필터링 (수만휘 f-e, 분당맘카페 2008bunsamo, 맘스홀릭 등 교육/맘 커뮤니티 전반)
    const rawCafeItems = cafeRes.items || [];
    const filteredCafeItems = rawCafeItems
      .filter(item => {
        const url = item.cafeurl || "";
        const name = item.cafename || "";
        return url.includes("f-e") || url.includes("2008bunsamo") || name.includes("수만휘") || name.includes("맘") || name.includes("분사모") || name.includes("학부모") || name.includes("공부");
      })
      .map(item => {
        const isSumanhwi = (item.cafeurl && item.cafeurl.includes("f-e")) || (item.cafename && item.cafename.includes("수만휘"));
        const prefix = isSumanhwi ? "[수만휘]" : "[입시/맘카페]";
        return {
          title: `${prefix} ${item.title.replace(/<[^>]+>/g, '')}`,
          description: item.description || "",
          link: item.link || "#",
          pubDate: item.pubDate || item.postdate || ""
        };
      });

    const parsedData = {
      kin: kinItems.map(formatItem),
      blog: (blogRes.items || []).map(formatItem),
      news: (newsRes.items || []).map(formatItem),
      cafe: filteredCafeItems.length > 0 ? filteredCafeItems : fallbackSocial.cafe,
      orbi: parsedOrbi
    };

    return NextResponse.json({
      success: true,
      mode: "naver-search-api",
      data: parsedData
    }, {
      headers: {
        "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate",
        "Pragma": "no-cache",
        "Expires": "0"
      }
    });

  } catch (error) {
    console.error("Failed to fetch NAVER Search API:", error);
    return NextResponse.json({ success: true, mode: "catch-fallback", data: { ...fallbackSocial, cafe: fallbackSocial.cafe, orbi: parsedOrbi } }, {
      headers: {
        "Cache-Control": "no-store, no-cache, must-revalidate, proxy-revalidate",
        "Pragma": "no-cache",
        "Expires": "0"
      }
    });
  }
}
