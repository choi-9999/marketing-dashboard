const fs = require('fs');
const https = require('https');
const path = require('path');

const SAVE_PATH = path.join(__dirname, 'live_cafe.json');

// 브라우저 헤더셋 정의
const HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
  "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
  "Connection": "keep-alive",
  "Host": "search.naver.com",
  "sec-ch-ua": '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
  "sec-ch-ua-mobile": "?0",
  "sec-ch-ua-platform": '"Windows"',
  "Upgrade-Insecure-Requests": "1"
};

function fetchHtmlLive(url) {
  return new Promise((resolve, reject) => {
    https.get(url, { headers: HEADERS }, (res) => {
      let data = "";
      res.on("data", (chunk) => { data += chunk; });
      res.on("end", () => { resolve(data); });
    }).on("error", (err) => { reject(err); });
  });
}

// 지점 목록 정의
const BRANCHES = ["분당정자", "대치", "수지", "이천기숙", "부천"];

async function crawlSingleBranch(branchName) {
  const shortBranch = branchName.substring(0, 2);
  const searchQuery = encodeURIComponent(`${shortBranch} 학원`);
  const searchUrl = `https://search.naver.com/search.naver?where=article&query=${searchQuery}&_ts=${Date.now()}`;
  
  const liveCafeArticles = [];
  try {
    const searchHtml = await fetchHtmlLive(searchUrl);
    const aTagRegex = /<a\s+([^>]*?)>([\s\S]*?)<\/a>/gi;
    let match;
    
    while ((match = aTagRegex.exec(searchHtml)) !== null) {
      const attrs = match[1];
      const rawTitle = match[2];
      
      const hrefMatch = attrs.match(/href="([^"]+)"/i);
      if (!hrefMatch) continue;
      const url = hrefMatch[1];
      
      const cafeMatch = url.match(/https:\/\/cafe\.naver\.com\/([a-zA-Z0-9_-]+)\/(\d+)/i);
      if (!cafeMatch) continue;
      
      const clubId = cafeMatch[1];
      const title = rawTitle.replace(/<[^>]+>/g, '').replace(/\s+/g, ' ').trim();
      
      if (title.length > 5 && !title.startsWith("naver.com") && !title.includes("카페 더보기")) {
        const isSumanhwi = clubId === "f-e" || title.includes("수만휘");
        const isBunsamo = clubId === "2008bunsamo" || title.includes("분사모") || clubId === "2008bunsamo";
        const isReply = title.startsWith("RE");
        
        let prefix = "[입시/맘카페]";
        if (isSumanhwi) prefix = "[수만휘]";
        else if (isBunsamo) prefix = "[분사모]";
        else if (isReply) prefix = "[카페 댓글]";
        
        if (!liveCafeArticles.some(art => art.link === url)) {
          liveCafeArticles.push({
            title: `${prefix} ${title}`,
            link: url,
            pubDate: new Date().toISOString().split("T")[0]
          });
        }
      }
    }
    console.log(`[${new Date().toLocaleTimeString()}] Crawled "${branchName}" - Found ${liveCafeArticles.length} live articles.`);
  } catch (err) {
    console.error(`[ERROR] Crawling failed for ${branchName}:`, err.message);
  }
  return liveCafeArticles;
}

async function startDaemon() {
  console.log("Starting Cafe Background Daemon...");
  
  while (true) {
    const database = {};
    for (const br of BRANCHES) {
      database[br] = await crawlSingleBranch(br);
      // 네이버 과도요청 차단 우회를 위해 지점별 2초 딜레이
      await new Promise(r => setTimeout(r, 2000));
    }
    
    // JSON 결과 디스크 쓰기
    try {
      fs.writeFileSync(SAVE_PATH, JSON.stringify(database, null, 2), 'utf8');
      console.log(`[OK] Saved results to live_cafe.json`);
    } catch (e) {
      console.error("[ERROR] Failed to save JSON database:", e.message);
    }
    
    // 30초마다 전체 지점 재수집 루프
    console.log("Sleeping 30 seconds before next cycle...");
    await new Promise(r => setTimeout(r, 30000));
  }
}

startDaemon();
