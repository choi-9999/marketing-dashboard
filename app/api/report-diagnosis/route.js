import { GoogleGenerativeAI } from "@google/generative-ai";

export const dynamic = "force-dynamic";
export const revalidate = 0;

export async function POST(request) {
  try {
    const body = await request.json();
    const {
      branch,
      region,
      score,
      grade,
      rank,
      totalRankedBranches,
      nationalAverage,
      regionAverage,
      nationalOperationAverage,
      regionOperationAverage,
      participationRate,
      programs,
      snsScore,
      snsGrade,
      blogScore,
      instagramScore,
      recentContentCount,
      latestBlogPosts,
      latestInstagramPosts,
      hasFacilityVideo,
      collabUrlCount,
      mentorCount,
      scholarshipAmount,
      contentAssetCount
    } = body;

    if (!branch) {
      return Response.json({ success: false, error: "지점명이 누락되었습니다." }, { status: 400 });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.warn("GEMINI_API_KEY is not defined. Falling back to local heuristic analysis.");
      return Response.json({
        success: true,
        fallback: true,
        summary: `${branch} 지점의 종합 운영 데이터 기반 표준 진단 결과입니다.`,
        diagnosis: [
          `전체 종합 점수는 ${score}점(${grade})이며, 전국 ${totalRankedBranches || 0}개 지점 중 ${rank || "-"}위입니다.`,
          `활성화 방안 참여율은 ${participationRate}%이며, SNS 종합 점수는 ${snsScore != null ? snsScore + "점" : "미집계"}입니다.`,
          recentContentCount < 8 ? "최근 30일 SNS 콘텐츠 수가 권장 기준(월 8개) 미만입니다." : "최근 SNS 콘텐츠가 활발히 운영되고 있습니다."
        ],
        recommendations: [
          participationRate < 50 ? `활성화 방안 참여율(${participationRate}%) 강화를 위해 미참여 프로그램부터 운영 일정에 반영하세요.` : "현재 우수한 활성화 방안 참여율을 유지하며 프로그램별 실행 품질을 고도화하세요.",
          (!snsScore || Number(snsScore) < 60) ? "SNS 채널(블로그·인스타그램)의 정기 발행 계획을 수립하고 학부모/수험생 타겟 콘텐츠를 보강하세요." : "SNS 채널의 양호한 흐름을 바탕으로 실제 상담 문의 전환율을 높이는 콜투액션(CTA)을 강화하세요.",
          recentContentCount < 8 ? `최근 30일 SNS 콘텐츠가 ${recentContentCount}개입니다. 채널 합산 월 8개 이상 정기 업로드를 1차 목표로 설정하세요.` : "최근 SNS 발행 주기를 유지하며 주요 게시물의 네이버 플레이스 및 지역 맘카페 연계 확산을 추진하세요.",
          !hasFacilityVideo ? "학습 환경 신뢰도 향상을 위해 시설 및 관리 시스템 안내 영상을 제작·연결하세요." : "등록된 시설 영상을 블로그 및 인스타그램 주요 하이라이트로 재확산하여 상담 전환에 활용하세요.",
          !collabUrlCount ? "본사 협업 캠페인 및 이벤트 링크를 확보하여 지점 마케팅 노출을 확대하세요." : "협업 이벤트 콘텐츠를 원내 상담 시 추가 안내 자료로 적극 활용하세요.",
          !mentorCount ? "합격 사례와 생생한 학습 경험을 증빙할 수 있는 장학생·멘토단 인물 콘텐츠를 발굴하세요." : "등록된 멘토단의 합격 수기 및 학습 팁을 카드뉴스로 가공하여 SNS에 정기 연재하세요."
        ]
      });
    }

    // Initialize Gemini Client
    const genAI = new GoogleGenerativeAI(apiKey);
    const modelName = process.env.GEMINI_MODEL || "gemini-3.5-flash-lite";
    const model = genAI.getGenerativeModel({ model: modelName });

    // Format program status
    const programSummary = Array.isArray(programs)
      ? programs.map((p) => `- ${p.name}: ${p.activeEvents > 0 ? "참여 (" + p.activeEvents + "건)" : "미참여"}`).join("\n")
      : "정보 없음";

    // Format recent post headlines
    const blogHeadlines = Array.isArray(latestBlogPosts) && latestBlogPosts.length > 0
      ? latestBlogPosts.map((p, idx) => `  [블로그 ${idx + 1}] ${p.title || p}`).join("\n")
      : "  (최근 블로그 글 없음)";

    const instaHeadlines = Array.isArray(latestInstagramPosts) && latestInstagramPosts.length > 0
      ? latestInstagramPosts.map((p, idx) => `  [인스타그램 ${idx + 1}] ${p.title || p.caption || p}`).join("\n")
      : "  (최근 인스타그램 글 없음)";

    const prompt = `
당신은 대한민국 대표 독학재수 전문학원 '이투스247학원' 본사의 마케팅 전략 수석 책임자이자 지점 운영 컨설팅 최고 전문가입니다.
다음은 '${branch}' 지점의 최신 운영 및 마케팅 종합 성과 데이터입니다:

=========================================
[지점 기본 및 종합 성과]
- 지점명: ${branch} (${region})
- 종합 운영 점수: ${score}점 (${grade}) / 순위: 전체 ${totalRankedBranches || 0}개 지점 중 ${rank || "-"}위
- 전국 평균 점수: ${nationalAverage}점 / 권역 평균 점수: ${regionAverage}점
- 활성화 방안 참여율: ${participationRate}%

[마케팅 활성화 프로그램 참여 현황]
${programSummary}

[SNS 채널 운영 현황]
- SNS 종합 점수: ${snsScore != null ? snsScore + "점" : "미집계"} (등급: ${snsGrade || "-"})
- 블로그 점수: ${blogScore}점 / 최근 발행 수: ${recentContentCount}개
- 인스타그램 점수: ${instagramScore}점
- 최근 블로그 콘텐츠:\n${blogHeadlines}
- 최근 인스타그램 콘텐츠:\n${instaHeadlines}

[마케팅 자산 및 인프라]
- 시설 영상 등록 여부: ${hasFacilityVideo ? "등록됨" : "미등록"}
- 본사 협업 이벤트 확산 링크: ${collabUrlCount}개
- 등록된 장학생/멘토단: ${mentorCount}명 (장학금: ${scholarshipAmount ? scholarshipAmount.toLocaleString("ko-KR") + "원" : "0원"})
- 등록된 콘텐츠 자산: ${contentAssetCount || 0}개
=========================================

위 지표 데이터의 상관관계와 강점/약점을 냉철하게 분석하여, 지점장 및 원장이 즉시 이해하고 실행할 수 있는 전략적 진단과 실행 제안을 도출해 주세요.

1. 총평 (summary):
   - 해당 지점의 종합적인 위상과 핵심 현황을 명쾌하게 짚어주는 1~2문장의 핵심 브리핑.

2. 핵심 종합 진단 (diagnosis):
   - 전국/권역 평균 대비 강점 및 즉시 보완이 필요한 취약점을 2~3가지 핵심 포인트로 도출 (배열 형식, 각 항목 1~2줄).
   - 실제 수치(점수, 순위, SNS 현황 등)를 근거로 구체적으로 언급할 것.

3. 전술적 우선 실행 제안 (recommendations):
   - 이번 달 지점에서 가장 우선적으로 실행해야 할 마케팅 및 원내 운영 액션 플랜 **정확히 6가지** (배열 형식).
   - 6가지 항목은 각각 다음 핵심 영역을 골고루 포괄해야 합니다:
     ① 활성화 방안 프로그램 연계 강화 방안
     ② 블로그/인스타그램 채널별 맞춤 콘텐츠 및 발행 주기 개선
     ③ 시설 영상 또는 원내 학습 환경 신뢰도 확산
     ④ 멘토/장학생/합격생 후기 자산 활용 방안
     ⑤ 본사 협업 이벤트 및 온·오프라인 마케팅 접점 확대
     ⑥ 실제 원내 방문/상담 문의 전환(CTA) 극대화 전략
   - 단순한 조언이 아닌 "무엇을 어떻게 실행하고 어디에 확산해야 하는지" 구체적이고 실천 가능한 지침으로 작성할 것.

반드시 아래 JSON 스키마를 준수하여 응답해 주세요.
`;

    const response = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: {
          type: "OBJECT",
          properties: {
            summary: {
              type: "STRING",
              description: "1~2문장의 핵심 총평"
            },
            diagnosis: {
              type: "ARRAY",
              items: { type: "STRING" },
              description: "2~3가지 핵심 진단 포인트"
            },
            recommendations: {
              type: "ARRAY",
              items: { type: "STRING" },
              description: "정확히 6가지 구체적인 전술적 우선 실행 제안"
            }
          },
          required: ["summary", "diagnosis", "recommendations"]
        }
      }
    });

    const responseText = response.response.text();
    const parsedData = JSON.parse(responseText);

    return Response.json({
      success: true,
      fallback: false,
      model: modelName,
      summary: parsedData.summary,
      diagnosis: parsedData.diagnosis,
      recommendations: parsedData.recommendations
    });

  } catch (error) {
    console.error("Gemini report diagnosis API error:", error);
    return Response.json({
      success: false,
      error: error.message,
      fallback: true
    });
  }
}
