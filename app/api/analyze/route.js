import { GoogleGenerativeAI } from "@google/generative-ai";

export const dynamic = "force-dynamic";
export const revalidate = 0;

export async function POST(request) {
  try {
    const body = await request.json();
    const { compName, promotions } = body;

    if (!compName || !promotions || !Array.isArray(promotions) || promotions.length === 0) {
      return Response.json({ success: false, error: "Invalid parameters" }, { status: 400 });
    }

    const apiKey = process.env.GEMINI_API_KEY;
    if (!apiKey) {
      console.warn("GEMINI_API_KEY is not defined. Falling back to local heuristic analysis.");
      return Response.json({ success: false, fallback: true, message: "API key missing" });
    }

    // Initialize Gemini Client
    const genAI = new GoogleGenerativeAI(apiKey);
    const model = genAI.getGenerativeModel({ model: "gemini-3.1-flash-lite" });

    // Format promotion data for the prompt
    const promoContext = promotions
      .map((p, idx) => `[${idx + 1}] 제목: ${p.title}\n내용 요약: ${p.content}`)
      .join("\n\n");

    const prompt = `
당신은 대한민국 대표 대입 독학재수 전문 학원인 '이투스247 분당정자점'의 원장 및 마케팅 전략 수석 책임자입니다.
경쟁사인 '${compName}'의 최신 마케팅 및 프로모션 글 정보는 다음과 같습니다:

${promoContext}

위 경쟁사의 프로모션 흐름과 동향을 냉철히 분석하여 다음 두 가지 분석 결과를 도출해 주세요.

1. 동향 요약 (trend):
- 이 경쟁사의 프로모션 움직임을 분석하여 어떤 마케팅에 집중하고 있는지 핵심 요약을 3가지(각각 1줄) 도출해 주세요.
- 줄바꿈으로 나누어 1., 2., 3. 번호 매기기 형식으로 작성해 주세요. (예: "1. '...' 등 실전 모의고사 운영을 통한 ...")
- 분석 시 경쟁사가 크롤링된 프로모션의 실제 제목 중 핵심 키워드나 맥락을 반드시 언급해 주어 독창적인 분석이 되게 하십시오.

2. AI 전술적 대응 지침 (guide):
- 이투스247 입장에서 경쟁사의 위의 3가지 동향에 1:1로 정확하게 대응할 수 있는 가장 현실적이고 날카로운 반박 영업/상담 전술 또는 원내 마케팅 지침 3가지(각각 1줄)를 도출해 주세요.
- 줄바꿈으로 나누어 1., 2., 3. 번호 매기기 형식으로 작성해 주세요. (예: "1. [모의고사 대응] 자사 풀시즌 모의고사 라인업과 1:1 진단 피드백의 비교 우위를 ...")

응답은 반드시 아래 JSON 스키마를 만족하는 올바른 JSON 형식이어야 합니다.
{
  "trend": "1. ...\\n2. ...\\n3. ...",
  "guide": "1. ...\\n2. ...\\n3. ..."
}
`;

    const response = await model.generateContent({
      contents: [{ role: "user", parts: [{ text: prompt }] }],
      generationConfig: {
        responseMimeType: "application/json",
        responseSchema: {
          type: "OBJECT",
          properties: {
            trend: {
              type: "STRING",
              description: "1. ...\n2. ...\n3. ... 형식의 3가지 동향 요약 텍스트"
            },
            guide: {
              type: "STRING",
              description: "1. ...\n2. ...\n3. ... 형식의 3가지 대응 지침 텍스트"
            }
          },
          required: ["trend", "guide"]
        }
      }
    });

    const responseText = response.response.text();
    const parsedData = JSON.parse(responseText);

    return Response.json({
      success: true,
      trend: parsedData.trend,
      guide: parsedData.guide
    });

  } catch (error) {
    console.error("Gemini API analyze error:", error);
    return Response.json({ success: false, error: error.message, fallback: true });
  }
}
