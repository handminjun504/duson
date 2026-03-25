const { GoogleGenerativeAI } = require('@google/generative-ai');
const logger = require('../utils/logger');

let genAI = null;

function initGemini() {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    logger.warn('GEMINI_API_KEY가 설정되지 않았습니다. AI 분석 기능이 비활성화됩니다.');
    return false;
  }
  genAI = new GoogleGenerativeAI(apiKey);
  return true;
}

async function analyzeComparison(sheetData, gyeongliData) {
  if (!genAI) {
    return {
      available: false,
      message: 'Gemini API 키가 설정되지 않았습니다. .env 파일에 GEMINI_API_KEY를 추가하세요.',
    };
  }

  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });

    const prompt = `당신은 회계 전문가입니다. 아래 두 데이터를 비교 분석해주세요.

## Google Sheet 매출 데이터
${JSON.stringify(sheetData?.slice(0, 20), null, 2)}

## 경리나라 매출 데이터
${JSON.stringify(gyeongliData?.slice(0, 3), null, 2)}

다음 항목을 분석해주세요:
1. 두 데이터 간 불일치 건이 있는지
2. 거래처명이 다르게 표기된 건은 없는지 (퍼지 매칭)
3. 누락된 거래 건이 있는지
4. 비고 자동 판단 (이월 여부, 공장 출고 건 구분)

JSON 형식으로 응답해주세요:
{
  "discrepancies": [{"description": "...", "severity": "high|medium|low"}],
  "nameMatches": [{"sheet": "...", "gyeongli": "...", "confidence": 0.95}],
  "missingEntries": [{"source": "sheet|gyeongli", "description": "..."}],
  "notesSuggestions": [{"product": "...", "deliveryDate": "...", "suggestedNote": "..."}],
  "summary": "전체 분석 요약"
}`;

    const result = await model.generateContent(prompt);
    const text = result.response.text();

    let parsed;
    try {
      const jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/) || text.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(jsonMatch ? (jsonMatch[1] || jsonMatch[0]) : text);
    } catch {
      parsed = { summary: text, raw: true };
    }

    logger.info('Gemini 매출 대조 분석 완료');
    return { available: true, analysis: parsed };
  } catch (err) {
    logger.error('Gemini API 호출 실패', { error: err.message });
    return { available: false, message: err.message };
  }
}

function truncateForGemini(data, maxItems = 50) {
  if (!Array.isArray(data)) return data;
  return data.slice(0, maxItems);
}

function summarizeSales(salesByClient) {
  if (!Array.isArray(salesByClient)) return salesByClient;
  return salesByClient.map(c => ({
    client: c.client,
    itemCount: c.items?.length || 0,
    totalAmount: (c.items || []).reduce((s, i) => s + (i.total || 0), 0),
    prevBalance: c.prevBalance || 0,
  }));
}

async function suggestDepositMatches(deposits, salesByClient) {
  if (!genAI) {
    return { available: false, message: 'Gemini API 키 미설정' };
  }

  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });

    const safeDeposits = truncateForGemini(deposits, 30).map(d => ({
      date: d.date, client: d.client, amount: d.amount, bank: d.bank,
    }));
    const safeSales = summarizeSales(salesByClient);

    const prompt = `회계 담당자입니다. 아래 입금 내역과 거래처별 미수금을 비교하여 매칭 제안을 해주세요.

## 입금 내역
${JSON.stringify(safeDeposits, null, 2)}

## 거래처별 매출 요약
${JSON.stringify(safeSales, null, 2)}

각 입금에 대해 어떤 매출 건과 매칭될 수 있는지 제안해주세요.
개인통장 이체건이 있을 수 있으니 금액이 정확히 일치하지 않을 수 있습니다.

JSON 형식:
{
  "matches": [
    {"depositIndex": 0, "matchedItems": ["정백당 2/3 납품분"], "confidence": 0.8, "reason": "금액 근사 일치"}
  ],
  "summary": "매칭 요약"
}`;

    const result = await model.generateContent(prompt);
    const text = result.response.text();

    let parsed;
    try {
      const jsonMatch = text.match(/```json\s*([\s\S]*?)\s*```/) || text.match(/\{[\s\S]*\}/);
      parsed = JSON.parse(jsonMatch ? (jsonMatch[1] || jsonMatch[0]) : text);
    } catch {
      parsed = { summary: text, raw: true };
    }

    logger.info('Gemini 입금 매칭 분석 완료');
    return { available: true, analysis: parsed };
  } catch (err) {
    logger.error('Gemini 입금 매칭 분석 실패', { error: err.message });
    return { available: false, message: err.message };
  }
}

async function generateMonthlyReport(allData) {
  if (!genAI) {
    return { available: false, message: 'Gemini API 키 미설정' };
  }

  try {
    const model = genAI.getGenerativeModel({ model: 'gemini-2.0-flash' });

    const safeData = {
      salesCount: allData.sales?.length || 0,
      totalSales: (allData.sales || []).reduce((s, c) =>
        s + (c.items || []).reduce((si, i) => si + (i.total || 0), 0), 0),
      clientSummary: (allData.sales || []).slice(0, 20).map(c => ({
        client: c.client,
        total: (c.items || []).reduce((s, i) => s + (i.total || 0), 0),
        prevBalance: c.prevBalance || 0,
      })),
      depositCount: allData.deposits?.length || 0,
      totalDeposit: (allData.deposits || []).reduce((s, d) => s + (d.amount || 0), 0),
      comparisonSummary: allData.comparison?.summary || null,
    };

    const prompt = `회계 담당자를 위한 월간 매출/미수금 요약 리포트를 작성해주세요.

## 데이터
${JSON.stringify(safeData, null, 2)}

다음을 포함해주세요:
1. 이번 달 총매출 현황
2. 거래처별 미수금 현황
3. 주의 필요 항목 (큰 미수금, 장기 미수 등)
4. 입금 매칭 현황

한국어로 간결하게 작성해주세요.`;

    const result = await model.generateContent(prompt);
    return { available: true, report: result.response.text() };
  } catch (err) {
    logger.error('Gemini 리포트 생성 실패', { error: err.message });
    return { available: false, message: err.message };
  }
}

module.exports = { initGemini, analyzeComparison, suggestDepositMatches, generateMonthlyReport };
