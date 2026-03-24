const express = require('express');
const path = require('path');
const fs = require('fs');
const router = express.Router();

const googleSheets = require('../collectors/googleSheets');
const gyeongliNara = require('../collectors/gyeongliNara');
const matcher = require('../processors/matcher');
const gemini = require('../processors/geminiAnalyzer');
const excelWriter = require('../writers/excelWriter');
const logger = require('../utils/logger');

const STATE_FILE = path.join(__dirname, '..', '..', '.appstate.json');

function loadState() {
  try {
    if (fs.existsSync(STATE_FILE)) {
      const raw = fs.readFileSync(STATE_FILE, 'utf8');
      const saved = JSON.parse(raw);
      logger.info(`저장된 상태 복원: Sheet ${saved.sheetData?.length || 0}건, 경리나라 ${saved.gyeongliSales?.length || 0}개 거래처, 입금 ${saved.gyeongliDeposits?.length || 0}건`);
      return { ...saved, comparison: null, depositMatch: null, geminiAnalysis: null, lastGenerated: null };
    }
  } catch (err) {
    logger.warn(`상태 복원 실패: ${err.message}`);
  }
  return { sheetData: null, gyeongliSales: null, gyeongliDeposits: null, comparison: null, depositMatch: null, geminiAnalysis: null, lastGenerated: null };
}

function saveState() {
  try {
    const toSave = {
      sheetData: appState.sheetData,
      gyeongliSales: appState.gyeongliSales,
      gyeongliDeposits: appState.gyeongliDeposits,
      savedAt: new Date().toISOString(),
    };
    fs.writeFileSync(STATE_FILE, JSON.stringify(toSave), 'utf8');
  } catch (err) {
    logger.warn(`상태 저장 실패: ${err.message}`);
  }
}

let appState = loadState();

router.get('/status', (req, res) => {
  res.json({
    sheetCollected: !!appState.sheetData,
    sheetCount: appState.sheetData?.length || 0,
    gyeongliCollected: !!appState.gyeongliSales,
    gyeongliClientCount: appState.gyeongliSales?.length || 0,
    gyeongliItemCount: appState.gyeongliSales?.reduce((s, c) => s + c.items.length, 0) || 0,
    depositCount: appState.gyeongliDeposits?.length || 0,
    compared: !!appState.comparison,
    matchedCount: appState.comparison?.summary?.matchedCount || 0,
    mismatchCount: appState.comparison?.summary?.mismatchCount || 0,
    sheetOnlyCount: appState.comparison?.summary?.sheetOnlyCount || 0,
    gyeongliOnlyCount: appState.comparison?.summary?.gyeongliOnlyCount || 0,
    geminiAvailable: !!process.env.GEMINI_API_KEY,
    lastGenerated: appState.lastGenerated,
  });
});

router.post('/collect/sheets', async (req, res) => {
  try {
    const { startDate, endDate } = req.body || {};
    appState.sheetData = await googleSheets.fetchSalesData({ startDate, endDate });
    saveState();
    res.json({ success: true, count: appState.sheetData.length });
  } catch (err) {
    logger.error('Google Sheets 수집 실패', { error: err.message });
    res.status(500).json({ success: false, error: err.message });
  }
});

router.post('/collect/gyeongli', async (req, res) => {
  try {
    const { startDate, endDate } = req.body || {};
    const sales = await gyeongliNara.collectSalesData({ startDate, endDate });
    appState.gyeongliSales = sales;
    const deposits = await gyeongliNara.collectDepositData({ startDate, endDate });
    appState.gyeongliDeposits = deposits;
    saveState();
    res.json({
      success: true,
      clients: sales.map(c => ({
        name: c.client,
        itemCount: c.items.length,
        totalAmount: c.items.reduce((s, i) => s + i.total, 0),
      })),
      depositCount: deposits.length,
    });
  } catch (err) {
    logger.error('경리나라 수집 실패', { error: err.message });
    res.status(500).json({ success: false, error: err.message });
  }
});

router.get('/compare', (req, res) => {
  if (!appState.sheetData || !appState.gyeongliSales) {
    return res.status(400).json({
      error: '양쪽 데이터를 먼저 수집하세요 (Google Sheet + 경리나라)',
      summary: { sheetCount: appState.sheetData?.length || 0, gyeongliCount: 0 },
    });
  }
  appState.comparison = matcher.compareSalesData(appState.sheetData, appState.gyeongliSales);
  saveState();
  res.json(appState.comparison);
});

router.post('/analyze', async (req, res) => {
  try {
    const analysis = await gemini.analyzeComparison(appState.sheetData, appState.gyeongliSales);
    appState.geminiAnalysis = analysis;
    res.json(analysis);
  } catch (err) {
    logger.error('Gemini 분석 실패', { error: err.message });
    res.status(500).json({ success: false, error: err.message });
  }
});

router.get('/clients', (req, res) => {
  if (!appState.gyeongliSales) {
    return res.json({ clients: [] });
  }

  const clientMap = {};
  for (const c of appState.gyeongliSales) {
    const name = c.client;
    if (!clientMap[name]) {
      clientMap[name] = {
        name,
        items: [],
        prevBalance: 0,
        depositAmount: 0,
        outstandingBalance: 0,
        dates: new Set(),
      };
    }
    clientMap[name].items.push(...(c.items || []));
    clientMap[name].prevBalance = clientMap[name].prevBalance || c.prevBalance || 0;
    clientMap[name].depositAmount += c.depositAmount || 0;
    clientMap[name].outstandingBalance = c.outstandingBalance || clientMap[name].outstandingBalance;
    if (c.date) clientMap[name].dates.add(c.date);
    c.items?.forEach(i => { if (i.deliveryDate) clientMap[name].dates.add(i.deliveryDate); });
  }

  const clients = Object.values(clientMap).map(c => ({
    name: c.name,
    itemCount: c.items.length,
    dateCount: c.dates.size,
    totalAmount: c.items.reduce((s, i) => s + (i.total || 0), 0),
    totalSupply: c.items.reduce((s, i) => s + (i.supplyAmount || 0), 0),
    totalVat: c.items.reduce((s, i) => s + (i.vat || 0), 0),
    prevBalance: c.prevBalance,
    depositAmount: c.depositAmount,
    outstandingBalance: c.outstandingBalance,
  }));

  res.json({ clients });
});

router.get('/clients/:name/transactions', (req, res) => {
  const clientName = decodeURIComponent(req.params.name);
  const matches = appState.gyeongliSales?.filter(c => c.client === clientName) || [];
  if (matches.length === 0) {
    return res.status(404).json({ error: '거래처를 찾을 수 없습니다' });
  }

  const allItems = [];
  let prevBalance = 0;
  let depositAmount = 0;
  let outstandingBalance = 0;

  for (const m of matches) {
    allItems.push(...(m.items || []));
    prevBalance = prevBalance || m.prevBalance || 0;
    depositAmount += m.depositAmount || 0;
    outstandingBalance = m.outstandingBalance || outstandingBalance;
  }

  const deposits = appState.gyeongliDeposits?.filter(d => {
    const dClient = (d.client || '').replace(/\s/g, '');
    const tClient = clientName.replace(/\s/g, '');
    return dClient.includes(tClient) || tClient.includes(dClient) ||
           (d.allCells || []).some(cell => cell.includes(clientName));
  }) || [];

  res.json({
    client: clientName,
    items: allItems,
    prevBalance,
    depositAmount,
    outstandingBalance,
    deposits,
    totalAmount: allItems.reduce((s, i) => s + (i.total || 0), 0),
  });
});

router.get('/deposits', (req, res) => {
  let matchResults = appState.depositMatch;
  if (!matchResults && appState.gyeongliDeposits?.length && appState.gyeongliSales?.length) {
    const basicMatch = matcher.matchDeposits(appState.gyeongliSales, appState.gyeongliDeposits);
    matchResults = { basic: basicMatch, gemini: null };
  }
  res.json({
    deposits: appState.gyeongliDeposits || [],
    matchResults,
  });
});

router.post('/match', async (req, res) => {
  try {
    const basicMatch = matcher.matchDeposits(appState.gyeongliSales, appState.gyeongliDeposits);
    const geminiMatch = await gemini.suggestDepositMatches(appState.gyeongliDeposits, appState.gyeongliSales);
    appState.depositMatch = { basic: basicMatch, gemini: geminiMatch };
    res.json(appState.depositMatch);
  } catch (err) {
    logger.error('입금 매칭 실패', { error: err.message });
    res.status(500).json({ success: false, error: err.message });
  }
});

router.post('/generate', async (req, res) => {
  try {
    if (!appState.gyeongliSales || appState.gyeongliSales.length === 0) {
      return res.status(400).json({ error: '먼저 경리나라에서 데이터를 수집하세요' });
    }

    const merged = {};
    for (const c of appState.gyeongliSales) {
      if (!merged[c.client]) {
        merged[c.client] = { client: c.client, items: [], priorBalance: 0 };
      }
      merged[c.client].items.push(...(c.items || []));
      merged[c.client].priorBalance = merged[c.client].priorBalance || c.prevBalance || 0;
    }
    const clientsData = Object.values(merged).map(c => ({
      ...c,
      deposits: appState.gyeongliDeposits?.filter(d => {
        const dc = (d.client || '').replace(/\s/g, '');
        const tc = c.client.replace(/\s/g, '');
        return dc.includes(tc) || tc.includes(dc);
      }) || [],
    }));

    const result = await excelWriter.generateLedger(clientsData);
    appState.lastGenerated = { ...result, timestamp: new Date().toISOString() };
    res.json({ success: true, ...result });
  } catch (err) {
    logger.error('Excel 생성 실패', { error: err.message });
    res.status(500).json({ success: false, error: err.message });
  }
});

router.get('/download', (req, res) => {
  if (!appState.lastGenerated) {
    return res.status(404).json({ error: '생성된 파일이 없습니다. 먼저 Excel을 생성하세요.' });
  }

  const filePath = appState.lastGenerated.path;
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: '파일을 찾을 수 없습니다' });
  }

  res.download(filePath, appState.lastGenerated.filename);
});

router.post('/report', async (req, res) => {
  try {
    const report = await gemini.generateMonthlyReport({
      sales: appState.gyeongliSales,
      deposits: appState.gyeongliDeposits,
      comparison: appState.comparison,
    });
    res.json(report);
  } catch (err) {
    logger.error('리포트 생성 실패', { error: err.message });
    res.status(500).json({ success: false, error: err.message });
  }
});

module.exports = router;
