const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
const logger = require('../utils/logger');

const HEADER_KEYWORDS = ['거래일자', '품목명', '수량', '단가', '거래처명'];

const COL_ALIASES = {
  '거래일자': 'tradeDate',
  '구분': 'type',
  '매출거래처명': 'salesClient',
  '매출 거래처명': 'salesClient',
  '사업자번호': '_bizNo',
  '매입거래처명': 'purchaseClient',
  '매입 거래처명': 'purchaseClient',
  '부가세구분': 'vatType',
  '부가세\n구분': 'vatType',
  '프로젝트현장': 'project',
  '프로젝트\n현장': 'project',
  '창고': 'warehouse',
  '품목월일': 'itemDate',
  '품목\n월일': 'itemDate',
  '품목코드': 'itemCode',
  '품목\n코드': 'itemCode',
  '품목명': 'productName',
  '규격': 'spec',
  '수량': 'qty',
  '단위': 'unit',
  '매출단가': 'unitPrice',
  '매출\n단가': 'unitPrice',
  '매출공급가액': 'supplyAmount',
  '매출\n공급가액': 'supplyAmount',
  '매출세액': 'vat',
  '매출 세액': 'vat',
  '매출합계금액': 'totalAmount',
  '매출\n합계금액': 'totalAmount',
  '비고': 'note',
  '변경항목': 'note',
  '변경': '_change',
};

function normalizeHeader(h) {
  return (h || '').replace(/\s+/g, ' ').trim();
}

function findAlias(rawHeader) {
  if (COL_ALIASES[rawHeader]) return COL_ALIASES[rawHeader];
  const normalized = rawHeader.replace(/[\s\n]+/g, '');
  for (const [key, alias] of Object.entries(COL_ALIASES)) {
    if (key.replace(/[\s\n]+/g, '') === normalized) return alias;
  }
  return null;
}

function pn(s) {
  return parseFloat((s || '').replace(/[,\s]/g, '')) || 0;
}

function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    if (!row) continue;
    const joined = row.join(' ');
    const matchCount = HEADER_KEYWORDS.filter(kw => joined.includes(kw)).length;
    if (matchCount >= 3) {
      logger.info(`헤더 행 발견: row ${i + 1} (매칭 키워드 ${matchCount}개)`);
      return i;
    }
  }
  return 0;
}

async function getAuthClient() {
  const keyPath = path.resolve(process.env.GOOGLE_SERVICE_ACCOUNT_KEY_PATH || './config/google-credentials.json');

  if (!fs.existsSync(keyPath)) {
    throw new Error(`Google 서비스 계정 키 파일을 찾을 수 없습니다: ${keyPath}\nconfig/google-credentials.json에 서비스 계정 키를 배치하세요.`);
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: keyPath,
    scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
  });

  return auth.getClient();
}

function toYMD(dateStr) {
  if (!dateStr) return '';
  const s = dateStr.replace(/[-./]/g, '');
  if (/^\d{8}$/.test(s)) return s;
  return '';
}

async function fetchSalesData(options = {}) {
  const { startDate, endDate } = options;
  const spreadsheetId = process.env.GOOGLE_SHEETS_SPREADSHEET_ID;
  if (!spreadsheetId) {
    throw new Error('GOOGLE_SHEETS_SPREADSHEET_ID가 .env에 설정되지 않았습니다');
  }

  const filterStart = toYMD(startDate);
  const filterEnd = toYMD(endDate);
  const hasFilter = filterStart || filterEnd;
  logger.info(`Google Sheets에서 매출 데이터 수집 시작... ${hasFilter ? `(${startDate} ~ ${endDate})` : '(전체)'}`);

  const client = await getAuthClient();
  const sheets = google.sheets({ version: 'v4', auth: client });

  const targetGid = process.env.GOOGLE_SHEETS_GID;
  const meta = await sheets.spreadsheets.get({ spreadsheetId });
  const allSheets = meta.data.sheets;

  let targetSheets = allSheets;
  if (targetGid) {
    const gidSheet = allSheets.find(s => String(s.properties.sheetId) === String(targetGid));
    if (gidSheet) {
      targetSheets = [gidSheet];
      logger.info(`GID ${targetGid} 시트 선택: "${gidSheet.properties.title}"`);
    }
  }

  const sheetNames = targetSheets.map(s => s.properties.title);
  logger.info(`대상 시트: ${sheetNames.join(', ')}`);

  const allData = [];

  for (const sheetName of sheetNames) {
    const targetSheet = targetSheets.find(s => s.properties.title === sheetName);
    const maxRow = targetSheet?.properties?.gridProperties?.rowCount || 5000;
    const maxCol = targetSheet?.properties?.gridProperties?.columnCount || 20;
    const lastCol = String.fromCharCode(64 + Math.min(maxCol, 26));
    const mainRange = `'${sheetName}'!A1:${lastCol}${maxRow}`;
    logger.info(`시트 범위: ${mainRange} (${maxRow}행 x ${maxCol}열)`);

    const mainRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: mainRange,
    });

    let noteRows = [];
    if (maxCol >= 31) {
      try {
        const noteRes = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `'${sheetName}'!AE1:AE${maxRow}`,
        });
        noteRows = noteRes.data.values || [];
      } catch (err) {
        logger.warn(`AE열 비고 조회 실패 (무시): ${err.message}`);
      }
    }

    const rows = mainRes.data.values || [];
    if (rows.length < 3) continue;

    const headerIdx = findHeaderRow(rows);
    const headerRow = rows[headerIdx];

    const colMap = {};
    const rawHeaders = [];
    for (let i = 0; i < headerRow.length; i++) {
      const raw = (headerRow[i] || '').trim();
      rawHeaders.push(raw);
      const alias = findAlias(raw);
      if (alias) {
        colMap[alias] = i;
      }
    }

    logger.info(`헤더 매핑: ${JSON.stringify(colMap)}`);
    logger.info(`원본 헤더 (${rawHeaders.length}개): ${rawHeaders.map((h, i) => `${String.fromCharCode(65 + i)}=${h || '(빈)'}`).join(', ')}`);

    if (!colMap.productName && !colMap.salesClient) {
      logger.warn(`시트 "${sheetName}": 필수 컬럼(품목명/거래처명) 없음 - 스킵`);
      continue;
    }

    const noteHeaderVal = (noteRows[headerIdx]?.[0] || '').trim();
    logger.info(`AE열 헤더: "${noteHeaderVal}", AE열 데이터: ${noteRows.length}행`);

    const dataRows = rows.slice(headerIdx + 1);
    let parsed = 0;
    let skipped = 0;

    for (let rowIdx = 0; rowIdx < dataRows.length; rowIdx++) {
      const row = dataRows[rowIdx];
      if (!row || row.every(c => !c || !c.trim())) { skipped++; continue; }

      const get = (field) => {
        const idx = colMap[field];
        return idx !== undefined ? (row[idx] || '').trim() : '';
      };

      const productName = get('productName');
      const salesClient = get('salesClient');

      if (!productName && !salesClient) { skipped++; continue; }

      const tradeDate = get('tradeDate');

      if (hasFilter) {
        const ymd = toYMD(tradeDate);
        if (ymd) {
          if (filterStart && ymd < filterStart) { skipped++; continue; }
          if (filterEnd && ymd > filterEnd) { skipped++; continue; }
        }
      }

      const record = {
        tradeDate,
        type: get('type'),
        salesClient,
        purchaseClient: get('purchaseClient'),
        vatType: get('vatType'),
        project: get('project'),
        warehouse: get('warehouse'),
        itemDate: get('itemDate'),
        itemCode: get('itemCode'),
        productName,
        spec: get('spec'),
        qty: pn(get('qty')),
        unit: get('unit'),
        unitPrice: pn(get('unitPrice')),
        supplyAmount: pn(get('supplyAmount')),
        vat: pn(get('vat')),
        totalAmount: pn(get('totalAmount')),
        note: get('note') || (noteRows[headerIdx + 1 + rowIdx]?.[0] || '').trim(),
        _sheet: sheetName,
        _raw: row.map(c => (c || '').trim()),
      };

      if (!record.totalAmount && record.supplyAmount) {
        record.totalAmount = record.supplyAmount + record.vat;
      }

      allData.push(record);
      parsed++;
    }

    logger.info(`시트 "${sheetName}": ${parsed}건 파싱, ${skipped}건 스킵${hasFilter ? ` (날짜 필터: ${filterStart}~${filterEnd})` : ''}`);
  }

  logger.info(`Google Sheets에서 총 ${allData.length}건 매출 데이터 수집 완료${hasFilter ? ` (${startDate} ~ ${endDate})` : ''}`);
  return allData;
}

module.exports = { fetchSalesData };
