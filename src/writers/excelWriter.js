const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');
const { COMPANY_INFO } = require('../../config/template');
const logger = require('../utils/logger');

const THIN_BORDER = {
  top: { style: 'thin' },
  left: { style: 'thin' },
  bottom: { style: 'thin' },
  right: { style: 'thin' },
};

const SENDER_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFF2F2F2' },
};

const TITLE_FONT = { name: '맑은 고딕', size: 22, bold: true };
const SENDER_FONT = { name: '맑은 고딕', size: 14, bold: true };
const HEADER_FONT = { name: '맑은 고딕', size: 14, bold: true };
const LABEL_FONT = { name: '맑은 고딕', size: 12, bold: true };
const DATA_FONT = { name: '맑은 고딕', size: 11 };
const NUM_FMT = '#,##0_ ';
const NUM_FMT_RED = '#,##0_);[Red](#,##0)';

const ROW = {
  TITLE: 1,
  SENDER: 2,
  HEADER: 3,
  PRIOR: 4,
  DATA_START: 5,
  DATA_END: 28,
  TOTAL: 29,
  DEPOSIT_START: 31,
  DEPOSIT_END: 41,
  OUTSTANDING: 42,
  BANK: 43,
};

const TOTAL_COLS = 14;
const COL_WIDTHS = [20.75, 24.13, 8.43, 10.38, 11.51, 10.26, 15.13, 11.88, 10.88, 12.01, 9.76, 11.13, 9.76, 12.63, 8.43, 10.76];
const HEADERS = ['매출처', '제품명', '수량', '단가', '공급가액', '부가세', '합계금액', '', '', '이월금액', '납품일', '수금액', '수금일', '비고'];

function setCell(ws, row, col, value, opts = {}) {
  const cell = ws.getRow(row).getCell(col);
  cell.value = value;
  cell.font = opts.font || DATA_FONT;
  cell.border = THIN_BORDER;
  if (opts.alignment) cell.alignment = opts.alignment;
  if (opts.fill) cell.fill = opts.fill;
  if (opts.numFmt) cell.numFmt = opts.numFmt;
  return cell;
}

function borderRow(ws, row, from, to) {
  for (let c = from; c <= to; c++) {
    ws.getRow(row).getCell(c).border = THIN_BORDER;
  }
}

function toDate(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  const s = String(val).trim();
  const m = s.match(/^(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function normalizeStr(s) {
  return (s || '').replace(/\s+/g, '').toLowerCase();
}

async function generateLedger(clientsData, options = {}) {
  const { month = new Date().getMonth() + 1, year = new Date().getFullYear() } = options;
  const monthStr = String(month).padStart(2, '0');

  const wb = new ExcelJS.Workbook();
  wb.creator = '두손푸드웨이';

  const usedNames = new Set();

  for (const clientData of clientsData) {
    const { client, items = [], deposits = [], priorBalance = 0, sheetNotes = {} } = clientData;

    let safeName = client.replace(/[\\/*?:\[\]]/g, '').slice(0, 31);
    if (usedNames.has(safeName)) {
      let suffix = 2;
      while (usedNames.has(`${safeName.slice(0, 28)}_${suffix}`)) suffix++;
      safeName = `${safeName.slice(0, 28)}_${suffix}`;
    }
    usedNames.add(safeName);

    const ws = wb.addWorksheet(safeName);
    COL_WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    // === R1: Title (merged A1:N1) ===
    ws.mergeCells(ROW.TITLE, 1, ROW.TITLE, TOTAL_COLS);
    setCell(ws, ROW.TITLE, 1, `${year}년 ${monthStr}월 매출 현황(${client})`, {
      font: TITLE_FONT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    ws.getRow(ROW.TITLE).height = 114.6;

    // === R2: Sender (merged A2:B2) ===
    ws.mergeCells(ROW.SENDER, 1, ROW.SENDER, 2);
    setCell(ws, ROW.SENDER, 1, `발신:   ${COMPANY_INFO.sender}`, {
      font: SENDER_FONT,
      fill: SENDER_FILL,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    borderRow(ws, ROW.SENDER, 3, TOTAL_COLS);
    ws.getRow(ROW.SENDER).height = 40.9;

    // === R3: Headers ===
    for (let c = 1; c <= TOTAL_COLS; c++) {
      setCell(ws, ROW.HEADER, c, HEADERS[c - 1], {
        font: HEADER_FONT,
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
      });
    }
    ws.getRow(ROW.HEADER).height = 50.45;

    // === R4: Prior balance ===
    setCell(ws, ROW.PRIOR, 1, '전기 외상매출금 미납액', {
      font: LABEL_FONT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    borderRow(ws, ROW.PRIOR, 2, TOTAL_COLS);
    if (priorBalance) {
      setCell(ws, ROW.PRIOR, 7, priorBalance, {
        font: { ...DATA_FONT, bold: true },
        numFmt: NUM_FMT,
        alignment: { horizontal: 'right' },
      });
    }
    ws.getRow(ROW.PRIOR).height = 50.45;

    // === R5~R28: Data rows (fixed 24 rows) ===
    const maxData = ROW.DATA_END - ROW.DATA_START + 1;
    if (items.length > maxData) {
      logger.warn(`${client}: 데이터 ${items.length}건이 최대 ${maxData}행 초과 — 초과분 생략`);
    }

    for (let i = 0; i < maxData; i++) {
      const r = ROW.DATA_START + i;
      const item = items[i];

      if (item) {
        // A: 매출처 (첫 번째 행에만)
        if (i === 0) {
          setCell(ws, r, 1, client, { alignment: { horizontal: 'left', vertical: 'middle' } });
        } else {
          ws.getRow(r).getCell(1).border = THIN_BORDER;
        }

        // B: 제품명
        setCell(ws, r, 2, item.product || '', { alignment: { horizontal: 'center' } });
        // C: 수량
        setCell(ws, r, 3, item.qty || '', { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
        // D: 단가
        setCell(ws, r, 4, item.unitPrice || '', { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
        // E: 공급가액 = C * D
        setCell(ws, r, 5, { formula: `C${r}*D${r}` }, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
        // F: 부가세 = E * 10%
        setCell(ws, r, 6, { formula: `E${r}*10%` }, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
        // G: 합계금액 = E + F
        setCell(ws, r, 7, { formula: `E${r}+F${r}` }, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });

        // H~J: 수금일/수금액/이월금액 (빈 칸)
        borderRow(ws, r, 8, 10);

        // K(11): 납품일
        const dateObj = toDate(item.deliveryDate);
        if (dateObj) {
          setCell(ws, r, 11, dateObj, { alignment: { horizontal: 'center' }, numFmt: 'YYYY-MM-DD' });
        } else {
          ws.getRow(r).getCell(11).border = THIN_BORDER;
        }

        // L~M: 수금액/수금일 (빈 칸)
        borderRow(ws, r, 12, 13);

        // N(14): 비고 — 시트 AE열 데이터 우선, 없으면 경리나라 비고
        const np = normalizeStr(item.product);
        const note = sheetNotes[np] || item.note || '';
        setCell(ws, r, 14, note, { alignment: { horizontal: 'center' } });
      } else {
        borderRow(ws, r, 1, TOTAL_COLS);
      }

      if (i === 0) ws.getRow(r).height = 22.15;
      else if (i === maxData - 1) ws.getRow(r).height = 18;
      else ws.getRow(r).height = 17.3;
    }

    // === R29: 월 합계 (SUM 수식) ===
    setCell(ws, ROW.TOTAL, 1, `${monthStr}월합계`, {
      font: LABEL_FONT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    borderRow(ws, ROW.TOTAL, 2, 4);
    setCell(ws, ROW.TOTAL, 5, { formula: `SUM(E${ROW.DATA_START}:E${ROW.DATA_END})` }, {
      font: { ...DATA_FONT, bold: true }, numFmt: NUM_FMT, alignment: { horizontal: 'center', vertical: 'middle' },
    });
    setCell(ws, ROW.TOTAL, 6, { formula: `SUM(F${ROW.DATA_START}:F${ROW.DATA_END})` }, {
      font: { ...DATA_FONT, bold: true }, numFmt: NUM_FMT, alignment: { horizontal: 'center', vertical: 'middle' },
    });
    setCell(ws, ROW.TOTAL, 7, { formula: `SUM(G${ROW.DATA_START}:G${ROW.DATA_END})` }, {
      font: { ...DATA_FONT, bold: true }, numFmt: NUM_FMT, alignment: { horizontal: 'center', vertical: 'middle' },
    });
    borderRow(ws, ROW.TOTAL, 8, TOTAL_COLS);
    ws.getRow(ROW.TOTAL).height = 18.8;

    // === R30: blank row ===
    borderRow(ws, 30, 1, TOTAL_COLS);
    ws.getRow(30).height = 18.8;

    // === R31~R41: 입금 rows (11 rows) ===
    const maxDeposits = ROW.DEPOSIT_END - ROW.DEPOSIT_START + 1;
    for (let i = 0; i < maxDeposits; i++) {
      const r = ROW.DEPOSIT_START + i;
      const dep = deposits[i];

      // A: "입금" (첫 행만)
      if (i === 0) {
        setCell(ws, r, 1, '입금', { font: LABEL_FONT, alignment: { horizontal: 'center', vertical: 'middle' } });
      } else {
        ws.getRow(r).getCell(1).border = THIN_BORDER;
      }

      // B~K: 빈 칸
      borderRow(ws, r, 2, 11);

      if (dep) {
        // L(12): 수금액
        setCell(ws, r, 12, dep.amount || 0, { numFmt: NUM_FMT_RED, alignment: { horizontal: 'right', vertical: 'middle' } });
        // M(13): 수금일
        const depDate = toDate(dep.date);
        if (depDate) {
          setCell(ws, r, 13, depDate, { alignment: { horizontal: 'center' }, numFmt: 'YYYY-MM-DD' });
        } else {
          ws.getRow(r).getCell(13).border = THIN_BORDER;
        }
      } else {
        borderRow(ws, r, 12, 13);
      }

      // N: 빈 칸
      ws.getRow(r).getCell(14).border = THIN_BORDER;

      // P(16): 입금 합계 SUM (첫 행에만)
      if (i === 0) {
        setCell(ws, r, 16, { formula: `SUM(L${ROW.DEPOSIT_START}:L${ROW.DEPOSIT_END})` }, {
          numFmt: NUM_FMT_RED, alignment: { horizontal: 'right', vertical: 'middle' },
        });
      }

      if (i === 0) ws.getRow(r).height = 18.8;
      else if (i === maxDeposits - 1) ws.getRow(r).height = 18;
      else if (i === 1) ws.getRow(r).height = 18;
      else ws.getRow(r).height = 17.3;
    }

    // === R42: 총외상매출금미납액 ===
    setCell(ws, ROW.OUTSTANDING, 1, '총외상매출금미납액', {
      font: LABEL_FONT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    borderRow(ws, ROW.OUTSTANDING, 2, 6);

    const depositCells = Array.from(
      { length: maxDeposits },
      (_, i) => `L${ROW.DEPOSIT_START + i}`,
    ).join('-');
    setCell(ws, ROW.OUTSTANDING, 7, { formula: `G${ROW.PRIOR}+G${ROW.TOTAL}-${depositCells}` }, {
      font: { ...DATA_FONT, bold: true },
      numFmt: NUM_FMT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    borderRow(ws, ROW.OUTSTANDING, 8, TOTAL_COLS);
    ws.getRow(ROW.OUTSTANDING).height = 53.45;

    // === R43: 입금계좌 ===
    setCell(ws, ROW.BANK, 1, '입금계좌', {
      font: SENDER_FONT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    setCell(ws, ROW.BANK, 2, COMPANY_INFO.bankAccount, {
      font: SENDER_FONT,
      alignment: { horizontal: 'left', vertical: 'middle' },
    });
    borderRow(ws, ROW.BANK, 3, TOTAL_COLS);
    ws.getRow(ROW.BANK).height = 20.25;
  }

  const outputDir = path.resolve(__dirname, '../../output');
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

  const filename = `거래처원장_${year}${monthStr}.xlsx`;
  const outputPath = path.join(outputDir, filename);
  await wb.xlsx.writeFile(outputPath);

  logger.info(`거래처원장 생성 완료: ${outputPath} (${clientsData.length}개 거래처)`);
  return { path: outputPath, filename, clientCount: clientsData.length };
}

module.exports = { generateLedger };
