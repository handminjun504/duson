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

const HEADER_FILL = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFF2F2F2' },
};

const HEADER_FONT = { name: '맑은 고딕', size: 10, bold: true };
const DATA_FONT = { name: '맑은 고딕', size: 10 };
const TITLE_FONT = { name: '맑은 고딕', size: 16, bold: true };
const NUM_FMT = '#,##0';

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

async function generateLedger(clientsData, options = {}) {
  const { month = new Date().getMonth() + 1, year = new Date().getFullYear() } = options;
  const monthStr = String(month).padStart(2, '0');

  const wb = new ExcelJS.Workbook();
  wb.creator = '두손푸드웨이';

  const COL_WIDTHS = [14, 18, 8, 10, 13, 12, 14, 10, 12, 12, 10, 10, 10, 10];
  const HEADERS = ['매출처', '제품명', '수량', '단가', '공급가액', '부가세', '합계금액', '납품일', '수금액', '수금일', '이월금액', '수금액', '수금일', '비고'];
  const TOTAL_COLS = 14;

  const usedNames = new Set();

  for (const clientData of clientsData) {
    const { client, items, deposits = [], priorBalance = 0 } = clientData;

    let safeName = client.replace(/[\\/*?:\[\]]/g, '').slice(0, 31);
    if (usedNames.has(safeName)) {
      let suffix = 2;
      while (usedNames.has(`${safeName.slice(0, 28)}_${suffix}`)) suffix++;
      safeName = `${safeName.slice(0, 28)}_${suffix}`;
    }
    usedNames.add(safeName);

    const ws = wb.addWorksheet(safeName);

    COL_WIDTHS.forEach((w, i) => { ws.getColumn(i + 1).width = w; });

    let r = 1;

    ws.mergeCells(r, 1, r, TOTAL_COLS);
    const titleCell = setCell(ws, r, 1, `${year}년 ${monthStr}월 매출 현황(${client})`, {
      font: TITLE_FONT,
      alignment: { horizontal: 'center', vertical: 'middle' },
    });
    ws.getRow(r).height = 36;
    r++;

    r++;
    ws.mergeCells(r, 1, r, 3);
    setCell(ws, r, 1, `발신:   ${COMPANY_INFO.sender}`, {
      font: { ...DATA_FONT, bold: true },
      alignment: { horizontal: 'left' },
    });
    for (let c = 4; c <= TOTAL_COLS; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
    r++;

    r++;
    for (let c = 1; c <= TOTAL_COLS; c++) {
      setCell(ws, r, c, HEADERS[c - 1], {
        font: HEADER_FONT,
        fill: HEADER_FILL,
        alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
      });
    }
    ws.getRow(r).height = 24;
    const headerRow = r;
    r++;

    ws.mergeCells(r, 1, r, 3);
    setCell(ws, r, 1, '전기 외상매출금 미납액', {
      font: { ...DATA_FONT, bold: true },
      alignment: { horizontal: 'left' },
    });
    for (let c = 4; c <= TOTAL_COLS; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
    r++;

    let totalSupply = 0;
    let totalVat = 0;
    let totalAmount = 0;

    if (items.length === 0) {
      setCell(ws, r, 1, client, { alignment: { horizontal: 'left' } });
      setCell(ws, r, 2, '(거래 없음)', { alignment: { horizontal: 'center' } });
      for (let c = 3; c <= TOTAL_COLS; c++) ws.getRow(r).getCell(c).border = THIN_BORDER;
      r++;
    }

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      const supply = item.supplyAmount || item.qty * item.unitPrice;
      const vat = item.vat != null ? item.vat : Math.round(supply * 0.1);
      const total = item.total || supply + vat;

      if (i === 0) {
        setCell(ws, r, 1, client, { alignment: { horizontal: 'left', vertical: 'middle' } });
      } else {
        ws.getRow(r).getCell(1).border = THIN_BORDER;
      }

      setCell(ws, r, 2, item.product, { alignment: { horizontal: 'center' } });
      setCell(ws, r, 3, item.qty, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
      setCell(ws, r, 4, item.unitPrice, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
      setCell(ws, r, 5, supply, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
      setCell(ws, r, 6, vat, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
      setCell(ws, r, 7, total, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });

      if (item.deliveryDate) {
        const dateStr = String(item.deliveryDate).replace(/^\d{4}[-/]/, '').replace(/[-/]/g, '/');
        setCell(ws, r, 8, dateStr, { alignment: { horizontal: 'center' } });
      } else {
        ws.getRow(r).getCell(8).border = THIN_BORDER;
      }

      for (let c = 9; c <= 13; c++) {
        ws.getRow(r).getCell(c).border = THIN_BORDER;
      }

      setCell(ws, r, 14, item.note || '', { alignment: { horizontal: 'center' } });

      totalSupply += supply;
      totalVat += vat;
      totalAmount += total;
      r++;
    }

    if (items.length > 1) {
      ws.mergeCells(headerRow + 2, 1, headerRow + 1 + items.length, 1);
    }

    r++;
    r++;
    setCell(ws, r, 1, `${monthStr}월합계`, {
      font: { ...DATA_FONT, bold: true },
      alignment: { horizontal: 'center' },
    });
    ws.getRow(r).getCell(2).border = THIN_BORDER;
    ws.getRow(r).getCell(3).border = THIN_BORDER;
    ws.getRow(r).getCell(4).border = THIN_BORDER;
    setCell(ws, r, 5, totalSupply, { font: { ...DATA_FONT, bold: true }, numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
    setCell(ws, r, 6, totalVat, { font: { ...DATA_FONT, bold: true }, numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
    setCell(ws, r, 7, totalAmount, { font: { ...DATA_FONT, bold: true }, numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
    for (let c = 8; c <= TOTAL_COLS; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
    r++;

    setCell(ws, r, 1, '입금', {
      font: { ...DATA_FONT, bold: true },
      alignment: { horizontal: 'left' },
    });
    for (let c = 2; c <= TOTAL_COLS; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
    r++;

    let totalDeposit = 0;
    if (deposits.length > 0) {
      for (const dep of deposits) {
        for (let c = 1; c <= 7; c++) {
          ws.getRow(r).getCell(c).border = THIN_BORDER;
        }
        setCell(ws, r, 8, dep.date || '', { alignment: { horizontal: 'center' } });
        setCell(ws, r, 9, dep.amount || 0, { numFmt: NUM_FMT, alignment: { horizontal: 'right' } });
        for (let c = 10; c <= TOTAL_COLS; c++) {
          ws.getRow(r).getCell(c).border = THIN_BORDER;
        }
        if (dep.client) {
          setCell(ws, r, 14, dep.client, { alignment: { horizontal: 'center' } });
        }
        totalDeposit += dep.amount || 0;
        r++;
      }
    }

    for (let pad = 0; pad < 5; pad++) {
      for (let c = 1; c <= TOTAL_COLS; c++) {
        ws.getRow(r).getCell(c).border = THIN_BORDER;
      }
      r++;
    }

    r++;
    setCell(ws, r, 1, '총외상매출금미납액', {
      font: { ...DATA_FONT, bold: true },
      alignment: { horizontal: 'left' },
    });
    for (let c = 2; c <= 6; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
    setCell(ws, r, 7, priorBalance + totalAmount - totalDeposit, {
      font: { ...DATA_FONT, bold: true },
      numFmt: NUM_FMT,
      alignment: { horizontal: 'right' },
    });
    for (let c = 8; c <= TOTAL_COLS; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
    r++;

    setCell(ws, r, 1, '입금계좌', {
      font: { ...DATA_FONT, bold: true },
      alignment: { horizontal: 'center' },
    });
    ws.mergeCells(r, 2, r, 5);
    setCell(ws, r, 2, COMPANY_INFO.bankAccount, {
      font: DATA_FONT,
      alignment: { horizontal: 'left' },
    });
    for (let c = 6; c <= TOTAL_COLS; c++) {
      ws.getRow(r).getCell(c).border = THIN_BORDER;
    }
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
