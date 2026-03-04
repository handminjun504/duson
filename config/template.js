const EXCEL_SERIAL_EPOCH = new Date(1899, 11, 30);

function serialToDate(serial) {
  if (!serial || typeof serial !== 'number') return null;
  const d = new Date(EXCEL_SERIAL_EPOCH.getTime() + serial * 86400000);
  return d;
}

function dateToSerial(date) {
  if (!(date instanceof Date)) date = new Date(date);
  return Math.round((date.getTime() - EXCEL_SERIAL_EPOCH.getTime()) / 86400000);
}

function formatDate(date) {
  if (!date) return '';
  if (typeof date === 'number') date = serialToDate(date);
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

const COLUMNS = {
  A: { index: 0, name: '매출처' },
  B: { index: 1, name: '제품명' },
  C: { index: 2, name: '수량' },
  D: { index: 3, name: '단가' },
  E: { index: 4, name: '공급가액' },
  F: { index: 5, name: '부가세' },
  G: { index: 6, name: '합계금액' },
  H: { index: 7, name: '수금일' },
  I: { index: 8, name: '수금액' },
  J: { index: 9, name: '이월금액' },
  K: { index: 10, name: '납품일' },
  L: { index: 11, name: '수금액2' },
  M: { index: 12, name: '수금일2' },
  N: { index: 13, name: '비고' },
};

const LAYOUT = {
  titleRow: 0,
  senderRow: 1,
  headerRow: 2,
  priorBalanceRow: 3,
  dataStartRow: 4,
  summaryLabel: '2월합계',
  depositLabel: '입금',
  totalLabel: '총외상매출금미납액',
  bankInfoLabel: '입금계좌',
};

const COMPANY_INFO = {
  sender: '주식회사 두손푸드웨이',
  bankAccount: '하나은행429-910017-29204',
};

module.exports = {
  COLUMNS,
  LAYOUT,
  COMPANY_INFO,
  serialToDate,
  dateToSerial,
  formatDate,
};
