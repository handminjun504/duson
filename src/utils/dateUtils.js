const { serialToDate, dateToSerial, formatDate } = require('../../config/template');

function getCurrentMonth() {
  const now = new Date();
  return {
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    label: `${now.getFullYear()}.${String(now.getMonth() + 1).padStart(2, '0')}`,
  };
}

function isInMonth(serial, year, month) {
  const date = serialToDate(serial);
  if (!date) return false;
  return date.getFullYear() === year && date.getMonth() + 1 === month;
}

function isCarryOver(serial, year, month) {
  const date = serialToDate(serial);
  if (!date) return false;
  const nextMonth = month === 12 ? 1 : month + 1;
  const nextYear = month === 12 ? year + 1 : year;
  return date.getFullYear() === nextYear && date.getMonth() + 1 === nextMonth;
}

module.exports = {
  serialToDate,
  dateToSerial,
  formatDate,
  getCurrentMonth,
  isInMonth,
  isCarryOver,
};
