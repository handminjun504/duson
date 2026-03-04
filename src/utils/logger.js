const LOG_LEVELS = { ERROR: 0, WARN: 1, INFO: 2, DEBUG: 3 };
let currentLevel = LOG_LEVELS.INFO;

function timestamp() {
  return new Date().toISOString().replace('T', ' ').slice(0, 19);
}

const logger = {
  setLevel(level) {
    currentLevel = LOG_LEVELS[level] ?? LOG_LEVELS.INFO;
  },
  error(msg, meta) {
    if (currentLevel >= LOG_LEVELS.ERROR)
      console.error(`[${timestamp()}] ERROR: ${msg}`, meta ?? '');
  },
  warn(msg, meta) {
    if (currentLevel >= LOG_LEVELS.WARN)
      console.warn(`[${timestamp()}] WARN: ${msg}`, meta ?? '');
  },
  info(msg, meta) {
    if (currentLevel >= LOG_LEVELS.INFO)
      console.log(`[${timestamp()}] INFO: ${msg}`, meta ?? '');
  },
  debug(msg, meta) {
    if (currentLevel >= LOG_LEVELS.DEBUG)
      console.log(`[${timestamp()}] DEBUG: ${msg}`, meta ?? '');
  },
};

module.exports = logger;
