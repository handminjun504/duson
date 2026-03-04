const dotenv = require('dotenv');
const result = dotenv.config();
if (result.error) {
  console.error('.env 로드 실패:', result.error.message);
} else {
  console.log('.env 로드 성공. 환경변수:', Object.keys(result.parsed).join(', '));
  console.log('GYEONGLI_ID:', process.env.GYEONGLI_ID ? '[설정됨]' : '[비어있음]');
  console.log('GYEONGLI_PW:', process.env.GYEONGLI_PW ? '[설정됨]' : '[비어있음]');
}

const express = require('express');
const cors = require('cors');
const path = require('path');
const logger = require('./src/utils/logger');
const apiRoutes = require('./src/routes/api');
const { initGemini } = require('./src/processors/geminiAnalyzer');

const app = express();
const PORT = process.env.PORT || 3055;
const HOST = process.env.HOST || '0.0.0.0';

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.use('/api', apiRoutes);

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.use((err, req, res, _next) => {
  logger.error('서버 에러', { error: err.message, stack: err.stack });
  res.status(500).json({ error: '서버 내부 오류가 발생했습니다' });
});

process.on('unhandledRejection', (reason) => {
  logger.error('Unhandled Promise Rejection', { error: reason?.message || String(reason) });
});

process.on('uncaughtException', (err) => {
  logger.error('Uncaught Exception', { error: err.message, stack: err.stack });
});

initGemini();

app.listen(PORT, HOST, () => {
  logger.info(`두손푸드웨이 거래처원장 관리 서버 시작: http://${HOST}:${PORT}`);
});
