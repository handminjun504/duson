const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const logger = require('../utils/logger');

const BASE_URL = 'https://ai.serp.co.kr';
const USER_DATA_DIR = path.join(__dirname, '..', '..', '.browser-data');
const NET_LOG_PATH = path.join(__dirname, '..', '..', 'logs', 'network-capture.log');
const DEBUG_DIR = path.join(__dirname, '..', '..', 'public', 'debug');

function appendNetLog(line) {
  try {
    const dir = path.dirname(NET_LOG_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.appendFileSync(NET_LOG_PATH, `[${new Date().toISOString()}] ${line}\n`);
  } catch { /* ignore */ }
}

async function saveDebugScreenshot(target, label) {
  try {
    if (!fs.existsSync(DEBUG_DIR)) fs.mkdirSync(DEBUG_DIR, { recursive: true });
    const ts = Date.now();
    const filename = `${label}-${ts}.png`;
    const filepath = path.join(DEBUG_DIR, filename);
    const screenshotTarget = target.page ? target.page() : target;
    await screenshotTarget.screenshot({ path: filepath, fullPage: true });
    logger.info(`디버그 스크린샷 저장: /debug/${filename}`);
    return `/debug/${filename}`;
  } catch (err) {
    logger.warn(`스크린샷 저장 실패: ${err.message}`);
    return null;
  }
}

const MENU_ACT = {
  s1110: '/trgm_m002_01.act',
  s3120: '/fnsh_0004_01.act',
  s3130: '/rcpt_0003_01.act',
};

let browser = null;
let page = null;
let context = null;
let isLoggedIn = false;
let isBusy = false;
let loginRetried = false;
const progress = { percent: 0, message: '', step: '', total: 0, current: 0 };

function updateProgress(percent, message, extra = {}) {
  progress.percent = Math.min(100, Math.round(percent));
  progress.message = message;
  Object.assign(progress, extra);
}

function getProgress() {
  return { ...progress };
}

async function closeBrowser() {
  try {
    if (context) await context.close().catch(() => null);
  } catch { /* ignore */ }
  browser = null;
  page = null;
  context = null;
  isLoggedIn = false;
}

async function ensureBrowser() {
  if (browser && page) {
    try {
      await page.evaluate(() => true);
      return;
    } catch {
      await closeBrowser();
    }
  }

  const headless = process.env.HEADLESS !== 'false';
  context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless,
    slowMo: headless ? 50 : 150,
    viewport: { width: 1400, height: 900 },
    locale: 'ko-KR',
    ignoreHTTPSErrors: true,
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-web-security',
      '--ignore-certificate-errors',
      '--allow-running-insecure-content',
      '--disable-features=IsolateOrigins,site-per-process',
    ],
  });
  browser = context;
  const pages = context.pages();
  page = pages.length > 0 ? pages[0] : await context.newPage();
  page.setDefaultTimeout(30000);

  page.on('response', async (response) => {
    const url = response.url();
    if (!url.includes('.act')) return;
    const req = response.request();
    const method = req.method();
    const postData = req.postData() || '';
    const status = response.status();
    const logLine = `${method} ${url} (${status})`;
    appendNetLog(`[REQ] ${logLine}`);
    if (postData) appendNetLog(`[POST] ${postData.substring(0, 1000)}`);
  });

  // missinstall 리다이렉트를 네트워크 레벨에서 차단
  await page.route('**/wserp_0003_01.act**', (route) => {
    logger.info('[ROUTE] missinstall 리다이렉트 차단');
    route.abort('blockedbyclient');
  });

  isLoggedIn = false;
  logger.info('Playwright 브라우저 실행됨 (영구 컨텍스트)');
}

async function login() {
  if (isLoggedIn) return;

  const loginUrl = process.env.GYEONGLI_URL || `${BASE_URL}/wserp_0002_01.act`;
  const userId = (process.env.GYEONGLI_ID || '').replace(/^["']|["']$/g, '');
  const userPw = (process.env.GYEONGLI_PW || '').replace(/^["']|["']$/g, '');

  if (!userId || !userPw) {
    throw new Error('경리나라 로그인 정보가 .env에 설정되지 않았습니다 (GYEONGLI_ID, GYEONGLI_PW)');
  }

  logger.info(`경리나라 로그인 시도 중... (ID: ${userId.substring(0, 4)}***)`);

  const currentUrl = page.url();
  logger.info(`현재 페이지 URL: ${currentUrl}`);

  if (currentUrl.includes('serp.co.kr') && !currentUrl.includes('0002_01') && !currentUrl.includes('about:blank')) {
    const hasContent = await page.locator('.toolbar2, .co_name, .gnb, .lnb, #main_iframe').count();
    if (hasContent > 0) {
      isLoggedIn = true;
      logger.info('경리나라 이미 로그인되어 있음 (영구 세션)');
      return;
    }
  }

  // alert 팝업 캡처 (로그인 실패 메시지 등)
  page.on('dialog', async (dialog) => {
    logger.info(`[DIALOG] ${dialog.type()}: ${dialog.message()}`);
    await page.evaluate((msg) => { window.__lastAlert = msg; }, dialog.message());
    await dialog.dismiss().catch(() => null);
  });

  // missinstall 쿠키를 미리 설정하여 리다이렉트 방지
  await page.context().addCookies([
    { name: 'missinstall', value: 'Y', domain: new URL(BASE_URL).hostname, path: '/' },
  ]);

  await page.goto(loginUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(1500);

  let pageUrl = page.url();
  logger.info(`로그인 페이지 URL: ${pageUrl}`);

  if (pageUrl.includes('missinstall') || pageUrl.includes('0003_01')) {
    logger.warn('missinstall 페이지 감지, 우회 시도...');

    // 1) 페이지 내 건너뛰기/다음에 설치/확인 버튼 클릭 시도
    const skipClicked = await page.evaluate(() => {
      const all = document.querySelectorAll('a, button, input[type="button"], span');
      for (const el of all) {
        const txt = (el.textContent || el.value || '').trim();
        if (/건너뛰|다음에|나중에|확인|skip|continue|close/i.test(txt)) {
          el.click();
          return txt;
        }
      }
      const links = document.querySelectorAll('a[href]');
      for (const a of links) {
        if (a.href.includes('0002_01') || a.href.includes('login')) {
          a.click();
          return a.href;
        }
      }
      return null;
    });
    logger.info(`missinstall 우회 클릭: ${skipClicked}`);
    await page.waitForTimeout(3000);
    pageUrl = page.url();

    // 2) 아직 missinstall이면 route intercept로 강제 이동
    if (pageUrl.includes('missinstall') || pageUrl.includes('0003_01')) {
      logger.info('route intercept로 로그인 페이지 강제 이동');
      await page.route('**/wserp_0003_01**', route => route.abort());
      await page.goto(loginUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
      await page.waitForTimeout(3000);
      pageUrl = page.url();
      await page.unroute('**/wserp_0003_01**');
    }

    // 3) 아직 안 되면 JavaScript로 location 강제 변경
    if (pageUrl.includes('missinstall') || pageUrl.includes('0003_01')) {
      logger.info('JavaScript location으로 강제 이동');
      await page.evaluate((url) => { window.location.href = url; }, loginUrl);
      await page.waitForTimeout(5000);
      pageUrl = page.url();
    }

    logger.info(`missinstall 우회 후 URL: ${pageUrl}`);
  }

  if (!pageUrl.includes('0002_01') && !pageUrl.includes('missinstall') && !pageUrl.includes('0003_01')) {
    const hasContent = await page.locator('.toolbar2, .co_name, .gnb, .lnb, #main_iframe').count();
    if (hasContent > 0) {
      isLoggedIn = true;
      logger.info('경리나라 이미 로그인되어 있음 (리다이렉트)');
      return;
    }
  }

  // 보안 플러그인 오버레이/팝업 강제 제거
  await page.evaluate(() => {
    document.querySelectorAll('[role="dialog"], .modal, .popup, .layer_pop, .dim, .dimmed').forEach(d => {
      const btn = d.querySelector('button, .close, .btn_close, [class*="close"]');
      if (btn) btn.click();
      else d.remove();
    });
    // pointer-events를 가로채는 전체화면 오버레이 제거
    document.querySelectorAll('div').forEach(el => {
      const style = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      if (rect.width > window.innerWidth * 0.5 && rect.height > window.innerHeight * 0.5
          && style.position === 'fixed' && parseInt(style.zIndex || '0') > 100) {
        el.remove();
      }
    });
  }).catch(() => null);
  await page.waitForTimeout(500);

  const pwField = page.locator('input[type="password"]').first();
  const hasPwField = await pwField.count();
  logger.info(`비밀번호 필드 존재: ${hasPwField > 0}`);

  if (hasPwField === 0) {
    const debugInfo = await page.evaluate(() => ({
      url: window.location.href,
      title: document.title,
      body: document.body?.innerText?.substring(0, 500) || '',
      inputs: Array.from(document.querySelectorAll('input')).map(i => `${i.type}:${i.name||i.id}`).join(', '),
      links: Array.from(document.querySelectorAll('a[href]')).slice(0, 5).map(a => a.href).join(', '),
    }));
    logger.error('로그인 폼 미발견 디버그 정보', debugInfo);
    throw new Error(`로그인 폼을 찾을 수 없습니다. URL: ${debugInfo.url}`);
  }

  // ID 필드 찾기: name/id 속성 우선, 그 다음 type="text"
  const idField = page.locator('input[name="userId"], input[name="user_id"], input[name="loginId"], input[id="userId"], input[id="user_id"], input[id="loginId"]').first();
  const idFieldCount = await idField.count();
  const idSelector = idFieldCount > 0
    ? idField
    : page.locator('input[type="text"], input[type="email"]').first();

  await idSelector.click({ force: true }).catch(() => null);
  await idSelector.fill('');
  await idSelector.type(userId, { delay: 10 });
  logger.info('ID 입력 완료');

  await pwField.click({ force: true }).catch(() => null);
  await pwField.fill('');
  await pwField.type(userPw, { delay: 10 });
  logger.info('PW 입력 완료');

  await page.waitForTimeout(200);

  logger.info('로그인 버튼 클릭 시도...');
  // 1차: Playwright locator로 로그인 버튼 클릭
  const loginBtn = page.locator('button:has-text("로그인"), a:has-text("로그인"), input[value*="로그인"], button:has-text("LOGIN"), a:has-text("LOGIN")').first();
  const loginBtnCount = await loginBtn.count();
  if (loginBtnCount > 0) {
    await loginBtn.click({ force: true });
    logger.info('로그인 버튼 클릭 (Playwright locator)');
  } else {
    // 2차: evaluate 폴백
    const loginClicked = await page.evaluate(() => {
      const btns = document.querySelectorAll('button, a, input[type="submit"], input[type="button"]');
      for (const btn of btns) {
        const txt = (btn.textContent || btn.value || '').trim();
        if (txt.includes('로그인') || txt.includes('LOGIN') || txt.includes('Log in')) {
          btn.click();
          return txt;
        }
      }
      const form = document.querySelector('form');
      if (form) { form.submit(); return 'form.submit'; }
      return null;
    });
    logger.info(`로그인 클릭 (evaluate): ${loginClicked}`);
  }

  // 로그인 후 페이지 전환 대기
  await Promise.race([
    page.waitForNavigation({ waitUntil: 'domcontentloaded', timeout: 10000 }).catch(() => null),
    page.waitForTimeout(5000),
  ]);

  const afterUrl = page.url();
  logger.info(`로그인 후 URL: ${afterUrl}`);

  const alertText = await page.evaluate(() => window.__lastAlert || null).catch(() => null);
  if (alertText) logger.warn(`alert 메시지: ${alertText}`);

  const hasToolbar = await page.locator('.toolbar2, .co_name, .gnb, .lnb, #main_iframe').count();
  const hasFrame = page.frames().length > 1;
  const hasPwStill = await page.locator('input[type="password"]').count();

  logger.info(`로그인 판별: toolbar=${hasToolbar} frames=${page.frames().length} pwField=${hasPwStill} url=${afterUrl}`);

  if (hasToolbar > 0 || hasFrame) {
    isLoggedIn = true;
    logger.info('경리나라 로그인 성공');
  } else if (!afterUrl.includes('0002_01') && afterUrl.includes('serp.co.kr')) {
    isLoggedIn = true;
    logger.info('경리나라 로그인 성공 (URL 기반 확인)');
  } else if (hasPwStill === 0 && afterUrl.includes('serp.co.kr')) {
    isLoggedIn = true;
    logger.info('경리나라 로그인 성공 (비밀번호 필드 사라짐)');
  } else {
    const bodyText = await page.evaluate(() => document.body?.innerText?.substring(0, 500) || '');
    logger.error(`로그인 실패 디버그`, { afterUrl, bodyText, alertText, frames: page.frames().length });

    if (!loginRetried) {
      loginRetried = true;
      logger.info('브라우저 재시작 후 로그인 재시도...');
      await closeBrowser();
      try {
        const browserDataDir = USER_DATA_DIR;
        if (fs.existsSync(browserDataDir)) {
          fs.rmSync(browserDataDir, { recursive: true, force: true });
          logger.info('.browser-data 초기화 완료');
        }
      } catch (e) { logger.warn(`browser-data 삭제 실패: ${e.message}`); }
      await ensureBrowser();
      return login();
    }
    loginRetried = false;
    throw new Error('경리나라 로그인 실패. 아이디/비밀번호를 확인하세요.');
  }
  loginRetried = false;
}

async function navigateToMenu(menuId) {
  const actFile = MENU_ACT[menuId];
  if (!actFile) throw new Error(`알 수 없는 메뉴 ID: ${menuId}`);

  logger.info(`메뉴 이동: ${menuId} → ${actFile}`);

  const currentUrl = page.url();
  if (!currentUrl.includes('serp.co.kr') || currentUrl.includes('0002_01')) {
    await page.goto(`${BASE_URL}/?MENU_ID=${menuId}`, {
      waitUntil: 'domcontentloaded',
      timeout: 30000,
    });
    await page.waitForTimeout(1000);
  }

  let frame = page.frame({ name: 'main_iframe' });

  if (!frame) {
    const frames = page.frames();
    frame = frames.find(f => f !== page.mainFrame() && f.url().includes('serp.co.kr'));
  }

  if (frame) {
    const targetUrl = `${BASE_URL}${actFile}`;
    await frame.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(800);
    logger.info(`iframe URL: ${frame.url()}`);
    return frame;
  }

  logger.warn('iframe을 찾을 수 없어 직접 페이지 이동');
  await page.goto(`${BASE_URL}${actFile}`, {
    waitUntil: 'domcontentloaded',
    timeout: 30000,
  });
  await page.waitForTimeout(800);
  return page.mainFrame();
}

async function scrollToLoadAll(target, label = '') {
  let prevRowCount = 0;
  for (let attempt = 0; attempt < 5; attempt++) {
    const info = await target.evaluate(() => {
      const allEls = document.querySelectorAll('div, section');
      for (const el of allEls) {
        const style = getComputedStyle(el);
        if ((style.overflowY === 'scroll' || style.overflowY === 'auto') && el.scrollHeight > el.clientHeight + 10) {
          el.scrollTop = el.scrollHeight;
        }
      }
      const docEl = document.scrollingElement || document.documentElement;
      docEl.scrollTop = docEl.scrollHeight;
      window.scrollTo(0, document.body.scrollHeight);
      return document.querySelectorAll('tr').length;
    }).catch(() => 0);

    if (attempt > 0 && info === prevRowCount) break;
    prevRowCount = info;
    await page.waitForTimeout(500);
  }
}

async function setDateRange(frame, startDate, endDate) {
  if (!startDate && !endDate) return;

  const startDash = (startDate || '').replace(/\//g, '-');
  const endDash = (endDate || '').replace(/\//g, '-');
  const startSlash = (startDate || '').replace(/-/g, '/');
  const endSlash = (endDate || '').replace(/-/g, '/');

  logger.info(`날짜 범위 설정: ${startDash} ~ ${endDash}`);

  const result = await frame.evaluate(({ sDash, eDash, sSlash, eSlash }) => {
    const log = [];

    function setDateField(el, dashVal, slashVal) {
      const usesSlash = el.value.includes('/');
      const val = usesSlash ? slashVal : dashVal;
      const before = el.value;

      el.value = val;
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')?.set;
      if (setter) setter.call(el, val);

      // jQuery datepicker 내부 상태 동기화
      try {
        if (typeof $ !== 'undefined' && $(el).hasClass('hasDatepicker')) {
          const parts = dashVal.split('-');
          const dateObj = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
          $(el).datepicker('setDate', dateObj);
          log.push(`  datepicker setDate: ${dateObj.toISOString().substring(0,10)}`);
        }
      } catch (e) { log.push(`  datepicker err: ${e.message}`); }

      // jQuery val() 도 시도
      try {
        if (typeof $ !== 'undefined') { $(el).val(val).trigger('change'); }
      } catch {}

      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));

      return `"${before}" -> "${el.value}"`;
    }

    const startIds = ['txtSrcStartDt', 'txtDtlSrcStartDt'];
    const endIds = ['txtSrcEndDt', 'txtDtlSrcEndDt'];
    let startSet = false, endSet = false;

    for (const id of startIds) {
      const el = document.getElementById(id);
      if (el) {
        const r = setDateField(el, sDash, sSlash);
        log.push(`start ${id}: ${r}`);
        startSet = true;
        break;
      }
    }
    for (const id of endIds) {
      const el = document.getElementById(id);
      if (el) {
        const r = setDateField(el, eDash, eSlash);
        log.push(`end ${id}: ${r}`);
        endSet = true;
        break;
      }
    }

    if (!startSet || !endSet) {
      const allInputs = document.querySelectorAll('input[type="text"], input:not([type])');
      for (const el of allInputs) {
        if (startSet && endSet) break;
        const v = el.value || '';
        const id = el.id || '';
        const name = el.name || '';
        const isDate = /\d{4}[-/]\d{2}[-/]\d{2}/.test(v) || el.classList.contains('hasDatepicker');
        if (!isDate || el.offsetParent === null) continue;
        if (!startSet && (/start|str|from/i.test(id + name) || v <= sDash)) {
          const r = setDateField(el, sDash, sSlash);
          log.push(`start fallback ${id||name}: ${r}`);
          startSet = true;
        } else if (!endSet) {
          const r = setDateField(el, eDash, eSlash);
          log.push(`end fallback ${id||name}: ${r}`);
          endSet = true;
        }
      }
    }

    // 설정 후 최종 확인
    const verify = [];
    for (const id of [...startIds, ...endIds]) {
      const el = document.getElementById(id);
      if (el) verify.push(`${id}="${el.value}"`);
    }
    log.push(`verify: ${verify.join(', ')}`);
    log.push(`set: start=${startSet}, end=${endSet}, jQuery=${typeof $ !== 'undefined'}`);

    try { if (typeof fn_search === 'function') { fn_search(); log.push('fn_search()'); return log; } } catch (e) { log.push('fn_search err: ' + e.message); }
    try { if (typeof fn_Search === 'function') { fn_Search(); log.push('fn_Search()'); return log; } } catch (e) { log.push('fn_Search err: ' + e.message); }
    try { if (typeof doSearch === 'function') { doSearch(); log.push('doSearch()'); return log; } } catch (e) { log.push('doSearch err: ' + e.message); }

    const btns = document.querySelectorAll('[onclick], button, a, input[type="button"], img, span');
    for (const btn of btns) {
      const oc = btn.getAttribute('onclick') || '';
      if (/fn_search|fn_Search|doSearch/.test(oc)) { btn.click(); log.push('onclick: ' + oc.substring(0, 50)); return log; }
      const txt = (btn.textContent?.trim() || btn.value || btn.alt || btn.title || '');
      if (txt === '조회' || txt === '검색') { btn.click(); log.push('click: ' + txt); return log; }
    }
    log.push('search button not found');
    return log;
  }, { sDash: startDash, eDash: endDash, sSlash: startSlash, eSlash: endSlash });

  logger.info(`날짜+조회 결과: ${result.join(' | ')}`);

  await page.waitForTimeout(1500);

  const afterCheck = await frame.evaluate(({ sDash, eDash, sSlash, eSlash }) => {
    const fields = ['txtSrcStartDt', 'txtSrcEndDt', 'txtDtlSrcStartDt', 'txtDtlSrcEndDt'];
    const vals = fields.map(id => { const el = document.getElementById(id); return el ? { id, value: el.value } : null; }).filter(Boolean);
    const links = document.querySelectorAll('[onclick*="PopupCall"], [onclick*="fn_PopupCall"]');

    const startField = vals.find(v => /start|str/i.test(v.id));
    const endField = vals.find(v => /end/i.test(v.id));
    const startOk = startField && (startField.value === sDash || startField.value === sSlash);
    const endOk = endField && (endField.value === eDash || endField.value === eSlash);

    return {
      fields: vals.map(v => `${v.id}=${v.value}`).join(', '),
      clients: links.length,
      startOk,
      endOk,
    };
  }, { sDash: startDash, eDash: endDash, sSlash: startSlash, eSlash: endSlash });

  logger.info(`조회 후: ${afterCheck.fields}, clients:${afterCheck.clients}, dateOk:start=${afterCheck.startOk},end=${afterCheck.endOk}`);

  if (!afterCheck.startOk || !afterCheck.endOk) {
    logger.warn('날짜가 리셋됨! Playwright type()으로 재설정 후 재조회...');

    const startSelectors = ['#txtSrcStartDt', '#txtDtlSrcStartDt'];
    const endSelectors = ['#txtSrcEndDt', '#txtDtlSrcEndDt'];

    for (const sel of startSelectors) {
      const el = frame.locator(sel).first();
      if (await el.count() > 0) {
        const curVal = await el.inputValue().catch(() => '');
        const usesSlash = curVal.includes('/');
        await el.click({ force: true }).catch(() => null);
        await el.fill('');
        await el.type(usesSlash ? startSlash : startDash, { delay: 30 });
        await el.dispatchEvent('change');
        logger.info(`시작일 재설정(${sel}): ${usesSlash ? startSlash : startDash}`);
        break;
      }
    }
    for (const sel of endSelectors) {
      const el = frame.locator(sel).first();
      if (await el.count() > 0) {
        const curVal = await el.inputValue().catch(() => '');
        const usesSlash = curVal.includes('/');
        await el.click({ force: true }).catch(() => null);
        await el.fill('');
        await el.type(usesSlash ? endSlash : endDash, { delay: 30 });
        await el.dispatchEvent('change');
        logger.info(`종료일 재설정(${sel}): ${usesSlash ? endSlash : endDash}`);
        break;
      }
    }

    await page.waitForTimeout(300);

    await frame.evaluate(() => {
      try { if (typeof fn_search === 'function') { fn_search(); return; } } catch {}
      try { if (typeof fn_Search === 'function') { fn_Search(); return; } } catch {}
      const btns = document.querySelectorAll('[onclick], button, a');
      for (const btn of btns) {
        const txt = (btn.textContent?.trim() || btn.value || '');
        if (txt === '조회' || txt === '검색') { btn.click(); return; }
      }
    });
    logger.info('재조회 실행');

    await page.waitForTimeout(2000);

    const recheck = await frame.evaluate(() => {
      const fields = ['txtSrcStartDt', 'txtSrcEndDt', 'txtDtlSrcStartDt', 'txtDtlSrcEndDt'];
      const vals = fields.map(id => { const el = document.getElementById(id); return el ? `${id}=${el.value}` : null; }).filter(Boolean);
      const links = document.querySelectorAll('[onclick*="PopupCall"], [onclick*="fn_PopupCall"]');
      return vals.join(', ') + `, clients:${links.length}`;
    });
    logger.info(`재조회 후: ${recheck}`);
  }
}

async function collectSalesData(options = {}) {
  if (isBusy) throw new Error('이미 수집 중입니다. 잠시 후 다시 시도하세요.');
  isBusy = true;
  updateProgress(0, '준비 중...');

  const { startDate, endDate } = options;

  try {
    updateProgress(5, '브라우저 시작...');
    await ensureBrowser();
    updateProgress(10, '로그인 중...');
    await login();

    updateProgress(12, '로그인 상태 확인...');
    logger.info(`로그인 후 페이지 URL: ${page.url()}, frames: ${page.frames().length}`);
    await saveDebugScreenshot(page, 'after-login');

    updateProgress(15, '거래명세표 페이지 이동...');
    logger.info(`매출(거래명세표) 페이지 이동... ${startDate ? `(${startDate} ~ ${endDate})` : '(기본 기간)'}`);
    const frame = await navigateToMenu('s1110');
    logger.info(`거래명세표 프레임 URL: ${frame.url()}`);
    await saveDebugScreenshot(page, 'after-navigate');

    updateProgress(20, '날짜 범위 설정...');
    await setDateRange(frame, startDate, endDate);

    await page.waitForTimeout(2000);

    const frameDebug = await frame.evaluate(() => ({
      url: window.location.href,
      title: document.title,
      dateFields: (() => {
        const ids = ['txtSrcStartDt', 'txtSrcEndDt', 'txtDtlSrcStartDt', 'txtDtlSrcEndDt'];
        return ids.map(id => { const el = document.getElementById(id); return el ? `${id}=${el.value}` : null; }).filter(Boolean);
      })(),
      allOnclicks: Array.from(document.querySelectorAll('[onclick]')).slice(0, 20).map(el => ({
        tag: el.tagName, text: el.textContent.trim().substring(0, 30), onclick: (el.getAttribute('onclick') || '').substring(0, 80),
      })),
      tableCount: document.querySelectorAll('table').length,
      trCount: document.querySelectorAll('tr').length,
      linkCount: document.querySelectorAll('a').length,
    }));
    logger.info(`프레임 상태: url=${frameDebug.url}, tables=${frameDebug.tableCount}, rows=${frameDebug.trCount}, dates=[${frameDebug.dateFields.join(', ')}]`);
    if (frameDebug.allOnclicks.length > 0) {
      logger.info(`onclick 요소 (${frameDebug.allOnclicks.length}개): ${JSON.stringify(frameDebug.allOnclicks.slice(0, 5))}`);
    }

    let clientLinks = await frame.evaluate(() => {
      const links = document.querySelectorAll('[onclick*="PopupCall"], [onclick*="fn_PopupCall"]');
      return Array.from(links)
        .map(el => ({
          name: el.textContent.trim(),
          onclick: el.getAttribute('onclick') || '',
        }))
        .filter(item => item.name.length > 1 && !/^[\d,.]+$/.test(item.name));
    });

    if (clientLinks.length === 0) {
      logger.warn('PopupCall 링크 없음, 대체 셀렉터 시도...');
      clientLinks = await frame.evaluate(() => {
        const results = [];
        const allLinks = document.querySelectorAll('a[onclick], td[onclick], span[onclick]');
        for (const el of allLinks) {
          const oc = el.getAttribute('onclick') || '';
          const name = el.textContent.trim();
          if (name.length > 1 && !/^[\d,.]+$/.test(name) && !['조회','검색','닫기','확인','삭제','수정'].includes(name)) {
            if (oc.includes('Popup') || oc.includes('popup') || oc.includes('detail') || oc.includes('Detail') || oc.includes('view') || oc.includes('View')) {
              results.push({ name, onclick: oc });
            }
          }
        }
        return results;
      });
      if (clientLinks.length > 0) {
        logger.info(`대체 셀렉터로 ${clientLinks.length}개 거래처 발견`);
      }
    }

    logger.info(`${clientLinks.length}개 거래처 발견: ${clientLinks.map(c => c.name).join(', ')}`);
    updateProgress(25, `${clientLinks.length}개 거래처 발견`, { total: clientLinks.length, current: 0 });

    if (clientLinks.length === 0) {
      const allText = await frame.evaluate(() => document.body?.innerText || '');
      logger.warn(`거래처 링크 없음. 페이지 텍스트 (${allText.length}자): ${allText.substring(0, 500)}`);
      await saveDebugScreenshot(page, 'no-clients');
      updateProgress(100, '거래처 없음 (디버그 스크린샷 저장됨)');
      return [];
    }

    const salesData = [];

    for (let idx = 0; idx < clientLinks.length; idx++) {
      const client = clientLinks[idx];
      const pct = 25 + Math.round((idx / clientLinks.length) * 60);
      updateProgress(pct, `거래처 수집: ${client.name} (${idx + 1}/${clientLinks.length})`, { current: idx + 1 });

      try {
        await page.evaluate(() => true);
      } catch {
        logger.warn('브라우저가 닫혔습니다. 수집을 중단합니다.');
        break;
      }

      try {
        logger.info(`거래처 수집: ${client.name}`);
        const clientDetail = await fetchClientData(frame, client);
        if (clientDetail) {
          salesData.push(clientDetail);
          logger.info(`  → ${client.name}: ${clientDetail.items.length}개 품목, 합계 ${clientDetail.totalAmount}`);
        }
      } catch (err) {
        if (err.message.includes('closed') || err.message.includes('destroyed')) {
          logger.warn('브라우저가 닫혔습니다. 수집을 중단합니다.');
          break;
        }
        logger.warn(`거래처 "${client.name}" 실패: ${err.message}`);
        salesData.push({ client: client.name, items: [], error: err.message });
      }
    }

    const totalItems = salesData.reduce((s, c) => s + (c.items?.length || 0), 0);

    const allDates = salesData.flatMap(c =>
      (c.items || []).map(i => i.deliveryDate || i.date || '').filter(Boolean)
    ).sort();
    if (allDates.length > 0) {
      logger.info(`수집된 날짜 범위: ${allDates[0]} ~ ${allDates[allDates.length - 1]} (${allDates.length}건)`);
    }

    updateProgress(85, `매출 수집 완료: ${salesData.length}개 거래처, ${totalItems}건`);
    logger.info(`매출 수집 완료: ${salesData.length}개 거래처, ${totalItems}건`);
    return salesData;

  } catch (err) {
    updateProgress(0, `오류: ${err.message}`);
    logger.error('경리나라 매출 데이터 수집 실패', { error: err.message });
    throw err;
  } finally {
    isBusy = false;
  }
}

const PARSE_POPUP_JS = `(function(doc) {
  var result = {
    client: '', date: '', items: [],
    prevBalance: 0, totalSupply: 0, totalVat: 0,
    totalAmount: 0, depositAmount: 0, outstandingBalance: 0,
  };
  var tables = doc.querySelectorAll('table');
  for (var ti = 0; ti < tables.length; ti++) {
    var rows = tables[ti].querySelectorAll('tr');
    for (var ri = 0; ri < rows.length; ri++) {
      var cells = Array.from(rows[ri].querySelectorAll('td, th'));
      var texts = cells.map(function(c) { return c.textContent.trim(); });
      var clientIdx = texts.findIndex(function(t) { return t.includes('거 래 처') || t.includes('거래처'); });
      if (clientIdx >= 0 && clientIdx < texts.length - 1) result.client = texts[clientIdx + 1] || '';
      var dateIdx = texts.findIndex(function(t) { return (t.includes('일') && t.includes('자')) && !t.includes('품'); });
      if (dateIdx >= 0 && dateIdx < texts.length - 1) result.date = texts[dateIdx + 1] || '';
    }
  }
  var itemTable = null;
  for (var ti2 = 0; ti2 < tables.length; ti2++) {
    var txt = tables[ti2].textContent || '';
    if ((txt.includes('품목') || txt.includes('품 목')) &&
        (txt.includes('수량') || txt.includes('수 량')) &&
        (txt.includes('단가') || txt.includes('단 가'))) { itemTable = tables[ti2]; break; }
  }
  if (itemTable) {
    var irows = itemTable.querySelectorAll('tr');
    var colMap = {};
    var headerFound = false;
    var colPatterns = {
      date: ['월일','일자','월 일','일 자'], product: ['품목','품 목','상품명'],
      spec: ['규격','규 격'], unit: ['단위','단 위'], qty: ['수량','수 량'],
      unitPrice: ['단가','단 가'], supply: ['공급가액','공급가','공 급 가 액'],
      vat: ['세액','세 액','부가세','부가'], note: ['비고','비 고'],
    };
    for (var iri = 0; iri < irows.length; iri++) {
      var icells = Array.from(irows[iri].querySelectorAll('td, th'));
      var itexts = icells.map(function(c) { return c.textContent.trim().replace(/\\s+/g,''); });
      if (itexts.some(function(t){return t.includes('품목');}) && itexts.some(function(t){return t.includes('수량');})) {
        for (var ci = 0; ci < itexts.length; ci++) {
          var ct = itexts[ci];
          var keys = Object.keys(colPatterns);
          for (var ki = 0; ki < keys.length; ki++) {
            var ps = colPatterns[keys[ki]];
            if (ps.some(function(p){return ct.replace(/\\s/g,'').includes(p.replace(/\\s/g,''));})) colMap[keys[ki]] = ci;
          }
        }
        headerFound = true; continue;
      }
      if (!headerFound) continue;
      if (itexts.length < 4) continue;
      if (itexts.every(function(t){return !t || t==='';})) continue;
      if (itexts.some(function(t){return t.includes('합계') || t.includes('소계');})) continue;
      var get = function(key) { return (colMap[key] !== undefined ? itexts[colMap[key]] : '') || ''; };
      var pn = function(s) { return parseFloat((s||'').replace(/[,\\s]/g,'')) || 0; };
      var product = get('product');
      var dateCell = get('date');
      if (product && product.length > 0 && product !== dateCell) {
        result.items.push({
          deliveryDate: dateCell, product: product,
          spec: get('spec'), qty: pn(get('qty')),
          unitPrice: pn(get('unitPrice')), supplyAmount: pn(get('supply')),
          vat: pn(get('vat')), total: pn(get('supply')) + pn(get('vat')),
          note: get('note'),
        });
      }
    }
    result._colMap = colMap;
  }
  for (var ti3 = 0; ti3 < tables.length; ti3++) {
    var srows = tables[ti3].querySelectorAll('tr');
    for (var sri = 0; sri < srows.length; sri++) {
      var text = srows[sri].textContent || '';
      var pn2 = function(s) { return parseFloat((s||'').replace(/[,\\s]/g,'')) || 0; };
      if (text.includes('전미수잔액')) {
        var nums = text.match(/[\\d,]+/g);
        if (nums) result.prevBalance = pn2(nums[nums.length-1]);
      }
      if (text.includes('합계') && !text.includes('총합계') && !text.includes('전미수')) {
        var vals = Array.from(srows[sri].querySelectorAll('td')).map(function(c){return c.textContent.trim();}).filter(function(t){return /^[\\d,]+$/.test(t);});
        if (vals.length >= 2) { result.totalSupply = pn2(vals[0]); result.totalVat = pn2(vals[1]); }
      }
      if (text.includes('총합계') || text.includes('합계금액')) {
        var m = text.match(/([\\d,]+)\\s*$/);
        if (m) result.totalAmount = pn2(m[1]);
      }
      if (text.includes('입금액')) {
        var dvals = Array.from(srows[sri].querySelectorAll('td')).map(function(c){return c.textContent.trim().replace(/[,\\s]/g,'');}).filter(function(v){return /^\\d+$/.test(v) && parseInt(v)>0;});
        if (dvals.length > 0) result.depositAmount = parseFloat(dvals[0]) || 0;
      }
      if (text.includes('총미수잔액') || text.includes('미수잔액')) {
        var ovals = Array.from(srows[sri].querySelectorAll('td')).map(function(c){return c.textContent.trim().replace(/[,\\s]/g,'');}).filter(function(v){return /^\\d+$/.test(v) && parseInt(v)>0;});
        if (ovals.length > 0) result.outstandingBalance = parseFloat(ovals[0]) || 0;
      }
    }
  }
  result.totalAmount = result.totalAmount || result.items.reduce(function(s,i){return s+i.total;},0);
  return result;
})`;

async function fetchClientData(frame, client) {
  const capturedUrl = await frame.evaluate((onclick) => {
    return new Promise((resolve) => {
      const origOpen = window.open;
      window.open = function(url) {
        window.open = origOpen;
        resolve(url);
        return null;
      };
      try {
        const fn = new Function(onclick);
        fn();
      } catch {
        window.open = origOpen;
        resolve(null);
      }
      setTimeout(() => { window.open = origOpen; resolve(null); }, 500);
    });
  }, client.onclick);

  if (!capturedUrl) {
    logger.warn(`${client.name}: URL 캡처 실패, 팝업 fallback`);
    return openClientPopupFallback(frame, client);
  }

  const fullUrl = capturedUrl.startsWith('http')
    ? capturedUrl
    : `${BASE_URL}${capturedUrl.startsWith('/') ? '' : '/'}${capturedUrl}`;
  logger.info(`${client.name}: fetch URL = ${fullUrl}`);
  appendNetLog(`[FETCH] ${client.name}: ${fullUrl}`);

  const data = await frame.evaluate(async ({ url, parseCode }) => {
    try {
      const resp = await fetch(url, { credentials: 'include' });
      if (!resp.ok) return { _fetchError: 'HTTP ' + resp.status };
      const html = await resp.text();
      if (!html || html.length < 100) return { _fetchError: 'empty response' };
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const parseFn = eval(parseCode);
      return parseFn(doc);
    } catch (e) {
      return { _fetchError: e.message };
    }
  }, { url: fullUrl, parseCode: PARSE_POPUP_JS });

  if (data && data._fetchError) {
    logger.warn(`${client.name}: fetch 실패 (${data._fetchError}), 팝업 fallback`);
    return openClientPopupFallback(frame, client);
  }

  if (!data || (!data.items?.length && !data.totalAmount)) {
    logger.warn(`${client.name}: fetch 데이터 부족, 팝업 fallback`);
    return openClientPopupFallback(frame, client);
  }

  if (!data.client) data.client = client.name;
  logger.info(`${client.name}: fetch 성공 (${data.items?.length || 0}건)`);
  return data;
}

async function openClientPopupFallback(frame, client) {
  const popupPromise = context.waitForEvent('page', { timeout: 8000 });

  await frame.evaluate((onclick) => {
    const fn = new Function(onclick);
    fn();
  }, client.onclick);

  let popup;
  try {
    popup = await popupPromise;
  } catch {
    throw new Error('팝업이 열리지 않음');
  }

  await popup.waitForLoadState('networkidle').catch(() => popup.waitForLoadState('domcontentloaded'));
  await popup.waitForTimeout(500);

  await scrollToLoadAll(popup, `팝업(${client.name})`);

  const data = await popup.evaluate((parseCode) => {
    const parseFn = eval(parseCode);
    return parseFn(document);
  }, PARSE_POPUP_JS);

  await popup.close().catch(() => null);
  await page.waitForTimeout(100);

  if (!data.client) data.client = client.name;
  logger.info(`${client.name}: 팝업 fallback 성공 (${data.items?.length || 0}건)`);
  return data;
}

async function collectDepositData(options = {}) {
  if (isBusy) throw new Error('이미 수집 중입니다. 잠시 후 다시 시도하세요.');
  isBusy = true;

  const { startDate, endDate } = options;

  try {
    updateProgress(88, '입금내역 수집 시작...');
    await ensureBrowser();
    await login();

    updateProgress(90, '입출내역조회 페이지 이동...');
    logger.info(`입출내역조회 페이지 이동... ${startDate ? `(${startDate} ~ ${endDate})` : '(기본 기간)'}`);
    const frame = await navigateToMenu('s3120');

    if (startDate || endDate) {
      await setDateRange(frame, startDate, endDate);
    } else {
      await page.waitForTimeout(1000);
    }

    const deposits = await frame.evaluate(() => {
      const results = [];
      const tables = document.querySelectorAll('table');

      for (const table of tables) {
        const rows = table.querySelectorAll('tr');
        let isDepositTable = false;

        for (const row of rows) {
          const text = row.textContent;
          if (text.includes('적요') || text.includes('입금') || text.includes('거래일')) {
            isDepositTable = true;
            continue;
          }

          if (!isDepositTable) continue;

          const cells = Array.from(row.querySelectorAll('td'));
          if (cells.length < 3) continue;

          const texts = cells.map(c => c.textContent.trim());
          if (texts.every(t => !t)) continue;

          const dateCell = texts.find(t => /\d{4}[-./]\d{2}[-./]\d{2}/.test(t)) ||
                           texts.find(t => /\d{2}[-./]\d{2}/.test(t) && t.length <= 10);
          const memoCell = texts.find(t =>
            t.length > 1 && !/^[\d,.\-/\s]+$/.test(t) &&
            !t.includes('합계') && !t.includes('잔액')
          );
          const amountCells = texts.filter(t => /^[\d,]+$/.test(t.replace(/\s/g, '')));

          if (amountCells.length > 0) {
            const amount = parseFloat(amountCells[0].replace(/[,\s]/g, '')) || 0;
            if (amount > 0) {
              results.push({
                date: dateCell || '',
                client: memoCell || '',
                amount,
                allCells: texts,
              });
            }
          }
        }
      }

      return results;
    });

    logger.info(`입금내역: ${deposits.length}건 수집`);

    if (deposits.length === 0) {
      await saveDebugScreenshot(page, 'no-deposits');
      logger.info('입금내역 수납확인 페이지로 대체 시도...');
      const confirmFrame = await navigateToMenu('s3130');
      if (startDate || endDate) {
        await setDateRange(confirmFrame, startDate, endDate);
      } else {
        await page.waitForTimeout(2000);
      }

      const confirmDeposits = await confirmFrame.evaluate(() => {
        const results = [];
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
          const rows = table.querySelectorAll('tr');
          for (const row of rows) {
            const cells = Array.from(row.querySelectorAll('td'));
            if (cells.length < 3) continue;
            const texts = cells.map(c => c.textContent.trim());
            if (texts.every(t => !t)) continue;

            const hasDate = texts.some(t => /\d{2}[-./]\d{2}/.test(t));
            const hasAmount = texts.some(t => /^[\d,]+$/.test(t.replace(/\s/g, '')));

            if (hasDate || hasAmount) {
              const dateCell = texts.find(t => /\d{4}[-./]\d{2}[-./]\d{2}/.test(t)) ||
                               texts.find(t => /\d{2}[-./]\d{2}/.test(t));
              const memoCell = texts.find(t =>
                t.length > 1 && !/^[\d,.\-/\s]+$/.test(t) &&
                !t.includes('합계')
              );
              const amountCells = texts.filter(t => /^[\d,]+$/.test(t.replace(/\s/g, '')));
              if (amountCells.length > 0) {
                const amount = parseFloat(amountCells[0].replace(/[,\s]/g, '')) || 0;
                if (amount > 0) {
                  results.push({
                    date: dateCell || '',
                    client: memoCell || '',
                    amount,
                    allCells: texts,
                  });
                }
              }
            }
          }
        }
        return results;
      });

      deposits.push(...confirmDeposits);
      logger.info(`입금내역 수납확인: ${confirmDeposits.length}건 추가 수집`);
    }

    for (const dep of deposits) {
      parseDepositDetail(dep);
    }

    updateProgress(100, `수집 완료: 입금 ${deposits.length}건`);
    logger.info(`입금 수집 완료: 총 ${deposits.length}건`);
    return deposits;

  } catch (err) {
    logger.error('경리나라 입금 데이터 수집 실패', { error: err.message });
    throw err;
  } finally {
    isBusy = false;
  }
}

function isTimeStr(s) { return /^\d{1,2}:\d{2}(:\d{2})?$/.test(s); }
function isAccountStr(s) { return /^\d{2,}-\d+-\d+/.test(s); }
function isDateStr(s) { return /\d{4}[-./]\d{2}[-./]\d{2}/.test(s) || /^\d{2}[-./]\d{2}$/.test(s); }
function isNumericStr(s) { return /^[\d,]+$/.test((s || '').replace(/\s/g, '')); }
const BANK_NAMES = ['하나은행','하나','국민은행','국민','신한은행','신한','우리은행','우리','기업은행','기업','농협','SC은행','씨티은행','카카오뱅크','토스뱅크','케이뱅크','새마을금고','수협','신협','우체국','대구은행','부산은행','경남은행','광주은행','전북은행','제주은행'];
function isBankStr(s) {
  return BANK_NAMES.some(b => s.includes(b)) || s.endsWith('은행') || s.endsWith('금고');
}
const SKIP_WORDS = ['수정', '삭제', '수정삭제', '수정 삭제', '합계', '잔액'];
function isSkipStr(s) { return SKIP_WORDS.includes(s.replace(/\s/g, '')); }

function parseDepositDetail(dep) {
  const cells = dep.allCells || [];
  let time = '';
  let bank = '';
  let account = '';
  let clientName = '';
  let found = false;

  for (const cell of cells) {
    if (cell.includes('/')) {
      const parts = cell.split('/').map(p => p.trim()).filter(Boolean);
      if (parts.length >= 3 && (parts.some(isTimeStr) || parts.some(isAccountStr))) {
        for (const p of parts) {
          if (isSkipStr(p)) continue;
          if (isTimeStr(p)) { time = p; }
          else if (isAccountStr(p)) { account = p; }
          else if (isBankStr(p)) { bank = p; }
          else if (p.length > 1 && !isNumericStr(p) && !isDateStr(p)) { clientName = p; }
        }
        found = true;
        break;
      }
    }
  }

  if (!found) {
    for (const cell of cells) {
      if (!cell || cell.length < 2) continue;
      if (isSkipStr(cell)) continue;
      if (isDateStr(cell) || isNumericStr(cell)) continue;
      if (cell === dep.date) continue;

      if (isTimeStr(cell)) { time = cell; }
      else if (isAccountStr(cell)) { account = cell; }
      else if (isBankStr(cell)) { bank = cell; }
      else if (!clientName && cell.length > 1) { clientName = cell; }
    }
  }

  if (clientName) dep.client = clientName;
  dep.time = time;
  dep.bank = bank;
  dep.account = account;
  dep.detail = [time, bank, account].filter(Boolean).join(' / ');

  logger.info(`입금 파싱: client="${dep.client}", bank="${bank}", account="${account}", time="${time}", cells=[${cells.join(' | ')}]`);
}

async function closeBrowser() {
  if (context) {
    await context.close().catch(() => null);
    browser = null;
    page = null;
    context = null;
    isLoggedIn = false;
    isBusy = false;
    logger.info('Playwright 브라우저 종료');
  }
}

module.exports = {
  collectSalesData,
  collectDepositData,
  closeBrowser,
  getProgress,
};
