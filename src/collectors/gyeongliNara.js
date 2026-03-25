const { chromium } = require('playwright');
const path = require('path');
const logger = require('../utils/logger');

const BASE_URL = 'https://ai.serp.co.kr';
const USER_DATA_DIR = path.join(__dirname, '..', '..', '.browser-data');

const MENU_ACT = {
  s1110: '/trgm_m002_01.act',
  s1310: '/trco_m012_01_v6.act?MENU=SALE',
  s3120: '/fnsh_0004_01.act',
  s3130: '/rcpt_0003_01.act',
};

let browser = null;
let page = null;
let context = null;
let isLoggedIn = false;
let isBusy = false;
let cachedSales = null;
let cachedDeposits = null;

async function ensureBrowser() {
  if (browser && page) {
    try {
      await page.evaluate(() => true);
      return;
    } catch {
      browser = null;
      page = null;
      context = null;
      isLoggedIn = false;
    }
  }

  const headless = process.env.HEADLESS !== 'false';
  context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless,
    slowMo: 150,
    viewport: { width: 1400, height: 900 },
    locale: 'ko-KR',
  });
  browser = context;
  const pages = context.pages();
  page = pages.length > 0 ? pages[0] : await context.newPage();
  page.setDefaultTimeout(30000);

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

  // missinstall 쿠키를 미리 설정하여 리다이렉트 방지
  await page.context().addCookies([
    { name: 'missinstall', value: 'Y', domain: new URL(BASE_URL).hostname, path: '/' },
  ]);

  await page.goto(loginUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(3000);

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

  // evaluate로 직접 값 주입 (팝업/오버레이 방해 회피)
  await page.evaluate(({ id, pw }) => {
    const idEl = document.querySelector('input[type="text"], input[type="email"]');
    const pwEl = document.querySelector('input[type="password"]');
    if (idEl) {
      idEl.value = id;
      idEl.dispatchEvent(new Event('input', { bubbles: true }));
      idEl.dispatchEvent(new Event('change', { bubbles: true }));
    }
    if (pwEl) {
      pwEl.value = pw;
      pwEl.dispatchEvent(new Event('input', { bubbles: true }));
      pwEl.dispatchEvent(new Event('change', { bubbles: true }));
    }
  }, { id: userId, pw: userPw });
  await page.waitForTimeout(500);

  logger.info('로그인 버튼 클릭 시도...');
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
  logger.info(`로그인 클릭: ${loginClicked}`);

  await page.waitForTimeout(6000);

  const afterUrl = page.url();
  logger.info(`로그인 후 URL: ${afterUrl}`);

  const hasToolbar = await page.locator('.toolbar2, .co_name, .gnb, .lnb, #main_iframe').count();
  if (hasToolbar > 0) {
    isLoggedIn = true;
    logger.info('경리나라 로그인 성공');
  } else if (!afterUrl.includes('0002_01')) {
    isLoggedIn = true;
    logger.info('경리나라 로그인 성공 (URL 기반 확인)');
  } else {
    const bodyText = await page.evaluate(() => document.body?.innerText?.substring(0, 500) || '');
    logger.error(`로그인 실패. 페이지 내용: ${bodyText}`);
    throw new Error('경리나라 로그인 실패. 아이디/비밀번호를 확인하세요.');
  }
}

async function getMainIframe() {
  await page.waitForTimeout(2000);

  let frame = page.frame({ name: 'main_iframe' });

  if (!frame) {
    const frames = page.frames();
    frame = frames.find(f => f.url().includes('serp.co.kr') && f !== page.mainFrame());
  }

  if (frame) {
    await frame.waitForLoadState('domcontentloaded').catch(() => null);
    await page.waitForTimeout(2000);
    return frame;
  }

  throw new Error('main_iframe을 찾을 수 없습니다');
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
    await page.waitForTimeout(5000);
  }

  let frame = page.frame({ name: 'main_iframe' });

  if (!frame) {
    const frames = page.frames();
    frame = frames.find(f => f !== page.mainFrame() && f.url().includes('serp.co.kr'));
  }

  if (frame) {
    const targetUrl = `${BASE_URL}${actFile}`;
    await frame.goto(targetUrl, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await page.waitForTimeout(3000);
    logger.info(`iframe URL: ${frame.url()}`);
    return frame;
  }

  logger.warn('iframe을 찾을 수 없어 직접 페이지 이동');
  await page.goto(`${BASE_URL}${actFile}`, {
    waitUntil: 'domcontentloaded',
    timeout: 30000,
  });
  await page.waitForTimeout(3000);
  return page.mainFrame();
}

async function setDateRange(frame, startDate, endDate) {
  if (!startDate && !endDate) return;

  const formatDate = (d) => (d || '').replace(/\//g, '-');
  const fmtStart = formatDate(startDate);
  const fmtEnd = formatDate(endDate);

  logger.info(`날짜 범위 설정: ${fmtStart || '(기본)'} ~ ${fmtEnd || '(기본)'}`);

  // hidden 필드가 많으므로 JavaScript evaluate()로 직접 값 설정
  const setResult = await frame.evaluate(({ start, end }) => {
    const log = [];
    const ids = [
      ['txtSrcStartDt', 'txtSrcEndDt'],
      ['txtDtlSrcStartDt', 'txtDtlSrcEndDt'],
    ];
    const names = [
      ['STR_DT', 'END_DT'],
    ];

    function setField(el, val) {
      if (!el || !val) return false;
      const nativeSetter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype, 'value'
      )?.set;
      if (nativeSetter) {
        nativeSetter.call(el, val);
      } else {
        el.value = val;
      }
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      el.dispatchEvent(new Event('blur', { bubbles: true }));
      return true;
    }

    let setCount = 0;
    for (const [startId, endId] of ids) {
      const startEl = document.getElementById(startId);
      const endEl = document.getElementById(endId);
      if (startEl && start) {
        const before = startEl.value;
        setField(startEl, start);
        log.push(`${startId}: "${before}" → "${start}"`);
        setCount++;
      }
      if (endEl && end) {
        const before = endEl.value;
        setField(endEl, end);
        log.push(`${endId}: "${before}" → "${end}"`);
        setCount++;
      }
    }

    for (const [startName, endName] of names) {
      const startEl = document.querySelector(`input[name="${startName}"]`);
      const endEl = document.querySelector(`input[name="${endName}"]`);
      if (startEl && start && !startEl.id) {
        setField(startEl, start);
        log.push(`name=${startName}: → "${start}"`);
        setCount++;
      }
      if (endEl && end && !endEl.id) {
        setField(endEl, end);
        log.push(`name=${endName}: → "${end}"`);
        setCount++;
      }
    }

    // 모든 hasDatepicker 필드도 확인 (위에서 놓친 게 있을 수 있으므로)
    const pickers = document.querySelectorAll('input.hasDatepicker');
    const pickerInfo = Array.from(pickers).map(el => `${el.id}="${el.value}"`);
    log.push(`전체 datepicker(${pickers.length}개): ${pickerInfo.join(', ')}`);

    return { log, setCount };
  }, { start: fmtStart, end: fmtEnd });

  logger.info(`날짜 설정: ${setResult.log.join(' | ')}`);

  if (setResult.setCount === 0) {
    logger.error('날짜 필드를 하나도 설정하지 못했습니다!');
    return;
  }

  await page.waitForTimeout(500);

  // 검색 실행 - fn_search 직접 호출 우선
  const searchResult = await frame.evaluate(() => {
    const log = [];

    // 1. 전역 함수 직접 호출 (가장 확실)
    try {
      if (typeof fn_search === 'function') { fn_search(); log.push('fn_search() 호출 성공'); return log; }
    } catch (e) { log.push(`fn_search 오류: ${e.message}`); }
    try {
      if (typeof fn_Search === 'function') { fn_Search(); log.push('fn_Search() 호출 성공'); return log; }
    } catch (e) { log.push(`fn_Search 오류: ${e.message}`); }
    try {
      if (typeof doSearch === 'function') { doSearch(); log.push('doSearch() 호출 성공'); return log; }
    } catch (e) { log.push(`doSearch 오류: ${e.message}`); }

    // 2. onclick 속성에 검색 함수가 있는 버튼
    const allEls = document.querySelectorAll('[onclick]');
    for (const el of allEls) {
      const oc = el.getAttribute('onclick') || '';
      if (oc.includes('fn_search') || oc.includes('fn_Search') || oc.includes('doSearch')) {
        el.click();
        log.push(`onclick 클릭: ${oc.substring(0, 50)}`);
        return log;
      }
    }

    // 3. 조회 텍스트 버튼
    const btns = document.querySelectorAll('button, a, input[type="button"], img, span');
    for (const btn of btns) {
      const txt = (btn.textContent?.trim() || btn.value || btn.alt || btn.title || '');
      if (txt === '조회' || txt === '검색') {
        btn.click();
        log.push(`텍스트 클릭: "${txt}"`);
        return log;
      }
    }

    log.push('조회 버튼/함수를 찾지 못함');
    return log;
  });

  logger.info(`검색 실행: ${searchResult.join(' | ')}`);
  await page.waitForTimeout(4000);

  // 조회 후 확인
  const afterCheck = await frame.evaluate(() => {
    const fields = ['txtSrcStartDt', 'txtSrcEndDt', 'txtDtlSrcStartDt', 'txtDtlSrcEndDt'];
    const vals = fields.map(id => {
      const el = document.getElementById(id);
      return el ? `${id}="${el.value}"` : null;
    }).filter(Boolean);

    const links = document.querySelectorAll('[onclick*="PopupCall"], [onclick*="fn_PopupCall"]');
    vals.push(`거래처 링크: ${links.length}개`);
    return vals.join(', ');
  });
  logger.info(`조회 후 상태: ${afterCheck}`);
}

async function collectSalesData(options = {}) {
  if (isBusy) throw new Error('이미 수집 중입니다. 잠시 후 다시 시도하세요.');
  isBusy = true;

  const { startDate, endDate } = options;

  try {
    await ensureBrowser();
    await login();

    logger.info(`매출(거래명세표) 페이지 이동... ${startDate ? `(${startDate} ~ ${endDate})` : '(기본 기간)'}`);
    const frame = await navigateToMenu('s1110');

    await setDateRange(frame, startDate, endDate);

    const clientLinks = await frame.evaluate(() => {
      const links = document.querySelectorAll('[onclick*="PopupCall"], [onclick*="fn_PopupCall"]');
      return Array.from(links)
        .map(el => ({
          name: el.textContent.trim(),
          onclick: el.getAttribute('onclick') || '',
        }))
        .filter(item => item.name.length > 1 && !/^[\d,.]+$/.test(item.name));
    });

    logger.info(`${clientLinks.length}개 거래처 발견: ${clientLinks.map(c => c.name).join(', ')}`);

    if (clientLinks.length === 0) {
      const allText = await frame.evaluate(() => document.body?.innerText || '');
      logger.warn(`거래처 링크 없음. 페이지 텍스트 (${allText.length}자): ${allText.substring(0, 300)}`);
      cachedSales = [];
      return [];
    }

    const salesData = [];

    for (const client of clientLinks) {
      try {
        await page.evaluate(() => true);
      } catch {
        logger.warn('브라우저가 닫혔습니다. 수집을 중단합니다.');
        break;
      }

      try {
        logger.info(`거래처 수집: ${client.name}`);
        const clientDetail = await openClientPopupAndExtract(frame, client);
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

    cachedSales = salesData;
    const totalItems = salesData.reduce((s, c) => s + (c.items?.length || 0), 0);
    logger.info(`매출 수집 완료: ${salesData.length}개 거래처, ${totalItems}건`);
    return salesData;

  } catch (err) {
    logger.error('경리나라 매출 데이터 수집 실패', { error: err.message });
    throw err;
  } finally {
    isBusy = false;
  }
}

async function openClientPopupAndExtract(frame, client) {
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

  await popup.waitForLoadState('domcontentloaded');
  await popup.waitForTimeout(2000);

  const data = await popup.evaluate(() => {
    const result = {
      client: '',
      date: '',
      items: [],
      prevBalance: 0,
      totalSupply: 0,
      totalVat: 0,
      totalAmount: 0,
      depositAmount: 0,
      outstandingBalance: 0,
    };

    const tables = document.querySelectorAll('table');

    for (const table of tables) {
      const rows = table.querySelectorAll('tr');
      for (const row of rows) {
        const cells = Array.from(row.querySelectorAll('td, th'));
        const texts = cells.map(c => c.textContent.trim());

        const clientIdx = texts.findIndex(t =>
          t.includes('거 래 처') || t.includes('거래처'));
        if (clientIdx >= 0 && clientIdx < texts.length - 1) {
          result.client = texts[clientIdx + 1] || '';
        }

        const dateIdx = texts.findIndex(t =>
          (t.includes('일') && t.includes('자')) && !t.includes('품'));
        if (dateIdx >= 0 && dateIdx < texts.length - 1) {
          result.date = texts[dateIdx + 1] || '';
        }
      }
    }

    let itemTable = null;
    for (const table of tables) {
      const txt = table.textContent || '';
      if ((txt.includes('품목') || txt.includes('품 목')) &&
          (txt.includes('수량') || txt.includes('수 량')) &&
          (txt.includes('단가') || txt.includes('단 가'))) {
        itemTable = table;
        break;
      }
    }

    if (itemTable) {
      const rows = itemTable.querySelectorAll('tr');
      const colMap = {};
      let headerFound = false;

      const colPatterns = {
        date: ['월일', '일자', '월 일', '일 자'],
        product: ['품목', '품 목', '상품명'],
        spec: ['규격', '규 격'],
        unit: ['단위', '단 위'],
        qty: ['수량', '수 량'],
        unitPrice: ['단가', '단 가'],
        supply: ['공급가액', '공급가', '공 급 가 액'],
        vat: ['세액', '세 액', '부가세', '부가'],
        note: ['비고', '비 고'],
      };

      for (const row of rows) {
        const cells = Array.from(row.querySelectorAll('td, th'));
        const texts = cells.map(c => c.textContent.trim().replace(/\s+/g, ''));

        const hasProduct = texts.some(t => t.includes('품목'));
        const hasQty = texts.some(t => t.includes('수량'));

        if (hasProduct && hasQty) {
          for (let i = 0; i < texts.length; i++) {
            const cellText = texts[i];
            for (const [key, patterns] of Object.entries(colPatterns)) {
              if (patterns.some(p => cellText.replace(/\s/g, '').includes(p.replace(/\s/g, '')))) {
                colMap[key] = i;
              }
            }
          }
          headerFound = true;
          continue;
        }

        if (!headerFound) continue;
        if (texts.length < 4) continue;
        if (texts.every(t => !t || t === '')) continue;
        if (texts.some(t => t.includes('합계') || t.includes('소계'))) continue;

        const get = (key) => (colMap[key] !== undefined ? texts[colMap[key]] : '') || '';
        const pn = (s) => parseFloat((s || '').replace(/[,\s]/g, '')) || 0;

        const product = get('product');
        const dateCell = get('date');

        if (product && product.length > 0 && product !== dateCell) {
          result.items.push({
            deliveryDate: dateCell,
            product,
            spec: get('spec'),
            qty: pn(get('qty')),
            unitPrice: pn(get('unitPrice')),
            supplyAmount: pn(get('supply')),
            vat: pn(get('vat')),
            total: pn(get('supply')) + pn(get('vat')),
            note: get('note'),
          });
        }
      }

      result._colMap = colMap;
    }

    for (const table of tables) {
      for (const row of table.querySelectorAll('tr')) {
        const text = row.textContent || '';
        const pn = (s) => parseFloat((s || '').replace(/[,\s]/g, '')) || 0;

        if (text.includes('전미수잔액')) {
          const nums = text.match(/[\d,]+/g);
          if (nums) result.prevBalance = pn(nums[nums.length - 1]);
        }
        if (text.includes('합계') && !text.includes('총합계') && !text.includes('전미수')) {
          const vals = Array.from(row.querySelectorAll('td'))
            .map(c => c.textContent.trim()).filter(t => /^[\d,]+$/.test(t));
          if (vals.length >= 2) {
            result.totalSupply = pn(vals[0]);
            result.totalVat = pn(vals[1]);
          }
        }
        if (text.includes('총합계') || text.includes('합계금액')) {
          const m = text.match(/([\d,]+)\s*$/);
          if (m) result.totalAmount = pn(m[1]);
        }
        if (text.includes('입금액')) {
          const vals = Array.from(row.querySelectorAll('td'))
            .map(c => c.textContent.trim().replace(/[,\s]/g, ''))
            .filter(v => /^\d+$/.test(v) && parseInt(v) > 0);
          if (vals.length > 0) result.depositAmount = parseFloat(vals[0]) || 0;
        }
        if (text.includes('총미수잔액') || text.includes('미수잔액')) {
          const vals = Array.from(row.querySelectorAll('td'))
            .map(c => c.textContent.trim().replace(/[,\s]/g, ''))
            .filter(v => /^\d+$/.test(v) && parseInt(v) > 0);
          if (vals.length > 0) result.outstandingBalance = parseFloat(vals[0]) || 0;
        }
      }
    }

    result.totalAmount = result.totalAmount ||
      result.items.reduce((s, i) => s + i.total, 0);

    return result;
  });

  await popup.close().catch(() => null);
  await page.waitForTimeout(300);

  if (!data.client) data.client = client.name;
  return data;
}

async function collectDepositData(options = {}) {
  if (isBusy) throw new Error('이미 수집 중입니다. 잠시 후 다시 시도하세요.');
  isBusy = true;

  const { startDate, endDate } = options;

  try {
    await ensureBrowser();
    await login();

    logger.info(`입출내역조회 페이지 이동... ${startDate ? `(${startDate} ~ ${endDate})` : '(기본 기간)'}`);
    const frame = await navigateToMenu('s3120');

    if (startDate || endDate) {
      await setDateRange(frame, startDate, endDate);
    } else {
      await page.waitForTimeout(2000);
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

    cachedDeposits = deposits;
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

function getCachedSales() { return cachedSales; }
function getCachedDeposits() { return cachedDeposits; }

module.exports = {
  collectSalesData,
  collectDepositData,
  closeBrowser,
  getCachedSales,
  getCachedDeposits,
};
