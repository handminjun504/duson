const API = '/api';

// === Tab Navigation ===
document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(`tab-${btn.dataset.tab}`).classList.add('active');

    if (btn.dataset.tab === 'raw-sheets') loadRawSheets();
    if (btn.dataset.tab === 'raw-gyeongli') loadRawGyeongli();
  });
});

// === Utilities ===
function showLoading(text = '처리 중...') {
  document.getElementById('loadingText').textContent = text;
  document.getElementById('loading').classList.add('show');
}

function hideLoading() {
  document.getElementById('loading').classList.remove('show');
}

function toast(message, type = 'info') {
  const container = document.getElementById('toastContainer');
  const el = document.createElement('div');
  el.className = `toast toast-${type}`;
  el.textContent = message;
  container.appendChild(el);
  setTimeout(() => { el.style.opacity = '0'; setTimeout(() => el.remove(), 300); }, 3000);
}

function formatNumber(n) {
  if (n == null || n === '') return '-';
  return Number(n).toLocaleString('ko-KR');
}

function esc(str) {
  if (!str) return '';
  const d = document.createElement('div');
  d.textContent = String(str);
  return d.innerHTML;
}

async function apiCall(method, endpoint, body = null) {
  const opts = { method, headers: { 'Content-Type': 'application/json' } };
  if (body) opts.body = JSON.stringify(body);
  const res = await fetch(`${API}${endpoint}`, opts);
  if (!res.ok) {
    const err = await res.json().catch(() => ({ error: '서버 오류' }));
    throw new Error(err.error || `HTTP ${res.status}`);
  }
  return res.json();
}

// === Dashboard ===
async function loadStatus() {
  try {
    const st = await apiCall('GET', '/status');
    document.getElementById('statSheetCount').textContent = st.sheetCount || '-';
    document.getElementById('statGyeongliCount').textContent = st.gyeongliItemCount || '-';
    document.getElementById('statMismatch').textContent = st.compared ? st.mismatchCount : '-';
    document.getElementById('statDeposits').textContent = st.depositCount || '-';
    document.getElementById('statusText').textContent = st.geminiAvailable ? 'Gemini 연결됨' : 'Gemini 미설정';
  } catch {
    document.getElementById('statusText').textContent = '서버 연결 실패';
  }
}

function addLog(html) {
  const log = document.getElementById('dashboardLog');
  if (log.querySelector('.empty-state')) log.innerHTML = '';
  const entry = document.createElement('div');
  entry.style.cssText = 'padding:8px 0;border-bottom:1px solid var(--lf-border);font-size:13px';
  entry.innerHTML = `<span style="color:var(--lf-text-light);font-size:11px">${new Date().toLocaleTimeString()}</span> ${html}`;
  log.prepend(entry);
}

// === Data Collection ===
async function collectSheets() {
  const dates = getDateParams();
  const label = dates.startDate ? `${dates.startDate} ~ ${dates.endDate}` : '전체';
  showLoading(`Google Sheet 매출 데이터 수집 중... (${label})`);
  try {
    const data = await apiCall('POST', '/collect/sheets', dates);
    toast(`Google Sheet ${data.count}건 수집 완료 (${label})`, 'success');
    addLog(`<span class="lf-badge lf-badge-success">완료</span> Google Sheet에서 <strong>${data.count}건</strong> 매출 데이터 수집 (${label})`);
    loadStatus();
  } catch (err) {
    toast(`수집 실패: ${err.message}`, 'error');
    addLog(`<span class="lf-badge lf-badge-danger">실패</span> Google Sheet 수집 오류: ${err.message}`);
  } finally {
    hideLoading();
  }
}

function setThisMonth() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  document.getElementById('dateStart').value = `${y}-${m}-01`;
  document.getElementById('dateEnd').value = `${y}-${m}-${String(new Date(y, now.getMonth() + 1, 0).getDate()).padStart(2, '0')}`;
}

function setLastMonth() {
  const now = new Date();
  const prev = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const y = prev.getFullYear();
  const m = String(prev.getMonth() + 1).padStart(2, '0');
  document.getElementById('dateStart').value = `${y}-${m}-01`;
  document.getElementById('dateEnd').value = `${y}-${m}-${String(new Date(y, prev.getMonth() + 1, 0).getDate()).padStart(2, '0')}`;
}

function setThisYear() {
  const y = new Date().getFullYear();
  document.getElementById('dateStart').value = `${y}-01-01`;
  document.getElementById('dateEnd').value = `${y}-12-31`;
}

function getDateParams() {
  const s = document.getElementById('dateStart').value;
  const e = document.getElementById('dateEnd').value;
  const body = {};
  if (s) body.startDate = s;
  if (e) body.endDate = e;
  return body;
}

async function collectGyeongli() {
  const dates = getDateParams();
  const label = dates.startDate ? `${dates.startDate} ~ ${dates.endDate}` : '기본 기간';
  showLoading(`경리나라 수집 중... (${label})`);
  try {
    const data = await apiCall('POST', '/collect/gyeongli', dates);
    const total = data.clients.reduce((s, c) => s + c.itemCount, 0);
    toast(`경리나라 ${data.clients.length}개 거래처 (${total}건) 수집 완료`, 'success');
    addLog(`<span class="lf-badge lf-badge-success">완료</span> 경리나라 <strong>${data.clients.length}개 거래처, ${total}건</strong> 매출 + <strong>${data.depositCount}건</strong> 입금 수집 (${label})`);
    loadStatus();
    loadClients();
    loadDeposits();
  } catch (err) {
    toast(`수집 실패: ${err.message}`, 'error');
    addLog(`<span class="lf-badge lf-badge-danger">실패</span> 경리나라 수집 오류: ${err.message}`);
  } finally {
    hideLoading();
  }
}

// === Compare ===
async function runCompare() {
  showLoading('매출 대조 실행 중...');
  try {
    const data = await apiCall('POST', '/compare');
    renderCompareResult(data);
    toast('매출 대조 완료', 'info');
    addLog(`<span class="lf-badge lf-badge-primary">대조</span> Google Sheet ${data.summary?.sheetCount || 0}건 vs 경리나라 ${data.summary?.gyeongliCount || 0}건`);
    loadStatus();
  } catch (err) {
    toast(`대조 실패: ${err.message}`, 'error');
  } finally {
    hideLoading();
  }
}

function renderCompareResult(data) {
  const el = document.getElementById('compareResult');
  const s = data.summary || {};

  if (!s.sheetCount && !s.gyeongliCount) {
    el.innerHTML = '<div class="empty-state"><i class="ri-check-double-line"></i><p>대조할 데이터가 없습니다. 양쪽 데이터를 먼저 수집하세요.</p></div>';
    return;
  }

  let html = `
    <div style="display:flex;gap:12px;margin-bottom:20px;flex-wrap:wrap">
      <div class="stat-card" style="flex:1;min-width:140px">
        <div class="stat-label">Google Sheet (매출)</div>
        <div class="stat-value" style="font-size:20px">${s.sheetCount || 0}건</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:140px">
        <div class="stat-label">경리나라</div>
        <div class="stat-value" style="font-size:20px">${s.gyeongliCount || 0}건</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:140px">
        <div class="stat-label">✅ 일치</div>
        <div class="stat-value" style="font-size:20px;color:#27AE60">${s.matchedCount || 0}건</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:140px">
        <div class="stat-label">⚠️ 불일치</div>
        <div class="stat-value" style="font-size:20px;color:#E67E22">${s.mismatchCount || 0}건</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:140px">
        <div class="stat-label">📋 시트만</div>
        <div class="stat-value" style="font-size:20px;color:#3498DB">${s.sheetOnlyCount || 0}건</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:140px">
        <div class="stat-label">📋 경리만</div>
        <div class="stat-value" style="font-size:20px;color:#9B59B6">${s.gyeongliOnlyCount || 0}건</div>
      </div>
    </div>`;

  let filterHtml = `<div style="margin-bottom:16px;display:flex;gap:8px;flex-wrap:wrap">
    <button class="lf-btn lf-btn-sm compare-filter active" data-filter="all" onclick="filterCompare('all')">전체</button>
    <button class="lf-btn lf-btn-sm compare-filter" data-filter="mismatch" onclick="filterCompare('mismatch')">⚠️ 불일치 (${s.mismatchCount || 0})</button>
    <button class="lf-btn lf-btn-sm compare-filter" data-filter="sheetOnly" onclick="filterCompare('sheetOnly')">시트만 (${s.sheetOnlyCount || 0})</button>
    <button class="lf-btn lf-btn-sm compare-filter" data-filter="gyeongliOnly" onclick="filterCompare('gyeongliOnly')">경리만 (${s.gyeongliOnlyCount || 0})</button>
    <button class="lf-btn lf-btn-sm compare-filter" data-filter="matched" onclick="filterCompare('matched')">✅ 일치 (${s.matchedCount || 0})</button>
  </div>`;
  html += filterHtml;

  html += `<div style="overflow-x:auto"><table class="data-table" id="compareTable"><thead><tr>
    <th>상태</th><th>날짜</th><th>거래처</th><th>품목명</th>
    <th class="num">수량</th>
    <th class="num">시트 공급가</th><th class="num">경리 공급가</th>
    <th class="num">시트 세액</th><th class="num">경리 세액</th>
    <th class="num">시트 합계</th><th class="num">경리 합계</th>
    <th>차이 내용</th></tr></thead><tbody>`;

  const allItems = [
    ...(data.mismatch || []).map(i => ({ ...i, _cat: 'mismatch' })),
    ...(data.sheetOnly || []).map(i => ({ ...i, _cat: 'sheetOnly' })),
    ...(data.gyeongliOnly || []).map(i => ({ ...i, _cat: 'gyeongliOnly' })),
    ...(data.matched || []).map(i => ({ ...i, _cat: 'matched' })),
  ];

  for (const item of allItems) {
    let badge = '';
    let rowStyle = '';
    if (item._cat === 'matched') {
      badge = '<span class="lf-badge lf-badge-success">일치</span>';
      if (item._note) badge += ` <span style="font-size:10px;color:var(--lf-text-light)">${esc(item._note)}</span>`;
    }
    else if (item._cat === 'mismatch') {
      const level = item._matchLevel || '';
      badge = level
        ? `<span class="lf-badge lf-badge-warning" title="${level}">불일치</span>`
        : '<span class="lf-badge lf-badge-warning">불일치</span>';
      rowStyle = 'background:#FFF8E1;';
    }
    else if (item._cat === 'sheetOnly') { badge = '<span class="lf-badge lf-badge-primary">시트만</span>'; rowStyle = 'background:#E3F2FD;'; }
    else { badge = '<span class="lf-badge" style="background:#F3E5F5;color:#7B1FA2">경리만</span>'; rowStyle = 'background:#F3E5F5;'; }

    const dateDisplay = item.sheetDate || item.gyeongliDate || item.tradeDate || item.deliveryDate || '';
    const ss = item.sheetSupply ?? item.supplyAmount ?? '';
    const gs = item.gyeongliSupply ?? item.supplyAmount ?? '';
    const sv = item.sheetVat ?? item.vat ?? '';
    const gv = item.gyeongliVat ?? item.vat ?? '';
    const st = item.sheetTotal ?? item.total ?? '';
    const gt = item.gyeongliTotal ?? item.total ?? '';
    const diffText = (item.diffs || []).join('<br>');

    html += `<tr class="compare-row" data-cat="${item._cat}" style="${rowStyle}">
      <td>${badge}</td>
      <td style="white-space:nowrap">${esc(dateDisplay)}</td>
      <td>${esc(item.client)}</td>
      <td>${esc(item.product || item.productName)}</td>
      <td class="num">${formatNumber(item.qty)}</td>
      <td class="num">${formatNumber(ss)}</td>
      <td class="num">${formatNumber(gs)}</td>
      <td class="num">${formatNumber(sv)}</td>
      <td class="num">${formatNumber(gv)}</td>
      <td class="num">${formatNumber(st)}</td>
      <td class="num">${formatNumber(gt)}</td>
      <td style="font-size:11px;max-width:220px">${diffText || '-'}</td>
    </tr>`;
  }

  html += '</tbody></table></div>';
  el.innerHTML = html;
}

function filterCompare(cat) {
  document.querySelectorAll('.compare-filter').forEach(b => b.classList.remove('active'));
  document.querySelector(`.compare-filter[data-filter="${cat}"]`)?.classList.add('active');
  document.querySelectorAll('.compare-row').forEach(row => {
    row.style.display = (cat === 'all' || row.dataset.cat === cat) ? '' : 'none';
  });
}

// === AI Analysis ===
async function runAnalysis() {
  showLoading('Gemini AI 분석 실행 중...');
  try {
    const data = await apiCall('POST', '/analyze');
    if (!data.available) {
      toast(data.message, 'error');
      addLog(`<span class="lf-badge lf-badge-warning">AI</span> ${data.message}`);
      return;
    }
    const summary = data.analysis?.summary || JSON.stringify(data.analysis);
    toast('AI 분석 완료', 'success');
    addLog(`<span class="ai-badge"><i class="ri-sparkling-line"></i> AI</span> ${summary}`);
  } catch (err) {
    toast(`AI 분석 실패: ${err.message}`, 'error');
  } finally {
    hideLoading();
  }
}

// === Clients ===
let allClientsData = [];
let allClientsDetail = {};

let clientsLoading = false;
async function loadClients() {
  if (clientsLoading) return;
  clientsLoading = true;
  try {
    const data = await apiCall('GET', '/clients');
    allClientsData = data.clients || [];

    // 상위 10개만 상세 조회 (나머지는 클릭 시 로드)
    const top = allClientsData.slice(0, 10);
    const details = await Promise.all(
      top.map(c =>
        apiCall('GET', `/clients/${encodeURIComponent(c.name)}/transactions`).catch(() => null)
      )
    );
    allClientsDetail = {};
    details.forEach(d => { if (d) allClientsDetail[d.client] = d; });

    renderAllClients(allClientsData);
  } catch (err) {
    toast(`거래처 목록 로드 실패: ${err.message}`, 'error');
  } finally {
    clientsLoading = false;
  }
}

function filterAllClients(keyword) {
  if (!keyword) {
    renderAllClients(allClientsData);
    return;
  }
  const kw = keyword.toLowerCase();
  renderAllClients(allClientsData.filter(c => c.name.toLowerCase().includes(kw)));
}

function renderAllClients(clients) {
  const el = document.getElementById('clientsAllView');

  if (!clients.length) {
    el.innerHTML = '<div class="empty-state"><i class="ri-building-line"></i><p>거래처 데이터 없음</p></div>';
    return;
  }

  let html = '';
  for (const c of clients) {
    const data = allClientsDetail[c.name];
    if (data) {
      html += renderClientLedger(data);
    } else {
      html += `<div class="ledger-card" id="client-${btoa(encodeURIComponent(c.name))}">
        <div class="ledger-header" style="cursor:pointer" onclick="loadClientDetail('${esc(c.name)}')">
          <div class="ledger-header-name">${esc(c.name)}</div>
          <div class="ledger-header-stats">
            <span class="ledger-stat">품목 <strong>${c.itemCount}</strong>건</span>
            <span class="ledger-stat">매출합계 <strong style="color:#4A8CDB">${formatNumber(c.totalAmount)}</strong></span>
            <span style="color:var(--lf-text-light);font-size:12px">클릭하여 상세 보기</span>
          </div>
        </div>
      </div>`;
    }
  }
  el.innerHTML = html;
}

async function loadClientDetail(name) {
  try {
    const data = await apiCall('GET', `/clients/${encodeURIComponent(name)}/transactions`);
    allClientsDetail[name] = data;
    const card = document.getElementById(`client-${btoa(encodeURIComponent(name))}`);
    if (card) card.outerHTML = renderClientLedger(data);
  } catch (err) {
    toast(`${name} 상세 로드 실패: ${err.message}`, 'error');
  }
}

function renderClientLedger(data) {
  const totalSupply = data.items.reduce((s, i) => s + (i.supplyAmount || 0), 0);
  const totalVat = data.items.reduce((s, i) => s + (i.vat || 0), 0);
  const totalAmount = data.items.reduce((s, i) => s + (i.total || 0), 0);

  const timeline = [];
  const dateGroups = {};
  for (const item of data.items) {
    const key = item.deliveryDate || '날짜 미상';
    if (!dateGroups[key]) dateGroups[key] = [];
    dateGroups[key].push(item);
  }
  for (const [date, items] of Object.entries(dateGroups)) {
    timeline.push({ date, type: '매출', items });
  }
  if (data.deposits) {
    for (const dep of data.deposits) {
      const bankInfo = [dep.bank, dep.account].filter(Boolean).join(' ');
      timeline.push({
        date: dep.date || '',
        type: '계좌',
        depositDetail: bankInfo || dep.detail || dep.client || '',
        depositAmount: dep.amount || 0,
      });
    }
  }
  timeline.sort((a, b) => (a.date || '').localeCompare(b.date || ''));

  let html = `<div class="ledger-card">
    <div class="ledger-header">
      <div class="ledger-header-name">${data.client}</div>
      <div class="ledger-header-stats">
        <span class="ledger-stat">이전잔액 <strong>${formatNumber(data.prevBalance || 0)}</strong></span>
        <span class="ledger-stat">매출 <strong style="color:#4A8CDB">${formatNumber(totalAmount)}</strong></span>
        <span class="ledger-stat">입금 <strong style="color:#2EAA5C">${formatNumber(data.depositAmount || 0)}</strong></span>
        <span class="ledger-stat">미수 <strong style="color:#E05252">${formatNumber(data.outstandingBalance || 0)}</strong></span>
      </div>
    </div>
    <table class="ledger-table">
      <thead><tr>
        <th style="width:95px">거래일자</th>
        <th style="width:55px">거래구분</th>
        <th>세부내역</th>
        <th class="num" style="width:65px">수량</th>
        <th class="num" style="width:75px">단가</th>
        <th class="num" style="width:95px">공급가</th>
        <th class="num" style="width:85px">부/가</th>
      </tr></thead><tbody>`;

  html += `<tr class="balance-row">
    <td colspan="2" style="font-weight:600;color:#8B95A5">이전잔액</td>
    <td colspan="3"></td>
    <td class="num" colspan="2" style="font-weight:700;font-size:13px">${formatNumber(data.prevBalance || 0)}</td>
  </tr>`;

  for (const entry of timeline) {
    if (entry.type === '매출') {
      const ds = entry.items.reduce((s, i) => s + (i.supplyAmount || 0), 0);
      const dv = entry.items.reduce((s, i) => s + (i.vat || 0), 0);

      html += `<tr class="date-row">
        <td style="font-weight:700;color:#2C3E50">${entry.date}</td>
        <td><span class="badge-sale">매출</span></td>
        <td colspan="3"></td>
        <td class="num" style="font-weight:700;color:#2C3E50">${formatNumber(ds)}</td>
        <td class="num" style="font-weight:600;color:#6B7A8D">${formatNumber(dv)}</td>
      </tr>`;

      for (const item of entry.items) {
        const name = item.spec ? `${item.product}/${item.spec}` : item.product;
        html += `<tr class="item-row">
          <td></td><td></td>
          <td>${name}</td>
          <td class="num">${formatNumber(item.qty)}</td>
          <td class="num">${formatNumber(item.unitPrice)}</td>
          <td class="num">${formatNumber(item.supplyAmount)}</td>
          <td class="num">${formatNumber(item.vat)}</td>
        </tr>`;
      }
    } else {
      html += `<tr class="deposit-row">
        <td style="font-weight:700;color:#2C3E50">${entry.date}</td>
        <td><span class="badge-deposit">계좌</span></td>
        <td colspan="3" style="color:#2EAA5C">${entry.depositDetail}</td>
        <td class="num" style="color:#2EAA5C;font-weight:700" colspan="2">${formatNumber(entry.depositAmount)}</td>
      </tr>`;
    }
  }

  html += `</tbody></table></div>`;
  return html;
}

// === Deposits ===
async function loadDeposits() {
  try {
    const data = await apiCall('GET', '/deposits');
    renderDeposits(data);
  } catch (err) {
    toast(`입금 내역 로드 실패: ${err.message}`, 'error');
  }
}

function renderDeposits(data) {
  const el = document.getElementById('depositResult');

  if (!data.deposits?.length) {
    el.innerHTML = '<div class="empty-state"><i class="ri-bank-card-line"></i><p>입금 내역이 없습니다</p></div>';
    return;
  }

  const totalDeposit = data.deposits.reduce((s, d) => s + (d.amount || 0), 0);
  const basic = data.matchResults?.basic;
  const matchedKeys = new Set();
  if (basic?.matched) {
    for (const m of basic.matched) {
      matchedKeys.add(`${m.date}|${m.amount}|${m.client}`);
    }
  }
  const unmatchedCount = basic ? (basic.unmatched?.length ?? 0) : data.deposits.length;

  let html = `
    <div style="display:flex;gap:16px;margin-bottom:16px;flex-wrap:wrap">
      <div class="stat-card" style="flex:1;min-width:200px">
        <div class="stat-label">입금 건수</div>
        <div class="stat-value" style="font-size:22px">${data.deposits.length}건</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:200px">
        <div class="stat-label">입금 합계</div>
        <div class="stat-value" style="font-size:22px;color:#27AE60">${formatNumber(totalDeposit)}원</div>
      </div>
      <div class="stat-card" style="flex:1;min-width:200px">
        <div class="stat-label">미매칭 건</div>
        <div class="stat-value" style="font-size:22px;color:#E74C3C">${unmatchedCount}건</div>
      </div>
    </div>
    <div style="margin-bottom:12px;display:flex;gap:8px;flex-wrap:wrap">
      <button class="lf-btn lf-btn-sm deposit-filter active" data-filter="all" onclick="filterDeposits('all')">전체</button>
      <button class="lf-btn lf-btn-sm deposit-filter" data-filter="unmatched" onclick="filterDeposits('unmatched')">매칭 안된 건만 (${unmatchedCount})</button>
    </div>
    <div style="overflow-x:auto">
    <table class="data-table" style="margin-bottom:24px" id="depositsTable">
      <thead><tr>
        <th>입금일</th><th>적요 / 거래처</th><th class="num">입금액</th><th>시간</th><th>은행 / 계좌</th><th>매칭 상태</th>
      </tr></thead><tbody>`;

  for (const d of data.deposits) {
    const clientName = d.client || '-';
    const time = d.time || '';
    const bankAccount = [d.bank, d.account].filter(Boolean).join(' / ') || d.detail || '-';
    const isMatched = matchedKeys.has(`${d.date}|${d.amount}|${d.client}`);
    const matchedTo = basic?.matched?.find(m => m.date === d.date && m.amount === d.amount && m.client === d.client)?.matchedTo;
    const rowClass = isMatched ? 'deposit-row-matched' : 'deposit-row-unmatched';
    const badge = isMatched
      ? `<span class="lf-badge lf-badge-success">매칭됨</span>${matchedTo ? ` <span style="font-size:11px;color:var(--lf-text-light)">&rarr; ${matchedTo}</span>` : ''}`
      : '<span class="lf-badge lf-badge-warning">미매칭</span>';
    const clientCell = isMatched
      ? clientName
      : `<strong style="color:#2C3E50">${clientName}</strong>`;
    const amountCell = isMatched
      ? formatNumber(d.amount)
      : `<strong>${formatNumber(d.amount)}</strong>`;

    html += `<tr class="deposit-table-row ${rowClass}" data-matched="${isMatched}">
      <td>${d.date || '-'}</td>
      <td>${clientCell}</td>
      <td class="num" style="color:#27AE60;font-weight:700">${amountCell}</td>
      <td style="font-size:12px;color:var(--lf-text-light)">${time}</td>
      <td style="font-size:12px;color:var(--lf-text-light)">${bankAccount}</td>
      <td>${badge}</td>
    </tr>`;
  }

  html += '</tbody></table></div>';

  // 미매칭 건 요약 카드
  const unmatchedDeposits = data.deposits.filter(d => !matchedKeys.has(`${d.date}|${d.amount}|${d.client}`));
  if (unmatchedDeposits.length > 0) {
    html += `<div class="detail-panel" style="margin-top:20px;border-left:4px solid #E74C3C">
      <h4 style="margin:0 0 12px;color:#E74C3C"><i class="ri-error-warning-line"></i> 미매칭 입금 내역 (${unmatchedDeposits.length}건)</h4>
      <table class="data-table"><thead><tr>
        <th>입금일</th><th>거래처</th><th class="num">입금액</th><th>은행 / 계좌</th>
      </tr></thead><tbody>`;
    for (const d of unmatchedDeposits) {
      html += `<tr style="background:#FFF8E1">
        <td style="font-weight:600">${d.date || '-'}</td>
        <td><strong style="color:#2C3E50">${esc(d.client || '-')}</strong></td>
        <td class="num" style="color:#E74C3C;font-weight:700;font-size:14px">${formatNumber(d.amount)}</td>
        <td style="font-size:12px;color:var(--lf-text-light)">${[d.bank, d.account].filter(Boolean).join(' / ') || '-'}</td>
      </tr>`;
    }
    html += '</tbody></table></div>';
  }

  if (data.matchResults?.gemini?.available) {
    html += `<div class="detail-panel" style="margin-top:16px">
      <h4><span class="ai-badge"><i class="ri-sparkling-line"></i> AI</span> 매칭 제안</h4>
      <div class="report-box" style="margin-top:12px">${data.matchResults.gemini.analysis?.summary || JSON.stringify(data.matchResults.gemini.analysis)}</div>
    </div>`;
  }

  el.innerHTML = html;
}

function filterDeposits(filter) {
  document.querySelectorAll('.deposit-filter').forEach(b => b.classList.remove('active'));
  document.querySelector(`.deposit-filter[data-filter="${filter}"]`)?.classList.add('active');
  document.querySelectorAll('.deposit-table-row').forEach(row => {
    const isMatched = row.dataset.matched === 'true';
    if (filter === 'all') {
      row.style.display = '';
    } else if (filter === 'unmatched') {
      row.style.display = isMatched ? 'none' : '';
    }
  });
}

async function runDepositMatch() {
  showLoading('입금 매칭 분석 중...');
  try {
    const data = await apiCall('POST', '/match');
    toast('입금 매칭 분석 완료', 'success');
    await loadDeposits();
  } catch (err) {
    toast(`입금 매칭 실패: ${err.message}`, 'error');
  } finally {
    hideLoading();
  }
}

// === Generate ===
async function generateExcel() {
  showLoading('거래처원장 Excel 생성 중...');
  try {
    const data = await apiCall('POST', '/generate');
    toast(`거래처원장 생성 완료! (${data.clientCount}개 거래처)`, 'success');
    document.getElementById('generateStatus').innerHTML = `
      <div style="padding:16px;background:#E8F5E9;border-radius:8px;display:inline-block">
        <i class="ri-check-line" style="color:#27AE60;font-size:20px"></i>
        <strong style="color:#27AE60">생성 완료!</strong>
        <span style="color:var(--lf-text-light);margin-left:8px">${data.filename} (${data.clientCount}개 거래처)</span>
      </div>`;
    document.getElementById('downloadBtn').style.display = 'inline-flex';
    loadStatus();

    generateReport();
  } catch (err) {
    toast(`생성 실패: ${err.message}`, 'error');
    document.getElementById('generateStatus').innerHTML = `
      <div style="padding:16px;background:#FDEDEC;border-radius:8px;display:inline-block">
        <i class="ri-error-warning-line" style="color:#E74C3C;font-size:20px"></i>
        <strong style="color:#E74C3C">생성 실패</strong>
        <span style="color:var(--lf-text-light);margin-left:8px">${err.message}</span>
      </div>`;
  } finally {
    hideLoading();
  }
}

function downloadExcel() {
  window.open(`${API}/download`, '_blank');
}

async function generateReport() {
  try {
    const data = await apiCall('POST', '/report');
    if (data.available && data.report) {
      document.getElementById('aiReport').style.display = 'block';
      document.getElementById('reportContent').textContent = data.report;
    }
  } catch {
    // 리포트는 보너스 기능이므로 실패해도 무시
  }
}

// === Raw Data Tabs ===
let _rawSheetsData = null;
let _rawGyeongliData = null;

async function loadRawSheets() {
  const container = document.getElementById('rawSheetsContent');
  const countEl = document.getElementById('rawSheetsCount');
  try {
    const res = await apiCall('GET', '/raw/sheets');
    _rawSheetsData = res.data || [];
    countEl.textContent = `(${_rawSheetsData.length}건)`;
    renderRawSheetsTable(_rawSheetsData);
  } catch {
    container.innerHTML = '<p style="color:var(--lf-text-light);text-align:center;padding:48px">시트 데이터를 먼저 수집하세요.</p>';
    countEl.textContent = '';
  }
}

function renderRawSheetsTable(rows) {
  const container = document.getElementById('rawSheetsContent');
  if (!rows || rows.length === 0) {
    container.innerHTML = '<p style="color:var(--lf-text-light);text-align:center;padding:48px">데이터 없음</p>';
    return;
  }

  let html = '<div style="overflow-x:auto"><table class="data-table"><thead><tr>';
  html += '<th>No</th><th>거래일자</th><th>거래처명</th><th>품목명</th><th>규격</th>';
  html += '<th class="num">수량</th><th class="num">단가</th><th class="num">공급가액</th>';
  html += '<th class="num">세액</th><th class="num">합계금액</th><th>비고</th>';
  html += '</tr></thead><tbody>';

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    html += `<tr>
      <td>${i + 1}</td>
      <td>${esc(r.date || r.transactionDate || '')}</td>
      <td>${esc(r.salesClient || r.client || '')}</td>
      <td>${esc(r.productName || r.product || '')}</td>
      <td>${esc(r.spec || r.standard || '')}</td>
      <td class="num">${formatNumber(r.quantity)}</td>
      <td class="num">${formatNumber(r.salesUnitPrice || r.unitPrice)}</td>
      <td class="num">${formatNumber(r.salesSupply || r.supplyAmount)}</td>
      <td class="num">${formatNumber(r.salesTax || r.tax)}</td>
      <td class="num">${formatNumber(r.salesTotal || r.total)}</td>
      <td>${esc(r.note || '')}</td>
    </tr>`;
  }
  html += '</tbody></table></div>';
  container.innerHTML = html;
}

async function loadRawGyeongli() {
  const container = document.getElementById('rawGyeongliContent');
  const countEl = document.getElementById('rawGyeongliCount');
  try {
    const res = await apiCall('GET', '/raw/gyeongli');
    _rawGyeongliData = res.data || [];
    countEl.textContent = `(${_rawGyeongliData.length}건)`;
    renderRawGyeongliTable(_rawGyeongliData);
  } catch {
    container.innerHTML = '<p style="color:var(--lf-text-light);text-align:center;padding:48px">경리나라 데이터를 먼저 수집하세요.</p>';
    countEl.textContent = '';
  }
}

function renderRawGyeongliTable(rows) {
  const container = document.getElementById('rawGyeongliContent');
  if (!rows || rows.length === 0) {
    container.innerHTML = '<p style="color:var(--lf-text-light);text-align:center;padding:48px">데이터 없음</p>';
    return;
  }

  let html = '<div style="overflow-x:auto"><table class="data-table"><thead><tr>';
  html += '<th>No</th><th>거래처</th><th>납품일</th><th>품목명</th><th>규격</th>';
  html += '<th class="num">수량</th><th class="num">단가</th><th class="num">공급가액</th>';
  html += '<th class="num">세액</th><th class="num">합계금액</th><th>비고</th>';
  html += '</tr></thead><tbody>';

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    html += `<tr>
      <td>${i + 1}</td>
      <td>${esc(r.client || '')}</td>
      <td>${esc(r.deliveryDate || r.date || '')}</td>
      <td>${esc(r.productName || r.product || '')}</td>
      <td>${esc(r.spec || r.standard || '')}</td>
      <td class="num">${formatNumber(r.quantity)}</td>
      <td class="num">${formatNumber(r.unitPrice)}</td>
      <td class="num">${formatNumber(r.supplyAmount)}</td>
      <td class="num">${formatNumber(r.vat || r.tax)}</td>
      <td class="num">${formatNumber(r.total)}</td>
      <td>${esc(r.note || r.remark || '')}</td>
    </tr>`;
  }
  html += '</tbody></table></div>';
  container.innerHTML = html;
}

document.getElementById('rawSheetsSearch')?.addEventListener('input', function() {
  if (!_rawSheetsData) return;
  const q = this.value.trim().toLowerCase();
  if (!q) { renderRawSheetsTable(_rawSheetsData); return; }
  const filtered = _rawSheetsData.filter(r => {
    const text = [r.salesClient, r.client, r.productName, r.product, r.note, r.date, r.transactionDate].join(' ').toLowerCase();
    return text.includes(q);
  });
  renderRawSheetsTable(filtered);
  document.getElementById('rawSheetsCount').textContent = `(${filtered.length}/${_rawSheetsData.length}건)`;
});

document.getElementById('rawGyeongliSearch')?.addEventListener('input', function() {
  if (!_rawGyeongliData) return;
  const q = this.value.trim().toLowerCase();
  if (!q) { renderRawGyeongliTable(_rawGyeongliData); return; }
  const filtered = _rawGyeongliData.filter(r => {
    const text = [r.client, r.productName, r.product, r.note, r.remark, r.deliveryDate, r.date].join(' ').toLowerCase();
    return text.includes(q);
  });
  renderRawGyeongliTable(filtered);
  document.getElementById('rawGyeongliCount').textContent = `(${filtered.length}/${_rawGyeongliData.length}건)`;
});

// === Init ===
setThisMonth();
loadStatus();
