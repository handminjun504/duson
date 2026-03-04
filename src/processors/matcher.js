const logger = require('../utils/logger');

function normalizeClient(name) {
  return (name || '')
    .replace(/\s+/g, '')
    .replace(/[()（）[\]]/g, '')
    .replace(/주식회사/g, '')
    .replace(/㈜/g, '')
    .replace(/\(주\)/g, '')
    .replace(/유한회사/g, '')
    .toLowerCase();
}

function normalizeProduct(name) {
  return (name || '').replace(/\s+/g, '').toLowerCase();
}

function pn(v) {
  if (typeof v === 'number') return v;
  return parseFloat((v || '').toString().replace(/[,\s]/g, '')) || 0;
}

function normalizeDate(raw) {
  if (!raw) return '';
  const s = String(raw).trim();

  // YYYYMMDD (20260304)
  const ymdNoSep = s.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (ymdNoSep) return `${ymdNoSep[2]}-${ymdNoSep[3]}`;

  // YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD
  const fullMatch = s.match(/(\d{4})[-./](\d{1,2})[-./](\d{1,2})/);
  if (fullMatch) return `${fullMatch[2].padStart(2, '0')}-${fullMatch[3].padStart(2, '0')}`;

  // MMDD (0304)
  const mmdd = s.match(/^(\d{2})(\d{2})$/);
  if (mmdd) return `${mmdd[1]}-${mmdd[2]}`;

  // MM/DD or MM-DD
  const mdMatch = s.match(/(\d{1,2})[-./](\d{1,2})/);
  if (mdMatch) return `${mdMatch[1].padStart(2, '0')}-${mdMatch[2].padStart(2, '0')}`;

  // D or DD (day only)
  const dayOnly = s.match(/^(\d{1,2})$/);
  if (dayOnly) return dayOnly[1].padStart(2, '0');

  return s;
}

function clientMatch(a, b) {
  const na = normalizeClient(a);
  const nb = normalizeClient(b);
  if (!na || !nb) return false;
  return na === nb || na.includes(nb) || nb.includes(na);
}

function dateMatch(d1, d2) {
  const n1 = normalizeDate(d1);
  const n2 = normalizeDate(d2);
  if (!n1 || !n2) return false;
  if (n1 === n2) return true;
  if (n1.endsWith(n2) || n2.endsWith(n1)) return true;
  return false;
}

function priceMatch(a, b, tolerance = 10) {
  return Math.abs(pn(a) - pn(b)) <= tolerance;
}

function makeKey(date, client, product, supply) {
  return `${normalizeDate(date)}|${normalizeClient(client)}|${normalizeProduct(product)}|${Math.round(pn(supply))}`;
}

function compareSalesData(sheetData, gyeongliData) {
  const results = {
    matched: [],
    sheetOnly: [],
    gyeongliOnly: [],
    mismatch: [],
    summary: {},
  };

  if (!sheetData || !gyeongliData) {
    logger.warn('대조 데이터가 부족합니다. 양쪽 데이터를 먼저 수집하세요.');
    return results;
  }

  const sheetSales = sheetData.filter(r => {
    return (r.salesClient || '').trim().length > 0;
  });

  // 경리나라 아이템을 거래처별로 그룹
  const gyeongliByClient = {};
  let gyeongliTotal = 0;
  for (const client of gyeongliData) {
    const nClient = normalizeClient(client.client);
    if (!gyeongliByClient[nClient]) {
      gyeongliByClient[nClient] = { name: client.client, items: [] };
    }
    for (const item of (client.items || [])) {
      gyeongliByClient[nClient].items.push({
        client: client.client,
        product: item.product || '',
        spec: item.spec || '',
        qty: pn(item.qty),
        unitPrice: pn(item.unitPrice),
        supplyAmount: pn(item.supplyAmount),
        vat: pn(item.vat),
        total: pn(item.total),
        deliveryDate: item.deliveryDate || '',
        note: item.note || '',
        _used: false,
      });
      gyeongliTotal++;
    }
  }

  logger.info(`매출 대조: Google Sheet ${sheetSales.length}건 (전체 ${sheetData.length}건) vs 경리나라 ${gyeongliTotal}건`);

  // 시트 레코드도 거래처별 그룹
  const sheetByClient = {};
  for (const s of sheetSales) {
    const nClient = normalizeClient(s.salesClient);
    if (!sheetByClient[nClient]) sheetByClient[nClient] = [];
    sheetByClient[nClient].push(s);
  }

  // 시트 거래처 → 경리나라 거래처 매핑
  const clientMapping = {};
  for (const sheetKey of Object.keys(sheetByClient)) {
    if (gyeongliByClient[sheetKey]) {
      clientMapping[sheetKey] = sheetKey;
    } else {
      for (const gKey of Object.keys(gyeongliByClient)) {
        if (sheetKey.includes(gKey) || gKey.includes(sheetKey)) {
          clientMapping[sheetKey] = gKey;
          break;
        }
      }
    }
  }

  logger.info(`거래처 매핑: ${Object.entries(clientMapping).map(([s, g]) => `${sheetByClient[s][0].salesClient} → ${gyeongliByClient[g].name}`).join(', ')}`);

  for (const [sheetClientKey, sheetItems] of Object.entries(sheetByClient)) {
    const gKey = clientMapping[sheetClientKey];
    if (!gKey || !gyeongliByClient[gKey]) {
      for (const s of sheetItems) {
        results.sheetOnly.push({
          client: s.salesClient || '',
          product: s.productName || '',
          spec: s.spec || '',
          qty: pn(s.qty),
          unitPrice: pn(s.unitPrice),
          supplyAmount: pn(s.supplyAmount),
          vat: pn(s.vat),
          total: pn(s.totalAmount),
          sheetDate: s.itemDate || s.tradeDate || '',
          _status: 'sheetOnly',
        });
      }
      continue;
    }

    const gItems = gyeongliByClient[gKey].items;
    const clientName = sheetItems[0].salesClient || gyeongliByClient[gKey].name;

    for (const s of sheetItems) {
      const sProduct = s.productName || '';
      const sDate = s.itemDate || s.tradeDate || '';
      const sQty = pn(s.qty);
      const sPrice = pn(s.unitPrice);
      const sSupply = pn(s.supplyAmount);
      const sVat = pn(s.vat);
      const sTotal = pn(s.totalAmount);
      const nProduct = normalizeProduct(sProduct);

      // 1단계: 거래처 내에서 품목+공급가액 일치
      let match = gItems.find(g =>
        !g._used && normalizeProduct(g.product) === nProduct && Math.abs(g.supplyAmount - sSupply) <= 1
      );

      // 2단계: 품목+단가 일치
      if (!match) {
        match = gItems.find(g =>
          !g._used && normalizeProduct(g.product) === nProduct && Math.abs(g.unitPrice - sPrice) <= 1
        );
      }

      // 3단계: 품목명만 일치
      if (!match) {
        match = gItems.find(g =>
          !g._used && normalizeProduct(g.product) === nProduct
        );
      }

      if (match) {
        match._used = true;

        const diffs = [];
        if (!dateMatch(sDate, match.deliveryDate)) {
          diffs.push(`날짜: 시트=${sDate || '없음'} / 경리=${match.deliveryDate || '없음'}`);
        }
        if (Math.abs(match.supplyAmount - sSupply) > 1) {
          diffs.push(`공급가: ${sSupply.toLocaleString()} → ${match.supplyAmount.toLocaleString()}`);
        }
        if (Math.abs(match.unitPrice - sPrice) > 1) {
          diffs.push(`단가: ${sPrice.toLocaleString()} → ${match.unitPrice.toLocaleString()}`);
        }
        if (Math.abs(match.vat - sVat) > 1) {
          diffs.push(`세액: ${sVat.toLocaleString()} → ${match.vat.toLocaleString()}`);
        }
        if (Math.abs(match.qty - sQty) > 0.01) {
          diffs.push(`수량: ${sQty} → ${match.qty}`);
        }

        const entry = {
          client: clientName,
          product: sProduct,
          spec: s.spec || match.spec || '',
          qty: sQty,
          sheetDate: sDate,
          gyeongliDate: match.deliveryDate,
          sheetPrice: sPrice,
          sheetSupply: sSupply,
          sheetVat: sVat,
          sheetTotal: sTotal,
          gyeongliPrice: match.unitPrice,
          gyeongliSupply: match.supplyAmount,
          gyeongliVat: match.vat,
          gyeongliTotal: match.total,
          note: match.note || '',
          diffs,
        };

        if (diffs.length > 0) {
          entry._status = 'mismatch';
          results.mismatch.push(entry);
        } else {
          entry._status = 'matched';
          results.matched.push(entry);
        }
      } else {
        results.sheetOnly.push({
          client: clientName,
          product: sProduct,
          spec: s.spec || '',
          qty: sQty,
          unitPrice: sPrice,
          supplyAmount: sSupply,
          vat: sVat,
          total: sTotal,
          sheetDate: sDate,
          _status: 'sheetOnly',
        });
      }
    }
  }

  // 경리나라에만 있는 항목
  for (const group of Object.values(gyeongliByClient)) {
    for (const g of group.items) {
      if (!g._used) {
        results.gyeongliOnly.push({
          client: g.client,
          product: g.product,
          spec: g.spec,
          qty: g.qty,
          unitPrice: g.unitPrice,
          supplyAmount: g.supplyAmount,
          vat: g.vat,
          total: g.total,
          gyeongliDate: g.deliveryDate,
          note: g.note,
          _status: 'gyeongliOnly',
        });
      }
    }
  }

  results.summary = {
    sheetCount: sheetSales.length,
    sheetTotalCount: sheetData.length,
    gyeongliCount: gyeongliTotal,
    matchedCount: results.matched.length,
    mismatchCount: results.mismatch.length,
    sheetOnlyCount: results.sheetOnly.length,
    gyeongliOnlyCount: results.gyeongliOnly.length,
  };

  logger.info(`대조 결과: 일치=${results.matched.length}, 불일치=${results.mismatch.length}, 시트만=${results.sheetOnly.length}, 경리만=${results.gyeongliOnly.length}`);

  return results;
}

function matchDeposits(salesData, depositData) {
  const results = {
    matched: [],
    unmatched: [],
    suggestions: [],
  };

  if (!salesData || !depositData) {
    logger.warn('매칭 데이터가 부족합니다.');
    return results;
  }

  const clientTotals = {};
  for (const c of salesData) {
    const name = normalizeClient(c.client);
    if (!clientTotals[name]) {
      clientTotals[name] = { client: c.client, total: 0, prevBalance: c.prevBalance || 0 };
    }
    clientTotals[name].total += (c.items || []).reduce((s, i) => s + (i.total || 0), 0);
  }

  for (const deposit of depositData) {
    const depClient = normalizeClient(deposit.client);
    const matchedClient = Object.keys(clientTotals).find(k =>
      k === depClient || k.includes(depClient) || depClient.includes(k)
    );

    if (matchedClient) {
      results.matched.push({
        ...deposit,
        matchedTo: clientTotals[matchedClient].client,
        clientTotal: clientTotals[matchedClient].total,
      });
    } else {
      results.unmatched.push({
        ...deposit,
        possibleMatches: Object.values(clientTotals)
          .filter(c => normalizeClient(c.client).substring(0, 2) === depClient.substring(0, 2))
          .map(c => c.client),
      });
    }
  }

  results.suggestions = results.unmatched.map(d => ({
    deposit: d,
    suggestion: d.possibleMatches.length > 0
      ? `"${d.client}" → 후보: ${d.possibleMatches.join(', ')}`
      : `"${d.client}" ${(d.amount || 0).toLocaleString()}원 - 매칭 거래처 없음`,
  }));

  return results;
}

module.exports = { compareSalesData, matchDeposits };
