/* 納品価格帳 — 弥生販売の台帳と個別価格表を読み込み、端末内で動く配達用アプリ */
'use strict';
const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));
const yen = n => (n == null || n === '') ? '' : Math.round(n).toLocaleString('ja-JP');
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } }
};
try { navigator.storage && navigator.storage.persist && navigator.storage.persist(); } catch (e) {}

let DB = store.get('nh.data', null);          // {customers, products, indiv, importedAt, files}
let SET = store.get('nh.settings', { jodaiKeywords: ['Aコープ', 'ロッキー'] });
let SLIPS = store.get('nh.slips', []);        // saved delivery records
let DRAFT = store.get('nh.draft', {});        // custCode -> {qty:{code:n}, jodai:bool}
let UI = { area: store.get('nh.area', 'すべて'), q: '', sel: null, view: 'cust', showAll: false, overlay: null, confirm: null, msg: null };

const TIER_IDX = { '上代': 0, '売上単価１': 1, '売上単価２': 2, '売上単価３': 3, '売上単価４': 4, '売上単価1': 1, '売上単価2': 2, '売上単価3': 3, '売上単価4': 4 };
const TIER_LABEL = ['上代（小売）', '単価１（一般）', '単価２（カナカン等）', '単価３（ロッキー等）', '単価４（道の駅）'];
const AREA_ORDER = ['富来', '門前', '志賀', '七尾', '羽咋', '金沢', '県外'];

/* ---------- 文字の正規化（ひらがな・カタカナ・半角を同一視） ---------- */
function norm(s) {
  s = String(s ?? '').normalize('NFKC').toLowerCase().replace(/[\s　・()（）]/g, '');
  return s.replace(/[ァ-ヶ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0x60));
}

/* ---------- 端数処理 ---------- */
function roundBy(x, mode) {
  if (mode === '四捨五入') return Math.round(x);
  if (mode === '切り上げ') return Math.ceil(x - 1e-9);
  return Math.floor(x + 1e-9);
}

/* ---------- 読み込み ---------- */
function rowsOf(ws) { return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true }); }
function findHeader(rows, key) {
  for (let i = 0; i < Math.min(rows.length, 20); i++) if (rows[i] && rows[i].some(v => String(v ?? '').trim() === key)) return i;
  return -1;
}
function colOf(hdr, name) { return hdr.findIndex(v => String(v ?? '').trim() === name); }
const str = v => v == null ? '' : String(v).trim();
const num = v => { const n = Number(String(v ?? '').replace(/,/g, '')); return isFinite(n) ? n : 0; };

function parseCustomers(rows) {
  const h = findHeader(rows, '単価種類'); const hd = rows[h];
  const c = n => colOf(hd, n);
  const cA = c('分類１');
  const out = [];
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const code = str(r[c('コード')]); const name = str(r[c('名称')]);
    if (!code || !name) continue;
    out.push({
      code, name, kana: str(r[c('フリガナ')]), short: str(r[c('略称')]),
      addr: [str(r[c('住所１')]), str(r[c('住所２')])].filter(Boolean).join(' '),
      tel: str(r[c('TEL')]), deal: str(r[c('取引区分')]), tier: str(r[c('単価種類')]),
      rate: num(r[c('掛率')]) || 100, area: cA >= 0 ? str(r[cA + 1]) : '',
      memo: str(r[c('メモ欄')]), amtRound: str(r[c('金額端数処理')]), taxRound: str(r[c('税端数処理')]),
      show: str(r[c('参照表示')]) !== '表示しない'
    });
  }
  return out;
}
function parseProducts(rows) {
  const h = findHeader(rows, '税抜上代'); const hd = rows[h];
  const c = n => colOf(hd, n);
  const pc = ['税抜上代', '税抜売上単価１', '税抜売上単価２', '税抜売上単価３', '税抜売上単価４'].map(c);
  const out = [];
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const code = str(r[c('コード')]); const name = str(r[c('名称')]);
    if (!code || !name) continue;
    const tax = str(r[c('課税区分')]);
    out.push({
      code, name, unit: str(r[c('単位')]), irisu: num(r[c('入数')]),
      p: pc.map(ix => ix >= 0 ? num(r[ix]) : 0),
      taxRate: tax.includes('軽') ? 8 : (tax.startsWith('課税') ? 10 : 0),
      show: str(r[c('参照表示')]) !== '表示しない' && !/^[×✕]/.test(name)
    });
  }
  return out;
}
function parseIndiv(rows) {
  const h = findHeader(rows, '得意先コード'); const hd = rows[h];
  const cc = colOf(hd, '得意先コード'), pc = colOf(hd, '商品コード');
  let vc = hd.findIndex(v => String(v ?? '').startsWith('個別価格'));
  const mc = hd.findIndex(v => String(v ?? '').includes('メモ'));
  const out = {}; let n = 0;
  for (let i = h + 1; i < rows.length; i++) {
    const r = rows[i]; if (!r) continue;
    const cu = str(r[cc]), pr = str(r[pc]), v = r[vc];
    if (!cu || !pr || v == null || v === '' || !isFinite(Number(v))) continue;
    (out[cu] = out[cu] || {})[pr] = { price: Number(v), note: mc >= 0 ? str(r[mc]) : '' };
    n++;
  }
  return { map: out, count: n };
}

async function importFiles(files) {
  const found = {}; const names = [];
  for (const f of files) {
    const buf = await f.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    let used = false;
    for (const sn of wb.SheetNames) {
      const rows = rowsOf(wb.Sheets[sn]);
      if (findHeader(rows, '得意先コード') >= 0 && !found.indiv) { found.indiv = parseIndiv(rows); used = true; }
      else if (findHeader(rows, '単価種類') >= 0 && findHeader(rows, '取引区分') >= 0 && !found.customers) { found.customers = parseCustomers(rows); used = true; }
      else if (findHeader(rows, '税抜上代') >= 0 && !found.products) { found.products = parseProducts(rows); used = true; }
    }
    names.push(f.name + (used ? '' : '（読めませんでした）'));
  }
  if (!found.customers && !found.products && !found.indiv) throw new Error('得意先台帳・商品台帳・個別価格表のどれとも判別できませんでした。弥生販売から書き出したExcelファイルか確認してください。');
  const next = Object.assign({ customers: [], products: [], indiv: {}, indivCount: 0 }, DB || {});
  if (found.customers) next.customers = found.customers;
  if (found.products) next.products = found.products;
  if (found.indiv) { next.indiv = found.indiv.map; next.indivCount = found.indiv.count; }
  next.importedAt = new Date().toISOString();
  next.files = names;
  if (!store.set('nh.data', next)) throw new Error('端末に保存できませんでした。ブラウザの保存領域を確認してください。');
  DB = next;
  return found;
}

/* ---------- 価格計算 ---------- */
const custByCode = code => DB && DB.customers.find(c => c.code === code);
const prodByCode = code => DB && DB.products.find(p => p.code === code);
function tierIdx(c) { return TIER_IDX[c.tier] ?? 1; }
function basePrice(c, p) {
  let v = p.p[tierIdx(c)];
  if (!v) return null;
  if (c.rate && c.rate !== 100) v = roundBy(v * c.rate / 100, c.amtRound);
  return v;
}
function priceFor(c, p) {
  const k = DB.indiv[c.code] && DB.indiv[c.code][p.code];
  if (k) return { v: k.price, indiv: true, note: k.note };
  const v = basePrice(c, p);
  return v == null ? null : { v, indiv: false };
}
function isJodaiCust(c) { return (SET.jodaiKeywords || []).some(k => k && norm(c.name).includes(norm(k))); }
function draftOf(c) {
  if (!DRAFT[c.code]) DRAFT[c.code] = { qty: {}, jodai: isJodaiCust(c) };
  return DRAFT[c.code];
}
function saveDraft() { for (const k in DRAFT) if (!Object.keys(DRAFT[k].qty).length && DRAFT[k].jodai === isJodaiCust(custByCode(k) || { name: '' })) delete DRAFT[k]; store.set('nh.draft', DRAFT); }

function buildSlip(c, d) {
  const lines = [];
  for (const p of DB.products) {
    const q = d.qty[p.code]; if (!q) continue;
    const pr = priceFor(c, p);
    const unit = pr ? pr.v : 0;
    const jodai = p.p[0] || 0;
    lines.push({ code: p.code, name: p.name, unitLabel: p.unit, qty: q, price: unit, amount: roundBy(unit * q, c.amtRound),
      indiv: pr ? pr.indiv : false, jprice: jodai, jamount: jodai * q, taxRate: p.taxRate, noPrice: !pr });
  }
  const sub = lines.reduce((a, l) => a + l.amount, 0);
  const byRate = {};
  lines.forEach(l => { if (l.taxRate) byRate[l.taxRate] = (byRate[l.taxRate] || 0) + l.amount; });
  const taxes = Object.keys(byRate).sort().map(r => ({ rate: +r, base: byRate[r], tax: roundBy(byRate[r] * r / 100, c.taxRound) }));
  const tax = taxes.reduce((a, t) => a + t.tax, 0);
  const jsub = lines.reduce((a, l) => a + l.jamount, 0);
  const count = lines.reduce((a, l) => a + l.qty, 0);
  return { lines, sub, taxes, tax, total: sub + tax, jsub, count, jodai: d.jodai };
}

/* ---------- 一覧 ---------- */
function activeCustomers() { return DB ? DB.customers.filter(c => c.show) : []; }
function areaOf(c) { return c.area || '未分類'; }
function areas() {
  const set = new Set(activeCustomers().map(areaOf));
  const ord = a => { const i = AREA_ORDER.indexOf(a); return i < 0 ? (a === '未分類' ? 99 : 50) : i; };
  return ['すべて', ...[...set].sort((a, b) => ord(a) - ord(b))];
}
function renderChips() {
  const as = areas(); if (!as.includes(UI.area)) UI.area = 'すべて';
  $('#chips').innerHTML = as.map(a => `<button class="chip" data-a="${esc(a)}" aria-pressed="${a === UI.area}">${esc(a)}</button>`).join('');
}
function renderList() {
  if (!DB) { $('#list').innerHTML = '<div class="empty">データがまだありません。「データ更新」から読み込んでください。</div>'; return; }
  const q = norm(UI.q);
  let rows = activeCustomers().filter(c => (UI.area === 'すべて' || areaOf(c) === UI.area) &&
    (!q || norm(c.name).includes(q) || norm(c.kana).includes(q) || norm(c.short).includes(q) || c.code.includes(q)));
  const ord = a => { const i = AREA_ORDER.indexOf(a); return i < 0 ? (a === '未分類' ? 99 : 50) : i; };
  rows.sort((a, b) => ord(areaOf(a)) - ord(areaOf(b)) || norm(a.kana || a.name).localeCompare(norm(b.kana || b.name), 'ja'));
  let html = '', cur = null;
  for (const c of rows) {
    const a = areaOf(c);
    if (a !== cur) { cur = a; html += `<div class="grp">${esc(a)}</div>`; }
    const ti = tierIdx(c);
    const hasInd = DB.indiv[c.code] && Object.keys(DB.indiv[c.code]).length;
    const dr = DRAFT[c.code] && Object.keys(DRAFT[c.code].qty).length;
    html += `<button class="row" data-c="${esc(c.code)}" aria-current="${UI.view === 'cust' && UI.sel === c.code}">
      <div class="grow"><div class="nm">${esc(c.name)}</div><div class="sub">${esc(c.code)}${c.addr ? '　' + esc(c.addr.replace(/^石川県/, '')) : ''}</div></div>
      ${dr ? '<span class="tag draft">入力中</span>' : ''}
      ${c.deal.includes('現金') ? '<span class="tag cash">現金</span>' : ''}
      ${hasInd ? '<span class="tag gold">個別</span>' : ''}
      <span class="tag ${ti === 1 && c.rate === 100 ? '' : 'acc'}">${esc(['上代', '単価1', '単価2', '単価3', '単価4'][ti])}${c.rate !== 100 ? ' ' + c.rate + '%' : ''}</span></button>`;
  }
  $('#list').innerHTML = html || `<div class="empty">当てはまる得意先がありません。地域を「すべて」にすると見つかる場合があります。</div>`;
}

/* ---------- 商品の分類 ---------- */
function groupOf(p) {
  const n = Number(p.code);
  if (n < 1000) return [9, '容器・空瓶'];
  const d = String(p.code)[0];
  if (/セット|ギフト|詰合/.test(p.name)) return [7, 'ギフト・セット'];
  if (/入り|×\d+本/.test(p.name)) return [6, 'ケース'];
  return ({ '1': [1, '醤油'], '2': [2, '味醤油'], '3': [3, '酢'], '4': [4, 'みりん・その他'], '5': [5, '味噌'], '8': [7, 'ギフト・セット'], '9': [8, '送料・資材'] })[d] || [8, 'その他'];
}

/* ---------- 詳細 ---------- */
function renderDetail() {
  const el = $('#detail');
  if (!DB) { el.innerHTML = welcomeHTML(); return; }
  if (UI.view === 'rec') { el.innerHTML = recordsHTML(); return; }
  const c = custByCode(UI.sel);
  if (!c) { el.innerHTML = `<div class="wrap"><p class="info">左の一覧から得意先を選ぶと、その店の納品価格が表示されます。</p>${dataSummaryHTML()}</div>`; return; }
  const d = draftOf(c);
  const ti = tierIdx(c);
  const indiv = DB.indiv[c.code] || {};
  const meta = [c.code, c.area, TIER_LABEL[ti] + (c.rate !== 100 ? `・掛率${c.rate}%` : '')].filter(Boolean);
  let h = `<div class="wrap">
    <button class="back" data-act="back">← 得意先一覧</button>
    <div><h1 class="cname">${esc(c.name)}</h1>
      <div class="meta">${meta.map(m => `<span>${esc(m)}</span>`).join('')}${c.deal ? `<span class="${c.deal.includes('現金') ? 'cash' : ''}">${esc(c.deal)}</span>` : ''}</div>
      ${c.addr || c.tel ? `<div class="info" style="margin-top:6px">${esc(c.addr)}${c.tel ? '　TEL ' + esc(c.tel) : ''}</div>` : ''}
    </div>`;
  if (c.memo) h += `<div class="note">${esc(c.memo)}</div>`;
  const indKeys = Object.keys(indiv);
  if (indKeys.length) {
    h += `<div class="box"><h2>この店だけの価格（税抜）</h2>` + indKeys.map(k => {
      const p = prodByCode(k);
      return `<div class="it"><span>${esc(p ? p.name : '商品コード ' + k)}</span><span class="pr">${yen(indiv[k].price)}<small style="font-size:13px;font-weight:400">円</small></span></div>`;
    }).join('') + `</div>`;
  }
  h += `<div class="tools">
    <div class="seg" role="group" aria-label="表示する商品">
      <button data-act="filter" data-v="0" aria-pressed="${!UI.showAll}">この店の価格がある商品</button>
      <button data-act="filter" data-v="1" aria-pressed="${UI.showAll}">すべての商品</button>
    </div>
    <button class="tgl" data-act="jodai" aria-pressed="${d.jodai}">上代（小売）も計算 ${d.jodai ? 'ON' : 'OFF'}</button>
  </div>`;
  const prods = DB.products.filter(p => p.show).map(p => ({ p, pr: priceFor(c, p), g: groupOf(p) }))
    .filter(x => UI.showAll || x.pr || d.qty[x.p.code])
    .sort((a, b) => a.g[0] - b.g[0] || Number(a.p.code) - Number(b.p.code));
  let rows = '', cat = null;
  for (const x of prods) {
    if (x.g[1] !== cat) { cat = x.g[1]; rows += `<div class="pcat">${esc(cat)}</div>`; }
    const q = d.qty[x.p.code] || '';
    rows += `<div class="prow ${q ? 'has' : ''}" data-p="${esc(x.p.code)}">
      <div class="pn"><span class="k">${esc(x.p.name)}</span><span class="s">${esc(x.p.code)}${x.p.unit ? '・' + esc(x.p.unit) : ''}${x.p.irisu ? '・入数' + x.p.irisu : ''}</span></div>
      <div class="pp">${x.pr ? `<span class="v">${x.pr.indiv ? '<span class="ind">個別</span>' : ''}${yen(x.pr.v)}</span>` : '<span class="v na">価格なし</span>'}
        ${d.jodai ? `<span class="j">上代 ${x.p.p[0] ? yen(x.p.p[0]) : '－'}</span>` : ''}</div>
      <div class="step"><button data-act="dec" aria-label="1つ減らす">−</button><input type="number" inputmode="numeric" min="0" value="${q}" placeholder="0" aria-label="${esc(x.p.name)} の数量" data-act="qty"><button class="plus" data-act="inc" aria-label="1つ増やす">＋</button></div>
    </div>`;
  }
  h += `<div class="plist">${rows || '<div class="empty">この店に設定された価格がありません。「すべての商品」を押してください。</div>'}</div>
    <p class="small">価格はすべて税抜。弥生販売の単価種類${c.rate !== 100 ? '・掛率' : ''}と個別価格表から計算しています。</p></div>`;
  const s = buildSlip(c, d);
  h += `<div class="slipbar"><div class="sum">
      <span><b>${s.lines.length}</b> 品・計 <b>${s.count}</b> 点</span>
      <span><small>金額（税抜）</small><b>${yen(s.sub)}</b> 円</span>
      ${d.jodai ? `<span><small>上代金額</small><b>${yen(s.jsub)}</b> 円</span>` : ''}
    </div><button data-act="slip" ${s.lines.length ? '' : 'disabled'}>納品書を見る</button></div>`;
  const st = el.scrollTop; el.innerHTML = h; el.scrollTop = st;
}

function slipTableHTML(s) {
  const j = s.jodai;
  return `<div class="tblwrap"><table><thead><tr><th>商品</th><th class="n">数量</th><th class="n">単価</th><th class="n">金額</th>${j ? '<th class="n">上代</th><th class="n">上代金額</th>' : ''}</tr></thead><tbody>
    ${s.lines.map(l => `<tr><td class="nm">${esc(l.name)}${l.indiv ? ' <span class="tag gold">個別</span>' : ''}${l.noPrice ? ' <span class="tag cash">価格なし</span>' : ''}</td><td class="n">${l.qty}${esc(l.unitLabel)}</td><td class="n">${yen(l.price)}</td><td class="n">${yen(l.amount)}</td>${j ? `<td class="n">${l.jprice ? yen(l.jprice) : '－'}</td><td class="n">${l.jprice ? yen(l.jamount) : '－'}</td>` : ''}</tr>`).join('')}
  </tbody><tfoot>
    <tr><td>合計（税抜）</td><td class="n">${s.count}</td><td></td><td class="n">${yen(s.sub)}</td>${j ? `<td></td><td class="n">${yen(s.jsub)}</td>` : ''}</tr>
    ${s.taxes.map(t => `<tr><td colspan="3" class="small">消費税 ${t.rate}%（対象 ${yen(t.base)}円・参考）</td><td class="n">${yen(t.tax)}</td>${j ? '<td></td><td></td>' : ''}</tr>`).join('')}
    <tr><td colspan="3">税込合計（参考）</td><td class="n">${yen(s.total)}</td>${j ? '<td></td><td></td>' : ''}</tr>
  </tfoot></table></div>`;
}

function overlayHTML() {
  const o = UI.overlay; if (!o) return '';
  if (o.type === 'slip') {
    const c = custByCode(UI.sel); const s = buildSlip(c, draftOf(c));
    const conf = UI.confirm === 'clear' ? `<div class="confirm">入力した数量をすべて消しますか？<span class="btns"><button class="btn" data-act="cancelc">やめる</button><button class="btn dan" data-act="doclear">消す</button></span></div>` : '';
    return `<div class="overlay" data-act="close"><div class="sheet" role="dialog" aria-modal="true" aria-label="納品書">
      <h2>${esc(c.name)} 様　納品内容</h2>
      <div class="small">${new Date().toLocaleDateString('ja-JP')}　${esc(TIER_LABEL[tierIdx(c)])}${c.rate !== 100 ? '・掛率' + c.rate + '%' : ''}　${esc(c.deal)}</div>
      ${slipTableHTML(s)}${conf}
      <div class="btns"><button class="btn dan" data-act="clear">数量をクリア</button><button class="btn" data-act="close">閉じる</button><button class="btn pri" data-act="save">納品記録に保存</button></div>
    </div></div>`;
  }
  if (o.type === 'rec') {
    const r = SLIPS.find(x => x.id === o.id); if (!r) return '';
    const conf = UI.confirm === 'del' ? `<div class="confirm">この記録を削除しますか？<span class="btns"><button class="btn" data-act="cancelc">やめる</button><button class="btn dan" data-act="dodel">削除する</button></span></div>` : '';
    return `<div class="overlay" data-act="close"><div class="sheet" role="dialog" aria-modal="true" aria-label="納品記録">
      <h2>${esc(r.custName)} 様</h2><div class="small">${new Date(r.at).toLocaleString('ja-JP')}　得意先コード ${esc(r.custCode)}</div>
      ${slipTableHTML(r.slip)}${conf}
      <div class="btns"><button class="btn dan" data-act="del">削除</button><button class="btn" data-act="close">閉じる</button></div>
    </div></div>`;
  }
  if (o.type === 'data') {
    return `<div class="overlay" data-act="close"><div class="sheet" role="dialog" aria-modal="true" aria-label="データ更新">
      <h2>データ更新</h2>
      ${UI.msg ? `<div class="msg ${UI.msg.err ? 'err' : ''}">${esc(UI.msg.text)}</div>` : ''}
      ${dataSummaryHTML()}
      <div class="small">弥生販売から書き出した「得意先台帳」「商品台帳」と「個別価格の得意先一覧」のExcelを選んでください。まとめて選んでも、1つずつでもかまいません。選んだファイルの分だけ置き換わります。</div>
      <label class="filebtn">ファイルを選んで読み込む<input type="file" id="files" multiple accept=".xlsx,.xls,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"></label>
      <div class="field"><label for="kw"><b>上代（小売）も計算する得意先</b>（名前に含まれる言葉。読点で区切る）</label>
        <input id="kw" value="${esc((SET.jodaiKeywords || []).join('、'))}"></div>
      <div class="btns"><button class="btn" data-act="close">閉じる</button><button class="btn pri" data-act="savekw">設定を保存</button></div>
    </div></div>`;
  }
  return '';
}
function dataSummaryHTML() {
  if (!DB) return '';
  const at = DB.importedAt ? new Date(DB.importedAt).toLocaleString('ja-JP') : '－';
  return `<div class="stat">
    <div><span class="small">得意先（取引中）</span><b>${activeCustomers().length}</b></div>
    <div><span class="small">商品</span><b>${DB.products.filter(p => p.show).length}</b></div>
    <div><span class="small">個別価格</span><b>${DB.indivCount || 0}</b></div>
    <div><span class="small">最終読み込み</span><b style="font-size:15px">${esc(at)}</b></div></div>`;
}
function welcomeHTML() {
  return `<div class="welcome"><h1>はじめに、データを読み込みます</h1>
    <p>会社のパソコンで弥生販売から次のファイルを書き出し、タブレットから開ける場所（Googleドライブなど）に保存してください。</p>
    <ol><li>得意先台帳（Excel）</li><li>商品台帳（Excel）</li><li>個別価格の得意先一覧（Excel）</li></ol>
    <p>読み込んだデータはこのタブレットの中だけに保存されます。インターネットがなくても使えます。</p>
    <div><button class="filebtn" data-act="data">データを読み込む</button></div></div>`;
}
function recordsHTML() {
  const list = SLIPS.slice().sort((a, b) => b.at.localeCompare(a.at));
  let h = `<div class="wrap"><button class="back" data-act="back">← 得意先一覧</button><h1 class="cname">納品記録</h1>
    <p class="info">「納品記録に保存」した納品内容です。このタブレットの中に保存されています。</p>`;
  let day = null;
  for (const r of list) {
    const d = new Date(r.at).toLocaleDateString('ja-JP', { year: 'numeric', month: 'long', day: 'numeric', weekday: 'short' });
    if (d !== day) { day = d; h += `<h2 style="font-size:15px;color:var(--muted);margin:8px 0 0">${esc(d)}</h2>`; }
    h += `<button class="rec" data-rec="${esc(r.id)}"><span><b>${esc(r.custName)}</b><br><span class="small">${new Date(r.at).toLocaleTimeString('ja-JP', { hour: '2-digit', minute: '2-digit' })}・${r.slip.lines.length}品 ${r.slip.count}点${r.slip.jodai ? '・上代 ' + yen(r.slip.jsub) + '円' : ''}</span></span><span class="amt">${yen(r.slip.sub)}円</span></button>`;
  }
  if (!list.length) h += '<p class="empty">まだ記録はありません。</p>';
  return h + '</div>';
}

function render() {
  renderChips(); renderList(); renderDetail();
  $('#ov').innerHTML = overlayHTML();
  $('#recBtn').setAttribute('aria-pressed', UI.view === 'rec');
  $('#qclr').hidden = !UI.q;
}

/* ---------- 操作 ---------- */
function setQty(code, v) {
  const c = custByCode(UI.sel); const d = draftOf(c);
  v = Math.max(0, Math.floor(Number(v) || 0));
  if (v) d.qty[code] = v; else delete d.qty[code];
  saveDraft(); renderDetail(); renderList();
}
$('#chips').addEventListener('click', e => { const b = e.target.closest('.chip'); if (!b) return; UI.area = b.dataset.a; store.set('nh.area', UI.area); renderChips(); renderList(); });
$('#list').addEventListener('click', e => {
  const b = e.target.closest('.row'); if (!b) return;
  UI.sel = b.dataset.c; UI.view = 'cust'; UI.showAll = false;
  document.body.classList.add('show-detail'); render(); $('#detail').scrollTop = 0;
});
$('#detail').addEventListener('click', e => {
  const t = e.target.closest('[data-act],[data-rec]'); if (!t) return;
  if (t.dataset.rec) { UI.overlay = { type: 'rec', id: t.dataset.rec }; UI.confirm = null; render(); return; }
  const a = t.dataset.act; const row = t.closest('.prow');
  if (a === 'back') { document.body.classList.remove('show-detail'); if (UI.view === 'rec') UI.view = 'cust'; render(); }
  else if (a === 'filter') { UI.showAll = t.dataset.v === '1'; renderDetail(); }
  else if (a === 'jodai') { const c = custByCode(UI.sel); const d = draftOf(c); d.jodai = !d.jodai; saveDraft(); renderDetail(); }
  else if (a === 'inc' || a === 'dec') { const c = custByCode(UI.sel); const cur = draftOf(c).qty[row.dataset.p] || 0; setQty(row.dataset.p, cur + (a === 'inc' ? 1 : -1)); }
  else if (a === 'slip') { UI.overlay = { type: 'slip' }; UI.confirm = null; render(); }
  else if (a === 'data') { openData(); }
});
$('#detail').addEventListener('change', e => { if (e.target.dataset.act === 'qty') setQty(e.target.closest('.prow').dataset.p, e.target.value); });
$('#ov').addEventListener('click', async e => {
  const t = e.target.closest('[data-act]'); if (!t) return;
  const a = t.dataset.act;
  if (a === 'close' && t !== e.target) return; // backdrop only when clicked directly
  if (a === 'close') { UI.overlay = null; UI.confirm = null; UI.msg = null; render(); }
  else if (a === 'clear') { UI.confirm = 'clear'; $('#ov').innerHTML = overlayHTML(); }
  else if (a === 'cancelc') { UI.confirm = null; $('#ov').innerHTML = overlayHTML(); }
  else if (a === 'doclear') { const c = custByCode(UI.sel); draftOf(c).qty = {}; saveDraft(); UI.overlay = null; UI.confirm = null; render(); }
  else if (a === 'save') {
    const c = custByCode(UI.sel); const d = draftOf(c); const slip = buildSlip(c, d);
    SLIPS.push({ id: Date.now().toString(36), at: new Date().toISOString(), custCode: c.code, custName: c.name, slip });
    store.set('nh.slips', SLIPS); d.qty = {}; saveDraft(); UI.overlay = null; render();
  }
  else if (a === 'del') { UI.confirm = 'del'; $('#ov').innerHTML = overlayHTML(); }
  else if (a === 'dodel') { SLIPS = SLIPS.filter(x => x.id !== UI.overlay.id); store.set('nh.slips', SLIPS); UI.overlay = null; UI.confirm = null; render(); }
  else if (a === 'savekw') {
    SET.jodaiKeywords = $('#kw').value.split(/[、,，\s]+/).map(s => s.trim()).filter(Boolean);
    store.set('nh.settings', SET); UI.msg = { text: '設定を保存しました。新しく開いた得意先から反映されます。' }; render();
  }
});
$('#ov').addEventListener('change', async e => {
  if (e.target.id !== 'files' || !e.target.files.length) return;
  try {
    const f = await importFiles([...e.target.files]);
    const parts = [];
    if (f.customers) parts.push(`得意先 ${f.customers.filter(c => c.show).length}件`);
    if (f.products) parts.push(`商品 ${f.products.filter(p => p.show).length}件`);
    if (f.indiv) parts.push(`個別価格 ${f.indiv.count}件`);
    UI.msg = { text: '読み込みました：' + parts.join('、') };
  } catch (err) { UI.msg = { err: true, text: err.message || String(err) }; }
  render();
});
function openData() { UI.overlay = { type: 'data' }; UI.msg = null; render(); }
$('#dataBtn').addEventListener('click', openData);
$('#recBtn').addEventListener('click', () => { UI.view = UI.view === 'rec' ? 'cust' : 'rec'; if (UI.view === 'rec') document.body.classList.add('show-detail'); render(); $('#detail').scrollTop = 0; });
$('#q').addEventListener('input', e => { UI.q = e.target.value; $('#qclr').hidden = !UI.q; renderList(); });
$('#qclr').addEventListener('click', () => { UI.q = ''; $('#q').value = ''; $('#qclr').hidden = true; renderList(); $('#q').focus(); });

render();
if ('serviceWorker' in navigator && location.protocol.startsWith('http')) {
  navigator.serviceWorker.register('sw.js').catch(() => {});
}
