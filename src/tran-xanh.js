/* Tab TRẬN XANH — KH OFF mua Chanté Tết / Tea365 Tết.
   Nguồn: window.TRAN_XANH (sheet "Trận Xanh" của file Tracking KPI) + T11_DATA (MCP, Đại Sứ, VIP) + KPI_MONTHS (CT/TH KPI Trận Xanh).
   Cài vào thanh menu giống tab BÁO CÁO NGÀY. */
(function(){
'use strict';
if (window.__txInstalled) return; window.__txInstalled = true;

var CSS = '#tran-xanh{--tx-g:#0f9d58;--tx-y:#d99a06;--tx-r:#c62828;--tx-o:#e05606;--tx-ink:#1d1b19;--tx-mut:#6b6560;--tx-line:#e3ddd8;--tx-card:#fff;max-width:1560px;margin:0 auto;padding:4px 0 90px;color:var(--tx-ink)}'+
'#tran-xanh *{box-sizing:border-box}'+
'#tran-xanh .tx-hero{background:linear-gradient(120deg,#0b6b3a,#13a35b 60%,#5fd08f);color:#fff;border-radius:20px;padding:22px 24px;display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;align-items:flex-end;box-shadow:0 10px 28px rgba(11,107,58,.25)}'+
'#tran-xanh .tx-hero small{letter-spacing:.12em;font-weight:700;opacity:.85;font-size:11px}#tran-xanh .tx-hero h2{margin:6px 0 4px;font-size:26px;line-height:1.2}#tran-xanh .tx-hero p{margin:0;opacity:.92;font-size:13px;max-width:760px;line-height:1.5}'+
'#tran-xanh .tx-time{text-align:right}#tran-xanh .tx-time b{display:block;font-size:20px}#tran-xanh .tx-time span{font-size:12px;opacity:.9}'+
'#tran-xanh .tx-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin:16px 0}'+
'#tran-xanh .tx-card{background:var(--tx-card);border:1px solid var(--tx-line);border-radius:16px;padding:14px 16px;box-shadow:0 2px 8px rgba(0,0,0,.04)}'+
'#tran-xanh .tx-card span{display:block;font-size:11px;font-weight:700;color:var(--tx-mut);letter-spacing:.06em;text-transform:uppercase}#tran-xanh .tx-card strong{display:block;font-size:26px;margin:6px 0 2px}#tran-xanh .tx-card small{color:var(--tx-mut);font-size:12px}'+
'#tran-xanh .tx-card.good strong{color:var(--tx-g)}#tran-xanh .tx-card.watch strong{color:var(--tx-y)}#tran-xanh .tx-card.bad strong{color:var(--tx-r)}'+
'#tran-xanh .tx-bar{height:7px;border-radius:9px;background:#eee;margin-top:8px;overflow:hidden}#tran-xanh .tx-bar i{display:block;height:100%;border-radius:9px;background:var(--tx-g)}'+
'#tran-xanh .tx-panel{background:var(--tx-card);border:1px solid var(--tx-line);border-radius:18px;padding:16px;margin-bottom:16px}'+
'#tran-xanh .tx-head{display:flex;justify-content:space-between;align-items:flex-start;gap:10px;flex-wrap:wrap;margin-bottom:10px}#tran-xanh .tx-head h3{margin:0;font-size:17px}#tran-xanh .tx-head p{margin:3px 0 0;color:var(--tx-mut);font-size:12px}'+
'#tran-xanh .tx-pill{background:#e7f6ee;color:#0b6b3a;border-radius:999px;padding:5px 11px;font-size:12px;font-weight:700;white-space:nowrap}'+
'#tran-xanh .tx-insight{background:#f3fbf6;border:1px solid #cfeedd;border-radius:14px;padding:12px 14px;margin:0 0 16px;font-size:13px;line-height:1.6}#tran-xanh .tx-insight b{color:#0b6b3a}'+
'#tran-xanh .tx-wrap{overflow:auto;-webkit-overflow-scrolling:touch;border-radius:12px;border:1px solid var(--tx-line)}'+
'#tran-xanh table{border-collapse:collapse;width:100%;font-size:13px;min-width:720px}#tran-xanh th{position:sticky;top:0;background:#f6f3f0;font-size:11px;text-transform:uppercase;letter-spacing:.03em;color:#4c4743;text-align:right;padding:9px 8px;border-bottom:1px solid var(--tx-line);white-space:nowrap}'+
'#tran-xanh td{padding:8px;border-bottom:1px solid #f0ebe7;text-align:right;white-space:nowrap}#tran-xanh th:first-child,#tran-xanh td:first-child,#tran-xanh th.l,#tran-xanh td.l{text-align:left}'+
'#tran-xanh td:first-child{position:sticky;left:0;background:var(--tx-card);font-weight:600}#tran-xanh tr.tot td{background:#fbf7f3;font-weight:800}'+
'#tran-xanh .c-good{color:var(--tx-g);font-weight:800}#tran-xanh .c-watch{color:var(--tx-y);font-weight:800}#tran-xanh .c-bad{color:var(--tx-r);font-weight:800}'+
'#tran-xanh .bg-good{background:#e4f6ea!important}#tran-xanh .bg-watch{background:#fff4d6!important}#tran-xanh .bg-bad{background:#fde6e6!important}'+
'#tran-xanh .tx-tools{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin-bottom:10px}'+
'#tran-xanh .tx-seg{display:inline-flex;background:#f1ece8;border-radius:999px;padding:3px;gap:2px;flex-wrap:wrap}#tran-xanh .tx-seg button{border:0;background:transparent;border-radius:999px;padding:7px 12px;font:600 13px system-ui,Segoe UI,sans-serif;color:#4c4743;cursor:pointer}#tran-xanh .tx-seg button.on{background:#fff;color:#0b6b3a;box-shadow:0 1px 4px rgba(0,0,0,.12)}'+
'#tran-xanh input.tx-q{flex:1;min-width:160px;border:1px solid var(--tx-line);border-radius:999px;padding:9px 14px;font-size:14px}'+
'#tran-xanh .tx-btn{border:0;background:#0b6b3a;color:#fff;border-radius:999px;padding:9px 14px;font:600 13px system-ui,Segoe UI,sans-serif;cursor:pointer}'+
'#tran-xanh .tx-tag{display:inline-block;border-radius:999px;padding:2px 8px;font-size:11px;font-weight:700}#tran-xanh .t-ds{background:#fff0e6;color:#b5480a}#tran-xanh .t-vip{background:#eef0ff;color:#3446b8}#tran-xanh .t-le{background:#f1f1f1;color:#555}'+
'#tran-xanh .ok{color:var(--tx-g);font-weight:800}#tran-xanh .no{color:#c4bdb7}'+
'#tran-xanh .tx-note{font-size:12px;color:var(--tx-mut);line-height:1.6;margin-top:10px}'+
'#tran-xanh .tx-more{display:block;margin:10px auto 0;border:1px solid var(--tx-line);background:#fff;border-radius:999px;padding:8px 16px;cursor:pointer;font-weight:600}'+
'@media(max-width:720px){#tran-xanh .tx-hero{padding:16px}#tran-xanh .tx-hero h2{font-size:20px}#tran-xanh .tx-time{text-align:left}#tran-xanh .tx-cards{grid-template-columns:repeat(2,minmax(0,1fr))}#tran-xanh .tx-card strong{font-size:21px}#tran-xanh table{font-size:12px}}'+
'html.tt-dark #tran-xanh{--tx-ink:#eee;--tx-mut:#b7b0aa;--tx-line:#3a3633;--tx-card:#221f1d}html.tt-dark #tran-xanh th{background:#2b2724;color:#ddd}html.tt-dark #tran-xanh tr.tot td{background:#2b2724}html.tt-dark #tran-xanh .tx-insight{background:#16261d;border-color:#24462f}html.tt-dark #tran-xanh .tx-seg{background:#2b2724}html.tt-dark #tran-xanh .tx-seg button.on,html.tt-dark #tran-xanh .tx-more{background:#3a3633;color:#9fe3bd}';

var UI = { list: 'bought', day: 'all', grp: 'all', q: '', limit: 60 };
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"]/g, function(c){ return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]; }); }
function nf(v, d){ return new Intl.NumberFormat('vi-VN', {minimumFractionDigits: d || 0, maximumFractionDigits: d || 0}).format(Number(v) || 0); }
function pc(v){ return nf((Number(v) || 0) * 100, 1) + '%'; }
function tone(v){ return v >= 1 ? 'good' : v >= 0.8 ? 'watch' : 'bad'; }
function norm(s){ return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/đ/gi, 'd').toLowerCase(); }
function curRep(){ try { return (typeof state !== 'undefined' && state && state.rep) ? state.rep : ''; } catch(e){ return ''; } }
function dayList(v){ return String(v || '').split(',').map(function(x){ return x.trim().replace('Thứ ', 'T'); }).filter(Boolean); }

function model(){
  var TX = window.TRAN_XANH; if (!TX || !TX.list) return null;
  var T = window.T11_DATA || {}, K = (window.KPI_DATA || (window.KPI_MONTHS && window.KPI_MONTHS[Object.keys(window.KPI_MONTHS).sort().pop()]) || {});
  var team = (T.reps || []).map(function(r){ return r.name; });
  var teamSet = {}; team.forEach(function(n){ teamSet[n] = 1; });
  var L = TX.listCols, ix = function(n){ return L.indexOf(n); };
  var iC = ix('Mã KH'), iN = ix('Tên KH'), iR = ix('Đại diện KD'), iD = ix('Thứ'), iF = ix('Tần suất'), iG = ix('Nhóm'), iA = ix('Chanté Tết'), iB = ix('Tea365 Tết'), iBoth = ix('Mua cả 2'), iP = ix('Phân loại');
  var bought = TX.list.map(function(r){ return { c: String(r[iC]), n: r[iN], rep: r[iR], visit: r[iD], freq: r[iF], grp: r[iG] || 'Lẻ', a: +r[iA] || 0, b: +r[iB] || 0, both: +r[iBoth] || 0, cls: r[iP] }; })
    .filter(function(x){ return !team.length || teamSet[x.rep]; });
  var bset = {}; bought.forEach(function(x){ bset[x.c] = 1; });
  var ds = {}; (T.dp || []).forEach(function(d){ if (/đại sứ/i.test(d.prog || '') || /đại sứ/i.test(d.src || '')) ds[String(d.c)] = 1; });
  var vip = {}; (T.hv || []).forEach(function(h){ if (/VIP3|VIP5/i.test(h.ty || '')) vip[String(h.c)] = 1; });
  var cust = (T.cust || []).filter(function(c){ return /off/i.test(c.ch || ''); });
  var custBy = {}; cust.forEach(function(c){ custBy[String(c.c)] = c; });
  bought.forEach(function(x){ var c = custBy[x.c]; if (c){ x.addr = c.addr; x.ward = c.ward; } });
  var notYet = cust.filter(function(c){ return !bset[String(c.c)]; }).map(function(c){
    var k = String(c.c); var grp = ds[k] ? 'Đại Sứ' : (vip[k] || /VIP3|VIP5/i.test(c.type || '') ? 'VIP3-5' : 'Lẻ');
    return { c: k, n: c.n, rep: c.rep, visit: c.visit, freq: c.freq, grp: grp, addr: c.addr, ward: c.ward, ds3: c.ds };
  });
  // KPI Trận Xanh từ KPI THÁNG
  var kp = {}; (K.kpis || []).forEach(function(k){ kp[k.rep] = { ct: +k.teaGreenTarget || 0, th: +k.teaGreenActual || 0 }; });
  // Tổng hợp theo ĐDKD từ sheet
  var S = TX.sumCols, sx = function(n){ return S.indexOf(n); };
  var sum = (TX.sum || []).map(function(r){ var o = {}; S.forEach(function(n, i){ o[n] = r[i]; }); return o; })
    .filter(function(o){ return !team.length || teamSet[o['Đại diện KD']]; });
  var meta = (K.meta || {}), tm = T.meta || {};
  return { TX: TX, team: team, bought: bought, notYet: notYet, kp: kp, sum: sum, meta: meta, tmeta: tm, sx: sx };
}

function scoped(arr, rep){ return rep ? arr.filter(function(x){ return x.rep === rep || x['Đại diện KD'] === rep; }) : arr; }
function tot(rows, key){ return rows.reduce(function(s, r){ return s + (Number(r[key]) || 0); }, 0); }

function cards(M, rep){
  var S = scoped(M.sum, rep);
  var off = tot(S, 'Tổng KH OFF'), buy = tot(S, 'KH có mua'), a = tot(S, 'Chanté Tết'), b = tot(S, 'Tea365 Tết'), both = tot(S, 'Mua cả 2');
  var reps = rep ? [rep] : M.team; var ct = 0, th = 0; reps.forEach(function(n){ var k = M.kp[n]; if (k){ ct += k.ct; th += k.th; } });
  var kpiPct = ct ? th / ct : 0, cov = off ? buy / off : 0;
  function card(lbl, val, sub, cls, bar){ return '<div class="tx-card ' + (cls || '') + '"><span>' + lbl + '</span><strong>' + val + '</strong><small>' + sub + '</small>' + (bar != null ? '<div class="tx-bar"><i style="width:' + Math.min(100, Math.max(0, bar * 100)).toFixed(1) + '%"></i></div>' : '') + '</div>'; }
  return '<div class="tx-cards">' +
    card('ASO Trận Xanh (KPI)', ct ? pc(kpiPct) : 'Thiếu dữ liệu', 'TH ' + nf(th) + ' / CT ' + nf(ct) + ' · sheet KPI THÁNG', ct ? tone(kpiPct) : '', ct ? kpiPct : null) +
    card('KH OFF có mua', nf(buy), 'trên ' + nf(off) + ' KH OFF MCP · ' + pc(cov), '', cov) +
    card('Mua Chanté Tết', nf(a), pc(off ? a / off : 0) + ' KH OFF') +
    card('Mua Tea365 Tết', nf(b), pc(off ? b / off : 0) + ' KH OFF') +
    card('Mua cả 2', nf(both), pc(buy ? both / buy : 0) + ' KH đã mua') +
    card('KH OFF chưa mua', nf(Math.max(0, off - buy)), 'Ưu tiên Đại Sứ → VIP3-5 → Lẻ') +
  '</div>';
}

function insight(M, rep){
  var rows = scoped(M.sum, rep).map(function(o){ var k = M.kp[o['Đại diện KD']] || {ct:0, th:0}; return { n: o['Đại diện KD'], cov: (+o['Tổng KH OFF']) ? (+o['KH có mua']) / (+o['Tổng KH OFF']) : 0, buy: +o['KH có mua'] || 0, pct: k.ct ? k.th / k.ct : 0, ct: k.ct, th: k.th }; });
  if (!rows.length) return '';
  var red = rows.filter(function(r){ return r.ct && r.pct < 0.8; }), best = rows.slice().sort(function(a, b){ return b.cov - a.cov; })[0], worst = rows.slice().sort(function(a, b){ return a.cov - b.cov; })[0];
  var ny = scoped(M.notYet, rep), dsLeft = ny.filter(function(x){ return x.grp === 'Đại Sứ'; }).length, vipLeft = ny.filter(function(x){ return x.grp === 'VIP3-5'; }).length;
  var s = [];
  if (!rep) s.push('<b>' + red.length + '/' + rows.length + ' ĐDKD</b> dưới 80% CT ASO Trận Xanh (KPI THÁNG).');
  if (!rep && best) s.push('Độ phủ KH OFF cao nhất: <b>' + esc(best.n) + '</b> ' + pc(best.cov) + ' · thấp nhất: <b>' + esc(worst.n) + '</b> ' + pc(worst.cov) + '.');
  s.push('Còn <b>' + nf(dsLeft) + ' KH Đại Sứ</b> và <b>' + nf(vipLeft) + ' KH VIP3-5</b> chưa mua SP Tết — nên đi trước theo lịch tuyến.');
  return '<div class="tx-insight">💡 ' + s.join(' ') + '</div>';
}

function repTable(M, rep){
  var rows = scoped(M.sum, rep);
  if (!rows.length) return '<p class="tx-note">Thiếu dữ liệu: không có dòng tổng hợp ĐDKD trong sheet Trận Xanh.</p>';
  var r2 = rows.map(function(o){ var k = M.kp[o['Đại diện KD']] || {ct:0, th:0}; return { o: o, ct: k.ct, th: k.th, pct: k.ct ? k.th / k.ct : 0, cov: (+o['Tổng KH OFF']) ? (+o['KH có mua']) / (+o['Tổng KH OFF']) : 0 }; })
    .sort(function(a, b){ return a.pct - b.pct || a.cov - b.cov; });
  var cols = ['Tổng KH OFF', 'KH có mua', 'Chanté Tết', 'Tea365 Tết', 'Mua cả 2', 'Đại Sứ mua', 'VIP3-5 mua', 'Lẻ mua', 'Đại Sứ + VIP3-5 mua cả 2'];
  cols = cols.filter(function(c){ return M.TX.sumCols.indexOf(c) >= 0; });
  var h = '<div class="tx-wrap"><table><thead><tr><th>ĐDKD</th><th>CT ASO</th><th>TH ASO</th><th>% đạt</th>' + cols.map(function(c){ return '<th>' + esc(c) + '</th>'; }).join('') + '<th>% KH OFF có mua</th></tr></thead><tbody>';
  var T = { ct: 0, th: 0 }; cols.forEach(function(c){ T[c] = 0; });
  r2.forEach(function(x){
    T.ct += x.ct; T.th += x.th; cols.forEach(function(c){ T[c] += Number(x.o[c]) || 0; });
    h += '<tr><td>' + esc(x.o['Đại diện KD']) + '</td><td>' + nf(x.ct) + '</td><td>' + nf(x.th) + '</td><td class="bg-' + (x.ct ? tone(x.pct) : '') + ' c-' + (x.ct ? tone(x.pct) : '') + '">' + (x.ct ? pc(x.pct) : '–') + '</td>' +
      cols.map(function(c){ return '<td>' + nf(x.o[c]) + '</td>'; }).join('') + '<td>' + pc(x.cov) + '</td></tr>';
  });
  if (r2.length > 1){ var tp = T.ct ? T.th / T.ct : 0, tc = T['Tổng KH OFF'] ? T['KH có mua'] / T['Tổng KH OFF'] : 0;
    h += '<tr class="tot"><td>TỔNG TEAM (' + r2.length + ')</td><td>' + nf(T.ct) + '</td><td>' + nf(T.th) + '</td><td class="c-' + tone(tp) + '">' + pc(tp) + '</td>' + cols.map(function(c){ return '<td>' + nf(T[c]) + '</td>'; }).join('') + '<td>' + pc(tc) + '</td></tr>'; }
  return h + '</tbody></table></div><div class="tx-note">CT/TH ASO lấy từ sheet KPI THÁNG (ASO Trận Xanh). % đạt = TH ÷ CT — xanh ≥100%, vàng 80–99%, đỏ &lt;80%. Các cột còn lại lấy nguyên từ bảng "Tổng hợp theo ĐDKD" của sheet Trận Xanh. % KH OFF có mua = KH có mua ÷ Tổng KH OFF. Xếp từ % đạt thấp lên cao.</div>';
}

function grpTag(g){ return '<span class="tx-tag ' + (g === 'Đại Sứ' ? 't-ds' : g === 'VIP3-5' ? 't-vip' : 't-le') + '">' + esc(g) + '</span>'; }
var GORD = { 'Đại Sứ': 0, 'VIP3-5': 1, 'Lẻ': 2 };

function custRows(M, rep){
  var src = UI.list === 'bought' ? M.bought : M.notYet;
  var L = scoped(src, rep);
  if (UI.day !== 'all') L = L.filter(function(x){ return dayList(x.visit).indexOf(UI.day) >= 0; });
  if (UI.grp !== 'all') L = L.filter(function(x){ return x.grp === UI.grp; });
  if (UI.q.trim()){ var q = norm(UI.q.trim()); L = L.filter(function(x){ return norm(x.c + ' ' + x.n + ' ' + x.rep + ' ' + (x.addr || '')).indexOf(q) >= 0; }); }
  return L.slice().sort(function(a, b){ return (GORD[a.grp] - GORD[b.grp]) || String(a.rep).localeCompare(String(b.rep)) || String(a.n).localeCompare(String(b.n)); });
}

function custPanel(M, rep){
  var base = scoped(UI.list === 'bought' ? M.bought : M.notYet, rep);
  var days = ['T2','T3','T4','T5','T6','T7','CN'].filter(function(d){ return base.some(function(x){ return dayList(x.visit).indexOf(d) >= 0; }); });
  var L = custRows(M, rep), show = L.slice(0, UI.limit);
  var seg = function(key, val, lbl, n){ return '<button type="button" data-tx="' + key + '" data-v="' + esc(val) + '" class="' + (UI[key] === val ? 'on' : '') + '">' + lbl + (n != null ? ' · ' + nf(n) : '') + '</button>'; };
  var nB = scoped(M.bought, rep).length, nN = scoped(M.notYet, rep).length;
  var h = '<div class="tx-tools"><div class="tx-seg">' + seg('list', 'bought', '✅ Đã mua', nB) + seg('list', 'notyet', '⏳ Chưa mua', nN) + '</div>' +
    '<div class="tx-seg">' + seg('grp', 'all', 'Tất cả nhóm') + seg('grp', 'Đại Sứ', 'Đại Sứ') + seg('grp', 'VIP3-5', 'VIP3-5') + seg('grp', 'Lẻ', 'Lẻ') + '</div></div>' +
    '<div class="tx-tools"><div class="tx-seg">' + seg('day', 'all', 'Cả tuần') + days.map(function(d){ return seg('day', d, d); }).join('') + '</div>' +
    '<input class="tx-q" id="txQ" placeholder="Tìm tên KH, mã KH, ĐDKD, địa chỉ…" value="' + esc(UI.q) + '"><button class="tx-btn" type="button" id="txExport">⇩ Xuất Excel</button></div>';
  if (UI.list === 'bought'){
    h += '<div class="tx-wrap"><table><thead><tr><th>Khách hàng</th><th class="l">Nhóm</th><th class="l">ĐDKD</th><th class="l">Thứ</th><th class="l">TS</th><th>Chanté Tết</th><th>Tea365 Tết</th><th class="l">Phân loại</th></tr></thead><tbody>' +
      show.map(function(x){ return '<tr><td>' + esc(x.n) + '<br><small style="color:#8a837d;font-weight:500">' + esc(x.c) + (x.addr ? ' · ' + esc(x.addr) : '') + '</small></td><td class="l">' + grpTag(x.grp) + '</td><td class="l">' + esc(x.rep) + '</td><td class="l">' + esc(x.visit) + '</td><td class="l">' + esc(x.freq) + '</td><td>' + (x.a ? '<span class="ok">✓</span>' : '<span class="no">–</span>') + '</td><td>' + (x.b ? '<span class="ok">✓</span>' : '<span class="no">–</span>') + '</td><td class="l">' + esc(x.cls) + '</td></tr>'; }).join('') + '</tbody></table></div>';
  } else {
    h += '<div class="tx-wrap"><table><thead><tr><th>Khách hàng</th><th class="l">Nhóm</th><th class="l">ĐDKD</th><th class="l">Thứ</th><th class="l">TS</th><th class="l">Phường/Xã</th><th>DS tháng (k)</th></tr></thead><tbody>' +
      show.map(function(x){ return '<tr><td>' + esc(x.n) + '<br><small style="color:#8a837d;font-weight:500">' + esc(x.c) + (x.addr ? ' · ' + esc(x.addr) : '') + '</small></td><td class="l">' + grpTag(x.grp) + '</td><td class="l">' + esc(x.rep) + '</td><td class="l">' + esc(x.visit) + '</td><td class="l">' + esc(x.freq) + '</td><td class="l">' + esc(x.ward || '') + '</td><td>' + nf((x.ds3 || 0) / 1000) + '</td></tr>'; }).join('') + '</tbody></table></div>';
  }
  if (!L.length) h += '<p class="tx-note">Không có khách hàng phù hợp bộ lọc.</p>';
  if (L.length > show.length) h += '<button type="button" class="tx-more" id="txMore">Xem thêm (' + nf(show.length) + '/' + nf(L.length) + ')</button>';
  h += '<div class="tx-note">' + (UI.list === 'bought' ? 'Danh sách KH OFF đã mua lấy từ sheet Trận Xanh.' : 'KH chưa mua = KH kênh Off Premise trong MCP chưa có trong danh sách đã mua. Nhóm: Đại Sứ (chương trình Đại Sứ ở T.Bày) → VIP3-5 (VIP3/VIP5) → Lẻ. DS tháng = doanh số tháng hiện tại của quầy (nghìn đồng).') + '</div>';
  return h;
}

function exportExcel(M, rep){
  var go = function(){
    var L = custRows(M, rep), wb = XLSX.utils.book_new();
    var S = scoped(M.sum, rep).map(function(o){ var k = M.kp[o['Đại diện KD']] || {ct:0, th:0}; var r = { 'ĐDKD': o['Đại diện KD'], 'CT ASO Trận Xanh': k.ct, 'TH ASO Trận Xanh': k.th, '% đạt': k.ct ? k.th / k.ct : 0 }; M.TX.sumCols.forEach(function(c){ if (c !== 'Đại diện KD') r[c] = o[c]; }); return r; });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(S), 'Tổng hợp ĐDKD');
    var rows = L.map(function(x){ return UI.list === 'bought' ? { 'Mã KH': x.c, 'Tên KH': x.n, 'Nhóm': x.grp, 'ĐDKD': x.rep, 'Thứ': x.visit, 'Tần suất': x.freq, 'Chanté Tết': x.a, 'Tea365 Tết': x.b, 'Mua cả 2': x.both, 'Phân loại': x.cls, 'Địa chỉ': x.addr || '' } : { 'Mã KH': x.c, 'Tên KH': x.n, 'Nhóm': x.grp, 'ĐDKD': x.rep, 'Thứ': x.visit, 'Tần suất': x.freq, 'Phường/Xã': x.ward || '', 'Địa chỉ': x.addr || '', 'DS tháng': x.ds3 || 0 }; });
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), UI.list === 'bought' ? 'KH đã mua' : 'KH chưa mua');
    XLSX.writeFile(wb, 'Tran_Xanh_' + (rep ? rep.replace(/\s+/g, '_') + '_' : '') + (UI.list === 'bought' ? 'da_mua' : 'chua_mua') + '.xlsx');
  };
  if (window.XLSX && XLSX.utils) return go();
  var s = document.createElement('script'); s.src = 'xlsx.full.min.js'; s.onload = go; s.onerror = function(){ alert('Không tải được thư viện Excel.'); }; document.head.appendChild(s);
}

window.renderTranXanh = function(){
  var root = document.getElementById('content'); if (!root) return;
  if (!document.getElementById('tx-css')){ var st = document.createElement('style'); st.id = 'tx-css'; st.textContent = CSS; document.head.appendChild(st); }
  var M = model(), rep = curRep();
  if (!M){ root.innerHTML = '<section id="tran-xanh"><div class="tx-hero"><div><small>SALES INTELLIGENCE · TEAM SS VĂN</small><h2>TRẬN XANH</h2><p>Thiếu dữ liệu: file đang dùng chưa có sheet "Trận Xanh" (cần các cột Mã KH, Tên KH, Đại diện KD, Nhóm, Chanté Tết, Tea365 Tết…). Nạp file Tracking KPI có sheet này để xem tab.</p></div></div></section>'; return; }
  var upd = (M.meta.displayUpdated || M.meta.updated || M.tmeta.updated || ''), el = M.tmeta.elapsed, wd = M.tmeta.workingDays;
  var sku = Object.keys(M.TX.sku || {}).map(function(k){ return '<b>' + esc(k) + '</b>: ' + esc((M.TX.sku[k] || []).join(', ')); }).join(' · ');
  root.innerHTML = '<section id="tran-xanh"><div class="tx-hero"><div><small>SALES INTELLIGENCE · ' + esc(rep ? rep.toUpperCase() : 'TEAM SS VĂN') + '</small><h2>' + esc(M.TX.title || 'TRẬN XANH') + '</h2><p>' + esc(M.TX.note || '') + '</p></div><div class="tx-time"><b>' + esc(upd) + '</b><span>' + (el && wd ? el + '/' + wd + ' ngày làm việc · tiến độ ' + pc(el / wd) : '') + '</span></div></div>' +
    cards(M, rep) + insight(M, rep) +
    '<div class="tx-panel"><div class="tx-head"><div><h3>TỔNG HỢP THEO ĐDKD</h3><p>ASO Trận Xanh theo KPI THÁNG và độ phủ KH OFF mua SP Tết</p></div><span class="tx-pill">' + (rep ? '1 ĐDKD' : nf(M.team.length) + ' ĐDKD') + '</span></div>' + repTable(M, rep) + '</div>' +
    '<div class="tx-panel"><div class="tx-head"><div><h3>DANH SÁCH KHÁCH HÀNG</h3><p>Lọc theo nhóm, thứ viếng thăm · chưa mua xếp ưu tiên Đại Sứ → VIP3-5 → Lẻ</p></div></div>' + custPanel(M, rep) + '</div>' +
    (sku ? '<div class="tx-note">Mã SP Tết đang tính: ' + sku + '</div>' : '') + '</section>';
  bind(M, rep);
};

function bind(M, rep){
  var root = document.getElementById('tran-xanh'); if (!root) return;
  root.querySelectorAll('[data-tx]').forEach(function(b){ b.onclick = function(){ UI[b.getAttribute('data-tx')] = b.getAttribute('data-v'); UI.limit = 60; if (b.getAttribute('data-tx') === 'list') UI.day = 'all'; window.renderTranXanh(); }; });
  var q = document.getElementById('txQ'); if (q){ var t; q.oninput = function(){ clearTimeout(t); t = setTimeout(function(){ UI.q = q.value; UI.limit = 60; var pos = q.selectionStart; window.renderTranXanh(); var q2 = document.getElementById('txQ'); if (q2){ q2.focus(); try { q2.setSelectionRange(pos, pos); } catch(e){} } }, 250); }; }
  var m = document.getElementById('txMore'); if (m) m.onclick = function(){ UI.limit += 100; window.renderTranXanh(); };
  var ex = document.getElementById('txExport'); if (ex) ex.onclick = function(){ exportExcel(M, rep); };
}

function fixHeader(){
  // Dòng phụ dưới logo trước đây ghi cứng "DỮ LIỆU 10/11 · NGÀY 8/25" → lấy theo dữ liệu đang nạp
  try {
    var m = (window.T11_DATA && window.T11_DATA.meta) || {}; if (!m.updated) return;
    var u = String(m.updated).split('/'), mon = (m.month || '').split('-');
    document.querySelectorAll('small').forEach(function(s){
      if (/DỮ LIỆU \d{1,2}\/\d{1,2} · NGÀY \d+\/\d+/.test(s.textContent))
        s.textContent = s.textContent.replace(/T\d{1,2}\/\d{4} · DỮ LIỆU \d{1,2}\/\d{1,2} · NGÀY \d+\/\d+/, 'T' + (mon[1] || u[1]) + '/' + (mon[0] || u[2]) + ' · DỮ LIỆU ' + u[0] + '/' + u[1] + ' · NGÀY ' + m.elapsed + '/' + m.workingDays);
    });
  } catch(e){}
}
function install(){
  fixHeader();
  var nav = document.getElementById('nav'); if (!nav) return;
  var btn = nav.querySelector('[data-view="tranxanh"]');
  if (!btn){
    btn = document.createElement('button'); btn.className = 'nav-item'; btn.setAttribute('data-view', 'tranxanh');
    btn.innerHTML = '<span><svg viewBox="0 0 24 24"><path d="M12 3c4 3 6 6.5 6 10a6 6 0 0 1-12 0c0-3.5 2-7 6-10zM12 21v-8"/></svg></span>TRẬN XANH';
    var anchor = nav.querySelector('[data-view="daily"]') || nav.querySelector('[data-view="mtn"]');
    if (anchor && anchor.nextSibling) nav.insertBefore(btn, anchor.nextSibling); else nav.appendChild(btn);
  }
  if (btn.dataset.txBound) return; btn.dataset.txBound = '1';
  btn.addEventListener('click', function(e){
    e.preventDefault(); e.stopPropagation();
    nav.querySelectorAll('.nav-item').forEach(function(x){ x.classList.remove('active'); }); btn.classList.add('active');
    window.__txActive = true; var title = document.getElementById('pageTitle'); if (title) title.textContent = 'TRẬN XANH';
    var c = document.getElementById('content'); if (c) c.setAttribute('data-view', 'tranxanh');
    window.renderTranXanh(); try { window.scrollTo(0, 0); } catch(e){}
  });
  nav.addEventListener('click', function(e){ var b = e.target.closest && e.target.closest('.nav-item'); if (b && b !== btn) window.__txActive = false; }, true);
  var rf = document.getElementById('repFilter'); if (rf) rf.addEventListener('change', function(){ if (window.__txActive) setTimeout(window.renderTranXanh, 30); });
  var rs = document.getElementById('resetFilters'); if (rs) rs.addEventListener('click', function(){ if (window.__txActive) setTimeout(window.renderTranXanh, 30); });
  if (/#tranxanh/i.test(location.hash) || (window.parent !== window && /#tranxanh/i.test((function(){ try { return window.parent.location.hash; } catch(e){ return ''; } })()))) btn.click();
}
if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install); else install();
setTimeout(install, 800);
})();
