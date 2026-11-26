// conv.js — chuyển file Tracking KPI T11 (.xlsb) sang cấu trúc dữ liệu của dashboard "KPI Studio".
// Chạy được trong Node (build) và trình duyệt (nút Nạp file T11).
(function (root) {
  const SHEETS = ['KPI THÁNG', 'MCP', '061 new', 'ĐH Combo total', 'DSKH Combo', 'Sheet10', 'HV11 • Tracking', 'Pivot', 'Cap Tran', 'T.Bày', 'Bonus', 'LE', 'SINH NHẬT THÁNG', 'Data_Chỉ tiêu', 'Trận Xanh'];
  const num = v => (typeof v === 'number' && isFinite(v)) ? v : (typeof v === 'string' && v.trim() !== '' && !isNaN(+v.replace(/,/g, '')) ? +v.replace(/,/g, '') : 0);
  const str = v => (v == null ? '' : String(v)).trim();
  const norm = s => str(s).normalize('NFC').replace(/\s+/g, ' ').toLowerCase();
  const isCode = v => /SF\.|SB\./.test(str(v));
  const DOWS = ['CN', 'Thứ 2', 'Thứ 3', 'Thứ 4', 'Thứ 5', 'Thứ 6', 'Thứ 7'];
  const pd = v => { if (typeof v === 'number') { const d = new Date(Math.round((v - 25569) * 864e5)); return [d.getUTCDate(), d.getUTCMonth() + 1, d.getUTCFullYear()]; } const m = str(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); return m ? [+m[1], +m[2], +m[3]] : null; };
  const p2 = n => String(n).padStart(2, '0');
  const dowOf = d => DOWS[new Date(Date.UTC(d[2], d[1] - 1, d[0])).getUTCDay()];
  const days = s => str(s).split(',').map(x => x.trim()).filter(Boolean);
  const SUBS = ['Seasoning', 'Convenience Foods', 'Home Care', 'Refreshment Drinks', 'Processed Meats', 'Nutrition', 'Coffee', 'Rice base', 'Snack and Dessert', 'Beer'];
  const XM = s => s === 'Processed Meats' || s === 'Beer';
  const ON7 = ['Canteen', 'Giải trí', 'Máy bán hàng tự động', 'Quán nước', 'Quốc lộ', 'Thể thao', 'Vận chuyển hành khách'];
  const sum = (a, f) => a.reduce((n, x) => n + (f ? f(x) : x), 0);
  const rt = (a, b) => b ? a / b : 0;
  function dflt(t) { if (Array.isArray(t)) return []; if (t && typeof t === 'object') { const o = {}; for (const k in t) o[k] = dflt(t[k]); return o; } if (typeof t === 'number') return 0; if (typeof t === 'boolean') return false; if (typeof t === 'string') return ''; return null; }
  const mk = (tpl, v) => Object.assign(dflt(tpl), v);


  // ===== Trận Xanh (sheet "Trận Xanh") — thêm cho bản GitHub =====
  function tranXanh(rows) {
    const S = v => (v == null ? '' : String(v)).trim();
    if (!rows || !rows.length) return null;
    const hr = rows.findIndex(r => r && r.some(v => S(v) === 'Mã KH') && r.some(v => S(v) === 'Tên KH'));
    if (hr < 0) return null;
    const H = rows[hr].map(S);
    const block = (c0) => { const cols = []; for (let c = c0; c < H.length && H[c]; c++) cols.push(c); return cols; };
    const c0 = H.indexOf('Mã KH'); const lc = block(c0);
    const s0 = H.findIndex((v, i) => i > lc[lc.length - 1] && v === 'Mã NVBH'); const sc = s0 >= 0 ? block(s0) : [];
    const k0 = H.findIndex((v, i) => i > (sc.length ? sc[sc.length - 1] : lc[lc.length - 1]) && /Chant/i.test(v)); const kc = k0 >= 0 ? block(k0).filter(c => !/^KH /i.test(H[c])) : [];
    const sec = rows[hr - 1] || [];
    const list = [], sum = []; let total = null; const sku = {};
    kc.forEach(c => sku[H[c]] = []);
    for (let i = hr + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      if (S(r[c0])) list.push(lc.map(c => r[c] == null ? '' : r[c]));
      if (sc.length) { const k = S(r[s0]); if (k === 'TỔNG' || /^TỔNG/i.test(k)) total = sc.map(c => r[c] == null ? '' : r[c]); else if (k && !/^Kiểm tra/i.test(k)) sum.push(sc.map(c => r[c] == null ? '' : r[c])); }
      kc.forEach(c => { if (S(r[c])) sku[H[c]].push(S(r[c])); });
    }
    return { title: S((rows[0] || [])[0]), note: S((rows[1] || [])[0]), listTitle: S(sec[c0]) || 'DANH SÁCH KH OFF CÓ MUA', sumTitle: s0 >= 0 ? (S(sec[s0]) || 'TỔNG HỢP THEO ĐDKD') : '',
      listCols: lc.map(c => H[c]), list, sumCols: sc.map(c => H[c]), sum, total, sku };
  }
  function convert(XLSX, wb, opt) {
    const TPL = opt.tpl, H = opt.hist || {};
    const A = n => wb.Sheets[n] && wb.Sheets[n]['!ref'] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: true }) : [];
    const warn = [];
    const teamKey = opt.teamKey || 'Văn', TEAM = opt.team || 'Võ Viết Văn';

    // ===== KPI THÁNG =====
    const K = A('KPI THÁNG');
    const hI = K.findIndex(r => r && str(r[0]) === 'ĐẠI DIỆN KD'); if (hI < 0) throw new Error('Không thấy sheet KPI THÁNG');
    const H1 = K[hI], H2 = K[hI + 1];
    const col = (re, row = H1) => row.findIndex(v => re.test(str(v)));
    let today = '', wd = 25, passed = 0, prog = 0;
    for (let i = 0; i < hI; i++) (K[i] || []).forEach((v, j) => { const s = str(v); if (/TODAY/i.test(s)) today = (s.match(/\d{1,2}\/\d{1,2}\/\d{4}/) || [''])[0]; if (s === 'Ngày làm việc') wd = num(K[i][j - 1]); if (s === 'Đã qua') passed = num(K[i][j - 1]); if (s === 'Tiến độ') prog = num(K[i][j + 1]); });
    if (!prog) prog = passed / wd;
    const td = pd(today) || [10, 11, 2026]; const MON = td[1], YEAR = td[2];
    const MK = `${YEAR}-${p2(MON)}`;
    const C = { mcp: col(/^MCP THÁNG/i), s: col(/^DOANH SỐ BÁN RA/i), aOn: col(/^ACTIVE ASO MCP ON\s*$/i), aOff: col(/^ACTIVE ASO MCP OFF/i), aAll: col(/^ACTIVE ASO MCP ON \+ OFF/i), tv: col(/^Trận Vàng/i), tx: col(/^Trận Xanh/i), pcT: col(/^Tổng số PC/i), lb: col(/^TRACKING NỀN TẢNG/i), pOff: col(/^PC OFF/i), pOn: col(/^PC ON/i), lp: col(/^LPPC TRỪ MEAT/i), lm: col(/^LPPC MEAT/i), tea: col(/^TỔNG QUẦY ON MUA TEA365/i), sn: col(/^KHÁCH HÀNG SINH NHẬT/i) };
    const cDay = col(/^Chỉ Tiêu Ngày/i, H2);
    const indC = SUBS.map(n => col(new RegExp('^' + n + '$')));
    const hdr = k => str(H1[C[k]]).replace(/\s+/g, ' ');
    const rows = [];
    let nppRow = null; const ssRows = [];
    for (let i = hI + 2; i < K.length; i++) { const r = K[i]; if (!r) continue; if (isCode(r[0]) && str(r[1])) rows.push(r); else if (/^SS\s/i.test(str(r[0]))) ssRows.push(r); else if (str(r[0]) && !str(r[1]) && num(r[C.s]) > 0 && !nppRow) nppRow = r; }
    const NPP = nppRow ? str(nppRow[0]) : 'NPP';

    // ===== MCP =====
    const M = A('MCP'); const mH = M.findIndex(r => r && str(r[0]) === 'Mã KH'); const mh = M[mH].map(str);
    const mi = n => mh.findIndex(x => norm(x) === norm(n));
    const mc = { pos: mi('Position code'), cr: mi('Ngày tạo'), a: mi('Địa chỉ'), w: mi('Phường/ Xã'), vt: mi('Vị trí CH'), l1: mi('Loại hình KH 1'), l2: mi('Loại hình KH 2'), l3: mi('Loại hình KH 3'), ss: mi('SS'), nv: mi('Mã NVBH'), thu: mi('Thứ'), f: mi('Tần suất'), t: mi('Type'), ds3: mi('DS 3 tháng'), tg: mi('TARGET KH'), tea: mh.findIndex(x => x === 'TEA365') };
    const OUT = {}; const ssOf = {};
    for (let i = mH + 1; i < M.length; i++) { const r = M[i]; if (!r || !str(r[0])) continue; const nv = str(r[mc.nv]); if (nv && str(r[mc.ss])) ssOf[nv] = str(r[mc.ss]);
      OUT[str(r[0])] = { c: str(r[0]), n: str(r[1]), nv, ss: str(r[mc.ss]), a: str(r[mc.a]), w: str(r[mc.w]), vt: str(r[mc.vt]), l1: str(r[mc.l1]), l2: str(r[mc.l2]), l3: str(r[mc.l3]), ch: /On Premise/i.test(str(r[mc.l1])) ? 'ON' : 'OFF', v: str(r[mc.thu]), f: str(r[mc.f]), t: str(r[mc.t]) || 'LẺ', ds3: num(r[mc.ds3]), tg: num(r[mc.tg]), cr: pd(r[mc.cr]) }; }
    const inTeam = nv => new RegExp(teamKey, 'i').test(ssOf[nv] || '');
    const repRows = rows.filter(r => inTeam(str(r[0])));
    if (!repRows.length) throw new Error('Không tìm thấy ĐDKD thuộc team ' + TEAM);
    const REPS = repRows.map(r => ({ code: str(r[0]), name: str(r[1]), r }));
    const byCode = Object.fromEntries(REPS.map(x => [x.code, x]));
    const outs = Object.values(OUT).filter(o => byCode[o.nv]);
    const outsBy = {}; outs.forEach(o => (outsBy[o.nv] = outsBy[o.nv] || []).push(o));
    const wardOf = {}; REPS.forEach(x => { const c = {}; (outsBy[x.code] || []).forEach(o => c[o.w] = (c[o.w] || 0) + 1); wardOf[x.code] = Object.entries(c).sort((a, b) => b[1] - a[1])[0]?.[0] || ''; });
    const ssRow = ssRows.find(r => new RegExp(teamKey, 'i').test(str(r[0])));

    // ===== 061 new =====
    const O = A('061 new'); const oh = (O[0]||[]).map(str); const oi = n => oh.findIndex(x => norm(x) === norm(n));
    const oc = { nv: oi('Mã NVBH'), id: oi('Mã đơn hàng'), st: oi('Tình trạng đơn hàng'), dt: oi('Ngày tạo đơn hàng'), pos: oi('Tên vị trí công việc'), c: oi('Mã CH'), sub: oi('Sub Division'), g: oi('Group STD SKU'), sku: oi('Mã sản phẩm'), sn: oi('Tên sản phẩm'), le: oi('Tổng lẻ'), v: oi('Thành tiền trước CK'), va: oi('Giá trị sau CK'), km: oi('Hàng KM'), note: oi('NOTE'), br: oi('MBS_Brand') };
    const L = []; const posOf = {};
    for (let i = 1; i < O.length; i++) { const r = O[i]; if (!r || !str(r[oc.nv])) continue; if (/hủy/i.test(str(r[oc.st]))) continue; const d = pd(r[oc.dt]); if (!d || d[1] !== MON) continue; const nv = str(r[oc.nv]); if (!byCode[nv]) continue;
      posOf[nv] = posOf[nv] || str(r[oc.pos]); const c = str(r[oc.c]); const o = OUT[c];
      const km = str(r[oc.km]) === 'Y';
      L.push({ nv, id: str(r[oc.id]), dd: p2(d[0]), d, c, sub: str(r[oc.sub]), g: str(r[oc.g]), sku: str(r[oc.sku]), sn: str(r[oc.sn]), le: num(r[oc.le]), v: km ? 0 : num(r[oc.v]), va: km ? 0 : num(r[oc.va]), km, ch: /On Premise/i.test(str(r[oc.note])) ? 'ON' : (/Off Premise/i.test(str(r[oc.note])) ? 'OFF' : (o ? o.ch : 'OFF')), br: str(r[oc.br]) || 'Khác' }); }
    // ===== Tuần chẵn/lẻ của KH tần suất F2 (suy từ ngày có đơn: Pivot + 061 new) =====
    const isoWk = d => { const t = new Date(Date.UTC(d[2], d[1] - 1, d[0])); const dy = (t.getUTCDay() + 6) % 7; t.setUTCDate(t.getUTCDate() - dy + 3); const f = new Date(Date.UTC(t.getUTCFullYear(), 0, 4)); return 1 + Math.round(((t - f) / 864e5 - 3 + ((f.getUTCDay() + 6) % 7)) / 7); };
    const PV = A('Pivot'); const vote = {}; const addV = (c, d) => { if (!c || !d) return; const w = isoWk(d) % 2; (vote[c] = vote[c] || [0, 0])[w]++; };
    PV.forEach((r, i) => { if (!r || i < 3) return; addV(str(r[2]), pd(r[4])); addV(str(r[20]), pd(r[22])); });
    L.forEach(l => addV(l.c, l.d));
    const parOf = o => { if (!/^F2$/i.test(o.f)) return 'Both'; const z = vote[o.c]; if (!z) return ''; return z[1] > z[0] ? 'Odd Week' : z[0] > z[1] ? 'Even Week' : 'Both'; };
    Object.values(OUT).forEach(o => { o.p = parOf(o); });
    const f2u = outs.filter(o => /^F2$/i.test(o.f) && !o.p).length; if (f2u) warn.push(`${f2u} KH tần suất F2 chưa xác định được tuần chẵn/lẻ (chưa có đơn trong T8 và T11) — chưa tính vào KH tuyến ngày`);
    const onRoute = (o, dd) => { const dw = dowOf([+dd, MON, YEAR]); if (!days(o.v).includes(dw)) return false; if (o.p === 'Both') return true; return o.p === (isoWk([+dd, MON, YEAR]) % 2 ? 'Odd Week' : 'Even Week'); };
    const isTV = l => /hạt nêm/i.test(l.sn) && /ngọt thanh/i.test(l.sn);
    const isTea = l => /tea\s?365/i.test(l.sn) || /tea 365/i.test(l.g);
    const DDS = [...new Set(L.map(l => l.dd))].sort().filter(dd => dowOf([+dd, MON, YEAR]) !== 'CN');
    if (!L.length) warn.push('Sheet 061 new không có đơn của team');
    // tổng hợp theo quầy
    const OS = {}; L.forEach(l => { const z = OS[l.c] = OS[l.c] || { ds: 0, dsa: 0, ord: new Set(), dd: new Set(), tea: 0, byDay: {}, ordV: {}, last: '' }; if (l.km) return; z.ds += l.v; z.dsa += l.va; if (l.v > 0) { z.ord.add(l.id); z.dd.add(l.dd); z.byDay[l.dd] = (z.byDay[l.dd] || 0) + l.v; z.ordV[l.id] = (z.ordV[l.id] || 0) + l.v; if (l.dd > z.last) z.last = l.dd; } if (isTea(l)) z.tea += l.v; });
    const ordVal = {}; L.forEach(l => { if (!l.km) ordVal[l.id] = (ordVal[l.id] || 0) + l.v; });
    const bought = c => OS[c] && OS[c].ds > 0;
    // chỉ số kênh: {pc(KH), lines, ds, orders}
    function chan(ls) { const k = new Set(), o = new Set(); let ln = 0, ds = 0; ls.forEach(l => { if (l.km || l.v <= 0) return; k.add(l.c); o.add(l.id); ln++; ds += l.v; }); return { pc: k.size, lines: ln, ds, orders: o.size }; }
    const LBy = {}; L.forEach(l => (LBy[l.nv] = LBy[l.nv] || []).push(l));

    // ===== Data_Chỉ tiêu (CT theo nhóm SP) =====
    const DC = A('Data_Chỉ tiêu'); const ctG = {}; // nv -> sub -> group -> ct
    DC.forEach(r => { if (!r || !isCode(r[9]) || !str(r[11]) || !str(r[12])) return; const nv = str(r[9]); if (!byCode[nv]) return; const s = str(r[11]), g = str(r[12]); ctG[nv] = ctG[nv] || {}; ctG[nv][s] = ctG[nv][s] || {}; ctG[nv][s][g] = (ctG[nv][s][g] || 0) + num(r[13]); });
    // brand theo group (từ đơn hàng)
    const brOfG = {}; L.forEach(l => { if (l.g && l.br && l.br !== 'Khác') brOfG[l.g] = l.br; });
    const BRK = [['Vinacafe', /vinacaf/i], ['Wake up 247', /247/], ['Wake up', /wake ?up/i], ['Omachi', /omachi/i], ['Kokomi', /kokomi|komi/i], ['Chanté', /chant/i], ['Homey', /homey/i], ['NET', /\bnet\b/i], ['Nam Ngư', /nam ngư/i], ['Chinsu', /chin-?su/i], ['Tam Thái Tử', /tam thái/i], ['Heo Cao Bồi', /cao bồi/i], ['Ponnie', /ponnie|bin & bon|bin bon/i], ['Bupnon TEA365', /tea ?365|búp non|bupnon/i], ["B'fast", /b'?fast/i], ['Red Ruby', /ruby/i], ['Vivant', /vivant/i], ['Compact', /compact/i], ['Joins', /joins/i]];
    const brandOf = g => brOfG[g] || (BRK.find(([, re]) => re.test(g)) || ['Khác'])[0];
    function tree(nv, ls) { const t = {}; const add = (s, b, g, ct, l) => { const S = t[s] = t[s] || { t: [0, 0, new Set(), 0, new Set()], b: {} }; const B = S.b[b] = S.b[b] || { t: [0, 0, new Set(), 0, new Set()], o: {} }; const G = B.o[g] = B.o[g] || [0, 0, new Set(), 0, new Set()];
        [S.t, B.t, G].forEach(a => { a[0] += ct; if (l) { a[1] += l.v; a[2].add(l.c); a[3]++; a[4].add(l.id); } }); };
      if (nv) { const cg = ctG[nv] || {}; for (const s in cg) for (const g in cg[s]) add(s, brandOf(g), g, cg[s][g], null); }
      ls.forEach(l => { if (l.km || l.v <= 0) return; add(l.sub || 'Khác', l.br !== 'Khác' ? l.br : brandOf(l.g), l.g || '(Chưa map)', 0, l); });
      const fix = a => [a[0], a[1], a[2].size, a[3], a[4].size];
      const out = {}; for (const s in t) { out[s] = { t: fix(t[s].t), b: {} }; for (const b in t[s].b) { out[s].b[b] = { t: fix(t[s].b[b].t), o: {} }; for (const g in t[s].b[b].o) out[s].b[b].o[g] = fix(t[s].b[b].o[g]); } } return out; }

    // ===== Combo =====
    const CB = A('ĐH Combo total'); const ch0 = (CB[0]||[]).map(str); const ci = n => ch0.findIndex(x => norm(x) === norm(n));
    const cbc = { id: ci('Mã đơn hàng'), st: ci('Trạng thái đơn'), pg: ci('Chương trình'), q: ci('Số suất đạt'), c: ci('Mã CH'), dt: ci('Ngày đơn chuẩn'), ok: ci('ĐH hợp lệ duy nhất'), nv: ci('Mã DDKD'), d2: ci('Ngày tạo đơn') };
    const cgrp = s => /800K|VIP1/i.test(s) ? 'off' : (/250K|Bev|Tea 365/i.test(s) ? 'on' : 'nh');
    const CEV = [];
    for (let i = 1; i < CB.length; i++) { const r = CB[i]; if (!r || !str(r[cbc.id])) continue; if (/hủy/i.test(str(r[cbc.st]))) continue; if (cbc.ok >= 0 && !num(r[cbc.ok])) continue; const nv = str(r[cbc.nv]); if (!byCode[nv]) continue; const d = pd(r[cbc.dt]) || pd(r[cbc.d2]); if (!d) continue;
      CEV.push({ c: str(r[cbc.c]), nv, rep: byCode[nv].name, d: `${YEAR}-${p2(d[1])}-${p2(d[0])}`, dd: p2(d[0]), g: cgrp(str(r[cbc.pg])), q: num(r[cbc.q]) || 1, v: ordVal[str(r[cbc.id])] || 0 }); }
    const DK = A('DSKH Combo'); const elig = {}; // c -> Set(g)
    if (DK.length) { const h = DK[0].map(str); const dx = n => h.findIndex(x => norm(x) === norm(n)); const k = { c: dx('OUTLET_CD'), nv: dx('SM UID'), grp: dx('PROMO_GROUP'), pg: dx('PROMOTION_NAME') };
      for (let i = 1; i < DK.length; i++) { const r = DK[i]; if (!r || !str(r[k.c])) continue; if (!byCode[str(r[k.nv])]) continue; const gs = str(r[k.grp]); const g = /OFF/i.test(gs) ? 'off' : (/ON/i.test(gs) ? 'on' : 'nh'); (elig[str(r[k.c])] = elig[str(r[k.c])] || new Set()).add(g); } }
    const cevBy = {}; CEV.forEach(e => (cevBy[e.c] = cevBy[e.c] || []).push(e));

    // ===== Sheet10 (gắn kết tăng trưởng) =====
    const S10 = A('Sheet10'); const s0 = (S10[0]||[]).map(str); const si = n => s0.findIndex(x => norm(x) === norm(n));
    const sc = { uid: si('SM UID'), c: si('Outlet Code'), n: si('Outlet Name'), pg: si('Tên Chương trình'), dsc: si('Detail Scheme'), ct: si('Chỉ tiêu doanh số'), th: si('Doanh số thực hiện (Paid & Invoiced)'), pt: si('Điểm thưởng dự kiến') };
    const GK = {}; const gparRaw = {};
    const gkey = s => /Chanté/i.test(s) ? 'ch' : /Cà phê/i.test(s) ? 'cf' : /Mì/i.test(s) ? 'mi' : /Bia/i.test(s) ? 'bia' : '';
    for (let i = 1; i < S10.length; i++) { const r = S10[i]; if (!r || !str(r[sc.c])) continue; const nv = str(r[sc.uid]); if (!byCode[nv]) continue; const k = gkey(str(r[sc.pg])); if (!k || k === 'bia') continue;
      const z = GK[str(r[sc.c])] = GK[str(r[sc.c])] || { nv, n: str(r[sc.n]) }; const ct = num(r[sc.ct]), th = num(r[sc.th]); z[k] = [ct, th, rt(th, ct), ct > 0 && th >= ct ? 1 : 0, num(r[sc.pt]) / 100];
      if (!gparRaw[k]) gparRaw[k] = str(r[sc.dsc]); }
    const gp = k => { const s = gparRaw[k] || ''; const th = +(s.match(/(\d+)%\s*trở lên/) || s.match(/từ\s*(\d+)%/) || [0, 100])[1] / 100; const pts = +(String((s.match(/(\d+(?:[.,]\d+)?)\s*lần số điểm/) || [0, 1])[1]).replace(',', '.')); const cap = s.match(/tối đa[^\d]*([\d,]+)\s*điểm/i); return [th, pts, cap ? num(cap[1].replace(/,/g, '')) / 100 : null]; };
    const GPAR = [['Chương trình', 'Mã', 'Ngưỡng % so nền tảng', 'Điểm / 1đ DS', 'Trần thưởng (đ)', '100 điểm = (đ)', null], ['Ngành Mì', 'MI', ...gp('mi'), 100, null], ['Chanté', 'CH', ...gp('ch'), 100, null], ['Ngành Cà phê', 'CF', ...gp('cf'), 100, null]];

    // ===== HV11 • Tracking + Cap Tran =====
    const HV = A('HV11 • Tracking'); const hvH = HV.findIndex(r => r && str(r[0]) === 'SM UID'); const HVR = [];
    if (hvH >= 0) { const h = HV[hvH].map(str); const hx = re => h.findIndex(x => re.test(x));
      const c = { ad: hx(/^Địa chỉ/), ss: hx(/^Giám sát/), thu: hx(/^Thứ$/), c: hx(/^Mã KH/), n: hx(/^Tên KH/), g: hx(/^Nhóm hội viên/), ct: hx(/^Chỉ tiêu DS nền/), th: hx(/^Doanh số tháng/), ch: hx(/^DS Chanté/), b4: hx(/^DS 4 nhãn/), cs: hx(/^DS Chinsu/), kk: hx(/^DS Kokomi/), nn: hx(/^DS Nam Ngư/), ds4: hx(/^Đại sứ: đủ 4/), vip: hx(/^VIP: 4 nhãn/), nh: hx(/^Điều kiện nhãn/), dk: hx(/^Điều kiện doanh số/), hs: hx(/^Hệ số/), kq: hx(/^Kết quả thưởng/), ly: hx(/^Lý do/), tt: hx(/^Doanh số tính thưởng sau cap/), pt: hx(/^Điểm thưởng/) };
      for (let i = hvH + 1; i < HV.length; i++) { const r = HV[i]; if (!r || !str(r[0])) continue; const nv = str(r[0]); if (!byCode[nv]) continue; HVR.push({ nv, c: str(r[c.c]), n: str(r[c.n]), v: str(r[c.thu]), a: str(r[c.ad]), grp: str(r[c.g]), ct: num(r[c.ct]), th: num(r[c.th]), ch: num(r[c.ch]), b4: num(r[c.b4]), cs: num(r[c.cs]), kk: num(r[c.kk]), nn: num(r[c.nn]), nh: str(r[c.nh]), dk: str(r[c.dk]), hs: num(r[c.hs]), kq: str(r[c.kq]), ly: str(r[c.ly]), tt: num(r[c.tt]), pt: num(r[c.pt]) }); } }
    const CT = A('Cap Tran'); const capBy = {}; const skuStat = {};
    for (let i = 1; i < CT.length; i++) { const r = CT[i]; if (!r || !str(r[3])) continue; const nv = str(r[0]); if (!byCode[nv]) continue; const sku = str(r[6]).replace(/^[A-Za-z ]+?(?=[A-ZĐ][^ ]*[À-ỹ]|Chin|Nam|Tam|Wake|Vina|Kokomi|Omachi|Tương|Mì|Xúc|Hạt|Nước)/, '').trim() || str(r[6]); const k = str(r[3]);
      const cap = num(r[8]), th = num(r[9]), ov = num(r[11]); (capBy[k] = capBy[k] || {})[sku] = [cap, th, cap - th, ov];
      const z = skuStat[sku] = skuStat[sku] || [0, 0, 0, 0, 0]; if (th > cap) z[0]++; else if (th > 0 && th >= 0.9 * cap) z[1]++; z[2] += cap; z[3] += th; z[4] += Math.max(0, ov); }
    const SKUS = Object.entries(skuStat).sort((a, b) => b[1][4] - a[1][4]).slice(0, 11).map(x => x[0]);

    // ===== T.Bày =====
    const TB = A('T.Bày'); const tH = TB.findIndex(r => r && r.some(v => str(v) === 'TARGET') && r.some(v => str(v) === 'Đại diện KD')); const DP = [];
    if (tH >= 0) { const h = TB[tH].map(str); const ix = n => h.findIndex(x => norm(x) === norm(n)); const c = { pc: ix('Column1'), p: ix('Column2'), m: ix('Column3'), k: ix('Column4'), n: ix('Column5'), a: ix('ĐỊA CHỈ'), w: ix('Phường/Xã'), d: ix('Thứ'), nv: ix('Mã Nv'), tg: ix('TARGET'), th: ix('THỰC HIỆN'), ch: ix('BỘ CHỤP'), cm: ix('BỘ CHẤM'), dt: ix('BỘ ĐẠT'), kq: ix('KQ Final') };
      for (let i = tH + 1; i < TB.length; i++) { const r = TB[i]; if (!r || !str(r[c.k])) continue; const nv = str(r[c.nv]); if (!byCode[nv]) continue; DP.push({ pc: str(r[c.pc]), p: str(r[c.p]), m: str(r[c.m]), c: str(r[c.k]), n: str(r[c.n]), a: str(r[c.a]), w: str(r[c.w]), v: str(r[c.d]), nv, tg: num(r[c.tg]), th: num(r[c.th]), shots: num(r[c.ch]), ok: num(r[c.dt]), kq: str(r[c.kq]) }); } }

    // ===== Bonus =====
    const B = A('Bonus'); const bT = B.findIndex(r => r && str(r[0]) === 'BONUS DDKD'); const BR = {}; let bFrom = 0; const BN = {};
    if (bT >= 0) { const tr0 = B[bT]; bFrom = tr0.findIndex(v => /BIA RUBY/i.test(str(v)));
      tr0.forEach((v, j) => { if (/MEAT THÁNG/i.test(str(v))) BN.meat = str(B[bT + 1][j]); if (/COMBO XỐT/i.test(str(v))) BN.combo = str(B[bT + 1][j]); if (/NƯỚC MẮM CAO CẤP/i.test(str(v))) BN.homey = str(B[bT + 1][j]); if (/CONVENIENCE FOODS/i.test(str(v))) BN.mi = str(B[bT + 1][j]); });
      for (let i = bT + 3; i < B.length; i++) { const r = B[i]; if (!r || !str(r[1])) continue; BR[isCode(r[0]) ? str(r[0]) : str(r[1])] = r.slice(bFrom, bFrom + 40); } }
    const bonusRec = (code, rep, ss, v) => { v = v || []; const g = i => num(v[i]);
      const ds = [g(10), g(13), g(16)], rw = [g(12), g(15), g(18)];
      const m1 = Math.min(1.5e6, ds[0] * .01) + Math.min(1.5e6, ds[1] * .03) + Math.min(1.5e6, ds[2] * .06), m2 = Math.min(1.5e6, ds[0] * .02) + Math.min(1.5e6, ds[1] * .05) + Math.min(1.5e6, ds[2] * .10);
      const mi = [g(32), g(33), rt(g(33), g(32)), g(36)], meat = [g(6), g(7), rt(g(7), g(6))], combo = [g(25), g(26), rt(g(26), g(25)), g(31)], homey = [g(21), g(22), rt(g(22), g(21)), g(24)];
      const tot = g(19); const all = mi[3] + tot + combo[3] + homey[3];
      return { code, rep, ss, mi, meat, lvl: str(v[9]) || 'CHƯA ĐẠT', ds, rw, tot, m1, m2, l100: Math.max(0, meat[0] - meat[1]), l105: Math.max(0, meat[0] * 1.05 - meat[1]), combo, homey, tbtl: [0, 0, 0, 0], tm: [0, 0, 0, 0, 0], sum: [mi[3], tot, combo[3], homey[3], 0, 0, all, 0] }; };

    // ===== Sinh nhật =====
    const SN = A('SINH NHẬT THÁNG'); const BD = {};
    if (SN.length) { const h = SN[0].map(str); const sx = n => h.findIndex(x => norm(x) === norm(n)); const c = { c: sx('Mã KH'), d: sx('Ngày SN'), u: sx('Ngày update SN') };
      for (let i = 1; i < SN.length; i++) { const r = SN[i]; if (!r || !str(r[c.c])) continue; const d = pd(r[c.d]); if (d) BD[str(r[c.c])] = { b: `${d[2]}-${p2(d[1])}-${p2(d[0])}`, m: d[1], u: (pd(r[c.u]) || null) }; } }

    // ===== LE =====
    const LE = A('LE'); const lH = LE.findIndex(r => r && str(r[0]) === 'Ngành Hàng'); const leInd = {}, leDd = {};
    if (lH >= 0) { let i = lH + 1; for (; i < LE.length; i++) { const r = LE[i]; if (!r || !str(r[0])) continue; if (str(r[0]) === 'TOTAL') { i++; break; } const s = SUBS.find(x => norm(str(r[0])).startsWith(norm(x))); if (s) leInd[s] = num(r[1]) ? num(r[6]) / num(r[1]) : 1; }
      for (; i < LE.length; i++) { const r = LE[i]; if (!r || !str(r[0]) || str(r[0]) === 'TOTAL') continue; if (num(r[1]) > 0) leDd[norm(r[0])] = num(r[6]) / num(r[1]); } }

    // =================== DỰNG CẤU TRÚC ===================
    const kv = (r, c, o = 0) => c >= 0 ? num(r[c + o]) : 0;
    // Trận Xanh: từ 26/11 khối cột đổi thành CT/TH Nhóm TB, CT/TH Nhóm TB+VIP, TH Nhóm Lẻ… → lấy CT/TH "Nhóm TB, VIP"; file cũ vẫn dùng CT (+0) / TH (+2)
    const TXO = (() => { if (C.tx < 0) return { ct: 0, th: 2 }; const nx = Object.values(C).filter(v => v > C.tx).reduce((m, v) => Math.min(m, v), C.tx + 12); let ct = -1, th = -1;
      for (let c = C.tx; c < nx; c++) { const h = str(H2[c]).replace(/\s+/g, ' ').trim(); if (/^CT Nhóm TB, ?VIP/i.test(h)) ct = c - C.tx; if (/^TH Nhóm TB, ?VIP/i.test(h)) th = c - C.tx; }
      return ct >= 0 && th >= 0 ? { ct, th } : { ct: 0, th: 2 }; })();
    const visitOf = o => o ? o.v : '';
    const team = TEAM;
    const KPIS = [], MBONUS = [], MMONTHLY = [], TREPS = [], TFUND = [], TBON = [], MTNREPS = [], FXREPS = [], FXON = [], FXX = [], FXNB = [], LEROWS = [];
    const TOV = {}, OVW = [];
    const nbList = [];
    const lastDays = DDS.slice(-2);
    // lịch sử
    const KH = H.K ? H.K.K.split(';').map(x => x.split(',').map(Number)) : [];
    const histBy = {}; KH.forEach(([d, a, i, m, t, v]) => { const n = H.K.D[d]; const z = (histBy[n] = histBy[n] || {})[m] = (histBy[n] || {})[m] || [0, 0]; z[0] += t * 1000; z[1] += v * 1000; });

    REPS.forEach((x, ri) => {
      const r = x.r, code = x.code, name = x.name, os = outsBy[code] || [], ls = LBy[code] || [];
      const mcpOff = kv(r, C.mcp, 0), mcpOn = kv(r, C.mcp, 1), mcpTot = kv(r, C.mcp, 2);
      const actOff = kv(r, C.aOff, 2), actOn = kv(r, C.aOn, 2);
      const target = kv(r, C.s), actual = kv(r, C.s, 1);
      const lpPc = kv(r, C.lp), lpLn = kv(r, C.lp, 1), lmPc = kv(r, C.lm), lm = kv(r, C.lm, 1);
      const gkr = Object.values(GK).filter(z => z.nv === code);
      const gsum = k => { const a = gkr.filter(z => z[k]); return [sum(a, z => z[k][0]), sum(a, z => z[k][1]), a.length, a.filter(z => z[k][3]).length, sum(a, z => z[k][3] ? z[k][4] : 0)]; };
      const gm = gsum('mi'), gc = gsum('ch'), gf = gsum('cf');
      const hvr = HVR.filter(h => h.nv === code);
      const cOff = os.filter(o => elig[o.c]?.has('off')), cOnNh = os.filter(o => elig[o.c]?.has('on') || elig[o.c]?.has('nh'));
      const evs = g => CEV.filter(e => e.nv === code && (Array.isArray(g) ? g.includes(e.g) : e.g === g));
      const cmbArr = (lst, g) => { const e = evs(g); const ks = new Set(e.map(z => z.c)); return [lst.length, sum(e, z => z.q), ks.size, sum([...ks], c => OS[c]?.ds || 0)]; };
      const cmbOff = cmbArr(cOff, 'off'), cmbOn = cmbArr(os.filter(o => elig[o.c]?.has('on')), 'on'), cmbNh = cmbArr(os.filter(o => elig[o.c]?.has('nh')), 'nh'), cmbOnNh = cmbArr(cOnNh, ['on', 'nh']);
      const fOff = chan(ls.filter(l => l.ch === 'OFF' && !XM(l.sub)));
      const k = {
        position: posOf[code] || '', code, rep: name, team, mcpOff, mcpOn, mcpTotal: mcpTot, mcpActive: kv(r, C.mcp, 3), mcpActiveOff: actOff, mcpActiveOn: actOn,
        target, actual, gap: actual - target, achievement: rt(actual, target), paceGap: rt(actual, target) - prog, dailyNeed: kv(r, C.s, 5),
        activeTarget: kv(r, C.aAll), activeActual: kv(r, C.aAll, 2), activeAchievement: rt(kv(r, C.aAll, 2), kv(r, C.aAll)),
        teaGreenTarget: kv(r, C.tx, TXO.ct), teaGreenActual: kv(r, C.tx, TXO.th), teaGreenAchievement: rt(kv(r, C.tx, TXO.th), kv(r, C.tx, TXO.ct)),
        homeyYellowTarget: kv(r, C.tv), homeyYellowActual: kv(r, C.tv, 2), homeyYellowAchievement: rt(kv(r, C.tv, 2), kv(r, C.tv)),
        pcTarget: kv(r, C.pOff), pcActual: kv(r, C.pOff, 3), pcBc: kv(r, C.pcT, 1), pcAchievement: rt(kv(r, C.pOff, 3), kv(r, C.pOff)),
        pcTotalTarget: kv(r, C.pcT), pcTotalActual: kv(r, C.pcT, 1),
        pcOffTarget: kv(r, C.pOff), pcOffActual: kv(r, C.pOff, 3), pcOffAchievement: rt(kv(r, C.pOff, 3), kv(r, C.pOff)), pcOffDaily: kv(r, C.pOff, 6),
        pcOnTarget: kv(r, C.pOn), pcOnActual: kv(r, C.pOn, 3), pcOnAchievement: rt(kv(r, C.pOn, 3), kv(r, C.pOn)),
        lppcSalaryPc: lpPc, lppcSalaryLine: lpLn, lppcSalary: rt(lpLn, lpPc), lppcPc: lpPc, lppcLine: lpLn, lppc: rt(lpLn, lpPc), lppcBq: rt(lpLn, lpPc),
        lppcMeatPc: lmPc, lppcMeatLine: Math.round(lmPc * lm), lppcMeat: lm,
        growthChanteTarget: gc[0], growthChanteActual: gc[1], growthChanteAchievement: rt(gc[1], gc[0]), growthChanteCustomers: gc[2], growthChanteEligible: gc[3],
        growthNoodleTarget: gm[0], growthNoodleActual: gm[1], growthNoodleAchievement: rt(gm[1], gm[0]), growthCoffeeTarget: gf[0], growthCoffeeActual: gf[1], growthCoffeeAchievement: rt(gf[1], gf[0]),
        comboTeaTarget: cmbOff[0], comboTeaOrders: cmbOff[1], comboTeaCustomers: cmbOff[2], comboTeaSales: cmbOff[3], comboChanteTarget: cmbOnNh[0], comboChanteOrders: cmbOnNh[1], comboChanteCustomers: cmbOnNh[2], comboChanteSales: cmbOnNh[3],
        comboOrders: cmbOff[1] + cmbOnNh[1], comboCustomers: cmbOff[2] + cmbOnNh[2], comboSales: cmbOff[3] + cmbOnNh[3],
        mbsCustomers: hvr.length, mbsReward: hvr.filter(h => /ĐẠT THƯỞNG/i.test(h.kq)).length, mbsTarget: sum(hvr, h => h.ct), mbsActual: sum(hvr, h => h.th), mbsAchievement: rt(sum(hvr, h => h.th), sum(hvr, h => h.ct)),
        foundationPc: fOff.orders, foundationLine: fOff.lines, foundationLppc: rt(fOff.lines, fOff.orders)
      };
      const SK = { Seasoning: 'Seasoning', 'Convenience Foods': 'Convenience', 'Home Care': 'HomeCare', 'Refreshment Drinks': 'Refreshment', 'Processed Meats': 'Processed', Nutrition: 'Nutrition', Coffee: 'Coffee', 'Rice base': 'Rice', Beer: 'Beer' };
      SUBS.forEach((s, i) => { if (!SK[s] || indC[i] < 0) return; const ct = num(r[indC[i]]), th = num(r[indC[i] + 1]); k['sellout' + SK[s] + 'Target'] = ct; k['sellout' + SK[s] + 'Actual'] = th; k['sellout' + SK[s] + 'Achievement'] = rt(th, ct); });
      KPIS.push(mk(TPL.M_kpis, k));
      // bonus
      const bv = BR[code]; TBON.push(bonusRec(code, name, TEAM, bv));
      const b = TBON[TBON.length - 1];
      MBONUS.push(mk(TPL.M_bonus, { code, rep: name, team, sauceAsoTarget: b.combo[0], sauceAsoActual: b.combo[1], sauceAsoAchievement: b.combo[2], sauceAsoReward: b.combo[3], homeyAsoTarget: b.homey[0], homeyAsoActual: b.homey[1], homeyAsoAchievement: b.homey[2], homeyAsoReward: b.homey[3], kkm90Target: b.mi[0], kkm90Actual: b.mi[1], kkm90Achievement: b.mi[2], kkm90Reward: b.mi[3], meatTarget: b.meat[0], meatActual: b.meat[1], meatAchievement: b.meat[2], meatEstimatedReward: b.tot, totalReward: b.sum[6], meatCurrySales: b.ds[0], meatCurryReward: b.rw[0], meatInnoSales: b.ds[1], meatInnoReward: b.rw[1], meatCoreSales: b.ds[2], meatCoreReward: b.rw[2], meatGroupReward: b.tot, meatMissing100: b.l100, meatMissing105: b.l105, meatStatus: b.lvl, kkm90SourceAvailable: { kkm90Target: true, kkm90Actual: true, kkm90Achievement: true, kkm90Target108: false, kkm90Reward: true } }));
      // monthly
      const hb = histBy[name] || {}; const months = [];
      for (let m = 1; m <= MON; m++) { if (m === MON) months.push({ month: `T${m}/${YEAR}`, target, actual, achievement: rt(actual, target), available: true }); else { const z = hb[m]; months.push({ month: `T${m}/${YEAR}`, target: z ? z[0] : 0, actual: z ? z[1] : 0, achievement: z ? rt(z[1], z[0]) : 0, available: !!z }); } }
      const av = months.filter(z => z.available); MMONTHLY.push({ rep: name, code, team, ward: wardOf[code], months, totalTarget: sum(av, z => z.target), totalActual: sum(av, z => z.actual), totalAchievement: rt(sum(av, z => z.actual), sum(av, z => z.target)), monthsAchieved: av.filter(z => z.achievement >= 1).length, growth: [] });
      // T.reps
      const hvDS = hvr.filter(h => /DAISU/i.test(h.grp)), hvV = hvr.filter(h => !/DAISU/i.test(h.grp));
      const lvl = h => rt(h.th, h.ct) >= 1.05 ? 3 : rt(h.th, h.ct) >= 1 ? 2 : rt(h.th, h.ct) >= .95 ? 1 : 0;
      const nv2ok = h => !/CHƯA/i.test(h.nh) && !!h.nh;
      TREPS.push(mk(TPL.T_reps, { pos: posOf[code] || '', code, name, team, ct: target, th: actual, cl: Math.max(0, target - actual), pct: rt(actual, target), mcp: { off: mcpOff, on: mcpOn }, cmb: { off: cmbOff, on: cmbOn, nh: cmbNh },
        hv: { tong: hvr.length, nv2: hvr.filter(nv2ok).length, nv1: hvr.filter(h => lvl(h) > 0).length, dat: hvr.filter(h => /ĐẠT THƯỞNG/i.test(h.kq)).length, nv2only: hvr.filter(h => nv2ok(h) && !lvl(h)).length, nv1only: hvr.filter(h => !nv2ok(h) && lvl(h)).length, zero: hvr.filter(h => h.th <= 0).length, ds_ct: sum(hvr, h => h.ct), ds_th: sum(hvr, h => h.th), pct: rt(sum(hvr, h => h.th), sum(hvr, h => h.ct)), nvip: hvDS.length, nvipOK: hvDS.filter(nv2ok).length, nvip2: hvV.length, ntotal: hvr.filter(h => /ĐẠT THƯỞNG/i.test(h.kq)).length },
        gk: { mi: [gm[0], gm[1], Math.max(0, gm[0] - gm[1]), rt(gm[1], gm[0]), gm[2], gm[3], gm[4]], ch: [gc[0], gc[1], Math.max(0, gc[0] - gc[1]), rt(gc[1], gc[0]), gc[2], gc[3], gc[4]], cf: [gf[0], gf[1], Math.max(0, gf[0] - gf[1]), rt(gf[1], gf[0]), gf[2], gf[3], gf[4]], all: [gm[0] + gc[0] + gf[0], gm[1] + gc[1] + gf[1], Math.max(0, gm[0] + gc[0] + gf[0] - gm[1] - gc[1] - gf[1]), rt(gm[1] + gc[1] + gf[1], gm[0] + gc[0] + gf[0]), new Set(gkr.map(z => z.n)).size, gm[3] + gc[3] + gf[3]] } }));
      // fund
      const grp = {}; const gk2 = o => o.ch === 'ON' ? 'ON' : (/VIP3/.test(o.t) ? 'VIP3' : /VIP5/.test(o.t) ? 'VIP5' : /SI|SỈ/i.test(o.t) ? 'VIPSI' : 'LE');
      os.forEach(o => { const g = gk2(o); const z = grp[g] = grp[g] || [0, 0, 0]; z[0]++; if (bought(o.c)) { z[1]++; z[2] += OS[o.c].ds; } });
      const onO = os.filter(o => o.ch === 'ON'), offO = os.filter(o => o.ch === 'OFF');
      const ordOn = new Set(ls.filter(l => l.ch === 'ON' && l.v > 0).map(l => l.id)).size, ordOff = new Set(ls.filter(l => l.ch === 'OFF' && l.v > 0).map(l => l.id)).size;
      const teaOn = onO.filter(o => OS[o.c]?.tea > 0).length;
      const onT = {}; onO.forEach(o => { const t = ON7.find(z => norm(z) === norm(o.l3) || norm(z) === norm(o.l2)) || (/quán ăn|nhà hàng/i.test(o.l3) ? 'Quán ăn' : 'Khác'); const z = onT[t] = onT[t] || [0, 0, 0]; z[0]++; if (bought(o.c)) z[1]++; if (OS[o.c]?.tea > 0) z[2]++; });
      const call = ch => { let v = 0, p = 0; lastDays.forEach(dd => { os.filter(o => o.ch === ch && onRoute(o, dd)).forEach(o => { v++; if (OS[o.c]?.byDay[dd]) p++; }); }); return [v, p]; };
      const ng = {}; SUBS.forEach(s => { const a = ls.filter(l => l.sub === s && !l.km && l.v > 0); if (a.length) ng[s] = { kh: new Set(a.map(l => l.c)).size, lines: a.length, orders: new Set(a.map(l => l.id)).size, ds: sum(a, l => l.v) }; });
      const ob = {}; const cg = ctG[code] || {}; SUBS.forEach(s => { const grs = new Set(Object.keys(cg[s] || {})); ls.filter(l => l.sub === s && l.v > 0).forEach(l => grs.add(l.g)); if (!grs.size) return; ob[s] = {}; grs.forEach(g => { const a = ls.filter(l => l.sub === s && l.g === g && l.v > 0); ob[s][g] = { ct: (cg[s] || {})[g] || 0, th: sum(a, l => l.v), kh: new Set(a.map(l => l.c)).size }; }); });
      TFUND.push(mk(TPL.T_fund, { name, code, team, ct: target, th: actual, mcp: [mcpOff, mcpOn, mcpTot], act: [actOff + actOn, actOff, actOn], call: { off: call('OFF'), on: call('ON') }, all: chan(ls), xm: chan(ls.filter(l => !XM(l.sub))), off: chan(ls.filter(l => l.ch === 'OFF' && !XM(l.sub))), offa: chan(ls.filter(l => l.ch === 'OFF')), offall: chan(ls.filter(l => l.ch === 'OFF')), on: chan(ls.filter(l => l.ch === 'ON')), ng, ob, grp, onoff: { mcpOn: onO.length, asoOn: onO.filter(o => bought(o.c)).length, ordOn, teaOn, on1: 0, mcpOff: offO.length, buyOff: offO.filter(o => bought(o.c)).length, ordOff }, onT }));
      // ov
      const vip = { VIPSI: grp.VIPSI || [0, 0, 0], VIP3: grp.VIP3 || [0, 0, 0], VIP5: grp.VIP5 || [0, 0, 0], LE: grp.LE || [0, 0, 0], ON: grp.ON || [0, 0, 0] };
      const ont = {}; ON7.forEach(t => { const a = onO.filter(o => norm(o.l3) === norm(t) || norm(o.l2) === norm(t)); ont[t] = [a.length, a.filter(o => bought(o.c)).length]; });
      TOV[name] = { tree: tree(code, ls), vip, onoff: { on: [onO.length, onO.filter(o => bought(o.c)).length, ordOn, teaOn, 0], off: [offO.length, offO.filter(o => bought(o.c)).length, ordOff] }, ont, ontAso: sum(Object.values(ont), z => z[1]) };
      // ovW theo phường
      const wards = [...new Set(os.map(o => o.w))].filter(Boolean);
      wards.forEach(w => { const ow = os.filter(o => o.w === w), cs = new Set(ow.map(o => o.c)); const lw = ls.filter(l => cs.has(l.c)); if (!ow.length) return; const og = {}; ow.forEach(o => { const g = gk2(o); const z = og[g] = og[g] || [0, 0, 0]; z[0]++; if (bought(o.c)) { z[1]++; z[2] += OS[o.c].ds; } });
        const onW = ow.filter(o => o.ch === 'ON'), offW = ow.filter(o => o.ch === 'OFF'); const ontW = {}; ON7.forEach(t => { const a = onW.filter(o => norm(o.l3) === norm(t) || norm(o.l2) === norm(t)); ontW[t] = [a.length, a.filter(o => bought(o.c)).length]; });
        const ch3 = a => { const c = chan(a); return [c.pc, c.lines, c.ds]; }; const cats = {}; lw.forEach(l => { if (l.v > 0) cats[l.sub] = (cats[l.sub] || 0) + l.v; });
        OVW.push({ w, rep: name, tree: tree(null, lw), vip: { VIPSI: og.VIPSI || [0, 0, 0], VIP3: og.VIP3 || [0, 0, 0], VIP5: og.VIP5 || [0, 0, 0], LE: og.LE || [0, 0, 0], ON: og.ON || [0, 0, 0] }, onoff: { on: [onW.length, onW.filter(o => bought(o.c)).length, new Set(lw.filter(l => l.ch === 'ON' && l.v > 0).map(l => l.id)).size, onW.filter(o => OS[o.c]?.tea > 0).length, 0], off: [offW.length, offW.filter(o => bought(o.c)).length, new Set(lw.filter(l => l.ch === 'OFF' && l.v > 0).map(l => l.id)).size] }, ont: ontW, ontAso: sum(Object.values(ontW), z => z[1]), chn: { all: ch3(lw), xm: ch3(lw.filter(l => !XM(l.sub))), offall: ch3(lw.filter(l => l.ch === 'OFF')), off: ch3(lw.filter(l => l.ch === 'OFF' && !XM(l.sub))), on: ch3(lw.filter(l => l.ch === 'ON')), meat: ch3(lw.filter(l => l.sub === 'Processed Meats')) }, cats, xanh: 0, vang: new Set(lw.filter(isTV).map(l => l.c)).size, tea: new Set(lw.filter(l => isTea(l) && l.v > 0).map(l => l.c)).size }); });
      // mtn reps
      const day = {}, mix = {}; DDS.forEach(dd => { const a = ls.filter(l => l.dd === dd && !l.km && l.v > 0); if (!a.length) return; day[dd] = sum(a, l => l.v); const ngd = {}, nh = {}; a.forEach(l => { ngd[l.sub] = (ngd[l.sub] || 0) + l.v; const b = l.br !== 'Khác' ? l.br : brandOf(l.g); nh[b] = (nh[b] || 0) + l.v; }); const ords = new Set(a.map(l => l.id));
        mix[dd] = { ng: ngd, nh, ord: ords.size, ordx: new Set(a.filter(l => !XM(l.sub)).map(l => l.id)).size, ln: a.length, pc: new Set(a.map(l => l.c)).size, pon: new Set(a.filter(l => l.ch === 'ON').map(l => l.id)).size, tv: new Set(a.filter(isTV).map(l => l.c)).size, tx: 0 }; });
      MTNREPS.push({ name, code, ss: TEAM, ct: target, th: actual, day, mix });
      // fx
      const fd = {}; DDS.forEach(dd => { const a = ls.filter(l => l.dd === dd && !l.km && l.v > 0); fd[`${dd}/${p2(MON)}`] = [new Set(a.map(l => l.c)).size, new Set(a.filter(isTV).map(l => l.c)).size, CEV.filter(e => e.nv === code && e.dd === dd).length, new Set(a.filter(l => isTea(l) && l.ch === 'ON').map(l => l.c)).size, new Set(a.filter(l => isTea(l) && l.ch === 'OFF').map(l => l.c)).size]; });
      FXREPS.push({ name, code, ss: TEAM, row: ri + 5, nOn: mcpOn, nOff: mcpOff, days: fd });
      const n7 = onO.filter(o => ON7.some(t => norm(t) === norm(o.l3) || norm(t) === norm(o.l2))); const pri = onO.filter(o => /quán nước|quốc lộ|thể thao|vận chuyển/i.test(o.l3 + ' ' + o.l2));
      FXON.push({ name, ss: TEAM, n: onO.length, aso: onO.filter(o => bought(o.c)).length, n7: n7.length, aso7: n7.filter(o => bought(o.c)).length, tea7: n7.filter(o => OS[o.c]?.tea > 0).length, teaOn, np: pri.length, asop: pri.filter(o => bought(o.c)).length, pcCt: kv(r, C.pOn), pcTh: kv(r, C.pOn, 3), pcCl: Math.max(0, kv(r, C.pOn) - kv(r, C.pOn, 3)), pcPct: rt(kv(r, C.pOn, 3), kv(r, C.pOn)) });
      FXX.push({ name, ss: TEAM, asoCt: kv(r, C.tx, TXO.ct), asoTh: kv(r, C.tx, TXO.th), chinsu: [0, 0, 0], omc: [0, 0, 0], vip3: [0, 0, 0], vip5: [0, 0, 0] });
      const nb = os.filter(o => o.ds3 <= 0); nb.forEach(o => nbList.push({ c: o.c, n: o.n, rep: name, ch: o.ch.toLowerCase(), l2: o.l2, w: o.w, a: o.a, cr: o.cr ? `${o.cr[2]}-${p2(o.cr[1])}-${p2(o.cr[0])}` : '', ty: o.t, tg: o.tg, th: OS[o.c]?.ds || 0, vis: o.v }));
      FXNB.push({ name, ss: TEAM, off: nb.filter(o => o.ch === 'OFF').length, on: nb.filter(o => o.ch === 'ON').length, offBuy: nb.filter(o => o.ch === 'OFF' && bought(o.c)).length, onBuy: nb.filter(o => o.ch === 'ON' && bought(o.c)).length, offMcp: offO.length, onMcp: onO.length });
      // LE
      const pleT = leDd[norm(name)] || 1; const ngLe = {}; let tct = 0, tth = 0, tle = 0;
      ['Seasoning', 'Convenience Foods', 'Home Care', 'Refreshment Drinks', 'Processed Meats', 'Nutrition', 'Coffee', 'Beer'].forEach(s => { const i = SUBS.indexOf(s); const ct = num(r[indC[i]]), th = num(r[indC[i] + 1]); const ple = leInd[s] || pleT; const le = ct * ple; tct += ct; tth += th; tle += le; ngLe[s] = { ct, th, pth: rt(th, ct), ple, le, cl: Math.max(0, le - th) }; });
      ngLe.TOTAL = { ct: target, th: actual, pth: rt(actual, target), ple: rt(tle, tct), le: tle, cl: Math.max(0, tle - actual) };
      LEROWS.push({ rep: name, code, ng: ngLe });
    });

    // ===== khách hàng =====
    const MMCP = [], TCUST = [], MCOMBO = [], KHL = [];
    const parity = dd => isoWk([+dd, MON, YEAR]) % 2 ? 'Odd Week' : 'Even Week';
    outs.forEach(o => {
      const z = OS[o.c]; const repN = byCode[o.nv].name; const pg = [...(elig[o.c] || [])]; const got = {}; (cevBy[o.c] || []).forEach(e => { const g = got[e.g] = got[e.g] || [0, z ? z.ds : 0]; g[0] += e.q; });
      const ordDates = z ? [...z.dd].sort().join(',') : '';
      MMCP.push(mk(TPL.M_mcp, { customerCode: o.c, customer: o.n, address: o.a, ward: o.w, type: o.t, storePosition: o.vt, channelPrimary: o.l1, channelSecondary: o.l2, storeType: o.l3, channel: o.ch, team, rep: repN, routeDay: o.v, threeMonthSales: o.ds3, unpurchased3Months: o.ds3 <= 0, unpurchasedT1: false, unpurchasedT2: false, t1Status: null, t2Status: null, birthdayCurrentMonth: BD[o.c]?.m === MON, birthdaySeptember: false, birthdayNote: null, activeTotal: z ? z.ds : 0, pc: z ? z.ord.size : 0, pc200: z ? Object.values(z.ordV).filter(v => v >= 2e5).length : 0, totalSales: z ? z.ds : 0, purchased: bought(o.c), tea365Sales: z ? z.tea : 0, tea365Purchased: !!(z && z.tea > 0), ordCount: z ? z.ord.size : 0, ordDays: z ? z.dd.size : 0, ordLast: z && z.last ? `${YEAR}-${p2(MON)}-${z.last}` : '', ordDates, freq: o.f, oddWeek: o.p === 'Both' ? '' : (o.p || 'Chưa rõ') }));
      TCUST.push({ c: o.c, n: o.n, rep: repN, team, visit: o.v, freq: o.f, type: o.t, ch: o.l1, l3: o.l3, tgt: o.tg, prog: pg, got, ds: z ? z.ds : 0, od: z ? z.dd.size : 0, addr: o.a, ward: o.w, pc: z ? z.ord.size : 0 });
      if (pg.length) MCOMBO.push(mk(TPL.M_combo, { customerCode: o.c, customer: o.n, address: o.a, ward: o.w, routeDay: o.v, rep: repN, team, npp: NPP, dmsLoc: '', comboOff: pg.includes('off'), comboOn: pg.includes('on'), comboNh: pg.includes('nh'), newOrder200k: !!(z && Object.values(z.ordV).some(v => v >= 2e5)), coreDetail: '' }));
      const oo = {}; if (z) for (const dd in z.byDay) oo[dd] = z.byDay[dd]; const cb = {}; (cevBy[o.c] || []).forEach(e => { const z = cb[e.dd] = cb[e.dd] || [0, 0, e.g]; z[0]++; z[1] += e.q; });
      const teaD = [...new Set(L.filter(l => l.c === o.c && isTea(l) && l.v > 0).map(l => l.dd))].sort();
      KHL.push({ c: o.c, n: o.n, r: repN, v: o.v, p: o.p, t: o.t, w: o.w, a: o.a, ch: o.ch.toLowerCase(), l7: ON7.some(t => norm(t) === norm(o.l3) || norm(t) === norm(o.l2)) && o.ch === 'ON' ? 1 : 0, tea: teaD, rr: o.ds3 / 3, hv: HVR.some(h => h.c === o.c) ? 1 : 0, pg, o: oo, cb });
    });

    // ===== hội viên =====
    const MMBS = [], THV = [];
    HVR.forEach(h => { const o = OUT[h.c] || {}; const repN = byCode[h.nv].name; const pct = rt(h.th, h.ct); const lvl = pct >= 1.05 ? 3 : pct >= 1 ? 2 : pct >= .95 ? 1 : 0; const ds = /DAISU/i.test(h.grp);
      const mark = (n, v) => `${n} ${ds ? (v >= 15e4 ? '✓' : 'X') : (v > 0 ? '✓' : 'X')}`; const det = [mark('Nam Ngư', h.nn), mark('Chinsu', h.cs), mark('Kokomi', h.kk), mark('Chanté', h.ch)].join(' | ');
      const nv2 = !/CHƯA/i.test(h.nh) && !!h.nh; const cap = SKUS.map(s => { const c = (capBy[h.c] || {})[s]; return c ? [c[0], c[1], c[2]] : [0, 0, 0]; });
      const ov = Object.entries(capBy[h.c] || {}).filter(([, v]) => v[1] > v[0]).map(([k]) => k), nr = Object.entries(capBy[h.c] || {}).filter(([, v]) => v[1] <= v[0] && v[1] > 0 && v[1] >= .9 * v[0]).map(([k]) => k);
      const dsNo = sum(Object.values(capBy[h.c] || {}), v => Math.max(0, v[3])); const capW = (ov.length ? 'VƯỢT cap: ' + ov.join(', ') + '.' : '') + (nr.length ? ' SẮP CHẠM: ' + nr.join(', ') + '.' : '');
      const thuong = h.pt / 100; const ok = /ĐẠT THƯỞNG/i.test(h.kq);
      THV.push({ ty: o.t || '', c: h.c, n: h.n, visit: h.v, rep: repN, team, grp: h.grp, ct: h.ct, th: h.th, m3: h.ct * 1.05, pct, lvl, nn: h.nn, cs: h.cs, kk: h.kk, ch: h.ch, b4: h.b4, nv2: nv2 ? 'ĐẠT' : 'CHƯA ĐẠT', dsTT: h.tt, kq: ok ? 'ĐẠT' : 'CHƯA ĐẠT', thuong, thieu: h.ly, nv1: lvl > 0 ? 1 : 0, nv2f: nv2 ? 1 : 0, dat: ok ? 1 : 0, capW: capW.trim(), nOv: ov.length, nNr: nr.length, dsNo, cap });
      MMBS.push(mk(TPL.M_mbs, { customerCode: h.c, customer: h.n, routeDay: h.v, rep: repN, team, address: h.a, ward: o.w || '', memberType: h.grp, target: h.ct, actual: h.th, gap: Math.max(0, h.ct - h.th), achievement: pct, salesStatus: lvl ? 'ĐẠT' : 'CHƯA ĐẠT', maintenanceDetail: det, maintenanceSales: h.b4, maintenanceRequired: 1, maintenanceAchieved: nv2 ? 1 : 0, maintenance: nv2 ? 'ĐẠT NV2' : 'CHƯA ĐẠT NV2', newBrandDetail: det, newBrandSales: h.b4, newBrand: '—', m3: h.ct * 1.05, lvl, nn: h.nn, cs: h.cs, kk: h.kk, chn: h.ch, b4: h.b4, nv2: nv2 ? 'ĐẠT' : 'CHƯA ĐẠT', dsTT: h.tt, nv1: lvl > 0 ? 1 : 0, nv2f: nv2 ? 1 : 0, rewardStatus: ok ? 'ĐẠT' : 'CHƯA ĐẠT', rewardRate: [0, .012, .015, .02][lvl], estimatedReward: thuong, missing: h.ly, gameApplicable: false, salesEligible: lvl > 0, catBrandEligible: nv2, rewardEligible: ok })); });

    // ===== tăng trưởng (gk) =====
    const MGR = [], TGK = [];
    Object.entries(GK).forEach(([c, z]) => { const o = OUT[c] || {}; const repN = byCode[z.nv].name; const g = k => z[k] || [0, 0, 0, 0, 0];
      const all = [g('mi')[0] + g('ch')[0] + g('cf')[0], g('mi')[1] + g('ch')[1] + g('cf')[1]]; all.push(rt(all[1], all[0]), g('mi')[3] || g('ch')[3] || g('cf')[3] ? 1 : 0);
      TGK.push({ c, n: z.n, rep: repN, team, visit: o.v || '', mi: g('mi').slice(0, 4), ch: g('ch').slice(0, 4), cf: g('cf').slice(0, 4), all });
      MGR.push(mk(TPL.M_growth, { customerCode: c, customer: z.n, rep: repN, address: o.a || '', ward: o.w || '', routeDay: o.v || '', team, noodleTarget: g('mi')[0], noodleActual: g('mi')[1], noodleGap: Math.max(0, g('mi')[0] - g('mi')[1]), noodleAchievement: g('mi')[3], chanteTarget: g('ch')[0], chanteActual: g('ch')[1], chanteGap: Math.max(0, g('ch')[0] - g('ch')[1]), chanteAchievement: g('ch')[3], coffeeTarget: g('cf')[0], coffeeActual: g('cf')[1], coffeeGap: Math.max(0, g('cf')[0] - g('cf')[1]), coffeeAchievement: g('cf')[3] })); });

    // ===== trưng bày =====
    const MDISP = [], TDP = [];
    DP.forEach(d => { const repN = byCode[d.nv].name; const o = OUT[d.c] || {}; const six = d.ok >= 6 ? 1 : 0;
      TDP.push({ prog: d.p, lvl: d.m, c: d.c, n: d.n, rep: repN, team, visit: d.v, tgt: d.tg, th: d.th, max: 0, shots: d.shots, ok: d.ok, src: d.p, cap: d.shots > 0 ? 1 : 0, six, el: d.tg > 0 ? 1 : 0, thd: d.th, tgd: d.tg, addr: d.a, ward: d.w || o.w || '' });
      MDISP.push(mk(TPL.M_display, { programCode: d.pc, program: d.p, level: d.m, customerCode: d.c, customer: d.n, address: d.a, ward: d.w || o.w || '', routeDay: d.v, rep: repN, target: d.tg, actual: d.th, gap: d.th - d.tg, maxReward: 0, photoCaptured: d.shots, photoRepCaptured: d.shots, photoApproved: d.ok, photoSixStatus: six ? 'ĐẠT' : '', sourceProgram: d.p, team, targetStatus: d.th >= d.tg && d.tg > 0 ? 'ĐẠT' : '', earnedReward: 0, sourceEligible: d.tg > 0, sourceRewardPoints: 0 })); });

    // ===== bonus tổng =====
    const ssB = bonusRec(null, 'SS ' + TEAM.toUpperCase(), TEAM, BR[Object.keys(BR).find(k => /^SS/i.test(k) && new RegExp(teamKey, 'i').test(k))]);
    const nppB = bonusRec(null, NPP + ' — TỔNG NPP', null, BR[Object.keys(BR).find(k => norm(k) === norm(NPP))]);
    const notes = { mi: BN.mi || TPL.T_bonusNotes.mi, meat: BN.meat || TPL.T_bonusNotes.meat, combo: BN.combo || TPL.T_bonusNotes.combo, homey: BN.homey || '' };

    // ===== meta =====
    const meta = JSON.parse(JSON.stringify(TPL.meta));
    const rw = (k, d) => { const m = hdr(k).match(/([\d.,]+)\s*(tr|K)/i); if (!m) return d; const v = /k/i.test(m[2]) ? num(m[1].replace(/\./g, '')) * 1e3 : num(m[1].replace(',', '.')) * 1e6; return v.toLocaleString('vi-VN') + 'đ'; };
    const salesRw = (hdr('s').match(/([\d.,]+)\s*tr/i) || [0, '5.6'])[1]; const salesV = (num(salesRw.replace(',', '.')) * 1e6).toLocaleString('vi-VN') + 'đ';
    const aThr = +(hdr('aAll').match(/>\s*(\d+)\s*%/) || [0, 80])[1]; const lCt = +(hdr('lp').match(/Mức\s*([\d.]+)/i) || [0, 4.7])[1]; const mCt = +(hdr('lm').match(/Mức\s*([\d.]+)/i) || [0, 3.8])[1];
    Object.assign(meta, { title: `BÁO CÁO KPI THÁNG ${p2(MON)}/${YEAR}  •  Cập nhật: ${today}`, updated: today, team: TEAM, bonusTeam: TEAM, customerScope: 'Kênh ON + OFF', mcpOffCustomers: sum(REPS, x => kv(x.r, C.mcp, 0)), mcpOnCustomers: sum(REPS, x => kv(x.r, C.mcp, 1)), excludedOnCustomers: 0,
      growthPrograms: [{ key: 'noodle', label: `Ngành Mì (≥${Math.round(gp('mi')[0] * 100)}%)` }, { key: 'chante', label: `Chanté (≥${Math.round(gp('ch')[0] * 100)}%)` }, { key: 'coffee', label: `Ngành Cà phê (≥${Math.round(gp('cf')[0] * 100)}%)` }],
      mbsRules: 'Hội viên T11: NV1 = DS total Masan đạt Mức 1/2/3 (95% / 100% / 105% chỉ tiêu nền) • NV2 = nhãn trọng điểm: Đại sứ ≥150.000đ/nhãn, VIP còn lại tổng 4 nhãn ≥500.000đ • Cap SKU trọng yếu 115% (sheet Cap Tran)',
      foundationSource: `sheet KPI THÁNG T${MON} • Team SS Văn • map theo tiêu đề cột`,
      foundationMetrics: [
        { label: 'Doanh số & tiến độ', targetKey: 'target', actualKey: 'actual', unit: 'money', reward: `Đạt 100% • Lương ${salesV}` },
        { label: 'Active ASO MCP OFF + ON', targetKey: 'activeTarget', actualKey: 'activeActual', unit: 'number', reward: `Đạt trên ${aThr}% • Lương ${rw('aAll', '1.040.000đ')}` },
        { label: 'ASO Trận Xanh • Điểm lẻ bao phủ SP Tết', targetKey: 'teaGreenTarget', actualKey: 'teaGreenActual', unit: 'number', reward: `Đạt 100%: ${rw('tx', '1.000.000đ')}` },
        { label: 'ASO Trận Vàng • Hạt nêm Chinsu Ngọt Thanh', targetKey: 'homeyYellowTarget', actualKey: 'homeyYellowActual', unit: 'number', reward: `Đạt 100%: ${rw('tv', '800.000đ')}` },
        { label: 'PC OFF • 65% MCP', targetKey: 'pcTarget', actualKey: 'pcActual', unit: 'number', reward: `Đạt 100% • Lương ${rw('pOff', '2.200.000đ')}` },
        { label: 'LPPC trừ Meat & Beer', targetValue: lCt, actualKey: 'lppcSalary', aggregation: 'average', unit: 'decimal', reward: `Mức 1: ≥${String(lCt).replace('.', ',')} • ${rw('lp', '2.000.000đ')}` },
        { label: 'LPPC Meat', targetValue: mCt, actualKey: 'lppcMeat', aggregation: 'average', unit: 'decimal', reward: `Mức 1: LPPC Meat từ ${String(mCt).replace('.', ',')} • ${rw('lm', '400.000đ')}` },
        { label: 'PC ON • Tất cả đơn hàng ON', targetKey: 'pcOnTarget', actualKey: 'pcOnActual', unit: 'number', reward: `Đạt 100% • Lương ${rw('pOn', '1.500.000đ')}` }],
      workingDays: wd, elapsedDays: passed, progress: prog, periodKey: MK, periodLabel: `Tháng ${p2(MON)}/${YEAR}`, monthNumber: MON, locked: false, sourceFile: opt.sourceFile || 'T11_26_Tracking_KPI.xlsb', displayUpdated: today, comboUpdated: today, salesUpdated: today, mbsUpdated: today, growthUpdated: today,
      salesSource: 'KPI THÁNG • cột TH nhóm DOANH SỐ BÁN RA • đối chiếu dòng SS Võ Viết Văn',
      salesReconciliation: { kpiTotal: sum(REPS, x => kv(x.r, C.s, 1)), mcpTotal: sum(Object.values(OS), z => z.ds), difference: sum(REPS, x => kv(x.r, C.s, 1)) - sum(Object.values(OS), z => z.ds), note: 'Tổng quan dùng doanh số KPI THÁNG; doanh số từng quầy tính từ sheet 061 new (không gồm hàng KM, đơn hủy).' },
      snapshotNote: `Tháng ${MON} có đơn hàng ngày ${DDS[0] || ''}–${DDS[DDS.length - 1] || ''}/${p2(MON)}. Tracking Team SS Võ Viết Văn theo file Tracking KPI T${MON}.`,
      comboPrograms: [{ key: 'comboOff', label: 'Combo OFF VIP1', sourceLabel: 'Bill ≥800K, có ≥100K 4 nhãn (Chinsu, Nam Ngư, Kokomi, Chanté)' }, { key: 'comboOn', label: 'Combo ON Tea/Amazon', sourceLabel: 'Bill BEV ≥250K + ≥80K Tea365 / Wake up 247 Amazon' }, { key: 'comboNh', label: 'Combo Nhà hàng VIP2 · VIP3', sourceLabel: 'Gia vị ≥500K + 200K Nam Ngư STK (hoặc 400K + 32K TTT Nhị Ca)' }],
      teams: [TEAM], scopeLabel: 'Team SS Văn', multiTeam: false, detailAvailability: { display: { [TEAM]: true }, mbs: { [TEAM]: true }, growth: { [TEAM]: true } },
      bonusScope: 'Team SS Văn', displaySource: 'T.Bày • danh sách KH trưng bày • BỘ CHỤP / BỘ ĐẠT', foundationRecalculation: { reason: 'PC, Line và LPPC nền tảng tính từ sheet 061 new (kênh OFF, trừ Meat & Bia).', sourceColumns: {}, fields: [], reps: [] },
      ordAsOf: `${YEAR}-${p2(MON)}-${DDS[DDS.length - 1] || p2(td[0])}`, pcOnNote: 'PC ON lấy từ KPI THÁNG (cột PC ON).', comboNames: ['Combo OFF VIP1', 'Combo ON + Nhà hàng'], teaGreenLabel: 'ASO Trận Xanh • Điểm lẻ bao phủ SP Tết', bonusSchema: TPL.meta.bonusSchema,
      bonusPrograms: [{ key: 'kkm90', label: 'Thưởng ngành Mì (CF)', target: 'kkm90Target', actual: 'kkm90Actual', achievement: 'kkm90Achievement', unit: 'money', icon: 'chart' }, { key: 'meat', label: 'Doanh số Meat • 3 nhóm', target: 'meatTarget', actual: 'meatActual', achievement: 'meatAchievement', unit: 'money', icon: 'trophy' }, { key: 'sauceAso', label: 'ASO combo xốt Mỹ Vị Chinsu', target: 'sauceAsoTarget', actual: 'sauceAsoActual', achievement: 'sauceAsoAchievement', unit: 'number', icon: 'layers' }, { key: 'homeyAso', label: 'ASO Nước mắm cao cấp', target: 'homeyAsoTarget', actual: 'homeyAsoActual', achievement: 'homeyAsoAchievement', unit: 'number', icon: 'gift' }] });
    const bonusRules = { kkm90: BN.mi || '', meat: BN.meat || '', sauceAso: BN.combo || '', homeyAso: BN.homey || '', premiumFishSauce: '', cooler: '' };

    const MONTHS = { [MK]: { meta, kpis: KPIS, growth: MGR, mbs: MMBS, display: MDISP, mcp: MMCP, aso: KPIS.map((k, i) => Object.assign({}, k, { ward: wardOf[k.code], lppcLevel: k.lppcSalary >= lCt ? 1 : 0 })), bonus: MBONUS, bonusRules, monthly: MMONTHLY, combo: MCOMBO, comboOn: [], leSheet: { labels: TPL.leSheetLabels, order: TPL.leOrder, rows: LEROWS, source: 'Sheet LE (%LE ngành hàng × CT KPI THÁNG)' } } };

    // ===== T11 =====
    const allDays = []; for (let d = 1; d <= new Date(Date.UTC(YEAR, MON, 0)).getUTCDate(); d++) { const w = dowOf([d, MON, YEAR]); if (w !== 'CN') allDays.push({ dd: p2(d), wd: w, par: parity(d) }); }
    const T11 = { meta: { month: MK, label: `Tháng ${p2(MON)}/${YEAR}`, updated: today, workingDays: wd, elapsed: passed, progress: prog, combo: TPL.T_meta.combo },
      reps: TREPS, capmeta: { sku: SKUS, note: `TRACKING CẢNH BÁO CAP TRẦN: ${THV.filter(h => h.nOv).length} KH có SKU VƯỢT cap • ${THV.filter(h => !h.nOv && h.nNr).length} KH chỉ SẮP CHẠM cap • tổng DS vượt cap ${Math.round(sum(THV, h => h.dsNo)).toLocaleString('en-US')} đ` },
      cev: CEV.map(e => ({ c: e.c, rep: e.rep, d: e.d, g: e.g, q: e.q, v: e.v })), cevEst: [], cust: TCUST, hv: THV, gk: TGK, dp: TDP, gpar: GPAR,
      bonus: { reps: TBON, ss: ssB, npp: nppB, notes }, fund: TFUND, ov: TOV, ovmeta: { meatPc: sum(TFUND, f => kv(byCode[f.code].r, C.lm)), meatLine: sum(KPIS, k => k.lppcMeatLine) },
      fx: { dates: DDS.map(dd => `${dd}/${p2(MON)}`), reps: FXREPS, on: FXON, xanh: { reps: FXX, list: [] }, nb: { reps: FXNB, list: nbList } },
      mtn: { days: allDays, reps: MTNREPS, kh: KHL, wd }, ovW: OVW, fxScope: TEAM, onTypes: ON7 };

    // ===== MONTHLY_DETAIL =====
    const sales = [];
    KH.forEach(([d, a, i, m, t, v]) => { const n = H.K.D[d]; sales.push({ month: `${YEAR}-${p2(m)}`, rep: n, code: (REPS.find(x => x.name === n) || {}).code || '', ward: H.K.A[a], sub: H.K.I[i], target: t * 1000, actual: v * 1000, source: 'Lịch sử T1–T10 (dashboard Team Văn)', team }); });
    REPS.forEach(x => SUBS.forEach((s, i) => { if (indC[i] < 0) return; const ct = num(x.r[indC[i]]), th = num(x.r[indC[i] + 1]); if (ct || th) sales.push({ month: MK, rep: x.name, code: x.code, ward: wardOf[x.code], sub: s, target: ct, actual: th, source: 'KPI THÁNG', team }); }));
    const periods = []; for (let m = 1; m <= MON; m++) periods.push(`${YEAR}-${p2(m)}`);
    const MD = Object.assign(JSON.parse(JSON.stringify(TPL.MD_rest)), { sales, aso: [], brandSales: [], periods, missingSalesMonths: periods.filter(p => !sales.some(s => s.month === p)), asoMissingMonths: periods, notes: ['ASO theo tháng chưa có nguồn trong file của Team Văn.'], source: 'T01–T10: số đã khóa (dashboard Team Văn) / T11: KPI THÁNG', scopeNote: 'Chỉ Team SS Văn • Không gồm NVBH vãng lai', brandSourceNote: 'Chưa có dữ liệu nhãn hàng theo tháng.', brandMappingNote: '' });

    const BDAY = outs.filter(o => BD[o.c]).map(o => ({ c: o.c, n: o.n, a: o.a, r: byCode[o.nv].name, t: o.t, b: BD[o.c].b, u: BD[o.c].u ? `${BD[o.c].u[2]}-${p2(BD[o.c].u[1])}-${p2(BD[o.c].u[0])}` : '' }));
    // quầy chưa có ngày SN
    outs.filter(o => !BD[o.c]).forEach(o => BDAY.push({ c: o.c, n: o.n, a: o.a, r: byCode[o.nv].name, t: o.t, b: '', u: '' }));
    const DAILY = { month: MK, asOf: meta.ordAsOf, subs: SUBS, dms: [], o: [], cb: {}, prog: {} };
    let TX = null; try { TX = tranXanh(A('Trận Xanh')); } catch (e) { warn.push('Không đọc được sheet Trận Xanh: ' + (e && e.message)); }
    if (!TX) warn.push('Không thấy sheet Trận Xanh — tab Trận Xanh sẽ trống');
    return { KPI_MONTHS: MONTHS, KPI_DAILY: DAILY, KPI_MONTHLY_DETAIL: MD, T11_DATA: T11, BDAY, TRAN_XANH: TX, warn, info: { reps: REPS.length, team: TEAM, npp: NPP, today } };
  }
  function readWorkbook(XLSX, fixXlsb, buf) { return XLSX.read(fixXlsb(XLSX, buf), { type: 'array', sheets: SHEETS, cellFormula: false, cellHTML: false, cellText: false, cellStyles: false }); }
  const api = { convert, readWorkbook, SHEETS };
  if (typeof module !== 'undefined') module.exports = api; else root.KPIConvert = api;
})(typeof window !== 'undefined' ? window : globalThis);

