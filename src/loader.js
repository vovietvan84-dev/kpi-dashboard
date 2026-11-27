(function(){
  if (window.__kpiLoader) return; window.__kpiLoader = true;
  var css = '.kl-fab{position:fixed;left:calc(16px + env(safe-area-inset-left,0px));bottom:calc(16px + env(safe-area-inset-bottom,0px));z-index:99991;display:flex;gap:8px;align-items:center}'+
    '.kl-btn{background:#163a7a;color:#fff;border:0;border-radius:999px;padding:11px 16px;font:600 14px/1 system-ui,Segoe UI,Roboto,sans-serif;box-shadow:0 4px 14px rgba(0,0,0,.25);cursor:pointer}'+
    '.kl-btn.alt{background:#fff;color:#163a7a;border:1px solid #c7d4f0}'+
    '.kl-ov{position:fixed;inset:0;z-index:99999;background:rgba(10,20,45,.6);display:grid;place-items:center;font:600 15px/1.5 system-ui,Segoe UI,Roboto,sans-serif}'+
    '.kl-box{background:#fff;color:#1f2a44;border-radius:16px;padding:22px 24px;max-width:420px;width:calc(100% - 32px);box-shadow:0 20px 60px rgba(0,0,0,.35)}'+
    '.kl-box b{display:block;font-size:17px;margin-bottom:6px;color:#1e3a8a}.kl-box small{display:block;color:#64708a;font-weight:500;margin-top:8px}'+
    '.kl-box .kl-err{color:#c62828}.kl-box button{margin-top:14px}';
  var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
  var fab = document.createElement('div'); fab.className = 'kl-fab';
  fab.innerHTML = '<button class="kl-btn" type="button" id="klLoad">⇪ Nạp file mới</button>' + (window.__kplUploaded ? '<button class="kl-btn alt" type="button" id="klReset">↺ Bản gốc</button>' : '');
  var inp = document.createElement('input'); inp.type = 'file'; inp.accept = '.xlsb,.xlsx,.xlsm'; inp.style.display = 'none';
  fab.style.display = 'none'; /* Ẩn nút Nạp file mới theo yêu cầu (27/11) — cập nhật số qua GitHub */
  document.body.appendChild(fab); document.body.appendChild(inp);
  function box(html){ var o = document.getElementById('klOv'); if(!o){ o = document.createElement('div'); o.id = 'klOv'; o.className = 'kl-ov'; document.body.appendChild(o); } o.innerHTML = '<div class="kl-box">' + html + '</div>'; return o; }
  function closeBox(){ var o = document.getElementById('klOv'); if (o) o.remove(); }
  document.getElementById('klLoad').onclick = function(){ inp.click(); };
  var rb = document.getElementById('klReset'); if (rb) rb.onclick = function(){ if (window.frameElement) window.frameElement.remove(); else location.reload(); };
  function loadXLSX(){ return new Promise(function(res, rej){ if (window.XLSX && window.XLSX.CFB) return res(); var s = document.createElement('script'); s.src = 'xlsx.full.min.js'; s.onload = function(){ res(); }; s.onerror = function(){ rej(new Error('Không tải được thư viện đọc Excel.')); }; document.head.appendChild(s); }); }
  function txt(id){ var el = document.getElementById(id); return el ? el.textContent : ''; }
  function b64u(s){ var bin = atob(s), u = new Uint8Array(bin.length); for (var i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return new TextDecoder().decode(u); }
  function js(o){ return JSON.stringify(o).replace(/<\//g, '<\\/'); }
  function makeHtml(d, fileName){
    var mk = d.info.today.split('/'); var key = mk[2] + '-' + mk[1];
    var html = b64u(txt('kl-src'))
      .replace('/*KPI_MONTHS*/', function(){ return 'window.KPI_MONTHS=' + js(d.KPI_MONTHS) + ';window.KPI_DATA=window.KPI_MONTHS["' + key + '"];'; })
      .replace('/*KPI_DAILY*/', function(){ return 'window.KPI_DAILY=' + js(d.KPI_DAILY) + ';'; })
      .replace('/*KPI_MONTHLY_DETAIL*/', function(){ return 'window.KPI_MONTHLY_DETAIL=' + js(d.KPI_MONTHLY_DETAIL) + ';'; })
      .replace('/*T11_DATA*/', function(){ return 'window.T11_DATA=' + js(d.T11_DATA) + ';window.__kplUploaded=' + js(fileName) + ';'; })
      .replace('/*BDAY*/', function(){ return js(d.BDAY); })
      .replace('/*TRAN_XANH*/', function(){ return 'window.TRAN_XANH=' + js(d.TRAN_XANH || null) + ';'; });
    if (html.indexOf('daily-report.css') < 0) html = html.replace('</head>', '<link rel="stylesheet" href="daily-report.css?v=196"></head>');
    if (html.indexOf('daily-report.js') < 0) html = html.replace('</body>', '<script src="daily-report.js?v=196"><\/script></body>');
    if (html.indexOf('ddkd-photos-v202.js') < 0) html = html.replace('</body>', '<script src="ddkd-photos-v202.js"><\/script></body>');
    if (html.indexOf('tran-xanh.js') < 0) html = html.replace('</body>', '<script src="tran-xanh.js?v=1"><\/script></body>');
    var keep = ['kl-src', 'kl-tpl', 'kl-hist', 'kl-lib'].map(function(id){ var el = document.getElementById(id); return el ? el.outerHTML : ''; }).join('');
    return html.replace('</body>', function(){ return keep + '</body>'; });
  }
  function mountSnapshot(d, fileName){
    var html = makeHtml(d, fileName); var old = document.getElementById('klFrame'); if (old) old.remove();
    var fr = document.createElement('iframe'); fr.id = 'klFrame'; fr.title = 'Dashboard đã lưu ' + fileName; fr.style.cssText = 'position:fixed;inset:0;width:100%;height:100%;border:0;z-index:2147483000;background:#f5f7fa'; fr.srcdoc = html; document.body.appendChild(fr); closeBox();
  }
  function gzipSnapshot(d){
    if(!window.CompressionStream) return Promise.reject(new Error('Trình duyệt chưa hỗ trợ lưu snapshot. Hãy dùng Chrome hoặc Edge mới.'));
    var blob = new Blob([JSON.stringify(d)], {type:'application/json'});
    return new Response(blob.stream().pipeThrough(new CompressionStream('gzip'))).arrayBuffer();
  }
  // ===== Lưu lên GitHub (thay cho máy chủ cũ) =====
  var GHK = 'kpi-gh-token', GHR = 'kpi-gh-repo';
  function ls(k, v){ try { if (v === undefined) return localStorage.getItem(k) || ''; if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch(e){ return ''; } }
  function guessRepo(){ var h = location.hostname, m = h.match(/^([^.]+)\.github\.io$/i); if (!m) return ''; var p = location.pathname.split('/').filter(Boolean)[0]; return m[1] + '/' + (p && !/\.html?$/i.test(p) ? p : (m[1] + '.github.io')); }
  function abToB64(ab){ var u = new Uint8Array(ab), s = '', ch = 0x8000; for (var i = 0; i < u.length; i += ch) s += String.fromCharCode.apply(null, u.subarray(i, i + ch)); return btoa(s); }
  function gh(path, opt){ var tk = ls(GHK), repo = ls(GHR) || guessRepo(); opt = opt || {}; opt.headers = Object.assign({ 'Accept': 'application/vnd.github+json', 'Authorization': 'Bearer ' + tk, 'X-GitHub-Api-Version': '2022-11-28' }, opt.headers || {}); return fetch('https://api.github.com/repos/' + repo + path, opt); }
  function saveSnapshot(d, fileName){
    return gzipSnapshot(d).then(function(body){
      var p = '/contents/data/snapshot.json.gz';
      return gh(p + '?ref=main', { cache: 'no-store' }).then(function(r){ return r.ok ? r.json() : {}; }).then(function(cur){
        var msg = 'Cập nhật dữ liệu dashboard: ' + fileName + ' · ' + (d.info && d.info.today || '');
        return gh(p, { method: 'PUT', headers: { 'content-type': 'application/json' }, body: JSON.stringify({ message: msg, content: abToB64(body), sha: cur && cur.sha || undefined, branch: 'main' }) });
      });
    }).then(function(r){ return r.json().then(function(j){ if (!r.ok) throw new Error((j && j.message) || ('GitHub báo lỗi ' + r.status)); return j; }); });
  }
  function askToken(){
    return new Promise(function(res){
      box('<b>Lưu lên link chung (GitHub)</b>Chỉ người quản lý dashboard cần làm 1 lần trên máy này. ĐDKD không cần.' +
        '<label style="display:block;margin-top:10px;font-size:13px">Kho GitHub (tài khoản/tên-kho)<input id="klRepo" style="width:100%;padding:8px;margin-top:4px;border:1px solid #c7d4f0;border-radius:8px" value="' + (ls(GHR) || guessRepo()).replace(/"/g, '') + '"></label>' +
        '<label style="display:block;margin-top:8px;font-size:13px">GitHub token (quyền Contents: Read and write)<input id="klTok" type="password" style="width:100%;padding:8px;margin-top:4px;border:1px solid #c7d4f0;border-radius:8px" placeholder="github_pat_…"></label>' +
        '<small>Token chỉ lưu trong trình duyệt của máy này.</small><button class="kl-btn" type="button" id="klTokOk">Lưu lên GitHub</button> <button class="kl-btn alt" type="button" id="klTokNo">Chỉ xem trên máy này</button>');
      document.getElementById('klTokOk').onclick = function(){ var t = document.getElementById('klTok').value.trim(), r = document.getElementById('klRepo').value.trim(); if (t) ls(GHK, t); if (r) ls(GHR, r); res(!!ls(GHK)); };
      document.getElementById('klTokNo').onclick = function(){ res(false); };
    });
  }
  function restoreLatest(){ /* bản GitHub: dữ liệu đã nằm sẵn trong trang, không cần tải lại */ }
  inp.onchange = function(){
    var f = inp.files && inp.files[0]; if (!f) return;
    box('<b>Đang đọc file ' + f.name.replace(/[<>&]/g, '') + '</b>Bước 1/3: tải thư viện đọc Excel…<small>Có thể mất 15–40 giây với file khoảng 40MB. Đừng đóng trang.</small>');
    loadXLSX().then(function(){ return new Promise(function(res, rej){ var rd = new FileReader(); rd.onload = function(){ res(rd.result); }; rd.onerror = function(){ rej(new Error('Không đọc được file.')); }; rd.readAsArrayBuffer(f); }); })
    .then(function(buf){ box('<b>Đang xử lý dữ liệu…</b>Bước 2/3: đọc các sheet KPI THÁNG, MCP, 061 new, Combo, HV, Trưng bày, Bonus, LE…<small>Trang có thể đứng vài giây.</small>'); return new Promise(function(r){ setTimeout(function(){ r(buf); }, 60); }); })
    .then(function(buf){
      var wb = window.KPIConvert.readWorkbook(XLSX, window.fixXlsb, new Uint8Array(buf));
      var d = window.KPIConvert.convert(XLSX, wb, { tpl: JSON.parse(txt('kl-tpl')), hist: { K: JSON.parse(txt('kl-hist')) }, sourceFile: f.name });
      return (ls(GHK) ? Promise.resolve(true) : askToken()).then(function(ok){
        if (!ok) { mountSnapshot(d, f.name); return; }
        box('<b>Đang lưu lên GitHub…</b>Bước 3/3: ' + d.info.reps + ' ĐDKD · dữ liệu ngày ' + d.info.today + '<small>Sau khi lưu, link chung tự cập nhật trong khoảng 2–3 phút.</small>');
        return saveSnapshot(d, f.name).then(function(){ mountSnapshot(d, f.name); setTimeout(function(){ try { (window.top || window).alert('Đã lưu lên GitHub. Link chung sẽ có số mới sau khoảng 2–3 phút.'); } catch(e){} }, 300); })
          .catch(function(e){ if (/credentials|401|403|Not Found|404/i.test(String(e && e.message))) ls(GHK, null); throw new Error('Đã đọc file nhưng chưa lưu được lên GitHub: ' + (e && e.message || e) + '. Kiểm tra lại token/tên kho rồi nạp lại.'); });
      });
    })
    .catch(function(e){ console.error(e); box('<b>Không nạp được file</b><span class="kl-err">' + String(e && e.message || e).replace(/[<>&]/g, '') + '</span><small>Cần đúng mẫu file Tracking KPI (có sheet KPI THÁNG, MCP, 061 new…).</small><button class="kl-btn" type="button" id="klClose">Đóng</button>'); document.getElementById('klClose').onclick = closeBox; })
    .finally(function(){ inp.value = ''; });
  };
  restoreLatest();
})();
