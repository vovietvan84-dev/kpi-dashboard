#!/usr/bin/env node
// Dựng dashboard tĩnh cho GitHub Pages.
//   node build/build.js --xlsb data/T11.xlsb   → đọc Excel, lưu data/snapshot.json.gz, dựng _site/
//   node build/build.js                        → dựng _site/ từ data/snapshot.json.gz có sẵn
const fs = require('fs'), path = require('path'), zlib = require('zlib');
const ROOT = path.resolve(__dirname, '..'), SRC = path.join(ROOT, 'src');
const arg = n => { const i = process.argv.indexOf('--' + n); return i > 0 ? process.argv[i + 1] : null; };
const OUT = path.resolve(ROOT, arg('out') || '_site');
const SNAP = path.join(ROOT, 'data', 'snapshot.json.gz');
const read = f => fs.readFileSync(path.join(SRC, f), 'utf8');
const js = o => JSON.stringify(o).replace(/<\//g, '<\\/');

let d, fileName = 'snapshot';
const xlsb = arg('xlsb');
if (xlsb) {
  const t0 = Date.now();
  const XLSX = require(path.join(SRC, 'xlsx.full.min.js'));
  const fixXlsb = require(path.join(SRC, 'lib', 'fixxlsb.js'));
  const C = require(path.join(SRC, 'lib', 'conv.js'));
  fileName = path.basename(xlsb);
  console.log('Đọc file', fileName);
  const wb = C.readWorkbook(XLSX, fixXlsb, fs.readFileSync(xlsb));
  d = C.convert(XLSX, wb, { tpl: JSON.parse(read('kl-tpl.json')), hist: { K: JSON.parse(read('kl-hist.json')) }, sourceFile: fileName });
  d.source = { file: fileName, builtAt: new Date().toISOString() };
  fs.mkdirSync(path.dirname(SNAP), { recursive: true });
  fs.writeFileSync(SNAP, zlib.gzipSync(Buffer.from(JSON.stringify(d)), { level: 9 }));
  console.log('Đã chuyển đổi trong', ((Date.now() - t0) / 1000).toFixed(1) + 's →', d.info, '\nCảnh báo:', d.warn);
} else {
  d = JSON.parse(zlib.gunzipSync(fs.readFileSync(SNAP)).toString('utf8'));
  console.log('Dùng snapshot có sẵn:', d.info);
}
if (!d || !d.info || !d.info.today) throw new Error('Snapshot không hợp lệ');

// ---- makeHtml (giống nút Nạp file trong trang) ----
const tpl = read('template.html');
const mk = d.info.today.split('/'); const key = mk[2] + '-' + mk[1];
let html = tpl
  .replace('/*KPI_MONTHS*/', () => 'window.KPI_MONTHS=' + js(d.KPI_MONTHS) + ';window.KPI_DATA=window.KPI_MONTHS["' + key + '"];')
  .replace('/*KPI_DAILY*/', () => 'window.KPI_DAILY=' + js(d.KPI_DAILY) + ';')
  .replace('/*KPI_MONTHLY_DETAIL*/', () => 'window.KPI_MONTHLY_DETAIL=' + js(d.KPI_MONTHLY_DETAIL) + ';')
  .replace('/*T11_DATA*/', () => 'window.T11_DATA=' + js(d.T11_DATA) + ';window.__kplUploaded=null;')
  .replace('/*BDAY*/', () => js(d.BDAY))
  .replace('/*TRAN_XANH*/', () => 'window.TRAN_XANH=' + js(d.TRAN_XANH || null) + ';');
const v = Date.now().toString(36);
html = html.replace('</head>', '<link rel="stylesheet" href="daily-report.css?v=' + v + '"></head>');
html = html.replace('</body>', () => '<script src="daily-report.js?v=' + v + '"></script><script src="ddkd-photos-v202.js?v=' + v + '"></script><script src="tran-xanh.js?v=' + v + '"></script></body>');
const lib = read('lib/fixxlsb.js') + '\n' + read('lib/conv.js') + '\n' + read('loader.js');
if (/<\/script/i.test(lib)) throw new Error('kl-lib chứa </script>');
const keep = '<script type="text/plain" id="kl-src">' + Buffer.from(tpl, 'utf8').toString('base64') + '</script>' +
  '<script type="application/json" id="kl-tpl">' + read('kl-tpl.json') + '</script>' +
  '<script type="application/json" id="kl-hist">' + read('kl-hist.json') + '</script>' +
  '<script id="kl-lib">' + lib + '</script>';
html = html.replace('</body>', () => keep + '</body>');

fs.rmSync(OUT, { recursive: true, force: true });
fs.mkdirSync(OUT, { recursive: true });
fs.writeFileSync(path.join(OUT, 'index.html'), html);
['daily-report.css', 'daily-report.js', 'ddkd-photos-v202.js', 'tran-xanh.js', 'xlsx.full.min.js'].forEach(f => fs.copyFileSync(path.join(SRC, f), path.join(OUT, f)));
fs.writeFileSync(path.join(OUT, '.nojekyll'), '');
fs.writeFileSync(path.join(OUT, 'version.json'), JSON.stringify({ updated: d.info.today, file: (d.source && d.source.file) || fileName, builtAt: new Date().toISOString() }));
console.log('Xong:', path.join(OUT, 'index.html'), (html.length / 1e6).toFixed(1) + 'MB');
