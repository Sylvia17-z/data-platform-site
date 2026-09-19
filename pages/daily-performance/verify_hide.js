// verify_hide.js — 验证"手动调整后展开的数据列不再被自动隐藏"。
// 场景：默认隐藏除最新列外的所有日期列；用户点"+展开"全部展开(或部分取消隐藏)后，模拟一次手动调整触发的重渲染，
// 断言展开的列仍可见（不再被 defaultHideOldDates 重隐藏）；反向：默认隐藏的列在重渲染后仍保持隐藏。
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const proj = 'D:/workbuddy/2026-08-15-10-10-03/daily-performance';
const appSrc = fs.readFileSync(path.join(proj, 'assets', 'app.js'), 'utf8');

function buildDom(){
  const html = `<!DOCTYPE html><html><body>
    <input id="cutoff-date" type="date">
    <div id="debug-banner"></div>
    <div id="vab-meta"></div>
    <div id="data-status"></div>
    <div id="orders-meta"></div>
    <div id="toast"></div>
    <div id="render-error"></div>
    <table class="report" id="report-table"><thead></thead><tbody id="report-body"></tbody></table>
    <table class="report" id="store-table"><thead></thead><tbody id="store-body"></tbody></table>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'outside-only', pretendToBeVisual: true, url: 'http://localhost/' });
  const { window } = dom;
  window.showToast = () => {};
  window.downloadBlob = () => {};
  return dom;
}

function loadApp(dom){
  const { window } = dom;
  Object.defineProperty(window.document, 'readyState', { value: 'loading', configurable: true });
  const src = '(function(localStorage){\n' + appSrc +
    '\n;window.__app = { state, recompute, renderReport, calculate, calculateDay };\n})(window.localStorage);';
  try { window.eval(src); } catch (e) { console.error('APP_EVAL_ERROR:', e && e.message); throw e; }
  return window.__app;
}

const checks = [];
const check = (n, c, d) => checks.push({ n, c: !!c, d });

const dom = buildDom();
const app = loadApp(dom);
const doc = dom.window.document;
const mk = (date, amt) => ({ payTime: new Date(date + 'T10:00:00'), cashPay: amt, channelClass: '商务部', channelL1: '商务部', channelL2: '商务部' });

// 准备数据：9/9 与 9/10 各一单，cutoff=9/10（生成 1~10 日共 10 个每日列）
app.state.orders = [ mk('2026-09-09', 1000), mk('2026-09-10', 3000) ];
app.state.cutoffDate = '2026-09-10';
doc.getElementById('cutoff-date').value = '2026-09-10';

const hiddenDates = () => [...doc.querySelectorAll('#report-table th.daily-col.col-hidden')].map(t => t.getAttribute('data-date'));
const visibleDates = () => [...doc.querySelectorAll('#report-table th.daily-col:not(.col-hidden)')].map(t => t.getAttribute('data-date'));
const hideTh = (pred) => doc.querySelectorAll('#report-table th.daily-col').forEach(th => { if(pred(th.getAttribute('data-date'))) th.classList.add('col-hidden'); });
const unhideAll = () => doc.querySelectorAll('#report-table .col-hidden, #store-table .col-hidden').forEach(el => el.classList.remove('col-hidden'));

// 第一次渲染（默认：仅最新列 9/10 可见，9/1~9/9 共 9 列隐藏）
app.recompute();
check('默认渲染：9 个旧日期列隐藏', hiddenDates().length === 9, 'hidden=' + hiddenDates().length);
check('默认渲染：仅 9/10 可见', visibleDates().length === 1 && visibleDates()[0] === '2026-09-10', JSON.stringify(visibleDates()));
check('默认渲染(门店表)：9 列隐藏', doc.querySelectorAll('#store-table th.daily-col.col-hidden').length === 9, 'hidden=' + doc.querySelectorAll('#store-table th.daily-col.col-hidden').length);

// 反向①：默认渲染后立刻"手动调整"重渲染，原本隐藏的 9 列应保持隐藏（含 9/9）
app.recompute();
let hd = hiddenDates();
check('默认隐藏列在重渲染后保持隐藏(含9/9)', hd.length === 9 && hd.includes('2026-09-09'), 'hidden=' + hd.length + ' has9/9=' + hd.includes('2026-09-09'));

// 场景A：用户点"+展开"全部展开 → 手动调整重渲染 → 展开的列应仍全部可见（核心修复）
unhideAll();
check('展开后：主表无隐藏列', doc.querySelectorAll('#report-table th.daily-col.col-hidden').length === 0, 'hidden=' + doc.querySelectorAll('#report-table th.daily-col.col-hidden').length);
check('展开后：门店表无隐藏列', doc.querySelectorAll('#store-table th.daily-col.col-hidden').length === 0, 'hidden=' + doc.querySelectorAll('#store-table th.daily-col.col-hidden').length);
app.recompute();
check('手动调整后：主表展开的列仍全部可见(不自动隐藏)', doc.querySelectorAll('#report-table th.daily-col.col-hidden').length === 0, 'hidden=' + doc.querySelectorAll('#report-table th.daily-col.col-hidden').length);
check('手动调整后：门店表展开的列仍全部可见(不自动隐藏)', doc.querySelectorAll('#store-table th.daily-col.col-hidden').length === 0, 'hidden=' + doc.querySelectorAll('#store-table th.daily-col.col-hidden').length);

// 场景B：从全可见状态，仅隐藏 9/1~9/8（保留 9/9、9/10 可见）→ 手动调整重渲染 → 仅 9/9 与 9/10 可见
hideTh(d => d !== '2026-09-09' && d !== '2026-09-10');
app.recompute();
check('部分隐藏(留9/9,9/10)后手动调整：仅 9/9 与 9/10 可见',
  JSON.stringify(visibleDates().slice().sort()) === JSON.stringify(['2026-09-09', '2026-09-10']),
  JSON.stringify(visibleDates()));

const failed = checks.filter(c => !c.c);
checks.forEach(c => console.log((c.c ? 'PASS' : 'FAIL') + ' | ' + c.n + ' | ' + c.d));
console.log('\n' + (failed.length ? 'FAILED ' + failed.length : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
