// verify_persist.js — real jsdom test of historical-order persistence across a page refresh.
// Scenario: import orders -> persistOrders() saves to localStorage -> simulate reload (fresh jsdom with
// same localStorage) -> restoreOrders() + recompute() -> monthly-cumulative & daily values restored.
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const proj = 'D:/workbuddy/2026-08-15-10-10-03/daily-performance';
const appSrc = fs.readFileSync(path.join(proj, 'assets', 'app.js'), 'utf8');

// Shared localStorage between "sessions"
const store = {};
const makeLocalStorage = () => ({
  getItem: k => (k in store ? store[k] : null),
  setItem: (k, v) => { store[k] = String(v); },
  removeItem: k => { delete store[k]; }
});

function buildDom(){
  const html = `<!DOCTYPE html><html><body>
    <input id="cutoff-date" type="date">
    <div id="debug-banner"></div>
    <div id="vab-meta"></div>
    <div id="data-status"></div>
    <div id="orders-meta"></div>
    <div id="toast"></div>
    <button id="file-orders"></button>
    <button id="file-vab"></button>
    <button id="edit-vab"></button>
    <button id="vab-save"></button>
    <button id="vab-cancel"></button>
    <button id="recompute"></button>
    <button id="clear-manual"></button>
    <button id="clear-vab"></button>
    <button id="targets-toggle"></button>
    <button id="export-png"></button>
    <button id="export-xlsx"></button>
    <button id="export-store-png"></button>
    <button id="export-store-xlsx"></button>
    <div id="targets-body"></div>
    <div id="targets-arrow"></div>
    <div id="render-error"></div>
    <table class="report" id="report-table"><thead></thead><tbody id="report-body"></tbody></table>
    <table class="report" id="store-table"><thead></thead><tbody id="store-body"></tbody></table>
  </body></html>`;
  const dom = new JSDOM(html, { runScripts: 'outside-only', pretendToBeVisual: true });
  const { window } = dom;
  // jsdom's localStorage is read-only; override with a stub backed by the SHARED `store`
  // so "session 1" writes and "session 2/3" (fresh pages) read the same persisted data.
  Object.defineProperty(window, 'localStorage', { value: makeLocalStorage(), configurable: true });
  window.showToast = () => {};
  window.downloadBlob = () => {};
  return dom;
}

// Run app.js inside the jsdom window scope; capture live symbols via window.__app.
// Force readyState='loading' so initApp is deferred (we drive restore/recompute manually).
function loadApp(dom){
  const { window } = dom;
  Object.defineProperty(window.document, 'readyState', { value: 'loading', configurable: true });
  // 用 IIFE 把 stub 作为 `localStorage` 形参注入：window.eval 的全局作用域下，Object.defineProperty
  // 的实例自有属性无法挡住 jsdom 原型的 localStorage getter，故显式以参数遮蔽，确保 app.js 内所有
  // localStorage 引用都解析到共享 store 的 stub（刷新周期可跨“会话”读取同一份持久化数据）。
  const stub = window.localStorage;
  const src = '(function(localStorage){\n' + appSrc +
    '\n;window.__app = { state, persistOrders, restoreOrders, persistColWidths, restoreColWidths, recompute, calculate, calculateDay };\n})(window.localStorage);';
  try {
    window.eval(src);
  } catch (e) {
    console.error('APP_EVAL_ERROR:', e && e.message, '\n', (e && e.stack || '').split('\n').slice(0,4).join('\n'));
    throw e;
  }
  if (!window.__app) throw new Error('window.__app not captured (app.js eval produced no __app)');
  return window.__app;
}

const checks = [];
const check = (n, c, d) => checks.push({ n, c: !!c, d });

// ---- Session 1: import orders, then persist (top-level restore already ran but orders empty) ----
const dom1 = buildDom();
const app1 = loadApp(dom1);
const mk = (date, amt, consultant) => ({ payTime: new Date(date + 'T10:00:00'), cashPay: amt, consultant: consultant || '', channelClass: '商务部', channelL1: '商务部', channelL2: '商务部' });
app1.state.orders = [
  mk('2026-09-09', 1000, '王爽'),
  mk('2026-09-09', 2000, '李娜'),
  mk('2026-09-10', 3000, '王爽')
];
app1.state.cutoffDate = '2026-09-10';
app1.persistOrders(); // simulate onFileOrders() persistence
const persisted1 = dom1.window.localStorage.getItem('dailyPerformance_orders');

check('session1: orders persisted to localStorage', !!persisted1, persisted1 ? 'len=' + JSON.parse(persisted1).orders.length : 'missing');
check('session1: persisted cutoff = 2026-09-10', JSON.parse(persisted1).cutoff === '2026-09-10', JSON.parse(persisted1).cutoff);

// ---- Session 2: fresh page load (same localStorage) ----
const dom2 = buildDom();
dom2.window.localStorage.setItem('dailyPerformance_orders', persisted1); // carry persisted data across "reload"
const app2 = loadApp(dom2); // top-level restoreOrders() already ran during script eval
check('session2: orders restored from localStorage', app2.state.orders.length === 3, 'restored=' + app2.state.orders.length);
check('session2: cutoff restored', app2.state.cutoffDate === '2026-09-10', app2.state.cutoffDate);

// mimic initApp restore path: set cutoff input + recompute
dom2.window.document.getElementById('cutoff-date').value = app2.state.cutoffDate;
app2.recompute();
const monthTotal = app2.state.results ? app2.state.results.values['合计'] : undefined;
check('session2: monthly-cumulative recomputed (=6000)', monthTotal === 6000, 'month=合计 ' + monthTotal);

// daily values derived per date via calculateDay (the same function recompute uses)
const day9_10 = app2.calculateDay(app2.state.orders, app2.state.vabSet, '2026-09-10');
const day9_9 = app2.calculateDay(app2.state.orders, app2.state.vabSet, '2026-09-09');
check('session2: 9/10 daily 合计 = 3000', day9_10 && day9_10.values['合计'] === 3000, day9_10 ? day9_10.values['合计'] : 'n/a');
check('session2: 9/9 daily 合计 = 3000', day9_9 && day9_9.values['合计'] === 3000, day9_9 ? day9_9.values['合计'] : 'n/a');

// ---- Session 3: delete one order, persist, reload -> reflects deletion ----
const ymd = d => d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
app2.state.orders = app2.state.orders.filter(r => ymd(r.payTime) !== '2026-09-10'); // drop 9/10 row
app2.persistOrders(); // simulate deleteOrder() persistence
const persisted2 = dom2.window.localStorage.getItem('dailyPerformance_orders');
const dom3 = buildDom();
dom3.window.localStorage.setItem('dailyPerformance_orders', persisted2); // carry across reload
const app3 = loadApp(dom3);
check('session3: after delete+reload, 9/10 row gone (2 orders left)', app3.state.orders.length === 2, 'restored=' + app3.state.orders.length);
dom3.window.document.getElementById('cutoff-date').value = app3.state.cutoffDate || '2026-09-10';
app3.recompute();
const monthTotal3 = app3.state.results ? app3.state.results.values['合计'] : undefined;
check('session3: monthly-cumulative now 3000 (1000+2000)', monthTotal3 === 3000, 'month=合计 ' + monthTotal3);

// ---- Session 4: column-width persistence across refresh (regression for COLW_STORAGE_KEY TDZ) ----
// 若 restoreColWidths() 在 const COLW_STORAGE_KEY 初始化前被调用，会抛 ReferenceError 被 catch 吞掉，
// 导致 state.colWidths 不恢复（仍为 {}）。本会话验证修复后列宽可跨刷新保留。
const dom4 = buildDom();
const app4 = loadApp(dom4);
app4.state.colWidths = { 0: 260, 2: 140, 5: 90 };
app4.persistColWidths();
const persistedCW = dom4.window.localStorage.getItem('dailyPerformance_colw');
check('session4: colw persisted to localStorage', !!persistedCW, persistedCW ? 'len=' + persistedCW.length : 'missing');
const dom5 = buildDom();
dom5.window.localStorage.setItem('dailyPerformance_colw', persistedCW); // carry across reload
const app5 = loadApp(dom5); // top-level restoreColWidths() runs during eval
check('session4: colw restored from localStorage', JSON.stringify(app5.state.colWidths) === JSON.stringify({ 0: 260, 2: 140, 5: 90 }), JSON.stringify(app5.state.colWidths));

const failed = checks.filter(c => !c.c);
checks.forEach(c => console.log((c.c ? 'PASS' : 'FAIL') + ' | ' + c.n + ' | ' + c.d));
console.log('\n' + (failed.length ? 'FAILED ' + failed.length : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
