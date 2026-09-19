// verify_detail_filter.js — 验证"数据明细支持按日期 / 月度筛选"（v20260910-63）
// 场景：
//  A. 无筛选：全部标签展示所有导入订单
//  B. 月度筛选：选 2026-08 → 仅该月；选 2026-09 → 仅该月
//  C. 日期筛选：选 2026-09-10 → 仅该日
//  D. 月度 + 日期交集：2026-09 + 2026-09-10 → 交集；2026-09 + 2026-08-15 → 空集（不崩溃）
//  E. 选项生成：月度/日期下拉选项来自订单实际出现的月份/日期（含"全部"）
//  F. 清除筛选按钮：置灰筛选后出现，点击后恢复全部
//  G. 与标签交集：today 标签 + 月度筛选 取正确交集
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
    <div id="import-details"></div>
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
    '\n;window.__app = { state, recompute, renderImportDetails, filterOrdersByTab, onImportDetailsChange, fmtMonth, fmtDate };\n})(window.localStorage);';
  try { window.eval(src); } catch (e) { console.error('APP_EVAL_ERROR:', e && e.message); throw e; }
  return window.__app;
}

const checks = [];
const check = (n, c, d) => checks.push({ n, c: !!c, d });

const dom = buildDom();
const app = loadApp(dom);
const win = dom.window;
const doc = win.document;

// 构造跨月/跨日的订单：2026-08 两条、2026-09 三条（其中 9/10 两条）
const mk = (date, amt, l1) => ({
  payTime: new Date(date + 'T10:00:00'), cashPay: amt, _uid: 'u_' + date + '_' + amt,
  name: '客户' + amt, memberNo: '', channelClass: '商务部', channelL1: l1 || '商务部',
  channelL2: l1 || '商务部', channelL3: '', consultant: '', cashier: '南媛'
});
app.state.orders = [
  mk('2026-08-15', 1000),
  mk('2026-08-20', 2000),
  mk('2026-09-01', 3000),
  mk('2026-09-10', 4000),
  mk('2026-09-10', 5000, '市场部')
];
app.state.cutoffDate = '2026-09-10';
app.recompute();
app.renderImportDetails();

const rows = () => doc.querySelectorAll('#import-details tbody tr').length;
const monthSel = () => doc.getElementById('detail-month-filter');
const dateSel  = () => doc.getElementById('detail-date-filter');
const setMonth = m => { const s = monthSel(); s.value = m; s.dispatchEvent(new win.Event('change', { bubbles: true })); };
const setDate  = d => { const s = dateSel(); s.value = d; s.dispatchEvent(new win.Event('change', { bubbles: true })); };
const clickTab = key => { const b = doc.querySelector(`.detail-tab[data-tab="${key}"]`); b.dispatchEvent(new win.MouseEvent('click', { bubbles: true })); };
const clickReset = () => { const b = doc.getElementById('detail-filter-reset'); if(b) b.dispatchEvent(new win.MouseEvent('click', { bubbles: true })); };

// A. 无筛选：5 行
check('A: 初始全部标签=5 行', rows() === 5, 'rows=' + rows());

// E. 选项生成
const mOpts = [...monthSel().options].map(o => o.value);
const dOpts = [...dateSel().options].map(o => o.value);
check('E: 月度选项含 [全部, 2026-08, 2026-09]', JSON.stringify(mOpts) === JSON.stringify(['', '2026-08', '2026-09']), 'opts=' + JSON.stringify(mOpts));
check('E: 日期选项含 4 个实际日期', dOpts.length === 5 && dOpts[0] === '', 'opts=' + JSON.stringify(dOpts));

// B. 月度筛选
setMonth('2026-08');
check('B: 月度=2026-08 → 2 行', rows() === 2, 'rows=' + rows());
check('B: filterOrdersByTab 返回 2', app.filterOrdersByTab('all').length === 2, 'len=' + app.filterOrdersByTab('all').length);
setMonth('2026-09');
check('B: 月度=2026-09 → 3 行', rows() === 3, 'rows=' + rows());
setMonth(''); // 复位

// C. 日期筛选
setDate('2026-09-10');
check('C: 日期=2026-09-10 → 2 行', rows() === 2, 'rows=' + rows());
check('C: 该 2 行均为 9/10', app.filterOrdersByTab('all').every(r => app.fmtDate(r.payTime) === '2026-09-10'), 'all9/10');
setDate('2026-09-01');
check('C: 日期=2026-09-01 → 1 行', rows() === 1, 'rows=' + rows());
setDate('');

// D. 月度 + 日期交集
setMonth('2026-09'); setDate('2026-09-10');
check('D: 2026-09 ∩ 2026-09-10 → 2 行', rows() === 2, 'rows=' + rows());
setDate('2026-08-15'); // 与 2026-09 月度无交集
check('D: 2026-09 ∩ 2026-08-15 → 0 行（不崩溃）', rows() === 0, 'rows=' + rows());
setMonth(''); setDate('');

// F. 清除筛选按钮
setMonth('2026-08');
check('F: 激活筛选后出现清除按钮', !!doc.getElementById('detail-filter-reset'), 'hasReset=' + !!doc.getElementById('detail-filter-reset'));
clickReset();
check('F: 点击清除筛选 → 恢复 5 行', rows() === 5, 'rows=' + rows());
check('F: 清除后按钮消失', !doc.getElementById('detail-filter-reset'), 'hasReset=' + !!doc.getElementById('detail-filter-reset'));

// G. 与标签交集：today 标签（最新日 = 2026-09-10）+ 月度筛选
clickTab('today');
check('G: today 标签（最新日）=2 行', rows() === 2, 'rows=' + rows());
setMonth('2026-09');
check('G: today + 月度2026-09 → 2 行', rows() === 2, 'rows=' + rows());
setMonth('2026-08');
check('G: today + 月度2026-08 → 0 行（交集为空）', rows() === 0, 'rows=' + rows());
setMonth('');
check('G: 取消月度筛选 → today 回到 2 行', rows() === 2, 'rows=' + rows());

const failed = checks.filter(c => !c.c);
checks.forEach(c => console.log((c.c ? 'PASS' : 'FAIL') + ' | ' + c.n + ' | ' + c.d));
console.log('\n' + (failed.length ? 'FAILED ' + failed.length : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
