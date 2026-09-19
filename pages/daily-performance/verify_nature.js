// verify_nature.js — 验证"自然到店"数据修复（v20260910-64）
// 修复前：品牌组-自然到店 行与 短视频组 行共用 targetKey '短视频组及自然到店'，
//         但该 key 不在 values 中（只有 '短视频组' 与 '品牌组-自然到店'），导致两行值均为 0/合并。
// 修复后：各用自身 valueKey，自然到店 = 仅营销部/品牌组/二级含"自然到店"的收款金额-现款支付，按日期独立汇总。
// 场景：
//  A. 短视频组 行月累计 = 短视频组记录现金之和（不含自然到店）
//  B. 品牌组-自然到店 行月累计 = 仅"自然到店"记录现金之和（不含短视频组）
//  C. 两者互不包含（自然到店 ≠ 短视频组+自然到店 合并值）
//  D. 品牌组-自然到店 某日每日列 = 该日自然到店现金之和（按日期汇总）
//  E. 月度目标格仍显示合并目标 3（合并 targetKey 不变）
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
  dom.window.showToast = () => {};
  dom.window.downloadBlob = () => {};
  return dom;
}

function loadApp(dom){
  const { window } = dom;
  Object.defineProperty(window.document, 'readyState', { value: 'loading', configurable: true });
  const src = '(function(localStorage){\n' + appSrc +
    '\n;window.__app = { state, recompute, renderReportInner };\n})(window.localStorage);';
  try { window.eval(src); } catch (e) { console.error('APP_EVAL_ERROR:', e && e.message); throw e; }
  return window.__app;
}

const dom = buildDom();
const app = loadApp(dom);
const win = dom.window, doc = win.document;
const num = t => Number(String(t).replace(/[^\d.-]/g, ''));

// 构造订单（均在 cutoff 2026-09 当月，确保被统计）
// 短视频组 r4：营销部/短视频组，金额 3000
// 自然到店 r5：营销部/品牌组/二级含"自然到店"，金额 5000（9/10）、2000（9/9）
const mk = (date, amt, l1, l2) => ({
  payTime: new Date(date + 'T10:00:00'), cashPay: amt, _uid: 'u_' + date + '_' + amt + '_' + l1,
  name: '客户', memberNo: '', channelClass: '营销部', channelL1: l1, channelL2: l2 || '',
  channelL3: '', consultant: '', cashier: '南媛'
});
app.state.orders = [
  mk('2026-09-10', 3000, '短视频组', '短视频投放'),   // r4
  mk('2026-09-10', 5000, '品牌组', '自然到店-门口'),   // r5 (9/10)
  mk('2026-09-09', 2000, '品牌组', '自然到店-转介绍'), // r5 (9/9)
];
app.state.cutoffDate = '2026-09-10';
app.recompute();
app.renderReportInner();

const monthCell = key => doc.querySelector(`#report-table td.num[data-rowkey="${key}"]:not([data-idx])`);
const dayCell = (key, date) => doc.querySelector(`#report-table td.num[data-rowkey="${key}"][data-idx][data-date="${date}"]`);

const checks = [];
const check = (n, c, d) => checks.push({ n, c: !!c, d });

// A. 短视频组 月累计 = 3000（不含自然到店）
check('A: 短视频组 月累计=3,000（仅短视频组记录）', num(monthCell('短视频组').textContent) === 3000, 'text=' + monthCell('短视频组').textContent);
// B. 品牌组-自然到店 月累计 = 5000+2000=7,000（仅自然到店）
check('B: 品牌组-自然到店 月累计=7,000（仅自然到店记录）', num(monthCell('品牌组-自然到店').textContent) === 7000, 'text=' + monthCell('品牌组-自然到店').textContent);
// C. 两者互不包含：自然到店 ≠ 短视频组+自然到店
check('C: 自然到店(7,000) ≠ 短视频组+自然到店合并值(10,000)', num(monthCell('品牌组-自然到店').textContent) !== 10000, 'nature=' + num(monthCell('品牌组-自然到店').textContent));
check('C2: 短视频组(3,000) ≠ 合并值(10,000)', num(monthCell('短视频组').textContent) !== 10000, 'short=' + num(monthCell('短视频组').textContent));
// D. 自然到店 按日期汇总：9/10=5000, 9/9=2000
check('D: 自然到店 每日 9/10=5,000', num(dayCell('品牌组-自然到店', '2026-09-10').textContent) === 5000, 'text=' + dayCell('品牌组-自然到店', '2026-09-10').textContent);
check('D: 自然到店 每日 9/9=2,000', num(dayCell('品牌组-自然到店', '2026-09-09').textContent) === 2000, 'text=' + dayCell('品牌组-自然到店', '2026-09-09').textContent);
// 短视频组 每日 9/10=3000（验证互不串）
check('D2: 短视频组 每日 9/10=3,000（不与自然到店混淆）', num(dayCell('短视频组', '2026-09-10').textContent) === 3000, 'text=' + dayCell('短视频组', '2026-09-10').textContent);
// E. 月目标合并格仍显示 3
const targetCell = doc.querySelector('#report-table td.target[data-targetkey="短视频组及自然到店"]');
check('E: 月目标合并格显示 3（合并 targetKey 不变）', targetCell && num(targetCell.textContent) === 3, 'text=' + (targetCell ? targetCell.textContent : 'null'));

const failed = checks.filter(c => !c.c);
checks.forEach(c => console.log((c.c ? 'PASS' : 'FAIL') + ' | ' + c.n + ' | ' + c.d));
console.log('\n' + (failed.length ? 'FAILED ' + failed.length : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
