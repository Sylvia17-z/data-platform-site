// verify_batch_expand.js — 验证「批量 / 展开」从报表表头移出、改为面板独立按钮
// ① 静态检查（源码层面）
// ② 功能检查（jsdom 加载真实 app.js，模拟点击独立按钮，断言列隐藏/展开行为不变）
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const proj = 'D:/workbuddy/2026-08-15-10-10-03/daily-performance';
const appSrc = fs.readFileSync(path.join(proj, 'assets', 'app.js'), 'utf8');
const htmlSrc = fs.readFileSync(path.join(proj, 'index.html'), 'utf8');

let pass = 0, fail = 0;
const check = (n, c, d='') => { if(c){pass++; console.log('  PASS '+n);} else {fail++; console.log('  FAIL '+n+(d?'  ['+d+']':''));} };

/* ---------- ① 静态 ---------- */
check('报表表头不再含 <th id="batch-toggle">', !/<th[^>]*id="batch-toggle"/.test(appSrc));
check('报表表头不再含 <th id="expand-all">', !/<th[^>]*id="expand-all"/.test(appSrc));
check('index.html 含独立按钮 <button id="batch-toggle">', /<button[^>]*id="batch-toggle"/.test(htmlSrc));
check('index.html 含独立按钮 <button id="expand-all">', /<button[^>]*id="expand-all"/.test(htmlSrc));
check('新增 bindBatchExpandButtons 函数', /function bindBatchExpandButtons\(tableEl\)/.test(appSrc));
check('initApp 中绑定独立按钮', /bindBatchExpandButtons\(\s*\$\('report-table'\)\s*\)/.test(appSrc));
check('bindDailyColumnToggle 不再绑定 batch/expand(已移出)', !/const batchBtn = tableEl\.querySelector\('#batch-toggle'\)/.test(appSrc));
check('updateMainTitleColspan 改为 5 + visibleDaily', /titleTh\.colSpan = 5 \+ visibleDaily/.test(appSrc));
check('totalCols 不再 +2', /const totalCols = 3 \+ 1 \+ 1 \+ dateList\.length;/.test(appSrc));

/* ---------- ② 功能（jsdom） ---------- */
function buildDom(){
  const html = `<!DOCTYPE html><html><body>
    <input id="cutoff-date" type="date">
    <div id="debug-banner"></div>
    <div id="vab-meta"></div><div id="data-status"></div>
    <div id="orders-meta"></div><div id="toast"></div><div id="render-error"></div>
    <button id="batch-toggle">☑ 批量</button>
    <button id="expand-all">+展开</button>
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
    '\n;window.__app = { state, recompute, renderReport, bindBatchExpandButtons };\n})(window.localStorage);';
  window.eval(src);
  return window.__app;
}

const dom = buildDom();
const app = loadApp(dom);
const doc = dom.window.document;
const mk = (date, amt) => ({ payTime: new Date(date + 'T10:00:00'), cashPay: amt, channelClass: '商务部', channelL1: '商务部', channelL2: '商务部' });
app.state.orders = [ mk('2026-09-09', 1000), mk('2026-09-10', 3000) ];
app.state.cutoffDate = '2026-09-10';
doc.getElementById('cutoff-date').value = '2026-09-10';
app.recompute();

const reportTable = doc.getElementById('report-table');
const dailyThs = () => [...reportTable.querySelectorAll('th.daily-col')];
const hiddenCount = () => reportTable.querySelectorAll('th.daily-col.col-hidden').length;
const batchBtn = doc.getElementById('batch-toggle');
const expandBtn = doc.getElementById('expand-all');

check('渲染后存在每日列(>=2)', dailyThs().length >= 2, 'n=' + dailyThs().length);

// 绑定独立按钮（仅一次）
app.bindBatchExpandButtons(reportTable);
app.bindBatchExpandButtons(reportTable); // 二次调用不应重复绑定

// 展开：默认旧日期列已隐藏，点独立「+展开」→ 全部可见
check('展开前存在隐藏列(默认旧日期隐藏)', hiddenCount() > 0, 'hidden=' + hiddenCount());
expandBtn.click();
check('点独立「+展开」后无隐藏列', hiddenCount() === 0, 'hidden=' + hiddenCount());

// 批量：进入批量模式 → 选中一列 → 完成隐藏 → 该列隐藏且退出批量
const target = dailyThs()[0];
batchBtn.click();
check('点批量后按钮变「✕ 完成隐藏」', batchBtn.textContent === '✕ 完成隐藏', 'txt=' + batchBtn.textContent);
check('点批量后表格进入 batch-mode', reportTable.classList.contains('batch-mode'));
target.click(); // 批量模式下点击日列 = 选中
check('日列被选中(selected)', target.classList.contains('selected'));
const targetDate = target.getAttribute('data-date');
batchBtn.click(); // 完成隐藏
check('完成后按钮恢复「☑ 批量」', batchBtn.textContent === '☑ 批量', 'txt=' + batchBtn.textContent);
check('完成后退出 batch-mode', !reportTable.classList.contains('batch-mode'));
const stillHidden = [...reportTable.querySelectorAll('th.daily-col.col-hidden')].map(t=>t.getAttribute('data-date'));
check('选中的日期列被隐藏', stillHidden.includes(targetDate), 'hidden=' + stillHidden.join(','));

// 重复绑定防护：二次绑定后再次进入/完成批量，仅触发一次（行为一致，不抛错）
let threw = false;
try { app.bindBatchExpandButtons(reportTable); batchBtn.click(); batchBtn.click(); } catch(e){ threw = true; }
check('重复绑定不抛错且仍可操作', !threw);

console.log((fail===0?'\nALL_PASS':'\nFAILED')+' ('+(pass+fail)+' checks)');
process.exit(fail===0?0:1);
