// verify_edit.js — 验证"手动调整数据后 月累计/其他数据页联动更新"修复 + "导出图片日期列收窄"。
// 场景：
//  A. 双击编辑主表「每日」单元格(商务部 9/10) → 该行「月累计」应联动更新
//  B. 双击编辑主表「月累计」单元格(商务部) → 该行月累计更新，且「合计」应联动更新
//  C. 门店表在编辑后正常重渲染
//  D. 导出克隆(PNG)：仅含最新日期列、日期列 90px（与页面每日列一致，不吸收剩余宽度）、无导出专用样式覆盖
//  D2. Excel 导出：最新日期列之前的每日列自动隐藏
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
    '\n;window.__app = { state, recompute, renderReport, buildExportClone, buildReportGrid, buildStoreGrid };\n})(window.localStorage);';
  try { window.eval(src); } catch (e) { console.error('APP_EVAL_ERROR:', e && e.message); throw e; }
  return window.__app;
}

const checks = [];
const check = (n, c, d) => checks.push({ n, c: !!c, d });
const num = t => Number(String(t).replace(/[^\d.-]/g, ''));

const dom = buildDom();
const app = loadApp(dom);
const win = dom.window;
const doc = win.document;
const mk = (date, amt) => ({ payTime: new Date(date + 'T10:00:00'), cashPay: amt, channelClass: '商务部', channelL1: '商务部', channelL2: '商务部' });

app.state.orders = [ mk('2026-09-09', 1000), mk('2026-09-10', 3000) ];
app.state.cutoffDate = '2026-09-10';
doc.getElementById('cutoff-date').value = '2026-09-10';
app.recompute();

function editCell(td, val){
  if(!td) throw new Error('editCell: td not found');
  td.dispatchEvent(new win.MouseEvent('dblclick', { bubbles: true }));
  const input = td.querySelector('input');
  if(!input) throw new Error('editCell: 未生成 input（dblclick 未绑定？）');
  input.value = String(val);
  input.dispatchEvent(new win.KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
}
const monthCell = key => doc.querySelector(`#report-table td.num[data-rowkey="${key}"]:not([data-idx])`);
const dayCell = (key, date) => doc.querySelector(`#report-table td.num[data-rowkey="${key}"][data-idx][data-date="${date}"]`);
const totalCell = () => doc.querySelector('#report-table td.num[data-rowkey="合计"]:not([data-idx])');

check('初始：商务部月累计=4,000', num(monthCell('商务部').textContent) === 4000, 'text=' + monthCell('商务部').textContent);
check('初始：合计月累计=4,000', num(totalCell().textContent) === 4000, 'text=' + totalCell().textContent);

// 场景A：编辑 商务部 9/10 每日格 = 555555 → 商务部月累计应联动为 1000+555555=556555
editCell(dayCell('商务部', '2026-09-10'), 555555);
check('A: 商务部 9/10 每日更新=555,555', num(dayCell('商务部','2026-09-10').textContent) === 555555, 'text=' + dayCell('商务部','2026-09-10').textContent);
check('A: 商务部月累计联动更新=556,555', num(monthCell('商务部').textContent) === 556555, 'text=' + monthCell('商务部').textContent);
check('A: 合计联动更新=556,555', num(totalCell().textContent) === 556555, 'text=' + totalCell().textContent);
// 一致性：单日"合计"单元格应随当日叶子部门覆盖同步（之前每日合计不反映当日手工覆盖 → 月合计≠Σ每日合计）
const dayTotalCell = date => doc.querySelector(`#report-table td.num[data-rowkey="合计"][data-idx][data-date="${date}"]`);
check('A: 合计行 9/10 每日=555,555（随当日商务部覆盖同步）', num(dayTotalCell('2026-09-10').textContent) === 555555, 'text=' + dayTotalCell('2026-09-10').textContent);
// 不变量：月合计 == 各日"合计"单元格之和
let sumDaily = 0;
doc.querySelectorAll('#report-table td.num[data-rowkey="合计"][data-idx]').forEach(td => { sumDaily += num(td.textContent); });
check('A: 月合计 == Σ每日合计 (' + sumDaily + ')', num(totalCell().textContent) === sumDaily, 'month=' + num(totalCell().textContent) + ' sum=' + sumDaily);

// 场景B：再编辑 商务部 月累计格=888,888 → 商务部月累计=888888，合计联动=888888
editCell(monthCell('商务部'), 888888);
check('B: 商务部月累计更新=888,888', num(monthCell('商务部').textContent) === 888888, 'text=' + monthCell('商务部').textContent);
check('B: 合计联动更新=888,888', num(totalCell().textContent) === 888888, 'text=' + totalCell().textContent);
check('B: state.manual.month["商务部"]=888888', app.state.manual.month['商务部'] === 888888, 'v=' + app.state.manual.month['商务部']);

// 场景E/F/G：单元格支持四则运算与相对增减。先清空手动覆盖回到基准(商务部=4000, 合计=4000)
app.state.manual.month = {}; app.state.manual.daily = {}; app.recompute();
check('E前: 商务部月累计回到基准=4,000', num(monthCell('商务部').textContent) === 4000, 'text=' + monthCell('商务部').textContent);
// E: 商务部 9/10 每日格输入表达式 "1000+500" → 1500；月累计联动=1000+1500=2500
editCell(dayCell('商务部', '2026-09-10'), '1000+500');
check('E: 商务部 9/10 每日=表达式结果 1,500', num(dayCell('商务部','2026-09-10').textContent) === 1500, 'text=' + dayCell('商务部','2026-09-10').textContent);
check('E: 商务部月累计联动=2,500', num(monthCell('商务部').textContent) === 2500, 'text=' + monthCell('商务部').textContent);
// F: 商务部 月累计格输入相对 "+2000"（基准 2,500）→ 4,500
editCell(monthCell('商务部'), '+2000');
check('F: 相对+2000 → 商务部月累计=4,500', num(monthCell('商务部').textContent) === 4500, 'text=' + monthCell('商务部').textContent);
// G: 商务部 月累计格输入表达式 "9000-1000" → 8,000（绝对表达式）
editCell(monthCell('商务部'), '9000-1000');
check('G: 表达式9000-1000 → 商务部月累计=8,000', num(monthCell('商务部').textContent) === 8000, 'text=' + monthCell('商务部').textContent);
check('G: 合计联动=8,000', num(totalCell().textContent) === 8000, 'text=' + totalCell().textContent);

// 场景C：门店表重渲染仍存在
check('C: 门店表行存在且重渲染', doc.querySelectorAll('#store-table tbody tr').length >= 1, 'rows=' + doc.querySelectorAll('#store-table tbody tr').length);

// 场景D：导出克隆 —— 日期列 90px（与页面每日列一致，不再吸收 width:100% 剩余宽度）、无导出专用样式覆盖（v-75）
const clone = app.buildExportClone(doc.getElementById('report-table'));
const allCols = [...clone.querySelectorAll('colgroup col')];
const visibleDaily = clone.querySelectorAll('th.daily-col').length;
const frontCount = allCols.length - visibleDaily;
const cols = allCols;
const dailyW = cols[frontCount] ? cols[frontCount].style.width : '(无)';
check('D: 导出日期列宽=90px（每日列标准宽，不再拉伸）', dailyW === '90px', 'width=' + dailyW);
check('D: 导出无 table-layout 覆盖（继承页面 auto）', !clone.style.tableLayout, 'tl=' + clone.style.tableLayout);
check('D: 导出仅含最新日期列(1列=9/10)', visibleDaily === 1, 'dailyCols=' + visibleDaily);
check('D: 最新日期列 9/10 出现在导出中', !!clone.querySelector('th.daily-col[data-date="2026-09-10"]'), 'has9/10=' + !!clone.querySelector('th.daily-col[data-date="2026-09-10"]'));
check('D: 非最新日期列 9/9 不出现在导出中', !clone.querySelector('th.daily-col[data-date="2026-09-09"]'), 'has9/9=' + !!clone.querySelector('th.daily-col[data-date="2026-09-09"]'));
check('D: 月目标表头含"万元"单位(未被裁切)', (()=>{ const t=[...clone.querySelectorAll('th.sub-head')].find(th=>th.textContent.includes('月目标')); return !!t && t.textContent.includes('万元'); })(), 'hdr=' + [...clone.querySelectorAll('th.sub-head')].map(t=>t.textContent.replace(/\s+/g,'')).join('|'));
// v-74 起导出不再注入内联样式：日期表头/数据格字号继承页面 CSS（12px），导出与页面展示一致
const dailyTh = clone.querySelector('th.daily-col');
const dailyTd = clone.querySelector('td.daily-col, td[class*="daily-col"]');
check('D: 导出日期表头无内联字号覆盖(继承页面CSS)', dailyTh && !dailyTh.style.fontSize, 'fs=' + (dailyTh && dailyTh.style.fontSize));
check('D: 导出日期数据格无内联字号覆盖(继承页面CSS)', dailyTd && !dailyTd.style.fontSize, 'fs=' + (dailyTd && dailyTd.style.fontSize));
// 导出图片日期列可容纳多位金额：宽度 90px（auto 布局下内容需要时会自然加宽，不截断）
check('D: 导出日期列宽足够容纳多位金额(>=90px)', parseInt(dailyW,10) >= 90, 'width=' + dailyW);

// 场景D2：Excel 导出——最新日期列之前的每日列自动隐藏（即使屏幕已展开全部旧日期列）
doc.querySelectorAll('#report-table [data-idx]').forEach(el => el.classList.remove('col-hidden'));
const grid = app.buildReportGrid();
const dailyIdxAll = [...doc.querySelectorAll('#report-table th.daily-col')].map(th => +th.dataset.idx);
const latestIdx = dailyIdxAll.length ? Math.max(...dailyIdxAll) : -1;
const latestSheetCol = 5 + latestIdx;
const preLatestSheetCols = dailyIdxAll.filter(i => i !== latestIdx).map(i => 5 + i);
const hiddenSet = new Set(grid.hiddenCols);
check('D2: Excel 隐藏了最新列之前的所有每日列', preLatestSheetCols.every(c => hiddenSet.has(c)), 'hidden=' + JSON.stringify(grid.hiddenCols));
check('D2: Excel 未隐藏最新日期列', !hiddenSet.has(latestSheetCol), 'latestSheetCol=' + latestSheetCol);

const failed = checks.filter(c => !c.c);
checks.forEach(c => console.log((c.c ? 'PASS' : 'FAIL') + ' | ' + c.n + ' | ' + c.d));
console.log('\n' + (failed.length ? 'FAILED ' + failed.length : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
