// v-76 回归：仅导出侧框线调细，页面 CSS 与其余导出样式均不动。
// 环境限制：jsdom 不可用，故采用「源码文本断言」确认变更落点，而非运行时 DOM 断言。
const fs = require('fs');
const APP = 'assets/app.js';
const IDX = 'index.html';
const checks = [];
function check(name, cond, extra) { checks.push({ name, pass: !!cond, extra: extra === undefined ? '' : String(extra) }); }

const app = fs.readFileSync(APP, 'utf8');
const idx = fs.readFileSync(IDX, 'utf8');

// ---- A. 页面 index.html 边框保持 1px（未被改动）----
const pageOuterOk = /table\.report\{[^}]*border:1px solid #000000;/.test(idx.replace(/\s+/g, ' '));
check('页面 table.report 外框仍为 1px（未改动）', pageOuterOk);

const pageCellOk = /table\.report th,table\.report td\{[^}]*border:1px solid #000000;/.test(idx.replace(/\s+/g, ' '));
check('页面 th/td 边框仍为 1px（未改动）', pageCellOk);

// ---- B. app.js 新增「仅导出侧」边框调细（步骤 9）----
check('buildExportClone 含 v-76 步骤 9 注释', app.includes('// 9)（v-76）'));
check('定义导出细边框常量 0.5px', app.includes("const EXPORT_BORDER_W = '0.5px'"), (app.match(/EXPORT_BORDER_W = '([^']+)'/) || [])[1]);
check('克隆外框 borderWidth 改写为细框', app.includes('clone.style.borderWidth = EXPORT_BORDER_W;'));
check('克隆单元格 th/td borderWidth 改写为细框',
  /clone\.querySelectorAll\('th, td'\)\.forEach\([^)]*\)\s*\{[^}]*cell\.style\.borderWidth = EXPORT_BORDER_W;/.test(app));

// ---- C. 版本号 bump 至 76，旧 75 残留清零 ----
const app75 = (app.match(/v20260910-75/g) || []).length;
const app76 = (app.match(/v20260910-76/g) || []).length;
const app77 = (app.match(/v20260910-77/g) || []).length;
check('app.js 无 v20260910-75 残留', app75 === 0, '剩 ' + app75);
check('app.js 无 v20260910-76 残留', app76 === 0, '剩 ' + app76);
check('app.js 含 v20260910-77（>=5 处标记）', app77 >= 5, '共 ' + app77);

const idx75 = (idx.match(/\?v=20260910-75/g) || []).length;
const idx76 = (idx.match(/\?v=20260910-76/g) || []).length;
const idx77 = (idx.match(/\?v=20260910-77/g) || []).length;
check('index.html 无 ?v=20260910-75 残留', idx75 === 0, '剩 ' + idx75);
check('index.html 无 ?v=20260910-76 残留', idx76 === 0, '剩 ' + idx76);
check('index.html 含 ?v=20260910-77（2 处 script）', idx77 === 2, '共 ' + idx77);

// ---- 汇总 ----
let allPass = true;
for (const c of checks) {
  if (!c.pass) allPass = false;
  console.log((c.pass ? 'PASS' : 'FAIL') + ' | ' + c.name + (c.extra ? ' | ' + c.extra : ''));
}
console.log(allPass ? '\nALL_PASS (' + checks.length + ' checks)' : '\nFAILED (' + checks.filter(c => !c.pass).length + ' checks)');
process.exit(allPass ? 0 : 1);
