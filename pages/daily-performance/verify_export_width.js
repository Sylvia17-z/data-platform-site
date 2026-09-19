// verify_export_width.js — verifies the pure export column-width planner computeExportColPx()
// introduced in v-75: the export clone's single daily col no longer absorbs the leftover width
// of the width:100% page table. Front cols lock to the live rendered px (fallback: colgroup
// spec — % scaled by refW, px used as-is); the daily col keeps its normal width (90px default,
// or the user-dragged live spec).
const fs = require('fs');
const path = require('path');
const appJs = path.join(__dirname, 'assets', 'app.js');
const src = fs.readFileSync(appJs, 'utf8');

// Extract the real pure function from app.js and eval it.
const m = src.match(/function computeExportColPx\([\s\S]*?\n}/);
if (!m) { console.error('FAIL | could not locate computeExportColPx in app.js'); process.exit(1); }
// eslint-disable-next-line no-eval
const fn = new Function(m[0] + '\nreturn computeExportColPx;')();

const checks = [];
function check(name, cond, detail) { checks.push({ name, pass: !!cond, detail }); }

// 结构自检：app.js 应包含量测函数，且不再有导出专用布局覆盖
check('app.js has measureLiveColWidths()', /function measureLiveColWidths\(/.test(src), /function measureLiveColWidths\(/.test(src));
check('app.js has no tableLayout override', !/clone\.style\.tableLayout/.test(src), !/clone\.style\.tableLayout/.test(src));
check('app.js step-8 uses computeExportColPx', /computeExportColPx\(frontCount, measured, liveSpec/.test(src), /computeExportColPx\(frontCount, measured, liveSpec/.test(src));

// 场景1：真实浏览器，前导列实测像素宽可用 → 前导列 = 实测，日期列 = 90px（不再吸收剩余宽度）
let r = fn(5, [82, 121, 137, 143, 204], ['7.5%','11%','12.5%','13%','18.5%','90px'], 1100, '90px', 90);
check('S1 front px = measured', JSON.stringify(r.frontPx) === '[82,121,137,143,204]', JSON.stringify(r.frontPx));
check('S1 daily = 90px (not stretched)', r.dailyPx === 90, 'dailyPx=' + r.dailyPx);
check('S1 total = 687+90 = 777', r.total === 777, 'total=' + r.total);

// 场景2：量不到（jsdom/无布局，实测全 0） → 退回 colgroup 百分比 × refW（Math.round）
r = fn(5, [0, 0, 0, 0, 0], ['7.5%','11%','12.5%','13%','18.5%','90px'], 1100, '90px', 90);
check('S2 front fallback = %xrefW', JSON.stringify(r.frontPx) === '[83,121,138,143,204]', JSON.stringify(r.frontPx));
check('S2 total = 689+90 = 779', r.total === 779, 'total=' + r.total);

// 场景3：用户拖拽过日期列（线上规格 '120px'） → 日期列 120px（尊重拖拽定制）
r = fn(5, [82, 121, 137, 143, 204], ['7.5%','11%','12.5%','13%','18.5%','120px'], 1100, '120px', 90);
check('S3 dragged daily width honored', r.dailyPx === 120, 'dailyPx=' + r.dailyPx);
check('S3 total = 687+120 = 807', r.total === 807, 'total=' + r.total);

// 场景4：日期列线上规格是百分比 → ×refW 换算
r = fn(5, [82, 121, 137, 143, 204], ['7.5%','11%','12.5%','13%','18.5%','10%'], 1000, '10%', 90);
check('S4 daily % spec converted', r.dailyPx === 100, 'dailyPx=' + r.dailyPx);
check('S4 total = 687+100 = 787', r.total === 787, 'total=' + r.total);

// 场景5：全空兜底（无实测、无规格、refW=0） → 日期列 = DAILY_COL_PX 兜底
r = fn(5, [0, 0, 0, 0, 0], ['', '', '', '', '', ''], 0, '', 90);
check('S5 empty fallback daily=90', r.dailyPx === 90, 'dailyPx=' + r.dailyPx);
check('S5 empty fallback total=90', r.total === 90, 'total=' + r.total);

// 场景6：门店表（2 前导列 + 1 日期列）
r = fn(2, [176, 220], ['16%','20%','90px','36px'], 1100, '90px', 90);
check('S6 store front = measured', JSON.stringify(r.frontPx) === '[176,220]', JSON.stringify(r.frontPx));
check('S6 store total = 396+90 = 486', r.total === 486, 'total=' + r.total);

const failed = checks.filter(c => !c.pass);
checks.forEach(c => console.log((c.pass ? 'PASS' : 'FAIL') + ' | ' + c.name + ' | ' + c.detail));
console.log('\n' + (failed.length ? ('FAILED ' + failed.length) : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
