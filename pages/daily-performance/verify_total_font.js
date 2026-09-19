// v-77 回归：合计行数据字号放大一号（12px/12pt → 14px/14pt），其他一律不动。
// 环境限制：jsdom 不可用，采用源码文本断言确认变更落点与未牵连其他样式。
const fs = require('fs');
const IDX = 'index.html';
const XLSX = 'assets/xlsx-styled.js';
const checks = [];
function check(name, cond, extra) { checks.push({ name, pass: !!cond, extra: extra === undefined ? '' : String(extra) }); }

const idx = fs.readFileSync(IDX, 'utf8');
const xlsx = fs.readFileSync(XLSX, 'utf8');

// ---- A. 页面 合计行（tr.row-orange）字号改为 14px，其余规则保留 ----
// A1 关键 !important 规则：合计行 td / .cat 为 14px
const totalRule = idx.match(/table\.report tr\.row-orange td,\s*table\.report tr\.row-orange \.cat\s*\{\s*([^}]*)\}/);
check('页面存在 合计行 !important 规则', !!totalRule);
if (totalRule) {
  const body = totalRule[1];
  check('合计行字号=14px !important', /font-size:14px !important/.test(body), body.trim());
  check('合计行仍加粗 font-weight:700 !important', /font-weight:700 !important/.test(body));
  check('合计行仍黑字 color:#000 !important', /color:#000 !important/.test(body));
}

// A2 仅放大合计行：正文/数据格/表头/标题 字号均未被改（仍为原值）
check('正文 table.report 字号未变(12.5px)', /table\.report\{[^}]*font-size:12\.5px/.test(idx.replace(/\s+/g, ' ')));
check('普通数据格 .num 字号未变(12px)', /table\.report \.num\{[^}]*font-size:12px/.test(idx.replace(/\s+/g, ' ')));
check('列标题 .sub-head 字号未变(12px)', /table\.report \.sub-head\{[^}]*font-size:12px/.test(idx.replace(/\s+/g, ' ')));
check('大标题 .title-cell 字号未变(15.5px)', /table\.report \.title-cell\{[^}]*font-size:15\.5px/.test(idx.replace(/\s+/g, ' ')));

// ---- B. Excel 合计行 font3 字号改为 14pt，底色橙不变 ----
check('xlsx font3 合计字号=14pt', /font3: 合计行黑粗14/.test(xlsx));
check('xlsx font3 实际 sz val="14"', /<font><b\/><sz val="14"\/><color rgb="FF000000"\/><name val="Microsoft YaHei"\/><\/font>' \+ {0,1}\s*\/\/ font3/.test(xlsx));
check('xlsx 合计行底色橙 fill:7 未变', /orange:\s*\{f:3, fill:7/.test(xlsx));
check('xlsx 列标题 font4 仍 12pt（未改动）', /font4 列标题黑粗12/.test(xlsx));

// ---- C. 版本号 bump 至 77，旧 76 残留清零 ----
const app = fs.readFileSync('assets/app.js', 'utf8');
check('app.js 无 v20260910-76 残留', (app.match(/v20260910-76/g) || []).length === 0, '剩 ' + (app.match(/v20260910-76/g) || []).length);
check('app.js 含 v20260910-77', (app.match(/v20260910-77/g) || []).length >= 5, '共 ' + (app.match(/v20260910-77/g) || []).length);
check('index.html 无 ?v=20260910-76 残留', (idx.match(/\?v=20260910-76/g) || []).length === 0, '剩 ' + (idx.match(/\?v=20260910-76/g) || []).length);
check('index.html 含 ?v=20260910-77', (idx.match(/\?v=20260910-77/g) || []).length === 2, '共 ' + (idx.match(/\?v=20260910-77/g) || []).length);

// ---- 汇总 ----
let allPass = true;
for (const c of checks) {
  if (!c.pass) allPass = false;
  console.log((c.pass ? 'PASS' : 'FAIL') + ' | ' + c.name + (c.extra ? ' | ' + c.extra : ''));
}
console.log(allPass ? '\nALL_PASS (' + checks.length + ' checks)' : '\nFAILED (' + checks.filter(c => !c.pass).length + ' checks)');
process.exit(allPass ? 0 : 1);
