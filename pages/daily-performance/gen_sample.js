// gen_sample.js — 用修复后的 xlsx-styled.js 生成一个「与真实导出同结构」的 13 行 xlsx，供直接打开验证：
//   1) 标题行藏蓝 + 白字 15pt（v-73 由 18pt 收小）
//   2) 列宽收窄（v-73）
//   3) 数据区下方/右侧空白区域为白色（v-73：cellXfs[0] 中性默认样式，修复"空白行全藏蓝"）
const fs = require('fs');
const JSZip = require('D:/workbuddy/2026-08-15-10-10-03/daily-performance/assets/vendor/jszip.min.js');
global.window = { JSZip };
const code = fs.readFileSync('D:/workbuddy/2026-08-15-10-10-03/daily-performance/assets/xlsx-styled.js', 'utf8');
eval(code);

const cell = (v, t, s) => ({ v, t: t || 's', s: s || 'num' });
const nCols = 6; // A/B/C=部门类别(3列) + D=月目标 + E=月累计 + F=唯一可见日期列
const rows = [];
const merges = [];
// 行1：标题（藏蓝，整行合并）
rows.push([cell('各部门业绩完成情况跟踪日报表(截止9月16日)', 's', 'title'), null, null, null, null, null]);
merges.push({ r: 0, c: 0, r2: 0, c2: nCols - 1 });
// 行2：列标题（浅蓝）
rows.push([cell('部门类别', 's', 'subhead'), null, null, cell('月目标(万元)', 's', 'subhead'), cell('月累计\n(单位：元)', 's', 'subhead'), cell('9月16日', 's', 'subhead')]);
merges.push({ r: 1, c: 0, r2: 1, c2: 2 });
// 行3-7：营销部块（主列 rowspan5 / 公司支持 rowspan2 / 部门业绩 rowspan3 / 月目标 3 行合并）
rows.push([cell('营销部', 's', 'cat'), cell('公司支持', 's', 'catsub'), cell('品牌组-其他', 's', 'catsub'), cell(400, 'n', 'target'), cell(0, 'n', 'num'), cell(0, 'n', 'num')]);
rows.push([null, null, cell('活动组-活动支持', 's', 'catsub'), null, cell(0, 'n', 'num'), cell(0, 'n', 'num')]);
rows.push([null, cell('部门业绩', 's', 'catsub'), cell('电商组', 's', 'catsub'), cell(77, 'n', 'target'), cell(134284.90, 'n', 'num'), cell(72731.71, 'n', 'num')]);
rows.push([null, null, cell('短视频组', 's', 'catsub'), cell(3, 'n', 'target'), cell(0, 'n', 'num'), cell(0, 'n', 'num')]);
rows.push([null, null, cell('品牌组-自然到店', 's', 'catsub'), null, cell(71097.00, 'n', 'num'), cell(0, 'n', 'num')]);
merges.push({ r: 2, c: 0, r2: 6, c2: 0 }, { r: 2, c: 1, r2: 3, c2: 1 }, { r: 4, c: 1, r2: 6, c2: 1 }, { r: 2, c: 3, r2: 3, c2: 3 }, { r: 5, c: 3, r2: 6, c2: 3 });
// 行8-9：商务部 / 其他
rows.push([cell('商务部', 's', 'cat'), null, null, cell(125, 'n', 'target'), cell(519928.25, 'n', 'num'), cell(66472.63, 'n', 'num')]);
rows.push([cell('其他', 's', 'cat'), null, null, cell(170, 'n', 'target'), cell(868507.03, 'n', 'num'), cell(526531.94, 'n', 'num')]);
merges.push({ r: 7, c: 0, r2: 7, c2: 2 }, { r: 8, c: 0, r2: 8, c2: 2 });
// 行10：合计（橙）
rows.push([cell('合计', 's', 'orange'), null, null, cell(400, 'n', 'orange'), cell(1593817.18, 'n', 'orange'), cell(540405.13, 'n', 'orange')]);
merges.push({ r: 9, c: 0, r2: 9, c2: 2 });
// 行11-13：基础业绩 / 新增业绩 / 管家部
rows.push([cell('基础业绩 （新增业绩之外）', 's', 'cat'), null, null, cell(350, 'n', 'target'), cell(1280005.65, 'n', 'num'), cell(540405.13, 'n', 'num')]);
rows.push([cell('新增业绩 （2026新增VAB）', 's', 'cat'), null, null, cell(50, 'n', 'target'), cell(313811.53, 'n', 'num'), cell(0, 'n', 'num')]);
rows.push([cell('管家部', 's', 'cat'), null, null, cell(270, 'n', 'target'), cell(1199232.52, 'n', 'num'), cell(506221.68, 'n', 'num')]);
merges.push({ r: 10, c: 0, r2: 10, c2: 2 }, { r: 11, c: 0, r2: 11, c2: 2 }, { r: 12, c: 0, r2: 12, c2: 2 });

const colWidths = [10.5, 9.5, 16, 12, 13, 11]; // v-73 收窄后口径

(async () => {
  const blob = await global.window.buildStyledXlsx({ sheetName: '日业绩', rows, merges, colWidths, hiddenCols: [] });
  const buf = Buffer.from(await blob.arrayBuffer());
  const out = 'D:/Download/样式修复验证-v73-2026-09-17.xlsx';
  fs.writeFileSync(out, buf);
  console.log('written', out, buf.length, 'bytes');
})().catch(e => { console.error(e); process.exit(1); });
