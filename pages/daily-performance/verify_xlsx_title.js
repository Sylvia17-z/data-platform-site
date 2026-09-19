/* verify_xlsx_title.js
 * 回归：导出 Excel 的标题行需与页面 .title-cell 一致 ——
 * 藏蓝背景(#002060) + 白色字体 + 加粗 + 15pt(v-73 由18pt收小) + 居中，且跨整行合并区均为该样式。
 * 另含 v-73 关键回归：cellXfs[0] 必须为中性默认样式（否则 WPS 把数据区外空白格套用 xf0 → 空白行全藏蓝）。
 * 用法：node verify_xlsx_title.js
 */
const fs = require('fs');
const path = require('path');
const JSZip = require(path.resolve('assets/vendor/jszip.min.js'));
global.window = { JSZip };
const code = fs.readFileSync(path.resolve('assets/xlsx-styled.js'), 'utf8');
eval(code);

let pass = 0, fail = 0;
function check(name, cond, info) {
  if (cond) { console.log('PASS | ' + name + (info ? ' | ' + info : '')); pass++; }
  else { console.log('FAIL | ' + name + (info ? ' | ' + info : '')); fail++; }
}
function strip(xml, tag) {
  const m = xml.match(new RegExp('<' + tag + '[\\s\\S]*?<\\/' + tag + '>'));
  return m ? m[0] : '';
}

(async () => {
  const cell = (v, t, s) => ({ v, t: t || 's', s: s || 'num' });
  const nCols = 6;
  const rows = [
    [cell('各部门业绩完成情况跟踪日报表(截止2026-09)', 's', 'title'), null, null, null, null, null],
    [cell('部门类别', 's', 'subhead'), cell('部门类别', 's', 'subhead'), cell('项目', 's', 'subhead'), cell('月目标(万元)', 's', 'subhead'), cell('月累计\n(单位：元)', 's', 'subhead'), cell('9月16日', 's', 'subhead')],
    [cell('合计', 's', 'orange'), cell('', 's', 'cat'), cell('', 's', 'cat'), cell(0, 'n', 'target'), cell(63663.96, 'n', 'num'), cell(63663.96, 'n', 'num')],
  ];
  const merges = [{ r: 0, c: 0, r2: 0, c2: nCols - 1 }];
  const colWidths = [16, 16, 16, 16, 18, 13];
  const blob = await global.window.buildStyledXlsx({ sheetName: '日业绩', rows, merges, colWidths, hiddenCols: [] });
  const buf = Buffer.from(await blob.arrayBuffer());
  const zip = await JSZip.loadAsync(buf);
  const styles = await zip.file('xl/styles.xml').async('string');
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');

  const fonts = strip(styles, 'fonts');
  const fills = strip(styles, 'fills');
  const cellXfs = strip(styles, 'cellXfs');
  // v-73：cellXfs[0] = 中性默认样式（自闭合 <xf .../>），业务样式整体后移一位（title=1, subhead=2, ..., orange=8）。
  // 注意不能用 split('</xf>') 索引——中性 xf 是自闭合的，没有 </xf> 闭合标签。
  // 两个完整模式并列：先尝试自闭合 <xf .../>，再尝试成对 <xf ...>...</xf>，
  // 避免共享后缀的 (?:\/>|>...) 贪婪分支把自闭合 xf 吞进下一个成对 xf。
  const xfs = cellXfs.match(/<xf\b[^>]*\/>|<xf\b[^>]*>[\s\S]*?<\/xf>/g) || [];
  const xf0 = xfs[0] || '';
  const titleXf = xfs[1] || '';
  const titleFont = fonts.split('</font>').filter(s => s.includes('<font')).map(s => s + '</font>')[1];
  const navyFill = fills.split('</fill>').filter(s => s.includes('<fill')).map(s => s + '</fill>')[2];

  // 0) v-73 关键回归：cellXfs[0] 必须是中性默认样式（fillId=0 且 borderId=0），否则空白区域全藏蓝
  check('cellXfs[0] 为中性默认样式(fillId=0/borderId=0)——修复空白行全藏蓝',
    /fillId="0"/.test(xf0) && /borderId="0"/.test(xf0) && /\/>$/.test(xf0) && !/fillId="[1-9]/.test(xf0) && !/borderId="[1-9]/.test(xf0),
    'xf0=' + xf0);

  // 1) 标题样式：藏蓝填充
  check('标题样式使用藏蓝填充(fillId=2 = #002060)', /fillId="2"/.test(titleXf) && /FF002060/.test(navyFill), 'xf=' + titleXf + ' fill=' + navyFill);
  // 2) 标题字体：白色 + 加粗 + 15pt（v-73：用户要求 15 号即可，不再放大到 18）
  const sz = (titleFont.match(/sz val="(\d+)"/) || [])[1];
  check('标题字体白色(bold + FFFFFFFF)', /<b\/>/.test(titleFont) && /FFFFFFFF/.test(titleFont), 'font=' + titleFont);
  check('标题字号15pt(当前 ' + sz + ')', sz && parseInt(sz, 10) === 15, 'sz=' + sz);
  // 3) 标题居中
  check('标题居中对齐(水平+垂直)', /horizontal="center"[^>]*vertical="center"|vertical="center"[^>]*horizontal="center"/.test(titleXf), 'xf=' + titleXf);
  // 4) 标题行高度足够容纳放大字号
  const row1 = sheet.match(/<row r="1"[\s\S]*?<\/row>/)[0];
  check('标题行高 >=32pt(容纳放大字号)', /ht="(\d+)"/.test(row1) && parseInt(row1.match(/ht="(\d+)"/)[1], 10) >= 32, row1.match(/ht="\d+"/)[0]);
  // 5) 合并区内每一格都是标题样式(s=1) —— 整行藏蓝
  const titleCells = (row1.match(/s="(\d+)"/g) || []).map(s => s.match(/\d+/)[0]);
  check('标题合并区全格为标题样式(s=1)', titleCells.length === nCols && titleCells.every(s => s === '1'), 'cells=' + titleCells.join(','));

  // 6) 列标题行(行2：部门类别/月目标/月累计/日期) 还原为浅蓝底 + 黑字加粗 + 12pt + 居中
  const subheadXf = xfs[2] || ''; // STYLE_MAP 第2个键 = subhead（v-73 后 xf 索引 2）
  const subheadFont = fonts.split('</font>').filter(s => s.includes('<font')).map(s => s + '</font>')[4];
  check('列标题行样式：浅蓝填充(fillId=3 = #DAE3F5)', /fillId="3"/.test(subheadXf), 'xf=' + subheadXf);
  check('列标题行样式：黑字加粗(FF000000+<b/>)', /<b\/>/.test(subheadFont) && /FF000000/.test(subheadFont), 'font=' + subheadFont);
  const subSz = (subheadFont.match(/sz val="(\d+)"/) || [])[1];
  check('列标题行字号12pt(当前 ' + subSz + ')', subSz && parseInt(subSz, 10) === 12, 'sz=' + subSz);
  const row2 = sheet.match(/<row r="2"[\s\S]*?<\/row>/)[0];
  const subCells = (row2.match(/s="(\d+)"/g) || []).map(s => s.match(/\d+/)[0]);
  check('列标题行全格为列标题样式(s=2, 浅蓝+黑粗)', subCells.length === nCols && subCells.every(s => s === '2'), 'cells=' + subCells.join(','));

  // 7) 合计行(orange) 字号改为 14pt 加粗（v-71 由 15pt 改 12pt，v-77 放大一号→14pt）
  const orangeXf = xfs[8] || ''; // STYLE_MAP 末键 = orange（v-73 后 xf 索引 8）
  const orangeFont = fonts.split('</font>').filter(s => s.includes('<font')).map(s => s + '</font>)[3];
  const oSz = (orangeFont.match(/sz val="(\d+)"/) || [])[1];
  check('合计行字号14pt(当前 ' + oSz + ')', oSz && parseInt(oSz, 10) === 14, 'sz=' + oSz);
  check('合计行底色橙(fillId=7 = #FFBA84)', /fillId="7"/.test(orangeXf), 'xf=' + orangeXf);

  // 8) 兼容性(v-72)：实色填充的 bgColor 必须同为 rgb，不得再用 indexed="64"
  //    （WPS/部分查看器会把 indexed=64 当作系统背景=白色渲染，导致藏蓝/浅蓝整行变白）
  const hasIndexedBg = /bgColor indexed="64"/.test(fills);
  check('实色填充不再使用 bgColor indexed="64"', !hasIndexedBg, 'hasIndexedBg=' + hasIndexedBg);
  check('藏蓝填充 bgColor 同为 rgb(FF002060)', /bgColor rgb="FF002060"/.test(fills), 'fills=' + (fills.match(/bgColor rgb="FF002060"/) || [''])[0]);
  check('浅蓝填充 bgColor 同为 rgb(FFDAE3F5)', /bgColor rgb="FFDAE3F5"/.test(fills), 'fills=' + (fills.match(/bgColor rgb="FFDAE3F5"/) || [''])[0]);

  // 9) 兼容(v-72)：styles.xml 须含 <cellStyles>，否则 WPS 整体忽略 cellXfs（整表显白）
  check('styles.xml 含 <cellStyles>(WPS 兼容)', /<cellStyles/.test(styles), 'has=' + /<cellStyles/.test(styles));

  console.log((fail === 0 ? '\nALL_PASS' : '\nFAILED') + ' (' + (pass + fail) + ' checks)');
  process.exit(fail === 0 ? 0 : 1);
})().catch(e => { console.error(e); process.exit(1); });
