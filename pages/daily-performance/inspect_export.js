// inspect_export.js — 解析真实导出的 xlsx，打印标题行/列标题行/合计行的实际样式索引与字体填充定义
const fs = require('fs');
const JSZip = require('D:/workbuddy/2026-08-15-10-10-03/daily-performance/assets/vendor/jszip.min.js');
const file = 'D:/Download/标题填充验证-2026-09-17.xlsx';
(async () => {
  const buf = fs.readFileSync(file);
  const zip = await JSZip.loadAsync(buf);
  const styles = await zip.file('xl/styles.xml').async('string');
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');

  const cellXfs = (styles.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [''])[0];
  const xfs = (cellXfs.match(/<xf\b[^>]*>/g) || []);
  console.log('=== cellXfs count:', xfs.length, '===');
  xfs.forEach((x, i) => {
    const fid = (x.match(/fontId="(\d+)"/) || [])[1];
    const fld = (x.match(/fillId="(\d+)"/) || [])[1];
    console.log('xf[' + i + '] fontId=' + fid + ' fillId=' + fld + '  ' + x);
  });

  const fonts = (styles.match(/<fonts[\s\S]*?<\/fonts>/) || [''])[0];
  const fontEls = (fonts.match(/<font>[\s\S]*?<\/font>/g) || []);
  fontEls.forEach((f, i) => console.log('font[' + i + ']', f));

  const fills = (styles.match(/<fills[\s\S]*?<\/fills>/) || [''])[0];
  const fillEls = (fills.match(/<fill>[\s\S]*?<\/fill>/g) || []);
  fillEls.forEach((f, i) => console.log('fill[' + i + ']', f));

  console.log('=== SHEET rows 1..4 ===');
  for (let r = 1; r <= 4; r++) {
    const rowM = sheet.match(new RegExp('<row r="' + r + '"[\\s\\S]*?</row>'));
    if (!rowM) { console.log('row ' + r + ' NOT FOUND'); continue; }
    const row = rowM[0];
    const cells = (row.match(/<c [^>]*>/g) || []);
    console.log('--- ROW ' + r + ' (cells=' + cells.length + ') ---');
    cells.forEach(c => {
      const ref = (c.match(/r="([^"]+)"/) || [])[1];
      const s = (c.match(/ s="(\d+)"/) || [])[1];
      const t = (c.match(/ t="([^"]+)"/) || [])[1];
      console.log('  ', ref, 's=' + s, 't=' + t);
    });
  }
  // 找标题文本所在单元格
  const titleCell = sheet.match(/各部门业绩完成情况跟踪日报表[^<]*/);
  console.log('=== title text present:', !!titleCell, titleCell ? titleCell[0].slice(0, 30) : '');
})().catch(e => { console.error(e); process.exit(1); });
