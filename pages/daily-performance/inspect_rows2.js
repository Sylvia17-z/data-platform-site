// inspect_rows2.js — 更彻底：容错匹配所有 <row 标签（含自闭合），并打印 sheet1.xml 尾部原文
const fs = require('fs');
const JSZip = require('D:/workbuddy/2026-08-15-10-10-03/daily-performance/assets/vendor/jszip.min.js');
const files = process.argv.slice(2);
(async () => {
  for (const file of files) {
    const buf = fs.readFileSync(file);
    const zip = await JSZip.loadAsync(buf);
    const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');
    console.log('======', file, ' len=', sheet.length);
    // 容错：所有 <row ...> / <row .../> 开标签
    const opens = sheet.match(/<row\b[^>]*>/g) || [];
    console.log('<row open tags> =', opens.length);
    opens.forEach(o => console.log('  ', o.slice(0, 60)));
    // mergeCells
    const mc = sheet.match(/<mergeCells count="(\d+)"/);
    console.log('mergeCells count =', mc ? mc[1] : 0);
    const refs = sheet.match(/<mergeCell ref="[^"]+"\/>/g) || [];
    refs.forEach(x => console.log('  ', x));
    // 尾部 600 字符
    console.log('--- tail ---');
    console.log(sheet.slice(-600));
    console.log('--- sheetView/pane ---');
    const sv = sheet.match(/<sheetViews[\s\S]*?<\/sheetViews>/);
    console.log(sv ? sv[0] : 'none');
  }
})().catch(e => { console.error(e); process.exit(1); });
