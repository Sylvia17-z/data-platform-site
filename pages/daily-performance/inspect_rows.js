// inspect_rows.js — 解析真实导出的 xlsx，统计总行数并按行打印样式分布（定位"空白行全藏蓝"来源）
const fs = require('fs');
const JSZip = require('D:/workbuddy/2026-08-15-10-10-03/daily-performance/assets/vendor/jszip.min.js');
const file = process.argv[2] || 'D:/Download/日业绩-2026-09-16(4).xlsx';
(async () => {
  const buf = fs.readFileSync(file);
  const zip = await JSZip.loadAsync(buf);
  const styles = await zip.file('xl/styles.xml').async('string');
  const sheet = await zip.file('xl/worksheets/sheet1.xml').async('string');

  // fill 色表
  const fills = (styles.match(/<fills[\s\S]*?<\/fills>/) || [''])[0];
  const fillEls = (fills.match(/<fill>[\s\S]*?<\/fill>/g) || []);
  const fillRgb = fillEls.map(f => (f.match(/fgColor rgb="([0-9A-F]+)"/) || [])[1] || 'none');

  // xf -> fillId
  const cellXfs = (styles.match(/<cellXfs[\s\S]*?<\/cellXfs>/) || [''])[0];
  const xfs = (cellXfs.match(/<xf\b[^>]*>/g) || []);
  const xfFill = xfs.map(x => +((x.match(/fillId="(\d+)"/) || [])[1] || 0));
  const xfFont = xfs.map(x => +((x.match(/fontId="(\d+)"/) || [])[1] || 0));
  console.log('xfFill =', xfFill.join(','), '  fillRgb =', fillRgb.join(','));

  // 全部行
  const rowRe = /<row r="(\d+)"([^>]*)>([\s\S]*?)<\/row>/g;
  let m, count = 0;
  const rowInfo = [];
  while((m = rowRe.exec(sheet))){
    count++;
    const r = +m[1];
    const cells = (m[3].match(/<c [^>]*?(?:\/>|><\/c>|>)/g) || []);
    const stylesUsed = {};
    let nonEmpty = 0;
    cells.forEach(c => {
      const s = (c.match(/ s="(\d+)"/) || [])[1];
      const hasVal = /<(v|is)>/.test(c);
      if(hasVal) nonEmpty++;
      stylesUsed[s] = (stylesUsed[s]||0)+1;
    });
    rowInfo.push({r, ht:(m[2].match(/ht="([\d.]+)"/)||[])[1], ncells:cells.length, nonEmpty, stylesUsed});
  }
  console.log('TOTAL ROWS =', count);
  rowInfo.forEach(x => {
    const styleDesc = Object.entries(x.stylesUsed).map(([s,n]) => 's'+s+'(fill '+xfFill[+s]+'='+fillRgb[xfFill[+s]]+' x'+n+')').join(' ');
    console.log('row '+x.r+'  ht='+x.ht+'  cells='+x.ncells+'  withVal='+x.nonEmpty+'  ['+styleDesc+']');
  });
  // 最后一个有值行之后的行
  const lastValRow = Math.max(...rowInfo.filter(x=>x.nonEmpty>0).map(x=>x.r));
  console.log('last row WITH VALUE =', lastValRow, '; rows after it:', rowInfo.filter(x=>x.r>lastValRow).length);
})().catch(e => { console.error(e); process.exit(1); });
