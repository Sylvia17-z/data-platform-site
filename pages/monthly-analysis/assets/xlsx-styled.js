/* xlsx-styled.js (monthly-analysis 定制版)
 * 用 JSZip 手写 OOXML，生成「与线上显示完全一致」的带样式 Excel（藏蓝主题 + 黑框 + 合并单元格）。
 * SheetJS 社区版在本环境无法写入单元格填充/字体样式，故走原生 OOXML（同 daily-performance 方案，
 * 保留 v-72（bgColor 同色修复）与 v-73（cellXfs[0] 中性默认样式）两个 WPS/Excel 兼容修复）。
 * 暴露 window.buildStyledXlsx({ sheetName, rows, merges, colWidths, freezeRows }) -> Promise<Blob>
 *   rows:   二维数组，每个元素为 {v, t:'n'|'s', s:<styleKey>} 或 null（被合并覆盖的格子）
 *   merges: [{r,c,r2,c2}] 0-based 合并区域（仅左上角写值，其余置 null）
 *   colWidths: 每列字符宽度（可选）
 *   freezeRows: 冻结表头行数（可选，默认 0）
 * 样式与页面 CSS 一一对应：
 *   表头 navy#002060 白粗 / 部门列 #d6e0f0 / 科室列 #f3f6fb / 医生列 白底左对齐
 *   斑马行 #fbfcfe / 合计行 #D7EFF9 / 全院合计 #d6e0f0 / 参考行 #f7f7f7 斜体灰 / 「/」灰字
 */
(function(){
  function esc(s){
    return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
  }
  function colLetter(c){
    let s=''; c=c+1;
    while(c>0){ const m=(c-1)%26; s=String.fromCharCode(65+m)+s; c=Math.floor((c-1)/26); }
    return s;
  }

  // numFmtId：0=General 3=#,##0 4=#,##0.00 164=0.0%（自定义）
  // font0 常规黑 / font1 白粗(表头) / font2 #003366粗(部门/合计) / font3 #0F172A(正文) /
  // font4 #5B6472(灰字/「/」) / font5 #5B6472斜体(参考行)
  // fill2 藏蓝#002060 / fill3 #d6e0f0 / fill4 #f3f6fb / fill5 白 / fill6 #fbfcfe / fill7 #f7f7f7 / fill8 #D7EFF9
  const STYLE_MAP = {
    title:   {f:1, fill:2, bd:1, nf:0,   align:'center'},   // 表头行：藏蓝底 + 白字加粗 + 居中
    dept:    {f:2, fill:3, bd:1, nf:0,   align:'center'},   // 部门列：#d6e0f0 底 + #003366 粗体居中
    clinic:  {f:2, fill:4, bd:1, nf:0,   align:'center'},   // 科室列：#f3f6fb 底 + #003366 粗体居中
    doc:     {f:3, fill:5, bd:1, nf:0,   align:'left'},     // 医生/名称列：白底左对齐
    docZ:    {f:3, fill:6, bd:1, nf:0,   align:'left'},     // 斑马偶数行：#fbfcfe
    docSoft: {f:4, fill:5, bd:1, nf:0,   align:'left'},     // 「无」行名称：灰字
    ref:     {f:5, fill:7, bd:1, nf:0,   align:'left'},     // 参考行名称：#f7f7f7 斜体灰
    refPlain:{f:5, fill:7, bd:1, nf:0,   align:'right'},    // 参考行目标列
    refN:    {f:5, fill:7, bd:1, nf:4,   align:'right'},    // 参考行金额
    refPct:  {f:5, fill:7, bd:1, nf:164, align:'right'},    // 参考行完成率
    num:     {f:3, fill:5, bd:1, nf:4,   align:'right'},    // 金额：白底右对齐 #,##0.00
    numZ:    {f:3, fill:6, bd:1, nf:4,   align:'right'},    // 金额斑马行
    numInt:  {f:3, fill:5, bd:1, nf:3,   align:'right'},    // 容量：白底右对齐 #,##0
    dim:     {f:4, fill:5, bd:1, nf:0,   align:'right'},    // 「/」：灰字右对齐
    plain:   {f:3, fill:5, bd:1, nf:0,   align:'right'},    // 目标列：General 右对齐
    pct:     {f:3, fill:5, bd:1, nf:164, align:'right'},    // 完成率：0.0%
    pctDim:  {f:4, fill:5, bd:1, nf:0,   align:'right'},    // 完成率「/」
    sumL:    {f:2, fill:8, bd:1, nf:0,   align:'center'},   // 部门合计行：#D7EFF9 粗体
    sumDoc:  {f:2, fill:8, bd:1, nf:0,   align:'left'},     // 表B「管家合计」标签：#D7EFF9 粗体左对齐（页面 td.doc）
    sumN:    {f:2, fill:8, bd:1, nf:4,   align:'right'},
    sumNInt: {f:2, fill:8, bd:1, nf:3,   align:'right'},
    totL:    {f:2, fill:3, bd:1, nf:0,   align:'center'},   // 全院合计行：#d6e0f0 粗体
    totDoc:  {f:2, fill:3, bd:1, nf:0,   align:'left'},     // 表B「全院合计」标签：#d6e0f0 粗体左对齐
    totN:    {f:2, fill:3, bd:1, nf:4,   align:'right'},
    totNInt: {f:2, fill:3, bd:1, nf:3,   align:'right'},
    totPlain:{f:2, fill:3, bd:1, nf:0,   align:'right'},   // 年度合计·目标（navy-100 底，navy 粗）
    totPctDim:{f:2, fill:3, bd:1, nf:0,   align:'right'},  // 年度合计·完成率「/」
    // 月度业绩完成情况（按截图：灰双表头 #D9D9D9 + 钢蓝全院目标 #8497B0）
    grayTitle:{f:2, fill:10, bd:1, nf:0,  align:'left'},    // 灰表头·标题行（左）
    grayHead: {f:2, fill:10, bd:1, nf:0,  align:'center'}, // 灰表头·列名（中）
    chSumL:   {f:2, fill:5,  bd:1, nf:0,  align:'left'},    // 合计行·标签（白底 navy 粗）
    chSumN:   {f:2, fill:5,  bd:1, nf:4,  align:'right'},   // 合计行·金额（白底 navy 粗）
    hospL:    {f:1, fill:9,  bd:1, nf:0,  align:'center'},  // 全院目标·标签（钢蓝底白粗）
    hospPlain:{f:1, fill:9,  bd:1, nf:0,  align:'right'},   // 全院目标·目标（钢蓝底白粗）
    hospN:    {f:1, fill:9,  bd:1, nf:4,  align:'right'},   // 全院目标·实际（钢蓝底白粗）
    hospPct:  {f:1, fill:9,  bd:1, nf:164,align:'right'},   // 全院目标·完成率（钢蓝底白粗）
    hospPctDim:{f:1,fill:9,  bd:1, nf:0,  align:'right'}    // 全院目标·完成率「/」
  };

  function stylesXml(){
    const numFmts = '<numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>';
    const fonts =
      '<fonts count="6">' +
        '<font><sz val="11"/><color rgb="FF0F172A"/><name val="Microsoft YaHei"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FF003366"/><name val="Microsoft YaHei"/></font>' +
        '<font><sz val="11"/><color rgb="FF0F172A"/><name val="Microsoft YaHei"/></font>' +
        '<font><sz val="11"/><color rgb="FF5B6472"/><name val="Microsoft YaHei"/></font>' +
        '<font><i/><sz val="11"/><color rgb="FF5B6472"/><name val="Microsoft YaHei"/></font>' +
      '</fonts>';
    const fills =
      '<fills count="11">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        // v-72 兼容修复：实色填充 bgColor 必须写同一 rgb，否则 WPS 下整行渲染成白底
        '<fill><patternFill patternType="solid"><fgColor rgb="FF002060"/><bgColor rgb="FF002060"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD6E0F0"/><bgColor rgb="FFD6E0F0"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF3F6FB"/><bgColor rgb="FFF3F6FB"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFBFCFE"/><bgColor rgb="FFFBFCFE"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF7F7F7"/><bgColor rgb="FFF7F7F7"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD7EFF9"/><bgColor rgb="FFD7EFF9"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FF8497B0"/><bgColor rgb="FF8497B0"/></patternFill></fill>' +  // fill9 全院目标钢蓝
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/><bgColor rgb="FFD9D9D9"/></patternFill></fill>' +  // fill10 灰表头
      '</fills>';
    const borders =
      '<borders count="2">' +
        '<border><left/><right/><top/><bottom/><diagonal/></border>' +
        '<border>' +
          '<left style="thin"><color rgb="FF000000"/></left>' +
          '<right style="thin"><color rgb="FF000000"/></right>' +
          '<top style="thin"><color rgb="FF000000"/></top>' +
          '<bottom style="thin"><color rgb="FF000000"/></bottom>' +
          '<diagonal/>' +
        '</border>' +
      '</borders>';
    const cellStyleXfs = '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    // v-73 兼容修复：cellXfs[0] 必须是中性默认样式（无填充/无边框），业务样式从索引 1 开始；
    // 否则数据区之外的空白单元格会整片套用第一个业务样式（"空白区全藏蓝"）。
    let cellXfs = '<cellXfs count="' + (Object.keys(STYLE_MAP).length + 1) + '">';
    cellXfs += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    Object.keys(STYLE_MAP).forEach((k)=>{
      const m = STYLE_MAP[k];
      const applyNf = m.nf ? ' applyNumberFormat="1"' : '';
      cellXfs += `<xf numFmtId="${m.nf}" fontId="${m.f}" fillId="${m.fill}" borderId="${m.bd}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"${applyNf}>` +
        `<alignment horizontal="${m.align}" vertical="center"/></xf>`;
    });
    cellXfs += '</cellXfs>';
    const cellStyles = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    const dxfs = '<dxfs count="0"/>';
    const tableStyles = '<tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmts + fonts + fills + borders + cellStyleXfs + cellXfs + cellStyles + dxfs + tableStyles +
      '</styleSheet>';
  }

  function sheetXml(rows, merges, colWidths, freezeRows){
    let cols = '';
    if(colWidths && colWidths.length){
      cols = '<cols>';
      colWidths.forEach((w, i)=>{
        cols += `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;
      });
      cols += '</cols>';
    }

    let maxC = 0;
    rows.forEach(row => { if(row) maxC = Math.max(maxC, row.length); });
    if(merges) merges.forEach(mg => { maxC = Math.max(maxC, mg.c2 + 1); });

    const key = (ri, ci) => ri * 100000 + ci;
    const grid = {};
    rows.forEach((row, ri) => {
      (row || []).forEach((cell, ci) => {
        if(cell !== null && cell !== undefined) grid[key(ri, ci)] = cell;
      });
    });

    // 合并区域补格：把左上角样式铺满整块区域，保证合并块整体有填充+框线
    if(merges && merges.length){
      merges.forEach(mg => {
        const tl = grid[key(mg.r, mg.c)];
        const s = (tl && tl.s) ? tl.s : 'num';
        for(let r = mg.r; r <= mg.r2; r++){
          for(let c = mg.c; c <= mg.c2; c++){
            if(grid[key(r, c)] === undefined){
              grid[key(r, c)] = { v: '', t: 's', s: s };
            }
          }
        }
      });
    }

    let body = '';
    const nHead = freezeRows || 0;
    rows.forEach((row, ri) => {
      const r = ri + 1;
      let attrs = `r="${r}"`;
      if(r <= nHead) attrs += ' ht="26" customHeight="1"';
      else attrs += ' ht="20" customHeight="1"';
      let cells = '';
      for(let ci = 0; ci < maxC; ci++){
        const cell = grid[key(ri, ci)];
        if(cell === undefined) continue;
        const ref = colLetter(ci) + r;
        const sIdx = Object.keys(STYLE_MAP).indexOf(cell.s) + 1;
        const sAttr = ` s="${sIdx < 1 ? 0 : sIdx}"`;
        if(cell.t === 'n'){
          const num = (cell.v === null || cell.v === undefined || isNaN(+cell.v)) ? 0 : +cell.v;
          cells += `<c r="${ref}"${sAttr}><v>${num}</v></c>`;
        } else {
          cells += `<c r="${ref}"${sAttr} t="inlineStr"><is><t xml:space="preserve">${esc(cell.v)}</t></is></c>`;
        }
      }
      body += `<row ${attrs}>${cells}</row>`;
    });

    let mergeXml = '';
    if(merges && merges.length){
      mergeXml = '<mergeCells count="' + merges.length + '">';
      merges.forEach(mg=>{
        const ref = colLetter(mg.c) + (mg.r+1) + ':' + colLetter(mg.c2) + (mg.r2+1);
        mergeXml += `<mergeCell ref="${ref}"/>`;
      });
      mergeXml += '</mergeCells>';
    }

    const pane = nHead > 0
      ? `<pane ySplit="${nHead}" topLeftCell="A${nHead+1}" activePane="bottomLeft" state="frozen"/>`
      : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetViews><sheetView workbookViewId="0">' + pane + '</sheetView></sheetViews>' +
        cols +
        '<sheetData>' + body + '</sheetData>' +
        mergeXml +
      '</worksheet>';
  }

  async function buildStyledXlsx(opts){
    const JSZip = window.JSZip;
    if(!JSZip) throw new Error('JSZip 未加载');
    const sheetName = (opts.sheetName || 'Sheet1').slice(0, 31);
    const contentTypes =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>';
    const rootRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';
    const workbook =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="' + esc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets>' +
      '</workbook>';
    const workbookRels =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>';

    const zip = new JSZip();
    zip.file('[Content_Types].xml', contentTypes);
    zip.folder('_rels').file('.rels', rootRels);
    const xl = zip.folder('xl');
    xl.file('workbook.xml', workbook);
    xl.folder('_rels').file('workbook.xml.rels', workbookRels);
    xl.file('styles.xml', stylesXml());
    xl.folder('worksheets').file('sheet1.xml', sheetXml(opts.rows, opts.merges || [], opts.colWidths || null, opts.freezeRows || 0));
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  window.buildStyledXlsx = buildStyledXlsx;
})();
