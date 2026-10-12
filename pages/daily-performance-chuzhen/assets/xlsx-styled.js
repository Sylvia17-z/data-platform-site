/* xlsx-styled.js
 * 用 JSZip 手写 OOXML，生成「与线上显示一致」的带样式 Excel（藏蓝主题 + 黑框 + 合计加粗）。
 * SheetJS 在本环境无法写入单元格填充/字体样式，故走原生 OOXML。
 * 暴露 window.buildStyledXlsx({ sheetName, rows, merges, colWidths }) -> Promise<Blob>
 *   rows:   二维数组，每个元素为 {v, t:'n'|'s', s:<styleKey>} 或 null（被合并覆盖的格子）
 *   merges: [{r,c,r2,c2}] 0-based 合并区域（仅写左上角，其余置 null）
 *   colWidths: 每列字符宽度（可选）
 * styleKey 映射见 STYLES
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

  // 样式键 -> (fontId, fillId, borderId, numFmtId, wrap)
  // 注意(v-73)：cellXfs[0] 必须是"中性默认样式"（无填充/无边框），在下方 stylesXml() 中单独写入，
  // 因此本表的键从 xf 索引 1 开始：title=1, subhead=2, cat=3, catsub=4, target=5, num=6, pink=7, orange=8。
  // 根因：WPS/Excel 会把「数据区之外的空白单元格」默认套用 cellXfs[0]；
  // 之前 title(藏蓝) 恰好占据索引 0，导致导出文件在数据下方/右侧的空白区域全部渲染成藏蓝（"空白行全藏蓝"）。
  // font0 常规 / font1 标题白粗15 / font2 黑粗11 / font3 合计黑粗14 / font4 列标题黑粗12
  // fill2 藏蓝 / fill3 浅蓝 / fill4 浅蓝2 / fill5 白 / fill6 粉 / fill7 橙
  const STYLE_MAP = {
    title:   {f:1, fill:2, bd:1, nf:0,  wrap:false},               // 标题行：藏蓝底 + 白字加粗15pt + 居中（v-73 由 18pt 改 15pt）
    subhead: {f:4, fill:3, bd:1, nf:0,  wrap:true},                // 列标题行：浅蓝底(#DAE3F5) + 黑字加粗12pt + 居中（还原 v-67 误改的藏蓝）
    cat:     {f:2, fill:3, bd:1, nf:0,  wrap:false},
    catsub:  {f:2, fill:4, bd:1, nf:0,  wrap:false},
    target:  {f:0, fill:5, bd:1, nf:3,  wrap:false}, // 月目标：整数千分位
    num:     {f:0, fill:5, bd:1, nf:4,  wrap:false}, // 金额：2位小数千分位
    pink:    {f:0, fill:6, bd:1, nf:4,  wrap:false},
    orange:  {f:3, fill:7, bd:1, nf:4,  wrap:false}
  };

  function stylesXml(){
    const fonts =
      '<fonts count="5">' +
        '<font><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>' +
        '<font><b/><sz val="15"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font>' +
        '<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>' +
        '<font><b/><sz val="14"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>' +  // font3: 合计行黑粗14（v20260910-77 放大一号：12→14）
        '<font><b/><sz val="12"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>' +  // font4: 列标题行黑粗12（v-71 由白字改黑字）
      '</fonts>';
    const fills =
      '<fills count="8">' +
        '<fill><patternFill patternType="none"/></fill>' +
        '<fill><patternFill patternType="gray125"/></fill>' +
        // 关键兼容修复(v-72)：实色填充必须把 bgColor 也写成同一 rgb。
        // 原 `<bgColor indexed="64"/>` 在 WPS/部分查看器下被当作"系统背景(白)"渲染，
        // 导致藏蓝/浅蓝等填充整行变白——这正是用户"标题行始终是白底"的根因。
        '<fill><patternFill patternType="solid"><fgColor rgb="FF002060"/><bgColor rgb="FF002060"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFDAE3F5"/><bgColor rgb="FFDAE3F5"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F4"/><bgColor rgb="FFD9E1F4"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor rgb="FFFFFFFF"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFC9C7"/><bgColor rgb="FFFFC9C7"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFFFBA84"/><bgColor rgb="FFFFBA84"/></patternFill></fill>' +
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
    // v-73 关键修复：cellXfs[0] 必须是中性默认样式（无填充/无边框/常规字）。
    // WPS/Excel 对数据区之外的空白单元格默认套用 cellXfs[0]；
    // 之前 title(藏蓝) 占据索引 0 → 空白行/空白列区域全被渲染成藏蓝。
    // Excel 自产的文件 cellXfs[0] 恒为空样式，此处对齐该约定；业务样式从索引 1 开始。
    let cellXfs = '<cellXfs count="' + (Object.keys(STYLE_MAP).length + 1) + '">';
    cellXfs += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    Object.keys(STYLE_MAP).forEach((k, i)=>{
      const m = STYLE_MAP[k];
      const applyNf = m.nf ? ' applyNumberFormat="1"' : '';
      const wrap = m.wrap ? '<alignment horizontal="center" vertical="center" wrapText="1"/>' : '<alignment horizontal="center" vertical="center"/>';
      cellXfs += `<xf numFmtId="${m.nf}" fontId="${m.f}" fillId="${m.fill}" borderId="${m.bd}" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"${applyNf}>${wrap}</xf>`;
    });
    cellXfs += '</cellXfs>';
    // v-72 兼容补强：WPS 较严格，缺 <cellStyles> 时会整体忽略 cellXfs（导致所有填充/字体不生效、整表显白）。
    // 补齐 Excel 标准结构（cellStyles/dxfs/tableStyles），与 Excel 导出的 styles.xml 对齐。
    const cellStyles = '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>';
    const dxfs = '<dxfs count="0"/>';
    const tableStyles = '<tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      fonts + fills + borders + cellStyleXfs + cellXfs + cellStyles + dxfs + tableStyles +
      '</styleSheet>';
  }

  function sheetXml(rows, merges, colWidths, hiddenCols){
    let cols = '';
    if(colWidths && colWidths.length){
      const hiddenSet = new Set(hiddenCols || []);
      cols = '<cols>';
      colWidths.forEach((w, i)=>{
        if(hiddenSet.has(i)){
          // 隐藏列：与页面一致，默认隐藏但保留数据，用户可在 Excel 中取消隐藏查看
          cols += `<col min="${i+1}" max="${i+1}" width="0" hidden="1" customWidth="1"/>`;
        }else{
          cols += `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`;
        }
      });
      cols += '</cols>';
    }

    // 计算最大列数（含合并区域覆盖的列）
    let maxC = 0;
    rows.forEach(row => { if(row) maxC = Math.max(maxC, row.length); });
    if(merges) merges.forEach(mg => { maxC = Math.max(maxC, mg.c2 + 1); });

    // 建立 (ri,ci)->cell 网格（仅存入非 null 的真实单元格）
    const key = (ri, ci) => ri * 100000 + ci;
    const grid = {};
    rows.forEach((row, ri) => {
      (row || []).forEach((cell, ci) => {
        if(cell !== null && cell !== undefined) grid[key(ri, ci)] = cell;
      });
    });

    // 合并区域补格：把合并区域「左上角」的样式铺到整块区域（含原本为 null 的空格），
    // 否则合并区域的填充/框线只落在左上角一格 → 标题行/分类行出现"只有首列有底色、框线残缺"。
    if(merges && merges.length){
      merges.forEach(mg => {
        const tl = grid[key(mg.r, mg.c)];
        const s = (tl && tl.s) ? tl.s : 'num'; // 左上角样式键（缺省按金额样式兜底）
        for(let r = mg.r; r <= mg.r2; r++){
          for(let c = mg.c; c <= mg.c2; c++){
            if(grid[key(r, c)] === undefined){
              grid[key(r, c)] = { v: '', t: 's', s: s }; // 补一个同样式空格，保证合并块整体有填充+框线
            }
          }
        }
      });
    }

    let body = '';
    rows.forEach((row, ri) => {
      const r = ri + 1;
      let attrs = `r="${r}"`;
      if(r === 1) attrs += ' ht="32" customHeight="1"';
      else if(r === 2) attrs += ' ht="32" customHeight="1"';
      else attrs += ' ht="22" customHeight="1"'; // 正文行提高一点列高，便于阅读
      let cells = '';
      for(let ci = 0; ci < maxC; ci++){
        const cell = grid[key(ri, ci)];
        if(cell === undefined) continue; // 完全空白（未被任何合并覆盖）→ 不写单元格
        const ref = colLetter(ci) + r;
        // v-73：cellXfs[0] 已被中性默认样式占用，业务样式整体后移一位（title=1, subhead=2, ...）
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

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
        '<sheetViews><sheetView workbookViewId="0">' +
          '<pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/>' +
        '</sheetView></sheetViews>' +
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
    xl.folder('worksheets').file('sheet1.xml', sheetXml(opts.rows, opts.merges || [], opts.colWidths || null, opts.hiddenCols || []));
    return zip.generateAsync({
      type: 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    });
  }

  window.buildStyledXlsx = buildStyledXlsx;
})();
