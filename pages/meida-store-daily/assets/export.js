'use strict';
/* =====================================================================
 * 美大店铺日数据 —— 导出（Excel / 图片）
 * ---------------------------------------------------------------------
 * 需求（2026-10-28）：「该页面需要支持图片以及Excel形式的导出，
 *   并且导出的格式显示需要与当前页面显示完全一致」
 *
 * 一致性策略：导出不另建一套版式，而是**直接按渲染表格时的同一份取值函数
 *   cellValue()/display() 生成导出网格**，颜色/列宽/行高/合并/对齐全部取自
 *   页面 CSS 的同一套色板（FILL 与 index.html :root 的 --xl-* 一一对应）。
 *   因此：切日期 → 导出该日；切「仅当日/仅月累计」→ 导出对应行；
 *   切「仅总表」→ 导出总表口径，屏幕与文件永远同源。
 *
 * Excel：JSZip 手写 OOXML（SheetJS 社区版无法写单元格样式）
 *   · 34 列 × 三层表头（标题 + 分组 + 字段）+ 7 店 × 当日/累计 + 合计行
 *   · 合并：标题整行、分组横向跨列、「统计口径」纵向跨两行
 *   · 底色/字体色/边框色全部照抄页面；冻结窗格 B4（= 页面吸附表头 + 冻结首列）
 *   · 数值保持数字型（金额 #,##0.00 / 整数 #,##0 / 百分比 0.00% / 秒 0.0"s"），
 *     可直接在 Excel 里继续求和，但显示与页面一致
 * 图片：html2canvas 离屏渲染整表（解除 sticky 与滚动容器裁剪，2× 高清）
 * ===================================================================== */
(function () {

  /* ---------------- 依赖懒加载 ---------------- */
  /* 2026-10-30：线上偶发「导出图片失败：依赖加载失败：assets/vendor/html2canvas.min.js」。
   * 根因是相对路径 + 一次性加载：外壳 iframe 的 base 异常、网关瞬时抖动（首页曾出现 3.3s 才返回）、
   * Service Worker 把网络失败包装成 503，都会让 <script> 触发 onerror；而失败结果又被 _lib 缓存，
   * 于是首次失败后整个会话都失败，必须刷新页面才能恢复。
   * 处理：
   *   ① 路径改为以「本脚本自身 URL」为基准解析，不再依赖文档 base 与相对路径；
   *   ② 本地失败自动换 CDN（与 vendor 同版本）兜底；
   *   ③ 每个候选地址重试一次并加时间戳，绕开被污染的缓存；
   *   ④ 只缓存成功结果，失败不落缓存，用户再点一次即可重新加载。 */
  const SELF = (document.currentScript && document.currentScript.src) || '';
  const ASSET_BASE = SELF ? SELF.replace(/[?#].*$/, '').replace(/[^/]*$/, '') : 'assets/';
  const SRC = {
    jszip: [
      ASSET_BASE + 'vendor/jszip.min.js',
      'assets/vendor/jszip.min.js',
      'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
    ],
    h2c: [
      ASSET_BASE + 'vendor/html2canvas.min.js',
      'assets/vendor/html2canvas.min.js',
      'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
    ],
  };
  const LIB_NAME = { jszip: 'JSZip', h2c: 'html2canvas' };
  const _lib = {};
  function inject(url) {
    return new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = url;
      s.async = true;
      s.onload = () => resolve(true);
      s.onerror = () => reject(new Error(url));
      document.head.appendChild(s);
    });
  }
  function loadScript(key) {
    if (_lib[key]) return _lib[key];
    const urls = SRC[key] || [key];
    const p = (async () => {
      for (const u of urls) {
        for (let i = 0; i < 2; i++) {
          const url = i === 0 ? u : u + (u.indexOf('?') >= 0 ? '&' : '?') + 'retry=' + Date.now();
          try { await inject(url); return true; } catch (_) { /* 换下一个候选 / 重试 */ }
        }
      }
      throw new Error('依赖加载失败：' + (LIB_NAME[key] || key) + '（已重试并尝试 CDN，请检查网络后重试）');
    })();
    p.catch(() => { delete _lib[key]; }); // 失败不缓存
    _lib[key] = p;
    return p;
  }

  /* ---------------- 色板（与 index.html :root 严格一致） ---------------- */
  const FILL = {
    title: '#FFFFFF',
    g1: '#1F497D', g2: '#17375E', g3: '#4472C4', g4: '#31859B',
    g5: '#60497A', g6: '#76923C', g7: '#C55A11',
    rowA: '#B8CCE4',      // 当日 / 累计行 B–N
    rowTodayB: '#D9E2F3', // 当日在 O–AB
    rowCumB: '#E2EFDA',   // 累计在 O–AB
    orange: '#FCE4D6',    // 投产比 + 好评区（AC–AH）
    sum: '#FFF2CC',       // 合计行
    period: '#EEF2FA',    // 统计口径列
  };
  const TEXT = {
    body: '#1F2D4D', head: '#FFFFFF', headDark: '#000000',
    navy: '#002060', sum: '#7A4A00', dash: '#9AA9C0', fallback: '#8A6A2A',
  };
  const BORDER = { thin: '#D6DDE8', title: '#9AA9C0' };

  // 数值格式：0=General 3=#,##0 4=#,##0.00 176=0.00% 177=0.0"s" 178=0.0
  const NF = { gen: 0, int: 3, money: 4, pct: 176, sec: 177, dec1: 178 };
  function nfOf(col) {
    switch (col.type) {
      case 'money': return NF.money;
      case 'money0': return NF.int;
      case 'int': return NF.int;
      case 'pct': return NF.pct;
      case 'sec': return NF.sec;
      case 'score': return NF.dec1;
      default: return NF.gen;
    }
  }
  const isTextCol = (col) => col.type === 'text' || col.type === 'scoreText';

  /* ---------------- 行高量测：取页面实际渲染高度，保证与屏幕同高 ---------------- */
  function measureHeights() {
    const t = document.getElementById('tbl');
    const out = { head: [30, 34, 50], body: 25 };
    try {
      if (t && t.tHead && t.tHead.rows.length === 3) {
        out.head = [...t.tHead.rows].map((r) => r.offsetHeight || 30);
      }
      const tr = t && t.tBodies[0] && t.tBodies[0].rows[0];
      if (tr && tr.offsetHeight) out.body = tr.offsetHeight;
    } catch (e) { /* 表格尚未渲染 → 用默认值 */ }
    return out;
  }

  /* =====================================================================
   * 一、导出网格（与页面 DOM 同源）
   * ===================================================================== */
  function dashCell(spec) {
    return { v: '—', t: 's', spec: Object.assign({}, spec, { color: TEXT.dash }) };
  }
  // 单个数据格：优先取 cellValue()，与 rowHTML() 完全同源
  function valueCell(col, storeKey, which, spec) {
    const { v, kind } = cellValue(storeKey, col.k, which);
    const base = Object.assign({}, spec);
    if (kind === 'sheet') { base.color = TEXT.fallback; base.italic = true; }
    if (v === null || v === undefined || v === '') return dashCell(base);
    if (typeof v === 'number' && isFinite(v)) {
      return { v: v, t: 'n', spec: Object.assign(base, { nf: nfOf(col), align: isTextCol(col) ? 'center' : 'right' }) };
    }
    return { v: String(v), t: 's', spec: Object.assign(base, { align: isTextCol(col) ? 'center' : 'right' }) };
  }
  // 合计行单格：与 sumRowHTML() 同源
  function sumCell(col, which, spec) {
    if (col.k === 'period') return { v: '—', t: 's', spec: Object.assign({}, spec, { color: TEXT.dash }) };
    if (isTextCol(col) || SUM_KEYS.indexOf(col.k) < 0) return dashCell(spec);
    let total = null;
    STORES.forEach((s) => {
      const { v } = cellValue(s.key, col.k, which);
      if (typeof v === 'number' && isFinite(v)) total = (total || 0) + v;
    });
    if (total === null) return dashCell(spec);
    return { v: total, t: 'n', spec: Object.assign({}, spec, { nf: nfOf(col), align: 'right' }) };
  }

  function buildGrid() {
    const period = state.period;
    const nCol = COLS.length;
    const H = measureHeights();
    const rows = [];
    const merges = [];
    const heights = [];

    /* —— 第 1 行：标题（白底 藏蓝字 加粗，下边线 #9AA9C0）—— */
    const r0 = new Array(nCol).fill(null);
    r0[0] = { v: '海蓓医疗七大店铺团购数据总表', t: 's', spec: { fill: FILL.title, color: TEXT.body, bold: true, size: 10.5, align: 'center', bd: 'bottom' } };
    rows.push(r0); heights.push(H.head[0]);
    merges.push({ r: 0, c: 0, r2: 0, c2: nCol - 1 });

    /* —— 第 2 行：分组表头（「统计口径」纵向跨两行）—— */
    const r1 = new Array(nCol).fill(null);
    GROUPS.forEach((g) => {
      const idxs = [];
      COLS.forEach((c, i) => { if (c.g === g.id) idxs.push(i); });
      if (!idxs.length) return;
      const start = idxs[0];
      const end = idxs[idxs.length - 1];
      const dark = g.id === 4;   // 「业绩」组：绿松石底 + 黑字（页面 .g4{color:#000!important}）
      r1[start] = {
        v: g.label, t: 's',
        spec: { fill: FILL[g.cls] || FILL.g3, color: dark ? TEXT.headDark : TEXT.head, bold: true, size: 9, align: 'center', wrap: true },
      };
      if (end > start || g.rowspan) merges.push({ r: 1, c: start, r2: g.rowspan ? 2 : 1, c2: end });
    });
    rows.push(r1); heights.push(H.head[1]);

    /* —— 第 3 行：字段名（#4472C4 白字加粗，高 50px 自动换行）—— */
    const r2 = new Array(nCol).fill(null);
    COLS.forEach((c, i) => {
      if (GROUPS.some((g) => g.id === c.g && g.rowspan)) return;   // 纵向合并列不写
      r2[i] = {
        v: c.label, t: 's',
        spec: { fill: FILL.g3, color: TEXT.head, bold: true, size: 8, align: c.k === 'store' ? 'left' : 'center', wrap: true },
      };
    });
    rows.push(r2); heights.push(H.head[2]);

    /* —— 数据行：7 店 × 当日 / 本月累计（合并列与页面同源：共享列 + 投产比/好评 5 列按值判定） —— */
    function buildDataRow(store, which, merge, ms) {
      const isToday = which === 'today';
      return COLS.map((c) => {
        if (merge && ms.has(c.k) && !isToday) return null;   // 第 2 行合并列由纵向合并覆盖
        // 合并列：当日为空而累计有值时取累计值（与页面同源，否则合并后这个值会消失）
        let src = which;
        if (merge && ms.has(c.k) && MERGED_FIELDS.indexOf(c.k) < 0
            && !rawTextOf(store.key, c.k, 'today') && rawTextOf(store.key, c.k, 'cum')) src = 'cum';
        if (c.k === 'store') {
          return { v: storeName(store), t: 's', spec: { fill: FILL.rowA, color: TEXT.navy, bold: true, size: 9, align: 'left', wrap: true } };
        }
        if (c.k === 'period') {
          return { v: isToday ? '当日数据' : '本月累计', t: 's', spec: { fill: FILL.period, color: TEXT.navy, bold: true, size: 9, align: 'center' } };
        }
        // 行底色：B–N 统一，O–AB 按当日/累计分色，AC–AH 统一橙色
        const fill = c.zone === 'b' ? (isToday ? FILL.rowTodayB : FILL.rowCumB)
          : c.zone === 'c' ? FILL.orange : FILL.rowA;
        const spec = { fill: fill, color: TEXT.body, size: 9 };
        if (c.k === 'pm' || c.k === 'op') {
          // 与页面同源：先取手工覆盖，再回落源表值 / 默认责任人（否则导出的责任人与屏幕不一致）
          const cv = cellValue(store.key, c.k, src).v;
          const v = (cv === null || cv === undefined || cv === '') ? (store[c.k] || '') : cv;
          return v === '' ? dashCell(Object.assign(spec, { align: 'center' }))
            : { v: String(v), t: 's', spec: Object.assign(spec, { align: 'center' }) };
        }
        return valueCell(c, store.key, src, Object.assign(spec, { align: isTextCol(c) ? 'center' : 'right' }));
      });
    }
    const dataRows = [];
    const dataMerges = [];
    let dr = rows.length;   // 数据区起始行（0-indexed）
    STORES.forEach((s) => {
      if (period === 'both') {
        const ms = mergeSetFor(s);
        const r1 = buildDataRow(s, 'today', true, ms);
        const r2 = buildDataRow(s, 'cum', true, ms);
        dataRows.push(r1, r2);
        COLS.forEach((c, ci) => { if (ms.has(c.k)) dataMerges.push({ r: dr, c: ci, r2: dr + 1, c2: ci }); });
        dr += 2;
      } else {
        const w = period === 'cum' ? 'cum' : 'today';
        dataRows.push(buildDataRow(s, w, false, mergeSetFor(s)));
        dr += 1;
      }
    });
    dataRows.forEach((r) => { rows.push(r); heights.push(H.body); });
    dataMerges.forEach((m) => merges.push(m));

    /* —— 合计行 —— */
    const sumWhich = period === 'cum' ? 'cum' : 'today';
    const sumLabel = period === 'cum' ? '合计（本月累计）' : '合计（当日数据）';
    rows.push(COLS.map((c) => {
      const fill = c.zone === 'c' ? FILL.orange : (c.k === 'period' ? FILL.period : FILL.sum);
      const spec = { fill: fill, color: c.k === 'period' ? TEXT.navy : TEXT.sum, bold: true, size: 9 };
      if (c.k === 'store') return { v: sumLabel, t: 's', spec: Object.assign(spec, { align: 'left' }) };
      return sumCell(c, sumWhich, spec);
    }));
    heights.push(H.body);

    return {
      rows: rows,
      merges: merges,
      colWidths: COLS.map((c) => Math.max(5, Math.round(((c.w - 5) / 7) * 10) / 10)),
      rowHeights: heights.map((px) => Math.round(px * 0.75 * 10) / 10),
      freeze: { xSplit: 1, ySplit: 3 },
      meta: { date: state.date, period: period, source: state.source, cols: nCol, rows: rows.length },
    };
  }

  /* =====================================================================
   * 二、OOXML：JSZip 手写（样式全部来自上面的网格 spec）
   * ===================================================================== */
  function xmlEsc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }
  function colLetter(c) {
    let s = ''; c += 1;
    while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
    return s;
  }
  const rgb = (hex) => 'FF' + String(hex || '#000000').replace('#', '').toUpperCase();

  // 样式簿：字体 / 填充 / 边框 / cellXfs 按 spec 去重生成
  function makeStyleBook() {
    const fonts = [{ b: 0, i: 0, sz: 9, color: TEXT.body }];      // 0 = 中性默认
    const fMap = new Map([[key0(fonts[0]), 0]]);
    const fills = ['none', 'gray125'];                            // 0/1 必须占位
    const flMap = new Map();
    const xfs = [null];                                           // 0 = 中性默认
    const xMap = new Map();
    function key0(f) { return [f.b, f.i, f.sz, rgb(f.color)].join('|'); }
    function fontId(sp) {
      const f = { b: sp.bold ? 1 : 0, i: sp.italic ? 1 : 0, sz: sp.size || 9, color: sp.color || TEXT.body };
      const k = key0(f);
      if (fMap.has(k)) return fMap.get(k);
      fonts.push(f); fMap.set(k, fonts.length - 1); return fonts.length - 1;
    }
    function fillId(sp) {
      if (!sp.fill) return 0;
      const k = rgb(sp.fill);
      if (flMap.has(k)) return flMap.get(k);
      fills.push(k); flMap.set(k, fills.length - 1); return fills.length - 1;
    }
    function xfId(sp) {
      const o = {
        f: fontId(sp), fill: fillId(sp),
        bd: sp.bd === false ? 0 : (sp.bd === 'bottom' ? 2 : 1),
        nf: sp.nf || 0, align: sp.align || 'center', wrap: sp.wrap ? 1 : 0,
      };
      const k = JSON.stringify(o);
      if (xMap.has(k)) return xMap.get(k);
      xfs.push(o); xMap.set(k, xfs.length - 1); return xfs.length - 1;
    }
    return { fonts, fills, xfs, xfId };
  }

  function stylesXml(sb) {
    const numFmts = '<numFmts count="3">' +
      '<numFmt numFmtId="176" formatCode="0.00%"/>' +
      '<numFmt numFmtId="177" formatCode="0.0&quot;s&quot;"/>' +
      '<numFmt numFmtId="178" formatCode="0.0"/>' +
      '</numFmts>';
    const fonts = '<fonts count="' + sb.fonts.length + '">' + sb.fonts.map((f) =>
      '<font>' + (f.b ? '<b/>' : '') + (f.i ? '<i/>' : '') +
      '<sz val="' + f.sz + '"/><color rgb="' + rgb(f.color) + '"/><name val="Microsoft YaHei"/><charset val="134"/></font>').join('') + '</fonts>';
    // v-72 兼容修复：实色填充 fgColor/bgColor 必须写同一 rgb，否则 WPS 下渲染为白底
    const fills = '<fills count="' + sb.fills.length + '">' + sb.fills.map((c) => {
      if (c === 'none') return '<fill><patternFill patternType="none"/></fill>';
      if (c === 'gray125') return '<fill><patternFill patternType="gray125"/></fill>';
      return '<fill><patternFill patternType="solid"><fgColor rgb="' + c + '"/><bgColor rgb="' + c + '"/></patternFill></fill>';
    }).join('') + '</fills>';
    const side = (name, style, color) => style
      ? '<' + name + ' style="' + style + '"><color rgb="' + color + '"/></' + name + '>'
      : '<' + name + '/>';
    const borders = '<borders count="3">' +
      '<border><left/><right/><top/><bottom/><diagonal/></border>' +
      '<border>' + side('left', 'thin', rgb(BORDER.thin)) + side('right', 'thin', rgb(BORDER.thin)) +
      side('top', 'thin', rgb(BORDER.thin)) + side('bottom', 'thin', rgb(BORDER.thin)) + '<diagonal/></border>' +
      '<border>' + side('left', 'thin', rgb(BORDER.thin)) + side('right', 'thin', rgb(BORDER.thin)) +
      side('top', 'thin', rgb(BORDER.thin)) + side('bottom', 'medium', rgb(BORDER.title)) + '<diagonal/></border>' +
      '</borders>';
    const cellStyleXfs = '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>';
    // v-73 兼容修复：cellXfs[0] 必须中性（无填充/无边框），业务样式从 1 开始，
    // 否则数据区之外的空白单元格会整片套用第一个业务样式
    let cellXfs = '<cellXfs count="' + sb.xfs.length + '">';
    cellXfs += '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>';
    sb.xfs.slice(1).forEach((o) => {
      const applyNf = o.nf ? ' applyNumberFormat="1"' : '';
      cellXfs += '<xf numFmtId="' + o.nf + '" fontId="' + o.f + '" fillId="' + o.fill + '" borderId="' + o.bd + '" xfId="0"' +
        ' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"' + applyNf + '>' +
        '<alignment horizontal="' + o.align + '" vertical="center"' + (o.wrap ? ' wrapText="1"' : '') + '/></xf>';
    });
    cellXfs += '</cellXfs>';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      numFmts + fonts + fills + borders + cellStyleXfs + cellXfs +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2"/>' +
      '</styleSheet>';
  }

  function sheetXml(g, sb) {
    const rows = g.rows, merges = g.merges || [];
    const colWidths = g.colWidths || [];
    const rowHeights = g.rowHeights || [];

    let cols = '';
    if (colWidths.length) {
      cols = '<cols>' + colWidths.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>';
    }

    let maxC = 0;
    rows.forEach((row) => { if (row) maxC = Math.max(maxC, row.length); });
    merges.forEach((m) => { maxC = Math.max(maxC, m.c2 + 1); });

    const key = (r, c) => r * 100000 + c;
    const grid = {};
    rows.forEach((row, ri) => {
      (row || []).forEach((cell, ci) => { if (cell) grid[key(ri, ci)] = cell; });
    });
    // 合并区域补格：把左上角样式铺满整块，保证合并块整体有底色 + 框线
    merges.forEach((m) => {
      const tl = grid[key(m.r, m.c)];
      const spec = (tl && tl.spec) ? tl.spec : null;
      if (!spec) return;
      for (let r = m.r; r <= m.r2; r++) {
        for (let c = m.c; c <= m.c2; c++) {
          if (!grid[key(r, c)]) grid[key(r, c)] = { v: '', t: 's', spec: spec };
        }
      }
    });

    let body = '';
    rows.forEach((row, ri) => {
      const r = ri + 1;
      const h = rowHeights[ri];
      let attrs = 'r="' + r + '"' + (h ? ' ht="' + h + '" customHeight="1"' : '');
      let cells = '';
      for (let ci = 0; ci < maxC; ci++) {
        const cell = grid[key(ri, ci)];
        if (!cell) continue;
        const ref = colLetter(ci) + r;
        const s = sb.xfId(cell.spec || {});
        if (cell.t === 'n') {
          const n = (cell.v === null || cell.v === undefined || isNaN(+cell.v)) ? 0 : +cell.v;
          cells += '<c r="' + ref + '" s="' + s + '"><v>' + n + '</v></c>';
        } else {
          const txt = cell.v == null ? '' : String(cell.v);
          cells += '<c r="' + ref + '" s="' + s + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(txt) + '</t></is></c>';
        }
      }
      body += '<row ' + attrs + '>' + cells + '</row>';
    });

    let mergeXml = '';
    if (merges.length) {
      mergeXml = '<mergeCells count="' + merges.length + '">' + merges.map((m) =>
        '<mergeCell ref="' + colLetter(m.c) + (m.r + 1) + ':' + colLetter(m.c2) + (m.r2 + 1) + '"/>').join('') + '</mergeCells>';
    }
    const fz = g.freeze || {};
    const pane = (fz.ySplit || fz.xSplit)
      ? '<pane' + (fz.xSplit ? ' xSplit="' + fz.xSplit + '"' : '') + (fz.ySplit ? ' ySplit="' + fz.ySplit + '"' : '') +
        ' topLeftCell="' + colLetter(fz.xSplit || 0) + ((fz.ySplit || 0) + 1) + '" activePane="bottomRight" state="frozen"/>'
      : '';

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetViews><sheetView workbookViewId="0">' + pane + '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="18"/>' +
      cols +
      '<sheetData>' + body + '</sheetData>' +
      mergeXml +
      '</worksheet>';
  }

  async function buildXlsxBlob(g) {
    await loadScript('jszip');
    const JSZip = window.JSZip;
    if (!JSZip) throw new Error('JSZip 未加载');
    const sb = makeStyleBook();
    const sheetName = ('美大店铺日数据 ' + (g.meta.date || '')).slice(0, 31);
    const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '</Types>';
    const rootRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>';
    const workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="' + xmlEsc(sheetName) + '" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const workbookRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
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
    const sheet = sheetXml(g, sb);
    xl.file('styles.xml', stylesXml(sb));
    xl.folder('worksheets').file('sheet1.xml', sheet);
    return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* =====================================================================
   * 三、图片：html2canvas 离屏渲染整表
   * ===================================================================== */
  function totalWidth() {
    return COLS.reduce((a, c) => a + c.w, 0);
  }
  async function buildPngCanvas() {
    await loadScript('h2c');
    if (typeof window.html2canvas !== 'function') throw new Error('html2canvas 未加载');
    const table = document.getElementById('tbl');
    if (!table || table.style.display === 'none') throw new Error('暂无可导出的表格');

    const W = totalWidth();
    const holder = document.createElement('div');
    holder.style.cssText = 'position:absolute;left:-30000px;top:0;background:#fff;z-index:-1;';
    const clone = table.cloneNode(true);
    clone.removeAttribute('id');
    clone.style.width = W + 'px';
    clone.style.minWidth = W + 'px';
    clone.style.margin = '0';
    holder.appendChild(clone);
    document.body.appendChild(holder);

    try {
      // 必须先挂到文档再读计算样式：脱离文档的节点 getComputedStyle 返回默认值，
      // 会把页面里的 position:sticky（吸附表头 / 冻结首列）误判成 static 而漏改，
      // 导致 html2canvas 把首列与表头绘制到错误位置（2026-10-28 实测：首列整列空白）。
      clone.querySelectorAll('th,td').forEach((el) => {
        if (getComputedStyle(el).position === 'sticky') {
          el.style.position = 'static';
          el.style.left = 'auto';
          el.style.top = 'auto';
          el.style.zIndex = 'auto';
        }
        el.removeAttribute('data-edit');   // 导出图不带交互属性
      });
      const canvas = await window.html2canvas(clone, {
        backgroundColor: '#ffffff',
        scale: 2,
        useCORS: true,
        logging: false,
        width: W,
        height: clone.scrollHeight || clone.offsetHeight,
        windowWidth: Math.max(W + 100, window.innerWidth),
      });
      return canvas;
    } finally {
      holder.remove();
    }
  }

  /* =====================================================================
   * 四、下载与按钮
   * ===================================================================== */
  function pad2(n) { return String(n).padStart(2, '0'); }
  function fileBase() {
    const d = state.date || '未选日期';
    const p = state.period === 'today' ? '_仅当日' : state.period === 'cum' ? '_仅月累计' : '';
    return '美大店铺日数据_' + d + p;
  }
  function downloadBlob(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(url); a.remove(); }, 2000);
  }
  function setBusy(on, msg) {
    const meta = document.getElementById('export-meta');
    if (meta) {
      meta.textContent = msg || (on ? '正在导出…' : '导出内容 = 当前屏幕上这份报表（同日期 / 同口径 / 同列）');
      meta.className = 'file-meta' + (on ? '' : '');
    }
    ['btn-export-xlsx', 'btn-export-png'].forEach((id) => {
      const b = document.getElementById(id);
      if (b) b.disabled = !!on;
    });
  }

  async function exportXlsx() {
    if (!state.date) { toast('暂无可导出的数据，请先导入订单表 / 店铺总表 / 平台数据', 'warn'); return; }
    setBusy(true, '正在生成 Excel…');
    try {
      const g = buildGrid();
      const blob = await buildXlsxBlob(g);
      downloadBlob(blob, fileBase() + '.xlsx');
      toast('Excel 已导出（' + g.meta.rows + ' 行 × ' + g.meta.cols + ' 列，版式与页面一致）', 'ok', 3600);
      setBusy(false, '已导出 Excel：' + fileBase() + '.xlsx');
    } catch (e) {
      console.error(e);
      toast('Excel 导出失败：' + e.message, 'err', 6000);
      setBusy(false, 'Excel 导出失败');
    }
  }

  async function exportPng() {
    if (!state.date) { toast('暂无可导出的数据，请先导入订单表 / 店铺总表 / 平台数据', 'warn'); return; }
    setBusy(true, '正在生成图片…');
    try {
      const canvas = await buildPngCanvas();
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/png'));
      if (!blob) throw new Error('画布导出失败');
      downloadBlob(blob, fileBase() + '.png');
      toast('图片已导出（' + canvas.width + '×' + canvas.height + 'px，2 倍高清）', 'ok', 3600);
      setBusy(false, '已导出图片：' + fileBase() + '.png');
    } catch (e) {
      console.error(e);
      toast('图片导出失败：' + e.message, 'err', 6000);
      setBusy(false, '图片导出失败');
    }
  }

  function bind() {
    const bx = document.getElementById('btn-export-xlsx');
    const bp = document.getElementById('btn-export-png');
    if (bx) bx.addEventListener('click', exportXlsx);
    if (bp) bp.addEventListener('click', exportPng);
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();

  // 供回归测试调用（不参与页面交互）
  window.MEIDA_EXPORT = {
    buildGrid: buildGrid,
    xlsxBlob: async () => buildXlsxBlob(buildGrid()),
    pngCanvas: buildPngCanvas,
    exportXlsx: exportXlsx,
    exportPng: exportPng,
    FILL: FILL,
  };
})();
