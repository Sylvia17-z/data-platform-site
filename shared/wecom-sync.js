/**
 * 页面侧「同步到企微在线表格」出口
 *
 * 只做三件事：把当前屏幕的网格序列化成 TSV、下载成一个约定命名的文件、
 * 并在右键时提供「复制为表格数据」的手动兜底。
 *
 * 真正的写入由本机 platform/tools/wecom-sync/push.mjs 完成 ——
 * 企微 API 没有跨域头，浏览器无法直连，也拿不到 access_token。
 *
 * 必须是 classic script（不能用 ESM）：美大 export.js 依赖与 app.js
 * 共享同一个全局词法作用域去访问 state / COLS / cellValue。
 *
 * 本文件里的序列化逻辑与 tools/wecom-sync/lib/tsv.mjs 逐字等价，
 * 由 selftest.mjs 的「跨端一致」用例守住。
 */
(function (global) {
  'use strict';
  if (global.HBWecomSync) return;

  var MAGIC = '#HBWECOM';
  var SCHEMA_VER = 1;

  /* ---------- 单元格地址 ---------- */

  function colLetter(i) {
    var n = Number(i);
    if (!isFinite(n) || n < 0) n = 0;
    var s = '';
    do {
      s = String.fromCharCode(65 + (n % 26)) + s;
      n = Math.floor(n / 26) - 1;
    } while (n >= 0);
    return s;
  }

  function rangeOf(rows, cols, anchor) {
    var a = anchor || 'A1';
    var r = Math.max(1, Math.floor(rows) || 1);
    var c = Math.max(1, Math.floor(cols) || 1);
    if (r === 1 && c === 1) return a;
    return a + ':' + colLetter(c - 1) + r;
  }

  /* ---------- 数字格式 ---------- */

  function killNegZero(s) {
    return /^-0\.?0*$/.test(s) ? s.slice(1) : s;
  }

  function formatNumber(v, nf) {
    if (typeof v !== 'number' || !isFinite(v)) return '';
    var x = v === 0 ? 0 : v;
    switch (nf) {
      case 176: return killNegZero((x * 100).toFixed(2)) + '%';
      case 177: return killNegZero(x.toFixed(1));
      case 178: return killNegZero(x.toFixed(1));
      case 4:   return killNegZero(x.toFixed(2));
      case 2:   return killNegZero(x.toFixed(2));
      case 3:
      case 1:   return String(Math.round(x));
      default:  return String(x);
    }
  }

  /* ---------- 文本清洗 ---------- */

  function sanitizeText(input, opts) {
    var o = opts || {};
    var keepNewlines = !!o.keepNewlines;
    var maxLen = o.maxCellLen || 5000;
    var guard = o.guardFormula !== false;
    var s = String(input == null ? '' : input);
    s = s.replace(/\u00A0/g, ' ')
         .replace(/[\u200B-\u200D\uFEFF]/g, '')
         .replace(/\t/g, ' ')
         .replace(keepNewlines ? /\r\n?/g : /[\r\n]+/g, keepNewlines ? '\n' : ' ');
    if (s.length > maxLen) s = s.slice(0, maxLen);
    if (guard && /^[=+@]/.test(s)) s = "'" + s;
    return s;
  }

  /* ---------- 单元格取值 ---------- */

  function cellToText(cell, opts) {
    var o = opts || {};
    if (cell == null) return '';
    if (typeof cell !== 'object') {
      return typeof cell === 'number' ? formatNumber(cell) : sanitizeText(cell, o);
    }
    var v = cell.v;
    if (v == null) return '';
    if (typeof v === 'number') {
      if (!isFinite(v)) return '';
      var nf = (cell.spec && cell.spec.nf != null) ? cell.spec.nf : cell.nf;
      return formatNumber(v, nf);
    }
    var s = String(v);
    if (s === '\u2014' || s === '\uFF0D') return o.dash ? String(o.dash) : '';
    return sanitizeText(s, o);
  }

  /* ---------- 网格展开 ---------- */

  function gridCols(grid) {
    var n = 0;
    var rows = (grid && grid.rows) || [];
    for (var i = 0; i < rows.length; i++) if (rows[i] && rows[i].length > n) n = rows[i].length;
    var merges = (grid && grid.merges) || [];
    for (var j = 0; j < merges.length; j++) {
      var m = merges[j];
      if (m && m.c2 + 1 > n) n = m.c2 + 1;
    }
    return n;
  }

  function expandGrid(grid, opts) {
    var o = Object.assign({ expandMerges: 'all', dash: '' }, opts || {});
    var rows = (grid && grid.rows) || [];
    var merges = (grid && grid.merges) || [];
    var nCol = gridCols(grid);
    var out = rows.map(function (r) {
      var arr = new Array(nCol).fill('');
      for (var c = 0; c < nCol; c++) arr[c] = cellToText(r ? r[c] : null, o);
      return arr;
    });
    if (o.expandMerges === 'none') return out;
    for (var k = 0; k < merges.length; k++) {
      var m = merges[k];
      if (!m) continue;
      var horz = m.c2 > m.c;
      if (o.expandMerges === 'vertical' && horz) continue;
      var val = (out[m.r] && out[m.r][m.c] != null) ? out[m.r][m.c] : '';
      if (val === '') continue;
      for (var r = m.r; r <= m.r2 && r < out.length; r++) {
        for (var c = m.c; c <= m.c2 && c < nCol; c++) out[r][c] = val;
      }
    }
    return out;
  }

  function gridToTsv(grid, opts) {
    return expandGrid(grid, opts).map(function (r) { return r.join('\t'); }).join('\r\n');
  }

  /* ---------- 表头断言 ---------- */

  function assertHeaderShape(grid, headerRows) {
    var rows = (grid && grid.rows) || [];
    var merges = (grid && grid.merges) || [];
    var nCol = gridCols(grid);
    if (!rows.length) throw new Error('表头校验失败：网格为空');
    if (nCol < 2) throw new Error('表头校验失败：列数异常（' + nCol + '）');
    if (headerRows < 1 || headerRows > rows.length) {
      throw new Error('表头校验失败：headerRows=' + headerRows + ' 超出实际行数 ' + rows.length);
    }
    var hasTitle = merges.some(function (m) {
      return m && m.r === 0 && m.c === 0 && m.r2 === 0 && m.c2 === nCol - 1;
    });
    if (!hasTitle) {
      throw new Error('表头结构变化：第 1 行不再是「A1:' + colLetter(nCol - 1) + '1」整行标题合并');
    }
    if (headerRows === 3) {
      var groups = merges.filter(function (m) { return m && m.r === 1 && m.c2 > m.c; }).length;
      if (groups < 2) throw new Error('表头结构变化：第 2 行分组合并只剩 ' + groups + ' 块，应为多层表头');
    } else if (headerRows === 2) {
      var ok = merges.some(function (m) { return m && m.r === 1 && m.c === 0 && m.c2 === 2; });
      if (!ok) throw new Error('表头结构变化：第 2 行「部门类别」不再是三列合并');
    }
    return { headerRows: headerRows, dataStartRow: headerRows, cols: nCol, rows: rows.length };
  }

  /* ---------- 文件名与 magic 头 ---------- */

  function pad(n) { return n < 10 ? '0' + n : String(n); }

  function stampOf(d) {
    d = d || new Date();
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) +
      '-' + pad(d.getHours()) + pad(d.getMinutes()) + pad(d.getSeconds());
  }

  function nonceOf() {
    return Math.random().toString(36).slice(2, 6) + Math.random().toString(36).slice(2, 4);
  }

  function buildFileName(module, stamp, nonce) {
    return 'wecom-sync__' + module + '__' + stamp + '__' + nonce + '.tsv';
  }

  /* ---------- 提示 ---------- */

  function notify(msg, type) {
    try {
      if (typeof global.toast === 'function') return global.toast(msg, type || 'ok');
      if (typeof global.showToast === 'function') return global.showToast(msg, type || 'ok');
    } catch (e) { /* 各页 toast 实现不同，失败不阻断 */ }
    if (global.console) console.log('[HBWecomSync] ' + msg);
  }

  /* ---------- 下载 ---------- */

  function download(text, filename) {
    var blob = new Blob(['\uFEFF' + text], { type: 'text/tab-separated-values;charset=utf-8' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 2000);
  }

  /* ---------- 对外接口 ---------- */

  /**
   * @param {object} grid  { rows, merges, ... }
   * @param {object} meta  { module, date, period, label, headerRows, expandMerges }
   */
  function push(grid, meta) {
    meta = meta || {};
    var module = meta.module || 'unknown';
    var label = meta.label || module;
    try {
      if (!grid || !grid.rows || !grid.rows.length) throw new Error('没有可同步的数据');
      if (meta.headerRows) assertHeaderShape(grid, meta.headerRows);

      var opts = { expandMerges: meta.expandMerges || 'all' };
      var cols = gridCols(grid);
      var rows = grid.rows.length;
      var tsv = gridToTsv(grid, opts);
      var stamp = stampOf();
      var nonce = nonceOf();

      var head = [
        MAGIC, SCHEMA_VER, module,
        meta.date || '', meta.period || '',
        cols, rows, stamp, nonce,
      ].join('\t');

      // magic 头与正文必须用 CRLF 分隔：正文本身也是 CRLF，混用会让第一行与表头粘连
      download(head + '\r\n' + tsv, buildFileName(module, stamp, nonce));

      try { localStorage.setItem('hb_wecom_last_submit_' + module, String(Date.now())); } catch (e) {}

      notify('已生成同步文件（' + cols + ' 列 × ' + rows + ' 行）→ 双击「同步到企微」完成写入「' + label + '」', 'ok');
      return { ok: true, cols: cols, rows: rows, stamp: stamp, nonce: nonce };
    } catch (e) {
      notify('生成同步文件失败：' + (e && e.message ? e.message : e), 'err');
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }

  /** 兜底：不依赖本机脚本，直接复制成可在企微表格 Ctrl+V 的文本 */
  function copyToClipboard(grid, meta) {
    meta = meta || {};
    try {
      var tsv = gridToTsv(grid, { expandMerges: meta.expandMerges || 'all' });
      if (global.HBCopy && typeof global.HBCopy.copy === 'function') {
        global.HBCopy.copy(tsv);
      } else if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(tsv);
      } else {
        throw new Error('当前浏览器不支持剪贴板');
      }
      var cols = gridCols(grid);
      notify('已复制 ' + cols + ' 列 × ' + grid.rows.length + ' 行，到在线表格点 A1 后 Ctrl+V', 'ok');
      return { ok: true, cols: cols, rows: grid.rows.length };
    } catch (e) {
      notify('复制失败：' + (e && e.message ? e.message : e), 'err');
      return { ok: false, error: e && e.message ? e.message : String(e) };
    }
  }

  /** 给按钮挂上右键 = 复制表格数据 */
  function bindContextMenu(btn, getGrid, meta) {
    if (!btn) return;
    btn.addEventListener('contextmenu', function (ev) {
      ev.preventDefault();
      var g = typeof getGrid === 'function' ? getGrid() : getGrid;
      copyToClipboard(g, meta);
    });
  }

  global.HBWecomSync = {
    MAGIC: MAGIC,
    SCHEMA_VER: SCHEMA_VER,
    push: push,
    copyToClipboard: copyToClipboard,
    bindContextMenu: bindContextMenu,
    gridToTsv: gridToTsv,
    expandGrid: expandGrid,
    cellToText: cellToText,
    sanitizeText: sanitizeText,
    formatNumber: formatNumber,
    colLetter: colLetter,
    rangeOf: rangeOf,
    gridCols: gridCols,
    assertHeaderShape: assertHeaderShape,
    buildFileName: buildFileName,
    stampOf: stampOf,
  };
})(window);
