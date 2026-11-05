/*!
 * shared/clipboard.js —— 全站统一剪贴板（移动端优先）
 *
 * 为什么需要它：移动浏览器（尤其华为/微信/夸克等内核，以及被 iframe 嵌入的页面）里
 *   · navigator.clipboard.writeText 需要安全上下文 + 权限策略 + 文档处于焦点，常直接 reject；
 *   · 而 reject 之后再走「屏外 textarea + execCommand」时已经脱离用户手势，必然失败；
 *   · 屏外（top:-9999px）或 opacity:0 的 textarea 会被多数移动内核判定为不可见而拒绝复制。
 * 因此本模块把顺序反过来：移动端先用「视口内 1px 的真实 textarea + execCommand」同步复制，
 * 仍未成功时再异步尝试 clipboard API，最后请求顶层窗口代理，三层都失败才交给手动复制浮层。
 *
 * 接口：
 *   HBCopy.copy(text)            -> Promise<boolean>   推荐；需在点击事件里直接调用
 *   HBCopy.copySync(text)        -> boolean           同步版本（用于版本号徽章等即时反馈场景）
 *   HBCopy.showManual(text,opt)  -> { close() }       手动复制浮层；opt: {title, hint, onSuccess}
 *   HBCopy.isMobile()
 */
(function (global) {
  'use strict';

  if (global.HBCopy) return;

  var IS_FRAME = false;
  try { IS_FRAME = global.top !== global.self; } catch (e) { IS_FRAME = true; }

  var MOBILE_RE = /Android|iPhone|iPad|iPod|HarmonyOS|OpenHarmony|Mobile|MicroMessenger|MQQBrowser|HuaweiBrowser|HeyTapBrowser|UCBrowser|Quark/i;

  function isMobile() {
    try {
      if (MOBILE_RE.test(navigator.userAgent || '')) return true;
      // UA 被裁剪时用触摸点数兜底，避免把桌面触屏误判成手机（取短边阈值）
      return (navigator.maxTouchPoints || 0) > 1 && Math.min(global.innerWidth, global.innerHeight) < 900;
    } catch (e) { return false; }
  }

  function norm(t) { return t == null ? '' : String(t); }

  /* ---------------- 同步复制（必须在用户手势的同步调用栈内） ---------------- */
  function execCopy(text) {
    var body = document.body || document.documentElement;
    if (!body) return false;

    var prevActive = null;
    try { prevActive = document.activeElement; } catch (e) {}

    var ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.setAttribute('aria-hidden', 'true');
    // 必须留在视口内并保持可渲染（1px + opacity .01），不能用 -9999px / opacity:0
    ta.style.cssText = 'position:fixed;top:50%;left:50%;width:1px;height:1px;'
      + 'padding:0;margin:0;border:0;outline:0;box-shadow:none;background:transparent;'
      + 'color:transparent;font-size:16px;line-height:1;resize:none;overflow:hidden;'
      + 'opacity:0.01;z-index:2147483647;';
    body.appendChild(ta);

    var ok = false;
    try {
      // iOS 下 readonly 的 textarea 调 select() 无效，需临时放开编辑态
      if (isMobile()) { ta.readOnly = false; ta.contentEditable = 'true'; }
      try { if (IS_FRAME && global.focus) global.focus(); } catch (e) {}
      ta.focus({ preventScroll: true });
      ta.select();
      try { ta.setSelectionRange(0, text.length); } catch (e) {}
      ok = document.execCommand('copy');
    } catch (e) { ok = false; }

    try { body.removeChild(ta); } catch (e) {}
    try { if (prevActive && prevActive.focus) prevActive.focus({ preventScroll: true }); } catch (e) {}
    return !!ok;
  }

  function copySync(text) {
    text = norm(text);
    if (!text) return false;
    try { return execCopy(text); } catch (e) { return false; }
  }

  /* ---------------- 异步 clipboard API ---------------- */
  function apiCopy(text) {
    return new Promise(function (resolve) {
      var cb = navigator.clipboard;
      if (!cb || typeof cb.writeText !== 'function') return resolve(false);
      try {
        cb.writeText(text).then(function () { resolve(true); }, function () { resolve(false); });
      } catch (e) { resolve(false); }
    });
  }

  /* ---------------- 顶层窗口代理（iframe 内最后一道自动方案） ---------------- */
  var seq = 0;
  var pending = {};

  function askParent(text) {
    if (!IS_FRAME) return Promise.resolve(false);
    var topWin = null;
    try { topWin = global.top; } catch (e) { return Promise.resolve(false); }
    if (!topWin || topWin === global) return Promise.resolve(false);

    var id = 'hb' + (++seq) + '_' + Date.now();
    return new Promise(function (resolve) {
      var settled = false;
      var timer = global.setTimeout(function () { settle(false); }, 1500);
      function settle(ok) {
        if (settled) return;
        settled = true;
        try { global.clearTimeout(timer); } catch (e) {}
        delete pending[id];
        resolve(!!ok);
      }
      pending[id] = settle;
      try {
        topWin.postMessage({ type: 'hb:copy:request', id: id, text: text }, '*');
      } catch (e) { settle(false); }
    });
  }

  function sameOrigin(origin) {
    if (!origin || origin === 'null') return true; // file:// 或 sandbox 场景
    try { return origin === global.location.origin; } catch (e) { return false; }
  }

  global.addEventListener('message', function (ev) {
    var d = ev.data;
    if (!d || typeof d !== 'object') return;
    if (d.type === 'hb:copy:result' && d.id) {
      var fn = pending[d.id];
      if (fn) fn(!!d.ok);
      return;
    }
    // 顶层：替同源 iframe 执行复制
    if (d.type === 'hb:copy:request' && d.id && !IS_FRAME) {
      if (!sameOrigin(ev.origin)) return;
      var src = ev.source || null;
      var text = norm(d.text);
      var reply = function (ok) {
        try { if (src) src.postMessage({ type: 'hb:copy:result', id: d.id, ok: !!ok }, '*'); } catch (e) {}
      };
      if (!text) return reply(false);
      if (execCopy(text)) return reply(true);
      apiCopy(text).then(reply);
    }
  }, false);

  /* ---------------- 统一入口 ---------------- */
  function copy(text) {
    text = norm(text);
    if (!text) return Promise.resolve(false);

    if (isMobile()) {
      // 移动端：同步路径唯一可靠，且此刻仍在用户手势内
      if (execCopy(text)) return Promise.resolve(true);
      return apiCopy(text).then(function (ok) { return ok ? true : askParent(text); });
    }
    // 桌面端：clipboard API 不抢焦点，优先；其失败后 execCommand 在桌面上不依赖手势
    return apiCopy(text).then(function (ok) {
      if (ok) return true;
      if (execCopy(text)) return true;
      return askParent(text);
    });
  }

  /* ---------------- 手动复制浮层 ---------------- */
  function showManual(text, opts) {
    opts = opts || {};
    text = norm(text);

    var old = document.getElementById('hb-copy-manual');
    if (old && old.parentNode) old.parentNode.removeChild(old);

    var mask = document.createElement('div');
    mask.id = 'hb-copy-manual';
    mask.style.cssText = 'position:fixed;inset:0;z-index:2147483000;background:rgba(15,23,42,.52);'
      + 'display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box;';

    var box = document.createElement('div');
    box.style.cssText = 'width:100%;max-width:440px;max-height:88vh;overflow:auto;background:#fff;border-radius:14px;'
      + 'box-shadow:0 18px 50px rgba(15,23,42,.32);padding:16px;box-sizing:border-box;'
      + 'font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;';

    var title = document.createElement('div');
    title.textContent = opts.title || '文本已全选，请长按并选择「复制」';
    title.style.cssText = 'font-size:15px;font-weight:600;color:#0f2a5c;margin-bottom:5px;line-height:1.4;';

    var hint = document.createElement('div');
    hint.textContent = opts.hint || '浏览器未允许自动写入剪贴板时，可长按下方文本 →「全选」→「复制」。';
    hint.style.cssText = 'font-size:12.5px;line-height:1.5;color:#5b6b86;margin-bottom:10px;';

    var ta = document.createElement('textarea');
    ta.value = text;
    ta.readOnly = true;
    ta.spellcheck = false;
    ta.setAttribute('wrap', 'off');
    ta.style.cssText = 'display:block;width:100%;height:min(34vh,190px);box-sizing:border-box;'
      + 'padding:10px;border:1.5px solid #b9cdea;border-radius:9px;background:#f8fbff;color:#12233d;'
      + 'font-size:16px;line-height:1.55;resize:none;overflow:auto;white-space:pre;'
      + '-webkit-user-select:text;user-select:text;-webkit-touch-callout:default;touch-action:auto;'
      + 'font-family:ui-monospace,Menlo,Consolas,monospace;';

    var status = document.createElement('div');
    status.style.cssText = 'font-size:12.5px;line-height:1.4;margin-top:8px;min-height:16px;color:#1a7f4b;';

    var row = document.createElement('div');
    row.style.cssText = 'display:flex;gap:8px;margin-top:12px;';

    function mkBtn(label, kind) {
      var b = document.createElement('button');
      b.type = 'button';
      b.textContent = label;
      var base = 'flex:1;padding:11px 8px;border-radius:9px;font-size:14px;cursor:pointer;'
        + 'font-family:inherit;line-height:1.2;';
      if (kind === 'primary') base += 'border:0;background:#3658d3;color:#fff;font-weight:600;';
      else base += 'border:1px solid #c9d8f0;background:#fff;color:#22406e;';
      b.style.cssText = base;
      return b;
    }

    var retry = mkBtn('再试一次', 'primary');
    var selectAll = mkBtn('全选');
    var closeBtn = mkBtn('关闭');

    function toast(msg) { status.textContent = msg || ''; }

    function doClose() {
      try { mask.parentNode && mask.parentNode.removeChild(mask); } catch (e) {}
    }

    retry.addEventListener('click', function () {
      // 这次点击仍在用户手势内，同步复制成功率最高
      if (execCopy(text)) {
        toast('已复制');
        if (typeof opts.onSuccess === 'function') { try { opts.onSuccess(); } catch (e) {} }
        global.setTimeout(doClose, 260);
      } else {
        toast('仍未成功，请长按上方文本手动复制');
      }
    });

    selectAll.addEventListener('click', function () {
      try {
        ta.focus({ preventScroll: true });
        ta.select();
        ta.setSelectionRange(0, text.length);
        toast('已全选，长按或点系统菜单「复制」');
      } catch (e) { toast('请长按文本选择'); }
    });

    closeBtn.addEventListener('click', doClose);
    mask.addEventListener('click', function (e) { if (e.target === mask) doClose(); });

    row.appendChild(retry);
    row.appendChild(selectAll);
    row.appendChild(closeBtn);
    box.appendChild(title);
    box.appendChild(hint);
    box.appendChild(ta);
    box.appendChild(status);
    box.appendChild(row);
    mask.appendChild(box);
    (document.body || document.documentElement).appendChild(mask);

    // 打开即全选一次，用户可直接点系统「复制」
    try {
      ta.focus({ preventScroll: true });
      ta.select();
      ta.setSelectionRange(0, text.length);
    } catch (e) {}

    return { close: doClose, element: mask };
  }

  global.HBCopy = {
    copy: copy,
    copySync: copySync,
    showManual: showManual,
    isMobile: isMobile,
    isEmbedded: IS_FRAME
  };
})(window);
