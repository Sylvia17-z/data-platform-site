/* ============================================================
 * shared/data-backup.js — 本机数据一键备份 / 恢复
 * ------------------------------------------------------------
 * 目的：平台的业务数据全部落在浏览器本机（localStorage + IndexedDB），
 *       按 origin 隔离。换电脑、换浏览器、换入口链接时数据不通用。
 *       本模块把整份本机数据导出为单个 JSON 文件，在任意入口一键还原。
 *
 * 覆盖范围
 *   - localStorage：除登录态 / 会话 / UI 偏好外的全部键
 *   - IndexedDB：本 origin 下全部数据库、全部 store、全部记录
 *
 * 不包含
 *   - 账号密码、登录 token（跨机无效且不安全）
 *
 * 依赖：无。IIFE，暴露 window.DATA_BACKUP = { open, collect, download, restore }
 * ============================================================ */
(function () {
  'use strict';
  if (window.DATA_BACKUP) return;

  var KIND = 'haibei-data-backup';
  var VER = 1;
  var DATE_TAG = '__hbdate__';

  // 已知数据库（indexedDB.databases() 不可用时的兜底）
  var KNOWN_DBS = ['dailyPerformance', 'dpBridge'];

  // 不导出 / 不恢复的键：登录态、心跳、UI 偏好、权限缓存（有时效，跨机无意义）
  var SKIP = [
    /^site-admin-session$/,
    /^workbuddy-cloud\.session\./,
    /^site-auth-ping$/,
    /^dp_auth_backup$/,
    /^acl_current_operator$/,
    /^acl_cache$/,
    /^sidebar-collapsed$/,
    /_filters_collapsed$/
  ];

  // 本平台键前缀：恢复时用于清理备份中不存在的残留键
  var PLATFORM_PREFIX = ['dailyPerformance_', 'meidaStore_', 'med_', 'monthly_', 'stats_', 'dp_', 'acl_'];

  function isSkip(k) {
    for (var i = 0; i < SKIP.length; i++) if (SKIP[i].test(k)) return true;
    return false;
  }
  function isPlatform(k) {
    for (var i = 0; i < PLATFORM_PREFIX.length; i++) if (k.indexOf(PLATFORM_PREFIX[i]) === 0) return true;
    return false;
  }
  function pad(n) { return n < 10 ? '0' + n : '' + n; }
  function stamp(d) {
    return '' + d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '_' + pad(d.getHours()) + pad(d.getMinutes());
  }
  function isoLocal(d) {
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' +
      pad(d.getHours()) + ':' + pad(d.getMinutes()) + ':' + pad(d.getSeconds());
  }
  function fmtBytes(n) {
    if (!n) return '0 B';
    if (n < 1024) return n + ' B';
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + ' KB';
    if (n < 1024 * 1024 * 1024) return (n / 1048576).toFixed(2) + ' MB';
    return (n / 1073741824).toFixed(2) + ' GB';
  }

  /* ---------------- 日期保真：IDB 内可能存 Date 对象 ---------------- */
  function replacer(k, v) {
    var raw = this[k];
    if (raw instanceof Date) {
      var o = {};
      o[DATE_TAG] = raw.toISOString();
      return o;
    }
    return v;
  }
  function reviver(k, v) {
    if (v && typeof v === 'object' && typeof v[DATE_TAG] === 'string') {
      var d = new Date(v[DATE_TAG]);
      if (!isNaN(d.getTime())) return d;
    }
    return v;
  }

  /* ---------------- localStorage ---------------- */
  function readLocal() {
    var map = {}, bytes = 0, skipped = 0;
    for (var i = 0; i < localStorage.length; i++) {
      var k = localStorage.key(i);
      if (!k) continue;
      if (isSkip(k)) { skipped++; continue; }
      var v = null;
      try { v = localStorage.getItem(k); } catch (e) { continue; }
      if (v == null) continue;
      map[k] = v;
      bytes += k.length + v.length;
    }
    return { map: map, bytes: bytes, skipped: skipped };
  }

  /* ---------------- IndexedDB 读 ---------------- */
  function openDb(name, version) {
    return new Promise(function (res, rej) {
      var req = version ? indexedDB.open(name, version) : indexedDB.open(name);
      req.onsuccess = function () { res(req.result); };
      req.onerror = function () { rej(req.error || new Error('idb-open-failed')); };
      req.onblocked = function () { rej(new Error('idb-blocked')); };
    });
  }

  function listDbNames() {
    if (indexedDB.databases) {
      return indexedDB.databases().then(function (list) {
        var names = (list || []).map(function (d) { return d && d.name; }).filter(Boolean);
        return names.length ? names : KNOWN_DBS.slice();
      }).catch(function () { return KNOWN_DBS.slice(); });
    }
    return Promise.resolve(KNOWN_DBS.slice());
  }

  function readStore(db, name) {
    return new Promise(function (res, rej) {
      var rows = [];
      var tx;
      try { tx = db.transaction(name, 'readonly'); } catch (e) { return res(rows); }
      var req;
      try { req = tx.objectStore(name).openCursor(); } catch (e) { return res(rows); }
      req.onsuccess = function (e) {
        var c = e.target.result;
        if (c) { rows.push({ k: c.key, v: c.value }); c.continue(); }
        else res(rows);
      };
      req.onerror = function () { rej(req.error); };
    });
  }

  function countStore(db, name) {
    return new Promise(function (res) {
      try {
        var tx = db.transaction(name, 'readonly');
        var req = tx.objectStore(name).count();
        req.onsuccess = function () { res(req.result || 0); };
        req.onerror = function () { res(0); };
      } catch (e) { res(0); }
    });
  }

  function readAllIdb() {
    return listDbNames().then(function (names) {
      var out = {}, records = 0, bytes = 0, failed = [];
      return names.reduce(function (p, name) {
        return p.then(function () {
          return openDb(name).then(function (db) {
            var stores = Array.prototype.slice.call(db.objectStoreNames);
            if (!stores.length) { db.close(); return; }
            return stores.reduce(function (q, s) {
              return q.then(function () {
                return readStore(db, s).then(function (rows) {
                  (out[name] = out[name] || {})[s] = rows;
                  records += rows.length;
                  for (var i = 0; i < rows.length; i++) {
                    try {
                      var raw = rows[i].v;
                      bytes += (typeof raw === 'string' ? raw.length : JSON.stringify(raw, replacer).length);
                    } catch (e) { /* 不可序列化：忽略体积统计 */ }
                  }
                }).catch(function () { failed.push(name + '/' + s); });
              });
            }, Promise.resolve()).then(function () { try { db.close(); } catch (e) {} });
          }).catch(function () { failed.push(name + '/*'); });
        });
      }, Promise.resolve()).then(function () {
        return { dbs: out, records: records, bytes: bytes, failed: failed };
      });
    });
  }

  /* 兜底：indexedDB.databases() 在个别浏览器/首次枚举时可能长时间不返回，
     不能让面板一直停在「统计中…」，到点用中性占位收尾。 */
  function withTimeout(promise, ms, fallback) {
    return new Promise(function (resolve) {
      var done = false;
      var timer = setTimeout(function () {
        if (done) return;
        done = true;
        resolve(fallback);
      }, ms);
      promise.then(function (v) {
        if (done) return;
        done = true;
        clearTimeout(timer);
        resolve(v);
      }, function () {
        if (done) return;
        done = true;
        clearTimeout(timer);
        resolve(fallback);
      });
    });
  }

  function countAllIdb() {
    return listDbNames().then(function (names) {
      var total = 0, dbNames = [];
      return names.reduce(function (p, name) {
        return p.then(function () {
          return openDb(name).then(function (db) {
            var stores = Array.prototype.slice.call(db.objectStoreNames);
            if (!stores.length) { db.close(); return; }
            dbNames.push(name);
            return Promise.all(stores.map(function (s) { return countStore(db, s); })).then(function (ns) {
              ns.forEach(function (n) { total += n; });
              try { db.close(); } catch (e) {}
            });
          }).catch(function () {});
        });
      }, Promise.resolve()).then(function () { return { records: total, dbs: dbNames }; });
    });
  }

  /* ---------------- 打包 ---------------- */
  function collect() {
    var ls = readLocal();
    return readAllIdb().then(function (idb) {
      var payload = {
        kind: KIND,
        version: VER,
        exportedAt: isoLocal(new Date()),
        exportedAtISO: new Date().toISOString(),
        siteVer: (typeof window.SITE_VER === 'string' ? window.SITE_VER : ''),
        origin: location.origin + location.pathname,
        ua: navigator.userAgent,
        localStorage: ls.map,
        indexedDB: idb.dbs,
        stats: {
          lsKeys: Object.keys(ls.map).length,
          lsBytes: ls.bytes,
          lsSkipped: ls.skipped,
          idbRecords: idb.records,
          idbTables: Object.keys(idb.dbs).reduce(function (n, d) { return n + Object.keys(idb.dbs[d] || {}).length; }, 0),
          idbBytes: idb.bytes,
          idbDbs: Object.keys(idb.dbs),
          idbFailed: idb.failed
        }
      };
      return payload;
    });
  }

  /* 导出：体积大（真实数据可达数十 MB），优先 gzip 压缩，便于手机 / 微信传输；
     浏览器不支持 CompressionStream 时退化为明文 JSON，恢复端两种都认。 */
  function buildBlob(text) {
    if (typeof CompressionStream === 'undefined') {
      return Promise.resolve({ blob: new Blob([text], { type: 'application/json;charset=utf-8' }), ext: '.json', rawBytes: text.length });
    }
    var compressed;
    try {
      compressed = new Blob([text]).stream().pipeThrough(new CompressionStream('gzip'));
    } catch (e) {
      return Promise.resolve({ blob: new Blob([text], { type: 'application/json;charset=utf-8' }), ext: '.json', rawBytes: text.length });
    }
    return new Response(compressed).blob().then(function (b) {
      return { blob: new Blob([b], { type: 'application/gzip' }), ext: '.json.gz', rawBytes: text.length };
    }).catch(function () {
      return { blob: new Blob([text], { type: 'application/json;charset=utf-8' }), ext: '.json', rawBytes: text.length };
    });
  }

  function saveBlob(blob, ext) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = '海蓓数据备份_' + stamp(new Date()) + ext;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(function () {
      try { document.body.removeChild(a); } catch (e) {}
      setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    }, 300);
  }

  function download(payload) {
    var text = JSON.stringify(payload, replacer);
    return buildBlob(text).then(function (r) {
      saveBlob(r.blob, r.ext);
      return { size: r.blob.size, rawBytes: r.rawBytes, ext: r.ext, gz: r.ext === '.json.gz' };
    });
  }

  /* ---------------- IndexedDB 写 ---------------- */
  function ensureDb(name, storeNames) {
    return new Promise(function (res, rej) {
      var req = indexedDB.open(name);
      req.onupgradeneeded = function () {
        var db = req.result;
        storeNames.forEach(function (s) { try { if (!db.objectStoreNames.contains(s)) db.createObjectStore(s); } catch (e) {} });
      };
      req.onsuccess = function () {
        var db = req.result;
        var missing = storeNames.filter(function (s) { return !db.objectStoreNames.contains(s); });
        if (!missing.length) return res(db);
        var next = db.version + 1;
        try { db.close(); } catch (e) {}
        var r2 = indexedDB.open(name, next);
        r2.onupgradeneeded = function () {
          var d2 = r2.result;
          missing.forEach(function (s) { try { d2.createObjectStore(s); } catch (e) {} });
        };
        r2.onsuccess = function () { res(r2.result); };
        r2.onerror = function () { rej(r2.error || new Error('idb-upgrade-failed')); };
      };
      req.onerror = function () { rej(req.error || new Error('idb-open-failed')); };
    });
  }

  function writeStore(db, name, rows) {
    var BATCH = 1500;
    var hasInline = false;
    try { hasInline = !!db.transaction(name, 'readonly').objectStore(name).keyPath; } catch (e) {}

    function chunk(from) {
      return new Promise(function (res, rej) {
        var slice = rows.slice(from, from + BATCH);
        var tx;
        try { tx = db.transaction(name, 'readwrite'); } catch (e) { return rej(e); }
        var os = tx.objectStore(name);
        if (from === 0) { try { os.clear(); } catch (e) {} }
        for (var i = 0; i < slice.length; i++) {
          try {
            if (hasInline) os.put(slice[i].v);
            else os.put(slice[i].v, slice[i].k);
          } catch (e) { /* 单条失败不阻塞整批 */ }
        }
        tx.oncomplete = function () { res(from + slice.length); };
        tx.onerror = function () { rej(tx.error || new Error('idb-write-failed')); };
        tx.onabort = function () { rej(tx.error || new Error('idb-aborted')); };
      });
    }

    return (function loop(i) {
      if (i >= rows.length) {
        if (!rows.length) {
          return new Promise(function (res, rej) {
            var tx;
            try { tx = db.transaction(name, 'readwrite'); } catch (e) { return rej(e); }
            try { tx.objectStore(name).clear(); } catch (e) {}
            tx.oncomplete = function () { res(0); };
            tx.onerror = function () { rej(tx.error); };
          });
        }
        return Promise.resolve(0);
      }
      return chunk(i).then(function (next) { return loop(next); });
    })(0);
  }

  /* ---------------- 恢复 ---------------- */
  function apply(payload, onProgress) {
    var report = { idbOk: 0, idbFail: [], lsOk: 0, lsFail: [], dropped: 0 };
    var dbs = payload.indexedDB || {};
    var dbNames = Object.keys(dbs);

    function stepIdb(i) {
      if (i >= dbNames.length) return Promise.resolve();
      var name = dbNames[i];
      var stores = dbs[name] || {};
      var storeNames = Object.keys(stores);
      if (!storeNames.length) return stepIdb(i + 1);
      if (onProgress) onProgress('正在写入数据库 ' + name + '（' + (i + 1) + '/' + dbNames.length + '）…');
      return ensureDb(name, storeNames).then(function (db) {
        return storeNames.reduce(function (p, s) {
          return p.then(function () {
            return writeStore(db, s, stores[s] || []).then(function () {
              report.idbOk += (stores[s] || []).length;
            }).catch(function (e) {
              report.idbFail.push(name + '/' + s + '：' + ((e && e.message) || '写入失败'));
            });
          });
        }, Promise.resolve()).then(function () { try { db.close(); } catch (e) {} })
          .then(function () { return stepIdb(i + 1); });
      }).catch(function (e) {
        report.idbFail.push(name + '：' + ((e && e.message) || '打开失败'));
        return stepIdb(i + 1);
      });
    }

    return stepIdb(0).then(function () {
      if (onProgress) onProgress('正在写入浏览器本地存储…');
      var map = payload.localStorage || {};

      // 1) 清掉本平台残留、但备份中不存在的键（彻底还原为备份状态）
      var drop = [];
      for (var i = 0; i < localStorage.length; i++) {
        var k = localStorage.key(i);
        if (!k || isSkip(k)) continue;
        if (isPlatform(k) && !(k in map)) drop.push(k);
      }
      for (var j = 0; j < drop.length; j++) {
        try { localStorage.removeItem(drop[j]); report.dropped++; } catch (e) {}
      }

      // 2) 小键优先写入，避免大键吃满配额导致配置项丢失
      var keys = Object.keys(map).sort(function (a, b) {
        return (map[a] ? map[a].length : 0) - (map[b] ? map[b].length : 0);
      });
      for (var m = 0; m < keys.length; m++) {
        var kk = keys[m];
        try { localStorage.setItem(kk, map[kk]); report.lsOk++; }
        catch (e) { report.lsFail.push(kk); }
      }
      return report;
    });
  }

  function validate(p) {
    if (!p || typeof p !== 'object') throw new Error('不是有效的备份文件');
    if (p.kind !== KIND) throw new Error('文件类型不匹配（不是本平台的备份）');
    if (typeof p.version !== 'number' || p.version > VER) throw new Error('备份版本 ' + p.version + ' 高于当前程序支持的 ' + VER);
    if (!p.localStorage && !p.indexedDB) throw new Error('备份内容为空');
    return true;
  }

  /* ---------------- UI ---------------- */
  var CSS = [
    '.bk-open{display:inline-flex;align-items:center;gap:5px;border:1px solid #cfd9e6;background:#fff;color:#123c73;border-radius:6px;padding:5px 10px;font-size:12.5px;line-height:1.2;font-family:inherit;cursor:pointer;white-space:nowrap}',
    '.bk-open:hover{background:#f2f6fc;border-color:#9fb6d4}',
    '.bk-mask{position:fixed;inset:0;background:rgba(16,32,56,.42);display:flex;align-items:center;justify-content:center;z-index:99990;padding:16px}',
    '.bk-panel{background:#fff;border-radius:12px;width:min(560px,100%);max-height:88vh;overflow:auto;box-shadow:0 16px 48px rgba(16,32,56,.28);font-family:inherit;color:#1c2b3f}',
    '.bk-hd{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid #e7edf5}',
    '.bk-hd b{font-size:15px;color:#123c73}',
    '.bk-x{border:0;background:transparent;font-size:17px;color:#7b8ba3;cursor:pointer;line-height:1;padding:2px 6px;font-family:inherit}',
    '.bk-x:hover{color:#123c73}',
    '.bk-body{padding:16px 18px 18px}',
    '.bk-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:10px;margin-bottom:14px}',
    '.bk-cell{background:#f6f9fd;border:1px solid #e7edf5;border-radius:8px;padding:10px 12px}',
    '.bk-cell .lb{font-size:11.5px;color:#7b8ba3;margin-bottom:4px}',
    '.bk-cell .vl{font-size:15px;font-weight:600;color:#123c73;font-variant-numeric:tabular-nums}',
    '.bk-acts{display:flex;flex-direction:column;gap:9px;margin:4px 0 12px}',
    '.bk-btn{display:inline-flex;align-items:center;justify-content:center;gap:7px;border-radius:8px;padding:11px 14px;font-size:13.5px;line-height:1.2;font-family:inherit;cursor:pointer;border:1px solid transparent}',
    '.bk-btn:disabled{opacity:.55;cursor:not-allowed}',
    '.bk-btn.primary{background:#123c73;color:#fff;border-color:#123c73}',
    '.bk-btn.primary:hover:not(:disabled){background:#0e2f5c}',
    '.bk-btn.ghost{background:#fff;color:#123c73;border-color:#c7d4e5}',
    '.bk-btn.ghost:hover:not(:disabled){background:#f2f6fc}',
    '.bk-btn.danger{background:#b3261e;color:#fff;border-color:#b3261e}',
    '.bk-btn.danger:hover:not(:disabled){background:#8f1d17}',
    '.bk-tip{font-size:12px;line-height:1.7;color:#5b6b82;background:#f6f9fd;border:1px solid #e7edf5;border-left:3px solid #123c73;border-radius:6px;padding:9px 11px}',
    '.bk-tip b{color:#123c73}',
    '.bk-toast{position:fixed;left:50%;top:22px;transform:translateX(-50%);z-index:99999;background:#0f5132;color:#fff;border-radius:10px;padding:12px 18px;font-size:13px;line-height:1.6;box-shadow:0 10px 30px rgba(16,32,56,.28);max-width:min(420px,92vw);text-align:center;transition:opacity .5s,transform .5s}',
    '.bk-toast.out{opacity:0;transform:translateX(-50%) translateY(-10px)}',
    '.bk-must{background:#fff4e5;border:1px solid #ffd7a8;color:#8a4b00;border-radius:6px;padding:8px 10px;margin-bottom:7px;font-size:12.5px}',
    '.bk-log{margin-top:11px;font-size:12px;line-height:1.7;color:#5b6b82;min-height:0}',
    '.bk-log .ok{color:#166534}.bk-log .err{color:#b3261e}.bk-log .warn{color:#a15c00}',
    '.bk-confirm{margin-top:12px;border:1px solid #f0d9a8;background:#fffbf0;border-radius:8px;padding:11px 12px;font-size:12.5px;line-height:1.75;color:#6b5320}',
    '.bk-confirm b{color:#8a5a00}',
    '.bk-kv{display:flex;gap:8px;flex-wrap:wrap;margin-top:6px}',
    '.bk-kv span{background:#fff;border:1px solid #eee1c2;border-radius:5px;padding:2px 7px;font-size:11.5px;color:#6b5320}',
    '.bk-foot{display:flex;gap:8px;justify-content:flex-end;margin-top:12px}',
    '@media(max-width:640px){.bk-open .bk-open-txt{display:none}.bk-grid{grid-template-columns:repeat(2,1fr)}.bk-body{padding:14px}}'
  ].join('');

  var mask = null, ui = null, pending = null, lastReport = null;

  function el(tag, cls, html) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  }

  function log(msg, kind) {
    if (!ui) return;
    var line = el('div', kind || '');
    line.textContent = msg;
    ui.log.innerHTML = '';
    ui.log.appendChild(line);
  }

  function logKeep(msg, kind) {
    if (!ui) return;
    var line = el('div', kind || '');
    line.textContent = msg;
    ui.log.appendChild(line);
  }

  function ensureStyle() {
    if (document.getElementById('bk-style')) return;
    var style = document.createElement('style');
    style.id = 'bk-style';
    style.textContent = CSS;
    document.head.appendChild(style);
  }

  function build() {
    ensureStyle();
    if (mask) return;

    mask = el('div', 'bk-mask');
    var panel = el('div', 'bk-panel');
    panel.innerHTML = [
      '<div class="bk-hd"><b>数据备份 / 恢复</b><button class="bk-x" type="button" aria-label="关闭">✕</button></div>',
      '<div class="bk-body">',
      '  <div class="bk-grid">',
      '    <div class="bk-cell"><div class="lb">本地存储键</div><div class="vl" data-k="lsn">—</div></div>',
      '    <div class="bk-cell"><div class="lb">本地存储体积</div><div class="vl" data-k="lsb">—</div></div>',
      '    <div class="bk-cell"><div class="lb">数据表项</div><div class="vl" data-k="idbn">统计中…</div></div>',
      '    <div class="bk-cell"><div class="lb">最近数据更新</div><div class="vl" data-k="ts">—</div></div>',
      '  </div>',
      '  <div class="bk-acts">',
      '    <button class="bk-btn primary" type="button" data-act="export">⬇️ 下载本机数据备份</button>',
      '    <button class="bk-btn ghost" type="button" data-act="pick">⬆️ 从备份文件恢复</button>',
      '  </div>',
      '  <div class="bk-tip">备份文件包含：订单明细、历史消费、VAB 名单、月目标、手工调整、美大店铺数据、医管考核数据等<b>全部业务数据</b>。<br>不包含<b>账号密码 / 登录状态</b>。换电脑、换浏览器、换入口链接后，导入这份文件即可恢复现状。<br>恢复支持 <b>.json</b> 与 <b>.json.gz</b> 两种格式（导出的压缩版体积更小，手机也能传）。</div>',
      '  <div class="bk-log" data-k="log"></div>',
      '  <div data-k="confirm"></div>',
      '  <div class="bk-foot" data-k="foot"></div>',
      '</div>'
    ].join('');
    mask.appendChild(panel);
    document.body.appendChild(mask);

    var q = function (k) { return panel.querySelector('[data-k="' + k + '"]'); };
    ui = {
      panel: panel,
      lsn: q('lsn'), lsb: q('lsb'), idbn: q('idbn'), ts: q('ts'),
      log: q('log'), confirm: q('confirm'), foot: q('foot'),
      file: (function () {
        var f = document.createElement('input');
        f.type = 'file';
        /* 不设 accept 白名单：导出文件名是「海蓓数据备份_日期.json.gz」，
           Windows 文件对话框按最后一个扩展名（.gz）过滤，accept=".json" 会让
           备份文件变灰、根本选不中（2026-11-05 用户实测反馈）。
           改为读取后校验内容：选错文件会得到明确提示，不再靠系统过滤。
           隐藏方式用视觉隐藏而非 display:none —— 后者在部分浏览器/扩展下
           无法被程序化 click 可靠唤起（与 11-04 上传按钮同一个坑）。 */
        f.setAttribute('aria-hidden', 'true');
        f.style.cssText = 'position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0 0 0 0);clip-path:inset(50%);white-space:nowrap;border:0;';
        panel.appendChild(f);
        return f;
      })()
    };

    mask.addEventListener('click', function (e) { if (e.target === mask) close(); });
    panel.querySelector('.bk-x').addEventListener('click', close);

    panel.addEventListener('click', function (e) {
      var b = e.target.closest('button[data-act]');
      if (!b) return;
      var act = b.getAttribute('data-act');
      if (act === 'export') doExport(b);
      else if (act === 'pick') ui.file.click();
      else if (act === 'apply') doApply(b);
      else if (act === 'reload') location.reload();
      else if (act === 'close') close();
    });

    ui.file.addEventListener('change', function () {
      var f = ui.file.files && ui.file.files[0];
      ui.file.value = '';
      if (f) readFile(f);
    });

    refreshStats();
  }

  function refreshStats() {
    var ls = readLocal();
    ui.lsn.textContent = Object.keys(ls.map).length + ' 个';
    ui.lsb.textContent = fmtBytes(ls.bytes);
    var ts = 0;
    try { ts = Number(localStorage.getItem('dailyPerformance_snap_ts') || 0); } catch (e) {}
    ui.ts.textContent = ts ? isoLocal(new Date(ts)).slice(5, 16) : '—';
    ui.ts.style.fontSize = ts ? '13px' : '';
    withTimeout(countAllIdb(), 2500, null).then(function (r) {
      if (!ui) return;
      if (!r) { ui.idbn.textContent = '—'; return; }
      ui.idbn.textContent = r.records ? r.records.toLocaleString() + ' 项' : (r.dbs.length ? '空' : '—');
    }).catch(function () { if (ui) ui.idbn.textContent = '—'; });
  }

  function doExport(btn) {
    btn.disabled = true;
    log('正在打包本机数据…');
    var t0 = Date.now();
    collect().then(function (payload) {
      return download(payload).then(function (info) {
        var s = payload.stats;
        var sizeTxt = info.gz
          ? fmtBytes(info.size) + '（原始 ' + fmtBytes(info.rawBytes) + '，已压缩）'
          : fmtBytes(info.size);
        log('已生成备份文件：' + sizeTxt + '（本地存储 ' + s.lsKeys + ' 项 / 数据库 ' +
          s.idbTables + ' 项，用时 ' + ((Date.now() - t0) / 1000).toFixed(1) + ' 秒）', 'ok');
        return payload;
      });
    }).then(function () {
      btn.disabled = false;
    }).catch(function (e) {
      log('打包失败：' + ((e && e.message) || e), 'err');
      btn.disabled = false;
    });
  }

  function gunzip(buf) {
    if (typeof DecompressionStream === 'undefined') {
      return Promise.reject(new Error('当前浏览器不支持解压该备份文件，请改用 Chrome / Edge 打开本页后重试'));
    }
    var s = new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'));
    return new Response(s).text();
  }

  function readFile(file) {
    log('正在读取 ' + file.name + ' …');
    file.arrayBuffer().then(function (buf) {
      var u8 = new Uint8Array(buf);
      var isGz = u8.length > 1 && u8[0] === 0x1f && u8[1] === 0x8b;
      return (isGz ? gunzip(buf) : Promise.resolve(new TextDecoder('utf-8').decode(buf)));
    }).then(function (text) {
      var p = JSON.parse(text, reviver);
      validate(p);
      pending = p;
      showConfirm(file, p);
    }).catch(function (e) {
      var msg = (e && e.message) || String(e);
      if (/JSON|Unexpected token|Unexpected end|Unterminated|not valid/i.test(msg)) {
        msg = '这不是本平台的备份文件（或文件已损坏）。请选择文件名以「海蓓数据备份_」开头的文件，支持 .json 与 .json.gz 两种格式';
      }
      log('无法读取「' + file.name + '」：' + msg, 'err');
    });
  }

  function showConfirm(file, p) {
    var s = p.stats || {};
    var lsKeys = p.localStorage ? Object.keys(p.localStorage).length : 0;
    var idbRecords = 0;
    var dbs = p.indexedDB || {};
    Object.keys(dbs).forEach(function (n) {
      Object.keys(dbs[n] || {}).forEach(function (st) { idbRecords += (dbs[n][st] || []).length; });
    });
    ui.log.innerHTML = '';
    ui.confirm.innerHTML = [
      '<div class="bk-confirm">',
      '  <div class="bk-must">⚠️ 文件已读取成功，还需<b>再点一次下方红色按钮</b>「确认覆盖并恢复」才会真正开始写入。</div>',
      '  <div>即将用此备份<b>覆盖本机现有数据</b>：</div>',
      '  <div class="bk-kv">',
      '    <span>📅 ' + (p.exportedAt || '未知时间') + '</span>',
      '    <span>🗂 ' + lsKeys + ' 个存储键</span>',
      '    <span>💽 ' + idbRecords.toLocaleString() + ' 项数据表</span>',
      '    <span>📦 ' + fmtBytes(file.size) + '</span>',
      '  </div>',
      '  <div style="margin-top:7px;color:#8a5a00">来源：' + (p.origin || '未知') + (s.lsKeys && s.lsKeys !== lsKeys ? '（备份时 ' + s.lsKeys + ' 键）' : '') + '</div>',
      '</div>'
    ].join('');
    ui.foot.innerHTML = '<button class="bk-btn ghost" type="button" data-act="close">取消</button>' +
      '<button class="bk-btn danger" type="button" data-act="apply">确认覆盖并恢复</button>';
    // 选完文件后把确认按钮带进视野并聚焦：否则用户容易以为「选完就恢复好了」
    var ab = ui.foot.querySelector('[data-act="apply"]');
    if (ab) {
      try { ui.confirm.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); } catch (e) {}
      try { ab.scrollIntoView({ block: 'center', behavior: 'smooth' }); } catch (e) {}
      try { ab.focus({ preventScroll: true }); } catch (e) {}
    }
  }

  function doApply(btn) {
    if (!pending) return;
    btn.disabled = true;
    ui.foot.innerHTML = '';
    ui.log.innerHTML = '';
    var t0 = Date.now();
    apply(pending, function (msg) { log(msg); }).then(function (report) {
      lastReport = report;
      var secs = ((Date.now() - t0) / 1000).toFixed(1);
      if (report.lsFail.length) {
        log('数据已导入（用时 ' + secs + ' 秒），但有 ' + report.lsFail.length + ' 个较大的键未能写入本地存储：' +
          report.lsFail.slice(0, 3).join('、') + (report.lsFail.length > 3 ? ' 等' : ''), 'warn');
        logKeep('这些键的正常副本已写入 IndexedDB，不影响使用（页面会从数据库读取）。', 'warn');
      } else {
        log('恢复完成：数据库 ' + report.idbOk.toLocaleString() + ' 条，本地存储 ' + report.lsOk +
          ' 键，清理旧残留 ' + report.dropped + ' 键，用时 ' + secs + ' 秒。', 'ok');
      }
      if (report.idbFail.length) {
        logKeep('部分数据表写入失败：' + report.idbFail.slice(0, 3).join('；'), 'err');
      }
      ui.foot.innerHTML = '<button class="bk-btn primary" type="button" data-act="reload">立即刷新查看</button>';

      // 恢复已落盘，但页面内存里还是旧数据（各模块只在加载时读取一次）。
      // 不重建页面的话，用户看到的仍是旧数字，会以为「没恢复成功」——
      // 2026-11-05 用户实测反馈正是这个现象。故：通知壳层失效内存缓存 + 自动刷新。
      try { window.postMessage({ type: 'dp:reset' }, '*'); } catch (e) {}
      try {
        sessionStorage.setItem('bk_restore_done', JSON.stringify({
          at: (pending && pending.exportedAt) || '',
          lsOk: report.lsOk,
          idbOk: report.idbOk,
          dropped: report.dropped
        }));
      } catch (e) {}
      logKeep('数据已写入本机，2 秒后自动刷新页面…', 'ok');
      setTimeout(function () { try { location.reload(); } catch (e) {} }, 2000);
    }).catch(function (e) {
      log('恢复失败：' + ((e && e.message) || e), 'err');
      ui.foot.innerHTML = '<button class="bk-btn ghost" type="button" data-act="close">关闭</button>';
    });
  }

  /* 恢复完成 → 自动刷新后，给一个明确的落地反馈，避免用户猜「到底生效没有」 */
  function flashRestoreResult() {
    var raw = null;
    try { raw = sessionStorage.getItem('bk_restore_done'); } catch (e) { return; }
    if (!raw) return;
    try { sessionStorage.removeItem('bk_restore_done'); } catch (e) {}
    var info = {};
    try { info = JSON.parse(raw) || {}; } catch (e) { return; }
    var box = document.createElement('div');
    box.className = 'bk-toast';
    box.innerHTML = '✅ 备份已恢复生效<br>' +
      '<span style="font-size:12px;line-height:1.7;display:inline-block;margin-top:3px">' +
      (info.lsOk || 0) + ' 个存储键 · ' + (info.idbOk || 0).toLocaleString() + ' 项数据表' +
      (info.at ? ' · 备份时间 ' + info.at : '') +
      '<br>请重新打开需要的功能模块查看</span>';
    function mount() {
      try { document.body.appendChild(box); } catch (e) { return; }
      setTimeout(function () { box.classList.add('out'); }, 4200);
      setTimeout(function () { try { box.parentNode.removeChild(box); } catch (e) {} }, 5000);
    }
    if (document.body) mount();
    else document.addEventListener('DOMContentLoaded', mount);
  }

  function open() { build(); mask.style.display = 'flex'; refreshStats(); }
  function close() { if (mask) mask.style.display = 'none'; }

  function mountButton() {
    if (window.top !== window.self) return;      // iframe 内不重复注入
    var slot = document.getElementById('sa-badge-slot');
    var host = slot && slot.parentNode;
    if (!host) return;
    if (document.getElementById('bk-open')) return;
    var btn = el('button', 'bk-open', '💾<span class="bk-open-txt">备份</span>');
    btn.id = 'bk-open';
    btn.type = 'button';
    btn.title = '本机数据备份 / 恢复';
    btn.addEventListener('click', open);
    host.insertBefore(btn, slot);
  }

  window.DATA_BACKUP = {
    open: open,
    close: close,
    collect: collect,
    download: download,
    apply: apply,
    validate: validate,
    fmtBytes: fmtBytes
  };

  ensureStyle();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { mountButton(); flashRestoreResult(); });
  } else {
    mountButton();
    flashRestoreResult();
  }
})();
