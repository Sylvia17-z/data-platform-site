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
  var DP_DB_NAME = 'dpBridge';           // 壳层共享数据桥接库（dpBridge/kv）
  var KNOWN_DBS = ['dailyPerformance', DP_DB_NAME, 'med_assess_db'];

  // 不导出 / 不恢复的键：登录态、心跳、UI 偏好、权限缓存（有时效，跨机无意义）
  var SKIP = [
    /^site-admin-session$/,
    /^workbuddy-cloud\.session\./,
    /^site-auth-ping$/,
    /^dp_auth_backup$/,
    /^acl_current_operator$/,
    /^acl_cache$/,
    /^sidebar-collapsed$/,
    /_filters_collapsed$/,
    /^bk_restore_guard$/,     // 恢复守护标记：跟设备/时刻绑定，跨机无意义
    /^bk_restore_done$/,
    /^bk_restore_pending$/
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
    var report = { idbOk: 0, idbFail: [], lsOk: 0, lsFail: [], dropped: 0, tsAligned: 0, copiesAligned: 0 };
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
        try {
          // 先删旧值再写：大键（订单整包常见 2~4MB）在配额紧张时 setItem 会抛错，
          // 若保留旧值，页面读到的仍是恢复前的数据 → 用户会认为「恢复没生效」。
          // 删掉旧值后即使写入失败，页面会回退到 IndexedDB（同样已写入备份内容）。
          try { localStorage.removeItem(kk); } catch (e0) {}
          localStorage.setItem(kk, map[kk]);
          report.lsOk++;
        } catch (e) { report.lsFail.push(kk); }
      }

      // 3) 恢复守护（2026-11-05 用户实测「恢复之后数据还是旧的」根因修复）
      //    只把数据写进本机是不够的：刷新后日业绩(dp_orders)、周客量(dp_stats)、
      //    月度分析(dp_monthly) 会拿本地「最后变更时间」与云端快照比新旧，
      //    而备份里的时间戳是备份那一刻的旧值 → 判定「云端更新」→ 云端旧快照立刻回灌，
      //    把刚恢复的数据覆盖回去，用户看到的还是旧数据。
      //    处理：把各模块的本地时间戳抬到当前时刻（本机成为最新），并留一个 24h
      //    守护标记，各同步脚本在首次对齐时据此强制以本地为准、并把本地回推云端。
      var now = Date.now();
      var TS_ALIGN = [
        'dailyPerformance_snap_ts', 'kpi_cloud_ts_stats', 'kpi_cloud_ts_monthly',
        'medassess_cloud_ts', 'meidaStore_cloud_ts'
      ];
      for (var t = 0; t < TS_ALIGN.length; t++) {
        var tk = TS_ALIGN[t];
        var has = (tk in map);
        if (!has) { try { has = localStorage.getItem(tk) != null; } catch (e) {} }
        if (!has) continue;
        try { localStorage.setItem(tk, String(now)); report.tsAligned++; } catch (e) {}
      }
      try {
        localStorage.setItem('bk_restore_guard', JSON.stringify({
          at: now,
          exp: now + 24 * 3600 * 1000,   // 守护窗口：当天反复刷新都不被云端回灌
          exportedAt: payload.exportedAt || '',
          origin: payload.origin || '',
          lsOk: report.lsOk,
          idbOk: report.idbOk
        }));
      } catch (e) {}
      return report;
    }).then(function (report) {
      // 4) 兜底：清掉备份里没有的桥接缓存，避免壳层把旧值优先推给子页
      return syncCopies(payload, report).then(function () { return report; });
    });
  }

  /* 恢复后的「副本一致性」对齐（2026-11-05 第二次实测「恢复后仍是旧数据」的根因）
     ------------------------------------------------------------------
     同一份业务数据在浏览器里有多份副本，且读取规则并不一致：
       · dailyPerformance_orders / hist / vab
           - localStorage（页内同步首选）
           - IndexedDB dailyPerformance/kv（chuzhen 兜底：ts 更新或条数更多者胜 → 会反过来覆盖 localStorage）
           - IndexedDB dpBridge/kv（壳层 dpGet 以它为准，直接推给子页）
       · med_order_list / med_charge_list / med_vab_list
           - localStorage（首选）
           - IndexedDB med_assess_db/kv（键名 order / charge / vab，localStorage 未命中才回退）
     备份时这些副本可能并不是同一份快照（某次写入失败、旧版本只写了一份、被云端回灌过旧值…）。
     恢复时若原样写回，页内「挑更新的那份」就会把旧副本挑出来盖掉刚恢复的数据。
     故：恢复末尾统一裁决一次（ts 更新 / 条数更多者胜），把同一份写进所有副本。 */
  var SHARED_STR_KEYS = ['dailyPerformance_orders', 'dailyPerformance_hist', 'dailyPerformance_vab'];
  var MED_KEYS = [
    { ls: 'med_order_list', idbKey: 'order' },
    { ls: 'med_charge_list', idbKey: 'charge' },
    { ls: 'med_vab_list', idbKey: 'vab' }
  ];
  var MED_DB = 'med_assess_db';

  function findKvRow(payload, dbName, key) {
    var stores = (payload.indexedDB || {})[dbName] || {};
    var names = Object.keys(stores);
    for (var i = 0; i < names.length; i++) {
      var rows = stores[names[i]] || [];
      for (var j = 0; j < rows.length; j++) {
        var r = rows[j];
        if (r && String(r.k) === String(key)) return r.v;
      }
    }
    return undefined;
  }
  function metaOf(v) {
    var o;
    try { o = (typeof v === 'string') ? JSON.parse(v) : v; } catch (e) { return null; }
    if (Array.isArray(o)) return { ts: 0, n: o.length };
    if (!o || typeof o !== 'object') return null;
    var arr = o.orders || o.list || o.rows || [];
    var ts = Number(o.ts) || Number(o.histTs) || 0;
    return { ts: ts, n: Array.isArray(arr) ? arr.length : 0 };
  }
  /* 两份里取更完整的一份（与页内读取规则保持一致），返回原始值（字符串或对象） */
  function betterOf(a, b) {
    var ma = metaOf(a), mb = metaOf(b);
    if (!ma) return b;
    if (!mb) return a;
    if (mb.ts > ma.ts) return b;
    if (ma.ts > mb.ts) return a;
    return mb.n > ma.n ? b : a;
  }
  function kvWrite(dbName, items, report) {
    if (!items.length) return Promise.resolve();
    return ensureDb(dbName, ['kv']).then(function (db) {
      return new Promise(function (res) {
        var tx;
        try { tx = db.transaction('kv', 'readwrite'); }
        catch (e) { try { db.close(); } catch (_) {} return res(); }
        var st = tx.objectStore('kv');
        items.forEach(function (it) {
          try { if (it.val == null) st.delete(it.key); else st.put(it.val, it.key); } catch (e) {}
        });
        tx.oncomplete = function () {
          report.copiesAligned = (report.copiesAligned || 0) + items.length;
          try { db.close(); } catch (e) {}
          res();
        };
        tx.onerror = tx.onabort = function () { try { db.close(); } catch (e) {} res(); };
      });
    }).catch(function () {});
  }
  function syncCopies(payload, report) {
    var lsMap = payload.localStorage || {};
    var planDp = [], planBridge = [], planMed = [];
    // 只覆盖备份里本来就有的桥接键：dpBridge 是壳层缓存，历史明细整包可达数 MB，
    // 无差别补写会让同一份数据在浏览器里存两遍，容易顶到配额。
    var bridgeHas = {};
    var bKv = ((payload.indexedDB || {})[DP_DB_NAME] || {}).kv || [];
    for (var bi = 0; bi < bKv.length; bi++) {
      var br = bKv[bi];
      if (br && br.k != null) bridgeHas[String(br.k)] = 1;
    }

    SHARED_STR_KEYS.forEach(function (k) {
      var lsVal = lsMap[k];
      var idbVal = findKvRow(payload, 'dailyPerformance', k);
      if (lsVal == null && idbVal == null) {
        // 备份里完全没有 → 本机旧副本也要清掉，否则读取时又会把旧值捡回来
        planDp.push({ key: k, val: null });
        planBridge.push({ key: k, val: null });
        return;
      }
      var win = (lsVal != null && idbVal != null) ? betterOf(lsVal, idbVal) : (lsVal != null ? lsVal : idbVal);
      var str = (typeof win === 'string') ? win : JSON.stringify(win);
      planDp.push({ key: k, val: str });
      // 备份里没有这份桥接缓存 → 清掉本机旧记录，让 dpGet 回退到刚恢复的 localStorage；
      // 有则对齐成同一份（不无差别补写，避免同一份大数据在浏览器里存两遍）
      planBridge.push({ key: k, val: bridgeHas[k] ? str : null });
      // 无论原 LS 是否命中，只要裁决结果不同就回写：IDB 更完整时也要同步到 LS，
      // 否则页内先读到的仍是 LS 里那份旧的，看起来像「没恢复」。
      if (lsVal !== str) { try { localStorage.setItem(k, str); } catch (e) {} }
    });

    MED_KEYS.forEach(function (m) {
      var lsVal = lsMap[m.ls];
      var idbVal = findKvRow(payload, MED_DB, m.idbKey);
      if (lsVal == null && idbVal == null) { planMed.push({ key: m.idbKey, val: null }); return; }
      var win = (lsVal != null && idbVal != null) ? betterOf(lsVal, idbVal) : (lsVal != null ? lsVal : idbVal);
      var str = (typeof win === 'string') ? win : JSON.stringify(win);
      var parsed = null;
      try { parsed = JSON.parse(str); } catch (e) { parsed = null; }
      planMed.push({ key: m.idbKey, val: parsed });
      if (lsVal !== str) { try { localStorage.setItem(m.ls, str); } catch (e) {} }
    });

    return kvWrite('dailyPerformance', planDp, report)
      .then(function () { return kvWrite(DP_DB_NAME, planBridge, report); })
      .then(function () { return kvWrite(MED_DB, planMed, report); });
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
    '.bk-content{font-size:12px;line-height:1.7;color:#41526b;background:#f6f9fd;border:1px dashed #dbe6f3;border-radius:8px;padding:8px 10px;margin:0 0 12px}',
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
      '  <div class="bk-content" data-k="content">—</div>',
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
      lsn: q('lsn'), lsb: q('lsb'), idbn: q('idbn'), ts: q('ts'), content: q('content'),
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

  /* 内容清单：把「备份里到底有什么」直接说清楚（用户需要能自查备份是否完整） */
  function countIn(str, paths) {
    var o;
    try { o = JSON.parse(str); } catch (e) { return 0; }
    if (Array.isArray(o)) return o.length;
    if (!o || typeof o !== 'object') return 0;
    for (var i = 0; i < paths.length; i++) {
      var v = o[paths[i]];
      if (Array.isArray(v)) return v.length;
      if (typeof v === 'number' && v > 0) return v;
    }
    return 0;
  }
  function contentSummary(map) {
    map = map || {};
    var parts = [];
    var orders = countIn(map['dailyPerformance_orders'], ['orders']);
    if (orders) parts.push('订单 ' + orders.toLocaleString() + ' 条');
    var hist = countIn(map['dailyPerformance_hist'], ['list', 'hist', 'rows']);
    if (hist) parts.push('历史明细 ' + hist.toLocaleString() + ' 条');
    else if (map['dailyPerformance_hist']) parts.push('历史明细 已存');
    var vab = countIn(map['dailyPerformance_vab'], ['list', 'rows']);
    if (vab) parts.push('VAB ' + vab + ' 人');
    var mv = countIn(map['med_vab_list'], []);
    if (mv) parts.push('医管VAB ' + mv + ' 人');
    var mo = countIn(map['med_order_list'], []);
    if (mo) parts.push('医管订单 ' + mo.toLocaleString() + ' 条');
    var mc = countIn(map['med_charge_list'], []);
    if (mc) parts.push('收费 ' + mc.toLocaleString() + ' 条');
    var tgt = 0;
    try { tgt = Object.keys(JSON.parse(map['monthly_mgr_targets_v1'] || '{}')).length; } catch (e) {}
    if (tgt) parts.push('月目标 ' + tgt + ' 个月');
    var meida = 0;
    for (var k in map) if (Object.prototype.hasOwnProperty.call(map, k) && k.indexOf('meidaStore_') === 0) meida++;
    if (meida) parts.push('美大店铺 ' + meida + ' 项');
    return parts.length ? parts.join(' · ') : '尚未导入任何业务数据';
  }

  function refreshStats() {
    var ls = readLocal();
    ui.lsn.textContent = Object.keys(ls.map).length + ' 个';
    ui.lsb.textContent = fmtBytes(ls.bytes);
    var ts = 0;
    try { ts = Number(localStorage.getItem('dailyPerformance_snap_ts') || 0); } catch (e) {}
    ui.ts.textContent = ts ? isoLocal(new Date(ts)).slice(5, 16) : '—';
    ui.ts.style.fontSize = ts ? '13px' : '';
    if (ui.content) ui.content.innerHTML = '<b>本机现有数据：</b>' + contentSummary(ls.map);
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
      '  <div style="margin-top:7px"><b>包含：</b>' + contentSummary(p.localStorage) + '</div>',
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

  /* 恢复前冻结子页回写（2026-11-05 第二次实测「恢复后仍是旧数据」的直接原因之二）
     子页在 beforeunload / pagehide / visibilitychange 时会把内存里的旧数据刷回
     localStorage + IndexedDB。恢复写盘到页面刷新之间的几百毫秒里，只要发生一次回写，
     刚恢复的数据就被旧值盖回去——与云端无关，管理员本地模式同样会中招。
     做法：恢复一开始就给所有 iframe 打 __dpFrozen 标记（同源直接写变量，跨域再补一条消息），
     刷新前把 iframe 从 DOM 移除（移除 iframe 只触发 unload，不触发 beforeunload）。 */
  function freezeFrames() {
    var fs = document.querySelectorAll('iframe');
    for (var i = 0; i < fs.length; i++) {
      var w = null;
      try { w = fs[i].contentWindow; } catch (e) {}
      if (!w) continue;
      try { w.__dpFrozen = true; } catch (e) {}
      try { w.postMessage({ type: 'dp:freeze' }, '*'); } catch (e) {}
    }
    return fs.length;
  }
  function detachFrames() {
    var fs = document.querySelectorAll('iframe');
    for (var i = 0; i < fs.length; i++) {
      try { fs[i].parentNode.removeChild(fs[i]); } catch (e) {}
    }
  }

  function doApply(btn) {
    if (!pending) return;
    btn.disabled = true;
    freezeFrames();          // 立刻冻结：后续 apply 可能耗时数秒，期间不许子页回写
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
        if (report.tsAligned) {
          log('已把 ' + report.tsAligned + ' 个模块的「本地更新时间」对齐到当前，避免刷新后被云端旧快照覆盖。', 'ok');
        }
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
          dropped: report.dropped,
          lsFail: report.lsFail.length,
          idbFail: report.idbFail.length,
          tsAligned: report.tsAligned || 0
        }));
      } catch (e) {}
      freezeFrames();        // 二次冻结（apply 期间可能有 iframe 重载丢失标记）
      logKeep('数据已写入本机，1 秒后自动刷新页面…', 'ok');
      // 尽快刷新：先把 iframe 摘掉再刷新，彻底断掉 beforeunload 把旧数据写回的路径。
      // 反馈由刷新后的顶部绿色提示承接。
      setTimeout(function () {
        detachFrames();
        try { location.reload(); } catch (e) {}
      }, 900);
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
    var warn = (info.lsFail > 0)
      ? '<br><span style="color:#b45309">⚠️ 有 ' + info.lsFail + ' 个大数据键因存储空间不足未能写入（旧值已清空），页面会改从数据库读取，功能不受影响。</span>'
      : '';
    if (info.idbFail > 0) {
      warn += '<br><span style="color:#b91c1c">⚠️ 有 ' + info.idbFail + ' 张数据表写入失败，该部分可能仍是旧数据，请再次导入备份。</span>';
    }
    box.innerHTML = '✅ 备份已恢复生效<br>' +
      '<span style="font-size:12px;line-height:1.7;display:inline-block;margin-top:3px">' +
      (info.lsOk || 0) + ' 个存储键 · ' + (info.idbOk || 0).toLocaleString() + ' 项数据表' +
      (info.at ? ' · 备份时间 ' + info.at : '') +
      '<br>已阻止云端旧快照回灌，请重新打开需要的功能模块查看' + warn + '</span>';
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
