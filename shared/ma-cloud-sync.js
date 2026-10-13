/* ================================================================
 * 医管考核页 云端同步（2026-10-13 需求：登录后数据多端同步）
 * 策略：轮询 + 内容哈希，不侵入页面内部代码
 * - 本地数据源：localStorage med_vab_list / med_order_list / med_charge_list
 *   + IndexedDB(med_assess_db/kv: order, charge)（大表镜像，本地权威）
 * - 云端：表 dp_medassess（RLS owner 隔离，单行/账号 upsert 最新）
 * - 加载/登录后：云端有数据且本地为空 → 写回本地 + 刷新一次恢复
 *                 本地有数据 → 上传云端
 * - 会话中：每 15s 检查本地哈希，变化即防抖上传（含主动清空→上传空包）
 * ================================================================ */
(function(){
  if(window.__MA_CLOUD_SYNC__) return; window.__MA_CLOUD_SYNC__ = true;

  var PK = (window.CLOUD_PK) || 'wbpk_KQXIeDHLXWDHN7WQp3zqUQ_Y5IlM6hJd7n10U8UN5bYmSGUYeevVOwl';
  var ENDPOINT = (window.CLOUD_ENDPOINT) || (typeof location !== 'undefined' ? location.origin : '');
  var TABLE = 'dp_medassess';
  var LS_VAB = 'med_vab_list', LS_ORDER = 'med_order_list', LS_CHARGE = 'med_charge_list';
  var IDB_DB = 'med_assess_db', IDB_ORDER = 'order', IDB_CHARGE = 'charge';
  var LS_TS = 'med_cloud_ts';           // 本地最后已知云端/本地同步时间戳
  var POLL_MS = 15000, DEBOUNCE_MS = 1500;

  var cloud = null, lastPushHash = null, lastCloudTs = 0, _timer = null, _reloading = false;
  // 可观测性：E2E/排障用
  window.__MA_SYNC_DEBUG = { started: false, initDone: false, lastError: null, lastAdopt: null, lastPush: null };

  function getCloud(){
    if(cloud) return cloud;
    if(typeof window.WorkBuddyCloud === 'undefined') return null;
    try{ cloud = window.WorkBuddyCloud.createWorkBuddyCloud({ endpoint: ENDPOINT, publishableKey: PK }); }catch(e){ cloud = null; }
    return cloud;
  }
  function lsGet(k){ try{ return localStorage.getItem(k); }catch(e){ return null; } }
  function lsSet(k,v){ try{ localStorage.setItem(k,v); return true; }catch(e){ return false; } }

  function idbOpen(){
    return new Promise(function(res, rej){
      try{
        if(typeof indexedDB === 'undefined' || !indexedDB) return rej(new Error('no idb'));
        var req = indexedDB.open(IDB_DB, 1);
        req.onupgradeneeded = function(){ try{ req.result.createObjectStore('kv'); }catch(_){} };
        req.onsuccess = function(){ res(req.result); };
        req.onerror = function(){ rej(req.error || new Error('idb open failed')); };
      }catch(e){ rej(e); }
    });
  }
  function idbGet(key){
    return idbOpen().then(function(db){ return new Promise(function(res, rej){
      try{
        var tx = db.transaction('kv','readonly'); var rq = tx.objectStore('kv').get(key);
        rq.onsuccess = function(){ res(rq.result); }; rq.onerror = function(){ rej(rq.error); };
      }catch(e){ rej(e); }
    }); }).catch(function(){ return null; });
  }
  function idbSet(key, val){
    return idbOpen().then(function(db){ return new Promise(function(res, rej){
      try{
        var tx = db.transaction('kv','readwrite');
        if(val == null) tx.objectStore('kv').delete(key); else tx.objectStore('kv').put(val, key);
        tx.oncomplete = function(){ res(); }; tx.onerror = tx.onabort = function(){ rej(tx.error); };
      }catch(e){ rej(e); }
    }); }).catch(function(){});
  }

  function parseSafe(s){ try{ var v = JSON.parse(s); return Array.isArray(v) ? v : null; }catch(e){ return null; } }
  function hashOf(o){
    var s = JSON.stringify(o || null);
    if(!s || s === 'null') return 'empty';
    // 轻量哈希：长度 + 首中尾采样，避免每 15s 对数万行做完整摘要
    var len = s.length;
    var sample = s.slice(0, 200) + '|' + s.slice(len >> 1, (len >> 1) + 200) + '|' + s.slice(-200);
    var h = 0; for(var i = 0; i < sample.length; i++){ h = ((h * 31) + sample.charCodeAt(i)) >>> 0; }
    return len + ':' + h.toString(36);
  }

  async function collectLocal(){
    var vab = parseSafe(lsGet(LS_VAB)) || [];
    var order = await idbGet(IDB_ORDER);  if(order == null) order = parseSafe(lsGet(LS_ORDER));
    var charge = await idbGet(IDB_CHARGE); if(charge == null) charge = parseSafe(lsGet(LS_CHARGE));
    order = Array.isArray(order) ? order : [];
    charge = Array.isArray(charge) ? charge : [];
    var projMap = null;
    try{ var raw = lsGet('med_proj_map'); if(raw) projMap = JSON.parse(raw); }catch(e){}
    return { vab: vab, order: order, charge: charge, projMap: projMap,
      counts: { vab: vab.length, order: order.length, charge: charge.length } };
  }
  function hashPayload(p){
    return hashOf(p.vab) + '|' + hashOf(p.order) + '|' + hashOf(p.charge) + '|' + hashOf(p.projMap);
  }
  function isEmptyPayload(p){ return !p.vab.length && !p.order.length && !p.charge.length && !p.projMap; }

  /* ---------- 云端读 / 写 ---------- */
  async function fetchCloudRow(){
    var c = getCloud(); if(!c) return null;
    try{
      var r = await c.database.from(TABLE).select('id, payload, updated_at').order('updated_at', { ascending: false }).limit(1);
      if(r.error) return null;
      var row = Array.isArray(r.data) ? r.data[0] : r.data;
      return row || null;
    }catch(e){ return null; }
  }
  async function pushCloud(payload, force){
    var c = getCloud(); if(!c) return false;
    try{
      var sel = await c.database.from(TABLE).select('id, payload, updated_at').order('updated_at', { ascending: false }).limit(1);
      if(sel.error) return false;
      var latest = Array.isArray(sel.data) && sel.data.length ? sel.data[0] : null;
      var now = Date.now();
      var body = { payload: payload, updated_at: new Date().toISOString() };
      if(latest && latest.id != null){
        var cloudTs = Number(latest.payload && latest.payload.ts) || 0;
        if(!force && cloudTs > (Number(localStorage.getItem(LS_TS)) || 0) && hashPayload(latest.payload || {}) !== hashPayload(payload)){
          return false;   // 其他设备刚写入且比本地新 → 跳过（下次拉取会对齐）
        }
        var up = await c.database.from(TABLE).update(body).eq('id', latest.id);
        if(up.error) return false;
        // 清理多余残留行
        try{
          var others = await c.database.from(TABLE).select('id');
          if(Array.isArray(others.data)){
            for(var i = 0; i < others.data.length; i++){
              var rr = others.data[i];
              if(rr && rr.id != null && rr.id !== latest.id){ try{ await c.database.from(TABLE).delete().eq('id', rr.id); }catch(_){} }
            }
          }
        }catch(_){}
      }else{
        var ins = await c.database.from(TABLE).insert(body);
        if(ins.error) return false;
      }
      lastCloudTs = now;
      try{ localStorage.setItem(LS_TS, String(now)); }catch(e){}
      return true;
    }catch(e){ return false; }
  }

  /* ---------- 首次对齐 ---------- */
  var _initDone = false;
  async function initialSync(){
    if(_initDone) return; _initDone = true;
    window.__MA_SYNC_DEBUG.initDone = true;
    try{
    var local = await collectLocal();
    var localEmpty = isEmptyPayload(local);
    var row = await fetchCloudRow();
    var cp = row && row.payload;
    if(cp && !isEmptyPayload({ vab: cp.vab || [], order: cp.order || [], charge: cp.charge || [], projMap: cp.projMap || null })){
      var cloudTs = Number(cp.ts) || 0;
      var localTs = Number(localStorage.getItem(LS_TS)) || 0;
      var cloudCounts = { v: (cp.vab || []).length, o: (cp.order || []).length, c: (cp.charge || []).length };
      var adopt = false;
      if(localEmpty) adopt = true;                                   // 本地空 → 直接采用云端
      else if(cloudTs > localTs) adopt = true;                       // 云端更新（其他设备写入）
      else if(cloudCounts.v + cloudCounts.o + cloudCounts.c > local.counts.vab + local.counts.order + local.counts.charge) adopt = true;
      if(adopt && !_reloading){
        window.__MA_SYNC_DEBUG.lastAdopt = { cloudTs: cloudTs, counts: cloudCounts };
        // 写回本地：vab/projMap → localStorage；order/charge → LS + IDB；然后刷新一次由页面自身恢复逻辑接管
        lsSet(LS_VAB, JSON.stringify(cp.vab || []));
        if(cp.projMap != null) lsSet('med_proj_map', JSON.stringify(cp.projMap));
        lsSet(LS_TS, String(cloudTs || Date.now()));
        await idbSet(IDB_ORDER, (cp.order && cp.order.length) ? cp.order : null);
        await idbSet(IDB_CHARGE, (cp.charge && cp.charge.length) ? cp.charge : null);
        try{ localStorage.setItem(LS_ORDER, JSON.stringify(cp.order || [])); }catch(e){}
        try{ localStorage.setItem(LS_CHARGE, JSON.stringify(cp.charge || [])); }catch(e){}
        _reloading = true;
        setTimeout(function(){ location.reload(); }, 300);
        return;
      }
    }
    // 云端空 且 本地有数据 → 首次上传
    if(localEmpty){
      lastPushHash = hashPayload(local);
      return;
    }
    var ok = await pushCloud(Object.assign({ ts: Date.now() }, local), true);
    window.__MA_SYNC_DEBUG.lastPush = ok;
    lastPushHash = hashPayload(local);
    }catch(e){ window.__MA_SYNC_DEBUG.lastError = String(e && e.message || e); }
  }

  /* ---------- 轮询上传（本地变化 → 防抖上传；含主动清空） ---------- */
  async function pollTick(){
    if(_reloading) return;
    var local = await collectLocal();
    var h = hashPayload(local);
    if(h === lastPushHash) return;
    if(_timer){ clearTimeout(_timer); }
    _timer = setTimeout(async function(){
      _timer = null;
      var cur = await collectLocal();
      var h2 = hashPayload(cur);
      if(h2 === lastPushHash) return;
      var wasEmpty = lastPushHash !== null && lastPushHash.split('|').every(function(x){ return x === 'empty'; });
      var nowEmpty = isEmptyPayload(cur);
      // 从有数据变空 = 用户主动清空 → 允许上传空包覆盖云端；初始即空不写云端
      if(nowEmpty && !wasEmpty && lastPushHash === null) return;
      var ok = await pushCloud(Object.assign({ ts: Date.now() }, cur), nowEmpty);
      if(ok) lastPushHash = h2;
    }, DEBOUNCE_MS);
  }

  function start(){
    window.__MA_SYNC_DEBUG.started = true;
    initialSync();
    setInterval(pollTick, POLL_MS);
    // 本地写入后尽快感知（storage 事件仅跨标签触发，本页轮询兜底）
    ['storage'].forEach(function(ev){ window.addEventListener(ev, function(e){
      if(e.key === LS_VAB || e.key === LS_ORDER || e.key === LS_CHARGE) pollTick();
    }); });
  }

  function boot(){
    if(window.__SITE_AUTH_GATE__ && window.addEventListener){
      window.addEventListener('site-auth-ready', function(){ start(); }, { once: true });
      // 若 gate 已解锁（事件早于本脚本），通过会话兜底检测
      var probe = async function(){
        if(_initDone) return;
        var c = getCloud(); if(!c){ setTimeout(probe, 800); return; }
        try{ var r = await c.auth.getSession(); if(r.data && r.data.user) start(); }catch(e){}
      };
      setTimeout(probe, 1200);
    } else {
      setTimeout(function(){ boot(); }, 500);
    }
  }
  boot();
})();
