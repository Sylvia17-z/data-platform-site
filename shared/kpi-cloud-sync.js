/* ================================================================
 * 通用云端快照同步（2026-10-16）
 * 适用：以「整包 JSONB 快照 + 单行 upsert + 按账号隔离」方式持久化的模块。
 *       当前使用者：月度分析面板（dp_monthly）、周客量统计平台（dp_stats）。
 * 策略（与 shared/ma-cloud-sync.js 一致，便于排障复用经验）：
 *   - 登录后首次对齐：本地空 → 采用云端；本地有且比云端新/更全 → 上传云端
 *   - 会话中：页面在数据变更后调用 KpiCloudSync.markDirty()，15s 轮询发现脏标记
 *     即防抖 1.2s 上传（无需每 15s 对大树做完整摘要）
 *   - 单行 upsert：按 updated_at 倒序取最新一条更新，并清理多余残留行
 *   - 未登录 / 云端未就绪 / 管理员本地账号 → 静默跳过（数据仍存本机）
 * 用法（页面内）：
 *   window.KpiCloudSync.start({
 *     table  : 'dp_monthly',
 *     tsKey  : 'kpi_cloud_ts_monthly',            // 本地记录同步时间的 key
 *     collect: function(){ return payload|null; }, // 本地快照（必须含 ts）
 *     apply  : function(payload){ ... },           // 用云端快照恢复本地并重渲染
 *     isEmpty: function(payload){ return bool; },  // 可选：空包判定
 *     countOf: function(payload){ return n; }      // 可选：新旧裁决用的规模量
 *   });
 * ================================================================ */
(function(){
  if(window.__KPI_CLOUD_SYNC__) return; window.__KPI_CLOUD_SYNC__ = true;

  var DEFAULT_PK = 'wbpk_KQXIeDHLXWDHN7WQp3zqUQ_Y5IlM6hJd7n10U8UN5bYmSGUYeevVOwl';
  var POLL_MS = 15000, DEBOUNCE_MS = 1200;

  // 凭证惰性读取：页面可能在脚本之后才设置 window.CLOUD_PK / window.CLOUD_ENDPOINT，
  // 故不在 IIFE 求值时锁定，而是在真正建 client 时再取。
  function pk(){ return (typeof window !== 'undefined' && window.CLOUD_PK) || DEFAULT_PK; }
  function endpoint(){ return (typeof window !== 'undefined' && window.CLOUD_ENDPOINT) || 'https://haibei-medical.app.workbuddy.host'; }

  var cloud = null, opts = null, timer = null, lastHash = null, initDone = false, dirty = false, flushing = false;

  window.__KPI_SYNC_DEBUG = { started:false, initDone:false, adminMode:false, lastError:null, lastAdopt:null, lastPush:null, table:null };

  /* 默认管理员（本地内置假会话）无云端权限：整体跳过，数据仅存本机 */
  function isAdminSession(){
    try{
      var s = JSON.parse(localStorage.getItem('site-admin-session') || 'null');
      return !!(s && s.user && Number(s.expiresAt) > Date.now());
    }catch(e){ return false; }
  }
  function getCloud(){
    if(cloud) return cloud;
    if(typeof window.WorkBuddyCloud === 'undefined') return null;
    try{ cloud = window.WorkBuddyCloud.createWorkBuddyCloud({ endpoint: endpoint(), publishableKey: pk() }); }catch(e){ cloud = null; }
    return cloud;
  }
  function tsGet(){ try{ return Number(localStorage.getItem(opts.tsKey)) || 0; }catch(e){ return 0; } }
  function tsSet(v){ try{ localStorage.setItem(opts.tsKey, String(v)); }catch(e){} }
  function isEmpty(p){ return opts.isEmpty ? !!opts.isEmpty(p) : (p == null); }
  function countOf(p){ return opts.countOf ? (Number(opts.countOf(p)) || 0) : 0; }
  function hashOf(o){
    if(o == null) return 'empty';
    // 排除易变的写入时间戳 ts：否则每次 collect() 都产生新哈希 → 无变更也反复上传
    var c = o;
    if(typeof o === 'object'){
      c = {}; for(var k in o){ if(k !== 'ts') c[k] = o[k]; }
    }
    var s = JSON.stringify(c);
    if(!s || s === 'null' || s === '{}') return 'empty';
    var len = s.length;
    var sample = s.slice(0, 200) + '|' + s.slice(len >> 1, (len >> 1) + 200) + '|' + s.slice(-200);
    var h = 0; for(var i = 0; i < sample.length; i++){ h = ((h * 31) + sample.charCodeAt(i)) >>> 0; }
    return len + ':' + h.toString(36);
  }

  async function fetchCloudRow(){
    var c = getCloud(); if(!c) return null;
    try{
      var r = await c.database.from(opts.table).select('id, payload, updated_at').order('updated_at', { ascending: false }).limit(1);
      if(r.error) return null;
      return (Array.isArray(r.data) ? r.data[0] : r.data) || null;
    }catch(e){ return null; }
  }
  async function pushCloud(payload, force){
    var c = getCloud(); if(!c) return false;
    try{
      var sel = await c.database.from(opts.table).select('id, payload, updated_at').order('updated_at', { ascending: false }).limit(1);
      if(sel.error) return false;
      var latest = Array.isArray(sel.data) && sel.data.length ? sel.data[0] : null;
      var body = { payload: payload, updated_at: new Date().toISOString() };
      if(latest && latest.id != null){
        var cloudTs = Number(latest.payload && latest.payload.ts) || 0;
        if(!force && cloudTs > tsGet()) return false;   // 其他设备刚写入且更新 → 跳过本次上传
        var up = await c.database.from(opts.table).update(body).eq('id', latest.id);
        if(up.error) return false;
        // 清理历史残留行，保证下次读取只命中最新一条
        try{
          var others = await c.database.from(opts.table).select('id');
          if(Array.isArray(others.data)){
            for(var i = 0; i < others.data.length; i++){
              var rr = others.data[i];
              if(rr && rr.id != null && rr.id !== latest.id){ try{ await c.database.from(opts.table).delete().eq('id', rr.id); }catch(_){} }
            }
          }
        }catch(_){}
      }else{
        var ins = await c.database.from(opts.table).insert(body);
        if(ins.error) return false;
      }
      tsSet(Number(payload.ts) || Date.now());
      return true;
    }catch(e){ return false; }
  }

  async function initialSync(){
    if(initDone) return; initDone = true;
    window.__KPI_SYNC_DEBUG.initDone = true;
    try{
      var local = await opts.collect();
      var localEmpty = isEmpty(local);
      var row = await fetchCloudRow();
      var cp = row && row.payload;
      if(cp && !isEmpty(cp)){
        var cloudTs = Number(cp.ts) || 0;
        var adopt = false;
        if(localEmpty) adopt = true;                       // 本地空 → 直接采用云端
        else if(cloudTs > tsGet()) adopt = true;           // 云端更新（其他设备写入）
        else if(countOf(cp) > countOf(local)) adopt = true; // 云端更完整
        if(adopt){
          window.__KPI_SYNC_DEBUG.lastAdopt = { table: opts.table, cloudTs: cloudTs };
          await opts.apply(cp);
          tsSet(cloudTs || Date.now());
          lastHash = hashOf(await opts.collect());
          return;
        }
      }
      if(localEmpty){ lastHash = hashOf(local); return; }   // 双方皆空：不写空包
      var ok = await pushCloud(local, true);
      window.__KPI_SYNC_DEBUG.lastPush = ok;
      lastHash = hashOf(local);
    }catch(e){ window.__KPI_SYNC_DEBUG.lastError = String((e && e.message) || e); }
  }

  async function flush(){
    if(flushing) return; flushing = true;
    try{
      var cur = await opts.collect();
      if(isEmpty(cur)) return;                       // 不用空包覆盖云端（主动清空请走页面自身的清空逻辑）
      var h = hashOf(cur);
      if(h === lastHash){ dirty = false; return; }
      var ok = await pushCloud(cur, false);
      if(ok){ lastHash = h; dirty = false; window.__KPI_SYNC_DEBUG.lastPush = true; }
    }catch(e){ window.__KPI_SYNC_DEBUG.lastError = String((e && e.message) || e); }
    finally { flushing = false; }
  }
  function schedule(){
    if(timer) clearTimeout(timer);
    timer = setTimeout(function(){ timer = null; flush(); }, DEBOUNCE_MS);
  }
  function pollTick(){
    // 兜底：页面若漏标（或标记发生在首次对齐之前），这里补一次
    if(initDone && dirty && lastHash !== null){ dirty = false; schedule(); }
  }

  function begin(){
    window.__KPI_SYNC_DEBUG.started = true;
    initialSync();
    setInterval(pollTick, POLL_MS);
  }

  window.KpiCloudSync = {
    start: function(o){
      if(opts) return; opts = o || {};
      window.__KPI_SYNC_DEBUG.table = opts.table || null;
      if(!opts.table || typeof opts.collect !== 'function' || typeof opts.apply !== 'function') return;
      if(isAdminSession()){ window.__KPI_SYNC_DEBUG.adminMode = true; return; }   // 管理员：本地模式
      var fired = false, tries = 0;
      function fire(){ if(fired) return; fired = true; begin(); }
      if(window.addEventListener){
        window.addEventListener('site-auth-ready', fire);
        // 兜底：脚本先于事件执行 / 页面独立打开（无 auth-gate）时，自行探测会话
        var probe = async function(){
          if(fired || tries++ > 25) return;
          var c = getCloud();
          if(c){ try{ var r = await c.auth.getSession(); if(r.data && r.data.user) return fire(); }catch(e){} }
          setTimeout(probe, 800);
        };
        setTimeout(probe, 1000);
      }else{ fire(); }
    },
    /* 页面在本地数据变更后调用。首次对齐完成（lastHash 就绪）后即刻防抖上传，
       不必等 15s 轮询，避免用户导入后立刻关页导致数据没上云。 */
    markDirty: function(){
      dirty = true;
      if(initDone && lastHash !== null) schedule();
    }
  };
})();
