/* ================================================================
 * 全站权限引擎（RBAC）+ 子页操作守卫（2026-10-16）
 * - 单文件、幂等（window.__RBAC__ 守卫）；壳层与子页 iframe 共引同一份。
 * - 权限模型：自定义角色 + 细粒度权限点（菜单可见性 / 导入 / 导出 / 清除 / 编辑 / 重算 / 管理）。
 * - 管理对象：本机操作员档案（local primary）+ 当前云登录账号映射到角色（cloud fallback）。
 * - 数据归属：全组织共享 —— 采用「共享云账号下的单一 dp_acl 快照行」落地（与 dp_monthly 同 RLS 隔离）。
 *   纯前端无法跨不同云账号真正共享 ACL，属架构固有限制；共用账号组织方为推荐形态。
 * - 混合登录模式：本机操作员档案优先，云快照兜底；首次无配置自动种入全权限超级管理员。
 *
 * 用法：
 *   壳层 public/app.js 启动时 await RBAC.ready() 后渲染菜单并广播 rbac:state 给各 iframe。
 *   子页只需引入本脚本 + 在敏感按钮加 data-rbac="import|export|clear|editVab|editTargets|recompute"，
 *   守卫自动禁用无权限元素（含 <label for=file> 导入触发器双重禁用）。
 * ================================================================ */
(function () {
  if (window.__RBAC__) return;
  window.__RBAC__ = true;

  /* ---------- 常量 ---------- */
  var PK = (window.CLOUD_PK) || 'wbpk_KQXIeDHLXWDHN7WQp3zqUQ_Y5IlM6hJd7n10U8UN5bYmSGUYeevVOwl';
  var ENDPOINT = (window.CLOUD_ENDPOINT) || (typeof location !== 'undefined' ? location.origin : '');
  var ADMIN_KEY = 'site-admin-session';     // 本地内置超级管理员（admin/123）
  var ACL_TABLE = 'dp_acl';                 // 云表：单行 ACL 快照
  var ACL_CACHE_KEY = 'acl_cache';          // 云不可用时本机兜底缓存
  var CURRENT_OP_KEY = 'acl_current_operator'; // 本机当前物理操作员 id

  /* ---------- 静态权限目录 ---------- */
  var CATALOG = {
    actions: [
      { key: 'import',      label: '导入数据' },
      { key: 'export',      label: '导出报表' },
      { key: 'clear',       label: '清除数据' },
      { key: 'editVab',     label: '编辑VAB会员' },
      { key: 'editTargets', label: '编辑月目标' },
      { key: 'recompute',   label: '重新计算' },
      { key: 'manageAcl',   label: '管理账号与角色' }
    ],
    modules: [
      { id: 'data-statistics',    name: '客量业绩统计报表' },
      { id: 'daily-performance',  name: '日业绩统计' },
      { id: 'monthly-analysis',   name: '月度分析面板' },
      { id: 'medical-assessment', name: '医管考核数据统计' },
      { id: 'account-roles',      name: '账号与角色权限' }
    ]
  };

  /* ---------- 内部状态 ---------- */
  var S = {
    ready: false, admin: false, cloudAcl: false,
    role: null, operator: null, acl: null,
    pendingSelect: false, remote: null,
    uid: null, email: null, _cloud: null, _rowId: null
  };
  var readyCbs = [], changeCbs = [];
  var resolved = false, initStarted = false;

  /* ---------- 工具 ---------- */
  function isAdminSession() {
    try {
      var s = JSON.parse(localStorage.getItem(ADMIN_KEY) || 'null');
      return !!(s && s.user && Number(s.expiresAt) > Date.now());
    } catch (e) { return false; }
  }
  function getCloud() {
    if (S._cloud) return S._cloud;
    if (typeof window.WorkBuddyCloud === 'undefined') return null;
    try { S._cloud = window.WorkBuddyCloud.createWorkBuddyCloud({ endpoint: ENDPOINT, publishableKey: PK }); }
    catch (e) { S._cloud = null; }
    return S._cloud;
  }
  function fullPerms() { var a = {}; CATALOG.actions.forEach(function (x) { a[x.key] = true; }); return a; }
  function allModules() { return CATALOG.modules.map(function (m) { return m.id; }); }
  function defaultAcl(uid, name) {
    return {
      ts: Date.now(),
      roles: [{ id: 'r_admin', name: '超级管理员', permissions: { actions: fullPerms(), modules: allModules() } }],
      operators: [{ id: 'op_owner', name: name || '管理员', accountId: uid, roleId: 'r_admin' }],
      assignments: {}
    };
  }
  function loadLocalAcl() {
    try { return JSON.parse(localStorage.getItem(ACL_CACHE_KEY) || 'null'); } catch (e) { return null; }
  }

  /* ---------- 从 ACL 解析当前操作员/角色 ---------- */
  function resolveFromAcl(acl, uid, email) {
    S.acl = acl || S.acl;
    var list = (S.acl && S.acl.operators) || [];
    // 仅在 uid/email 有效时匹配，防止匿名/未解析会话误绑到 accountId 为空的操作员
    var matches = list.filter(function (o) {
      return (uid && o.accountId === uid) || (email && o.accountId === email);
    });
    var op = null;
    if (matches.length === 1) op = matches[0];
    else if (matches.length > 1) {
      var cur = localStorage.getItem(CURRENT_OP_KEY);
      op = matches.filter(function (o) { return o.id === cur; })[0] || null;
      S.pendingSelect = !op;
    }
    S.operator = op;
    S.role = op ? ((S.acl.roles || []).filter(function (r) { return r.id === op.roleId; })[0] || null) : null;
  }

  /* ---------- 公共判定 ---------- */
  function can(action) {
    if (S.admin) return true;
    if (S.remote) return !!S.remote.actions[action];           // 嵌入态：用壳层下发的快照
    if (S.role && S.role.permissions && S.role.permissions.actions) return !!S.role.permissions.actions[action];
    return false;                                              // 无角色：拒绝操作（菜单仍默认可见，避免锁死）
  }
  function canViewModule(mid) {
    if (mid === 'account-roles') return can('manageAcl');      // 管理入口仅 manageAcl 可见
    if (S.admin) return true;
    if (S.remote) return S.remote.modules.indexOf(mid) >= 0;
    if (!S.role) return true;                                  // 降级：菜单可见，避免锁死
    var ms = (S.role.permissions && S.role.permissions.modules) || [];
    return ms.indexOf(mid) >= 0;
  }
  function currentRole() { return S.role; }
  function currentOperator() { return S.operator; }
  function isAdmin() { return S.admin; }
  function isCloudAcl() { return S.cloudAcl; }

  function snapshotForChild() {
    return {
      isAdmin: S.admin,
      isCloudAcl: S.cloudAcl,
      actions: (S.role && S.role.permissions && S.role.permissions.actions) || (S.admin ? fullPerms() : {}),
      modules: (S.role && S.role.permissions && S.role.permissions.modules) || (S.admin ? allModules() : []),
      operator: S.operator ? { id: S.operator.id, name: S.operator.name } : null
    };
  }

  /* ---------- 读写云表 dp_acl（单行快照，复用 kpi-cloud-sync 模式） ---------- */
  async function loadAclFromCloud(c) {
    try {
      var r = await c.database.from(ACL_TABLE).select('id,payload,updated_at').order('updated_at', { ascending: false }).limit(1);
      if (r.error) return null;
      var d = Array.isArray(r.data) ? r.data[0] : r.data;
      return d || null;
    } catch (e) { return null; }
  }
  async function persistAclToCloud(acl) {
    var c = getCloud(); if (!c) return false;
    var body = { payload: acl, updated_at: new Date().toISOString() };
    try {
      if (S._rowId != null) { await c.database.from(ACL_TABLE).update(body).eq('id', S._rowId); }
      else {
        var r = await c.database.from(ACL_TABLE).insert(body);
        if (r.data && r.data[0]) S._rowId = r.data[0].id;
      }
      return true;
    } catch (e) { return false; }
  }
  var _dirty = false, _flushing = false;
  function markAclDirty() { _dirty = true; if (S.ready && !_flushing) setTimeout(pushNow, 800); }
  async function pushNow() {
    if (_flushing) return; _flushing = true;
    try { if (_dirty && S.acl) await persistAclToCloud(S.acl); } catch (e) {}
    finally { _dirty = false; _flushing = false; }
  }

  /* ---------- 初始化 ---------- */
  function finishInit() {
    resolved = true; S.ready = true;
    readyCbs.forEach(function (cb) { try { cb(); } catch (e) {} });
    readyCbs = [];
  }
  function onReady(cb) { if (resolved) cb(); else readyCbs.push(cb); }
  function onChange(cb) { changeCbs.push(cb); }
  function fireChange() { changeCbs.forEach(function (cb) { try { cb(); } catch (e) {} }); }

  function startInit() {
    initStarted = true;
    if (isAdminSession()) {
      // 本地管理员：全权限，配置仅本机（无真实云端账号）
      S.admin = true; S.cloudAcl = false; S.uid = 'admin'; S.email = 'admin';
      var local = loadLocalAcl();
      S.acl = local || defaultAcl('admin', '管理员');
      resolveFromAcl(S.acl, 'admin', 'admin');
      finishInit();
      return;
    }
    var fired = false;
    /* 无论由 site-auth-ready 事件还是探测触发，都必须先解析会话拿到 uid/email，
       再初始化 ACL —— 否则操作员按 accountId 匹配必然落空（uid=null）。 */
    function go() {
      if (fired) return; fired = true;
      (async function () {
        var c = getCloud();
        if (c) {
          try {
            var r = await c.auth.getSession();
            if (r.data && r.data.user) { S.uid = r.data.user.id; S.email = r.data.user.email || r.data.user.username; }
          } catch (e) {}
        }
        initCloud();
      })();
    }
    if (window.addEventListener) window.addEventListener('site-auth-ready', go);
    // 兜底探测：auth-gate 早于/晚于本脚本时，自行探测会话；始终无会话则降级完成（不阻塞壳层渲染）
    var tries = 0;
    (function probe() {
      if (fired) return;
      if (tries++ > 25) { fired = true; S.cloudAcl = false; finishInit(); return; }
      var c = getCloud();
      if (c) {
        try {
          c.auth.getSession().then(function (r) {
            if (r.data && r.data.user) go();
            else setTimeout(probe, 800);
          }).catch(function () { setTimeout(probe, 800); });
        } catch (e) { setTimeout(probe, 800); }
      } else setTimeout(probe, 800);
    })();
  }
  async function initCloud() {
    var c = getCloud();
    if (!c || !S.uid) { S.cloudAcl = false; finishInit(); return; }   // 无云组件 / 未登录：降级为仅本机
    var row = await loadAclFromCloud(c);
    var acl;
    if (row && row.payload) {
      acl = row.payload; S.cloudAcl = true; S._rowId = row.id;
    } else {
      // 首跑 bootstrap：自动种入全权限超级管理员 + 绑定当前账号操作员（仅在确有登录身份时）
      acl = defaultAcl(S.uid, S.email);
      S.cloudAcl = true;
      try {
        var ins = await c.database.from(ACL_TABLE).insert({ payload: acl, updated_at: new Date().toISOString() });
        if (ins.data && ins.data[0]) S._rowId = ins.data[0].id;
      } catch (e) {}
    }
    resolveFromAcl(acl, S.uid, S.email);
    finishInit();
  }

  function ready() {
    if (resolved) return Promise.resolve();
    return new Promise(function (resolve) {
      onReady(resolve);
      if (!initStarted) startInit();
    });
  }

  /* ---------- 操作员选择 / ACL 读写 ---------- */
  function selectOperator(opId) {
    localStorage.setItem(CURRENT_OP_KEY, opId || '');
    if (S.acl) resolveFromAcl(S.acl, S.uid || 'admin', S.email);
    S.pendingSelect = false;
    fireChange();
  }
  function getAcl() { return S.acl; }
  function saveAcl(acl) {
    acl.ts = Date.now();
    S.acl = acl;
    resolveFromAcl(acl, S.uid || 'admin', S.email);
    if (S.admin) { try { localStorage.setItem(ACL_CACHE_KEY, JSON.stringify(acl)); } catch (e) {} }
    else if (S.cloudAcl) { markAclDirty(); }
    fireChange();
  }

  /* ---------- 子页守卫 ---------- */
  function injectStyle() {
    if (document.getElementById('rbac-style')) return;
    var s = document.createElement('style');
    s.id = 'rbac-style';
    s.textContent = '.rbac-disabled{opacity:.5;cursor:not-allowed;}';
    if (document.head) document.head.appendChild(s);
  }
  function applyGuard(doc) {
    doc = doc || document;
    if (!doc || !doc.querySelectorAll) return;
    // 嵌入态且尚未收到壳层快照：保持原状（降级为开放），避免误将全部按钮判为无权限而全禁
    if ((window.self !== window.top) && !S.remote) return;
    injectStyle();
    var acts = S.remote ? S.remote.actions
      : ((S.role && S.role.permissions && S.role.permissions.actions) || (S.admin ? fullPerms() : {}));
    var allow = function (a) { return a == null ? false : !!acts[a]; };
    var nodes = doc.querySelectorAll('[data-rbac]');
    for (var i = 0; i < nodes.length; i++) {
      var el = nodes[i];
      var key = el.getAttribute('data-rbac');
      var ok = allow(key);
      if (ok) {
        el.classList.remove('rbac-disabled');
        el.removeAttribute('data-rbac-denied');
        el.style.pointerEvents = '';
        el.setAttribute('aria-disabled', 'false');
        if (el.tagName === 'BUTTON') el.disabled = false;
        if (el.tagName === 'INPUT' && el.type === 'file') el.disabled = false;
        if (el.tagName === 'LABEL' && el.getAttribute('for')) {
          var inp = doc.getElementById(el.getAttribute('for'));
          if (inp) { inp.disabled = false; inp.removeAttribute('data-rbac-denied'); }
        }
      } else {
        el.classList.add('rbac-disabled');
        el.setAttribute('data-rbac-denied', '1');
        el.style.pointerEvents = 'none';
        el.setAttribute('aria-disabled', 'true');
        if (el.tagName === 'BUTTON') el.disabled = true;
        if (el.tagName === 'INPUT' && el.type === 'file') el.disabled = true; // 直接禁用 file input，阻断 label / drop 触发
        if (el.tagName === 'LABEL' && el.getAttribute('for')) {
          var inp2 = doc.getElementById(el.getAttribute('for'));
          if (inp2) { inp2.disabled = true; inp2.setAttribute('data-rbac-denied', '1'); } // label 无 disabled，需双重禁用被引用 input
        }
      }
    }
  }
  function initGuardIfEmbedded() {
    if (typeof window === 'undefined') return;
    var embedded = (window.self !== window.top);
    if (!embedded) return;   // 壳层自身不跑守卫
    window.addEventListener('message', function (e) {
      var m = e && e.data;
      if (!m || m.type !== 'rbac:state') return;
      S.remote = m.payload || null;
      applyGuard(document);
    });
    function request() { if (window.parent && window.parent.postMessage) window.parent.postMessage({ type: 'rbac:request' }, '*'); }
    window.addEventListener('DOMContentLoaded', function () { applyGuard(document); request(); });
    if (document.readyState !== 'loading') { applyGuard(document); request(); }
    if ('MutationObserver' in window) {
      // attributes 必须 true：子页映射脚本在 DOMContentLoaded 才补 data-rbac 标签，需触发一次守卫复扫
      var obs = new MutationObserver(function () { applyGuard(document); });
      obs.observe(document.documentElement, { childList: true, subtree: true, attributes: true, attributeFilter: ['data-rbac'] });
    }
  }

  /* ---------- 暴露 API ---------- */
  window.RBAC = {
    permissionCatalog: CATALOG,
    ready: ready,
    can: can,
    canViewModule: canViewModule,
    currentRole: currentRole,
    currentOperator: currentOperator,
    isAdmin: isAdmin,
    isCloudAcl: isCloudAcl,
    snapshotForChild: snapshotForChild,
    selectOperator: selectOperator,
    onReady: onReady,
    onChange: onChange,
    getAcl: getAcl,
    saveAcl: saveAcl,
    applyGuard: applyGuard
  };

  initGuardIfEmbedded();
})();
