'use strict';

/* ============================================================
 * 数据统计平台 · 前端
 * - 多标签页：每个子功能页独立 iframe 保活，切换不丢状态，可同时打开多个
 * - 侧边栏可收起 / 展开（桌面端）
 * 子应用以 iframe 原样运行，框架不侵入其逻辑。
 * ============================================================ */

const state = { modules: [], tabs: [], current: null };

// 发布版本号：由 index.html 顶部的 <script>window.SITE_VER=... 注入，作为唯一权威来源，避免双文件版本漂移。
// 此处保留兜底常量，仅在脱离 Shell 独立打开时生效。用于 iframe 缓存穿透，确保线上改动即时可见。
// file:// 双击打开时不加查询串（带查询串的 iframe 在 file:// 下可能无法加载）。
const SITE_VER = (typeof window !== 'undefined' && window.SITE_VER) || '20261105f';
// 每次页面加载的唯一戳：拼到所有子页 iframe URL 之后，确保浏览器每次都重新拉取最新脚本，
// 彻底规避「服务器无 Cache-Control 时浏览器启发式缓存旧脚本导致刷新丢数据」的问题。
// file:// 双击打开时不加查询串。
const PAGE_NONCE = Date.now();
function withVer(src) {
  if (location.protocol === 'file:') return src;
  const sep = src.indexOf('?') >= 0 ? '&' : '?';
  return src + sep + 'v=' + SITE_VER + '_' + PAGE_NONCE;
}

// 模块注册表（与 server.js 的 MODULES 保持一致；内嵌以便纯静态部署，无需 /api/modules）
const MODULES = [
  {
    id: 'home',
    name: '首页',
    icon: '🏠',
    group: '工作台',
    type: 'embed',
    src: '/pages/home/index.html',
    desc: '平台工作台首页（本周数据/快捷入口/平台动态）'
  },
  {
    id: 'data-statistics',
    name: '客量业绩统计报表',
    icon: '📊',
    group: '数据统计',
    type: 'embed',
    src: '/pages/data-statistics/index.html',
    desc: '客量 / 业绩统计（周客量、周业绩、管家业绩、预约量）',
    children: [
      { id: 'ds-report', name: '周客量业绩统计报表', icon: '📊', mod: 'a' },
      { id: 'ds-appt', name: '预约量统计', icon: '📅', mod: 'b' },
    ],
  },
  {
    // 2026-10-28：升为父级「日业绩」，下挂「日业绩报表」（原初诊版）与「美大店铺日数据」两个子模块
    id: 'daily-performance',
    name: '日业绩',
    icon: '🩺',
    group: '数据统计',
    type: 'embed',
    src: '/pages/daily-performance-chuzhen/index.html',
    desc: '日业绩系列：日业绩报表（订单 / VAB / 月目标 / 导出）、美大店铺日数据',
    children: [
      { id: 'dp-report', name: '日业绩报表', icon: '🩺', src: '/pages/daily-performance-chuzhen/index.html' },
      { id: 'dp-meida', name: '美大店铺日数据', icon: '🏪', src: '/pages/meida-store-daily/index.html' },
    ],
  },
  {
    id: 'monthly-analysis',
    name: '月度分析面板',
    icon: '📆',
    group: '数据统计',
    type: 'embed',
    src: '/pages/monthly-analysis/index.html',
    desc: '导入「订单项目收退款表-明细」类 Excel，按月聚合现款支付 / 合计 / 充值、一级渠道 / 开单科室 / 到访类型维度',
  },
  {
    id: 'medical-assessment',
    name: '医管考核数据统计',
    icon: '🏥',
    group: '数据统计',
    type: 'embed',
    src: '/pages/medical-assessment/index.html',
    desc: '医疗管理考核相关的数据统计（VAB 会员清单、收退款明细、收费明细、统计表）',
    children: [
      { id: 'ma-vab', name: 'VAB会员清单', icon: '📇', mod: 'vab' },
      { id: 'ma-order', name: '领健数据', icon: '💰', mod: 'order' },
      { id: 'ma-charge', name: 'ABC数据', icon: '🧾', mod: 'charge' },
      { id: 'ma-stats', name: '到店率统计', icon: '📊', mod: 'stats' },
      { id: 'ma-projects', name: '重点项目达成情况统计', icon: '🎯', mod: 'projects' },
      { id: 'ma-consumption', name: '累计消费统计', icon: '💳', mod: 'consumption' },
    ],
  },
  {
    id: 'account-roles',
    name: '账号与角色权限',
    icon: '🔐',
    group: '系统管理',
    type: 'embed',
    src: '/pages/account-roles/index.html',
    desc: '管理与维护操作员档案、角色及细粒度权限（菜单可见性 / 导入 / 导出 / 清除 / 编辑）',
  },
];

const $ = (sel) => document.querySelector(sel);

// 子页面（iframe）持久化桥接：把 iframe 内部的订单数据交由稳定的顶层壳窗口保管，
// 避免部署环境下 iframe 处于一次性 / 跨域 origin 导致其内部 localStorage / IndexedDB 刷新后丢失。
// 协议：iframe → 壳：dp:persist（写入）/ dp:request（索取）；壳 → iframe：dp:payload（推送 / 应答）。
// 共享数据（订单 / 历史消费明细 / VAB 名单）在壳这一侧的双层保管：
// localStorage 只有 ~5MB，历史明细整包（数万行）往往远超配额、写入静默失败 → 子页读不到 →
// 「美大店铺日数据」的初诊规则B停用、月累计比日业绩报表偏小。故再落一份到壳窗口的 IndexedDB
//（容量大），postMessage 传输不受体积限制，读取时以 IndexedDB 为准、localStorage 兜底。
const DP_DB = 'dpBridge';
const DP_KEYS = ['dailyPerformance_orders', 'dailyPerformance_hist', 'dailyPerformance_vab'];
const dpMem = new Map();
function dpOpen(){
  return new Promise((res, rej) => {
    try{
      const r = indexedDB.open(DP_DB, 1);
      r.onupgradeneeded = () => { try{ if(!r.result.objectStoreNames.contains('kv')) r.result.createObjectStore('kv'); }catch(_){} };
      r.onsuccess = () => res(r.result);
      r.onerror = () => rej(r.error);
    }catch(e){ rej(e); }
  });
}
function dpPut(key, payload){
  dpMem.set(key, payload);
  dpOpen().then(db => new Promise(res => {
    try{
      const tx = db.transaction('kv', 'readwrite');
      tx.objectStore('kv').put(payload, key);
      tx.oncomplete = () => res(true);
      tx.onerror = () => res(false);
    }catch(e){ res(false); }
  })).catch(() => {});
}
function dpGet(key){
  if(dpMem.has(key)) return Promise.resolve(dpMem.get(key));
  let ls = null;
  try{ ls = localStorage.getItem(key); }catch(_){}
  return dpOpen().then(db => new Promise(res => {
    try{
      const r = db.transaction('kv', 'readonly').objectStore('kv').get(key);
      r.onsuccess = () => res(r.result || ls);
      r.onerror = () => res(ls);
    }catch(e){ res(ls); }
  })).catch(() => ls);
}
function pushEmbeddedState(frame){
  if(!frame.contentWindow) return;
  // 三份共享数据都要推：少推历史明细会让「美大店铺日数据」页的初诊判定规则B停用。
  DP_KEYS.forEach(k => {
    dpGet(k).then(payload => {
      if(payload && frame.contentWindow) frame.contentWindow.postMessage({ type:'dp:payload', key:k, payload }, '*');
    }).catch(() => {});
  });
}
window.addEventListener('message', (e) => {
  const m = (e && e.data) || null;
  if(!m || typeof m !== 'object') return;
  if(m.type === 'dp:persist' && typeof m.key === 'string' && typeof m.payload === 'string'){
    dpPut(m.key, m.payload);
    try { localStorage.setItem(m.key, m.payload); } catch(_){}
  } else if(m.type === 'dp:request' && typeof m.key === 'string'){
    dpGet(m.key).then(payload => {
      try { if(e.source && e.source.postMessage) e.source.postMessage({ type:'dp:payload', key:m.key, payload: payload || null }, '*'); } catch(_){}
    });
  } else if(m.type === 'dp:reset'){
    // 备份恢复已直接改写 localStorage / IndexedDB，内存缓存必须失效，
    // 否则后续 pushEmbeddedState 或 dp:request 应答会把旧数据回灌给子页。
    dpMem.clear();
  } else if(m.type === 'rbac:request'){
    // 子页（iframe）请求当前权限快照：回发 rbac:state 供其守卫禁用无权限操作
    try { if (e.source && e.source.postMessage) e.source.postMessage({ type:'rbac:state', payload: RBAC.snapshotForChild() }, '*'); } catch(_){}
  }
});

/* ---------- 启动 ---------- */
async function boot() {
  startClock();
  setupVerTag();
  setupDrawer();
  setupSidebarToggle();
  state.modules = MODULES;
  if (!state.modules.length) {
    $('#content').innerHTML = '<div class="loading">当前没有可集成的模块。</div>';
    return;
  }
  // 权限引擎就绪后：菜单按角色过滤、权限变更时重渲并广播给各 iframe
  RBAC.onChange(() => { renderNav(); broadcastRbacState(); });
  await RBAC.ready();
  renderNav();
  const hash = location.hash.replace('#', '');
  const startId = resolveTarget(hash) ? hash : firstVisibleId();
  openTarget(startId);
  window.addEventListener('hashchange', () => {
    const id = location.hash.replace('#', '');
    if (resolveTarget(id)) openTarget(id);
  });
}

/* ---------- 模块可见性（RBAC） ---------- */
function canShowModule(m) {
  if (m.id === 'home') return true; // 首页对所有人可见（不参与 RBAC 模块可见性）
  if (typeof RBAC === 'undefined') return true;
  if (m.id === 'account-roles') return RBAC.can('manageAcl'); // 管理入口仅 manageAcl 可见
  return RBAC.canViewModule(m.id);
}
function firstVisibleId() {
  for (const m of state.modules) {
    if (!canShowModule(m)) continue;
    if (m.children && m.children.length) return m.children[0].id;
    return m.id;
  }
  const f = state.modules[0];
  return (f && f.children && f.children.length) ? f.children[0].id : (f ? f.id : '');
}
/* 向所有已打开的 iframe 广播当前权限快照（子页守卫据此禁用无权限操作） */
function broadcastRbacState() {
  if (typeof RBAC === 'undefined') return;
  const snap = RBAC.snapshotForChild();
  state.tabs.forEach((t) => {
    const page = document.getElementById('page-' + t.id);
    const f = page && page.querySelector('iframe');
    if (f && f.contentWindow) { try { f.contentWindow.postMessage({ type: 'rbac:state', payload: snap }, '*'); } catch (_) {} }
  });
}

/* ---------- 侧边栏模块导航（支持二级子菜单） ---------- */
function renderNav() {
  const nav = $('#nav');
  const groups = {};
  state.modules.forEach((m) => { (groups[m.group] ||= []).push(m); });
  let html = '';
  Object.keys(groups).forEach((g) => {
    html += `<div class="nav-group">${g}</div>`;
    groups[g].forEach((m) => {
      if (!canShowModule(m)) return;               // 无权限的模块不在侧边栏出现
      if (m.children && m.children.length) {
        const sub = m.children.map((c) => `
          <div class="nav-subitem" data-id="${c.id}">
            <span class="ic">${c.icon || '•'}</span><span class="label">${c.name}</span>
          </div>`).join('');
        html += `<div class="nav-parent" data-id="${m.id}">
          <span class="ic">${m.icon}</span>
          <span class="label">${m.name}</span>
          <span class="caret">▸</span>
        </div>
        <div class="nav-sub">${sub}</div>`;
      } else {
        html += `<div class="nav-item" data-id="${m.id}">
          <span class="ic">${m.icon}</span><span class="label">${m.name}</span></div>`;
      }
    });
  });
  nav.innerHTML = html;
  nav.querySelectorAll('.nav-item').forEach((el) => el.addEventListener('click', () => openTarget(el.dataset.id)));
  nav.querySelectorAll('.nav-parent').forEach((el) => el.addEventListener('click', () => toggleParent(el.dataset.id)));
  nav.querySelectorAll('.nav-subitem').forEach((el) => el.addEventListener('click', (e) => { e.stopPropagation(); openTarget(el.dataset.id); }));
}

/* ---------- 二级子功能解析与打开 ---------- */
function resolveTarget(id) {
  /* 旧版 hash 兼容：合并模块前的一级菜单 id → 新的子功能 id */
  // 2026-10-23：标准版「日业绩报表」已移除；
  // 2026-10-28：再升为父级「日业绩」，报表本体成为子项 dp-report —— 旧 id 全部落到它。
  if (id === 'dp-std' || id === 'dp-chuzhen' || id === 'daily-performance-chuzhen') id = 'dp-report';
  // 父级 hash（#daily-performance）在 openTarget 里只会展开、不开页 → 会首屏空白、书签失效，
  // 因此这里直接解析为第一个子项（日业绩报表）。
  if (id === 'daily-performance') {
    const p = state.modules.find((m) => m.id === 'daily-performance');
    if (p && p.children && p.children.length) return { kind: 'child', mod: p, child: p.children[0] };
  }
  const mod = state.modules.find((m) => m.id === id);
  if (mod) return { kind: 'module', mod };
  for (const m of state.modules) {
    if (m.children) {
      const c = m.children.find((x) => x.id === id);
      if (c) return { kind: 'child', mod: m, child: c };
    }
  }
  return null;
}

function openTarget(id) {
  const t = resolveTarget(id);
  if (!t) return;
  if (t.kind === 'module') {
    if (t.mod.children && t.mod.children.length) toggleParent(t.mod.id);
    else openModule(t.mod.id);
  } else {
    openChild(t.mod, t.child);
  }
}

function toggleParent(id) {
  const parent = document.querySelector('.nav-parent[data-id="' + id + '"]');
  if (!parent) return;
  const sub = parent.nextElementSibling;
  if (!sub || !sub.classList.contains('nav-sub')) return;
  const open = sub.classList.toggle('open');
  parent.classList.toggle('open', open);
}

function openChild(parent, child) {
  const key = child.id;
  let tab = state.tabs.find((t) => t.moduleId === key);
  if (!tab) {
    tab = { id: 'tab-' + key, moduleId: key, title: child.name, parentName: parent.name, mod: child.mod };
    state.tabs.push(tab);
    createChildTab(tab, parent, child);
  }
  activateTab(tab.id);
}

function createChildTab(tab, parent, child) {
  const page = document.createElement('div');
  page.className = 'tab-page';
  page.id = 'page-' + tab.id;
  const frame = document.createElement('iframe');
  frame.className = 'embed-frame';
  frame.title = parent.name + ' · ' + child.name;
  frame.setAttribute('loading', 'lazy');
  // 注意：不要给 iframe src 加 ?mod= 查询串——在 file:// 双击打开时，带查询串的 iframe 可能无法加载。
  // 改用 postMessage 把子功能（a/b）发给子应用，兼容 file:// 与 http/https 两种打开方式。
  // 独立页面（如月度分析面板，现已作为一级菜单）使用 src 直接嵌入，无需 postMessage。
  frame.src = withVer((child.src || parent.src).replace(/^\//, ''));
  const post = () => { if (!child.mod) return; try { frame.contentWindow.postMessage({ type: 'app:setMod', mod: child.mod }, '*'); } catch (e) {} };
  frame.addEventListener('load', () => { pushEmbeddedState(frame); post(); setTimeout(post, 300); broadcastRbacState(); });
  page.appendChild(frame);
  $('#content').appendChild(page);
}

/* ---------- 多标签页 ---------- */
function openModule(id) {
  const mod = state.modules.find((m) => m.id === id);
  if (!mod) return;
  let tab = state.tabs.find((t) => t.moduleId === id);
  if (!tab) {
    tab = { id: 'tab-' + id, moduleId: id, title: mod.name };
    state.tabs.push(tab);
    createTabPage(tab, mod);
  }
  activateTab(tab.id);
}

function createTabPage(tab, mod) {
  const page = document.createElement('div');
  page.className = 'tab-page';
  page.id = 'page-' + tab.id;
  const frame = document.createElement('iframe');
  frame.className = 'embed-frame';
  frame.title = mod.name;
  frame.setAttribute('loading', 'lazy');
  frame.src = withVer(mod.src.replace(/^\//, '')); // 去掉前导斜杠，兼容 GitHub Pages 子路径；http(s) 下追加版本号防缓存
  frame.addEventListener('load', () => { pushEmbeddedState(frame); broadcastRbacState(); });
  page.appendChild(frame);
  $('#content').appendChild(page);
}

function activateTab(tabId) {
  state.tabs.forEach((t) => {
    const page = document.getElementById('page-' + t.id);
    if (page) page.classList.toggle('active', t.id === tabId);
  });
  state.current = state.tabs.find((t) => t.id === tabId);
  if (!state.current) return;
  location.hash = state.current.moduleId;
  const cur = state.current.moduleId;
  document.querySelectorAll('.nav-item, .nav-subitem').forEach((el) => {
    el.classList.toggle('active', el.dataset.id === cur);
  });
  document.querySelectorAll('.nav-parent').forEach((el) => {
    const sub = el.nextElementSibling;
    if (sub && sub.classList.contains('nav-sub')) {
      const anyActive = [].some.call(sub.children, (c) => c.classList.contains('active'));
      el.classList.toggle('open', anyActive);
      sub.classList.toggle('open', anyActive);
    }
  });
  $('#crumb').textContent = (state.current.parentName ? state.current.parentName + ' / ' : '') + state.current.title;
  renderTabbar();
}

function closeTab(tabId) {
  const idx = state.tabs.findIndex((t) => t.id === tabId);
  if (idx < 0) return;
  const page = document.getElementById('page-' + state.tabs[idx].id);
  if (page) page.remove();
  const closed = state.tabs[idx];
  state.tabs.splice(idx, 1);
  if (state.tabs.length === 0) {
    state.current = null;
    $('#content').innerHTML = '<div class="loading">从左侧选择一个功能模块开始（可同时打开多个）</div>';
    $('#crumb').textContent = '数据统计平台';
    document.querySelectorAll('.nav-item').forEach((el) => el.classList.remove('active'));
    renderTabbar();
    return;
  }
  if (state.current && state.current.id === closed.id) {
    activateTab(state.tabs[Math.max(0, idx - 1)].id);
  } else {
    renderTabbar();
  }
}

function renderTabbar() {
  const bar = $('#tabbar');
  if (!state.tabs.length) { bar.style.display = 'none'; bar.innerHTML = ''; return; }
  bar.style.display = 'flex';
  bar.innerHTML = state.tabs.map((t) => `
    <div class="tab ${state.current && t.id === state.current.id ? 'active' : ''}" data-id="${t.id}">
      <span class="tab-title">${t.title}</span>
      <span class="tab-close" data-close="${t.id}" title="关闭">×</span>
    </div>`).join('');
  bar.querySelectorAll('.tab').forEach((el) => {
    el.addEventListener('click', (e) => {
      const closeId = e.target.dataset.close;
      if (closeId) { e.stopPropagation(); closeTab(closeId); }
      else activateTab(el.dataset.id);
    });
  });
}

/* ---------- 侧边栏收起 / 展开（桌面端） ---------- */
function setupSidebarToggle() {
  const btn = $('#sidebar-toggle');
  const saved = localStorage.getItem('sidebar-collapsed') === '1';
  if (saved) document.body.classList.add('sidebar-collapsed');
  syncToggle(btn);
  btn.addEventListener('click', () => {
    const collapsed = document.body.classList.toggle('sidebar-collapsed');
    localStorage.setItem('sidebar-collapsed', collapsed ? '1' : '0');
    syncToggle(btn);
  });
}
function syncToggle(btn) {
  btn.textContent = document.body.classList.contains('sidebar-collapsed') ? '»' : '«';
}

/* ---------- 手机版抽屉（与电脑端共用同一批模块，功能完全一致） ---------- */
function setupDrawer() {
  const sb = $('.sidebar'), ov = $('#overlay'), btn = $('#hamburger');
  const open = () => { sb.classList.add('open'); ov.classList.add('show'); };
  const close = () => { sb.classList.remove('open'); ov.classList.remove('show'); };
  btn.addEventListener('click', () => {
    sb.classList.contains('open') ? close() : open();
  });
  ov.addEventListener('click', close);
  $('#nav').addEventListener('click', (e) => { if (e.target.closest('.nav-item')) close(); });
  window.addEventListener('resize', () => { if (window.innerWidth > 768) close(); });
}

/* ---------- 时钟 ---------- */
function startClock() {
  const el = $('#clock');
  const tick = () => { el.textContent = new Date().toLocaleString('zh-CN', { hour12: false }); };
  tick(); setInterval(tick, 1000);
}

/* ---------- 顶栏版本号徽章 ---------- */
// 显示当前发布版本（唯一权威来源 = index.html 的 window.SITE_VER），点击复制。
// 作用：用户不用开 F12 就能确认「浏览器里跑的是不是最新一版」——排查缓存问题时第一眼要看的。
function copyText(t) {
  try {
    const ta = document.createElement('textarea');
    ta.value = t;
    ta.setAttribute('readonly', '');
    ta.style.cssText = 'position:fixed;left:-9999px;top:0;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return ok;
  } catch (e) { return false; }
}
function setupVerTag() {
  const el = $('#ver-tag');
  if (!el) return;
  const text = 'v' + SITE_VER;
  el.textContent = text;
  el.title = '当前发布版本 ' + SITE_VER + '（点击复制）';
  let timer = 0;
  const flash = () => {
    el.textContent = '已复制';
    el.classList.add('copied');
    clearTimeout(timer);
    timer = setTimeout(() => { el.textContent = text; el.classList.remove('copied'); }, 1100);
  };
  el.addEventListener('click', () => {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(SITE_VER).then(flash, () => { if (copyText(SITE_VER)) flash(); });
        return;
      }
    } catch (e) {}
    if (copyText(SITE_VER)) flash();
  });
}

boot();
