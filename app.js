'use strict';

/* ============================================================
 * 数据统计平台 · 前端
 * - 多标签页：每个子功能页独立 iframe 保活，切换不丢状态，可同时打开多个
 * - 侧边栏可收起 / 展开（桌面端）
 * 子应用以 iframe 原样运行，框架不侵入其逻辑。
 * ============================================================ */

const state = { modules: [], tabs: [], current: null };

// 模块注册表（与 server.js 的 MODULES 保持一致；内嵌以便纯静态部署，无需 /api/modules）
const MODULES = [
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
    id: 'daily-performance',
    name: '每日业绩日报表',
    icon: '📈',
    group: '数据统计',
    type: 'embed',
    src: '/pages/daily-performance/index.html',
    desc: '各部门业绩完成情况跟踪日报表（订单 / VAB / 月目标 / 导出）',
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
    desc: '医疗管理考核相关的数据统计（VAB 会员清单、收退款明细、收费明细三类报表）',
    children: [
      { id: 'ma-vab', name: 'VAB会员清单', icon: '📇', mod: 'vab' },
      { id: 'ma-order', name: '收退款数据明细', icon: '💰', mod: 'order' },
      { id: 'ma-charge', name: '收费明细数据明细', icon: '🧾', mod: 'charge' },
    ],
  },
];

const $ = (sel) => document.querySelector(sel);

/* ---------- 启动 ---------- */
async function boot() {
  startClock();
  setupDrawer();
  setupSidebarToggle();
  state.modules = MODULES;
  if (!state.modules.length) {
    $('#content').innerHTML = '<div class="loading">当前没有可集成的模块。</div>';
    return;
  }
  renderNav();
  const hash = location.hash.replace('#', '');
  const first = state.modules[0];
  const startId = resolveTarget(hash) ? hash
    : (first.children && first.children.length ? first.children[0].id : first.id);
  openTarget(startId);
  window.addEventListener('hashchange', () => {
    const id = location.hash.replace('#', '');
    if (resolveTarget(id)) openTarget(id);
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
  frame.src = (child.src || parent.src).replace(/^\//, '');
  const post = () => { if (!child.mod) return; try { frame.contentWindow.postMessage({ type: 'app:setMod', mod: child.mod }, '*'); } catch (e) {} };
  frame.addEventListener('load', () => { post(); setTimeout(post, 300); });
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
  frame.src = mod.src.replace(/^\//, ''); // 去掉前导斜杠，兼容 GitHub Pages 子路径
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

boot();
