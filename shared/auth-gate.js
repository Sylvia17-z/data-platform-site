/* ================================================================
 * 整站共享登录门（2026-10-13 改版：独立登录页方案）
 * - 同源共享 SDK 会话（localStorage），一次登录全站/全部 iframe 生效
 * - 未登录：整页跳转至 /pages/login/index.html?next=<原页面>（独立登录界面）
 *   登录成功后登录页跳回 next 目标页
 * - 已登录：右下角账号徽章 → 点「账号」打开账号信息面板
 *   （查看当前登录账号：邮箱/用户ID；修改密码；退出登录）
 * - 解锁后广播 window 事件 site-auth-ready + localStorage 'site-auth-ping'
 *   （其他 iframe / 标签页监听该键自动解锁；退出时同样广播→各页跳登录页）
 * ================================================================ */
(function(){
  if(window.__SITE_AUTH_GATE__) return; window.__SITE_AUTH_GATE__ = true;

  var PK = (window.CLOUD_PK) || 'wbpk_KQXIeDHLXWDHN7WQp3zqUQ_Y5IlM6hJd7n10U8UN5bYmSGUYeevVOwl';
  var ENDPOINT = (window.CLOUD_ENDPOINT) || (typeof location !== 'undefined' ? location.origin : '');
  var SESSION_KEY = 'workbuddy-cloud.session.' + PK;
  var PING_KEY = 'site-auth-ping';
  var LOGIN_URL = '/pages/login/index.html';
  var ADMIN_KEY = 'site-admin-session';   // 默认管理员本地会话（登录页 seed，30 天）

  function readAdminSession(){
    try{
      var s = JSON.parse(localStorage.getItem(ADMIN_KEY) || 'null');
      return (s && s.user && Number(s.expiresAt) > Date.now()) ? s : null;
    }catch(e){ return null; }
  }

  /* ---------- SDK 加载（若宿主页已加载则直接复用） ---------- */
  function loadSdk(cb){
    if(window.WorkBuddyCloud) return cb();
    var urls = [
      'https://cdn.jsdelivr.net/npm/@tencent-ai/workbuddy-cloud-sdk@dev/lib/index.global.js',
      'https://unpkg.com/@tencent-ai/workbuddy-cloud-sdk@dev/lib/index.global.js',
      'https://registry.npmmirror.com/@tencent-ai/workbuddy-cloud-sdk/latest/files/lib/index.global.js'
    ];
    var i = 0;
    (function tryNext(){
      if(i >= urls.length){ setTimeout(function(){ tryNext(); }, 5000); return; }  // 全部失败 5s 后重试
      var s = document.createElement('script');
      s.src = urls[i++];
      s.onload = function(){ if(window.WorkBuddyCloud) cb(); else tryNext(); };
      s.onerror = function(){ tryNext(); };
      document.head.appendChild(s);
    })();
  }

  var cloud = null, user = null;

  function getCloud(){
    if(cloud) return cloud;
    if(typeof window.WorkBuddyCloud === 'undefined') return null;
    try{ cloud = window.WorkBuddyCloud.createWorkBuddyCloud({ endpoint: ENDPOINT, publishableKey: PK }); }
    catch(e){ cloud = null; }
    return cloud;
  }

  /* ---------- 跳转登录页 ---------- */
  function isTop(){ try{ return window.self === window.top; }catch(e){ return false; } }
  function goLogin(){
    if(window.__SA_NO_REDIRECT) return;                       // 测试钩子：禁止跳转
    try{
      var p = (typeof location !== 'undefined' && location.pathname) || '';
      if(p.indexOf('/pages/login/') === 0) return;            // 已在登录页，防循环
      var next = p + ((typeof location !== 'undefined' && location.search) || '');
      var url = LOGIN_URL + '?next=' + encodeURIComponent(next);
      if(isTop()) location.replace(url);
      else { try{ window.top.location.replace(url); }catch(e){} }
    }catch(e){}
  }

  /* ---------- UI：徽章 + 账号信息面板 ---------- */
  var CSS = ''
    + '#sa-badge{position:fixed;right:12px;bottom:12px;z-index:2147483000;background:#fff;border:1px solid #d5dde7;border-radius:999px;box-shadow:0 4px 14px rgba(15,35,64,.18);padding:6px 12px;font-size:12px;color:#1c2b3a;display:none;align-items:center;gap:8px;font-family:inherit;}'
    + '#sa-badge b{color:#123c73;font-weight:600;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}'
    + '#sa-badge button{border:none;background:none;color:#123c73;cursor:pointer;font-size:12px;font-family:inherit;padding:2px;}'
    + '#sa-badge button:hover{text-decoration:underline;}'
    + '#sa-panel{position:fixed;right:12px;bottom:52px;z-index:2147483000;background:#fff;border:1px solid #d5dde7;border-radius:14px;box-shadow:0 14px 44px rgba(15,35,64,.3);padding:16px 18px;width:min(320px,92vw);font-family:inherit;color:#1c2b3a;display:none;}'
    + '#sa-panel h4{margin:0 0 10px;font-size:14px;color:#123c73;}'
    + '#sa-panel .sa-info{background:#f4f7fb;border-radius:8px;padding:10px 12px;font-size:12px;line-height:1.9;color:#1c2b3a;}'
    + '#sa-panel .sa-info b{color:#123c73;word-break:break-all;}'
    + '#sa-panel label{display:block;font-size:12px;color:#5a6b80;margin:10px 0 4px;}'
    + '#sa-panel input{width:100%;box-sizing:border-box;border:1px solid #d5dde7;border-radius:8px;padding:8px 10px;font-size:13px;font-family:inherit;outline:none;}'
    + '#sa-panel input:focus{border-color:#2a5da8;}'
    + '#sa-panel .sa-row{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;}'
    + '#sa-panel button{border:1px solid #123c73;background:#fff;color:#123c73;border-radius:8px;padding:8px 12px;font-size:13px;font-family:inherit;cursor:pointer;}'
    + '#sa-panel button.sa-pri{background:#123c73;color:#fff;}'
    + '#sa-panel button.sa-danger{border-color:#c0392b;color:#c0392b;}'
    + '#sa-panel button:disabled{opacity:.55;cursor:default;}'
    + '#sa-panel .sa-chg-tt{margin-top:14px;padding-top:12px;border-top:1px dashed #e4ebf3;font-size:12px;color:#5a6b80;font-weight:600;}'
    + '#sa-msg{margin-top:8px;font-size:12px;min-height:15px;color:#c0392b;}';
  var style = document.createElement('style'); style.textContent = CSS; document.head.appendChild(style);

  var badge = document.createElement('div'); badge.id = 'sa-badge';
  badge.innerHTML = '<span>👤</span><b id="sa-badge-name"></b><button id="sa-badge-acct">账号</button>';
  var panel = document.createElement('div'); panel.id = 'sa-panel';
  panel.innerHTML = ''
    + '<h4>👤 当前登录账号</h4>'
    + '<div class="sa-info">'
    + '邮箱：<b id="sa-info-email">—</b><br>'
    + '用户ID：<b id="sa-info-uid">—</b>'
    + '</div>'
    + '<div class="sa-chg-tt">修改密码</div>'
    + '<label>当前密码</label><input id="sa-old-pwd" type="password" autocomplete="current-password">'
    + '<label>新密码（至少 6 位）</label><input id="sa-new-pwd" type="password" autocomplete="new-password">'
    + '<label>确认新密码</label><input id="sa-new-pwd2" type="password" autocomplete="new-password">'
    + '<div id="sa-msg"></div>'
    + '<div class="sa-row">'
    + '<button class="sa-pri" id="sa-chg">🔑 修改密码</button>'
    + '<button class="sa-danger" id="sa-logout">退出登录</button>'
    + '</div>';

  var _uiAppended = false;
  function appendUI(){
    if(_uiAppended || !document.body) return;
    _uiAppended = true;
    document.body.appendChild(badge);
    document.body.appendChild(panel);
    if(user){ badge.style.display = 'flex'; fillInfo(); }   // 解锁早于 DOM 就绪时补显徽章
  }
  document.addEventListener('DOMContentLoaded', appendUI);
  if(document.readyState !== 'loading'){ appendUI(); }

  function $(id){ return document.getElementById(id); }
  function msg(t, isErr){ var el = $('sa-msg'); if(el){ el.textContent = t || ''; el.style.color = isErr ? '#c0392b' : '#1a7a3a'; } }
  function fillInfo(){
    var e = $('sa-info-email'), u = $('sa-info-uid'), n = $('sa-badge-name');
    // 服务端用户模型以 username 承载邮箱（验证码登录返回的 user 可能无 email 字段）
    var email = (user && (user.email || user.username || user.phone || user.name)) || '—';
    if(e) e.textContent = email;
    if(u) u.textContent = (user && user.id) || '—';
    if(n) n.textContent = email;
  }

  function unlock(){
    badge.style.display = 'flex';   // 显式覆盖 CSS 默认 display:none
    fillInfo();
    try{ localStorage.setItem(PING_KEY, String(Date.now())); }catch(e){}
    try{ window.dispatchEvent(new Event('site-auth-ready')); }catch(e){}
  }

  /* ---------- 登录态检查 ---------- */
  async function checkSession(){
    var admin = readAdminSession();
    if(admin){ user = admin.user; unlock(); return; }   // 管理员本地会话有效 → 直接解锁（不走云端，避免 SDK 触碰本地假会话）
    var c = getCloud(); if(!c){ setTimeout(function(){ checkSession(); }, 800); return; }
    var session = null;
    try{ var r = await c.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
    if(session && session.user){ user = session.user; unlock(); startKeepalive(); }
    else { user = null; goLogin(); }
  }
  var _ka = null;
  function startKeepalive(){
    if(_ka) return;
    _ka = setInterval(async function(){
      if(!document.visibilityState || document.visibilityState === 'visible'){
        var c = getCloud(); if(!c) return;
        try{ await c.auth.getSession(); }catch(e){}
      }
    }, 20 * 60 * 1000);
  }

  /* ---------- 动作 ---------- */
  async function actChange(){
    if(readAdminSession()){ msg('默认管理员为内置本地账号，不支持在线修改密码', true); return; }
    var c = getCloud(); if(!c){ msg('云端组件加载中，请稍候重试', true); return; }
    var o = ($('sa-old-pwd')||{}).value || '', p1 = ($('sa-new-pwd')||{}).value || '', p2 = ($('sa-new-pwd2')||{}).value || '';
    if(!o){ msg('请输入当前密码', true); return; }
    if(!p1 || p1.length < 6){ msg('新密码至少 6 位', true); return; }
    if(p1 !== p2){ msg('两次输入的新密码不一致', true); return; }
    if(p1 === o){ msg('新密码不能与当前密码相同', true); return; }
    var b = $('sa-chg'); b.disabled = true;
    var r = await c.auth.resetPasswordForOld({ oldPassword: o, newPassword: p1 });
    b.disabled = false;
    if(r.error){
      var em = (r.error.message || '') + '';
      // 「verification failed」= 服务端验当前密码失败：最常见原因是该账号从未设置过密码（仅验证码登录创建）
      if(/verification/i.test(em) || /验证失败/.test(em)){
        msg('当前密码验证失败。若该账号从未设置过密码（此前仅用验证码登录），请先退出登录，在登录页点「忘记密码？」通过邮箱验证码设置密码', true);
        return;
      }
      msg('修改失败：' + (em || '请检查当前密码'), true); return;
    }
    ['sa-old-pwd','sa-new-pwd','sa-new-pwd2'].forEach(function(id){ var el = $(id); if(el) el.value = ''; });
    msg('密码修改成功', false);
  }
  async function actLogout(){
    var c = getCloud(); if(!c) return;
    try{ await c.auth.signOut(); }catch(e){}
    user = null;
    try{ localStorage.removeItem(SESSION_KEY); localStorage.removeItem('dp_auth_backup'); localStorage.removeItem(ADMIN_KEY); }catch(e){}
    try{ localStorage.setItem(PING_KEY, String(Date.now())); }catch(e){}   // 通知其他 iframe/标签页
    if(panel) panel.style.display = 'none';
    if(badge) badge.style.display = 'none';
    goLogin();
  }

  /* ---------- 事件 ---------- */
  function bind(){
    if(!badge || !panel) return;
    badge.querySelector('#sa-badge-acct').addEventListener('click', function(){
      var on = panel.style.display === 'block';
      panel.style.display = on ? 'none' : 'block';
      if(!on) fillInfo();
    });
    panel.querySelector('#sa-chg').addEventListener('click', actChange);
    panel.querySelector('#sa-logout').addEventListener('click', actLogout);
    window.addEventListener('storage', function(e){
      if(e.key === PING_KEY){ checkSession(); }   // 其他 iframe/标签登录或退出 → 本页同步（登录=解锁，退出=跳登录页）
    });
  }

  loadSdk(function(){
    bind();
    checkSession();
  });
})();
