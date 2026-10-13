/* ================================================================
 * 整站共享登录门（2026-10-13 需求：进入网站任意页面都需登录，登录后多端同步）
 * - 同源共享 SDK 会话（localStorage），一次登录全站/全部 iframe 生效
 * - 未登录：全屏遮罩强制登录（验证码 / 密码登录 / 忘记密码重置）
 * - 已登录：右下角账号徽章（修改密码 / 退出登录），20 分钟定时续期
 * - 解锁后广播 window 事件 site-auth-ready + localStorage 'site-auth-ping'
 *   （其他 iframe / 标签页监听该键自动解锁）
 * ================================================================ */
(function(){
  if(window.__SITE_AUTH_GATE__) return; window.__SITE_AUTH_GATE__ = true;

  var PK = (window.CLOUD_PK) || 'wbpk_KQXIeDHLXWDHN7WQp3zqUQ_Y5IlM6hJd7n10U8UN5bYmSGUYeevVOwl';
  var ENDPOINT = (window.CLOUD_ENDPOINT) || (typeof location !== 'undefined' ? location.origin : '');
  var SESSION_KEY = 'workbuddy-cloud.session.' + PK;
  var PING_KEY = 'site-auth-ping';

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

  var cloud = null, user = null, pendingOtp = null, pwdReset = null;

  function getCloud(){
    if(cloud) return cloud;
    if(typeof window.WorkBuddyCloud === 'undefined') return null;
    try{ cloud = window.WorkBuddyCloud.createWorkBuddyCloud({ endpoint: ENDPOINT, publishableKey: PK }); }
    catch(e){ cloud = null; }
    return cloud;
  }

  /* ---------- UI ---------- */
  var CSS = ''
    + '#sa-mask{position:fixed;inset:0;background:rgba(15,35,64,.55);backdrop-filter:blur(3px);display:none;align-items:center;justify-content:center;z-index:2147483000;}'
    + '#sa-mask.show{display:flex;}'
    + '#sa-box{background:#fff;border-radius:14px;box-shadow:0 18px 60px rgba(15,35,64,.35);padding:22px 24px;width:min(400px,94vw);font-family:inherit;color:#1c2b3a;}'
    + '#sa-box h3{margin:0 0 14px;font-size:16px;color:#123c73;}'
    + '#sa-box label{display:block;font-size:12px;color:#5a6b80;margin:10px 0 4px;}'
    + '#sa-box input{width:100%;box-sizing:border-box;border:1px solid #d5dde7;border-radius:8px;padding:8px 10px;font-size:13px;font-family:inherit;}'
    + '#sa-box .sa-row{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px;}'
    + '#sa-box button{border:1px solid #123c73;background:#fff;color:#123c73;border-radius:8px;padding:8px 12px;font-size:13px;font-family:inherit;cursor:pointer;}'
    + '#sa-box button.sa-pri{background:#123c73;color:#fff;}'
    + '#sa-box button:disabled{opacity:.55;cursor:default;}'
    + '#sa-tip{margin-top:12px;font-size:11px;color:#8a97a8;line-height:1.5;}'
    + '#sa-msg{margin-top:8px;font-size:12px;min-height:16px;color:#c0392b;}'
    + '#sa-badge{position:fixed;right:12px;bottom:12px;z-index:2147483000;background:#fff;border:1px solid #d5dde7;border-radius:999px;box-shadow:0 4px 14px rgba(15,35,64,.18);padding:6px 12px;font-size:12px;color:#1c2b3a;display:none;align-items:center;gap:8px;font-family:inherit;}'
    + '#sa-badge b{color:#123c73;font-weight:600;max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}'
    + '#sa-badge button{border:none;background:none;color:#123c73;cursor:pointer;font-size:12px;font-family:inherit;padding:2px;}'
    + '#sa-badge button:hover{text-decoration:underline;}';
  var style = document.createElement('style'); style.textContent = CSS; document.head.appendChild(style);

  var mask = document.createElement('div'); mask.id = 'sa-mask';
  mask.innerHTML = ''
    + '<div id="sa-box" role="dialog" aria-modal="true">'
    + '<h3 id="sa-title">🔐 登录后使用（数据多端同步）</h3>'
    + '<label>邮箱</label><input id="sa-email" type="email" placeholder="you@example.com" autocomplete="username">'
    + '<div id="sa-code-wrap" style="display:none;"><label>验证码</label><input id="sa-code" type="text" inputmode="numeric" placeholder="6 位验证码"></div>'
    + '<div id="sa-pwd-wrap"><label>密码（已注册用户）</label><input id="sa-pwd" type="password" placeholder="密码" autocomplete="current-password"></div>'
    + '<div id="sa-reset-wrap" style="display:none;"><label>重置密码（验证码已发送至上方邮箱）</label><input id="sa-reset-code" type="text" inputmode="numeric" placeholder="6 位验证码"><input id="sa-reset-pwd" type="password" placeholder="新密码（至少 6 位）" style="margin-top:6px;" autocomplete="new-password"></div>'
    + '<div id="sa-chg-wrap" style="display:none;"><label>修改密码</label><input id="sa-old-pwd" type="password" placeholder="当前密码" autocomplete="current-password"><input id="sa-new-pwd" type="password" placeholder="新密码（至少 6 位）" style="margin-top:6px;" autocomplete="new-password"><input id="sa-new-pwd2" type="password" placeholder="确认新密码" style="margin-top:6px;" autocomplete="new-password"></div>'
    + '<div id="sa-msg"></div>'
    + '<div class="sa-row">'
    + '<button class="sa-pri" id="sa-send">📧 获取验证码</button>'
    + '<button class="sa-pri" id="sa-verify">✅ 验证并登录</button>'
    + '<button id="sa-pwdlogin">🔑 密码登录</button>'
    + '<button id="sa-forgot">忘记密码？</button>'
    + '<button id="sa-reset" style="display:none;">✅ 重置密码</button>'
    + '<button id="sa-chg" style="display:none;">🔑 修改密码</button>'
    + '<button id="sa-logout" style="display:none;">退出登录</button>'
    + '</div>'
    + '<div id="sa-tip">首次使用邮箱即自动注册（新账号请同时填写密码）；数据按账号隔离，仅本人可读取；登录后各页面数据自动云端同步。</div>'
    + '</div>';
  var badge = document.createElement('div'); badge.id = 'sa-badge';
  badge.innerHTML = '<span>👤</span><b id="sa-badge-name"></b><button id="sa-badge-acct">账号</button>';
  var _uiAppended = false;
  function appendUI(){
    if(_uiAppended || !document.body) return;
    _uiAppended = true;
    document.body.appendChild(mask);
    document.body.appendChild(badge);
    if(user) badge.style.display = 'flex';   // 解锁早于 DOM 就绪时补显徽章
  }
  document.addEventListener('DOMContentLoaded', appendUI);
  if(document.readyState !== 'loading'){ appendUI(); }

  function $(id){ return document.getElementById(id) || (mask && mask.querySelector('#' + id)) || null; }
  function msg(t, isErr){ var el = $('sa-msg'); if(el){ el.textContent = t || ''; el.style.color = isErr ? '#c0392b' : '#1a7a3a'; } }
  function show(el, on){ if(el) el.style.display = on ? '' : 'none'; }

  function renderMode(){
    var logged = !!user;
    $('sa-title').textContent = logged ? '👤 账号' : '🔐 登录后使用（数据多端同步）';
    show($('sa-email'), !logged); show($('sa-pwd-wrap'), !logged);
    show($('sa-code-wrap'), !logged && !!pendingOtp);
    show($('sa-send'), !logged); show($('sa-verify'), !logged);
    show($('sa-pwdlogin'), !logged);
    show($('sa-forgot'), !logged && !pwdReset);
    show($('sa-reset'), !logged && !!pwdReset);
    show($('sa-reset-wrap'), !logged && !!pwdReset);
    show($('sa-chg'), logged); show($('sa-chg-wrap'), logged);
    show($('sa-logout'), logged);
  }
  function lock(){ mask.classList.add('show'); renderMode(); }
  function unlock(){
    mask.classList.remove('show');
    badge.style.display = 'flex';   // 显式覆盖 CSS 默认 display:none
    if(badge){ var n = $('sa-badge-name'); if(n) n.textContent = (user && (user.email || user.name || user.id)) || '已登录'; }
    try{ localStorage.setItem(PING_KEY, String(Date.now())); }catch(e){}
    try{ window.dispatchEvent(new Event('site-auth-ready')); }catch(e){}
  }

  /* ---------- 登录态检查 ---------- */
  async function checkSession(){
    var c = getCloud(); if(!c){ setTimeout(function(){ checkSession(); }, 800); return; }
    var session = null;
    try{ var r = await c.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
    if(session && session.user){ user = session.user; unlock(); startKeepalive(); }
    else { user = null; lock(); }
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
  async function actSend(){
    var email = ($('sa-email')||{}).value || '';
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ msg('请输入有效的邮箱', true); return; }
    var c = getCloud(); if(!c){ msg('云端组件加载中，请稍候重试', true); return; }
    var b = $('sa-send'); b.disabled = true; var old = b.textContent; b.textContent = '发送中…';
    var sent = await c.auth.sendOtp({ email: email });
    b.disabled = false; b.textContent = old;
    if(sent.error){ msg('验证码发送失败：' + (sent.error.message || ''), true); return; }
    pendingOtp = { email: email, verificationId: sent.data && sent.data.verificationId, isExistingUser: sent.data && sent.data.isExistingUser };
    msg('验证码已发送至 ' + email, false);
    renderMode();
  }
  async function actVerify(){
    if(!pendingOtp){ msg('请先获取验证码', true); return; }
    var c = getCloud(); if(!c){ msg('云端组件加载中，请稍候重试', true); return; }
    var code = ($('sa-code')||{}).value || '';
    var pwd = ($('sa-pwd')||{}).value || '';
    if(!code.trim()){ msg('请输入验证码', true); return; }
    var params = { email: pendingOtp.email, verificationId: pendingOtp.verificationId, isExistingUser: pendingOtp.isExistingUser, token: code.trim() };
    if(!pendingOtp.isExistingUser) params.password = pwd;   // 新账号注册需同时填密码
    var done = await c.auth.verifyOtp(params);
    if(done.error){ msg('登录失败：' + (done.error.message || ''), true); return; }
    pendingOtp = null; user = (done.data && done.data.user) || null; msg(''); unlock(); startKeepalive();
  }
  async function actPwdLogin(){
    var c = getCloud(); if(!c){ msg('云端组件加载中，请稍候重试', true); return; }
    var email = ($('sa-email')||{}).value || '';
    var pwd = ($('sa-pwd')||{}).value || '';
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ msg('请输入有效的邮箱', true); return; }
    if(!pwd){ msg('请输入密码', true); return; }
    var b = $('sa-pwdlogin'); b.disabled = true;
    var r = await c.auth.signInWithPassword({ email: email, password: pwd });
    b.disabled = false;
    if(r.error){ msg('登录失败：' + (r.error.message || ''), true); return; }
    user = (r.data && r.data.user) || null; msg(''); unlock(); startKeepalive();
  }
  async function actForgot(){
    var email = ($('sa-email')||{}).value || '';
    if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)){ msg('请输入有效的邮箱', true); return; }
    var c = getCloud(); if(!c){ msg('云端组件加载中，请稍候重试', true); return; }
    var b = $('sa-forgot'); b.disabled = true; var old = b.textContent; b.textContent = '发送中…';
    var r = await c.auth.resetPasswordForEmail(email);
    b.disabled = false; b.textContent = old;
    if(r.error){ msg('重置验证码发送失败：' + (r.error.message || ''), true); return; }
    if(r.data && typeof r.data.updateUser === 'function'){
      pwdReset = { email: email, updateUser: r.data.updateUser };
      msg('重置验证码已发送至 ' + email, false); renderMode();
    } else msg('重置流程返回异常，请稍候重试', true);
  }
  async function actReset(){
    if(!pwdReset){ msg('请先点击「忘记密码？」获取验证码', true); return; }
    var code = ($('sa-reset-code')||{}).value || '';
    var pwd = ($('sa-reset-pwd')||{}).value || '';
    if(!code.trim()){ msg('请输入邮箱验证码', true); return; }
    if(!pwd || pwd.length < 6){ msg('新密码至少 6 位', true); return; }
    var b = $('sa-reset'); b.disabled = true;
    var r = await pwdReset.updateUser({ nonce: code.trim(), password: pwd });
    b.disabled = false;
    if(r.error){ msg('密码重置失败：' + (r.error.message || '请检查验证码'), true); return; }
    pwdReset = null;
    ['sa-reset-code','sa-reset-pwd'].forEach(function(id){ var el = $(id); if(el) el.value = ''; });
    user = (r.data && r.data.user) || null; msg(''); unlock(); startKeepalive();
  }
  async function actChange(){
    var c = getCloud(); if(!c){ msg('云端组件加载中，请稍候重试', true); return; }
    var o = ($('sa-old-pwd')||{}).value || '', p1 = ($('sa-new-pwd')||{}).value || '', p2 = ($('sa-new-pwd2')||{}).value || '';
    if(!o){ msg('请输入当前密码', true); return; }
    if(!p1 || p1.length < 6){ msg('新密码至少 6 位', true); return; }
    if(p1 !== p2){ msg('两次输入的新密码不一致', true); return; }
    if(p1 === o){ msg('新密码不能与当前密码相同', true); return; }
    var b = $('sa-chg'); b.disabled = true;
    var r = await c.auth.resetPasswordForOld({ oldPassword: o, newPassword: p1 });
    b.disabled = false;
    if(r.error){ msg('修改失败：' + (r.error.message || '请检查当前密码'), true); return; }
    ['sa-old-pwd','sa-new-pwd','sa-new-pwd2'].forEach(function(id){ var el = $(id); if(el) el.value = ''; });
    msg('密码修改成功', false);
  }
  async function actLogout(){
    var c = getCloud(); if(!c) return;
    try{ await c.auth.signOut(); }catch(e){}
    user = null;
    try{ localStorage.removeItem(SESSION_KEY); localStorage.removeItem('dp_auth_backup'); }catch(e){}
    badge.style.display = 'none';
    lock();
  }

  /* ---------- 事件 ---------- */
  function bind(){
    var q = function(id){ return mask.querySelector('#' + id) || document.getElementById(id); };
    q('sa-send').addEventListener('click', actSend);
    q('sa-verify').addEventListener('click', actVerify);
    q('sa-pwdlogin').addEventListener('click', actPwdLogin);
    q('sa-forgot').addEventListener('click', actForgot);
    q('sa-reset').addEventListener('click', actReset);
    q('sa-chg').addEventListener('click', actChange);
    q('sa-logout').addEventListener('click', actLogout);
    badge.querySelector('#sa-badge-acct').addEventListener('click', function(){ lock(); });   // 徽章「账号」→ 打开账号面板
    mask.addEventListener('click', function(e){ if(e.target === mask){ /* 强制登录：点空白不关闭 */ } });
    window.addEventListener('storage', function(e){
      if(e.key === PING_KEY){ checkSession(); }                             // 其他 iframe/标签登录 → 本页解锁
    });
  }

  loadSdk(function(){
    bind();
    checkSession();
  });
})();
