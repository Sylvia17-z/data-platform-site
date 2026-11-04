/* ================================================================
 * 海蓓医管平台 —— 腾讯云 CloudBase 后端桥接（2026-10-31）
 * 替代原 @tencent-ai/workbuddy-cloud-sdk (WorkBuddy Cloud Service)。
 *
 * 其它模块通过 window.CB 访问，尽量保持与原 createWorkBuddyCloud 客户端一致的调用面：
 *   CB.loadSdk(cb)           加载 SDK 并初始化（幂等）
 *   CB.client()    -> { auth, database:{ from(t){...} } }   // 兼容旧 c 对象
 *     c.auth.getSession()     -> { data:{ user, session, accessToken, refreshToken, expiresAt }, error }
 *     c.auth.signInWithPassword / signInWithOtp / signUp / resetPasswordForEmail / signOut / updateUser
 *     c.database.from(t).select(...).order(...).limit(n) -> { data, error }   // 底层 app.rdb() (PostgreSQL)
 *
 * 登录方式：用户名/密码 signInWithPassword 已就绪（无需 SMTP）；
 *          邮箱验证码/注册/找回 走 CloudBase v3 回调式（需控制台配置 SMTP 后生效）。
 * 管理员 admin/123 仍为本地内置账号（localStorage 假会话），不走云端、云同步整体跳过。
 * ================================================================ */
(function () {
  if (window.CB) return;

  // —— 目标环境（复用现有 trial：sylvia17）——
  var ENV = 'sylvia17-d1gcg5nv128d47eed';
  var REGION = 'ap-shanghai';
  // 发布可公开密钥（publishable key，公开设计，置于前端无妨）
  var PK = 'eyJhbGciOiJSUzI1NiIsImtpZCI6IjQ0NzIzYTVhLTZmZTUtNDI1Zi1iZGIzLWIyZmNiNzY5ODFjYSJ9.eyJpc3MiOiJodHRwczovL3N5bHZpYTE3LWQxZ2NnNW52MTI4ZDQ3ZWVkLmFwLXNoYW5naGFpLnRjYi1hcGkudGVuY2VudGNsb3VkYXBpLmNvbSIsInN1YiI6ImFub24iLCJhdWQiOiJzeWx2aWExNy1kMWdjZzVudjEyOGQ0N2VlZCIsImV4cCI6NDA5NzExOTkzNSwiaWF0IjoxNzkzNDM2NzM1LCJub25jZSI6InU5RTdmcW13UU9leTJLR2RxcTA0YWciLCJhdF9oYXNoIjoidTlFN2ZxbXdRT2V5MktHZHFxMDRhZyIsIm5hbWUiOiJBbm9ueW1vdXMiLCJzY29wZSI6ImFub255bW91cyIsInByb2plY3RfaWQiOiJzeWx2aWExNy1kMWdjZzVudjEyOGQ0N2VlZCIsIm1ldGEiOnsicGxhdGZvcm0iOiJQdWJsaXNoYWJsZUtleSJ9LCJyb2xlIjoiYW5vbiIsImlzX2Fub255bW91cyI6dHJ1ZSwiYXBwX21ldGFkYXRhIjp7InByb3ZpZGVyIjoiYW5vbnltb3VzIiwicHJvdmlkZXJzIjpbImFub255bW91cyJdfSwidXNlcl9tZXRhZGF0YSI6eyJuYW1lIjoiQW5vbnltb3VzIn0sInVzZXJfdHlwZSI6IiIsImNsaWVudF90eXBlIjoiY2xpZW50X3VzZXIiLCJpc19zeXN0ZW1fYWRtaW4iOmZhbHNlfQ.H_SRaRcQMCUFKXBJXeb6aJYv1Own9tOeDHNMPs8O0Qmc_KR9nz4IMYrEvMJtVji3pS5x9AGn60c7nMjFGSMFSDZ4yP-bZaGD1vmmqklNIn2R9UuseANKmriF_KA2DWKlwgwvx8MTXnfdOy2fFEl6aA8RsdfnHbr8w1oYuGs_iKawcAjSnjRPorYwJzXZgXs05_WJgDp_TvUnpggQ0NSQ-KJX2mMKnizf69Z70ETtzWnDRBtvy1l7wpcJU3l1Uff4LIjl0jW6d-7bXwJfzpJXBa2VJWIABOI7_ZWzRQibSYI4kF9HP9NZKk_unu-CxaVfrbdckq63bjgap44WBLtKfQ';
  var STATIC = 'https://sylvia17-d1gcg5nv128d47eed-1505667687.tcloudbaseapp.com';

  var SDK_URLS = [
    'https://static.cloudbase.net/cloudbase-js-sdk/3.10.4/cloudbase.full.js',
    'https://static.cloudbase.net/cloudbase-js-sdk/latest/cloudbase.full.js'
  ];

  var _app = null, _loading = false, _queue = [];
  function ensureApp() {
    if (_app) return _app;
    if (!window.cloudbase) return null;
    try {
      _app = window.cloudbase.init({
        env: ENV,
        region: REGION,
        accessKey: PK,
        auth: { detectSessionInUrl: true }
      });
    } catch (e) { _app = null; }
    return _app;
  }
  var _readyFired = false, _failedFired = false;
  function fireReady() { if (_readyFired) return; _readyFired = true; try { document.dispatchEvent(new Event('cloud-sdk-ready')); } catch (e) {} }
  function fireFailed() { if (_failedFired) return; _failedFired = true; try { document.dispatchEvent(new Event('cloud-sdk-failed')); } catch (e) {} }
  function loadSdk(cb) {
    if (_app) { fireReady(); if (cb) cb(); return; }
    if (window.cloudbase) { ensureApp(); fireReady(); if (cb) cb(); return; }
    if (_loading) { if (cb) _queue.push(cb); return; }
    _loading = true;
    var i = 0;
    (function next() {
      if (i >= SDK_URLS.length) { _loading = false; fireFailed(); if (cb) cb(); return; }
      var s = document.createElement('script');
      s.src = SDK_URLS[i++];
      s.onload = function () {
        _loading = false;
        ensureApp();
        var q = _queue.splice(0);
        q.forEach(function (f) { if (f) f(); });
        if (window.cloudbase) fireReady();
        if (cb) cb();
      };
      s.onerror = function () { next(); };
      document.head.appendChild(s);
    })();
  }
  // 任一页面引入桥接即主动拉起 SDK，确保 client() 首次调用前 SDK 已就绪
  loadSdk();

  // 适配 CloudBase v3 会话对象 → 旧代码期望的 { data:{ user, ... }, error }
  function adaptSession(res) {
    var sess = (res && res.data && res.data.session) || (res && res.session) || null;
    var err = (res && res.error) || null;
    return {
      data: {
        user: sess ? sess.user : null,
        session: sess,
        accessToken: sess ? sess.access_token : null,
        refreshToken: sess ? sess.refresh_token : null,
        expiresAt: sess ? sess.expires_at : 0
      },
      error: err
    };
  }

  function makeAuth(real) {
    if (!real) return null;
    return {
      getSession: function () {
        return real.getSession().then(adaptSession).catch(function (e) { return { data: { user: null, session: null }, error: e }; });
      },
      // 旧代码 actChange 用 resetPasswordForOld({oldPassword,newPassword})；CloudBase v3 改为登录态下直接改密码
      resetPasswordForOld: function (p) {
        return real.updateUser({ password: (p && p.newPassword) || (p && p.password) });
      },
      signInWithPassword: function (p) { return real.signInWithPassword(p); },
      signInWithOtp: function (p) { return real.signInWithOtp(p); },
      signUp: function (p) { return real.signUp(p); },
      resetPasswordForEmail: function (email) { return real.resetPasswordForEmail(email); },
      signOut: function () { return real.signOut(); },
      updateUser: function (p) { return real.updateUser(p); },
      onAuthStateChange: function (cb) { return real.onAuthStateChange ? real.onAuthStateChange(cb) : null; }
    };
  }

  function client() {
    var app = ensureApp();
    if (!app) return null;
    return {
      auth: makeAuth(app.auth()),
      // 底层 PostgreSQL：app.rdb().from(table) 返回 postgREST 风格 { data, error }
      database: { from: function (t) { return app.rdb().from(t); } }
    };
  }

  window.CB = {
    ENV: ENV, REGION: REGION, PK: PK, STATIC: STATIC,
    loadSdk: loadSdk,
    client: client,
    getApp: function () { return ensureApp(); }
  };
  // 兼容旧全局名：部分页面曾读 window.CLOUD_ENDPOINT / CLOUD_PK
  window.CLOUD_ENDPOINT = window.CLOUD_ENDPOINT || STATIC;
  window.CLOUD_PK = window.CLOUD_PK || PK;
})();
