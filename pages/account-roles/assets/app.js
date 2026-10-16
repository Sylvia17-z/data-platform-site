/* ================================================================
 * 账号与角色权限管理 · 页面逻辑
 * 依赖：/shared/rbac.js（window.RBAC）
 * 入口：先 await RBAC.ready() 解析会话与 ACL，再校验 manageAcl。
 * 编辑在本地 work 副本上进行，保存时统一 RBAC.saveAcl(work)（写云 dp_acl 或本机缓存）。
 * ================================================================ */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  function clone(o) { try { return JSON.parse(JSON.stringify(o)); } catch (e) { return null; } }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  var work = null;          // 本地编辑副本
  var editingRoleId = null; // 当前角色编辑器（null=新建）
  var editingOpId = null;   // 当前操作员编辑器

  function toast(msg, kind) {
    var t = $('toast'); if (!t) return;
    t.textContent = msg; t.className = 'toast show' + (kind ? ' ' + kind : '');
    clearTimeout(t._t); t._t = setTimeout(function () { t.className = 'toast'; }, 2600);
  }

  function fullPermsMap() { var a = {}; RBAC.permissionCatalog.actions.forEach(function (x) { a[x.key] = true; }); return a; }
  function allModuleIds() { return RBAC.permissionCatalog.modules.map(function (m) { return m.id; }); }

  /* ---------- 入口 ---------- */
  RBAC.ready().then(function () {
    if (!RBAC.can('manageAcl')) {
      $('admin-ui').style.display = 'none';
      $('save-bar').style.display = 'none';
      $('no-permission').style.display = 'block';
      return;
    }
    $('admin-ui').style.display = 'block';
    $('save-bar').style.display = 'flex';
    reload();
  });

  function reload() {
    work = clone(RBAC.getAcl()) || { ts: Date.now(), roles: [], operators: [], assignments: {} };
    if (!work.roles) work.roles = [];
    if (!work.operators) work.operators = [];
    if (!work.assignments) work.assignments = {};
    renderBanner();
    renderCurrentOp();
    renderRoles();
    renderOps();
  }

  /* ---------- 横幅 ---------- */
  function renderBanner() {
    var slot = $('banner-slot');
    if (RBAC.isCloudAcl()) {
      slot.innerHTML = '<div class="banner ok">☁️ 云端共享模式：配置已写入云端共享账号，<b>全组织（共用该云账号的所有设备）</b>自动生效。独立云账号用户仅能查看自己那份，属纯前端架构限制。</div>';
    } else {
      slot.innerHTML = '<div class="banner warn">⚠️ 本地管理员模式：当前为内置本地账号，权限配置<b>仅本机有效，不会同步到其他设备</b>。请使用共享云账号登录以启用全组织共享。</div>';
    }
  }

  /* ---------- ① 本机当前操作员 ---------- */
  function renderCurrentOp() {
    var sel = $('current-op');
    sel.innerHTML = work.operators.map(function (o) {
      return '<option value="' + esc(o.id) + '">' + esc(o.name) + (o.accountId ? '（' + esc(o.accountId) + '）' : '') + '</option>';
    }).join('') || '<option value="">（暂无操作员）</option>';
    var cur = localStorage.getItem('acl_current_operator') || (work.operators[0] && work.operators[0].id) || '';
    sel.value = cur;
    updateOpInfo();
    sel.onchange = function () {
      RBAC.selectOperator(sel.value);
      updateOpInfo();
      toast('已切换本机操作员：' + (sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : ''));
    };
  }
  function updateOpInfo() {
    var id = $('current-op').value;
    var op = work.operators.filter(function (o) { return o.id === id; })[0];
    var info = $('current-op-info');
    if (!op) { info.textContent = ''; return; }
    var role = work.roles.filter(function (r) { return r.id === op.roleId; })[0];
    info.textContent = '→ 角色：' + (role ? role.name : '（未分配角色）');
  }

  /* ---------- ② 角色 CRUD ---------- */
  function renderRoles() {
    var list = $('role-list');
    if (!work.roles.length) { list.innerHTML = '<div class="empty">暂无角色，点击「新建角色」创建。</div>'; }
    else {
      list.innerHTML = work.roles.map(function (r) {
        var pa = r.permissions && r.permissions.actions ? r.permissions.actions : {};
        var actCnt = RBAC.permissionCatalog.actions.filter(function (a) { return pa[a.key]; }).length;
        var modCnt = r.permissions && r.permissions.modules ? r.permissions.modules.length : 0;
        return '<div class="card"><h3>' + esc(r.name) + '</h3>' +
          '<div class="mut">操作权限：' + actCnt + ' / ' + RBAC.permissionCatalog.actions.length + '　菜单可见：' + modCnt + ' / ' + allModuleIds().length + '</div></div>';
      }).join('');
    }
    var rs = $('role-select');
    rs.innerHTML = work.roles.map(function (r) { return '<option value="' + esc(r.id) + '">' + esc(r.name) + '</option>'; }).join('');
    $('role-editor').style.display = 'none';
    editingRoleId = null;
  }
  function roleById(id) { return work.roles.filter(function (r) { return r.id === id; })[0]; }

  function openRoleEditor(id) {
    editingRoleId = id || null;
    var r = id ? roleById(id) : { id: 'r_' + Date.now(), name: '', permissions: { actions: {}, modules: [] } };
    if (!r) return;
    var acts = r.permissions && r.permissions.actions ? r.permissions.actions : {};
    var mods = r.permissions && r.permissions.modules ? r.permissions.modules : [];
    var html = '<h3>' + (id ? '编辑角色：' + esc(r.name) : '新建角色') + '</h3>' +
      '<div class="row"><span class="label">角色名称</span><input id="re-name" class="input" value="' + esc(r.name) + '" placeholder="如：数据录入员" style="min-width:200px;"></div>' +
      '<div class="grp-title">操作权限</div><div class="checks col2">' +
      RBAC.permissionCatalog.actions.map(function (a) {
        return '<label class="chk"><input type="checkbox" data-act="' + a.key + '" ' + (acts[a.key] ? 'checked' : '') + '>' + esc(a.label) + '</label>';
      }).join('') + '</div>' +
      '<div class="grp-title">菜单可见</div><div class="checks col2">' +
      RBAC.permissionCatalog.modules.map(function (m) {
        return '<label class="chk"><input type="checkbox" data-mod="' + m.id + '" ' + (mods.indexOf(m.id) >= 0 ? 'checked' : '') + '>' + esc(m.name) + '</label>';
      }).join('') + '</div>' +
      '<div class="row" style="margin-top:12px;"><button class="btn btn-primary btn-sm" id="re-save">保存角色</button><button class="btn btn-sm" id="re-cancel">取消</button></div>';
    var box = $('role-editor'); box.innerHTML = html; box.style.display = 'block';
    $('re-save').onclick = function () {
      var name = $('re-name').value.trim();
      if (!name) { toast('请填写角色名称', 'err'); return; }
      var actions = {}; box.querySelectorAll('[data-act]').forEach(function (c) { if (c.checked) actions[c.getAttribute('data-act')] = true; });
      var modules = []; box.querySelectorAll('[data-mod]').forEach(function (c) { if (c.checked) modules.push(c.getAttribute('data-mod')); });
      if (id) { r.name = name; r.permissions = { actions: actions, modules: modules }; }
      else { work.roles.push({ id: r.id, name: name, permissions: { actions: actions, modules: modules } }); }
      renderRoles(); toast('角色已保存（未提交云端，点「保存配置」生效）');
    };
    $('re-cancel').onclick = function () { box.style.display = 'none'; editingRoleId = null; };
  }

  function deleteRole(id) {
    if (!id) return;
    if (!confirm('确认删除该角色？已绑定此角色的操作员将变为「未分配角色」。')) return;
    work.roles = work.roles.filter(function (r) { return r.id !== id; });
    work.operators.forEach(function (o) { if (o.roleId === id) o.roleId = null; });
    renderRoles(); renderOps(); toast('角色已删除（点「保存配置」生效）');
  }

  $('role-new').onclick = function () { openRoleEditor(null); };
  $('role-edit').onclick = function () { var id = $('role-select').value; if (id) openRoleEditor(id); else toast('请先选择角色', 'err'); };
  $('role-del').onclick = function () { var id = $('role-select').value; if (id) deleteRole(id); else toast('请先选择角色', 'err'); };

  /* ---------- ③ 操作员 CRUD ---------- */
  function renderOps() {
    var list = $('op-list');
    if (!work.operators.length) { list.innerHTML = '<div class="empty">暂无操作员，点击「新建操作员」创建。</div>'; }
    else {
      list.innerHTML = work.operators.map(function (o) {
        var role = work.roles.filter(function (r) { return r.id === o.roleId; })[0];
        var cur = localStorage.getItem('acl_current_operator');
        return '<div class="card"><h3>' + esc(o.name) +
          (cur === o.id ? ' · 本机' : '') + '</h3>' +
          '<div class="mut">账号：' + esc(o.accountId || '（未绑定）') + '</div>' +
          '<div class="mut">角色：' + (role ? esc(role.name) : '（未分配）') + '</div>' +
          '<div class="row" style="margin-top:8px;">' +
          '<button class="btn btn-sm" data-opedit="' + esc(o.id) + '">编辑</button>' +
          '<button class="btn btn-sm btn-danger" data-opdel="' + esc(o.id) + '">删除</button></div></div>';
      }).join('');
      list.querySelectorAll('[data-opedit]').forEach(function (b) { b.onclick = function () { openOpEditor(b.getAttribute('data-opedit')); }; });
      list.querySelectorAll('[data-opdel]').forEach(function (b) { b.onclick = function () { deleteOp(b.getAttribute('data-opdel')); }; });
    }
    renderCurrentOp();
  }
  function opById(id) { return work.operators.filter(function (o) { return o.id === id; })[0]; }

  function openOpEditor(id) {
    editingOpId = id || null;
    var o = id ? opById(id) : { id: 'op_' + Date.now(), name: '', accountId: '', roleId: '' };
    if (!o) return;
    var html = '<h3>' + (id ? '编辑操作员：' + esc(o.name) : '新建操作员') + '</h3>' +
      '<div class="row"><span class="label">姓名</span><input id="oe-name" class="input" value="' + esc(o.name) + '" placeholder="如：张三" style="min-width:160px;"></div>' +
      '<div class="row"><span class="label">云账号</span><input id="oe-account" class="input" value="' + esc(o.accountId) + '" placeholder="登录云端的 uid 或邮箱" style="min-width:220px;"></div>' +
      '<div class="row"><span class="label">角色</span><select id="oe-role" class="select">' +
      '<option value="">（未分配）</option>' +
      work.roles.map(function (r) { return '<option value="' + esc(r.id) + '" ' + (r.id === o.roleId ? 'selected' : '') + '>' + esc(r.name) + '</option>'; }).join('') +
      '</select></div>' +
      '<div class="mut">说明：云账号用于匹配当前登录身份；同一云账号多人共用时，再用「本机当前操作员」区分具体是谁。</div>' +
      '<div class="row" style="margin-top:12px;"><button class="btn btn-primary btn-sm" id="oe-save">保存操作员</button><button class="btn btn-sm" id="oe-cancel">取消</button></div>';
    var box = $('op-editor'); box.innerHTML = html; box.style.display = 'block';
    $('oe-save').onclick = function () {
      var name = $('oe-name').value.trim();
      if (!name) { toast('请填写姓名', 'err'); return; }
      var account = $('oe-account').value.trim();
      var roleId = $('oe-role').value;
      if (id) { o.name = name; o.accountId = account; o.roleId = roleId || null; }
      else { work.operators.push({ id: o.id, name: name, accountId: account, roleId: roleId || null }); }
      renderOps(); toast('操作员已保存（点「保存配置」生效）');
    };
    $('oe-cancel').onclick = function () { box.style.display = 'none'; editingOpId = null; };
  }

  function deleteOp(id) {
    if (!id) return;
    if (!confirm('确认删除该操作员档案？')) return;
    work.operators = work.operators.filter(function (o) { return o.id !== id; });
    if (localStorage.getItem('acl_current_operator') === id) localStorage.removeItem('acl_current_operator');
    renderOps(); toast('操作员已删除（点「保存配置」生效）');
  }

  $('op-new').onclick = function () { openOpEditor(null); };

  /* ---------- 保存 / 放弃 ---------- */
  $('save-btn').onclick = function () {
    if (!work.roles.length) { if (!confirm('当前没有任何角色，所有登录用户将被拒绝全部操作（菜单仍可见）。确定保存？')) return; }
    RBAC.saveAcl(work);
    $('save-status').textContent = RBAC.isCloudAcl() ? '已写入云端共享账号 ✓' : '已写入本机（本地管理员模式）';
    toast('配置已保存', 'ok');
    // 保存后重新载入，保持 work 与引擎一致（含自动生成的 id）
    reload();
  };
  $('reset-btn').onclick = function () {
    reload(); $('save-status').textContent = ''; toast('已放弃未保存的修改');
  };
})();
