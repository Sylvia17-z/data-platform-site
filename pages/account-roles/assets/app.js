/* ================================================================
 * 账号与角色权限管理 · 页面逻辑（方案A · 左右主从布局，2026-10-16 重构）
 * 布局：左侧「角色列表 + 操作员列表」点选 → 右侧编辑器随选中项切换（点谁配谁）。
 * 依赖：/shared/rbac.js（window.RBAC）
 * 入口：先 await RBAC.ready() 解析会话与 ACL，再校验 manageAcl。
 * 编辑在本地 work 副本上进行，保存时统一 RBAC.saveAcl(work)（写云 dp_acl 或本机缓存）。
 * 兼容性约定：保留 #banner-slot / #no-permission / #admin-ui / #save-bar / #toast 等既有锚点
 * （test-account-roles.js 与 auth-gate 依赖这些 ID）。
 * ================================================================ */
(function () {
  'use strict';
  var $ = function (id) { return document.getElementById(id); };
  function clone(o) { try { return JSON.parse(JSON.stringify(o)); } catch (e) { return null; } }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  var work = null;        // 本地编辑副本
  var selRole = null;     // 当前选中角色 id（'__new__' = 正在新建）
  var selOp = null;       // 当前选中操作员 id（'__new__' = 正在新建）
  var NEW = '__new__';

  function toast(msg, kind) {
    var t = $('toast'); if (!t) return;
    t.textContent = msg; t.className = 'toast show' + (kind ? ' ' + kind : '');
    clearTimeout(t._t); t._t = setTimeout(function () { t.className = 'toast'; }, 2600);
  }

  function allModuleIds() { return RBAC.permissionCatalog.modules.map(function (m) { return m.id; }); }
  function roleById(id) { return work.roles.filter(function (r) { return r.id === id; })[0]; }
  function opById(id) { return work.operators.filter(function (o) { return o.id === id; })[0]; }

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
    // 选中项失效清理
    if (selRole && selRole !== NEW && !roleById(selRole)) selRole = null;
    if (selOp && selOp !== NEW && !opById(selOp)) selOp = null;
    renderBanner();
    renderCurrentOp();
    renderRoleList();
    renderOpList();
    renderEditor();
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

  /* ---------- ① 本机当前操作员（顶条） ---------- */
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
      renderOpList(); // 「本机」标记跟随切换
      toast('已切换本机操作员：' + (sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text : ''));
    };
  }
  function updateOpInfo() {
    var id = $('current-op').value;
    var op = opById(id);
    var info = $('current-op-info');
    if (!op) { info.textContent = ''; info.style.display = 'none'; return; }
    info.style.display = '';
    var role = work.roles.filter(function (r) { return r.id === op.roleId; })[0];
    info.textContent = '→ 角色：' + (role ? role.name : '（未分配角色）');
  }

  /* ---------- 左侧：角色列表 ---------- */
  function renderRoleList() {
    var ul = $('role-list');
    $('role-count').textContent = '（' + work.roles.length + '）';
    if (!work.roles.length) {
      ul.innerHTML = '<li class="empty-li">暂无角色，点击下方「＋ 新建角色」创建</li>';
      return;
    }
    ul.innerHTML = work.roles.map(function (r) {
      var pa = r.permissions && r.permissions.actions ? r.permissions.actions : {};
      var actCnt = RBAC.permissionCatalog.actions.filter(function (a) { return pa[a.key]; }).length;
      var modCnt = r.permissions && r.permissions.modules ? r.permissions.modules.length : 0;
      return '<li data-roleid="' + esc(r.id) + '"' + (selRole === r.id ? ' class="on"' : '') + '>' +
        '<span class="nm">' + esc(r.name) + '</span>' +
        '<span class="cnt">' + actCnt + '/' + RBAC.permissionCatalog.actions.length + ' 权限 · ' + modCnt + '/' + allModuleIds().length + ' 菜单</span></li>';
    }).join('');
    Array.prototype.forEach.call(ul.querySelectorAll('[data-roleid]'), function (li) {
      li.onclick = function () { selRole = li.getAttribute('data-roleid'); selOp = null; renderRoleList(); renderOpList(); renderEditor(); };
    });
  }

  /* ---------- 左侧：操作员列表 ---------- */
  function renderOpList() {
    var ul = $('op-list');
    $('op-count').textContent = '（' + work.operators.length + '）';
    var cur = localStorage.getItem('acl_current_operator');
    if (!work.operators.length) {
      ul.innerHTML = '<li class="empty-li">暂无操作员，点击下方「＋ 新建操作员」创建</li>';
      return;
    }
    ul.innerHTML = work.operators.map(function (o) {
      var role = work.roles.filter(function (r) { return r.id === o.roleId; })[0];
      return '<li data-opid="' + esc(o.id) + '"' + (selOp === o.id ? ' class="on"' : '') + '>' +
        '<span class="nm">' + esc(o.name) + (cur === o.id ? ' · 本机' : '') + '</span>' +
        '<span class="cnt">' + (role ? esc(role.name) : '（未分配）') + '</span></li>';
    }).join('');
    Array.prototype.forEach.call(ul.querySelectorAll('[data-opid]'), function (li) {
      li.onclick = function () { selOp = li.getAttribute('data-opid'); selRole = null; renderRoleList(); renderOpList(); renderEditor(); };
    });
  }

  /* ---------- 右侧编辑器：随选中项切换 ---------- */
  function renderEditor() {
    if (selRole) renderRoleEditor();
    else if (selOp) renderOpEditor();
    else renderEmptyEditor();
  }

  function renderEmptyEditor() {
    $('editor-head').innerHTML = '<h2>权限配置</h2>';
    $('editor-body').innerHTML = '<div class="placeholder">从左侧选择一个角色或操作员开始配置<br>（点击「＋ 新建」可创建新角色 / 操作员）</div>';
  }

  /* ----- 角色 编辑器 ----- */
  function renderRoleEditor() {
    var isNew = selRole === NEW;
    var r = isNew ? null : roleById(selRole);
    if (!isNew && !r) { selRole = null; renderEditor(); return; }
    var acts = r && r.permissions && r.permissions.actions ? r.permissions.actions : {};
    var mods = r && r.permissions && r.permissions.modules ? r.permissions.modules : [];
    var bound = isNew ? 0 : work.operators.filter(function (o) { return o.roleId === selRole; }).length;

    $('editor-head').innerHTML = '<h2>' + (isNew ? '新建角色' : '编辑角色：' + esc(r.name)) + '</h2>' +
      (isNew ? '' : '<span class="chip">绑定操作员 ' + bound + ' 人</span>') +
      (isNew ? '' : '<button type="button" class="btn btn-danger btn-sm" id="re-del" style="margin-left:auto;">删除角色</button>');

    var html = '<div class="hint">改动先存本地副本，点底部「保存配置」统一生效（写入' + (RBAC.isCloudAcl() ? '云端共享账号' : '本机') + '）。</div>' +
      '<div class="namerow"><span class="label">角色名称</span><input id="re-name" class="input" value="' + esc(r ? r.name : '') + '" placeholder="如：数据录入员" style="min-width:220px;"></div>' +
      '<div class="grp-title">操作权限（<span id="re-act-cnt"></span> / ' + RBAC.permissionCatalog.actions.length + '）</div>' +
      '<div class="checks">' +
      RBAC.permissionCatalog.actions.map(function (a) {
        return '<label class="chk' + (acts[a.key] ? ' on' : '') + '"><input type="checkbox" data-act="' + a.key + '"' + (acts[a.key] ? ' checked' : '') + '>' + esc(a.label) + '</label>';
      }).join('') + '</div>' +
      '<div class="grp-title">菜单可见（<span id="re-mod-cnt"></span> / ' + allModuleIds().length + '）</div>' +
      '<div class="checks">' +
      RBAC.permissionCatalog.modules.map(function (m) {
        return '<label class="chk' + (mods.indexOf(m.id) >= 0 ? ' on' : '') + '"><input type="checkbox" data-mod="' + m.id + '"' + (mods.indexOf(m.id) >= 0 ? ' checked' : '') + '>' + esc(m.name) + '</label>';
      }).join('') + '</div>' +
      '<div class="foot"><button type="button" class="btn btn-primary" id="re-save">保存角色</button><button type="button" class="btn" id="re-cancel">取消修改</button></div>';
    var body = $('editor-body');
    body.innerHTML = html;

    // 勾选联动：外层 label 高亮 + 计数刷新
    function refreshCnt() {
      var on = body.querySelectorAll('[data-act]:checked').length;
      $('re-act-cnt').textContent = on;
      $('re-mod-cnt').textContent = body.querySelectorAll('[data-mod]:checked').length;
    }
    Array.prototype.forEach.call(body.querySelectorAll('.chk input'), function (c) {
      c.addEventListener('change', function () {
        c.parentNode.classList.toggle('on', c.checked);
        refreshCnt();
      });
    });
    refreshCnt();

    $('re-save').onclick = function () {
      var name = $('re-name').value.trim();
      if (!name) { toast('请填写角色名称', 'err'); return; }
      var actions = {}; body.querySelectorAll('[data-act]').forEach(function (c) { if (c.checked) actions[c.getAttribute('data-act')] = true; });
      var modules = []; body.querySelectorAll('[data-mod]').forEach(function (c) { if (c.checked) modules.push(c.getAttribute('data-mod')); });
      if (isNew) {
        var nr = { id: 'r_' + Date.now(), name: name, permissions: { actions: actions, modules: modules } };
        work.roles.push(nr);
        selRole = nr.id;
      } else {
        r.name = name; r.permissions = { actions: actions, modules: modules };
      }
      renderRoleList(); renderOpList(); renderEditor();
      toast('角色已保存（未提交云端，点「保存配置」生效）');
    };
    $('re-cancel').onclick = function () {
      if (isNew) { selRole = null; renderRoleList(); renderEditor(); }
      else { renderRoleList(); renderEditor(); toast('已放弃该角色的未保存修改'); }
    };
    var del = $('re-del');
    if (del) del.onclick = function () { deleteRole(selRole); };
  }

  function deleteRole(id) {
    if (!id) return;
    if (!confirm('确认删除该角色？已绑定此角色的操作员将变为「未分配角色」。')) return;
    work.roles = work.roles.filter(function (r) { return r.id !== id; });
    work.operators.forEach(function (o) { if (o.roleId === id) o.roleId = null; });
    selRole = null;
    renderRoleList(); renderOpList(); renderEditor();
    toast('角色已删除（点「保存配置」生效）');
  }

  /* ----- 操作员 编辑器 ----- */
  function renderOpEditor() {
    var isNew = selOp === NEW;
    var o = isNew ? null : opById(selOp);
    if (!isNew && !o) { selOp = null; renderEditor(); return; }

    $('editor-head').innerHTML = '<h2>' + (isNew ? '新建操作员' : '编辑操作员：' + esc(o.name)) + '</h2>' +
      (isNew ? '' : '<span class="chip">' + esc(o.accountId || '（未绑定云账号）') + '</span>') +
      (isNew ? '' : '<button type="button" class="btn btn-danger btn-sm" id="oe-del" style="margin-left:auto;">删除操作员</button>');

    var html = '<div class="hint">云账号用于匹配当前登录身份；同一云账号多人共用时，再用「本机当前操作员」（顶部）区分具体是谁。</div>' +
      '<div class="namerow"><span class="label">姓名</span><input id="oe-name" class="input" value="' + esc(o ? o.name : '') + '" placeholder="如：张三" style="min-width:180px;"></div>' +
      '<div class="namerow"><span class="label">云账号</span><input id="oe-account" class="input" value="' + esc(o ? o.accountId : '') + '" placeholder="登录云端的 uid 或邮箱" style="min-width:260px;"></div>' +
      '<div class="namerow"><span class="label">角色</span><select id="oe-role" class="select" style="min-width:180px;">' +
      '<option value="">（未分配）</option>' +
      work.roles.map(function (r) { return '<option value="' + esc(r.id) + '"' + (o && r.id === o.roleId ? ' selected' : '') + '>' + esc(r.name) + '</option>'; }).join('') +
      '</select></div>' +
      '<div class="foot"><button type="button" class="btn btn-primary" id="oe-save">保存操作员</button><button type="button" class="btn" id="oe-cancel">取消修改</button></div>';
    var body = $('editor-body');
    body.innerHTML = html;

    $('oe-save').onclick = function () {
      var name = $('oe-name').value.trim();
      if (!name) { toast('请填写姓名', 'err'); return; }
      var account = $('oe-account').value.trim();
      var roleId = $('oe-role').value;
      if (isNew) {
        var no = { id: 'op_' + Date.now(), name: name, accountId: account, roleId: roleId || null };
        work.operators.push(no);
        selOp = no.id;
      } else {
        o.name = name; o.accountId = account; o.roleId = roleId || null;
      }
      renderOpList(); renderCurrentOp(); renderEditor();
      toast('操作员已保存（点「保存配置」生效）');
    };
    $('oe-cancel').onclick = function () {
      if (isNew) { selOp = null; renderOpList(); renderEditor(); }
      else { renderOpList(); renderEditor(); toast('已放弃该操作员的未保存修改'); }
    };
    var del = $('oe-del');
    if (del) del.onclick = function () { deleteOp(selOp); };
  }

  function deleteOp(id) {
    if (!id) return;
    if (!confirm('确认删除该操作员档案？')) return;
    work.operators = work.operators.filter(function (o) { return o.id !== id; });
    if (localStorage.getItem('acl_current_operator') === id) localStorage.removeItem('acl_current_operator');
    selOp = null;
    renderOpList(); renderCurrentOp(); renderEditor();
    toast('操作员已删除（点「保存配置」生效）');
  }

  /* ---------- 左侧「＋ 新建」入口 ---------- */
  function newRole() { selRole = NEW; selOp = null; renderRoleList(); renderOpList(); renderEditor(); var n = $('re-name'); if (n) n.focus(); }
  function newOp() { selOp = NEW; selRole = null; renderRoleList(); renderOpList(); renderEditor(); var n = $('oe-name'); if (n) n.focus(); }
  $('role-new').onclick = newRole;
  $('role-new-mini').onclick = newRole;
  $('op-new').onclick = newOp;

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
