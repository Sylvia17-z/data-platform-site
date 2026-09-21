/* =====================================================================
 * 各部门业绩完成情况跟踪日报表 - 核心逻辑
 * 数据源：订单项目收退款表-明细（每日导出）+ 2026 VAB 会员列表（可选）
 * ===================================================================== */

// 0. 按需懒加载第三方库（避免打开页面时同步解析 xlsx/html2canvas/jszip 造成卡顿）
const _libCache = {};
function loadScript(src){
  return new Promise((resolve, reject) => {
    if(_libCache[src]) return resolve(true);
    const s = document.createElement('script');
    s.src = src;
    s.async = true;
    s.onload = () => { _libCache[src] = true; resolve(true); };
    s.onerror = () => reject(new Error('脚本加载失败: ' + src));
    document.head.appendChild(s);
  });
}
function ensureXlsx(){ return loadScript('assets/vendor/xlsx.full.min.js'); }
function ensureHtml2canvas(){ return loadScript('assets/vendor/html2canvas.min.js'); }
async function ensureXlsxStyled(){
  await loadScript('assets/vendor/xlsx.full.min.js');
  await loadScript('assets/vendor/jszip.min.js');
  await loadScript('assets/xlsx-styled.js?v=20260917-01');
  return true;
}

// 1. 渠道分类枚举（用于选项下拉）
const CHANNEL_CLASSES = ['营销部','商务部','市场部','公司所有','经营科室'];
const L1_CHANNELS = ['品牌组','活动组','电商组','短视频组','商保业务','商务直客','中医科','儿科','客户转介绍','企业团检','市场田梦代理人','商务渠道-韩通代理人','公司所有'];

// 月目标默认值（万元）- 来自 sheet2 默认值
const DEFAULT_TARGETS = {
  '品牌组-其他':0,'活动组-活动支持':0,
  '电商组':77,'短视频组及自然到店':3,
  '公司支持':360,
  '商务部':125,
  '其他':170,        // 原 155 + 市场部 15（市场部已并入"其他"，不再单列）
  '合计':360,
  '基础业绩':310,'新增业绩':50,'管家部':238
};
// 无业绩目标的部门（高新店 / 经开店）：不在「② 月目标」面板列出，也不在报表中显示目标
const NO_TARGET_KEYS = new Set(['高新','北郊']);
// ② 月目标面板分组（美化：按业务板块归类，便于核对）
const TARGET_GROUPS = [
  {title:'营销部板块', keys:['品牌组-其他','活动组-活动支持','电商组','短视频组及自然到店','公司支持']},
  {title:'商务部 / 其他', keys:['商务部','其他']},
  {title:'汇总校验', keys:['合计','基础业绩','新增业绩','管家部']},
];

// 管家部业绩口径（用户最终确认——纯按「开单咨询师」字段值判定，渠道不参与）：
//   开单咨询师 不为空 且 ≠ 无/许顺华/空/营销部/吴雅琴/张帆/王勤红 → 计入管家部
// 注意：「营销部」判断的是【开单咨询师字段值】=营销部，而不是任何渠道字段（客户来源一级/二级渠道等均不参与管家部判定）。
function normName(v){
  return String(v == null ? '' : v).trim().replace(/[ 　]+/g, ' ').toLowerCase();
}
const STEWARD_EXCLUDE = new Set(['无','许顺华','空','营销部','吴雅琴','张帆','王勤红'].map(normName));

// 一条订单是否属于管家部：开单咨询师 非空 且 不在排除名单（大小写/全半角空格容错）
function isSteward(r){
  const c = normName(r.consultant);
  if(!c) return false;                       // 空 / 无 → 排除
  if(STEWARD_EXCLUDE.has(c)) return false;   // 排除名单（含"营销部"作为开单咨询师取值）
  return true;                               // 其余非空咨询师 → 计入管家部
}

// 二级渠道匹配包含关系
function includesAny(haystack, needles){
  if(!haystack) return false;
  return needles.some(n => String(haystack).includes(n));
}

// 状态
const state = {
  orders: [],      // 解析后的订单明细
  vabSet: new Set(),// VAB 会员号（规范化集合，含小写/仅数字/去前导零等容错形态）
  vabList: [],     // VAB 会员号原始字符串（可读清单，去重；用于页面展示，不混入容错形态）
  vabRows: [],     // VAB 完整行（导入表格的全部字段，用于按原表字段展示）
  demoOrders: false, // 当前订单是否为自动加载的示例（上传真实文件后应清除）
  demoVab: false,    // 当前 VAB 是否为自动加载的示例
  targets: {...DEFAULT_TARGETS},
  cutoffDate: null,// 截止日期
  results: null,   // 最新计算结果
  colWidths: {},   // 用户拖拽调整的列宽（index -> width）
  manual: {        // 手动修改的数据（覆盖计算值）
    month: {},     // month[rowKey] = 手动月累计值
    daily: {},     // daily[dateStr][rowKey] = 手动每日值
  },
};

// 订单稳定主键（_uid）：用于增量合并/删除时精确定位行，避免按内容匹配出错。
// 必须在使用它的 restoreOrders()（脚本加载时顶层调用）之前声明，否则 nextUid() 在
// _orderUidSeq 初始化前触发 TDZ 报错，会被 restoreOrders 的 catch 静默吞掉，导致刷新丢数据。
let _orderUidSeq = 0;
function nextUid(){ return 'ord_' + (++_orderUidSeq); }
function ensureUid(r){ if(!r._uid) r._uid = nextUid(); return r; }

// 手动修改数据持久化（localStorage：月目标 / 手动覆盖值）
const TARGETS_STORAGE_KEY = 'dailyPerformance_targets';
const MANUAL_STORAGE_KEY = 'dailyPerformance_manual';
function persistTargets(){ try{ localStorage.setItem(TARGETS_STORAGE_KEY, JSON.stringify(state.targets)); }catch(e){} }
function restoreTargets(){ try{ const raw = localStorage.getItem(TARGETS_STORAGE_KEY); if(raw) state.targets = {...DEFAULT_TARGETS, ...JSON.parse(raw)}; }catch(e){} }
function persistManual(){ try{ localStorage.setItem(MANUAL_STORAGE_KEY, JSON.stringify(state.manual)); }catch(e){} }
function restoreManual(){ try{ const raw = localStorage.getItem(MANUAL_STORAGE_KEY); if(raw){ const m = JSON.parse(raw); state.manual.month = m.month || {}; state.manual.daily = m.daily || {}; } }catch(e){} }

// 历史导入订单本地持久化（localStorage：刷新网页后自动恢复；月累计/每日数据均由订单重算得到）
// 存储订单原始行（含 _uid 稳定主键）+ 截止日期；不存派生结果，刷新后由 recompute() 重新生成，保证一致。
const ORDERS_STORAGE_KEY = 'dailyPerformance_orders';
// 列宽持久化键：必须在 restoreColWidths()（脚本加载时顶层调用）之前声明，否则与订单 TDZ 同理，
// const 在初始化前被访问会抛 ReferenceError，被恢复函数的 catch 静默吞掉 → 拖拽列宽刷新后不保留。
const COLW_STORAGE_KEY = 'dailyPerformance_colw';
function persistOrders(){
  try{ localStorage.setItem(ORDERS_STORAGE_KEY, JSON.stringify({ orders: state.orders, cutoff: state.cutoffDate || null })); }
  catch(e){ /* 配额超限等：静默降级，下次导入/删除再尝试保存 */ }
}
function restoreOrders(){
  try{
    const raw = localStorage.getItem(ORDERS_STORAGE_KEY);
    if(!raw) return;
    const obj = JSON.parse(raw);
    if(obj && Array.isArray(obj.orders)){
      // JSON 不保留 Date：恢复时把 payTime 字符串还原为 Date（fmtDate/mergeOrdersByDate 均依赖 Date）
      state.orders = obj.orders.map(r => {
        if(r && typeof r.payTime === 'string') r.payTime = new Date(r.payTime);
        if(r && !r._uid) r._uid = nextUid();
        return r;
      });
      if(obj.cutoff) state.cutoffDate = obj.cutoff;
    }
  }catch(e){}
}
function clearManualOverrides(){
  state.manual = { month:{}, daily:{} };
  persistManual();
  renderReport();
  showToast('已清除所有手动调整，月累计/每日数据恢复为自动计算值', 'ok');
}
restoreTargets();
restoreManual();

// 手动编辑单元格：双击任意数值单元格进入编辑，Enter/失焦保存，Esc 取消
// 单元格编辑：支持四则运算表达式（如 1000+500、2000-300*2、4000/2）与相对增减（+500 / -300 表示在基准值上增减）。
// 仅允许数字与 + - * / ( ) . 字符，杜绝代码注入；结果非有限数时返回 NaN 由调用方降级为 0。
function evalArithmetic(raw, base){
  const s = String(raw).replace(/[,¥元\s]/g, '').trim();
  if(s === '') return NaN;
  let expr = s;
  // 以 +/- 开头 → 在基准值上增减（如单元格原值 4000，输入 +500 → 4500）
  if(/^[+\-]/.test(s)){
    const b = Number(base);
    expr = (isNaN(b) ? 0 : b) + s;
  }
  if(!/^[0-9+\-*/().]+$/.test(expr)) return NaN;
  try {
    const r = Function('"use strict";return (' + expr + ');')();
    return (typeof r === 'number' && isFinite(r)) ? r : NaN;
  } catch(e){ return NaN; }
}

function makeEditable(el, onSave, opts={}){
  el.addEventListener('dblclick', ()=>{
    if(el.querySelector('input')) return;
    const oldVal = el.textContent.trim().replace(/[,¥元]/g,'');
    const input = document.createElement('input');
    input.type = 'text';           // 不显示上下调整(原生 spinner)按钮，直接手动输入
    input.inputMode = 'decimal';   // 移动端弹出数字键盘
    input.value = oldVal || '0';
    input.title = '支持四则运算，如 1000+500；以 +/- 开头表示在现值的基数上增减';
    input.style.cssText = 'width:92%;text-align:right;border:1px solid var(--c-border);border-radius:4px;padding:1px 3px;font-size:12px;';
    el.textContent = '';
    el.appendChild(input);
    input.focus();
    input.select();
    let done = false;
    const finish = (save)=>{
      if(done) return;
      done = true;
      const raw = input.value.trim();
      if(save && raw !== ''){
        const v = evalArithmetic(raw, oldVal);
        onSave(isNaN(v) ? 0 : v);
      }
    };
    input.addEventListener('keydown', e=>{
      if(e.key === 'Enter'){ finish(true); }
      else if(e.key === 'Escape'){ finish(false); }
    });
    input.addEventListener('blur', ()=>finish(true));
  });
}

// VAB 会员本地持久化（localStorage，无需每次重新导入；上传一次即为默认，可覆盖更新）
const VAB_STORAGE_KEY = 'dailyPerformance_vab';
let _vabStorageWarned = false;
function persistVab(){
  try{
    // set 保存规范化集合用于匹配；list 保存原始会员号（展示/去重）；rows 保存导入表格全部字段用于按原表字段展示
    localStorage.setItem(VAB_STORAGE_KEY, JSON.stringify({
      set:[...state.vabSet],
      list: state.vabList || [],
      rows: state.vabRows || [],
      demo: !!state.demoVab
    }));
  }catch(e){
    if(!_vabStorageWarned){
      _vabStorageWarned = true;
      showToast('⚠ 本地保存不可用（可能是 file:// 或隐私模式），VAB 不会跨会话保留；建议用本地服务器 http://localhost:8800 打开', 'warn', 5200);
    }
    console.warn('VAB 保存失败', e);
  }
}
// 用完整行对象集合刷新 VAB 状态：vabRows(展示全部字段) / vabList(会员号) / vabSet(匹配) 三者同步
function setVabRows(rows){
  const clean = (rows || []).filter(r => r && typeof r === 'object');
  state.vabRows = clean;
  state.vabList = clean.map(normalizeVab).filter(Boolean);
  state.vabSet = buildVabSet(state.vabList);
  state.demoVab = false;
}
function restoreVab(){
  try{
    const raw = localStorage.getItem(VAB_STORAGE_KEY);
    if(!raw) return;
    const data = JSON.parse(raw);
    if(data && Array.isArray(data.rows) && data.rows.length){
      setVabRows(data.rows);                       // 新格式：保留全部字段
    }else if(data && Array.isArray(data.list) && data.list.length){
      setVabRows(data.list.map(x => ({'会员号': x}))); // 兼容旧格式（仅会员号）
    }else if(Array.isArray(data)){                  // 极旧纯数组格式
      setVabRows(data.map(x => ({'会员号': x})));
    }
    if(data && typeof data.demo === 'boolean') state.demoVab = data.demo;
  }catch(e){ /* 忽略损坏数据 */ }
}
// 页面加载时恢复本地 VAB
restoreVab();
restoreColWidths(); // 恢复列宽（拖拽调整后保留）
restoreOrders(); // 恢复历史导入订单（刷新网页后保留，含月累计/每日所需原始数据）
// 统一刷新 VAB 状态文案（示例 vs 已保存默认，且可覆盖更新）
function refreshVabStatus(){
  const meta = $('vab-meta');
  if(!meta) return;
  if(vabUniqueCount() === 0){
    meta.textContent = '⚠ 未上传 VAB 名单 · 「新增业绩」将按 0 计，请上传真实会员表';
    meta.className = 'chip warn';
  }else if(state.demoVab){
    meta.textContent = `⚠ 示例VAB（${vabUniqueCount()}人）· 请上传真实会员表，上传后即为默认`;
    meta.className = 'chip warn';
  }else{
    meta.textContent = `✔ 已保存为默认（${vabUniqueCount()}人）· 重新上传/编辑可覆盖更新`;
    meta.className = 'chip ok';
  }
  renderVabList(); // 同步刷新页面上的会员清单
}
// 在页面上渲染 VAB 会员清单（按导入表格的全部字段展示），默认展开呈现，支持单元格双击编辑 + 行内删除 + 页面内新增
function renderVabList(){
  const box = $('vab-list');
  if(!box) return;
  const rows = state.vabRows || [];
  // summary 显示人数（折叠态也可见）：用去重后的有效会员数，与状态条保持一致
  const sum = $('vab-list-summary');
  if(sum) sum.textContent = `📋 VAB 会员清单（共 ${vabUniqueCount()} 人）`;
  // 动态列头：取所有行键的并集，保持首次出现顺序（按原表字段显示）
  const headers = [];
  rows.forEach(r => Object.keys(r).forEach(k => { if(!headers.includes(k)) headers.push(k); }));
  const headHtml = headers.length
    ? '<tr>' + headers.map(h=>`<th>${esc(h)}</th>`).join('') + '<th style="width:44px;text-align:center;">操作</th></tr>'
    : '';
  const bodyHtml = rows.length
    ? rows.map((r,i)=>{
        const m = normalizeVab(r);
        const cells = headers.map(h=>`<td class="vab-cell" data-ri="${i}" data-field="${esc(h)}" title="双击编辑该字段">${esc(r[h] ?? '')}</td>`).join('');
        return `<tr data-m="${esc(m)}">${cells}<td style="text-align:center;"><button class="vab-del" type="button" title="删除该会员">×</button></td></tr>`;
      }).join('')
    : '';
  const tableBlock = rows.length
    ? `<div class="table-wrap vab-table-wrap" style="max-height:320px;overflow:auto;border:1px solid var(--c-border);border-radius:8px;margin:8px 0;">
         <table class="data"><thead>${headHtml}</thead><tbody>${bodyHtml}</tbody></table></div>`
    : `<div class="file-meta" style="margin:8px 0;">（暂无 VAB 会员，点「＋ 新增会员」弹窗录入，或重新上传会员表）</div>`;
  box.innerHTML =
    tableBlock +
    `<div class="vab-add-row">` +
      `<button id="vab-add-btn" class="btn btn-primary" type="button">＋ 新增会员</button>` +
      `<span class="file-meta">弹窗录入会员号/姓名/组别/手机号等全部字段</span>` +
    `</div>`;
  // 事件委托只需绑定一次（容器常驻，内部元素随 innerHTML 重建）
  if(!box.dataset.bound){
    box.addEventListener('click', onVabListClick);
    box.addEventListener('dblclick', onVabListDblClick);
    box.dataset.bound = '1';
  }
}
// 列表双击：进入单元格编辑（事件委托到 #vab-list 容器）
function onVabListDblClick(e){
  const td = e.target.closest && e.target.closest('td.vab-cell');
  if(td) startVabCellEdit(td);
}
// 单元格就地编辑：双击单元格 → 显示输入框 → 回车/失焦保存，Esc 取消；保存到 state.vabRows 并持久化
function startVabCellEdit(td){
  if(td.querySelector('input')) return;
  const ri = Number(td.dataset.ri);
  const field = td.dataset.field;
  const row = state.vabRows[ri];
  if(!row) return;
  const oldVal = row[field] ?? '';
  const input = document.createElement('input');
  input.type = 'text';
  input.value = oldVal;
  input.title = '编辑该字段，回车或失焦保存，Esc 取消';
  input.style.cssText = 'width:96%;border:1px solid var(--c-accent);border-radius:4px;padding:1px 3px;font-size:12px;box-sizing:border-box;';
  td.textContent = '';
  td.appendChild(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (save) => {
    if(done) return;
    done = true;
    if(save){
      const v = input.value;
      state.vabRows[ri][field] = v;
      setVabRows(state.vabRows);
      persistVab();
      recompute();
      refreshVabStatus(); // 更新状态条并刷新清单（含新值）
      showToast(`已更新会员「${field}」：${oldVal} → ${v}`, 'ok');
    }else{
      renderVabList(); // 取消：恢复原显示
    }
  };
  input.addEventListener('keydown', e => {
    if(e.key === 'Enter'){ e.preventDefault(); finish(true); }
    else if(e.key === 'Escape'){ e.preventDefault(); finish(false); }
  });
  input.addEventListener('blur', () => finish(true));
}
// 列表点击：删除按钮 / 新增按钮（事件委托到 #vab-list 容器）
function onVabListClick(e){
  const del = e.target.closest && e.target.closest('.vab-del');
  if(del){ const tr = del.closest('tr'); deleteVabMember(tr && tr.dataset.m); return; }
  if(e.target.id === 'vab-add-btn'){ openAddVabModal(); }
}
// 新增会员改为弹窗（见 openAddVabModal / saveAddVabModal），不再内联输入会员号
// 新增会员弹窗：按当前 VAB 字段模板生成输入项，可录入会员号/姓名/组别等全部字段
function buildVabHeaders(){
  const set = [];
  (state.vabRows || []).forEach(r => Object.keys(r).forEach(k => { if(!set.includes(k)) set.push(k); }));
  if(!set.includes('会员号')) set.unshift('会员号'); // 会员号始终为首项
  return set;
}
function openAddVabModal(){
  const fields = buildVabHeaders();
  const box = $('vab-add-fields');
  if(box) box.innerHTML = fields.map(f =>
    `<div class="modal-field"><label>${esc(f)}</label><input class="modal-input" data-field="${esc(f)}" type="text" placeholder="${esc(f)}" /></div>`
  ).join('');
  const m = $('vab-add-modal');
  if(m) m.style.display = 'flex';
  const first = box && box.querySelector('input');
  if(first) first.focus();
}
function closeAddVabModal(){
  const m = $('vab-add-modal');
  if(m) m.style.display = 'none';
}
// 保存弹窗新增的会员（按字段构建整行对象，去重、持久化、联动统计）
function saveAddVabModal(){
  const box = $('vab-add-fields');
  if(!box) return;
  const row = {};
  box.querySelectorAll('input').forEach(inp => {
    const f = inp.dataset.field;
    const v = (inp.value || '').trim();
    if(v !== '') row[f] = v;
  });
  if(!row['会员号']){
    showToast('会员号不能为空', 'warn'); return;
  }
  if(state.vabList.some(x => x.toLowerCase() === String(row['会员号']).toLowerCase())){
    showToast('该会员已存在：'+row['会员号'], 'warn'); return;
  }
  state.vabRows.push(row);
  setVabRows(state.vabRows);
  persistVab();
  recompute();
  refreshVabStatus(); // 刷新状态条与清单（含新行）
  closeAddVabModal();
  showToast('已新增会员：'+row['会员号'], 'ok');
}
// 页面内删除单个会员（同步集合与持久化、刷新统计）
function deleteVabMember(m){
  if(!m) return;
  const key = String(m).toLowerCase();
  state.vabRows = state.vabRows.filter(r => String(normalizeVab(r)||'').toLowerCase() !== key);
  setVabRows(state.vabRows);
  persistVab(); refreshVabStatus(); recompute();
  showToast('已删除会员：'+m, 'ok');
}
// 清空整个 VAB 清单（用于重置旧版本残留的示例/历史数据，清空后重新上传真实会员表即可覆盖）
function clearVabList(){
  if(vabUniqueCount() === 0){ showToast('当前 VAB 清单已为空', 'warn'); return; }
  if(!window.confirm('确定清空 VAB 会员清单（' + vabUniqueCount() + ' 人）？清空后可重新上传真实会员表。')) return;
  state.vabRows = [];
  state.vabList = [];
  state.vabSet = new Set();
  state.demoVab = false;
  persistVab(); refreshVabStatus(); recompute();
  showToast('已清空 VAB 清单，请重新上传真实会员表', 'ok');
}

// 清除本机保存的全部页面数据（订单/VAB会员/月目标/手动调整/列宽），重置为空
function clearAllData(){
  const hasData = state.orders.length > 0
    || Object.keys(state.manual.month).length > 0 || Object.keys(state.manual.daily).length > 0
    || Object.keys(state.colWidths).length > 0;
  if(!hasData){ showToast('当前本机数据已为空，无需清除', 'warn'); return; }
  if(!window.confirm('确定清除本机保存的全部页面数据？\n（订单、月目标、手动调整、列宽将全部清空，刷新后恢复为空；VAB会员不受影响。此操作不可撤销）')) return;
  try{
    localStorage.removeItem(ORDERS_STORAGE_KEY);
    localStorage.removeItem(MANUAL_STORAGE_KEY);
    localStorage.removeItem(TARGETS_STORAGE_KEY);
    localStorage.removeItem(COLW_STORAGE_KEY);
  }catch(e){}
  state.orders = [];
  state.manual = { month:{}, daily:{} };
  state.targets = Object.assign({}, DEFAULT_TARGETS);
  state.results = null; state.dailyResults = []; state.colWidths = {};
  if($('cutoff-date')) $('cutoff-date').value = todayIso();
  renderReport(); renderStoreReport([], []); renderImportDetails(); renderTargets();
  if($('orders-meta')){ $('orders-meta').textContent = '未导入'; $('orders-meta').className = 'file-meta'; }
  if($('data-status')){ $('data-status').textContent = '待导入数据'; $('data-status').className = 'chip'; }
  showToast('已清除本机全部页面数据（订单/月目标/手动调整/列宽），VAB会员已保留', 'ok', 3600);
}

// 工具
const $ = (id) => document.getElementById(id);
const fmt = (n, d=2) => (n==null||isNaN(n)) ? '0' : Number(n).toLocaleString('zh-CN',{minimumFractionDigits:d,maximumFractionDigits:d});
const fmtInt = (n) => (n==null||isNaN(n)) ? '0' : Math.round(Number(n)).toLocaleString('zh-CN');
// 复制专用：无千分位、无强制小数位，便于 Excel 粘贴为数值
const fmtPlain = (n) => (n==null||isNaN(n)) ? '0' : String(Math.round(Number(n)*100)/100);
// HTML 转义：订单字段(name/memberNo 等)经 innerHTML 渲染前必须转义，防止自注入
const esc = (s) => String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
const todayIso = () => new Date().toISOString().slice(0,10);
function showToast(msg, type='ok', timeout=2400){
  const t = $('toast');
  t.textContent = msg;
  t.className = 'toast show ' + (type==='ok'?'ok':type==='err'?'err':'');
  clearTimeout(showToast._t);
  showToast._t = setTimeout(()=>{ t.className='toast'; }, timeout);
}

// 颜色提示
function tagClassFor(arr, tag){
  const tagOk = tag==='ok'; const tagErr = tag==='err';
  return tagOk ? 'chip ok' : tagErr ? 'chip err' : 'chip warn';
}

// 2. Excel 解析
async function parseExcelFile(file){
  const buf = await file.arrayBuffer();
  // 注意：cellDates:true 会把 Excel 日期序列按 UTC 解读，在 +8 时区含时间(尤其深夜)的订单会被偏移一天，
  // 导致"当日"列与月累计归属错位。改为 cellDates:false（raw:false 已是格式化字符串），
  // 日期单元格以"yyyy-mm-dd hh:mm:ss"字符串返回，由 parseDateLoose 按本地时区解析，避免跨天。
  const wb = XLSX.read(buf, {type:'array', cellDates:false});
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, {defval:'', raw:false});
  return rows;
}

// 3. 列名归一化（应对表格字段名变化）
function pickField(row, candidates){
  for(const k of Object.keys(row)){
    for(const c of candidates){
      if(String(k).replace(/\s|（实时|（\u5b9e\u65f6）/g,'').includes(c)) return row[k];
    }
  }
  return '';
}

function normalizeOrder(r){
  return {
    name: pickField(r, ['客户姓名','客户']),
    // 会员号列名兼容多种写法：会员号 / 会员编号 / 会员ID / 会员卡号 / 卡号 / VAB会员号
    // 若真实订单表用"会员编号"等命名而只匹配"会员号"，memberNo 会为空 → 全部 VAB 匹配失败 → 新增业绩恒为 0
    memberNo: pickField(r, ['会员号','会员编号','会员ID','会员卡号','卡号','VAB会员号','member_no']),
    developer: pickField(r, ['客户开发人']),
    consultant: pickField(r, ['开单健康管家','开单咨询师','咨询师','服务顾问','开单人','主诊咨询师']),
    channelClass: pickField(r, ['客户来源渠道分类','渠道分类']),
    channelL1: pickField(r, ['客户来源一级渠道','一级渠道']),
    channelL2: pickField(r, ['客户来源二级渠道','二级渠道']),
    channelL3: pickField(r, ['客户来源三级渠道','三级渠道','渠道L3','客户来源三级','三级来源']),
    cashPay: Number(pickField(r, ['现款支付'])) || 0,
    visitType: pickField(r, ['到访类型']),
    payTime: parseDateLoose(pickField(r, ['收款时间'])),
    receiptNo: pickField(r, ['收款单号']),
    cashier: pickField(r, ['收款员工']),
    payType: pickField(r, ['收款类型']),
    payMethod: pickField(r, ['收款详情汇总','收款方式','详情']),
  };
}

function normalizeVab(r){
  return String(pickField(r, ['会员号','VAB会员号','VAB','会员编号','会员卡号','卡号','member_no'])).trim();
}

// 构建 VAB 匹配集合：同时收录「原值(去空格/小写)」「仅数字」「仅数字·去前导零」
// 三种形态，以容忍订单会员号与 VAB 表之间的前缀(V)/前导零(0263..)/大小写差异
function buildVabSet(arr){
  const s = new Set();
  for(const x of arr){
    const n = String(x == null ? '' : x).trim().toLowerCase();
    if(!n) continue;
    s.add(n);
    const d = n.replace(/\D/g, '');
    if(d){
      s.add(d);
      const ds = d.replace(/^0+/, ''); // 去前导零，容忍 026374801 / 26374801
      if(ds) s.add(ds);
    }
  }
  return s;
}

// 去重后的有效 VAB 会员人数（用户-facing 的「人」数，不含空值、重复、规范化形态扩张）
function vabUniqueCount(){
  return new Set(state.vabList).size;
}

function parseDateLoose(v){
  if(!v) return null;
  if(v instanceof Date) return v;
  // 字符串
  const s = String(v).trim();
  // 中文日期时间：2026年9月10日 10:00:00 / 2026年9月10日-10:00
  let m = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日[-\sT]?(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if(m) return new Date(+m[1], +m[2]-1, +m[3], +(m[4]||0), +(m[5]||0), +(m[6]||0));
  // yyyy-mm-dd hh:mm:ss 或 yyyy-mm-dd-hh:mm:ss（横杠连接时间，用户文件常见）
  m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})[-\sT](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if(m) return new Date(+m[1], +m[2]-1, +m[3], +(m[4]||0), +(m[5]||0), +(m[6]||0));
  // yyyy-mm-dd 或 yyyy/mm/dd（无时间）
  m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if(m) return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]));
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}

function inDateRange(d, cutoff){
  if(!d) return false;
  if(cutoff){
    const c = new Date(cutoff);
    c.setHours(23,59,59,999);
    return d <= c;
  }
  return true;
}

// 4. 业绩计算引擎 - 实现全部15条规则
function calculate(orders, vabSet, cutoff){
  // 月累计口径 = 截止日所在"本月"累计（当月1日 ~ 截止日），不含历史月份数据
  // 否则混入示例/往期数据会让本月月累计被撑大（与"月目标"语义不匹配）
  const cy = cutoff.getFullYear(), cm = cutoff.getMonth();
  orders = orders.filter(r => {
    if(!r.payTime) return false;
    const d = r.payTime;
    if(d.getFullYear() !== cy || d.getMonth() !== cm) return false; // 仅统计截止月
    return inDateRange(d, cutoff);
  });

  const sum = (rows) => rows.reduce((a,b)=>a + (Number(b.cashPay)||0), 0);

  // --- 1. 品牌组-其他 ---
  const r1 = orders.filter(r =>
    r.channelClass === '营销部' &&
    r.channelL1 === '品牌组' &&
    ['微信小程序','微信公众号','客服手机','前台座机'].includes(r.channelL2) ||
    (r.channelClass === '营销部' && r.channelL1 === '品牌组' && (r.channelL2||'').includes('阿里') && (r.channelL2||'').includes('高德'))
  );

  // --- 2. 活动组-活动支持 ---
  const r2 = orders.filter(r => r.channelClass==='营销部' && r.channelL1==='活动组' && r.channelL2==='活动支持');

  // --- 3. 电商组（接受 一级渠道=电商组 / 客户转介绍 / 企业团检） ---
  const r3 = orders.filter(r => r.channelClass==='营销部' && ['电商组','客户转介绍','企业团检'].includes(r.channelL1));

  // --- 4. 短视频组 ---
  const r4 = orders.filter(r => r.channelClass==='营销部' && r.channelL1==='短视频组');

  // --- 5. 品牌组-自然到店（包含自然到店关键字） ---
  const r5 = orders.filter(r => r.channelClass==='营销部' && r.channelL1==='品牌组' && (r.channelL2||'').includes('自然到店'));

  // --- 6. 商务部（口径：客户来源渠道分类/一级渠道/二级渠道 任一为「商务部」或「市场部」；
  //      市场部数据并入商务部统一核算，故不再单独落入"其他"） ---
  //      注：渠道分类=推荐 的订单改由下方 r6b 单独处理，避免与 chanHas 商务部/市场部 重复计入，
  //      故此处先把推荐类排除，r6b 再并入商务部，二者互不重叠。
  const r6 = orders.filter(r => (chanHas(r, '商务部') || chanHas(r, '市场部')) && (r.channelClass||'') !== '推荐');

  // --- 6b. 推荐渠道补充（新增逻辑，不动 r1~r6 原有规则）---
  // 渠道分类=推荐 且 一级分类包含「营销部」→ 计入营销部业绩：
  //   二级含「自然到店」→ 品牌组-自然到店(r5b)，否则 → 电商组(r3b)
  // 渠道分类=推荐 且 一级分类包含「商务部」→ 计入商务部业绩(r6b)
  const r3b = orders.filter(r => isRecToYingxiao(r) && !(r.channelL2||'').includes('自然到店'));
  const r5b = orders.filter(r => isRecToYingxiao(r) && (r.channelL2||'').includes('自然到店'));
  const r6b = orders.filter(r => isRecToShangwu(r));

  // --- 7. 合计（全部行） ---
  const rAll = orders;

  // --- 9. 新增业绩（VAB会员） ---
  // 会员号匹配：精确 -> 仅数字 -> 仅数字·去前导零 三级容错，
  // 容忍订单会员号与 VAB 表之间的前缀(如 V25062001225)/前导零/大小写差异
  const r9 = orders.filter(r => {
    if(!r.memberNo) return false;
    const raw = String(r.memberNo).trim().toLowerCase();
    if(!raw) return false;
    if(vabSet.has(raw)) return true;
    const digits = raw.replace(/\D/g, '');
    if(digits){
      if(vabSet.has(digits)) return true;
      const ds = digits.replace(/^0+/, '');
      if(ds && vabSet.has(ds)) return true;
    }
    return false;
  });

  // --- 10. 基础业绩 = 合计 - 新增 ---
  const sumAll = sum(rAll);
  const sumVab = sum(r9);
  const sumBasic = sumAll - sumVab;

  // --- 11. 管家部：开单咨询师 不为空 且 不属于排除名单；开单咨询师为空时按渠道兜底（见 isSteward） ---
  const r11 = orders.filter(r => isSteward(r));

  // --- 14. 北郊：收款员工≠南媛 ---
  const r14 = orders.filter(r => (r.cashier||'').toString().trim() !== '南媛');
  // --- 15. 高新 = 合计 - 北郊（即收款员工=南媛） ---
  const rGaoXin = orders.filter(r => (r.cashier||'').toString().trim() === '南媛');

  // --- 15. 高新 = 合计 - 北郊 ---
  const sumBeiJiao = sum(r14);
  const sumGaoXin = sumAll - sumBeiJiao;
  const sumStoreTotal = sumGaoXin + sumBeiJiao; // 门店合计 = 高新店 + 经开店

  // --- 8. 其他 = 合计 - (品牌组-其他 + 活动组-活动支持 + 电商组 + 短视频组 + 品牌组-自然到店 + 商务部) ---
  // 注：市场部已并入"商务部"统一核算，故"其他"的扣减项里的「商务部」已含市场部金额，不会重复计入。
  // 注：原"活动组-自拓活动"板块已整体下线（不展示、不计算），其金额并入"其他"
  const sumBrandOther = sum(r1);
  const sumActSupport = sum(r2);
  const sumEcom = sum(r3) + sum(r3b);
  const sumShort = sum(r4);
  const sumNature = sum(r5) + sum(r5b);
  const sumCommerce = sum(r6) + sum(r6b);
  const sumOther = sumAll - (sumBrandOther + sumActSupport + sumEcom + sumShort + sumNature + sumCommerce);

  return {
    rows: {
      '品牌组-其他': r1, '活动组-活动支持': r2,
      '电商组': [...r3, ...r3b], '短视频组': r4,
      '品牌组-自然到店': [...r5, ...r5b],
      '商务部': [...r6, ...r6b], '其他': orders.filter(()=>true),
      '合计': rAll, '新增业绩': r9, '基础业绩': orders, '管家部': r11,
      '北郊': r14, '高新': rGaoXin,
    },
    values: {
      '品牌组-其他': sumBrandOther,
      '活动组-活动支持': sumActSupport,
      '电商组': sumEcom,
      '短视频组': sumShort,
      '品牌组-自然到店': sumNature,
      '商务部': sumCommerce,
      '其他': sumOther,
      '合计': sumAll,
      '新增业绩': sumVab,
      '基础业绩': sumBasic,
      '管家部': sum(r11),
      '北郊': sumBeiJiao,
      '高新': sumGaoXin,
      '门店合计': sumStoreTotal,
    }
  };
}

// 5. 报表渲染：表结构完全对齐截图 + 原表配色
const REPORT_TEMPLATE = [
  // 营销部（rowspan 5 = 公司支持2 + 部门业绩3）
  // 月目标：公司支持块2行共享360；部门业绩块下 电商组独立(77)，短视频组+品牌组-自然到店 合并为一格(3, rowspanTarget:2)
  {main:'营销部', sub:'公司支持', item:'品牌组-其他', targetKey:'品牌组-其他', rowspanMain:5, rowspanSub:2, rowspanTarget:2, indent:1},
  {main:'',       sub:'',        item:'活动组-活动支持', targetKey:'活动组-活动支持', indent:1},
  {main:'',       sub:'部门业绩', item:'电商组', targetKey:'电商组', rowspanSub:3, indent:1},
  {main:'',       sub:'',        item:'短视频组', targetKey:'短视频组及自然到店', valueKey:'短视频组', rowspanTarget:2, indent:1},
  {main:'',       sub:'',        item:'品牌组-自然到店', targetKey:'短视频组及自然到店', valueKey:'品牌组-自然到店', indent:1},
  {main:'商务部', sub:'',        item:'', targetKey:'商务部', mergedLabel:true},
  {main:'其他',   sub:'',        item:'', targetKey:'其他', mergedLabel:true},
  {main:'合计',   sub:'',        item:'', targetKey:'合计', mergedLabel:true, highlight:true},
  {main:'基础业绩', sub:'（新增业绩之外）', item:'', targetKey:'基础业绩', mergedLabel:true},
  {main:'新增业绩', sub:'（2026新增VAB）', item:'', targetKey:'新增业绩', mergedLabel:true},
  {main:'管家部',   sub:'',        item:'', targetKey:'管家部', mergedLabel:true},
];

function renderTargets(){
  const el = $('targets');
  const html = TARGET_GROUPS.map(g=>{
    const items = g.keys.filter(k => (k in DEFAULT_TARGETS) && !NO_TARGET_KEYS.has(k));
    if(!items.length) return '';
    const cells = items.map(k=>`
      <div class="target-row">
        <span class="label">${k}</span>
        <input class="input target-input" data-key="${k}" type="text" inputmode="decimal" value="${state.targets[k] ?? DEFAULT_TARGETS[k]}" />
        <span class="unit">万元</span>
      </div>`).join('');
    return `<div class="target-group"><div class="target-group__title">${g.title}</div><div class="target-group__grid">${cells}</div></div>`;
  }).filter(Boolean).join('');
  el.innerHTML = html;
  el.querySelectorAll('.target-input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      state.targets[e.target.dataset.key] = Number(e.target.value) || 0;
      persistTargets();
      renderReport();
    });
  });
}

function readTargets(){
  document.querySelectorAll('.target-input').forEach(inp=>{
    state.targets[inp.dataset.key] = Number(inp.value)||0;
  });
}

function renderReport(){
  try {
    renderReportInner();
    // 渲染成功时清除错误条
    const errEl = $('render-error');
    if(errEl) errEl.style.display = 'none';
  } catch(e) {
    console.error('renderReport 异常:', e);
    const msg = (e && e.message) || String(e);
    $('report-table').querySelector('thead').innerHTML = '<tr><th colspan="6" class="title-cell" style="background:#b3261e;color:#fff;">⚠ 渲染异常：'+esc(msg)+'</th></tr>';
    $('report-body').innerHTML = '<tr><td colspan="6" style="color:#b3261e;padding:24px;text-align:center;">'+esc(msg)+'<br><br>请检查数据日期格式（如"2026-08-15 10:00:00"）或刷新页面重试。</td></tr>';
    // 在数据导入区也显示错误条
    const errEl = $('render-error');
    if(errEl){
      errEl.textContent = '⚠ 报表渲染异常：'+msg+'（已打开浏览器控制台查看详情）';
      errEl.style.display = '';
    }
    showToast('报表渲染异常：'+msg, 'err');
  }
}

// 派生行联动：在「计算结果 + 手动覆盖」基础上，重算派生行，使修改任一输入后相关数据同步更新
// - 优先采用手动覆盖值（state.manual.month / state.manual.daily[date]）
// - 未被手动覆盖的派生行按公式重算：
//     基础业绩 = 合计 - 新增业绩
//     其他     = 合计 - (品牌组-其他+活动组-活动支持+电商组+短视频组+品牌组-自然到店+商务部)
//     高新     = 顶层合计 - 北郊   （主表顶层=合计；门店表顶层=门店合计）
//     门店合计 = 高新 + 北郊
function recomputeDerived(base, manual){
  const m = Object.assign({}, base || {});
  const man = manual || {};
  for(const k in man){ if(man[k] !== undefined && man[k] !== null && man[k] !== '') m[k] = Number(man[k]) || 0; }
  // 合计的组成项：6个部门块 + 其他（残差）。注意：新增业绩、管家部 是另一维度的子集，不计入合计。
  const BLOCKS = ['品牌组-其他','活动组-活动支持','电商组','短视频组','品牌组-自然到店','商务部'];
  const g = k => (k in m ? (Number(m[k]) || 0) : 0);
  // 1) 合计：手动覆盖优先；否则当某部门块被手动覆盖时，仅按该块的增量调整原始合计；
  //    否则保持原合计（即 calculate 的真实总和 / 月视图的 Σ每日合计），不做任何重算。
  if(!('合计' in man)){
    let delta = 0;
    for(const k of BLOCKS){ if(k in man) delta += (Number(m[k]) || 0) - (Number(base[k]) || 0); }
    if(delta !== 0) m['合计'] = (Number(base['合计']) || 0) + delta;
  }
  // 2) 其他：手动覆盖优先；否则 = 合计 - 6部门块之和（残差，随合计/部门块变化同步）
  if(!('其他' in man)) m['其他'] = g('合计') - (g('品牌组-其他') + g('活动组-活动支持') + g('电商组') + g('短视频组') + g('品牌组-自然到店') + g('商务部'));
  // 3) 基础业绩 = 合计 - 新增业绩
  if(!('基础业绩' in man)) m['基础业绩'] = g('合计') - g('新增业绩');
  // 4) 高新 = 合计 - 北郊；门店合计 = 高新 + 北郊（恒等于合计）
  if(!('高新' in man)) m['高新'] = g('合计') - g('北郊');
  if(!('门店合计' in man)) m['门店合计'] = g('高新') + g('北郊');
  return m;
}

// 月累计展示值 = 各日(含当日手动覆盖)求和；若某行有手动月累计覆盖则以手动值优先。
// 作用：手动改「每日」单元格 → 该行月累计联动更新；手动改「月累计」单元格 → 合计/派生行联动更新。
// 无手动编辑时，Σ每日值 与 calculate() 的月累计口径(当月1日~截止日)数值等价，无回归。
// 合计 = 各叶子部门月累计之和（含手动覆盖），使手动改任一部门后"合计/基础业绩/其他"等派生行同步更新。
function computeMonthValues(dateList, dailyResults){
  const m = {};
  dateList.forEach((d, idx) => {
    const dr = dailyResults[idx];
    const manDay = (state.manual.daily && state.manual.daily[d]) || {};
    // 用「当日重算后(已含当日覆盖)」的值做累加：
    //  - 月合计 = Σ当日合计（当日覆盖如实并入月累计）
    //  - 新增业绩/管家部 等另一维度子集不会误并入合计
    const dayVals = dr ? recomputeDerived(dr.values, manDay)
                       : (manDay && Object.keys(manDay).length ? Object.assign({}, manDay) : {});
    for(const k in dayVals){ m[k] = (m[k] || 0) + (Number(dayVals[k]) || 0); }
  });
  // 应用手动月累计覆盖（覆盖优先于按日求和）
  const mm = state.manual.month || {};
  for(const k in mm){ if(mm[k] !== undefined && mm[k] !== '') m[k] = Number(mm[k]) || 0; }
  // 合计 = 6个部门块 + 其他（新增业绩/管家部不计入）；手动"合计"覆盖优先
  const BLOCKS = ['品牌组-其他','活动组-活动支持','电商组','短视频组','品牌组-自然到店','商务部'];
  let sumBlocks = 0;
  for(const k of BLOCKS) sumBlocks += (Number(m[k]) || 0);
  const otherVal = (Number(m['其他']) || 0);
  if(mm['合计'] !== undefined && mm['合计'] !== '') m['合计'] = Number(mm['合计']) || 0;
  else m['合计'] = sumBlocks + otherVal;
  return m;
}

// 默认隐藏除最新日期外的所有日期列（仅保留截止日那列可见）
function defaultHideOldDates(tableEl, dateList){
  const lastIdx = dateList.length - 1;
  if(lastIdx < 1) return; // 仅1列时不隐藏
  for(let idx=0; idx<lastIdx; idx++){
    tableEl.querySelectorAll(`[data-idx="${idx}"]`).forEach(el=>el.classList.add('col-hidden'));
  }
  if(tableEl.id === 'store-table') updateStoreTitleColspan(tableEl);
  else updateMainTitleColspan(tableEl);
}

// 渲染前捕获用户当前"已隐藏"的每日列日期集合，使手动调整/重算重渲染时保留展开状态
// 返回 null 表示表尚未构建（首渲染）→ 走默认仅显最新；返回 Set（可能为空=用户已全部展开）
function captureHiddenDates(tableEl){
  let els;
  try { els = tableEl.querySelectorAll('th.daily-col'); } catch(e){ return null; }
  if(!els || !els.length) return null;
  const set = new Set();
  els.forEach(th => { if(th.classList.contains('col-hidden') && th.dataset && th.dataset.date) set.add(th.dataset.date); });
  return set;
}
// 渲染后恢复用户可见性：仅隐藏"之前就隐藏"的日期列；首次渲染(null)则默认仅显最新
function applyHiddenState(tableEl, dateList, prevHidden){
  if(prevHidden === null){ defaultHideOldDates(tableEl, dateList); return; }
  if(prevHidden.size){
    // 单次扫描替代「每个隐藏日期各扫一遍全表」，避免 O(隐藏列数 × 单元格数) 的重复查询
    const cells = tableEl.querySelectorAll('td[data-idx],th[data-idx]');
    cells.forEach(el => { const i = +el.dataset.idx; if(prevHidden.has(dateList[i])) el.classList.add('col-hidden'); });
  }
  if(tableEl.id === 'store-table') updateStoreTitleColspan(tableEl);
  else updateMainTitleColspan(tableEl);
}

function renderReportInner(){
  readTargets();
  // 无数据时显示简洁占位
  if(!state.results){
    $('report-table').querySelector('thead').innerHTML = '<tr><th colspan="6" class="sub-head" style="background:#fff;color:var(--c-text);">请先在"① 数据导入"上传订单表</th></tr>';
    $('report-body').innerHTML = '';
    // 门店业绩占位
    $('store-table').querySelector('thead').innerHTML = '<tr><th colspan="4" class="sub-head" style="background:#fff;color:var(--c-text);">请先在"① 数据导入"上传订单表</th></tr>';
    $('store-body').innerHTML = '';
    return;
  }
  // 导入数据的最新日期
  const latest = latestOrderDate();
  const latestDate = latest ? new Date(latest) : null;

  // 截止日期显示（标题用 cutoff）
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const monthStr = (cutoff.getMonth()+1) + '月' + cutoff.getDate() + '日';

  // 当月每日列：1日到 cutoff
  const dateList = generateDateRange(cutoff);
  const totalCols = 3 + 1 + 1 + dateList.length; // 分类3 + 月目标 + 月累计 + 每日
  // 预按日期分组一次，避免每天重复全量过滤（性能优化）
  const byDate = new Map();
  for(const r of state.orders){
    if(!r.payTime) continue;
    const k = fmtDate(r.payTime);
    if(!byDate.has(k)) byDate.set(k, []);
    byDate.get(k).push(r);
  }
  const dailyResults = dateList.map(d => {
    const g = byDate.get(d);
    return (g && g.length) ? calculate(g, state.vabSet, new Date(d)) : null;
  });
  // 缓存：供"按日期复制整列"复用，避免重复计算
  state.dateList = dateList;
  state.dailyResults = dailyResults;

  // 月累计展示值 = 各日求和(含当日手动覆盖)；手动月累计覆盖优先。
  // 手动改每日→联动月累计；手动改月累计→联动合计/派生行（dateList/dailyResults 在此处已就绪）
  const monthCalc = computeMonthValues(dateList, dailyResults);
  const dispValues = recomputeDerived(monthCalc, {});

  // 表头：标题行 + 一行表头，与参考图片一致
  //   行1：标题（深藏蓝）
  //   行2：部门类别(colspan=3) | 月目标 | 月累计 | 每日日期列 | 批量 | 展开
  // 标题行只跨主数据列（部门类别3+月目标+月累计+每日），不合并「批量」「展开」两列
  const titleCols = 3 + 1 + 1 + dateList.length;
  const thead = `
    <tr><th colspan="${titleCols}" class="title-cell">各部门业绩完成情况跟踪日报表(截止${monthStr})</th></tr>
    <tr>
      <th colspan="3" class="sub-head">部门类别</th>
      <th class="sub-head">月目标<br>(单位：万元)</th>
      <th class="sub-head">月累计<br>(单位：元)</th>
      ${dateList.map((d, i) => {
        const dt = new Date(d);
        return `<th class="sub-head daily-col" data-date="${d}" data-idx="${i}" style="cursor:pointer;position:relative;" title="点击收起该列（批量模式下点击为多选）">
          <span class="daily-label">${dt.getMonth()+1}月${dt.getDate()}日</span>
          <span class="copy-col" data-date="${d}" data-idx="${i}" title="复制该日期整列数据">复制</span>
        </th>`;
      }).join('')}
    </tr>
  `;

  const tbody = $('report-body');
  // 计算 colspan 跳过已被 rowspan 占用的格子
  let skipMain = 0, skipSub = 0, skipTarget = 0;
  let html = '';
  REPORT_TEMPLATE.forEach((row)=>{
    if(row.hideRow) return;
    const rowKey = row.valueKey || row.targetKey || row.sub || row.item;
    // rowspanTarget > 1 的合并行，月目标用 row.sub 作为 key（如「公司支持」），否则用 row.targetKey
    const targetKey = (row.rowspanTarget > 1 && row.sub) ? row.sub : row.targetKey;
    const target = state.targets[targetKey] ?? 0;
    // 月累计取值键必须与单元格 data-rowkey（手动编辑写入的键）完全一致：统一用 rowKey = targetKey||sub||item
    // 旧逻辑先读 row.sub/row.item 会导致合并行（如「短视频组及自然到店」对应 短视频组/品牌组-自然到店 两子行）
    // 显示与写入键不一致 → 双击手动调整后月累计不刷新。
    const monthVal = dispValues[rowKey] ?? 0;
    const cls = row.rowClass || (row.highlight ? 'row-orange' : '');
    // 合计/重点行：inline font-weight + 更大字号 防止浏览器缓存旧 CSS 导致未加粗
    const rowStyle = row.highlight ? ' style="font-weight:700;font-size:15px;"' : '';

    // 每日列单元格：override-aware（含派生行联动），不使用黄色高亮
    const dailyCells = dailyResults.map((dr, idx) => {
      const dateStr = dateList[idx];
      const dispDay = recomputeDerived(dr ? dr.values : {}, state.manual.daily[dateStr] || {});
      const v = dispDay[rowKey] ?? 0;
      const cls = 'num daily-col-' + idx;
      return `<td class="${cls}" data-idx="${idx}" data-rowkey="${rowKey}" data-date="${dateStr}" title="${dateStr}（双击可编辑）">${fmt(v)}</td>`;
    }).join('');

    if(row.mergedLabel){
      // 独立汇总行（商务部/其他/合计/基础业绩/新增业绩/管家部/北郊/高新）：
      // 仅把「部门类别」3 列合并为标签（colspan=3），月目标、月累计仍各自显示
      const label = `${row.main}${row.sub?' '+row.sub:''}`;
      const labelCell = `<td class="cat" colspan="3" style="white-space:nowrap;">${label}</td>`;
      // 月目标列：双击可编辑
      const targetCell = `<td class="target" data-targetkey="${row.targetKey}" title="双击修改月目标">${target||'-'}</td>`;
      // 月累计列：override-aware（已含派生行联动，不含黄色高亮）
      const monthCell = `<td class="num" data-rowkey="${rowKey}" title="双击可编辑">${fmt(monthVal)}</td>`;
      html += `<tr class="${cls}"${rowStyle}>${labelCell}${targetCell}${monthCell}${dailyCells}</tr>`;
      return;
    }

    let mainCell='', subCell='', itemCell='', targetCell='';
    if(skipMain > 0){
      skipMain--;
    } else if(row.main){
      mainCell = `<td class="cat" rowspan="${row.rowspanMain||1}">${row.main}</td>`;
      if(row.rowspanMain>1) skipMain = row.rowspanMain-1;
    }

    if(skipSub > 0){
      skipSub--;
    } else if(row.sub){
      subCell = `<td class="cat sub" rowspan="${row.rowspanSub||1}">${row.sub}</td>`;
      if(row.rowspanSub>1) skipSub = row.rowspanSub-1;
    }

    if(row.item || row.main || row.sub){
      const ind = row.indent ? 'padding-left:18px;' : '';
      itemCell = `<td class="cat sub" style="${ind}">${row.item||''}</td>`;
    }

    // 月目标列：跨行共享（rowspanTarget>1 时首行输出 rowspan，后续行跳过）；双击可编辑
    if(skipTarget > 0){
      skipTarget--;
    } else if(row.rowspanTarget > 1){
      targetCell = `<td class="target" rowspan="${row.rowspanTarget}" data-targetkey="${targetKey}" title="双击修改月目标">${target||'-'}</td>`;
      skipTarget = row.rowspanTarget - 1;
    } else {
      targetCell = `<td class="target" data-targetkey="${targetKey}" title="双击修改月目标">${target||'-'}</td>`;
    }

    // 月累计列：override-aware（已含派生行联动，不含黄色高亮）
    html += `<tr class="${cls}">${mainCell}${subCell}${itemCell}${targetCell}<td class="num" data-rowkey="${rowKey}" title="双击可编辑">${fmt(monthVal)}</td>${dailyCells}</tr>`;
  });

  // 替换 thead（渲染前先捕获用户当前已展开的列，避免手动调整后旧列被重隐藏）
  const tableEl = $('report-table');
  const prevHidden = captureHiddenDates(tableEl);
  tableEl.querySelector('thead').innerHTML = thead;
  tbody.innerHTML = html;
  // 更新 colgroup 列宽（5 列基础 + 每日列）
  setupColgroup(tableEl, totalCols, dateList.length);
  setupResizeHandles(tableEl);
  bindDailyColumnToggle(tableEl);
  // 保留用户展开/收起状态：仅重新隐藏"之前就隐藏"的日期列；首次渲染走默认仅显最新
  applyHiddenState(tableEl, dateList, prevHidden);
  // 绑定双击编辑：月目标 / 月累计 / 每日数据
  bindManualEdit(tableEl);

  // 数据明细（导入数据明细 ⑥ 保留，业绩明细 ⑤ 已按需求取消）
  renderImportDetails();

  // ④ 门店业绩（高新 / 北郊）
  renderStoreReport(dateList, dailyResults);
}

// 双击编辑绑定：月目标、月累计、每日数据
function bindManualEdit(tableEl){
  // 月目标列（双击 → 保存到 state.targets + 同步"② 月目标"面板 input）
  tableEl.querySelectorAll('td.target[data-targetkey]').forEach(td=>{
    const key = td.dataset.targetkey;
    makeEditable(td, v=>{
      state.targets[key] = v;
      const inp = document.querySelector(`.target-input[data-key="${key}"]`);
      if(inp) inp.value = v;
      persistTargets();
      renderReport();
      showToast(`月目标「${key}」已改为 ${v} 万元`, 'ok');
    }, {step:'0.1'});
  });
  // 月累计列（双击 → 保存到 state.manual.month）
  tableEl.querySelectorAll('td.num[data-rowkey]:not([data-idx])').forEach(td=>{
    const key = td.dataset.rowkey;
    makeEditable(td, v=>{
      state.manual.month[key] = v;
      persistManual();
      renderReport();
      showToast(`「${key}」月累计已手动调整为 ${fmt(v)}`, 'ok');
    });
  });
  // 每日数据列（双击 → 保存到 state.manual.daily[date][rowKey]）
  tableEl.querySelectorAll('td.num[data-idx][data-rowkey]').forEach(td=>{
    const key = td.dataset.rowkey;
    const dateStr = td.dataset.date;
    makeEditable(td, v=>{
      if(!state.manual.daily[dateStr]) state.manual.daily[dateStr] = {};
      state.manual.daily[dateStr][key] = v;
      persistManual();
      renderReport();
      showToast(`${dateStr}「${key}」已手动调整为 ${fmt(v)}`, 'ok');
    });
  });
}

// ④ 门店业绩（高新店 / 经开店）：独立表格，显示每天各门店的业绩（不显示月目标）
// 高新店 = 收款员工为「南媛」；经开店 = 收款员工≠「南媛」
const STORE_ROWS = [
  {name:'高新店', targetKey:'高新', highlight:false},
  {name:'经开店', targetKey:'北郊', highlight:false},
  {name:'合计', targetKey:'门店合计', highlight:true},
];

function renderStoreReport(dateList, dailyResults){
  const storeTable = $('store-table');
  const storeBody = $('store-body');
  if(!state.results){
    storeTable.querySelector('thead').innerHTML = '<tr><th colspan="3" class="sub-head" style="background:#fff;color:var(--c-text);">请先在"① 数据导入"上传订单表</th></tr>';
    storeBody.innerHTML = '';
    return;
  }
  // 月累计展示值 = 各日求和(含当日手动覆盖)；手动月累计覆盖优先（与 ③ 主表一致）
  const monthCalc = computeMonthValues(dateList, dailyResults);
  const dispValues = recomputeDerived(monthCalc, {});
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const monthStr = (cutoff.getMonth()+1) + '月' + cutoff.getDate() + '日';
  const totalCols = 1 + 1 + dateList.length + 1; // 门店类别 + 月累计 + 每日 + 展开
  const storeTitleCols = 1 + 1 + dateList.length; // 门店类别 + 月累计 + 每日（标题不合并「展开」列）

  const thead = `
    <tr><th colspan="${storeTitleCols}" class="title-cell">门店业绩完成情况（高新店 / 经开店，截止${monthStr}）</th></tr>
    <tr>
      <th class="sub-head">门店类别</th>
      <th class="sub-head">月累计(单位：元)</th>
      ${dateList.map((d,i)=>{
        const dt = new Date(d);
        return `<th class="sub-head daily-col" data-date="${d}" data-idx="${i}" style="cursor:pointer;position:relative;" title="点击收起该列">${dt.getMonth()+1}月${dt.getDate()}日<span class="copy-col" data-date="${d}" data-idx="${i}" title="复制该日期整列数据">复制</span></th>`;
      }).join('')}
      <th class="sub-head" id="store-expand" style="width:46px;cursor:pointer;" title="展开所有日期列">+展开</th>
    </tr>
  `;

  let html = '';
  STORE_ROWS.forEach(row=>{
    const monthVal = dispValues[row.targetKey] ?? 0;
    const cls = row.highlight ? 'row-orange' : '';
    const rowStyle = row.highlight ? ' style="font-weight:700;font-size:15px;"' : '';
    // 每日列：override-aware（含派生行联动），不使用黄色高亮
    const dailyCells = dailyResults.map((dr, idx)=>{
      const dateStr = dateList[idx];
      const dispDay = recomputeDerived(dr ? dr.values : {}, state.manual.daily[dateStr] || {});
      const v = dispDay[row.targetKey] ?? 0;
      const cls = 'num daily-col-' + idx;
      const title = dateStr + '（双击可编辑）';
      return `<td class="${cls}" data-idx="${idx}" data-rowkey="${row.targetKey}" data-date="${dateStr}" title="${title}">${fmt(v)}</td>`;
    }).join('');
    const monthCell = `<td class="num" data-rowkey="${row.targetKey}" title="双击可编辑">${fmt(monthVal)}</td>`;
    html += `<tr class="${cls}"${rowStyle}><td class="cat">${row.name}</td>${monthCell}${dailyCells}</tr>`;
  });

  const prevHiddenStore = captureHiddenDates(storeTable);
  storeTable.querySelector('thead').innerHTML = thead;
  storeBody.innerHTML = html;

  // 列宽：前2列按比例，每日列固定 65px，展开列 36px
  storeTable.querySelectorAll('colgroup').forEach(c=>c.remove());
  const colgroup = document.createElement('colgroup');
  ['16%','20%'].forEach(w=>{ const c=document.createElement('col'); c.style.width=w; colgroup.appendChild(c); });
  for(let i=0;i<dateList.length;i++){ const c=document.createElement('col'); c.style.width=DAILY_COL_PX+'px'; colgroup.appendChild(c); }
  const ce=document.createElement('col'); ce.style.width='36px'; colgroup.appendChild(ce);
  storeTable.insertBefore(colgroup, storeTable.firstChild);

  bindStoreColumnToggle(storeTable);
  // 保留用户展开/收起状态：仅重新隐藏"之前就隐藏"的日期列；首次渲染走默认仅显最新
  applyHiddenState(storeTable, dateList, prevHiddenStore);
  bindManualEdit(storeTable);
}

// 门店表每日列：单列点击收起 + "+展开"恢复（不引入批量模式，避免与 ③ 全局 batchMode 冲突）
function updateStoreTitleColspan(tableEl){
  // 标题跨整个表头可见区域：门店类别 + 月累计 + 可见每日列 + 展开。
  const visibleDaily = tableEl.querySelectorAll('th.daily-col:not(.col-hidden)').length;
  const titleTh = tableEl.querySelector('thead tr:first-child th');
  if(titleTh) titleTh.colSpan = 3 + visibleDaily;
}

function bindStoreColumnToggle(tableEl){
  const expandBtn = tableEl.querySelector('#store-expand');
  tableEl.querySelectorAll('th.daily-col').forEach(th=>{
    th.addEventListener('click', ()=>{
      const idx = th.dataset.idx;
      const hide = !th.classList.contains('col-hidden');
      tableEl.querySelectorAll(`[data-idx="${idx}"]`).forEach(el=>el.classList.toggle('col-hidden', hide));
      updateStoreTitleColspan(tableEl);
    });
  });
  // 复制整列（停止冒泡，避免触发收起）
  tableEl.querySelectorAll('th.daily-col .copy-col').forEach(btn=>{
    btn.addEventListener('click', e=>{
      e.stopPropagation();
      copyDailyColumn(btn.dataset.date, +btn.dataset.idx, true);
    });
  });
  if(expandBtn){
    expandBtn.addEventListener('click', ()=>{
      tableEl.querySelectorAll('.col-hidden').forEach(el=>el.classList.remove('col-hidden'));
      updateStoreTitleColspan(tableEl);
    });
  }
  // 初始同步一次
  updateStoreTitleColspan(tableEl);
}

// ===== 按日期复制整列数据 =====
// 复制格式：首行为表头(部门类别<TAB>日期)，其后每行 "板块名<TAB>数值"，换行分隔 → 可直接粘贴进 Excel 成为两列日快照
function copyDailyColumn(dateStr, idx, isStore){
  const dailyResults = state.dailyResults || [];
  const dr = dailyResults[idx];
  const rows = isStore ? STORE_ROWS : REPORT_TEMPLATE;
  const lines = [];
  // 仅复制数值（纯列，无表头/类别名），每行一个，粘贴进 Excel 即成一列
  rows.forEach(row=>{
    if(row.hideRow) return;
    const rowKey = row.valueKey || row.targetKey || row.sub || row.item;
    // 优先取手动修改值，否则取当日计算结果
    const manualV = state.manual.daily[dateStr]?.[rowKey];
    const v = manualV !== undefined ? manualV
      : (dr ? (dr.values[row.sub] ?? dr.values[row.item] ?? dr.values[row.targetKey] ?? 0) : 0);
    lines.push(fmtPlain(v));
  });
  const text = lines.join('\n');
  copyTextToClipboard(text).then(ok=>{
    showToast(ok ? `已复制 ${dateStr} 整列数据（${lines.length} 行）` : '复制失败，请手动选择复制', ok ? 'ok' : 'err', 2600);
  });
}

// 写入剪贴板：优先 navigator.clipboard（http/localhost 安全上下文），
// 失败或 file:// 非安全上下文时回退到 textarea + execCommand
function copyTextToClipboard(text){
  return new Promise(resolve=>{
    try{
      if(navigator.clipboard && navigator.clipboard.writeText){
        navigator.clipboard.writeText(text).then(()=>resolve(true)).catch(()=>resolve(fallbackCopy(text)));
        return;
      }
    }catch(e){ /* 落到回退 */ }
    resolve(fallbackCopy(text));
  });
}
function fallbackCopy(text){
  try{
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.top = '-9999px';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.focus(); ta.select();
    const ok = document.execCommand('copy');
    document.body.removeChild(ta);
    return ok;
  }catch(e){ return false; }
}

// 表格列宽设置与拖拽：动态列数
// 前5列默认比例（部门类别3+月目标+月累计），每日列固定 90px
const FRONT_COL_PCT_5 = [7.5, 11.0, 12.5, 13.0, 18.5];
const DAILY_COL_PX = 90;
// （v-75）导出图片列宽：前导列锁定线上实测像素宽、日期列固定每日列宽（90px），
// 不再吸收 width:100% 的剩余宽度（v-74 原样继承会导致收起状态日期列约 37% 过宽）；
// 原 EXPORT_DAILY_COL_PX=95 已随 v-74 "导出专用样式"一并废除。

function setupColgroup(tableEl, totalCols, dailyCount){
  tableEl.querySelectorAll('colgroup').forEach(c=>c.remove());
  const colgroup = document.createElement('colgroup');
  for(let i=0;i<totalCols;i++){
    const col = document.createElement('col');
    if(state.colWidths && state.colWidths[i]){
      col.style.width = state.colWidths[i];
    }else if(i < FRONT_COL_PCT_5.length){
      col.style.width = FRONT_COL_PCT_5[i] + '%';
    }else{
      col.style.width = DAILY_COL_PX + 'px';
    }
    colgroup.appendChild(col);
  }
  tableEl.insertBefore(colgroup, tableEl.firstChild);
  return colgroup;
}

function setupResizeHandles(tableEl){
  tableEl.querySelectorAll('.resize-handle').forEach(h=>h.remove());
  const headerRow = tableEl.querySelector('thead tr:last-child');
  if(!headerRow) return;
  const ths = [...headerRow.querySelectorAll('th')];
  const colgroup = tableEl.querySelector('colgroup');
  let colCursor = 0;
  ths.forEach(th=>{
    // 跳过"+展开"和"☑ 批量"列（不参与拖拽）
    if(th.id === 'expand-all' || th.id === 'batch-toggle') return;
    const span = +th.colSpan || 1;
    const lastColIdx = colCursor + span - 1;
    colCursor += span;
    if(span !== 1) return;
    th.style.position = 'relative';
    const handle = document.createElement('div');
    handle.className = 'resize-handle';
    th.appendChild(handle);
    handle.addEventListener('mousedown', e=>{
      e.preventDefault();
      e.stopPropagation();
      const startX = e.clientX;
      const startW = colgroup.children[lastColIdx].offsetWidth || th.offsetWidth;
      const onMove = ev=>{
        const w = Math.max(40, startW + (ev.clientX - startX));
        colgroup.children[lastColIdx].style.width = w + 'px';
        state.colWidths[lastColIdx] = w + 'px';
      };
      const onUp = ()=>{
        document.removeEventListener('mousemove', onMove);
        document.removeEventListener('mouseup', onUp);
        document.body.style.cursor = '';
        persistColWidths(); // 拖拽结束后保存列宽
      };
      document.body.style.cursor = 'col-resize';
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
  });
}

// 列宽持久化（拖拽调整后的列宽刷新后仍保留）
function persistColWidths(){
  try{ localStorage.setItem(COLW_STORAGE_KEY, JSON.stringify(state.colWidths || {})); }catch(e){}
}
function restoreColWidths(){
  try{
    const raw = localStorage.getItem(COLW_STORAGE_KEY);
    if(raw){ const o = JSON.parse(raw); if(o && typeof o === 'object') state.colWidths = o; }
  }catch(e){}
}

// 每日列收起/展开 + 批量选择隐藏：点"☑ 批量"进入多选模式（黄色高亮），再点"✕ 完成隐藏"一次性隐藏选中列
let batchMode = false;
function updateMainTitleColspan(tableEl){
  // 标题跨整个表头可见区域：部门类别3 + 月目标 + 月累计 + 可见每日列。
  // 统一全宽后，无论 display:none 的隐藏列在不同浏览器/html2canvas 下如何折叠，标题始终覆盖全部可见列。
  const visibleDaily = tableEl.querySelectorAll('th.daily-col:not(.col-hidden)').length;
  const titleTh = tableEl.querySelector('thead tr:first-child th');
  if(titleTh) titleTh.colSpan = 5 + visibleDaily;
}

function bindDailyColumnToggle(tableEl){
  const dailyThs = [...tableEl.querySelectorAll('th.daily-col')];

  dailyThs.forEach(th => {
    th.addEventListener('click', ()=>{
      if(batchMode){
        // 批量模式：切换选中（黄色高亮）
        th.classList.toggle('selected');
      } else {
        // 单列收起
        const idx = th.dataset.idx;
        const hide = !th.classList.contains('col-hidden');
        tableEl.querySelectorAll(`[data-idx="${idx}"]`).forEach(el=>{
          el.classList.toggle('col-hidden', hide);
        });
        updateMainTitleColspan(tableEl);
      }
    });
  });
  // 复制整列（停止冒泡，避免触发收起/批量选中）
  tableEl.querySelectorAll('th.daily-col .copy-col').forEach(btn=>{
    btn.addEventListener('click', e=>{
      e.stopPropagation();
      copyDailyColumn(btn.dataset.date, +btn.dataset.idx, false);
    });
  });

  // 初始同步一次
  updateMainTitleColspan(tableEl);
}

// 「批量 / 展开」已移出报表表头、作为面板独立按钮（见 index.html ③ 业绩报表标题栏）。
// 仅在初始化绑定一次；点击时实时查询日期列，避免每次重渲染重复绑定、引用已被替换的旧节点。
function bindBatchExpandButtons(tableEl){
  const expandBtn = document.querySelector('#expand-all');
  const batchBtn = document.querySelector('#batch-toggle');
  const dailyThsNow = ()=> [...tableEl.querySelectorAll('th.daily-col')];
  const exitBatch = ()=>{
    batchMode = false;
    tableEl.classList.remove('batch-mode');
    if(batchBtn) batchBtn.textContent = '☑ 批量';
    dailyThsNow().forEach(th=>th.classList.remove('selected'));
  };
  if(batchBtn && !batchBtn.dataset.bound){
    batchBtn.dataset.bound = '1';
    batchBtn.addEventListener('click', ()=>{
      const dailyThs = dailyThsNow();
      if(batchMode){
        // 完成：隐藏所有选中列
        const selected = dailyThs.filter(th=>th.classList.contains('selected'));
        selected.forEach(th=>{
          const idx = th.dataset.idx;
          tableEl.querySelectorAll(`[data-idx="${idx}"]`).forEach(el=>{
            el.classList.add('col-hidden');
          });
        });
        if(selected.length > 0) showToast(`已隐藏 ${selected.length} 个日期列`, 'ok');
        exitBatch();
        updateMainTitleColspan(tableEl);
      } else {
        // 进入批量模式
        batchMode = true;
        tableEl.classList.add('batch-mode');
        batchBtn.textContent = '✕ 完成隐藏';
      }
    });
  }
  if(expandBtn && !expandBtn.dataset.bound){
    expandBtn.dataset.bound = '1';
    expandBtn.addEventListener('click', ()=>{
      tableEl.querySelectorAll('.col-hidden').forEach(el=>el.classList.remove('col-hidden'));
      if(batchMode) exitBatch();
      updateMainTitleColspan(tableEl);
    });
  }
}

function countGroupRows(grp){
  return REPORT_TEMPLATE.filter(r => r.grp === grp).length;
}

// 6. 详细数据明细（分规则展示）
// 6b. VAB 会员命中判断（与 buildVabSet 的三级容错对齐：原值小写 / 仅数字 / 仅数字去前导零）
function isVabMember(memberNo){
  const raw = String(memberNo||'').trim().toLowerCase();
  if(state.vabSet.has(raw)) return true;
  const dig = raw.replace(/\D/g,'');
  if(dig && state.vabSet.has(dig)) return true;
  const digNoZero = dig.replace(/^0+/,'');
  if(digNoZero && state.vabSet.has(digNoZero)) return true;
  return false;
}

// 6c. 导入数据明细：标签筛选（当前选中标签持久于 detailTab，上传/重算后保持）
// 渠道类页签容错匹配：真实导出表里"部门"可能落在 渠道分类/一级渠道/二级渠道 任一列，
// 故任一字段等于目标值即命中（营销部/商务部/市场部 实际在 channelClass；公司所有 在 channelL1）。
function chanHas(r, val){
  return (r.channelClass||'') === val || (r.channelL1||'') === val || (r.channelL2||'') === val;
}

// 推荐渠道归类（新增逻辑，不改原有 r1~r6）：渠道分类=推荐 时按「一级渠道」落地到营销部/商务部
// 一级分类包含「营销部」→ 计入营销部业绩（二级含自然到店 → 品牌组-自然到店，否则 → 电商组）
// 一级分类包含「商务部」→ 计入商务部业绩
function isRecToYingxiao(r){
  return (r.channelClass||'') === '推荐' && (r.channelL1||'').includes('营销部');
}
function isRecToShangwu(r){
  return (r.channelClass||'') === '推荐' && (r.channelL1||'').includes('商务部');
}

const DETAIL_TABS = [
  {key:'all',           label:'全部数据'},
  {key:'today',         label:'当日数据'},
  {key:'today_gaoxin',  label:'当日-高新'},
  {key:'today_beijiao', label:'当日-北郊'},
  {key:'shangwu',       label:'商务部'},
  {key:'newperf',       label:'新增业绩'},
  {key:'bigspend',      label:'消费≥2万'},
  {key:'yingxiao',      label:'营销部'},
  {key:'recommend',     label:'推荐'},
  {key:'staff',         label:'员工消费'},
  {key:'steward',       label:'管家部'},
];
let detailTab = 'all';
let detailMonth = ''; // 月度筛选：''=全部，'YYYY-MM'
let detailDate = '';  // 日期筛选：''=全部，'YYYY-MM-DD'
let detailSort = { key:'', dir:1 }; // 列排序：''=未排序；dir 1=升序 -1=降序（再次点击同列切换）

// 明细表排序取值：日期按时间戳、金额按数值、VAB 按 0/1，其余按中文字符串比较
function detailSortVal(r, key){
  switch(key){
    case 'payTime': return r.payTime ? r.payTime.getTime() : 0;
    case 'cashPay': return Number(r.cashPay)||0;
    case 'vab':     return isVabMember(r.memberNo) ? 1 : 0;
    default:        return String(r[key]||'');
  }
}
// 可排序列定义（# 与 操作 列不排序）
const DETAIL_SORT_COLS = [
  ['name','姓名'], ['memberNo','会员号'], ['channelClass','渠道分类'], ['channelL1','一级渠道'],
  ['channelL2','二级渠道'], ['channelL3','三级渠道'], ['consultant','咨询师'], ['cashier','收款员工'],
  ['payTime','收款时间'], ['cashPay','现款支付'], ['vab','VAB'],
];
// 对筛选后的列表就地排序（无排序状态时保持导入顺序）
function applyDetailSort(list){
  if(!detailSort.key) return list;
  const k = detailSort.key, dir = detailSort.dir;
  return list.slice().sort((a,b)=>{
    const va = detailSortVal(a,k), vb = detailSortVal(b,k);
    if(typeof va === 'string' && typeof vb === 'string') return va.localeCompare(vb,'zh-CN')*dir;
    return (va - vb)*dir;
  });
}

// 按标签筛选导入订单；"当日"=数据中的最新日期（latestOrderDate）；月度/日期筛选叠加取交集
function filterOrdersByTab(key){
  const orders = state.orders || [];
  const latest = latestOrderDate();
  const isToday = r => fmtDate(r.payTime) === latest;
  let list;
  if(key === 'all') list = orders.slice();
  else if(key === 'today') list = orders.filter(r => isToday(r));
  else if(key === 'today_gaoxin')  list = orders.filter(r => isToday(r) && (r.cashier||'').trim() === '南媛');
  else if(key === 'today_beijiao') list = orders.filter(r => isToday(r) && (r.cashier||'').trim() !== '南媛');
  else if(key === 'shangwu') list = orders.filter(r => isToday(r) && (chanHas(r,'商务部') || chanHas(r,'市场部') || isRecToShangwu(r)));
  else if(key === 'newperf')  list = orders.filter(r => isToday(r) && isVabMember(r.memberNo));
  else if(key === 'bigspend'){
    // 按日 + 单会员号 聚合现款，≥20000 的(日,会员)组整组显示（单会员号现款消费≥两万）
    const sumMap = new Map();
    orders.forEach(r=>{
      const m = String(r.memberNo||'').trim();
      if(!m) return;
      const k = fmtDate(r.payTime) + '|' + m;
      sumMap.set(k, (sumMap.get(k)||0) + (Number(r.cashPay)||0));
    });
    const hit = new Set([...sumMap.entries()].filter(([k,v])=>v>=20000).map(([k])=>k));
    list = orders.filter(r=>{ const m=String(r.memberNo||'').trim(); return m && hit.has(fmtDate(r.payTime)+'|'+m); });
  }
  else if(key === 'yingxiao') list = orders.filter(r => isToday(r) && (chanHas(r,'营销部') || isRecToYingxiao(r)));
  // 推荐页签：渠道分类=推荐 的订单（当日口径，与商务部/营销部页签一致），
  // 用于核对新归类逻辑落位（一级含营销部→营销部，含商务部→商务部，未匹配→其他）
  else if(key === 'recommend') list = orders.filter(r => isToday(r) && (r.channelClass||'') === '推荐');
  // 员工消费：渠道分类=公司所有 且 客户渠道标记为"员工本人及家属"。
  // 实测真实导出表里该标记落在【二级渠道】(L2)，旧导出可能在【三级渠道】(L3)，故 L2/L3 任一命中即可；
  // 原代码只查 L3 导致整组 0 行。按全月统计（不限当日），与「管家部」全月审计口径一致。
  else if(key === 'staff')    list = orders.filter(r => chanHas(r,'公司所有') && ((r.channelL2||'').includes('员工本人及家属') || (r.channelL3||'').includes('员工本人及家属')));
  else if(key === 'steward')  list = orders.filter(r => isSteward(r)); // 管家部：开单咨询师口径 + 渠道兜底（全月审计，不限于当日）
  else list = orders.slice();
  // 月度 / 日期 筛选叠加（与标签取交集）
  if(detailMonth) list = list.filter(r => r.payTime && fmtMonth(r.payTime) === detailMonth);
  if(detailDate)  list = list.filter(r => r.payTime && fmtDate(r.payTime) === detailDate);
  return list;
}

// 6c. 导入数据明细（扁平展示全部导入订单 + 标签筛选，便于核对原始数据 / VAB 命中）
function renderImportDetails(){
  const el = $('import-details');
  if(!el) return;
  const orders = state.orders || [];
  if(orders.length === 0){ el.innerHTML = '<div class="file-meta" style="padding:12px 0;">尚未导入订单数据</div>'; return; }
  const listRaw = filterOrdersByTab(detailTab);
  const list = applyDetailSort(listRaw);
  const vabCount = list.filter(r=>isVabMember(r.memberNo)).length;
  const total = list.reduce((a,b)=>a+(Number(b.cashPay)||0),0);
  const body = list.map((r,i)=>{
    const vab = isVabMember(r.memberNo);
    return `<tr data-uid="${esc(r._uid)}">
      <td style="text-align:center;color:var(--c-muted);">${i+1}</td>
      <td>${esc(r.name)}</td>
      <td>${esc(r.memberNo)}</td>
      <td>${esc(r.channelClass)}</td>
      <td>${esc(r.channelL1)}</td>
      <td>${esc(r.channelL2)}</td>
      <td>${esc(r.channelL3)}</td>
      <td>${esc(r.consultant)}</td>
      <td>${esc(r.cashier)}</td>
      <td>${r.payTime?r.payTime.toLocaleString('zh-CN'):''}</td>
      <td style="text-align:right;">${fmt(r.cashPay)}</td>
      <td style="text-align:center;${vab?'color:#1a73e8;font-weight:600;':''}">${vab?'✔':'—'}</td>
      <td style="text-align:center;"><button class="detail-del" type="button" data-uid="${esc(r._uid)}" title="删除该条数据">×</button></td>
    </tr>`;
  }).join('');
  // 月度 / 日期 筛选下拉（选项来自当前导入订单中实际出现的月份/日期，按时间升序）
  const months = [...new Set(orders.map(r=>r.payTime?fmtMonth(r.payTime):null).filter(Boolean))].sort();
  const dates  = [...new Set(orders.map(r=>r.payTime?fmtDate(r.payTime):null).filter(Boolean))].sort();
  const monthOpts = ['<option value="">全部月份</option>']
    .concat(months.map(m=>`<option value="${m}" ${m===detailMonth?'selected':''}>${m}</option>`)).join('');
  const dateOpts = ['<option value="">全部日期</option>']
    .concat(dates.map(d=>`<option value="${d}" ${d===detailDate?'selected':''}>${d}</option>`)).join('');
  const filterBar = `
    <div class="detail-filters">
      <label class="detail-filter">月度
        <select id="detail-month-filter" class="detail-select">${monthOpts}</select>
      </label>
      <label class="detail-filter">日期
        <select id="detail-date-filter" class="detail-select">${dateOpts}</select>
      </label>
      ${(detailMonth||detailDate)?'<button type="button" id="detail-filter-reset" class="detail-reset" title="清除月度/日期筛选">清除筛选</button>':''}
    </div>`;
  const tabs = DETAIL_TABS.map(t=>`<button class="detail-tab ${t.key===detailTab?'active':''}" data-tab="${t.key}">${t.label}</button>`).join('');
  // 表头：# 与 操作 不排序，其余列点击排序（当前排序列显示 ▲升序 / ▼降序）
  const sortThs = ['<th>#</th>']
    .concat(DETAIL_SORT_COLS.map(([k,label])=>{
      const active = detailSort.key===k;
      const arrow = active ? (detailSort.dir===1?' ▲':' ▼') : '';
      return `<th data-sort="${k}" title="点击按${label}${active?(detailSort.dir===1?'升序':'降序'):'排序'}" `
        + `style="cursor:pointer;user-select:none;white-space:nowrap;${active?'color:#1a4f8b;text-decoration:underline;':''}">${label}${arrow}</th>`;
    }))
    .concat(['<th>操作</th>']).join('');
  const tabLabel = DETAIL_TABS.find(t=>t.key===detailTab).label;
  const filterNote = (detailMonth?('月度 '+detailMonth+' · '):'')+(detailDate?('日期 '+detailDate+' · '):'');
  el.innerHTML = `
    ${filterBar}
    <div class="detail-tabs">${tabs}</div>
    <div style="padding:6px 0 10px;color:var(--c-muted);font-size:12px;">
      当前标签「${tabLabel}」· ${filterNote}共 ${list.length} 行 · 现款支付合计 ¥${fmt(total)} · 命中VAB ${vabCount} 行
    </div>
    <div class="table-wrap" style="max-height:420px;overflow:auto;border:1px solid var(--c-border);border-radius:8px;">
      <table class="data">
        <thead><tr>${sortThs}</tr></thead>
        <tbody>${body}</tbody>
        <tfoot><tr style="font-weight:700;background:var(--c-blue-tint);color:var(--c-head);">
          <td colspan="10" style="text-align:right;">合计（${list.length} 行 · 命中VAB ${vabCount} 行）</td>
          <td style="text-align:right;">${fmt(total)}</td>
          <td style="text-align:center;">${vabCount}</td>
          <td></td>
        </tr></tfoot>
      </table>
    </div>`;
  // 事件委托（绑定一次）：标签切换 + 删除某条数据（删除后全表同步更新）+ 月度/日期筛选
  if(!el.dataset.bound){
    el.addEventListener('click', onImportDetailsClick);
    el.addEventListener('change', onImportDetailsChange);
    el.dataset.bound = '1';
  }
}

// 导入明细点击：切换标签 / 删除某条订单 / 清除筛选（data-uid 精确删除，删除后各统计表联动重算）
function onImportDetailsClick(e){
  const del = e.target.closest && e.target.closest('.detail-del');
  if(del){ deleteOrder(del.dataset.uid); return; }
  const tab = e.target.closest && e.target.closest('.detail-tab');
  if(tab){ detailTab = tab.dataset.tab; renderImportDetails(); return; }
  // 表头排序：同列再次点击切换升/降序，不同列默认升序
  const th = e.target.closest && e.target.closest('th[data-sort]');
  if(th){
    const k = th.dataset.sort;
    if(detailSort.key === k) detailSort.dir = -detailSort.dir;
    else detailSort = { key:k, dir:1 };
    renderImportDetails();
    return;
  }
  const reset = e.target.closest && e.target.closest('#detail-filter-reset');
  if(reset){ detailMonth = ''; detailDate = ''; renderImportDetails(); }
}

// 导入明细筛选：月度 / 日期下拉变更（与当前标签取交集，重新渲染明细）
function onImportDetailsChange(e){
  if(!e || !e.target) return;
  const m = e.target.closest && e.target.closest('#detail-month-filter');
  if(m){ detailMonth = m.value || ''; renderImportDetails(); return; }
  const d = e.target.closest && e.target.closest('#detail-date-filter');
  if(d){ detailDate = d.value || ''; renderImportDetails(); }
}

// 删除一条订单并同步刷新所有统计表
function deleteOrder(uid){
  if(!uid) return;
  const before = state.orders.length;
  state.orders = state.orders.filter(r => String(r._uid) !== String(uid));
  if(state.orders.length === before){ showToast('未找到该条数据', 'warn'); return; }
  // 删除数据即视为数据快照变化：清空所有手动调整（月累计/每日），避免残留的单元格覆盖值
  // 导致"月累计刷新、当日不刷新"等不一致。数据确定后用户可再双击重新修正。
  if(Object.keys(state.manual.month).length || Object.keys(state.manual.daily).length){
    state.manual = { month:{}, daily:{} };
    persistManual();
  }
  if(state.orders.length === 0){
    // 删空：清空计算结果，手动渲染占位（recompute 在空数据时提前返回，不渲染）
    state.results = null;
    renderReport();
    renderImportDetails();
    persistOrders(); // 删空后同步持久化（空数组），刷新不再出现旧数据
    showToast('已删除全部数据，手动调整已清空，各统计表已同步更新', 'ok');
    return;
  }
  recompute(); // 重算并渲染 ③主表 / ④门店表 / ⑥导入明细（当月/当日均按剩余数据重新计算）
  persistOrders(); // 删除后同步持久化，刷新仍保留剩余数据
  showToast('已删除 1 条数据并清空手动调整，各统计表（含当日列）已重新计算', 'ok');
}

// 日期合并：仅替换 incoming 中最大日期对应的历史订单；保留更早日期与手动调整。
function mergeOrdersByDate(existing, incoming){
  // incoming 覆盖到的日期集合（incoming 里出现过的所有日期）
  const incomingDates = new Set();
  incoming.forEach(r => { if(r.payTime) incomingDates.add(fmtDate(r.payTime)); });
  // 保留 existing 中「不在 incoming 日期集合内」的订单：这些日期未被新文件覆盖，应予保留（新增/不一致部分）
  // incoming 覆盖到的每个日期：整批用 incoming 的对应订单替换（覆盖），避免重复计数
  const kept = existing.filter(r => !r.payTime || !incomingDates.has(fmtDate(r.payTime)));
  const replaced = existing.length - kept.length;
  const newMaxDate = incomingDates.size ? [...incomingDates].sort().pop() : null;
  return { orders: kept.concat(incoming), newMaxDate, replaced, coveredDates: incomingDates.size };
}

// 7. 文件上传处理
async function onFileOrders(input){
  const f = input.files?.[0];
  if(!f) return;
  state.demoOrders = false; // 上传真实订单后，不再视为示例
  const mergeMode = $('merge-mode') ? $('merge-mode').checked : true;
  try{
    await ensureXlsx();
    const rows = (await parseExcelFile(f)).map(normalizeOrder).filter(r=>r.payTime).map(ensureUid);
    if(mergeMode && state.orders.length > 0){
      const { orders, newMaxDate, replaced, coveredDates } = mergeOrdersByDate(state.orders, rows);
      state.orders = orders;
      $('orders-meta').textContent = `✔ 合并导入成功：本次${rows.length}行，累计${state.orders.length}行`;
      showToast(`已合并 ${rows.length} 条（覆盖 ${coveredDates} 个日期共 ${replaced} 条历史订单，保留其它日期与手动调整）`, 'ok');
    }else{
      state.orders = rows;
      $('orders-meta').textContent = `✔ 已导入 ${rows.length} 行（订单）`;
      showToast(`订单表导入成功，共 ${state.orders.length} 条记录`, 'ok');
    }
    $('orders-meta').className = 'file-meta';
    // 自动将截止日期更新为数据中的最大日期（当日列显示最新一天）
    const maxDate = state.orders.reduce((a,b)=>{ if(!b.payTime) return a; const s=fmtDate(b.payTime); return (!a || s>a) ? s : a; }, null);
    if(maxDate && $('cutoff-date').value !== maxDate){ $('cutoff-date').value = maxDate; }
    recompute();
    persistOrders(); // 持久化导入的订单（含截止日期），刷新网页后自动恢复
    if(state.demoVab){
      showToast('⚠ 当前"新增业绩"仍使用示例VAB会员，请上传真实VAB会员表', 'warn', 4200);
    }
  }catch(e){
    console.error(e);
    showToast('订单表解析失败：'+e.message, 'err');
  }
}


async function onFileVab(input){
  const f = input.files?.[0];
  if(!f) return;
  try{
    await ensureXlsx();
    const rows = await parseExcelFile(f);
    setVabRows(rows); // 保留导入表格全部字段，并同步会员号/匹配集合
    state.demoVab = false; // 真实文件上传后，不再标记为示例
    persistVab(); // 保存到本地
    refreshVabStatus();
    recompute();
    showToast(`VAB 会员列表导入成功，已保存到本地`, 'ok');
  }catch(e){
    showToast('VAB 表解析失败：'+e.message, 'err');
  }
}

// 「编辑会员」= 展开/收起 VAB 会员清单（清单内单元格双击编辑、行尾 × 删除，见 renderVabList/startVabCellEdit）
function toggleVabList(){
  const details = document.querySelector('.vab-list-details');
  if(!details) return;
  details.open = !details.open;
  if(details.open){
    renderVabList(); // 展开时确保内容为最新
    if(details.scrollIntoView) details.scrollIntoView({behavior:'smooth', block:'nearest'});
  }
}

// 导入数据中的最新日期（YYYY-MM-DD）
function latestOrderDate(){
  if(!state.orders.length) return todayIso();
  return state.orders.reduce((a,b)=>{
    if(!b.payTime) return a;
    const s = fmtDate(b.payTime);
    return (!a || s > a) ? s : a;
  }, null) || todayIso();
}

// 生成当月 1 日到 cutoff 的所有日期（YYYY-MM-DD 数组）
function generateDateRange(cutoff){
  if(!cutoff) return [];
  const c = new Date(cutoff);
  if(isNaN(c.getTime())) return []; // 截止日期无效保护
  const year = c.getFullYear();
  const month = c.getMonth();
  const lastDay = c.getDate();
  if(!Number.isFinite(lastDay) || lastDay < 1) return [];
  const days = [];
  for(let d=1; d<=lastDay; d++){
    const dt = new Date(year, month, d);
    days.push(fmtDate(dt));
  }
  return days;
}

// 仅计算某一天各板块业绩（复用 calculate 但只过滤当天的数据）
function calculateDay(orders, vabSet, dateStr){
  const dayOrders = orders.filter(r=>r.payTime && fmtDate(r.payTime) === dateStr);
  if(dayOrders.length === 0) return null; // 无数据
  return calculate(dayOrders, vabSet, new Date(dateStr));
}

// 8. 重新计算
function fmtDate(d){
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}
// 仅取 年-月（用于月度筛选）
function fmtMonth(d){
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
}

function updateOrdersMeta(){
  if(state.orders.length === 0){ return; }
  const dates = state.orders.map(r=>r.payTime).filter(Boolean).map(fmtDate);
  let range = '';
  if(dates.length){
    const min = dates.reduce((a,b)=>a<b?a:b);
    const max = dates.reduce((a,b)=>a>b?a:b);
    range = (min===max) ? min : `${min} ~ ${max}`;
  }
  $('orders-meta').textContent = `✔ ${state.orders.length} 条 | 覆盖 ${range}`;
  $('orders-meta').className = 'file-meta';
}

function recompute(){
  if(state.orders.length === 0){
    showToast('请先导入订单表', 'err');
    return;
  }
  state.cutoffDate = $('cutoff-date').value || todayIso();
  updateOrdersMeta();
  state.results = calculate(state.orders, state.vabSet, new Date(state.cutoffDate));
  renderReport();
  refreshDebug();
}

// 刷新调试条（含管家部诊断），在初始化与每次重算后调用
function refreshDebug(){
  const dbg = $('debug-banner');
  if(dbg){
    const t = new Date().toLocaleTimeString('zh-CN');
    dbg.textContent = `⚙ v20260910-77 | ${t} | 订单 ${state.orders.length} 条 | VAB ${vabUniqueCount()} 人 | cutoff ${state.cutoffDate||'(未设)'} | results ${state.results?'已计算':'未计算'} | ${stewardDiag()}${manualOverrideDiag()}`;
    dbg.style.display = '';
  }
}

// 9. 示例数据已移除：本应用不再注入任何示例/演示数据，仅使用用户上传的真实订单与 VAB 会员表。

// 10. 导出
function downloadBlob(blob, filename){
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url), 1500);
}

// ===== 带样式 Excel 导出（与线上显示一致，仅主数据区域）=====
const cell = (v, t, s) => ({v, t, s});

// 主报表网格：部门类别(3) + 月目标 + 月累计 + 全部每日列（含隐藏）；含跨行合并
function buildReportGrid(){
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const dateList = state.dateList || generateDateRange(cutoff);
  const dailyResults = state.dailyResults && state.dailyResults.length ? state.dailyResults : dateList.map(d => calculateDay(state.orders, state.vabSet, d));
  const tableEl = $('report-table');
  const dailyIdx = [...tableEl.querySelectorAll('th.daily-col')].map(th => +th.dataset.idx);
  // 导出 Excel：最新日期列【之前】的每日列一律自动隐藏（与屏幕展开状态无关）。
  const lastDailyIdx = dailyIdx.length ? Math.max(...dailyIdx) : -1;
  const hiddenDailyIdx = dailyIdx.filter(i => i !== lastDailyIdx);
  const nCols = 3 + 1 + 1 + dailyIdx.length;
  const monthStr = (cutoff.getMonth()+1)+'月'+cutoff.getDate()+'日';
  const rows = [];
  const merges = [];
  // 标题行
  rows.push(Array.from({length:nCols}, () => null));
  rows[0][0] = cell('各部门业绩完成情况跟踪日报表(截止'+monthStr+')', 's', 'title');
  merges.push({r:0, c:0, r2:0, c2:nCols-1});
  // 表头行
  rows.push(Array.from({length:nCols}, () => null));
  rows[1][0] = cell('部门类别', 's', 'subhead'); merges.push({r:1, c:0, r2:1, c2:2});
  rows[1][3] = cell('月目标(万元)', 's', 'subhead');
  rows[1][4] = cell('月累计\n(单位：元)', 's', 'subhead');
  dailyIdx.forEach((di,k)=>{ const dt=new Date(dateList[di]); rows[1][5+k] = cell((dt.getMonth()+1)+'月'+dt.getDate()+'日', 's', 'subhead'); });
  // 表体
  let skipMain=0, skipSub=0, skipTarget=0, r=2;
  REPORT_TEMPLATE.forEach(row=>{
    if(row.hideRow) return;
    const rowKey = row.valueKey || row.targetKey || row.sub || row.item;
    const targetKey = (row.rowspanTarget > 1 && row.sub) ? row.sub : row.targetKey;
    const target = state.targets[targetKey] ?? 0;
    // 月累计展示值必须与屏幕一致：用 computeMonthValues（按日重算后求和，含每日覆盖 + 月累计覆盖优先），
    // 而非原始 state.results.values（不含每日覆盖），否则「双击改每日」后屏幕月累计更新、导出月累计却仍是旧值。
    const dispValues = recomputeDerived(computeMonthValues(dateList, dailyResults), {});
    // 月累计取值键与 HTML 表一致：统一用 rowKey（避免合并行 短视频组及自然到店 显示/写入键错位）
    const monthVal = dispValues[rowKey] ?? 0;
    const isPink = row.rowClass === 'row-pink';
    const isOrange = !!row.highlight;
    const arr = Array.from({length:nCols}, () => null);
    const catStyle = isPink ? 'pink' : (isOrange ? 'orange' : 'cat');
    const catSubStyle = isPink ? 'pink' : (isOrange ? 'orange' : 'catsub');
    const targetStyle = isPink ? 'pink' : (isOrange ? 'orange' : 'target');
    const dataStyle = isPink ? 'pink' : (isOrange ? 'orange' : 'num');

    if(row.mergedLabel){
      const label = `${row.main}${row.sub ? ' '+row.sub : ''}`.trim();
      arr[0] = cell(label, 's', isOrange ? 'orange' : (isPink ? 'pink' : 'cat'));
      merges.push({r, c:0, r2:r, c2:2});
    } else {
      if(skipMain > 0){ skipMain--; }
      else if(row.main){ arr[0] = cell(row.main,'s',catStyle); if(row.rowspanMain>1){ merges.push({r,c:0,r2:r+row.rowspanMain-1,c2:0}); skipMain = row.rowspanMain-1; } }
      if(skipSub > 0){ skipSub--; }
      else if(row.sub){ arr[1] = cell(row.sub,'s',catSubStyle); if(row.rowspanSub>1){ merges.push({r,c:1,r2:r+row.rowspanSub-1,c2:1}); skipSub = row.rowspanSub-1; } }
      if(row.item || row.main || row.sub){ arr[2] = cell(row.item||'','s',catSubStyle); }
    }
    // 月目标
    if(skipTarget > 0){ skipTarget--; }
    else if(row.rowspanTarget > 1){ arr[3] = cell(target,'n',targetStyle); merges.push({r,c:3,r2:r+row.rowspanTarget-1,c2:3}); skipTarget = row.rowspanTarget-1; }
    else { arr[3] = cell(target,'n',targetStyle); }
    // 月累计（override-aware，已含派生行联动）
    arr[4] = cell(monthVal, 'n', isOrange ? 'orange' : (isPink ? 'pink' : 'num'));
    // 每日列（override-aware）
    dailyIdx.forEach((di,k)=>{
      const dateStr = dateList[di];
      const dispDay = recomputeDerived(dailyResults[di] ? dailyResults[di].values : {}, state.manual.daily[dateStr] || {});
      const v = dispDay[rowKey] ?? 0;
      arr[5+k] = cell(v, 'n', dataStyle);
    });
    rows.push(arr); r++;
  });
  // v-73 列宽收窄：部门类别区 A/B 收窄（C=16 保住"活动组-活动支持"不截断），
  // 月目标 14→12、月累计 18→13、每日列 13→11（最长金额 506,221.68=10 字符仍放得下）
  const colWidths = [10.5, 9.5, 16, 12, 13];
  dailyIdx.forEach(() => colWidths.push(11));
  // 把隐藏列的「数据列索引」映射到 Excel 工作表真实列号（跳过前面固定的 5 列：部门类别3+月目标+月累计）
  const hiddenCols = hiddenDailyIdx.map(idx => 5 + idx);
  return { rows, merges, colWidths, hiddenCols };
}

// 门店报表网格：门店类别 + 月累计 + 全部每日列（含隐藏）
function buildStoreGrid(){
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const dateList = state.dateList || generateDateRange(cutoff);
  const dailyResults = state.dailyResults && state.dailyResults.length ? state.dailyResults : dateList.map(d => calculateDay(state.orders, state.vabSet, d));
  const tableEl = $('store-table');
  const dailyIdx = [...tableEl.querySelectorAll('th.daily-col')].map(th => +th.dataset.idx);
  // 导出 Excel：最新日期列【之前】的每日列一律自动隐藏（与屏幕展开状态无关）。
  const lastDailyIdx = dailyIdx.length ? Math.max(...dailyIdx) : -1;
  const hiddenDailyIdx = dailyIdx.filter(i => i !== lastDailyIdx);
  const nCols = 1 + 1 + dailyIdx.length;
  const monthStr = (cutoff.getMonth()+1)+'月'+cutoff.getDate()+'日';
  const rows = [];
  const merges = [];
  rows.push(Array.from({length:nCols}, () => null));
  rows[0][0] = cell('门店业绩完成情况（高新店 / 经开店，截止'+monthStr+'）', 's', 'title');
  merges.push({r:0, c:0, r2:0, c2:nCols-1});
  rows.push(Array.from({length:nCols}, () => null));
  rows[1][0] = cell('门店类别', 's', 'subhead');
  rows[1][1] = cell('月累计\n(单位：元)', 's', 'subhead');
  dailyIdx.forEach((di,k)=>{ const dt=new Date(dateList[di]); rows[1][2+k] = cell((dt.getMonth()+1)+'月'+dt.getDate()+'日', 's', 'subhead'); });
  let r=2;
  STORE_ROWS.forEach(row=>{
    const isOrange = !!row.highlight;
    const arr = Array.from({length:nCols}, () => null);
    arr[0] = cell(row.name, 's', isOrange ? 'orange' : 'cat');
    const dispValues = recomputeDerived(computeMonthValues(dateList, dailyResults), {});
    const monthVal = dispValues[row.targetKey] ?? 0;
    arr[1] = cell(monthVal, 'n', isOrange ? 'orange' : 'num');
    dailyIdx.forEach((di,k)=>{
      const dateStr = dateList[di];
      const dispDay = recomputeDerived(dailyResults[di] ? dailyResults[di].values : {}, state.manual.daily[dateStr] || {});
      const v = dispDay[row.targetKey] ?? 0;
      arr[2+k] = cell(v, 'n', isOrange ? 'orange' : 'num');
    });
    rows.push(arr); r++;
  });
  // v-73 列宽收窄：与主报表口径一致（门店类别 16→10.5、月累计 16→13、每日列 13→11）
  const colWidths = [10.5, 13];
  dailyIdx.forEach(() => colWidths.push(11));
  // 门店表固定列：门店类别(1) + 月累计(1)，之后是每日列
  const hiddenCols = hiddenDailyIdx.map(idx => 2 + idx);
  return { rows, merges, colWidths, hiddenCols };
}

// （v-75）量测线上表格各逻辑列的实际渲染像素宽：
// 对每个逻辑列找一个 colspan=1 且可见(offsetWidth>0)的单元格取其 offsetWidth（即该列真实列宽）。
// 返回 widthByCol[c]（像素；量不到为 0，由调用方退回 colgroup 规格）。
function measureLiveColWidths(tableEl){
  const widths = [];
  if(!tableEl || !tableEl.rows || !tableEl.rows.length) return widths;
  const grid = [];
  const cellsInfo = [];
  [...tableEl.rows].forEach((row, r) => {
    if(!grid[r]) grid[r] = [];
    let c = 0;
    [...row.cells].forEach(cell => {
      while(grid[r][c]) c++;
      const cs = cell.colSpan || 1, rs = cell.rowSpan || 1;
      cellsInfo.push({ cell, start: c, span: cs });
      for(let i = 0; i < rs; i++){
        if(!grid[r + i]) grid[r + i] = [];
        for(let j = 0; j < cs; j++) grid[r + i][c + j] = cell;
      }
      c += cs;
    });
  });
  let maxC = 0;
  grid.forEach(g => { if(g && g.length > maxC) maxC = g.length; });
  for(let c = 0; c < maxC; c++){
    const hit = cellsInfo.find(ci => ci.start === c && ci.span === 1 && ci.cell.offsetWidth > 0);
    widths[c] = hit ? hit.cell.offsetWidth : 0;
  }
  return widths;
}

// （v-75，纯函数便于回归）导出克隆列宽锁定计算：
// 前导列优先取线上实测像素宽，量不到退回线上 colgroup 规格（% 按 refW 换算 / px 直接用）；
// 日期列取线上规格（可能被拖拽改宽），无规格时用兜底宽 —— 不再吸收 width:100% 的剩余宽度。
// 返回 { frontPx:[...px], dailyPx, total }（单位 px）。
function computeExportColPx(frontCount, frontMeasured, liveSpec, refW, dailySpec, fallbackDailyPx){
  function specToPx(sw){
    if(!sw) return 0;
    if(sw.endsWith('%')) return refW ? Math.round(parseFloat(sw) * refW / 100) : 0;
    const v = parseFloat(sw);
    return isNaN(v) ? 0 : v;
  }
  const frontPx = [];
  let total = 0;
  for(let k = 0; k < frontCount; k++){
    const w = frontMeasured[k] || specToPx(liveSpec[k] || '');
    frontPx.push(w);
    total += w;
  }
  let dailyPx = specToPx(dailySpec || '');
  if(!dailyPx) dailyPx = fallbackDailyPx || 0;
  total += dailyPx;
  return { frontPx, dailyPx, total };
}

// 导出 PNG 前构建干净的表格克隆：物理删除隐藏日期列（display:none 在部分 html2canvas
// 版本下仍占位导致右侧留白），并离屏渲染，保证任何 html2canvas 版本/缓存状态下都紧贴内容、无空白。
function buildExportClone(tableEl){
  const clone = tableEl.cloneNode(true);
  clone.removeAttribute('id');

  // 0) 图片导出：日期列只保留【最新日期列】，其余日期列一律隐藏（不出现在导出图中）。
  //    与屏幕展开状态无关——即使屏幕上展开了旧日期列，导出图片也只含最新日期列。
  const dailyHeads = [...clone.querySelectorAll('th.daily-col')];
  const lastDailyIdx = dailyHeads.length ? Math.max(...dailyHeads.map(th => +th.dataset.idx)) : -1;
  if(lastDailyIdx >= 0){
    clone.querySelectorAll('th.daily-col, td[class*="daily-col"]').forEach(el => {
      const di = +el.dataset.idx;
      if(di === lastDailyIdx) el.classList.remove('col-hidden');
      else el.classList.add('col-hidden');
    });
  }

  // 1) 找出所有表头行里需要移除的列：
  //    - 日期列已在步骤0全部展开，此处仅移除纯 UI 控件列：批量/展开/门店展开（导出的图片不应出现这些按钮）
  //    表头采用 2 行结构：标题行 + 一行表头，仍用通用网格扫描以兼容 body 行的 rowspan。
  const headRowsArr = clone.tHead ? [...clone.tHead.rows] : [];
  const hiddenIdx = new Set();
  const occupied = headRowsArr.map(() => []);
  headRowsArr.forEach((row, r) => {
    let c = 0;
    [...row.cells].forEach(th => {
      while (occupied[r][c]) c++;
      const cs = th.colSpan || 1;
      const rs = th.rowSpan || 1;
      if (th.classList.contains('col-hidden') || th.id === 'batch-toggle' || th.id === 'expand-all' || th.id === 'store-expand') {
        for (let i = 0; i < cs; i++) hiddenIdx.add(c + i);
      }
      for (let i = 0; i < rs; i++) {
        if (!occupied[r + i]) occupied[r + i] = [];
        for (let j = 0; j < cs; j++) occupied[r + i][c + j] = true;
      }
      c += cs;
    });
  });

  // 1b) 记录原始每日列数（后续步骤会物理删除隐藏列，导致 querySelectorAll 数量变少）
  const originalDailyCount = clone.querySelectorAll('th.daily-col').length;

  // 2) 移除 colgroup 中对应的 <col>
  const colgroup = clone.querySelector('colgroup');
  if (colgroup) [...colgroup.children].forEach((col, i) => { if (hiddenIdx.has(i)) col.remove(); });

  // 3) 给整张表每个“ originating cell”标上逻辑起始列和跨度（考虑 rowspan/colspan），
  //    不能直接用 cellIndex——body 行因上方 rowspan 而 cellIndex 左移，会删错列。
  const allRows = [];
  if (clone.tHead) headRowsArr.forEach(r => allRows.push(r));
  if (clone.tBodies) [...clone.tBodies].forEach(b => [...b.rows].forEach(r => allRows.push(r)));

  const grid = []; // grid[r][c] = cell element occupying that logical slot
  allRows.forEach((row, r) => {
    if (!grid[r]) grid[r] = [];
    let c = 0;
    [...row.cells].forEach(cell => {
      while (grid[r][c]) c++; // 跳过被上方 rowspan 占用的格子
      const cs = cell.colSpan || 1;
      const rs = cell.rowSpan || 1;
      cell.dataset.logicalStart = String(c);
      cell.dataset.logicalSpan = String(cs);
      for (let i = 0; i < rs; i++) {
        if (!grid[r + i]) grid[r + i] = [];
        for (let j = 0; j < cs; j++) grid[r + i][c + j] = cell;
      }
      c += cs;
    });
  });

  // 4) 按行处理 originating cells：若某 cell 的跨度覆盖隐藏列，删除它或缩减其 colspan
  allRows.forEach((row, r) => {
    const cells = [...row.cells];
    for (let k = cells.length - 1; k >= 0; k--) {
      const cell = cells[k];
      const start = Number(cell.dataset.logicalStart) || 0;
      const span = Number(cell.dataset.logicalSpan) || 1;
      const covered = [];
      for (let i = 0; i < span; i++) if (hiddenIdx.has(start + i)) covered.push(start + i);
      if (covered.length === 0) continue; // 不涉及隐藏列
      if (covered.length === span) {
        cell.remove(); // 整个 cell 都在隐藏列上 → 直接删掉
      } else {
        // 部分隐藏：缩减 colspan（只保留可见列）
        const newSpan = span - covered.length;
        cell.setAttribute('colspan', String(newSpan > 0 ? newSpan : 1));
      }
    }
  });

  // 5) 修正标题行 colspan：按「克隆后实际保留的每日列数」计算
  //    隐藏列已在步骤 4 被物理删除（PNG 导出保持干净、只保留可见列），
  //    因此 colspan 必须用剩余可见每日列数，不能用 originalDailyCount（否则会多出幻影列、标题被拉伸过宽）。
  //    主表基础列 = 部门类别3 + 月目标 + 月累计 = 5；门店表基础列 = 门店类别 + 月累计 = 2。
  const isStore = clone.id === 'store-table';
  const baseCols = isStore ? 2 : 5;
  const visibleDailyAfter = clone.querySelectorAll('th.daily-col').length;
  if (headRowsArr.length) {
    headRowsArr[0].querySelectorAll('.title-cell').forEach(th => th.setAttribute('colspan', String(baseCols + visibleDailyAfter)));
  }

  // 6) 移除复制/拖拽等纯 UI 控件
  clone.querySelectorAll('.copy-col,.resize-handle').forEach(n => n.remove());

  // 7)（v-74）导出样式与页面展示完全一致：不再做任何导出专用样式覆盖。
  //    此前此处强制过 0.5px 细框线 / 9px 大内边距 / 表头扁平浅蓝黑字 / 标题平涂藏蓝不换行 /
  //    日期列 95px / fixed 布局，导致导出图与页面不一致（行更高、框线更细、表头字色不同、列宽不同）。
  //    克隆与页面同文档同 CSS：隐藏日期列已在步骤 0-5 物理删除、复制/拖拽控件已在步骤 6 移除，
  //    其余渲染（1px 黑框、6px 内边距、浅蓝表头+藏蓝字、渐变标题、90px 日期列、auto 布局）
  //    全部原样继承页面 —— 唯一差异即「非最新日期列被隐藏」。

  // 8)（v-75）导出宽度 = 前导列线上实际渲染像素宽之和 + 日期列宽（默认 90px，与页面每日列一致）。
  //    根因：线上表格 width:100%，收起状态只剩 1 个日期列时会吸收全部剩余宽度（约 37%，导出图日期列过宽）。
  //    修复：前导列逐列量测线上真实渲染宽并锁定为像素（量不到退回 colgroup 规格），
  //    日期列固定取每日列宽、不再吸收剩余宽度 —— 前导列与页面逐像素一致，仅日期列收窄为正常每日列宽。
  const refW = tableEl && (tableEl.offsetWidth || tableEl.scrollWidth) ? (tableEl.offsetWidth || tableEl.scrollWidth) : 0;
  const trimmedCols = [...clone.querySelectorAll('colgroup col')];
  const liveCols = tableEl ? [...tableEl.querySelectorAll('colgroup col')] : [];
  const liveSpec = liveCols.map(c => (c.style && c.style.width) ? c.style.width : '');
  const frontCount = trimmedCols.length - 1;        // 裁剪后 colgroup = 前导列 + 唯一日期列
  const dailyColIdx = frontCount + lastDailyIdx;    // 该日期列在线上 colgroup 中的下标
  if(trimmedCols.length){
    const measured = measureLiveColWidths(tableEl);
    const plan = computeExportColPx(frontCount, measured, liveSpec, refW, liveSpec[dailyColIdx] || '', DAILY_COL_PX);
    trimmedCols.forEach((col, k) => {
      if(k < frontCount){ if(plan.frontPx[k] > 0) col.style.width = plan.frontPx[k] + 'px'; }
      else col.style.width = plan.dailyPx + 'px';
    });
    clone.style.width = plan.total > 0 ? plan.total + 'px' : 'auto';
  }else{
    clone.style.width = refW > 0 ? refW + 'px' : 'auto';
  }
  clone.style.margin = '0';

  // 9)（v-76）仅导出侧框线调细：用户反馈导出 PNG 框线偏粗。
  //    页面 index.html 的 table.report 边框保持 1px 不变；此处只对本克隆设置内联更细边框，
  //    不改变页面，也不影响导出的其他样式（颜色/内边距/字体/列宽/布局均原样继承页面）。
  //    外框（table 元素）与各单元格（th/td）的 border-width 由继承的 1px 改为 0.5px，
  //    边框样式 solid 与颜色 #000000 仍来自页面 CSS，导出图框线约为页面的一半粗细。
  const EXPORT_BORDER_W = '0.5px';
  clone.style.borderWidth = EXPORT_BORDER_W;
  clone.querySelectorAll('th, td').forEach(function(cell){ cell.style.borderWidth = EXPORT_BORDER_W; });

  return clone;
}

async function exportPng(){
  await ensureHtml2canvas();
  // 图片导出：仅含屏幕可见列（隐藏的旧日期列在克隆阶段被物理删除），保持截图干净
  const table = $('report-table');
  const clone = buildExportClone(table);
  const holder = document.createElement('div');
  holder.style.cssText = 'position:absolute;left:-99999px;top:0;width:auto;background:#fff;z-index:-1;';
  holder.appendChild(clone);
  document.body.appendChild(holder);
  html2canvas(clone, {backgroundColor:'#fff', scale:2, useCORS:true}).then(canvas=>{
    canvas.toBlob(blob=>{
      downloadBlob(blob, `日业绩跟踪表-${state.cutoffDate||todayIso()}.png`);
      showToast('图片已导出 (v20260910-77)', 'ok');
    });
  }).catch(e=>{
    console.error(e);
    showToast('导出失败：'+e.message, 'err');
  }).finally(()=>{
    holder.remove();
  });
}

async function exportXlsx(){
  await ensureXlsxStyled();
  if(!state.results){ showToast('请先计算', 'err'); return; }
  const { rows, merges, colWidths, hiddenCols } = buildReportGrid();
  buildStyledXlsx({ sheetName:'日业绩', rows, merges, colWidths, hiddenCols }).then(blob=>{
    downloadBlob(blob, `日业绩-${state.cutoffDate||todayIso()}.xlsx`);
    showToast('Excel已导出(与线上显示一致, v20260910-77)', 'ok');
  }).catch(e=>{
    console.error(e);
    showToast('导出失败：'+e.message, 'err');
  });
}

// 门店业绩导出
async function exportStorePng(){
  await ensureHtml2canvas();
  // 门店图片导出：仅含屏幕可见列（隐藏的旧日期列在克隆阶段被物理删除）
  const table = $('store-table');
  const clone = buildExportClone(table);
  const holder = document.createElement('div');
  holder.style.cssText = 'position:absolute;left:-99999px;top:0;width:auto;background:#fff;z-index:-1;';
  holder.appendChild(clone);
  document.body.appendChild(holder);
  html2canvas(clone, {backgroundColor:'#fff', scale:2, useCORS:true}).then(canvas=>{
    canvas.toBlob(blob=>{
      downloadBlob(blob, `门店业绩-${state.cutoffDate||todayIso()}.png`);
      showToast('门店业绩图片已导出 (v20260910-77)', 'ok');
    });
  }).catch(e=>{ console.error(e); showToast('导出失败：'+e.message, 'err'); })
    .finally(()=>{ holder.remove(); });
}

async function exportStoreXlsx(){
  await ensureXlsxStyled();
  if(!state.results){ showToast('请先计算', 'err'); return; }
  const { rows, merges, colWidths, hiddenCols } = buildStoreGrid();
  buildStyledXlsx({ sheetName:'门店业绩', rows, merges, colWidths, hiddenCols }).then(blob=>{
    downloadBlob(blob, `门店业绩-${state.cutoffDate||todayIso()}.xlsx`);
    showToast('门店业绩Excel已导出(与线上显示一致)', 'ok');
  }).catch(e=>{
    console.error(e);
    showToast('导出失败：'+e.message, 'err');
  });
}

// 11. 事件绑定
// ② 月目标面板 收起/展开（默认收起）
let targetsCollapsed = true;
function toggleTargetsPanel(){
  targetsCollapsed = !targetsCollapsed;
  $('targets-body').style.display = targetsCollapsed ? 'none' : '';
  $('targets-arrow').textContent = targetsCollapsed ? '▼ 展开' : '▲ 收起';
}

function initApp(){
  $('cutoff-date').value = todayIso();
  renderTargets();
  // 月目标面板默认收起
  $('targets-body').style.display = 'none';
  $('targets-arrow').textContent = '▼ 展开';

  // 调试条：显示当前状态（确认页面是新版本、加载成功）
  const dbg = $('debug-banner');
  if(dbg){
    const t = new Date().toLocaleTimeString('zh-CN');
    dbg.textContent = `⚙ v20260910-77 | ${t} | 订单 ${state.orders.length} 条 | VAB ${vabUniqueCount()} 人 | cutoff ${state.cutoffDate||'(未设)'} | results ${state.results?'已计算':'未计算'}`;
    dbg.style.display = '';
  }
  if(vabUniqueCount() > 0){
    $('vab-meta').textContent = `✔ 本地VAB（${vabUniqueCount()} 人）`;
    $('vab-meta').className = 'file-meta';
  }

  $('file-orders').addEventListener('change', e=>onFileOrders(e.target));
  $('file-vab').addEventListener('change', e=>onFileVab(e.target));
  $('edit-vab').addEventListener('click', toggleVabList);
  $('recompute').addEventListener('click', recompute);
  $('cutoff-date').addEventListener('change', ()=>{ persistOrders(); recompute(); }); // 切换截止日后持久化并重算月累计/每日
  $('clear-manual').addEventListener('click', clearManualOverrides);
  $('clear-vab').addEventListener('click', clearVabList);
  $('clear-all').addEventListener('click', clearAllData);
  // VAB 新增会员弹窗
  $('vab-add-save').addEventListener('click', saveAddVabModal);
  $('vab-add-cancel').addEventListener('click', closeAddVabModal);
  const vabModal = $('vab-add-modal');
  if(vabModal) vabModal.addEventListener('click', e => { if(e.target === vabModal) closeAddVabModal(); });
  document.addEventListener('keydown', e => { if(e.key === 'Escape'){ const m = $('vab-add-modal'); if(m && m.style.display === 'flex') closeAddVabModal(); } });
  $('targets-toggle').addEventListener('click', toggleTargetsPanel);
  $('export-png').addEventListener('click', exportPng);
  $('export-xlsx').addEventListener('click', exportXlsx);
  $('export-store-png').addEventListener('click', exportStorePng);
  $('export-store-xlsx').addEventListener('click', exportStoreXlsx);
  // 「批量 / 展开」报表独立按钮（已移出表头）：仅绑定一次
  bindBatchExpandButtons($('report-table'));

  // 不再自动注入示例数据：打开即为空，仅显示用户上传的真实数据（VAB 已从本机恢复）
  refreshVabStatus(); // 还原后统一刷新 VAB 状态（已保存默认 vs 未上传）

  // 历史订单已在脚本加载时从本地恢复（restoreOrders）；若有效则恢复截止日期、自动重算月累计/每日数据
  if(state.orders.length > 0){
    if(state.cutoffDate) $('cutoff-date').value = state.cutoffDate;
    // 防止新导入订单的 _uid 与已恢复订单冲突：把序列号抬到已恢复最大值之上
    const maxUid = state.orders.reduce((m,r)=>{ const n = r._uid && /^ord_(\d+)$/.exec(r._uid); return n ? Math.max(m, +n[1]) : m; }, 0);
    if(maxUid > _orderUidSeq) _orderUidSeq = maxUid;
    // 关键性能修复：先即时渲染空表骨架（首帧），再在浏览器空闲时做重计算，
    // 避免点击菜单后因“解析+重算+整表重渲染”全阻塞主线程而长时间白屏/卡死
    renderReport();
    const runRecompute = () => { recompute(); showToast(`已从本地恢复 ${state.orders.length} 条历史订单，月累计/每日数据已自动刷新`, 'ok'); };
    if(window.requestIdleCallback) window.requestIdleCallback(runRecompute, { timeout: 400 });
    else setTimeout(runRecompute, 30);
  } else {
    renderReport();
  }

  // 调试条（末尾位置，记录最终状态）
  refreshDebug();
}

// 管家部诊断：展示「开单健康管家」字段的取值分布与计入/排除状态，便于定位 管家部=0 原因
function stewardDiag(){
  const orders = state.orders || [];
  const dist = {};
  const byDate = {};
  let hitLen = 0;
  // 单次遍历同时累计「开单健康管家分布」与「管家部命中日期分布」，避免对全量订单多次扫描
  for(const r of orders){
    const raw = (r.consultant||'').toString().trim();
    const k = raw ? raw : '（空）';
    dist[k] = (dist[k] || 0) + 1;
    if(isSteward(r)){
      hitLen++;
      const d = r.payTime ? fmtDate(r.payTime) : '（无日期）';
      byDate[d] = (byDate[d] || 0) + 1;
    }
  }
  const parts = Object.entries(dist).map(([k,v]) => `${k}${v}${isSteward({consultant:k==='（空）'?'':k})?'✓':'✗'}`);
  // 命中订单的「日期分布」：揭示管家部业绩落在哪一天（合并多日上传时可能不在截止日那列 → 当日列显示0但月累计有值）
  const datePart = Object.keys(byDate).length
    ? ' | 管家部日期分布 ' + Object.entries(byDate).map(([d,n]) => `${d.slice(5)}:${n}`).join('·')
    : ' | 管家部无命中';
  return `管家部命中${hitLen}单${datePart} | 开单健康管家分布 ${parts.join(' ')}`;
}

function manualOverrideDiag(){
  const m = state.manual || {month:{}, daily:{}};
  const keys = Object.keys(m.month || {});
  if(keys.length === 0) return '';
  const parts = keys.map(k => `${k}=${m.month[k]}`);
  return ` | 手动覆盖[${parts.join(', ')}]`;
}

// 兼容：脚本在 body 底部，DOM 可能已就绪；若仍处于 loading 则等 DOMContentLoaded
if(document.readyState === 'loading'){
  document.addEventListener('DOMContentLoaded', initApp);
}else{
  initApp();
}
