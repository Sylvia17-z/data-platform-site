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
  await loadScript('assets/xlsx-styled.js?v=20261024-01');
  return true;
}

// 1. 渠道分类枚举（用于选项下拉）
const CHANNEL_CLASSES = ['营销部','商务部','市场部','公司所有','经营科室'];
const L1_CHANNELS = ['品牌组','活动组','电商组','新媒体组','短视频组','商保业务','商务直客','中医科','儿科','客户转介绍','企业团检','市场田梦代理人','商务渠道-韩通代理人','公司所有'];

// 月目标默认值（万元）- 来自 sheet2 默认值
const DEFAULT_TARGETS = {
  '品牌组-其他':0,'活动组-活动支持':0,
  '电商组及短视频组':77,'品牌组-自然到店':3,
  '公司支持':360,
  '商务部':125,
  '其他':170,        // 原 155 + 市场部 15（市场部已并入"其他"，不再单列）
  '合计':360,
  '基础业绩':310,'新增业绩':50,'管家部':238
};
// 无业绩目标的部门（高新店 / 经开店）：不在「② 月目标」面板列出，也不在报表中显示目标
const NO_TARGET_KEYS = new Set(['高新','北郊']);
// 门店归属判定（高新店 / 经开店），全局单一口径：
//   高新店 = 收款员工为「南媛」或「系统管理员」
//   经开店 = 收款员工既非「南媛」也非「系统管理员」（即经开店数据不含收款员工是系统管理员的数据）
function isGaoXinCashier(r){
  const c = (r.cashier||'').toString().trim();
  return c === '南媛' || c === '系统管理员';
}
// ② 月目标面板分组（美化：按业务板块归类，便于核对）
const TARGET_GROUPS = [
  {title:'营销部板块', keys:['品牌组-其他','活动组-活动支持','电商组及短视频组','品牌组-自然到店','公司支持'], subtotal:true},
  {title:'商务部 / 其他', keys:['商务部','其他']},
  {title:'汇总校验', keys:['合计','基础业绩','新增业绩','管家部']},
];

// 管家部业绩口径（用户最终确认——纯按「开单咨询师」字段值判定，渠道不参与）：
//   开单咨询师 不为空 且 ≠ 无/许顺华/空/营销部/吴雅琴/张帆 → 计入管家部（王勤红不在此排除名单，须计入管家部）
// 注意：「营销部」判断的是【开单咨询师字段值】=营销部，而不是任何渠道字段（客户来源一级/二级渠道等均不参与管家部判定）。
function normName(v){
  return String(v == null ? '' : v).trim().replace(/[ 　]+/g, ' ').toLowerCase();
}
const STEWARD_EXCLUDE = new Set(['无','许顺华','空','营销部','吴雅琴','张帆'].map(normName));

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
  ordersTs: 0,     // 订单快照写入时刻（epoch ms），用于本地/云端新旧裁决（防云端旧快照覆盖本地新数据）
  snapTs: 0,       // 整包快照变更时刻（epoch ms）：orders/VAB/月目标/手动覆盖 任一变更都更新，随快照上云做跨设备裁决
  results: null,   // 最新计算结果
  histOrders: [],  // 近一年消费明细（2026-10-16 新增）：仅用于「营销部顾客一年内无消费 → 初诊客户」判定，不计入业绩金额
  histTs: 0,       // 历史明细写入时刻（epoch ms），参与初诊键集合缓存失效签名
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
// 云端整包快照版本号（单一时间戳）：orders / VAB 名单 / 月目标 / 手动覆盖 任一变更都会更新它。
// 用于「同一账号跨设备」的整包新旧裁决——修复「两个浏览器登录同一账号，订单同步了，
// 但 VAB 会员名单 / 月目标没同步」：以前这些本地状态不在云端快照里，故换机后回落为 0 / 默认值。
const SNAP_TS_STORAGE_KEY = 'dailyPerformance_snap_ts';
function loadSnapTs(){ try{ return Number(localStorage.getItem(SNAP_TS_STORAGE_KEY)) || 0; }catch(e){ return 0; } }
function saveSnapTs(v){
  try{ (window.__snapLog = window.__snapLog || []).push('set ' + (v||0) + ' @ ' + (new Error().stack || '').split('\n').slice(2,4).join(' <- ')); }catch(_){}
  try{ localStorage.setItem(SNAP_TS_STORAGE_KEY, String(v || 0)); }catch(e){}
}
function bumpSnapTs(){ state.snapTs = Date.now(); saveSnapTs(state.snapTs); return state.snapTs; }
state.snapTs = loadSnapTs(); // 本机最后一次本地变更时刻（0 = 本机从未产生本地变更）
function persistTargets(){ writeTargetsLocal(); markCloudDirty(); }
function restoreTargets(){ try{ const raw = localStorage.getItem(TARGETS_STORAGE_KEY); if(raw) state.targets = {...DEFAULT_TARGETS, ...JSON.parse(raw)}; }catch(e){} }
function persistManual(){ writeManualLocal(); markCloudDirty(); }
function restoreManual(){ try{ const raw = localStorage.getItem(MANUAL_STORAGE_KEY); if(raw){ const m = JSON.parse(raw); state.manual.month = m.month || {}; state.manual.daily = m.daily || {}; } }catch(e){ } }
// 纯本地写入（不触发云端上传）：供「应用云端快照」时落盘使用，避免 A/B 两端互相回写形成 ping-pong。
function writeTargetsLocal(){ try{ localStorage.setItem(TARGETS_STORAGE_KEY, JSON.stringify(state.targets)); }catch(e){} }
function writeManualLocal(){ try{ localStorage.setItem(MANUAL_STORAGE_KEY, JSON.stringify(state.manual)); }catch(e){} }
function writeVabLocal(){ try{ localStorage.setItem(VAB_STORAGE_KEY, JSON.stringify({ set:[...state.vabSet], list: state.vabList || [], rows: state.vabRows || [], demo: !!state.demoVab })); }catch(e){} }
// 当前报表月份键（YYYY-MM）。手动月累计覆盖按月隔离存储：
// 修复「上月手动调整的月累计残留到本月」——旧格式为扁平 {行名: 值}（无月份），新读取只认 {YYYY-MM: {行名: 值}}，
// 旧扁平键自然失效（正是本次数据异常的根源：新增业绩月累计被上月覆盖值污染，导致 基础业绩=合计−新增 变为负数）。
function curMonthKey(){
  const c = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  return c.getFullYear() + '-' + String(c.getMonth()+1).padStart(2,'0');
}

// 历史导入订单本地持久化（localStorage + IndexedDB 双层：刷新网页后自动恢复；月累计/每日数据均由订单重算得到）
// 存储订单原始行（含 _uid 稳定主键）+ 截止日期；不存派生结果，刷新后由 recompute() 重新生成，保证一致。
// 双写策略：localStorage 同步（最快恢复路径）+ IndexedDB 异步（配额大、防 localStorage 静默失败）；
// 并在 beforeunload 时兜底刷写，确保「不主动点击清除」时数据始终保留在页面上（即使某些代码路径未显式 persistOrders）。
const ORDERS_STORAGE_KEY = 'dailyPerformance_orders';
const ORDERS_IDB_NAME = 'dailyPerformance';
const ORDERS_IDB_STORE = 'kv';
// 列宽持久化键：必须在 restoreColWidths()（脚本加载时顶层调用）之前声明，否则与订单 TDZ 同理，
// const 在初始化前被访问会抛 ReferenceError，被恢复函数的 catch 静默吞掉 → 拖拽列宽刷新后不保留。
const COLW_STORAGE_KEY = 'dailyPerformance_colw';

// ---------- IndexedDB 辅助（容错：环境不支持或失败则降级到 localStorage） ----------
function openOrdersIDB(){
  return new Promise((resolve, reject) => {
    if(typeof indexedDB === 'undefined'){ reject(new Error('no-idb')); return; }
    let req;
    try{ req = indexedDB.open(ORDERS_IDB_NAME, 1); }
    catch(e){ reject(e); return; }
    req.onupgradeneeded = () => { try{ if(!req.result.objectStoreNames.contains(ORDERS_IDB_STORE)) req.result.createObjectStore(ORDERS_IDB_STORE); }catch(_){} };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('idb-open-fail'));
  });
}
function idbSet(key, value){
  return openOrdersIDB().then(db => new Promise((resolve, reject) => {
    try{
      const tx = db.transaction(ORDERS_IDB_STORE, 'readwrite');
      tx.objectStore(ORDERS_IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => reject(tx.error || new Error('idb-put-fail'));
    }catch(e){ reject(e); }
  }));
}
function idbGet(key){
  return openOrdersIDB().then(db => new Promise((resolve, reject) => {
    try{
      const tx = db.transaction(ORDERS_IDB_STORE, 'readonly');
      const r = tx.objectStore(ORDERS_IDB_STORE).get(key);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error || new Error('idb-get-fail'));
    }catch(e){ reject(e); }
  }));
}
function idbDel(key){
  return openOrdersIDB().then(db => new Promise(resolve => {
    try{
      const tx = db.transaction(ORDERS_IDB_STORE, 'readwrite');
      tx.objectStore(ORDERS_IDB_STORE).delete(key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    }catch(e){ resolve(false); }
  }));
}

// 把存储原始字符串还原到 state.orders（JSON 不保留 Date，需把 payTime 还原为 Date）
function applyOrdersRaw(raw){
  if(!raw || typeof raw !== 'string') return false;
  try{
    const obj = JSON.parse(raw);
    if(obj && Array.isArray(obj.orders)){
      state.orders = obj.orders.map(r => {
        if(r && typeof r.payTime === 'string') r.payTime = new Date(r.payTime);
        if(r && !r._uid) r._uid = nextUid();
        return r;
      });
      if(obj.cutoff) state.cutoffDate = obj.cutoff;
      if(obj.ts) state.ordersTs = Number(obj.ts) || 0;
      return true;
    }
  }catch(e){}
  return false;
}

function persistOrders(force, syncNow){
  // 空态守卫：内存 orders 为空、而已存储中存在非空数据时，绝不用空数据覆写真数据。
  // 根因：beforeunload 与多处路径都会调用 persistOrders()，若此刻内存态为空/残缺
  //（如刚加载未恢复、切走标签、渲染中断），无条件写回会把已存的真实订单清空，
  // 导致「刷新后日业绩数据消失」。仅显式清空（clearAllData 传 force=true）才允许写空。
  if(!force && (!Array.isArray(state.orders) || state.orders.length === 0)){
    let storedHasData = false;
    try{
      const raw = localStorage.getItem(ORDERS_STORAGE_KEY);
      if(raw){ const o = JSON.parse(raw); storedHasData = !!(o && Array.isArray(o.orders) && o.orders.length > 0); }
    }catch(e){ /* 解析失败按无数据处理 */ }
    if(storedHasData){
      // 内存空但存储有数据 → 不覆写、不清空，仅静默返回保护既有数据。
      return;
    }
    // 空 localStorage 的真·新设备（2026-10-19 新增）：此处空写入会凭空制造一个「更新的 snapTs」，
    // 使 syncFromCloud 裁决为 local，从而压制云端整包（订单/历史明细/VAB/月目标全部不恢复）。
    // 本地从未有过真实快照（snap_ts 不存在/为 0）→ 跳过本次空持久化；显式清空走 force=true 不受影响。
    let snapTs = 0;
    try{ snapTs = Number(localStorage.getItem(SNAP_TS_STORAGE_KEY)) || 0; }catch(e){}
    if(!snapTs){
      try{ (window.__guardLog = window.__guardLog || []).push('BLOCKED @ ' + (new Error().stack || '').split('\n').slice(1,4).join(' <- ')); }catch(_){}
      return;
    }
  }
  state.ordersTs = Date.now();
  state.snapTs = state.ordersTs;      // 订单变更同样是「整包快照」的一次变更
  saveSnapTs(state.snapTs);
  const payload = JSON.stringify({ orders: state.orders, cutoff: state.cutoffDate || null, ts: state.ordersTs });
  let lsOk = false;
  // 1) localStorage 同步（最快恢复路径，刷新/卸载前都会执行，最可靠）
  try{ localStorage.setItem(ORDERS_STORAGE_KEY, payload); lsOk = true; }
  catch(e){ /* 配额超限/隐私模式：降级到 IndexedDB */ }
  // 2) IndexedDB 异步兜底（更大配额，防 localStorage 静默失败）
  idbSet(ORDERS_STORAGE_KEY, payload).catch(()=>{});
  // 3) 桥接父窗口：跨域 / 一次性 origin 的 iframe 内部存储不可靠，转交稳定的顶层壳窗口持久化
  try{ if(window.parent && window.parent !== window) window.parent.postMessage({ type:'dp:persist', key:ORDERS_STORAGE_KEY, payload }, '*'); }catch(e){}
  showOrdersSaved(lsOk);
  // 云端持久化：订单变更后异步同步服务端（防抖 800ms），仅发布域名生效；本地/预览环境自动降级为仅本机。
  // syncNow=true 时跳过防抖立即同步——用于「导入订单」等关键落盘点，避免用户紧接着刷新/关页导致云端尚未写入。
  saveOrdersToCloud(!!syncNow);
}
function restoreOrders(){
  // 同步优先：localStorage（顶层调用不阻塞页面解析）
  let raw = null;
  try{ raw = localStorage.getItem(ORDERS_STORAGE_KEY); }catch(e){ raw = null; }
  if(raw) applyOrdersRaw(raw);
  // 异步回退：IndexedDB（localStorage 为空或不可用时兜底恢复）
  idbGet(ORDERS_STORAGE_KEY).then(saved => {
    if(saved && typeof saved === 'string' && !raw){
      if(applyOrdersRaw(saved) && typeof _postRestore === 'function') _postRestore();
    }
  }).catch(()=>{});
  // 桥接父窗口：向壳请求持久化数据（与壳在 iframe load 时主动推送互补，覆盖壳监听尚未就绪等时序）
  try{ if(window.parent && window.parent !== window) window.parent.postMessage({ type:'dp:request', key:ORDERS_STORAGE_KEY }, '*'); }catch(e){}
}
// IDB 异步恢复成功后由 initApp 注入的回调：重新渲染并触发重算
let _postRestore = null;
// 保存状态可见化（依赖 HTML 中的 #save-status 元素）
function showOrdersSaved(ok){
  const el = $('save-status');
  if(!el) return;
  if(state.orders.length === 0){ el.textContent = ''; el.className = 'file-meta'; return; }
  el.textContent = ok ? '· 已本地保存' : '· 已本地保存(IDB)';
  el.className = 'file-meta';
}
// 接收套壳父窗口推送 / 应答的持久化数据（跨域 iframe 内部存储不可靠时的恢复通道）。
// 仅当本页尚未自行恢复（state.orders 为空）时应用，避免同源场景下的重复重渲染。
// 直接调用 runRestoredRecompute（函数声明，随时可用）：即便 initApp 已先以空数据渲染，
// 此处也会用恢复出的数据重新渲染 + 重算，确保表格不空白。
window.addEventListener('message', (e) => {
  const m = (e && e.data) || null;
  if(!m || m.type !== 'dp:payload' || !m.key) return;
  if(m.key === ORDERS_STORAGE_KEY){
    if(state.orders.length > 0) return;
    if(applyOrdersRaw(m.payload)) runRestoredRecompute();
    return;
  }
  // 近一年消费明细：壳窗口推送 / 应答恢复（恢复后缓存失效；订单已在场时立即重算，
  // 2026-10-20 修复「待下次重算生效」实际等不到下次重算 → 刷新后新客规则失效）
  if(m.key === HIST_STORAGE_KEY){
    if(state.histOrders.length > 0) return;
    if(applyHistRaw(m.payload)){
      invalidateMktKeyCaches();
      // 2026-10-23 修复：壳窗口桥接恢复时也恢复 histGz，避免该设备后续订单保存把云端 hist 清空
      try{ const _o = JSON.parse(m.payload); if(_o && typeof _o.gz === 'string'){ state.histGz = { gz: _o.gz, histTs: Number(_o.ts) || 0 }; } }catch(_){}
      if(state.orders.length > 0 && typeof runRestoredRecompute === 'function') runRestoredRecompute({ silent:true });
    }
    return;
  }
});

/* =================== 近一年消费明细（历史）本地持久化 ===================
 * 用途：2026-10-16 新增「营销部顾客一年内无消费记录 → 视为初诊客户」判定规则的数据源。
 * 与订单表相互独立：仅参与初诊判定，不计入任何业绩金额；未导入时该规则自动停用（保持既有口径）。
 * 持久化与订单同构：localStorage 同步 + IndexedDB 兜底 + 壳窗口桥接（dp:persist / dp:request）。 */
const HIST_STORAGE_KEY = 'dailyPerformance_hist';
// hist 是否已成功上云/从云端恢复（用于状态文案；2026-10-20 起仅明细也可上云，故不再等同「订单已导入」）
let histClouded = false;
// VAB 名单是否已成功上云/从云端恢复（用于状态文案；随订单整包上云，本机无订单时走 saveVabOnlyToCloud 单独上云）
let vabClouded = false;

function applyHistRaw(raw){
  if(!raw || typeof raw !== 'string') return false;
  try{
    const obj = JSON.parse(raw);
    if(obj && Array.isArray(obj.histOrders)){
      state.histOrders = obj.histOrders.map(r => {
        if(r && typeof r.payTime === 'string') r.payTime = new Date(r.payTime);
        if(r && !r._uid) r._uid = nextUid();
        return r;
      });
      if(obj.ts) state.histTs = Number(obj.ts) || 0;
      return true;
    }
  }catch(e){}
  return false;
}

/* ---- hist 上云压缩（2026-10-19）：明细行数大（实测 2.6 万行瘦行后约 3MB，整包超云端网关请求体限制报 NetworkError），
   用浏览器原生 CompressionStream gzip 压缩后再以 base64 存入 payload.hist.gz（明细重复度高，约压缩 10 倍）。
   不可用环境（旧浏览器）自动降级：不上传 hist（仅本地），订单数据不受影响。 ---- */
async function gzipToB64(obj){
  const json = JSON.stringify(obj);
  if(typeof CompressionStream === 'undefined') throw new Error('no CompressionStream');
  const stream = new Blob([json]).stream().pipeThrough(new CompressionStream('gzip'));
  const buf = new Uint8Array(await new Response(stream).arrayBuffer());
  let bin = ''; const CH = 0x8000;
  for(let i = 0; i < buf.length; i += CH) bin += String.fromCharCode.apply(null, buf.subarray(i, i + CH));
  return btoa(bin);
}
async function b64ToObj(b64){
  const bin = atob(b64);
  const buf = new Uint8Array(bin.length);
  for(let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
  const stream = new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'));
  return JSON.parse(await new Response(stream).text());
}
/* ---- 云端读写加固（2026-10-24）----
   现象：整包 payload 涨到 ~1.4MB 后，弱网/代理环境下 SELECT/UPDATE 偶发 fetch 网络层中断，
   界面报「云端保存失败：TypeError: NetworkError when attempting to fetch resource」
   （实测 08:11:39 报错、08:11:48 同网络同尺寸写入成功 → 瞬时抖动，重试即可自愈）。
   措施：① 订单数组随快照 gzip 压缩（ordersGz，复用 hist 的压缩设施，整包 1.39MB → ~0.6MB）；
         ② 网络层错误自动重试 2 次（1.2s / 3s 退避），业务错误（权限/字段等）不重试直接上抛。 ---- */
const NET_ERR_RE = /networkerror|network error|failed to fetch|fetch failed|load failed|when attempting to fetch/i;
// 包一层 SDK 请求：网络层错误自动重试；最终仍失败时在 message 上标注重试次数
async function cloudRun(makeReq, tries){
  const N = tries || 3;
  let lastMsg = '';
  for(let i = 0; i < N; i++){
    if(i) await new Promise(r => setTimeout(r, i === 1 ? 1200 : 3000));
    try{
      const { data, error } = await makeReq();
      if(!error) return { data, error: null };
      lastMsg = String(error.message || '');
      if(!NET_ERR_RE.test(lastMsg)) return { data, error };
    }catch(e){
      lastMsg = String((e && e.message) || e);
      if(!NET_ERR_RE.test(lastMsg)) throw e;
    }
  }
  return { data: null, error: { message: '网络异常（已自动重试 ' + (N - 1) + ' 次）' } };
}
// 网络类错误转友好文案（最终失败提示用）
function netErrMsg(m){ return NET_ERR_RE.test(String(m || '')) ? '网络异常（请求未送达，请检查网络后重试）' : (m || ''); }
// 解出云端快照里的订单数组：orders 数组 = 旧格式（直读）；ordersGz = 压缩格式（2026-10-24 起）
async function cloudOrdersArr(payload){
  if(payload && Array.isArray(payload.orders)) return payload.orders;
  if(payload && payload.ordersGz){
    try{ const arr = await b64ToObj(payload.ordersGz); if(Array.isArray(arr)) return arr; }catch(e){}
  }
  return [];
}
// hist 压缩缓存（内存态，不持久化）：persistHist / 云端恢复 时填充，cloudPayload 同步取用
function slimHistRows(rows){
  return (rows || []).map(r => ({
    memberNo: r.memberNo || '', name: r.name || '',
    payTime: (r.payTime instanceof Date) ? r.payTime.toISOString() : (r.payTime || ''),
    _uid: r._uid
  }));
}
function queueHistGz(){
  if(!(state.histOrders || []).length){ state.histGz = null; return; }
  gzipToB64({ histOrders: slimHistRows(state.histOrders), ts: state.histTs || 0 })
    .then(gz => { state.histGz = { gz, histTs: state.histTs || 0 }; })
    .catch(() => { state.histGz = null; }); // 压缩失败：本次上传不带 hist（仅本地），订单不受影响
}
function persistHist(){
  state.histTs = Date.now();
  const basePayload = JSON.stringify({ histOrders: state.histOrders, ts: state.histTs });
  // 先落盘「无 gz 版」（保证本机可读）；gz 压缩完成后会就地补写（见 .then），
  // 使刷新后 state.histGz 可直接从本机恢复 —— 否则重载后下次订单保存会把云端 hist 清空（2026-10-23 修复）。
  try{ localStorage.setItem(HIST_STORAGE_KEY, basePayload); }catch(e){ /* 配额超限：IndexedDB 兜底 */ }
  idbSet(HIST_STORAGE_KEY, basePayload).catch(()=>{});
  try{ if(window.parent && window.parent !== window) window.parent.postMessage({ type:'dp:persist', key:HIST_STORAGE_KEY, payload: basePayload }, '*'); }catch(e){}
  if((state.histOrders||[]).length === 0) return; // 无任何明细可传
  // 云端持久化（2026-10-20）：hist 始终上云，不再要求「必须有订单」。
  //   有订单 → 整包上传（含 hist）；无订单（仅明细）→ 合并云端已有订单后上传，避免清空云端订单。
  bumpSnapTs();
  gzipToB64({ histOrders: slimHistRows(state.histOrders), ts: state.histTs || 0 })
    .then(gz => {
      state.histGz = { gz, histTs: state.histTs || 0 };
      // 补写带 gz 的快照到本机（localStorage + IDB + 壳窗口桥接），供刷新后直接恢复 histGz（2026-10-23 修复）
      const withGz = JSON.stringify({ histOrders: state.histOrders, ts: state.histTs, gz });
      try{ localStorage.setItem(HIST_STORAGE_KEY, withGz); }catch(e){}
      idbSet(HIST_STORAGE_KEY, withGz).catch(()=>{});
      try{ if(window.parent && window.parent !== window) window.parent.postMessage({ type:'dp:persist', key:HIST_STORAGE_KEY, payload: withGz }, '*'); }catch(e){}
      if(Array.isArray(state.orders) && state.orders.length > 0){
        // 注意：histClouded 在 saveOrdersToCloud 真正上传成功后才置真（见其 success 分支），
        // 不能在此处提前置真——否则无云端/未登录时会误报「已同步云端」。
        saveOrdersToCloud(true);
      }else{
        saveHistOnlyToCloud(gz);
      }
    })
    .catch(() => {
      if(Array.isArray(state.orders) && state.orders.length > 0){ state.histGz = null; saveOrdersToCloud(true); }
      // 压缩失败且无订单：明细仅留本机（IDB/localStorage），下次有订单时再上云
    });
}

// 安全解析 localStorage/IDB 里的 hist 字符串 → { orders, ts }
function safeParseHist(raw){
  if(!raw || typeof raw !== 'string') return null;
  try{ const o = JSON.parse(raw); if(o && Array.isArray(o.histOrders)) return { orders: o.histOrders, ts: Number(o.ts) || 0 }; }catch(e){}
  return null;
}
function restoreHist(){
  let raw = null;
  histClouded = false; // 本地恢复默认未上云（云端恢复会在 syncFromCloud 兜底里置真）
  try{ raw = localStorage.getItem(HIST_STORAGE_KEY); }catch(e){ raw = null; }
  const lsParsed = safeParseHist(raw);
  // 合并恢复（2026-10-20 修复「陈旧 localStorage 遮蔽新 IDB 数据」）：
  // 同源两路——localStorage（快、但共享 5MB 配额易因其他页面数据超限而写入失败）与 IndexedDB（大容量可靠）。
  // 两路都拿到后选「行数更多」者；IDB 优先于同等数量的 localStorage（localStorage 失败时残留的旧小快照会
  // 遮蔽新写入的 IDB 大体积数据，表现为「导入后刷新数据被回退/清除」）。绝不直接用 localStorage 同步覆盖。
  idbGet(HIST_STORAGE_KEY).then(saved => {
    const idbParsed = safeParseHist(saved && typeof saved === 'string' ? saved : null);
    let chosen = null, chosenRaw = null;
    if(idbParsed && idbParsed.orders.length){ chosen = idbParsed; chosenRaw = saved; }
    if(lsParsed && lsParsed.orders.length && (!chosen || lsParsed.orders.length > chosen.orders.length)){ chosen = lsParsed; chosenRaw = raw; }
    if(chosen && chosen.orders.length){
      if(applyHistRaw(chosenRaw)){
        invalidateMktKeyCaches();
        if(typeof refreshHistMeta === 'function') refreshHistMeta();
        // 2026-10-23 修复：从本机快照恢复 histGz，避免重载后下次订单保存把云端 hist 清空
        try{ const _o = JSON.parse(chosenRaw); if(_o && typeof _o.gz === 'string'){ state.histGz = { gz: _o.gz, histTs: Number(_o.ts) || 0 }; } }catch(_){}
        // 2026-10-20 修复「刷新后新客规则失效」：hist 为异步恢复，订单先恢复并完成首次重算时
        // hist 尚未就位 → 报表按「规则B停用」出数；hist 恢复后必须再重算一次，否则报表停留旧口径。
        if(state.orders.length > 0 && typeof runRestoredRecompute === 'function') runRestoredRecompute({ silent:true });
        // 自愈：数据此前只在 IDB（localStorage 写失败）→ 恢复后推回云端
        if(Array.isArray(state.orders) && state.orders.length > 0){ bumpSnapTs(); saveOrdersToCloud(true); }
      }
    }
  }).catch(()=>{
    // IDB 不可用 → 退回 localStorage
    if(lsParsed && lsParsed.orders.length){ if(applyHistRaw(raw)){ invalidateMktKeyCaches(); if(typeof refreshHistMeta==='function') refreshHistMeta(); if(state.orders.length > 0 && typeof runRestoredRecompute === 'function') runRestoredRecompute({ silent:true }); try{ const _o = JSON.parse(raw); if(_o && typeof _o.gz === 'string'){ state.histGz = { gz: _o.gz, histTs: Number(_o.ts) || 0 }; } }catch(_){} } }
  });
  // 桥接父窗口：向壳索取持久化数据（与壳的主动推送互补）
  try{ if(window.parent && window.parent !== window) window.parent.postMessage({ type:'dp:request', key:HIST_STORAGE_KEY }, '*'); }catch(e){}
}

// 初诊键集合缓存失效（历史明细导入/清除/恢复后调用；随后 recompute() 重算生效）
function invalidateMktKeyCaches(){
  _histKeysCache = { sig:'', date:'', set:null };
  _mktKeysCache = { sig:'', set:null };
}

/* =================== 云端持久化（WorkBuddy Cloud Service，按账号隔离） ===================
 * 根因：发布/访问环境会抹掉 localStorage/IndexedDB 的跨会话持久性（sandbox 不透明源 / origin 漂移），
 * 导致刷新/重进丢数据。改为订单数据落地云端（服务端），刷新 / 换设备 / 关机均不丢。
 * 仅在注册域名（haibei-medical.app.workbuddy.host）生效；本地/预览环境会优雅降级为「仅本机」。 */
const CLOUD_ENDPOINT = (typeof window !== 'undefined' && window.CLOUD_ENDPOINT) || 'https://haibei-medical.app.workbuddy.host';
const CLOUD_PK = (typeof window !== 'undefined' && window.CLOUD_PK) || 'wbpk_KQXIeDHLXWDHN7WQp3zqUQ_Y5IlM6hJd7n10U8UN5bYmSGUYeevVOwl';
let _cloudClient = null;
let _cloudUser = null;
let _cloudSaveTimer = null;
let _pendingOtp = null;

function getCloud(){
  if(_cloudClient) return _cloudClient;
  if(typeof window.WorkBuddyCloud === 'undefined') return null;
  try{ _cloudClient = window.WorkBuddyCloud.createWorkBuddyCloud({ endpoint: CLOUD_ENDPOINT, publishableKey: CLOUD_PK }); }
  catch(e){ _cloudClient = null; }
  return _cloudClient;
}

/* ---- 登录态持久化（2026-10-13 需求：登录不退出的情况下保持为登录状态）----
 * 问题：SDK 会话存于 localStorage 键 workbuddy-cloud.session.<PK>，但发布环境下
 * localStorage 跨会话可能被清（origin 漂移），且 access_token 过期后若刷新失败
 * SDK 会 abandon() 清空会话 → 刷新页面就变「未登录」。
 * 方案（三层）：
 *   1. 备份：每次拿到有效会话时，把 {accessToken,refreshToken,expiresAt,user}
 *      备份到独立键 dp_auth_backup（与 SDK 键分离，防单点丢失）；
 *   2. 恢复：启动时 getSession() 为空但备份存在 → 把备份写回 SDK 存储键，
 *      再次 getSession() 触发 SDK 用 refreshToken 静默续期，成功则仍是登录态；
 *   3. 保活：页面打开期间每 20 分钟 + 每次切回前台时 getSession() 一次，
 *      让 token 在过期前完成滚动刷新，避免长期闲置后 refreshToken 整体失效。 */
const AUTH_BACKUP_KEY = 'dp_auth_backup';
const SDK_SESSION_KEY = 'workbuddy-cloud.session.' + CLOUD_PK;
function backupAuthSession(session){
  try{
    if(!session || !session.user) return;
    localStorage.setItem(AUTH_BACKUP_KEY, JSON.stringify({
      accessToken: session.accessToken, refreshToken: session.refreshToken,
      expiresAt: session.expiresAt, user: session.user, savedAt: Date.now(),
    }));
  }catch(e){}
}
function readAuthBackup(){
  try{
    const raw = localStorage.getItem(AUTH_BACKUP_KEY);
    if(!raw) return null;
    const s = JSON.parse(raw);
    return (s && s.refreshToken && s.user) ? s : null;
  }catch(e){ return null; }
}
function clearAuthBackup(){ try{ localStorage.removeItem(AUTH_BACKUP_KEY); }catch(e){} }
async function restoreAuthSession(){
  const cloud = getCloud(); if(!cloud) return null;
  const bk = readAuthBackup(); if(!bk) return null;
  try{
    // 把备份会话写回 SDK 的存储键，再走 getSession() → ensure() → refreshOnce 静默续期
    localStorage.setItem(SDK_SESSION_KEY, JSON.stringify({
      accessToken: bk.accessToken, refreshToken: bk.refreshToken,
      expiresAt: bk.expiresAt || 0, user: bk.user,
    }));
    const r = await cloud.auth.getSession();
    if(r && r.data && r.data.user){ backupAuthSession(r.data); return r.data; }
  }catch(e){}
  clearAuthBackup();
  return null;
}
let _cloudKeepaliveTimer = null;
function startCloudKeepalive(){
  if(_cloudKeepaliveTimer) return;
  _cloudKeepaliveTimer = setInterval(() => {
    const cloud = getCloud(); if(!cloud || !_cloudUser) return;
    cloud.auth.getSession().then(r => {
      if(r && r.data && r.data.user) backupAuthSession(r.data);
    }).catch(() => {});
  }, 20 * 60 * 1000);
  // 切回前台立即续一次（移动端/后台标签页休眠后回来自动保持登录）
  document.addEventListener('visibilitychange', () => {
    if(document.visibilityState !== 'visible' || !_cloudUser) return;
    const cloud = getCloud(); if(!cloud) return;
    cloud.auth.getSession().then(r => { if(r && r.data && r.data.user) backupAuthSession(r.data); }).catch(() => {});
  });
}
function setCloudStatus(text, kind){
  const el = $('cloud-status'); if(!el) return;
  el.textContent = text || '';
  el.className = 'file-meta' + (kind ? ' cloud-' + kind : '');
}
// 登录态下显示「从云端恢复 / 推送到云端」手动按钮（2026-10-24）：给用户显式控制权，
// 无论自动裁决结果如何都能一键补齐（手机端刷新后仍无数据 → 点「从云端恢复」；PC 登录后未入库 → 点「推送到云端」）。
function showCloudActions(show){
  const r = $('cloud-restore'), p = $('cloud-push');
  if(r) r.style.display = show ? '' : 'none';
  if(p) p.style.display = show ? '' : 'none';
}
function applyOrdersObj(obj){
  if(obj && Array.isArray(obj.orders)){
    state.orders = obj.orders.map(r => {
      if(r && typeof r.payTime === 'string'){ try{ r.payTime = new Date(r.payTime); }catch(_){} }
      if(r && !r._uid) r._uid = nextUid();
      return r;
    });
    if(obj.cutoff) state.cutoffDate = obj.cutoff;
    if(obj.ts) state.ordersTs = Number(obj.ts) || 0;
    return true;
  }
  return false;
}
function cloudPayload(){
  return {
    orders: state.orders,
    cutoff: state.cutoffDate || null,
    // 同账号跨设备同步范围扩展（2026-10-16）：VAB 会员名单 / 月目标 / 手动覆盖也随快照上云，
    // 否则换设备后会出现「订单在、VAB 名单为 0、月目标回落默认值」的割裂状态。
    vab: { list: state.vabList || [], rows: state.vabRows || [], demo: !!state.demoVab },
    targets: state.targets,
    manual: state.manual,
    // 近一年消费明细随快照上云（2026-10-19）：瘦行仅保留判定所需 4 字段（memberNo/name/payTime/_uid，见 histKeysCached）
    hist: state.histGz || undefined,
    ts: state.snapTs || Date.now(),
  };
}
// 应用云端整包快照：订单 + VAB 会员名单 + 月目标 + 手动覆盖。
// 期间置 _applyingCloud，禁止任何写入路径把「刚拉下来的云端数据」又立刻推回云端（防 ping-pong）。
let _applyingCloud = false;
async function applyCloudObj(obj){
  if(!obj || typeof obj !== 'object') return false;
  _applyingCloud = true;
  let touched = false;
  try{
    // 订单：orders 数组 = 旧格式直读；ordersGz = 压缩格式（2026-10-24）→ 先解压再应用。
    // 不改写入参 obj（调用方随后还要原样落 localStorage）。
    if(!Array.isArray(obj.orders) && obj.ordersGz){
      try{
        const arr = await b64ToObj(obj.ordersGz);
        if(Array.isArray(arr)) obj = Object.assign({}, obj, { orders: arr });
      }catch(e){}
    }
    if(applyOrdersObj(obj)) touched = true;
    const v = obj.vab;
    if(v && typeof v === 'object'){
      // 只有快照里「确实带了 vab 字段」才处理（旧版快照没有该字段 → 不动本机，避免误清空）；
      // 带了字段但为空数组 = 云端显式清空 VAB → 同步清空本机。
      const hasVabKey = Array.isArray(v.rows) || Array.isArray(v.list);
      if(hasVabKey){
        const rows = (Array.isArray(v.rows) && v.rows.length) ? v.rows
                   : ((Array.isArray(v.list) && v.list.length) ? v.list.map(x => ({'会员号': x})) : []);
        setVabRows(rows);
        state.demoVab = rows.length ? !!v.demo : false;
        writeVabLocal();
        vabClouded = true; // 从云端恢复 = 名单已在云端
        refreshVabStatus(); // 立即刷新状态文案（含「已同步云端」），不等外部调用
        touched = true;
      }
    }
    if(obj.targets && typeof obj.targets === 'object' && Object.keys(obj.targets).length){
      state.targets = Object.assign({}, DEFAULT_TARGETS, obj.targets);
      writeTargetsLocal();
      // 必须同步重绘「② 月目标」输入框：renderReportInner() 里的 readTargets() 会把 DOM 输入值
      // 反向写回 state.targets，若 DOM 仍是旧值（如默认 360），刚恢复的云端值（350）会被立刻覆盖回去。
      if(typeof renderTargets === 'function') renderTargets();
      touched = true;
    }
    const h = obj.hist;
  if(h && typeof h === 'object'){
    // 只有快照里确实带了 hist 字段才处理（旧版快照没有 → 不动本机，避免误清空）；
    // gz = 压缩格式（现行为）；histOrders 数组 = 旧格式（直读，兼容历史快照）。
    // 分支内不调 persistHist（_applyingCloud 期间防回写）。
    if(h.gz){
      b64ToObj(h.gz).then(obj2 => {
        if(obj2 && Array.isArray(obj2.histOrders)){
          // 先置云端标志，再刷新状态文案——否则 refreshHistMeta 在 histClouded=false 时定格为「暂仅保存在本机」
          histClouded = true; // 从云端恢复了 hist
          state.histGz = { gz: h.gz, histTs: h.histTs || 0 }; // 缓存 gz，回写时复用
          if(applyHistRaw(JSON.stringify({ histOrders: obj2.histOrders, ts: h.histTs || Number(obj.ts) || Date.now() }))){
            invalidateMktKeyCaches();
            if(typeof refreshHistMeta === 'function') refreshHistMeta();
            if(typeof recompute === 'function') recompute(); // 异步解压完成后再重算一次（⑤清单判定）
          }
        }
      }).catch(() => {});
      touched = true;
    } else if(Array.isArray(h.histOrders)){
      histClouded = true; // 从云端恢复了 hist（旧版未压缩路径）
      if(applyHistRaw(JSON.stringify({ histOrders: h.histOrders, ts: h.histTs || Number(obj.ts) || Date.now() }))){
        invalidateMktKeyCaches();
        if(typeof refreshHistMeta === 'function') refreshHistMeta();
        if(state.orders.length > 0 && typeof runRestoredRecompute === 'function') runRestoredRecompute({ silent:true }); // 与 gz 分支对齐：恢复后重算
        touched = true;
      }
    }
  }
  if(obj.manual && typeof obj.manual === 'object'){
      state.manual.month = obj.manual.month || {};
      state.manual.daily = obj.manual.daily || {};
      writeManualLocal();
      touched = true;
    }
    if(Number(obj.ts)){
      state.ordersTs = Number(obj.ts) || 0;
      state.snapTs = Number(obj.ts) || 0;
      saveSnapTs(state.snapTs);
    }
  }finally{
    _applyingCloud = false;
  }
  return touched;
}
// 标记整包快照为「脏」：更新版本号并触发（防抖）上传。orders / VAB / 月目标 / 手动覆盖 任一变更均调用。
function markCloudDirty(force){
  if(_applyingCloud) return;              // 正在应用云端快照 → 不回写
  bumpSnapTs();
  // 无订单时不推送（清空流程走 clearOrdersInCloud），避免把云端已有订单误覆盖成空快照
  if(Array.isArray(state.orders) && state.orders.length > 0) saveOrdersToCloud(!!force);
}
// 取一组订单的「最新月份」（形如 2026-10）；无有效日期返回 ''。用于本地/云端新旧裁决。
function monthOfOrders(orders){
  let max = '';
  if(!Array.isArray(orders)) return max;
  for(const r of orders){
    if(!r || !r.payTime) continue;
    const d = (r.payTime instanceof Date) ? r.payTime : new Date(r.payTime);
    if(isNaN(d)) continue;
    const m = d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0');
    if(m > max) max = m;
  }
  return max;
}

async function syncFromCloud(opts){
  const silent = !!(opts && opts.silent);   // 后台轮询触发：不弹恢复提示
  const forcePull = !!(opts && opts.forcePull); // 「从云端恢复」手动按钮：忽略本地裁决，强制采用云端
  const cloud = getCloud();
  if(!cloud){ setCloudStatus('云端组件加载中…（多源回退）请稍候', 'warn'); return; }
  let session = null;
  try{ const r = await cloud.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
  if(!session){
    // 会话丢失（SDK 存储被清/token 过期 abandon）→ 尝试用本地备份静默恢复登录态
    session = await restoreAuthSession();
  }
  if(!session){
    _cloudUser = null;
    showCloudActions(false);
    // 不自动弹窗（避免每次刷新都打断）；把状态条变为可点击的登录入口，用户需要时再登录
    setCloudStatus('未登录 · 未同步云端', 'warn');
    return;
  }
  _cloudUser = session.user || null;
  showCloudActions(true);       // 已登录 → 显示手动「从云端恢复 / 推送到云端」按钮
  backupAuthSession(session);   // 保持登录：每次拿到有效会话都刷新本地备份
  startCloudKeepalive();        // 页面打开期间定时续期，避免长期闲置后 refreshToken 失效
  setCloudStatus('已登录 · 同步云端', 'ok');
  try{
    const { data, error } = await cloudRun(() => cloud.database.from('dp_orders')
      .select('id, payload, updated_at').order('updated_at', { ascending:false }).limit(1));
    if(error){ setCloudStatus('云端读取失败：' + netErrMsg(error.message), 'err'); return; }
    const row = Array.isArray(data) ? data[0] : data;
    const snap = (row && row.payload) ? row.payload : null; // 函数级快照引用，供下方 hist 兜底块使用（data2 在块内声明，块外不可见）
    if(row && row.payload){
      const data2 = row;
      // 整包新旧裁决（2026-10-16 重写）：orders + VAB 名单 + 月目标 + 手动覆盖 视为「一个快照」整体裁决，
      // 避免出现「订单采用云端、VAB/月目标却仍是本机旧值」的割裂（用户实测：同账号两台机器，订单一致，
      // 但一台 VAB 64 人/月目标 350，另一台 VAB 0 人/月目标 360）。
      // 规则（ts = 本机最后一次本地变更时刻 state.snapTs，随快照上云）：
      //   - 本机无任何本地变更记录（snapTs=0）→ 采用云端；
      //   - cloudTs > localTs → 采用云端（其他设备更新）；
      //   - cloudTs < localTs → 保留本地并推回云端；
      //   - ts 相同 → 订单条数多者胜；完全持平则判定「已一致」，不动作（避免每次刷新都无谓上传/互相覆盖）。
      // 安全网：云端订单月份落后于本地 → 判定云端为陈旧快照，保留本地（防止旧快照把新数据倒退）。
      const cloudTs   = Number(data2.payload && data2.payload.ts) || 0;
      const localTs   = Number(state.snapTs) || 0;
      const cloudOrdersA = await cloudOrdersArr(data2.payload); // 兼容 orders / ordersGz 两种云端格式
      const cloudCnt  = cloudOrdersA.length;
      const localCnt  = Array.isArray(state.orders) ? state.orders.length : 0;
      const cloudMon  = monthOfOrders(cloudOrdersA);
      const localMon  = monthOfOrders(state.orders);
      // 本地「真实数据」标志：orders / VAB / 历史明细 任一非空。
      // 本地全空（如新设备 / 清过缓存）时 snapTs 可能被空写入抬高，此时无权压制云端——必拉云端，
      // 否则会出现「换设备后一直显示已同步云端(2 条) 但数据为空」的假同步（2026-10-19 修复）。
      const localEmpty = localCnt === 0
        && !(state.vabRows || []).length
        && !(state.histOrders || []).length;
      let verdict;
      if(!localTs || localEmpty || cloudTs > localTs) verdict = 'cloud';
      else if(cloudTs < localTs) verdict = 'local';
      else verdict = (cloudCnt === localCnt) ? 'same' : (cloudCnt > localCnt ? 'cloud' : 'local');
      // 「从云端恢复」手动按钮：跳过本地裁决，强制以云端为准（换设备/数据缺失时一键拉全量）
      if(forcePull) verdict = 'cloud';
      // 安全网：云端订单月份落后于本地 → 判定云端为陈旧快照，保留本地（防止旧快照把新数据倒退）；
      // 手动「从云端恢复」时用户意图明确，不触发该回退。
      if(!forcePull && verdict === 'cloud' && localMon && cloudMon && cloudMon < localMon) verdict = 'local';
      // 升级迁移（2026-10-16）：旧版快照不含 vab/targets/manual 字段，且这些老机器 state.snapTs 恒为 0，
      // 会让「本机有 VAB、云端没有」的机器误判为采用云端 → VAB 永远推不上去（另一台机器始终显示 0 人）。
      // 规则：本机无变更记录 + 云端完全没有 VAB 数据 + 本机有 VAB → 改判「本地优先推送」，把 VAB/月目标补进云端。
      const cloudHasVab = !!(data2.payload.vab && ((data2.payload.vab.list && data2.payload.vab.list.length) || (data2.payload.vab.rows && data2.payload.vab.rows.length)));
      const localHasVab = (state.vabList || []).length > 0;
      if(localTs === 0 && !cloudHasVab && localHasVab) verdict = 'local';
      const cloudVabN = (data2.payload.vab && data2.payload.vab.list && data2.payload.vab.list.length) || 0;
      if(verdict === 'same'){
        setCloudStatus('已同步云端（' + cloudCnt + ' 条 · VAB ' + cloudVabN + ' 人）', 'ok');
      }else if(verdict === 'cloud'){
        await applyCloudObj(data2.payload);
        // 云端快照落本地缓存前先解压回明文 orders（本地 localStorage 始终存明文格式）
        try{
          const plainSnap = Object.assign({}, data2.payload);
          if(!Array.isArray(plainSnap.orders) && plainSnap.ordersGz){
            plainSnap.orders = await cloudOrdersArr(plainSnap);
            delete plainSnap.ordersGz;
          }
          localStorage.setItem(ORDERS_STORAGE_KEY, JSON.stringify(plainSnap));
        }catch(e){}
        if(typeof runRestoredRecompute === 'function') runRestoredRecompute({ silent });
        try{ refreshVabStatus(); }catch(e){}
        setCloudStatus('已同步云端（' + cloudCnt + ' 条 · VAB ' + (state.vabList || []).length + ' 人）', 'ok');
      }else{
        // 本地更新 → 覆盖云端，避免云端长期停留在旧数据导致下次刷新又倒退。
        // 2026-10-24 修复：不再仅看 orders，hist 仅/VAB 仅 的本地数据也一并推上去（见 pushLocalSnapshotToCloud）。
        if(localCnt > 0 || (state.histOrders||[]).length > 0 || (state.vabList||[]).length > 0){
          await pushLocalSnapshotToCloud();
        }else{
          setCloudStatus('已同步云端（' + cloudCnt + ' 条）', 'ok');
        }
      }
    }else if((Array.isArray(state.orders) && state.orders.length > 0) || (state.histOrders||[]).length > 0 || (state.vabList||[]).length > 0){
      // 云端尚无任何记录（首次登录）→ 把本机已有数据（订单/hist/VAB 任一）推上去，杜绝「仅 hist/VAB 设备登录后永不入库」
      await pushLocalSnapshotToCloud();
    }else{
      setCloudStatus('已登录 · 云端暂无数据', 'ok');
    }
    // hist 兜底（2026-10-20）：云端快照含 hist 而本机恢复后为空/陈旧 → 强制回灌，
    // 不依赖整体 verdict（尤其 verdict==='same' 时整包不会被 applyCloudObj，会漏掉 hist，导致刷新后明细丢失）。
    try{
      const ch = snap && snap.hist;
      if(ch && typeof ch === 'object'){
        const chTs = Number(ch.histTs) || Number(snap.ts) || 0;
        const loTs = Number(state.histTs) || 0;
        if((state.histOrders||[]).length === 0 || chTs > loTs){
          applyCloudObj({ hist: ch, ts: snap.ts });
        }else{
          // 本机已是最新（ts 持平）→ 不重复回灌，但云端确有该明细，置云端标志避免误显「暂仅保存在本机」
          histClouded = true;
          if(typeof refreshHistMeta === 'function') refreshHistMeta();
        }
      }
    }catch(e){}
  }catch(e){
    setCloudStatus('云端同步异常：' + (e && e.message ? e.message : '网络错误'), 'err');
  }
}
async function saveOrdersToCloud(force){
  if(_cloudSaveTimer){ clearTimeout(_cloudSaveTimer); _cloudSaveTimer = null; }
  // 失败可见性（2026-10-16）：force=true 表示用户显式动作（导入 / 清空 / 重算）触发的保存，
  // 此时若云端没写成，除状态条外再弹一次 toast，避免「以为已存云端、其实只在本机」的静默失败。
  const failSave = (msg) => {
    setCloudStatus(msg, 'err');
    if(force) showToast(msg, 'err', 5200);
  };
  const doSave = async () => {
    _cloudSaveTimer = null;
    const cloud = getCloud();
    if(!cloud){ failSave('云端组件未就绪：数据仅保存在本机，请刷新页面重试'); return; }
    let session = null;
    try{ const r = await cloud.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
    if(!session){ failSave('未登录云端：数据仅保存在本机，登录成功后才会同步'); return; }
    const payload = cloudPayload();
    // 空数据不写云端（清数据走 clearOrdersInCloud），避免把好的云端数据覆盖/误删成空的
    if(payload.orders.length === 0) return;
    try{
      // 显式按 updated_at 倒序取最新一条并 upsert，避免 maybeSingle() 在存在多条历史残留记录时
      // 直接报错导致「云端保存失败」→ 云端长期停留在旧快照，刷新后把新数据倒退覆盖。
      const { data: rows, error: selErr } = await cloudRun(() => cloud.database.from('dp_orders')
        .select('id, updated_at, payload').order('updated_at', { ascending:false }).limit(1));
      if(selErr){ failSave('云端保存失败：' + netErrMsg(selErr.message)); return; }
      const latest = Array.isArray(rows) && rows.length ? rows[0] : null;
      // 跨设备保护（2026-10-20）：本地无 hist 但云端已有 hist → 保留云端 hist，避免「仅订单设备」上传时把云端明细清空
      if(!payload.hist && latest && latest.payload && latest.payload.hist){ payload.hist = latest.payload.hist; }
      // 整包上传前压缩：订单数组 gzip（ordersGz），整包 1.39MB → ~0.6MB，
      // 防 payload 超云端网关请求体限制报 NetworkError（与 hist 压缩同思路，2026-10-24）。
      // 压缩不可用的旧浏览器自动保留明文 orders 兜底；云端读取侧两种格式都兼容。
      const upPayload = Object.assign({}, payload);
      if(typeof CompressionStream !== 'undefined'){
        try{ upPayload.ordersGz = await gzipToB64(payload.orders); delete upPayload.orders; }catch(e){}
      }
      if(latest && latest.id != null){
        // 跨设备保护：云端已有比本地更新的数据（如另一台设备刚写入）且非强制保存时，
        // 不用过期本地数据覆盖云端，改为拉取云端保持一致
        const cloudTsNow = Number(latest.payload && latest.payload.ts) || 0;
        if(!force && cloudTsNow > (Number(state.snapTs) || 0)){
          setCloudStatus('云端数据更新（其他设备已保存），已跳过本次上传', 'warn');
          return;
        }
        const { error } = await cloudRun(() => cloud.database.from('dp_orders').update({ payload: upPayload, updated_at: new Date().toISOString() }).eq('id', latest.id));
        if(error){ failSave('云端保存失败：' + netErrMsg(error.message)); return; }
        // 清理可能存在的多余旧记录，保证下次读取只命中最新一条
        try{
          const { data: others } = await cloudRun(() => cloud.database.from('dp_orders').select('id'));
          if(Array.isArray(others)){
            for(const r of others){
              if(r && r.id != null && r.id !== latest.id){
                try{ await cloud.database.from('dp_orders').delete().eq('id', r.id); }catch(_){}
              }
            }
          }
        }catch(_){}
      }else{
        const { error } = await cloudRun(() => cloud.database.from('dp_orders').insert({ payload: upPayload, updated_at: new Date().toISOString() }));
        if(error){ failSave('云端保存失败：' + netErrMsg(error.message)); return; }
      }
      state.ordersTs = Number(payload.ts) || Date.now();
      state.snapTs = state.ordersTs; saveSnapTs(state.snapTs);
      setCloudStatus('已同步云端（' + payload.orders.length + ' 条 · VAB ' + (payload.vab && payload.vab.list ? payload.vab.list.length : 0) + ' 人）', 'ok');
      // hist 随整包上云成功 → 同步置云端标志并刷新状态文案（修复「导入后状态文案滞后为『暂仅保存在本机』」）
      if(payload.hist){ histClouded = true; }
      vabClouded = true; // VAB 恒随整包上云
      if(typeof refreshHistMeta === 'function') refreshHistMeta();
      refreshVabStatus();
    }catch(e){
      failSave('云端保存异常：' + (e && e.message ? e.message : '网络错误'));
    }
  };
  if(force){ await doSave(); } else { _cloudSaveTimer = setTimeout(doSave, 800); }
}
// 仅明细（无订单）时上传：把 hist 合并进云端已有订单快照后 upsert，避免「仅明细设备」把云端订单清空（2026-10-20）
async function saveHistOnlyToCloud(gz){
  const cloud = getCloud();
  if(!cloud){ setCloudStatus('云端组件未就绪：明细仅保存在本机', 'err'); return; }
  let session = null;
  try{ const r = await cloud.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
  if(!session){ setCloudStatus('未登录云端：明细仅保存在本机', 'err'); return; }
  try{
    const { data: rows, error } = await cloudRun(() => cloud.database.from('dp_orders').select('id, payload').order('updated_at',{ascending:false}).limit(1));
    if(error){ setCloudStatus('云端读取失败：'+netErrMsg(error.message), 'err'); return; }
    const latest = Array.isArray(rows) && rows.length ? rows[0] : null;
    const base = (latest && latest.payload && typeof latest.payload === 'object') ? latest.payload
               : { orders: [], cutoff: null, vab: { list: [], rows: [], demo: false }, targets: {}, manual: {} };
    const payload = Object.assign({}, base, { hist: { gz, histTs: state.histTs || 0 }, ts: state.snapTs || Date.now() });
    if(latest && latest.id != null){
      const { error: e2 } = await cloudRun(() => cloud.database.from('dp_orders').update({ payload, updated_at: new Date().toISOString() }).eq('id', latest.id));
      if(e2){ setCloudStatus('云端保存失败：'+netErrMsg(e2.message), 'err'); return; }
    }else{
      const { error: e2 } = await cloudRun(() => cloud.database.from('dp_orders').insert({ payload, updated_at: new Date().toISOString() }));
      if(e2){ setCloudStatus('云端保存失败：'+netErrMsg(e2.message), 'err'); return; }
    }
    state.histGz = { gz, histTs: state.histTs || 0 };
    histClouded = true;
    state.ordersTs = payload.ts; state.snapTs = payload.ts; saveSnapTs(state.snapTs);
    setCloudStatus('已同步云端（仅明细 · '+(state.histOrders||[]).length+' 行）', 'ok');
  }catch(e){ setCloudStatus('云端同步异常：'+(e&&e.message||'网络错误'), 'err'); }
}
// 仅 VAB 名单（本机无订单）时上云：合并进云端已有快照后 upsert，避免「仅 VAB 设备」把云端订单/明细清空（2026-10-24）
async function saveVabOnlyToCloud(){
  _vabCloudTimer = null;
  const cloud = getCloud();
  if(!cloud){ setCloudStatus('云端组件未就绪：VAB 名单仅保存在本机', 'err'); return; }
  let session = null;
  try{ const r = await cloud.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
  if(!session){ setCloudStatus('未登录云端：VAB 名单仅保存在本机', 'err'); return; }
  try{
    const { data: rows, error } = await cloudRun(() => cloud.database.from('dp_orders').select('id, payload').order('updated_at',{ascending:false}).limit(1));
    if(error){ setCloudStatus('云端读取失败：'+netErrMsg(error.message), 'err'); return; }
    const latest = Array.isArray(rows) && rows.length ? rows[0] : null;
    // 跨设备保护：云端已有比本地更新的整包（其他设备刚写入）→ 不用过期本机 VAB 覆盖，保持一致即可
    const cloudTsNow = Number(latest && latest.payload && latest.payload.ts) || 0;
    if(latest && cloudTsNow > (Number(state.snapTs) || 0)){
      setCloudStatus('云端数据更新（其他设备已保存），已跳过本次上传', 'warn');
      return;
    }
    const base = (latest && latest.payload && typeof latest.payload === 'object') ? latest.payload
               : { orders: [], cutoff: null, vab: { list: [], rows: [], demo: false }, targets: {}, manual: {} };
    const payload = Object.assign({}, base, { vab: { list: state.vabList || [], rows: state.vabRows || [], demo: !!state.demoVab }, ts: state.snapTs || Date.now() });
    if(latest && latest.id != null){
      const { error: e2 } = await cloudRun(() => cloud.database.from('dp_orders').update({ payload, updated_at: new Date().toISOString() }).eq('id', latest.id));
      if(e2){ setCloudStatus('云端保存失败：'+netErrMsg(e2.message), 'err'); return; }
    }else{
      const { error: e2 } = await cloudRun(() => cloud.database.from('dp_orders').insert({ payload, updated_at: new Date().toISOString() }));
      if(e2){ setCloudStatus('云端保存失败：'+netErrMsg(e2.message), 'err'); return; }
    }
    state.snapTs = payload.ts; saveSnapTs(state.snapTs);
    vabClouded = true;
    setCloudStatus('已同步云端（VAB ' + vabUniqueCount() + ' 人）', 'ok');
    refreshVabStatus();
  }catch(e){ setCloudStatus('云端同步异常：'+(e&&e.message||'网络错误'), 'err'); }
}
// 登录态下把「本机已有但云端可能缺失」的数据推上去（2026-10-24 修复根因）：
// 旧逻辑仅当 state.orders 非空时才上传，导致「只导入了 hist 明细 / 只导入了 VAB」的设备在登录云端后
// 永远走 else 分支报「云端暂无数据」，数据从不上云，换设备/手机端永远拉不到。
// 优先级：有订单 → 整包上传（含 hist/VAB）；否则 有 hist → 仅明细上云；否则 有 VAB → 仅 VAB 上云。
async function pushLocalSnapshotToCloud(){
  if(Array.isArray(state.orders) && state.orders.length > 0){ await saveOrdersToCloud(true); return; }
  if((state.histOrders||[]).length > 0){
    if(state.histGz && state.histGz.gz){ await saveHistOnlyToCloud(state.histGz.gz); }
    else { await persistHist(); } // 本机有明文明细但无 gz 缓存 → 重新压缩并上云
    return;
  }
  if((state.vabList||[]).length > 0){ await saveVabOnlyToCloud(); return; }
}
async function clearOrdersInCloud(){
  const cloud = getCloud(); if(!cloud) return;
  let session = null;
  try{ const r = await cloud.auth.getSession(); session = r.data; if(r.error) session = null; }catch(e){ session = null; }
  if(!session) return;
  try{
    const { data: rows } = await cloud.database.from('dp_orders').select('id');
    if(Array.isArray(rows)){
      for(const r of rows){ if(r && r.id != null){ try{ await cloud.database.from('dp_orders').delete().eq('id', r.id); }catch(_){} } }
    }
  }catch(e){}
}

/* 登录浮层已移除（2026-10-14）：登录/改密/退出统一走整站登录页 + 右下角账号面板（shared/auth-gate.js） */

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
      // allowEmpty：清空输入 = 删除该格手动覆盖（onSave 收到 ''），恢复自动计算值
      if(save && raw === '' && opts.allowEmpty){ onSave(''); return; }
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
  // VAB 名单同样属于「整包快照」的一部分：变更后打脏标记，随订单一起上云（修复「换机后 VAB 名单为 0」）。
  // 本机无订单时 markCloudDirty 的「无订单不推送」保护会跳过 → 改走 saveVabOnlyToCloud 单独上云（2026-10-24），
  // 避免「新设备先传 VAB、后传订单」期间 VAB 长期仅存本机。
  vabClouded = false; refreshVabStatus();
  if(Array.isArray(state.orders) && state.orders.length > 0){ markCloudDirty(); }
  else { queueVabCloudSave(); }
}
// VAB 单独上云防抖（与订单保存的 _cloudSaveTimer 相互独立，避免互相取消）
let _vabCloudTimer = null;
function queueVabCloudSave(){
  if(_vabCloudTimer) clearTimeout(_vabCloudTimer);
  _vabCloudTimer = setTimeout(saveVabOnlyToCloud, 800);
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
restoreHist(); // 恢复近一年消费明细（刷新网页后保留，用于「一年内无消费→初诊」新客判定）
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
    meta.textContent = `✔ 已保存为默认（${vabUniqueCount()}人）· 重新上传/编辑可覆盖更新 · ` + (vabClouded ? '已同步云端' : '暂仅保存在本机');
    meta.className = 'chip ok';
  }
  renderVabList(); // 同步刷新页面上的会员清单
}
// 在页面上渲染 VAB 会员清单（按导入表格的全部字段展示），默认展开呈现，支持单元格双击编辑 + 行内删除 + 页面内新增
/* ---- VAB 会员清单分页状态（默认 10 条/页，可选 10/20/50/100） ---- */
let vabPage = 1, vabPageSize = 10;
// VAB 清单滚动容器高度自适应：上限随视口伸缩（64vh，600px 封顶），并对齐「表头+整数倍行高」，
// 避免底部行被裁一半造成「遮盖显示不全」的观感；折叠态（不可见）跳过，展开/resize 时重算
function fitVabTableWrap(wrap){
  if(!wrap || !wrap.offsetParent) return;
  const tbl = wrap.querySelector('table');
  if(!tbl) return;
  const trs = tbl.querySelectorAll('tbody tr');
  if(!trs.length) return;
  // ≤200 行：完整展开（显式 none 覆盖页面 CSS 的 .table-wrap{max-height:300px}）
  if(trs.length <= 200){ wrap.style.maxHeight = 'none'; return; }
  // 行距用相邻行 offsetTop 差值（精确整数，不受字体/亚像素测量影响）
  const rowH = trs.length > 1 ? (trs[1].offsetTop - trs[0].offsetTop) : trs[0].offsetHeight;
  if(!rowH) return;
  const theadH = trs[0].offsetTop; // 首行顶(相对表格顶) = 表头高
  const cap = Math.max(280, window.innerHeight * 0.8); // 与内滚形态 80vh 一致
  const n = Math.max(3, Math.floor((cap - theadH) / rowH));
  wrap.style.maxHeight = (theadH + n * rowH) + 'px';
  // 自校正：以实际渲染结果消除 border/box-sizing/字体晚载造成的半行偏移
  const off = (wrap.clientHeight - trs[0].offsetTop) % rowH;
  if(off !== 0) wrap.style.maxHeight = (parseFloat(wrap.style.maxHeight) - off) + 'px';
  // 字体加载会改变表头/行高 → 就绪后重算一次
  if(document.fonts && document.fonts.status !== 'loaded') document.fonts.ready.then(() => fitVabTableWrap(wrap)).catch(()=>{});
}
// 窗口尺寸变化时重算全部清单容器（全局只绑一次，防抖 180ms）
if(!window.__vabFitResizeBound){
  window.__vabFitResizeBound = true;
  window.addEventListener('resize', () => {
    clearTimeout(window.__vabFitT);
    window.__vabFitT = setTimeout(() => {
      document.querySelectorAll('.vab-table-wrap').forEach(w => fitVabTableWrap(w));
    }, 180);
  });
}
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
    ? '<tr>' + headers.map(h=>`<th>${esc(h)}</th>`).join('') + '<th style="width:82px;text-align:center;">操作</th></tr>'
    : '';
  /* ---- 分页：默认 10 条/页 ---- */
  const total = rows.length;
  const pages = Math.max(1, Math.ceil(total / vabPageSize));
  if(vabPage > pages) vabPage = pages;
  if(vabPage < 1) vabPage = 1;
  const vabStart = (vabPage - 1) * vabPageSize;
  const pageRows = rows.slice(vabStart, vabStart + vabPageSize);
  const bodyHtml = pageRows.length
    ? pageRows.map((r,k)=>{
        const i = vabStart + k; // 真实行索引（编辑/删除定位 state.vabRows 用）
        const m = normalizeVab(r);
        const cells = headers.map(h=>`<td class="vab-cell" data-ri="${i}" data-field="${esc(h)}" title="双击编辑该字段">${esc(r[h] ?? '')}</td>`).join('');
        return `<tr data-m="${esc(m)}">${cells}<td style="text-align:center;white-space:nowrap;"><button class="vab-edit" type="button" data-ri="${i}" title="编辑该会员（弹窗修改全部字段）">编辑</button><button class="vab-del" type="button" title="删除该会员">×</button></td></tr>`;
      }).join('')
    : '';
  // 分页后每页行数有限：表格完整展开（显式 none 覆盖页面 CSS .table-wrap{max-height:300px}），
  // 横向滚动截在容器内（配合 .vab-list-details/.vab-list-wrap 的 min-width:0 防撑开）
  const tableBlock = rows.length
    ? `<div class="table-wrap vab-table-wrap" style="max-height:none;overflow:auto;border:1px solid var(--c-border);border-radius:8px;margin:8px 0;">
         <table class="data"><thead>${headHtml}</thead><tbody>${bodyHtml}</tbody></table></div>`
    : `<div class="file-meta" style="margin:8px 0;">（暂无 VAB 会员，点「＋ 新增会员」弹窗录入，或重新上传会员表）</div>`;
  const pagerHtml = rows.length
    ? `<div class="detail-pager" style="margin:6px 2px 12px;">
         <button type="button" class="pg-btn" data-vpg="first" title="第一页"${vabPage<=1?' disabled':''}>«</button>
         <button type="button" class="pg-btn" data-vpg="prev" title="上一页"${vabPage<=1?' disabled':''}>‹ 上一页</button>
         <span class="pg-info" style="margin:0 8px;">第 ${vabPage} / ${pages} 页 · 共 ${total} 人</span>
         <button type="button" class="pg-btn" data-vpg="next" title="下一页"${vabPage>=pages?' disabled':''}>下一页 ›</button>
         <button type="button" class="pg-btn" data-vpg="last" title="最后一页"${vabPage>=pages?' disabled':''}>»</button>
         <select id="vab-page-size" class="detail-select" style="height:26px;width:auto;margin-left:8px;" title="每页显示条数">
           ${[10,20,50,100].map(n=>`<option value="${n}"${vabPageSize===n?' selected':''}>${n} 条/页</option>`).join('')}
         </select>
       </div>`
    : '';
  box.innerHTML =
    tableBlock +
    pagerHtml +
    `<div class="vab-add-row">` +
      `<button id="vab-add-btn" class="btn btn-primary" type="button">＋ 新增会员</button>` +
      `<span class="file-meta">弹窗录入会员号/姓名/组别/手机号等全部字段</span>` +
    `</div>`;
  // 渲染后立即对齐滚动高度（消除半行裁切）；清单折叠态测不到行高则跳过，展开 toggle 时会重算
  const vwrap = box.querySelector('.vab-table-wrap');
  if(vwrap) fitVabTableWrap(vwrap);
  const vdetails = box.closest('.vab-list-details');
  if(vdetails && !vdetails.dataset.fitBound){
    vdetails.dataset.fitBound = '1';
    vdetails.addEventListener('toggle', () => {
      const w = box.querySelector('.vab-table-wrap');
      if(w) fitVabTableWrap(w);
    });
  }
  // 事件委托只需绑定一次（容器常驻，内部元素随 innerHTML 重建）
  if(!box.dataset.bound){
    box.addEventListener('click', onVabListClick);
    box.addEventListener('dblclick', onVabListDblClick);
    box.addEventListener('change', onVabListChange);
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
// 列表点击：编辑按钮 / 删除按钮 / 新增按钮（事件委托到 #vab-list 容器）
function onVabListClick(e){
  const edt = e.target.closest && e.target.closest('.vab-edit');
  if(edt){ openEditVabModal(Number(edt.dataset.ri)); return; }
  const del = e.target.closest && e.target.closest('.vab-del');
  if(del){ const tr = del.closest('tr'); deleteVabMember(tr && tr.dataset.m); return; }
  if(e.target.id === 'vab-add-btn'){ openAddVabModal(); return; }
  const vpg = e.target.closest && e.target.closest('[data-vpg]');
  if(vpg){
    const pages = Math.max(1, Math.ceil((state.vabRows || []).length / vabPageSize));
    const a = vpg.dataset.vpg;
    if(a === 'first') vabPage = 1;
    else if(a === 'prev') vabPage = Math.max(1, vabPage - 1);
    else if(a === 'next') vabPage = Math.min(pages, vabPage + 1);
    else if(a === 'last') vabPage = pages;
    renderVabList();
  }
}
// 清单分页：每页条数切换
function onVabListChange(e){
  if(e.target && e.target.id === 'vab-page-size'){
    vabPageSize = Number(e.target.value) || 10;
    vabPage = 1;
    renderVabList();
  }
}
// 新增会员改为弹窗（见 openAddVabModal / saveAddVabModal），不再内联输入会员号
// 新增会员弹窗：按当前 VAB 字段模板生成输入项，可录入会员号/姓名/组别等全部字段
function buildVabHeaders(){
  const set = [];
  (state.vabRows || []).forEach(r => Object.keys(r).forEach(k => { if(!set.includes(k)) set.push(k); }));
  if(!set.includes('会员号')) set.unshift('会员号'); // 会员号始终为首项
  return set;
}
/* ---- 弹窗模式：vabEditIndex=null 为新增，数字为编辑 state.vabRows[该索引]（新增/编辑共用同一弹窗） ---- */
let vabEditIndex = null;
function setVabModalMode(editIdx){
  vabEditIndex = (typeof editIdx === 'number' && editIdx >= 0) ? editIdx : null;
  const title = document.querySelector('#vab-add-modal .modal-title');
  if(title) title.textContent = vabEditIndex === null ? '＋ 新增 VAB 会员' : '✎ 编辑 VAB 会员';
  const save = $('vab-add-save');
  if(save) save.textContent = vabEditIndex === null ? '💾 保存' : '💾 保存修改';
}
// 新增会员弹窗：按当前 VAB 字段模板生成输入项，可录入会员号/姓名/组别等全部字段
function openAddVabModal(){
  setVabModalMode(null);
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
// 编辑会员：复用新增弹窗，预填该行全部字段；保存仅更新该行，停留当前分页
function openEditVabModal(ri){
  const row = state.vabRows && state.vabRows[ri];
  if(!row){ showToast('未找到该会员，请刷新清单后重试', 'warn'); return; }
  setVabModalMode(ri);
  const fields = buildVabHeaders();
  const box = $('vab-add-fields');
  if(box) box.innerHTML = fields.map(f =>
    `<div class="modal-field"><label>${esc(f)}</label><input class="modal-input" data-field="${esc(f)}" type="text" value="${esc(row[f] ?? '')}" placeholder="${esc(f)}" /></div>`
  ).join('');
  const m = $('vab-add-modal');
  if(m) m.style.display = 'flex';
  const first = box && box.querySelector('input');
  if(first){ first.focus(); first.select(); }
}
function closeAddVabModal(){
  setVabModalMode(null); // 复位为新增模式，避免下次打开残留编辑态
  const m = $('vab-add-modal');
  if(m) m.style.display = 'none';
}
// 保存弹窗新增的会员（按字段构建整行对象，去重、持久化、联动统计）；编辑模式转 saveEditVabModal
function saveAddVabModal(){
  if(vabEditIndex !== null){ saveEditVabModal(); return; }
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
  vabPage = Math.max(1, Math.ceil((state.vabRows || []).length / vabPageSize)); // 新增后跳末页，看到新会员
  refreshVabStatus(); // 刷新状态条与清单（含新行）
  closeAddVabModal();
  showToast('已新增会员：'+row['会员号'], 'ok');
}
// 保存弹窗编辑：仅更新当前行（清空字段=删除该字段；会员号查重排除自身；停留当前分页）
function saveEditVabModal(){
  const box = $('vab-add-fields');
  if(!box) return;
  const ri = vabEditIndex;
  const row = state.vabRows && state.vabRows[ri];
  if(!row){ closeAddVabModal(); return; }
  let newNo = '';
  const vals = [];
  box.querySelectorAll('input').forEach(inp => {
    const f = inp.dataset.field;
    const v = (inp.value || '').trim();
    vals.push([f, v]);
    if(f === '会员号') newNo = v;
  });
  if(!newNo){ showToast('会员号不能为空', 'warn'); return; }
  const dup = state.vabRows.some((r, idx) => idx !== ri && String(r['会员号'] || '').toLowerCase() === newNo.toLowerCase());
  if(dup){ showToast('会员号已存在：' + newNo, 'warn'); return; }
  // 校验通过后再写回（失败时不留脏值）
  vals.forEach(([f, v]) => { if(v === '') delete row[f]; else row[f] = v; });
  setVabRows(state.vabRows); // 重建 vabList/vabSet（含新会员号）
  persistVab();
  recompute();
  refreshVabStatus(); // 停留当前分页刷新清单
  closeAddVabModal();
  showToast('已保存会员修改：' + newNo, 'ok');
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
  if(!window.confirm('确定清除本机保存的全部页面数据？\n（订单、月目标、手动调整、列宽将全部清空，刷新后恢复为空；VAB会员与近一年消费明细不受影响。此操作不可撤销）')) return;
  try{
    localStorage.removeItem(ORDERS_STORAGE_KEY);
    localStorage.removeItem(MANUAL_STORAGE_KEY);
    localStorage.removeItem(TARGETS_STORAGE_KEY);
    localStorage.removeItem(COLW_STORAGE_KEY);
  }catch(e){}
  // 同步清除 IndexedDB 持久化层，避免刷新后从 IDB 重新恢复已被清除的订单
  idbDel(ORDERS_STORAGE_KEY).catch(()=>{});
  clearOrdersInCloud(); // 云端一并清空（异步，不阻塞本机清除流程）
  // 本机变更时刻推到最新（不清零）：确保「本机主动清空」不会被云端旧快照在刷新后回灌（ts 裁决 → 保留本地空态）
  bumpSnapTs();
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


// 清除指定日期的业绩记录（按「收款时间」所在日整日清除；数据变化即清空手动调整并全量重算，与删除单条逻辑一致）
function clearOrdersByDate(){
  if(!state.orders.length){ showToast('当前没有任何业绩数据，无需清除', 'warn'); return; }
  // 数据中实际存在的日期集合（降序），供用户核对输入
  const dateSet = new Set();
  state.orders.forEach(r => { if(r.payTime) dateSet.add(fmtDate(r.payTime)); });
  const dates = Array.from(dateSet).sort().reverse();
  const def = state.cutoffDate || dates[0] || '';
  const input = window.prompt('请输入要清除业绩记录的日期（YYYY-MM-DD）：\n当前数据包含的日期：' + dates.slice(0, 10).join('、') + (dates.length > 10 ? ' 等 ' + dates.length + ' 天' : ''), def);
  if(input === null) return; // 用户取消
  const dateStr = String(input).trim();
  if(!/^\d{4}-\d{2}-\d{2}$/.test(dateStr)){ showToast('日期格式应为 YYYY-MM-DD，未执行清除', 'warn'); return; }
  if(!dateSet.has(dateStr)){ showToast('数据中不存在 ' + dateStr + ' 的业绩记录，未执行清除', 'warn'); return; }
  const n = state.orders.filter(r => r.payTime && fmtDate(r.payTime) === dateStr).length;
  if(!window.confirm('确定清除 ' + dateStr + ' 的业绩记录？\n共 ' + n + ' 条订单将被删除，该日在各统计表（月累计/每日列/门店业绩）中的数据将一并消失；\n手动调整将清空恢复自动计算。此操作不可撤销。')) return;
  state.orders = state.orders.filter(r => !(r.payTime && fmtDate(r.payTime) === dateStr));
  // 数据快照变化：清空所有手动调整（月累计/每日），避免残留单元格覆盖值导致"月累计刷新、当日不刷新"等不一致
  if(Object.keys(state.manual.month).length || Object.keys(state.manual.daily).length){
    state.manual = { month:{}, daily:{} };
    persistManual();
  }
  if($('orders-meta')){
    $('orders-meta').textContent = '\u2714 已清除 ' + dateStr + ' 的 ' + n + ' 条记录，剩余 ' + state.orders.length + ' 行';
    $('orders-meta').className = 'file-meta';
  }
  if(state.orders.length === 0){
    // 删空：彻底清空持久化层（localStorage + IndexedDB + 云端），重置为待导入状态
    state.results = null;
    try{ localStorage.removeItem(ORDERS_STORAGE_KEY); }catch(e){}
    idbDel(ORDERS_STORAGE_KEY).catch(()=>{});
    clearOrdersInCloud();
    bumpSnapTs(); // 同 clearAllData：标记本机已有更新动作，防云端旧快照回灌
    renderReport(); renderStoreReport([], []); renderImportDetails(); renderTargets();
    if($('data-status')){ $('data-status').textContent = '待导入数据'; $('data-status').className = 'chip'; }
    showToast('已清除 ' + dateStr + ' 全部记录（页面已无剩余数据），手动调整已清空', 'ok', 3600);
    return;
  }
  // 若被清日期晚于剩余数据最大日，截止日期回退到剩余最大日，保证「当日列」仍有意义
  const maxDate = state.orders.reduce((a,b)=>{ if(!b.payTime) return a; const s=fmtDate(b.payTime); return (!a || s>a) ? s : a; }, null);
  if(state.cutoffDate && maxDate && state.cutoffDate > maxDate){
    state.cutoffDate = maxDate;
    if($('cutoff-date')) $('cutoff-date').value = maxDate;
  }
  recompute(); // 重算并渲染 ③主表 / ④门店表 / ⑥导入明细
  persistOrders(); // 持久化剩余数据（本地 + 云端快照），刷新后保留
  showToast('已清除 ' + dateStr + ' 的 ' + n + ' 条记录并清空手动调整，剩余 ' + state.orders.length + ' 条，各统计表已重新计算', 'ok', 3600);
}

// 工具
const $ = (id) => document.getElementById(id);
const fmt = (n, d=2) => (n==null||isNaN(n)) ? '0' : Number(n).toLocaleString('zh-CN',{minimumFractionDigits:d,maximumFractionDigits:d});
const fmtInt = (n) => (n==null||isNaN(n)) ? '0' : Math.round(Number(n)).toLocaleString('zh-CN');
// 复制专用：无千分位、无强制小数位，便于 Excel 粘贴为数值
const fmtPlain = (n) => (n==null||isNaN(n)) ? '0' : String(Math.round(Number(n)*100)/100);
// HTML 转义：订单字段(name/memberNo 等)经 innerHTML 渲染前必须转义，防止自注入
const esc = (s) => String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
// 本地时区的「今天」(YYYY-MM-DD)。
// 注意：不可用 toISOString()（那是 UTC），否则东八区凌晨 0~8 点会退回前一天，
// 导致 initApp 把截止日设成昨天、报表标题与恢复的月份错位。
const todayIso = () => {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
};
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
  // 字符串字段统一 trim：Excel 单元格常带首尾空格，不清洗会导致
  // 「微信小程序」/「线上400」等精确/包含匹配失效，金额被误落入"其他"
  const s = (v) => String(v == null ? '' : v).trim();
  return {
    name: s(pickField(r, ['客户姓名','客户'])),
    // 会员号列名兼容多种写法：会员号 / 会员编号 / 会员ID / 会员卡号 / 卡号 / VAB会员号
    // 若真实订单表用"会员编号"等命名而只匹配"会员号"，memberNo 会为空 → 全部 VAB 匹配失败 → 新增业绩恒为 0
    memberNo: s(pickField(r, ['会员号','会员编号','会员ID','会员卡号','卡号','VAB会员号','member_no'])),
    developer: s(pickField(r, ['客户开发人'])),
    consultant: s(pickField(r, ['开单健康管家','开单咨询师','咨询师','服务顾问','开单人','主诊咨询师'])),
    channelClass: s(pickField(r, ['客户来源渠道分类','渠道分类'])),
    channelL1: s(pickField(r, ['客户来源一级渠道','一级渠道'])),
    channelL2: s(pickField(r, ['客户来源二级渠道','二级渠道'])),
    channelL3: s(pickField(r, ['客户来源三级渠道','三级渠道','渠道L3','客户来源三级','三级来源'])),
    cashPay: Number(pickField(r, ['现款支付'])) || 0,
    assetDeduct: Number(pickField(r, ['收款金额-资产抵扣','资产抵扣'])) || 0,
    visitType: s(pickField(r, ['到访类型'])),
    payTime: parseDateLoose(pickField(r, ['收款时间'])),
    receiptNo: s(pickField(r, ['收款单号'])),
    cashier: s(pickField(r, ['收款员工'])),
    payType: s(pickField(r, ['收款类型'])),
    payMethod: s(pickField(r, ['收款详情汇总','收款方式','详情'])),
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

// 「近一年消费明细」键集合（2026-10-16 新增）：历史明细中 payTime 在 [截止日-365天, 截止日]
// 窗口内的客户键（会员号优先，无会员号按姓名）。规则B数据源：营销部渠道顾客若不在此集合中
// → 一年内无消费记录 → 视为初诊客户。带缓存（按 截止日 + 历史行数 + 时间戳 失效）。
let _histKeysCache = { sig:'', date:'', set:null };
function histKeysCached(cutoff){
  const d = (cutoff && !isNaN(cutoff.getTime())) ? fmtDate(cutoff) : '';
  const sig = (state.histOrders||[]).length + '|' + (state.histTs||0);
  if(_histKeysCache.sig !== sig || _histKeysCache.date !== d){
    const s = new Set();
    if(d && (state.histOrders||[]).length){
      // 近一年窗口：cutoff 往前推 366 天（含容差，避免闰年/导出口径漂移把临界消费挡在窗外）
      const start = fmtDate(new Date(cutoff.getTime() - 366*86400000));
      state.histOrders.forEach(r => {
        if(!r.payTime) return;
        const dd = fmtDate(r.payTime);
        if(dd < start || dd > d) return;
        const k = mktFirstClientKey(r);
        if(k && k !== 'name:') s.add(k);
      });
    }
    _histKeysCache = { sig, date:d, set:s };
  }
  return _histKeysCache.set;
}

// 「营销部初诊客户」键集合（会员号优先，无会员号按姓名）——营销部各类别业绩仅计入这些客户的消费。
// 规则A（既有口径）：滚动 30 天（含两端，可跨月）窗口内 来源渠道=营销部 且 到访类型=初诊；
//   窗口起点=截止日往前推 29 天：9.30 → 9.1~9.30，10.1 → 9.2~10.1（2026-10-15 需求）。
// 规则B（2026-10-16 新增）：营销部渠道顾客（渠道分类=营销部 或 推荐→营销部）在「近一年消费明细」
//   中无任何消费记录（截止日往前推一年窗口内无该客户键）→ 视为初诊客户。
//   仅导入「近一年消费明细」后启用（集合非空才算启用），未导入时保持既有口径，防止复诊老客误判。
// 两条规则取并集；客户键口径一致（mktFirstClientKey）。
function buildMktFirstKeys(list, cutoff){
  const s = new Set();
  if(!cutoff || isNaN(cutoff.getTime())) return s;
  const end = fmtDate(cutoff);
  const start = fmtDate(new Date(cutoff.getTime() - 29*86400000));
  const histKeys = histKeysCached(cutoff); // 未导入历史明细时为空集 → 规则B自动停用
  (list||[]).forEach(r => {
    if(!r.payTime) return;
    const d = fmtDate(r.payTime);
    if(d >= start && d <= end && isMktChuzhenOrder(r)){
      // 规则A：30天窗口内 营销部渠道 且 到访类型=初诊
      const k = mktFirstClientKey(r);
      if(k && k !== 'name:') s.add(k);
    }else if(histKeys.size && (chanHas(r,'营销部') || isRecToYingxiao(r))){
      // 规则B：营销部渠道顾客 且 近一年历史消费明细中无该客户键 → 新客视为初诊
      const k = mktFirstClientKey(r);
      if(k && k !== 'name:' && !histKeys.has(k)) s.add(k);
    }
  });
  return s;
}

// 4. 业绩计算引擎 - 实现全部15条规则
function calculate(orders, vabSet, cutoff, mktFirstKeysPre){
  // 「营销部初诊客户」集合：必须在月度过滤之前用全量订单构建（初诊可落在上月、仍在30天窗口内）
  const mktFirstKeys = (mktFirstKeysPre instanceof Set) ? mktFirstKeysPre : buildMktFirstKeys(orders, cutoff);
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
  // 渠道字段统一 trim 后比较：容忍历史持久化数据 / Excel 单元格的首尾空格，
  // 避免「微信小程序 」「品牌组 」等值匹配失效而误落入"其他"
  const CC  = r => (r.channelClass||'').trim();
  const L1  = r => (r.channelL1||'').trim();
  const L2  = r => (r.channelL2||'').trim();

  // --- 0. 部门类别归属覆盖（⑥明细「报表归属」手动指定优先）---
  // 在明细表把某条明细的报表归属改为指定部门后，该行不再走渠道自动规则，直接计入指定部门；
  // 「其他」覆盖 = 从所有具体部门规则中剥离，落入残差「其他」；未指定（自动）的行走原有规则。
  const DEPT_KEYS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','短视频组','品牌组-自然到店','商务部','其他'];
  const hasOv = r => DEPT_KEYS.includes(String(r.deptOverride||'').trim());
  const auto  = orders.filter(r => !hasOv(r));
  const ov    = k => orders.filter(r => String(r.deptOverride||'').trim() === k);

  // 营销部各类别口径（2026-10-12 需求）：仅「30 天内营销部初诊客户」的消费计入
  // 品牌组-其他/活动组-活动支持/电商组/短视频组/品牌组-自然到店（含推荐→营销部落位）；
  // 非初诊客户的营销部订单落入残差「其他」；商务部口径不变；手动归属覆盖优先于本门槛。
  // 注意：集合已在函数开头用全量订单按 30 天窗口构建（初诊可跨月），此处直接使用。
  const isMktFirst = r => mktFirstKeys.has(mktFirstClientKey(r));
  const autoYx = auto.filter(isMktFirst);

  // --- 1. 品牌组-其他 ---
  // （括号明确分组：两个分支均要求 营销部+品牌组，避免 && / || 优先级歧义）
  const r1 = [
    ...autoYx.filter(r =>
      (CC(r) === '营销部' &&
       L1(r) === '品牌组' &&
       ['微信小程序','微信公众号','客服手机','前台座机'].includes(L2(r))) ||
      (CC(r) === '营销部' && L1(r) === '品牌组' && L2(r).includes('阿里') && L2(r).includes('高德'))
    ),
    ...ov('品牌组-其他'),
  ];

  // --- 2. 活动组-活动支持 ---
  const r2 = [
    ...autoYx.filter(r => CC(r)==='营销部' && L1(r)==='活动组' && L2(r)==='活动支持'),
    ...ov('活动组-活动支持'),
  ];

  // --- 3. 电商组（接受 一级渠道=电商组 / 客户转介绍 / 企业团检） ---
  const r3 = [
    ...autoYx.filter(r => CC(r)==='营销部' && ['电商组','客户转介绍','企业团检'].includes(L1(r))),
    ...ov('电商组'),
  ];

  // --- 3c. 电商组补充（2026-09-29 需求）：营销部 + 品牌组 + 二级渠道=线上400 → 归入电商组 ---
  const r3c = autoYx.filter(r => CC(r)==='营销部' && L1(r)==='品牌组' && L2(r).includes('线上400'));

  // --- 3d. 新媒体组（2026-10-16 需求）：渠道分类=营销部 + 一级分类=新媒体组 → 独立部门类别 ---
  //      口径与营销部其他类别完全一致：仅「30天内营销部初诊客户」的消费计入（autoYx 已含该门槛）；
  //      ⑥明细的「报表归属」手动指定 新媒体组 时同样优先（ov）。
  const r3d = [
    ...autoYx.filter(r => CC(r)==='营销部' && L1(r)==='新媒体组'),
    ...ov('新媒体组'),
  ];

  // --- 4. 短视频组 ---
  const r4 = [
    ...autoYx.filter(r => CC(r)==='营销部' && L1(r)==='短视频组'),
    ...ov('短视频组'),
  ];

  // --- 5. 品牌组-自然到店（包含自然到店关键字） ---
  const r5 = [
    ...autoYx.filter(r => CC(r)==='营销部' && L1(r)==='品牌组' && L2(r).includes('自然到店')),
    ...ov('品牌组-自然到店'),
  ];

  // --- 6. 商务部（口径：客户来源渠道分类/一级渠道/二级渠道 任一为「商务部」或「市场部」；
  //      市场部数据并入商务部统一核算，故不再单独落入"其他"；不适用初诊客户门槛） ---
  //      注：渠道分类=推荐 的订单改由下方 r6b 单独处理，避免与 chanHas 商务部/市场部 重复计入，
  //      故此处先把推荐类排除，r6b 再并入商务部，二者互不重叠。
  //      ⚠ 手动「报表归属」=商务部 的行必须并入 r6（ov）：
  //      覆盖行已从 auto 中剔除，若 r6 不拼 ov('商务部')，该行金额既不进商务部也不进 r6b，
  //      会被残差「其他」吞掉 → 表现为「改了报表归属，业绩报表没同步」。
  const r6 = [
    ...auto.filter(r => (chanHas(r, '商务部') || chanHas(r, '市场部')) && (r.channelClass||'').trim() !== '推荐'),
    ...ov('商务部'),
  ];

  // --- 6b. 推荐渠道补充（新增逻辑，不动 r1~r6 原有规则）---
  // 渠道分类=推荐 且 一级分类包含「营销部」→ 计入营销部业绩（同样受初诊客户门槛约束）：
  //   二级含「自然到店」→ 品牌组-自然到店(r5b)，否则 → 电商组(r3b)
  // 渠道分类=推荐 且 一级分类包含「商务部」→ 计入商务部业绩(r6b)
  const r3b = autoYx.filter(r => isRecToYingxiao(r) && !(r.channelL2||'').includes('自然到店'));
  const r5b = autoYx.filter(r => isRecToYingxiao(r) && (r.channelL2||'').includes('自然到店'));
  const r6b = auto.filter(r => isRecToShangwu(r));

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

  // 门店归属判定（高新店 / 经开店）：高新店=收款员工为「南媛」或「系统管理员」；经开店=其余
  // --- 14. 北郊（经开店）：收款员工 非 南媛 且 非 系统管理员 ---
  const r14 = orders.filter(r => !isGaoXinCashier(r));
  // --- 15. 高新 = 收款员工为 南媛 或 系统管理员（系统管理员数据计入高新店） ---
  const rGaoXin = orders.filter(r => isGaoXinCashier(r));

  // --- 15. 高新 = 合计 - 北郊 ---
  const sumBeiJiao = sum(r14);
  const sumGaoXin = sumAll - sumBeiJiao;
  const sumStoreTotal = sumGaoXin + sumBeiJiao; // 门店合计 = 高新店 + 经开店

  // --- 8. 其他 = 合计 - (品牌组-其他 + 活动组-活动支持 + 电商组 + 新媒体组 + 品牌组-自然到店 + 商务部) ---
  // 注：市场部已并入"商务部"统一核算，故"其他"的扣减项里的「商务部」已含市场部金额，不会重复计入。
  // 注：原"活动组-自拓活动"板块已整体下线（不展示、不计算），其金额并入"其他"
  // 注：短视频组行已隐藏（2026-10-19），其金额并入"其他"，合计守恒；rows/values 字典仍保留该键供内部计算
  const sumBrandOther = sum(r1);
  const sumActSupport = sum(r2);
  const sumEcom = sum(r3) + sum(r3b) + sum(r3c);
  const sumNewMedia = sum(r3d);
  const sumShort = sum(r4);
  const sumNature = sum(r5) + sum(r5b);
  const sumCommerce = sum(r6) + sum(r6b);
  const sumOther = sumAll - (sumBrandOther + sumActSupport + sumEcom + sumNewMedia + sumNature + sumCommerce);

  return {
    rows: {
      '品牌组-其他': r1, '活动组-活动支持': r2,
      '电商组': [...r3, ...r3b, ...r3c], '新媒体组': r3d, '短视频组': r4,
      '品牌组-自然到店': [...r5, ...r5b],
      '商务部': [...r6, ...r6b], '其他': orders.filter(()=>true),
      '合计': rAll, '新增业绩': r9, '基础业绩': orders, '管家部': r11,
      '北郊': r14, '高新': rGaoXin,
    },
    values: {
      '品牌组-其他': sumBrandOther,
      '活动组-活动支持': sumActSupport,
      '电商组': sumEcom,
      '新媒体组': sumNewMedia,
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
  // 营销部（rowspan 5 = 公司支持2 + 部门业绩3；短视频组行已隐藏 2026-10-19，其金额并入"其他"）
  // 月目标：公司支持块2行共享360；部门业绩块下 电商组+新媒体组 合并为一格(77, rowspanTarget:2)，品牌组-自然到店 单独一格(3)
  {main:'营销部', sub:'公司支持', item:'品牌组-其他', targetKey:'品牌组-其他', rowspanMain:5, rowspanSub:2, rowspanTarget:2, indent:1},
  {main:'',       sub:'',        item:'活动组-活动支持', targetKey:'活动组-活动支持', indent:1},
  {main:'',       sub:'部门业绩', item:'电商组', targetKey:'电商组及短视频组', valueKey:'电商组', targetMergeKey:'电商组及短视频组', rowspanSub:3, rowspanTarget:2, indent:1},
  {main:'',       sub:'',        item:'新媒体组', targetKey:'电商组及短视频组', valueKey:'新媒体组', indent:1},
  {main:'',       sub:'',        item:'品牌组-自然到店', targetKey:'品牌组-自然到店', valueKey:'品牌组-自然到店', indent:1},
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
    let subtotalHtml = '';
    if(g.subtotal){
      const sum = items.reduce((s,k)=>s+(Number(state.targets[k])||0),0);
      subtotalHtml = `<div class="target-subtotal" data-keys='${JSON.stringify(items)}'><span class="t-label">板块小计</span><span class="t-val">${sum}</span><span class="t-unit">万元</span></div>`;
    }
    return `<div class="target-group"><div class="target-group__title">${g.title}</div><div class="target-group__grid">${cells}</div>${subtotalHtml}</div>`;
  }).filter(Boolean).join('');
  el.innerHTML = html;
  el.querySelectorAll('.target-input').forEach(inp=>{
    inp.addEventListener('input', e=>{
      state.targets[e.target.dataset.key] = Number(e.target.value) || 0;
      persistTargets();
      renderReport();
      // 实时刷新各板块小计（只读格随输入联动）
      document.querySelectorAll('.target-subtotal').forEach(st=>{
        const keys = JSON.parse(st.dataset.keys);
        st.querySelector('.t-val').textContent = keys.reduce((s,k)=>s+(Number(state.targets[k])||0),0);
      });
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
//     其他     = 合计 - (品牌组-其他+活动组-活动支持+电商组+新媒体组+短视频组+品牌组-自然到店+商务部)
//     高新     = 顶层合计 - 北郊   （主表顶层=合计；门店表顶层=门店合计）
//     门店合计 = 高新 + 北郊
function recomputeDerived(base, manual){
  const m = Object.assign({}, base || {});
  const man = manual || {};
  for(const k in man){ if(man[k] !== undefined && man[k] !== null && man[k] !== '') m[k] = Number(man[k]) || 0; }
  // 合计的组成项：6个部门块 + 其他（残差）。注意：新增业绩、管家部 是另一维度的子集，不计入合计。
  const BLOCKS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','短视频组','品牌组-自然到店','商务部'];
  const g = k => (k in m ? (Number(m[k]) || 0) : 0);
  // 1) 合计：手动覆盖优先；否则当某部门块被手动覆盖时，仅按该块的增量调整原始合计；
  //    否则保持原合计（即 calculate 的真实总和 / 月视图的 Σ每日合计），不做任何重算。
  if(!('合计' in man)){
    let delta = 0;
    for(const k of BLOCKS){ if(k in man) delta += (Number(m[k]) || 0) - (Number(base[k]) || 0); }
    if(delta !== 0) m['合计'] = (Number(base['合计']) || 0) + delta;
  }
  // 2) 其他：手动覆盖优先；否则 = 合计 - 6部门块之和（残差，随合计/部门块变化同步）
  if(!('其他' in man)) m['其他'] = g('合计') - (g('品牌组-其他') + g('活动组-活动支持') + g('电商组') + g('新媒体组') + g('短视频组') + g('品牌组-自然到店') + g('商务部'));
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
function computeMonthValues(dateList, dailyResults, monthKey){
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
  // 应用手动月累计覆盖（覆盖优先于按日求和）——按月份隔离：只应用当前报表月份的覆盖，
  // 旧格式扁平键（无月份）不再生效，避免历史月份的手动值污染本月月累计。
  const mm = (state.manual.month && state.manual.month[monthKey || curMonthKey()]) || {};
  for(const k in mm){ if(mm[k] !== undefined && mm[k] !== '') m[k] = Number(mm[k]) || 0; }
  // 合计 = 6个部门块 + 其他（新增业绩/管家部不计入）；手动"合计"覆盖优先
  const BLOCKS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','短视频组','品牌组-自然到店','商务部'];
  let sumBlocks = 0;
  for(const k of BLOCKS) sumBlocks += (Number(m[k]) || 0);
  const otherVal = (Number(m['其他']) || 0);
  if(mm['合计'] !== undefined && mm['合计'] !== '') m['合计'] = Number(mm['合计']) || 0;
  else m['合计'] = sumBlocks + otherVal;
  return m;
}

// 月累计展示值派生（专用，勿用 recomputeDerived——2026-10-13 BUG 修复）
// computeMonthValues 已含「Σ每日(含当日手动覆盖) + 月累计覆盖优先」；
// 此处仅补算可由月值安全派生的 其他/基础业绩，
// 【不重算】高新/门店合计/北郊：它们的月累计=Σ每日，若按「合计-北郊」公式重算，
// 会把日值手动覆盖（如改高新某日值）后的正确 Σ 覆盖回主表合计值（表现为"手动修改后总业绩不变"）。
function monthDispValues(monthCalc){
  const m = Object.assign({}, monthCalc || {});
  const g = k => (Number(m[k]) || 0);
  m['其他'] = g('合计') - (g('品牌组-其他') + g('活动组-活动支持') + g('电商组') + g('新媒体组') + g('短视频组') + g('品牌组-自然到店') + g('商务部'));
  m['基础业绩'] = g('合计') - g('新增业绩');
  return m;
}

// 无数据的每日日期集合（renderReportInner 内按订单日期填充）：这些列默认隐藏，「+展开」也不展开
let emptyDailyDates = new Set();

// 默认隐藏：无数据日期列一律隐藏，其余仅保留「最新有数据的日期」列（2026-10-15 需求）
function defaultHideOldDates(tableEl, dateList){
  let lastWithData = -1;
  for(let i=dateList.length-1; i>=0; i--){
    if(!emptyDailyDates.has(dateList[i])){ lastWithData = i; break; }
  }
  for(let idx=0; idx<dateList.length; idx++){
    if(idx === lastWithData) continue;   // 最新有数据的列保持可见
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
// 渲染后恢复用户可见性：仅隐藏"之前就隐藏"的日期列；首次渲染(null)则按默认规则。
// 无数据日期列始终强制隐藏（即使此前被展开过）——「+展开」也会跳过这些列。
function applyHiddenState(tableEl, dateList, prevHidden){
  if(prevHidden === null){ defaultHideOldDates(tableEl, dateList); return; }
  const needScan = (prevHidden && prevHidden.size) || emptyDailyDates.size;
  if(needScan){
    // 单次扫描替代「每个隐藏日期各扫一遍全表」，避免 O(隐藏列数 × 单元格数) 的重复查询
    const cells = tableEl.querySelectorAll('td[data-idx],th[data-idx]');
    cells.forEach(el => {
      const d = dateList[+el.dataset.idx];
      if((prevHidden && prevHidden.has(d)) || emptyDailyDates.has(d)) el.classList.add('col-hidden');
    });
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
    // ⑤ 电商组二级渠道业绩占位
    renderEcSubReport();
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
  // 无数据日期列集合：当日无任何订单的日期列默认隐藏，「+展开」也不展开（2026-10-15 需求）
  emptyDailyDates = new Set(dateList.filter(d => !byDate.has(d)));
  const dailyResults = dateList.map(d => {
    const g = byDate.get(d);
    // 必须把「30天内营销部初诊客户」集合用【全量订单】按当日为截止预构建后传入：
    // 否则 calculate 会退化为用「当天订单」构建集合，等于要求初诊订单与消费同日，
    // 导致「上月/前几日初诊、当日消费」的客户业绩被挡在营销部各类别之外（落入「其他」）。
    // 修复案例：宋鑫丹 9-18 初诊、10-9 消费 785（营销部-电商组）应计入 10-9 电商组。
    return (g && g.length) ? calculate(g, state.vabSet, new Date(d), buildMktFirstKeys(state.orders, new Date(d))) : null;
  });
  // 缓存：供"按日期复制整列"复用，避免重复计算
  state.dateList = dateList;
  state.dailyResults = dailyResults;

  // 月累计展示值 = 各日求和(含当日手动覆盖)；手动月累计覆盖优先。
  // 手动改每日→联动月累计；手动改月累计→联动合计/派生行（dateList/dailyResults 在此处已就绪）
  const monthCalc = computeMonthValues(dateList, dailyResults);
  const dispValues = monthDispValues(monthCalc);

  // 健康检查：子集行（新增业绩/管家部）月累计不应超过「合计」；超限通常意味着历史手动覆盖残留
  // （如跨月残留的「新增业绩」手动月累计，使 基础业绩=合计−新增 变为负数）。提示用户清除手动调整。
  // 用 window.__subsetWarn 做会话级去重，避免每次渲染重复弹窗。
  window.__subsetWarn = window.__subsetWarn || {};
  ['新增业绩','管家部'].forEach(k=>{
    const sv = Number(dispValues[k]) || 0, tv = Number(dispValues['合计']) || 0;
    if(tv > 0 && sv > tv){
      if(!window.__subsetWarn[k]){
        window.__subsetWarn[k] = true;
        showToast(`⚠ 「${k}」月累计(${fmt(sv)}) 大于「合计」(${fmt(tv)})，疑似历史手动覆盖残留——请点击「🧹 清除手动调整」或双击该单元格清空后恢复自动计算`, 'warn', 9000);
      }
    } else { window.__subsetWarn[k] = false; }
  });

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
    const targetKey = (row.rowspanTarget > 1) ? (row.targetMergeKey || row.sub || row.targetKey) : row.targetKey;
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

  // ⑤ 电商组各二级渠道业绩（初诊，日/月）
  renderEcSubReport();
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
  // 月累计列（双击 → 保存到 state.manual.month[当前月份]；清空输入 = 删除覆盖恢复自动计算）
  tableEl.querySelectorAll('td.num[data-rowkey]:not([data-idx])').forEach(td=>{
    const key = td.dataset.rowkey;
    const mk = curMonthKey();
    makeEditable(td, v=>{
      if(v === ''){
        if(state.manual.month[mk]) delete state.manual.month[mk][key];
        persistManual();
        renderReport();
        showToast(`「${key}」月累计手动覆盖已清除，恢复自动计算`, 'ok');
        return;
      }
      if(!state.manual.month[mk]) state.manual.month[mk] = {};
      state.manual.month[mk][key] = v;
      persistManual();
      renderReport();
      showToast(`「${key}」月累计已手动调整为 ${fmt(v)}（仅 ${mk} 生效）`, 'ok');
    }, {allowEmpty:true});
  });
  // 每日数据列（双击 → 保存到 state.manual.daily[date][rowKey]；清空输入 = 删除覆盖恢复自动计算）
  tableEl.querySelectorAll('td.num[data-idx][data-rowkey]').forEach(td=>{
    const key = td.dataset.rowkey;
    const dateStr = td.dataset.date;
    makeEditable(td, v=>{
      if(v === ''){
        if(state.manual.daily[dateStr]) delete state.manual.daily[dateStr][key];
        persistManual();
        renderReport();
        showToast(`${dateStr}「${key}」手动覆盖已清除，恢复自动计算`, 'ok');
        return;
      }
      if(!state.manual.daily[dateStr]) state.manual.daily[dateStr] = {};
      state.manual.daily[dateStr][key] = v;
      persistManual();
      renderReport();
      showToast(`${dateStr}「${key}」已手动调整为 ${fmt(v)}`, 'ok');
    }, {allowEmpty:true});
  });
}

// ④ 门店业绩（高新店 / 经开店）：独立表格，显示每天各门店的业绩（不显示月目标）
// 高新店 = 收款员工为「南媛」或「系统管理员」；经开店 = 收款员工既非「南媛」也非「系统管理员」
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
  const dispValues = monthDispValues(monthCalc);
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const monthStr = (cutoff.getMonth()+1) + '月' + cutoff.getDate() + '日';
  const storeTitleCols = 1 + 1 + dateList.length; // 门店类别 + 月累计 + 每日（批量/展开已移至面板标题栏按钮）

  const thead = `
    <tr><th colspan="${storeTitleCols}" class="title-cell">门店业绩完成情况（高新店 / 经开店，截止${monthStr}）</th></tr>
    <tr>
      <th class="sub-head">门店类别</th>
      <th class="sub-head">月累计(单位：元)</th>
      ${dateList.map((d,i)=>{
        const dt = new Date(d);
        return `<th class="sub-head daily-col" data-date="${d}" data-idx="${i}" style="cursor:pointer;position:relative;" title="点击收起该列">${dt.getMonth()+1}月${dt.getDate()}日<span class="copy-col" data-date="${d}" data-idx="${i}" title="复制该日期整列数据">复制</span></th>`;
      }).join('')}
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

  // 列宽：前2列按比例，每日列固定 65px（批量/展开为面板按钮，不再占表尾列）
  storeTable.querySelectorAll('colgroup').forEach(c=>c.remove());
  const colgroup = document.createElement('colgroup');
  ['16%','20%'].forEach(w=>{ const c=document.createElement('col'); c.style.width=w; colgroup.appendChild(c); });
  for(let i=0;i<dateList.length;i++){ const c=document.createElement('col'); c.style.width=DAILY_COL_PX+'px'; colgroup.appendChild(c); }
  storeTable.insertBefore(colgroup, storeTable.firstChild);

  bindStoreColumnToggle(storeTable);
  // 保留用户展开/收起状态：仅重新隐藏"之前就隐藏"的日期列；首次渲染走默认仅显最新
  applyHiddenState(storeTable, dateList, prevHiddenStore);
  bindManualEdit(storeTable);
}

// 门店表每日列：单列点击收起；批量模式多选（与 ③ 业绩报表一致），「+展开」为面板按钮
function updateStoreTitleColspan(tableEl){
  // 标题跨整个表头可见区域：门店类别 + 月累计 + 可见每日列。
  const visibleDaily = tableEl.querySelectorAll('th.daily-col:not(.col-hidden)').length;
  const titleTh = tableEl.querySelector('thead tr:first-child th');
  if(titleTh) titleTh.colSpan = 2 + visibleDaily;
}

function bindStoreColumnToggle(tableEl){
  tableEl.querySelectorAll('th.daily-col').forEach(th=>{
    th.addEventListener('click', ()=>{
      if(storeBatchMode){
        // 批量模式：切换选中（黄色高亮），点面板「✕ 完成隐藏」一次性隐藏
        th.classList.toggle('selected');
        return;
      }
      const idx = th.dataset.idx;
      const hide = !th.classList.contains('col-hidden');
      tableEl.querySelectorAll(`[data-idx="${idx}"]`).forEach(el=>el.classList.toggle('col-hidden', hide));
      updateStoreTitleColspan(tableEl);
    });
  });
  // 复制整列（停止冒泡，避免触发收起/批量选中）
  tableEl.querySelectorAll('th.daily-col .copy-col').forEach(btn=>{
    btn.addEventListener('click', e=>{
      e.stopPropagation();
      copyDailyColumn(btn.dataset.date, +btn.dataset.idx, true);
    });
  });
  // 初始同步一次
  updateStoreTitleColspan(tableEl);
}

// 门店表批量模式（与 ③ 业绩报表「☑ 批量 / +展开」同一交互）：
// 「☑ 批量」→ 黄色多选日期列 → 「✕ 完成隐藏」一次性隐藏；「+展开」恢复全部隐藏列
let storeBatchMode = false;
function bindStoreBatchExpandButtons(tableEl){
  const expandBtn = document.querySelector('#store-expand-all');
  const batchBtn = document.querySelector('#store-batch-toggle');
  const dailyThsNow = ()=> [...tableEl.querySelectorAll('th.daily-col')];
  const exitBatch = ()=>{
    storeBatchMode = false;
    tableEl.classList.remove('batch-mode');
    if(batchBtn) batchBtn.textContent = '☑ 批量';
    dailyThsNow().forEach(th=>th.classList.remove('selected'));
  };
  if(batchBtn && !batchBtn.dataset.bound){
    batchBtn.dataset.bound = '1';
    batchBtn.addEventListener('click', ()=>{
      const dailyThs = dailyThsNow();
      if(storeBatchMode){
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
        updateStoreTitleColspan(tableEl);
      } else {
        // 进入批量模式
        storeBatchMode = true;
        tableEl.classList.add('batch-mode');
        batchBtn.textContent = '✕ 完成隐藏';
      }
    });
  }
  if(expandBtn && !expandBtn.dataset.bound){
    expandBtn.dataset.bound = '1';
    expandBtn.addEventListener('click', ()=>{
      // 「+展开」跳过无数据日期列（这些列保持隐藏，2026-10-15 需求）
      tableEl.querySelectorAll('th.daily-col.col-hidden').forEach(th=>{
        if(emptyDailyDates.has(th.dataset.date)) return;
        tableEl.querySelectorAll(`[data-idx="${th.dataset.idx}"]`).forEach(el=>el.classList.remove('col-hidden'));
      });
      if(storeBatchMode) exitBatch();
      updateStoreTitleColspan(tableEl);
    });
  }
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
      // 「+展开」跳过无数据日期列（这些列保持隐藏，2026-10-15 需求）
      tableEl.querySelectorAll('th.daily-col.col-hidden').forEach(th=>{
        if(emptyDailyDates.has(th.dataset.date)) return;
        tableEl.querySelectorAll(`[data-idx="${th.dataset.idx}"]`).forEach(el=>el.classList.remove('col-hidden'));
      });
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
  return (r.channelClass||'').trim() === val || (r.channelL1||'').trim() === val || (r.channelL2||'').trim() === val;
}

// 推荐渠道归类（新增逻辑，不改原有 r1~r6）：渠道分类=推荐 时按「一级渠道」落地到营销部/商务部
// 一级分类包含「营销部」→ 计入营销部业绩（二级含自然到店 → 品牌组-自然到店，否则 → 电商组）
// 一级分类包含「商务部」→ 计入商务部业绩
function isRecToYingxiao(r){
  return (r.channelClass||'').trim() === '推荐' && (r.channelL1||'').includes('营销部');
}
function isRecToShangwu(r){
  return (r.channelClass||'').trim() === '推荐' && (r.channelL1||'').includes('商务部');
}

// ===== 报表归属（部门类别）：明细表可手动指定，未指定按渠道规则自动归属 =====
// 与 calculate 的 r1~r6b 规则镜像（仅用于展示/排序/筛选，金额计算仍以 calculate 为准）
const DEPT_OPTIONS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','短视频组','品牌组-自然到店','商务部','其他'];
// 营销部板块内的部门类别（手动把某行归属到这些类别时，该行也应在「营销部」页签可见）
const YX_DEPT_KEYS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','短视频组','品牌组-自然到店'];
// ③ 报表中「营销部-部门业绩」实际展示的 4 行（品牌组-自然到店/电商组/新媒体组）+「公司支持」2 行
//   = 品牌组-其他 / 活动组-活动支持 / 电商组 / 新媒体组 / 品牌组-自然到店。
//   注意：「短视频组」行报表已隐藏（2026-10-19）且金额并入残差「其他」，故不属部门业绩口径。
const REPORT_YX_KEYS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','品牌组-自然到店'];
// 该明细是否被手动指定为某部门类别（只看覆盖值，不含自动归属）
function deptOverrideIs(r, dept){ return String(r.deptOverride||'').trim() === dept; }
// 该明细是否被手动指定到营销部板块（渠道可能是公司所有/商务部等，与报表归属无关）
function isYxDeptOverride(r){ return YX_DEPT_KEYS.indexOf(String(r.deptOverride||'').trim()) >= 0; }
// 按【收款日】缓存「30天内营销部初诊客户」集合 —— 与 ③ 报表每日列
//   calculate(当日订单, vabSet, 当日, buildMktFirstKeys(state.orders, 当日)) 完全同源。
// 明细页签用它逐笔判定，避免「用单一窗口过滤整月」把初诊前的消费算进来（2026-10-21 用户反馈）。
let _mktGateByDate = { sig:'', map:new Map() };
function mktGateSetByDate(orders, dateStr){
  const sig = (state.orders||[]).length + '|' + (state.ordersTs||0)
    + '|H' + (state.histOrders||[]).length + '|' + (state.histTs||0);
  if(_mktGateByDate.sig !== sig){ _mktGateByDate.sig = sig; _mktGateByDate.map = new Map(); }
  let s = _mktGateByDate.map.get(dateStr);
  if(!s){ s = buildMktFirstKeys(orders, new Date(dateStr)); _mktGateByDate.map.set(dateStr, s); }
  return s;
}
// 该明细在「③ 报表 → 营销部 → 部门业绩/公司支持」中的归属（镜像 calculate 的 r1/r2/r3/r3b/r3c/r3d/r5/r5b）
// 返回 null = 不计入部门业绩（短视频组已隐藏并入「其他」、营销部渠道但无对应部门类别、
//                 非初诊客户的营销部订单），用于让「营销部·初诊」明细与报表部门业绩合计严格一致。
function yxReportDept(r, gateSet){
  const ov = String(r.deptOverride||'').trim();
  if(REPORT_YX_KEYS.indexOf(ov) >= 0) return ov;   // 手动指定优先（calculate 的 ov 不受初诊门槛约束）
  const cc=(r.channelClass||'').trim(), l1=(r.channelL1||'').trim(), l2=(r.channelL2||'').trim();
  if(!(cc==='营销部' || isRecToYingxiao(r))) return null;
  const gate = gateSet || mktFirstKeysCached();
  const yxGate = gate.has(mktFirstClientKey(r));
  if(cc==='营销部' && l1==='品牌组' && (['微信小程序','微信公众号','客服手机','前台座机'].includes(l2) || (l2.includes('阿里') && l2.includes('高德')))) return yxGate?'品牌组-其他':null;
  if(cc==='营销部' && l1==='活动组' && l2==='活动支持') return yxGate?'活动组-活动支持':null;
  if(cc==='营销部' && ['电商组','客户转介绍','企业团检'].includes(l1)) return yxGate?'电商组':null;
  if(cc==='营销部' && l1==='新媒体组') return yxGate?'新媒体组':null;
  if(cc==='营销部' && l1==='品牌组' && l2.includes('线上400')) return yxGate?'电商组':null;
  if(cc==='营销部' && l1==='品牌组' && l2.includes('自然到店')) return yxGate?'品牌组-自然到店':null;
  if(isRecToYingxiao(r)) return yxGate ? (l2.includes('自然到店') ? '品牌组-自然到店' : '电商组') : null;
  return null;  // 营销部渠道但不属于任何部门类别（如一级=商务合作、二级=外拓活动、短视频组）→ 报表「其他」
}
// 「30天内营销部初诊客户」键集合（带缓存，按 截止月+订单数+时间戳 失效）——与 calculate 口径一致
let _mktKeysCache = { sig:'', set:null };
function mktFirstKeysCached(){
  const month = state.cutoffDate ? String(state.cutoffDate).slice(0,7) : '';
  // 签名含历史明细签名（行数+时间戳）：导入/清除「近一年消费明细」后初诊键集合立即重建
  const sig = month + '|' + (state.orders||[]).length + '|' + (state.ordersTs||0)
    + '|H' + (state.histOrders||[]).length + '|' + (state.histTs||0);
  if(_mktKeysCache.sig !== sig){
    // cutoff 未设置时默认 today（与 renderMktFirstClients 对齐），避免 gateSet 静默为空导致⑤清单空白
    const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
    _mktKeysCache = { sig, set: buildMktFirstKeys(state.orders, cutoff) };
  }
  return _mktKeysCache.set;
}
function autoDeptOf(r){
  const cc=(r.channelClass||'').trim(), l1=(r.channelL1||'').trim(), l2=(r.channelL2||'').trim();
  // 营销部各类别门槛：仅「30天内营销部初诊客户」的消费计入（与 calculate 一致），否则落「其他」
  const yxGate = mktFirstKeysCached().has(mktFirstClientKey(r));
  if(cc==='营销部' && l1==='品牌组' && (['微信小程序','微信公众号','客服手机','前台座机'].includes(l2) || (l2.includes('阿里') && l2.includes('高德')))) return yxGate ? '品牌组-其他' : '其他';
  if(cc==='营销部' && l1==='活动组' && l2==='活动支持') return yxGate ? '活动组-活动支持' : '其他';
  if(cc==='营销部' && ['电商组','客户转介绍','企业团检'].includes(l1)) return yxGate ? '电商组' : '其他';
  if(cc==='营销部' && l1==='新媒体组') return yxGate ? '新媒体组' : '其他';
  if(cc==='营销部' && l1==='品牌组' && l2.includes('线上400')) return yxGate ? '电商组' : '其他';
  if(cc==='营销部' && l1==='短视频组') return yxGate ? '短视频组' : '其他';
  if(cc==='营销部' && l1==='品牌组' && l2.includes('自然到店')) return yxGate ? '品牌组-自然到店' : '其他';
  if(isRecToYingxiao(r)) return yxGate ? (l2.includes('自然到店') ? '品牌组-自然到店' : '电商组') : '其他';
  if((chanHas(r,'商务部') || chanHas(r,'市场部')) && cc!=='推荐') return '商务部';
  if(isRecToShangwu(r)) return '商务部';
  return '其他';
}
function effectiveDept(r){
  const o = String(r.deptOverride||'').trim();
  return DEPT_OPTIONS.includes(o) ? o : autoDeptOf(r);
}
// 归属下拉单元格（ov=手动覆盖值；自动选项显示当前自动归属结果）
function deptSelectHTML(r){
  const ov = String(r.deptOverride||'').trim();
  const autoDept = autoDeptOf(r);
  const style = ov ? 'color:#1a4f8b;font-weight:600;' : '';
  return `<select class="detail-select dept-sel" data-uid="${esc(r._uid)}" style="${style}"
    title="修改该明细在业绩报表中归属的部门类别（自动=按渠道规则，${esc(autoDept)}）">
    <option value="" ${ov?'':'selected'}>自动·${esc(autoDept)}</option>
    ${DEPT_OPTIONS.map(k=>`<option value="${k}" ${ov===k?'selected':''}>${k}</option>`).join('')}
  </select>`;
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
  {key:'yx_chuzhen',    label:'营销部·初诊'},
  {key:'yx_blank',      label:'营销部·类型空白'},
  {key:'recommend',     label:'推荐'},
  {key:'staff',         label:'员工消费'},
  {key:'steward',       label:'管家部'},
];
let detailTab = 'all';
let detailMonth = []; // 月度筛选：选中月份数组（[]=全部；'__blank__'=收款时间为空白的行）
let detailDate = '__latest__';  // 日期筛选三态：'__latest__'=默认显示当日（最新日期）；''=全部日期；'YYYY-MM-DD'=指定日期（日期控件为点选式非下拉，保持单选）
let detailSort = { key:'', dir:1 }; // 列排序：''=未排序；dir 1=升序 -1=降序（再次点击同列切换）
let detailFilters = {}; // 字段筛选：key -> 选中值数组（[]=全部；'__blank__'=该字段为空白）；name/memberNo 模糊字符串；cashPay/assetDeduct 数值字符串（配 *Op）；vab 为 ['hit']/['miss']
// 多选下拉（msel）运行态：当前展开的 key + 各 key 的候选项缓存（供「全选」用）
const MS_BLANK = '__blank__';
const mselRuntime = { open: null, options: {}, search: null };  // search={key,q}：面板内关键词（重渲染后恢复）

// 数值比较（2026-10-13 需求：笔数/金额筛选改为「运算符+数值」形式，参考 Excel 筛选）
// op ∈ eq(等于)/ne(不等于)/gt(大于)/ge(大于或等于)/lt(小于)/le(小于或等于)，''/缺省 = 等于
function cmpNum(v, op, t){
  switch(op){
    case 'ne': return v !== t;
    case 'gt': return v >  t;
    case 'ge': return v >= t;
    case 'lt': return v <  t;
    case 'le': return v <= t;
    default:   return v === t;
  }
}
const NUM_OPS = [['','等于'],['ne','不等于'],['gt','大于'],['ge','大于或等于'],['lt','小于'],['le','小于或等于']];

// 明细表排序取值：日期按时间戳、金额按数值、VAB 按 0/1，其余按中文字符串比较
function detailSortVal(r, key){
  switch(key){
    case 'payTime': return r.payTime ? r.payTime.getTime() : 0;
    case 'cashPay': return Number(r.cashPay)||0;
    case 'assetDeduct': return Number(r.assetDeduct)||0;
    case 'dept': return effectiveDept(r);
    case 'vab':     return isVabMember(r.memberNo) ? 1 : 0;
    default:        return String(r[key]||'');
  }
}
// 可排序列定义（# 与 操作 列不排序）
// 2026-10-23：新增「客户开发人」列（原表字段：客户所属(实时)-客户开发人），各页签共用此表
const DETAIL_SORT_COLS = [
  ['name','姓名'], ['memberNo','会员号'], ['developer','客户开发人'], ['channelClass','渠道分类'], ['channelL1','一级渠道'],
  ['channelL2','二级渠道'], ['visitType','到访类型'], ['consultant','咨询师'], ['cashier','收款员工'],
  ['payTime','收款时间'], ['cashPay','现款支付'], ['assetDeduct','资产抵扣金额'], ['dept','报表归属'], ['vab','VAB'],
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

// 按标签筛选导入订单。
// 日期口径说明：所有页签不再内置"当日"过滤，统一由下方日期筛选承担 ——
// detailDate 默认 '__latest__'（=数据中最新日期，即"当日"），可通过日期下拉切换到其他日期或"全部日期"。
function filterOrdersByTab(key, opts){
  const orders = state.orders || [];
  // opts.ignoreFilters=true：只按页签自身口径取全量数据（忽略月度/日期/字段/金额/VAB 筛选），
  // 用于「导出全部页签」——每个页签导出其完整数据，不受当前筛选状态影响。
  const ignoreFilters = !!(opts && opts.ignoreFilters);
  const latest = latestOrderDate();
  // 日期筛选的生效日期需【先于】页签筛选算出：营销部·初诊页签的初诊客户集合
  // 以「筛选日期」为截止日构建（与 ③ 报表每日列 calculate(g, cutoff=当日) 完全一致）。
  // '__latest__'=默认当日（最新日期），''=全部日期（此时用报表截止日，对应月累计口径）
  let effDate = detailDate;
  if(effDate === '__latest__') effDate = latest;
  let list;
  if(key === 'all') list = orders.slice();
  else if(key === 'today') list = orders.slice(); // 当日口径由日期筛选（默认当日）承担
  else if(key === 'today_gaoxin')  list = orders.filter(r => isGaoXinCashier(r));
  else if(key === 'today_beijiao') list = orders.filter(r => !isGaoXinCashier(r));
  else if(key === 'shangwu') list = orders.filter(r => chanHas(r,'商务部') || chanHas(r,'市场部') || isRecToShangwu(r) || deptOverrideIs(r,'商务部'));
  else if(key === 'newperf')  list = orders.filter(r => isVabMember(r.memberNo));
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
  else if(key === 'yingxiao') list = orders.filter(r => chanHas(r,'营销部') || isRecToYingxiao(r) || isYxDeptOverride(r));
  // 营销部·初诊：与 ③ 报表每日列口径完全一致——
  // 以「筛选日期」为截止日构建「30天内营销部初诊客户」集合（buildMktFirstKeys 语义：
  // 窗口内有任一营销部初诊的客户即入集合），集合内客户的【全部】营销部订单明细计入
  // （不限到访类型：初诊/复诊/空白均计入，即"30天内有初诊的客户多次到访都核算为初诊业绩"）。
  // 日期筛选=某日 d 时 → 集合截止=d，与报表 d 日列逐笔一致；全部日期 → 集合截止=报表截止日（月累计口径）。
  else if(key === 'yx_chuzhen'){
    // 口径与「③ 报表 → 营销部 → 公司支持/部门业绩」严格一致（2026-10-21 用户两次反馈）：
    //  ① 部门维度：原先只按「营销部渠道 + 初诊客户」筛选，短视频组（报表行已隐藏、金额并入残差「其他」）、
    //     一级渠道=商务合作、二级=外拓活动等行会进明细但不进报表 → 明细合计 > 报表部门业绩合计。
    //     现用 yxReportDept 精确镜像 calculate 的部门归属：null = 不计入部门业绩，不展示。
    //  ② 初诊门槛维度：③ 报表每个每日列用【各自的日期】构建初诊集合
    //     （calculate(当日订单, ..., buildMktFirstKeys(state.orders, 当日))，见 renderReportInner），
    //     而本页签原先用【筛选日/最新日】一次性 gateSet 过滤整月 →
    //     「初诊发生在某次消费之后」的客户，其初诊前的消费会被明细算进来、报表却不算 → 仍不一致。
    //     现改为按每笔订单的【收款日】判定，与报表每日列逐日口径逐笔对齐。
    list = orders.filter(r => r.payTime && yxReportDept(r, mktGateSetByDate(orders, fmtDate(r.payTime))) !== null);
  }
  // 营销部·类型空白：审计视图——营销部渠道下到访类型空白（未填写）的【全部】订单，
  // 不受「30天初诊客户」集合限制（2026-10-13 需求：10月闫语函/冯梓芯/魏周朝等无初诊记录的
  // 空白类型订单需在此页签可见；此类客户在报表中落「其他」，不影响营销部业绩金额）。
  else if(key === 'yx_blank')   list = orders.filter(r => (chanHas(r,'营销部') || isRecToYingxiao(r)) && !(r.visitType||'').trim());
  // 推荐页签：渠道分类=推荐 的订单（用于核对新归类逻辑落位：一级含营销部→营销部，含商务部→商务部，未匹配→其他）
  else if(key === 'recommend') list = orders.filter(r => (r.channelClass||'').trim() === '推荐');
  // 员工消费：渠道分类=公司所有 且 客户渠道标记为"员工本人及家属"。
  // 实测真实导出表里该标记落在【二级渠道】(L2)，旧导出可能在【三级渠道】(L3)，故 L2/L3 任一命中即可；
  // 原代码只查 L3 导致整组 0 行。按全月统计（不限当日），与「管家部」全月审计口径一致。
  else if(key === 'staff')    list = orders.filter(r => chanHas(r,'公司所有') && ((r.channelL2||'').includes('员工本人及家属') || (r.channelL3||'').includes('员工本人及家属')));
  else if(key === 'steward')  list = orders.filter(r => isSteward(r)); // 管家部：开单咨询师口径 + 渠道兜底（全月审计，不限于当日）
  else list = orders.slice();
  // 月度 / 日期 筛选叠加（与标签取交集）。effDate 已在函数开头算出（供 yx_chuzhen 集合门控复用）
  // 月度筛选叠加（与标签取交集）：多选数组命中；'__blank__' 匹配收款时间为空白的行
  if(!ignoreFilters && detailMonth.length) list = list.filter(r => {
    const m = r.payTime ? fmtMonth(r.payTime) : '';
    if(m) return detailMonth.includes(m);
    return detailMonth.includes(MS_BLANK);
  });
  if(effDate && !ignoreFilters) list = list.filter(r => r.payTime && fmtDate(r.payTime) === effDate);
  // 字段筛选叠加（展示的全部字段）：姓名/会员号模糊包含，其余多选精确命中；收款时间用月度/日期承担，VAB 按命中状态
  const f = detailFilters || {};
  if(ignoreFilters){
    // 导出全部页签：跳过下方全部字段/金额/VAB 筛选，仅保留页签口径
    return list;
  }
  if(f.name){ const q=String(f.name).toLowerCase(); list=list.filter(r=>String(r.name||'').toLowerCase().includes(q)); }
  if(f.memberNo){ const q=String(f.memberNo).toLowerCase(); list=list.filter(r=>String(r.memberNo||'').toLowerCase().includes(q)); }
  // 多选命中：空数组=不过滤；'__blank__' 匹配「字段值为空白」的行，其余按 trim 后精确匹配
  const multiHit = (set, raw) => {
    const v = String(raw==null?'':raw).trim();
    if(set.includes(MS_BLANK) && v==='') return true;
    return set.includes(v);
  };
  ['channelClass','channelL1','channelL2','visitType','consultant','developer','cashier'].forEach(k=>{
    if(f[k] && f[k].length) list=list.filter(r=>multiHit(f[k], r[k]));
  });
  if(f.dept && f.dept.length) list=list.filter(r=>multiHit(f.dept, effectiveDept(r)));
  // 金额筛选（运算符+数值）：数值非空才生效；op 存于 <key>Op（''=等于）
  const applyAmt = (field) => {
    const raw = String(f[field]||'').trim();
    if(raw === '' || isNaN(Number(raw))) return;
    const t = Number(raw), op = f[field+'Op'] || 'eq';
    list = list.filter(r => cmpNum(Number(r[field])||0, op, t));
  };
  applyAmt('cashPay');
  applyAmt('assetDeduct');
  if(f.vab && f.vab.length){
    const hit=f.vab.includes('hit'), miss=f.vab.includes('miss');
    if(hit && !miss) list=list.filter(r=>isVabMember(r.memberNo));
    else if(miss && !hit) list=list.filter(r=>!isVabMember(r.memberNo));
  }
  return list;
}

// ===== 「⑥ 导入数据明细」导出（各页签均可导出）=====
// 表头与页面列一致：# / 姓名 / 会员号 / 客户开发人 / 渠道分类 / 一级渠道 / 二级渠道 / 到访类型 /
// 咨询师 / 收款员工 / 收款时间 / 现款支付 / 资产抵扣金额 / 报表归属 / VAB
const DETAIL_EXPORT_HEAD = ['#','姓名','会员号','客户开发人','渠道分类','一级渠道','二级渠道','到访类型','咨询师','收款员工','收款时间','现款支付','资产抵扣金额','报表归属','VAB'];
const DETAIL_EXPORT_COLW = [6,12,16,12,12,14,16,10,12,12,12,14,14,16,8];

function detailExportRow(r, idx){
  return [
    idx+1,
    r.name||'', r.memberNo||'', r.developer||'', r.channelClass||'', r.channelL1||'', r.channelL2||'',
    r.visitType||'', r.consultant||'', r.cashier||'',
    r.payTime?fmtDate(r.payTime):'',
    Number(r.cashPay)||0, Number(r.assetDeduct)||0,
    effectiveDept(r),                 // 报表归属：手动覆盖优先，否则自动归属
    isVabMember(r.memberNo)?'✔':''
  ];
}

// 生成导出用的二维表（末行附合计）
function detailExportRows(list, tabLabel){
  const rows = [DETAIL_EXPORT_HEAD.slice()];
  list.forEach((r,i)=>rows.push(detailExportRow(r,i)));
  const total = list.reduce((a,b)=>a+(Number(b.cashPay)||0),0);
  const totalAsset = list.reduce((a,b)=>a+(Number(b.assetDeduct)||0),0);
  const foot = new Array(DETAIL_EXPORT_HEAD.length).fill('');
  foot[0] = '合计';
  foot[1] = list.length + ' 行';
  foot[11] = total;          // 现款支付（新增「客户开发人」列后下标 +1）
  foot[12] = totalAsset;     // 资产抵扣金额
  rows.push(foot);
  return rows;
}

// 导出当前页签（所见即所得：含当前筛选与排序；带样式单 sheet）
async function exportDetailCurrent(){
  const list = applyDetailSort(filterOrdersByTab(detailTab));
  if(!list.length){ showToast('当前页签没有可导出的数据','warn'); return; }
  const tabLabel = (DETAIL_TABS.find(t=>t.key===detailTab)||{label:'明细'}).label;
  try{
    await ensureXlsxStyled();
    const raw = detailExportRows(list, tabLabel);
    const rows = raw.map((r,ri)=> r.map((v,ci)=>{
      if(ri===0) return cell(v,'s','subhead');
      if(ri===raw.length-1) return cell(v, typeof v==='number'?'n':'s', ci===11||ci===12?'num':'cat');
      if(ci===11||ci===12) return cell(v,'n','num');
      return cell(v,'s', ci===13?'cat':'');
    }));
    const blob = await buildStyledXlsx({ sheetName:'导入数据明细', rows, merges:[], colWidths:DETAIL_EXPORT_COLW });
    downloadBlob(blob, `日业绩明细-${tabLabel}-${state.cutoffDate||todayIso()}.xlsx`);
    showToast(`已导出「${tabLabel}」页签 ${list.length} 行`, 'ok');
  }catch(e){
    console.error(e);
    showToast('导出失败：'+e.message,'err');
  }
}

// 导出全部页签（每个页签一个 sheet；忽略当前筛选，导出各页签完整数据）
async function exportDetailAll(){
  if(!(state.orders||[]).length){ showToast('尚未导入订单数据','err'); return; }
  try{
    await ensureXlsx();
    const wb = XLSX.utils.book_new();
    let totalRows = 0;
    DETAIL_TABS.forEach(t=>{
      const list = applyDetailSort(filterOrdersByTab(t.key, {ignoreFilters:true}));
      totalRows += list.length;
      const aoa = detailExportRows(list, t.label);
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = DETAIL_EXPORT_COLW.map(w=>({wch:w}));
      XLSX.utils.book_append_sheet(wb, ws, t.label.replace(/[\\\/\?\*\[\]:]/g,'_').slice(0,28));
    });
    XLSX.writeFile(wb, `日业绩明细-全部页签-${state.cutoffDate||todayIso()}.xlsx`);
    showToast(`已导出全部 ${DETAIL_TABS.length} 个页签（共 ${totalRows} 行）`, 'ok');
  }catch(e){
    console.error(e);
    showToast('导出失败：'+e.message,'err');
  }
}

// ===== 手动覆盖（双击单元格）与「报表归属」改动的联动 =====
// 报表的月累计/每日列允许双击手动覆盖（state.manual.month / state.manual.daily），
// 手动值优先级高于自动计算。若用户此前做过手动调整，随后在明细表改了「报表归属」，
// 被手动覆盖的单元格仍显示旧手动值 → 表现为「报表没跟着更新」。
// 因此：改动归属/到访类型后，清除受影响部门行（含残差「其他」）的手动覆盖，让报表回到自动口径。
function clearManualForDeptKeys(keys){
  const want = new Set((keys||[]).map(k=>String(k||'')).filter(Boolean));
  if(!want.size) return 0;
  let n = 0;
  const sweep = (bag)=>{
    Object.keys(bag||{}).forEach(k=>{
      const inner = bag[k] || {};
      Object.keys(inner).forEach(kk=>{
        if(want.has(String(kk))){ delete inner[kk]; n++; }
      });
      if(!Object.keys(inner).length) delete bag[k];
    });
  };
  sweep(state.manual.month);
  sweep(state.manual.daily);
  if(n) writeManualLocal();
  return n;
}

// 导入明细分页状态（2026-10-13 需求：所有页签默认每页 10 条）
let detailPage = 1;
let detailPageSize = 10;

// 筛选区折叠状态（⑥明细）：localStorage 持久化，手机端/电脑端一致
let detailFiltersCollapsed = false;
try{ detailFiltersCollapsed = localStorage.getItem('dp_detail_filters_collapsed') === '1'; }catch(e){}

// 是否手机端视图（2026-10-13 需求：手机端一律卡片形式）
// 判定：窄屏（max-width:640px，与 index.html 断点一致）或 触屏设备（pointer:coarse，
// 覆盖平板/大屏手机/横屏等宽屏触屏场景——触屏下表格点击删除/改类型不便，统一用卡片）
function isMobileView(){
  try {
    if(window.matchMedia && window.matchMedia('(max-width: 640px)').matches) return true;
    if(window.matchMedia && window.matchMedia('(pointer: coarse)').matches) return true;
    return false;
  }
  catch(e){ return false; }
}

// 桌面端表格行（含分页起始偏移，# 列显示全局行号）
function detailRowCells(r, i, start){
  const vab = isVabMember(r.memberNo);
  return `<tr data-uid="${esc(r._uid)}">
    <td style="text-align:center;color:var(--c-muted);">${start+i+1}</td>
    <td>${esc(r.name)}</td>
    <td>${esc(r.memberNo)}</td>
    <td>${esc(r.developer)}</td>
    <td>${esc(r.channelClass)}</td>
    <td>${esc(r.channelL1)}</td>
    <td>${esc(r.channelL2)}</td>
    <td><input type="text" class="detail-select vt-input" data-uid="${esc(r._uid)}" list="vt-options" value="${esc(r.visitType||'')}" placeholder="（空白）" title="点击可手动修改到访类型（如：初诊/复诊），修改后各统计表联动重算" style="min-width:0;width:78px;"></td>
    <td>${esc(r.consultant)}</td>
    <td>${esc(r.cashier)}</td>
    <td>${r.payTime?fmtDate(r.payTime):''}</td>
    <td style="text-align:right;">${fmt(r.cashPay)}</td>
    <td style="text-align:right;">${fmt(r.assetDeduct)}</td>
    <td style="white-space:nowrap;">${deptSelectHTML(r)}</td>
    <td style="text-align:center;${vab?'color:#1a73e8;font-weight:600;':''}">${vab?'✔':'—'}</td>
    <td style="text-align:center;"><button class="detail-del" type="button" data-uid="${esc(r._uid)}" title="删除该条数据">×</button></td>
  </tr>`;
}

// 手机端卡片（每字段 标签:值，删除按钮在右上角）
function detailCardHTML(r){
  const vab = isVabMember(r.memberNo);
  const row = (k,v,cls='') => `<span>${k}</span><b class="${cls}">${v}</b>`;
  return `<div class="detail-card" data-uid="${esc(r._uid)}">
    <div class="dc-head">
      <span class="dc-name">${esc(r.name||'—')}</span>
      <span class="dc-vab ${vab?'hit':''}">${vab?'VAB ✔':'VAB —'}</span>
      <button class="detail-del dc-del" type="button" data-uid="${esc(r._uid)}" title="删除该条数据">×</button>
    </div>
    <div class="dc-rows">
      ${row('会员号', esc(r.memberNo))}
      ${row('客户开发人', esc(r.developer))}
      ${row('渠道分类', esc(r.channelClass))}
      ${row('一级渠道', esc(r.channelL1))}
      ${row('二级渠道', esc(r.channelL2))}
      ${row('到访类型', esc(r.visitType))}
      ${row('咨询师', esc(r.consultant))}
      ${row('收款员工', esc(r.cashier))}
      ${row('收款时间', r.payTime?esc(fmtDate(r.payTime)):'')}
      ${row('现款支付', esc(fmt(r.cashPay)), 'amt')}
      ${row('资产抵扣金额', esc(fmt(r.assetDeduct)), 'amt')}
      ${row('报表归属', esc(effectiveDept(r)) + (String(r.deptOverride||'').trim()?'（手动）':''))}
    </div>
    <div class="dc-rows" style="margin-top:6px;">
      <label style="display:inline-flex;align-items:center;gap:6px;">到访类型
        <input type="text" class="detail-select vt-input" data-uid="${esc(r._uid)}" list="vt-options" value="${esc(r.visitType||'')}" placeholder="（空白）" style="min-width:0;width:96px;">
      </label>
      ${deptSelectHTML(r)}
    </div>
  </div>`;
}

// 分页工具条：上一页/下一页 + 窗口化页码 + 每页条数选择
function renderPager(page, pageCount){
  if(pageCount <= 1) return '';
  const mk = (label, pageNo, opts) => {
    opts = opts || {};
    return `<button type="button" class="pg-btn ${opts.active?'active':''}" data-page="${pageNo}" ${opts.disabled?'disabled':''}>${label}</button>`;
  };
  const nums = [...new Set([1, pageCount, page, page-1, page+1].filter(n => n>=1 && n<=pageCount))].sort((a,b)=>a-b);
  const parts = [ mk('上一页', page-1, {disabled: page<=1}) ];
  let prev = 0;
  for(const n of nums){
    if(n - prev > 1) parts.push('<span class="pg-sep">…</span>');
    parts.push(mk(String(n), n, {active: n===page}));
    prev = n;
  }
  parts.push(mk('下一页', page+1, {disabled: page>=pageCount}));
  const sizeOpts = [10,20,50,100].map(s=>`<option value="${s}" ${s===detailPageSize?'selected':''}>每页 ${s}</option>`).join('');
  return `<div class="detail-pager">
    ${parts.join('')}
    <span class="pg-info">第 ${page}/${pageCount} 页</span>
    <label class="detail-filter pager-size">每页
      <select id="detail-pagesize-filter" class="detail-select" style="min-width:auto;">${sizeOpts}</select>
    </label>
  </div>`;
}

// —— 多选下拉组件（msel）：按钮 + 复选面板；支持多选、（空白）筛选、全选/清空 ——
// 展开态存 mselRuntime.open（key），重渲染后由 applyMselOpenState 恢复，勾选后面板不闪没
function mselGet(key){
  if(key === '__month__') return detailMonth;
  if(key === 'mkt:channel') return Array.isArray(mktFilters.channel) ? mktFilters.channel : [];
  const v = detailFilters[key];
  return Array.isArray(v) ? v : [];
}
function mselSet(key, arr){
  if(key === '__month__'){ detailMonth = arr; return; }
  if(key === 'mkt:channel'){ mktFilters.channel = arr; return; }
  detailFilters[key] = arr;
}
function mselLabelOf(key, val){
  const arr = mselRuntime.options[key] || [];
  const it = arr.find(x => (Array.isArray(x)?x[0]:x) === val);
  return it ? (Array.isArray(it)?it[1]:it) : val;
}
function mselBtnText(label, key, set){
  // 标签由 .df-label 单独呈现，按钮内不重复标签，只显示当前取值状态
  if(!set || !set.length) return '全部';
  if(set.length === 1) return set[0] === MS_BLANK ? '（空白）' : mselLabelOf(key, set[0]);
  return '已选' + set.length + '项';
}
function mselHTML(key, label, arr, extraCls){
  mselRuntime.options[key] = arr;
  const set = mselGet(key);
  const items = [];
  if(key !== 'vab' && key !== 'mkt:channel')   // VAB 仅命中/未命中两态；营销渠道候选里「—」即空白行，无额外空白概念
    items.push('<label class="msel-item"><input type="checkbox" value="'+MS_BLANK+'" '+(set.includes(MS_BLANK)?'checked':'')+'><span>（空白）</span></label>');
  arr.forEach(x=>{
    const v = Array.isArray(x)?x[0]:x, t = Array.isArray(x)?x[1]:x;
    items.push('<label class="msel-item"><input type="checkbox" value="'+esc(v)+'" '+(set.includes(v)?'checked':'')+'><span>'+esc(t)+'</span></label>');
  });
  return '<span class="detail-filter'+(extraCls?' '+extraCls:'')+'"><span class="df-label">'+esc(label)+'</span><span class="msel" data-mskey="'+key+'">'
    + '<button type="button" class="msel-btn detail-select" title="'+esc(label)+'（可多选、可搜索）">'+esc(mselBtnText(label,key,set))+'<span class="msel-caret">▾</span></button>'
    + '<span class="msel-panel"><span class="msel-top">'
    + '<input type="text" class="msel-search" placeholder="输入关键词筛选选项">'
    + '<span class="msel-actions">'
    + '<button type="button" class="msel-act" data-msact="all">全选</button>'
    + '<button type="button" class="msel-act" data-msact="none">清空</button>'
    + '</span></span>' + items.join('') + '</span></span></span>';
}
// 面板内关键词过滤：仅控制候选项显隐，不触发重渲染
function applyMselSearchFilter(box, q){
  q = (q || '').toLowerCase();
  box.querySelectorAll('label.msel-item').forEach(it=>{
    it.style.display = (!q || it.textContent.toLowerCase().includes(q)) ? '' : 'none';
  });
}
function applyMselOpenState(){
  document.querySelectorAll('.msel').forEach(b=>{
    const key = b.dataset.mskey;
    const open = key === mselRuntime.open;
    b.classList.toggle('open', open);
    if(open){
      // 恢复面板内搜索词（勾选触发的重渲染后不丢关键词）
      const q = (mselRuntime.search && mselRuntime.search.key === key) ? (mselRuntime.search.q || '') : '';
      const inp = b.querySelector('.msel-search');
      if(inp){ inp.value = q; applyMselSearchFilter(b, q); }
    }
  });
}

// 6c. 导入数据明细（扁平展示全部导入订单 + 标签筛选，便于核对原始数据 / VAB 命中）
function renderImportDetails(){
  const el = $('import-details');
  if(!el) return;
  const orders = state.orders || [];
  if(orders.length === 0){ el.innerHTML = '<div class="file-meta" style="padding:12px 0;">尚未导入订单数据</div>'; return; }
  // 级联收敛（渠道 分类→一级→二级）：上游选择变化后，裁剪下游已选但候选中不再存在的值
  const hitPool = (set, raw) => {
    if(!set || !set.length) return true;
    const v = String(raw==null?'':raw).trim();
    if(set.includes(MS_BLANK)) return v==='' || set.includes(v);
    return set.includes(v);
  };
  const poolC = orders.filter(r => hitPool(detailFilters.channelClass, r.channelClass));
  if(detailFilters.channelL1 && detailFilters.channelL1.length){
    const l1Has = new Set(poolC.map(r=>String(r.channelL1||'').trim()).filter(Boolean));
    detailFilters.channelL1 = detailFilters.channelL1.filter(v => v===MS_BLANK
      ? poolC.some(r=>!String(r.channelL1||'').trim())
      : l1Has.has(v));
  }
  const poolL1 = poolC.filter(r => hitPool(detailFilters.channelL1, r.channelL1));
  if(detailFilters.channelL2 && detailFilters.channelL2.length){
    const l2Has = new Set(poolL1.map(r=>String(r.channelL2||'').trim()).filter(Boolean));
    detailFilters.channelL2 = detailFilters.channelL2.filter(v => v===MS_BLANK
      ? poolL1.some(r=>!String(r.channelL2||'').trim())
      : l2Has.has(v));
  }
  const listRaw = filterOrdersByTab(detailTab);
  const list = applyDetailSort(listRaw);
  const vabCount = list.filter(r=>isVabMember(r.memberNo)).length;
  const total = list.reduce((a,b)=>a+(Number(b.cashPay)||0),0);
  const totalAsset = list.reduce((a,b)=>a+(Number(b.assetDeduct)||0),0);
  // 分页计算（排序后的完整列表分页展示；合计/行数统计仍基于完整 list）
  const pageCount = Math.max(1, Math.ceil(list.length / detailPageSize));
  if(detailPage > pageCount) detailPage = pageCount;
  if(detailPage < 1) detailPage = 1;
  const start = (detailPage-1) * detailPageSize;
  const pageItems = list.slice(start, start + detailPageSize);
  const isMobile = isMobileView();
  // 月度候选（多选下拉）；日期保持点选式日期控件（单选三态）
  const latest = latestOrderDate();
  const months = [...new Set(orders.map(r=>r.payTime?fmtMonth(r.payTime):null).filter(Boolean))].sort();
  const dates  = [...new Set(orders.map(r=>r.payTime?fmtDate(r.payTime):null).filter(Boolean))].sort();
  // 日期筛选改为日期控件（type=date）：值='' 表示全部日期；默认显示当日（最新日期）；
  // 「当日」按钮可随时回到默认口径。min/max 限定在数据实际日期范围内，便于控件选择。
  const dateMin = dates[0] || '';
  const dateMax = dates[dates.length-1] || '';
  // 字段筛选下拉选项（取自全部订单，与页签无关，保证跨页签稳定）
  const distinct = key => [...new Set(orders.map(r=>String(r[key]||'').trim()).filter(Boolean))]
    .sort((a,b)=>a.localeCompare(b,'zh-CN'));
  // 级联候选：一级渠道随「渠道分类」收敛；二级渠道随「渠道分类+一级渠道」收敛
  const distinctPool = (key, pool) => [...new Set(pool.map(r=>String(r[key]||'').trim()).filter(Boolean))]
    .sort((a,b)=>a.localeCompare(b,'zh-CN'));
  const l1Candidates = distinctPool('channelL1', poolC);
  const l2Candidates = distinctPool('channelL2', poolL1);
  // 到访类型候选（datalist）：数据中已出现的值 + 常用「初诊/复诊」兜底，允许自由输入
  const vtVals = [...new Set(distinct('visitType').concat(
    ['初诊','复诊'].filter(v=>!distinct('visitType').includes(v))))].sort((a,b)=>a.localeCompare(b,'zh-CN'));
  const vtOptions = vtVals.map(v=>`<option value="${esc(v)}"></option>`).join('');
  const fText = (key,label) => `<label class="detail-filter"><span class="df-label">${label}</span>`
    + `<input type="text" class="detail-select" data-dfilter-text="${key}" value="${esc(detailFilters[key]||'')}" `
    + `placeholder="含…"></label>`;
  // 数值筛选（运算符+数值，2026-10-13 需求）：op 存 <key>Op（''=等于），数值存 <key>；数值为空=不筛选
  const fNum = (key,label) => `<span class="detail-filter num-filter"><span class="df-label">${label}</span>`
    + `<span class="num-line">`
    + `<select class="detail-select op-sel" data-dfilter="${key}Op" title="选择比较方式">`
    + NUM_OPS.map(([v,l])=>`<option value="${v}" ${String(detailFilters[key+'Op']||'')===v?'selected':''}>${l}</option>`).join('')
    + '</select>'
    + `<input type="number" step="any" class="detail-select" data-dfilter-text="${key}" value="${esc(detailFilters[key]||'')}" `
    + `placeholder="数值"></span></span>`;
  const effDateOn = !!detailDate && detailDate !== '__latest__';
  // 运算符键（*Op，''=等于）不计入「已激活筛选」，避免仅选运算符未填数值时出现清除按钮/计数虚高
  const activeFilterEntries = Object.entries(detailFilters||{}).filter(([k,v]) => (Array.isArray(v)?v.length:!!v) && !k.endsWith('Op'));
  const anyDetailFilter = !!(detailMonth.length||effDateOn||activeFilterEntries.length);
  const filterBar = `
    <div class="detail-filters ${detailFiltersCollapsed?'collapsed':''}">
      <button type="button" id="detail-filter-toggle" class="ftoggle" title="展开/收起筛选区"><span>⚙</span><span>筛选</span>${activeFilterEntries.length?`<span class="ft-count">${activeFilterEntries.length}</span>`:''}<span class="ft-caret">${detailFiltersCollapsed?'▸':'▴'}</span></button>
      <div class="filters-body">
      ${mselHTML('__month__','月度',months)}
      <label class="detail-filter date-filter"><span class="df-label">日期</span>
        <span class="date-line">
          <input type="date" id="detail-date-filter" class="detail-select"
            min="${dateMin}" max="${dateMax}"
            value="${detailDate==='__latest__'?latest:(detailDate||'')}"
            title="选择筛选日期；清空=全部日期">
          <button type="button" id="detail-date-today" class="detail-reset" title="回到当日（最新日期 ${latest}）">当日</button>
        </span>
      </label>
      ${fText('name','姓名')}
      ${fText('memberNo','会员号')}
      ${mselHTML('channelClass','渠道分类',distinct('channelClass'))}
      ${mselHTML('channelL1','一级渠道',l1Candidates)}
      ${mselHTML('channelL2','二级渠道',l2Candidates)}
      ${mselHTML('visitType','到访类型',distinct('visitType'))}
      ${mselHTML('consultant','咨询师',distinct('consultant'))}
      ${mselHTML('developer','客户开发人',distinct('developer'))}
      ${mselHTML('cashier','收款员工',distinct('cashier'))}
      ${fNum('cashPay','现款支付')}
      ${fNum('assetDeduct','资产抵扣金额')}
      ${mselHTML('dept','报表归属',DEPT_OPTIONS)}
      ${mselHTML('vab','VAB',[['hit','✔ 命中'],['miss','— 未命中']],'span2')}
      ${anyDetailFilter?'<button type="button" id="detail-filter-reset" class="detail-reset" title="清除全部筛选">清除筛选</button>':''}
      </div>
    </div>`;
  const tabs = DETAIL_TABS.map(t=>`<button class="detail-tab ${t.key===detailTab?'active':''}" data-tab="${t.key}">${t.label}</button>`).join('');
  // 导出区（2026-10-21 需求：各页签数据均支持导出）
  const exportOps = `<div class="detail-export-ops">
      <button type="button" class="btn btn-sm" id="detail-export-current" data-rbac="export"
        title="导出当前页签的全部明细（含当前筛选与排序），带样式 Excel">⬇ 导出本页签</button>
      <button type="button" class="btn btn-sm" id="detail-export-all" data-rbac="export"
        title="一次导出全部页签（每个页签一个 sheet），忽略当前筛选">⬇ 导出全部页签</button>
    </div>`;
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
  const dateNote = detailDate === '__latest__' ? ('当日 '+latest+' · ') : (detailDate ? ('日期 '+detailDate+' · ') : '');
  const activeFieldCount = activeFilterEntries.length;
  const filterNote = (detailMonth.length?('月度选'+detailMonth.length+'项 · '):'')+dateNote
    +(activeFieldCount?('字段筛选 '+activeFieldCount+' 项 · '):'');
  // 2026-10-23：全站统一不显示口径说明文字——营销部两类页签的口径提示已从界面移除（仅保留筛选/行数/金额等数据）
  const mktTabNote = '';
  // 手机端排序条（桌面端隐藏，由表格表头排序代替）
  const mobileSort = `
    <div class="detail-mobile-bar">
      <label class="detail-filter">排序
        <select id="detail-mobile-sort" class="detail-select" style="min-width:auto;">
          <option value="">默认顺序</option>
          ${DETAIL_SORT_COLS.map(([k,l])=>`<option value="${k}" ${detailSort.key===k?'selected':''}>${l}</option>`).join('')}
        </select>
      </label>
    </div>`;
  // 内容区：手机端卡片式 / 桌面端表格
  const content = isMobile
    ? `<div class="detail-cards">${pageItems.map(r=>detailCardHTML(r)).join('')}</div>`
    : `<div class="table-wrap" style="max-height:none;overflow:auto;border:1px solid var(--c-border);border-radius:8px;">
        <table class="data">
          <thead><tr>${sortThs}</tr></thead>
          <tbody>${pageItems.map((r,i)=>detailRowCells(r,i,start)).join('')}</tbody>
          <tfoot><tr style="font-weight:700;background:var(--c-blue-tint);color:var(--c-head);">
            <td colspan="11" style="text-align:right;">合计（${list.length} 行 · 命中VAB ${vabCount} 行）</td>
            <td style="text-align:right;">${fmt(total)}</td>
            <td style="text-align:right;">${fmt(totalAsset)}</td>
            <td></td>
            <td style="text-align:center;">${vabCount}</td>
            <td></td>
          </tr></tfoot>
        </table>
      </div>`;
  const rangeNote = list.length > detailPageSize
    ? ` · 本页 ${start+1}-${Math.min(start+detailPageSize, list.length)}`
    : '';
  el.innerHTML = `
    ${filterBar}
    <div class="detail-tabs-row">
      <div class="detail-tabs">${tabs}</div>
      ${exportOps}
    </div>
    ${mobileSort}
    <div style="padding:6px 0 10px;color:var(--c-muted);font-size:12px;">
      当前标签「${tabLabel}」· ${mktTabNote}${filterNote}共 ${list.length} 行 · 现款支付合计 ¥${fmt(total)} · 资产抵扣合计 ¥${fmt(totalAsset)} · 命中VAB ${vabCount} 行${rangeNote}
    </div>
    ${content}
    ${renderPager(detailPage, pageCount)}
    <datalist id="vt-options">${vtOptions}</datalist>`;
  // 事件委托（绑定一次）：标签切换 + 删除某条数据（删除后全表同步更新）+ 筛选交互
  if(!el.dataset.bound){
    el.addEventListener('click', onImportDetailsClick);
    el.addEventListener('change', onImportDetailsChange);
    el.addEventListener('input', onImportDetailsInput);
    // 点击面板外部收起多选下拉（捕获阶段，避免与其他点击逻辑竞争）
    document.addEventListener('click', (ev)=>{
      if(mselRuntime.open && !(ev.target.closest && ev.target.closest('.msel'))){
        mselRuntime.open = null;
        mselRuntime.search = null;
        applyMselOpenState();
      }
    }, true);
    el.dataset.bound = '1';
  }
  // 恢复多选下拉展开态（重渲染后面板不闪没）
  applyMselOpenState();
}

// 导入明细点击：切换标签 / 删除某条订单 / 清除筛选（data-uid 精确删除，删除后各统计表联动重算）
function onImportDetailsClick(e){
  const del = e.target.closest && e.target.closest('.detail-del');
  if(del){ deleteOrder(del.dataset.uid); return; }
  // 筛选区展开/收起（2026-10-13：手机端与电脑端均支持，状态持久化）
  const ftg = e.target.closest && e.target.closest('#detail-filter-toggle');
  if(ftg){
    detailFiltersCollapsed = !detailFiltersCollapsed;
    try{ localStorage.setItem('dp_detail_filters_collapsed', detailFiltersCollapsed?'1':'0'); }catch(err){}
    renderImportDetails();
    return;
  }
  const tab = e.target.closest && e.target.closest('.detail-tab');
  if(tab){ detailTab = tab.dataset.tab; detailPage = 1; renderImportDetails(); return; }
  // 导出当前页签 / 全部页签（2026-10-21 需求）
  const exCur = e.target.closest && e.target.closest('#detail-export-current');
  if(exCur){ exportDetailCurrent(); return; }
  const exAll = e.target.closest && e.target.closest('#detail-export-all');
  if(exAll){ exportDetailAll(); return; }
  // 分页按钮（含上一页/下一页/页码）：data-page 为 1 基页码
  const pg = e.target.closest && e.target.closest('[data-page]');
  if(pg){ detailPage = parseInt(pg.dataset.page, 10) || 1; renderImportDetails(); return; }
  // 表头排序：同列再次点击切换升/降序，不同列默认升序
  const th = e.target.closest && e.target.closest('th[data-sort]');
  if(th){
    const k = th.dataset.sort;
    if(detailSort.key === k) detailSort.dir = -detailSort.dir;
    else detailSort = { key:k, dir:1 };
    detailPage = 1;
    renderImportDetails();
    return;
  }
  const dtBtn = e.target.closest && e.target.closest('#detail-date-today');
  if(dtBtn){ detailDate='__latest__'; detailPage=1; renderImportDetails(); return; }
  const reset = e.target.closest && e.target.closest('#detail-filter-reset');
  if(reset){
    detailMonth = []; detailDate = '__latest__'; detailFilters = {}; detailPage = 1;
    mselRuntime.open = null; mselRuntime.search = null;
    renderImportDetails();
    return;
  }
  // 多选下拉：开合按钮（换面板/新开面板时清空搜索词）
  const mb = e.target.closest && e.target.closest('.msel-btn');
  if(mb){
    const key = mb.closest('.msel').dataset.mskey;
    if(mselRuntime.open !== key) mselRuntime.search = null;
    mselRuntime.open = mselRuntime.open === key ? null : key;
    applyMselOpenState();
    return;
  }
  // 多选下拉：全选（当前可见候选值，不含「空白」）/ 清空（含「空白」一并清除）
  const act = e.target.closest && e.target.closest('.msel-act');
  if(act){
    const box = act.closest('.msel');
    const key = box.dataset.mskey;
    if(act.dataset.msact === 'all'){
      const vals = [...box.querySelectorAll('label.msel-item')]
        .filter(it => it.style.display !== 'none')
        .map(it => it.querySelector('input').value)
        .filter(v => v !== MS_BLANK);
      mselSet(key, vals);
    } else mselSet(key, []);
    detailPage = 1;
    mselRuntime.open = key;
    renderImportDetails();
    return;
  }
}

// 导入明细筛选：月度 / 日期 / 每页条数 / 手机端排序 / 字段下拉 / 报表归属 下拉变更
function onImportDetailsChange(e){
  if(!e || !e.target) return;
  const ds = e.target.closest && e.target.closest('select.dept-sel');
  if(ds){
    const row = state.orders.find(r => String(r._uid) === String(ds.dataset.uid));
    if(row){
      const v = String(ds.value||'').trim();
      const oldDept = effectiveDept(row);            // 改前生效归属（手动覆盖优先）
      if(DEPT_OPTIONS.includes(v)) row.deptOverride = v;
      else delete row.deptOverride; // 恢复自动归属
      const newDept = effectiveDept(row);
      // 清除受影响部门行（含残差「其他」）的手动覆盖 → 报表立即回到按新归属自动计算的值
      const cleared = clearManualForDeptKeys([oldDept, newDept, '其他']);
      persistOrders(); // 归属覆盖随订单持久化（localStorage/IndexedDB/云端）
      recompute(); // 重算 ③主表 / ④门店表 / ⑥明细，报表归属立即生效
      const tip = v ? `该明细报表归属已改为「${v}」` : '已恢复自动归属';
      showToast(tip + '，各统计表已重新计算' + (cleared ? `（已清除 ${cleared} 处手动调整以恢复自动口径）` : ''), 'ok');
    }
    return;
  }
  const vt = e.target.closest && e.target.closest('input.vt-input');
  if(vt){
    const row = state.orders.find(r => String(r._uid) === String(vt.dataset.uid));
    if(row){
      const v = String(vt.value||'').trim();
      const old = String(row.visitType||'').trim();
      if(v !== old){
        row.visitType = v;
        // 到访类型会改变营销部初诊门槛 → 该行归属可能变化，同样清除相关手动覆盖
        const cleared = clearManualForDeptKeys([effectiveDept(row), '其他']);
        persistOrders(); // 到访类型随订单持久化（localStorage/IndexedDB/云端）
        recompute(); // 到访类型影响：营销部初诊门槛 / ⑤清单 / 营销部·初诊与类型空白页签，全联动重算
        showToast((v ? `该明细到访类型已改为「${v}」` : '已清空该明细到访类型') + '，各统计表已重新计算'
          + (cleared ? `（已清除 ${cleared} 处手动调整以恢复自动口径）` : ''), 'ok');
      }
    }
    return;
  }
  const d = e.target.closest && e.target.closest('#detail-date-filter');
  if(d){ detailDate = d.value || ''; detailPage = 1; renderImportDetails(); return; }
  // 数值筛选运算符下拉（cashPayOp / assetDeductOp；值筛选下拉已改为 msel 多选，不再走此分支）
  const op = e.target.closest && e.target.closest('select[data-dfilter]');
  if(op){ detailFilters[op.dataset.dfilter] = op.value || ''; detailPage = 1; renderImportDetails(); return; }
  // 多选下拉复选框：更新选中集合后重渲染（展开态由 mselRuntime.open 恢复，面板不闪没）
  const cb = e.target.closest && e.target.closest('.msel-panel input[type="checkbox"]');
  if(cb){
    const box = cb.closest('.msel');
    const key = box.dataset.mskey;
    const set = new Set(mselGet(key));
    if(cb.checked) set.add(cb.value); else set.delete(cb.value);
    mselSet(key, [...set]);
    detailPage = 1;
    mselRuntime.open = key;
    renderImportDetails();
    return;
  }
  const ps = e.target.closest && e.target.closest('#detail-pagesize-filter');
  if(ps){ detailPageSize = parseInt(ps.value, 10) || 10; detailPage = 1; renderImportDetails(); return; }
  const ms = e.target.closest && e.target.closest('#detail-mobile-sort');
  if(ms){
    const v = ms.value;
    detailSort = v ? { key:v, dir:1 } : { key:'', dir:1 };
    detailPage = 1;
    renderImportDetails();
  }
}

// 字段文本筛选（姓名/会员号模糊包含）：防抖 300ms 后重渲染，并恢复焦点与光标到行尾
let _dfilterTimer = null;
function onImportDetailsInput(e){
  // msel 面板搜索：仅本地过滤候选项显隐，不重渲染
  const sr = e.target.closest && e.target.closest('.msel-search');
  if(sr){
    const box = sr.closest('.msel');
    mselRuntime.search = { key: box.dataset.mskey, q: sr.value };
    applyMselSearchFilter(box, sr.value);
    return;
  }
  const ft = e.target.closest && e.target.closest('input[data-dfilter-text]');
  if(!ft) return;
  const key = ft.dataset.dfilterText;
  const val = ft.value.trim();
  clearTimeout(_dfilterTimer);
  _dfilterTimer = setTimeout(()=>{
    detailFilters[key] = val;
    detailPage = 1;
    renderImportDetails();
    const inp = document.querySelector('input[data-dfilter-text="'+key+'"]');
    if(inp){ inp.focus(); try{ inp.setSelectionRange(inp.value.length, inp.value.length); }catch(err){} }
  }, 300);
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
    renderMktFirstClients();
    persistOrders(); // 删空后同步持久化（空数组），刷新不再出现旧数据
    clearOrdersInCloud(); // 本地已空 → 同步清空云端，避免重新登录后旧数据回灌
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
    const parsed = await parseExcelFile(f);
    const allRows = parsed.map(normalizeOrder);
    // 静默失败防护：若「收款时间」解析失败，.filter(r=>r.payTime) 会把所有行丢弃，
    // 导致 rows 为空、合并 0 条、cutoff 仍停在旧月份——用户看到「刷新后还是 9 月」却毫无线索。
    // 故在此显式拦截并给出可操作提示。
    const rows = allRows.filter(r => r.payTime).map(ensureUid);
    if(parsed.length > 0 && rows.length === 0){
      state.demoOrders = false;
      showOrdersSaved(true);
      showToast(`❌ 导入失败：已读到 ${parsed.length} 行，但「收款时间」全部无法解析（订单已全部丢弃，未写入任何数据）。请确认表头含「收款时间」列且为可识别日期格式`, 'err', 9000);
      $('orders-meta').textContent = `❌ 导入失败：${parsed.length} 行全部缺少可解析的「收款时间」`;
      $('orders-meta').className = 'file-meta';
      return;
    }
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
    persistOrders(false, true); // 持久化导入的订单（含截止日期）；syncNow=true 立即同步云端，确保刷新/关页不丢
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

// ===== 近一年消费明细导入（2026-10-16 新增）=====
// 「营销部顾客一年内无消费记录 → 视为初诊客户」规则的数据源。表结构与订单表一致（订单项目收退款表），
// 仅读取 会员号/姓名/收款时间 参与判定；该数据不计入任何业绩金额；未导入时新客规则自动停用。
async function onFileHistory(input){
  const f = input.files?.[0];
  if(!f) return;
  try{
    await ensureXlsx();
    const parsed = await parseExcelFile(f);
    const allRows = parsed.map(normalizeOrder);
    const rows = allRows.filter(r => r.payTime).map(ensureUid);
    if(parsed.length > 0 && rows.length === 0){
      showToast(`❌ 历史明细导入失败：已读到 ${parsed.length} 行，但「收款时间」全部无法解析。请确认表头含「收款时间」列且为可识别日期格式`, 'err', 9000);
      if($('hist-meta')){ $('hist-meta').textContent = '❌ 导入失败：收款时间全部无法解析'; $('hist-meta').className = 'file-meta'; }
      return;
    }
    state.histOrders = rows;
    persistHist(); // 内部刷新 histTs 并落盘（localStorage + IndexedDB + 壳窗口桥接）
    invalidateMktKeyCaches();
    refreshHistMeta();
    recompute();
    showToast(`近一年消费明细导入成功（${rows.length} 行）。「营销部顾客一年内无消费记录 → 初诊客户」规则已启用`, 'ok', 5200);
  }catch(e){
    console.error(e);
    showToast('历史消费明细解析失败：'+e.message, 'err');
  }
}

// 清除近一年消费明细（不影响订单/VAB/月目标）；清除后新客规则自动停用
function clearHistData(){
  if(!state.histOrders.length){ showToast('尚未导入近一年消费明细，无需清除', 'warn'); return; }
  if(!window.confirm('确定清除已导入的近一年消费明细？\n清除后「营销部顾客一年内无消费记录 → 初诊客户」规则将停用，业绩报表恢复既有口径（30天内 到访类型=初诊）。\n订单 / VAB 会员 / 月目标不受影响。此操作不可撤销。')) return;
  state.histOrders = [];
  histClouded = false;
  persistHist(); // 写入空快照（含壳窗口桥接，防刷新后旧数据回灌）
  invalidateMktKeyCaches();
  refreshHistMeta();
  recompute();
  showToast('已清除近一年消费明细，新客规则已停用', 'ok');
}

// 历史明细状态文案（导入 / 本地恢复 / 清除 后统一刷新）
function refreshHistMeta(){
  const meta = $('hist-meta');
  if(!meta) return;
  if(state.histOrders.length){
    // 云同步状态提示：历史明细随订单整包上云；2026-10-20 起仅明细（无订单）也可单独上云，
    // 故以 histClouded 标志反映真实同步状态，而非「是否导入了订单」。
    meta.textContent = `✔ 历史明细（${state.histOrders.length} 行）· 新客规则已启用 · ` + (histClouded ? '已同步云端' : '暂仅保存在本机');
  }else{
    meta.textContent = '未导入（新客规则停用）';
  }
  meta.className = 'file-meta';
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

// ===== ⑤ 营销部初诊客户（近30天）=====
// 口径：来源渠道=营销部（与「营销部」页签一致，含推荐→营销部落位）且 到访类型=初诊；
// 近30天 = 截止日起往前30天（含两端，可跨月）；按客户去重（会员号优先，无会员号按姓名）。
// 近30天滚动窗口（截止日往前推29天，含两端共30天，可跨月）：⑤清单 / ③报表门槛 / ⑥营销部页签 统一口径
function mktWindow(cutoffOverride){
  const cutoff = cutoffOverride || (state.cutoffDate ? new Date(state.cutoffDate) : new Date());
  if(isNaN(cutoff.getTime())) return { start:'', end:'' };
  return { start: fmtDate(new Date(cutoff.getTime() - 29*86400000)), end: fmtDate(cutoff) };
}
// 客户键归一化（2026-10-20 修复「历史明细会员号格式与订单不一致 → 新客规则失效」）：
// 订单表与历史明细导出的会员号常带 V 前缀 / 前导零 / 空格（如 M26374801 vs V26374801 vs 026374801），
// 原 mktFirstClientKey 用原始字符串精确匹配 → 永远 mismatch → 规则B 误把回头老客也判成新客。
// 在此统一「去非数字 + 去前导零」（与 VAB buildVabSet 同思路），使历史↔订单跨格式可靠匹配；
// 无会员号时回退到「name:姓名」。histKeysCached 与 buildMktFirstKeys 两侧均经此函数 → 对称一致。
function normMemberKey(no){
  const s = String(no == null ? '' : no).trim().toLowerCase();
  const d = s.replace(/\D/g, '');
  return d ? d.replace(/^0+/, '') : '';
}
function mktFirstClientKey(r){
  const nk = normMemberKey(r.memberNo);
  return nk ? nk : ('name:'+String(r.name||'').trim());
}
function isMktChuzhenOrder(r){
  return (chanHas(r,'营销部') || isRecToYingxiao(r)) && (r.visitType||'').trim() === '初诊';
}
// ⑤ 筛选状态：姓名/会员号模糊包含，来源渠道/初诊日期精确，消费笔数=运算符(cntOp)+数值(cnt)
let mktFilters = { name:'', memberNo:'', channel:[], date:'', cnt:'', cntOp:'' }; // channel 为多选数组（[]=全部）
let _mktFilterTimer = null;
let mktPage = 1;      // ⑤清单分页：当前页（1 基）
let mktPageSize = 10; // ⑤清单分页：每页条数
// 筛选区折叠状态（⑤清单）：localStorage 持久化
let mktFiltersCollapsed = false;
try{ mktFiltersCollapsed = localStorage.getItem('dp_mkt_filters_collapsed') === '1'; }catch(e){}
function mktChannelText(r){
  return [r.channelClass, r.channelL1, r.channelL2].map(v=>String(v||'').trim()).filter(Boolean).join('-');
}
function renderMktFirstClients(){
  const el = $('mkt-first-clients');
  if(!el) return;
  if(!state.results || !state.orders.length){
    el.innerHTML = '<div class="file-meta" style="padding:12px 0;">导入订单数据后显示近30天营销部初诊客户</div>';
    return;
  }
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  // 近30天滚动窗口（与业绩归属口径一致）：截止日往前推29天，含两端共30天，可跨月
  const winEnd = fmtDate(cutoff);
  const winStart = fmtDate(new Date(cutoff.getTime() - 29*86400000));
  const winLabel = '近30天（' + winStart + ' ~ ' + winEnd + '）';
  const inWin = r => { if(!r.payTime) return false; const d = fmtDate(r.payTime); return d >= winStart && d <= winEnd; };
  // 按客户分组（初诊客户 = 规则A ∪ 规则B，与业绩口径完全一致）：
  // 规则A：窗口内 营销部渠道 且 到访类型=初诊；规则B：营销部渠道 且 近一年无消费·新客（导入历史明细后启用）。
  // 实现上直接用与 calculate 一致的初诊键集合（mktFirstKeysCached）过滤窗口内营销部订单。
  const gateSet = mktFirstKeysCached();
  // 判定依据预计算：该客户任意一笔订单 到访类型=初诊 → 规则A；否则（在集合中）→ 规则B新客
  const chuzhenTypeKeys = new Set();
  state.orders.forEach(r => { if(isMktChuzhenOrder(r)){ const k = mktFirstClientKey(r); if(k) chuzhenTypeKeys.add(k); } });
  const groups = new Map();
  state.orders.forEach(r=>{
    if(!inWin(r)) return;
    if(!(chanHas(r,'营销部') || isRecToYingxiao(r))) return;
    const key = mktFirstClientKey(r);
    if(!key || key === 'name:' || !gateSet.has(key)) return;
    if(!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });
  if(groups.size === 0){
    el.innerHTML = '<div class="file-meta" style="padding:12px 0;">近30天（'+winStart+' ~ '+winEnd+'）暂无营销部初诊客户（判定=到访类型初诊 ∪ 一年内无消费·新客）</div>';
    return;
  }
  // 来源渠道 = 渠道分类-一级渠道-二级渠道 拼接（空段跳过）
  const allClients = [...groups.entries()].map(([key, rows])=>{
    // 初诊日期 = 该客户当月最早一笔初诊消费日期
    let first = null;
    rows.forEach(r=>{ if(r.payTime && (!first || r.payTime < first.payTime)) first = r; });
    const r0 = rows[0];
    return { key, name: r0.name || '—',
      memberNo: String(r0.memberNo||'').trim() || '—',
      channel: mktChannelText(r0) || '—',
      basis: chuzhenTypeKeys.has(key) ? '到访类型=初诊' : '一年内无消费·新客',
      firstDate: first && first.payTime ? fmtDate(first.payTime) : '—',
      cnt: state.orders.filter(r=>inWin(r) && mktFirstClientKey(r)===key).length };
  });
  // 筛选（选项取自全部客户，与当前筛选无关，保证跨筛选稳定）
  const f = mktFilters || {};
  const channelOpts = [...new Set(allClients.map(c=>c.channel))].sort((a,b)=>a.localeCompare(b,'zh-CN'));
  const dateOpts = [...new Set(allClients.map(c=>c.firstDate))].sort();
  let clients = allClients;
  if(f.name){ const q=String(f.name).toLowerCase(); clients=clients.filter(c=>String(c.name).toLowerCase().includes(q)); }
  if(f.memberNo){ const q=String(f.memberNo).toLowerCase(); clients=clients.filter(c=>String(c.memberNo).toLowerCase().includes(q)); }
  if(f.channel && f.channel.length) clients=clients.filter(c=>f.channel.includes(c.channel));
  if(f.date) clients=clients.filter(c=>c.firstDate===f.date);
  // 消费笔数（运算符+数值，2026-10-13 需求）：数值非空才生效；op 存 cntOp（''=等于）
  {
    const raw = String(f.cnt||'').trim();
    if(raw !== '' && !isNaN(Number(raw))){
      const t = Number(raw), op = f.cntOp || 'eq';
      clients = clients.filter(c => cmpNum(c.cnt, op, t));
    }
  }
  const anyF = !!(f.name||f.memberNo||(f.channel&&f.channel.length)||f.date||f.cnt);
  const filterBar = `
    <div class="detail-filters ${mktFiltersCollapsed?'collapsed':''}" style="margin:0 0 8px;">
      <button type="button" id="mkt-filter-toggle" class="ftoggle" title="展开/收起筛选区"><span>⚙</span><span>筛选</span>${anyF?'<span class="ft-count">生效</span>':''}<span class="ft-caret">${mktFiltersCollapsed?'▸':'▴'}</span></button>
      <div class="filters-body">
      <label class="detail-filter"><span class="df-label">姓名</span>
        <input type="text" class="detail-select" data-mkt-filter="name" value="${esc(f.name||'')}" placeholder="含…">
      </label>
      <label class="detail-filter"><span class="df-label">会员号</span>
        <input type="text" class="detail-select" data-mkt-filter="memberNo" value="${esc(f.memberNo||'')}" placeholder="含…">
      </label>
      ${mselHTML('mkt:channel','来源渠道',channelOpts,'span2')}
      <label class="detail-filter date-filter"><span class="df-label">初诊日期</span>
        <span class="date-line">
          <input type="date" class="detail-select" data-mkt-filter="date"
            min="${dateOpts[0]||''}" max="${dateOpts[dateOpts.length-1]||''}" value="${esc(f.date||'')}"
            title="按初诊日期筛选；清空=全部">
        </span>
      </label>
      <span class="detail-filter num-filter"><span class="df-label">消费笔数</span>
        <span class="num-line">
          <select class="detail-select op-sel" data-mkt-filter="cntOp" title="选择比较方式">
            ${NUM_OPS.map(([v,l])=>`<option value="${v}" ${String(f.cntOp||'')===v?'selected':''}>${l}</option>`).join('')}
          </select>
          <input type="number" step="any" class="detail-select" data-mkt-filter="cnt" value="${esc(f.cnt||'')}"
            placeholder="笔数" title="输入笔数；清空=全部">
        </span>
      </span>
      <span class="filters-foot">
        ${anyF?'<button type="button" id="mkt-filter-reset" class="detail-reset" title="清除全部筛选">清除筛选</button>':''}
        <span>${clients.length}/${allClients.length} 位</span>
      </span>
      </div>
    </div>`;
  clients.sort((a,b)=> String(a.firstDate).localeCompare(String(b.firstDate)) || String(a.name).localeCompare(String(b.name),'zh-CN'));
  // 分页（与⑥明细一致的分页器样式；筛选后列表分页，页码窗口化）
  const pageCount = Math.max(1, Math.ceil(clients.length / mktPageSize));
  if(mktPage > pageCount) mktPage = pageCount;
  if(mktPage < 1) mktPage = 1;
  const pStart = (mktPage-1) * mktPageSize;
  const pageItems = clients.slice(pStart, pStart + mktPageSize);
  const pEnd = Math.min(pStart + mktPageSize, clients.length);
  const nums = [...new Set([1, pageCount, mktPage, mktPage-1, mktPage+1].filter(n=>n>=1&&n<=pageCount))].sort((a,b)=>a-b);
  const pagerParts = [`<button type="button" class="pg-btn" data-mpage="${mktPage-1}" ${mktPage<=1?'disabled':''}>上一页</button>`];
  let prev = 0;
  nums.forEach(n=>{
    if(n-prev>1) pagerParts.push('<span class="pg-sep">…</span>');
    pagerParts.push(`<button type="button" class="pg-btn ${n===mktPage?'active':''}" data-mpage="${n}">${n}</button>`);
    prev = n;
  });
  pagerParts.push(`<button type="button" class="pg-btn" data-mpage="${mktPage+1}" ${mktPage>=pageCount?'disabled':''}>下一页</button>`);
  const sizeOpts = [10,20,50,100].map(s=>`<option value="${s}" ${s===mktPageSize?'selected':''}>每页 ${s}</option>`).join('');
  const pagerHTML = `
    <div class="detail-pager">
      ${pagerParts.join('')}
      <span class="pg-info">第 ${mktPage}/${pageCount} 页</span>
      <label class="detail-filter pager-size">每页
        <select id="mkt-pagesize" class="detail-select" style="min-width:auto;">${sizeOpts}</select>
      </label>
    </div>`;
  // 内容区：手机端卡片 / 桌面端表格（2026-10-13：手机端⑤清单改为卡片形式）
  const isMobile = isMobileView();
  const listHTML = clients.length === 0
    ? '<div class="file-meta" style="padding:12px 0;">无匹配客户（可调整或清除筛选）</div>'
    : (isMobile
      ? `<div class="detail-cards">${pageItems.map((c,i)=>{
          const rows = state.orders.filter(r=>inWin(r) && mktFirstClientKey(r)===c.key);
          const cnt = rows.length;
          const sum = rows.reduce((a,b)=>a+(Number(b.cashPay)||0),0);
          return `<div class="detail-card mkt-first-card" data-key="${esc(c.key)}" title="点击查看该客户近30天消费明细">
            <div class="dc-head">
              <span class="dc-name">${esc(c.name)}</span>
              <span class="dc-vab">#${pStart+i+1} · ${cnt} 笔</span>
            </div>
            <div class="dc-rows">
              ${`<span>会员号</span><b>${esc(c.memberNo)}</b>`}
              ${`<span>来源渠道</span><b>${esc(c.channel)}</b>`}
              ${`<span>判定依据</span><b>${esc(c.basis)}</b>`}
              ${`<span>初诊日期</span><b>${esc(c.firstDate)}</b>`}
              ${`<span>近30天消费笔数</span><b>${cnt}</b>`}
              ${`<span>近30天现款合计</span><b class="amt">${fmt(sum)}</b>`}
            </div>
          </div>`;
        }).join('')}</div>`
      : `<div class="table-wrap" style="max-height:none;overflow:auto;border:1px solid var(--c-border);border-radius:8px;">
      <table class="data">
        <thead><tr><th>#</th><th>客户姓名</th><th>客户会员号</th><th>来源渠道</th><th>判定依据</th><th>初诊日期</th><th>近30天消费笔数</th><th>近30天现款合计</th></tr></thead>
        <tbody>${pageItems.map((c,i)=>{
          const rows = state.orders.filter(r=>inWin(r) && mktFirstClientKey(r)===c.key);
          const cnt = rows.length;
          const sum = rows.reduce((a,b)=>a+(Number(b.cashPay)||0),0);
          return `<tr class="mkt-first-row" data-key="${esc(c.key)}" style="cursor:pointer;" title="点击查看该客户近30天消费明细">
            <td style="text-align:center;color:var(--c-muted);">${pStart+i+1}</td>
            <td>${esc(c.name)}</td>
            <td>${esc(c.memberNo)}</td>
            <td>${esc(c.channel)}</td>
            <td style="text-align:center;font-size:12px;color:${c.basis==='到访类型=初诊'?'var(--c-muted)':'#1f8a4a'};">${esc(c.basis)}</td>
            <td style="text-align:center;">${esc(c.firstDate)}</td>
            <td style="text-align:center;">${cnt}</td>
            <td style="text-align:right;">${fmt(sum)}</td>
          </tr>`;
        }).join('')}
        </tbody>
      </table>
    </div>`);
  el.innerHTML = filterBar + listHTML + pagerHTML + `
    <div style="padding:6px 0 0;color:var(--c-muted);font-size:12px;">共 ${clients.length} 位初诊客户（${winLabel}）· 本页 ${pStart+1}-${pEnd} · 来源渠道=渠道分类-一级渠道-二级渠道 · 点击${isMobile?'卡片':'数据行'}查看该客户近30天消费明细</div>`;
  if(!el.dataset.bound){
    el.addEventListener('click', onMktFirstClick);
    el.addEventListener('click', onMktFirstFilterClick);
    el.addEventListener('change', onMktFirstFilterChange);
    el.addEventListener('input', onMktFirstFilterInput);
    el.dataset.bound = '1';
  }
  // 恢复多选下拉展开态（重渲染后面板不闪没）
  applyMselOpenState();
}

// ⑤ 筛选/分页交互：翻页按钮 + 下拉/重置 + 文本防抖（保持焦点）+ 多选渠道
function onMktFirstFilterClick(e){
  const pg = e.target.closest && e.target.closest('[data-mpage]');
  if(pg){ mktPage = parseInt(pg.dataset.mpage, 10) || 1; renderMktFirstClients(); return; }
  // 筛选区展开/收起（2026-10-13：手机端与电脑端均支持，状态持久化）
  const ftg = e.target.closest && e.target.closest('#mkt-filter-toggle');
  if(ftg){
    mktFiltersCollapsed = !mktFiltersCollapsed;
    try{ localStorage.setItem('dp_mkt_filters_collapsed', mktFiltersCollapsed?'1':'0'); }catch(err){}
    renderMktFirstClients();
    return;
  }
  // 多选下拉：开合按钮 / 全选（可见候选）/ 清空
  const mb = e.target.closest && e.target.closest('.msel-btn');
  if(mb){
    const key = mb.closest('.msel').dataset.mskey;
    if(mselRuntime.open !== key) mselRuntime.search = null;
    mselRuntime.open = mselRuntime.open === key ? null : key;
    applyMselOpenState();
    return;
  }
  const act = e.target.closest && e.target.closest('.msel-act');
  if(act){
    const box = act.closest('.msel');
    const key = box.dataset.mskey;
    if(act.dataset.msact === 'all'){
      const vals = [...box.querySelectorAll('label.msel-item')]
        .filter(it => it.style.display !== 'none')
        .map(it => it.querySelector('input').value)
        .filter(v => v !== MS_BLANK);
      mselSet(key, vals);
    } else mselSet(key, []);
    mktPage = 1;
    mselRuntime.open = key;
    renderMktFirstClients();
    return;
  }
  const rst = e.target.closest && e.target.closest('#mkt-filter-reset');
  if(rst){ mktFilters = { name:'', memberNo:'', channel:[], date:'', cnt:'', cntOp:'' }; mktPage = 1; mselRuntime.open = null; renderMktFirstClients(); }
}
function onMktFirstFilterChange(e){
  // 多选下拉复选框：更新选中集合后重渲染（展开态由 mselRuntime.open 恢复）
  const cb = e.target.closest && e.target.closest('.msel-panel input[type="checkbox"]');
  if(cb){
    const box = cb.closest('.msel');
    const key = box.dataset.mskey;
    const set = new Set(mselGet(key));
    if(cb.checked) set.add(cb.value); else set.delete(cb.value);
    mselSet(key, [...set]);
    mktPage = 1;
    mselRuntime.open = key;
    renderMktFirstClients();
    return;
  }
  const ps = e.target.closest && e.target.closest('#mkt-pagesize');
  if(ps){ mktPageSize = parseInt(ps.value, 10) || 10; mktPage = 1; renderMktFirstClients(); return; }
  // 下拉与日期控件（type=date）统一走 change：清空=全部
  const sel = e.target.closest && e.target.closest('select[data-mkt-filter], input[type="date"][data-mkt-filter]');
  if(sel){ mktFilters[sel.dataset.mktFilter] = sel.value || ''; mktPage = 1; renderMktFirstClients(); }
}
function onMktFirstFilterInput(e){
  // msel 面板搜索：仅本地过滤候选项显隐，不重渲染
  const sr = e.target.closest && e.target.closest('.msel-search');
  if(sr){
    const box = sr.closest('.msel');
    mselRuntime.search = { key: box.dataset.mskey, q: sr.value };
    applyMselSearchFilter(box, sr.value);
    return;
  }
  const inp = e.target.closest && e.target.closest('input[data-mkt-filter]');
  if(!inp || inp.type === 'date') return; // 日期控件由 change 事件承担，避免防抖打断选择
  const key = inp.dataset.mktFilter;
  const val = inp.value.trim();
  clearTimeout(_mktFilterTimer);
  _mktFilterTimer = setTimeout(()=>{
    mktFilters[key] = val;
    mktPage = 1;
    renderMktFirstClients();
    const inp2 = document.querySelector('#mkt-first-clients input[data-mkt-filter="'+key+'"]');
    if(inp2){ inp2.focus(); try{ inp2.setSelectionRange(inp2.value.length, inp2.value.length); }catch(err){} }
  }, 300);
}

// ⑤ 行点击 → 弹窗展示该客户近30天消费明细（该客户窗口内全部订单，不限营销部初诊）
function onMktFirstClick(e){
  const tr = (e.target.closest && (e.target.closest('tr.mkt-first-row') || e.target.closest('.mkt-first-card')));
  if(!tr) return;
  const key = tr.dataset.key;
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const winEnd = fmtDate(cutoff);
  const winStart = fmtDate(new Date(cutoff.getTime() - 29*86400000));
  const inWin = r => { if(!r.payTime) return false; const d = fmtDate(r.payTime); return d >= winStart && d <= winEnd; };
  const rows = state.orders.filter(r=>inWin(r) && mktFirstClientKey(r)===key);
  if(!rows.length){ showToast('未找到该客户近30天数据', 'warn'); return; }
  showClientMonthModal(rows[0], rows, '近30天 ' + winStart + ' ~ ' + winEnd);
}

// 客户近30天消费明细弹窗（复用 .modal-mask/.modal-box 体系）
function showClientMonthModal(r0, rows, month){
  const mask = document.createElement('div');
  mask.className = 'modal-mask';
  mask.style.zIndex = '1100';
  const sorted = rows.slice().sort((a,b)=>(a.payTime&&b.payTime)?(a.payTime-b.payTime):0);
  const total = sorted.reduce((a,b)=>a+(Number(b.cashPay)||0),0);
  const totalAsset = sorted.reduce((a,b)=>a+(Number(b.assetDeduct)||0),0);
  // 手机端：明细以卡片形式展示（2026-10-13）；桌面端保持表格
  const bodyHTML = isMobileView()
    ? `<div class="detail-cards">${sorted.map((r,i)=>{
        const row = (k,v,cls='') => `<span>${k}</span><b class="${cls}">${v}</b>`;
        return `<div class="detail-card">
          <div class="dc-head"><span class="dc-name">#${i+1} · ${r.payTime?esc(fmtDate(r.payTime)):''}</span><span class="dc-vab">${esc(r.visitType)||'（空白）'}</span></div>
          <div class="dc-rows">
            ${row('渠道分类', esc(r.channelClass))}
            ${row('一级渠道', esc(r.channelL1))}
            ${row('开单咨询师', esc(r.consultant))}
            ${row('收款员工', esc(r.cashier))}
            ${row('现款支付', esc(fmt(r.cashPay)), 'amt')}
            ${row('资产抵扣金额', esc(fmt(r.assetDeduct)), 'amt')}
          </div>
        </div>`;
      }).join('')}
      <div class="detail-card" style="background:var(--c-blue-tint);">
        <div class="dc-rows">
          <span>合计（${sorted.length} 笔）</span><b class="amt">现款 ${fmt(total)}</b>
          <span></span><b class="amt">资产抵扣 ${fmt(totalAsset)}</b>
        </div>
      </div>
    </div>`
    : `<div class="table-wrap" style="max-height:54vh;overflow:auto;border:1px solid var(--c-border);border-radius:8px;">
        <table class="data">
          <thead><tr><th>#</th><th>收款时间</th><th>到访类型</th><th>渠道分类</th><th>一级渠道</th><th>开单咨询师</th><th>收款员工</th><th>现款支付</th><th>资产抵扣金额</th></tr></thead>
          <tbody>${sorted.map((r,i)=>`
            <tr>
              <td style="text-align:center;color:var(--c-muted);">${i+1}</td>
              <td>${r.payTime?fmtDate(r.payTime):''}</td>
              <td>${esc(r.visitType)}</td>
              <td>${esc(r.channelClass)}</td>
              <td>${esc(r.channelL1)}</td>
              <td>${esc(r.consultant)}</td>
              <td>${esc(r.cashier)}</td>
              <td style="text-align:right;">${fmt(r.cashPay)}</td>
              <td style="text-align:right;">${fmt(r.assetDeduct)}</td>
            </tr>`).join('')}
          </tbody>
          <tfoot><tr style="font-weight:700;background:var(--c-blue-tint);color:var(--c-head);">
            <td colspan="7" style="text-align:right;">合计</td>
            <td style="text-align:right;">${fmt(total)}</td>
            <td style="text-align:right;">${fmt(totalAsset)}</td>
          </tr></tfoot>
        </table>
      </div>`;
  mask.innerHTML = `
    <div class="modal-box" role="dialog" aria-modal="true" style="width:min(820px,96vw);">
      <div class="modal-title">📋 ${esc(r0.name||'—')}（会员号 ${esc(String(r0.memberNo||'').trim()||'无')}）· ${month} 消费明细 · 共 ${sorted.length} 笔</div>
      ${bodyHTML}
      <div class="row" style="justify-content:flex-end;margin-top:12px;">
        <button class="btn" id="client-modal-close">关闭</button>
      </div>
    </div>`;
  const close = ()=> mask.remove();
  mask.addEventListener('click', e=>{ if(e.target === mask) close(); });
  mask.querySelector('#client-modal-close').addEventListener('click', close);
  document.body.appendChild(mask);
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
  // 「30天内营销部初诊客户」集合：用全量订单 + 当日为截止构建（buildMktFirstKeys 内部做窗口过滤），
  // 避免按日计算时初诊订单不在当天/当月导致该客户后续消费被误挡在营销部各类别之外
  const c = new Date(dateStr);
  return calculate(dayOrders, vabSet, c, buildMktFirstKeys(orders, c));
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
  renderMktFirstClients();
  refreshDebug();
}

// 刷新调试条：调试横幅已按需求下线（2026-09-29），此处仅确保横幅保持隐藏，不再展示任何调试信息
function refreshDebug(){
  const dbg = $('debug-banner');
  if(dbg){ dbg.style.display = 'none'; }
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
    const targetKey = (row.rowspanTarget > 1) ? (row.targetMergeKey || row.sub || row.targetKey) : row.targetKey;
    const target = state.targets[targetKey] ?? 0;
    // 月累计展示值必须与屏幕一致：用 computeMonthValues（按日重算后求和，含每日覆盖 + 月累计覆盖优先），
    // 而非原始 state.results.values（不含每日覆盖），否则「双击改每日」后屏幕月累计更新、导出月累计却仍是旧值。
    const dispValues = monthDispValues(computeMonthValues(dateList, dailyResults));
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
    const dispValues = monthDispValues(computeMonthValues(dateList, dailyResults));
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

/* ===== ⑤ 电商组各二级渠道业绩（初诊口径，日/月） =====
 * 需求（2026-10-24）：门店报表下方新增报表，统计「电商组」各二级渠道的日业绩与本月累计业绩，
 * 口径与 ③ 业绩总表「营销部」一致（初诊业绩 = 30天/近一年无消费 的营销部初诊客户消费，计入现款支付）。
 * 结构（10-24 晚定稿）：行沿用最初格式——每个二级渠道 2 行（当日数据 / 月累计），列仍为每日一列；
 *   默认只显示「最新有数据的日期」列，支持单列收起、「☑ 批量」多选隐藏、「+展开」（与 ③ 业绩报表一致）。
 */
const EC_SUB_PLATFORMS = {
  '美大-高端体检': '海蓓国际医疗·高端体检中心',
  '美大-国际门诊': '海蓓国际医疗·高新中心医院',
  '美大-口腔保健': '海蓓国际医疗·口腔保健美学中心',
  '美大-新妇科·抗衰': '海蓓国际医疗·女性健康中心',
  '美大-新妇科·私密': '海蓓Anna·皮肤私密养护中心',
  '美大-北郊保健': '海蓓·私享保健中心',
  '美大-北郊医疗': '海蓓国际医疗·医疗保健中心（未央分院）',
};
const EC_SUB_ORDER = Object.keys(EC_SUB_PLATFORMS);
// 与 calculate 内 DEPT_KEYS 保持一致（手动「报表归属」键）
const EC_SUB_OV_KEYS = ['品牌组-其他','活动组-活动支持','电商组','新媒体组','短视频组','品牌组-自然到店','商务部','其他'];

// ⑤ 列可见性状态：null = 尚未手动调整（默认仅显示最新有数据日期列）；Set = 用户自定义隐藏的日期集合
let ecSubHidden = null;
let ecSubBatchMode = false;

function computeEcSubReport(){
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const dateStr = fmtDate(cutoff);
  const dateList = generateDateRange(cutoff); // 当月 1 日 ~ 截止日，每天一列
  const res = { dateStr, dateList, rows: [], totalMonth: 0, totalCount: 0, totalDays: {}, totalCum: {}, totalCntDays: {}, defaultIdx: -1, hasOrders: Array.isArray(state.orders) && state.orders.length > 0 };
  if(!res.hasOrders) return res;

  const CC = r => (r.channelClass || '').trim();
  const L1 = r => (r.channelL1 || '').trim();
  const L2 = r => (r.channelL2 || '').trim();
  const mktFirstKeys = buildMktFirstKeys(state.orders, cutoff);
  const isMktFirst = r => mktFirstKeys.has(mktFirstClientKey(r));
  const hasOv = r => EC_SUB_OV_KEYS.includes(String(r.deptOverride || '').trim());

  // 电商组订单集：完全复刻 ③ 主表 电商组 行（r3 + r3c + r3b + ov），再按二级渠道下钻
  const ecAuto = state.orders.filter(r => {
    if(hasOv(r)) return false;
    if(!isMktFirst(r)) return false;                       // 初诊门槛（与主表 autoYx 一致）
    if(CC(r)==='营销部' && ['电商组','客户转介绍','企业团检'].includes(L1(r))) return true;
    if(CC(r)==='营销部' && L1(r)==='品牌组' && L2(r).includes('线上400')) return true;
    if(isRecToYingxiao(r) && !L2(r).includes('自然到店')) return true;
    return false;
  });
  const ecOv = state.orders.filter(r => String(r.deptOverride || '').trim() === '电商组');
  const ecOrders = ecAuto.concat(ecOv);

  // 月累计：截止日所在「本月」且 <= 截止日（与 calculate 同口径，不含历史月份）
  const cy = cutoff.getFullYear(), cm = cutoff.getMonth();
  const monthSet = ecOrders.filter(r => r.payTime
    && r.payTime.getFullYear() === cy && r.payTime.getMonth() === cm && inDateRange(r.payTime, cutoff));

  // 桶：7 个平台先建好（保证「没有数据也显示 0」），数据中出现的其他二级渠道再追加
  // dayCnt：按日去重客户集合（2026-10-26 客量区分当日/累计）
  // ccCnt/l1Cnt：渠道分类 / 一级渠道 众数统计（2026-10-26 界面新增两列，取桶内月订单最高频非空值）
  const map = new Map();
  const newBucket = () => ({ month: 0, cnt: new Set(), days: new Map(), dayCnt: new Map(), ccCnt: new Map(), l1Cnt: new Map() });
  EC_SUB_ORDER.forEach(l2 => map.set(l2, newBucket()));
  const bump = (m, v) => { if(v) m.set(v, (m.get(v) || 0) + 1); };
  const topVal = m => { let bv = '—', bc = 0; m.forEach((c, v) => { if(c > bc){ bc = c; bv = v; } }); return bv; };
  monthSet.forEach(r => {
    const l2 = L2(r) || '（未分类）';
    if(!map.has(l2)) map.set(l2, newBucket());
    const b = map.get(l2);
    const amt = Number(r.cashPay) || 0;
    b.month += amt;
    const ck = mktFirstClientKey(r);
    if(ck && ck !== 'name:') b.cnt.add(ck);           // 客量：去重客户数
    if(r.payTime){
      const d = fmtDate(r.payTime);
      b.days.set(d, (b.days.get(d) || 0) + amt);
      if(ck && ck !== 'name:'){
        if(!b.dayCnt.has(d)) b.dayCnt.set(d, new Set());
        b.dayCnt.get(d).add(ck);                      // 当日客量：当日去重客户数
      }
    }
    bump(b.ccCnt, CC(r));                             // 渠道分类 / 一级渠道 众数
    bump(b.l1Cnt, L1(r));
  });

  const ordered = EC_SUB_ORDER.slice();
  [...map.keys()].forEach(k => { if(!ordered.includes(k)) ordered.push(k); });
  res.rows = ordered.map(l2 => {
    const b = map.get(l2) || newBucket();
    const days = {}, cum = {}, cntDays = {};
    let acc = 0;
    dateList.forEach(d => {
      const v = b.days.get(d) || 0;
      days[d] = v; acc += v; cum[d] = acc;      // 月累计行：累计至该日
      cntDays[d] = b.dayCnt.has(d) ? b.dayCnt.get(d).size : 0;   // 当日客量：当日去重客户数
    });
    return { l2, platform: EC_SUB_PLATFORMS[l2] || '—', month: b.month, cnt: b.cnt.size, days, cum, cntDays,
      channelClass: topVal(b.ccCnt), channelL1: topVal(b.l1Cnt) };
  });

  res.totalDays = {}; res.totalCum = {}; res.totalCntDays = {};
  let tAcc = 0;
  dateList.forEach(d => {
    const dv = res.rows.reduce((a,r) => a + (r.days[d] || 0), 0);
    res.totalDays[d] = dv; tAcc += dv; res.totalCum[d] = tAcc;
    res.totalCntDays[d] = res.rows.reduce((a,r) => a + (r.cntDays[d] || 0), 0);   // 当日客量合计 = 各渠道之和
  });
  res.totalMonth = res.rows.reduce((a,r) => a + r.month, 0);
  res.totalCount = res.rows.reduce((a,r) => a + r.cnt, 0);
  // 默认可见列：最后一个「有数据」的日期；全月无数据则取最后一列
  for(let i = dateList.length - 1; i >= 0; i--){
    if((res.totalDays[dateList[i]] || 0) > 0){ res.defaultIdx = i; break; }
  }
  if(res.defaultIdx < 0 && dateList.length) res.defaultIdx = dateList.length - 1;
  return res;
}

// ⑤ 当前应隐藏的日期集合：用户未手动调整过（null）时按「仅显示最新有数据日期列」的默认规则
function ecSubHiddenSet(rep){
  const list = (rep && rep.dateList) || [];
  if(ecSubHidden === null){
    const s = new Set(list);
    if(rep && rep.defaultIdx >= 0) s.delete(list[rep.defaultIdx]);
    return s;
  }
  return new Set(list.filter(d => ecSubHidden.has(d)));
}

function renderEcSubReport(){
  const table = $('ec-sub-table');
  if(!table) return;
  const head = $('ec-sub-head'), body = $('ec-sub-body');
  const rep = computeEcSubReport();
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const monthStr = (cutoff.getMonth()+1) + '月' + cutoff.getDate() + '日';
  const dateList = rep.dateList || [];
  const nCols = 6 + dateList.length;

  if(!state.results || !rep.hasOrders){
    const msg = !state.results ? '请先在"① 数据导入"上传订单表' : '当前数据中无「电商组」初诊业绩';
    head.innerHTML = `<tr><th colspan="${nCols}" class="sub-head" style="background:#fff;color:var(--c-text);">${msg}</th></tr>`;
    body.innerHTML = '';
    table.querySelectorAll('colgroup').forEach(c=>c.remove());
    // 无数据：列可见性回到默认规则
    ecSubHidden = null; ecSubBatchMode = false;
    table.classList.remove('batch-mode');
    const bb = $('ec-sub-batch-toggle'); if(bb) bb.textContent = '☑ 批量';
    return;
  }

  // 只渲染「可见」日期列（隐藏列不进 DOM，表格不会被幽灵列撑出空白）
  const hidden = ecSubHiddenSet(rep);
  const vDates = dateList.filter(d => !hidden.has(d));
  const idxOf = new Map(dateList.map((d,i)=>[d,i]));

  head.innerHTML = `
    <tr><th colspan="${6 + vDates.length}" class="title-cell">电商组各二级渠道业绩（初诊业绩，截止${monthStr}）</th></tr>
    <tr>
      <th class="sub-head">平台名称</th>
      <th class="sub-head">二级渠道</th>
      <th class="sub-head">渠道分类</th>
      <th class="sub-head">一级渠道</th>
      <th class="sub-head">业绩类型</th>
      <th class="sub-head">客量<br>(人)<span class="copy-col" data-col="cnt" title="复制客量整列（每渠道 2 值：当日/累计）">复制</span></th>
      ${vDates.map(d=>{
        const dt = new Date(d);
        return `<th class="sub-head daily-col" data-date="${d}" data-idx="${idxOf.get(d)}" title="点击收起该日期列">${dt.getMonth()+1}月${dt.getDate()}日<span class="copy-col" data-col="d:${d}" title="复制该日期整列数据">复制</span></th>`;
      }).join('')}
    </tr>`;

  // 客量(人)列拆两行（2026-10-26）：当日数据行=当日客量（取最新可见日期列）、月累计行=当月去重客量
  const latestVisible = vDates.length ? vDates[vDates.length - 1] : null;

  let html = '';
  rep.rows.forEach(r => {
    // 每个二级渠道 2 行：当日数据 / 月累计（平台·渠道·分类·客量纵向合并 2 行；客量拆两行不再合并）
    const cntDay = latestVisible ? (r.cntDays[latestVisible] || 0) : 0;
    const cntDayTitle = latestVisible ? `当日客量（${latestVisible.slice(5).replace('-','月')}日）` : '当日客量';
    html += `<tr class="ec-sub-today">
      <td class="cat" rowspan="2">${esc(r.platform)}</td>
      <td rowspan="2">${esc(r.l2)}</td>
      <td rowspan="2" class="chan-cell">${esc(r.channelClass || '—')}</td>
      <td rowspan="2" class="chan-cell">${esc(r.channelL1 || '—')}</td>
      <td class="type-cell">当日数据</td>
      <td class="num" data-col="cnt" data-val="${cntDay}" title="${cntDayTitle}">${cntDay}</td>
      ${vDates.map(d=>`<td class="num" data-idx="${idxOf.get(d)}" data-date="${d}" data-col="d:${d}" data-val="${r.days[d] || 0}">${fmt(r.days[d] || 0)}</td>`).join('')}
    </tr>
    <tr class="ec-sub-cum">
      <td class="type-cell">月累计</td>
      <td class="num" data-col="cnt" data-val="${r.cnt}" title="当月去重客量">${r.cnt}</td>
      ${vDates.map(d=>`<td class="num" data-idx="${idxOf.get(d)}" data-date="${d}" data-col="d:${d}" data-val="${r.cum[d] || 0}">${fmt(r.cum[d] || 0)}</td>`).join('')}
    </tr>`;
  });
  // 合计：同样 2 行（当日合计 / 累计合计）；data-total 用于「单列复制」时跳过合计，避免粘贴后重复汇总
  const tCntDay = latestVisible ? (rep.totalCntDays[latestVisible] || 0) : 0;
  html += `<tr class="row-orange" data-total="1">
      <td class="cat" colspan="5" rowspan="2">合计</td>
      <td class="num" data-col="cnt" data-val="${tCntDay}">${tCntDay}</td>
      ${vDates.map(d=>`<td class="num" data-idx="${idxOf.get(d)}" data-date="${d}" data-col="d:${d}" data-val="${rep.totalDays[d] || 0}">${fmt(rep.totalDays[d] || 0)}</td>`).join('')}
    </tr>
    <tr class="row-orange" data-total="1">
      <td class="num" data-col="cnt" data-val="${rep.totalCount}">${rep.totalCount}</td>
      ${vDates.map(d=>`<td class="num" data-idx="${idxOf.get(d)}" data-date="${d}" data-col="d:${d}" data-val="${rep.totalCum[d] || 0}">${fmt(rep.totalCum[d] || 0)}</td>`).join('')}
    </tr>`;
  body.innerHTML = html;

  // 列宽：前 6 列固定像素（内容更长时由 auto 布局自动撑开），每日列与 ③ 主表一致 90px
  table.querySelectorAll('colgroup').forEach(c=>c.remove());
  const cg = document.createElement('colgroup');
  const addCol = w => { const c = document.createElement('col'); c.style.width = w; cg.appendChild(c); };
  addCol('230px'); addCol('120px'); addCol('90px'); addCol('90px'); addCol('92px'); addCol('62px');
  vDates.forEach(()=>addCol(DAILY_COL_PX + 'px'));
  table.insertBefore(cg, table.firstChild);

  if(ecSubBatchMode) table.classList.add('batch-mode');
}

// 标题跨可见列：6 个固定列 + 可见日期列
function updateEcSubTitleColspan(){
  const table = $('ec-sub-table');
  if(!table) return;
  const visible = table.querySelectorAll('th.daily-col').length;   // 隐藏列已不渲染
  const titleTh = table.querySelector('thead tr:first-child th');
  if(titleTh) titleTh.colSpan = 6 + visible;
}

// ⑤ 页面 DOM/CSS 兜底：index.html 万一被 CDN/浏览器缓存成旧版本（缺按钮或宽度规则），
// 由 app.js（带 Date.now() 穿透）补齐，保证新旧 HTML 混载时渲染依旧正确。幂等，可重复调用。
function ensureEcSubDom(){
  const panel = $('panel-ec-sub');
  if(!panel) return;
  if(!document.getElementById('ec-sub-fallback-style')){
    const st = document.createElement('style');
    st.id = 'ec-sub-fallback-style';
    st.textContent = 'table.report#ec-sub-table{width:auto;}'
      + ' table.report .type-cell{background:var(--c-blue-sub);color:#002060;font-weight:600;font-size:12px;}'
      + ' table.report tr.ec-sub-cum td{background:#f7f9fd;}'
      + ' table.report tr.ec-sub-cum td.type-cell{background:#dbe5f7;}';
    document.head.appendChild(st);
  }
  if($('ec-sub-batch-toggle') && $('ec-sub-expand-all')) return;
  const bar = panel.querySelector('.row .row') || panel.querySelector('.row');
  if(!bar) return;
  if(!$('ec-sub-batch-toggle')){
    const b = document.createElement('button');
    b.className = 'btn'; b.id = 'ec-sub-batch-toggle';
    b.title = '批量选择要隐藏的日期列'; b.textContent = '☑ 批量';
    bar.insertBefore(b, bar.firstChild);
  }
  if(!$('ec-sub-expand-all')){
    const b = document.createElement('button');
    b.className = 'btn'; b.id = 'ec-sub-expand-all';
    b.title = '展开所有日期列'; b.textContent = '+展开';
    const batch = $('ec-sub-batch-toggle');
    bar.insertBefore(b, batch ? batch.nextSibling : bar.firstChild);
  }
}

// ⑤ 日期列交互（事件委托，只绑定一次）：单列收起、「☑ 批量」多选隐藏、「+展开」
// 隐藏列直接不渲染（避免幽灵列占宽产生空白），故所有操作都改 ecSubHidden 后重渲染
function bindEcSubControls(){
  ensureEcSubDom();
  const table = $('ec-sub-table');
  if(!table || table.dataset.bound) return;
  table.dataset.bound = '1';
  const head = $('ec-sub-head');
  const batchBtn = $('ec-sub-batch-toggle');
  const expandBtn = $('ec-sub-expand-all');
  const exitBatch = ()=>{
    ecSubBatchMode = false;
    table.classList.remove('batch-mode');
    if(batchBtn) batchBtn.textContent = '☑ 批量';
    table.querySelectorAll('th.daily-col.selected').forEach(th=>th.classList.remove('selected'));
  };
  // 至少保留一个日期列（否则「当日数据/月累计」两行失去意义）
  const applyHidden = (nextSet) => {
    const rep = computeEcSubReport();
    const total = (rep.dateList || []).length;
    if(total > 0 && nextSet.size >= total){
      showToast('至少保留一个日期列，请用「+展开」恢复', 'err', 2400);
      return false;
    }
    ecSubHidden = nextSet;
    renderEcSubReport();
    return true;
  };
  const curHidden = () => ecSubHiddenSet(computeEcSubReport());
  if(head){
    head.addEventListener('click', e => {
      const cbtn = e.target.closest('.copy-col');
      if(cbtn){ e.stopPropagation(); copyEcSubColumn(cbtn.dataset.col); return; }
      const th = e.target.closest('th.daily-col');
      if(!th) return;
      if(ecSubBatchMode){ th.classList.toggle('selected'); return; }   // 批量模式：点选（黄色高亮）
      const next = curHidden(); next.add(th.dataset.date);
      applyHidden(next);
    });
  }
  if(batchBtn){
    batchBtn.addEventListener('click', ()=>{
      if(ecSubBatchMode){
        const sel = [...table.querySelectorAll('th.daily-col.selected')].map(th => th.dataset.date);
        const next = curHidden();
        sel.forEach(d => next.add(d));
        if(sel.length && applyHidden(next)) showToast(`已隐藏 ${sel.length} 个日期列`, 'ok');
        exitBatch();
      } else {
        ecSubBatchMode = true;
        table.classList.add('batch-mode');
        batchBtn.textContent = '✕ 完成隐藏';
      }
    });
  }
  if(expandBtn){
    expandBtn.addEventListener('click', ()=>{
      if(ecSubBatchMode) exitBatch();
      ecSubHidden = new Set();     // 全部展开 → 记住「无隐藏」，重渲染不再回到默认
      renderEcSubReport();
    });
  }
}

// 复制 ⑤ 报表的某一整列（月累计 / 客量 / 某个日期）：读取单元格 data-val 原始数值
// 只取各平台明细行，跳过末尾「合计」行（避免粘贴到表格后重复汇总）
// 输出走 fmtPlain（无千分位、无强制小数），与 ③ 主表 copyDailyColumn 一致，粘贴进 Excel 即为数值
function copyEcSubColumn(colKey){
  const body = $('ec-sub-body');
  if(!body) return;
  const tds = [...body.querySelectorAll(`tr:not([data-total]) td[data-col="${colKey}"]`)];
  if(!tds.length){ showToast('该列暂无可复制数据', 'err', 1800); return; }
  const vals = tds.map(td => {
    const raw = td.dataset.val;
    return raw !== undefined ? fmtPlain(raw) : (td.textContent || '').replace(/,/g, '').trim();
  });
  copyTextToClipboard(vals.join('\n')).then(ok=>{
    showToast(ok ? `已复制该列 ${vals.length} 行` : '复制失败，请手动选择复制', ok ? 'ok' : 'err', 2200);
  });
}

// ⑤ 导出网格：平台名称 | 二级渠道 | 渠道分类 | 一级渠道 | 业绩类型 | 客量 | 每日列；每渠道 2 行（当日数据/月累计），末尾合计 2 行
// 客量拆两行（2026-10-26）：当日数据行=当日客量（最新可见日期→导出时取截止日，即最后一个日期列）、月累计行=当月去重客量
// 导出始终包含「全部」日期列（不跟随屏幕收起状态），便于后续分析
function buildEcSubGrid(){
  const rep = computeEcSubReport();
  const cutoff = state.cutoffDate ? new Date(state.cutoffDate) : new Date();
  const monthStr = (cutoff.getMonth()+1) + '月' + cutoff.getDate() + '日';
  const dateList = rep.dateList || [];
  const nCols = 6 + dateList.length;
  const cntDayDate = dateList.length ? dateList[dateList.length - 1] : null;   // 导出「当日客量」基准：截止日
  const rows = [];
  const merges = [];
  rows.push(Array.from({length:nCols}, () => null));
  rows[0][0] = cell('电商组各二级渠道业绩（初诊业绩，截止'+monthStr+'）', 's', 'title');
  merges.push({r:0, c:0, r2:0, c2:nCols-1});
  rows.push(Array.from({length:nCols}, () => null));
  rows[1][0] = cell('平台名称', 's', 'subhead');
  rows[1][1] = cell('二级渠道', 's', 'subhead');
  rows[1][2] = cell('渠道分类', 's', 'subhead');
  rows[1][3] = cell('一级渠道', 's', 'subhead');
  rows[1][4] = cell('业绩类型', 's', 'subhead');
  rows[1][5] = cell('客量\n(人)', 's', 'subhead');
  dateList.forEach((d,k)=>{ const dt = new Date(d); rows[1][6+k] = cell((dt.getMonth()+1)+'月'+dt.getDate()+'日', 's', 'subhead'); });

  rep.rows.forEach((row, ri) => {
    const r1 = 2 + ri * 2, r2 = r1 + 1;
    const a = Array.from({length:nCols}, () => null);   // 当日数据
    const b = Array.from({length:nCols}, () => null);   // 月累计
    a[0] = cell(row.platform, 's', 'cat');
    a[1] = cell(row.l2, 's', 'cat');
    a[2] = cell(row.channelClass || '—', 's', 'cat');
    a[3] = cell(row.channelL1 || '—', 's', 'cat');
    a[4] = cell('当日数据', 's', 'catsub');
    a[5] = cell(cntDayDate ? (row.cntDays[cntDayDate] || 0) : 0, 'n', 'target');
    b[4] = cell('月累计', 's', 'catsub');
    b[5] = cell(row.cnt, 'n', 'target');
    dateList.forEach((d,k)=>{
      a[6+k] = cell(row.days[d] || 0, 'n', 'num');
      b[6+k] = cell(row.cum[d] || 0, 'n', 'num');
    });
    rows.push(a); rows.push(b);
    merges.push({r:r1, c:0, r2:r2, c2:0});   // 平台名称 纵向合并
    merges.push({r:r1, c:1, r2:r2, c2:1});   // 二级渠道 纵向合并
    merges.push({r:r1, c:2, r2:r2, c2:2});   // 渠道分类 纵向合并
    merges.push({r:r1, c:3, r2:r2, c2:3});   // 一级渠道 纵向合并
  });
  // 合计：2 行（当日合计 / 累计合计），标签与客量纵向合并
  const tr = 2 + rep.rows.length * 2;
  const t1 = Array.from({length:nCols}, () => null);
  const t2 = Array.from({length:nCols}, () => null);
  t1[0] = cell('合计', 's', 'orange');
  t1[5] = cell(cntDayDate ? (rep.totalCntDays[cntDayDate] || 0) : 0, 'n', 'orange');
  t2[5] = cell(rep.totalCount, 'n', 'orange');
  dateList.forEach((d,k)=>{
    t1[6+k] = cell(rep.totalDays[d] || 0, 'n', 'orange');
    t2[6+k] = cell(rep.totalCum[d] || 0, 'n', 'orange');
  });
  rows.push(t1); rows.push(t2);
  merges.push({r:tr, c:0, r2:tr+1, c2:4});
  merges.push({r:tr, c:5, r2:tr+1, c2:5});

  const colWidths = [26, 18, 10, 10, 12, 8];
  dateList.forEach(()=>colWidths.push(11));
  return { rows, merges, colWidths, hiddenCols:[] };
}

async function exportEcSubXlsx(){
  await ensureXlsxStyled();
  const { rows, merges, colWidths } = buildEcSubGrid();
  buildStyledXlsx({ sheetName:'电商组二级渠道', rows, merges, colWidths }).then(blob=>{
    downloadBlob(blob, `电商组二级渠道业绩-${state.cutoffDate||todayIso()}.xlsx`);
    showToast('Excel已导出（含全部日期列）', 'ok');
  }).catch(e=>{
    console.error(e);
    showToast('导出失败：'+e.message, 'err');
  });
}

async function exportEcSubPng(){
  await ensureHtml2canvas();
  const table = $('ec-sub-table');
  if(!table) return;
  const holder = document.createElement('div');
  holder.style.cssText = 'position:absolute;left:-99999px;top:0;width:auto;background:#fff;z-index:-1;';
  const clone = table.cloneNode(true);
  clone.removeAttribute('id');                                 // 避免与线上表 id 冲突
  clone.querySelectorAll('.copy-col').forEach(n => n.remove()); // 图片不含「复制」按钮
  clone.style.width = (table.offsetWidth || 0) + 'px';
  clone.style.margin = '0';
  holder.appendChild(clone);
  document.body.appendChild(holder);
  html2canvas(clone, {backgroundColor:'#fff', scale:2, useCORS:true}).then(canvas=>{
    canvas.toBlob(blob=>{
      downloadBlob(blob, `电商组二级渠道业绩-${state.cutoffDate||todayIso()}.png`);
      showToast('图片已导出', 'ok');
    });
  }).catch(e=>{
    console.error(e);
    showToast('导出失败：'+e.message, 'err');
  }).finally(()=>{ holder.remove(); });
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
      if (th.classList.contains('col-hidden') || th.id === 'batch-toggle' || th.id === 'expand-all') {
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

  // 调试横幅已下线（2026-09-29），不再展示；#debug-banner 默认隐藏
  if(vabUniqueCount() > 0){
    $('vab-meta').textContent = `✔ 本地VAB（${vabUniqueCount()} 人）`;
    $('vab-meta').className = 'file-meta';
  }

  $('file-orders').addEventListener('change', e=>onFileOrders(e.target));
  // 「合并历史数据」勾选框：勾选本身不导入任何数据，仍需点击「上传订单表」选择文件。
  // 这里给出显式提示，避免用户误以为「勾了就已经把 10 月数据并进来了」。
  const mmBox = $('merge-mode');
  if(mmBox) mmBox.addEventListener('change', () => {
    if(mmBox.checked) showToast('已勾选「合并历史数据」——但这不会自动导入数据；仍需点击「📤 上传订单表」并选择 10 月的订单文件（重叠日期会覆盖，其它日期保留）', 'warn', 7000);
  });
  $('file-vab').addEventListener('change', e=>onFileVab(e.target));
  // 近一年消费明细导入 / 清除（2026-10-16 新增：一年内无消费→初诊 新客规则数据源）
  $('file-history').addEventListener('change', e=>onFileHistory(e.target));
  $('clear-hist').addEventListener('click', clearHistData);
  refreshHistMeta(); // 本地恢复后的历史明细状态文案
  $('edit-vab').addEventListener('click', toggleVabList);
  $('recompute').addEventListener('click', recompute);
  $('cutoff-date').addEventListener('change', ()=>{ persistOrders(); recompute(); }); // 切换截止日后持久化并重算月累计/每日
  $('clear-manual').addEventListener('click', clearManualOverrides);
  $('clear-vab').addEventListener('click', clearVabList);
  $('clear-all').addEventListener('click', clearAllData);
  $('clear-date').addEventListener('click', clearOrdersByDate);
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
  // ⑤ 电商组二级渠道业绩导出
  const ecSubXlsxBtn = $('export-ec-sub-xlsx');
  if(ecSubXlsxBtn) ecSubXlsxBtn.addEventListener('click', exportEcSubXlsx);
  const ecSubPngBtn = $('export-ec-sub-png');
  if(ecSubPngBtn) ecSubPngBtn.addEventListener('click', exportEcSubPng);
  bindEcSubControls();   // ⑤ 日期列：单列收起 /「☑ 批量」/「+展开」
  // 「批量 / 展开」报表独立按钮（已移出表头）：仅绑定一次；④ 门店业绩同款
  bindBatchExpandButtons($('report-table'));
  bindStoreBatchExpandButtons($('store-table'));

  // 不再自动注入示例数据：打开即为空，仅显示用户上传的真实数据（VAB 已从本机恢复）
  refreshVabStatus(); // 还原后统一刷新 VAB 状态（已保存默认 vs 未上传）

  // 历史订单已在脚本加载时从本地恢复（restoreOrders）；若有效则恢复截止日期、自动重算月累计/每日数据。
  // 该渲染/重算逻辑抽成 runRestoredRecompute，供 initApp 与 IndexedDB 异步恢复回调（_postRestore）共用。
  _postRestore = runRestoredRecompute;
  runRestoredRecompute();

  // 卸载前兜底刷写：确保「不主动点击清除全部数据」时，任何代码路径产生的订单变更都已落盘。
  // beforeunload 内的 localStorage.setItem 为同步写，刷新/关页/切换平台页签销毁 iframe 前都能可靠落盘。
  window.addEventListener('beforeunload', persistOrders);

  // 调试条已按需求下线（2026-09-29），仅保持横幅隐藏
  refreshDebug();

  // ---- 云端：登录/改密/退出已统一至整站登录页与账号面板，页内不再提供登录入口 ----
  // 首次进入：已登录则拉取云端订单；未登录则把状态条变为登录入口（不强制弹窗打断）
  syncFromCloud();
  // 手动云端操作按钮（2026-10-24）：给用户显式控制权，绕过自动裁决兜底
  const cloudRestoreBtn = $('cloud-restore');
  if(cloudRestoreBtn) cloudRestoreBtn.addEventListener('click', () => {
    if(!_cloudUser){ showToast('请先登录云端', 'warn'); return; }
    showToast('正在从云端恢复…', 'info');
    syncFromCloud({ forcePull: true, silent: false });
  });
  const cloudPushBtn = $('cloud-push');
  if(cloudPushBtn) cloudPushBtn.addEventListener('click', () => {
    if(!_cloudUser){ showToast('请先登录云端', 'warn'); return; }
    const hasLocal = (Array.isArray(state.orders) && state.orders.length > 0)
      || (state.histOrders||[]).length > 0 || (state.vabList||[]).length > 0;
    if(!hasLocal){ showToast('本机暂无数据可推送', 'warn'); return; }
    showToast('正在推送到云端…', 'info');
    pushLocalSnapshotToCloud();
  });
  // SDK 多源回退晚于首帧加载完成时：加载成功→重试云端同步刷新状态条；全部失败→明确提示
  document.addEventListener('cloud-sdk-ready', () => { if(!_cloudUser) syncFromCloud(); });
  document.addEventListener('cloud-sdk-failed', () => { if(!_cloudUser) setCloudStatus('云端组件加载失败 · 请检查网络后刷新重试（数据仍保存在本机）', 'err'); });
  // 跨设备近实时同步（2026-10-16）：同一账号在另一台电脑/浏览器改动后，本机定时/切回页面时轻量重拉一次。
  // 保护：本地存在待写入变更（防抖未触发）或正在应用云端快照时跳过，绝不用云端覆盖尚未上传的本地改动。
  const cloudPoll = () => {
    if(document.hidden) return;
    if(_cloudSaveTimer) return;
    if(_applyingCloud) return;
    syncFromCloud({ silent: true });
  };
  setInterval(cloudPoll, 30000);
  document.addEventListener('visibilitychange', () => { if(!document.hidden) cloudPoll(); });
}

// 历史订单恢复后的渲染与重算（localStorage 同步恢复 / IndexedDB 异步恢复共用）
// opts.silent=true：跨设备定时同步等「后台刷新」场景，不再弹「已从本地恢复 N 条」提示，避免打扰
function runRestoredRecompute(opts){
  const silent = !!(opts && opts.silent);
  if(state.orders.length > 0){
    if(state.cutoffDate) $('cutoff-date').value = state.cutoffDate;
    // 防止新导入订单的 _uid 与已恢复订单冲突：把序列号抬到已恢复最大值之上
    const maxUid = state.orders.reduce((m,r)=>{ const n = r._uid && /^ord_(\d+)$/.exec(r._uid); return n ? Math.max(m, +n[1]) : m; }, 0);
    if(maxUid > _orderUidSeq) _orderUidSeq = maxUid;
    // 关键性能修复：先即时渲染空表骨架（首帧），再在浏览器空闲时做重计算，
    // 避免点击菜单后因“解析+重算+整表重渲染”全阻塞主线程而长时间白屏/卡死
    renderReport();
    showOrdersSaved(true);
    const runRecompute = () => {
      recompute();
      if(silent) return;
      showToast(`已从本地恢复 ${state.orders.length} 条历史订单，月累计/每日数据已自动刷新`, 'ok');
      // 若已存数据的最新月份早于当前月，说明存在「比存储更新的数据未落盘」的情况
      // （例如上一会话看过 10 月但未成功导入/保存）。此时明确提示，避免用户误以为数据被系统吞掉。
      try{
        const nowMonth = todayIso().slice(0,7);                 // 形如 2026-10
        let maxMonth = '';
        for(const r of state.orders){
          if(!r.payTime) continue;
          const m = fmtMonth(r.payTime);
          if(m > maxMonth) maxMonth = m;
        }
        if(maxMonth && maxMonth < nowMonth){
          const [y, mo] = nowMonth.split('-');
          setTimeout(()=>showToast(
            `⚠ 已保存数据最新为 ${maxMonth.replace('-','年')}月，少于当前 ${y}年${Number(mo)}月。若你此前看过 ${y}年${Number(mo)}月数据，说明那份数据未成功保存——请重新上传订单表（勾选「合并历史数据」可与 ${maxMonth.replace('-','年')}月数据合并，不重复计数）`,
            'warn', 9000), 900);
        }
      }catch(_){}
    };
    if(window.requestIdleCallback) window.requestIdleCallback(runRecompute, { timeout: 400 });
    else setTimeout(runRecompute, 30);
  } else {
    renderReport();
  }
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
