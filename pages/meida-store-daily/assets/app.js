'use strict';
/* =====================================================================
 * 美大店铺日数据（meida-store-daily）
 * ---------------------------------------------------------------------
 * 需求（2026-10-28）：
 *   「美大店铺日数据 按照此文件格式创建报表 需要支持数据导入 日期筛选
 *     业绩以及客量的显示 从日业绩报表 电商组二级渠道业绩（初诊）平移」
 *
 * 两条数据来源
 *   1) 订单表（.xlsx）—— 与「日业绩报表」共用同一份持久化数据（localStorage
 *      dailyPerformance_orders + IndexedDB），两页天然一致，只导入一次即可。
 *      → 自动算出「初诊业绩 / 初诊成交客量 / 有效客资人数 / 新增VAB客量 / 客单价」
 *   2) 七大店铺数据总表（.xlsx）—— 每个 sheet 名即日期（如 10.23），取 A4:AH17
 *      的 7 店 × 2 行数据，填充评分 / 美团漏斗 / 投产比 / 好评等人工指标。
 *   3) 平台日报三表（可一次多选，按表头自动识别类型）—— 覆盖「美团数据」「回复时长/回复率」：
 *        客流数据    → 曝光人数(O)、访问人数(P)
 *        在线咨询数据 → 昨日咨客平均回复时长(AG)、30s内回复率(AH)
 *        门店交易数据 → 下单人数(Q)、下单券数(R)、下单金(S)、核销人数(T)、核销券数(U)、总核销金额(V)
 *      当日取该日行；月累计 = 当月 1 日~所选日逐日累加（回复时长/回复率按「在线咨询人数」加权）。
 *      三表存在时优先于七大店铺总表的同名列；缺失的格子仍回落总表。
 *
 * 口径（与 ③④⑤ 的「电商组二级渠道业绩（初诊）」严格同源，改一处必须同步另一处）
 *   渠道范围 isEcSubScope()：营销部 + 一级渠道∈{电商组,客户转介绍,企业团检}
 *     ∪ 营销部 + 品牌组 + 二级含「线上400」∪ 推荐→营销部（二级不含自然到店）
 *     ∪ 手动报表归属=电商组
 *   初诊门槛 buildMktFirstKeys()：30 天窗口内营销部初诊客户集合
 *     （规则A 初诊类型 / 规则B 非初诊但近一年历史明细无记录 / 规则C 空白类型且 365 天无消费）
 *   「本月累计」= 当月 1 日至所选日期的逐日累加（不取源表快照）。
 * ===================================================================== */

/* ---------------- 基础工具 ---------------- */
const $ = (id) => document.getElementById(id);
const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const pad2 = (n) => String(n).padStart(2, '0');
const fmtDate = (d) => (d ? `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}` : '');
const iso2date = (s) => { const m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/); return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null; };
const dayLabel = (s) => { const d = iso2date(s); return d ? `${d.getMonth() + 1}月${d.getDate()}日` : s; };
const isNum = (v) => typeof v === 'number' && isFinite(v);
const money = (n, d = 2) => (isNum(n) ? Number(n).toLocaleString('zh-CN', { minimumFractionDigits: d, maximumFractionDigits: d }) : '—');
const int0 = (n) => (isNum(n) ? Math.round(n).toLocaleString('zh-CN') : '—');
const $id = (s) => document.querySelector(s);

let _toastTimer = 0;
function toast(msg, kind, ms) {
  const el = $('toast');
  if (!el) return;
  el.textContent = msg;
  el.className = 'toast show' + (kind ? ' ' + kind : '');
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(() => { el.className = 'toast'; }, ms || 3000);
}

/* ---------------- 7 大店铺（与 ⑤ 的 EC_SUB_PLATFORMS 一一对应） ----------------
 * match = 平台导出文件里的门店名匹配规则（美团/大众点评导出名与本表名不完全一致，
 *   例：本表「医疗保健中心（未央分院）」，平台导出为「长寿养护保健中心（未央分院）」）。
 * 顺序有意义：先具体的，后宽松的。
 */
const STORES = [
  { key: '美大-高端体检', name: '海蓓国际医疗·高端体检中心', pm: '李群', op: '段聪', match: /高端体检/ },
  { key: '美大-国际门诊', name: '海蓓国际医疗·高新中心医院', pm: '李群', op: '段聪', match: /高新/ },
  { key: '美大-口腔保健', name: '海蓓国际医疗·口腔保健美学中心', pm: '杜巧娟', op: '（待招）', match: /口腔/ },
  { key: '美大-新妇科·抗衰', name: '海蓓国际医疗·女性健康中心', pm: '杜巧娟', op: '（待招）', match: /女性健康/ },
  { key: '美大-新妇科·私密', name: '海蓓Anna·皮肤私密养护中心', pm: '杜巧娟', op: '（待招）', match: /私密|anna/i },
  { key: '美大-北郊保健', name: '海蓓·私享保健中心', pm: '杜巧娟', op: '杜巧娟', match: /私享|北郊保健/ },
  { key: '美大-北郊医疗', name: '海蓓国际医疗·医疗保健中心（未央分院）', pm: '杜巧娟', op: '杜巧娟', match: /未央|北郊医疗/ },
];
const STORE_BY_NAME = new Map(STORES.map((s) => [s.name, s]));
const STORE_KEYS = STORES.map((s) => s.key);

// 平台导出门店名 → 本表店铺（先精确，再按 match 关键字容错）
function resolveStoreByName(name) {
  const s = String(name == null ? '' : name).trim();
  if (!s) return null;
  if (STORE_BY_NAME.has(s)) return STORE_BY_NAME.get(s);
  for (const st of STORES) {
    if (st.match && st.match.test(s)) return st;
  }
  // 兜底：本表名前 8 字前缀匹配（历史行为）
  return STORES.find((x) => s.indexOf(x.name.slice(0, 8)) === 0) || null;
}

/* ---------------- 34 列表头（严格对应原 Excel A..AH） ----------------
 * g   = 第 2 行分组id
 * zone= 行底色分区：''(B–N) / 'b'(O–AB) / 'c'(AC–AH)
 * auto= 该列由订单自动计算
 * sch = 源表列字母（解析七大店铺总表时用）
 */
const COLS = [
  { k: 'store', label: '平台名称', g: 1, w: 196, sch: 'A', type: 'text' },
  { k: 'pm', label: '项目经理', g: 2, w: 76, sch: 'B', type: 'text' },
  { k: 'op', label: '平台运营', g: 2, w: 76, sch: 'C', type: 'text' },
  { k: 'mt', label: '美团', g: 3, w: 64, sch: 'D', type: 'score' },
  { k: 'dp', label: '大众点评', g: 3, w: 76, sch: 'E', type: 'scoreText' },
  { k: 'rank', label: '榜单', g: 3, w: 84, sch: 'F', type: 'text' },
  { k: 'period', label: '', g: 4, w: 72, sch: 'G', type: 'text' },
  { k: 'tgtAmt', label: '初诊目标（T30）', g: 5, w: 92, sch: 'H', type: 'int' },
  { k: 'actAmt', label: '初诊业绩（T30）', g: 5, w: 94, sch: 'I', type: 'money', auto: 1 },
  { k: 'aov', label: '客单价', g: 5, w: 76, sch: 'J', type: 'money0', auto: 1 },
  { k: 'lead', label: '有效客资人数', g: 6, w: 94, sch: 'K', type: 'int', auto: 1 },
  { k: 'tgtVis', label: '客量目标（T30）', g: 6, w: 88, sch: 'L', type: 'int' },
  { k: 'firstVis', label: '初诊成交客量（T30）', g: 6, w: 96, sch: 'M', type: 'int', auto: 1 },
  { k: 'newVab', label: '新增VAB客量', g: 6, w: 92, sch: 'N', type: 'int', auto: 1 },
  { k: 'expo', label: '曝光人数', g: 7, w: 68, sch: 'O', type: 'int', zone: 'b' },
  { k: 'visit', label: '访问人数', g: 7, w: 68, sch: 'P', type: 'int', zone: 'b' },
  { k: 'orderU', label: '下单人数', g: 7, w: 68, sch: 'Q', type: 'int', zone: 'b' },
  { k: 'orderC', label: '下单券数', g: 7, w: 68, sch: 'R', type: 'int', zone: 'b' },
  { k: 'orderAmt', label: '下单金（元）', g: 7, w: 86, sch: 'S', type: 'money', zone: 'b' },
  { k: 'woU', label: '核销人数', g: 7, w: 68, sch: 'T', type: 'int', zone: 'b' },
  { k: 'woC', label: '核销券数', g: 7, w: 68, sch: 'U', type: 'int', zone: 'b' },
  { k: 'woAmt', label: '总核销金额（元）', g: 7, w: 94, sch: 'V', type: 'money', zone: 'b' },
  { k: 'vcC', label: '代金券张数（非新客）', g: 7, w: 88, sch: 'W', type: 'int', zone: 'b' },
  { k: 'vcAmt', label: '代金券金额（非新客）', g: 7, w: 94, sch: 'X', type: 'money', zone: 'b' },
  { k: 'realWo', label: '实际核销（元）', g: 7, w: 94, sch: 'Y', type: 'money', zone: 'b' },
  // Z（推广费）归入「投产比」分组（推广费 / 目标 / 实际 三列同属投产比，与原始模板表头一致）
  { k: 'promoFee', label: '推广费（元）', g: 8, w: 86, sch: 'Z', type: 'money', zone: 'b' },
  { k: 'roiTgt', label: '目标', g: 8, w: 64, sch: 'AA', type: 'text', zone: 'c' },
  { k: 'roiAct', label: '实际', g: 8, w: 84, sch: 'AB', type: 'text', zone: 'c' },
  { k: 'revTgt', label: '好评目标', g: 9, w: 70, sch: 'AC', type: 'int', zone: 'c' },
  { k: 'revCnt', label: '月累计好评数量', g: 9, w: 88, sch: 'AD', type: 'int', zone: 'c' },
  { k: 'revRate', label: '好评完成率', g: 9, w: 78, sch: 'AE', type: 'pct', zone: 'c' },
  { k: 'qrScan', label: '优惠码扫码量（近7日）', g: 9, w: 116, sch: 'AF', type: 'int', zone: 'c' },
  { k: 'replySec', label: '昨日咨客平均回复时长（秒）', g: 9, w: 150, sch: 'AG', type: 'sec', zone: 'c' },
  { k: 'reply30', label: '30s内回复率', g: 9, w: 84, sch: 'AH', type: 'pct', zone: 'c' },
];
// 与 Excel 的差异：Excel 把「好评目标」的上级留空（AC2 未合并），此处统一归入「美团榜单维护及客资转化」分组，视觉更完整。
const GROUPS = [
  { id: 1, label: '美大店铺', cls: 'g1' },
  { id: 2, label: '责任人', cls: 'g1' },
  { id: 3, label: '店铺评分/星级/榜单', cls: 'g2' },
  { id: 4, label: '统计口径', cls: 'g3', rowspan: 2 },
  { id: 5, label: '业绩', cls: 'g4' },
  { id: 6, label: '客量', cls: 'g2' },
  { id: 7, label: '美团数据（曝光 / 下单 / 核销）填前一天数据', cls: 'g5' },
  { id: 8, label: '投产比（按T30初诊业绩核算）', cls: 'g6' },
  { id: 9, label: '美团榜单维护及客资转化（填前一天数据）', cls: 'g7' },
];

// 合计行参与求和的列（非派生/非比率列）
const SUM_KEYS = ['tgtAmt', 'actAmt', 'lead', 'tgtVis', 'firstVis', 'newVab',
  'expo', 'visit', 'orderU', 'orderC', 'orderAmt', 'woU', 'woC', 'woAmt', 'vcC', 'vcAmt', 'realWo', 'promoFee', 'revTgt', 'revCnt'];
// 店级共享字段：当日数据与本月累计两行取值相同，渲染时纵向合并（rowspan=2），与原始模板表一致
const MERGED_FIELDS = ['store', 'pm', 'op', 'mt', 'dp', 'rank'];
/* 「投产比」+「好评」两组里的 5 列：原始模板表里同店两行是纵向合并的（目标 1:4 / 实际 1:0.47 …），
   只有当两行取值确实不同时才保留两行（如好评目标 20 / 40）。故按「值相同 或 其中一行为空」动态合并。
   推广费（元）在模板里逐行独立，不参与。 */
const COND_MERGE_FIELDS = ['roiTgt', 'roiAct', 'revTgt', 'revCnt', 'revRate'];
const AUTO_KEYS = COLS.filter((c) => c.auto).map((c) => c.k);

/* ---------------- 持久化键 ---------------- */
const LS = {
  orders: 'dailyPerformance_orders',   // 与日业绩报表共用
  hist: 'dailyPerformance_hist',       // 与日业绩报表共用
  vab: 'dailyPerformance_vab',         // 与日业绩报表共用
  rows: 'meidaStore_rows',             // 七大店铺总表解析结果
  daily: 'meidaStore_daily',           // 平台三表（客流/咨询/交易）逐日明细
  promo: 'meidaStore_promo',          // 门店报表（推广费，按 日期×门店）
  manual: 'meidaStore_manual',         // 本页手工覆盖
  names: 'meidaStore_names',           // 店铺名手工覆盖 { storeKey: 名称 }
  meta: 'meidaStore_meta',             // 导入时间/年份等
};
const IDB_NAME = 'dailyPerformance';
const IDB_STORE = 'kv';

/* ---------------- 状态 ---------------- */
const state = {
  orders: [],          // 订单（含 Date 化的 payTime）
  histOrders: [],      // 历史消费明细（仅用于规则B）
  vabSet: new Set(),   // VAB 会员规范化集合
  rows: {},            // { 'YYYY-MM-DD': { storeKey: { pm,op,mt,dp,rank, today:{…}, cum:{…} } } }
  daily: {},           // { 'YYYY-MM-DD': { storeKey: { expo,visit,orderU,orderC,orderAmt,woU,woC,woAmt,replySec,reply30,chatCnt } } }
  promo: {},           // { 'YYYY-MM-DD': { storeKey: 推广费（元） } }
  manual: {},          // { 'date|storeKey|field': value }
  names: {},           // { storeKey: 店铺名 } 手工改名（不随日期变化）
  date: '',            // 当前报表日期 'YYYY-MM-DD'
  period: 'both',      // both | today | cum
  source: 'auto',      // auto | sheet
  editMode: false,     // 编辑模式（点「编辑」开启：单击任意格可改）
  ordersTs: 0,
  meta: {},
};

/* =====================================================================
 * 一、口径函数（镜像自 daily-performance-chuzhen/assets/app.js，勿单独修改）
 * ===================================================================== */
function chanHas(r, val) {
  return (r.channelClass || '').trim() === val || (r.channelL1 || '').trim() === val || (r.channelL2 || '').trim() === val;
}
function isRecToYingxiao(r) {
  return (r.channelClass || '').trim() === '推荐' && (r.channelL1 || '').includes('营销部');
}
function normMemberKey(no) {
  const s = String(no == null ? '' : no).trim().toLowerCase();
  const d = s.replace(/\D/g, '');
  return d ? d.replace(/^0+/, '') : '';
}
function mktFirstClientKey(r) {
  const nk = normMemberKey(r.memberNo);
  return nk ? nk : 'name:' + String(r.name || '').trim();
}
function isMktChuzhenOrder(r) {
  return (chanHas(r, '营销部') || isRecToYingxiao(r)) && (r.visitType || '').trim() === '初诊';
}
const DEPT_OPTIONS = ['品牌组-其他', '活动组-活动支持', '电商组', '新媒体组', '短视频组', '品牌组-自然到店', '商务部', '其他'];
// 「电商组」渠道范围（= ③ 主表「电商组」行 r3 ∪ r3b ∪ r3c）
function isEcSubScope(r) {
  const ovk = String(r.deptOverride || '').trim();
  if (DEPT_OPTIONS.includes(ovk)) return ovk === '电商组';
  const cc = (r.channelClass || '').trim();
  const l1 = (r.channelL1 || '').trim();
  const l2 = (r.channelL2 || '').trim();
  if (cc === '营销部' && ['电商组', '客户转介绍', '企业团检'].includes(l1)) return true;
  if (cc === '营销部' && l1 === '品牌组' && l2.indexOf('线上400') >= 0) return true;
  if (isRecToYingxiao(r) && l2.indexOf('自然到店') < 0) return true;
  return false;
}
// 近一年历史明细的客户键集合（未导入历史明细时为空 → 规则B 自动停用，规则C 不受影响）
let _histCache = { sig: '', date: '', set: new Set() };
function histKeysFor(cutoff) {
  const d = cutoff ? fmtDate(cutoff) : '';
  const sig = state.histOrders.length + '|' + (state.meta.histTs || 0);
  if (_histCache.sig !== sig || _histCache.date !== d) {
    const s = new Set();
    if (d && state.histOrders.length) {
      const start = fmtDate(new Date(cutoff.getTime() - 366 * 86400000));
      state.histOrders.forEach((r) => {
        if (!r.payTime) return;
        const dd = fmtDate(r.payTime);
        if (dd < start || dd > d) return;
        const k = mktFirstClientKey(r);
        if (k && k !== 'name:') s.add(k);
      });
    }
    _histCache = { sig, date: d, set: s };
  }
  return _histCache.set;
}
// 30 天内「营销部初诊客户」集合（逐日口径：每个报表日期各自一个 cutoff）
function buildMktFirstKeys(list, cutoff) {
  const s = new Set();
  if (!cutoff || isNaN(cutoff.getTime())) return s;
  const end = fmtDate(cutoff);
  const start = fmtDate(new Date(cutoff.getTime() - 29 * 86400000));
  const histKeys = histKeysFor(cutoff);
  const winStart365 = fmtDate(new Date(cutoff.getTime() - 366 * 86400000));
  const firstSeen = new Map();
  (list || []).forEach((r) => {
    if (!r.payTime) return;
    const d = fmtDate(r.payTime);
    if (d < winStart365 || d > end) return;
    const k = mktFirstClientKey(r);
    if (!k || k === 'name:') return;
    const prev = firstSeen.get(k);
    if (prev === undefined || d < prev) firstSeen.set(k, d);
  });
  (list || []).forEach((r) => {
    if (!r.payTime) return;
    const d = fmtDate(r.payTime);
    const k = mktFirstClientKey(r);
    if (!k || k === 'name:') return;
    if (d >= start && d <= end && isMktChuzhenOrder(r)) {
      s.add(k);
    } else if (d >= start && d <= end && (chanHas(r, '营销部') || isRecToYingxiao(r))) {
      const vt = (r.visitType || '').trim();
      const hasEarlier = firstSeen.get(k) !== undefined && firstSeen.get(k) < d;
      const hasHist = histKeys.size > 0 && histKeys.has(k);
      if (vt === '') {
        if (!hasEarlier && !hasHist) s.add(k);
      } else if (histKeys.size && !histKeys.has(k)) {
        s.add(k);
      }
    }
  });
  return s;
}

/* =====================================================================
 * 二、数据读取（共用日业绩报表的存储）
 * ===================================================================== */
function openIDB() {
  return new Promise((resolve, reject) => {
    if (typeof indexedDB === 'undefined') return reject(new Error('no-idb'));
    let req;
    try { req = indexedDB.open(IDB_NAME, 1); } catch (e) { return reject(e); }
    req.onupgradeneeded = () => { try { if (!req.result.objectStoreNames.contains(IDB_STORE)) req.result.createObjectStore(IDB_STORE); } catch (_) {} };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error || new Error('idb-open-fail'));
  });
}
function idbGet(key) {
  return openIDB().then((db) => new Promise((resolve, reject) => {
    try {
      const r = db.transaction(IDB_STORE, 'readonly').objectStore(IDB_STORE).get(key);
      r.onsuccess = () => resolve(r.result);
      r.onerror = () => reject(r.error);
    } catch (e) { reject(e); }
  })).catch(() => null);
}
function idbSet(key, value) {
  return openIDB().then((db) => new Promise((resolve) => {
    try {
      const tx = db.transaction(IDB_STORE, 'readwrite');
      tx.objectStore(IDB_STORE).put(value, key);
      tx.oncomplete = () => resolve(true);
      tx.onerror = () => resolve(false);
    } catch (e) { resolve(false); }
  })).catch(() => false);
}
function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
function lsSet(k, v) { try { localStorage.setItem(k, v); return true; } catch (e) { return false; } }

function reviveOrders(arr) {
  return (arr || []).map((r) => {
    if (r && typeof r.payTime === 'string') r.payTime = new Date(r.payTime);
    return r;
  }).filter((r) => r && r.payTime);
}

// 读取订单：先 localStorage（快），再 IndexedDB（大容量兜底），取行数更多者
async function loadSharedOrders() {
  let best = [];
  const raw = lsGet(LS.orders);
  if (raw) {
    try { const o = JSON.parse(raw); if (Array.isArray(o.orders)) best = reviveOrders(o.orders); state.ordersTs = Number(o.ts) || 0; } catch (e) {}
  }
  const idbRaw = await idbGet(LS.orders);
  if (typeof idbRaw === 'string') {
    try {
      const o = JSON.parse(idbRaw);
      if (Array.isArray(o.orders) && o.orders.length > best.length) { best = reviveOrders(o.orders); state.ordersTs = Number(o.ts) || 0; }
    } catch (e) {}
  }
  state.orders = best;
  return best;
}
async function loadSharedVab() {
  let raw = lsGet(LS.vab);
  if (!raw) raw = await idbGet(LS.vab);
  if (typeof raw !== 'string') return;
  try {
    const o = JSON.parse(raw);
    if (Array.isArray(o.set)) state.vabSet = new Set(o.set.map((x) => String(x)));
  } catch (e) {}
}
async function loadSharedHist() {
  let raw = lsGet(LS.hist);
  if (!raw) raw = await idbGet(LS.hist);
  if (typeof raw !== 'string') return;
  try {
    const o = JSON.parse(raw);
    if (Array.isArray(o.histOrders)) {
      state.histOrders = o.histOrders.map((r) => {
        if (r && typeof r.payTime === 'string') r.payTime = new Date(r.payTime);
        return r;
      }).filter((r) => r && r.payTime);
      state.meta.histTs = Number(o.ts) || 0;
    }
  } catch (e) {}
}

// 写入订单（与日业绩报表同格式同键，写后需回到该页刷新才会重算 ③④⑤）
function persistSharedOrders(orders) {
  const payload = JSON.stringify({ orders, cutoff: state.date, ts: Date.now() });
  lsSet(LS.orders, payload);
  idbSet(LS.orders, payload);
}

/* =====================================================================
 * 三、订单表导入（字段解析镜像自日业绩报表 normalizeOrder）
 * ===================================================================== */
function pickField(row, candidates) {
  for (const k of Object.keys(row)) {
    for (const c of candidates) {
      if (String(k).replace(/\s|（实时|（\u5b9e\u65f6）/g, '').includes(c)) return row[k];
    }
  }
  return '';
}
function parseDateLoose(v) {
  if (!v) return null;
  if (v instanceof Date) return v;
  const s = String(v).trim();
  let m = s.match(/(\d{4})年(\d{1,2})月(\d{1,2})日[-\sT]?(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})[-\sT](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
  m = s.match(/(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const dt = new Date(s);
  return isNaN(dt) ? null : dt;
}
function normalizeOrder(r) {
  const s = (v) => String(v == null ? '' : v).trim();
  return {
    name: s(pickField(r, ['客户姓名', '客户'])),
    memberNo: s(pickField(r, ['会员号', '会员编号', '会员ID', '会员卡号', '卡号', 'VAB会员号', 'member_no'])),
    developer: s(pickField(r, ['客户开发人'])),
    consultant: s(pickField(r, ['开单健康管家', '开单咨询师', '咨询师', '服务顾问', '开单人', '主诊咨询师'])),
    channelClass: s(pickField(r, ['客户来源渠道分类', '渠道分类'])),
    channelL1: s(pickField(r, ['客户来源一级渠道', '一级渠道'])),
    channelL2: s(pickField(r, ['客户来源二级渠道', '二级渠道'])),
    channelL3: s(pickField(r, ['客户来源三级渠道', '三级渠道', '渠道L3', '客户来源三级', '三级来源'])),
    cashPay: Number(pickField(r, ['现款支付'])) || 0,
    assetDeduct: Number(pickField(r, ['收款金额-资产抵扣', '资产抵扣'])) || 0,
    visitType: s(pickField(r, ['到访类型'])),
    payTime: parseDateLoose(pickField(r, ['收款时间'])),
    receiptNo: s(pickField(r, ['收款单号'])),
    cashier: s(pickField(r, ['收款员工'])),
    payType: s(pickField(r, ['收款类型'])),
    payMethod: s(pickField(r, ['收款详情汇总', '收款方式', '详情'])),
  };
}
let _xlsxLoading = null;
function ensureXlsx() {
  if (typeof XLSX !== 'undefined') return Promise.resolve();
  if (_xlsxLoading) return _xlsxLoading;
  _xlsxLoading = new Promise((resolve, reject) => {
    const s = document.createElement('script');
    s.src = 'assets/vendor/xlsx.full.min.js';
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('xlsx 库加载失败'));
    document.head.appendChild(s);
  });
  return _xlsxLoading;
}
async function readSheetObjects(file) {
  await ensureXlsx();
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: 'array', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}
// 按「收款日期」合并：本次导入覆盖同名日期，其它日期保留（与日业绩报表 mergeOrdersByDate 同意图）
function mergeByDate(existing, incoming) {
  const dates = new Set(incoming.map((r) => fmtDate(r.payTime)));
  const kept = existing.filter((r) => r && r.payTime && !dates.has(fmtDate(r.payTime)));
  let replaced = 0;
  existing.forEach((r) => { if (r && r.payTime && dates.has(fmtDate(r.payTime))) replaced++; });
  return { orders: kept.concat(incoming), replaced, coveredDates: dates.size };
}

async function onFileOrders(file) {
  if (!file) return;
  try {
    const parsed = await readSheetObjects(file);
    const rows = parsed.map(normalizeOrder).filter((r) => r.payTime);
    if (parsed.length && !rows.length) {
      toast('导入失败：已读到 ' + parsed.length + ' 行，但「收款时间」全部无法解析，请确认表头含该列', 'err', 8000);
      return;
    }
    const existed = state.orders.length > 0;
    let msg;
    if (existed) {
      const { orders, replaced, coveredDates } = mergeByDate(state.orders, rows);
      state.orders = orders;
      msg = `已合并 ${rows.length} 条（覆盖 ${coveredDates} 个日期共 ${replaced} 条历史订单，保留其它日期）`;
    } else {
      state.orders = rows;
      msg = `订单表导入成功，共 ${rows.length} 条`;
    }
    persistSharedOrders(state.orders);
    syncDateRange();
    autoPickDate();
    render();
    toast(msg, 'ok', 4200);
  } catch (e) {
    console.error(e);
    toast('订单表解析失败：' + e.message, 'err', 6000);
  }
}

/* =====================================================================
 * 四、七大店铺数据总表解析
 * ===================================================================== */
function cleanCell(v, type) {
  if (v == null || v === '') return null;
  if (v instanceof Date) {
    // 个别单元格被录成时间（如 10.26 的 AB12 = 01:13），按 HH:MM 文本保留
    return pad2(v.getHours()) + ':' + pad2(v.getMinutes());
  }
  if (typeof v === 'number') return isFinite(v) ? v : null;
  let s = String(v).trim();
  if (!s) return null;
  if (/^(#(DIV\/0!|VALUE!|REF!|N\/A|NAME\?)|-{1,2}|\/|—|－|无)$/i.test(s)) return null;
  if (s.indexOf('合并到') >= 0) return { merged: true, text: s };
  let m = s.match(/^(-?[\d.]+)\s*万$/);
  if (m) return Number(m[1]) * 10000;
  if (type === 'sec') {
    m = s.match(/^(-?[\d.]+)\s*(s|秒)$/i); if (m) return Number(m[1]);
    m = s.match(/^(-?[\d.]+)\s*(min|分钟|分)$/i); if (m) return Number(m[1]) * 60;
  }
  if (type === 'pct') {
    m = s.match(/^(-?[\d.]+)\s*%$/); if (m) return Number(m[1]) / 100;
  }
  const n = Number(s.replace(/,/g, ''));
  return isNaN(n) ? s : n;   // 保留文本（如 1:44 / 未推广）
}
// 'M.D' 或 'M月D日' 形式的 sheet 名 → 归一化日期
function sheetNameToDate(name, year) {
  const s = String(name || '').trim();
  let m = s.match(/^(\d{1,2})[.．\-/月](\d{1,2})日?$/);
  if (!m) return null;
  const mo = +m[1], da = +m[2];
  if (mo < 1 || mo > 12 || da < 1 || da > 31) return null;
  return `${year}-${pad2(mo)}-${pad2(da)}`;
}

async function onFileStores(file) {
  if (!file) return;
  try {
    await ensureXlsx();
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array', cellDates: true });
    // 年份推断：文件名含 20xx → 用它；否则用订单最大日期年份；再否则用当前年
    let year = new Date().getFullYear();
    const ym = String(file.name).match(/(20\d{2})/);
    if (ym) year = +ym[1];
    else if (state.orders.length) {
      const max = state.orders.reduce((a, r) => (r.payTime && fmtDate(r.payTime) > a ? fmtDate(r.payTime) : a), '');
      if (max) year = +max.slice(0, 4);
    }
    const parsedDates = [];
    let storeHits = 0;
    wb.SheetNames.forEach((sn) => {
      const d = sheetNameToDate(sn, year);
      if (!d) return;                              // 跳过「工作表1」等辅助 sheet
      const ws = wb.Sheets[sn];
      const grid = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: true, raw: true });
      if (!grid || grid.length < 5) return;
      const day = state.rows[d] || (state.rows[d] = {});
      let last = null;
      for (let i = 3; i < Math.min(grid.length, 17); i++) {   // Excel 行 4..17
        const row = grid[i] || [];
        const a = row[0] == null ? '' : String(row[0]).trim();
        if (a) {
          const st = STORE_BY_NAME.get(a) || STORES.find((x) => a.indexOf(x.name.slice(0, 8)) === 0);
          if (st) last = st;
        }
        if (!last) continue;
        const rec = day[last.key] || (day[last.key] = { today: {}, cum: {} });
        if (!rec.today) rec.today = {};
        if (!rec.cum) rec.cum = {};
        // 「统计口径」决定这一行落到 today 还是 cum 桶（同一列在两行有不同值，不能互相覆盖）
        const periodTxt = String(row[6] == null ? '' : row[6]).trim();
        const bucket = /累计/.test(periodTxt) ? 'cum' : 'today';
        // A–F 为两行纵向合并（值只写在首行），此处按「有值才写」合并到店级共享字段
        ['pm', 'op', 'mt', 'dp', 'rank'].forEach((k, idx) => {
          const col = COLS.find((c) => c.k === k);
          const val = cleanCell(row[1 + idx], col.type);
          if (val !== null) rec[k] = val;
        });
        COLS.forEach((col, ci) => {
          if (ci <= 6) return;                      // A–G 已处理
          const val = cleanCell(row[ci], col.type);
          if (val !== null) rec[bucket][col.k] = val;
        });
        storeHits++;
      }
      parsedDates.push(d);
    });
    if (!parsedDates.length) {
      toast('未识别到日期 sheet（sheet 名需形如「10.23」），请确认上传的是七大店铺数据总表', 'err', 7000);
      return;
    }
    state.meta.storeYear = year;
    saveLocal();
    syncDateRange();
    autoPickDate();
    render();
    toast(`七大店铺总表解析完成：${parsedDates.length} 个日期（${parsedDates.sort().join('、')}），${storeHits} 行店铺数据`, 'ok', 6000);
  } catch (e) {
    console.error(e);
    toast('七大店铺总表解析失败：' + e.message, 'err', 6000);
  }
}

/* =====================================================================
 * 四之二、平台日报三表解析（客流 / 在线咨询 / 门店交易，支持一次多选）
 * ---------------------------------------------------------------------
 * 三份导出的结构一致：第 1 行表头，其后每行 = 一个「日期 × 门店」。
 * 类型靠**表头**判定（不看文件名，用户可能改过名）：
 *   门店交易 含「下单券数」且「核销券数」
 *   客流     含「曝光人数」或「访问人数」
 *   在线咨询 含「平均响应时长」或「30秒内回复率」
 * 列一律**精确名优先**：如「下单金额」不能误取「下单金额（原价）」，
 * 「平均响应时长」不能误取「平均响应时长均值 / 优秀值」。
 * ===================================================================== */
const PLAT_FIELDS = ['expo', 'visit', 'orderU', 'orderC', 'orderAmt', 'woU', 'woC', 'woAmt', 'replySec', 'reply30'];
const PLAT_TYPE_LABEL = { flow: '客流', chat: '在线咨询', trade: '门店交易' };
const PLAT_TYPE_SHORT = { flow: '客流', chat: '咨询', trade: '交易' };

function platNum(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = Number(String(v).replace(/[,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}
// 日期单元格 → 'YYYY-MM-DD'（兼容文本 / Date / Excel 序列号）
function parseDateCell(v) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return fmtDate(v);
  if (typeof v === 'number' && v > 20000 && v < 80000) {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }
  const d = parseDateLoose(v);
  return d ? fmtDate(d) : '';
}
function findHeaderRow(grid) {
  for (let i = 0; i < Math.min(grid.length, 8); i++) {
    const row = (grid[i] || []).map((v) => String(v == null ? '' : v).trim());
    if (row.includes('日期') && row.some((h) => h.includes('门店'))) return i;
  }
  return -1;
}
// 精确名优先；否则包含匹配（可用 exclude 排除「…原价 / …均值 / …优秀值」这类近似列）
function colIdx(hdr, names, exclude) {
  for (const nm of names) {
    const i = hdr.indexOf(nm);
    if (i >= 0) return i;
  }
  for (const nm of names) {
    const i = hdr.findIndex((h) => h.includes(nm) && (!exclude || !h.includes(exclude)));
    if (i >= 0) return i;
  }
  return -1;
}
function classifyPlat(hdr) {
  const has = (t) => hdr.some((h) => h.includes(t));
  if (has('下单券数') && has('核销券数')) return 'trade';
  if (has('曝光人数') || has('访问人数')) return 'flow';
  if (has('平均响应时长') || has('30秒内回复率')) return 'chat';
  return null;
}
/* 门店报表（推广费）表头：日期 + 推广门店/门店 + 花费（元）/推广费/消耗。
   单独「上传门店报表」按钮与「平台数据（可多选）」共用这一判定，避免两份解析逻辑。 */
function isPromoHeader(hdr) {
  const has = (t) => hdr.some((h) => h.includes(t));
  if (!has('日期')) return false;
  if (!(has('推广门店') || has('门店名称') || hdr.indexOf('门店') >= 0)) return false;
  return has('花费') || has('推广费') || has('消耗');
}

async function parsePlatFile(file) {
  await ensureXlsx();
  const buf = await file.arrayBuffer();
  // .csv 不能按二进制 array 读（SheetJS 会解析不出工作表），先解码成文本再按 string 解析
  const wb = /\.csv$/i.test(String(file.name || ''))
    ? XLSX.read(new TextDecoder('utf-8').decode(buf), { type: 'string', raw: true })
    : XLSX.read(buf, { type: 'array', cellDates: true });
  const out = { type: null, rows: [] };
  // 门店报表的日期是「10-27」这种不带年份的写法，年份从文件名 / 订单最大日期 / 当前年推断
  let promoYear = new Date().getFullYear();
  const ym = String(file.name || '').match(/(20\d{2})/);
  if (ym) promoYear = +ym[1];
  else if (state.orders.length) {
    const mx = state.orders.reduce((a, r) => (r.payTime && fmtDate(r.payTime) > a ? fmtDate(r.payTime) : a), '');
    if (mx) promoYear = +mx.slice(0, 4);
  }
  for (const sn of wb.SheetNames) {
    const grid = XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, defval: null, blankrows: false, raw: true });
    if (!grid || grid.length < 2) continue;
    const hi = findHeaderRow(grid);
    if (hi < 0) continue;
    const hdr = (grid[hi] || []).map((v) => String(v == null ? '' : v).trim());
    // —— 门店报表（推广费）：日期 × 门店 → 花费 ——
    if (isPromoHeader(hdr)) {
      const iPDate = colIdx(hdr, ['日期']);
      const iPStore = colIdx(hdr, ['推广门店', '门店名称', '门店']);
      const iFee = colIdx(hdr, ['花费（元）', '花费(元)', '花费', '推广费', '消耗']);
      if (iPDate >= 0 && iPStore >= 0 && iFee >= 0) {
        const promo = [];
        for (let r = hi + 1; r < grid.length; r++) {
          const row = grid[r] || [];
          const date = parsePromoDate(row[iPDate], promoYear);
          const name = row[iPStore] == null ? '' : String(row[iPStore]).trim();
          const rawFee = row[iFee];
          if (!date || !name || rawFee == null || rawFee === '') continue;
          const st = resolveStoreByName(name);
          promo.push({ date, name, store: st ? st.key : '', fee: platNum(rawFee) });
        }
        if (promo.length) { out.type = 'promo'; out.promo = (out.promo || []).concat(promo); continue; }
      }
    }
    const type = classifyPlat(hdr);
    if (!type) continue;
    const iDate = colIdx(hdr, ['日期']);
    const iStore = colIdx(hdr, ['门店名称', '门店']);
    if (iDate < 0 || iStore < 0) continue;
    const idx = {};
    if (type === 'flow') {
      idx.expo = colIdx(hdr, ['曝光人数']);
      idx.visit = colIdx(hdr, ['访问人数']);
    } else if (type === 'chat') {
      idx.replySec = colIdx(hdr, ['平均响应时长']);
      idx.reply30 = colIdx(hdr, ['30秒内回复率']);
      idx.chatCnt = colIdx(hdr, ['在线咨询人数']);
    } else {
      idx.orderU = colIdx(hdr, ['下单人数']);
      idx.orderC = colIdx(hdr, ['下单券数']);
      idx.orderAmt = colIdx(hdr, ['下单金额'], '原价');
      idx.woU = colIdx(hdr, ['核销人数']);
      idx.woC = colIdx(hdr, ['核销券数']);
      idx.woAmt = colIdx(hdr, ['核销金额'], '原价');
    }
    const rows = [];
    for (let r = hi + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const date = parseDateCell(row[iDate]);
      if (!date) continue;
      const st = resolveStoreByName(row[iStore]);
      if (!st) continue;
      const vals = {};
      if (type === 'flow') {
        vals.expo = platNum(row[idx.expo]);
        vals.visit = platNum(row[idx.visit]);
      } else if (type === 'trade') {
        vals.orderU = platNum(row[idx.orderU]);
        vals.orderC = platNum(row[idx.orderC]);
        vals.orderAmt = platNum(row[idx.orderAmt]);
        vals.woU = platNum(row[idx.woU]);
        vals.woC = platNum(row[idx.woC]);
        vals.woAmt = platNum(row[idx.woAmt]);
      } else {
        const sec = cleanCell(row[idx.replySec], 'sec');
        const p30 = cleanCell(row[idx.reply30], 'pct');
        if (isNum(sec)) vals.replySec = sec;
        if (isNum(p30)) vals.reply30 = p30;
        vals.chatCnt = idx.chatCnt >= 0 ? platNum(row[idx.chatCnt]) : 0;
      }
      rows.push({ date, store: st.key, vals });
    }
    if (rows.length) { out.type = type; out.rows = rows; break; }
  }
  return out;
}

async function onFilesPlatform(fileList) {
  const list = Array.from(fileList || []);
  if (!list.length) return;
  const fileCnt = { flow: 0, chat: 0, trade: 0 };
  const rowCnt = { flow: 0, chat: 0, trade: 0 };
  const skipped = [];
  // 门店报表（推广费）也可能被一起选进来：按表头自动分流，不用单独点「上传门店报表」
  let promoFiles = 0, promoRows = 0;
  const promoUnmatched = new Set();
  // 导入前的最大日期：只有本次导入真的把可用日期往后延伸时，才自动跳到最新一天
  const prevDates = allDates();
  const prevMax = prevDates.length ? prevDates[prevDates.length - 1] : '';
  try {
    for (const f of list) {
      const res = await parsePlatFile(f);
      if (res.type === 'promo') {
        const { rows, unmatched } = applyPromoRows(res.promo);
        if (rows) promoFiles++;
        promoRows += rows;
        unmatched.forEach((n) => promoUnmatched.add(n));
        continue;
      }
      if (!res.type || !res.rows.length) { skipped.push(f.name); continue; }
      fileCnt[res.type]++;
      res.rows.forEach((r) => {
        const day = state.daily[r.date] || (state.daily[r.date] = {});
        const rec = day[r.store] || (day[r.store] = {});
        Object.assign(rec, r.vals);
        rowCnt[res.type]++;
      });
    }
    const got = fileCnt.flow + fileCnt.chat + fileCnt.trade;
    if (!got && !promoFiles) {
      toast('未识别到可导入的数据表。支持：客流 / 在线咨询 / 门店交易（平台导出）与门店报表（推广费），表头需含 日期 + 门店名称', 'err', 8000);
      return;
    }
    const dates = Object.keys(state.daily)
      .filter((d) => Object.keys(state.daily[d]).length)
      .sort();
    const storeSet = new Set();
    dates.forEach((d) => Object.keys(state.daily[d]).forEach((k) => storeSet.add(k)));
    const prev = state.meta.plat || {};
    const types = Object.assign({}, prev.types || {});
    Object.keys(fileCnt).forEach((t) => { if (fileCnt[t]) types[t] = 1; });
    state.meta.plat = {
      types,
      files: (prev.files || 0) + got,
      dates: dates.length,
      from: dates[0] || '',
      to: dates[dates.length - 1] || '',
      stores: storeSet.size,
      ts: Date.now(),
    };
    if (promoFiles) refreshPromoMeta();     // 推广费元信息与门店报表按钮同源
    saveLocal();
    _platCache.key = '';
    if (promoFiles) _promoCache.key = '';
    syncDateRange();
    // 新数据的日期范围更靠后 → 直接定位到最新一天；否则保留用户当前选择
    const nextDates = allDates();
    const nextMax = nextDates.length ? nextDates[nextDates.length - 1] : '';
    if (nextMax && nextMax > prevMax) state.date = nextMax;
    autoPickDate();
    render();
    const parts = Object.keys(fileCnt)
      .filter((t) => fileCnt[t])
      .map((t) => `${PLAT_TYPE_LABEL[t]} ${fileCnt[t]} 个（${rowCnt[t]} 行）`);
    if (promoFiles) parts.push(`推广费 ${promoFiles} 个（${promoRows} 条）`);
    let msg = `数据导入完成：${parts.join(' · ')}`;
    if (got) msg += `；平台数据覆盖 ${storeSet.size} 店 × ${dates.length} 天`;
    if (promoFiles) msg += promoUnmatchedTip(promoUnmatched);
    if (skipped.length) msg += `；未识别 ${skipped.length} 个文件`;
    toast(msg, promoUnmatched.size ? 'warn' : 'ok', promoUnmatched.size ? 9000 : 6500);
  } catch (e) {
    console.error(e);
    toast('平台数据解析失败：' + e.message, 'err', 7000);
  }
}

/* =====================================================================
 * 四之二·乙、门店报表（推广费）解析
 * ---------------------------------------------------------------------
 * 门店报表_*.xlsx：A=日期(M-D) / B=推广门店 / E=花费（元）= 推广费。
 * 按 日期×门店 写入 state.promo；年分从文件名 / 订单最大日期 / 当前年 推断。
 * ===================================================================== */
function parsePromoDate(v, year) {
  if (v == null || v === '') return '';
  if (v instanceof Date) return fmtDate(v);
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})-(\d{1,2})$/);
  if (m) return `${year}-${pad2(+m[1])}-${pad2(+m[2])}`;
  const d = parseDateLoose(v);
  return d ? fmtDate(d) : '';
}

/* 把 parsePlatFile 解析出的推广费行写入 state.promo；返回 { rows, unmatched } */
function applyPromoRows(promoRows) {
  const unmatched = new Set();
  let rows = 0;
  (promoRows || []).forEach((p) => {
    if (!p.store) { if (p.name) unmatched.add(p.name); return; }
    const day = state.promo[p.date] || (state.promo[p.date] = {});
    day[p.store] = p.fee;
    rows++;
  });
  return { rows, unmatched };
}
function refreshPromoMeta() {
  const dates = Object.keys(state.promo).filter((d) => Object.keys(state.promo[d] || {}).length);
  const stores = new Set();
  let rows = 0;
  dates.forEach((d) => {
    const keys = Object.keys(state.promo[d]);
    rows += keys.length;
    keys.forEach((k) => stores.add(k));
  });
  state.meta.promo = { ts: Date.now(), rows, stores: stores.size, dates: dates.length };
  return state.meta.promo;
}
function promoUnmatchedTip(unmatched) {
  return (unmatched && unmatched.size) ? '；未匹配门店：' + [...unmatched].join('、') : '';
}

async function onFilePromo(file) {
  if (!file) return;
  try {
    await ensureXlsx();
    const res = await parsePlatFile(file);
    if (res.type !== 'promo' || !res.promo || !res.promo.length) {
      toast('未识别到门店报表。表头需含「日期」「推广门店」「花费（元）」，请确认上传的是门店报表导出文件', 'err', 8000);
      return;
    }
    const { rows, unmatched } = applyPromoRows(res.promo);
    if (!rows) {
      toast('未解析到任何推广费记录' + (unmatched.size ? '，门店名均未匹配：' + [...unmatched].join('、') : ''), 'err', 8000);
      return;
    }
    const m = refreshPromoMeta();
    saveLocal();
    _promoCache.key = '';
    syncDateRange();
    autoPickDate();
    render();
    toast(`门店报表解析完成：${rows} 条推广费，覆盖 ${m.dates} 天 × ${m.stores} 店` + promoUnmatchedTip(unmatched),
      unmatched.size ? 'warn' : 'ok', unmatched.size ? 9000 : 6000);
  } catch (e) {
    console.error(e);
    toast('门店报表解析失败：' + e.message, 'err', 6000);
  }
}

/* =====================================================================
 * 四之三、平台三表的 当日 / 月累计 聚合
 * ---------------------------------------------------------------------
 * 当日 = 该日期行的原值；月累计 = 当月 1 日 ~ 所选日逐日累加。
 * 回复时长 / 30s 回复率为**比率**，不能相加 —— 按当日「在线咨询人数」加权平均
 * （无该列时权重取 1，退化为简单平均）。
 * ===================================================================== */
let _platCache = { key: '', data: null };
const PLAT_SUM_KEYS = ['expo', 'visit', 'orderU', 'orderC', 'orderAmt', 'woU', 'woC', 'woAmt'];

/* 「美团数据（曝光/下单/核销）」「美团榜单维护及客资转化」两组，表头写的就是「填前一天数据」：
   美团/点评后台当天出不了当天数，报表日期 D 的这一行填的是 D-1 的数。
   故平台指标统一按「前一天」取数：当日行 = D-1 的平台数据，月累计 = 当月 1 日 ~ D-1。
   D-1 当天没有平台数据时（周末/节假日出数会晚），回退到 ≤ D-1 的最近一个有数据的一天；
   只有 D 当天有数据时（当天导出已到）才退回用 D，避免整列空着。 */
function addDays(iso, n) {
  const d = iso2date(iso);
  if (!d) return '';
  d.setDate(d.getDate() + n);
  return fmtDate(d);
}
function platDate() {
  const date = state.date;
  if (!date) return '';
  const prev = addDays(date, -1);
  if (!prev) return '';
  const ds = Object.keys(state.daily).filter((d) => d && Object.keys(state.daily[d] || {}).length);
  if (!ds.length) return prev;                                  // 还没有平台数据：仍按前一天（显示 —）
  if (ds.indexOf(prev) >= 0) return prev;
  const earlier = ds.filter((d) => d < prev).sort();
  if (earlier.length) return earlier[earlier.length - 1];       // 用 ≤ D-1 的最近一天
  return ds.indexOf(date) >= 0 ? date : prev;
}

function platAgg() {
  const date = state.date;
  if (!date) return null;
  const pdate = platDate() || date;
  const key = pdate + '|' + (state.meta.plat ? state.meta.plat.ts : 0) + '|' + Object.keys(state.daily).length;
  if (_platCache.key === key) return _platCache.data;
  const day = iso2date(date);                       // 月份边界按报表日期算（「本月累计」= 报表日期所在月）
  const start = fmtDate(new Date(day.getFullYear(), day.getMonth(), 1));
  const today = {}, cum = {};
  STORE_KEYS.forEach((k) => { today[k] = null; cum[k] = null; });
  STORE_KEYS.forEach((k) => {
    const r = (state.daily[pdate] || {})[k];
    if (r) today[k] = r;
  });
  Object.keys(state.daily)
    .filter((d) => d >= start && d <= pdate)
    .sort()
    .forEach((d) => {
      const dayRec = state.daily[d] || {};
      STORE_KEYS.forEach((k) => {
        const r = dayRec[k];
        if (!r) return;
        const c = cum[k] || (cum[k] = {});
        PLAT_SUM_KEYS.forEach((f) => { if (isNum(r[f])) c[f] = (c[f] || 0) + r[f]; });
        const w = isNum(r.chatCnt) && r.chatCnt > 0 ? r.chatCnt : 1;
        if (isNum(r.replySec)) { c._secW = (c._secW || 0) + r.replySec * w; c._secN = (c._secN || 0) + w; }
        if (isNum(r.reply30)) { c._p30W = (c._p30W || 0) + r.reply30 * w; c._p30N = (c._p30N || 0) + w; }
      });
    });
  STORE_KEYS.forEach((k) => {
    const c = cum[k];
    if (!c) return;
    if (c._secN) c.replySec = c._secW / c._secN;
    if (c._p30N) c.reply30 = c._p30W / c._p30N;
    delete c._secW; delete c._secN; delete c._p30W; delete c._p30N;
  });
  const data = { today, cum };
  _platCache = { key, data };
  return data;
}

/* =====================================================================
 * 四之四、门店报表（推广费）聚合
 * ---------------------------------------------------------------------
 * 数据源：门店报表_*.xlsx（A=日期 M-D / B=推广门店 / E=花费（元）= 推广费）。
 * 当日取该日值；月累计 = 当月 1 日 ~ 所选日逐日累加（与平台三表口径一致）。
 * 优先级：手工覆盖 > 门店报表文件 > 七大店铺总表（cellValue 内裁决）。
 * ===================================================================== */
let _promoCache = { key: '', data: null };

function promoAgg() {
  const date = state.date;
  if (!date) return null;
  const key = date + '|' + (state.meta.promo ? state.meta.promo.ts : 0) + '|' + Object.keys(state.promo).length;
  if (_promoCache.key === key) return _promoCache.data;
  const day = iso2date(date);
  const start = fmtDate(new Date(day.getFullYear(), day.getMonth(), 1));
  const today = {}, cum = {};
  STORE_KEYS.forEach((k) => { today[k] = null; cum[k] = null; });
  STORE_KEYS.forEach((k) => {
    const r = (state.promo[date] || {})[k];
    if (r != null) today[k] = r;
  });
  Object.keys(state.promo)
    .filter((d) => d >= start && d <= date)
    .sort()
    .forEach((d) => {
      const rec = state.promo[d] || {};
      STORE_KEYS.forEach((k) => { const v = rec[k]; if (isNum(v)) cum[k] = (isNum(cum[k]) ? cum[k] : 0) + v; });
    });
  const data = { today, cum };
  _promoCache = { key, data };
  return data;
}

/* =====================================================================
 * 五、计算：每店 当日 / 月累计 的 初诊业绩、初诊客量、全部客量、VAB
 * ===================================================================== */
let _aggCache = { key: '', data: null };

function computeAgg() {
  const date = state.date;
  if (!date) return null;
  const key = date + '|' + state.orders.length + '|' + state.ordersTs + '|' + state.histOrders.length + '|' + state.vabSet.size;
  if (_aggCache.key === key) return _aggCache.data;

  const day = iso2date(date);
  const monthStart = new Date(day.getFullYear(), day.getMonth(), 1);
  // 当月 1 日 ~ 所选日 的全部日期
  const dates = [];
  for (let t = new Date(monthStart); t <= day; t = new Date(t.getTime() + 86400000)) dates.push(fmtDate(t));

  // 按日期分组（只保留电商组范围内的订单，其它可忽略）
  const byDate = new Map();
  state.orders.forEach((r) => {
    if (!r.payTime) return;
    const d = fmtDate(r.payTime);
    if (d < fmtDate(monthStart) || d > date) return;
    if (!isEcSubScope(r)) return;
    if (!byDate.has(d)) byDate.set(d, []);
    byDate.get(d).push(r);
  });

  const blank = () => {
    const o = { actAmt: 0, firstVis: 0, lead: 0, newVab: 0, allAmt: 0, aov: 0, has: false };
    return o;
  };
  const today = {}, cum = {};
  STORE_KEYS.forEach((k) => { today[k] = blank(); cum[k] = blank(); });

  dates.forEach((d) => {
    const dayOrders = byDate.get(d) || [];
    if (!dayOrders.length) return;
    // 逐日复用 ③ 的初诊集合（每个日期用它自己的 cutoff —— 与报表每日列完全一致）
    const cutoff = iso2date(d);
    const gate = buildMktFirstKeys(state.orders, cutoff);
    STORE_KEYS.forEach((storeKey) => {
      const list = dayOrders.filter((r) => (r.channelL2 || '').trim() === storeKey);
      if (!list.length) return;
      let amt = 0; const vabCli = new Set(); const firstCli = new Set(); const allCli = new Set();
      list.forEach((r) => {
        // ⑤ 的「电商组」行 = r3∪r3b∪r3c：
        //   自动规则行必须先过 30 天初诊门槛（autoYx）；
        //   而手动「报表归属 = 电商组」的行（ov）不受门槛约束，直接计入。
        const isFirst = String(r.deptOverride || '').trim() === '电商组' ? true : gate.has(mktFirstClientKey(r));
        const ck = mktFirstClientKey(r);
        if (ck && ck !== 'name:') allCli.add(ck);
        if (!isFirst) return;
        amt += Number(r.cashPay) || 0;
        if (ck && ck !== 'name:') {
          firstCli.add(ck);
          if (state.vabSet.size && state.vabSet.has(normMemberKey(r.memberNo))) vabCli.add(ck);
        }
      });
      const allAmt = list.reduce((a, r) => a + (Number(r.cashPay) || 0), 0);
      // 当日：只累加「所选日期」这一天（此前误在整月循环里累加，会把当日做成月合计）
      if (d === date) {
        const t = today[storeKey];
        t.actAmt += amt; t.firstVis += firstCli.size; t.lead += allCli.size; t.newVab += vabCli.size; t.allAmt += allAmt;
      }
      // 月累计（逐日累加；客量按整月去重，与 ⑤ 的 cumCntDays 同语义）
      const c = cum[storeKey];
      c.actAmt += amt; c.newVab += vabCli.size; c.allAmt += allAmt;
      c._firstCli = c._firstCli || new Set(); c._allCli = c._allCli || new Set();
      firstCli.forEach((x) => c._firstCli.add(x));
      allCli.forEach((x) => c._allCli.add(x));
    });
  });

  STORE_KEYS.forEach((k) => {
    ['today', 'cum'].forEach((w) => {
      const o = (w === 'today' ? today : cum)[k];
      if (w === 'cum') {
        o.firstVis = (o._firstCli ? o._firstCli.size : 0);
        o.lead = (o._allCli ? o._allCli.size : 0);
        delete o._firstCli; delete o._allCli;
      }
      o.aov = o.firstVis > 0 ? o.actAmt / o.firstVis : null;
      o.has = o.actAmt !== 0 || o.firstVis !== 0 || o.lead !== 0;
    });
  });

  const data = { today, cum, dates };
  _aggCache = { key, data };
  return data;
}

/* =====================================================================
 * 六、取值（含手工覆盖 / 源表回落）
 * ===================================================================== */
function manualKey(date, storeKey, field) { return date + '|' + storeKey + '|' + field; }
function storeByKey(key) { return STORES.find((s) => s.key === key) || null; }
/** 店铺展示名：手工改名优先（不随日期变化，导入时的名称匹配仍用 STORES 原名） */
function storeName(store) {
  if (!store) return '';
  const v = state.names[store.key];
  return (v === null || v === undefined || v === '') ? store.name : v;
}

/**
 * 取某店某日某列的展示值
 * 返回 { v, kind }  kind: 'auto' | 'sheet' | 'manual' | 'merged' | null
 * 结构：state.rows[date][storeKey] = { pm,op,mt,dp,rank, today:{…}, cum:{…} }
 *   店级共享字段（责任人 / 评分 / 榜单）在两行都一样；其余按「当日数据 / 本月累计」分桶。
 */
function cellValue(storeKey, field, which) {
  const date = state.date;
  const col = COLS.find((c) => c.k === field);
  const mk = manualKey(date, storeKey, field);
  if (Object.prototype.hasOwnProperty.call(state.manual, mk)) {
    return { v: state.manual[mk], kind: 'manual' };
  }
  const rec = (state.rows[date] || {})[storeKey] || {};
  const shared = ['pm', 'op', 'mt', 'dp', 'rank'];
  let sheetVal;
  if (shared.includes(field)) sheetVal = rec[field];
  else sheetVal = (rec[which] || {})[field];

  // 平台三表（客流 / 在线咨询 / 门店交易）优先于七大店铺总表的同名列；
  // 「仅总表」模式（source === 'sheet'）下跳过，用于和原表核对。
  if (state.source === 'auto' && PLAT_FIELDS.indexOf(field) >= 0) {
    const p = platAgg();
    const pr = p && p[which] ? p[which][storeKey] : null;
    if (pr && isNum(pr[field])) return { v: pr[field], kind: 'plat' };
  }

  // 自动列
  if (col && col.auto && state.source === 'auto') {
    const agg = computeAgg();
    const o = agg ? agg[which][storeKey] : null;
    let autoVal = null;
    if (field === 'actAmt') autoVal = o.actAmt;
    else if (field === 'firstVis') autoVal = o.firstVis;
    else if (field === 'lead') autoVal = o.lead;
    else if (field === 'newVab') autoVal = o.newVab;
    else if (field === 'aov') autoVal = o.aov;
    if (autoVal !== null && autoVal !== undefined && (autoVal !== 0 || !hasSheetNumber(sheetVal))) {
      return { v: autoVal, kind: 'auto' };
    }
    // 订单侧无数据 → 回落源表
    if (sheetVal !== undefined) return normSheet(sheetVal);
    return { v: autoVal, kind: 'auto' };
  }
  // 推广费（元）：用户指定数据源 = 门店报表文件，优先于七大店铺总表的同名 Z 列
  if (col && col.k === 'promoFee' && state.source === 'auto') {
    const pa = promoAgg();
    const pv = pa && pa[which] ? pa[which][storeKey] : null;
    if (isNum(pv)) return { v: pv, kind: 'promo' };
  }
  if (sheetVal !== undefined) return normSheet(sheetVal);
  return { v: null, kind: null };
}
function hasSheetNumber(v) { return typeof v === 'number' && v !== 0; }
function normSheet(v) {
  if (v && typeof v === 'object' && v.merged) return { v: v.text, kind: 'merged' };
  return { v, kind: 'sheet' };
}

/* ---------------- 同店两行的纵向合并判定 ---------------- */
function rawTextOf(storeKey, field, which) {
  const { v } = cellValue(storeKey, field, which);
  if (v === null || v === undefined || v === '') return '';
  return String(v);
}
// 投产比 / 好评 5 列：两行取值相同、或其中一行为空 → 合并（模板表就是这么合并的）；
// 两行都有值且不同 → 保留两行（如好评目标 20 / 40）。
function pairMergeable(store, field) {
  const a = rawTextOf(store.key, field, 'today');
  const b = rawTextOf(store.key, field, 'cum');
  if (!a || !b) return true;
  return a === b;
}
// 某店「当日 + 月累计」两行需要纵向合并的列集合
function mergeSetFor(store) {
  const set = new Set(MERGED_FIELDS);
  COND_MERGE_FIELDS.forEach((f) => { if (pairMergeable(store, f)) set.add(f); });
  return set;
}

/* ---------------- 渲染格式化 ---------------- */
function display(kind, v, col) {
  if (v === null || v === undefined || v === '') return '<span class="dash">—</span>';
  if (typeof v === 'string') return esc(v);
  switch (col.type) {
    case 'money': return money(v, 2);
    case 'money0': return money(v, 0);
    case 'int': return int0(v);
    case 'pct': return (v * 100).toFixed(2) + '%';
    case 'sec': return Number(v).toFixed(1) + 's';
    case 'score': return Number(v).toFixed(1);
    case 'scoreText': return esc(String(v));
    default: return esc(String(v));
  }
}

/* =====================================================================
 * 七、渲染
 * ===================================================================== */
function ensureColgroup() {
  const cg = $('tbl-colgroup');
  if (cg.childElementCount) return;
  cg.innerHTML = COLS.map((c) => `<col style="width:${c.w}px">`).join('');
}

function renderThead() {
  const totalW = COLS.reduce((a, c) => a + c.w, 0);
  let row2 = '';
  // 按分组顺序输出，period 单独 rowspan=2
  const order = [];
  COLS.forEach((c) => {
    if (!order.includes(c.g)) order.push(c.g);
  });
  GROUPS.forEach((g) => {
    if (!order.includes(g.id)) return;
    const span = COLS.filter((c) => c.g === g.id).length;
    if (g.rowspan) row2 += `<th class="${g.cls}" rowspan="2" style="width:${COLS.find((c) => c.g === g.id).w}px">${esc(g.label)}</th>`;
    else row2 += `<th class="${g.cls}" colspan="${span}">${esc(g.label)}</th>`;
  });
  const row3 = COLS.map((c) => {
    if (c.k === 'period') return '';
    const isFirst = c.k === 'store';
    return `<th class="g3${isFirst ? ' c-store' : ''}">${esc(c.label)}</th>`;
  }).join('');
  $('tbl-thead').innerHTML = `
    <tr><th class="t-title" colspan="${COLS.length}" style="width:${totalW}px">海蓓医疗七大店铺团购数据总表</th></tr>
    <tr>${row2}</tr>
    <tr>${row3}</tr>`;
}

function cellHTML(c, store, which, merge, mergedSet) {
  const isToday = which === 'today';
  const shared = mergedSet ? mergedSet.has(c.k) : MERGED_FIELDS.indexOf(c.k) >= 0;
  if (merge && shared && !isToday) return '';   // 第 2 行：共享列已由第 1 行 rowspan 覆盖
  const rs = (merge && shared) ? ' rowspan="2"' : '';
  // 按值合并的列（投产比 / 好评）：当日为空、累计有值时用累计值填这一格，
  // 否则合并后第 2 行被吃掉，这个值就消失了。
  let src = which;
  if (merge && shared && MERGED_FIELDS.indexOf(c.k) < 0
      && !rawTextOf(store.key, c.k, 'today') && rawTextOf(store.key, c.k, 'cum')) src = 'cum';
  if (c.k === 'store') {
    const edited = Object.prototype.hasOwnProperty.call(state.names, store.key) && state.names[store.key];
    return `<td class="c-store${edited ? ' name-edited' : ''}" data-edit="1" data-store="${store.key}" data-field="name" data-which="${src}"${rs}>${esc(storeName(store))}</td>`;
  }
  if (c.k === 'pm' || c.k === 'op') {
    const { v: mv, kind: mk } = cellValue(store.key, c.k, src);
    const def = c.k === 'pm' ? store.pm : store.op;
    const txt = (mv === null || mv === undefined || mv === '') ? (def || '—') : mv;
    return `<td class="txt editable${mk === 'manual' ? ' manual' : ''}" data-edit="1" data-store="${store.key}" data-field="${c.k}" data-which="${src}"${rs}>${esc(txt)}</td>`;
  }
  if (c.k === 'period') return `<td class="c-period"${rs}>${isToday ? '当日数据' : '本月累计'}</td>`;
  if (c.auto) {
    const { v, kind } = cellValue(store.key, c.k, src);
    const inner = (kind === 'auto' && typeof v === 'number')
      ? '<span class="adot" title="由订单自动计算"></span>' + display(kind, v, c)
      : display(kind, v, c);
    const klass = 'num' + (kind === 'manual' ? ' editable manual' : '') + (kind === 'sheet' ? ' fb' : '') + (c.zone ? ' z-' + c.zone : '');
    return `<td class="${klass}" data-edit="1" data-store="${store.key}" data-field="${c.k}" data-which="${src}"${rs}>${inner}</td>`;
  }
  const { v, kind } = cellValue(store.key, c.k, src);
  const klass = (c.type === 'text' || c.type === 'scoreText' ? 'txt' : 'num')
    + ' editable' + (c.zone ? ' z-' + c.zone : '');
  // 平台三表算出的格子加蓝点（区别于订单自动计算的绿点）
  const inner = (kind === 'plat' && typeof v === 'number')
    ? '<span class="pdot" title="由平台日报数据汇总（客流 / 在线咨询 / 门店交易）"></span>' + display(kind, v, c)
    : display(kind, v, c);
  return `<td class="${klass}" data-edit="1" data-store="${store.key}" data-field="${c.k}" data-which="${src}"${rs}>${inner}</td>`;
}
function rowHTML(store, which) {
  const cls = which === 'today' ? 'r-today' : 'r-cum';
  return `<tr class="${cls}">${COLS.map((c) => cellHTML(c, store, which, false)).join('')}</tr>`;
}
// 「当日 + 月累计」同行展示：店级共享列纵向合并（rowspan=2），与原始模板表一致
function storeBlockHTML(store) {
  const ms = mergeSetFor(store);
  const row1 = COLS.map((c) => cellHTML(c, store, 'today', true, ms)).join('');
  const row2 = COLS.map((c) => cellHTML(c, store, 'cum', true, ms)).join('');
  return `<tr class="r-today">${row1}</tr><tr class="r-cum">${row2}</tr>`;
}

function sumRowHTML(label, which) {
  const cells = COLS.map((c) => {
    if (c.k === 'store') return `<td class="c-store" style="font-weight:700;">${esc(label)}</td>`;
    if (c.k === 'period') return `<td class="c-period">—</td>`;
    if (c.type === 'text' || c.type === 'scoreText' || !SUM_KEYS.includes(c.k)) {
      return `<td class="num${c.zone ? ' z-' + c.zone : ''}"><span class="dash">—</span></td>`;
    }
    let total = null;
    STORES.forEach((s) => {
      const { v } = cellValue(s.key, c.k, which);
      if (typeof v === 'number') total = (total || 0) + v;
    });
    return `<td class="num${c.zone ? ' z-' + c.zone : ''}">${total === null ? '<span class="dash">—</span>' : display(null, total, c)}</td>`;
  }).join('');
  return `<tr class="r-sum">${cells}</tr>`;
}

function renderTable() {
  const hasOrders = state.orders.length > 0;
  const hasSheet = Object.keys(state.rows).length > 0;
  const hasPlat = Object.keys(state.daily).some((d) => Object.keys(state.daily[d] || {}).length);
  if (!hasOrders && !hasSheet && !hasPlat) {
    $('tbl').style.display = 'none';
    $('tbl-empty').style.display = '';
    return;
  }
  $('tbl-empty').style.display = 'none';
  $('tbl').style.display = '';
  ensureColgroup();
  renderThead();
  let body = '';
  STORES.forEach((s) => {
    if (state.period === 'both') body += storeBlockHTML(s);
    else body += rowHTML(s, state.period);
  });
  $('tbl-body').innerHTML = body;
  const label = state.period === 'cum' ? '合计（本月累计）' : '合计（当日数据）';
  $('tbl-foot').innerHTML = sumRowHTML(label, state.period === 'cum' ? 'cum' : 'today');
}

function renderKpis() {
  const el = $('kpis');
  if (!state.date) { el.innerHTML = ''; return; }
  // 概览卡全部走 cellValue，与表格逐格同源（尊重「自动计算 / 源表值」切换与源表回落），
  // 避免出现「卡片显示 0、表格显示源表值」的口径分裂。
  const sumOf = (field, which) => {
    let total = 0, any = false;
    STORES.forEach((s) => {
      const { v } = cellValue(s.key, field, which);
      if (typeof v === 'number') { total += v; any = true; }
    });
    return any ? total : null;
  };
  const amtToday = sumOf('actAmt', 'today');
  const amtCum = sumOf('actAmt', 'cum');
  const tgt = sumOf('tgtAmt', 'today');
  const visToday = sumOf('firstVis', 'today');
  const leadToday = sumOf('lead', 'today');
  const rate = (tgt && tgt > 0 && amtCum !== null) ? amtCum / tgt : null;
  const aov = (amtToday !== null && visToday) ? amtToday / visToday : null;
  const y = (n) => (n === null ? '<span class="dash">—</span>' : '¥' + money(n, 2));
  const card = (k, v, extra) => `<div class="kpi"><div class="k">${k}</div><div class="v${extra && extra.cls ? ' ' + extra.cls : ''}">${v}${extra && extra.unit ? `<small>${extra.unit}</small>` : ''}</div></div>`;
  el.innerHTML = [
    card('初诊业绩（当日）', y(amtToday)),
    card('初诊业绩（月累计）', y(amtCum)),
    card('初诊目标 / 达成率', (tgt && rate !== null)
      ? ('¥' + money(tgt, 0) + ' <small>/ ' + (rate * 100).toFixed(1) + '%</small>')
      : '<span class="dash">—</span>'),
    card('初诊成交客量（当日）', visToday === null ? '<span class="dash">—</span>' : int0(visToday), { unit: '人' }),
    card('有效客资人数（当日）', leadToday === null ? '<span class="dash">—</span>' : int0(leadToday), { unit: '人' }),
    card('当日客单价', aov === null ? '<span class="dash">—</span>' : '¥' + money(aov, 0)),
  ].join('');
}

function renderStatus() {
  const om = $('orders-meta');
  if (state.orders.length) {
    const dates = state.orders.map((r) => fmtDate(r.payTime)).filter(Boolean).sort();
    om.textContent = `✔ 订单 ${state.orders.length} 行 · ${dates[0]} ~ ${dates[dates.length - 1]}`;
    om.className = 'file-meta ok';
  } else {
    om.textContent = '未导入（可先在「日业绩报表」导入，本页自动共用）';
    om.className = 'file-meta';
  }
  const sm = $('stores-meta');
  const ds = Object.keys(state.rows).sort();
  if (ds.length) {
    sm.textContent = `✔ 七大店铺总表 ${ds.length} 天 · ${ds[0]} ~ ${ds[ds.length - 1]}`;
    sm.className = 'file-meta ok';
  } else {
    sm.textContent = '未导入';
    sm.className = 'file-meta';
  }
  const pm = $('platform-meta');
  if (pm) {
    const p = state.meta.plat;
    if (p && p.dates) {
      const t = p.types || {};
      const label = Object.keys(PLAT_TYPE_SHORT).filter((k) => t[k]).map((k) => PLAT_TYPE_SHORT[k]).join('+');
      pm.textContent = `✔ 平台三表 ${label || '已导入'} · ${p.stores} 店 × ${p.dates} 天`;
      pm.className = 'file-meta ok';
    } else {
      pm.textContent = '未导入';
      pm.className = 'file-meta';
    }
  }
  const prm = $('promo-meta');
  if (prm) {
    const p = state.meta.promo;
    if (p && p.dates) {
      prm.textContent = `✔ 门店报表（推广费）· ${p.stores} 店 × ${p.dates} 天`;
      prm.className = 'file-meta ok';
    } else {
      prm.textContent = '未导入';
      prm.className = 'file-meta';
    }
  }
  const hint = $('date-hint');
  const hasAuto = state.orders.length > 0;
  const hasSheet = ds.length > 0;
  const hasPlat = !!(state.meta.plat && state.meta.plat.dates);
  const srcLabel = state.source === 'auto'
    ? (hasAuto || hasPlat ? '自动汇总（订单 + 平台三表）' : '订单与平台数据均为空 → 回落总表')
    : '仅七大店铺总表';
  const miss = [];
  if (!hasSheet) miss.push('未导入七大店铺总表，评分/榜单/好评列为空');
  if (!hasPlat) miss.push('未导入平台三表，曝光/下单/核销/回复时长沿用总表');
  const pd = platDate();
  // 页面已不再展示口径说明文字，这里保留计算但做空值保护（元素不存在时不报错）
  if (hint) {
    hint.innerHTML = state.date
      ? `当前：<b>${state.date}</b>（${dayLabel(state.date)}）· 取数：${srcLabel}` +
        (pd && pd !== state.date ? ` · 平台指标取 <b>${pd}</b>（前一天）` : '') +
        (miss.length ? ' · ' + miss.join(' · ') : '')
      : '暂无可用日期';
  }
}

function render() {
  renderStatus();
  renderKpis();
  renderTable();
  const dEl = $('report-date');
  if (dEl && state.date) dEl.value = state.date;
}

/* =====================================================================
 * 八、日期范围与选择
 * ===================================================================== */
function allDates() {
  const set = new Set();
  state.orders.forEach((r) => { if (r.payTime) set.add(fmtDate(r.payTime)); });
  Object.keys(state.rows).forEach((d) => set.add(d));
  Object.keys(state.daily).forEach((d) => { if (Object.keys(state.daily[d] || {}).length) set.add(d); });
  Object.keys(state.promo).forEach((d) => set.add(d));
  return [...set].sort();
}
function syncDateRange() {
  const ds = allDates();
  const el = $('report-date');
  if (!ds.length) { el.removeAttribute('min'); el.removeAttribute('max'); return; }
  el.min = ds[0];
  el.max = ds[ds.length - 1];
}
function autoPickDate() {
  const ds = allDates();
  if (!ds.length) { state.date = ''; return; }
  if (state.date && ds.includes(state.date)) return;
  // 默认取最新一天
  state.date = ds[ds.length - 1];
}
function gotoLatest() {
  const ds = allDates();
  if (!ds.length) { toast('暂无可选日期，请先导入数据', 'warn'); return; }
  state.date = ds[ds.length - 1];
  render();
  toast('已跳到最新日期 ' + state.date, 'ok', 1600);
}

/* =====================================================================
 * 九、手工编辑
 *   ① 常态：双击数据格（不含店铺名）就地修改；
 *   ② 编辑模式（点「编辑」按钮）：单击任意格（含店铺名）即可修改。
 *   两者写入同一份覆盖表（manual / names），优先级最高，「清除修改」可还原。
 * ===================================================================== */
let _pendingEdit = null;   // 上一个格子 blur 触发重绘后，需要接着打开的新格子

function findEditTd(key) {
  const tbl = $('tbl');
  if (!tbl || !key) return null;
  return tbl.querySelector(`td[data-edit][data-store="${key.store}"][data-field="${key.field}"][data-which="${key.which}"]`);
}

function startEdit(td) {
  if (!td || td.querySelector('input')) return;
  const field = td.dataset.field;
  const storeKey = td.dataset.store;
  const which = td.dataset.which || 'today';
  const col = field === 'name' ? { k: 'name', type: 'text' } : COLS.find((c) => c.k === field);
  if (!col) return;
  const cur = field === 'name' ? storeName(storeByKey(storeKey)) : cellValue(storeKey, field, which).v;
  const raw = (cur === null || cur === undefined) ? '' : String(cur);
  const inp = document.createElement('input');
  inp.type = 'text';
  inp.value = raw;
  inp.style.cssText = 'width:96%;text-align:center;border:1px solid var(--c-accent);border-radius:4px;padding:1px 3px;font-size:12px;font-family:inherit;';
  td.textContent = '';
  td.appendChild(inp);
  inp.focus(); inp.select();
  let done = false;
  const finish = (save) => {
    if (done) return;
    done = true;
    if (save) {
      const v = inp.value.trim();
      if (field === 'name') {
        // 店铺名：清空或改回原名即视为无覆盖
        const def = (storeByKey(storeKey) || {}).name || '';
        if (v === '' || v === def) delete state.names[storeKey];
        else state.names[storeKey] = v;
      } else {
        const mk = manualKey(state.date, storeKey, field);
        if (v === '') delete state.manual[mk];
        else {
          const textCol = col.type === 'text' || field === 'pm' || field === 'op';
          if (textCol) state.manual[mk] = v;
          else {
            const n = Number(v.replace(/,/g, ''));
            state.manual[mk] = isNaN(n) ? v : n;
          }
        }
      }
      saveLocal();
    }
    render();
    // 编辑模式下连续点选：上一个格子落库重绘后，接着打开这次点中的格子
    if (_pendingEdit) {
      const next = findEditTd(_pendingEdit);
      _pendingEdit = null;
      if (next) startEdit(next);
    }
  };
  inp.addEventListener('keydown', (ev) => {
    if (ev.key === 'Enter') finish(true);
    else if (ev.key === 'Escape') { _pendingEdit = null; finish(false); }
  });
  inp.addEventListener('blur', () => finish(true));
}

function bindEdit() {
  const tbl = $('tbl');
  if (tbl.dataset.editBound) return;
  // 编辑模式：单击即改（pointerdown 早于 blur，先记住目标格再落库重绘，避免这次点击被重绘吞掉）
  tbl.addEventListener('pointerdown', (e) => {
    if (!state.editMode) return;
    const td = e.target.closest && e.target.closest('td[data-edit]');
    if (!td) return;
    if (td.querySelector('input')) return;      // 点在正在编辑的输入框里
    const open = tbl.querySelector('td[data-edit] input');
    if (open) {
      // 必须先 preventDefault：否则这次 mousedown 会把焦点从旧输入框移走，
      // 旧框 blur→重绘之后，新开的输入框又被同一串事件抢走焦点，编辑会「点一下闪一下」。
      e.preventDefault();
      _pendingEdit = { store: td.dataset.store, field: td.dataset.field, which: td.dataset.which || 'today' };
      open.blur();                               // 触发 finish(true) → render() → 打开新格
      return;
    }
    e.preventDefault();
    startEdit(td);
  });
  // 常态：双击数据格（店铺名仅在编辑模式下可改）
  tbl.addEventListener('dblclick', (e) => {
    const td = e.target.closest && e.target.closest('td[data-edit]');
    if (!td) return;
    if (td.dataset.field === 'name' && !state.editMode) return;
    startEdit(td);
  });
  tbl.dataset.editBound = '1';
}

function updateEditHint() {
  const el = $('edit-hint');
  if (!el) return;
  el.textContent = state.editMode
    ? '编辑中：单击任意格（含店铺名）即可改，Enter 保存 / Esc 取消'
    : '双击数据格可改；点「编辑」后单击任意格（含店铺名）即可改';
  el.classList.toggle('on', !!state.editMode);
}

function toggleEditMode() {
  state.editMode = !state.editMode;
  document.body.classList.toggle('edit-mode', state.editMode);
  const b = $('btn-edit');
  if (b) {
    b.classList.toggle('on', state.editMode);
    b.textContent = state.editMode ? '✅ 完成编辑' : '✏️ 编辑';
  }
  if (!state.editMode) {
    const open = $('tbl') && $('tbl').querySelector('td[data-edit] input');
    if (open) open.blur();
  }
  updateEditHint();
  toast(state.editMode ? '编辑模式已开启：单击任意单元格（含店铺名）即可修改' : '已退出编辑模式', 'ok', 2000);
}

/* =====================================================================
 * 十、本地持久化（本页自己的数据）
 * ===================================================================== */
let _saveTimer = 0;
function saveLocal() {
  clearTimeout(_saveTimer);
  _saveTimer = setTimeout(() => {
    lsSet(LS.rows, JSON.stringify(state.rows));
    lsSet(LS.daily, JSON.stringify(state.daily));
    lsSet(LS.promo, JSON.stringify(state.promo));
    lsSet(LS.manual, JSON.stringify(state.manual));
    lsSet(LS.names, JSON.stringify(state.names));
    lsSet(LS.meta, JSON.stringify(state.meta));
  }, 120);
}
function restoreLocal() {
  try { const r = lsGet(LS.rows); if (r) state.rows = JSON.parse(r) || {}; } catch (e) {}
  try { const d = lsGet(LS.daily); if (d) state.daily = JSON.parse(d) || {}; } catch (e) {}
  try { const p = lsGet(LS.promo); if (p) state.promo = JSON.parse(p) || {}; } catch (e) {}
  try { const m = lsGet(LS.manual); if (m) state.manual = JSON.parse(m) || {}; } catch (e) {}
  try { const n = lsGet(LS.names); if (n) state.names = JSON.parse(n) || {}; } catch (e) {}
  try { const t = lsGet(LS.meta); if (t) state.meta = Object.assign(state.meta, JSON.parse(t) || {}); } catch (e) {}
}

/* =====================================================================
 * 十一、初始化与事件
 * ===================================================================== */
async function init() {
  restoreLocal();
  // 事件与首屏先就绪：本地存储（localStorage/IndexedDB）可能被其它标签页的长事务阻塞，
  // 若把绑定排在 await 之后，一旦读库卡住整页就完全无响应（2026-10-28 实测复现）。
  bindUI();
  bindEdit();
  updateEditHint();
  syncDateRange();
  autoPickDate();
  render();

  // 共用数据（订单 / VAB / 历史明细）异步补入，最多等 4s；超时也不影响页面可用
  await Promise.race([
    (async () => { await loadSharedOrders(); await loadSharedVab(); await loadSharedHist(); })(),
    new Promise((r) => setTimeout(r, 4000)),
  ]);
  syncDateRange();
  autoPickDate();
  _aggCache.key = '';
  _platCache.key = '';
  render();
}

function bindUI() {
  $('btn-orders').addEventListener('click', () => $('file-orders').click());
  $('btn-stores').addEventListener('click', () => $('file-stores').click());
  $('btn-platform').addEventListener('click', () => $('file-platform').click());
  $('btn-promo').addEventListener('click', () => $('file-promo').click());
  $('file-orders').addEventListener('change', (e) => { onFileOrders(e.target.files && e.target.files[0]); e.target.value = ''; });
  $('file-stores').addEventListener('change', (e) => { onFileStores(e.target.files && e.target.files[0]); e.target.value = ''; });
  // 平台三表：一次可选多个文件（客流 / 在线咨询 / 门店交易任意组合）
  $('file-platform').addEventListener('change', (e) => {
    const files = e.target.files ? Array.from(e.target.files) : [];
    e.target.value = '';
    if (files.length) onFilesPlatform(files);
  });
  $('file-promo').addEventListener('change', (e) => { onFilePromo(e.target.files && e.target.files[0]); e.target.value = ''; });
  $('btn-latest').addEventListener('click', gotoLatest);

  $('report-date').addEventListener('change', (e) => {
    const v = e.target.value;
    if (!v) return;
    state.date = v;
    _aggCache.key = '';
    _platCache.key = '';
    render();
  });

  $('period-seg').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-p]');
    if (!b) return;
    state.period = b.dataset.p;
    [...$('period-seg').querySelectorAll('button')].forEach((x) => x.classList.toggle('active', x === b));
    renderTable();
  });
  $('src-seg').addEventListener('click', (e) => {
    const b = e.target.closest('button[data-s]');
    if (!b) return;
    state.source = b.dataset.s;
    [...$('src-seg').querySelectorAll('button')].forEach((x) => x.classList.toggle('active', x === b));
    if (state.source === 'auto' && !state.orders.length) {
      toast('尚未导入订单表，「自动计算」暂无数据，已回落到源表值', 'warn', 4200);
    }
    render();
  });

  $('btn-clear').addEventListener('click', () => {
    if (!confirm('将清除本页的「七大店铺总表」解析结果、平台三表数据与手工覆盖值（不会删除订单数据）。继续？')) return;
    state.rows = {}; state.daily = {}; state.promo = {}; state.manual = {}; state.names = {}; state.meta = {};
    lsSet(LS.rows, '{}'); lsSet(LS.daily, '{}'); lsSet(LS.promo, '{}'); lsSet(LS.manual, '{}'); lsSet(LS.names, '{}'); lsSet(LS.meta, '{}');
    _aggCache.key = ''; _platCache.key = '';
    syncDateRange(); autoPickDate(); render();
    toast('本页数据已清除', 'ok');
  });
  $('btn-edit').addEventListener('click', toggleEditMode);
  $('btn-clear-edit').addEventListener('click', () => {
    const nManual = Object.keys(state.manual).length;
    const nNames = Object.keys(state.names).length;
    if (!nManual && !nNames) { toast('当前没有手动修改', 'warn'); return; }
    if (!confirm('将清除本页所有手工修改（双击或编辑模式下改过的单元格值，含店铺名）。继续？')) return;
    state.manual = {}; state.names = {};
    lsSet(LS.manual, '{}'); lsSet(LS.names, '{}');
    _aggCache.key = ''; _platCache.key = '';
    render();
    toast('已清除 ' + nManual + ' 处单元格修改' + (nNames ? '、' + nNames + ' 处店铺名' : ''), 'ok');
  });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
else init();
