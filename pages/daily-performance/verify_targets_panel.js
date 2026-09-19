/* 验证 ② 月目标面板美化 + 高新/北郊无目标 */
const fs = require('fs');
const path = require('path');
const src = fs.readFileSync(path.join(__dirname, 'assets/app.js'), 'utf8');

let pass = 0, fail = 0;
function check(name, cond, extra='') { if(cond){pass++; console.log('  PASS '+name);} else {fail++; console.log('  FAIL '+name+(extra?'  ['+extra+']':''));} }

// 1) 提取 DEFAULT_TARGETS 对象字面量
const mDef = src.match(/const DEFAULT_TARGETS\s*=\s*\{([\s\S]*?)\n\};/);
if(!mDef){ console.error('无法定位 DEFAULT_TARGETS'); process.exit(1); }
const DEFAULT_TARGETS = eval('({' + mDef[1] + '})');
check('DEFAULT_TARGETS 已移除「北郊」', !('北郊' in DEFAULT_TARGETS), 'keys='+Object.keys(DEFAULT_TARGETS).join(','));
check('DEFAULT_TARGETS 已移除「高新」', !('高新' in DEFAULT_TARGETS));

// 2) 提取 NO_TARGET_KEYS
const mNo = src.match(/const NO_TARGET_KEYS\s*=\s*new Set\(\[([\s\S]*?)\]\)/);
const NO_TARGET_KEYS = new Set((mNo ? mNo[1] : '').split(',').map(s=>s.trim().replace(/['"]/g,'')).filter(Boolean));
check('NO_TARGET_KEYS 含「高新」', NO_TARGET_KEYS.has('高新'));
check('NO_TARGET_KEYS 含「北郊」', NO_TARGET_KEYS.has('北郊'));

// 3) 提取 TARGET_GROUPS
const mGrp = src.match(/const TARGET_GROUPS\s*=\s*(\[[\s\S]*?\n\];)/);
const TARGET_GROUPS = eval(mGrp[1].replace(/;$/,''));
check('TARGET_GROUPS 存在且为数组', Array.isArray(TARGET_GROUPS) && TARGET_GROUPS.length>=2, 'len='+(TARGET_GROUPS&&TARGET_GROUPS.length));
const groupedKeys = TARGET_GROUPS.flatMap(g=>g.keys);
const missing = Object.keys(DEFAULT_TARGETS).filter(k=>!NO_TARGET_KEYS.has(k) && !groupedKeys.includes(k));
check('所有有目标的部门均已分组（无遗漏）', missing.length===0, 'missing='+missing.join(','));
const leaked = groupedKeys.filter(k=>NO_TARGET_KEYS.has(k));
check('分组中不含无目标部门', leaked.length===0, 'leaked='+leaked.join(','));

// 4) renderTargets 必须产出分组卡片结构
check('renderTargets 使用 .target-group 分组', /target-group/.test(src));
check('renderTargets 逐条 .target-row 渲染', /target-row/.test(src));
check('renderTargets 仍保留 data-key 绑定', /data-key="\$\{k\}"/.test(src));
check('renderTargets 仍绑定 input 即时写回', /persistTargets\(\)/.test(src) && /renderReport\(\)/.test(src));

// 5) index.html 美化 CSS 与说明
const html = fs.readFileSync(path.join(__dirname,'index.html'),'utf8');
check('index.html 含 .target-group 美化样式', /\.target-group\b/.test(html));
check('index.html 含 .target-row 样式', /\.target-row\b/.test(html));
check('index.html 含高新/北郊无目标说明', /高新店 \/ 经开店 不设月度目标/.test(html));
check('index.html 已去掉 grid-2 残留用法', !/id="targets" class="grid-2"/.test(html));

console.log((fail===0?'\nALL_PASS':'\nFAILED')+' ('+(pass+fail)+' checks)');
process.exit(fail===0?0:1);
