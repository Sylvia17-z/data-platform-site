// verify_colwidth.js — value-level check of the real column-width constants in app.js
// Reads the actual FRONT_COL_PCT_5 / DAILY_COL_PX from source and asserts the
// width-assignment formula used by setupColgroup() yields the expected strings.
const fs = require('fs');
const path = require('path');

const appJs = path.join(__dirname, 'assets', 'app.js');
const src = fs.readFileSync(appJs, 'utf8');

function extractConst(name) {
  const re = new RegExp('const\\s+' + name + '\\s*=\\s*([^;]+);');
  const m = src.match(re);
  if (!m) throw new Error('const ' + name + ' not found in app.js');
  return m[1].trim();
}

const pctRaw = extractConst('FRONT_COL_PCT_5');
const pxRaw = extractConst('DAILY_COL_PX');
const FRONT_COL_PCT_5 = eval(pctRaw);
const DAILY_COL_PX = eval(pxRaw);

// Replicate the exact branch from setupColgroup() (app.js ~947-953)
function colWidth(i, state) {
  if (state.colWidths && state.colWidths[i]) return state.colWidths[i];
  if (i < FRONT_COL_PCT_5.length) return FRONT_COL_PCT_5[i] + '%';
  return DAILY_COL_PX + 'px';
}

const checks = [];
function check(name, cond, detail) {
  checks.push({ name, pass: !!cond, detail });
}

// 1. constants match the widened values
check('FRONT_COL_PCT_5 widened', JSON.stringify(FRONT_COL_PCT_5) === '[7.5,11,12.5,10.5,21]',
  'got ' + JSON.stringify(FRONT_COL_PCT_5));
check('DAILY_COL_PX widened', DAILY_COL_PX === 90, 'got ' + DAILY_COL_PX);
check('front 5 sum <=100', FRONT_COL_PCT_5.reduce((a, b) => a + b, 0) <= 100,
  'sum=' + FRONT_COL_PCT_5.reduce((a, b) => a + b, 0));

// 2. width formula for a sample layout: 5 front + 3 daily = 8 cols
const state = {};
const sample = [0,1,2,3,4,5,6,7].map(i => colWidth(i, state));
check('front col0 = 7.5%', sample[0] === '7.5%', sample[0]);
check('front col1 = 11%', sample[1] === '11%', sample[1]);
check('front col4 = 21%', sample[4] === '21%', sample[4]);
check('daily col5 = 90px', sample[5] === '90px', sample[5]);
check('daily col7 = 90px', sample[7] === '90px', sample[7]);

// 3. manual colWidth override still takes precedence (persisted widths)
const state2 = { colWidths: { 2: '15%' } };
check('manual override wins', colWidth(2, state2) === '15%', colWidth(2, state2));

const failed = checks.filter(c => !c.pass);
checks.forEach(c => console.log((c.pass ? 'PASS' : 'FAIL') + ' | ' + c.name + ' | ' + c.detail));
console.log('\n' + (failed.length ? ('FAILED ' + failed.length) : 'ALL_PASS') +
  ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
