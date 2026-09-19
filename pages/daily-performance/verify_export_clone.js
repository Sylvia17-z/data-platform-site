// verify_export_clone.js — real jsdom render of buildExportClone() to confirm the
// v-75 export width behavior: front cols lock to px (jsdom has no layout -> measured
// widths are 0 -> falls back to colgroup % scaled by the stubbed offsetWidth), and the
// single surviving daily col is locked to its normal width (90px) instead of absorbing
// the leftover width of the width:100% page table.
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const proj = 'D:/workbuddy/2026-08-15-10-10-03/daily-performance';
const appSrc = fs.readFileSync(path.join(proj, 'assets', 'app.js'), 'utf8');

const html = `<!DOCTYPE html><html><body>
<div class="report-stage" style="overflow-x:auto;max-width:100%">
  <table class="report" id="report-table" style="width:100%;border-collapse:collapse;table-layout:auto;">
    <colgroup>
      <col style="width:7.5%"><col style="width:11%"><col style="width:12.5%"><col style="width:10.5%"><col style="width:21%"><col style="width:90px"><col style="width:90px"><col style="width:60px"><col style="width:46px">
    </colgroup>
    <thead>
      <tr class="title-row"><th class="title-cell" colspan="8">日业绩跟踪表</th></tr>
      <tr class="head-row">
        <th class="cat" colspan="3">部门类别</th>
        <th class="target">月目标</th>
        <th>月累计</th>
        <th class="daily-col" data-idx="0" data-date="2026-09-09">9/9</th>
        <th class="daily-col" data-idx="1" data-date="2026-09-10">9/10</th>
        <th id="batch-toggle" style="width:60px">批量</th>
        <th id="expand-all" style="width:46px">展开</th>
      </tr>
    </thead>
    <tbody>
      <tr><td class="cat" colspan="3">合计</td><td class="target">1000</td><td>500</td><td>200</td><td>300</td><td></td><td></td></tr>
    </tbody>
  </table>
</div>
</body></html>`;

const dom = new JSDOM(html, { runScripts: 'outside-only', pretendToBeVisual: true });
const { window } = dom;
const document = window.document;

// Stub the globals app.js touches at top level / in buildExportClone.
window.localStorage = { getItem: () => null, setItem: () => {}, removeItem: () => {} };
const noop = () => {};
window.showToast = noop; window.downloadBlob = noop; window.refreshDebug = noop;
window.todayIso = () => '2026-09-10';
window.state = { cutoffDate: '2026-09-10', results: { ok: true }, manual: { month: {}, daily: {} }, colWidths: {} };

// Expose buildExportClone by evaluating app.js in this window context, then grabbing the function.
// app.js defines buildExportClone as a top-level function declaration -> becomes window.buildExportClone.
const runner = new window.Function(appSrc + '\nreturn (typeof buildExportClone==="function")?buildExportClone:null;');
const buildExportClone = runner.call(window);
if (!buildExportClone) { console.error('FAIL | buildExportClone not found after eval'); process.exit(1); }

const table = document.getElementById('report-table');
// Simulate a wide on-screen render: give the table an offsetWidth (jsdom returns 0 otherwise).
// Cell offsetWidth stays 0 in jsdom -> measureLiveColWidths() yields 0 -> colgroup % fallback.
Object.defineProperty(table, 'offsetWidth', { value: 1400, configurable: true });

const clone = buildExportClone(table);
const cloneWidth = clone.style.width;
const cloneCols = [...clone.querySelectorAll('colgroup col')];

const checks = [];
const check = (n, c, d) => checks.push({ n, c: !!c, d });

// v-75: front cols locked to px (colgroup % x 1400 fallback), daily col = 90px.
// front = 105+154+175+147+294 = 875; total = 875 + 90 = 965.
check('clone width = front(875) + daily(90) = 965px', cloneWidth === '965px', cloneWidth);
// colgroup trimmed to 5 front + 1 daily (hidden 9/9, batch 60px, expand 46px removed) = 6
check('colgroup keeps front+latest-daily cols (6)', cloneCols.length === 6, 'cols=' + cloneCols.length);
check('front col[0] locked to 105px (7.5%x1400)', cloneCols[0] && cloneCols[0].style.width === '105px', cloneCols[0] && cloneCols[0].style.width);
check('front col[4] locked to 294px (21%x1400)', cloneCols[4] && cloneCols[4].style.width === '294px', cloneCols[4] && cloneCols[4].style.width);
check('daily col locked to 90px (not stretched)', cloneCols[5] && cloneCols[5].style.width === '90px', cloneCols[5] && cloneCols[5].style.width);
// batch/expand th removed from clone
check('batch-toggle removed', !clone.querySelector('#batch-toggle'), 'present=' + !!clone.querySelector('#batch-toggle'));
check('expand-all removed', !clone.querySelector('#expand-all'), 'present=' + !!clone.querySelector('#expand-all'));
// hidden 9/9 removed, latest 9/10 kept
check('non-latest daily 9/9 removed', !clone.querySelector('th.daily-col[data-date="2026-09-09"]'), 'present=' + !!clone.querySelector('th.daily-col[data-date="2026-09-09"]'));
check('latest daily 9/10 kept', !!clone.querySelector('th.daily-col[data-date="2026-09-10"]'), 'present=' + !!clone.querySelector('th.daily-col[data-date="2026-09-10"]'));
// title colspan recomputed to baseCols(5)+visibleDaily(1)=6
const titleColspan = clone.querySelector('.title-cell').getAttribute('colspan');
check('title colspan = 6', titleColspan === '6', 'colspan=' + titleColspan);

const failed = checks.filter(c => !c.c);
checks.forEach(c => console.log((c.c ? 'PASS' : 'FAIL') + ' | ' + c.n + ' | ' + c.d));
console.log('\n' + (failed.length ? ('FAILED ' + failed.length) : 'ALL_PASS') + ' (' + checks.length + ' checks)');
process.exit(failed.length ? 1 : 0);
