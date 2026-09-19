
(function(){
  "use strict";
  var XLSX = window.XLSX, JSZip = window.JSZip, htmlToImage = window.htmlToImage;

  // 统计说明（三报表平台：页面与导出件共用）
  var NOTE_TEXT = "统计说明：数据来源 领健系统每周一至周六数据（截止时间：周六 12:00）\n1、初诊客量~周区间内到访类型为初诊 复诊客量~周区间内到访类型为复诊/疗程内/再消费/复查/为空\n2、成交率 现款支付不为0并且不是退款的消费客量 占总客量的比例\n3、如果一个客户 周一来体检 周三来解读报告 计算客量为1";

  // 预约量统计模块专用统计说明（按所选统计日，不分周区间）
  var NEXT_NOTE_TEXT = "统计说明：\n1、数据来源 领健系统预约列表\n2、初诊预约客量~预约类型为初诊 复诊预约客量~预约类型为复诊/疗程内/再消费/复查/为空";

  var APPT = [];   // {date,visit,mem,type,source,butler,doctor}
  var ORDER = [];  // {date,mem,consultant,deptSrc,cash,employee}
  var RANGE = {start:null, end:null};
  var APP_VERSION = "v2.7 改为两大独立模块：三报表平台(订单+预约)与预约量统计(仅预约)，顶部按钮切换 2026-09-16";
  var OV = { perf:{}, butler:{} };   // 手动调整覆盖：perf[日期]={gao,jing}；butler[姓名]={perf}

  // 手工调整持久化：按数据区间 key 存入 localStorage，刷新/重新导入同区间数据后仍在
  function ovPeriod(){ return (RANGE.start && RANGE.end) ? (RANGE.start + "~" + RANGE.end) : ""; }
  function saveOV(){
    try { localStorage.setItem("stats_ov", JSON.stringify({ period: ovPeriod(), ov: OV })); } catch(e){}
  }
  function restoreOV(){
    try {
      var raw = localStorage.getItem("stats_ov");
      if(!raw){ OV = { perf:{}, butler:{} }; return; }
      var obj = JSON.parse(raw);
      if(obj && obj.period === ovPeriod() && obj.ov){ OV = obj.ov; }
      else { OV = { perf:{}, butler:{} }; }
    } catch(e){ OV = { perf:{}, butler:{} }; }
  }
  function clearOV(){
    var n = Object.keys(OV.perf||{}).length + Object.keys(OV.butler||{}).length;
    OV = { perf:{}, butler:{} };
    try { localStorage.removeItem("stats_ov"); } catch(e){}
    recompute();
    if(n>0){
      status.innerHTML = '<b style="color:#0a7a3a">✅ 已清除全部手工调整（共 '+n+' 项）</b>，已按真实数据重算。 · ' + status.innerHTML;
    }else{
      status.innerHTML = '<b style="color:#8a5a00">ℹ️ 当前没有手工调整记录</b>，管家业绩直接来自订单真实计算（非覆盖值）。若仍显示异常，多为浏览器缓存了旧页面，请硬刷新（Ctrl/Cmd+Shift+R）。 · ' + status.innerHTML;
    }
  }
  var HI = ['南媛','系统管理员'];

  // ---------- 工具 ----------
  function findCol(headers, keys){
    var best=-1, bestScore=0;
    for(var i=0;i<headers.length;i++){
      var h = headers[i]!=null ? String(headers[i]) : "";
      var score=0; for(var k=0;k<keys.length;k++){ if(h.indexOf(keys[k])>=0) score++; }
      if(score>bestScore){ bestScore=score; best=i; }
    }
    return bestScore>0 ? best : -1;
  }
  function parseDateVal(v){
    if(v==null || v==="") return null;
    if(v instanceof Date) return v.getFullYear()+"-"+pad(v.getMonth()+1)+"-"+pad(v.getDate());
    var s = String(v).replace(/\//g,"-").replace("T"," ").trim();
    var m = s.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    return m ? (m[1]+"-"+pad(+m[2])+"-"+pad(+m[3])) : null;
  }
  function pad(n){ return (n<10?"0":"")+n; }
  function fmtMD(d){ if(!d) return ""; var p=d.split("-"); return +p[1]+"."+ +p[2]; }
  function fmtSlash(d){ if(!d) return ""; var p=d.split("-"); return +p[1]+"/"+ +p[2]; }
  function addDays(str,n){ if(!str) return null; var p=str.split("-"); var d=new Date(+p[0],+p[1]-1,+p[2]); d.setDate(d.getDate()+n); return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
  function todayStr(){ var d=new Date(); return d.getFullYear()+"-"+pad(d.getMonth()+1)+"-"+pad(d.getDate()); }
  var WD = ['周日','周一','周二','周三','周四','周五','周六'];
  function weekday(d){ if(!d) return ""; var p=d.split("-"); return WD[new Date(+p[0],+p[1]-1,+p[2]).getDay()]; }
  function escapeHtml(s){ return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); }
  function money(x){ return (Math.round(x*100)/100).toFixed(2); }
  function pctStr(r){ return (r*100).toFixed(0)+"%"; }
  function deptAppt(src){
    var s=String(src||"");
    if(s.indexOf("营销部")>=0) return "营销部";
    if(s.indexOf("市场部")>=0 || s.indexOf("商务部")>=0) return "商务部";
    return "其他(公司所有)";
  }
  function deptOrder(src){
    var s=String(src||"");
    if(s.indexOf("营销部")>=0) return "营销部";
    if(s.indexOf("市场部")>=0 || s.indexOf("商务部")>=0) return "商务部";
    return "其他(公司所有)";
  }

  // ---------- 解析 ----------
  // 在工作簿的「所有 sheet」×「每个 sheet 的前几行」里寻找关键列命中最全的一组作为表头。
  // 背景：领健导出结构不固定——如 预约列表-2026-09-05.xlsx 的第一个 sheet 是透视表（Sheet4），
  // 明细在第二个 sheet（名字是日期、不含"预约"二字），且表头前可能有空行/标题行。
  // colDefs: [{key, names:[关键词], req:是否必需}]；返回 {ws, ci, hrow, hit} 或 null
  function scanSheets(wb, colDefs, maxScan){
    var best=null;
    for(var s=0;s<wb.SheetNames.length;s++){
      var name=wb.SheetNames[s], ws=wb.Sheets[name];
      if(!ws) continue;
      var aoa;
      try{ aoa=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true}); }catch(e){ continue; }
      for(var i=0;i<Math.min(aoa.length,maxScan||6);i++){
        var row=aoa[i]; if(!row) continue;
        var H=row.map(function(c){return c==null?"":String(c);});
        var ci={},hit=0,reqFail=false;
        colDefs.forEach(function(d){
          var idx=findCol(H,d.names);
          ci[d.key]=idx;
          if(idx>=0) hit++;
          else if(d.req) reqFail=true;
        });
        if(reqFail) continue;
        if(!best||hit>best.hit){ best={ws:ws,ci:ci,hrow:i,hit:hit}; }
      }
    }
    return best;
  }

  function parseAppt(wb){
    var b=scanSheets(wb,[
      {key:"date",names:["日期"]},
      {key:"visit",names:["是否到访"],req:true},
      {key:"mem",names:["会员号"],req:true},
      {key:"type",names:["预约类型"]},
      {key:"source",names:["客户来源"],req:true},
      {key:"butler",names:["预约健康管家","管家","健康管家","服务管家"]},
      {key:"doctor",names:["医生","预约医生","接诊医生","主诊医生","医生姓名","责任医生"]}
    ]);
    if(!b) throw new Error("预约列表缺少关键列（会员号/是否到访/客户来源）");
    var aoa=XLSX.utils.sheet_to_json(b.ws,{header:1,defval:null,raw:true});
    var ci=b.ci, out=[];
    for(var r=b.hrow+1;r<aoa.length;r++){
      var row=aoa[r]; if(!row) continue;
      var get=function(i){ return (i>=0&&i<row.length)?row[i]:null; };
      var mem=get(ci.mem); if(mem==null||String(mem).trim()==="") continue;
      out.push({ date:parseDateVal(get(ci.date)), visit:String(get(ci.visit)||"").trim(),
        mem:String(mem).trim(), type:String(get(ci.type)||"").trim(),
        source:String(get(ci.source)||"").trim(), butler:String(get(ci.butler)||"").trim(),
        doctor:String(get(ci.doctor)||"").trim() });
    }
    return out;
  }
  function parseOrder(wb){
    var b=scanSheets(wb,[
      {key:"mem",names:["会员号"],req:true},
      {key:"consultant",names:["开单咨询师"],req:true},
      {key:"cash",names:["现款支付"],req:true},
      {key:"dept",names:["客户来源渠道分类","渠道分类"]},
      {key:"time",names:["收款时间"]},
      {key:"emp",names:["收款员工"]},
      {key:"type",names:["收款类型"]}
    ]);
    if(!b) throw new Error("订单明细缺少关键列（会员号/开单咨询师/现款支付）");
    var aoa=XLSX.utils.sheet_to_json(b.ws,{header:1,defval:null,raw:true});
    var ci=b.ci, out=[];
    for(var r=b.hrow+1;r<aoa.length;r++){
      var row=aoa[r]; if(!row) continue;
      var get=function(i){ return (i>=0&&i<row.length)?row[i]:null; };
      var mem=get(ci.mem); if(mem==null||String(mem).trim()==="") continue;
      var cash=get(ci.cash); cash=(typeof cash==="number")?cash:(parseFloat(cash)||0);
      out.push({ date:parseDateVal(get(ci.time)), mem:String(mem).trim(),
        consultant:String(get(ci.consultant)||"").trim(),
        deptSrc:String(get(ci.dept)||"").trim(), cash:cash,
        employee:String(get(ci.emp)||"").trim(),
        type:(ci.type>=0?String(get(ci.type)||""):"").trim() });
    }
    return out;
  }

  // ---------- 计算 ----------
  function inRange(d){ return (!RANGE.start || (d&&d>=RANGE.start)) && (!RANGE.end || (d&&d<=RANGE.end)); }

  function compute(){
    var depts=["营销部","商务部","其他(公司所有)"];
    var kv={}; depts.forEach(function(d){ kv[d]={tot:new Set(),chu:new Set(),cash:new Set()}; });
    APPT.forEach(function(r){
      if(r.visit!=="是"||!inRange(r.date)) return;
      var d=deptAppt(r.source); if(!kv[d]) return;
      kv[d].tot.add(r.mem);
      if(r.type==="初诊") kv[d].chu.add(r.mem);
    });
    ORDER.forEach(function(r){
      if(r.cash===0||r.type==="退款"||!inRange(r.date)) return;
      var d=deptOrder(r.deptSrc); if(!kv[d]) return;
      kv[d].cash.add(r.mem);
    });
    var kvR={};
    depts.forEach(function(d){
      var tot=kv[d].tot.size, chu=kv[d].chu.size, cash=kv[d].cash.size;
      kvR[d]={tot:tot,chu:chu,fu:tot-chu,cash:cash,rate:tot?cash/tot:0};
    });
    var gt=0,gc=0,gf=0,gk=0;
    depts.forEach(function(d){ gt+=kvR[d].tot;gc+=kvR[d].chu;gf+=kvR[d].fu;gk+=kvR[d].cash; });
    kvR["合计"]={tot:gt,chu:gc,fu:gf,cash:gk,rate:gt?gk/gt:0};

    // 周业绩
    var gaoDay={},jingDay={},visitDay={},bookDay={};
    ORDER.forEach(function(r){
      if(!inRange(r.date)) return;
      var add=Math.round(r.cash*100)/100;
      if(HI.indexOf(r.employee)>=0) gaoDay[r.date]=Math.round(((gaoDay[r.date]||0)+add)*100)/100;
      else jingDay[r.date]=Math.round(((jingDay[r.date]||0)+add)*100)/100;
    });
    APPT.forEach(function(r){
      if(!inRange(r.date)||!r.mem) return;
      (bookDay[r.date]=bookDay[r.date]||new Set()).add(r.mem);
      if(r.visit==="是") (visitDay[r.date]=visitDay[r.date]||new Set()).add(r.mem);
    });
    var daySet={}, k;
    [gaoDay,jingDay,visitDay,bookDay].forEach(function(o){ for(k in o) daySet[k]=1; });
    var days=Object.keys(daySet).sort();
    var perfRows=days.map(function(d){
      var g=gaoDay[d]||0, j=jingDay[d]||0;
      var va=(visitDay[d]||new Set()).size, vb=(bookDay[d]||new Set()).size;
      return {date:d,gao:g,jing:j,total:g+j,rank:0,visit:va,book:vb,rate:vb?va/vb:0};
    });
    var sorted=perfRows.slice().sort(function(a,b){return b.total-a.total;});
    sorted.forEach(function(r,i){ for(var x=0;x<perfRows.length;x++){ if(perfRows[x].date===r.date){ perfRows[x].rank=i+1; break; } } });
    var wkVisit=new Set(), wkBook=new Set();
    APPT.forEach(function(r){ if(!inRange(r.date)||!r.mem) return; wkBook.add(r.mem); if(r.visit==="是") wkVisit.add(r.mem); });
    var pSum={gao:0,jing:0,total:0};
    perfRows.forEach(function(r){ pSum.gao+=r.gao; pSum.jing+=r.jing; pSum.total+=r.total; });
    var maxDay="",minDay="",maxV=-1,minV=1e18;
    perfRows.forEach(function(r){ if(r.total>maxV){maxV=r.total;maxDay=r.date;} if(r.total<minV){minV=r.total;minDay=r.date;} });
    var avg = perfRows.length ? pSum.total/perfRows.length : 0;

    // 管家：纯按"开单咨询师"分组；仅王勤红/空咨询师归入"无"
    var perfByName={};
    ORDER.forEach(function(r){
      if(!inRange(r.date)) return;
      var add=Math.round(r.cash*100)/100;
      var c=r.consultant||"";
      var key=(c==="王勤红" || c==="营销部" || c==="")?"无":c;
      perfByName[key]=Math.round(((perfByName[key]||0)+add)*100)/100;
    });
    var bBook={},bVisit={};
    APPT.forEach(function(r){
      if(!inRange(r.date)||!r.mem) return; if(!r.butler) return;
      (bBook[r.butler]=bBook[r.butler]||new Set()).add(r.mem);
      if(r.visit==="是") (bVisit[r.butler]=bVisit[r.butler]||new Set()).add(r.mem);
    });
    var named=Object.keys(perfByName).filter(function(n){return n!=="无";}).sort(function(a,b){return perfByName[b]-perfByName[a];});
    var namedBook=0,namedVisit=0;
    named.forEach(function(n){ namedBook+=(bBook[n]?bBook[n].size:0); namedVisit+=(bVisit[n]?bVisit[n].size:0); });
    var noBook=Math.max(0,wkBook.size-namedBook), noVisit=Math.max(0,wkVisit.size-namedVisit);
    var butlerRows=named.map(function(n){
      var bk=bBook[n]?bBook[n].size:0, vt=bVisit[n]?bVisit[n].size:0;
      return {name:n,perf:perfByName[n],rank:0,book:bk,visit:vt,rate: bk? vt/bk : null};
    });
    butlerRows.forEach(function(r,i){ r.rank=i+1; });
    butlerRows.push({name:"无",perf:perfByName["无"]||0,rank:"/",book:noBook,visit:noVisit,rate: noBook? noVisit/noBook : null});
    var bSumP=0,bSumB=0,bSumV=0;
    butlerRows.forEach(function(r){ bSumP+=r.perf; bSumB+=r.book; bSumV+=r.visit; });

    return { kv:kvR, depts:depts,
      perf:{rows:perfRows, sum:pSum, maxDay:maxDay, minDay:minDay, avg:avg, wkVisit:wkVisit.size, wkBook:wkBook.size},
      butler:{rows:butlerRows, sumP:bSumP, sumB:bSumB, sumV:bSumV} };
  }

  // ---------- 渲染 ----------
  function render(R, periodText){
    document.getElementById("kvTitle").textContent="周客量统计(数据区间："+periodText+")";
    document.getElementById("perfTitle").textContent="周业绩统计(数据区间："+periodText+")";
    document.getElementById("butlerTitle").textContent="管家业绩统计(数据区间："+periodText+")";
    var noteEl=document.getElementById("reportNote");
    if(noteEl) noteEl.innerHTML=NOTE_TEXT;

    // 周客量
    var kv=R.kv, d=R.depts;
    var kvHtml =
      '<thead><tr><th colspan="2">部门</th><th>总到访客量</th><th>初诊客量</th><th>复诊客量</th><th>有成交客量</th><th>成交率</th></tr></thead>'+
      '<tbody>'+
      '<tr><td class="ch" rowspan="1">线上</td><td class="dept">营销部</td>'+
        '<td>'+kv[d[0]].tot+'</td><td>'+kv[d[0]].chu+'</td><td>'+kv[d[0]].fu+'</td><td>'+kv[d[0]].cash+'</td><td class="pct">'+pctStr(kv[d[0]].rate)+'</td></tr>'+
      '<tr><td class="ch" rowspan="2">线下</td><td class="dept">商务部</td>'+
        '<td>'+kv[d[1]].tot+'</td><td>'+kv[d[1]].chu+'</td><td>'+kv[d[1]].fu+'</td><td>'+kv[d[1]].cash+'</td><td class="pct">'+pctStr(kv[d[1]].rate)+'</td></tr>'+
      '<tr><td class="dept">其他(公司所有)</td>'+
        '<td>'+kv[d[2]].tot+'</td><td>'+kv[d[2]].chu+'</td><td>'+kv[d[2]].fu+'</td><td>'+kv[d[2]].cash+'</td><td class="pct">'+pctStr(kv[d[2]].rate)+'</td></tr>'+
      '<tr class="total"><td colspan="2">合计</td>'+
        '<td>'+kv["合计"].tot+'</td><td>'+kv["合计"].chu+'</td><td>'+kv["合计"].fu+'</td><td>'+kv["合计"].cash+'</td><td class="pct">'+pctStr(kv["合计"].rate)+'</td></tr>'+
      '</tbody>';
    document.getElementById("tblKv").innerHTML=kvHtml;

    // 周业绩
    var p=R.perf;
    var perfHtml='<thead><tr><th>日期</th><th>高新(元)</th><th>经开(元)</th><th>总业绩(元)<br>（高新+经开）</th><th>业绩排名</th><th>总到访客量</th><th>总预约客量</th><th>到访率</th></tr></thead><tbody>';
    p.rows.forEach(function(r){
      perfHtml+='<tr data-date="'+r.date+'"><td>'+r.date.replace(/-/g,"/")+'('+weekday(r.date)+')</td>'+
        '<td class="editable" contenteditable="true" data-table="perf" data-date="'+r.date+'" data-field="gao">'+money(r.gao)+'</td>'+
        '<td class="editable" contenteditable="true" data-table="perf" data-date="'+r.date+'" data-field="jing">'+money(r.jing)+'</td>'+
        '<td>'+money(r.total)+'</td>'+
        '<td>'+r.rank+'</td><td>'+r.visit+'</td><td>'+r.book+'</td><td class="pct">'+pctStr(r.rate)+'</td></tr>';
    });
    perfHtml+='<tr class="total"><td>合计</td><td>'+money(p.sum.gao)+'</td><td>'+money(p.sum.jing)+'</td><td>'+money(p.sum.total)+'</td>'+
      '<td>/</td><td>'+p.wkVisit+'</td><td>'+p.wkBook+'</td><td class="pct">'+pctStr(p.wkBook?p.wkVisit/p.wkBook:0)+'</td></tr>';
    perfHtml+='<tr class="summary"><td colspan="8">业绩最高值是'+fmtMD(p.maxDay)+'('+weekday(p.maxDay)+')　业绩最低值是'+fmtMD(p.minDay)+'('+weekday(p.minDay)+')　业绩平均值：'+money(p.avg)+'元</td></tr>';
    perfHtml+='</tbody>';
    document.getElementById("tblPerf").innerHTML=perfHtml;

    // 管家
    var b=R.butler;
    var bh='<thead><tr><th>姓名</th><th>总业绩(元)</th><th>业绩排名</th><th>总预约量</th><th>总到访量</th><th colspan="3">到访率</th></tr></thead><tbody>';
    b.rows.forEach(function(r){
      var rate = (r.rate==null) ? "/" : pctStr(r.rate);
      bh+='<tr data-name="'+r.name+'"><td>'+escapeHtml(r.name)+'</td>'+
        '<td class="editable" contenteditable="true" data-table="butler" data-name="'+r.name+'" data-field="perf">'+money(r.perf)+'</td>'+
        '<td>'+r.rank+'</td><td>'+r.book+'</td><td>'+r.visit+'</td><td colspan="3" class="pct">'+rate+'</td></tr>';
    });
    bh+='<tr class="total"><td>合计</td><td>'+money(b.sumP)+'</td><td>/</td><td>'+b.sumB+'</td><td>'+b.sumV+'</td><td colspan="3" class="pct">'+pctStr(b.sumB?b.sumV/b.sumB:0)+'</td></tr>';
    bh+='</tbody>';
    document.getElementById("tblButler").innerHTML=bh;
  }

  // ---------- 手动调整 ----------
  // 把 OV 覆盖应用到 compute() 结果，并重算所有联动指标（排名/合计/最高最低平均）
  function effectiveR(){
    var R=compute();
    R.perf.rows.forEach(function(r){
      var o=OV.perf[r.date];
      if(o){ if(o.gao!=null) r.gao=o.gao; if(o.jing!=null) r.jing=o.jing; r.total=r.gao+r.jing; }
    });
    var ps=R.perf.rows.slice().sort(function(a,b){return b.total-a.total;});
    ps.forEach(function(r,i){ for(var x=0;x<R.perf.rows.length;x++){ if(R.perf.rows[x].date===r.date){ R.perf.rows[x].rank=i+1; break; } } });
    var sg=0,sj=0,st=0; R.perf.rows.forEach(function(r){ sg+=r.gao; sj+=r.jing; st+=r.total; });
    R.perf.sum={gao:sg,jing:sj,total:st};
    var maxDay="",minDay="",maxV=-1,minV=1e18;
    R.perf.rows.forEach(function(r){ if(r.total>maxV){maxV=r.total;maxDay=r.date;} if(r.total<minV){minV=r.total;minDay=r.date;} });
    R.perf.avg=R.perf.rows.length?st/R.perf.rows.length:0; R.perf.maxDay=maxDay; R.perf.minDay=minDay;
    R.butler.rows.forEach(function(r){ var o=OV.butler[r.name]; if(o&&o.perf!=null) r.perf=o.perf; });
    var bs=R.butler.rows.filter(function(r){return r.name!=="无";}).slice().sort(function(a,b){return b.perf-a.perf;});
    bs.forEach(function(r,i){ for(var x=0;x<R.butler.rows.length;x++){ if(R.butler.rows[x].name===r.name){ R.butler.rows[x].rank=i+1; break; } } });
    var sp=0; R.butler.rows.forEach(function(r){ sp+=r.perf; }); R.butler.sumP=sp;
    return R;
  }
  function refreshDerived(){
    var R=effectiveR();
    var tblPerf=document.getElementById("tblPerf"); if(tblPerf){
      var tb=tblPerf.getElementsByTagName("tbody")[0];
      [].forEach.call(tb.rows,function(tr){
        var d=tr.getAttribute("data-date"); if(!d) return;
        var r=null; for(var i=0;i<R.perf.rows.length;i++){ if(R.perf.rows[i].date===d){ r=R.perf.rows[i]; break; } } if(!r) return;
        tr.cells[3].textContent=money(r.total);
        tr.cells[4].textContent=r.rank;
      });
      var tot=tblPerf.querySelector("tr.total");
      if(tot){ tot.cells[1].textContent=money(R.perf.sum.gao); tot.cells[2].textContent=money(R.perf.sum.jing);
        tot.cells[3].textContent=money(R.perf.sum.total); tot.cells[5].textContent=R.perf.wkVisit; tot.cells[6].textContent=R.perf.wkBook;
        tot.cells[7].textContent=pctStr(R.perf.wkBook?R.perf.wkVisit/R.perf.wkBook:0); }
      var sum=tblPerf.querySelector("tr.summary");
      if(sum){ sum.cells[0].textContent="业绩最高值是"+fmtMD(R.perf.maxDay)+"("+weekday(R.perf.maxDay)+")　业绩最低值是"+fmtMD(R.perf.minDay)+"("+weekday(R.perf.minDay)+")　业绩平均值："+money(R.perf.avg)+"元"; }
    }
    var tblB=document.getElementById("tblButler"); if(tblB){
      var tb2=tblB.getElementsByTagName("tbody")[0];
      [].forEach.call(tb2.rows,function(tr){
        var n=tr.getAttribute("data-name"); if(!n) return;
        var r=null; for(var i=0;i<R.butler.rows.length;i++){ if(R.butler.rows[i].name===n){ r=R.butler.rows[i]; break; } } if(!r) return;
        tr.cells[2].textContent=r.rank;
      });
      var bt=tblB.querySelector("tr.total");
      if(bt){ bt.cells[1].textContent=money(R.butler.sumP); }
    }
  }
  // 安全算式求值（不使用 eval）：递归下降解析，支持 + - * / ( ) 及 @（当前单元格原值）
  function calcExpr(expr, base){
    expr=String(expr==null?"":expr).replace(/@/g, String(Number(base)||0));
    if(!/^[\d\s.+*/()-]+$/.test(expr)) return NaN;
    var s=expr.replace(/\s+/g,"");
    if(s==="") return NaN;
    var pos=0;
    function parseExpr(){
      var v=parseTerm();
      while(pos<s.length && (s.charAt(pos)==="+"||s.charAt(pos)==="-")){
        var op=s.charAt(pos++); var r=parseTerm();
        v = op==="+" ? v+r : v-r;
      }
      return v;
    }
    function parseTerm(){
      var v=parseFactor();
      while(pos<s.length && (s.charAt(pos)==="*"||s.charAt(pos)==="/")){
        var op=s.charAt(pos++); var r=parseFactor();
        v = op==="*" ? v*r : (r===0?0:v/r);
      }
      return v;
    }
    function parseFactor(){
      if(pos>=s.length) return NaN;
      var ch=s.charAt(pos);
      if(ch==="+"){ var p=s.charAt(pos-1); if(p==="+"||p==="*"||p==="/") return NaN; pos++; return parseFactor(); }
      if(ch==="-"){ pos++; return -parseFactor(); }
      if(ch==="("){ pos++; var v=parseExpr(); if(s.charAt(pos)!==")") return NaN; pos++; return v; }
      var m=s.slice(pos).match(/^\d+\.?\d*/);
      if(!m) return NaN;
      pos+=m[0].length;
      return parseFloat(m[0]);
    }
    var result=parseExpr();
    if(pos<s.length) return NaN;
    return isFinite(result)?result:NaN;
  }
  // 取当前单元格有效数值（手改 @ 引用时使用）
  function cellBaseValue(c, table){
    var R=effectiveR(); if(!R) return 0;
    if(table==="perf"){
      var d=c.getAttribute("data-date"), f=c.getAttribute("data-field");
      for(var i=0;i<R.perf.rows.length;i++){ if(R.perf.rows[i].date===d) return Number(R.perf.rows[i][f])||0; }
    } else {
      var n=c.getAttribute("data-name");
      for(var j=0;j<R.butler.rows.length;j++){ if(R.butler.rows[j].name===n) return Number(R.butler.rows[j].perf)||0; }
    }
    return 0;
  }
  // 解析手改输入：返回 数字 / null(清空) / NaN(算式未完成，暂不更新)
  function computeEditValue(raw, c, table){
    raw=(raw||"").trim();
    if(raw==="") return null;
    if(raw.charAt(0)==="=" && /[+*/-]/.test(raw.slice(1))){
      var v=calcExpr(raw.slice(1), cellBaseValue(c, table));
      return isNaN(v)?NaN:v;
    }
    var ns=raw.replace(/[,\s¥￥]/g,"").replace(/[^\d.\-]/g,"");
    var n=parseFloat(ns); return isNaN(n)?0:n;
  }
  function onEditInput(e){
    var c=e.target; if(!c.getAttribute || !c.getAttribute("data-field")) return;
    var table=c.getAttribute("data-table");
    var raw=(c.textContent||"").trim();
    if(raw===""){
      if(table==="perf"){ var d=c.getAttribute("data-date"); if(OV.perf[d]){ delete OV.perf[d][c.getAttribute("data-field")]; if(Object.keys(OV.perf[d]).length===0) delete OV.perf[d]; } }
      else { var n=c.getAttribute("data-name"); delete OV.butler[n]; }
      refreshDerived(); return;
    }
    var ev;
    if(raw.charAt(0)==="="){                 // 算式：未含运算符或非法时暂不覆盖，避免误写
      if(!/[+*/-]/.test(raw.slice(1))) return;
      ev=calcExpr(raw.slice(1), cellBaseValue(c, table));
      if(isNaN(ev)) return;
    } else {
      var ns=raw.replace(/[,\s¥￥]/g,"").replace(/[^\d.\-]/g,"");
      ev=parseFloat(ns); if(isNaN(ev)) ev=0;
    }
    if(table==="perf"){
      var dd=c.getAttribute("data-date"), f=c.getAttribute("data-field");
      OV.perf[dd]=OV.perf[dd]||{}; OV.perf[dd][f]=ev;
    } else {
      var nn=c.getAttribute("data-name");
      OV.butler[nn]=OV.butler[nn]||{}; OV.butler[nn].perf=ev;
    }
    refreshDerived(); saveOV();
  }
  function onEditBlur(e){
    var c=e.target; if(!c.getAttribute || !c.getAttribute("data-field")) return;
    var table=c.getAttribute("data-table");
    var raw=(c.textContent||"").trim();
    if(raw.charAt(0)==="=" && /[+*/-]/.test(raw.slice(1))){
      var v=calcExpr(raw.slice(1), cellBaseValue(c, table));
      if(isNaN(v)) v=cellBaseValue(c, table);   // 算式非法 → 回退原值
      if(table==="perf"){ var d=c.getAttribute("data-date"), f=c.getAttribute("data-field"); OV.perf[d]=OV.perf[d]||{}; OV.perf[d][f]=v; }
      else { var n=c.getAttribute("data-name"); OV.butler[n]=OV.butler[n]||{}; OV.butler[n].perf=v; }
      saveOV();
      c.textContent=money(v);
      refreshDerived(); return;
    }
    var R=effectiveR();
    if(table==="perf"){
      var d2=c.getAttribute("data-date"), f2=c.getAttribute("data-field");
      var row=null; for(var i=0;i<R.perf.rows.length;i++){ if(R.perf.rows[i].date===d2){ row=R.perf.rows[i]; break; } }
      if(row) c.textContent=money(row[f2]);
    } else {
      var n2=c.getAttribute("data-name");
      var br=null; for(var j=0;j<R.butler.rows.length;j++){ if(R.butler.rows[j].name===n2){ br=R.butler.rows[j]; break; } }
      if(br) c.textContent=money(br.perf);
    }
    refreshDerived();
  }

  // ---------- 导出 Excel (手写 OOXML，还原模板合并+公式) ----------
  function buildXlsxZip(R, periodText){
    var HDR="FF01418B";
    var note=NOTE_TEXT;
    var kv=R.kv, d=R.depts, p=R.perf, b=R.butler;

    var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'+
    ' <numFmts count="1"><numFmt numFmtId="164" formatCode="0%"/></numFmts>\n'+
    ' <fonts count="5">\n'+
    '  <font><sz val="11"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><b/><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>\n'+
    ' </fonts>\n'+
    ' <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'+
    '  <fill><patternFill patternType="solid"><fgColor rgb="'+HDR+'"/><bgColor indexed="64"/></patternFill></fill></fills>\n'+
    ' <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>\n'+
    '  <border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>\n'+
    ' <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'+
    // 字体统一：全部 Microsoft YaHei 11 号；字重规则——标题行/表头行/合计行/业绩汇总行加粗，其余数据行不加粗
    // 样式索引：0默认 / 1标题·表头(加粗白字蓝底) / 2普通 / 3数据(居中不加粗) / 4说明(左对齐换行) /
    //            5百分比 / 6百分比(数据·不加粗) / 7百分比(合计·加粗) / 8合计(居中加粗) / 9业绩汇总(加粗左对齐换行)
    // 注意：count 必须与实际 <xf> 条数一致，否则末位索引的样式失效（会退回默认格式，百分比显示成小数）
    ' <cellXfs count="10">\n'+
    '  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'+
    '  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>\n'+
    '  <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\n'+
    '  <xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="164" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\n'+
    ' </cellXfs>\n'+
    ' <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\n'+
    '</styleSheet>';

    var rows=[]; var merges=[];
    var rowH={};   // 行号 -> 行高：合并单元格不会自动撑高，统计说明行/标题行/表头行必须显式指定
    function addRow(cells, ht){ rows.push(cells); if(ht) rowH[rows.length]=ht; return rows.length; } // cells: array of {v,t,s,f,span}
    function col(c){ return String.fromCharCode(65+c); }
    function C(v,t,s,f,span){ return {v:v,t:t==null?(typeof v==="number"?"n":"s"):t,s:s,f:f,span:span||1}; }

    // 统计说明（仅顶部显示一次）
    // 说明：单元格的 span 会在序列化时自动占位并生成同行合并，此处不要再手动 push 同类 merge（重复会让 Excel 报「需要修复」）
    addRow([C(note,"s",4,null,8)], 78);   // 统计说明：4 行文本，合并单元格不会自动撑高，必须显式指定行高
    // 周客量（列：A大类 B部门 C总到访 D初诊 E复诊 F有成交 G:H成交率）
    addRow([C("周客量统计(数据区间："+periodText+")","s",1,null,8)], 24);
    addRow([C("部门","s",1,null,2),C("总到访客量","s",1),C("初诊客量","s",1),C("复诊客量","s",1),C("有成交客量","s",1),C("成交率","s",1,null,2)], 30);
    var r4n=rows.length+1;
    var r4=addRow([C("线上","s",3),C("营销部","s",3),C(kv[d[0]].tot,"n",3),C(kv[d[0]].chu,"n",3),C("=C"+r4n+"-D"+r4n,"f",3),C(kv[d[0]].cash,"n",3),C(kv[d[0]].rate,"n",6,null,2)]);
    var r5n=rows.length+1;
    var r5=addRow([C("线下","s",3),C("商务部","s",3),C(kv[d[1]].tot,"n",3),C(kv[d[1]].chu,"n",3),C("=C"+r5n+"-D"+r5n,"f",3),C(kv[d[1]].cash,"n",3),C(kv[d[1]].rate,"n",6,null,2)]);
    var r6n=rows.length+1;
    var r6=addRow([C("","s",3),C("其他(公司所有)","s",3),C(kv[d[2]].tot,"n",3),C(kv[d[2]].chu,"n",3),C("=C"+r6n+"-D"+r6n,"f",3),C(kv[d[2]].cash,"n",3),C(kv[d[2]].rate,"n",6,null,2)]);
    merges.push("A"+r5+":A"+r6);   // 「线下」跨 商务部、其他(公司所有) 两行
    var r7n=rows.length+1;
    var r7=addRow([C("合计","s",8,null,2),C("=SUM(C"+r4n+":C"+r6n+")","f",8),C("=SUM(D"+r4n+":D"+r6n+")","f",8),C("=SUM(E"+r4n+":E"+r6n+")","f",8),C("=SUM(F"+r4n+":F"+r6n+")","f",8),C("=F"+r7n+"/C"+r7n,"f",7,null,2)]);

    // 周业绩（列：A日期 B高新 C经开 D总业绩 E排名 F总到访 G总预约 H到访率）
    addRow([C("","s",0,null,0)]); // spacer
    addRow([C("周业绩统计(数据区间："+periodText+")","s",1,null,8)], 24);
    addRow([C("日期","s",1),C("高新(元)","s",1),C("经开(元)","s",1),C("总业绩(元)\n（高新+经开）","s",1),C("业绩排名","s",1),C("总到访客量","s",1),C("总预约客量","s",1),C("到访率","s",1)], 32);
    var firstDay=rows.length+1;
    p.rows.forEach(function(r){
      addRow([C(r.date.replace(/-/g,"/")+"("+weekday(r.date)+")","s",3),C(r.gao,"n",3),C(r.jing,"n",3),
        C("=B"+(rows.length+1)+"+C"+(rows.length+1),"f",3),C(r.rank,"n",3),C(r.visit,"n",3),C(r.book,"n",3),C(r.rate,"n",6)]);
    });
    var lastDay=rows.length;
    var rc=addRow([C("合计","s",8),C("=SUM(B"+firstDay+":B"+lastDay+")","f",8),C("=SUM(C"+firstDay+":C"+lastDay+")","f",8),
      C("=SUM(D"+firstDay+":D"+lastDay+")","f",8),C("/","s",8),C(p.wkVisit,"n",8),C(p.wkBook,"n",8),C(p.wkBook?p.wkVisit/p.wkBook:0,"n",7)]);
    addRow([C("业绩最高值是"+fmtMD(p.maxDay)+"("+weekday(p.maxDay)+")  业绩最低值是"+fmtMD(p.minDay)+"("+weekday(p.minDay)+")  业绩平均值："+money(p.avg)+"元","s",9,null,8)], 30);   // 业绩汇总行：加粗 + 行高 30

    // 管家（列：A姓名 B总业绩 C排名 D总预约 E总到访 F:H到访率）
    addRow([C("","s",0,null,0)]); // spacer
    addRow([C("管家业绩统计(数据区间："+periodText+")","s",1,null,8)], 24);
    addRow([C("姓名","s",1),C("总业绩(元)","s",1),C("业绩排名","s",1),C("总预约量","s",1),C("总到访量","s",1),C("到访率","s",1,null,3)], 30);
    var fb=rows.length+1;
    b.rows.forEach(function(r){
      var rate = (r.rate==null) ? "/" : r.rate;
      var bk = r.book>0 ? r.book : "/";
      var vt = r.visit>0 ? r.visit : "/";
      // 「无」行的排名是字符串 "/"，必须按字符串写，否则数值单元格里出现文本会被 Excel 判为损坏
      addRow([C(r.name,"s",3),C(r.perf,"n",3),(r.rank==="/")?C("/","s",3):C(r.rank,"n",3),
        (bk==="/")?C("/","s",3):C(bk,"n",3),
        (vt==="/")?C("/","s",3):C(vt,"n",3),
        (r.rate==null)?C("/","s",3,null,3):C(rate,"n",6,null,3)]);
    });
    var lb=rows.length;
    // 「合计」只占 A 列：若跨 A:B 会把 B 列（总业绩合计）挤到 C 列造成整行错位
    addRow([C("合计","s",8),C("=SUM(B"+fb+":B"+lb+")","f",8),C("/","s",8),C("=SUM(D"+fb+":D"+lb+")","f",8),C("=SUM(E"+fb+":E"+lb+")","f",8),C(b.sumB?b.sumV/b.sumB:0,"n",7,null,3)]);

    // serialize
    // 列宽需同时照顾三张表：A=日期/部门大类/姓名 B=部门/金额 C=经开·初诊 D=总业绩(换行最宽) E=排名·复诊 F=有成交·总到访 G=总预约 H=到访率
    var colXml="<cols><col min=\"1\" max=\"1\" width=\"18\" customWidth=\"1\"/><col min=\"2\" max=\"2\" width=\"16\" customWidth=\"1\"/><col min=\"3\" max=\"3\" width=\"13\" customWidth=\"1\"/><col min=\"4\" max=\"4\" width=\"18\" customWidth=\"1\"/><col min=\"5\" max=\"5\" width=\"13\" customWidth=\"1\"/><col min=\"6\" max=\"6\" width=\"13\" customWidth=\"1\"/><col min=\"7\" max=\"7\" width=\"13\" customWidth=\"1\"/><col min=\"8\" max=\"8\" width=\"12\" customWidth=\"1\"/></cols>";
    var rowsXml="";
    for(var ri=0;ri<rows.length;ri++){
      var cells=rows[ri]; if(!cells) continue;
      var rownum=ri+1;
      rowsXml+='<row r="'+rownum+'"'+(rowH[rownum]?' ht="'+rowH[rownum]+'" customHeight="1"':'')+'>';
      // 列游标：span>1 的单元格必须占满 span 列，否则其后所有单元格会整体左移一列（表头与数据错位、公式自引用）
      var colIdx=0;
      for(var ci2=0;ci2<cells.length;ci2++){
        var cell=cells[ci2]; if(cell==null) continue;
        var span=(cell.span>1)?cell.span:1;
        var ref=col(colIdx)+rownum;
        if(span>1) merges.push(ref+":"+col(colIdx+span-1)+rownum);
        if(cell.t==="f"){
          rowsXml+='<c r="'+ref+'" s="'+cell.s+'"><f>'+escapeXml(String(cell.v).slice(1))+'</f><v>0</v></c>';
        } else if(cell.t==="n"){
          rowsXml+='<c r="'+ref+'" s="'+cell.s+'"><v>'+cell.v+'</v></c>';
        } else {
          rowsXml+='<c r="'+ref+'" s="'+cell.s+'" t="inlineStr"><is><t xml:space="preserve">'+escapeXml(cell.v)+'</t></is></c>';
        }
        // 合并区右侧补空单元格，保证黑框线连续
        for(var k=1;k<span;k++){ rowsXml+='<c r="'+col(colIdx+k)+rownum+'" s="'+cell.s+'"/>'; }
        colIdx+=span;
      }
      rowsXml+='</row>';
    }
    var mergeXml='<mergeCells count="'+merges.length+'">'+merges.map(function(m){return '<mergeCell ref="'+m+'"/>';}).join("")+'</mergeCells>';
    var sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'+
      ' '+colXml+'\n <sheetData>'+rowsXml+'</sheetData>\n '+mergeXml+'\n</worksheet>';

    var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
      '<Default Extension="xml" ContentType="application/xml"/>'+
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'+
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    var rootRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var wbRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="周客量统计" sheetId="1" r:id="rId1"/></sheets></workbook>';

    var zip=new JSZip();
    zip.file("[Content_Types].xml",ct);
    zip.folder("_rels").file(".rels",rootRels);
    var xl=zip.folder("xl"); xl.file("workbook.xml",workbook); xl.file("styles.xml",styles);
    xl.folder("_rels").file("workbook.xml.rels",wbRels);
    xl.folder("worksheets").file("sheet1.xml",sheet);
    return zip;
  }
  function escapeXml(s){ return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

  function exportExcel(){
    if(!APPT.length && !ORDER.length) return;
    var R=effectiveR(); if(!R) return;
    var period=rangeText();
    var zip=buildXlsxZip(R, period);
    var safe=period.replace(/[\\/:*?"<>|]/g,"_");
    zip.generateAsync({type:"blob"}).then(function(blob){
      var a=document.createElement("a"); a.href=URL.createObjectURL(blob);
      a.download="周客量业绩统计_"+safe+".xlsx"; document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(a.href); },2000);
    }).catch(function(err){ alert("导出Excel失败："+err); });
  }
  // 导出图片：先算出「完整内容宽度」（含被横向滚动隐藏的列），
  // 再离屏克隆一份、取消横向滚动让表格按完整宽度展开后截图，
  // 避免手机端屏幕显示不全时导出图片同样被裁切。
  // 导出图片：把指定节点离屏克隆、按完整内容宽度展开后截图，避免横向滚动导致裁切。
  function exportNodePng(node, name){
    if(!node) return;
    // 取「内容实际宽度」：预约量统计的卡片/表格为 max-width:820px 居中，
    // 若直接用容器宽度（更宽）会产生左右空白；这里取卡片真实宽度作为画布宽度。
    var fullW = 0;
    [].forEach.call(node.querySelectorAll(".report-card"),function(c){
      var w=Math.ceil(c.getBoundingClientRect().width)||c.offsetWidth||0;
      if(w>fullW) fullW=w;
    });
    [].forEach.call(node.querySelectorAll(".stat-note,.report-note"),function(c){
      var w=Math.ceil(c.getBoundingClientRect().width)||c.offsetWidth||0;
      if(w>fullW) fullW=w;
    });
    [].forEach.call(node.querySelectorAll("table.tbl"),function(t){
      var w=Math.ceil(t.getBoundingClientRect().width)||t.offsetWidth||0;
      if(w>fullW) fullW=w;
    });
    [].forEach.call(node.querySelectorAll(".scroll"),function(sc){
      if(sc.scrollWidth>fullW) fullW=sc.scrollWidth;
    });
    if(!fullW) fullW = node.offsetWidth || node.clientWidth || node.scrollWidth || 0;

    var holder=document.createElement("div");
    holder.setAttribute("data-export-holder","1");
    holder.style.cssText="position:fixed;left:-10000px;top:0;z-index:-1;background:#ffffff;overflow:visible;";
    holder.style.width=fullW+"px";

    var clone=node.cloneNode(true);
    clone.style.width=fullW+"px";
    clone.style.maxWidth="none";
    clone.style.margin="0";
    // 子元素去掉 820px 上限与自动外边距，让内容刚好铺满画布、不留左右空白
    [].forEach.call(clone.querySelectorAll(".report-card,.stat-note,.report-note"),function(c){
      c.style.maxWidth="none"; c.style.width="100%"; c.style.marginLeft="0"; c.style.marginRight="0";
    });
    [].forEach.call(clone.querySelectorAll(".scroll"),function(sc){
      sc.style.overflow="visible"; sc.style.width="100%"; sc.style.maxWidth="none";
    });
    [].forEach.call(clone.querySelectorAll("table.tbl"),function(t){
      t.style.width="100%"; t.style.minWidth="0";
    });

    holder.appendChild(clone);
    document.body.appendChild(holder);

    var fullH = clone.offsetHeight || clone.scrollHeight;

    htmlToImage.toPng(clone,{pixelRatio:2,backgroundColor:"#ffffff",cacheBust:true,width:fullW,height:fullH})
      .then(function(dataUrl){
        var a=document.createElement("a"); a.href=dataUrl; a.download=name+".png"; a.click();
      })
      .catch(function(err){ alert("导出图片失败："+err); })
      .then(function(){ if(holder.parentNode) holder.parentNode.removeChild(holder); });
  }
  // 导出图片：导出三报表平台（模块A）的报表区；预约量统计（模块B）有独立的「导出此区块图片」按钮
  function exportPng(){ exportNodePng(document.getElementById("reports"), "周客量业绩统计_"+rangeText()); }

  // ---------- 预约量统计（模块B）导出 Excel：管家维度 + 医生维度 两个工作表 ----------
  function buildNextXlsxZip(target, bRows, dRows, totalKe, totalChu, totalFu){
    var HDR="FF01418B";
    var note=NEXT_NOTE_TEXT;
    // 样式：0默认 / 1标题(蓝底白字加粗) / 2数据 / 3合计(加粗) / 4说明(左对齐换行) / 5百分比 / 6百分比(合计)
    var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'+
    ' <numFmts count="1"><numFmt numFmtId="164" formatCode="0.0%"/></numFmts>\n'+
    ' <fonts count="5">\n'+
    '  <font><sz val="11"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><b/><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>\n'+
    '  <font><sz val="11"/><color rgb="FF000000"/><name val="Microsoft YaHei"/></font>\n'+
    ' </fonts>\n'+
    ' <fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'+
    '  <fill><patternFill patternType="solid"><fgColor rgb="'+HDR+'"/><bgColor indexed="64"/></patternFill></fill>'+
    '  <fill><patternFill patternType="solid"><fgColor rgb="FFD9E6F5"/><bgColor indexed="64"/></patternFill></fill></fills>\n'+
    ' <borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>\n'+
    '  <border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border></borders>\n'+
    ' <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>\n'+
    ' <cellXfs count="7">\n'+
    '  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>\n'+
    '  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>\n'+
    '  <xf numFmtId="164" fontId="2" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    '  <xf numFmtId="164" fontId="3" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center"/></xf>\n'+
    ' </cellXfs>\n'+
    ' <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>\n'+
    '</styleSheet>';

    function col(c){ return String.fromCharCode(65+c); }
    function C(v,t,s,f,span){ return {v:v,t:t==null?(typeof v==="number"?"n":"s"):t,s:s,f:f,span:span||1}; }
    function buildSheet(title, dimLabel, rows){
      var rws=[], merges=[], rowH={};
      function addRow(cells, ht){ rws.push(cells); if(ht) rowH[rws.length]=ht; return rws.length; }
      addRow([C(note,"s",4,null,5)], 62);
      addRow([C(title,"s",1,null,5)], 24);
      addRow([C(dimLabel,"s",1),C("预约客量","s",1),C("初诊客量","s",1),C("复诊客量","s",1),C("占比","s",1)], 20);
      rows.forEach(function(r){
        addRow([C(r.name,"s",2),C(r.ke,"n",2),C(r.chu,"n",2),C(r.fu,"n",2),C(totalKe?r.ke/totalKe:0,"n",5)]);
      });
      // 合计行使用全局去重口径（与页面一致），管家/医生两个工作表合计完全相同
      addRow([C("合计","s",3),C(totalKe||0,"n",3),C(totalChu||0,"n",3),C(totalFu||0,"n",3),C(totalKe?1:0,"n",6)]);
      var rowsXml="";
      for(var ri=0;ri<rws.length;ri++){
        var cells=rws[ri]; if(!cells) continue;
        var rownum=ri+1;
        rowsXml+='<row r="'+rownum+'"'+(rowH[rownum]?' ht="'+rowH[rownum]+'" customHeight="1"':'')+'>';
        var colIdx=0;
        for(var ci2=0;ci2<cells.length;ci2++){
          var cell=cells[ci2]; if(cell==null) continue;
          var span=(cell.span>1)?cell.span:1;
          var ref=col(colIdx)+rownum;
          if(span>1) merges.push(ref+":"+col(colIdx+span-1)+rownum);
          if(cell.t==="f"){ rowsXml+='<c r="'+ref+'" s="'+cell.s+'"><f>'+escapeXml(String(cell.v).slice(1))+'</f><v>0</v></c>'; }
          else if(cell.t==="n"){ rowsXml+='<c r="'+ref+'" s="'+cell.s+'"><v>'+cell.v+'</v></c>'; }
          else { rowsXml+='<c r="'+ref+'" s="'+cell.s+'" t="inlineStr"><is><t xml:space="preserve">'+escapeXml(cell.v)+'</t></is></c>'; }
          for(var k=1;k<span;k++){ rowsXml+='<c r="'+col(colIdx+k)+rownum+'" s="'+cell.s+'"/>'; }
          colIdx+=span;
        }
        rowsXml+='</row>';
      }
      var colXml='<cols><col min="1" max="1" width="28" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/><col min="3" max="3" width="12" customWidth="1"/><col min="4" max="4" width="12" customWidth="1"/><col min="5" max="5" width="10" customWidth="1"/></cols>';
      var mergeXml='<mergeCells count="'+merges.length+'">'+merges.map(function(m){return '<mergeCell ref="'+m+'"/>';}).join("")+'</mergeCells>';
      return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'+
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'+
        ' '+colXml+'\n <sheetData>'+rowsXml+'</sheetData>\n '+mergeXml+'\n</worksheet>';
    }

    var ct='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'+
      '<Default Extension="xml" ContentType="application/xml"/>'+
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'+
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'+
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>';
    var rootRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    var wbRels='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'+
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'+
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    var workbook='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'+
      '<sheets><sheet name="管家维度" sheetId="1" r:id="rId1"/><sheet name="医生维度" sheetId="2" r:id="rId2"/></sheets></workbook>';

    var zip=new JSZip();
    zip.file("[Content_Types].xml",ct);
    zip.folder("_rels").file(".rels",rootRels);
    var xl=zip.folder("xl"); xl.file("workbook.xml",workbook); xl.file("styles.xml",styles);
    xl.folder("_rels").file("workbook.xml.rels",wbRels);
    var ws=xl.folder("worksheets");
    ws.file("sheet1.xml",buildSheet("管家维度 · "+target+" 预约客量","预约健康管家",bRows));
    ws.file("sheet2.xml",buildSheet("医生维度 · "+target+" 预约客量","预约医生",dRows));
    return zip;
  }
  // 供导出按钮取用最近一次渲染结果（避免重复计算、保证与页面数字一致）
  var NEXT_CACHE = { target:"", bRows:[], dRows:[], total:0 };
  function exportNextExcel(){
    var c=NEXT_CACHE;
    if(!APPT.length || !c.target){ alert("请先导入《预约列表》后再导出。"); return; }
    var zip=buildNextXlsxZip(c.target, c.bRows, c.dRows, c.total, c.totalChu, c.totalFu);
    var safe=c.target.replace(/[\\/:*?"<>|]/g,"_");
    zip.generateAsync({type:"blob"}).then(function(blob){
      var a=document.createElement("a"); a.href=URL.createObjectURL(blob);
      a.download="预约客量统计_"+safe+".xlsx"; document.body.appendChild(a); a.click(); document.body.removeChild(a);
      setTimeout(function(){ URL.revokeObjectURL(a.href); },2000);
    }).catch(function(err){ alert("导出Excel失败："+err); });
  }

  // ---------- 预约量统计（按具体日期，支持历史日期）子功能 ----------
  // 只依赖已导入的《预约列表》(APPT)，按管家 / 医生双维度统计所选「统计日」的预约客量。
  // 统计日直接选取（可选任意历史日期），不再固定为基准日+1；标题与说明直接展示具体日期。
  function nextTableHtml(rows, totals, dimLabel){
    var head='<thead><tr><th>'+dimLabel+'</th><th>预约客量</th><th>初诊客量</th><th>复诊客量</th><th>占比</th></tr></thead>';
    if(!rows.length) return head+'<tbody><tr><td colspan="5" style="color:#5b6472">该日无预约数据</td></tr></tbody>';
    var body="";
    rows.forEach(function(r){
      var pct = totals.ke? (r.ke/totals.ke*100).toFixed(1)+"%" : "—";
      body+='<tr><td class="dim">'+escapeHtml(r.name)+'</td><td>'+r.ke+'</td><td>'+r.chu+'</td><td>'+r.fu+'</td><td class="pct">'+pct+'</td></tr>';
    });
    var totPct = totals.ke? "100.0%" : "—";
    // 合计行使用全局去重口径（totals 由 renderNext 统一计算），管家/医生两维度合计必然一致
    body+='<tr class="total"><td class="dim">合计</td><td>'+totals.ke+'</td><td>'+totals.chu+'</td><td>'+totals.fu+'</td><td class="pct">'+totPct+'</td></tr>';
    return head+'<tbody>'+body+'</tbody>';
  }
  function setNextEmpty(target){
    var emptyButler='<thead><tr><th>预约健康管家</th><th>预约客量</th><th>初诊客量</th><th>复诊客量</th><th>占比</th></tr></thead><tbody><tr><td colspan="5" style="color:#5b6472">导入预约列表后展示</td></tr></tbody>';
    var emptyDoctor='<thead><tr><th>预约医生</th><th>预约客量</th><th>初诊客量</th><th>复诊客量</th><th>占比</th></tr></thead><tbody><tr><td colspan="5" style="color:#5b6472">导入预约列表后展示</td></tr></tbody>';
    document.getElementById("tblNextButler").innerHTML=emptyButler;
    document.getElementById("tblNextDoctor").innerHTML=emptyDoctor;
    document.getElementById("nextButlerTitle").textContent="管家维度 · "+target+" 预约客量";
    document.getElementById("nextDoctorTitle").textContent="医生维度 · "+target+" 预约客量";
    NEXT_CACHE={ target:"", bRows:[], dRows:[], total:0 };
  }
  function renderNext(){
    var statEl=document.getElementById("nextStatDate");
    var target = statEl && statEl.value ? statEl.value : todayStr();
    var nd=document.getElementById("nextDate"); if(nd) nd.textContent = target || "--";
    var range=apptDateRange();
    if(!APPT.length){
      var st0=document.getElementById("nextStatus");
      if(st0){ st0.className="status"; st0.innerHTML="请先导入《预约列表》（与三报表平台一同导入即可）。"; }
      setNextEmpty(target);
      return;
    }
    var tgt=APPT.filter(function(r){ return r.date===target; });
    var butlerMap={}, doctorMap={}, totalKe=new Set(), gChu=new Set(), gFu=new Set();
    function isFuType(t){ return t==="复诊"||t==="疗程内"||t==="再消费"||t==="复查"||t===""; }
    tgt.forEach(function(r){
      totalKe.add(r.mem);
      var bk=r.butler || "未预约管家";
      var dk=r.doctor || "未预约医生";
      var isChu = (r.type==="初诊");
      if(isChu) gChu.add(r.mem); else if(isFuType(r.type)) gFu.add(r.mem);
      butlerMap[bk]=butlerMap[bk]||{all:new Set(),chu:new Set(),fu:new Set(),unc:new Set()};
      doctorMap[dk]=doctorMap[dk]||{all:new Set(),chu:new Set(),fu:new Set(),unc:new Set()};
      butlerMap[bk].all.add(r.mem); doctorMap[dk].all.add(r.mem);
      // 预约量统计：初诊客量=预约类型「初诊」；复诊客量=预约类型「复诊/疗程内/再消费/复查/空」
      if(isChu){ butlerMap[bk].chu.add(r.mem); doctorMap[dk].chu.add(r.mem); }
      else if(isFuType(r.type)){ butlerMap[bk].fu.add(r.mem); doctorMap[dk].fu.add(r.mem); }
      else { butlerMap[bk].unc.add(r.mem); doctorMap[dk].unc.add(r.mem); }
    });
    function toRows(map){ var a=Object.keys(map).map(function(k){return {name:k,ke:map[k].all.size,chu:map[k].chu.size,fu:map[k].fu.size};}); a.sort(function(x,y){ var ux=/^未预约/.test(x.name), uy=/^未预约/.test(y.name); if(ux&&!uy) return 1; if(!ux&&uy) return -1; return y.ke-x.ke; }); return a; }
    var bRows=toRows(butlerMap), dRows=toRows(doctorMap);
    // 合计使用「全局去重」口径：同一客户跨多个管家/医生时只在总客量中计一次，
    // 因此管家维度与医生维度的合计必然一致（都是当日全局去重会员数）。
    var totals={ ke: totalKe.size, chu: gChu.size, fu: gFu.size };
    document.getElementById("nextButlerTitle").textContent="管家维度 · "+target+" 预约客量";
    document.getElementById("nextDoctorTitle").textContent="医生维度 · "+target+" 预约客量";
    document.getElementById("tblNextButler").innerHTML=nextTableHtml(bRows, totals, "预约健康管家");
    document.getElementById("tblNextDoctor").innerHTML=nextTableHtml(dRows, totals, "预约医生");
    // 缓存本次渲染结果，供「导出 Excel」直接取用（与页面数字完全一致）
    NEXT_CACHE={ target:target, bRows:bRows, dRows:dRows, total:totals.ke, totalChu:totals.chu, totalFu:totals.fu };
    var st=document.getElementById("nextStatus");
    var rangeTxt = (range.min&&range.max) ? "（数据区间 "+range.min+" ~ "+range.max+"）" : "";
    if(totalKe.size===0){ st.className="status warn"; st.innerHTML="<b>"+target+" 暂无预约数据。</b>可点「用数据最晚日」查看数据最末日；或直接选其它历史日期。"+rangeTxt; }
    else { st.className="status"; st.innerHTML="已导入预约列表 <b>"+APPT.length+" 行</b>；统计日 <b>"+target+"</b> 共 <b>"+totalKe.size+" 客</b>。"+rangeTxt; }
  }
  function apptDateRange(){
    var ds=APPT.map(function(r){return r.date;}).filter(Boolean).sort();
    return ds.length? {min:ds[0], max:ds[ds.length-1]} : {min:null,max:null};
  }
  function defaultNextStat(){
    var range=apptDateRange();
    var el=document.getElementById("nextStatDate");
    if(el) el.value = range.max || todayStr();
  }

  // ---------- 流程 ----------
  function rangeText(){
    if(RANGE.start&&RANGE.end) return fmtMD(RANGE.start)+"~"+fmtMD(RANGE.end)+" 12:00";
    return "";
  }
  function autoRange(){
    var min=null,max=null;
    APPT.concat(ORDER).forEach(function(r){
      if(!r.date) return;
      if(min==null||r.date<min) min=r.date;
      if(max==null||r.date>max) max=r.date;
    });
    if(min){ RANGE.start=min; RANGE.end=max; dateStart.value=min; dateEnd.value=max; }
  }
  function recompute(){
    if(!APPT.length || !ORDER.length){
      status.innerHTML='<span class="warn">需同时导入《预约列表》与《订单项目收退款明细》两份文件后才能计算。</span>';
      return;
    }
    try{
      var R=effectiveR();
      render(R, rangeText());
      // 更新 KPI 总览卡片（用 R.kv，避免与下方 var kv=R.kv 赋值顺序冲突）
      var kpiT=document.getElementById("kpiTot"), kpiC=document.getElementById("kpiChu");
      var kpiF=document.getElementById("kpiFu"),  kpiR=document.getElementById("kpiRate");
      if(kpiT) kpiT.textContent=R.kv["合计"].tot;
      if(kpiC) kpiC.textContent=R.kv["合计"].chu;
      if(kpiF) kpiF.textContent=R.kv["合计"].fu;
      if(kpiR) kpiR.textContent=pctStr(R.kv["合计"].rate);
      btnPng.disabled=false; btnXlsx.disabled=false; btnClearOV.disabled=false;
      var kv=R.kv;
      status.innerHTML='已导入 <b>预约列表 '+APPT.length+' 行</b> / <b>订单明细 '+ORDER.length+' 行</b> · 数据区间 <b>'+rangeText()+'</b> · '+
        '周客量合计到访 <b>'+kv["合计"].tot+'</b> 人，成交率 <b>'+pctStr(kv["合计"].rate)+'</b>；周业绩合计 <b>'+money(R.perf.sum.total)+'</b> 元。'
        + ' · <span style="color:#888">版本 '+APP_VERSION+'</span>';
    }catch(err){ status.innerHTML='<span class="warn">计算失败：'+(err.message||err)+'</span>'; }
  }

  var fileBoth=document.getElementById("fileBoth");
  var dateStart=document.getElementById("dateStart");
  var dateEnd=document.getElementById("dateEnd");
  var btnCalc=document.getElementById("btnCalc");
  var btnPng=document.getElementById("btnPng");
  var btnXlsx=document.getElementById("btnXlsx");
  var btnClearOV=document.getElementById("btnClearOV");
  var status=document.getElementById("status");

  // 依据内容自动识别：订单明细含「开单咨询师/现款支付」，预约列表含「是否到访」
  function classifyAndLoad(wb){
    try{ return {kind:"order", data:parseOrder(wb)}; }catch(e){}
    try{ return {kind:"appt", data:parseAppt(wb)}; }catch(e){}
    return {kind:null};
  }
  function readAsArrayBuffer(f){
    return new Promise(function(res,rej){
      var r=new FileReader();
      r.onload=function(){ res(r.result); };
      r.onerror=function(){ rej(r.error||new Error("文件读取失败")); };
      r.readAsArrayBuffer(f);
    });
  }
  fileBoth.addEventListener("change",function(e){
    var files=[].slice.call(e.target.files||[]);
    if(!files.length) return;
    status.innerHTML='正在解析 <b>'+files.length+'</b> 个文件…';
    Promise.all(files.map(function(f){
      return readAsArrayBuffer(f).then(function(buf){
        var wb=XLSX.read(buf,{type:"array",cellDates:true});
        var r=classifyAndLoad(wb);
        if(r.kind==="order"){ ORDER=r.data; return '订单明细 '+ORDER.length+' 行'; }
        if(r.kind==="appt"){ APPT=r.data; return '预约列表 '+APPT.length+' 行'; }
        return '<span class="warn">无法识别（非预约列表/订单明细）：'+f.name+'</span>';
      }).catch(function(err){
        return '<span class="warn">读取失败：'+f.name+' '+(err.message||err)+'</span>';
      });
    })).then(function(msgs){
      var both = APPT.length && ORDER.length;
      if(both){ autoRange(); restoreOV(); recompute(); }
      else { status.innerHTML=msgs.join(' ； '); }
      // 三报表平台导入只生成本模块报表；预约量统计由模块B独立「导入预约列表」按钮生成，这里不再联动渲染
    });
  });
  btnCalc.addEventListener("click",function(){
    RANGE.start=dateStart.value||null; RANGE.end=dateEnd.value||null; recompute();
  });
  btnPng.addEventListener("click",exportPng);
  btnXlsx.addEventListener("click",exportExcel);
  btnClearOV.addEventListener("click",clearOV);

  // 业绩单元格手动编辑：input 时实时联动重算，blur 时把编辑框重新格式化为金额
  var reportsEl=document.getElementById("reports");
  reportsEl.addEventListener("input",onEditInput);
  reportsEl.addEventListener("blur",onEditBlur,true);

  // 预约量统计（按具体日期，支持历史日期）：控件与事件
  var nextStatDate=document.getElementById("nextStatDate");
  var btnNextMax=document.getElementById("btnNextMax");
  var btnNextPng=document.getElementById("btnNextPng");
  if(nextStatDate) nextStatDate.addEventListener("change",renderNext);
  if(btnNextMax) btnNextMax.addEventListener("click",function(){
    if(APPT.length){ defaultNextStat(); renderNext(); }
    else { var s=document.getElementById("nextStatus"); if(s){ s.className="status warn"; s.innerHTML="尚无预约数据，无法取最晚日。"; } }
  });
  if(btnNextPng) btnNextPng.addEventListener("click",function(){
    if(!APPT.length) return;
    var stEl=document.getElementById("nextStatDate");
    var target=stEl&&stEl.value?stEl.value:todayStr();
    exportNodePng(document.getElementById("nextReport"), "预约客量_"+target);
  });
  var btnNextXlsx=document.getElementById("btnNextXlsx");
  if(btnNextXlsx) btnNextXlsx.addEventListener("click",exportNextExcel);
  // 初始化：默认统计日=今天，渲染区块（空态）
  var nextNoteEl=document.getElementById("nextNote");
  if(nextNoteEl) nextNoteEl.textContent=NEXT_NOTE_TEXT;
  if(nextStatDate) nextStatDate.value=todayStr();
  renderNext();

  // ---------- 模块B：预约量统计 专用导入（仅导入《预约列表》） ----------
  // 与模块A（一次性导入订单+预约）完全独立：此导入只把预约列表写入 APPT，再刷新预约量报表。
  var fileAppt=document.getElementById("fileAppt");
  if(fileAppt){
    fileAppt.addEventListener("change",function(e){
      var f=e.target.files&&e.target.files[0]; if(!f) return;
      var st=document.getElementById("nextStatus");
      if(st){ st.className="status"; st.innerHTML="正在解析《预约列表》…"; }
      readAsArrayBuffer(f).then(function(buf){
        var wb=XLSX.read(buf,{type:"array",cellDates:true});
        var r=classifyAndLoad(wb);
        if(r.kind==="appt"){ APPT=r.data; defaultNextStat(); renderNext(); }
        else if(r.kind==="order"){ if(st){ st.className="status warn"; st.innerHTML="该文件被识别为「订单明细」，预约量统计只需《预约列表》。"; } }
        else { if(st){ st.className="status warn"; st.innerHTML="无法识别为《预约列表》，请确认文件。"; } }
      }).catch(function(err){ if(st){ st.className="status warn"; st.innerHTML="读取失败："+(err.message||err); } });
    });
  }

  // ---------- 模块切换（两个完全独立的模块） ----------
  // 模块A：三报表平台（导入订单明细+预约明细）；模块B：预约量统计（仅导入预约明细）。
  // 两个模块彼此独立，点击顶部按钮分别进入。
  var modnav=document.getElementById("modnav");
  if(modnav){
    [].forEach.call(modnav.querySelectorAll(".modbtn"),function(btn){
      btn.addEventListener("click",function(){
        var key=btn.getAttribute("data-mod");
        [].forEach.call(modnav.querySelectorAll(".modbtn"),function(b){
          if(b===btn) b.classList.add("active"); else b.classList.remove("active");
        });
        [].forEach.call(document.querySelectorAll(".mod-panel"),function(p){
          if(p.getAttribute("data-mod")===key) p.classList.add("active"); else p.classList.remove("active");
        });
      });
    });
  }
})();
