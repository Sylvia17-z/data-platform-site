const fs=require('fs');
const src=fs.readFileSync('app.js','utf8');
const m=src.match(/function calcExpr\(expr, base\)\{[\s\S]*?return isFinite\(result\)\?result:NaN;\n  \}/);
if(!m){console.log('提取 calcExpr 失败');process.exit(1);}
const calcExpr=eval('('+m[0].replace('function calcExpr','function')+')');
const approx=(a,b)=>Math.abs(a-b)<1e-9;
const cases=[
 ['1000+2000',0,3000],
 ['@-3000',5000,2000],
 ['@+1500',8000,9500],
 ['1000+2000*3',0,7000],
 ['(1000+2000)*3',0,9000],
 ['10/4',0,2.5],
 ['@*0.9',1000,900],
 ['1000+',0,NaN],
 ['1000++2000',0,NaN],
 ['5*+3',0,NaN],
 ['5*-3',0,-15],
 ['5+-3',0,2],
 ['5--3',0,8],
 ['',0,NaN],
 ['1000+2000)+5',0,NaN],
 ['-5+3',0,-2],
 ['+7',0,7],
 ['2*(3+4)',0,14],
 ['1000.5+2000.5',0,3001],
 ['500/0',0,0],
 ['((1+2)*3)',0,9],
 ['@/2',100,50],
 ['3+4*2/(1-5)',0,1],
];
let pass=0,fail=0;
for(const [e,base,exp] of cases){
  const got=calcExpr(e,base);
  const ok=(isNaN(exp)&&isNaN(got))||approx(got,exp);
  console.log((ok?'PASS':'FAIL'),JSON.stringify(e),'base='+base,'=>',got, ok?'':'(期望 '+exp+')');
  ok?pass++:fail++;
}
console.log('\n总计 '+(pass+fail)+'，PASS '+pass+'，FAIL '+fail);
process.exit(fail?1:0);
