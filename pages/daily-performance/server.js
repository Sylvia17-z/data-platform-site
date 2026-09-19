const http=require('http'),fs=require('fs'),path=require('path');
const root = String.raw`D:\workbuddy\2026-08-15-10-10-03\daily-performance`;
const types={'.html':'text/html','.js':'application/javascript','.css':'text/css','.json':'application/json'};
http.createServer((req,res)=>{
  let url=req.url.split('?')[0]; if(url==='/')url='/index.html';
  const fp=path.join(root,url);
  if(!fp.startsWith(root)){res.statusCode=403;return res.end();}
  fs.readFile(fp,(e,d)=>{ if(e){res.statusCode=404;return res.end();}
    res.setHeader('Content-Type',types[path.extname(fp)]||'application/octet-stream');
    res.setHeader('Cache-Control','no-store, no-cache, must-revalidate');
    res.setHeader('Pragma','no-cache');
    res.setHeader('Expires','0');
    res.end(d);});
}).listen(8800,()=>console.log('业绩平台服务已启动: http://localhost:8800/'));
