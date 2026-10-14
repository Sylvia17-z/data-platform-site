/* PWA service worker（2026-10-14 方案C：H5 桌面化）
 * 职责仅一个：存在 fetch handler 以满足 Chrome「添加到主屏幕」安装条件。
 * 刻意不做任何磁盘缓存 —— 页面与脚本永远走网络拉最新，
 * 与各页 app.js 的 Date.now() 缓存穿透策略保持一致，避免「刷新后仍跑旧版本」。
 * 注意：仅接管「同源」请求；跨域（云 SDK CDN 等）不经过 SW，
 * 避免改变其网络路径（透明代理 + 不缓存的语义）。 */
self.addEventListener('install', function(){ self.skipWaiting(); });
self.addEventListener('activate', function(e){ e.waitUntil(self.clients.claim()); });
self.addEventListener('fetch', function(e){
  if(e.request.method !== 'GET') return;
  var sameOrigin = false;
  try{ sameOrigin = new URL(e.request.url).origin === self.location.origin; }catch(err){ return; }
  if(!sameOrigin) return;   // 跨域请求不接管
  e.respondWith(
    fetch(e.request).catch(function(){
      return new Response('网络不可用，请检查网络后重试', { status: 503, statusText: 'offline' });
    })
  );
});
