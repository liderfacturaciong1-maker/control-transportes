/* Sesión de navegación entre módulos (misma pestaña / ventanas abiertas desde la app).
   PIN solo en sessionStorage, NUNCA en localStorage; vence a las 8 horas.
   Cada módulo comprueba su autorización con Supabase antes de mostrar datos. */
(()=>{'use strict';
 const KEY='ct_navigation_session_v3', CHANNEL='ct_navigation_logout_v3', RETURN_KEY='ct_navigation_return_once_v1';
 const TTL=8*60*60*1000;
 const valid=p=>/^\d{1,8}$/.test(String(p||''));
 const storage=()=>{try{return window.sessionStorage}catch(_){return null}};
 function clear(notify=false){try{storage()?.removeItem(KEY);storage()?.removeItem(RETURN_KEY)}catch(_){}if(notify){try{new BroadcastChannel(CHANNEL).postMessage('logout')}catch(_){}}}
 function get(){try{const raw=storage()?.getItem(KEY);if(!raw)return '';const s=JSON.parse(raw);if(!valid(s.pin)||!Number.isFinite(s.until)||Date.now()>s.until){clear();return ''}return s.pin}catch(_){clear();return ''}}
 function save(pin){if(!valid(pin))return;try{storage()?.setItem(KEY,JSON.stringify({pin:String(pin),until:Date.now()+TTL}))}catch(_){}}
 try{const bc=new BroadcastChannel(CHANNEL);bc.onmessage=e=>{if(e.data==='logout'){clear();if(get()=== '' && document.getElementById('logout') && !document.getElementById('logout').classList.contains('hidden'))location.reload()}}}catch(_){}
 function markReturn(){try{storage()?.setItem(RETURN_KEY,'1')}catch(_){}}
 function consumeReturn(){try{const v=storage()?.getItem(RETURN_KEY)==='1';storage()?.removeItem(RETURN_KEY);return v}catch(_){return false}}
 // A return from a module to the root may restore a validated navigation session.
 document.addEventListener('click',event=>{
   const a=event.target.closest?.('a[href]');if(!a)return;
   try{const u=new URL(a.href,location.href);
     if(u.origin===location.origin&&u.pathname.replace(/\/index\.html$/,'/')==='/'&&location.pathname!=='/'&&location.pathname!=='/index.html') markReturn();
   }catch(_){}
 },true);
 window.ctNavigationSession={get,save,clear,markReturn,consumeReturn};
})();
