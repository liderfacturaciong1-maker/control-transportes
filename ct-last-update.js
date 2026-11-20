/* Indicador temporal por pestaña. La hora de consulta local no equivale a modificación compartida. */
(()=>{'use strict';
 const fmt=new Intl.DateTimeFormat('es-CO',{timeZone:'America/Bogota',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:true});
 const header=document.querySelector('header');if(!header)return;
 const box=document.createElement('div');box.id='ctLastUpdateInfo';box.setAttribute('role','status');
 box.style.cssText='font-size:12px;line-height:1.5;padding:5px 11px;border-radius:8px;background:#eef6fb;color:#174563;margin:5px 0;display:inline-flex;align-items:center;gap:6px;flex-wrap:wrap;';
 header.append(box);
 let current='general';const modified=new Map();const consulted=new Map();
 function show(){
   const value=modified.get(current),read=consulted.get(current);
   box.textContent=value?'🕒 Última modificación registrada: '+fmt.format(value)+' (Colombia)':read?'🕒 Última consulta en este dispositivo: '+fmt.format(read)+' (Colombia)':'🕒 Aún no hay fecha de actualización verificada';
 }
 window.ctLastDataUpdate=(tab,date)=>{if(!date)return;const d=new Date(date);if(!Number.isNaN(d.getTime())){modified.set(tab,d);show();}};
 window.ctLastDataConsulted=(tab)=>{consulted.set(tab,new Date());show();};
 function getCurrentTab(){
   const active=document.querySelector('.tabs [data-tab].active, nav [data-tab].active, [role="tab"][aria-selected="true"], .tab.active[data-tab]');
   if(active)return active.dataset.tab||active.id||'general';
   const visible=document.querySelector('main section:not(.hidden)[id], main article:not(.hidden)[id]');return visible?.id||'general';
 }
 function change(){current=getCurrentTab();show();}
 document.addEventListener('click',e=>{if(e.target.closest('[data-tab], [role="tab"], .tab, .nav-item'))setTimeout(change,0);});
 window.addEventListener('focus',change);
 window.ctShowLastUpdateTab=change;
 show();
})();
