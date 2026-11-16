/* Control Transportes: inicio de sesion en modulo desde ventana principal.
   No usa localStorage, sessionStorage ni cookies para guardar el PIN. */
(() => {
  'use strict';
  if(window.parent !== window) return; // Las vistas embebidas conservan su autenticacion actual.
  const openerAvailable=!!(window.opener && window.opener!==window);
  if(!openerAvailable && !window.ctNavigationSession?.get()) return;
  if(typeof crypto?.randomUUID !== 'function') return;
  const nonce=crypto.randomUUID();
  let handled=false;
  let timer;
  function request(){
    try { if(window.opener && !window.opener.closed) {
      window.opener.postMessage({type:'CT_SSO_REQUEST_V1', nonce}, location.origin);
    } } catch(e) { console.warn('No se pudo solicitar la sesión del módulo:',e); }
  }
  function accept(event){
    if(event.origin!==location.origin || event.source!==window.opener || !event.data)return;
    if(event.data.type==='CT_SSO_LOGOUT_V1'){ window.location.reload(); return; }
    if(event.data.type!=='CT_SSO_REPLY_V1'||event.data.nonce!==nonce||handled)return;
    if(!/^\d{1,8}$/.test(String(event.data.pin||'')))return;
    const field=document.getElementById('pin');
    const button=['enter','connect','loginBtn','login'].map(id=>document.getElementById(id)).find(Boolean);
    if(!field||!button)return;
    handled=true;
    clearTimeout(timer);
    field.value=String(event.data.pin);
    window.ctNavigationSession?.save(event.data.pin);
    // La validación de permisos del módulo continúa ejecutándose en Supabase.
    button.click();
  }
  window.addEventListener('message',accept);
  function start(){
    const cached=window.ctNavigationSession?.get();
    if(cached){
      const field=document.getElementById('pin');
      const button=['enter','connect','loginBtn','login'].map(id=>document.getElementById(id)).find(Boolean);
      if(field&&button){handled=true;field.value=cached;button.click();return;}
    }
    if(openerAvailable)request();
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start,{once:true});
  else start();
  timer=setTimeout(()=>{if(!handled){window.removeEventListener('message',accept);}},12000);
})();
