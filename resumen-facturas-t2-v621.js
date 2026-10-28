(()=>{'use strict';const $=id=>document.getElementById(id);const clean=s=>String(s??'').trim().toUpperCase().replace(/\s+/g,'');const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));let busy=false;
const status=(s,bad=false)=>{const el=$('ctT2SummaryMessage');if(el){el.textContent=s;el.style.color=bad?'#b42318':'#54738a'}};
const rowHtml=r=>'<tr><td>'+esc(r.transporte)+'</td><td>'+esc(r.viaje)+'</td><td><strong>'+esc(r.placa)+'</strong></td><td>'+(r.recibido===true?'<span class="ctT2Received">✓ Sí</span>':'<span class="ctT2NotReceived">Pendiente</span>')+'</td></tr>';
async function refresh(){if(busy||!$('ctT2SummaryPending'))return;if(!window.ctRotationApi){status('No hay sesión activa en Control Transportes.',true);return;}busy=true;const b=$('ctT2SummaryRefresh');if(b)b.disabled=true;status('Consultando facturas y restricciones actualizadas…');try{
 const data=await window.ctRotationApi('ct_t2_invoice_summary_get');const pending=Array.isArray(data?.pending)?data.pending:[],held=Array.isArray(data?.restricted)?data.restricted:[];
 $('ctT2SummaryPending').innerHTML=pending.map(rowHtml).join('')||'<tr><td colspan="4">Sin facturas disponibles pendientes.</td></tr>';
 $('ctT2SummaryRestricted').innerHTML=held.map(rowHtml).join('')||'<tr><td colspan="4">Sin facturas con restricción activa.</td></tr>';
 $('ctT2SummaryPendingCount').textContent=pending.length;$('ctT2SummaryHeldCount').textContent=held.length;
 $('ctT2SummaryTotal').textContent='Pendientes: '+(pending.length+held.length);$('ctT2SummaryHeld').textContent='En restricción: '+held.length;
 status('Actualizado: '+new Date().toLocaleString('es-CO')+'. Solo se muestran facturas que no han sido entregadas.');
 }catch(err){status('No fue posible sincronizar con Validaciones Operación: '+err.message,true);}
 finally{busy=false;if(b)b.disabled=false}}
window.addEventListener('ct-refresh-all',()=>{if(document.querySelector('[data-handover-pane="T2"]')?.hidden===false)void refresh()});
function init(){if(!$('ctT2SummaryPending'))return;$('ctT2SummaryRefresh').onclick=refresh;document.querySelector('[data-handover-tab="T2"]')?.addEventListener('click',refresh);window.addEventListener('ct-billing-opened',()=>{if(document.querySelector('[data-handover-pane="T2"]')?.hidden===false)refresh()});document.addEventListener('visibilitychange',()=>{if(!document.hidden&&document.querySelector('[data-handover-pane="T2"]')?.hidden===false)refresh()});}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();})();
