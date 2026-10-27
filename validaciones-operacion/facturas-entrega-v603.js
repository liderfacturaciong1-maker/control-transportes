(()=>{'use strict';const $=id=>document.getElementById(id);const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));let sb,pin='',rows=[],restrictions=[],write=false,canUpload=false,canDelete=false,canPurge=false,restrictionsOk=false;const normalize=s=>String(s??'').trim().toUpperCase();const clean=s=>normalize(s).replace(/[^A-Z0-9]/g,'');const alertStatus=(m,bad=false)=>{$('message').textContent=m;$('message').classList.toggle('error',bad)};
async function rpc(name,params={}){const {data,error}=await sb.rpc(name,{p_pin:pin,...params});if(error)throw Error(error.message);return data}
function groupRestricted(row){
 if(row.entregada||row.recibido!==true)return false;
 const plate=clean(row.placa),transport=clean(row.transporte);
 return restrictions.some(r=>{
  if(String(r.estado_liberacion||'ACTIVA').trim().toUpperCase()!=='ACTIVA')return false;
  // La restricción aplica a la placa aun si el transporte de Restricciones
  // es distinto al transporte del archivo de facturas para entrega.
  const restrictedPlate=clean(r.placa),restrictedTransport=clean(r.transporte);
  return (plate&&restrictedPlate===plate)||
         (transport&&restrictedTransport===transport);
 });
}
function buttons(row){let state=row.recibido===true?'yes':row.recibido===false?'no':'';return `<td><button class="status ${state}" title="Alternar entre recibido y no recibido" data-id="${row.id}" data-field="recibido" ${!write?'disabled':''}>${row.recibido===true?'✓':row.recibido===false?'✕':'—'}</button></td><td><button class="status" title="Marcar como entregada" data-id="${row.id}" data-field="entregada" ${(!write||row.recibido!==true||groupRestricted(row))?'disabled':''}>✓ Entregar</button></td>`;}
function render(){const active=rows.filter(r=>!r.entregada),waiting=active.filter(r=>!groupRestricted(r)),held=active.filter(groupRestricted),done=rows.filter(r=>r.entregada);$('readyCount').textContent=waiting.length;$('restrictedCount').textContent=held.length;$('deliveredCount').textContent=done.length;
const tr=r=>`<td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td><b>${esc(r.placa)}</b></td>`;
$('readyRows').innerHTML=waiting.map(r=>`<tr>${tr(r)}${buttons(r)}</tr>`).join('')||'<tr><td colspan="5" class="empty">Sin facturas pendientes.</td></tr>';
$('restrictedRows').innerHTML=held.map(r=>`<tr>${tr(r)}</tr>`).join('')||'<tr><td colspan="3" class="empty">Sin facturas restringidas.</td></tr>';
$('deliveredRows').innerHTML=done.map(r=>`<tr>${tr(r)}<td>${r.entregada_el?new Date(r.entregada_el).toLocaleString('es-CO'):'—'}</td></tr>`).join('')||'<tr><td colspan="4" class="empty">Sin facturas entregadas.</td></tr>';
}
async function load(){rows=await rpc('ct_validation_invoices_get');let restrictionError='';try{restrictions=await rpc('ct_finance_restrictions_list')}catch(e){restrictions=[];restrictionError=e.message}restrictionsOk=!restrictionError;render();if(restrictionError)alertStatus('No se pudieron consultar restricciones. No se puede determinar con seguridad cuáles facturas están restringidas: '+restrictionError,true);else alertStatus('Información actualizada. Cruce con Reporte Restricciones activo.');$('complete').disabled=!write||!restrictionsOk||!rows.some(r=>!r.entregada&&r.recibido!==true);}
async function connect(){pin=$('pin').value.trim();if(!/^\d{1,8}$/.test(pin))return alertStatus('Ingresa un PIN válido.',true);try{const cfg=window.CT_CONFIG||window.config||window.APP_CONFIG||{};let url=cfg.SUPABASE_URL,key=cfg.SUPABASE_ANON_KEY;if(!url||!key){for(const keyName of Object.keys(window)){if(keyName.toLowerCase().includes('config')){const obj=window[keyName];if(obj&&typeof obj==='object'&&obj.SUPABASE_URL){url=obj.SUPABASE_URL;key=obj.SUPABASE_ANON_KEY;break}}}}if(!url||!key)throw Error('No se encontró la configuración Supabase de la aplicación.');sb=window.supabase.createClient(url,key,{auth:{persistSession:false}});const access=await rpc('ct_validaciones_can_access');if(access!==true)throw Error('Este usuario no tiene autorización para Validaciones Operación.');write=(await rpc('ct_validation_invoice_auth',{p_write:true}))===true;
// Los permisos complementarios no deben impedir el ingreso al módulo.
// Ante errores, se dejan bloqueadas las acciones administrativas.
canUpload=false;canDelete=false;canPurge=false;
let rightsWarning='';
try{
 const rights=await rpc('ct_validation_invoice_rights_get');
 canUpload=rights?.upload===true;
 canDelete=rights?.delete===true;
}catch(err){
 rightsWarning='No se pudieron consultar permisos de carga/eliminación. Verifica ACTUALIZAR_PERMISOS_VALIDACIONES_v608.sql. '+err.message;
 // Fallar de forma segura: no habilitar cargas ni eliminación por omisión.
}
try{canPurge=(await rpc('ct_validation_invoice_can_purge'))===true;}
catch(err){rightsWarning+=(rightsWarning?' | ':'')+'No se pudo consultar permiso de borrado total: '+err.message;}
$('purgeAll').hidden=!canPurge;$('workspace').hidden=false;$('authState').textContent=write?'Administrador · Puede modificar':'Autorizado · Solo visualizar';$('pin').value='';['excel','fileLabel'].forEach(id=>{$(id).disabled=!canUpload;if(id==='fileLabel')$(id).style.display=canUpload?'':'none'});$('clear').disabled=!canDelete;$('clear').hidden=!canDelete;$('complete').disabled=!write;await load();if(rightsWarning)alertStatus('Ingreso permitido. '+rightsWarning,true);}catch(e){alertStatus('No se pudo iniciar sesión: '+e.message,true)}}
async function purgeAll(){
 if(!canPurge)return alertStatus('No tienes autorización para eliminar todos los datos.',true);
 const total=rows.length;
 if(!confirm(`ADVERTENCIA: se eliminarán definitivamente ${total} registros de Validaciones Operación, incluidos pendientes, restringidos y entregados. ¿Continuar?`))return;
 try{
  const deleted=await rpc('ct_validation_invoices_purge',{p_confirmation:'ELIMINAR TODO'});
  $('meta').textContent='Base reiniciada · Sin archivo activo';
  await load();
  alertStatus(`Se eliminaron ${deleted} registros. La base de Facturas para entrega quedó vacía.`);
 }catch(e){alertStatus('No se pudo eliminar: '+e.message,true)}
}
async function importFile(file){if(!file||!canUpload)return;try{const book=XLSX.read(await file.arrayBuffer(),{type:'array'}),ws=book.Sheets[book.SheetNames[0]];let a=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:''});let at=a.findIndex(r=>r.some(c=>normalize(c)==='TRANSPORTE')&&r.some(c=>normalize(c)==='VIAJE')&&r.some(c=>normalize(c)==='PLACA'));if(at<0)throw Error('El archivo debe tener las columnas Transporte, Viaje y Placa.');const header=a[at].map(normalize);const indices=['TRANSPORTE','VIAJE','PLACA'].map(x=>header.indexOf(x));const loaded=a.slice(at+1).map(r=>({transporte:String(r[indices[0]]||'').trim(),viaje:String(r[indices[1]]||'').trim(),placa:normalize(r[indices[2]])})).filter(r=>r.transporte&&r.placa);if(!loaded.length)throw Error('El Excel no contiene filas válidas.');if(!confirm(`¿Agregar ${loaded.length} registros del Excel? Se conservarán todas las facturas pendientes, restringidas y entregadas de días anteriores, sin duplicar registros abiertos.`))return;await rpc('ct_validation_invoices_import',{p_rows:loaded});await load();$('meta').textContent=`Último archivo: ${file.name} · ${loaded.length} registros importados`;}catch(e){alertStatus('Error de importación: '+e.message,true)}finally{$('excel').value=''}}
async function modify(id,field){if(!write)return;const item=rows.find(r=>r.id===id);if(!item)return;if(!restrictionsOk)return alertStatus('No se puede modificar: primero debe funcionar el cruce con restricciones.',true);const val=field==='recibido'?item.recibido!==true:true;if(field==='entregada'&&groupRestricted(item))return alertStatus('La factura está en restricción y no se puede entregar hasta que se libere.',true);if(field==='entregada'&&item.recibido!==true)return alertStatus('Primero confirma Recibido antes de entregar la factura.',true);if(field==='entregada'&&!confirm(`¿Marcar como entregada la factura del transporte ${item.transporte}?`))return;try{await rpc('ct_validation_invoice_status',{p_id:id,p_field:field,p_value:val});await load()}catch(e){alertStatus(e.message,true)}}
async function complete(){if(!write||!restrictionsOk)return;const pending=rows.filter(r=>!r.entregada&&r.recibido!==true);if(!pending.length)return;if(!confirm(`¿Marcar las ${pending.length} facturas pendientes como recibidas?`))return;let errors=0;for(const r of pending){try{await rpc('ct_validation_invoice_status',{p_id:r.id,p_field:'recibido',p_value:true})}catch(e){errors++}}await load();if(errors)alertStatus(`${errors} facturas no se pudieron actualizar.`,true)}
function init(){if(new URLSearchParams(location.search).get('embedded')==='t2'&&window.parent!==window){window.addEventListener('message',ev=>{if(ev.origin!==location.origin||ev.source!==window.parent||ev.data?.type!=='CT_T2_PIN_REPLY')return;const field=$('pin');if(!field||!ev.data.pin)return;field.value=String(ev.data.pin);connect();},{once:true});window.parent.postMessage({type:'CT_T2_PIN_REQUEST'},location.origin);}document.querySelectorAll('[data-panel]').forEach(b=>b.onclick=()=>{const box=$(b.dataset.panel);box.classList.toggle('ctCollapsed');b.textContent=box.classList.contains('ctCollapsed')?'Mostrar ▼':'Ocultar ▲'});document.querySelectorAll('[data-show]').forEach(b=>b.onclick=()=>{const box=$(b.dataset.show);box.classList.remove('ctCollapsed');box.querySelector('.ctHidePanel').textContent='Ocultar ▲';});$('connect').onclick=connect;$('purgeAll').onclick=purgeAll;$('reload').onclick=()=>load().catch(e=>alertStatus(e.message,true));$('excel').onchange=e=>importFile(e.target.files?.[0]);$('clear').onclick=async()=>{if(!canDelete||!confirm('¿Quitar la referencia del archivo cargado? Las facturas pendientes, restringidas y entregadas permanecerán registradas.'))return;try{await rpc('ct_validation_invoices_clear');$('meta').textContent='Archivo diario retirado · El historial y los pendientes se conservan';await load()}catch(e){alertStatus(e.message,true)}};$('complete').onclick=complete;$('readyRows').onclick=e=>{const b=e.target.closest('button[data-id]');if(b)modify(Number(b.dataset.id),b.dataset.field)};setInterval(()=>{if(pin&&sb)load().catch(()=>{})},60000)}if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();})();
