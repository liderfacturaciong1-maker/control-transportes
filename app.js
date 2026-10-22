(()=>{'use strict';
const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const norm=s=>String(s??'').trim().replace(/\.0$/,'').toUpperCase();
// El prefijo S es solo de presentación; no cambia la clave de cruce.
const transportKey=s=>norm(s).replace(/^S(?=\d)/,'');
const displayedTransport=s=>{const key=transportKey(s);return key?`S${key}`:'';};
const heading=s=>norm(s).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9]/g,'');
const ix=(cols,variants)=>cols.findIndex(x=>variants.includes(heading(x)));
let statRows=new Map(),statLoaded=false,statFile='',loading=false;
const SYNC_VERSION='v5.23 indice status diario'; // Motor de sincronización v5.21
let cargoFallback=[],cargoFetching=false,cargoLastAttempt=0,cargoError='';
let statusIndex=new Map(),statusIndexReady=false,statusIndexError='',statusIndexBatch='';
let statusIndexFetching=false;
async function loadStatusIndex(force=false){
 if(statusIndexFetching||typeof sb==='undefined'||!sb||!loginPin)return;
 if(statusIndexReady&&!force)return;
 statusIndexFetching=true;statusIndexError='';
 try{
  const {data,error}=await sb.rpc('ct_report_info',{p_pin:loginPin,p_kind:'rutero'});
  if(error)throw error;
  const meta=data?.[0];
  if(!meta){statusIndex=new Map();statusIndexReady=false;statusIndexError='No hay un Rutero guardado.';return;}
  const index=meta.metadata?.statusIndex;
  if(!Array.isArray(index)){statusIndex=new Map();statusIndexReady=false;statusIndexError='El Rutero vigente todavía no tiene índice de Status; vuelve a cargarlo desde Vista rápida con la versión 5.23.';return;}
  statusIndex=new Map(index.map(x=>[transportKey(x[0]),String(x[1]??'')]).filter(x=>x[0]));
  statusIndexReady=true;statusIndexBatch=String(meta.batch_id||'');window.refreshSmartSearch?.();
 }catch(e){statusIndexError=e?.message||String(e);}
 finally{statusIndexFetching=false;render();}
}

function cargoSource(){
 // Fuente preferente: los mismos registros usados para construir Cargue y Despacho.
 let direct;try{direct=window.getCargoSyncRecords?.();}catch(e){direct=null;}
 if(Array.isArray(direct)&&direct.length)return direct;
 if(Array.isArray(window.syncCargoRecords)&&window.syncCargoRecords.length)return window.syncCargoRecords;
 if(Array.isArray(direct))return direct.length?direct:cargoFallback;
 return cargoFallback;
}

async function loadCargoSource(force=false){
 if(cargoFetching)return;
 if(!force && cargoSource().length){render();return;}
 if(!force && Date.now()-cargoLastAttempt<8000)return;
 if(typeof sb==='undefined'||!sb)return;
 cargoFetching=true;cargoLastAttempt=Date.now();cargoError='';
 try{
   const {data,error}=await sb.rpc('list_transports_with_category');
   if(error)throw error;
   cargoFallback=Array.isArray(data)?data:[];
 }catch(err){cargoError=err?.message||String(err);}
 finally{cargoFetching=false;render();}
}

function canEdit(){return typeof canModule==='function'&&canModule('cargo','edit');}
function alertStatus(t){$('syncStatMessage').textContent=t;}
async function parseFile(file){
 const buf=await file.arrayBuffer();
 const book=XLSX.read(buf,{type:'array',cellDates:false});
 for(const name of book.SheetNames){
  const sh=book.Sheets[name];if(!sh||!sh['!ref'])continue;
  const grid=XLSX.utils.sheet_to_json(sh,{header:1,defval:'',raw:false,blankrows:false});
  // Estructura real del archivo STAT: B = ID de ruta; D, E, F = códigos de estado.
  const headerIndex=grid.findIndex(r=>Array.isArray(r)&&heading(r[1])==='IDDERUTA');
  if(headerIndex<0)continue;
  const map=new Map();let count=0;
  for(const row of grid.slice(headerIndex+1)){
   const key=transportKey(row[1]);if(!/^\d+$/.test(key))continue;
   count++;
   const codes=[3,4,5].map(i=>String(row[i]??'').trim().toUpperCase());
   const state=codes.some(Boolean)?codes.map(v=>v||'—').join('-'):'';
   if(!map.has(key))map.set(key,new Set());
   if(state)map.get(key).add(state);
  }
  if(!count)continue;
  return {map,items:[...map.entries()].map(([k,v])=>({transporte:k,estados:[...v]})),count,hasStatus:true};
 }
 throw Error('No se encontró la columna B «ID de ruta» con transportes válidos. Verifica el formato de Sincronización STAT.');
}
async function upload(){const file=$('syncStatFile')?.files?.[0];if(!file)return alertStatus('Selecciona Sincronización STAT con la columna B «ID de ruta».');if(!canEdit())return alertStatus('No tienes permiso para cargar STAT.');
 const btn=$('syncStatUpload');btn.disabled=true;alertStatus('Procesando archivo STAT…');
 try{const parsed=await parseFile(file);if(!window.CTShared)throw Error('No está disponible el almacenamiento de Supabase');
  alertStatus('Guardando archivo compartido…');await window.CTShared.save('sync_stat',file.name,parsed.items,{rows:parsed.count,hasStatus:parsed.hasStatus});
  statRows=parsed.map;statLoaded=true;statFile=file.name;
  alertStatus(`${file.name}: ${parsed.count.toLocaleString('es-CO')} filas leídas; ${parsed.items.length.toLocaleString('es-CO')} transportes. Guardado en Supabase.`);render();
 }catch(err){alertStatus('No se pudo cargar STAT: '+err.message);}finally{btn.disabled=false;}
}
async function clearDailyStat(){
 if(!canEdit()){alertStatus('No tienes permiso para eliminar el archivo STAT.');return;}
 if(!window.CTShared){alertStatus('No está disponible la conexión con Supabase.');return;}
 if(!window.confirm('¿Eliminar el archivo Sincronización STAT guardado para TODOS los usuarios? No se borrarán los viajes de Cargue y Despacho ni el Rutero de Vista rápida.'))return;
 const button=$('syncStatClearBtn');button.disabled=true;
 try{
  alertStatus('Eliminando STAT anterior de Supabase…');
  await CTShared.clear('sync_stat');
  statRows=new Map();statLoaded=false;statFile='';$('syncStatFile').value='';
  alertStatus('STAT anterior eliminado de Supabase. Selecciona el archivo actualizado y pulsa Cargar STAT.');
  render();
 }catch(err){alertStatus('No se pudo eliminar el STAT: '+(err?.message||err));}
 finally{button.disabled=false;}
}
async function restore(){if(loading||statLoaded||!window.CTShared)return;loading=true;
 try{const data=await window.CTShared.load('sync_stat');if(!data){alertStatus('Sin archivo STAT cargado.');return;}
 const map=new Map();for(const r of data.items||[])if(transportKey(r.transporte))map.set(transportKey(r.transporte),new Set((r.estados||[]).map(String)));
 statRows=map;statLoaded=true;statFile=data.filename||'STAT';alertStatus('STAT compartido: '+statFile+' · '+map.size.toLocaleString('es-CO')+' transportes.');render();
 }catch(err){alertStatus('No se pudo recuperar STAT: '+err.message);}finally{loading=false;}}
let currentSyncRows=[];
function render(){if(!$('syncRows'))return;try{
 // SOURCE OF TRUTH: Cargue y Despacho, not Rutero / STAT.
 const cargo=cargoSource();
 let report={headers:[],rows:[],loaded:false};try{report=window.getRuteroSyncData?.()||report;}catch(err){console.warn('Vista rápida no disponible',err);}
 const ti=ix(report.headers||[],['TRANSPORTE','TRANSPORT','NUMERODETRANSPORTE','NROTRANSPORTE']);
 // Vista rápida de despacho: Transporte por encabezado; Status = columna I del Excel (índice 8).
 // Si los encabezados cambian de posición, el cruce sigue identificando Transporte por nombre.
 const statusHeaderIndex=ix(report.headers||[],['STATUS']);
 // El archivo nuevo ubica Status en I; algunos Ruteros anteriores tienen Cant en I.
 // Verificar el encabezado antes de interpretar una cantidad como estado.
 const si=heading(report.headers?.[8])==='STATUS' ? 8 : statusHeaderIndex;
 const quick=new Map();
 if(report.loaded&&ti>=0&&si>=0){
   for(const r of (Array.isArray(report.rows)?report.rows:[])){
     if(!Array.isArray(r))continue;
     const k=transportKey(r[ti]); // S8009001225 y 8009001225 son equivalentes.
     if(!k)continue;
     if(!quick.has(k))quick.set(k,new Set());
     const status=String(r[si]??'').trim();
     if(status)quick.get(k).add(status);
   }
 }
 const q=norm($('syncSearch')?.value),f=$('syncFilter')?.value||'';
 const rows=cargo.filter(r=>r&&typeof r==='object').map(r=>{const key=transportKey(r.transporte);const statuses=quick.get(key);const status=statusIndex.get(key)|| (statuses?.size?[...statuses].join(' · '):'—');
  let result;if(statLoaded){const vals=statRows.get(key);result=vals?([...(vals||[])].filter(Boolean).join(' · ')||'En STAT'):'No está en STAT';}
  else result=!report.loaded?'Sin archivo':statuses?.size?'Sincronizado':'No encontrado';
  return {transporte:displayedTransport(r.transporte),viaje:r.viaje??'',placa:r.placa??'',status,result};});
 currentSyncRows=rows;
 const filtered=rows.filter(r=>(!f||r.result===f)&&(!q||[r.transporte,r.viaje,r.placa,r.status,r.result].some(v=>norm(v).includes(q))));
 $('syncTotal').textContent=filtered.length.toLocaleString('es-CO');
 $('syncOk').textContent=filtered.filter(r=>statLoaded?(statRows.has(transportKey(r.transporte))):r.result==='Sincronizado').length.toLocaleString('es-CO');
 $('syncMissing').textContent=filtered.filter(r=>statLoaded?!statRows.has(transportKey(r.transporte)):r.result==='No encontrado').length.toLocaleString('es-CO');
 let syncInfo='';
 if(cargoError&&!cargo.length)syncInfo='No se pudo consultar Cargue y Despacho: '+cargoError;
 else if(!cargo.length)syncInfo=cargoFetching?'Consultando transportes de Cargue y Despacho…':'No hay transportes disponibles. Pulsa Actualizar cruce.';
 else syncInfo=`${filtered.length.toLocaleString('es-CO')} registros de Cargue y Despacho visibles. ${statusIndexReady?'Status consultado desde índice de Rutero guardado ('+statusIndex.size+' transportes).':report.loaded&&si>=0?'Status consultado en Vista rápida.':(statusIndexError||'Recuperando índice Status de Rutero…')} ${statLoaded?'Cruce STAT: '+statFile+'.':''}`;
 $('syncMessage').textContent='Motor de sincronización v5.23 · '+syncInfo;
 updateUnsyncedCount();
 $('syncRows').innerHTML=filtered.slice(0,1000).map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td>${esc(r.placa)}</td><td>${esc(r.status)}</td><td><span class="syncBadge ${/^(Sincronizado|En STAT)$/i.test(r.result)?'syncGood':/^(No encontrado|No está en STAT)$/i.test(r.result)?'syncBad':'syncUnknown'}">${esc(r.result)}</span></td></tr>`).join('') || `<tr><td colspan="5" style="text-align:center;padding:20px">${esc(cargoError ? 'Error al consultar transportes: '+cargoError : cargoFetching ? 'Cargando transportes...' : 'No hay transportes disponibles. Pulsa Actualizar cruce.')}</td></tr>`;
 }catch(err){const msg='Error al presentar Sincronización: '+(err?.message||String(err));const m=$('syncMessage');if(m)m.textContent=msg;const body=$('syncRows');if(body)body.innerHTML='<tr><td colspan="5">'+esc(msg)+'</td></tr>';console.error('[Sincronización]',err);}
}

// Solo alertar estados STAT incompletos de transportes con Status 5 en Rutero.
function hasStatusFive(status){return String(status??'').split(/[·,;|\/\s]+/).some(v=>v.trim()==='5');}
function isSyncAlert(r){return hasStatusFive(r.status)&&String(r.result||'').trim().toUpperCase()!=='S-S-U';}
function openUnsynced(){
 const modal=$('syncAlertModal');const holder=$('syncAlertItems');if(!modal||!holder)return;
 const invalid=currentSyncRows.filter(isSyncAlert);
 $('syncAlertTotal').textContent=invalid.length.toLocaleString('es-CO');
 holder.innerHTML=invalid.length?invalid.map((r,i)=>`<details class="syncAlertItem"><summary><strong>${esc(r.transporte)}</strong><span>${esc(r.result)}</span></summary><dl><div><dt>Transporte</dt><dd>${esc(r.transporte)}</dd></div><div><dt>Placa</dt><dd>${esc(r.placa)}</dd></div><div><dt>Viaje</dt><dd>${esc(r.viaje)}</dd></div><div><dt>Estado sincronización</dt><dd>${esc(r.result)}</dd></div></dl></details>`).join(''):'<p>No hay transportes con Status 5 y sincronización diferente de S-S-U.</p>';
 modal.classList.remove('hidden');document.body.classList.add('syncAlertOpen');$('syncAlertClose')?.focus();
}
function closeUnsynced(){ $('syncAlertModal')?.classList.add('hidden');document.body.classList.remove('syncAlertOpen');$('syncAlertBtn')?.focus(); }
function updateUnsyncedCount(){const n=currentSyncRows.filter(isSyncAlert).length;const el=$('syncAlertCount');if(el)el.textContent=n.toLocaleString('es-CO');}

window.getTransportRuteroStatus=(transporte)=>statusIndex.get(transportKey(transporte))||'';window.renderSyncReport=()=>{render();if(!cargoSource().length)void loadCargoSource();void loadStatusIndex();};window.loadSyncStatReport=()=>{render();void loadCargoSource(!cargoSource().length);void restore();void loadStatusIndex();};
window.addEventListener('rutero-updated',()=>{statusIndexReady=false;void loadStatusIndex(true);render();});
window.addEventListener('cargo-records-updated',render);
window.addEventListener('cargo-records-updated',()=>{cargoFallback=[];render();});
window.addEventListener('focus',()=>{if(!$('syncModuleView')?.classList.contains('hidden')){render();void loadCargoSource();}});
function initializeSync514(){for(const id of ['syncSearch','syncFilter'])$(id)?.addEventListener('input',render);$('syncStatUpload')?.addEventListener('click',upload);$('syncStatClearBtn')?.addEventListener('click',clearDailyStat);$('syncRefresh')?.addEventListener('click',()=>{void loadCargoSource(true);void restore();void loadStatusIndex(true);render();});$('syncAlertBtn')?.addEventListener('click',openUnsynced);$('syncAlertClose')?.addEventListener('click',closeUnsynced);$('syncAlertModal')?.addEventListener('click',e=>{if(e.target===$('syncAlertModal'))closeUnsynced();});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('syncAlertModal')?.classList.contains('hidden'))closeUnsynced();});render();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initializeSync514,{once:true});else initializeSync514();
})();

const cfg=window.APP_CONFIG||{};let sb=null,user=null,onlineTrips=[],records=[],pendingImportRows=null,pendingTapaRows=null,tapaRecords=[],loginPin="",selectedTransportId=null,selectedCheckinId=null,activeModule="main",activeTopModule="cargo",activeCategory="Viajes HYM y URBAN",selectedCategories=[],categoryMultiMode=false,permissions={is_admin:false,can_upload:false,can_create_users:false,can_delete_all_transports:false,can_arm:false,can_count:false,can_load:false,can_edit_availability:false,can_edit_physical_plate:false,can_manage_checkin:false};window.getCargoSyncRecords = () => records;
const $=id=>document.getElementById(id);
function toast(m){$("toast").textContent=m;$("toast").classList.add("show");setTimeout(()=>$("toast").classList.remove("show"),3000)}
function configured(){return cfg.SUPABASE_URL&&!cfg.SUPABASE_URL.includes("TU-PROYECTO")&&cfg.SUPABASE_ANON_KEY&&!cfg.SUPABASE_ANON_KEY.includes("TU_ANON")}
async function init(){if(!configured()){$("loginMsg").innerHTML="Falta configurar <b>config.js</b>. Revisa README.";return}if(!window.supabase?.createClient){$("loginMsg").textContent="No cargó el servicio de inicio de sesión. Verifica internet y recarga la página.";return;}sb=supabase.createClient(cfg.SUPABASE_URL,cfg.SUPABASE_ANON_KEY);$("mainModuleBtn").onclick=()=>switchModule("main");$("tapaModuleBtn").onclick=()=>switchModule("tapa");$("cargoPrimaryBtn").onclick=()=>switchModule(activeTopModule==="cargo"?(activeModule==="tapa"?"tapa":"main"):"main");$("checkinPrimaryBtn").onclick=()=>switchModule("checkin");$("brujulaPrimaryBtn").onclick=()=>switchModule("brujula");$("syncPrimaryBtn").onclick=()=>switchModule("sync");$("tapaFamilyFilter").onchange=renderTapa;$("tapaClientFilter").onchange=renderTapa;$("tapaSearch").oninput=renderTapa;$("loginBtn").onclick=login;$("pin").onkeydown=e=>e.key==="Enter"&&login();$("logout").onclick=logout;$("ctReturnsEntry").onclick=ctOpenReturns;$("ctFinanceEntry").onclick=ctOpenFinance;$("ctValidacionesEntry").onclick=ctOpenValidaciones;$("ctReturnsBack").onclick=ctCloseReturns;$("users").addEventListener("click",e=>{const finance=e.target.closest('.ctFinanceUserToggle');if(finance){void ctChangeFinancePermission(finance.dataset.user,finance.dataset.enabled!=='true');return;}const b=e.target.closest('.ctReturnsUserToggle');if(!b)return;void ctChangeReturnsPermission(b.dataset.user,b.dataset.enabled!=='true');});$("ctManualRefresh")?.addEventListener("click",()=>{void ctManualRefreshAll()});for(const id of ["smartPlate","smartTransport","smartStatus"])$(id).addEventListener(id==="smartStatus"?"change":"input",renderSmartSearch);$("smartClear").onclick=()=>{for(const id of ["smartPlate","smartTransport","smartStatus"])$(id).value="";renderSmartSearch()};$("smartResults").addEventListener("click",e=>{const b=e.target.closest(".smartGoTo");if(b)openSmartResult(b.dataset.category,b.dataset.transporte)});$("openPorteriaPending").onclick=openPorteriaPending;$("closePorteriaPending").onclick=closePorteriaPending;$("porteriaPendingModal").addEventListener("click",e=>{if(e.target.id==="porteriaPendingModal")closePorteriaPending()});document.addEventListener("keydown",e=>{if(e.key==="Escape")closePorteriaPending()});function updateCategoryTabs(){document.querySelectorAll(".categoryTab[data-category]").forEach(btn=>{const on=categoryMultiMode?selectedCategories.includes(btn.dataset.category):btn.dataset.category===activeCategory;btn.classList.toggle("active",on);btn.setAttribute("aria-selected",String(on))});$("clearCategorySelection").style.display=categoryMultiMode?"inline-block":"none";$("categorySelectionHint").textContent=categoryMultiMode?"Selección múltiple activa: haz clic en las demás pestañas para agregarlas o quitarlas.":"Para combinar pestañas, haz doble clic en la primera pestaña y luego haz clic en las demás."}document.querySelectorAll(".categoryTab[data-category]").forEach(btn=>{btn.addEventListener("dblclick",e=>{e.preventDefault();const cat=btn.dataset.category;if(cat==="Todos"){categoryMultiMode=false;selectedCategories=[];activeCategory="Todos";toast("La pestaña Todos ya reúne todas las categorías.");}else{categoryMultiMode=true;selectedCategories=[cat];activeCategory="__MULTI__";toast("Selección múltiple activada. Ahora elige las demás pestañas.");}updateCategoryTabs();render()});btn.addEventListener("click",()=>{const cat=btn.dataset.category;if(categoryMultiMode){if(cat==="Todos"){categoryMultiMode=false;selectedCategories=[];activeCategory="Todos"}else{selectedCategories=selectedCategories.includes(cat)?selectedCategories.filter(x=>x!==cat):[...selectedCategories,cat];if(!selectedCategories.length){categoryMultiMode=false;activeCategory="Viajes HYM y URBAN"}}}else{activeCategory=cat;selectedCategories=[]}updateCategoryTabs();render()})});$("clearCategorySelection").addEventListener("click",()=>{categoryMultiMode=false;selectedCategories=[];activeCategory="Viajes HYM y URBAN";updateCategoryTabs();render()});updateCategoryTabs();setupCategoryHoverPreview();$("importBtn").onclick=prepareImport;$("confirmImportBtn").onclick=confirmImport;$("cancelImportBtn").onclick=cancelImport;$("previewClose").onclick=cancelImport;$("downloadBtn").onclick=downloadExcel;$("beesCompareBtn").onclick=compareBeesCsv;$("beesClearBtn").onclick=clearBeesReport;$("porteriaUploadBtn").onclick=loadPorteriaExcel;$("porteriaClearBtn").onclick=clearPorteria;$("addUser").onclick=addUser;$("saveModulePermissions").onclick=saveModulePermissionEdits;$("cancelModulePermissions").onclick=()=>$("modulePermissionEditor").classList.add("hidden");$("newRole").onchange=updateRoleOptions;$("deleteAllBtn").onclick=deleteAllTransports;$("cancelDeleteBtn").onclick=closeDeleteModal;$("confirmDeleteBtn").onclick=confirmDeleteAll;$("deleteConfirmModal").addEventListener("click",e=>{if(e.target.id==="deleteConfirmModal")closeDeleteModal()});document.querySelectorAll(".statCard").forEach(card=>{card.addEventListener("click",()=>openStatChart(card.dataset.stat));card.addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();openStatChart(card.dataset.stat)}})});$("statChartClose").onclick=closeStatChart;$("statChartModal").addEventListener("click",e=>{if(e.target.id==="statChartModal")closeStatChart()});document.addEventListener("keydown",e=>{if(e.key==="Escape")closeStatChart()})}
async function login(){
 const msg=$("loginMsg"),button=$("loginBtn"),pin=$("pin").value.trim();
 if(!/^\d{1,8}$/.test(pin)){msg.textContent="Ingresa tu PIN numérico de 1 a 8 dígitos.";return;}
 if(!sb){msg.textContent="No se pudo conectar con la base de datos. Recarga la página.";return;}
 button.disabled=true;msg.textContent="Verificando acceso…";
 const limited=async (name,args)=>{
  let timeoutId;
  try{return await Promise.race([sb.rpc(name,args),new Promise((_,reject)=>{timeoutId=setTimeout(()=>reject(new Error("Tiempo de espera agotado. Verifica la conexión e intenta de nuevo.")),18000);})]);}
  finally{clearTimeout(timeoutId);}
 };
 try{
  const {data,error}=await limited("login_by_pin",{p_pin:pin});
  if(error){console.error("Error autenticación:",error);msg.textContent="No se pudo verificar el PIN por un error del servidor. Intenta de nuevo o contacta al administrador.";return;}
  if(!Array.isArray(data)||!data.length){msg.textContent="PIN incorrecto o usuario desactivado.";return;}
  const pr=await limited("get_user_permissions_by_pin",{p_pin:pin});
  if(pr.error||!pr.data?.length){console.error("Error permisos base:",pr.error);msg.textContent="Tu PIN fue aceptado, pero no fue posible consultar los permisos principales. Contacta al administrador.";return;}
  const dp=await limited("get_user_all_permissions_by_pin",{p_pin:pin});
  if(dp.error||!dp.data?.length){console.error("Error permisos adicionales:",dp.error);msg.textContent="Tu PIN fue aceptado, pero no fue posible consultar los permisos adicionales. Contacta al administrador.";return;}
  user=data[0];loginPin=pin;permissions={...pr.data[0],...dp.data[0]};
  try{sessionStorage.setItem("transport_user",JSON.stringify(user));}catch(e){console.warn("SessionStorage:",e);}
  await loadModuleRights();await ctLoadReturnsPermission();await ctFetchActions();await showApp();msg.textContent="";
 }catch(e){console.error("No se completó el inicio de sesión:",e);msg.textContent=e.message?.includes("Tiempo de espera")?e.message:"No se pudo completar el ingreso. Verifica tu conexión y vuelve a intentarlo.";}
 finally{button.disabled=false;}
}
const DEFAULT_MODULE_RIGHTS={cargo_view:true,cargo_edit:true,checkin_view:true,checkin_edit:true,rutero_view:true,rutero_edit:true};
let moduleRights={...DEFAULT_MODULE_RIGHTS},selectedPermissionUser=null;
function canModule(module,level){return permissions.is_admin||!!moduleRights[module+"_"+level]}
async function loadModuleRights(){if(permissions.is_admin){moduleRights={cargo_view:true,cargo_edit:true,checkin_view:true,checkin_edit:true,rutero_view:true,rutero_edit:true};return}const{data,error}=await sb.rpc("module_permissions_for_pin",{p_pin:loginPin});if(error){moduleRights={cargo_view:false,cargo_edit:false,checkin_view:false,checkin_edit:false,rutero_view:false,rutero_edit:false};toast("Falta ejecutar SQL de permisos v5.0: "+error.message);return}moduleRights={...DEFAULT_MODULE_RIGHTS,...(data?.[0]||{})};}
function applyModuleRights(){for(const [key,id] of [["cargo","cargoPrimaryBtn"],["checkin","checkinPrimaryBtn"],["rutero","brujulaPrimaryBtn"]])$(id).classList.toggle("hidden",!canModule(key,"view"));$("syncPrimaryBtn").classList.toggle("hidden",!canModule("cargo","view"));$("noModuleAccess").classList.toggle("hidden",["cargo","checkin","rutero"].some(x=>canModule(x,"view")));$("uploadPanel").classList.toggle("hidden",!(canModule("cargo","edit")&&(permissions.can_upload||permissions.is_admin)));$("importBtn").disabled=!canModule("cargo","edit");$("beesCompareBtn").disabled=!canModule("checkin","edit");$("porteriaUploadBtn").disabled=!canModule("checkin","edit");$("brujulaUploadBtn").disabled=!canModule("rutero","edit");if($("brujulaClearBtn"))$("brujulaClearBtn").disabled=!canModule("rutero","edit");if($("syncStatClearBtn"))$("syncStatClearBtn").disabled=!canModule("cargo","edit");if(!canModule("checkin","edit")){ $("beesClearBtn").disabled=true;$("porteriaClearBtn").disabled=true;}}
function setupCheckinCollapses(){document.querySelectorAll(".checkinCollapse").forEach(btn=>{const el=$(btn.dataset.collapse),key="ct_checkin_collapsed_"+btn.dataset.collapse;let hidden=false;try{hidden=localStorage.getItem(key)==="1"}catch(e){}el.classList.toggle("hidden",hidden);btn.textContent=hidden?"Mostrar":"Ocultar";btn.onclick=()=>{const hide=!el.classList.contains("hidden");el.classList.toggle("hidden",hide);btn.textContent=hide?"Mostrar":"Ocultar";try{localStorage.setItem(key,hide?"1":"0")}catch(e){}}})}
// Presentaciones breves: 500 ms totales, solo al ingresar al módulo principal.
let splashTimer=0, splashExitTimer=0;
function playModuleSplash(module){
 const ids={cargo:"cargoSplash",checkin:"checkinSplash",rutero:"ruteroSplash"};
 const id=ids[module];if(!id||!canModule(module,"view"))return;
 clearTimeout(splashTimer);clearTimeout(splashExitTimer);
 for(const splashId of Object.values(ids)){
  const previous=$(splashId);if(previous){previous.classList.add("hidden");previous.classList.remove("depart")}
 }
 const el=$(id);if(!el)return;
 el.classList.remove("hidden","depart");void el.offsetWidth;
 splashTimer=setTimeout(()=>{
  el.classList.add("depart");
  splashExitTimer=setTimeout(()=>el.classList.add("hidden"),200);
 },300);
}
function playCargoSplash(){playModuleSplash("cargo")}
async function editModulePermissions(id,name){if(!permissions.is_admin)return;selectedPermissionUser=id;$("modulePermissionWho").textContent=name;const{data,error}=await sb.rpc("admin_get_module_permissions",{p_admin_pin:loginPin,p_user_id:id});if(error){toast(error.message);return}const rights=data?.[0]||DEFAULT_MODULE_RIGHTS;for(const mod of ["Cargo","Checkin","Rutero"])for(const mode of ["View","Edit"])$("edit"+mod+mode).checked=!!rights[mod.toLowerCase()+"_"+mode.toLowerCase()];$("modulePermissionEditor").classList.remove("hidden");$("modulePermissionEditor").scrollIntoView({behavior:"smooth",block:"center"})}
async function saveModulePermissionEdits(){if(!selectedPermissionUser||!permissions.is_admin)return;const rights={};for(const mod of ["Cargo","Checkin","Rutero"])for(const mode of ["View","Edit"])rights[mod.toLowerCase()+"_"+mode.toLowerCase()]=$("edit"+mod+mode).checked;const{error}=await sb.rpc("admin_set_module_permissions",{p_admin_pin:loginPin,p_user_id:selectedPermissionUser,p_rights:rights});if(error){toast(error.message);return}$("modulePermissionEditor").classList.add("hidden");selectedPermissionUser=null;toast("Permisos de módulos guardados");await loadUsers()}
function setupAdminModuleToggles(){document.querySelectorAll(".moduleToggle").forEach(btn=>{btn.classList.toggle("hidden",!permissions.is_admin);const panel=document.querySelector(`[data-module-panel="${btn.dataset.target}"]`);if(!panel)return;let collapsed=false;try{collapsed=localStorage.getItem(`transport_module_${btn.dataset.target}_collapsed`)==="1"}catch(e){}panel.classList.toggle("collapsed",collapsed);btn.textContent=collapsed?"Desplegar":"Ocultar";btn.onclick=()=>{if(!permissions.is_admin)return;const next=!panel.classList.contains("collapsed");panel.classList.toggle("collapsed",next);btn.textContent=next?"Desplegar":"Ocultar";try{localStorage.setItem(`transport_module_${btn.dataset.target}_collapsed`,next?"1":"0")}catch(e){}}})}
// v5.57: módulo independiente; los permisos se consultan en Supabase, no se confía solo en ocultar botones.
let ctFinanceAllowed=false;
async function ctLoadFinancePermission(){ctFinanceAllowed=false;try{if(loginPin){const {data,error}=await sb.rpc('ct_finance_can_access',{p_pin:loginPin});if(error)throw error;ctFinanceAllowed=data===true;}}catch(e){console.warn('Permiso de Liquidación Financiera:',e.message);} $('ctFinanceEntry')?.classList.toggle('hidden',!ctFinanceAllowed);}
function ctOpenFinance(){if(!ctFinanceAllowed){toast('Sin permiso para Liquidación Financiera');return;}window.open('liquidacion-financiera/index.html','_blank','noopener,noreferrer');}
// Validaciones Operación: acceso predeterminado denegado; Supabase determina la visibilidad.
let ctValidacionesAllowed=false;
async function ctLoadValidacionesPermission(){
 ctValidacionesAllowed=false;
 try{if(loginPin){const {data,error}=await sb.rpc('ct_validaciones_can_access',{p_pin:loginPin});if(error)throw error;ctValidacionesAllowed=data===true;}}
 catch(e){console.warn('Permiso de Validaciones Operación:',e.message);}
 $('ctValidacionesEntry')?.classList.toggle('hidden',!ctValidacionesAllowed);
}
async function ctOpenValidaciones(){
 if(!loginPin)return;
 await ctLoadValidacionesPermission();
 if(!ctValidacionesAllowed){toast('Sin permiso para Validaciones Operación');return;}
 window.open('validaciones-operacion/index.html','_blank','noopener,noreferrer');
}
let ctReturnsAllowed=false,ctReturnsOpen=false;
async function ctLoadReturnsPermission(){ctReturnsAllowed=false;try{const {data,error}=await sb.rpc('ct_devoluciones_can_access',{p_pin:loginPin});if(error)throw error;ctReturnsAllowed=data===true;}catch(e){console.warn('No se pudo verificar acceso a Devoluciones:',e.message);ctReturnsAllowed=false;}$('ctReturnsEntry')?.classList.toggle('hidden',!ctReturnsAllowed);await ctLoadFinancePermission();await ctLoadValidacionesPermission();}
async function ctOpenReturns(){if(!loginPin)return;await ctLoadReturnsPermission();if(!ctReturnsAllowed){toast('Sin permiso para Seguimiento de Devoluciones');return;}ctReturnsOpen=true;$('app').classList.add('hidden');$('ctReturnsModule').classList.remove('hidden');$('ctReturnsBack').classList.remove('hidden');$('ctReturnsEntry').classList.add('hidden');$('ctManualRefresh')?.classList.add('hidden');window.scrollTo({top:0,behavior:'smooth'});}
function ctCloseReturns(){ctReturnsOpen=false;$('ctReturnsModule').classList.add('hidden');$('app').classList.remove('hidden');$('ctReturnsBack').classList.add('hidden');$('ctReturnsEntry').classList.toggle('hidden',!ctReturnsAllowed);$('ctManualRefresh')?.classList.remove('hidden');window.scrollTo({top:0,behavior:'smooth'});}
async function ctPaintReturnsAdminPermissions(){if(!permissions.is_admin)return;const {data,error}=await sb.rpc('ct_devoluciones_admin_list',{p_admin_pin:loginPin});if(error){console.warn('Permisos devoluciones:',error.message);return;}const rights=new Map((data||[]).map(r=>[String(r.user_id),!!r.can_view]));document.querySelectorAll('.ctReturnsUserToggle').forEach(b=>{const on=rights.get(b.dataset.user)===true;b.dataset.enabled=String(on);b.textContent=on?'📦 Devoluciones: Permitido':'📦 Devoluciones: Sin acceso';b.classList.toggle('ctReturnsGranted',on);});}
async function ctPaintFinanceAdminPermissions(){
 if(!permissions.is_admin)return;
 const {data,error}=await sb.rpc('ct_finance_admin_list',{p_admin_pin:loginPin});
 if(error){toast('Permisos de Liquidación: '+error.message);return;}
 const rights=new Map((data||[]).map(r=>[String(r.user_id),!!r.can_view]));
 document.querySelectorAll('.ctFinanceUserToggle').forEach(b=>{
  const on=rights.get(b.dataset.user)===true;
  b.dataset.enabled=String(on);
  b.textContent=on?'💵 Financiera: Permitido':'💵 Financiera: Sin acceso';
  b.classList.toggle('ctReturnsGranted',on);
 });
}
async function ctChangeFinancePermission(id,enabled){
 if(!permissions.is_admin)return;
 const {error}=await sb.rpc('ct_finance_admin_set_access',{p_admin_pin:loginPin,p_user_id:id,p_can_view:enabled});
 if(error){toast('No se pudo guardar el permiso financiero: '+error.message);return;}
 toast(enabled?'Acceso a Liquidación Financiera autorizado':'Acceso a Liquidación Financiera revocado');
 await ctPaintFinanceAdminPermissions();
}
async function ctChangeReturnsPermission(id,enabled){if(!permissions.is_admin)return;const {error}=await sb.rpc('ct_devoluciones_admin_set_access',{p_admin_pin:loginPin,p_user_id:id,p_can_view:enabled});if(error){toast('No se pudo guardar permiso: '+error.message);return;}toast(enabled?'Acceso a Devoluciones autorizado':'Acceso a Devoluciones revocado');await ctPaintReturnsAdminPermissions();}
async function showApp(){$("ctReturnsEntry")?.classList.toggle("hidden",!ctReturnsAllowed);$("ctFinanceEntry")?.classList.toggle("hidden",!ctFinanceAllowed);$("ctValidacionesEntry")?.classList.toggle("hidden",!ctValidacionesAllowed);$("ctManualRefresh")?.classList.remove("hidden");$("login").classList.add("hidden");$("app").classList.remove("hidden");$("logout").classList.remove("hidden");$("userName").textContent=user.name;$("connection").textContent="● Conectado";if(permissions.can_upload||permissions.is_admin)$("uploadPanel").classList.remove("hidden");if(permissions.can_create_users||permissions.is_admin)$("usersPanel").classList.remove("hidden");if(permissions.can_delete_all_transports||permissions.is_admin)$("deletePanel").classList.remove("hidden");if(permissions.is_admin||permissions.can_create_users){$("checkinPermissionLabel").classList.remove("hidden")}$("historyPanel").classList.toggle("hidden",!permissions.is_admin);setupAdminModuleToggles();setupCheckinCollapses();ctSetupActionSettings();applyModuleRights();if(permissions.is_admin){$("permissionLabel").classList.remove("hidden");$("createPermissionLabel").classList.remove("hidden");$("deletePermissionLabel").classList.remove("hidden");$("statusPermissionLabels").classList.remove("hidden")}updateRoleOptions();await refresh();await refreshCheckin();await loadSharedCheckin();subscribe();if(permissions.is_admin)loadUsers();if(!canModule("cargo","view"))switchModule(canModule("checkin","view")?"checkin":"brujula");else playCargoSplash()}
function logout(){loginPin="";sessionStorage.clear();location.reload()}
function updateRoleOptions(){const admin=permissions.is_admin;const opt=$("newRole").querySelector('option[value="Administrador"]');if(opt)opt.disabled=!admin;if(!admin&&$("newRole").value==="Administrador")$("newRole").value="Operativo";if(!admin){$("newCanUpload").checked=false;$("newCanCreateUsers").checked=false;$("newCanDeleteAll").checked=false;["newCanArm","newCanCount","newCanLoad","newCanAvailability","newCanPhysicalPlate"].forEach(id=>$(id).checked=false)}}

// v5.46: vista flotante al hacer clic en una categoría, conservando su filtro normal.
function setupCategoryHoverPreview(){
 const tabs=[...document.querySelectorAll('.categoryTab[data-category]')];
 if(!tabs.length||document.getElementById('ctCategoryHover'))return;
 const panel=document.createElement('section');
 panel.id='ctCategoryHover';panel.className='ctCategoryHover';
 panel.setAttribute('role','region');panel.setAttribute('aria-label','Vista previa de transportes por categoría');
 panel.innerHTML=`<div class="ctCategoryHoverTop"><div><div class="ctCategoryHoverEyebrow">VISTA RÁPIDA POR CATEGORÍA</div><h3 id="ctCategoryHoverTitle">Transportes</h3><p id="ctCategoryHoverCount">0 transportes</p></div><button type="button" class="ctCategoryHoverClose" aria-label="Cerrar vista previa">✕</button></div><div class="ctCategoryHoverScroll"><table><thead><tr><th>Transporte</th><th>Viaje</th><th>Placa</th><th>Estado</th></tr></thead><tbody id="ctCategoryHoverRows"></tbody></table></div><p class="ctCategoryHoverFoot">La pestaña filtra la tabla principal. Desplázate aquí para revisar todos los transportes.</p>`;
 document.body.appendChild(panel);
 let opened=null,hideTimer=0;
 const esc=(value)=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 function rowsFor(category){
   const excel=records.filter(r=>r.estado!=='Cargado');
   const online=onlineTrips.filter(r=>r.estado!=='Cargado').map(r=>({...r,categoria:'Viajes en línea'}));
   const combined=[...excel,...online];
   return category==='Todos'?combined:combined.filter(r=>(r.categoria||'Sin categoría')===category);
 }
 function show(category){
   clearTimeout(hideTimer);opened=category;
   const list=rowsFor(category);
   panel.querySelector('#ctCategoryHoverTitle').textContent=category;
   panel.querySelector('#ctCategoryHoverCount').textContent=`${list.length.toLocaleString('es-CO')} transportes`;
   const body=panel.querySelector('#ctCategoryHoverRows');
   body.innerHTML=list.length?list.map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td>${esc(r.placa)}</td><td>${esc(r.estado||'Pendiente')}</td></tr>`).join(''):'<tr><td colspan="4" class="ctCategoryEmpty">No hay transportes pendientes en esta categoría.</td></tr>';
   panel.querySelector('.ctCategoryHoverScroll').scrollTop=0;
   panel.classList.add('ctCategoryHoverVisible');
 }
 function hide(){clearTimeout(hideTimer);opened=null;panel.classList.remove('ctCategoryHoverVisible')}
 function delayHide(){clearTimeout(hideTimer);hideTimer=setTimeout(hide,280)}
 tabs.forEach(tab=>{
  tab.addEventListener('click',()=>{
    const category=tab.dataset.category;
    if(opened===category)hide(); else show(category);
  });
 });
 // El recuadro permanece abierto hasta cerrar, cambiar de pestaña o hacer clic fuera.
 panel.querySelector('.ctCategoryHoverClose').addEventListener('click',hide);
 document.addEventListener('keydown',e=>{if(e.key==='Escape')hide()});
 document.addEventListener('click',e=>{if(opened&&!e.target.closest('.ctCategoryHover')&&!e.target.closest('.categoryTab[data-category]'))hide()});
}

async function loadOnlineTrips(){
 if(!sb||!loginPin)return;
 const result=await sb.rpc('ct_online_trip_list_secure',{p_pin:loginPin});
 if(result.error){console.warn('[Viajes en línea]',result.error.message);onlineTrips=[];return;}
 onlineTrips=Array.isArray(result.data)?result.data:[];
}

async function refresh(){const{data,error}=await sb.rpc("list_transports_with_category");if(error){toast(error.message);return}records=data||[];
// La fuente de la categoria es el Excel de Cargue y Despacho, no el texto del viaje.
const categoryResult=await sb.rpc('ct_transport_category_rows_secure',{p_pin:loginPin});
if(categoryResult.error){
  toast('Categorías no disponibles: ejecuta ACTUALIZACION_SUPABASE_v5_29.sql. '+categoryResult.error.message);
  console.error('[Categorías]',categoryResult.error);
  records=records.map(r=>({...r,categoria:r.categoria||'Sin categoría'}));
}else{
  const ck=(t,v,pl)=>[t,v,pl].map(x=>String(x??'').trim().toUpperCase()).join('|');
  const byKey=new Map((categoryResult.data||[]).map(x=>[ck(x.transporte,x.viaje,x.placa),x.categoria]));
  records=records.map(r=>({...r,categoria:byKey.get(ck(r.transporte,r.viaje,r.placa))||r.categoria||'Sin categoría'}));
  const noCat=records.filter(r=>r.categoria==='Sin categoría').length;
  if(noCat){toast(`${noCat} viajes sin categoría. Reimporta PRB07.xlsx después de instalar SQL v5.29.`);}
}
await loadOnlineTrips();
const extra=await sb.rpc("list_transport_vehicle_fields");if(!extra.error){const fields=new Map((extra.data||[]).map(x=>[String(x.id),x]));records=records.map(r=>({...r,...(fields.get(String(r.id))||{})}))}window.syncCargoRecords = records;
window.dispatchEvent(new CustomEvent('cargo-records-updated',{detail:{count:records.length}}));
render();renderSmartSearch();if(permissions.is_admin)loadHistory()}
window.refreshSmartSearch=()=>renderSmartSearch();function renderSmartSearch(){
 const panel=$("smartSearchPanel");if(!panel)return;
 const plate=$("smartPlate").value.trim().toUpperCase().replace(/\s+/g,"");
 const transport=$("smartTransport").value.trim().toUpperCase().replace(/^S(?=\d)/,"");
 const option=$("smartStatus").value;
 const summary=$("smartSummary"),wrap=$("smartResultsWrap"),tbody=$("smartResults");
 if(!plate&&!transport&&!option){summary.textContent="Escribe una placa, un número de transporte o selecciona un Status para consultar todas las categorías.";wrap.classList.add("hidden");tbody.innerHTML="";return}
 const all=[...records,...onlineTrips.map((r,i)=>({...r,id:-1000000-Number(r.id),categoria:"Viajes en línea",estado:r.estado||"Pendiente",is_online:true}))];
 const getStatus=window.getTransportRuteroStatus||(()=>"");
 const matching=all.filter(r=>{
  const key=String(r.transporte||"").toUpperCase().replace(/^S(?=\d)/,"");
  const actualPlate=String(r.placa||"").toUpperCase().replace(/\s+/g,"");
  const actualPhysical=String(r.cambio_placa||"").toUpperCase().replace(/\s+/g,"");
  if(plate&&!actualPlate.includes(plate)&&!actualPhysical.includes(plate))return false;
  if(transport&&!key.includes(transport))return false;
  if(option){const [kind,val]=option.split(":");if(kind==="estado"&&String(r.estado||"").toUpperCase()!==val.toUpperCase())return false;
  if(kind==="status"){const status=String(getStatus(r.transporte)||"");if(!status.split(/[·,;|\/\s]+/).includes(val))return false;}}
  return true;
 });
 summary.textContent=`${matching.length.toLocaleString("es-CO")} transporte(s) encontrados en ${new Set(matching.map(r=>r.categoria||"Sin categoría")).size} pestaña(s).`+(option.startsWith("status:")&&!matching.length?" Si acabas de cargar el Rutero, actualiza el reporte para recuperar su índice de Status.":"");
 wrap.classList.remove("hidden");
 tbody.innerHTML=matching.slice(0,400).map(r=>`<tr><td><b>${esc(r.transporte)}</b></td><td>${esc(r.viaje)}</td><td>${esc(r.placa)}</td><td><span class="smartCategory">${esc(r.categoria||"Sin categoría")}</span></td><td>${esc(r.estado||"—")}</td><td>${esc(getStatus(r.transporte)||"—")}</td><td><button type="button" class="secondary smartGoTo" data-category="${esc(r.categoria||"Sin categoría")}" data-transporte="${esc(r.transporte)}">Ver pestaña</button></td></tr>`).join("")||'<tr><td colspan="7">No hay coincidencias en las categorías cargadas.</td></tr>';
 if(matching.length>400)summary.textContent+=" Se muestran los primeros 400 resultados; escribe más caracteres para precisar.";
}
function openSmartResult(category,transporte){
 const tabs=[...document.querySelectorAll(".categoryTab[data-category]")];
 const dest=tabs.find(el=>el.dataset.category===category);
 categoryMultiMode=false;selectedCategories=[];activeCategory=dest?category:"Todos";
 document.querySelectorAll(".categoryTab[data-category]").forEach(btn=>{const on=btn.dataset.category===activeCategory;btn.classList.toggle("active",on);btn.setAttribute("aria-selected",String(on))});
 $("clearCategorySelection").style.display="none";
 $("categorySelectionHint").textContent=`Categoría localizada: ${category}. Transporte: ${transporte}.`;
 render();
 const row=[...$("rows").querySelectorAll("tr")].find(tr=>tr.cells?.[1]?.textContent.trim().toUpperCase()===String(transporte).toUpperCase());
 if(row){row.classList.add("smartFoundRow");row.scrollIntoView({behavior:"smooth",block:"center"});}else{$("mainModuleView").scrollIntoView({behavior:"smooth",block:"start"})}
}
function render(){
const allRecords=[...records,...onlineTrips.map(r=>({id:-1000000-Number(r.id),transporte:r.transporte,viaje:r.viaje,placa:r.placa,categoria:"Viajes en línea",estado:r.estado||"Pendiente",novedad:"",is_online:true,created_at:r.created_at}))];
const filterCategory=arr=>activeCategory==="Todos"?arr:(categoryMultiMode&&selectedCategories.length?arr.filter(r=>selectedCategories.includes(r.categoria||"Sin categoría")):arr.filter(r=>(r.categoria||"Sin categoría")===activeCategory));
// Las tarjetas y las barras utilizan el universo completo, incluso los Cargados que ya están en CHECKIN.
const reportList=filterCategory(allRecords);
// La tabla de trabajo conserva únicamente los viajes aún pendientes de CHECKIN.
const list=reportList.filter(r=>r.estado!=="Cargado");
window.currentCargoFilteredRecords=list;
window.currentCargoAllStatusRecords=reportList;
window.syncCargoRecords=records;
window.renderSyncReport?.();updatePorteriaPendingCount();
const total=reportList.length;
const armadoCount=reportList.filter(r=>r.estado==="Armado").length;
const contadoCount=reportList.filter(r=>r.estado==="Contado").length;
const cargadoCount=reportList.filter(r=>r.estado==="Cargado").length;
const pendienteCount=total-cargadoCount;
$("total").textContent=total;
$("pendientes").textContent=pendienteCount;
$("armados").textContent=armadoCount;
$("contados").textContent=contadoCount;
$("cargados").textContent=cargadoCount;
const pct=n=>total?Math.round(n*100/total):0;
for(const [name,count] of [["Pendiente",pendienteCount],["Armado",armadoCount],["Contado",contadoCount],["Cargado",cargadoCount]]){
$("progress"+name+"Count").textContent=count;
$("progress"+name+"Pct").textContent=pct(count)+"%";
$("progress"+name+"Bar").style.width=pct(count)+"%";
}
$("rows").innerHTML=list.map(r=>{const has=String(r.novedad||"").trim()!=="";const ev=r.estado_vehiculo||"";const selected=String(selectedTransportId)===String(r.id);const canAvailability=canModule("cargo","edit")&&(permissions.is_admin||permissions.can_edit_availability);const canPlate=canModule("cargo","edit")&&(permissions.is_admin||permissions.can_edit_physical_plate);const nextAction=ctNextAction(r.estado);const canAction=canModule("cargo","edit")&&!!nextAction&&ctActionAllowed(nextAction);return `<tr class="${has?"hasNovelty ":""}${selected?"transportSelected":""}"><td>${`<button type="button" class="${selected?"selectedTransportBtn":"secondary selectTransportBtn"}" onclick="selectTransport(${r.id})">${selected?"✓ Seleccionado":"Seleccionar"}</button>`}</td><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td><b>${esc(r.placa)}</b></td><td>${has?`<span class="novelty">${esc(r.novedad)}</span>`:"—"}</td><td><select class="vehicleState ${vehicleClass(ev)}" aria-label="Disponibilidad De Vehiculos para ${esc(r.placa)}" ${(r.is_online||!selected||!canAvailability)?"disabled":""} onchange="saveVehicleField(${r.id},'estado_vehiculo',this.value)"><option value="">Seleccionar...</option><option ${ev==="Bodega"?"selected":""}>Bodega</option><option ${ev==="Sobredimensionado"?"selected":""}>Sobredimensionado</option><option ${ev==="OK para utilizar"?"selected":""}>OK para utilizar</option></select></td><td><input class="plateChange" aria-label="Placa Fisica para ${esc(r.placa)}" ${(r.is_online||!selected||!canPlate)?"disabled":""} maxlength="8" value="${esc(r.cambio_placa||"")}" placeholder="Máx. 8 caracteres" onchange="this.value=normalizePhysicalPlate(this.value);saveVehicleField(${r.id},'cambio_placa',this.value)" oninput="this.value=normalizePhysicalPlate(this.value)"></td><td><span class="badge ${r.estado}">${esc(r.estado)}</span></td><td><button onclick="advance(${r.id})" ${(!selected||!canAction||r.estado==="Cargado")?"disabled":""}>${nextAction?ctActionLabel(nextAction):"—"}</button></td><td>${r.is_online?"<b style=\"color:#047857\">En línea</b>":"—"}</td><td><span class="porteriaBadge ${getPorteriaStatus(r.transporte)==="Salida"?"porteriaYes":getPorteriaStatus(r.transporte)==="Sin salida"?"porteriaNo":"porteriaUnknown"}">${getPorteriaStatus(r.transporte)}</span></td></tr>`}).join("")||'<tr><td colspan="11">No hay registros con estos filtros.</td></tr>';const selectedRecord=allRecords.find(r=>String(r.id)===String(selectedTransportId));if(selectedTransportId&&!selectedRecord)selectedTransportId=null}
function selectTransport(id){selectedTransportId=String(selectedTransportId)===String(id)?null:String(id);const r=Number(id)<0?onlineTrips.find(x=>String(-1000000-x.id)===String(id)):records.find(x=>String(x.id)===String(id));toast(selectedTransportId?`Transporte seleccionado: ${r?.transporte||""} · Viaje ${r?.viaje||""}. Ya puedes modificar los campos permitidos.`:"Selección cancelada.");render()}
function openStatChart(key){const visible=window.currentCargoAllStatusRecords||[];const total=visible.length;const config={total:{label:"Total",rows:visible,color:"#1769aa"},pendientes:{label:"Pendientes por cargar",rows:visible.filter(r=>r.estado!=="Cargado"),color:"#dc2626"},armados:{label:"Armados",rows:visible.filter(r=>r.estado==="Armado"),color:"#f97316"},contados:{label:"Contados",rows:visible.filter(r=>r.estado==="Contado"),color:"#2563eb"},cargados:{label:"Cargados",rows:visible.filter(r=>r.estado==="Cargado"),color:"#22c55e"},novedades:{label:"Con novedad",rows:visible.filter(r=>String(r.novedad||"").trim()!==""),color:"#ef4444"}};const item=config[key];if(!item)return;const count=item.rows.length,percent=total?Math.round(count*100/total):0;$("statChartTitle").textContent="Avance: "+item.label;$("statChartCount").textContent=count.toLocaleString("es-CO");$("statChartPercent").textContent=percent+"% de los registros filtrados";$("statChartFill").style.width=percent+"%";$("statChartFill").style.background=item.color;$("statChartTruck").style.left=`calc(${Math.min(percent,96)}% - 4px)`;$("statChartTruck").setAttribute("aria-label",`Camión en ${percent} por ciento`);$("statChartTotal").textContent=`Total filtrado: ${total.toLocaleString("es-CO")} transportes`;$("statChartVehicleHeading").textContent=`Transportes en categoría: ${count}`;$("statChartVehicles").innerHTML=count?item.rows.map(r=>`<div class="statVehicleRow"><span class="statVehicleName">${esc(String(r.transporte??"Sin nombre").trim())} - ${esc(String(r.placa??"Sin placa").trim())}</span></div>`).join(""):'<p class="statVehicleEmpty">No hay transportes en esta categoría.</p>';$("statChartModal").classList.remove("hidden");$("statChartClose").focus()}function closeStatChart(){$("statChartModal")?.classList.add("hidden")}
function normalizePhysicalPlate(value){let v=String(value||"").toLocaleUpperCase("es").replace(/\s+/g,"");if(v&&!v.startsWith("CO"))v="CO"+v;return v.slice(0,8)}
function vehicleClass(v){return v==="Bodega"?"vehicle-bodega":v==="Sobredimensionado"?"vehicle-sobredimensionado":v==="OK para utilizar"?"vehicle-ok":""}
async function saveVehicleField(id,field,value){if(!canModule("cargo","edit")){toast("Sin permiso de modificación en el módulo.");return}if(field==="cambio_placa")value=normalizePhysicalPlate(value);if(field==="cambio_placa"&&String(value).length>8){toast("Placa Fisica: máximo 8 caracteres.");await refresh();return}const{error}=await sb.rpc("update_transport_vehicle_field_secure",{p_transport_id:id,p_field:field,p_value:value,p_pin:loginPin});if(error){toast("No se pudo guardar: "+error.message);await refresh();return}const r=records.find(x=>Number(x.id)===Number(id));if(r)r[field]=value;toast(field==="estado_vehiculo"?"Disponibilidad De Vehiculos guardada.":"Placa Fisica guardada.")}
setInterval(()=>{if(loginPin&&sb&&!document.hidden)ctFetchActions()},15000);document.addEventListener("visibilitychange",()=>{if(!document.hidden&&loginPin&&sb)void ctFetchActions()});
const CT_ACTION_BUILD='5.50-20261022';
const ctActionDefaults={armado:false,contado:false,cargado:false};let ctActionFlags={...ctActionDefaults},ctActionCanManage=false,ctActionsLoaded=false;
function ctNextAction(estado){const order=['Armado','Contado','Cargado'];const index=estado==='Pendiente'?-1:order.indexOf(estado);if(index===-1&&estado!=='Pendiente')return null;return order.slice(index+1).find(x=>ctActionFlags[x.toLowerCase()])||null}
function ctActionAllowed(next){return permissions.is_admin||(next==='Armado'?permissions.can_arm:next==='Contado'?permissions.can_count:permissions.can_load)}
function ctActionLabel(next){return next==='Armado'?'Armar':next==='Contado'?'Contar':next==='Cargado'?'Cargar':'Sin acción'}
function ctUpdateActionUI(){const marker=$('ctActionBuildMarker');if(marker)marker.textContent='DIAGNÓSTICO '+CT_ACTION_BUILD+' · Supabase: '+(ctActionsLoaded?('Armado '+ctActionFlags.armado+' / Contado '+ctActionFlags.contado+' / Cargado '+ctActionFlags.cargado):'Sin conexión de configuración');const b=$('ctActionConfigure');if(b)b.classList.toggle('hidden',!ctActionCanManage);const summary=$('ctActionSummary');if(summary)summary.textContent='Acciones habilitadas: '+Object.entries(ctActionFlags).filter(([,v])=>v).map(([k])=>k[0].toUpperCase()+k.slice(1)).join(' · ')||'Sin acciones habilitadas';render()}
function ctBoolean(value){return value===true||value===1||String(value).trim().toLowerCase()==='true'||String(value).trim()==='1'}
async function ctFetchActions(){
 if(!sb||!loginPin)return false;
 try{
  const {data,error}=await sb.rpc('ct_cargo_action_config_get',{p_pin:loginPin});
  if(error)throw error;
  // PostgREST puede devolver una fila como objeto o un conjunto de filas como arreglo.
  const row=Array.isArray(data)?data[0]:data;
  if(!row||typeof row!=='object'||!['armado','contado','cargado'].every(key=>Object.prototype.hasOwnProperty.call(row,key))){throw new Error('La consulta de configuración no devolvió armado, contado y cargado.');}
  ctActionFlags={armado:ctBoolean(row.armado),contado:ctBoolean(row.contado),cargado:ctBoolean(row.cargado)};
  ctActionCanManage=ctBoolean(row.can_manage);
  ctActionsLoaded=true;
  ctUpdateActionUI();
  return true;
 }catch(e){
  console.warn('Acciones de Cargue:',e.message);const marker=$('ctActionBuildMarker');if(marker)marker.textContent='DIAGNÓSTICO '+CT_ACTION_BUILD+' · ERROR SUPABASE: '+String(e.message||e);
  ctActionFlags={...ctActionDefaults};ctActionCanManage=false;ctActionsLoaded=false;
  ctUpdateActionUI();const markerError=$('ctActionBuildMarker');if(markerError)markerError.textContent='DIAGNÓSTICO '+CT_ACTION_BUILD+' · ERROR SUPABASE: '+String(e.message||e);
  const summary=$('ctActionSummary');if(summary)summary.textContent='No se pudo leer la configuración de Supabase. Acciones temporalmente deshabilitadas.';
  return false;
 }
}
async function ctSaveActions(){if(!ctActionCanManage)return;const btn=$('ctActionSave');btn.disabled=true;try{const args={p_pin:loginPin,p_armado:$('ctActionArmado').checked,p_contado:$('ctActionContado').checked,p_cargado:$('ctActionCargado').checked};const {error}=await sb.rpc('ct_cargo_action_config_save',args);if(error)throw error;await ctFetchActions();$('ctActionSettings').classList.add('hidden');toast('Acciones actualizadas para todos los usuarios')}catch(e){$('ctActionSettingsStatus').textContent=e.message}finally{btn.disabled=false}}
function ctSetupActionSettings(){const b=$('ctActionConfigure');if(!b)return;b.onclick=()=>{if(!ctActionCanManage)return;for(const key of ['Armado','Contado','Cargado'])$('ctAction'+key).checked=ctActionFlags[key.toLowerCase()];$('ctActionSettings').classList.toggle('hidden')};$('ctActionSave').onclick=ctSaveActions}
async function advance(id){if(!canModule("cargo","edit")){toast("Sin permiso de modificación en el módulo.");return}const isOnline=Number(id)<0;
const r=isOnline?onlineTrips.find(x=>Number(-1000000-x.id)===Number(id)):records.find(x=>Number(x.id)===Number(id));if(!r)return;
if(String(selectedTransportId)!==String(id)){toast("Primero selecciona el transporte.");return}
// Volver a consultar antes de cada cambio: nunca usar una configuración obsoleta.
if(!await ctFetchActions()){toast('No fue posible verificar las acciones habilitadas.');return}
const next=ctNextAction(r.estado);
if(!next)return;
const allowed=ctActionAllowed(next);
if(!allowed){toast("No tienes permiso para avanzar este estado.");return}
if(!confirm(`¿Confirmas el cambio de estado?\n\nPlaca: ${r.placa}\nEstado: ${r.estado} → ${next}`))return;
const rpc="ct_cargo_action_advance";
const args={p_id:isOnline?r.id:id,p_pin:loginPin,p_online:isOnline,p_target:next};
const{error}=await sb.rpc(rpc,args);if(error){toast("No se pudo actualizar: "+error.message);return}
selectedTransportId=null;toast(`Estado actualizado: ${r.estado} → ${next}`);await refresh();if(next==="Cargado")await refreshCheckin()}
async function loadHistory(){const{data,error}=await sb.rpc("recent_history",{p_limit:30});if(error)return;$("history").innerHTML=(data||[]).map(h=>`<div class="historyItem"><b>${esc(h.usuario)}</b> · ${esc(h.placa)} · Viaje ${esc(h.viaje)}<br>${h.estado_anterior} → <b>${h.estado_nuevo}</b> · ${new Date(h.created_at).toLocaleString("es-CO")}</div>`).join("")||"Sin movimientos."}
let ctAdminUsers=[];
const ctGeneralFlags=[['can_upload','Cargar y descargar archivos'],['can_create_users','Crear usuarios'],['can_manage_checkin','Modificar pendientes de Checkin'],['can_delete_all_transports','Eliminar todos los transportes'],['can_arm','Cambiar a Armado'],['can_count','Cambiar a Contado'],['can_load','Cambiar a Cargado'],['can_edit_availability','Editar disponibilidad'],['can_edit_physical_plate','Editar placa física']];
const ctFinanceTabs={base:'Pendientes por Liquidar',novedades:'Novedades',restricciones:'Restricciones',documentos:'Documentos Extraviados',reporte:'Reporte Restricciones'};
const ctModuleLabels={cargo:'Cargue y Despacho',checkin:'CHECKIN',rutero:'Brújula y Rutero'};
let ctEditorUser=null;
function ctRenderAdminUsers(){
 const term=($('ctUsersSearch')?.value||'').trim().toLowerCase(),filter=$('ctUsersFilter')?.value||'all';
 const result=ctAdminUsers.filter(u=>(filter==='all'||!!u.active===(filter==='active'))&&[u.name,u.cedula,u.role].some(v=>String(v||'').toLowerCase().includes(term)));
 $('ctUsersTotal').textContent=result.length+' de '+ctAdminUsers.length+' usuarios';
 const root=$('users');root.classList.add('ctUserList');root.replaceChildren();
 result.forEach(u=>{
  const card=document.createElement('div');card.className='ctUserCard'+(u.active?'':' inactive');
  const info=document.createElement('div');info.className='ctUserIdentity';
  const name=document.createElement('strong');name.textContent=u.name;
  const description=document.createElement('small');description.textContent='C.C. '+(u.cedula||'Sin cédula')+' · '+(u.role||'Operativo')+' · '+(u.active?'Activo':'Sin acceso');info.append(name,description);
  const actions=document.createElement('div');actions.className='ctUserCardActions';
  const edit=document.createElement('button');edit.type='button';edit.textContent='⚙️ Administrar permisos';edit.disabled=!permissions.is_admin;edit.onclick=()=>ctOpenPermissionEditor(u.id);actions.append(edit);
  if(permissions.is_admin&&!u.is_admin){
   const toggle=document.createElement('button');toggle.type='button';toggle.className=u.active?'danger':'secondary';toggle.textContent=u.active?'Quitar acceso':'Restaurar acceso';toggle.onclick=()=>ctToggleUserAccess(u);actions.append(toggle);
   const returns=document.createElement('button');returns.type='button';returns.className='secondary ctReturnsUserToggle';returns.dataset.user=u.id;returns.dataset.name=u.name;returns.textContent='📦 Devoluciones';actions.append(returns);
   const finance=document.createElement('button');finance.type='button';finance.className='secondary ctFinanceUserToggle';finance.dataset.user=u.id;finance.dataset.enabled='false';finance.textContent='💵 Financiera: Cargando…';actions.append(finance);
  }
  card.append(info,actions);root.append(card);
 });
 if(permissions.is_admin){void ctPaintReturnsAdminPermissions();void ctPaintFinanceAdminPermissions();}
}
async function loadUsers(){const{data,error}=await sb.rpc('admin_users_with_all_permissions',{p_admin_pin:loginPin});if(error){toast(error.message);return}ctAdminUsers=Array.isArray(data)?data:[];ctRenderAdminUsers();}
async function ctOpenPermissionEditor(id){
 if(!permissions.is_admin)return;const u=ctAdminUsers.find(v=>v.id===id);if(!u)return;
 ctEditorUser=id;$('ctEditorTitle').textContent='Permisos de '+u.name;$('ctEditorSummary').textContent='Cédula: '+(u.cedula||'—')+' · '+(u.active?'Cuenta activa':'Cuenta sin acceso');$('ctEditorRole').value=u.role||'Operativo';
 const gen=$('ctEditorGeneral');gen.replaceChildren();ctGeneralFlags.forEach(([key,label])=>{const l=document.createElement('label'),ch=document.createElement('input');ch.type='checkbox';ch.dataset.flag=key;ch.checked=!!u[key];l.append(ch,document.createTextNode(label));gen.append(l)});
 const {data,error}=await sb.rpc('admin_get_module_permissions',{p_admin_pin:loginPin,p_user_id:id});if(error){toast('No se pudieron consultar módulos: '+error.message);return}
 const rights=data?.[0]||DEFAULT_MODULE_RIGHTS,modules=$('ctEditorModules');modules.replaceChildren();
 Object.entries(ctModuleLabels).forEach(([key,label])=>{
  const div=document.createElement('div');div.className='ctEditorModule';const title=document.createElement('strong');title.textContent=label;div.append(title);
  for(const mode of ['view','edit']){const l=document.createElement('label'),ch=document.createElement('input');ch.type='checkbox';ch.dataset.module=key;ch.dataset.mode=mode;ch.checked=!!rights[key+'_'+mode];l.append(ch,document.createTextNode(mode==='view'?'Ver pestaña':'Modificar'));div.append(l)}modules.append(div)
 });
 modules.querySelectorAll('input[data-mode="edit"]').forEach(ch=>ch.addEventListener('change',()=>{if(ch.checked)modules.querySelector('input[data-module="'+ch.dataset.module+'"][data-mode="view"]').checked=true}));
 modules.querySelectorAll('input[data-mode="view"]').forEach(ch=>ch.addEventListener('change',()=>{if(!ch.checked)modules.querySelector('input[data-module="'+ch.dataset.module+'"][data-mode="edit"]').checked=false}));
 const vr=$('ctEditorValidaciones');vr.replaceChildren();
  const actionResult=await sb.rpc('ct_cargo_action_admin_user_rights',{p_admin_pin:loginPin,p_user_id:id});
 if(actionResult.error){toast('Permiso configuración de acciones: '+actionResult.error.message);return;}
 const actionLabel=document.createElement('label'),actionCheck=document.createElement('input');actionCheck.type='checkbox';actionCheck.id='ctManageCargoActions';actionCheck.checked=actionResult.data===true;actionLabel.append(actionCheck,document.createTextNode(' Puede configurar acciones globales de Cargue y Despacho'));$('ctEditorGeneral').append(actionLabel);
 const vresult=await sb.rpc('ct_validaciones_admin_user_rights',{p_admin_pin:loginPin,p_user_id:id});
 if(vresult.error){toast('No se pudo consultar Validaciones Operación: '+vresult.error.message);return;}
 const vwrap=document.createElement('div');vwrap.className='ctEditorModule';
 const vtitle=document.createElement('strong');vtitle.textContent='Validaciones Operación';
 const vlabel=document.createElement('label'),vcheck=document.createElement('input');vcheck.type='checkbox';vcheck.id='ctValidacionesView';vcheck.checked=vresult.data===true;
 vlabel.append(vcheck,document.createTextNode(' Permitir ingresar'));vwrap.append(vtitle,vlabel);vr.append(vwrap);
 const financeRoot=$('ctEditorFinance');financeRoot.replaceChildren();
 const fresult=await sb.rpc('ct_finance_admin_user_rights',{p_admin_pin:loginPin,p_user_id:id});
 if(fresult.error){toast('No se pudieron consultar permisos financieros: '+fresult.error.message);return;}
 const fdata=fresult.data||{};
 Object.entries(ctFinanceTabs).forEach(([tab,label])=>{
  const div=document.createElement('div');div.className='ctEditorModule';const title=document.createElement('strong');title.textContent=label;div.append(title);
  for(const mode of ['view','edit']){const l=document.createElement('label'),ch=document.createElement('input');ch.type='checkbox';ch.dataset.financeTab=tab;ch.dataset.financeMode=mode;ch.checked=!!fdata[tab+'_'+mode];l.append(ch,document.createTextNode(mode==='view'?' Ver':' Modificar'));div.append(l)}financeRoot.append(div);
 });
 financeRoot.querySelectorAll('[data-finance-mode="edit"]').forEach(ch=>ch.addEventListener('change',()=>{if(ch.checked)financeRoot.querySelector('[data-finance-tab="'+ch.dataset.financeTab+'"][data-finance-mode="view"]').checked=true}));
 financeRoot.querySelectorAll('[data-finance-mode="view"]').forEach(ch=>ch.addEventListener('change',()=>{if(!ch.checked)financeRoot.querySelector('[data-finance-tab="'+ch.dataset.financeTab+'"][data-finance-mode="edit"]').checked=false}));
 $('ctEditorDeactivate').textContent=u.active?'Quitar acceso al usuario':'Restaurar acceso';$('ctEditorDeactivate').className=u.active?'danger':'secondary';$('ctEditorStatus').textContent='';$('ctUserPermissionDrawer').classList.remove('hidden');$('ctUserPermissionDrawer').scrollIntoView({behavior:'smooth',block:'nearest'});
}
async function ctSaveUserRights(){
 if(!permissions.is_admin||!ctEditorUser)return;
 const u=ctAdminUsers.find(x=>x.id===ctEditorUser);if(!u)return;
 const flag=k=>!!$('ctEditorGeneral').querySelector('[data-flag="'+k+'"]').checked;
 const role=$('ctEditorRole').value;
 if(role==='Administrador'&&!confirm('Convertir a este usuario en administrador le concederá permisos elevados. ¿Continuar?'))return;
 const args={p_admin_pin:loginPin,p_user_id:u.id,p_role:role,p_can_upload:flag('can_upload'),p_can_create_users:flag('can_create_users'),p_can_delete_all_transports:flag('can_delete_all_transports'),p_can_arm:flag('can_arm'),p_can_count:flag('can_count'),p_can_load:flag('can_load'),p_can_edit_availability:flag('can_edit_availability'),p_can_edit_physical_plate:flag('can_edit_physical_plate'),p_can_manage_checkin:flag('can_manage_checkin')};
 const rights={};for(const mod of Object.keys(ctModuleLabels))for(const mode of ['view','edit'])rights[mod+'_'+mode]=$('ctEditorModules').querySelector('[data-module="'+mod+'"][data-mode="'+mode+'"]').checked;
 const btn=$('ctEditorSave');btn.disabled=true;$('ctEditorStatus').textContent='Guardando permisos…';
 try{
  const general=await sb.rpc('admin_set_user_all_permissions_secure',args);if(general.error)throw general.error;
  const modules=await sb.rpc('admin_set_module_permissions',{p_admin_pin:loginPin,p_user_id:u.id,p_rights:rights});if(modules.error)throw Error('Permisos generales guardados, pero falló el acceso a módulos: '+modules.error.message);
  const financeRights={};for(const tab of Object.keys(ctFinanceTabs))for(const mode of ['view','edit'])financeRights[tab+'_'+mode]=$('ctEditorFinance').querySelector('[data-finance-tab="'+tab+'"][data-finance-mode="'+mode+'"]').checked;
  const fresult=await sb.rpc('ct_finance_admin_set_user_rights',{p_admin_pin:loginPin,p_user_id:u.id,p_rights:financeRights});if(fresult.error)throw Error('Permisos generales guardados, pero falló Liquidación Financiera: '+fresult.error.message);
  const aresult=await sb.rpc('ct_cargo_action_admin_set_user_rights',{p_admin_pin:loginPin,p_user_id:u.id,p_enabled:$('ctManageCargoActions').checked});if(aresult.error)throw Error('No se pudo guardar permiso para configurar acciones: '+aresult.error.message);
  const vresult=await sb.rpc('ct_validaciones_admin_set_access',{p_admin_pin:loginPin,p_user_id:u.id,p_can_view:$('ctValidacionesView').checked});if(vresult.error)throw Error('No se pudo guardar Validaciones Operación: '+vresult.error.message);
  $('ctEditorStatus').textContent='Permisos guardados correctamente';toast('Permisos actualizados');await loadUsers();
 }catch(e){$('ctEditorStatus').textContent='No se pudo completar: '+e.message;toast(e.message)}finally{btn.disabled=false}
}
async function ctToggleUserAccess(u){if(!permissions.is_admin||u.is_admin)return;
 if(!confirm(u.active?'¿Quitar el acceso a '+u.name+'? Su historial se conservará y podrás restaurar la cuenta.':'¿Restaurar el acceso de '+u.name+'?'))return;
 const {error}=await sb.rpc('admin_set_user_active_secure',{p_admin_pin:loginPin,p_user_id:u.id,p_active:!u.active});if(error){toast(error.message);return}await loadUsers();if(ctEditorUser===u.id)$('ctUserPermissionDrawer').classList.add('hidden');toast(u.active?'Acceso retirado':'Acceso restaurado');
}
async function ctDeleteUserForever(){
 if(!permissions.is_admin||!ctEditorUser)return;
 const u=ctAdminUsers.find(x=>x.id===ctEditorUser);if(!u||u.is_admin){toast('No es posible eliminar administradores desde este control.');return;}
 if(!confirm('¿Eliminar DEFINITIVAMENTE la cuenta de '+u.name+'? No podrás restaurarla. Esta operación no elimina transportes ni registros históricos.'))return;
 const answer=prompt('Para confirmar, escribe exactamente la cédula del usuario: '+u.cedula);
 if(answer===null||answer.trim()!==String(u.cedula||'').trim()){toast('Cédula incorrecta. No se eliminó la cuenta.');return;}
 const b=$('ctEditorDeleteForever');b.disabled=true;
 try{const {error}=await sb.rpc('ct_admin_delete_user_permanent',{p_admin_pin:loginPin,p_user_id:u.id,p_confirm_cedula:answer.trim()});if(error)throw error;
 $('ctUserPermissionDrawer').classList.add('hidden');ctEditorUser=null;await loadUsers();toast('Usuario eliminado permanentemente');
 }catch(e){toast('No se pudo eliminar: '+e.message)}finally{b.disabled=false}
}
const ctSetupUserPanel=()=>{const create=$('ctToggleCreateUser');if(create)create.onclick=()=>{const pane=$('ctCreateUserForm');pane.classList.toggle('ctCreateUserCollapsed');create.textContent=pane.classList.contains('ctCreateUserCollapsed')?'＋ Nuevo usuario':'− Ocultar creación'};
 $('ctUsersSearch')?.addEventListener('input',ctRenderAdminUsers);$('ctUsersFilter')?.addEventListener('change',ctRenderAdminUsers);
 $('ctEditorClose')?.addEventListener('click',()=>$('ctUserPermissionDrawer').classList.add('hidden'));
 $('ctEditorSave')?.addEventListener('click',ctSaveUserRights);
 $('ctEditorDeactivate')?.addEventListener('click',()=>{const u=ctAdminUsers.find(x=>x.id===ctEditorUser);if(u)ctToggleUserAccess(u)});
 $('ctEditorDeleteForever')?.addEventListener('click',ctDeleteUserForever);
};
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',ctSetupUserPanel);else ctSetupUserPanel();
async function addUser(){if(!permissions.can_create_users&&!permissions.is_admin){toast("No tienes permiso para crear usuarios.");return}const name=$("newName").value.trim(),cedula=$("newCedula").value.trim(),pin=$("newPin").value.trim(),role=$("newRole").value;if(!name||!cedula||!pin){toast("Nombre, cédula y PIN son obligatorios");return}if(role==="Administrador"&&!permissions.is_admin){toast("Solo un administrador puede crear administradores.");return}const{data:newUserId,error}=await sb.rpc("admin_create_user_full_secure",{p_name:name,p_cedula:cedula,p_pin:pin,p_creator_pin:loginPin,p_role:role,p_can_upload:permissions.is_admin&&$("newCanUpload").checked,p_can_create_users:permissions.is_admin&&$("newCanCreateUsers").checked,p_can_delete_all_transports:permissions.is_admin&&$("newCanDeleteAll").checked,p_can_arm:permissions.is_admin&&$("newCanArm").checked,p_can_count:permissions.is_admin&&$("newCanCount").checked,p_can_load:permissions.is_admin&&$("newCanLoad").checked,p_can_edit_availability:permissions.is_admin&&$("newCanAvailability").checked,p_can_edit_physical_plate:permissions.is_admin&&$("newCanPhysicalPlate").checked,p_can_manage_checkin:(permissions.is_admin||permissions.can_create_users)&&$("newCanManageCheckin").checked});if(error){toast(error.message);return}$("newName").value="";$("newCedula").value="";$("newPin").value="";$("newRole").value="Operativo";["newCanUpload","newCanCreateUsers","newCanDeleteAll","newCanArm","newCanCount","newCanLoad","newCanAvailability","newCanPhysicalPlate","newCanManageCheckin"].forEach(id=>$(id).checked=false);if(permissions.is_admin&&newUserId){const rights={};for(const mod of ["Cargo","Checkin","Rutero"])for(const mode of ["View","Edit"])rights[mod.toLowerCase()+"_"+mode.toLowerCase()]=$("new"+mod+mode).checked;const res=await sb.rpc("admin_set_module_permissions",{p_admin_pin:loginPin,p_user_id:newUserId,p_rights:rights});if(res.error){toast("Usuario creado, pero faltan permisos de módulos: "+res.error.message);return}}if(permissions.is_admin)await loadUsers();toast("Usuario creado correctamente")}
async function setUserPermissions(id){const{data,error}=await sb.rpc("admin_users_with_all_permissions",{p_admin_pin:loginPin});if(error){toast(error.message);return}const u=(data||[]).find(x=>x.id===id);if(!u)return;const role=prompt("Rol: Administrador, Operativo o Supervisor",u.role||"Operativo");if(role===null)return;if(!["Administrador","Operativo","Supervisor"].includes(role.trim())){toast("Rol no válido");return}const ask=(label,value)=>confirm(label+"\nAceptar = Sí / Cancelar = No");const canUpload=ask("¿Puede cargar/descargar archivos?",u.can_upload),canCreate=ask("¿Puede crear usuarios?",u.can_create_users),canDelete=ask("¿Puede eliminar todos los transportes?",u.can_delete_all_transports),canArm=ask("¿Puede cambiar a estado Armado?",u.can_arm),canCount=ask("¿Puede cambiar a estado Contado?",u.can_count),canLoad=ask("¿Puede cambiar a estado Cargado?",u.can_load),canAvail=ask("¿Puede editar Disponibilidad de Vehículos?",u.can_edit_availability),canPlate=ask("¿Puede editar Placa Física?",u.can_edit_physical_plate);const{error:saveError}=await sb.rpc("admin_set_user_all_permissions_secure",{p_admin_pin:loginPin,p_user_id:id,p_role:role.trim(),p_can_upload:canUpload,p_can_create_users:canCreate,p_can_delete_all_transports:canDelete,p_can_arm:canArm,p_can_count:canCount,p_can_load:canLoad,p_can_edit_availability:canAvail,p_can_edit_physical_plate:canPlate,p_can_manage_checkin:ask("¿Puede modificar datos en Pendientes Por Checkin?",u.can_manage_checkin)});if(saveError){toast(saveError.message);return}await loadUsers();toast("Permisos actualizados")}
async function toggleUser(id,active){const{error}=await sb.rpc("admin_set_user_active_secure",{p_admin_pin:loginPin,p_user_id:id,p_active:!active});if(error){toast(error.message);return}loadUsers()}
function norm(s){return String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s/g,"")}
async function prepareImport(){if(!canModule("cargo","edit")){toast("Sin permiso de modificación en el módulo.");return}if(!permissions.can_upload&&!permissions.is_admin){toast("No tienes permiso para cargar archivos.");return}const f=$("excelFile").files[0];if(!f){toast("Selecciona un archivo");return}const buf=await f.arrayBuffer(),wb=XLSX.read(buf,{type:"array"});const required=["Viajes HYM y URBAN","Viajes KA","Viajes Tradicional"];const optional=["Viajes 2KA","Viajes 2","Custodias"];const expected=[...required];const sheetFor=label=>wb.SheetNames.find(n=>norm(n)===norm(label));const missing=required.filter(n=>!sheetFor(n));if(missing.length){toast("Faltan pestañas en el Excel: "+missing.join(", "));return}for(const category of optional){if(sheetFor(category))expected.push(category)}const tapaSheet=sheetFor("Tapa Códigos");if(!tapaSheet){toast("Falta la pestaña Tapa Códigos en el Excel.");return}const tapaData=XLSX.utils.sheet_to_json(wb.Sheets[tapaSheet],{defval:""});if(tapaData.length){const tk=Object.keys(tapaData[0]),findT=n=>tk.find(k=>norm(k)===norm(n));const tTrans=findT("transporte"),tVeh=findT("vehiculo"),tViaje=findT("viaje"),tMat=findT("material"),tDen=findT("denominacion"),tCant=findT("cantidad"),tFam=findT("familia"),tDesc=findT("descripcion");if(!tTrans||!tVeh||!tViaje||!tMat||!tDen||!tCant||!tFam||!tDesc){toast("Tapa Códigos debe tener FAMILIA, TRANSPORTE, VEHICULO, VIAJE, MATERIAL, DENOMINACION, CANTIDAD y DESCRIPCION.");return}pendingTapaRows=tapaData.map(x=>({familia:String(x[tFam]??"").trim(),descripcion:String(x[tDesc]??"").trim(),transporte:String(x[tTrans]??"").trim(),vehiculo:String(x[tVeh]??"").trim(),viaje:String(x[tViaje]??"").trim(),material:String(x[tMat]??"").trim(),denominacion:String(x[tDen]??"").trim(),cantidad:Number(String(x[tCant]??"0").replace(/,/g,""))||0})).filter(x=>x.familia||x.descripcion||x.transporte||x.vehiculo||x.viaje||x.material||x.denominacion)}else{pendingTapaRows=[]}let rows=[];for(const category of expected){const data=XLSX.utils.sheet_to_json(wb.Sheets[sheetFor(category)],{defval:""});if(!data.length)continue;const keys=Object.keys(data[0]),find=n=>keys.find(k=>norm(k)===n),kt=find("transporte"),kv=find("viaje"),kp=find("placa"),kn=find("novedad");if(!kt||!kv||!kp||!kn){toast("La pestaña "+category+" debe contener Transporte, Viaje, Placa y Novedad");return}rows.push(...data.map(x=>({transporte:String(x[kt]??"").trim(),viaje:String(x[kv]??"").trim(),placa:String(x[kp]??"").trim(),novedad:String(x[kn]??"").trim(),categoria:category})).filter(x=>x.transporte||x.viaje||x.placa))}if(!rows.length){toast("Las pestañas de viajes están vacías");return}const seen=new Set();let duplicates=0;const unique=rows.filter(x=>{const k=[x.categoria,x.transporte,x.viaje,x.placa].join("|").toLowerCase();if(seen.has(k)){duplicates++;return false}seen.add(k);return true});const existingKeys=new Set(records.map(r=>[r.categoria||"Sin categoría",r.transporte,r.viaje,r.placa].map(v=>String(v??"").trim()).join("|").toLowerCase()));const existing=unique.filter(x=>existingKeys.has([x.categoria,x.transporte,x.viaje,x.placa].map(v=>String(v??"").trim()).join("|").toLowerCase())).length;pendingImportRows=unique;const importMsg=$("importResultMessage");if(importMsg){importMsg.textContent="";importMsg.classList.add("hidden");}$("previewText").innerHTML=`<b>${rows.length}</b> filas leídas entre las pestañas de viajes y <b>${pendingTapaRows.length}</b> filas de Tapa Códigos.<br>✅ <b>${unique.length-existing}</b> nuevas para agregar.<br>🔄 <b>${existing}</b> ya existen: se actualizará su Novedad y se conservarán estado, placa física y disponibilidad.<br>♻️ <b>${duplicates}</b> duplicadas dentro del Excel no se agregarán.`;$("previewModal").classList.remove("hidden")}
async function confirmImport(){
  const btn=$('confirmImportBtn');
  const msg=$('importResultMessage');
  const display=(message,isError=false)=>{
    if(msg){msg.textContent=message;msg.style.color=isError?'#9f2733':'#124c74';msg.classList.remove('hidden');}
    if(isError){console.error('[Importación de viajes]',message);toast(message);}
  };
  if(!btn){console.error('[Importación] Falta confirmImportBtn');return;}
  if(!canModule('cargo','edit')||(!permissions.can_upload&&!permissions.is_admin)){display('Sin permisos para importar. Revisa con el administrador.',true);return;}
  if(!Array.isArray(pendingImportRows)||!pendingImportRows.length){display('No hay viajes pendientes por importar. Selecciona de nuevo el Excel.',true);return;}
  if(btn.dataset.saving==='1')return;
  btn.dataset.saving='1';btn.disabled=true;
  const previous=btn.textContent;
  let timer;
  let phase='Enviando viajes a Supabase';
  const started=Date.now();
  let saved=false;
  try{
    if(!sb||typeof sb.rpc!=='function')throw Error('La conexión Supabase no está inicializada. Vuelve a iniciar sesión.');
    const rows=pendingImportRows.map(x=>({...x}));
    btn.textContent='Guardando viajes…';
    display(`Paso 1/3: ${phase} (${rows.length} registros). No cierres esta ventana ni repitas la carga.`);
    timer=setInterval(()=>{const sec=Math.floor((Date.now()-started)/1000);display(`Paso ${saved?'3/3':'1/3'}: ${phase}. Transcurridos ${sec} s.${sec>=15&&!saved?' La respuesta tarda más de lo esperado; no confirmes nuevamente hasta comprobar la base.':''}`);},2000);
    console.info('[Importación] Enviando',rows.length,'registros al procedimiento replace_transports_secure');
    const {data:result,error}=await sb.rpc('replace_transports_secure',{p_rows:rows,p_pin:loginPin});
    if(error){display('Supabase rechazó la carga: '+error.message,true);return;}
    saved=true;
    console.info('[Importación] Confirmación Supabase',result);if(result?.categories_updated)console.info("Categorias corregidas:",result.categories_updated);
    const added=Number(result?.added??0),existing=Number(result?.existing??0);
    phase='Viajes confirmados por Supabase';
    display(`Paso 2/3: Supabase confirmó la carga (${added} nuevos; ${existing} ya existentes).`);
    if(Array.isArray(pendingTapaRows)&&pendingTapaRows.length){
      phase='Guardando Tapa Códigos';btn.textContent='Guardando Tapa Códigos…';
      const tapa=await sb.rpc('replace_tapa_codigos_ka_secure',{p_rows:pendingTapaRows,p_pin:loginPin});
      if(tapa.error){display(`Viajes confirmados (${added} nuevos), pero Tapa Códigos falló: ${tapa.error.message}. No repitas la carga completa.`,true);return;}
      await refreshTapa();
    }
    phase='Actualizando tabla de Cargue y Despacho';
    display(`Paso 3/3: ${phase}.`);
    try{await refresh();if(records.some(r=>r.categoria==='Sin categoría')){display('Viajes guardados, pero hay categorías sin asignar. Verifica el SQL v5.29 y vuelve a importar el Excel.',true);return;}}catch(refreshError){display(`Supabase confirmó la carga (${added} nuevos), pero falló la actualización de pantalla: ${refreshError.message}. Vuelve a abrir Cargue y Despacho; NO repitas el Excel.`,true);return;}
    display(`Carga confirmada: ${added} nuevos, ${existing} existentes.`);
    pendingImportRows=null;pendingTapaRows=null;
    $('previewModal').classList.add('hidden');$('excelFile').value='';
    toast(`Carga confirmada: ${added} nuevos, ${existing} existentes.`);
  }catch(e){
    display(`${saved?'Los viajes se guardaron, pero hubo un error posterior':'No se pudo confirmar la importación'}: ${e?.message||String(e)}. ${saved?'No repitas la carga.':'Consulta primero si los transportes ya aparecen en Cargue y Despacho.'}`,true);
  }finally{
    clearInterval(timer);btn.disabled=false;btn.dataset.saving='0';btn.textContent=previous;
  }
}
function cancelImport(){pendingImportRows=null;$("previewModal").classList.add("hidden")}
async function downloadExcel(){
if(!permissions.can_upload&&!permissions.is_admin){toast("No tienes permiso para descargar archivos.");return}
const btn=$("downloadBtn"),oldText=btn.textContent;btn.disabled=true;btn.textContent="Preparando consolidado…";
try{
 const {data,error}=await sb.rpc("list_transports_with_category");if(error)throw Error(error.message);
 const fields=await sb.rpc("list_transport_vehicle_fields");if(fields.error)throw Error(fields.error.message);
 const vehicleMap=new Map((fields.data||[]).map(x=>[String(x.id),x]));
 const normTransport=x=>String(x??'').trim().replace(/^S(?=\d)/i,'').replace(/\.0$/,'').replace(/^0+(?=\d)/,'');
 const normHeader=x=>String(x??'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');
 // Cliente = Solic. (identificador único), NO el número de filas del Rutero.
 const customerCounts=new Map();let ruteroAvailable=false;
 const ingestRutero=(headers,rows)=>{
  const cols=headers.map(normHeader),ti=cols.indexOf('transporte');
  const ci=cols.findIndex(x=>['solic','solicitante','destinatario','destinat','idcliente','codigocliente'].includes(x));
  if(ti<0||ci<0)return false;
  const sets=new Map();
  for(const row of rows){const t=normTransport(row[ti]),id=String(row[ci]??'').trim();if(!t||!id)continue;
   if(!sets.has(t))sets.set(t,new Set());sets.get(t).add(id);
  }
  sets.forEach((clients,t)=>customerCounts.set(t,clients.size));return true;
 };
 const current=window.getRuteroSyncData?.();
 if(current?.loaded&&current.rows?.length)ruteroAvailable=ingestRutero(current.headers||[],current.rows);
 if(!ruteroAvailable&&window.CTShared){
  btn.textContent='Consultando clientes del Rutero…';
  try{
   const saved=await window.CTShared.load('rutero',(n,total)=>{btn.textContent=`Leyendo Rutero ${n}/${total}…`;});
   if(saved?.items?.length)ruteroAvailable=ingestRutero(saved.metadata?.headers||[],saved.items);
  }catch(err){console.warn('[Exportar clientes]',err);throw Error('No se pudo consultar el Rutero para contar clientes: '+err.message);}
 }
 if(!ruteroAvailable)toast('No hay Rutero con columna Solic. disponible. Los clientes saldrán como Sin dato.');
 const categories=['Viajes HYM y URBAN','Viajes KA','Viajes Tradicional','Viajes 2KA','Viajes 2','Custodias'];
 const base=(data||[]).map(r=>{const v=vehicleMap.get(String(r.id))||{};return {
  Categoría:r.categoria||'Sin categoría',Transporte:r.transporte||'',Viaje:r.viaje||'',Placa:r.placa||'',
  'Placa física':v.cambio_placa||'',Estado:r.estado||'Pendiente',
  Clientes:ruteroAvailable?(customerCounts.get(normTransport(r.transporte))??0):'Sin dato'
 };});
 const wb=XLSX.utils.book_new();
 const dateText=new Date().toLocaleDateString('es-CO',{year:'numeric',month:'long',day:'numeric'});
 const navy='123047',blue='1769AA',pale='EAF1F7',white='FFFFFF',gray='536477';
 const headers=['Categoría','Transporte','Viaje','Placa','Placa física','Estado','Clientes'];
 function createSheet(title,rows,label){
  const values=[['OPERACIÓN CD GIRARDOTA'],[title],['Fecha',dateText],['Categorías',label],['Total transportes',rows.length],[],headers,...rows.map(r=>headers.map(h=>r[h]))];
  const ws=XLSX.utils.aoa_to_sheet(values);ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:6}},{s:{r:1,c:0},e:{r:1,c:6}}];
  ws['!cols']=[{wch:25},{wch:19},{wch:22},{wch:18},{wch:20},{wch:16},{wch:13}];
  ws['!freeze']={xSplit:0,ySplit:7};
  ws['!autofilter']={ref:XLSX.utils.encode_range({s:{r:6,c:0},e:{r:Math.max(6,values.length-1),c:6}})};
  ws['!rows']=[{hpt:32},{hpt:24},{hpt:21},{hpt:21},{hpt:21},{hpt:10},{hpt:30}];
  ws['!pageSetup']={orientation:'landscape',paperSize:9,fitToWidth:1,fitToHeight:0,fitToPage:true};
  ws['!printOptions']={horizontalCentered:true};
  ws['!margins']={left:.25,right:.25,top:.5,bottom:.5,header:.2,footer:.2};
  for(let c=0;c<7;c++){
   const addr=XLSX.utils.encode_cell({r:6,c});if(ws[addr])ws[addr].s={fill:{fgColor:{rgb:blue}},font:{bold:true,color:{rgb:white},name:'Aptos',sz:11},alignment:{horizontal:'center',vertical:'center'},border:{bottom:{style:'medium',color:{rgb:navy}}}};
  }
  for(let r=7;r<values.length;r++){ws['!rows'][r]={hpt:22};for(let c=0;c<7;c++){
   const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell)cell.s={font:{name:'Aptos',sz:10,color:{rgb:'243447'}},fill:{fgColor:{rgb:r%2?pale:white}},alignment:{horizontal:c===6?'center':'left',vertical:'center'},border:{bottom:{style:'thin',color:{rgb:'D8E1EA'}}}};
  }}
  for(const addr of ['A1','A2'])if(ws[addr])ws[addr].s={fill:{fgColor:{rgb:navy}},font:{name:'Aptos Display',bold:true,color:{rgb:white},sz:addr==='A1'?18:12},alignment:{vertical:'center'}};
  for(const addr of ['A3','A4','A5'])if(ws[addr])ws[addr].s={font:{name:'Aptos',sz:10,italic:true,color:{rgb:gray}}};
  return ws;
 }
 // El reporte imprimible siempre incluye TODAS las categorías, aunque la interfaz filtre HYM.
 XLSX.utils.book_append_sheet(wb,createSheet('REPORTE CONSOLIDADO PARA IMPRESIÓN',base,'Todas las categorías'),'Reporte imprimible');
 const summary=[['OPERACIÓN CD GIRARDOTA'],['RESUMEN GENERAL DE CARGUE Y DESPACHO'],['Fecha',dateText],['Categorías','Todas las categorías'],[],['Categoría','Transportes'],...categories.map(c=>[c,base.filter(r=>r.Categoría===c).length])];
 const missing=base.filter(r=>!categories.includes(r.Categoría));if(missing.length)summary.push(['Sin categoría',missing.length]);
 summary.push(['TOTAL GENERAL',base.length]);
 const sws=XLSX.utils.aoa_to_sheet(summary);sws['!merges']=[{s:{r:0,c:0},e:{r:0,c:1}},{s:{r:1,c:0},e:{r:1,c:1}}];sws['!cols']=[{wch:36},{wch:25}];
 for(const addr of ['A1','A2'])if(sws[addr])sws[addr].s={fill:{fgColor:{rgb:navy}},font:{name:'Aptos Display',bold:true,color:{rgb:white},sz:addr==='A1'?17:12}};
 for(const addr of ['A6','B6'])if(sws[addr])sws[addr].s={fill:{fgColor:{rgb:blue}},font:{bold:true,color:{rgb:white}}};
 XLSX.utils.book_append_sheet(wb,sws,'Resumen');
 for(const c of categories){const cr=base.filter(r=>r.Categoría===c);if(cr.length)XLSX.utils.book_append_sheet(wb,createSheet(c.toUpperCase(),cr,c),c.slice(0,31));}
 XLSX.writeFile(wb,'Operacion_CD_Girardota_Consolidado_'+new Date().toISOString().slice(0,10)+'.xlsx',{cellStyles:true});
} catch(err){console.error('[Exportación consolidada]',err);toast('No se pudo generar el reporte: '+err.message)}
finally{btn.disabled=false;btn.textContent=oldText;}
}
// v5.53: actualizacion colaborativa. Notificaciones Realtime + reconciliacion periodica.
let ctLiveChannel=null,ctLivePoll=null,ctLiveQueued=false,ctLiveBusy=false,ctLiveLast=0,ctLiveLastTapa=0;
function ctLiveEditing(){
 const e=document.activeElement;
 if(!e)return false;
 if(e.closest('#welcomeScreen,#login'))return true;
 return e.matches('input,textarea,select,[contenteditable="true"]')&&!e.matches('[readonly],[disabled]');
}
function ctLiveStatus(label,ok=true){const el=$('connection');if(el)el.textContent=(ok?'● ':'◌ ')+label;}
async function ctLiveUpdate(force=false){
 if(!sb||!loginPin||document.hidden||ctLiveBusy)return;
 if(!force&&ctLiveEditing()){ctLiveQueued=true;return;}
 ctLiveBusy=true;ctLiveQueued=false;
 try{
  await refresh();
  await refreshCheckin();
  if(activeModule==='tapa'&&Date.now()-ctLiveLastTapa>30000){await refreshTapa();ctLiveLastTapa=Date.now();}
  if(activeModule==='sync'&&typeof window.renderSyncReport==='function')window.renderSyncReport();
  ctLiveLast=Date.now();ctLiveStatus('Sincronizado · '+new Intl.DateTimeFormat('es-CO',{hour:'2-digit',minute:'2-digit',second:'2-digit',hour12:false,timeZone:'America/Bogota'}).format(new Date()));
 }catch(e){console.warn('[Sincronizacion]',e);ctLiveStatus('Reconectando datos…',false)}
 finally{ctLiveBusy=false;}
}
let ctLiveTimer=null;
function ctLiveSchedule(){
 clearTimeout(ctLiveTimer);
 ctLiveTimer=setTimeout(()=>void ctLiveUpdate(),450);
}
// v5.56: Sincronización manual global. No recarga la página ni modifica la pestaña activa.
let ctManualRefreshBusy=false;
async function ctManualRefreshAll(){
 const button=$('ctManualRefresh');
 if(ctManualRefreshBusy||!sb||!loginPin||!button)return;
 if(ctLiveEditing()&&!window.confirm('Hay un campo en edición. ¿Deseas refrescar los datos? Los cambios sin guardar podrían perderse.'))return;
 ctManualRefreshBusy=true;
 button.disabled=true;button.classList.add('is-refreshing');
 button.setAttribute('aria-busy','true');
 const label=button.querySelector('.ctRefreshLabel');if(label)label.textContent='Actualizando';
 const errors=[];
 async function attempt(name,fn){try{await fn()}catch(err){errors.push(name);console.warn('[Refrescar]',name,err)}}
 try{
  // Esperar una actualización automática en curso para evitar superposiciones.
  if(ctLiveBusy){for(let k=0;k<25&&ctLiveBusy;k++)await new Promise(resolve=>setTimeout(resolve,100))}
  await attempt('Cargue y Despacho',refresh);
  await attempt('CHECKIN',refreshCheckin);
  await attempt('Reportes CHECKIN',loadSharedCheckin);
  await attempt('Tapa Códigos KA',refreshTapa);
  await attempt('Sincronización',async()=>{window.loadSyncStatReport?.();window.renderSyncReport?.();});
  await attempt('Vista rápida de despacho',async()=>{window.loadBrujulaReport?.();});
  if(permissions.is_admin)await attempt('Historial',loadHistory);
  ctLiveLast=Date.now();
  if(errors.length)toast('Actualización parcial. No se pudo consultar: '+errors.join(', '));
  else {toast('Datos refrescados en todas las pestañas.');ctLiveStatus('Sincronizado manualmente');}
 }finally{
  ctManualRefreshBusy=false;button.disabled=false;button.classList.remove('is-refreshing');
  button.removeAttribute('aria-busy');if(label)label.textContent='Refrescar';
 }
}
function subscribe(){
 if(ctLiveChannel)return;
 ctLiveChannel=sb.channel('transportes-live-v553')
  .on('postgres_changes',{event:'*',schema:'public',table:'transportes'},ctLiveSchedule)
  .on('postgres_changes',{event:'*',schema:'public',table:'ct_viajes_en_linea'},ctLiveSchedule)
  .on('postgres_changes',{event:'*',schema:'public',table:'ct_checkin_archived'},ctLiveSchedule)
  .subscribe(status=>{if(status==='CHANNEL_ERROR'||status==='TIMED_OUT')ctLiveStatus('Verificando cambios automáticamente…',false)});
 // Fallback para bases sin publicaciones Realtime y cambios hechos por RPC.
 ctLivePoll=setInterval(()=>{if(document.hidden||!loginPin)return;void ctLiveUpdate()},12000);
 window.addEventListener('focus',()=>{if(loginPin)ctLiveSchedule()});
 document.addEventListener('visibilitychange',()=>{if(!document.hidden&&loginPin)ctLiveSchedule()});
 document.addEventListener('focusout',()=>{if(ctLiveQueued&&loginPin)setTimeout(()=>{if(!ctLiveEditing())void ctLiveUpdate(true)},500)},true);
}

function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]))}
function switchModule(module){const target=(module==="main"||module==="tapa")?"cargo":module==="brujula"?"rutero":module==="sync"?"cargo":"checkin";if(!canModule(target,"view")){toast("Sin permiso para acceder a esta pestaña.");return}const previousTop=(activeModule==="main"||activeModule==="tapa")?"cargo":activeModule==="brujula"?"rutero":activeModule==="sync"?"cargo":"checkin";
if(previousTop!==target && module!=="tapa")playModuleSplash(target);
activeModule=module;const cargo=module==="main"||module==="tapa",main=module==="main",checkin=module==="checkin",tapa=module==="tapa",brujula=module==="brujula",sync=module==="sync";if(cargo)activeTopModule="cargo";else activeTopModule=module;$("mainModuleView").classList.toggle("hidden",!main);$("checkinModuleView").classList.toggle("hidden",!checkin);$("tapaModuleView").classList.toggle("hidden",!tapa);$("brujulaModuleView").classList.toggle("hidden",!brujula);$("syncModuleView").classList.toggle("hidden",!sync);$("cargoSubnav").classList.toggle("hidden",!cargo);$("cargoPrimaryBtn").classList.toggle("active",cargo);$("checkinPrimaryBtn").classList.toggle("active",checkin);$("brujulaPrimaryBtn").classList.toggle("active",brujula);$("syncPrimaryBtn").classList.toggle("active",sync);$("cargoPrimaryBtn").setAttribute("aria-selected",String(cargo));$("checkinPrimaryBtn").setAttribute("aria-selected",String(checkin));$("brujulaPrimaryBtn").setAttribute("aria-selected",String(brujula));$("syncPrimaryBtn").setAttribute("aria-selected",String(sync));$("mainModuleBtn").classList.toggle("active",main);$("tapaModuleBtn").classList.toggle("active",tapa);$("mainModuleBtn").setAttribute("aria-selected",String(main));$("tapaModuleBtn").setAttribute("aria-selected",String(tapa));if(checkin)refreshCheckin();if(checkin)refreshCheckin();if(tapa)refreshTapa();if(brujula)loadBrujulaReport();if(sync){window.renderSyncReport?.();window.loadSyncStatReport?.();if(!records.length&&sb&&loginPin)void refresh()}}
let checkinRecords=[];
let checkinClosedArchive=new Set();
let checkinView='pending';
let checkinHasExported=false;
function checkinKey(row){return (row.is_online||Number(row.transport_id)<0?'online:':'normal:')+String(row.is_online ? -1000000-Number(row.transport_id) : row.transport_id);}
function checkinClosedRows(){return checkinRecords.filter(r=>r.checkin_estado==='Realizado'&&!checkinClosedArchive.has(checkinKey(r)));}
function showCheckinView(view){checkinView=view==='closed'?'closed':'pending';renderCheckin();}
function dateStamp(){const fmt=new Intl.DateTimeFormat('es-CO',{timeZone:'America/Bogota',day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit',hour12:false});return fmt.format(new Date());}
function downloadCheckinClosure(){
 const data=checkinClosedRows();
 if(!data.length){toast('No hay transportes realizados para descargar.');return;}
 if(typeof XLSX==='undefined'){toast('No se encuentra la biblioteca Excel.');return;}
 const now=new Date();
 const datePart=new Intl.DateTimeFormat('es-CO',{timeZone:'America/Bogota',day:'2-digit',month:'2-digit',year:'numeric'}).format(now);
 const timePart=new Intl.DateTimeFormat('es-CO',{timeZone:'America/Bogota',hour:'2-digit',minute:'2-digit',hour12:false}).format(now);
 // Encabezado y tabla con la misma estructura del formato impreso de referencia.
 const values=[
  ['CIERRE DE CHECKIN','','','','','',''],
  ['Fecha y hora:',`${datePart}    -    ${timePart}`,'','','','',''],
  ['Total realizados:',data.length,'','','','',''],
  ['','','','','','',''],
  ['TRANSPORTE','VIAJE','PLACA','PLACA FÍSICA','CLASE CARGUE','ACCIÓN','NOVEDAD']
 ];
 for(const r of data){values.push([
  String(r.transporte||''),String(r.viaje||''),String(r.placa||''),String(r.placa_fisica||''),
  r.is_online||Number(r.transport_id)<0?'En línea':'Preventa',
  'Realizado',String(r.checkin_novedad||'')
 ]);}
 const ws=XLSX.utils.aoa_to_sheet(values);
 const orange='F79646',dark='000000',light='FBD4B4',pale='FCE8D6',white='FFFFFF';
 const fineBorder={top:{style:'thin',color:{rgb:'D7D7D7'}},bottom:{style:'thin',color:{rgb:'D7D7D7'}},left:{style:'thin',color:{rgb:'D7D7D7'}},right:{style:'thin',color:{rgb:'D7D7D7'}}};
 const mainBorder={top:{style:'medium',color:{rgb:dark}},bottom:{style:'medium',color:{rgb:dark}},left:{style:'medium',color:{rgb:dark}},right:{style:'medium',color:{rgb:dark}}};
 const cellStyle=(r,c,style)=>{const addr=XLSX.utils.encode_cell({r,c});if(!ws[addr])ws[addr]={t:'s',v:''};ws[addr].s=style;};
 for(let c=0;c<7;c++)cellStyle(0,c,{font:{name:'Aptos',bold:true,sz:14,color:{rgb:white}},fill:{patternType:'solid',fgColor:{rgb:orange}},alignment:{horizontal:'center',vertical:'center'},border:mainBorder});
 for(let r=1;r<=2;r++)for(let c=0;c<2;c++)cellStyle(r,c,{font:{name:'Aptos',bold:c===0,sz:11,color:{rgb:c===0?white:dark}},fill:{patternType:'solid',fgColor:{rgb:c===0?orange:white}},alignment:{horizontal:c===0?'center':'left',vertical:'center',indent:c===1?1:0},border:mainBorder});
 for(let c=0;c<7;c++)cellStyle(4,c,{font:{name:'Aptos',bold:true,sz:11,color:{rgb:white}},fill:{patternType:'solid',fgColor:{rgb:orange}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:mainBorder});
 for(let r=5;r<values.length;r++)for(let c=0;c<7;c++){
  const fill=r%2===1?light:pale;
  cellStyle(r,c,{font:{name:'Aptos',sz:10,color:{rgb:dark}},fill:{patternType:'solid',fgColor:{rgb:fill}},alignment:{horizontal:'center',vertical:'center',wrapText:c===6},border:fineBorder});
 }
 ws['!merges']=[{s:{r:0,c:0},e:{r:0,c:6}}];
 ws['!cols']=[{wch:20},{wch:24},{wch:17},{wch:20},{wch:22},{wch:17},{wch:38}];
 ws['!rows']=[{hpt:26},{hpt:22},{hpt:22},{hpt:13},{hpt:31},...data.map(()=>({hpt:24}))];
 ws['!autofilter']={ref:`A5:G${values.length}`};
 ws['!pageSetup']={orientation:'landscape',paperSize:9,fitToWidth:1,fitToHeight:0};
 ws['!margins']={left:0.25,right:0.25,top:0.35,bottom:0.35,header:0.15,footer:0.15};
 ws['!print_area']=`A1:G${values.length}`;
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Cierre CHECKIN');
 const fileDate=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Bogota',year:'numeric',month:'2-digit',day:'2-digit'}).format(now);
 XLSX.writeFile(wb,`Cierre_Checkin_${fileDate}.xlsx`,{cellStyles:true});
 checkinHasExported=true;
 toast('Excel de cierre generado con formato de impresión. Conserva el archivo antes de limpiar.');
}

async function clearCheckinClosure(){const rows=checkinClosedRows();if(!rows.length){toast('No hay datos de cierre para limpiar.');return;}
 if(!canModule('checkin','edit')||!(permissions.is_admin||permissions.can_manage_checkin)){toast('Sin permiso para limpiar cierres.');return;}
 if(!checkinHasExported){toast('Primero descarga el Excel de Cierre CHECKIN.');return;}
 if(!confirm('¿Confirmas que guardaste el Excel? Se archivarán '+rows.length+' transportes realizados y desaparecerán del Cierre CHECKIN para todos los usuarios. Esta acción no se puede deshacer desde la aplicación.'))return;
 const keys=rows.map(checkinKey);
 const {error}=await sb.rpc('ct_checkin_close_archive_secure',{p_pin:loginPin,p_keys:keys});
 if(error){toast('Error al limpiar cierre: '+error.message);return;}
 keys.forEach(x=>checkinClosedArchive.add(x));checkinHasExported=false;renderCheckin();toast('Cierre limpiado correctamente. Los transportes originales se conservaron.');
}

async function refreshTapa(){if(!sb||!loginPin)return;const{data,error}=await sb.rpc("list_tapa_codigos_ka_secure",{p_pin:loginPin});if(error){toast("No se pudo cargar Tapa Códigos KA. Ejecuta el SQL incluido: "+error.message);return}tapaRecords=data||[];await fetchTapaStates();const select=$("tapaFamilyFilter"),current=select.value,clientSelect=$("tapaClientFilter"),currentClient=clientSelect.value;const clients=[...new Set(tapaRecords.map(r=>String(r.descripcion||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));clientSelect.innerHTML='<option value="">Todos los clientes</option>'+clients.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join("");if(clients.includes(currentClient))clientSelect.value=currentClient;const families=[...new Set(tapaRecords.map(r=>String(r.familia||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));select.innerHTML='<option value="">Todas las familias</option>'+families.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join("");if(families.includes(current))select.value=current;renderTapa()}
let tapaStatusMap=new Map();
async function fetchTapaStates(){if(!canModule("cargo","view"))return;const{data,error}=await sb.rpc("list_tapa_status_secure",{p_pin:loginPin});if(error){toast("No fue posible consultar estados de tapa códigos: "+error.message);return}tapaStatusMap=new Map((data||[]).map(x=>[x.group_key,x.status]));syncTapaStates()}
function paintTapaState(el){el.classList.remove("tapaComplete","tapaPartial","tapaMissing");if(el.value==="Completo")el.classList.add("tapaComplete");else if(el.value==="Parcial")el.classList.add("tapaPartial");else if(el.value==="Sin tapa códigos")el.classList.add("tapaMissing");}
function syncTapaStates(){document.querySelectorAll(".tapaState").forEach(el=>{const key=decodeURIComponent(el.dataset.tapaKey);el.value=tapaStatusMap.get(key)||"";paintTapaState(el);el.onchange=async()=>{if(!canModule("cargo","edit")){toast("No tienes permiso de edición");el.value=tapaStatusMap.get(key)||"";paintTapaState(el);return}const previous=tapaStatusMap.get(key)||"";paintTapaState(el);const{error}=await sb.rpc("set_tapa_status_secure",{p_pin:loginPin,p_group_key:key,p_status:el.value||null});if(error){el.value=previous;paintTapaState(el);toast(error.message);return}tapaStatusMap.set(key,el.value);toast("Estado de Tapa Códigos guardado")}})}
function renderTapa(){const family=$("tapaFamilyFilter")?.value||"",client=$("tapaClientFilter")?.value||"",q=($("tapaSearch")?.value||"").trim().toLocaleLowerCase("es");const filtered=tapaRecords.filter(r=>(!family||String(r.familia||"")===family)&&(!client||String(r.descripcion||"")===client)&&(!q||[r.transporte,r.vehiculo,r.viaje,r.material,r.denominacion,r.descripcion].some(v=>String(v||"").toLocaleLowerCase("es").includes(q))));const groups=new Map();for(const r of filtered){const key=[r.transporte,r.vehiculo,r.viaje,r.material,r.denominacion].map(v=>String(v||"").trim()).join("\u001f");const item=groups.get(key)||{transporte:r.transporte||"",vehiculo:r.vehiculo||"",viaje:r.viaje||"",material:r.material||"",denominacion:r.denominacion||"",cantidad:0};item.cantidad+=Number(r.cantidad)||0;groups.set(key,item)}const rows=[...groups.values()].sort((a,b)=>[a.transporte,a.vehiculo,a.viaje,a.material,a.denominacion].join("|").localeCompare([b.transporte,b.vehiculo,b.viaje,b.material,b.denominacion].join("|"),"es"));$("tapaRows").innerHTML=rows.map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.vehiculo)}</td><td>${esc(r.viaje)}</td><td>${esc(r.material)}</td><td>${esc(r.denominacion)}</td><td>${(r.cantidad||0).toLocaleString("es-CO",{maximumFractionDigits:3})}</td><td><select class="tapaState" data-tapa-key="${encodeURIComponent([r.transporte,r.vehiculo,r.viaje,r.material,r.denominacion].map(v=>String(v||"").trim()).join("\u001f"))}" ${!canModule("cargo","edit")?"disabled":""}><option value="">Seleccionar</option><option value="Completo">Completo</option><option value="Parcial">Parcial</option><option value="Sin tapa códigos">Sin tapa códigos</option></select></td></tr>`).join("")||'<tr><td colspan="7">No hay datos para los filtros seleccionados.</td></tr>';$("tapaTotal").textContent=rows.reduce((s,r)=>s+r.cantidad,0).toLocaleString("es-CO",{maximumFractionDigits:3});$("tapaSummary").textContent=`${rows.length} combinaciones · ${filtered.length} registros de origen`;syncTapaStates();}

async function refreshCheckin(){if(!sb||!loginPin)return;await loadOnlineTrips();const{data,error}=await sb.rpc("list_checkin_transports_secure",{p_pin:loginPin});if(error){toast("No se pudo cargar Checkin. Ejecuta el SQL de Checkin: "+error.message);return}const onlineLoaded=onlineTrips.filter(r=>r.estado==="Cargado").map(r=>({transport_id:-1000000-Number(r.id),transporte:r.transporte,viaje:r.viaje,placa:r.placa,placa_fisica:"",checkin_estado:r.checkin_estado||"Pendiente",checkin_novedad:r.checkin_novedad||"",is_online:true}));
checkinRecords=[...(data||[]),...onlineLoaded];
const archived=await sb.rpc('ct_checkin_closed_list_secure',{p_pin:loginPin});
if(!archived.error)checkinClosedArchive=new Set((archived.data||[]).map(x=>String(x.trip_key)));
renderCheckin()}
function renderCheckin(){
 const rows=$('checkinRows');if(!rows)return;
 const open=checkinRecords.filter(r=>r.checkin_estado!=='Realizado'&&!checkinClosedArchive.has(checkinKey(r)));
 const closed=checkinClosedRows();const total=open.length+closed.length;
 const stats=$('ctCheckinStats');if(stats)stats.innerHTML=`<div class="ctCheckinNotebook"><span>📘</span><b>${total}</b><small>Total pendientes por CHECKIN</small></div><div class="ctCheckinNotebook"><span>📘</span><b>${closed.length}</b><small>Realizados</small></div><div class="ctCheckinNotebook red"><span>📕</span><b>${open.length}</b><small>Pendientes</small></div>`;
 const pendingBtn=$('ctCheckinPendingTab'),closedBtn=$('ctCheckinClosedTab');
 pendingBtn?.classList.toggle('active',checkinView==='pending');closedBtn?.classList.toggle('active',checkinView==='closed');
 const standardPanel=$('ctCheckinStandardPanel'),closurePanel=$('ctCheckinClosurePanel');
 if(standardPanel)standardPanel.classList.toggle('hidden',checkinView!=='pending');
 if(closurePanel)closurePanel.classList.toggle('hidden',checkinView!=='closed');
 if(checkinView==='closed'){
  const body=$('ctCheckinClosureRows');if(body)body.innerHTML=closed.map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td>${esc(r.placa)}</td><td>${esc(r.placa_fisica||'')}</td><td>${r.is_online||Number(r.transport_id)<0?'En línea':'Preventa'}</td><td>Realizado</td><td>${esc(r.checkin_novedad||'')}</td></tr>`).join('')||'<tr><td colspan="7">No hay transportes realizados pendientes de cierre.</td></tr>';
  $('ctCheckinClosureCount').textContent=closed.length+' transportes para cierre';return;
 }
 rows.innerHTML=open.map(r=>{const sel=String(selectedCheckinId)===String(r.transport_id);const canEdit=canModule('checkin','edit')&&(permissions.is_admin||permissions.can_manage_checkin);const cls=r.checkin_estado==='Pendiente'?'status-pendiente':'status-novedad';const beesStatus=getBeesStatus(r.transporte);const porteriaStatus=getPorteriaStatus(r.transporte);const beesClass=beesStatus==='REALIZADO'?'beesMatch':beesStatus==='N/A'?'beesMissing':'';
 return `<tr class="${sel?'checkinSelected':''}"><td><button type="button" class="${sel?'selectedTransportBtn':'secondary'}" onclick="selectCheckin(${r.transport_id})">${sel?'✓ Seleccionado':'Seleccionar'}</button></td><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td><b>${esc(r.placa)}</b></td><td><input class="checkinPhysicalPlate" value="${esc(r.placa_fisica||'')}" readonly aria-label="Placa Física" /></td><td><span class="checkinLoadClass ${r.is_online||Number(r.transport_id)<0?'checkinLoadOnline':'checkinLoadPreventa'}">${r.is_online||Number(r.transport_id)<0?'En línea':'Preventa'}</span></td><td><select class="checkinStatus ${cls}" ${(!sel||!canEdit)?'disabled':''} onchange="saveCheckin(${r.transport_id},this.value,null)"><option value="Pendiente" ${r.checkin_estado==='Pendiente'?'selected':''}>Pendiente</option><option value="Realizado">Realizado</option><option value="Con novedad" ${r.checkin_estado==='Con novedad'?'selected':''}>Con novedad</option></select></td><td><input class="checkinNovelty" value="${esc(r.checkin_novedad||'')}" placeholder="Escribir novedad…" ${(!sel||!canEdit)?'disabled':''} onchange="saveCheckin(${r.transport_id},null,this.value)" /></td><td><span class="beesBadge ${beesClass}">${beesStatus}</span></td><td><span class="porteriaBadge ${porteriaStatus==='Salida'?'porteriaYes':porteriaStatus==='Sin salida'?'porteriaNo':'porteriaUnknown'}">${porteriaStatus}</span></td></tr>`}).join('')||'<tr><td colspan="10">No hay transportes pendientes de CHECKIN.</td></tr>';
}
function selectCheckin(id){selectedCheckinId=String(selectedCheckinId)===String(id)?null:String(id);const r=checkinRecords.find(x=>String(x.transport_id)===String(id));toast(selectedCheckinId?`Seleccionado: ${r?.transporte||""} · Viaje ${r?.viaje||""}`:"Selección cancelada");renderCheckin()}
async function saveCheckin(id,status,novelty){if(!canModule("checkin","edit")){toast("Sin permiso de modificación en el módulo.");return}if(String(selectedCheckinId)!==String(id)){toast("Primero selecciona el transporte.");await refreshCheckin();return}if(!(permissions.is_admin||permissions.can_manage_checkin)){toast("No tienes permiso para modificar Checkin.");await refreshCheckin();return}const r=checkinRecords.find(x=>String(x.transport_id)===String(id));const isOnline=Number(id)<0;
const rpc=isOnline?"ct_online_trip_checkin_secure":"update_checkin_transport_secure";
const args=isOnline?{p_id:-1000000-Number(id),p_status:status===null?r?.checkin_estado:status,p_novedad:novelty===null?r?.checkin_novedad:novelty,p_pin:loginPin}:{p_transport_id:id,p_status:status===null?r?.checkin_estado:status,p_novedad:novelty===null?r?.checkin_novedad:novelty,p_pin:loginPin};
const{error}=await sb.rpc(rpc,args);if(error){toast(error.message);await refreshCheckin();return}toast("Check-in actualizado.");await refreshCheckin()}


let porteriaTransportes=null; // null: no se ha cargado archivo
function getPorteriaStatus(transporte){if(porteriaTransportes===null)return "Sin datos";return porteriaTransportes.has(normalizeTransport(transporte))?"Salida":"Sin salida"}
async function loadPorteriaExcel(){if(!canModule("checkin","edit")){toast("Sin permiso de modificación en el módulo.");return}
 const file=$("porteriaExcelFile")?.files?.[0];
 if(!file){toast("Selecciona el archivo Excel Salidas Por Portería.");return}
 try{
   const wb=XLSX.read(await file.arrayBuffer(),{type:"array",raw:false});
   const sh=wb.Sheets[wb.SheetNames[0]];
   if(!sh)throw Error("No se encontró la primera hoja del Excel.");
   const rows=XLSX.utils.sheet_to_json(sh,{defval:"",raw:false});
   if(!rows.length)throw Error("El Excel no contiene filas de datos.");
   const normalizeHeader=t=>String(t||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toLowerCase().replace(/\s+/g," ");
   const transportKey=Object.keys(rows[0]).find(k=>normalizeHeader(k)==="transporte");
   if(!transportKey)throw Error("No existe una columna llamada Transporte.");
   const ids=new Set(rows.map(r=>normalizeTransport(r[transportKey])).filter(Boolean));
   if(!ids.size)throw Error("La columna Transporte no contiene identificadores válidos.");
   await CTShared.save('porteria',file.name,[...ids],{source_rows:rows.length});
   porteriaTransportes=ids;
   const coincidencias=checkinRecords.filter(r=>ids.has(normalizeTransport(r.transporte))).length;
   $("porteriaSummary").textContent=`Archivo: ${file.name} · ${rows.length.toLocaleString("es-CO")} registros · ${ids.size.toLocaleString("es-CO")} transportes únicos · ${coincidencias.toLocaleString("es-CO")} coincidencias en CHECKIN. Cruce guardado en Supabase para los usuarios autorizados.`;
   renderCheckin();render();toast("Salidas por Portería cargadas correctamente.");
 }catch(err){console.error(err);$("porteriaSummary").textContent="Error: "+err.message;toast("No se pudo procesar Portería: "+err.message)}
}
async function clearPorteria(){if(!canModule('checkin','edit'))return toast('Sin permiso');try{await CTShared.clear('porteria')}catch(e){return toast(e.message)}porteriaTransportes=null;$("porteriaExcelFile").value="";$("porteriaSummary").textContent="Sin archivo de portería: el estado aparecerá como Sin datos.";renderCheckin();render();toast("Cruce de Portería limpiado.")}

let beesCsvRows=[];
function getBeesStatus(transport){if(!beesCsvRows.length)return "Sin datos";const first=beesCsvRows[0]||{};const routeKey=Object.keys(first).find(k=>String(k).trim().toLowerCase()==="route_external_id");if(!routeKey)return "Sin datos";const id=normalizeTransport(transport);return beesCsvRows.some(r=>normalizeTransport(r[routeKey])===id)?"REALIZADO":"N/A"}
function normalizeTransport(value){
  let s=String(value??"").trim().toUpperCase().replace(/\s+/g,"");
  s=s.replace(/^S(?=\d)/,"");
  return s.replace(/[^A-Z0-9]/g,"");
}
function csvCell(row,names){
  const keys=Object.keys(row||{});
  const key=keys.find(k=>names.includes(String(k).trim().toLowerCase()));
  return key?row[key]:"";
}
async function compareBeesCsv(){if(!canModule("checkin","edit")){toast("Sin permiso de modificación en el módulo.");return}
  const file=$("beesCsvFile")?.files?.[0];
  if(!file){toast("Selecciona primero el CSV descargado de BEES.");return}
  try{
    const buffer=await file.arrayBuffer();
    const workbook=XLSX.read(buffer,{type:"array",raw:false});
    const sheet=workbook.Sheets[workbook.SheetNames[0]];
    const raw=XLSX.utils.sheet_to_json(sheet,{defval:"",raw:false});
    if(!raw.length){toast("El archivo de BEES está vacío.");return}
    const first=raw[0], routeKey=Object.keys(first).find(k=>String(k).trim().toLowerCase()==="route_external_id");
    if(!routeKey){toast("No encuentro la columna route_external_id en el CSV.");return}
    const newBeesRows=raw.filter(r=>normalizeTransport(r[routeKey]));
    await CTShared.save('bees',file.name,newBeesRows,{source_rows:raw.length});
    beesCsvRows=newBeesRows;
    const byTransport=new Map();
    for(const r of beesCsvRows){
      const id=normalizeTransport(r[routeKey]);
      if(!byTransport.has(id))byTransport.set(id,{rows:[],plates:new Set(),drivers:new Set()});
      const item=byTransport.get(id);item.rows.push(r);
      const plate=csvCell(r,["vehicle_license_plate"]);if(plate)item.plates.add(String(plate).trim());
      const driver=csvCell(r,["driver_name"]);if(driver)item.drivers.add(String(driver).trim());
    }
    const appMap=new Map();
    for(const r of checkinRecords){const id=normalizeTransport(r.transporte);if(id&&!appMap.has(id))appMap.set(id,[]);if(id)appMap.get(id).push(r)}
    const results=[];
    for(const [id,appRows] of appMap){
      const b=byTransport.get(id);
      results.push({id,status:b?"Coincide":"Pendiente en BEES",cls:b?"beesMatch":"beesMissing",plates:b?[...b.plates].join(", "):"",drivers:b?[...b.drivers].join(", "):"",count:b?b.rows.length:0});
    }
    for(const [id,b] of byTransport){
      if(!appMap.has(id))results.push({id,status:"Solo en BEES",cls:"beesOnly",plates:[...b.plates].join(", "),drivers:[...b.drivers].join(", "),count:b.rows.length});
    }
    results.sort((a,b)=>a.id.localeCompare(b.id,undefined,{numeric:true}));
    $("beesReportRows").innerHTML=results.map(r=>`<tr class="${r.cls}"><td><b>${esc(r.id)}</b></td><td class="beesBadge">${esc(r.status)}</td><td>${esc(r.plates)}</td><td>${esc(r.drivers)}</td><td>${r.count}</td></tr>`).join("")||'<tr><td colspan="5">No hay diferencias ni coincidencias para mostrar.</td></tr>';
    const matched=results.filter(r=>r.status==="Coincide").length,missing=results.filter(r=>r.status==="Pendiente en BEES").length,only=results.filter(r=>r.status==="Solo en BEES").length;
    $("beesReportSummary").textContent=`Archivo: ${file.name} · ${byTransport.size} transportes únicos en BEES · ${appMap.size} transportes en Checkin · Coinciden: ${matched} · Pendientes en BEES: ${missing} · Solo en BEES: ${only}.`;
    $("beesReport").classList.remove("hidden");
    renderCheckin();
    toast("Cruce BEES completado. No se modificaron registros.");
  }catch(e){console.error(e);toast("No se pudo leer el CSV. Descárgalo nuevamente desde BEES.")}
}
async function clearBeesReport(){
  if(!canModule('checkin','edit'))return toast('Sin permiso');try{await CTShared.clear('bees')}catch(e){return toast(e.message)}
  beesCsvRows=[];$("beesCsvFile").value="";$("beesReportRows").innerHTML="";$("beesReport").classList.add("hidden");$("beesReportSummary").textContent="Todavía no se ha cargado un archivo de BEES.";renderCheckin();toast("Reporte limpiado.");
}

init();

function deleteAllTransports(){if(!permissions.can_delete_all_transports&&!permissions.is_admin){toast("No tienes permiso para eliminar todos los transportes.");return}if(!records.length){toast("No hay transportes para eliminar.");return}$("deleteConfirmCount").textContent=String(records.length);$("deleteConfirmModal").classList.remove("hidden");$("confirmDeleteBtn").focus()}function closeDeleteModal(){$("deleteConfirmModal").classList.add("hidden")}async function confirmDeleteAll(){if(!permissions.can_delete_all_transports&&!permissions.is_admin){closeDeleteModal();toast("No tienes permiso para eliminar todos los transportes.");return}if(!records.length){closeDeleteModal();toast("No hay transportes para eliminar.");return}const count=records.length;$("confirmDeleteBtn").disabled=true;$("confirmDeleteBtn").textContent="Eliminando…";const{data,error}=await sb.rpc("delete_all_transports_secure",{p_pin:loginPin});$("confirmDeleteBtn").disabled=false;$("confirmDeleteBtn").textContent="Sí, eliminar todos";if(error){closeDeleteModal();toast("No se pudo eliminar: "+error.message);return}closeDeleteModal();toast(`Se eliminaron ${data?.deleted??count} transportes y se vació Tapa Códigos KA. El historial se conservó.`);await refresh()}

async function loadSharedCheckin(){
 if(!canModule('checkin','view')&&!canModule('cargo','view'))return;
 try{
  const [port,bees]=await Promise.all([CTShared.load('porteria'),canModule('checkin','view')?CTShared.load('bees'):Promise.resolve(null)]);
  porteriaTransportes=port?new Set(port.items):null;
  if(port)$('porteriaSummary').textContent=`Archivo compartido: ${port.filename} · ${port.items.length.toLocaleString('es-CO')} transportes · Actualizado: ${new Date(port.saved_at).toLocaleString('es-CO')}`;
  else $('porteriaSummary').textContent='Sin archivo de portería guardado. El cruce permanecerá hasta pulsar Limpiar cruce.';
  beesCsvRows=bees?.items||[];
  if(bees){$('beesReportSummary').textContent=`BEES compartido: ${bees.filename} · ${beesCsvRows.length.toLocaleString('es-CO')} registros · Actualizado: ${new Date(bees.saved_at).toLocaleString('es-CO')}`;renderBeesSavedComparison();}
  renderCheckin();render();
 }catch(e){console.error(e);toast('No se pudieron recuperar los reportes compartidos: '+e.message)}
}
function renderBeesSavedComparison(){
 const key=Object.keys(beesCsvRows[0]||{}).find(k=>String(k).trim().toLowerCase()==='route_external_id');if(!key)return;
 const by=new Map();for(const row of beesCsvRows){const id=normalizeTransport(row[key]);if(!id)continue;let v=by.get(id);if(!v){v={plates:new Set(),drivers:new Set(),count:0};by.set(id,v)}v.count++;const plate=csvCell(row,['vehicle_license_plate']);const driver=csvCell(row,['driver_name']);if(plate)v.plates.add(plate);if(driver)v.drivers.add(driver)}
 const app=new Set(checkinRecords.map(r=>normalizeTransport(r.transporte)).filter(Boolean));const results=[];
 for(const id of app){const v=by.get(id);results.push({id,status:v?'Coincide':'Pendiente en BEES',cls:v?'beesMatch':'beesMissing',plates:v?[...v.plates].join(', '):'',drivers:v?[...v.drivers].join(', '):'',count:v?.count||0})}
 for(const [id,v] of by)if(!app.has(id))results.push({id,status:'Solo en BEES',cls:'beesOnly',plates:[...v.plates].join(', '),drivers:[...v.drivers].join(', '),count:v.count});
 results.sort((a,b)=>a.id.localeCompare(b.id,undefined,{numeric:true}));
 $('beesReportRows').innerHTML=results.map(r=>`<tr class="${r.cls}"><td><b>${esc(r.id)}</b></td><td>${esc(r.status)}</td><td>${esc(r.plates)}</td><td>${esc(r.drivers)}</td><td>${r.count}</td></tr>`).join('');
 $('beesReport').classList.remove('hidden');
}

/* v5.5: vehículos pendientes de salida en portería */
function getPendingPorteriaVehicles(){
  if(typeof porteriaTransportes==='undefined'||porteriaTransportes===null)return null;
  return (window.currentCargoFilteredRecords||[]).filter(r=>getPorteriaStatus(r.transporte)==='Sin salida');
}
function updatePorteriaPendingCount(){
  const counter=document.getElementById('porteriaPendingCount');
  const button=document.getElementById('openPorteriaPending');
  if(!counter||!button)return;
  const list=getPendingPorteriaVehicles();
  counter.textContent=list===null?'—':list.length.toLocaleString('es-CO');
  button.setAttribute('aria-label',list===null?'Consultar vehículos sin salida: no hay archivo de portería':`Ver ${list.length} vehículos sin salida`);
}
function openPorteriaPending(){
  if(!canModule('cargo','view'))return;
  const list=getPendingPorteriaVehicles();
  const summary=document.getElementById('porteriaPendingSummary');
  const rows=document.getElementById('porteriaPendingRows');
  if(list===null){summary.textContent='No hay archivo de Salidas por Portería cargado. Cárgalo desde CHECKIN para realizar el cruce.';rows.innerHTML='';}
  else{
    summary.textContent=`${list.length.toLocaleString('es-CO')} transporte(s) sin salida entre los registros filtrados.`;
    rows.innerHTML=list.length?list.map(r=>`<tr><td><strong>${esc(r.transporte)}</strong></td><td>${esc(r.viaje||'—')}</td><td>${esc(r.placa||'—')}</td></tr>`).join(''):'<tr><td colspan="3" class="porteriaNoPending">No hay vehículos sin salida con los filtros actuales.</td></tr>';
  }
  document.getElementById('porteriaPendingModal').classList.remove('hidden');
  document.getElementById('closePorteriaPending').focus();
}
function closePorteriaPending(){
  const modal=document.getElementById('porteriaPendingModal');
  if(!modal||modal.classList.contains('hidden'))return;
  modal.classList.add('hidden');
  document.getElementById('openPorteriaPending')?.focus();
}
