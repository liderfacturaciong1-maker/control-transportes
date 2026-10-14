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
const SYNC_VERSION='v5.18 status columna I'; // Motor de sincronización v5.18
let cargoFallback=[],cargoFetching=false,cargoLastAttempt=0,cargoError='';
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
 const si=8;
 const quick=new Map();
 if(report.loaded&&ti>=0){
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
 const rows=cargo.filter(r=>r&&typeof r==='object').map(r=>{const key=transportKey(r.transporte);const statuses=quick.get(key);const status=statuses?.size?[...statuses].join(' · '):'—';
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
 else syncInfo=`${filtered.length.toLocaleString('es-CO')} registros de Cargue y Despacho visibles. ${report.loaded?'Status leído de la columna I de Vista rápida.':'Sin informe de Vista rápida cargado; Status pendiente.'} ${statLoaded?'Cruce STAT: '+statFile+'.':''}`;
 $('syncMessage').textContent='Motor de sincronización v5.18 · '+syncInfo;
 updateUnsyncedCount();
 $('syncRows').innerHTML=filtered.slice(0,1000).map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td>${esc(r.placa)}</td><td>${esc(r.status)}</td><td><span class="syncBadge ${/^(Sincronizado|En STAT)$/i.test(r.result)?'syncGood':/^(No encontrado|No está en STAT)$/i.test(r.result)?'syncBad':'syncUnknown'}">${esc(r.result)}</span></td></tr>`).join('') || `<tr><td colspan="5" style="text-align:center;padding:20px">${esc(cargoError ? 'Error al consultar transportes: '+cargoError : cargoFetching ? 'Cargando transportes...' : 'No hay transportes disponibles. Pulsa Actualizar cruce.')}</td></tr>`;
 }catch(err){const msg='Error al presentar Sincronización: '+(err?.message||String(err));const m=$('syncMessage');if(m)m.textContent=msg;const body=$('syncRows');if(body)body.innerHTML='<tr><td colspan="5">'+esc(msg)+'</td></tr>';console.error('[Sincronización]',err);}
}

function openUnsynced(){
 const modal=$('syncAlertModal');const holder=$('syncAlertItems');if(!modal||!holder)return;
 const invalid=currentSyncRows.filter(r=>String(r.result||'').trim().toUpperCase()!=='S-S-U');
 $('syncAlertTotal').textContent=invalid.length.toLocaleString('es-CO');
 holder.innerHTML=invalid.length?invalid.map((r,i)=>`<details class="syncAlertItem"><summary><strong>${esc(r.transporte)}</strong><span>${esc(r.result)}</span></summary><dl><div><dt>Transporte</dt><dd>${esc(r.transporte)}</dd></div><div><dt>Placa</dt><dd>${esc(r.placa)}</dd></div><div><dt>Viaje</dt><dd>${esc(r.viaje)}</dd></div><div><dt>Estado sincronización</dt><dd>${esc(r.result)}</dd></div></dl></details>`).join(''):'<p>Todos los transportes tienen el estado S-S-U.</p>';
 modal.classList.remove('hidden');document.body.classList.add('syncAlertOpen');$('syncAlertClose')?.focus();
}
function closeUnsynced(){ $('syncAlertModal')?.classList.add('hidden');document.body.classList.remove('syncAlertOpen');$('syncAlertBtn')?.focus(); }
function updateUnsyncedCount(){const n=currentSyncRows.filter(r=>String(r.result||'').trim().toUpperCase()!=='S-S-U').length;const el=$('syncAlertCount');if(el)el.textContent=n.toLocaleString('es-CO');}

window.renderSyncReport=()=>{render();if(!cargoSource().length)void loadCargoSource();};window.loadSyncStatReport=()=>{render();void loadCargoSource(!cargoSource().length);void restore();};
window.addEventListener('rutero-updated',render);
window.addEventListener('cargo-records-updated',render);
window.addEventListener('cargo-records-updated',()=>{cargoFallback=[];render();});
window.addEventListener('focus',()=>{if(!$('syncModuleView')?.classList.contains('hidden')){render();void loadCargoSource();}});
function initializeSync514(){for(const id of ['syncSearch','syncFilter'])$(id)?.addEventListener('input',render);$('syncStatUpload')?.addEventListener('click',upload);$('syncRefresh')?.addEventListener('click',()=>{void loadCargoSource(true);void restore();render();});$('syncAlertBtn')?.addEventListener('click',openUnsynced);$('syncAlertClose')?.addEventListener('click',closeUnsynced);$('syncAlertModal')?.addEventListener('click',e=>{if(e.target===$('syncAlertModal'))closeUnsynced();});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('syncAlertModal')?.classList.contains('hidden'))closeUnsynced();});render();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initializeSync514,{once:true});else initializeSync514();
})();

const cfg=window.APP_CONFIG||{};let sb=null,user=null,records=[],pendingImportRows=null,pendingTapaRows=null,tapaRecords=[],loginPin="",selectedTransportId=null,selectedCheckinId=null,activeModule="main",activeTopModule="cargo",activeCategory="Viajes HYM y URBAN",selectedCategories=[],categoryMultiMode=false,permissions={is_admin:false,can_upload:false,can_create_users:false,can_delete_all_transports:false,can_arm:false,can_count:false,can_load:false,can_edit_availability:false,can_edit_physical_plate:false,can_manage_checkin:false};window.getCargoSyncRecords = () => records;
const $=id=>document.getElementById(id);
function toast(m){$("toast").textContent=m;$("toast").classList.add("show");setTimeout(()=>$("toast").classList.remove("show"),3000)}
function configured(){return cfg.SUPABASE_URL&&!cfg.SUPABASE_URL.includes("TU-PROYECTO")&&cfg.SUPABASE_ANON_KEY&&!cfg.SUPABASE_ANON_KEY.includes("TU_ANON")}
async function init(){if(!configured()){$("loginMsg").innerHTML="Falta configurar <b>config.js</b>. Revisa README.";return}sb=supabase.createClient(cfg.SUPABASE_URL,cfg.SUPABASE_ANON_KEY);$("mainModuleBtn").onclick=()=>switchModule("main");$("tapaModuleBtn").onclick=()=>switchModule("tapa");$("cargoPrimaryBtn").onclick=()=>switchModule(activeTopModule==="cargo"?(activeModule==="tapa"?"tapa":"main"):"main");$("checkinPrimaryBtn").onclick=()=>switchModule("checkin");$("brujulaPrimaryBtn").onclick=()=>switchModule("brujula");$("syncPrimaryBtn").onclick=()=>switchModule("sync");$("tapaFamilyFilter").onchange=renderTapa;$("tapaClientFilter").onchange=renderTapa;$("tapaSearch").oninput=renderTapa;$("loginBtn").onclick=login;$("pin").onkeydown=e=>e.key==="Enter"&&login();$("logout").onclick=logout;$("search").oninput=render;$("openPorteriaPending").onclick=openPorteriaPending;$("closePorteriaPending").onclick=closePorteriaPending;$("porteriaPendingModal").addEventListener("click",e=>{if(e.target.id==="porteriaPendingModal")closePorteriaPending()});document.addEventListener("keydown",e=>{if(e.key==="Escape")closePorteriaPending()});function updateCategoryTabs(){document.querySelectorAll(".categoryTab[data-category]").forEach(btn=>{const on=categoryMultiMode?selectedCategories.includes(btn.dataset.category):btn.dataset.category===activeCategory;btn.classList.toggle("active",on);btn.setAttribute("aria-selected",String(on))});$("clearCategorySelection").style.display=categoryMultiMode?"inline-block":"none";$("categorySelectionHint").textContent=categoryMultiMode?"Selección múltiple activa: haz clic en las demás pestañas para agregarlas o quitarlas.":"Para combinar pestañas, haz doble clic en la primera pestaña y luego haz clic en las demás."}document.querySelectorAll(".categoryTab[data-category]").forEach(btn=>{btn.addEventListener("dblclick",e=>{e.preventDefault();const cat=btn.dataset.category;if(cat==="Todos"){categoryMultiMode=false;selectedCategories=[];activeCategory="Todos";toast("La pestaña Todos ya reúne todas las categorías.");}else{categoryMultiMode=true;selectedCategories=[cat];activeCategory="__MULTI__";toast("Selección múltiple activada. Ahora elige las demás pestañas.");}updateCategoryTabs();render()});btn.addEventListener("click",()=>{const cat=btn.dataset.category;if(categoryMultiMode){if(cat==="Todos"){categoryMultiMode=false;selectedCategories=[];activeCategory="Todos"}else{selectedCategories=selectedCategories.includes(cat)?selectedCategories.filter(x=>x!==cat):[...selectedCategories,cat];if(!selectedCategories.length){categoryMultiMode=false;activeCategory="Viajes HYM y URBAN"}}}else{activeCategory=cat;selectedCategories=[]}updateCategoryTabs();render()})});$("clearCategorySelection").addEventListener("click",()=>{categoryMultiMode=false;selectedCategories=[];activeCategory="Viajes HYM y URBAN";updateCategoryTabs();render()});updateCategoryTabs();$("statusFilter").onchange=render;$("noveltyFilter").onchange=render;$("importBtn").onclick=prepareImport;$("confirmImportBtn").onclick=confirmImport;$("cancelImportBtn").onclick=cancelImport;$("previewClose").onclick=cancelImport;$("downloadBtn").onclick=downloadExcel;$("beesCompareBtn").onclick=compareBeesCsv;$("beesClearBtn").onclick=clearBeesReport;$("porteriaUploadBtn").onclick=loadPorteriaExcel;$("porteriaClearBtn").onclick=clearPorteria;$("addUser").onclick=addUser;$("saveModulePermissions").onclick=saveModulePermissionEdits;$("cancelModulePermissions").onclick=()=>$("modulePermissionEditor").classList.add("hidden");$("newRole").onchange=updateRoleOptions;$("deleteAllBtn").onclick=deleteAllTransports;$("cancelDeleteBtn").onclick=closeDeleteModal;$("confirmDeleteBtn").onclick=confirmDeleteAll;$("deleteConfirmModal").addEventListener("click",e=>{if(e.target.id==="deleteConfirmModal")closeDeleteModal()});document.querySelectorAll(".statCard").forEach(card=>{card.addEventListener("click",()=>openStatChart(card.dataset.stat));card.addEventListener("keydown",e=>{if(e.key==="Enter"||e.key===" "){e.preventDefault();openStatChart(card.dataset.stat)}})});$("statChartClose").onclick=closeStatChart;$("statChartModal").addEventListener("click",e=>{if(e.target.id==="statChartModal")closeStatChart()});document.addEventListener("keydown",e=>{if(e.key==="Escape")closeStatChart()})}
async function login(){$("loginMsg").textContent="Verificando…";const pin=$("pin").value.trim();const{data,error}=await sb.rpc("login_by_pin",{p_pin:pin});if(error||!data?.length){$("loginMsg").textContent="PIN incorrecto o usuario desactivado.";return}user=data[0];loginPin=pin;const pr=await sb.rpc("get_user_permissions_by_pin",{p_pin:loginPin});if(pr.error||!pr.data?.length){$("loginMsg").textContent="No se pudieron verificar los permisos. Ejecuta primero la actualización SQL.";return}permissions=pr.data[0];const dp=await sb.rpc("get_user_all_permissions_by_pin",{p_pin:loginPin});if(dp.error||!dp.data?.length){$("loginMsg").textContent="No se pudieron verificar los permisos nuevos. Ejecuta el SQL de permisos adicionales.";return}permissions={...permissions,...dp.data[0]};sessionStorage.setItem("transport_user",JSON.stringify(user));await loadModuleRights();await showApp()}
const DEFAULT_MODULE_RIGHTS={cargo_view:true,cargo_edit:true,checkin_view:true,checkin_edit:true,rutero_view:true,rutero_edit:true};
let moduleRights={...DEFAULT_MODULE_RIGHTS},selectedPermissionUser=null;
function canModule(module,level){return permissions.is_admin||!!moduleRights[module+"_"+level]}
async function loadModuleRights(){if(permissions.is_admin){moduleRights={cargo_view:true,cargo_edit:true,checkin_view:true,checkin_edit:true,rutero_view:true,rutero_edit:true};return}const{data,error}=await sb.rpc("module_permissions_for_pin",{p_pin:loginPin});if(error){moduleRights={cargo_view:false,cargo_edit:false,checkin_view:false,checkin_edit:false,rutero_view:false,rutero_edit:false};toast("Falta ejecutar SQL de permisos v5.0: "+error.message);return}moduleRights={...DEFAULT_MODULE_RIGHTS,...(data?.[0]||{})};}
function applyModuleRights(){for(const [key,id] of [["cargo","cargoPrimaryBtn"],["checkin","checkinPrimaryBtn"],["rutero","brujulaPrimaryBtn"]])$(id).classList.toggle("hidden",!canModule(key,"view"));$("syncPrimaryBtn").classList.toggle("hidden",!canModule("cargo","view"));$("noModuleAccess").classList.toggle("hidden",["cargo","checkin","rutero"].some(x=>canModule(x,"view")));$("uploadPanel").classList.toggle("hidden",!(canModule("cargo","edit")&&(permissions.can_upload||permissions.is_admin)));$("importBtn").disabled=!canModule("cargo","edit");$("beesCompareBtn").disabled=!canModule("checkin","edit");$("porteriaUploadBtn").disabled=!canModule("checkin","edit");$("brujulaUploadBtn").disabled=!canModule("rutero","edit");if(!canModule("checkin","edit")){ $("beesClearBtn").disabled=true;$("porteriaClearBtn").disabled=true;}}
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
async function showApp(){$("login").classList.add("hidden");$("app").classList.remove("hidden");$("logout").classList.remove("hidden");$("userName").textContent=user.name;$("connection").textContent="● Conectado";if(permissions.can_upload||permissions.is_admin)$("uploadPanel").classList.remove("hidden");if(permissions.can_create_users||permissions.is_admin)$("usersPanel").classList.remove("hidden");if(permissions.can_delete_all_transports||permissions.is_admin)$("deletePanel").classList.remove("hidden");if(permissions.is_admin||permissions.can_create_users){$("checkinPermissionLabel").classList.remove("hidden")}$("historyPanel").classList.toggle("hidden",!permissions.is_admin);setupAdminModuleToggles();setupCheckinCollapses();applyModuleRights();if(permissions.is_admin){$("permissionLabel").classList.remove("hidden");$("createPermissionLabel").classList.remove("hidden");$("deletePermissionLabel").classList.remove("hidden");$("statusPermissionLabels").classList.remove("hidden")}updateRoleOptions();await refresh();await refreshCheckin();await loadSharedCheckin();subscribe();if(permissions.is_admin)loadUsers();if(!canModule("cargo","view"))switchModule(canModule("checkin","view")?"checkin":"brujula");else playCargoSplash()}
function logout(){loginPin="";sessionStorage.clear();location.reload()}
function updateRoleOptions(){const admin=permissions.is_admin;const opt=$("newRole").querySelector('option[value="Administrador"]');if(opt)opt.disabled=!admin;if(!admin&&$("newRole").value==="Administrador")$("newRole").value="Operativo";if(!admin){$("newCanUpload").checked=false;$("newCanCreateUsers").checked=false;$("newCanDeleteAll").checked=false;["newCanArm","newCanCount","newCanLoad","newCanAvailability","newCanPhysicalPlate"].forEach(id=>$(id).checked=false)}}
async function refresh(){const{data,error}=await sb.rpc("list_transports_with_category");if(error){toast(error.message);return}records=data||[];const extra=await sb.rpc("list_transport_vehicle_fields");if(!extra.error){const fields=new Map((extra.data||[]).map(x=>[String(x.id),x]));records=records.map(r=>({...r,...(fields.get(String(r.id))||{})}))}window.syncCargoRecords = records;
window.dispatchEvent(new CustomEvent('cargo-records-updated',{detail:{count:records.length}}));
render();if(permissions.is_admin)loadHistory()}
function render(){const categoryRecords=activeCategory==="Todos"?records:(categoryMultiMode&&selectedCategories.length?records.filter(r=>selectedCategories.includes(r.categoria||"Viajes Tradicional")):records.filter(r=>(r.categoria||"Viajes Tradicional")===activeCategory));const q=($("search").value||"").toLowerCase(),sf=$("statusFilter").value,nf=$("noveltyFilter").value;const list=categoryRecords.filter(r=>{const text=[r.transporte,r.viaje,r.placa,r.novedad,r.estado_vehiculo,r.cambio_placa].join(" ").toLowerCase(),has=String(r.novedad||"").trim()!=="";return(!sf||r.estado===sf)&&(!nf||(nf==="con"?has:!has))&&(!q||text.includes(q))});window.currentCargoFilteredRecords=list;window.syncCargoRecords=records;window.renderSyncReport?.();updatePorteriaPendingCount();$("total").textContent=list.length;$("pendientes").textContent=list.filter(r=>r.estado==="Pendiente").length;$("armados").textContent=list.filter(r=>r.estado==="Armado").length;$("contados").textContent=list.filter(r=>r.estado==="Contado").length;$("cargados").textContent=list.filter(r=>r.estado==="Cargado").length;$("novedades").textContent=list.filter(r=>String(r.novedad||"").trim()!=="").length;
const total=list.length||1,novedadCount=list.filter(r=>String(r.novedad||"").trim()!=="").length,armadoCount=list.filter(r=>r.estado==="Armado").length,contadoCount=list.filter(r=>r.estado==="Contado").length,cargadoCount=list.filter(r=>r.estado==="Cargado").length;const pct=n=>Math.round(n*100/total);$("progressNovedadCount").textContent=novedadCount;$("progressNovedadPct").textContent=pct(novedadCount)+"%";$("progressNovedadBar").style.width=pct(novedadCount)+"%";$("progressArmadoCount").textContent=armadoCount;$("progressArmadoPct").textContent=pct(armadoCount)+"%";$("progressArmadoBar").style.width=pct(armadoCount)+"%";$("progressContadoCount").textContent=contadoCount;$("progressContadoPct").textContent=pct(contadoCount)+"%";$("progressContadoBar").style.width=pct(contadoCount)+"%";$("progressCargadoCount").textContent=cargadoCount;$("progressCargadoPct").textContent=pct(cargadoCount)+"%";$("progressCargadoBar").style.width=pct(cargadoCount)+"%";
$("rows").innerHTML=list.map(r=>{const has=String(r.novedad||"").trim()!=="";const ev=r.estado_vehiculo||"";const selected=String(selectedTransportId)===String(r.id);const canAvailability=canModule("cargo","edit")&&(permissions.is_admin||permissions.can_edit_availability);const canPlate=canModule("cargo","edit")&&(permissions.is_admin||permissions.can_edit_physical_plate);const canAction=canModule("cargo","edit")&&(permissions.is_admin||(r.estado==="Pendiente"?permissions.can_arm:r.estado==="Armado"?permissions.can_count:r.estado==="Contado"?permissions.can_load:false));return `<tr class="${has?"hasNovelty ":""}${selected?"transportSelected":""}"><td><button type="button" class="${selected?"selectedTransportBtn":"secondary selectTransportBtn"}" onclick="selectTransport(${r.id})">${selected?"✓ Seleccionado":"Seleccionar"}</button></td><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td><b>${esc(r.placa)}</b></td><td>${has?`<span class="novelty">${esc(r.novedad)}</span>`:"—"}</td><td><select class="vehicleState ${vehicleClass(ev)}" aria-label="Disponibilidad De Vehiculos para ${esc(r.placa)}" ${(!selected||!canAvailability)?"disabled":""} onchange="saveVehicleField(${r.id},'estado_vehiculo',this.value)"><option value="">Seleccionar...</option><option ${ev==="Bodega"?"selected":""}>Bodega</option><option ${ev==="Sobredimensionado"?"selected":""}>Sobredimensionado</option><option ${ev==="OK para utilizar"?"selected":""}>OK para utilizar</option></select></td><td><input class="plateChange" aria-label="Placa Fisica para ${esc(r.placa)}" ${(!selected||!canPlate)?"disabled":""} maxlength="8" value="${esc(r.cambio_placa||"")}" placeholder="Máx. 8 caracteres" onchange="this.value=normalizePhysicalPlate(this.value);saveVehicleField(${r.id},'cambio_placa',this.value)" oninput="this.value=normalizePhysicalPlate(this.value)"></td><td><span class="badge ${r.estado}">${esc(r.estado)}</span></td><td><button onclick="advance(${r.id})" ${(!selected||!canAction||r.estado==="Cargado")?"disabled":""}>${r.estado==="Pendiente"?"Armar":r.estado==="Armado"?"Contar":"Cargar"}</button></td><td><span class="porteriaBadge ${getPorteriaStatus(r.transporte)==="Salida"?"porteriaYes":getPorteriaStatus(r.transporte)==="Sin salida"?"porteriaNo":"porteriaUnknown"}">${getPorteriaStatus(r.transporte)}</span></td></tr>`}).join("")||'<tr><td colspan="10">No hay registros con estos filtros.</td></tr>';const selectedRecord=records.find(r=>String(r.id)===String(selectedTransportId));if(selectedTransportId&&!selectedRecord)selectedTransportId=null}
function selectTransport(id){selectedTransportId=String(selectedTransportId)===String(id)?null:String(id);const r=records.find(x=>String(x.id)===String(id));toast(selectedTransportId?`Transporte seleccionado: ${r?.transporte||""} · Viaje ${r?.viaje||""}. Ya puedes modificar los campos permitidos.`:"Selección cancelada.");render()}
function openStatChart(key){const visible=window.currentCargoFilteredRecords||[];const total=visible.length;const config={total:{label:"Total",rows:visible,color:"#1769aa"},pendientes:{label:"Pendientes",rows:visible.filter(r=>r.estado==="Pendiente"),color:"#64748b"},armados:{label:"Armados",rows:visible.filter(r=>r.estado==="Armado"),color:"#f97316"},contados:{label:"Contados",rows:visible.filter(r=>r.estado==="Contado"),color:"#2563eb"},cargados:{label:"Cargados",rows:visible.filter(r=>r.estado==="Cargado"),color:"#22c55e"},novedades:{label:"Con novedad",rows:visible.filter(r=>String(r.novedad||"").trim()!==""),color:"#ef4444"}};const item=config[key];if(!item)return;const count=item.rows.length,percent=total?Math.round(count*100/total):0;$("statChartTitle").textContent="Avance: "+item.label;$("statChartCount").textContent=count.toLocaleString("es-CO");$("statChartPercent").textContent=percent+"% de los registros filtrados";$("statChartFill").style.width=percent+"%";$("statChartFill").style.background=item.color;$("statChartTruck").style.left=`calc(${Math.min(percent,96)}% - 4px)`;$("statChartTruck").setAttribute("aria-label",`Camión en ${percent} por ciento`);$("statChartTotal").textContent=`Total filtrado: ${total.toLocaleString("es-CO")} transportes`;$("statChartVehicleHeading").textContent=`Transportes en categoría: ${count}`;$("statChartVehicles").innerHTML=count?item.rows.map(r=>`<div class="statVehicleRow"><span class="statVehicleName">${esc(String(r.transporte??"Sin nombre").trim())} - ${esc(String(r.placa??"Sin placa").trim())}</span></div>`).join(""):'<p class="statVehicleEmpty">No hay transportes en esta categoría.</p>';$("statChartModal").classList.remove("hidden");$("statChartClose").focus()}function closeStatChart(){$("statChartModal")?.classList.add("hidden")}
function normalizePhysicalPlate(value){let v=String(value||"").toLocaleUpperCase("es").replace(/\s+/g,"");if(v&&!v.startsWith("CO"))v="CO"+v;return v.slice(0,8)}
function vehicleClass(v){return v==="Bodega"?"vehicle-bodega":v==="Sobredimensionado"?"vehicle-sobredimensionado":v==="OK para utilizar"?"vehicle-ok":""}
async function saveVehicleField(id,field,value){if(!canModule("cargo","edit")){toast("Sin permiso de modificación en el módulo.");return}if(field==="cambio_placa")value=normalizePhysicalPlate(value);if(field==="cambio_placa"&&String(value).length>8){toast("Placa Fisica: máximo 8 caracteres.");await refresh();return}const{error}=await sb.rpc("update_transport_vehicle_field_secure",{p_transport_id:id,p_field:field,p_value:value,p_pin:loginPin});if(error){toast("No se pudo guardar: "+error.message);await refresh();return}const r=records.find(x=>Number(x.id)===Number(id));if(r)r[field]=value;toast(field==="estado_vehiculo"?"Disponibilidad De Vehiculos guardada.":"Placa Fisica guardada.")}
async function advance(id){if(!canModule("cargo","edit")){toast("Sin permiso de modificación en el módulo.");return}const r=records.find(x=>Number(x.id)===Number(id));if(!r)return;const next=r.estado==="Pendiente"?"Armado":r.estado==="Armado"?"Contado":r.estado==="Contado"?"Cargado":null;if(!next)return;if(!confirm(`¿Confirmas el cambio de estado?\n\nPlaca: ${r.placa}\nEstado: ${r.estado} → ${next}`))return;const{error}=await sb.rpc("advance_transport_secure",{p_transport_id:id,p_pin:loginPin});if(error){toast(error.message);return}toast(`Estado actualizado: ${r.estado} → ${next}`);await refresh()}
async function loadHistory(){const{data,error}=await sb.rpc("recent_history",{p_limit:30});if(error)return;$("history").innerHTML=(data||[]).map(h=>`<div class="historyItem"><b>${esc(h.usuario)}</b> · ${esc(h.placa)} · Viaje ${esc(h.viaje)}<br>${h.estado_anterior} → <b>${h.estado_nuevo}</b> · ${new Date(h.created_at).toLocaleString("es-CO")}</div>`).join("")||"Sin movimientos."}
async function loadUsers(){const{data,error}=await sb.rpc("admin_users_with_all_permissions",{p_admin_pin:loginPin});if(error){toast(error.message);return}$("users").innerHTML=(data||[]).map(u=>`<div class="userRow"><span><b>${esc(u.name)}</b> · C.C. ${esc(u.cedula||"No registrada")} · ${esc(u.role|| (u.is_admin?"Administrador":"Operativo"))} · ${u.active?"Activo":"Inactivo"}<br><small>Armado: ${u.can_arm?"Sí":"No"} · Contado: ${u.can_count?"Sí":"No"} · Cargado: ${u.can_load?"Sí":"No"}<br>Disponibilidad: ${u.can_edit_availability?"Sí":"No"} · Placa física: ${u.can_edit_physical_plate?"Sí":"No"} · Modificar Checkin: ${u.can_manage_checkin?"Sí":"No"}<br>Carga: ${u.can_upload?"Sí":"No"} · Crear usuarios: ${u.can_create_users?"Sí":"No"} · Eliminar todos: ${u.can_delete_all_transports?"Sí":"No"}</small></span><span>PIN: ••••</span>${permissions.is_admin?`<button onclick="toggleUser('${u.id}',${u.active})" class="secondary">${u.active?"Desactivar":"Activar"}</button>`:""}${permissions.is_admin&&!u.is_admin?`<button class="secondary" onclick="setUserPermissions('${u.id}')">Editar permisos</button><button class="secondary" onclick="editModulePermissions('${u.id}','${esc(u.name).replaceAll("&#39;", "")}')">Acceso a pestañas</button>`:""}</div>`).join("")}
async function addUser(){if(!permissions.can_create_users&&!permissions.is_admin){toast("No tienes permiso para crear usuarios.");return}const name=$("newName").value.trim(),cedula=$("newCedula").value.trim(),pin=$("newPin").value.trim(),role=$("newRole").value;if(!name||!cedula||!pin){toast("Nombre, cédula y PIN son obligatorios");return}if(role==="Administrador"&&!permissions.is_admin){toast("Solo un administrador puede crear administradores.");return}const{data:newUserId,error}=await sb.rpc("admin_create_user_full_secure",{p_name:name,p_cedula:cedula,p_pin:pin,p_creator_pin:loginPin,p_role:role,p_can_upload:permissions.is_admin&&$("newCanUpload").checked,p_can_create_users:permissions.is_admin&&$("newCanCreateUsers").checked,p_can_delete_all_transports:permissions.is_admin&&$("newCanDeleteAll").checked,p_can_arm:permissions.is_admin&&$("newCanArm").checked,p_can_count:permissions.is_admin&&$("newCanCount").checked,p_can_load:permissions.is_admin&&$("newCanLoad").checked,p_can_edit_availability:permissions.is_admin&&$("newCanAvailability").checked,p_can_edit_physical_plate:permissions.is_admin&&$("newCanPhysicalPlate").checked,p_can_manage_checkin:(permissions.is_admin||permissions.can_create_users)&&$("newCanManageCheckin").checked});if(error){toast(error.message);return}$("newName").value="";$("newCedula").value="";$("newPin").value="";$("newRole").value="Operativo";["newCanUpload","newCanCreateUsers","newCanDeleteAll","newCanArm","newCanCount","newCanLoad","newCanAvailability","newCanPhysicalPlate","newCanManageCheckin"].forEach(id=>$(id).checked=false);if(permissions.is_admin&&newUserId){const rights={};for(const mod of ["Cargo","Checkin","Rutero"])for(const mode of ["View","Edit"])rights[mod.toLowerCase()+"_"+mode.toLowerCase()]=$("new"+mod+mode).checked;const res=await sb.rpc("admin_set_module_permissions",{p_admin_pin:loginPin,p_user_id:newUserId,p_rights:rights});if(res.error){toast("Usuario creado, pero faltan permisos de módulos: "+res.error.message);return}}if(permissions.is_admin)await loadUsers();toast("Usuario creado correctamente")}
async function setUserPermissions(id){const{data,error}=await sb.rpc("admin_users_with_all_permissions",{p_admin_pin:loginPin});if(error){toast(error.message);return}const u=(data||[]).find(x=>x.id===id);if(!u)return;const role=prompt("Rol: Administrador, Operativo o Supervisor",u.role||"Operativo");if(role===null)return;if(!["Administrador","Operativo","Supervisor"].includes(role.trim())){toast("Rol no válido");return}const ask=(label,value)=>confirm(label+"\nAceptar = Sí / Cancelar = No");const canUpload=ask("¿Puede cargar/descargar archivos?",u.can_upload),canCreate=ask("¿Puede crear usuarios?",u.can_create_users),canDelete=ask("¿Puede eliminar todos los transportes?",u.can_delete_all_transports),canArm=ask("¿Puede cambiar a estado Armado?",u.can_arm),canCount=ask("¿Puede cambiar a estado Contado?",u.can_count),canLoad=ask("¿Puede cambiar a estado Cargado?",u.can_load),canAvail=ask("¿Puede editar Disponibilidad de Vehículos?",u.can_edit_availability),canPlate=ask("¿Puede editar Placa Física?",u.can_edit_physical_plate);const{error:saveError}=await sb.rpc("admin_set_user_all_permissions_secure",{p_admin_pin:loginPin,p_user_id:id,p_role:role.trim(),p_can_upload:canUpload,p_can_create_users:canCreate,p_can_delete_all_transports:canDelete,p_can_arm:canArm,p_can_count:canCount,p_can_load:canLoad,p_can_edit_availability:canAvail,p_can_edit_physical_plate:canPlate,p_can_manage_checkin:ask("¿Puede modificar datos en Pendientes Por Checkin?",u.can_manage_checkin)});if(saveError){toast(saveError.message);return}await loadUsers();toast("Permisos actualizados")}
async function toggleUser(id,active){const{error}=await sb.rpc("admin_set_user_active_secure",{p_admin_pin:loginPin,p_user_id:id,p_active:!active});if(error){toast(error.message);return}loadUsers()}
function norm(s){return String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s/g,"")}
async function prepareImport(){if(!canModule("cargo","edit")){toast("Sin permiso de modificación en el módulo.");return}if(!permissions.can_upload&&!permissions.is_admin){toast("No tienes permiso para cargar archivos.");return}const f=$("excelFile").files[0];if(!f){toast("Selecciona un archivo");return}const buf=await f.arrayBuffer(),wb=XLSX.read(buf,{type:"array"});const required=["Viajes HYM y URBAN","Viajes KA","Viajes Tradicional"];const optional=["Viajes 2KA","Viajes 2","Custodias"];const expected=[...required];const sheetFor=label=>wb.SheetNames.find(n=>norm(n)===norm(label));const missing=required.filter(n=>!sheetFor(n));if(missing.length){toast("Faltan pestañas en el Excel: "+missing.join(", "));return}for(const category of optional){if(sheetFor(category))expected.push(category)}const tapaSheet=sheetFor("Tapa Códigos");if(!tapaSheet){toast("Falta la pestaña Tapa Códigos en el Excel.");return}const tapaData=XLSX.utils.sheet_to_json(wb.Sheets[tapaSheet],{defval:""});if(tapaData.length){const tk=Object.keys(tapaData[0]),findT=n=>tk.find(k=>norm(k)===norm(n));const tTrans=findT("transporte"),tVeh=findT("vehiculo"),tViaje=findT("viaje"),tMat=findT("material"),tDen=findT("denominacion"),tCant=findT("cantidad"),tFam=findT("familia"),tDesc=findT("descripcion");if(!tTrans||!tVeh||!tViaje||!tMat||!tDen||!tCant||!tFam||!tDesc){toast("Tapa Códigos debe tener FAMILIA, TRANSPORTE, VEHICULO, VIAJE, MATERIAL, DENOMINACION, CANTIDAD y DESCRIPCION.");return}pendingTapaRows=tapaData.map(x=>({familia:String(x[tFam]??"").trim(),descripcion:String(x[tDesc]??"").trim(),transporte:String(x[tTrans]??"").trim(),vehiculo:String(x[tVeh]??"").trim(),viaje:String(x[tViaje]??"").trim(),material:String(x[tMat]??"").trim(),denominacion:String(x[tDen]??"").trim(),cantidad:Number(String(x[tCant]??"0").replace(/,/g,""))||0})).filter(x=>x.familia||x.descripcion||x.transporte||x.vehiculo||x.viaje||x.material||x.denominacion)}else{pendingTapaRows=[]}let rows=[];for(const category of expected){const data=XLSX.utils.sheet_to_json(wb.Sheets[sheetFor(category)],{defval:""});if(!data.length)continue;const keys=Object.keys(data[0]),find=n=>keys.find(k=>norm(k)===n),kt=find("transporte"),kv=find("viaje"),kp=find("placa"),kn=find("novedad");if(!kt||!kv||!kp||!kn){toast("La pestaña "+category+" debe contener Transporte, Viaje, Placa y Novedad");return}rows.push(...data.map(x=>({transporte:String(x[kt]??"").trim(),viaje:String(x[kv]??"").trim(),placa:String(x[kp]??"").trim(),novedad:String(x[kn]??"").trim(),categoria:category})).filter(x=>x.transporte||x.viaje||x.placa))}if(!rows.length){toast("Las pestañas de viajes están vacías");return}const seen=new Set();let duplicates=0;const unique=rows.filter(x=>{const k=[x.categoria,x.transporte,x.viaje,x.placa].join("|").toLowerCase();if(seen.has(k)){duplicates++;return false}seen.add(k);return true});const existingKeys=new Set(records.map(r=>[r.categoria||"Viajes Tradicional",r.transporte,r.viaje,r.placa].map(v=>String(v??"").trim()).join("|").toLowerCase()));const existing=unique.filter(x=>existingKeys.has([x.categoria,x.transporte,x.viaje,x.placa].map(v=>String(v??"").trim()).join("|").toLowerCase())).length;pendingImportRows=unique;const importMsg=$("importResultMessage");if(importMsg){importMsg.textContent="";importMsg.classList.add("hidden");}$("previewText").innerHTML=`<b>${rows.length}</b> filas leídas entre las pestañas de viajes y <b>${pendingTapaRows.length}</b> filas de Tapa Códigos.<br>✅ <b>${unique.length-existing}</b> nuevas para agregar.<br>🔄 <b>${existing}</b> ya existen: se actualizará su Novedad y se conservarán estado, placa física y disponibilidad.<br>♻️ <b>${duplicates}</b> duplicadas dentro del Excel no se agregarán.`;$("previewModal").classList.remove("hidden")}
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
    console.info('[Importación] Confirmación Supabase',result);
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
    try{await refresh();}catch(refreshError){display(`Supabase confirmó la carga (${added} nuevos), pero falló la actualización de pantalla: ${refreshError.message}. Vuelve a abrir Cargue y Despacho; NO repitas el Excel.`,true);return;}
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
const{data,error}=await sb.rpc("list_transports_with_category");
if(error){toast(error.message);return}
const fields=await sb.rpc("list_transport_vehicle_fields");
const map=new Map((fields.data||[]).map(x=>[String(x.id),x]));
const wb=XLSX.utils.book_new();
const categories=["Viajes HYM y URBAN","Viajes KA","Viajes Tradicional","Viajes 2KA","Viajes 2","Custodias"];
const allRows=(data||[]).map(r=>{const x=map.get(String(r.id))||{};return {categoria:r.categoria||"Viajes Tradicional",Transporte:r.transporte||"",Viaje:r.viaje||"",Placa:r.placa||"", "Placa física":x.cambio_placa||"",Estado:r.estado||"Pendiente",Clientes:""}});
const visible=categoryMultiMode&&selectedCategories.length?allRows.filter(r=>selectedCategories.includes(r.categoria)):activeCategory==="Todos"?allRows:allRows.filter(r=>r.categoria===activeCategory);
const selected=visible.length?visible:allRows;
const dateText=new Date().toLocaleDateString("es-CO",{year:"numeric",month:"long",day:"numeric"});
const navy="123047",blue="1769AA",pale="EAF1F7",white="FFFFFF",gray="5B6573";
function styleSheet(ws,{titleRow=1,headerRow=3,lastCol=7,autoFilter=true}={}){
const range=XLSX.utils.decode_range(ws["!ref"]||"A1:A1");
ws["!cols"]=[{wch:20},{wch:16},{wch:16},{wch:20},{wch:18},{wch:32}];
ws["!freeze"]={xSplit:0,ySplit:headerRow};
ws["!autofilter"]=autoFilter?{ref:XLSX.utils.encode_range({s:{r:headerRow-1,c:0},e:{r:range.e.r,c:lastCol}})}:undefined;
ws["!pageSetup"]={orientation:"landscape",paperSize:9,fitToWidth:1,fitToHeight:0,scale:0};
ws["!printOptions"]={horizontalCentered:true};
ws["!margins"]={left:0.25,right:0.25,top:0.5,bottom:0.5,header:0.2,footer:0.2};
ws["!pageSetup"].fitToPage=true;
for(let c=0;c<=lastCol;c++){
 const h=ws[XLSX.utils.encode_cell({r:headerRow-1,c})];
 if(h)h.s={fill:{fgColor:{rgb:blue}},font:{bold:true,color:{rgb:white},name:"Aptos",sz:10},alignment:{horizontal:"center",vertical:"center",wrapText:true},border:{bottom:{style:"medium",color:{rgb:navy}}}};
}
ws["!rows"]=ws["!rows"]||[];
ws["!rows"][0]={hpt:28};ws["!rows"][headerRow-1]={hpt:30};
for(let r=headerRow;r<=range.e.r;r++){
 ws["!rows"][r]={hpt:22};
 for(let c=0;c<=lastCol;c++){const cell=ws[XLSX.utils.encode_cell({r,c})];if(cell)cell.s={font:{name:"Aptos",sz:10,color:{rgb:"243447"}},fill:{fgColor:{rgb:r%2===0?white:pale}},alignment:{vertical:"center",wrapText:true},border:{bottom:{style:"thin",color:{rgb:"D8E1EA"}}}};}
}
}
const summary=[["CONTROL DE TRANSPORTES"],["REPORTE OPERATIVO PARA IMPRESIÓN"],["Generado el",dateText],["Categoría",categoryMultiMode?selectedCategories.join(" + "):activeCategory],[],["RESUMEN POR CATEGORÍA","CANTIDAD"]];
categories.forEach(c=>summary.push([c,allRows.filter(r=>r.categoria===c).length]));
summary.push(["TOTAL GENERAL",allRows.length]);
const sws=XLSX.utils.aoa_to_sheet(summary);
sws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:1}},{s:{r:1,c:0},e:{r:1,c:1}}];
sws["!cols"]=[{wch:36},{wch:24}];sws["!rows"]=[{hpt:30},{hpt:22}];
for(const addr of ["A1","A2"]){if(sws[addr])sws[addr].s={fill:{fgColor:{rgb:navy}},font:{name:"Aptos Display",bold:true,color:{rgb:white},sz:addr==="A1"?18:12},alignment:{vertical:"center"}}}
["A6","B6"].forEach(a=>{if(sws[a])sws[a].s={fill:{fgColor:{rgb:blue}},font:{bold:true,color:{rgb:white},name:"Aptos",sz:10},alignment:{horizontal:"center"}}});
for(let r=6;r<summary.length;r++)for(let c=0;c<2;c++){const cell=sws[XLSX.utils.encode_cell({r,c})];if(cell)cell.s={font:{name:"Aptos",sz:11,bold:r===summary.length-1},fill:{fgColor:{rgb:r%2===0?pale:white}},border:{bottom:{style:"thin",color:{rgb:"D8E1EA"}}}}}
sws["!pageSetup"]={orientation:"portrait",paperSize:9,fitToWidth:1,fitToHeight:1,scale:0};sws["!pageSetup"].fitToPage=true;sws["!margins"]={left:0.5,right:0.5,top:0.6,bottom:0.6,header:0.2,footer:0.2};
XLSX.utils.book_append_sheet(wb,sws,"Resumen");
const reportRows=rows=>rows.map(({categoria,...r})=>r);
for(const category of categories){
 const categoryRows=allRows.filter(r=>r.categoria===category);
 if(!categoryRows.length)continue;
 const dataRows=reportRows(categoryRows);
 const ws=XLSX.utils.aoa_to_sheet([["CONTROL DE TRANSPORTES"],[category.toUpperCase()+" · REPORTE PARA ENTREGA"],["Fecha de generación",dateText],[],Object.keys(dataRows[0]||{Transporte:"",Viaje:"",Placa:"",Novedad:"", "Disponibilidad de vehículos":"","Placa física":"","Estado":""}),...dataRows.map(r=>Object.values(r))]);
 ws["!merges"]=[{s:{r:0,c:0},e:{r:0,c:5}},{s:{r:1,c:0},e:{r:1,c:5}}];
 styleSheet(ws,{headerRow:5,lastCol:5});
 for(const addr of ["A1","A2"]){if(ws[addr])ws[addr].s={fill:{fgColor:{rgb:navy}},font:{name:"Aptos Display",bold:true,color:{rgb:white},sz:addr==="A1"?18:12},alignment:{vertical:"center"}}}
 if(ws["A3"])ws["A3"].s={font:{name:"Aptos",color:{rgb:gray},italic:true,sz:9}};
 XLSX.utils.book_append_sheet(wb,ws,category.slice(0,31));
}
const chosenRows=reportRows(selected);
const combined=XLSX.utils.aoa_to_sheet([["CONTROL DE TRANSPORTES"],["REPORTE CONSOLIDADO PARA IMPRESIÓN"],["Fecha de generación",dateText],["Categorías",categoryMultiMode?selectedCategories.join(", "):activeCategory],[],Object.keys(chosenRows[0]||{Transporte:"",Viaje:"",Placa:"",Novedad:"", "Disponibilidad de vehículos":"","Placa física":"","Estado":""}),...chosenRows.map(r=>Object.values(r))]);
combined["!merges"]=[{s:{r:0,c:0},e:{r:0,c:5}},{s:{r:1,c:0},e:{r:1,c:5}}];
styleSheet(combined,{headerRow:6,lastCol:5});
for(const addr of ["A1","A2"]){if(combined[addr])combined[addr].s={fill:{fgColor:{rgb:navy}},font:{name:"Aptos Display",bold:true,color:{rgb:white},sz:addr==="A1"?18:12},alignment:{vertical:"center"}}}
if(combined["A3"])combined["A3"].s={font:{name:"Aptos",color:{rgb:gray},italic:true,sz:9}};
XLSX.utils.book_append_sheet(wb,combined,"Reporte imprimible");
XLSX.writeFile(wb,"Control_Transportes_Reporte_Ejecutivo.xlsx",{cellStyles:true});
}
function subscribe(){sb.channel("transportes-live").on("postgres_changes",{event:"*",schema:"public",table:"transportes"},()=>refresh()).subscribe()}
function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]))}
function switchModule(module){const target=(module==="main"||module==="tapa")?"cargo":module==="brujula"?"rutero":module==="sync"?"cargo":"checkin";if(!canModule(target,"view")){toast("Sin permiso para acceder a esta pestaña.");return}const previousTop=(activeModule==="main"||activeModule==="tapa")?"cargo":activeModule==="brujula"?"rutero":activeModule==="sync"?"cargo":"checkin";
if(previousTop!==target && module!=="tapa")playModuleSplash(target);
activeModule=module;const cargo=module==="main"||module==="tapa",main=module==="main",checkin=module==="checkin",tapa=module==="tapa",brujula=module==="brujula",sync=module==="sync";if(cargo)activeTopModule="cargo";else activeTopModule=module;$("mainModuleView").classList.toggle("hidden",!main);$("checkinModuleView").classList.toggle("hidden",!checkin);$("tapaModuleView").classList.toggle("hidden",!tapa);$("brujulaModuleView").classList.toggle("hidden",!brujula);$("syncModuleView").classList.toggle("hidden",!sync);$("cargoSubnav").classList.toggle("hidden",!cargo);$("cargoPrimaryBtn").classList.toggle("active",cargo);$("checkinPrimaryBtn").classList.toggle("active",checkin);$("brujulaPrimaryBtn").classList.toggle("active",brujula);$("syncPrimaryBtn").classList.toggle("active",sync);$("cargoPrimaryBtn").setAttribute("aria-selected",String(cargo));$("checkinPrimaryBtn").setAttribute("aria-selected",String(checkin));$("brujulaPrimaryBtn").setAttribute("aria-selected",String(brujula));$("syncPrimaryBtn").setAttribute("aria-selected",String(sync));$("mainModuleBtn").classList.toggle("active",main);$("tapaModuleBtn").classList.toggle("active",tapa);$("mainModuleBtn").setAttribute("aria-selected",String(main));$("tapaModuleBtn").setAttribute("aria-selected",String(tapa));if(checkin)refreshCheckin();if(checkin)refreshCheckin();if(tapa)refreshTapa();if(brujula)loadBrujulaReport();if(sync){window.renderSyncReport?.();window.loadSyncStatReport?.();if(!records.length&&sb&&loginPin)void refresh()}}
let checkinRecords=[];
async function refreshTapa(){if(!sb||!loginPin)return;const{data,error}=await sb.rpc("list_tapa_codigos_ka_secure",{p_pin:loginPin});if(error){toast("No se pudo cargar Tapa Códigos KA. Ejecuta el SQL incluido: "+error.message);return}tapaRecords=data||[];await fetchTapaStates();const select=$("tapaFamilyFilter"),current=select.value,clientSelect=$("tapaClientFilter"),currentClient=clientSelect.value;const clients=[...new Set(tapaRecords.map(r=>String(r.descripcion||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));clientSelect.innerHTML='<option value="">Todos los clientes</option>'+clients.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join("");if(clients.includes(currentClient))clientSelect.value=currentClient;const families=[...new Set(tapaRecords.map(r=>String(r.familia||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));select.innerHTML='<option value="">Todas las familias</option>'+families.map(f=>`<option value="${esc(f)}">${esc(f)}</option>`).join("");if(families.includes(current))select.value=current;renderTapa()}
let tapaStatusMap=new Map();
async function fetchTapaStates(){if(!canModule("cargo","view"))return;const{data,error}=await sb.rpc("list_tapa_status_secure",{p_pin:loginPin});if(error){toast("No fue posible consultar estados de tapa códigos: "+error.message);return}tapaStatusMap=new Map((data||[]).map(x=>[x.group_key,x.status]));syncTapaStates()}
function syncTapaStates(){document.querySelectorAll(".tapaState").forEach(el=>{const key=decodeURIComponent(el.dataset.tapaKey);el.value=tapaStatusMap.get(key)||"";el.onchange=async()=>{if(!canModule("cargo","edit")){toast("No tienes permiso de edición");return}const previous=tapaStatusMap.get(key)||"";const{error}=await sb.rpc("set_tapa_status_secure",{p_pin:loginPin,p_group_key:key,p_status:el.value||null});if(error){el.value=previous;toast(error.message);return}tapaStatusMap.set(key,el.value);toast("Estado de Tapa Códigos guardado")}})}
function renderTapa(){const family=$("tapaFamilyFilter")?.value||"",client=$("tapaClientFilter")?.value||"",q=($("tapaSearch")?.value||"").trim().toLocaleLowerCase("es");const filtered=tapaRecords.filter(r=>(!family||String(r.familia||"")===family)&&(!client||String(r.descripcion||"")===client)&&(!q||[r.transporte,r.vehiculo,r.viaje,r.material,r.denominacion,r.descripcion].some(v=>String(v||"").toLocaleLowerCase("es").includes(q))));const groups=new Map();for(const r of filtered){const key=[r.transporte,r.vehiculo,r.viaje,r.material,r.denominacion].map(v=>String(v||"").trim()).join("\u001f");const item=groups.get(key)||{transporte:r.transporte||"",vehiculo:r.vehiculo||"",viaje:r.viaje||"",material:r.material||"",denominacion:r.denominacion||"",cantidad:0};item.cantidad+=Number(r.cantidad)||0;groups.set(key,item)}const rows=[...groups.values()].sort((a,b)=>[a.transporte,a.vehiculo,a.viaje,a.material,a.denominacion].join("|").localeCompare([b.transporte,b.vehiculo,b.viaje,b.material,b.denominacion].join("|"),"es"));$("tapaRows").innerHTML=rows.map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.vehiculo)}</td><td>${esc(r.viaje)}</td><td>${esc(r.material)}</td><td>${esc(r.denominacion)}</td><td>${(r.cantidad||0).toLocaleString("es-CO",{maximumFractionDigits:3})}</td><td><select class="tapaState" data-tapa-key="${encodeURIComponent([r.transporte,r.vehiculo,r.viaje,r.material,r.denominacion].map(v=>String(v||"").trim()).join("\u001f"))}" ${!canModule("cargo","edit")?"disabled":""}><option value="">Seleccionar</option><option value="Completo">Completo</option><option value="Parcial">Parcial</option><option value="Sin tapa códigos">Sin tapa códigos</option></select></td></tr>`).join("")||'<tr><td colspan="7">No hay datos para los filtros seleccionados.</td></tr>';$("tapaTotal").textContent=rows.reduce((s,r)=>s+r.cantidad,0).toLocaleString("es-CO",{maximumFractionDigits:3});$("tapaSummary").textContent=`${rows.length} combinaciones · ${filtered.length} registros de origen`;syncTapaStates();}

async function refreshCheckin(){if(!sb||!loginPin)return;const{data,error}=await sb.rpc("list_checkin_transports_secure",{p_pin:loginPin});if(error){toast("No se pudo cargar Checkin. Ejecuta el SQL de Checkin: "+error.message);return}checkinRecords=data||[];renderCheckin()}
function renderCheckin(){const rows=$("checkinRows");if(!rows)return;rows.innerHTML=checkinRecords.map(r=>{const sel=String(selectedCheckinId)===String(r.transport_id);const canEdit=canModule("checkin","edit")&&(permissions.is_admin||permissions.can_manage_checkin);const cls=r.checkin_estado==="Pendiente"?"status-pendiente":r.checkin_estado==="Realizado"?"status-realizado":"status-novedad";const beesStatus=getBeesStatus(r.transporte);const porteriaStatus=getPorteriaStatus(r.transporte);const beesClass=beesStatus==="REALIZADO"?"beesMatch":beesStatus==="N/A"?"beesMissing":"";return `<tr class="${sel?"checkinSelected":""}"><td><button type="button" class="${sel?"selectedTransportBtn":"secondary"}" onclick="selectCheckin(${r.transport_id})">${sel?"✓ Seleccionado":"Seleccionar"}</button></td><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td><b>${esc(r.placa)}</b></td><td><input class="checkinPhysicalPlate" value="${esc(r.placa_fisica||"")}" readonly aria-label="Placa Física" /></td><td><select class="checkinStatus ${cls}" ${(!sel||!canEdit)?"disabled":""} onchange="saveCheckin(${r.transport_id},this.value,null)"><option value="Pendiente" ${r.checkin_estado==="Pendiente"?"selected":""}>Pendiente</option><option value="Realizado" ${r.checkin_estado==="Realizado"?"selected":""}>Realizado</option><option value="Con novedad" ${r.checkin_estado==="Con novedad"?"selected":""}>Con novedad</option></select></td><td><input class="checkinNovelty" value="${esc(r.checkin_novedad||"")}" placeholder="Escribir novedad…" ${(!sel||!canEdit)?"disabled":""} onchange="saveCheckin(${r.transport_id},null,this.value)" /></td><td><span class="beesBadge ${beesClass}">${beesStatus}</span></td><td><span class="porteriaBadge ${porteriaStatus==="Salida"?"porteriaYes":porteriaStatus==="Sin salida"?"porteriaNo":"porteriaUnknown"}">${porteriaStatus}</span></td></tr>`}).join("")||'<tr><td colspan="9">No hay transportes en estado Cargado para check-in.</td></tr>'}function selectCheckin(id){selectedCheckinId=String(selectedCheckinId)===String(id)?null:String(id);const r=checkinRecords.find(x=>String(x.transport_id)===String(id));toast(selectedCheckinId?`Seleccionado: ${r?.transporte||""} · Viaje ${r?.viaje||""}`:"Selección cancelada");renderCheckin()}
async function saveCheckin(id,status,novelty){if(!canModule("checkin","edit")){toast("Sin permiso de modificación en el módulo.");return}if(String(selectedCheckinId)!==String(id)){toast("Primero selecciona el transporte.");await refreshCheckin();return}if(!(permissions.is_admin||permissions.can_manage_checkin)){toast("No tienes permiso para modificar Checkin.");await refreshCheckin();return}const r=checkinRecords.find(x=>String(x.transport_id)===String(id));const{error}=await sb.rpc("update_checkin_transport_secure",{p_transport_id:id,p_status:status===null?r?.checkin_estado:status,p_novedad:novelty===null?r?.checkin_novedad:novelty,p_pin:loginPin});if(error){toast(error.message);await refreshCheckin();return}toast("Check-in actualizado.");await refreshCheckin()}


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
