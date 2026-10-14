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
let cargoFallback=[],cargoFetching=false,cargoLastAttempt=0,cargoError='';
function cargoSource(){
 // Los registros originales de Cargue y Despacho son la unica fuente.
 if(typeof records!=='undefined' && Array.isArray(records) && records.length) return records;
 if(Array.isArray(window.syncCargoRecords) && window.syncCargoRecords.length) return window.syncCargoRecords;
 return cargoFallback;
}
async function loadCargoSource(force=false){
 if(cargoFetching)return;
 if(!force && cargoSource().length)return;
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
function render(){if(!$('syncRows'))return;
 // SOURCE OF TRUTH: Cargue y Despacho, not Rutero / STAT.
 const cargo=cargoSource();
 const report=window.getRuteroSyncData?.()||{headers:[],rows:[],loaded:false};
 const ti=ix(report.headers||[],['TRANSPORTE','TRANSPORT','NUMERODETRANSPORTE','NROTRANSPORTE']);
 const si=ix(report.headers||[],['STATUS','ESTADO']);
 const quick=new Map();if(report.loaded&&ti>=0){for(const r of report.rows){const k=transportKey(r[ti]);if(!k)continue;if(!quick.has(k))quick.set(k,new Set());const s=si>=0?String(r[si]??'').trim():'';if(s)quick.get(k).add(s);}}
 const q=norm($('syncSearch')?.value),f=$('syncFilter')?.value||'';
 const rows=cargo.map(r=>{const key=transportKey(r.transporte);const statuses=quick.get(key);const status=statuses?.size?[...statuses].join(' · '):'—';
  let result;if(statLoaded){const vals=statRows.get(key);result=vals?([...(vals||[])].filter(Boolean).join(' · ')||'En STAT'):'No está en STAT';}
  else result=!report.loaded?'Sin archivo':statuses?.size?'Sincronizado':'No encontrado';
  return {transporte:displayedTransport(r.transporte),viaje:r.viaje??'',placa:r.placa??'',status,result};});
 currentSyncRows=rows;
 const filtered=rows.filter(r=>(!f||r.result===f)&&(!q||[r.transporte,r.viaje,r.placa,r.status,r.result].some(v=>norm(v).includes(q))));
 $('syncTotal').textContent=filtered.length.toLocaleString('es-CO');
 $('syncOk').textContent=filtered.filter(r=>statLoaded?(statRows.has(transportKey(r.transporte))):r.result==='Sincronizado').length.toLocaleString('es-CO');
 $('syncMissing').textContent=filtered.filter(r=>statLoaded?!statRows.has(transportKey(r.transporte)):r.result==='No encontrado').length.toLocaleString('es-CO');
 $('syncMessage').textContent=cargoError && !cargo.length?('No se pudo consultar Cargue y Despacho: '+cargoError):(!cargo.length?(cargoFetching?'Consultando transportes de Cargue y Despacho…':'No hay transportes disponibles en Cargue y Despacho. Pulsa Actualizar cruce.'):`${filtered.length.toLocaleString('es-CO')} registros de Cargue y Despacho visibles. ${statLoaded?'Cruce STAT: '+statFile:report.loaded?'Status consultado en Vista rápida.':'Sin archivo de referencia cargado.'}`);
 updateUnsyncedCount();
 $('syncRows').innerHTML=filtered.slice(0,1000).map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td>${esc(r.placa)}</td><td>${esc(r.status)}</td><td><span class="syncBadge ${/^(Sincronizado|En STAT)$/i.test(r.result)?'syncGood':/^(No encontrado|No está en STAT)$/i.test(r.result)?'syncBad':'syncUnknown'}">${esc(r.result)}</span></td></tr>`).join('');
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

window.renderSyncReport=()=>{render();if(!cargoSource().length)void loadCargoSource();};window.loadSyncStatReport=()=>{void restore();void loadCargoSource();};
window.addEventListener('rutero-updated',render);
window.addEventListener('focus',()=>{if(!$('syncModuleView')?.classList.contains('hidden')){render();void loadCargoSource();}});
document.addEventListener('DOMContentLoaded',()=>{for(const id of ['syncSearch','syncFilter'])$(id)?.addEventListener('input',render);$('syncStatUpload')?.addEventListener('click',upload);$('syncRefresh')?.addEventListener('click',()=>{void loadCargoSource(true);void restore();render();});$('syncAlertBtn')?.addEventListener('click',openUnsynced);$('syncAlertClose')?.addEventListener('click',closeUnsynced);$('syncAlertModal')?.addEventListener('click',e=>{if(e.target===$('syncAlertModal'))closeUnsynced();});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('syncAlertModal')?.classList.contains('hidden'))closeUnsynced();});});
})();