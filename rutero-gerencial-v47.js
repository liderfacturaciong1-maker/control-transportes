/* Brújula y Rutero: consulta y depuración no destructiva del archivo de origen.
   Excel pivots and formula recalculation REQUIRE an Excel-compatible server engine. */
(()=>{'use strict';
const $=id=>document.getElementById(id), safe=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
let headers=[], rows=[], filtered=[], processed=null, audit=[], lastName='', columnFilters=[];
let filterTimer=null;
const norm=s=>String(s??'').trim().toLocaleLowerCase('es');
const val=c=>c?.w??c?.v??'';
function unpack(ws){if(!ws||typeof ws!=='object')throw Error('La hoja no fue devuelta por el lector Excel.');const range=XLSX.utils.decode_range(ws['!ref']||'A1');const out=[];for(let r=range.s.r;r<=range.e.r;r++){const row=[];let has=false;for(let c=0;c<=Math.max(range.e.c,33);c++){const v=val(ws[XLSX.utils.encode_cell({r,c})]);row.push(v);if(String(v).trim())has=true;}if(has)out.push(row);}return out;}
function dedup(data,start,stop){if(!data.length)return {data,removed:0};const seen=new Set(),out=[];for(const row of data){const key=JSON.stringify(row.slice(start,stop).map(norm));if(key==='[]'||row.slice(start,stop).every(x=>!String(x??'').trim())){out.push(row);continue;}if(seen.has(key))continue;seen.add(key);out.push(row);}return {data:out,removed:data.length-out.length};}
function choices(el,values,label){el.innerHTML=`<option value="">${safe(label)}</option>`+[...new Set(values.map(String).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es')).map(v=>`<option value="${safe(v)}">${safe(v)}</option>`).join('');}
function pickColumn(pattern,fallback){return headers.findIndex(h=>pattern.test(norm(h)));}
function draw(){if(!rows.length)return;const query=norm($('brujulaSearch').value), filters=[['brujulaCenterFilter',/centro|^cd$|origen/,0],['brujulaStatusFilter',/status|estado/,1],['brujulaDestinationFilter',/destino|destinatario/,2]];filtered=rows.filter(row=> (!query||row.some(v=>norm(v).includes(query)))&&filters.every(([id,pattern,fallback])=>!$(id).value||String(row[pickColumn(pattern,fallback)]??'')===$(id).value)&&columnFilters.every((needle,i)=>!needle||norm(row[i]).includes(norm(needle))));$('brujulaCount').textContent=filtered.length.toLocaleString('es-CO')+' de '+rows.length.toLocaleString('es-CO')+' registros';const transportIndex=pickColumn(/^transporte$/);$('brujulaRows').innerHTML=filtered.slice(0,250).map(row=>'<tr>'+headers.map((_,i)=>i===transportIndex?`<td><button type="button" class="byrTransportLink" data-transport="${safe(row[i])}" title="Ver resumen del transporte">${safe(row[i])}</button></td>`:`<td>${safe(row[i])}</td>`).join('')+'</tr>').join('');$('brujulaLimit').textContent=filtered.length>250?'Se muestran los primeros 250 registros. Exporta para consultar el total.':'';const status=pickColumn(/status|estado/,1);const counts={};for(const row of filtered){const k=status<0?'Sin columna de estado':String(row[status]||'Sin estado');counts[k]=(counts[k]||0)+1;}$('byrKpiTotal').textContent=rows.length.toLocaleString('es-CO');$('byrKpiFiltered').textContent=filtered.length.toLocaleString('es-CO');$('byrKpiDup').textContent=audit.reduce((a,b)=>a+(Number(b.removed)||0),0).toLocaleString('es-CO');$('byrStatusChart').innerHTML=Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10).map(([k,v])=>`<div class="byrBarRow"><span title="${safe(k)}">${safe(k)}</span><div class="byrBar"><i style="width:${Math.max(2,100*v/Math.max(1,filtered.length))}%"></i></div><b>${v}</b></div>`).join('')||'<p>Sin datos</p>';}
function busy(txt){$('brujulaMessage').textContent=txt;}
async function importFile(){if(typeof canModule==='function'&&!canModule('rutero','edit')){busy('No tienes permiso para modificar Brújula y Rutero.');return}
 const file=$('brujulaFile').files[0];
 if(!file)return busy('Selecciona un archivo Excel.');
 if(!/\.(xlsx|xlsb|xls)$/i.test(file.name))return busy('Formato no permitido. Sube un archivo .xlsx, .xlsb o .xls.');
 $('brujulaUploadBtn').disabled=true;
 busy('Leyendo Rutero; los archivos grandes pueden tardar unos segundos…');
 try{
  await new Promise(r=>setTimeout(r,30));
  const buffer=await file.arrayBuffer();
  const wb=XLSX.read(buffer,{type:'array',cellFormula:false,cellText:true,cellDates:false});
  const sheetName=wb.SheetNames.find(n=>norm(n)==='rutero') || (wb.SheetNames.length===1 ? wb.SheetNames[0] : null);
  if(!sheetName)throw Error('No se encontró la hoja Rutero. Hojas disponibles: '+wb.SheetNames.join(', ')+'.');
  const ws=wb.Sheets?.[sheetName];
  if(!ws || !ws['!ref'])throw Error('La hoja Rutero está vacía o el lector no pudo interpretarla. Guarda una copia como .xlsx (valores, sin fórmulas) e inténtalo de nuevo.');
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,blankrows:false});
  const first=raw.findIndex(r=>Array.isArray(r)&&r.some(v=>String(v??'').trim()));
  if(first<0)throw Error('La hoja Rutero no contiene registros.');
  const matrix=raw.slice(first);
  const width=Math.max(matrix[0].length,...matrix.slice(1,100).map(r=>r.length));
  headers=Array.from({length:width},(_,i)=>String(matrix[0][i]||'Columna '+XLSX.utils.encode_col(i)).trim());
  const sourceRows=matrix.slice(1).filter(r=>r.some(v=>String(v??'').trim()));
  const unique=dedup(sourceRows,0,width);
  rows=unique.data.map(r=>Array.from({length:width},(_,i)=>r[i]??''));
  audit=[{name:'Rutero',original:sourceRows.length,remaining:rows.length,removed:unique.removed}];
  lastName=file.name;
  busy('Guardando el informe compartido en Supabase…');
  await CTShared.save('rutero',file.name,rows,{headers,audit});
  processed={Rutero:true};window.dispatchEvent(new Event("rutero-updated"));
  initReportView();
  busy('Informe gerencial generado a partir de Rutero: '+rows.length.toLocaleString('es-CO')+' registros únicos. No se recalculan fórmulas ni tablas dinámicas.');
 }catch(e){processed=null;busy('No se pudo guardar el informe compartido: '+e.message+'. El informe publicado anteriormente se conserva.');console.error(e);}finally{$('brujulaUploadBtn').disabled=false;}
}

window.getRuteroSyncData=()=>({headers,rows,loaded:Boolean(processed)});
function initReportView(){
  const selectConfigs=[['brujulaCenterFilter',/centro|^cd$|origen/,0,'Todos los centros'],['brujulaStatusFilter',/status|estado/,1,'Todos los estados'],['brujulaDestinationFilter',/destino|destinatario/,2,'Todos los destinos']];
  columnFilters=Array(headers.length).fill('');$('brujulaHead').innerHTML='<tr>'+headers.map((h,i)=>`<th><span>${safe(h)}</span><input class="byrColumnFilter" data-column="${i}" type="search" placeholder="Filtrar…" aria-label="Filtrar ${safe(h)}" autocomplete="off"></th>`).join('')+'</tr>';
  for(const [id,pattern,fallback,label] of selectConfigs){
   const col=headers.findIndex(h=>pattern.test(norm(h)));
   const el=$(id);
   if(col<0){el.innerHTML='<option value="">'+safe(label)+' (no disponible)</option>';el.disabled=true;}
   else{el.disabled=false;choices(el,rows.map(r=>r[col]),label);}
  }
  $('brujulaSearch').value='';
  $('brujulaLoadedAt').textContent='Archivo: '+lastName+' · '+new Date().toLocaleString('es-CO');
  $('byrAudit').innerHTML=audit.map(x=>`<tr><td>${safe(x.name)}</td><td>${x.original}</td><td>${x.remaining}</td><td>${x.removed}</td></tr>`).join('');
  draw();
}
async function restoreSharedRutero(){
 if(typeof canModule==='function'&&!canModule('rutero','view'))return;
 if(processed)return;
 busy('Recuperando último informe compartido de Supabase…');
 try{
  const saved=await CTShared.load('rutero');
  if(!saved){busy('No hay informe compartido. Un usuario autorizado puede cargar el archivo Rutero.');return;}
  headers=saved.metadata.headers||[];rows=saved.items;audit=saved.metadata.audit||[];
  processed={Rutero:true};lastName=saved.filename;window.dispatchEvent(new Event("rutero-updated"));
  initReportView();$('brujulaLoadedAt').textContent='Archivo compartido: '+saved.filename+' · '+new Date(saved.saved_at).toLocaleString('es-CO');
  busy('Informe gerencial recuperado de Supabase: '+rows.length.toLocaleString('es-CO')+' registros.');
 }catch(e){busy('No se pudo recuperar Rutero: '+e.message);console.error(e)}
}
function col(name){return headers.findIndex(h=>norm(h)===norm(name));}
function distinctValues(dataset,index){if(index<0)return [];return [...new Set(dataset.map(r=>String(r[index]??'').trim()).filter(Boolean))];}
function numberOf(value){const s=String(value??'').trim().replace(/\s/g,'');if(!s)return 0;const standardized=s.includes(',')?s.replace(/\./g,'').replace(',','.'):s;const x=Number(standardized);return Number.isFinite(x)?x:0;}
function showTransport(transport){
 const idx=col('Transporte');if(idx<0)return;
 const group=rows.filter(r=>String(r[idx]??'').trim()===String(transport).trim());if(!group.length)return;
 const find=(...names)=>{for(const name of names){const k=headers.findIndex(h=>norm(h).replace(/[^a-z0-9]/g,'')===norm(name).replace(/[^a-z0-9]/g,''));if(k>=0)return k;}return -1;};
 const values=(...names)=>distinctValues(group,find(...names));
 const field=(label,names)=>{const v=values(...names);return `<div class="byrDetailField"><small>${safe(label)}</small><strong>${safe(v.length?v.join(' · '):'Sin información')}</strong></div>`;};
 const multi=(label,names)=>{const v=values(...names);return `<div class="byrDetailField byrDropdown"><small>${safe(label)} (${v.length})</small><details><summary>${v.length?('Ver '+v.length+' registros'):'Sin información'}</summary><ul>${v.map(x=>`<li>${safe(x)}</li>`).join('')}</ul></details></div>`;};
 const total=(label,names)=>{const i=find(...names);const val=i<0?'Sin columna':group.reduce((n,r)=>n+numberOf(r[i]),0).toLocaleString('es-CO',{maximumFractionDigits:2});return `<div class="byrDetailField byrDetailTotal"><small>${safe(label)}</small><strong>${safe(val)}</strong></div>`;};
 $('byrDetailTitle').textContent='Transporte '+transport;
 $('byrDetailMeta').textContent=group.length.toLocaleString('es-CO')+' registros asociados';
 $('byrDetailContent').innerHTML=
 field('Viaje',['Viaje'])+field('Placa',['Placa'])+field('AgServTran',['AgServTran'])+field('Status',['Status'])+field('Destino',['Destino'])+field('Clientes',['Clientes'])+field('APP',['APP'])+field('Teléfono',['Telefono','Teléfono','Celular'])+field('Fecha',['Fecha'])+
 total('Cantidad de cajas',['Cantidad de Cajas','Cantidad','Cajas'])+total('HL',['HL'])+
 multi('Establecimientos',['Establecimiento'])+multi('Pedidos',['Pedido'])+multi('Entregas',['Entrega']);
 const modal=$('byrTransportModal');
 // Ensure the overlay is a direct child of body: nested panels with backdrop-filter
 // create a containing block for position:fixed and push the dialog below the table.
 if(modal.parentElement!==document.body) document.body.appendChild(modal);
 modal.classList.remove('hidden');
 document.body.classList.add('byrDialogOpen');
 const pane=modal.querySelector('.byrTransportCard');if(pane)pane.scrollTop=0;
 $('byrTransportClose').focus();
}
function hideTransport(){$('byrTransportModal').classList.add('hidden');document.body.classList.remove('byrDialogOpen');}

function exportFile(){if(!rows.length)return busy('Primero carga un archivo.');const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([headers,...filtered]),'Rutero filtrado');XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(audit.map(x=>({Hoja:x.name,'Filas originales':x.original,'Filas conservadas':x.remaining,'Filas duplicadas':x.removed}))), 'Auditoría');XLSX.writeFile(wb,'Informe_Rutero_'+new Date().toISOString().slice(0,10)+'.xlsx');}
window.loadBrujulaReport=function(){if(!processed)restoreSharedRutero();};
function initialize(){if(!$('brujulaUploadBtn'))return;
 $('brujulaHead').addEventListener('input',e=>{if(!e.target.matches('.byrColumnFilter'))return;columnFilters[Number(e.target.dataset.column)]=e.target.value;clearTimeout(filterTimer);filterTimer=setTimeout(draw,180);});
 $('brujulaRows').addEventListener('click',e=>{const b=e.target.closest('.byrTransportLink');if(b)showTransport(b.dataset.transport);});
 $('byrTransportClose').addEventListener('click',hideTransport);$('byrTransportModal').addEventListener('click',e=>{if(e.target.id==='byrTransportModal')hideTransport();});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('byrTransportModal').classList.contains('hidden'))hideTransport();});busy('Versión Rutero 4.7 activa · Solo se requiere la hoja Rutero.');$('brujulaUploadBtn').onclick=importFile;$('brujulaDownloadBtn').onclick=exportFile;['brujulaSearch','brujulaCenterFilter','brujulaStatusFilter','brujulaDestinationFilter'].forEach(id=>$(id).addEventListener(id==='brujulaSearch'?'input':'change',draw));}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',initialize);else initialize();
})();
