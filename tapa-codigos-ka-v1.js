/* Tapa Códigos KA - importación independiente y resumen de diez clientes, solo 36 SKU seleccionados. */
(()=>{'use strict';
const $=id=>document.getElementById(id);
const CUSTOMERS=new Map(Object.entries({
 '10522923':'Exito Verde','14112183':'Exito San Pedro De Los Milagros',
 '13243394':'Exito','10564376':'Exito Bello',
 '11291509':'Exito Express EDS La Sierra','13114886':'Exito Plaza Fabricato',
 '10564408':'Exito Bello Centro','14156330':'Exito Gardel',
 '14156331':'Exito Puerto Berrio','10564327':'Exito Robledo'
}));
const SELECTED_SKUS=new Set('1428 2174 2641 2976 3228 5150 5252 5255 5630 5632 5902 7993 8532 9672 9909 12283 13631 14703 14709 16210 16466 17155 20469 21146 21159 21189 21730 22102 22242 22256 22265 22270 22271 23013 23047 23121'.split(' '));
const norm=s=>String(s??'').trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-zA-Z0-9]/g,'').toLowerCase();
const code=s=>String(s??'').trim().replace(/\.0+$/,'').replace(/\s/g,'');
const txt=v=>String(v??'').trim();
const qty=v=>typeof v==='number'?v:(Number(String(v??'0').replace(/\s/g,'').replace(/,/g,''))||0);
const fmt=v=>Number(v||0).toLocaleString('es-CO',{maximumFractionDigits:3});
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const findSheet=(wb,label)=>wb.SheetNames.find(n=>norm(n)===norm(label));
const headers=(a,row=0)=>{const m=new Map();(a[row]||[]).forEach((v,i)=>{if(norm(v)&&!m.has(norm(v)))m.set(norm(v),i)});return m};
let completedBook=null, reportRows=[],localProcessedRows=null;
function parseWorkbook(wb){
 const baseName=findSheet(wb,'Bases Datos')||findSheet(wb,'Base Datos');
 const tapaName=findSheet(wb,'Tapa Codigos')||findSheet(wb,'Tapa Códigos');
 if(!baseName||!tapaName)throw Error('El Excel debe contener las pestañas Bases Datos y Tapa Codigos.');
 const source=XLSX.utils.sheet_to_json(wb.Sheets[baseName],{header:1,defval:'',raw:true});
 // Cruces por posiciones fijas del archivo, no por nombres ambiguos de encabezados:
 // Bases Datos A (código cliente) → B (nombre cliente)
 // Bases Datos D (SKU) → E (descripción del producto)
 // Tapa Codigos M (SOLIC.) → T (Cliente)
 // Tapa Codigos G (MATERIAL) → S (Familia: texto de columna E solicitado)
 const hd=headers(source,1);
 if(hd.get('cliente')!==0||hd.get('nombre3')!==1||hd.get('sku')!==3||hd.get('descripcion')!==4){
  throw Error('Bases Datos no tiene el formato esperado: A=Cliente, B=Nombre 3, D=SKU, E=Descripcion (encabezados en fila 2).');
 }
 const clients=new Map(),families=new Map();
 for(const row of source.slice(2)){
  const c=code(row[0]);if(c&&!clients.has(c))clients.set(c,txt(row[1]));
  const sku=code(row[3]);if(sku&&!families.has(sku))families.set(sku,txt(row[4]));
 }
 const sheet=wb.Sheets[tapaName];const data=XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:''});
 const th=headers(data);
 if(th.get('solic')!==12||th.get('material')!==6||th.get('familia')!==18||th.get('cliente')!==19){
  throw Error('Tapa Codigos debe tener G=MATERIAL, M=SOLIC., S=Familia y T=Cliente. Comprueba las columnas.');
 }
 for(const k of ['transporte','vehiculo','viaje','denominacion','cantidad'])if(!th.has(k))throw Error('Falta '+k+' en Tapa Codigos.');
 const result=[];let foundClient=0,foundFamily=0,missingClient=0,missingFamily=0;
 for(let i=1;i<data.length;i++){
  const row=data[i]||[];if(row.every(v=>txt(v)===''))continue;
  const clientCode=code(row[th.get('solic')]);const sku=code(row[th.get('material')]);
  const clientName=clients.get(clientCode)||'';
  const family=families.get(sku)||'';
  if(clientName)foundClient++;else missingClient++;
  if(family)foundFamily++;else missingFamily++;
  // Columna Cliente: nombre legible; la descripción almacenada incluye código para el filtro exacto.
  const ec=XLSX.utils.encode_cell({r:i,c:19});
  const ef=XLSX.utils.encode_cell({r:i,c:18});
  sheet[ec]={...(sheet[ec]||{}),t:'s',v:clientName};
  sheet[ef]={...(sheet[ef]||{}),t:'s',v:family};
  result.push({
   transporte:txt(row[th.get('transporte')]),
   descripcion:clientCode?(clientCode+' | '+clientName):clientName,
   vehiculo:txt(row[th.get('vehiculo')]),viaje:txt(row[th.get('viaje')]),
   material:sku,denominacion:txt(row[th.get('denominacion')]),
   cantidad:qty(row[th.get('cantidad')]),familia:family
  });
 }
 if(!result.length)throw Error('La pestaña Tapa Codigos no contiene registros.');
 return {rows:result,foundClient,foundFamily,missingClient,missingFamily};
}
function groupSelected(rows){
 const groups=new Map();
 for(const r of rows||[]){
  const label=txt(r.descripcion);const match=label.match(/^(\d+)\s*\|\s*(.*)$/);
  const clientId=match?.[1]||'';
  if(!CUSTOMERS.has(clientId)||!SELECTED_SKUS.has(code(r.material)))continue;
  const item={transporte:txt(r.transporte),descripcion:txt(match[2])||CUSTOMERS.get(clientId),vehiculo:txt(r.vehiculo),viaje:txt(r.viaje),material:txt(r.material),denominacion:txt(r.denominacion),cantidad:qty(r.cantidad)};
  const key=[clientId,item.transporte,item.vehiculo,item.viaje,item.material,item.denominacion].join('\u001f');
  if(groups.has(key))groups.get(key).cantidad+=item.cantidad;
  else groups.set(key,item);
 }
 return [...groups.values()].sort((a,b)=>[a.transporte,a.descripcion,a.vehiculo,a.viaje,a.material,a.denominacion].join('|').localeCompare([b.transporte,b.descripcion,b.vehiculo,b.viaje,b.material,b.denominacion].join('|'),'es',{numeric:true}));
}
function render(){
 const records=localProcessedRows||window.ctTapaKaReadRows?.()||[];
 reportRows=groupSelected(records);
 const tbody=$('ctTapaReportRows');if(!tbody)return;
 tbody.innerHTML=reportRows.map(r=>`<tr><td>${escape(r.transporte)}</td><td>${escape(r.descripcion)}</td><td>${escape(r.vehiculo)}</td><td>${escape(r.viaje)}</td><td>${escape(r.material)}</td><td>${escape(r.denominacion)}</td><td>${fmt(r.cantidad)}</td></tr>`).join('')||'<tr><td colspan="7">Sin coincidencias simultáneas de los 10 clientes Éxito y los 36 SKU. Revisa la vista previa del cruce M→T y G→S en Cargar archivos.</td></tr>';
 $('ctTapaReportTotal').textContent=fmt(reportRows.reduce((s,r)=>s+r.cantidad,0));
 const skuMatches=records.filter(r=>SELECTED_SKUS.has(code(r.material))).length;
 const customerMatches=records.filter(r=>CUSTOMERS.has(code(String(r.descripcion||'').split('|')[0]))).length;
 $('ctTapaReportSummary').textContent=`${reportRows.length} combinaciones · ${fmt(reportRows.reduce((s,r)=>s+r.cantidad,0))} unidades. Origen: ${records.length} filas · SKU permitidos: ${skuMatches} · clientes Éxito: ${customerMatches} · coinciden ambos filtros: ${records.filter(r=>SELECTED_SKUS.has(code(r.material))&&CUSTOMERS.has(code(String(r.descripcion||'').split('|')[0]))).length}.`;
}
window.ctTapaKaRenderSummary=render;
function excelDownload(wb,name){XLSX.writeFile(wb,name,{compression:true});}
function downloadFull(){if(!completedBook)return;excelDownload(completedBook,'Tapa_Codigos_Completado.xlsx');}
function downloadReport(){
 if(!reportRows.length){$('ctTapaReportSummary').textContent='No hay registros de los SKU seleccionados para los 10 clientes para descargar.';return;}
 const wb=XLSX.utils.book_new();const cols=['TRANSPORTE','DESCRIPCIÓN','VEHÍCULO','VIAJE','MATERIAL','DENOMINACIÓN','CANTIDAD'];
 const aoa=[cols,...reportRows.map(r=>[r.transporte,r.descripcion,r.vehiculo,r.viaje,r.material,r.denominacion,r.cantidad])];
 aoa.push(['TOTAL GENERAL','','','','','',reportRows.reduce((s,r)=>s+r.cantidad,0)]);
 const ws=XLSX.utils.aoa_to_sheet(aoa);ws['!cols']=[{wch:17},{wch:35},{wch:16},{wch:14},{wch:16},{wch:39},{wch:17}];ws['!autofilter']={ref:`A1:G${reportRows.length+1}`};
 for(let c=0;c<7;c++){
  const addr=XLSX.utils.encode_cell({r:0,c});
  ws[addr].s={font:{bold:true,color:{rgb:'FFFFFF'}},fill:{fgColor:{rgb:'196989'}},alignment:{horizontal:'center'}};
  const tail=XLSX.utils.encode_cell({r:aoa.length-1,c});
  if(ws[tail])ws[tail].s={font:{bold:true,color:{rgb:'232F3D'}},fill:{fgColor:{rgb:'FCD039'}}};
 }
 XLSX.utils.book_append_sheet(wb,ws,'Exito SKU');excelDownload(wb,'Resumen_Tapa_Codigos_KA_Exito_SKU.xlsx');
}
function showCrossPreview(rows,info){
 const host=$('ctTapaCrossPreview');if(!host)return;
 const both=rows.filter(r=>SELECTED_SKUS.has(code(r.material))&&CUSTOMERS.has(code(String(r.descripcion||'').split('|')[0]))).length;
 const shown=rows.slice(0,20);
 host.hidden=false;
 host.innerHTML=`<h4 style="margin:12px 0 6px">Verificación del cruce (primeras ${shown.length} filas)</h4>
 <p class="hint">Cliente M→A→B→T: ${info.foundClient} encontrados, ${info.missingClient} sin coincidencia. Material G→D→E→S: ${info.foundFamily} encontrados, ${info.missingFamily} sin coincidencia. Coinciden con los 10 clientes y 36 SKU del resumen: <b>${both}</b> registros.</p>
 <div class="tableWrap"><table><thead><tr><th>M · SOLIC.</th><th>T · Cliente</th><th>G · MATERIAL</th><th>S · Descripción</th></tr></thead><tbody>${shown.map(r=>{const m=String(r.descripcion||'').split('|');return `<tr><td>${escape(txt(m[0]))}</td><td>${escape(txt(m.slice(1).join('|')))}</td><td>${escape(r.material)}</td><td>${escape(r.familia)}</td></tr>`}).join('')}</tbody></table></div>`;
}
async function upload(){
 const status=$('ctTapaImportStatus'),btn=$('ctTapaProcessBtn'),file=$('ctTapaExcelFile')?.files?.[0];
 if(!file){status.textContent='Selecciona el Excel de Tapa Códigos.';return;}
 if(!window.ctTapaKaCanUpload?.()){status.textContent='No tienes permiso de modificación o carga de archivos. Inicia sesión con un usuario autorizado.';return;}
 btn.disabled=true;completedBook=null;localProcessedRows=null;$('ctTapaFullExcelBtn').disabled=true;
 try{
  status.textContent='Leyendo hojas y cruzando códigos de clientes y materiales…';
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellStyles:true,cellDates:true});
  const result=parseWorkbook(wb);
  completedBook=wb;localProcessedRows=result.rows;
  $('ctTapaFullExcelBtn').disabled=false;showCrossPreview(result.rows,result);render();
  status.textContent=`Cruce realizado: ${result.rows.length} registros, ${result.foundClient} clientes encontrados y ${result.foundFamily} descripciones de producto encontradas. Guardando datos…`;
  await window.ctTapaKaSaveRows(result.rows);
  status.textContent=`Carga exitosa: ${result.rows.length} registros. Cliente: ${result.foundClient} encontrados / ${result.missingClient} sin coincidencia. Descripción en S: ${result.foundFamily} encontradas / ${result.missingFamily} sin coincidencia. Puedes descargar el Excel actualizado y consultar el resumen en Tapa Códigos KA.`;
  render();
 }catch(err){status.textContent='Error: '+(err?.message||err)+(completedBook?' El cruce local sí se realizó: puedes descargar el Excel completado. No se confirmó el guardado en Supabase.':'');console.error('[Tapa Códigos KA]',err)}
 finally{btn.disabled=false}
}
function init(){
 $('ctTapaProcessBtn')?.addEventListener('click',upload);
 $('ctTapaFullExcelBtn')?.addEventListener('click',downloadFull);
 $('ctTapaReportDownload')?.addEventListener('click',downloadReport);
 $('tapaModuleBtn')?.addEventListener('click',()=>setTimeout(render,150));
 render();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
