/* Cargue WIP — consolida exclusivamente hoja DATOS, sin modificar Supabase. */
(()=>{'use strict';
 const $=id=>document.getElementById(id);
 const HEADERS=['TRANSPORTE','CENTRO','FECHA DESPACHO','TRANSPORTISTA','VEHICULO','TIPO ENTREGA','VIAJE','MATERIAL','DENOMINACION','CANTIDAD','UM'];
 const SOURCE=['TRANSPORTE','CENTRO','FECHA DESPACHO','TRANSPORTISTA','VEHICULO','TIPO ENTREGA','VIAJE','MATERIAL','DENOMINACION','UM','CANTIDAD'];
 const normalize=value=>String(value??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]+/g,'');
 const safe=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const string=value=>String(value??'').trim();
 const numeric=value=>{
  if(typeof value==='number')return value;
  let raw=string(value).replace(/\s/g,''); if(!raw)return NaN;
  if(/^-?\d{1,3}(?:\.\d{3})+(?:,\d+)?$/.test(raw))raw=raw.replace(/\./g,'').replace(',','.');
  else if(/^-?\d{1,3}(?:,\d{3})+(?:\.\d+)?$/.test(raw))raw=raw.replace(/,/g,'');
  else if(raw.includes(',')&&!raw.includes('.'))raw=raw.replace(',','.');
  return Number(raw);
 };
 let items=[],sourceFilename='',countOriginal=0;
 function message(msg,error=false){const p=$('wipMessage');if(p){p.textContent=msg;p.classList.toggle('wipError',error);}}
 function date(value){
  if(value instanceof Date&&!isNaN(value))return `${String(value.getDate()).padStart(2,'0')}.${String(value.getMonth()+1).padStart(2,'0')}.${value.getFullYear()}`;
  if(typeof value==='number'&&value>30000&&value<100000){const d=XLSX.SSF.parse_date_code(value);if(d)return `${String(d.d).padStart(2,'0')}.${String(d.m).padStart(2,'0')}.${d.y}`;}
  return string(value);
 }
 function parse(workbook){
  const sheetName=workbook.SheetNames.find(n=>normalize(n)==='DATOS');
  if(!sheetName)throw Error('El archivo debe contener una pestaña llamada DATOS. No se procesará DINAMICA.');
  const matrix=XLSX.utils.sheet_to_json(workbook.Sheets[sheetName],{header:1,defval:'',raw:true,blankrows:false});
  const headerIndex=matrix.findIndex(row=>{const h=new Set(row.map(normalize));return SOURCE.every(x=>h.has(normalize(x)));});
  if(headerIndex<0)throw Error('DATOS no contiene las columnas requeridas: '+SOURCE.join(', '));
  const header=matrix[headerIndex].map(normalize);
  const pos=Object.fromEntries(SOURCE.map(key=>[key,header.indexOf(normalize(key))]));
  const aggregated=new Map();let count=0,skipped=0;
  for(const raw of matrix.slice(headerIndex+1)){
   const values=Object.fromEntries(SOURCE.map(key=>[key,raw[pos[key]]]));
   const transport=string(values.TRANSPORTE),material=string(values.MATERIAL);
   if(!transport&&!material)continue;
   if(!transport||!material){skipped++;continue;}
   const qty=numeric(values.CANTIDAD);
   if(!Number.isFinite(qty)){skipped++;continue;}
   count++;
   const row=[string(values.TRANSPORTE),string(values.CENTRO),date(values['FECHA DESPACHO']),string(values.TRANSPORTISTA),string(values.VEHICULO),string(values['TIPO ENTREGA']),string(values.VIAJE),string(values.MATERIAL),string(values.DENOMINACION),0,string(values.UM)];
   const key=JSON.stringify(row.filter((_,i)=>i!==9));
   if(!aggregated.has(key))aggregated.set(key,row);
   aggregated.get(key)[9]+=qty;
  }
  if(!count)throw Error('No se encontraron registros de transporte/material con cantidad válida.');
  const result=[...aggregated.values()];
  // Orden tipo tabla dinámica: transportes recientes primero, luego dimensiones y material.
  result.sort((a,b)=>{const primary=b[0].localeCompare(a[0],"es",{numeric:true});if(primary)return primary;for(const i of [1,2,3,4,5,6,7,8,10]){const diff=a[i].localeCompare(b[i],"es",{numeric:true});if(diff)return diff;}return 0;});
  return {result,count,skipped};
 }
 function render(){
  const body=$('wipRows');if(!body)return;
  const fmt=new Intl.NumberFormat('es-CO',{maximumFractionDigits:0});
  $('wipOriginal').textContent=countOriginal.toLocaleString('es-CO');
  $('wipConsolidated').textContent=items.length.toLocaleString('es-CO');
  $('wipDownload').disabled=!items.length;
  body.innerHTML=items.length?items.slice(0,100).map(row=>'<tr>'+row.map((v,i)=>`<td>${safe(i===9?fmt.format(Math.round(v)):v)}</td>`).join('')+'</tr>').join(''):'<tr><td colspan="11">Selecciona el archivo Excel y pulsa Generar tabla.</td></tr>';
  $('wipPreviewNote').textContent=items.length?`Vista previa de ${Math.min(items.length,100)} de ${items.length.toLocaleString('es-CO')} filas. La descarga incluye todas las filas.`:'Se mostrarán las primeras 100 filas del resultado.';
 }
 async function load(){
  const file=$('wipFile').files[0];if(!file)return message('Selecciona un archivo Excel .xlsx o .xls.',true);
  if(!/\.xlsx?$/i.test(file.name))return message('Selecciona un archivo con extensión .xlsx o .xls.',true);
  const btn=$('wipGenerate');btn.disabled=true;items=[];countOriginal=0;render();message('Leyendo hoja DATOS y consolidando cantidades…');
  try{
   if(!window.XLSX)throw Error('No está disponible el lector Excel. Recarga la aplicación.');
   const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:false});
   const parsed=parse(wb);items=parsed.result;countOriginal=parsed.count;sourceFilename=file.name;render();
   message(`Listo: ${parsed.count.toLocaleString('es-CO')} registros consolidados en ${items.length.toLocaleString('es-CO')} filas.${parsed.skipped?` Se omitieron ${parsed.skipped} filas incompletas.`:''} CANTIDAD se exporta como número entero.`);
  }catch(e){message('Error al procesar: '+(e.message||e),true);}finally{btn.disabled=false;}
 }
 function download(){
  if(!items.length)return;
  // Mantener valores numéricos, no texto; no exportar ceros decimales ni puntos finales.
  const data=[HEADERS,...items.map(r=>r.map((v,i)=>i===9?Math.round(v):v))];
  const sheet=XLSX.utils.aoa_to_sheet(data);
  sheet['!cols']=[16,10,18,16,18,16,13,13,39,14,10].map(wch=>({wch}));
  sheet['!autofilter']={ref:`A1:K${data.length}`};
  for(let i=2;i<=data.length;i++){
   const c=sheet[`J${i}`];if(c){c.t='n';c.z='0';}
   const transport=sheet[`A${i}`];if(transport){transport.t='s';transport.v=String(transport.v);}
   const material=sheet[`H${i}`];if(material){material.t='s';material.v=String(material.v);}
  }
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,sheet,'CARGUE WIP');
  XLSX.writeFile(wb,'Cargue_WIP_Consolidado.xlsx',{compression:true});
  message(`Excel descargado: ${items.length.toLocaleString('es-CO')} filas, 11 columnas. Archivo origen: ${sourceFilename}.`);
 }
 function init(){if(!$('wipGenerate'))return;$('wipGenerate').addEventListener('click',load);$('wipDownload').addEventListener('click',download);$('wipFile').addEventListener('change',()=>{items=[];countOriginal=0;render();message('Archivo seleccionado. Haz clic en Generar tabla.');});render();}
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init,{once:true});else init();
})();
