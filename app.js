/* BRÚJULA Y RUTERO — pegar dentro de app.js, antes de init(); */
const BRUJULA_COLUMNS = ["Centro","Fecha","Transporte","Placa","Viaje","AgServTran","Denominación","Solic.","Nombre","Establecimiento","Calle","Pedido","Fecha Preferente Entrega","Status","Destino","APP","Telefono"];
let brujulaReportRows = [], brujulaReportTimestamp = null;

function brujulaNorm(v){return String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").trim().toLowerCase().replace(/\s+/g," ")}
function brujulaEsc(v){return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]))}
function brujulaCell(row,headers,name){let i=headers.findIndex(h=>brujulaNorm(h)===brujulaNorm(name));if(i<0&&name==="Solic.")i=headers.findIndex(h=>brujulaNorm(h).replace(/\.$/,"")==="solic");return i<0?"":String(row[i]??"").trim()}
function brujulaFindHeader(grid){for(let i=0;i<Math.min(30,grid.length);i++){const h=(grid[i]||[]).map(x=>brujulaNorm(x));if(h.includes("transporte")&&h.includes("placa")&&h.includes("agservtran"))return i}return -1}
async function processBrujulaFile(){
  if(!(permissions.is_admin||permissions.can_upload)){toast("No tienes permiso para cargar informes.");return}
  const file=$("brujulaFile").files[0];if(!file){$("brujulaMessage").textContent="Selecciona el archivo BYR APP.";return}
  const btn=$("brujulaUploadBtn");btn.disabled=true;$("brujulaMessage").textContent="Leyendo el libro y preparando el informe…";
  try{
    const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
    const sheetName=wb.SheetNames.find(n=>brujulaNorm(n)==="rutero");
    if(!sheetName)throw Error('No se encontró la pestaña "Rutero".');
    const grid=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:"",raw:false});
    const hi=brujulaFindHeader(grid);if(hi<0)throw Error('No se reconocieron los encabezados de "Rutero".');
    const headers=grid[hi].map(x=>String(x??"").trim());
    const missing=BRUJULA_COLUMNS.filter(k=>!headers.some(h=>brujulaNorm(h)===brujulaNorm(k)));
    if(missing.length)throw Error("Faltan columnas: "+missing.join(", ")+". Verifica el archivo y las celdas calculadas.");
    const next=grid.slice(hi+1).filter(r=>r.some(v=>String(v??"").trim()!=="")).map(r=>{const o={};BRUJULA_COLUMNS.forEach(k=>o[k]=brujulaCell(r,headers,k));return o}).filter(r=>r.Transporte||r.Placa||r.Pedido);
    if(!next.length)throw Error('No se encontraron registros en "Rutero".');
    const {error}=await sb.rpc("brujula_save_report",{p_rows:next,p_pin:loginPin});if(error)throw error;
    brujulaReportRows=next;brujulaReportTimestamp=new Date().toISOString();populateBrujulaFilters();renderBrujulaReport();
    $("brujulaMessage").textContent="Informe reemplazado correctamente: "+next.length.toLocaleString("es-CO")+" registros.";
  }catch(e){$("brujulaMessage").textContent="No se pudo procesar el archivo: "+(e?.message||e)}
  finally{btn.disabled=false}
}
async function loadBrujulaReport(){
  if(!sb||!loginPin)return;
  const {data,error}=await sb.rpc("brujula_get_report",{p_pin:loginPin});
  if(error){$("brujulaMessage").textContent="No se pudo consultar el informe. Ejecuta primero brujula_upgrade.sql en Supabase.";return}
  brujulaReportRows=Array.isArray(data?.[0]?.report_rows)?data[0].report_rows:[];
  brujulaReportTimestamp=data?.[0]?.uploaded_at||null;populateBrujulaFilters();renderBrujulaReport();
  $("brujulaMessage").textContent=brujulaReportRows.length?"Informe compartido cargado. La próxima carga reemplazará todos los registros.":"Todavía no hay un informe cargado.";
}
function populateBrujulaFilters(){
  [["brujulaCenterFilter","Centro","Todos los centros"],["brujulaStatusFilter","Status","Todos los Status"],["brujulaDestinationFilter","Destino","Todos los destinos"]].forEach(([id,key,label])=>{
    const el=$(id),old=el.value,vals=[...new Set(brujulaReportRows.map(r=>String(r[key]||"").trim()).filter(Boolean))].sort((a,b)=>a.localeCompare(b,"es"));
    el.innerHTML='<option value="">'+brujulaEsc(label)+'</option>'+vals.map(v=>'<option value="'+brujulaEsc(v)+'">'+brujulaEsc(v)+'</option>').join("");if(vals.includes(old))el.value=old;
  });
}
function renderBrujulaReport(){
  const q=brujulaNorm($("brujulaSearch")?.value||""),center=$("brujulaCenterFilter")?.value||"",status=$("brujulaStatusFilter")?.value||"",dest=$("brujulaDestinationFilter")?.value||"";
  const rows=brujulaReportRows.filter(r=>(!center||String(r.Centro||"")===center)&&(!status||String(r.Status||"")===status)&&(!dest||String(r.Destino||"")===dest)&&(!q||BRUJULA_COLUMNS.some(k=>brujulaNorm(r[k]).includes(q))));
  $("brujulaHead").innerHTML="<tr>"+BRUJULA_COLUMNS.map(k=>"<th>"+brujulaEsc(k)+"</th>").join("")+"</tr>";
  $("brujulaRows").innerHTML=rows.slice(0,1500).map(r=>"<tr>"+BRUJULA_COLUMNS.map(k=>"<td>"+brujulaEsc(r[k])+"</td>").join("")+"</tr>").join("");
  $("brujulaCount").textContent=rows.length.toLocaleString("es-CO")+" registros"+(rows.length>1500?" · mostrando los primeros 1.500":"");
  $("brujulaLoadedAt").textContent=brujulaReportTimestamp?"Última carga: "+new Date(brujulaReportTimestamp).toLocaleString("es-CO"):"Informe aún no cargado";
}
function downloadBrujulaReport(){
  const q=brujulaNorm($("brujulaSearch").value),center=$("brujulaCenterFilter").value,status=$("brujulaStatusFilter").value,dest=$("brujulaDestinationFilter").value;
  const rows=brujulaReportRows.filter(r=>(!center||String(r.Centro||"")===center)&&(!status||String(r.Status||"")===status)&&(!dest||String(r.Destino||"")===dest)&&(!q||BRUJULA_COLUMNS.some(k=>brujulaNorm(r[k]).includes(q))));
  if(!rows.length){toast("No hay filas para descargar con los filtros actuales.");return}
  const data=[BRUJULA_COLUMNS,...rows.map(r=>BRUJULA_COLUMNS.map(k=>r[k]??""))],ws=XLSX.utils.aoa_to_sheet(data);
  ws["!cols"]=BRUJULA_COLUMNS.map(k=>({wch:Math.max(14,k.length+2)}));ws["!autofilter"]={ref:ws["!ref"]};
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Informe Rutero");XLSX.writeFile(wb,"Informe_Gerencial_Brujula_Rutero.xlsx");
}
