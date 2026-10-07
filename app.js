const cfg=window.APP_CONFIG||{};
let sb=null,user=null,records=[],timer=null;
const $=id=>document.getElementById(id);
const STATES=["Pendiente","Armado","Contado","Cargado"];
function toast(m){$("toast").textContent=m;$("toast").classList.add("show");setTimeout(()=>$("toast").classList.remove("show"),2500)}
function configured(){return cfg.SUPABASE_URL&&!cfg.SUPABASE_URL.includes("TU-PROYECTO")&&cfg.SUPABASE_ANON_KEY&&!cfg.SUPABASE_ANON_KEY.includes("TU_ANON")}
async function init(){
 if(!configured()){$("loginMsg").innerHTML="Falta configurar <b>config.js</b>. Revisa README.";return}
 sb=supabase.createClient(cfg.SUPABASE_URL,cfg.SUPABASE_ANON_KEY);
 $("loginBtn").onclick=login;$("pin").onkeydown=e=>e.key==="Enter"&&login();
 $("logout").onclick=logout;$("search").oninput=render;$("statusFilter").onchange=render;
 $("importBtn").onclick=importExcel;$("downloadBtn").onclick=downloadExcel;$("addUser").onclick=addUser;
}
async function login(){
 $("loginMsg").textContent="Verificando…";
 const pin=$("pin").value.trim();
 const {data,error}=await sb.rpc("login_by_pin",{p_pin:pin});
 if(error||!data?.length){$("loginMsg").textContent="PIN incorrecto o usuario desactivado.";return}
 user=data[0]; sessionStorage.setItem("transport_user",JSON.stringify(user));
 await showApp();
}
async function showApp(){
 $("login").classList.add("hidden");$("app").classList.remove("hidden");$("logout").classList.remove("hidden");
 $("userName").textContent=user.name;$("connection").textContent="● Conectado";
 if(user.is_admin)$("adminPanel").classList.remove("hidden");
 await refresh(); subscribe(); if(user.is_admin)loadUsers();
}
function logout(){sessionStorage.clear();location.reload()}
async function refresh(){
 const {data,error}=await sb.rpc("list_transports");
 if(error){toast(error.message);return} records=data||[];render();loadHistory();
}
function render(){
 const q=($("search").value||"").toLowerCase(), sf=$("statusFilter").value;
 const list=records.filter(r=>(!sf||r.estado===sf)&&(!q||[r.transporte,r.viaje,r.placa].join(" ").toLowerCase().includes(q)));
 for(const s of STATES)$("rows");
 $("total").textContent=records.length;
 $("pendientes").textContent=records.filter(r=>r.estado==="Pendiente").length;
 $("armados").textContent=records.filter(r=>r.estado==="Armado").length;
 $("contados").textContent=records.filter(r=>r.estado==="Contado").length;
 $("cargados").textContent=records.filter(r=>r.estado==="Cargado").length;
 $("rows").innerHTML=list.map(r=>`<tr><td>${esc(r.transporte)}</td><td>${esc(r.viaje)}</td><td><b>${esc(r.placa)}</b></td><td><span class="badge ${r.estado}">${r.estado}</span></td><td><button onclick="advance(${r.id})" ${r.estado==="Cargado"?"disabled":""}>${r.estado==="Pendiente"?"Armar":r.estado==="Armado"?"Contar":"Cargar"}</button></td></tr>`).join("")||'<tr><td colspan="5">No hay registros.</td></tr>';
}
async function advance(id){
 const {data,error}=await sb.rpc("advance_transport",{p_transport_id:id});
 if(error){toast(error.message);return} toast("Estado actualizado");await refresh();
}
async function loadHistory(){
 const {data,error}=await sb.rpc("recent_history",{p_limit:30});
 if(error)return;
 $("history").innerHTML=(data||[]).map(h=>`<div class="historyItem"><b>${esc(h.usuario)}</b> · ${esc(h.placa)} · Viaje ${esc(h.viaje)}<br>${h.estado_anterior} → <b>${h.estado_nuevo}</b> · ${new Date(h.created_at).toLocaleString("es-CO")}</div>`).join("")||"Sin movimientos.";
}
async function loadUsers(){
 const {data,error}=await sb.rpc("admin_users"); if(error)return;
 $("users").innerHTML=(data||[]).map(u=>`<div class="userRow"><span>${esc(u.name)} ${u.is_admin?"👑":""}</span><span>PIN: ••••</span><button onclick="toggleUser('${u.id}',${u.active})" class="secondary">${u.active?"Desactivar":"Activar"}</button></div>`).join("");
}
async function addUser(){
 const name=$("newName").value.trim(),pin=$("newPin").value.trim();
 if(!name||!pin){toast("Nombre y PIN son obligatorios");return}
 const {error}=await sb.rpc("admin_create_user",{p_name:name,p_pin:pin});
 if(error){toast(error.message);return} $("newName").value="";$("newPin").value="";loadUsers();toast("Usuario creado");
}
async function toggleUser(id,active){
 const {error}=await sb.rpc("admin_set_user_active",{p_user_id:id,p_active:!active});
 if(error){toast(error.message);return}loadUsers();
}
async function importExcel(){
 const f=$("excelFile").files[0];if(!f){toast("Selecciona un archivo");return}
 const buf=await f.arrayBuffer(),wb=XLSX.read(buf,{type:"array"}),sheet=wb.Sheets[wb.SheetNames[0]],data=XLSX.utils.sheet_to_json(sheet,{defval:""});
 if(!data.length){toast("Excel vacío");return}
 const keys=Object.keys(data[0]),norm=s=>String(s).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\s/g,"");
 const find=n=>keys.find(k=>norm(k)===n);const kt=find("transporte"),kv=find("viaje"),kp=find("placa");
 if(!kt||!kv||!kp){toast("Faltan Transporte, Viaje o Placa");return}
 const rows=data.map(x=>({transporte:String(x[kt]),viaje:String(x[kv]),placa:String(x[kp])}));
 const {error}=await sb.rpc("replace_transports",{p_rows:rows});
 if(error){toast(error.message);return}toast(`${rows.length} transportes cargados`);refresh();
}
async function downloadExcel(){
 const {data,error}=await sb.rpc("list_transports");if(error)return;
 const out=(data||[]).map(r=>({Transporte:r.transporte,Viaje:r.viaje,Placa:r.placa,Estado:r.estado}));
 const ws=XLSX.utils.json_to_sheet(out),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Transportes");XLSX.writeFile(wb,"transportes_actualizados.xlsx");
}
function subscribe(){
 sb.channel("transportes-live").on("postgres_changes",{event:"*",schema:"public",table:"transportes"},()=>refresh()).subscribe();
}
function esc(s){return String(s??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m]))}
init();