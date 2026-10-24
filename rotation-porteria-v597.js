(()=>{'use strict';const $=id=>document.getElementById(id);const api=(n,p)=>window.ctRotationApi(n,p);const admin=()=>window.ctRotationIsAdmin?.()===true;
const defaults={variations:['06:00 - 14:00','14:00 - 22:00','22:00 - 06:00'],people:['Marcy','Luisa','Daniela','Claudia'],shifts:['Turno 1 (06:00 - 14:00)','Turno 2 (14:00 - 22:00)','Turno 3 (22:00 - 06:00) Liquidación','Turno 3 (22:00 - 06:00) Preventa']};let conf=structuredClone(defaults),start='',draft={},saved=[],editing=false;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));const iso=d=>{const dt=new Date(d.getTime()-d.getTimezoneOffset()*60000);return dt.toISOString().slice(0,10)};
const parse=s=>new Date(s+'T12:00:00');const add=(s,n)=>{const d=parse(s);d.setDate(d.getDate()+n);return iso(d)};const saturday=s=>{const d=parse(s);d.setDate(d.getDate()-(d.getDay()+1)%7);return iso(d)};const fmt=s=>parse(s).toLocaleDateString('es-CO',{day:'2-digit',month:'2-digit',year:'numeric'});const day=s=>parse(s).toLocaleDateString('es-CO',{weekday:'long'}).toUpperCase();
const STANDARD_SHIFT_TIMES=['06:00 - 14:00','14:00 - 22:00','22:00 - 06:00'];
const status=s=>{$('gateRMessage').textContent=s};const slotInfo=x=>typeof x==='string'?{name:x,rest:false}:({name:String(x?.name||''),rest:!!x?.rest,shift:String(x?.shift||'')});
const hasStandardTime=label=>STANDARD_SHIFT_TIMES.some(t=>String(label||'').includes(t));
const slotTone=(shiftLabel,v)=>{if(v?.name&&v?.rest)return 'gateRToneRest';if(v?.name&&!v?.rest&&!hasStandardTime(v.shift||shiftLabel))return 'gateRToneCustom';return ''};
function option(values,chosen){return '<option value="">— Sin asignar —</option>'+values.map(p=>'<option value="'+esc(p)+'" '+(p===chosen?'selected':'')+'>'+esc(p)+'</option>').join('')}
function shiftDefault(sh){return conf.variations.find(t=>String(sh).includes(t))||conf.variations[0]||'';}
function render(){if(!$('gateRWeekDate'))return;$('gateRWeekDate').value=start;$('gateRWeekLabel').textContent=fmt(start)+' — '+fmt(add(start,6));$('gateRThead').innerHTML='<tr><th>Fecha</th><th>DÍA</th>'+conf.shifts.map(x=>'<th>'+esc(x)+'</th>').join('')+'</tr>';
$('gateRTbody').innerHTML=Array.from({length:7},(_,i)=>{const date=add(start,i),row=draft[date]||{};return '<tr class="rotMiniHead"><td colspan="2">'+day(date)+' · '+fmt(date)+'</td>'+conf.shifts.map((sh,j)=>{const v=slotInfo(row[String(j)]),sel=v.shift||shiftDefault(sh);return '<td class="rotDayHeaderCell"><span class="rotDayShiftTitle">'+esc(sh)+'</span><select class="rotHeaderShift '+slotTone(sh,v)+'" aria-label="Horario para '+esc(sh)+' el '+date+'" data-type="shift" data-date="'+date+'" data-slot="'+j+'" '+(!admin()?'disabled':'')+'>'+[...new Set([sel,...conf.variations].filter(Boolean))].map(t=>'<option value="'+esc(t)+'" '+(sel===t?'selected':'')+'>'+esc(t)+'</option>').join('')+'</select></td>'}).join('')+'</tr><tr class="rotDayRow"><td>'+fmt(date)+'</td><td>'+day(date)+'</td>'+conf.shifts.map((sh,j)=>{const v=slotInfo(row[String(j)]),off=!admin()?'disabled':'',tone=slotTone(sh,v),cls=tone?(' '+tone):'';return '<td class="rotAssignment"><div class="rotSlotCompact'+cls+'"><select class="rotPersonChoice'+cls+'" aria-label="Persona '+esc(sh)+' '+date+'" data-type="person" data-date="'+date+'" data-slot="'+j+'" '+off+'>'+option(conf.people,v.name)+'</select><select class="rotWorkChoice'+cls+'" aria-label="Trabaja o descansa" data-type="work" data-date="'+date+'" data-slot="'+j+'" '+off+'><option value="trabaja" '+(!v.rest?'selected':'')+'>Trabaja</option><option value="descansa" '+(v.rest?'selected':'')+'>Descansa</option></select></div></td>'}).join('')+'</tr>'}).join('');
$('gateRSave').disabled=!admin();['gateRPeopleToggle','gateRShiftsToggle','gateRVariationsToggle'].forEach(id=>$(id).disabled=!admin());}
function editor(kind){const list=conf[kind],root=$(kind==='people'?'gateRPeopleRows':kind==='variations'?'gateRVariationRows':'gateRShiftRows');root.innerHTML=list.map((v,i)=>'<div class="rotEditRow"><input data-item="'+i+'" value="'+esc(v)+'" aria-label="'+(kind==='people'?'Persona':kind==='variations'?'Variación de horario':'Turno')+' '+(i+1)+'"><button type="button" data-remove="'+i+'">✕</button></div>').join('');root.querySelectorAll('[data-remove]').forEach(b=>b.onclick=()=>{list.splice(Number(b.dataset.remove),1);editor(kind)});root.querySelectorAll('input').forEach(el=>el.oninput=()=>{list[Number(el.dataset.item)]=el.value});}
async function loadConfig(){try{const x=await api('ct_gate_rotation_config_get');if(x?.people?.length)conf.people=x.people;if(x?.shifts?.length)conf.shifts=x.shifts;if(x?.variations?.length)conf.variations=x.variations;}catch(e){status('No se cargó la configuración: '+e.message)}editor('people');editor('shifts');editor('variations');render()}
async function saveConfig(){if(!admin())return;const people=[...new Set(conf.people.map(x=>x.trim()).filter(Boolean))];const shifts=conf.shifts.map(x=>x.trim()).filter(Boolean);const variations=[...new Set(conf.variations.map(x=>x.trim()).filter(Boolean))];if(!people.length||!shifts.length||shifts.length>8){alert('Debe existir al menos una persona y un turno (máximo 8).');return}if(!variations.length)return alert('Agrega al menos una variación de horario.');conf={people,shifts,variations};try{await api('ct_gate_rotation_config_save',{p_data:conf});editor('people');editor('shifts');editor('variations');render();status('Configuración guardada en Supabase.')}catch(e){status('Error al guardar configuración: '+e.message)}}
async function loadWeek(){draft={};editing=false;try{const x=await api('ct_gate_rotation_weeks_get');saved=Array.isArray(x)?x:[];const one=saved.find(w=>w.week_start===start);if(one){draft=structuredClone(one.schedule||{});status('Semana guardada. Los cambios requieren Confirmar y guardar.');}else status('Semana sin confirmar: selecciona las personas y presiona Confirmar y guardar.');history();render()}catch(e){status('No fue posible consultar Supabase: '+e.message)}}
function history(){
 const list=saved.slice().sort((a,b)=>b.week_start.localeCompare(a.week_start));
 const count=Math.max(4,...list.map(w=>(w.shifts||[]).length));
 const head=$('gateRHistoryHead');
 if(head)head.innerHTML='<tr><th>Fecha</th><th>Día</th>'+Array.from({length:count},(_,j)=>'<th>Turno '+(j+1)+'</th>').join('')+'</tr>';
 const rows=[];
 for(const w of list){
  const shifts=w.shifts||[];
  const buttons=admin()?'<button type="button" class="rotAction" data-edit="'+esc(w.week_start)+'">Editar semana</button> <button type="button" class="rotAction rotDanger" data-delete="'+esc(w.week_start)+'">Eliminar semana</button>':'';
  rows.push('<tr class="rotHistoryWeek"><td colspan="'+(count+2)+'"><div class="rotHistoryWeekTitle"><strong>Semana '+fmt(w.week_start)+' — '+fmt(add(w.week_start,6))+'</strong><span>'+buttons+'</span></div></td></tr>');
  rows.push('<tr class="rotHistoryShiftLabels"><td colspan="2">TURNOS ASIGNADOS</td>'+Array.from({length:count},(_,j)=>'<td>'+esc(shifts[j]||'—')+'</td>').join('')+'</tr>');
  for(let i=0;i<(w.schedule?.[add(w.week_start,7)]?8:7);i++){
   const date=add(w.week_start,i);
   let cells='';
   for(let j=0;j<count;j++){
    const v=slotInfo(w.schedule?.[date]?.[String(j)]);
    const shift=shifts[j]||'';
    if(!v.name){cells+='<td class="rotHistoryEmpty">—</td>';continue;}
    const worktime=v.shift||shiftDefault(shift);
    const special=!v.rest&&!hasStandardTime(worktime);
    const cls=v.rest?'gateRHistoryRest':special?'gateRHistorySpecial':'gateRHistoryNormal';
    cells+='<td><div class="rotHistoryAssignment '+cls+'"><strong>'+esc(v.name)+'</strong><span>'+esc(worktime)+'</span><small>'+(v.rest?'DESCANSA':'TRABAJA')+'</small></div></td>';
   }
   rows.push('<tr class="rotHistoryDay"><td>'+fmt(date)+'</td><td>'+day(date)+'</td>'+cells+'</tr>');
  }
 }
 const picker=$('gateRPhotoWeek');if(picker){const previous=picker.value;picker.innerHTML=list.map(w=>'<option value="'+esc(w.week_start)+'">'+fmt(w.week_start)+' — '+fmt(add(w.week_start,6))+'</option>').join('');if(list.some(w=>w.week_start===previous))picker.value=previous;}
 $('gateRShowHistory').disabled=!list.length;
 $('gateRHistoryBody').innerHTML=rows.join('')||'<tr><td colspan="'+(count+2)+'">Aún no hay turnos asignados en semanas confirmadas.</td></tr>';
 $('gateRHistoryBody').querySelectorAll('[data-edit]').forEach(b=>b.onclick=()=>{start=b.dataset.edit;const w=saved.find(x=>x.week_start===start);if(w){conf.shifts=[...w.shifts];draft=structuredClone(w.schedule||{});editor('shifts');render();status('Editando semana histórica '+fmt(start)+'. Guarda para confirmar los cambios.')}});
 $('gateRHistoryBody').querySelectorAll('[data-delete]').forEach(b=>b.onclick=async()=>{const week=b.dataset.delete;if(!confirm('¿Eliminar de forma permanente toda la semana '+fmt(week)+'?'))return;try{await api('ct_gate_rotation_week_delete',{p_week_start:week});await loadWeek();status('Semana eliminada.')}catch(e){status('No se pudo eliminar: '+e.message)}});
}


function fitPhotoToScreen(){
 const viewport=$('gateRPhotoModal')?.querySelector('.rotPhotoViewport'),sheet=$('gateRPhotoSheet');
 if(!viewport||!sheet||$('gateRPhotoModal').hidden)return;
 // El transform CSS reduce la imagen, pero no su altura de diseño. Ubicarla
 // absolutamente evita que la fila viernes sea recortada por esa altura fantasma.
 sheet.style.transform='none';sheet.style.left='0px';sheet.style.top='0px';
 const maxW=Math.max(1,viewport.clientWidth-8),maxH=Math.max(1,viewport.clientHeight-8);
 const naturalW=sheet.offsetWidth,naturalH=sheet.scrollHeight;
 const scale=Math.min(1,maxW/naturalW,maxH/naturalH);
 sheet.style.transform='scale('+scale+')';
 sheet.style.left=Math.max(0,(viewport.clientWidth-naturalW*scale)/2)+'px';
 sheet.style.top=Math.max(0,(viewport.clientHeight-naturalH*scale)/2)+'px';
}
function showPhoto(weekStart){
 const modal=$('gateRPhotoModal'),w=saved.find(x=>x.week_start===weekStart);
 if(!modal||!w)return;
 $('gateRPhotoPeriod').textContent='Semana del '+fmt(w.week_start)+' al '+fmt(add(w.week_start,6));
 const shifts=w.shifts||[],count=Math.max(4,shifts.length);
 let html='<table class="rotPhotoTable"><thead><tr><th>FECHA</th><th>DÍA</th>'+Array.from({length:count},(_,j)=>'<th>Turno '+(j+1)+'</th>').join('')+'</tr></thead><tbody>';
 html+='<tr class="rotPhotoShiftRow"><td colspan="2">TURNOS ASIGNADOS</td>'+Array.from({length:count},(_,j)=>'<td>'+esc(shifts[j]||'—')+'</td>').join('')+'</tr>';
 for(let i=0;i<(w.schedule?.[add(w.week_start,7)]?8:7);i++){
  const date=add(w.week_start,i);html+='<tr><td class="rotPhotoDate">'+fmt(date)+'</td><td class="rotPhotoDate">'+day(date)+'</td>';
  for(let j=0;j<count;j++){
   const v=slotInfo(w.schedule?.[date]?.[String(j)]),shift=shifts[j]||'';
   if(!v.name){html+='<td class="rotPhotoEmpty">—</td>';continue;}
   const t=v.shift||shiftDefault(shift),cls=v.rest?' rotPhotoRest':!hasStandardTime(t)?' rotPhotoSpecial':'';
   html+='<td><div class="rotPhotoCell'+cls+'"><b>'+esc(v.name)+'</b><span>'+esc(t)+'</span><small>'+(v.rest?'DESCANSA':'TRABAJA')+'</small></div></td>';
  }
  html+='</tr>';
 }
 $('gateRPhotoContent').innerHTML=html+'</tbody></table>';
 $('gateRPhotoWeek').value=w.week_start;modal.hidden=false;document.body.style.overflow='hidden';fitPhotoToScreen();requestAnimationFrame(()=>requestAnimationFrame(fitPhotoToScreen));$('gateRPhotoClose').focus();
}
function closePhoto(){const m=$('gateRPhotoModal');if(m)m.hidden=true;document.body.style.overflow='';}

async function saveWeek(){if(!admin())return;if(!confirm('¿Confirmar y guardar la programación del '+fmt(start)+' al '+fmt(add(start,6))+'?'))return;try{await api('ct_gate_rotation_week_save',{p_week_start:start,p_shifts:conf.shifts,p_schedule:draft});await loadWeek();status('✅ Programación semanal guardada en el acumulado.')}catch(e){status('No se pudo guardar: '+e.message)}}
async function copyHistory(){
 const lines=[];
 for(const w of saved.slice().sort((a,b)=>a.week_start.localeCompare(b.week_start))){
  const shifts=w.shifts||[];
  lines.push('Semana del '+fmt(w.week_start)+' al '+fmt(add(w.week_start,6)));
  lines.push(['Fecha','Día',...shifts].join('\t'));
  for(let i=0;i<(w.schedule?.[add(w.week_start,7)]?8:7);i++){
   const date=add(w.week_start,i);
   lines.push([fmt(date),day(date),...shifts.map((s,j)=>{
    const v=slotInfo(w.schedule?.[date]?.[String(j)]);
    if(!v.name)return '';
    return [v.name,v.shift||shiftDefault(s),v.rest?'DESCANSA':'TRABAJA'].join(' | ');
   })].join('\t'));
  }
  lines.push('');
 }
 try{await navigator.clipboard.writeText(lines.join('\n'));status('Resumen horizontal copiado. Puedes pegarlo en Excel o Outlook.')}catch(e){status('No fue posible copiar: '+e.message)}
}

function init(){if(!$('gateRTbody'))return;window.addEventListener('resize',fitPhotoToScreen);start=saturday(iso(new Date()));$('gateRTbody').addEventListener('change',e=>{const el=e.target.closest('[data-date][data-slot]');if(!el)return;const row=(draft[el.dataset.date]??={});const existing=slotInfo(row[el.dataset.slot]);if(el.dataset.type==='person')existing.name=el.value;else if(el.dataset.type==='work')existing.rest=el.value==='descansa';else if(el.dataset.type==='shift')existing.shift=el.value;row[el.dataset.slot]=existing;render();});$('gateRPrev').onclick=()=>{start=add(start,-7);loadWeek()};$('gateRNext').onclick=()=>{start=add(start,7);loadWeek()};$('gateRWeekDate').onchange=e=>{if(!e.target.value)return;start=saturday(e.target.value);loadWeek()};$('gateRPeopleToggle').onclick=()=>{$('gateRPeopleEditor').hidden=!$('gateRPeopleEditor').hidden};$('gateRShiftsToggle').onclick=()=>{$('gateRShiftsEditor').hidden=!$('gateRShiftsEditor').hidden};$('gateRVariationsToggle').onclick=()=>{$('gateRVariationsEditor').hidden=!$('gateRVariationsEditor').hidden};$('gateRAddPerson').onclick=()=>{conf.people.push('');editor('people')};$('gateRAddShift').onclick=()=>{if(conf.shifts.length>=8)return alert('Máximo 8 turnos.');conf.shifts.push('');editor('shifts')};$('gateRAddVariation').onclick=()=>{conf.variations.push('');editor('variations')};$('gateRQuickAddVariation').disabled=!admin();$('gateRQuickAddVariation').onclick=()=>{if(!admin())return;$('gateRVariationsEditor').hidden=false;conf.variations.push('');editor('variations');const fields=$('gateRVariationRows').querySelectorAll('input');const last=fields[fields.length-1];last?.focus();last?.scrollIntoView({behavior:'smooth',block:'center'});status('Escribe el nuevo horario (por ejemplo 08:00 - 16:00) y presiona Guardar horarios para incluirlo en las listas.');};$('gateRSaveVariations').onclick=saveConfig;$('gateRSavePeople').onclick=saveConfig;$('gateRSaveShifts').onclick=saveConfig;$('gateRSave').onclick=saveWeek;$('gateRRefreshHistory').onclick=loadWeek;$('gateRCopyHistory').onclick=copyHistory;$('gateRShowHistory').onclick=()=>{if(saved.length)showPhoto($('gateRPhotoWeek').value||saved.slice().sort((a,b)=>b.week_start.localeCompare(a.week_start))[0].week_start)};$('gateRPhotoClose').onclick=closePhoto;$('gateRPhotoWeek').onchange=e=>showPhoto(e.target.value);$('gateRPhotoModal').addEventListener('click',e=>{if(e.target===$('gateRPhotoModal'))closePhoto()});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('gateRPhotoModal').hidden)closePhoto()});document.getElementById('portPorteriaSubTab')?.addEventListener('click',()=>{loadConfig().then(loadWeek)});window.addEventListener('ct-billing-opened',()=>{loadConfig().then(loadWeek)});render();}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();