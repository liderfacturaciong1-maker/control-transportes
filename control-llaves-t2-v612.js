(()=>{'use strict';const $=id=>document.getElementById(id);if(!$('ctKeysPlate'))return;
let mode='LLAVE RECIBIDA', busy=false;
const PLATE_PATTERN=/^[A-Z]{3,6}\d{2,4}(?:-\d{1,2})?$/;
const normalizePlate=v=>String(v??'').trim().toUpperCase().replace(/[\s\u00a0]+/g,'').replace(/[–—−]/g,'-');
const safe=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const rpc=(name,args={})=>{if(!window.ctRotationApi)throw Error('Inicia sesión en Control Transportes');return window.ctRotationApi(name,args)};
function editable(){return window.ctBillingTabCanEdit?.('handover')===true}
function msg(t,bad=false){$('ctKeysMsg').textContent=t;$('ctKeysMsg').style.color=bad?'#b91c1c':'#166534'}
function setMode(m){mode=m;document.querySelectorAll('[data-key-action]').forEach(b=>b.classList.toggle('selected',b.dataset.keyAction===m));$('ctKeysPersonWrap').hidden=m!=='LLAVE PRESTADA';$('ctKeysPlate').focus()}
function render(rows){$('ctKeysRows').innerHTML=rows.length?rows.map(r=>{const klass=r.estado==='LLAVE RECIBIDA'?'received':r.estado==='LLAVE PRESTADA'?'loaned':'delivered';return `<tr><td><strong>${safe(r.placa)}</strong></td><td>${safe(r.transportes||'SIN ASIGNACIÓN')}</td><td>${safe(r.prestada_a||'—')}</td><td>${safe(r.updated_at?new Date(r.updated_at).toLocaleString('es-CO'):'—')}</td></tr>`}).join(''):'<tr><td colspan="4">Sin movimientos registrados.</td></tr>'}
async function load(){try{const rows=await rpc('ct_keys_t2_get');render(rows||[]);$('ctKeysApply').disabled=!editable();$('ctKeysPerson').disabled=!editable();$('ctKeysPlate').disabled=!editable();}catch(e){msg('No se pudo consultar llaves: '+e.message,true)}}
async function apply(){if(busy)return;const placa=normalizePlate($('ctKeysPlate').value);const persona=$('ctKeysPerson').value.trim();if(!PLATE_PATTERN.test(placa))return msg('La lectura no corresponde a una placa. Vuelve a escanear (ej. COJTX912 o LWL645).',true);if(mode==='LLAVE PRESTADA'&&!persona)return msg('Escribe el nombre de quien recibe la llave',true);if(!editable())return msg('Tu usuario no tiene permiso de modificar Entrega de turno',true);busy=true;$('ctKeysApply').disabled=true;try{const result=await rpc('ct_keys_t2_set',{p_plate:placa,p_state:mode,p_borrower:mode==='LLAVE PRESTADA'?persona:null});msg(`${placa}: ${mode}. Transporte(s): ${result?.transportes||'Sin asignación en Facturas'}`);$('ctKeysPlate').value='';if(mode==='LLAVE PRESTADA')$('ctKeysPerson').value='';await load();$('ctKeysPlate').focus()}catch(e){msg('No se guardó: '+e.message,true)}finally{busy=false;$('ctKeysApply').disabled=!editable()}}

// Cámara móvil: API nativa BarcodeDetector, con alternativa ZXing cuando no está disponible.
let scanning=false, scanStream=null,scanTimer=null,zxingReader=null,zxingControls=null,torch=false,lastCandidate='',candidateCount=0;
const scanCanvas=document.createElement('canvas');const scanCtx=scanCanvas.getContext('2d',{willReadFrequently:true});
async function setZoom(){const tr=scanStream?.getVideoTracks()?.[0];if(!tr)return;const target=Number($('ctKeysZoom')?.value||1);const caps=tr.getCapabilities?.()||{};if(caps.zoom){try{await tr.applyConstraints({advanced:[{zoom:Math.max(caps.zoom.min,Math.min(caps.zoom.max,target))}]});return}catch(_){}} // Visual zoom only when hardware zoom unavailable
const video=$('ctKeysVideo');video.style.transform=`scale(${Math.min(target,2)})`;video.style.transformOrigin='center center';}
async function toggleTorch(){const tr=scanStream?.getVideoTracks()?.[0],caps=tr?.getCapabilities?.()||{};if(!caps.torch)return;try{torch=!torch;await tr.applyConstraints({advanced:[{torch}]});$('ctKeysTorch').textContent=torch?'🔦 Apagar luz':'🔦 Encender luz'}catch(_){}}

function cameraNote(t){$('ctKeysCameraMsg').textContent=t}
function stopCamera(){scanning=false;lastCandidate='';candidateCount=0;if(scanTimer){clearTimeout(scanTimer);scanTimer=null}try{zxingControls?.stop()}catch(_){}zxingControls=null;try{zxingReader?.reset?.()}catch(_){}zxingReader=null;if(scanStream){scanStream.getTracks().forEach(t=>t.stop());scanStream=null}const video=$('ctKeysVideo');try{video.pause()}catch(_){}video.srcObject=null;video.style.transform='';torch=false;$('ctKeysTorch').hidden=true;$('ctKeysCamera').hidden=true}
function scannedCode(raw){
 const plate=normalizePlate(raw);
 if(!PLATE_PATTERN.test(plate)){
  cameraNote('Lectura descartada: '+plate.slice(0,24)+'. El código debe contener la placa, no únicamente números.');return false;
 }
 if(lastCandidate!==plate){lastCandidate=plate;candidateCount=1;cameraNote('Placa '+plate+' detectada. Mantén el código estable para confirmar.');return false;}
 candidateCount++;
 if(candidateCount<2)return false;
 $('ctKeysPlate').value=plate;stopCamera();msg('Placa '+plate+' identificada. Verifica la operación y presiona Registrar.');$('ctKeysApply').focus();return true;
}
function loadZxing(){return new Promise((resolve,reject)=>{if(window.ZXingBrowser)return resolve(window.ZXingBrowser);const tag=document.createElement('script');tag.src='https://cdn.jsdelivr.net/npm/@zxing/browser@0.1.5/umd/zxing-browser.min.js';tag.onload=()=>window.ZXingBrowser?resolve(window.ZXingBrowser):reject(Error('No se pudo inicializar el lector alternativo'));tag.onerror=()=>reject(Error('No se pudo cargar el lector de cámara alternativo. Comprueba la conexión a internet.'));document.head.appendChild(tag)})}
async function openCamera(){if(scanning)return;if(!window.isSecureContext||!navigator.mediaDevices?.getUserMedia)return msg('La cámara requiere HTTPS y permisos del navegador. Usa el lector físico o escribe la placa.',true);if(!editable())return msg('No tienes permiso para registrar movimientos de llaves',true);scanning=true;$('ctKeysCamera').hidden=false;cameraNote('Solicitando permiso de cámara…');try{
 const stream=await navigator.mediaDevices.getUserMedia({audio:false,video:{facingMode:{ideal:'environment'},width:{ideal:1920},height:{ideal:1080},resizeMode:'none'}});if(!scanning){stream.getTracks().forEach(t=>t.stop());return}scanStream=stream;const video=$('ctKeysVideo');video.srcObject=stream;await video.play();const caps=stream.getVideoTracks()[0]?.getCapabilities?.()||{};$('ctKeysTorch').hidden=!caps.torch;await setZoom();cameraNote('Código de 3 cm: prueba zoom 1,5× o 2×, evita reflejos y espera que enfoque.');
 if('BarcodeDetector' in window){const scanner=new BarcodeDetector({formats:['code_128']});const loop=async()=>{if(!scanning)return;try{let values=await scanner.detect(video);
 if(values?.length&&values.some(v=>v.format==='code_128'&&scannedCode(v.rawValue)))return;
 // Amplificación central: escanea la región donde se enfoca la etiqueta pequeña.
 if(scanCtx&&video.videoWidth&&video.videoHeight){const w=video.videoWidth,h=video.videoHeight;const cw=Math.floor(w*.66),ch=Math.floor(h*.36);scanCanvas.width=Math.min(1600,cw);scanCanvas.height=Math.round(scanCanvas.width*ch/cw);scanCtx.imageSmoothingEnabled=true;scanCtx.imageSmoothingQuality='high';scanCtx.drawImage(video,(w-cw)/2,(h-ch)/2,cw,ch,0,0,scanCanvas.width,scanCanvas.height);values=await scanner.detect(scanCanvas);if(values?.length&&values.some(v=>v.format==='code_128'&&scannedCode(v.rawValue)))return;}}catch(_){}if(scanning)scanTimer=setTimeout(loop,140)};loop();}
 else{cameraNote('Activando lector compatible con este dispositivo…');try{const ZX=await loadZxing();if(!scanning)return;scanStream?.getTracks().forEach(t=>t.stop());scanStream=null;video.srcObject=null;zxingReader=new ZX.BrowserMultiFormatReader();zxingControls=await zxingReader.decodeFromVideoDevice(undefined,video,(result)=>{if(result)scannedCode(result.getText())});cameraNote('Enfoca las barras para leer la placa.')}catch(e){stopCamera();msg('Este navegador no pudo activar el escáner: '+e.message+' Usa el lector Bluetooth o el ingreso manual.',true)}}
 }catch(e){stopCamera();msg('No fue posible abrir la cámara: '+e.message,true)}}
// Préstamos históricos: cada evento se conserva incluso al registrar una devolución.
async function loadSummary(){
 try{const list=await rpc('ct_keys_t2_loans_get');const items=Array.isArray(list)?list:[];
 $('ctKeysSummaryCount').textContent=items.length+' préstamo'+(items.length===1?'':'s');
 $('ctKeysSummaryRows').innerHTML=items.length?items.map(r=>`<tr><td>${safe(new Date(r.registrado_el).toLocaleString('es-CO'))}</td><td><strong>${safe(r.placa)}</strong></td><td>${safe(r.prestada_a||'—')}</td><td>${safe(r.usuario||'—')}</td></tr>`).join(''):'<tr><td colspan="4">No hay préstamos registrados.</td></tr>';
 $('ctKeysSummaryDelete').hidden=!editable();
 }catch(e){$('ctKeysSummaryRows').innerHTML='<tr><td colspan="4">No fue posible consultar el historial.</td></tr>';msg('Error de historial: '+e.message,true)}
}
async function deleteSummary(){if(!editable())return msg('No tienes autorización para borrar préstamos históricos',true);
 if(!confirm('¿Eliminar definitivamente TODO el histórico de préstamos de llaves? Esta acción no se puede deshacer y no modifica el estado actual de las llaves.'))return;
 try{await rpc('ct_keys_t2_loans_clear');await loadSummary();msg('Histórico de préstamos eliminado. Los estados de llaves y facturas se conservaron.')}catch(e){msg('No se pudo borrar el histórico: '+e.message,true)}
}
function openSummary(){$('ctKeysSummaryModal').hidden=false;loadSummary()}
function closeSummary(){$('ctKeysSummaryModal').hidden=true}
function init(){$('ctKeysSummaryOpen').onclick=openSummary;$('ctKeysSummaryClose').onclick=closeSummary;$('ctKeysSummaryRefresh').onclick=loadSummary;$('ctKeysSummaryDelete').onclick=deleteSummary;$('ctKeysSummaryModal').addEventListener('click',e=>{if(e.target===$('ctKeysSummaryModal'))closeSummary()});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('ctKeysSummaryModal').hidden)closeSummary()});$('ctKeysZoom').onchange=setZoom;$('ctKeysTorch').onclick=toggleTorch;$('ctKeysCameraOpen').onclick=openCamera;$('ctKeysCameraClose').onclick=stopCamera;document.addEventListener('visibilitychange',()=>{if(document.hidden)stopCamera()});document.querySelectorAll('[data-handover-tab]').forEach(b=>b.addEventListener('click',()=>{if(b.dataset.handoverTab!=='T2')stopCamera()}));document.querySelectorAll('[data-key-action]').forEach(b=>b.onclick=()=>setMode(b.dataset.keyAction));$('ctKeysApply').onclick=apply;$('ctKeysPlate').addEventListener('input',e=>e.target.value=e.target.value.toUpperCase());$('ctKeysPlate').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();apply()}});$('ctKeysRefresh').onclick=load;document.querySelector('[data-handover-tab="T2"]')?.addEventListener('click',load);window.addEventListener('ct-billing-opened',load);setMode(mode)}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
