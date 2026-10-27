(()=>{'use strict';const $=id=>document.getElementById(id);if(!$('ctKeysPlate'))return;
let mode='LLAVE RECIBIDA', busy=false;
const PLATE_PATTERN=/^[A-Z]{3,6}\d{2,4}(?:-\d{1,2})?$/;
const normalizePlate=v=>String(v??'').trim().toUpperCase().replace(/[\s\u00a0]+/g,'').replace(/[–—−]/g,'-');
const safe=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const rpc=(name,args={})=>{if(!window.ctValidationKeysRpc)throw Error('Inicia sesión en Validaciones Operación');return window.ctValidationKeysRpc(name,args)};
function editable(){return window.ctValidationKeysCanEdit?.()===true}
function msg(t,bad=false){$('ctKeysMsg').textContent=t;$('ctKeysMsg').style.color=bad?'#b91c1c':'#166534'}
function setMode(m){mode=m;document.querySelectorAll('[data-key-action]').forEach(b=>b.classList.toggle('selected',b.dataset.keyAction===m));$('ctKeysPersonWrap').hidden=m!=='LLAVE PRESTADA';$('ctKeysPlate').focus()}
function render(rows){$('ctKeysRows').innerHTML=rows.length?rows.map(r=>{const klass=r.estado==='LLAVE RECIBIDA'?'received':r.estado==='LLAVE PRESTADA'?'loaned':'delivered';return `<tr><td><strong>${safe(r.placa)}</strong></td><td>${safe(r.transportes||'SIN ASIGNACIÓN')}</td><td>${safe(r.prestada_a||'—')}</td><td>${safe(r.updated_at?new Date(r.updated_at).toLocaleString('es-CO'):'—')}</td></tr>`}).join(''):'<tr><td colspan="4">Sin movimientos registrados.</td></tr>'}
async function load(){try{const rows=await rpc('ct_validation_keys_get');render(rows||[]);$('ctKeysApply').disabled=!editable();$('ctKeysPerson').disabled=!editable();$('ctKeysPlate').disabled=!editable();}catch(e){msg('No se pudo consultar llaves: '+e.message,true)}}
async function apply(){if(busy)return;const placa=normalizePlate($('ctKeysPlate').value);const persona=$('ctKeysPerson').value.trim();if(!PLATE_PATTERN.test(placa))return msg('La lectura no corresponde a una placa. Vuelve a escanear (ej. COJTX912 o LWL645).',true);if(mode==='LLAVE PRESTADA'&&!persona)return msg('Escribe el nombre de quien recibe la llave',true);if(!editable())return msg('Tu usuario no tiene permiso de modificar Entrega de turno',true);busy=true;$('ctKeysApply').disabled=true;try{const result=await rpc('ct_validation_keys_set',{p_plate:placa,p_state:mode,p_borrower:mode==='LLAVE PRESTADA'?persona:null});msg(`${placa}: ${mode}. Transporte(s): ${result?.transportes||'Sin asignación en Facturas'}`);$('ctKeysPlate').value='';if(mode==='LLAVE PRESTADA')$('ctKeysPerson').value='';await load();$('ctKeysPlate').focus()}catch(e){msg('No se guardó: '+e.message,true)}finally{busy=false;$('ctKeysApply').disabled=!editable()}}

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
function loadZxing(){return new Promise((resolve,reject)=>{if(window.ZXingBrowser)return resolve(window.ZXingBrowser);const tag=document.createElement('script');tag.src='https://cdn.jsdelivr.net/npm/@zxing/browser@0.2.1/umd/zxing-browser.min.js';tag.onload=()=>window.ZXingBrowser?resolve(window.ZXingBrowser):reject(Error('No se pudo inicializar el lector alternativo'));tag.onerror=()=>reject(Error('No se descargó el lector alternativo; comprueba internet.'));document.head.appendChild(tag)})}
async function useZxing(video){
 cameraNote('Activando lector alternativo Code 128…');
 try{
  const ZX=await loadZxing();if(!scanning)return;
  scanStream?.getTracks().forEach(t=>t.stop());scanStream=null;video.srcObject=null;
  const hints=new Map();
  // Code 128 primero; evitamos interpretar códigos EAN o numéricos.
  if(window.ZXing?.DecodeHintType&&window.ZXing?.BarcodeFormat){
   hints.set(window.ZXing.DecodeHintType.POSSIBLE_FORMATS,[window.ZXing.BarcodeFormat.CODE_128]);
   hints.set(window.ZXing.DecodeHintType.TRY_HARDER,true);
  }
  zxingReader=new ZX.BrowserMultiFormatReader(hints);
  zxingControls=await zxingReader.decodeFromVideoDevice(undefined,video,(result)=>{if(result)scannedCode(result.getText())});
  cameraNote('Lector alternativo activo. Coloca el código a 10–15 cm y aléjalo/acércalo hasta verlo nítido.');
 }catch(e){stopCamera();msg('No se activó lector alternativo: '+e.message+'. Usa lector Bluetooth o placa manual.',true)}
}
async function openCamera(){
 if(scanning)return;
 if(!window.isSecureContext||!navigator.mediaDevices?.getUserMedia)return msg('La cámara requiere HTTPS y permisos del navegador.',true);
 if(!editable())return msg('No tienes permiso para registrar movimientos de llaves',true);
 scanning=true;$('ctKeysCamera').hidden=false;cameraNote('Solicitando permiso de cámara…');
 try{
  const stream=await navigator.mediaDevices.getUserMedia({audio:false,video:{facingMode:{ideal:'environment'},width:{ideal:1920},height:{ideal:1080}}});
  if(!scanning){stream.getTracks().forEach(t=>t.stop());return}
  scanStream=stream;const video=$('ctKeysVideo');video.srcObject=stream;await video.play();
  const caps=stream.getVideoTracks()[0]?.getCapabilities?.()||{};$('ctKeysTorch').hidden=!caps.torch;await setZoom();
  cameraNote('Buscando Code 128… Si no lo detecta, cambiará automáticamente al lector alternativo.');
  let nativeReady=false;
  if(typeof BarcodeDetector==='function'){
   try{const formats=await BarcodeDetector.getSupportedFormats();nativeReady=formats.includes('code_128')}catch(_){nativeReady=false}
  }
  if(!nativeReady){await useZxing(video);return}
  let scanner;
  try{scanner=new BarcodeDetector({formats:['code_128']})}catch(_){await useZxing(video);return}
  const nativeStarted=Date.now();
  const loop=async()=>{
   if(!scanning)return;
   if(Date.now()-nativeStarted>7000){await useZxing(video);return}
   try{
    let values=await scanner.detect(video);
    if(values?.some(v=>v.format==='code_128'&&scannedCode(v.rawValue)))return;
    if(scanCtx&&video.videoWidth&&video.videoHeight){
     const w=video.videoWidth,h=video.videoHeight,cw=Math.floor(w*.7),ch=Math.floor(h*.38);
     scanCanvas.width=cw;scanCanvas.height=ch;scanCtx.drawImage(video,(w-cw)/2,(h-ch)/2,cw,ch,0,0,cw,ch);
     values=await scanner.detect(scanCanvas);
     if(values?.some(v=>v.format==='code_128'&&scannedCode(v.rawValue)))return;
    }
   }catch(_){/* continuar al lector alternativo */}
   if(scanning)scanTimer=setTimeout(loop,220);
  };
  loop();
 }catch(e){stopCamera();msg('No fue posible abrir la cámara: '+e.message,true)}
}
// Préstamos históricos: cada evento se conserva incluso al registrar una devolución.
async function loadSummary(){
 try{const list=await rpc('ct_validation_keys_loans_get');const items=Array.isArray(list)?list:[];
 $('ctKeysSummaryCount').textContent=items.length+' préstamo'+(items.length===1?'':'s');
 $('ctKeysSummaryRows').innerHTML=items.length?items.map(r=>`<tr><td>${safe(new Date(r.registrado_el).toLocaleString('es-CO'))}</td><td><strong>${safe(r.placa)}</strong></td><td>${safe(r.prestada_a||'—')}</td><td>${safe(r.usuario||'—')}</td></tr>`).join(''):'<tr><td colspan="4">No hay préstamos registrados.</td></tr>';
 $('ctKeysSummaryDelete').hidden=true;
 }catch(e){$('ctKeysSummaryRows').innerHTML='<tr><td colspan="4">No fue posible consultar el historial.</td></tr>';msg('Error de historial: '+e.message,true)}
}
async function deleteSummary(){if(!editable())return msg('No tienes autorización para borrar préstamos históricos',true);
 if(!confirm('¿Eliminar definitivamente TODO el histórico de préstamos de llaves? Esta acción no se puede deshacer y no modifica el estado actual de las llaves.'))return;
 try{await rpc('ct_validation_keys_loans_clear');await loadSummary();msg('Histórico de préstamos eliminado. Los estados de llaves y facturas se conservaron.')}catch(e){msg('No se pudo borrar el histórico: '+e.message,true)}
}
function openSummary(){$('ctKeysSummaryModal').hidden=false;loadSummary()}
function closeSummary(){$('ctKeysSummaryModal').hidden=true}
function init(){$('ctKeysSummaryOpen').onclick=openSummary;$('ctKeysSummaryClose').onclick=closeSummary;$('ctKeysSummaryRefresh').onclick=loadSummary;$('ctKeysSummaryDelete').onclick=deleteSummary;$('ctKeysSummaryModal').addEventListener('click',e=>{if(e.target===$('ctKeysSummaryModal'))closeSummary()});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('ctKeysSummaryModal').hidden)closeSummary()});$('ctKeysZoom').onchange=setZoom;$('ctKeysTorch').onclick=toggleTorch;$('ctKeysCameraOpen').onclick=openCamera;$('ctKeysCameraClose').onclick=stopCamera;document.addEventListener('visibilitychange',()=>{if(document.hidden)stopCamera()});document.querySelectorAll('[data-key-action]').forEach(b=>b.onclick=()=>setMode(b.dataset.keyAction));$('ctKeysApply').onclick=apply;$('ctKeysPlate').addEventListener('input',e=>e.target.value=e.target.value.toUpperCase());$('ctKeysPlate').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();apply()}});$('ctKeysRefresh').onclick=load;window.addEventListener('ct-validation-keys-ready',load);setMode(mode)}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();