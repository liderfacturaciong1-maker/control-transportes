(()=>{'use strict';const $=id=>document.getElementById(id);if(!$('ctKeysPlate'))return;
let mode='LLAVE RECIBIDA', busy=false;
const PLATE_PATTERN=/^[A-Z]{3,6}\d{2,4}(?:-\d{1,2})?$/;
const normalizePlate=v=>{let x=String(v??'').trim().toUpperCase().replace(/[\s\u00a0]+/g,'').replace(/[–—−]/g,'-');if(x && /^[A-Z]/.test(x) && !x.startsWith('CO'))x='CO'+x;return x};
const safe=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const rpc=(name,args={})=>{if(!window.ctValidationKeysRpc)throw Error('Inicia sesión en Validaciones Operación');return window.ctValidationKeysRpc(name,args)};
function editable(){return window.ctValidationKeysCanEdit?.()===true}
function msg(t,bad=false){$('ctKeysMsg').textContent=t;$('ctKeysMsg').style.color=bad?'#b91c1c':'#166534'}
function setMode(m){if(!['LLAVE RECIBIDA','LLAVE PRESTADA','LLAVE ENTREGADA'].includes(m))return;mode=m;document.querySelectorAll('#ctValidKeysPanel [data-key-action]').forEach(b=>{const selected=b.dataset.keyAction===m;b.classList.toggle('selected',selected);b.setAttribute('aria-pressed',String(selected));});$('ctKeysPersonWrap').hidden=m==='LLAVE RECIBIDA';$('ctKeysReceiverWrap').hidden=m!=='LLAVE RECIBIDA';const label=$('ctKeysModeLabel622');if(label)label.textContent='Operación seleccionada: '+(m==='LLAVE PRESTADA'?'PRÉSTAMO DE LLAVES':m==='LLAVE ENTREGADA'?'ENTREGA DE LLAVES':'RECIBO DE LLAVES');$('ctKeysApply').textContent=m==='LLAVE PRESTADA'?'✓ Registrar préstamo':m==='LLAVE ENTREGADA'?'✓ Registrar entrega':'✓ Registrar recibo';}
async function load(){try{$('ctKeysApply').disabled=!editable();$('ctKeysPerson').disabled=!editable();$('ctKeysPlate').disabled=!editable()}catch(e){msg('No se pudo verificar acceso: '+e.message,true)}}
async function apply(){if(busy)return;const placa=normalizePlate($('ctKeysPlate').value);const persona=$('ctKeysPerson').value.trim();if(!PLATE_PATTERN.test(placa))return msg('La lectura no corresponde a una placa. Vuelve a escanear (ej. COJTX912 o LWL645).',true);if(mode!=='LLAVE RECIBIDA'&&!persona)return msg('Escribe el nombre de quien solicita la llave',true);const responsible=$('ctKeysReceiver').value;if(mode==='LLAVE RECIBIDA'&&!responsible)return msg('Selecciona quién recibe la llave',true);if(!editable())return msg('Tu usuario no tiene permiso de modificar Entrega de turno',true);busy=true;$('ctKeysApply').disabled=true;try{const result=await rpc('ct_validation_keys_set',{p_plate:placa,p_state:mode,p_borrower:mode==='LLAVE RECIBIDA'?responsible:persona});msg(`${placa}: ${mode}. Transporte(s): ${result?.transportes||'Sin asignación en Facturas'}`);$('ctKeysPlate').value='';if(mode!=='LLAVE RECIBIDA')$('ctKeysPerson').value='';await load();$('ctKeysPlate').focus()}catch(e){msg('No se guardó: '+e.message,true)}finally{busy=false;$('ctKeysApply').disabled=!editable()}}

// Cámara móvil: API nativa BarcodeDetector, con alternativa ZXing cuando no está disponible.
let receiverOptions=['DIANA','MARIA JOSE','ANGELA'];let adminReceivers=false;
async function loadReceivers(){try{const data=await rpc('ct_validation_key_receivers_get');receiverOptions=Array.isArray(data?.names)?data.names:receiverOptions;adminReceivers=!!data?.admin;const previous=$('ctKeysReceiver').value;$('ctKeysReceiver').innerHTML=receiverOptions.map(n=>`<option value="${safe(n)}">${safe(n)}</option>`).join('');if(receiverOptions.includes(previous))$('ctKeysReceiver').value=previous;$('ctKeysManageReceivers').hidden=!adminReceivers;}catch(e){msg('No se pudieron consultar responsables: '+e.message,true)}}
function renderReceivers(){$('ctKeysReceiverList').innerHTML=receiverOptions.map(n=>`<div class="receiverRow"><strong>${safe(n)}</strong><button type="button" data-remove-receiver="${safe(n)}">Quitar</button></div>`).join('');$('ctKeysReceiverList').querySelectorAll('[data-remove-receiver]').forEach(b=>b.onclick=()=>updateReceiver('remove',b.dataset.removeReceiver))}
async function updateReceiver(action,name){if(!adminReceivers)return;try{await rpc('ct_validation_key_receivers_modify',{p_action:action,p_name:name});await loadReceivers();renderReceivers();$('ctKeysNewReceiver').value='';$('ctKeysManageMsg').textContent='Opciones actualizadas.'}catch(e){$('ctKeysManageMsg').textContent=e.message}}
let scanning=false, scanStream=null,scanTimer=null,zxingReader=null,zxingControls=null,torch=false,lastCandidate='',candidateCount=0,lastCandidateTime=0;
const scanCanvas=document.createElement('canvas');const scanCtx=scanCanvas.getContext('2d',{willReadFrequently:true});
async function setZoom(){const tr=scanStream?.getVideoTracks()?.[0];if(!tr)return;const target=Number($('ctKeysZoom')?.value||1);const caps=tr.getCapabilities?.()||{};if(caps.zoom){try{await tr.applyConstraints({advanced:[{zoom:Math.max(caps.zoom.min,Math.min(caps.zoom.max,target))}]});return}catch(_){}} // Visual zoom only when hardware zoom unavailable
const video=$('ctKeysVideo');video.style.transform=`scale(${Math.min(target,2)})`;video.style.transformOrigin='center center';}
async function toggleTorch(){const tr=scanStream?.getVideoTracks()?.[0],caps=tr?.getCapabilities?.()||{};if(!caps.torch)return;try{torch=!torch;await tr.applyConstraints({advanced:[{torch}]});$('ctKeysTorch').textContent=torch?'🔦 Apagar luz':'🔦 Encender luz'}catch(_){}}

function cameraNote(t){$('ctKeysCameraMsg').textContent=t}
function stopCamera(){scanning=false;lastCandidate='';candidateCount=0;lastCandidateTime=0;if(scanTimer){clearTimeout(scanTimer);scanTimer=null}try{zxingControls?.stop()}catch(_){}zxingControls=null;try{zxingReader?.reset?.()}catch(_){}zxingReader=null;if(scanStream){scanStream.getTracks().forEach(t=>t.stop());scanStream=null}const video=$('ctKeysVideo');try{video.pause()}catch(_){}video.srcObject=null;video.style.transform='';torch=false;$('ctKeysTorch').hidden=true;$('ctKeysCamera').hidden=true}
function scannedCode(raw){
 const plate=normalizePlate(raw);
 if(!PLATE_PATTERN.test(plate)){
  cameraNote('Lectura descartada: '+plate.slice(0,24)+'. El QR debe contener exactamente la placa, por ejemplo COTMX037-1.');return false;
 }
 if(lastCandidate!==plate||Date.now()-lastCandidateTime>3000){lastCandidate=plate;candidateCount=1;lastCandidateTime=Date.now();cameraNote('Placa '+plate+' detectada. Mantén el código estable para confirmar.');return false;}
 candidateCount++;lastCandidateTime=Date.now();
 if(candidateCount<2)return false;
 $('ctKeysPlate').value=normalizePlate(plate);stopCamera();msg('Placa '+plate+' identificada. Verifica la operación y presiona Registrar.');$('ctKeysApply').focus();return true;
}
function loadZxing(){return new Promise((resolve,reject)=>{if(window.ZXingBrowser)return resolve(window.ZXingBrowser);const tag=document.createElement('script');tag.src='https://cdn.jsdelivr.net/npm/@zxing/browser@0.2.1/umd/zxing-browser.min.js';tag.onload=()=>window.ZXingBrowser?resolve(window.ZXingBrowser):reject(Error('No se pudo inicializar el lector alternativo'));tag.onerror=()=>reject(Error('No se pudo cargar ZXing desde la red. Comprueba internet o prueba otro navegador (Chrome en Android).'));document.head.appendChild(tag)})}
async function useZxing(video){
 cameraNote('Activando lector alternativo QR y Code 128…');
 try{
  const ZX=await loadZxing();if(!scanning)return;
  scanStream?.getTracks().forEach(t=>t.stop());scanStream=null;video.srcObject=null;
  const hints=new Map();
  // Code 128 primero; evitamos interpretar códigos EAN o numéricos.
  if(window.ZXing?.DecodeHintType&&window.ZXing?.BarcodeFormat){
   hints.set(window.ZXing.DecodeHintType.POSSIBLE_FORMATS,[window.ZXing.BarcodeFormat.QR_CODE,window.ZXing.BarcodeFormat.CODE_128]);
   hints.set(window.ZXing.DecodeHintType.TRY_HARDER,true);
  }
  zxingReader=new ZX.BrowserMultiFormatReader(hints);
  zxingControls=await zxingReader.decodeFromVideoDevice(undefined,video,(result)=>{if(result)scannedCode(result.getText())});
  cameraNote('Lector QR activo. Mantén el cuadrado QR completo dentro de la cámara, sin brillos.');
 }catch(e){stopCamera();msg('No se activó lector alternativo: '+e.message+'. Usa lector Bluetooth o placa manual.',true)}
}
async function openCamera(){
 if(scanning)return;
 if(!window.isSecureContext||!navigator.mediaDevices?.getUserMedia)return msg('La cámara requiere HTTPS y permisos del navegador.',true);
 if(!editable())return msg('No tienes permiso para registrar movimientos de llaves',true);
 scanning=true;$('ctKeysCamera').hidden=false;cameraNote('Solicitando permiso de cámara…');
 try{
  const stream=await navigator.mediaDevices.getUserMedia({audio:false,video:{facingMode:{ideal:'environment'},width:{ideal:2560},height:{ideal:1440}}});
  if(!scanning){stream.getTracks().forEach(t=>t.stop());return}
  scanStream=stream;try{const track=stream.getVideoTracks()[0],caps=track.getCapabilities?.()||{};const settings={advanced:[]};if((caps.focusMode||[]).includes('continuous'))settings.advanced.push({focusMode:'continuous'});if((caps.exposureMode||[]).includes('continuous'))settings.advanced.push({exposureMode:'continuous'});if(settings.advanced.length)await track.applyConstraints(settings)}catch(_){}const video=$('ctKeysVideo');video.srcObject=stream;await video.play();
  const caps=stream.getVideoTracks()[0]?.getCapabilities?.()||{};$('ctKeysTorch').hidden=!caps.torch;await setZoom();
  cameraNote('Buscando QR de placa (también Code 128)… Centra el QR completo en la cámara, con luz uniforme; mueve el celular lentamente hasta que enfoque. Si no se detecta en 15 s, se intentará el lector alternativo.');
  let nativeReady=false;let supportedFormats=[];
  if(typeof BarcodeDetector==='function'){
   try{const formats=await BarcodeDetector.getSupportedFormats();supportedFormats=formats;nativeReady=formats.includes('qr_code')||formats.includes('code_128')}catch(_){nativeReady=false}
  }
  if(!nativeReady){await useZxing(video);return}
  let scanner;
  try{scanner=new BarcodeDetector({formats:['qr_code','code_128'].filter(f=>supportedFormats.includes(f))})}catch(_){await useZxing(video);return}
  const nativeStarted=Date.now();
  const loop=async()=>{
   if(!scanning)return;
   if(Date.now()-nativeStarted>15000){await useZxing(video);return}
   try{
    let values=await scanner.detect(video);
    if(values?.some(v=>(v.format==='qr_code'||v.format==='code_128')&&scannedCode(v.rawValue)))return;
    if(scanCtx&&video.videoWidth&&video.videoHeight){
     const w=video.videoWidth,h=video.videoHeight,cw=Math.floor(w*.86),ch=Math.floor(h*.50);
     scanCanvas.width=cw;scanCanvas.height=ch;scanCtx.drawImage(video,(w-cw)/2,(h-ch)/2,cw,ch,0,0,cw,ch);
     values=await scanner.detect(scanCanvas);
     if(values?.some(v=>(v.format==='qr_code'||v.format==='code_128')&&scannedCode(v.rawValue)))return;
    }
   }catch(_){/* continuar al lector alternativo */}
   if(scanning)scanTimer=setTimeout(loop,220);
  };
  loop();
 }catch(e){stopCamera();msg('No fue posible abrir la cámara: '+e.message,true)}
}
// Bitácora completa: consulta solo desde Resumen, sin depender de cargas de Excel.
let historyItems=[];
const historyLabel={"LLAVE PRESTADA":"PRÉSTAMO DE LLAVES","LLAVE ENTREGADA":"ENTREGA DE LLAVES","LLAVE RECIBIDA":"RECIBO DE LLAVES"};
function renderSummary(){
 const filter=$('ctKeysHistoryFilter').value;
 const q=$('ctKeysHistorySearch').value.trim().toUpperCase();
 const records=historyItems.filter(r=>(!filter||r.estado===filter)&&(!q||String(r.placa||'').toUpperCase().includes(q)||String(r.prestada_a||'').toUpperCase().includes(q)));
 $('ctKeysSummaryCount').textContent=records.length+' de '+historyItems.length+' movimientos';
 $('ctKeysSummaryRows').innerHTML=records.length?records.map(r=>`<tr><td>${safe(r.registrado_el?new Date(r.registrado_el).toLocaleString('es-CO'):'—')}</td><td><strong>${safe(r.placa)}</strong></td><td><span class="ctKeysHistoryBadge ${r.estado==='LLAVE RECIBIDA'?'rec':r.estado==='LLAVE PRESTADA'?'pre':'ent'}">${safe(historyLabel[r.estado]||r.estado)}</span></td><td>${safe(r.prestada_a||'—')}</td><td>${safe(r.usuario||'—')}</td></tr>`).join(''):'<tr><td colspan="5">No hay movimientos con este filtro.</td></tr>';
}
async function loadSummary(){
 try{const result=await rpc('ct_validation_keys_history_get');historyItems=Array.isArray(result?.items)?result.items:[];
 $('ctKeysSummaryDelete').hidden=result?.admin!==true;renderSummary()}
 catch(e){$('ctKeysSummaryRows').innerHTML='<tr><td colspan="5">No fue posible consultar el historial.</td></tr>';msg('Error de historial: '+e.message,true)}
}
async function deleteSummary(){
 if($('ctKeysSummaryDelete').hidden)return;
 if(!confirm('ATENCIÓN: ¿Borrar TODO el historial de recibos, entregas y préstamos? Este borrado es definitivo. No modifica los estados actuales de las llaves ni las facturas.'))return;
 const word=prompt('Para confirmar escribe BORRAR TODO');if(word!=='BORRAR TODO')return;
 try{const count=await rpc('ct_validation_keys_history_clear',{p_confirmation:'BORRAR TODO'});await loadSummary();msg('Se eliminaron '+count+' movimientos del historial. Estados actuales conservados.')}catch(e){msg('No se pudo borrar: '+e.message,true)}
}
function openSummary(){$('ctKeysSummaryModal').hidden=false;loadSummary()}
function closeSummary(){$('ctKeysSummaryModal').hidden=true}
function init(){$('ctKeysManageReceivers').onclick=()=>{if(!adminReceivers)return;renderReceivers();$('ctKeysManageModal').hidden=false};$('ctKeysCloseReceivers').onclick=()=>{$('ctKeysManageModal').hidden=true};$('ctKeysAddReceiver').onclick=()=>updateReceiver('add',$('ctKeysNewReceiver').value);$('ctKeysSummaryOpen').onclick=openSummary;$('ctKeysSummaryClose').onclick=closeSummary;$('ctKeysSummaryRefresh').onclick=loadSummary;$('ctKeysHistoryFilter').onchange=renderSummary;$('ctKeysHistorySearch').oninput=renderSummary;$('ctKeysSummaryDelete').onclick=deleteSummary;$('ctKeysSummaryModal').addEventListener('click',e=>{if(e.target===$('ctKeysSummaryModal'))closeSummary()});document.addEventListener('keydown',e=>{if(e.key==='Escape'&&!$('ctKeysSummaryModal').hidden)closeSummary()});$('ctKeysZoom').onchange=setZoom;$('ctKeysTorch').onclick=toggleTorch;$('ctKeysCameraOpen').onclick=openCamera;$('ctKeysCameraClose').onclick=stopCamera;document.addEventListener('visibilitychange',()=>{if(document.hidden)stopCamera()});document.querySelectorAll('#ctValidKeysPanel [data-key-action]').forEach(b=>{b.onclick=e=>{e.preventDefault();e.stopPropagation();setMode(b.getAttribute('data-key-action'));};});const collapse=$('ctKeysToggle622');if(collapse)collapse.addEventListener('click',()=>{const panel=$('ctValidKeysPanel');const hidden=panel.classList.toggle('ctKeysIsCollapsed622');const body=$('ctKeysBody622');body.hidden=hidden;body.style.display=hidden?'none':'';collapse.textContent=hidden?'Mostrar ▼':'Ocultar ▲';collapse.setAttribute('aria-expanded',String(!hidden));if(hidden)stopCamera();});$('ctKeysApply').onclick=apply;$('ctKeysPlate').addEventListener('input',e=>{const el=e.target;const previous=el.value;const fixed=normalizePlate(previous);if(previous!==fixed){const pos=el.selectionStart??previous.length;el.value=fixed;try{el.setSelectionRange(Math.max(0,pos+(fixed.length-previous.length)),Math.max(0,pos+(fixed.length-previous.length)))}catch(_){}}});$('ctKeysPlate').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();apply()}});$('ctKeysRefresh').onclick=load;window.addEventListener('ct-validation-keys-ready',()=>{load();loadReceivers()});setMode(mode);$('ctKeysReceiverWrap').hidden=false}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
