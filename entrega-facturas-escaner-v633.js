(()=>{'use strict';
const $=id=>document.getElementById(id), norm=x=>String(x??'').trim().toUpperCase().replace(/\s+/g,'');
let stream=null,reading=false,scanTimer=null,controls=null,reader=null,pending=null,busy=false;
const msg=(s,err=false)=>{const e=$('ctInvoiceScanFeedback');e.textContent=s;e.style.color=err?'#b42318':'#174868'};
// GS1 ejemplo: (415)0000000032319(8020)80088139410006072803.
// AI 8020 puede contener datos adicionales: solo los primeros diez dígitos son el DT para este formato de factura.
function parseCode(raw){const text=String(raw||'').trim();let field=null;
 const m=text.match(/\(8020\)\s*(\d{10,})/);if(m)field=m[1];
 if(!field){const ix=text.indexOf('8020');if(ix>=0){const suffix=text.slice(ix+4).replace(/^[\x1d\s]+/,'');const n=suffix.match(/^\d{10,}/);if(n)field=n[0]}}
 if(!field&&/^\d{10}$/.test(text))field=text;
 if(!field)throw Error('No se identificó un número de transporte válido (AI 8020). Comprueba la lectura.');
 return {transport:field.slice(0,10),raw:text};
}
function stop(){reading=false;clearTimeout(scanTimer);try{controls?.stop?.()}catch{}controls=null;try{reader?.reset?.()}catch{}reader=null;stream?.getTracks().forEach(t=>t.stop());stream=null;$('ctInvoiceScanVideo').srcObject=null;$('ctInvoiceScanCameraPanel').hidden=true}
function acceptCode(value){if(!reading)return;const text=String(value||'').trim();try{parseCode(text)}catch{return} $('ctInvoiceBarcode').value=text;stop();process();}
async function startCamera(){if(!window.ctInvoiceScanApi?.snapshot().write)return msg('Acceso de solo lectura: no puedes entregar facturas.',true);stop();
 try{stream=await navigator.mediaDevices.getUserMedia({audio:false,video:{facingMode:{ideal:'environment'},width:{ideal:1920},height:{ideal:1080}}});const video=$('ctInvoiceScanVideo');video.srcObject=stream;$('ctInvoiceScanCameraPanel').hidden=false;await video.play();reading=true;msg('Enfoca el código de barras completo, incluidos sus extremos.');
 let formats=[];try{formats=await BarcodeDetector.getSupportedFormats()}catch{};
 if(typeof BarcodeDetector!=='undefined'&&formats.includes('code_128')){
  const detector=new BarcodeDetector({formats:['code_128']});let attempts=0;
  const tick=async()=>{if(!reading)return;try{for(const v of await detector.detect(video)){if(v.format==='code_128'){acceptCode(v.rawValue);return}}}catch{}if(++attempts>=28){await fallback(video);return}scanTimer=setTimeout(tick,250)};tick();
 }else await fallback(video);
 }catch(e){msg('No se pudo activar la cámara: '+e.message,true);stop()}}
async function fallback(video){if(!reading)return;try{if(!window.ZXingBrowser){await new Promise((resolve,reject)=>{const script=document.createElement('script');script.src='https://cdn.jsdelivr.net/npm/@zxing/browser@0.2.1/umd/zxing-browser.min.js';script.onload=resolve;script.onerror=()=>reject(Error('Sin conexión con el lector alternativo ZXing'));document.head.appendChild(script)})}if(!reading)return;reader=new window.ZXingBrowser.BrowserMultiFormatReader();controls=await reader.decodeFromVideoElement(video,(result)=>{if(result)acceptCode(result.getText())});}catch(e){msg('El lector alternativo no está disponible: '+e.message,true)}}
function confirmDelivery(title,details,cb){pending=cb;const modal=$('ctInvoiceScanConfirm');$('ctInvoiceScanConfirmTitle').textContent=title;$('ctInvoiceScanConfirmText').textContent=details;modal.showModal()}
async function deliver(row){if(busy)return;busy=true;try{await window.ctInvoiceScanApi.deliver(row.id);msg('✓ FACTURA ENTREGADA · Transporte '+row.transporte+' · Placa '+row.placa);$('ctInvoiceBarcode').value='';$('ctInvoiceBarcode').focus()}catch(e){msg('No se pudo guardar la entrega: '+e.message,true)}finally{busy=false}}
async function process(){if(busy)return;const api=window.ctInvoiceScanApi;if(!api)return msg('Espera a que cargue Validaciones Operación.',true);const state=api.snapshot();if(!state.write)return msg('No tienes permiso para modificar facturas.',true);if(!state.restrictionsOk)return msg('No está disponible el cruce de restricciones. No se permite entregar.',true);
 let parsed;try{parsed=parseCode($('ctInvoiceBarcode').value)}catch(e){return msg(e.message,true)}
 const plate=norm($('ctInvoiceExpectedPlate').value);const candidates=state.rows.filter(r=>!r.entregada&&String(r.transporte).trim()===parsed.transport);const byPlate=plate?state.rows.filter(r=>!r.entregada&&norm(r.placa)===plate):[];
 if(!candidates.length){if(byPlate.length)return confirmDelivery('⚠️ INFORMACIÓN NO COINCIDE, ¿SEGURO DE ENTREGAR?',`Código: ${parsed.transport}. La placa ${plate} existe, pero corresponde al transporte ${byPlate[0].transporte}. Aceptar entregará la factura de la placa verificada, no la indicada por el código.`,()=>deliver(byPlate[0]));return msg('Transporte '+parsed.transport+' no encontrado en las facturas pendientes. No se puede entregar automáticamente.',true)}
 if(candidates.length>1){return msg('El transporte '+parsed.transport+' tiene '+candidates.length+' facturas pendientes. Selecciona la factura manualmente para evitar una entrega incorrecta.',true)}
 const r=candidates[0];if(plate&&norm(r.placa)!==plate)return confirmDelivery('⚠️ INFORMACIÓN NO COINCIDE, ¿SEGURO DE ENTREGAR?',`El código identifica el transporte ${r.transporte}, vinculado en Excel con la placa ${r.placa}; la placa indicada es ${plate}. Aceptar entregará el transporte ${r.transporte}.`,()=>deliver(r));
 if(state.restrictedIds.includes(r.id))return confirmDelivery('⛔ FACTURA EN RESTRICCIÓN, ¿SEGURO DE ENTREGAR?',`Transporte: ${r.transporte}. Placa: ${r.placa}. Se registrará la entrega pese a la restricción activa.`,()=>deliver(r));
 await deliver(r);
}
function init(){const panel=$('ctInvoiceScanner');if(!panel)return;const observer=new MutationObserver(()=>{panel.hidden=$('workspace').hidden});observer.observe($('workspace'),{attributes:true,attributeFilter:['hidden']});
 $('ctInvoiceScanProcess').addEventListener('click',process);$('ctInvoiceBarcode').addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();process()}});$('ctInvoiceScanCamera').addEventListener('click',startCamera);$('ctInvoiceScanStop').addEventListener('click',stop);
 $('ctInvoiceScanCancel').addEventListener('click',()=>{pending=null;$('ctInvoiceScanConfirm').close();msg('Entrega cancelada. No se modificó ninguna factura.')});$('ctInvoiceScanAccept').addEventListener('click',async()=>{const fn=pending;pending=null;$('ctInvoiceScanConfirm').close();if(fn)await fn()});
 document.addEventListener('visibilitychange',()=>{if(document.hidden)stop()});window.addEventListener('pagehide',stop);
}if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
