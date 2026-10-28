(()=>{'use strict';
const KEY='ct_sync_ping_v631', name='ct-sync-v631';
let channel=null, last=0, pending=null;
try{if('BroadcastChannel'in window)channel=new BroadcastChannel(name)}catch(e){console.warn('[Sync] BroadcastChannel',e)}
function receive(data){if(!data||data.kind!=='ct-change'||!Number.isFinite(data.at)||Date.now()-data.at>120000)return;
  if(data.at<=last)return;last=data.at;
  window.dispatchEvent(new CustomEvent('ct-external-change',{detail:{module:data.module||'all'}}));
}
if(channel)channel.onmessage=e=>receive(e.data);
window.addEventListener('storage',e=>{if(e.key===KEY&&e.newValue){try{receive(JSON.parse(e.newValue))}catch(_){}}});
window.ctSyncNotify=function(module='all'){
 const msg={kind:'ct-change',module:String(module).slice(0,48),at:Date.now()};
 try{channel?.postMessage(msg)}catch(_){ }
 try{localStorage.setItem(KEY,JSON.stringify(msg))}catch(_){ }
 window.dispatchEvent(new CustomEvent('ct-local-change',{detail:{module:msg.module}}));
};
window.ctSyncVersion='6.31';
})();