/* v5.2: datasets compartidos, almacenamiento en Supabase. */
window.CTShared={
 async _retry(action,attempts,label){
  let error;
  for(let attempt=1;attempt<=attempts;attempt++){
   try{return await action();}catch(e){error=e;if(attempt===attempts)break;await new Promise(r=>setTimeout(r,300*attempt));}
  }
  throw Error(label+': '+(error?.message||String(error)));
 },
 async save(kind,filename,items,metadata={},onProgress){
  if(!sb||!loginPin)throw Error('Debes iniciar sesión para guardar el reporte');
  const pin=loginPin;const chunkSize=kind==='rutero'?180:350;
  const chunks=[];for(let i=0;i<items.length;i+=chunkSize)chunks.push(items.slice(i,i+chunkSize));
  const begin=await sb.rpc('ct_report_begin',{p_pin:pin,p_kind:kind,p_filename:filename,p_metadata:metadata});
  if(begin.error)throw Error(begin.error.message);
  const batch=begin.data;
  for(let i=0;i<chunks.length;i++){
   await this._retry(async()=>{
    const result=await sb.rpc('ct_report_chunk',{p_pin:pin,p_batch:batch,p_index:i,p_payload:chunks[i]});
    if(result.error)throw Error(result.error.message);
   },3,'guardar bloque '+(i+1));
   if(typeof onProgress==='function' && (i%5===0||i===chunks.length-1))onProgress(i+1,chunks.length,'save');
  }
  const done=await sb.rpc('ct_report_commit',{p_pin:pin,p_batch:batch,p_expected:chunks.length});
  if(done.error)throw Error(done.error.message);
  return chunks.length;
 },
 async load(kind,onProgress){
  if(!sb||!loginPin)return null;
  const pin=loginPin;const info=await sb.rpc('ct_report_info',{p_pin:pin,p_kind:kind});
  if(info.error)throw Error(info.error.message);
  const meta=info.data?.[0];if(!meta)return null;
  const items=[];
  // Un bloque por petición evita respuestas grandes que suelen provocar Failed to fetch.
  for(let start=0;start<meta.chunk_count;start++){
   const parts=await this._retry(async()=>{
    const page=await sb.rpc('ct_report_get_chunks',{p_pin:pin,p_kind:kind,p_start:start,p_limit:1});
    if(page.error)throw Error(page.error.message);
    return page.data||[];
   },3,'recuperar bloque '+(start+1));
   if(parts.length!==1||Number(parts[0].chunk_index)!==start||!Array.isArray(parts[0].payload))
    throw Error('Falta el bloque '+(start+1)+' de '+meta.chunk_count+'. El informe no se ha reemplazado.');
   items.push(...parts[0].payload);
   if(typeof onProgress==='function' && (start%5===0||start===meta.chunk_count-1))onProgress(start+1,meta.chunk_count,'load');
  }
  return {items,filename:meta.filename,metadata:meta.metadata||{},saved_at:meta.saved_at,batch_id:meta.batch_id};
 },
 async clear(kind){
  if(!sb||!loginPin)throw Error('Debes iniciar sesión');
  const r=await sb.rpc('ct_report_clear',{p_pin:loginPin,p_kind:kind});if(r.error)throw Error(r.error.message);
 }
};
