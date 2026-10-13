/* v5.2: datasets compartidos, almacenamiento en Supabase. */
window.CTShared={
 async save(kind,filename,items,metadata={}){
  if(!sb||!loginPin)throw Error('Debes iniciar sesión para guardar el reporte');
  const pin=loginPin;const chunkSize=kind==='rutero'?180:350;
  const chunks=[];for(let i=0;i<items.length;i+=chunkSize)chunks.push(items.slice(i,i+chunkSize));
  const begin=await sb.rpc('ct_report_begin',{p_pin:pin,p_kind:kind,p_filename:filename,p_metadata:metadata});
  if(begin.error)throw Error(begin.error.message);
  const batch=begin.data;
  for(let i=0;i<chunks.length;i++){
   const result=await sb.rpc('ct_report_chunk',{p_pin:pin,p_batch:batch,p_index:i,p_payload:chunks[i]});
   if(result.error)throw Error('Error al guardar bloque '+(i+1)+'/'+chunks.length+': '+result.error.message);
  }
  const done=await sb.rpc('ct_report_commit',{p_pin:pin,p_batch:batch,p_expected:chunks.length});
  if(done.error)throw Error(done.error.message);
  return chunks.length;
 },
 async load(kind){
  if(!sb||!loginPin)return null;
  const pin=loginPin;const info=await sb.rpc('ct_report_info',{p_pin:pin,p_kind:kind});
  if(info.error)throw Error(info.error.message);
  const meta=info.data?.[0];if(!meta)return null;
  const items=[];
  for(let start=0;start<meta.chunk_count;start+=10){
   const page=await sb.rpc('ct_report_get_chunks',{p_pin:pin,p_kind:kind,p_start:start,p_limit:10});
   if(page.error)throw Error(page.error.message);
   for(const part of page.data||[])items.push(...part.payload);
  }
  return {items,filename:meta.filename,metadata:meta.metadata||{},saved_at:meta.saved_at,batch_id:meta.batch_id};
 },
 async clear(kind){
  if(!sb||!loginPin)throw Error('Debes iniciar sesión');
  const r=await sb.rpc('ct_report_clear',{p_pin:loginPin,p_kind:kind});if(r.error)throw Error(r.error.message);
 }
};
