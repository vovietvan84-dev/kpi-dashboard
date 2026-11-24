// Gỡ phần metadata.bin (SheetJS 0.18.5 không đọc được) trước khi parse
function fixXlsb(XLSX,buf){
  const z=XLSX.CFB.read(buf,{type: (typeof Buffer!=='undefined'&&Buffer.isBuffer&&Buffer.isBuffer(buf))?'buffer':'array'});
  const meta=z.FullPaths.filter(p=>/xl\/metadata\.bin$/i.test(p));
  if(!meta.length) return buf;
  meta.forEach(p=>XLSX.CFB.utils.cfb_del(z,p));
  const ctPath=z.FullPaths.find(p=>/\[Content_Types\]\.xml$/.test(p));
  const ct=XLSX.CFB.find(z,ctPath);
  let s=new TextDecoder().decode(ct.content);
  s=s.replace(/<Override[^>]*metadata\.bin[^>]*\/>/gi,'');
  ct.content=new TextEncoder().encode(s); ct.size=ct.content.length;
  const relPath=z.FullPaths.find(p=>/xl\/_rels\/workbook\.bin\.rels$/i.test(p));
  if(relPath){const r=XLSX.CFB.find(z,relPath); let t=new TextDecoder().decode(r.content); t=t.replace(/<Relationship[^>]*metadata\.bin[^>]*\/>/gi,''); r.content=new TextEncoder().encode(t); r.size=r.content.length;}
  return XLSX.CFB.write(z,{type:'array',fileType:'zip'});
}
if(typeof module!=='undefined') module.exports=fixXlsb;

if(typeof window!=='undefined') window.fixXlsb=fixXlsb;
