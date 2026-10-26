import { get, put } from '@vercel/blob'
const PATH='iruda/sales-transactions.json'
export default async function handler(req,res){
 res.setHeader('Cache-Control','private, no-store')
 if(!process.env.BLOB_READ_WRITE_TOKEN)return res.status(503).json({error:'비공개 저장소 연결이 필요합니다.'})
 try{
  if(req.method==='GET'){
   const blob=await get(PATH,{access:'private',token:process.env.BLOB_READ_WRITE_TOKEN})
   if(!blob)return res.status(200).json({rows:[],updatedAt:null})
   const content=await new Response(blob.stream).text()
   const parsed=JSON.parse(content)
   return res.status(200).json(parsed)
  }
  if(req.method==='PUT'){
   const rows=req.body?.rows
   if(!Array.isArray(rows)||rows.length>50000)return res.status(400).json({error:'거래 데이터 형식 또는 건수가 올바르지 않습니다.'})
   const clean=rows.map(r=>({id:String(r.id||''),date:String(r.date||''),product:String(r.product||''),name:String(r.name||''),amount:Number(r.amount),refund:Number(r.refund),start:String(r.start||''),end:String(r.end||''),status:String(r.status||'')}))
   if(clean.some(r=>!/^\d{4}-\d{2}-\d{2}$/.test(r.date)||!Number.isFinite(r.amount)||!Number.isFinite(r.refund)))return res.status(400).json({error:'거래 날짜 또는 금액에 오류가 있습니다.'})
   const updatedAt=new Date().toISOString()
   await put(PATH,JSON.stringify({rows:clean,updatedAt}),{access:'private',allowOverwrite:true,addRandomSuffix:false,contentType:'application/json',token:process.env.BLOB_READ_WRITE_TOKEN})
   return res.status(200).json({ok:true,count:clean.length,updatedAt})
  }
  return res.status(405).json({error:'지원하지 않는 요청입니다.'})
 }catch(e){console.error('sales store',e);return res.status(500).json({error:'저장소 처리에 실패했습니다.'})}
}
