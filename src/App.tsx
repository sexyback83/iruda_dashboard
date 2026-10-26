import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import './index.css'

type Tx = { id:string; date:string; product:string; name:string; amount:number; refund:number; start:string; end:string; status:string }
type Issue = { level:'오류'|'주의'; row:number; message:string }
const CATEGORIES = ['1회시간권','정액시간권','기간권','스터디룸','사물함','기타']
const FIELDS: Record<keyof Tx,string[]> = {
 id:['No','거래번호','주문번호','결제번호','transaction_id','id'],
 date:['주문일시','결제일시','결제일','거래일','날짜','date'],
 product:['구분','매출분류','상품분류','분류','product','category'],
 name:['주문명','상품명','이용권명','name'],
 amount:['합계금액','실제 결제금액','결제금액','매출액','금액','amount'],
 refund:['환불금액','취소금액','refund'],
 start:['이용시작일','시작일','start'],
 end:['이용종료일','종료일','end'],
 status:['주문상태','거래상태','상태','status']
}
const money=(n:number)=>Math.round(n).toLocaleString('ko-KR')+'원'
const dateText=(value:unknown):string=>{
 if(value instanceof Date) return value.toISOString().slice(0,10)
 if(typeof value==='number' && value>20000 && value<100000){const d=XLSX.SSF.parse_date_code(value);return d?`${d.y}-${String(d.m).padStart(2,'0')}-${String(d.d).padStart(2,'0')}`:''}
 const s=String(value??'').trim();if(!s)return ''
 const m=s.match(/^(\d{4})[.\-/년 ]+(\d{1,2})[.\-/월 ]+(\d{1,2})/)
 if(m)return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`
 return s.slice(0,10)
}
const validDate=(s:string)=>{if(!/^\d{4}-\d{2}-\d{2}$/.test(s))return false;const d=new Date(s+'T00:00:00');return !isNaN(d.getTime())&&d.toISOString().slice(0,10)===s}
const numeric=(v:unknown)=>{if(typeof v==='number')return Number.isFinite(v)?v:NaN;const s=String(v??'').replace(/[,₩원\s]/g,'');return s===''?0:Number(s)}
const category=(s:string)=>CATEGORIES.find(x=>s.includes(x))??(s.includes('일일')?'1회시간권':s.includes('정액')?'정액시간권':s.includes('고정')?'기간권':'기타')
const STORAGE='iruda-dashboard-v3'
function load():Tx[]{try{const a=JSON.parse(localStorage.getItem(STORAGE)||'[]');return Array.isArray(a)?a:[]}catch{return []}}
function validate(rows:Tx[]):Issue[]{
 const seen=new Set<string>();const issues:Issue[]=[]
 rows.forEach((r,i)=>{const row=i+2
 if(!r.id)issues.push({level:'주의',row,message:'거래번호 누락'})
 else if(seen.has(r.id))issues.push({level:'오류',row,message:'중복 거래번호: '+r.id})
 seen.add(r.id)
 if(!validDate(r.date))issues.push({level:'오류',row,message:'결제일 오류'})
 if(!Number.isFinite(r.amount)||!Number.isFinite(r.refund)||r.amount<0||r.refund<0)issues.push({level:'오류',row,message:'금액 형식 오류'})
 if(r.status==='환불'&&r.refund===0)issues.push({level:'주의',row,message:'환불 거래 금액을 확인해 주세요'})
 if(r.status==='환불'&&r.refund===r.amount){} // 환불 행의 금액은 환불액으로 간주
 if(r.refund>r.amount)issues.push({level:'주의',row,message:'환불금액이 결제금액보다 큼 — 별도 환불 거래인지 확인'})
 if(r.start&&!validDate(r.start)||r.end&&!validDate(r.end))issues.push({level:'오류',row,message:'이용기간 날짜 형식 오류'})
 if(r.start&&r.end&&validDate(r.start)&&validDate(r.end)&&r.start>r.end)issues.push({level:'오류',row,message:'이용종료일이 시작일보다 빠름'})
 if(r.product==='기타')issues.push({level:'주의',row,message:'상품 분류 확인 필요'})
 })
 return issues
}
function exportExcel(rows:object[],name:string){const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.json_to_sheet(rows),'매출');XLSX.writeFile(book,name+'.xlsx')}
function App(){
 const [tab,setTab]=useState('통합 대시보드')
 const [rows,setRows]=useState<Tx[]>(load)
 const [cloudStatus,setCloudStatus]=useState('서버 데이터 확인 중')
 const [serverUpdated,setServerUpdated]=useState<string|null>(null)
 const [draft,setDraft]=useState<Record<string,unknown>[]>([])
 const [mapping,setMapping]=useState<Record<string,string>>({})
 const [fileName,setFileName]=useState('')
 const [message,setMessage]=useState('')
 const [year,setYear]=useState('전체')
 const [month,setMonth]=useState('전체')
 const [product,setProduct]=useState('전체')
 const [basis,setBasis]=useState<'결제일'|'이용기간'>('결제일')
 
 const [search,setSearch]=useState('')
 const [showRefunds,setShowRefunds]=useState(true)
 const [customCategories,setCustomCategories]=useState<string[]>(()=>{try{return JSON.parse(localStorage.getItem('iruda-categories')||'[]')}catch{return []}})
 const [newCategory,setNewCategory]=useState('')
 const [showMenu,setShowMenu]=useState(false)
 useEffect(()=>{try{localStorage.setItem(STORAGE,JSON.stringify(rows))}catch{setMessage('브라우저 저장 공간이 부족합니다. 데이터를 엑셀로 백업해 주세요.')}},[rows])
 useEffect(()=>{let active=true;fetch('/api/sales',{credentials:'same-origin',cache:'no-store'}).then(async response=>{if(!response.ok)throw Error('서버 응답 '+response.status);return response.json()}).then(data=>{if(!active)return;if(Array.isArray(data.rows)&&data.rows.length){setRows(data.rows);setServerUpdated(data.updatedAt||null);setCloudStatus('서버 데이터 연결됨')}else setCloudStatus('서버 데이터 없음 · 첫 업로드 필요')}).catch(()=>{if(active)setCloudStatus('서버 연결 실패 · 브라우저 데이터 사용')});return()=>{active=false}},[])
 const saveCloud=async(data:Tx[])=>{setCloudStatus('비공개 서버 저장 중');try{const response=await fetch('/api/sales',{method:'PUT',headers:{'Content-Type':'application/json'},credentials:'same-origin',body:JSON.stringify({rows:data})});const result=await response.json();if(!response.ok)throw Error(result.error||'서버 오류');setServerUpdated(result.updatedAt);setCloudStatus('서버 데이터 연결됨');setMessage('비공개 서버에 '+data.length.toLocaleString()+'건을 저장했습니다. 다른 기기에서도 자동으로 불러옵니다.')}catch(e){setCloudStatus('서버 저장 실패');setMessage('로컬 데이터는 유지되었습니다. 서버 저장 실패: '+String(e))}}
 useEffect(()=>{localStorage.setItem('iruda-categories',JSON.stringify(customCategories))},[customCategories])
 const issues=useMemo(()=>validate(rows),[rows])
 const valid=useMemo(()=>rows.filter(r=>validDate(r.date)&&Number.isFinite(r.amount)&&Number.isFinite(r.refund)&&r.amount>=0&&r.refund>=0),[rows])
 const years=useMemo(()=>[...new Set(valid.map(r=>r.date.slice(0,4)))].sort().reverse(),[valid])
 const expanded=useMemo(()=>{
  const out:{date:string;product:string;amount:number;id:string}[]=[]
  for(const r of valid){
   const amount=r.status==='환불'?(showRefunds?-Math.abs(r.refund||r.amount):0):r.amount-r.refund
   if(basis==='이용기간'&&validDate(r.start)&&validDate(r.end)&&r.start<=r.end){
    const from=new Date(r.start+'T00:00:00Z').getTime(),to=new Date(r.end+'T00:00:00Z').getTime()
    const count=Math.round((to-from)/86400000)+1
    if(count>0&&count<=3660){for(let i=0;i<count;i++){const date=new Date(from+i*86400000).toISOString().slice(0,10);out.push({date,product:r.product,amount:amount/count,id:r.id})}continue}
   }
   out.push({date:r.date,product:r.product,amount,id:r.id})
  }
  return out
 },[valid,basis,showRefunds])
 const filtered=useMemo(()=>expanded.filter(r=>(year==='전체'||r.date.startsWith(year))&&(month==='전체'||r.date.slice(5,7)===month)&&(product==='전체'||r.product===product)),[expanded,year,month,product])
 const sum=(a:{amount:number}[])=>a.reduce((n,r)=>n+r.amount,0)
 const total=sum(filtered)
 const completed=valid.filter(r=>r.status!=='환불')
 const refunds=valid.filter(r=>r.status==='환불')
 const refundTotal=refunds.reduce((a,r)=>a+(r.refund||r.amount),0)
 const grossTotal=completed.reduce((a,r)=>a+r.amount,0)
 const group=(mode:'day'|'month'|'year'|'product')=>{
  const map=new Map<string,{amount:number;count:Set<string>}>()
  for(const r of filtered){const k=mode==='day'?r.date:mode==='month'?r.date.slice(0,7):mode==='year'?r.date.slice(0,4):r.product
   const x=map.get(k)||{amount:0,count:new Set<string>()};x.amount+=r.amount;x.count.add(r.id);map.set(k,x)}
  return [...map].map(([label,v])=>({label,amount:v.amount,count:v.count.size})).sort((a,b)=>a.label.localeCompare(b.label,'ko'))
 }
 const daily=group('day'),monthly=group('month'),yearly=group('year'),products=group('product')
 const current=tab==='일별 매출 분석'?daily:tab==='월별 매출 분석'?monthly:tab==='연도별 매출 분석'?yearly:tab==='상품별 매출 분석'?products:tab==='이용기간 매출 분석'?monthly:monthly
 const columns=draft.length?Object.keys(draft[0]):[]
 const mapped=(field:keyof Tx)=>(mapping[field]||'')
 const convert=():Tx[]=>draft.map((r,i)=>({id:String(r[mapped('id')]??'').trim()||'ROW-'+String(i+2),date:dateText(r[mapped('date')]),product:category(String(r[mapped('product')]??r[mapped('name')]??'')),name:String(r[mapped('name')]??''),amount:numeric(r[mapped('amount')]),refund:String(r[mapped('status')]??'').trim()==='환불'?numeric(r[mapped('amount')]):numeric(r[mapped('refund')]),start:dateText(r[mapped('start')]),end:dateText(r[mapped('end')]),status:String(r[mapped('status')]??'')}))
 const read=async(file?:File)=>{if(!file)return;try{
  if(file.size>15*1024*1024)throw Error('파일 크기는 15MB 이하로 제한됩니다.')
  const book=XLSX.read(await file.arrayBuffer(),{type:'array',cellDates:true})
  const data=XLSX.utils.sheet_to_json<Record<string,unknown>>(book.Sheets[book.SheetNames[0]],{defval:''})
  if(!data.length)throw Error('파일에 데이터가 없습니다.')
  const keys=Object.keys(data[0]);const map:Record<string,string>={}
  for(const [field,aliases] of Object.entries(FIELDS)){map[field]=keys.find(k=>aliases.some(a=>a.toLowerCase()===k.trim().toLowerCase()))||''}
  setDraft(data);setMapping(map);setFileName(file.name);setTab('매출 데이터 관리');setMessage(`${data.length.toLocaleString()}행을 읽었습니다. 열 연결을 확인한 후 가져오세요.`)
 }catch(e){setMessage('파일 읽기 실패: '+String(e))}}
 const importRows=()=>{if(!mapped('date')||!mapped('amount')){setMessage('결제일과 실제 결제금액 열은 반드시 연결해야 합니다.');return}
 const converted=convert();const errs=validate(converted).filter(x=>x.level==='오류')
 if(errs.length){setMessage(`가져오기 중단: 오류 ${errs.length}건. 데이터 검증 화면에서 확인하기 전에 원본 파일을 수정해 주세요. 첫 오류: ${errs[0].message} (${errs[0].row}행)`);return}
 setRows(converted);setDraft([]);setMessage(`${converted.length.toLocaleString()}건의 거래를 적용했습니다. 서버 저장을 진행합니다.`);void saveCloud(converted)}
 const navigation=['통합 대시보드','일별 매출 분석','월별 매출 분석','연도별 매출 분석','상품별 매출 분석','이용기간 매출 분석','매출 데이터 관리','데이터 검증','설정']
 const table=(items:{label:string;amount:number;count:number}[])=> <div className="panel" style={{padding:20,overflowX:'auto',marginTop:16}}><table style={{width:'100%',borderCollapse:'collapse',fontSize:13}}><thead><tr><th style={{textAlign:'left',padding:12}}>기간 / 분류</th><th style={{textAlign:'right'}}>순매출</th><th style={{textAlign:'right'}}>거래 건수</th><th style={{textAlign:'right'}}>비중</th></tr></thead><tbody>{items.map(r=><tr key={r.label} style={{borderTop:'1px solid #e2e8f0'}}><td style={{padding:12}}>{r.label}</td><td style={{textAlign:'right'}}>{money(r.amount)}</td><td style={{textAlign:'right'}}>{r.count.toLocaleString()}건</td><td style={{textAlign:'right'}}>{total?((r.amount/total)*100).toFixed(1):'0.0'}%</td></tr>)}</tbody></table>{!items.length&&<p style={{padding:16,color:'#64748b'}}>조회할 거래가 없습니다. 매출 데이터를 업로드해 주세요.</p>}</div>
 const bars=(items:{label:string;amount:number;count:number}[])=>{const max=Math.max(1,...items.map(x=>Math.abs(x.amount)));return <div className="panel" style={{padding:20,marginTop:16}}><h2 style={{margin:'0 0 18px',fontSize:17}}>매출 추이</h2><div style={{display:'flex',alignItems:'end',gap:8,overflowX:'auto',height:225}}>{items.map(x=><div key={x.label} title={x.label+' '+money(x.amount)} style={{flex:'1 0 35px',minWidth:35,textAlign:'center'}}><div style={{height:175,display:'flex',alignItems:'end',justifyContent:'center'}}><div style={{height:Math.max(2,Math.abs(x.amount)/max*165),background:x.amount<0?'#dc2626':'#2563eb',borderRadius:'5px 5px 0 0',width:'65%'}}/></div><small style={{fontSize:10,whiteSpace:'nowrap'}}>{x.label.slice(-7)}</small></div>)}</div></div>}
 const filters=<section className="filter-card" style={{flexWrap:'wrap',marginBottom:18}}><div className="filter-group"><span className="filter-label">연도</span><select value={year} onChange={e=>setYear(e.target.value)}><option>전체</option>{years.map(y=><option key={y}>{y}</option>)}</select></div><div className="filter-group"><span className="filter-label">월</span><select value={month} onChange={e=>setMonth(e.target.value)}><option>전체</option>{Array.from({length:12},(_,i)=><option key={i} value={String(i+1).padStart(2,'0')}>{i+1}월</option>)}</select></div><div className="filter-group"><span className="filter-label">상품</span><select value={product} onChange={e=>setProduct(e.target.value)}><option>전체</option>{[...new Set([...CATEGORIES,...customCategories,...rows.map(r=>r.product)])].map(p=><option key={p}>{p}</option>)}</select></div><div className="filter-group"><span className="filter-label">집계 기준</span><select value={basis} onChange={e=>setBasis(e.target.value as '결제일'|'이용기간')}><option>결제일</option><option>이용기간</option></select></div><label style={{fontSize:12,display:'flex',alignItems:'center',gap:6}}><input type="checkbox" checked={showRefunds} onChange={e=>setShowRefunds(e.target.checked)}/>환불 차감</label><button className="button" onClick={()=>{setYear('전체');setMonth('전체');setProduct('전체');setBasis('결제일');setShowRefunds(true)}}>초기화</button></section>
 return <div className="app-shell"><aside className={`sidebar ${showMenu?'sidebar-open':''}`}><div className="brand"><div className="brand-mark"><span/><span/><span/></div><div><strong>이루다</strong><small>STUDY CAFE</small></div></div><nav className="nav">{navigation.map((item,i)=><button key={item} className={`nav-item ${tab===item?'active':''}`} onClick={()=>{setTab(item);setShowMenu(false)}}><span>{i+1}. {item}</span></button>)}</nav><div className="sidebar-foot"><div className="sample-dot"/><div><strong>{rows.length?'업로드된 실제 거래':'거래 데이터 없음'}</strong><small>{rows.length.toLocaleString()}건 · {cloudStatus}</small></div></div></aside><div className="content-shell"><header className="mobile-header"><button className="button" onClick={()=>setShowMenu(!showMenu)}>☰ 메뉴</button><strong>이루다</strong></header><main className="main-content"><section className="page-heading"><div><div className="eyebrow">IRUDA SALES ANALYTICS · 실제 거래 기반</div><h1>{tab}</h1><p>이루다 수지구청점 결제로그 기준 · 1회시간권 / 정액시간권 / 기간권 / 스터디룸 / 사물함 · 실제 결제 데이터만 집계합니다.</p><p style={{fontSize:12,color:'#64748b'}}>비공개 저장소: {cloudStatus}{serverUpdated?' · 최종 저장 '+new Date(serverUpdated).toLocaleString('ko-KR'):''}</p></div><div className="heading-actions"><button className="button" onClick={()=>setTab('매출 데이터 관리')}>엑셀 업로드</button><button className="button button-primary" onClick={()=>exportExcel(filtered,'이루다_매출분석')}>결과 다운로드</button></div></section>{message&&<div className="info-banner" role="status" onClick={()=>setMessage('')}>{message} (눌러서 닫기)</div>}
 {tab==='매출 데이터 관리'?<><label className="upload-zone" onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();void read(e.dataTransfer.files[0])}}><input type="file" accept=".xlsx,.xls,.csv" onChange={e=>void read(e.target.files?.[0])}/><strong>엑셀 또는 CSV 파일 선택 / 끌어놓기</strong><p>이루다 결제로그(15열) 자동 인식 · 환불 행은 음수 매출 처리 · 원본 파일은 서버로 전송하지 않습니다.</p><b>파일 선택</b></label>{draft.length>0&&<section className="panel" style={{padding:20,marginTop:16}}><h2>열 연결 확인 · {fileName} · {draft.length}행</h2><p>결제일과 결제금액은 필수입니다. 이루다 결제로그의 주문상태=환불 행은 차감 거래로 처리합니다. 잘못 연결된 열을 직접 수정하세요.</p><div style={{display:'grid',gridTemplateColumns:'repeat(auto-fit,minmax(190px,1fr))',gap:12}}>{Object.entries(FIELDS).map(([field,aliases])=><label key={field} style={{fontSize:12}}>{aliases[0]} {['date','amount'].includes(field)?'*':''}<select style={{display:'block',width:'100%',padding:8,marginTop:5}} value={mapping[field]||''} onChange={e=>setMapping(m=>({...m,[field]:e.target.value}))}><option value="">연결 안 함</option>{columns.map(c=><option key={c}>{c}</option>)}</select></label>)}</div><div style={{marginTop:16,display:'flex',gap:10}}><button className="button button-primary" onClick={importRows}>검증 후 {draft.length}건 적용 (기존 자료 교체)</button><button className="button" onClick={()=>{setDraft([]);setMessage('가져오기를 취소했습니다.')}}>취소</button></div><div className="preview-scroll" style={{marginTop:16}}><table><thead><tr>{columns.slice(0,8).map(c=><th key={c}>{c}</th>)}</tr></thead><tbody>{draft.slice(0,5).map((r,i)=><tr key={i}>{columns.slice(0,8).map(c=><td key={c}>{String(r[c]??'')}</td>)}</tr>)}</tbody></table></div></section>}{rows.length>0&&<section className="panel" style={{padding:20,marginTop:16}}><h2>저장된 거래 {rows.length.toLocaleString()}건</h2><p>비공개 Vercel Blob 저장소에 동기화합니다. 최초 업로드 후 다른 기기에서도 관리자 로그인으로 조회할 수 있습니다.</p><button className="button" onClick={()=>exportExcel(rows,'이루다_원본백업')}>전체 데이터 백업</button> <button className="button" onClick={()=>void saveCloud(rows)}>비공개 서버에 다시 저장</button> <button className="button" onClick={()=>{if(window.confirm('브라우저의 거래를 비웁니다. 서버 데이터는 유지됩니다. 계속하시겠습니까?')){setRows([]);setMessage('브라우저 데이터를 비웠습니다. 서버에 저장된 거래는 유지됩니다.')}}}>전체 삭제</button></section>}</>:
 tab==='데이터 검증'?<><section className="metrics analysis-metrics"><article className="metric-card"><p>전체 거래</p><strong>{rows.length}건</strong></article><article className="metric-card"><p>오류</p><strong>{issues.filter(x=>x.level==='오류').length}건</strong></article><article className="metric-card"><p>주의</p><strong>{issues.filter(x=>x.level==='주의').length}건</strong></article></section><div className="panel" style={{padding:20,marginTop:16}}><h2>실제 데이터 검증 결과</h2>{issues.length?issues.map((x,i)=><div key={i} style={{padding:10,borderBottom:'1px solid #eee'}}><strong>{x.level}</strong> · 원본 {x.row}행 · {x.message}</div>):<p>검증 오류가 없습니다. 데이터가 없으면 업로드 후 검사하세요.</p>}<button className="button" onClick={()=>exportExcel(issues,'이루다_검증결과')}>검증 결과 다운로드</button></div></>:
 tab==='설정'?<section className="panel" style={{padding:24}}><h2>상품 분류 설정</h2><p>추가 분류는 현재 브라우저에 저장되며, 업로드할 때 상품명이 자동 분류되지 않는 항목은 ‘기타’로 표시됩니다.</p><div style={{display:'flex',gap:8,flexWrap:'wrap'}}>{[...CATEGORIES,...customCategories].map(c=><span key={c} style={{padding:8,background:'#eff6ff',borderRadius:6}}>{c} {customCategories.includes(c)&&<button onClick={()=>setCustomCategories(a=>a.filter(x=>x!==c))}>×</button>}</span>)}</div><div style={{marginTop:16}}><input value={newCategory} onChange={e=>setNewCategory(e.target.value)} placeholder="새 상품 분류" style={{padding:10,border:'1px solid #ddd'}}/> <button className="button" onClick={()=>{const v=newCategory.trim();if(v&&!CATEGORIES.includes(v)&&!customCategories.includes(v))setCustomCategories(a=>[...a,v]);setNewCategory('')}}>추가</button></div><p style={{marginTop:20}}>접근 제어: Vercel 프로젝트의 인증 설정에서 관리합니다. 이 앱은 자체 비밀번호를 저장하지 않습니다.</p></section>:
 <>{filters}{basis==='이용기간'&&<div className="info-banner">이용 시작일~종료일에 금액을 일할 배분합니다. 기간 정보가 없으면 결제일에 반영합니다. 내부 경영분석용이며 회계상 매출 인식과 다를 수 있습니다.</div>}<section className="metrics analysis-metrics"><article className="metric-card"><p>순매출</p><strong>{money(total)}</strong><div className="metric-detail">환불 차감 {showRefunds?'적용':'해제'} · 선택한 기간 및 상품 기준</div></article><article className="metric-card"><p>결제완료 총액</p><strong>{money(grossTotal)}</strong><div className="metric-detail">전체 업로드 데이터 기준</div></article><article className="metric-card"><p>환불 총액</p><strong>{money(refundTotal)}</strong><div className="metric-detail">환불 {refunds.length.toLocaleString()}건 · 전체 업로드 데이터 기준</div></article><article className="metric-card"><p>거래 건수</p><strong>{new Set(filtered.map(x=>x.id)).size.toLocaleString()}건</strong></article><article className="metric-card"><p>일평균 매출</p><strong>{money(daily.length?total/daily.length:0)}</strong><div className="metric-detail">매출 발생일 기준</div></article><article className="metric-card"><p>검증 이슈</p><strong>{issues.length}건</strong><button className="button button-ghost" onClick={()=>setTab('데이터 검증')}>확인</button></article></section>{tab==='통합 대시보드'?<>{bars(monthly)}{table(products)}<div className="panel" style={{padding:20,marginTop:16}}><h2>최근 거래</h2><input placeholder="거래번호·상품명 검색" value={search} onChange={e=>setSearch(e.target.value)} style={{padding:10,border:'1px solid #ddd',width:'100%',maxWidth:320}}/><div className="preview-scroll"><table><thead><tr>{['거래번호','결제일','상품','결제금액','환불금액','순매출'].map(x=><th key={x}>{x}</th>)}</tr></thead><tbody>{valid.filter(r=>(r.id+' '+r.name).includes(search)).slice(-30).reverse().map((r,i)=><tr key={r.id+i}><td>{r.id}</td><td>{r.date}</td><td>{r.product} · {r.name}</td><td>{money(r.amount)}</td><td>{money(r.refund)}</td><td>{money(r.status==='환불'?-(r.refund||r.amount):r.amount-r.refund)}</td></tr>)}</tbody></table></div></div></>:<>{bars(current)}{table(current)}</>}</>}
 </main></div></div>
}
export default App
