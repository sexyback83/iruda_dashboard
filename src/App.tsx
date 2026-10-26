import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'

type Product = '시간권' | '기간권' | '고정석' | '일일권' | '사물함' | '스터디룸'
type IconName =
  | 'dashboard'
  | 'calendar'
  | 'chart'
  | 'trend'
  | 'box'
  | 'clock'
  | 'database'
  | 'check'
  | 'settings'
  | 'upload'
  | 'refresh'
  | 'chevron'
  | 'arrowUp'
  | 'arrowDown'
  | 'info'
  | 'menu'
  | 'close'

const products: Product[] = ['시간권', '기간권', '고정석', '일일권', '사물함', '스터디룸']
const productColors: Record<Product, string> = {
  시간권: '#2563eb',
  기간권: '#7c3aed',
  고정석: '#0891b2',
  일일권: '#16a34a',
  사물함: '#f59e0b',
  스터디룸: '#e11d48',
}

const navItems: { label: string; icon: IconName }[] = [
  { label: '통합 대시보드', icon: 'dashboard' },
  { label: '일별 매출 분석', icon: 'calendar' },
  { label: '월별 매출 분석', icon: 'chart' },
  { label: '연도별 매출 분석', icon: 'trend' },
  { label: '상품별 매출 분석', icon: 'box' },
  { label: '이용기간 매출 분석', icon: 'clock' },
  { label: '매출 데이터 관리', icon: 'database' },
  { label: '데이터 검증', icon: 'check' },
  { label: '설정', icon: 'settings' },
]

const baseMonthly: Record<Product, number[]> = {
  시간권: [22, 24, 27, 25, 30, 34, 31, 36, 38, 41, 0, 0],
  기간권: [34, 37, 41, 39, 45, 47, 43, 49, 52, 57, 0, 0],
  고정석: [18, 20, 22, 21, 24, 26, 25, 27, 30, 32, 0, 0],
  일일권: [13, 14, 17, 16, 18, 21, 23, 25, 24, 26, 0, 0],
  사물함: [8, 9, 9, 10, 11, 11, 12, 13, 14, 14, 0, 0],
  스터디룸: [11, 13, 14, 13, 16, 18, 17, 20, 22, 24, 0, 0],
}

function Icon({ name, size = 18 }: { name: IconName; size?: number }) {
  const paths: Record<IconName, React.ReactNode> = {
    dashboard: <><rect x="3" y="3" width="7" height="7" rx="2" /><rect x="14" y="3" width="7" height="7" rx="2" /><rect x="3" y="14" width="7" height="7" rx="2" /><rect x="14" y="14" width="7" height="7" rx="2" /></>,
    calendar: <><path d="M7 2v4m10-4v4M3 9h18" /><rect x="3" y="4" width="18" height="18" rx="3" /><path d="M8 13h.01M12 13h.01M16 13h.01M8 17h.01M12 17h.01" /></>,
    chart: <><path d="M4 20V10m6 10V4m6 16v-7m4 7H2" /></>,
    trend: <><path d="m3 17 6-6 4 4 8-9" /><path d="M15 6h6v6" /></>,
    box: <><path d="m4 7 8-4 8 4-8 4-8-4Z" /><path d="m4 7 8 4 8-4v10l-8 4-8-4V7Z" /><path d="M12 11v10" /></>,
    clock: <><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></>,
    database: <><ellipse cx="12" cy="5" rx="8" ry="3" /><path d="M4 5v7c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12v7c0 1.7 3.6 3 8 3s8-1.3 8-3v-7" /></>,
    check: <><path d="m9 11 2 2 4-4" /><rect x="3" y="3" width="18" height="18" rx="3" /><path d="M8 17h8" /></>,
    settings: <><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.9l.1.1-2.8 2.8-.1-.1a1.7 1.7 0 0 0-1.9-.3 1.7 1.7 0 0 0-1 1.6v.2h-4V21a1.7 1.7 0 0 0-1-1.6 1.7 1.7 0 0 0-1.9.3l-.1.1L4.2 17l.1-.1a1.7 1.7 0 0 0 .3-1.9A1.7 1.7 0 0 0 3 14H2.8v-4H3a1.7 1.7 0 0 0 1.6-1 1.7 1.7 0 0 0-.3-1.9L4.2 7 7 4.2l.1.1A1.7 1.7 0 0 0 9 4.6 1.7 1.7 0 0 0 10 3V2.8h4V3a1.7 1.7 0 0 0 1 1.6 1.7 1.7 0 0 0 1.9-.3l.1-.1L19.8 7l-.1.1a1.7 1.7 0 0 0-.3 1.9 1.7 1.7 0 0 0 1.6 1h.2v4H21a1.7 1.7 0 0 0-1.6 1Z" /></>,
    upload: <><path d="M12 16V4m0 0L7 9m5-5 5 5" /><path d="M5 15v4h14v-4" /></>,
    refresh: <><path d="M20 7v5h-5M4 17v-5h5" /><path d="M6.1 8a7 7 0 0 1 11.7-1.8L20 9M4 15l2.2 2.8A7 7 0 0 0 18 16" /></>,
    chevron: <path d="m8 10 4 4 4-4" />,
    arrowUp: <><path d="m7 14 5-5 5 5" /><path d="M12 19V9" /></>,
    arrowDown: <><path d="m7 10 5 5 5-5" /><path d="M12 5v10" /></>,
    info: <><circle cx="12" cy="12" r="9" /><path d="M12 11v6M12 7h.01" /></>,
    menu: <path d="M4 7h16M4 12h16M4 17h16" />,
    close: <path d="m6 6 12 12M18 6 6 18" />,
  }
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{paths[name]}</svg>
}

function Button({ children, variant = 'default', className = '', onClick }: { children: React.ReactNode; variant?: 'default' | 'primary' | 'ghost'; className?: string; onClick?: () => void }) {
  return <button className={`button button-${variant} ${className}`} onClick={onClick}>{children}</button>
}

function Select({ value, onChange, children, label }: { value: string; onChange: (value: string) => void; children: React.ReactNode; label: string }) {
  return <label className="select-wrap"><span className="sr-only">{label}</span><select value={value} onChange={(event) => onChange(event.target.value)}>{children}</select><Icon name="chevron" size={15} /></label>
}

function Sidebar({ active, onChange, open, onClose }: { active: string; onChange: (value: string) => void; open: boolean; onClose: () => void }) {
  return <aside className={`sidebar ${open ? 'sidebar-open' : ''}`}>
    <div className="brand">
      <div className="brand-mark"><span /><span /><span /></div>
      <div><strong>이루다</strong><small>STUDY CAFE</small></div>
      <Button variant="ghost" className="sidebar-close" onClick={onClose}><Icon name="close" /></Button>
    </div>
    <nav className="nav">
      <p className="nav-label">매출 분석</p>
      {navItems.slice(0, 6).map((item) => <button key={item.label} className={`nav-item ${active === item.label ? 'active' : ''}`} onClick={() => { onChange(item.label); onClose() }}><Icon name={item.icon} /><span>{item.label}</span></button>)}
      <p className="nav-label nav-label-secondary">데이터 및 관리</p>
      {navItems.slice(6).map((item) => <button key={item.label} className={`nav-item ${active === item.label ? 'active' : ''}`} onClick={() => { onChange(item.label); onClose() }}><Icon name={item.icon} /><span>{item.label}</span>{item.label === '데이터 검증' && <b className="nav-badge">3</b>}</button>)}
    </nav>
    <div className="sidebar-foot">
      <div className="sample-dot" />
      <div><strong>샘플 데이터 사용 중</strong><small>2026년 1월–10월</small></div>
    </div>
  </aside>
}

function MetricCard({ label, value, detail, change, tone = 'neutral' }: { label: string; value: string; detail: string; change?: string; tone?: 'up' | 'down' | 'neutral' }) {
  return <article className="metric-card">
    <p>{label}</p>
    <strong>{value}</strong>
    <div className={`metric-detail ${tone}`}>
      {tone !== 'neutral' && <span className="change-pill"><Icon name={tone === 'up' ? 'arrowUp' : 'arrowDown'} size={13} />{change}</span>}
      <span>{detail}</span>
    </div>
  </article>
}

function TrendChart({ data, comparison }: { data: number[]; comparison: number[] }) {
  const width = 720
  const height = 230
  const max = Math.max(...data, ...comparison) * 1.14
  const points = (values: number[]) => values.map((v, index) => `${40 + index * ((width - 65) / (values.length - 1))},${height - 25 - (v / max) * (height - 50)}`).join(' ')
  const area = `40,${height - 25} ${points(data)} ${width - 25},${height - 25}`
  const labels = data.length === 12 ? data.map((_, index) => `${index + 1}월`) : data.map((_, index) => `${index + 1}일`)
  return <div className="chart-shell">
    <svg className="trend-svg" viewBox={`0 0 ${width} ${height}`} role="img" aria-label="매출 추이 비교 그래프">
      <defs><linearGradient id="chartFill" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stopColor="#2563eb" stopOpacity=".18" /><stop offset="1" stopColor="#2563eb" stopOpacity="0" /></linearGradient></defs>
      {[0, 1, 2, 3].map((line) => <line key={line} x1="40" x2={width - 25} y1={30 + line * 52} y2={30 + line * 52} className="grid-line" />)}
      <polygon points={area} fill="url(#chartFill)" />
      <polyline points={points(comparison)} className="line previous" />
      <polyline points={points(data)} className="line current" />
      {data.map((value, index) => <g key={index} className="point-group">
        <circle cx={40 + index * ((width - 65) / (data.length - 1))} cy={height - 25 - (value / max) * (height - 50)} r="9" className="hover-target" />
        <circle cx={40 + index * ((width - 65) / (data.length - 1))} cy={height - 25 - (value / max) * (height - 50)} r="3" className="point" />
        <g className="svg-tooltip" transform={`translate(${Math.min(width - 150, Math.max(10, 40 + index * ((width - 65) / (data.length - 1)) - 60))},${Math.max(5, height - 85 - (value / max) * (height - 50))})`}>
          <rect width="122" height="49" rx="7" />
          <text x="10" y="19">{labels[index]}</text><text x="10" y="38">{(value * 10000).toLocaleString()}원</text>
        </g>
      </g>)}
      {labels.map((label, index) => <text key={label} x={40 + index * ((width - 65) / (labels.length - 1))} y={height - 5} className="axis-label">{label}</text>)}
    </svg>
  </div>
}

function Dashboard() {
  const [year, setYear] = useState('2026')
  const [month, setMonth] = useState('전체')
  const [basis, setBasis] = useState<'결제일' | '이용기간'>('결제일')
  const [selectedProducts, setSelectedProducts] = useState<Product[]>(products)
  const [productOpen, setProductOpen] = useState(false)
  const [range, setRange] = useState<'월별' | '일별'>('월별')
  const [notice, setNotice] = useState('')

  const factor = basis === '이용기간' ? 0.94 : 1
  const totals = useMemo(() => products.reduce((acc, product) => {
    const indexes = month === '전체' ? Array.from({ length: 10 }, (_, i) => i) : [Number(month) - 1]
    acc[product] = selectedProducts.includes(product) ? indexes.reduce((sum, index) => sum + baseMonthly[product][index] * 1000000 * factor, 0) : 0
    return acc
  }, {} as Record<Product, number>), [month, basis, selectedProducts])
  const total = Object.values(totals).reduce((sum, value) => sum + value, 0)
  const payments = Math.round(total / 73500)
  const trend = useMemo(() => {
    if (range === '월별') return Array.from({ length: 12 }, (_, monthIndex) => selectedProducts.reduce((sum, product) => sum + baseMonthly[product][monthIndex] * factor, 0))
    const monthIndex = month === '전체' ? 9 : Number(month) - 1
    const monthTotal = selectedProducts.reduce((sum, product) => sum + baseMonthly[product][monthIndex] * factor, 0)
    return Array.from({ length: 30 }, (_, day) => monthTotal / 30 * (0.72 + ((day * 7) % 11) / 20))
  }, [range, selectedProducts, month, factor])
  const comparison = trend.map((value, index) => value * (0.82 + (index % 4) * 0.018))
  const sortedProducts = [...products].sort((a, b) => totals[b] - totals[a])
  const gradient = sortedProducts.reduce((result, product) => {
    const prior = result.end
    const part = total ? (totals[product] / total) * 100 : 0
    return { css: `${result.css}${result.css ? ',' : ''}${productColors[product]} ${prior}% ${prior + part}%`, end: prior + part }
  }, { css: '', end: 0 })

  const toggleProduct = (product: Product) => setSelectedProducts((current) => current.includes(product) ? (current.length === 1 ? current : current.filter((item) => item !== product)) : [...current, product])
  const reset = () => { setYear('2026'); setMonth('전체'); setBasis('결제일'); setSelectedProducts(products); setRange('월별') }

  return <main className="main-content">
    <section className="page-heading">
      <div><div className="eyebrow"><span className="status-dot" /> 샘플 데이터</div><h1>통합 대시보드</h1><p>매출 현황과 핵심 변동 요인을 한눈에 확인하세요.</p></div>
      <div className="heading-actions"><Button onClick={() => setNotice('새로고침되었습니다')}><Icon name="refresh" />새로고침</Button><Button variant="primary" onClick={() => setNotice('엑셀 업로드는 데이터 관리 화면에서 연결됩니다')}><Icon name="upload" />엑셀 업로드</Button></div>
    </section>

    {notice && <button className="notice" onClick={() => setNotice('')}><Icon name="check" />{notice}<Icon name="close" /></button>}

    <section className="filter-card">
      <div className="filter-group"><span className="filter-label">조회 기간</span><div className="filter-row"><Select value={year} onChange={setYear} label="조회 연도"><option>2026</option></Select><Select value={month} onChange={setMonth} label="조회 월"><option>전체</option>{Array.from({ length: 10 }, (_, i) => <option key={i + 1} value={i + 1}>{i + 1}월</option>)}</Select></div></div>
      <div className="filter-divider" />
      <div className="filter-group"><span className="filter-label">매출 집계 기준</span><div className="segmented"><button className={basis === '결제일' ? 'selected' : ''} onClick={() => setBasis('결제일')}>결제일 기준</button><button className={basis === '이용기간' ? 'selected' : ''} onClick={() => setBasis('이용기간')}>이용기간 기준</button></div></div>
      <div className="filter-divider" />
      <div className="filter-group product-filter"><span className="filter-label">매출 항목</span><button className="product-trigger" onClick={() => setProductOpen(!productOpen)}>전체 {selectedProducts.length}개 항목 <Icon name="chevron" size={15} /></button>
        {productOpen && <div className="product-popover">{products.map((product) => <label key={product}><input type="checkbox" checked={selectedProducts.includes(product)} onChange={() => toggleProduct(product)} /><span style={{ background: productColors[product] }} />{product}</label>)}</div>}
      </div>
      <Button variant="ghost" className="reset-button" onClick={reset}>초기화</Button>
    </section>

    {basis === '이용기간' && <div className="info-banner"><Icon name="info" /><span>이용기간 기준 매출은 내부 경영분석용 지표이며 법정 회계 매출과 다를 수 있습니다.</span></div>}

    <section className="metrics">
      <MetricCard label="선택 기간 총매출" value={`${Math.round(total / 10000).toLocaleString()}만원`} detail={`${basis} 기준 · 순매출`} />
      <MetricCard label="전년 동기 대비" value={`+${Math.round(total * .164 / 10000).toLocaleString()}만원`} change="16.4%" detail="비교 기간 대비" tone="up" />
      <MetricCard label="직전 기간 대비" value="+8.7%" change="8.7%" detail="동일 길이 기간" tone="up" />
      <MetricCard label="일평균 매출" value={`${Math.round(total / (month === '전체' ? 304 : 30) / 10000).toLocaleString()}만원`} detail={`${month === '전체' ? '304' : '30'}일 기준`} />
      <MetricCard label="총 결제 건수" value={`${payments.toLocaleString()}건`} change="11.2%" detail="전년 동기 대비" tone="up" />
      <MetricCard label="평균 결제금액" value={`${Math.round(payments ? total / payments : 0).toLocaleString()}원`} change="2.1%" detail="전년 동기 대비" tone="down" />
    </section>

    <section className="dashboard-grid">
      <article className="panel trend-panel">
        <div className="panel-head"><div><h2>매출 추이 비교</h2><p>2026년과 전년도 동일 기간 매출을 비교합니다.</p></div><div className="panel-tools"><div className="legend"><span><i className="legend-current" />2026년</span><span><i className="legend-previous" />2025년</span></div><div className="segmented small"><button className={range === '일별' ? 'selected' : ''} onClick={() => setRange('일별')}>일별</button><button className={range === '월별' ? 'selected' : ''} onClick={() => setRange('월별')}>월별</button></div></div></div>
        <TrendChart data={trend} comparison={comparison} />
      </article>
      <article className="panel share-panel">
        <div className="panel-head"><div><h2>상품별 매출 비중</h2><p>상품을 선택해 필터링할 수 있습니다.</p></div><Button variant="ghost">상세보기</Button></div>
        <div className="donut-wrap"><div className="donut" style={{ background: `conic-gradient(${gradient.css || '#e2e8f0 0 100%'})` }}><div className="donut-center"><span>총매출</span><strong>{Math.round(total / 100000000)}억원</strong></div></div></div>
        <div className="share-list">{sortedProducts.map((product) => <button key={product} onClick={() => setSelectedProducts([product])}><span className="color-dot" style={{ background: productColors[product] }} /><strong>{product}</strong><span>{Math.round(totals[product] / 10000).toLocaleString()}만원</span><b>{total ? ((totals[product] / total) * 100).toFixed(1) : '0.0'}%</b></button>)}</div>
      </article>
    </section>

    <section className="panel change-panel">
      <div className="panel-head"><div><h2>매출 증감 분석</h2><p>직전 기간 대비 상품별 매출 기여도를 분석합니다.</p></div><Button variant="ghost">분석 리포트 보기 <span>→</span></Button></div>
      <div className="change-table">
        <div className="table-head"><span>상품</span><span>이번 기간 매출</span><span>직전 기간 매출</span><span>증감액</span><span>증감률</span><span>기여도</span></div>
        {sortedProducts.slice(0, 4).map((product, index) => {
          const rate = [18.4, 12.7, -6.2, 5.8][index]
          const previous = totals[product] / (1 + rate / 100)
          return <div className="table-row" key={product}><span><i style={{ background: productColors[product] }} />{product}</span><strong>{Math.round(totals[product] / 10000).toLocaleString()}만원</strong><span>{Math.round(previous / 10000).toLocaleString()}만원</span><span className={rate >= 0 ? 'positive' : 'negative'}>{rate >= 0 ? '+' : ''}{Math.round((totals[product] - previous) / 10000).toLocaleString()}만원</span><span><b className={`rate-badge ${rate >= 0 ? 'up' : 'down'}`}><Icon name={rate >= 0 ? 'arrowUp' : 'arrowDown'} size={12} />{Math.abs(rate)}%</b></span><span><i className="progress"><i style={{ width: `${Math.abs(rate) * 3}%`, background: rate >= 0 ? '#2563eb' : '#dc2626' }} /></i></span></div>
        })}
      </div>
    </section>
  </main>
}

const sampleTransactions = Array.from({ length: 12 }, (_, index) => ({
  id: `ERD-2609-${String(index + 1).padStart(4, '0')}`,
  date: `2026-09-${String(index + 2).padStart(2, '0')}`,
  product: products[index % products.length],
  name: ['100시간권', '4주 기간권', '고정석 4주', '1일권', '사물함 4주', '스터디룸 2시간'][index % 6],
  amount: [120000, 160000, 210000, 15000, 30000, 40000][index % 6],
  status: index === 7 ? '부분환불' : '결제완료',
}))

function AnalysisScreen({ type }: { type: string }) {
  const [selected, setSelected] = useState<string | null>(null)
  const [sortDesc, setSortDesc] = useState(true)
  const months = Array.from({ length: 10 }, (_, i) => {
    const total = products.reduce((sum, product) => sum + baseMonthly[product][i] * 1000000, 0)
    const previous = total * (.81 + (i % 3) * .025)
    return { label: `${i + 1}월`, total, previous, change: (total / previous - 1) * 100 }
  })
  const days = Array.from({ length: 14 }, (_, i) => {
    const total = 5200000 + ((i * 1300000) % 4100000)
    return { label: `9월 ${i + 1}일`, total, previous: total * (.84 + (i % 4) * .03), change: 6 + (i % 5) * 2.7 }
  })
  const isDaily = type === '일별 매출 분석'
  const isYearly = type === '연도별 매출 분석'
  const rows = isDaily ? days : months
  const title = type
  const subtitle = isDaily ? '일자별 매출과 전년 동일 날짜·요일 실적을 비교합니다.' : isYearly ? '2026년 기준 연도의 연중 누적 실적을 관리합니다.' : '기간별 매출 변화와 비교 지표를 분석합니다.'
  const ordered = [...rows].sort((a, b) => sortDesc ? b.total - a.total : a.total - b.total)
  return <main className="main-content">
    <section className="page-heading"><div><div className="eyebrow"><span className="status-dot" /> 샘플 데이터</div><h1>{title}</h1><p>{subtitle}</p></div><div className="heading-actions"><Button onClick={() => exportRows(rows, title)}><Icon name="upload" />분석 결과 다운로드</Button></div></section>
    <section className="filter-card compact-filter"><div className="filter-group"><span className="filter-label">조회 연도</span><Select value="2026" onChange={() => {}} label="연도"><option>2026</option></Select></div><div className="filter-group"><span className="filter-label">조회 기간</span><Select value={isDaily ? '9월' : '전체'} onChange={() => {}} label="기간"><option>{isDaily ? '9월' : '전체'}</option></Select></div><div className="filter-group"><span className="filter-label">매출 기준</span><div className="segmented"><button className="selected">결제일 기준</button><button>이용기간 기준</button></div></div></section>
    <section className="metrics analysis-metrics">
      <MetricCard label={isYearly ? '연중 누적 매출' : '기간 총매출'} value="12억 4,680만원" detail="2026년 실적" />
      <MetricCard label="비교 기간 대비" value="+14.8%" change="14.8%" detail={isYearly ? '기준 연도 · 비교 데이터 없음' : '동일 기간 대비'} tone="up" />
      <MetricCard label="평균 매출" value={isDaily ? '684만원' : '1억 2,468만원'} detail={isDaily ? '일평균' : '월평균'} />
      <MetricCard label="평균 결제금액" value="73,520원" change="2.4%" detail="이전 기간 대비" tone="up" />
    </section>
    <section className="panel analysis-chart-panel">
      <div className="panel-head"><div><h2>{isDaily ? '일별 매출 현황' : isYearly ? '연간 누적 매출' : '월별 매출 비교'}</h2><p>막대에 마우스를 올리면 상세 금액을 확인할 수 있습니다.</p></div><div className="legend"><span><i className="legend-current" />2026년</span><span><i className="legend-previous" />비교 기간</span></div></div>
      <div className="bar-chart">{rows.map((row) => <button key={row.label} className="bar-column" onClick={() => setSelected(row.label)} title={`${row.label} · ${row.total.toLocaleString()}원`}><span className="bar-value">{Math.round(row.total / 10000).toLocaleString()}만</span><span className="bar-pair"><i style={{ height: `${Math.max(18, row.previous / 1800000)}%` }} /><b style={{ height: `${Math.max(22, row.total / 1800000)}%` }} /></span><small>{row.label.replace('9월 ', '')}</small></button>)}</div>
    </section>
    <section className="panel analysis-table-panel">
      <div className="panel-head"><div><h2>{isDaily ? '일별 상세 내역' : '기간별 매출 내역'}</h2><p>행을 선택하면 관련 거래를 확인합니다.</p></div><Button variant="ghost" onClick={() => setSortDesc(!sortDesc)}>매출액 {sortDesc ? '내림차순' : '오름차순'}</Button></div>
      <div className="simple-table"><div className="simple-head"><span>기간</span><span>당해 매출</span><span>비교 매출</span><span>증감액</span><span>증감률</span><span>상태</span></div>{ordered.map((row) => <button className="simple-row" key={row.label} onClick={() => setSelected(row.label)}><strong>{row.label}</strong><span>{row.total.toLocaleString()}원</span><span>{Math.round(row.previous).toLocaleString()}원</span><span className="positive">+{Math.round(row.total - row.previous).toLocaleString()}원</span><span className="positive">+{row.change.toFixed(1)}%</span><span><b className="status-chip">정상</b></span></button>)}</div>
    </section>
    {selected && <TransactionModal date={selected} onClose={() => setSelected(null)} />}
  </main>
}

function ProductScreen() {
  const [activeProduct, setActiveProduct] = useState<Product>('기간권')
  return <main className="main-content">
    <section className="page-heading"><div><div className="eyebrow"><span className="status-dot" /> 6개 상품 분류</div><h1>상품별 매출 분석</h1><p>상품별 매출액, 판매 건수, 할인과 환불을 비교합니다.</p></div><Button onClick={() => exportRows(sampleTransactions, '상품별 매출')}><Icon name="upload" />다운로드</Button></section>
    <section className="product-cards">{products.map((product) => { const total = baseMonthly[product].reduce((a, b) => a + b, 0) * 1000000; return <button key={product} className={`product-card ${activeProduct === product ? 'selected' : ''}`} onClick={() => setActiveProduct(product)}><i style={{ background: productColors[product] }} /><span>{product}</span><strong>{Math.round(total / 10000).toLocaleString()}만원</strong><small>전월 대비 <b>+{(6.2 + products.indexOf(product) * 1.3).toFixed(1)}%</b></small></button> })}</section>
    <section className="dashboard-grid product-detail-grid"><article className="panel analysis-chart-panel"><div className="panel-head"><div><h2>{activeProduct} 월별 매출 추이</h2><p>2026년 1월–10월</p></div></div><div className="bar-chart product-bars">{baseMonthly[activeProduct].slice(0, 10).map((value, i) => <div className="bar-column" key={i}><span className="bar-value">{value * 100}만</span><span className="bar-pair"><b style={{ height: `${value * 1.7}%`, background: productColors[activeProduct] }} /></span><small>{i + 1}월</small></div>)}</div></article><article className="panel product-summary"><h2>상품 성과 요약</h2>{[['판매 건수', '1,284건'], ['평균 결제금액', '128,420원'], ['전체 매출 비중', '27.8%'], ['할인금액', '1,860만원'], ['환불금액', '420만원']].map(([label, value]) => <div key={label}><span>{label}</span><strong>{value}</strong></div>)}</article></section>
  </main>
}

function TransactionModal({ date, onClose }: { date: string; onClose: () => void }) {
  return <div className="modal-backdrop" onMouseDown={onClose}><section className="modal-card" onMouseDown={(e) => e.stopPropagation()}><div className="modal-head"><div><h2>{date} 거래 상세</h2><p>결제 완료 및 환불 거래를 포함합니다.</p></div><Button variant="ghost" onClick={onClose}><Icon name="close" /></Button></div><div className="transaction-list">{sampleTransactions.slice(0, 6).map((row) => <div key={row.id}><span><strong>{row.name}</strong><small>{row.id} · 카드결제</small></span><b>{row.amount.toLocaleString()}원</b><i className={row.status === '결제완료' ? 'ok' : 'warn'}>{row.status}</i></div>)}</div><div className="modal-foot"><span>총 6건</span><strong>575,000원</strong></div></section></div>
}

function exportRows(rows: object[], name: string) {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), '분석결과')
  XLSX.writeFile(workbook, `${name}_2026.xlsx`)
}

function DataManagement() {
  const [rows, setRows] = useState<Record<string, unknown>[]>([])
  const [fileName, setFileName] = useState('')
  const readFile = async (file?: File) => {
    if (!file) return
    const workbook = XLSX.read(await file.arrayBuffer())
    const parsed = XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[workbook.SheetNames[0]], { defval: '' })
    setRows(parsed)
    setFileName(file.name)
    localStorage.setItem('iruda-upload', JSON.stringify(parsed.slice(0, 500)))
  }
  return <main className="main-content">
    <section className="page-heading"><div><div className="eyebrow">브라우저 내 처리</div><h1>매출 데이터 관리</h1><p>원본 파일은 서버로 전송되지 않으며 현재 브라우저에서만 처리됩니다.</p></div>{rows.length > 0 && <Button variant="primary" onClick={() => exportRows(rows, '전체_매출내역')}><Icon name="upload" />엑셀 다운로드</Button>}</section>
    <label className="upload-zone" onDragOver={(e) => e.preventDefault()} onDrop={(e) => { e.preventDefault(); readFile(e.dataTransfer.files[0]) }}><input type="file" accept=".xlsx,.xls,.csv" onChange={(e) => readFile(e.target.files?.[0])} /><span className="upload-icon"><Icon name="upload" size={24} /></span><strong>XLSX, XLS 또는 CSV 파일을 선택하세요</strong><p>파일을 이 영역에 끌어다 놓을 수도 있습니다.</p><b>파일 선택</b></label>
    {rows.length > 0 ? <section className="panel upload-result"><div className="panel-head"><div><h2>데이터 미리보기</h2><p>{fileName} · 총 {rows.length.toLocaleString()}행</p></div><span className="success-label"><Icon name="check" />형식 확인 완료</span></div><div className="preview-scroll"><table><thead><tr>{Object.keys(rows[0]).slice(0, 8).map((key) => <th key={key}>{key}</th>)}</tr></thead><tbody>{rows.slice(0, 8).map((row, index) => <tr key={index}>{Object.keys(rows[0]).slice(0, 8).map((key) => <td key={key}>{String(row[key] ?? '')}</td>)}</tr>)}</tbody></table></div><div className="mapping-note"><Icon name="info" /><span>헤더명이 표준 필드와 다르면 다음 단계에서 열 매핑이 필요합니다. 원본에 없는 값은 생성하지 않습니다.</span></div></section> : <section className="panel format-guide"><h2>표준 데이터 필드</h2><div>{['거래번호', '결제일시', '상품명', '매출분류', '정상가격', '할인금액', '실제 결제금액', '결제수단', '이용시작일', '이용종료일', '거래상태', '원거래번호', '환불금액'].map((field) => <span key={field}>{field}</span>)}</div></section>}
  </main>
}

function ValidationScreen() {
  const [checked, setChecked] = useState(false)
  const issues = [
    ['오류', '원거래번호 없는 환불', 'ERD-2609-0081', '원거래 연결이 필요합니다.'],
    ['오류', '이용 종료일 오류', 'ERD-2609-0124', '종료일이 시작일보다 빠릅니다.'],
    ['주의', '상품 미분류', 'ERD-2609-0192', '상품 분류를 직접 지정하세요.'],
  ]
  return <main className="main-content"><section className="page-heading"><div><div className="eyebrow">마지막 검사 · 방금 전</div><h1>데이터 검증</h1><p>중복, 누락, 환불 연결과 이용기간 오류를 자동 점검합니다.</p></div><Button variant="primary" onClick={() => setChecked(true)}><Icon name="refresh" />오류 재검증</Button></section>{checked && <div className="info-banner success"><Icon name="check" />재검증을 완료했습니다. 수정되지 않은 3개 항목이 남아 있습니다.</div>}<section className="validation-summary"><article><span className="validation-icon ok"><Icon name="check" /></span><div><strong>1,842건</strong><small>정상</small></div></article><article><span className="validation-icon warn"><Icon name="info" /></span><div><strong>1건</strong><small>주의</small></div></article><article><span className="validation-icon error"><Icon name="close" /></span><div><strong>2건</strong><small>오류</small></div></article></section><section className="panel issue-panel"><div className="panel-head"><div><h2>검증 필요 항목</h2><p>오류를 수정한 뒤 재검증하세요.</p></div></div>{issues.map(([level, title, id, detail]) => <button className="issue-row" key={id}><b className={level === '오류' ? 'error' : 'warn'}>{level}</b><span><strong>{title}</strong><small>{detail}</small></span><code>{id}</code><span>수정하기 →</span></button>)}</section></main>
}

export default function App() {
  const [active, setActive] = useState('통합 대시보드')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  return <div className="app-shell">
    <Sidebar active={active} onChange={setActive} open={sidebarOpen} onClose={() => setSidebarOpen(false)} />
    <div className="content-shell">
      <header className="mobile-header"><Button variant="ghost" onClick={() => setSidebarOpen(true)}><Icon name="menu" /></Button><strong>이루다</strong><span>{active}</span></header>
      {active === '통합 대시보드' ? <Dashboard />
        : ['일별 매출 분석', '월별 매출 분석', '연도별 매출 분석', '이용기간 매출 분석'].includes(active) ? <AnalysisScreen type={active} />
        : active === '상품별 매출 분석' ? <ProductScreen />
        : active === '매출 데이터 관리' ? <DataManagement />
        : active === '데이터 검증' ? <ValidationScreen />
        : <main className="main-content"><section className="page-heading"><div><h1>설정</h1><p>상품 분류 및 분석 기준 설정 화면입니다.</p></div></section><section className="placeholder panel"><Icon name="settings" size={28} /><h2>상품 설정</h2><p>세부 상품 추가·수정 기능은 다음 단계에서 연결됩니다.</p></section></main>}
    </div>
  </div>
}
