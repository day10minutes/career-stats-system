import React, { useState, useEffect, useMemo } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, AreaChart, Area
} from 'recharts';
import { 
  Upload, FileSpreadsheet, Download, FileText, Settings, 
  Users, BarChart3, PieChart as PieIcon, LayoutDashboard, PlusCircle, AlertCircle
} from 'lucide-react';
import { ExcelService } from './utils/ExcelService';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';

const COLORS = ['#2563eb', '#7c3aed', '#db2777', '#ea580c', '#16a34a', '#0891b2', '#4f46e5'];

export default function App() {
  const [data, setData] = useState([]);
  const [stats, setStats] = useState(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [mapping, setMapping] = useState({
    college: '단과대',
    grade: '학년',
    status: '참석여부',
    date: '참여일시',
    major: '전공(학과)',
    unit: '단위',
    type: '구분',
    program: '프로그램명',
    education: '교육명'
  });
  
  // 마스터 데이터 (관리자가 수정 가능)
  const [categories, setCategories] = useState({
    units: ['개발 EP. 2-3-2', '개발 EP. 2-4-1', '현장 유입'],
    types: ['진로', '취업', '창업', '기타'],
    programs: ['진로취업컨설팅', '면접코칭', '자기소개서 첨삭'],
    educations: ['진로개발', '실전면접', '직무분석']
  });

  const [isConfigOpen, setIsConfigOpen] = useState(false);
  const [insight, setInsight] = useState("");

  // 파일 업로드 핸들러
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const json = await ExcelService.parseExcel(file);
    setData(json);
    
    // 만약 데이터에 매핑된 헤더가 하나도 없으면 매핑 요청
    const headers = Object.keys(json[0] || {});
    const hasHeader = headers.some(h => Object.values(mapping).includes(h));
    
    if (!hasHeader) {
      setIsMappingMode(true);
    }
  };

  // 통계 계산 감시
  useEffect(() => {
    if (data.length > 0) {
      const calculated = ExcelService.calculateStats(data, mapping);
      setStats(calculated);
      generateInsight(calculated);
    }
  }, [data, mapping]);

  // 인사이트 텍스트 자동 생성
  const generateInsight = (s) => {
    if (!s) return;
    const topCollege = Object.entries(s.byCollege).sort((a,b) => b[1] - a[1])[0];
    const topGrade = Object.entries(s.byGrade).sort((a,b) => b[1] - a[1])[0];
    
    const text = `신청자 수(${s.applicants}명) 대비 실제 참여자 수(${s.participants}명)로 대조되었습니다. 
    가장 참여가 활발한 곳은 ${topCollege?.[0]}이며, ${topGrade?.[0]} 학생들의 유입이 두드러집니다.`;
    setInsight(text);
  };

  // PDF 내보내기
  const exportPDF = () => {
    const input = document.getElementById('dashboard-view');
    html2canvas(input).then((canvas) => {
      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF('p', 'mm', 'a4');
      const imgProps= pdf.getImageProperties(imgData);
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
      pdf.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
      pdf.save("진로취업_통계_리포트.pdf");
    });
  };

  // 차트용 데이터 변환
  const chartDataCollege = useMemo(() => 
    Object.entries(stats?.byCollege || {}).map(([name, value]) => ({ name, value })),
    [stats]
  );
  
  const chartDataGrade = useMemo(() => 
    Object.entries(stats?.byGrade || {}).map(([name, value]) => ({ name, value })),
    [stats]
  );

  return (
    <div className="dashboard-container">
      {/* Sidebar */}
      <aside>
        <div style={{display:'flex', alignItems:'center', gap:'10px', color:'var(--primary)', marginBottom:'1rem'}}>
          <LayoutDashboard size={28} />
          <h2 style={{fontSize:'1.2rem'}}>StatSystem</h2>
        </div>
        
        <nav style={{display:'flex', flexDirection:'column', gap:'10px'}}>
          <button className={`btn ${activeTab === 'dashboard' ? 'btn-primary' : 'btn-outline'}`} onClick={() => setActiveTab('dashboard')}>
            <BarChart3 size={18} /> 대시보드
          </button>
          <button className={`btn ${activeTab === 'data' ? 'btn-outline' : 'btn-outline'}`} onClick={() => setActiveTab('data')}>
            <FileSpreadsheet size={18} /> 참여자 명단
          </button>
          <button className={`btn btn-outline`} onClick={() => setIsMappingMode(!isMappingMode)}>
            <Settings size={18} /> 항목(컬럼) 관리
          </button>
        </nav>

        <div style={{marginTop:'auto'}}>
           <label className="btn btn-primary" style={{width:'100%', justifyContent:'center'}}>
             <Upload size={18} /> 데이터 업로드
             <input type="file" hidden accept=".xlsx, .xls" onChange={handleFileUpload} />
           </label>
        </div>
      </aside>

      {/* Main Content */}
      <main id="dashboard-view">
        <header style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'2rem'}}>
          <div>
            <h1>취업 프로그램 참여 통계</h1>
            <p style={{color:'var(--text-muted)'}}>참여자 데이터를 기반으로 실시간 인사이트를 제공합니다.</p>
          </div>
          <div style={{display:'flex', gap:'10px'}}>
            <button className="btn btn-outline" onClick={() => ExcelService.exportToExcel(data)}>
              <Download size={18} /> 엑셀 다운로드
            </button>
            <button className="btn btn-primary" onClick={exportPDF}>
              <FileText size={18} /> PDF 리포트 추출
            </button>
          </div>
        </header>

        {data.length === 0 ? (
          <div className="card" style={{height:'400px', display:'flex', flexDirection:'column', alignItems:'center', justifyContent:'center', borderStyle:'dashed'}}>
            <PlusCircle size={48} color="#94a3b8" style={{marginBottom:'1rem'}} />
            <h3>분석할 엑셀 파일을 업로드해 주세요</h3>
            <p style={{color:'var(--text-muted)'}}>엑셀을 드래그하거나 업로드 버튼을 클릭하세요.</p>
          </div>
        ) : (
          <>
            {/* Stat Cards */}
            <div className="stat-grid">
              <div className="card">
                <span style={{color:'var(--text-muted)', fontSize:'0.9rem'}}>총 신청자</span>
                <div style={{fontSize:'1.8rem', fontWeight:700}}>{stats?.total}명</div>
              </div>
              <div className="card">
                <span style={{color:'var(--text-muted)', fontSize:'0.9rem', color:'var(--primary)'}}>실제 참여자</span>
                <div style={{fontSize:'1.8rem', fontWeight:700, color:'var(--primary)'}}>{stats?.participants}명</div>
              </div>
              <div className="card">
                <span style={{color:'var(--text-muted)', fontSize:'0.9rem'}}>참여율</span>
                <div style={{fontSize:'1.8rem', fontWeight:700}}>
                  {((stats?.participants / stats?.total) * 100).toFixed(1)}%
                </div>
              </div>
            </div>

            {/* Analysis Row */}
            <div className="card" style={{backgroundColor:'#f0f7ff', border:'1px solid #bae6fd'}}>
              <div style={{display:'flex', gap:'10px', alignItems:'center', marginBottom:'0.5rem'}}>
                <AlertCircle size={20} color="var(--primary)" />
                <h3 style={{fontSize:'1rem'}}>자동 데이터 인사이트</h3>
              </div>
              <p style={{margin:0, lineHeight:'1.6'}}>{insight}</p>
            </div>

            <div style={{display:'grid', gridTemplateColumns:'1fr 1fr', gap:'1.5rem'}}>
              <div className="card">
                <h3>단과대별 분포</h3>
                <div style={{height:'300px', marginTop:'1rem'}}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie data={chartDataCollege} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} label>
                        {chartDataCollege.map((entry, index) => <Cell key={index} fill={COLORS[index % COLORS.length]} />)}
                      </Pie>
                      <Tooltip />
                    </PieChart>
                  </ResponsiveContainer>
                </div>
              </div>
              <div className="card">
                <h3>학년별 참여 현황</h3>
                <div style={{height:'300px', marginTop:'1rem'}}>
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={chartDataGrade}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="name" />
                      <YAxis />
                      <Tooltip />
                      <Bar dataKey="value" fill="var(--primary)" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>
          </>
        )}

        {/* Dynamic Mapping Sidebar (Overlay) */}
        {isMappingMode && (
          <div className="card" style={{position:'fixed', top:'20px', right:'20px', width:'300px', zIndex:100, border:'2px solid var(--primary)'}}>
            <h3>항목(컬럼) 매핑 설정</h3>
            <p style={{fontSize:'0.8rem', color:'var(--text-muted)'}}>엑셀의 어느 컬럼이 다음 항목인지 지정하세요.</p>
            <div style={{display:'flex', flexDirection:'column', gap:'10px', marginTop:'1rem'}}>
              {Object.keys(mapping).map(key => (
                <div key={key}>
                  <label style={{fontSize:'0.75rem', fontWeight:600}}>{key === 'college' ? '단과대' : key === 'grade' ? '학년' : key}</label>
                  <select 
                    value={mapping[key]} 
                    onChange={(e) => setMapping({...mapping, [key]: e.target.value})}
                    style={{width:'100%', padding:'5px', borderRadius:'4px', border:'1px solid #ccc'}}
                  >
                    {data.length > 0 && Object.keys(data[0]).map(h => <option key={h} value={h}>{h}</option>)}
                  </select>
                </div>
              ))}
            </div>
            <button className="btn btn-primary" style={{width:'100%', marginTop:'1rem'}} onClick={() => setIsMappingMode(false)}>적용 및 닫기</button>
          </div>
        )}
      </main>
    </div>
  );
}
