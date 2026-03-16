import React, { useState, useEffect, useMemo } from 'react';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer,
  PieChart, Pie, Cell, LineChart, Line, AreaChart, Area
} from 'recharts';
import { 
  Upload, FileSpreadsheet, Download, FileText, Settings, 
  Users, BarChart3, PieChart as PieIcon, LayoutDashboard, PlusCircle, AlertCircle, Edit, Save, CheckCircle
} from 'lucide-react';
import { ExcelService } from './utils/ExcelService';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';

const COLORS = [
  '#2563eb', '#7c3aed', '#db2777', '#ea580c', '#16a34a', '#0891b2', '#4f46e5',
  '#ef4444', '#f59e0b', '#10b981', '#3b82f6', '#6366f1', '#a855f7', '#ec4899',
  '#f43f5e', '#f97316', '#84cc16', '#06b6d4', '#8b5cf6', '#d946ef'
];

export default function App() {
  const [data, setData] = useState([]);
  const [stats, setStats] = useState(null);
  const [activeTab, setActiveTab] = useState('dashboard');
  const [workbookData, setWorkbookData] = useState({});
  const [sheetNames, setSheetNames] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('ALL');
  const [isMappingMode, setIsMappingMode] = useState(false);
  const [mapping, setMapping] = useState({
    college: '단과대',
    grade: '학년',
    status: '참석여부',
    date: '일정/일자',
    major: '전공(학과)',
    unit: '단위',
    studentId: '학번',
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
  const [managerComment, setManagerComment] = useState("");
  const [isEditing, setIsEditing] = useState(false);

  // 파일 업로드 핸들러
  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const result = await ExcelService.parseExcel(file);
    setWorkbookData(result.workbook);
    setSheetNames(result.sheetNames);
    
    // 기본적으로 'ALL' 또는 첫 번째 시트 선택
    const initialSheet = result.sheetNames.includes('ALL') ? 'ALL' : result.sheetNames[0];
    setSelectedSheet(initialSheet);
    setData(result.workbook[initialSheet]);
    
    // 만약 데이터에 매핑된 헤더가 하나도 없으면 매핑 요청
    const firstData = result.workbook[initialSheet];
    const headers = Object.keys(firstData[0] || {});
    const hasHeader = headers.some(h => Object.values(mapping).includes(h));
    
    if (!hasHeader) {
      setIsMappingMode(true);
    }
  };

  // 시트 변경 핸들러
  const handleSheetChange = (sheetName) => {
    setSelectedSheet(sheetName);
    setData(workbookData[sheetName] || []);
  };

  // 통계 계산 감시
  useEffect(() => {
    if (data.length > 0) {
      const calculated = ExcelService.calculateStats(data, mapping);
      setStats(calculated);
      generateInsight(calculated);
      
      // 저장된 코멘트 불러오기
      const saved = localStorage.getItem(`mgr_comment_${selectedSheet}`);
      setManagerComment(saved || "");
      setIsEditing(false);
    }
  }, [data, mapping, selectedSheet]);

  // 인사이트 텍스트 자동 생성
  const generateInsight = (s) => {
    if (!s) return;
    const isAll = selectedSheet.toUpperCase().includes('ALL');
    
    let text = `데이터 분석 결과, 전반적으로 <strong>${s.topCollege}</strong>의 참여가 가장 비중이 높으며, <strong>${s.topGrade}</strong>의 참여 비율이 가장 높게 나타납니다. `;
    
    if (isAll) {
      text += `중복 참여가 가장 많은 프로그램은 <strong>${s.mostDuplicatedProgram}</strong>이며, <strong>${s.topMonth}</strong>에 가장 높은 참여율을 보이고 있습니다.`;
    } else {
      text += `개설했던 교육 중 <strong>${s.topProgram}</strong> 교육이 <strong>${s.participants}명</strong>으로 가장 참여자가 많았습니다.`;
    }
    if (s.needsCheckCount > 0) {
      text += ` <br><span style="color:#e11d48; font-size:0.9rem;">⚠️ 참고: 참석여부 빈 데이터 ${s.needsCheckCount}건이 확인되어 참여자로 가집계되었습니다.</span>`;
    }
    setInsight(text);
  };

  // 코멘트 저장
  const handleSaveComment = () => {
    localStorage.setItem(`mgr_comment_${selectedSheet}`, managerComment);
    setIsEditing(false);
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
    Object.entries(stats?.byCollege || {})
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value),
    [stats]
  );
  
  const chartDataGrade = useMemo(() => 
    Object.entries(stats?.byGrade || {}).map(([name, value]) => ({ name, value })),
    [stats]
  );

  return (
    <div className="dashboard-container">
      {/* Sidebar */}
      <aside style={{backgroundColor: '#fff', padding: '1.5rem'}}>
        <div style={{display:'flex', flexDirection:'column', alignItems:'center', marginBottom: '2rem', gap: '15px'}}>
          <img src="/assets/logo-cdc.png" alt="이화여자대학교 인재개발원" style={{width: '100%', maxWidth: '200px', borderRadius: '4px'}} />
          <div style={{textAlign:'center'}}>
            <div style={{fontSize: '1.1rem', fontWeight: 800, color: '#00462a'}}>진로취업 프로그램 통계</div>
          </div>
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
          <div style={{display:'flex', gap:'10px', alignItems:'center'}}>
            {sheetNames.length > 0 && (
              <div style={{marginRight: '1rem'}}>
                <label style={{fontSize: '0.8rem', color: 'var(--text-muted)', display: 'block', marginBottom: '4px'}}>분석 시트 선택</label>
                <select 
                  value={selectedSheet} 
                  onChange={(e) => handleSheetChange(e.target.value)}
                  className="btn btn-outline"
                  style={{padding: '0.4rem 0.8rem', minWidth: '120px'}}
                >
                  {sheetNames.map(name => (
                    <option key={name} value={name}>{name}</option>
                  ))}
                </select>
              </div>
            )}
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
                {stats?.needsCheckCount > 0 && (
                  <div style={{color:'#e11d48', fontSize:'0.75rem', marginTop:'4px'}}>
                    ⚠️ 확인요망(빈칸): {stats.needsCheckCount}건
                  </div>
                )}
              </div>
            </div>

            {/* Analysis Row */}
            <div className="card" style={{borderLeft: '6px solid #00462a', backgroundColor: '#fff', padding: '2rem'}}>
              <div style={{marginBottom: '2rem'}}>
                <div style={{display:'flex', gap:'10px', alignItems:'center', marginBottom:'1.2rem'}}>
                  <BarChart3 size={24} color="#00462a" />
                  <h3 style={{fontSize:'1.2rem', margin: 0, border: 'none', padding: 0}}>📈 참여 정밀 분석 리포트</h3>
                </div>
                <p style={{margin:0, lineHeight:'1.8', fontSize: '1.05rem', color: '#334155'}} dangerouslySetInnerHTML={{__html: insight}} />
                <div style={{marginTop: '1.2rem', fontSize: '0.9rem', color: 'var(--text-muted)', borderTop: '1px solid var(--border)', paddingTop: '1.2rem'}}>
                  <div style={{display:'flex', gap:'30px'}}>
                    <span style={{display:'flex', alignItems:'center', gap: '5px'}}><CheckCircle size={14} /> 순 참여자: <strong>{stats?.participants}명</strong> 중 고유 인원 <strong>{stats?.netUniqueCount}명</strong></span>
                    <span style={{display:'flex', alignItems:'center', gap: '5px'}}><Users size={14} /> 중복 참여: <strong>{stats?.duplicateUserCount}명</strong>이 다회 이용 중</span>
                  </div>
                </div>
              </div>

              <div style={{backgroundColor: '#f8fafc', padding: '1.5rem', borderRadius: '12px', border: '1px solid var(--border)'}}>
                <div style={{display:'flex', justifyContent: 'space-between', alignItems:'center', marginBottom:'1rem'}}>
                  <div style={{display:'flex', gap:'8px', alignItems:'center'}}>
                      <FileText size={20} color="#00462a" />
                      <h4 style={{fontSize:'1rem', margin: 0, fontWeight: 700}}>📝 담당자 분석 코멘트</h4>
                  </div>
                  {!isEditing ? (
                    <button 
                      onClick={() => setIsEditing(true)}
                      className="btn btn-outline"
                      style={{display:'flex', alignItems:'center', gap:'5px', padding: '6px 16px', fontSize: '0.85rem'}}
                    >
                      <Edit size={16} /> 편집하기
                    </button>
                  ) : (
                    <button 
                      onClick={handleSaveComment}
                      className="btn btn-primary"
                      style={{display:'flex', alignItems:'center', gap:'5px', padding: '6px 16px', fontSize: '0.85rem', backgroundColor: '#00462a'}}
                    >
                      <Save size={16} /> 저장하기
                    </button>
                  )}
                </div>
                
                {isEditing ? (
                  <textarea 
                    value={managerComment}
                    onChange={(e) => setManagerComment(e.target.value)}
                    style={{
                      width: '100%', 
                      height: '110px', 
                      borderRadius: '8px', 
                      border: '2px solid #00462a', 
                      padding: '12px',
                      fontSize: '0.95rem',
                      fontFamily: 'inherit',
                      resize: 'none',
                      backgroundColor: '#fff',
                      lineHeight: '1.6'
                    }}
                    placeholder="이곳에 직접 분석 내용을 입력하세요..."
                  />
                ) : (
                  <div style={{
                    minHeight: '80px', 
                    padding: '12px', 
                    fontSize: '1rem', 
                    lineHeight: '1.7', 
                    color: '#1e293b',
                    whiteSpace: 'pre-wrap'
                  }}>
                    {managerComment || <span style={{color: '#94a3b8'}}>작성된 코멘트가 없습니다. 편집하기 버튼을 눌러 내용을 입력하세요.</span>}
                  </div>
                )}
              </div>
            </div>

            <div className="stat-grid" style={{marginBottom: '2rem'}}>
              <div className="card">
                <h3>단과대별 구성 (마우스 오버 시 상세 전공 확인)</h3>
                <div style={{height:'300px', marginTop:'1rem'}}>
                  <ResponsiveContainer width="100%" height="100%">
                    <PieChart>
                      <Pie 
                        data={chartDataCollege} 
                        dataKey="value" 
                        nameKey="name" 
                        cx="50%" cy="50%" 
                        outerRadius={80} 
                        label={({name, percent}) => `${name} ${(percent * 100).toFixed(0)}%`}
                      >
                        {chartDataCollege.map((entry, index) => <Cell key={index} fill={COLORS[index % COLORS.length]} />)}
                      </Pie>
                      <Tooltip content={({ active, payload }) => {
                        if (active && payload && payload.length) {
                          const name = payload[0].name;
                          const details = stats?.collegeDetails?.[name] || {};
                          const sortedDetails = Object.entries(details).sort((a,b) => b[1] - a[1]);
                          return (
                            <div className="card" style={{padding: '10px', fontSize: '0.8rem', border: '1px solid var(--primary)'}}>
                              <p style={{fontWeight: 700, marginBottom: '5px'}}>{name} ({payload[0].value}명)</p>
                              <hr />
                              {sortedDetails.slice(0, 5).map(([d, count]) => (
                                <p key={d}>{d}: {count}명</p>
                              ))}
                              {sortedDetails.length > 5 && <p>...외 {sortedDetails.length - 5}건</p>}
                            </div>
                          );
                        }
                        return null;
                      }} />
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
                      <YAxis allowDecimals={false} />
                      <Tooltip cursor={{fill: 'rgba(0,0,0,0.05)'}} />
                      <Bar dataKey="value" fill="var(--primary)" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* Detailed Report Table */}
            <div className="card">
              <h3>{selectedSheet.toUpperCase().includes('ALL') ? '📋 전체 운영 프로그램별 리포트' : '🗓️ 월별 세부 리포트'}</h3>
              <div style={{marginTop: '1.5rem'}}>
                {selectedSheet.toUpperCase().includes('ALL') ? (
                  /* ALL 시트용: 프로그램별 뷰 */
                  <div className="table-wrapper">
                    <table className="table-5col">
                      <thead>
                        <tr>
                          <th>프로그램명</th>
                          <th>인원</th>
                          <th>주요 단과대</th>
                          <th>주요 학년</th>
                          <th>주요 월</th>
                        </tr>
                      </thead>
                      <tbody>
                        {Object.entries(stats?.programMatrix || {})
                          .sort((a,b) => b[1].total - a[1].total)
                          .map(([name, s]) => {
                            const sortedCols = Object.entries(s.cols).sort((a,b)=>b[1]-a[1]);
                            const sortedGrads = Object.entries(s.grads).sort((a,b)=>b[1]-a[1]);
                            const topCol = sortedCols[0]?.[0] || '-';
                            const topGrad = sortedGrads[0]?.[0] || '-';
                             const sortedMArr = Object.entries(s.months)
                               .filter(([m]) => m && m.includes('월') && !m.includes('미확인'))
                               .sort((a,b)=>b[1]-a[1]);
                             const topMonth = sortedMArr[0]?.[0] || (s.months['일정 미확인'] ? '일정 미확인' : '-');
                            
                            const colTooltip = sortedCols.map(([k,v])=>`${k}: ${v}명`).join('\n');
                            const gradTooltip = sortedGrads.map(([k,v])=>`${k}: ${v}명`).join('\n');

                            return (
                              <tr key={name}>
                                <td style={{fontWeight: 700}}>{name}</td>
                                <td style={{color:'var(--primary)', fontWeight: 700}}>{s.total}명</td>
                                <td title={colTooltip} style={{cursor:'help', textDecoration:'underline dotted'}}>{topCol}</td>
                                <td title={gradTooltip} style={{cursor:'help', textDecoration:'underline dotted'}}>{topGrad}</td>
                                <td>{topMonth}</td>
                              </tr>
                            );
                          })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  /* 일반 시트용: 월별 뷰 */
                  Object.entries(stats?.matrix || {})
                    .sort((a, b) => {
                      const order = (m) => {
                        if (m === '일정 미확인') return 999;
                        const n = parseInt(m);
                        return isNaN(n) ? 998 : n;
                      };
                      return order(a[0]) - order(b[0]);
                    })
                    .map(([month, groups]) => {
                      const mTotal = Object.values(groups).reduce((acc, curr) => acc + curr.total, 0);
                      return (
                        <div key={month} style={{marginBottom: '2rem'}}>
                          <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:'1rem'}}>
                            <h4 style={{margin:0}}>📅 {month} 운영 현황</h4>
                            <span style={{backgroundColor:'#00462a', color:'#fff', padding:'4px 12px', borderRadius:'20px', fontSize:'0.8rem', fontWeight:700}}>
                              월 참여 합계: {mTotal}명
                            </span>
                          </div>
                          <div className="table-wrapper">
                            <table className="table-4col" style={{width:'100%', borderCollapse:'collapse'}}>
                              <thead>
                                <tr style={{borderBottom:'2px solid var(--border)', textAlign:'left'}}>
                                  <th style={{padding:'12px'}}>교육명(유형)</th>
                                  <th style={{padding:'12px'}}>인원</th>
                                  <th style={{padding:'12px'}}>주요 단과대</th>
                                  <th style={{padding:'12px'}}>주요 학년</th>
                                </tr>
                              </thead>
                              <tbody>
                                {Object.entries(groups).map(([name, s]) => {
                                  const sortedCols = Object.entries(s.cols).sort((a,b)=>b[1]-a[1]);
                                  const sortedGrads = Object.entries(s.grads).sort((a,b)=>b[1]-a[1]);
                                  const topCol = sortedCols[0]?.[0] || '-';
                                  const topGrad = sortedGrads[0]?.[0] || '-';

                                  const colTooltip = sortedCols.map(([k,v])=>`${k}: ${v}명`).join('\n');
                                  const gradTooltip = sortedGrads.map(([k,v])=>`${k}: ${v}명`).join('\n');

                                  return (
                                    <tr key={name}>
                                      <td style={{fontWeight: 700}}>{name}</td>
                                      <td style={{color:'var(--primary)', fontWeight: 700}}>{s.total}명</td>
                                      <td title={colTooltip} style={{cursor:'help', textDecoration:'underline dotted'}}>{topCol}</td>
                                      <td title={gradTooltip} style={{cursor:'help', textDecoration:'underline dotted'}}>{topGrad}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      );
                    })
                )}
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
