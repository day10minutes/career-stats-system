import * as XLSX from 'xlsx';
import { format, parse } from 'date-fns';

const COLLEGE_LIST = [
  '인문과학대학', '사회과학대학', '자연과학대학', '공과대학(엘텍포함)', 
  '음악대학', '조형예술대학', '사범대학', '경영대학', '신산업융합대학', 
  '의과대학', '간호대학', '약학대학', '스크랜튼대학', '인공지능대학', 
  '호크마교양대학', '건강과학대학'
];

/**
 * 엑셀 데이터를 읽고 정제하는 서비스
 */
export const ExcelService = {
  // 엑셀 바이너리 데이터를 JSON으로 변환
  parseExcel: (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          const result = {
            workbook: {},
            sheetNames: []
          };
          
          let combinedData = [];
          const hasExistingAll = workbook.SheetNames.some(name => name.toUpperCase() === 'ALL');
          
          workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            result.workbook[sheetName] = json;
            result.sheetNames.push(sheetName);
            
            // 만약 ALL 시트가 파일에 이미 있다면 그것이 합쳐진 데이터일 수 있으므로 그대로 사용
            // 만약 ALL 시트가 없다면 모든 시트를 합침
            if (!hasExistingAll) {
              combinedData = combinedData.concat(json);
            }
          });
          
          // ALL 시트가 없는 경우 가상의 'ALL' 시트 추가
          if (!hasExistingAll) {
            result.workbook['ALL'] = combinedData;
            result.sheetNames.unshift('ALL');
          } else {
            // 이미 ALL 시트가 있는 경우, 해당 시트의 이름을 대문자 'ALL'로 표준화하거나 그대로 둠
            // 여기서는 'ALL'이 목록의 맨 앞에 오도록 조정
            const allIndex = result.sheetNames.findIndex(n => n.toUpperCase() === 'ALL');
            const actualName = result.sheetNames.splice(allIndex, 1)[0];
            result.sheetNames.unshift(actualName);
          }
          
          resolve(result);
        } catch (error) {
          reject(error);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  },

  // 동적 통계 산출 로직
  calculateStats: (data, mapping) => {
    if (!data || data.length === 0) return null;

    const stats = {
      total: data.length,
      participants: 0,
      applicants: data.length,
      byCollege: {},
      byGrade: {},
      byMonth: {},
      raw: data
    };

    const matrix = {};
    const programMatrix = {};
    const collegeDetails = {}; 
    const studentTrace = {};
    const netUniqueSet = new Set();
    let needsCheckCount = 0;

    data.forEach(item => {
      // 시트별/행별 헤더 명칭이 다를 수 있으므로 동적 매핑 보정
      const keys = Object.keys(item);
      const find = (keywords) => {
        for (const kw of keywords) {
          const match = keys.find(k => k.includes(kw));
          if (match) return match;
        }
        return null;
      };
      const m_row = {
        college: find(['단과대', '소속', '대학']),
        grade: find(['학년']),
        status: find(['참석', '참여여부', '상태', '결과']),
        date: find(['일정', '일시', '일자', '날짜', '시간', '시점']),
        program: find(['프로그램명', '프로그램']),
        education: find(['교육명', '유형', '프로그램명']),
        studentId: find(['학번', 'ID']),
        name: find(['성명', '이름']),
        major: find(['전공', '학과'])
      };

      // 참여 여부 확인
      const sVal = String(item[m_row.status] || '').trim();
      let isAttended = false;
      if (sVal === '') {
        isAttended = true; // 빈칸은 참여로
        needsCheckCount++;
      } else if (sVal.includes('불참')) {
        isAttended = false;
      } else if (sVal.includes('참') || sVal.includes('Y') || sVal.includes('완료')) {
        isAttended = true;
      }
      
      if (isAttended) {
        stats.participants++;

        // 단과대 분류
        const rawCollege = item[m_row.college] || '';
        let college = '기타';
        
        if (rawCollege.includes('공과대학') || rawCollege.includes('엘텍')) {
            college = '공과대학(엘텍포함)';
        } else {
            for (const cName of COLLEGE_LIST) {
                if (rawCollege.includes(cName)) {
                    college = cName;
                    break;
                }
            }
        }
        stats.byCollege[college] = (stats.byCollege[college] || 0) + 1;

        // 단과대별 전공/원래 이름 집계
        const major = item[m_row.major] || '미기입 전공';
        if (!collegeDetails[college]) collegeDetails[college] = {};
        
        let detailKey = major;
        if (college === '기타') detailKey = rawCollege || '정보없음';
        if (college === '공과대학(엘텍포함)') {
            const subName = rawCollege.includes('엘텍') ? '엘텍공과' : '공과대학';
            detailKey = `[${subName}] ${major}`;
        }
        collegeDetails[college][detailKey] = (collegeDetails[college][detailKey] || 0) + 1;

        // 학년별 통계 (숫자/문자열 혼합 처리)
        let grade = item[m_row.grade];
        if (grade !== undefined && grade !== null) {
          grade = String(grade).trim();
          if (!grade.includes('학년') && !isNaN(grade) && grade !== '') grade = grade + '학년';
        } else {
          grade = '미기입';
        }
        stats.byGrade[grade] = (stats.byGrade[grade] || 0) + 1;

        // 날짜/월 추출
        let month = "일정 미확인";
        const dateValue = item[m_row.date];
        if (dateValue) {
          if (typeof dateValue === 'number') {
            const date = new Date((dateValue - 25569) * 86400 * 1000);
            month = (date.getMonth() + 1) + '월';
          } else {
            const cleaned = String(dateValue).replace(/\s+/g, '');
            const parts = cleaned.match(/\d+/g);
            if (parts && parts.length > 0) {
              let m = -1;
              if (parts.length >= 3) m = parseInt(parts[1]);
              else if (parts.length === 2) {
                const p0 = parseInt(parts[0]);
                const p1 = parseInt(parts[1]);
                if (p0 > 12) m = p1;
                else m = p0;
              } else if (parts.length === 1) {
                m = parseInt(parts[0]);
              }
              if (m >= 1 && m <= 12) month = m + '월';
            }
          }
        }
        stats.byMonth[month] = (stats.byMonth[month] || 0) + 1;

        // 프로그램/교육명
        const prog = item[m_row.program] || item[m_row.education] || '기타 프로그램';
        const edu = item[m_row.education] || item[m_row.program] || '기타 유형';
        const studentId = String(item[m_row.studentId] || '미기입');
        const name = item[m_row.name] || '미기입';

        if (!studentTrace[studentId]) studentTrace[studentId] = { name: name, count: 0 };
        studentTrace[studentId].count++;
        netUniqueSet.add(`${studentId}|${prog}`);

        // 월별 메트릭스 (일반 시트용)
        if (!matrix[month]) matrix[month] = {};
        if (!matrix[month][edu]) matrix[month][edu] = { total: 0, cols: {}, grads: {} };
        matrix[month][edu].total++;
        matrix[month][edu].cols[college] = (matrix[month][edu].cols[college] || 0) + 1;
        matrix[month][edu].grads[grade] = (matrix[month][edu].grads[grade] || 0) + 1;

        // 프로그램별 메트릭스 (ALL 시트용)
        if (!programMatrix[prog]) programMatrix[prog] = { total: 0, cols: {}, grads: {}, months: {}, trace: {} };
        programMatrix[prog].total++;
        programMatrix[prog].cols[college] = (programMatrix[prog].cols[college] || 0) + 1;
        programMatrix[prog].grads[grade] = (programMatrix[prog].grads[grade] || 0) + 1;
        programMatrix[prog].months[month] = (programMatrix[prog].months[month] || 0) + 1;
        programMatrix[prog].trace[studentId] = (programMatrix[prog].trace[studentId] || 0) + 1;
      }
    });

    // 최종 통계 정돈
    const sortedCols = Object.entries(stats.byCollege).sort((a,b)=>b[1]-a[1]);
    const sortedGrads = Object.entries(stats.byGrade).sort((a,b)=>b[1]-a[1]);
    const sortedMonths = Object.entries(stats.byMonth)
        .filter(([m]) => m && m.includes('월') && !m.includes('미확인'))
        .sort((a,b)=>b[1]-a[1]);
    
    stats.topMonth = sortedMonths[0]?.[0] || (stats.byMonth['일정 미확인'] ? '일정 미확인' : '-');
    stats.topCollege = sortedCols[0]?.[0] || '-';
    stats.topGrade = sortedGrads[0]?.[0] || '-';
    const sortedProgs = Object.entries(programMatrix).sort((a,b) => b[1].total - a[1].total);
    stats.topProgram = sortedProgs[0]?.[0] || '-';
    stats.netUniqueCount = netUniqueSet.size;
    stats.duplicateUserCount = Object.values(studentTrace).filter(s => s.count > 1).length;

    // 중복 참여가 가장 많은 프로그램 찾기 (ALL 시트용)
    let maxDupeProg = '-';
    let maxDupeCount = 0;
    Object.entries(programMatrix).forEach(([name, data]) => {
        const dupeInProg = Object.values(data.trace || {}).filter(c => c > 1).length;
        if (dupeInProg > maxDupeCount) {
            maxDupeCount = dupeInProg;
            maxDupeProg = name;
        }
    });
    stats.mostDuplicatedProgram = maxDupeProg;

    stats.matrix = matrix;
    stats.programMatrix = programMatrix;
    stats.collegeDetails = collegeDetails;
    stats.needsCheckCount = needsCheckCount;
    return stats;
  },

  // 엑셀 내보내기
  exportToExcel: (data, fileName = '통계데이터.xlsx') => {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Results");
    XLSX.writeFile(workbook, fileName);
  }
};
