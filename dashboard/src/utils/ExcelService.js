import * as XLSX from 'xlsx';
import { format, parse } from 'date-fns';

/**
 * 엑셀 데이터를 읽고 정제하는 서비스
 */
export const ExcelService = {
  // 엑셀 바이너리 데이터를 JSON으로 변환
  parseExcel: (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(worksheet);
        resolve(json);
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

    data.forEach(item => {
      // 참여 여부 (기본: '참여', '참석' 등 긍정어 확인)
      const statusValue = item[mapping.status] || '';
      const isAttended = statusValue.includes('참') || statusValue.includes('Y');
      if (isAttended) stats.participants++;

      // 단과대별 통계
      const college = item[mapping.college] || '기타/미정';
      stats.byCollege[college] = (stats.byCollege[college] || 0) + 1;

      // 학년별 통계
      const grade = item[mapping.grade] || '미기입';
      stats.byGrade[grade] = (stats.byGrade[grade] || 0) + 1;

      // 월별 통계 (날짜 데이터가 있을 경우)
      if (mapping.date && item[mapping.date]) {
         try {
           // 엑셀 날짜 형식 처리 (숫자 or 문자열)
           let date;
           if (typeof item[mapping.date] === 'number') {
             date = XLSX. some_date_logic; // 생략 가능성 있음, 안전하게 처리
           } else {
             const cleanedDate = String(item[mapping.date]).replace(/[^\d.-]/g, '');
             // 월 추출 (YYYY-MM-DD 등에서 MM 추출)
             const monthMatch = cleanedDate.match(/-(\d{2})-/) || cleanedDate.match(/\.(\d{2})\./);
             const month = monthMatch ? parseInt(monthMatch[1]) + '월' : '기타';
             stats.byMonth[month] = (stats.byMonth[month] || 0) + 1;
           }
         } catch(e) { /* ignore error */ }
      }
    });

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
