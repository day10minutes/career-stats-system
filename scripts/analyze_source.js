const XLSX = require('xlsx');
const fs = require('fs');

try {
    const workbook = XLSX.readFile('C:/Users/82108/.gemini/antigravity/scratch/career-stats-system/data/source.xlsx');
    const result = {
        sheets: workbook.SheetNames,
        sheetDetails: {}
    };

    workbook.SheetNames.forEach(name => {
        const sheet = workbook.Sheets[name];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        result.sheetDetails[name] = {
            headers: rows[0] || [],
            sampleRow: rows[1] || [],
            rowCount: rows.length - 1
        };
    });

    fs.writeFileSync('C:/Users/82108/.gemini/antigravity/scratch/career-stats-system/data/analysis_result.json', JSON.stringify(result, null, 2));
    console.log('Analysis Complete. Results saved to analysis_result.json');
} catch (err) {
    console.error('Analysis failed:', err.message);
}
