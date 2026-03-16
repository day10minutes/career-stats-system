const XLSX = require('xlsx');

function testParsing(sheetNames) {
    console.log(`Testing with sheets: ${sheetNames.join(', ')}`);
    
    // Simulate workbook
    const workbook = {
        SheetNames: sheetNames,
        Sheets: {}
    };
    
    sheetNames.forEach(name => {
        workbook.Sheets[name] = XLSX.utils.json_to_sheet([{ data: name + '_row1' }]);
    });

    // Mock logic from ExcelService.js
    const allSheetName = workbook.SheetNames.find(name => name.toUpperCase() === 'ALL');
    let combinedData = [];
    
    if (allSheetName) {
        const worksheet = workbook.Sheets[allSheetName];
        combinedData = XLSX.utils.sheet_to_json(worksheet);
    } else {
        workbook.SheetNames.forEach(sheetName => {
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);
            combinedData = combinedData.concat(json);
        });
    }
    
    console.log('Result length:', combinedData.length);
    console.log('Result data:', JSON.stringify(combinedData));
    console.log('---');
}

testParsing(['ALL', 'Sheet1']);
testParsing(['Sheet1', 'Sheet2']);
testParsing(['all', 'Sheet1']); // Case sensitivity check
