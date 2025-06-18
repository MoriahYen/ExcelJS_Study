const ExcelJs = require('exceljs');
const path = require('path');

async function writeExcelTest() {
    const workbook = new ExcelJs.Workbook();
    // await workbook.xlsx.readFile(
    //     'C:\Users\Rawst\Documents\Study\PlaywrightStudy\excelDownloadTest.xlsx',
    // );
    // gpt改的  不然會報錯
    const filePath = path.join(
        'C:',
        'Users',
        'Rawst',
        'Documents',
        'Study',
        'PlaywrightStudy',
        'excelDownloadTest.xlsx',
    );

    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet('Sheet1');
    let output = { row: -1, column: -1 };
    // for loop
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if (cell.value === 'Banana') {
                output.row = rowNumber;
                output.column = colNumber;
            }
        });
    });

    const cell = worksheet.getCell(3, 2);
    cell.value = 'Republic';
    await workbook.xlsx.writeFile(filePath);
}

writeExcelTest();
