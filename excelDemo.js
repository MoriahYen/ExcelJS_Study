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
    // for loop
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            console.log(cell.value);
        });
    });
}

writeExcelTest();
