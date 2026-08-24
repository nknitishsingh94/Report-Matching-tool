const xlsx = require('xlsx');
const ExcelJS = require('exceljs');

async function run() {
    try {
        const csv = `Call Date,,Disposition,,Total Time,Talk Time,Ring Time,Caller ID,Target Name,Target Number
7/28/2026 13:07,,ANSWER,,'00:06:53,'00:06:46,'00:00:07,13175578549,Luk 1707,18332251707`;
        
        const workbook = xlsx.read(csv, { type: 'string' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { defval: "" });

        console.log("Headers:", Object.keys(data[0]));

        const newWorkbook = new ExcelJS.Workbook();
        const updatedSheet = newWorkbook.addWorksheet('Validated Numbers');
        
        const headers = Object.keys(data[0]);
        updatedSheet.columns = headers.map(header => ({ header: header, key: header, width: 20 }));
        
        data.forEach(rowData => {
            const row = updatedSheet.addRow(rowData);
        });

        console.log("Success");
    } catch (e) {
        console.error(e);
    }
}
run();
