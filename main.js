const { app, BrowserWindow, ipcMain } = require('electron');
const xlsx = require('xlsx');
const path = require('path');
const ExcelJS = require('exceljs');


function createWindow() {
    const mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            contextIsolation: false,
            nodeIntegration: true
        }
    });

    mainWindow.loadFile('index.html');

    ipcMain.on('readExcelData', async (event, filePaths) => {
        try {
            const outputPath = path.join(__dirname, 'Book1.xlsx'); // Đường dẫn cho file mới

            // Đọc workbook gốc
            const workbookData = new ExcelJS.Workbook();
            const workbook = await workbookData.xlsx.readFile(filePaths[0]);

            /*--- 届出と附票シートにデータを取得する---*/
            const worksheet = workbook.getWorksheet('届出と附票');
            //② 人口動態処理業務（全件）
            const allCaseValue = worksheet.getCell('B11').value.result;
            //③ 附表関連処理業務
            let processValue = null;
            worksheet.getColumn('K').eachCell({ includeEmpty: false }, function (cell, rowNumber) {
                if (cell.value === '合計') {
                    processValue = worksheet.getCell('O' + rowNumber).value;
                    return false;
                }
            });
            //① 届出処理業務（送付分のみ）
            let sendValue = null;
            worksheet.getColumn('A').eachCell({ includeEmpty: false }, function (cell, rowNumber) {
                if (cell.value === '合計') {
                    sendValue = worksheet.getCell('B' + rowNumber).value;
                    return false;
                }
            });
            // 営業日
            let lastRow = worksheet.getColumn(1).values.length;
            // Duyệt qua từng hàng từ 5 đến lastRow
            let count = 0;
            for (let m = 5; m <= lastRow; m++) {
                const date_chk = worksheet.getCell(m, 11).value;
                if (typeof date_chk === 'object' && date_chk !== null) {
                    if (date_chk.result instanceof Date) {
                        count = count + 1;
                    }
                }
            }
            // const sendValue = worksheet.getCell('B46').value.result;

            // Tạo một workbook mới và sao chép dữ liệu từ workbook hiện tại
            const newWorkbook = new ExcelJS.Workbook();
            const newWorksheet = newWorkbook.addWorksheet('Sheet1');

            // Đọc workbook đích
            const docbookOriginal = new ExcelJS.Workbook();
            const docbook = await docbookOriginal.xlsx.readFile(outputPath);
            const docName = 'Sheet1';
            const docsheet = docbook.getWorksheet(docName);

            // Sao chép dữ liệu từ workbook đích sang workbook mới
            docsheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
                row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                    const newCell = newWorksheet.getCell(rowNumber, colNumber);
                    Object.assign(newCell, cell); // Sao chép tất cả các thuộc tính của ô, bao gồm cả màu sắc và định dạng
                    newCell.font = Object.assign({}, cell.font); // Sao chép cả phông chữ
                });
            });

            // Cập nhật giá trị của ô D5 và D6, D7 trong workbook mới
            const newCellD5 = newWorksheet.getCell('D5');
            const newCellD6 = newWorksheet.getCell('D6');
            const newCellD7 = newWorksheet.getCell('D7');
            newCellD5.value = sendValue.result;
            newCellD6.value = allCaseValue;
            newCellD7.value = processValue.result;

            // Lưu workbook mới vào file mới
            await newWorkbook.xlsx.writeFile(outputPath);

            console.log('susscess....');
            event.reply('excelDataWritten', outputPath);
            event.reply('excelData', sendValue);
        } catch (error) {
            console.error('error file Excel:', error);
        }
    });
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});
