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
            const sheetName = '届出と附票';
            const worksheet = workbook.getWorksheet(sheetName);
            const cellValue = worksheet.getCell('B11').value.result;
            const cellValue1 = worksheet.getCell('O28').value.result;

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

            // Cập nhật giá trị của ô D5 và D6 trong workbook mới
            const newCellD5 = newWorksheet.getCell('D5');
            const newCellD6 = newWorksheet.getCell('D6');
            newCellD5.value = cellValue;
            newCellD6.value = cellValue1;

            // Lưu workbook mới vào file mới
            await newWorkbook.xlsx.writeFile(outputPath);

            console.log('Dữ liệu đã được cập nhật và ghi vào file Excel mới thành công.');
            event.reply('excelDataWritten', outputPath);
            event.reply('excelData', cellValue);
        } catch (error) {
            console.error('Lỗi khi thao tác với file Excel:', error);
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
