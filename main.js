const { app, BrowserWindow, ipcMain } = require('electron');
const xlsx = require('xlsx');
const path = require('path');

// Example function to read Excel file
function readExcelFile(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    return data;
}

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

    // Lắng nghe yêu cầu đọc dữ liệu Excel từ renderer process
    ipcMain.on('readExcelData', (event, filePaths) => {
        // const fullPath = fileList[0].path;
        const workbook = xlsx.readFile(filePaths[0]);
        const sheetName = '届出と附票';
        const sheet = workbook.Sheets[sheetName];
        const cellAddress = 'B11';
        const cellValue = sheet[cellAddress].v;

        // Ghi dữ liệu vào một file Excel mới
        const newWorkbook = xlsx.utils.book_new();
        const newSheet = xlsx.utils.aoa_to_sheet([[cellValue]]);
        xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');
        const outputPath = path.join(__dirname, 'Book1.xlsx');
        // const outputPath = 'C:\Users\DSN\Desktop\Gangter\shiyaku\test.xlsx'; // Đường dẫn và tên file Excel mới
        xlsx.writeFile(newWorkbook, outputPath);

        // Gửi đường dẫn của file Excel mới về renderer process nếu cần
        event.reply('excelDataWritten', outputPath);

        // Gửi dữ liệu về renderer process
        event.reply('excelData', cellValue);
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
