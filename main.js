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

        const outputPath = path.join(__dirname, 'Book1.xlsx');
        var book = xlsx.readFile(outputPath);
        const sheet2 = book.Sheets["Sheet1"];
        sheet2["D5"] = { t: "s", v: "hoge1", w: "hoge1" };
        book.Sheets["Sheet1"] = sheet2;
        xlsx.writeFile(book, outputPath);

        event.reply('excelDataWritten', outputPath);
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
