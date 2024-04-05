const { app, BrowserWindow, ipcMain } = require('electron');
const xlsx = require('xlsx');
const path = require('path');

// Example function to read Excel file
function readExcelFile(filePath) {
    const workbook = xlsx.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
    console.log(data);
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

    const filePath = path.join(__dirname, 'test.xlsx');
    const excelData = readExcelFile(filePath);

    // Gửi dữ liệu sang renderer process khi cửa sổ đã được tạo
    mainWindow.webContents.on('did-finish-load', () => {
        mainWindow.webContents.send('sendExcelData', excelData);
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
