const { app, BrowserWindow, ipcMain } = require('electron');
const printPDF = require('./src/controller/PdfController');
const readExcelData = require('./src/controller/ReadDataController');
const path = require('path');

async function createWindow() {
    const mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            contextIsolation: false,
            nodeIntegration: true
        }
    });

    const indexPath = path.join(__dirname, 'src', 'template', 'html', 'index.html');
    mainWindow.loadFile(indexPath);

}

app.whenReady().then(createWindow);

// データを読み込む、書き込む
ipcMain.on('readExcelData', async (event, filePaths) => {
    await readExcelData(event, filePaths);
});

// PDFにデータを書き込む
ipcMain.on('printPDF', async (event, year) => {
    await printPDF(event, year);
});

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
