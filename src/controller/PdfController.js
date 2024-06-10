const { BrowserWindow, ipcMain, app } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
require('dotenv').config();
const fs = require('fs');
const { exec } = require('child_process');
const os = require('os');
const handlebars = require('handlebars');


//データを保存するファイルのURL
const outputPath = path.join(__dirname, '..', '..', '年度処理件数集計ツール.xlsx');
const maxLength = 12;

module.exports = async function printPDF(event, year) {

    try {

        const dataPdf = await caulateDataforPdf()
        console.log('Data from cell D6:', dataPdf.totals, dataPdf.sumA, dataPdf.sumB, dataPdf.sumTotalA, dataPdf.sumTotalB);
        // Puppeteerのブラウザを作成する
        const browser = await puppeteer.launch();

        // ブラウザでページを新たな開く
        const page = await browser.newPage();

        //　HTMLファイルにデータを書き込むこと
        const htmlFilePath = path.join(app.getAppPath(), 'src/template/html/anken.html');
        const htmlContent = fs.readFileSync(htmlFilePath, 'utf8');

        const template = handlebars.compile(htmlContent);
        const renderedHtml = template({ dataPdf, year });

        const cssFilePath = path.join(app.getAppPath(), 'src/template/css/anken.css');
        const cssContent = fs.readFileSync(cssFilePath, 'utf8');

        await page.setContent(renderedHtml, { waitUntil: 'domcontentloaded' });
        await page.addStyleTag({ url: 'https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css' });
        await page.addStyleTag({ content: cssContent, type: 'text/css' });

        // PDFを印刷すること
        const pdfPath = path.join(app.getPath('downloads'), `【電算】_${year}年度処理件数実績.pdf`);
        await page.pdf({ path: pdfPath, format: 'A4', printBackground: true });

        console.log('PDF created:', pdfPath);
        openPdf(pdfPath);
        event.reply('notifySuceess');

        // ブラウザが閉まる
        await browser.close();
    } catch (error) {
        console.error('Error creating PDF:', error);
    }
};

function openPdf(pdfPath) {
    if (os.platform() === 'win32' || os.platform() === 'win64') {
        exec(`start ${pdfPath}`);
    } else if (os.platform() === 'darwin') {
        exec(`open ${pdfPath}`);
    } else {
        exec(`xdg-open ${pdfPath}`);
    }
}

async function caulateDataforPdf() {

    const workbookRender = new ExcelJS.Workbook();
    await workbookRender.xlsx.readFile(outputPath);
    const worksheetRender = workbookRender.getWorksheet('出力シート');

    let totals = {
        totalDate: 0,
        totalProcessValue: 0,
        totalSendValue: 0,
        totalAllCaseValue: 0,
        totalUsevalue: 0,
        totalPublicValue: 0,
        totalReceivedValue: 0,
        totalReturndValue: 0,
        totalReceivedPublic: 0,
        totalReturndPublic: 0
    };

    for (let index = 0; index < maxLength; index++) {
        const values = [
            worksheetRender.getCell(3, index + 4).value, // totalDate
            worksheetRender.getCell(5, index + 4).value, // totalAllCaseValue
            worksheetRender.getCell(6, index + 4).value, // totalProcessValue
            worksheetRender.getCell(7, index + 4).value, // totalSendValue
            worksheetRender.getCell(8, index + 4).value, // totalUsevalue
            worksheetRender.getCell(9, index + 4).value, // totalPublicValue
            worksheetRender.getCell(15, index + 4).value, // totalReceivedValue
            worksheetRender.getCell(16, index + 4).value, // totalReturndValue
            worksheetRender.getCell(17, index + 4).value, // totalReceivedPublic
            worksheetRender.getCell(18, index + 4).value  // totalReturndPublic
        ];

        Object.keys(totals).forEach((key, i) => {
            if (values[i] !== undefined) {
                totals[key] += values[i];
            }
        });
    }

    let sumA = parseInt(totals.totalUsevalue) + parseInt(totals.totalPublicValue);
    let sumB = parseInt(totals.totalReceivedPublic) + parseInt(totals.totalReturndPublic);
    // 総件数
    let sumTotalA = parseInt(totals.totalProcessValue) + parseInt(totals.totalSendValue) + parseInt(totals.totalAllCaseValue)
        + parseInt(totals.totalUsevalue) + parseInt(totals.totalPublicValue)
    // 総件数
    let sumTotalB = parseInt(totals.totalReceivedValue) + parseInt(totals.totalReturndValue) + parseInt(totals.totalReceivedPublic)
        + parseInt(totals.totalReturndPublic)
    sumA = sumA.toLocaleString('ja-JP');
    sumB = sumB.toLocaleString('ja-JP');
    sumTotalA = sumTotalA.toLocaleString('ja-JP');
    sumTotalB = sumTotalB.toLocaleString('ja-JP');
    totals.totalProcessValue = totals.totalProcessValue.toLocaleString('ja-JP');
    totals.totalSendValue = totals.totalSendValue.toLocaleString('ja-JP');
    totals.totalAllCaseValue = totals.totalAllCaseValue.toLocaleString('ja-JP');
    totals.totalUsevalue = totals.totalUsevalue.toLocaleString('ja-JP');
    totals.totalPublicValue = totals.totalPublicValue.toLocaleString('ja-JP');
    totals.totalReceivedValue = totals.totalReceivedValue.toLocaleString('ja-JP');
    totals.totalReturndValue = totals.totalReturndValue.toLocaleString('ja-JP');
    totals.totalReceivedPublic = totals.totalReceivedPublic.toLocaleString('ja-JP');
    totals.totalReturndPublic = totals.totalReturndPublic.toLocaleString('ja-JP');
    return { totals, sumA, sumB, sumTotalA, sumTotalB };
}
