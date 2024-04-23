const { BrowserWindow, ipcMain } = require('electron');
const path = require('path');
const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
require('dotenv').config();

//データを保存するファイルのURL
const outputPath = path.join(__dirname, '..', '..', process.env.EXCEL_FILE_PATH);
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
        await page.goto(`file://${path.join(__dirname, '../template/html/anken.html')}`);
        const elements = [
            //年度
            { id: 'year_tile', value: year },
            { id: 'year', value: year },

            // PDFの上のデータ
            { id: 'dataContainerA', value: dataPdf.sumTotalA },
            { id: 'dataProcessValue', value: dataPdf.totals.totalProcessValue },
            { id: 'dataSendValue', value: dataPdf.totals.totalSendValue },
            { id: 'dataCaseValue', value: dataPdf.totals.totalAllCaseValue },
            { id: 'dataUseValue', value: dataPdf.totals.totalUsevalue },
            { id: 'dataPublicValue', value: dataPdf.totals.totalPublicValue },
            { id: 'dataSumA', value: dataPdf.sumA },

            //　PDFの下のデータ
            { id: 'dataContainerB', value: dataPdf.sumTotalB },
            { id: 'dataReceivedValue', value: dataPdf.totals.totalReceivedValue },
            { id: 'dataReturndValue', value: dataPdf.totals.totalReturndValue },
            { id: 'dataReceivedPublic', value: dataPdf.totals.totalReceivedPublic },
            { id: 'dataReturndPublic', value: dataPdf.totals.totalReturndPublic },
            { id: 'dataSumB', value: dataPdf.sumB },
            //　日
            { id: 'dataOfDate', value: dataPdf.totals.totalDate },

        ];

        // idに対応するデータを割り当てる
        for (const element of elements) {
            await page.evaluate(({ id, value }) => {
                document.getElementById(id).innerText = value;
            }, element);
        }


        // PDFを印刷すること
        const pdfPath = path.join(__dirname, '..', '..', 'pdf', `${year}_年度.pdf`);
        await page.pdf({ path: pdfPath, format: 'A4', printBackground: true });

        console.log('PDF created:', pdfPath);
        event.reply('notifySuceess');

        // ブラウザが閉まる
        await browser.close();
    } catch (error) {
        console.error('Error creating PDF:', error);
    }
};


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

    return { totals, sumA, sumB, sumTotalA, sumTotalB };
}
