const { app, BrowserWindow, ipcMain } = require('electron');
const xlsx = require('xlsx');
const path = require('path');
const ExcelJS = require('exceljs');
require('dotenv').config();

//データを保存するファイルのURL
const outputPath = path.join(__dirname, '..', '..', process.env.EXCEL_FILE_PATH);
const maxLength = 12;


module.exports = async function readExcelData(event, filePaths) {
    try {
        for (let index = 0; index < filePaths.length; index++) {
            console.log("index", filePaths.length);

            /*ーーーーー取得データ関数ーーーーーー*/
            // オリジナルのエクセルを読み込み
            const workbookData = new ExcelJS.Workbook();
            const workbook = await workbookData.xlsx.readFile(filePaths[index]);
            const data = await getData(workbook);

            // 選択フォルダのデータをコービーするために、新たなworkbookを作成します。
            const newWorkbook = new ExcelJS.Workbook();
            let newWorksheet = await saveFile(newWorkbook, outputPath);

            // データを年度処理件数集計ツール.xlsxにの出力シートに更新します。
            await saveData(newWorksheet, index, data)

            // データを年度処理件数集計ツール.xlsxにの出力シートに書き込みます
            await newWorkbook.xlsx.writeFile(outputPath);
            console.log('susscess....');
            event.reply('excelDataWritten', outputPath);
            // event.reply('excelData', data);
        }
        if (filePaths.length < 12) {
            const lengthFile = maxLength - filePaths.length
            for (let index = 0; index < lengthFile; index++) {

                const newWorkbook = new ExcelJS.Workbook();
                let newWorksheet = await saveFile(newWorkbook, outputPath);

                // データを年度処理件数集計ツール.xlsxにの出力シートに更新します。
                await saveData(newWorksheet, filePaths.length + index, [])

                // データを年度処理件数集計ツール.xlsxにの出力シートに書き込みます
                console.log("not enough...");
                await newWorkbook.xlsx.writeFile(outputPath);
            }
        }
    } catch (error) {
        console.error('error file Excel:', error);
    }
};

async function getData(workbook) {

    /*ーーーーー届出と附票シートにデータを取得するーーーーーー*/
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
    let countDate = 0;
    for (let m = 5; m <= lastRow; m++) {
        const date_chk = worksheet.getCell(m, 11).value;
        if (typeof date_chk === 'object' && date_chk !== null) {
            if (date_chk.result instanceof Date) {
                countDate++;
            }
        }
    }

    /*ーーーーーー戸籍集計報告・グラフ【郵送】ーーーーーーー*/
    const worksheetUse = workbook.getWorksheet('戸籍集計報告・グラフ【郵送】');
    let lastUseColumn = worksheetUse.getRow(6).actualCellCount + 1;
    let lastUseRow = worksheetUse.getColumn(lastUseColumn).values.length - 1;
    const useValue = worksheetUse.getCell(lastUseRow, lastUseColumn).value;

    /*ーーーーー戸籍集計報告・グラフ【公用】ーーーーーー*/
    const worksheetPublic = workbook.getWorksheet('戸籍集計報告・グラフ【公用】');
    let lastPulicColumn = worksheetPublic.getColumn(1).values.length;
    const publicValue = worksheetPublic.getCell('Z' + lastPulicColumn).value;

    /*ーーーーー住民票集計報告ーーーーーーー*/
    const worksheetReport = workbook.getWorksheet('住民票集計報告');
    let lastRowRport = worksheetReport.getColumn(1).values.length - 1;
    //② 公用請求分（送付分のみ
    const receivedValue = worksheetReport.getCell('F' + lastRowRport).value;
    //④ 郵送住民票返戻（該当なし）（送付分のみ）公用
    const returnValue = worksheetReport.getCell('G' + lastRowRport).value;
    //① 一般請求分（送付分のみ）
    const receivedValue2 = worksheetReport.getCell('N' + lastRowRport).value;
    //④ 郵送住民票返戻（該当なし）（送付分のみ）一般
    const returnValue2 = worksheetReport.getCell('P' + lastRowRport).value;

    return [countDate, sendValue.result, allCaseValue, processValue.result,
        useValue.result, publicValue.result, receivedValue2.result,
        receivedValue.result, returnValue2.result, returnValue.result]
}

async function saveFile(newWorkbook, outputPath) {

    //　選択フォルダのデータをコービーするために、新たなworkbookを作成します。
    const newWorksheet = newWorkbook.addWorksheet('出力シート');

    // オリジナルを読み込む
    const docbookOriginal = new ExcelJS.Workbook();
    const docbook = await docbookOriginal.xlsx.readFile(outputPath);
    const docName = '出力シート';
    const docsheet = docbook.getWorksheet(docName);

    // 色と文字フォントなどを全部コービーします。
    docsheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const newCell = newWorksheet.getCell(rowNumber, colNumber);
            Object.assign(newCell, cell); // 色や書式設定を含むすべてのセルのプロパティをコピーします
            newCell.font = Object.assign({}, cell.font); // 文字のフォントをコービーします。
        });
    });
    return newWorksheet
}

async function saveData(newWorksheet, index, data) {

    //データを書き込み
    newWorksheet.getCell(3, index + 4).value = data.length !== 0 ? data[0] : '';
    newWorksheet.getCell(5, index + 4).value = data.length !== 0 ? data[1] : '';
    newWorksheet.getCell(6, index + 4).value = data.length !== 0 ? data[2] : '';
    newWorksheet.getCell(7, index + 4).value = data.length !== 0 ? data[3] : '';
    newWorksheet.getCell(8, index + 4).value = data.length !== 0 ? data[4] : '';
    newWorksheet.getCell(9, index + 4).value = data.length !== 0 ? data[5] : '';
    newWorksheet.getCell(15, index + 4).value = data.length !== 0 ? data[6] : '';
    newWorksheet.getCell(16, index + 4).value = data.length !== 0 ? data[7] : '';
    newWorksheet.getCell(17, index + 4).value = data.length !== 0 ? data[8] : '';
    newWorksheet.getCell(18, index + 4).value = data.length !== 0 ? data[9] : '';
}