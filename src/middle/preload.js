const { ipcRenderer } = require('electron');
const path = require('path');

let year = 0;
document.getElementById('directoryInput').addEventListener('change', function (event) {
    const fileList = event.target.files;

    if (fileList.length > 0) {
        const fileExtensions = [".xls", ".xlsx", ".xlsm"]; // Các phần mở rộng của tệp Excel
        let isValid = true;

        for (let i = 0; i < fileList.length; i++) {
            const fileName = fileList[i].name;
            const fileExtension = fileName.slice((fileName.lastIndexOf(".") - 1 >>> 0) + 2).toLowerCase(); // Lấy phần mở rộng của tệp

            if (!fileExtensions.includes("." + fileExtension)) {
                isValid = false;
                break;
            }
        }

        if (isValid) {
            const fullPath = fileList[0].path;
            const lastSlashIndex = fullPath.lastIndexOf("\\");
            const folderPath = fullPath.substring(0, lastSlashIndex);
            year = path.basename(folderPath)
            document.getElementById('folder-path').innerText = "選択フォルダURL: " + folderPath;
            const filePaths = Array.from(fileList).map(file => file.path);
            for (let index = 0; index < 12; index++) {
                document.getElementById(`${index + 4}_month`).innerText = ''
            }
            // Gửi yêu cầu đọc dữ liệu từ file Excel đến main process
            ipcRenderer.send('readExcelData', filePaths);

        } else {
            alert("Excel (.xls, .xlsx, .xlsm)のみが受け入れられます。");
        }

    } else {
        document.getElementById('folder-path').innerText = "No folder selected";
    }
});


document.getElementById('printPDFButton').addEventListener('click', () => {
    // IPC経由でメインプロセスにリクエストを送信
    ipcRenderer.send('printPDF', year);
});


// Lắng nghe sự kiện reply từ main process để nhận dữ liệu Excel và hiển thị
ipcRenderer.on('showFile', (event, filePaths, index) => {
    document.getElementById(`${index + 4}_month`).innerText = filePaths ? path.basename(filePaths) : '';
});
