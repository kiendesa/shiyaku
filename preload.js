const { ipcRenderer } = require('electron');
const xlsx = require('xlsx');

document.getElementById('directoryInput').addEventListener('change', function (event) {
    const fileList = event.target.files;

    if (fileList.length > 0) {
        console.log("file list", fileList);
        //pathを表示する
        const fullPath = fileList[0].path;
        const lastSlashIndex = fullPath.lastIndexOf("\\");
        const folderPath = fullPath.substring(0, lastSlashIndex);
        document.getElementById('folder-path').innerText = "Selected folder: " + folderPath;
        const filePaths = Array.from(fileList).map(file => file.path);
        // Gửi yêu cầu đọc dữ liệu từ file Excel đến main process
        ipcRenderer.send('readExcelData', filePaths);

    } else {
        document.getElementById('folder-path').innerText = "No folder selected";
    }
});

// Lắng nghe sự kiện reply từ main process để nhận dữ liệu Excel và hiển thị
ipcRenderer.on('excelData', (event, cellValue) => {
    console.log("Data from Excel:", cellValue);
    // Hiển thị dữ liệu trên giao diện người dùng
});
