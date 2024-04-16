window.onload = () => {
    const electron = require('electron');
    electron.ipcRenderer.on('data', (event, excelData) => {
        console.log("excelData", excelData);
        // Sử dụng giá trị của ô Excel để render trong HTML
        const htmlContent = `<div>${excelData}</div>`;
        // Đưa nội dung HTML vào một phần tử trên trang web của bạn
        document.getElementById('dataContainer').innerHTML = htmlContent;
    });

    electron.ipcRenderer.on('excelData', (event, excelData) => {
        console.log("excelData", excelData);
        // Sử dụng giá trị của ô Excel để render trong HTML
        const htmlContent = `<div>${excelData}</div>`;
        // Đưa nội dung HTML vào một phần tử trên trang web của bạn
        document.getElementById('excelValueContainer').innerHTML = htmlContent;
    });

    // Gửi yêu cầu để nhận dữ liệu Excel
    electron.ipcRenderer.send('requestExcelData');
};