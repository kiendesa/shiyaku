window.onload = () => {
    console.log("uuuuu");
    const electron = require('electron');
    electron.ipcRenderer.on('sendExcelData', (event, excelData) => {
        const tableContainer = document.getElementById('excel-table');
        const table = document.createElement('table');
        const thead = document.createElement('thead');
        const tbody = document.createElement('tbody');

        // Create table header
        const headerRow = document.createElement('tr');
        Object.keys(excelData[0]).forEach(key => {
            const th = document.createElement('th');
            th.textContent = key;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow);

        // Create table body
        excelData.forEach(row => {
            const tr = document.createElement('tr');
            Object.values(row).forEach(value => {
                const td = document.createElement('td');
                td.textContent = value;
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });

        // Add thead and tbody to table
        table.appendChild(thead);
        table.appendChild(tbody);

        // Add table to container
        tableContainer.appendChild(table);
    });

    // Gửi yêu cầu để nhận dữ liệu Excel
    electron.ipcRenderer.send('requestExcelData');
};