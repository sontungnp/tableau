'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

        worksheet.getSummaryDataAsync().then(function (sumData) {
            let columns = sumData.columns.map(col => col.fieldName);
            let data = sumData.data;

            // Thêm header
            columns.forEach(col => {
                $('#table-header').append(`<th>${col}</th>`);
            });

            // Thêm dữ liệu
            data.forEach(row => {
                let rowHTML = '<tr>';
                row.forEach(cell => {
                    rowHTML += `<td>${cell.formattedValue}</td>`;
                });
                rowHTML += '</tr>';
                $('#table-body').append(rowHTML);
            });

            // Kích hoạt DataTables.js
            $('#data-table').DataTable({
                paging: true,      // Phân trang
                searching: true,   // Cho phép filter
                ordering: true,    // Cho phép sắp xếp
                pageLength: 10     // Số hàng mỗi trang
            });
        });
    });
});
