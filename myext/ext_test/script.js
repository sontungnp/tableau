'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

        worksheet.getSummaryDataAsync().then(function (sumData) {
            let columns = sumData.columns.map(col => col.fieldName);
            let data = sumData.data;

            // Thêm header cột
            columns.forEach(col => {
                $('#table-header').append(`<th>${col}</th>`);
                $('#filter-row').append(`<th><select id="filter-${col}" onchange="filterColumn('${col}')">
                    <option value="">All</option>
                </select></th>`);
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

            // Khởi tạo DataTable với Export Excel
            let table = $('#data-table').DataTable({
                paging: true,
                searching: true,
                ordering: true,
                pageLength: 10,
                dom: 'Bfrtip',
                buttons: [
                    {
                        extend: 'excelHtml5',
                        text: 'Export to Excel',
                        title: 'Exported_Data'
                    }
                ]
            });

            // Lấy danh sách giá trị duy nhất và thêm vào filter
            columns.forEach((col, index) => {
                let uniqueValues = [...new Set(data.map(row => row[index].formattedValue))];
                uniqueValues.forEach(value => {
                    $(`#filter-${col}`).append(`<option value="${value}">${value}</option>`);
                });
            });

            // Hàm filter cột
            window.filterColumn = function (col) {
                let colIndex = columns.indexOf(col);
                let val = $(`#filter-${col}`).val();
                table.column(colIndex).search(val ? `^${val}$` : '', true, false).draw();
            };
        });
    });
});
