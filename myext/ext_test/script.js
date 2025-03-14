'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

        worksheet.getSummaryDataAsync().then(function (sumData) {
            let columns = sumData.columns.map(col => col.fieldName);
            let data = sumData.data;

            // Tạo hàng tiêu đề
            columns.forEach(col => {
                $('#table-header').append(`<th>${col}</th>`);
            });

            // Tạo hàng filter (dropdown cho từng cột)
            columns.forEach((col, index) => {
                let uniqueValues = [...new Set(data.map(row => row[index].formattedValue))];
                let select = `<select class="column-filter" id="filter-${index}" onchange="filterColumn(${index})">
                                <option value="">All ${col}</option>`;
                uniqueValues.forEach(value => {
                    select += `<option value="${value}">${value}</option>`;
                });
                select += `</select>`;
                $('#table-filters').append(`<th>${select}</th>`);
            });

            // Thêm dữ liệu vào bảng
            data.forEach(row => {
                let rowHTML = '<tr>';
                row.forEach(cell => {
                    rowHTML += `<td>${cell.formattedValue}</td>`;
                });
                rowHTML += '</tr>';
                $('#table-body').append(rowHTML);
            });

            // Kích hoạt DataTable với Export Excel
            let table = $('#data-table').DataTable({
                paging: true,
                searching: true,
                ordering: true,
                pageLength: 10,
                dom: 'lBfrtip',
                buttons: [
                    {
                        extend: 'excelHtml5',
                        text: 'Export to Excel',
                        title: 'Exported_Data'
                    }
                ]
            });

            // Hàm filter theo từng cột
            window.filterColumn = function (index) {
                let val = $(`#filter-${index}`).val();
                table.column(index).search(val ? `^${val}$` : '', true, false).draw();
            };
        });
    });
});
