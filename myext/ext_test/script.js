'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

        worksheet.getSummaryDataAsync().then(function (sumData) {
            let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

            worksheet.getSummaryDataAsync().then(function (sumData) {
                let columns = sumData.columns.map(col => col.fieldName);
                let data = sumData.data;

                // Tạo header
                columns.forEach(col => {
                    $('#table-header').append(`<th>${col}</th>`);
                });

                // Tạo dropdown filter
                columns.forEach((col, index) => {
                    let uniqueValues = [...new Set(data.map(row => row[index].formattedValue))];
                    let select = `<select id="filter-${index}" onchange="filterColumn(${index})">
                                    <option value="">All ${col}</option>`;
                    uniqueValues.forEach(value => {
                        select += `<option value="${value}">${value}</option>`;
                    });
                    select += `</select>`;
                    $('#filters').append(select);
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
                    dom: 'Bfrtip',
                    buttons: [
                        {
                            extend: 'excelHtml5',
                            text: 'Export to Excel',
                            title: 'Exported_Data'
                        }
                    ]
                });

                // Hàm filter theo cột
                window.filterColumn = function (index) {
                    let val = $(`#filter-${index}`).val();
                    table.column(index).search(val ? `^${val}$` : '', true, false).draw();
                };
            });
        });
    });
});
