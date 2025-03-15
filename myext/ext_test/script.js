'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

        worksheet.getSummaryDataAsync().then(function (sumData) {
            let columns = sumData.columns.map(col => col.fieldName);
            let data = sumData.data;

            // Kiểm tra có cột Measure Names không
            const measureNameIndex = columns.indexOf("Measure Names");

            if (measureNameIndex !== -1) {
                // Pivot dữ liệu
                let pivotData = {};
                let dimensionCols = columns.filter((_, idx) => idx !== measureNameIndex);
                let measureCols = [...new Set(data.map(row => row[measureNameIndex].formattedValue))];

                // Chuẩn bị header
                dimensionCols.forEach(col => $('#table-header').append(`<th>${col}</th>`));
                measureCols.forEach(measure => $('#table-header').append(`<th>${measure}</th>`));

                // Pivot dữ liệu
                data.forEach(row => {
                    let key = dimensionCols.map((_, idx) => row[idx].formattedValue).join("|");
                    if (!pivotData[key]) {
                        pivotData[key] = Array(dimensionCols.length + measureCols.length).fill("");
                        dimensionCols.forEach((_, idx) => pivotData[key][idx] = row[idx].formattedValue);
                    }
                    let measureIndex = measureCols.indexOf(row[measureNameIndex].formattedValue);
                    pivotData[key][dimensionCols.length + measureIndex] = row[measureNameIndex + 1].formattedValue;
                });

                // Đưa dữ liệu vào bảng
                Object.values(pivotData).forEach(row => {
                    $('#table-body').append(`<tr>${row.map(cell => `<td>${cell}</td>`).join('')}</tr>`);
                });
            } else {

                // Tạo hàng tiêu đề
                columns.forEach(col => {
                    $('#table-header').append(`<th>${col}</th>`);
                });

                // Tạo hàng filter ngay dưới tiêu đề
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
            }

            // Kích hoạt DataTable
            let table = $('#data-table').DataTable({
                paging: true,
                searching: true,
                ordering: true,
                pageLength: 10,
                dom: '<"top-controls"lBf>rtip', // Định vị controls lên trên
                buttons: [
                    {
                        extend: 'excelHtml5',
                        text: 'Export to Excel',
                        title: 'Exported_Data'
                    }
                ]
            });

            // Di chuyển các control vào vị trí mong muốn
            $('#table-length').html($('.dataTables_length'));
            $('#table-search').html($('.dataTables_filter'));
            $('#table-export').html($('.dt-buttons'));

            // Hàm filter theo từng cột
            window.filterColumn = function (index) {
                let val = $(`#filter-${index}`).val();
                table.column(index).search(val ? `^${val}$` : '', true, false).draw();
            };
        });
    });
});
