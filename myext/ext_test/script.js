'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];

        worksheet.getSummaryDataAsync().then(function (sumData) {
            let columns = sumData.columns.map(col => col.fieldName);
            let data = sumData.data;

            // Kiểm tra có cột Measure Names không
            const measureNameIndex = columns.indexOf("Measure Names");
            const measureValueIndex = columns.indexOf("Measure Values");

            // Bỏ các cột không cần thiết như "AGG(index)"
            const validDimensionCols = columns.filter((col, idx) =>
                !col.includes("AGG") && idx !== measureNameIndex && idx !== measureValueIndex
            );

            let measureCols = [];
            if (measureNameIndex !== -1) {
                // Tìm tất cả các giá trị measure
                measureCols = [...new Set(data.map(row => row[measureNameIndex].formattedValue))];

                // Tạo header và filter (bỏ cột không cần thiết)
                validDimensionCols.forEach(col => {
                    $('#table-header').append(`<th>${col}</th>`);
                    $('#table-filters').append(`<th><input type="text" class="column-filter" placeholder="Filter ${col}" /></th>`);
                });
                measureCols.forEach(measure => {
                    $('#table-header').append(`<th>${measure}</th>`);
                    $('#table-filters').append(`<th><input type="text" class="column-filter" placeholder="Filter ${measure}" /></th>`);
                });

                // Pivot dữ liệu
                let pivotData = {};
                data.forEach(row => {
                    let dimensionKey = validDimensionCols.map(col =>
                        row[columns.indexOf(col)].formattedValue
                    ).join("|");

                    if (!pivotData[dimensionKey]) {
                        pivotData[dimensionKey] = {};
                        validDimensionCols.forEach(col => {
                            pivotData[dimensionKey][col] = row[columns.indexOf(col)].formattedValue;
                        });
                        measureCols.forEach(measure => {
                            pivotData[dimensionKey][measure] = "";
                        });
                    }

                    let measureName = row[measureNameIndex].formattedValue;
                    let measureValue = row[measureValueIndex].formattedValue;
                    pivotData[dimensionKey][measureName] = measureValue;
                });

                // Hiển thị dữ liệu pivot trong bảng
                Object.values(pivotData).forEach(row => {
                    let rowHTML = "<tr>";
                    validDimensionCols.forEach(col => rowHTML += `<td>${row[col]}</td>`);
                    measureCols.forEach(measure => rowHTML += `<td>${row[measure]}</td>`);
                    rowHTML += "</tr>";
                    $('#table-body').append(rowHTML);
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
