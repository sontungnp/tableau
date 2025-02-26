'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        const dashboard = tableau.extensions.dashboardContent.dashboard;

        setToDateDefaultValue();

        document.getElementById("refreshButton").addEventListener("click", () => {
            updateAndRefreshData();
        });

        document.getElementById("resetButton").addEventListener("click", () => {
            resetData();
        });

        function updateAndRefreshData() {
            setLoading(true); // Bắt đầu loading
            const endDate = document.getElementById('endDate').value;

            // lưu vào localstorage
            // console.log('start luu vao localstoreage');
            localStorage.setItem("endDate", endDate);
            // console.log('end luu vao localstoreage');

            // Chuyển giá trị sang định dạng yyyymmdd
            const date = new Date(endDate);
            const formattedEndDate = date.getFullYear().toString() +
                (date.getMonth() + 1).toString().padStart(2, '0') +
                date.getDate().toString().padStart(2, '0');
            // Nối hai giá trị với dấu phẩy
            const fromDateToDate = `${formattedEndDate}`;

            // Get the parameters P_Fd_Td
            dashboard.getParametersAsync().then(parameters => {
                const fdtdParam = parameters.find(param => param.name === 'P_Fd_Td');
                // Cập nhật P_FD bằng giá trị của fromDateToDate
                fdtdParam.changeValueAsync(fromDateToDate).then(() => {
                    console.log('thay doi tham so P_Fd_Td');
                    setLoading(false); // Kết thúc loading
                });
            }).catch(err => {
                setLoading(false); // Kết thúc loading
                console.error("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String).");
                alert("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String)");
                // console.error("P_Fd_Td = " + fromDateToDate);
                // alert("P_Fd_Td = " + fromDateToDate);
                console.error(err);
            });
        }

        function resetData() {
            setLoading(false); // Kết thúc loading
            document.getElementById('endDate').value = null;

            // lưu vào localstorage
            // console.log('start luu vao localstoreage');
            localStorage.setItem("endDate", null);
            localStorage.setItem("departmentCode", 'ALL');
            // console.log('end luu vao localstoreage');
            
            const fromDateToDate = 10000101;
            // Get the parameters P_Fd_Td
            dashboard.getParametersAsync().then(parameters => {
                const fdtdParam = parameters.find(param => param.name === 'P_Fd_Td');
                // Cập nhật P_FD bằng giá trị của fromDateToDate
                fdtdParam.changeValueAsync(fromDateToDate).then(() => {
                    console.log('thay doi tham so P_Fd_Td');
                    setLoading(false); // Kết thúc loading
                }).catch(err => {
                    setLoading(false); // Kết thúc loading
                    console.error("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String).");
                    alert("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String)");
                    console.error(err);
                });
                
            }).catch(err => {
                setLoading(false); // Kết thúc loading
                console.error("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String).");
                alert("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String)");
                console.error(err);
            });

            clearAllFilters();
        }

        // Thêm trạng thái loading
        function setLoading(isLoading) {
            if (isLoading) {
                document.body.style.cursor = 'wait'; // Thay đổi con trỏ chuột sang "loading"
                refreshButton.classList.add('loading'); // Thêm class loading
                refreshButton.disabled = true; // Vô hiệu hóa nút khi đang loading
            } else {
                document.body.style.cursor = 'default'; // Trả lại trạng thái bình thường
                refreshButton.classList.remove('loading'); // Bỏ class loading
                refreshButton.disabled = false; // Kích hoạt lại nút
            }
        }

        function clearAllFilters() {
            dashboard.worksheets.forEach((worksheet) => {
                worksheet.getFiltersAsync().then((filters) => {
                    filters.forEach((filter) => {
                        if (filter.fieldName !== "Trans_Date_Calc" && filter.fieldName !== "Measure Names" ) { 
                            worksheet.clearFilterAsync(filter.fieldName);
                        }
                    });
                });
            });
        }

        function setToDateDefaultValue() {
            // console.log('start setFromDateToDateDefaultValue');
            let endDateLocal = localStorage.getItem("endDate");

            if (endDateLocal) {
                // console.log('start doc tu local storage');
                document.getElementById("endDate").value = endDateLocal;
                // console.log('end doc tu local storage');
            }
        };
    }).catch(err => {
        console.error("Đã có lỗi xảy ra.");
        setLoading(false); // Kết thúc loading
    });
});
