'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const worksheets = dashboard.worksheets;
        const parameterName = "P_Fd_Td"; // Replace with your parameter's name
        const defaultValue = "1000-01-01,1000-01-01"; // Replace with your default value

        // init();

        document.getElementById("refreshButton").addEventListener("click", () => {
            updateAndRefreshData();
        });

        tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(function (parameters) {
            parameters.forEach(function (p) {
                p.addEventListener(tableau.TableauEventType.ParameterChanged, onParameterChange);
            });
        });

        function onParameterChange(parameterChangeEvent) {
            parameterChangeEvent.getParameterAsync().then(function (param) {
                console.log("Change parameter");
            });
        }

        function updateAndRefreshData() {
            setCursorLoading(true); // Hiển thị con trỏ loading
            setLoading(true); // Bắt đầu loading
            const startDate = document.getElementById('startDate').value;
            const endDate = document.getElementById('endDate').value;
            // Chuyển giá trị sang định dạng yyyy-mm-dd
            const formattedStartDate = new Date(startDate).toISOString().split('T')[0];
            const formattedEndDate = new Date(endDate).toISOString().split('T')[0];
            // Nối hai giá trị với dấu phẩy
            const fromDateToDate = `${formattedStartDate},${formattedEndDate}`;

            // Get the parameters P_Fd_Td
            dashboard.getParametersAsync().then(parameters => {
                const fdtdParam = parameters.find(param => param.name === 'P_Fd_Td');
                // Cập nhật P_FD bằng giá trị của fromDateToDate
                fdtdParam.changeValueAsync(fromDateToDate).then(() => {
                    console.log('thay doi tham so P_Fd_Td');
                    setCursorLoading(false); // Tắt con trỏ loading sau khi hoàn thành
                    setLoading(false); // Kết thúc loading
                });
            }).catch(err => {
                setCursorLoading(false); // Tắt con trỏ loading nếu không tìm thấy tham số
                setLoading(false); // Kết thúc loading
                console.error("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_From_Date (date), P_To_Date (Date), P_Fd_Td (String).");
                alert("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_From_Date (date), P_To_Date (Date), P_Fd_Td (String)");
                // console.error("P_Fd_Td = " + fromDateToDate);
                // alert("P_Fd_Td = " + fromDateToDate);
                console.error(err);
            });
        }

        function init() {
            // Get the parameters P_FROM_DATE and P_TO_DATE
            dashboard.getParametersAsync().then(parameters => {
                const fromDateToDate = '1000-01-01,1000-01-01';
                const fdtdParam = parameters.find(param => param.name === 'P_Fd_Td');

                const fdtdDefaultValue = fdtdParam.currentValue.value;
                const [startDateDefault, endDateDefault] = fdtdDefaultValue.split(",");

                // Gán giá trị cho các input type="date"
                document.getElementById("startDate").value = startDateDefault;
                document.getElementById("endDate").value = endDateDefault;

                // Cập nhật P_FD bằng giá trị của P_From_Date + ',' + P_To_Date
                fdtdParam.changeValueAsync(fromDateToDate).then(() => {
                    console.log('thay doi tham so P_Fd_Td');
                });
            }).catch(err => {
                console.error("Đã có lỗi xảy ra. Khi init dữ liệu");
            });
        }

        // Register the unload event to reset the parameter
        window.addEventListener("beforeunload", () => {
            const dashboard = tableau.extensions.dashboardContent.dashboard;

            // Find and set the parameter
            dashboard.findParameterAsync(parameterName).then((parameter) => {
                if (parameter) {
                    parameter.changeValueAsync(defaultValue).then(() => {
                        console.log(`Parameter ${parameterName} reset to default: ${defaultValue}`);
                    });
                } else {
                    console.error(`Parameter ${parameterName} not found`);
                }
            });
        });

        function setCursorLoading(isLoading) {
            if (isLoading) {
                document.body.style.cursor = 'wait'; // Thay đổi con trỏ chuột sang "loading"
            } else {
                document.body.style.cursor = 'default'; // Trả lại trạng thái bình thường
            }
        }

        // Thêm trạng thái loading
        function setLoading(isLoading) {
            if (isLoading) {
                refreshButton.classList.add('loading'); // Thêm class loading
                refreshButton.disabled = true; // Vô hiệu hóa nút khi đang loading
            } else {
                refreshButton.classList.remove('loading'); // Bỏ class loading
                refreshButton.disabled = false; // Kích hoạt lại nút
            }
        }
    });
});
