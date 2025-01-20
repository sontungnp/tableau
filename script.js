'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        const worksheets = dashboard.worksheets;

        document.getElementById("refreshButton").addEventListener("click", () => {
            updateAndRefreshData();
        });

        tableau.extensions.dashboardContent.dashboard.getParametersAsync().then(function (parameters) {
            parameters.forEach(function (p) {
              p.addEventListener(tableau.TableauEventType.ParameterChanged, onParameterChange);
            });
          });

          function onParameterChange (parameterChangeEvent) {
            parameterChangeEvent.getParameterAsync().then(function (param) {
                console.log("Change parameter");
            });
          }

        

        function updateAndRefreshData() {
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
                });
            }).catch(err => {
                console.error("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_From_Date (date), P_To_Date (Date), P_Fd_Td (String).");
                alert("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_From_Date (date), P_To_Date (Date), P_Fd_Td (String)");
                // console.error("P_Fd_Td = " + fromDateToDate);
                // alert("P_Fd_Td = " + fromDateToDate);
                console.error(err);
            });
        }
    });
});
