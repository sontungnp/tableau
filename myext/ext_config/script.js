'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        const dashboard = tableau.extensions.dashboardContent.dashboard;

        document.getElementById("refreshButton").addEventListener("click", () => {
            updateAndRefreshData();
        });

        function updateAndRefreshData() {
            const endDate = document.getElementById('endDate').value;
            // Chuyển giá trị sang định dạng yyyy-mm-dd
            const formattedEndDate = new Date(endDate).toISOString().split('T')[0];
            
            // Get the parameters P_Fd_Td
            dashboard.getParametersAsync().then(parameters => {
                const fdtdParam = parameters.find(param => param.name === 'P_Fd_Td');
                // Cập nhật P_FD bằng giá trị của formattedEndDate
                fdtdParam.changeValueAsync(formattedEndDate).then(() => {
                    console.log('thay doi tham so P_Fd_Td');
                });
            }).catch(err => {
                console.error("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String).");
                alert("Đã có lỗi xảy ra. Đảm bảo có đủ tham số: P_Fd_Td (String)");
                // console.error("P_Fd_Td = " + formattedEndDate);
                // alert("P_Fd_Td = " + formattedEndDate);
                console.error(err);
            });
        }
    });
});
