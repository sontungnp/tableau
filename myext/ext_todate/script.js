'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        const dashboard = tableau.extensions.dashboardContent.dashboard;

        document.getElementById("refreshButton").addEventListener("click", () => {
            updateAndRefreshData();
        });

        function updateAndRefreshData() {
            setLoading(true); // Bắt đầu loading
            const endDate = document.getElementById('endDate').value;
            // Chuyển giá trị sang định dạng yyyy-mm-dd
            const formattedEndDate = new Date(endDate).toISOString().split('T')[0];
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
    });
});
