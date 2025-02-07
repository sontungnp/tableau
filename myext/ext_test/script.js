'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        document.getElementById("dropdown-toggle").addEventListener("click", () => {
            let popupUrl = window.location.origin + "/tableau/myext/ext_test/popup.html"; // URL của file popup
            tableau.extensions.ui.displayDialogAsync(popupUrl, "", { width: 400, height: 300 })
                .then((payload) => {
                    console.log("Popup đóng với dữ liệu: " + payload);
                })
                .catch((error) => {
                    console.log("Lỗi khi mở popup: " + error.message);
                });
        });

        tableau.extensions.initializeDialogAsync().then(() => {
            console.log("Popup mở thành công!");

            document.getElementById("closePopup").addEventListener("click", () => {
                tableau.extensions.ui.closeDialog("Dữ liệu trả về từ popup");
            });
        });
    });
});
