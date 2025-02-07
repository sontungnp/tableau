'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        document.getElementById("openPopup").addEventListener("click", () => {
            let popupUrl = window.location.origin + "/tableau/myext/ext_config/popup.html"; // URL của file popup
            tableau.extensions.ui.displayDialogAsync(popupUrl, "", { width: 400, height: 300 })
                .then((payload) => {
                    console.log("Popup đóng với dữ liệu: " + payload);
                })
                .catch((error) => {
                    console.log("Lỗi khi mở popup: " + error.message);
                });
        });
    });
});
