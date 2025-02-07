'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        document.getElementById("resizeButton").addEventListener("click", () => {
            let newWidth = 600; // Đổi sang kích thước mong muốn
            let newHeight = 400;
            
            tableau.extensions.ui.dashboardContent.setFrameSize(newWidth, newHeight);
            console.log(`Đã đổi kích thước: ${newWidth}x${newHeight}`);
        });
    });
});
