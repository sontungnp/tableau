'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let selectedValue = "Chọn giá trị";
        const button = document.getElementById("combo-button");
        const selectedDisplay = document.getElementById("selected-value");
        
        // Gắn popup vào document của trang cha (Tableau report)
        const popup = parent.document.createElement("div");
        popup.id = "combo-popup";
        popup.className = "combo-popup";
        parent.document.body.appendChild(popup);
        
        button.addEventListener("click", function (event) {
            const rect = button.getBoundingClientRect();
            popup.style.top = `${rect.bottom + parent.window.scrollY}px`;
            popup.style.left = `${rect.left + parent.window.scrollX}px`;
            popup.style.display = "block";
        });
        
        parent.document.addEventListener("click", function (event) {
            if (!button.contains(event.target) && !popup.contains(event.target)) {
                popup.style.display = "none";
            }
        });
        
        function fetchData() {
            let fieldValues = ["Giá trị 1", "Giá trị 2", "Giá trị 3", "Giá trị 4", "Giá trị 5"];
            popup.innerHTML = "";
            fieldValues.forEach(value => {
                let item = parent.document.createElement("div");
                item.textContent = value;
                item.addEventListener("click", function () {
                    selectedValue = value;
                    selectedDisplay.textContent = selectedValue;
                    popup.style.display = "none";
                });
                popup.appendChild(item);
            });
        }
        
        fetchData();
    });
});
