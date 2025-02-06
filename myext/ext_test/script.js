'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let selectedValue = "Chọn giá trị";
        const button = document.getElementById("combo-button");
        const selectedDisplay = document.getElementById("selected-value");
        
        const popup = document.createElement("div");
        popup.id = "combo-popup";
        popup.className = "combo-popup";
        document.body.appendChild(popup);
        
        button.addEventListener("click", function (event) {
            const rect = button.getBoundingClientRect();
            popup.style.top = `${rect.bottom + window.scrollY}px`;
            popup.style.left = `${rect.left + window.scrollX}px`;
            popup.style.display = "block";
        });
        
        document.addEventListener("click", function (event) {
            if (!button.contains(event.target) && !popup.contains(event.target)) {
                popup.style.display = "none";
            }
        });
        
        function fetchData() {
            let fieldValues = ["Giá trị 1", "Giá trị 2", "Giá trị 3", "Giá trị 4", "Giá trị 5"];
            popup.innerHTML = "";
            fieldValues.forEach(value => {
                let item = document.createElement("div");
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
