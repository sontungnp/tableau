'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let selectedValue = "Ch?n giá tr?";
        const button = document.getElementById("combo-button");
        const popup = document.getElementById("combo-popup");
        const overlay = document.getElementById("overlay");
        const selectedDisplay = document.getElementById("selected-value");
        
        button.addEventListener("click", function () {
            popup.style.display = "block";
            overlay.style.display = "block";
        });
        
        overlay.addEventListener("click", function () {
            popup.style.display = "none";
            overlay.style.display = "none";
        });
        
        function fetchData() {
            // D? li?u m?u d? test
            let fieldValues = ["Giá tr? 1", "Giá tr? 2", "Giá tr? 3", "Giá tr? 4", "Giá tr? 5"];
            popup.innerHTML = "";
            fieldValues.forEach(value => {
                let item = document.createElement("div");
                item.textContent = value;
                item.addEventListener("click", function () {
                    selectedValue = value;
                    selectedDisplay.textContent = selectedValue;
                    popup.style.display = "none";
                    overlay.style.display = "none";
                });
                popup.appendChild(item);
            });
        }
        
        fetchData();
    });
});
