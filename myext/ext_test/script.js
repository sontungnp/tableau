'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        document.getElementById('openPopupButton').addEventListener('click', function() {
            document.getElementById('popup').style.display = 'block';
        });
        
        function closePopup() {
            document.getElementById('popup').style.display = 'none';
        }
    });
});
