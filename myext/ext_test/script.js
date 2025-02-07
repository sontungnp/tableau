'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        let selectedNodes = new Set();
        let treeData = [];
        fetchData();

        function fetchData() {
            const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            worksheet.getSummaryDataAsync().then(data => {
                treeData = transformDataToTree(data);
            });
        }

        function transformDataToTree(data) {
            const nodes = {};
            data.data.forEach(row => {
                const id = row[0].value;
                const parentId = row[1].value;
                const label = row[2].value;
                nodes[id] = nodes[id] || { id, name: label, children: [], parent: null };
                if (parentId !== null) {
                    nodes[parentId] = nodes[parentId] || { id: parentId, name: "", children: [], parent: null };
                    nodes[parentId].children.push(nodes[id]);
                    nodes[id].parent = nodes[parentId];
                }
            });
            return Object.values(nodes).find(node => !node.parent) || [];
        }
        
        document.getElementById("dropdown-toggle").addEventListener("click", () => {
            let popupUrl = window.location.origin + "/tableau/myext/ext_test/popup.html"; // URL của file popup
            tableau.extensions.ui.displayDialogAsync(popupUrl, JSON.stringify(treeData), { width: 400, height: 300 })
                .then((payload) => {
                    console.log("Popup đóng với dữ liệu: " + payload);
                })
                .catch((error) => {
                    console.log("Lỗi khi mở popup: " + error.message);
                });
        });
    });
});
