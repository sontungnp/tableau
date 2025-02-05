'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let selectedNode = null;

        document.addEventListener("DOMContentLoaded", function () {
            tableau.extensions.initializeAsync().then(() => {
                fetchData();
            });
        });

        function fetchData() {
            const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            worksheet.getSummaryDataAsync().then(data => {
                const treeData = transformDataToTree(data);
                renderTree(treeData);
            });
        }

        function transformDataToTree(data) {
            const nodes = {};
            data.data.forEach(row => {
                const id = row[0].value;
                const parentId = row[1].value;
                const label = row[2].value;

                nodes[id] = nodes[id] || { id, name: label, children: [] };
                if (parentId !== null) {
                    nodes[parentId] = nodes[parentId] || { id: parentId, name: "", children: [] };
                    nodes[parentId].children.push(nodes[id]);
                }
            });
            return Object.values(nodes).find(node => !node.parentId) || [];
        }

        function renderTree(treeData) {
            const container = document.getElementById("tree-container");
            container.innerHTML = createTreeHtml(treeData);

            document.querySelectorAll(".node").forEach(node => {
                node.addEventListener("click", function () {
                    selectedNode = this.getAttribute("data-name");
                    document.getElementById("apply-filter").disabled = false;
                });
            });
        }

        function createTreeHtml(node) {
            if (!node) return "";
            let html = `<div class='node' data-name='${node.name}'>${node.name}</div>`;
            if (node.children.length) {
                html += `<div style='margin-left:20px;'>`;
                node.children.forEach(child => {
                    html += createTreeHtml(child);
                });
                html += `</div>`;
            }
            return html;
        }

        document.getElementById("apply-filter").addEventListener("click", function () {
            if (!selectedNode) return;
            const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            worksheet.applyFilterAsync("Category", selectedNode, tableau.FilterUpdateType.Replace);
        });
    });
});
