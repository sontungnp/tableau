'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let selectedNode = null;

        document.addEventListener("DOMContentLoaded", function () {
            tableau.extensions.initializeAsync().then(() => {
                fetchData();
            });
        });

        fetchData();

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

            document.querySelectorAll(".node input").forEach(checkbox => {
                checkbox.addEventListener("change", function () {
                    if (this.checked) {
                        selectedNodes.add(this.getAttribute("data-name"));
                    } else {
                        selectedNodes.delete(this.getAttribute("data-name"));
                    }
                    document.getElementById("apply-filter").disabled = selectedNodes.size === 0;
                });
            });
        }

        function createTreeHtml(node) {
            if (!node) return "";
            let html = `<div class='node'><input type='checkbox' data-name='${node.name}'>${node.name}</div>`;
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
            if (selectedNodes.size === 0) return;
            const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            worksheet.applyFilterAsync("Category", Array.from(selectedNodes), tableau.FilterUpdateType.Replace);
        });
    });
});
