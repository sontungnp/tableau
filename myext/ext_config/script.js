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
                renderTree(treeData, document.getElementById("tree-container"));
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

        function renderTree(node, container) {
            if (!node) return;
            let div = document.createElement("div");
            div.classList.add("node");
            div.innerHTML = `<input type='checkbox' data-name='${node.name}'>${node.name}`;
            container.appendChild(div);
            
            let checkbox = div.querySelector("input");
            checkbox.addEventListener("change", function () {
                toggleChildren(checkbox, node.children, this.checked);
                updateSelectedNodes();
            });
            
            if (node.children.length) {
                let childrenContainer = document.createElement("div");
                childrenContainer.classList.add("children");
                container.appendChild(childrenContainer);
                node.children.forEach(child => renderTree(child, childrenContainer));
            }
        }

        function toggleChildren(parentCheckbox, children, checked) {
            children.forEach(child => {
                let checkbox = document.querySelector(`input[data-name='${child.name}']`);
                if (checkbox) {
                    checkbox.checked = checked;
                    toggleChildren(checkbox, child.children, checked);
                }
            });
        }

        function updateSelectedNodes() {
            selectedNodes.clear();
            document.querySelectorAll(".node input:checked").forEach(checkbox => {
                selectedNodes.add(checkbox.getAttribute("data-name"));
            });
            document.getElementById("apply-filter").disabled = selectedNodes.size === 0;
        }

        document.getElementById("apply-filter").addEventListener("click", function () {
            if (selectedNodes.size === 0) return;
            const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            worksheet.applyFilterAsync("Category", Array.from(selectedNodes), tableau.FilterUpdateType.Replace);
        });
    });
});
