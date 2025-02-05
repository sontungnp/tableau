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

                nodes[id] = nodes[id] || { id, name: label, children: [], parent: null };
                if (parentId !== null) {
                    nodes[parentId] = nodes[parentId] || { id: parentId, name: "", children: [], parent: null };
                    nodes[parentId].children.push(nodes[id]);
                    nodes[id].parent = nodes[parentId];
                }
            });
            return Object.values(nodes).find(node => !node.parent) || [];
        }

        function renderTree(node, container) {
            if (!node) return;
            let div = document.createElement("div");
            div.classList.add("node");
            div.innerHTML = `<input type='checkbox' data-id='${node.id}'>${node.name}`;
            container.appendChild(div);
            
            let checkbox = div.querySelector("input");
            checkbox.addEventListener("change", function () {
                toggleChildren(node, this.checked);
                toggleParent(node);
                updateSelectedNodes();
            });
            
            if (node.children.length) {
                let childrenContainer = document.createElement("div");
                childrenContainer.classList.add("children");
                container.appendChild(childrenContainer);
                node.children.forEach(child => renderTree(child, childrenContainer));
            }
        }

        function toggleChildren(node, checked) {
            node.children.forEach(child => {
                let checkbox = document.querySelector(`input[data-id='${child.id}']`);
                if (checkbox) {
                    checkbox.checked = checked;
                    toggleChildren(child, checked);
                }
            });
        }

        function toggleParent(node) {
            if (!node.parent) return;
            let parentCheckbox = document.querySelector(`input[data-id='${node.parent.id}']`);
            if (parentCheckbox) {
                let allChecked = node.parent.children.every(child => document.querySelector(`input[data-id='${child.id}']`).checked);
                parentCheckbox.checked = allChecked;
                toggleParent(node.parent);
            }
        }

        function updateSelectedNodes() {
            selectedNodes.clear();
            document.querySelectorAll(".node input:checked").forEach(checkbox => {
                selectedNodes.add(checkbox.getAttribute("data-id"));
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
