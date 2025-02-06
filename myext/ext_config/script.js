'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        let selectedNodes = new Set();
        let treeData = [];

        document.getElementById("dropdown-toggle").addEventListener("click", function () {
            // Mở cửa sổ mới khi click vào combo box
            let popupWindow = window.open('', '', 'width=800,height=600');
            popupWindow.document.write(`
                <html>
                    <head>
                        <title>Tableau Tree Filter</title>
                        <style>
                            body {
                                font-family: Arial, sans-serif;
                                padding: 20px;
                            }
                            .node {
                                display: flex;
                                align-items: center;
                                gap: 10px;
                                padding: 2px 5px;
                                cursor: pointer;
                            }
                            .toggle {
                                cursor: pointer;
                                width: 18px;
                                text-align: center;
                            }
                            .checkbox {
                                margin-right: 10px;
                            }
                            .children {
                                margin-left: 20px;
                                display: none;
                            }
                            .expanded > .children {
                                display: block;
                            }
                        </style>
                    </head>
                    <body>
                        <h2>Tree Filter</h2>
                        <div id="tree-container"></div>
                        <script src="https://cdnjs.cloudflare.com/ajax/libs/tableau-api/2.3.0/tableau.min.js"></script>
                        <script>
                            // Nhúng code JavaScript từ script.js vào popup
                            ${fetchData.toString()}
                            ${transformDataToTree.toString()}
                            ${renderTree.toString()}

                            fetchData(); // Gọi hàm để tải dữ liệu và render tree
                        </script>
                    </body>
                </html>
            `);
        });

        fetchData();

        function fetchData() {
            const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            worksheet.getSummaryDataAsync().then(data => {
                treeData = transformDataToTree(data);
                renderTree(treeData, popupWindow.document.getElementById("tree-container"));
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

            let toggle = document.createElement("span");
            toggle.classList.add("toggle");
            toggle.textContent = node.children.length ? "▶" : "";
            toggle.addEventListener("click", function (event) {
                event.stopPropagation();
                let parent = this.parentElement;
                let childrenContainer = parent.nextElementSibling;
                if (childrenContainer) {
                    let isExpanded = childrenContainer.style.display === "block";
                    childrenContainer.style.display = isExpanded ? "none" : "block";
                    this.textContent = isExpanded ? "▶" : "▼";
                }
            });

            let checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.classList.add("checkbox");
            checkbox.dataset.id = node.id;
            checkbox.addEventListener("change", function () {
                toggleChildren(node, this.checked);
                updateParentState(node.parent);
                updateSelectedNodes();
            });

            div.appendChild(toggle);
            div.appendChild(checkbox);
            div.appendChild(document.createTextNode(node.name));
            container.appendChild(div);

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
                    checkbox.indeterminate = false;
                    toggleChildren(child, checked);
                }
            });
        }

        function updateParentState(node) {
            if (!node) return;
            let parentCheckbox = document.querySelector(`input[data-id='${node.id}']`);
            let childCheckboxes = node.children.map(child => document.querySelector(`input[data-id='${child.id}']`));

            let allChecked = childCheckboxes.every(checkbox => checkbox.checked);
            let someChecked = childCheckboxes.some(checkbox => checkbox.checked || checkbox.indeterminate);

            parentCheckbox.checked = allChecked;
            parentCheckbox.indeterminate = !allChecked && someChecked;

            updateParentState(node.parent);
        }
    });
});
