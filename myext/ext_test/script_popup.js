'use strict';

tableau.extensions.initializeDialogAsync().then(async (payload) => { // Sử dụng async ở đây
    console.log("Popup mở thành công!");

    console.log(payload);

    document.getElementById("closePopup").addEventListener("click", () => {
        tableau.extensions.ui.closeDialog("Dữ liệu trả về từ popup");
    });

    document.getElementById("loadData").addEventListener("click", () => {
        console.log("Load Data");
    });

    let selectedNodes = new Set();
    let treeData = [];

    await fetchData(); // Đợi fetchData chạy xong trước khi tiếp tục

    console.log(treeData);

    renderTree(treeData, document.getElementById("tree-container"));

    async function fetchData() {
        const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
        const data = await worksheet.getSummaryDataAsync(); // Đợi dữ liệu được lấy về
        treeData = transformDataToTree(data);
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

    function filterTree() {
        let query = document.getElementById("search-box").value.toLowerCase();
        document.querySelectorAll(".node").forEach(node => {
            let text = node.textContent.toLowerCase();
            node.style.display = text.includes(query) ? "flex" : "none";
        });
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
