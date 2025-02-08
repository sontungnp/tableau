'use strict';

tableau.extensions.initializeDialogAsync().then(async (payload) => { // Sử dụng async ở đây
    let selectedItems = [];
    let expandLevel = 2; // Giá trị này có thể nhận từ tham số truyền vào

    console.log("Popup mở thành công!");

    console.log(payload);

    document.getElementById("closePopup").addEventListener("click", () => {
        tableau.extensions.ui.closeDialog("Dữ liệu trả về từ popup");
        // tableau.extensions.ui.closeDialog(JSON.stringify(selectedItems));
    });

    let treeData = JSON.parse(payload);
    
    renderTree(treeData, document.getElementById("tree-container"), null, 1, expandLevel);
    let container = document.getElementById("tree-container");
    container.style.display = container.style.display === "block" ? "none" : "block";
    

    function renderTree(node, container, parent = null, level = 1, expandLevel = 2) {
        if (!node) return;
        node.parent = parent;

        let div = document.createElement("div");
        div.classList.add("node");

        let toggle = document.createElement("span");
        toggle.classList.add("toggle");
        toggle.textContent = node.children.length ? (level <= expandLevel ? "▼" : "▶") : "";
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
            updateSelectedItems(); // 🔥 CẬP NHẬT DANH SÁCH 🔥
        });

        div.appendChild(toggle);
        div.appendChild(checkbox);
        div.appendChild(document.createTextNode(node.name));
        container.appendChild(div);

        if (node.children.length) {
            let childrenContainer = document.createElement("div");
            childrenContainer.classList.add("children");
            container.appendChild(childrenContainer);

            if (level <= expandLevel) {
                childrenContainer.style.display = "block"; // Mở rộng theo tham số truyền vào
            }

            node.children.forEach(child => renderTree(child, childrenContainer, node, level + 1, expandLevel));
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
                checkbox.indeterminate = false; // Bổ sung để đảm bảo không có trạng thái trung gian
                toggleChildren(child, checked);
            }
        });
    }

    function updateParentState(node) {
        if (!node) return;
        let parentCheckbox = document.querySelector(`input[data-id='${node.id}']`);
        let childCheckboxes = node.children.map(child => document.querySelector(`input[data-id='${child.id}']`)).filter(checkbox => checkbox !== null);
    
        let allChecked = childCheckboxes.length > 0 && childCheckboxes.every(checkbox => checkbox.checked);
        let someChecked = childCheckboxes.some(checkbox => checkbox.checked || checkbox.indeterminate);
    
        parentCheckbox.checked = allChecked;
        parentCheckbox.indeterminate = !allChecked && someChecked;
    
        updateParentState(node.parent);
    }

    function updateSelectedItems() { 
        selectedItems = [];  // 🔥 XÓA DANH SÁCH CŨ 🔥
        document.querySelectorAll("input[type='checkbox']:checked").forEach(checkbox => {
            let id = checkbox.dataset.id;
            let node = findNodeById(treeData, id);
            if (node) {
                let isBranch = node.children.length > 0;
                let isFullySelected = isBranch ? node.children.every(child => document.querySelector(`input[data-id='${child.id}']`).checked) : false;
    
                selectedItems.push({
                    id: node.id,
                    name: node.name,
                    level: getLevel(node),
                    type: isBranch ? "Cành" : "Lá",
                    selection: isBranch ? (isFullySelected ? "Tất cả" : "Một phần") : "N/A"
                });
            }
        });
        renderSelectedItemsTable();  // 🔥 CẬP NHẬT BẢNG 🔥
    }
    
    function renderSelectedItemsTable() { 
        let table = document.getElementById("selected-items-table"); 
        let tbody = table.querySelector("tbody"); 
        tbody.innerHTML = ""; // 🔥 XÓA DỮ LIỆU CŨ 🔥
    
        selectedItems.forEach(item => { 
            let row = document.createElement("tr"); 
            row.innerHTML = `
                <td>${item.id}</td>
                <td>${item.name}</td>
                <td>${item.level}</td>
                <td>${item.type}</td>
                <td>${item.selection}</td>
            `; 
            tbody.appendChild(row); 
        }); 
    }
    
    function findNodeById(node, id) {
        if (!node) return null;
        if (node.id == id) return node;
        for (let child of node.children) {
            let found = findNodeById(child, id);
            if (found) return found;
        }
        return null;
    }

    function getLevel(node) {
        let level = 1;
        while (node.parent) {
            level++;
            node = node.parent;
        }
        return level;
    }
});
