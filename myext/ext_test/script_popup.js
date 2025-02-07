'use strict';

tableau.extensions.initializeDialogAsync().then(async (payload) => { // Sử dụng async ở đây
    console.log("Popup mở thành công!");

    console.log(payload);

    document.getElementById("closePopup").addEventListener("click", () => {
        tableau.extensions.ui.closeDialog("Dữ liệu trả về từ popup");
    });

    let treeData = JSON.parse(payload);

    renderTree(treeData, document.getElementById("tree-container"));
    let container = document.getElementById("tree-container");
    container.style.display = container.style.display === "block" ? "none" : "block";
    

    function renderTree(node, container, parent = null) {
        if (!node) return;
        node.parent = parent; // Gán parent cho node
    
        let div = document.createElement("div");
        div.classList.add("node");
    
        let toggle = document.createElement("span");
        toggle.classList.add("toggle");
        toggle.textContent = node.children.length ? "▶" : "";
    
        let checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.dataset.id = node.id;
        checkbox.addEventListener("change", function () {
            toggleChildren(node, this.checked);
            updateParentState(node.parent);
        });
    
        div.appendChild(toggle);
        div.appendChild(checkbox);
        div.appendChild(document.createTextNode(node.name));
        container.appendChild(div);
    
        if (node.children.length) {
            let childrenContainer = document.createElement("div");
            childrenContainer.classList.add("children");
    
            // Nếu là cấp 2 (con của root), thì hiển thị mặc định
            if (parent !== null && parent.parent === null) {
                childrenContainer.style.display = "block"; 
                toggle.textContent = "▼";  // Cập nhật icon toggle
            }
    
            container.appendChild(childrenContainer);
            node.children.forEach(child => renderTree(child, childrenContainer, node));
        }
    
        // Toggle khi click
        toggle.addEventListener("click", function (event) {
            event.stopPropagation();
            let isExpanded = this.textContent === "▼";
            let childrenContainer = div.nextElementSibling;
            if (childrenContainer) {
                childrenContainer.style.display = isExpanded ? "none" : "block";
                this.textContent = isExpanded ? "▶" : "▼";
            }
        });
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
});
