'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        let selectedNodes = new Set();
        let popupData = {};
        let treeData = [];

        // khởi tạo giá trị lần đầu load extension lên
        let selectedData = {
            "action": "INIT",
            "selectedLeafIds": [],
            "showIds": ["ALL"],
            "isAll": "ALL",
            "maxLevel": 2
        };
        fetchData();

        function fetchData() {
            const worksheetName = "OrgCodeSheet"; // Tên worksheet cần lấy
            const dashboard = tableau.extensions.dashboardContent.dashboard;
            // const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            const worksheet = dashboard.worksheets.find(ws => ws.name === worksheetName);
            worksheet.getSummaryDataAsync().then(data => {
                treeData = transformDataToTree(data);
            });
        }

        // function transformDataToTree(data) {
        //     const nodes = {};
        //     data.data.forEach(row => {
        //         const id = row[0].value;
        //         const parentId = row[1].value;
        //         const label = row[2].value;
        //         nodes[id] = nodes[id] || { id, name: label, children: [], parent: null };
        //         if (parentId !== null) {
        //             nodes[parentId] = nodes[parentId] || { id: parentId, name: "", children: [], parent: null };
        //             nodes[parentId].children.push(nodes[id]);
        //             nodes[id].parent = nodes[parentId];
        //         }
        //     });
        //     return Object.values(nodes).find(node => !node.parent) || [];
        // }

        function transformDataToTree(data) {
            if (!data.data.length) return null; // Nếu dữ liệu rỗng, trả về null
        
            const nodes = {};
            let rootId = data.data[0][0].value; // Lấy ID của dòng đầu tiên làm root
        
            data.data.forEach(row => {
                const id = row[0].value;
                const parentId = row[1].value;
                const label = row[2].value;
        
                if (!nodes[id]) {
                    nodes[id] = { id, name: label, children: [] };
                } else {
                    nodes[id].name = label;
                }
        
                if (parentId !== null) {
                    if (!nodes[parentId]) {
                        nodes[parentId] = { id: parentId, name: "", children: [] };
                    }
                    nodes[parentId].children.push(nodes[id]);
                }
            });
        
            return nodes[rootId] || null; // Trả về node gốc đã chọn
        }
        
        
        document.getElementById("dropdown-toggle").addEventListener("click", () => {
            let popupUrl = window.location.origin + "/tableau/myext/ext_test/popup.html"; // URL của file popup

            function removeParentRefs(node) {
                if (!node) return;
                node.children.forEach(child => removeParentRefs(child));
                delete node.parent; // ❌ Xóa thuộc tính parent
            }
        
            removeParentRefs(treeData); // Xóa vòng lặp trước khi truyền
            
            popupData = {
                "treeData": treeData,
                "selectedData": selectedData
            };

            tableau.extensions.ui.displayDialogAsync(popupUrl, JSON.stringify(popupData), { width: 600, height: 800 }) // JSON.stringify(treeData)
                .then((payload) => {
                    console.log("Popup đóng với dữ liệu: " + payload);
                    let receivedValue  = JSON.parse(payload);
                    if (receivedValue.action === 'ok') {
                        console.log("Ok");
                        selectedData = {
                            "selectedLeafIds": receivedValue.selectedLeafIds, 
                            "showIds": receivedValue.showIds, 
                            "isAll": receivedValue.isAll,
                            "maxLevel": receivedValue.maxLevel
                        }

                        document.getElementById("search-box").value = arrayToString(selectedData.showIds);

                        setFilterOrgCode(selectedData.selectedLeafIds, selectedData.isAll);
                    } else {
                        console.log("Calcel");
                    }
                })
                .catch((error) => {
                    console.log("Lỗi khi mở popup: " + error.message);
                });
        });

        function arrayToString(arr) {
            return arr.join(",");
        }

        async function setFilterOrgCode(filterValue, isAll) {
            try {
                const dashboard = tableau.extensions.dashboardContent.dashboard;
                const filterField = "Orgid"; // 🔴 Đổi tên filter nếu cần
        
                let worksheets = dashboard.worksheets;
        
                // Chuyển filterValue về chuỗi hoặc giá trị mặc định
                let filterStr = (filterValue !== null && filterValue !== undefined) ? String(filterValue).toUpperCase() : "ALL";
        
                for (const ws of worksheets) {
                    // 🔹 Lấy danh sách filters hiện có trên worksheet
                    let filters = await ws.getFiltersAsync();
                    
                    // Tìm xem worksheet có filter này không -> nếu ko có thì continue sang worksheet khác
                    let hasFilter = filters.some(f => f.fieldName === filterField);
        
                    if (!hasFilter) {
                        console.warn(`Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`);
                        continue;
                    }
        
                    if (!filterValue || filterStr === "ALL" || filterStr.trim() === "" || isAll === "ALL") {
                        // 🔹 Nếu filterValue rỗng hoặc là "ALL" => Clear filter
                        await ws.clearFilterAsync(filterField);
                    } else {
                        // 🔹 Kiểm tra nếu filterValue là một mảng thì truyền mảng, nếu không thì truyền giá trị đơn lẻ
                        await ws.applyFilterAsync(filterField, filterValue, "replace");
                    }
                }
        
                alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
            } catch (error) {
                console.error("Error setting filter:", error);
                alert("Failed to set filter. Check console for details.");
            }
        }
    });
});
