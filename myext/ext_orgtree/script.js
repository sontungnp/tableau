'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        let selectedNodes = new Set();
        let popupData = {};
        let treeData = [];
        let selectedData = {};
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        let worksheets = dashboard.worksheets;
        const worksheetName = "OrgCodeSheet"; // Tên worksheet cần lấy
        const filterField = "Departmentcode"; // 🔴 Đổi tên filter nếu cần

        // lấy từ localstorage
        // treeData = JSON.parse(localStorage.getItem("treeData"));
        fetchData();

        selectedData = JSON.parse(localStorage.getItem("selectedData"));

        // if (treeData && treeData.length !== 0) {
        //     console.log('fetch tree data thành công')
        //     removeParentRefs(treeData); // Xóa vòng lặp trước khi truyền
        //     // luu vao localStorage
        //     localStorage.setItem("treeData", JSON.stringify(treeData));
        // } else {
        //     console.log('fetch tree data không thành công')
        // }

        if (!selectedData) {
            // khởi tạo giá trị lần đầu load extension lên
            selectedData = {
                "action": "INIT",
                "selectedIds": [],
                "selectedCodes": "ALL",
                "showIds": ["ALL"],
                "isAll": "ALL",
                "maxLevel": 2
            };
            localStorage.setItem("selectedData", JSON.stringify(selectedData));
            localStorage.setItem("departmentCode", selectedData.selectedCodes);
        }

        document.getElementById("selected-box").value = selectedData.selectedCodes;
        
        document.getElementById("dropdown-toggle").addEventListener("click", () => {
            let popupUrl = window.location.href + "popup.html"; // URL của file popup
            
            popupData = {
                "treeData": treeData,
                "selectedData": selectedData
            };

            tableau.extensions.ui.displayDialogAsync(popupUrl, JSON.stringify(popupData), { width: 600, height: 800 })
                .then((payload) => {
                    
                    // // kiểm tra treeData nếu rỗng thì lấy lại từ localstorage
                    // if (!treeData || treeData.length === 0) {
                    //     treeData = JSON.parse(localStorage.getItem("treeData"));
                    //     console.log('treeData null -> lay lai tu localstorage');
                    // }

                    // // kiểm tra nếu vẫn rỗng thì láy từ localstorage
                    // if (!treeData || treeData.length === 0) {
                    //     console.log('treeData vẫn null -> fetch lại dữ liệu');
                    //     fetchData();
                    //     if (!treeData || treeData.length === 0)  {
                    //         removeParentRefs(treeData); // Xóa vòng lặp trước khi truyền
                    //         // luu vao localStorage
                    //         localStorage.setItem("treeData", JSON.stringify(treeData));
                    //         console.log('fetch lại tree data thành công')
                    //     } else {
                    //         console.log('fetch lại tree data lỗi - vẫn null')
                    //     }
                    // }

                    let receivedValue  = JSON.parse(payload);
                    if (receivedValue.action === 'ok') {
                        selectedData = {
                            "action": "OK",
                            "selectedIds": receivedValue.selectedIds, 
                            "selectedCodes": receivedValue.selectedCodes,
                            "showIds": receivedValue.showIds, 
                            "isAll": receivedValue.isAll,
                            "maxLevel": receivedValue.maxLevel
                        }

                        localStorage.setItem("selectedData", JSON.stringify(selectedData));
                        localStorage.setItem("departmentCode", selectedData.selectedCodes);

                        document.getElementById("selected-box").value = selectedData.selectedCodes;

                        setFilterOrgCodeByDepartmentCode(selectedData.selectedCodes, selectedData.isAll);
                    } else {
                        console.log("Calcel");
                    }
                })
                .catch((error) => {
                    console.log("Lỗi khi mở popup: " + error.message);
                });
        });

        document.getElementById("clear").addEventListener("click", clearOrgFilters);

        // check thay đổi lcalstorage do nut reset từ extension khác
        // window.addEventListener("storage", function(event) {
        //     if (event.key === "departmentCode") {
        //         console.log("departmentCode đã thay đổi:", event.newValue);
        //         if (event.newValue === null || event.newValue === 'ALL') {
        //             selectedData = {
        //                     "action": "INIT",
        //                     "selectedIds": [],
        //                     "selectedCodes": "ALL",
        //                     "showIds": ["ALL"],
        //                     "isAll": "ALL",
        //                     "maxLevel": 2
        //                 }
        //             localStorage.setItem("selectedData", JSON.stringify(selectedData));
        //             localStorage.setItem("departmentCode", selectedData.selectedCodes);
        //         } else {
        //             selectedData.selectedCodes = event.newValue
        //         }
                
        //         document.getElementById("selected-box").value = selectedData.selectedCodes
        //     }
        // });

        function fetchData() {
            // const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
            const worksheet = worksheets.find(ws => ws.name === worksheetName);
            worksheet.getSummaryDataAsync().then(data => {
                treeData = transformDataToTree(data);
            });
        }

        function transformDataToTree(data) {
            if (!data.data.length) return null; // Nếu dữ liệu rỗng, trả về null
        
            const nodes = {};
            let rootId = data.data[0][0].value; // Lấy ID của dòng đầu tiên làm root
        
            data.data.forEach(row => {
                const id = row[0].value;
                const parentId = row[1].value;
                const label = row[2].value;
                const code = row[3].value; // Đọc thêm cột code
        
                if (!nodes[id]) {
                    nodes[id] = { id, name: label, code, children: [] };
                } else {
                    nodes[id].name = label;
                    nodes[id].code = code; // Gán giá trị code nếu node đã tồn tại
                }
        
                if (parentId !== null) {
                    if (!nodes[parentId]) {
                        nodes[parentId] = { id: parentId, name: "", code: "", children: [] };
                    }
                    nodes[parentId].children.push(nodes[id]);
                }
            });
        
            return nodes[rootId] || null; // Trả về node gốc đã chọn
        }

        async function setFilterOrgCodeByDepartmentCode(lstDepartmentCode, isAll) {
            try {
                // Chuyển filterValue về chuỗi hoặc giá trị mặc định
                let filterStr = (lstDepartmentCode !== null && lstDepartmentCode !== undefined) ? String(lstDepartmentCode).toUpperCase() : "ALL";

                await Promise.allSettled(worksheets.map(async (ws) => {
                    // 🔹 Lấy danh sách filters hiện có trên worksheet
                    let filters = await ws.getFiltersAsync();

                    // Tìm xem worksheet có filter này không -> nếu không có thì bỏ qua
                    if (!filters.some(f => f.fieldName === filterField)) {
                        console.warn(`Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`);
                        return;
                    }

                    if (!lstDepartmentCode || lstDepartmentCode === "ALL" || lstDepartmentCode.trim() === "" || isAll === "ALL") {
                        // 🔹 Nếu filterValue rỗng hoặc là "ALL" => Clear filter
                        document.getElementById("selected-box").value = 'ALL';
                        await ws.clearFilterAsync(filterField);
                    } else {
                        // 🔹 Kiểm tra nếu filterValue là một mảng thì truyền mảng, nếu không thì truyền giá trị đơn lẻ
                        await ws.applyFilterAsync(filterField, lstDepartmentCode.split(",").map(item => item.trim()), "replace");
                    }
                }));

                // alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
            } catch (error) {
                console.error("Error setting filter:", error);
                alert("Failed to set filter. Check console for details.");
            }
        }

        async function clearOrgFilters() {
            // thiết lập giá trị khởi tạo ban đầu
            selectedData = {
                "action": "INIT",
                "selectedIds": [],
                "selectedCodes": "ALL",
                "showIds": ["ALL"],
                "isAll": "ALL",
                "maxLevel": 2
            };

            localStorage.setItem("selectedData", JSON.stringify(selectedData));
            localStorage.setItem("departmentCode", selectedData.selectedCodes);

            document.getElementById("selected-box").value = 'ALL';

            try {
                for (const ws of worksheets) {
                    // 🔹 Lấy danh sách filters hiện có trên worksheet
                    let filters = await ws.getFiltersAsync();
                    
                    // Tìm xem worksheet có filter này không -> nếu ko có thì continue sang worksheet khác
                    let hasFilter = filters.some(f => f.fieldName === filterField);
        
                    if (!hasFilter) {
                        console.warn(`Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`);
                        continue;
                    } else {
                        await ws.clearFilterAsync(filterField);
                    }
                }
        
                // alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
            } catch (error) {
                console.error("Error clear filter:" + filterField, error);
            }
        }

        function removeParentRefs(node) {
            if (!node) return;
            node.children.forEach(child => removeParentRefs(child));
            delete node.parent; // ❌ Xóa thuộc tính parent
        }

    });
});
