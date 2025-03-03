'use strict';

document.addEventListener("DOMContentLoaded", () => {
    tableau.extensions.initializeAsync().then(() => {
        console.log("Extension initialized");

        let selectedNodes = new Set();
        let popupData = {};
        let treeData = [];
        const dashboard = tableau.extensions.dashboardContent.dashboard;
        let worksheets = dashboard.worksheets;
        const worksheetName = "OrgCodeSheet"; // Tên worksheet cần lấy
        // const filterField = "Orgid"; // 🔴 Đổi tên filter nếu cần
        const filterField = "filter_reset_Departmentcode"; // 🔴 Đổi tên filter nếu cần

        // addEventListenerFilter();

        // khởi tạo giá trị lần đầu load extension lên
        let selectedData = {
            "action": "INIT",
            "selectedIds": [],
            "selectedCodes": "ALL",
            "showIds": ["ALL"],
            "isAll": "ALL",
            "maxLevel": 2
        };
        
        // lấy từ localstorage
        selectedData.selectedCodes = localStorage.getItem("departmentCode");
        if (selectedData.selectedCodes === null || selectedData.selectedCodes.trim() === "") {
            selectedData.selectedCodes = 'ALL'
        }
        document.getElementById("selected-box").value = selectedData.selectedCodes;
        
        fetchData();

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
        
        
        document.getElementById("dropdown-toggle").addEventListener("click", () => {
            // Sử dụng localStorage để gửi tín hiệu
            let currentState = localStorage.getItem("popupVisible");
            if (currentState) {
                localStorage.setItem("popupVisible", !currentState);
            } else {
                localStorage.setItem("popupVisible", false);
            }
            
            
            // Gửi sự kiện qua Storage API (Extension2 sẽ nghe sự kiện này)
            window.dispatchEvent(new Event("storage"));

            // let popupUrl = window.location.href + "popup.html"; // URL của file popup
            // // console.log('Vị trí: ', window.location)

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

            // luu vao localstorage
            localStorage.setItem("popupData", JSON.stringify(popupData));

            // tableau.extensions.ui.displayDialogAsync(popupUrl, JSON.stringify(popupData), { width: 600, height: 800 }) // JSON.stringify(treeData)
            //     .then((payload) => {
            //         console.log("Popup đóng với dữ liệu: " + payload);
            //         let receivedValue  = JSON.parse(payload);
            //         if (receivedValue.action === 'ok') {
            //             console.log("Ok");
            //             selectedData = {
            //                 "selectedIds": receivedValue.selectedIds, 
            //                 "selectedCodes": receivedValue.selectedCodes,
            //                 "showIds": receivedValue.showIds, 
            //                 "isAll": receivedValue.isAll,
            //                 "maxLevel": receivedValue.maxLevel
            //             }

            //             document.getElementById("selected-box").value = selectedData.selectedCodes;

            //             // lưu vào localstorage
            //             localStorage.setItem("departmentCode", selectedData.selectedCodes);

            //             // setFilterOrgCode(selectedData.selectedIds, selectedData.isAll);
            //             setFilterOrgCodeByDepartmentCode(selectedData.selectedCodes, selectedData.isAll);
            //         } else {
            //             console.log("Calcel");
            //         }
            //     })
            //     .catch((error) => {
            //         console.log("Lỗi khi mở popup: " + error.message);
            //     });
        });

        function arrayToString(arr) {
            return arr.join(",");
        }

        /*
        async function setFilterOrgCode(filterValue, isAll) {
            try {
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
                        document.getElementById("selected-box").value = 'ALL';
                        await ws.clearFilterAsync(filterField);
                    } else {
                        // 🔹 Kiểm tra nếu filterValue là một mảng thì truyền mảng, nếu không thì truyền giá trị đơn lẻ
                        await ws.applyFilterAsync(filterField, filterValue, "replace");
                    }
                }
        
                // alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
            } catch (error) {
                console.error("Error setting filter:", error);
                alert("Failed to set filter. Check console for details.");
            }
        }
            
        */

        async function setFilterOrgCode(filterValue, isAll) {
            try {
                // Chuyển filterValue về chuỗi hoặc giá trị mặc định
                let filterStr = (filterValue !== null && filterValue !== undefined) ? String(filterValue).toUpperCase() : "ALL";

                await Promise.allSettled(worksheets.map(async (ws) => {
                    // 🔹 Lấy danh sách filters hiện có trên worksheet
                    let filters = await ws.getFiltersAsync();

                    // Tìm xem worksheet có filter này không -> nếu không có thì bỏ qua
                    if (!filters.some(f => f.fieldName === filterField)) {
                        console.warn(`Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`);
                        return;
                    }

                    if (!filterValue || filterStr === "ALL" || filterStr.trim() === "" || isAll === "ALL") {
                        // 🔹 Nếu filterValue rỗng hoặc là "ALL" => Clear filter
                        document.getElementById("selected-box").value = 'ALL';
                        await ws.clearFilterAsync(filterField);
                    } else {
                        // 🔹 Kiểm tra nếu filterValue là một mảng thì truyền mảng, nếu không thì truyền giá trị đơn lẻ
                        await ws.applyFilterAsync(filterField, filterValue, "replace");
                    }
                }));

                // alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
            } catch (error) {
                console.error("Error setting filter:", error);
                alert("Failed to set filter. Check console for details.");
            }
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


        document.getElementById("clear").addEventListener("click", clearOrgFilters);

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

            document.getElementById("selected-box").value = 'ALL';

            // lưu vào localstorage
            localStorage.setItem("departmentCode", 'ALL');

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

        function clearAllFilters() {
            worksheets.forEach((worksheet) => {
                worksheet.getFiltersAsync().then((filters) => {
                    filters.forEach((filter) => {
                        worksheet.clearFilterAsync(filter.fieldName);
                    });
                });
            });
        }

        function filterChangedHandler(event) {
            event.getFilterAsync().then(updatedFilter => {
                if (updatedFilter.fieldName === filterField) {
                    if (updatedFilter.appliedValues.length === 0) {
                        selectedData = {
                            "action": "INIT",
                            "selectedIds": [],
                            "selectedCodes": "ALL",
                            "showIds": ["ALL"],
                            "isAll": "ALL",
                            "maxLevel": 2
                        }

                        // document.getElementById("selected-box").value = 'ALL';

                        setFilterOrgCode(selectedData.selectedIds, selectedData.isAll);
                    }
                    console.log(`filter_reset_Departmentcode đã bị thay đổi sang giá trị: ${updatedFilter.appliedValues.map(v => v.formattedValue).join(", ")}`);
                }
            });
        }

        async function addEventListenerFilter() {
            for (const ws of worksheets) {
                // 🔹 Lấy danh sách filters hiện có trên worksheet
                let filters = await ws.getFiltersAsync();
    
                let hasFilter = filters.some(f => f.fieldName === filterField);
            
                if (!hasFilter) {
                    console.warn(`Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`);
                    continue;
                } else {
                    ws.addEventListener(tableau.TableauEventType.FilterChanged, filterChangedHandler);
                }
            }
        }

        window.addEventListener("storage", function(event) {
            if (event.key === "departmentCode") {
                console.log("departmentCode đã thay đổi:", event.newValue);
                if (event.newValue === null || event.newValue === 'ALL') {
                    selectedData = {
                            "action": "INIT",
                            "selectedIds": [],
                            "selectedCodes": "ALL",
                            "showIds": ["ALL"],
                            "isAll": "ALL",
                            "maxLevel": 2
                        }
                } else {
                    selectedData.selectedCodes = event.newValue
                }
                
                document.getElementById("selected-box").value = selectedData.selectedCodes
            }
        });

    });
});
