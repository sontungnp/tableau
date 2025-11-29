'use strict'

let selectedCellValue = null
let expandListenersBound = false // <-- thêm dòng này
let extractRefreshTime = ''

let gridApi = null
let nestedData = []

let currentExpandedLevel = 1
let maxTreeLevel = 1

function setAllExpanded(nodes, expanded) {
  if (!nodes || !nodes.length) return
  for (const n of nodes) {
    if (n.children && n.children.length) {
      n.expanded = expanded
      setAllExpanded(n.children, expanded)
    }
  }
}

// Mở rộng toàn bộ subtree của 1 node
function setSubtreeExpanded(node, expanded) {
  if (!node) return
  node.expanded = expanded
  if (node.children) {
    node.children.forEach((child) => setSubtreeExpanded(child, expanded))
  }
}

// Tìm node theo ID trong nestedData
function findNodeById(nodes, id) {
  for (const n of nodes) {
    if (n.id == id) return n
    if (n.children) {
      const f = findNodeById(n.children, id)
      if (f) return f
    }
  }
  return null
}

function getMaxTreeLevel(nodes) {
  let max = 1

  function walk(list) {
    for (const n of list) {
      max = Math.max(max, n.level)
      if (n.children) walk(n.children)
    }
  }

  walk(nodes)
  return max
}

function applyExpandLevel(nodes, level) {
  for (const n of nodes) {
    n.expanded = n.level < level // mở tất cả level < currentExpandedLevel

    if (n.children) {
      applyExpandLevel(n.children, level)
    }
  }
}

// Hàm chuẩn hóa chỉ để đồng bộ Unicode, không bỏ dấu
function normalizeUnicode(str) {
  return str ? str.normalize('NFC').toLowerCase().trim() : ''
}

// Pivot Measure Names/Values
function pivotMeasureValues(
  table,
  excludeIndexes = [],
  fieldFormat = 'snake_case'
) {
  // 🔹 Hàm chuyển format cho key field
  const formatField = (str) => {
    switch (fieldFormat) {
      case 'camelCase':
        return str
          .replace(/(?:^\w|[A-Z]|\b\w|\s+)/g, (match, index) =>
            index === 0 ? match.toLowerCase() : match.toUpperCase()
          )
          .replace(/\s+/g, '')
      case 'snake_case':
        return str.replace(/\s+/g, '_')
      default:
        return str // Giữ nguyên
    }
  }

  const cols = table.columns.map((c) => c.fieldName)
  const rows = table.data.map((r) =>
    r.map((c) => {
      if (c.nativeValue === null || c.nativeValue === undefined) return ''

      // 🔹 Nếu là kiểu ngày hợp lệ (Date object hoặc chuỗi ngày)
      if (c.nativeValue instanceof Date) {
        // Định dạng dd/MM/yyyy có thêm số 0
        return c.nativeValue.toLocaleDateString('vi-VN', {
          day: '2-digit',
          month: '2-digit',
          year: 'numeric'
        })
      }

      return c.formattedValue
    })
  )

  // 🔹 Xác định vị trí Measure Names / Values
  const measureNameIdx = cols.findIndex((c) =>
    c.toLowerCase().includes('measure names')
  )
  const measureValueIdx = cols.findIndex((c) =>
    c.toLowerCase().includes('measure values')
  )

  const dimensionIdxs = cols
    .map((c, i) => i)
    .filter((i) => i !== measureNameIdx && i !== measureValueIdx)

  // 🔹 Pivot dữ liệu
  const pivotMap = new Map()
  const measureSet = new Set()

  rows.forEach((r) => {
    const dimKey = dimensionIdxs.map((i) => r[i]).join('||')
    const mName = r[measureNameIdx]
    const mValue = r[measureValueIdx]

    measureSet.add(mName)

    if (!pivotMap.has(dimKey)) {
      pivotMap.set(dimKey, {
        dims: dimensionIdxs.map((i) => r[i]),
        measures: {}
      })
    }
    pivotMap.get(dimKey).measures[mName] = mValue
  })

  // console.log('pivotMap', JSON.stringify(Object.fromEntries(pivotMap), null, 2))

  const measureNames = Array.from(measureSet)
  const headers = [...dimensionIdxs.map((i) => cols[i]), ...measureNames]
  const isMeasure = [
    ...dimensionIdxs.map(() => false),
    ...measureNames.map(() => true)
  ]

  // 🔹 Loại bỏ các cột có tên bắt đầu bằng "hiden" hoặc "AGG("
  const headerIndexesToKeep = headers
    .map((header, index) => ({ header, index }))
    .filter(({ header }) => {
      const cleanHeader = header.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
      return (
        !cleanHeader.toLowerCase().startsWith('hiden') &&
        !cleanHeader.startsWith('AGG(')
      )
    })
    .map(({ index }) => index)

  const filteredHeaders = headerIndexesToKeep.map((index) => headers[index])
  const filteredIsMeasure = headerIndexesToKeep.map((index) => isMeasure[index])

  // ⚡ Sinh dữ liệu dạng object (key = field format) - chỉ giữ các cột hợp lệ
  const data = Array.from(pivotMap.values()).map((entry) => {
    const row = {}
    filteredHeaders.forEach((h, idx) => {
      const originalIdx = headerIndexesToKeep[idx]
      const cleanHeader = h.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
      const key = formatField(cleanHeader)

      if (originalIdx < dimensionIdxs.length) {
        // Là dimension
        row[key] = entry.dims[originalIdx]
      } else {
        // Là measure
        const mName = measureNames[originalIdx - dimensionIdxs.length]
        const rawValue = entry.measures[mName] || ''
        const numValue = parseFloat(rawValue.toString().replace(/,/g, ''))
        row[key] = !isNaN(numValue) ? numValue : rawValue
      }
    })
    return row
  })

  // ⚡ columnDefs khớp field format, có xử lý width và numericColumn
  let demTree = 0
  const tmpColumnDefs = headers.map((h, idx) => {
    const widthMatch = h.match(/\((\d+)\)/)
    const width = widthMatch ? parseInt(widthMatch[1], 10) : 150 // mặc định 150
    const cleanHeader = h.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
    const fieldName = formatField(cleanHeader)
    // console.log('demTree', demTree)

    if (fieldName.startsWith('tree_lv')) {
      if (demTree === 0) {
        demTree = demTree + 1
        return {
          headerName: 'Cấu trúc cây',
          field: 'name',
          width: 300,
          cellRenderer: (params) => {
            const node = params.data
            if (!node) return ''

            const indent = '<span class="tree-indent"></span>'.repeat(
              node.level - 1
            )
            if (node.leaf) {
              return indent + '' + (node.name || '')
            } else {
              const symbol = node.expanded ? '▾' : '▸'
              return (
                indent +
                // `<span class="toggle-btn" data-id="${node.id}">${symbol}</span> 📁 ` +
                `<span class="toggle-btn" data-id="${node.id}">${symbol}</span> ` +
                node.name
              )
            }
          }
        }
      }
    } else {
      const colDef = {
        field: fieldName,
        headerName: cleanHeader,
        wrapText: true,
        // autoHeight: true,
        width: width,
        minWidth: 30,
        maxWidth: 500,
        cellStyle: (params) => {
          // Căn phải cho số, căn trái cho text
          return isMeasure[idx]
            ? { textAlign: 'right', fontVariantNumeric: 'tabular-nums' }
            : { textAlign: 'left' }
        }
      }

      if (isMeasure[idx]) {
        colDef.type = 'numericColumn'
        colDef.valueFormatter = (params) => {
          const v = params.value
          if (v == null || v === '') return ''
          const num = Number(v)
          if (isNaN(num)) return v
          // 🔹 Format với phân tách hàng nghìn, tối đa 2 chữ số thập phân
          // return num.toLocaleString('vi-VN', { maximumFractionDigits: 2 })
          return num.toLocaleString('en-US', { maximumFractionDigits: 2 })
        }

        // 🔹 ĐỔI MÀU ĐỎ nếu giá trị âm
        colDef.cellStyle = (params) => {
          const val = Number(params.value)
          if (!isNaN(val) && val < 0) {
            return {
              color: 'red',
              textAlign: 'right',
              fontVariantNumeric: 'tabular-nums'
            }
          }
          // Mặc định vẫn căn phải, giữ format số
          return { textAlign: 'right', fontVariantNumeric: 'tabular-nums' }
        }
      }

      return colDef
    }
  })

  const columnDefs = tmpColumnDefs.filter(
    (item) => item !== null && item !== undefined
  )

  return { headers, data, isMeasure, columnDefs }
}

// ======================
// 2️⃣ Hàm tạo dữ liệu tree
// ======================
function buildTree(data) {
  let idCounter = 0
  const rootMap = {}

  for (const row of data) {
    // Lấy tất cả các cấp tree_lv1...tree_lvN
    const treeLevels = Object.keys(row)
      .filter((k) => k.startsWith('tree_lv'))
      .sort((a, b) => {
        const na = parseInt(a.replace('tree_lv', ''))
        const nb = parseInt(b.replace('tree_lv', ''))
        return na - nb
      })

    let currentLevel = rootMap
    let parent = null

    // Duyệt từng cấp
    treeLevels.forEach((key, i) => {
      const value = row[key]
      if (!currentLevel[value]) {
        currentLevel[value] = {
          id: ++idCounter,
          name: value,
          level: i + 1,
          expanded: false,
          leaf: false,
          children: {}
        }
      }
      parent = currentLevel[value]
      currentLevel = parent.children
    })

    // 3️⃣ Cấp cuối cùng -> thêm dòng dữ liệu leaf (động theo keys)
    const leafNode = {
      id: ++idCounter,
      name: null,
      level: treeLevels.length + 1,
      leaf: true
    }

    // ✅ Copy toàn bộ field KHÔNG thuộc tree_lv vào leaf
    for (const [key, val] of Object.entries(row)) {
      if (!key.startsWith('tree_lv')) {
        leafNode[key] = val
      }
    }

    parent.children[`leaf_${idCounter}`] = leafNode
  }

  return Object.values(rootMap).map((n) => normalizeTree(n))
}

function normalizeTree(node) {
  if (node.children && !Array.isArray(node.children)) {
    node.children = Object.values(node.children).map((n) => normalizeTree(n))
  }
  return node
}

// 🔹 Cộng dồn giá trị từ con lên cha cho các cột measure
function aggregateTreeValues(nodes, numericCols) {
  for (const node of nodes) {
    // Nếu có children → xử lý đệ quy
    if (node.children && node.children.length > 0) {
      aggregateTreeValues(node.children, numericCols)

      // Khởi tạo tổng của cha
      numericCols.forEach((col) => {
        node[col] = 0
      })

      // Cộng dồn từ các con
      for (const child of node.children) {
        numericCols.forEach((col) => {
          const val = Number(child[col])
          if (!isNaN(val)) {
            node[col] += val
          }
        })
      }
    }
  }
}

// ======================
// 3️⃣ Flatten tree (để hiển thị)
// ======================
function flattenTree(nodes) {
  let result = []
  for (const n of nodes) {
    result.push(n)
    if (n.expanded && n.children) {
      result = result.concat(flattenTree(n.children))
    }
  }
  return result
}

// ======================
// 7️⃣ Toggle expand/collapse
// ======================
function toggleNode(nodeId) {
  // Tìm node theo ID trong dữ liệu gốc
  function recursiveToggle(nodes) {
    for (const n of nodes) {
      if (n.id == nodeId) {
        n.expanded = !n.expanded
        return true
      }
      if (n.children && recursiveToggle(n.children)) return true
    }
    return false
  }

  recursiveToggle(nestedData)

  const flatData = flattenTree(nestedData)
  // ✅ FIX: Đẩy việc cập nhật rowData vào event loop tiếp theo
  setTimeout(() => {
    gridApi.setGridOption('rowData', flatData)

    // Sau khi render xong, cuộn đến đúng node vừa click
    const rowNode = gridApi.getDisplayedRowAtIndex(
      flatData.findIndex((r) => r.id == nodeId)
    )
    if (rowNode) {
      gridApi.ensureNodeVisible(rowNode, 'middle')
    }
  }, 0) // <--- Thêm setTimeout(..., 0)
}

// search cu

// xport cu

// copy cu

// --- Hàm thực hiện copy ---
function copySelectedRows() {
  const selectedNodes = []
  // gridApi.forEachNode((node) => {
  //   if (node.isSelected()) selectedNodes.push(node)
  // })
  gridApi.forEachNodeAfterFilterAndSort((node) => {
    if (node.isSelected()) selectedNodes.push(node)
  })

  if (selectedNodes.length === 0) {
    alert('⚠️ Chưa chọn dòng nào!')
    return
  }

  const selectedData = selectedNodes.map((node) => node.data)
  const text = selectedData
    .map((row) => Object.values(row).join('\t'))
    .join('\n')

  // --- Fallback cổ điển, tương thích mọi trình duyệt / Tableau Extension ---
  const textarea = document.createElement('textarea')
  textarea.value = text
  textarea.style.position = 'fixed'
  textarea.style.top = '-9999px'
  document.body.appendChild(textarea)
  textarea.focus()
  textarea.select()

  try {
    const success = document.execCommand('copy')
    if (success) {
      console.log(`✅ Đã copy ${selectedData.length} dòng vào clipboard!`)
    } else {
      console.log('⚠️ Copy không thành công.')
    }
  } catch (err) {
    console.error('Copy lỗi:', err)
    alert('❌ Không thể copy (trình duyệt không cho phép).')
  }

  document.body.removeChild(textarea)
}

// ======= 3️⃣ TÍNH TỔNG =======
function calcTotalsTree(nodes, numericCols) {
  const totals = {}
  numericCols.forEach((col) => (totals[col] = 0))

  function traverse(nodeList) {
    for (const node of nodeList) {
      // Nếu node có children → duyệt tiếp
      if (node.children && node.children.length > 0) {
        traverse(node.children)
      }

      // Nếu node là leaf → cộng giá trị numeric
      if (node.leaf) {
        numericCols.forEach((col) => {
          const val = Number(node[col])
          if (!isNaN(val)) {
            totals[col] += val
          }
        })
      }
    }
  }

  traverse(nodes)
  return totals
}

// ✅ Tính toán chiều cao khả dụng của extension
function adjustGridHeight() {
  const container = document.querySelector('.container')
  const toolbar = document.querySelector('.toolbar')
  // const notebar = document.querySelector('.notebar')
  const gridContainer = document.getElementById('gridContainer')

  // Chiều cao toàn bộ extension
  const totalHeight = window.innerHeight

  // Trừ phần toolbar + padding + margin
  const toolbarHeight = toolbar.offsetHeight
  const notebarHeight = notebar.offsetHeight
  const padding = 20 // tổng trên + dưới
  const extraSpacing = 10 // khoảng cách phụ nếu có

  const gridHeight =
    totalHeight - toolbarHeight - notebarHeight - padding - extraSpacing
  gridContainer.style.height = `${gridHeight}px`
}

// ======================
// Helper cho export: Flatten tree với path và tính max level (FIX: không thừa level cho leaf)
// ======================
// function exportFlattenWithPath(
//   nodes,
//   currentPath = [],
//   result = [],
//   maxLevelRef = { max: 0 }
// ) {
//   for (const node of nodes) {
//     // Chỉ tính max cho non-leaf (cha có children), tránh thừa level từ leaf rỗng
//     if (!node.leaf && node.name) {
//       const nodePath = [...currentPath, node.name]
//       maxLevelRef.max = Math.max(maxLevelRef.max, nodePath.length)
//       const row = { ...node, path: nodePath } // Copy node + path cho cha
//       result.push(row)
//     } else if (node.leaf) {
//       // Leaf: dùng path của parent (không thêm level rỗng), copy data measures
//       const leafRow = { ...node, path: currentPath } // Path không thêm ''
//       // Copy measures từ leaf (nếu có aggregate từ con, nhưng leaf gốc có data)
//       result.push(leafRow)
//     }

//     // Recurse children (flatten hết cho export full)
//     if (node.children && node.children.length > 0) {
//       exportFlattenWithPath(
//         node.children,
//         node.children.length > 0 ? [...currentPath, node.name] : currentPath,
//         result,
//         maxLevelRef
//       )
//     }
//   }
//   return result
// }

// ======================
// Helper cho export: Flatten tree với path và tính max level (FIX: chỉ visible theo expanded, không thừa level cho leaf)
// ======================
function exportFlattenWithPath(
  nodes,
  currentPath = [],
  result = [],
  maxLevelRef = { max: 0 }
) {
  for (const node of nodes) {
    // Luôn push node hiện tại (vì nếu đến đây thì node này visible)
    if (!node.leaf && node.name) {
      // Non-leaf (cha): thêm name vào path
      const nodePath = [...currentPath, node.name]
      maxLevelRef.max = Math.max(maxLevelRef.max, nodePath.length)
      const row = { ...node, path: nodePath } // Copy node + path cho cha
      result.push(row)
    } else if (node.leaf) {
      // Leaf: dùng path của parent (không thêm level rỗng), copy data measures
      const leafRow = { ...node, path: currentPath } // Path không thêm ''
      result.push(leafRow)
    }

    // Recurse children CHỈ NẾU expanded (để chỉ lấy visible)
    if (node.expanded && node.children && node.children.length > 0) {
      exportFlattenWithPath(
        node.children,
        [...currentPath, node.name],
        result,
        maxLevelRef
      )
    }
  }
  return result
}

// Load lại dữ liệu và render
function loadAndRender(worksheet) {
  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
    let idCounter = 0

    // ======================
    // 1️⃣ Dữ liệu gốc
    // ======================

    // console.log('sumData', sumData)

    // Xác định cột cần loại bỏ
    const excludeCols = sumData.columns
      .map((col, idx) => ({ name: col.fieldName, idx }))
      .filter(
        (c) =>
          c.name.toLowerCase().startsWith('hiden') || c.name.startsWith('AGG')
      )
      .map((c) => c.idx)

    const { headers, data, isMeasure, columnDefs } = pivotMeasureValues(
      sumData,
      excludeCols
    )

    // console.log('headers', headers)
    // console.log('columnDefs', columnDefs)
    // console.log('data', data)

    // console.log('isMeasure', isMeasure)

    // ======= DÒNG TỔNG =======
    function updateFooterTotals() {
      if (!gridApi) return

      const allData = []
      gridApi.forEachNodeAfterFilterAndSort((node) => {
        if (!node.rowPinned) {
          // Chỉ lấy dòng thường, không lấy dòng pinned
          allData.push(node.data)
        }
      })

      const numericCols = columnDefs
        .filter((col) => col.type === 'numericColumn')
        .map((col) => col.field)

      const totals = calcTotalsTree(allData, numericCols)

      const totalRow = {}
      columnDefs.forEach((col) => {
        const field = col.field
        if (numericCols.includes(field)) {
          totalRow[field] = totals[field]
        } else if (field === columnDefs[0].field) {
          totalRow[field] = 'Tổng cộng'
        } else {
          totalRow[field] = ''
        }
      })

      totalRow.leaf = true

      gridApi.setGridOption('pinnedBottomRowData', [totalRow])
    }

    funcTionWait4ToUpdateTotal(1000)

    // ======================
    // 4️⃣ Tree data + Flatten ban đầu
    // ======================
    nestedData = buildTree(data)

    maxTreeLevel = getMaxTreeLevel(nestedData)
    currentExpandedLevel = 1 // ban đầu chỉ hiển thị root

    // ✅ Xác định các cột numeric
    const numericCols = columnDefs
      .filter((col) => col.type === 'numericColumn')
      .map((col) => col.field)

    // ✅ Gọi hàm cộng dồn giá trị
    aggregateTreeValues(nestedData, numericCols)

    // ✅ Sau đó mới flatten để render
    let flatData = flattenTree(nestedData)

    // console.log('data', data)
    // console.log('nestedData', nestedData)
    // console.log('flatData', flatData)

    // ======================
    // 6️⃣ Cấu hình AG Grid
    // ======================
    const gridOptions = {
      columnDefs,
      rowData: flatData,
      defaultColDef: {
        filter: false, // chuyển sang false vì ko dùng filter nữa
        sortable: true,
        resizable: true
        // bỏ tham số filter đi vì không dùng filter nữa
        // filterParams: {
        //   textFormatter: (value) => normalizeUnicode(value)
        // }
      },
      // 🔹 Làm nổi bật các dòng tổng (cha)
      getRowStyle: (params) => {
        const node = params.data
        if (!node) return null

        // ✅ THÊM KIỂM TRA TẠI ĐÂY
        // if (!columnDefs || columnDefs.length === 0 || !columnDefs[0].field) {
        //   return null
        // }

        // ✅ Nếu là dòng "Tổng cộng"
        if (node[columnDefs[0].field] === 'Tổng cộng') {
          return {
            fontWeight: 'bold',
            color: '#d00000',
            backgroundColor: '#fabcbcff' // nền nhạt cho dễ nhìn
          }
        }

        // Dòng cha (có children) → in đậm
        if (node.children && node.children.length > 0) {
          return {
            fontWeight: 'bold',
            backgroundColor: '#f7f7f7' // nhẹ cho dễ nhìn, có thể bỏ
          }
        }

        // Dòng leaf → style bình thường
        return null
      },

      rowSelection: {
        mode: 'multiRow',
        checkboxes: true,
        enableClickSelection: false
      },
      // suppressRowClickSelection: false,
      // suppressInjectStyles: true, // ✅ fix bug injection CSS

      // sự kiện click vào 1 cell
      onCellClicked: (params) => {
        const el = params.event.target
        if (el.classList.contains('toggle-btn')) {
          toggleNode(el.dataset.id)
        } else {
          selectedCellValue = params.value
          console.log('Selected cell value:', selectedCellValue)
          // Bỏ chọn tất cả dòng khác
          gridApi.deselectAll()
          // Chọn dòng hiện tại
          params.node.setSelected(true)
        }
      },
      onGridReady: (params) => {
        gridApi = params.api
        // safeUpdateTotals()
        // updateFooterTotals()
        // setTimeout(() => updateFooterTotals(), 1000)
        funcTionWait4ToUpdateTotal(1000)
        console.log('run onGridReady.')
      },
      // onFirstDataRendered: () => updateFooterTotals(),
      onFirstDataRendered: (params) => {
        // updateFooterTotals()
        funcTionWait4ToUpdateTotal(1000)
        console.log('run onFirstDataRendered.')
      },
      // onFilterChanged: () => safeUpdateTotals(), xxx4
      // onSortChanged: () => safeUpdateTotals(), xxx5
      onCellContextMenu: (params) => {
        const node = params.data
        if (!node || !node.id) return

        params.event.preventDefault() // chặn menu mặc định của trình duyệt

        showContextMenu(params.event.pageX, params.event.pageY, node.id)
      }
    }

    const eGridDiv = document.querySelector('#gridContainer')

    if (!gridApi) {
      // ❗ Chỉ tạo grid 1 lần
      gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    } else {
      // ✅ Cập nhật lại dữ liệu
      gridApi.setGridOption('rowData', flatData)
      gridApi.setGridOption('columnDefs', columnDefs)
      // updateFooterTotals()
      // xxx1
      // setTimeout(() => {
      //   safeUpdateTotals()
      // }, 100)
    }

    // Code mở tất cả và đóng tất cả tree
    if (!expandListenersBound) {
      const btnExpand = document.getElementById('btnExpandAll')
      const btnCollapse = document.getElementById('btnCollapseAll')

      btnExpand1Level.addEventListener('click', () => {
        if (currentExpandedLevel < maxTreeLevel) {
          currentExpandedLevel += 1
        }

        applyExpandLevel(nestedData, currentExpandedLevel)

        const flat = flattenTree(nestedData)
        setTimeout(() => {
          gridApi.setGridOption('rowData', flat)
        }, 0)
      })

      btnCollapse1Level.addEventListener('click', () => {
        if (currentExpandedLevel > 1) {
          currentExpandedLevel -= 1
        }

        applyExpandLevel(nestedData, currentExpandedLevel)

        const flat = flattenTree(nestedData)
        setTimeout(() => {
          gridApi.setGridOption('rowData', flat)
        }, 0)
      })

      if (btnExpand) {
        btnExpand.addEventListener('click', () => {
          // Lấy node đang chọn
          const selectedNodes = []
          gridApi.forEachNode((node) => {
            if (node.isSelected()) selectedNodes.push(node.data)
          })

          // Node mục tiêu để scroll lại (nếu có chọn)
          const targetId = selectedNodes.length > 0 ? selectedNodes[0].id : null

          // Expand logic
          if (!targetId) {
            setAllExpanded(nestedData, true)
          } else {
            const node = findNodeById(nestedData, targetId)
            if (node) setSubtreeExpanded(node, true)
          }

          const flat = flattenTree(nestedData)
          // ✅ FIX: Sử dụng setTimeout(..., 0) để cập nhật rowData bất đồng bộ
          setTimeout(() => {
            gridApi.setGridOption('rowData', flat)

            // === GIỐNG toggleNode() ===
            if (targetId) {
              // requestAnimationFrame được giữ lại bên trong setTimeout để đảm bảo grid đã render
              requestAnimationFrame(() => {
                const idx = flat.findIndex((r) => r.id == targetId)
                const rowNode = gridApi.getDisplayedRowAtIndex(idx)
                if (rowNode) {
                  gridApi.ensureNodeVisible(rowNode, 'middle')
                }
              })
            }
          }, 0) // <-- Thêm setTimeout

          currentExpandedLevel = maxTreeLevel
        })
      }

      if (btnCollapse) {
        btnCollapse.addEventListener('click', () => {
          // Lấy node đang chọn
          const selectedNodes = []
          gridApi.forEachNode((node) => {
            if (node.isSelected()) selectedNodes.push(node.data)
          })

          // Node mục tiêu
          const targetId = selectedNodes.length > 0 ? selectedNodes[0].id : null

          if (!targetId) {
            setAllExpanded(nestedData, false)
          } else {
            const node = findNodeById(nestedData, targetId)
            if (node) setSubtreeExpanded(node, false)
          }

          const flat = flattenTree(nestedData)
          // ✅ FIX: Thêm setTimeout xxx
          // setTimeout(() => {
          //   gridApi.setGridOption('rowData', flat)
          //   safeUpdateTotals(gridApi)

          //   // === GIỐNG toggleNode() ===
          //   if (targetId) {
          //     requestAnimationFrame(() => {
          //       const idx = flat.findIndex((r) => r.id == targetId)
          //       const rowNode = gridApi.getDisplayedRowAtIndex(idx)
          //       if (rowNode) {
          //         gridApi.ensureNodeVisible(rowNode, 'middle')
          //       }
          //     })
          //   }
          // }, 0) // <-- Thêm setTimeout

          currentExpandedLevel = 1
        })
      }

      expandListenersBound = true
    }

    // ======================
    // Tìm kiếm toàn bộ
    // ======================
    document.getElementById('globalSearch').addEventListener('input', (e) => {
      gridApi.setGridOption('quickFilterText', normalizeUnicode(e.target.value))
      // safeUpdateTotals() // ✅ gọi đúng xxx7
      // updateFooterTotals()
    })

    function funcTionWait4ToUpdateTotal(secondsamt) {
      setTimeout(() => {
        document.getElementById('updateTotal').click() // 👈 Tự động kích nút
      }, secondsamt)
    }

    document
      .getElementById('clearAllFilterBtn')
      .addEventListener('click', () => {
        if (!gridApi) return

        // 🔹 1️⃣ Xoá toàn bộ filter theo cột
        // gridApi.setFilterModel(null)  // bỏ đi vì không dùng filter nữa
        gridApi.onFilterChanged()

        // 🔹 2️⃣ Xoá luôn filter toàn cục (search box)
        const globalSearch = document.getElementById('globalSearch')
        if (globalSearch) {
          globalSearch.value = ''
          gridApi.setGridOption('quickFilterText', '')
        }

        // 🔹 3️⃣ Cập nhật lại dòng tổng
        // safeUpdateTotals() // ✅ gọi đúng xxx8
        // updateFooterTotals()
      })

    document
      .getElementById('updateTotal')
      .addEventListener('click', updateFooterTotals)
  })
}

// Khi DOM ready
document.addEventListener('DOMContentLoaded', () => {
  tableau.extensions.initializeAsync().then(() => {
    const worksheet =
      tableau.extensions.dashboardContent.dashboard.worksheets.find(
        (ws) => ws.name === 'DataTableExtSheet'
      )

    if (!worksheet) {
      console.error("❌ Không tìm thấy worksheet tên 'DataTableExtSheet'")
      return
    }

    function refreshExtractTime() {
      worksheet.getDataSourcesAsync().then((dataSources) => {
        dataSources.forEach((ds) => {
          if (ds.isExtract) {
            extractRefreshTime = 'Extract Refresh Time: ' + ds.extractUpdateTime
          } else {
            extractRefreshTime = ''
          }

          document.getElementById('extractRefreshTime').innerText =
            extractRefreshTime
        })
      })
    }

    refreshExtractTime()

    // Load lần đầu
    loadAndRender(worksheet)

    // ======================
    // Export CSV -> tree không thò thụt được khi export csv
    // ======================
    // document.getElementById('exportExcel').addEventListener('click', () => {
    //   gridApi.exportDataAsCsv({
    //     fileName: 'tree_data.csv'
    //   })
    // })

    // ======================
    // Export CSV -> tree thò thụt được khi export csv
    // ======================
    // document.getElementById('exportExcel').addEventListener('click', () => {
    //   const allRows = []
    //   gridApi.forEachNodeAfterFilterAndSort((node) => {
    //     allRows.push(node.data)
    //   })

    //   // 🔹 Lấy pinned bottom rows (ví dụ: dòng tổng cộng)
    //   const pinnedRows = gridApi.getPinnedBottomRowCount()
    //     ? Array.from(
    //         { length: gridApi.getPinnedBottomRowCount() },
    //         (_, i) => gridApi.getPinnedBottomRow(i).data
    //       )
    //     : []

    //   // 🔹 Gộp lại (dòng tổng ở cuối)
    //   const exportRows = [...allRows, ...pinnedRows]

    //   const displayedCols = gridApi.getColumnDefs().map((c) => c.field)
    //   const headers = displayedCols.join(',')

    //   const csvRows = exportRows.map((row) => {
    //     return displayedCols
    //       .map((col) => {
    //         let val = row[col] ?? ''
    //         if (col === 'name' && row.level) {
    //           const indent = '  '.repeat(row.level - 1)
    //           val = indent + val
    //         }
    //         // Escape CSV nếu có dấu phẩy, nháy kép hoặc xuống dòng
    //         if (typeof val === 'string' && val.match(/[",\n]/)) {
    //           val = '"' + val.replace(/"/g, '""') + '"'
    //         }
    //         return val
    //       })
    //       .join(',')
    //   })

    //   // ⚡ Thêm BOM UTF-8 để Excel đọc đúng tiếng Việt
    //   const bom = '\uFEFF'
    //   const csvContent = [headers, ...csvRows].join('\n')

    //   const blob = new Blob([bom + csvContent], {
    //     type: 'text/csv;charset=utf-8;'
    //   })

    //   const link = document.createElement('a')
    //   link.href = URL.createObjectURL(blob)
    //   link.download = 'tree_data.csv'
    //   document.body.appendChild(link)
    //   link.click()
    //   document.body.removeChild(link)
    // })

    // ======================
    // Export CSV -> tree với mỗi level là cột riêng (chỉ sửa phần này)
    // ======================
    // document.getElementById('exportExcel').addEventListener('click', () => {
    //   if (!gridApi || !nestedData || nestedData.length === 0) {
    //     alert('⚠️ Không có dữ liệu để export!')
    //     return
    //   }

    //   // Flatten tree với path (full data, ignore filter/sort cho export toàn bộ)
    //   const maxLevelRef = { max: 0 }
    //   const exportRows = exportFlattenWithPath(nestedData, [], [], maxLevelRef)
    //   const maxTreeLevel = maxLevelRef.max

    //   // Lấy pinned bottom rows (dòng tổng)
    //   const pinnedRows =
    //     gridApi.getPinnedBottomRowCount() > 0
    //       ? Array.from(
    //           { length: gridApi.getPinnedBottomRowCount() },
    //           (_, i) => gridApi.getPinnedBottomRow(i).data
    //         )
    //       : []

    //   // Gộp rows (thêm pinned ở cuối)
    //   const allExportRows = [...exportRows, ...pinnedRows]

    //   // Lấy columnDefs hiện tại (cột name là cột đầu, bỏ nó đi vì ta dùng levels thay thế)
    //   const currentColumnDefs = gridApi.getColumnDefs()
    //   const firstField = currentColumnDefs[0].field // 'name'
    //   const otherCols = currentColumnDefs.slice(1).map((c) => c.field) // Các cột measure khác

    //   // Headers: Level 1 to max + other cols
    //   const levelHeaders = Array.from(
    //     { length: maxTreeLevel },
    //     (_, i) => `Level ${i + 1}`
    //   )
    //   const exportHeaders = [...levelHeaders, ...otherCols]
    //   const headers = exportHeaders.join(',')

    //   // Build CSV rows
    //   const csvRows = allExportRows.map((row) => {
    //     let rowVals = []
    //     const isTotal = row[firstField] === 'Tổng cộng' // Dòng tổng
    //     if (isTotal) {
    //       // Dòng tổng: 'Tổng cộng' ở Level 1, rỗng các level khác
    //       rowVals.push('Tổng cộng')
    //       for (let i = 1; i < maxTreeLevel; i++) {
    //         rowVals.push('')
    //       }
    //     } else {
    //       // Dòng tree: dùng path để điền levels
    //       const path = row.path || []
    //       for (let i = 0; i < maxTreeLevel; i++) {
    //         rowVals.push(path[i] || '')
    //       }
    //     }

    //     // Thêm other cols (measures, v.v.)
    //     otherCols.forEach((col) => {
    //       let val = row[col] ?? ''
    //       // Escape CSV nếu cần (phẩy, nháy, xuống dòng)
    //       if (typeof val === 'string' && val.match(/[",\n]/)) {
    //         val = '"' + val.replace(/"/g, '""') + '"'
    //       }
    //       rowVals.push(val)
    //     })

    //     return rowVals.join(',')
    //   })

    //   // Tạo file CSV với BOM UTF-8 cho tiếng Việt
    //   const bom = '\uFEFF'
    //   const csvContent = [headers, ...csvRows].join('\n')
    //   const blob = new Blob([bom + csvContent], {
    //     type: 'text/csv;charset=utf-8;'
    //   })
    //   const link = document.createElement('a')
    //   link.href = URL.createObjectURL(blob)
    //   link.download = 'tree_data.csv'
    //   document.body.appendChild(link)
    //   link.click()
    //   document.body.removeChild(link)

    //   console.log(
    //     `✅ Đã export ${allExportRows.length} rows với ${maxTreeLevel} levels!`
    //   )
    // })

    // ======================
    // Export EXCEL -> tree với mỗi level là cột riêng (chỉ sửa phần này)
    // fix lỗi liên quan đến mất số 0 ở đầu
    // ======================
    document.getElementById('exportExcel').addEventListener('click', () => {
      if (!gridApi || !nestedData || nestedData.length === 0) {
        alert('⚠️ Không có dữ liệu để export!')
        return
      }

      // Flatten tree (full data)
      const maxLevelRef = { max: 0 }
      const exportRows = exportFlattenWithPath(nestedData, [], [], maxLevelRef)
      const maxTreeLevel = maxLevelRef.max

      // Pinned bottom rows
      const pinnedRows =
        gridApi.getPinnedBottomRowCount() > 0
          ? Array.from(
              { length: gridApi.getPinnedBottomRowCount() },
              (_, i) => gridApi.getPinnedBottomRow(i).data
            )
          : []

      const allExportRows = [...exportRows, ...pinnedRows]

      const currentColumnDefs = gridApi.getColumnDefs()
      const firstField = currentColumnDefs[0].field // 'name'
      const otherCols = currentColumnDefs.slice(1).map((c) => c.field)

      // Headers
      const levelHeaders = Array.from(
        { length: maxTreeLevel },
        (_, i) => `Level ${i + 1}`
      )
      const exportHeaders = [...levelHeaders, ...otherCols]

      // Build worksheet data
      const worksheetData = []
      worksheetData.push(exportHeaders)

      allExportRows.forEach((row) => {
        const rowVals = []
        const isTotal = row[firstField] === 'Tổng cộng'

        if (isTotal) {
          rowVals.push('Tổng cộng')
          for (let i = 1; i < maxTreeLevel; i++) rowVals.push('')
        } else {
          const path = row.path || []
          for (let i = 0; i < maxTreeLevel; i++) rowVals.push(path[i] || '')
        }

        // Add other columns, giữ 0 đầu bằng cách ép thành string
        otherCols.forEach((col) => {
          let val = row[col] ?? ''
          if (typeof val === 'number') {
            rowVals.push(val)
          } else {
            // ép dạng text EXCEL để giữ 0 đầu
            rowVals.push(val.toString())
          }
        })

        worksheetData.push(rowVals)
      })

      // Tạo workbook XLSX
      const ws = XLSX.utils.aoa_to_sheet(worksheetData)

      // ⭐ Force tất cả dimension dạng text (giữ số 0 đầu)
      Object.keys(ws).forEach((cell) => {
        if (!cell.startsWith('!')) {
          const value = ws[cell].v
          if (typeof value === 'string' && /^\d+$/.test(value)) {
            ws[cell].t = 's' // string
          }
        }
      })

      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, ws, 'TreeData')

      // Xuất file
      XLSX.writeFile(wb, 'tree_data.xlsx')

      console.log(`✅ Export Excel thành công (${allExportRows.length} dòng)!`)
    })

    // --- Copy bằng nút bấm ---
    document.getElementById('copyRow').addEventListener('click', () => {
      copySelectedRows()
    })

    document.getElementById('copyCellBtn').addEventListener('click', () => {
      if (selectedCellValue === null) {
        alert('Chưa chọn ô nào để copy!')
        return
      }

      const text = selectedCellValue.toString()

      // --- Fallback cổ điển ---
      const textarea = document.createElement('textarea')
      textarea.value = text
      textarea.style.position = 'fixed'
      textarea.style.top = '-9999px'
      document.body.appendChild(textarea)
      textarea.focus()
      textarea.select()

      try {
        const success = document.execCommand('copy')
        if (success) {
          console.log(`✅ Đã copy ô: ${text}`)
        } else {
          console.log('⚠️ Copy không thành công.')
        }
      } catch (err) {
        console.error('Copy lỗi:', err)
        alert('❌ Không thể copy (trình duyệt không cho phép).')
      }

      document.body.removeChild(textarea)
    })

    // Lắng nghe filter và parameter change
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, () => {
      // console.log('vao day roi')
      refreshExtractTime()
      loadAndRender(worksheet)
    })

    tableau.extensions.dashboardContent.dashboard
      .getParametersAsync()
      .then(function (parameters) {
        parameters.forEach(function (p) {
          p.addEventListener(tableau.TableauEventType.ParameterChanged, () => {
            // console.log('vao day roi 2')
            refreshExtractTime()
            loadAndRender(worksheet)
          })
        })
      })

    // Gọi khi load trang và khi resize
    adjustGridHeight()
    window.addEventListener('resize', adjustGridHeight)
  })
})
