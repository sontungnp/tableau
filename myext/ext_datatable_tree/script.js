'use strict'

let selectedCellValue = null
let expandListenersBound = false // <-- thêm dòng này
let extractRefreshTime = ''

let gridApi = null
let nestedData = []

function setAllExpanded(nodes, expanded) {
  if (!nodes || !nodes.length) return
  for (const n of nodes) {
    if (n.children && n.children.length) {
      n.expanded = expanded
      setAllExpanded(n.children, expanded)
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
              return indent + '📄 ' + (node.name || '')
            } else {
              const symbol = node.expanded ? '➖' : '➕'
              return (
                indent +
                `<span class="toggle-btn" data-id="${node.id}">${symbol}</span> 📁 ` +
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
        autoHeight: true,
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
          return num.toLocaleString('vi-VN', { maximumFractionDigits: 2 })
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
  gridApi.setGridOption('rowData', flatData)

  // Sau khi render xong, cuộn đến đúng node vừa click
  requestAnimationFrame(() => {
    const rowNode = gridApi.getDisplayedRowAtIndex(
      flatData.findIndex((r) => r.id == nodeId)
    )
    if (rowNode) {
      gridApi.ensureNodeVisible(rowNode, 'middle')
    }
  })
}

// search cu

// xport cu

// copy cu

// --- Hàm thực hiện copy ---
function copySelectedRows() {
  const selectedNodes = []
  gridApi.forEachNode((node) => {
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
      const allData = []
      gridApi.forEachNodeAfterFilterAndSort((node) => allData.push(node.data))

      const numericCols = columnDefs
        .filter((col) => col.type === 'numericColumn')
        .map((col) => col.field)

      const totals = calcTotalsTree(allData, numericCols)

      // 🟢 Tạo 1 dòng "tổng cộng"
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

      // ✅ Gán dòng này thành pinned bottom row
      gridApi.setGridOption('pinnedBottomRowData', [totalRow])
    }

    // ======================
    // 4️⃣ Tree data + Flatten ban đầu
    // ======================
    nestedData = buildTree(data)
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
        filter: true,
        sortable: true,
        resizable: true,
        filterParams: {
          textFormatter: (value) => normalizeUnicode(value)
        }
      },
      // 🔹 Làm nổi bật các dòng tổng (cha)
      getRowStyle: (params) => {
        const node = params.data
        if (!node) return null

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
        checkboxes: true
      },
      suppressRowClickSelection: false,
      suppressInjectStyles: true, // ✅ fix bug injection CSS

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
      onGridReady: () => updateFooterTotals(),
      onFilterChanged: () => updateFooterTotals(),
      onSortChanged: () => updateFooterTotals()
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
      setTimeout(() => {
        updateFooterTotals()
      }, 300)
    }

    // Code mở tất cả và đóng tất cả tree
    if (!expandListenersBound) {
      const btnExpand = document.getElementById('btnExpandAll')
      const btnCollapse = document.getElementById('btnCollapseAll')

      if (btnExpand) {
        btnExpand.addEventListener('click', () => {
          // mở toàn bộ cây
          setAllExpanded(nestedData, true)
          const flat = flattenTree(nestedData)
          gridApi.setGridOption('rowData', flat)
          // updateFooterTotals //&& updateFooterTotals()
          setTimeout(() => {
            updateFooterTotals()
          }, 300)
          // nếu muốn scroll tới đầu:
          // const vp = gridApi.gridBodyCtrl?.eBodyViewport; if (vp) vp.scrollTop = 0
        })
      }

      if (btnCollapse) {
        btnCollapse.addEventListener('click', () => {
          // đóng toàn bộ cây
          setAllExpanded(nestedData, false)
          const flat = flattenTree(nestedData)
          gridApi.setGridOption('rowData', flat)
          // updateFooterTotals && updateFooterTotals()
          setTimeout(() => {
            updateFooterTotals()
          }, 300)
        })
      }

      expandListenersBound = true
    }

    // ======================
    // Tìm kiếm toàn bộ
    // ======================
    document.getElementById('globalSearch').addEventListener('input', (e) => {
      gridApi.setGridOption('quickFilterText', normalizeUnicode(e.target.value))
      // updateFooterTotals()
      setTimeout(() => {
        updateFooterTotals()
      }, 300)
    })

    document
      .getElementById('clearAllFilterBtn')
      .addEventListener('click', () => {
        if (!gridApi) return

        // 🔹 1️⃣ Xoá toàn bộ filter theo cột
        gridApi.setFilterModel(null)
        gridApi.onFilterChanged()

        // 🔹 2️⃣ Xoá luôn filter toàn cục (search box)
        const globalSearch = document.getElementById('globalSearch')
        if (globalSearch) {
          globalSearch.value = ''
          gridApi.setGridOption('quickFilterText', '')
        }

        // 🔹 3️⃣ Cập nhật lại dòng tổng
        setTimeout(() => {
          updateFooterTotals()
        }, 300)
      })
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
    // Export CSV
    // ======================
    document.getElementById('exportExcel').addEventListener('click', () => {
      gridApi.exportDataAsCsv({
        fileName: 'tree_data.csv'
      })
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
