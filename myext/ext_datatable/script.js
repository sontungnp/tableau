'use strict'

let selectedCellValue = null
let extractRefreshTime = ''
let gridApi = null

// Hàm chuẩn hóa chỉ để đồng bộ Unicode, không bỏ dấu
function normalizeUnicode(str) {
  return str ? str.normalize('NFC').toLowerCase().trim() : ''
}

// Hàm format ngày tháng (Date object) cho trường hợp đặc biệt
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date)) return ''
  // Định dạng dd/MM/yyyy có thêm số 0
  return date.toLocaleDateString('vi-VN', {
    day: '2-digit',
    month: '2-digit',
    year: 'numeric'
  })
}

// Hàm format số cho Grand Total và ô không có valueFormatter
function formatNumber(num) {
  if (num == null || isNaN(Number(num))) return num
  const parsedNum = Number(num)
  // Format với phân tách hàng nghìn, tối đa 2 chữ số thập phân
  return parsedNum.toLocaleString('en-US', { maximumFractionDigits: 2 })
}

// 🎯 Hàm Pivot Measure Names/Values
function pivotMeasureValues(table, fieldFormat = 'snake_case') {
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
  const rows = table.data

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

  // 🔹 Pivot dữ liệu (Tối ưu: chỉ lấy nativeValue/formattedValue thô, xử lý chuyển đổi sau)
  const pivotMap = new Map()
  const measureSet = new Set()

  rows.forEach((r) => {
    // Tối ưu: Chỉ lấy giá trị, không thực hiện định dạng phức tạp trong vòng lặp lớn
    const dims = dimensionIdxs.map((i) => r[i])

    // Key phải là chuỗi
    const dimKey = dims
      .map((c) => (c.nativeValue === null ? '' : c.nativeValue.toString()))
      .join('||')

    const mName = r[measureNameIdx].formattedValue
    const mValue = r[measureValueIdx] // Giữ nguyên CellValue object

    measureSet.add(mName)

    if (!pivotMap.has(dimKey)) {
      pivotMap.set(dimKey, {
        dims: dims,
        measures: {}
      })
    }
    // Lưu CellValue object để giữ cả nativeValue và formattedValue
    pivotMap.get(dimKey).measures[mName] = mValue
  })

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
        const cellValue = entry.dims[originalIdx]
        if (cellValue.nativeValue instanceof Date) {
          // Chỉ định dạng Date ở đây (ít tốn kém hơn so với định dạng chuỗi/số)
          row[key] = formatDate(cellValue.nativeValue)
        } else {
          row[key] =
            cellValue.formattedValue === 'Null' ? '' : cellValue.formattedValue
        }
      } else {
        // Là measure
        const mName = measureNames[originalIdx - dimensionIdxs.length]
        const cellValue = entry.measures[mName]

        // 🟢 Tối ưu: Chỉ lấy nativeValue dạng số nếu có, để ag-Grid valueFormatter lo phần định dạng.
        let value = ''
        if (cellValue && cellValue.nativeValue !== null) {
          if (typeof cellValue.nativeValue === 'number') {
            value = cellValue.nativeValue
          } else {
            // Trường hợp Measure Value là chuỗi (ví dụ: "$100.00")
            // Thử chuyển đổi chuỗi formattedValue sang số (cần làm sạch chuỗi)
            const numValue = parseFloat(
              cellValue.formattedValue.toString().replace(/[^0-9.-]+/g, '')
            )
            value = !isNaN(numValue) ? numValue : cellValue.formattedValue
          }
        }
        row[key] = value
      }
    })
    return row
  })

  // ⚡ columnDefs khớp field format, có xử lý width và numericColumn
  const columnDefs = filteredHeaders.map((h, idx) => {
    const widthMatch = h.match(/\((\d+)\)/)
    const width = widthMatch ? parseInt(widthMatch[1], 10) : 150 // mặc định 150
    const cleanHeader = h.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
    const fieldName = formatField(cleanHeader)

    const colDef = {
      field: fieldName,
      headerName: cleanHeader,
      wrapText: true,
      autoHeight: true,
      width: width,
      minWidth: 30,
      maxWidth: 500,
      // Đảm bảo chỉ căn phải cho cột số
      cellStyle: (params) => {
        return filteredIsMeasure[idx]
          ? { textAlign: 'right', fontVariantNumeric: 'tabular-nums' }
          : { textAlign: 'left' }
      }
    }

    if (filteredIsMeasure[idx]) {
      colDef.type = 'numericColumn'
      colDef.valueFormatter = (params) => {
        const v = params.value
        if (v == null || v === '') return ''
        const num = Number(v)
        if (isNaN(num)) return v
        // 🔹 Dùng hàm formatNumber chung
        return formatNumber(num)
      }
    }

    return colDef
  })

  return {
    data,
    columnDefs
  }
}

// Load lại dữ liệu và render
function loadAndRender(worksheet) {
  // Bắt đầu đo thời gian xử lý JS
  const startTime = performance.now()

  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
    // ⏰ Thời gian tải dữ liệu Tableau (API call)
    console.log(`[${new Date().toISOString()}] Tableau API returned data.`)

    const { data, columnDefs } = pivotMeasureValues(sumData)

    // ⏰ Thời gian Pivot Dữ liệu
    const pivotTime = performance.now()
    console.log(
      `[${new Date().toISOString()}] Pivot Time: ${(
        pivotTime - startTime
      ).toFixed(2)}ms`
    )

    // ======= 3️⃣ CẤU HÌNH & TÍNH TỔNG =======
    // Logic tính tổng này sẽ chỉ chạy sau khi filter/sort đã ổn định.

    const gridOptions = {
      theme: 'legacy',
      columnDefs,
      rowData: data,
      animateRows: true,
      defaultColDef: {
        sortable: true,
        filter: true,
        resizable: true,
        filterParams: {
          textFormatter: (value) => normalizeUnicode(value)
        }
      },
      rowSelection: {
        mode: 'multiRow',
        checkboxes: true
      },

      getRowStyle: (params) => {
        // Nếu là dòng pinned bottom (Grand Total)
        if (params.node.rowPinned === 'bottom') {
          return {
            color: 'red', // chữ màu đỏ
            fontWeight: 'bold', // đậm cho nổi bật
            backgroundColor: '#fff5f5', // nền nhẹ (tùy chọn)
            height: 18
          }
        }
        return null
      },

      getRowHeight: (params) => {
        if (params.data && params.data.name === 'Grand Total') return 15 // Hoặc 'auto'
        return undefined // Mặc định
      },

      onCellClicked: (params) => {
        selectedCellValue = params.value
        console.log('Selected cell value:', selectedCellValue)
        gridApi.deselectAll()
        params.node.setSelected(true)
      },

      // sự kiện click vào 1 dòng
      // onRowClicked: (event) => {
      //   // Bỏ chọn tất cả dòng khác
      //   gridApi.deselectAll()
      //   // Chọn dòng hiện tại
      //   event.node.setSelected(true)
      // },

      domLayout: 'normal',
      // Thêm sự kiện sau khi grid đã render xong dữ liệu (tùy chọn)
      onGridReady: (params) => {
        // Lần đầu tiên chạy
        gridApi = params.api
        // Bắt đầu tính tổng
        updateFooterTotals()
      },
      onFilterChanged: () => {
        console.log(`[${new Date().toISOString()}] Filter changed`)
        renderActiveFilters() // ✅ cập nhật danh sách button filter
        funcTionWait4ToUpdateTotal(1000)
      },
      onSortChanged: () => {
        funcTionWait4ToUpdateTotal(1000)
      }
    }

    const eGridDiv = document.querySelector('#myGrid')

    if (!gridApi) {
      // 🟢 Khởi tạo lần đầu
      gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    } else {
      // 🟢 Cập nhật (Tối ưu: Chỉ dùng setRowData và tránh set columnDefs nếu có thể)

      // Cập nhật columnDefs CHỈ KHI CẦN (nếu cấu trúc worksheet thay đổi)
      // Trong trường hợp này, vì pivot tạo ra cột động, ta vẫn phải cập nhật, nhưng
      // ag-Grid sẽ tự tối ưu nếu cấu trúc tương tự.
      gridApi.setGridOption('columnDefs', columnDefs)

      // Tối ưu quan trọng: dùng setRowData
      // gridApi.setRowData(data)
      gridApi.setGridOption('rowData', data)

      // Đảm bảo tổng được tính lại sau khi set dữ liệu mới
      funcTionWait4ToUpdateTotal(1000)
    }

    // ⏰ Thời gian Render ag-Grid
    const endTime = performance.now()
    console.log(
      `[${new Date().toISOString()}] Render Time (Total JS): ${(
        endTime - startTime
      ).toFixed(2)}ms`
    )

    // ======= 5️⃣ TÌM KIẾM (Chỉ khởi tạo listener 1 lần) =======
    const searchBox = document.getElementById('searchBox')
    if (searchBox && !searchBox.hasListener) {
      searchBox.addEventListener('input', function () {
        gridApi.setGridOption('quickFilterText', normalizeUnicode(this.value))
      })
      searchBox.hasListener = true // Đánh dấu đã gắn listener
    }

    function funcTionWait4ToUpdateTotal(secondsamt) {
      setTimeout(() => {
        document.getElementById('updateTotal').click() // 👈 Tự động kích nút
      }, secondsamt)
    }

    // ======= 7️⃣ DÒNG TỔNG (Giữ nguyên, cần chạy sau khi filter/sort ổn định) =======
    function calcTotals(data, numericCols) {
      const totals = {}
      numericCols.forEach((col) => {
        totals[col] = data.reduce(
          (sum, row) => sum + (Number(row[col]) || 0),
          0
        )
      })
      return totals
    }

    function updateFooterTotals() {
      if (!gridApi) return

      const allData = []
      // Tối ưu: ag-Grid nhanh hơn khi lặp qua node so với tính toán lại từ đầu
      gridApi.forEachNodeAfterFilterAndSort((node) => {
        if (!node.rowPinned) {
          allData.push(node.data)
        }
      })

      const numericCols = columnDefs
        .filter((col) => col.type === 'numericColumn')
        .map((col) => col.field)

      const totals = calcTotals(allData, numericCols)

      // 🟢 Tạo 1 dòng "Grand Total"
      const totalRow = {}
      columnDefs.forEach((col) => {
        const field = col.field
        if (numericCols.includes(field)) {
          // Format tổng số bằng hàm chung
          totalRow[field] = totals[field]
        } else if (field === columnDefs[0].field) {
          totalRow[field] = 'Grand Total'
        } else {
          totalRow[field] = ''
        }
      })

      // ✅ Gán dòng này thành pinned bottom row
      gridApi.setGridOption('pinnedBottomRowData', [totalRow])
    }

    // --- Các hàm và listener khác (chỉ khởi tạo 1 lần) ---
    // const isListenersInitialized = document.body.getAttribute(
    //   'data-listeners-initialized'
    // )
    // if (!isListenersInitialized) {
    document
      .getElementById('updateTotal')
      .addEventListener('click', updateFooterTotals)

    document
      .getElementById('clearAllFilterBtn')
      .addEventListener('click', () => {
        if (!gridApi) return
        gridApi.setFilterModel(null)
        const searchBox = document.getElementById('searchBox')
        if (searchBox) {
          searchBox.value = ''
          gridApi.setGridOption('quickFilterText', '')
        }
        gridApi.onFilterChanged()
        funcTionWait4ToUpdateTotal(1000)
      })

    // Giữ nguyên các hàm renderActiveFilters, copySelectedRows, copyCellBtn, copyBtn...

    function renderActiveFilters() {
      if (!gridApi) return

      const filterModel = gridApi.getFilterModel()
      const filterArea = document.getElementById('filter-area')
      filterArea.innerHTML = ''

      if (Object.keys(filterModel).length === 0) {
        filterArea.innerHTML = `<span style="color:#888;">Không có filter nào</span>`
        return
      }

      Object.keys(filterModel).forEach((col) => {
        const btn = document.createElement('button')
        btn.textContent = col
        btn.addEventListener('click', () => {
          const model = gridApi.getFilterModel()
          delete model[col]
          gridApi.setFilterModel(model)
          gridApi.onFilterChanged()
          renderActiveFilters()
        })
        filterArea.appendChild(btn)
      })
    }

    function copySelectedRows() {
      const selectedNodes = []
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

    document
      .getElementById('copyBtn')
      .addEventListener('click', copySelectedRows)

    document.getElementById('copyCellBtn').addEventListener('click', () => {
      if (selectedCellValue === null) {
        alert('Chưa chọn ô nào để copy!')
        return
      }

      const text = selectedCellValue.toString()

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

    document.body.setAttribute('data-listeners-initialized', 'true')
    // }
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

    // ======= 6️⃣ EXPORT EXCEL (Chỉ khởi tạo listener 1 lần) =======
    const exportBtn = document.getElementById('exportBtn')
    if (exportBtn && !exportBtn.hasListener) {
      exportBtn.addEventListener('click', function () {
        gridApi.exportDataAsCsv({
          fileName: 'data_export.csv',
          processCellCallback: (params) => {
            // Sử dụng raw value để export chính xác hơn
            return params.value
          }
        })
      })
      exportBtn.hasListener = true
    }

    // Lắng nghe filter và parameter change
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, () => {
      refreshExtractTime()
      loadAndRender(worksheet)
    })

    tableau.extensions.dashboardContent.dashboard
      .getParametersAsync()
      .then(function (parameters) {
        parameters.forEach(function (p) {
          p.addEventListener(tableau.TableauEventType.ParameterChanged, () => {
            refreshExtractTime()
            loadAndRender(worksheet)
          })
        })
      })

    // ✅ Tính toán chiều cao khả dụng của extension
    function adjustGridHeight() {
      const toolbar = document.querySelector('.toolbar')
      const notebar = document.querySelector('.notebar')
      const gridContainer = document.getElementById('myGrid')

      const totalHeight = window.innerHeight
      const toolbarHeight = toolbar ? toolbar.offsetHeight : 0
      const notebarHeight = notebar ? notebar.offsetHeight : 0
      const padding = 20 // tổng trên + dưới
      const extraSpacing = 10 // khoảng cách phụ nếu có

      const gridHeight =
        totalHeight - toolbarHeight - notebarHeight - padding - extraSpacing

      if (gridContainer) {
        gridContainer.style.height = `${gridHeight}px`
      }

      // Kích hoạt ag-Grid điều chỉnh lại layout
      if (gridApi) {
        gridApi.sizeColumnsToFit()
      }
    }

    // Gọi khi load trang và khi resize
    adjustGridHeight()
    window.addEventListener('resize', adjustGridHeight)
  })
})
