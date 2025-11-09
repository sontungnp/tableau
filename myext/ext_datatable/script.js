'use strict'

let selectedCellValue = null
let extractRefreshTime = ''

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

  // ⚡ columnDefs khớp field format, có xử lý width và numericColumn - chỉ giữ các cột hợp lệ
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
      cellStyle: (params) => {
        // Căn phải cho số, căn trái cho text
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
        // 🔹 Format với phân tách hàng nghìn, tối đa 2 chữ số thập phân
        return num.toLocaleString('en-US', { maximumFractionDigits: 2 })
        // return num.toLocaleString('vi-VN', { maximumFractionDigits: 2 })
      }
    }

    return colDef
  })

  return {
    data,
    columnDefs
  }
}

let gridApi = null

// Load lại dữ liệu và render
function loadAndRender(worksheet) {
  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
    // console.log('sumData', sumData)

    const { data, columnDefs } = pivotMeasureValues(sumData)

    // console.log('headers', headers)
    // console.log('columnDefs', columnDefs)
    // console.log('data', data)
    // console.log('result', result)

    // console.log('isMeasure', isMeasure)

    // ======= 3️⃣ TÍNH TỔNG =======
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

    // ======= 4️⃣ CẤU HÌNH GRID =======
    const gridOptions = {
      theme: 'legacy',
      columnDefs,
      rowData: data,
      animateRows: true,
      suppressAggFuncInHeader: true,
      alwaysShowHorizontalScroll: true,
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
        // Nếu là dòng pinned bottom (Tổng cộng)
        if (params.node.rowPinned === 'bottom') {
          return {
            color: 'red', // chữ màu đỏ
            fontWeight: 'bold', // đậm cho nổi bật
            backgroundColor: '#fff5f5' // nền nhẹ (tùy chọn)
          }
        }
        return null
      },

      // sự kiện click vào 1 cell
      onCellClicked: (params) => {
        selectedCellValue = params.value
        console.log('Selected cell value:', selectedCellValue)

        // Bỏ chọn tất cả dòng khác
        gridApi.deselectAll()
        // Chọn dòng hiện tại
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
      onGridReady: (params) => {
        gridApi = params.api
        updateFooterTotals()
      },
      // onFirstDataRendered: () => safeUpdateTotals(gridApi),
      onFilterChanged: () => {
        console.log('Filter changed -> model updated incoming')
      },
      onModelUpdated: () => {
        console.log('Model updated -> rows might change')
      },
      onRowDataUpdated: () => {
        console.log('Row data updated -> safe to calculate totals')
      },
      onDisplayedColumnsChanged: () => {
        console.log('Displayed columns changed -> grid layout ready')
        updateFooterTotalsSafe()
      },
      onSortChanged: () => {
        console.log('Timeout - 268')
        safeUpdateTotals(gridApi)
      }
    }

    const eGridDiv = document.querySelector('#myGrid')
    // const gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    if (!gridApi) {
      gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    } else {
      // ✅ Cập nhật lại dữ liệu và đảm bảo tổng được tính
      gridApi.setGridOption('rowData', data)
      gridApi.setGridOption('columnDefs', columnDefs)

      // Đảm bảo tổng được tính lại sau khi set dữ liệu mới
      console.log('vong 2 timeout - 285')

      setTimeout(() => {
        safeUpdateTotals()
      }, 300)
    }

    // ======= 5️⃣ TÌM KIẾM =======
    document.getElementById('searchBox').addEventListener('input', function () {
      gridApi.setGridOption('quickFilterText', normalizeUnicode(this.value))
      // console.log('Timeout - 289')
      safeUpdateTotals() // Đảm bảo gọi đúng hàm
    })

    // export cu

    // ======= 7️⃣ DÒNG TỔNG =======
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

      const totals = calcTotals(allData, numericCols)

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

    function updateFooterTotalsSafe() {
      if (!gridApi) return

      // Đảm bảo grid DOM đã vẽ pinned container
      gridApi.ensurePinnedBottomDisplayed()

      const allData = []
      gridApi.forEachNodeAfterFilterAndSort((node) => {
        if (!node.rowPinned) allData.push(node.data)
      })

      const numericCols = gridApi
        .getColumnDefs()
        .filter((col) => col.type === 'numericColumn')
        .map((col) => col.field)

      const totals = {}
      numericCols.forEach((col) => {
        totals[col] = allData.reduce((sum, r) => sum + (Number(r[col]) || 0), 0)
      })

      const totalRow = {}
      gridApi.getColumnDefs().forEach((col, idx) => {
        if (numericCols.includes(col.field)) {
          totalRow[col.field] = totals[col.field]
        } else if (idx === 0) {
          totalRow[col.field] = 'Tổng cộng'
        } else {
          totalRow[col.field] = ''
        }
      })

      gridApi.setPinnedBottomRowData([totalRow])
    }

    function safeUpdateTotals(delay = 300) {
      if (!gridApi) return

      requestAnimationFrame(() => {
        setTimeout(() => updateFooterTotals(), delay)
        console.log('xxxxxx')
      })
    }

    // --- Copy bằng nút bấm ---
    document.getElementById('copyBtn').addEventListener('click', () => {
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

    document
      .getElementById('clearAllFilterBtn')
      .addEventListener('click', () => {
        if (!gridApi) return

        // 🔹 1️⃣ Xoá toàn bộ filter theo cột
        gridApi.setFilterModel(null)
        gridApi.onFilterChanged()

        // 🔹 2️⃣ Xoá luôn filter toàn cục (search box)
        const searchBox = document.getElementById('searchBox')
        if (searchBox) {
          searchBox.value = ''
          gridApi.setGridOption('quickFilterText', '')
        }

        // 🔹 3️⃣ Cập nhật lại dòng tổng
        console.log('Timeout - 396')
        safeUpdateTotals(gridApi)
      })

    // --- Copy khi Ctrl + C ---
    // document.addEventListener('keydown', (e) => {
    //   if (e.ctrlKey && e.key.toLowerCase() === 'c') {
    //     copySelectedRows()
    //   }
    // })

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

    // ======= 6️⃣ EXPORT EXCEL =======
    document.getElementById('exportBtn').addEventListener('click', function () {
      gridApi.exportDataAsCsv({
        fileName: 'data_export.csv',
        processCellCallback: (params) => params.value // lấy raw value
      })
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

    // ✅ Tính toán chiều cao khả dụng của extension
    function adjustGridHeight() {
      const container = document.querySelector('.container')
      const toolbar = document.querySelector('.toolbar')
      // const notebar = document.querySelector('.notebar')
      const gridContainer = document.getElementById('myGrid')

      // Chiều cao toàn bộ extension
      const totalHeight = window.innerHeight
      // console.log('totalHeight', totalHeight)

      // Trừ phần toolbar + padding + margin
      const toolbarHeight = toolbar.offsetHeight
      const notebarHeight = notebar.offsetHeight
      const padding = 20 // tổng trên + dưới
      const extraSpacing = 10 // khoảng cách phụ nếu có

      // console.log('toolbarHeight', toolbarHeight)

      const gridHeight =
        totalHeight - toolbarHeight - notebarHeight - padding - extraSpacing

      // console.log('gridHeight', gridHeight)
      gridContainer.style.height = `${gridHeight}px`
    }

    // Gọi khi load trang và khi resize
    adjustGridHeight()
    window.addEventListener('resize', adjustGridHeight)
  })
})
