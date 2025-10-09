'use strict'

let selectedCellValue = null

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
    r.map((c) =>
      c.formattedValue === null || c.formattedValue === undefined
        ? ''
        : c.formattedValue
    )
  )

  // 🔹 Loại bỏ cột không cần
  const filteredCols = cols.filter((_, i) => !excludeIndexes.includes(i))
  const filteredRows = rows.map((r) =>
    r.filter((_, i) => !excludeIndexes.includes(i))
  )

  // 🔹 Xác định vị trí Measure Names / Values
  const measureNameIdx = filteredCols.findIndex((c) =>
    c.toLowerCase().includes('measure names')
  )
  const measureValueIdx = filteredCols.findIndex((c) =>
    c.toLowerCase().includes('measure values')
  )

  const dimensionIdxs = filteredCols
    .map((c, i) => i)
    .filter((i) => i !== measureNameIdx && i !== measureValueIdx)

  const pivotMap = new Map()
  const measureSet = new Set()

  filteredRows.forEach((r) => {
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

  const measureNames = Array.from(measureSet)
  const headers = [
    ...dimensionIdxs.map((i) => filteredCols[i]),
    ...measureNames
  ]
  const isMeasure = [
    ...dimensionIdxs.map(() => false),
    ...measureNames.map(() => true)
  ]

  // ⚡ Sinh dữ liệu dạng object (key = field format)
  const data = Array.from(pivotMap.values()).map((entry) => {
    const row = {}
    headers.forEach((h, idx) => {
      // Bỏ phần (width) nếu có
      const cleanHeader = h.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
      const key = formatField(cleanHeader)

      if (idx < dimensionIdxs.length) {
        row[key] = entry.dims[idx]
      } else {
        const mName = measureNames[idx - dimensionIdxs.length]
        const rawValue = entry.measures[mName] || ''
        const numValue = parseFloat(rawValue.toString().replace(/,/g, ''))
        row[key] = !isNaN(numValue) ? numValue : rawValue
      }
    })
    return row
  })

  // ⚡ columnDefs khớp field format, có xử lý width và numericColumn
  const columnDefs = headers.map((h, idx) => {
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
    }

    return colDef
  })

  return { headers, data, isMeasure, columnDefs }
}

let gridApi = null

// Load lại dữ liệu và render
function loadAndRender(worksheet) {
  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
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
    // console.log('result', result)

    console.log('isMeasure', isMeasure)

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
      onGridReady: () => updateFooterTotals(),
      onFilterChanged: () => updateFooterTotals(),
      onSortChanged: () => updateFooterTotals()
    }

    const eGridDiv = document.querySelector('#myGrid')
    // const gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    if (!gridApi) {
      // ❗ Chỉ tạo grid 1 lần
      gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    } else {
      // ✅ Cập nhật lại dữ liệu
      gridApi.setGridOption('rowData', data)
      gridApi.setGridOption('columnDefs', columnDefs)
      updateFooterTotals()
    }

    // ======= 5️⃣ TÌM KIẾM =======
    document.getElementById('searchBox').addEventListener('input', function () {
      gridApi.setGridOption('quickFilterText', normalizeUnicode(this.value))
      updateFooterTotals()
    })

    // ======= 6️⃣ EXPORT EXCEL =======
    document.getElementById('exportBtn').addEventListener('click', function () {
      gridApi.exportDataAsCsv({
        fileName: 'data_export.csv',
        processCellCallback: (params) => params.value // lấy raw value
      })
    })

    // ======= 7️⃣ DÒNG TỔNG =======
    function updateFooterTotals() {
      const allData = []
      gridApi.forEachNodeAfterFilterAndSort((node) => allData.push(node.data))

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

    // --- Copy khi Ctrl + C ---
    // document.addEventListener('keydown', (e) => {
    //   if (e.ctrlKey && e.key.toLowerCase() === 'c') {
    //     copySelectedRows()
    //   }
    // })

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
  })
}

// Khi DOM ready
document.addEventListener('DOMContentLoaded', () => {
  tableau.extensions.initializeAsync().then(() => {
    const worksheet =
      tableau.extensions.dashboardContent.dashboard.worksheets[0]

    // Load lần đầu
    loadAndRender(worksheet)

    // Lắng nghe filter và parameter change
    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, () => {
      // console.log('vao day roi')

      loadAndRender(worksheet)
    })

    tableau.extensions.dashboardContent.dashboard
      .getParametersAsync()
      .then(function (parameters) {
        parameters.forEach(function (p) {
          p.addEventListener(tableau.TableauEventType.ParameterChanged, () => {
            // console.log('vao day roi 2')
            loadAndRender(worksheet)
          })
        })
      })
  })
})
