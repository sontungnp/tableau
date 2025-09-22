'use strict'

const activeFilters = {} // key: column index, value: array các giá trị được chọn

// Hàm chuẩn hóa chỉ để đồng bộ Unicode, không bỏ dấu
function normalizeUnicode(str) {
  return str ? str.normalize('NFC').toLowerCase().trim() : ''
}

function adjustHeaderContainerWidth() {
  const tableEl = document.getElementById('data-table')
  const headerContainer = document.querySelector('.header-container')
  if (tableEl && headerContainer) {
    headerContainer.style.width = tableEl.offsetWidth + 'px'
  }
}

// Hàm đo độ rộng text
function getTextWidth(text, font = '14px Arial') {
  const canvas = document.createElement('canvas')
  const context = canvas.getContext('2d')
  context.font = font
  return context.measureText(text).width
}

// Hàm format số
function formatNumber(value) {
  if (value === null || value === undefined || value === '') return ''
  const num = Number(value.toString().replace(/,/g, ''))
  if (isNaN(num)) return value
  return num.toLocaleString('en-US') // có thể đổi sang 'vi-VN'
}

// Render table
function renderTable(headers, data, colWidths, isMeasure) {
  const thead = document.getElementById('table-header')
  const tfilter = document.getElementById('table-filter')
  const tbody = document.getElementById('table-body')
  thead.innerHTML = ''
  tfilter.innerHTML = ''
  tbody.innerHTML = ''

  // Xác định các cột cần ẩn
  const columnsToHide = headers
    .map((header, index) => ({ header, index }))
    .filter((item) => item.header.toLowerCase().startsWith('hiden'))
    .map((item) => item.index)

  // Lọc chỉ các cột hiển thị
  const visibleHeaders = headers.filter(
    (header, index) => !columnsToHide.includes(index)
  )
  const visibleColWidths = colWidths.filter(
    (width, index) => !columnsToHide.includes(index)
  )
  const visibleIsMeasure = isMeasure.filter(
    (measure, index) => !columnsToHide.includes(index)
  )

  // Lọc dữ liệu - chỉ giữ các cột visible
  const visibleData = data.map((row) =>
    row.filter((cell, index) => !columnsToHide.includes(index))
  )

  // Header căn giữa
  visibleHeaders.forEach((h, idx) => {
    const th = document.createElement('th')
    th.textContent = h
    th.style.backgroundColor = '#f2f2f2' // nền xám nhạt
    th.style.fontWeight = 'bold'
    th.style.minWidth = visibleColWidths[idx] + 'px'
    th.style.textAlign = 'center'
    thead.appendChild(th)
  })

  // filter căn giữa
  visibleHeaders.forEach((h, idx) => {
    const th = document.createElement('th')
    th.style.minWidth = visibleColWidths[idx] + 'px'
    th.style.textAlign = 'center'
    th.style.backgroundColor = '#f2f2f2' // nền xám nhạt

    const btn = document.createElement('button')
    btn.textContent = 'Filter'
    btn.onclick = (e) => {
      e.stopPropagation()
      th.innerHTML = '' // bỏ button đi

      // distinct values
      const values = visibleData.map((row) => row[idx])
      const distinct = [...new Set(values)].sort()

      // wrapper
      const comboWrapper = document.createElement('div')
      comboWrapper.style.position = 'relative'
      comboWrapper.style.width = '100%'

      // ô hiển thị (giả combobox)
      const display = document.createElement('div')
      display.textContent = '(Tất cả)'
      display.style.border = '1px solid #ccc'
      display.style.padding = '2px 24px 2px 6px'
      display.style.cursor = 'pointer'
      display.style.background = '#fff'
      display.style.borderRadius = '4px'
      display.style.whiteSpace = 'nowrap'
      display.style.overflow = 'hidden'
      display.style.textOverflow = 'ellipsis'
      display.style.position = 'relative'
      display.style.width = visibleColWidths[idx] + 'px' // giữ độ rộng cố định
      display.style.maxWidth = visibleColWidths[idx] + 'px'

      // icon mũi tên
      const arrow = document.createElement('span')
      arrow.innerHTML = '▼'
      arrow.style.position = 'absolute'
      arrow.style.right = '8px'
      arrow.style.top = '50%'
      arrow.style.transform = 'translateY(-50%)'
      arrow.style.fontSize = '10px'
      arrow.style.color = '#666'
      display.appendChild(arrow)

      // dropdown list
      const dropdown = document.createElement('div')
      dropdown.style.position = 'absolute'
      dropdown.style.top = '100%'
      dropdown.style.left = '0'
      dropdown.style.width = '100%'
      dropdown.style.border = '1px solid #ccc'
      dropdown.style.background = '#fff'
      dropdown.style.boxShadow = '0 2px 6px rgba(0,0,0,0.2)'
      dropdown.style.zIndex = '1000'
      dropdown.style.maxHeight = '200px'
      dropdown.style.overflowY = 'auto'
      dropdown.style.display = 'none'
      dropdown.style.textAlign = 'left'

      // search box trong dropdown
      const searchBox = document.createElement('input')
      searchBox.type = 'text'
      searchBox.placeholder = 'Tìm...'
      searchBox.style.width = '100%' // Căng toàn bộ chiều ngang
      searchBox.style.margin = '4px 0' // Chỉ giữ margin trên-dưới, bỏ trái-phải
      searchBox.style.padding = '2px 8px' // Padding trái-phải để chữ không sát mép
      searchBox.style.border = '1px solid #ccc'
      searchBox.style.position = 'sticky'
      searchBox.style.top = '0' // Sửa '1' thành '0' cho chuẩn vị trí sticky
      searchBox.style.background = '#ffb6c1' // Màu hồng phấn
      searchBox.style.zIndex = '1'
      searchBox.style.boxSizing = 'border-box' // Đảm bảo padding không làm vượt kích thước
      dropdown.appendChild(searchBox)

      // option: tất cả
      const allDiv = document.createElement('div')
      const allCb = document.createElement('input')
      allCb.type = 'checkbox'
      allCb.checked = true
      allCb.style.marginRight = '6px'
      const allLbl = document.createElement('span')
      allLbl.textContent = '(Tất cả)'
      allDiv.appendChild(allCb)
      allDiv.appendChild(allLbl)
      allDiv.style.position = 'sticky'
      allDiv.style.top = '21px' // Giữ nguyên nếu chiều cao searchBox không đổi
      allDiv.style.background = '#b0c4de' // Màu xanh xám
      allDiv.style.zIndex = '1'
      allDiv.style.padding = '4px 8px' // Padding trái-phải để chữ không sát mép, trên-dưới giữ nhỏ
      allDiv.style.width = '100%' // Căng toàn bộ chiều ngang
      allDiv.style.boxSizing = 'border-box' // Đảm bảo padding không làm vượt kích thước
      dropdown.appendChild(allDiv)

      const hr = document.createElement('hr')
      hr.style.margin = '0'
      hr.style.position = 'sticky'
      hr.style.top = '45px' // Khoảng sau searchBox (30px) + allDiv (30px)
      hr.style.zIndex = '1'
      dropdown.appendChild(hr)

      // options distinct
      distinct.forEach((v) => {
        const item = document.createElement('div')
        item.setAttribute('data-value', v)
        const cb = document.createElement('input')
        cb.type = 'checkbox'
        cb.value = v
        cb.checked = true
        cb.style.marginRight = '6px'
        const lbl = document.createElement('span')
        lbl.textContent = v
        item.appendChild(cb)
        item.appendChild(lbl)
        item.style.padding = '6px 8px' // Tăng padding để giãn dòng
        item.style.margin = '2px 0' // Thêm margin trên-dưới để các dòng cách nhau
        dropdown.appendChild(item)
      })

      // Lọc option theo search
      searchBox.addEventListener('input', () => {
        const keyword = normalizeUnicode(searchBox.value)
        dropdown.querySelectorAll('div[data-value]').forEach((item) => {
          const text = normalizeUnicode(item.innerText)
          item.style.display = text.includes(keyword) ? '' : 'none'
        })
      })

      // mở/đóng dropdown
      display.onclick = (ev) => {
        ev.stopPropagation()
        dropdown.style.display =
          dropdown.style.display === 'none' ? 'block' : 'none'

        adjustHeaderContainerWidth()
      }

      // áp dụng filter
      function applyFilter() {
        const selected = []
        dropdown.querySelectorAll('input[type=checkbox]').forEach((cb) => {
          if (cb.checked && cb.value) selected.push(cb.value)
        })

        let textShow
        if (allCb.checked || selected.length === 0) {
          textShow = '(Tất cả)'
        } else if (selected.length <= 2) {
          textShow = selected.join(', ')
        } else {
          textShow =
            selected.slice(0, 2).join(', ') + ` (+${selected.length - 2})`
        }

        // đổi text hiển thị
        display.childNodes[0].nodeValue = textShow

        // cập nhật filter cho cột này
        activeFilters[idx] = allCb.checked ? [] : selected

        // lọc bảng dựa trên tất cả filter
        tbody.querySelectorAll('tr').forEach((tr) => {
          let show = true
          for (const [colIdx, values] of Object.entries(activeFilters)) {
            if (values.length === 0) continue
            const idxNum = parseInt(colIdx, 10) // 👈 ép về số
            const cell = tr.children[idxNum]
            if (!cell) continue // tránh undefined
            const cellValue = cell.textContent
            if (!values.includes(cellValue)) {
              show = false
              break
            }
          }

          tr.style.display = show ? '' : 'none'
        })
      }

      // check/uncheck tất cả
      allCb.onchange = () => {
        const checked = allCb.checked
        dropdown.querySelectorAll('input[type=checkbox]').forEach((cb) => {
          if (cb !== allCb) cb.checked = checked
        })
        applyFilter()
      }

      // gắn cho từng checkbox con
      dropdown.querySelectorAll('input[type=checkbox]').forEach((cb) => {
        if (cb !== allCb) {
          cb.onchange = () => {
            // nếu tất cả con đều check thì tick lại "Tất cả"
            const allChildren = Array.from(
              dropdown.querySelectorAll('input[type=checkbox]')
            ).filter((x) => x !== allCb)
            allCb.checked = allChildren.every((x) => x.checked)
            applyFilter()
          }
        }
      })

      // click ra ngoài thì đóng dropdown
      document.addEventListener('click', function closeDropdown(ev) {
        if (!comboWrapper.contains(ev.target)) {
          dropdown.style.display = 'none'
        }
      })

      comboWrapper.appendChild(display)
      comboWrapper.appendChild(dropdown)
      th.appendChild(comboWrapper)

      adjustHeaderContainerWidth()
    }

    th.appendChild(btn)
    tfilter.appendChild(th)
  })

  // === Quản lý chọn dòng ===
  let lastSelectedIndex = null

  // Body
  visibleData.forEach((row, rowIndex) => {
    const tr = document.createElement('tr')
    row.forEach((cell, idx) => {
      const td = document.createElement('td')
      td.textContent = visibleIsMeasure[idx] ? formatNumber(cell) : cell
      td.style.minWidth = visibleColWidths[idx] + 'px'
      td.style.textAlign = visibleIsMeasure[idx] ? 'right' : 'left'
      tr.appendChild(td)
    })

    // Click chọn dòng
    tr.addEventListener('click', (e) => {
      if (e.shiftKey && lastSelectedIndex !== null) {
        // chọn range
        const trs = Array.from(tbody.querySelectorAll('tr'))
        const start = Math.min(lastSelectedIndex, rowIndex)
        const end = Math.max(lastSelectedIndex, rowIndex)
        for (let i = start; i <= end; i++) {
          trs[i].classList.add('selected-row')
        }
      } else if (e.ctrlKey || e.metaKey) {
        // toggle
        tr.classList.toggle('selected-row')
        lastSelectedIndex = rowIndex
      } else {
        // chỉ chọn 1
        tbody
          .querySelectorAll('tr')
          .forEach((tr2) => tr2.classList.remove('selected-row'))
        tr.classList.add('selected-row')
        lastSelectedIndex = rowIndex
      }
    })

    tbody.appendChild(tr)
  })

  // Copy khi Ctrl+C
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key.toLowerCase() === 'c') {
      const selected = tbody.querySelectorAll('.selected-row')
      if (selected.length > 0) {
        const text = Array.from(selected)
          .map((tr) => tr.innerText) // lấy toàn bộ nội dung dòng
          .join('\n')

        // --- Fallback cách cổ điển ---
        const textarea = document.createElement('textarea')
        textarea.value = text
        document.body.appendChild(textarea)
        textarea.select()
        try {
          document.execCommand('copy')
        } catch (err) {
          console.error('Copy thất bại:', err)
        }
        document.body.removeChild(textarea)
      }
    }
  })

  // === Dòng tổng cuối bảng ===
  const totals = []
  visibleIsMeasure.forEach((isM, idx) => {
    if (isM) {
      let sum = 0
      visibleData.forEach((r) => {
        const val = Number(r[idx].toString().replace(/,/g, ''))
        if (!isNaN(val)) sum += val
      })
      totals.push(sum)
    } else {
      totals.push('') // cột dimension thì để trống
    }
  })

  const totalRow = document.createElement('tr')
  totalRow.classList.add('total-row') // 👈 thêm dòng này
  totalRow.style.fontWeight = 'bold'
  totalRow.style.backgroundColor = '#f2f2f2' // nền xám nhạt
  totalRow.style.color = 'red' // ✅ Thêm màu chữ đỏ

  const dimCount = visibleIsMeasure.filter((v) => !v).length
  let firstDimHandled = false

  visibleIsMeasure.forEach((isM, idx) => {
    if (!isM && !firstDimHandled) {
      const td = document.createElement('td')
      td.textContent = 'Tổng cộng'
      td.colSpan = dimCount
      td.style.textAlign = 'left'
      totalRow.appendChild(td)
      firstDimHandled = true
    } else if (isM) {
      const td = document.createElement('td')
      td.textContent = formatNumber(totals[idx])
      td.style.textAlign = 'right'
      totalRow.appendChild(td)
    }
  })
  tbody.appendChild(totalRow)
}

// Pivot Measure Names/Values
function pivotMeasureValues(table) {
  const cols = table.columns.map((c) => c.fieldName)
  const rows = table.data.map((r) => r.map((c) => c.formattedValue))

  const measureNameIdx = cols.findIndex((c) =>
    c.toLowerCase().includes('measure names')
  )
  const measureValueIdx = cols.findIndex((c) =>
    c.toLowerCase().includes('measure values')
  )

  const dimensionIdxs = cols
    .map((c, i) => i)
    .filter((i) => i !== measureNameIdx && i !== measureValueIdx)

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

  const measureNames = Array.from(measureSet)
  const headers = [...dimensionIdxs.map((i) => cols[i]), ...measureNames]

  const isMeasure = [
    ...dimensionIdxs.map(() => false),
    ...measureNames.map(() => true)
  ]

  const data = Array.from(pivotMap.values()).map((entry) => {
    return [...entry.dims, ...measureNames.map((m) => entry.measures[m] || '')]
  })

  return { headers, data, isMeasure }
}

// Load lại dữ liệu và render
function loadAndRender(worksheet) {
  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
    const { headers, data, isMeasure } = pivotMeasureValues(sumData)

    const colWidths = headers.map((h, idx) => {
      const headerWidth = getTextWidth(h)
      const maxCellWidth = Math.max(
        ...data.map((r) => getTextWidth(r[idx] || ''))
      )
      const rawWidth = Math.max(headerWidth, maxCellWidth) + 20
      return Math.min(300, Math.max(30, rawWidth)) // giới hạn min = 30, max = 300
    })

    renderTable(headers, data, colWidths, isMeasure)

    // Sau khi render xong, lấy width thực của table
    const tableEl = document.getElementById('data-table')
    const headerContainer = document.querySelector('.header-container')
    if (tableEl && headerContainer) {
      const tableWidth = tableEl.offsetWidth
      headerContainer.style.width = tableWidth + 'px'
    }

    attachGlobalSearch()
  })
}

// Gắn search toàn cục
function attachGlobalSearch() {
  const searchInput = document.getElementById('global-search')
  if (!searchInput) return

  searchInput.addEventListener('input', () => {
    const keyword = normalizeUnicode(searchInput.value)
    const tbody = document.getElementById('table-body')

    tbody.querySelectorAll('tr').forEach((tr) => {
      const rowText = normalizeUnicode(tr.textContent)
      tr.style.display = rowText.includes(keyword) ? '' : 'none'
    })
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
      loadAndRender(worksheet)
    })

    tableau.extensions.dashboardContent.dashboard
      .getParametersAsync()
      .then(function (parameters) {
        parameters.forEach(function (p) {
          p.addEventListener(tableau.TableauEventType.ParameterChanged, () =>
            loadAndRender(worksheet)
          )
        })
      })
  })
})
