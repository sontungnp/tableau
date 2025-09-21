'use strict'

const activeFilters = {} // key: column index, value: array các giá trị được chọn

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

  // Header căn giữa
  headers.forEach((h, idx) => {
    const th = document.createElement('th')
    th.textContent = h
    th.style.backgroundColor = '#f2f2f2' // nền xám nhạt
    th.style.fontWeight = 'bold'
    th.style.minWidth = colWidths[idx] + 'px'
    th.style.textAlign = 'center'
    thead.appendChild(th)
  })

  // filter căn giữa
  headers.forEach((h, idx) => {
    const th = document.createElement('th')
    th.style.minWidth = colWidths[idx] + 'px'
    th.style.textAlign = 'center'
    th.style.backgroundColor = '#f2f2f2' // nền xám nhạt

    const btn = document.createElement('button')
    btn.textContent = 'Filter'
    btn.onclick = (e) => {
      e.stopPropagation()
      th.innerHTML = '' // bỏ button đi

      // distinct values
      const values = data.map((row) => row[idx])
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
      display.style.width = colWidths[idx] + 'px' // giữ độ rộng cố định
      display.style.maxWidth = colWidths[idx] + 'px'

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
      dropdown.appendChild(allDiv)
      dropdown.appendChild(document.createElement('hr'))

      // options distinct
      distinct.forEach((v) => {
        const item = document.createElement('div')
        const cb = document.createElement('input')
        cb.type = 'checkbox'
        cb.value = v
        cb.checked = true
        cb.style.marginRight = '6px'
        const lbl = document.createElement('span')
        lbl.textContent = v
        item.appendChild(cb)
        item.appendChild(lbl)
        dropdown.appendChild(item)
      })

      // mở/đóng dropdown
      display.onclick = (ev) => {
        ev.stopPropagation()
        dropdown.style.display =
          dropdown.style.display === 'none' ? 'block' : 'none'
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
    }

    th.appendChild(btn)
    tfilter.appendChild(th)
  })

  // === Quản lý chọn dòng ===
  let lastSelectedIndex = null

  // Body
  data.forEach((row, rowIndex) => {
    const tr = document.createElement('tr')
    row.forEach((cell, idx) => {
      const td = document.createElement('td')
      td.textContent = isMeasure[idx] ? formatNumber(cell) : cell
      td.style.minWidth = colWidths[idx] + 'px'
      td.style.textAlign = isMeasure[idx] ? 'right' : 'left'
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
  isMeasure.forEach((isM, idx) => {
    if (isM) {
      let sum = 0
      data.forEach((r) => {
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

  let firstDimHandled = false
  isMeasure.forEach((isM, idx) => {
    const td = document.createElement('td')
    if (!isM && !firstDimHandled) {
      td.textContent = 'Tổng cộng'
      td.colSpan = isMeasure.findIndex((v) => v) // gộp hết dimension
      td.style.textAlign = 'left'
      totalRow.appendChild(td)
      firstDimHandled = true
    }
    if (isM) {
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
      return Math.max(headerWidth, maxCellWidth) + 20
    })

    renderTable(headers, data, colWidths, isMeasure)
    attachGlobalSearch()
  })
}

// Gắn search toàn cục
function attachGlobalSearch() {
  const searchInput = document.getElementById('global-search')
  if (!searchInput) return

  searchInput.addEventListener('input', () => {
    const keyword = searchInput.value.toLowerCase()
    const tbody = document.getElementById('table-body')
    tbody.querySelectorAll('tr').forEach((tr) => {
      const rowText = tr.textContent.toLowerCase()
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
