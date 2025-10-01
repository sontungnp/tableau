'use strict'

const activeFilters = {} // key: column index, value: {mode: 'multi' | 'advanced', values: [] | {operator: '', value: ''}}

function debounce(fn, delay) {
  let timeout
  return (...args) => {
    clearTimeout(timeout)
    timeout = setTimeout(() => fn(...args), delay)
  }
}

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
const canvas = document.createElement('canvas')
const ctx = canvas.getContext('2d')
ctx.font = '14px Arial'
function getTextWidth(text) {
  return ctx.measureText(text).width
}

// Hàm format số
function formatNumber(value) {
  if (value === null || value === undefined || value === '') return ''
  const num = Number(value.toString().replace(/,/g, ''))
  if (isNaN(num)) return value
  return num.toLocaleString('en-US') // có thể đổi sang 'vi-VN'
}

// Hàm apply advanced filter cho một row và cột
function applyAdvancedFilter(cellValue, operator, filterValue, isMeasure) {
  const numValue = Number(cellValue.toString().replace(/,/g, ''))
  const strValue = normalizeUnicode(cellValue)

  if (isMeasure) {
    const numFilter = Number(filterValue)
    if (isNaN(numValue) || isNaN(numFilter)) return false
    switch (operator) {
      case '=':
        return numValue === numFilter
      case '!=':
        return numValue !== numFilter
      case '>':
        return numValue > numFilter
      case '<':
        return numValue < numFilter
      case '>=':
        return numValue >= numFilter
      case '<=':
        return numValue <= numFilter
      default:
        return true
    }
  } else {
    const strFilter = normalizeUnicode(filterValue)
    switch (operator) {
      case 'contains':
        return strValue.includes(strFilter)
      case 'startsWith':
        return strValue.startsWith(strFilter)
      case 'endsWith':
        return strValue.endsWith(strFilter)
      case '=':
        return strValue === strFilter
      case '!=':
        return strValue !== strFilter
      default:
        return true
    }
  }
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
    .filter(
      (item) =>
        item.header.toLowerCase().startsWith('hiden') ||
        item.header.startsWith('AGG')
    )
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
      comboWrapper.className = 'combo-wrapper' // ✅ Thêm dòng này
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
      searchBox.style.width = '100%'
      searchBox.style.margin = '4px 0'
      searchBox.style.padding = '2px 8px'
      searchBox.style.border = '1px solid #ccc'
      searchBox.style.position = 'sticky'
      searchBox.style.top = '0'
      searchBox.style.background = '#ffb6c1' // Màu hồng phấn
      searchBox.style.zIndex = '1'
      searchBox.style.boxSizing = 'border-box'
      dropdown.appendChild(searchBox)

      // ✅ THÊM NÚT NÂNG CAO (sticky sau searchBox)
      const advancedBtn = document.createElement('button')
      advancedBtn.textContent = 'Nâng cao'
      advancedBtn.style.display = 'block'
      advancedBtn.style.width = '100%'
      advancedBtn.style.padding = '4px 8px'
      advancedBtn.style.background = '#4CAF50'
      advancedBtn.style.color = 'white'
      advancedBtn.style.border = 'none'
      advancedBtn.style.borderRadius = '4px'
      advancedBtn.style.margin = '4px 0'
      advancedBtn.style.cursor = 'pointer'
      advancedBtn.style.fontSize = '12px'
      advancedBtn.style.position = 'sticky'
      advancedBtn.style.top = '21px'
      advancedBtn.style.zIndex = '1'
      dropdown.appendChild(advancedBtn)

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
      allDiv.style.top = '52px' // Điều chỉnh sau khi thêm advancedBtn (21px search + ~20px btn + margin)
      allDiv.style.background = '#b0c4de' // Màu xanh xám
      allDiv.style.zIndex = '1'
      allDiv.style.padding = '4px 8px'
      allDiv.style.width = '100%'
      allDiv.style.boxSizing = 'border-box'
      dropdown.appendChild(allDiv)

      const hr = document.createElement('hr')
      hr.style.margin = '0'
      hr.style.position = 'sticky'
      hr.style.top = '76px' // Điều chỉnh tương ứng
      hr.style.zIndex = '1'
      dropdown.appendChild(hr)

      // options distinct (ban đầu hiển thị nếu mode multi)
      const optionsContainer = document.createElement('div')
      optionsContainer.id = `options-${idx}` // Để toggle visibility
      distinct.forEach((v) => {
        const item = document.createElement('div')
        item.setAttribute('data-value', v)
        item.setAttribute('data-normalized', normalizeUnicode(v))
        const cb = document.createElement('input')
        cb.type = 'checkbox'
        cb.value = v

        // Khởi tạo activeFilters nếu chưa có
        if (!activeFilters[idx]) {
          activeFilters[idx] = { mode: 'multi', values: distinct.slice() }
        }
        const filter = activeFilters[idx]
        if (filter.mode === 'multi') {
          cb.checked = filter.values.includes(v)
        } else {
          cb.style.display = 'none' // Ẩn checkboxes nếu advanced mode
        }

        cb.style.marginRight = '6px'
        const lbl = document.createElement('span')
        lbl.textContent = v
        item.appendChild(cb)
        item.appendChild(lbl)
        item.style.padding = '6px 8px'
        item.style.margin = '2px 0'
        optionsContainer.appendChild(item)
      })
      dropdown.appendChild(optionsContainer)

      // Lọc option theo search (chỉ cho multi mode)
      searchBox.addEventListener(
        'input',
        debounce(() => {
          if (activeFilters[idx]?.mode !== 'multi') return
          const keyword = normalizeUnicode(searchBox.value)
          optionsContainer
            .querySelectorAll('div[data-value]')
            .forEach((item) => {
              const text = item.getAttribute('data-normalized')
              item.style.display = text.includes(keyword) ? '' : 'none'
            })
        }, 250)
      )

      // Mở advanced modal
      advancedBtn.onclick = (e) => {
        e.stopPropagation()
        openAdvancedModal(idx, visibleIsMeasure[idx], visibleData, distinct)
      }

      // Toggle dropdown visibility và mode
      function toggleOptionsVisibility() {
        const show = activeFilters[idx]?.mode === 'multi'
        optionsContainer.style.display = show ? 'block' : 'none'
        allDiv.style.display = show ? '' : 'none'
        hr.style.display = show ? '' : 'none'
        searchBox.style.display = show ? '' : 'none'
        advancedBtn.style.display = show ? '' : 'block' // Luôn show nút nâng cao
      }
      toggleOptionsVisibility() // Khởi tạo

      // mở/đóng dropdown
      display.onclick = (ev) => {
        ev.stopPropagation()
        dropdown.style.display = 'block'
        dropdown.classList.add('dropdown-open')

        adjustHeaderContainerWidth()
      }

      // áp dụng filter (cập nhật để handle advanced)
      function applyFilter() {
        const filterCols = Object.entries(activeFilters).filter(([_, f]) => {
          if (f.mode === 'multi') return f.values.length > 0
          return f.mode === 'advanced' && f.operator && f.value !== ''
        })

        // Reset tổng
        let totals = Array(visibleIsMeasure.length).fill(0)

        tbody.querySelectorAll('tr:not(.total-row)').forEach((tr, rowIndex) => {
          const row = visibleData[rowIndex]
          const show = filterCols.every(([colIdxStr, f]) => {
            const colIdx = parseInt(colIdxStr)
            if (f.mode === 'multi') {
              return f.values.includes(row[colIdx])
            } else if (f.mode === 'advanced') {
              return applyAdvancedFilter(
                row[colIdx],
                f.operator,
                f.value,
                visibleIsMeasure[colIdx]
              )
            }
            return true
          })
          tr.style.display = show ? '' : 'none'

          // Nếu dòng được hiển thị thì cộng vào tổng
          if (show) {
            row.forEach((cell, cIdx) => {
              if (visibleIsMeasure[cIdx]) {
                const val = Number(cell.toString().replace(/,/g, ''))
                if (!isNaN(val)) totals[cIdx] += val
              }
            })
          }
        })

        // Cập nhật lại dòng tổng
        const totalRow = tbody.querySelector('.total-row')
        if (totalRow) {
          totalRow.innerHTML = ''
          let firstDimHandled = false
          visibleIsMeasure.forEach((isM, cIdx) => {
            if (!isM && !firstDimHandled) {
              const td = document.createElement('td')
              td.textContent = 'Tổng cộng'
              td.colSpan = visibleIsMeasure.filter((v) => !v).length
              td.style.textAlign = 'left'
              totalRow.appendChild(td)
              firstDimHandled = true
            } else if (isM) {
              const td = document.createElement('td')
              td.textContent = formatNumber(totals[cIdx])
              td.style.textAlign = 'right'
              totalRow.appendChild(td)
            }
          })
        }

        // Cập nhật display label
        updateDisplayLabel()
      }

      // Cập nhật label hiển thị (như code cũ cho multi, cải thiện advanced với icon + tooltip)
      function updateDisplayLabel() {
        const filter = activeFilters[idx]
        if (filter.mode === 'multi') {
          const selected = filter.values
          if (selected.length === distinct.length) {
            display.textContent = '(Tất cả)'
          } else if (selected.length === 0) {
            display.textContent = '(Trống)'
          } else {
            // Giữ nguyên code cũ: join bằng comma, CSS sẽ tự ellipsis nếu dài
            display.textContent = selected.join(', ')
          }
        } else {
          // ✅ Cải thiện advanced: Thêm icon ⚙️, text ngắn gọn, và tooltip full info
          // ✅ Fix: Thêm quote cho tất cả keys string
          const operatorLabel =
            {
              '=': '=',
              '!=': '≠',
              '>': '>',
              '<': '<',
              '>=': '≥',
              '<=': '≤',
              contains: 'Chứa',
              startsWith: 'Bắt đầu',
              endsWith: 'Kết thúc'
            }[filter.operator] || filter.operator // Symbol ngắn gọn cho operator
          display.textContent = `⚙️ ${operatorLabel} ${filter.value}` // Icon + short label
          // Tooltip: Hiển thị full khi hover (không cần mở dropdown)
          const colType = visibleIsMeasure[idx] ? 'số' : 'chuỗi'
          display.title = `Advanced Filter đang áp dụng: ${filter.operator} ${filter.value} (cột ${colType})`
        }
        display.appendChild(arrow)
      }

      // check/uncheck tất cả (chỉ cho multi)
      allCb.onchange = () => {
        if (activeFilters[idx]?.mode !== 'multi') return
        const checked = allCb.checked
        optionsContainer
          .querySelectorAll('input[type=checkbox]')
          .forEach((cb) => {
            if (cb !== allCb) cb.checked = checked
          })
        activeFilters[idx].values = checked ? distinct.slice() : []
        updateDisplayLabel()
        applyFilter()
      }

      // gắn cho từng checkbox con (chỉ cho multi)
      optionsContainer
        .querySelectorAll('input[type=checkbox]')
        .forEach((cb) => {
          if (cb !== allCb) {
            cb.onchange = () => {
              if (activeFilters[idx]?.mode !== 'multi') return
              const allChildren = Array.from(
                optionsContainer.querySelectorAll('input[type=checkbox]')
              ).filter((x) => x !== allCb)
              allCb.checked = allChildren.every((x) => x.checked)
              const selected = allChildren
                .filter((x) => x.checked)
                .map((x) => x.value)
              activeFilters[idx].values = selected
              updateDisplayLabel()
              applyFilter()
            }
          }
        })

      // click ra ngoài thì đóng dropdown
      document.addEventListener('click', function closeDropdown(ev) {
        if (!comboWrapper.contains(ev.target)) {
          dropdown.style.display = 'none'
          dropdown.classList.remove('dropdown-open')
        }
      })

      comboWrapper.appendChild(display)
      comboWrapper.appendChild(dropdown)
      th.appendChild(comboWrapper)

      adjustHeaderContainerWidth()
      applyFilter() // Áp dụng initial filter
    }

    th.appendChild(btn)
    tfilter.appendChild(th)
  })

  // === Quản lý chọn dòng ===
  let lastSelectedIndex = null

  // Body
  const fragment = document.createDocumentFragment()
  visibleData.forEach((row, rowIndex) => {
    const tr = document.createElement('tr')
    tr.innerHTML = row
      .map((cell, idx) => {
        const align = visibleIsMeasure[idx] ? 'right' : 'left'
        const safeCell =
          !cell ||
          (typeof cell === 'string' && cell.trim().toLowerCase() === 'null')
            ? ''
            : cell

        const content = visibleIsMeasure[idx]
          ? formatNumber(safeCell)
          : safeCell

        return `<td style="min-width:${visibleColWidths[idx]}px;text-align:${align}">${content}</td>`
      })
      .join('')

    // ✅ Gắn event để highlight dòng khi chọn
    tr.addEventListener('click', (e) => {
      if (e.ctrlKey) {
        // Multi-select với Ctrl
        tr.classList.toggle('row-selected')
      } else if (e.shiftKey && lastSelectedIndex !== null) {
        // Chọn nhiều dòng liên tục với Shift
        const start = Math.min(lastSelectedIndex, rowIndex)
        const end = Math.max(lastSelectedIndex, rowIndex)
        tbody.querySelectorAll('tr').forEach((r, i) => {
          if (i >= start && i <= end) {
            r.classList.add('row-selected')
          }
        })
      } else {
        // Chọn 1 dòng
        tbody
          .querySelectorAll('tr')
          .forEach((r) => r.classList.remove('row-selected'))
        tr.classList.add('row-selected')
      }
      lastSelectedIndex = rowIndex
    })

    fragment.appendChild(tr)
  })
  tbody.appendChild(fragment)

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

// Modal cho advanced filter (global)
let advancedModal = null
function createAdvancedModal() {
  if (advancedModal) return advancedModal

  advancedModal = document.createElement('dialog')
  advancedModal.style.position = 'fixed'
  advancedModal.style.top = '50%'
  advancedModal.style.left = '50%'
  advancedModal.style.transform = 'translate(-50%, -50%)'
  advancedModal.style.width = '300px'
  advancedModal.style.padding = '20px'
  advancedModal.style.border = 'none'
  advancedModal.style.borderRadius = '8px'
  advancedModal.style.boxShadow = '0 4px 12px rgba(0,0,0,0.3)'
  advancedModal.style.background = '#fff'
  advancedModal.innerHTML = `
    <h3>Tìm kiếm nâng cao</h3>
    <label>Điều kiện:</label>
    <select id="adv-operator"></select>
    <input type="text" id="adv-value" placeholder="Nhập giá trị..." style="width:100%; margin:10px 0; padding:5px;">
    <div style="text-align:right;">
      <button id="adv-apply" style="margin-right:10px;">Áp dụng</button>
      <button id="adv-clear">Xóa</button>
    </div>
  `
  document.body.appendChild(advancedModal)

  // Close khi click outside (polyfill cho dialog cũ)
  advancedModal.addEventListener('click', (e) => {
    if (e.target === advancedModal) advancedModal.close()
  })

  return advancedModal
}

function openAdvancedModal(colIdx, isMeasure, visibleData, distinct) {
  const modal = createAdvancedModal()
  const operatorSelect = modal.querySelector('#adv-operator')
  const valueInput = modal.querySelector('#adv-value')
  const applyBtn = modal.querySelector('#adv-apply')
  const clearBtn = modal.querySelector('#adv-clear')

  // Operators tùy loại cột
  const operators = isMeasure
    ? [
        { value: '=', label: 'Bằng (=)' },
        { value: '!=', label: 'Không bằng (!=)' },
        { value: '>', label: 'Lớn hơn (>)' },
        { value: '<', label: 'Nhỏ hơn (<)' },
        { value: '>=', label: 'Lớn hơn hoặc bằng (>=-)' },
        { value: '<=', label: 'Nhỏ hơn hoặc bằng (<=)' }
      ]
    : [
        { value: 'contains', label: 'Chứa (contains)' },
        { value: 'startsWith', label: 'Bắt đầu bằng (startsWith)' },
        { value: 'endsWith', label: 'Kết thúc bằng (endsWith)' },
        { value: '=', label: 'Bằng (=)' },
        { value: '!=', label: 'Không bằng (!=)' }
      ]

  operatorSelect.innerHTML = operators
    .map((op) => `<option value="${op.value}">${op.label}</option>`)
    .join('')

  // Load current filter
  const currentFilter = activeFilters[colIdx]
  if (currentFilter?.mode === 'advanced') {
    operatorSelect.value = currentFilter.operator
    valueInput.value = currentFilter.value
  } else {
    // Default: contains hoặc = cho string/number
    operatorSelect.value = isMeasure ? '=' : 'contains'
    valueInput.value = ''
  }

  // Validate input cho number
  valueInput.oninput = () => {
    if (isMeasure && valueInput.value && isNaN(Number(valueInput.value))) {
      valueInput.style.borderColor = 'red'
    } else {
      valueInput.style.borderColor = '#ccc'
    }
  }

  // Apply
  applyBtn.onclick = () => {
    const operator = operatorSelect.value
    const value = valueInput.value.trim()
    if (!value) {
      alert('Vui lòng nhập giá trị!')
      return
    }
    if (isMeasure && isNaN(Number(value))) {
      alert('Giá trị phải là số!')
      return
    }

    activeFilters[colIdx] = { mode: 'advanced', operator, value }
    applyFilterGlobal(colIdx) // Truyền colIdx để update label
    modal.close()
  }

  // Clear
  clearBtn.onclick = () => {
    delete activeFilters[colIdx]
    activeFilters[colIdx] = { mode: 'multi', values: distinct.slice() } // Reset về multi all
    applyFilterGlobal(colIdx) // Truyền colIdx để update label
    modal.close()
  }

  modal.showModal()
}

// Global applyFilter (gọi từ modal) – Fix: Update label thủ công cho changed col
function applyFilterGlobal(changedColIdx = null) {
  if (window.currentVisibleData && window.currentVisibleIsMeasure) {
    const tbody = document.getElementById('table-body')
    if (!tbody) return

    const filterCols = Object.entries(activeFilters).filter(([_, f]) => {
      if (f.mode === 'multi') return f.values.length > 0
      return f.mode === 'advanced' && f.operator && f.value !== ''
    })

    let totals = Array(window.currentVisibleIsMeasure.length).fill(0)

    tbody.querySelectorAll('tr:not(.total-row)').forEach((tr, rowIndex) => {
      const row = window.currentVisibleData[rowIndex]
      const show = filterCols.every(([colIdxStr, f]) => {
        const colIdx = parseInt(colIdxStr)
        if (f.mode === 'multi') {
          return f.values.includes(row[colIdx])
        } else if (f.mode === 'advanced') {
          return applyAdvancedFilter(
            row[colIdx],
            f.operator,
            f.value,
            window.currentVisibleIsMeasure[colIdx]
          )
        }
        return true
      })
      tr.style.display = show ? '' : 'none'

      if (show) {
        row.forEach((cell, cIdx) => {
          if (window.currentVisibleIsMeasure[cIdx]) {
            const val = Number(cell.toString().replace(/,/g, ''))
            if (!isNaN(val)) totals[cIdx] += val
          }
        })
      }
    })

    // Update total row (giữ nguyên)
    const totalRow = tbody.querySelector('.total-row')
    if (totalRow) {
      totalRow.innerHTML = ''
      let firstDimHandled = false
      window.currentVisibleIsMeasure.forEach((isM, cIdx) => {
        if (!isM && !firstDimHandled) {
          const td = document.createElement('td')
          td.textContent = 'Tổng cộng'
          td.colSpan = window.currentVisibleIsMeasure.filter((v) => !v).length
          td.style.textAlign = 'left'
          totalRow.appendChild(td)
          firstDimHandled = true
        } else if (isM) {
          const td = document.createElement('td')
          td.textContent = formatNumber(totals[cIdx])
          td.style.textAlign = 'right'
          totalRow.appendChild(td)
        }
      })
    }

    // ✅ Fix: Update label cho cột thay đổi (thủ công, không hack click)
    if (changedColIdx !== null && activeFilters[changedColIdx]) {
      const filter = activeFilters[changedColIdx]
      const th = document.querySelector(
        `#table-filter th:nth-child(${changedColIdx + 1})`
      ) // nth-child(1) cho col 0
      if (th) {
        // Tìm display: th > comboWrapper (div) > display (div đầu tiên)
        const display = th.querySelector('div > div') // Selector ổn định dựa trên structure
        if (display) {
          const arrow = display.querySelector('span') // Giữ arrow nếu có
          if (filter.mode === 'multi') {
            // Cho multi: Cần distinct (lấy từ local, nhưng vì global apply, dùng length so sánh với total rows nếu approx)
            // Để đơn giản (vì multi không gọi global), giữ "(Tất cả)" tạm hoặc skip – nhưng advanced là focus
            display.textContent = '(Tất cả)' // Fallback, hoặc reload nếu cần
          } else {
            // Advanced: Compute như updateDisplayLabel
            const operatorMap = {
              '=': '=',
              '!=': '≠',
              '>': '>',
              '<': '<',
              '>=': '≥',
              '<=': '≤',
              contains: 'Chứa',
              startsWith: 'Bắt đầu',
              endsWith: 'Kết thúc'
            }
            const operatorLabel =
              operatorMap[filter.operator] || filter.operator
            const newText = `⚙️ ${operatorLabel} ${filter.value}`
            display.textContent = newText
            // Tooltip
            const colType = window.currentVisibleIsMeasure[changedColIdx]
              ? 'số'
              : 'chuỗi'
            display.title = `Advanced Filter đang áp dụng: ${filter.operator} ${filter.value} (cột ${colType})`
          }
          if (arrow) display.appendChild(arrow) // Đảm bảo arrow ở cuối
        }
      }
    }
  } else {
    console.error('Global data not loaded!') // Debug
  }
}

// Pivot Measure Names/Values
function pivotMeasureValues(table) {
  console.log('table.columns', table.columns)

  const cols = table.columns.map((c) => c.fieldName)
  const rows = table.data.map((r) =>
    r.map((c) =>
      c.formattedValue === null || c.formattedValue === undefined
        ? ''
        : c.formattedValue
    )
  )

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

    // Lưu global cho applyFilterGlobal
    const columnsToHide = headers
      .map((header, index) => ({ header, index }))
      .filter(
        (item) =>
          item.header.toLowerCase().startsWith('hiden') ||
          item.header.startsWith('AGG')
      )
      .map((item) => item.index)

    window.currentVisibleData = data.map((row) =>
      row.filter((cell, index) => !columnsToHide.includes(index))
    ) // Sử dụng window để global
    window.currentVisibleIsMeasure = isMeasure.filter(
      (measure, index) => !columnsToHide.includes(index)
    )

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

  // Copy khi Ctrl+C
  document.addEventListener('keydown', (e) => {
    if (e.ctrlKey && e.key.toLowerCase() === 'c') {
      const tbody = document.getElementById('table-body') // 👈 thêm dòng này
      const selected = tbody.querySelectorAll('.row-selected')
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

  document.addEventListener('click', (ev) => {
    document.querySelectorAll('.dropdown-open').forEach((dd) => {
      if (!dd.contains(ev.target)) {
        dd.style.display = 'none'
        dd.classList.remove('dropdown-open')
      }
    })
  })

  // === Export Excel ===
  document.getElementById('btn-export').addEventListener('click', () => {
    const table = document.getElementById('data-table')
    if (!table) {
      alert('Không tìm thấy bảng dữ liệu để export!')
      return
    }

    // Lấy header (chỉ dòng header chính)
    const headerCells = Array.from(
      table.querySelectorAll('thead tr#table-header th')
    )
    const columnsCount = headerCells.length
    const headers = headerCells.map((th) => th.innerText.trim())
    const rows = [headers]

    // Helper: chuyển text -> number nếu có thể (loại bỏ dấu phẩy)
    function parseCellText(txt) {
      const s = (txt || '').toString().trim()
      if (s === '') return ''
      // loại bỏ dấu phẩy/space
      const n = s.replace(/,/g, '').replace(/\s+/g, '')
      if (!isNaN(Number(n))) return Number(n)
      return s
    }

    const tbodyRows = Array.from(table.querySelectorAll('tbody tr'))
    let totalRowEl = null

    // Duyệt body, skip .total-row (xử lý sau), chỉ lấy các dòng đang hiển thị
    tbodyRows.forEach((tr) => {
      if (tr.classList.contains('total-row')) {
        totalRowEl = tr
        return
      }
      if (tr.style.display === 'none') return // filter đã ẩn -> bỏ
      const tds = Array.from(tr.querySelectorAll('td'))
      const values = tds.map((td) => parseCellText(td.innerText))

      // đảm bảo độ dài bằng columnsCount (pad/truncate nếu cần)
      if (values.length < columnsCount) {
        while (values.length < columnsCount) values.push('')
      } else if (values.length > columnsCount) {
        values.length = columnsCount
      }
      rows.push(values)
    })

    // Xử lý total-row (nếu có): mở rộng colspan để khớp column count
    if (totalRowEl) {
      const tds = Array.from(totalRowEl.querySelectorAll('td'))
      if (tds.length > 0) {
        const firstTd = tds[0]
        const colspanAttr = firstTd.getAttribute('colspan')
        const colspan = colspanAttr ? parseInt(colspanAttr, 10) || 1 : 1

        const totalCells = []
        // đặt nội dung ô đầu tiên, sau đó thêm (colspan-1) ô rỗng để "giả" colspan
        totalCells.push(firstTd.innerText.trim())
        for (let i = 1; i < colspan; i++) totalCells.push('')

        // phần còn lại là các ô measures
        for (let i = 1; i < tds.length; i++) {
          totalCells.push(parseCellText(tds[i].innerText))
        }

        // pad/truncate để đạt đúng columnsCount
        if (totalCells.length < columnsCount) {
          while (totalCells.length < columnsCount) totalCells.push('')
        } else if (totalCells.length > columnsCount) {
          totalCells.length = columnsCount
        }

        // Cuối cùng push vào rows (luôn ở cuối giống trên web)
        rows.push(totalCells)
      }
    }

    // Tạo workbook và sheet từ array-of-arrays
    const wb = XLSX.utils.book_new()
    const ws = XLSX.utils.aoa_to_sheet(rows)

    // Optional: set column widths dựa trên header hiển thị (wpx)
    try {
      ws['!cols'] = headerCells.map((th) => ({ wpx: th.offsetWidth || 80 }))
    } catch (e) {
      // ignore if offsetWidth không khả dụng
    }

    XLSX.utils.book_append_sheet(wb, ws, 'Data')
    XLSX.writeFile(wb, 'export.xlsx')
  })
})
