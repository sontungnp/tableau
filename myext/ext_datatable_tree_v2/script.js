'use strict'

// ======================
// CONSTANTS & CONFIG
// ======================
const OPERATORS = {
  divide: (a, b) => (b !== 0 ? a / b : 0),
  ratio_complement: (a, b) => (b !== 0 ? 1 - a / b : 0)
}

const FORMATTERS = {
  percent:
    (opts = {}) =>
    (params) => {
      const v = Number(params.value)
      if (isNaN(v)) return params.value ?? ''
      const precision = opts.precision ?? 2
      return `${(v * 100).toLocaleString('en-US', {
        minimumFractionDigits: precision,
        maximumFractionDigits: precision
      })} %`
    },
  currency:
    (opts = {}) =>
    (params) => {
      const v = Number(params.value)
      if (isNaN(v)) return params.value ?? ''
      return v.toLocaleString('en-US', {
        style: 'currency',
        currency: opts.currency || 'VND',
        maximumFractionDigits: opts.precision ?? 0
      })
    },
  number:
    (opts = {}) =>
    (params) => {
      const v = Number(params.value)
      if (isNaN(v)) return params.value ?? ''
      const precision = opts.precision ?? 2
      return v.toLocaleString('en-US', {
        minimumFractionDigits: precision,
        maximumFractionDigits: precision
      })
    }
}

// ======================
// GLOBAL STATE
// ======================
let state = {
  list_column_horizontal: null,
  list_column_vertical: null,
  list_column_measure: null,
  list_exclude_column_config: [],
  order_lv1: null,
  order_lv2: null,
  leaf_in_tree: null,
  dimensionColumns: [],
  measureColumns: [],
  pivotDataOutput: null,
  pivot_column_config: [],
  formated_columns: null,
  agGridColumnDefs: null,
  agGridColumnDefs_flat: null,
  selectedCellValue: null,
  extractRefreshTime: '',
  gridApi: null,
  nestedData: [],
  currentExpandedLevel: 1,
  maxTreeLevel: 1,
  levelSortRules: null,
  showGrandTotal: 1,
  percent_columns: [],
  expandListenersBound: false,
  columnFormatMatchers: [],
  font_size: '15px',
  row_height: 28,
  fit_content: 0,
  fit_window: 0,
  grandTotalPosition: 'bottom', // 'top' hoặc 'bottom', mặc định 'bottom'
  toolbar_position: 'top'
}

// ======================
// UTILITY FUNCTIONS
// ======================
const parseOrder = (order) => {
  if (order === 'asc' || order === 'desc') return { type: order }
  const list = Array.isArray(order)
    ? order
    : String(order)
        .split(',')
        .map((v) => v.trim())
  return { type: 'custom', list }
}

const normalizeUnicode = (str) =>
  str ? str.normalize('NFC').toLowerCase().trim() : ''

const extractTreeHeaderName = (header) => {
  const cleanHeader = header.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
  const match = cleanHeader.match(/tree_lv\d+\s*\(\s*(.+?)\s*\)/)
  return match ? match[1].trim() : 'Cấu trúc cây'
}

const getMaxTreeLevel = (nodes) => {
  let max = 1
  const walk = (list) => {
    for (const n of list) {
      max = Math.max(max, n.level)
      if (n.children) walk(n.children)
    }
  }
  walk(nodes)
  return max
}

// ======================
// SORT FUNCTIONS
// ======================
function sortColumns(list_columns, order_lv1, order_lv2) {
  const lv1Order = parseOrder(order_lv1)
  const lv2Order = parseOrder(order_lv2)

  const parsed = list_columns.map((c) => {
    const [lv1, lv2] = c.split('_')
    return { raw: c, lv1, lv2 }
  })

  const groupMap = new Map()
  parsed.forEach((item) => {
    if (!groupMap.has(item.lv1)) groupMap.set(item.lv1, [])
    groupMap.get(item.lv1).push(item)
  })

  let lv1Keys = Array.from(groupMap.keys())
  if (lv1Order.type === 'asc') lv1Keys.sort((a, b) => a.localeCompare(b))
  else if (lv1Order.type === 'desc') lv1Keys.sort((a, b) => b.localeCompare(a))
  else {
    const idx = new Map(lv1Order.list.map((v, i) => [v, i]))
    lv1Keys.sort((a, b) => (idx.get(a) ?? 9999) - (idx.get(b) ?? 9999))
  }

  const result = []
  lv1Keys.forEach((lv1) => {
    const rows = groupMap.get(lv1)
    if (lv2Order.type === 'asc') rows.sort((a, b) => a.lv2.localeCompare(b.lv2))
    else if (lv2Order.type === 'desc')
      rows.sort((a, b) => b.lv2.localeCompare(a.lv2))
    else {
      const idx = new Map(lv2Order.list.map((v, i) => [v.padStart(2, '0'), i]))
      rows.sort((a, b) => (idx.get(a.lv2) ?? 9999) - (idx.get(b.lv2) ?? 9999))
    }
    result.push(...rows.map((r) => r.raw))
  })
  return result
}

function customValueComparator(orderList) {
  const orderMap = new Map(orderList.map((v, i) => [v, i]))
  return (a, b) =>
    (orderMap.has(a.name) ? orderMap.get(a.name) : Infinity) -
    (orderMap.has(b.name) ? orderMap.get(b.name) : Infinity)
}

// ======================
// COLUMN DEFS
// ======================
function initColumnFormatMatchers(formated_columns) {
  const matchers = []
  if (!formated_columns) return matchers

  try {
    const formats =
      typeof formated_columns === 'string'
        ? JSON.parse(formated_columns)
        : formated_columns
    formats.forEach((f) => {
      if (!f.field || !f.formatType) return
      const pattern = f.field
      let matcher
      if (pattern.startsWith('%') && pattern.endsWith('%')) {
        const key = pattern.slice(1, -1)
        matcher = (field) => field.includes(key)
      } else if (pattern.startsWith('%')) {
        const key = pattern.slice(1)
        matcher = (field) => field.endsWith(key)
      } else if (pattern.endsWith('%')) {
        const key = pattern.slice(0, -1)
        matcher = (field) => field.startsWith(key)
      } else {
        matcher = (field) => field === pattern
      }
      matchers.push({ matcher, config: f })
    })
  } catch (e) {
    console.error('Invalid formated_columns JSON', e)
  }
  return matchers
}

function resolveFormatedColumns(field, matchers) {
  let matched = null,
    priority = -1
  matchers.forEach(({ matcher, config }) => {
    if (!matcher(field)) return
    let p = 0
    const pattern = config.field
    if (!pattern.includes('%')) p = 4
    else if (pattern.startsWith('%') && pattern.endsWith('%')) p = 1
    else if (pattern.startsWith('%') || pattern.endsWith('%')) p = 2
    if (p > priority) {
      priority = p
      matched = config
    }
  })
  return matched
}

function getFixedWidthColumns() {
  const fixedWidthFields = new Set()

  // 1️⃣ Xử lý từ pivot_column_config (exact match)
  if (state.pivot_column_config) {
    state.pivot_column_config.forEach((col) => {
      if (
        col.field &&
        col.width &&
        col.width !== 'auto' &&
        col.width !== undefined
      ) {
        fixedWidthFields.add(col.field)
      }
    })
  }

  // 2️⃣ Xử lý từ formated_columns (hỗ trợ wildcard %)
  if (state.formated_columns) {
    try {
      const formats =
        typeof state.formated_columns === 'string'
          ? JSON.parse(state.formated_columns)
          : state.formated_columns

      formats.forEach((f) => {
        // Nếu có width và không phải auto
        if (f.width && f.width !== 'auto' && f.width !== undefined && f.field) {
          const pattern = f.field

          // Lưu pattern để sau này match với field thực tế
          if (pattern.startsWith('%') && pattern.endsWith('%')) {
            // %abc% → contains
            const key = pattern.slice(1, -1)
            fixedWidthFields.add({ type: 'contains', key, pattern })
          } else if (pattern.startsWith('%')) {
            // %abc → endsWith
            const key = pattern.slice(1)
            fixedWidthFields.add({ type: 'endsWith', key, pattern })
          } else if (pattern.endsWith('%')) {
            // abc% → startsWith
            const key = pattern.slice(0, -1)
            fixedWidthFields.add({ type: 'startsWith', key, pattern })
          } else {
            // exact match
            fixedWidthFields.add({ type: 'exact', key: pattern, pattern })
          }
        }
      })
    } catch (e) {
      console.error('Invalid formated_columns JSON', e)
    }
  }

  return fixedWidthFields
}

// Hàm kiểm tra field có bị fixed width không
function isFieldFixedWidth(field, fixedWidthColumns) {
  for (const rule of fixedWidthColumns) {
    // Nếu rule là string (từ pivot_column_config)
    if (typeof rule === 'string') {
      if (field === rule) return true
      continue
    }

    // Nếu rule là object (từ formated_columns với wildcard)
    switch (rule.type) {
      case 'contains':
        if (field.includes(rule.key)) return true
        break
      case 'startsWith':
        if (field.startsWith(rule.key)) return true
        break
      case 'endsWith':
        if (field.endsWith(rule.key)) return true
        break
      case 'exact':
        if (field === rule.key) return true
        break
    }
  }
  return false
}

function applyColumnSizing() {
  // 1. Auto size tất cả columns để Fit nội dung trước
  // params.api.autoSizeAllColumns()

  // 2. Co giãn các cột để vừa với grid width, Đảm bảo tổng width không vượt quá grid
  // params.api.sizeColumnsToFit()

  // 3. Lắng nghe resize window để fit lại
  // window.addEventListener('resize', () => {
  //   params.api.sizeColumnsToFit()
  // })

  // ==========================================

  // ✅ Lấy danh sách cột có width cố định từ cả 2 nguồn
  // const fixedWidthColumns = getFixedWidthColumns()
  // console.log('fixedWidthColumns: ', fixedWidthColumns)

  // ✅ Lấy tất cả cột hiển thị
  // const allColumns = params.api.getColumns()
  // const visibleColumns = allColumns.filter((col) => col.isVisible())

  // ✅ Phân loại cột cần auto fit
  // const columnsToAutoFit = visibleColumns.filter((col) => {
  //   const colId = col.getColId()
  //   return !isFieldFixedWidth(colId, fixedWidthColumns)
  // })

  // console.log('columnsToAutoFit: ', columnsToAutoFit)

  // ✅ Auto fit các cột không có width cố định
  // if (columnsToAutoFit.length) {
  //   params.api.autoSizeColumns(columnsToAutoFit)
  // }

  // ====================================

  if (!state.gridApi) return

  const fixedWidthColumns = getFixedWidthColumns()

  // 1️⃣ Lấy tất cả cột hiển thị
  const allColumns = state.gridApi.getColumns()
  const visibleColumns = allColumns.filter((col) => col.isVisible())

  // 2️⃣ Phân loại cột
  const fixedColumns = [] // Cột có width cố định
  const autoFitColumns = [] // Cột cần auto fit

  visibleColumns.forEach((col) => {
    const colId = col.getColId()
    if (isFieldFixedWidth(colId, fixedWidthColumns)) {
      fixedColumns.push(col)
    } else {
      autoFitColumns.push(col)
    }
  })

  console.log('fixedColumns: ', fixedColumns)
  console.log('autoFitColumns: ', autoFitColumns)

  // cấu hình để fit content các cột
  if (state.fit_content == 1) {
    console.log('Vao fit_content', state.fit_content)
    // ✅ Auto fit các cột không có width cố định
    if (autoFitColumns.length) {
      state.gridApi.autoSizeColumns(autoFitColumns)
    }
  }

  // Cấu hình để fit window
  if (state.fit_window == 1) {
    console.log('Vao fit_window', state.fit_window)

    // 3️⃣ Tạm thời khóa các cột cố định trước khi fit
    const originalSuppressSizeToFit = {}
    fixedColumns.forEach((col) => {
      const colDef = col.getColDef()
      originalSuppressSizeToFit[col.getId()] = colDef.suppressSizeToFit
      colDef.suppressSizeToFit = true // Ngăn sizeColumnsToFit thay đổi cột này
    })

    // 4️⃣ Auto fit nội dung cho các cột không cố định
    if (autoFitColumns.length) {
      state.gridApi.autoSizeColumns(autoFitColumns)
    }

    // 5️⃣ Fit tổng thể vào grid, nhưng chỉ ảnh hưởng cột không bị khóa
    state.gridApi.sizeColumnsToFit()

    // 6️⃣ Khôi phục lại suppressSizeToFit
    fixedColumns.forEach((col) => {
      const colDef = col.getColDef()
      colDef.suppressSizeToFit = originalSuppressSizeToFit[col.getId()] || false
    })
  }

  // ====================================
}

function createColumnDefs(
  dimensionColumns,
  measureColumns,
  customConfig,
  excludeColumns = [],
  formated_columns
) {
  const columnFormatMatchers = initColumnFormatMatchers(formated_columns)

  const filteredDimensionColumns = dimensionColumns.filter(
    (field) => !field.startsWith('tree_lv') && !excludeColumns.includes(field)
  )

  const filteredMeasureColumns = sortColumns(
    measureColumns.filter((field) => !excludeColumns.includes(field)),
    state.order_lv1,
    state.order_lv2
  )

  const configMap = new Map()
  const orderedFieldsFromConfig = []
  customConfig.forEach((col) => {
    if (excludeColumns.includes(col.field)) return
    if (
      col.field !== 'name' &&
      !filteredDimensionColumns.includes(col.field) &&
      !filteredMeasureColumns.includes(col.field)
    )
      return
    configMap.set(col.field, col)
    orderedFieldsFromConfig.push(col.field)
  })

  const mergedFields = [
    ...new Set([
      'name',
      ...orderedFieldsFromConfig,
      ...filteredDimensionColumns,
      ...filteredMeasureColumns
    ])
  ].filter((f) => !excludeColumns.includes(f))

  const columnDefs = mergedFields.map((field) => {
    const customProps = configMap.get(field)
    const isDimensionOrName =
      filteredDimensionColumns.includes(field) || field === 'name'
    const isMeasure = filteredMeasureColumns.includes(field)

    let columnDef = {
      field,
      headerName: field,
      width: isDimensionOrName ? 250 : 200,
      wrapHeaderText: true,
      autoHeaderHeight: true
    }

    if (isDimensionOrName) columnDef.cellStyle = { textAlign: 'left' }
    else if (isMeasure) {
      columnDef.cellStyle = { textAlign: 'right' }
      columnDef.valueFormatter = (params) => {
        const v = params.value
        if (v == null || v === '') return ''
        const num = Number(v)
        return isNaN(num)
          ? v
          : num.toLocaleString('en-US', { maximumFractionDigits: 2 })
      }
    }

    if (field === 'name') {
      columnDef = {
        ...columnDef,
        cellRenderer: nameCellRenderer,
        headerName: customProps?.headerName || 'Tên'
      }
    }

    if (customProps) columnDef = { ...columnDef, ...customProps }

    const formatConfig = resolveFormatedColumns(field, columnFormatMatchers)
    if (formatConfig && FORMATTERS[formatConfig.formatType]) {
      columnDef.valueFormatter =
        FORMATTERS[formatConfig.formatType](formatConfig)
      columnDef.type = 'numericColumn'
      columnDef.cellStyle = { textAlign: 'right' }
      if (formatConfig.width) columnDef.width = formatConfig.width
      if (formatConfig.minWidth) {
        columnDef.minWidth = formatConfig.minWidth
      }
      if (formatConfig.maxWidth) {
        columnDef.maxWidth = formatConfig.maxWidth
      }
      // ✅ Thêm xử lý hide
      if (formatConfig.hide === true) {
        columnDef.hide = true
      }
      if (formatConfig.wrapText === true) {
        columnDef.wrapText = true
        columnDef.autoHeight = true // autoHeight cần đi kèm wrapText
      }
    }

    if (!columnDef.headerName) columnDef.headerName = field
    return columnDef
  })

  state.agGridColumnDefs_flat = columnDefs
  return groupMeasureColumns(columnDefs, filteredMeasureColumns)
}

function groupMeasureColumns(columnDefs, filteredMeasureColumns) {
  const measureGroups = {}
  const columnDefMap = new Map(columnDefs.map((col) => [col.field, col]))

  filteredMeasureColumns.forEach((field) => {
    const parts = field.split('_')
    if (parts.length === 2) {
      const [prefix, suffix] = parts
      if (!measureGroups[prefix]) measureGroups[prefix] = []
      const originalDef = columnDefMap.get(field)
      measureGroups[prefix].push({ ...originalDef, headerName: suffix })
    }
  })

  if (Object.keys(measureGroups).length === 0) return columnDefs

  const finalColumnDefs = []
  const processedPrefixes = new Set()

  for (const col of columnDefs) {
    const parts = col.field.split('_')
    const prefix = parts[0]
    if (parts.length === 2 && measureGroups[prefix]) {
      if (!processedPrefixes.has(prefix)) {
        processedPrefixes.add(prefix)
        finalColumnDefs.push({
          headerName: prefix,
          headerClass: 'header-center',
          children: measureGroups[prefix]
        })
      }
      continue
    }
    finalColumnDefs.push(col)
  }
  return finalColumnDefs
}

// ======================
// TREE FUNCTIONS
// ======================
function buildTree(data) {
  let idCounter = 0
  const rootMap = {}

  for (const row of data) {
    const treeLevels = Object.keys(row)
      .filter((k) => k.startsWith('tree_lv'))
      .sort(
        (a, b) =>
          parseInt(a.replace('tree_lv', '')) -
          parseInt(b.replace('tree_lv', ''))
      )

    let currentLevel = rootMap
    let parent = null

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

    if (state.leaf_in_tree === '0') {
      const leafNode = {
        id: ++idCounter,
        name: null,
        level: treeLevels.length + 1,
        leaf: true
      }
      for (const [key, val] of Object.entries(row)) {
        if (!key.startsWith('tree_lv')) leafNode[key] = val
      }
      parent.children[`leaf_${idCounter}`] = leafNode
    } else {
      parent.leaf = true
      for (const [key, val] of Object.entries(row)) {
        if (!key.startsWith('tree_lv')) parent[key] = val
      }
      delete parent.children
    }
  }

  return Object.values(rootMap).map((n) => normalizeTree(n))
}

function normalizeTree(node) {
  if (node.children && !Array.isArray(node.children)) {
    node.children = Object.values(node.children).map((n) => normalizeTree(n))
  }
  return node
}

function aggregateTreeValues(nodes, numericCols, formulaRules = []) {
  for (const node of nodes) {
    if (node.children?.length) {
      aggregateTreeValues(node.children, numericCols, formulaRules)

      numericCols.forEach((col) => (node[col] = 0))

      for (const child of node.children) {
        numericCols.forEach((col) => {
          if (!isFormulaColumn(col, formulaRules)) {
            const val = Number(child[col])
            if (!isNaN(val)) node[col] += val
          }
        })
      }

      numericCols.forEach((col) => {
        const rule = getFormulaRule(col, formulaRules)
        if (!rule) return
        const prefix = col.replace(rule.targetSuffix, '')
        const numerator = Number(node[prefix + rule.numeratorSuffix]) || 0
        const denominator = Number(node[prefix + rule.denominatorSuffix]) || 0
        const operatorFn = OPERATORS[rule.operator]
        if (operatorFn) node[col] = operatorFn(numerator, denominator)
      })
    }
  }
}

function flattenTree(nodes, level = 1) {
  let result = []
  const orderList = state.levelSortRules?.[level]
  const sortedNodes = orderList
    ? [...nodes].sort(customValueComparator(orderList))
    : nodes

  for (const n of sortedNodes) {
    result.push({ ...n, level })
    if (n.expanded && n.children?.length) {
      result = result.concat(flattenTree(n.children, level + 1))
    }
  }
  return result
}

function setAllExpanded(nodes, expanded) {
  if (!nodes?.length) return
  for (const n of nodes) {
    if (n.children?.length) {
      n.expanded = expanded
      setAllExpanded(n.children, expanded)
    }
  }
}

function setSubtreeExpanded(node, expanded) {
  if (!node) return
  node.expanded = expanded
  if (node.children)
    node.children.forEach((child) => setSubtreeExpanded(child, expanded))
}

function findNodeById(nodes, id) {
  for (const n of nodes) {
    if (n.id == id) return n
    if (n.children) {
      const found = findNodeById(n.children, id)
      if (found) return found
    }
  }
  return null
}

function applyExpandLevel(nodes, level) {
  for (const n of nodes) {
    n.expanded = n.level < level
    if (n.children) applyExpandLevel(n.children, level)
  }
}

// ======================
// FORMULA HELPERS
// ======================
function isFormulaColumn(col, rules) {
  return rules.some((r) => col.endsWith(r.targetSuffix))
}

function getFormulaRule(col, rules) {
  return rules.find((r) => col.endsWith(r.targetSuffix))
}

function calcTotalsTree(nodes, numericCols) {
  const totals = Object.fromEntries(numericCols.map((col) => [col, 0]))

  const traverse = (nodeList) => {
    for (const node of nodeList) {
      if (node.children?.length) traverse(node.children)
      if (node.leaf) {
        numericCols.forEach((col) => {
          const val = Number(node[col])
          if (!isNaN(val)) totals[col] += val
        })
      }
    }
  }

  traverse(nodes)
  return totals
}

// ======================
// PIVOT FUNCTIONS
// ======================
function pivotData(
  list_columns,
  list_rows,
  list_column_horizontal,
  list_column_vertical,
  list_column_measure
) {
  const rows = list_rows.map((rowArr) => {
    const obj = {}
    list_columns.forEach((col, i) => (obj[col] = rowArr[i]))
    return obj
  })

  const resultMap = new Map()
  const verticalKeys = new Set()

  for (const row of rows) {
    const horizontalKey = list_column_horizontal.map((c) => row[c]).join('||')
    const verticalKey = list_column_vertical.map((c) => row[c]).join('_')
    verticalKeys.add(verticalKey)

    let measureValue = row[list_column_measure]
    if (measureValue != null) measureValue = Number(measureValue)

    if (!resultMap.has(horizontalKey)) {
      const base = {}
      list_column_horizontal.forEach((c) => (base[c] = row[c]))
      resultMap.set(horizontalKey, base)
    }
    resultMap.get(horizontalKey)[verticalKey] = measureValue
  }

  return {
    pivotData: Array.from(resultMap.values()),
    dimensionColumns: [...list_column_horizontal],
    measureColumns: Array.from(verticalKeys)
  }
}

// ======================
// CELL RENDERER
// ======================
const nameCellRenderer = (params) => {
  const node = params.data
  if (!node) return ''
  const indent = '<span class="tree-indent"></span>'.repeat(node.level - 1)
  if (node.leaf)
    return indent + '<span class="tree-indent"></span> ' + (node.name || '')
  const symbol = node.expanded ? '▾' : '▸'
  return (
    indent +
    `<span class="toggle-btn" data-id="${node.id}">${symbol}</span> ` +
    node.name
  )
}

// ======================
// GRID FUNCTIONS
// ======================
function toggleNode(nodeId) {
  const recursiveToggle = (nodes) => {
    for (const n of nodes) {
      if (n.id == nodeId) {
        n.expanded = !n.expanded
        return true
      }
      if (n.children && recursiveToggle(n.children)) return true
    }
    return false
  }
  recursiveToggle(state.nestedData)

  const flatData = flattenTree(state.nestedData)
  setTimeout(() => {
    state.gridApi.setGridOption('rowData', flatData)
    const rowNode = state.gridApi.getDisplayedRowAtIndex(
      flatData.findIndex((r) => r.id == nodeId)
    )
    if (rowNode) state.gridApi.ensureNodeVisible(rowNode, 'middle')
  }, 0)
}

function copySelectedRows() {
  const selectedNodes = []
  state.gridApi.forEachNodeAfterFilterAndSort((node) => {
    if (node.isSelected()) selectedNodes.push(node)
  })

  if (!selectedNodes.length) {
    alert('⚠️ Chưa chọn dòng nào!')
    return
  }

  const displayedCols = state.gridApi.getAllDisplayedColumns()
  const text = selectedNodes
    .map((node) =>
      displayedCols
        .map((col) => {
          let v = node.data[col.getColId()]
          if (v == null) return ''
          if (typeof v === 'object') return ''
          return v.toString()
        })
        .join('\t')
    )
    .join('\n')

  const textarea = document.createElement('textarea')
  textarea.value = text
  textarea.style.position = 'fixed'
  textarea.style.top = '-9999px'
  document.body.appendChild(textarea)
  textarea.select()
  document.execCommand('copy')
  document.body.removeChild(textarea)
}

function adjustGridHeight() {
  const container = document.querySelector('.container')
  const toolbar = document.querySelector('.toolbar')
  const notebar = document.querySelector('.notebar')
  const gridContainer = document.getElementById('gridContainer')

  const totalHeight = window.innerHeight
  const toolbarHeight = toolbar?.offsetHeight || 0
  const notebarHeight = notebar?.offsetHeight || 0
  const gridHeight = totalHeight - toolbarHeight - notebarHeight - 5
  if (gridContainer) gridContainer.style.height = `${gridHeight}px`
}

// ======================
// EXPORT FUNCTIONS
// ======================
function exportFlattenWithPath(
  nodes,
  currentPath = [],
  result = [],
  maxLevelRef = { max: 0 },
  level = 1
) {
  const orderList = state.levelSortRules?.[level]
  const sortedNodes = orderList
    ? [...nodes].sort(customValueComparator(orderList))
    : nodes

  for (const node of sortedNodes) {
    if (!node.leaf && node.name) {
      const nodePath = [...currentPath, node.name]
      maxLevelRef.max = Math.max(maxLevelRef.max, nodePath.length)
      result.push({ ...node, path: nodePath })
    } else if (node.leaf) {
      const leafPath = [...currentPath, node.name]
      maxLevelRef.max = Math.max(maxLevelRef.max, leafPath.length)
      result.push({ ...node, path: leafPath })
    }

    if (node.expanded && node.children?.length) {
      exportFlattenWithPath(
        node.children,
        [...currentPath, node.name],
        result,
        maxLevelRef,
        level + 1
      )
    }
  }
  return result
}

// ======================
// RENDER FUNCTION
// ======================
function loadAndRender(worksheet) {
  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
    const cols = sumData.columns.map((c) => c.fieldName)
    const rows = sumData.data.map((r) =>
      r.map((c) => {
        if (c.nativeValue == null) return ''
        if (c.nativeValue instanceof Date)
          return c.nativeValue.toLocaleDateString('vi-VN')
        if (typeof c.nativeValue === 'number') return c.nativeValue
        return c.formattedValue
      })
    )

    state.pivotDataOutput = pivotData(
      cols,
      rows,
      state.list_column_horizontal,
      state.list_column_vertical,
      state.list_column_measure
    )

    state.nestedData = buildTree(state.pivotDataOutput.pivotData)
    aggregateTreeValues(
      state.nestedData,
      state.pivotDataOutput.measureColumns,
      state.percent_columns
    )

    let flatData = flattenTree(state.nestedData)
    state.maxTreeLevel = getMaxTreeLevel(state.nestedData)
    state.currentExpandedLevel = 1

    state.agGridColumnDefs = createColumnDefs(
      state.pivotDataOutput.dimensionColumns,
      state.pivotDataOutput.measureColumns,
      state.pivot_column_config,
      state.list_exclude_column_config,
      state.formated_columns
    )

    const gridOptions = {
      columnDefs: state.agGridColumnDefs,
      rowData: flatData,
      defaultColDef: {
        filter: false,
        sortable: true,
        resizable: true,
        cellStyle: { fontSize: state.font_size }
      },
      suppressFieldDotNotation: true,
      getRowStyle: (params) => {
        const node = params.data
        if (!node) return null
        if (node.name === 'Grand Total')
          return {
            fontWeight: 'bold',
            color: '#d00000',
            backgroundColor: '#fabcbcff'
          }
        if (node.children?.length)
          return { fontWeight: 'bold', backgroundColor: '#f7f7f7' }
        return null
      },
      getRowHeight: (params) => {
        if (params.data && params.data.name === 'Grand Total')
          return state.row_height
        return state.row_height // Mặc định undefined
      },
      rowSelection: {
        mode: 'multiRow',
        checkboxes: true,
        enableClickSelection: false
      },
      onCellClicked: (params) => {
        const el = params.event.target
        if (el.classList.contains('toggle-btn')) {
          toggleNode(el.dataset.id)
        } else {
          state.selectedCellValue = params.value
          state.gridApi.deselectAll()
          params.node.setSelected(true)
        }
      },
      // onFilterChanged: () => {
      //   applyColumnSizing()
      // },
      onGridReady: (params) => {
        state.gridApi = params.api

        // ✅ Gọi áp dụng sizing
        applyColumnSizing()

        if (state.showGrandTotal == 1) updateFooterTotals()
      },
      onFirstDataRendered: () => {
        // ✅ Đảm bảo sizing lại sau khi data rendered lần đầu
        applyColumnSizing()

        if (state.showGrandTotal == 1) updateFooterTotals()
      }
    }

    const eGridDiv = document.querySelector('#gridContainer')
    if (!state.gridApi) state.gridApi = agGrid.createGrid(eGridDiv, gridOptions)
    else {
      state.gridApi.setGridOption('rowData', flatData)
      state.gridApi.setGridOption('columnDefs', state.agGridColumnDefs)

      // ✅ Gọi lại fit sizing sau khi cập nhật
      setTimeout(() => {
        applyColumnSizing()
      }, 200)
    }

    setupExpandButtons()
    setupEventListeners()
    if (state.showGrandTotal == 1) {
      console.log('vao day', state.showGrandTotal)

      setTimeout(() => document.getElementById('updateTotal')?.click(), 1000)
    } else {
      console.log('ko vao day', state.showGrandTotal)
    }
  })
}

function updateFooterTotals() {
  console.log('vao ham updateFooterTotals')

  if (!state.gridApi) return

  const allData = []
  state.gridApi.forEachNodeAfterFilterAndSort((node) => {
    if (!node.rowPinned) allData.push(node.data)
  })

  const numericCols = state.pivotDataOutput.measureColumns
  const totals = calcTotalsTree(allData, numericCols)

  const totalRow = {}
  state.agGridColumnDefs_flat.forEach((col) => {
    const field = col.field
    if (numericCols.includes(field)) totalRow[field] = totals[field] || 0
    else if (field === state.agGridColumnDefs_flat[0]?.field)
      totalRow[field] = 'Grand Total'
    else totalRow[field] = ''
  })

  numericCols.forEach((col) => {
    const rule = getFormulaRule(col, state.percent_columns)
    if (!rule) return
    const prefix = col.replace(rule.targetSuffix, '')
    const numerator = Number(totalRow[prefix + rule.numeratorSuffix]) || 0
    const denominator = Number(totalRow[prefix + rule.denominatorSuffix]) || 0
    const operatorFn = OPERATORS[rule.operator]
    if (operatorFn) totalRow[col] = operatorFn(numerator, denominator)
  })

  totalRow.leaf = true

  // 👈 Xử lý hiển thị theo vị trí
  if (state.grandTotalPosition === 'top') {
    // Hiển thị ở trên cùng: dùng pinnedTopRowData
    state.gridApi.setGridOption('pinnedTopRowData', [totalRow])
    state.gridApi.setGridOption('pinnedBottomRowData', []) // Xóa bottom nếu có
  } else {
    // Hiển thị ở dưới cùng (mặc định): dùng pinnedBottomRowData
    state.gridApi.setGridOption('pinnedBottomRowData', [totalRow])
    state.gridApi.setGridOption('pinnedTopRowData', []) // Xóa top nếu có
  }
}

function setupExpandButtons() {
  if (state.expandListenersBound) return

  const btnExpand1Level = document.getElementById('btnExpand1Level')
  const btnCollapse1Level = document.getElementById('btnCollapse1Level')
  const btnExpandAll = document.getElementById('btnExpandAll')
  const btnCollapseAll = document.getElementById('btnCollapseAll')

  if (btnExpand1Level) {
    btnExpand1Level.addEventListener('click', () => {
      if (state.currentExpandedLevel < state.maxTreeLevel)
        state.currentExpandedLevel++
      applyExpandLevel(state.nestedData, state.currentExpandedLevel)
      setTimeout(
        () =>
          state.gridApi.setGridOption('rowData', flattenTree(state.nestedData)),
        0
      )
    })
  }

  if (btnCollapse1Level) {
    btnCollapse1Level.addEventListener('click', () => {
      if (state.currentExpandedLevel > 1) state.currentExpandedLevel--
      applyExpandLevel(state.nestedData, state.currentExpandedLevel)
      setTimeout(
        () =>
          state.gridApi.setGridOption('rowData', flattenTree(state.nestedData)),
        0
      )
    })
  }

  if (btnExpandAll) {
    btnExpandAll.addEventListener('click', () => {
      const selectedNodes = []
      state.gridApi.forEachNode((node) => {
        if (node.isSelected()) selectedNodes.push(node.data)
      })
      const targetId = selectedNodes.length ? selectedNodes[0].id : null

      if (!targetId) setAllExpanded(state.nestedData, true)
      else {
        const node = findNodeById(state.nestedData, targetId)
        if (node) setSubtreeExpanded(node, true)
      }

      const flat = flattenTree(state.nestedData)
      setTimeout(() => {
        state.gridApi.setGridOption('rowData', flat)
        if (targetId) {
          requestAnimationFrame(() => {
            const idx = flat.findIndex((r) => r.id == targetId)
            const rowNode = state.gridApi.getDisplayedRowAtIndex(idx)
            if (rowNode) state.gridApi.ensureNodeVisible(rowNode, 'middle')
          })
        }
      }, 0)
      state.currentExpandedLevel = state.maxTreeLevel
    })
  }

  if (btnCollapseAll) {
    btnCollapseAll.addEventListener('click', () => {
      const selectedNodes = []
      state.gridApi.forEachNode((node) => {
        if (node.isSelected()) selectedNodes.push(node.data)
      })
      const targetId = selectedNodes.length ? selectedNodes[0].id : null

      if (!targetId) setAllExpanded(state.nestedData, false)
      else {
        const node = findNodeById(state.nestedData, targetId)
        if (node) setSubtreeExpanded(node, false)
      }

      const flat = flattenTree(state.nestedData)
      setTimeout(() => {
        state.gridApi.setGridOption('rowData', flat)
        if (targetId) {
          requestAnimationFrame(() => {
            const idx = flat.findIndex((r) => r.id == targetId)
            const rowNode = state.gridApi.getDisplayedRowAtIndex(idx)
            if (rowNode) state.gridApi.ensureNodeVisible(rowNode, 'middle')
          })
        }
      }, 0)
      state.currentExpandedLevel = 1
    })
  }

  state.expandListenersBound = true
}

function setupEventListeners() {
  document
    .getElementById('clearAllFilterBtn')
    ?.addEventListener('click', () => {
      if (state.gridApi) state.gridApi.onFilterChanged()
    })

  document
    .getElementById('updateTotal')
    ?.addEventListener('click', updateFooterTotals)

  document
    .getElementById('copyRow')
    ?.addEventListener('click', copySelectedRows)

  document.getElementById('copyCellBtn')?.addEventListener('click', () => {
    if (state.selectedCellValue === null) {
      alert('Chưa chọn ô nào để copy!')
      return
    }
    const textarea = document.createElement('textarea')
    textarea.value = state.selectedCellValue.toString()
    textarea.style.position = 'fixed'
    textarea.style.top = '-9999px'
    document.body.appendChild(textarea)
    textarea.select()
    document.execCommand('copy')
    document.body.removeChild(textarea)
  })

  setupExportExcel()
}

function setupExportExcel() {
  document.getElementById('exportExcel')?.addEventListener('click', () => {
    if (!state.gridApi || !state.nestedData?.length) {
      alert('⚠️ Không có dữ liệu để export!')
      return
    }

    const maxLevelRef = { max: 0 }
    const exportRows = exportFlattenWithPath(
      state.nestedData,
      [],
      [],
      maxLevelRef
    )
    const maxTreeLevel = maxLevelRef.max

    // 👈 Lấy cả top và bottom nếu có
    let pinnedRows = []

    // Lấy pinned top rows nếu có
    if (
      state.grandTotalPosition === 'top' &&
      state.gridApi.getPinnedTopRowCount() > 0
    ) {
      pinnedRows = Array.from(
        { length: state.gridApi.getPinnedTopRowCount() },
        (_, i) => state.gridApi.getPinnedTopRow(i).data
      )
    }

    // Lấy pinned bottom rows nếu có
    if (
      state.grandTotalPosition === 'bottom' &&
      state.gridApi.getPinnedBottomRowCount() > 0
    ) {
      pinnedRows = Array.from(
        { length: state.gridApi.getPinnedBottomRowCount() },
        (_, i) => state.gridApi.getPinnedBottomRow(i).data
      )
    }

    // 👈 Sắp xếp thứ tự rows cho export
    let allExportRows = [...exportRows]

    if (state.grandTotalPosition === 'top') {
      // Grand Total ở trên cùng
      allExportRows = [...pinnedRows, ...exportRows]
    } else {
      // Grand Total ở dưới cùng
      allExportRows = [...exportRows, ...pinnedRows]
    }
    const currentColumnDefs = state.gridApi.getColumnDefs()
    const otherColumnDefs = currentColumnDefs.slice(1)

    const getLeafColumns = (cols, parentPath = []) => {
      const leaves = []
      cols.forEach((col) => {
        if (!col.children?.length) {
          leaves.push({
            field: col.field,
            headerPath: [...parentPath, col.headerName || col.field || ''],
            colDef: col
          })
        } else {
          leaves.push(
            ...getLeafColumns(col.children, [
              ...parentPath,
              col.headerName || ''
            ])
          )
        }
      })
      return leaves
    }

    const leafColumns = getLeafColumns(otherColumnDefs)
    const maxHeaderDepth = Math.max(
      ...leafColumns.map((col) => col.headerPath.length),
      0
    )

    const headerRows = []
    for (let depth = 0; depth < maxHeaderDepth; depth++) {
      const headerRow = []
      for (let i = 0; i < maxTreeLevel; i++)
        headerRow.push(depth === 0 ? `Level ${i + 1}` : '')
      leafColumns.forEach((col) =>
        headerRow.push(
          depth < col.headerPath.length ? col.headerPath[depth] : ''
        )
      )
      headerRows.push(headerRow)
    }

    const wsData = [...headerRows]
    allExportRows.forEach((row) => {
      const rowVals = []
      const firstField = currentColumnDefs[0]?.field
      const isTotal = row[firstField] === 'Grand Total'

      if (isTotal) {
        rowVals.push('Grand Total')
        for (let i = 1; i < maxTreeLevel; i++) rowVals.push('')
      } else {
        const path = row.path || []
        for (let i = 0; i < maxTreeLevel; i++) rowVals.push(path[i] || '')
      }

      leafColumns.forEach((colDef) => {
        let val = row[colDef.field] ?? ''
        rowVals.push(typeof val === 'number' ? val : val.toString())
      })
      wsData.push(rowVals)
    })

    const ws = XLSX.utils.aoa_to_sheet(wsData)
    ws['!merges'] = ws['!merges'] || []

    // Merge logic remains similar, simplified here for brevity
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, ws, 'TreeData')
    XLSX.writeFile(
      wb,
      `tree_data_${new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)}.xlsx`
    )
  })
}

// ======================
// INITIALIZATION
// ======================
function refreshExtractTime(worksheet) {
  worksheet.getDataSourcesAsync().then((dataSources) => {
    dataSources.forEach((ds) => {
      state.extractRefreshTime = ds.isExtract
        ? `Extract Refresh Time: ${ds.extractUpdateTime}`
        : ''
      const el = document.getElementById('extractRefreshTime')
      if (el) el.innerText = state.extractRefreshTime
    })
  })
}

document.addEventListener('DOMContentLoaded', () => {
  tableau.extensions.initializeAsync().then(() => {
    const worksheet =
      tableau.extensions.dashboardContent.dashboard.worksheets.find(
        (ws) => ws.name === 'DataTableExtSheet'
      )
    if (!worksheet)
      return console.error("❌ Không tìm thấy worksheet 'DataTableExtSheet'")

    const worksheetConfig =
      tableau.extensions.dashboardContent.dashboard.worksheets.find(
        (ws) => ws.name === 'ConfigSheet'
      )
    if (!worksheetConfig)
      return console.error("❌ Không tìm thấy worksheet 'ConfigSheet'")

    worksheetConfig.getSummaryDataAsync({ maxRows: 0 }).then((configData) => {
      configData.data.forEach((item) => {
        const key = item[0]?.formattedValue
        const value = item[1]?.formattedValue
        switch (key) {
          case 'list_column_horizontal':
            state.list_column_horizontal = value.split(',')
            break
          case 'list_column_vertical':
            state.list_column_vertical = value.split(',')
            break
          case 'list_column_measure':
            state.list_column_measure = value.split(',')
            break
          case 'pivot_column_config':
            state.pivot_column_config = JSON.parse(value)
            break
          case 'list_exclude_column_config':
            state.list_exclude_column_config = value.split(',')
            break
          case 'order_lv1':
            state.order_lv1 = value
            break
          case 'order_lv2':
            state.order_lv2 = value
            break
          case 'formated_columns':
            state.formated_columns = value
            break
          case 'leaf_in_tree':
            state.leaf_in_tree = value
            break
          case 'level_sort_rules':
            state.levelSortRules = JSON.parse(value)
            break
          case 'show_grand_total':
            console.log('nhan bien show_grand_total:', value)

            state.showGrandTotal = value
            break
          case 'percent_columns':
            state.percent_columns = JSON.parse(value)
            break
          case 'font_size':
            console.log('font_size', state.font_size)

            state.font_size = value
            console.log('font_size', state.font_size)
            break
          case 'row_height':
            state.row_height = value
            break
          case 'fit_content':
            state.fit_content = value
            break
          case 'fit_window':
            state.fit_window = value
            break
          case 'grand_total_position': // 👈 Thêm case mới
            state.grandTotalPosition = value || 'bottom'
            break
          case 'toolbar_position': // 👈 Thêm case mới
            const toolbar = document.querySelector('.toolbar')
            const grid = document.getElementById('gridContainer')
            state.toolbar_position = value || 'down'
            if (state.toolbar_position === 'down') {
              toolbar.style.order = 2
              grid.style.order = 1
            } else {
              toolbar.style.order = 1
              grid.style.order = 2
            }
            break
        }
      })
    })

    refreshExtractTime(worksheet)
    loadAndRender(worksheet)

    worksheet.addEventListener(tableau.TableauEventType.FilterChanged, () => {
      refreshExtractTime(worksheet)
      loadAndRender(worksheet)
    })

    tableau.extensions.dashboardContent.dashboard
      .getParametersAsync()
      .then((parameters) => {
        parameters.forEach((p) =>
          p.addEventListener(tableau.TableauEventType.ParameterChanged, () => {
            refreshExtractTime(worksheet)
            loadAndRender(worksheet)
          })
        )
      })

    adjustGridHeight()
    window.addEventListener('resize', adjustGridHeight)
  })
})
