'use strict'

let list_column_horizontal
let list_column_vertical
let list_column_measure
let list_exclude_column_config
let order_lv1
let order_lv2
let leaf_in_tree
let datanew
let newnestedData
let dimensionColumns
let measureColumns
let pivotDataOutput
let pivot_column_config
let formated_columns
let numericCols

let agGridColumnDefs
let agGridColumnDefs_flat

let selectedCellValue = null
let expandListenersBound = false // <-- thêm dòng này
let extractRefreshTime = ''

let gridApi = null
let nestedData = []

let currentExpandedLevel = 1
let maxTreeLevel = 1

let levelSortRules
let showGrandTotal = 1

// ⭐ Hàm cellRenderer tùy chỉnh cho cột 'name' (giữ nguyên)
const nameCellRenderer = (params) => {
  const node = params.data
  if (!node) return ''

  const indent = '<span class="tree-indent"></span>'.repeat(node.level - 1)
  if (node.leaf) {
    return (
      indent + '<span class="tree-indent"></span>' + ' ' + (node.name || '')
    )
  } else {
    const symbol = node.expanded ? '▾' : '▸'
    return (
      indent +
      `<span class="toggle-btn" data-id="${node.id}">${symbol}</span> ` +
      node.name
    )
  }
}

// ⭐ Hàm Style cho giá trị âm (Negative Cell Style)
// Trả về style cho AG Grid, khiến số âm có màu đỏ
const negativeCellStyle = (params) => {
  const style = { textAlign: 'right' }
  // if (params.value < 0) {
  //   style.color = 'red'
  // }
  return style
}

function sortColumns(list_columns, order_lv1, order_lv2) {
  // ---- helper: parse order config ----
  const parseOrder = (order) => {
    if (order === 'asc' || order === 'desc') {
      return { type: order }
    }
    // cho phép truyền string CSV hoặc array
    const list = Array.isArray(order)
      ? order
      : String(order)
          .split(',')
          .map((v) => v.trim())
    return { type: 'custom', list }
  }

  const lv1Order = parseOrder(order_lv1)
  const lv2Order = parseOrder(order_lv2)

  // ---- parse columns ----
  const parsed = list_columns.map((c) => {
    const [lv1, lv2] = c.split('_')
    return { raw: c, lv1, lv2 }
  })

  // ---- group by level1 ----
  const groupMap = new Map()
  parsed.forEach((item) => {
    if (!groupMap.has(item.lv1)) groupMap.set(item.lv1, [])
    groupMap.get(item.lv1).push(item)
  })

  // ---- sort level1 keys ----
  let lv1Keys = Array.from(groupMap.keys())

  if (lv1Order.type === 'asc') {
    lv1Keys.sort((a, b) => a.localeCompare(b))
  } else if (lv1Order.type === 'desc') {
    lv1Keys.sort((a, b) => b.localeCompare(a))
  } else {
    const idx = new Map(lv1Order.list.map((v, i) => [v, i]))
    lv1Keys.sort((a, b) => (idx.get(a) ?? 9999) - (idx.get(b) ?? 9999))
  }

  // ---- sort level2 inside each group ----
  const result = []

  lv1Keys.forEach((lv1) => {
    const rows = groupMap.get(lv1)

    if (lv2Order.type === 'asc') {
      rows.sort((a, b) => a.lv2.localeCompare(b.lv2))
    } else if (lv2Order.type === 'desc') {
      rows.sort((a, b) => b.lv2.localeCompare(a.lv2))
    } else {
      const idx = new Map(lv2Order.list.map((v, i) => [v.padStart(2, '0'), i]))
      rows.sort((a, b) => (idx.get(a.lv2) ?? 9999) - (idx.get(b.lv2) ?? 9999))
    }

    rows.forEach((r) => result.push(r.raw))
  })

  return result
}

/**
 * Chuyển đổi cấu hình tuỳ chỉnh và danh sách cột pivot thành AG Grid columnDefs.
 * @param {Array<string>} dimensionColumns
 * @param {Array<string>} measureColumns
 * @param {Array<object>} customConfig
 * @param {Array<string>} excludeColumns - Danh sách cột cần loại bỏ
 * @param {string json} formatedColumns - Danh sách cột được format dữ liệu
 * @returns {Array<object>}
 */
function createColumnDefs(
  dimensionColumns,
  measureColumns,
  customConfig,
  excludeColumns = [],
  formated_columns
) {
  const FORMATTERS = {
    percent:
      (opts = {}) =>
      (params) => {
        const v = Number(params.value)
        if (isNaN(v)) return params.value ?? ''
        const precision = opts.precision ?? 2
        return `${(v * 100).toFixed(precision)} %`
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
        return v.toLocaleString('en-US', {
          maximumFractionDigits: opts.precision ?? 2
        })
      }
  }

  const columnFormatMatchers = []

  if (formated_columns) {
    try {
      const formats =
        typeof formated_columns === 'string'
          ? JSON.parse(formated_columns)
          : formated_columns

      formats.forEach((f) => {
        if (!f.field || !f.formatType) return

        let matcher

        const pattern = f.field

        if (pattern.startsWith('%') && pattern.endsWith('%')) {
          // %abc% → contains
          const key = pattern.slice(1, -1)
          matcher = (field) => field.includes(key)
        } else if (pattern.startsWith('%')) {
          // %abc → endsWith
          const key = pattern.slice(1)
          matcher = (field) => field.endsWith(key)
        } else if (pattern.endsWith('%')) {
          // abc% → startsWith
          const key = pattern.slice(0, -1)
          matcher = (field) => field.startsWith(key)
        } else {
          // exact match
          matcher = (field) => field === pattern
        }

        columnFormatMatchers.push({
          matcher,
          config: f
        })
      })
    } catch (e) {
      console.error('Invalid formated_columns JSON', e)
    }
  }

  // console.log("columnFormatMatchers", columnFormatMatchers);

  function resolveFormatedColumns(field) {
    let matched = null
    let priority = -1

    columnFormatMatchers.forEach(({ matcher, config }) => {
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

  // 1. Lọc dimension và measure
  const filteredDimensionColumns = dimensionColumns.filter(
    (field) => !field.startsWith('tree_lv') && !excludeColumns.includes(field)
  )

  // const filteredMeasureColumns = measureColumns
  //   .filter((field) => !excludeColumns.includes(field))
  //   .sort() // Sắp xếp tăng dần theo thứ tự alphabet

  const filteredMeasureColumns = sortColumns(
    measureColumns.filter((field) => !excludeColumns.includes(field)),
    order_lv1,
    order_lv2
  )

  console.log('filteredDimensionColumns', filteredDimensionColumns)
  console.log('filteredMeasureColumns', filteredMeasureColumns)
  console.log(
    'filteredMeasureColumns',
    sortColumns(filteredMeasureColumns, order_lv1, order_lv2)
  )

  // 2. Map custom config
  const configMap = new Map()
  const orderedFieldsFromConfig = []

  customConfig.forEach((col) => {
    if (excludeColumns.includes(col.field)) return

    // Quan trọng: PHẢI CHỈ GIỮ col.field nếu nó nằm trong dimension/measure đã lọc
    if (
      col.field !== 'name' &&
      !filteredDimensionColumns.includes(col.field) &&
      !filteredMeasureColumns.includes(col.field)
    ) {
      return // bỏ khỏi output hoàn toàn
    }

    configMap.set(col.field, col)
    orderedFieldsFromConfig.push(col.field)
  })

  // 3. Merge field theo thứ tự:
  // - name
  // - customConfig hợp lệ
  // - dimension + measure còn lại
  const mergedFields = Array.from(
    new Set([
      'name',
      ...orderedFieldsFromConfig,
      ...filteredDimensionColumns,
      ...filteredMeasureColumns
    ])
  ).filter((f) => !excludeColumns.includes(f))

  // 4. Build columnDefs
  const columnDefs = mergedFields.map((field) => {
    const customProps = configMap.get(field)

    const isDimensionOrName =
      filteredDimensionColumns.includes(field) || field === 'name'
    const isMeasure = filteredMeasureColumns.includes(field)

    // Default
    let columnDef = {
      field,
      headerName: field, //.replace(/_/g, '\n'),
      width: isDimensionOrName ? 250 : 200,
      wrapHeaderText: true,
      autoHeaderHeight: true
    }

    // console.log('======== field, columnDef', field, columnDef)

    // Alignment & formatter
    if (isDimensionOrName) {
      columnDef.cellStyle = { textAlign: 'left' }
    } else if (isMeasure) {
      columnDef.cellStyle = negativeCellStyle
      columnDef.valueFormatter = (params) => {
        const v = params.value
        if (v == null || v === '') return ''
        const num = Number(v)
        if (isNaN(num)) return v
        return num.toLocaleString('en-US', { maximumFractionDigits: 2 })
      }
    }

    // Cột name
    if (field === 'name') {
      columnDef = {
        ...columnDef,
        cellRenderer: nameCellRenderer,
        headerName: customProps?.headerName || 'Tên'
      }
    }

    // Override bằng customConfig
    if (customProps) {
      columnDef = {
        ...columnDef,
        ...customProps
      }
    }

    // Override formatedColumns ở đây (support wildcard %)
    const formatConfig = resolveFormatedColumns(field)
    // console.log('==========formatConfig: ', field, formatConfig)

    if (formatConfig && FORMATTERS[formatConfig.formatType]) {
      columnDef.valueFormatter =
        FORMATTERS[formatConfig.formatType](formatConfig)
      columnDef.type = 'numericColumn'
      columnDef.cellStyle = { textAlign: 'right' }
      if (formatConfig.width) {
        columnDef.width = formatConfig.width
      }

      // console.log('columnDef',columnDef);
    }

    // Ensure headerName
    if (!columnDef.headerName) {
      columnDef.headerName = field
    }

    return columnDef
  })

  agGridColumnDefs_flat = columnDefs
  // console.log('==========columnDefs', columnDefs)

  // -----------------------------------------------------------
  // 5. GROUPING các measure columns dạng PREFIX_SUFFIX
  // -----------------------------------------------------------

  const measureGroups = {} // { prefix: [childColumns] }

  // Tạo map để lấy lại columnDef gốc
  const columnDefMap = new Map()
  columnDefs.forEach((col) => columnDefMap.set(col.field, col))

  filteredMeasureColumns.forEach((field) => {
    const parts = field.split('_')
    if (parts.length === 2) {
      const [prefix, suffix] = parts

      if (!measureGroups[prefix]) {
        measureGroups[prefix] = []
      }

      // Lấy lại columnDef gốc
      const originalDef = columnDefMap.get(field)

      // Clone + override headerName
      const childDef = {
        ...originalDef,
        headerName: suffix //.padStart(2, '0')
      }

      measureGroups[prefix].push(childDef)
    }
  })

  // Nếu không có group nào → giữ nguyên
  if (Object.keys(measureGroups).length === 0) return columnDefs

  // Áp dụng grouping
  const finalColumnDefs = []
  const processedPrefixes = new Set()

  for (const col of columnDefs) {
    const field = col.field

    const parts = field.split('_')
    const prefix = parts[0]

    // Đây là measure dạng PREFIX_SUFFIX
    if (parts.length === 2 && measureGroups[prefix]) {
      if (!processedPrefixes.has(prefix)) {
        processedPrefixes.add(prefix)

        finalColumnDefs.push({
          headerName: prefix,
          headerClass: 'header-center',
          children: measureGroups[prefix]
        })
      }
      continue // Không add cột lẻ nữa
    }

    // Không phải grouped measure → giữ nguyên
    finalColumnDefs.push(col)
  }

  return finalColumnDefs
}

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

// Helper: Extract tên cột từ định dạng tree_lvN("Tên")
function extractTreeHeaderName(header) {
  // Clean width nếu có: tree_lv1("Tên")(300) -> tree_lv1("Tên")
  const cleanHeader = header.replace(/\(\s*\d+\s*\)\s*$/, '').trim()
  // Parse ngoặc: tree_lv1("Cây tổ chức") -> "Cây tổ chức"
  const match = cleanHeader.match(/tree_lv\d+\s*\(\s*(.+?)\s*\)/)
  return match ? match[1].trim() : 'Cấu trúc cây' // Fallback nếu không match
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

    if (leaf_in_tree === '0') {
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
    } else {
      // ✅ Không có ITEM → node tree cuối là leaf luôn
      parent.leaf = true

      for (const [key, val] of Object.entries(row)) {
        if (!key.startsWith('tree_lv')) {
          parent[key] = val
        }
      }

      // ❗ không cần children nữa
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

function customValueComparator(orderList) {
  const orderMap = new Map(orderList.map((v, i) => [v, i]))

  return (a, b) => {
    const ia = orderMap.has(a.name) ? orderMap.get(a.name) : Infinity
    const ib = orderMap.has(b.name) ? orderMap.get(b.name) : Infinity
    return ia - ib
  }
}

// ======================
// 3️⃣ Flatten tree (để hiển thị)
// ======================
// function flattenTree(nodes) {
//   let result = []
//   for (const n of nodes) {
//     result.push(n)
//     if (n.expanded && n.children) {
//       result = result.concat(flattenTree(n.children))
//     }
//   }
//   return result
// }
function flattenTree(nodes, level = 1) {
  let result = []

  const orderList = levelSortRules?.[level]
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

// function copySelectedRows() {
//   const selectedNodes = []
//   gridApi.forEachNodeAfterFilterAndSort((node) => {
//     if (node.isSelected()) selectedNodes.push(node)
//   })

//   if (selectedNodes.length === 0) {
//     alert('⚠️ Chưa chọn dòng nào!')
//     return
//   }

//   // Lấy danh sách cột đang hiển thị trên giao diện
//   const displayedCols = gridApi.getColumnDefs().map((c) => c.field)

//   // Build text cần copy
//   const text = selectedNodes
//     .map((node) => {
//       return displayedCols
//         .map((col) => {
//           let v = node.data[col]

//           if (v === null || v === undefined) return ''
//           if (typeof v === 'object') return '' // tránh [object Object]

//           return v.toString()
//         })
//         .join('\t')
//     })
//     .join('\n')

//   // Copy vào clipboard
//   const textarea = document.createElement('textarea')
//   textarea.value = text
//   textarea.style.position = 'fixed'
//   textarea.style.top = '-9999px'
//   document.body.appendChild(textarea)
//   textarea.focus()
//   textarea.select()

//   try {
//     document.execCommand('copy')
//   } catch (err) {
//     alert('❌ Không thể copy.')
//   }

//   document.body.removeChild(textarea)
// }

function copySelectedRows() {
  const selectedNodes = []
  gridApi.forEachNodeAfterFilterAndSort((node) => {
    if (node.isSelected()) selectedNodes.push(node)
  })

  if (selectedNodes.length === 0) {
    alert('⚠️ Chưa chọn dòng nào!')
    return
  }

  // Lấy danh sách cột đang hiển thị (flattened, bao gồm các cột con)
  const displayedCols = gridApi.getAllDisplayedColumns()

  // Build text cần copy
  const text = selectedNodes
    .map((node) => {
      return displayedCols
        .map((col) => {
          let v = node.data[col.getColId()]

          if (v === null || v === undefined) return ''
          if (typeof v === 'object') return '' // tránh [object Object]

          return v.toString()
        })
        .join('\t')
    })
    .join('\n')

  // Copy vào clipboard
  const textarea = document.createElement('textarea')
  textarea.value = text
  textarea.style.position = 'fixed'
  textarea.style.top = '-9999px'
  document.body.appendChild(textarea)
  textarea.focus()
  textarea.select()

  try {
    document.execCommand('copy')
  } catch (err) {
    alert('❌ Không thể copy.')
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
  const padding = 0 // tổng trên + dưới
  const extraSpacing = -5 // khoảng cách phụ nếu có

  const gridHeight =
    totalHeight - toolbarHeight - notebarHeight - padding - extraSpacing
  gridContainer.style.height = `${gridHeight}px`
}

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
  // const headers = [...dimensionIdxs.map((i) => cols[i]), ...measureNames]
  const headers = [
    ...dimensionIdxs.map((i) => cols[i]),
    ...measureNames
  ].filter((h) => h !== undefined && h !== null)

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
        const treeHeaderName = extractTreeHeaderName(h)
        return {
          headerName: treeHeaderName,
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

function pivotData(
  list_columns,
  list_rows,
  list_column_horizontal,
  list_column_vertical,
  list_column_measure
) {
  // B1: convert row array → object {colName: value}
  const rows = list_rows.map((rowArr) => {
    const obj = {}
    list_columns.forEach((col, i) => (obj[col] = rowArr[i]))
    return obj
  })

  const resultMap = new Map()
  // Set để lưu trữ các khóa cột ngang (vertical keys) duy nhất
  const verticalKeys = new Set()

  for (const row of rows) {
    // Khóa của hàng dọc (horizontal)
    const horizontalKey = list_column_horizontal.map((c) => row[c]).join('||')

    // Khóa của cột ngang (vertical) - đây là tên cột mới
    const verticalKey = list_column_vertical.map((c) => row[c]).join('_')
    verticalKeys.add(verticalKey) // Thêm vào danh sách cột measure mới

    // ⭐ THAY ĐỔI: Chuyển đổi Measure value sang kiểu Number
    let measureValue = row[list_column_measure]
    if (measureValue !== null && measureValue !== undefined) {
      measureValue = Number(measureValue)
      // Nếu kết quả không phải là số (NaN), bạn có thể chọn đặt nó thành 0
      // hoặc giữ nguyên NaN, tùy thuộc vào yêu cầu xử lý lỗi của bạn.
      // Ở đây, tôi sẽ giữ nguyên Number(measureValue)
    }

    // Nếu chưa có dòng này trong kết quả → tạo
    if (!resultMap.has(horizontalKey)) {
      const base = {}
      list_column_horizontal.forEach((c) => (base[c] = row[c]))
      resultMap.set(horizontalKey, base)
    }

    // Gán giá trị measure vào đúng cột
    const target = resultMap.get(horizontalKey)
    target[verticalKey] = measureValue
  }

  // Danh sách các cột dimension (chiều)
  const dimensionColumns = [...list_column_horizontal]

  // Danh sách các cột measure (đo lường)
  const measureColumns = Array.from(verticalKeys)

  // Trả về đối tượng chứa cả dữ liệu đã pivot và danh sách các cột
  return {
    pivotData: Array.from(resultMap.values()),
    dimensionColumns: dimensionColumns,
    measureColumns: measureColumns
  }
}

// Load lại dữ liệu và render
function loadAndRender(worksheet) {
  worksheet.getSummaryDataAsync({ maxRows: 0 }).then((sumData) => {
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

      const numericCols = pivotDataOutput.measureColumns

      const totals = calcTotalsTree(allData, numericCols)

      const totalRow = {}
      agGridColumnDefs_flat.forEach((col) => {
        const field = col.field

        if (numericCols.includes(field)) {
          totalRow[field] = totals[field]
        } else if (field === agGridColumnDefs_flat[0].field) {
          totalRow[field] = 'Grand Total'
        } else {
          totalRow[field] = ''
        }
      })

      totalRow.leaf = true

      gridApi.setGridOption('pinnedBottomRowData', [totalRow])
    }

    if (showGrandTotal === 1) {
      funcTionWait4ToUpdateTotal(1000)
    }

    const cols = sumData.columns.map((c) => c.fieldName)
    const rows = sumData.data.map((r) =>
      r.map((c) => {
        // console.log('c', c)

        if (c.nativeValue === null || c.nativeValue === undefined) return ''

        // 🔹 Nếu là kiểu ngày hợp lệ (Date object hoặc chuỗi ngày)
        if (c.nativeValue instanceof Date) {
          // Định dạng dd/MM/yyyy có thêm số 0
          return c.nativeValue.toLocaleDateString('vi-VN', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
          })
        } else if (typeof c.nativeValue === 'number') {
          return c.nativeValue
        } else {
          return c.formattedValue
        }
      })
    )

    // Pivot table
    pivotDataOutput = pivotData(
      cols,
      rows,
      list_column_horizontal,
      list_column_vertical,
      list_column_measure
    )

    // console.log('pivotDataOutput', pivotDataOutput)

    // 5. Build tree

    nestedData = buildTree(pivotDataOutput.pivotData)
    // console.log('nestedData', nestedData)

    // ✅ Gọi hàm cộng dồn giá trị
    aggregateTreeValues(nestedData, pivotDataOutput.measureColumns)

    // 6. Flat tree
    let flatData = flattenTree(nestedData)
    console.log('flatData', flatData)

    maxTreeLevel = getMaxTreeLevel(nestedData)
    currentExpandedLevel = 1 // ban đầu chỉ hiển thị root

    console.log('formated_columns 2', formated_columns)

    // 7. Build cấu hình để truyền vào AG Grid
    agGridColumnDefs = createColumnDefs(
      pivotDataOutput.dimensionColumns,
      pivotDataOutput.measureColumns,
      pivot_column_config,
      list_exclude_column_config,
      formated_columns
    )

    console.log('agGridColumnDefs', agGridColumnDefs)

    // ======================
    // 6️⃣ Cấu hình AG Grid
    // ======================
    const gridOptions = {
      columnDefs: agGridColumnDefs,
      rowData: flatData,
      defaultColDef: {
        filter: false, // chuyển sang false vì ko dùng filter nữa
        sortable: true,
        resizable: true
      },
      suppressFieldDotNotation: true,
      // 🔹 Làm nổi bật các dòng tổng (cha)
      getRowStyle: (params) => {
        const node = params.data
        if (!node) return null

        // ✅ Nếu là dòng "Grand Total"
        // Dòng Grand Total
        if (node.name === 'Grand Total') {
          return {
            fontWeight: 'bold',
            color: '#d00000',
            backgroundColor: '#fabcbcff'
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
      getRowHeight: (params) => {
        if (params.data && params.data.name === 'Grand Total') return 25 // Hoặc 'auto'
        return undefined // Mặc định
      },

      rowSelection: {
        mode: 'multiRow',
        checkboxes: true,
        enableClickSelection: false
      },

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
        if (showGrandTotal === 1) {
          funcTionWait4ToUpdateTotal(1000)
        }
        console.log('run onGridReady.')
      },
      onFirstDataRendered: (params) => {
        if (showGrandTotal === 1) {
          funcTionWait4ToUpdateTotal(1000)
        }
        console.log('run onFirstDataRendered.')
      },
      onCellContextMenu: (params) => {
        const node = params.data
        if (!node || !node.id) return

        params.event.preventDefault()

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
      gridApi.setGridOption('columnDefs', agGridColumnDefs)
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
          setTimeout(() => {
            gridApi.setGridOption('rowData', flat)

            // === GIỐNG toggleNode() ===
            if (targetId) {
              requestAnimationFrame(() => {
                const idx = flat.findIndex((r) => r.id == targetId)
                const rowNode = gridApi.getDisplayedRowAtIndex(idx)
                if (rowNode) {
                  gridApi.ensureNodeVisible(rowNode, 'middle')
                }
              })
            }
          }, 0) // <-- Thêm setTimeout

          currentExpandedLevel = 1
        })
      }

      expandListenersBound = true
    }

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
        gridApi.onFilterChanged()
      })

    document
      .getElementById('updateTotal')
      .addEventListener('click', updateFooterTotals)
  })
}

// Khi DOM ready
document.addEventListener('DOMContentLoaded', () => {
  tableau.extensions.initializeAsync().then(() => {
    // 1. Nhận dữ liệu từ Worksheet
    const worksheet =
      tableau.extensions.dashboardContent.dashboard.worksheets.find(
        (ws) => ws.name === 'DataTableExtSheet'
      )
    if (!worksheet) {
      console.error("❌ Không tìm thấy worksheet tên 'DataTableExtSheet'")
      return
    }

    // 2. Nhận dữ liệu cấu hình
    const worksheetConfig =
      tableau.extensions.dashboardContent.dashboard.worksheets.find(
        (ws) => ws.name === 'ConfigSheet'
      )
    if (!worksheetConfig) {
      console.error("❌ Không tìm thấy worksheet tên 'worksheetConfig'")
      return
    }

    worksheetConfig.getSummaryDataAsync({ maxRows: 0 }).then((configData) => {
      console.log('configData.data', configData.data)

      configData.data.forEach((item) => {
        switch (item[0].formattedValue) {
          case 'list_column_horizontal':
            list_column_horizontal = item[1].formattedValue.split(',')
            console.log('list_column_horizontal', list_column_horizontal)
            break
          case 'list_column_vertical':
            list_column_vertical = item[1].formattedValue.split(',')
            break
          case 'list_column_measure':
            list_column_measure = item[1].formattedValue.split(',')
            break
          case 'pivot_column_config':
            pivot_column_config = JSON.parse(item[1].formattedValue)
            break
          case 'list_exclude_column_config':
            list_exclude_column_config = item[1].formattedValue.split(',')
            break
          case 'order_lv1':
            order_lv1 = item[1].formattedValue
            break
          case 'order_lv2':
            order_lv2 = item[1].formattedValue
            break
          case 'formated_columns':
            formated_columns = item[1].formattedValue
            break
          case 'leaf_in_tree':
            leaf_in_tree = item[1].formattedValue
            break
          case 'level_sort_rules':
            levelSortRules = JSON.parse(item[1].formattedValue)
            break
          case 'show_grand_total':
            showGrandTotal = item[1].formattedValue
            break
        }
      })

      console.log('list_column_horizontal', list_column_horizontal)
      console.log('list_column_vertical', list_column_vertical)
      console.log('list_column_measure', list_column_measure)
      console.log('pivot_column_config', pivot_column_config)
      console.log('list_exclude_column_config', list_exclude_column_config)
      console.log('formated_columns', formated_columns)
      console.log('leaf_in_tree', leaf_in_tree)
      console.log('levelSortRules', levelSortRules)
      console.log('showGrandTotal: ', showGrandTotal)

      if (!pivot_column_config) {
        pivot_column_config = JSON.parse(
          '[{"field":"name","headerName":"Org Tree","width":150,"wrapText":true},{"field":"Dept code","headerName":"Dept code","width":150,"wrapText":true},{"field":"Amount","headerName":"Amount","width":150,"wrapText":true},{"field":"Amount Adj","headerName":"Amount Adj","width":150,"wrapText":true}]'
        )
      }
      console.log('pivot_column_config', pivot_column_config)
    })

    refreshExtractTime()

    // Load lần đầu
    loadAndRender(worksheet)

    // ======================
    // Export EXCEL -> tree với mỗi level là cột riêng (chỉ sửa phần này)
    // fix lỗi liên quan đến mất số 0 ở đầu
    // merge row trong group và format number cho measure
    // ======================
    // document.getElementById('exportExcel').addEventListener('click', () => {
    //   if (!gridApi || !nestedData || nestedData.length === 0) {
    //     alert('⚠️ Không có dữ liệu để export!')
    //     return
    //   }

    //   const maxLevelRef = { max: 0 }
    //   const exportRows = exportFlattenWithPath(nestedData, [], [], maxLevelRef)
    //   const maxTreeLevel = maxLevelRef.max

    //   const pinnedRows =
    //     gridApi.getPinnedBottomRowCount() > 0
    //       ? Array.from(
    //           { length: gridApi.getPinnedBottomRowCount() },
    //           (_, i) => gridApi.getPinnedBottomRow(i).data
    //         )
    //       : []

    //   const allExportRows = [...exportRows, ...pinnedRows]

    //   const currentColumnDefs = gridApi.getColumnDefs()
    //   const firstField = currentColumnDefs[0].field
    //   const otherCols = currentColumnDefs.slice(1).map((c) => c.field)

    //   const levelHeaders = Array.from(
    //     { length: maxTreeLevel },
    //     (_, i) => `Level ${i + 1}`
    //   )
    //   const exportHeaders = [...levelHeaders, ...otherCols]

    //   const wsData = []
    //   wsData.push(exportHeaders)

    //   allExportRows.forEach((row) => {
    //     const rowVals = []
    //     const isTotal = row[firstField] === 'Grand Total'

    //     if (isTotal) {
    //       rowVals.push('Grand Total')
    //       for (let i = 1; i < maxTreeLevel; i++) rowVals.push('')
    //     } else {
    //       const path = row.path || []
    //       for (let i = 0; i < maxTreeLevel; i++) {
    //         rowVals.push(path[i] || '')
    //       }
    //     }

    //     otherCols.forEach((col) => {
    //       let val = row[col] ?? ''
    //       if (typeof val === 'number') rowVals.push(val)
    //       else rowVals.push(val.toString())
    //     })

    //     wsData.push(rowVals)
    //   })

    //   const ws = XLSX.utils.aoa_to_sheet(wsData)
    //   ws['!merges'] = ws['!merges'] || []

    //   // ======================================================
    //   // ⭐ MERGE GROUP (GIỮ NHƯ CŨ)
    //   // ======================================================
    //   for (let col = 0; col < maxTreeLevel; col++) {
    //     let start = 1
    //     for (let row = 2; row <= wsData.length; row++) {
    //       const curr = wsData[row - 1][col]
    //       const prev = wsData[row - 2][col]

    //       const isEmpty = (v) => v === null || v === undefined || v === ''

    //       if (curr !== prev || isEmpty(prev)) {
    //         if (row - 1 > start && !isEmpty(prev)) {
    //           ws['!merges'].push({
    //             s: { r: start, c: col },
    //             e: { r: row - 2, c: col }
    //           })
    //         }
    //         start = row - 1
    //       }

    //       if (row === wsData.length && row - 1 > start && !isEmpty(curr)) {
    //         ws['!merges'].push({
    //           s: { r: start, c: col },
    //           e: { r: row - 1, c: col }
    //         })
    //       }
    //     }
    //   }

    //   // ======================================================
    //   // ⭐ VERTICAL TOP ALIGN
    //   // ======================================================
    //   for (let r = 1; r < wsData.length; r++) {
    //     for (let c = 0; c < maxTreeLevel; c++) {
    //       const cell = XLSX.utils.encode_cell({ r, c })
    //       if (ws[cell]) {
    //         ws[cell].s = ws[cell].s || {}
    //         ws[cell].s.alignment = { vertical: 'top' }
    //       }
    //     }
    //   }

    //   // ======================================================
    //   // ⭐ MEASURE COLUMNS ACCOUNTING FORMAT
    //   // ======================================================
    //   const accFmt = '_(* #,##0.00_);_(* (#,##0.00)_);_(* "-"??_);_(@_)'

    //   for (let C = maxTreeLevel; C < exportHeaders.length; C++) {
    //     for (let R = 1; R < wsData.length; R++) {
    //       const ref = XLSX.utils.encode_cell({ r: R, c: C })
    //       if (!ws[ref]) continue
    //       if (typeof ws[ref].v === 'number') {
    //         ws[ref].t = 'n'
    //         ws[ref].z = accFmt
    //       }
    //     }
    //   }

    //   // ======================================================
    //   // ⭐ NEW FEATURE: BOLD CÁC DÒNG NHÓM (NON-LEAF + CHILDREN)
    //   // ======================================================
    //   for (let R = 1; R < wsData.length; R++) {
    //     const originalRow = allExportRows[R - 1]
    //     const isGroup =
    //       originalRow &&
    //       !originalRow.leaf &&
    //       originalRow.children &&
    //       originalRow.children.length > 0

    //     if (isGroup) {
    //       for (let C = 0; C < exportHeaders.length; C++) {
    //         const ref = XLSX.utils.encode_cell({ r: R, c: C })
    //         if (!ws[ref]) continue

    //         ws[ref].s = ws[ref].s || {}
    //         ws[ref].s.font = { bold: true } // ⭐ IN ĐẬM GROUP
    //       }
    //     }
    //   }

    //   // ======================================================
    //   // EXPORT
    //   // ======================================================
    //   const wb = XLSX.utils.book_new()
    //   XLSX.utils.book_append_sheet(wb, ws, 'TreeData')
    //   XLSX.writeFile(wb, 'tree_data.xlsx')

    //   console.log('✅ Export Excel OK!')
    // })

    // ======================
    // Export EXCEL -> tree với mỗi level là cột riêng (chỉ sửa phần này)
    // fix lỗi liên quan đến mất số 0 ở đầu
    // merge row trong group và format number cho measure
    // ======================
    document.getElementById('exportExcel').addEventListener('click', () => {
      if (!gridApi || !nestedData || nestedData.length === 0) {
        alert('⚠️ Không có dữ liệu để export!')
        return
      }

      const maxLevelRef = { max: 0 }
      const exportRows = exportFlattenWithPath(nestedData, [], [], maxLevelRef)
      const maxTreeLevel = maxLevelRef.max

      const pinnedRows =
        gridApi.getPinnedBottomRowCount() > 0
          ? Array.from(
              { length: gridApi.getPinnedBottomRowCount() },
              (_, i) => gridApi.getPinnedBottomRow(i).data
            )
          : []

      const allExportRows = [...exportRows, ...pinnedRows]

      const currentColumnDefs = gridApi.getColumnDefs()

      // ======================================================
      // ⭐ PHẦN SỬA: Xử lý header nhiều cấp (BỎ CỘT TREE ĐẦU TIÊN)
      // ======================================================

      // Bỏ cột đầu tiên (tree column) và chỉ lấy các cột còn lại
      const otherColumnDefs = currentColumnDefs.slice(1)

      // Lấy tất cả các leaf columns (cột cuối cùng không có children)
      const getLeafColumns = (cols, parentPath = []) => {
        const leaves = []

        cols.forEach((col) => {
          if (!col.children || col.children.length === 0) {
            // Nếu là leaf column, lấy đường dẫn đầy đủ
            const fullPath = [...parentPath, col.headerName || col.field || '']
            leaves.push({
              field: col.field,
              headerPath: fullPath,
              colDef: col
            })
          } else {
            // Nếu có children, đệ quy để lấy leaf columns
            const childLeaves = getLeafColumns(col.children, [
              ...parentPath,
              col.headerName || ''
            ])
            leaves.push(...childLeaves)
          }
        })

        return leaves
      }

      const leafColumns = getLeafColumns(otherColumnDefs)

      // Tạo headers: chỉ cần level headers + các leaf columns
      const headerRows = []

      // Tìm max depth của header
      const maxHeaderDepth =
        leafColumns.length > 0
          ? Math.max(...leafColumns.map((col) => col.headerPath.length))
          : 0

      // Tạo từng hàng header từ trên xuống
      for (let depth = 0; depth < maxHeaderDepth; depth++) {
        const headerRow = []

        // Thêm các cột tree levels (bỏ cột tree gốc)
        for (let i = 0; i < maxTreeLevel; i++) {
          if (depth === 0) {
            headerRow.push(`Level ${i + 1}`)
          } else {
            headerRow.push('')
          }
        }

        // Thêm các cột leaf columns
        leafColumns.forEach((col) => {
          if (depth < col.headerPath.length) {
            headerRow.push(col.headerPath[depth])
          } else {
            headerRow.push('')
          }
        })

        headerRows.push(headerRow)
      }

      const wsData = [...headerRows]

      // Thêm dữ liệu từng dòng
      allExportRows.forEach((row) => {
        const rowVals = []

        // Kiểm tra xem có phải Grand Total không (không có cột tree gốc nữa)
        // Lấy field của cột đầu tiên để check Grand Total
        const firstField = currentColumnDefs[0].field
        const isTotal = row[firstField] === 'Grand Total'

        if (isTotal) {
          // Đối với Grand Total: chỉ cần "Grand Total" ở cột đầu tiên
          rowVals.push('Grand Total')
          // Các cột level còn lại để trống
          for (let i = 1; i < maxTreeLevel; i++) {
            rowVals.push('')
          }
        } else {
          // Thêm các giá trị từ path (đã được duỗi thành các level)
          const path = row.path || []
          for (let i = 0; i < maxTreeLevel; i++) {
            rowVals.push(path[i] || '')
          }
        }

        // Thêm giá trị cho các leaf columns
        leafColumns.forEach((colDef) => {
          let val = row[colDef.field] ?? ''
          if (typeof val === 'number') {
            rowVals.push(val)
          } else {
            rowVals.push(val.toString())
          }
        })

        wsData.push(rowVals)
      })

      const ws = XLSX.utils.aoa_to_sheet(wsData)
      ws['!merges'] = ws['!merges'] || []

      // ======================================================
      // ⭐ MERGE HEADER MULTI-LEVEL (PHẦN MỚI)
      // ======================================================
      if (maxHeaderDepth > 1) {
        // Bắt đầu từ cột level cuối cùng + 1 (vì đã bỏ cột tree gốc)
        const startDataCol = maxTreeLevel

        // Merge các header cells theo chiều dọc (vertical merge)
        for (
          let col = startDataCol;
          col < startDataCol + leafColumns.length;
          col++
        ) {
          for (let row = 0; row < maxHeaderDepth; row++) {
            // Tìm tất cả các cells có cùng parent header ở hàng trên
            if (row < maxHeaderDepth - 1) {
              const currentHeader = wsData[row][col]
              if (currentHeader && currentHeader.trim() !== '') {
                // Đếm số rows bên dưới có cùng giá trị
                let mergeCount = 1
                for (
                  let nextRow = row + 1;
                  nextRow < maxHeaderDepth;
                  nextRow++
                ) {
                  if (
                    wsData[nextRow][col] === '' &&
                    wsData[nextRow][col - 1] === ''
                  ) {
                    mergeCount++
                  } else {
                    break
                  }
                }

                if (mergeCount > 1) {
                  ws['!merges'].push({
                    s: { r: row, c: col },
                    e: { r: row + mergeCount - 1, c: col }
                  })
                }
              }
            }
          }
        }

        // Merge các header cells theo chiều ngang (horizontal merge) cho cùng level
        for (let row = 0; row < maxHeaderDepth; row++) {
          let startCol = startDataCol
          let currentHeader = wsData[row][startCol]

          for (
            let col = startDataCol + 1;
            col < startDataCol + leafColumns.length;
            col++
          ) {
            if (wsData[row][col] === currentHeader && currentHeader !== '') {
              // Tiếp tục tìm cho đến khi khác
              if (col === startDataCol + leafColumns.length - 1) {
                ws['!merges'].push({
                  s: { r: row, c: startCol },
                  e: { r: row, c: col }
                })
              }
            } else {
              if (col - 1 > startCol) {
                ws['!merges'].push({
                  s: { r: row, c: startCol },
                  e: { r: row, c: col - 1 }
                })
              }
              startCol = col
              currentHeader = wsData[row][col]
            }
          }
        }
      }

      // ======================================================
      // ⭐ MERGE GROUP (GIỮ NHƯ CŨ) - CHỈ MERGE CÁC CỘT LEVEL
      // ======================================================
      for (let col = 0; col < maxTreeLevel; col++) {
        let start = maxHeaderDepth // Bắt đầu từ sau header rows
        for (let row = maxHeaderDepth + 2; row <= wsData.length; row++) {
          const curr = wsData[row - 1][col]
          const prev = wsData[row - 2][col]

          const isEmpty = (v) => v === null || v === undefined || v === ''

          if (curr !== prev || isEmpty(prev)) {
            if (row - 1 > start && !isEmpty(prev)) {
              ws['!merges'].push({
                s: { r: start, c: col },
                e: { r: row - 2, c: col }
              })
            }
            start = row - 1
          }

          if (row === wsData.length && row - 1 > start && !isEmpty(curr)) {
            ws['!merges'].push({
              s: { r: start, c: col },
              e: { r: row - 1, c: col }
            })
          }
        }
      }

      // ======================================================
      // ⭐ MEASURE COLUMNS ACCOUNTING FORMAT
      // ======================================================
      const accFmt = '_(* #,##0.00_);_(* (#,##0.00)_);_(* "-"??_);_(@_)'

      // Chỉ format các cột measure (bắt đầu từ sau các cột level)
      const startMeasureCol = maxTreeLevel
      for (
        let C = startMeasureCol;
        C < startMeasureCol + leafColumns.length;
        C++
      ) {
        for (let R = maxHeaderDepth; R < wsData.length; R++) {
          const ref = XLSX.utils.encode_cell({ r: R, c: C })
          if (!ws[ref]) continue
          if (typeof ws[ref].v === 'number') {
            ws[ref].t = 'n'
            ws[ref].z = accFmt
          }
        }
      }

      // ======================================================
      // EXPORT
      // ======================================================
      const wb = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wb, ws, 'TreeData')

      // Tạo tên file với timestamp
      const timestamp = new Date()
        .toISOString()
        .replace(/[:.]/g, '-')
        .slice(0, 19)
      const filename = `tree_data_${timestamp}.xlsx`

      XLSX.writeFile(wb, filename)

      console.log('✅ Export Excel OK!')
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
  })
})
