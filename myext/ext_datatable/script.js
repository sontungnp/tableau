'use strict'

document.addEventListener('DOMContentLoaded', () => {
  tableau.extensions.initializeAsync().then(() => {
    let worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0]

    console.log(worksheet)

    worksheet.getSummaryDataAsync().then(function (sumData) {
      console.log('sumData', sumData)

      let columns = sumData.columns.map((col) => col.fieldName)
      let data = sumData.data

      console.log('columns', columns)

      // Detect tree columns: tree_l1, tree_l2, ...
      const treeCols = columns.filter((col) => /^tree_l\d+$/i.test(col))
      treeCols.sort((a, b) => {
        const getLevel = (s) => parseInt(s.match(/\d+/)[0])
        return getLevel(a) - getLevel(b)
      })

      // Detect Measure Names / Measure Values
      const measureNameIndex = columns.indexOf('Measure Names')
      const measureValueIndex = columns.indexOf('Measure Values')

      console.log('measureNameIndex', measureNameIndex)
      console.log('measureValueIndex', measureValueIndex)

      // Base valid (dimension-like) columns: remove AGG, hidden, and Measure Name/Value columns
      const baseValidCols = columns.filter(
        (col, idx) =>
          !col.includes('AGG') &&
          !col.startsWith('hiden') &&
          idx !== measureNameIndex &&
          idx !== measureValueIndex
      )

      // Build pivoted measure list if any
      let measureCols = []
      if (measureNameIndex !== -1 && measureValueIndex !== -1) {
        measureCols = [
          ...new Set(data.map((row) => row[measureNameIndex].formattedValue))
        ]
      }

      console.log('measureCols', measureCols)

      // === RENDERING HELPERS ===
      function addHeaderAndFilterCells(headers, filterSpecBuilder) {
        headers.forEach((h, idx) => {
          $('#table-header').append(`<th>${h}</th>`)
          const spec = filterSpecBuilder ? filterSpecBuilder(h, idx) : null
          if (spec && spec.type === 'button') {
            $('#table-filters').append(
              `<th><button class="show-filter-btn" data-coltype="${spec.coltype}" data-idx="${spec.idx}">Filter</button></th>`
            )
          } else if (spec && spec.type === 'input') {
            $('#table-filters').append(
              `<th><input type="text" class="column-filter" id="filter-${spec.idx}" placeholder="Filter ${h}" /></th>`
            )
          } else {
            $('#table-filters').append('<th></th>')
          }
        })
      }

      // === BRANCH 1: TREE VIEW (with optional pivot) ===
      if (treeCols.length > 0) {
        // Separate non-tree dimensions and measures
        let dimCols = baseValidCols.filter((c) => !treeCols.includes(c))
        let valueCols = []

        if (measureCols.length > 0) {
          valueCols = measureCols
        } else {
          valueCols = [] // no pivot, will use numeric/agg measures directly if any
        }

        // Build a hierarchical node list
        let nodeList = []
        let nodeMap = {}

        data.forEach((row) => {
          // Build full path for the row
          let fullPath = []
          treeCols.forEach((col) => {
            let val = row[columns.indexOf(col)].formattedValue
            if (val !== undefined && val !== null && `${val}`.trim() !== '') {
              fullPath.push(val)
            }
          })
          if (fullPath.length === 0) return

          // Create nodes along the path
          fullPath.forEach((label, idx) => {
            let nodeId = fullPath.slice(0, idx + 1).join('|')
            let parentId = idx === 0 ? null : fullPath.slice(0, idx).join('|')

            if (!nodeMap[nodeId]) {
              nodeMap[nodeId] = {
                id: nodeId,
                parent: parentId,
                label: label,
                level: idx + 1,
                isLeaf: idx === fullPath.length - 1,
                dimValues: {},
                measures: {}
              }
              nodeList.push(nodeMap[nodeId])
            }

            // Attach measures + dimension values at leaf level
            if (idx === fullPath.length - 1) {
              const n = nodeMap[nodeId]
              // copy dimension values
              dimCols.forEach((colName) => {
                n.dimValues[colName] =
                  row[columns.indexOf(colName)].formattedValue
              })

              if (measureCols.length > 0) {
                const mName = row[measureNameIndex].formattedValue
                const mVal = row[measureValueIndex].formattedValue
                if (mName != null && mName !== '') {
                  n.measures[mName] = mVal
                }
                valueCols.forEach((m) => {
                  if (n.measures[m] === undefined) n.measures[m] = ''
                })
              } else {
                // Non-pivot: no extra measures handled here
              }
            }
          })
        })

        // Build headers: Tree + dimCols + valueCols
        addHeaderAndFilterCells(['Tree'], (h, idx) => ({
          type: 'button',
          coltype: 'tree',
          idx: 0
        }))
        addHeaderAndFilterCells(dimCols, (h, idx) => ({
          type: 'button',
          coltype: 'dim',
          idx: idx
        }))
        addHeaderAndFilterCells(valueCols, (h, idx) => ({
          type: 'button',
          coltype: 'measure',
          idx: idx
        }))

        // Render body rows
        nodeList.forEach((node) => {
          let indent = (node.level - 1) * 20
          let toggleBtn = node.isLeaf
            ? ''
            : `<span class="toggle-btn">[+]</span> `
          let attrs = node.parent
            ? `data-parent-id="${node.parent}" style="display:none"`
            : ''
          let rowHTML = `<tr data-node-id="${node.id}" ${attrs}>`
          rowHTML += `<td style="padding-left:${indent}px">${toggleBtn}${node.label}</td>`
          dimCols.forEach((col) => {
            rowHTML += `<td>${node.dimValues[col] || ''}</td>`
          })
          valueCols.forEach((col) => {
            rowHTML += `<td style="text-align:right">${
              node.measures[col] || ''
            }</td>`
          })
          rowHTML += `</tr>`
          $('#table-body').append(rowHTML)
        })

        // Toggle handlers
        $(document).on('click', '.toggle-btn', function () {
          const $btn = $(this)
          const $row = $btn.closest('tr')
          const nodeId = $row.data('node-id')
          const isExpanded = $row.hasClass('expanded')
          if (isExpanded) {
            collapseChildren(nodeId)
            $btn.text('[+]')
            $row.removeClass('expanded')
          } else {
            expandChildren(nodeId)
            $btn.text('[−]')
            $row.addClass('expanded')
          }
        })
        function collapseChildren(parentId) {
          $(`tr[data-parent-id="${parentId}"]`).each(function () {
            let childId = $(this).data('node-id')
            $(this).hide().removeClass('expanded')
            $(this).find('.toggle-btn').text('[+]')
            collapseChildren(childId)
          })
        }
        function expandChildren(parentId) {
          $(`tr[data-parent-id="${parentId}"]`).each(function () {
            $(this).show()
          })
        }

        // Filter button → select creators (Excel-like)
        $(document).on('click', '.show-filter-btn', function () {
          let btn = $(this)
          let colType = btn.data('coltype')
          let idx = parseInt(btn.data('idx'))

          if (colType === 'tree') {
            const unique = [...new Set(nodeList.map((n) => n.label))]
            let select = `<select class="column-filter" id="filter-0" onchange="filterColumn(0)">`
            select += `<option value="">All Tree</option>`
            unique.forEach(
              (v) => (select += `<option value="${v}">${v}</option>`)
            )
            select += `</select>`
            btn.replaceWith(select)
          } else if (colType === 'dim') {
            const tableColIdx = 1 + idx
            const dName = dimCols[idx]
            const unique = [...new Set(nodeList.map((n) => n.dimValues[dName]))]
            let select = `<select class="column-filter" id="filter-${tableColIdx}" onchange="filterColumn(${tableColIdx})">`
            select += `<option value="">All ${dName}</option>`
            unique.forEach(
              (v) => (select += `<option value="${v}">${v}</option>`)
            )
            select += `</select>`
            btn.replaceWith(select)
          } else if (colType === 'measure') {
            const tableColIdx = 1 + dimCols.length + idx
            const mName = valueCols[idx]
            const unique = [...new Set(nodeList.map((n) => n.measures[mName]))]
            let select = `<select class="column-filter" id="filter-${tableColIdx}" onchange="filterColumn(${tableColIdx})">`
            select += `<option value="">All ${mName}</option>`
            unique.forEach(
              (v) => (select += `<option value="${v}">${v}</option>`)
            )
            select += `</select>`
            btn.replaceWith(select)
          }
        })
      } else {
        if (measureCols.length > 0) {
          // PIVOT MODE
          const validDimensionCols = baseValidCols

          let pivotMap = {}
          data.forEach((row) => {
            let dimKey = validDimensionCols
              .map((col) => row[columns.indexOf(col)].formattedValue)
              .join('|')

            if (!pivotMap[dimKey]) {
              pivotMap[dimKey] = {}
              validDimensionCols.forEach((col) => {
                pivotMap[dimKey][col] = row[columns.indexOf(col)].formattedValue
              })
              measureCols.forEach((m) => (pivotMap[dimKey][m] = ''))
            }

            let mName = row[measureNameIndex].formattedValue
            let mVal = row[measureValueIndex].formattedValue
            pivotMap[dimKey][mName] = mVal
          })

          // Headers + filter buttons
          addHeaderAndFilterCells(validDimensionCols, (h, idx) => ({
            type: 'button',
            coltype: 'dim',
            idx
          }))
          addHeaderAndFilterCells(measureCols, (h, idx) => ({
            type: 'button',
            coltype: 'measure-flat',
            idx
          }))

          // Body
          Object.values(pivotMap).forEach((row) => {
            let rowHTML = '<tr>'
            validDimensionCols.forEach(
              (col) => (rowHTML += `<td>${row[col]}</td>`)
            )
            measureCols.forEach(
              (m) => (rowHTML += `<td style="text-align:right">${row[m]}</td>`)
            )
            rowHTML += '</tr>'
            $('#table-body').append(rowHTML)
          })

          // Filter button → select (Excel-like)
          $(document).on('click', '.show-filter-btn', function () {
            let btn = $(this)
            let colType = btn.data('coltype')
            let idx = parseInt(btn.data('idx'))

            if (colType === 'dim') {
              // Dimension column
              let colName = validDimensionCols[idx]
              let unique = [
                ...new Set(
                  data.map(
                    (row) => row[columns.indexOf(colName)].formattedValue
                  )
                )
              ]

              let select = `<select class="column-filter" id="filter-${idx}" onchange="filterColumn(${idx})">`
              select += `<option value="">All ${colName}</option>`
              unique.forEach(
                (v) => (select += `<option value="${v}">${v}</option>`)
              )
              select += `</select>`

              btn.replaceWith(select)
            } else if (colType === 'measure-flat') {
              // Measure column
              const tableColIdx = validDimensionCols.length + idx
              const mName = measureCols[idx]
              const unique = [
                ...new Set(Object.values(pivotMap).map((row) => row[mName]))
              ]

              let select = `<select class="column-filter" id="filter-${tableColIdx}" onchange="filterColumn(${tableColIdx})">`
              select += `<option value="">All ${mName}</option>`
              unique.forEach(
                (v) => (select += `<option value="${v}">${v}</option>`)
              )
              select += `</select>`

              btn.replaceWith(select)
            }
          })

          // Filter button handler giống pivot trong code cũ...
        } else {
          // SIMPLE FLAT MODE
          console.log('vao day roi')

          columns.forEach((col) => {
            $('#table-header').append(`<th>${col}</th>`)
          })
          columns.forEach((col, index) => {
            let unique = [
              ...new Set(data.map((row) => row[index].formattedValue))
            ]
            let select = `<select class="column-filter" id="filter-${index}" onchange="filterColumn(${index})">`
            select += `<option value="">All ${col}</option>`
            unique.forEach(
              (v) => (select += `<option value="${v}">${v}</option>`)
            )
            select += `</select>`
            $('#table-filters').append(`<th>${select}</th>`)
          })
          data.forEach((row) => {
            let rowHTML = '<tr>'
            row.forEach((cell) => {
              rowHTML += `<td>${cell.formattedValue}</td>`
            })
            rowHTML += '</tr>'
            $('#table-body').append(rowHTML)
          })
        }
      }

      // === COMMON ===
      $.fn.dataTable.ext.type.search.string = function (data) {
        return !data
          ? ''
          : data
              .toString()
              .normalize('NFKD')
              .replace(/[\u0300-\u036f]/g, '')
              .toLowerCase()
              .trim()
      }

      let table = $('#data-table').DataTable({
        paging: true,
        searching: true,
        ordering: false,
        pageLength: 500,
        scrollY: '100%', // 👈 thêm dòng này
        dom: '<"top-controls"lBf>rtip',
        fixedHeader: true, // 👈 cố định thead
        lengthMenu: [
          [10, 50, 100, 500, 1000, 2000, 5000],
          [10, 50, 100, 500, 1000, 2000, 5000]
        ],
        buttons: [
          {
            extend: 'excelHtml5',
            text: 'Export to Excel',
            title: 'Exported_Data'
          }
        ]
      })

      $('#table-length').html($('.dataTables_length'))
      $('#table-search').html($('.dataTables_filter'))
      $('#table-export').html($('.dt-buttons'))

      window.filterColumn = function (index) {
        let val = $(`#filter-${index}`).val()
        table
          .column(index)
          .search(val ? `^${val}$` : '', true, false)
          .draw()
      }

      let lastSelectedRow = null
      $('#table-body').on('click', 'tr', function (event) {
        if (event.ctrlKey) {
          $(this).toggleClass('highlight')
        } else if (event.shiftKey && lastSelectedRow) {
          let rows = $('#table-body tr')
          let start = rows.index(lastSelectedRow)
          let end = rows.index(this)
          let [min, max] = [Math.min(start, end), Math.max(start, end)]
          rows.slice(min, max + 1).addClass('highlight')
        } else {
          $('#table-body tr').removeClass('highlight')
          $(this).addClass('highlight')
        }
        lastSelectedRow = this
      })

      function copySelectedRows() {
        let copiedText = ''
        $('.highlight').each(function () {
          let rowData = $(this)
            .find('td')
            .map(function () {
              return $(this).text().trim()
            })
            .get()
            .join('\t')
          copiedText += rowData + '\n'
        })

        if (copiedText) {
          let textarea = $('<textarea>')
            .val(copiedText)
            .appendTo('body')
            .select()
          document.execCommand('copy')
          textarea.remove()
          alert('Copied to clipboard!')
        } else {
          alert('No rows selected!')
        }
      }

      $(document).on('keyup', function (event) {
        if (event.ctrlKey && (event.key === 'c' || event.key === 'C')) {
          copySelectedRows()
        }
      })

      $('#copy-btn').on('click', function () {
        copySelectedRows()
      })

      let contextMenu = $(
        '<ul id="context-menu" class="context-menu"><li id="copy-selected">Copy</li></ul>'
      )
      $('body').append(contextMenu)
      $(document).on('click', function () {
        $('#context-menu').hide()
      })
      $('#table-body').on('contextmenu', 'tr.highlight', function (event) {
        event.preventDefault()
        $('#context-menu')
          .css({ top: event.pageY + 'px', left: event.pageX + 'px' })
          .show()
      })
      $('#copy-selected').on('click', function () {
        copySelectedRows()
        $('#context-menu').hide()
      })
    })
  })
})
