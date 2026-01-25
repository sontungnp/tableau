'use strict'

document.addEventListener('DOMContentLoaded', () => {
  tableau.extensions.initializeAsync().then(() => {
    console.log('Extension initialized')

    let popupData = {}
    let arrAllAccount = []
    // let objInitSelectedAccounts = {}
    let objInitSelectedAccounts = {
        action: 'INIT',
        arrSelectedItems: [],
        strSelectedItems: 'ALL',
        isAll: true
      }
    let objSelectedAccounts = {}
    const dashboard = tableau.extensions.dashboardContent.dashboard
    let worksheets = dashboard.worksheets
    const worksheetName = 'AccCodeSheet' // Tên worksheet cần lấy
    const filterField = 'filter_accountcode' // 🔴 Đổi tên filter nếu cần

    fetchData()

    objSelectedAccounts = JSON.parse(
      localStorage.getItem('objSelectedAccounts')
    )

    console.log('load tu localstorage: ', objSelectedAccounts)

    if (
      !objSelectedAccounts || // null hoặc undefined
      Object.keys(objSelectedAccounts).length === 0 || // object rỗng {}
      !objSelectedAccounts.strSelectedItems // thiếu trường quan trọng
    ) {
      // khởi tạo giá trị lần đầu load extension lên
      objInitSelectedAccounts = {
        action: 'INIT',
        arrSelectedItems: [],
        strSelectedItems: 'ALL',
        isAll: true
      }
      objSelectedAccounts = objInitSelectedAccounts

      localStorage.setItem(
        'objSelectedAccounts',
        JSON.stringify(objSelectedAccounts)
      )
      localStorage.setItem('accountCode', objSelectedAccounts.selectedCodes)
    }

    document.getElementById('selected-box').value =
      objSelectedAccounts.strSelectedItems

    document.getElementById('dropdown-toggle').addEventListener('click', () => {
      let popupUrl = window.location.href + 'popup.html' // URL của file popup

      popupData = {
        arrAllAccount: arrAllAccount,
        objSelectedAccounts: objSelectedAccounts
      }

      tableau.extensions.ui
        .displayDialogAsync(popupUrl, JSON.stringify(popupData), {
          width: 600,
          height: 800
        })
        .then((payload) => {
          let receivedValue = JSON.parse(payload)
          if (receivedValue.action === 'ok') {
            objSelectedAccounts = {
              action: 'OK',
              arrSelectedItems: receivedValue.arrSelectedItems,
              strSelectedItems: receivedValue.strSelectedItems,
              isAll: receivedValue.isAll
            }

            console.log('payload nhan ve: ', objSelectedAccounts)

            localStorage.setItem(
              'objSelectedAccounts',
              JSON.stringify(objSelectedAccounts)
            )
            localStorage.setItem(
              'accountCode',
              objSelectedAccounts.strSelectedItems
            )

            document.getElementById('selected-box').value =
              objSelectedAccounts.strSelectedItems

            setFilterAccCode(
              objSelectedAccounts.arrSelectedItems,
              objSelectedAccounts.isAll
            )
          } else {
            console.log('Calcel')
          }
        })
        .catch((error) => {
          console.log('Lỗi khi mở popup: ' + error.message)
        })
    })

    document.getElementById('clear').addEventListener('click', clearAccFilters)

    // check thay đổi lcalstorage do nut reset từ extension khác
    // window.addEventListener('storage', function (event) {
    //   if (event.key === 'accountCode') {
    //     console.log('accountCode đã thay đổi:', event.newValue)
    //     if (event.newValue === null || event.newValue === 'ALL') {
    //       objSelectedAccounts = objInitSelectedAccounts
    //       localStorage.setItem(
    //         'objSelectedAccounts',
    //         JSON.stringify(objSelectedAccounts)
    //       )
    //       localStorage.setItem(
    //         'accountCode',
    //         objSelectedAccounts.strSelectedItems
    //       )
    //     } else {
    //       objSelectedAccounts.strSelectedItems = event.newValue
    //     }

    //     document.getElementById('selected-box').value =
    //       objSelectedAccounts.strSelectedItems
    //   }
    // })

    function fetchData() {
      // const worksheet = tableau.extensions.dashboardContent.dashboard.worksheets[0];
      const worksheet = worksheets.find((ws) => ws.name === worksheetName)
      worksheet.getSummaryDataAsync().then((data) => {
        arrAllAccount = data.data
      })
    }

    async function setFilterAccCode(lstAccCode, isAll) {
      try {
        await Promise.allSettled(
          worksheets.map(async (ws) => {
            // 🔹 Lấy danh sách filters hiện có trên worksheet
            let filters = await ws.getFiltersAsync()

            // Tìm xem worksheet có filter này không -> nếu không có thì bỏ qua
            if (!filters.some((f) => f.fieldName === filterField)) {
              console.warn(
                `Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`
              )
              return
            }

            if (isAll === 'ALL') {
              // 🔹 Nếu filterValue rỗng hoặc là "ALL" => Clear filter
              document.getElementById('selected-box').value = 'ALL'
              await ws.clearFilterAsync(filterField)
            } else {
              // 🔹 Kiểm tra nếu filterValue là một mảng thì truyền mảng, nếu không thì truyền giá trị đơn lẻ
              console.log('vao day roi  nhe')
              await ws.applyFilterAsync(filterField, lstAccCode, 'replace')
            }
          })
        )

        // alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
      } catch (error) {
        console.error('Error setting filter:', error)
        alert('Failed to set filter. Check console for details.')
      }
    }

    async function clearAccFilters() {
      // thiết lập giá trị khởi tạo ban đầu
      objSelectedAccounts = objInitSelectedAccounts

      localStorage.setItem(
        'objSelectedAccounts',
        JSON.stringify(objSelectedAccounts)
      )
      localStorage.setItem('accountCode', objSelectedAccounts.selectedCodes)

      document.getElementById('selected-box').value = 'ALL'

      try {
        for (const ws of worksheets) {
          // 🔹 Lấy danh sách filters hiện có trên worksheet
          let filters = await ws.getFiltersAsync()

          // Tìm xem worksheet có filter này không -> nếu ko có thì continue sang worksheet khác
          let hasFilter = filters.some((f) => f.fieldName === filterField)

          if (!hasFilter) {
            console.warn(
              `Worksheet "${ws.name}" does not have filter "${filterField}". Skipping...`
            )
            continue
          } else {
            await ws.clearFilterAsync(filterField)
          }
        }

        // alert(`Filter "${filterField}" set to: ${filterValue} on all worksheets`);
      } catch (error) {
        console.error('Error clear filter:' + filterField, error)
      }
    }
  })
})
