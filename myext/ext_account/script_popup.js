'use strict'

tableau.extensions.initializeDialogAsync().then(async (payload) => {
  // Sử dụng async ở đây
  let arrSelectedItems = []
  let strSelectedItems = ''
  let isALL = 'NOTALL'

  let popupData = JSON.parse(payload)
  console.log('popupData:', popupData)
  let arrAllAccount = popupData.arrAllAccount
  let objSelectedAccounts = popupData.objSelectedAccounts
  arrSelectedItems = objSelectedAccounts.arrSelectedItems
  strSelectedItems = objSelectedAccounts.strSelectedItems

  let container = document.getElementById('acc-table-container')
  let selected_item_list_detail = document.getElementById('selected-detail-box')
  selected_item_list_detail.value = strSelectedItems

  container.style.display = 'block'

  renderTable(arrAllAccount)

  let checkAllBox = document.getElementById('checkAll')
  let checkShowSelectedItems = document.getElementById('show-selected-items-ck')

  checkShowSelectedItems.addEventListener('change', function () {
    filterSelectedAcc()
  })

  // ===== Vẽ bảng từ arrAllAccount =====
  function renderTable(data) {
    console.log('=========data', data)

    let tableHtml = `
    <table border="1" cellspacing="0" cellpadding="4" style="width:100%; border-collapse: collapse;">
      <thead>
        <tr style="background:#f0f0f0;">
          <th style="width:40px;"><input type="checkbox" id="checkAll"></th>
          <th>Code</th>
          <th>Name</th>
        </tr>
      </thead>
      <tbody>
        ${data
          .map((row) => {
            let code = row[0]._value
            let isChecked = arrSelectedItems.includes(code) ? 'checked' : ''
            return `
              <tr>
                <td style="text-align:center;">
                  <input type="checkbox" class="row-check" data-code="${code}" ${isChecked}>
                </td>
                <td>${row[0]._value}</td>
                <td>${row[1]._value}</td>
              </tr>
            `
          })
          .join('')}
      </tbody>
    </table>
  `
    container.innerHTML = tableHtml

    let checkAll = document.getElementById('checkAll')
    let allChecks = container.querySelectorAll('.row-check')

    // --- Set trạng thái checkAll + isALL ban đầu ---
    checkAll.checked = Array.from(allChecks).every((chk) => chk.checked)
    isALL =
      arrSelectedItems.length === 0 ||
      arrSelectedItems.length === arrAllAccount.length
        ? 'ALL'
        : 'NOTALL'

    // --- Check All ---
    checkAll.addEventListener('change', function () {
      allChecks.forEach((chk) => {
        chk.checked = this.checked
        updateSelectedItems(chk.dataset.code, chk.checked)
      })
      // cập nhật biến isALL
      isALL =
        arrSelectedItems.length === 0 ||
        arrSelectedItems.length === arrAllAccount.length
          ? 'ALL'
          : 'NOTALL'
    })

    // --- Check từng dòng ---
    allChecks.forEach((chk) => {
      chk.addEventListener('change', function () {
        updateSelectedItems(this.dataset.code, this.checked)

        // Cập nhật lại trạng thái checkAll
        checkAll.checked = Array.from(allChecks).every((chk) => chk.checked)
        isALL =
          arrSelectedItems.length === 0 ||
          arrSelectedItems.length === arrAllAccount.length
            ? 'ALL'
            : 'NOTALL'
      })
    })
  }

  function updateSelectedItems(code, isChecked) {
    if (isChecked) {
      if (!arrSelectedItems.includes(code)) {
        arrSelectedItems.push(code)
      }
    } else {
      arrSelectedItems = arrSelectedItems.filter((item) => item !== code)
    }

    showNewListDetailItem()
  }

  function showNewListDetailItem() {
    strSelectedItems = arrSelectedItems.join(',')
    selected_item_list_detail.value = strSelectedItems
  }

  // ====== Search Filter ======
  document.getElementById('search-box').addEventListener('input', function () {
    // Bỏ tích checkbox nếu đang được tích
    if (checkShowSelectedItems.checked) {
      checkShowSelectedItems.checked = false
      // Kích hoạt sự kiện change nếu cần
      checkShowSelectedItems.dispatchEvent(new Event('change'))
    }

    // Gọi hàm filterAcc
    filterAcc()
  })

  function normalizeStr(str) {
    return str
      .normalize('NFKD') // Chuẩn hóa Unicode
      .replace(/[\u0300-\u036f]/g, '') // Bỏ dấu tổ hợp (nếu cần cho tiếng Việt)
      .toLowerCase()
      .trim()
  }

  function filterAcc() {
    let keyword = normalizeStr(document.getElementById('search-box').value)
    let filtered = arrAllAccount.filter(
      (row) =>
        normalizeStr(row[0]._value).includes(keyword) ||
        normalizeStr(row[1]._value).includes(keyword)
    )
    renderTable(filtered)
  }

  function filterSelectedAcc() {
    if (document.getElementById('show-selected-items-ck').checked) {
      let filtered = arrAllAccount.filter((row) =>
        arrSelectedItems.includes(row[0]._value)
      )
      renderTable(filtered)
    } else {
      renderTable(arrAllAccount)
    }
  }

  document.getElementById('okPopup').addEventListener('click', () => {
    returnData('ok')
  })

  document.getElementById('closePopup').addEventListener('click', function () {
    returnData('cancel')
  })

  function returnData(action) {
    let selectedCodes = document.getElementById('selected-box').value
    if (!selectedCodes || selectedCodes.trim() === '') {
      selectedCodes = 'ALL'
    }

    let returnValues = {
      action: action,
      arrSelectedItems: arrSelectedItems,
      strSelectedItems: strSelectedItems,
      isAll: isALL
    }

    tableau.extensions.ui.closeDialog(JSON.stringify(returnValues))
  }

  // ====== Tick checkbox theo code nhập ======
  document.getElementById('checking-buttons').addEventListener('click', () => {
    if (checkAllBox.checked) {
      checkAllBox.checked = false // bỏ check
      // Trigger lại sự kiện 'change'
      checkAllBox.dispatchEvent(new Event('change'))
    }
    tickNodeByTypingCode()

    // chỉ show những item được chọn thôi.
    if (!checkShowSelectedItems.checked) {
      checkShowSelectedItems.checked = true // tích checkbox
      checkShowSelectedItems.dispatchEvent(new Event('change')) // kích hoạt sự kiện change
    }
  })

  function tickNodeByTypingCode() {
    let inputValue = document.getElementById('selected-box').value.trim()
    if (!inputValue) return

    let unitCodes = inputValue
      .split(',')
      .map((code) => code.trim().toLowerCase())

    // Lọc arrAllAccount theo pattern dựa trên row[0]._value
    let matchedCodes = arrAllAccount
      .map((row) => row[0]._value) // lấy tất cả code từ row[0]._value
      .filter((codeVal) => {
        let lowerCode = codeVal.toLowerCase()
        return unitCodes.some((pattern) => {
          if (pattern.includes('%')) {
            // LIKE pattern: đổi % thành regex .*
            let regexPattern = '^' + pattern.replace(/%/g, '.*') + '$'
            return new RegExp(regexPattern, 'i').test(lowerCode)
          } else {
            return lowerCode === pattern
          }
        })
      })

    // Cập nhật arrSelectedItems dựa trên matchedCodes
    arrAllAccount.forEach((row) => {
      let code = row[0]._value
      updateSelectedItems(code, matchedCodes.includes(code))
    })

    filterSelectedAcc()
  }
})
