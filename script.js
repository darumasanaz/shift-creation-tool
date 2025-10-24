
document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  // --- 1. HTMLの要素をすべて取得 ---
  const staffNameInput = document.getElementById('staff-name');
  const addStaffButton = document.getElementById('add-staff-button');
  const staffList = document.getElementById('staff-list');
  
  const staffSelect = document.getElementById('dayoff-staff');
  const dayoffDateInput = document.getElementById('dayoff-date');
  const addDayoffButton = document.getElementById('add-dayoff-button');
  const dayoffList = document.getElementById('dayoff-list');

  const yearSelect = document.getElementById('target-year');
  const monthSelect = document.getElementById('target-month');

  const generateButton = document.getElementById('generate-btn');

  // --- 2. アプリケーションの状態を管理する場所 ---
  const state = {
    staff: [],
    dayoffs: [],
  };

  // --- 3. すべての関数をここで定義 ---

  /** スタッフを追加する関数 */
  function addStaff(event) {
    event.preventDefault();
    const name = staffNameInput.value.trim();
    if (name && !state.staff.includes(name)) { // 重複もチェック
      state.staff.push(name);
      renderStaffList();
      staffNameInput.value = '';
    }
  }

  /** スタッフリストとプルダウンを画面に描画する関数 */
  function renderStaffList() {
    staffList.innerHTML = '';
    staffSelect.innerHTML = '<option value="">スタッフを選択</option>';
    state.staff.forEach(name => {
      const li = document.createElement('li');
      li.textContent = name;
      staffList.appendChild(li);
      const option = document.createElement('option');
      option.value = name;
      option.textContent = name;
      staffSelect.appendChild(option);
    });
  }

  /** 希望休を登録する関数 */
  function addDayoff(event) {
    event.preventDefault();
    const staffName = staffSelect.value;
    const date = dayoffDateInput.value;
    if (staffName && date) {
      // 同じ日の希望休が重複しないようにチェック
      const isDuplicate = state.dayoffs.some(d => d.staff === staffName && d.date === date);
      if (!isDuplicate) {
        state.dayoffs.push({ staff: staffName, date: date });
        renderDayoffList();
      }
    }
  }

  /** 希望休リストを画面に描画する関数 */
  function renderDayoffList() {
    dayoffList.innerHTML = '';
    state.dayoffs.forEach(dayoff => {
      const li = document.createElement('li');
      li.textContent = `${dayoff.staff} - ${dayoff.date}`;
      dayoffList.appendChild(li);
    });
  }

  /** 年の選択肢を作成する関数 */
  function populateYears() {
    const currentYear = new Date().getFullYear();
    for (let i = -1; i <= 2; i++) {
      const year = currentYear + i;
      const option = document.createElement('option');
      option.value = year;
      option.textContent = `${year}年`;
      if (i === 0) option.selected = true;
      yearSelect.appendChild(option);
    }
  }

  /** 月の選択肢を作成する関数 */
  function populateMonths() {
    const currentMonth = new Date().getMonth() + 1;
    for (let month = 1; month <= 12; month++) {
      const option = document.createElement('option');
      option.value = month;
      option.textContent = `${month}月`;
      if (month === currentMonth) option.selected = true;
      monthSelect.appendChild(option);
    }
  }
  
  /** 結果表のヘッダーを作成する関数 */
  function renderHeader() {
    const resultTable = document.querySelector('#result-area table');
    if(!resultTable) return;
    resultTable.innerHTML = ''; 

    const year = parseInt(yearSelect.value, 10);
    const month = parseInt(monthSelect.value, 10);
    if (!year || !month) return;

    const daysInMonth = new Date(year, month, 0).getDate();
    const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    const weekdayRow = document.createElement('tr');
    headerRow.innerHTML = '<th>スタッフ</th>';
    weekdayRow.innerHTML = '<th></th>';

    for (let day = 1; day <= daysInMonth; day++) {
      const thDate = document.createElement('th');
      thDate.textContent = day;
      headerRow.appendChild(thDate);
      const date = new Date(year, month - 1, day);
      const thWd = document.createElement('th');
      thWd.textContent = weekdays[date.getDay()];
      if (date.getDay() === 0) thWd.style.color = 'red';
      if (date.getDay() === 6) thWd.style.color = 'blue';
      weekdayRow.appendChild(thWd);
    }

    thead.appendChild(headerRow);
    thead.appendChild(weekdayRow);
    resultTable.appendChild(thead);
    
    const tbody = document.createElement('tbody');
    tbody.id = 'result-body';
    resultTable.appendChild(tbody);
  }

  /** ★★★ メインのシフト生成ロジック ★★★ */
  function generateShift() {
    const tableBody = document.getElementById('result-body');
    if (!tableBody) return;
    tableBody.innerHTML = '';

    const year = parseInt(yearSelect.value, 10);
    const month = parseInt(monthSelect.value, 10);
    if (!year || !month) return;
    
    const daysInMonth = new Date(year, month, 0).getDate();

    state.staff.forEach(staffName => {
      const row = document.createElement('tr');
      const nameCell = document.createElement('td');
      nameCell.textContent = staffName;
      row.appendChild(nameCell);

      for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('td');
        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const isDayOff = state.dayoffs.some(d => d.staff === staffName && d.date === dateStr);
        
        if (isDayOff) {
          cell.textContent = '休み';
          cell.style.backgroundColor = '#ffdcdc';
        } else {
          // MVPでは、希望休以外は単純に空白にしておきます
          cell.textContent = ''; 
        }
        row.appendChild(cell);
      }
      tableBody.appendChild(row);
    });
  }

  // --- 4. 初期化とイベント設定 ---
  populateYears();
  populateMonths();
  renderHeader();

  if(addStaffButton) addStaffButton.addEventListener('click', addStaff);
  if(addDayoffButton) addDayoffButton.addEventListener('click', addDayoff);
  if(yearSelect) yearSelect.addEventListener('change', renderHeader);
  if(monthSelect) monthSelect.addEventListener('change', renderHeader);
  if(generateButton) generateButton.addEventListener('click', generateShift);
});
