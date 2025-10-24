document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  const SHIFT_PATTERNS = ['早番', '日勤A', '日勤B', '遅番', '夜勤', '明け', '休み'];
  const WEEKDAY_INDEX_TO_KEY = ['sun', 'mon', 'tue', 'wed', 'thu', 'fri', 'sat'];
  const WEEKDAY_KEY_TO_INDEX = WEEKDAY_INDEX_TO_KEY.reduce((acc, key, index) => {
    acc[key] = index;
    return acc;
  }, {});

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

  const staffModal = document.getElementById('staff-modal');
  const modalTitle = document.getElementById('modal-title');
  const modalForm = document.getElementById('modal-form');
  const modalShifts = document.getElementById('modal-shifts');
  const modalWeekdays = document.getElementById('modal-weekdays');
  const modalMaxDays = document.getElementById('modal-max-days');
  const modalSaveBtn = document.getElementById('modal-save-btn');
  const modalCancelBtn = document.getElementById('modal-cancel-btn');

  const state = {
    staff: [],
    dayoffs: [],
    editingStaffId: null,
  };

  function generateStaffId() {
    return `staff-${Date.now()}-${Math.random().toString(16).slice(2)}`;
  }

  function createStaff(name) {
    return {
      id: generateStaffId(),
      name,
      availableShifts: [...SHIFT_PATTERNS],
      fixedHolidays: [],
      maxWorkingDays: null,
    };
  }

  function findStaffById(staffId) {
    return state.staff.find(staff => staff.id === staffId) || null;
  }

  function getWeekdayIndexFromKey(key) {
    return Object.prototype.hasOwnProperty.call(WEEKDAY_KEY_TO_INDEX, key)
      ? WEEKDAY_KEY_TO_INDEX[key]
      : null;
  }

  function normalizeFixedHolidays(staff) {
    if (!staff || !Array.isArray(staff.fixedHolidays)) {
      if (staff) staff.fixedHolidays = [];
      return;
    }

    const normalized = staff.fixedHolidays
      .map(value => {
        if (typeof value === 'number' && value >= 0 && value <= 6) {
          return value;
        }
        if (typeof value === 'string') {
          const numeric = Number(value);
          if (!Number.isNaN(numeric) && numeric >= 0 && numeric <= 6) {
            return numeric;
          }
          const indexFromKey = getWeekdayIndexFromKey(value);
          if (indexFromKey != null) {
            return indexFromKey;
          }
        }
        return null;
      })
      .filter(value => value != null);

    staff.fixedHolidays = Array.from(new Set(normalized));
  }

  function formatDate(year, month, day) {
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  function markCellAsDayOff(cell) {
    if (!cell) return;
    cell.textContent = '休み';
    cell.style.backgroundColor = '#ffdcdc';
  }

  function resetCell(cell) {
    if (!cell) return;
    cell.textContent = '';
    cell.style.backgroundColor = '';
  }

  function addStaff(event) {
    event.preventDefault();
    const name = staffNameInput ? staffNameInput.value.trim() : '';
    if (!name) return;

    const isDuplicate = state.staff.some(staff => staff.name === name);
    if (isDuplicate) return;

    state.staff.push(createStaff(name));
    renderStaffList();
    if (staffNameInput) {
      staffNameInput.value = '';
    }
  }

  function renderStaffList() {
    if (staffList) {
      staffList.innerHTML = '';
    }
    if (staffSelect) {
      staffSelect.innerHTML = '<option value="">スタッフを選択</option>';
    }

    state.staff.forEach(staff => {
      normalizeFixedHolidays(staff);
      if (staffList) {
        const li = document.createElement('li');
        const nameSpan = document.createElement('span');
        nameSpan.textContent = staff.name;
        li.appendChild(nameSpan);

        const editButton = document.createElement('button');
        editButton.type = 'button';
        editButton.className = 'edit-btn';
        editButton.textContent = '編集';
        editButton.setAttribute('data-staff-id', staff.id);
        li.appendChild(editButton);

        staffList.appendChild(li);
      }

      if (staffSelect) {
        const option = document.createElement('option');
        option.value = staff.id;
        option.textContent = staff.name;
        staffSelect.appendChild(option);
      }
    });
  }

  function addDayoff(event) {
    event.preventDefault();
    if (!staffSelect || !dayoffDateInput) return;

    const staffId = staffSelect.value;
    const date = dayoffDateInput.value;
    if (!staffId || !date) return;

    const isDuplicate = state.dayoffs.some(dayoff => dayoff.staffId === staffId && dayoff.date === date);
    if (isDuplicate) return;

    state.dayoffs.push({ staffId, date });
    renderDayoffList();
  }

  function renderDayoffList() {
    if (!dayoffList) return;
    dayoffList.innerHTML = '';

    state.dayoffs.forEach(dayoff => {
      const li = document.createElement('li');
      const staff = findStaffById(dayoff.staffId);
      const staffName = staff ? staff.name : '不明なスタッフ';
      li.textContent = `${staffName} - ${dayoff.date}`;
      dayoffList.appendChild(li);
    });
  }

  function populateYears() {
    if (!yearSelect) return;
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

  function populateMonths() {
    if (!monthSelect) return;
    const currentMonth = new Date().getMonth() + 1;
    for (let month = 1; month <= 12; month++) {
      const option = document.createElement('option');
      option.value = month;
      option.textContent = `${month}月`;
      if (month === currentMonth) option.selected = true;
      monthSelect.appendChild(option);
    }
  }

  function renderHeader() {
    const resultTable = document.querySelector('#result-area table');
    if (!resultTable) return;
    resultTable.innerHTML = '';

    if (!yearSelect || !monthSelect) return;
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

  function generateShift() {
    const tableBody = document.getElementById('result-body');
    if (!tableBody || !yearSelect || !monthSelect) return;
    tableBody.innerHTML = '';

    const year = parseInt(yearSelect.value, 10);
    const month = parseInt(monthSelect.value, 10);
    if (!year || !month) return;

    const daysInMonth = new Date(year, month, 0).getDate();

    const scheduleCells = new Map();

    state.staff.forEach(staffObject => {
      normalizeFixedHolidays(staffObject);
      const row = document.createElement('tr');
      const nameCell = document.createElement('td');
      nameCell.textContent = staffObject.name;
      row.appendChild(nameCell);

      const cellMap = new Map();

      for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('td');
        const dateStr = formatDate(year, month, day);
        cell.dataset.date = dateStr;
        resetCell(cell);
        row.appendChild(cell);
        cellMap.set(dateStr, cell);
      }

      scheduleCells.set(staffObject.id, cellMap);
      tableBody.appendChild(row);
    });

    state.staff.forEach(staffObject => {
      normalizeFixedHolidays(staffObject);
      const cellMap = scheduleCells.get(staffObject.id);
      if (!cellMap) return;

      for (let day = 1; day <= daysInMonth; day++) {
        const weekdayIndex = new Date(year, month - 1, day).getDay();
        if (!staffObject.fixedHolidays.includes(weekdayIndex)) continue;

        const dateStr = formatDate(year, month, day);
        const cell = cellMap.get(dateStr);
        markCellAsDayOff(cell);
      }
    });

    state.dayoffs.forEach(dayoff => {
      const staffObject = findStaffById(dayoff.staffId);
      if (!staffObject) return;

      const dateParts = dayoff.date.split('-').map(Number);
      if (dateParts.length < 3 || dateParts.some(Number.isNaN)) return;
      const [offYear, offMonth] = dateParts;
      if (offYear !== year || offMonth !== month) return;

      const cellMap = scheduleCells.get(staffObject.id);
      if (!cellMap) return;

      const cell = cellMap.get(dayoff.date);
      markCellAsDayOff(cell);
    });
  }

  function populateShiftCheckboxes(staff) {
    if (!modalShifts) return;
    modalShifts.innerHTML = '';

    SHIFT_PATTERNS.forEach(pattern => {
      const label = document.createElement('label');
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.name = 'modal-shift';
      checkbox.value = pattern;
      checkbox.checked = staff.availableShifts.includes(pattern);
      label.appendChild(checkbox);
      label.appendChild(document.createTextNode(` ${pattern}`));
      modalShifts.appendChild(label);
    });
  }

  function populateWeekdayCheckboxes(staff) {
    if (!modalWeekdays) return;
    const inputs = modalWeekdays.querySelectorAll('input[type="checkbox"][name="modal-weekday"]');
    inputs.forEach(input => {
      const weekdayIndex = getWeekdayIndexFromKey(input.value);
      input.checked = weekdayIndex != null && staff.fixedHolidays.includes(weekdayIndex);
    });
  }

  function openStaffModal(staffId) {
    if (!staffModal) return;
    const staff = findStaffById(staffId);
    if (!staff) return;
    normalizeFixedHolidays(staff);

    state.editingStaffId = staffId;
    if (modalTitle) {
      modalTitle.textContent = `${staff.name}の詳細設定`;
    }

    populateShiftCheckboxes(staff);
    populateWeekdayCheckboxes(staff);
    if (modalMaxDays) {
      modalMaxDays.value = staff.maxWorkingDays != null ? staff.maxWorkingDays : '';
    }

    staffModal.style.display = 'flex';
  }

  function closeStaffModal() {
    if (!staffModal) return;
    staffModal.style.display = 'none';
    state.editingStaffId = null;
    if (modalForm) {
      modalForm.reset();
    }
    if (modalShifts) {
      modalShifts.innerHTML = '';
    }
  }

  function handleStaffListClick(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    const editButton = target.closest('.edit-btn');
    if (!editButton) return;
    const staffId = editButton.getAttribute('data-staff-id');
    if (staffId) {
      openStaffModal(staffId);
    }
  }

  function handleModalSave(event) {
    event.preventDefault();
    if (!state.editingStaffId) return;
    const staff = findStaffById(state.editingStaffId);
    if (!staff) return;
    normalizeFixedHolidays(staff);

    if (modalShifts) {
      const selectedShifts = Array.from(modalShifts.querySelectorAll('input[type="checkbox"][name="modal-shift"]'))
        .filter(input => input.checked)
        .map(input => input.value);
      staff.availableShifts = selectedShifts.length ? selectedShifts : [];
    }

    if (modalWeekdays) {
      const selectedWeekdays = Array.from(
        modalWeekdays.querySelectorAll('input[type="checkbox"][name="modal-weekday"]')
      )
        .filter(input => input.checked)
        .map(input => getWeekdayIndexFromKey(input.value))
        .filter(index => index != null);
      staff.fixedHolidays = Array.from(new Set(selectedWeekdays));
      normalizeFixedHolidays(staff);
    }

    if (modalMaxDays) {
      const maxDaysRaw = modalMaxDays.value.trim();
      staff.maxWorkingDays = maxDaysRaw === '' ? null : Number(maxDaysRaw);
    }

    renderStaffList();
    renderDayoffList();
    closeStaffModal();
  }

  if (addStaffButton) addStaffButton.addEventListener('click', addStaff);
  if (staffList) staffList.addEventListener('click', handleStaffListClick);
  if (addDayoffButton) addDayoffButton.addEventListener('click', addDayoff);
  if (yearSelect) yearSelect.addEventListener('change', renderHeader);
  if (monthSelect) monthSelect.addEventListener('change', renderHeader);
  if (generateButton) generateButton.addEventListener('click', generateShift);
  if (modalSaveBtn) modalSaveBtn.addEventListener('click', handleModalSave);
  if (modalCancelBtn) modalCancelBtn.addEventListener('click', closeStaffModal);
  if (staffModal) {
    staffModal.addEventListener('click', event => {
      if (event.target === staffModal) {
        closeStaffModal();
      }
    });
  }

  populateYears();
  populateMonths();
  renderHeader();
});
