document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  const SHIFT_DEFINITIONS = [
    { name: '早番', start: 7, end: 16 },
    { name: '日勤A', start: 9, end: 18 },
    { name: '日勤B', start: 9, end: 16 },
    { name: '遅番', start: 15, end: 21 },
    { name: '夜勤A', start: 16, end: 33 },
    { name: '夜勤B', start: 18, end: 33 },
    { name: '夜勤C', start: 21, end: 31 },
  ];

  const HOURLY_NEEDS = {
    bathDay: {
      '7-9': 3,
      '9-15': 5,
      '16-18': 3,
      '18-24': 2,
      '0-7': 2,
    },
    normalDay: {
      '7-9': 3,
      '9-15': 4,
      '16-18': 3,
      '18-24': 2,
      '0-7': 2,
    },
    wednesday: {
      '7-9': 3,
      '9-15': 2,
      '16-18': 4,
      '18-24': 2,
      '0-7': 2,
    },
  };

  const MAX_CONSECUTIVE_WORKDAYS = 5;
  const NIGHT_SHIFTS = ['夜勤A', '夜勤B', '夜勤C'];

  const today = new Date();
  const nextMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
  const DEFAULT_YEAR = nextMonthDate.getFullYear();
  const DEFAULT_MONTH = nextMonthDate.getMonth() + 1;

  const SHIFT_PATTERNS = SHIFT_DEFINITIONS.map(pattern => pattern.name);
  const WEEKDAY_INDEX_MAP = {
    sun: '0',
    mon: '1',
    tue: '2',
    wed: '3',
    thu: '4',
    fri: '5',
    sat: '6',
  };

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

  function saveState() {
    try {
      localStorage.setItem('shiftToolState', JSON.stringify(state));
    } catch (error) {
      console.error('Failed to save shift tool state:', error);
    }
  }

  function loadState() {
    try {
      const stored = localStorage.getItem('shiftToolState');
      if (!stored) {
        renderStaffList();
        renderDayoffList();
        return;
      }

      const parsed = JSON.parse(stored);
      if (parsed && typeof parsed === 'object') {
        if (Array.isArray(parsed.staff)) {
          state.staff = parsed.staff.map(item => {
            const available = Array.isArray(item.availableShifts)
              ? item.availableShifts.filter(shift => SHIFT_PATTERNS.includes(shift))
              : [...SHIFT_PATTERNS];
            const fixedHolidays = Array.isArray(item.fixedHolidays)
              ? item.fixedHolidays.map(value => String(value))
              : [];
            let maxWorkingDays = null;
            if (item.maxWorkingDays === null) {
              maxWorkingDays = null;
            } else if (typeof item.maxWorkingDays === 'number' && !Number.isNaN(item.maxWorkingDays)) {
              maxWorkingDays = item.maxWorkingDays;
            }

            return {
              id: item.id || generateStaffId(),
              name: typeof item.name === 'string' ? item.name : '',
              availableShifts: available,
              fixedHolidays,
              maxWorkingDays,
            };
          });
        }

        if (Array.isArray(parsed.dayoffs)) {
          state.dayoffs = parsed.dayoffs
            .map(dayoff => ({
              staffId: dayoff.staffId || null,
              staffName: dayoff.staffName || null,
              date: dayoff.date,
            }))
            .filter(dayoff => typeof dayoff.date === 'string');
        }
      }
    } catch (error) {
      console.error('Failed to load shift tool state:', error);
    }

    renderStaffList();
    renderDayoffList();
  }

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

    saveState();
  }

  function renderStaffList() {
    if (staffList) {
      staffList.innerHTML = '';
    }
    if (staffSelect) {
      staffSelect.innerHTML = '<option value="">スタッフを選択</option>';
    }

    state.staff.forEach(staff => {
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

    const staff = state.staff.find(item => item.id === staffId);
    const staffName = staff ? staff.name : '';

    const isDuplicate = state.dayoffs.some(dayoff => {
      const matchesStaffId = dayoff.staffId === staffId;
      const matchesStaffName = dayoff.staffName && staffName && dayoff.staffName === staffName;
      return (matchesStaffId || matchesStaffName) && dayoff.date === date;
    });
    if (isDuplicate) return;

    state.dayoffs.push({ staffId, staffName, date });
    renderDayoffList();

    saveState();
  }

  function renderDayoffList() {
    if (!dayoffList) return;
    dayoffList.innerHTML = '';

    state.dayoffs.forEach(dayoff => {
      const li = document.createElement('li');
      const staff = state.staff.find(item => item.id === dayoff.staffId);
      const staffName = dayoff.staffName || (staff ? staff.name : '不明なスタッフ');
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
      option.value = String(year);
      option.textContent = `${year}年`;
      if (year === DEFAULT_YEAR) option.selected = true;
      yearSelect.appendChild(option);
    }

    if (!yearSelect.value) {
      yearSelect.value = String(DEFAULT_YEAR);
    }
  }

  function populateMonths() {
    if (!monthSelect) return;
    for (let month = 1; month <= 12; month++) {
      const option = document.createElement('option');
      option.value = String(month);
      option.textContent = `${month}月`;
      if (month === DEFAULT_MONTH) option.selected = true;
      monthSelect.appendChild(option);
    }

    if (!monthSelect.value) {
      monthSelect.value = String(DEFAULT_MONTH);
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

  function normalizeFixedHolidays(staffObject) {
    if (!staffObject) return [];
    if (!Array.isArray(staffObject.fixedHolidays)) {
      staffObject.fixedHolidays = [];
      return staffObject.fixedHolidays;
    }

    const normalized = staffObject.fixedHolidays
      .map(value => {
        if (value == null) return null;
        const strValue = String(value).trim();
        if (strValue === '') return null;
        if (/^[0-6]$/.test(strValue)) return strValue;
        const mapped = WEEKDAY_INDEX_MAP[strValue];
        return mapped != null ? mapped : null;
      })
      .filter(value => value != null);

    const uniqueValues = Array.from(new Set(normalized));
    staffObject.fixedHolidays = uniqueValues;
    return uniqueValues;
  }

  function getDayType(dayOfWeek) {
    if (dayOfWeek === 3) return 'wednesday';
    if (dayOfWeek === 0 || dayOfWeek === 6) return 'normalDay';
    return 'bathDay';
  }

  function buildHourlyNeeds(dayType) {
    const hourlyTemplate = HOURLY_NEEDS[dayType];
    const hourlyNeeds = new Array(24).fill(0);
    if (!hourlyTemplate) return hourlyNeeds;

    Object.entries(hourlyTemplate).forEach(([range, count]) => {
      const [startStr, endStr] = range.split('-');
      const start = parseInt(startStr, 10);
      const end = parseInt(endStr, 10);
      if (Number.isNaN(start) || Number.isNaN(end)) return;
      for (let hour = start; hour < end; hour++) {
        hourlyNeeds[hour % 24] = count;
      }
    });

    return hourlyNeeds;
  }

  function computeShiftRequirements(dayType) {
    const hourlyNeeds = buildHourlyNeeds(dayType);
    const remainingNeeds = hourlyNeeds.slice();
    const requirements = {};

    for (let iteration = 0; iteration < 500; iteration++) {
      let bestShift = null;
      let bestScore = 0;

      SHIFT_DEFINITIONS.forEach(shift => {
        let score = 0;
        for (let hour = shift.start; hour < shift.end; hour++) {
          score += remainingNeeds[hour % 24];
        }
        if (score > bestScore) {
          bestScore = score;
          bestShift = shift;
        }
      });

      if (!bestShift || bestScore === 0) {
        break;
      }

      requirements[bestShift.name] = (requirements[bestShift.name] || 0) + 1;
      for (let hour = bestShift.start; hour < bestShift.end; hour++) {
        const index = hour % 24;
        if (remainingNeeds[index] > 0) {
          remainingNeeds[index] -= 1;
        }
      }

      if (!remainingNeeds.some(value => value > 0)) {
        break;
      }
    }

    return requirements;
  }

  function markCellAsOff(cellRecord, backgroundColor = '#ffdcdc') {
    if (!cellRecord) return;
    cellRecord.assignment = '休み';
    cellRecord.cell.textContent = '休み';
    cellRecord.cell.style.backgroundColor = backgroundColor;
    cellRecord.isLockedOff = true;
  }

  function markNightShiftRest(cellRecord) {
    markCellAsOff(cellRecord, '#fff2cc');
  }

  function markForcedRest(cellRecord) {
    markCellAsOff(cellRecord, '#fff2cc');
  }

  function assignShiftToCell(cellRecord, shiftName) {
    if (!cellRecord) return;
    cellRecord.assignment = shiftName;
    cellRecord.cell.textContent = shiftName;
    cellRecord.cell.style.backgroundColor = '#e6f7ff';
  }

  function canAssignShift(record, dayIndex, shiftName) {
    if (!record) return false;
    const cellRecord = record.cells[dayIndex];
    if (!cellRecord || cellRecord.isLockedOff) return false;
    if (cellRecord.assignment) return false;

    const available = Array.isArray(record.staffObject.availableShifts)
      ? record.staffObject.availableShifts
      : [];
    if (!available.includes(shiftName)) return false;

    const maxDays = record.staffObject.maxWorkingDays;
    if (typeof maxDays === 'number' && !Number.isNaN(maxDays) && record.workingDays >= maxDays) {
      return false;
    }

    return true;
  }

  function generateShift() {
    const tableBody = document.getElementById('result-body');
    if (!tableBody || !yearSelect || !monthSelect) return;
    tableBody.innerHTML = '';

    const year = parseInt(yearSelect.value, 10);
    const month = parseInt(monthSelect.value, 10);
    if (!year || !month) return;

    const daysInMonth = new Date(year, month, 0).getDate();
    const staffRecords = state.staff.map(staff => {
      const row = document.createElement('tr');
      const nameCell = document.createElement('td');
      nameCell.textContent = staff.name;
      row.appendChild(nameCell);

      const cellRecords = [];
      for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('td');
        cell.textContent = '';
        cell.style.backgroundColor = '';

        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = new Date(year, month - 1, day).getDay();

        cellRecords.push({
          cell,
          dateStr,
          dayOfWeek,
          assignment: '',
          isLockedOff: false,
        });
        row.appendChild(cell);
      }

      tableBody.appendChild(row);

      return {
        staffObject: staff,
        fixedHolidays: normalizeFixedHolidays(staff),
        workingDays: 0,
        nightShiftRestDays: new Set(),
        cells: cellRecords,
      };
    });

    // Apply fixed holidays first.
    staffRecords.forEach(record => {
      if (!record.fixedHolidays.length) return;
      record.cells.forEach(cellRecord => {
        if (record.fixedHolidays.includes(String(cellRecord.dayOfWeek))) {
          markCellAsOff(cellRecord);
        }
      });
    });

    // Apply requested day-offs afterwards to overwrite as needed.
    staffRecords.forEach(record => {
      record.cells.forEach(cellRecord => {
        if (isDayOff(record.staffObject, cellRecord.dateStr)) {
          markCellAsOff(cellRecord);
        }
      });
    });

    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      staffRecords.forEach(record => {
        const cellRecord = record.cells[dayIndex];
        if (!cellRecord || cellRecord.isLockedOff) {
          if (record.nightShiftRestDays.has(dayIndex)) {
            record.nightShiftRestDays.delete(dayIndex);
          }
          return;
        }

        if (record.nightShiftRestDays.has(dayIndex)) {
          markNightShiftRest(cellRecord);
          record.nightShiftRestDays.delete(dayIndex);
          return;
        }

        const prevDayIndex = dayIndex - 1;
        if (prevDayIndex >= 0) {
          const prevAssignment = record.cells[prevDayIndex]?.assignment;
          if (prevAssignment && NIGHT_SHIFTS.includes(prevAssignment)) {
            markNightShiftRest(cellRecord);
            return;
          }
        }

        let consecutiveWorkdays = 0;
        for (let back = dayIndex - 1; back >= 0; back--) {
          const previous = record.cells[back];
          if (!previous || !previous.assignment || previous.assignment === '休み') {
            break;
          }
          consecutiveWorkdays += 1;
        }

        if (consecutiveWorkdays >= MAX_CONSECUTIVE_WORKDAYS) {
          markForcedRest(cellRecord);
          return;
        }
      });

      const dayType = getDayType(dayOfWeek);
      const requirements = computeShiftRequirements(dayType);

      SHIFT_DEFINITIONS.forEach(shift => {
        let remaining = requirements[shift.name] || 0;
        while (remaining > 0) {
          const candidate = staffRecords.find(record => canAssignShift(record, dayIndex, shift.name));
          if (!candidate) {
            break;
          }

          const cellRecord = candidate.cells[dayIndex];
          assignShiftToCell(cellRecord, shift.name);
          candidate.workingDays += 1;

          if (NIGHT_SHIFTS.includes(shift.name)) {
            const nextDayIndex = dayIndex + 1;
            if (nextDayIndex < candidate.cells.length) {
              candidate.nightShiftRestDays.add(nextDayIndex);
            }
          }

          remaining -= 1;
        }
      });
    }
  }

  function isDayOff(staffObject, dateStr) {
    return state.dayoffs.some(dayoff => {
      const matchesStaffByName = dayoff.staffName && dayoff.staffName === staffObject.name;
      const matchesStaffById = !dayoff.staffName && dayoff.staffId && dayoff.staffId === staffObject.id;
      if (!matchesStaffByName && !matchesStaffById) return false;
      return dayoff.date === dateStr;
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
    const normalizedFixedHolidays = normalizeFixedHolidays(staff);
    const inputs = modalWeekdays.querySelectorAll('input[type="checkbox"][name="modal-weekday"]');
    inputs.forEach(input => {
      const weekdayIndex = WEEKDAY_INDEX_MAP[input.value];
      input.checked = weekdayIndex != null && normalizedFixedHolidays.includes(weekdayIndex);
    });
  }

  function openStaffModal(staffId) {
    if (!staffModal) return;
    const staff = state.staff.find(item => item.id === staffId);
    if (!staff) return;

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
    const staff = state.staff.find(item => item.id === state.editingStaffId);
    if (!staff) return;

    if (modalShifts) {
      const selectedShifts = Array.from(modalShifts.querySelectorAll('input[type="checkbox"][name="modal-shift"]'))
        .filter(input => input.checked)
        .map(input => input.value);
      staff.availableShifts = selectedShifts.length ? selectedShifts : [];
    }

    if (modalWeekdays) {
      const selectedWeekdays = Array.from(modalWeekdays.querySelectorAll('input[type="checkbox"][name="modal-weekday"]'))
        .filter(input => input.checked)
        .map(input => WEEKDAY_INDEX_MAP[input.value])
        .filter(value => value != null);
      staff.fixedHolidays = selectedWeekdays;
      normalizeFixedHolidays(staff);
    }

    if (modalMaxDays) {
      const maxDaysRaw = modalMaxDays.value.trim();
      staff.maxWorkingDays = maxDaysRaw === '' ? null : Number(maxDaysRaw);
    }

    renderStaffList();
    renderDayoffList();
    closeStaffModal();

    saveState();
  }

  loadState();

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
