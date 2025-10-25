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
  const MIN_WORKDAY_GOAL_RATIO = 0;
  const MIN_GOAL_BONUS_WEIGHT = 10;
  const FAIR_WEIGHTS = {
    progressPenalty: 4,
    monthOver: 2,
    weekRisk: 6,
    streakNearMax: 2,
  };

  const DAYTIME_SHIFTS = ['早番', '日勤A', '日勤B', '遅番'];
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

  const generateButton = document.getElementById('generate-btn');
  const exportCsvButton = document.getElementById('export-csv-btn');
  const csvOutput = document.getElementById('csv-output');

  const staffModal = document.getElementById('staff-modal');
  const modalTitle = document.getElementById('modal-title');
  const modalForm = document.getElementById('modal-form');
  const modalShifts = document.getElementById('modal-shifts');
  const modalWeekdays = document.getElementById('modal-weekdays');
  const modalMaxDays = document.getElementById('modal-max-days');
  const modalMaxDaysPerWeek = document.getElementById('modal-max-days-per-week');
  const modalSaveBtn = document.getElementById('modal-save-btn');
  const modalCancelBtn = document.getElementById('modal-cancel-btn');

  const state = {
    staff: [],
    dayoffs: [],
    editingStaffId: null,
    targetYear: null,
    targetMonth: null,
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
            if (item.maxWorkingDays != null && item.maxWorkingDays !== '') {
              const parsedMaxWorking = Number(item.maxWorkingDays);
              if (Number.isFinite(parsedMaxWorking)) {
                maxWorkingDays = parsedMaxWorking;
              }
            }

            let maxDaysPerWeek = null;
            if (item.maxDaysPerWeek != null && item.maxDaysPerWeek !== '') {
              const parsedWeekly = Number(item.maxDaysPerWeek);
              if (Number.isFinite(parsedWeekly)) {
                maxDaysPerWeek = parsedWeekly;
              }
            }

            return {
              id: item.id || generateStaffId(),
              name: typeof item.name === 'string' ? item.name : '',
              availableShifts: available,
              fixedHolidays,
              maxWorkingDays,
              maxDaysPerWeek,
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

        if (parsed.targetYear != null) {
          const parsedYear = Number(parsed.targetYear);
          state.targetYear = Number.isFinite(parsedYear) ? Math.trunc(parsedYear) : null;
        }

        if (parsed.targetMonth != null) {
          const parsedMonth = Number(parsed.targetMonth);
          const monthInt = Number.isFinite(parsedMonth) ? Math.trunc(parsedMonth) : null;
          state.targetMonth = monthInt && monthInt >= 1 && monthInt <= 12 ? monthInt : null;
        }
      }
    } catch (error) {
      console.error('Failed to load shift tool state:', error);
    }

    renderStaffList();
    renderDayoffList();
  }

  function setNextMonthTarget() {
    const today = new Date();
    const nextMonthDate = new Date(today.getFullYear(), today.getMonth() + 1, 1);
    const nextYear = nextMonthDate.getFullYear();
    const nextMonth = nextMonthDate.getMonth() + 1;

    const changed = state.targetYear !== nextYear || state.targetMonth !== nextMonth;
    state.targetYear = nextYear;
    state.targetMonth = nextMonth;
    return changed;
  }

  function updateDayoffDateBounds() {
    if (!dayoffDateInput) return;

    const year = state.targetYear;
    const month = state.targetMonth;
    if (year == null || month == null) {
      dayoffDateInput.removeAttribute('min');
      dayoffDateInput.removeAttribute('max');
      return;
    }

    const paddedMonth = String(month).padStart(2, '0');
    const start = `${year}-${paddedMonth}-01`;
    const daysInMonth = new Date(year, month, 0).getDate();
    const end = `${year}-${paddedMonth}-${String(daysInMonth).padStart(2, '0')}`;

    dayoffDateInput.min = start;
    dayoffDateInput.max = end;

    if (dayoffDateInput.value) {
      if (dayoffDateInput.value < start || dayoffDateInput.value > end) {
        dayoffDateInput.value = '';
      }
    }
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
      maxDaysPerWeek: null,
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

        const deleteButton = document.createElement('button');
        deleteButton.type = 'button';
        deleteButton.className = 'delete-btn';
        deleteButton.textContent = '削除';
        deleteButton.setAttribute('data-staff-id', staff.id);
        li.appendChild(deleteButton);

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

  function renderHeader() {
    const year = state.targetYear;
    const month = state.targetMonth;

    updateDayoffDateBounds();

    const resultTable = document.querySelector('#result-area table');
    if (!resultTable) return;
    if (year == null || month == null) return;

    resultTable.innerHTML = '';

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

  function exportToCSV() {
    const resultTable = document.querySelector('#result-area table');
    const csvOutput = document.getElementById('csv-output');

    if (!resultTable || !csvOutput) {
      console.error('CSV出力に必要な要素が見つかりません。');
      return;
    }

    let csvString = '';
    const rows = [];

    const headerRow = resultTable.querySelector('thead tr:first-child');
    if (headerRow) {
      const headers = Array.from(headerRow.querySelectorAll('th')).map(th => `"${th.textContent.trim()}"`);
      rows.push(headers.join(','));
    }

    const dataRows = resultTable.querySelectorAll('tbody tr');
    dataRows.forEach(row => {
      const cols = Array.from(row.querySelectorAll('td')).map(td => `"${td.textContent.trim()}"`);
      rows.push(cols.join(','));
    });

    csvString = rows.join('\n');

    csvOutput.value = csvString;
    csvOutput.style.display = 'block';
    csvOutput.select();
    document.execCommand('copy');
    alert('CSVデータをクリップボードにコピーしました！');
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

  function createHourlyNeedsMap(dayType) {
    const map = new Array(24).fill(0);
    const hourlyTemplate = HOURLY_NEEDS ? HOURLY_NEEDS[dayType] : null;
    if (!hourlyTemplate) {
      return map;
    }

    Object.entries(hourlyTemplate).forEach(([range, rawCount]) => {
      const [startStr, endStr] = range.split('-');
      const start = Number(startStr);
      const end = Number(endStr);
      const count = Number(rawCount);

      if (!Number.isFinite(start) || !Number.isFinite(end) || !Number.isFinite(count)) {
        return;
      }

      const normalizedEnd = end > start ? end : end + 24;
      for (let hour = start; hour < normalizedEnd; hour++) {
        map[hour % 24] = count;
      }
    });

    return map;
  }

  function createHourlySupplyMap(assignedShifts) {
    const map = new Array(24).fill(0);
    if (!Array.isArray(assignedShifts) || !assignedShifts.length) {
      return map;
    }

    assignedShifts.forEach(entry => {
      if (!entry) return;

      let shiftDefinition = entry.shift;
      if (typeof shiftDefinition === 'string') {
        shiftDefinition = SHIFT_DEFINITIONS.find(def => def.name === shiftDefinition) || null;
      }

      if (!shiftDefinition || typeof shiftDefinition.start !== 'number' || typeof shiftDefinition.end !== 'number') {
        return;
      }

      const start = shiftDefinition.start;
      const end = shiftDefinition.end;
      const normalizedEnd = end > start ? end : end + 24;

      for (let hour = start; hour < normalizedEnd; hour++) {
        const index = hour % 24;
        map[index] = (map[index] || 0) + 1;
      }
    });

    return map;
  }

  function calculateDeficitMap(needsMap, supplyMap) {
    const deficit = new Array(24).fill(0);
    for (let hour = 0; hour < 24; hour++) {
      const need = Array.isArray(needsMap) ? needsMap[hour] : needsMap ? needsMap[hour] : undefined;
      const supply = Array.isArray(supplyMap) ? supplyMap[hour] : supplyMap ? supplyMap[hour] : undefined;
      const needValue = Number.isFinite(need) ? need : 0;
      const supplyValue = Number.isFinite(supply) ? supply : 0;
      deficit[hour] = needValue - supplyValue;
    }
    return deficit;
  }

  function markCellAsOff(cellRecord, backgroundColor = '#ffdcdc') {
    if (!cellRecord) return;
    cellRecord.assignment = '休み';
    cellRecord.backgroundColor = backgroundColor;
    cellRecord.isLocked = true;
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
    cellRecord.backgroundColor = '#e6f7ff';
    cellRecord.isLocked = true;
  }

  function isFiniteNumber(value) {
    return typeof value === 'number' && Number.isFinite(value);
  }

  function isWorkingAssignment(value) {
    return value && value !== '休み';
  }

  function countConsecutiveWorkdays(record, dayIndex) {
    let consecutive = 0;
    for (let back = dayIndex - 1; back >= 0; back--) {
      const previous = record.cells[back];
      if (!previous || !isWorkingAssignment(previous.assignment)) {
        break;
      }
      consecutive += 1;
    }
    return consecutive;
  }

  function canAssignShift(record, dayIndex, shiftName) {
    if (!record) return false;
    const cellRecord = record.cells[dayIndex];
    if (!cellRecord || cellRecord.isLocked) return false;
    if (cellRecord.assignment) return false;

    const available = Array.isArray(record.staffObject.availableShifts)
      ? record.staffObject.availableShifts
      : [];
    if (!available.includes(shiftName)) return false;

    const maxDays = record.staffObject.maxWorkingDays;
    if (isFiniteNumber(maxDays) && record.workdaysInMonth >= maxDays) {
      return false;
    }

    const maxDaysPerWeek = record.staffObject.maxDaysPerWeek;
    if (isFiniteNumber(maxDaysPerWeek)) {
      const weeklyWorked = record.workdaysInWeek || 0;
      if (weeklyWorked >= maxDaysPerWeek) {
        return false;
      }
    }

    return true;
  }

  function collectEligibleRecords(staffRecords, dayIndex, allowedShiftNames = null) {
    return staffRecords.filter(record => {
      const cell = record.cells[dayIndex];
      if (!cell || cell.isLocked || cell.assignment) {
        return false;
      }

      const available = Array.isArray(record.staffObject.availableShifts)
        ? record.staffObject.availableShifts
        : [];
      const candidateShifts = allowedShiftNames
        ? available.filter(name => allowedShiftNames.includes(name))
        : available.slice();
      if (!candidateShifts.length) {
        return false;
      }

      const maxDays = record.staffObject.maxWorkingDays;
      if (isFiniteNumber(maxDays) && record.workdaysInMonth >= maxDays) {
        return false;
      }

      const maxDaysPerWeek = record.staffObject.maxDaysPerWeek;
      if (isFiniteNumber(maxDaysPerWeek) && record.workdaysInWeek >= maxDaysPerWeek) {
        return false;
      }

      return true;
    });
  }

  function collectAssignmentsForDay(staffRecords, dayIndex) {
    const assignments = [];
    staffRecords.forEach(record => {
      const cell = record.cells[dayIndex];
      if (!cell || !isWorkingAssignment(cell.assignment)) {
        return;
      }

      const shiftDefinition = SHIFT_DEFINITIONS.find(def => def.name === cell.assignment);
      if (!shiftDefinition) {
        return;
      }

      assignments.push({
        staff: record.staffObject,
        shift: shiftDefinition,
        record,
      });
    });
    return assignments;
  }

  function resetWorkCounters(staffRecords) {
    staffRecords.forEach(record => {
      record.workdaysInMonth = 0;
      record.workdaysInWeek = 0;
    });
  }

  function hasPositiveDeficitForShifts(deficitMap, shiftNames) {
    return shiftNames.some(shiftName => {
      const definition = SHIFT_DEFINITIONS.find(def => def.name === shiftName);
      if (!definition) return false;
      for (let hour = definition.start; hour < definition.end; hour++) {
        if ((deficitMap[hour % 24] || 0) > 0) {
          return true;
        }
      }
      return false;
    });
  }

  function enforceConsecutiveRest(staffRecords, daysInMonth) {
    let changed = true;
    while (changed) {
      changed = false;
      staffRecords.forEach(record => {
        let consecutive = 0;
        for (let dayIndex = 0; dayIndex < daysInMonth; dayIndex++) {
          const cell = record.cells[dayIndex];
          if (!cell) continue;

          if (cell.assignment === '休み') {
            consecutive = 0;
            continue;
          }

          consecutive += 1;
          if (consecutive > MAX_CONSECUTIVE_WORKDAYS) {
            if (!cell.isLocked || cell.assignment !== '休み') {
              markForcedRest(cell);
              changed = true;
            }
            consecutive = 0;
          }
        }
      });
    }
  }

  function findBestAssignment(
    availableRecords,
    deficitMap,
    dayIndex,
    allowedShiftNames = null,
    daysInMonth = 30
  ) {
    if (!Array.isArray(availableRecords) || !availableRecords.length) {
      return null;
    }

    let bestMove = null;

    availableRecords.forEach(record => {
      if (!record || !record.staffObject || !Array.isArray(record.cells)) return;

      const available = Array.isArray(record.staffObject.availableShifts)
        ? record.staffObject.availableShifts
        : [];
      if (!available.length) return;

      const candidateShifts = allowedShiftNames
        ? available.filter(name => allowedShiftNames.includes(name))
        : available.slice();
      if (!candidateShifts.length) return;

      const prevAssignment =
        typeof dayIndex === 'number' && dayIndex > 0 ? record.cells[dayIndex - 1]?.assignment || '' : '';
      const prevWasNight = prevAssignment && NIGHT_SHIFTS.includes(prevAssignment);
      const prevWasLateDay = prevAssignment === '日勤A' || prevAssignment === '日勤B';

      const consecutiveWorkdays = typeof dayIndex === 'number' ? countConsecutiveWorkdays(record, dayIndex) : 0;

      candidateShifts.forEach(shiftName => {
        const shiftDefinition = SHIFT_DEFINITIONS.find(def => def.name === shiftName);
        if (!shiftDefinition) return;

        if (prevWasLateDay && shiftDefinition.name === '早番') {
          return;
        }

        if (prevWasNight) {
          return;
        }

        const wouldBeConsecutive = consecutiveWorkdays + 1;
        if (wouldBeConsecutive > MAX_CONSECUTIVE_WORKDAYS) {
          return;
        }

        const maxDays = record.staffObject.maxWorkingDays;
        if (isFiniteNumber(maxDays) && record.workdaysInMonth + 1 > maxDays) {
          return;
        }

        const maxDaysPerWeek = record.staffObject.maxDaysPerWeek;
        if (isFiniteNumber(maxDaysPerWeek) && record.workdaysInWeek + 1 > maxDaysPerWeek) {
          return;
        }

        let score = 0;
        for (let hour = shiftDefinition.start; hour < shiftDefinition.end; hour++) {
          const deficit = deficitMap[hour % 24] || 0;
          if (deficit > 0) {
            score += deficit;
          }
        }

        const afterMonth = (record.workdaysInMonth || 0) + 1;
        const target = isFiniteNumber(record.targetWorkdays)
          ? record.targetWorkdays
          : Math.ceil(daysInMonth * 0.55);
        const expected = Math.round(target * ((dayIndex + 1) / daysInMonth));
        const progressOver = Math.max(0, afterMonth - expected);
        score -= FAIR_WEIGHTS.progressPenalty * progressOver;

        if (afterMonth > target) {
          score -= FAIR_WEIGHTS.monthOver * (afterMonth - target);
        }

        const weekAfter = (record.workdaysInWeek || 0) + 1;
        if (isFiniteNumber(maxDaysPerWeek)) {
          const threshold = Math.max(1, maxDaysPerWeek - 1);
          if (weekAfter > threshold) {
            score -= FAIR_WEIGHTS.weekRisk * (weekAfter - threshold);
          }
        }

        if (wouldBeConsecutive >= MAX_CONSECUTIVE_WORKDAYS) {
          score -= FAIR_WEIGHTS.streakNearMax * (wouldBeConsecutive - MAX_CONSECUTIVE_WORKDAYS + 1);
        }

        if (score <= 0) {
          return;
        }

        const availableCount = available.length;
        if (NIGHT_SHIFTS.includes(shiftDefinition.name)) {
          if (availableCount > 0 && availableCount <= 3) {
            score += 50;
          }
        } else if (DAYTIME_SHIFTS.includes(shiftDefinition.name)) {
          const canWorkNight = available.some(name => NIGHT_SHIFTS.includes(name));
          if (!canWorkNight) {
            score += 30;
          }
        }

        if (!bestMove || score > bestMove.score) {
          bestMove = {
            staffRecord: record,
            shift: shiftDefinition,
            score,
          };
          return;
        }

        if (bestMove && score === bestMove.score) {
          const currentMonthWork = record.workdaysInMonth || 0;
          const bestMonthWork = bestMove.staffRecord ? bestMove.staffRecord.workdaysInMonth || 0 : 0;
          if (currentMonthWork < bestMonthWork) {
            bestMove = {
              staffRecord: record,
              shift: shiftDefinition,
              score,
            };
          }
        }
      });
    });

    return bestMove;
  }

  function generateShift() {
    if (state.targetYear == null || state.targetMonth == null) {
      console.error('対象年月が設定されていないため、シフトを生成できません。');
      return;
    }

    renderHeader();

    const tableBody = document.getElementById('result-body');
    if (!tableBody) return;
    tableBody.innerHTML = '';

    const year = state.targetYear;
    const month = state.targetMonth;
    const daysInMonth = new Date(year, month, 0).getDate();

    const staffRecords = state.staff.map(staff => {
      const row = document.createElement('tr');
      const nameCell = document.createElement('td');
      nameCell.textContent = staff.name;
      row.appendChild(nameCell);

      const cells = [];
      for (let day = 1; day <= daysInMonth; day++) {
        const cell = document.createElement('td');
        row.appendChild(cell);

        const dateStr = `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayOfWeek = new Date(year, month - 1, day).getDay();

        cells.push({
          cell,
          dateStr,
          dayOfWeek,
          assignment: '',
          isLocked: false,
          backgroundColor: '',
        });
      }

      tableBody.appendChild(row);

      return {
        staffObject: staff,
        rowElement: row,
        fixedHolidays: normalizeFixedHolidays(staff),
        cells,
        workdaysInMonth: 0,
        workdaysInWeek: 0,
        targetWorkdays: isFiniteNumber(staff.maxWorkingDays)
          ? staff.maxWorkingDays
          : Math.ceil(daysInMonth * 0.55),
        minWorkdaysGoal: null,
      };
    });

    // Phase 1: establish rest blocks from fixed and requested holidays
    staffRecords.forEach(record => {
      record.cells.forEach(cellRecord => {
        if (record.fixedHolidays.includes(String(cellRecord.dayOfWeek))) {
          markCellAsOff(cellRecord);
        }
      });
    });

    staffRecords.forEach(record => {
      record.cells.forEach(cellRecord => {
        if (cellRecord.isLocked) return;
        if (isDayOff(record.staffObject, cellRecord.dateStr)) {
          markCellAsOff(cellRecord);
        }
      });
    });

    enforceConsecutiveRest(staffRecords, daysInMonth);

    // Phase 2: allocate night shifts first
    resetWorkCounters(staffRecords);
    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      if (dayOfWeek === 0) {
        staffRecords.forEach(record => {
          record.workdaysInWeek = 0;
        });
      }

      const dayType = getDayType(dayOfWeek);
      const needsMap = createHourlyNeedsMap(dayType);
      let assignedShiftsThisDay = collectAssignmentsForDay(staffRecords, dayIndex);
      let supplyMap = createHourlySupplyMap(assignedShiftsThisDay);

      while (true) {
        const deficitMap = calculateDeficitMap(needsMap, supplyMap);
        if (!hasPositiveDeficitForShifts(deficitMap, NIGHT_SHIFTS)) {
          break;
        }

        const availableRecords = collectEligibleRecords(staffRecords, dayIndex, NIGHT_SHIFTS);
        if (!availableRecords.length) {
          break;
        }

        const bestMove = findBestAssignment(
          availableRecords,
          deficitMap,
          dayIndex,
          NIGHT_SHIFTS,
          daysInMonth
        );
        if (!bestMove) {
          break;
        }

        const { staffRecord, shift } = bestMove;
        if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
          markForcedRest(staffRecord.cells[dayIndex]);
          continue;
        }

        assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
        staffRecord.workdaysInMonth += 1;
        staffRecord.workdaysInWeek += 1;

        assignedShiftsThisDay.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
        supplyMap = createHourlySupplyMap(assignedShiftsThisDay);

        const nextIndex = dayIndex + 1;
        if (nextIndex < daysInMonth) {
          const nextCell = staffRecord.cells[nextIndex];
          if (nextCell && !nextCell.isLocked) {
            markNightShiftRest(nextCell);
          }
        }
      }
    }

    // Phase 3: cover daytime gaps with remaining capacity
    resetWorkCounters(staffRecords);
    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      if (dayOfWeek === 0) {
        staffRecords.forEach(record => {
          record.workdaysInWeek = 0;
        });
      }

      const dayType = getDayType(dayOfWeek);
      const needsMap = createHourlyNeedsMap(dayType);
      let assignedShiftsThisDay = collectAssignmentsForDay(staffRecords, dayIndex);

      assignedShiftsThisDay.forEach(entry => {
        entry.record.workdaysInMonth += 1;
        entry.record.workdaysInWeek += 1;
      });

      let supplyMap = createHourlySupplyMap(assignedShiftsThisDay);

      while (true) {
        const deficitMap = calculateDeficitMap(needsMap, supplyMap);
        const remainingNeeds = deficitMap.some(value => value > 0);
        if (!remainingNeeds) {
          break;
        }

        const availableRecords = collectEligibleRecords(staffRecords, dayIndex, DAYTIME_SHIFTS);
        if (!availableRecords.length) {
          break;
        }

        const bestMove = findBestAssignment(
          availableRecords,
          deficitMap,
          dayIndex,
          DAYTIME_SHIFTS,
          daysInMonth
        );
        if (!bestMove) {
          break;
        }

        const { staffRecord, shift } = bestMove;
        if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
          markForcedRest(staffRecord.cells[dayIndex]);
          continue;
        }

        assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
        staffRecord.workdaysInMonth += 1;
        staffRecord.workdaysInWeek += 1;

        assignedShiftsThisDay.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
        supplyMap = createHourlySupplyMap(assignedShiftsThisDay);
      }
    }

    // Final rendering: push schedule back to DOM
    staffRecords.forEach(record => {
      record.cells.forEach(cellRecord => {
        const content = cellRecord.assignment || '';
        cellRecord.cell.textContent = content;
        cellRecord.cell.style.backgroundColor = cellRecord.backgroundColor || '';
      });
    });
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

    if (modalMaxDaysPerWeek) {
      modalMaxDaysPerWeek.value = staff.maxDaysPerWeek != null ? staff.maxDaysPerWeek : '';
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
    const deleteButton = target.closest('.delete-btn');
    if (deleteButton) {
      const staffId = deleteButton.getAttribute('data-staff-id');
      if (!staffId) return;
      const staff = state.staff.find(item => item.id === staffId);
      if (!staff) return;

      const confirmed = window.confirm(
        `${staff.name}さんを本当に削除しますか？関連する希望休もすべて削除されます。`
      );
      if (!confirmed) return;

      state.staff = state.staff.filter(item => item.id !== staffId);
      state.dayoffs = state.dayoffs.filter(dayoff => {
        if (dayoff.staffId) {
          return dayoff.staffId !== staffId;
        }
        if (dayoff.staffName) {
          return dayoff.staffName !== staff.name;
        }
        return true;
      });

      if (state.editingStaffId === staffId) {
        closeStaffModal();
      }

      renderStaffList();
      renderDayoffList();
      saveState();
      return;
    }

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
      if (maxDaysRaw === '') {
        staff.maxWorkingDays = null;
      } else {
        const parsedMaxDays = Number(maxDaysRaw);
        staff.maxWorkingDays = Number.isFinite(parsedMaxDays) ? parsedMaxDays : null;
      }
    }

    if (modalMaxDaysPerWeek) {
      const maxDaysPerWeekRaw = modalMaxDaysPerWeek.value.trim();
      if (maxDaysPerWeekRaw === '') {
        staff.maxDaysPerWeek = null;
      } else {
        const parsedWeekly = Number(maxDaysPerWeekRaw);
        staff.maxDaysPerWeek = Number.isFinite(parsedWeekly) ? parsedWeekly : null;
      }
    }

    renderStaffList();
    renderDayoffList();
    closeStaffModal();

    saveState();
  }
  renderHeader();

  loadState();
  const targetChanged = setNextMonthTarget();
  if (targetChanged) {
    saveState();
  }
  renderHeader();

  if (addStaffButton) addStaffButton.addEventListener('click', addStaff);
  if (staffList) staffList.addEventListener('click', handleStaffListClick);
  if (addDayoffButton) addDayoffButton.addEventListener('click', addDayoff);
  if (generateButton) generateButton.addEventListener('click', generateShift);
  if (exportCsvButton) exportCsvButton.addEventListener('click', exportToCSV);
  if (modalSaveBtn) modalSaveBtn.addEventListener('click', handleModalSave);
  if (modalCancelBtn) modalCancelBtn.addEventListener('click', closeStaffModal);
  if (staffModal) {
    staffModal.addEventListener('click', event => {
      if (event.target === staffModal) {
        closeStaffModal();
      }
    });
  }

});
