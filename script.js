document.addEventListener('DOMContentLoaded', function () {
  'use strict';

  function safeRun(fn) {
    try {
      if (typeof fn === 'function') {
        return fn();
      }
    } catch (e) {
      console.warn('[init skipped]', e);
    }
    return undefined;
  }

  window.addEventListener('error', e => {
    console.warn('[global error]', (e && e.message) || e);
  });

  const SHIFT_DEFINITIONS = [
    { name: '早番', start: 7, end: 16 },
    { name: '日勤A', start: 9, end: 18 },
    { name: '日勤B', start: 9, end: 16 },
    { name: '遅番', start: 15, end: 21 },
    { name: '夜勤A', start: 16, end: 33 },
    { name: '夜勤B', start: 18, end: 33 },
    { name: '夜勤C', start: 21, end: 31 },
  ];

  // 夜勤専従判定：夜勤B or 夜勤C を持っているスタッフは常に夜勤専従として扱う
  function isNightDedicated(recordOrStaff) {
    const s = recordOrStaff && recordOrStaff.staffObject ? recordOrStaff.staffObject : recordOrStaff;
    const shifts = Array.isArray(s?.availableShifts) ? s.availableShifts : [];
    return shifts.includes('夜勤B') || shifts.includes('夜勤C');
  }

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
  const NIGHT_STRICT_HOURS = [21, 22, 23, 0, 1, 2, 3, 4, 5, 6];
  const EVENING_HOURS = [18, 19, 20];
  const DAY_NEED_WEIGHT = 10;
  const DAY_OVERSUP_ALLOW = 1;
  const DAY_OVERSUP_PENALTY = 4;
  const NIGHT_PRIORITY_ORDER = ['夜勤C', '夜勤B', '夜勤A'];

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
  const clearDayoffButton = document.getElementById('clear-dayoff-button');

  const generateButton = document.getElementById('generate-btn');
  const exportCsvButton = document.getElementById('export-csv-btn');
  const exportXlsxButton = document.getElementById('export-xlsx-btn');
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
    // console.debug('[addStaff] clicked name=', name);
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

    state.dayoffs.forEach((dayoff, idx) => {
      const li = document.createElement('li');

      const staff = state.staff.find(item => item.id === dayoff.staffId);
      const staffName = dayoff.staffName || (staff ? staff.name : '不明なスタッフ');

      li.innerHTML = `
      <span>${staffName} - ${dayoff.date}</span>
      <button
        type="button"
        class="delete-btn delete-dayoff-btn"
        data-index="${idx}"
        aria-label="${staffName} の ${dayoff.date} の希望休を削除"
      >削除</button>
    `;

      dayoffList.appendChild(li);
    });
  }

  function handleDayoffListClick(event) {
    const btn = event.target.closest('.delete-dayoff-btn');
    if (!btn) return;

    const idx = Number(btn.dataset.index);
    if (!Number.isInteger(idx) || idx < 0 || idx >= state.dayoffs.length) return;

    const target = state.dayoffs[idx];
    const staff = state.staff.find(s => s.id === target.staffId);
    const staffName = target.staffName || (staff ? staff.name : '不明なスタッフ');
    const ok = window.confirm(`${staffName} の ${target.date} の希望休を削除しますか？`);
    if (!ok) return;

    state.dayoffs.splice(idx, 1);
    renderDayoffList();
    saveState();
  }

  function handleClearDayoffs() {
    if (!state.dayoffs.length) return;
    const ok = window.confirm(`登録済みの希望休（${state.dayoffs.length}件）を全て削除します。よろしいですか？`);
    if (!ok) return;

    state.dayoffs = [];
    renderDayoffList();
    saveState();
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

  function exportToXLSX() {
    const resultTable = document.querySelector('#result-area table');
    if (!resultTable) {
      console.error('Excel出力に必要な要素が見つかりません。');
      return;
    }

    const y = state.targetYear ?? '';
    const m = state.targetMonth != null ? String(state.targetMonth).padStart(2, '0') : '';
    const filename = `シフト_${y}-${m}.xlsx`;

    if (typeof XLSX !== 'undefined' && XLSX && XLSX.utils && XLSX.writeFile) {
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.table_to_sheet(resultTable);
      XLSX.utils.book_append_sheet(wb, ws, 'シフト');
      XLSX.writeFile(wb, filename);
      return;
    }

    const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body>${resultTable.outerHTML}</body></html>`;
    const blob = new Blob([html], { type: 'application/vnd.ms-excel;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename.replace(/\.xlsx$/i, '.xls');
    a.click();
    URL.revokeObjectURL(url);
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

  const DAYTIME_NEEDS_CACHE = new Map();

  function getCachedDaytimeNeeds(dayType) {
    if (!dayType) {
      return null;
    }
    if (!DAYTIME_NEEDS_CACHE.has(dayType)) {
      DAYTIME_NEEDS_CACHE.set(dayType, createHourlyNeedsMap(dayType));
    }
    return DAYTIME_NEEDS_CACHE.get(dayType);
  }

  function getMinMaxForHour(dayType, hour, isNextMorning = false) {
    if (hour == null) {
      return { min: 0, max: Number.POSITIVE_INFINITY };
    }

    const normalized = ((hour % 24) + 24) % 24;

    if (isNextMorning) {
      if (normalized >= 0 && normalized <= 6) {
        return { min: 2, max: 2 };
      }
      return { min: 0, max: Number.POSITIVE_INFINITY };
    }

    if (normalized >= 21 && normalized <= 23) {
      return { min: 2, max: 2 };
    }

    if (normalized >= 18 && normalized <= 20) {
      return { min: 2, max: 3 };
    }

    if (normalized >= 7 && normalized <= 17) {
      const needsMap = getCachedDaytimeNeeds(dayType);
      const min = needsMap ? needsMap[normalized] || 0 : 0;
      return { min, max: Number.POSITIVE_INFINITY };
    }

    return { min: 0, max: Number.POSITIVE_INFINITY };
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

  function wouldViolateHardCaps(dayIndex, shiftDefinition, context = null) {
    if (!shiftDefinition) {
      return false;
    }
    const start = shiftDefinition.start;
    const end = shiftDefinition.end;
    const normalizedEnd = end > start ? end : end + 24;
    const supplyMap = context && Array.isArray(context.supplyMap) ? context.supplyMap : null;
    const dayType = context ? context.dayType : null;
    const nextDayType = context ? context.nextDayType : null;

    for (let hour = start; hour < normalizedEnd; hour++) {
      const normalizedHour = ((hour % 24) + 24) % 24;
      const isNextMorning = hour >= 24;
      const hourDayType = isNextMorning ? nextDayType : dayType;
      const { min, max } = getMinMaxForHour(hourDayType, normalizedHour, isNextMorning);
      const current = supplyMap ? supplyMap[normalizedHour] || 0 : 0;
      const after = current + 1;

      if (Number.isFinite(max) && after > max) {
        return true;
      }
    }

    return false;
  }

  function isHourCoveredByShift(shiftDefinition, targetHour) {
    if (!shiftDefinition || targetHour == null) {
      return false;
    }
    const start = shiftDefinition.start;
    const end = shiftDefinition.end;
    const normalizedEnd = end > start ? end : end + 24;
    for (let hour = start; hour < normalizedEnd; hour++) {
      if (hour % 24 === targetHour) {
        return true;
      }
    }
    return false;
  }

  function isStrictNightHour(hour) {
    return NIGHT_STRICT_HOURS.includes(hour);
  }

  function isEveningHour(hour) {
    return EVENING_HOURS.includes(hour);
  }

  function calculateNightCoverageNeeds(supplyMap, dayType, nextDayType) {
    const status = {
      strictBelowMin: false,
      eveningBelowMin: false,
    };
    if (!Array.isArray(supplyMap)) {
      return status;
    }

    [21, 22, 23].forEach(hour => {
      const { min } = getMinMaxForHour(dayType, hour, false);
      if ((supplyMap[hour] || 0) < min) {
        status.strictBelowMin = true;
      }
    });

    for (let hour = 0; hour <= 6; hour++) {
      const { min } = getMinMaxForHour(nextDayType, hour, true);
      if ((supplyMap[hour] || 0) < min) {
        status.strictBelowMin = true;
      }
    }

    EVENING_HOURS.forEach(hour => {
      const { min } = getMinMaxForHour(dayType, hour, false);
      if ((supplyMap[hour] || 0) < min) {
        status.eveningBelowMin = true;
      }
    });

    return status;
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

  // 朝7-9の最優先充足（>=3名）。昼の候補は夜勤専従を除く集合からのみ選ぶ。
  function fillMorningBand7to9(candidateRecords, dayIndex, dayType, nextDayType, daysInMonth) {
    const allowed = ['早番', '日勤A'];
    let assigned = collectAssignmentsForDay(candidateRecords, dayIndex);
    let supply = createHourlySupplyMap(assigned);

    while (true) {
      const ok7 = (supply[7] || 0) >= 3;
      const ok8 = (supply[8] || 0) >= 3;
      if (ok7 && ok8) break;

      const needsMap = createHourlyNeedsMap(dayType);
      const deficit = calculateDeficitMap(needsMap, supply);
      const avail = collectEligibleRecords(candidateRecords, dayIndex, allowed);
      if (!avail.length) break;

      const best = findBestAssignment(
        avail,
        deficit,
        dayIndex,
        allowed,
        dayType,
        nextDayType,
        daysInMonth,
        supply
      );
      if (!best) break;

      const { staffRecord, shift } = best;
      if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
        markForcedRest(staffRecord.cells[dayIndex]);
        continue;
      }

      assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
      staffRecord.workdaysInMonth += 1;
      staffRecord.workdaysInWeek += 1;

      assigned.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
      supply = createHourlySupplyMap(assigned);
    }
  }

  function fillDaytimeBandsForDay(candidateRecords, dayIndex, dayType, nextDayType, daysInMonth) {
    const bands = [
      { hours: [9, 15], allowed: ['日勤A', '日勤B'] },
      { hours: [16, 18], allowed: ['日勤A', '遅番'] },
    ];
    let assigned = collectAssignmentsForDay(candidateRecords, dayIndex);
    let supply = createHourlySupplyMap(assigned);
    const needsMap = createHourlyNeedsMap(dayType);

    for (const band of bands) {
      while (true) {
        const deficit = calculateDeficitMap(needsMap, supply);
        const [a, b] = band.hours;
        const hasDef = deficit.slice(a, b).some(v => v > 0);
        if (!hasDef) break;

        const avail = collectEligibleRecords(candidateRecords, dayIndex, band.allowed);
        if (!avail.length) break;

        const best = findBestAssignment(
          avail,
          deficit,
          dayIndex,
          band.allowed,
          dayType,
          nextDayType,
          daysInMonth,
          supply
        );
        if (!best) break;

        const { staffRecord, shift } = best;
        if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
          markForcedRest(staffRecord.cells[dayIndex]);
          continue;
        }

        assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
        staffRecord.workdaysInMonth += 1;
        staffRecord.workdaysInWeek += 1;

        assigned.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
        supply = createHourlySupplyMap(assigned);
      }
    }
  }

  function buildDailyCoverageMap(staffRecords, dayIndex) {
    const map = new Array(24).fill(0);
    const currentAssignments = collectAssignmentsForDay(staffRecords, dayIndex);
    currentAssignments.forEach(entry => {
      const shift = entry.shift;
      if (!shift) return;
      const start = shift.start;
      const end = shift.end;
      const normalizedEnd = end > start ? end : end + 24;
      for (let hour = start; hour < normalizedEnd; hour++) {
        if (hour >= 24) {
          break;
        }
        if (hour < 0) {
          continue;
        }
        const normalizedHour = hour % 24;
        map[normalizedHour] = (map[normalizedHour] || 0) + 1;
      }
    });

    if (dayIndex > 0) {
      const prevAssignments = collectAssignmentsForDay(staffRecords, dayIndex - 1);
      prevAssignments.forEach(entry => {
        const shift = entry.shift;
        if (!shift) return;
        const start = shift.start;
        const end = shift.end;
        const normalizedEnd = end > start ? end : end + 24;
        for (let hour = start; hour < normalizedEnd; hour++) {
          if (hour < 24) {
            continue;
          }
          const normalizedHour = hour - 24;
          if (normalizedHour >= 0 && normalizedHour < 24) {
            map[normalizedHour] = (map[normalizedHour] || 0) + 1;
          }
        }
      });
    }

    return map;
  }

  function validateSchedule(staffRecords, daysInMonth, year, month) {
    let daytimeShortageSlots = 0;
    let daytimeOversupplySlots = 0;
    let nightViolations = 0;

    for (let dayIndex = 0; dayIndex < daysInMonth; dayIndex++) {
      const coverage = buildDailyCoverageMap(staffRecords, dayIndex);
      const currentDate = new Date(year, month - 1, dayIndex + 1);
      const dayType = getDayType(currentDate.getDay());
      const nextDate = dayIndex + 1 < daysInMonth ? new Date(year, month - 1, dayIndex + 2) : null;
      const nextDayType = nextDate ? getDayType(nextDate.getDay()) : null;
      const needsMap = getCachedDaytimeNeeds(dayType) || new Array(24).fill(0);
      const issues = [];

      EVENING_HOURS.forEach(hour => {
        const { min, max } = getMinMaxForHour(dayType, hour, false);
        const value = coverage[hour] || 0;
        if (value < min || (Number.isFinite(max) && value > max)) {
          issues.push({
            type: 'night',
            hour,
            expected: `${min}-${Number.isFinite(max) ? max : '∞'}`,
            actual: value,
          });
          nightViolations += 1;
        }
      });

      [21, 22, 23].forEach(hour => {
        const { min } = getMinMaxForHour(dayType, hour, false);
        const value = coverage[hour] || 0;
        if (value !== min) {
          issues.push({
            type: 'night',
            hour,
            expected: `${min}`,
            actual: value,
          });
          nightViolations += 1;
        }
      });

      if (dayIndex > 0) {
        for (let hour = 0; hour <= 6; hour++) {
          const { min } = getMinMaxForHour(nextDayType, hour, true);
          const value = coverage[hour] || 0;
          if (value !== min) {
            issues.push({
              type: 'night',
              hour,
              expected: `${min}`,
              actual: value,
            });
            nightViolations += 1;
          }
        }
      }

      for (let hour = 7; hour <= 17; hour++) {
        const need = needsMap[hour] || 0;
        const value = coverage[hour] || 0;
        if (value < need) {
          issues.push({
            type: 'day-deficit',
            hour,
            expected: `>=${need}`,
            actual: value,
          });
          daytimeShortageSlots += need - value;
        }

        const allowed = need + DAY_OVERSUP_ALLOW;
        if (value > allowed) {
          issues.push({
            type: 'day-oversupply',
            hour,
            expected: `<=${allowed}`,
            actual: value,
          });
          daytimeOversupplySlots += value - allowed;
        }
      }

      if (issues.length) {
        console.warn(`Coverage issues on day ${dayIndex + 1}:`, issues);
      }
    }

    console.info(
      `Daytime shortages (staff-hours): ${daytimeShortageSlots}, oversupply beyond need + ${DAY_OVERSUP_ALLOW} (staff-hours): ${daytimeOversupplySlots}, night coverage warnings: ${nightViolations}`
    );

    return {
      daytimeShortageSlots,
      daytimeOversupplySlots,
      nightViolations,
    };
  }

  function computeShortageSummary(staffRecords, daysInMonth, year, month) {
    const rows = [];
    for (let dayIndex = 0; dayIndex < daysInMonth; dayIndex++) {
      const coverage = buildDailyCoverageMap(staffRecords, dayIndex);
      const date = new Date(year, month - 1, dayIndex + 1);
      const dayType = getDayType(date.getDay());
      const nextDate = dayIndex + 1 < daysInMonth ? new Date(year, month - 1, dayIndex + 2) : null;
      const nextDayType = nextDate ? getDayType(nextDate.getDay()) : null;

      const bands = [
        { name: '7-9', hours: [7, 9] },
        { name: '9-15', hours: [9, 15] },
        { name: '16-18', hours: [16, 18] },
        { name: '18-21', hours: [18, 21] },
        { name: '21-24', hours: [21, 24] },
        { name: '0-7', hours: [0, 7], nextMorning: true },
      ];

      let totalShort = 0;
      const rec = { day: dayIndex + 1 };

      for (const band of bands) {
        const [a, b] = band.hours;
        let short = 0;
        for (let h = a; h < b; h++) {
          const hour = (h + 24) % 24;
          const isNext = band.nextMorning === true;
          const typ = isNext ? nextDayType : dayType;
          const { min } = getMinMaxForHour(typ, hour, isNext);
          const actual = coverage[hour] || 0;
          if (actual < min) {
            short += min - actual;
          }
        }
        rec[band.name] = short;
        totalShort += short;
      }
      rec.total = totalShort;
      rows.push(rec);
    }
    return rows;
  }

  function renderShortageTable(rows) {
    const tbl = document.getElementById('shortage-table');
    if (!tbl) return;
    tbl.innerHTML = '';
    const thead = document.createElement('thead');
    thead.innerHTML = `<tr>
    <th>日</th><th>7-9</th><th>9-15</th><th>16-18</th><th>18-21</th><th>21-24</th><th>0-7</th><th>合計不足</th>
  </tr>`;
    tbl.appendChild(thead);
    const tbody = document.createElement('tbody');
    rows.forEach(r => {
      const tr = document.createElement('tr');
      tr.innerHTML = `<td>${r.day}</td><td>${r['7-9']}</td><td>${r['9-15']}</td><td>${r['16-18']}</td>
                    <td>${r['18-21']}</td><td>${r['21-24']}</td><td>${r['0-7']}</td><td><b>${r.total}</b></td>`;
      if (r.total > 0) tr.style.backgroundColor = '#fff5f5';
      tbody.appendChild(tr);
    });
    tbl.appendChild(tbody);
  }

  function exportShortageCSV(rows) {
    const header = ['日', '7-9', '9-15', '16-18', '18-21', '21-24', '0-7', '合計不足'];
    const lines = [header.join(',')].concat(
      rows.map(r =>
        [r.day, r['7-9'], r['9-15'], r['16-18'], r['18-21'], r['21-24'], r['0-7'], r.total].join(',')
      )
    );
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = '不足サマリー.csv';
    a.click();
    URL.revokeObjectURL(url);
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

  function isNightOnlyRecord(record) {
    if (!record || !record.staffObject) {
      return false;
    }
    const available = Array.isArray(record.staffObject.availableShifts)
      ? record.staffObject.availableShifts
      : [];
    const canNight = available.some(name => NIGHT_SHIFTS.includes(name));
    const canDay = available.some(name => DAYTIME_SHIFTS.includes(name));
    return canNight && !canDay;
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

          if (!isWorkingAssignment(cell.assignment)) {
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
    dayType = null,
    nextDayType = null,
    daysInMonth = 30,
    currentSupplyMap = null
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

      let candidateShifts = allowedShiftNames
        ? available.filter(name => allowedShiftNames.includes(name))
        : available.slice();
      if (
        allowedShiftNames &&
        allowedShiftNames.length &&
        allowedShiftNames.every(name => NIGHT_SHIFTS.includes(name))
      ) {
        candidateShifts = candidateShifts.sort((a, b) => {
          const aIndex = NIGHT_PRIORITY_ORDER.indexOf(a);
          const bIndex = NIGHT_PRIORITY_ORDER.indexOf(b);
          return (aIndex === -1 ? NIGHT_PRIORITY_ORDER.length : aIndex) -
            (bIndex === -1 ? NIGHT_PRIORITY_ORDER.length : bIndex);
        });
      }
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

        if (
          currentSupplyMap &&
          wouldViolateHardCaps(dayIndex, shiftDefinition, {
            supplyMap: currentSupplyMap,
            dayType,
            nextDayType,
          })
        ) {
          return;
        }

        let score = 0;
        for (let hour = shiftDefinition.start; hour < shiftDefinition.end; hour++) {
          const normalizedHour = ((hour % 24) + 24) % 24;
          const isNextMorning = hour >= 24;
          const hourDayType = isNextMorning ? nextDayType : dayType;
          const { min, max } = getMinMaxForHour(hourDayType, normalizedHour, isNextMorning);
          const currentCoverage = currentSupplyMap ? currentSupplyMap[normalizedHour] || 0 : 0;
          const afterCoverage = currentCoverage + 1;
          const deficit = deficitMap[normalizedHour] || 0;
          if (deficit > 0) {
            score += deficit;
          }

          const isNightStrict = isStrictNightHour(normalizedHour);
          const isEvening = isEveningHour(normalizedHour);

          if ((isNightStrict || isEvening) && min > 0) {
            const improvement = Math.min(afterCoverage, min) - currentCoverage;
            if (improvement > 0) {
              const weight = isNightStrict ? 50 : 25;
              score += improvement * weight;
            }
          }

          if (!isNextMorning && isEvening && currentCoverage >= min && (!Number.isFinite(max) || currentCoverage < max)) {
            score += 5;
          }

          const isDaytimeHour = !isNextMorning && normalizedHour >= 7 && normalizedHour <= 17;
          if (isDaytimeHour) {
            if (currentCoverage < min) {
              score += DAY_NEED_WEIGHT * (min - currentCoverage);
            }
            const overAmount = afterCoverage - (min + DAY_OVERSUP_ALLOW);
            if (overAmount > 0) {
              score -= DAY_OVERSUP_PENALTY * overAmount;
            }
          }
        }

        if (NIGHT_SHIFTS.includes(shiftDefinition.name)) {
          const priorityIndex = NIGHT_PRIORITY_ORDER.indexOf(shiftDefinition.name);
          if (priorityIndex !== -1) {
            score += (NIGHT_PRIORITY_ORDER.length - priorityIndex) * 2;
          }
          const hasDayShifts = (record.staffObject.availableShifts || []).some(name =>
            DAYTIME_SHIFTS.includes(name)
          );
          if (!hasDayShifts) {
            score += 80;
          } else {
            score -= 40;
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

    const dedicatedNightRecords = staffRecords.filter(r => isNightDedicated(r));
    const nonDedicatedRecords = staffRecords.filter(r => !isNightDedicated(r));

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

    // Phase 2: allocate night shifts before daytime coverage
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
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;
      const needsMap = createHourlyNeedsMap(dayType);
      let assignedShiftsThisDay = collectAssignmentsForDay(staffRecords, dayIndex);

      assignedShiftsThisDay.forEach(entry => {
        entry.record.workdaysInMonth += 1;
        entry.record.workdaysInWeek += 1;
      });

      let supplyMap = createHourlySupplyMap(assignedShiftsThisDay);

      while (true) {
        const deficitMap = calculateDeficitMap(needsMap, supplyMap);
        const nightStatus = calculateNightCoverageNeeds(supplyMap, dayType, nextDayType);
        if (!nightStatus.strictBelowMin && !nightStatus.eveningBelowMin) {
          break;
        }

        let allowedShifts = null;
        if (nightStatus.strictBelowMin) {
          allowedShifts = NIGHT_SHIFTS;
        } else if (nightStatus.eveningBelowMin) {
          allowedShifts = ['遅番', '夜勤B', '夜勤A'];
        }
        if (!allowedShifts) {
          break;
        }

        let candidates = collectEligibleRecords(dedicatedNightRecords, dayIndex, allowedShifts);
        if (!candidates.length) {
          candidates = collectEligibleRecords(nonDedicatedRecords, dayIndex, allowedShifts);
        }
        if (!candidates.length) {
          break;
        }

        const bestMove = findBestAssignment(
          candidates,
          deficitMap,
          dayIndex,
          allowedShifts,
          dayType,
          nextDayType,
          daysInMonth,
          supplyMap
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

        if (NIGHT_SHIFTS.includes(shift.name)) {
          const nextIndex = dayIndex + 1;
          if (nextIndex < daysInMonth) {
            const nextCell = staffRecord.cells[nextIndex];
            if (nextCell) {
              if (isWorkingAssignment(nextCell.assignment)) {
                nextCell.assignment = '';
                nextCell.backgroundColor = '';
                nextCell.isLocked = false;
              }
              markNightShiftRest(nextCell);
            }
          }

          if (shift.name === '夜勤A') {
            const next2 = dayIndex + 2;
            if (next2 < daysInMonth) {
              const c2 = staffRecord.cells[next2];
              if (c2 && !c2.isLocked) {
                markNightShiftRest(c2);
              }
            }
          }
        }
      }
    }

    // Phase 3: prioritize morning coverage (7-9) with non-night-dedicated staff
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
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;

      fillMorningBand7to9(nonDedicatedRecords, dayIndex, dayType, nextDayType, daysInMonth);
    }

    // Phase 4: fill remaining daytime bands with non-night-dedicated staff
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
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;

      fillDaytimeBandsForDay(nonDedicatedRecords, dayIndex, dayType, nextDayType, daysInMonth);
    }

    const validation = safeRun(() => validateSchedule(staffRecords, daysInMonth, year, month));
    if (validation) {
      console.debug('Validation summary:', validation);
    }

    const rows = computeShortageSummary(staffRecords, daysInMonth, year, month);
    renderShortageTable(rows);
    const shortageBtn = document.getElementById('export-shortage-csv-btn');
    if (shortageBtn) {
      shortageBtn.onclick = () => exportShortageCSV(rows);
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

  if (addStaffButton) addStaffButton.addEventListener('click', addStaff);
  if (staffList) staffList.addEventListener('click', handleStaffListClick);
  if (addDayoffButton) addDayoffButton.addEventListener('click', addDayoff);
  if (dayoffList) dayoffList.addEventListener('click', handleDayoffListClick);
  if (clearDayoffButton) clearDayoffButton.addEventListener('click', handleClearDayoffs);
  if (generateButton) generateButton.addEventListener('click', generateShift);
  if (exportCsvButton) exportCsvButton.addEventListener('click', exportToCSV);
  if (exportXlsxButton) exportXlsxButton.addEventListener('click', exportToXLSX);
  if (modalSaveBtn) modalSaveBtn.addEventListener('click', handleModalSave);
  if (modalCancelBtn) modalCancelBtn.addEventListener('click', closeStaffModal);
  if (staffModal) {
    staffModal.addEventListener('click', event => {
      if (event.target === staffModal) {
        closeStaffModal();
      }
    });
  }

  safeRun(() => renderHeader());

  safeRun(() => {
    loadState();
  });
  safeRun(() => {
    const targetChanged = setNextMonthTarget();
    if (targetChanged) {
      saveState();
    }
  });
  safeRun(() => renderHeader());

});
