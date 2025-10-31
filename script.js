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

  const SAMPLE_OUTPUTS = {
    dec2025: {
      label: '2025年12月（現場データ・日本語表記）',
      path: 'frontend/public/samples/output_dec2025.json',
    },
  };

  // 週インデックス（Mon–Sun）: 0=第1週
  function getWeekIndex(year, month, day) {
    const d = new Date(year, month - 1, day);
    const dayOfWeek = d.getDay();
    const first = new Date(year, month - 1, 1);
    const offset = ((first.getDay() + 6) % 7);
    const nth = Math.floor((day - 1 + offset) / 7);
    return nth;
  }

  // 各スタッフの使用カウンタ（週/⽉）を初期化
  function buildUsageTrackers(staffRecords, year, month) {
    const daysInMonth = new Date(year, month, 0).getDate();
    const usage = new Map();
    staffRecords.forEach(r => {
      usage.set(r.staffObject.id, {
        month: 0,
        week: {},
      });
    });
    return { usage, daysInMonth };
  }

  function getWeeklyUsed(usage, staffId, weekIndex) {
    const u = usage.get(staffId);
    return u.week[weekIndex] || 0;
  }

  function incUsage(usage, staffId, weekIndex) {
    const u = usage.get(staffId);
    u.month += 1;
    u.week[weekIndex] = (u.week[weekIndex] || 0) + 1;
  }

  // その日が固定休/希望休ならtrue
  function isHardOffDay(record, dateStr) {
    const dow = new Date(dateStr).getDay();
    const fixed = record.fixedHolidays || [];
    const fixedHit = fixed.includes(String(dow));
    const reqOff = isDayOff(record.staffObject, dateStr);
    return !!(fixedHit || reqOff);
  }

  // 夜勤Aの翌2日、夜勤Cの翌1⽇を休みにできるか（セル空/未ロック/既休でOK）を確認
  function canReserveRestWindow(record, dayIndex, days, daysInMonth) {
    for (let i = 1; i <= days; i++) {
      const idx = dayIndex + i;
      if (idx >= daysInMonth) return false;
      const cell = record.cells[idx];
      if (!cell) return false;
      if (cell.assignment && cell.assignment !== '休み') return false;
    }
    return true;
  }

  function applyRestWindow(record, dayIndex, days) {
    for (let i = 1; i <= days; i++) {
      const cell = record.cells[dayIndex + i];
      if (!cell) continue;
      if (!cell.assignment || cell.assignment !== '休み') {
        markForcedRest(cell);
      }
    }
  }

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

  const sampleLoadButtons = document.querySelectorAll('[data-sample-output]');

  const generateButton = document.getElementById('generate-btn');
  const exportCsvButton = document.getElementById('export-csv-btn');
  const exportXlsxButton = document.getElementById('export-xlsx-btn');
  const csvOutput = document.getElementById('csv-output');

  const staffModal = document.getElementById('staff-modal');
  const modalTitle = document.getElementById('modal-title');
  const modalForm = document.getElementById('modal-form');
  const modalShifts = document.getElementById('modal-shifts');
  const modalWeekdays = document.getElementById('modal-weekdays');
  const modalMinDays = document.getElementById('modal-min-days');
  const modalMaxDays = document.getElementById('modal-max-days');
  const modalMinDaysPerWeek = document.getElementById('modal-min-days-per-week');
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

            const parseNullableNumber = value => {
              if (value === null || value === undefined) return null;
              if (typeof value === 'string' && value.trim() === '') return null;
              const num = Number(value);
              return Number.isFinite(num) ? num : null;
            };

            let minWorkingDays = parseNullableNumber(item.minWorkingDays);
            let maxWorkingDays = parseNullableNumber(item.maxWorkingDays);
            let minDaysPerWeek = parseNullableNumber(item.minDaysPerWeek);
            let maxDaysPerWeek = parseNullableNumber(item.maxDaysPerWeek);

            if (Number.isFinite(maxWorkingDays) && !Number.isFinite(minWorkingDays)) {
              minWorkingDays = 0;
            }
            if (Number.isFinite(maxDaysPerWeek) && !Number.isFinite(minDaysPerWeek)) {
              minDaysPerWeek = 0;
            }

            if (Number.isFinite(minWorkingDays) && Number.isFinite(maxWorkingDays) && minWorkingDays > maxWorkingDays) {
              minWorkingDays = maxWorkingDays;
            }
            if (Number.isFinite(minDaysPerWeek) && Number.isFinite(maxDaysPerWeek) && minDaysPerWeek > maxDaysPerWeek) {
              minDaysPerWeek = maxDaysPerWeek;
            }

            const normalizeRangeValue = (value, upperLimit = null) => {
              if (!Number.isFinite(value)) return null;
              let normalized = Math.max(0, value);
              if (Number.isFinite(upperLimit)) {
                normalized = Math.min(upperLimit, normalized);
              }
              return normalized;
            };

            minWorkingDays = normalizeRangeValue(minWorkingDays);
            maxWorkingDays = normalizeRangeValue(maxWorkingDays);
            minDaysPerWeek = normalizeRangeValue(minDaysPerWeek, 7);
            maxDaysPerWeek = normalizeRangeValue(maxDaysPerWeek, 7);

            return {
              id: item.id || generateStaffId(),
              name: typeof item.name === 'string' ? item.name : '',
              availableShifts: available,
              fixedHolidays,
              minWorkingDays,
              maxWorkingDays,
              minDaysPerWeek,
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
      minWorkingDays: null,
      maxWorkingDays: null,
      minDaysPerWeek: null,
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

  function buildCsvFromResultTable() {
    const resultTable = document.querySelector('#result-area table');
    if (!resultTable) return '';

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

    return rows.join('\n');
  }

  function refreshCsvPreview() {
    const csvOutput = document.getElementById('csv-output');
    if (!csvOutput) return;

    const csvString = buildCsvFromResultTable();
    if (!csvString) {
      csvOutput.value = '';
      csvOutput.style.display = 'none';
      return;
    }

    csvOutput.value = csvString;
    csvOutput.style.display = 'block';
  }

  function exportToCSV() {
    const csvOutput = document.getElementById('csv-output');

    if (!csvOutput) {
      console.error('CSV出力に必要な要素が見つかりません。');
      return;
    }

    const csvString = buildCsvFromResultTable();
    if (!csvString) {
      alert('出力するシフト表がありません。');
      return;
    }

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

  function fillDedicatedNightBCForDay(
    dedicatedRecords,
    allRecords,
    dayIndex,
    dayType,
    nextDayType,
    daysInMonth,
    year,
    month,
    usageObj
  ) {
    if (!Array.isArray(dedicatedRecords) || !dedicatedRecords.length) {
      return;
    }

    const allowed = ['夜勤C', '夜勤B'];
    let assigned = collectAssignmentsForDay(allRecords, dayIndex);
    let supply = createHourlySupplyMap(assigned);
    const needsMap = createHourlyNeedsMap(dayType);

    while (true) {
      const deficit = calculateDeficitMap(needsMap, supply);
      if (!hasPositiveDeficitForShifts(deficit, allowed)) {
        break;
      }

      const avail = collectEligibleRecords(dedicatedRecords, dayIndex, allowed, year, month, usageObj);
      if (!avail.length) {
        break;
      }

      const best = findBestAssignment(
        avail,
        deficit,
        dayIndex,
        allowed,
        dayType,
        nextDayType,
        daysInMonth,
        supply,
        year,
        month,
        usageObj
      );
      if (!best) {
        break;
      }

      const { staffRecord, shift } = best;
      if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
        markForcedRest(staffRecord.cells[dayIndex]);
        continue;
      }

      assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
      if (usageObj && usageObj.usage) {
        const weekIndex = getWeekIndex(year, month, dayIndex + 1);
        incUsage(usageObj.usage, staffRecord.staffObject.id, weekIndex);
      }

      if (shift.name === '夜勤C') {
        applyRestWindow(staffRecord, dayIndex, 1);
      }

      assigned.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
      supply = createHourlySupplyMap(assigned);
    }
  }

  // 朝7-9の最優先充足（>=3名）。昼の候補は夜勤専従を除く集合からのみ選ぶ。
  function fillMorningBand7to9(
    candidateRecords,
    allRecords,
    dayIndex,
    dayType,
    nextDayType,
    daysInMonth,
    year,
    month,
    usageObj
  ) {
    const allowed = ['早番', '日勤A'];
    let assigned = collectAssignmentsForDay(allRecords, dayIndex);
    let supply = createHourlySupplyMap(assigned);

    while (true) {
      const ok7 = (supply[7] || 0) >= 3;
      const ok8 = (supply[8] || 0) >= 3;
      if (ok7 && ok8) break;

      const needsMap = createHourlyNeedsMap(dayType);
      const deficit = calculateDeficitMap(needsMap, supply);
      const avail = collectEligibleRecords(candidateRecords, dayIndex, allowed, year, month, usageObj);
      if (!avail.length) break;

      const best = findBestAssignment(
        avail,
        deficit,
        dayIndex,
        allowed,
        dayType,
        nextDayType,
        daysInMonth,
        supply,
        year,
        month,
        usageObj
      );
      if (!best) break;

      const { staffRecord, shift } = best;
      if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
        markForcedRest(staffRecord.cells[dayIndex]);
        continue;
      }

      assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
      if (usageObj && usageObj.usage) {
        const weekIndex = getWeekIndex(year, month, dayIndex + 1);
        incUsage(usageObj.usage, staffRecord.staffObject.id, weekIndex);
      }

      assigned.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
      supply = createHourlySupplyMap(assigned);
    }
  }

  function fillDaytimeBandsForDay(
    candidateRecords,
    allRecords,
    dayIndex,
    dayType,
    nextDayType,
    daysInMonth,
    year,
    month,
    usageObj
  ) {
    const bands = [
      { hours: [9, 15], allowed: ['日勤A', '日勤B'] },
      { hours: [16, 18], allowed: ['日勤A', '遅番'] },
    ];
    let assigned = collectAssignmentsForDay(allRecords, dayIndex);
    let supply = createHourlySupplyMap(assigned);
    const needsMap = createHourlyNeedsMap(dayType);

    for (const band of bands) {
      while (true) {
        const deficit = calculateDeficitMap(needsMap, supply);
        const [a, b] = band.hours;
        const hasDef = deficit.slice(a, b).some(v => v > 0);
        if (!hasDef) break;

        const avail = collectEligibleRecords(candidateRecords, dayIndex, band.allowed, year, month, usageObj);
        if (!avail.length) break;

        const best = findBestAssignment(
          avail,
          deficit,
          dayIndex,
          band.allowed,
          dayType,
          nextDayType,
          daysInMonth,
          supply,
          year,
          month,
          usageObj
        );
        if (!best) break;

        const { staffRecord, shift } = best;
        if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
          markForcedRest(staffRecord.cells[dayIndex]);
          continue;
        }

        assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
        if (usageObj && usageObj.usage) {
          const weekIndex = getWeekIndex(year, month, dayIndex + 1);
          incUsage(usageObj.usage, staffRecord.staffObject.id, weekIndex);
        }

        assigned.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
        supply = createHourlySupplyMap(assigned);
      }
    }
  }

  function fillRemainingNightShiftsForDay(
    allRecords,
    dayIndex,
    dayType,
    nextDayType,
    daysInMonth,
    year,
    month,
    usageObj
  ) {
    const allowed = ['夜勤A'];
    let assigned = collectAssignmentsForDay(allRecords, dayIndex);
    let supply = createHourlySupplyMap(assigned);
    const needsMap = createHourlyNeedsMap(dayType);

    while (true) {
      const deficit = calculateDeficitMap(needsMap, supply);
      if (!hasPositiveDeficitForShifts(deficit, allowed)) {
        break;
      }

      const avail = collectEligibleRecords(allRecords, dayIndex, allowed, year, month, usageObj);
      if (!avail.length) {
        break;
      }

      const best = findBestAssignment(
        avail,
        deficit,
        dayIndex,
        allowed,
        dayType,
        nextDayType,
        daysInMonth,
        supply,
        year,
        month,
        usageObj
      );
      if (!best) {
        break;
      }

      const { staffRecord, shift } = best;
      if (!canAssignShift(staffRecord, dayIndex, shift.name)) {
        markForcedRest(staffRecord.cells[dayIndex]);
        continue;
      }

      assignShiftToCell(staffRecord.cells[dayIndex], shift.name);
      if (usageObj && usageObj.usage) {
        const weekIndex = getWeekIndex(year, month, dayIndex + 1);
        incUsage(usageObj.usage, staffRecord.staffObject.id, weekIndex);
      }

      if (shift.name === '夜勤A') {
        applyRestWindow(staffRecord, dayIndex, 2);
      }

      assigned.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
      supply = createHourlySupplyMap(assigned);
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
    const caption = document.createElement('caption');
    caption.textContent = '※スタッフの上限・固定休など全ハード制約を順守した上での不足量です。';
    tbl.appendChild(caption);
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

    return true;
  }

  function collectEligibleRecords(
    staffRecords,
    dayIndex,
    allowedShiftNames = null,
    year = null,
    month = null,
    usageObj = null
  ) {
    return staffRecords.filter(record => {
      const cell = record.cells[dayIndex];
      if (!cell || cell.isLocked || cell.assignment) return false;

      if (isHardOffDay(record, cell.dateStr)) return false;

      const available = Array.isArray(record.staffObject.availableShifts)
        ? record.staffObject.availableShifts
        : [];
      const candidateShifts = allowedShiftNames
        ? available.filter(name => allowedShiftNames.includes(name))
        : available.slice();
      if (!candidateShifts.length) return false;

      if (year != null && month != null && usageObj) {
        const { usage } = usageObj;
        if (usage && usage.has(record.staffObject.id)) {
          const staffId = record.staffObject.id;
          const weekIndex = getWeekIndex(year, month, dayIndex + 1);
          const maxDaysPerWeek = record.staffObject.maxDaysPerWeek;
          const maxMonth = record.staffObject.maxWorkingDays;

          if (
            Number.isFinite(maxDaysPerWeek) &&
            getWeeklyUsed(usage, staffId, weekIndex) >= maxDaysPerWeek
          ) {
            return false;
          }
          if (Number.isFinite(maxMonth) && usage.get(staffId).month >= maxMonth) {
            return false;
          }
        }
      }

      if (countConsecutiveWorkdays(record, dayIndex) >= MAX_CONSECUTIVE_WORKDAYS) {
        return false;
      }

      if (allowedShiftNames && allowedShiftNames.length === 1) {
        const name = allowedShiftNames[0];
        if (name === '夜勤A') {
          if (!canReserveRestWindow(record, dayIndex, 2, record.cells.length)) return false;
        }
        if (name === '夜勤C') {
          if (!canReserveRestWindow(record, dayIndex, 1, record.cells.length)) return false;
        }
      }

      // 将来的に月間・週間の下限設定を導入する場合はここで扱う（現時点では未使用）

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

  function findBestAssignment(
    availableRecords,
    deficitMap,
    dayIndex,
    allowedShiftNames = null,
    dayType = null,
    nextDayType = null,
    daysInMonth = 30,
    currentSupplyMap = null,
    year = null,
    month = null,
    usageObj = null
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

      const staffId = record.staffObject.id;
      const tracker = usageObj ? usageObj.usage : null;
      const weekIndex = year != null && month != null ? getWeekIndex(year, month, dayIndex + 1) : null;
      const usedMonth = tracker && tracker.has(staffId) ? tracker.get(staffId).month : 0;
      const usedWeek = tracker && tracker.has(staffId) && weekIndex != null ? getWeeklyUsed(tracker, staffId, weekIndex) : 0;

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
        if (Number.isFinite(maxDays) && usedMonth >= maxDays) {
          return;
        }

        const maxDaysPerWeek = record.staffObject.maxDaysPerWeek;
        if (Number.isFinite(maxDaysPerWeek) && usedWeek >= maxDaysPerWeek) {
          return;
        }

        if (shiftDefinition.name === '夜勤A') {
          if (!canReserveRestWindow(record, dayIndex, 2, daysInMonth)) {
            return;
          }
        }

        if (shiftDefinition.name === '夜勤C') {
          if (!canReserveRestWindow(record, dayIndex, 1, daysInMonth)) {
            return;
          }
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
          const currentMonthUsage = tracker && tracker.has(staffId) ? tracker.get(staffId).month : 0;
          const bestStaffId = bestMove.staffRecord?.staffObject?.id;
          const bestMonthUsage = tracker && bestStaffId != null && tracker.has(bestStaffId)
            ? tracker.get(bestStaffId).month
            : 0;
          if (currentMonthUsage < bestMonthUsage) {
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
        targetWorkdays: isFiniteNumber(staff.minWorkingDays)
          ? staff.minWorkingDays
          : Math.ceil(daysInMonth * 0.55),
        minWorkdaysGoal: isFiniteNumber(staff.minWorkingDays) ? staff.minWorkingDays : null,
      };
    });

    const dedicatedNightRecords = staffRecords.filter(r => isNightDedicated(r));
    const nightExclusiveRecords = dedicatedNightRecords.filter(record => {
      const shifts = Array.isArray(record.staffObject.availableShifts)
        ? record.staffObject.availableShifts
        : [];
      return shifts.length > 0 && shifts.every(name => name === '夜勤B' || name === '夜勤C');
    });
    const nonDedicatedRecords = staffRecords.filter(r => !isNightDedicated(r));

    // 固定休・希望休を事前に反映
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
    const usageObj = buildUsageTrackers(staffRecords, year, month);

    staffRecords.forEach(record => {
      record.cells.forEach((cellRecord, index) => {
        if (!isWorkingAssignment(cellRecord.assignment)) return;
        const weekIndex = getWeekIndex(year, month, index + 1);
        incUsage(usageObj.usage, record.staffObject.id, weekIndex);
      });
    });

    // Phase 1: 夜勤B/C専従の先行割付
    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      const dayType = getDayType(dayOfWeek);
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;

      fillDedicatedNightBCForDay(
        nightExclusiveRecords,
        staffRecords,
        dayIndex,
        dayType,
        nextDayType,
        daysInMonth,
        year,
        month,
        usageObj
      );
    }

    // Phase 2: 朝帯7-9の最優先充足
    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      const dayType = getDayType(dayOfWeek);
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;

      fillMorningBand7to9(
        nonDedicatedRecords,
        staffRecords,
        dayIndex,
        dayType,
        nextDayType,
        daysInMonth,
        year,
        month,
        usageObj
      );
    }

    // Phase 3: 9-15帯の充足
    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      const dayType = getDayType(dayOfWeek);
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;

      fillDaytimeBandsForDay(
        nonDedicatedRecords,
        staffRecords,
        dayIndex,
        dayType,
        nextDayType,
        daysInMonth,
        year,
        month,
        usageObj
      );
    }

    // Phase 4: 残り夜勤（主に夜勤A）
    for (let day = 1; day <= daysInMonth; day++) {
      const dayIndex = day - 1;
      const currentDate = new Date(year, month - 1, day);
      const dayOfWeek = currentDate.getDay();

      const dayType = getDayType(dayOfWeek);
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;

      fillRemainingNightShiftsForDay(
        staffRecords,
        dayIndex,
        dayType,
        nextDayType,
        daysInMonth,
        year,
        month,
        usageObj
      );
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

    refreshCsvPreview();
  }

  async function loadSampleOutput(sampleKey) {
    const definition = SAMPLE_OUTPUTS[sampleKey];
    if (!definition) {
      console.warn('未定義のサンプルキーです:', sampleKey);
      return;
    }

    try {
      const response = await fetch(definition.path);
      if (!response.ok) {
        throw new Error(`Failed to fetch sample: ${response.status}`);
      }
      const data = await response.json();
      applySampleResult(data);
    } catch (error) {
      console.error('サンプルの読み込みに失敗しました', error);
      alert('サンプルの読み込みに失敗しました。');
    }
  }

  function applySampleResult(data) {
    if (!data || !Array.isArray(data.assignments)) {
      console.warn('サンプルデータの形式が不正です。');
      return;
    }

    if (typeof data.year === 'number' && typeof data.month === 'number') {
      state.targetYear = data.year;
      state.targetMonth = data.month;
    }

    renderHeader();

    const tbody = document.getElementById('result-body');
    if (!tbody) return;
    tbody.innerHTML = '';

    const daysInMonth = new Date(state.targetYear, state.targetMonth, 0).getDate();

    data.assignments.forEach((assignment, index) => {
      const tr = document.createElement('tr');
      const nameCell = document.createElement('td');
      const displayName =
        assignment.displayName ||
        assignment.staffName ||
        assignment.staffId ||
        assignment.id ||
        `スタッフ${index + 1}`;
      nameCell.textContent = displayName;
      tr.appendChild(nameCell);

      const shifts = Array.isArray(assignment.shifts) ? assignment.shifts : [];
      for (let dayIndex = 0; dayIndex < daysInMonth; dayIndex++) {
        const td = document.createElement('td');
        td.textContent = shifts[dayIndex] || '';
        tr.appendChild(td);
      }

      tbody.appendChild(tr);
    });

    const shortageRows = Array.isArray(data.shortageSummary) ? data.shortageSummary : [];
    renderShortageTable(shortageRows);
    const shortageBtn = document.getElementById('export-shortage-csv-btn');
    if (shortageBtn) {
      shortageBtn.onclick = () => exportShortageCSV(shortageRows);
    }

    refreshCsvPreview();
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
    if (modalMinDays) {
      modalMinDays.value = staff.minWorkingDays != null ? staff.minWorkingDays : '';
    }
    if (modalMaxDays) {
      modalMaxDays.value = staff.maxWorkingDays != null ? staff.maxWorkingDays : '';
    }

    if (modalMinDaysPerWeek) {
      modalMinDaysPerWeek.value = staff.minDaysPerWeek != null ? staff.minDaysPerWeek : '';
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

    const toNumOrNull = value => {
      if (value == null) return null;
      const trimmed = String(value).trim();
      if (trimmed === '') return null;
      const parsed = Number(trimmed);
      return Number.isFinite(parsed) ? parsed : null;
    };

    const sanitizeMonthValue = value => {
      if (value == null) return null;
      return Math.max(0, value);
    };

    const sanitizeWeekValue = value => {
      if (value == null) return null;
      return Math.min(7, Math.max(0, value));
    };

    let minMonthValue = modalMinDays ? sanitizeMonthValue(toNumOrNull(modalMinDays.value)) : null;
    let maxMonthValue = modalMaxDays ? sanitizeMonthValue(toNumOrNull(modalMaxDays.value)) : null;
    let minWeekValue = modalMinDaysPerWeek ? sanitizeWeekValue(toNumOrNull(modalMinDaysPerWeek.value)) : null;
    let maxWeekValue = modalMaxDaysPerWeek ? sanitizeWeekValue(toNumOrNull(modalMaxDaysPerWeek.value)) : null;

    if (minMonthValue != null && maxMonthValue != null && minMonthValue > maxMonthValue) {
      alert('月間の下限は上限以下にしてください');
      return;
    }

    if (minWeekValue != null && maxWeekValue != null && minWeekValue > maxWeekValue) {
      alert('週の下限は上限以下にしてください');
      return;
    }

    staff.minWorkingDays = minMonthValue;
    staff.maxWorkingDays = maxMonthValue;
    staff.minDaysPerWeek = minWeekValue;
    staff.maxDaysPerWeek = maxWeekValue;

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
  sampleLoadButtons.forEach(button => {
    button.addEventListener('click', () => {
      const key = button.getAttribute('data-sample-output');
      if (key) {
        loadSampleOutput(key);
      }
    });
  });
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
