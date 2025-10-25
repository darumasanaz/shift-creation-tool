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
  const NIGHT_STRICT_HOURS = [21, 22, 23, 0, 1, 2, 3, 4, 5, 6];
  const EVENING_HOURS = [18, 19, 20];
  const WEIGHT_DAYB_CORE = 10;
  const WEIGHT_LATE_CORE = 12;
  const WEIGHT_DAY_DEFICIT = 8;
  const WEIGHT_DAYB_DEFICIT = 10;
  const PENALTY_GEN_NIGHTS = 8;
  const DAY_OVERSUP_ALLOW = 1;
  const DAY_OVERSUP_PENALTY = 5;
  const GENERALIST_NIGHT_THRESHOLD = 6;
  const NIGHT_SPECIALIST_BONUS = 6;
  const GENERALIST_NIGHT_ROLE_PENALTY = 4;
  const NIGHT_PRIORITY_ORDER = ['夜勤C', '夜勤B', '夜勤A'];

  const DAYTIME_SHIFTS = ['早番', '日勤A', '日勤B', '遅番'];
  const SHIFT_PATTERNS = SHIFT_DEFINITIONS.map(pattern => pattern.name);
  const ROLE_CATEGORIES = {
    nightSpecialist: 'night-specialist',
    daybCore: 'dayb-core',
    lateCore: 'late-core',
    generalist: 'generalist',
    dayOnly: 'day-only',
  };
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

  function inferRoleCategory(staff) {
    if (!staff) {
      return ROLE_CATEGORIES.dayOnly;
    }

    const available = Array.isArray(staff.availableShifts)
      ? staff.availableShifts.filter(shift => SHIFT_PATTERNS.includes(shift))
      : [];
    if (!available.length) {
      return ROLE_CATEGORIES.dayOnly;
    }

    const dayShifts = available.filter(name => DAYTIME_SHIFTS.includes(name));
    const nightShifts = available.filter(name => NIGHT_SHIFTS.includes(name));

    if (dayShifts.length === 0 && nightShifts.length > 0) {
      return ROLE_CATEGORIES.nightSpecialist;
    }

    const lateCount = dayShifts.filter(name => name === '遅番').length;
    if (lateCount > 0) {
      const dayOnlyLate = dayShifts.length === lateCount;
      const lateDominant = lateCount >= Math.max(1, dayShifts.length - 1);
      if ((dayOnlyLate || lateDominant) && nightShifts.length <= 1) {
        return ROLE_CATEGORIES.lateCore;
      }
    }

    if (dayShifts.includes('日勤B')) {
      const strictMonthlyCap = isFiniteNumber(staff.maxWorkingDays) && staff.maxWorkingDays <= 10;
      if (nightShifts.length === 0 || strictMonthlyCap) {
        return ROLE_CATEGORIES.daybCore;
      }
    }

    if (dayShifts.length > 0 && nightShifts.length > 0) {
      return ROLE_CATEGORIES.generalist;
    }

    return ROLE_CATEGORIES.dayOnly;
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
      const isNightWindow =
        (normalizedHour >= 18 && normalizedHour <= 23) || (isNextMorning && normalizedHour >= 0 && normalizedHour <= 6);
      if (!isNightWindow) {
        continue;
      }
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
    let daytimeOversupplyHours = 0;
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
          daytimeOversupplyHours += 1;
        }
      }

      if (issues.length) {
        console.warn(`Coverage issues on day ${dayIndex + 1}:`, issues);
      }
    }

    if (daytimeOversupplyHours > 0) {
      console.warn(
        `Daytime oversupply hours beyond need + ${DAY_OVERSUP_ALLOW}: ${daytimeOversupplyHours}`
      );
    }

    console.info(
      `Daytime shortages (staff-hours): ${daytimeShortageSlots}, oversupply beyond need + ${DAY_OVERSUP_ALLOW} (staff-hours): ${daytimeOversupplySlots}, night coverage warnings: ${nightViolations}`
    );
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

      const roleCategory = record.roleCategory || inferRoleCategory(record.staffObject);
      record.roleCategory = roleCategory;

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
        let dayDeficitCovered = false;
        let daybDeficitWeight = 0;
        let dayOversupplyPenalty = 0;
        let lateCoreBonus = 0;
        let nightContribution = 0;

        for (let hour = shiftDefinition.start; hour < shiftDefinition.end; hour++) {
          const normalizedHour = ((hour % 24) + 24) % 24;
          const isNextMorning = hour >= 24;
          const hourDayType = isNextMorning ? nextDayType : dayType;
          const { min } = getMinMaxForHour(hourDayType, normalizedHour, isNextMorning);
          const currentCoverage = currentSupplyMap ? currentSupplyMap[normalizedHour] || 0 : 0;
          const afterCoverage = currentCoverage + 1;
          const deficitValue = Math.max(0, deficitMap[normalizedHour] || 0);
          const isDaytimeHour = !isNextMorning && normalizedHour >= 7 && normalizedHour <= 17;
          const isLateWindow = !isNextMorning && normalizedHour >= 16 && normalizedHour <= 18;
          const isNightWindow = !isDaytimeHour && (normalizedHour >= 18 || isNextMorning || normalizedHour <= 6);

          if (isDaytimeHour) {
            if (currentCoverage < min) {
              const shortage = min - currentCoverage;
              dayDeficitCovered = true;
              score += WEIGHT_DAY_DEFICIT * shortage;
              if (shiftDefinition.name === '日勤B') {
                daybDeficitWeight += WEIGHT_DAYB_DEFICIT * shortage;
              }
            }

            const overAmount = afterCoverage - (min + DAY_OVERSUP_ALLOW);
            if (overAmount > 0) {
              dayOversupplyPenalty += DAY_OVERSUP_PENALTY * overAmount;
            }

            if (
              !lateCoreBonus &&
              shiftDefinition.name === '遅番' &&
              roleCategory === ROLE_CATEGORIES.lateCore &&
              isLateWindow &&
              deficitValue > 0
            ) {
              lateCoreBonus = WEIGHT_LATE_CORE;
            }
          } else if (isNightWindow && min > 0) {
            const improvement = Math.min(afterCoverage, min) - currentCoverage;
            if (improvement > 0) {
              const weight = normalizedHour >= 21 || isNextMorning ? 50 : 25;
              nightContribution += improvement * weight;
            }
            if (deficitValue > 0) {
              nightContribution += deficitValue * 5;
            }
          }
        }

        score += daybDeficitWeight;
        score -= dayOversupplyPenalty;
        score += lateCoreBonus;
        score += nightContribution;

        if (shiftDefinition.name === '日勤B' && roleCategory === ROLE_CATEGORIES.daybCore && dayDeficitCovered) {
          score += WEIGHT_DAYB_CORE;
        }

        if (NIGHT_SHIFTS.includes(shiftDefinition.name)) {
          const priorityIndex = NIGHT_PRIORITY_ORDER.indexOf(shiftDefinition.name);
          if (priorityIndex !== -1) {
            score += (NIGHT_PRIORITY_ORDER.length - priorityIndex) * 2;
          }
          if (roleCategory === ROLE_CATEGORIES.nightSpecialist) {
            score += NIGHT_SPECIALIST_BONUS;
          } else if (roleCategory === ROLE_CATEGORIES.generalist) {
            score -= GENERALIST_NIGHT_ROLE_PENALTY;
            const nightCount = record.nightShiftsAssigned || 0;
            const projectedNightCount = nightCount + 1;
            const overNights = Math.max(0, projectedNightCount - GENERALIST_NIGHT_THRESHOLD);
            if (overNights > 0) {
              score -= PENALTY_GEN_NIGHTS * overNights;
            }
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
        nightShiftsAssigned: 0,
        roleCategory: inferRoleCategory(staff),
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
      const nextDayOfWeek = day < daysInMonth ? new Date(year, month - 1, day + 1).getDay() : null;
      const nextDayType = nextDayOfWeek != null ? getDayType(nextDayOfWeek) : null;
      const needsMap = createHourlyNeedsMap(dayType);
      let assignedShiftsThisDay = collectAssignmentsForDay(staffRecords, dayIndex);
      let supplyMap = createHourlySupplyMap(assignedShiftsThisDay);

      while (true) {
        const deficitMap = calculateDeficitMap(needsMap, supplyMap);
        const nightStatus = calculateNightCoverageNeeds(supplyMap, dayType, nextDayType);
        if (!nightStatus.strictBelowMin && !nightStatus.eveningBelowMin) {
          break;
        }

        const allowedShifts = nightStatus.strictBelowMin ? NIGHT_SHIFTS : ['遅番'];
        const availableRecords = collectEligibleRecords(staffRecords, dayIndex, allowedShifts);
        if (!availableRecords.length) {
          break;
        }

        const bestMove = findBestAssignment(
          availableRecords,
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
        if (NIGHT_SHIFTS.includes(shift.name)) {
          staffRecord.nightShiftsAssigned = (staffRecord.nightShiftsAssigned || 0) + 1;
        }

        assignedShiftsThisDay.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
        supplyMap = createHourlySupplyMap(assignedShiftsThisDay);

        if (NIGHT_SHIFTS.includes(shift.name)) {
          const nextIndex = dayIndex + 1;
          if (nextIndex < daysInMonth) {
            const nextCell = staffRecord.cells[nextIndex];
            if (nextCell && !nextCell.isLocked) {
              markNightShiftRest(nextCell);
            }
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
        if (NIGHT_SHIFTS.includes(shift.name)) {
          staffRecord.nightShiftsAssigned = (staffRecord.nightShiftsAssigned || 0) + 1;
        }

        assignedShiftsThisDay.push({ staff: staffRecord.staffObject, shift, record: staffRecord });
        supplyMap = createHourlySupplyMap(assignedShiftsThisDay);
      }
    }

    validateSchedule(staffRecords, daysInMonth, year, month);

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
