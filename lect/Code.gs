const CONFIG = {
  LECT_SPREADSHEET_ID: '1U5kHU7VnNdOLyaqoUqzqqVb3SdZDW9IpmNZZYOayjQc',
  STUDENT_SPREADSHEET_ID: '1YQocFXT5NyWvv6YzFy4NfMFXahkjl2KKUmBEgk-Hdzs',
  SHEET_LECT: 'DATA_LECT',
  SHEET_CURRENT_STUDENTS: 'current_students',
  SHEET_INACTIVE_STUDENTS: 'inactive_students',
  SHEET_RATIO: 'RATIO_T_S',
  TARGET_RATIO: 8,
  DEFAULT_PLANNED_INTAKE: 100,
  CACHE_SECONDS: 300
};

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'dashboard';

    if (action !== 'dashboard') {
      return jsonOutput_({
        success: false,
        message: 'Unknown action'
      });
    }

    const cache = CacheService.getScriptCache();
    const cacheKey = 'lecturer_student_dashboard_v1';
    const cached = cache.get(cacheKey);

    if (cached) {
      return jsonOutput_(JSON.parse(cached));
    }

    const result = getDashboardData_();
    cache.put(cacheKey, JSON.stringify(result), CONFIG.CACHE_SECONDS);
    return jsonOutput_(result);
  } catch (err) {
    return jsonOutput_({
      success: false,
      message: err.message || String(err)
    });
  }
}

function getDashboardData_() {
  const lectSs = SpreadsheetApp.openById(CONFIG.LECT_SPREADSHEET_ID);
  const studentSs = SpreadsheetApp.openById(CONFIG.STUDENT_SPREADSHEET_ID);

  const lectSheet = lectSs.getSheetByName(CONFIG.SHEET_LECT);
  const ratioSheet = lectSs.getSheetByName(CONFIG.SHEET_RATIO);
  const currentStudentSheet = studentSs.getSheetByName(CONFIG.SHEET_CURRENT_STUDENTS);
  const inactiveStudentSheet = studentSs.getSheetByName(CONFIG.SHEET_INACTIVE_STUDENTS);

  if (!lectSheet) throw new Error('ไม่พบชีต ' + CONFIG.SHEET_LECT);
  if (!currentStudentSheet) throw new Error('ไม่พบชีต ' + CONFIG.SHEET_CURRENT_STUDENTS);

  const lecturers = readSheetObjects_(lectSheet);
  const currentStudents = readSheetObjects_(currentStudentSheet);
  const inactiveStudents = inactiveStudentSheet ? readSheetObjects_(inactiveStudentSheet) : [];

  const activeLecturerCount = lecturers.filter(isActiveLecturer_).length;
  const activeStudentRows = currentStudents.filter(isCurrentStudent_);
  const currentStudentCount = activeStudentRows.length;
  const plannedIntake = readPlannedIntake_(ratioSheet);

  const ratioValue = activeLecturerCount > 0 ? currentStudentCount / activeLecturerCount : 0;
  const intakeCapacityWithoutHiring = Math.max(
    0,
    (activeLecturerCount * CONFIG.TARGET_RATIO) - currentStudentCount
  );
  const additionalLecturersRequired = Math.max(
    0,
    Math.ceil((currentStudentCount + plannedIntake) / CONFIG.TARGET_RATIO) - activeLecturerCount
  );

  const studentsByCohort = buildStudentsByCohort_(activeStudentRows);

  return {
    success: true,
    generatedAt: new Date().toISOString(),
    kpis: {
      totalActiveLecturers: activeLecturerCount,
      totalCurrentStudents: currentStudentCount,
      currentRatioValue: Number(ratioValue.toFixed(2)),
      currentRatioDisplay: activeLecturerCount > 0 ? '1 : ' + ratioValue.toFixed(2) : 'ไม่มีอาจารย์',
      intakeCapacityWithoutHiring: intakeCapacityWithoutHiring,
      plannedNewIntake: plannedIntake,
      additionalLecturersRequired: additionalLecturersRequired,
      targetRatio: CONFIG.TARGET_RATIO
    },
    charts: {
      studentsByCohort: studentsByCohort,
      cohortComparison: buildCohortComparison_(studentsByCohort),
      retentionByCohort: buildRetentionByCohort_(currentStudents, inactiveStudents)
    }
  };
}

function isActiveLecturer_(row) {
  return normalizeBool_(row.ACTIVE) && String(row.PERSON_STATUS || '').trim() === 'อาจารย์';
}

function isCurrentStudent_(row) {
  const studentId = String(row['รหัสนักศึกษา'] || '').trim();
  const isActive = normalizeBool_(row.STUDENT_ACTIVE);
  const statusGroup = String(row.STATUS_GROUP || '').trim().toUpperCase();
  return studentId !== '' && isActive && (!statusGroup || statusGroup === 'ACTIVE');
}

function buildStudentsByCohort_(rows) {
  const counts = {};

  rows.forEach(function(row) {
    const cohort = normalizeCohort_(row.COHORT_YEAR || row.CLASS_YEAR || 'ไม่ระบุ');
    counts[cohort] = (counts[cohort] || 0) + 1;
  });

  return Object.keys(counts)
    .sort(sortCohorts_)
    .map(function(cohort) {
      return {
        cohort: cohort,
        currentCount: counts[cohort]
      };
    });
}

function buildCohortComparison_(cohorts) {
  return cohorts
    .slice()
    .sort(function(a, b) {
      return sortCohorts_(a.cohort, b.cohort);
    })
    .slice(-2)
    .map(function(item) {
      return {
        cohort: item.cohort,
        count: item.currentCount
      };
    });
}

function buildRetentionByCohort_(currentRows, inactiveRows) {
  const summary = {};

  currentRows.forEach(function(row) {
    const studentId = String(row['รหัสนักศึกษา'] || '').trim();
    if (!studentId) return;

    const cohort = normalizeCohort_(row.COHORT_YEAR || row.CLASS_YEAR || 'ไม่ระบุ');
    ensureCohortSummary_(summary, cohort);
    summary[cohort].originalIds[studentId] = true;
    if (isCurrentStudent_(row)) summary[cohort].currentIds[studentId] = true;
  });

  inactiveRows.forEach(function(row) {
    const studentId = String(row['รหัสนักศึกษา'] || '').trim();
    if (!studentId) return;

    const cohort = normalizeCohort_(row.COHORT_YEAR || row.CLASS_YEAR || 'ไม่ระบุ');
    ensureCohortSummary_(summary, cohort);
    summary[cohort].originalIds[studentId] = true;
  });

  return Object.keys(summary)
    .sort(sortCohorts_)
    .map(function(cohort) {
      const originalCount = Object.keys(summary[cohort].originalIds).length;
      const currentCount = Object.keys(summary[cohort].currentIds).length;
      const retentionRate = originalCount > 0
        ? Number(((currentCount / originalCount) * 100).toFixed(2))
        : 0;

      return {
        cohort: cohort,
        originalCount: originalCount,
        currentCount: currentCount,
        retentionRate: retentionRate
      };
    });
}

function ensureCohortSummary_(summary, cohort) {
  if (!summary[cohort]) {
    summary[cohort] = {
      originalIds: {},
      currentIds: {}
    };
  }
}

function readPlannedIntake_(ratioSheet) {
  if (!ratioSheet) return CONFIG.DEFAULT_PLANNED_INTAKE;

  const values = ratioSheet.getDataRange().getValues();
  for (var i = 0; i < values.length; i++) {
    const key = String(values[i][0] || '').trim().toLowerCase();
    const value = toNumber_(values[i][1]);

    if (
      key.indexOf('planned new intake') !== -1 ||
      key.indexOf('แผนการรับนักศึกษาใหม่') !== -1 ||
      key.indexOf('แผนรับนักศึกษาใหม่') !== -1
    ) {
      return value || CONFIG.DEFAULT_PLANNED_INTAKE;
    }
  }

  return CONFIG.DEFAULT_PLANNED_INTAKE;
}

function readSheetObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(normalizeHeader_);

  return values.slice(1)
    .filter(function(row) {
      return row.some(function(cell) {
        return cell !== '' && cell !== null;
      });
    })
    .map(function(row) {
      const obj = {};
      headers.forEach(function(header, index) {
        if (!header) return;
        obj[header] = row[index];
      });
      return obj;
    });
}

function normalizeHeader_(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeCohort_(value) {
  const text = String(value || '').trim();
  const match = text.match(/\d+/);
  if (!match) return text || 'ไม่ระบุ';

  const numberValue = Number(match[0]);
  if (numberValue >= 2500) return String(numberValue - 2500);
  return String(numberValue);
}

function sortCohorts_(a, b) {
  const numberA = Number(a);
  const numberB = Number(b);
  if (isFinite(numberA) && isFinite(numberB)) return numberA - numberB;
  return String(a).localeCompare(String(b), 'th');
}

function normalizeBool_(value) {
  if (value === true) return true;
  if (value === false) return false;

  const text = String(value || '').trim().toUpperCase();
  return ['TRUE', 'YES', 'Y', '1', 'ACTIVE'].indexOf(text) !== -1;
}

function toNumber_(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return value;

  const numberValue = Number(String(value).replace(/,/g, '').trim());
  return isNaN(numberValue) ? 0 : numberValue;
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
