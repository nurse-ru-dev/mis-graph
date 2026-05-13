const BUDGET_SHEET_NAME = 'BUDGET_INPUT';
const BUDGET_HEADERS = [
  'ROW_ID',
  'BUDGET_YEAR',
  'CATEGORY_TYPE',
  'CATEGORY_NAME',
  'ROW_TYPE',
  'TERM_LABEL',
  'COHORT_CODE',
  'STUDY_YEAR',
  'SEMESTER',
  'STUDENT_COUNT',
  'FEE_RATE',
  'AMOUNT',
  'INPUT_AMOUNT',
  'SORT_ORDER',
  'REMARK',
  'IS_ACTIVE',
  'UPDATED_AT'
];

const BUDGET_TEMPLATE_ROWS = [
  ['INCOME', 'รวมรายได้', 'SUMMARY'],
  ['INCOME', 'ค่าธรรมเนียมการศึกษา', 'INPUT'],
  ['INCOME', 'รายได้จากการบริการวิชาการ', 'INPUT'],
  ['INCOME', 'รายได้อื่น ๆ (PN)', 'INPUT'],
  ['INCOME', 'ค่ารับสมัคร', 'INPUT'],
  ['EXPENSE', 'ก.แผนงานจัดการศึกษา (รวม)', 'SUMMARY'],
  ['EXPENSE', 'ก.แผนงานจัดการศึกษา (งบบุคลากร)', 'INPUT'],
  ['EXPENSE', 'ก.แผนงานจัดการศึกษา (การจัดการเรียนการสอน)', 'INPUT'],
  ['EXPENSE', 'ข.แผนงานบริการวิชาการแก่สังคม', 'INPUT'],
  ['EXPENSE', 'ค.แผนงานอนุรักษ์ ส่งเสริมและพัฒนาศาสนา ศิลปะและวัฒนธรรม', 'INPUT'],
  ['EXPENSE', 'ง.แผนงานสนับสนุนวิชาการ', 'INPUT'],
  ['EXPENSE', 'จ.แผนงานวิจัย', 'INPUT'],
  ['EXPENSE', 'รวมรายจ่าย', 'SUMMARY']
];

function doGet(e) {
  return budgetHandle_(e);
}

function doPost(e) {
  return budgetHandle_(e);
}

function budgetHandle_(e) {
  try {
    const payload = budgetPayload_(e);
    const action = String(payload.action || '').trim();
    const aliases = {
      getDashboard: 'getBudgetDashboard',
      saveRow: 'saveBudgetInput',
      deleteRow: 'deleteBudgetInput'
    };
    const normalizedAction = aliases[action] || action;

    switch (normalizedAction) {
      case 'getBudgetYears':
        return budgetJson_({ ok: true, data: budgetGetYears_() });
      case 'getBudgetDashboard':
        return budgetJson_({ ok: true, data: budgetGetDashboard_(payload.budgetYear || payload.year) });
      case 'createBudgetYear':
        return budgetJson_(budgetCreateYear_(payload.year || payload.budgetYear));
      case 'saveBudgetInput':
        return budgetJson_(budgetSaveInput_(payload));
      case 'deleteBudgetInput':
        return budgetJson_(budgetDeleteInput_(payload.rowId));
      case 'cloneBudgetYear':
        return budgetJson_(budgetCloneYear_(payload.sourceYear, payload.targetYear));
      default:
        return budgetJson_({ ok: false, message: 'Unknown action: ' + action });
    }
  } catch (err) {
    return budgetJson_({ ok: false, message: err && err.message ? err.message : String(err) });
  }
}

function budgetPayload_(e) {
  const params = Object.assign({}, e && e.parameter ? e.parameter : {});
  const contents = e && e.postData && e.postData.contents;
  if (contents) {
    try {
      return Object.assign(params, JSON.parse(contents));
    } catch (_) {
      return params;
    }
  }
  return params;
}

function budgetJson_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function budgetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BUDGET_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(BUDGET_SHEET_NAME);
  budgetEnsureHeaders_(sheet);
  return sheet;
}

function budgetEnsureHeaders_(sheet) {
  const firstRow = sheet.getRange(1, 1, 1, BUDGET_HEADERS.length).getValues()[0];
  const hasHeaders = BUDGET_HEADERS.every((header, index) => firstRow[index] === header);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, BUDGET_HEADERS.length).setValues([BUDGET_HEADERS]);
    sheet.setFrozenRows(1);
  }
}

function budgetRows_() {
  const sheet = budgetSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, BUDGET_HEADERS.length).getValues();
  return values.map((row, index) => budgetObjectFromRow_(row, index + 2));
}

function budgetObjectFromRow_(row, rowIndex) {
  const record = {};
  BUDGET_HEADERS.forEach((header, index) => record[header] = row[index]);
  record._rowIndex = rowIndex;
  return record;
}

function budgetGetYears_() {
  const years = {};
  budgetRows_().forEach(row => {
    const year = String(row.BUDGET_YEAR || '').trim();
    if (year && budgetBool_(row.IS_ACTIVE, true)) years[year] = true;
  });
  return Object.keys(years).sort();
}

function budgetGetDashboard_(year) {
  const budgetYear = String(year || '').trim();
  if (!budgetYear) throw new Error('Missing budgetYear');

  const rows = budgetRows_()
    .filter(row => String(row.BUDGET_YEAR || '').trim() === budgetYear)
    .map(budgetNormalizeRow_);

  return {
    budgetYear,
    incomeRows: rows.filter(row => row.categoryType === 'INCOME'),
    expenseRows: rows.filter(row => row.categoryType === 'EXPENSE')
  };
}

function budgetNormalizeRow_(row) {
  return {
    rowId: String(row.ROW_ID || ''),
    budgetYear: String(row.BUDGET_YEAR || ''),
    categoryType: String(row.CATEGORY_TYPE || ''),
    categoryName: String(row.CATEGORY_NAME || ''),
    rowType: String(row.ROW_TYPE || 'INPUT'),
    termLabel: String(row.TERM_LABEL || ''),
    cohortCode: budgetNumber_(row.COHORT_CODE),
    studyYear: budgetNumber_(row.STUDY_YEAR),
    semester: budgetNumber_(row.SEMESTER),
    studentCount: budgetNumber_(row.STUDENT_COUNT),
    feeRate: budgetNumber_(row.FEE_RATE),
    amount: budgetNumber_(row.AMOUNT),
    inputAmount: row.INPUT_AMOUNT === '' || row.INPUT_AMOUNT === null || row.INPUT_AMOUNT === undefined
      ? budgetNumber_(row.AMOUNT)
      : budgetNumber_(row.INPUT_AMOUNT),
    sortOrder: budgetNumber_(row.SORT_ORDER),
    remark: String(row.REMARK || ''),
    isActive: budgetBool_(row.IS_ACTIVE, true)
  };
}

function budgetCreateYear_(year) {
  const targetYear = String(year || '').trim();
  if (!targetYear) return { ok: false, message: 'Missing year' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = budgetRows_().some(row => String(row.BUDGET_YEAR || '').trim() === targetYear);
    if (existing) return { ok: false, message: 'ปีงบประมาณนี้มีอยู่แล้ว' };

    const now = new Date();
    const values = BUDGET_TEMPLATE_ROWS.map((item, index) => {
      const categoryType = item[0];
      const categoryName = item[1];
      const rowType = item[2];
      return [
        Utilities.getUuid(),
        targetYear,
        categoryType,
        categoryName,
        rowType,
        '',
        '',
        '',
        '',
        0,
        0,
        0,
        0,
        index + 1,
        '',
        true,
        now
      ];
    });
    budgetSheet_().getRange(budgetSheet_().getLastRow() + 1, 1, values.length, BUDGET_HEADERS.length).setValues(values);
    return { ok: true, data: { year: targetYear } };
  } finally {
    lock.releaseLock();
  }
}

function budgetSaveInput_(payload) {
  const rowId = String(payload.rowId || '').trim();
  if (!rowId) return { ok: false, message: 'Missing rowId' };

  const sheet = budgetSheet_();
  const target = budgetRows_().find(row => String(row.ROW_ID || '') === rowId);
  if (!target) return { ok: false, message: 'ไม่พบแถวที่ต้องการบันทึก' };

  const current = budgetNormalizeRow_(target);
  const next = {
    ROW_ID: rowId,
    BUDGET_YEAR: payload.budgetYear || current.budgetYear,
    CATEGORY_TYPE: payload.categoryType || current.categoryType,
    CATEGORY_NAME: payload.categoryName || current.categoryName,
    ROW_TYPE: target.ROW_TYPE || current.rowType || 'INPUT',
    TERM_LABEL: payload.termLabel || '',
    COHORT_CODE: payload.cohortCode || '',
    STUDY_YEAR: payload.studyYear || '',
    SEMESTER: payload.semester || '',
    STUDENT_COUNT: budgetNumber_(payload.studentCount),
    FEE_RATE: budgetNumber_(payload.feeRate),
    AMOUNT: budgetNumber_(payload.amount),
    INPUT_AMOUNT: budgetNumber_(payload.inputAmount || payload.amount),
    SORT_ORDER: payload.sortOrder || current.sortOrder || '',
    REMARK: payload.remark || '',
    IS_ACTIVE: true,
    UPDATED_AT: new Date()
  };

  sheet.getRange(target._rowIndex, 1, 1, BUDGET_HEADERS.length).setValues([BUDGET_HEADERS.map(header => next[header])]);
  return { ok: true, data: { rowId } };
}

function budgetDeleteInput_(rowId) {
  const id = String(rowId || '').trim();
  if (!id) return { ok: false, message: 'Missing rowId' };

  const sheet = budgetSheet_();
  const target = budgetRows_().find(row => String(row.ROW_ID || '') === id);
  if (!target) return { ok: false, message: 'ไม่พบแถวที่ต้องการลบ' };

  const activeColumn = BUDGET_HEADERS.indexOf('IS_ACTIVE') + 1;
  const updatedColumn = BUDGET_HEADERS.indexOf('UPDATED_AT') + 1;
  sheet.getRange(target._rowIndex, activeColumn).setValue(false);
  sheet.getRange(target._rowIndex, updatedColumn).setValue(new Date());
  return { ok: true, data: { rowId: id } };
}

function budgetCloneYear_(sourceYear, targetYear) {
  const source = String(sourceYear || '').trim();
  const target = String(targetYear || '').trim();
  if (!source || !target) return { ok: false, message: 'Missing sourceYear or targetYear' };
  if (source === target) return { ok: false, message: 'ปีต้นทางและปลายทางต้องไม่เหมือนกัน' };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rows = budgetRows_();
    if (rows.some(row => String(row.BUDGET_YEAR || '').trim() === target)) {
      return { ok: false, message: 'ปีปลายทางมีข้อมูลอยู่แล้ว' };
    }
    const sourceRows = rows.filter(row => String(row.BUDGET_YEAR || '').trim() === source);
    if (!sourceRows.length) return { ok: false, message: 'ไม่พบข้อมูลปีต้นทาง' };

    const now = new Date();
    const values = sourceRows.map(row => BUDGET_HEADERS.map(header => {
      if (header === 'ROW_ID') return Utilities.getUuid();
      if (header === 'BUDGET_YEAR') return target;
      if (header === 'UPDATED_AT') return now;
      return row[header];
    }));
    budgetSheet_().getRange(budgetSheet_().getLastRow() + 1, 1, values.length, BUDGET_HEADERS.length).setValues(values);
    return { ok: true, data: { sourceYear: source, targetYear: target } };
  } finally {
    lock.releaseLock();
  }
}

function budgetNumber_(value) {
  const n = Number(value);
  return isFinite(n) ? n : 0;
}

function budgetBool_(value, fallback) {
  if (value === '' || value === null || value === undefined) return fallback;
  if (value === true || value === false) return value;
  return !['FALSE', 'false', '0', 'NO', 'no'].includes(String(value).trim());
}
