'use strict';

function escHtml(str) {
  return String(str ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setBtnLoading(btn, loading, defaultHTML) {
  if (!btn) return;
  if (loading) {
    btn.disabled = true;
    btn.classList.add('is-loading');
    btn.setAttribute('aria-busy', 'true');
    btn.dataset.origHtml = btn.innerHTML;
    btn.innerHTML = '<span class="spinner-sm"></span>&nbsp;กำลังบันทึก...';
  } else {
    btn.disabled = false;
    btn.classList.remove('is-loading');
    btn.removeAttribute('aria-busy');
    btn.innerHTML = btn.dataset.origHtml || defaultHTML || btn.innerHTML;
  }
}

function hydrateLoadingBlocks(root = document) {
  root.querySelectorAll('#loading-block, [data-loading-block]').forEach((block) => {
    if (block.dataset.loadingHydrated === 'true') return;
    const label = block.dataset.loadingLabel || 'กำลังเตรียมข้อมูล...';
    block.dataset.loadingHydrated = 'true';
    block.innerHTML = '<div class="loading-card"><span class="spinner-sm"></span><strong>' + escHtml(label) + '</strong></div>';
  });
}

// ── Budget dashboard module ───────────────────────────────────
const BUDGET_API_URL_KEY = 'BUDGET_DASHBOARD_API_URL';
const BUDGET_DEFAULT_API_URL = 'https://script.google.com/macros/s/AKfycbyKKEYiogeYm7bCYc26KAA_zHAfguQROfl-fUc5DRiawvnaCA4qMji0zuxOTKiiXqU1mQ/exec';
const BUDGET_DEFAULT_YEAR = '2569';
const budgetFmt = new Intl.NumberFormat('th-TH');
const BUDGET_CATEGORY_COLOR_MAP = {
  'ก.แผนงานจัดการศึกษา (รวม)': '#4f46e5',
  'ก.แผนงานจัดการศึกษา (งบบุคลากร)': '#14b8a6',
  'ก.แผนงานจัดการศึกษา (การจัดการเรียนการสอน)': '#f59e0b',
  'ข.แผนงานบริการวิชาการแก่สังคม': '#ec4899',
  'ค.แผนงานอนุรักษ์ ส่งเสริมและพัฒนาศาสนา ศิลปะและวัฒนธรรม': '#0ea5e9',
  'ง.แผนงานสนับสนุนวิชาการ': '#84cc16',
  'จ.แผนงานวิจัย': '#f97316'
};
const BUDGET_CATEGORY_FALLBACK_COLORS = ['#4f46e5', '#14b8a6', '#f59e0b', '#ec4899', '#0ea5e9', '#84cc16', '#f97316', '#ef4444'];
const BUDGET_ACTION_FALLBACKS = {
  getBudgetDashboard: ['getDashboard'],
  saveBudgetInput: ['saveRow'],
  deleteBudgetInput: ['deleteRow']
};

const budgetState = {
  budgetYears: [],
  budgetYear: '',
  dashboard: null,
  forecastDashboards: {},
  dashboardCache: {},
  forecastRequestId: 0,
  renderToken: 0,
  computedCache: {
    sourceDashboard: null,
    sourceForecastDashboards: null,
    interactionKey: '',
    value: null
  },
  interaction: {
    metricKey: '',
    expenseCategory: '',
    forecastYear: ''
  }
};

let budgetBarChart = null;
let budgetDoughnutSummaryChart = null;
let budgetDoughnutDetailChart = null;
let budgetForecastChart = null;
let budgetChartRenderFrame = 0;

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => {
    if (document.body?.dataset.page === 'budget') initBudgetPage();
  }, { once: true });
} else if (document.body?.dataset.page === 'budget') {
  initBudgetPage();
}

async function initBudgetPage() {
  try {
    hydrateLoadingBlocks();
    budgetBindEvents();
    budgetResetForm();
    budgetShowPage();
    await budgetLoadYears();
  } catch (err) {
    budgetSetStatus('เริ่มระบบไม่สำเร็จ: ' + budgetGetErrorMessage(err), 'error');
    budgetShowPage();
  }
}

function budgetBindEvents() {
  document.getElementById('budgetYearSelect')?.addEventListener('change', async event => {
    budgetState.budgetYear = event.target.value;
    await budgetLoadDashboard(budgetState.budgetYear);
  });

  document.getElementById('btnReload')?.addEventListener('click', async () => {
    if (!budgetState.budgetYear) return;
    await budgetLoadDashboard(budgetState.budgetYear, { forceRefresh: true });
  });

  document.getElementById('btnOpenAddModal')?.addEventListener('click', () => budgetOpenCreateYearModal());
  document.getElementById('btnOpenGuide')?.addEventListener('click', () => budgetOpenModal('budgetGuideModal'));
  document.getElementById('btnOpenClone')?.addEventListener('click', () => budgetOpenCloneModal());
  document.getElementById('btnSaveRow')?.addEventListener('click', async () => budgetSaveCurrentForm());
  document.getElementById('btnResetForm')?.addEventListener('click', () => budgetResetForm());
  document.getElementById('btnCloneCancel')?.addEventListener('click', () => budgetCloseModal('budgetCloneModal'));
  document.getElementById('btnCloneConfirm')?.addEventListener('click', async () => budgetCloneBudgetYear());
  document.getElementById('btnCreateYearCancel')?.addEventListener('click', () => budgetCloseModal('budgetCreateYearModal'));
  document.getElementById('btnCreateYearConfirm')?.addEventListener('click', async () => budgetCreateBudgetYearAction());
  document.getElementById('btnBudgetClearFocus')?.addEventListener('click', () => budgetClearInteractionFocus());
  document.getElementById('btnBudgetOpenForecastYear')?.addEventListener('click', async () => budgetOpenSelectedForecastYear());
  document.getElementById('categoryName')?.addEventListener('change', budgetSyncCategoryFields);
  document.getElementById('categoryType')?.addEventListener('change', budgetSyncCategoryFields);
  document.getElementById('studentCount')?.addEventListener('input', budgetRecalcFormAmount);
  document.getElementById('feeRate')?.addEventListener('input', budgetRecalcFormAmount);
  document.getElementById('semester')?.addEventListener('input', budgetRecalcTermLabel);
  document.getElementById('studyYear')?.addEventListener('input', budgetRecalcTermLabel);
  document.getElementById('budgetYearInput')?.addEventListener('input', budgetRecalcTermLabel);
  document.getElementById('amount')?.addEventListener('input', budgetRecalcFormAmount);

  document.querySelectorAll('[data-budget-close]').forEach(el => {
    el.addEventListener('click', () => budgetCloseModal(el.getAttribute('data-budget-close')));
  });

  document.addEventListener('click', async event => {
    if (document.body?.dataset.page !== 'budget') return;
    const editButton = event.target.closest('[data-budget-edit-row]');
    if (editButton) {
      budgetEditRow(editButton.getAttribute('data-budget-edit-row'));
      return;
    }

    const deleteButton = event.target.closest('[data-budget-delete-row]');
    if (deleteButton) {
      await budgetDeleteRowAction(deleteButton.getAttribute('data-budget-delete-row'));
    }
  });
}

function budgetShowPage() {
  const loadingBlock = document.getElementById('loading-block');
  const pageContent = document.getElementById('page-content');
  if (loadingBlock) loadingBlock.style.display = 'none';
  if (pageContent) pageContent.style.display = 'block';
}

async function budgetLoadYears() {
  budgetSetStatus('กำลังโหลดปีงบประมาณ...', 'info');
  const res = await budgetApiGetCompat('getBudgetYears');
  if (!res.ok && budgetIsUnknownActionMessage(res.message)) {
    budgetState.budgetYears = budgetInferFallbackYears();
    budgetRenderYearSelect();
    budgetState.budgetYear = budgetState.budgetYears[budgetState.budgetYears.length - 1] || '';
    budgetSetValue('budgetYearSelect', budgetState.budgetYear);
    if (budgetState.budgetYear) {
      await budgetLoadDashboard(budgetState.budgetYear);
      budgetSetStatus('โหลดข้อมูลด้วย route เดิมของ budget สำเร็จ', 'success');
      return;
    }
  }
  if (!res.ok) throw new Error(res.message || 'โหลดปีงบประมาณไม่สำเร็จ');

  budgetState.budgetYears = Array.isArray(res.data) ? res.data : [];
  budgetRenderYearSelect();

  if (!budgetState.budgetYears.length) {
    budgetState.budgetYear = '';
    budgetState.dashboard = null;
    budgetRenderDashboard();
    budgetSetStatus('ยังไม่มีข้อมูลปีงบประมาณในชีต กรุณากด "สร้างปีงบ"', 'info');
    return;
  }

  budgetState.budgetYear = budgetState.budgetYears.includes(BUDGET_DEFAULT_YEAR)
    ? BUDGET_DEFAULT_YEAR
    : budgetState.budgetYears[0];
  budgetSetValue('budgetYearSelect', budgetState.budgetYear);
  await budgetLoadDashboard(budgetState.budgetYear);
}

function budgetRenderYearSelect() {
  const select = document.getElementById('budgetYearSelect');
  const cloneSource = document.getElementById('cloneSourceYear');
  const options = budgetState.budgetYears.map(year => `<option value="${escHtml(year)}">${escHtml(year)}</option>`).join('');
  if (select) select.innerHTML = options;
  if (cloneSource) cloneSource.innerHTML = options;
}

async function budgetLoadDashboard(year, options = {}) {
  const forceRefresh = options.forceRefresh === true;
  if (!year) {
    budgetState.dashboard = null;
    budgetState.forecastDashboards = {};
    budgetRenderDashboard();
    return;
  }

  if (!forceRefresh && budgetState.dashboardCache[year]) {
    budgetState.dashboard = budgetState.dashboardCache[year];
    budgetRenderDashboard();
    budgetSetStatus('โหลดข้อมูลเรียบร้อย', 'success');
    budgetScheduleForecastLoad(year);
    return;
  }

  budgetSetStatus(`กำลังโหลดข้อมูลปี ${year} ...`, 'info');
  const res = await budgetApiGetCompat('getBudgetDashboard', { budgetYear: year });
  if (!res.ok) throw new Error(res.message || 'โหลด Dashboard ไม่สำเร็จ');
  budgetState.dashboard = res.data || null;
  budgetState.dashboardCache[year] = budgetState.dashboard;
  budgetRenderDashboard();
  budgetSetStatus('โหลดข้อมูลเรียบร้อย', 'success');
  budgetScheduleForecastLoad(year);
}

function budgetRenderDashboard() {
  const computed = budgetGetComputedDashboard();
  const renderToken = ++budgetState.renderToken;
  budgetRenderStats(computed);
  budgetRenderInsight(computed);
  budgetRenderTables(computed);
  budgetScheduleChartRender(computed, renderToken);
}

function budgetRenderStats(computed) {
  const stats = (computed || budgetGetComputedDashboard()).stats;
  budgetSetText('totalIncome', budgetFmtNum(stats.totalIncome));
  budgetSetText('totalExpense', budgetFmtNum(stats.totalExpense));
  budgetSetText('netBalance', budgetFmtNum(stats.netBalance));
  budgetSetText('budgetYearStat', budgetState.budgetYear || '-');

  const netEl = document.getElementById('netBalance');
  const hintEl = document.getElementById('netHint');
  if (!netEl || !hintEl) return;

  netEl.className = 'budget-stat-value';
  if (stats.netBalance > 0) {
    netEl.classList.add('positive');
    hintEl.textContent = 'รายรับมากกว่ารายจ่าย';
  } else if (stats.netBalance < 0) {
    netEl.classList.add('negative');
    hintEl.textContent = 'รายจ่ายสูงกว่ารายรับ';
  } else {
    netEl.classList.add('warning');
    hintEl.textContent = 'สมดุลพอดี';
  }
}

function budgetRenderTables(computed) {
  const incomeBody = document.getElementById('incomeTableBody');
  const expenseBody = document.getElementById('expenseTableBody');
  const incomeHint = document.getElementById('incomeTableHint');
  const expenseHint = document.getElementById('expenseTableHint');
  const dashboardComputed = computed || budgetGetComputedDashboard();
  const interaction = budgetState.interaction || {};
  const incomeRows = dashboardComputed.incomeRows;
  const expenseRows = dashboardComputed.expenseRows;
  let visibleIncomeRows = incomeRows;
  let visibleExpenseRows = expenseRows;

  if (interaction.metricKey === 'income') {
    visibleExpenseRows = [];
  } else if (interaction.metricKey === 'expense') {
    visibleIncomeRows = [];
  }

  if (interaction.expenseCategory) {
    visibleExpenseRows = expenseRows.filter(row =>
      row.rowType === 'SUMMARY' || row.categoryName === interaction.expenseCategory
    );
  }

  if (incomeHint) {
    incomeHint.textContent = interaction.metricKey === 'expense'
      ? 'ซ่อนชั่วคราว เพราะกำลังโฟกัสรายจ่ายจากกราฟ'
      : interaction.metricKey === 'income'
        ? 'กำลังโฟกัสรายการรายรับจากกราฟ'
        : 'รายรับจากค่าเทอมและรายได้อื่น';
  }

  if (expenseHint) {
    expenseHint.textContent = interaction.expenseCategory
      ? `กำลังกรองเฉพาะหมวด "${interaction.expenseCategory}"`
      : interaction.metricKey === 'income'
        ? 'ซ่อนชั่วคราว เพราะกำลังโฟกัสรายรับจากกราฟ'
        : 'รายการงบรายจ่ายที่ใช้งานอยู่';
  }

  if (incomeBody) {
    incomeBody.innerHTML = visibleIncomeRows.length
      ? visibleIncomeRows.map(row => `
          <tr>
            <td>${budgetActionButtons(row)}</td>
            <td>${escHtml(row.categoryName)}</td>
            <td>${escHtml(row.termLabel || '—')}</td>
            <td>${escHtml(row.cohortCode || '—')}</td>
            <td>${escHtml(row.studyYear || '—')}</td>
            <td>${escHtml(row.semester || '—')}</td>
            <td>${budgetFmtNum(row.feeRate || 0)}</td>
            <td>${budgetFmtNum(row.studentCount || 0)}</td>
            <td>${budgetFmtNum(row.displayAmount || 0)}</td>
            <td>${escHtml(row.remark || '—')}</td>
          </tr>
        `).join('')
      : '<tr class="budget-empty-row"><td colspan="10">ไม่มีรายการรายรับในมุมมองที่เลือก</td></tr>';
  }

  if (expenseBody) {
    expenseBody.innerHTML = visibleExpenseRows.length
      ? visibleExpenseRows.map(row => `
          <tr>
            <td>${budgetActionButtons(row)}</td>
            <td>${escHtml(row.categoryName)}</td>
            <td>${budgetFmtNum(row.displayAmount || 0)}</td>
            <td>${escHtml(row.sortOrder || '—')}</td>
            <td>${escHtml(row.remark || '—')}</td>
          </tr>
        `).join('')
      : '<tr class="budget-empty-row"><td colspan="5">ไม่มีรายการรายจ่ายในมุมมองที่เลือก</td></tr>';
  }
}

function budgetRenderInsight(computed) {
  const pillEl = document.getElementById('budgetFocusPill');
  const valueEl = document.getElementById('budgetInsightValue');
  const textEl = document.getElementById('budgetInsightText');
  const openYearBtn = document.getElementById('btnBudgetOpenForecastYear');
  const interaction = budgetState.interaction || {};
  const dashboardComputed = computed || budgetGetComputedDashboard();
  const stats = dashboardComputed.stats;
  const forecastRows = dashboardComputed.charts?.forecastRows || [];
  let pill = 'ภาพรวมทั้งปี';
  let value = 'ยังไม่ได้เลือกจุดวิเคราะห์';
  let text = 'คลิกกราฟแท่งเพื่อโฟกัสรายรับ/รายจ่าย คลิกโดนัทเพื่อกรองหมวดงบ และคลิกกราฟแนวโน้มเพื่อเทียบปีล่วงหน้า';

  if (interaction.metricKey === 'income') {
    pill = 'Focus: รายรับ';
    value = `รายรับรวม ${budgetFmtNum(stats.totalIncome)} บาท`;
    text = 'ใช้ตารางรายรับด้านล่างเพื่อตรวจสอบว่ารายได้หลักมาจากค่าเทอมหรือรายได้เสริม และดูว่าควรเพิ่มจำนวนผู้เรียนหรือกระจายแหล่งรายได้';
  } else if (interaction.metricKey === 'expense') {
    pill = 'Focus: รายจ่าย';
    value = `รายจ่ายรวม ${budgetFmtNum(stats.totalExpense)} บาท`;
    text = 'ตารางรายจ่ายถูกเน้นไว้แล้ว เพื่อดูว่าหมวดใดกดต้นทุนรวมมากที่สุดและควรควบคุมก่อน';
  } else if (interaction.metricKey === 'net') {
    const margin = stats.totalIncome > 0 ? (stats.netBalance / stats.totalIncome) * 100 : 0;
    pill = 'Focus: ดุลสุทธิ';
    value = `ดุลสุทธิ ${budgetFmtNum(stats.netBalance)} บาท`;
    text = stats.netBalance >= 0
      ? `งบยังเป็นบวกประมาณ ${budgetFmtPct(margin)} ของรายรับรวม สามารถใช้เป็น buffer สำหรับตัดสินใจลงทุนหรือกันความเสี่ยง`
      : `งบติดลบประมาณ ${budgetFmtPct(Math.abs(margin))} ของรายรับรวม ควรพิจารณาลดรายจ่ายหมวดใหญ่หรือเพิ่มรายรับที่ทำได้เร็ว`;
  }

  if (interaction.expenseCategory) {
    const expenseRows = dashboardComputed.charts?.expenseDetailRows || [];
    const totalExpense = expenseRows.reduce((sum, row) => sum + Number(row.amount || 0), 0);
    const selected = expenseRows.find(row => row.label === interaction.expenseCategory);
    const amount = Number(selected?.amount || 0);
    const percent = totalExpense > 0 ? (amount / totalExpense) * 100 : 0;
    pill = 'Focus: หมวดรายจ่าย';
    value = `${interaction.expenseCategory} ${budgetFmtNum(amount)} บาท`;
    text = `หมวดนี้คิดเป็น ${budgetFmtPct(percent)} ของรายจ่ายทั้งหมด ตารางรายจ่ายด้านล่างถูกกรองแล้วเพื่อช่วยดูรายละเอียดเฉพาะหมวด`;
  }

  if (interaction.forecastYear) {
    const selectedForecast = forecastRows.find(row => String(row.year) === String(interaction.forecastYear));
    const baseForecast = forecastRows[0] || { income: 0, expense: 0, net: 0 };
    const delta = Number(selectedForecast?.net || 0) - Number(baseForecast.net || 0);
    pill = `Forecast: ปี ${interaction.forecastYear}`;
    value = `ดุลคาดการณ์ ${budgetFmtNum(selectedForecast?.net || 0)} บาท`;
    text = `เทียบกับปี ${budgetState.budgetYear || '-'} ต่างกัน ${budgetFmtNum(delta)} บาท และมีรายรับ ${budgetFmtNum(selectedForecast?.income || 0)} บาท / รายจ่าย ${budgetFmtNum(selectedForecast?.expense || 0)} บาท`;
  }

  if (pillEl) pillEl.textContent = pill;
  if (valueEl) valueEl.textContent = value;
  if (textEl) textEl.textContent = text;
  if (openYearBtn) {
    const canOpen = Boolean(interaction.forecastYear && (budgetState.budgetYears || []).includes(String(interaction.forecastYear)));
    openYearBtn.disabled = !canOpen;
  }
}

function budgetActionButtons(row) {
  if (!row || row.rowType === 'SUMMARY') return '<span class="text-muted text-sm">ระบบ</span>';
  return `
      <div class="budget-action-buttons">
        <button type="button" class="btn btn-edit" data-budget-edit-row="${budgetEscapeAttr(row.rowId)}">แก้ไข</button>
        <button type="button" class="btn btn-delete" data-budget-delete-row="${budgetEscapeAttr(row.rowId)}">ลบ</button>
      </div>
  `;
}

function budgetOpenCreateYearModal() {
  budgetSetValue('createBudgetYearInput', budgetState.budgetYear || '');
  budgetOpenModal('budgetCreateYearModal');
}

async function budgetCreateBudgetYearAction() {
  const targetYear = budgetGetValue('createBudgetYearInput');
  if (!targetYear) {
    budgetSetStatus('กรุณาระบุปีงบประมาณที่จะสร้าง', 'error');
    return;
  }

  const button = document.getElementById('btnCreateYearConfirm');
  setBtnLoading(button, true, 'สร้างปีงบ');
  budgetSetStatus(`กำลังสร้างปีงบ ${targetYear} ...`, 'info');

  try {
    const res = await budgetApiPostCompat('createBudgetYear', { year: targetYear });
    if (!res.ok) throw new Error(res.message || 'สร้างปีงบไม่สำเร็จ');
    budgetCloseModal('budgetCreateYearModal');
    await budgetRefreshYearsThenDashboard(targetYear);
    budgetSetStatus(`สร้างปีงบ ${targetYear} เรียบร้อย`, 'success');
  } catch (err) {
    budgetSetStatus(budgetGetErrorMessage(err), 'error');
  } finally {
    setBtnLoading(button, false, 'สร้างปีงบ');
  }
}

function budgetEditRow(rowId) {
  const computed = budgetGetComputedDashboard();
  const allRows = [...computed.incomeRows, ...computed.expenseRows];
  const row = allRows.find(item => item.rowId === rowId);
  if (!row || row.rowType === 'SUMMARY') return;

  budgetSetText('rowModalTitle', 'แก้ไขรายการ');
  budgetSetValue('rowId', row.rowId || '');
  budgetSetValue('budgetYearInput', row.budgetYear || '');
  budgetSetValue('categoryType', row.categoryType || 'INCOME');
  budgetSetValue('categoryName', row.categoryName || '');
  budgetSetValue('termLabel', row.termLabel || '');
  budgetSetValue('cohortCode', row.cohortCode || '');
  budgetSetValue('studyYear', row.studyYear || '');
  budgetSetValue('semester', row.semester || '');
  budgetSetValue('studentCount', row.studentCount || 0);
  budgetSetValue('feeRate', row.feeRate || 0);
  budgetSetValue('amount', row.displayAmount || row.inputAmount || row.amount || 0);
  budgetSetValue('sortOrder', row.sortOrder || '');
  budgetSetValue('remark', row.remark || '');
  budgetSyncCategoryFields();
  budgetOpenModal('budgetRowModal');
}

async function budgetSaveCurrentForm() {
  const payload = budgetGetFormData();
  const validationMessage = budgetValidateForm(payload);
  if (validationMessage) {
    budgetSetStatus(validationMessage, 'error');
    return;
  }

  budgetSetStatus('กำลังบันทึกข้อมูล...', 'info');
  const saveButton = document.getElementById('btnSaveRow');
  setBtnLoading(saveButton, true, 'บันทึกแถวนี้');

  try {
    const res = await budgetApiPostCompat('saveBudgetInput', payload);
    if (!res.ok) throw new Error(res.message || 'บันทึกข้อมูลไม่สำเร็จ');

    budgetCloseModal('budgetRowModal');
    budgetResetForm();
    budgetState.budgetYear = payload.budgetYear;
    await budgetRefreshYearsThenDashboard(payload.budgetYear);
    budgetSetStatus('บันทึกข้อมูลเรียบร้อย', 'success');
  } catch (err) {
    budgetSetStatus(budgetGetErrorMessage(err), 'error');
  } finally {
    setBtnLoading(saveButton, false, 'บันทึกแถวนี้');
  }
}

async function budgetDeleteRowAction(rowId) {
  if (!window.confirm('ยืนยันการลบรายการนี้หรือไม่? ระบบจะทำ soft delete เท่านั้น')) return;
  budgetSetStatus('กำลังลบข้อมูล...', 'info');
  try {
    const res = await budgetApiPostCompat('deleteBudgetInput', { rowId });
    if (!res.ok) throw new Error(res.message || 'ลบข้อมูลไม่สำเร็จ');
    delete budgetState.dashboardCache[budgetState.budgetYear];
    await budgetLoadDashboard(budgetState.budgetYear, { forceRefresh: true });
    budgetSetStatus('ลบข้อมูลเรียบร้อย', 'success');
  } catch (err) {
    budgetSetStatus(budgetGetErrorMessage(err), 'error');
  }
}

function budgetGetFormData() {
  const amountValue = budgetGetValue('amount');
  return {
    rowId: budgetGetValue('rowId'),
    budgetYear: budgetGetValue('budgetYearInput'),
    categoryType: budgetGetValue('categoryType'),
    categoryName: budgetGetValue('categoryName'),
    termLabel: budgetGetValue('termLabel'),
    cohortCode: budgetGetValue('cohortCode'),
    studyYear: budgetGetValue('studyYear'),
    semester: budgetGetValue('semester'),
    studentCount: budgetGetValue('studentCount'),
    feeRate: budgetGetValue('feeRate'),
    amount: amountValue,
    inputAmount: amountValue,
    sortOrder: budgetGetValue('sortOrder'),
    remark: budgetGetValue('remark')
  };
}

function budgetValidateForm(payload) {
  if (!payload.rowId) return 'กรุณาเลือกแถวข้อมูลจากตารางก่อนแก้ไข';
  if (!payload.budgetYear) return 'กรุณาระบุปีงบประมาณ';
  if (!payload.categoryType) return 'กรุณาเลือกประเภท';
  if (!payload.categoryName) return 'กรุณาเลือกรายการ';
  return '';
}

function budgetResetForm() {
  document.getElementById('rowForm')?.reset();
  budgetSetText('rowModalTitle', 'แก้ไขรายการ');
  budgetSetValue('rowId', '');
  budgetSetValue('budgetYearInput', budgetState.budgetYear || '');
  budgetSetValue('termLabel', '');
  budgetSetValue('categoryType', 'INCOME');
  budgetSetValue('studentCount', 0);
  budgetSetValue('feeRate', 0);
  budgetSetValue('amount', 0);
  budgetSetValue('sortOrder', '');
  budgetSetValue('remark', '');
  budgetSyncCategoryFields();
}

function budgetSyncCategoryFields() {
  const categoryName = budgetGetValue('categoryName');
  const isIncome = [
    'ค่าธรรมเนียมการศึกษา',
    'รายได้จากการบริการวิชาการ',
    'รายได้อื่น ๆ (PN)',
    'ค่ารับสมัคร'
  ].includes(categoryName);
  const categoryTypeEl = document.getElementById('categoryType');
  const feeRateEl = document.getElementById('feeRate');
  const studentCountEl = document.getElementById('studentCount');
  const amountEl = document.getElementById('amount');
  const isTuition = categoryName === 'ค่าธรรมเนียมการศึกษา';

  if (categoryTypeEl && categoryName) categoryTypeEl.value = isIncome ? 'INCOME' : 'EXPENSE';
  if (feeRateEl) feeRateEl.disabled = !isTuition;
  if (studentCountEl) studentCountEl.disabled = !isTuition;
  if (amountEl) amountEl.disabled = isTuition;
  budgetRecalcTermLabel();
  budgetRecalcFormAmount();
}

function budgetOpenCloneModal() {
  if (!budgetState.budgetYears.length) {
    budgetSetStatus('ยังไม่มีปีงบสำหรับ Clone', 'error');
    return;
  }
  budgetSetValue('cloneSourceYear', budgetState.budgetYear || budgetState.budgetYears[budgetState.budgetYears.length - 1]);
  budgetSetValue('cloneTargetYear', '');
  budgetOpenModal('budgetCloneModal');
}

async function budgetCloneBudgetYear() {
  const sourceYear = budgetGetValue('cloneSourceYear');
  const targetYear = budgetGetValue('cloneTargetYear');
  if (!sourceYear || !targetYear) {
    budgetSetStatus('กรุณาเลือกปีต้นทางและกรอกปีปลายทาง', 'error');
    return;
  }
  if (sourceYear === targetYear) {
    budgetSetStatus('ปีต้นทางและปลายทางต้องไม่เหมือนกัน', 'error');
    return;
  }

  const cloneButton = document.getElementById('btnCloneConfirm');
  setBtnLoading(cloneButton, true, 'ยืนยัน Clone');
  budgetSetStatus(`กำลัง Clone ปี ${sourceYear} ไป ${targetYear} ...`, 'info');

  try {
    const res = await budgetApiPostCompat('cloneBudgetYear', { sourceYear, targetYear });
    if (!res.ok) throw new Error(res.message || 'Clone ไม่สำเร็จ');
    budgetCloseModal('budgetCloneModal');
    await budgetRefreshYearsThenDashboard(targetYear);
    budgetSetStatus(`Clone ปี ${sourceYear} ไป ${targetYear} เรียบร้อย`, 'success');
  } catch (err) {
    budgetSetStatus(budgetGetErrorMessage(err), 'error');
  } finally {
    setBtnLoading(cloneButton, false, 'ยืนยัน Clone');
  }
}

async function budgetRefreshYearsThenDashboard(targetYear) {
  const res = await budgetApiGetCompat('getBudgetYears');
  if (res.ok) {
    budgetState.budgetYears = Array.isArray(res.data) ? res.data : [];
    budgetRenderYearSelect();
  } else if (budgetIsUnknownActionMessage(res.message)) {
    budgetState.budgetYears = budgetInferFallbackYears(targetYear);
    budgetRenderYearSelect();
  }
  budgetState.budgetYear = targetYear;
  budgetSetValue('budgetYearSelect', targetYear);
  delete budgetState.dashboardCache[targetYear];
  await budgetLoadDashboard(targetYear, { forceRefresh: true });
}

function getBudgetCategoryColor(label) {
  if (BUDGET_CATEGORY_COLOR_MAP[label]) return BUDGET_CATEGORY_COLOR_MAP[label];
  const hash = Array.from(String(label || '')).reduce((sum, char) => sum + char.charCodeAt(0), 0);
  return BUDGET_CATEGORY_FALLBACK_COLORS[hash % BUDGET_CATEGORY_FALLBACK_COLORS.length];
}

function getBudgetCategoryColors(labels) {
  return (Array.isArray(labels) ? labels : []).map(getBudgetCategoryColor);
}

function budgetInitCharts() {
  const barCtx = document.getElementById('barChart');
  const doughnutSummaryCtx = document.getElementById('doughnutChartSummary');
  const doughnutDetailCtx = document.getElementById('doughnutChartDetail');
  const forecastCtx = document.getElementById('forecastChart');
  const doughnutLegendOptions = {
    position: 'right',
    align: 'center',
    onClick(_event, legendItem, legend) {
      const chart = legend?.chart;
      const label = legendItem?.text ? String(legendItem.text).replace(/\s+\(.+\)$/, '') : '';
      if (!chart || !label) return;
      budgetToggleExpenseCategoryFocus(label);
    },
    labels: {
      usePointStyle: true,
      pointStyle: 'circle',
      boxWidth: 10,
      boxHeight: 10,
      padding: 14,
      generateLabels(chart) {
        const data = chart?.data || {};
        const labels = Array.isArray(data.labels) ? data.labels : [];
        const values = Array.isArray(data.datasets?.[0]?.data) ? data.datasets[0].data : [];
        const colors = Array.isArray(data.datasets?.[0]?.backgroundColor) ? data.datasets[0].backgroundColor : [];
        const total = values.reduce((sum, value) => sum + Number(value || 0), 0);
        return labels.map((label, index) => {
          const amount = Number(values[index] || 0);
          const percent = total > 0 ? (amount / total) * 100 : 0;
          return {
            text: `${label} (${budgetFmtPct(percent)})`,
            fillStyle: colors[index],
            strokeStyle: colors[index],
            lineWidth: 0,
            hidden: false,
            index,
            pointStyle: 'circle'
          };
        });
      },
      font: {
        size: 11
      }
    }
  };

  if (barCtx && !budgetBarChart) {
    budgetBarChart = new Chart(barCtx, {
      type: 'bar',
      data: {
        labels: ['รายรับรวม', 'รายจ่ายรวม', 'ดุลสุทธิ'],
        datasets: [{
          data: [0, 0, 0],
          backgroundColor: ['#4f46e5', '#f59e0b', '#ef4444'],
          borderRadius: 12,
          maxBarThickness: 64
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        onClick(_event, elements) {
          const index = elements?.[0]?.index;
          if (index === undefined) return;
          const metricMap = ['income', 'expense', 'net'];
          const metricKey = metricMap[index] || '';
          budgetState.interaction.metricKey = budgetState.interaction.metricKey === metricKey ? '' : metricKey;
          budgetRenderDashboard();
        },
        plugins: { legend: { display: false } },
        scales: {
          y: {
            ticks: {
              callback(value) {
                return budgetFmt.format(value);
              }
            }
          }
        }
      }
    });
  }

  if (doughnutSummaryCtx && !budgetDoughnutSummaryChart) {
    budgetDoughnutSummaryChart = new Chart(doughnutSummaryCtx, {
      type: 'doughnut',
      data: {
        labels: [],
        datasets: [{ data: [], backgroundColor: [], borderWidth: 0 }]
      },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          onClick(_event, elements) {
            const index = elements?.[0]?.index;
            const label = index === undefined ? '' : String(budgetDoughnutSummaryChart?.data?.labels?.[index] || '');
            budgetToggleExpenseCategoryFocus(label);
          },
          plugins: {
            legend: doughnutLegendOptions,
            tooltip: {
              callbacks: {
                label(context) {
                  const value = Number(context.raw || 0);
                  const total = (context.dataset?.data || []).reduce((sum, item) => sum + Number(item || 0), 0);
                  const percent = total > 0 ? (value / total) * 100 : 0;
                  return `${context.label}: ${budgetFmtNum(value)} บาท (${budgetFmtPct(percent)})`;
                }
              }
            }
          },
          cutout: '62%'
        }
      });
    }

  if (doughnutDetailCtx && !budgetDoughnutDetailChart) {
    budgetDoughnutDetailChart = new Chart(doughnutDetailCtx, {
      type: 'doughnut',
      data: {
        labels: [],
        datasets: [{ data: [], backgroundColor: [], borderWidth: 0 }]
      },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          onClick(_event, elements) {
            const index = elements?.[0]?.index;
            const label = index === undefined ? '' : String(budgetDoughnutDetailChart?.data?.labels?.[index] || '');
            budgetToggleExpenseCategoryFocus(label);
          },
          plugins: {
            legend: doughnutLegendOptions,
            tooltip: {
              callbacks: {
                label(context) {
                  const value = Number(context.raw || 0);
                  const total = (context.dataset?.data || []).reduce((sum, item) => sum + Number(item || 0), 0);
                  const percent = total > 0 ? (value / total) * 100 : 0;
                  return `${context.label}: ${budgetFmtNum(value)} บาท (${budgetFmtPct(percent)})`;
                }
              }
            }
          },
          cutout: '62%'
        }
      });
    }

  if (forecastCtx && !budgetForecastChart) {
    budgetForecastChart = new Chart(forecastCtx, {
      type: 'line',
      data: {
        labels: [],
        datasets: [
          { label: 'รายรับ', data: [], borderColor: '#4f46e5', backgroundColor: 'rgba(79, 70, 229, 0.15)', tension: 0.28, fill: false },
          { label: 'รายจ่าย', data: [], borderColor: '#f59e0b', backgroundColor: 'rgba(245, 158, 11, 0.15)', tension: 0.28, fill: false },
          { label: 'ดุลสุทธิ', data: [], borderColor: '#16a34a', backgroundColor: 'rgba(22, 163, 74, 0.15)', tension: 0.28, fill: false }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        interaction: {
          mode: 'index',
          intersect: false
        },
        onClick(_event, elements) {
          const index = elements?.[0]?.index;
          const year = index === undefined ? '' : String(budgetForecastChart?.data?.labels?.[index] || '');
          if (!year) return;
          budgetState.interaction.forecastYear = budgetState.interaction.forecastYear === year ? '' : year;
          budgetRenderDashboard();
        },
        plugins: {
          legend: {
            position: 'bottom',
            onClick(_event, legendItem) {
              const metricKey = budgetMapForecastLegendToMetric_(legendItem?.text);
              if (!metricKey) return;
              budgetState.interaction.metricKey = budgetState.interaction.metricKey === metricKey ? '' : metricKey;
              budgetRenderDashboard();
            }
          }
        },
        scales: {
          y: {
            ticks: {
              callback(value) {
                return budgetFmt.format(value);
              }
            }
          }
        }
      }
    });
  }
}

function budgetUpdateCharts() {
  const computed = budgetGetComputedDashboard();
  const stats = computed.stats;
  const chartData = computed.charts;

  if (budgetBarChart) {
    budgetBarChart.data.datasets[0].data = [stats.totalIncome, stats.totalExpense, stats.netBalance];
    budgetBarChart.data.datasets[0].backgroundColor = ['#4f46e5', '#f59e0b', stats.netBalance >= 0 ? '#16a34a' : '#ef4444'];
    budgetBarChart.setActiveElements(budgetGetBarActiveElements_());
    budgetBarChart.update();
  }

  if (budgetDoughnutSummaryChart) {
    const labels = chartData.expenseSummaryRows.map(row => row.label);
    budgetDoughnutSummaryChart.data.labels = labels;
    budgetDoughnutSummaryChart.data.datasets[0].data = chartData.expenseSummaryRows.map(row => row.amount);
    budgetDoughnutSummaryChart.data.datasets[0].backgroundColor = getBudgetCategoryColors(labels);
    budgetDoughnutSummaryChart.setActiveElements(budgetGetDoughnutActiveElements_(budgetDoughnutSummaryChart));
    budgetDoughnutSummaryChart.update();
  }

  if (budgetDoughnutDetailChart) {
    const labels = chartData.expenseDetailRows.map(row => row.label);
    budgetDoughnutDetailChart.data.labels = labels;
    budgetDoughnutDetailChart.data.datasets[0].data = chartData.expenseDetailRows.map(row => row.amount);
    budgetDoughnutDetailChart.data.datasets[0].backgroundColor = getBudgetCategoryColors(labels);
    budgetDoughnutDetailChart.setActiveElements(budgetGetDoughnutActiveElements_(budgetDoughnutDetailChart));
    budgetDoughnutDetailChart.update();
  }

  if (budgetForecastChart) {
    budgetForecastChart.data.labels = chartData.forecastRows.map(row => row.year);
    budgetForecastChart.data.datasets[0].data = chartData.forecastRows.map(row => row.income);
    budgetForecastChart.data.datasets[1].data = chartData.forecastRows.map(row => row.expense);
    budgetForecastChart.data.datasets[2].data = chartData.forecastRows.map(row => row.net);
    budgetForecastChart.setActiveElements(budgetGetForecastActiveElements_());
    budgetForecastChart.update();
  }
}

function budgetGetBarActiveElements_() {
  const metricMap = { income: 0, expense: 1, net: 2 };
  const metricIndex = metricMap[budgetState.interaction?.metricKey];
  return metricIndex === undefined ? [] : [{ datasetIndex: 0, index: metricIndex }];
}

function budgetGetDoughnutActiveElements_(chart) {
  const label = budgetState.interaction?.expenseCategory;
  if (!label) return [];
  const labels = Array.isArray(chart?.data?.labels) ? chart.data.labels : [];
  const index = labels.findIndex(item => String(item) === String(label));
  return index < 0 ? [] : [{ datasetIndex: 0, index }];
}

function budgetGetForecastActiveElements_() {
  const year = budgetState.interaction?.forecastYear;
  if (!year || !budgetForecastChart) return [];
  const labels = Array.isArray(budgetForecastChart.data?.labels) ? budgetForecastChart.data.labels : [];
  const index = labels.findIndex(item => String(item) === String(year));
  if (index < 0) return [];
  return [0, 1, 2].map(datasetIndex => ({ datasetIndex, index }));
}

function budgetToggleExpenseCategoryFocus(label) {
  if (!label) return;
  budgetState.interaction.metricKey = 'expense';
  budgetState.interaction.expenseCategory = budgetState.interaction.expenseCategory === label ? '' : label;
  budgetRenderDashboard();
}

function budgetMapForecastLegendToMetric_(label) {
  if (label === 'รายรับ') return 'income';
  if (label === 'รายจ่าย') return 'expense';
  if (label === 'ดุลสุทธิ') return 'net';
  return '';
}

function budgetClearInteractionFocus() {
  budgetState.interaction = {
    metricKey: '',
    expenseCategory: '',
    forecastYear: ''
  };
  budgetRenderDashboard();
}

async function budgetOpenSelectedForecastYear() {
  const year = String(budgetState.interaction?.forecastYear || '');
  if (!year) return;
  if (!(budgetState.budgetYears || []).includes(year)) {
    budgetSetStatus(`ยังไม่มี dashboard ปี ${year} ให้เปิดแบบเต็มหน้า`, 'info');
    return;
  }
  budgetState.interaction.forecastYear = '';
  budgetState.budgetYear = year;
  budgetSetValue('budgetYearSelect', year);
  await budgetLoadDashboard(year);
}

async function budgetApiGet(action, params = {}) {
  const url = new URL(localStorage.getItem(BUDGET_API_URL_KEY) || BUDGET_DEFAULT_API_URL);
  url.searchParams.set('action', action);
  Object.entries(params).forEach(([key, value]) => url.searchParams.set(key, value));
  const res = await fetch(url.toString(), { method: 'GET' });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

async function budgetApiPost(action, body = {}) {
  const url = new URL(localStorage.getItem(BUDGET_API_URL_KEY) || BUDGET_DEFAULT_API_URL);
  url.searchParams.set('action', action);
  const res = await fetch(url.toString(), {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify(body)
  });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  return res.json();
}

async function budgetApiGetCompat(action, params = {}) {
  const actions = [action].concat(BUDGET_ACTION_FALLBACKS[action] || []);
  let lastResponse = null;
  for (const candidate of actions) {
    const res = await budgetApiGet(candidate, params);
    lastResponse = res;
    if (res && res.ok) return res;
    if (!budgetIsUnknownActionMessage(res && res.message)) return res;
  }
  return lastResponse || { ok: false, message: 'Unknown action' };
}

async function budgetApiPostCompat(action, body = {}) {
  const actions = [action].concat(BUDGET_ACTION_FALLBACKS[action] || []);
  let lastResponse = null;
  for (const candidate of actions) {
    const res = await budgetApiPost(candidate, body);
    lastResponse = res;
    if (res && res.ok) return res;
    if (!budgetIsUnknownActionMessage(res && res.message)) return res;
  }
  return lastResponse || { ok: false, message: 'Unknown action' };
}

function budgetSetStatus(message, type = 'info') {
  const box = document.getElementById('budgetStatusBar');
  if (!box) return;
  box.textContent = message || '';
  box.className = `budget-status show ${type}`;
}

function budgetOpenModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('hidden');
}

function budgetCloseModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('hidden');
}

function budgetSetText(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function budgetSetValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

function budgetGetValue(id) {
  const el = document.getElementById(id);
  return el ? String(el.value).trim() : '';
}

function budgetFmtNum(n) {
  return budgetFmt.format(Number(n || 0));
}

function budgetFmtPct(n) {
  return `${Number(n || 0).toFixed(1)}%`;
}

function budgetGetErrorMessage(err) {
  return err && err.message ? err.message : String(err);
}

function budgetEscapeAttr(str) {
  return escHtml(str).replace(/`/g, '&#096;');
}

function budgetIsUnknownActionMessage(message) {
  return String(message || '').toLowerCase().indexOf('unknown action') !== -1;
}

function budgetInferFallbackYears(preferredYear) {
  const selectValue = document.getElementById('budgetYearSelect')?.value || '';
  const currentYear = new Date().getFullYear() + 543;
  const years = [preferredYear, selectValue, String(currentYear)].filter(Boolean);
  return Array.from(new Set(years)).sort();
}

function budgetGetComputedDashboard() {
  const sourceDashboard = budgetState.dashboard || {};
  const sourceForecastDashboards = budgetState.forecastDashboards || {};
  const interactionKey = [
    budgetState.interaction?.metricKey || '',
    budgetState.interaction?.expenseCategory || '',
    budgetState.interaction?.forecastYear || ''
  ].join('|');
  const cache = budgetState.computedCache || {};
  if (
    cache.sourceDashboard === sourceDashboard &&
    cache.sourceForecastDashboards === sourceForecastDashboards &&
    cache.interactionKey === interactionKey &&
    cache.value
  ) {
    return cache.value;
  }

  const value = budgetComputeDashboardFromSource_(sourceDashboard, sourceForecastDashboards);
  budgetState.computedCache = {
    sourceDashboard,
    sourceForecastDashboards,
    interactionKey,
    value
  };
  return value;
}

function budgetScheduleChartRender(computed, renderToken) {
  if (budgetChartRenderFrame) {
    window.cancelAnimationFrame(budgetChartRenderFrame);
  }
  budgetChartRenderFrame = window.requestAnimationFrame(() => {
    if (renderToken !== budgetState.renderToken) return;
    budgetEnsureCharts();
    budgetUpdateCharts(computed);
  });
}

function budgetEnsureCharts() {
  if (
    budgetBarChart &&
    budgetDoughnutSummaryChart &&
    budgetDoughnutDetailChart &&
    budgetForecastChart
  ) {
    return;
  }
  if (typeof Chart === 'undefined') return;
  budgetInitCharts();
}

function budgetComputeDashboardFromSource_(dashboard, forecastDashboards) {
  const incomeRows = budgetPrepareComputedRows(Array.isArray(dashboard.incomeRows) ? dashboard.incomeRows : []);
  const expenseRows = budgetPrepareComputedRows(Array.isArray(dashboard.expenseRows) ? dashboard.expenseRows : []);
  const stats = budgetComputeStatsFromRows_(incomeRows, expenseRows);

  incomeRows.forEach(row => {
    if (row.categoryName === 'รวมรายได้') row.displayAmount = stats.totalIncome;
  });
  expenseRows.forEach(row => {
    if (row.categoryName === 'ก.แผนงานจัดการศึกษา (รวม)') row.displayAmount = stats.educationTotal;
    if (row.categoryName === 'รวมรายจ่าย') row.displayAmount = stats.totalExpense;
  });

  const expenseInputRows = expenseRows.filter(row => row.isActive && row.rowType !== 'SUMMARY');
  const expenseSummaryRows = [];
  if (stats.educationTotal > 0 || expenseRows.some(row => row.categoryName === 'ก.แผนงานจัดการศึกษา (รวม)')) {
    expenseSummaryRows.push({ label: 'ก.แผนงานจัดการศึกษา (รวม)', amount: stats.educationTotal });
  }
  expenseInputRows.forEach(row => {
    if (row.categoryName === 'ก.แผนงานจัดการศึกษา (งบบุคลากร)') return;
    if (row.categoryName === 'ก.แผนงานจัดการศึกษา (การจัดการเรียนการสอน)') return;
    expenseSummaryRows.push({ label: row.categoryName, amount: row.displayAmount });
  });

  const forecastRows = budgetBuildForecastRows_(dashboard, forecastDashboards);

  return {
    incomeRows,
    expenseRows,
    stats: {
      totalIncome: stats.totalIncome,
      totalExpense: stats.totalExpense,
      netBalance: stats.netBalance
    },
    charts: {
      expenseSummaryRows,
      expenseDetailRows: expenseInputRows.map(row => ({ label: row.categoryName, amount: row.displayAmount })),
      forecastRows
    }
  };
}

function budgetComputeStatsFromRows_(incomeRows, expenseRows) {
  const incomeInputRows = incomeRows.filter(row => row.isActive && row.rowType !== 'SUMMARY');
  const expenseInputRows = expenseRows.filter(row => row.isActive && row.rowType !== 'SUMMARY');
  const educationDetailRows = expenseInputRows.filter(row =>
    row.categoryName === 'ก.แผนงานจัดการศึกษา (งบบุคลากร)' ||
    row.categoryName === 'ก.แผนงานจัดการศึกษา (การจัดการเรียนการสอน)'
  );
  const educationTotal = educationDetailRows.reduce((sum, row) => sum + row.displayAmount, 0);
  const totalIncome = incomeInputRows.reduce((sum, row) => sum + row.displayAmount, 0);
  const otherExpenseTotal = expenseInputRows
    .filter(row =>
      row.categoryName !== 'ก.แผนงานจัดการศึกษา (งบบุคลากร)' &&
      row.categoryName !== 'ก.แผนงานจัดการศึกษา (การจัดการเรียนการสอน)'
    )
    .reduce((sum, row) => sum + row.displayAmount, 0);
  const totalExpense = educationTotal + otherExpenseTotal;

  return {
    totalIncome,
    totalExpense,
    educationTotal,
    netBalance: totalIncome - totalExpense
  };
}

function budgetPrepareComputedRows(rows) {
  return rows.map(row => {
    const categoryName = String(row.categoryName || '').trim();
    const normalizedRowType = budgetNormalizeRowType(row.rowType, categoryName);
    const displayAmount = budgetCalcRowAmount(row);
    return {
      ...row,
      rowType: normalizedRowType,
      canEdit: normalizedRowType !== 'SUMMARY',
      displayAmount
    };
  }).sort((a, b) => Number(a.sortOrder || 0) - Number(b.sortOrder || 0));
}

function budgetCalcRowAmount(row) {
  if (!row) return 0;
  const categoryName = String(row.categoryName || '').trim();
  const rowType = String(row.rowType || '').toUpperCase();
  if (rowType === 'SUMMARY') return Number(row.finalTotal || row.totalIncome || row.totalExpense || 0);
  if (categoryName === 'ค่าธรรมเนียมการศึกษา') {
    return Number(row.studentCount || 0) * Number(row.feeRate || 0);
  }
  return Number(row.inputAmount || row.amount || row.finalTotal || 0);
}

function budgetNormalizeRowType(rowType, categoryName) {
  const normalized = String(rowType || '').trim().toUpperCase();
  if (normalized === 'INPUT' || normalized === 'SUMMARY') return normalized;
  if (
    categoryName === 'รวมรายได้' ||
    categoryName === 'ก.แผนงานจัดการศึกษา (รวม)' ||
    categoryName === 'รวมรายจ่าย'
  ) {
    return 'SUMMARY';
  }
  return 'INPUT';
}

async function budgetLoadForecastDashboards(baseYear) {
  const requestId = ++budgetState.forecastRequestId;
  const startYear = Number(baseYear || 0);
  const dashboards = { ...(budgetState.forecastDashboards || {}) };
  if (!startYear) {
    budgetState.forecastDashboards = dashboards;
    return;
  }

  const baseYearString = String(baseYear);
  if (budgetState.dashboard && baseYearString === String(budgetState.dashboard?.budgetYear || baseYearString)) {
    dashboards[baseYearString] = budgetState.dashboard;
  }
  budgetState.forecastDashboards = dashboards;

  const requests = [];
  for (let offset = 1; offset < 4; offset += 1) {
    const year = String(startYear + offset);
    if (dashboards[year]) continue;
    requests.push(
      budgetApiGetCompat('getBudgetDashboard', { budgetYear: year })
        .then(res => ({ year, res }))
        .catch(() => ({ year, res: null }))
    );
  }

  if (!requests.length) {
    budgetUpdateCharts();
    return;
  }

  const results = await Promise.all(requests);
  if (requestId !== budgetState.forecastRequestId) return;
  results.forEach(({ year, res }) => {
    if (res && res.ok && res.data) dashboards[year] = res.data;
  });
  budgetState.forecastDashboards = dashboards;
  if (String(budgetState.budgetYear || '') === baseYearString) {
    budgetUpdateCharts();
  }
}

function budgetScheduleForecastLoad(baseYear) {
  const run = () => budgetLoadForecastDashboards(baseYear);
  if (typeof window.requestIdleCallback === 'function') {
    window.requestIdleCallback(run, { timeout: 1200 });
    return;
  }
  window.setTimeout(run, 180);
}

function budgetBuildForecastRows_(dashboard, forecastDashboards) {
  const baseYear = Number(dashboard?.budgetYear || budgetState.budgetYear || 0);
  if (!baseYear) return [];

  const rows = [];
  for (let offset = 0; offset < 4; offset += 1) {
    const year = String(baseYear + offset);
    const source = forecastDashboards?.[year] || (year === String(dashboard?.budgetYear || '') ? dashboard : null);
    if (!source) {
      rows.push({ year, income: 0, expense: 0, net: 0 });
      continue;
    }
    const incomeRows = budgetPrepareComputedRows(Array.isArray(source.incomeRows) ? source.incomeRows : []);
    const expenseRows = budgetPrepareComputedRows(Array.isArray(source.expenseRows) ? source.expenseRows : []);
    const stats = budgetComputeStatsFromRows_(incomeRows, expenseRows);
    rows.push({
      year,
      income: stats.totalIncome,
      expense: stats.totalExpense,
      net: stats.netBalance
    });
  }
  return rows;
}

function budgetRecalcTermLabel() {
  const termLabelEl = document.getElementById('termLabel');
  const categoryName = budgetGetValue('categoryName');
  if (!termLabelEl) return;
  if (categoryName !== 'ค่าธรรมเนียมการศึกษา') {
    termLabelEl.value = '';
    return;
  }
  const budgetYear = Number(budgetGetValue('budgetYearInput'));
  const semester = Number(budgetGetValue('semester'));
  if (!budgetYear || !semester) {
    termLabelEl.value = '';
    return;
  }
  termLabelEl.value = `ภาค${semester}/${semester === 1 ? budgetYear : budgetYear - 1}`;
}

function budgetRecalcFormAmount() {
  const amountEl = document.getElementById('amount');
  if (!amountEl) return;
  const categoryName = budgetGetValue('categoryName');
  if (categoryName !== 'ค่าธรรมเนียมการศึกษา') return;
  const studentCount = Number(budgetGetValue('studentCount') || 0);
  const feeRate = Number(budgetGetValue('feeRate') || 0);
  amountEl.value = String(studentCount * feeRate);
}
