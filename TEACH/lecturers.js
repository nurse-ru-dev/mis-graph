const CONFIG = {
  API_BASE_URL: 'https://script.google.com/macros/s/AKfycbzrXF0uWFMry_gZBTetDPyj-mKZtrEWU5Oq3Kz_ZlzcRApumP2tpLUOL9a7b7mIZV8cFQ/exec',
  TARGET_RATIO: 8,
  CACHE_KEY: 'student-dashboard-lecturers-cache-v1',
  CACHE_TTL_MS: 5 * 60 * 1000
};

const state = {
  raw: null,
  scenario: {
    lecturers: 0,
    students: 0,
    plannedIntake: 0
  },
  renderFrame: null,
  loadingTimer: null
};

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btnRefresh')?.addEventListener('click', () => loadDashboard({ forceRefresh: true }));
  document.getElementById('btnApplyScenario')?.addEventListener('click', applyScenarioFromInputs);
  document.getElementById('btnResetScenario')?.addEventListener('click', resetScenarioToRaw);

  document.querySelectorAll('.step-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.stepTarget;
      const step = Number(btn.dataset.step || 0);
      stepInput(target, step);
      applyScenarioFromInputs();
    });
  });

  ['inputLecturers', 'inputStudents', 'inputPlannedIntake'].forEach(id => {
    document.getElementById(id)?.addEventListener('input', scheduleScenarioRender);
  });

  loadDashboard();
});

async function loadDashboard(options = {}) {
  const { forceRefresh = false } = options;
  const cached = !forceRefresh ? readCachedDashboard() : null;

  if (cached) {
    hydrateDashboard(cached);
    showLoading(false);
  } else {
    showLoading(true);
  }

  try {
    const url = forceRefresh
      ? `${CONFIG.API_BASE_URL}?action=dashboard&_ts=${Date.now()}`
      : `${CONFIG.API_BASE_URL}?action=dashboard`;
    const res = await fetch(url, {
      cache: forceRefresh ? 'no-store' : 'default'
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);

    const data = await res.json();
    if (!data.success) throw new Error(data.message || 'โหลดข้อมูลไม่สำเร็จ');

    writeCachedDashboard(data);
    hydrateDashboard(data);
    clearError();
  } catch (err) {
    console.error(err);
    if (!cached) {
      showError(err.message || 'โหลดข้อมูลไม่สำเร็จ');
    }
  } finally {
    showLoading(false);
  }
}

function hydrateDashboard(data) {
  state.raw = data;
  state.scenario.lecturers = Number(data.kpis.totalActiveLecturers || 0);
  state.scenario.students = Number(data.kpis.totalCurrentStudents || 0);
  state.scenario.plannedIntake = Number(data.kpis.plannedNewIntake || 0);
  syncInputsWithScenario();
  renderAll();
}

function applyScenarioFromInputs() {
  state.scenario.lecturers = clampNonNegative(readNumber('inputLecturers'));
  state.scenario.students = clampNonNegative(readNumber('inputStudents'));
  state.scenario.plannedIntake = clampNonNegative(readNumber('inputPlannedIntake'));
  renderAll();
}

function scheduleScenarioRender() {
  if (state.renderFrame) {
    cancelAnimationFrame(state.renderFrame);
  }

  state.renderFrame = requestAnimationFrame(() => {
    state.renderFrame = null;
    applyScenarioFromInputs();
  });
}

function resetScenarioToRaw() {
  if (!state.raw) return;
  state.scenario.lecturers = Number(state.raw.kpis.totalActiveLecturers || 0);
  state.scenario.students = Number(state.raw.kpis.totalCurrentStudents || 0);
  state.scenario.plannedIntake = Number(state.raw.kpis.plannedNewIntake || 0);
  syncInputsWithScenario();
  renderAll();
}

function syncInputsWithScenario() {
  setInputValue('inputLecturers', state.scenario.lecturers);
  setInputValue('inputStudents', state.scenario.students);
  setInputValue('inputPlannedIntake', state.scenario.plannedIntake);
}

function renderAll() {
  if (!state.raw) return;

  const kpis = calculateScenarioKpis(
    state.scenario.lecturers,
    state.scenario.students,
    state.scenario.plannedIntake
  );

  renderKpis(kpis);
  renderGauge(kpis);
}

function calculateScenarioKpis(lecturers, students, plannedIntake) {
  const ratioValue = lecturers > 0 ? students / lecturers : 0;
  const ratioDisplay = lecturers > 0 ? `1 : ${ratioValue.toFixed(2)}` : 'ไม่มีอาจารย์';
  const intakeCapacityWithoutHiring = Math.max(0, (lecturers * CONFIG.TARGET_RATIO) - students);
  const additionalLecturersRequired = Math.max(
    0,
    Math.ceil((students + plannedIntake) / CONFIG.TARGET_RATIO) - lecturers
  );

  return {
    totalActiveLecturers: lecturers,
    totalCurrentStudents: students,
    currentRatioValue: Number(ratioValue.toFixed(2)),
    currentRatioDisplay: ratioDisplay,
    intakeCapacityWithoutHiring,
    plannedNewIntake: plannedIntake,
    additionalLecturersRequired
  };
}

function renderKpis(kpis) {
  setText('kpiLecturers', formatNumber(kpis.totalActiveLecturers));
  setText('kpiRatio', kpis.currentRatioDisplay);
  setText('kpiCapacity', `${formatNumber(kpis.intakeCapacityWithoutHiring)} คน`);
  setText('kpiNeedLecturers', `${formatNumber(kpis.additionalLecturersRequired)} คน`);
}

function renderGauge(kpis) {
  setText('gaugeRatioText', kpis.currentRatioDisplay);

  const status = kpis.currentRatioValue <= CONFIG.TARGET_RATIO
    ? 'อยู่ในเกณฑ์มาตรฐาน'
    : 'เกินเกณฑ์มาตรฐาน';

  setText('gaugeStatusText', `เกณฑ์เป้าหมายไม่เกิน 1 : 8 • สถานะ: ${status}`);
}

function stepInput(target, step) {
  const map = {
    lecturers: 'inputLecturers',
    students: 'inputStudents',
    plannedIntake: 'inputPlannedIntake'
  };

  const id = map[target];
  const el = document.getElementById(id);
  if (!el) return;

  const current = clampNonNegative(Number(el.value || 0));
  el.value = Math.max(0, current + step);
}

function readNumber(id) {
  return Number(document.getElementById(id)?.value || 0);
}

function setInputValue(id, value) {
  const el = document.getElementById(id);
  if (el) el.value = value;
}

function clampNonNegative(value) {
  return Math.max(0, Number.isFinite(value) ? value : 0);
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('th-TH');
}

function showLoading(isLoading) {
  const el = document.getElementById('loadingState');
  if (!el) return;

  if (isLoading) {
    if (state.loadingTimer) return;

    state.loadingTimer = setTimeout(() => {
      el.style.display = 'flex';
      state.loadingTimer = null;
    }, 180);
    return;
  }

  if (state.loadingTimer) {
    clearTimeout(state.loadingTimer);
    state.loadingTimer = null;
  }

  el.style.display = 'none';
}

function showError(message) {
  const el = document.getElementById('errorState');
  if (!el) return;
  el.textContent = message;
  el.style.display = 'block';
}

function clearError() {
  const el = document.getElementById('errorState');
  if (!el) return;
  el.textContent = '';
  el.style.display = 'none';
}

function readCachedDashboard() {
  try {
    const raw = localStorage.getItem(CONFIG.CACHE_KEY);
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (!parsed?.timestamp || !parsed?.data) return null;
    if ((Date.now() - parsed.timestamp) > CONFIG.CACHE_TTL_MS) return null;

    return parsed.data;
  } catch (err) {
    console.warn('Failed to read cached lecturer dashboard', err);
    return null;
  }
}

function writeCachedDashboard(data) {
  try {
    localStorage.setItem(CONFIG.CACHE_KEY, JSON.stringify({
      timestamp: Date.now(),
      data
    }));
  } catch (err) {
    console.warn('Failed to cache lecturer dashboard', err);
  }
}
