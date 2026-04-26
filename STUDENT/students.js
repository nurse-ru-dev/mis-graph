const CONFIG = {
  API_BASE_URL: 'https://script.google.com/macros/s/AKfycbzrXF0uWFMry_gZBTetDPyj-mKZtrEWU5Oq3Kz_ZlzcRApumP2tpLUOL9a7b7mIZV8cFQ/exec',
  CURRENT_FIRST_YEAR_COHORT: 68,
  CACHE_KEY: 'student-dashboard-students-cache-v1',
  CACHE_TTL_MS: 5 * 60 * 1000
};

const state = {
  raw: null,
  loadingTimer: null
};

document.addEventListener('DOMContentLoaded', () => {
  document.getElementById('btnRefresh')?.addEventListener('click', () => loadDashboard({ forceRefresh: true }));
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
  renderAll();
}

function renderAll() {
  if (!state.raw) return;

  setText('kpiStudents', formatNumber(state.raw.kpis.totalCurrentStudents || 0));
  renderStudentsByCohort(state.raw.charts.studentsByCohort || []);
  renderRetentionChart(state.raw.charts.retentionByCohort || []);
}

function renderStudentsByCohort(items) {
  const el = document.getElementById('cohortChart');
  if (!el) return;

  if (!items.length) {
    el.innerHTML = `<div class="empty-state">ไม่พบข้อมูล</div>`;
    return;
  }

  const sortedItems = [...items].sort((a, b) => {
    const yearA = Number(getStudyYearNumber(a.cohort));
    const yearB = Number(getStudyYearNumber(b.cohort));
    return yearA - yearB;
  });

  const max = Math.max(...sortedItems.map(x => x.currentCount || 0), 1);

  el.innerHTML = sortedItems.map(item => {
    const height = Math.max(12, (Number(item.currentCount || 0) / max) * 100);

    return `
      <div class="vertical-bar-item">
        <div class="vertical-bar-value">${formatNumber(item.currentCount)} คน</div>
        <div class="vertical-bar-wrap">
          <div class="vertical-bar" style="height:${height}%"></div>
        </div>
        <div class="vertical-bar-label">${escapeHtml(getStudyYearLabel(item.cohort))}</div>
      </div>
    `;
  }).join('');
}
function getStudyYearNumber(cohortValue) {
  const cohort = Number(cohortValue);
  if (!Number.isFinite(cohort)) return 999;

  const studyYear = CONFIG.CURRENT_FIRST_YEAR_COHORT - cohort + 1;
  return studyYear;
}
function renderRetentionChart(items) {
  const el = document.getElementById('retentionChart');
  if (!el) return;

  if (!items.length) {
    el.innerHTML = `<div class="empty-state">ไม่พบข้อมูล</div>`;
    return;
  }

  el.innerHTML = items.map(item => {
    const originalCount = Number(item.originalCount || 0);
    const currentCount = Number(item.currentCount || 0);
    const retentionRate = Math.max(0, Math.min(100, Number(item.retentionRate || 0)));
    const hasRetentionData = originalCount > 0 || currentCount > 0;

    return `
    <div class="retention-row ${hasRetentionData ? '' : 'retention-row-empty'}">
      <div class="retention-head">
        <div class="retention-title">${escapeHtml(getStudyYearLabel(item.cohort))}</div>
        <div class="retention-value">${hasRetentionData ? `${retentionRate.toFixed(2)}%` : 'ยังไม่มีข้อมูล'}</div>
      </div>
      <div class="retention-lines">
        <div class="retention-line">
          <div class="retention-line-head">
            <span class="retention-line-label">
              <span class="retention-dot ${hasRetentionData ? 'retention-dot-target' : 'retention-dot-empty'}"></span>
              100%
            </span>
            <span class="retention-line-value">${hasRetentionData ? 'เกณฑ์เต็ม' : 'ยังไม่มีข้อมูล'}</span>
          </div>
          <div class="retention-track ${hasRetentionData ? 'retention-track-target' : 'retention-track-empty'}">
            <div class="retention-fill ${hasRetentionData ? 'retention-fill-target' : 'retention-fill-empty'}" style="width:100%"></div>
          </div>
        </div>
        <div class="retention-line">
          <div class="retention-line-head">
            <span class="retention-line-label">
              <span class="retention-dot retention-dot-original"></span>
              เดิม
            </span>
            <span class="retention-line-value">${formatNumber(originalCount)} คน</span>
          </div>
          <div class="retention-track retention-track-original">
            <div class="retention-fill retention-fill-original" style="width:100%"></div>
          </div>
        </div>
        <div class="retention-line">
          <div class="retention-line-head">
            <span class="retention-line-label">
              <span class="retention-dot ${hasRetentionData ? 'retention-dot-current' : 'retention-dot-empty'}"></span>
              ปัจจุบัน
            </span>
            <span class="retention-line-value">${formatNumber(currentCount)} คน</span>
          </div>
          <div class="retention-track ${hasRetentionData ? 'retention-track-current' : 'retention-track-empty'}">
            <div class="retention-fill ${hasRetentionData ? 'retention-fill-current' : 'retention-fill-empty'}" style="width:${hasRetentionData ? retentionRate : 100}%"></div>
          </div>
        </div>
      </div>
      <div class="retention-meta">
        ${hasRetentionData ? `อัตราการคงอยู่ ${retentionRate.toFixed(2)}%` : 'รอข้อมูลนักศึกษารุ่นนี้'}
      </div>
    </div>
  `;
  }).join('');
}

function getStudyYearLabel(cohortValue) {
  const cohort = Number(cohortValue);
  if (!Number.isFinite(cohort)) return `รุ่น ${cohortValue}`;

  const studyYear = CONFIG.CURRENT_FIRST_YEAR_COHORT - cohort + 1;
  if (studyYear >= 1 && studyYear <= 10) return `ชั้นปีที่ ${studyYear}`;
  return `รุ่น ${cohortValue}`;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('th-TH');
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
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
    console.warn('Failed to read cached student dashboard', err);
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
    console.warn('Failed to cache student dashboard', err);
  }
}
