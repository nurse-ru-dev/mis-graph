const API_URL = 'https://script.google.com/macros/s/AKfycbx5BtQNOOZ5KFcXb8XhnofzL8Io5tsyFCCvPXvlo8-Iz1w6YxwNFx-2GsopfxeptwE4Fw/exec';
const CACHE_KEY = 'mgstd-chart-dashboard-cache-v1';
const CACHE_TTL_MS = 5 * 60 * 1000;

const STATE = {
  currentStudents: [],
  inactiveStudents: [],
  lastUpdated: '',
  loading: false,
  activeClassYear: ''
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
  } else {
    showAlert('กำลังโหลดข้อมูลจากระบบ...', 'info');
  }

  try {
    STATE.loading = true;

    const [currentResult, inactiveResult] = await Promise.all([
      apiGet('getCurrentStudents', forceRefresh),
      apiGet('getInactiveStudents', forceRefresh)
    ]);

    if (!currentResult.success) throw new Error(currentResult.message || 'โหลดข้อมูลนักศึกษาปัจจุบันไม่สำเร็จ');
    if (!inactiveResult.success) throw new Error(inactiveResult.message || 'โหลดข้อมูล inactive_students ไม่สำเร็จ');

    const payload = {
      currentStudents: Array.isArray(currentResult.data) ? currentResult.data : [],
      inactiveStudents: Array.isArray(inactiveResult.data) ? inactiveResult.data : [],
      lastUpdated: new Date().toISOString()
    };

    writeCachedDashboard(payload);
    hydrateDashboard(payload);
    hideAlert();
  } catch (error) {
    console.error(error);
    if (!cached) {
      hydrateDashboard({
        currentStudents: [],
        inactiveStudents: [],
        lastUpdated: ''
      });
      showAlert('โหลดข้อมูลไม่สำเร็จ: ' + (error.message || error), 'error');
      return;
    }

    showAlert('ใช้ข้อมูลล่าสุดที่บันทึกไว้ชั่วคราว เนื่องจากโหลดข้อมูลใหม่ไม่สำเร็จ', 'error');
  } finally {
    STATE.loading = false;
  }
}

function hydrateDashboard(payload) {
  STATE.currentStudents = payload.currentStudents || [];
  STATE.inactiveStudents = payload.inactiveStudents || [];
  STATE.lastUpdated = payload.lastUpdated || '';
  renderDashboard(buildDashboardFromState());
  renderDashboardCharts();
  renderLastUpdated();
  renderStudentTable('รายชื่อนักศึกษาปัจจุบันทั้งหมด', addSourceToRows(STATE.currentStudents || [], 'current_students'));
}

function renderDashboard(data) {
  setText('totalCurrent', formatNumber(data.totalCurrent));
  setText('year1', formatNumber(data.year1));
  setText('year2', formatNumber(data.year2));
  setText('year3', formatNumber(data.year3));
  setText('year4', formatNumber(data.year4));
  setText('totalInactive', formatNumber(data.totalInactive));
  setText('totalGraduated', formatNumber(data.totalGraduated));

  document.querySelectorAll('.stat-card').forEach(card => {
    const label = card.querySelector('p')?.textContent || '';
    const value = card.querySelector('h2')?.textContent || '0';
    card.title = `${label}: ${value} คน`;
  });
}

function buildDashboardFromState() {
  const current = STATE.currentStudents || [];
  const inactive = STATE.inactiveStudents || [];
  const graduated = inactive.filter(student => getStatusValue(student, true) === 'สำเร็จการศึกษา').length;

  return {
    totalCurrent: current.length,
    year1: current.filter(student => getClassYear(student) === '1').length,
    year2: current.filter(student => getClassYear(student) === '2').length,
    year3: current.filter(student => getClassYear(student) === '3').length,
    year4: current.filter(student => getClassYear(student) === '4').length,
    totalInactive: inactive.length,
    totalGraduated: graduated
  };
}

function renderDashboardCharts() {
  const current = STATE.currentStudents || [];
  const inactive = STATE.inactiveStudents || [];
  const allStudents = current.concat(inactive);
  const activeClassYear = STATE.activeClassYear;
  const linkedCurrent = activeClassYear
    ? current.filter(student => getClassYear(student) === activeClassYear)
    : current;
  const linkedInactive = activeClassYear
    ? inactive.filter(student => getClassYear(student) === activeClassYear)
    : inactive;
  const linkedAllStudents = linkedCurrent.concat(linkedInactive);
  const visibleYears = activeClassYear ? [Number(activeClassYear)] : [1, 2, 3, 4];
  const currentByClass = countByClassYear(current);
  const initialByClass = countByClassYear(allStudents);
  const statusByClassRows = buildStatusByClassRows(linkedAllStudents)
    .filter(row => !activeClassYear || row.classYear === activeClassYear);

  renderVerticalBarChart('classYearChart', [
    { label: 'ปี 1', value: currentByClass['1'], classYear: '1', color: '#0b1f3a' },
    { label: 'ปี 2', value: currentByClass['2'], classYear: '2', color: '#d6a62c' },
    { label: 'ปี 3', value: currentByClass['3'], classYear: '3', color: '#157347' },
    { label: 'ปี 4', value: currentByClass['4'], classYear: '4', color: '#2563eb' }
  ], activeClassYear);

  renderRetentionChart('retentionChart', visibleYears.map(year => {
    const initial = initialByClass[String(year)] || 0;
    const remaining = currentByClass[String(year)] || 0;

    return {
      label: `ปี ${year}`,
      initial,
      remaining,
      percent: percentOf(remaining, initial)
    };
  }));

  renderTrendChart('admissionTrendChart', buildAdmissionTrendRows(linkedAllStudents));
  renderStatusByClassChart('statusChart', statusByClassRows);
  renderExecutiveSummary(linkedCurrent, linkedInactive, statusByClassRows);
}

function countByClassYear(rows) {
  const counts = { '1': 0, '2': 0, '3': 0, '4': 0 };

  rows.forEach(student => {
    const classYear = getClassYear(student);
    if (counts[classYear] !== undefined) counts[classYear]++;
  });

  return counts;
}

function buildStatusByClassRows(rows) {
  const classRows = ['1', '2', '3', '4'].map(classYear => ({
    classYear,
    label: `ปี ${classYear}`,
    total: 0,
    statuses: {}
  }));
  const byClass = classRows.reduce((acc, row) => {
    acc[row.classYear] = row;
    return acc;
  }, {});

  rows.forEach(student => {
    const classYear = getClassYear(student);
    if (!byClass[classYear]) return;

    const isInactive = isInactiveStudent(student);
    const status = getStatusValue(student, isInactive) || 'ไม่ระบุสถานะ';
    byClass[classYear].total++;
    byClass[classYear].statuses[status] = (byClass[classYear].statuses[status] || 0) + 1;
  });

  return classRows;
}

function buildAdmissionTrendRows(rows) {
  const counts = {};

  rows.forEach(student => {
    const year = String(getVal(student, ['ADMISSION_YEAR', 'admission_year', 'COHORT_YEAR', 'cohort_year', 'PRE_NO']) || '').trim();
    if (!year) return;
    counts[year] = (counts[year] || 0) + 1;
  });

  const dataYears = Object.keys(counts)
    .map(year => Number(year))
    .filter(year => !Number.isNaN(year));
  const currentThaiYear = new Date().getFullYear() + 543;
  const endYear = Math.max(currentThaiYear, ...dataYears);
  const startYear = endYear - 4;

  return Array.from({ length: 5 }, (_, index) => {
    const year = String(startYear + index);
    return { year, count: counts[year] || 0 };
  });
}

function renderVerticalBarChart(elementId, rows, activeClassYear = '') {
  const target = document.getElementById(elementId);
  if (!target) return;

  const max = Math.max(...rows.map(row => row.value), 1);
  const axisMax = Math.max(max, 4);
  const ticks = buildAxisTicks(axisMax);

  target.innerHTML = `
    <div class="vertical-chart">
      <div class="y-axis">${ticks.map(tick => `<span>${tick}</span>`).join('')}</div>
      <div class="plot-area">
        <div class="grid-lines">${ticks.map(() => '<span></span>').join('')}</div>
        <div class="vertical-bars">
          ${rows.map(row => {
            const height = Math.max((row.value / axisMax) * 100, row.value > 0 ? 4 : 0);
            const tooltip = `${row.label}: ${row.value} คน`;

            return `
              <button class="vbar ${activeClassYear === row.classYear ? 'is-selected' : ''}" style="--bar-height:${height}%;--bar-color:${row.color}" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}" onclick="selectDashboardGroup('currentClass', '${escapeAttr(row.classYear)}')">
                <span class="vbar-value">${formatNumber(row.value)}</span>
                <span class="vbar-fill"></span>
                <span class="vbar-label">${escapeHtml(row.label)}</span>
              </button>
            `;
          }).join('')}
        </div>
      </div>
    </div>
    <div class="axis-caption">
      <span>แกน Y = จำนวนคน</span>
      <span>แกน X = ชั้นปี</span>
    </div>
  `;
}

function renderRetentionChart(elementId, rows) {
  const target = document.getElementById(elementId);
  if (!target) return;

  const max = Math.max(...rows.flatMap(row => [row.initial, row.remaining]), 1);
  const axisMax = Math.max(max, 4);
  const ticks = buildAxisTicks(axisMax);

  target.innerHTML = `
    <div class="chart-legend">
      <span><i class="legend-initial"></i>จำนวนแรก</span>
      <span><i class="legend-remaining"></i>คงอยู่</span>
    </div>
    <div class="vertical-chart">
      <div class="y-axis">${ticks.map(tick => `<span>${tick}</span>`).join('')}</div>
      <div class="plot-area">
        <div class="grid-lines">${ticks.map(() => '<span></span>').join('')}</div>
        <div class="vertical-bars grouped">
          ${rows.map(row => `
            <div class="vbar-group" title="${escapeAttr(`${row.label}: คงอยู่ ${row.remaining} จาก ${row.initial} คน (${formatPercent(row.percent)}%)`)}">
              ${renderGroupedVerticalBar(row.initial, axisMax, 'จำนวนแรก', row.label, 'initial', percentOf(row.initial, row.initial || max))}
              ${renderGroupedVerticalBar(row.remaining, axisMax, 'คงอยู่', row.label, 'remaining', row.percent)}
              <span class="vbar-label">${escapeHtml(row.label)}</span>
            </div>
          `).join('')}
        </div>
      </div>
    </div>
    <div class="axis-caption">
      <span>แกน Y = จำนวนคน</span>
      <span>แกน X = ชั้นปี</span>
    </div>
  `;
}

function renderGroupedVerticalBar(value, max, series, label, type, percent) {
  const height = Math.max((value / max) * 100, value > 0 ? 4 : 0);
  const tooltip = `${label} ${series}: ${value} คน (${formatPercent(percent)}%)`;
  const classYear = label.replace('ปี ', '');
  const group = type === 'initial' ? 'allClass' : 'currentClass';

  return `
    <button class="vbar mini ${type}" style="--bar-height:${height}%" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}" onclick="selectDashboardGroup('${group}', '${escapeAttr(classYear)}')">
      <span class="vbar-value">${formatNumber(value)}</span>
      <span class="vbar-percent">${formatPercent(percent)}%</span>
      <span class="vbar-fill"></span>
    </button>
  `;
}

function renderTrendChart(elementId, rows) {
  const target = document.getElementById(elementId);
  if (!target) return;

  if (!rows.length) {
    target.innerHTML = `<div class="empty">ยังไม่มีข้อมูลปีที่รับเข้า</div>`;
    return;
  }

  const max = Math.max(...rows.map(row => row.count), 1);
  const axisMax = Math.max(max, 4);
  const ticks = buildAxisTicks(axisMax);
  const points = rows.map((row, index) => ({
    x: rows.length === 1 ? 50 : (index / (rows.length - 1)) * 100,
    y: 100 - ((row.count / axisMax) * 100),
    row
  }));
  const polyline = points.map(point => `${point.x},${point.y}`).join(' ');

  target.innerHTML = `
    <div class="line-chart">
      <div class="y-axis">${ticks.map(tick => `<span>${tick}</span>`).join('')}</div>
      <div class="line-plot">
        <div class="grid-lines">${ticks.map(() => '<span></span>').join('')}</div>
        <svg class="line-svg" viewBox="0 0 100 100" preserveAspectRatio="none" aria-hidden="true">
          <polyline points="${polyline}" />
        </svg>
        <div class="line-points">
          ${points.map(point => {
            const tooltip = `ปีรับเข้า ${point.row.year}: ${point.row.count} คน`;
            return `
              <button class="line-point" style="left:${point.x}%;top:${point.y}%" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}" onclick="selectDashboardGroup('admissionYear', '${escapeAttr(point.row.year)}')">
                <span class="line-value">${formatNumber(point.row.count)}</span>
              </button>
            `;
          }).join('')}
        </div>
        <div class="line-x-axis">${rows.map(row => `<span>${escapeHtml(row.year)}</span>`).join('')}</div>
      </div>
    </div>
    <div class="axis-caption">
      <span>แกน Y = จำนวนคน</span>
      <span>แกน X = ปีที่รับเข้า ย้อนหลัง 5 ปี</span>
    </div>
  `;
}

function renderStatusByClassChart(elementId, rows) {
  const target = document.getElementById(elementId);
  if (!target) return;

  if (!rows.length || rows.every(row => row.total === 0)) {
    target.innerHTML = `<div class="empty">ยังไม่มีข้อมูลสถานะ</div>`;
    return;
  }

  const colors = ['#d6a62c', '#0b1f3a', '#157347', '#b42318', '#2563eb', '#7c3aed', '#64748b', '#f97316'];
  const statusColorMap = {};
  let colorIndex = 0;

  target.innerHTML = `
    <div class="status-class-grid">
      ${rows.map(row => {
        const statusEntries = Object.entries(row.statuses)
          .map(([status, count]) => ({ status, count }))
          .sort((a, b) => b.count - a.count || a.status.localeCompare(b.status));

        return `
          <section class="status-class-card" title="${escapeAttr(`${row.label}: รวม ${row.total} คน`)}">
            <div class="status-class-head">
              <strong>${escapeHtml(row.label)}</strong>
              <span>${formatNumber(row.total)} คน</span>
            </div>
            ${statusEntries.length ? `
              <div class="status-stack" aria-label="${escapeAttr(`${row.label} แยกตามสถานะ`)}">
                ${statusEntries.map(item => {
                  if (!statusColorMap[item.status]) {
                    statusColorMap[item.status] = colors[colorIndex % colors.length];
                    colorIndex++;
                  }
                  const percent = percentOf(item.count, row.total);
                  const tooltip = `${row.label} ${item.status}: ${item.count} คน (${formatPercent(percent)}%)`;

                  return `
                    <button class="status-segment" style="width:${Math.max(percent, item.count > 0 ? 6 : 0)}%;background:${statusColorMap[item.status]}" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}" onclick="selectDashboardGroup('statusClass', '${escapeAttr(row.classYear)}', '${escapeAttr(item.status)}')">
                      ${percent >= 18 ? formatNumber(item.count) : ''}
                    </button>
                  `;
                }).join('')}
              </div>
              <div class="status-breakdown">
                ${statusEntries.map(item => {
                  const percent = percentOf(item.count, row.total);
                  const tooltip = `${row.label} ${item.status}: ${item.count} คน (${formatPercent(percent)}%)`;

                  return `
                    <button class="status-mini-item" title="${escapeAttr(tooltip)}" aria-label="${escapeAttr(tooltip)}" onclick="selectDashboardGroup('statusClass', '${escapeAttr(row.classYear)}', '${escapeAttr(item.status)}')">
                      <span class="status-dot" style="background:${statusColorMap[item.status]}"></span>
                      <span>${escapeHtml(item.status)}</span>
                      <strong>${formatNumber(item.count)}</strong>
                      <small>${formatPercent(percent)}%</small>
                    </button>
                  `;
                }).join('')}
              </div>
            ` : `<div class="empty small-empty">ยังไม่มีข้อมูล</div>`}
          </section>
        `;
      }).join('')}
    </div>
  `;
}

function renderExecutiveSummary(current, inactive, statusByClassRows) {
  const target = document.getElementById('executiveSummary');
  if (!target) return;

  const totalCurrent = current.length;
  const totalInactive = inactive.length;
  const totalAll = totalCurrent + totalInactive;
  const graduated = inactive.filter(student => getStatusValue(student, true) === 'สำเร็จการศึกษา').length;
  const largestClass = statusByClassRows.slice().sort((a, b) => b.total - a.total)[0];
  const riskStatuses = ['ลาออก', 'ไล่ออก', 'พ้นสภาพ', 'โอนย้ายสถานศึกษา', 'เสียชีวิต'];
  const riskCount = inactive.filter(student => riskStatuses.includes(getStatusValue(student, true))).length;
  const retention = percentOf(totalCurrent, totalAll || totalCurrent);
  const graduateRate = percentOf(graduated, totalAll);

  target.innerHTML = `
    <h3>สรุปเพื่อการตัดสินใจ</h3>
    <p>
      ปัจจุบันมีนักศึกษาที่คงอยู่ในระบบ ${formatNumber(totalCurrent)} คน จากรายชื่อทั้งหมด ${formatNumber(totalAll || totalCurrent)} คน
      คิดเป็นอัตราคงอยู่ประมาณ ${formatPercent(retention)}% โดยชั้นปีที่มีรายชื่อมากที่สุดคือ ${escapeHtml(largestClass?.label || '-')}
      จำนวน ${formatNumber(largestClass?.total || 0)} คน
    </p>
    <p>
      กลุ่มที่สำเร็จการศึกษามี ${formatNumber(graduated)} คน หรือประมาณ ${formatPercent(graduateRate)}% ของรายชื่อทั้งหมด
      และกลุ่มที่ควรติดตามเป็นพิเศษมี ${formatNumber(riskCount)} คน
    </p>
    <div class="summary-actions">
      <button type="button" onclick="selectDashboardGroup('currentAll')">นักศึกษาคงอยู่</button>
      <button type="button" onclick="selectDashboardGroup('inactiveAll')">กลุ่มที่ต้องติดตาม</button>
    </div>
  `;
}

function selectDashboardGroup(type, value = '', extra = '') {
  const current = addSourceToRows(STATE.currentStudents || [], 'current_students');
  const inactive = addSourceToRows(STATE.inactiveStudents || [], 'inactive_students');
  const all = current.concat(inactive);
  const isClassFilter = ['currentClass', 'allClass', 'statusClass'].includes(type);
  const selectedClassYear = isClassFilter ? String(value) : '';

  if ((type === 'currentClass' || type === 'allClass') && STATE.activeClassYear === selectedClassYear) {
    clearDashboardSelection();
    return;
  }

  if (isClassFilter && selectedClassYear) {
    STATE.activeClassYear = selectedClassYear;
    renderDashboardCharts();
  } else if (['currentAll', 'inactiveAll', 'graduated'].includes(type)) {
    STATE.activeClassYear = '';
    renderDashboardCharts();
  }

  const scopedAll = STATE.activeClassYear
    ? all.filter(row => getClassYear(row) === STATE.activeClassYear)
    : all;
  let rows = [];
  let title = 'เลือกข้อมูลจากกราฟ';

  if (type === 'currentAll') {
    rows = current;
    title = 'นักศึกษาปัจจุบันทั้งหมด';
  } else if (type === 'currentClass') {
    rows = current.filter(row => getClassYear(row) === String(value));
    title = `นักศึกษาปัจจุบัน ปี ${value}`;
  } else if (type === 'inactiveAll') {
    rows = inactive;
    title = 'พ้นสภาพ/จบ/ลาออกทั้งหมด';
  } else if (type === 'graduated') {
    rows = inactive.filter(row => getStatusValue(row, true) === 'สำเร็จการศึกษา');
    title = 'สำเร็จการศึกษา';
  } else if (type === 'allClass') {
    rows = all.filter(row => getClassYear(row) === String(value));
    title = `รายชื่อทั้งหมด ปี ${value}`;
  } else if (type === 'statusClass') {
    rows = all.filter(row => {
      const isInactive = row._source === 'inactive_students';
      return getClassYear(row) === String(value) && getStatusValue(row, isInactive) === extra;
    });
    title = `ปี ${value} สถานะ ${extra}`;
  } else if (type === 'admissionYear') {
    rows = scopedAll.filter(row => {
      const admissionYear = String(getVal(row, ['ADMISSION_YEAR', 'admission_year', 'COHORT_YEAR', 'cohort_year', 'PRE_NO']) || '').trim();
      return admissionYear === String(value);
    });
    title = STATE.activeClassYear
      ? `นักศึกษาเข้าใหม่ปี ${value} เฉพาะปี ${STATE.activeClassYear}`
      : `นักศึกษาเข้าใหม่ปี ${value}`;
  }

  renderSelectionSummary(title, rows);
  renderStudentTable(title, rows);
}

function renderSelectionSummary(title, rows) {
  const panel = document.getElementById('selectionPanel');
  const titleEl = document.getElementById('selectionTitle');
  const subtitleEl = document.getElementById('selectionSubtitle');
  if (!panel || !titleEl || !subtitleEl) return;

  titleEl.textContent = title;
  subtitleEl.textContent = `พบ ${formatNumber(rows.length)} คน จากกลุ่มที่เลือก`;
  panel.classList.remove('hidden');
  panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function clearDashboardSelection() {
  STATE.activeClassYear = '';
  document.getElementById('selectionPanel')?.classList.add('hidden');
  renderDashboardCharts();
  renderStudentTable('รายชื่อนักศึกษาปัจจุบันทั้งหมด', addSourceToRows(STATE.currentStudents || [], 'current_students'));
}

function renderStudentTable(title, rows) {
  const titleEl = document.getElementById('studentTableTitle');
  const subtitleEl = document.getElementById('studentTableSubtitle');
  const countEl = document.getElementById('studentTableCount');
  const tbody = document.getElementById('studentTableBody');
  if (!titleEl || !subtitleEl || !countEl || !tbody) return;

  titleEl.textContent = title;
  subtitleEl.textContent = rows.length
    ? 'รายชื่อนี้ถูกกรองตามการเลือกจากการ์ดหรือกราฟ'
    : 'ไม่พบรายชื่อในกลุ่มที่เลือก';
  countEl.textContent = `${formatNumber(rows.length)} คน`;

  if (!rows.length) {
    tbody.innerHTML = `<tr><td colspan="7" class="table-empty">ไม่พบข้อมูลนักศึกษา</td></tr>`;
    return;
  }

  tbody.innerHTML = rows.map(row => {
    const isInactive = row._source === 'inactive_students' || isInactiveStudent(row);
    const code = getVal(row, ['รหัสนักศึกษา', 'student_code', 'STUDENT_CODE']);
    const name = getVal(row, ['FULL_NAME', 'full_name', 'ชื่อ-นามสกุล', 'NAME']);
    const classYear = getClassYear(row);
    const cohort = getVal(row, ['ADMISSION_YEAR', 'admission_year', 'COHORT_YEAR', 'cohort_year', 'PRE_NO']);
    const status = getStatusValue(row, isInactive) || '-';
    const email = getVal(row, ['EMAIL', 'email']);
    const source = row._source === 'inactive_students' ? 'inactive_students' : 'current_students';

    return `
      <tr>
        <td>${escapeHtml(code)}</td>
        <td><strong>${escapeHtml(name || '-')}</strong></td>
        <td>${classYear ? `ปี ${escapeHtml(classYear)}` : '-'}</td>
        <td>${escapeHtml(cohort || '-')}</td>
        <td>${renderStatusPill(status, isInactive)}</td>
        <td>${escapeHtml(email || '-')}</td>
        <td><span class="source-pill ${source === 'inactive_students' ? 'source-inactive' : ''}">${escapeHtml(source)}</span></td>
      </tr>
    `;
  }).join('');
}

function renderStatusPill(status, inactive = false) {
  const className = inactive ? 'status-pill inactive' : 'status-pill active';
  return `<span class="${className}">${escapeHtml(status || '-')}</span>`;
}

function buildAxisTicks(max) {
  const top = Math.max(4, Math.ceil(max / 5) * 5);
  const step = Math.max(1, Math.ceil(top / 4));
  const ticks = [];

  for (let value = top; value >= 0; value -= step) {
    ticks.push(value);
  }

  if (ticks[ticks.length - 1] !== 0) ticks.push(0);
  return ticks;
}

function percentOf(value, total) {
  if (total <= 0) return 0;
  return Number(((Number(value || 0) / total) * 100).toFixed(2));
}

function formatPercent(value) {
  const number = Number(value || 0);
  return number.toFixed(2).replace(/\.00$/, '');
}

function getStatusValue(student, inactive = false) {
  if (inactive) {
    return String(getVal(student, ['INACTIVE_STATUS', 'inactive_status', 'STUDENT_STATUS', 'student_status']) || '').trim();
  }

  return String(getVal(student, ['STUDENT_STATUS', 'student_status']) || '').trim();
}

function getClassYear(row) {
  return String(getVal(row, ['CLASS_YEAR', 'class_year']) || '').trim();
}

function isInactiveStudent(student) {
  return String(getVal(student, ['STATUS_GROUP', 'status_group']) || '').toUpperCase() === 'INACTIVE'
    || String(getVal(student, ['INACTIVE_STATUS', 'inactive_status']) || '').trim() !== ''
    || student._source === 'inactive_students';
}

function addSourceToRows(rows, source) {
  return rows.map(row => Object.assign({ _source: source }, row));
}

function getVal(obj, keys) {
  for (const key of keys) {
    if (obj && obj[key] !== undefined && obj[key] !== null && obj[key] !== '') {
      return obj[key];
    }
  }

  return '';
}

async function apiGet(action, forceRefresh = false) {
  const url = new URL(API_URL);
  url.searchParams.set('action', action);
  if (forceRefresh) url.searchParams.set('_ts', Date.now());

  const response = await fetch(url.toString(), {
    method: 'GET',
    redirect: 'follow',
    cache: forceRefresh ? 'no-store' : 'default'
  });
  const text = await response.text();

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error('API ไม่ได้ส่ง JSON กลับมา กรุณาตรวจ doGet ใน Apps Script');
  }
}

function renderLastUpdated() {
  const el = document.getElementById('lastUpdated');
  if (!el) return;

  if (!STATE.lastUpdated) {
    el.textContent = 'ยังไม่มีข้อมูลอัปเดต';
    return;
  }

  const date = new Date(STATE.lastUpdated);
  el.textContent = `อัปเดตข้อมูล: ${date.toLocaleDateString('th-TH', {
    day: '2-digit',
    month: 'short',
    year: 'numeric'
  })} ${date.toLocaleTimeString('th-TH', {
    hour: '2-digit',
    minute: '2-digit'
  })}`;
}

function setText(id, value) {
  const el = document.getElementById(id);
  if (el) el.textContent = value;
}

function formatNumber(value) {
  return Number(value || 0).toLocaleString('th-TH');
}

function showAlert(message, type = 'info') {
  const alertBox = document.getElementById('alertBox');
  if (!alertBox) return;
  alertBox.textContent = message;
  alertBox.className = `alert ${type}`;
  alertBox.classList.remove('hidden');
}

function hideAlert() {
  document.getElementById('alertBox')?.classList.add('hidden');
}

function readCachedDashboard() {
  try {
    const raw = localStorage.getItem(CACHE_KEY);
    if (!raw) return null;

    const parsed = JSON.parse(raw);
    if (!parsed?.timestamp || !parsed?.data) return null;
    if ((Date.now() - parsed.timestamp) > CACHE_TTL_MS) return null;

    return parsed.data;
  } catch (error) {
    console.warn('Failed to read cached student dashboard', error);
    return null;
  }
}

function writeCachedDashboard(data) {
  try {
    localStorage.setItem(CACHE_KEY, JSON.stringify({
      timestamp: Date.now(),
      data
    }));
  } catch (error) {
    console.warn('Failed to cache student dashboard', error);
  }
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll('`', '&#096;');
}
