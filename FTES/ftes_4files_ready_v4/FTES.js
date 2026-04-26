const API_URL = "https://script.google.com/macros/s/AKfycbwjVL_hoCe_qucVpftJjU-f01vevdwh8QnJQrl5TinYXXlMaf5rVOK5Ylr1L8yAidcnsg/exec";

const state = {
  raw: [],
  selectedRow: null,
  compareChart: null,
  ratioChart: null
};

const els = {};

document.addEventListener("DOMContentLoaded", async () => {
  bindElements();
  bindEvents();
  await init();
});

function bindElements() {
  els.openLevelSelect = document.getElementById("openLevelSelect");
  els.summaryText = document.getElementById("summaryText");

  els.kpiTotalFTES = document.getElementById("kpiTotalFTES");
  els.kpiTeacher = document.getElementById("kpiTeacher");
  els.kpiRatio = document.getElementById("kpiRatio");

  els.insightList = document.getElementById("insightList");
  els.summaryTableBody = document.getElementById("summaryTableBody");
  els.loadingOverlay = document.getElementById("loadingOverlay");

  els.compareChartCanvas = document.getElementById("compareChart");
  els.ratioChartCanvas = document.getElementById("ratioChart");
}

function bindEvents() {
  els.openLevelSelect.addEventListener("change", updateDashboard);
}

async function init() {
  try {
    showLoading(true);
    const res = await fetch(`${API_URL}?type=summary&_=${Date.now()}`);
    const json = await res.json();

    if (json.error || json.ok === false) {
      throw new Error(json.error || "โหลดข้อมูลไม่สำเร็จ");
    }

    state.raw = dedupeRows(Array.isArray(json.data) ? json.data : []);
    buildOpenLevelFilter();
    updateDashboard();
  } catch (err) {
    console.error(err);
    alert(`โหลดข้อมูลไม่สำเร็จ: ${err.message}`);
  } finally {
    showLoading(false);
  }
}

function dedupeRows(rows) {
  const map = new Map();

  rows.forEach(row => {
    const year = String(row.year || "").trim();
    const level = String(row.level || "").trim();
    const openLevels = String(row.openLevels || "").trim();
    const key = openLevels || `${year}__${level}`;

    if (!year && !level && !openLevels) return;
    if (!map.has(key)) {
      map.set(key, {
        ...row,
        year,
        level,
        openLevels
      });
    }
  });

  return [...map.values()].sort(sortRows);
}

function buildOpenLevelFilter() {
  const options = state.raw
    .map(row => String(row.openLevels || "").trim())
    .filter(Boolean);

  els.openLevelSelect.innerHTML = options.map(value => `
    <option value="${escapeHtml(value)}">${escapeHtml(value)}</option>
  `).join("");
}

function updateDashboard() {
  const selectedOpenLevel = String(els.openLevelSelect.value || "").trim();
  state.selectedRow = state.raw.find(row => String(row.openLevels || "").trim() === selectedOpenLevel) || null;

  renderSummaryText(state.selectedRow);
  renderKpis(state.selectedRow);
  renderInsights(state.selectedRow);
  renderTable(state.selectedRow);
  renderCharts(state.selectedRow);
}

function renderSummaryText(row) {
  if (!row) {
    els.summaryText.textContent = "ไม่พบข้อมูล";
    return;
  }

  els.summaryText.textContent = `ปีการศึกษา: ${row.year} | ชั้นปี: ${row.level} | เปิดสอน: ${row.openLevels}`;
}

function renderKpis(row) {
  if (!row) {
    els.kpiTotalFTES.textContent = "-";
    els.kpiTeacher.textContent = "-";
    els.kpiRatio.textContent = "-";
    return;
  }

  els.kpiTotalFTES.textContent = formatNumber(row.ftes, 2);
  els.kpiTeacher.textContent = formatNumber(row.teacher, 0);
  els.kpiRatio.textContent = `1 : ${formatNumber(row.ratio, 2)}`;
}

function renderInsights(row) {
  if (!row) {
    els.insightList.innerHTML = `<li>ไม่พบข้อมูลสำหรับวิเคราะห์</li>`;
    return;
  }

  const insights = [
    `ข้อมูลชุดนี้เป็นปีการศึกษา ${row.year} ชั้นปี ${row.level} และเปิดสอน ${row.openLevels}.`,
    `มีค่า FTES เท่ากับ ${formatNumber(row.ftes, 2)} และมีอาจารย์ ${formatNumber(row.teacher, 0)} คน.`,
    `อัตราส่วนจากชีตคือ FTES/อาจารย์ = ${formatNumber(row.ratio, 2)}.`,
    `เมื่อแสดงเป็นรูปแบบอาจารย์ : FTES จะเป็น 1 : ${formatNumber(row.ratio, 2)}.`
  ];

  els.insightList.innerHTML = insights.map(item => `<li>${escapeHtml(item)}</li>`).join("");
}

function renderTable(row) {
  if (!row) {
    els.summaryTableBody.innerHTML = "";
    return;
  }

  els.summaryTableBody.innerHTML = `
    <tr>
      <td>${escapeHtml(row.year)}</td>
      <td>${escapeHtml(row.level)}</td>
      <td>${formatNumber(row.sch, 0)}</td>
      <td>${formatNumber(row.ftes, 2)}</td>
      <td>${formatNumber(row.teacher, 0)}</td>
      <td>${escapeHtml(row.openLevels)}</td>
      <td>${formatNumber(row.ratio, 2)}</td>
    </tr>
  `;
}

function renderCharts(row) {
  const labels = row ? [`${row.year} ${row.level}`] : [];
  const ftesData = row ? [toNumber(row.ftes)] : [];
  const teacherData = row ? [toNumber(row.teacher)] : [];
  const ratioData = row ? [toNumber(row.ratio)] : [];

  if (state.compareChart) state.compareChart.destroy();
  if (state.ratioChart) state.ratioChart.destroy();

  state.compareChart = new Chart(els.compareChartCanvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        { label: "FTES", data: ftesData },
        { label: "อาจารย์", data: teacherData }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });

  state.ratioChart = new Chart(els.ratioChartCanvas, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          label: "FTES/อาจารย์",
          data: ratioData
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

function sortRows(a, b) {
  const openLevelA = String(a.openLevels || "");
  const openLevelB = String(b.openLevels || "");
  return openLevelA.localeCompare(openLevelB, "th");
}

function toNumber(value) {
  const num = Number(value);
  return Number.isFinite(num) ? num : 0;
}

function formatNumber(value, digits = 2) {
  return Number(value || 0).toLocaleString("th-TH", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits
  });
}

function escapeHtml(str) {
  return String(str ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function showLoading(show) {
  els.loadingOverlay.classList.toggle("hidden", !show);
}
