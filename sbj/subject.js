const CONFIG = {
  WEB_APP_URL: "https://script.google.com/macros/s/AKfycbzMPZ7KcZT5LKxvzmCm3yKpsDtZ-BfiQrl9QjDrCM7jPM6TzO54fE0KasaKzrsAiPqIWA/exec"
};


const state = {
  rawData: [],
  filteredData: [],
  activeYearTab: "all"
};

const els = {
  filterAcademicYear: document.getElementById("filterAcademicYear"),
  filterSemester: document.getElementById("filterSemester"),
  filterYearLevel: document.getElementById("filterYearLevel"),
  sheetInfo: document.getElementById("sheetInfo"),
  lastUpdated: document.getElementById("lastUpdated"),

  orgStatusBadge: document.getElementById("orgStatusBadge"),
  orgStatusText: document.getElementById("orgStatusText"),
  statusSummaryInline: document.getElementById("statusSummaryInline"),

  kpiTotalCourses: document.getElementById("kpiTotalCourses"),
  kpiOffered: document.getElementById("kpiOffered"),
  kpiCompleted: document.getElementById("kpiCompleted"),
  kpiPending: document.getElementById("kpiPending"),
  kpiProgress: document.getElementById("kpiProgress"),
  kpiProgressTrend: document.getElementById("kpiProgressTrend"),

  criticalSummary: document.getElementById("criticalSummary"),
  watchList: document.getElementById("watchList"),
  analysisList: document.getElementById("analysisList"),

  yearChart: document.getElementById("yearChart"),
  progressChart: document.getElementById("progressChart"),

  courseTableBody: document.getElementById("courseTableBody"),
  yearTabs: document.getElementById("yearTabs")
};

document.addEventListener("DOMContentLoaded", init);

async function init() {
  try {
    els.orgStatusText.textContent = "กำลังโหลดข้อมูลจาก Apps Script...";

    const response = await fetch(`${CONFIG.WEB_APP_URL}?action=getCourses`, {
      method: "GET"
    });

    if (!response.ok) {
      throw new Error("ไม่สามารถเรียก Apps Script Web App ได้");
    }

    const result = await response.json();

    if (!result.ok) {
      throw new Error(result.message || "เกิดข้อผิดพลาดจาก backend");
    }

    state.rawData = Array.isArray(result.data) ? result.data : [];

    populateFilters(state.rawData);
    initYearTabs();
    applyFilters();

    els.lastUpdated.textContent = result.lastUpdated || formatThaiDateTime(new Date());
    els.sheetInfo.textContent = "Apps Script Web App";
  } catch (error) {
    console.error(error);
    renderEmptyState(error.message);
  }
}

function populateFilters(data) {
  const academicYears = uniqueSorted(data.map(item => item.academicYear));
  const semesters = uniqueSorted(data.map(item => item.semester));
  const years = uniqueSorted(data.map(item => item.yearLevel));

  fillSelect(els.filterAcademicYear, academicYears, "ทั้งหมด");
  fillSelect(els.filterSemester, semesters, "ทั้งหมด");
  fillSelect(els.filterYearLevel, years, "ทั้งหมด");

  els.filterAcademicYear.addEventListener("change", applyFilters);
  els.filterSemester.addEventListener("change", applyFilters);
  els.filterYearLevel.addEventListener("change", applyFilters);
}

function fillSelect(selectEl, values, allLabel) {
  selectEl.innerHTML = `<option value="all">${allLabel}</option>`;
  values.forEach(value => {
    selectEl.insertAdjacentHTML(
      "beforeend",
      `<option value="${escapeHtml(String(value))}">${escapeHtml(String(value))}</option>`
    );
  });
}

function initYearTabs() {
  if (!els.yearTabs) return;

  els.yearTabs.addEventListener("click", (event) => {
    const btn = event.target.closest(".year-tab");
    if (!btn) return;

    state.activeYearTab = btn.dataset.tab || "all";

    els.yearTabs.querySelectorAll(".year-tab").forEach(tab => {
      tab.classList.remove("active");
    });

    btn.classList.add("active");
    renderCourseTable(state.filteredData);
  });
}

function applyFilters() {
  const academicYear = els.filterAcademicYear.value;
  const semester = els.filterSemester.value;
  const yearLevel = els.filterYearLevel.value;

  state.filteredData = state.rawData.filter(item => {
    const matchAcademicYear = academicYear === "all" || item.academicYear === academicYear;
    const matchSemester = semester === "all" || String(item.semester) === semester;
    const matchYearLevel = yearLevel === "all" || String(item.yearLevel) === yearLevel;
    return matchAcademicYear && matchSemester && matchYearLevel;
  });

  renderDashboard(state.filteredData);
}

function renderDashboard(data) {
  renderExecutiveStatus(data);
  renderKpis(data);
  renderCriticalSummary(data);
  renderWatchList(data);
  renderAnalysis(data);
  renderYearChart(data);
  renderProgressChart(data);
  renderTabsCount(data);
  renderCourseTable(data);
}

function renderExecutiveStatus(data) {
  const offered = data.filter(item => item.isOffered).length;
  const completed = data.filter(item => item.isCompleted).length;
  const pending = data.filter(item => item.isOffered && !item.isCompleted).length;
  const notOpened = data.filter(item => !item.isOffered).length;
  const progress = offered > 0 ? (completed / offered) * 100 : 0;

  let badgeText = "สถานะ: ปกติ";
  let badgeClass = "org-status-good";
  let message = "ภาพรวมอยู่ในเกณฑ์ดี";
  let summary = `เปิดสอน ${offered} | เสร็จสิ้น ${completed} | ค้าง ${pending}`;

  if (notOpened > 0 && progress < 90) {
    badgeText = "สถานะ: มีความเสี่ยง";
    badgeClass = "org-status-risk";
    message = "มีรายวิชายังไม่เปิด และมีงานค้างที่ต้องเร่งติดตาม";
  } else if (pending > 0 || notOpened > 0 || progress < 100) {
    badgeText = "สถานะ: ต้องติดตาม";
    badgeClass = "org-status-watch";
    message = "ภาพรวมอยู่ในระดับดี แต่ยังมีประเด็นที่ต้องเฝ้าระวัง";
  }

  els.orgStatusBadge.className = `org-status-badge ${badgeClass}`;
  els.orgStatusBadge.textContent = badgeText;
  els.orgStatusText.textContent = message;
  els.statusSummaryInline.textContent = summary;
}

function renderKpis(data) {
  const total = data.length;
  const offered = data.filter(item => item.isOffered).length;
  const completed = data.filter(item => item.isCompleted).length;
  const pending = data.filter(item => item.isOffered && !item.isCompleted).length;
  const progress = offered > 0 ? (completed / offered) * 100 : 0;

  els.kpiTotalCourses.textContent = total.toLocaleString("th-TH");
  els.kpiOffered.textContent = offered.toLocaleString("th-TH");
  els.kpiCompleted.textContent = completed.toLocaleString("th-TH");
  els.kpiPending.textContent = pending.toLocaleString("th-TH");
  els.kpiProgress.textContent = `${progress.toFixed(2)}%`;
  els.kpiProgressTrend.textContent = progress >= 95 ? "▲" : progress >= 80 ? "●" : "▼";
}

function renderCriticalSummary(data) {
  const criticalCount = data.filter(item => !item.isOffered).length;
  const warningCount = data.filter(item => item.isOffered && !item.isCompleted).length;

  const pills = [];

  if (criticalCount > 0) {
    pills.push(`<span class="summary-pill summary-pill-critical">Critical ${criticalCount} รายการ</span>`);
  }

  if (warningCount > 0) {
    pills.push(`<span class="summary-pill summary-pill-warning">Warning ${warningCount} รายการ</span>`);
  }

  if (criticalCount === 0 && warningCount === 0) {
    pills.push(`<span class="summary-pill summary-pill-good">ไม่มีรายการต้องติดตาม</span>`);
  }

  els.criticalSummary.innerHTML = pills.join("");
}

function renderWatchList(data) {
  const criticalItems = data
    .filter(item => !item.isOffered)
    .map(item => ({ ...item, priority: 1 }));

  const warningItems = data
    .filter(item => item.isOffered && !item.isCompleted)
    .map(item => ({ ...item, priority: 2 }));

  const watchItems = [...criticalItems, ...warningItems]
    .sort((a, b) => a.priority - b.priority || a.yearLevel - b.yearLevel || String(a.courseCode).localeCompare(String(b.courseCode), "th"))
    .slice(0, 5);

  if (watchItems.length === 0) {
    els.watchList.innerHTML = `
      <div class="watch-item">
        <div class="watch-name">ไม่มีรายการที่ต้องติดตามเพิ่มเติม</div>
        <div class="watch-status">
          <span class="badge badge-success">ภาพรวมปกติ</span>
        </div>
      </div>
    `;
    return;
  }

  els.watchList.innerHTML = watchItems.map(item => {
    const isCritical = !item.isOffered;
    const mainBadge = isCritical
      ? `<span class="badge badge-critical">ยังไม่เปิดสอน</span>`
      : `<span class="badge badge-warning">เปิดสอนแล้ว แต่ยังไม่เสร็จ</span>`;

    const levelBadge = isCritical
      ? `<span class="badge badge-exec">Critical</span>`
      : `<span class="badge badge-exec">Warning</span>`;

    return `
      <div class="watch-item">
        <div class="watch-meta-top">
          <div>
            <div class="watch-code">${escapeHtml(item.courseCode)}</div>
            <div class="watch-name">${escapeHtml(item.courseName)}</div>
          </div>
          <div class="watch-year">ชั้นปี ${escapeHtml(String(item.yearLevel || "-"))}</div>
        </div>
        <div class="watch-status">
          ${mainBadge}
          ${levelBadge}
        </div>
      </div>
    `;
  }).join("");
}

function renderAnalysis(data) {
  const yearLevels = [1, 2, 3, 4];
  const insightLines = [];

  yearLevels.forEach(year => {
    const group = data.filter(item => item.yearLevel === year);
    const total = group.length;
    const offered = group.filter(item => item.isOffered).length;
    const completed = group.filter(item => item.isCompleted).length;
    const pending = group.filter(item => item.isOffered && !item.isCompleted).length;

    if (total === 0) {
      insightLines.push(`ชั้นปี ${year} ไม่มีข้อมูลในช่วงที่เลือก`);
      return;
    }

    if (offered === 0) {
      insightLines.push(`ชั้นปี ${year}: ยังไม่อยู่ในช่วงเปิดการเรียนการสอน`);
      return;
    }

    if (completed === offered) {
      insightLines.push(`ชั้นปี ${year}: ดำเนินการสอนแล้วเสร็จครบ 100% อยู่ในระดับปกติ`);
      return;
    }

    if (pending > 0) {
      insightLines.push(`ชั้นปี ${year}: มี ${pending} รายวิชาที่ยังดำเนินการไม่แล้วเสร็จ ควรติดตามใกล้ชิด`);
      return;
    }

    insightLines.push(`ชั้นปี ${year}: เปิดสอน ${offered} จาก ${total} รายวิชา`);
  });

  const offered = data.filter(item => item.isOffered).length;
  const completed = data.filter(item => item.isCompleted).length;
  const notOpened = data.filter(item => !item.isOffered).length;
  const progress = offered > 0 ? (completed / offered) * 100 : 0;

  if (notOpened > 0 && progress < 90) {
    insightLines.push(`ภาพรวม: ความก้าวหน้า ${progress.toFixed(2)}% อยู่ในระดับเสี่ยง และควรเร่งติดตามรายการที่ยังไม่เปิดสอน`);
  } else if (progress < 100) {
    insightLines.push(`ภาพรวม: ความก้าวหน้า ${progress.toFixed(2)}% อยู่ในระดับดี แต่ยังมีบางรายวิชาที่ต้องติดตาม`);
  } else {
    insightLines.push(`ภาพรวม: ความก้าวหน้า ${progress.toFixed(2)}% และไม่มีประเด็นคงค้าง`);
  }

  els.analysisList.innerHTML = insightLines.map(line => `<li>${escapeHtml(line)}</li>`).join("");
}

function renderYearChart(data) {
  const yearLevels = [1, 2, 3, 4];
  const maxValue = Math.max(
    1,
    ...yearLevels.flatMap(year => {
      const group = data.filter(item => item.yearLevel === year);
      return [
        group.length,
        group.filter(item => item.isOffered).length,
        group.filter(item => item.isCompleted).length
      ];
    })
  );

  els.yearChart.innerHTML = yearLevels.map(year => {
    const group = data.filter(item => item.yearLevel === year);
    const total = group.length;
    const offered = group.filter(item => item.isOffered).length;
    const completed = group.filter(item => item.isCompleted).length;

    const totalHeight = Math.max((total / maxValue) * 220, total ? 18 : 4);
    const offeredHeight = Math.max((offered / maxValue) * 220, offered ? 18 : 4);
    const completedHeight = Math.max((completed / maxValue) * 220, completed ? 18 : 4);

    return `
      <div class="year-col">
        <div class="year-bars">
          <div class="bar bar-total" style="height:${totalHeight}px" title="ทั้งหมด ${total}"></div>
          <div class="bar bar-offered" style="height:${offeredHeight}px" title="เปิดสอน ${offered}"></div>
          <div class="bar bar-completed" style="height:${completedHeight}px" title="สอนเสร็จ ${completed}"></div>
        </div>
        <div class="year-label">ชั้นปี ${year}</div>
      </div>
    `;
  }).join("");
}

function renderProgressChart(data) {
  const yearLevels = [1, 2, 3, 4];

  els.progressChart.innerHTML = yearLevels.map(year => {
    const group = data.filter(item => item.yearLevel === year);
    const offered = group.filter(item => item.isOffered).length;
    const completed = group.filter(item => item.isCompleted).length;
    const percent = offered > 0 ? (completed / offered) * 100 : 0;
    const color =
      percent >= 100 ? "var(--primary)" :
      percent >= 80 ? "var(--green)" :
      percent > 0 ? "var(--orange)" :
      "#9fb4d9";

    return `
      <div class="progress-row">
        <div class="progress-name">ปี ${year}</div>
        <div class="progress-track">
          <div class="progress-fill" style="width:${percent}%; background:${color};"></div>
        </div>
        <div class="progress-value">${percent.toFixed(0)}%</div>
      </div>
    `;
  }).join("");
}

function renderTabsCount(data) {
  if (!els.yearTabs) return;

  const counts = {
    all: data.length,
    1: data.filter(item => String(item.yearLevel) === "1").length,
    2: data.filter(item => String(item.yearLevel) === "2").length,
    3: data.filter(item => String(item.yearLevel) === "3").length,
    4: data.filter(item => String(item.yearLevel) === "4").length
  };

  els.yearTabs.querySelectorAll(".year-tab").forEach(tab => {
    const key = tab.dataset.tab;
    if (key === "all") {
      tab.textContent = `ทั้งหมด (${counts.all})`;
    } else {
      tab.textContent = `ชั้นปี ${key} (${counts[key] || 0})`;
    }
  });
}

function renderCourseTable(data) {
  let tableData = [...data];

  if (state.activeYearTab !== "all") {
    tableData = tableData.filter(
      item => String(item.yearLevel) === String(state.activeYearTab)
    );
  }

  if (tableData.length === 0) {
    els.courseTableBody.innerHTML = `
      <tr>
        <td colspan="9" style="text-align:center; color:#6b7d90; padding:24px;">
          ไม่พบข้อมูลรายวิชาในแท็บนี้
        </td>
      </tr>
    `;
    return;
  }

  els.courseTableBody.innerHTML = tableData.map(item => `
    <tr>
      <td>${escapeHtml(item.academicYear || "-")}</td>
      <td>${escapeHtml(String(item.semester || "-"))}</td>
      <td>${escapeHtml(item.courseCode)}</td>
      <td>${escapeHtml(item.courseName)}</td>
      <td>${escapeHtml(item.credits || "-")}</td>
      <td>${escapeHtml(String(item.yearLevel || "-"))}</td>
      <td>${booleanBadge(item.isOffered, "offered")}</td>
      <td>${booleanBadge(item.isCompleted, "completed")}</td>
      <td>${statusBadge(item)}</td>
    </tr>
  `).join("");
}

function renderEmptyState(message) {
  els.orgStatusBadge.className = "org-status-badge org-status-risk";
  els.orgStatusBadge.textContent = "สถานะ: โหลดไม่สำเร็จ";
  els.orgStatusText.textContent = message;
  els.statusSummaryInline.textContent = "-";

  els.kpiTotalCourses.textContent = "0";
  els.kpiOffered.textContent = "0";
  els.kpiCompleted.textContent = "0";
  els.kpiPending.textContent = "0";
  els.kpiProgress.textContent = "0%";
  els.kpiProgressTrend.textContent = "●";

  els.criticalSummary.innerHTML = `<span class="summary-pill summary-pill-critical">ไม่สามารถสรุปข้อมูลได้</span>`;
  els.watchList.innerHTML = `<div class="watch-item"><div class="watch-name">${escapeHtml(message)}</div></div>`;
  els.analysisList.innerHTML = `<li>${escapeHtml(message)}</li>`;
  els.yearChart.innerHTML = "";
  els.progressChart.innerHTML = "";
  els.courseTableBody.innerHTML = `
    <tr>
      <td colspan="9" style="text-align:center; color:#6b7d90; padding:24px;">
        ${escapeHtml(message)}
      </td>
    </tr>
  `;
}

function booleanBadge(value, type = "default") {
  if (type === "offered") {
    return value
      ? `<span class="badge badge-success">เปิดสอนแล้ว</span>`
      : `<span class="badge badge-muted">ยังไม่เปิดสอน</span>`;
  }

  if (type === "completed") {
    return value
      ? `<span class="badge badge-success">ดำเนินการเสร็จสิ้น</span>`
      : `<span class="badge badge-warning">อยู่ระหว่างดำเนินการ</span>`;
  }

  return value
    ? `<span class="badge badge-success">ใช่</span>`
    : `<span class="badge badge-muted">ไม่ใช่</span>`;
}

function statusBadge(item) {
  if (!item.isOffered && !item.isCompleted) {
    return `<span class="badge badge-critical">ยังไม่เปิดสอน</span>`;
  }
  if (item.isOffered && !item.isCompleted) {
    return `<span class="badge badge-warning">ค้างดำเนินการ</span>`;
  }
  return `<span class="badge badge-success">สอนเสร็จสิ้นแล้ว</span>`;
}

function uniqueSorted(arr) {
  return [...new Set(arr)]
    .filter(v => String(v).trim() !== "")
    .sort((a, b) => {
      if (!isNaN(a) && !isNaN(b)) return Number(a) - Number(b);
      return String(a).localeCompare(String(b), "th");
    });
}

function formatThaiDateTime(date) {
  return new Intl.DateTimeFormat("th-TH", {
    dateStyle: "long",
    timeStyle: "short"
  }).format(date);
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}