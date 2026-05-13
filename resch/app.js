/* ============================================================================
 * ไฟล์: app.js
 * หน้าที่: เป็น Frontend ทั้งหมด — ดึงข้อมูลจาก Apps Script, กรอง, นับ, วาดกราฟ,
 *          แสดงตาราง และจัดการ event ทั้งหมดบนหน้าเว็บ
 *
 * ลำดับการทำงานเมื่อเปิดหน้า:
 *   1) DOMContentLoaded → init()
 *   2) fetch ข้อมูลจาก API_URL (Google Apps Script)
 *   3) initFilters() เติม dropdown ตัวเลือก
 *   4) bindEvents() ผูก event ทุกตัว (filter / search / pagination)
 *   5) renderDashboard() = renderKpis + renderCharts + renderTable
 *
 * ★ จุดที่มัก "ต้องแก้บ่อย" จะมีคอมเมนต์ ★ กำกับ
 * ============================================================================ */


/* --- URL ของ Web App (ได้จากการ Deploy code.gs เป็น Web App) ---------------
 * ★ ถ้า Deploy ใหม่ → ต้องอัปเดต URL ตรงนี้ ไม่งั้นหน้าเว็บจะโหลดข้อมูลไม่ได้
 * -------------------------------------------------------------------------- */
const API_URL = 'https://script.google.com/macros/s/AKfycbyP9dsRXysUbrbEVbhcZTiCc6Hf4ykpyFlYQ6qGxJ13q8OqZu9CUP_-aKfgV-ZvLW5J3g/exec';


/* --- ตัวแปรกลาง (state) ใช้ร่วมกันทั้งไฟล์ -------------------------------- */
let rawData = [];        // ข้อมูลดิบทั้งหมดจากชีต (ไม่เคยกรอง)
let filteredData = [];   // ข้อมูลที่ผ่านการกรองจากผู้ใช้แล้ว → ใช้แสดงผล
let charts = {};         // เก็บ instance ของ Chart.js เพื่อไว้ destroy ก่อนวาดใหม่
let currentPage = 1;     // หน้าปัจจุบันของตาราง
let pageSize = 10;       // จำนวนรายการต่อหน้า (ค่าเริ่มต้น)

/* ตัวกรองพิเศษจากการคลิกกราฟ (publish pie chart)
   ★ ถ้าจะเพิ่มตัวกรองจากการคลิกกราฟอื่น → เพิ่ม key ในออบเจ็กต์นี้ */
let chartFilter = {
  publishStatus: '',
  fundingStatus: ''
};
let activeChartSelection = '';

/* --- Plugin Chart.js: แสดงตัวเลขกำกับบนกราฟ --------------------------------
 * วาดตัวเลขเหนือแท่ง / บนชิ้นวงกลม หลัง Chart.js วาดเสร็จ
 * - กราฟแท่ง/เส้น: แสดงตัวเลขจำนวน
 * - กราฟวงกลม/โดนัท: แสดงจำนวน + เปอร์เซ็นต์
 *
 * ★ อยากเปลี่ยนสี/ขนาดตัวเลข → แก้ ctx.font / ctx.fillStyle ด้านล่าง
 * ★ อยากซ่อนตัวเลขในกราฟบางชนิด → เพิ่ม if (chart.config.type === 'xxx') return
 * -------------------------------------------------------------------------- */
const valueLabelPlugin = {
  id: 'valueLabelPlugin',
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    const isCircleChart =
      chart.config.type === 'pie' || chart.config.type === 'doughnut';
    const isStackedBarChart =
      chart.config.type === 'bar' &&
      (chart.options?.scales?.x?.stacked || chart.options?.scales?.y?.stacked);

    ctx.save();
    ctx.font = '600 12px IBM Plex Sans Thai';
    ctx.fillStyle = '#1f2d3d';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';

    chart.data.datasets.forEach((dataset, datasetIndex) => {
      if (!chart.isDatasetVisible(datasetIndex)) return;

      const meta = chart.getDatasetMeta(datasetIndex);
      const visibleTotal = isCircleChart
        ? dataset.data.reduce((sum, n, index) => {
            const arc = meta.data[index];
            if (!arc || arc.hidden) return sum;
            return sum + Number(n || 0);
          }, 0)
        : 0;

      meta.data.forEach((element, index) => {
        const value = dataset.data[index];

        if (!value || value === 0) return;

        if (isCircleChart) {
          const percent = visibleTotal
            ? ((value / visibleTotal) * 100).toFixed(1)
            : 0;

          // ชิ้นที่เล็กมากมักทำให้ label ชนกันเมื่อซ่อนบาง legend ไว้
          if (Number(percent) < 5) return;

          const position = element.tooltipPosition();
          ctx.fillText(`${value} (${percent}%)`, position.x, position.y);
          return;
        }

        if (isStackedBarChart) {
          const { x, y, base } = element.getProps(['x', 'y', 'base'], true);
          const centerY = y + ((base - y) / 2);
          ctx.fillText(value, x, centerY);
          return;
        }

        const position = element.tooltipPosition();
        ctx.fillText(value, position.x, position.y - 10);
      });
    });

    ctx.restore();
  }
};

// ลงทะเบียน plugin ให้ Chart.js รู้จัก
Chart.register(valueLabelPlugin);


/* --- รอ DOM โหลดเสร็จก่อน แล้วค่อยเรียก init() ----------------------------- */
document.addEventListener('DOMContentLoaded', init);


/* --- init: จุดเริ่มต้นของแอป -----------------------------------------------
 * 1) แสดงข้อความ "กำลังโหลด..."
 * 2) fetch JSON จาก API_URL
 * 3) ถ้าโหลดได้ → เก็บลง rawData แล้วเริ่ม render หน้าเว็บ
 * 4) ถ้าโหลดไม่ได้ → แสดงข้อความข้อผิดพลาดในตาราง
 * ★ ถ้าจะแก้ตอนเรียก API (เช่น เพิ่ม header, parameter) → แก้ที่ fetch()
 * -------------------------------------------------------------------------- */
async function init() {
  setHeaderLoading(true);
  showLoading();

  try {
    const res = await fetch(API_URL);
    const json = await res.json();

    if (!json.success) {
      throw new Error(json.message || 'โหลดข้อมูลไม่สำเร็จ');
    }

    rawData = json.data || [];
    filteredData = [...rawData];

    initFilters();
    bindEvents();
    renderDashboard();
    setHeaderLoading(false);

  } catch (error) {
    console.error(error);
    setHeaderLoading(false);

    document.getElementById('researchTableBody').innerHTML = `
      <tr>
        <td colspan="6" class="empty">
          โหลดข้อมูลไม่สำเร็จ กรุณาตรวจสอบ Apps Script Web App URL
        </td>
      </tr>
    `;
  }
}

/* --- showLoading: แสดงข้อความระหว่างโหลด ----------------------------------
 * แสดงข้อความระหว่างโหลดให้ครอบทั้ง 6 คอลัมน์ของตาราง
 * -------------------------------------------------------------------------- */
function showLoading() {
  document.getElementById('researchTableBody').innerHTML = `
    <tr>
      <td colspan="6" class="loading">กำลังโหลดข้อมูล...</td>
    </tr>
  `;
}

function setHeaderLoading(isLoading) {
  document.getElementById('headerLoadingText').hidden = !isLoading;
}


/* --- initFilters: เติมตัวเลือกใน dropdown 3 ตัวจากข้อมูลจริง ----------------
 * ★ ถ้าจะเพิ่มตัวกรองใหม่ → เพิ่ม fillSelect(...) อีก 1 บรรทัดตรงนี้
 *   และอย่าลืมเพิ่มเงื่อนไขใน applyFilters() ด้วย
 * -------------------------------------------------------------------------- */
function initFilters() {
  fillSelect('yearFilter', uniqueValues(rawData.map(r => r.YEAR)).sort());
  fillSelect('indexFilter', uniqueValues(rawData.map(r => r.INDEX_STATUS)).sort());
  fillSelect('nameFilter', uniqueValues(rawData.map(r => r.NAME)).sort());
  fillSelect('majorFilter', uniqueValues(rawData.map(r => r.MAJOR)).sort());
}

// helper: ใส่ <option> ลงใน <select> ตาม id ที่ระบุ
function fillSelect(id, values) {
  const select = document.getElementById(id);

  values.forEach(value => {
    if (!value) return;

    const option = document.createElement('option');
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}

// helper: ดึงค่าที่ "ไม่ซ้ำ" ออกจาก array (ใช้กับ dropdown)
function uniqueValues(arr) {
  return [...new Set(
    arr.map(v => String(v || '').trim()).filter(Boolean)
  )];
}


/* --- bindEvents: ผูก event ทั้งหมดของหน้าเว็บ ------------------------------
 * รวม: filter dropdown, ช่องค้นหา, ปุ่มรีเซ็ต, จำนวนต่อหน้า, ปุ่มเปลี่ยนหน้า
 *
 * ⚠️⚠️⚠️ บั๊กสำคัญ! ⚠️⚠️⚠️
 * บล็อก event ของ pageSizeSelect / prevPageBtn / nextPageBtn (ระหว่างเครื่องหมาย ★)
 * อยู่ "ภายใน" callback ของ resetBtn click event
 * → หมายความว่า event ของ pagination จะถูกผูกซ้ำทุกครั้งที่กด "รีเซ็ต"
 *   และตอนเปิดหน้าครั้งแรก (ยังไม่กดรีเซ็ต) ปุ่มเปลี่ยนหน้ายังไม่ทำงาน!
 *
 * ★ วิธีแก้: ย้ายบล็อก pageSizeSelect / prevPageBtn / nextPageBtn ออกจาก
 *   callback ของ resetBtn ให้อยู่ระดับเดียวกับ event อื่น ๆ (ระดับ bindEvents)
 * -------------------------------------------------------------------------- */
function bindEvents() {

  // 1) เมื่อเปลี่ยน dropdown ตัวกรอง → กรองใหม่
  ['yearFilter', 'indexFilter', 'nameFilter', 'majorFilter'].forEach(id => {
    document.getElementById(id).addEventListener('change', applyManualFilters);
  });

  // 2) เมื่อพิมพ์ค้นหา → กรองใหม่ทันที
  document.getElementById('searchInput')
    .addEventListener('input', applyManualFilters);

  // 3) ปุ่มรีเซ็ต: ล้างทุก filter แล้วกรองใหม่
  document.getElementById('resetBtn')
    .addEventListener('click', resetFilters);

  document.getElementById('toggleResetBtn')
    .addEventListener('click', toggleResetTools);

  document.getElementById('pageSizeSelect').addEventListener('change', e => {
    pageSize = e.target.value === 'all' ? 'all' : Number(e.target.value);
    currentPage = 1;
    renderTable();
  });

  document.getElementById('prevPageBtn').addEventListener('click', () => {
    if (currentPage > 1) {
      currentPage--;
      renderTable();
    }
  });

  document.getElementById('nextPageBtn').addEventListener('click', () => {
    const totalPages = getTotalPages();
    if (currentPage < totalPages) {
      currentPage++;
      renderTable();
    }
  });

  bindTableDragScroll();
}

function applyManualFilters() {
  activeChartSelection = '';
  applyFilters();
}

function resetFilters() {
  document.getElementById('yearFilter').value = '';
  document.getElementById('indexFilter').value = '';
  document.getElementById('nameFilter').value = '';
  document.getElementById('majorFilter').value = '';
  document.getElementById('searchInput').value = '';

  chartFilter.publishStatus = '';
  chartFilter.fundingStatus = '';
  activeChartSelection = '';

  applyFilters();
}

function bindTableDragScroll() {
  const tableWrap = document.querySelector('.table-wrap');
  if (!tableWrap) return;

  let isDragging = false;
  let didDrag = false;
  let startX = 0;
  let startY = 0;
  let scrollLeft = 0;
  let scrollTop = 0;

  tableWrap.addEventListener('pointerdown', event => {
    if (event.button !== undefined && event.button !== 0) return;
    if (event.target.closest('a, button, input, select, textarea')) return;

    isDragging = true;
    didDrag = false;
    startX = event.clientX;
    startY = event.clientY;
    scrollLeft = tableWrap.scrollLeft;
    scrollTop = tableWrap.scrollTop;

    tableWrap.classList.add('is-dragging');
    tableWrap.setPointerCapture?.(event.pointerId);
  });

  tableWrap.addEventListener('pointermove', event => {
    if (!isDragging) return;

    const deltaX = event.clientX - startX;
    const deltaY = event.clientY - startY;

    if (Math.abs(deltaX) > 4 || Math.abs(deltaY) > 4) {
      didDrag = true;
    }

    tableWrap.scrollLeft = scrollLeft - deltaX;
    tableWrap.scrollTop = scrollTop - deltaY;

    if (didDrag) {
      event.preventDefault();
    }
  });

  ['pointerup', 'pointercancel', 'pointerleave'].forEach(eventName => {
    tableWrap.addEventListener(eventName, event => {
      if (!isDragging) return;

      isDragging = false;
      tableWrap.classList.remove('is-dragging');
      tableWrap.releasePointerCapture?.(event.pointerId);
    });
  });

  tableWrap.addEventListener('click', event => {
    if (!didDrag) return;

    event.preventDefault();
    event.stopPropagation();
    didDrag = false;
  }, true);
}

function toggleResetTools() {
  const tools = document.getElementById('floatingResetTools');
  const toggleBtn = document.getElementById('toggleResetBtn');
  const isCollapsed = tools.classList.toggle('collapsed');

  toggleBtn.textContent = isCollapsed ? '+' : '−';
  toggleBtn.setAttribute('aria-expanded', String(!isCollapsed));
  toggleBtn.setAttribute(
    'aria-label',
    isCollapsed ? 'แสดงปุ่มรีเซ็ตตัวกรอง' : 'ย่อปุ่มรีเซ็ตตัวกรอง'
  );
}

/* --- applyFilters: คัดข้อมูลตามตัวกรองทั้งหมด แล้วสั่ง re-render -----------
 * รวมเงื่อนไข: ปี + ฐานข้อมูล + ผู้วิจัย + สถานะตีพิมพ์ (จากคลิกกราฟ) + คำค้น
 * ★ ถ้าจะเพิ่มตัวกรองใหม่ → เพิ่มทั้ง matchXxx และต่อท้าย return
 * -------------------------------------------------------------------------- */
function applyFilters() {
  const year = document.getElementById('yearFilter').value;
  const index = document.getElementById('indexFilter').value;
  const name = document.getElementById('nameFilter').value;
  const major = document.getElementById('majorFilter').value;
  const search = document.getElementById('searchInput')
    .value.toLowerCase().trim();

  filteredData = rawData.filter(row => {
    const matchYear = !year || String(row.YEAR) === year;
    const matchIndex = !index || String(row.INDEX_STATUS) === index;
    const matchName = !name || String(row.NAME) === name;
    const matchMajor = !major || String(row.MAJOR) === major;

    const publishText = normalizePublishStatus(row.PUBLISH_STATUS);
    const matchPublish =
      !chartFilter.publishStatus ||
      publishText === chartFilter.publishStatus;

    const fundingText = hasFunding(row) ? 'มีทุน' : 'ไม่มีทุน';
    const matchFunding =
      !chartFilter.fundingStatus ||
      fundingText === chartFilter.fundingStatus;

    const text = [
      row.YEAR,
      row.NAME,
      row.TITLE,
      row.JOURNAL,
      row.INDEX_STATUS,
      row.MAJOR,
      row.PUBLISH_STATUS,
      row.FUNDING,
      row.FUND_SOURCE,
      row.FUND_AMOUNT
    ].join(' ').toLowerCase();

    const matchSearch = !search || text.includes(search);

    return (
      matchYear &&
      matchIndex &&
      matchName &&
      matchMajor &&
      matchPublish &&
      matchFunding &&
      matchSearch
    );
  });

  renderDashboard();
}

/* --- renderDashboard: เรียกฟังก์ชัน render ทั้งหมดเรียงลำดับ ---------------- */
function renderDashboard() {
  renderKpis();
  renderCharts();
  renderTable();
}


/* --- renderKpis: นับและใส่ตัวเลขลงใน KPI Cards 4 ใบ -----------------------
 * นับเฉพาะแถวที่ TYPE === "วิจัย" (ตัดบทความวิชาการ ฯลฯ ออก)
 *   - total = จำนวนงานวิจัยทั้งหมด
 *   - firstAuthor = คนที่เป็น First Author
 *   - corresponding = Corresponding Author
 *   - olderThan5 = งานวิจัยอายุ > 5 ปี (สีแดง เตือน)
 *
 * ★ ถ้าหัวคอลัมน์ในชีตเปลี่ยน → แก้เงื่อนไขที่นี่ และอย่าลืมแก้
 *   normalizeResearchRow_() ใน code.gs ให้สอดคล้องด้วย
 * ★ ถ้าจะเปลี่ยนคำที่ถือว่าเป็น "ใช่" → แก้ที่ฟังก์ชัน isYes() ด้านล่าง
 * -------------------------------------------------------------------------- */
function renderKpis() {
  const researchOnly = filteredData.filter(
    r => String(r.TYPE || '').trim() === 'วิจัย'
  );

  const total = researchOnly.length;

  const firstAuthor = researchOnly.filter(r =>
    isYes(r.FIRST_AUTHOR) ||
    String(r.AUTHOR_LEVEL || '').toLowerCase().includes('first') ||
    String(r.AUTHOR_LEVEL || '').includes('ผู้วิจัย')
  ).length;

  const corresponding = researchOnly.filter(r =>
    isYes(r.CORRESPONDING_AUTHOR) ||
    String(r.AUTHOR_LEVEL || '').toLowerCase().includes('corresponding')
  ).length;

  const olderThan5 = researchOnly.filter(r =>
    String(r.MORE_THAN_5_YEARS || '').includes('มากกว่า') ||
    String(r.MORE_THAN_5_YEARS || '').toLowerCase() === 'true' ||
    String(r.MORE_THAN_5_YEARS || '') === '1'
  ).length;

  document.getElementById('totalResearch').textContent = total;
  document.getElementById('firstAuthor').textContent = firstAuthor;
  document.getElementById('correspondingAuthor').textContent = corresponding;
  document.getElementById('olderThan5').textContent = olderThan5;
}

/* helper: ตัดสินใจว่าค่าในเซลล์ "นับเป็น Yes" ไหม
   ★ ถ้าในชีตใช้คำอื่น เช่น "Y", "ถูกต้อง" → เพิ่มลงใน array ด้านล่าง */
function isYes(value) {
  const v = String(value || '').toLowerCase().trim();
  return ['true', 'yes', '1', 'ใช่', 'เป็น'].includes(v);
}


/* --- renderCharts: เรียกฟังก์ชันวาดกราฟทั้ง 4 ตัว -------------------------- */
function renderCharts() {
  renderYearLineChart();    // กราฟเส้น
  renderIndexBarChart();    // กราฟแท่งสะสม
  renderMajorBarChart();    // กราฟแท่งแยกสาขาวิชา
  renderIndexDonutChart();  // กราฟโดนัท
  renderPublishPieChart();  // กราฟพาย
  renderFundingAmountChart(); // กราฟแท่งจำนวนเงินทุน
  renderFundingStatusChart(); // กราฟโดนัท มีทุน/ไม่มีทุน
}

/* --- renderYearLineChart: กราฟเส้น แนวโน้มจำนวนงานวิจัยรายปี --------------
 * ★ ถ้าจะเปลี่ยนเป็นกราฟแท่ง → เปลี่ยน type: 'line' เป็น 'bar'
 * ★ tension = 0.35 → ทำให้เส้นโค้งนุ่ม (0 = หักมุม)
 * -------------------------------------------------------------------------- */
function renderYearLineChart() {
  const grouped = countBy(filteredData, 'YEAR');
  const labels = Object.keys(grouped).sort();
  const data = labels.map(label => grouped[label]);

  createOrUpdateChart('yearLineChart', {
    type: 'line',
    data: {
      labels,
      datasets: [{
        label: 'จำนวนงานวิจัย',
        data,
        tension: 0.35,
        fill: true
      }]
    }
  });
}

/* --- renderIndexBarChart: กราฟแท่งสะสม แยกฐานข้อมูล × ปี ------------------
 * stacked: true ทั้งแกน X และ Y → แท่งซ้อนกันแสดงสัดส่วนในแต่ละปี
 * ★ ถ้าจะให้แท่งเรียงข้างกัน (ไม่ซ้อน) → ตั้ง stacked: false ทั้ง 2 แกน
 * -------------------------------------------------------------------------- */
function renderIndexBarChart() {
  const years = uniqueValues(filteredData.map(r => r.YEAR)).sort();
  const indexes = uniqueValues(filteredData.map(r => r.INDEX_STATUS)).sort();

  const datasets = indexes.map(index => ({
    label: index || 'ไม่ระบุ',
    data: years.map(year =>
      filteredData.filter(r =>
        String(r.YEAR) === String(year) &&
        String(r.INDEX_STATUS || 'ไม่ระบุ') === String(index || 'ไม่ระบุ')
      ).length
    )
  }));

  createOrUpdateChart('indexBarChart', {
    type: 'bar',
    data: {
      labels: years,
      datasets
    },
    options: {
      scales: {
        x: {
          stacked: true,
          grid: {
            display: false
          }
        },
        y: {
          stacked: true,
          beginAtZero: true,
          ticks: {
            precision: 0
          }
        }
      }
    }
  });
}

/* --- renderIndexDonutChart: กราฟโดนัท สัดส่วนฐานข้อมูล --------------------
 * ★ ถ้าจะเปลี่ยนเป็นพาย (ทั้งวง) → เปลี่ยน type: 'doughnut' เป็น 'pie'
 * -------------------------------------------------------------------------- */
function renderIndexDonutChart() {
  const grouped = countBy(filteredData, 'INDEX_STATUS');
  const labels = Object.keys(grouped);
  const data = labels.map(label => grouped[label]);

  createOrUpdateChart('indexDonutChart', {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data }]
    }
  });
}

/* --- renderPublishPieChart: กราฟพาย สัดส่วนสถานะ (ตีพิมพ์/โครงร่าง) -------
 * นับด้วย countPublishStatus() ที่ผ่านการ "normalize" ค่ามาแล้ว
 * -------------------------------------------------------------------------- */
function renderPublishPieChart() {
  const grouped = countPublishStatus(filteredData);
  const labels = Object.keys(grouped);
  const data = labels.map(label => grouped[label]);

  createOrUpdateChart('publishPieChart', {
    type: 'pie',
    data: {
      labels,
      datasets: [{ data }]
    }
  });
}

/* --- renderMajorBarChart: กราฟแท่ง จำนวนงานวิจัยแยกตามสาขาวิชา ---------- */
function renderMajorBarChart() {
  const grouped = countBy(filteredData, 'MAJOR');
  const labels = Object.keys(grouped).sort((a, b) => grouped[b] - grouped[a]);
  const data = labels.map(label => grouped[label]);

  createOrUpdateChart('majorBarChart', {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'จำนวนงานวิจัย',
        data,
        backgroundColor: '#2a9d8f',
        borderColor: '#1f776d',
        borderWidth: 1
      }]
    },
    options: {
      indexAxis: 'y',
      scales: {
        x: {
          beginAtZero: true,
          ticks: {
            precision: 0
          }
        },
        y: {
          grid: {
            display: false
          }
        }
      }
    }
  });
}

function renderFundingAmountChart() {
  const years = uniqueValues(filteredData.map(r => r.YEAR)).sort();
  const data = years.map(year =>
    filteredData.reduce((sum, row) => {
      if (String(row.YEAR) !== String(year)) return sum;
      return sum + parseFundingAmount(row.FUND_AMOUNT);
    }, 0)
  );

  createOrUpdateChart('fundingAmountChart', {
    type: 'bar',
    data: {
      labels: years,
      datasets: [{
        label: 'จำนวนเงินทุน',
        data,
        backgroundColor: '#d4a017',
        borderColor: '#a87900',
        borderWidth: 1
      }]
    },
    options: {
      plugins: {
        tooltip: {
          callbacks: {
            label: function (context) {
              const value = Number(context.raw || 0);
              return `${context.dataset.label}: ${formatCurrency(value)}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: {
            display: false
          }
        },
        y: {
          beginAtZero: true,
          ticks: {
            callback: value => formatCurrencyShort(value)
          }
        }
      }
    }
  });
}

function renderFundingStatusChart() {
  const grouped = filteredData.reduce((acc, row) => {
    const label = hasFunding(row) ? 'มีทุน' : 'ไม่มีทุน';
    acc[label] = (acc[label] || 0) + 1;
    return acc;
  }, {});

  const labels = ['มีทุน', 'ไม่มีทุน'].filter(label => grouped[label] > 0);
  const data = labels.map(label => grouped[label]);

  createOrUpdateChart('fundingStatusChart', {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: ['#0d3b66', '#cfd8e3']
      }]
    }
  });
}

/* helper: นับจำนวนที่ซ้ำใน array → return object { ค่า: จำนวน } */
function countBy(data, key) {
  return data.reduce((acc, row) => {
    const value = String(row[key] || 'ไม่ระบุ').trim() || 'ไม่ระบุ';
    acc[value] = (acc[value] || 0) + 1;
    return acc;
  }, {});
}

/* helper เฉพาะของสถานะตีพิมพ์ – ใช้ normalizePublishStatus เพื่อรวมคำเขียนต่างกัน */
function countPublishStatus(data) {
  return data.reduce((acc, row) => {
    const status = normalizePublishStatus(row.PUBLISH_STATUS);
    acc[status] = (acc[status] || 0) + 1;
    return acc;
  }, {});
}

/* --- normalizePublishStatus: รวมคำเขียนหลากหลาย ให้เหลือแค่ 3 หมวด ---------
 * ในชีตอาจมีค่า: "ตีพิมพ์แล้ว", "Published 2566", "Draft", "Proposal" ฯลฯ
 * → ฟังก์ชันนี้ลดให้เหลือ "ตีพิมพ์" / "โครงร่าง" / "ไม่ระบุ"
 *
 * ★ ถ้าจะเพิ่มสถานะใหม่ (เช่น "อยู่ระหว่างดำเนินการ") → เพิ่ม if อีก 1 บล็อก
 * -------------------------------------------------------------------------- */
function normalizePublishStatus(value) {
  const v = String(value || '').toLowerCase();

  if (v.includes('ตีพิมพ์') || v.includes('published')) {
    return 'ตีพิมพ์';
  }

  if (
    v.includes('โครงร่าง') ||
    v.includes('draft') ||
    v.includes('proposal')
  ) {
    return 'โครงร่าง';
  }

  return 'ไม่ระบุ';
}

function parseFundingAmount(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  const cleaned = String(value || '')
    .replace(/[,฿\s]/g, '')
    .replace(/[^\d.-]/g, '');

  const amount = Number(cleaned);
  return Number.isFinite(amount) ? amount : 0;
}

function hasFunding(row) {
  const funding = String(row.FUNDING || '').trim();
  const source = String(row.FUND_SOURCE || '').trim();
  const amount = parseFundingAmount(row.FUND_AMOUNT);

  const hasFundingText = funding && !isNoFundingValue(funding);
  const hasFundingSource = source && !isNoFundingValue(source);

  return Boolean(hasFundingText || hasFundingSource || amount > 0);
}

function isNoFundingValue(value) {
  const normalized = String(value || '').trim().toLowerCase();

  return [
    '',
    '-',
    'ไม่มี',
    'ไม่มีทุน',
    'ไม่รับทุน',
    'ไม่ได้รับทุน',
    'no',
    'none',
    'n/a',
    'na'
  ].includes(normalized);
}

function formatCurrency(value) {
  return new Intl.NumberFormat('th-TH', {
    style: 'currency',
    currency: 'THB',
    maximumFractionDigits: 2
  }).format(Number(value || 0));
}

function formatCurrencyShort(value) {
  const amount = Number(value || 0);

  if (Math.abs(amount) >= 1000000) {
    return `${(amount / 1000000).toFixed(1)} ลบ.`;
  }

  if (Math.abs(amount) >= 1000) {
    return `${(amount / 1000).toFixed(0)}k`;
  }

  return `${amount}`;
}

/* --- createOrUpdateChart: ฟังก์ชันกลางสำหรับสร้าง/อัปเดต Chart ทุกตัว ------
 * รวม "ค่า default" ของกราฟทั้งหมด (legend, tooltip, animation, click)
 * แล้ว merge กับ config ที่ส่งเข้ามา
 *
 * ★ อยากเปลี่ยนสไตล์ tooltip / ฟอนต์ legend / สีพื้น tooltip → แก้ตรง
 *   defaultOptions.plugins.tooltip / .legend
 * ★ อยากปิดการกดที่กราฟแล้วกรอง → ลบ onClick: ออก
 * -------------------------------------------------------------------------- */
function createOrUpdateChart(canvasId, config) {
  const ctx = document.getElementById(canvasId);
  const isCircleChart =
    config.type === 'pie' || config.type === 'doughnut';

  const defaultOptions = {
    responsive: true,
    maintainAspectRatio: false,

    interaction: {
      mode: 'nearest',
      intersect: false
    },

    plugins: {
      legend: {
        position: 'bottom',
        labels: {
          usePointStyle: true,
          padding: 16,
          font: {
            family: 'IBM Plex Sans Thai',
            size: 12
          }
        }
      },

      tooltip: {
        enabled: true,
        backgroundColor: '#0d3b66',
        padding: 12,
        cornerRadius: 12,
        callbacks: {
          label: function (context) {
            const label =
              context.dataset.label || context.label || '';
            const value = context.raw || 0;

            if (isCircleChart) {
              const data = context.dataset.data || [];
              const total = data.reduce(
                (sum, n) => sum + Number(n || 0),
                0
              );
              const percent = total
                ? ((value / total) * 100).toFixed(1)
                : 0;

              return `${label}: ${value} รายการ (${percent}%)`;
            }

            return `${label}: ${value} รายการ`;
          }
        }
      }
    },

    scales: isCircleChart
      ? undefined
      : {
          x: {
            ticks: {
              font: {
                family: 'IBM Plex Sans Thai'
              }
            },
            grid: {
              display: false
            }
          },
          y: {
            beginAtZero: true,
            ticks: {
              precision: 0,
              font: {
                family: 'IBM Plex Sans Thai'
              }
            }
          }
        },

    animation: {
      duration: 700,
      easing: 'easeOutQuart'
    },

    onHover: (event, chartElement) => {
      event.native.target.style.cursor =
        chartElement.length ? 'pointer' : 'default';
    },

    onClick: (event, elements, chart) => {
      if (!elements.length) return;

      const clicked = elements[0];
      const label = chart.data.labels[clicked.index];
      const datasetLabel =
        chart.data.datasets[clicked.datasetIndex]?.label || '';

      handleChartClick(canvasId, label, datasetLabel);
    }
  };

  // ทำลายกราฟตัวเก่าก่อน (ป้องกัน Chart.js error: canvas already in use)
  if (charts[canvasId]) {
    charts[canvasId].destroy();
  }

  // สร้างกราฟใหม่ และเก็บ instance ไว้ในตัวแปรกลาง charts[]
  charts[canvasId] = new Chart(ctx, {
    ...config,
    options: {
      ...defaultOptions,
      ...(config.options || {}),
      plugins: {
        ...defaultOptions.plugins,
        ...(config.options?.plugins || {})
      }
    }
  });
}

/* --- handleChartClick: เมื่อคลิกที่กราฟ ให้ตั้ง filter อัตโนมัติ ----------
 * - คลิก "yearLineChart"  → ตั้ง filter ปี
 * - คลิก "indexBarChart"  → ตั้งทั้งปีและฐานข้อมูล
 * - คลิก "indexDonutChart"→ ตั้ง filter ฐานข้อมูล
 * - คลิก "publishPieChart"→ ตั้ง chartFilter.publishStatus
 * แล้วเรียก applyFilters() ให้กรองและ re-render ทันที
 * ★ ถ้าจะเพิ่มกราฟใหม่ที่คลิกแล้วกรองได้ → เพิ่ม if อีก 1 บล็อก
 * -------------------------------------------------------------------------- */
function handleChartClick(canvasId, label, datasetLabel) {
  const selectionKey = [
    canvasId,
    String(label || ''),
    String(datasetLabel || '')
  ].join('|');

  if (activeChartSelection === selectionKey) {
    resetFilters();
    return;
  }

  activeChartSelection = selectionKey;

  if (canvasId === 'yearLineChart') {
    document.getElementById('yearFilter').value = label;
    chartFilter.publishStatus = '';
  }

  if (canvasId === 'indexBarChart') {
    document.getElementById('yearFilter').value = label;

    if (datasetLabel) {
      document.getElementById('indexFilter').value = datasetLabel;
    }

    chartFilter.publishStatus = '';
  }

  if (canvasId === 'indexDonutChart') {
    document.getElementById('indexFilter').value = label;
    chartFilter.publishStatus = '';
  }

  if (canvasId === 'majorBarChart') {
    document.getElementById('majorFilter').value = label;
    chartFilter.publishStatus = '';
  }

  if (canvasId === 'publishPieChart') {
    chartFilter.publishStatus = label;
  }

  if (canvasId === 'fundingAmountChart') {
    document.getElementById('yearFilter').value = label;
    chartFilter.publishStatus = '';
  }

  if (canvasId === 'fundingStatusChart') {
    chartFilter.fundingStatus = label;
  }

  applyFilters();
}

/* --- renderTable: วาดตารางรายละเอียดงานวิจัย พร้อม pagination -------------
 * ขั้นตอน:
 *   1) คำนวณจำนวนหน้าทั้งหมด
 *   2) ตัดข้อมูลเฉพาะหน้าปัจจุบัน
 *   3) แปลงเป็น HTML แถวละ <tr>
 *   4) ใส่จำนวนรวม / หน้าปัจจุบัน / เปิด-ปิดปุ่ม prev/next
 *
 * ⚠️ "ไม่พบข้อมูล" ใช้ colspan="6" — ถูกต้องแล้ว (ตารางมี 6 คอลัมน์)
 * ★ ถ้าจะเพิ่ม/ลบคอลัมน์ → ต้องแก้:
 *     - <thead> ใน index.html
 *     - tbody.innerHTML = displayData.map(...) ตรงนี้
 *     - colspan ของแถว "ไม่พบข้อมูล" / "loading"
 *     - ความกว้างคอลัมน์ใน style.css (th:nth-child)
 * -------------------------------------------------------------------------- */
function renderTable() {
  const tbody = document.getElementById('researchTableBody');

  const totalItems = filteredData.length;
  const totalPages = getTotalPages();

  if (currentPage > totalPages) currentPage = totalPages;

  const displayData = getPagedData();

  document.getElementById('tableCount').textContent =
    `${totalItems} รายการ`;

  document.getElementById('pageInfo').textContent =
    `หน้า ${currentPage} / ${totalPages}`;

  document.getElementById('prevPageBtn').disabled = currentPage <= 1;
  document.getElementById('nextPageBtn').disabled = currentPage >= totalPages;

  if (!displayData.length) {
    tbody.innerHTML = `
      <tr>
        <td colspan="6" class="empty">ไม่พบข้อมูล</td>
      </tr>
    `;
    return;
  }

  tbody.innerHTML = displayData.map(row => `
    <tr>
      <td>${escapeHtml(row.YEAR)}</td>
      <td>
        <div class="researcher-name">${escapeHtml(row.NAME)}</div>
        <div class="researcher-role">${escapeHtml(row.AUTHOR_LEVEL || '-')}</div>
      </td>
      <td class="title-cell" title="${escapeHtml(row.TITLE)}">
        <span class="title-text">${escapeHtml(row.TITLE)}</span>
        ${renderInlineLink(row.LINK)}
      </td>
      <td>${escapeHtml(row.JOURNAL)}</td>
      <td>${escapeHtml(row.INDEX_STATUS)}</td>
      <td>${escapeHtml(row.FUNDING || row.FUND_SOURCE || row.FUND_AMOUNT || '-')}</td>
    </tr>
  `).join('');
}

function renderInlineLink(link) {
  if (!link || link === '-') return '';

  return `
    <a class="inline-link-icon" href="${escapeHtml(link)}" target="_blank" rel="noopener" aria-label="เปิดลิงก์งานวิจัยในแท็บใหม่" title="เปิดลิงก์งานวิจัยในแท็บใหม่">
      ↗
    </a>
  `;
}

/* helper: ป้องกัน XSS – แปลงอักขระพิเศษใน HTML ให้ปลอดภัยก่อนแสดง
   ★ อย่าลบฟังก์ชันนี้! เพราะข้อความในชีตอาจมี <, >, " แล้วทำให้หน้าเสีย */
function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}
/* --- getTotalPages: คำนวณจำนวนหน้าทั้งหมดของตาราง ------------------------- */
function getTotalPages() {
  if (pageSize === 'all') return 1;
  return Math.max(1, Math.ceil(filteredData.length / pageSize));
}

/* --- getPagedData: ตัดข้อมูลเฉพาะหน้าปัจจุบัน
 *   เช่น หน้า 2, pageSize 20 → return รายการที่ index 20 ถึง 39
 * -------------------------------------------------------------------------- */
function getPagedData() {
  const sortedData = getSortedTableData();

  if (pageSize === 'all') return sortedData;

  const start = (currentPage - 1) * pageSize;
  const end = start + pageSize;

  return sortedData.slice(start, end);
}

function getSortedTableData() {
  return filteredData
    .map((row, index) => ({ row, index }))
    .sort((a, b) => {
      const yearA = getSortableYear(a.row.YEAR);
      const yearB = getSortableYear(b.row.YEAR);

      if (yearA !== yearB) return yearB - yearA;

      return a.index - b.index;
    })
    .map(item => item.row);
}

function getSortableYear(value) {
  const match = String(value || '').match(/\d+/);
  const year = match ? Number(match[0]) : Number.NEGATIVE_INFINITY;

  return Number.isFinite(year) ? year : Number.NEGATIVE_INFINITY;
}
