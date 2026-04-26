/* ==========================================================================
 * ไฟล์: code.gs  (Google Apps Script)
 * หน้าที่: เป็น Backend ทำหน้าที่อ่านข้อมูลจาก Google Sheets
 *          แล้วส่งกลับไปให้ฝั่งหน้าเว็บ (app.js) ในรูปแบบ JSON
 *
 * วิธีใช้งาน:
 *   1) เปิด Google Sheets ที่ใช้เก็บข้อมูล
 *   2) ไปที่ Extensions → Apps Script แล้ววางโค้ดนี้
 *   3) Deploy → New deployment → Web app  (เลือก "Anyone")
 *   4) เอา URL ที่ได้ ไปวางในตัวแปร API_URL ที่บรรทัดบนสุดของ app.js
 * ========================================================================== */


/* --- ส่วนตั้งค่า (ถ้าเปลี่ยนชีต/เปลี่ยนชื่อแท็บ ต้องแก้ตรงนี้) ----------------- */
const CONFIG = {
  // ID ของ Google Sheet (ดูได้จาก URL ของชีต ส่วนระหว่าง /d/...../edit)
  // ★ ถ้าจะใช้ชีตอื่น → แก้ค่านี้
  SPREADSHEET_ID: '1UDKqqeoXAtEVXQ8JHSOfTQkkULqjEwD1QeBJY0OIaDY',

  // ชื่อแท็บ (Sheet) ภายในไฟล์ที่เก็บข้อมูล
  // ★ ถ้าเปลี่ยนชื่อแท็บในไฟล์ Google Sheets → ต้องแก้ตรงนี้ให้ตรงกัน
  SHEET_NAME: 'RESEARCH_DATA'
};

/* --- doGet: จุดรับ request จากหน้าเว็บ ----------------------------------------
 * เมื่อหน้าเว็บ (app.js) ยิง fetch มาที่ URL ของ Web App
 * ฟังก์ชันนี้จะถูกเรียกอัตโนมัติ
 * รับพารามิเตอร์ ?action=xxx ผ่าน e.parameter
 * ★ ถ้าจะเพิ่ม endpoint ใหม่ (เช่น เพิ่มข้อมูล/แก้ข้อมูล) → เพิ่ม if ตรงนี้
 * -------------------------------------------------------------------------- */
function doGet(e) {
  const action = e.parameter.action || 'getResearchData';

  if (action === 'getResearchData') {
    return jsonOutput(getResearchData_());
  }

  return jsonOutput({
    success: false,
    message: 'Invalid action'
  });
}

/* --- getResearchData_: อ่านข้อมูลทุกแถวในชีต แล้วแปลงเป็น array ของ object -----
 * ขั้นตอน:
 *   1) เปิดไฟล์ชีตด้วย ID
 *   2) อ่านทุกค่า → แถวแรก = ชื่อหัวคอลัมน์ (headers)
 *   3) แถวที่เหลือ = ข้อมูลจริง → แปลงเป็น object ที่ key คือชื่อคอลัมน์
 *   4) ส่งผ่าน normalizeResearchRow_ เพื่อจัดชื่อฟิลด์ให้เป็นมาตรฐาน
 * -------------------------------------------------------------------------- */
function getResearchData_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    return {
      success: false,
      message: 'ไม่พบชีตชื่อ ' + CONFIG.SHEET_NAME,
      data: []
    };
  }

  // ดึงข้อมูลทุกเซลล์ในตาราง
  const values = sheet.getDataRange().getValues();

  // shift() = เอาแถวแรกออกมาเป็น headers (ชื่อคอลัมน์)
  const headers = values.shift().map(h => String(h).trim());

  const data = values
    // กรองแถวว่างทิ้ง
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        obj[h] = row[i];
      });
      return normalizeResearchRow_(obj);
    });

  return {
    success: true,
    count: data.length,
    data: data
  };
}

/* --- normalizeResearchRow_: จับคู่ชื่อคอลัมน์ในชีต ↔ ชื่อ key ที่ frontend ใช้ ---
 * เพราะหัวคอลัมน์ในชีตอาจเขียนด้วยภาษาไทยหรืออังกฤษ หรือสะกดต่างกัน
 * เราจึงใช้ || (or) เพื่อ "ลองชื่อหลายแบบ" จนกว่าจะเจอที่มีค่า
 *
 * ★ ถ้าหัวคอลัมน์ในชีตของคุณชื่อแปลกใหม่ → เพิ่มชื่อนั้นต่อท้ายในบรรทัดที่เกี่ยว
 *   ตัวอย่าง: ถ้าคอลัมน์ผู้วิจัยชื่อ "ผู้แต่ง" → เพิ่ม row['ผู้แต่ง'] เข้าไปบรรทัด NAME
 *
 * ⚠️ หมายเหตุ: บรรทัด AUTHOR_LEVEL ปรากฏ 2 ครั้ง (บรรทัดแรกถูกบรรทัดสองทับ)
 *   ไม่กระทบการทำงาน แต่ถ้าจะลบให้สะอาด ลบบรรทัดล่าง (ตำแหน่งซ้ำ) ออกได้
 * -------------------------------------------------------------------------- */
function normalizeResearchRow_(row) {
  return {
    // ปีที่ตีพิมพ์ (พ.ศ.)
    YEAR: row['ปี พ.ศ.'] || row.YEAR_BE || row.YEAR || '',
    // ชื่อผู้วิจัย
    NAME: row.NAME || row.FULL_NAME || row['ชื่อผู้วิจัย'] || '',
    // ชื่อเรื่อง/บทความ
    TITLE: row.TITLE || row['ชื่อเรื่อง'] || '',
    // วารสารที่ตีพิมพ์
    JOURNAL: row.JOURNAL || row['วารสาร'] || '',
    // ฐานข้อมูลที่ index (เช่น TCI1, Scopus, ISI)
    INDEX_STATUS: row.INDEX_STATUS || row['INDEX STATUS'] || row['ฐานข้อมูล'] || row.Level || '',
    // สถานะการตีพิมพ์ (ตีพิมพ์/โครงร่าง/อยู่ระหว่างดำเนินการ)
    PUBLISH_STATUS: row.PUBLISH_STATUS || row['ตีพิมพ์'] || row['สถานะ'] || '',
    // ประเภทงาน เช่น "วิจัย" "บทความวิชาการ"
    TYPE: row.TYPE || row['ประเภท'] || '',
    // ระดับผู้แต่ง (First / Co / Corresponding)
    AUTHOR_LEVEL: row.AUTHOR_LEVEL || row.Level || row['Level'] || '',
    // ผู้วิจัยหลัก (First Author) – ค่ามักเป็น Yes/No
    FIRST_AUTHOR: row.FIRST_AUTHOR || row['First Author'] || '',
    // ผู้รับผิดชอบบทความ (Corresponding Author)
    CORRESPONDING_AUTHOR: row.CORRESPONDING_AUTHOR || row['Corresponding Author'] || '',
    // ⚠️ บรรทัดซ้ำ (ทับบรรทัด AUTHOR_LEVEL ด้านบน) – ลบได้ไม่กระทบ
    AUTHOR_LEVEL: row.AUTHOR_LEVEL || row.Level || row['Level'] || '',
    // ทุนวิจัย: ทุน / แหล่งทุน / จำนวนเงิน
    FUNDING: row.FUNDING || row['ทุน'] || '',
    FUND_SOURCE: row.FUND_SOURCE || row['แหล่งทุน'] || '',
    FUND_AMOUNT: row.FUND_AMOUNT || row['จำนวนเงิน'] || '',
    // ลิงก์บทความ
    LINK: row.LINK || row['Link'] || row['URL'] || ''
  };
}


/* --- jsonOutput: helper แปลง object เป็น JSON response --------------------- */
function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}