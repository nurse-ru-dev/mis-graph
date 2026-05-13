const SHEET_ID = "19uDHafBByN5jRyJUstDF0GgXLAXZDol_Uj094hJY00c";
const SHEET_CAL = "FTES_CAL";
const CACHE_SECONDS = 300;

function doGet(e) {
  try {
    const cache = CacheService.getScriptCache();
    const cacheKey = "FTES_summary_all";
    const cached = cache.get(cacheKey);

    if (cached) {
      return outputJson(JSON.parse(cached));
    }

    const result = getSummaryData_();
    cache.put(cacheKey, JSON.stringify(result), CACHE_SECONDS);
    return outputJson(result);

  } catch (err) {
    return outputJson({
      ok: false,
      error: err.message || String(err)
    });
  }
}

function getSummaryData_() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_CAL);
  if (!sh) throw new Error("ไม่พบชีท FTES_CAL");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) {
    return { ok: true, data: [] };
  }

  const headers = values[0].map(h => normalizeHeader_(h));

  const idxYear = findIndex_(headers, ["ปีการศึกษา"]);
  const idxLevel = findIndex_(headers, ["ชั้นปี"]);
  const idxSCH = findIndex_(headers, ["SCH รวม", "SCH"]);
  const idxFTES = findIndex_(headers, ["FTES"]);
  const idxTeacher = findIndex_(headers, ["อาจารย์"]);
  const idxOpenLevels = findIndex_(headers, ["จำนวนชั้นปีที่เปิดสอน"]);
  const idxRatio = findIndex_(headers, ["FTES/อาจารย์", "FTES ต่ออาจารย์"]);

  const data = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const year = safeValue_(row, idxYear);
    const level = safeValue_(row, idxLevel);

    if (!year && !level) continue;

    data.push({
      year: String(year || "").trim(),
      level: String(level || "").trim(),
      sch: toNumber_(safeValue_(row, idxSCH)),
      ftes: toNumber_(safeValue_(row, idxFTES)),
      teacher: toNumber_(safeValue_(row, idxTeacher)),
      openLevels: String(safeValue_(row, idxOpenLevels) || "").trim(),
      ratio: toNumber_(safeValue_(row, idxRatio))
    });
  }

  return {
    ok: true,
    data: data
  };
}

function findIndex_(headers, names) {
  for (var i = 0; i < names.length; i++) {
    var idx = headers.indexOf(normalizeHeader_(names[i]));
    if (idx !== -1) return idx;
  }
  return -1;
}

function normalizeHeader_(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .replace(/\u00A0/g, " ")
    .trim();
}

function safeValue_(row, idx) {
  if (idx < 0 || idx >= row.length) return "";
  return row[idx];
}

function toNumber_(value) {
  if (value === null || value === undefined || value === "") return 0;
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/,/g, "").trim();
  const num = Number(cleaned);
  return isNaN(num) ? 0 : num;
}

function outputJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
