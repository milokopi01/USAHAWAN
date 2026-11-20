/**
 * =========================================================================
 * KH PACKAGING — KEUSAHAWANAN PRO (Google Apps Script Backend)
 * =========================================================================
 * Panduan Pemasangan:
 * 1. Muat naik fail DATABASE_KEUSAHAWANAN_PRO_2024_2025.xlsx ke Google Drive.
 * 2. Buka fail tersebut dengan Google Sheets -> File -> Save as Google Sheets.
 * 3. Pergi ke menu Extensions -> Apps Script.
 * 4. Padam semua kod sedia ada, tampal (paste) keseluruhan kod ini & tekan Save.
 * 5. Klik Deploy -> New deployment -> Type: Web app.
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Salin Web App URL dan tampal ke dalam menu Tetapan aplikasi PWA anda.
 * =========================================================================
 */

const SHEET_DAILY   = "Ringkasan_Harian";
const SHEET_ITEMS   = "Rekod_Jualan";
const SHEET_CATALOG = "Katalog_Produk";
const SHEET_SET     = "Tetapan_App";

const MALAY_MONTHS = [
  "JANUARI", "FEBRUARI", "MAC", "APRIL", "MEI", "JUN",
  "JULAI", "OGOS", "SEPTEMBER", "OKTOBER", "NOVEMBER", "DISEMBER"
];
const MALAY_DAYS = ["Ahad", "Isnin", "Selasa", "Rabu", "Khamis", "Jumaat", "Sabtu"];

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "ping";
    if (action === "ping") {
      return jsonResponse({ ok: true, message: "Sambungan Google Sheet Berjaya!", app: getSettings().NAMA_PERNIAGAAN || "KH KEUSAHAWANAN PRO" });
    }
    if (action === "getBootstrap") {
      return jsonResponse({ ok: true, data: getBootstrapData() });
    }
    if (action === "getDailyDetail") {
      const tarikh = e.parameter.tarikh || "";
      return jsonResponse({ ok: true, data: getDailyDetail(tarikh) });
    }
    return jsonResponse({ ok: false, error: "Tindakan GET tidak dikenali: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function doPost(e) {
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    const payload = JSON.parse(raw);
    const action = payload.action || "";

    if (action === "saveDailySales") {
      return jsonResponse({ ok: true, data: saveDailySales(payload) });
    }
    if (action === "deleteDailyRecord") {
      return jsonResponse({ ok: true, data: deleteDailyRecord(payload) });
    }
    if (action === "saveProduct") {
      return jsonResponse({ ok: true, data: saveProduct(payload) });
    }
    if (action === "deleteProduct") {
      return jsonResponse({ ok: true, data: deleteProduct(payload) });
    }
    return jsonResponse({ ok: false, error: "Tindakan POST tidak dikenali: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatDateYMD(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const str = String(val).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) return str.substring(0, 10);
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return str;
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_SET);
  const res = {
    NAMA_PERNIAGAAN: "KH PACKAGING — KEUSAHAWANAN PRO",
    MATA_WANG: "RM",
    HAD_AMARAN_VARIANS: "5.00",
    PIN_ADMIN: "1234"
  };
  if (!sh) return res;
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0] || "").trim();
    const v = String(rows[i][1] || "").trim();
    if (k) res[k] = v;
  }
  return res;
}

function getCatalog() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CATALOG);
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  const list = [];
  for (let i = 1; i < rows.length; i++) {
    const nama = String(rows[i][1] || "").trim();
    if (!nama) continue;
    list.push({
      kod: String(rows[i][0] || ("PRD-" + i)).trim(),
      kategori: nama,
      harga: Number(rows[i][2]) || 0,
      ikon: String(rows[i][3] || "📦").trim(),
      status: String(rows[i][4] || "AKTIF").trim().toUpperCase()
    });
  }
  return list;
}

function getBootstrapData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  const catalog = getCatalog();

  // Ringkasan Harian
  const shDaily = ss.getSheetByName(SHEET_DAILY);
  const dailyList = [];
  if (shDaily && shDaily.getLastRow() > 1) {
    const rows = shDaily.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const tarikh = formatDateYMD(rows[i][1]);
      if (!tarikh) continue;
      dailyList.push({
        id: String(rows[i][0] || ""),
        tarikh: tarikh,
        bulan: String(rows[i][2] || ""),
        tahun: Number(rows[i][3]) || Number(tarikh.substring(0, 4)),
        hari: String(rows[i][4] || ""),
        bilItem: Number(rows[i][5]) || 0,
        jumlahUnit: Number(rows[i][6]) || 0,
        jualanDijangka: Number(rows[i][7]) || 0,
        jualanSebenar: Number(rows[i][8]) || 0,
        varians: Number(rows[i][9]) || 0,
        status: String(rows[i][10] || ""),
        catatan: String(rows[i][11] || "")
      });
    }
  }
  // Sort newest first
  dailyList.sort((a, b) => b.tarikh.localeCompare(a.tarikh));

  // Rekod Jualan (for Category Analytics)
  const shItems = ss.getSheetByName(SHEET_ITEMS);
  const catMap = {};
  if (shItems && shItems.getLastRow() > 1) {
    const rows = shItems.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const kat = String(rows[i][4] || "").trim();
      if (!kat) continue;
      const unit = Number(rows[i][9]) || 0;
      const rm = Number(rows[i][11]) || 0;
      if (!catMap[kat]) catMap[kat] = { kategori: kat, unit: 0, jualan: 0, rekod: 0 };
      catMap[kat].unit += unit;
      catMap[kat].jualan += rm;
      catMap[kat].rekod += 1;
    }
  }
  const categoryStats = Object.values(catMap).sort((a, b) => b.jualan - a.jualan);

  return {
    settings: {
      NAMA_PERNIAGAAN: settings.NAMA_PERNIAGAAN,
      MATA_WANG: settings.MATA_WANG,
      HAD_AMARAN_VARIANS: Number(settings.HAD_AMARAN_VARIANS) || 5
    },
    catalog: catalog,
    dailyList: dailyList,
    categoryStats: categoryStats
  };
}

function getDailyDetail(tarikh) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shItems = ss.getSheetByName(SHEET_ITEMS);
  const shDaily = ss.getSheetByName(SHEET_DAILY);
  const items = [];
  let summary = null;

  if (shItems && shItems.getLastRow() > 1) {
    const rows = shItems.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (formatDateYMD(rows[i][1]) === tarikh) {
        items.push({
          kategori: String(rows[i][4] || ""),
          stokAwal: Number(rows[i][5]) || 0,
          stokTambah: Number(rows[i][6]) || 0,
          jumlahStok: Number(rows[i][7]) || 0,
          bakiStok: Number(rows[i][8]) || 0,
          unitTerjual: Number(rows[i][9]) || 0,
          hargaSeunit: Number(rows[i][10]) || 0,
          jualanDijangka: Number(rows[i][11]) || 0
        });
      }
    }
  }

  if (shDaily && shDaily.getLastRow() > 1) {
    const rows = shDaily.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (formatDateYMD(rows[i][1]) === tarikh) {
        summary = {
          id: String(rows[i][0] || ""),
          tarikh: tarikh,
          jualanDijangka: Number(rows[i][7]) || 0,
          jualanSebenar: Number(rows[i][8]) || 0,
          varians: Number(rows[i][9]) || 0,
          status: String(rows[i][10] || ""),
          catatan: String(rows[i][11] || "")
        };
        break;
      }
    }
  }

  return { tarikh: tarikh, summary: summary, items: items };
}

function saveDailySales(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shDaily = ss.getSheetByName(SHEET_DAILY);
    const shItems = ss.getSheetByName(SHEET_ITEMS);
    if (!shDaily || !shItems) throw new Error("Tab Ringkasan_Harian atau Rekod_Jualan tidak dijumpai!");

    const tarikh = String(payload.tarikh || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(tarikh)) throw new Error("Format tarikh tidak sah (YYYY-MM-DD).");

    const items = Array.isArray(payload.items) ? payload.items.filter(x => (Number(x.stokAwal) || 0) > 0 || (Number(x.stokTambah) || 0) > 0) : [];
    if (items.length === 0) throw new Error("Sila masukkan sekurang-kurangnya satu produk dengan stok.");

    const jualanSebenar = Number(payload.jualanSebenar) || 0;
    const catatan = String(payload.catatan || "").trim();

    // Parse Malay month/day
    const parts = tarikh.split("-").map(Number);
    const dtObj = new Date(parts[0], parts[1] - 1, parts[2]);
    const bulanStr = MALAY_MONTHS[dtObj.getMonth()] + " " + parts[0];
    const hariStr = MALAY_DAYS[dtObj.getDay()];

    // Remove existing rows for this date (Update/Overwrite mode cleanly)
    removeRowsByDate(shItems, 2, tarikh);
    removeRowsByDate(shDaily, 2, tarikh);

    // Append item rows with live formulas
    let totalExpected = 0;
    let totalUnit = 0;
    let nextItemRow = shItems.getLastRow() + 1;

    items.forEach((it, idx) => {
      const sAwal = Number(it.stokAwal) || 0;
      const sTambah = Number(it.stokTambah) || 0;
      const baki = Number(it.bakiStok) || 0;
      const harga = Number(it.hargaSeunit) || 0;
      const terjual = (sAwal + sTambah) - baki;
      const exp = terjual * harga;
      totalUnit += terjual;
      totalExpected += exp;

      const r = nextItemRow + idx;
      const idRek = "RK-" + tarikh.replace(/-/g, "") + "-" + (idx + 1);
      shItems.appendRow([
        idRek,
        tarikh,
        bulanStr,
        parts[0],
        String(it.kategori || "").toUpperCase(),
        sAwal,
        sTambah,
        `=F${r}+G${r}`,
        baki,
        `=H${r}-I${r}`,
        harga,
        `=J${r}*K${r}`,
        "PWA APP"
      ]);
    });

    // Append daily summary row with live formulas
    const dr = shDaily.getLastRow() + 1;
    const idHari = "HR-" + tarikh.replace(/-/g, "");
    shDaily.appendRow([
      idHari,
      tarikh,
      bulanStr,
      parts[0],
      hariStr,
      items.length,
      totalUnit,
      Number(totalExpected.toFixed(2)),
      Number(jualanSebenar.toFixed(2)),
      `=I${dr}-H${dr}`,
      `=IF(ROUND(J${dr},2)=0,"SEIMBANG",IF(J${dr}>0,"LEBIHAN (+)","KURANGAN (-)"))`,
      catatan,
      "PWA APP"
    ]);

    SpreadsheetApp.flush();
    return {
      saved: true,
      tarikh: tarikh,
      jualanDijangka: Number(totalExpected.toFixed(2)),
      jualanSebenar: Number(jualanSebenar.toFixed(2)),
      varians: Number((jualanSebenar - totalExpected).toFixed(2))
    };
  } finally {
    lock.releaseLock();
  }
}

function removeRowsByDate(sheet, dateCol1Based, tarikh) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;
  const vals = sheet.getRange(2, dateCol1Based, lastRow - 1, 1).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (formatDateYMD(vals[i][0]) === tarikh) {
      sheet.deleteRow(i + 2);
    }
  }
}

function deleteDailyRecord(payload) {
  const settings = getSettings();
  if (String(payload.pin || "").trim() !== String(settings.PIN_ADMIN || "1234")) {
    throw new Error("Kod PIN Admin tidak sah!");
  }
  const tarikh = String(payload.tarikh || "").trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  removeRowsByDate(ss.getSheetByName(SHEET_ITEMS), 2, tarikh);
  removeRowsByDate(ss.getSheetByName(SHEET_DAILY), 2, tarikh);
  SpreadsheetApp.flush();
  return { deleted: true, tarikh: tarikh };
}

function saveProduct(payload) {
  const settings = getSettings();
  if (String(payload.pin || "").trim() !== String(settings.PIN_ADMIN || "1234")) {
    throw new Error("Kod PIN Admin tidak sah!");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CATALOG);
  const kod = String(payload.kod || "").trim() || ("PRD-" + (sh.getLastRow()));
  const kategori = String(payload.kategori || "").trim().toUpperCase();
  const harga = Number(payload.harga) || 0;
  const ikon = String(payload.ikon || "📦").trim();
  const status = String(payload.status || "AKTIF").trim().toUpperCase();

  if (!kategori || harga <= 0) throw new Error("Nama kategori dan harga mesti diisi.");

  const rows = sh.getDataRange().getValues();
  let updated = false;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim() === kod) {
      sh.getRange(i + 1, 1, 1, 5).setValues([[kod, kategori, harga, ikon, status]]);
      updated = true;
      break;
    }
  }
  if (!updated) {
    sh.appendRow([kod, kategori, harga, ikon, status]);
  }
  return { saved: true, catalog: getCatalog() };
}

function deleteProduct(payload) {
  const settings = getSettings();
  if (String(payload.pin || "").trim() !== String(settings.PIN_ADMIN || "1234")) {
    throw new Error("Kod PIN Admin tidak sah!");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_CATALOG);
  const kod = String(payload.kod || "").trim();
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === kod) {
      sh.deleteRow(i + 1);
      break;
    }
  }
  return { deleted: true, catalog: getCatalog() };
}
