/**
 * =========================================================================
 * KH PACKAGING — KEUSAHAWANAN PRO v5 (Google Apps Script Backend)
 * Termasuk Modul Jualan Harian + Audit Tunai + Gaji Pekerja PPKI
 * =========================================================================
 * Panduan Kemas Kini Apps Script:
 * 1. Buka Google Sheet anda -> Extensions -> Apps Script.
 * 2. Gantikan keseluruhan kod dengan kod ini & tekan Save.
 * 3. Klik Deploy -> Manage deployments -> ikon Pensel (Edit) -> Version: New version -> Deploy.
 *    (URL Web App anda kekal SAMA, tidak perlu tukar index.html!)
 * =========================================================================
 */

const SHEET_DAILY   = "Ringkasan_Harian";
const SHEET_ITEMS   = "Rekod_Jualan";
const SHEET_CATALOG = "Katalog_Produk";
const SHEET_SET     = "Tetapan_App";
const SHEET_GAJI    = "Rekod_Gaji";
const SHEET_PEKERJA = "Ringkasan_Pekerja";

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

    if (action === "saveDailySales") return jsonResponse({ ok: true, data: saveDailySales(payload) });
    if (action === "deleteDailyRecord") return jsonResponse({ ok: true, data: deleteDailyRecord(payload) });
    if (action === "saveProduct") return jsonResponse({ ok: true, data: saveProduct(payload) });
    if (action === "deleteProduct") return jsonResponse({ ok: true, data: deleteProduct(payload) });
    if (action === "saveGaji") return jsonResponse({ ok: true, data: saveGaji(payload) });
    if (action === "savePosSettings") return jsonResponse({ ok: true, data: savePosSettings(payload) });
    if (action === "savePosSale") return jsonResponse({ ok: true, data: savePosSale(payload) });
    if (action === "deleteGaji") return jsonResponse({ ok: true, data: deleteGaji(payload) });

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
    PIN_ADMIN: "1234",
    POS_MOD: "ON",
    POS_MOD_LAJU: "OFF",
    POS_MOD_LATIHAN: "OFF",
    POS_PIN_GURU: "1234",
    POS_JENIS_PENCETAK: "SISTEM",
    POS_SAIZ_KERTAS: "80",
    POS_RESIT_NAMA: "KH PACKAGING",
    POS_RESIT_HEADER2: "KEUSAHAWANAN PRO",
    POS_RESIT_ALAMAT: "Koperasi Sekolah",
    POS_RESIT_TELEFON: "",
    POS_RESIT_LOGO: "KH",
    POS_RESIT_FOOTER: "Terima kasih kerana menyokong Keusahawanan Pelajar!",
    POS_RESIT_QR_TEKS: "Imbas untuk maklum balas"
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

function getGajiData(ss) {
  const shGaji = ss.getSheetByName(SHEET_GAJI);
  const gajiList = [];
  if (shGaji && shGaji.getLastRow() > 1) {
    const rows = shGaji.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const nama = String(rows[i][3] || "").trim();
      if (!nama) continue;
      gajiList.push({
        id: String(rows[i][0] || ("GJ-" + i)),
        tarikh: formatDateYMD(rows[i][1]),
        bulan: String(rows[i][2] || ""),
        nama: nama.toUpperCase(),
        jenis: String(rows[i][4] || "GAJI HARIAN"),
        jumlah: Number(rows[i][5]) || 0,
        catatan: String(rows[i][6] || "")
      });
    }
  }
  gajiList.sort((a, b) => b.tarikh.localeCompare(a.tarikh));

  const staffSet = {};
  const shStaff = ss.getSheetByName(SHEET_PEKERJA);
  if (shStaff && shStaff.getLastRow() > 1) {
    const sRows = shStaff.getDataRange().getValues();
    for (let i = 1; i < sRows.length; i++) {
      const nm = String(sRows[i][1] || "").trim().toUpperCase();
      if (nm && nm !== "JUMLAH KESELURUHAN") staffSet[nm] = true;
    }
  }
  gajiList.forEach(g => { staffSet[g.nama] = true; });
  return { gajiList: gajiList, staffNames: Object.keys(staffSet).sort() };
}

function getBootstrapData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings();
  const catalog = getCatalog();

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
  dailyList.sort((a, b) => b.tarikh.localeCompare(a.tarikh));

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
  const gajiObj = getGajiData(ss);

  return {
    settings: {
      NAMA_PERNIAGAAN: settings.NAMA_PERNIAGAAN,
      MATA_WANG: settings.MATA_WANG,
      HAD_AMARAN_VARIANS: Number(settings.HAD_AMARAN_VARIANS) || 5
    },
    catalog: catalog,
    dailyList: dailyList,
    categoryStats: categoryStats,
    gajiList: gajiObj.gajiList,
    staffNames: gajiObj.staffNames
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
          catatan: String(rows[i][11]) || ""
        };
        break;
      }
    }
  }

  return { tarikh: tarikh, summary: summary, items: items };
}


function savePosSettings(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_SET);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SET);
    sh.appendRow(["Kunci", "Nilai"]);
  }
  const settings = payload.settings || {};
  const allowed = ["POS_MOD","POS_MOD_LAJU","POS_MOD_LATIHAN","POS_PIN_GURU","POS_JENIS_PENCETAK","POS_SAIZ_KERTAS","POS_RESIT_NAMA","POS_RESIT_HEADER2","POS_RESIT_ALAMAT","POS_RESIT_TELEFON","POS_RESIT_LOGO","POS_RESIT_FOOTER","POS_RESIT_QR_TEKS"];
  const rows = sh.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0] || "").trim();
    if (k) rowByKey[k] = i + 1;
  }
  let saved = 0;
  allowed.forEach(function (k) {
    if (settings[k] === undefined) return;
    const v = String(settings[k]);
    if (rowByKey[k]) sh.getRange(rowByKey[k], 2).setValue(v);
    else { sh.appendRow([k, v]); rowByKey[k] = sh.getLastRow(); }
    saved++;
  });
  SpreadsheetApp.flush();
  return { saved: true, count: saved };
}

function savePosSale(payload) {
  // Gabung jualan POS hari ini dengan rekod harian sedia ada, supaya tidak memadam rekod Rekod Harian.
  const tarikh = String(payload.tarikh || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tarikh)) throw new Error("Format tarikh tidak sah (YYYY-MM-DD).");
  const incoming = Array.isArray(payload.items) ? payload.items : [];
  if (incoming.length === 0) throw new Error("Tiada item POS untuk direkod.");
  const merged = {};
  try {
    const existing = getDailyDetail(tarikh);
    if (existing && existing.items) {
      existing.items.forEach(function (ex) {
        const sold = Math.max(0, (Number(ex.stokAwal) || 0) + (Number(ex.stokTambah) || 0) - (Number(ex.bakiStok) || 0));
        if (sold > 0) merged[String(ex.kategori || "").toUpperCase()] = { kategori: String(ex.kategori || "").toUpperCase(), stokAwal: sold, stokTambah: 0, bakiStok: 0, hargaSeunit: Number(ex.hargaSeunit) || 0, unitTerjual: sold, jualanDijangka: Number((sold * (Number(ex.hargaSeunit) || 0)).toFixed(2)) };
      });
    }
  } catch (e) { /* tiada rekod lama: teruskan */ }
  incoming.forEach(function (it) {
    const key = String(it.kategori || "").toUpperCase();
    const qty = Number(it.unitTerjual || it.stokAwal || 0);
    const harga = Number(it.hargaSeunit) || 0;
    if (!merged[key]) merged[key] = { kategori: key, stokAwal: qty, stokTambah: 0, bakiStok: 0, hargaSeunit: harga, unitTerjual: qty, jualanDijangka: Number((qty * harga).toFixed(2)) };
    else { merged[key].stokAwal += qty; merged[key].unitTerjual = merged[key].stokAwal; merged[key].jualanDijangka = Number((merged[key].stokAwal * merged[key].hargaSeunit).toFixed(2)); }
  });
  const items = Object.keys(merged).map(function (k) { return merged[k]; });
  const totalExpected = items.reduce(function (a, x) { return a + x.jualanDijangka; }, 0);
  return saveDailySales({ tarikh: tarikh, jualanSebenar: Number(totalExpected.toFixed(2)), catatan: String(payload.catatan || "POS"), items: items });
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

    const parts = tarikh.split("-").map(Number);
    const dtObj = new Date(parts[0], parts[1] - 1, parts[2]);
    const bulanStr = MALAY_MONTHS[dtObj.getMonth()] + " " + parts[0];
    const hariStr = MALAY_DAYS[dtObj.getDay()];

    removeRowsByDate(shItems, 2, tarikh);
    removeRowsByDate(shDaily, 2, tarikh);

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
        idRek, tarikh, bulanStr, parts[0],
        String(it.kategori || "").toUpperCase(),
        sAwal, sTambah, `=F${r}+G${r}`, baki, `=H${r}-I${r}`, harga, `=J${r}*K${r}`, "PWA APP"
      ]);
    });

    const dr = shDaily.getLastRow() + 1;
    const idHari = "HR-" + tarikh.replace(/-/g, "");
    shDaily.appendRow([
      idHari, tarikh, bulanStr, parts[0], hariStr,
      items.length, totalUnit,
      Number(totalExpected.toFixed(2)),
      Number(jualanSebenar.toFixed(2)),
      `=I${dr}-H${dr}`,
      `=IF(ROUND(J${dr},2)=0,"SEIMBANG",IF(J${dr}>0,"LEBIHAN (+)","KURANGAN (-)"))`,
      catatan, "PWA APP"
    ]);

    SpreadsheetApp.flush();
    return {
      saved: true, tarikh: tarikh,
      jualanDijangka: Number(totalExpected.toFixed(2)),
      jualanSebenar: Number(jualanSebenar.toFixed(2)),
      varians: Number((jualanSebenar - totalExpected).toFixed(2))
    };
  } finally {
    lock.releaseLock();
  }
}

function ensureGajiSheet(ss) {
  let sh = ss.getSheetByName(SHEET_GAJI);
  if (!sh) {
    sh = ss.insertSheet(SHEET_GAJI);
    sh.appendRow(["ID_Gaji", "Tarikh", "Bulan_Rujukan", "Nama_Pekerja", "Jenis_Transaksi", "Jumlah_RM", "Catatan"]);
  }
  return sh;
}

function saveGaji(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ensureGajiSheet(ss);
    const tarikh = String(payload.tarikh || "").trim();
    const nama = String(payload.nama || "").trim().toUpperCase();
    const jenis = String(payload.jenis || "GAJI HARIAN").trim().toUpperCase();
    const rawAmt = Math.abs(Number(payload.jumlah) || 0);
    const catatan = String(payload.catatan || "").trim();

    if (!tarikh || !nama) throw new Error("Sila masukkan tarikh dan nama pekerja.");
    if (rawAmt <= 0 && jenis.indexOf("HADIR") === -1) throw new Error("Jumlah RM mestilah lebih daripada 0.");

    const signedAmt = (jenis.indexOf("BAYAR") !== -1 || jenis.indexOf("PENDAHULUAN") !== -1) ? -rawAmt : rawAmt;
    const parts = tarikh.split("-").map(Number);
    const bulanStr = (parts.length >= 2 && parts[1] >= 1 && parts[1] <= 12) ? MALAY_MONTHS[parts[1] - 1].substring(0, 3) : "APP";
    const idGaji = "GJ-" + new Date().getTime().toString().slice(-6);

    sh.appendRow([idGaji, tarikh, bulanStr, nama, jenis, Number(signedAmt.toFixed(2)), catatan]);
    SpreadsheetApp.flush();
    return { saved: true, id: idGaji };
  } finally {
    lock.releaseLock();
  }
}

function deleteGaji(payload) {
  const settings = getSettings();
  if (String(payload.pin || "").trim() !== String(settings.PIN_ADMIN || "1234")) {
    throw new Error("Kod PIN Admin tidak sah!");
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_GAJI);
  if (!sh) throw new Error("Tab Rekod_Gaji tidak dijumpai.");
  const id = String(payload.id || "").trim();
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][0]).trim() === id) {
      sh.deleteRow(i + 1);
      break;
    }
  }
  SpreadsheetApp.flush();
  return { deleted: true, id: id };
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
  if (!updated) sh.appendRow([kod, kategori, harga, ikon, status]);
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
