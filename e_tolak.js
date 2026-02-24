// ====================================================================
// PENGATURAN KINERJA (UBAH DI SINI)
// ====================================================================

// --- 1. PILIH MODE PERHITUNGAN ANDA ---
// (Pilih SALAH SATU dari 4 opsi di bawah ini)

// 'HABIS_SAJA'       -> Hanya melacak Habis (1 persentase di G3)
// 'TIDAK_PUNYA_SAJA' -> Hanya melacak Tidak Punya (1 persentase di G3)
// 'GABUNGAN'         -> Melacak total gabungan (1 persentase di G3)
// 'DUA_BATAS'        -> Melacak Habis (di G3) DAN Tidak Punya (di H3) secara terpisah

const METODE_PERHITUNGAN = 'HABIS_SAJA'; // <-- GANTI MODE DI SINI
const BATAS_MODE_HABIS_SAJA = 0;

// ====================================================================
// FUNGSI ANALISIS KINERJA
// ====================================================================

function hitungPersentaseKinerja() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAnalisis = ss.getSheetByName("Analisis_Etolak");
  if (!sheetAnalisis) return;

  const sheetHabis = ss.getSheetByName("Etolak_Habis");
  const sheetTP = ss.getSheetByName("Etolak_Tidak_Punya");
  if (!sheetHabis || !sheetTP) return;

  const hitungCeklis = (sheet) => {
    const data = sheet.getLastRow() > 2 ? sheet.getRange("D3:D" + sheet.getLastRow()).getValues() : [];
    return data.filter(r => r[0] === true).length;
  };
  
  const totalHabis = hitungCeklis(sheetHabis);
  const totalTP = hitungCeklis(sheetTP);

  const targetG = sheetAnalisis.getRange("G2");
  const targetH = sheetAnalisis.getRange("H2");

  switch (METODE_PERHITUNGAN) {
    
    // --- MODE 1: HANYA HABIS ---
    case 'HABIS_SAJA': {
      let persentase = 1.0;
      
      if (BATAS_MODE_HABIS_SAJA === 0) {
        if (totalHabis > 0) {
          persentase = 1.0 / (totalHabis + 1);
        }
      } else if (totalHabis > BATAS_MODE_HABIS_SAJA) {
        persentase = BATAS_MODE_HABIS_SAJA / totalHabis;
      }
      
      targetG.setValue(persentase).setNumberFormat("0%");
      targetH.clearContent();
      break;
    }
      
    // --- MODE 2: HANYA TIDAK PUNYA ---
    case 'TIDAK_PUNYA_SAJA': {
      let persentase = 1.0;
      
      if (BATAS_MODE_TIDAK_PUNYA_SAJA === 0) {
        if (totalTP > 0) {
          persentase = 1.0 / (totalTP + 1);
        }
      } else if (totalTP > BATAS_MODE_TIDAK_PUNYA_SAJA) {
        persentase = BATAS_MODE_TIDAK_PUNYA_SAJA / totalTP;
      }
      
      targetG.setValue(persentase).setNumberFormat("0%");
      targetH.clearContent();
      break;
    }
      
    // --- MODE 3: GABUNGAN ---
    case 'GABUNGAN': {
      let persentase = 1.0;
      const totalData = totalHabis + totalTP;
      
      if (BATAS_MODE_GABUNGAN === 0) {
        if (totalData > 0) {
          persentase = 1.0 / (totalData + 1);
        }
      } else if (totalData > BATAS_MODE_GABUNGAN) {
        persentase = BATAS_MODE_GABUNGAN / totalData;
      }
      
      targetG.setValue(persentase).setNumberFormat("0%");
      targetH.clearContent();
      break;
    }
      
    // --- MODE 4: DUA BATAS (HABIS & TIDAK PUNYA) ---
    case 'DUA_BATAS': {
      let skorHabis = 1.0;
      if (BATAS_DUAL_HABIS === 0) {
        if (totalHabis > 0) skorHabis = 1.0 / (totalHabis + 1);
      } else if (totalHabis > BATAS_DUAL_HABIS) {
        skorHabis = BATAS_DUAL_HABIS / totalHabis;
      }
      
      let skorTP = 1.0;
      if (BATAS_DUAL_TIDAK_PUNYA === 0) {
        if (totalTP > 0) skorTP = 1.0 / (totalTP + 1);
      } else if (totalTP > BATAS_DUAL_TIDAK_PUNYA) {
        skorTP = BATAS_DUAL_TIDAK_PUNYA / totalTP;
      }
      
      targetG.setValue(skorHabis).setNumberFormat("0%");
      targetH.setValue(skorTP).setNumberFormat("0%");
      break;
    }

    default:
      Logger.log("METODE_PERHITUNGAN tidak valid. Cek pengaturan.");
      targetG.clearContent();
      targetH.clearContent();
      break;
  }
}

// ====================================================================
// FUNGSI UTAMA (EVENT HANDLER) - VERSI PERBAIKAN RACE CONDITION + KOREKSI KASING
// ====================================================================

/**
 * Fungsi onEdit(e) utama untuk menangani semua interaksi.
 * @param {Object} e Event object dari onEdit.
 */
function handlerEdit(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  const sheetName = sheet.getName();

  // --- MULAI: BLOK TIMESTAMP 'Tambah_Produk' ---
  if (sheetName === "Tambah_Produk" && row >= 2 && [2, 3, 4].includes(col)) {
    const cellTanggal = sheet.getRange(row, 1);
    const cellB = sheet.getRange(row, 2);
    const cellC = sheet.getRange(row, 3);
    const cellD = sheet.getRange(row, 4);
    const isTriggerFilled = cellB.getValue() || cellC.getValue() || cellD.getValue();
    const isTanggalEmpty = !cellTanggal.getValue();

    if (isTriggerFilled && isTanggalEmpty) {
      cellTanggal.setValue(new Date()).setNumberFormat("dd/MM/yyyy"); 
    }
    
    const isAllTriggersEmpty = !cellB.getValue() && !cellC.getValue() && !cellD.getValue();
    if (isAllTriggersEmpty) {
      cellTanggal.clearContent();
    }
  }

  const isEtolakSheet = sheetName === "Etolak_Habis" || sheetName === "Etolak_Tidak_Punya";
  if (!isEtolakSheet || row < 3) return;

  const cellTanggal = sheet.getRange(row, 1); // A
  const cellInput = sheet.getRange(row, 2); // B
  const cellProduk = sheet.getRange(row, 3); // C
  const cellCheckbox = sheet.getRange(row, 4); // D
  const cellUser = sheet.getRange(row, 5); // E
  const db = ss.getSheetByName("DataBase");
  
  const tz = ss.getSpreadsheetTimeZone();
  const nowLocal = () => new Date(Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss"));
  const norm = v => String(v || "").toLowerCase();
  const getFilled = (colLetter) =>
    db.getRange(colLetter + "2:" + colLetter).getValues()
      .map(r => r[0]).filter(v => v !== "" && v !== null);

  if (col === 2) {
    const rawValue = e.value || cellInput.getValue();
    
    const typedRaw = String(rawValue);
    const typed = typedRaw.trim();
    const n = s => String(s || "").toLowerCase();

    if (typed === "" || typed === "undefined") {
      if (cellInput.getValue() === "") { 
          cellInput.clearDataValidations();
          cellTanggal.clearContent();     
          cellProduk.clearContent();       
          cellCheckbox.clearContent();     
          cellCheckbox.removeCheckboxes();  
          cellUser.clearContent();        
          cellUser.clearDataValidations(); 
          return; 
      }
    }

    const dbCol = sheetName === "Etolak_Habis" ? "B" : "C";
    const source = getFilled(dbCol);
    const filtered = source.filter(x => n(x).includes(n(typed))).slice(0, 20);
    const allowInvalid = (sheetName === "Etolak_Tidak_Punya"); 
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(filtered, true)
      .setAllowInvalid(allowInvalid)
      .build();
    cellInput.setDataValidation(rule);

    const exactFromDB = source.find(x => n(x) === n(typed));
    let finalName = exactFromDB || (allowInvalid ? typed : null);

    if (finalName) {
      cellProduk.setValue(finalName);
      
      if (!cellTanggal.getValue()) cellTanggal.setValue(nowLocal());
    } else {
      cellProduk.clearContent(); 
    }

    const userValid = !!cellUser.getValue();
    if (cellProduk.getValue() && userValid && !cellCheckbox.isChecked()) {
      cellCheckbox.insertCheckboxes();
      cellCheckbox.setValue(false);
    } else if (!cellProduk.getValue()) {
      cellCheckbox.clearContent();
      cellCheckbox.removeCheckboxes();
    }
  }

// ---------- 3) USER (kolom E) ----------
  if (col === 5) {
    try {
      Utilities.sleep(800); 
    } catch (err) {}
    if (!cellProduk.getValue()) {
        cellUser.clearContent(); 
        cellUser.clearDataValidations();
        SpreadsheetApp.getActive().toast("Harap isi 'Input Produk' (Kolom B) terlebih dahulu.", "Validasi Gagal", 3);
        return; 
    }

    const curr = String(cellUser.getValue() || "").trim();
    const n = s => String(s || "").toLowerCase(); 

    if (!curr) {
      cellUser.clearDataValidations();
      cellCheckbox.clearContent();
      cellCheckbox.removeCheckboxes();
      return;
    }
    
    const users = getFilled("D");
    const typedNorm = n(curr);
    const exactDBName = users.find(u => n(u) === typedNorm);
    const filteredUsers = users.filter(u => n(u).includes(typedNorm)).slice(0, 20);
    const userRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(filteredUsers.length > 0 ? filteredUsers : [curr]) 
      .setAllowInvalid(true)
  	.build();
    cellUser.setDataValidation(userRule);
    if (exactDBName) {
      if (cellUser.getValue() !== exactDBName) {
        cellUser.setValue(exactDBName);
      }
      
  	if (cellProduk.getValue() && !cellCheckbox.isChecked()) {
  	    if (!cellTanggal.getValue()) cellTanggal.setValue(nowLocal());
  	    cellCheckbox.insertCheckboxes();
  	    cellCheckbox.setValue(false);
  	}
    } else {
      cellUser.clearContent(); 
      cellUser.clearDataValidations(); 
  	  cellCheckbox.clearContent();
  	  cellCheckbox.removeCheckboxes();
      SpreadsheetApp.getActive().toast("Nama Petugas '" + curr + "' tidak ada di database.", "Input Ditolak", 4);
  	}
  }

  // ---------- 4) Tambah ke DB saat checkbox TRUE (khusus sheet Tidak Punya) ----------
  if (sheetName === "Etolak_Tidak_Punya" && col === 4 && e.value === "TRUE") {
    const produk = norm(cellProduk.getValue());
    const allC = getFilled("C").map(x => norm(x));
    
    if (produk && !allC.includes(produk)) {
      const cVals = db.getRange("C2:C").getValues();
      const emptyIndex = cVals.findIndex(r => !r[0]);
      const targetRow = emptyIndex !== -1 ? emptyIndex + 2 : allC.length + 2;
      db.getRange(targetRow, 3).setValue(cellProduk.getValue()); 
    }
  }

  // ---------- 5) Trigger analisis saat data final (checkbox di-tick) ----------
  if (col === 4) {
      try {
        hitungAnalisisStok();
        hitungTotalEtolak();
        hitungRekapBulanan();
        hitungPersentaseKinerja();
      } catch (err) {
        SpreadsheetApp.getActive().toast(`Error update analisis: ${err.message}`);
      }
  }
}

// ====================================================================
// FUNGSI ANALISIS
// ====================================================================
/**
 * Menghitung rekap produk (Analisis 1)
 */
function hitungAnalisisStok() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetHabis = ss.getSheetByName("Etolak_Habis");
  const sheetTidakPunya = ss.getSheetByName("Etolak_Tidak_Punya");
  // INI BAGIAN YANG DIPERBAIKI:
  const sheetAnalisis = ss.getSheetByName("Analisis_Etolak"); 

  if (!sheetHabis || !sheetTidakPunya || !sheetAnalisis) {
    // INI BAGIAN YANG DIPERBAIKI:
    Logger.log("Error: Salah satu sheet (Etolak_Habis, Etolak_Tidak_Punya, Analisis_Etolak) tidak ditemukan.");
    return;
  }

  const dataHabis = sheetHabis.getLastRow() > 2 ? sheetHabis.getRange("A3:D" + sheetHabis.getLastRow()).getValues() : [];
  const dataTidakPunya = sheetTidakPunya.getLastRow() > 2 ? sheetTidakPunya.getRange("A3:D" + sheetTidakPunya.getLastRow()).getValues() : [];

  const formatBulanTeks = (date) => {
    try {
      if (!(date instanceof Date)) date = new Date(date);
      if (isNaN(date.getTime())) return "";
      const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
        "Agustus", "September", "Oktober", "November", "Desember"];
      return `${bulan[date.getMonth()]} ${date.getFullYear()}`;
    } catch {
      return "";
    }
  };

  const mapHabis = new Map();
  const mapTidakPunya = new Map();

  // === Data Habis ===
  dataHabis.forEach(row => {
    const tanggal = row[0];
    const produk = row[2];
    const checked = row[3] === true;
    if (produk && checked && tanggal) {
      const periode = formatBulanTeks(tanggal);
      if (!mapHabis.has(produk)) mapHabis.set(produk, { jumlah: 0, periode });
      mapHabis.get(produk).jumlah++;
    }
  });

  // === Data Tidak Punya ===
  dataTidakPunya.forEach(row => {
    const tanggal = row[0];
    const produk = row[2];
    const checked = row[3] === true;
    if (produk && checked && tanggal) {
      const periode = formatBulanTeks(tanggal);
      if (!mapTidakPunya.has(produk)) mapTidakPunya.set(produk, { jumlah: 0, periode });
      mapTidakPunya.get(produk).jumlah++;
    }
  });

  // === Kosongkan Isi Laporan Analisis ===
  // Hapus data lama di A4:F
  sheetAnalisis.getRange("A4:F" + sheetAnalisis.getMaxRows()).clearContent();

  // === Susun data habis (kolom A–C) ===
  const habisOutput = Array.from(mapHabis.entries()).map(([produk, val]) => [produk, val.jumlah, val.periode]);
  habisOutput.sort((a, b) => a[0].localeCompare(b[0]));

  // === Susun data tidak punya (kolom D–F) ===
  const tidakPunyaOutput = Array.from(mapTidakPunya.entries()).map(([produk, val]) => [produk, val.jumlah, val.periode]);
  tidakPunyaOutput.sort((a, b) => a[0].localeCompare(b[0]));

  // === Gabungkan output ===
  const maxRows = Math.max(habisOutput.length, tidakPunyaOutput.length);
  const finalOutput = [];
  for (let i = 0; i < maxRows; i++) {
    const habisRow = habisOutput[i] || ["", "", ""];
    const tidakPunyaRow = tidakPunyaOutput[i] || ["", "", ""];
    finalOutput.push([...habisRow, ...tidakPunyaRow]);
  }

  if (finalOutput.length > 0) {
    sheetAnalisis.getRange(4, 1, finalOutput.length, 6).setValues(finalOutput);
  }
}

/**
 * Menghitung rekap bulanan per petugas (sekarang digabung ke Analisis_Etolak)
 */
function hitungRekapBulanan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAnalisis = ss.getSheetByName("Analisis_Etolak"); 
  const sheetHabis = ss.getSheetByName("Etolak_Habis");
  const sheetTP = ss.getSheetByName("Etolak_Tidak_Punya");

  if (!sheetAnalisis || !sheetHabis || !sheetTP) {
    Logger.log("Error: Salah satu sheet (Analisis_Etolak, Etolak_Habis, Etolak_Tidak_Punya) tidak ditemukan.");
    return;
  }

  const bulanList = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli",
    "Agustus", "September", "Oktober", "November", "Desember"];
  const fmtPeriode = (d) => {
    if (!(d instanceof Date)) d = new Date(d);
    if (isNaN(d.getTime())) return "";
    return `${bulanList[d.getMonth()]} ${d.getFullYear()}`;
  };
  const idxPeriode = (p) => {
    const [b, y] = String(p).split(" ");
    return parseInt(y, 10) * 12 + bulanList.indexOf(b);
  };

  // Helper untuk memproses data (tidak ada perubahan di sini)
  const prosesData = (sheet) => {
    const data = sheet.getLastRow() > 2 ? sheet.getRange("A3:E" + sheet.getLastRow()).getValues() : [];
    const mapPeriode = new Map();
    data.forEach(r => {
      const tgl = r[0], produk = r[2], chk = (r[3] === true), user = r[4];
      if (produk && chk && user && tgl) {
        const per = fmtPeriode(tgl);
        if (!mapPeriode.has(per)) mapPeriode.set(per, new Map());
        const uMap = mapPeriode.get(per);
        uMap.set(user, (uMap.get(user) || 0) + 1);
      }
    });
    return mapPeriode;
  };

  const mapPH = prosesData(sheetHabis);
  const mapPTP = prosesData(sheetTP);

  // Helper untuk menyusun output (tidak ada perubahan di sini)
  const susunOutput = (mapPeriode) => {
    const periods = Array.from(mapPeriode.keys()).sort((a, b) => idxPeriode(a) - idxPeriode(b));
    const output = [];
    periods.forEach(p => {
      const uMap = mapPeriode.get(p);
      for (let [u, c] of uMap.entries()) {
        output.push([u, c, p]);
      }
    });
    return output;
  };

  const outH = susunOutput(mapPH);
  const outTP = susunOutput(mapPTP);

  // === UBAH DI SINI: Tulis ke LOKASI BARU di sheet Analisis_Etolak ===
  const targetRangeHabis = sheetAnalisis.getRange("M4:O" + sheetAnalisis.getMaxRows());
  targetRangeHabis.clearContent();
  if (outH.length) {
    sheetAnalisis.getRange(4, 13, outH.length, 3).setValues(outH);
  }

  const targetRangeTdkPunya = sheetAnalisis.getRange("P4:R" + sheetAnalisis.getMaxRows());
  targetRangeTdkPunya.clearContent();
  if (outTP.length) {
    sheetAnalisis.getRange(4, 16, outTP.length, 3).setValues(outTP);
  }
}