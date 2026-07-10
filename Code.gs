/**
 * Fungsi utama untuk melayani halaman HTML (opsional jika dijalankan di dalam GAS)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Jam Plus Minus Kompen')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Fungsi API POST untuk diakses dari HTML eksternal (Website Live)
 */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    let result = {};
    
    if (action === 'login') {
      result = login(payload.username, payload.password);
    } else if (action === 'searchData') {
      result = searchData(payload.nim, payload.offset, payload.limit);
    } else if (action === 'addData') {
      result = addData(payload.data);
    } else {
      result = { success: false, message: 'Action tidak valid' };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Fungsi untuk proses login
 */
function login(username, password) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Akun_Dosen');
    
    if (!sheet) {
      return { success: false, message: 'Sheet Akun_Dosen tidak ditemukan.' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == username && data[i][1] == password) {
        return { success: true, namaDosen: data[i][2] };
      }
    }
    return { success: false, message: 'Username atau password salah.' };
  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }
}

/**
 * Fungsi pembantu untuk mencari nama mahasiswa berdasar tahun angkatan dari NIM
 */
function getStudentName(nim) {
  try {
    const nimStr = nim.toString().trim();
    // Hilangkan tanda strip agar 2023-6-050 sama dengan 20236050
    const cleanNim = nimStr.replace(/-/g, '');
    const yearMatch = cleanNim.match(/^(\d{4})/);
    let targetTk = 1; // Default Tk 1
    
    if (yearMatch) {
      const year = parseInt(yearMatch[1], 10);
      // Memori algoritma tebakan level tingkat:
      if (year >= 2025) targetTk = 1;
      else if (year === 2024) targetTk = 2;
      else if (year === 2023) targetTk = 3;
      else if (year <= 2022) targetTk = 4;
    }
    
    // Urutan pencarian: mulai dari targetTk, lalu naik ke tingkat berikutnya (jika ga valid, cek ke atasnya)
    let sheetsToSearch = [];
    for (let i = 0; i < 4; i++) {
      let tk = targetTk + i;
      if (tk > 4) tk -= 4; // Wrap around (misal target 4, lalu ngecek 1, 2, 3)
      sheetsToSearch.push(`Data_Mahasiswa Tk.${tk}`);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    for (let i = 0; i < sheetsToSearch.length; i++) {
      const sheetName = sheetsToSearch[i];
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      const data = sheet.getDataRange().getValues();
      // Kolom A (0): No Coin, Kolom B (1): Nama Lengkap, Kolom C (2): Nim
      for (let r = 1; r < data.length; r++) {
        let sheetNim = data[r][2] ? data[r][2].toString().trim().replace(/-/g, '') : "";
        if (sheetNim === cleanNim && sheetNim !== "") {
          return data[r][1].toString().trim(); // Return Nama Lengkap
        }
      }
    }
    return ""; // Tidak ditemukan
  } catch (error) {
    return ""; // Abaikan error untuk silent fallback
  }
}

/**
 * Fungsi untuk mencari data riwayat mahasiswa berdasarkan NIM
 */
function searchData(nim, offset = 0, limit = 10) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ['DATABASE Plus', 'DATABASE Minus', 'DATABASE Kompen'];
    let results = [];
    
    // Cari nama mahasiswa otomatis dari data Sheet Master
    const studentName = getStudentName(nim);
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      const kategori = sheetName.replace('DATABASE ', '');
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][3].toString().trim() === nim.toString().trim()) {
          results.push({
            kategori: kategori,
            timestamp: new Date(data[i][0]).getTime(), 
            timestampStr: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
            namaInstruktur: data[i][1],
            jumlahJam: data[i][2],
            nim: data[i][3],
            namaMahasiswa: data[i][4],
            section: data[i][5],
            tanggalKejadian: data[i][6],
            keterangan: data[i][7]
          });
        }
      }
    });
    
    results.sort((a, b) => b.timestamp - a.timestamp);
    
    const paginated = results.slice(offset, offset + limit);
    const hasMore = (offset + limit) < results.length;
    
    return { success: true, data: paginated, hasMore: hasMore, studentName: studentName };
  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }
}

/**
 * Fungsi untuk menyimpan data baru ke sheet yang sesuai
 */
function addData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = 'DATABASE ' + payload.kategori;
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return { success: false, message: 'Sheet ' + sheetName + ' tidak ditemukan.' };
    }
    
    const timestamp = new Date();
    
    sheet.appendRow([
      timestamp,
      payload.namaInstruktur,
      payload.jumlahJam,
      payload.nim,
      payload.namaMahasiswa,
      payload.section,
      payload.tanggalKejadian,
      payload.keterangan
    ]);
    
    return { success: true, message: 'Data berhasil disimpan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan data: ' + error.toString() };
  }
}
