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
    } else if (action === 'getAvailableYears') {
      result = getAvailableYears();
    } else if (action === 'searchData') {
      result = searchData(payload.nim, payload.kategori, payload.offset, payload.limit);
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
 * Fungsi untuk mencari data riwayat mahasiswa berdasarkan NIM pada kategori spesifik
 */
function searchData(nim, kategori, offset = 0, limit = 10) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let results = [];
    
    // Cari nama mahasiswa otomatis dari data Sheet Master
    const studentName = getStudentName(nim);
    
    const sheetName = 'DATABASE ' + kategori;
    const sheet = ss.getSheetByName(sheetName);
    
    // Ambil data Kompen tambahan jika sedang memuat kategori Minus
    let kompenMap = {};
    if (kategori === 'Minus') {
      const kompenSheet = ss.getSheetByName('DATABASE Kompen');
      if (kompenSheet) {
        const kData = kompenSheet.getDataRange().getValues();
        for (let j = 1; j < kData.length; j++) {
           // NIM di kData[j][2], Jam Kompen di kData[j][8]
           if (kData[j][2] && kData[j][2].toString().trim() === nim.toString().trim()) {
              const kTime = new Date(kData[j][0]).getTime();
              kompenMap[kTime] = kData[j][8];
           }
        }
      }
    }
    
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      
      for (let i = 1; i < data.length; i++) {
        // NIM di data[i][2]
        if (data[i][2] && data[i][2].toString().trim() === nim.toString().trim()) {
          const tTime = new Date(data[i][0]).getTime();
          results.push({
            kategori: kategori,
            timestamp: tTime, 
            timestampStr: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"),
            namaInstruktur: data[i][1],
            nim: data[i][2],
            namaMahasiswa: data[i][3],
            section: data[i][4],
            statusJam: data[i][5],
            tanggalKejadian: data[i][6],
            keterangan: data[i][7],
            jumlahJam: data[i][8],
            jumlahJamKompen: kompenMap[tTime] || null
          });
        }
      }
    }
    
    results.sort((a, b) => b.timestamp - a.timestamp);
    
    const paginated = results.slice(offset, offset + limit);
    const hasMore = (offset + limit) < results.length;
    
    return { success: true, data: paginated, hasMore: hasMore, studentName: studentName };
  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }
}

/**
 * Fungsi untuk mengambil daftar tahun angkatan secara dinamis dari sheet Master
 */
function getAvailableYears() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let years = [];
    
    const sheets = ['Data_Mahasiswa Tk.1', 'Data_Mahasiswa Tk.2', 'Data_Mahasiswa Tk.3', 'Data_Mahasiswa Tk.4'];
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        // Ambil NIM pertama di baris ke-2 (asumsi baris 1 adalah header). Kolom C (3)
        const nimCell = sheet.getRange(2, 3).getValue();
        if (nimCell) {
          const nimStr = nimCell.toString().trim();
          const match = nimStr.match(/^(\d{4})/);
          if (match) {
            const year = parseInt(match[1]);
            if (!years.includes(year)) {
              years.push(year);
            }
          }
        }
      }
    });
    
    years.sort((a, b) => b - a);
    
    if (years.length === 0) {
      const currentYear = new Date().getFullYear();
      years = [currentYear, currentYear-1, currentYear-2, currentYear-3];
    }
    
    return { success: true, years: years };
  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }
}

/**
 * Fungsi untuk menambahkan data baru ke dalam Sheet berdasarkan Kategori
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
      payload.nim,
      payload.namaMahasiswa,
      payload.section,
      payload.kategori,         // Kolom F: Status jam
      payload.tanggalKejadian,
      payload.keterangan,
      payload.jumlahJam
    ]);
    
    // Jika Kategori Minus dan user mengisi Jam Kompen (Double Insert)
    if (payload.kategori === 'Minus' && payload.jumlahJamKompen) {
      const sheetKompen = ss.getSheetByName('DATABASE Kompen');
      if (sheetKompen) {
        sheetKompen.appendRow([
          timestamp,
          payload.namaInstruktur,
          payload.nim,
          payload.namaMahasiswa,
          payload.section,
          'Kompen',                 // Kolom F: Status jam untuk kompen
          payload.tanggalKejadian,
          payload.keterangan,
          payload.jumlahJamKompen
        ]);
      }
    }
    
    return { success: true, message: 'Data berhasil disimpan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan data: ' + error.toString() };
  }
}
