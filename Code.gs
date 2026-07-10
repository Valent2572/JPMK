/**
 * Fungsi utama untuk melayani halaman HTML (opsional jika dijalankan di dalam GAS)
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
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
 * @param {string} username 
 * @param {string} password 
 * @returns {object} status dan nama dosen jika berhasil
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
 * Fungsi untuk mencari data riwayat mahasiswa berdasarkan NIM
 */
function searchData(nim, offset = 0, limit = 10) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ['DATABASE Plus', 'DATABASE Minus', 'DATABASE Kompen'];
    let results = [];
    
    sheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      
      const data = sheet.getDataRange().getValues();
      const kategori = sheetName.replace('DATABASE ', '');
      
      for (let i = 1; i < data.length; i++) {
        if (data[i][3].toString() === nim.toString()) {
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
    
    return { success: true, data: paginated, hasMore: hasMore };
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
