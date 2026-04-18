const SPREADSHEET_ID = "1ObGH4HxcsDdRxF_NlOC6Cy9pehuxOpkWSVMj4Ahl5Qw";
const SHEET_NAME = "DataPegawai";
const FOLDER_FOTO_ID = "1J3JJD1FG1QdiRwYArSCMY74HuNzNEiiF"; 

const ADMIN_USER = "admin";
const ADMIN_PASS = "12345";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Si DAKON - UINSA') 
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** * PERBAIKAN LOGIN: 
 * Menambahkan status dan timestamp agar callback di HTML tidak "hang" atau blank.
 */
function checkLogin(username, password) {
  try {
    if (username === ADMIN_USER && password === ADMIN_PASS) {
      return { 
        success: true, 
        message: "Login Berhasil!",
        session: new Date().getTime()
      };
    } else {
      return { success: false, message: "Username atau Password Salah!" };
    }
  } catch (e) {
    return { success: false, message: "Error Sistem: " + e.toString() };
  }
}

function getSheetConnection() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error("Sheet '" + SHEET_NAME + "' tidak ditemukan!");
  return sheet;
}

function getRawData() {
  try {
    const sheet = getSheetConnection();
    const range = sheet.getDataRange();
    const data = range.getDisplayValues(); 
    
    if (data.length <= 1) return { headers: [], values: [] };
    
    return { 
      headers: data[0], 
      values: data.slice(1).filter(row => row.join("").trim() !== "") 
    };
  } catch (e) { 
    Logger.log("Error getRawData: " + e.message);
    return { error: e.toString() }; 
  }
}

function getDashboardStats() {
  try {
    const sheet = getSheetConnection();
    const data = sheet.getDataRange().getDisplayValues();
    if (data.length <= 1) return { total: 0, pria: 0, wanita: 0, units: {}, pangkat: {} };

    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    const rows = data.slice(1).filter(row => row.join("").trim() !== "");
    
    const idxGender = headers.findIndex(h => h.includes("jenis kelamin") || h.includes("gender"));
    const idxJabatan = headers.findIndex(h => h.includes("jabatan"));
    const idxPangkat = headers.findIndex(h => h.includes("pangkat") || h.includes("golongan"));

    let stats = { total: rows.length, pria: 0, wanita: 0, units: {}, pangkat: {} };

    rows.forEach(r => {
      if (idxGender !== -1) {
        const g = r[idxGender].toUpperCase().trim();
        if (['L', 'LAKI-LAKI', 'PRIA', 'Laki-Laki'].some(val => g.includes(val.toUpperCase()))) stats.pria++;
        else if (['P', 'PEREMPUAN', 'WANITA', 'Perempuan'].some(val => g.includes(val.toUpperCase()))) stats.wanita++;
      }
      if (idxJabatan !== -1) {
        const j = r[idxJabatan] || "Lainnya";
        stats.units[j] = (stats.units[j] || 0) + 1;
      }
      if (idxPangkat !== -1) {
        let p = r[20] || "Lainnya";
        stats.pangkat[p] = (stats.pangkat[p] || 0) + 1;
      }
    });
    return stats;
  } catch (e) { return { error: e.toString() }; }
}
function processForm(obj) {
  try {
    const sheet = getSheetConnection();
    // Gunakan getLastColumn agar dinamis
    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    let rowData = [];

    // Jika Update (Sudah ada row_index)
    if (obj.row_index && obj.row_index !== "") {
      const targetRow = parseInt(obj.row_index);
      const existingRow = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
      
      rowData = headers.map((h, i) => {
        const headLower = h.toLowerCase();
        // Cek apakah ada foto baru dari upload
        if (headLower.includes("foto")) {
          return obj.col_foto_url ? obj.col_foto_url : existingRow[i];
        }
        // Pastikan col_i ada di objek
        return (obj['col_' + i] !== undefined) ? obj['col_' + i] : existingRow[i];
      });
      
      sheet.getRange(targetRow, 1, 1, rowData.length).setValues([rowData]);
      return "Data Berhasil Diperbarui!";
    } 
    
    // Jika Data Baru
    else {
      rowData = headers.map((h, i) => {
        const headLower = h.toLowerCase();
        if (headLower.includes("foto")) return obj.col_foto_url || "";
        return obj['col_' + i] || "";
      });
      sheet.appendRow(rowData);
      return "Data Baru Ditambahkan!";
    }
  } catch (e) {
    Logger.log("Error processForm: " + e.message);
    return "Error: " + e.message;
  }
}

function processWithFile(formObject, base64Data, fileName) {
  try {
    const sheet = getSheetConnection();
    const headers = sheet.getDataRange().getValues()[0];
    
    const idxNip = headers.findIndex(h => {
      const hn = h.toUpperCase();
      return hn.includes("NIP") || hn.includes("NIDN") || hn.includes("ID");
    });
    
    const idPegawai = formObject['col_' + (idxNip !== -1 ? idxNip : 1)] || "IMG";
    const extension = fileName.split('.').pop();
    const newFileName = "FOTO_" + idPegawai + "_" + new Date().getTime() + "." + extension;

    const folder = DriveApp.getFolderById(FOLDER_FOTO_ID);
    const bytes = Utilities.base64Decode(base64Data.split(',')[1] || base64Data);
    const blob = Utilities.newBlob(bytes, 'image/jpeg', newFileName);
    const file = folder.createFile(blob);

    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    formObject['col_foto_url'] = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";

    return processForm(formObject);
  } catch (e) { 
    return "Error Upload: " + e.toString(); 
  }
}