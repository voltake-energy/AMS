// Fungsi ini hanya untuk trigger kotak kebenaran (authorization) jika perlu
function setupAuth() {
  DriveApp.getRootFolder();
  SpreadsheetApp.openById('19LIta_Dhtm_EO_g0u58QIBhQur1DARQU5vc9icAHFhI').getSheetByName('MASTERLIST');
  SpreadsheetApp.openById('19LIta_Dhtm_EO_g0u58QIBhQur1DARQU5vc9icAHFhI').getSheetByName('CALIBRATION LOG');
  SpreadsheetApp.openById('19LIta_Dhtm_EO_g0u58QIBhQur1DARQU5vc9icAHFhI').getSheetByName('Log Mobiliti Aset');
  SpreadsheetApp.openById('19LIta_Dhtm_EO_g0u58QIBhQur1DARQU5vc9icAHFhI').getSheetByName('ID PASSWORD'); // Trigger kebenaran tab auth
}

function doPost(e) {
  // Tetapan Output (Bypass CORS)
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    // Parsing Data dari Frontend
    var data = JSON.parse(e.postData.contents);
    var action = data.action || "add"; // Sama ada 'add', 'edit', 'recalibrate', 'audit', 'mobiliti', 'return', 'approve_return', 'update_profile'

    // === KONFIGURASI ID ANDA ===
    var SHEET_ID = '19LIta_Dhtm_EO_g0u58QIBhQur1DARQU5vc9icAHFhI';
    var MASTER_TAB = 'MASTERLIST';
    var CALIB_TAB = 'CALIBRATION LOG';
    var MOBILITI_TAB = 'Log Mobiliti Aset';
    var AUTH_TAB = 'ID PASSWORD'; // Tab untuk data pekerja
    var PDF_FOLDER_ID = '1sS5bKgSOZ0Mxt8ZqTXOtE4Rh_g78E125';
    var IMG_FOLDER_ID = '1AX0t9gqKFJGUCj9Mr26uIwpvFtHcMp0F';
    var RETURN_IMG_FOLDER_ID = '1AMZG-Nj380MMzHLucxPjVmyOF8-cqhQK'; // Folder khas gambar return
    var PROFILE_IMG_FOLDER_ID = '119f0qxH51XdXsmu8S8en39XpzeYzOoEf'; // Folder gambar profil (Baru)

    var certUrl = data.existingPdf || ""; 
    var picUrl = data.existingImg || "";  
    var logAmaran = [];

    // Jika ada muat naik PDF (Add/Edit/Recalibrate)
    if (data.pdfBase64 && data.pdfBase64 !== "") {
      try {
        var pdfFolder = DriveApp.getFolderById(PDF_FOLDER_ID);
        var pdfBlob = Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), data.pdfMimeType, data.pdfName);
        var pdfFile = pdfFolder.createFile(pdfBlob);
        certUrl = pdfFile.getUrl();
      } catch (err) {
        certUrl = "RALAT UPLOAD PDF: " + err.message;
        logAmaran.push("Sijil PDF gagal dimuat naik.");
      }
    }

    // Jika ada muat naik Gambar Aset/Pulang/Profil
    if (data.imgBase64 && data.imgBase64 !== "") {
      try {
        // Tentukan folder berdasarkan action
        var targetImgFolderId = IMG_FOLDER_ID;
        if (action === "return") targetImgFolderId = RETURN_IMG_FOLDER_ID;
        if (action === "update_profile") targetImgFolderId = PROFILE_IMG_FOLDER_ID;
        
        var imgFolder = DriveApp.getFolderById(targetImgFolderId);
        var imgBlob = Utilities.newBlob(Utilities.base64Decode(data.imgBase64), data.imgMimeType, data.imgName);
        var imgFile = imgFolder.createFile(imgBlob);
        picUrl = imgFile.getUrl(); // Akan digunakan untuk Profile, Aset Add/Edit, Return
      } catch (err) {
        picUrl = "RALAT UPLOAD GAMBAR: " + err.message;
        logAmaran.push("Gambar gagal dimuat naik.");
      }
    }

    // Jika ada muat naik Gambar Audit (Khusus untuk action === 'audit')
    var auditImgUrl = "";
    if (action === "audit" && data.auditImgBase64 && data.auditImgBase64 !== "") {
      try {
        var AUDIT_IMG_FOLDER_ID = '1OBwBUGyNWE-LNBie3cYoUF-GRmGo3CxW';
        var auditImgFolder = DriveApp.getFolderById(AUDIT_IMG_FOLDER_ID);
        var auditImgBlob = Utilities.newBlob(Utilities.base64Decode(data.auditImgBase64), data.auditImgMimeType, data.auditImgName);
        var auditImgFile = auditImgFolder.createFile(auditImgBlob);
        auditImgUrl = auditImgFile.getUrl();
      } catch (err) {
        logAmaran.push("Gambar Audit rosak/hilang gagal dimuat naik: " + err.message);
      }
    }

    var ss = SpreadsheetApp.openById(SHEET_ID);
    
    // ==========================================
    // ROUTING BERDASARKAN ACTION
    // ==========================================

    if (action === "update_profile") {
      // --- LOGIK: KEMASKINI PROFIL PEKERJA ---
      var sheetAuth = ss.getSheetByName(AUTH_TAB);
      if (!sheetAuth) throw new Error("Tab '" + AUTH_TAB + "' tidak ditemui!");

      var userId = data.userId;
      var newPassword = data.newPassword;
      var authDataValues = sheetAuth.getDataRange().getValues();
      var userRow = -1;

      // Cari dari data row 2 (index 1) ke bawah, anggap ID di Kolum A (Index 0)
      for (var i = 1; i < authDataValues.length; i++) {
        if (authDataValues[i][0] == userId) {
          userRow = i + 1; // +1 untuk nombor baris sebenar di Google Sheets
          break;
        }
      }

      if (userRow > -1) {
        if (newPassword && newPassword !== "") {
          sheetAuth.getRange(userRow, 4).setValue(newPassword); // Column D (Index 4)
        }
        if (picUrl !== "" && !picUrl.includes("RALAT")) {
          sheetAuth.getRange(userRow, 5).setValue(picUrl); // Column E (Index 5)
        }
      } else {
        throw new Error("ID Pekerja tidak dijumpai dalam pangkalan data auth.");
      }

    } else {
      // BAGI SEMUA ACTION SELAIN PROFIL, KITA PERLUKAN MASTERLIST
      var sheetMaster = ss.getSheetByName(MASTER_TAB);
      if (!sheetMaster) {
        throw new Error("Tab dengan nama 'MASTERLIST' tidak ditemui!");
      }

      if (action === "add") {
        // --- LOGIK: TAMBAH ASET BAHARU ---
        var lastRow = sheetMaster.getLastRow();
        var no = 1;
        if (lastRow > 1) {
          var lastNo = sheetMaster.getRange(lastRow, 1).getValue();
          no = (!isNaN(lastNo) && lastNo !== "") ? Number(lastNo) + 1 : lastRow;
        }

        var rowData = [
          no, data.category || "", data.make || "", data.model || "", data.serialNo || "",
          data.location || "", data.expiredDate || "", certUrl, data.calibBy || "", picUrl
        ];
        sheetMaster.appendRow(rowData);

      } else if (action === "edit") {
        // --- LOGIK: KEMASKINI ASET ---
        var rowIndex = Number(data.rowId);
        var rowData = [
          data.no || "", data.category || "", data.make || "", data.model || "", data.serialNo || "",
          data.location || "", data.expiredDate || "", certUrl, data.calibBy || "", picUrl
        ];
        sheetMaster.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);

      } else if (action === "recalibrate") {
        // --- LOGIK: KALIBRASI SEMULA ---
        var rowIndex = Number(data.rowId);

        // 1. UPDATE Tab MASTERLIST (Overwrite spesifik Column)
        if (data.expiredDate) sheetMaster.getRange(rowIndex, 7).setValue(data.expiredDate);
        if (certUrl) sheetMaster.getRange(rowIndex, 8).setValue(certUrl);
        if (data.calibBy) sheetMaster.getRange(rowIndex, 9).setValue(data.calibBy);

        // 2. INSERT Tab CALIBRATION LOG
        var sheetLog = ss.getSheetByName(CALIB_TAB);
        if (!sheetLog) {
          throw new Error("Tab 'CALIBRATION LOG' tidak ditemui! Pastikan ejaan tab betul.");
        }

        var logData = [
          data.category || "",
          data.make || "",
          data.model || "",
          data.serialNo || "",
          data.dateCalibrate || "",
          data.calibBy || ""
        ];
        sheetLog.appendRow(logData);

      } else if (action === "audit") {
        // --- LOGIK: AUDIT ASET ---
        var rowIndex = Number(data.rowId);
        
        // Column M (Index 13) = Tarikh Audit
        sheetMaster.getRange(rowIndex, 13).setValue(data.auditDate || "");
        
        // Column N (Index 14) = Audit Status
        sheetMaster.getRange(rowIndex, 14).setValue(data.auditStatus || "");
        
        // Column O (Index 15) = Remark
        sheetMaster.getRange(rowIndex, 15).setValue(data.remark || "");

        // Column P (Index 16) = Gambar Audit
        if (auditImgUrl !== "") {
          sheetMaster.getRange(rowIndex, 16).setValue(auditImgUrl);
        }
        
      } else if (action === "mobiliti") {
        // --- LOGIK: CHECKOUT MOBILITI ASET ---
        var sheetLogMobiliti = ss.getSheetByName(MOBILITI_TAB);
        
        if (!sheetLogMobiliti) {
          sheetLogMobiliti = ss.insertSheet(MOBILITI_TAB);
        }
        
        if (sheetLogMobiliti.getLastRow() === 0) {
          sheetLogMobiliti.appendRow([
            "Nama Peminjam", "Tarikh (Timestamp)", "Location", "Condition", "Remark", 
            "Tempoh Pinjaman", "Nama Projek", "Category", "Make", "Model", "Serial No",
            "Tarikh Pulang", "Kondisi Semasa Pulang", "Catatan Pulang", "Gambar Bukti",
            "Tarikh Disahkan (Admin)", "Kondisi Akhir", "Admin Pengesah", "Status Log"
          ]);
          sheetLogMobiliti.getRange("A1:S1").setFontWeight("bold");
          sheetLogMobiliti.setFrozenRows(1);
        }
        
        var timestamp = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
        var assets = data.assets;
        
        for (var i = 0; i < assets.length; i++) {
          var asset = assets[i];
          
          // 1. Simpan rekod dalam Tab Log Mobiliti (Dengan tambahan column baru)
          sheetLogMobiliti.appendRow([
            data.nama || "", timestamp, data.lokasi || "", data.kondisi || "", data.remark || "",
            data.tempoh || "", data.projek || "", asset.category || "", asset.make || "", asset.model || "", asset.serialNo || "",
            "", "", "", "", "", "", "", "Dipinjam" // Mengisi kolum L hingga S
          ]);
          
          // 2. Kemaskini status & lokasi pemegang baharu dalam MASTERLIST
          var rowId = Number(asset.rowId); 
          sheetMaster.getRange(rowId, 11).setValue(data.lokasi || ""); // Current Location (Kolum K)
          sheetMaster.getRange(rowId, 12).setValue(data.nama || ""); // Current Holder (Kolum L)
          sheetMaster.getRange(rowId, 17).setValue("Dipinjam"); // Status Pinjaman (Kolum Q)
        }
      } else if (action === "return") {
        // --- LOGIK: STAF PULANGKAN ASET ---
        var sheetLogMobiliti = ss.getSheetByName(MOBILITI_TAB);
        if (!sheetLogMobiliti) throw new Error("Tab 'Log Mobiliti Aset' tidak ditemui!");

        var serialNoTarget = data.serialNo;
        var dataLog = sheetLogMobiliti.getDataRange().getValues();
        var logRowToUpdate = -1;

        // Cari dari bawah (latest) baris pinjaman staf ini
        for (var i = dataLog.length - 1; i >= 1; i--) {
          // Kolum K (Index 10) = Serial No, Kolum S (Index 18) = Status Log
          if (dataLog[i][10] == serialNoTarget && dataLog[i][18] == "Dipinjam") {
            logRowToUpdate = i + 1; // +1 sebab index array mula 0, row sheet mula 1
            break;
          }
        }

        if (logRowToUpdate > -1) {
          // Update Log Mobiliti
          sheetLogMobiliti.getRange(logRowToUpdate, 12).setValue(data.returnDate); // Tarikh Pulang
          sheetLogMobiliti.getRange(logRowToUpdate, 13).setValue(data.returnCondition); // Kondisi
          sheetLogMobiliti.getRange(logRowToUpdate, 14).setValue(data.returnRemark); // Catatan
          if (picUrl !== "") {
            sheetLogMobiliti.getRange(logRowToUpdate, 15).setValue(picUrl); // Gambar
          }
          sheetLogMobiliti.getRange(logRowToUpdate, 19).setValue("Menunggu Pengesahan"); // Tukar Status Log
          
          // Update Masterlist
          var rowId = Number(data.rowId); 
          sheetMaster.getRange(rowId, 17).setValue("Menunggu Pengesahan"); // Tukar Status Pinjaman
        } else {
          throw new Error("Rekod pinjaman aktif bagi aset ini tidak dijumpai di dalam pangkalan data.");
        }

      } else if (action === "approve_return") {
        // --- LOGIK: ADMIN SAHKAN PEMULANGAN ---
        var sheetLogMobiliti = ss.getSheetByName(MOBILITI_TAB);
        if (!sheetLogMobiliti) throw new Error("Tab 'Log Mobiliti Aset' tidak ditemui!");

        var serialNoTarget = data.serialNo;
        var dataLog = sheetLogMobiliti.getDataRange().getValues();
        var logRowToUpdate = -1;

        // Cari dari bawah (latest) baris yang menunggu pengesahan
        for (var i = dataLog.length - 1; i >= 1; i--) {
          if (dataLog[i][10] == serialNoTarget && dataLog[i][18] == "Menunggu Pengesahan") {
            logRowToUpdate = i + 1;
            break;
          }
        }

        if (logRowToUpdate > -1) {
          var adminDate = Utilities.formatDate(new Date(), "GMT+8", "dd/MM/yyyy HH:mm:ss");
          
          // Update Log Mobiliti
          sheetLogMobiliti.getRange(logRowToUpdate, 16).setValue(adminDate); // Tarikh Admin Sahkan
          sheetLogMobiliti.getRange(logRowToUpdate, 17).setValue(data.approveCondition); // Kondisi Akhir
          sheetLogMobiliti.getRange(logRowToUpdate, 18).setValue(data.adminName); // Nama Admin
          sheetLogMobiliti.getRange(logRowToUpdate, 19).setValue("Selesai"); // Tukar Status Log

          // Update Masterlist
          var rowId = Number(data.rowId); 
          var finalLocation = data.approveLocation;

          sheetMaster.getRange(rowId, 11).setValue(finalLocation); // Current Location (Kolum K)
          
          if (finalLocation.toLowerCase() === "hq") {
              sheetMaster.getRange(rowId, 12).setValue(""); // Kosongkan Current Holder jika HQ
          } else {
              sheetMaster.getRange(rowId, 12).setValue("-"); // Standard kekosongan jika bukan HQ tapi dah pulang
          }
          
          sheetMaster.getRange(rowId, 17).setValue("Tersedia"); // Status Pinjaman ke Tersedia
        } else {
          throw new Error("Rekod menunggu pengesahan bagi aset ini tidak dijumpai di dalam pangkalan data.");
        }
      }
    } // End of else block (Non-Profile Actions)

    var successMsg = "Data berjaya diproses!";
    if (logAmaran.length > 0) successMsg += " Namun: " + logAmaran.join(", ");

    return output.setContent(JSON.stringify({
      status: "success",
      message: successMsg
    }));

  } catch (error) {
    return output.setContent(JSON.stringify({
      status: "error",
      message: "Ralat Sistem: " + error.toString()
    }));
  }
}

// Handler untuk akses rentas-domain (CORS) dari PWA Frontend ke Apps Script
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeader("Access-Control-Allow-Origin", "*")
    .setHeader("Access-Control-Allow-Methods", "POST, OPTIONS")
    .setHeader("Access-Control-Allow-Headers", "Content-Type");
}
