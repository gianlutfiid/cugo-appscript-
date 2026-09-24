function doGet() {



  return HtmlService.createHtmlOutputFromFile('index')



      .setTitle('Cugo App')



      .addMetaTag('viewport', 'width=device-width, initial-scale=1');



}







function getAppData() {



  var ss = SpreadsheetApp.getActiveSpreadsheet();



  SpreadsheetApp.flush();



 



  var data = {



    pesanan: getSheetDataBypassCache(ss, "Pesanan"),



    keuangan: getSheetDataBypassCache(ss, "Keuangan"),



    log: getSheetDataBypassCache(ss, "Log_Pekerjaan"),



    kedisiplinan: getSheetDataBypassCache(ss, "Kedisiplinan"),



    stok: getSheetDataBypassCache(ss, "Master_Stok"),



    log_stok: getSheetDataBypassCache(ss, "Log_Pengambilan"),



    karyawan: getSheetDataBypassCache(ss, "Data_Karyawan")



  };



  return JSON.stringify(data);



}







function getSheetDataBypassCache(ss, sheetName) {



  var sheet = ss.getSheetByName(sheetName);



  if (!sheet) return [];



  var lastRow = sheet.getLastRow();



  var lastCol = sheet.getLastColumn();



  if (lastRow === 0 || lastCol === 0) return [];



  return sheet.getRange(1, 1, lastRow, lastCol).getValues();



}







function bersihkanNota(notaInput) {



    if (!notaInput) return "";



    return notaInput.toString().replace(/[^a-zA-Z0-9]/g, "").trim().toLowerCase();



}







function cekPoinSudahMasukHariIni(ss, namaKaryawan, idPesananClean, tahapan, namaItem) {



  var logSheet = ss.getSheetByName("Log_Pekerjaan");



  if (!logSheet) return false;



  var logData = getSheetDataBypassCache(ss, "Log_Pekerjaan");



  var tglHariIni = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");



 



  for (var i = logData.length - 1; i >= 1; i--) {



      var row = logData[i];



      var logTgl = row[0];



      if (logTgl instanceof Date) logTgl = Utilities.formatDate(logTgl, "Asia/Jakarta", "dd/MM/yyyy");



     



      if (logTgl === tglHariIni && row[1].toString().toLowerCase().trim() === namaKaryawan.toLowerCase()) {



          var ket = row[2].toString().toLowerCase();



          var ketClean = bersihkanNota(ket);



         



          var tahapanCheck = tahapan.toLowerCase().trim();



          if (tahapanCheck === "diantar" || tahapanCheck === "pengantaran kurir" || tahapanCheck === "penjemputan") tahapanCheck = "kurir";



         



          var isTahapanMatch = ket.includes(tahapanCheck) || (tahapanCheck === "kurir" && (ket.includes("diantar") || ket.includes("pengantaran") || ket.includes("trip") || ket.includes("jemput") || ket.includes("pickup")));







          if (ketClean.includes(idPesananClean) && isTahapanMatch) {



              if (namaItem && namaItem !== "" && namaItem.toLowerCase() !== "layanan") {



                  var itemClean = bersihkanNota(namaItem);



                  if (ketClean.includes(itemClean)) return true;



              } else {



                  return true;



              }



          }



      }



  }



  return false;



}







function prosesUploadKasir(dataBaru) {



  var ss = SpreadsheetApp.getActiveSpreadsheet();



  var sheet = ss.getSheetByName("Pesanan");



 



  if (!sheet) {



    sheet = ss.insertSheet("Pesanan");



    sheet.appendRow(["Timestamp", "Tgl Terima", "No Nota", "Customer", "Rincian Layanan", "Total Tagihan", "Pembayaran", "Progres", "Kurir/Pembuat", "Pengambilan"]);



  }



 



  SpreadsheetApp.flush();



  var existingData = getSheetDataBypassCache(ss, "Pesanan");



  var existingNotas = {};



 



  for (var i = 1; i < existingData.length; i++) {



    var nota = existingData[i][2];



    if (nota) existingNotas[bersihkanNota(nota)] = i + 1;



  }



 



  for (var j = 0; j < dataBaru.length; j++) {



    var item = dataBaru[j];



    var layananTeks = item.layanan.join("\n");



    var rowData = [



      new Date(), item.tglTerima, item.nota, item.customer, layananTeks,                    



      item.totalTagihan, item.pembayaran, item.progres, item.pembuatNota, item.pengambilan                



    ];



   



    var notaKey = bersihkanNota(item.nota);



    if (existingNotas[notaKey]) {



      sheet.getRange(existingNotas[notaKey], 1, 1, 10).setValues([rowData]);



    } else {



      sheet.appendRow(rowData);



    }



  }



  SpreadsheetApp.flush();



  return "Sukses upload " + dataBaru.length + " data pesanan!";



}







function simpanPengeluaran(data) {



  var ss = SpreadsheetApp.getActiveSpreadsheet();



  var sheet = ss.getSheetByName("Keuangan");



  if (!sheet) {



    sheet = ss.insertSheet("Keuangan");



    sheet.appendRow(["Tanggal", "Tipe", "Kategori", "Keterangan", "Nominal"]);



  }



  var tglStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");



  sheet.appendRow([tglStr, "Keluar", data.kategori, data.keterangan, data.nominal]);



  SpreadsheetApp.flush();



  return "Sukses menyimpan pengeluaran!";



}







function simpanLogPekerjaan(data) {



  var ss = SpreadsheetApp.getActiveSpreadsheet();



  var sheet = ss.getSheetByName("Log_Pekerjaan");



  if (!sheet) {



    sheet = ss.insertSheet("Log_Pekerjaan");



    sheet.appendRow(["Tanggal", "Karyawan", "Jenis Pekerjaan", "Poin"]);



  }



  var tglStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "dd/MM/yyyy");



  sheet.appendRow([tglStr, data.karyawan, data.jenisPekerjaan, data.jumlah]);



  SpreadsheetApp.flush();



  return "Sukses mencatat poin karyawan!";



}







function simpanProgressPekerjaan(data) {



  var ss = SpreadsheetApp.getActiveSpreadsheet();



  var sheet = ss.getSheetByName("Pesanan");



  if (!sheet) return "Gagal: Sheet Pesanan tidak ditemukan.";



 



  SpreadsheetApp.flush();



  var existingData = getSheetDataBypassCache(ss, "Pesanan");



  var headers = existingData[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ""; });



 



  var notaTarget = bersihkanNota(data.nota);



  var rowToUpdate = 0;



 



  for (var i = 1; i < existingData.length; i++) {



    if (bersihkanNota(existingData[i][2]) === notaTarget) {



      rowToUpdate = i + 1;



      break;



    }



  }







  if (data.tahap && data.karyawan) {



      var tStr = data.tahap.toLowerCase().trim();



      var searchStr = tStr;



      if (searchStr === "pengantaran kurir" || searchStr === "pengantaran" || searchStr === "kurir" || searchStr === "penjemputan") searchStr = "diantar";



     



      var colIdx = -1;



      for (var c = 10; c < headers.length; c++) {



          if (headers[c] === searchStr || headers[c] === tStr) { colIdx = c; break; }



      }



      if (colIdx === -1) {



          colIdx = headers.indexOf(searchStr);



          if (colIdx === -1) colIdx = headers.indexOf(tStr);



      }



     



      var finalWorkerStr = data.karyawan;



     



      if (colIdx !== -1 && rowToUpdate > 0) {



          var currentWorker = existingData[rowToUpdate - 1][colIdx];



          if (currentWorker && currentWorker.toString().trim() !== "") {



              var arrWorkers = currentWorker.toString().toLowerCase().split(',').map(function(item) { return item.trim(); });



              if (arrWorkers.indexOf(data.karyawan.toLowerCase()) === -1) {



                  finalWorkerStr = currentWorker.toString().trim() + ", " + data.karyawan;



              } else {



                  finalWorkerStr = currentWorker.toString().trim();



              }



          }



          sheet.getRange(rowToUpdate, colIdx + 1).setValue(finalWorkerStr);



      } else if (colIdx === -1 && rowToUpdate > 0) {



          var maxCol = sheet.getLastColumn();



          sheet.getRange(1, maxCol + 1).setValue(data.tahap);



          sheet.getRange(rowToUpdate, maxCol + 1).setValue(data.karyawan);



      }







      var baseItemName = data.item ? data.item.split('(')[0].trim() : "";



      if (!cekPoinSudahMasukHariIni(ss, data.karyawan, bersihkanNota(data.nota), data.tahap, baseItemName)) {



          var qtyInput = parseFloat(data.qty) || 1;



          var itemLow = data.item ? data.item.toLowerCase() : "";



         



          var isSatuan = itemLow.includes("satuan") || itemLow.includes("pcs") || itemLow.includes("sepatu") || itemLow.includes("tas") || itemLow.includes("boneka") || itemLow.includes("sprei") || itemLow.includes("potong") || itemLow.includes("pasang") || itemLow.includes("helai") || itemLow.includes("ckl") || itemLow.includes("cks");



          var multiplier = 1;



          var unitStr = isSatuan ? "PCS" : "KG";



         



          if (tStr.includes("setrika")) multiplier = isSatuan ? 5 : 3;



          else if (tStr.includes("cuci")) multiplier = isSatuan ? 5 : 1;



          else if (tStr.includes("pengantaran") || tStr.includes("diantar") || tStr.includes("kurir") || tStr.includes("trip") || tStr.includes("jemput") || tStr.includes("pickup") || tStr.includes("penjemputan")) { multiplier = 3; unitStr = "Trip"; }



          else if (tStr.includes("packing") || tStr.includes("lipat")) { multiplier = 1; }



         



          var namaItemLog = data.item ? " - " + data.item : "";



          var ketFinal = "Update Nota " + data.nota + namaItemLog + " (" + data.tahap + ") " + qtyInput + " " + unitStr;



         



          simpanLogPekerjaan({ karyawan: data.karyawan, jenisPekerjaan: ketFinal, jumlah: qtyInput * multiplier });



      }



  }



 



  if (rowToUpdate > 0) {



      if (data.tahap === "Pengantaran Kurir" || data.tahap === "Penjemputan" || data.tahap === "Diambil") {



          var statusBaru = "Di Laundry";



          if (data.tahap === "Pengantaran Kurir") statusBaru = "Sudah Diantar";



          if (data.tahap === "Diambil") statusBaru = "Sudah Diambil";



         



          if (data.tahap !== "Penjemputan") sheet.getRange(rowToUpdate, 10).setValue(statusBaru);







          if (data.tahap === "Pengantaran Kurir" || data.tahap === "Penjemputan") {



            var kurirLama = (existingData[rowToUpdate-1][8] || "").toString().trim();



            if (kurirLama === "-" || kurirLama === "") kurirLama = "Tidak ada";



            if (!kurirLama.includes("/")) sheet.getRange(rowToUpdate, 9).setValue(kurirLama + " / " + data.karyawan);



            else sheet.getRange(rowToUpdate, 9).setValue(kurirLama.split("/")[0].trim() + " / " + data.karyawan);



          }



      }



      if (data.bayar && data.bayar !== "") {



          sheet.getRange(rowToUpdate, 7).setValue(data.bayar);



          simpanLogPekerjaan({ karyawan: (data.karyawan || "Admin"), jenisPekerjaan: "Pembayaran Nota " + data.nota + " -> " + data.bayar, jumlah: 0 });



      }



  }



  SpreadsheetApp.flush();



  return "Sukses mencatat progres pesanan!";



}







function prosesUploadPekerjaan(dataBaru) {



  var ss = SpreadsheetApp.getActiveSpreadsheet();



  var sheet = ss.getSheetByName("Pesanan");



  if (!sheet) return JSON.stringify({success: false, message: "Sheet Pesanan tidak ditemukan!"});







  SpreadsheetApp.flush();



  var existingData = getSheetDataBypassCache(ss, "Pesanan");



  var headers = existingData[0].map(function(h) { return h ? h.toString().toLowerCase().trim() : ""; });



 



  var tahapanWajib = ["cuci", "setrika", "lipat", "packing"];



  var maxCol = sheet.getLastColumn();



  for (var k = 0; k < tahapanWajib.length; k++) {



    var t = tahapanWajib[k];



    if (headers.indexOf(t) === -1) {



      maxCol++;



      sheet.getRange(1, maxCol).setValue(t.charAt(0).toUpperCase() + t.slice(1));



      headers.push(t);



      existingData[0][maxCol-1] = t.charAt(0).toUpperCase() + t.slice(1);



    }



  }






