// --- KONFIGURASI SCRIPT ---
const FOLDER_ID = '1E9c9t6d-BmQVJ27n_jTeNYQyG4z5Bggz'; // ID Folder tujuan penyimpanan PDF
const SHEET_DATA_NAME = 'req'; // Nama sheet sumber data yang sudah terpadu
const SHEET_FORM_NAME = 'form'; // Nama sheet template surat
const FORM_ID_CELL = 'C5'; // Sel di sheet 'Form' tempat No Surat disalin

// INDEKS KOLOM (Kolom dimulai dari 1)
const BAST_ID_COL_INDEX = 15; // Kolom O (No Surat BAST)
const DATE_COL_INDEX = 2; // Kolom B (Tanggal Request) - Diperlukan untuk format yyyymmdd
const PDF_LINK_COL_INDEX = 29; // Kolom AC (PDF File)
const PRINT_TRIGGER_COL_INDEX = 30; // Kolom AD (Print/Ceklist)

/**
 * Fungsi utama untuk membuat menu kustom di Google Sheet
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('PDF')
      .addItem('Buat PDF BAST', 'createPDFs')
      .addToUi();
}

/**
 * Memproses data di sheet 'req', menarik data berdasarkan trigger,
 * menyalin ID ke sheet 'form', lalu membuat PDF dengan konfigurasi cetak.
 */
function createPDFs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(SHEET_DATA_NAME);
  const formSheet = ss.getSheetByName(SHEET_FORM_NAME);
  
  if (!dbSheet || !formSheet) {
    SpreadsheetApp.getUi().alert('Error: Pastikan sheet "' + SHEET_DATA_NAME + '" dan "' + SHEET_FORM_NAME + '" ada.');
    return;
  }
  
  let folder;
  try {
    folder = DriveApp.getFolderById(FOLDER_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ID Folder Drive tidak valid.');
    return;
  }
  
  const dataRange = dbSheet.getDataRange();
  const data = dataRange.getValues();
  let processedCount = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowNumber = i + 1;
    
    const isPrintChecked = row[PRINT_TRIGGER_COL_INDEX - 1];
    const noSurat = row[BAST_ID_COL_INDEX - 1]; // Menggunakan No Surat BAST, bukan ReqID
    const rawDate = row[DATE_COL_INDEX - 1];
    
    if (isPrintChecked === true && noSurat !== '') {
      try {
        let formattedDate = '';
        if (rawDate instanceof Date) {
            formattedDate = Utilities.formatDate(rawDate, ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
        } else {
            formattedDate = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyyMMdd');
        }

        formSheet.getRange(FORM_ID_CELL).setValue(noSurat);
        SpreadsheetApp.flush(); 
        
        const safeNoSurat = noSurat.replace(/[\/\\?%*:|"<>]/g, '_');
        const fileName = formattedDate + '-BAST-' + safeNoSurat + '.pdf';
        
        const urlOptions = {
          size: 'A4', scale: 2, topMargin: 0.75, rightMargin: 0.75,
          bottomMargin: 0.75, leftMargin: 0.75, portrait: true
        };
        
        const pdfUrl = getPdfUrlWithConfig(formSheet.getSheetId(), urlOptions);
        
        const response = UrlFetchApp.fetch(pdfUrl, {
          headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
          muteHttpExceptions: true
        });
        
        const finalPdfBlob = response.getBlob().setName(fileName);
        const pdfFile = folder.createFile(finalPdfBlob);
        
        const pdfLinkDownload = pdfFile.getDownloadUrl(); 
        dbSheet.getRange(rowNumber, PDF_LINK_COL_INDEX).setValue(pdfLinkDownload);
        dbSheet.getRange(rowNumber, PRINT_TRIGGER_COL_INDEX).setValue(false);
        
        processedCount++;
        
      } catch (error) {
        SpreadsheetApp.getUi().alert('Gagal memproses baris ' + rowNumber + ' Error: ' + error.message);
      }
    }
  }
  
  formSheet.getRange(FORM_ID_CELL).clearContent();
  SpreadsheetApp.getUi().alert('Selesai! Total ' + processedCount + ' PDF dibuat.');
}

function getPdfUrlWithConfig(sheetId, options) {
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  let url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?';
  
  const parameters = {
    exportFormat: 'pdf', format: 'pdf', gid: sheetId,
    top_margin: options.topMargin, right_margin: options.rightMargin,
    bottom_margin: options.bottomMargin, left_margin: options.leftMargin,
    scale: options.scale || 2, portrait: options.portrait === true, 
    size: 'A4', printtitle: false, sheetnames: false, pagenumbers: false,
    gridlines: false, fzr: false
  };

  const queryString = Object.keys(parameters).map(key => key + '=' + parameters[key]).join('&');
  return url + queryString;
}
