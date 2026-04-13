const SPREADSHEET_ID = '1c_fuN8xvDARWAVVd_-zjmuhj_XvQgqq0ZQuxyxGkm80';
const PDF_FOLDER_ID = '1E9c9t6d-BmQVJ27n_jTeNYQyG4z5Bggz';
const EXPIRE_HOURS = 48;

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'createReq') {
      return handleCreateReq(data);
    } else if (action === 'createBAST') {
      return handleCreateBAST(data);
    }

    return jsonResponse({ status: 'Error', message: 'Unknown post action ' + action });
  } catch (err) {
    // Fallback if form-urlencoded
    const paramAction = e.parameter.action;
    if (paramAction === 'createReq') return handleCreateReq(e.parameter);
    if (paramAction === 'createBAST') return handleCreateBAST(e.parameter);
    return jsonResponse({ status: 'Error', message: err.toString() });
  }
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    
    if (action === 'listReq') {
      return handleListReq();
    } else if (action === 'getReq') {
      return handleGetReq(e.parameter.id);
    } else if (action === 'approve') {
      return handleApprovalPage(e.parameter.id);
    }

    return ContentService.createTextOutput("MIS BAST API is running.");
  } catch (err) {
    return ContentService.createTextOutput("Error: " + err.toString());
  }
}

function handleCreateReq(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('req');
  
  if (!sheet) return jsonResponse({ status: 'Error', message: "Sheet 'req' tidak ditemukan!" });

  const id = 'REQ-' + Math.floor(Math.random() * 10000000) + '-' + new Date().getTime().toString().substr(-4);
  const impactLabels = { "1": "Unit", "2": "Departemen", "3": "Multi Departemen", "4": "Perusahaan" };
  const cakupanTeks = impactLabels[data.impact_scale] || "Unit";

  // Data to insert (columns A to N)
  // Kolom A(0): ID, B(1): Tgl, C(2): Unit, D(3): Dept, E(4): PIC, F(5): Jabatan, G(6): Tipe, H(7): Tujuan, I(8): Cakupan, J(9): Status, K(10): Progress, L(11): Remarks, M(12): Email, N(13): Nama Req
  const newRow = [
    id, new Date(), data.unit_kerja, data.departemen, data.pic_name,
    data.pic_position, data.request_type, data.purpose, cakupanTeks,
    'Pending Approval', 0, "", data.emailReq || "", data.pic_name
  ];

  sheet.appendRow(newRow);
  return jsonResponse({ status: 'Success', message: 'Request berhasil disimpan', id: id });
}

function handleListReq() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("req");
  if (!sheet) return jsonResponse({ status: 'Error', message: 'No db sheet' });

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return jsonResponse({ status: 'Success', data: [] });

  const numRows = Math.min(lastRow - 1, 100); 
  const data = sheet.getRange(lastRow - numRows + 1, 1, numRows, 25).getValues();
  
  const result = data.reverse().map(function(r) {
    return {
      id: r[0], // Kolom A
      date: r[1] instanceof Date ? r[1].toISOString() : r[1], // Kolom B
      unit: r[2],
      departemen: r[3],
      pic: r[4],
      position: r[5], 
      type: r[6],
      purpose: r[7],
      scale: r[8],    
      status: r[9] || 'Pending Review', 
      progress: isNaN(parseInt(r[10])) ? 0 : parseInt(r[10]),
      remarks: r[11] || '',
      email: r[12] || '',
      bastNoSurat: r[14] || '', // Kolom O (Index 14)
      bastStatus: r[24] || ''   // Kolom Y (Index 24)
    };
  });

  return jsonResponse({ status: 'Success', data: result });
}

function handleGetReq(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("req");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { // Kolom A is ID
      const r = data[i];
      return jsonResponse({
        status: 'Success',
        data: {
          id: r[0],
          date: r[1],
          unit: r[2],
          departemen: r[3],
          pic: r[4],
          position: r[5], 
          type: r[6],
          purpose: r[7]
        }
      });
    }
  }
  return jsonResponse({ status: 'Error', message: 'Request tidak ditemukan' });
}

function handleCreateBAST(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName("req");
  const table = sheet.getDataRange().getValues();

  const reqId = data.reqId; // Dicari
  let rowIdx = -1;
  let targetRow = null;

  for (let i = 1; i < table.length; i++) {
    if (table[i][0] === reqId) {
      rowIdx = i + 1; // getRange requires 1-based index
      targetRow = table[i];
      break;
    }
  }

  if (rowIdx === -1) {
    return jsonResponse({ status: 'Error', message: 'ReqID tidak valid atau tidak ditemukan.' });
  }

  const now = new Date();
  const expiredAt = new Date(now.getTime() + (EXPIRE_HOURS * 60 * 60 * 1000));
  
  sheet.getRange(rowIdx, 15).setValue(data.noSurat);     // Kolom O(15): BAST_NoSurat
  sheet.getRange(rowIdx, 16).setValue(data.perihal);     // Kolom P(16): BAST_Perihal
  sheet.getRange(rowIdx, 17).setValue(data.urlApp);      // Kolom Q(17): BAST_URL
  sheet.getRange(rowIdx, 18).setValue(data.fungsi);      // Kolom R(18): BAST_Fungsi
  sheet.getRange(rowIdx, 19).setValue(data.nama1);       // Kolom S(19): BAST_Nama1
  sheet.getRange(rowIdx, 20).setValue(data.email1);      // Kolom T(20): BAST_Email1
  sheet.getRange(rowIdx, 21).setValue(data.atasan1);     // Kolom U(21): BAST_Atasan1
  sheet.getRange(rowIdx, 22).setValue(data.emailAtasan1);// Kolom V(22): BAST_EmailAtasan1
  sheet.getRange(rowIdx, 23).setValue(data.email2);      // Kolom W(23): BAST_Email2 (Requestor)
  sheet.getRange(rowIdx, 24).setValue(data.atasan2);     // Kolom X(24): BAST_Atasan2
  sheet.getRange(rowIdx, 25).setValue(data.emailAtasan2);// Kolom Y(25): BAST_EmailAtasan2
  sheet.getRange(rowIdx, 26).setValue('Level 1');        // Kolom Z(26): BAST_StatusLevel
  sheet.getRange(rowIdx, 27).setValue(expiredAt);        // Kolom AA(27): BAST_ExpiredAt

  sendApprovalEmail(data.email1, data.nama1, data.perihal, reqId, 1);
  return jsonResponse({ status: 'Success', message: 'BAST terbuat. Email Level 1 ke ' + data.email1 });
}

function handleApprovalPage(id) {
  const userEmail = Session.getActiveUser().getEmail().toLowerCase();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dbSheet = ss.getSheetByName('req');
  const formSheet = ss.getSheetByName('form');
  const data = dbSheet.getDataRange().getValues();
  
  let rowIdx = -1;
  let r = null;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { rowIdx = i + 1; r = data[i]; break; }
  }

  if (!r) return HtmlService.createHtmlOutput("Data tidak ditemukan.");
  // BAST Data dimulai dari Kolom O(14)
  // r[14]: BAST_NoSurat
  // r[15]: BAST_Perihal
  // r[17]: BAST_Fungsi
  // r[18]: BAST_Nama1
  // r[19]: BAST_Email1
  // r[20]: BAST_Atasan1
  // r[21]: BAST_EmailAtasan1
  // r[22]: BAST_Email2 (Requestor)
  // r[23]: BAST_Atasan2
  // r[24]: BAST_EmailAtasan2
  // r[25]: BAST_StatusLevel
  // r[26]: BAST_ExpiredAt

  const now = new Date();
  const expiredAt = r[26] ? new Date(r[26]) : new Date();
  if (now > expiredAt && r[25] !== 'Approved') return HtmlService.createHtmlOutput("<h2>Link Expired</h2>");
  if (r[25] === 'Approved') return HtmlService.createHtmlOutput("<h2>BAST Sudah Selesai</h2>Dokumen sudah disetujui semua pihak.");

  let currentLevel = parseInt(String(r[25]).replace('Level ', '')) || 0;
  if(currentLevel === 0) return HtmlService.createHtmlOutput("<h2>Error Status</h2>Status belum siap di-approve.");

  let targetEmail = "";
  let targetName = "";
  
  if (currentLevel === 1) { targetEmail = r[19]; targetName = r[18]; } // Email1, Nama1
  else if (currentLevel === 2) { targetEmail = r[21]; targetName = r[20]; } // EmailAtasan1, Atasan1
  else if (currentLevel === 3) { targetEmail = r[22]; targetName = r[4]; } // Email2(Req), PIC(Req) -> wait PIC Req is Kolom E(r[4])
  else if (currentLevel === 4) { targetEmail = r[24]; targetName = r[23]; } // EmailAtasan2, Atasan2

  if (userEmail && userEmail !== targetEmail.toLowerCase()) {
    return HtmlService.createHtmlOutput(`<h2>Akses Tertunda</h2>Saat ini giliran <b>${targetName}</b> (${targetEmail}) untuk menyetujui.`);
  }

  // Set nomor surat
  formSheet.getRange("C5").setValue(r[14]); // r[14] is BAST_NoSurat

  const signInfo = `SIGNED DIGITALLY BY ${targetName}`;
  const dateInfo = `ON ${now.toLocaleString('id-ID')}`;

  if (currentLevel === 1) { 
    formSheet.getRange("A41").setValue(signInfo);
    formSheet.getRange("A42").setValue(dateInfo);
  } else if (currentLevel === 2) { 
    formSheet.getRange("D41").setValue(signInfo);
    formSheet.getRange("D42").setValue(dateInfo);
  } else if (currentLevel === 3) { 
    formSheet.getRange("F41").setValue(signInfo);
    formSheet.getRange("F42").setValue(dateInfo);
  } else if (currentLevel === 4) { 
    formSheet.getRange("H41").setValue(signInfo);
    formSheet.getRange("H42").setValue(dateInfo);
  }

  if (currentLevel < 4) {
    let nextLevel = currentLevel + 1;
    let nextStatus = "Level " + nextLevel;
    let nextEmail = "";
    let nextName = "";

    if (currentLevel === 1) { nextEmail = r[21]; nextName = r[20]; }
    else if (currentLevel === 2) { nextEmail = r[22]; nextName = r[4]; }
    else if (currentLevel === 3) { nextEmail = r[24]; nextName = r[23]; }
    
    dbSheet.getRange(rowIdx, 26).setValue(nextStatus); // Kolom Z
    sendApprovalEmail(nextEmail, nextName, r[15], id, nextLevel);
    return HtmlService.createHtmlOutput(`<h2>Approval Berhasil</h2>Terima kasih ${targetName}. Email selanjutnya telah dikirim ke ${nextName}.`);
  } else {
    // FINAL
    dbSheet.getRange(rowIdx, 26).setValue('Approved');
    dbSheet.getRange(rowIdx, 28).setValue("Final Approved by " + targetName + " on " + now); // Kolom AB Log
    
    const pdfFile = generatePdfFromSheet(ss, r);
    const allEmails = [r[19], r[21], r[22], r[24], "mis-dept@company.com"].join(",");
    
    MailApp.sendEmail({
      to: allEmails,
      subject: `[FINAL] BAST Approved: ${r[15]}`,
      htmlBody: `BAST <b>${r[15]}</b> telah selesai disetujui oleh semua pihak. Dokumen PDF terlampir.`,
      attachments: [pdfFile.getAs(MimeType.PDF)]
    });
    
    return HtmlService.createHtmlOutput("<h2>Approval Final Berhasil</h2>Dokumen PDF telah dikirim ke semua pihak.");
  }
}

function sendApprovalEmail(email, name, perihal, id, level) {
  const link = `${ScriptApp.getService().getUrl()}?action=approve&id=${id}`;
  MailApp.sendEmail({
    to: email,
    subject: `[ACTION] Approval BAST Level ${level}: ${perihal}`,
    htmlBody: `Halo <b>${name}</b>, <br><br>Mohon berikan persetujuan berjenjang (Level ${level}) untuk BAST <b>${perihal}</b>.<br><br>
    <a href="${link}" style="background:#1a73e8; color:white; padding:10px 20px; text-decoration:none; border-radius:5px;">SETUJUI SEKARANG</a>`
  });
}

function generatePdfFromSheet(ss, record) {
  const formSheet = ss.getSheetByName('form');
  SpreadsheetApp.flush();
  Utilities.sleep(2000); 

  const folder = DriveApp.getFolderById(PDF_FOLDER_ID);
  const tempSs = SpreadsheetApp.create(`Temp_${record[14]}`);
  formSheet.copyTo(tempSs).setName('BAST');
  tempSs.deleteSheet(tempSs.getSheets()[0]);
  
  const blob = tempSs.getAs(MimeType.PDF).setName(`BAST_${record[14]}.pdf`);
  const file = folder.createFile(blob);
  DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  
  return file;
}
