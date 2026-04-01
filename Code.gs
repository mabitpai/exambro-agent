function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    // Pastikan ini berjalan untuk GET/POST. Untuk GET, gunakan param ?pin=1234
    const pin = e.parameter && e.parameter.pin ? e.parameter.pin : (e.postData ? JSON.parse(e.postData.contents).pin : null);
    const nis = e.parameter && e.parameter.nis ? e.parameter.nis : (e.postData ? JSON.parse(e.postData.contents).nis : null);
    if (!pin) {
      return ContentService
        .createTextOutput(JSON.stringify({ valid: false, error: 'PIN tidak dikirim' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Ganti dengan ID Spreadsheet Anda (hanya diperlukan jika script tidak bound)
    const spreadsheetId = 'YOUR_SPREADSHEET_ID';
    const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Ujian');
    if (!sheet) {
      throw new Error('Sheet Ujian tidak ditemukan');
    }

    if (!nis) {
      return ContentService
        .createTextOutput(JSON.stringify({ valid: false, error: 'NIS tidak dikirim' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    let examFound = null;
    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0].toString() === pin.toString() && values[i][4].toString().toLowerCase() === 'aktif') {
        examFound = {
          subject: values[i][1],
          className: values[i][2],
          formUrl: values[i][3],
          durationMinutes: Number(values[i][5]) || 90
        };
        break;
      }
    }

    if (!examFound) {
      return ContentService
        .createTextOutput(JSON.stringify({ valid: false, error: 'PIN tidak valid atau ujian tidak aktif' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Cek data siswa di sheet Siswa/Peserta
    const studentSheet = ss.getSheetByName('Siswa') || ss.getSheetByName('Peserta');
    if (!studentSheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ valid: false, error: 'Sheet Siswa/Peserta tidak ditemukan' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const studentValues = studentSheet.getDataRange().getValues();
    let studentData = null;
    for (let j = 1; j < studentValues.length; j++) {
      if (studentValues[j][0] && studentValues[j][0].toString() === nis.toString()) {
        studentData = {
          name: studentValues[j][1] || '',
          studentClass: studentValues[j][2] || '',
          school: studentValues[j][3] || '',
          absen: studentValues[j][4] || '',
          participantNumber: studentValues[j][5] || ''
        };
        break;
      }
    }

    if (!studentData) {
      return ContentService
        .createTextOutput(JSON.stringify({ valid: false, error: 'NIS tidak ditemukan di data siswa' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify(Object.assign({ valid: true }, examFound, studentData)))
      .setMimeType(ContentService.MimeType.JSON);

    return ContentService
      .createTextOutput(JSON.stringify({ valid: false }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ valid: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}