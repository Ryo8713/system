// 處理前端的 GET 請求 (例如：獲取可用設備清單)
function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getAvailableItems') {
    const sheet = ss.getSheetByName('Inventory');
    const data = sheet.getDataRange().getValues();
    data.shift();

    const allItems = data
      .filter(row => row[0] !== "")
      .map(row => ({
        id: row[0],
        name: row[0]
      }));

    return ContentService.createTextOutput(JSON.stringify(allItems))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getUsers') {
    const sheet = ss.getSheetByName('Users');
    const data = sheet.getDataRange().getValues();
    data.shift();

    const users = data
      .filter(row => row[0] !== '')
      .map(row => ({
        borrowerId: row[0],
        borrowerName: row[1],
        phone: row[2],
        classroom: row[3] // 如果你 Users 沒有 classroom，這行可刪
      }));

    return ContentService.createTextOutput(JSON.stringify(users))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getOpenRecordsByBorrower') {
    const borrowerId = (e.parameter.borrowerId || '').trim();
    const sheet = ss.getSheetByName('Records');
    const data = sheet.getDataRange().getValues();
    data.shift();

    const openRecords = data
    .filter(row => String(row[1]).trim() === borrowerId && String(row[5] || '').trim() === '')
    .map(row => ({
      recordId: row[0],
      borrowerId: row[1],
      classroom: row[2],
      items: row[3],
      borrowTime: row[4],
      returnTime: row[5],
      note: row[6]
    }));

    return ContentService.createTextOutput(JSON.stringify(openRecords))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'getAllRecords') {
    const sheet = ss.getSheetByName('Records');
    const data = sheet.getDataRange().getValues();
    data.shift(); // remove header row

    const records = data.map(row => ({
      recordId: row[0],
      borrowerId: row[1],
      classroom: row[2],
      items: row[3],
      borrowTime: row[4],
      returnTime: row[5],
      note: row[6]
    }));

    return ContentService.createTextOutput(JSON.stringify(records))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (params.action === 'borrow') {
    const requestId = String(params.requestId || '').trim();
    const borrowerId = String(params.borrowerId || '').trim();
    const borrowerName = String(params.borrowerName || '').trim();
    const phone = String(params.phone || '').trim();
    const classroom = String(params.classroom || '').trim();
    const items = Array.isArray(params.items) ? params.items : [];
    const note = String(params.note || '').trim();

    if (!borrowerId || !borrowerName || !phone || !classroom) {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'error', message: 'missing fields' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (requestId) {
      const dedupeSheet = ss.getSheetByName('Requests');
      const dedupeData = dedupeSheet.getDataRange().getValues();
      
      for (let i = 1; i < dedupeData.length; i++) {
        if (String(dedupeData[i][0]).trim() === requestId) {
          return ContentService.createTextOutput(JSON.stringify({ result: 'success', isDuplicate: true }))
            .setMimeType(ContentService.MimeType.JSON);
        }
      }
      
      dedupeSheet.appendRow([requestId, new Date()]);
    }

    // 1) 儲存借用者到 Users（若不存在）
    const userSheet = ss.getSheetByName('Users');
    const userData = userSheet.getDataRange().getValues();
    const userExists = userData.some(row => String(row[0]).trim() === borrowerId);
    let userCreated = false;

    if (!userExists) {
      const phoneText = "'" + phone;
      userSheet.appendRow([borrowerId, borrowerName, phoneText]);
      userCreated = true;
    }

    // 2) 新增借用紀錄（不記錄 requestId）
    const recordSheet = ss.getSheetByName('Records');
    recordSheet.appendRow([
      Utilities.getUuid(),
      borrowerId,
      classroom,
      JSON.stringify(items),
      new Date(),
      "",
      note
    ]);

    return ContentService.createTextOutput(JSON.stringify({ result: 'success', userCreated }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.action === 'return') {
    const recordIds = Array.isArray(params.recordIds) ? params.recordIds : [];
    if (recordIds.length === 0) {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'error', message: 'missing recordIds' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getSheetByName('Records');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const recordId = data[i][0];
      if (recordIds.includes(recordId) && data[i][5] === '') {
        sheet.getRange(i + 1, 6).setValue(new Date());
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.action === 'clearAll') {
    const password = String(params.password || '').trim();
    if (password !== 'admin123') {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'error', message: 'invalid password' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const recordSheet = ss.getSheetByName('Records');
    const lastRow = recordSheet.getLastRow();
    if (lastRow > 1) {
      recordSheet
        .getRange(2, 1, lastRow - 1, recordSheet.getLastColumn())
        .clearContent();
    }

    const requestSheet = ss.getSheetByName('Requests');
    const requestLastRow = requestSheet.getLastRow();
    if (requestLastRow > 1) {
      requestSheet
        .getRange(2, 1, requestLastRow - 1, requestSheet.getLastColumn())
        .clearContent();
    }

    return ContentService.createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (params.action === 'update') {
    const recordId = String(params.recordId || '').trim();
    const classroom = String(params.classroom || '').trim();
    const items = Array.isArray(params.items) ? params.items : [];

    if (!recordId || !classroom) {
      return ContentService.createTextOutput(
        JSON.stringify({ result: 'error', message: 'missing fields' })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = ss.getSheetByName('Records');
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      const rowRecordId = data[i][0];
      const returnTime = data[i][5];
      if (rowRecordId === recordId && returnTime === '') {
        sheet.getRange(i + 1, 3).setValue(classroom);
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(items));
        return ContentService.createTextOutput(
          JSON.stringify({ result: 'success' })
        ).setMimeType(ContentService.MimeType.JSON);
      }
    }

    return ContentService.createTextOutput(
      JSON.stringify({ result: 'error', message: 'record not found or already returned' })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}