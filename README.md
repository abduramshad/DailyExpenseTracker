Google Sheet Sync Setup

1. Create a new Google Sheet. Click Extensions > Apps Script.
2. Delete any existing code, paste the script below, and click Save. Copy
// Copy the code

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevents simultaneous edits
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    var email = data.email;
    var txs = data.transactions || [];

    // Add headers if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID', 'Date', 'Type', 'Title', 'Category', 'Payment Method', 'Note', 'Amount', 'Location', 'Email']);
    }

    var allData = sheet.getDataRange().getValues();
    var idToRow = {};
    
    // Map existing rows for this specific user
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][9] === email) { // Email is now at index 9
        idToRow[allData[i][0]] = i + 1;
      }
    }

    var rowsToDelete = [];

    // Process incoming transactions
    txs.forEach(function(t) {
      var row = idToRow[t.id];
      // Force date to be a clean string (YYYY-MM-DD) to prevent timezone shifts
      var cleanDate = t.date.split('T')[0]; 
      
      var rowData = [t.id, cleanDate, t.type, t.title, t.category, t.paymentMethod, t.note || '', t.amount, t.location || '', email];

      if (t.deleted) {
        if (row) rowsToDelete.push(row);
      } else {
        if (row) {
          sheet.getRange(row, 1, 1, 10).setValues([rowData]); // Update
          delete idToRow[t.id]; 
        } else {
          sheet.appendRow(rowData); // Insert
        }
      }
    });

    // Delete rows from bottom to top
    rowsToDelete.sort(function(a, b) { return b - a; });
    rowsToDelete.forEach(function(row) { sheet.deleteRow(row); });

    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()})).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var email = e.parameter.email;

  if (!email) return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);

  var allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);

  var headers = allData[0];
  var emailIdx = headers.indexOf('Email');
  var txs = [];

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][emailIdx] === email) {
      // Force date output to string format
      var rawDate = allData[i][1];
      var cleanDate = rawDate instanceof Date ? Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd") : rawDate.split('T')[0];
      
      txs.push({
        id: allData[i][0],
        date: cleanDate,
        type: allData[i][2],
        title: allData[i][3],
        category: allData[i][4],
        paymentMethod: allData[i][5],
        note: allData[i][6],
        amount: parseFloat(allData[i][7]),
        location: allData[i][8] || ''
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify({"transactions": txs})).setMimeType(ContentService.MimeType.JSON);
}


3. Click Deploy > New Deployment. Select type: Web App.
4. Set "Execute as" to Me and "Who has access" to Anyone. Click Deploy.
5. Authorize permissions, copy the Web App URL, and paste it into the Sync input box.
