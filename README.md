Google Sheet Sync Setup

1. Create a new Google Sheet. Click Extensions > Apps Script.
2. Delete any existing code, paste the script below, and click Save. Copy


function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevents simultaneous edits from different devices
  
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    var email = data.email;
    var txs = data.transactions || [];

    // If sheet is empty, add headers
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID', 'Date', 'Type', 'Title', 'Category', 'Payment Method', 'Note', 'Amount', 'Email']);
    }

    var allData = sheet.getDataRange().getValues();
    var idToRow = {};
    
    // Map existing rows for this specific user
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][8] === email) {
        idToRow[allData[i][0]] = i + 1;
      }
    }

    var rowsToDelete = [];

    // Process incoming transactions
    txs.forEach(function(t) {
      var row = idToRow[t.id];
      var rowData = [t.id, t.date, t.type, t.title, t.category, t.paymentMethod, t.note || '', t.amount, email];

      if (t.deleted) {
        if (row) rowsToDelete.push(row); // Mark for deletion
      } else {
        if (row) {
          sheet.getRange(row, 1, 1, 9).setValues([rowData]); // Update existing
          delete idToRow[t.id]; 
        } else {
          sheet.appendRow(rowData); // Add new
        }
      }
    });

    // Delete rows from bottom to top so row numbers don't shift
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

  if (!email) {
    return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);
  }

  var allData = sheet.getDataRange().getValues();
  if (allData.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);
  }

  var headers = allData[0];
  var emailIdx = headers.indexOf('Email');
  var txs = [];

  for (var i = 1; i < allData.length; i++) {
    if (allData[i][emailIdx] === email) {
      txs.push({
        id: allData[i][0],
        date: allData[i][1],
        type: allData[i][2],
        title: allData[i][3],
        category: allData[i][4],
        paymentMethod: allData[i][5],
        note: allData[i][6],
        amount: parseFloat(allData[i][7])
      });
    }
  }

  return ContentService.createTextOutput(JSON.stringify({"transactions": txs})).setMimeType(ContentService.MimeType.JSON);
}



3. Click Deploy > New Deployment. Select type: Web App.
4. Set "Execute as" to Me and "Who has access" to Anyone. Click Deploy.
5. Authorize permissions, copy the Web App URL, and paste it into the Sync input box.
