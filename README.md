Google Sheet Sync Setup

1. Create a new Google Sheet. Click Extensions > Apps Script.
2. Delete any existing code, paste the script below, and click Save. Copy
// Copy the code

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); 
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);
    var email = data.email;
    var txs = data.transactions || [];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID', 'Date', 'Type', 'Title', 'Category', 'Payment Method', 'Note', 'Amount', 'Location', 'Email']);
    }

    var allData = sheet.getDataRange().getValues();
    var headers = allData[0];
    var emailIdx = -1, idIdx = -1;
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h === 'email') emailIdx = i;
      if (h === 'id') idIdx = i;
    }

    if (emailIdx === -1 || idIdx === -1) return ContentService.createTextOutput(JSON.stringify({"error": "ID or Email column missing"})).setMimeType(ContentService.MimeType.JSON);

    var idToRow = {};
    for (var i = 1; i < allData.length; i++) {
      if (allData[i][emailIdx] && String(allData[i][emailIdx]).toLowerCase().trim() === String(email).toLowerCase().trim()) {
        idToRow[allData[i][idIdx]] = i + 1;
      }
    }

    var rowsToDelete = [];

    txs.forEach(function(t) {
      var row = idToRow[t.id];
      
      // STRICT DATE: Ensure it's exactly 10 characters (YYYY-MM-DD)
      var rawDate = t.date ? String(t.date).trim() : '';
      var cleanDate = rawDate.substring(0, 10); 
      
      var rowData = [t.id, cleanDate, t.type, t.title, t.category, t.paymentMethod, t.note || '', t.amount, t.location || '', email];

      if (t.deleted) {
        if (row) rowsToDelete.push(row);
      } else {
        if (row) {
          sheet.getRange(row, 1, 1, 10).setValues([rowData]); 
          sheet.getRange(row, 2).setNumberFormat('@'); // FORCE PLAIN TEXT on update
          delete idToRow[t.id]; 
        } else {
          sheet.appendRow(rowData);
          sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@'); // FORCE PLAIN TEXT on insert
        }
      }
    });

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
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var email = e.parameter.email;
    if (!email) return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);
    
    var allData = sheet.getDataRange().getValues();
    if (allData.length <= 1) return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);
    
    var headers = allData[0];
    var emailIdx = -1, idIdx = -1, dateIdx = -1, typeIdx = -1, titleIdx = -1, catIdx = -1, payIdx = -1, noteIdx = -1, amtIdx = -1, locIdx = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var h = String(headers[i]).trim().toLowerCase();
      if (h === 'email') emailIdx = i;
      else if (h === 'id') idIdx = i;
      else if (h === 'date') dateIdx = i;
      else if (h === 'type') typeIdx = i;
      else if (h === 'title') titleIdx = i;
      else if (h === 'category') catIdx = i;
      else if (h === 'payment method') payIdx = i;
      else if (h === 'note') noteIdx = i;
      else if (h === 'amount') amtIdx = i;
      else if (h === 'location') locIdx = i;
    }
    
    if (emailIdx === -1) return ContentService.createTextOutput(JSON.stringify({"transactions": [], "error": "Email column missing"})).setMimeType(ContentService.MimeType.JSON);

    var txs = [];
    for (var i = 1; i < allData.length; i++) {
      var row = allData[i];
      if (row[emailIdx] && String(row[emailIdx]).toLowerCase().trim() === String(email).toLowerCase().trim()) {
         
         var rawDate = dateIdx !== -1 ? row[dateIdx] : "";
         var cleanDate = "";
         
         if (rawDate) {
             if (rawDate instanceof Date) {
                 // If Sheets already corrupted it into a Date object
                 cleanDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
             } else {
                 // Force it to be exactly YYYY-MM-DD
                 var str = String(rawDate).trim();
                 cleanDate = str.substring(0, 10); 
                 if (cleanDate.length < 10 || cleanDate.includes(' ')) {
                     var parsed = new Date(str);
                     if (!isNaN(parsed.getTime())) {
                         cleanDate = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
                     }
                 }
             }
         }
         
         txs.push({
           id: idIdx !== -1 && row[idIdx] ? String(row[idIdx]) : "",
           date: cleanDate,
           type: typeIdx !== -1 && row[typeIdx] ? String(row[typeIdx]) : "expense",
           title: titleIdx !== -1 && row[titleIdx] ? String(row[titleIdx]) : "",
           category: catIdx !== -1 && row[catIdx] ? String(row[catIdx]) : "other-e",
           paymentMethod: payIdx !== -1 && row[payIdx] ? String(row[payIdx]) : "Cash",
           note: noteIdx !== -1 && row[noteIdx] ? String(row[noteIdx]) : "",
           amount: amtIdx !== -1 ? (parseFloat(row[amtIdx]) || 0) : 0,
           location: locIdx !== -1 && row[locIdx] ? String(row[locIdx]) : ""
         });
      }
    }
    return ContentService.createTextOutput(JSON.stringify({"transactions": txs})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({"error": err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}


3. Click Deploy > New Deployment. Select type: Web App.
4. Set "Execute as" to Me and "Who has access" to Anyone. Click Deploy.
5. Authorize permissions, copy the Web App URL, and paste it into the Sync input box.
