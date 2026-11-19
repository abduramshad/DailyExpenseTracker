Google Sheet Sync Setup

1. Create a new Google Sheet. Click Extensions > Apps Script.
2. Delete any existing code, paste the script below, and click Save. Copy


function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  if (data.action === 'sync') {
    sheet.clear();
    sheet.appendRow(['ID', 'Date', 'Type', 'Title', 'Category', 'Payment Method', 'Note', 'Amount', 'Email']);
    
    if (data.transactions && data.transactions.length > 0) {
      var rows = data.transactions.map(function(t) {
        return [t.id, t.date, t.type, t.title, t.category, t.paymentMethod, t.note, t.amount, data.email];
      });
      sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    }
    return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var email = e.parameter.email;
  var data = sheet.getDataRange().getValues();
  
  if (data.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify({"transactions": []})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var headers = data[0];
  var emailIdx = headers.indexOf('Email');
  var txs = [];
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][emailIdx] === email) {
      txs.push({
        id: data[i][0], date: data[i][1], type: data[i][2],
        title: data[i][3], category: data[i][4], paymentMethod: data[i][5],
        note: data[i][6], amount: parseFloat(data[i][7])
      });
    }
  }
  return ContentService.createTextOutput(JSON.stringify({"transactions": txs})).setMimeType(ContentService.MimeType.JSON);
}


3. Click Deploy > New Deployment. Select type: Web App.
4. Set "Execute as" to Me and "Who has access" to Anyone. Click Deploy.
5. Authorize permissions, copy the Web App URL, and paste it into the Sync input box.
