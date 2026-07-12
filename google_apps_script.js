function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var name = sheet.getName();
    var data = sheet.getDataRange().getValues();
    if (data.length === 0) {
      result[name] = [];
      continue;
    }
    
    var headers = data[0];
    var rows = [];
    for (var r = 1; r < data.length; r++) {
      var row = {};
      var hasValue = false;
      for (var c = 0; c < headers.length; c++) {
        var val = data[r][c];
        row[headers[c]] = val;
        if (val !== "") hasValue = true;
      }
      if (hasValue) {
        rows.push(row);
      }
    }
    result[name] = rows;
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Invalid JSON"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var action = params.action;
  
  if (action === "bulk_init") {
    // 1. Setup Index Sheet in bulk
    var indexSheet = ss.getSheetByName("Client_Index") || ss.insertSheet("Client_Index");
    indexSheet.clear();
    var indexHeaders = ["Client No", "Party Name", "Address", "Ledger Page", "PDF Page", "Notes"];
    var indexValues = [indexHeaders];
    params.index_data.forEach(function(row) {
      indexValues.push([
        String(row.no || ""),
        String(row.party_name || ""),
        String(row.address || ""),
        row.page !== null ? Number(row.page) : "",
        row.pdf_page !== null ? Number(row.pdf_page) : "",
        String(row.notes || "")
      ]);
    });
    indexSheet.getRange(1, 1, indexValues.length, indexHeaders.length).setValues(indexValues);
    
    // 2. Setup Debit Sheet in bulk
    var debitSheet = ss.getSheetByName("Debit_Transactions") || ss.insertSheet("Debit_Transactions");
    debitSheet.clear();
    var debitHeaders = ["Client Name", "Ledger Page", "No", "Date", "Details", "Description", "Size", "Model", "PD", "Bill No", "Qty", "Rate", "Total", "Remarks"];
    var debitValues = [debitHeaders];
    params.debit_data.forEach(function(row) {
      debitValues.push([
        String(row.party_name || ""),
        row.ledger_page !== null ? String(row.ledger_page) : "",
        String(row.no || ""),
        String(row.date || ""),
        String(row.bi_ka || ""),
        String(row.description || ""),
        String(row.size || ""),
        String(row.model || ""),
        String(row.pd || ""),
        String(row.bill || ""),
        row.qty !== null ? String(row.qty) : "",
        row.taka !== null ? String(row.taka) : "",
        row.total !== null ? String(row.total) : "",
        String(row.remarks || "")
      ]);
    });
    debitSheet.getRange(1, 1, debitValues.length, debitHeaders.length).setValues(debitValues);
    
    // 3. Setup Credit Sheet in bulk
    var creditSheet = ss.getSheetByName("Credit_Transactions") || ss.insertSheet("Credit_Transactions");
    creditSheet.clear();
    var creditHeaders = ["Client Name", "Ledger Page", "No", "Date", "Amount", "Remarks"];
    var creditValues = [creditHeaders];
    params.credit_data.forEach(function(row) {
      creditValues.push([
        String(row.party_name || ""),
        row.ledger_page !== null ? String(row.ledger_page) : "",
        String(row.no || ""),
        String(row.date || ""),
        row.amount !== null ? String(row.amount) : "",
        String(row.remarks || "")
      ]);
    });
    creditSheet.getRange(1, 1, creditValues.length, creditHeaders.length).setValues(creditValues);
    
    return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Bulk initialization complete"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Single Row CRUD Operations
  if (action === "add_debit") {
    var sheet = ss.getSheetByName("Debit_Transactions");
    sheet.appendRow([
      params.party_name, params.ledger_page, params.no, params.date, params.bi_ka,
      params.description, params.size, params.model, params.pd, params.bill,
      params.qty, params.taka, params.total, params.remarks
    ]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "add_credit") {
    var sheet = ss.getSheetByName("Credit_Transactions");
    sheet.appendRow([
      params.party_name, params.ledger_page, params.no, params.date, params.amount, params.remarks
    ]);
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "edit_profile") {
    var sheet = ss.getSheetByName("Client_Index");
    var data = sheet.getDataRange().getValues();
    var clientNo = params.client_no;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(clientNo)) {
        sheet.getRange(i+1, 2).setValue(params.party_name);
        sheet.getRange(i+1, 3).setValue(params.address);
        sheet.getRange(i+1, 6).setValue(params.notes);
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "edit_debit" || action === "delete_debit") {
    var sheet = ss.getSheetByName("Debit_Transactions");
    var data = sheet.getDataRange().getValues();
    var name = params.party_name;
    var ledgerPage = params.ledger_page;
    var rowNo = String(params.no);
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === name && String(data[i][1]) === String(ledgerPage) && String(data[i][2]) === rowNo) {
        if (action === "delete_debit") {
          sheet.deleteRow(i+1);
        } else {
          sheet.getRange(i+1, 4).setValue(params.date);
          sheet.getRange(i+1, 5).setValue(params.bi_ka);
          sheet.getRange(i+1, 6).setValue(params.description);
          sheet.getRange(i+1, 7).setValue(params.size);
          sheet.getRange(i+1, 8).setValue(params.model);
          sheet.getRange(i+1, 9).setValue(params.pd);
          sheet.getRange(i+1, 10).setValue(params.bill);
          sheet.getRange(i+1, 11).setValue(params.qty);
          sheet.getRange(i+1, 12).setValue(params.taka);
          sheet.getRange(i+1, 13).setValue(params.total);
          sheet.getRange(i+1, 14).setValue(params.remarks);
        }
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "edit_credit" || action === "delete_credit") {
    var sheet = ss.getSheetByName("Credit_Transactions");
    var data = sheet.getDataRange().getValues();
    var name = params.party_name;
    var ledgerPage = params.ledger_page;
    var rowNo = String(params.no);
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === name && String(data[i][1]) === String(ledgerPage) && String(data[i][2]) === rowNo) {
        if (action === "delete_credit") {
          sheet.deleteRow(i+1);
        } else {
          sheet.getRange(i+1, 4).setValue(params.date);
          sheet.getRange(i+1, 5).setValue(params.amount);
          sheet.getRange(i+1, 6).setValue(params.remarks);
        }
        break;
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Action not found"}))
    .setMimeType(ContentService.MimeType.JSON);
}
