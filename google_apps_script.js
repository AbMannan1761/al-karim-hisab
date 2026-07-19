// Helper to safely parse strings to integer numbers (removing decimals)
function parseNumeric(val) {
  if (val === null || val === undefined || val === "") return "";
  var str = String(val).trim().replace(/,/g, '');
  if (!str || str === '〃' || str === '"') return val; // preserve ditto marks
  var parsed = parseFloat(str);
  return isNaN(parsed) ? val : Math.round(parsed);
}

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result = {};
  
  var targetSheets = ["Client_Index", "Debit_Transactions", "Credit_Transactions"];
  for (var i = 0; i < targetSheets.length; i++) {
    var name = targetSheets[i];
    var sheet = ss.getSheetByName(name);
    if (!sheet) continue;
    
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
    var debitHeaders = ["Client Name", "Ledger Page", "No", "Date", "বিবরণ", "কাপড়", "Size", "Model", "Bill No", "Qty", "Rate", "মোট", "সর্বশেষ বিল", "Remarks"];
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
        String(row.bill || ""),
        row.qty !== null ? parseNumeric(row.qty) : "",
        row.taka !== null ? parseNumeric(row.taka) : "",
        String(row.pd || ""),
        row.total !== null ? parseNumeric(row.total) : "",
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
        row.amount !== null ? parseNumeric(row.amount) : "",
        String(row.remarks || "")
      ]);
    });
    creditSheet.getRange(1, 1, creditValues.length, creditHeaders.length).setValues(creditValues);
    
    SpreadsheetApp.flush(); // Commit all pending sheet writes to allow pivot range validation
    syncClientSheets(ss);
    return ContentService.createTextOutput(JSON.stringify({status: "success", message: "Bulk initialization complete"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Single Row CRUD Operations
  if (action === "add_debit") {
    var sheet = ss.getSheetByName("Debit_Transactions");
    sheet.appendRow([
      params.party_name, params.ledger_page, params.no, params.date, params.bi_ka,
      params.description, params.size, params.model, params.bill, parseNumeric(params.qty),
      parseNumeric(params.taka), params.pd, parseNumeric(params.total), params.remarks
    ]);
    syncClientSheets(ss, params.party_name);
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "add_credit") {
    var sheet = ss.getSheetByName("Credit_Transactions");
    sheet.appendRow([
      params.party_name, params.ledger_page, params.no, params.date, parseNumeric(params.amount), params.remarks
    ]);
    syncClientSheets(ss, params.party_name);
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "edit_profile") {
    var sheet = ss.getSheetByName("Client_Index");
    var data = sheet.getDataRange().getValues();
    var clientNo = params.client_no;
    var oldName = "";
    var newName = params.party_name;
    var rowIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(clientNo)) {
        oldName = String(data[i][1]).trim();
        sheet.getRange(i+1, 2).setValue(newName);
        sheet.getRange(i+1, 3).setValue(params.address);
        sheet.getRange(i+1, 6).setValue(params.notes);
        rowIdx = i;
        break;
      }
    }
    if (oldName && oldName !== newName && rowIdx !== -1) {
      var oldCleanName = oldName.replace(/[\\\/\?\:\*\[\]]/g, "");
      var oldSheetName = clientNo + ". " + oldCleanName;
      if (oldSheetName.length > 31) oldSheetName = oldSheetName.substring(0, 31);
      var oldSheet = ss.getSheetByName(oldSheetName);
      if (oldSheet) {
        ss.deleteSheet(oldSheet);
      }
    }
    syncClientSheets(ss, newName);
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
          sheet.getRange(i+1, 9).setValue(params.bill);
          sheet.getRange(i+1, 10).setValue(parseNumeric(params.qty));
          sheet.getRange(i+1, 11).setValue(parseNumeric(params.taka));
          sheet.getRange(i+1, 12).setValue(params.pd);
          sheet.getRange(i+1, 13).setValue(parseNumeric(params.total));
          sheet.getRange(i+1, 14).setValue(params.remarks);
        }
        break;
      }
    }
    syncClientSheets(ss, name);
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
          sheet.getRange(i+1, 5).setValue(parseNumeric(params.amount));
          sheet.getRange(i+1, 6).setValue(params.remarks);
        }
        break;
      }
    }
    syncClientSheets(ss, name);
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "add_client") {
    var sheet = ss.getSheetByName("Client_Index");
    var data = sheet.getDataRange().getValues();
    var newNo = 1;
    if (data.length > 1) {
      var maxNo = 0;
      for (var i = 1; i < data.length; i++) {
        var num = parseInt(data[i][0], 10);
        if (!isNaN(num) && num > maxNo) {
          maxNo = num;
        }
      }
      newNo = maxNo + 1;
    }
    sheet.appendRow([
      newNo,
      params.party_name,
      params.address,
      params.ledger_page,
      params.pdf_page || "",
      params.notes || ""
    ]);
    SpreadsheetApp.flush();
    syncClientSheets(ss, params.party_name);
    return ContentService.createTextOutput(JSON.stringify({status: "success", client_no: newNo}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (action === "delete_client") {
    var sheet = ss.getSheetByName("Client_Index");
    var data = sheet.getDataRange().getValues();
    var clientNo = params.client_no;
    var partyName = "";
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(clientNo)) {
        partyName = String(data[i][1]).trim();
        sheet.deleteRow(i + 1);
        break;
      }
    }
    if (partyName) {
      var cleanName = partyName.replace(/[\\\/\?\:\*\[\]]/g, "");
      var sheetName = clientNo + ". " + cleanName;
      if (sheetName.length > 31) sheetName = sheetName.substring(0, 31);
      var clientSheet = ss.getSheetByName(sheetName);
      if (clientSheet) {
        ss.deleteSheet(clientSheet);
      }
      
      var debitSheet = ss.getSheetByName("Debit_Transactions");
      if (debitSheet) {
        var debitData = debitSheet.getDataRange().getValues();
        for (var i = debitData.length - 1; i >= 1; i--) {
          if (String(debitData[i][0]).trim() === partyName) {
            debitSheet.deleteRow(i + 1);
          }
        }
      }
      var creditSheet = ss.getSheetByName("Credit_Transactions");
      if (creditSheet) {
        var creditData = creditSheet.getDataRange().getValues();
        for (var i = creditData.length - 1; i >= 1; i--) {
          if (String(creditData[i][0]).trim() === partyName) {
            creditSheet.deleteRow(i + 1);
          }
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "bulk_add_debit") {
    var sheet = ss.getSheetByName("Debit_Transactions");
    var rows = params.rows;
    rows.forEach(function(row) {
      sheet.appendRow([
        row.party_name, row.ledger_page, row.no, row.date, row.bi_ka,
        row.description, row.size, row.model, row.bill, parseNumeric(row.qty),
        parseNumeric(row.taka), row.pd, parseNumeric(row.total), row.remarks
      ]);
    });
    if (rows.length > 0) {
      syncClientSheets(ss, rows[0].party_name);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "bulk_add_credit") {
    var sheet = ss.getSheetByName("Credit_Transactions");
    var rows = params.rows;
    rows.forEach(function(row) {
      sheet.appendRow([
        row.party_name, row.ledger_page, row.no, row.date, parseNumeric(row.amount), row.remarks
      ]);
    });
    if (rows.length > 0) {
      syncClientSheets(ss, rows[0].party_name);
    }
    return ContentService.createTextOutput(JSON.stringify({status: "success"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Action not found"}))
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Al Karim Tools')
      .addItem('Sync Client Sheets', 'syncClientSheetsMenu')
      .addItem('Sort Sheets', 'sortSheetsMenu')
      .addToUi();
}

function syncClientSheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Syncing client sheets... this may take a moment.', 'Sync Status', -1);
  syncClientSheets(ss);
  ss.toast('Sorting sheets... please wait.', 'Sort Status', -1);
  sortSheets(ss);
  ss.toast('Sync and sorting complete!', 'Sync Status', 5);
}

function syncClientSheets(ss, targetClientName) {
  // Clean database sheet numeric formats to remove text decimals
  cleanDatabaseSheets(ss);
  
  // First propagate any name updates from Client_Index to Debit_Transactions and Credit_Transactions
  syncNamesInTransactions(ss);
  
  var indexSheet = ss.getSheetByName("Client_Index");
  if (!indexSheet) return;
  var indexData = indexSheet.getDataRange().getValues();
  
  var fontName = "Segoe UI";
  
  // 1. Collect all active sheet names based on Client_Index
  var activeSheetNames = {};
  for (var i = 1; i < indexData.length; i++) {
    var clientNo = String(indexData[i][0] || "").trim();
    var partyName = String(indexData[i][1] || "").trim();
    var ledgerPage = indexData[i][3];
    if (!partyName || !ledgerPage) continue;
    
    var cleanName = partyName.replace(/[\\\/\?\:\*\[\]]/g, "");
    var sheetName = clientNo + ". " + cleanName;
    if (sheetName.length > 31) {
      sheetName = sheetName.substring(0, 31);
    }
    activeSheetNames[sheetName] = true;
  }
  
  // 2. Clean up any orphan client sheets (e.g., if renamed)
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    // If it is a client sheet (starts with P followed by number or number followed by dot)
    if (/^P\d+\s*-/.test(name) || /^\d+\.\s*/.test(name)) {
      if (!activeSheetNames[name]) {
        ss.deleteSheet(sheet);
      }
    }
  });
  
  // 3. Process sheets
  for (var i = 1; i < indexData.length; i++) {
    var clientNo = String(indexData[i][0] || "").trim();
    var partyName = String(indexData[i][1] || "").trim();
    var address = String(indexData[i][2] || "").trim();
    var ledgerPage = indexData[i][3];
    var notes = String(indexData[i][5] || "").trim();
    
    if (!partyName || !ledgerPage) continue;
    
    // If targetClientName is provided, only process that specific client
    if (targetClientName && partyName !== targetClientName) {
      continue;
    }
    
    var cleanName = partyName.replace(/[\\\/\?\:\*\[\]]/g, "");
    var sheetName = clientNo + ". " + cleanName;
    if (sheetName.length > 31) {
      sheetName = sheetName.substring(0, 31);
    }
    
    var sheet = ss.getSheetByName(sheetName);
    var isNew = false;
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      isNew = true;
    } else {
      // Check if D2 (Name), D3 (Address), D4 (Notes) match in a single call to improve performance
      var metaValues = sheet.getRange("D2:D4").getValues();
      var currentName = String(metaValues[0][0] || "").trim();
      var currentAddress = String(metaValues[1][0] || "").trim();
      var currentNotes = String(metaValues[2][0] || "").trim();
      if (currentName === partyName && currentAddress === address && currentNotes === notes) {
        sheet.setFrozenRows(8); // Ensure rows are frozen even if we skip formatting
        
        // Ensure column widths are updated to 75px (~1 inch) for existing skipped sheets
        for (var col = 1; col <= 17; col++) {
          if (col === 1) {
            sheet.setColumnWidth(col, 25); // Column A is 0.25 inches (25px)
          } else if (col === 5) {
            sheet.setColumnWidth(col, 45); // Column E (Size) is set to 45px
          } else if (col === 13) {
            sheet.setColumnWidth(col, 25); // separator column
          } else {
            sheet.setColumnWidth(col, 75); // approx. 1 inch
          }
        }
        
        // Force update headers in Row 8 even if skipping full format
        var fontName = "Segoe UI";
        var debitHeaders = ["No", "Date", "Details (বিঃ কাঃ)", "Description", "Size", "Model", "Bill No", "Qty", "Rate", "মোট", "সর্বশেষ বিল", "Remarks"];
        debitHeaders.forEach(function(headerText, index) {
          var cell = sheet.getRange(8, index + 1);
          cell.setValue(headerText)
              .setFontFamily(fontName)
              .setFontSize(10)
              .setFontWeight("bold")
              .setFontColor("#FFFFFF")
              .setBackground("#5C82AD")
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");
        });
        
        continue; // Skip formatting if metadata is identical
      }
      sheet.clear();
      sheet.clearFormats();
    }
    
    // 1. Header Banner
    sheet.getRange("A1:Q1").merge();
    var bannerCell = sheet.getRange("A1");
    bannerCell.setValue("আল কারিম কম্পিউটার এন্ড এমব্রয়ডারি গার্মেন্টস - গ্রাহক খতিয়ান");
    bannerCell.setFontColor("#FFFFFF")
              .setBackground("#047857")
              .setFontFamily(fontName)
              .setFontSize(14)
              .setFontWeight("bold")
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");
    sheet.setRowHeight(1, 40);
    
    // 2. Metadata Block (Row 2-4) with merged cells to prevent column width issues
    sheet.getRange("A2:C2").merge();
    sheet.getRange("A2").setValue("গ্রাহকের নাম (Name):").setFontWeight("bold").setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    sheet.getRange("D2:K2").merge();
    sheet.getRange("D2").setValue(partyName).setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    
    sheet.getRange("L2:N2").merge();
    sheet.getRange("L2").setValue("লেজার পৃষ্ঠা (Page):").setFontWeight("bold").setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("right");
    sheet.getRange("O2:Q2").merge();
    sheet.getRange("O2").setValue(ledgerPage).setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    
    sheet.getRange("A3:C3").merge();
    sheet.getRange("A3").setValue("ঠিকানা (Address):").setFontWeight("bold").setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    sheet.getRange("D3:K3").merge();
    sheet.getRange("D3").setValue(address).setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    
    sheet.getRange("L3:N3").merge();
    sheet.getRange("L3").setValue("ফোন নম্বর (Phone):").setFontWeight("bold").setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("right");
    sheet.getRange("O3:Q3").merge();
    sheet.getRange("O3").setValue("N/A").setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    
    sheet.getRange("A4:C4").merge();
    sheet.getRange("A4").setValue("মন্তব্য (Notes):").setFontWeight("bold").setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    sheet.getRange("D4:K4").merge();
    sheet.getRange("D4").setValue(notes).setFontFamily(fontName).setFontSize(10).setHorizontalAlignment("left");
    
    // 3. KPI Summary cards (Row 5-6)
    // Card 1: Total Sales (A5:E6)
    sheet.getRange("A5:E5").merge();
    sheet.getRange("A6:E6").merge();
    sheet.getRange("A5").setValue("সর্বমোট বিক্রয় (Total Purchases)")
         .setFontFamily(fontName)
         .setFontSize(9)
         .setFontColor("#555555")
         .setHorizontalAlignment("center");
         
    var card1Val = sheet.getRange("A6");
    card1Val.setFormula("=SUM(K9:K)")
            .setFontFamily(fontName)
            .setFontSize(14)
            .setFontWeight("bold")
            .setFontColor("#366092")
            .setHorizontalAlignment("center")
            .setNumberFormat("#,##0");
            
    // Card 2: Total Payments (G5:K6)
    sheet.getRange("G5:K5").merge();
    sheet.getRange("G6:K6").merge();
    sheet.getRange("G5").setValue("সর্বমোট আদায় (Total Payments)")
         .setFontFamily(fontName)
         .setFontSize(9)
         .setFontColor("#555555")
         .setHorizontalAlignment("center");
         
    var card2Val = sheet.getRange("G6");
    card2Val.setFormula("=SUM(P9:P)")
            .setFontFamily(fontName)
            .setFontSize(14)
            .setFontWeight("bold")
            .setFontColor("#2E7D32")
            .setHorizontalAlignment("center")
            .setNumberFormat("#,##0");
            
    // Card 3: Outstanding (M5:Q6)
    sheet.getRange("M5:Q5").merge();
    sheet.getRange("M6:Q6").merge();
    sheet.getRange("M5").setValue("অবশিষ্ট বকেয়া (Outstanding)")
         .setFontFamily(fontName)
         .setFontSize(9)
         .setFontColor("#555555")
         .setHorizontalAlignment("center");
         
    var card3Val = sheet.getRange("M6");
    card3Val.setFormula("=A6-G6")
            .setFontFamily(fontName)
            .setFontSize(14)
            .setFontWeight("bold")
            .setFontColor("#C62828")
            .setHorizontalAlignment("center")
            .setNumberFormat("#,##0");
            
    // Apply styling to KPI cards
    var kpiRanges = [sheet.getRange("A5:E6"), sheet.getRange("G5:K6"), sheet.getRange("M5:Q6")];
    kpiRanges.forEach(function(rng) {
      rng.setBackground("#F2F5F8")
         .setBorder(true, true, true, true, false, false, "#B0C4DE", SpreadsheetApp.BorderStyle.SOLID);
    });
    
    // 4. Table Headers (Row 7)
    sheet.getRange("A7:L7").merge();
    sheet.getRange("A7").setValue("ডেবিট এন্ট্রি সমূহ (Debit - Bills/Sales)")
         .setFontFamily(fontName)
         .setFontSize(11)
         .setFontWeight("bold")
         .setFontColor("#FFFFFF")
         .setBackground("#366092")
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
         
    sheet.getRange("N7:Q7").merge();
    sheet.getRange("N7").setValue("ক্রেডিট এন্ট্রি সমূহ (Credit - Payments)")
         .setFontFamily(fontName)
         .setFontSize(11)
         .setFontWeight("bold")
         .setFontColor("#FFFFFF")
         .setBackground("#2E7D32")
         .setHorizontalAlignment("center")
         .setVerticalAlignment("middle");
    sheet.setRowHeight(7, 24);
    
    // Column Sub-headers (Row 8)
    var debitHeaders = ["No", "Date", "Details (বিঃ কাঃ)", "Description", "Size", "Model", "Bill No", "Qty", "Rate", "মোট", "সর্বশেষ বিল", "Remarks"];
    var creditHeaders = ["No", "Date", "Amount", "Remarks"];
    
    debitHeaders.forEach(function(headerText, index) {
      var cell = sheet.getRange(8, index + 1);
      cell.setValue(headerText)
          .setFontFamily(fontName)
          .setFontSize(10)
          .setFontWeight("bold")
          .setFontColor("#FFFFFF")
          .setBackground("#5C82AD")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
    });
    
    creditHeaders.forEach(function(headerText, index) {
      var cell = sheet.getRange(8, index + 14); // starts from Column N (14)
      cell.setValue(headerText)
          .setFontFamily(fontName)
          .setFontSize(10)
          .setFontWeight("bold")
          .setFontColor("#FFFFFF")
          .setBackground("#4E9F5D")
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
    });
    sheet.setRowHeight(8, 24);
    
    // 5. Clear old data and write static values
    if (sheet.getLastRow() >= 9) {
      sheet.getRange(9, 1, sheet.getLastRow() - 8, 12).clearContent();
      sheet.getRange(9, 14, sheet.getLastRow() - 8, 4).clearContent();
    }
    
    var debits = [];
    var debitSheet = ss.getSheetByName("Debit_Transactions");
    if (debitSheet) {
      var debitData = debitSheet.getDataRange().getValues();
      for (var d = 1; d < debitData.length; d++) {
        if (String(debitData[d][0]).trim() === partyName) {
          var rowNum = 9 + debits.length;
          debits.push([
            debitData[d][2], // No
            debitData[d][3], // Date
            debitData[d][4], // Details
            debitData[d][5], // Description
            debitData[d][6], // Size
            debitData[d][7], // Model
            debitData[d][8], // Bill No
            debitData[d][9], // Qty
            debitData[d][10], // Rate
            "=H" + rowNum + "*I" + rowNum, // মোট (previously PD)
            "=SUM(J$9:J" + rowNum + ")", // সর্বশেষ বিল (previously Total)
            debitData[d][13]  // Remarks
          ]);
        }
      }
    }
    if (debits.length > 0) {
      sheet.getRange(9, 1, debits.length, 12).setValues(debits);
    }
    
    var credits = [];
    var creditSheet = ss.getSheetByName("Credit_Transactions");
    if (creditSheet) {
      var creditData = creditSheet.getDataRange().getValues();
      for (var cr = 1; cr < creditData.length; cr++) {
        if (String(creditData[cr][0]).trim() === partyName) {
          credits.push([
            creditData[cr][2], // No
            creditData[cr][3], // Date
            creditData[cr][4], // Amount
            creditData[cr][5]  // Remarks
          ]);
        }
      }
    }
    if (credits.length > 0) {
      sheet.getRange(9, 14, credits.length, 4).setValues(credits);
    }
    
    // Format dynamic formula data columns (Row 9:1000) to be centered, top-aligned, and wrapped
    var dataRange = sheet.getRange("A9:Q1000");
    dataRange.setHorizontalAlignment("center")
             .setVerticalAlignment("top")
             .setWrap(true);
             
    sheet.getRange(9, 8, 992, 4).setNumberFormat("#,##0"); // Columns H to K (Qty, Rate, PD, Total)
    sheet.getRange("P9:P1000").setNumberFormat("#,##0");
    
    // Freeze rows 1-8 so the dashboard header remains visible when scrolling down
    sheet.setFrozenRows(8);
    
    // Set column widths so that everything fits on one screen
    for (var col = 1; col <= 17; col++) {
      if (col === 1) {
        sheet.setColumnWidth(col, 25); // Column A is 0.25 inches (25px)
      } else if (col === 5) {
        sheet.setColumnWidth(col, 45); // Column E (Size) is set to 45px
      } else if (col === 13) {
        sheet.setColumnWidth(col, 25); // Column M is the blank separator column
      } else {
        sheet.setColumnWidth(col, 75); // ~1 inch (75 pixels)
      }
    }
  }
  
  // Create business dashboard and monthly pivot sheets
  createDashboardSheet(ss);
  createPivotTableSheet(ss);
}

function sortSheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Sorting sheets... please wait.', 'Sort Status', -1);
  sortSheets(ss);
  ss.toast('Sheets sorted successfully!', 'Sort Status', 5);
}

function sortSheets(ss) {
  var sheets = ss.getSheets();
  
  var mainList = new Array(5); // index 0: Dashboard, index 1: Pivot, index 2: Client Index, index 3: Debit, index 4: Credit
  var clientList = [];
  
  sheets.forEach(function(sheet) {
    var rawName = sheet.getName();
    var normName = rawName.toLowerCase().replace(/[\s_-]/g, "");
    
    if (normName.indexOf("dashboard") !== -1 || normName.indexOf("ড্যাশবোর্ড") !== -1) {
      mainList[0] = {name: rawName, sheet: sheet};
    } else if (normName.indexOf("pivot") !== -1 || normName.indexOf("পিভট") !== -1) {
      mainList[1] = {name: rawName, sheet: sheet};
    } else if (normName.indexOf("clientindex") !== -1 || normName.indexOf("গ্রাহকসূচী") !== -1) {
      mainList[2] = {name: rawName, sheet: sheet};
    } else if (normName.indexOf("debit") !== -1 || normName.indexOf("ডেবিট") !== -1) {
      mainList[3] = {name: rawName, sheet: sheet};
    } else if (normName.indexOf("credit") !== -1 || normName.indexOf("ক্রেডিট") !== -1) {
      mainList[4] = {name: rawName, sheet: sheet};
    } else {
      clientList.push({name: rawName, sheet: sheet});
    }
  });
  
  // Filter out any missing main sheets from the list
  var cleanMainList = [];
  for (var i = 0; i < mainList.length; i++) {
    if (mainList[i]) {
      cleanMainList.push(mainList[i]);
    }
  }
  
  // Sort client sheets numerically based on the leading number in their name
  clientList.sort(function(a, b) {
    var matchA = a.name.match(/^(\d+)\./);
    var matchB = b.name.match(/^(\d+)\./);
    var numA = matchA ? parseInt(matchA[1], 10) : 9999;
    var numB = matchB ? parseInt(matchB[1], 10) : 9999;
    return numA - numB;
  });
  
  var sortedList = cleanMainList.concat(clientList);
  
  // Move sheets to their sorted positions
  for (var i = 0; i < sortedList.length; i++) {
    var targetSheet = sortedList[i].sheet;
    ss.setActiveSheet(targetSheet);
    ss.moveActiveSheet(i + 1);
  }
}

function syncNamesInTransactions(ss) {
  var indexSheet = ss.getSheetByName("Client_Index");
  if (!indexSheet) return;
  var indexData = indexSheet.getDataRange().getValues();
  
  // Create a mapping from page number (both even page and odd page) to the party name
  var pageToNameMap = {};
  for (var i = 1; i < indexData.length; i++) {
    var partyName = String(indexData[i][1] || "").trim();
    var ledgerPage = parseInt(indexData[i][3], 10);
    if (!partyName || isNaN(ledgerPage)) continue;
    
    pageToNameMap[ledgerPage] = partyName;      // Left (Debit) page
    pageToNameMap[ledgerPage + 1] = partyName;  // Right (Credit) page
  }
  
  // 1. Sync names in Debit_Transactions
  var debitSheet = ss.getSheetByName("Debit_Transactions");
  if (debitSheet) {
    var debitRange = debitSheet.getDataRange();
    var debitData = debitRange.getValues();
    var updated = false;
    for (var r = 1; r < debitData.length; r++) {
      var rowPage = parseInt(debitData[r][1], 10); // Column B is ledger_page
      if (isNaN(rowPage)) continue;
      
      var correctName = pageToNameMap[rowPage];
      if (correctName && debitData[r][0] !== correctName) { // Column A is party_name
        debitData[r][0] = correctName;
        updated = true;
      }
    }
    if (updated) {
      debitRange.setValues(debitData);
    }
  }
  
  // 2. Sync names in Credit_Transactions
  var creditSheet = ss.getSheetByName("Credit_Transactions");
  if (creditSheet) {
    var creditRange = creditSheet.getDataRange();
    var creditData = creditRange.getValues();
    var updated = false;
    for (var r = 1; r < creditData.length; r++) {
      var rowPage = parseInt(creditData[r][1], 10); // Column B is ledger_page
      if (isNaN(rowPage)) continue;
      
      var correctName = pageToNameMap[rowPage];
      if (correctName && creditData[r][0] !== correctName) { // Column A is party_name
        creditData[r][0] = correctName;
        updated = true;
      }
    }
    if (updated) {
      creditRange.setValues(creditData);
    }
  }
}

// Automatically cleans decimal strings inside the raw transaction tables and parses them to numbers
function cleanDatabaseSheets(ss) {
  var targetSheets = ["Debit_Transactions", "Credit_Transactions"];
  targetSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    var range = sheet.getDataRange();
    var values = range.getValues();
    var numRows = values.length;
    if (numRows <= 1) return;
    var numCols = values[0].length;
    
    var updated = false;
    for (var r = 1; r < numRows; r++) {
      for (var c = 0; c < numCols; c++) {
        var val = values[r][c];
        if (val !== null && val !== undefined && val !== "") {
          var str = String(val).trim().replace(/,/g, '');
          if (str && str !== '〃' && str !== '"') {
            var parsed = parseFloat(str);
            if (!isNaN(parsed)) {
              var rounded = Math.round(parsed);
              if (values[r][c] !== rounded) {
                values[r][c] = rounded;
                updated = true;
              }
            }
          }
        }
      }
    }
    if (updated) {
      range.setValues(values);
    }
  });
}

// Generates a beautiful executive dashboard with overall metrics and client balances
function createDashboardSheet(ss) {
  var sheet = ss.getSheetByName("Dashboard");
  if (sheet) {
    sheet.clear();
    // remove existing charts
    var charts = sheet.getCharts();
    for (var i = 0; i < charts.length; i++) {
      sheet.removeChart(charts[i]);
    }
  } else {
    sheet = ss.insertSheet("Dashboard");
  }
  
  // Ensure enough columns exist for the chart to anchor without fallback
  if (sheet.getMaxColumns() < 15) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 15 - sheet.getMaxColumns());
  }
  
  var fontName = "Segoe UI";
  
  // 1. Title Banner
  sheet.getRange("A1:G1").merge()
       .setValue("আল কারিম কম্পিউটার এন্ড এমব্রয়ডারি গার্মেন্টস - ব্যবসায়িক ড্যাশবোর্ড (Global Dashboard)")
       .setFontFamily(fontName)
       .setFontSize(16)
       .setFontWeight("bold")
       .setFontColor("#FFFFFF")
       .setBackground("#047857")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("center");
  sheet.setRowHeight(1, 50);
  
  // 2. KPI Cards
  sheet.setRowHeight(3, 24);
  sheet.setRowHeight(4, 96);
  
  // Total Sales (A3:B4)
  sheet.getRange("A3:B3").merge().setValue("সর্বমোট বিক্রয় (Total Sales)")
       .setFontFamily(fontName).setFontSize(9).setFontColor("#555555").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("A4:B4").merge().setFormula("=SUM(Debit_Transactions!M2:M)")
       .setFontFamily(fontName).setFontSize(22).setFontWeight("bold").setFontColor("#366092")
       .setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("#,##0");
       
  // Total Collected (C3:D3) -> D3:E4
  sheet.getRange("D3:E3").merge().setValue("সর্বমোট আদায় (Total Payments)")
       .setFontFamily(fontName).setFontSize(9).setFontColor("#555555").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("D4:E4").merge().setFormula("=SUM(Credit_Transactions!E2:E)")
       .setFontFamily(fontName).setFontSize(22).setFontWeight("bold").setFontColor("#2E7D32")
       .setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("#,##0");
       
  // Outstanding (G3:H4)
  sheet.getRange("G3:H3").merge().setValue("অবশিষ্ট বকেয়া (Outstanding)")
       .setFontFamily(fontName).setFontSize(9).setFontColor("#555555").setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("G4:H4").merge().setFormula("=A4-D4")
       .setFontFamily(fontName).setFontSize(22).setFontWeight("bold").setFontColor("#C62828")
       .setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("#,##0");
       
  // Format KPI Card Borders
  var kpiRanges = [sheet.getRange("A3:B4"), sheet.getRange("D3:E4"), sheet.getRange("G3:H4")];
  kpiRanges.forEach(function(rng) {
    rng.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    rng.setBackground("#f8fafc");
  });
  
  // 3. Client Summary Table Header
  var tableHeaders = ["Sl", "গ্রাহকের নাম (Client Name)", "ঠিকানা (Address)", "লেজার পৃষ্ঠা (Page)", "মোট ক্রয় (Sales)", "মোট পরিশোধ (Payments)", "বকেয়া (Balance)"];
  var headerRange = sheet.getRange(6, 1, 1, 7);
  headerRange.setValues([tableHeaders])
             .setFontFamily(fontName)
             .setFontSize(10)
             .setFontWeight("bold")
             .setFontColor("#FFFFFF")
             .setBackground("#047857")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("center");
  sheet.setRowHeight(6, 30);
  sheet.setFrozenRows(6);
  
  // 4. Data Rows
  var indexSheet = ss.getSheetByName("Client_Index");
  if (indexSheet) {
    var indexData = indexSheet.getDataRange().getValues();
    var numClients = indexData.length - 1;
    if (numClients > 0) {
      var rows = [];
      for (var i = 1; i <= numClients; i++) {
        var r = i + 6; // starts at row 7
        rows.push([
          i,
          "=Client_Index!B" + (i + 1),
          "=Client_Index!C" + (i + 1),
          "=Client_Index!D" + (i + 1),
          "=SUMIF(Debit_Transactions!A:A, B" + r + ", Debit_Transactions!M:M)",
          "=SUMIF(Credit_Transactions!A:A, B" + r + ", Credit_Transactions!E:E)",
          "=E" + r + "-F" + r
        ]);
      }
      
      var dataRange = sheet.getRange(7, 1, numClients, 7);
      dataRange.setFormulas(rows)
               .setFontFamily(fontName)
               .setFontSize(10)
               .setVerticalAlignment("center")
               .setBorder(true, true, true, true, true, true, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
               
      // Alignment
      sheet.getRange(7, 1, numClients, 1).setHorizontalAlignment("center"); // Sl
      sheet.getRange(7, 2, numClients, 2).setHorizontalAlignment("left");   // Name, Address
      sheet.getRange(7, 4, numClients, 1).setHorizontalAlignment("center"); // Page
      sheet.getRange(7, 5, numClients, 3).setHorizontalAlignment("right").setNumberFormat("#,##0"); // Sales, Payments, Balance
      
      // Auto-striping
      for (var i = 0; i < numClients; i++) {
        var rowRange = sheet.getRange(7 + i, 1, 1, 7);
        if (i % 2 === 1) {
          rowRange.setBackground("#f8fafc");
        } else {
          rowRange.setBackground("#ffffff");
        }
      }
      
      // Total Row
      var totalRow = 7 + numClients;
      sheet.getRange(totalRow, 1).setValue("");
      sheet.getRange(totalRow, 2).setValue("সর্বমোট (Total)").setFontWeight("bold").setHorizontalAlignment("left");
      sheet.getRange(totalRow, 3, 1, 2).merge();
      sheet.getRange(totalRow, 5).setFormula("=SUM(E7:E" + (totalRow-1) + ")").setFontWeight("bold").setNumberFormat("#,##0").setHorizontalAlignment("right");
      sheet.getRange(totalRow, 6).setFormula("=SUM(F7:F" + (totalRow-1) + ")").setFontWeight("bold").setNumberFormat("#,##0").setHorizontalAlignment("right");
      sheet.getRange(totalRow, 7).setFormula("=SUM(G7:G" + (totalRow-1) + ")").setFontWeight("bold").setNumberFormat("#,##0").setHorizontalAlignment("right");
      sheet.getRange(totalRow, 1, 1, 7)
           .setFontFamily(fontName)
           .setFontSize(10)
           .setBackground("#e2e8f0")
           .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
           
      // Add a chart
      var chart = sheet.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(sheet.getRange("B6:B" + (totalRow-1))) // Client Names
          .addRange(sheet.getRange("G6:G" + (totalRow-1))) // Outstanding Balances
          .setPosition(3, 8, 0, 0) // Anchor at H3 (inside frozen pane)
          .setOption("title", "গ্রাহকদের বকেয়া হিসাব (Client Outstanding Balances)")
          .setOption("colors", ["#C62828"])
          .setOption("width", 650)
          .setOption("height", 118) // Fits completely inside frozen Rows 3 & 4 (total height 120px)
          .build();
      sheet.insertChart(chart);
      SpreadsheetApp.flush(); // Commit changes immediately
    }
  }
  
  // Set Column Widths
  sheet.setColumnWidth(1, 35);  // Sl
  sheet.setColumnWidth(2, 180); // Name
  sheet.setColumnWidth(3, 180); // Address
  sheet.setColumnWidth(4, 75);  // Page
  sheet.setColumnWidth(5, 90);  // Sales
  sheet.setColumnWidth(6, 90);  // Payments
  sheet.setColumnWidth(7, 90);  // Balance
  sheet.setColumnWidth(8, 25);  // blank spacer
}

// Generates an interactive pivot table sheet grouping transactions
function createPivotTableSheet(ss) {
  var pivotSheet = ss.getSheetByName("Monthly_Pivot");
  if (pivotSheet) {
    ss.deleteSheet(pivotSheet);
  }
  pivotSheet = ss.insertSheet("Monthly_Pivot");
  
  var debitSheet = ss.getSheetByName("Debit_Transactions");
  if (!debitSheet) return;
  
  var sourceRange = debitSheet.getDataRange();
  var pivotTable = pivotSheet.getRange("A1").createPivotTable(sourceRange);
  
  // Row Group: Client Name (Column A, index 1)
  var rowGroup = pivotTable.addRowGroup(1);
  rowGroup.showTotals(true);
  
  // Column Group: Details (Column E, index 5)
  var colGroup = pivotTable.addColumnGroup(5);
  colGroup.showTotals(true);
  
  // Values: Total (Column M, index 13)
  pivotTable.addPivotValue(13, SpreadsheetApp.PivotTableSummarizeFunction.SUM).setDisplayName("Total Sales");
}

// ==============================================
// LIVE TWO-WAY SYNCHRONIZATION ON EDIT
// ==============================================
function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var sheetName = sheet.getName();
  var row = range.getRow();
  var col = range.getColumn();
  var newValue = e.value;
  var ss = e.source;

  // Prevent running on header/metadata rows
  if (row < 2) return;

  // Case 1: Edit on Debit_Transactions
  if (sheetName === "Debit_Transactions") {
    if (col < 3) return; // Skip headers/metadata (Client Name, Ledger Page)
    var clientName = sheet.getRange(row, 1).getValue();
    var rowNo = sheet.getRange(row, 3).getValue();
    if (!clientName || !rowNo) return;
    
    // Auto-calculate "মোট" (Column 12) if Qty (10) or Rate (11) is edited
    if (col === 10 || col === 11) {
      var qty = parseFloat(sheet.getRange(row, 10).getValue()) || 0;
      var rate = parseFloat(sheet.getRange(row, 11).getValue()) || 0;
      var currentPdCell = sheet.getRange(row, 12);
      var calculatedPd = Math.round(qty * rate);
      if (currentPdCell.getValue() !== calculatedPd) {
        currentPdCell.setValue(calculatedPd);
      }
    }
    
    // Recalculate running totals in Column 13 ("সর্বশেষ বিল") for this client
    recalculateRunningTotalsInDebitTransactions(sheet, clientName);
    
    // Update the client sheet
    var clientSheet = getClientSheetByName(ss, clientName);
    if (clientSheet) {
      syncClientSheets(ss, clientName);
    }
  }
  // Case 2: Edit on Credit_Transactions
  else if (sheetName === "Credit_Transactions") {
    if (col < 3) return;
    var clientName = sheet.getRange(row, 1).getValue();
    var rowNo = sheet.getRange(row, 3).getValue();
    if (!clientName || !rowNo) return;
    
    var clientSheet = getClientSheetByName(ss, clientName);
    if (clientSheet) {
      var clientRow = findRowInClientSheet(clientSheet, 14, rowNo);
      if (clientRow !== -1) {
        var cell = clientSheet.getRange(clientRow, col + 11);
        if (cell.getValue() !== newValue) {
          cell.setValue(newValue);
        }
      }
    }
  }
  // Case 3: Edit on Client Sheet
  else if (/^\d+\.\s*/.test(sheetName) || /^P\d+\s*-/.test(sheetName)) {
    if (row < 9) return; // Data starts at row 9
    var clientName = sheet.getRange("D2").getValue();
    if (!clientName) return;

    if (col >= 1 && col <= 12) { // Debit Edit
      var rowNo = sheet.getRange(row, 1).getValue();
      if (!rowNo) return;
      var debitSheet = ss.getSheetByName("Debit_Transactions");
      if (debitSheet) {
        var debitRow = findRowInTransactionSheet(debitSheet, clientName, rowNo);
        if (debitRow !== -1) {
          var targetCol = col + 2;
          var cell = debitSheet.getRange(debitRow, targetCol);
          if (cell.getValue() !== newValue) {
            cell.setValue(newValue);
            
            // If Qty (8) or Rate (9) was edited in client sheet, we let the Case 1 trigger handle
            // the pd/total recalculation in Debit_Transactions, which will then sync back!
          }
        }
      }
    } else if (col >= 14 && col <= 17) { // Credit Edit
      var rowNo = sheet.getRange(row, 14).getValue();
      if (!rowNo) return;
      var creditSheet = ss.getSheetByName("Credit_Transactions");
      if (creditSheet) {
        var creditRow = findRowInTransactionSheet(creditSheet, clientName, rowNo);
        if (creditRow !== -1) {
          var cell = creditSheet.getRange(creditRow, col - 11);
          if (cell.getValue() !== newValue) {
            cell.setValue(newValue);
          }
        }
      }
    }
  }
}

// Helpers
function getClientSheetByName(ss, clientName) {
  var indexSheet = ss.getSheetByName("Client_Index");
  if (!indexSheet) return null;
  var data = indexSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === String(clientName).trim()) {
      var clientNo = String(data[i][0]).trim();
      var cleanName = String(clientName).replace(/[\\\/\?\:\*\[\]]/g, "");
      var name1 = clientNo + ". " + cleanName;
      if (name1.length > 31) name1 = name1.substring(0, 31);
      var s = ss.getSheetByName(name1);
      if (s) return s;
    }
  }
  return null;
}

function findRowInClientSheet(sheet, colIndex, rowNo) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 9) return -1;
  var data = sheet.getRange(9, colIndex, lastRow - 8, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(rowNo).trim()) {
      return i + 9;
    }
  }
  return -1;
}

function findRowInTransactionSheet(sheet, clientName, rowNo) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(clientName).trim() && String(data[i][2]).trim() === String(rowNo).trim()) {
      return i + 1;
    }
  }
  return -1;
}

function recalculateRunningTotalsInDebitTransactions(sheet, clientName) {
  var data = sheet.getDataRange().getValues();
  var runningTotal = 0;
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(clientName).trim()) {
      var rowTotal = parseFloat(data[i][11]) || 0; // Column L (12th)
      runningTotal += rowTotal;
      var cell = sheet.getRange(i + 1, 13); // Column M (13th)
      if (cell.getValue() !== runningTotal) {
        cell.setValue(runningTotal);
      }
    }
  }
}
