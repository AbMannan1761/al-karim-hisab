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
    
    syncClientSheets(ss);
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
      var oldSheetName = "P" + data[rowIdx][3] + " - " + oldCleanName;
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
  // First propagate any name updates from Client_Index to Debit_Transactions and Credit_Transactions
  syncNamesInTransactions(ss);
  
  var indexSheet = ss.getSheetByName("Client_Index");
  if (!indexSheet) return;
  var indexData = indexSheet.getDataRange().getValues();
  
  var fontName = "Segoe UI";
  
  // 1. Collect all active sheet names based on Client_Index
  var activeSheetNames = {};
  for (var i = 1; i < indexData.length; i++) {
    var partyName = String(indexData[i][1] || "").trim();
    var ledgerPage = indexData[i][3];
    if (!partyName || !ledgerPage) continue;
    
    var cleanName = partyName.replace(/[\\\/\?\:\*\[\]]/g, "");
    var sheetName = "P" + ledgerPage + " - " + cleanName;
    if (sheetName.length > 31) {
      sheetName = sheetName.substring(0, 31);
    }
    activeSheetNames[sheetName] = true;
  }
  
  // 2. Clean up any orphan client sheets (e.g., if renamed)
  var sheets = ss.getSheets();
  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    // If it is a client sheet (starts with P followed by number)
    if (/^P\d+\s*-/.test(name)) {
      if (!activeSheetNames[name]) {
        ss.deleteSheet(sheet);
      }
    }
  });
  
  // 3. Process sheets
  for (var i = 1; i < indexData.length; i++) {
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
    var sheetName = "P" + ledgerPage + " - " + cleanName;
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
          if (col === 13) {
            sheet.setColumnWidth(col, 25); // separator column
          } else {
            sheet.setColumnWidth(col, 75); // approx. 1 inch
          }
        }
        
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
            .setNumberFormat("#,##0.00");
            
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
            .setNumberFormat("#,##0.00");
            
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
            .setNumberFormat("#,##0.00");
            
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
    var debitHeaders = ["No", "Date", "Details (বিঃ কাঃ)", "Description", "Size", "Model", "PD", "Bill No", "Qty", "Rate", "Total", "Remarks"];
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
    
    // 5. Insert filter formulas in Row 9
    sheet.getRange("A9").setFormula('=IFERROR(FILTER(Debit_Transactions!C2:N, Debit_Transactions!A2:A = D2), "")');
    sheet.getRange("N9").setFormula('=IFERROR(FILTER(Credit_Transactions!C2:F, Credit_Transactions!A2:A = D2), "")');
    
    // Format dynamic formula data columns (Row 9:1000)
    sheet.getRange("I9:K1000").setNumberFormat("#,##0.00").setHorizontalAlignment("right");
    sheet.getRange("P9:P1000").setNumberFormat("#,##0.00").setHorizontalAlignment("right");
    sheet.getRange("A9:H1000").setHorizontalAlignment("left");
    sheet.getRange("L9:L1000").setHorizontalAlignment("left");
    sheet.getRange("N9:O1000").setHorizontalAlignment("left");
    sheet.getRange("Q9:Q1000").setHorizontalAlignment("left");
    
    // Freeze rows 1-8 so the dashboard header remains visible when scrolling down
    sheet.setFrozenRows(8);
    
    // Set column widths to 75px (~1 inch) so that everything fits on one screen
    for (var col = 1; col <= 17; col++) {
      if (col === 13) {
        sheet.setColumnWidth(col, 25); // Column M is the blank separator column
      } else {
        sheet.setColumnWidth(col, 75); // ~1 inch (75 pixels)
      }
    }
  }
}

function sortSheetsMenu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Sorting sheets... please wait.', 'Sort Status', -1);
  sortSheets(ss);
  ss.toast('Sheets sorted successfully!', 'Sort Status', 5);
}

function sortSheets(ss) {
  var sheets = ss.getSheets();
  
  var mainList = new Array(3); // index 0: Client Index, index 1: Debit, index 2: Credit
  var clientList = [];
  
  sheets.forEach(function(sheet) {
    var rawName = sheet.getName();
    var normName = rawName.toLowerCase().replace(/[\s_-]/g, "");
    
    if (normName.indexOf("clientindex") !== -1 || normName.indexOf("গ্রাহকসূচী") !== -1) {
      mainList[0] = {name: rawName, sheet: sheet};
    } else if (normName.indexOf("debit") !== -1 || normName.indexOf("ডেবিট") !== -1) {
      mainList[1] = {name: rawName, sheet: sheet};
    } else if (normName.indexOf("credit") !== -1 || normName.indexOf("ক্রেডিট") !== -1) {
      mainList[2] = {name: rawName, sheet: sheet};
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
  
  // Sort client sheets alphabetically by name (Bengali locales supported)
  clientList.sort(function(a, b) {
    return a.name.localeCompare(b.name, 'bn');
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
