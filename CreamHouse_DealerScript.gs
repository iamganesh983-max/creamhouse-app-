// ================================================================
//  CREAM HOUSE — Dealer Order Apps Script
//  Add this to the SAME Apps Script project as your production app
//  OR create a new separate project linked to the same Sheet
//
//  SHEETS REQUIRED (auto-created on first use):
//  - "Dealers"       — your dealer list with PINs
//  - "Dealer Orders" — all incoming orders
// ================================================================

// ── GET: handles login + order history ──────────────────────────
function doGet_dealer(e) {
  var action = e.parameter.action;

  if (action === "login") {
    return handleLogin(e.parameter.name, e.parameter.pin);
  }

  if (action === "history") {
    return handleHistory(e.parameter.dealerId);
  }

  return ContentService
    .createTextOutput(JSON.stringify({success:true, message:"Dealer API ready"}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── LOGIN: validate dealer name + PIN ───────────────────────────
function handleLogin(name, pin) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Dealers");

    if (!sheet) {
      // Auto-create Dealers sheet with sample data
      sheet = ss.insertSheet("Dealers");
      setupDealersSheet(sheet);
    }

    var data = sheet.getDataRange().getValues();
    // Columns: A=DealerID, B=Name, C=PIN, D=Outstanding, E=Active
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowName = String(row[1]||"").trim().toLowerCase();
      var rowPin  = String(row[2]||"").trim();
      var rowActive = row[4];

      if (rowName === name.trim().toLowerCase() && rowPin === pin.trim()) {
        if (rowActive === false || rowActive === "No") {
          return json({success:false, error:"Your account is inactive. Contact Cream House."});
        }
        return json({
          success: true,
          dealer: {
            name:         row[1],
            id:           row[0],
            outstanding:  row[3] || 0,
            totalOrders:  getDealerOrderCount(row[0]),
          }
        });
      }
    }

    return json({success:false, error:"Incorrect name or PIN. Please try again."});

  } catch(err) {
    return json({success:false, error:"Server error: " + err.message});
  }
}

// ── ORDER HISTORY: last 10 orders for a dealer ──────────────────
function handleHistory(dealerId) {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Dealer Orders");

    if (!sheet || sheet.getLastRow() < 2) {
      return json({success:true, orders:[]});
    }

    var data = sheet.getDataRange().getValues();
    // Columns: A=OrderNo, B=Date, C=Time, D=DealerID, E=DealerName,
    //          F=Items, G=Total, H=Notes, I=Status

    var orders = [];
    for (var i = data.length - 1; i >= 1 && orders.length < 10; i--) {
      var row = data[i];
      if (String(row[3]||"") === String(dealerId||"") ||
          String(row[4]||"").toLowerCase() === String(dealerId||"").toLowerCase()) {
        orders.push({
          orderNo: row[0],
          date:    row[1],
          time:    row[2],
          items:   row[5],
          total:   row[6],
          notes:   row[7],
          status:  row[8] || "Pending"
        });
      }
    }

    return json({success:true, orders:orders});

  } catch(err) {
    return json({success:true, orders:[], error:err.message});
  }
}

// ── POST: save incoming dealer order ────────────────────────────
function doPost_dealer(e) {
  try {
    var data;
    if (e.parameter && e.parameter.payload) {
      data = JSON.parse(e.parameter.payload);
    } else if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      return json({success:false, error:"No data received"});
    }

    if (data.type === "dealer_order") {
      return saveDealerOrder(data.order);
    }

    return json({success:false, error:"Unknown type: " + data.type});

  } catch(err) {
    return json({success:false, error:err.message});
  }
}

// ── SAVE ORDER ───────────────────────────────────────────────────
function saveDealerOrder(order) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Dealer Orders");

  if (!sheet) {
    sheet = ss.insertSheet("Dealer Orders");
    setupOrdersSheet(sheet);
  }

  // Generate order number
  var lastRow = sheet.getLastRow();
  var orderNo = "ORD-" + String(lastRow).padStart(4, "0");

  // Format items as readable string
  var itemsStr = order.items.map(function(it){
    return it.name + " ₹" + it.mrp + " ×" + it.qty;
  }).join(", ");

  sheet.appendRow([
    orderNo,
    order.date,
    order.time,
    order.dealerId || "",
    order.dealer   || "",
    itemsStr,
    order.total    || 0,
    order.notes    || "",
    "Pending",          // Status — you update this to Confirmed/Delivered
    new Date()          // Timestamp
  ]);

  // Update dealer's outstanding amount in Dealers sheet
  updateDealerOutstanding(order.dealer, order.total);

  return json({success:true, orderNo:orderNo});
}

// ── UPDATE OUTSTANDING ───────────────────────────────────────────
function updateDealerOutstanding(dealerName, amount) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dealers");
    if (!sheet) return;
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]||"").toLowerCase() === String(dealerName||"").toLowerCase()) {
        var current = parseFloat(data[i][3]) || 0;
        sheet.getRange(i+1, 4).setValue(current + amount);
        break;
      }
    }
  } catch(e) {}
}

// ── COUNT DEALER ORDERS ──────────────────────────────────────────
function getDealerOrderCount(dealerId) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dealer Orders");
    if (!sheet) return 0;
    var data = sheet.getDataRange().getValues();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][3]) === String(dealerId)) count++;
    }
    return count;
  } catch(e) { return 0; }
}

// ── SHEET SETUP ──────────────────────────────────────────────────
function setupDealersSheet(sheet) {
  sheet.getRange(1, 1, 1, 6).setValues([[
    "Dealer ID", "Name", "PIN", "Outstanding (₹)", "Active", "Phone"
  ]]);
  sheet.getRange("1:1").setBackground("#C0392B").setFontColor("#fff").setFontWeight("bold");

  // Sample dealers — REPLACE WITH YOUR ACTUAL DEALERS
  var samples = [
    ["DLR-001", "Ravi Kumar",   "1234", 0, "Yes", "9876543210"],
    ["DLR-002", "Suresh Babu",  "5678", 0, "Yes", "9876543211"],
    ["DLR-003", "Mahesh Rao",   "9012", 0, "Yes", "9876543212"],
  ];
  sheet.getRange(2, 1, samples.length, 6).setValues(samples);
  sheet.setFrozenRows(1);

  Logger.log("Dealers sheet created with 3 sample dealers.");
  Logger.log("Go to the Dealers sheet and update names, PINs and phone numbers.");
}

function setupOrdersSheet(sheet) {
  sheet.getRange(1, 1, 1, 10).setValues([[
    "Order No", "Date", "Time", "Dealer ID", "Dealer Name",
    "Items", "Total MRP (₹)", "Notes", "Status", "Timestamp"
  ]]);
  sheet.getRange("1:1").setBackground("#1D9E75").setFontColor("#fff").setFontWeight("bold");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(6, 300);
}

// ── HELPER ───────────────────────────────────────────────────────
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── HOW TO INTEGRATE WITH YOUR EXISTING SCRIPT ──────────────────
// Option A (recommended): Add to your existing CreamHouse_AppsScript.gs
// In your existing doGet(), add:
//   if (action === "login" || action === "history") return doGet_dealer(e);
// In your existing doPost(), add before the type check:
//   if (data.type === "dealer_order") return saveDealerOrder(data.order);
//
// Option B: Deploy this as a SEPARATE Web App on the same Sheet
// Then use a different URL in the dealer app.
