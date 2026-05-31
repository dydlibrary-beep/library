
function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// FIXED: Added missing headers to match columns used in formatting (18, 19, 20)
const ORDER_HEADERS = [
  "Order No",
  "Buyer",
  "Books",
  "Summary",
  "Total",
  "Discount %",
  "Discount ₹",
  "Courier",
  "Net Total",
  "Address",
  "Mobile",
  "WhatsApp",
  "Billable",
  "Status",
  "Order Date",
  "Restored",
  "PDF ID",
  "Track ID",
  "Track Status",
  "Payment Status"
];

const BOOK_HEADERS = [
  "Book No",
  "Book Name",
  "Language",
  "Publication",
  "Author",
  "Stock",
  "Price"
];
const BOOK_START_ROW = 7;
const ORDER_START_ROW = 5; // Data starts at Row 5 (Row 4 is Headers)
const STATUS = {
  PENDING: "⏳ Pending",
  PROGRESS: "🔄 In Progress",
  COMPLETE: "✅ Complete",
  CANCELLED: "❌ Cancelled"
};

function getSheets() {
  var ss = getSS();

  // FIXED: Case sensitivity matching
  var books = ss.getSheetByName("All Book List");
  var orders = ss.getSheetByName("Order Book List");

  if (!books || !orders) {
    throw new Error("Required sheets missing");
  }

  return {
    BOOKS: books,
    ORDERS: orders
  };
}
const ADMIN_PASSWORD = "Admin@123";


// =====================================================
// ON OPEN
// =====================================================

function onOpen() {
  // OPTIMIZATION: Removed flush() for faster load
  createMenus();
}

// =====================================================
// CREATE MENUS
// =====================================================

function createMenus() {
  var ui = SpreadsheetApp.getUi();

  // ================= STOCK MENU =================
  var stockMenu = ui.createMenu("📦 Stock")
    .addItem("🆕 Add New Book", "showBookForm")
    .addItem("🛠 Manage Stock", "adminShowManageBookForm")
    .addItem("📤 Import Books from StockData", "adminImportBooksFromSheet")
    .addItem("📥 Download Stock Report", "downloadStockReport");

  // ================= ORDER MENU =================
  var orderMenu = ui.createMenu("📚 Orders")
    .addItem("🆕 New Order", "showOrderForm")
    .addItem("👁 Manage Order", "openSelectedOrder")
    .addItem("📝 Update Order Details", "showUpdateOrderForm")
    .addItem("🧾 Print Bill", "printBillFromSelectedRow")
    .addItem("📲 Send Bill WhatsApp", "sendBillToWhatsApp");

  // ================= STATUS MENU =================
  var statusMenu = ui.createMenu("📌 Status")
    .addItem("⏳ Pending", "markPending")
    .addItem("🔄 In Progress", "markProgress")
    .addItem("✅ Complete", "markComplete")
    .addItem("❌ Cancelled", "markCancelled");

  // ================= ADMIN MENU =================
  var adminMenu = ui.createMenu("🛡 Admin")
    .addItem("🔑 Login", "openAdminMenu")
    .addItem("🚪 Logout", "adminLogout")
    .addSeparator()
    .addItem("🚀 Run Setup", "adminInitialSetup")
    .addItem("🔓 Unlock Editing", "adminUnlockEditing")
    .addItem("🔒 Lock Editing", "adminLockEditing")
    .addSeparator()
    .addItem("🧹 Cleanup", "cleanUp")
    .addItem("🗑 Reset Entire Sheet", "adminDeleteAll");

  // ================= TRACKING MENU =================
  var trackingMenu = ui.createMenu("🔍 Tracking")
    .addItem("📬 Set / Update Tracking ID", "showSetTrackingIdDialog")
    .addItem("🚚 Start Tracking Update", "updateAllTrackingStatuses");

  var paymentMenu = ui.createMenu("💳 Payment")
    .addItem("⏳ Pending", "markPaymentPending")
    .addItem("✅ Received", "markPaymentReceived");

  // ================= MAIN MENU =================
  ui.createMenu("🏢 DYD Library")
    .addSubMenu(stockMenu)
    .addSubMenu(orderMenu)
    .addSubMenu(trackingMenu)
    .addSubMenu(statusMenu)
    .addSubMenu(paymentMenu)
    .addSubMenu(adminMenu)
    .addSeparator()
    .addItem("🔄 Refresh", "dailyRefresh")
    .addToUi();
}

// =====================================================
// CHECK ADMIN
// =====================================================

function isAdminLoggedIn() {
  return (
    PropertiesService
      .getUserProperties()
      .getProperty("ADMIN_LOGGED_IN") === "true"
  );
}

// =====================================================
// REQUIRE ADMIN
// =====================================================

function requireAdmin() {
  if (!isAdminLoggedIn()) {
    throw new Error("🔒 Admin login required");
  }
}

// =====================================================
// ADMIN LOGIN
// =====================================================

function openAdminMenu() {
  var ui = SpreadsheetApp.getUi();

  // ALREADY LOGGED IN
  if (isAdminLoggedIn()) {
    ui.alert("✅ Already logged in");
    return;
  }

  var result = ui.prompt(
    "🔐 Admin Access",
    "Enter Admin Password",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  var password = result.getResponseText();

  if (password !== ADMIN_PASSWORD) {
    ui.alert("❌ Wrong Password");
    return;
  }

  // SAVE SESSION
  PropertiesService
    .getUserProperties()
    .setProperty("ADMIN_LOGGED_IN", "true");

  ui.alert("✅ Admin Login Successful");
}

// =====================================================
// ADMIN LOGOUT
// =====================================================

function adminLogout() {
  PropertiesService
    .getUserProperties()
    .deleteProperty("ADMIN_LOGGED_IN");

  SpreadsheetApp
    .getUi()
    .alert("✅ Logged Out Successfully");
}

// =====================================================
// SECURED ADMIN FUNCTIONS
// =====================================================

function adminInitialSetup() {
  requireAdmin();
  initialSetup();
}

function adminUnlockEditing() {
  requireAdmin();
  unlockEditing();
}

function adminLockEditing() {
  requireAdmin();
  lockEditing();
}

function adminDeleteAll() {
  requireAdmin();
  var ui = SpreadsheetApp.getUi();
  var confirm = ui.alert(
    "⚠️ DELETE EVERYTHING",
    "This will permanently erase ALL data.\n\nContinue?",
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    return;
  }
  deleteAll();
}

function adminShowManageBookForm() {
  requireAdmin();
  showManageBookForm();
}

function markPending() {
  updateSelectedOrderStatus(STATUS.PENDING);
}

function markProgress() {
  updateSelectedOrderStatus(STATUS.PROGRESS);
}

function markComplete() {
  updateSelectedOrderStatus(STATUS.COMPLETE);
}

function markCancelled() {
  updateSelectedOrderStatus(STATUS.CANCELLED);
}

// =====================================================
// STOCK TRANSACTION LOGIC (LOCKED)
// =====================================================

function updateSelectedOrderStatus(status) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    var sheet = getSheets().ORDERS;
    var range = sheet.getActiveRange();
    var row = range.getRow();
    var col = range.getColumn();

    if (col !== 14) {
      showToast("Select Status column first", "warning");
      return;
    }

    if (row < ORDER_START_ROW) {
      showToast("Select a valid order row", "warning");
      return;
    }

    var orderNo = String(sheet.getRange(row, 1).getValue()).trim();
    var buyer = String(sheet.getRange(row, 2).getValue()).trim();

    if (!orderNo || !buyer) {
      showToast("Invalid order row", "error");
      return;
    }

    var oldStatus = String(sheet.getRange(row, 14).getValue()).trim();
    if (oldStatus === status) {
      showToast("Same status already", "info");
      return;
    }

    // Set new status temporarily; rollback if stock update fails
    sheet.getRange(row, 14).setValue(status);
    SpreadsheetApp.flush();

    try {
      handleStatusChange(row);
    } catch (stockErr) {
      // Rollback status
      sheet.getRange(row, 14).setValue(oldStatus);
      SpreadsheetApp.flush();
      showToast("Stock update failed – status reverted", "error");
      return;
    }

    applyOrderStatusStyle(sheet, row);
    applyOrderRowStyle(sheet, row);
    refreshOrderDashboardStats();

    showToast("Status updated", "success");
  } catch (err) {
    Logger.log(err);
    showToast("Error updating status: " + err.message, "error");
  } finally {
    lock.releaseLock();
  }
}

function handleStatusChange(row) {
  var sheet = getSheets().ORDERS;

  var status = String(sheet.getRange(row, 14).getValue()).trim();
  var restoredCell = sheet.getRange(row, 16);
  var alreadyRestored = restoredCell.getValue() === "RESTORED";
  var items = parseOrderedBooks(sheet.getRange(row, 3).getValue());

  // CANCELLED - RESTORE STOCK
  if (status === STATUS.CANCELLED && !alreadyRestored) {
    var transactions = items.map(function (item) {
      return {
        bookNo: item.bookNo,
        qty: Number(item.qty) || 0,
        action: "ADD"
      };
    });
    processStockTransactions(transactions, "ORDER_CANCEL");
    restoredCell.setValue("RESTORED");
  }

  // REOPEN ORDER - DEDUCT STOCK AGAIN
  if (status !== STATUS.CANCELLED && alreadyRestored) {
    var transactions = items.map(function (item) {
      return {
        bookNo: item.bookNo,
        qty: Number(item.qty) || 0,
        action: "REMOVE"
      };
    });
    processStockTransactions(transactions, "ORDER_REOPEN");
    restoredCell.setValue("");
  }
}

/* ================= SAFE NEXT ROW ================= */
function getNextOrderRow(sheet) {
  var START_ROW = ORDER_START_ROW;
  var lastRow = Math.max(sheet.getLastRow(), START_ROW);
  var values = sheet
    .getRange(START_ROW, 1, lastRow - START_ROW + 1, 2)
    .getValues();

  for (var i = 0; i < values.length; i++) {
    var orderNo = String(values[i][0] || "").trim();
    var buyer = String(values[i][1] || "").trim();
    if (orderNo === "") {
      // Clean any leftover buyer text so the row is fully reusable
      if (buyer !== "") {
        sheet.getRange(START_ROW + i, 2).setValue("");
      }
      return START_ROW + i;
    }
  }
  return lastRow + 1;
}
/* ================= BOOK ================= */

function showBookForm() {
  var html = HtmlService.createHtmlOutputFromFile("BookForm")
    .setWidth(500)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, "📚 Add New Book");
}

function saveBook(data) {
  var sheet = getSheets().BOOKS;

  var bookNo =
    "B" + Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyyMMdd-HHmmss"
    );
  var qty = parseInt(data.qty, 10);
  var price = parseFloat(data.price);

  if (!data.bookName || data.bookName.trim() === "")
    throw new Error("Book name required");

  if (isNaN(qty) || qty < 0)
    throw new Error("Invalid Qty");

  if (isNaN(price) || price <= 0)
    throw new Error("Invalid Price");

  var row = getNextBookRow(sheet);   // <-- changed line

  sheet.getRange(row, 1, 1, 7).setValues([[
    bookNo,
    data.bookName.trim(),
    data.language,
    data.publication,
    data.author,
    qty,
    price
  ]]);

  sheet.getRange(row, 7)
    .setNumberFormat('₹#,##0.00');

  applyBookRowStyle(row);
  refreshBookDashboardStats();
  return "Book saved! " + bookNo;
}
function getNextBookRow(sheet) {
  var START_ROW = BOOK_START_ROW;
  var lastRow = Math.max(sheet.getLastRow(), START_ROW);
  var values = sheet
    .getRange(START_ROW, 1, lastRow - START_ROW + 1, 1)
    .getValues();

  for (var i = 0; i < values.length; i++) {
    var bookNo = String(values[i][0] || "").trim();
    if (bookNo === "") {
      return START_ROW + i;
    }
  }
  return lastRow + 1;
}
/* ================= BOOK LIST ================= */

function getBooks() {
  try {
    var sheet = getSheets().BOOKS;
    if (!sheet) return [];

    var lastRow = sheet.getLastRow();
    if (lastRow < BOOK_START_ROW) return [];

    var data = sheet.getRange(
      BOOK_START_ROW,
      1,
      lastRow - BOOK_START_ROW + 1,
      7
    ).getValues();

    return data
      .filter(r => r[1] && r[1].toString().trim() !== "")
      .map(r => {
        return {
          no: r[0],
          name: r[1],
          language: r[2],
          publication: r[3],
          author: r[4],
          qty: Number(r[5]) || 0,
          price: Number(r[6]) || 0
        };
      });

  } catch (err) {
    Logger.log(err);
    return [];
  }
}

/* ================= ORDER FORM ================= */

function showOrderForm() {
  var html = HtmlService.createHtmlOutputFromFile("Form")
    .setWidth(1100)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, "New Book Order");
}
/* ================= FIND BOOK ================= */

function findBook(bookNo) {
  var sheet = getSheets().BOOKS;
  var lastRow = sheet.getLastRow();
  var data =
    sheet.getRange(
      BOOK_START_ROW,
      1,
      lastRow - BOOK_START_ROW + 1,
      7
    )
      .getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] == bookNo) {
      return {
        row: i + BOOK_START_ROW,
        no: data[i][0],
        name: data[i][1],
        language: data[i][2],
        qty: Number(data[i][5]),
        price: Number(data[i][6])
      };
    }
  }
  return null;
}


/* ================= MANAGE ================= */

function showManageOrderForm(row) {
  const html = HtmlService
    .createTemplateFromFile("ManageOrder");

  html.row = row || "";

  SpreadsheetApp
    .getUi()
    .showModalDialog(
      html.evaluate()
        .setWidth(100)
        .setHeight(800),
      "Manage Order"
    );
}

function saveManagedOrder(row, originalItems, items) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var orderSheet = getSheets().ORDERS;

    if (!row || row < ORDER_START_ROW) {
      throw new Error("Invalid row");
    }

    if (!items || !items.length) {
      throw new Error("No items found");
    }

    var status = String(orderSheet.getRange(row, 14).getValue()).trim();
    if (status === STATUS.CANCELLED) {
      throw new Error("Cannot modify a cancelled/restored order");
    }

    var originalMap = {};
    var newMap = {};

    (originalItems || []).forEach(function (i) {
      if (!i || !i.bookNo) return;
      originalMap[i.bookNo] = Number(i.qty) || 0;
    });

    items.forEach(function (i) {
      if (!i || !i.bookNo) return;
      newMap[i.bookNo] = Number(i.qty) || 0;
    });

    var allBookNos = {};
    Object.keys(originalMap).forEach(function (bookNo) { allBookNos[bookNo] = true; });
    Object.keys(newMap).forEach(function (bookNo) { allBookNos[bookNo] = true; });

    var transactions = [];

    Object.keys(allBookNos).forEach(function (bookNo) {
      var oldQty = Number(originalMap[bookNo] || 0);
      var newQty = Number(newMap[bookNo] || 0);

      if (newQty > oldQty) {
        transactions.push({ bookNo: bookNo, qty: newQty - oldQty, action: "REMOVE" });
      } else if (oldQty > newQty) {
        transactions.push({ bookNo: bookNo, qty: oldQty - newQty, action: "ADD" });
      }
    });

    processStockTransactions(transactions, "MANAGE_ORDER");

    var orderedText = items
      .map(function (i) { return i.bookNo + ":" + i.qty; })
      .join(", ");

    orderSheet.getRange(row, 3).setValue(orderedText);
    rebuildSummaryFromOrderRow(row);

    SpreadsheetApp.flush();
    applyOrderRowStyle(orderSheet, row);
    
    return "Saved Successfully";

  } catch (err) {
    Logger.log(err);
    throw new Error("Save Managed Order Failed : " + err.message);
  } finally {
    try {
      lock.releaseLock();
    } catch (e) { }
  }
}
/* ================= SAVE BACK ================= */

function saveItemsBack(sheet, row, items) {
  var text = items.map(i => i.bookNo + ":" + i.qty).join(", ");
  sheet.getRange(row, 3).setValue(text);
  rebuildSummaryFromOrderRow(row);
}

function rebuildSummaryFromOrderRow(row) {
  var sheet = getSheets().ORDERS;
  var items = parseOrderedBooks(sheet.getRange(row, 3).getValue());
  var summary = "";
  var total = 0;
  var bookMap = getBookMap();

  items.forEach(function (i) {
    var stock = bookMap[i.bookNo];
    if (!stock) return;
    var t = i.qty * stock.price;
    total += t;
    summary += stock.name +
      " | " + i.bookNo +
      " | Qty: " + i.qty +
      " × ₹" + stock.price +
      " = ₹" + t + " ; ";
  });

  summary += "--\nTOTAL = ₹" + total;

  sheet.getRange(row, 4).setValue(summary);
  sheet.getRange(row, 5).setValue(total);
  autoFillFormulas(row);
  updateBillable(row);
}
/* ================= PARSER ================= */

function parseOrderedBooks(text) {
  var items = [];
  if (!text) return items;
  text.split(",").forEach(function (part) {
    part = String(part).trim();
    if (!part) return;
    var match = part.match(/^(.+?):(\d+)$/);
    if (!match) return;
    items.push({
      bookNo: match[1].trim(),
      qty: Number(match[2]) || 0
    });
  });
  return items;
}
/* ================= GET ORDER BOOKS ================= */
function getOrderBooks(row) {
  var sheet = getSheets().ORDERS;
  row = Number(row);
  if (!row || row < ORDER_START_ROW) {
    throw new Error("Please select valid order row");
  }
  var data = sheet.getRange(row, 1, 1, 17).getValues()[0];
  var orderNo = String(data[0] || "").trim();
  if (!orderNo) {
    throw new Error("Order not found");
  }
  var rawBooks = String(data[2] || "").trim();
  var items = parseOrderedBooks(rawBooks);
  var bookMap = buildBookMap();
  var books = [];

  items.forEach(function (i) {
    var stock = bookMap[i.bookNo];
    books.push({
      bookNo: i.bookNo,
      qty: Number(i.qty) || 0,
      name: stock ? stock.name : i.bookNo,
      language: stock ? stock.language : "",
      totalStock: stock ? Number(stock.qty) || 0 : 0
    });
  });

  return {
    row: row,
    items: books
  };
}
/* ================= PRINT ================= */
function printBillFromSelectedRow() {
  var sheet = getSheets().ORDERS;
  var row = sheet.getActiveRange().getRow();

  if (!row || row < ORDER_START_ROW) {
    showToast("Select a valid order row", "warning");
    return;
  }

  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

  var billable = String(data[12] || "").trim();
  if (billable !== "✅") {
    showToast("Not billable", "error");
    return;
  }

  var rawItems = String(data[2] || "").trim();
  if (!rawItems) {
    showToast("No books in order", "error");
    return;
  }
  
  var html = HtmlService.createTemplateFromFile("Bill");
   html.logo = getLogoBase64();
  html.billNo = data[0];
  html.name = data[1];
  html.address = formatAddress(data[9]);
  html.mobile = data[10];
  html.date = Utilities.formatDate(new Date(data[14]), Session.getScriptTimeZone(), "d MMMM yyyy");

  html.discountPrice = Number(data[6]) || 0;
  html.courierPrice = Number(data[7]) || 0;
  html.grandTotal = Number(data[4]) || 0;
  html.netPrice = html.grandTotal - html.discountPrice + html.courierPrice;

  var items = parseOrderedBooks(rawItems);
  var bookMap = getBookMap();

  html.items = items.map(function (b) {
    var stock = bookMap[b.bookNo];
    var bookName = b.bookNo;
    var price = 0;
    if (stock) {
      bookName = stock.name;
      price = Number(stock.price) || 0;
    }
    return {
      name: bookName,
      qty: Number(b.qty) || 0,
      price: price,
      total: (Number(b.qty) || 0) * price
    };
  });

  var output = html.evaluate().setWidth(1100).setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(output, "🧾 Print Bill");
}

function autoFillFormulas(row) {
  var sheet = getSheets().ORDERS;
  if (!row || row < ORDER_START_ROW) return;

  var total = Number(sheet.getRange(row, 5).getValue()) || 0;
  var discountPct = Number(sheet.getRange(row, 6).getValue()) || 0;
  var courier = Number(sheet.getRange(row, 8).getValue()) || 0;

  var discountAmount = total * discountPct / 100;
  var netTotal = total - discountAmount + courier;

  sheet.getRange(row, 7).setValue(discountAmount);
  sheet.getRange(row, 9).setValue(netTotal);
}
/* ================= FORMAT ================= */

function formatSheets() {
  [getSheets().BOOKS, getSheets().ORDERS]
    .forEach(formatSheet);
}

function formatSheet(sheet) {
  if (!sheet) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1) return;

  sheet.getRange(1, 1, lastRow, lastCol)
    .setFontFamily("Arial")
    .setFontSize(10)
    .setBorder(true, true, true, true, true, true);

  if (sheet.getName() === "All Book List") {
    sheet.setFrozenRows(BOOK_START_ROW - 1);
  }
  if (sheet.getName() === "Order Book List") {
    sheet.setFrozenRows(ORDER_START_ROW - 1);
  }
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();

    if (sheet.getName() !== "Order Book List")
      return;

    var row = e.range.getRow();
    var col = e.range.getColumn();

    if (row < ORDER_START_ROW) return;
    if (col !== 14) return;

    var lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);
      handleStatusChange(row);
      applyOrderStatusStyle(sheet, row);
      refreshOrderDashboardStats();
    } catch (lockErr) {
      Logger.log("Lock error in onEdit: " + lockErr);
    } finally {
      lock.releaseLock();
    }

  } catch (err) {
    Logger.log(err);
    SpreadsheetApp.getActiveSpreadsheet().toast(err.message, "Error", 5);
  }
}
function getLogoBase64() {

  var blob = UrlFetchApp.fetch(
    "https://raw.githubusercontent.com/dydlibrary-beep/my-assets/main/dydlogonew.jpeg"
  ).getBlob();

  return "data:image/jpeg;base64," +
         Utilities.base64Encode(blob.getBytes());
}
function getBillPdfUrlByActiveRow() {
  var sheet = getSheets().ORDERS;
  var row = sheet.getActiveRange().getRow();

  if (!row || row < ORDER_START_ROW) {
    throw new Error("No row selected");
  }

  var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var html = HtmlService.createTemplateFromFile("Bill");
  html.logo = getLogoBase64();
  html.billNo = data[0];
  html.name = data[1];
  html.address = formatAddress(data[9]).replace(/\n/g, "<br>");
  html.mobile = data[10];

  html.date = Utilities.formatDate(new Date(data[14]), Session.getScriptTimeZone(), "d MMMM yyyy");

  html.discountPrice = Number(data[6]) || 0;
  html.courierPrice = Number(data[7]) || 0;
  html.grandTotal = Number(data[4]) || 0;
  html.netPrice = html.grandTotal - html.discountPrice + html.courierPrice;

  var items = parseOrderedBooks(data[2] || "");
  var bookMap = getBookMap();
  Logger.log(html.logo.substring(0,100));
  Logger.log("Length = " + html.logo.length);
  html.items = items.map(function (b) {
    var stock = bookMap[b.bookNo];
    var bookName = b.bookNo;
    var price = 0;
    if (stock) {
      bookName = stock.name;
      price = stock.price;
    }
    return {
      name: bookName,
      qty: b.qty,
      price: price,
      total: price * b.qty
    };
  });

  var blob = html.evaluate().getAs("application/pdf").setName("Bill_" + data[0] + ".pdf");

  const BILL_FOLDER_ID = PropertiesService.getScriptProperties().getProperty("BILL_FOLDER_ID");
  if (!BILL_FOLDER_ID) {
    throw new Error("BILL_FOLDER_ID missing");
  }

  var existingFileId = String(sheet.getRange(row, 17).getValue() || "").trim();
  var file;
  
  if (existingFileId) {
    try {
      // ---------- FIX ----------
      file = DriveApp.getFileById(existingFileId);
      file.setBlob(blob);                       // replaces Drive.Files.update
      // -------------------------
    } catch (err) {
      Logger.log("Existing file not found, creating new PDF: " + err);
      file = DriveApp.getFolderById(BILL_FOLDER_ID).createFile(blob);
      existingFileId = file.getId();
    }
  } else {
    file = DriveApp.getFolderById(BILL_FOLDER_ID).createFile(blob);
    existingFileId = file.getId();
  }

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  sheet.getRange(row, 17).setValue(existingFileId);
  
  return file.getUrl();
}

function formatAddress(address) {
  return String(address || "")
    .replace(/<br\s*\/?>/gi, ",")
    .replace(/\r?\n/g, ",")
    .split(",")
    .map(part => part.trim())
    .filter(part => part !== "")
    .join("<br>");
}

function isValidIndianMobile(number) {
  if (number === null || number === undefined) {
    return false;
  }
  number = String(number).trim();
  number = number.replace(/\D/g, "");

  if (number.startsWith("91") && number.length === 12) {
    number = number.substring(2);
  }

  if (number.length !== 10) {
    return false;
  }

  return /^[6-9]\d{9}$/.test(number);
}

function deleteAll() {
  var ss = getSS();
  ss.getSheets().forEach(function (sheet) {
    var maxRows = sheet.getMaxRows();
    var maxCols = sheet.getMaxColumns();
    var fullRange = sheet.getRange(1, 1, maxRows, maxCols);

    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    try { if (sheet.getFilter()) sheet.getFilter().remove(); } catch (err) { }
    try {
      var merged = sheet.getDataRange().getMergedRanges();
      merged.forEach(function (r) { try { r.breakApart(); } catch (err) { } });
    } catch (err) { }
    try { sheet.getBandings().forEach(function (b) { try { b.remove(); } catch (err) { } }); } catch (err) { }
    try { sheet.setConditionalFormatRules([]); } catch (err) { }
    try {
      var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protections.forEach(function (p) { try { p.remove(); } catch (err) { } });
    } catch (err) { }
    try { fullRange.clearDataValidations(); } catch (err) { }
    try { fullRange.clearNote(); } catch (err) { }
    try { fullRange.clearComment(); } catch (err) { }
    try { fullRange.clearContent(); } catch (err) { }
    try { fullRange.clearFormat(); } catch (err) { }

    fullRange
      .setBackground("#ffffff")
      .setFontColor("#000000")
      .setFontWeight("normal")
      .setFontStyle("normal")
      .setFontLine("none")
      .setFontSize(10)
      .setFontFamily("Arial")
      .setHorizontalAlignment("general")
      .setVerticalAlignment("middle")
      .setWrap(false)
      .setTextRotation(0)
      .setBorder(false, false, false, false, false, false);

    try {
      for (var c = 1; c <= maxCols; c++) {
        sheet.setColumnWidth(c, 100);
      }
    } catch (err) { }
    try {
      for (var r = 1; r <= maxRows; r++) {
        sheet.setRowHeight(r, 21);
      }
    } catch (err) { }

    sheet.setHiddenGridlines(false);
  });

  SpreadsheetApp.flush();
  SpreadsheetApp.getActiveSpreadsheet().toast("Reset successful", "RESET", 5);
}

function saveOrder(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    var sheets = getSheets();
    var orderSheet = sheets.ORDERS;

    if (!data) throw new Error("No order data");
    if (!data.buyer || String(data.buyer).trim() === "") throw new Error("Buyer name required");
    if (!data.items || !data.items.length) throw new Error("No books selected");

    var bookMap = buildBookMap();

    var merged = {};
    data.items.forEach(function (item) {
      if (!item || !item.bookNo) return;
      var qty = Number(item.qty) || 0;
      if (qty <= 0) return;
      if (!merged[item.bookNo]) {
        merged[item.bookNo] = { bookNo: item.bookNo, qty: 0 };
      }
      merged[item.bookNo].qty += qty;
    });

    var items = Object.keys(merged).map(function (key) { return merged[key]; });
    if (!items.length) throw new Error("Invalid items");

    // Stock check
    items.forEach(function (item) {
      var stock = bookMap[item.bookNo];
      if (!stock) throw new Error(item.bookNo + " not found");
      if (Number(stock.qty) < Number(item.qty)) throw new Error(stock.name + " out of stock");
    });

    var summary = "";
    var booksText = [];
    var grandTotal = 0;

    items.forEach(function (item) {
      var stock = bookMap[item.bookNo];
      var qty = Number(item.qty);
      var price = Number(stock.price) || 0;
      var total = qty * price;
      grandTotal += total;
      booksText.push(item.bookNo + ":" + qty);
      summary += stock.name + " | Qty: " + qty + " × ₹" + price + " = ₹" + total + " ; ";
    });

    summary += "-- TOTAL = ₹" + grandTotal;

    // ================= DISCOUNT (AUTOMATIC TIER‑BASED) =================
    var discountPercent = 0;
    if (grandTotal >= 10000) {
      discountPercent = 30;
    } else if (grandTotal >= 5000) {
      discountPercent = 25;
    } else if (grandTotal >= 4000) {
      discountPercent = 20;
    } else if (grandTotal >= 2000) {
      discountPercent = 15;
    } else {   // grandTotal < 2000
      discountPercent = 10;
    }

    var discountAmount = grandTotal * (discountPercent / 100);
    var courier = Number(data.courier || 0);
    var netTotal = grandTotal - discountAmount + courier;

    var orderNo = "ORD-" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");

    var row = getNextOrderRow(orderSheet);

    orderSheet.getRange(row, 1, 1, 20).setValues([[
      orderNo,
      String(data.buyer || "").trim(),
      booksText.join(", "),
      summary,
      grandTotal,
      discountPercent,
      discountAmount,
      courier,
      netTotal,
      String(data.address || "").trim(),
      cleanIndianMobile(data.mobile),
      cleanIndianMobile(data.whatsapp),
      "", // Billable
      STATUS.PENDING,
      new Date(),
      "", // Restored
      "", // PDF ID
      "", // Track ID
      "", // Track Status
      ""  // Payment Status
    ]]);

    var transactions = items.map(function (item) {
      return { bookNo: item.bookNo, qty: item.qty, action: "REMOVE" };
    });
    processStockTransactions(transactions, "SAVE_ORDER");

    updateBillable(row);
    applyOrderRowStyle(orderSheet, row);
    refreshOrderDashboardStats();

    SpreadsheetApp.flush();

    return {
      success: true,
      orderNo: orderNo,
      row: row
    };

  } catch (err) {
    Logger.log(err);
    return {
      success: false,
      message: err.message
    };
  } finally {
    try { lock.releaseLock(); } catch (e) { }
  }
}

function getOrderSnapshot(row) {
  var sheet = getSheets().ORDERS;
  var oldText = sheet.getRange(row, 3).getValue();
  var oldItems = parseOrderedBooks(oldText);
  return oldItems;
}

function downloadBillPdf() {
  var url = getBillPdfUrlByActiveRow();
  var html = HtmlService.createHtmlOutput(
    '<script>' +
    'window.open("' + url + '","_blank");' +
    'google.script.host.close();' +
    '</script>'
  );
  SpreadsheetApp.getUi().showModalDialog(html, "Downloading...");
}

function showOrderDetails(orderText, row) {
  const html = HtmlService.createTemplateFromFile("OrderDetails");
  html.orderText = orderText;
  html.row = row;
  SpreadsheetApp.getUi().showModalDialog(html.evaluate().setWidth(900).setHeight(700), "Order Details");
}

function getSelectedOrderUpdateData() {
  var sheet = getSheets().ORDERS;
  var activeCell = sheet.getActiveCell();

  if (!activeCell) throw new Error("Please select any order row");

  var row = activeCell.getRow();
  if (!row || row < ORDER_START_ROW) throw new Error("Please select valid order row");

  var rowData = sheet.getRange(row, 1, 1, 2).getValues()[0];
  var orderNo = String(rowData[0] || "").trim();
  if (!orderNo) throw new Error("Order not found");

  var data = sheet.getRange(row, 1, 1, 20).getValues()[0];
  var rawBooks = String(data[2] || "").trim();
  var items = parseOrderedBooks(rawBooks);
  var bookMap = buildBookMap();
  var books = [];
  var total = 0;

  items.forEach(function (i) {
    try {
      var stock = bookMap[i.bookNo];
      var qty = Number(i.qty) || 0;
      var price = stock ? Number(stock.price) || 0 : 0;
      var amount = qty * price;
      total += amount;

      books.push({
        bookNo: i.bookNo,
        name: stock ? stock.name : i.bookNo,
        qty: qty,
        price: price,
        amount: amount
      });
    } catch (err) { Logger.log(err); }
  });

  return {
    row: row,
    orderNo: String(data[0] || ""),
    buyer: String(data[1] || ""),
    books: books,
    total: Number(data[4]) || total || 0,
    discountPercent: Number(data[5]) || 0,
    discountAmount: Number(data[6]) || 0,
    courier: Number(data[7]) || 0,
    netTotal: Number(data[8]) || 0,
    address: String(data[9] || ""),
    mobile: String(data[10] || ""),
    whatsapp: String(data[11] || ""),
    status: String(data[13] || "")
  };
}

function showUpdateOrderForm() {
  var html = HtmlService.createHtmlOutputFromFile("UpdateOrder")
    .setWidth(700)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, "📝 Update Order Details");
}

function processStockTransactions(transactions, operation) {
  if (!transactions || !transactions.length) return;

  var sheet = getSheets().BOOKS;
  var bookMap = buildBookMap();

  transactions.forEach(function (t) {
    var stock = bookMap[t.bookNo];
    if (!stock) throw new Error(t.bookNo + " not found");

    var qty = Number(t.qty) || 0;
    if (qty <= 0) throw new Error("Invalid qty for " + t.bookNo);

    var current = Number(stock.qty) || 0;
    var finalQty = (t.action === "ADD") ? current + qty : current - qty;
    if (finalQty < 0) throw new Error(stock.name + " stock below zero");
  });

  transactions.forEach(function (t) {
    var stock = bookMap[t.bookNo];
    var qty = Number(t.qty) || 0;
    var current = Number(stock.qty) || 0;
    var finalQty = (t.action === "ADD") ? current + qty : current - qty;

    sheet.getRange(stock.row, 6).setValue(finalQty);
    stock.qty = finalQty;

    Logger.log(operation + " | " + t.bookNo + " | " + t.action + " | QTY = " + qty + " | FINAL = " + finalQty);
  });

  SpreadsheetApp.flush();
  
}

function updateStock(bookNo, qtyChange, operation) {
  var sheet = getSheets().BOOKS;
  var lastRow = sheet.getLastRow();

  if (lastRow < BOOK_START_ROW) throw new Error("No books found");

  var data = sheet.getRange(BOOK_START_ROW, 1, lastRow - BOOK_START_ROW + 1, 7).getValues();

  var foundRow = null;
  var stockName = "";
  var currentQty = 0;

  for (var i = 0; i < data.length; i++) {
    var currentBookNo = String(data[i][0] || "").trim();
    if (currentBookNo === String(bookNo).trim()) {
      foundRow = BOOK_START_ROW + i;
      stockName = String(data[i][1] || "").trim();
      currentQty = Number(data[i][5]) || 0;
      break;
    }
  }

  if (!foundRow) throw new Error(bookNo + " not found");

  var finalQty = currentQty + Number(qtyChange);
  if (finalQty < 0) throw new Error(stockName + " stock below zero");
  finalQty = Math.max(0, finalQty);

  sheet.getRange(foundRow, 6).setValue(finalQty);
  SpreadsheetApp.flush();
  refreshBookDashboardStats();

  Logger.log(operation + " | " + bookNo + " | OLD = " + currentQty + " | CHANGE = " + qtyChange + " | FINAL = " + finalQty);
  return finalQty;
}

function updateOrderDetails(data) {
  var sheet = getSheets().ORDERS;
  var row = Number(data.row);

  if (!row || row < ORDER_START_ROW) throw new Error("Invalid row");

  var total = Number(sheet.getRange(row, 5).getValue()) || 0;
  var discountPercent = Number(data.discountPercent) || 0;
  var discountAmount = Number(data.discountAmount) || 0;
  var courier = Number(data.courier) || 0;
  var netTotal = total - discountAmount + courier;

  sheet.getRange(row, 6, 1, 4).setValues([[discountPercent, discountAmount, courier, netTotal]]);
  sheet.getRange(row, 10).setValue(data.address || "");

  var mobile = cleanIndianMobile(data.mobile);
  if (mobile !== "" && !/^[6-9]\d{9}$/.test(mobile)) throw new Error("Enter valid 10-digit mobile number");

  var whatsapp = cleanIndianMobile(data.whatsapp);
  if (whatsapp !== "" && !/^[6-9]\d{9}$/.test(whatsapp)) throw new Error("Enter valid 10-digit WhatsApp number");

  sheet.getRange(row, 11).setValue(mobile);
  sheet.getRange(row, 12).setValue(whatsapp);

  updateBillable(row);
  applyOrderRowStyle(sheet, row);
  SpreadsheetApp.flush();

  return "Order updated successfully";
}

function openSelectedOrder() {
  try {
    var sheet = getSheets().ORDERS;
    var row = sheet.getActiveRange().getRow();

    if (!row || row < ORDER_START_ROW) {
      showToast("Select a valid order row", "warning");
      return;
    }

    var orderNo = String(sheet.getRange(row, 1).getValue()).trim();
    if (!orderNo) {
      showToast("No order found", "error");
      return;
    }

    var text = String(sheet.getRange(row, 4).getValue()).trim();
    if (!text) text = "No summary available";
    showOrderDetails(text, row);

  } catch (err) {
    Logger.log(err);
    SpreadsheetApp.getUi().alert("Open Order Failed:\n\n" + err.message);
  }
}

function cleanIndianMobile(number) {
  if (number === null || number === undefined) return "";
  number = String(number)
    .trim()
    .replace(/\.0$/, "")
    .replace(/\s+/g, "")
    .replace(/\D/g, "");

  if (number.startsWith("91") && number.length > 10) {
    number = number.slice(-10);
  }
  return number;
}

function sendBillToWhatsApp() {
  try {
    var sheet = getSheets().ORDERS;
    var row = sheet.getActiveRange().getRow();

    if (row < ORDER_START_ROW) {
      showToast("Select an order row", "warning");
      return;
    }

    var data = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    var billable = String(data[12] || "").trim();
    if (billable !== "✅") {
      showToast("Not billable", "error");
      return;
    }

    var mobile = cleanIndianMobile(data[11]);
    if (!mobile) {
      showToast("No WhatsApp number", "error");
      return;
    }
    if (mobile.length !== 10 || !/^[6-9]\d{9}$/.test(mobile)) {
      showToast("Invalid WhatsApp number", "error");
      return;
    }

    mobile = "91" + mobile;
    var customerName = String(data[1] || "").trim();
    var pdfUrl = getBillPdfUrlByActiveRow();
    var amount = Number(data[8]) || 0;

    if (amount <= 0) {
      showToast("Invalid amount", "error");
      return;
    }

    var upiId = "darsh99782@barodampay";
    var upiLink = "upi://pay?pa=" + encodeURIComponent(upiId) +
      "&pn=" + encodeURIComponent("Darshan Yog Dham") +
      "&am=" + encodeURIComponent(amount) +
      "&cu=INR" +
      "&tn=" + encodeURIComponent(data[0]);

    var message = "Hello " + customerName + ",\n\nYour bill is ready.\n\n🧾 Download Bill:\n " + pdfUrl + "\n\n💳 Pay Now:\n " + upiLink + "\n\nAmount: ₹" + amount + "\n\nSupported Apps:\n• Google Pay\n• PhonePe\n• Paytm\n• BHIM UPI\n\nAfter payment please send screenshot.\n\nThank you.";

    var whatsappUrl = "https://wa.me/" + mobile + "?text=" + encodeURIComponent(message);

    var html = HtmlService.createHtmlOutput(
      '<script>' +
      'window.open("' + whatsappUrl + '","_blank");' +
      'google.script.host.close();' +
      '</script>'
    );

    SpreadsheetApp.getUi().showModalDialog(html, "Opening WhatsApp...");

  } catch (err) {
    Logger.log(err);
    showToast(err.message, "error");
  }
}

function deleteOldBillPdfs() {
  const BILL_FOLDER_ID = PropertiesService.getScriptProperties().getProperty("BILL_FOLDER_ID");
  var folder = DriveApp.getFolderById(BILL_FOLDER_ID);
  var files = folder.getFiles();
  var now = new Date().getTime();

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();

    if (name.indexOf("Bill_") === 0) {
      var created = file.getDateCreated().getTime();
      var ageDays = (now - created) / (1000 * 60 * 60 * 24);

      if (ageDays > 7) {
        file.setTrashed(true);
        Logger.log("Deleted: " + file.getName());
      }
    }
  }
}

/* ================= STOCK WINDOW ================= */

function showManageBookForm() {
  var html = HtmlService.createHtmlOutputFromFile("ManageBook")
    .setWidth(750)
    .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, "📚 Manage Book");
}

function saveBookChanges(data) {
  var sheet = getSheets().BOOKS;
  var stock = getBookMap()[data.bookNo];

  if (!stock) throw new Error("Book not found");

  var name = String(data.name || "").trim();
  if (!name) throw new Error("Book name required");

  var price = Number(data.price);
  if (isNaN(price) || price <= 0) throw new Error("Invalid price");

  var stockQty = Number(data.stockQty) || 0;
  if (stockQty < 0) throw new Error("Invalid stock quantity");

  sheet.getRange(stock.row, 2).setValue(name);
  sheet.getRange(stock.row, 3).setValue(data.language);
  sheet.getRange(stock.row, 4).setValue(data.publication);
  sheet.getRange(stock.row, 5).setValue(data.author);
  sheet.getRange(stock.row, 7).setValue(price);

  var current = Number(stock.qty) || 0;
  var finalQty = current;

  if (stockQty > 0) {
    if (data.stockAction === "add") {
      finalQty = current + stockQty;
      sheet.getRange(stock.row, 6).setValue(finalQty);
    } else {
      if (current < stockQty) throw new Error("Stock cannot go below zero");
      finalQty = current - stockQty;
      sheet.getRange(stock.row, 6).setValue(finalQty);
    }
  }

  SpreadsheetApp.flush();
  refreshBookDashboardStats();

  return { name: name, oldQty: current, newQty: finalQty };
}

function getSelectedOrderBooks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var row = sheet.getActiveCell().getRow();

  if (row < ORDER_START_ROW) throw new Error("Please select order row");
  return getOrderBooks(row);
}

function updateBookStock(data) {
  var sheet = getSheets().BOOKS;
  var stock = getBookMap()[data.bookNo];

  if (!stock) throw new Error("Book not found");

  var currentQty = Number(stock.qty) || 0;
  var changeQty = Number(data.qty) || 0;
  var finalQty;

  if (data.type === "add") {
    finalQty = currentQty + changeQty;
  } else {
    finalQty = currentQty - changeQty;
    if (finalQty < 0) throw new Error("Stock cannot go below zero");
  }

  sheet.getRange(stock.row, 6).setValue(finalQty);
  SpreadsheetApp.flush();
  refreshBookDashboardStats();

  return { oldQty: currentQty, newQty: finalQty, book: stock.name };
}

/* ================= DOWNLOAD STOCK REPORT ================= */

function downloadStockReport() {
  var sheet = getSheets().BOOKS;
  if (!sheet) return;

  var lastRow = sheet.getLastRow();

  if (lastRow < BOOK_START_ROW) {
    showToast('No stock data to download', 'warning');
    return;
  }

  var headerRange = sheet.getRange(6, 1, 1, BOOK_HEADERS.length);
  var dataRange = sheet.getRange(
    BOOK_START_ROW,
    1,
    lastRow - BOOK_START_ROW + 1,
    BOOK_HEADERS.length
  );

  var headers = headerRange.getValues()[0];
  var data = dataRange.getValues();

  var csvLines = [];

  csvLines.push(
    headers.map(function (cell) {
      return '"' + String(cell || '').replace(/"/g, '""') + '"';
    }).join(',')
  );

  data.forEach(function (row) {
    csvLines.push(
      row.map(function (cell) {
        return '"' + String(cell || '').replace(/"/g, '""') + '"';
      }).join(',')
    );
  });

  var csvContent = csvLines.join('\n');

  var bom = '\uFEFF';
  var blob = Utilities.newBlob(
    bom + csvContent,
    'text/csv;charset=utf-8',
    'Total_Stock_Report.csv'
  );

  var file = DriveApp.createFile(blob);
  var url = file.getDownloadUrl();

  var html = HtmlService.createHtmlOutput(
    '<script>' +
    'window.open("' + url + '");' +
    'google.script.host.close();' +
    '</script>'
  );

  SpreadsheetApp.getUi().showModalDialog(html, 'Downloading...');
}

function buildBookMap() {
  var sheet = getSheets().BOOKS;
  var START_ROW = BOOK_START_ROW;
  var lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) return {};

  var data = sheet.getRange(START_ROW, 1, lastRow - START_ROW + 1, 7).getValues();

  var map = {};

  data.forEach(function (r, i) {
    if (!r[0]) return;
    map[r[0]] = {
      row: START_ROW + i,
      no: r[0],
      name: r[1],
      language: r[2],
      qty: Number(r[5]) || 0,
      price: Number(r[6]) || 0
    };
  });

  return map;
}

function setupBillFolder(folderId) {
  if (!folderId) throw new Error("Folder ID required");
  PropertiesService.getScriptProperties().setProperty("BILL_FOLDER_ID", folderId);
}

function getBookMap() {
  return buildBookMap();
}

// ================= TOAST FUNCTION =================

function getSelectedOrderDetails(row) {
  var sheet = getSheets().ORDERS;
  row = Number(row);

  if (!row || row < ORDER_START_ROW) throw new Error("Invalid row");

  var data = sheet.getRange(row, 1, 1, 20).getValues()[0];

  var items = parseOrderedBooks(data[2]);
  var bookMap = getBookMap();

  var books = items.map(function (i) {
    var stock = bookMap[i.bookNo];
    var price = stock ? stock.price : 0;
    return {
      bookNo: i.bookNo,
      name: stock ? stock.name : i.bookNo,
      qty: i.qty,
      price: price,
      total: i.qty * price
    };
  });

  return {
    row: row,
    items: books,
    total: Number(data[4]) || 0,
    discountPercent: Number(data[5]) || 0,
    discountAmount: Number(data[6]) || 0,
    courier: Number(data[7]) || 0,
    netTotal: Number(data[8]) || 0,
    address: data[9] || "",
    mobile: data[10] || "",
    whatsapp: data[11] || ""
  };
}

function createBillFolderAndSaveId() {
  var props = PropertiesService.getScriptProperties();
  var folderId = props.getProperty("BILL_FOLDER_ID");

  if (folderId) {
    try {
      var existingFolder = DriveApp.getFolderById(folderId);
      SpreadsheetApp.getActiveSpreadsheet().toast("Bill folder already exists", "Info", 5);
      Logger.log("Existing Folder: " + existingFolder.getName());
      return folderId;
    } catch (err) {
      Logger.log("Old folder missing. Creating new folder...");
      props.deleteProperty("BILL_FOLDER_ID");
    }
  }

  var folder = DriveApp.createFolder("Library Bills");
  props.setProperty("BILL_FOLDER_ID", folder.getId());
  SpreadsheetApp.getActiveSpreadsheet().toast("Bill folder created", "Success", 5);
  Logger.log("Folder URL: " + folder.getUrl());
  return folder.getId();
}

/* ================= BILLABLE RULE ================= */

function updateBillable(row) {
  if (row < ORDER_START_ROW) return;

  var sheet = getSheets().ORDERS;
  var d = sheet.getRange(row, 1, 1, 14).getValues()[0];

  function isValid(v) {
    return v !== null && v !== undefined && String(v).trim().length > 0;
  }

  var isComplete =
    isValid(d[0]) &&
    isValid(d[1]) &&
    isValid(d[2]) &&
    isValid(d[3]) &&
    !isNaN(Number(d[4])) &&
    Number(d[4]) > 0 &&
    isValid(d[9]) &&
    isValid(d[10]);

  var cell = sheet.getRange(row, 13);

  if (isComplete) {
    cell
      .setValue("✅")
      .setBackground("#dcfce7")
      .setFontColor("#166534")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  } else {
    cell
      .setValue("❌")
      .setBackground("#fee2e2")
      .setFontColor("#991b1b")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
  }
}

// ================= Sheet Design methods =================

function setUpBillableFields() {
  var sheet = getSheets().ORDERS;
  var startRow = ORDER_START_ROW;
  var lastRow = sheet.getLastRow();

  if (lastRow < startRow) return;

  for (var row = startRow; row <= lastRow; row++) {
    var cell = sheet.getRange(row, 13);
    var value = String(cell.getValue()).trim().toUpperCase();

    if (value === "YES" || value === "✅") {
      cell
        .setValue("✅")
        .setBackground("#dcfce7")
        .setFontColor("#166534")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
    } else {
      cell
        .setValue("❌")
        .setBackground("#fee2e2")
        .setFontColor("#991b1b")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
    }
  }
  SpreadsheetApp.flush();
}

function formatOrderSheetLayout() {
  var sheet = getSheets().ORDERS;
  if (!sheet) return;

  var numHeaders = ORDER_HEADERS.length;
  var lastRow = Math.max(sheet.getLastRow(), 4);

  while (sheet.getMaxColumns() < numHeaders) {
    sheet.insertColumnAfter(sheet.getMaxColumns());
  }

  try { sheet.getRange("1:4").breakApart(); } catch (e) { }
  sheet.getRange("1:4").clearContent().clearFormat();

  try {
    var images = sheet.getImages();
    images.forEach(function (img) {
      var row = img.getAnchorRow();
      if (row >= 1 && row <= 3) {
        img.remove();
      }
    });
  } catch (e) { }

  try { if (sheet.getFilter()) sheet.getFilter().remove(); } catch (e) { }
  sheet.setHiddenGridlines(true);

  var widths = [
    150, 250, 80, 300, 110, 110, 110, 110, 120, 220, 140, 150, 110, 150, 160, 120, 120, 130, 150, 150
  ];

  for (var i = 0; i < widths.length; i++) {
    sheet.setColumnWidth(i + 1, widths[i]);
  }

  sheet.getRange(1, 1, 1, numHeaders)
    .merge()
    .setValue("📋 ORDER MANAGEMENT DASHBOARD")
    .setBackground("#1E293B")
    .setFontColor("#FFFFFF")
    .setFontSize(16)
    .setFontWeight("bold")
    .setFontFamily("Inter")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(1, 48);

  var cards = [
    { range: "A2:B2", title: "📦 TOTAL ORDERS", value: "0", bg: "#EFF6FF", color: "#1E3A8A" },
    { range: "C2:D2", title: "🚚 IN PROGRESS", value: "0", bg: "#FEF3C7", color: "#92400E" },
    { range: "E2:F2", title: "✅ COMPLETED", value: "0", bg: "#D1FAE5", color: "#065F46" },
    { range: "G2:H2", title: "❌ CANCELLED", value: "0", bg: "#FEE2E2", color: "#991B1B" },
    { range: "I2:J2", title: "💰 TOTAL SALES", value: "₹0", bg: "#F3E8FF", color: "#6B21A8" }
  ];

  cards.forEach(function (card) {
    sheet.getRange(card.range)
      .merge()
      .setValue(card.title + "\n" + card.value)
      .setBackground(card.bg)
      .setFontColor(card.color)
      .setFontWeight("bold")
      .setFontSize(12)
      .setFontFamily("Inter")
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true);
  });

  if (numHeaders > 10) {
    sheet.getRange(2, 11, 1, numHeaders - 10)
      .merge()
      .setValue("")
      .setBackground("#F8FAFC");
  }
  sheet.setRowHeight(2, 56);

  sheet.getRange(3, 1, 1, numHeaders).setBackground("#3B82F6");
  sheet.setRowHeight(3, 35);

  sheet.getRange(4, 1, 1, numHeaders)
    .setValues([ORDER_HEADERS])
    .setBackground("#2563EB")
    .setFontColor("#FFFFFF")
    .setFontWeight("bold")
    .setFontSize(11)
    .setFontFamily("Inter")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 32);

  var dataStartRow = 5;
  var totalRows = lastRow - (dataStartRow - 1);

  if (totalRows > 0) {
    var dataRange = sheet.getRange(dataStartRow, 1, totalRows, numHeaders);

    var bgArray = [];
    for (var r = 0; r < totalRows; r++) {
      var bg = (r % 2 === 0) ? "#FFFFFF" : "#F8FAFC";
      var rowColors = [];
      for (var c = 0; c < numHeaders; c++) {
        rowColors.push(bg);
      }
      bgArray.push(rowColors);
    }

    dataRange.setBackgrounds(bgArray)
      .setFontFamily("Inter")
      .setFontSize(10)
      .setVerticalAlignment("middle");

    sheet.getRange(dataStartRow, 1, totalRows, 1).setHorizontalAlignment("center");
    sheet.getRange(dataStartRow, 3, totalRows, 1).setHorizontalAlignment("center");
    sheet.getRange(dataStartRow, 8, totalRows, 1).setHorizontalAlignment("center");
    sheet.getRange(dataStartRow, 9, totalRows, 1).setHorizontalAlignment("right");
    sheet.getRange(dataStartRow, 2, totalRows, 1).setWrap(true);
    sheet.getRange(dataStartRow, 4, totalRows, 1).setWrap(true);
    sheet.getRange(dataStartRow, 10, totalRows, 1).setWrap(true);

    dataRange.setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);

    try { sheet.autoResizeRows(dataStartRow, totalRows); } catch (e) { }
  }

  sheet.setFrozenRows(4);

  try {
    if (sheet.getFilter()) sheet.getFilter().remove();
    sheet.getRange(4, 1, Math.max(lastRow - 3, 1), numHeaders).createFilter();
  } catch (e) { }

  sheet.setTabColor("#0EA5E9");
  SpreadsheetApp.flush();

  try { refreshOrderDashboardStats(); } catch (e) { Logger.log(e); }
}

function centerHeaderIcons() {
  var sheet = getSheets().ORDERS;
  if (!sheet) return;

  var numHeaders = ORDER_HEADERS.length;
  var iconSize = 24;

  sheet.setRowHeight(3, 35);

  var images = sheet.getImages();
  images.forEach(function (img) {
    var col = img.getAnchorColumn();
    var row = img.getAnchorRow();
    if (row === 3 && col <= numHeaders) {
      img.setWidth(iconSize);
      img.setHeight(iconSize);
    }
  });

  sheet.getRange(3, 1, 1, numHeaders)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert('✅ All icons centered successfully!');
}

function formatAllBookSheetLayout() {
  var sheet = getSheets().BOOKS;
  if (!sheet) return;

  var lastCol = BOOK_HEADERS.length;
  var lastRow = Math.max(sheet.getLastRow(), 7);

  try { if (sheet.getFilter()) sheet.getFilter().remove(); } catch (e) { }
  sheet.getBandings().forEach(function (b) { b.remove(); });

  sheet.getRange("1:5").clearContent().clearFormat();

  var totalTitles = 0, totalStock = 0, totalValue = 0, lowStock = 0;
  var languageMap = {};

  if (lastRow >= 7) {
    var data = sheet.getRange(7, 1, lastRow - 6, lastCol).getValues();
    data.forEach(function (r) {
      var bookNo = r[0];
      var lang = String(r[2] || "").trim();
      var qty = Number(r[5]) || 0;
      var price = Number(r[6]) || 0;
      if (bookNo) totalTitles++;
      totalStock += qty;
      totalValue += qty * price;
      if (qty <= 5) lowStock++;
      if (lang) {
        if (!languageMap[lang]) languageMap[lang] = { qty: 0 };
        languageMap[lang].qty += qty;
      }
    });
  }

  sheet.getRange("A1:B1").merge()
    .setValue("📘 LIBRARY")
    .setBackground("#134e4a")
    .setFontColor("#fbbf24")
    .setFontSize(14)
    .setFontWeight("bold")
    .setFontFamily("Poppins")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#fbbf24", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  sheet.getRange("C1:G1").merge()
    .setValue("BOOK INVENTORY DASHBOARD")
    .setBackground("#134e4a")
    .setFontColor("#ffffff")
    .setFontSize(16)
    .setFontWeight("bold")
    .setFontFamily("Poppins")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#fbbf24", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(1, 50);

  var cards = [
    { range: "A2:B3", title: "📘 TOTAL TITLES", value: totalTitles, bg: "#ecfdf5", color: "#065f46" },
    { range: "C2:D3", title: "📦 TOTAL STOCK", value: totalStock, bg: "#d1fae5", color: "#065f46" },
    { range: "E2:F3", title: "💰 STOCK VALUE", value: "₹" + totalValue.toLocaleString("en-IN"), bg: "#a7f3d0", color: "#064e3b" },
    { range: "G2:G3", title: "⚠ LOW STOCK", value: lowStock, bg: "#ffe4e6", color: "#9f1239" }
  ];

  cards.forEach(function (card) {
    sheet.getRange(card.range).merge()
      .setValue(card.title + "\n" + card.value)
      .setBackground(card.bg)
      .setFontColor(card.color)
      .setFontWeight("bold")
      .setFontFamily("Inter")
      .setFontSize(13)
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle")
      .setWrap(true)
      .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  });

  var languageText = [];
  Object.keys(languageMap)
    .sort(function (a, b) { return languageMap[b].qty - languageMap[a].qty; })
    .slice(0, 10)
    .forEach(function (lang) { languageText.push(lang + " : " + languageMap[lang].qty); });

  sheet.getRange("A4").setValue("🌎 Top Languages")
    .setBackground("#ccfbf1").setFontColor("#115e59").setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  sheet.getRange("B4:G4").merge()
    .setValue(languageText.join("   |   "))
    .setBackground("#ffffff").setFontColor("#111827").setFontWeight("bold")
    .setWrap(false).setHorizontalAlignment("left").setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  var header = sheet.getRange(6, 1, 1, lastCol);
  header.setValues([BOOK_HEADERS])
    .setBackground("#0f766e")
    .setFontColor("#ffffff")
    .setFontWeight("bold")
    .setFontFamily("Inter")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, "#0f766e", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.setRowHeight(6, 40);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 510);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 340);
  sheet.setColumnWidth(5, 340);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 180);

  if (lastRow >= 7) {
    sheet.setRowHeights(7, lastRow - 6, 42);
    sheet.getRange(7, 1, lastRow - 6, 1).setHorizontalAlignment("center");
    sheet.getRange(7, 3, lastRow - 6, 1).setHorizontalAlignment("center");
    sheet.getRange(7, 6, lastRow - 6, 2).setHorizontalAlignment("center");
    sheet.getRange(7, 2, lastRow - 6, 4).setWrap(true);
    sheet.getRange(7, 6, lastRow - 6, 1).setNumberFormat("0");
    sheet.getRange(7, 7, lastRow - 6, 1).setNumberFormat("₹#,##0.00");

    sheet.getRange(7, 1, lastRow - 6, lastCol)
      .setFontFamily("Inter")
      .setFontSize(10)
      .setBorder(true, true, true, true, false, false, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);
  }

  sheet.setHiddenGridlines(true);
  sheet.setTabColor("#059669");
  sheet.setFrozenRows(6);

  try {
    if (!sheet.getFilter()) {
      sheet.getRange(6, 1, Math.max(lastRow - 5, 1), lastCol).createFilter();
    }
  } catch (e) { }
}

function applyBookRowStyle(row) {
  var sheet = getSheets().BOOKS;
  if (row < 7) return;

  var lastCol = BOOK_HEADERS.length;
  sheet.setRowHeight(row, 42);

  var baseBg = (row % 2 === 0) ? "#ffffff" : "#f0fdf4";
  var rowRange = sheet.getRange(row, 1, 1, lastCol);
  rowRange
    .setBackground(baseBg)
    .setFontFamily("Inter")
    .setFontSize(10)
    .setFontColor("#1e293b")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, false, false, "#e2e8f0", SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange(row, 1)
    .setBackground("#ccfbf1")
    .setFontWeight("bold")
    .setFontColor("#0f766e")
    .setHorizontalAlignment("center")
    .setFontSize(10);

  sheet.getRange(row, 2)
    .setFontWeight("bold")
    .setFontColor("#0f172a")
    .setFontSize(11)
    .setHorizontalAlignment("left");

  sheet.getRange(row, 3)
    .setBackground("#e2e8f0")
    .setFontColor("#1e40af")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.getRange(row, 4)
    .setFontColor("#334155")
    .setHorizontalAlignment("left");

  sheet.getRange(row, 5)
    .setFontColor("#475569")
    .setFontStyle("italic")
    .setHorizontalAlignment("left");

  var stock = Number(sheet.getRange(row, 6).getValue()) || 0;
  var stockCell = sheet.getRange(row, 6);
  stockCell
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setFontSize(11);

  if (stock <= 0) {
    stockCell.setBackground("#7f1d1d").setFontColor("#ffffff");
  } else if (stock <= 5) {
    stockCell.setBackground("#fee2e2").setFontColor("#b91c1c");
  } else if (stock <= 15) {
    stockCell.setBackground("#fef3c7").setFontColor("#92400e");
  } else {
    stockCell.setBackground("#dcfce7").setFontColor("#166534");
  }

  sheet.getRange(row, 7)
    .setBackground("#eff6ff")
    .setFontColor("#1d4ed8")
    .setFontWeight("bold")
    .setFontSize(11)
    .setHorizontalAlignment("center")
    .setNumberFormat("₹#,##0.00");

  rowRange.setBorder(false, false, true, false, false, false, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
}

function refreshOrderDashboardStats() {
  var sheet = getSheets().ORDERS;
  if (!sheet) {
    Logger.log("ERROR: Order sheet not found");
    return;
  }

  var DATA_START_ROW = ORDER_START_ROW;
  var lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    sheet.getRange("A2:B2").setValue("📦 TOTAL\n0");
    sheet.getRange("C2:D2").setValue("🚚 PROGRESS\n0");
    sheet.getRange("E2:F2").setValue("✅ COMPLETE\n0");
    sheet.getRange("G2:H2").setValue("❌ CANCELLED\n0");
    sheet.getRange("I2:J2").setValue("💰 SALES\n₹0");
    return;
  }

  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, ORDER_HEADERS.length).getValues();

  var totalOrders = 0, pending = 0, progress = 0, complete = 0, cancelled = 0, sales = 0;

  data.forEach(function (r) {
    var orderNo = String(r[0] || "").trim();
    var buyer = String(r[1] || "").trim();
    if (!orderNo && !buyer) return;

    totalOrders++;
    var status = String(r[13] || "").trim();
    var netTotal = Number(r[8]) || 0;

    if (status === STATUS.PENDING) pending++;
    else if (status === STATUS.PROGRESS) progress++;
    else if (status === STATUS.COMPLETE) { complete++; sales += netTotal; }
    else if (status === STATUS.CANCELLED) cancelled++;
  });

  sheet.getRange("A2:B2").setValue("📦 TOTAL\n" + totalOrders);
  sheet.getRange("C2:D2").setValue("🚚 PROGRESS\n" + progress);
  sheet.getRange("E2:F2").setValue("✅ COMPLETE\n" + complete);
  sheet.getRange("G2:H2").setValue("❌ CANCELLED\n" + cancelled);
  sheet.getRange("I2:J2").setValue("💰 SALES\n₹" + sales.toLocaleString("en-IN"));

  SpreadsheetApp.flush();
}

function refreshBookDashboardStats() {
  var sheet = getSheets().BOOKS;
  if (!sheet) return;

  var HEADER_ROW = 6;
  var DATA_START_ROW = 7;
  var lastCol = BOOK_HEADERS.length;
  var lastRow = sheet.getLastRow();

  if (lastRow < DATA_START_ROW) {
    sheet.getRange("A2:B3").setValue("📘 TOTAL TITLES\n0");
    sheet.getRange("C2:D3").setValue("📦 TOTAL STOCK\n0");
    sheet.getRange("E2:F3").setValue("💰 STOCK VALUE\n₹0");
    sheet.getRange("G2:G3").setValue("⚠ LOW\n0");
    sheet.getRange("B4:G4").setValue("");
    return;
  }

  var data = sheet.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, lastCol).getValues();

  var totalTitles = 0;
  var totalStock = 0;
  var totalValue = 0;
  var lowStock = 0;
  var languageMap = {};

  data.forEach(function (r) {
    var bookNo = String(r[0] || "").trim();
    var lang = String(r[2] || "").trim();
    var qty = Number(r[5]) || 0;
    var price = Number(r[6]) || 0;

    if (bookNo) totalTitles++;
    totalStock += qty;
    totalValue += (qty * price);
    if (qty <= 5 && bookNo) lowStock++;
    if (lang) {
      if (!languageMap[lang]) languageMap[lang] = { qty: 0 };
      languageMap[lang].qty += qty;
    }
  });

  sheet.getRange("A2:B3").setValue("📘 TOTAL TITLES\n" + totalTitles);
  sheet.getRange("C2:D3").setValue("📦 TOTAL STOCK\n" + totalStock);
  sheet.getRange("E2:F3").setValue("💰 STOCK VALUE\n₹" + totalValue.toLocaleString("en-IN"));
  sheet.getRange("G2:G3").setValue("⚠ LOW\n" + lowStock);

  var languageText = [];
  Object.keys(languageMap)
    .sort(function (a, b) { return (languageMap[b].qty - languageMap[a].qty); })
    .slice(0, 10)
    .forEach(function (lang) { languageText.push(lang + " : " + languageMap[lang].qty); });

  sheet.getRange("B4:G4").setValue(languageText.join("   |   "));
}

function applyOrderRowStyle(sheet, row) {
  if (!sheet || row < ORDER_START_ROW) return;

  var totalCols = ORDER_HEADERS.length;
  var orderNo = sheet.getRange(row, 1).getValue();
  if (!orderNo) return;

  sheet.setRowHeight(row, 55);

  var rowRange = sheet.getRange(row, 1, 1, totalCols);
  rowRange
    .setFontFamily("Arial")
    .setFontSize(10)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#e5e7eb", SpreadsheetApp.BorderStyle.SOLID)
    .setBackground(null);

  sheet.getRange(row, 1)
    .setFontColor("#1d4ed8")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.getRange(row, 2)
    .setFontWeight("bold")
    .setFontColor("#111827");

  sheet.getRange(row, 3)
    .setFontWeight("bold")
    .setFontColor("#7c3aed");

  sheet.getRange(row, 4)
    .setWrap(false)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setFontColor("#334155");

  sheet.getRange(row, 5)
    .setFontSize(11)
    .setNumberFormat("₹#,##0.00")
    .setFontWeight("bold")
    .setFontColor("#1d4ed8");

  sheet.getRange(row, 6)
    .setFontSize(11)
    .setNumberFormat("0.00")
    .setFontColor("#92400e")
    .setFontWeight("bold");

  sheet.getRange(row, 7)
    .setFontSize(11)
    .setNumberFormat("₹#,##0.00")
    .setFontColor("#dc2626")
    .setFontWeight("bold");

  sheet.getRange(row, 8)
    .setFontSize(11)
    .setNumberFormat("₹#,##0.00")
    .setFontColor("#7c3aed")
    .setFontWeight("bold");

  sheet.getRange(row, 9)
    .setFontSize(11)
    .setNumberFormat("₹#,##0.00")
    .setFontColor("#166534")
    .setFontWeight("bold");

  sheet.getRange(row, 10)
    .setWrap(false)
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setFontColor("#7c2d12");

  sheet.getRange(row, 11)
    .setFontColor("#155e75")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  sheet.getRange(row, 12)
    .setFontColor("#166534")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  applyOrderStatusStyle(sheet, row);

  sheet.getRange(row, 15)
    .setNumberFormat("dd-mmm-yyyy hh:mm AM/PM")
    .setFontColor("#0f766e")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");

  sheet.getRange(row, 16)
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  sheet.getRange(row, 17)
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setFontColor("#2563eb");

  sheet.getRange(row, 18)
    .setFontFamily("Courier New")
    .setFontSize(10)
    .setFontColor("#0f172a")
    .setHorizontalAlignment("center")
    .setFontWeight("bold");

  var trackStatusCell = sheet.getRange(row, 19);
  var trackStatus = String(trackStatusCell.getValue() || '').trim();

  trackStatusCell
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setBackground(null);

  if (/delivered/i.test(trackStatus)) {
    trackStatusCell.setFontColor("#166534");
  } else if (/out for delivery/i.test(trackStatus)) {
    trackStatusCell.setFontColor("#1d4ed8");
  } else if (/transit/i.test(trackStatus)) {
    trackStatusCell.setFontColor("#92400e");
  } else if (/returned|rto/i.test(trackStatus)) {
    trackStatusCell.setFontColor("#b91c1c");
  } else if (/hold/i.test(trackStatus)) {
    trackStatusCell.setFontColor("#6d28d9");
  } else {
    trackStatusCell.setFontColor("#334155");
  }

  var paymentCell = sheet.getRange(row, 20);
  var paymentStatus = String(paymentCell.getValue() || '').trim();

  paymentCell
    .setHorizontalAlignment("center")
    .setFontWeight("bold")
    .setBackground(null);

  if (/received/i.test(paymentStatus)) {
    paymentCell.setFontColor("#166534");
  } else if (/pending/i.test(paymentStatus)) {
    paymentCell.setFontColor("#92400e");
  } else {
    paymentCell.setFontColor("#334155");
  }

  var statusValue = String(sheet.getRange(row, 14).getValue()).trim();

  if (statusValue === STATUS.COMPLETE) {
    sheet.getRange(row, 1, 1, sheet.getLastColumn())
      .setBackground("#D4EDDA")
      .setFontColor("#155724");
  } else if (statusValue === STATUS.CANCELLED) {
    sheet.getRange(row, 1, 1, sheet.getLastColumn())
      .setBackground("#F8D7DA")
      .setFontColor("#721C24");
  }
}

function applyOrderStatusStyle(sheet, row) {
  if (!sheet || row < ORDER_START_ROW) return;

  var cell = sheet.getRange(row, 14);
  var status = String(cell.getValue()).trim();

  switch (status) {
    case STATUS.PENDING:
      cell.setBackground("#FFF3CD").setFontColor("#856404");
      break;
    case STATUS.PROGRESS:
      cell.setBackground("#D1ECF1").setFontColor("#0C5460");
      break;
    case STATUS.COMPLETE:
      cell.setBackground("#D4EDDA").setFontColor("#155724");
      break;
    case STATUS.CANCELLED:
      cell.setBackground("#F8D7DA").setFontColor("#721C24");
      break;
    default:
      cell.setBackground("#FFFFFF").setFontColor("#000000");
  }

  cell
    .setFontWeight("bold")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
}

function applyBookConditionalFormatting(sheet) {
  var rules = [];
  var dataStartRow = BOOK_START_ROW;
  var lastRow = sheet.getLastRow();
  if (lastRow < dataStartRow) return;

  var stockRange = sheet.getRange(dataStartRow, 6, lastRow - dataStartRow + 1, 1);

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(5)
      .setBackground("#fee2e2")
      .setFontColor("#991b1b")
      .setBold(true)
      .setRanges([stockRange])
      .build()
  );

  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(5)
      .setBackground("#dcfce7")
      .setFontColor("#166534")
      .setRanges([stockRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}


function showToast(message, type) {
  type = type || "info";
  var icons = {
    success: "✅",
    error: "❌",
    warning: "⚠️",
    info: "ℹ️"
  };

  var html = HtmlService.createTemplateFromFile("Toast");
  html.message = message;
  html.type = type;
  html.icon = icons[type] || "ℹ️";

  SpreadsheetApp.getUi().showModelessDialog(
    html.evaluate().setWidth(450).setHeight(140),
    " "
  );
}

function initialSetup() {
  cleanupTriggers();
  setUpBillableFields();
  createRequiredSheets();
  applyAllSecurity();
  formatOrderSheetLayout();
  formatAllBookSheetLayout();
  applyBookConditionalFormatting(getSheets().BOOKS);
  createBillFolderAndSaveId();
  createBillCleanupTrigger();
  loadLibraryIcons();
  centerHeaderIcons();
  centerAllBookIcons();
  refreshOrderDashboardStats();
  refreshBookDashboardStats();
  createSheetHeaders() 
  showToast("Setup complete", "success");
}

function cleanUp() {
  requireAdmin();                       // only admin can clean
  removeAllSecurity();                  // unlock temporarily
  cleanEmptyRowsFromSheets();           // now cleaning can work
  cleanupTriggers();
  applyAllSecurity();                   // re‑lock
  showToast("Cleanup complete", "success");
}

function cleanEmptyRowsFromSheets() {
  cleanSheetPreservingData(getSheets().ORDERS, ORDER_START_ROW, 1, ORDER_HEADERS.length);
  cleanSheetPreservingData(getSheets().BOOKS, BOOK_START_ROW, 1, BOOK_HEADERS.length);
}

function cleanSheetPreservingData(sheet, startRow, checkCol, totalCols) {
  if (!sheet) return;
  var maxRows = sheet.getMaxRows();
  if (maxRows < startRow) return;

  // 1. Clear ALL formatting in the data range (from startRow to bottom)
  var range = sheet.getRange(startRow, 1, maxRows - startRow + 1, totalCols);
  range.clearFormat();

  // 2. Re-apply formatting to rows that contain a value in the check column (Book No / Order No)
  var values = range.getValues();
  for (var i = 0; i < values.length; i++) {
    var row = startRow + i;
    var checkValue = String(values[i][checkCol - 1] || "").trim();
    if (checkValue !== "") {
      if (sheet.getName() === "Order Book List") {
        applyOrderRowStyle(sheet, row);
      } else if (sheet.getName() === "All Book List") {
        applyBookRowStyle(row);
      }
    }
  }
}

function dailyRefresh() {
  refreshOrderDashboardStats();
  refreshBookDashboardStats();
  showToast("Refreshed", "success");
}

function createRequiredSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var requiredSheets = [
    { name: "All Book List", headers: BOOK_HEADERS },
    { name: "Order Book List", headers: ORDER_HEADERS }
  ];

  requiredSheets.forEach(function (s) {
    var sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      Logger.log("Created Sheet: " + s.name);
    }
  });

  SpreadsheetApp.flush();
  showToast("Sheets ready", "success");
}

function cleanupTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var seen = {};

  triggers.forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (seen[fn]) {
      ScriptApp.deleteTrigger(t);
    } else {
      seen[fn] = true;
    }
  });
}

function createBillCleanupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function (t) {
    if (t.getHandlerFunction() === "deleteOldBillPdfs") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("deleteOldBillPdfs")
    .timeBased()
    .everyHours(1)
    .create();

  Logger.log("Cleanup trigger created");
}

function applyAllSecurity() {
  requireAdmin();
  var ss = getSS();
  removeAllSecurity();
  var me = Session.getEffectiveUser();

  ss.getSheets().forEach(function (sheet) {
    var protection = sheet.protect();
    protection.setDescription("Protected : " + sheet.getName());
    protection.setWarningOnly(false);

    try { protection.setDomainEdit(false); } catch (err) {
      Logger.log("Domain edit not supported: " + err);
    }

    var editors = protection.getEditors();
    if (editors.length > 0) protection.removeEditors(editors);

    protection.addEditor(me);
    protection.setUnprotectedRanges([]);
  });

  SpreadsheetApp.flush();
  showToast("Security applied", "success");
}

function removeAllSecurity() {
  var ss = getSS();

  ss.getSheets().forEach(function (sheet) {
    var rangeProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    rangeProtections.forEach(function (p) { try { p.remove(); } catch (err) { Logger.log(err); } });

    var sheetProtections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    sheetProtections.forEach(function (p) { try { p.remove(); } catch (err) { Logger.log(err); } });
  });

  SpreadsheetApp.flush();
  showToast("Security removed", "success");
}

function unlockEditing() {
  removeAllSecurity();
  showToast("Editing unlocked", "success");
}

function lockEditing() {
  applyAllSecurity();
  showToast("Editing locked", "success");
}

function centerAllBookIcons() {
  var sheet = getSheets().BOOKS;
  if (!sheet) return;

  var iconRow = 5;
  var numIcons = 7;
  var iconSize = 30;

  sheet.setRowHeight(iconRow, 40);
  sheet.getRange(iconRow, 1, 1, numIcons)
    .setBackground("#ccfbf1")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  var images = sheet.getImages();
  images.forEach(function (img) {
    var row = img.getAnchorRow();
    var col = img.getAnchorColumn();
    if (row === iconRow && col >= 1 && col <= numIcons) {
      img.setWidth(iconSize);
      img.setHeight(iconSize);
    }
  });

  SpreadsheetApp.flush();
}

function loadLibraryIcons() {
  const FOLDER_NAME = "DYD_Library_Icons_Backup";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const orderSheet = ss.getSheetByName("Order Book List");
  const allBookSheet = ss.getSheetByName("All Book List");

  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!folders.hasNext()) {
    throw new Error("Folder not found");
  }

  const folder = folders.next();
  const fileMap = {};
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    fileMap[file.getName().toLowerCase()] = file.getId();
  }

  for (let col = 1; col <= 17; col++) {
    const fileName = col + ".png";
    if (!fileMap[fileName]) continue;
    const fileId = fileMap[fileName];
    const imageUrl = "https://drive.google.com/thumbnail?id=" + fileId;
    orderSheet.getRange(3, col).setFormula(`=IMAGE("${imageUrl}")`);
  }

  const letters = ["a", "b", "c", "d", "e", "f", "g"];
  for (let i = 0; i < letters.length; i++) {
    const col = i + 1;
    const fileName = letters[i] + ".png";
    if (!fileMap[fileName]) continue;
    const fileId = fileMap[fileName];
    const imageUrl = "https://drive.google.com/thumbnail?id=" + fileId;
    allBookSheet.getRange(5, col).setFormula(`=IMAGE("${imageUrl}")`);
  }

  SpreadsheetApp.flush();
  Logger.log("Images loaded successfully");
}

// =========================================================
// TRACKING STATUS SYSTEM
// =========================================================

function updateAllTrackingStatuses() {
  var sheet = getSheets().ORDERS;
  var lastRow = sheet.getLastRow();

  if (lastRow < ORDER_START_ROW) {
    showToast('No orders to update', 'warning');
    return;
  }

  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('TRACKING_CANCEL');
  props.deleteProperty('TRACKING_PROGRESS');

  var html = HtmlService
    .createHtmlOutputFromFile('TrackingUpdateProgress')
    .setWidth(420)
    .setHeight(220);

  SpreadsheetApp
    .getUi()
    .showModelessDialog(
      html,
      'Updating Tracking Statuses'
    );
}

function startTrackingUpdateProcessing() {
  var sheet = getSheets().ORDERS;
  var lastRow = sheet.getLastRow();
  var props = PropertiesService.getScriptProperties();

  if (lastRow < ORDER_START_ROW) {
    props.setProperty(
      'TRACKING_PROGRESS',
      JSON.stringify({ current: 0, total: 0, message: 'No rows to process' })
    );
    return;
  }

  var data = sheet
    .getRange(ORDER_START_ROW, 14, lastRow - ORDER_START_ROW + 1, 6)
    .getValues();

  var eligibleRows = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var orderStatus = String(row[0] || '').trim();
    var trackId = String(row[4] || '').trim();
    var trackStatus = String(row[5] || '').trim();
    var isDelivered = /delivered/i.test(trackStatus);

    if (
      orderStatus !== STATUS.CANCELLED &&
      trackId !== '' &&
      !isDelivered
    ) {
      eligibleRows.push({
        row: ORDER_START_ROW + i,
        trackId: trackId
      });
    }
  }

  var total = eligibleRows.length;

  if (total === 0) {
    props.setProperty(
      'TRACKING_PROGRESS',
      JSON.stringify({ current: 0, total: 0, message: 'No rows to process' })
    );
    return;
  }

  var progress = { current: 0, total: total, message: 'Processing 0/' + total };
  props.setProperty('TRACKING_PROGRESS', JSON.stringify(progress));

  for (var j = 0; j < total; j++) {
    if (props.getProperty('TRACKING_CANCEL') === 'true') {
      progress.message = 'Cancelled at ' + j + '/' + total;
      props.setProperty('TRACKING_PROGRESS', JSON.stringify(progress));
      props.deleteProperty('TRACKING_CANCEL');
      return;
    }

    var item = eligibleRows[j];
    var newStatus = checkIndiaPostStatus(item.trackId);

    sheet.getRange(item.row, 19).setValue(newStatus);
    Utilities.sleep(500);

    progress.current = j + 1;
    progress.message = 'Processing ' + (j + 1) + '/' + total;
    props.setProperty('TRACKING_PROGRESS', JSON.stringify(progress));
  }

  progress.message = 'Completed ' + total + '/' + total;
  props.setProperty('TRACKING_PROGRESS', JSON.stringify(progress));
  props.deleteProperty('TRACKING_CANCEL');

  if (eligibleRows.length > 0) {
    applyOrderRowStyle(sheet, eligibleRows[eligibleRows.length - 1].row);
  }
}

function getTrackingUpdateProgress() {
  var props = PropertiesService.getScriptProperties();
  var json = props.getProperty('TRACKING_PROGRESS');
  if (!json) {
    return { current: 0, total: 0, message: 'Initialising...' };
  }
  return JSON.parse(json);
}

function cancelTrackingUpdate() {
  PropertiesService.getScriptProperties().setProperty('TRACKING_CANCEL', 'true');
}

function checkIndiaPostStatus(trackingId) {
  if (!trackingId) return '❌ Invalid Tracking ID';

  var primaryUrl = 'https://myspeedpost.com/track?number=' + encodeURIComponent(trackingId);
  var fallbackUrls = [
    'https://www.ordertracker.com/track/' + encodeURIComponent(trackingId),
    'https://track24.net/?code=' + encodeURIComponent(trackingId),
    'https://parcelsapp.com/en/tracking/' + encodeURIComponent(trackingId),
    'https://www.trackingmore.com/track/en/' + encodeURIComponent(trackingId)
  ];

  var allUrls = [primaryUrl].concat(fallbackUrls);

  for (var i = 0; i < allUrls.length; i++) {
    try {
      var response = UrlFetchApp.fetch(allUrls[i], {
        muteHttpExceptions: true,
        followRedirects: true,
        validateHttpsCertificates: true,
        timeout: 15000,
        headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' }
      });

      if (response.getResponseCode() !== 200) continue;

      var html = response.getContentText();
      var lowerHtml = html.toLowerCase();

      var jsonStatus = extractFromJsonLd(html);
      if (jsonStatus) return formatStatus(jsonStatus);

      var htmlStatus = extractFromHtmlPatterns(html, lowerHtml);
      if (htmlStatus) return formatStatus(htmlStatus);

      var tableStatus = extractLatestStatusFromHtml(lowerHtml);
      if (tableStatus) return formatStatus(tableStatus);

      var keywordStatus = extractFromKeywords(lowerHtml);
      if (keywordStatus) return keywordStatus;

    } catch (e) { Logger.log('Tracking fetch failed: ' + e); }
  }
  return '❌ Could not fetch status';
}

function extractFromJsonLd(html) {
  var matches = html.match(/<script type="application\/ld\+json"[^>]*>([\s\S]*?)<\/script>/gi);
  if (!matches) return null;

  for (var i = 0; i < matches.length; i++) {
    try {
      var cleaned = matches[i].replace(/<script[^>]*>/i, '').replace(/<\/script>/i, '');
      var data = JSON.parse(cleaned);
      var status = data?.deliveryStatus || data?.status || data?.tracking?.status || data?.orderStatus;
      if (status && typeof status === 'string') return status.trim();
    } catch (e) { }
  }
  return null;
}

function extractFromHtmlPatterns(html, lowerHtml) {
  var m1 = html.match(/<span[^>]*data-testid="shipment-status"[^>]*>(.*?)<\/span>/i);
  if (m1) return m1[1].replace(/<[^>]*>/g, '').trim();

  var m2 = html.match(/<td[^>]*>\s*status\s*<\/td>\s*<td[^>]*>(.*?)<\/td>/i);
  if (m2) return m2[1].replace(/<[^>]*>/g, '').trim();

  var m3 = html.match(/<[^>]*class="[^"]*status[^"]*"[^>]*>(.*?)<\//i);
  if (m3) {
    var text = m3[1].replace(/<[^>]*>/g, '').trim();
    if (text.length > 2 && text.length < 150) return text;
  }
  return null;
}

function extractLatestStatusFromHtml(lowerHtml) {
  var tbodyMatch = lowerHtml.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/);
  if (!tbodyMatch) return null;

  var rows = tbodyMatch[1].match(/<tr[^>]*>[\s\S]*?<\/tr>/g);
  if (!rows || rows.length === 0) return null;

  var firstRow = rows[0];
  var cells = firstRow.match(/<td[^>]*>([\s\S]*?)<\/td>/gi);
  if (!cells) return null;

  var statusText = '';
  for (var i = 0; i < cells.length; i++) {
    var text = cells[i].replace(/<[^>]*>/g, '').trim();
    var lowerText = text.toLowerCase();

    if (
      lowerText.indexOf('delivered') !== -1 ||
      lowerText.indexOf('transit') !== -1 ||
      lowerText.indexOf('dispatched') !== -1 ||
      lowerText.indexOf('out for delivery') !== -1 ||
      lowerText.indexOf('booked') !== -1 ||
      lowerText.indexOf('hold') !== -1 ||
      lowerText.indexOf('attempt') !== -1 ||
      lowerText.indexOf('returned to sender') !== -1 ||
      lowerText.indexOf('return to sender') !== -1
    ) {
      return text;
    }

    if (text.length > statusText.length) {
      statusText = text;
    }
  }
  return statusText || null;
}

function extractFromKeywords(lowerHtml) {
  if (/delivered/.test(lowerHtml)) return '✅ Delivered';
  if (/out for delivery/.test(lowerHtml)) return '🚚 Out for Delivery';
  if (/delivery attempted|failed attempt/.test(lowerHtml)) return '⚠️ Delivery Failed Attempt';
  if (/available for pickup/.test(lowerHtml)) return '📬 Available for Pickup';
  if (/received at/.test(lowerHtml)) return '📥 Received at Hub';
  if (/returned to sender|return to sender|rto delivered/i.test(lowerHtml)) return '↩️ Returned';
  if (/held|on hold/.test(lowerHtml)) return '⏸️ On Hold';
  if (/in transit|transit/.test(lowerHtml)) return '📦 In Transit';
  if (/dispatched/.test(lowerHtml)) return '📤 Dispatched';
  if (/booked|info received/.test(lowerHtml)) return '📦 Booked';
  return null;
}

function formatStatus(raw) {
  if (!raw) return 'ℹ️ Status unavailable';
  raw = raw.toLowerCase().trim();

  if (/delivered/.test(raw)) return '✅ Delivered';
  if (/out for delivery/.test(raw)) return '🚚 Out for Delivery';
  if (/dispatched/.test(raw)) return '📤 Dispatched';
  if (/booked|info received/.test(raw)) return '📦 Booked';
  if (/in transit|transit/.test(raw)) return '📦 In Transit';
  if (/returned|rto/.test(raw)) return '↩️ Returned';
  if (/held|on hold/.test(raw)) return '⏸️ On Hold';
  if (/received at/.test(raw)) return '📥 Received at Hub';
  if (/failed attempt/.test(raw)) return '⚠️ Delivery Failed Attempt';
  if (/available for pickup/.test(raw)) return '📬 Available for Pickup';

  return ('ℹ️ ' + raw.charAt(0).toUpperCase() + raw.slice(1));
}

// =========================================================
// UI FUNCTIONS
// =========================================================

function showSetTrackingIdDialog() {
  var sheet = getSheets().ORDERS;
  var row = sheet.getActiveRange().getRow();

  if (row < ORDER_START_ROW) {
    showToast('Please select a valid order row', 'warning');
    return;
  }

  var html = HtmlService
    .createHtmlOutputFromFile('SetTrackingID')
    .setWidth(500)
    .setHeight(280);

  SpreadsheetApp
    .getUi()
    .showModalDialog(html, '📬 Update Tracking ID');
}

function getTrackingIdData() {
  var sheet = getSheets().ORDERS;
  var row = sheet.getActiveRange().getRow();

  if (row < ORDER_START_ROW) throw new Error('No valid order selected');

  var trackId = String(sheet.getRange(row, 18).getValue() || '').trim();

  return {
    row: row,
    trackingId: trackId
  };
}

function saveTrackingId(row, trackingId) {
  var sheet = getSheets().ORDERS;

  if (row < ORDER_START_ROW) throw new Error('Invalid row');

  trackingId = String(trackingId || '').trim();

  sheet.getRange(row, 18).setValue(trackingId);

  var currentStatus = String(sheet.getRange(row, 14).getValue()).trim();

  if (
    currentStatus !== STATUS.CANCELLED &&
    currentStatus !== STATUS.COMPLETE
  ) {
    if (trackingId !== '') {
      sheet.getRange(row, 14).setValue(STATUS.PROGRESS);
    } else {
      sheet.getRange(row, 14).setValue(STATUS.PENDING);
    }

    applyOrderRowStyle(sheet, row);
  }

  return 'Tracking ID saved successfully!';
}

function markPaymentPending() {
  setPaymentStatus('⏳ Pending');
}

function markPaymentReceived() {
  setPaymentStatus('✅ Received');
}

function setPaymentStatus(status) {
  var sheet = getSheets().ORDERS;
  var range = sheet.getActiveRange();

  if (!range) {
    showToast('No row selected', 'warning');
    return;
  }

  var row = range.getRow();

  if (row < ORDER_START_ROW) {
    showToast('Select a valid order row', 'warning');
    return;
  }

  var orderNo = String(sheet.getRange(row, 1).getValue()).trim();

  if (!orderNo) {
    showToast('No order found in selected row', 'error');
    return;
  }

  var cell = sheet.getRange(row, 20);
  cell.setValue(status);

  if (status.indexOf('Received') !== -1) {
    cell.setBackground('#DCFCE7').setFontColor('#166534').setFontWeight('bold');
  } else {
    cell.setBackground('#FEF3C7').setFontColor('#92400E').setFontWeight('bold');
  }

  applyOrderRowStyle(sheet, row);
  showToast('Payment Status Updated: ' + status, 'success');
}
// ==============================================
// CUSTOMER DATABASE LOOKUP (Internal Sheet)
// ==============================================

function lookupCustomerByMobile(mobile) {
  mobile = cleanIndianMobile(mobile);
  if (!mobile) return { found: false };

  var ss = getSS();
  var sheet = ss.getSheetByName("Customer Database");
  if (!sheet) {
    Logger.log("Customer Database sheet not found");
    return { found: false };
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { found: false };   // only header row

  // Find columns by header name (case‑sensitive)
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colMobile  = headers.indexOf("Mobile") + 1;
  var colName    = headers.indexOf("Name") + 1;
  var colAddress = headers.indexOf("Address") + 1;
  var colArea    = headers.indexOf("Area") + 1;
  var colCity    = headers.indexOf("City") + 1;
  var colState   = headers.indexOf("State") + 1;
  var colZip     = headers.indexOf("Zip") + 1;

  if (colMobile === 0 || colName === 0 || colAddress === 0) {
    Logger.log("Required column headers (Mobile, Name, Address) missing in Customer Database sheet");
    return { found: false };
  }

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var dbMobile = cleanIndianMobile(row[colMobile - 1]);
    if (dbMobile === mobile) {
      var name    = String(row[colName - 1] || "").trim();
      var home    = String(row[colAddress - 1] || "").trim();
      var area    = colArea    > 0 ? String(row[colArea - 1] || "").trim() : "";
      var city    = colCity    > 0 ? String(row[colCity - 1] || "").trim() : "";
      var state   = colState   > 0 ? String(row[colState - 1] || "").trim() : "";
      var zip     = colZip     > 0 ? String(row[colZip - 1] || "").trim() : "";

      // Clean the home field – remove parts already in area/city/state/zip
      var cleanedHome = cleanAddressHome(home, area, city, state, zip);

      // Build location parts (area, city, state) with deduplication
      var locationParts = [];
      [area, city, state].forEach(function(part) {
        if (part && locationParts.indexOf(part) === -1) {
          locationParts.push(part);
        }
      });

      // Handle zip: attach to the last location part with a dash, if possible
      if (zip) {
        if (locationParts.length > 0) {
          // Replace the last location part with "lastPart - zip"
          var lastIdx = locationParts.length - 1;
          locationParts[lastIdx] = locationParts[lastIdx] + " - " + zip;
        } else {
          // No location parts – just add the zip as a separate part
          locationParts.push(zip);
        }
      }

      // Assemble final address
      var finalParts = [];
      if (cleanedHome) finalParts.push(cleanedHome);
      finalParts = finalParts.concat(locationParts);

      return {
        found: true,
        name: name,
        address: finalParts.join(", ")
      };
    }
  }
  return { found: false };
}


/**
 * Removes from 'home' any text already in area/city/state/zip.
 * Case‑insensitive. After removal, extra spaces/punctuation are cleaned up.
 */
function cleanAddressHome(home, area, city, state, zip) {
  var fieldsToRemove = [area, city, state, zip];
  var cleaned = home;

  fieldsToRemove.forEach(function(fieldValue) {
    if (!fieldValue) return;
    var escaped = fieldValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    var regex = new RegExp(escaped, 'gi');
    cleaned = cleaned.replace(regex, '');
  });

  cleaned = cleaned.replace(/\s{2,}/g, ' ');
  cleaned = cleaned.replace(/^\s*[, -]+\s*|\s*[, -]+\s*$/g, '');
  cleaned = cleaned.replace(/\s*,\s*,/g, ',');
  cleaned = cleaned.trim();

  return cleaned;
}
function lookupCustomerByName(searchText) {
  searchText = String(searchText || '').trim().toLowerCase();
  if (!searchText) return [];

  var ss = getSS();
  var sheet = ss.getSheetByName("Customer Database");
  if (!sheet) return [];

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];   // only header row

  // Find column indices (same as in your existing function)
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colName    = headers.indexOf("Name") + 1;
  var colMobile  = headers.indexOf("Mobile") + 1;
  var colAddress = headers.indexOf("Address") + 1;
  var colArea    = headers.indexOf("Area") + 1;
  var colCity    = headers.indexOf("City") + 1;
  var colState   = headers.indexOf("State") + 1;
  var colZip     = headers.indexOf("Zip") + 1;

  if (colName === 0) return [];   // "Name" column mandatory

  var data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  var results = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var name = String(row[colName - 1] || "").trim();
    if (!name) continue;

    if (name.toLowerCase().indexOf(searchText) !== -1) {
      var mobile  = colMobile  > 0 ? cleanIndianMobile(row[colMobile - 1]) : "";
      var home    = colAddress > 0 ? String(row[colAddress - 1] || "").trim() : "";
      var area    = colArea    > 0 ? String(row[colArea - 1] || "").trim() : "";
      var city    = colCity    > 0 ? String(row[colCity - 1] || "").trim() : "";
      var state   = colState   > 0 ? String(row[colState - 1] || "").trim() : "";
      var zip     = colZip     > 0 ? String(row[colZip - 1] || "").trim() : "";

      // Build address exactly as in lookupCustomerByMobile
      var cleanedHome = cleanAddressHome(home, area, city, state, zip);
      var locationParts = [];
      [area, city, state].forEach(function(part) {
        if (part && locationParts.indexOf(part) === -1) locationParts.push(part);
      });
      if (zip) {
        if (locationParts.length > 0)
          locationParts[locationParts.length - 1] += " - " + zip;
        else
          locationParts.push(zip);
      }
      var finalParts = [];
      if (cleanedHome) finalParts.push(cleanedHome);
      finalParts = finalParts.concat(locationParts);
      var fullAddress = finalParts.join(", ");

      results.push({
        name: name,
        mobile: mobile,
        address: fullAddress
      });
    }
    // Limit to 10 matches to keep the UI fast
    if (results.length >= 10) break;
  }

  return results;
}
// =========================================================
// BULK IMPORT FROM "StockData" SHEET
// =========================================================

// =========================================================
// BULK IMPORT FROM "StockData" SHEET (CREATE ONLY)
// =========================================================

var IMPORT_SHEET_NAME = "StockData";

function adminImportBooksFromSheet() {
  requireAdmin();
  var ss = getSS();
  var sheet = ss.getSheetByName(IMPORT_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Sheet "' + IMPORT_SHEET_NAME + '" not found. Please create it first.'
    );
    return;
  }
  var html = HtmlService.createHtmlOutputFromFile("ImportFromSheet")
    .setWidth(900)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(
    html,
    "📤 Import Books from " + IMPORT_SHEET_NAME
  );
}

/**
 * Reads the StockData sheet, ensures a "Book No" and "Status" column exist,
 * and returns headers, preview rows, and total count.
 */
function getImportPreview() {
  var ss = getSS();
  var sheet = ss.getSheetByName(IMPORT_SHEET_NAME);
  if (!sheet) throw new Error("Sheet not found");

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) throw new Error("Sheet is empty");

  var lastCol = sheet.getLastColumn();
  var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = allData[0].map(function (h) {
    return String(h).trim();
  });

  // Remove trailing empty rows
  var rows = allData.slice(1);
  while (
    rows.length &&
    rows[rows.length - 1].every(function (c) { return String(c).trim() === ""; })
  ) {
    rows.pop();
  }

  // Column indices (case‑insensitive)
  var colMap = {};
  headers.forEach(function (h, i) {
    colMap[h.toLowerCase()] = i;
  });

  // Required columns (Book No is optional – we ignore it and generate new ones)
  if (colMap["book name"] === undefined) {
    throw new Error('Missing required column: "Book Name"');
  }

  // Ensure "Status" column exists
  if (colMap["status"] === undefined) {
    sheet.insertColumnAfter(lastCol);
    sheet.getRange(1, lastCol + 1).setValue("Status");
    lastCol = sheet.getLastColumn();
    allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    headers = allData[0].map(String);
    rows = allData.slice(1);
    colMap = {};
    headers.forEach(function (h, i) { colMap[h.toLowerCase()] = i; });
  }

  // Ensure "Book No" column exists (we need to write the generated numbers back)
  if (colMap["book no"] === undefined) {
    sheet.insertColumnAfter(1); // insert right after column A (position 2)
    sheet.getRange(1, 2).setValue("Book No");
    lastCol = sheet.getLastColumn();
    allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    headers = allData[0].map(String);
    rows = allData.slice(1);
    colMap = {};
    headers.forEach(function (h, i) { colMap[h.toLowerCase()] = i; });
  }

  // Build row objects (ignore existing Book No, we'll generate)
  var rowObjects = rows.map(function (row) {
    return {
      rowNum: null,
      name:        String(row[colMap["book name"]] || "").trim(),
      language:    colMap["language"] !== undefined
        ? String(row[colMap["language"]] || "").trim() : "",
      publication: colMap["publication"] !== undefined
        ? String(row[colMap["publication"]] || "").trim() : "",
      author:      colMap["author"] !== undefined
        ? String(row[colMap["author"]] || "").trim() : "",
      stock:       colMap["stock"] !== undefined
        ? Number(row[colMap["stock"]]) || 0 : 0,
      price:       colMap["price"] !== undefined
        ? Number(row[colMap["price"]]) || 0 : 0,
      status:      String(row[colMap["status"]] || "").trim()
    };
  }).filter(function (r) { return r.name !== ""; });

  // Set row numbers (2‑based, first data row is row 2)
  for (var i = 0; i < rowObjects.length; i++) {
    rowObjects[i].rowNum = i + 2;
  }

  // Cache for batch processing
  var cache = CacheService.getScriptCache();
  cache.put("IMPORT_DATA_" + IMPORT_SHEET_NAME, JSON.stringify(rowObjects), 600);

  // Preview uses current headers (including the added Book No column)
  return {
    headers: headers,
    previewRows: rowObjects.slice(0, 20).map(function (r) {
      // Return a row in the same order as headers (columns)
      return headers.map(function (h) {
        var key = h.toLowerCase();
        if (key === "book no") return "";       // empty, will be generated
        if (key === "book name") return r.name;
        if (key === "language") return r.language;
        if (key === "publication") return r.publication;
        if (key === "author") return r.author;
        if (key === "stock") return r.stock;
        if (key === "price") return r.price;
        if (key === "status") return r.status;
        return "";
      });
    }),
    total: rowObjects.length
  };
}

/**
 * Processes a batch of rows – adds new books only, checks for duplicates.
 * Writes "Done", "Duplicate", or "Error" back to the StockData sheet.
 */
function processImportBatch(offset, batchSize) {
  var cache = CacheService.getScriptCache();
  var json = cache.get("IMPORT_DATA_" + IMPORT_SHEET_NAME);
  if (!json) throw new Error("Cache expired. Please preview again.");

  var allRows = JSON.parse(json);
  var total = allRows.length;
  var end = Math.min(offset + batchSize, total);
  var batch = allRows.slice(offset, end);

  var bookSheet = getSheets().BOOKS;
  var importSheet = getSS().getSheetByName(IMPORT_SHEET_NAME);
  var bookMap = buildBookMap();   // existing books for duplicate check

  var added = 0;
  var duplicates = 0;
  var errors = [];

  // Find column indices in importSheet (needed to write Book No and Status)
  var importHeaders = importSheet
    .getRange(1, 1, 1, importSheet.getLastColumn())
    .getValues()[0];
  var bookNoCol = 1; // default fallback
  var statusCol = importSheet.getLastColumn();
  for (var c = 0; c < importHeaders.length; c++) {
    var h = String(importHeaders[c]).toLowerCase().trim();
    if (h === "book no") bookNoCol = c + 1;
    if (h === "status") statusCol = c + 1;
  }

  for (var i = 0; i < batch.length; i++) {
    var row = batch[i];
    try {
      // Duplicate check: same Name + Language in existing books
      if (isDuplicateBook(row, bookMap)) {
        importSheet.getRange(row.rowNum, statusCol).setValue("Duplicate");
        duplicates++;
        continue;
      }

      // Generate new Book No (same format as saveBook)
      var newBookNo = generateBookNo();
      var newRow = bookSheet.getLastRow() + 1;
      if (newRow < BOOK_START_ROW) newRow = BOOK_START_ROW;

      bookSheet.getRange(newRow, 1, 1, 7).setValues([[
        newBookNo, row.name, row.language, row.publication,
        row.author, row.stock, row.price
      ]]);

      // Write generated Book No back to import sheet
      importSheet.getRange(row.rowNum, bookNoCol).setValue(newBookNo);
      importSheet.getRange(row.rowNum, statusCol).setValue("Done");

      // Add to bookMap for subsequent duplicate checks within this batch
      bookMap[newBookNo] = {
        row: newRow,
        name: row.name,
        language: row.language
      };
      added++;
    } catch (e) {
      importSheet
        .getRange(row.rowNum, statusCol)
        .setValue("Error: " + e.message);
      errors.push("Row " + row.rowNum + ": " + e.message);
    }
  }

  SpreadsheetApp.flush();

  // On final batch, apply formatting and refresh stats
  if (offset + batchSize >= total) {
    applyBookFormattingToAllRows();
    refreshBookDashboardStats();
  }

  return {
    processed: end,
    total: total,
    done: end >= total,
    message:
      "Added: " + added + ", Duplicates: " + duplicates +
      (errors.length ? " | " + errors.length + " errors" : ""),
    errors: errors.slice(0, 5)
  };
}

/**
 * Check for duplicate by Name + Language (case‑insensitive) in the book map.
 */
function isDuplicateBook(row, bookMap) {
  var nameLow = row.name.toLowerCase().trim();
  var langLow = row.language.toLowerCase().trim();
  for (var key in bookMap) {
    var b = bookMap[key];
    if (
      b.name.toLowerCase().trim() === nameLow &&
      b.language.toLowerCase().trim() === langLow
    ) {
      return true;
    }
  }
  return false;
}

/**
 * Generate a Book No in the same format as saveBook():
 * "B" + yyyyMMdd-HHmmss
 */
function generateBookNo() {
  var now = new Date();
  var datePart = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss");
  return "B" + datePart;
}

/**
 * Apply book row style to all data rows.
 */
function applyBookFormattingToAllRows() {
  var sheet = getSheets().BOOKS;
  var lastRow = sheet.getLastRow();
  for (var r = BOOK_START_ROW; r <= lastRow; r++) {
    var bookNo = String(sheet.getRange(r, 1).getValue()).trim();
    if (bookNo) {
      applyBookRowStyle(r);
    }
  }
}
/**
 * Creates or updates the headers for "Customer Database" and "StockData" sheets.
 * Run this once to set up the sheets.
 */
function createSheetHeaders() {
  // ---------- 1. Customer Database ----------
  const customerSheetName = "Customer Database";
  let customerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(customerSheetName);
  if (!customerSheet) {
    customerSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(customerSheetName);
  }

  const customerHeaders = ["Mobile", "Name", "Address", "Area", "City", "State", "Zip"];

  // Place headers in row 1
  customerSheet.getRange(1, 1, 1, customerHeaders.length).setValues([customerHeaders]);

  // Formatting
  customerSheet.getRange(1, 1, 1, customerHeaders.length)
    .setFontWeight("bold")
    .setBackground("#FF9933")   // orange theme from your bill
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  customerSheet.setFrozenRows(1);

  // Auto‑resize columns (optional)
  customerSheet.autoResizeColumns(1, customerHeaders.length);

  // ---------- 2. StockData ----------
  const stockSheetName = "StockData";
  let stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(stockSheetName);
  if (!stockSheet) {
    stockSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(stockSheetName);
  }

  const stockHeaders = [
    "Book Name","Language", "Publication",
    "Author", "Stock", "Price", "Status"
  ];

  stockSheet.getRange(1, 1, 1, stockHeaders.length).setValues([stockHeaders]);

  stockSheet.getRange(1, 1, 1, stockHeaders.length)
    .setFontWeight("bold")
    .setBackground("#1d4ed8")   // blue theme
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  stockSheet.setFrozenRows(1);
  stockSheet.autoResizeColumns(1, stockHeaders.length);
}
/**
 * Ensures all four core sheets exist.
 * Creates any missing sheet without adding headers.
 * Call your existing header/formatter functions after this.
 */
function ensureAllSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const requiredSheetNames = [
    "All Book List",
    "Order Book List",
    "Customer Database",
    "StockData"
  ];

  requiredSheetNames.forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
      Logger.log("Created missing sheet: " + name);
    }
  });
  SpreadsheetApp.flush();
  showToast("All required sheets exist", "success");
}
