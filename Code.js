/**
 * Diamond Ledger & Portfolio Management System - Backend Engine
 * Supports full CRUD for Purchases, Sales, and Consolidated Sales Reports.
 * Features automated Google Sheet synchronization, Script Lock concurrency protection,
 * and intelligent normalized header mapping.
 */

function doGet() {
  let template;
  try {
    template = HtmlService.createTemplateFromFile('index');
  } catch (e) {
    template = HtmlService.createTemplateFromFile('Index');
  }
  return template.evaluate()
      .setTitle('Diamond Ledger & Portfolio Management')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Normalizes any string to alphanumeric lowercase for ultra-flexible column matching
 * e.g. "GIA No" -> "giano", "ADV D AMT" -> "advdamt", "Profit %" -> "profit"
 */
function normKey(str) {
  return String(str || "").replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
}

/* ---- FIX HELPERS (added) ---- */
function num_(v)    { const n = parseFloat(v); return isNaN(n) ? 0 : n; }
function round2_(v) { return Math.round((Number(v) + Number.EPSILON) * 100) / 100; }
function round3_(v) { return Math.round((Number(v) + Number.EPSILON) * 1000) / 1000; }

/** Finds the "Profit %" column by its raw header (normKey would turn it into "profit"). */
function getProfitPctColIdx_(headers) {
  for (let c = 0; c < headers.length; c++) {
    const raw = String(headers[c]).trim().toLowerCase();
    if (raw.indexOf('%') !== -1 && normKey(raw) === 'profit') return c;
    if (normKey(raw) === 'profitpct') return c;
  }
  return -1;
}

/** Adds the "BillPct" column to an existing Purchases sheet if it is missing. */
function ensureBillPctColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if (headers.some(h => normKey(h) === "billpct")) return;
  const col = lastCol + 1;
  sheet.getRange(1, col).setValue("BillPct");
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, col, lastRow - 1, 1).setValue(0);
}

/** Adds the "Type" column to an existing Purchases sheet if it is missing, and backfills empty cells. */
function ensureTypeColumn_(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  let typeColIdx = -1;
  for (let c = 0; c < headers.length; c++) {
    const nk = normKey(headers[c]);
    if (nk === "type" || nk === "certtype" || nk === "certificatetype") {
      typeColIdx = c;
      break;
    }
  }

  if (typeColIdx === -1) {
    const col = lastCol + 1;
    sheet.getRange(1, col).setValue("Type");
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, col, lastRow - 1, 1).setValue("IGI");
  } else {
    // If column already exists, backfill any empty cells with "IGI" so all stones have a valid Type
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const range = sheet.getRange(2, typeColIdx + 1, lastRow - 1, 1);
      const vals = range.getValues();
      let changed = false;
      for (let r = 0; r < vals.length; r++) {
        if (!vals[r][0] || String(vals[r][0]).trim() === "") {
          vals[r][0] = "IGI";
          changed = true;
        }
      }
      if (changed) {
        range.setValues(vals);
      }
    }
  }
}

/** Only these emails may change the connected Google Sheet. Put your email here. */
const OWNER_EMAILS = ["your-email@gmail.com"];
const ENFORCE_OWNER_CHECK = false; // set true only if deployed as "Execute as: User accessing the web app"

function assertOwner_() {
  if (!ENFORCE_OWNER_CHECK) return;
  const email = (Session.getActiveUser().getEmail() || "").toLowerCase();
  if (OWNER_EMAILS.map(e => e.toLowerCase()).indexOf(email) === -1) {
    throw new Error("Only the owner can change the connected Google Sheet.");
  }
}

/**
 * Standard headers matching the exact diamond business structure
 */
const PURCHASE_HEADERS = [
  "PurId", "GIA No", "Purchase Date", "From Name", "Size", "Shape", "Color", 
  "Quality", "Flor.", "Cut", "Pol.", "Sym.", "Measur", "Less", "PerCarat", 
  "DAmount", "Rate", "IndAmount", "BType", "PaymentDay", "Broker", "User", "Status", "BillPct", "Type"
];

const SALE_HEADERS = [
  "SaleId", "PurId", "Date", "ClientName", "Country", "SalePerCarat", "SaleLess", 
  "SaleDFinal", "ADV D AMT", "DUE AMT", "SaleRemark", "User", "SaleDProfit", "Profit %"
];

/**
 * =============================================================================
 * GOOGLE SHEET CONNECTION CONFIGURATION
 * =============================================================================
 * If your script is created via Google Sheets (Extensions -> Apps Script), leave this empty.
 * If your script is standalone at script.google.com, you can paste your Google Sheet URL or ID below:
 */
const SPREADSHEET_ID_OR_URL = "";

/**
 * Robust Spreadsheet Resolver:
 * 1. Checks active spreadsheet (if container-bound).
 * 2. Checks configured SPREADSHEET_ID_OR_URL constant.
 * 3. Checks the SPREADSHEET_ID saved for THIS user (User Properties), so every user gets their own sheet.
 * 4. If none found, automatically creates a new Google Sheet in Drive and links it!
 */
function getSpreadsheet() {
  // 1. Container-bound check
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;

  // 2. Configured constant check
  if (typeof SPREADSHEET_ID_OR_URL !== 'undefined' && SPREADSHEET_ID_OR_URL.trim() !== "") {
    let target = SPREADSHEET_ID_OR_URL.trim();
    let matches = target.match(/\/d\/([a-zA-Z0-9-_]+)/);
    let id = matches && matches[1] ? matches[1] : target;
    try {
      return SpreadsheetApp.openById(id);
    } catch(e) {
      Logger.log("Failed to open configured spreadsheet: " + e.message);
    }
  }

  // 3. Stored property check
  let savedId = PropertiesService.getUserProperties().getProperty('SPREADSHEET_ID');
  if (savedId) {
    try {
      return SpreadsheetApp.openById(savedId);
    } catch(e) {
      Logger.log("Stored spreadsheet ID inaccessible: " + e.message);
    }
  }

  // 4. Auto-create new Google Sheet database if none linked
  try {
    let newSs = SpreadsheetApp.create("Diamond Portfolio & Ledger Database");
    PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', newSs.getId());
    return newSs;
  } catch(e) {
    throw new Error(
      "No Google Sheet linked. Please paste your Google Sheet URL into SPREADSHEET_ID_OR_URL in Code.gs, " +
      "or open your Google Sheet and click Extensions -> Apps Script."
    );
  }
}

/**
 * Returns or initializes a sheet with structured headers
 */
function getSheetByName(name) {
  let ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === "Purchases") {
      sheet.appendRow(PURCHASE_HEADERS);
      sheet.setFrozenRows(1);
    } else if (name === "Sales") {
      sheet.appendRow(SALE_HEADERS);
      sheet.setFrozenRows(1);
    }
  }
  return sheet;
}

/**
 * Dynamic Sheet Connector: Allows employee or owner to link any shared Google Sheet directly
 */
function setConnectedSpreadsheet(urlOrId) {
  assertOwner_();
  if (!urlOrId || !urlOrId.trim()) {
    throw new Error("Please enter a valid Google Sheet URL or ID.");
  }
  let target = urlOrId.trim();
  let matches = target.match(/\/d\/([a-zA-Z0-9-_]+)/);
  let id = matches && matches[1] ? matches[1] : target;

  let ss = SpreadsheetApp.openById(id);
  PropertiesService.getUserProperties().setProperty('SPREADSHEET_ID', id);

  // Initialize standard tabs if they don't exist
  getSheetByName("Purchases");
  getSheetByName("Sales");

  return {
    success: true,
    sheetUrl: ss.getUrl(),
    purchases: getAllPurchases(),
    sales: getAllSales()
  };
}

/**
 * Robust ID Generator: Scans column A for the maximum numeric suffix.
 * Prevents race conditions, sorting bugs, and NaN errors.
 */
function getNextId(sheet, prefix, defaultStart) {
  let data = sheet.getDataRange().getValues();
  if (data.length <= 1) return prefix + "-" + defaultStart;

  let maxNum = defaultStart - 1;
  let regex = new RegExp("^" + prefix + "-(\\d+)$");

  for (let i = 1; i < data.length; i++) {
    let val = String(data[i][0]).trim();
    let match = val.match(regex);
    if (match) {
      let num = parseInt(match[1], 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  return prefix + "-" + (maxNum + 1);
}

/**
 * Safe date string normalization (yyyy-MM-dd)
 */
function formatDateSafe(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  let str = String(val).trim();
  // Handle DD-MM-YYYY format
  if (/^\d{2}-\d{2}-\d{4}$/.test(str)) {
    let parts = str.split('-');
    return parts[2] + "-" + parts[1] + "-" + parts[0];
  }
  return str;
}

/**
 * READ Operations: Fetches all purchases, sales, and session info
 */
function getInitialData() {
  let ss = getSpreadsheet();
  return {
    purchases: getAllPurchases(),
    sales: getAllSales(),
    currentUser: Session.getActiveUser().getEmail() || "",
    sheetUrl: ss ? ss.getUrl() : ""
  };
}

function getAllPurchases() {
  let sheet = getSheetByName("Purchases");
  ensureBillPctColumn_(sheet);
  ensureTypeColumn_(sheet);
  let data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let rawHeaders = data[0];
  let purIdColIdx = -1;
  let statusColIdx = -1;

  for (let c = 0; c < rawHeaders.length; c++) {
    let nk = normKey(rawHeaders[c]);
    if (nk === "purid") purIdColIdx = c;
    if (nk === "status") statusColIdx = c;
  }

  // Auto-repair: Insert PurId column at col 1 if missing
  if (purIdColIdx === -1) {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("PurId");
    data = sheet.getDataRange().getValues();
    rawHeaders = data[0];
    purIdColIdx = 0;
  }

  // Auto-repair: Backfill PurId for any empty cells (never reuses an existing number)
  let maxP = 1000;
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][purIdColIdx]).trim().match(/^P-(\d+)$/);
    if (m) maxP = Math.max(maxP, parseInt(m[1], 10));
  }
  for (let i = 1; i < data.length; i++) {
    let currentId = String(data[i][purIdColIdx] || "").trim();
    if (!currentId) {
      maxP++;
      let generatedId = "P-" + maxP;
      sheet.getRange(i + 1, purIdColIdx + 1).setValue(generatedId);
      data[i][purIdColIdx] = generatedId;
    }
  }

  // Auto-repair: Insert Status column if missing
  if (statusColIdx === -1) {
    let lastCol = rawHeaders.length + 1;
    sheet.getRange(1, lastCol).setValue("Status");
    statusColIdx = lastCol - 1;
    for (let i = 1; i < data.length; i++) {
      sheet.getRange(i + 1, lastCol).setValue("IN STOCK");
    }
    data = sheet.getDataRange().getValues();
    rawHeaders = data[0];
  }

  return data.slice(1).map((row, index) => {
    let rowMap = {};
    rawHeaders.forEach((h, i) => {
      rowMap[normKey(h)] = row[i];
    });

    let purDate = rowMap["purchasedate"] || rowMap["date"] || "";
    let purId = rowMap["purid"] ? String(rowMap["purid"]).trim() : "P-" + (1001 + index);

    return {
      _rowIndex: index + 2,
      PurId: purId,
      GIANo: String(rowMap["giano"] || rowMap["gianumber"] || ""),
      PurchaseDate: formatDateSafe(purDate),
      FromName: String(rowMap["fromname"] || rowMap["fromcomp"] || ""),
      Size: rowMap["size"] || rowMap["sizecarat"] || "",
      Shape: String(rowMap["shape"] || ""),
      Color: String(rowMap["color"] || ""),
      Quality: String(rowMap["quality"] || ""),
      Flor: String(rowMap["flor"] || rowMap["fluor"] || ""),
      Cut: String(rowMap["cut"] || ""),
      Pol: String(rowMap["pol"] || ""),
      Sym: String(rowMap["sym"] || ""),
      Measur: String(rowMap["measur"] || rowMap["measurements"] || ""),
      Less: rowMap["less"] !== undefined && rowMap["less"] !== "" ? rowMap["less"] : 0,
      PerCarat: rowMap["percarat"] || "",
      DAmount: rowMap["damount"] || rowMap["amount"] || "",
      Rate: rowMap["rate"] || "",
      IndAmount: rowMap["indamount"] || "",
      BType: String(rowMap["btype"] || rowMap["bill"] || ""),
      PaymentDay: rowMap["paymentday"] || "",
      Broker: String(rowMap["broker"] || ""),
      User: String(rowMap["user"] || ""),
      Status: String(rowMap["status"] || "IN STOCK").toUpperCase(),
      BillPct: num_(rowMap["billpct"]),
      Type: String(rowMap["type"] || rowMap["certtype"] || rowMap["certificatetype"] || "IGI")
    };
  });
}

function getAllSales() {
  let sheet = getSheetByName("Sales");
  let data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  let rawHeaders = data[0];
  let saleIdColIdx = -1;

  for (let c = 0; c < rawHeaders.length; c++) {
    let nk = normKey(rawHeaders[c]);
    if (nk === "saleid") saleIdColIdx = c;
  }

  // Auto-repair: Insert SaleId column at col 1 if missing
  if (saleIdColIdx === -1) {
    sheet.insertColumnBefore(1);
    sheet.getRange(1, 1).setValue("SaleId");
    data = sheet.getDataRange().getValues();
    rawHeaders = data[0];
    saleIdColIdx = 0;
  }

  // Auto-repair: Backfill SaleId for any empty cells (never reuses an existing number)
  let maxS = 2000;
  for (let i = 1; i < data.length; i++) {
    const m = String(data[i][saleIdColIdx]).trim().match(/^S-(\d+)$/);
    if (m) maxS = Math.max(maxS, parseInt(m[1], 10));
  }
  for (let i = 1; i < data.length; i++) {
    let currentId = String(data[i][saleIdColIdx] || "").trim();
    if (!currentId) {
      maxS++;
      let generatedId = "S-" + maxS;
      sheet.getRange(i + 1, saleIdColIdx + 1).setValue(generatedId);
      data[i][saleIdColIdx] = generatedId;
    }
  }

  const pctColIdx = getProfitPctColIdx_(rawHeaders);

  return data.slice(1).map((row, index) => {
    let rowMap = {};
    rawHeaders.forEach((h, i) => {
      rowMap[normKey(h)] = row[i];
    });

    let saleId = rowMap["saleid"] ? String(rowMap["saleid"]).trim() : "S-" + (2001 + index);

    return {
      _rowIndex: index + 2,
      SaleId: saleId,
      PurId: String(rowMap["purid"] || ""),
      Date: formatDateSafe(rowMap["date"]),
      ClientName: String(rowMap["clientname"] || ""),
      Country: String(rowMap["country"] || ""),
      SalePerCarat: rowMap["salepercarat"] || rowMap["itemrate"] || "",
      SaleLess: rowMap["saleless"] !== undefined && rowMap["saleless"] !== "" ? rowMap["saleless"] : (rowMap["less"] || 0),
      SaleDFinal: rowMap["saledfinal"] || rowMap["final"] || "",
      AdvDAmt: rowMap["advdamt"] !== undefined && rowMap["advdamt"] !== "" ? rowMap["advdamt"] : (rowMap["advamt"] || 0),
      DueAmt: rowMap["dueamt"] !== undefined && rowMap["dueamt"] !== "" ? rowMap["dueamt"] : 0,
      SaleRemark: String(rowMap["saleremark"] || rowMap["remark"] || ""),
      User: String(rowMap["user"] || ""),
      SaleDProfit: rowMap["saledprofit"] !== undefined ? rowMap["saledprofit"] : "",
      ProfitPct: pctColIdx !== -1 ? row[pctColIdx] : ""
    };
  });
}

/**
 * CREATE & UPDATE: Purchase Record
 */
function savePurchaseRecord(formData) {
  let lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    let sheet = getSheetByName("Purchases");
    ensureBillPctColumn_(sheet);
    ensureTypeColumn_(sheet);
    let data = sheet.getDataRange().getValues();
    let rawHeaders = data[0];
    
    let purIdColIdx = -1;
    let giaColIdx = -1;
    for (let c = 0; c < rawHeaders.length; c++) {
      let nk = normKey(rawHeaders[c]);
      if (nk === "purid") purIdColIdx = c;
      if (nk === "giano" || nk === "gianumber") giaColIdx = c;
    }
    if (purIdColIdx === -1) purIdColIdx = 0;

    let rowIndex = -1;
    let id = formData.PurId ? String(formData.PurId).trim() : "";

    // 1. Match by PurId if present
    if (id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][purIdColIdx]).trim() === id) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    // 2. Match by _rowIndex if passed
    if (rowIndex === -1 && formData._rowIndex) {
      let rIdx = parseInt(formData._rowIndex, 10);
      if (rIdx > 1 && rIdx <= data.length) {
        rowIndex = rIdx;
        if (!id && purIdColIdx !== -1) {
          id = String(data[rowIndex - 1][purIdColIdx]).trim();
          formData.PurId = id;
        }
      }
    }

    // 3. Block duplicate GIA / Lot numbers (delete this block if you reuse lot numbers)
    if (giaColIdx !== -1 && formData.GIANo) {
      const targetGia = String(formData.GIANo).trim();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][giaColIdx]).trim() === targetGia && (i + 1) !== rowIndex) {
          throw new Error("GIA / Lot No " + targetGia + " already exists (" + data[i][purIdColIdx] + ").");
        }
      }
    }

    // If new record, generate new ID
    if (rowIndex === -1) {
      id = getNextId(sheet, "P", 1001);
      formData.PurId = id;
      formData.Status = "IN STOCK";
    }

    if (!formData.Status) formData.Status = "IN STOCK";

    // Recalculate amounts on the server (do not trust the browser)
    // Bill charge: % added on top of the total, only when payment type is BILL
    const isBill = String(formData.BType || "").toUpperCase() === "BILL";
    formData.BillPct = isBill ? num_(formData.BillPct) : 0;
    const baseCost = num_(formData.Size) * num_(formData.PerCarat) * (1 + num_(formData.Less) / 100);
    const pCost = baseCost * (1 + formData.BillPct / 100);
    formData.DAmount = round3_(pCost);
    formData.IndAmount = round3_(pCost * num_(formData.Rate));
    formData.Type = formData.Type ? String(formData.Type).trim() : "IGI";

    // Build form lookup by normalized key
    let formNormMap = {};
    Object.keys(formData).forEach(k => {
      formNormMap[normKey(k)] = formData[k];
    });

    // Alias common variants
    if (formData.PurchaseDate) formNormMap["date"] = formData.PurchaseDate;
    if (formData.GIANo) formNormMap["gianumber"] = formData.GIANo;
    if (formData.FromName) formNormMap["fromcomp"] = formData.FromName;
    if (formData.Size) formNormMap["sizecarat"] = formData.Size;
    if (formData.Flor) formNormMap["fluor"] = formData.Flor;
    if (formData.Measur) formNormMap["measurements"] = formData.Measur;
    if (formData.BType) formNormMap["bill"] = formData.BType;
    if (formData.DAmount) formNormMap["amount"] = formData.DAmount;
    if (formData.PurId) formNormMap["purid"] = formData.PurId;
    if (formData.Type) {
      formNormMap["type"] = formData.Type;
      formNormMap["certtype"] = formData.Type;
      formNormMap["certificatetype"] = formData.Type;
    }

    let rowValues = rawHeaders.map(header => {
      let val = formNormMap[normKey(header)];
      return val !== undefined && val !== null ? val : "";
    });

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
      recalcSalesForPurchase_(id, num_(formData.DAmount));
    } else {
      sheet.appendRow(rowValues);
    }

    return {
      success: true,
      purId: id,
      purchases: getAllPurchases()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * DELETE: Purchase Record (Blocked if diamond has already been sold)
 */
function deletePurchaseRecord(purId, rowIndexParam) {
  let lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    purId = String(purId || "").trim();
    let salesSheet = getSheetByName("Sales");
    let salesData = salesSheet.getDataRange().getValues();
    let sHeaders = salesData[0] || [];
    let sPurIdCol = -1;
    for (let c = 0; c < sHeaders.length; c++) {
      if (normKey(sHeaders[c]) === "purid") {
        sPurIdCol = c;
        break;
      }
    }
    if (sPurIdCol === -1) sPurIdCol = 1;
    
    // Check if PurId exists in active sales
    if (purId) {
      for (let i = 1; i < salesData.length; i++) {
        if (String(salesData[i][sPurIdCol]).trim() === purId) {
          throw new Error("Cannot delete: This diamond is already recorded in Sale #" + salesData[i][0]);
        }
      }
    }

    let pSheet = getSheetByName("Purchases");
    let pData = pSheet.getDataRange().getValues();
    let pHeaders = pData[0] || [];
    let pPurIdCol = -1;
    for (let c = 0; c < pHeaders.length; c++) {
      if (normKey(pHeaders[c]) === "purid") {
        pPurIdCol = c;
        break;
      }
    }
    if (pPurIdCol === -1) pPurIdCol = 0;

    let rowIndex = -1;
    if (purId) {
      for (let i = 1; i < pData.length; i++) {
        if (String(pData[i][pPurIdCol]).trim() === purId) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex === -1 && rowIndexParam) {
      let rIdx = parseInt(rowIndexParam, 10);
      if (rIdx > 1 && rIdx <= pData.length) {
        rowIndex = rIdx;
      }
    }

    if (rowIndex > 0) {
      pSheet.deleteRow(rowIndex);
    } else {
      throw new Error("Purchase record " + (purId || "") + " was not found in the sheet.");
    }

    return {
      success: true,
      purchases: getAllPurchases()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * CREATE & UPDATE: Sale Record
 * Automatically updates linked Purchase Status to 'SOLD'
 */
function saveSaleRecord(formData) {
  let lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    let sSheet = getSheetByName("Sales");
    let sData = sSheet.getDataRange().getValues();
    let rawHeaders = sData[0];
    
    let saleIdColIdx = -1;
    let salePurIdColIdx = -1;
    for (let c = 0; c < rawHeaders.length; c++) {
      const nk = normKey(rawHeaders[c]);
      if (nk === "saleid") saleIdColIdx = c;
      if (nk === "purid") salePurIdColIdx = c;
    }
    if (saleIdColIdx === -1) saleIdColIdx = 0;
    if (salePurIdColIdx === -1) salePurIdColIdx = 1;

    let rowIndex = -1;
    let saleId = formData.SaleId ? String(formData.SaleId).trim() : "";

    if (saleId) {
      for (let i = 1; i < sData.length; i++) {
        if (String(sData[i][saleIdColIdx]).trim() === saleId) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    if (rowIndex === -1 && formData._rowIndex) {
      let rIdx = parseInt(formData._rowIndex, 10);
      if (rIdx > 1 && rIdx <= sData.length) {
        rowIndex = rIdx;
      }
    }

    if (rowIndex === -1) {
      saleId = getNextId(sSheet, "S", 2001);
      formData.SaleId = saleId;
    }

    if (!formData.Date) {
      formData.Date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    }

    // ---- Client name and country are mandatory ----
    formData.ClientName = String(formData.ClientName || "").trim();
    formData.Country = String(formData.Country || "").trim();
    if (!formData.ClientName) throw new Error("Client name is required.");
    if (!formData.Country) throw new Error("Country is required.");

    // ---- Validate + recalculate on the server ----
    const purIdTrim = String(formData.PurId || "").trim();
    const pur = getAllPurchases().find(p => String(p.PurId).trim() === purIdTrim);
    if (!pur) throw new Error("Linked purchase " + purIdTrim + " was not found.");

    for (let i = 1; i < sData.length; i++) {
      if (String(sData[i][salePurIdColIdx]).trim() === purIdTrim && (i + 1) !== rowIndex) {
        throw new Error("This diamond is already sold in Sale " + sData[i][saleIdColIdx] + ".");
      }
    }

    formData.PurId = purIdTrim;
    const sFinal  = num_(pur.Size) * num_(formData.SalePerCarat) * (1 - num_(formData.SaleLess) / 100);
    const sProfit = sFinal - num_(pur.DAmount);
    formData.SaleDFinal  = round2_(sFinal);
    formData.SaleDProfit = round2_(sProfit);
    formData.ProfitPct   = sFinal > 0 ? round2_(sProfit / sFinal * 100) : 0;
    formData.DueAmt      = round2_(sFinal - num_(formData.AdvDAmt));

    // Build form lookup by normalized key
    let formNormMap = {};
    Object.keys(formData).forEach(k => {
      formNormMap[normKey(k)] = formData[k];
    });

    // Alias common variants
    if (formData.SalePerCarat) formNormMap["itemrate"] = formData.SalePerCarat;
    if (formData.SaleLess) formNormMap["less"] = formData.SaleLess;
    if (formData.SaleDFinal) formNormMap["final"] = formData.SaleDFinal;
    if (formData.AdvDAmt) formNormMap["advamt"] = formData.AdvDAmt;
    if (formData.SaleRemark) formNormMap["remark"] = formData.SaleRemark;

    const pctIdx = getProfitPctColIdx_(rawHeaders);
    let rowValues = rawHeaders.map((header, c) => {
      if (c === pctIdx) return formData.ProfitPct !== undefined ? formData.ProfitPct : "";
      let val = formNormMap[normKey(header)];
      return val !== undefined && val !== null ? val : "";
    });

    if (rowIndex > 0) {
      sSheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sSheet.appendRow(rowValues);
    }

    // Mark linked Purchase as SOLD
    if (formData.PurId) {
      updatePurchaseStatus(formData.PurId, "SOLD");
    }

    return {
      success: true,
      saleId: saleId,
      sales: getAllSales(),
      purchases: getAllPurchases()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * DELETE: Sale Record
 * Automatically resets the linked Purchase Status back to 'IN STOCK'
 */
function deleteSaleRecord(saleId, rowIndexParam) {
  let lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    saleId = String(saleId || "").trim();
    let sSheet = getSheetByName("Sales");
    let sData = sSheet.getDataRange().getValues();
    let rawHeaders = sData[0] || [];
    let saleIdCol = -1;
    let sPurIdCol = -1;

    for (let c = 0; c < rawHeaders.length; c++) {
      let nk = normKey(rawHeaders[c]);
      if (nk === "saleid") saleIdCol = c;
      if (nk === "purid") sPurIdCol = c;
    }
    if (saleIdCol === -1) saleIdCol = 0;
    if (sPurIdCol === -1) sPurIdCol = 1;

    let rowIndex = -1;
    let linkedPurId = "";

    if (saleId) {
      for (let i = 1; i < sData.length; i++) {
        if (String(sData[i][saleIdCol]).trim() === saleId) {
          rowIndex = i + 1;
          linkedPurId = String(sData[i][sPurIdCol]).trim();
          break;
        }
      }
    }

    if (rowIndex === -1 && rowIndexParam) {
      let rIdx = parseInt(rowIndexParam, 10);
      if (rIdx > 1 && rIdx <= sData.length) {
        rowIndex = rIdx;
        linkedPurId = String(sData[rowIndex - 1][sPurIdCol]).trim();
      }
    }

    if (rowIndex > 0) {
      sSheet.deleteRow(rowIndex);
      // Reset linked purchase back to IN STOCK
      if (linkedPurId) {
        updatePurchaseStatus(linkedPurId, "IN STOCK");
      }
    } else {
      throw new Error("Sale record " + (saleId || "") + " was not found.");
    }

    return {
      success: true,
      sales: getAllSales(),
      purchases: getAllPurchases()
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Helper to update status in Purchases sheet
 */
function updatePurchaseStatus(purId, newStatus) {
  let pSheet = getSheetByName("Purchases");
  let pData = pSheet.getDataRange().getValues();
  let headers = pData[0];
  let statusColIdx = -1;
  let purIdColIdx = -1;
  let giaColIdx = -1;

  for (let c = 0; c < headers.length; c++) {
    let nk = normKey(headers[c]);
    if (nk === "status") statusColIdx = c;
    if (nk === "purid") purIdColIdx = c;
    if (nk === "giano" || nk === "gianumber") giaColIdx = c;
  }
  if (statusColIdx === -1) {
    pSheet.insertColumnAfter(headers.length);
    pSheet.getRange(1, headers.length + 1).setValue("Status");
    statusColIdx = headers.length;
  }
  if (purIdColIdx === -1) purIdColIdx = 0;

  let targetId = String(purId).trim();
  for (let i = 1; i < pData.length; i++) {
    let pId = String(pData[i][purIdColIdx]).trim();
    let gia = giaColIdx !== -1 ? String(pData[i][giaColIdx]).trim() : "";
    if (pId === targetId || gia === targetId) {
      pSheet.getRange(i + 1, statusColIdx + 1).setValue(newStatus);
      break;
    }
  }
}

/**
 * Refreshes profit on linked sales when a purchase's cost is edited
 */
function recalcSalesForPurchase_(purId, newCost) {
  const sheet = getSheetByName("Sales");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const h = data[0];
  const iPur    = h.findIndex(x => normKey(x) === "purid");
  const iFinal  = h.findIndex(x => normKey(x) === "saledfinal");
  const iProfit = h.findIndex(x => normKey(x) === "saledprofit");
  const iPct    = getProfitPctColIdx_(h);
  if (iPur < 0 || iFinal < 0 || iProfit < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][iPur]).trim() !== String(purId).trim()) continue;
    const fin = num_(data[i][iFinal]);
    const profit = round2_(fin - newCost);
    sheet.getRange(i + 1, iProfit + 1).setValue(profit);
    if (iPct >= 0) sheet.getRange(i + 1, iPct + 1).setValue(fin > 0 ? round2_(profit / fin * 100) : 0);
  }
}

/**
 * REPORT QUERY Engine: Relational Join of Sales and Purchases
 * Outputs exactly the 31 consolidated business columns matching export file
 */
function getSalesReportByDate(fromDate, toDate) {
  let purchases = getAllPurchases();
  let sales = getAllSales();

  let fromStr = fromDate ? String(fromDate).trim() : "1900-01-01";
  let toStr = toDate ? String(toDate).trim() : "2099-12-31";

  // Build Purchase HashMap for fast O(1) lookups
  let purchaseMap = {};
  purchases.forEach(p => {
    purchaseMap[p.PurId] = p;
  });

  return sales.filter(s => {
    let sDate = s.Date || "";
    return sDate >= fromStr && sDate <= toStr;
  }).map(s => {
    let p = purchaseMap[s.PurId] || {};
    return {
      // Purchase Fields
      GIANo: p.GIANo || "",
      PurchaseDate: p.PurchaseDate || "",
      FromName: p.FromName || "",
      Size: p.Size || "",
      Shape: p.Shape || "",
      Color: p.Color || "",
      Quality: p.Quality || "",
      Flor: p.Flor || "",
      Cut: p.Cut || "",
      Pol: p.Pol || "",
      Sym: p.Sym || "",
      Measur: p.Measur || "",
      Less: p.Less !== undefined ? p.Less : "",
      PerCarat: p.PerCarat || "",
      DAmount: p.DAmount || "",
      Rate: p.Rate || "",
      IndAmount: p.IndAmount || "",
      BType: p.BType || "",
      BillPct: p.BillPct !== undefined ? p.BillPct : "",
      PaymentDay: p.PaymentDay || "",
      Broker: p.Broker || "",
      PurUser: p.User || "",
      PurId: p.PurId || s.PurId || "",
      Type: p.Type || "IGI",
      PurchaseMissing: !purchaseMap[s.PurId],
      // Sale Fields
      SaleId: s.SaleId || "",
      SaleDate: s.Date || "",
      ClientName: s.ClientName || "",
      Country: s.Country || "",
      SalePerCarat: s.SalePerCarat || "",
      SaleLess: s.SaleLess !== undefined ? s.SaleLess : "",
      SaleDFinal: s.SaleDFinal || "",
      AdvDAmt: s.AdvDAmt !== undefined ? s.AdvDAmt : "",
      DueAmt: s.DueAmt !== undefined ? s.DueAmt : "",
      SaleRemark: s.SaleRemark || "",
      SaleUser: s.User || "",
      SaleDProfit: s.SaleDProfit !== undefined ? s.SaleDProfit : "",
      ProfitPct: s.ProfitPct !== undefined ? s.ProfitPct : ""
    };
  });
}

/**
 * Optional Utility: Pre-populate Google Sheet with the 26 sample transactions
 * provided in your specifications for immediate testing.
 */
function seedInitialSampleData() {
  let sampleData = [
    ["793647470","2026-07-01","KD",3.01,"OVAL","E","VVS2","NONE","EX","EX","EX","12.11*8.03*4.76",0,100.3,301.903,95.24,28753.242,"CASH",21,"Keyur Zalavadiya","2026-07-01","Hala",121.27,0,365.02,0,365.02,"","Keyur Zalavadiya",63.12,17.29],
     ["811660767","2026-07-14","RK BRILLIANCE PRIVATE LIMITED",0.51,"PEAR","D","VVS2","NONE","EX","EX","EX","7.03*4.45*2.79",1.5,97,50.212,95.63,4801.774,"BILL",7,"Keyur Zalavadiya","2026-07-14","Diana Lady-D-World",137.26,0,70,0,70,"","Keyur Zalavadiya",19.79,28.27],
    ["811683771","2026-07-14","RK BRILLIANCE PRIVATE LIMITED",0.5,"PEAR","D","VVS2","NONE","EX","EX","EX","7.02*4.42*2.74",1.5,97,49.228,95.63,4707.674,"BILL",7,"Keyur Zalavadiya","2026-07-14","Diana Lady-D-World",140,0,70,0,70,"","Keyur Zalavadiya",20.77,29.67]
  ];

  let pSheet = getSheetByName("Purchases");
  let sSheet = getSheetByName("Sales");

  pSheet.clear();
  sSheet.clear();
  pSheet.appendRow(PURCHASE_HEADERS);
  sSheet.appendRow(SALE_HEADERS);
  pSheet.setFrozenRows(1);
  sSheet.setFrozenRows(1);

  sampleData.forEach((row, idx) => {
    let purId = "P-" + (1001 + idx);
    let saleId = "S-" + (2001 + idx);

    let pRow = [
      purId, row[0], row[1], row[2], row[3], row[4], row[5], row[6],
      row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14],
      row[15], row[16], row[17], row[18], "", row[19], "SOLD", 0, "IGI"
    ];
    pSheet.appendRow(pRow);

    let sRow = [
      saleId, purId, row[20], row[21], "", row[22], row[23],
      row[24], row[25], row[26], row[27], row[28], row[29], row[30]
    ];
    sSheet.appendRow(sRow);
  });

  Logger.log("Seed complete: " + sampleData.length + " transactions populated.");
}


/**
 * ONE-TIME REPAIR: recalculates the "Profit %" column for all existing sales.
 * Run once from the Apps Script editor (select repairProfitPercent -> Run).
 */
function repairProfitPercent() {
  const sheet = getSheetByName("Sales");
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  const h = data[0];
  const iFinal  = h.findIndex(x => normKey(x) === "saledfinal");
  const iProfit = h.findIndex(x => normKey(x) === "saledprofit");
  const iPct    = getProfitPctColIdx_(h);
  if (iFinal < 0 || iProfit < 0 || iPct < 0) throw new Error("Required columns not found.");

  const out = [];
  for (let i = 1; i < data.length; i++) {
    const fin = num_(data[i][iFinal]);
    const prof = num_(data[i][iProfit]);
    out.push([fin > 0 ? round2_(prof / fin * 100) : 0]);
  }
  sheet.getRange(2, iPct + 1, out.length, 1).setValues(out);
  Logger.log("Repaired " + out.length + " rows");
}
