var DEBUG = true;

function log() {
  if (DEBUG) Logger.log.apply(Logger, arguments);
}

// ================================================
//  code.gs — Procurement System (Fixed)
// ================================================

function showForm() {
  const html = HtmlService
    .createHtmlOutputFromFile('form')
    .setWidth(2000)
    .setHeight(2300);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/* =================================
   GET SUMMARY SHEET
================================= */
function getSheet() {
  return SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("SUMMARY");
}

function cleanPrNo(str) {
  return String(str)
    .replace(/[\u200B-\u200F\u2028-\u202E\u2060-\u2069\uFEFF\u00A0\u00AD\u034F\u061C\u115F\u1160\u17B4\u17B5\u180E\u2000-\u200A\u202F\u205F\u206A-\u206F\u3000\u2800\u2D7F\uA670-\uA673\uD7CB-\uD7FB\uFFA0]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function parseNum(v) {
  if (typeof v === 'number' && !isNaN(v)) return v;
  return parseFloat(String(v).replace(/[^0-9.\-]/g, '')) || 0;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('My Menu')
    .addItem('Start Canvass', 'showCanvass')
    .addToUi();
}

function showCanvass() {
  var url = 'https://docs.google.com/spreadsheets/d/1JgayRq38k4H7JrQ3i_L10hSuA0Z82A0D/edit?gid=328039259#gid=328039259';
  
  var html = HtmlService.createHtmlOutput(
    '<script>window.open("' + url + '", "_blank"); google.script.host.close();</script>'
  );
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Canvass...');
}

/* =================================
   GET NEXT UNIQUE PR NO
   Generates the next unique PR No.
   based on the selected office.
   Pattern: PR-{Initials}-{YYYY}-{MM}-{NNN}
   Checks SUMMARY sheet to ensure the
   number is not already in use.
================================= */
function getNextUniquePRNo(office) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  var today = new Date();
  var yyyy = today.getFullYear();
  var mm = String(today.getMonth() + 1).padStart(2, '0');

  var initials = getOfficeInitials(office);
  var prefix = 'PR-' + initials + '-' + yyyy + '-' + mm + '-';

  // Collect existing PR Nos from sheet (col B)
  var existing = {};
  for (var i = 1; i < data.length; i++) {
    var pr = cleanPrNo(data[i][1]).toUpperCase();
    if (pr) existing[pr] = true;
  }

  // Find the next unique NNN
  var seq = 1;
  while (existing[prefix + String(seq).padStart(3, '0')]) {
    seq++;
  }
  return prefix + String(seq).padStart(3, '0');
}

/**
 * Generates the next unique PO No. based on the office
 * associated with the given PR No.
 * Pattern: PO-{OfficeInitials}-{YYYY}-{MM}-{NNN}
 * Checks existing PO sheet names and AOQ_SupplierData for uniqueness.
 */
function getNextUniquePONo(prNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var today = new Date();
  var yyyy = today.getFullYear();
  var mm = String(today.getMonth() + 1).padStart(2, '0');

  // Look up office from SUMMARY for this PR No.
  var office = '';
  if (summarySheet) {
    var summaryData = summarySheet.getDataRange().getValues();
    var nPrNo = cleanPrNo(prNo).toLowerCase();
    for (var i = 1; i < summaryData.length; i++) {
      if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo) {
        office = String(summaryData[i][3] || '').trim();
        if (office) break;
      }
    }
  }

  var initials = office ? getOfficeInitials(office) : 'RO';
  var prefix = 'PO-' + initials + '-' + yyyy + '-' + mm + '-';

  // Collect existing PO Nos from sheet names (PO - SupplierName) and AOQ_SupplierData
  var existing = {};
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (name.indexOf('PO - ') === 0) {
      existing[name] = true;
    }
  }

  // Also check AOQ_SupplierData column J for existing PO Nos
  var aoqDataSheet = ss.getSheetByName("AOQ_SupplierData");
  if (aoqDataSheet) {
    var aoqData = aoqDataSheet.getDataRange().getValues();
    for (var i = 1; i < aoqData.length; i++) {
      var poNo = String(aoqData[i][9] || '').trim().toUpperCase();
      if (poNo) existing[poNo] = true;
    }
  }

  // Find the next unique NNN
  var seq = 1;
  while (existing[prefix + String(seq).padStart(3, '0')]) {
    seq++;
  }
  return prefix + String(seq).padStart(3, '0');
}

/* =================================
   GET OFFICE INITIALS
   Extracts the first letter of each
   word in the office name to form
   initials for PR No. generation.
   e.g. "Regional Office" → "RO"
================================= */
function getOfficeInitials(office) {
  return office.split(' ').map(function(word) {
    return word.charAt(0).toUpperCase();
  }).join('');
}

/* =================================
   SAVE SINGLE ITEM
================================= */
function saveItem(data) {
  if (
    !data.date || !data.prNo || !data.office ||
    !data.itemDescription || !data.quantity || !data.unitCost
  ) {
    return "Please fill in all required fields!";
  }

  const sheet = getSheet();
  const nextRow = sheet.getLastRow() + 1;
  const totalCost = Number(data.quantity || 0) * Number(data.unitCost || 0);

  sheet.getRange(nextRow, 1, 1, 19).setValues([[
    data.date, cleanPrNo(data.prNo), data.fundCluster, data.office, data.unit,
    data.itemDescription, data.quantity, data.unitCost, totalCost,
    data.purpose, data.deliveryTerm, data.paymentTerm, data.deliveryPeriod,
    data.supplier, data.supplierAddress, data.tin,
    data.modeProcurement, data.placeDelivery, data.dateDelivery
  ]]);

  return "Item Added Successfully!";
}

/* =================================
   SAVE MULTIPLE ITEMS (batch)
================================= */
function saveMultipleItems(itemsArray) {
  if (!itemsArray || itemsArray.length === 0) {
    return "No items to save!";
  }

  for (var i = 0; i < itemsArray.length; i++) {
    var d = itemsArray[i];
    if (!d.date || !d.prNo || !d.office || !d.itemDescription ||
        !d.quantity || !d.unitCost) {
      return "Item #" + (i + 1) + " is missing required fields. Save cancelled.";
    }
  }

  var rows = itemsArray.map(function(d) {
    var totalCost = Number(d.quantity || 0) * Number(d.unitCost || 0);
    return [
      d.date, cleanPrNo(d.prNo), d.fundCluster, d.office, d.unit,
      d.itemDescription, d.quantity, d.unitCost, totalCost,
      d.purpose, d.deliveryTerm, d.paymentTerm, d.deliveryPeriod,
      d.supplier, d.supplierAddress, d.tin,
      d.modeProcurement, d.placeDelivery, d.dateDelivery
    ];
  });

  var sheet   = getSheet();
  var nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, rows.length, 19).setValues(rows);

  try {
    writeToPRSheet(cleanPrNo(itemsArray[0].prNo));
  } catch(e) {
    Logger.log("writeToPRSheet error: " + e);
  }

  return rows.length + " Item(s) Saved Successfully!";
}

/* =================================
   SEARCH ITEM (from SUMMARY sheet)
   Returns all rows matching prNo,
   regardless of form type.
================================= */
function searchItem(prNo) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  let results = [];
  const nInput = cleanPrNo(prNo).toLowerCase();

  console.log({msg: "searchItem called", prNo: prNo, nInput: nInput, totalRows: data.length});

  for (let i = 1; i < data.length; i++) {
    const stored = data[i][1];
    const nStored = cleanPrNo(stored).toLowerCase();
    const match = nStored === nInput;
    console.log({msg: "searchItem compare", row: i, stored: stored, nStored: nStored, nInput: nInput, match: match});
    if (match) {
      results.push({
        row: i + 1,
        date: formatDate(data[i][0]),
        prNo: cleanPrNo(data[i][1]),
        fundCluster: data[i][2],
        office: data[i][3],
        unit: data[i][4],
        itemDescription: data[i][5],
        quantity: data[i][6],
        unitCost: data[i][7],
        totalCost: data[i][8],
        purpose: data[i][9],
        deliveryTerm: data[i][10],
        paymentTerm: data[i][11],
        deliveryPeriod: formatDate(data[i][12]),
        supplier: data[i][13],
        supplierAddress: data[i][14],
        tin: data[i][15],
        modeProcurement: data[i][16],
        placeDelivery: data[i][17],
        dateDelivery: formatDate(data[i][18])
      });
    }
  }

  // When no results found, attach diagnostics to help debug
  if (results.length === 0) {
    var diagSamples = [];
    var maxSamples = Math.min(data.length - 1, 20);
    for (var d = 1; d <= maxSamples; d++) {
      var raw = data[d][1];
      if (raw !== undefined && raw !== null && raw !== '') {
        diagSamples.push({
          row: d + 1,
          raw: String(raw),
          rawChars: String(raw).split('').map(function(c){return c.charCodeAt(0);}),
          cleaned: cleanPrNo(raw)
        });
      }
    }
    results._diagnostics = {
      searchedInputRaw: prNo,
      inputChars: String(prNo).split('').map(function(c){return c.charCodeAt(0);}),
      cleanedInput: nInput,
      storedRowCount: data.length - 1,
      diagnosticSamples: diagSamples,
      sampleCount: diagSamples.length
    };
  }

  return results;
}

/* =================================
   FIX: SEARCH ITEM BY TYPE
   For AOQ / RFQ / PO form views —
   currently searches the same SUMMARY
   sheet. If you separate sheets per
   type in the future, update here.
================================= */
function searchItemByType(prNo, formType) {
  return searchItem(cleanPrNo(prNo));
}

/* =================================
   SEARCH SUPPLIERS BY PR No.
   Groups items by supplier name so
   the PO form shows suppliers instead
   of individual items.
================================= */
function searchSuppliersByPR(prNo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nInput = cleanPrNo(prNo).toLowerCase();

  // Build item lookup from SUMMARY to get full item details
  const summarySheet = getSheet();
  const summaryData = summarySheet.getDataRange().getValues();
  const itemLookup = {};
  for (let i = 1; i < summaryData.length; i++) {
    const nStored = cleanPrNo(summaryData[i][1]).toLowerCase();
    if (nStored !== nInput) continue;
    const desc = String(summaryData[i][5] || '').trim().toLowerCase();
    if (!desc) continue;
    itemLookup[desc] = {
      row: i + 1,
      date: formatDate(summaryData[i][0]),
      prNo: cleanPrNo(summaryData[i][1]),
      fundCluster: summaryData[i][2],
      office: summaryData[i][3],
      unit: summaryData[i][4],
      itemDescription: summaryData[i][5],
      quantity: Number(summaryData[i][6]) || 0,
      unitCost: Number(summaryData[i][7]) || 0,
      totalCost: Number(summaryData[i][8]) || 0,
      purpose: summaryData[i][9],
      deliveryTerm: summaryData[i][10],
      paymentTerm: summaryData[i][11],
      deliveryPeriod: formatDate(summaryData[i][12]),
      supplier: summaryData[i][13] || '',
      supplierAddress: summaryData[i][14] || '',
      tin: summaryData[i][15] || '',
      modeProcurement: summaryData[i][16] || '',
      placeDelivery: summaryData[i][17] || '',
      dateDelivery: formatDate(summaryData[i][18])
    };
  }

  // Read supplier selections from AOQ_SupplierData
  const aoqSheet = ss.getSheetByName("AOQ_SupplierData");
  if (!aoqSheet) return { prNo: cleanPrNo(prNo), suppliers: [] };

  const aoqData = aoqSheet.getDataRange().getValues();
  const grouped = {};

  console.log({msg: "searchSuppliersByPR reading AOQ_SupplierData + SUMMARY", prNo: prNo, nInput: nInput, aoqRows: aoqData.length, summaryRows: summaryData.length});

  for (let i = 1; i < aoqData.length; i++) {
    if (cleanPrNo(String(aoqData[i][0] || '')).toLowerCase() !== nInput) continue;
    if (String(aoqData[i][7] || '').toUpperCase() !== 'TRUE') continue;

    const supplierName = String(aoqData[i][1] || '').trim() || 'Unassigned';
    if (!grouped[supplierName]) {
      grouped[supplierName] = {
        supplier: supplierName,
        supplierAddress: String(aoqData[i][2] || '').trim(),
        tin: String(aoqData[i][3] || '').trim(),
        modeProcurement: '',
        placeDelivery: '',
        dateDelivery: '',
        items: [],
        itemCount: 0,
        totalAmount: 0
      };
    }

    const group = grouped[supplierName];
    const aoqUnitCost = Number(aoqData[i][5]) || 0;
    const aoqTotalCost = Number(aoqData[i][6]) || 0;
    const aoqDesc = String(aoqData[i][4] || '').trim();
    const descKey = aoqDesc.toLowerCase();

    // Use AOQ unit cost if available, otherwise fall back to SUMMARY
    const summaryItem = itemLookup[descKey] || {};
    const unitCost = aoqUnitCost || summaryItem.unitCost || 0;
    const totalCost = aoqTotalCost || summaryItem.totalCost || 0;

    group.items.push({
      row: summaryItem.row || (i + 1),
      date: summaryItem.date || '',
      prNo: cleanPrNo(prNo),
      fundCluster: summaryItem.fundCluster || '',
      office: summaryItem.office || '',
      unit: summaryItem.unit || '',
      itemDescription: summaryItem.itemDescription || aoqDesc,
      quantity: summaryItem.quantity || 0,
      unitCost: unitCost,
      totalCost: totalCost,
      purpose: summaryItem.purpose || '',
      deliveryTerm: summaryItem.deliveryTerm || '',
      paymentTerm: summaryItem.paymentTerm || '',
      deliveryPeriod: summaryItem.deliveryPeriod || '',
      supplier: supplierName,
      supplierAddress: String(aoqData[i][2] || '').trim(),
      tin: String(aoqData[i][3] || '').trim(),
      modeProcurement: summaryItem.modeProcurement || '',
      placeDelivery: summaryItem.placeDelivery || '',
      dateDelivery: summaryItem.dateDelivery || ''
    });

    group.itemCount++;
    group.totalAmount += totalCost;
  }

  const suppliers = Object.keys(grouped).map(function(key) { return grouped[key]; });

  console.log({msg: "searchSuppliersByPR result", prNo: prNo, supplierCount: suppliers.length, supplierNames: suppliers.map(function(s){return s.supplier;})});
  return { prNo: cleanPrNo(prNo), suppliers: suppliers };
}

/* =================================
   BATCH UPDATE PO FIELDS FOR SUPPLIER
   Syncs PO-level fields across all
   items of the same supplier+PR combo.
================================= */
function updatePOFieldsForSupplier(prNo, supplier, poFields) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();
  const nInput = cleanPrNo(prNo).toLowerCase();
  let updatedCount = 0;

  for (let i = 1; i < data.length; i++) {
    const nStored = cleanPrNo(data[i][1]).toLowerCase();
    const storedSupplier = (data[i][13] || '').trim();

    if (nStored === nInput && storedSupplier === supplier) {
      const row = i + 1;
      const current = sheet.getRange(row, 1, 1, 19).getValues()[0];

      current[13] = poFields.supplier;
      current[14] = poFields.supplierAddress;
      current[15] = poFields.tin;
      current[16] = poFields.modeProcurement;
      current[17] = poFields.placeDelivery;
      current[18] = poFields.dateDelivery;

      sheet.getRange(row, 1, 1, 19).setValues([current]);
      updatedCount++;
    }
  }

  // Also update AOQ_SupplierData to keep in sync
  const aoqSheet = ss.getSheetByName("AOQ_SupplierData");
  if (aoqSheet) {
    const aoqData = aoqSheet.getDataRange().getValues();
    for (let i = 1; i < aoqData.length; i++) {
      const nPrNo = cleanPrNo(String(aoqData[i][0] || '')).toLowerCase();
      const storedSup = String(aoqData[i][1] || '').trim();
      if (nPrNo === nInput && storedSup === supplier) {
        aoqSheet.getRange(i + 1, 2).setValue(poFields.supplier);
        aoqSheet.getRange(i + 1, 3).setValue(poFields.supplierAddress);
        aoqSheet.getRange(i + 1, 4).setValue(poFields.tin);
      }
    }
  }

  SpreadsheetApp.flush();
  return "PO fields synced for " + updatedCount + " item(s) under supplier: " + supplier;
}

/* =================================
   UPDATE ITEM
================================= */
function updateItem(data) {
  const sheet = getSheet();

  if (
    !data.date ||
    !data.office ||
    !data.itemDescription ||
    !data.quantity ||
    !data.unitCost ||
    !data.purpose
  ) {
    return "Please fill in all required fields!";
  }

  const targetRow = Number(data.row);

  if (!targetRow || targetRow < 2) {
    return "Invalid row number.";
  }

  const existing = sheet
    .getRange(targetRow, 1, 1, 19)
    .getValues()[0];

  if (!existing || !existing[1]) {
    return "Selected row no longer exists.";
  }

  // Preserve original PR No.
  const originalPrNo = cleanPrNo(existing[1]);

  const totalCost =
    Number(data.quantity || 0) *
    Number(data.unitCost || 0);

  const updatedRow = [
    data.date,                     // A Date
    originalPrNo,                  // B PR No. (LOCKED)
    data.fundCluster,              // C Fund Cluster
    data.office,                   // D Office
    data.unit,                     // E Unit
    data.itemDescription,          // F Item Description
    Number(data.quantity),         // G Quantity
    Number(data.unitCost),         // H Unit Cost
    totalCost,                     // I Total Cost
    data.purpose,                  // J Purpose
    data.deliveryTerm,             // K Delivery Term
    data.paymentTerm,              // L Payment Term
    data.deliveryPeriod,           // M Delivery Period
    data.supplier,                 // N Supplier
    data.supplierAddress,          // O Supplier Address
    data.tin,                      // P TIN
    data.modeProcurement,          // Q Mode of Procurement
    data.placeDelivery,            // R Place of Delivery
    data.dateDelivery              // S Date of Delivery
  ];

  Logger.log("===== BEFORE UPDATE =====");
  Logger.log({
    row: targetRow,
    originalPrNo: originalPrNo
  });

  sheet
    .getRange(targetRow, 1, 1, 19)
    .setValues([updatedRow]);

  SpreadsheetApp.flush();

  const verifyRow = sheet
    .getRange(targetRow, 1, 1, 19)
    .getValues()[0];

  Logger.log("===== AFTER UPDATE =====");
  Logger.log(JSON.stringify(verifyRow));

  Logger.log("PR NO SAVED = [" + verifyRow[1] + "]");

  return "Item Updated Successfully!";
}


/* =================================
   DELETE ITEM
   FIX: Physically removes the row
   from the SUMMARY sheet by row
   number. Row numbers in the form
   shift after deletion — the front-
   end reloads fresh data after delete
   to stay in sync.
================================= */
function deleteItem(row) {
  var sheet = getSheet();
  row = Number(row);

  if (!row || row < 2) return "Invalid row — cannot delete header or missing row.";

  var lastRow = sheet.getLastRow();
  if (row > lastRow) return "Row " + row + " does not exist in the sheet.";

  sheet.deleteRow(row);
  return "Item Deleted Successfully!";
}

/* =================================
   ACTIVATE SHEET & GET URL
================================= */
function activateSheetAndGetUrl(sheetName, prNo) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      var allSheets = ss.getSheets();
      for (var i = 0; i < allSheets.length; i++) {
        if (allSheets[i].getName().toLowerCase() === sheetName.toLowerCase()) {
          sheet = allSheets[i];
          break;
        }
      }
    }

    if (!sheet) {
      return { url: null, error: "Sheet not found: " + sheetName };
    }

    sheet.activate();

    if (prNo) {
      var data = sheet.getDataRange().getValues();
      var nPrNo = cleanPrNo(prNo).toLowerCase();
      for (var r = 1; r < data.length; r++) {
        if (cleanPrNo(data[r][1]).toLowerCase() === nPrNo) {
          sheet.setActiveRange(sheet.getRange(r + 1, 1));
          break;
        }
      }
    }

    var ssUrl = ss.getUrl();
    var gid   = sheet.getSheetId();
    return { url: ssUrl + '#gid=' + gid };

  } catch (e) {
    return { url: null, error: e.toString() };
  }
}

/* =================================
   FORMAT DATE
================================= */
function formatDate(date) {
  if (!date) return "";
  return Utilities.formatDate(
    new Date(date),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

/* =================================
   GET PR HEADER (first matching row)
================================= */
function getPRHeader(prNo) {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName("SUMMARY");

  const data = sheet.getDataRange().getValues();
  const nPrNo = cleanPrNo(prNo).toLowerCase();

  for (let i = 1; i < data.length; i++) {
    if (cleanPrNo(data[i][1]).toLowerCase() === nPrNo) {
      return {
        date:            formatDate(data[i][0]),
        prNo:            cleanPrNo(data[i][1]),
        fundCluster:     data[i][2],
        office:          data[i][3],
        unit:            data[i][4],
        itemDescription: data[i][5],
        quantity:        data[i][6],
        unitCost:        data[i][7],
        totalCost:       data[i][8],
        purpose:         data[i][9],
        deliveryTerm:    data[i][10],
        paymentTerm:     data[i][11],
        deliveryPeriod:  formatDate(data[i][12]),
        supplier:        data[i][13],
        supplierAddress: data[i][14],
        tin:             data[i][15],
        modeProcurement: data[i][16],
        placeDelivery:   data[i][17],
        dateDelivery:    formatDate(data[i][18])
      };
    }
  }

  return null;
}

/* =================================
   WRITE TO RFQ SHEET
   Populates the "RFQ" sheet from
   the SUMMARY data for the given PR No.
================================= */
function writeToRFQSheet(prNo) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var rfqSheet     = ss.getSheetByName("RFQ");

  if (!summarySheet) return "SUMMARY sheet not found.";
  if (!rfqSheet)     return "RFQ sheet not found.";

  var summaryData = summarySheet.getDataRange().getValues();
  var items = [];
  var nPrNo = cleanPrNo(prNo).toLowerCase();

  for (var i = 1; i < summaryData.length; i++) {
    if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo) {
      var rawDate  = summaryData[i][0];
      var dateObj  = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      var formatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");

      items.push({
        date:            formatted,
        prNo:            cleanPrNo(summaryData[i][1]),
        fundCluster:     summaryData[i][2],
        office:          summaryData[i][3],
        unit:            summaryData[i][4],
        itemDescription: summaryData[i][5],
        quantity:        parseNum(summaryData[i][6]),
        unitCost:        parseNum(summaryData[i][7]),
        totalCost:       parseNum(summaryData[i][8]),
        purpose:         summaryData[i][9],
        deliveryTerm:    summaryData[i][10],
        paymentTerm:     summaryData[i][11],
        deliveryPeriod:  summaryData[i][12],
        supplier:        summaryData[i][13],
        supplierAddress: summaryData[i][14],
        tin:             summaryData[i][15]
      });
    }
  }

  if (items.length === 0) return "No items found for PR No.: " + prNo;

  var DATA_START_ROW = 17;
  var MAX_ITEM_ROWS  = 26;

  // Clear previous data
  rfqSheet.getRange(DATA_START_ROW, 1, MAX_ITEM_ROWS, 5).clearContent();
  rfqSheet.getRange("A9").clearContent();
  rfqSheet.getRange("A10").clearContent();
  rfqSheet.getRange("F8").clearContent();
  rfqSheet.getRange("F9").clearContent();
  rfqSheet.getRange("A13").clearContent();
  rfqSheet.getRange("F15").clearContent();
  rfqSheet.getRange("C45").clearContent();
  rfqSheet.getRange("C46").clearContent();

  var first = items[0];

  // RFQ Date Selection at F8 (merged F:G) — from the RFQ form's date picker
  var rfqDate = first.deliveryPeriod ? Utilities.formatDate(new Date(first.deliveryPeriod), Session.getScriptTimeZone(), "MM/dd/yyyy") : "";
  rfqSheet.getRange("F8").setValue(rfqDate);

  // Date label at F9 (merged F:G)
  rfqSheet.getRange("F9").setValue("Date");

  // Supplier label at A9 (merged A:C)
  rfqSheet.getRange("A9").setValue("Supplier: ");

  // Supplier Address label at A10 (merged A:C)
  rfqSheet.getRange("A10").setValue("Supplier Address: ");

  // QUOTATION/OFFER label at F15 (merged F:G and 15:16)
  rfqSheet.getRange("F15").setValue("QUOTATION/OFFER");

  // Concatenated intro text + due date (date from F8 + 3 days) at A13 (merged A:G and 13:14)
  var f8Value = rfqSheet.getRange("F8").getValue();
  var f8Date = (f8Value instanceof Date) ? f8Value : new Date(f8Value);
  f8Date.setDate(f8Date.getDate() + 3);
  var dueDate = Utilities.formatDate(f8Date, Session.getScriptTimeZone(), "MM/dd/yyyy");
  var introText = "This Office is undergoing procurement of labor and materials for the repainting of the interior walls.  As a known supplier that typically provides the required goods and services, we are sending this Request of Quotation.  Please quote your possible lowest price by filling in the blanks below and return to this Office duly accomplished on or before";
  rfqSheet.getRange("A13").setValue(introText + " " + dueDate);

  // Write items starting at row 17, cols A-E
  // A: Quantity, B: Unit, C: Description, D: Unit Cost, E: Total Cost
  var rows = items.map(function(item) {
    return [
      item.quantity,        // A: Quantity
      item.unit,            // B: Unit
      item.itemDescription, // C: Item Description
      item.unitCost,        // D: Unit Cost
      item.totalCost        // E: Total Cost
    ];
  });

  rfqSheet.getRange(DATA_START_ROW, 1, rows.length, 5).setValues(rows);

  var lastDataRow = DATA_START_ROW + items.length - 1;

  // Grand total formula at row 43
  rfqSheet.getRange("E43").setFormula("=SUM(E" + DATA_START_ROW + ":E" + lastDataRow + ")");

  // Payment Term at C45
  if (first.paymentTerm) rfqSheet.getRange("C45").setValue(first.paymentTerm);

  // Delivery Term at C46
  if (first.deliveryTerm) rfqSheet.getRange("C46").setValue(first.deliveryTerm);

  return "RFQ Sheet populated with " + items.length + " item(s) for PR No. " + prNo;
}


/* =================================
   WRITE TO AOQ SHEET
   Populates the "Abstract of Qoutations" sheet
   from SUMMARY data for the given PR No.
================================= */
function writeToAOQSheet(prNo, suppliersJson, pricingJson, selectionsJson) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var aoqSheet     = ss.getSheetByName("Abstract of Qoutations");

  if (!summarySheet) return "SUMMARY sheet not found.";
  if (!aoqSheet)     return "Abstract of Qoutations sheet not found.";

  log("writeToAOQSheet called — prNo: " + prNo);
  log("suppliersJson: " + (suppliersJson ? suppliersJson.substring(0, 200) : "EMPTY/UNDEFINED"));
  log("pricingJson: " + (pricingJson ? pricingJson.substring(0, 200) : "EMPTY/UNDEFINED"));
  log("selectionsJson: " + (selectionsJson ? selectionsJson.substring(0, 200) : "EMPTY/UNDEFINED"));

  var summaryData = summarySheet.getDataRange().getValues();
  var items = [];
  var nPrNo = cleanPrNo(prNo).toLowerCase();
  for (var i = 1; i < summaryData.length; i++) {
    if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo) {
      var rawDate  = summaryData[i][0];
      var dateObj  = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      var formatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");

      items.push({
        date:            formatted,
        prNo:            cleanPrNo(summaryData[i][1]),
        fundCluster:     summaryData[i][2],
        office:          summaryData[i][3],
        unit:            summaryData[i][4],
        itemDescription: summaryData[i][5],
        quantity:        parseNum(summaryData[i][6]),
        unitCost:        parseNum(summaryData[i][7]),
        totalCost:       parseNum(summaryData[i][8]),
        purpose:         summaryData[i][9],
        deliveryTerm:    summaryData[i][10],
        paymentTerm:     summaryData[i][11],
        deliveryPeriod:  summaryData[i][12],
        supplier:        summaryData[i][13],
        supplierAddress: summaryData[i][14],
        tin:             summaryData[i][15]
      });
    }
  }

  if (items.length === 0) return "No items found for PR No.: " + prNo;
  log("items found: " + items.length);

  var DATA_START_ROW = 16;
  var MAX_ITEM_ROWS  = 35;

  aoqSheet.getRange(DATA_START_ROW, 1, MAX_ITEM_ROWS, 7).clearContent();

  // Clear per-supplier pricing columns
  var supPriceCols = [8, 11, 14, 17, 20]; // H, K, N, Q, T
  for (var ps = 0; ps < supPriceCols.length; ps++) {
    aoqSheet.getRange(DATA_START_ROW, supPriceCols[ps], MAX_ITEM_ROWS, 2).clearContent();
  }

  var rows = items.map(function(item, i) {
    return [
      i + 1,                // A: Item No. (auto-numbered)
      item.quantity,        // B: Quantity
      item.unit,            // C: Unit
      item.itemDescription, // D: Item Description
      "",                   // E: blank
      item.unitCost,        // F: Unit Cost
      item.totalCost        // G: Total Cost
    ];
  });

  aoqSheet.getRange(DATA_START_ROW, 1, rows.length, 7).setValues(rows);

  var first = items[0];
  aoqSheet.getRange("G10").clearContent();
  aoqSheet.getRange("A12").clearContent();
  aoqSheet.getRange("D11").setValue(first.prNo);
  aoqSheet.getRange("E11").clearContent();
  aoqSheet.getRange("F11").setValue("Date: " + first.date);
  if (first.purpose) aoqSheet.getRange("A11").setValue(first.purpose);

  // Clear footer area F42:G45 before writing new values
  aoqSheet.getRange("F42:G45").clearContent();

  var lastDataRow = DATA_START_ROW + items.length - 1;
  aoqSheet.getRange("A43").setValue("TOTAL CONTRACT AMOUNT (FOR AWARD):");
  aoqSheet.getRange("G42").setFormula("=SUM(G" + DATA_START_ROW + ":G" + lastDataRow + ")");
  // Per-supplier grand total formulas (I42, L42, O42, R42, U42)
  var supTotalCols = [9, 12, 15, 18, 21];
  for (var st = 0; st < supTotalCols.length; st++) {
    var cl = String.fromCharCode(64 + supTotalCols[st]);
    aoqSheet.getRange(42, supTotalCols[st]).setFormula("=SUM(" + cl + DATA_START_ROW + ":" + cl + lastDataRow + ")");
  }

  if (first.paymentTerm) aoqSheet.getRange("F44").setValue(first.paymentTerm);
  if (first.deliveryPeriod) aoqSheet.getRange("F45").setValue(first.deliveryPeriod);

  // Write supplier names and addresses into merged cells
  var supplierCells = ['H14', 'K14', 'N14', 'Q14', 'T14'];
  var selectionCells = ['H15', 'K15', 'N15', 'Q15', 'T15'];
  for (var s = 0; s < supplierCells.length; s++) {
    aoqSheet.getRange(supplierCells[s]).clearContent();
    aoqSheet.getRange(selectionCells[s]).clearContent();
  }

  if (suppliersJson) {
    try {
      var suppliers = JSON.parse(suppliersJson);
      log("parsed suppliers count: " + suppliers.length);
      for (var s = 0; s < Math.min(suppliers.length, 5); s++) {
        var sup = suppliers[s];
        log("supplier[" + s + "]: name='" + sup.name + "', address='" + (sup.address || '') + "'");
        if (sup.name) {
          var value = sup.name;
          if (sup.address) value += '\n' + sup.address;
          aoqSheet.getRange(supplierCells[s]).setValue(value);
        }
      }
    } catch (e) {
      Logger.log("Error parsing suppliersJson: " + e);
    }
  } else {
    log("suppliersJson is EMPTY or UNDEFINED — supplier names will NOT be written");
  }

  // Write per-supplier pricing data when provided — FIXED ROW MAPPING
  if (pricingJson) {
    try {
      var pricingData = JSON.parse(pricingJson);
      log("parsed pricing count: " + pricingData.length);

      // Build description→row map from the `items` array (which was written
      // sequentially to the AOQ sheet starting at DATA_START_ROW).
      var aoqItemMap = {};
      for (var mi = 0; mi < items.length; mi++) {
        var desc = String(items[mi].itemDescription || '').trim().toLowerCase();
        if (desc) {
          aoqItemMap[desc] = DATA_START_ROW + mi;
        }
      }
      log("aoqItemMap built with " + Object.keys(aoqItemMap).length + " entries");

      for (var s = 0; s < Math.min(pricingData.length, 5); s++) {
        var pSup = pricingData[s];
        log("pricing[" + s + "]: name='" + (pSup ? pSup.name : 'null') + "', items=" + (pSup && pSup.items ? pSup.items.length : 0));
        if (!pSup || !pSup.items || pSup.items.length === 0) continue;

        var startCol = 8 + s * 3;  // Unit Cost column

        for (var r = 0; r < pSup.items.length; r++) {
          var pItem = pSup.items[r];
          var itemDesc = String(pItem.itemDescription || '').trim();
          if (!itemDesc) {
            Logger.log("Supplier '" + pSup.name + "' item at index " + r + " has no description — skipping");
            continue;
          }

          var lookupKey = itemDesc.toLowerCase();
          var aoqRow = aoqItemMap[lookupKey];

          if (!aoqRow) {
            Logger.log("WARNING: Item '" + itemDesc + "' from supplier '" + pSup.name + "' not found in AOQ — skipping");
            continue;
          }

          var uc = parseNum(pItem.unitCost);
          var tc = parseNum(pItem.totalCost);

          Logger.log("Supplier '" + pSup.name + "' → Item: '" + itemDesc + "' → AOQ Row: " + aoqRow + ", UC: " + uc + ", TC: " + tc + ", Col: " + startCol);

          aoqSheet.getRange(aoqRow, startCol).setValue(uc);
          aoqSheet.getRange(aoqRow, startCol + 1).setValue(tc);
        }
      }
    } catch (e) {
      Logger.log("Error parsing pricingJson: " + e);
    }
  } else {
    log("pricingJson is EMPTY or UNDEFINED — supplier pricing will NOT be written");
  }

  // Write per-supplier selection state (selected item descriptions) to row 15
  if (selectionsJson) {
    try {
      var selectionsData = JSON.parse(selectionsJson);
      for (var s = 0; s < Math.min(selectionsData.length, 5); s++) {
        var selEntry = selectionsData[s];
        if (selEntry && selEntry.selectedItems && selEntry.selectedItems.length > 0) {
          var selCol = 8 + s * 3;
          aoqSheet.getRange(15, selCol).setValue(selEntry.selectedItems.join(','));
        }
      }
    } catch (e) {
      Logger.log("Error parsing selectionsJson: " + e);
    }
  }

  return "Abstract of Qoutations Sheet populated with " + items.length + " item(s) for PR No. " + prNo;
}


/**
 * Writes to AOQ sheet while preserving any existing supplier data.
 * Used during search to avoid wiping supplier names/pricing.
 */
function writeToAOQSheetWithExistingData(prNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aoqSheet = ss.getSheetByName("Abstract of Qoutations");
  if (!aoqSheet) return writeToAOQSheet(prNo);

  // Read existing supplier data before writeToAOQSheet clears it
  var nPrNo = cleanPrNo(prNo).toLowerCase();
  var storedPrNo = cleanPrNo(String(aoqSheet.getRange("D11").getValue() || '')).toLowerCase();
  var suppliers = { list: [], pricing: [] };

  log("writeToAOQSheetWithExistingData — prNo: " + prNo + ", nPrNo: " + nPrNo + ", storedPrNo: '" + storedPrNo + "', match: " + (storedPrNo === nPrNo));

  var supplierCols = [8, 11, 14, 17, 20];
  var selections = [];

  if (storedPrNo === nPrNo) {
    var supplierRow = 14;
    var selectionRow = 15;
    var DATA_START_ROW = 16;
    var MAX_ITEM_ROWS = 35;

    var itemDescs = aoqSheet.getRange(DATA_START_ROW, 4, MAX_ITEM_ROWS, 1).getValues();
    var qtys = aoqSheet.getRange(DATA_START_ROW, 2, MAX_ITEM_ROWS, 1).getValues();

    for (var s = 0; s < supplierCols.length; s++) {
      var col = supplierCols[s];
      var nameCell = aoqSheet.getRange(supplierRow, col).getValue();
      var fullName = String(nameCell || '').trim();

      var name = '';
      var address = '';
      if (fullName) {
        var nlIdx = fullName.indexOf('\n');
        if (nlIdx > -1) {
          name = fullName.substring(0, nlIdx).trim();
          address = fullName.substring(nlIdx + 1).trim();
        } else {
          name = fullName;
        }
      }

      suppliers.list.push({ name: name, address: address });

      var unitCosts = aoqSheet.getRange(DATA_START_ROW, col, MAX_ITEM_ROWS, 1).getValues();
      var totalCosts = aoqSheet.getRange(DATA_START_ROW, col + 1, MAX_ITEM_ROWS, 1).getValues();
      var items = [];
      for (var r = 0; r < MAX_ITEM_ROWS; r++) {
        var desc = String(itemDescs[r] || '').trim();
        if (!desc) continue;
        var uCost = Number(unitCosts[r] || 0);
        var tCost = Number(totalCosts[r] || 0);
        if (tCost === 0 && uCost > 0) {
          var qty = Number(qtys[r] || 0);
          tCost = uCost * qty;
        }
        items.push({ itemDescription: desc, unitCost: uCost, totalCost: tCost });
      }
      suppliers.pricing.push({ name: name, items: items });
    }

    // Also read selection state (row 15)
    for (var s = 0; s < supplierCols.length; s++) {
      var selCell = aoqSheet.getRange(selectionRow, supplierCols[s]).getValue();
      var selStr = String(selCell || '').trim();
      selections.push(selStr);
    }
  } else {
    // Fallback: read supplier data from AOQ_SupplierData when AOQ sheet has a different PR
    var aoqDataSheet = ss.getSheetByName("AOQ_SupplierData");
    if (aoqDataSheet) {
      var aoqData = aoqDataSheet.getDataRange().getValues();
      if (aoqData.length > 0 && String(aoqData[0][0] || '').toLowerCase() === 'pr no.') {
        var aoqRows = [];
        for (var i = 1; i < aoqData.length; i++) {
          if (cleanPrNo(String(aoqData[i][0] || '')).toLowerCase() === nPrNo) {
            aoqRows.push({
              supplierName:    String(aoqData[i][1] || '').trim(),
              supplierAddress: String(aoqData[i][2] || '').trim(),
              itemDescription: aoqData[i][4],
              unitCost:        Number(aoqData[i][5] || 0),
              totalCost:       Number(aoqData[i][6] || 0),
              selected:        String(aoqData[i][7] || '').toUpperCase() === 'TRUE',
              colIdx:          Number(aoqData[i][8] || 0)
            });
          }
        }
        if (aoqRows.length > 0) {
          var groups = {};
          var selByCol = {};
          aoqRows.forEach(function(row) {
            var key = row.colIdx;
            if (!groups[key]) {
              groups[key] = { name: row.supplierName, address: row.supplierAddress, items: [] };
              selByCol[key] = [];
            }
            groups[key].items.push({ itemDescription: row.itemDescription, unitCost: row.unitCost, totalCost: row.totalCost });
            if (row.selected) selByCol[key].push(row.itemDescription);
          });
          // Build arrays aligned by colIdx (0-4) — always 5 entries for correct column mapping
          for (var col = 0; col < 5; col++) {
            if (groups[col] && groups[col].name) {
              suppliers.list.push({ name: groups[col].name, address: groups[col].address });
              suppliers.pricing.push({ name: groups[col].name, items: groups[col].items });
            } else {
              suppliers.list.push({ name: '', address: '' });
              suppliers.pricing.push({ name: '', items: groups[col] ? groups[col].items : [] });
            }
            selections[col] = selByCol[col] ? selByCol[col].join(',') : '';
          }
        }
      }
    }
  }

  log("writeToAOQSheetWithExistingData — suppliers.list: " + suppliers.list.length + ", suppliers.pricing: " + suppliers.pricing.length);

  var suppliersJson = suppliers.list.length > 0 ? JSON.stringify(suppliers.list) : '';
  var pricingJson = suppliers.pricing.length > 0 ? JSON.stringify(suppliers.pricing) : '';

  log("suppliersJson built: " + (suppliersJson ? suppliersJson.substring(0, 150) : "EMPTY"));
  log("pricingJson built: " + (pricingJson ? pricingJson.substring(0, 150) : "EMPTY"));

  // Build selections JSON from row 15 data (preserved during search)
  var selectionsData = [];
  for (var s = 0; s < supplierCols.length; s++) {
    if (selections[s]) {
      var items_arr = selections[s].split(',').map(function(item) { return item.trim(); }).filter(function(item) { return item; });
      if (items_arr.length > 0) {
        selectionsData.push({ supplierIdx: s, selectedItems: items_arr });
      }
    }
  }
  var selectionsJson = selectionsData.length > 0 ? JSON.stringify(selectionsData) : '';
  if (selectionsJson) Logger.log("writeToAOQSheetWithExistingData — selections: " + selectionsJson);

  return writeToAOQSheet(prNo, suppliersJson, pricingJson, selectionsJson);
}


/* =================================
   LOAD SUPPLIER PRICING
   Reads per-supplier pricing data from
   the "Abstract of Qoutations" sheet
   for the given PR No. and returns it
   as JSON for the AOQ form view.
   Schema: { suppliers: [{ name, address, items: [{ itemDescription, unitCost, totalCost }], selectedItems: [string] }] }
   Falls back to SUMMARY sheet if AOQ sheet has no matching PR data.
================================= */
function loadSupplierPricing(prNo) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var aoqSheet = ss.getSheetByName("Abstract of Qoutations");
  var summarySheet = ss.getSheetByName("SUMMARY");

  var supplierCols = [8, 11, 14, 17, 20]; // 1-indexed: H, K, N, Q, T
  var supplierRow = 14;
  var selectionRow = 15;
  var DATA_START_ROW = 16;
  var MAX_ITEM_ROWS = 35;

  // --- Try AOQ sheet first ---
  if (aoqSheet) {
    var nPrNo = cleanPrNo(prNo).toLowerCase();
    var storedPrNo = cleanPrNo(String(aoqSheet.getRange("D11").getValue() || '')).toLowerCase();
    if (storedPrNo === nPrNo) {
      // Read item descriptions (col D) and quantities (col B)
      var itemDescs = aoqSheet.getRange(DATA_START_ROW, 4, MAX_ITEM_ROWS, 1).getValues();
      var qtys = aoqSheet.getRange(DATA_START_ROW, 2, MAX_ITEM_ROWS, 1).getValues();

      var suppliers = [];

      for (var s = 0; s < supplierCols.length; s++) {
        var col = supplierCols[s];
        var nameCell = aoqSheet.getRange(supplierRow, col).getValue();
        var fullName = String(nameCell || '').trim();

        var name = '';
        var address = '';
        if (fullName) {
          var nlIdx = fullName.indexOf('\n');
          if (nlIdx > -1) {
            name = fullName.substring(0, nlIdx).trim();
            address = fullName.substring(nlIdx + 1).trim();
          } else {
            name = fullName;
          }
        }

        // Read selection state from row 15
        var selCell = aoqSheet.getRange(selectionRow, col).getValue();
        var selStr = String(selCell || '').trim();
        var selectedItems = selStr ? selStr.split(',').map(function(item) { return item.trim(); }).filter(function(item) { return item; }) : [];

        var unitCosts = aoqSheet.getRange(DATA_START_ROW, col, MAX_ITEM_ROWS, 1).getValues();
        var totalCosts = aoqSheet.getRange(DATA_START_ROW, col + 1, MAX_ITEM_ROWS, 1).getValues();

        var items = [];
        for (var r = 0; r < MAX_ITEM_ROWS; r++) {
          var desc = String(itemDescs[r] || '').trim();
          if (!desc) continue;
          var uCost = Number(unitCosts[r] || 0);
          var tCost = Number(totalCosts[r] || 0);
          if (tCost === 0 && uCost > 0) {
            var qty = Number(qtys[r] || 0);
            tCost = uCost * qty;
          }
          items.push({
            itemDescription: desc,
            unitCost: uCost,
            totalCost: tCost
          });
        }

        suppliers.push({
          name: name,
          address: address,
          items: items,
          selectedItems: selectedItems
        });
      }

      return JSON.stringify({ suppliers: suppliers });
    }
  }

  // --- Fallback: reconstruct from AOQ_SupplierData sheet ---
  var aoqDataSheet = ss.getSheetByName("AOQ_SupplierData");
  if (aoqDataSheet) {
    var aoqData = aoqDataSheet.getDataRange().getValues();
    var nPrNo2 = cleanPrNo(prNo).toLowerCase();
    // Read header row to validate columns
    if (aoqData.length > 0 && String(aoqData[0][0] || '').toLowerCase() === 'pr no.') {
      var aoqRows = [];
      for (var i = 1; i < aoqData.length; i++) {
        if (cleanPrNo(String(aoqData[i][0] || '')).toLowerCase() === nPrNo2) {
          aoqRows.push({
            prNo:            aoqData[i][0],
            supplierName:    String(aoqData[i][1] || '').trim(),
            supplierAddress: String(aoqData[i][2] || '').trim(),
            tin:             String(aoqData[i][3] || '').trim(),
            itemDescription: aoqData[i][4],
            unitCost:        Number(aoqData[i][5] || 0),
            totalCost:       Number(aoqData[i][6] || 0),
            selected:        String(aoqData[i][7] || '').toUpperCase() === 'TRUE',
            colIdx:          Number(aoqData[i][8] || 0),
            poNo:            String(aoqData[i][9] || '').trim()
          });
        }
      }

      if (aoqRows.length > 0) {
        // Group by column index (supplier)
        var supplierGroups = {};
        aoqRows.forEach(function(row) {
          var key = row.colIdx;
          if (!supplierGroups[key]) {
            supplierGroups[key] = {
              name: row.supplierName,
              address: row.supplierAddress,
              poNo: row.poNo,
              colIdx: key,
              items: [],
              selectedItems: []
            };
          }
          supplierGroups[key].items.push({
            itemDescription: row.itemDescription,
            unitCost: row.unitCost,
            totalCost: row.totalCost
          });
          if (row.selected) {
            supplierGroups[key].selectedItems.push(row.itemDescription);
          }
        });

        var fallbackSuppliers = [];
        var keys = Object.keys(supplierGroups).sort();
        keys.forEach(function(k) {
          var g = supplierGroups[k];
          if (g.name) {
            fallbackSuppliers.push({
              name: g.name,
              address: g.address,
              poNo: g.poNo,
              items: g.items,
              selectedItems: g.selectedItems
            });
          }
        });

        if (fallbackSuppliers.length > 0) {
          return JSON.stringify({ suppliers: fallbackSuppliers });
        }
      }
    }
  }

  // --- Last resort fallback: reconstruct from SUMMARY sheet ---
  if (summarySheet) {
    var summaryData = summarySheet.getDataRange().getValues();
    var nPrNo3 = cleanPrNo(prNo).toLowerCase();
    var summaryItems = [];
    for (var i = 1; i < summaryData.length; i++) {
      if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo3) {
        summaryItems.push({
          itemDescription: summaryData[i][5],
          quantity: parseNum(summaryData[i][6]),
          unitCost: parseNum(summaryData[i][7]),
          totalCost: parseNum(summaryData[i][8]),
          supplier: summaryData[i][13],
          supplierAddress: summaryData[i][14]
        });
      }
    }

    if (summaryItems.length > 0) {
      var supplierMap = {};
      summaryItems.forEach(function(si) {
        var sName = String(si.supplier || '').trim();
        if (!sName) return;
        if (!supplierMap[sName]) {
          supplierMap[sName] = { name: sName, address: String(si.supplierAddress || '').trim(), items: [] };
        }
        supplierMap[sName].items.push({
          itemDescription: si.itemDescription,
          unitCost: si.unitCost,
          totalCost: si.totalCost
        });
      });

      var fallbackSuppliers = [];
      for (var key in supplierMap) {
        if (supplierMap.hasOwnProperty(key)) {
          fallbackSuppliers.push(supplierMap[key]);
        }
      }

      if (fallbackSuppliers.length > 0) {
        return JSON.stringify({ suppliers: fallbackSuppliers });
      }
    }
  }

  return JSON.stringify({ suppliers: [] });
}


/* =================================
   WRITE TO PO (PURCHASE ORDER) SHEET
   Populates the "Purchase Order" sheet
   from SUMMARY data — uses stored
   totalCost (col I) for accurate grand total.
================================= */
function writeToPOSheet(prNo) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var aoqDataSheet = ss.getSheetByName("AOQ_SupplierData");
  var summarySheet = ss.getSheetByName("SUMMARY");
  var poSheet      = ss.getSheetByName("Purchase Order");

  if (!aoqDataSheet) return "AOQ_SupplierData sheet not found.";
  if (!summarySheet) return "SUMMARY sheet not found.";
  if (!poSheet)      return "Purchase Order sheet not found.";

  var aoqData = aoqDataSheet.getDataRange().getValues();
  var nPrNo = cleanPrNo(prNo).toLowerCase();

  // Build lookup of selected items from AOQ_SupplierData
  var selectedPrices = {};
  var supplierName = '';
  var supplierAddress = '';
  var supplierTin = '';

  for (var i = 1; i < aoqData.length; i++) {
    if (cleanPrNo(String(aoqData[i][0] || '')).toLowerCase() === nPrNo &&
        String(aoqData[i][7] || '').toUpperCase() === 'TRUE') {
      var desc = String(aoqData[i][4] || '').trim().toLowerCase();
      if (desc) {
        selectedPrices[desc] = {
          unitCost:  Number(aoqData[i][5] || 0),
          totalCost: Number(aoqData[i][6] || 0)
        };
      }
      if (!supplierName) {
        supplierName    = String(aoqData[i][1] || '').trim();
        supplierAddress = String(aoqData[i][2] || '').trim();
        supplierTin     = String(aoqData[i][3] || '').trim();
      }
    }
  }

  if (Object.keys(selectedPrices).length === 0) {
    return "No selected items found for PR No.: " + prNo;
  }

  // Read item details from SUMMARY, using AOQ_SupplierData pricing
  var summaryData = summarySheet.getDataRange().getValues();
  var items = [];

  for (var i = 1; i < summaryData.length; i++) {
    if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo) {
      var desc = String(summaryData[i][5] || '').trim().toLowerCase();
      if (!selectedPrices[desc]) continue;

      var rawDate  = summaryData[i][0];
      var dateObj  = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      var formatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");

      var qty       = parseNum(summaryData[i][6]);
      var unitCost  = selectedPrices[desc].unitCost;
      var totalCost = selectedPrices[desc].totalCost;
      if (totalCost === 0 && qty > 0 && unitCost > 0) {
        totalCost = qty * unitCost;
      }

      items.push({
        date:            formatted,
        prNo:            cleanPrNo(summaryData[i][1]),
        fundCluster:     summaryData[i][2],
        office:          summaryData[i][3],
        unit:            summaryData[i][4],
        itemDescription: summaryData[i][5],
        quantity:        parseNum(summaryData[i][6]),
        unitCost:        parseNum(summaryData[i][7]),
        totalCost:       parseNum(summaryData[i][8]),
        purpose:         summaryData[i][9],
        deliveryTerm:    summaryData[i][10],
        paymentTerm:     summaryData[i][11],
        deliveryPeriod:  summaryData[i][12],
        supplier:        summaryData[i][13],
        supplierAddress: summaryData[i][14],
        tin:             summaryData[i][15],
        modeProcurement: summaryData[i][16],
        placeDelivery:   summaryData[i][17],
        dateDelivery:    formattedD
      });
    }
  }

  if (items.length === 0) return "No items found for PR No.: " + prNo;

  var DATA_START_ROW = 19;
  var MAX_ITEM_ROWS  = 26;

  poSheet.getRange(DATA_START_ROW, 1, MAX_ITEM_ROWS, 7).clearContent();

  var first = items[0];

  poSheet.getRange("A10").setValue("Supplier: " + first.supplier || "");
  poSheet.getRange("A11").setValue("Address: " + first.supplierAddress || "");
  poSheet.getRange("A12").setValue("Tin: " + first.tin || "");

  poSheet.getRange("E11").setValue("Date: " + first.date);

  poSheet.getRange("E12").setValue("Mode of Procurement: " + first.modeProcurement || "");

  poSheet.getRange("A16").setValue("Place of Delivery: " + first.placeDelivery || "");
  if (first.dateDelivery) {
    var dd = (first.dateDelivery instanceof Date) ? first.dateDelivery : new Date(first.dateDelivery);
    poSheet.getRange("A17").setValue(Utilities.formatDate(dd, Session.getScriptTimeZone(), "MM/dd/yyyy"));
  }

  poSheet.getRange("E16").setValue("Delivery Term: " + first.deliveryTerm || "");
  poSheet.getRange("E17").setValue("Payment Term: " + first.paymentTerm || "");

  var rows = items.map(function(item) {
    return [
      "",                   // A: Stock/Property No.
      item.unit,            // B: Unit
      item.itemDescription, // C: Item Description (merged C:D)
      "",                   // D: blank (merged with C)
      item.quantity,        // E: Quantity
      item.unitCost,        // F: Unit Cost
      item.totalCost        // G: Amount
    ];
  });

  poSheet.getRange(DATA_START_ROW, 1, rows.length, 7).setValues(rows);

  var lastDataRow    = DATA_START_ROW + items.length - 1;
  var grandTotalCell = poSheet.getRange(45, 7);
  grandTotalCell.setFormula("=SUM(G" + DATA_START_ROW + ":G" + lastDataRow + ")");

  return "Purchase Order sheet populated with " + items.length + " item(s) for PR No. " + prNo;
}


/* =================================
   SAVE SUPPLIER DATA
   Saves supplier-filled RFQ rows
   back to SUMMARY (updating unitCost,
   totalCost, supplier fields), then
   writes to RFQ and AOQ sheets.
================================= */
function saveSupplierData(supplierRows) {
  if (!supplierRows || supplierRows.length === 0) return "No supplier data to save!";

  var sheet     = getSheet();
  var sheetData = sheet.getDataRange().getValues();
  var prNo      = supplierRows[0].prNo;

  // Match existing rows in SUMMARY by prNo + itemDescription and update them
  var updatedCount = 0;

  var nPrNo = cleanPrNo(prNo).toLowerCase();
  for (var s = 0; s < supplierRows.length; s++) {
    var sd   = supplierRows[s];
    var desc = String(sd.itemDescription || '').trim().toLowerCase();

    for (var r = 1; r < sheetData.length; r++) {
      if (
        cleanPrNo(sheetData[r][1]).toLowerCase() === nPrNo &&
        String(sheetData[r][5]).trim().toLowerCase() === desc
      ) {
        var rowNum    = r + 1;
        var qty       = Number(sheetData[r][6] || 0);
        var unitCost  = Number(sd.unitCost || 0);
        var totalCost = unitCost * qty;

        // Update unitCost (col H=8), totalCost (col I=9), supplier (col N=14),
        // supplierAddress (col O=15), TIN (col P=16)
        sheet.getRange(rowNum, 8).setValue(unitCost);
        sheet.getRange(rowNum, 9).setValue(totalCost);
        sheet.getRange(rowNum, 14).setValue(sd.supplier        || '');
        sheet.getRange(rowNum, 15).setValue(sd.supplierAddress || '');
        sheet.getRange(rowNum, 16).setValue(sd.tin             || '');
        updatedCount++;
        break; // each item only matches once
      }
    }
  }

  if (updatedCount === 0) {
    return "No matching items found in SUMMARY sheet for PR No. " + prNo + ". Make sure RFQ items are saved first.";
  }

  // After updating SUMMARY, generate PO number for tracking and write to PO sheet
  var poNo = '';
  var supplierName = String(supplierRows[0].supplier || '').trim();
  if (supplierName) {
    poNo = getNextUniquePONo(prNo);
  }
  try { writeToPOSheet(prNo); } catch(e) { Logger.log("writeToPOSheet error: " + e); }

  return JSON.stringify({ message: updatedCount + " supplier item(s) saved Successfully! PO sheet updated.", poNo: poNo });
}

/**
 * Ensures the AOQ_SupplierData sheet exists, creating it if needed.
 * This sheet stores all supplier quotations per PR, avoiding the
 * single-supplier limit of the SUMMARY sheet.
 */
function ensureAOQSupplierDataSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("AOQ_SupplierData");
  if (sheet) return sheet;
  sheet = ss.insertSheet("AOQ_SupplierData");
  sheet.appendRow(["PR No.", "Supplier Name", "Supplier Address", "TIN", "Item Description", "Unit Cost", "Total Cost", "Selected", "Col Idx", "PO No."]);
  sheet.setFrozenRows(1);
  return sheet;
}

/**
 * Saves the full supplier data set for a PR No. into the AOQ_SupplierData
 * sheet and re-populates the Abstract of Qoutations sheet.
 * Called from the client after all per-supplier saveSupplierData calls.
 *
 * @param {string} prNo - The PR number
 * @param {string} allDataJson - JSON string with structure:
 *   { suppliers: [{name, address, items: [{itemDescription, unitCost, totalCost}]}],
 *     selections: [{supplierIdx, selectedItems: [string]}],
 *     supplierRows: [{prNo, itemDescription, quantity, unitCost, totalCost, supplier, supplierAddress, selected}] }
 */
function saveFullSupplierData(prNo, allDataJson) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var allData = JSON.parse(allDataJson);

  // --- Write to AOQ_SupplierData sheet ---
  var dataSheet = ensureAOQSupplierDataSheet();
  var data = dataSheet.getDataRange().getValues();
  var nPrNo = cleanPrNo(prNo).toLowerCase();

  // Remove existing rows for this PR No., preserving header + other PRs
  var rowsToKeep = [];
  // Always use the full 10-column header
  rowsToKeep.push(["PR No.", "Supplier Name", "Supplier Address", "TIN", "Item Description", "Unit Cost", "Total Cost", "Selected", "Col Idx", "PO No."]);
  for (var i = 1; i < data.length; i++) {
    if (cleanPrNo(String(data[i][0] || '')).toLowerCase() !== nPrNo) {
      var row = data[i];
      while (row.length < 10) row.push('');  // pad to 10 columns if needed
      rowsToKeep.push(row);
    }
  }

  dataSheet.clearContents();
  if (rowsToKeep.length > 0) {
    dataSheet.getRange(1, 1, rowsToKeep.length, 10).setValues(rowsToKeep);
  }

  // Append new rows for this PR
  var newRows = [];
  var allSuppliers = allData.suppliers || [];
  var supplierRows = allData.supplierRows || [];
  var supplierPONos = allData.supplierPONos || {};

  // Build a lookup: prNo+itemDescription -> { unitCost, totalCost, supplier, supplierAddress, selected }
  var rowLookup = {};
  supplierRows.forEach(function(sr) {
    var key = (cleanPrNo(sr.prNo).toLowerCase() + '|' + (sr.itemDescription || '').trim().toLowerCase());
    rowLookup[key] = sr;
  });

  for (var s = 0; s < allSuppliers.length; s++) {
    var sup = allSuppliers[s];
    if (!sup.name) continue;
    var poNo = supplierPONos[sup.name] || '';
    for (var i = 0; i < (sup.items || []).length; i++) {
      var item = sup.items[i];
      var lookupKey = (cleanPrNo(prNo).toLowerCase() + '|' + (item.itemDescription || '').trim().toLowerCase());
      var rowData = rowLookup[lookupKey] || {};
      newRows.push([
        prNo,
        sup.name,
        sup.address || '',
        rowData.tin || '',
        item.itemDescription,
        Number(item.unitCost) || 0,
        Number(item.totalCost) || 0,
        rowData.selected !== false ? 'TRUE' : 'FALSE',
        s,  // Col Idx (0-4)
        poNo  // PO No.
      ]);
    }
  }

  if (newRows.length > 0) {
    var startRow = dataSheet.getLastRow() > 0 ? dataSheet.getLastRow() + 1 : 2;
    dataSheet.getRange(startRow, 1, newRows.length, 10).setValues(newRows);
  }

  // --- Re-populate Abstract of Qoutations sheet ---
  // Always produce exactly 5 entries for correct column mapping
  var suppliersList = [];
  var pricingList = [];
  for (var si = 0; si < 5; si++) {
    var sup = allSuppliers[si] || { name: '', address: '', items: [] };
    suppliersList.push({ name: sup.name || '', address: sup.address || '' });
    pricingList.push({ name: sup.name || '', items: sup.items || [] });
  }

  var selectionsData = allData.selections || [];
  var selectionsJson = selectionsData.length > 0 ? JSON.stringify(selectionsData) : '';

  return writeToAOQSheet(
    prNo,
    JSON.stringify(suppliersList),
    JSON.stringify(pricingList),
    selectionsJson
  );
}

/* =================================
   CREATE SUPPLIER RFQ SHEET
   Copies the RFQ template and
   populates it with supplier data
   from SUMMARY for the given PR.
   Sheet name: "RFQ - {supplierName}"
   Leaves the original RFQ tab
   untouched (template only).
================================= */
function createSupplierRFQSheet(prNo, supplierName) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var rfqTemplate  = ss.getSheetByName("RFQ");

  if (!summarySheet) return "SUMMARY sheet not found.";
  if (!rfqTemplate)  return "RFQ template sheet not found.";

  var summaryData = summarySheet.getDataRange().getValues();
  var items = [];
  var nPrNo = cleanPrNo(prNo).toLowerCase();

  for (var i = 1; i < summaryData.length; i++) {
    if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo) {
      var rawDate  = summaryData[i][0];
      var dateObj  = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      var formatted = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "MM/dd/yyyy");

      items.push({
        date:            formatted,
        prNo:            cleanPrNo(summaryData[i][1]),
        fundCluster:     summaryData[i][2],
        office:          summaryData[i][3],
        unit:            summaryData[i][4],
        itemDescription: summaryData[i][5],
        quantity:        parseNum(summaryData[i][6]),
        unitCost:        parseNum(summaryData[i][7]),
        totalCost:       parseNum(summaryData[i][8]),
        purpose:         summaryData[i][9],
        deliveryTerm:    summaryData[i][10],
        paymentTerm:     summaryData[i][11],
        deliveryPeriod:  summaryData[i][12],
        supplier:        summaryData[i][13],
        supplierAddress: summaryData[i][14],
        tin:             summaryData[i][15]
      });
    }
  }

  if (items.length === 0) return "No items found for PR No.: " + prNo;

  var first = items[0];
  var rawName = String(supplierName || first.supplier || 'Unknown').trim();
  var baseName = rawName.replace(/[\[\]:\?\*\/\\]/g, '').substring(0, 95);
  if (!baseName) baseName = 'Unknown';
  var sheetName = 'RFQ - ' + baseName;

  // Handle duplicate sheet names
  var existing = ss.getSheetByName(sheetName);
  var counter = 1;
  while (existing) {
    counter++;
    sheetName = 'RFQ - ' + baseName + ' (' + counter + ')';
    existing = ss.getSheetByName(sheetName);
  }

  // Copy the RFQ template
  var rfqSheet = rfqTemplate.copyTo(ss);
  rfqSheet.setName(sheetName);

  var DATA_START_ROW = 18;

  // Populate the copy with same layout as writeToRFQSheet
  // RFQ Date Selection at F8 (merged F:G) — from the RFQ form's date picker
  var rfqDate = first.deliveryPeriod ? Utilities.formatDate(new Date(first.deliveryPeriod), Session.getScriptTimeZone(), "MM/dd/yyyy") : "";
  rfqSheet.getRange("F8").setValue("Date: " + rfqDate);

  // Date label at F9 (merged F:G)
  rfqSheet.getRange("F9").setValue("Date");

  // Supplier label at A9 (merged A:C)
  rfqSheet.getRange("A9").setValue("Supplier");

  // Supplier Address label at A10 (merged A:C)
  rfqSheet.getRange("A10").setValue("Supplier Address");

  // QUOTATION/OFFER label at F15 (merged F:G and 15:16)
  rfqSheet.getRange("F15").setValue("QUOTATION/OFFER");

  // Concatenated intro text + due date (date from F8 + 3 days) at A13 (merged A:G and 13:14)
  var f8Value = rfqSheet.getRange("F8").getValue();
  var f8Date = (f8Value instanceof Date) ? f8Value : new Date(f8Value);
  f8Date.setDate(f8Date.getDate() + 3);
  var dueDate = Utilities.formatDate(f8Date, Session.getScriptTimeZone(), "MM/dd/yyyy");
  var introText = "This Office is undergoing procurement of labor and materials for the repainting of the interior walls.  As a known supplier that typically provides the required goods and services, we are sending this Request of Quotation.  Please quote your possible lowest price by filling in the blanks below and return to this Office duly accomplished on or before";
  rfqSheet.getRange("A13").setValue(introText + " " + dueDate);

  // Write items (cols A-E)
  var rows = items.map(function(item) {
    return [
      item.quantity,
      item.unit,
      item.itemDescription,
      item.unitCost,
      item.totalCost
    ];
  });

  rfqSheet.getRange(DATA_START_ROW, 1, rows.length, 5).setValues(rows);

  var lastDataRow = DATA_START_ROW + items.length - 1;
  rfqSheet.getRange("E43").setFormula("=SUM(E" + DATA_START_ROW + ":E" + lastDataRow + ")");

  if (first.paymentTerm) rfqSheet.getRange("C45").setValue(first.paymentTerm);
  if (first.deliveryTerm) rfqSheet.getRange("C46").setValue(first.deliveryTerm);

  return "Supplier RFQ sheet created: \"" + sheetName + "\"";
}





function writeToPRSheet(prNo) {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var summarySheet = ss.getSheetByName("SUMMARY");
  var prSheet      = ss.getSheetByName("Purchase Request");

  if (!summarySheet || !prSheet) return "Sheet not found.";

  var summaryData = summarySheet.getDataRange().getValues();

  var items = [];
  var nPrNo = cleanPrNo(prNo).toLowerCase();

  // FIX: start at index 1 (skip header row only — row index 0)
  // Previously started at 4, which silently skipped the first 3 data rows.
  for (var i = 1; i < summaryData.length; i++) {
    if (cleanPrNo(summaryData[i][1]).toLowerCase() === nPrNo) {
      var qty      = parseNum(summaryData[i][6]);  // col G → Quantity
      var unitCost = parseNum(summaryData[i][7]);  // col H → Unit Cost

      // FIX: Read the pre-computed Total Cost from col I (index 8)
      // instead of recomputing, so it matches exactly what was saved.
      // Safety: if stored total is 0 but qty and cost aren't, recompute.
      var totalCost = parseNum(summaryData[i][8]); // col I → Total Cost
      if (totalCost === 0 && qty > 0 && unitCost > 0) {
        totalCost = qty * unitCost;
      }

      // FIX: Date formatting — removed the stray "Date: " literal that was
      // accidentally inside the format pattern string in the original code.
      var rawDate  = summaryData[i][0];
      var dateObj  = (rawDate instanceof Date) ? rawDate : new Date(rawDate);
      var formatted = Utilities.formatDate(
        dateObj,
        Session.getScriptTimeZone(),
        "MM/dd/yyyy"   // ← clean pattern, no embedded label
      );

      items.push({
        date:            formatted,
        prNo:            cleanPrNo(summaryData[i][1]),
        fundCluster:     summaryData[i][2],
        office:          summaryData[i][3],
        unit:            summaryData[i][4],
        itemDescription: summaryData[i][5],
        quantity:        qty,
        unitCost:        unitCost,
        totalCost:       totalCost,   // FIX: use stored value, not recomputed
        purpose:         summaryData[i][9]
      });
    }
  }

  if (items.length === 0) return "No items found for PR No.: " + prNo;

  var DATA_START_ROW = 15;
  var MAX_ITEM_ROWS  = 26;

  // Clear previous data in item rows (cols A-G)
  prSheet.getRange(DATA_START_ROW, 1, MAX_ITEM_ROWS, 7).clearContent();

  // Build row array — columns A–G:
  // A: Stock/Property No. (blank)
  // B: Unit
  // C: Item Description (C+D merged on the template sheet)
  // D: (merged into C)
  // E: Quantity
  // F: Unit Cost
  // G: Total Cost
  var rows = items.map(function(item) {
    return [
      "",                   // A: Stock/Property No.
      item.unit,            // B: Unit
      item.itemDescription, // C: Item Description (merged C+D)
      "",                   // D: (blank — merged with C)
      item.quantity,        // E: Quantity
      item.unitCost,        // F: Unit Cost
      item.totalCost        // G: Total Cost
    ];
  });

  prSheet.getRange(DATA_START_ROW, 1, rows.length, 7).setValues(rows);

  // ── Header fields ──────────────────────────────────────────
  var first = items[0];

  // Fund Cluster → G10
  prSheet.getRange("G10").setValue(first.fundCluster || "");

  // Office/Section → A11
  prSheet.getRange("A11").setValue("Office/Section : " + (first.office || ""));

  // PR No. → D11
  prSheet.getRange("D11").setValue(first.prNo);
  prSheet.getRange("E11").clearContent();

  // FIX: Date value uses the cleanly-formatted string (no stray label)
  prSheet.getRange("F11").setValue("Date: " + first.date);

  // Clear row 42, set Purpose at row 43
  prSheet.getRange(42, 1, 1, 6).clearContent();
  if (first.purpose) {
    prSheet.getRange(43, 1).setValue("Purpose: " + first.purpose);
  }

  // Grand Total formula — SUM of col G (Total Cost) across all written item rows
  var lastDataRow    = DATA_START_ROW + items.length - 1;
  var grandTotalCell = prSheet.getRange(DATA_START_ROW + MAX_ITEM_ROWS, 7); // row 41, col G
  grandTotalCell.setFormula("=SUM(G" + DATA_START_ROW + ":G" + lastDataRow + ")");

  return "PR Sheet populated with " + items.length + " item(s) for PR No. " + prNo;
}