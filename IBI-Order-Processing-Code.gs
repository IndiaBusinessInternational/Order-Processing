function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);

    // ── ROUTE: pull pending orders from the IBI Package Tracker ─────────────
    // Staff upload the day's PDFs in the Package Tracker; this relays the saved
    // records server-to-server (no CORS), and drops any Order ID we've already
    // processed here, so the form lists only what's still pending.
    if (data && data.action === 'getTrackerOrders') {
      var trk;
      try { trk = handleGetTrackerOrders_(); }
      catch (ex) { trk = { ok:false, error: ex.toString(), orders:[] }; }
      return ContentService.createTextOutput(JSON.stringify(trk)).setMimeType(ContentService.MimeType.JSON);
    }

    // ── ROUTE: dismiss Tracker orders (hide them from the pending list) ─────
    if (data && data.action === 'dismissTrackerOrders') {
      var dsm;
      try { dsm = handleDismissTrackerOrders_(data); }
      catch (ex) { dsm = { ok:false, error: ex.toString() }; }
      return ContentService.createTextOutput(JSON.stringify(dsm)).setMimeType(ContentService.MimeType.JSON);
    }

    var ss    = SpreadsheetApp.openById("1Y1sE5fPODjevfYJXhTeJzi0djWGo5pdl2xy-obxhO0Q");
    var sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];

    // GENERATE TIMESTAMP IN IST (Asia/Kolkata)
    // Always correct regardless of browser timezone or sheet timezone
    var now = new Date();
    var entryDateTime = Utilities.formatDate(now, "Asia/Kolkata", "dd MMMM yyyy  HH:mm:ss");

    // AUTO-CREATE HEADER ROW if sheet is empty
    if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() === "") {
      var headers = [
        "Entry Date & Time",
        "Serial Number",
        "Pickup Date",
        "Pickup Day",
        "Pickup Time",
        "Platform",
        "Courier Service",
        "Product Name",
        "Retail Price (Rs)",
        "Payment Type",
        "Prod L (cm)",
        "Prod W (cm)",
        "Prod H (cm)",
        "Prod Wt (gm)",
        "Pkg L (cm)",
        "Pkg W (cm)",
        "Pkg H (cm)",
        "Pkg Wt (gm)",
        "B.Wt (gms)",
        "Shipment Type",
        "Name of Buyer",
        "Location / City",
        "State / UT"
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length)
           .setBackground("#1e293b")
           .setFontColor("#ffffff")
           .setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    // Ensure the "Order ID" column exists (added for IBI ERP unification).
    // Works for a freshly-created sheet and for an older one made before this
    // column — it is appended as a new trailing header so existing rows stay
    // aligned and new rows can carry the marketplace Order ID.
    var hdrRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasOrderId = hdrRow.some(function (h) { return String(h).trim() === "Order ID"; });
    if (!hasOrderId) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Order ID")
           .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    }

    // Ensure the "Invoice Number" column exists (same trailing-append approach).
    var hdrRow2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasInvoice = hdrRow2.some(function (h) { return String(h).trim() === "Invoice Number"; });
    if (!hasInvoice) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Invoice Number")
           .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold");
    }

    // SERIAL NUMBER: starts from 604042, auto-increments
    var SERIAL_START = parseInt(data.serialStart) || 604042;
    var lastRow = sheet.getLastRow();
    var lastSerial = 0;
    if (lastRow > 1) {
      var v = sheet.getRange(lastRow, 2).getValue();
      lastSerial = parseInt(v) || 0;
    }
    var serialNumber = Math.max(SERIAL_START, lastSerial + 1);

    // APPEND DATA ROW
    var newRowNum = sheet.getLastRow() + 1;
    sheet.appendRow([
      entryDateTime,
      serialNumber,
      data.pickupDate,
      data.pickupDay,
      data.pickupTime,
      data.platform,
      data.courier,
      data.productName,
      data.retailPrice,
      data.paymentType,
      data.pLen,
      data.pWid,
      data.pHei,
      data.pWgt,
      data.pkgLen,
      data.pkgWid,
      data.pkgHei,
      data.pkgWgt,
      data.bWgt,
      data.shipmentTo,
      data.buyerName,
      data.location,
      data.state
    ]);

    // FORCE Column A (Entry Date & Time) to Plain Text
    // Prevents Google Sheets from re-interpreting it as a date serial
    sheet.getRange(newRowNum, 1).setNumberFormat("@");

    // Write the marketplace Order ID into its column (located by header name, so
    // it lands correctly wherever that column sits). This is the key the IBI ERP
    // uses to unify this packing entry with the order it imports from the Tracker.
    var hdr2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var oc = 0; oc < hdr2.length; oc++) {
      if (String(hdr2[oc]).trim() === "Order ID") {
        sheet.getRange(newRowNum, oc + 1).setNumberFormat("@").setValue(data.orderId || "");
        break;
      }
    }
    for (var ic = 0; ic < hdr2.length; ic++) {
      if (String(hdr2[ic]).trim() === "Invoice Number") {
        sheet.getRange(newRowNum, ic + 1).setNumberFormat("@").setValue(data.invoiceNumber || "");
        break;
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", serialNumber: serialNumber }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ============================================================================
   READ ENDPOINT — added for IBI ERP sync (doPost above is unchanged)
   ----------------------------------------------------------------------------
   Returns every row of the "Orders" sheet as JSON so the IBI ERP can pull the
   day's processed orders. Supports JSONP (?callback=fn) so the ERP reads it
   cross-origin with no CORS setup — the same mechanism the Pickup Manager and
   Package Tracker backends use. Each record is keyed by the sheet's own header
   names (Serial Number, Platform, Product Name, …) so the ERP can map fields
   by name and stays correct even if column order changes later.
   ============================================================================ */
function doGet(e) {
  var cb = (e && e.parameter && e.parameter.callback) ? String(e.parameter.callback) : '';
  try {
    var ss    = SpreadsheetApp.openById("1Y1sE5fPODjevfYJXhTeJzi0djWGo5pdl2xy-obxhO0Q");
    var sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];

    var orders = [];
    if (sheet.getLastRow() > 1) {
      var values  = sheet.getDataRange().getValues();
      var headers = values[0].map(function (h) { return String(h).trim(); });
      for (var r = 1; r < values.length; r++) {
        var row = values[r];
        // Skip blank rows (no Serial Number in column B)
        if (row[1] === '' || row[1] === null || typeof row[1] === 'undefined') continue;
        var rec = {};
        for (var c = 0; c < headers.length; c++) {
          var val = row[c];
          rec[headers[c]] = (val instanceof Date)
            ? Utilities.formatDate(val, "Asia/Kolkata", "yyyy-MM-dd")
            : val;
        }
        orders.push(rec);
      }
    }

    var payload = JSON.stringify({ status: "success", count: orders.length, orders: orders });
    return _send_(payload, cb);

  } catch (err) {
    return _send_(JSON.stringify({ status: "error", message: err.toString() }), cb);
  }
}

// Shared responder: JSONP when a callback is supplied, plain JSON otherwise.
function _send_(payload, cb) {
  if (cb) {
    return ContentService
      .createTextOutput(cb + "(" + payload + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

/* ════════════════════════════════════════════════════════════════════════
   handleGetTrackerOrders_
   Reads saved records from the IBI Package Tracker (its loadPackages action),
   server-to-server via UrlFetchApp so there is no browser CORS issue, then
   removes any Order ID already present in our own "Orders" sheet. Returns the
   still-pending orders (newest first) for the form to list and pre-fill.
   ──────────────────────────────────────────────────────────────────────── */
function handleGetTrackerOrders_() {
  // Live /exec URL of the IBI Package Tracker web app.
  var TRACKER_URL = 'https://script.google.com/macros/s/AKfycbxLPmBJmkgFbO-5Ew5uObdVGwMNBXQ8oJCmssIue_Av_G2iaRAaArGzbRNaF7gMzFYsEg/exec';

  // 1) Pull all packages from the Package Tracker.
  var pkgs = [];
  var resp = UrlFetchApp.fetch(TRACKER_URL, {
    method: 'post',
    payload: { action: 'loadPackages' },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var json = JSON.parse(resp.getContentText());
  if (json && json.status === 'success' && Array.isArray(json.data)) pkgs = json.data;

  // 2) Order IDs we have already processed (present in our Orders sheet).
  var processed = {};
  try {
    var ss    = SpreadsheetApp.openById("1Y1sE5fPODjevfYJXhTeJzi0djWGo5pdl2xy-obxhO0Q");
    var sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];
    var values = sheet.getDataRange().getValues();
    if (values.length) {
      var hdr = values[0], oc = -1;
      for (var c = 0; c < hdr.length; c++) { if (String(hdr[c]).trim() === "Order ID") { oc = c; break; } }
      if (oc >= 0) {
        for (var r = 1; r < values.length; r++) {
          var id = String(values[r][oc] || '').trim();
          if (id) processed[id] = true;
        }
      }
    }
  } catch (e2) { /* if the processed list can't be read, just show everything */ }

  // 2b) Keys we have explicitly dismissed (kept in the "Dismissed" tab).
  var dismissed = _opDismissedSet_();

  // 3) Keep only unprocessed, non-dismissed orders, de-duped, newest first
  //    (loadPackages is already newest-first).
  var out = { ok: true, orders: [], total: pkgs.length };
  var seen = {};
  for (var i = 0; i < pkgs.length; i++) {
    var p   = pkgs[i];
    var oid = String(p.orderId || '').trim();
    if (oid && processed[oid]) continue;
    var key = oid || String(p.awb || '') || String(p.invoiceNo || '');
    if (key && dismissed[key]) continue;
    if (key && seen[key]) continue;
    if (key) seen[key] = true;
    out.orders.push({
      key:        key,
      orderId:    oid,
      invoiceNo:  String(p.invoiceNo  || ''),
      platform:   String(p.platform   || ''),
      courier:    String(p.courier    || ''),
      awb:        String(p.awb        || ''),
      shipDate:   String(p.shipDate   || ''),
      orderDate:  String(p.orderDate  || ''),
      products:   String(p.products   || ''),
      qty:        String(p.qty        || ''),
      amount:     String(p.amount     || ''),
      payType:    String(p.payType    || ''),
      buyerName:  String(p.buyerName  || ''),
      buyerPhone: String(p.buyerPhone || ''),
      address:    String(p.shippingAddress || ''),
      pincode:    String(p.pincode    || ''),
      savedOn:    String(p.savedOn    || '')
    });
  }
  return out;
}

/* ════════════════════════════════════════════════════════════════════════
   authorizeExternalRequest  —  RUN THIS ONCE
   In the Apps Script editor, choose this function from the dropdown, click
   Run, and press "Allow" on the permission dialog. That grants the
   script.external_request permission UrlFetchApp needs. Then redeploy a New
   version (Deploy → Manage deployments → Edit → New version).
   ──────────────────────────────────────────────────────────────────────── */
function authorizeExternalRequest() {
  var TRACKER_URL = 'https://script.google.com/macros/s/AKfycbxLPmBJmkgFbO-5Ew5uObdVGwMNBXQ8oJCmssIue_Av_G2iaRAaArGzbRNaF7gMzFYsEg/exec';
  var r = UrlFetchApp.fetch(TRACKER_URL, { method:'post', payload:{ action:'loadPackages' }, muteHttpExceptions:true, followRedirects:true });
  Logger.log('External request OK. HTTP ' + r.getResponseCode());
}

/* ════════════════════════════════════════════════════════════════════════
   DISMISSED BACKLOG
   The "Dismissed" tab (in the Orders spreadsheet) holds the keys of Tracker
   records we never want to process here — used to clear the legacy backlog so
   the pending count reflects only real work. Records are NOT deleted from the
   Package Tracker; they are simply hidden from this form's pending list. To
   restore everything, just delete the rows in the "Dismissed" tab.
   ──────────────────────────────────────────────────────────────────────── */
function _opDismissedSheet_() {
  var ss = SpreadsheetApp.openById("1Y1sE5fPODjevfYJXhTeJzi0djWGo5pdl2xy-obxhO0Q");
  var sh = ss.getSheetByName("Dismissed");
  if (!sh) {
    sh = ss.insertSheet("Dismissed");
    sh.appendRow(["Key", "Dismissed On", "Buyer", "Platform", "Saved On"]);
  }
  return sh;
}
function _opDismissedSet_() {
  var set = {};
  try {
    var sh = _opDismissedSheet_();
    var v  = sh.getDataRange().getValues();
    for (var i = 1; i < v.length; i++) {
      var k = String(v[i][0] || '').trim();
      if (k) set[k] = true;
    }
  } catch (e) {}
  return set;
}
function handleDismissTrackerOrders_(data) {
  var items = (data && data.items) || [];
  if (!items.length && data && data.keys) {
    items = data.keys.map(function (k) { return { key: k }; });
  }
  if (!items.length) return { ok: true, dismissed: 0 };

  var sh       = _opDismissedSheet_();
  var existing = _opDismissedSet_();
  var now      = Utilities.formatDate(new Date(), "Asia/Kolkata", "dd/MM/yyyy HH:mm:ss");
  var rows = [], added = 0;
  for (var i = 0; i < items.length; i++) {
    var it = items[i] || {};
    var k  = String(it.key || '').trim();
    if (!k || existing[k]) continue;
    existing[k] = true;
    rows.push([k, now, String(it.buyerName || ''), String(it.platform || ''), String(it.savedOn || '')]);
    added++;
  }
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  return { ok: true, dismissed: added };
}
