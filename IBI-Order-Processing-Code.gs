/* IBI Package Tracker — its spreadsheet (read directly, fast path) and its web
   app (fallback only). Same Google account owns both, and this project's
   manifest already holds the full spreadsheets scope, so no re-authorisation. */
var OP_TRACKER_SHEET_ID   = '1VjK5oA6mCZVXZ2AhfZYtb0kVAkB549bSUB4iQ4eW7f0';
var OP_TRACKER_TAB        = 'Package Tracker';
var OP_TRACKER_URL        = 'https://script.google.com/macros/s/AKfycbxLPmBJmkgFbO-5Ew5uObdVGwMNBXQ8oJCmssIue_Av_G2iaRAaArGzbRNaF7gMzFYsEg/exec';
var OP_PENDING_CACHE_KEY  = 'op_pending_v1';
var OP_PENDING_CACHE_SEC  = 60;

function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);

    // ── ROUTE: pull pending orders from the IBI Package Tracker ─────────────
    // Staff upload the day's PDFs in the Package Tracker; this relays the saved
    // records server-to-server (no CORS), and drops any Order ID we've already
    // processed here, so the form lists only what's still pending.
    if (data && data.action === 'getTrackerOrders') {
      var trk;
      try { trk = handleGetTrackerOrders_(data); }
      catch (ex) { trk = { ok:false, error: ex.toString(), orders:[] }; }
      return ContentService.createTextOutput(JSON.stringify(trk)).setMimeType(ContentService.MimeType.JSON);
    }

    // ── ROUTE: verify the Owner PIN, issue a short-lived signed token ──────
    // The PIN itself lives ONLY in Script Properties (Project Settings →
    // Script Properties → OWNER_PIN). It is never sent back to the browser;
    // a successful check returns an HMAC-signed token the client stores for
    // the session and presents on privileged actions below.
    if (data && data.action === 'verifyOwnerPin') {
      var vp;
      try { vp = handleVerifyOwnerPin_(data); }
      catch (ex) { vp = { ok:false, error: ex.toString() }; }
      return ContentService.createTextOutput(JSON.stringify(vp)).setMimeType(ContentService.MimeType.JSON);
    }

    // ── ROUTE: dismiss Tracker orders (hide them from the pending list) ─────
    // OWNER-ONLY: requires a valid token from verifyOwnerPin. Without it the
    // endpoint refuses, so knowing the /exec URL alone is not enough.
    if (data && data.action === 'dismissTrackerOrders') {
      if (!_opCheckOwnerToken_(data && data.token)) {
        return ContentService.createTextOutput(JSON.stringify({ ok:false, error:'unauthorized' })).setMimeType(ContentService.MimeType.JSON);
      }
      var dsm;
      try { CacheService.getScriptCache().remove(OP_PENDING_CACHE_KEY); } catch (eC3) {}
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

    // Ensure the "Qty Sold" column exists. Until now the units-sold figure was
    // collected on the form, used to deduct jewellery stock, and then thrown
    // away — so a 2-unit order was saved as a row indistinguishable from a
    // 1-unit one, and every downstream reader (ERP, audits, reprints) saw one
    // piece. It is a real column now. "Retail Price (Rs)" beside it stays the
    // ORDER TOTAL, not a per-unit price — do not multiply the two.
    var hdrRow3 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var hasQty = hdrRow3.some(function (h) { return String(h).trim() === "Qty Sold"; });
    if (!hasQty) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue("Qty Sold")
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

    // Units packed. Written as a NUMBER so it can be summed, and never left
    // blank — a blank would read as "unknown" and invite the old guess of 1.
    var qtySold = parseInt(data.saleQty, 10);
    if (isNaN(qtySold) || qtySold < 1) qtySold = 1;
    for (var qc = 0; qc < hdr2.length; qc++) {
      if (String(hdr2[qc]).trim() === "Qty Sold") {
        sheet.getRange(newRowNum, qc + 1).setValue(qtySold);
        break;
      }
    }

    // This order is no longer pending — drop the cached list so another device
    // opening the form does not offer it again for up to a minute.
    try { CacheService.getScriptCache().remove(OP_PENDING_CACHE_KEY); } catch (eC2) {}

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
function handleGetTrackerOrders_(data) {
  // A 60-second cache. Staff open this page all day and several devices hit it
  // at once; without it every open paid the full build again.
  var cache = CacheService.getScriptCache();
  if (!(data && data.fresh)) {
    var hit = cache.get(OP_PENDING_CACHE_KEY);
    if (hit) { try { var c = JSON.parse(hit); c.cached = true; return c; } catch (eC) {} }
  }

  // 1) Packages from the Package Tracker — read STRAIGHT off its spreadsheet.
  //    The old path called the Tracker's own web app for `loadPackages`, which
  //    serialises all 1,518 rows × 23 fields (~1.5 MB) and took ~6 s, so that
  //    the form could keep about three of them. It also burned a second Apps
  //    Script execution per page open, and this account's quota is shared by
  //    every IBI app. Reading the sheet in-process needs no new permission
  //    (the manifest already carries the full spreadsheets scope) and returns
  //    only the fields this form uses. The web app stays as the fallback.
  var pkgs = null;
  try { pkgs = _opReadTrackerSheet_(); } catch (e1) { pkgs = null; }
  if (!pkgs) pkgs = _opReadTrackerHttp_();

  // 2) Order IDs we have already processed (present in our Orders sheet).
  //    Only the Order ID column is read — pulling all 25 columns × 1,442 rows
  //    to look at one of them was pure waste.
  var processed = {};
  try {
    var ss    = SpreadsheetApp.openById("1Y1sE5fPODjevfYJXhTeJzi0djWGo5pdl2xy-obxhO0Q");
    var sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];
    var lastR = sheet.getLastRow();
    if (lastR > 1) {
      var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0], oc = -1;
      for (var c = 0; c < hdr.length; c++) { if (String(hdr[c]).trim() === "Order ID") { oc = c; break; } }
      if (oc >= 0) {
        var ids = sheet.getRange(2, oc + 1, lastR - 1, 1).getValues();
        for (var r = 0; r < ids.length; r++) {
          var id = String(ids[r][0] || '').trim();
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

  // CacheService rejects values over 100 KB — skip the cache rather than throw.
  try {
    var payload = JSON.stringify(out);
    if (payload.length < 95000) cache.put(OP_PENDING_CACHE_KEY, payload, OP_PENDING_CACHE_SEC);
  } catch (e3) {}
  return out;
}

/* ── Tracker sheet, read directly ──────────────────────────────────────────
   Columns are located BY HEADER NAME, never by fixed position: if the Tracker
   ever inserts a column this keeps working, and if the headers it needs are
   missing it returns null so the caller falls back to the web app instead of
   silently handing back nonsense. */
function _opReadTrackerSheet_() {
  var sh = SpreadsheetApp.openById(OP_TRACKER_SHEET_ID).getSheetByName(OP_TRACKER_TAB);
  if (!sh) return null;
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0]
               .map(function (h) { return String(h).trim().toLowerCase(); });
  function col(name) { return head.indexOf(String(name).toLowerCase()); }

  var iSaved = col('Saved On'),   iPlat  = col('Platform'),  iCour = col('Courier'),
      iOrder = col('Order ID'),   iAwb   = col('AWB / Tracking No'),
      iInv   = col('Invoice No'), iOdate = col('Order Date'),
      iBuyer = col('Buyer Name'), iPhone = col('Buyer Phone'),
      iAddr  = col('Shipping Address'), iPin = col('Pincode'),
      iProd  = col('Products / SKU'),   iQty = col('Qty'),
      iAmt   = col('Amount (₹)'),       iPay = col('Payment Type'),
      iShip  = col('Ship Date');
  if (iSaved < 0 || iOrder < 0 || iAwb < 0 || iQty < 0) return null;   // layout changed

  var rng  = sh.getRange(2, 1, lastRow - 1, lastCol);
  var vals = rng.getValues();
  var disp = rng.getDisplayValues();          // literal text, for the date fix below
  function g(row, i) { return i < 0 ? '' : row[i]; }

  var out = [];
  for (var i = vals.length - 1; i >= 0; i--) {   // bottom row first = newest first
    var r = vals[i];
    if (!r[iAwb] && !r[iOrder]) continue;
    out.push({
      savedOn:         _opSavedOn_(r[iSaved], disp[i][iSaved]),
      platform:        g(r, iPlat),  courier:    g(r, iCour),
      orderId:         g(r, iOrder), awb:        g(r, iAwb),
      invoiceNo:       g(r, iInv),   orderDate:  g(r, iOdate),
      buyerName:       g(r, iBuyer), buyerPhone: g(r, iPhone),
      shippingAddress: g(r, iAddr),  pincode:    g(r, iPin),
      products:        g(r, iProd),  qty:        g(r, iQty),
      amount:          g(r, iAmt),   payType:    g(r, iPay),
      shipDate:        g(r, iShip)
    });
  }
  return out;
}

/* Fallback: the Tracker's own web app, as before. */
function _opReadTrackerHttp_() {
  var resp = UrlFetchApp.fetch(OP_TRACKER_URL, {
    method: 'post',
    payload: { action: 'loadPackages' },
    muteHttpExceptions: true,
    followRedirects: true
  });
  var json = JSON.parse(resp.getContentText());
  return (json && json.status === 'success' && Array.isArray(json.data)) ? json.data : [];
}

/* "Saved On", day-first — ported from the Tracker's savedOnOut_.
   ⚠ The cell comes back with DAY AND MONTH TRANSPOSED whenever the day is 12
   or lower, so the displayed text is parsed day-first rather than trusting the
   Date object. A bare spreadsheet serial (cells left on General format) is
   converted too, otherwise those rows are invisible to every date filter. */
function _opSavedOn_(raw, shown) {
  var m = String(shown || '').trim()
    .match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})(?:[ ,T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) {
    function p2(v) { return ('0' + v).slice(-2); }
    return m[3] + '-' + p2(m[2]) + '-' + p2(m[1]) + 'T' + p2(m[4] || '00') + ':' + p2(m[5] || '00') + ':' + p2(m[6] || '00');
  }
  if (raw instanceof Date && !isNaN(raw.getTime())) {
    return Utilities.formatDate(raw, 'Asia/Kolkata', "yyyy-MM-dd'T'HH:mm:ss");
  }
  if (typeof raw === 'number' && isFinite(raw) && raw > 20000 && raw < 90000) {
    return Utilities.formatDate(new Date(Math.round((raw - 25569) * 86400000)), 'UTC', "yyyy-MM-dd'T'HH:mm:ss");
  }
  return raw;
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

/* ════════════════════════════════════════════════════════════════════════
   OWNER AUTHENTICATION  (PIN check moved server-side)
   ----------------------------------------------------------------------------
   The Owner PIN is stored in Script Properties (OWNER_PIN) — never in this
   source file and never sent to the browser. verifyOwnerPin compares the
   submitted PIN, and on a match returns a stateless token: "<exp>.<sig>",
   where sig = HMAC-SHA256(OWNER_TOKEN_SECRET, exp). Privileged actions
   recompute the signature and check expiry — no server-side session storage.

   ONE-TIME SETUP (do this once in the Apps Script editor):
     1. Run  opSetupOwnerAuth  → generates OWNER_TOKEN_SECRET automatically.
     2. Project Settings → Script Properties → add  OWNER_PIN  = your PIN.
        (Use a NEW pin, not the old 8899 that leaked in the screenshot.)
     3. Deploy → Manage deployments → Edit → New version.
   ──────────────────────────────────────────────────────────────────────── */
var OP_TOKEN_TTL_MS = 12 * 60 * 60 * 1000;   // token valid for 12 hours

function opSetupOwnerAuth() {
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('OWNER_TOKEN_SECRET')) {
    props.setProperty('OWNER_TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  Logger.log('OWNER_TOKEN_SECRET: set ✓');
  Logger.log('OWNER_PIN: ' + (props.getProperty('OWNER_PIN')
    ? 'set ✓'
    : 'NOT SET — add it under Project Settings → Script Properties'));
}

function handleVerifyOwnerPin_(data) {
  var props = PropertiesService.getScriptProperties();
  var realPin = props.getProperty('OWNER_PIN');
  if (!realPin) return { ok:false, error:'OWNER_PIN not configured' };
  var given = (data && data.pin != null) ? String(data.pin) : '';
  if (given !== String(realPin)) return { ok:false, error:'incorrect' };
  return { ok:true, token: _opMakeOwnerToken_() };
}

function _opSign_(msg) {
  var secret = PropertiesService.getScriptProperties().getProperty('OWNER_TOKEN_SECRET') || '';
  var raw = Utilities.computeHmacSha256Signature(String(msg), secret);
  return Utilities.base64EncodeWebSafe(raw);
}

function _opMakeOwnerToken_() {
  var exp = (new Date()).getTime() + OP_TOKEN_TTL_MS;
  return exp + '.' + _opSign_(String(exp));
}

function _opCheckOwnerToken_(token) {
  if (!token || typeof token !== 'string') return false;
  var dot = token.indexOf('.');
  if (dot < 1) return false;
  var exp = token.substring(0, dot);
  var sig = token.substring(dot + 1);
  if (!/^\d+$/.test(exp)) return false;
  if ((new Date()).getTime() > parseInt(exp, 10)) return false;   // expired
  var expected = _opSign_(exp);
  if (sig.length !== expected.length) return false;
  // constant-ish comparison
  var diff = 0;
  for (var i = 0; i < expected.length; i++) diff |= (sig.charCodeAt(i) ^ expected.charCodeAt(i));
  return diff === 0;
}
