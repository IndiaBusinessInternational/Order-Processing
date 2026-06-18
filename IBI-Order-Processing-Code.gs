function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);

    // ── ROUTE: image scan → OCR + DeepSeek extraction (no sheet write) ──────
    // The form posts {action:'scanExtract', imageBase64, imageMime} when staff
    // scan a label/invoice. We OCR the image and ask DeepSeek to return the
    // fields as JSON, then send them back to pre-fill the form (staff verify).
    if (data && data.action === 'scanExtract') {
      var out;
      try { out = handleScanExtract_(data); }
      catch (ex) { out = { status: 'error', message: ex.toString() }; }
      return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
    }

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

/* ============================================================================
   SCAN EXTRACTION — OCR (Google Drive) + DeepSeek field extraction
   ----------------------------------------------------------------------------
   Called from doPost when action === 'scanExtract'. Pipeline:
     1. Decode the camera image and OCR it with Google Drive's built-in OCR
        (free, server-side — no second API key, no CORS).
     2. Send the recognised text to DeepSeek with a strict JSON-only prompt.
     3. Recover the JSON even if the model wraps it in prose, and normalise.
   The DeepSeek key lives ONLY here, in Script Properties (never in the page).

   SETUP (one time):
     • Project Settings → Script properties → add  DEEPSEEK_API_KEY = sk-…
     • Editor → Services (+) → add "Drive API" (advanced service) so the OCR
       call below works. Then Deploy → Manage deployments → New version.
   ============================================================================ */
function handleScanExtract_(data) {
  var b64  = String(data.imageBase64 || '');
  var mime = String(data.imageMime || 'image/jpeg');
  if (!b64) return { status: 'error', message: 'No image supplied' };

  var ocrText = ocrImage_(Utilities.newBlob(Utilities.base64Decode(b64), mime, 'scan'));
  if (!ocrText || !ocrText.trim()) {
    return { status: 'error', message: 'OCR found no readable text — enter the details manually' };
  }

  var fields = deepSeekExtract_(ocrText);

  // Meesho ships via Delhivery and the label carries no Ship Date — default it
  // to today (IST) so staff don't have to look it up.
  var todayIST = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyy-MM-dd');
  if (String(fields.platform || '').toLowerCase().indexOf('meesho') >= 0 && !fields.shipDate) {
    fields.shipDate = todayIST;
  }

  return { status: 'success', fields: fields, ocrText: ocrText };
}

// Google Drive OCR: upload the image as a Google Doc with OCR, read the text,
// then delete the temporary file. Requires the "Drive API" advanced service.
function ocrImage_(blob) {
  var tmp = Drive.Files.insert(
    { title: 'ibi-ocr-' + Date.now(), mimeType: 'application/vnd.google-apps.document' },
    blob,
    { ocr: true, ocrLanguage: 'en' }
  );
  var text = '';
  try { text = DocumentApp.openById(tmp.id).getBody().getText(); } catch (e) {}
  try { Drive.Files.remove(tmp.id); }
  catch (e) { try { DriveApp.getFileById(tmp.id).setTrashed(true); } catch (e2) {} }
  return text;
}

// Ask DeepSeek to turn the OCR text into the exact fields we need, as JSON only.
function deepSeekExtract_(ocrText) {
  var key = PropertiesService.getScriptProperties().getProperty('DEEPSEEK_API_KEY');
  if (!key) throw new Error('DEEPSEEK_API_KEY is not set in Script Properties');

  var sys = [
    'You read Indian e-commerce shipping labels and tax invoices and return ONLY a JSON object.',
    'No prose, no markdown, no code fences — just the JSON.',
    'Keys (use "" when a value is not present):',
    'orderId, invoiceNumber, platform, shipDate, productName, retailPrice, paymentType, buyerName, city, state.',
    'platform must be one of: Amazon, Amazon Bazaar, Flipkart, Shopsy, ShopClues, Meesho, IBI Website, Offline.',
    'shipDate must be YYYY-MM-DD (convert any printed date), else "".',
    'retailPrice: digits only (no currency symbol/commas).',
    'paymentType: "Prepaid" or "COD".',
    'Flipkart Order IDs look like OD + 15-22 digits; Amazon like 408-1234567-1234567.'
  ].join(' ');

  var payload = {
    model: 'deepseek-chat',
    temperature: 0,
    response_format: { type: 'json_object' },
    messages: [
      { role: 'system', content: sys },
      { role: 'user',   content: 'Extract from this label/invoice text:\n\n' + ocrText }
    ]
  };

  var res = UrlFetchApp.fetch('https://api.deepseek.com/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code === 401 || code === 403) throw new Error('DeepSeek rejected the API key (HTTP ' + code + ')');
  if (code === 402) throw new Error('DeepSeek billing exhausted (HTTP 402) — top up credits');
  if (code === 429) throw new Error('DeepSeek rate-limited (HTTP 429) — try again in a moment');
  if (code < 200 || code >= 300) throw new Error('DeepSeek HTTP ' + code + ': ' + body.slice(0, 300));

  var data = JSON.parse(body);
  var content = (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || '';
  var f = recoverJson_(content) || {};

  // Normalise to exactly the keys we expect, all strings.
  var keys = ['orderId','invoiceNumber','platform','shipDate','productName','retailPrice','paymentType','buyerName','city','state'];
  var out = {};
  keys.forEach(function (k) { out[k] = (f[k] == null) ? '' : String(f[k]).trim(); });
  out.retailPrice = out.retailPrice.replace(/[^0-9.]/g, '');
  return out;
}

// Recover a JSON object from a model reply even if it's wrapped in prose or
// ```json fences — balanced-brace scan from the first "{" (same robust approach
// the Package Tracker uses across providers).
function recoverJson_(s) {
  s = String(s || '').replace(/```json/gi, '').replace(/```/g, '').trim();
  try { return JSON.parse(s); } catch (e) {}
  var start = s.indexOf('{');
  if (start < 0) return null;
  var depth = 0, inStr = false, esc = false;
  for (var i = start; i < s.length; i++) {
    var c = s[i];
    if (inStr) {
      if (esc) esc = false;
      else if (c === '\\') esc = true;
      else if (c === '"') inStr = false;
    } else if (c === '"') inStr = true;
    else if (c === '{') depth++;
    else if (c === '}') { depth--; if (depth === 0) {
      try { return JSON.parse(s.slice(start, i + 1)); } catch (e) { return null; }
    } }
  }
  return null;
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

  // 3) Keep only unprocessed orders, de-duped, newest first (loadPackages is
  //    already newest-first).
  var out = { ok: true, orders: [], total: pkgs.length };
  var seen = {};
  for (var i = 0; i < pkgs.length; i++) {
    var p   = pkgs[i];
    var oid = String(p.orderId || '').trim();
    if (oid && processed[oid]) continue;
    var key = oid || String(p.awb || '') || String(p.invoiceNo || '');
    if (key && seen[key]) continue;
    if (key) seen[key] = true;
    out.orders.push({
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
