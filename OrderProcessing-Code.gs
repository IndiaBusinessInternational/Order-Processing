function doPost(e) {
  try {
    var data  = JSON.parse(e.postData.contents);
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
