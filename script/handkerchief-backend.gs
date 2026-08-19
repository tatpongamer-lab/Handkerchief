const CONFIG = {
  SHEET_ORDERS: "orders",
  SHEET_META: "meta",
  DRIVE_FOLDER_ID: "1qXN26r_43RSLwt9YSofEcTXfXSE8Cxx5",
  ADMIN_PASSWORD: "2530",
  ORDER_PREFIX: "HANKY",
  PRODUCT_PRICE: 200,
  SHIPPING_FEE: 30,
  PRODUCT_NAME: "แคแสดในความทรงจำ | Cassia Memory"
};

// ================= doPost — Entry Point =================
function doPost(e) {
  try {
    if (!e.postData || !e.postData.contents) {
      return json({ success: false, message: "No payload" });
    }

    const data = JSON.parse(e.postData.contents);

    if (!data.action) {
      return json({ success: false, message: "Missing action" });
    }

    switch (data.action) {
      case "createOrder":       return handleCreateOrder(data);
      case "getProducts":       return handleGetProducts();
      case "searchOrders":      return handleSearchOrders(data);
      case "getAdminData":      return handleGetAdminData(data);
      case "updateOrderStatus": return handleUpdateOrderStatus(data);
      case "toggleStore":       return handleToggleStore(data);
      case "setPreorderCutoff": return handleSetPreorderCutoff(data);
      default:
        return json({ success: false, message: "Invalid action: " + escapeHtml(data.action) });
    }

  } catch (err) {
    return json({ success: false, message: escapeHtml(err.toString()) });
  }
}

// ================= doGet — Entry Point =================
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || "";

    if (e.parameter && e.parameter.payload) {
      var data = JSON.parse(e.parameter.payload);
      if (!data.action) return json({ success: false, message: "Missing action in payload" });

      switch (data.action) {
        case "getAdminData":      return handleGetAdminData(data);
        case "updateOrderStatus": return handleUpdateOrderStatus(data);
        case "toggleStore":       return handleToggleStore(data);
        case "setPreorderCutoff": return handleSetPreorderCutoff(data);
        case "searchOrders":      return handleSearchOrders(data);
        default:
          return json({ success: false, message: "Invalid payload action: " + escapeHtml(data.action) });
      }
    }

    switch (action) {
      case "getProducts": return handleGetProducts();
      default:
        return json({ success: false, message: "Invalid action" });
    }
  } catch (err) {
    return json({ success: false, message: escapeHtml(err.toString()) });
  }
}

// =============================================================================
// getProducts — ดึงข้อมูลสินค้า + สถานะร้าน + กำหนด Pre-order
// =============================================================================
function handleGetProducts() {
  const meta = getMetaMap();

  return json({
    success: true,
    products: [{
      id: "hanky",
      name: CONFIG.PRODUCT_NAME,
      price: CONFIG.PRODUCT_PRICE
    }],
    storeStatus: meta["STORE_STATUS"] || "OPEN",
    preorderCutoff: meta["PREORDER_CUTOFF"] || ""
  });
}

// =============================================================================
// createOrder — สร้างคำสั่งซื้อ
// =============================================================================
function handleCreateOrder(data) {
  // ─── Validation ───
  if (!data.name)           throw "Missing name";
  if (!data.phone)          throw "Missing phone";
  if (!data.deliveryMethod) throw "Missing deliveryMethod";
  if (!data.quantity || toInt(data.quantity) < 1) throw "Invalid quantity";

  // ─── Deduplication ด้วย CacheService ───
  const cache = CacheService.getScriptCache();
  const requestId = data.requestId || "";
  if (requestId) {
    if (cache.get("order_" + requestId)) {
      return json({ success: true, message: "duplicate" });
    }
    cache.put("order_" + requestId, "1", 600);
  }

  const meta = getMetaMap();
  const isAdminFreeOrder = verifyPassword(data.adminPassword);
  
  if (!isAdminFreeOrder && (meta["STORE_STATUS"] || "OPEN") === "CLOSED") {
    return json({ success: false, message: "ขออภัย ร้านค้าปิดรับคำสั่งซื้อชั่วคราว" });
  }

  // ─── LockService ───
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    // ─── Server-side Calculation ───
    const qty = toInt(data.quantity);
    let subtotal = qty * CONFIG.PRODUCT_PRICE;
    let shippingFee = (data.deliveryMethod === "ship") ? CONFIG.SHIPPING_FEE : 0;
    
    if (isAdminFreeOrder) {
      subtotal = 0;
      shippingFee = 0;
    }
    
    const total = subtotal + shippingFee;

    // ─── สร้าง Order ID ───
    const orderId = getNextOrderId();

    // ─── อัปโหลดสลิป ───
    let slipUrl = "";
    if (!isAdminFreeOrder && data.slipBase64) {
      if (Array.isArray(data.slipBase64)) {
        let urls = [];
        for (let i = 0; i < data.slipBase64.length; i++) {
           urls.push(saveSlipToDrive(data.slipBase64[i], orderId + "-" + (i+1)));
        }
        slipUrl = urls.join(",");
      } else {
        slipUrl = saveSlipToDrive(data.slipBase64, orderId);
      }
    }

    let paymentStatus = "PENDING";
    if (isAdminFreeOrder) {
      paymentStatus = "PAID";
    }

    // ─── บันทึกลง Sheet ───
    const now = new Date();
    const sheet = getSheet(CONFIG.SHEET_ORDERS);
    
    const customerName = isAdminFreeOrder ? (data.name || "") + " (Sponsor)" : (data.name || "");

    const row = [[
      orderId,                           // 0: OrderID
      customerName,                      // 1: CustomerName
      data.phone || "",                  // 2: Phone
      data.deliveryMethod || "",         // 3: DeliveryMethod
      data.address || "",                // 4: Address
      qty,                               // 5: Quantity
      subtotal,                          // 6: Subtotal
      shippingFee,                       // 7: ShippingFee
      total,                             // 8: Total
      paymentStatus,                     // 9: PaymentStatus
      "UNFULFILLED",                     // 10: FulfillmentStatus
      slipUrl,                           // 11: SlipURL
      "",                                // 12: TrackingNumber
      now                                // 13: CreatedAt
    ]];

    const newRow = sheet.getLastRow() + 1;
    sheet.getRange(newRow, 1, 1, row[0].length).setValues(row);
    // Phone text format
    sheet.getRange(newRow, 3).setNumberFormat("@");
    sheet.getRange(newRow, 3).setValue(data.phone || "");

    return json({
      success: true,
      orderId: orderId,
      paymentStatus: paymentStatus
    });

  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// searchOrders — ค้นหาคำสั่งซื้อด้วยเบอร์โทร
// =============================================================================
function handleSearchOrders(data) {
  const phoneQuery = String(data.phone || "").trim();
  if (!phoneQuery) {
    return json({ success: false, message: "กรุณากรอกเบอร์โทรศัพท์" });
  }

  const sheet = getSheet(CONFIG.SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();
  const results = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const sheetPhone = String(r[2]).trim().replace(/^0+/, "");
    const qPhone = phoneQuery.replace(/^0+/, "");
    
    if (sheetPhone === qPhone) {
      results.push({
        orderId:           r[0],
        customerName:      r[1],
        phone:             r[2],
        deliveryMethod:    r[3],
        address:           r[4],
        quantity:          toInt(r[5]),
        subtotal:          toInt(r[6]),
        shippingFee:       toInt(r[7]),
        total:             toInt(r[8]),
        paymentStatus:     r[9]  || "PENDING",
        fulfillmentStatus: r[10] || "UNFULFILLED",
        slipUrl:           r[11] || "",
        trackingNumber:    r[12] || "",
        date:              r[13] || ""
      });
    }
  }

  results.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });

  return json({ success: true, orders: results });
}

// =============================================================================
// getAdminData — ดึงข้อมูลทั้งหมด
// =============================================================================
function handleGetAdminData(data) {
  if (!verifyPassword(data.password)) {
    return json({ success: false, message: "รหัสผ่านไม่ถูกต้อง" });
  }

  const sheet = getSheet(CONFIG.SHEET_ORDERS);
  const rows = sheet.getDataRange().getValues();

  const orders = [];
  const stats = { totalOrders: 0, totalRevenue: 0, pendingPayments: 0 };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const total = toInt(r[8]);
    const pStatus = r[9] || "PENDING";

    stats.totalOrders++;
    if (pStatus === "PAID") stats.totalRevenue += total;
    if (pStatus === "PENDING") stats.pendingPayments++;

    orders.push({
      orderId:           r[0],
      customerName:      r[1],
      phone:             r[2],
      deliveryMethod:    r[3],
      address:           r[4],
      quantity:          toInt(r[5]),
      subtotal:          toInt(r[6]),
      shippingFee:       toInt(r[7]),
      total:             total,
      paymentStatus:     pStatus,
      fulfillmentStatus: r[10] || "UNFULFILLED",
      slipUrl:           r[11] || "",
      trackingNumber:    r[12] || "",
      date:              r[13] || "",
      carrier:           "",
      adminNote:         ""
    });
  }

  orders.sort(function(a, b) { return new Date(b.date) - new Date(a.date); });

  const meta = getMetaMap();

  return json({
    success: true,
    stats: stats,
    orders: orders,
    storeStatus: meta["STORE_STATUS"] || "OPEN",
    preorderCutoff: meta["PREORDER_CUTOFF"] || ""
  });
}

// =============================================================================
// updateOrderStatus — อัปเดตสถานะคำสั่งซื้อ (Admin)
// =============================================================================
function handleUpdateOrderStatus(data) {
  if (!verifyPassword(data.password)) {
    return json({ success: false, message: "Unauthorized" });
  }
  if (!data.orderId) throw "Missing orderId";

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {
    const sheet = getSheet(CONFIG.SHEET_ORDERS);
    const rows = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === data.orderId) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      return json({ success: false, message: "ไม่พบคำสั่งซื้อ: " + escapeHtml(data.orderId) });
    }

    if (data.paymentStatus !== undefined && data.paymentStatus !== null) {
      sheet.getRange(rowIndex, 10).setValue(data.paymentStatus);
    }
    if (data.fulfillmentStatus !== undefined && data.fulfillmentStatus !== null) {
      sheet.getRange(rowIndex, 11).setValue(data.fulfillmentStatus);
    }
    if (data.trackingNumber !== undefined && data.trackingNumber !== null) {
      sheet.getRange(rowIndex, 13).setValue(data.trackingNumber);
    }
    // Also supporting carrier/adminNote if needed in future (using cols beyond 14) but let's stick to spec.

    return json({ success: true });

  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// toggleStore — เปิด/ปิดร้าน (Admin)
// =============================================================================
function handleToggleStore(data) {
  if (!verifyPassword(data.password)) {
    return json({ success: false, message: "Unauthorized" });
  }

  const status = (data.status === "CLOSED") ? "CLOSED" : "OPEN";
  setMetaValue("STORE_STATUS", status);

  return json({ success: true, storeStatus: status });
}

// =============================================================================
// setPreorderCutoff — กำหนดวันหมดเขต Pre-order (Admin)
// =============================================================================
function handleSetPreorderCutoff(data) {
  if (!verifyPassword(data.password)) {
    return json({ success: false, message: "Unauthorized" });
  }

  const cutoffDate = data.cutoffDate || "";
  setMetaValue("PREORDER_CUTOFF", cutoffDate);

  return json({ success: true, preorderCutoff: cutoffDate });
}

// =============================================================================
// ORDER ID COUNTER
// =============================================================================
function getNextOrderId() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getSheet(CONFIG.SHEET_META);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === "HANKY_COUNTER") {
        rowIndex = i + 1;
        break;
      }
    }

    let next = 1;
    if (rowIndex === -1) {
      sheet.appendRow(["HANKY_COUNTER", next]);
    } else {
      const current = toInt(sheet.getRange(rowIndex, 2).getValue());
      next = current + 1;
      sheet.getRange(rowIndex, 2).setValue(next);
    }

    return CONFIG.ORDER_PREFIX + "-" + String(next).padStart(4, "0");
  } finally {
    lock.releaseLock();
  }
}

// =============================================================================
// META SHEET
// =============================================================================
function getMetaMap() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_META);
  if (!sheet || sheet.getLastRow() < 1) return {};

  const data = sheet.getDataRange().getValues();
  const map = {};
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (key) map[key] = data[i][1] !== undefined ? String(data[i][1]).trim() : "";
  }
  return map;
}

function setMetaValue(key, value) {
  const sheet = getSheet(CONFIG.SHEET_META);
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// =============================================================================
// SLIP UPLOAD
// =============================================================================
function saveSlipToDrive(base64Data, orderId) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

    let rawBase64 = base64Data;
    if (base64Data.indexOf(",") > -1) {
      rawBase64 = base64Data.split(",")[1];
    }

    const decoded = Utilities.base64Decode(rawBase64);
    const blob = Utilities.newBlob(decoded, "image/jpeg", orderId + ".jpg");
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return "https://lh3.googleusercontent.com/d/" + file.getId();

  } catch (err) {
    throw "Drive Upload Error: " + err.toString();
  }
}

// =============================================================================
// UTILS
// =============================================================================
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function toInt(v) {
  return parseInt(v || 0, 10) || 0;
}

function escapeHtml(unsafe) {
  return String(unsafe)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function verifyPassword(password) {
  return password === CONFIG.ADMIN_PASSWORD;
}

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw "Sheet not found: " + name;
  return sheet;
}
