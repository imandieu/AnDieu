/**
 * HỆ THỐNG QUẢN LÝ TỰ ĐỘNG NHQT - PHIÊN BẢN CUỐI CÙNG (MASTER)
 * Tự động hóa: Rule, Công thức (CT), Xuất đơn (OUT), Nhật ký (AutoLog)
 * Tích hợp: Telegram Bot (Cá nhân & Group), VietQR
 */

// --- CẤU HÌNH HỆ THỐNG NHQT ---
const BOT_TOKEN = "8531326834:AAEv9UP4ZbQmnETghBywAISIysyPRZdmSdQ"; 

// Đây là ID nhóm Quản trị của bạn (Nhận TOÀN BỘ thông báo)
const ADMIN_GROUP_ID = "-1003806745108";

// --- 2. HÀM KÍCH HOẠT CHÍNH (ON EDIT) ---
function installableOnEdit(e) {
  if (!e || !e.range) return; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value;
  
  if (row <= 1) return;
function installableOnEdit(e) {
  if (!e || !e.range) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  const value = e.value;

  if (row <= 1) return;

  // --- LOGIC MỚI: TỰ ĐỘNG FILL BILL (KHI NHẬP MÃ ĐƠN VÀO Ô B2) ---
  if (sheetName === "Fill" && row === 2 && col === 2) {
    if (!value) return clearBillTemplate(sheet); // Xóa trắng bill nếu xóa mã đơn
    fillBillData(ss, sheet, value);
    return;
  }

  // --- CÁC LOGIC CŨ (CT, RULE, LOG...) ---
  if (sheetName === "CT" || sheetName === "Công thức") {
    // ... Giữ nguyên phần code sinh mã đơn và gửi Telegram bạn đang có ...
    // Đảm bảo bọc writeAutoLog trong try-catch như tôi đã hướng dẫn để tránh treo mã đơn.
    try {
      writeAutoLog(ss, sheetName, range, e.oldValue || "", value || "", "");
    } catch(err) { console.log("Lỗi log: " + err); }
  }

  // XỬ LÝ SHEET "RULE" (NHẬP -> LƯU -> XÓA)
  if (sheetName === "Rule" && row === 2 && col === 5 && value) {
    handleRuleInput(ss, sheet, value);
  }
}
  // XỬ LÝ SHEET "CT" HOẶC "CÔNG THỨC"
  if (sheetName === "CT" || sheetName === "Công thức") {
    const staff = sheet.getRange(row, 13).getValue(); // M: Nhân viên
    const date = sheet.getRange(row, 1).getValue();   // A: Thời gian

    // --- 4. XỬ LÝ SHEET "CT" (MÃ ĐƠN & TELEGRAM) ---
  if (sheetName === "CT" || sheetName === "Công thức") {
    const staffLetter = sheet.getRange(row, 13).getValue(); // Cột M: Nhân viên
    const dateValue = sheet.getRange(row, 1).getValue();   // Cột A: Thời gian
    
    // Tự sinh mã đơn khi cột Thời gian (A) hoặc Nhân viên (M) được điền
    if ((col <= 6 || col === 13) && dateValue instanceof Date && staffLetter) {
      const orderID = generateOrderID(sheet, dateValue, row, staffLetter);
      sheet.getRange(row, 2).setValue(orderID);  // B: Mã đơn
      sheet.getRange(row, 12).setValue(orderID); // L: Nội dung CK
      
      const productValue = sheet.getRange(row, 3).getValue();
      if (productValue) {
        const avgRate = getAvgRateFromRule(ss, dateValue, productValue);
        sheet.getRange(row, 14).setValue(avgRate); // N: Tỷ giá mua từ Rule
      }
    }

    // B. GHI NHẬT KÝ (AUTOLOG - PHÒNG GIAN LẬN)
    writeAutoLog(ss, sheetName, range, e.oldValue || "", value || "", staff);

    // C. GỬI TELEGRAM (CỘT P)
    if (col === 16 && (value === "Gửi công nợ" || value === "Gửi bill CK")) {
      handleTelegramAction(ss, sheet, row, value);
      range.clearContent();
    }

// Chuyển sang OUT nếu cột O chọn Yes
if (col === 15 && range.getValue() === "Yes") { // Cột O
    const outSheet = SpreadsheetApp.getActive().getSheetByName("OUT");
    if (!outSheet) {
      SpreadsheetApp.getActive().toast("Lỗi: Không tìm thấy sheet OUT");
      return;
    }
    const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues();
    outSheet.appendRow(rowData[0]);
    sheet.deleteRow(row);
    SpreadsheetApp.getActive().toast("Đã chuyển dữ liệu sang sheet OUT");
  }
}

// --- 3. CÁC HÀM XỬ LÝ CHI TIẾT ---

function handleRuleInput(ss, sheet, staffName) {
  const input = sheet.getRange("A2:F2").getValues()[0];
  if (!input[0] || !input[1] || !input[2] || !input[5]) return ss.toast("⚠️ Thiếu dữ liệu nhập!");
  
  let historyG = sheet.getRange("G:G").getValues();
  let targetRow = 2;
  for (let i = historyG.length - 1; i >= 0; i--) {
    if (historyG[i][0] !== "") { targetRow = i + 2; break; }
  }

  sheet.getRange("A2:D2").copyTo(sheet.getRange(targetRow, 7), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange("F2").copyTo(sheet.getRange(targetRow, 13), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  sheet.getRange(targetRow, 11).setValue(staffName);
  sheet.getRange("E2").copyTo(sheet.getRange(targetRow, 11), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.getRange(targetRow, 14).setValue(Number(input[2]) * Number(input[5]));
  
  updateAverageRateRule(sheet, input[0], input[1]);
  sheet.getRange("A2:F2").clearContent();
  sheet.getRange("E2").clearContent();
  ss.toast("✅ Đã lưu Nhật ký Rule!");
}
function moveDataToOutSheet(ss, ctSheet, row) {
  const outSheet = ss.getSheetByName("OUT");
  if (!outSheet) return ss.toast("Lỗi: Không tìm thấy sheet OUT!");
  const data = ctSheet.getRange(row, 1, 1, 14).getValues()[0];
  outSheet.appendRow(data);
  ctSheet.getRange(row, 1, 1, 15).clearContent();
  ss.toast("Đã lưu đơn vào OUT và xóa ở CT");
}

function updateFillSheetMapping(ss, ctSheet, ctRow) {
  const fillSheet = ss.getSheetByName("Fill");
  if (!fillSheet) return;
  const d = ctSheet.getRange(ctRow, 1, 1, 14).getValues()[0];
  if (!d[1]) return; // Không có mã đơn thì bỏ qua

  fillSheet.getRange("C1").setValue(d[1]);  fillSheet.getRange("C2").setValue(d[3]);
  fillSheet.getRange("C3").setValue(d[0]);  fillSheet.getRange("C6").setValue(d[2]);
  fillSheet.getRange("D6").setValue(d[4]);  fillSheet.getRange("E6").setValue(d[5]);
  fillSheet.getRange("F6").setValue(d[6]);  fillSheet.getRange("C16").setValue(d[7]);
  fillSheet.getRange("C17").setValue(d[8]); fillSheet.getRange("C18").setValue(d[9]);

  const bankRaw = d[10].toString();
  if (bankRaw.includes('|')) {
    const b = bankRaw.split('|');
    fillSheet.getRange("C20").setValue(b[1]); fillSheet.getRange("C21").setValue(b[2]);
    const qrUrl = `https://img.vietqr.io/image/${b[0]}-${b[1]}-compact2.png?amount=${d[9]}&addInfo=${encodeURIComponent(d[1])}&accountName=${encodeURIComponent(b[2] || "")}`;
    fillSheet.getRange("D16").setFormula(`=IMAGE("${qrUrl}")`);
  }
}
// --- HÀM GỬI TELEGRAM (LOGIC CHÍNH - PHỤ / BẢN 2 CỘT) ---
function handleTelegramAction(ss, sheet, row, type) {
  const d = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const customerName = d[3]; // Cột D: Tên khách hàng
  const customerGroupId = getCustomerGroupId(customerName); // Lấy ID từ sheet Định nghĩa
  const amountClean = Math.round(String(d[9]).replace(/[^0-9]/g, ""));

  let msg = `<b>NHQT gửi thông tin thanh toán cho đơn hàng:</b> ${d[1]}\n` +

            `<b>Thông tin sản phẩm:</b>\n` +
            `${d[2]}: ${d[4]} x ${Number(d[5]).toLocaleString()} = ${Number(d[6]).toLocaleString()} đ\n` +
            `-----------------\n` +
            `<b>Tổng tiền:</b> ${Number(d[7]).toLocaleString()} đ\n` +
            `<b>Đã thanh toán:</b> ${Number(d[8]).toLocaleString()} đ\n` +
            `<b>Số tiền cần thanh toán:</b> <u>${Number(d[9]).toLocaleString()} đ</u>`;

   // Danh sách nhận tin: Luôn có Admin + Khách (nếu ID khách khác ID Admin)
  let targets = [ADMIN_GROUP_ID];
  if (customerGroupId && customerGroupId !== ADMIN_GROUP_ID) {
    targets.push(customerGroupId);
  }

  if (type === "Gửi bill CK") {
    const b = d[10].toString().split('|'); // Bank|STK|Tên
    if (b.length < 2) return ss.toast("⚠️ Cột K sai định dạng Bank|STK|Tên");
    
    // Tạo mã QR VietQR
    const qr = `https://img.vietqr.io/image/${b[0]}-${b[1]}-compact2.png?amount=${amountClean}&addInfo=${encodeURIComponent(d[1])}`;

    // GỬI 1 TIN DUY NHẤT: ẢNH + CAPTION cho tất cả
    targets.forEach(id => {
      sendTelegram(id, qr, msg, true);
    });

  } else if (type === "Gửi công nợ") {
    // GỬI 1 TIN DUY NHẤT: CHỈ VĂN BẢN cho tất cả
    targets.forEach(id => {
      sendTelegram(id, null, msg, false);
    });
  }
}
    // Gửi công nợ (Chỉ văn bản)
    // 1. Gửi Admin
    sendTelegram(ADMIN_GROUP_ID, null, `\n${msg}`, false);

    // 2. Gửi nhóm khách
    if (customerGroupId) {
      sendTelegram(customerGroupId, null, msg, false);
    } else {
      ss.toast("⚠️ Không tìm thấy ID Nhóm cho khách: " + customerName);
    }

// --- 4. HÀM PHỤ TRỢ (HELPER FUNCTIONS) ---


function sendTelegram(chatId, photo, text, isPhoto) {
  const method = isPhoto ? "sendPhoto" : "sendMessage";
  const url = "https://api.telegram.org/bot" + BOT_TOKEN + "/" + method;
  const payload = { chat_id: chatId, parse_mode: "HTML" };
  if (isPhoto) { payload.photo = photo; payload.caption = text; } else { payload.text = text; }
  UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", payload: JSON.stringify(payload), muteHttpExceptions: true });
}

function generateOrderID(sheet, date, currentRow, staff) {
  const ds = Utilities.formatDate(new Date(date), "GMT+7", "yyyyMMdd");
  const mm = Utilities.formatDate(new Date(date), "GMT+7", "MM"), dd = Utilities.formatDate(new Date(date), "GMT+7", "dd"), yy = Utilities.formatDate(new Date(date), "GMT+7", "yy");
  const allDates = sheet.getRange(1, 1, currentRow).getValues();
  let count = 0;
  for (let i = 0; i < currentRow - 1; i++) {
    if (allDates[i][0] instanceof Date && Utilities.formatDate(allDates[i][0], "GMT+7", "yyyyMMdd") === ds) count++;
  }
  count++; 
  return staff + mm + dd + (count < 10 ? "0" + count : count) + yy;
}

function getAvgRateFromRule(ss, date, product) {
  const ruleSheet = ss.getSheetByName("Rule");
  if (!ruleSheet) return "";
  const target = Utilities.formatDate(new Date(date), "GMT+7", "yyyyMMdd");
  const data = ruleSheet.getRange("G:L").getValues();
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] instanceof Date && data[i][1] === product && Utilities.formatDate(new Date(data[i][0]), "GMT+7", "yyyyMMdd") === target) return data[i][5];
  }
  return "";
}

function updateAverageRateRule(sheet, date, product) {
  const target = Utilities.formatDate(new Date(date), "GMT+7", "yyyyMMdd");
  const data = sheet.getRange("G:I").getValues();
  let total = 0, count = 0, rows = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] instanceof Date && data[i][1] === product && Utilities.formatDate(data[i][0], "GMT+7", "yyyyMMdd") === target) {
      total += Number(data[i][2]); count++; rows.push(i + 1);
    }
  }
  if (count > 0) { const avg = total / count; rows.forEach(r => sheet.getRange(r, 12).setValue(avg)); }
}

function moveDataToOutSheet(ss, ctSheet, row) {
  const out = ss.getSheetByName("OUT");
  const data = ctSheet.getRange(row, 1, 1, 14).getValues()[0];
  const colA = out.getRange("A:A").getValues();
  let lastR = 1;
  for (let i = colA.length - 1; i >= 0; i--) { if (colA[i][0] !== "") { lastR = i + 1; break; } }
  out.getRange(lastR + 1, 1, 1, 14).setValues([data]);
  if (out.getLastRow() > 1) out.getRange(2, 1, out.getLastRow() - 1, 14).sort({column: 1, ascending: true});
  ctSheet.getRange(row, 1, 1, 15).clearContent();
}

function writeAutoLog(ss, sheetName, range, oldV, newV, staff) {
  if (sheetName === "AutoLog") return;
  let log = ss.getSheetByName("AutoLog") || ss.insertSheet("AutoLog");
  if (log.getLastRow() === 0) log.appendRow(["Thời gian", "Sheet", "Ô", "Cũ", "Mới", "Nhân viên"]);
  log.appendRow([new Date(), sheetName, range.getA1Notation(), oldV, newV, staff || "Hệ thống"]);
}
// --- HÀM TRA CỨU ID NHÓM KHÁCH (BẢN 2 CỘT: A=Tên, B=ID Nhóm) ---
function getCustomerGroupId(name) {
  if (!name) return null;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Định nghĩa");
  if (!sheet) return null;
  
  // Lấy dữ liệu cột A và B (Cột 1 và 2)
  const data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  
  for (let i = 0; i < data.length; i++) {
    // So sánh tên khách (Cột A - data[i][0])
    if (data[i][0].toString().trim().toLowerCase() === name.toString().trim().toLowerCase()) {
      // Trả về ID Nhóm (Cột B - data[i][1])
      const groupId = data[i][1].toString().trim();
      // Kiểm tra ID có hợp lệ (bắt đầu bằng dấu trừ)
      return groupId.startsWith("-") ? groupId : null; 
    }
  }
  return null;
}
function testCheckCustomer() {
  var testName = "An Diêu"; // <--- THAY TÊN KHÁCH BẠN ĐANG TEST VÀO ĐÂY
  var resultId = getCustomerGroupId(testName);
  
  if (resultId) {
    Logger.log("✅ Tìm thấy ID cho khách [" + testName + "]: " + resultId);
    if (!resultId.startsWith("-")) {
      Logger.log("❌ LỖI: ID nhóm phải bắt đầu bằng dấu trừ '-'");
    }
  } else {
    Logger.log("❌ LỖI: Không tìm thấy tên [" + testName + "] trong sheet Định nghĩa. Hãy kiểm tra dấu cách hoặc chính tả.");
  }
}
// Hàm ghi Nhật ký sửa đổi (Sửa lỗi ReferenceError)
function writeAutoLog(ss, sheetName, range, oldV, newV, staff) {
  if (sheetName === "AutoLog") return;
  let log = ss.getSheetByName("AutoLog") || ss.insertSheet("AutoLog");
  if (log.getLastRow() === 0) log.appendRow(["Thời gian", "Sheet", "Ô", "Cũ", "Mới", "Nhân viên"]);
  log.appendRow([new Date(), sheetName, range.getA1Notation(), oldV, newV, staff || "Hệ thống"]);
}
// --- HÀM SINH MÃ ĐƠN TỰ ĐỘNG ---
function generateOrderID(sheet, date, currentRow, staff) {
  const ds = Utilities.formatDate(new Date(date), "GMT+7", "yyyyMMdd");
  const mm = Utilities.formatDate(new Date(date), "GMT+7", "MM");
  const dd = Utilities.formatDate(new Date(date), "GMT+7", "dd");
  const yy = Utilities.formatDate(new Date(date), "GMT+7", "yy");
  const allDates = sheet.getRange(1, 1, sheet.getLastRow()).getValues();
  let count = 0;
  for (let i = 0; i < currentRow; i++) {
    if (allDates[i][0] instanceof Date && Utilities.formatDate(new Date(allDates[i][0]), "GMT+7", "yyyyMMdd") === ds) {
      count++;
    }
  }
  const stt = count < 10 ? "0" + count : count.toString();
  return staff + mm + dd + stt + yy;
    }
  }
  // Hàm tìm và điền dữ liệu vào Bill
function fillBillData(ss, billSheet, orderId) {
  clearBillTemplate(billSheet);

  const sources = ["CT", "OUT"];
  let items = [];

  sources.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    data.forEach(r => {
      if (r[1] === orderId) items.push(r); // Cột B = Mã đơn
    });
  });

  if (items.length === 0) {
    ss.toast("❌ Không tìm thấy mã đơn");
    return;
  }

  // ===== HEADER =====
  billSheet.getRange("B1").setValue(items[0][1]); // Mã đơn
  billSheet.getRange("B2").setValue(items[0][3]); // Khách hàng
  billSheet.getRange("B3").setValue(items[0][0]); // Thời gian

  // ===== BẢNG SẢN PHẨM =====
  let total = 0;
  let paid = 0;

  items.forEach((r, i) => {
    const row = 6 + i;
    billSheet.getRange(row, 2).setValue(i + 1); // STT
    billSheet.getRange(row, 3).setValue(r[2]); // Sản phẩm
    billSheet.getRange(row, 4).setValue(r[4]); // Số tiền CNY
    billSheet.getRange(row, 5).setValue(r[5]); // Tỷ giá
    billSheet.getRange(row, 6).setValue(r[6]); // Thành tiền VND

    total += Number(r[6]) || 0;
    paid += Number(r[8]) || 0;
  });

  // ===== TỔNG HỢP =====
  billSheet.getRange("B16").setValue(total);
  billSheet.getRange("B17").setValue(paid);
  billSheet.getRange("B18").setValue(total - paid);

  // ===== THANH TOÁN =====
  billSheet.getRange("C19").setValue("Chuyển khoản");
  billSheet.getRange("B20").setValue(items[0][10]); // STK ngân hàng
  billSheet.getRange("B21").setValue("LE THI THUY"); // Chủ TK (có thể map cột nếu bạn thêm)
  }
}
