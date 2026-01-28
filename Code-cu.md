Có sang sheet OUT nhưng không xoá ở CT
----
Code cũ

/**
 * HỆ THỐNG QUẢN LÝ TỰ ĐỘNG: Rule, CT, OUT, BC, Fill & AutoLog
 * Cập nhật: 16/01/2026 - Fix Dropdown & AutoLog chống gian lận
 */

function installableOnEdit(e) {
  if (!e || !e.range) return; 

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  if (row <= 1) return; // Bỏ qua dòng tiêu đề

  // --- 1. XÁC ĐỊNH NHÂN SỰ TỪ DROPDOWN ---
  let staffLetter = "";
  if (sheetName === "Rule") {
    staffLetter = sheet.getRange(row, 5).getValue(); // Cột E: Người update
  } else if (sheetName === "CT" || sheetName === "Công thức") {
    staffLetter = sheet.getRange(row, 13).getValue(); // Cột M: Nhân viên
  }

  // --- 2. GHI NHẬT KÝ (AUTOLOG) ---
  // Ghi lại ký hiệu nhân sự và nội dung sửa đổi để tránh gian lận
  writeAutoLog(ss, sheetName, range, e.oldValue || "", e.value || "", staffLetter);

  // --- 3. XỬ LÝ SHEET "RULE" (NHẬT KÝ NẠP TIỀN) ---
  if (sheetName === "Rule" && col <= 6) {
    if (!staffLetter) return; // Chỉ chạy khi đã chọn người update ở cột E

    const inputData = sheet.getRange(row, 1, 1, 6).getValues()[0];
    const historyColG = sheet.getRange("G:G").getValues();
    let nextRow = 2;
    for (let i = historyColG.length - 1; i >= 0; i--) {
      if (historyColG[i][0] !== "") { nextRow = i + 2; break; }
    }
    
    // Sao chép sang vùng Nhật ký G-M
    sheet.getRange(nextRow, 7, 1, 4).setValues([[inputData[0], inputData[1], inputData[2], inputData[3]]]);
    sheet.getRange(nextRow, 11).setValue(staffLetter); // K: Người update
    sheet.getRange(nextRow, 13).setValue(inputData[5]); // M: Số lượng nạp CNY
    
    updateAverageRateRule(sheet, inputData[0], inputData[1]);
  }

  // --- 4. XỬ LÝ SHEET "CT" (NHẬP ĐƠN) ---
  if (sheetName === "CT" || sheetName === "Công thức") {
    if (!staffLetter) return; // Chỉ chạy khi đã chọn nhân viên ở cột M
    
    const dateValue = sheet.getRange(row, 1).getValue(); // A: Thời gian
    const productValue = sheet.getRange(row, 3).getValue(); // C: Sản phẩm
    
    if (dateValue instanceof Date) {
      // Tự sinh mã đơn (B) và Map nội dung CK (L) dựa trên nhân sự dropdown
      const mm = Utilities.formatDate(dateValue, "GMT+7", "MM");
      const dd = Utilities.formatDate(dateValue, "GMT+7", "dd");
      const yy = Utilities.formatDate(dateValue, "GMT+7", "yy");
      const stt = getSequenceInDay(sheet, dateValue, row);
      
      const orderID = staffLetter + mm + dd + stt + yy; //
      sheet.getRange(row, 2).setValue(orderID);  
      sheet.getRange(row, 12).setValue(orderID); 
      
      if (productValue) {
        const avgRate = getAvgRateFromRule(ss, dateValue, productValue);
        sheet.getRange(row, 14).setValue(avgRate); // N: Tỷ giá mua
      }
    }

    // Chuyển sang OUT nếu chọn "Yes" ở cột O
    if (col === 15 && e.value === "Yes") {
      moveDataToOutSheet(ss, sheet, row);
    } else {
      updateFillSheetMapping(ss, sheet, row);
    }
  }
}

// --- CÁC HÀM HỖ TRỢ (Hàm con) ---

function writeAutoLog(ss, sheetName, range, oldV, newV, staff) {
  if (sheetName === "AutoLog") return;
  let log = ss.getSheetByName("AutoLog") || ss.insertSheet("AutoLog");
  if (log.getLastRow() === 0) {
    log.appendRow(["Thời gian", "Sheet", "Ô", "Cũ", "Mới", "Nhân viên thực hiện"]);
  }
  log.appendRow([new Date(), sheetName, range.getA1Notation(), oldV, newV, staff || "Chưa chọn NV"]);
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

function updateAverageRateRule(sheet, date, product) {
  const target = Utilities.formatDate(date, "GMT+7", "yyyyMMdd");
  const data = sheet.getRange("G:I").getValues();
  let total = 0, count = 0, rows = [];
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] instanceof Date && data[i][1] === product && Utilities.formatDate(data[i][0], "GMT+7", "yyyyMMdd") === target) {
      total += Number(data[i][2]); count++; rows.push(i + 1);
    }
  }
  if (count > 0) rows.forEach(r => sheet.getRange(r, 12).setValue(total / count));
}

function getAvgRateFromRule(ss, date, product) {
  const rule = ss.getSheetByName("Rule");
  const data = rule.getRange("G:L").getValues();
  const target = Utilities.formatDate(date, "GMT+7", "yyyyMMdd");
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] instanceof Date && data[i][1] === product && Utilities.formatDate(data[i][0], "GMT+7", "yyyyMMdd") === target) return data[i][5];
  }
  return "";
}

function getSequenceInDay(sheet, date, row) {
  const ds = Utilities.formatDate(date, "GMT+7", "yyyyMMdd");
  const allDates = sheet.getRange(1, 1, row).getValues();
  let count = 0;
  for (let i = 0; i < allDates.length; i++) {
    if (allDates[i][0] instanceof Date && Utilities.formatDate(allDates[i][0], "GMT+7", "yyyyMMdd") === ds) count++;
  }
  return count < 10 ? "0" + count : count.toString();
}
