// --- HÀM GỬI TELEGRAM CHUẨN (Fix lỗi msg) ---
function handleTelegramAction(ss, sheet, row, type) {
  const d = sheet.getRange(row, 1, 1, 14).getValues()[0];
  const customerName = d[3]; 
  const customerGroupId = getCustomerGroupId(customerName);
  const amountClean = Math.round(String(d[9]).replace(/[^0-9]/g, ""));

  // KHAI BÁO BIẾN MSG Ở ĐÂY
let msg = `<b>NHQT gửi thông tin thanh toán cho đơn hàng:</b> ${d[1]}\n` +

            `<b>Thông tin sản phẩm:</b>\n` +
            `${d[2]}: ${d[4]} x ${Number(d[5]).toLocaleString()} = ${Number(d[6]).toLocaleString()} đ\n` +
            `-----------------\n` +
            `<b>Tổng tiền:</b> ${Number(d[7]).toLocaleString()} đ\n` +
            `<b>Đã thanh toán:</b> ${Number(d[8]).toLocaleString()} đ\n` +
            `<b>Số tiền cần thanh toán:</b> <u>${Num

  let targets = [ADMIN_GROUP_ID]; // Luôn gửi Admin
  if (customerGroupId && customerGroupId !== ADMIN_GROUP_ID) targets.push(customerGroupId);

  if (type === "Gửi bill CK") {
    const b = d[10].toString().split('|');
    const qr = `https://img.vietqr.io/image/${b[0]}-${b[1]}-compact2.png?amount=${amountClean}&addInfo=${encodeURIComponent(d[1])}`;
    targets.forEach(id => sendTelegram(id, qr, msg, true));
  } else {
    targets.forEach(id => sendTelegram(id, null, msg, false));
  }
}

// --- HÀM TỰ ĐIỀN DỮ LIỆU VÀO BILL ---
function fillBillData(ss, billSheet, orderId) {
  const ctData = ss.getSheetByName("CT").getDataRange().getValues();
  const outData = ss.getSheetByName("OUT").getDataRange().getValues();
  const allData = ctData.concat(outData);
  
  let found = false;
  for (let i = 0; i < allData.length; i++) {
    if (allData[i][1] == orderId) {
      const d = allData[i];
      billSheet.getRange("B3").setValue(d[3]); // Khách (D)
      billSheet.getRange("B4").setValue(d[0]); // Ngày (A)
      billSheet.getRange("C6").setValue(d[2]); // SP (C)
      billSheet.getRange("F6").setValue(d[6]); // Tiền (G)
      billSheet.getRange("B18").setValue(d[9]); // Còn lại (J)
      found = true; break;
    }
  }
  if (!found) ss.toast("❌ Không tìm thấy mã đơn!");
}
