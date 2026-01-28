Dưới đây là bảng thống kê chi tiết các sheet trong hệ thống **NHQT Payment System** của bạn, được trình bày theo cấu trúc **User Story (US)** và **Acceptance Criteria (AC)** để phục vụ quá trình "làm mịn sản phẩm" (Product Refinement):

### 1. Sheet: CT (Công thức)

* **User Story:** Là một **Nhân viên**, tôi muốn nhập liệu thông tin đơn hàng tại đây, để hệ thống tự động sinh mã định danh và hỗ trợ gửi thông báo thanh toán cho khách hàng.
* **Acceptance Criteria (AC):**
* **AC1:** Tự động sinh **Mã đơn (Cột B)** và **ND CK (Cột L)** ngay khi điền Ngày (Cột A) và chọn Nhân viên (Cột M).
* **AC2:** Mã đơn phải tuân thủ định dạng: `Nhân viên + Tháng + Ngày + STT + Năm`.
* **AC3:** Khi chọn "Gửi bill CK" hoặc "Gửi công nợ" tại **Cột P**, hệ thống phải gửi tin nhắn đến đúng nhóm Telegram của khách và nhóm Admin.
* **AC4:** Khi chọn "Yes" tại **Cột O**, toàn bộ dòng dữ liệu phải được chuyển tự động sang sheet **OUT** và xóa tại sheet **CT**.



### 2. Sheet: Fill (Bill)

* **User Story:** Là một **Nhân viên**, tôi muốn nhập mã đơn vào một ô duy nhất để hệ thống tự động đổ toàn bộ dữ liệu vào mẫu in, giúp tiết kiệm thời gian tạo hóa đơn.
* **Acceptance Criteria (AC):**
* **AC1:** Hệ thống tự động kích hoạt khi nhập Mã đơn vào ô **B2**.
* **AC2:** Dữ liệu phải được truy xuất từ cả hai nguồn là sheet **CT** và sheet **OUT**.
* **AC3:** Nếu một mã đơn có nhiều dòng sản phẩm, hệ thống phải liệt kê đầy đủ tất cả các dòng vào bảng hàng hóa (Dòng 6-15).
* **AC4:** Phần tài chính (Tổng thu, Đã TT, Còn lại) phải được cộng dồn chính xác từ tất cả các dòng có chung mã đơn.
* **AC5:** Tự động xóa trắng các ô dữ liệu khi ô B2 bị xóa trống.



### 3. Sheet: Định nghĩa

* **User Story:** Là một **Quản trị viên**, tôi muốn thiết lập bản đồ định danh khách hàng, để Bot biết chính xác địa chỉ (ID nhóm) cần gửi tin nhắn cho từng khách.
* **Acceptance Criteria (AC):**
* **AC1:** **Cột A (Tên):** Phải chứa tên khách hàng khớp 100% với tên tại cột D của sheet CT.
* **AC2:** **Cột B (ID Nhóm):** Phải chứa ID Telegram hợp lệ (thường bắt đầu bằng dấu `-` hoặc `-100`).
* **AC3:** Mỗi khách hàng chỉ được tồn tại một dòng duy nhất để tránh xung đột dữ liệu gửi Bot.



### 4. Sheet: OUT

* **User Story:** Là một **Quản trị viên**, tôi muốn có một kho lưu trữ các đơn hàng đã hoàn thành, để giữ cho sheet làm việc (CT) luôn gọn gàng nhưng vẫn có thể tra cứu lại khi cần.
* **Acceptance Criteria (AC):**
* **AC1:** Dữ liệu được đẩy sang phải giữ nguyên cấu trúc các cột để phục vụ việc tra cứu ngược từ sheet Fill.
* **AC2:** Chỉ tiếp nhận dữ liệu từ sheet CT khi người dùng xác nhận "Yes".



### 5. Sheet: AutoLog

* **User Story:** Là một **Quản trị viên**, tôi muốn theo dõi mọi thay đổi trên hệ thống, để kiểm soát sai sót và bảo mật dữ liệu.
* **Acceptance Criteria (AC):**
* **AC1:** Tự động ghi lại: Thời gian, Tên Sheet, Tên Ô, Giá trị cũ, Giá trị mới và Nhân viên thực hiện.
* **AC2:** Quá trình ghi Log không được làm gián đoạn các tính năng khác (như sinh mã đơn) nếu gặp lỗi.



### 6. Sheet: Rule

* **User Story:** Là một **Quản trị viên**, tôi muốn quản lý các quy tắc nghiệp vụ như tỷ giá mua/bán, để hệ thống tự động tính toán tài chính mà không cần nhập tay từng dòng.
* **Acceptance Criteria (AC):**
* **AC1:** Cung cấp dữ liệu tỷ giá tham chiếu cho các hàm tính toán tại sheet CT.



Bảng thống kê này giúp bạn có cái nhìn tổng thể để kiểm tra tính sẵn sàng của từng tính năng trước khi đưa vào vận hành thực tế. Bạn có muốn điều chỉnh thêm tiêu chí nào cho từng sheet không?
