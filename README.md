## Phần 1. Phân tích: So sánh bảng Excel "Xấu" và "Đúng"

| Tiêu chí | Bảng Excel "Xấu" (Vi phạm quy tắc vàng) | Bảng Excel "Đúng" (Chuẩn chuẩn hóa dữ liệu) |
| --- | --- | --- |
| **Cấu trúc & Trình bày** | Lạm dụng **Merge Cell** (trộn ô) ở các tiêu đề chính hoặc danh mục, gây lỗi khi lọc dữ liệu (Filter) hoặc dùng hàm. | Dữ liệu dạng bảng phẳng (Tabular/Flat table), mỗi cột là một trường dữ liệu (`Field`), mỗi dòng là một bản ghi (`Record`). |
| **Định dạng dữ liệu** | Trộn lẫn kiểu dữ liệu (ô số chứa cả chữ như `100 triệu VNĐ`), không định dạng căn lề nhất quán. | Tách biệt số liệu (`100000000`) và đơn vị; định dạng kiểu tiền tệ (`VND`) bằng format ô để dễ dàng tính toán. |
| **Công thức & Tham chiếu** | Dùng số cứng trong công thức hoặc tham chiếu thủ công dễ sai sót; không cố định địa chỉ ô khi kéo công thức. | Sử dụng các hàm tra cứu tiêu chuẩn (`VLOOKUP`, `XLOOKUP`) với địa chỉ tuyệt định rõ ràng (`$`). |

---

## Phần 2. Triển khai: Hồ sơ dự án Smart-Home

### 1. Tài liệu SRS (Software Requirements Specification) - Mô tả nền tảng

* **Heading 1: Giới thiệu hệ thống**
* *Heading 2: Mục tiêu dự án* — Xây dựng hệ sinh thái nhà thông minh tự động hóa cao, kết nối các thiết bị IoT qua Cloud.
* *Heading 2: Phạm vi ứng dụng* — Quản lý trạng thái thiết bị, tự động bật/tắt cảm biến ánh sáng, nhiệt độ và an ninh từ xa.


* **Heading 1: Yêu cầu chức năng**
* *Heading 2: Quản lý thiết bị* — Thêm, sửa, xóa, giám sát trạng thái thiết bị thời gian thực.
* *Heading 2: Cảnh báo tự động* — Gửi thông báo khi phát hiện rò rỉ khí gas hoặc có chuyển động bất thường qua cảm biến PIR.



### 2. File Excel: Quản lý danh sách nhà đất & Tính toán hoa hồng

* **Bảng Danh Sách Nhà Đất:** Cấu trúc các cột chuẩn gồm: `Mã BĐS`, `Tên Căn Hộ`, `Giá Bán (VNĐ)`, `Loại Hình`, `Khu Vực`, `Diện Tích (m2)`, `Hoa Hồng (VNĐ)`.
* **Bảng Tra Cứu Hoa Hồng (Tham chiếu phụ):**
* Cột 1: `Loại BĐS` (Ví dụ: Căn hộ, Biệt thự, Đất nền).
* Cột 2: `Tỷ lệ hoa hồng` (Ví dụ: `1.5%`, `2%`, `2.5%`).


* **Công thức tính:** Sử dụng hàm `VLOOKUP` kết hợp định địa chỉ tuyệt đối để dò tìm tỷ lệ hoa hồng chính xác từ bảng phụ, nhân với giá bán để ra số tiền hoa hồng thực nhận.

### 3. Bộ Slide Pitch Deck gọi vốn (Quy tắc 10/20/30)

* **Slide 1:** Tiêu đề dự án **Smart-Home** – Giải pháp không gian sống thông minh. *(1 thông điệp chính: Hiện thực hóa cuộc sống hiện đại).*
* **Slide 2:** Vấn đề thị trường – Nhu cầu an ninh và tiết kiệm năng lượng gia tăng.
* **Slide 3:** Giải pháp cốt lõi – Hệ sinh thái tích hợp cảm biến và tự động hóa qua di động.
* **Slide 4:** Mô hình kinh doanh & Doanh thu từ dịch vụ môi giới/thiết bị.
* **Slide 5:** Kế hoạch tài chính & Lời kêu gọi đầu tư (Call to Action).

---

## Phần 3. Kiểm thử: Khắc phục lỗi trôi vùng dữ liệu khi kéo công thức VLOOKUP

* **Vấn đề:** Khi bạn viết công thức `VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])` và kéo công thức xuống các dòng dưới, vùng tham chiếu `table_array` (bảng phụ chứa dữ liệu tra cứu) sẽ bị tự động dịch chuyển dòng (ví dụ từ `B2:C10` thành `B3:C11`, rồi `B4:C12`...). Điều này khiến các dòng phía dưới bị lỗi `#N/A` do bỏ sót dữ liệu ở đầu bảng.
* **Giải pháp sử dụng dấu đô la ($):**
* Cần cố định tuyệt đối tọa độ của bảng tra cứu bằng cách thêm dấu `$` vào trước tên cột và số dòng.
* *Ví dụ thay vì:* `VLOOKUP(E2, B2:C10, 2, FALSE)`
* *Hãy sửa thành:* `VLOOKUP(E2, $B$2:$C$10, 2, FALSE)`
* Khi có dấu `$`, dù bạn kéo công thức đi đâu, vùng tra cứu `$B$2:$C$10` vẫn giữ nguyên cố định, giúp hàm luôn tìm kiếm chính xác mà không bị trôi vùng dữ liệu.
