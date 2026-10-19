## 1. PHÂN TÍCH BÀI TẬP (I/O)

|Cột cần điền|lookup_value | table_array| col_index_num|
|---|---|---|---|
|D - Tên Tác Giả	| TRIM(C2)	| 'DanhMucTacGia'!$A$2:$E$81|	2|
|E – Phòng trình bày	| TRIM(C2)	| 'DanhMucTacGia'!$A$2:$E$81|	4|
|F – Thù lao| TRIM(C2)| 'DanhMucTacGia'!$A$2:$E$81 |	5|

## 2. Chiến lược xử lý bẫy
Bẫy 1 (Khoảng trắng): Dùng ```TRIM```(C2) để tự động cắt bỏ khoảng trắng thừa của mã tác giả.

Bẫy 2 (Lỗi #N/A): ```Dùng IFERROR(..., "Chưa cập nhật")``` để đổi lỗi thành chữ "Chưa cập nhật" cho 3 tác giả ngoài danh sách.

Bẫy 3 (Trôi vùng dò tìm): Dùng ký tự cố định tuyệt đối $ cho bảng tra cứu ($A$2:$E$81).


## 3. Các bước thực hiện trong Excel
Chọn ô D2, viết hàm VLOOKUP kết hợp ```TRIM```,``` IFERROR ```và dấu ```$```.

Kéo (fill handle) công thức từ D2 xuống D81.

Copy công thức sang E2 và F2, sau đó thay đổi số thứ tự cột ```(Col_index_num)``` thành 4 và 5.

### 4. Công thức ô D2
```
Excel
=IFERROR(VLOOKUP(TRIM(C2), DanhMucTacGia!$A$2:$E$81, 2, 0), "Chưa cập nhật")
```
### 5. Công thức ô E2 & F2

Ô E2 (Phòng):``` =IFERROR(VLOOKUP(TRIM(C2), DanhMucTacGia!$A$2:$E$81, 4, 0), "Chưa cập nhật")```

Ô F2 (Thù lao):``` =IFERROR(VLOOKUP(TRIM(C2), DanhMucTacGia!$A$2:$E$81, 5, 0), "Chưa cập nhật")```

Giải thích: Do chung bảng nguồn và chung giá trị dò tìm, ta chỉ cần thay đổi số thứ tự cột ```(Col_index_num)``` thành 4 và 5 mà không cần viết lại toàn bộ hàm.
### 6. Hiện tượng khi bỏ dấu $
Nếu bỏ ```$```, khi kéo công thức từ D2 xuống D5 (dịch xuống 3 dòng), vùng tra cứu sẽ bị trôi theo thành ```DanhMucTacGia!A5:E84```.

Hậu quả: Bảng tra bị lệch dòng, không tìm thấy mã tác giả và trả về lỗi #N/A hàng loạt từ dòng 3 trở đi.