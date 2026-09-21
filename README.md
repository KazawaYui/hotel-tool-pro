# Hotel Tool Pro

Công cụ lễ tân Tân Hotel: biến dữ liệu khách hằng ngày thành hồ sơ nộp công an
(KBTT cho khách nước ngoài, VNM cho khách Việt, sổ ĐK14), kèm quy đổi tỷ giá,
Regcard/ARR, đối chiếu lưu trú và sổ giao ca.

## Chạy

```bash
pip install -r requirements.txt
streamlit run app.py
```

Sổ giao ca lưu trên Supabase nếu có cấu hình; không có thì lưu tạm trên đĩa
server theo ngày. Xem `secrets.toml.example`.

## Cấu trúc

`app.py` chỉ chứa giao diện. Toàn bộ logic nghiệp vụ nằm trong `hotel/`, xếp
thành 4 tầng không có phụ thuộc vòng:

```
common                            giờ VN, đọc ô Excel, chuẩn hoá phòng/tên/ngày
  ├── lunar        âm lịch, mùa và khung giờ cho màn chào
  ├── progress     tiến độ ca, lưu trên đĩa để sống sót qua tải lại trang
  ├── db           sổ giao ca trên Supabase
  ├── assets       mẫu file .b64 và ảnh nền
  └── lookups      bảng tra mã quốc tịch + địa danh hành chính
        ├── reconcile   đối chiếu Smile ↔ trang lưu trú
        ├── regcard     Regcard PDF + file ARR
        ├── xlsx_io     quy đổi tỷ giá, chia giá phòng connecting
        ├── dk14        sổ ĐK14
        ├── docs        hồ sơ KBTT + VNM
        └── reports     kiểm tra dữ liệu, báo cáo ngày
```

`hotel_core.py` gộp tất cả lại; `app.py` và bộ test chỉ cần
`from hotel_core import *`.

Mỗi màn công cụ trong `app.py` được bọc trong một `@st.fragment` riêng, nên
thao tác bên trong một màn chỉ vẽ lại màn đó thay vì cả trang.

## Test

```bash
python -m pytest          # 342 test, ~5 giây
```

CI chạy bộ test này trên mỗi push và pull request, kèm 3 bước kiểm tra app
dựng được, `config.toml` hợp lệ, và ảnh nền còn đủ file.

## Những cái bẫy đã gặp (đừng vấp lại)

Mỗi mục dưới đây đều là lỗi **đã xảy ra thật trên bản đang chạy**, đã sửa và
có test chặn. Ghi lại vì chúng đều thuộc loại sai mà không báo lỗi gì.

**Chữ Đ không phải chữ D có dấu.** `Đ`/`đ` là ký tự riêng (U+0110/U+0111) nên
`unicodedata.normalize('NFD', …)` không tách ra được; lọc `[^a-z0-9]` sau đó sẽ
**xoá hẳn** chữ đ. `"Đà Nẵng"` từng thành `"a nang"`, `"Ấn Độ"` thành `"ano"`.
Tệ nhất: `"Đức"` thành `"uc"` và chiếm mất chỗ của `"Úc"` — khách Úc bị khai
thành khách Đức trên hồ sơ công an. Dùng `_strip_accents()` trong `common`.

**Quy đổi tỷ giá TRƯỚC, chia giá connecting SAU.** Chia trước thì giá ngoại tệ
(278.40 EUR) bị `int()` cắt còn 278, chia 139/139 rồi quy đổi từng nửa → lệch
gần 12.000đ so với PMS. Thứ tự đúng nằm trong khối xử lý hằng ngày của `app.py`.

**Bảng tra phải tra được bằng chính tên hiển thị của nó.** PMS xuất quốc tịch
bằng tên tiếng Anh đầy đủ (`"United States of America"`) còn bảng chỉ có tên
Việt và vài alias rút gọn → 10 nước ra hồ sơ trống mã. `lookups` nay tự sinh
khoá từ tên hiển thị; có test bất biến chặn tái phát.

**Danh sách cặp phòng connecting có 13 tầng.** Nguồn `C_p_CNT.xlsx` có 14 cột;
hai lần đọc trước đều bị cắt (9 rồi 11 cột) làm mất tầng 14→17. Nay đủ 78 cặp,
có test đếm.

**Đừng `copy()` style từng ô.** openpyxl chia sẻ style qua bảng chung của
workbook. Chụp một lần bằng `snapshot_styles()` rồi dán bằng `apply_style()` —
nhanh gấp 1,9 lần. Nhưng không bỏ `copy()` được ở mọi chỗ: `cell.font` trả về
`StyleProxy` không hash được, phải `copy()` để gỡ proxy.

**Ảnh nền để trong `static/`, không nhúng base64 vào CSS.** Inline CSS không bao
giờ được trình duyệt cache nên ảnh nhúng phải tải lại mỗi lần mở trang. Dùng
đường dẫn **tương đối** `app/static/…` để còn đúng khi deploy dưới đường dẫn con.

**`.streamlit/config.toml` hỏng thì Streamlit im lặng bỏ qua TẤT CẢ.** Thêm một
section `[server]` thứ hai là đủ làm hỏng cú pháp TOML, và cả bảng màu theme
cũng biến mất mà không có thông báo nào. CI có bước kiểm tra file này.

## Khoảng trống dữ liệu còn tồn tại

Bảng quốc tịch KBTT **không có** Campuchia, Lào và Nam Phi. App trả nguyên văn
kèm cảnh báo để lễ tân tự điền — không tự đoán mã. Cần bổ sung thì lấy mã chuẩn
từ danh mục của công an rồi thêm vào `hotel/lookups.py`.
