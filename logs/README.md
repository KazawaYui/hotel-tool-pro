# Nhật ký hoạt động ca (logs/)

Thư mục này chứa các file `shift_YYYY-MM-DD_HHMM.json` được lễ tân xuất từ màn
hình **Tổng quan ca trực** (nút "Xuất nhật ký hoạt động") và commit thủ công
vào đây sau mỗi ca.

## Nội dung mỗi file

Chỉ gồm **số liệu thống kê tổng hợp** và **trình tự thao tác** trong ca:
- `nav_sequence`: trình tự mở các công cụ trong ca kèm giờ (không có dữ liệu khách)
- `daily_processing` / `regcard_arr` / `recon_person` / `recon_room`: đếm số
  liệu (tổng khách, số booking, số lỗi dữ liệu, số cảnh báo visa, số chênh
  lệch đối chiếu...) — không có tên, số hộ chiếu, hay số phòng cụ thể của
  từng khách
- `handover`: chỉ đếm số ghi chú theo phân loại (vd "Khách nợ": 2) — **không**
  kèm nội dung ghi chú tự do (nội dung đó có thể chứa thông tin nhạy cảm)

## Mục đích

Được một quy trình phân tích tự động đọc định kỳ để nhận diện thói quen sử
dụng / quy trình ca làm việc, từ đó đề xuất cải tiến cho web app — xem
"Rà soát lỗi tiềm ẩn" và routine phân tích quy trình ca làm việc.

**Trước khi commit**, hãy dùng ô "Xem trước nội dung sẽ xuất" trong app để tự
kiểm tra file không chứa dữ liệu cá nhân nào.
