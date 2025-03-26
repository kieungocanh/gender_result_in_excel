Mục tiêu: nhằm cung cấp số liệu thống kê tổng quan về thị trường cho khách hàng cần 1 hệ thống sẽ tự động sinh ra các bảng kết quả theo một template Excel mẫu định sẵn từ dữ liệu đã được chuẩn hoá.
Ý tưởng:
1. Module tính toán
- Chứa các hàm xử lý dữ liệu, tính toán số liệu thống kê.
- Thiết kế dưới dạng class để dễ mở rộng và tái sử dụng.
- Mỗi bảng kết quả tương ứng với một DataFrame.
  
2. Config bảng
- Định nghĩa tên bảng và vị trí tương ứng trong file Excel template.

3. Sinh bảng kết quả
- Đọc file template mẫu.
- Ghi các DataFrame vào đúng vị trí đã định nghĩa.
