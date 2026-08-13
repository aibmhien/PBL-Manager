# PBL Manager V0.3.0

Ứng dụng web thử nghiệm cho **PBL1. Thiết kế hệ thống cơ khí**.

## Các thay đổi chính trong V0.3.0

- Tiêu đề trang quản lý hiển thị tên **Giảng viên hướng dẫn** từ phần Cấu hình.
- Tiêu đề chính: **PBL1. Thiết kế hệ thống cơ khí**.
- Tiêu đề phụ: **Ứng dụng để quản lý PBL1 và theo dõi tiến độ thực hiện của sinh viên**.
- Có trang **Phân công nhóm** để:
  - chuyển sinh viên sang nhóm khác;
  - tạo nhóm mới;
  - xử lý sinh viên chưa có thông tin Đề;
  - đổi loại DA của cả nhóm và cấp lại bộ số liệu phù hợp.
- Cổng sinh viên có chức năng **báo cáo tiến độ + upload nhiều file**:
  - `.docx` thuyết minh;
  - `.xls`, `.xlsx` tính toán;
  - `.dwg`, `.dxf` bản vẽ AutoCAD.
- Giảng viên xem và tải file sinh viên đã nộp ngay trong trang Chi tiết nhóm.
- Mỗi file upload tối đa 50 MB.
- Trang quản trị có mật khẩu đơn giản để tránh sinh viên sửa dữ liệu khi website được đưa lên Internet.
- Vẫn hỗ trợ in / **Save as PDF** bằng Chrome hoặc Edge trên Windows.

## 1. Chạy trên Windows

Giải nén toàn bộ thư mục, sau đó chạy:

`run_windows.bat`

Mở trình duyệt:

- Trang giảng viên: `http://127.0.0.1:5000/`
- Cổng sinh viên: `http://127.0.0.1:5000/student`

Mật khẩu giảng viên mặc định của bản thử nghiệm: `pbl123`.

Để đổi mật khẩu trước khi chạy, trong Command Prompt:

```bat
set PBL_ADMIN_PASSWORD=MatKhauCuaThay
run_windows.bat
```

## 2. Vì sao 127.0.0.1 không dùng được trên máy khác?

`127.0.0.1` luôn có nghĩa là **chính máy đang mở trình duyệt**. Nó không phải địa chỉ để chia sẻ cho máy khác.

Trong cùng Wi-Fi/LAN, hãy tìm IPv4 của máy giảng viên bằng lệnh:

```bat
ipconfig
```

Ví dụ IPv4 là `192.168.1.20`, sinh viên trong cùng mạng dùng:

`http://192.168.1.20:5000/student`

Ứng dụng đã lắng nghe trên `0.0.0.0`, nên về phía phần mềm đã cho phép truy cập LAN. Nếu vẫn không vào được, cần cho phép Python hoặc cổng TCP 5000 trong Windows Firewall.

## 3. Đưa lên GitHub và tạo link Internet

Không có "file định dạng http" để chạy website. HTTP là giao thức; ứng dụng cần một máy chủ chạy Python liên tục.

GitHub dùng để lưu mã nguồn. **GitHub Pages không chạy được FastAPI/Python backend**, vì vậy cách đơn giản cho V0.3.0 là:

1. Tạo một repository mới trên GitHub.
2. Upload toàn bộ nội dung thư mục `PBL_Manager_V0.3.0` vào repository.
3. Tạo tài khoản Render.
4. Chọn **New > Blueprint** và kết nối repository GitHub.
5. Render đọc file `render.yaml` và tạo web service.
6. Khi Render yêu cầu biến `PBL_ADMIN_PASSWORD`, nhập mật khẩu riêng cho giảng viên.
7. Sau khi deploy, Render cấp một địa chỉ HTTPS, ví dụ:
   `https://pbl-manager-v03.onrender.com`
8. Link chia cho sinh viên nên là:
   `https://.../student`

### Lưu ý dữ liệu

`render.yaml` cấu hình thư mục dữ liệu tại `/var/data` để database SQLite và file sinh viên upload không nằm trong source code. Khi triển khai thật, cần kiểm tra gói hosting đang dùng có hỗ trợ persistent disk theo cấu hình này hay không.

## 4. Quy trình sử dụng đề xuất

### Giảng viên

1. Đăng nhập trang quản lý.
2. Vào **Cấu hình** để nhập tên giảng viên, tuần bắt đầu/kết thúc, năm học.
3. Import danh sách sinh viên.
   - Chỉ bắt buộc MSSV/Số thẻ và Họ tên.
   - Cột Đề có thể có, để trống, hoặc không có.
4. Import bảng số liệu DA1–DA5.
5. Vào **Phân công nhóm** để xử lý các trường hợp đổi nhóm/chưa có đề.
6. Mở Chi tiết nhóm để xem tiến độ, tải file sinh viên và đánh giá.

### Sinh viên

1. Vào `/student`.
2. Nhập MSSV/Số thẻ.
3. Xem đề và in/Save as PDF.
4. Ở phần trang tiến độ, chọn mốc công việc, nhập mô tả tiến độ và upload file.
5. Xem trạng thái, điểm và nhận xét của giảng viên ở lần truy cập sau.

## 5. Hạn chế có chủ ý của V0.3.0

Đây vẫn là bản prototype. Việc truy cập sinh viên hiện dùng MSSV, chưa có token riêng; file nộp chưa có versioning; chưa có phân quyền nhiều giảng viên; chưa có email/thông báo; và AutoCAD chỉ được lưu/tải xuống chứ website chưa xem trực tiếp `.dwg`.

Các mục này phù hợp để xem xét cho V0.4 sau khi luồng sử dụng V0.3 được đánh giá.

## Ghi chú V0.3.1 - Render Free
Nếu dùng Render Free, hãy dùng `render.yaml` trong gói V0.3.1. Bản này đã bỏ persistent disk và ghim Python 3.12. Không cần bật GitHub Pages. Xem `DEPLOY_RENDER_FREE.md`.
