# ssmanage

Hệ thống quản lý thu mua thủy hải sản trong nước và xuất khẩu (Super Seafoods). Google Sheet là bộ não, trang web chỉ đọc và hiển thị.

## Cấu trúc

| Đường dẫn | Nội dung |
|---|---|
| `index.html` | Bản xem mẫu, dữ liệu giả định, mở thẳng trên trình duyệt hoặc GitHub Pages. Không nối Sheet. |
| `apps-script/Code.gs` | Mã Apps Script: đọc Sheet (`getData`), thêm dòng (`addRow`), sửa dòng (`updateRows`). |
| `apps-script/Index.html` | Giao diện bản thật, dán vào Apps Script với tên `Index`. Đọc dữ liệu từ Sheet qua `google.script.run`. |
| `data/Thuy-hai-san-bo-nao.xlsx` | Sheet mẫu đủ các tab, nhập vào Google Sheet làm bộ não. |
| `` | Hướng dẫn cài từng bước và prompt cho Director. |

## Cài đặt nhanh

1. Mở Google Sheet **QUẢN LÝ SUPER SEAFOODS**, nhập file `data/Thuy-hai-san-bo-nao.xlsx` (Tệp > Nhập > Tải lên > Thay thế bảng tính).
2. Tiện ích mở rộng > Apps Script, dán `apps-script/Code.gs`, thêm file HTML tên `Index` và dán `apps-script/Index.html`.
3. Triển khai thành ứng dụng web, chạy dưới tên người dùng đang truy cập.

Chi tiết xem ``.

## Lưu ý

- Người dùng cần quyền **Người chỉnh sửa** trên Sheet để dùng các nút thêm, sửa. Người chỉ xem vẫn xem được dữ liệu.
- Dữ liệu trong `index.html` và file xlsx là dữ liệu mẫu (có chữ "(mẫu)"), thay bằng dữ liệu thật trong Sheet.
- Danh sách chứng từ theo thị trường và điều kiện nhà cung cấp trong bản mẫu chỉ là gợi ý, cần xác nhận trước khi dùng.
