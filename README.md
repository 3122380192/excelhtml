# HTML Excel (`HTMLEXXCEL`)

Ứng dụng spreadsheet chạy 100% trên trình duyệt — mở / sửa / lưu **CSV** giống Excel, copy–paste nhiều ô giống **Google Sheets**.

## Cách mở

1. Mở file `index.html` bằng Chrome, Edge hoặc Firefox (double-click).
2. Hoặc kéo thả file `.csv` vào cửa sổ app.

Không cần cài đặt, không cần server.

## Cấu trúc dự án

```
HTMLEXXCEL/
├── index.html          # Giao diện chính
├── css/
│   └── styles.css      # Giao diện (sáng / tối)
├── js/
│   ├── csv.js          # Parse / ghi CSV & clipboard TSV
│   ├── formulas.js     # Máy tính công thức (=SUM, =A1*B1…)
│   └── app.js          # Logic spreadsheet
└── README.md
```

## Tính năng

| Nhóm | Chi tiết |
|------|----------|
| **File** | Mở CSV/TSV, lưu CSV UTF-8 BOM, kéo thả, sheet mới, dữ liệu mẫu |
| **Lưới** | Ô giống Excel, thanh công thức, header A/B/C…, số hàng |
| **Clipboard** | Copy / cắt / dán nhiều ô (TSV) — tương thích Excel & Google Sheets |
| **Công thức** | `=SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `IF`, `ROUND`, `ABS`, `LEN`, `CONCAT`, `A1+B2`, vùng `A1:B5` |
| **Sửa** | Undo / Redo, tìm & thay thế, xóa vùng, thêm/xóa hàng cột |
| **Sắp xếp** | Sort A→Z / Z→A theo cột đang chọn (giữ hàng header) |
| **UI** | Resize cột, auto-fit (double-click mép), fill handle, ghim header, dark mode |
| **Status bar** | Tổng / TB / Min / Max / Đếm vùng chọn (như Excel) |

## Phím tắt chính

- `Ctrl+O` Mở · `Ctrl+S` Lưu  
- `Ctrl+C` / `X` / `V` Copy / Cắt / Dán  
- `Ctrl+Z` / `Y` Undo / Redo  
- `Ctrl+F` / `H` Tìm / Thay thế  
- `F2` Sửa ô · `F1` Trợ giúp  

## Ví dụ công thức

```
=B2*C2
=SUM(D2:D10)
=AVERAGE(B2:B20)
=IF(A1>0,"OK","—")
=MAX(C2:C100)
```

## Ghi chú

- Công thức được **tính khi hiển thị**; khi **Lưu CSV** vẫn ghi **chuỗi gốc** (ví dụ `=SUM(D2:D5)`), mở lại vẫn dùng được trong app.
- Lần đầu dán từ clipboard, trình duyệt có thể xin quyền — chọn **Allow**.
- File cũ `csv-excel.html` (nếu còn ở thư mục cha) có thể xóa; dùng bản trong folder này.

## Giấy phép

Dùng tự do cho mục đích cá nhân / nội bộ.
