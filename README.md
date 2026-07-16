# HTML Excel (`HTMLEXXCEL`)

Ứng dụng spreadsheet chạy 100% trên trình duyệt — mở / sửa / lưu nhiều định dạng bảng tính, giao diện giống Excel, copy–paste nhiều ô giống **Google Sheets**.

## Cách mở

1. Mở file `index.html` bằng Chrome, Edge hoặc Firefox (double-click).
2. Hoặc kéo thả file vào cửa sổ app.

Không cần cài đặt, không cần server.

## Định dạng hỗ trợ

| Thao tác | Định dạng |
|----------|-----------|
| **Mở** | `.xlsx` · `.xls` · `.xlsm` · `.xlsb` · `.ods` · `.csv` · `.tsv` · `.txt` · `.html` |
| **Lưu** | `.xlsx` · `.xls` · `.ods` · `.csv` · `.tsv` |

Workbook Excel/ODS có **nhiều sheet** — chuyển tab ở thanh dưới lưới. Double-click tab để đổi tên; nút `+` thêm sheet.

## Cấu trúc dự án

```
HTMLEXXCEL/
├── index.html
├── README.md
├── css/
│   └── styles.css
└── js/
    ├── app.js              # Logic spreadsheet
    ├── csv.js              # Parse / ghi CSV & clipboard TSV
    ├── io.js               # Đọc/ghi đa định dạng
    ├── formulas.js         # Máy tính công thức
    └── vendor/
        └── xlsx.full.min.js  # SheetJS (Excel/ODS)
```

## Tính năng

| Nhóm | Chi tiết |
|------|----------|
| **File** | Mở/lưu đa định dạng, kéo thả, nhiều sheet, dữ liệu mẫu |
| **Lưới** | Ô giống Excel, thanh công thức, header A/B/C… |
| **Clipboard** | Copy / cắt / dán nhiều ô (TSV) — Excel & Google Sheets |
| **Công thức** | `=SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `IF`, `ROUND`… |
| **Sửa** | Undo / Redo, tìm & thay thế, thêm/xóa hàng cột |
| **Sắp xếp** | Sort A→Z / Z→A theo cột đang chọn |
| **UI** | Resize cột, fill handle, ghim header, dark mode |
| **Status bar** | Tổng / TB / Min / Max / Đếm vùng chọn |

## Phím tắt chính

- `Ctrl+O` Mở · `Ctrl+S` Lưu (định dạng hiện tại)
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

- Thư viện **SheetJS** (`js/vendor/xlsx.full.min.js`) dùng để đọc/ghi Excel & ODS.
- Công thức được tính khi **hiển thị**; khi lưu vẫn ghi chuỗi gốc (vd. `=SUM(D2:D5)`).
- Lưu CSV/TSV khi có nhiều sheet: chỉ xuất **sheet đang mở**.
- Lần đầu dán clipboard, trình duyệt có thể xin quyền — chọn **Allow**.

## Giấy phép

Dùng tự do cho mục đích cá nhân / nội bộ. SheetJS community edition theo giấy phép của SheetJS.
