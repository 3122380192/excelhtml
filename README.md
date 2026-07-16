# HTML Excel

Ứng dụng bảng tính **giống Microsoft Excel** chạy 100% trong trình duyệt — ribbon, định dạng ô, công thức, nhiều sheet, mở/lưu XLSX.

![HTML Excel](https://img.shields.io/badge/Excel-like-217346?style=flat-square)

## Mở nhanh

Double-click `index.html` (Chrome / Edge / Firefox). Không cần cài đặt hay server.

## Giao diện giống Excel

- **Title bar** xanh Excel + tên file
- **Ribbon**: Tệp · Trang chủ · Chèn · Công thức · Dữ liệu · Xem
- **Backstage (Tệp)**: Mới / Mở / Lưu / Xuất / Thông tin
- **Thanh công thức** + Name box + ✓/✕
- **Lưới** header A/B/C…, số hàng, fill handle
- **Sheet tabs** dưới lưới
- **Status bar** xanh: Tổng / TB / Min / Max / Zoom

## Định dạng hỗ trợ

| | |
|--|--|
| **Mở** | XLSX, XLS, XLSM, ODS, CSV, TSV, TXT, HTML |
| **Lưu** | XLSX, XLS, ODS, CSV, TSV |

## Tính năng chính

| Nhóm | Chi tiết |
|------|----------|
| **Clipboard** | Cut / Copy / Paste nhiều ô (TSV — Sheets & Excel) |
| **Phông chữ** | Family, size, **B** / *I* / U, màu chữ, màu nền |
| **Căn lề** | Trái / giữa / phải, wrap text |
| **Số** | Chung, số, tiền ₫, %, thập phân |
| **Viền** | Ngoài, dưới, xóa viền |
| **Công thức** | SUM, AVERAGE, COUNT, MIN, MAX, IF, `=A1*B1`… |
| **AutoSum** | Nút ∑ hoặc `Alt`+`=` |
| **Dữ liệu** | Sort A↔Z, tìm/thay, tách cột |
| **Xem** | Ghim header, đường lưới, zoom 50–200% |
| **Sheet** | Nhiều sheet, thêm/đổi tên/xóa |
| **Khác** | Undo/Redo, format painter, menu chuột phải, dark mode |

## Phím tắt

| Phím | Hành động |
|------|-----------|
| `Ctrl+O` / `S` / `N` | Mở / Lưu / Mới |
| `Ctrl+C` `X` `V` | Copy / Cắt / Dán |
| `Ctrl+B` `I` `U` | Đậm / Nghiêng / Gạch dưới |
| `Ctrl+Z` / `Y` | Undo / Redo |
| `Ctrl+F` / `H` | Tìm / Thay |
| `Alt+*`=` | AutoSum |
| `F2` | Sửa ô |
| `F1` | Trợ giúp |
| Chuột phải | Menu ngữ cảnh |

## Cấu trúc

```
HTMLEXXCEL/
├── index.html
├── css/styles.css
└── js/
    ├── app.js
    ├── csv.js
    ├── io.js
    ├── formulas.js
    └── vendor/xlsx.full.min.js   # SheetJS
```

## Ghi chú

- Style ô (màu, đậm…) lưu trong phiên làm việc; export XLSX/CSV hiện xuất **giá trị & công thức** (không full style OOXML).
- Công thức tính khi hiển thị; file lưu giữ chuỗi gốc `=SUM(...)`.

## Giấy phép

Dùng tự do. SheetJS community theo license của SheetJS.
