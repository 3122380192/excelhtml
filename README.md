# HTML Sheets

Bảng tính trong trình duyệt với giao diện **giống Google Sheets**, hỗ trợ XLSX/CSV, công thức, máy tính và **panel gợi ý hàm**.

## Mở

Double-click `index.html` (Chrome / Edge / Firefox).

## Giao diện (Google Sheets style)

- Logo Sheets + **đổi tên file** trên header  
- Menu: **File · Edit · View · Insert · Format · Data · Tools · Help**  
- Toolbar: undo, font, **B/I/U**, màu, căn lề, số, ƒx…  
- Thanh công thức + Name box  
- Selection **xanh dương** (Sheets blue)  
- Sheet tabs **xanh lá** phía dưới  
- Dark mode (View → theme)

## Panel Functions (tab gợi ý hàm)

- Cột phải **Functions** — tìm kiếm + lọc Math / Logic / Finance / Text / Date  
- Gõ **`=`** trên thanh công thức → **autocomplete**  
- Click hàm → chèn vào ô  
- Phím: `↑` `↓` chọn · `Tab` chèn · `Esc` đóng  
- Toggle: nút **∑?** trên toolbar hoặc rail **ƒx**

## File & tính toán

| | |
|--|--|
| **Mở / tải** | XLSX, XLS, ODS, CSV, TSV, HTML |
| **Công thức** | SUM, AVERAGE, IF, PMT, FV… (xem panel Functions) |
| **Tools** | Máy tính, VAT, lãi, trả góp, % tổng… |
| **Clipboard** | Copy/paste nhiều ô như Sheets |

## Phím tắt

`Ctrl+O/S/N` · `Ctrl+C/X/V` · `Ctrl+B/I/U` · `Ctrl+Z/Y` · `Ctrl+F/H` · `Alt+=` AutoSum · `Ctrl+Shift+C` Calculator · `F1` Help

## Cấu trúc

```
HTMLEXXCEL/
├── index.html
├── css/styles.css
└── js/
    ├── app.js          # Core spreadsheet
    ├── sheets-ui.js    # Menus + function panel
    ├── fx-catalog.js   # Function catalog
    ├── formulas.js
    ├── calc.js
    ├── csv.js / io.js
    └── vendor/xlsx.full.min.js
```

Repo: https://github.com/3122380192/excelhtml
