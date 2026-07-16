/**
 * Function catalog for suggestions (Google Sheets style)
 */
(function (global) {
  "use strict";

  const CATALOG = [
    { name: "SUM", cat: "Math", sig: "SUM(value1, [value2, ...])", desc: "Cộng tất cả số trong vùng hoặc danh sách.", ex: "=SUM(A1:A10)" },
    { name: "AVERAGE", cat: "Math", sig: "AVERAGE(value1, [value2, ...])", desc: "Trung bình cộng các số.", ex: "=AVERAGE(B2:B20)" },
    { name: "COUNT", cat: "Math", sig: "COUNT(value1, [value2, ...])", desc: "Đếm ô chứa số.", ex: "=COUNT(A1:A100)" },
    { name: "COUNTA", cat: "Math", sig: "COUNTA(value1, [value2, ...])", desc: "Đếm ô không trống.", ex: "=COUNTA(A1:A100)" },
    { name: "MIN", cat: "Math", sig: "MIN(value1, [value2, ...])", desc: "Giá trị nhỏ nhất.", ex: "=MIN(C2:C50)" },
    { name: "MAX", cat: "Math", sig: "MAX(value1, [value2, ...])", desc: "Giá trị lớn nhất.", ex: "=MAX(C2:C50)" },
    { name: "PRODUCT", cat: "Math", sig: "PRODUCT(value1, [value2, ...])", desc: "Nhân tất cả số.", ex: "=PRODUCT(A1:A5)" },
    { name: "MEDIAN", cat: "Math", sig: "MEDIAN(value1, [value2, ...])", desc: "Trung vị của dãy số.", ex: "=MEDIAN(A1:A20)" },
    { name: "STDEV", cat: "Math", sig: "STDEV(value1, [value2, ...])", desc: "Độ lệch chuẩn mẫu.", ex: "=STDEV(A1:A20)" },
    { name: "VAR", cat: "Math", sig: "VAR(value1, [value2, ...])", desc: "Phương sai mẫu.", ex: "=VAR(A1:A20)" },
    { name: "ROUND", cat: "Math", sig: "ROUND(number, [places])", desc: "Làm tròn số.", ex: "=ROUND(A1, 2)" },
    { name: "ABS", cat: "Math", sig: "ABS(number)", desc: "Giá trị tuyệt đối.", ex: "=ABS(A1)" },
    { name: "SQRT", cat: "Math", sig: "SQRT(number)", desc: "Căn bậc hai.", ex: "=SQRT(A1)" },
    { name: "POWER", cat: "Math", sig: "POWER(base, exponent)", desc: "Lũy thừa.", ex: "=POWER(2, 10)" },
    { name: "MOD", cat: "Math", sig: "MOD(number, divisor)", desc: "Phần dư phép chia.", ex: "=MOD(A1, 2)" },
    { name: "INT", cat: "Math", sig: "INT(number)", desc: "Phần nguyên (làm tròn xuống).", ex: "=INT(A1)" },
    { name: "CEILING", cat: "Math", sig: "CEILING(number)", desc: "Làm tròn lên số nguyên.", ex: "=CEILING(A1)" },
    { name: "FLOOR", cat: "Math", sig: "FLOOR(number)", desc: "Làm tròn xuống số nguyên.", ex: "=FLOOR(A1)" },
    { name: "LOG", cat: "Math", sig: "LOG(number, [base])", desc: "Logarithm.", ex: "=LOG(100, 10)" },
    { name: "LN", cat: "Math", sig: "LN(number)", desc: "Logarit tự nhiên.", ex: "=LN(A1)" },
    { name: "EXP", cat: "Math", sig: "EXP(number)", desc: "e mũ số.", ex: "=EXP(1)" },
    { name: "PI", cat: "Math", sig: "PI()", desc: "Hằng số π ≈ 3.14159.", ex: "=PI()" },
    { name: "SIN", cat: "Math", sig: "SIN(angle)", desc: "Sin (radian).", ex: "=SIN(PI()/2)" },
    { name: "COS", cat: "Math", sig: "COS(angle)", desc: "Cos (radian).", ex: "=COS(0)" },
    { name: "TAN", cat: "Math", sig: "TAN(angle)", desc: "Tan (radian).", ex: "=TAN(PI()/4)" },
    { name: "FACT", cat: "Math", sig: "FACT(number)", desc: "Giai thừa n!.", ex: "=FACT(5)" },
    { name: "GCD", cat: "Math", sig: "GCD(n1, n2, ...)", desc: "Ước chung lớn nhất.", ex: "=GCD(24, 36)" },
    { name: "LCM", cat: "Math", sig: "LCM(n1, n2, ...)", desc: "Bội chung nhỏ nhất.", ex: "=LCM(4, 6)" },
    { name: "RANDBETWEEN", cat: "Math", sig: "RANDBETWEEN(low, high)", desc: "Số nguyên ngẫu nhiên trong khoảng.", ex: "=RANDBETWEEN(1, 100)" },
    { name: "RAND", cat: "Math", sig: "RAND()", desc: "Số ngẫu nhiên 0–1.", ex: "=RAND()" },
    { name: "IF", cat: "Logic", sig: "IF(condition, value_if_true, value_if_false)", desc: "Trả về giá trị theo điều kiện.", ex: '=IF(A1>0,"OK","—")' },
    { name: "AND", cat: "Logic", sig: "AND(logical1, [logical2, ...])", desc: "TRUE nếu mọi điều kiện đúng.", ex: "=AND(A1>0, B1>0)" },
    { name: "OR", cat: "Logic", sig: "OR(logical1, [logical2, ...])", desc: "TRUE nếu ít nhất một đúng.", ex: "=OR(A1>0, B1>0)" },
    { name: "NOT", cat: "Logic", sig: "NOT(logical)", desc: "Đảo logic.", ex: "=NOT(A1=0)" },
    { name: "PMT", cat: "Finance", sig: "PMT(rate, nper, pv)", desc: "Khoản trả định kỳ (trả góp).", ex: "=PMT(0.01, 12, -10000000)" },
    { name: "FV", cat: "Finance", sig: "FV(rate, nper, pmt, [pv])", desc: "Giá trị tương lai.", ex: "=FV(0.01, 12, -1000000)" },
    { name: "PV", cat: "Finance", sig: "PV(rate, nper, pmt, [fv])", desc: "Giá trị hiện tại.", ex: "=PV(0.01, 12, -1000000)" },
    { name: "NPV", cat: "Finance", sig: "NPV(rate, value1, [value2, ...])", desc: "Giá trị hiện tại ròng.", ex: "=NPV(0.1, 100, 200, 300)" },
    { name: "LEN", cat: "Text", sig: "LEN(text)", desc: "Độ dài chuỗi.", ex: "=LEN(A1)" },
    { name: "UPPER", cat: "Text", sig: "UPPER(text)", desc: "Chữ HOA.", ex: "=UPPER(A1)" },
    { name: "LOWER", cat: "Text", sig: "LOWER(text)", desc: "Chữ thường.", ex: "=LOWER(A1)" },
    { name: "TRIM", cat: "Text", sig: "TRIM(text)", desc: "Xóa khoảng trắng thừa.", ex: "=TRIM(A1)" },
    { name: "LEFT", cat: "Text", sig: "LEFT(text, [count])", desc: "Lấy ký tự bên trái.", ex: "=LEFT(A1, 3)" },
    { name: "RIGHT", cat: "Text", sig: "RIGHT(text, [count])", desc: "Lấy ký tự bên phải.", ex: "=RIGHT(A1, 3)" },
    { name: "MID", cat: "Text", sig: "MID(text, start, count)", desc: "Lấy chuỗi con.", ex: "=MID(A1, 2, 5)" },
    { name: "CONCAT", cat: "Text", sig: "CONCAT(text1, [text2, ...])", desc: "Nối chuỗi.", ex: '=CONCAT(A1, " ", B1)' },
    { name: "TODAY", cat: "Date", sig: "TODAY()", desc: "Ngày hôm nay.", ex: "=TODAY()" },
    { name: "NOW", cat: "Date", sig: "NOW()", desc: "Ngày giờ hiện tại.", ex: "=NOW()" },
    { name: "PERCENT", cat: "Math", sig: "PERCENT(part, whole)", desc: "part / whole.", ex: "=PERCENT(A1, B1)" },
  ];

  const CATS = ["All", "Math", "Logic", "Finance", "Text", "Date"];

  function search(query, cat) {
    const q = (query || "").trim().toUpperCase().replace(/^=/, "");
    return CATALOG.filter((f) => {
      if (cat && cat !== "All" && f.cat !== cat) return false;
      if (!q) return true;
      return f.name.includes(q) || f.desc.toUpperCase().includes(q) || f.sig.toUpperCase().includes(q);
    });
  }

  function get(name) {
    return CATALOG.find((f) => f.name === String(name).toUpperCase()) || null;
  }

  global.FxCatalog = { CATALOG, CATS, search, get };
})(typeof window !== "undefined" ? window : globalThis);
