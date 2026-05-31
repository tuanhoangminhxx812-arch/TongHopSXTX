# ⚡ Ứng dụng Tổng hợp Báo cáo INV
**Công ty Điện lực Vũng Tàu - TCKT_PCVT**

---

## 📋 Mô tả

Ứng dụng Streamlit tự động tổng hợp dữ liệu từ các file Excel nguồn (INV-017A, INV-011A, GL-BU-006A, ...) thành 1 file báo cáo tổng hợp giống sheet **tonghop** trong file Báo cáo tháng.

---

## 🚀 Cách chạy

**Cách 1**: Double-click file `chay_app.bat`

**Cách 2**: Mở PowerShell, gõ:
```
python -m streamlit run app.py
```

Sau đó mở trình duyệt tại: **http://localhost:8501**

---

## 📁 Cấu trúc thư mục

```
app_tonghop/
├── app.py              ← Ứng dụng chính
├── phan_loai_mau.json  ← Dữ liệu mẫu T2/2026
├── requirements.txt    ← Thư viện cần thiết
├── chay_app.bat        ← Script chạy app nhanh
└── README.md           ← Hướng dẫn này

Mỗi tháng tạo 1 thư mục riêng:
T2/
├── INV-017A-Tháng 02-2026.xlsx
├── INV-011A-Tháng 02-2026.xlsx
├── GL-BU-006A-Tháng 02-2026.xlsx
└── phan_loai.xlsx          ← Anh tự cập nhật
```

---

## 📋 File phân loại (`phan_loai.xlsx`)

| Cột | Nội dung |
|-----|---------|
| Mã hồ sơ | VTA26011N05024, ... |
| Date mở HS | dd/mm/yyyy |
| Date END | dd/mm/yyyy |
| User | DIEU, SA, ... |
| Loại CP | GMDK, SCTX, BTDK, SUCO, ... |
| Đội | Tên đội |
| PA/Ttr | Số PA/Tờ trình |
| **Phân loại** | **Gia hạn / Quyết toán / Đã quyết toán** |
| Quyết toán | Số tiền quyết toán (VND) |
| Ghi chú | Ghi chú thêm |

---

## 🔢 Tài khoản cố định (Cột H→O)

| Cột | Tài khoản |
|-----|----------|
| H | 24231 |
| I | 24232 |
| J | 33195 |
| K | 2428 |
| L | 154163 |
| M | 627961 |
| N | 6416211 |
| O | 6276211 |

---

## 📊 Cột trong file tổng hợp

| Cột | Tên | Nguồn |
|-----|-----|-------|
| A-C | Date mở HS, Date END, User | File phân loại |
| D | Mã hồ sơ | File phân loại |
| E-G | Loại CP, Đội, PA/Ttr | File phân loại |
| H→O | XUAT-NHAP theo tài khoản | SUMIFS từ 17A |
| P | Tổng XN | =SUM(H:O) |
| Q | Quyết toán | File phân loại |
| R | INV-QToán | =P-Q |
| T | INV017a chi phí | SUM(L:O) |
| U | BU6A chi phí | SUMIF từ BU6A |
| V | Chênh lệch | =T-U |
| W | INV017a tổng | SUMIFS từ 17A (tất cả TK) |
| X | Cl INV017a | =P-W |
| Y | INV039 | SUMIFS từ 39 |
| Z | Cl INV039 | =W-Y |

---

## 📞 Liên hệ hỗ trợ

Nếu gặp lỗi hoặc cần thêm tính năng, liên hệ team AI.
