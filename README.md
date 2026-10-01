# Tool tổng hợp CRM – NTS Hanoi

Tool HTML **một file duy nhất, chạy offline** để tổng hợp và phân tích dữ liệu thương vụ (deals) export từ CRM.
Dữ liệu được xử lý ngay trên trình duyệt: không cần internet, không gửi đi đâu.

**Dùng tool:**
- Online: https://pm-ntshn.github.io/Tool-CRM/ (GitHub Pages, phục vụ file `index.html`). File Excel nạp vào chỉ được xử lý trên trình duyệt của người dùng, không tải lên máy chủ.
- Offline: mở file [`Tool tổng hợp CRM.html`](Tool%20tổng%20hợp%20CRM.html) (giống hệt `index.html`) bằng Chrome / Edge / Firefox (bản mới), kéo – thả file Excel vào.

## Luồng xử lý
1. Nạp file **raw export từ CRM** (`export.deals.report…xlsx`), hoặc file **template chuẩn** đã convert (sheet `Data`). Tool tự nhận diện loại file.
2. File raw được convert sang template chuẩn bằng `src/convert.js`, bản JavaScript chuyển thể 1:1 từ `reference/convert_deals.py` (cùng quy tắc, tham số và nội dung cờ chất lượng).
3. Mọi phân tích dùng dữ liệu template. Nút **File chuẩn** tải về `CRM_deals_template_dd_mm_yy.xlsx` (gồm 3 sheet `Data`, `Huong_dan`, `Danh_muc`, giống bản Python).
4. Nút **Xuất Excel tổng hợp** tạo file `Tổng hợp CRM_ Tháng M.YYYY.xlsx` gồm các sheet: Tổng quan, Danh sách thương vụ, Theo Sale / Hãng / Reseller / Nhóm KH / Khách hàng, Sale × Giai đoạn, Forecast theo quý, Cần nhắc cập nhật, Vấn đề dữ liệu, Data chuẩn và Ghi chú.

**Ngày chốt số liệu** dùng để tính cờ nhắc và timeline. Tool lấy ngày này từ tên file (`…01.10.26.xlsx` hoặc `…_01_10_26.xlsx`); nếu tên file không có ngày thì lấy ngày hôm nay. Có thể sửa ngày chốt ở màn hình nạp file hoặc trên thanh tiêu đề.

## Cấu trúc mã nguồn
| File | Nội dung |
|---|---|
| `src/index.html`, `src/styles.css` | Khung giao diện, nhận diện màu NTS |
| `src/xlsx-io.js` | Bộ đọc/ghi .xlsx tự viết, không phụ thuộc thư viện ngoài |
| `src/convert.js` | Convert raw sang template, bản JS của `reference/convert_deals.py` |
| `src/model.js` | Mô hình deal, cờ nhắc, các chỉ số tổng hợp |
| `src/report.js` | Dựng file Excel báo cáo |
| `src/app.js` | Dashboard: bộ lọc, gợi ý tìm kiếm, biểu đồ, bảng, chi tiết, soạn nội dung nhắc |
| `vendor/chart.umd.min.js` | Chart.js 4.5.1 (MIT) |

Build lại file HTML (tạo cả `Tool tổng hợp CRM.html` và `index.html`) sau khi sửa mã nguồn:
```bash
node build.mjs
```

Kiểm tra bản JS cho kết quả giống hệt bản Python (dữ liệu thật không commit vào repo):
```bash
python reference/convert_deals.py raw.xlsx py_out.xlsx --ref-date 01/10/2026
node test/convert.test.mjs raw.xlsx py_out.xlsx 01/10/2026
```

> Khi sửa quy tắc convert (danh mục sale, hãng, phân nhóm KH…), sửa đồng thời `reference/convert_deals.py` và `src/convert.js`, sau đó chạy lại test.
