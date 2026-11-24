# Dashboard KPI Team SS Văn — bản GitHub

Giao diện giữ nguyên như dashboard cũ (đủ 16 tab), có thêm tab **TRẬN XANH**. Số liệu đi theo file Tracking KPI. Ai có link đều xem được, kể cả ĐDKD mở trên điện thoại hoặc qua Zalo.

> ⚠️ Link GitHub Pages là **công khai**. Ai có link đều xem được tên, địa chỉ và ngày sinh của KH. Chỉ gửi link trong nhóm Zalo nội bộ.

---

## A. Cài đặt lần đầu (khoảng 15 phút)

1. **Tạo tài khoản GitHub**: vào github.com → Sign up. Nếu Anh đã có tài khoản thì bỏ qua bước này.
2. **Tạo kho mới**: bấm dấu **+** (góc trên bên phải) → *New repository*.
   - Repository name: `kpi-dashboard`
   - Chọn **Public**. Bản GitHub miễn phí chỉ làm được link xem cho kho Public.
   - Bấm *Create repository*.
3. **Tải mã nguồn lên**: trong kho vừa tạo, bấm *uploading an existing file*.
   - Giải nén file zip em gửi.
   - Mở thư mục `kpi-dashboard-github`.
   - Chọn **tất cả** mục bên trong (gồm `.github`, `build`, `data`, `src`, `HUONG-DAN.md`, `.gitignore`), kéo thả vào trang GitHub rồi bấm *Commit changes*.
   - Nếu thư mục `.github` không lên: bấm *Add file → Create new file*, gõ tên `.github/workflows/deploy.yml`, rồi dán nội dung file `deploy.yml` trong zip vào.
4. **Bật trang web**: vào *Settings → Pages* → mục *Source* chọn **GitHub Actions**.
5. **Chạy lần đầu**: vào tab *Actions* → chọn *Cập nhật dashboard* → bấm *Run workflow*. Đợi khoảng 2 phút cho tới khi hiện dấu ✅.
6. **Link dashboard**: `https://<tên-tài-khoản>.github.io/kpi-dashboard/`. Anh gửi link này vào nhóm Zalo ĐDKD.
   - Muốn mở thẳng tab Trận Xanh: thêm `#tranxanh` vào cuối link.

## B. Tạo "chìa khoá" để cập nhật từ dashboard (làm 1 lần)

1. Vào GitHub → ảnh đại diện → *Settings → Developer settings → Personal access tokens → Fine-grained tokens → Generate new token*.
2. Điền như sau:
   - Token name: `kpi-dashboard`
   - Expiration: 1 năm
   - Repository access: *Only select repositories* → chọn `kpi-dashboard`
   - Permissions → Repository permissions → **Contents: Read and write**
3. Bấm *Generate token*, rồi copy chuỗi `github_pat_…`. Không gửi chuỗi này cho ai.

## C. Cập nhật số liệu hằng ngày

**Cách 1 — thả file ngay trên dashboard (khuyên dùng):**

1. Mở link dashboard trên **máy tính** (Chrome hoặc Edge).
2. Bấm nút **⇪ Nạp file mới** ở góc dưới bên trái, rồi chọn file Tracking KPI `.xlsb`. File 40MB đọc mất khoảng 20–40 giây.
3. Lần đầu dashboard sẽ hỏi tên kho (`tên-tài-khoản/kpi-dashboard`) và token. Anh dán token rồi bấm **Lưu lên GitHub**. Máy sẽ nhớ token cho những lần sau.
4. Sau **2–3 phút**, link chung có số mới. ĐDKD chỉ cần tải lại trang.

Nếu ĐDKD bấm *Nạp file mới*, họ chỉ xem được trên máy của họ (chọn *Chỉ xem trên máy này*). Không có token thì họ không ghi đè được dữ liệu chung.

**Cách 2 — thả file vào GitHub:**

- Kéo file `.xlsb` vào thư mục `data/` của kho. GitHub tự đọc file, cập nhật số liệu và xoá file Excel gốc.
- Lưu ý: GitHub trên web chỉ nhận file **dưới 25MB**. File 40MB phải dùng GitHub Desktop hoặc Cách 1.

## D. Ghi chú kỹ thuật

- File cần có các sheet: KPI THÁNG, MCP, 061 new, DSKH Combo, Sheet10, HV11 • Tracking, Pivot, Cap Tran, T.Bày, Bonus, LE, SINH NHẬT THÁNG, Data_Chỉ tiêu, **Trận Xanh**. Thiếu sheet nào thì phần liên quan sẽ để trống, dashboard không báo lỗi.
- Tab **Trận Xanh**:
  - Tổng hợp theo ĐDKD và danh sách KH đã mua lấy nguyên từ sheet Trận Xanh.
  - CT/TH ASO Trận Xanh lấy từ KPI THÁNG.
  - Danh sách **KH chưa mua** = KH Off Premise trong MCP chưa có trong danh sách đã mua, xếp ưu tiên Đại Sứ → VIP3-5 → Lẻ.
  - Có lọc theo ĐDKD, thứ, nhóm và tìm kiếm. Có nút xuất Excel.
- Ảnh tab "DDKD Chụp" (`src/ddkd-photos-v202.js`) là dữ liệu cố định từ file DanhSachXetBoHinhTrungBay. File Tracking KPI không cập nhật phần này.
- Muốn xem thử trên máy mà chưa cần GitHub: mở `XEM-THU/index.html`.
