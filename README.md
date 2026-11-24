# Ứng Dụng Chuyển Đổi Hồ Sơ Thuế XML Sang Excel, Đối Chiếu QHNS (kèm Thuế Nhà Thầu) & Tính Tiền Chậm Nộp (v1.6)

Công cụ ETL (Extract – Transform – Load) chuyên dụng giúp đọc, làm sạch, chuẩn hóa và tổng hợp **100% dữ liệu từ hồ sơ khai thuế XML (HTKK / eTax)** sang file Excel chuẩn mẫu kế toán - thanh tra, kiểm tra thuế.

## Tính năng nổi bật (Phiên bản v1.6)
- **Lọc đúng Mã Cơ quan thuế quản lý khi lên QHNS & tách riêng CQT khác (`TQ_CQT_Khac`):**
  - Cho phép nhập trực tiếp **`Mã CQT quản lý (QHNS)`** trên giao diện `app.py` (mặc định `7901`, có thể đổi sang bất kỳ mã CQT nào như `7903`, `0101`... hoặc để trống nếu lấy tất cả).
  - Biểu `QHNS`, đối chiếu `F/G` trên từng năm và `DoiChieu_QHNS` chỉ lấy đúng số liệu của CQT quản lý đã nhập; toàn bộ phát sinh/đã nộp tại các CQT khác tự động tách riêng sang sheet **`TQ_CQT_Khac`** (gồm Bảng 1 tổng hợp theo Mã CQT × Năm × Tiểu mục và Bảng 2 chi tiết từng dòng giao dịch).
- **Tổng hợp & Đối chiếu trọn bộ Tờ khai Thuế Nhà Thầu Nước Ngoài (`01/NTNN`) riêng biệt:**
  - **Khu vực chọn riêng trên giao diện (`1b. Tờ khai Thuế Nhà Thầu XML`):** Tách riêng nút chọn nhiều file / thư mục XML Nhà thầu (MST nộp thay riêng) để không lẫn với MST công ty chính, kèm bộ lọc `Nhà thầu (NTNN)`.
  - **Tự động nhận diện chuẩn mẫu theo từng thời kỳ (tuyệt đối không lệch cột):** Phân biệt chính xác mẫu **TT 103/2014** (`ct6..ct13`: `ct6`=DT GTGT, `ct7`=Tỷ lệ GTGT, `ct8`=Thuế GTGT, `ct9`=DT TNDN, `ct12`=Thuế TNDN, `ct13`=Tổng thuế) và mẫu **TT 80/2021** (`ct6`=Ngày thanh toán, `ct7`=DT GTGT, `ct9`=Thuế GTGT, `ct10`=DT TNDN, `ct13`=Thuế TNDN, `ct14`=Tổng thuế), cũng như các thẻ ngữ nghĩa trực tiếp.
  - **2 sheet tổng hợp chuẩn giống hệt Thuế GTGT:**
    1. `So Thue NTNN TH`: Toàn bộ tờ khai nhà thầu (chính thức + bổ sung) theo dòng thời gian.
    2. `So Thue NTNN Cuoi`: Chốt bản hiệu lực cuối theo từng kỳ/ngày phát sinh, nhóm theo từng năm kèm dòng `Tổng cộng năm YYYY` (`=SUM(...)`) ngay bên dưới.
  - **Liên kết công thức tự động vào từng sheet năm & Đối chiếu `TQ_NTNN`:**
    - Tự động link `C95` (Doanh thu chưa bao gồm thuế GTGT), `C97` (Thuế GTGT NTNN) và `C98` (Thuế TNDN NTNN) từ `So Thue NTNN Cuoi`.
    - Đối chiếu `C97`/`C98` với `TQ_NTNN` (`F97/G97`, `F98/G98`) theo đúng quy tắc GTGT/TNDN/TNCN: khi phát hiện Tờ khai > Tổng quan NTNN (`> 1.000đ`), tự động đưa số phát sinh TQ vào Cột (4) và số chênh lệch vào Cột (5) `Số xử lý qua thanh tra` trên `QHNS`, đồng thời đối chiếu khớp 0 đồng trên `DoiChieu_QHNS`.
- **Hỗ trợ Sổ Tổng Quan Thuế Nhà Thầu riêng (`TQ_NTNN`):**
  - Cho phép chọn thêm file Sổ Tổng Quan riêng của MST nộp thay nhà thầu nước ngoài (`TQ_NTNN`) ngay trên giao diện `app.py`.
  - Tự động chép sheet `TQ_NTNN`, đối chiếu cột `F/G` dòng `97` (`Thuế GTGT Nhà thầu`) & `98` (`Thuế TNDN Nhà thầu`) trên từng sheet năm, đổ số liệu vào mục `4. Thuế nhà thầu` trên `QHNS`, và đưa đủ các tiểu mục `1701 (NTNN)`, `1052 (NTNN)` (cùng mọi tiểu mục phát sinh khác trên `TQ_NTNN`) vào `DoiChieu_QHNS` khớp 0 đồng.
- **Tự động lập Báo cáo tổng hợp số liệu đưa vào Biên bản kiểm tra (`BaoCao_BienBan`):**
  - Cài sẵn **100% công thức Excel liên kết tự động** từ các sheet Năm (`2020`..`2025`...) và `Tinh_TienChamNop` qua đủ 10 bảng: Thuế GTGT đầu ra, GTGT khấu trừ kỳ trước & đầu vào, Điều chỉnh tăng/giảm khấu trừ, Tổng hợp GTGT phải nộp & còn được khấu trừ, Thu nhập chịu thuế, Thuế TNDN ưu đãi miễn giảm, Thuế TNDN phải nộp, Thu nhập & Thuế TNCN, Thuế Nhà thầu nước ngoài, và Bảng Kết luận tổng hợp (Truy thu chi tiết từng năm, Tiền chậm nộp, Phạt 20%, Phạt trốn thuế, Phạt VPHC thủ tục, Giảm thuế GTGT khấu trừ chuyển kỳ sau & Giảm lỗ chuyển kỳ sau).
- **Tự động lập Biểu tính tiền truy thu, xử phạt VPHC & tiền chậm nộp (`Tinh_TienChamNop`):**
  - **1 ô điều khiển trung tâm (`D4`):** Chỉ cần sửa ngày lập biên bản tại ô `D4` (hoặc nhập sẵn trên giao diện `app.py`) để tự động tính lại số ngày chậm nộp, tiền chậm nộp và kiểm tra thời hiệu cho toàn bảng.
  - **Đa giai đoạn tỷ lệ chậm nộp:** Tự động tách số ngày đến hết `30/06/2016` (`0,05%/ngày` theo Luật 71/2014/QH14) và từ `01/07/2016` trở đi (`0,03%/ngày` theo Luật 106/2016/QH13 & Khoản 2 Điều 59 Luật QLT 38/2019/QH14).
  - **Tự động kiểm tra thời hiệu 5 năm & thời hạn truy thu 10 năm (Điều 8 NĐ 125/2020/NĐ-CP):** Quá 5 năm tự động đặt Phạt 20% & Phạt trốn thuế = `0` và cảnh báo trên cột Tình trạng thời hiệu, vẫn giữ nguyên truy thu & tiền chậm nộp.
  - **Tự động nhận diện kỳ kê khai & năm tài chính lệch kỳ:** Tự nhận diện GTGT theo tháng/quý (áp dụng đúng hạn nộp trước và sau 01/07/2020), tự tính hạn nộp quyết toán TNDN cho doanh nghiệp năm tài chính lệch kỳ, tự động liên kết số chênh lệch kiểm tra TNDN (`'YYYY'!E76`) & TNCN (`'YYYY'!E90`), tách riêng mục Phạt VPHC thủ tục thuế và tổng hợp theo từng sắc thuế kèm Phụ lục tra cứu căn cứ pháp lý.
- **Đối chiếu Sổ Tổng Quan (`TQ`) & Lập 2 sheet độc lập (`QHNS` & `DoiChieu_QHNS`):**
  - `QHNS`: Biểu vào Biên bản kiểm tra (đủ GTGT, TNDN, TNCN, Nhà thầu, Lệ phí môn bài; không ghi nhận số nộp thừa kỳ trước).
  - `DoiChieu_QHNS`: Đối chiếu khớp 100% đến 0 đồng cho **tất cả sắc thuế/tiểu mục có trên Sổ TQ**.
- **Hỗ trợ đa sắc thuế & đa thời kỳ Thông tư:**
  - **Thuế TNDN (`03/TNDN`):** Tự động phân tách và ánh xạ chuẩn chỉ tiêu Thông tư 156/151 (`maTKhai=03`, các năm `<= 2020`) và Thông tư 80 (`maTKhai=892`, các năm `>= 2021`) trên các sheet `KKTNDN`, `BCHDKD` (`03-1A/TNDN`), phụ lục chuyển lỗ `03-2/TNDN` và phụ lục Giao dịch liên kết `PL_GDLK2-01` (NĐ 132/2020/NĐ-CP).
  - **Báo cáo tài chính (`BCTC` - TT200/TT133):** Xuất Bảng cân đối kế toán (`B01-DN`), Báo cáo kết quả hoạt động SXKD (`B02-DN` - `PL_KQHDSXKD`) và Báo cáo lưu chuyển tiền tệ (`B03-DN` - `PL_LCTTGT`) qua các năm.
  - **Quyết toán thuế TNCN (`05/QTT-TNCN`):** Hỗ trợ đầy đủ cả Thông tư 92/156 (`maTKhai=395`) và Thông tư 80 (`maTKhai=953`) kèm 3 sheet bảng kê chuẩn mẫu `05-1`, `05-2`, `05-3` (tự động kế thừa phụ lục `05-3` từ lần kê khai trước khi tờ khai bổ sung chỉ điều chỉnh `05-1`/`05-2`).
  - **Thuế GTGT (`01/GTGT`):** Tạo sổ theo dõi toàn bộ lịch sử kê khai (`So Thue GTGT TH`), sổ chốt tờ khai hiệu lực cuối cùng theo từng năm (`So Thue GTGT Cuoi`), bảng đối chiếu tự động bằng công thức (`Chenh lech thue GTGT - TNDN`), và các phụ lục giảm thuế GTGT (`NQ43`, `NQ101`, `NQ142`).
  - **Tờ khai khác (`01/LPMB`, ...):** Bóc tách chi tiết người nộp lệ phí môn bài và toàn bộ các chi nhánh, địa điểm kinh doanh phụ thuộc (`To_Khai_Khac`).
  - **Tờ khai bổ sung (`01/KHBS` & `01-1/KHBS`):** Tổng hợp lịch sử khai bổ sung và chi tiết từng chỉ tiêu điều chỉnh tăng/giảm.
- **Bảo toàn đúng bản chất dữ liệu:**
  - Giữ nguyên số `0` đứng đầu cho `MST`, `CCCD/CMND`, `Mã CQT`, `Mã xã/phường`, `Mã tiểu mục`, `Mã hồ sơ` (định dạng `TEXT @`).
  - Định dạng chuẩn số tiền (`#,##0`), số nguyên/năm (`0`), và công thức Excel động (`=SUM(...)`).

## Cài đặt & Khởi chạy trên máy mới

### 1. Cài đặt thư viện Python
```bash
pip install -r requirements.txt
```

### 2. Chạy giao diện ứng dụng (Desktop GUI)
```bash
python app.py
```
Hoặc nhấp đúp vào file `run.bat` trên Windows.

### 3. Chạy kiểm thử tự động (Unit Tests)
```bash
python test_app.py
```

## Cấu trúc thư mục & Sổ tay nghiệp vụ ("Bộ não thứ 2")
- `SO_TAY_NGHIEP_VU_QHNS.md`: **Sổ tay nghiệp vụ chuẩn** lưu trữ toàn bộ quy tắc đối chiếu giữa Tờ khai thuế (`GTGT`, `TNDN`, `TNCN`, `Nhà thầu`) với Sổ Tổng Quan (`TQ`) trên hệ thống TMS và cách lập biểu **Quan Hệ Ngân Sách (`QHNS`)** khớp tuyệt đối với số dư cuối cùng của năm tài chính.
- `qhns_summarizer.py`: Module đối chiếu dữ liệu Sổ Tổng Quan (`TQ`), kiểm tra chênh lệch Tờ khai vs Tổng quan (`F23/G23`, `F76/G76`, `F90/G90`) và tổng hợp 2 biểu `QHNS` & `DoiChieu_QHNS`.
- `penalty_calculator.py`: Module sinh sheet `Tinh_TienChamNop` (biểu tính thuế truy thu, phạt khai sai 20%, phạt trốn thuế 1–3 lần, phạt VPHC thủ tục và tiền chậm nộp đa giai đoạn kèm kiểm tra thời hiệu 5/10 năm).
- `report_builder.py`: Module sinh sheet `BaoCao_BienBan` (báo cáo tổng hợp 10 bảng đưa vào Biên bản kiểm tra, tự động cài công thức liên kết từ các sheet năm và `Tinh_TienChamNop`).
- `payment_summarizer.py`: Module tổng hợp sổ nộp ngân sách (`so_nop.xlsx`) theo Năm Ngày nộp thuế và Tiểu mục.
- `tax_Extractor_app.py` / `app.py`: Giao diện người dùng Desktop.
- `xml_parser.py` / `tax_parser.py`: Bộ phân tích cú pháp & trích xuất dữ liệu XML thuế (Extract).
- `excel_writer.py` / `excel_exporter.py`: Bộ làm sạch, chuẩn hóa và kết xuất báo cáo Excel kèm các sheet tổng hợp kiểm tra theo từng năm (`2020`..`2025`).
- `split_catalogs.py`, `catalogs_dump.py`, `tncn_catalogs_dump.py`, `tncn_templates.py`: Danh mục chỉ tiêu chuẩn theo Thông tư 156/151/92/200 và Thông tư 80.
- `test_app.py`: Bộ kiểm thử tự động (Unit Tests).
