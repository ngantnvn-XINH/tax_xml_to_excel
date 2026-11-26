# Quy Chuẩn Kỹ Thuật & Nghiệp Vụ Dự Án `tax_xml_to_excel`

Mọi chỉnh sửa mã nguồn, mẫu Excel, test và tài liệu trong dự án `tax_xml_to_excel` phải tuân thủ tuyệt đối các nguyên tắc sau:

## 1. Độ Chính Xác Dữ Liệu Tờ Khai (Zero-Error Policy)
- **NGUYÊN TẮC BẤT DI BẤT DỊCH — COPY NGUYÊN BẢN TIÊU ĐỀ & CHỈ TIÊU TỜ KHAI GỐC:**
  - Nhiệm vụ khi tổng hợp tờ khai là **copy nguyên bản 100% tiêu đề, thứ tự cột và mã chỉ tiêu đúng như mẫu tờ khai gốc** theo từng thông tư/thời kỳ (VD: `01/NTNN` TT103 13 cột `(1)..(13)` vs TT80 15 cột `(1a), (1b), (2)..(14)`; `01/GTGT`, `03/TNDN`, `05/QTT-TNCN` theo từng thông tư).
  - **TUYỆT ĐỐI KHÔNG** tự ý sáng tạo, thêm bớt, gộp cột, đổi tên chỉ tiêu hay suy diễn cột khác với tờ khai gốc làm sai lệch bản chất số liệu.
- **Năm tài chính lệch kỳ (VD: 07/YYYY–06/(YYYY+1) hoặc 04/YYYY–03/(YYYY+1)):**
  - Khi xác định năm tính thuế (`TaxDeclaration.year`) cho `03/TNDN` và `BCTC`, **bắt buộc ưu tiên thẻ `<kyKKhai>` trước `den_ngay`** để tuyệt đối không bị dịch +1 năm.
  - Tự động nhận diện khoảng kỳ tài chính của `03/TNDN`; nếu là năm tài chính lệch kỳ, bắt buộc sinh thêm sheet `'So Thue GTGT Cuoi (Lech ky)'` tổng hợp tờ khai GTGT theo đúng khung tháng tài chính của TNDN và nối tự động sang `'Chenh lech thue GTGT - TNDN'`.
- **Tờ khai bổ sung cuối cùng:**
  - Nhận diện đầy đủ các thẻ lần bổ sung trong XML (`soLan`, `soLanBoSung`, `soLanBSung`, `lanBoSung`, `lanBSung`, `loaiTKhai=="B"`).
  - Luôn chọn bản có `so_lan` lớn nhất (cùng lần thì chọn `ngay_lap` mới nhất) cho bảng tổng hợp cuối và gắn Comment phiên bản tờ khai đã chọn tại tiêu đề năm trên `KKTNDN` / `BCHDKD`.

## 2. Quy Chuẩn Trình Bày & Đối Chiếu Trên Sheet Năm
- **Định dạng GTGT:** Không tô đậm tràn lan toàn bộ bảng; chỉ tô đậm các dòng chỉ tiêu tổng/chính (`{1, 4, 11, 14, 18, 22, 23, 24}`).
- **Đối chiếu Tờ khai vs Tổng quan (`TQ`):** Luôn đặt cột `F` (Số phát sinh trên `TQ`), cột `G` (`=C-F` chênh lệch) và Comment giải trình cho đủ 3 sắc thuế chính:
  - GTGT: Dòng `23` (`C23`, `F23`, `G23`)
  - TNDN: Dòng `76` (`C76`, `F76`, `G76`)
  - TNCN: Dòng `90` (`C90`, `F90`, `G90`)

## 3. Bảo Mật & Ẩn Danh Dữ Liệu Doanh Nghiệp (Bắt Buộc Trước Khi Push GitHub)
- **Tuyệt đối không** đưa tên doanh nghiệp thật, mã số thuế (MST) thật, số quyết định/chứng từ thật, đường dẫn thư mục chứa tên công ty thật, hay file Excel dữ liệu thật vào mã nguồn (`.py`), tài liệu (`.md`), unit test (`test_app.py`) hoặc lịch sử commit GitHub.
- Mọi ví dụ minh họa và unit test chỉ dùng dữ liệu giả lập (VD: `MST 0100000001`, `CTY A`).
- Luôn chạy toàn bộ bộ kiểm thử `python test_app.py` (100% PASS) trước khi commit & push.

## 4. Bài Học Kỹ Thuật Bắt Buộc Ghi Nhớ (Không Lặp Lại)
- **Định dạng thuần đen – trắng cho sheet copy ra Biên bản Word (`BaoCao_BienBan`):** Tuyệt đối không tô màu nền (`PatternFill(fill_type=None)`) và không dùng chữ màu; chỉ dùng chữ đen `Times New Roman`, viền đen mảnh, in đậm tiêu đề và dòng tổng để copy-paste trực tiếp vào file Word Biên bản không bị vỡ định dạng.
- **Ưu tiên công thức Excel động thay vì số tĩnh Python trên sheet đối chiếu (`DoiChieu_QHNS`):** Cột `F` (`Chênh lệch làm tròn Tờ khai vs TQ`) bắt buộc dùng công thức động (`='YYYY'!G23`, `='YYYY'!G76`, `='YYYY'!G90`), không ghi số tĩnh tính tay bằng Python. Đồng thời `_eval_simple_formula` phải hỗ trợ đệ quy công thức lồng trong `SUM(...)`, `IF(...>0, ROUND(...))` và độ sâu `> 20` để không bao giờ lệch ảo khi file Excel mới xuất chưa mở lưu bằng MS Excel.
