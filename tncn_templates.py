# -*- coding: utf-8 -*-
"""
Cấu hình danh mục chuẩn và định nghĩa tiêu đề cột theo đúng mẫu Excel của người dùng:
- Sheet TK (Tờ khai QTT TNCN 05/QTT-TNCN)
- Sheet 05-1 (Bảng kê cá nhân có HĐLĐ)
- Sheet 05-2 (Bảng kê cá nhân không cư trú / vãng lai)
- Sheet 05-3 (Bảng kê người phụ thuộc)
"""

from typing import List, Tuple

# Tiêu đề cột và cấu trúc của Sheet 05-1: (Hàng 1, Hàng 2, Hàng 3 Tag)
COLS_05_1: List[Tuple[str, str, str]] = [
    ("STT", "", "[06]"),
    ("Họ và tên", "", "[07]"),
    ("Mã số thuế", "", "[08]"),
    ("Loại giấy tờ", "", ""),
    ("Số CMND/CCCD/SĐDCN Hộ chiếu(trường hợp chưa có MST)", "", "[09]"),
    ("Cá nhân uỷ quyền quyết toán thay", "", "[10]"),
    ("CN nước ngoài ủy quyền quyết toán dưới 12 tháng", "", "[11]"),
    ("Thu nhập chịu thuế (TNCT)", "Tổng thu nhập chịu thuế", "[12]"),
    ("", "Trong đó: TNCT tại tổ chức trước khi điều chuyển (trường hợp có đánh dấu vào chỉ tiêu [04] tại Tờ khai 05/QTT-TNCN)", "[13]"),
    ("", "Trong đó: thu nhập chịu thuế được miễn theo Hiệp định", "[14]"),
    ("", "Trong đó: thu nhập chịu thuế được miễn theo quy định của Hợp đồng dầu khí", "[15]"),
    ("Các khoản giảm trừ", "Số lượng NPT tính giảm trừ", "[16]"),
    ("", "Tổng số tiền giảm trừ gia cảnh", "[17]"),
    ("", "Từ thiện, nhân đạo, khuyến học", "[18]"),
    ("", "Bảo hiểm được trừ", "[19]"),
    ("", "Quĩ hưu trí tự nguyện được trừ", "[20]"),
    ("Thu nhập tính thuế ([21]= [12]-[14]-[15]-[17])", "", "[21]"),
    ("Số thuế TNCN đã khấu trừ ", "Tổng số thuế TNCN đã khấu trừ", "[22]"),
    ("", "Trong đó: số thuế đã khấu trừ tại tổ chức trước khi điều chuyển (trường hợp có đánh dấu vào chỉ tiêu [04] tại Tờ khai 05/QTT", "[23]"),
    ("Chi tiết kết quả quyết toán thay cho cá nhân nộp thuế", "Tổng số thuế phải nộp", "[24]"),
    ("", "Số thuế đã nộp thừa", "[25]"),
    ("", "Số thuế còn phải nộp", "[26]"),
    ("Cá nhân có số thuế được miễn do có số thuế còn phải nộp từ 50.000 đồng trở xuống", "", "[27]"),
]

# Các ô merge ở tiêu đề 05-1
MERGES_05_1 = [
    "A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:E2", "F1:F2", "G1:G2",
    "H1:K1", "L1:P1", "Q1:Q2", "R1:S1", "T1:V1", "W1:W2"
]

# Tiêu đề cột và cấu trúc của Sheet 05-2:
COLS_05_2: List[Tuple[str, str, str]] = [
    ("STT", "", "[06]"),
    ("Họ và tên", "", "[07]"),
    ("Mã số thuế", "", "[08]"),
    ("Loại giấy tờ", "", ""),
    ("Số CMND/CCCD/Hộ chiếu", "", "[09]"),
    ("Cá nhân không cư trú", "", "[10]"),
    ("Thu nhập chịu thuế (TNCT)", "Tổng số", "[11]"),
    ("", "Trong đó: TNCT từ phí mua BH nhân thọ, BH không bắt buộc khác của DN BH không thành lập tại Việt Nam cho người lao động", "[12]"),
    ("Trong đó thu nhập chịu thuế được miễn theo Hiệp định", "", "[13]"),
    ("Trong đó TN chịu thuế được miễn theo quy định của Hợp đồng dầu khí", "", "[14]"),
    ("Số thuế thu nhập cá nhân (TNCN) đã khấu trừ", "Tổng số", "[15]"),
    ("", "Trong đó: Số thuế từ phí mua BH nhân thọ, BH không bắt buộc khác của DN BH không thành lập tại Việt Nam cho người lao động", "[16]"),
]

MERGES_05_2 = [
    "A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:E2", "F1:F2",
    "G1:H1", "I1:I2", "J1:J2", "K1:L1"
]

# Tiêu đề cột và cấu trúc của Sheet 05-3:
COLS_05_3: List[Tuple[str, str, str]] = [
    ("STT", "", "[06]"),
    ("Họ và tên người nộp thuế là người lao động tại tổ chức trả thu nhập ", "", "[07]"),
    ("MST của người nộp thuế là người lao động tại tổ chức trả thu nhập", "", "[08]"),
    ("Họ và tên người phụ thuộc", "", "[09]"),
    ("Ngày sinh người phụ thuộc", "", "[10]"),
    ("MST của người phụ thuộc", "", "[11]"),
    ("Loại giấy tờ (Số CMND/ CCCD/Hộ chiếu/GKS) người phụ thuộc", "", "[12]"),
    ("Số giấy tờ ", "", "[13]"),
    ("Quan hệ với người nộp thuế ", "", "[14]"),
    ("Thời gian tính giảm trừ trong năm tính thuế", "Từ tháng", "[15]"),
    ("", "Đến tháng", "[16]"),
]

MERGES_05_3 = [
    "A1:A2", "B1:B2", "C1:C2", "D1:D2", "E1:E2", "F1:F2", "G1:G2", "H1:H2", "I1:I2",
    "J1:K1"
]
