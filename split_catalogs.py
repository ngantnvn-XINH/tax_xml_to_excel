# -*- coding: utf-8 -*-
"""
Danh mục chỉ tiêu chuẩn tách riêng theo Thông tư 156/2013/TT-BTC (và TT151/2014/TT-BTC, áp dụng <= 2020)
và Thông tư 80/2021/TT-BTC (áp dụng >= 2021) cho:
- Tờ khai Quyết toán thuế TNDN (Mẫu 03/TNDN) -> Sheet 'KKTNDN'
- Phụ lục Kết quả hoạt động sản xuất kinh doanh (Mẫu 03-1A/TNDN) -> Sheet 'BCHDKD'
"""

from typing import List, Tuple

# ==============================================================================
# 1. TỜ KHAI QUYẾT TOÁN THUẾ TNDN (03/TNDN) - THÔNG TƯ 156 / 151 (CÁC NĂM <= 2020, maTKhai='03')
# Khớp 100% cấu trúc thẻ XML <CTieuTKhaiChinh> của mẫu 03/TNDN (TT151/2014/TT-BTC):
# B8 = B9 + B10 + B11 (3 mục giảm); B12 = Tổng TNCT; B13 = TNCT HĐ SXKD; B14 = TNCT BĐS;
# C1..C16; D, D1..D3; E, E1..E3; G, G1..G3; H; I.
# ==============================================================================
KKTNDN_156_ROWS: List[Tuple[str, str, str]] = [
    ("A", "Kết quả kinh doanh ghi nhận theo báo cáo tài chính", "A"),
    ("1", "Tổng lợi nhuận kế toán trước thuế thu nhập doanh nghiệp", "A1"),
    ("B", "Xác định thu nhập chịu thuế theo Luật thuế thu nhập doanh nghiệp", "B"),
    (
        "1",
        "Điều chỉnh tăng tổng lợi nhuận trước thuế thu nhập doanh nghiệp (B1=B2+B3+B4+B5+B6+B7)",
        "B1",
    ),
    ("1.1", "Các khoản điều chỉnh tăng doanh thu", "B2"),
    ("1.2", "Chi phí của phần doanh thu điều chỉnh giảm", "B3"),
    ("1.3", "Các khoản chi không được trừ khi xác định thu nhập chịu thuế", "B4"),
    ("1.4", "Thuế thu nhập đã nộp cho phần thu nhập nhận được ở nước ngoài", "B5"),
    (
        "1.5",
        "Điều chỉnh tăng lợi nhuận do xác định giá thị trường đối với giao dịch liên kết",
        "B6",
    ),
    ("1.6", "Các khoản điều chỉnh làm tăng lợi nhuận trước thuế khác", "B7"),
    (
        "2",
        "Điều chỉnh giảm tổng lợi nhuận trước thuế thu nhập doanh nghiệp (B8=B9+B10+B11)",
        "B8",
    ),
    ("2.1", "Giảm trừ các khoản doanh thu đã tính thuế năm trước", "B9"),
    ("2.2", "Chi phí của phần doanh thu điều chỉnh tăng", "B10"),
    ("2.3", "Các khoản điều chỉnh làm giảm lợi nhuận trước thuế khác", "B11"),
    ("3", "Tổng thu nhập chịu thuế (B12=A1+B1-B8)", "B12"),
    ("3.1", "Thu nhập chịu thuế từ hoạt động sản xuất kinh doanh", "B13"),
    ("3.2", "Thu nhập chịu thuế từ hoạt động chuyển nhượng BĐS", "B14"),
    (
        "C",
        "Thuế thu nhập doanh nghiệp (TNDN) phải nộp từ hoạt động sản xuất kinh doanh",
        "",
    ),
    ("1", "Thu nhập chịu thuế (C1 = B13)", "C1"),
    ("2", "Thu nhập miễn thuế", "C2"),
    ("3", "Chuyển lỗ và bù trừ lãi, lỗ (C3=C3a+C3b)", "C3"),
    ("", "Trong đó:", ""),
    ("3.1", "+ Lỗ từ hoạt động SXKD được chuyển trong kỳ", "C3a"),
    (
        "3.2",
        "+ Lỗ từ chuyển nhượng BĐS được bù trừ với lãi của hoạt động SXKD",
        "C3b",
    ),
    ("4", "Thu nhập tính thuế (TNTT) (C4=C1-C2-C3)", "C4"),
    ("5", "Trích lập quỹ khoa học công nghệ (nếu có)", "C5"),
    (
        "6",
        "TNTT sau khi đã trích lập quỹ khoa học công nghệ (C6=C4-C5=C7+C8+C9)",
        "C6",
    ),
    ("", "Trong đó:", ""),
    ("6.1", "+ Tính theo thuế suất 22%", "C7"),
    ("6.2", "+ Tính theo thuế suất 20%", "C8"),
    ("6.3", "+ Tính theo thuế suất không ưu đãi khác", "C9"),
    ("6.4", "+ Thuế suất không ưu đãi khác (%)", "C9a"),
    (
        "7",
        "Thuế TNDN từ hoạt động SXKD tính theo thuế suất không ưu đãi (C10=(C7 x 22%) + (C8 x 20%) + (C9 x C9a))",
        "C10",
    ),
    ("8", "Thuế TNDN chênh lệch do áp dụng mức thuế suất ưu đãi", "C11"),
    ("9", "Thuế TNDN được miễn, giảm trong kỳ", "C12"),
    ("", "Trong đó:", ""),
    ("9.1", "+ Số thuế được miễn, giảm theo Hiệp định", "C13"),
    ("9.2", "+ Số thuế được miễn, giảm không theo Luật thuế TNDN", "C14"),
    ("10", "Thuế thu nhập đã nộp ở nước ngoài được trừ trong kỳ tính thuế", "C15"),
    (
        "11",
        "Thuế TNDN phải nộp của hoạt động sản xuất kinh doanh (C16=C10-C11-C12-C15)",
        "C16",
    ),
    ("D", "Thuế TNDN phải nộp từ hoạt động chuyển nhượng BĐS (D=D1+D2+D3)", "D"),
    ("1", "Thuế TNDN phải nộp của hoạt động chuyển nhượng BĐS", "D1"),
    (
        "2",
        "Thuế TNDN phải nộp của hoạt động chuyển nhượng cơ sở hạ tầng, nhà có thu tiền theo tiến độ",
        "D2",
    ),
    ("3", "Thuế TNDN phải nộp khác (nếu có)", "D3"),
    ("E", "Số thuế TNDN phải nộp quyết toán trong kỳ (E=E1+E2+E3)", "E"),
    ("1", "Thuế TNDN của hoạt động sản xuất kinh doanh (E1 = C16)", "E1"),
    ("2", "Thuế TNDN từ hoạt động chuyển nhượng bất động sản (E2 = D)", "E2"),
    ("3", "Thuế TNDN phải nộp khác (nếu có)", "E3"),
    ("G", "Số thuế TNDN đã tạm nộp trong năm (G=G1+G2+G3)", "G"),
    ("1", "Thuế TNDN đã tạm nộp của hoạt động sản xuất kinh doanh", "G1"),
    ("2", "Thuế TNDN đã tạm nộp của hoạt động chuyển nhượng BĐS", "G2"),
    (
        "3",
        "Thuế TNDN đã tạm nộp các kỳ trước và trong năm quyết toán của hoạt động chuyển nhượng cơ sở hạ tầng, nhà có thu tiền theo tiến độ",
        "G3",
    ),
    ("H", "Chênh lệch giữa số thuế phải nộp và số thuế đã tạm nộp trong năm (H=E-G)", "H"),
    (
        "I",
        "Số thuế TNDN còn phải nộp đến thời hạn nộp hồ sơ khai quyết toán thuế (I=E-G)",
        "I",
    ),
]


# ==============================================================================
# 2. TỜ KHAI QUYẾT TOÁN THUẾ TNDN (03/TNDN) - THÔNG TƯ 80 (CÁC NĂM >= 2021, maTKhai='892')
# Khớp 100% cấu trúc thẻ XML <CTieuTKhaiChinh> của mẫu 03/TNDN (TT80/2021/TT-BTC)
# ==============================================================================
KKTNDN_80_ROWS: List[Tuple[str, str, str]] = [
    ("A", "Kết quả kinh doanh ghi nhận theo báo cáo tài chính", "A"),
    ("1", "Tổng lợi nhuận kế toán trước thuế thu nhập doanh nghiệp", "A1"),
    ("B", "Xác định thu nhập chịu thuế theo Luật thuế thu nhập doanh nghiệp", "B"),
    (
        "1",
        "Điều chỉnh tăng tổng lợi nhuận trước thuế thu nhập doanh nghiệp (B1= B2+B3+B4+B5+B6+B7)",
        "B1",
    ),
    ("1.1", "Các khoản điều chỉnh tăng doanh thu", "B2"),
    ("1.2", "Chi phí của phần doanh thu điều chỉnh giảm", "B3"),
    ("1.3", "Các khoản chi không được trừ khi xác định thu nhập chịu thuế", "B4"),
    ("1.4", "Thuế thu nhập đã nộp cho phần thu nhập nhận được ở nước ngoài", "B5"),
    (
        "1.5",
        "Điều chỉnh tăng lợi nhuận do xác định giá thị trường đối với giao dịch liên kết",
        "B6",
    ),
    ("1.6", "Các khoản điều chỉnh làm tăng lợi nhuận trước thuế khác", "B7"),
    (
        "2",
        "Điều chỉnh giảm tổng lợi nhuận trước thuế thu nhập doanh nghiệp (B8=B9+B10+B11+B12)",
        "B8",
    ),
    ("2.1", "Giảm trừ các khoản doanh thu đã tính thuế năm trước", "B9"),
    ("2.2", "Chi phí của phần doanh thu điều chỉnh tăng", "B10"),
    (
        "2.3",
        "Chi phí lãi vay không được trừ kỳ trước được chuyển sang kỳ này của doanh nghiệp có giao dịch liên kết",
        "B11",
    ),
    ("2.4", "Các khoản điều chỉnh làm giảm lợi nhuận trước thuế khác", "B12"),
    ("3", "Tổng thu nhập chịu thuế (B13=A1+B1-B8)", "B13"),
    ("3.1", "Thu nhập chịu thuế từ hoạt động sản xuất kinh doanh", "B14"),
    ("3.2", "Thu nhập chịu thuế từ hoạt động chuyển nhượng BĐS", "B15"),
    (
        "C",
        "Thuế thu nhập doanh nghiệp (TNDN) phải nộp từ hoạt động sản xuất kinh doanh",
        "",
    ),
    ("1", "Thu nhập chịu thuế (C1 = B14)", "C1"),
    ("2", "Thu nhập miễn thuế", "C2"),
    ("", "Loại thu nhập miễn thuế", "ten_loaiMien"),
    ("3", "Chuyển lỗ và bù trừ lãi, lỗ (C3=C3a+C3b)", "C3"),
    ("", "Trong đó:", ""),
    ("3.1", "+ Lỗ từ hoạt động SXKD được chuyển trong kỳ", "C3a"),
    (
        "3.2",
        "+ Lỗ từ chuyển nhượng BĐS được bù trừ với lãi của hoạt động SXKD",
        "C3b",
    ),
    ("4", "Thu nhập tính thuế (TNTT) (C4=C1-C2-C3)", "C4"),
    ("5", "Trích lập quỹ khoa học công nghệ (nếu có)", "C5"),
    (
        "6",
        "TNTT sau khi đã trích lập quỹ khoa học công nghệ (C6=C4-C5=C7+C8)",
        "C6",
    ),
    ("", "Trong đó:", ""),
    ("6.1", "+ Thu nhập tính thuế áp dụng thuế suất 20%", "C7"),
    ("6.2", "+ Thu nhập tính thuế tính theo thuế suất không ưu đãi khác", "C8"),
    ("6.3", "+ Thuế suất không ưu đãi khác (%)", "C8a"),
    (
        "7",
        "Thuế TNDN từ hoạt động SXKD tính theo thuế suất không ưu đãi(C9 =(C7 x 20%) + (C8 x C8a))",
        "C9",
    ),
    (
        "8",
        "Thuế TNDN được ưu đãi theo Luật thuế TNDN(C10 = C11 + C12 + C13)",
        "C10",
    ),
    ("", "Trong đó:", ""),
    ("8.1", "+ Thuế TNDN chênh lệch do áp dụng mức thuế suất ưu đãi", "C11"),
    ("8.2", "+ Thuế TNDN được miễn trong kỳ", "C12"),
    ("8.3", "+ Thuế TNDN được giảm trong kỳ", "C13"),
    ("9", "Thuế TNDN được miễn, giảm theo Hiệp định thuế", "C14"),
    ("10", "Thuế TNDN được miễn, giảm theo từng thời kỳ", "C15"),
    ("11", "Thuế thu nhập đã nộp ở nước ngoài được trừ trong kỳ tính thuế", "C16"),
    (
        "12",
        "Thuế TNDN phải nộp của hoạt động sản xuất kinh doanh(C17=C9-C10-C14-C15-C16)",
        "C17",
    ),
    ("D", "Thuế TNDN phải nộp từ hoạt động chuyển nhượng BĐS", "D"),
    ("1", "Thu nhập chịu thuế (D1 = B15)", "D1"),
    ("2", "Lỗ từ hoạt động chuyển nhượng BĐS được chuyển trong kỳ", "D2"),
    ("3", "Thu nhập tính thuế (D3=D1-D2)", "D3"),
    ("4", "Trích lập quỹ khoa học công nghệ (nếu có)", "D4"),
    ("5", "TNTT sau khi đã trích lập quỹ khoa học công nghệ (D5=D3-D4)", "D5"),
    ("6", "Thuế TNDN phải nộp của hoạt động chuyển nhượng BĐS trong kỳ", "D6"),
    (
        "7",
        "Thuế TNDN chênh lệch do áp dụng mức thuế suất ưu đãi đối với thu nhập từ thực hiện dự án đầu tư - kinh doanh nhà ở xã hội để bán, cho thuê, cho thuê mua",
        "D7",
    ),
    (
        "8",
        "Thuế TNDN của hoạt động chuyển nhượng BĐS còn phải nộp kỳ này (D8=D6-D7)",
        "D8",
    ),
    ("E", "Số thuế TNDN phải nộp quyết toán trong kỳ (E=E1+E2+E5)", "E"),
    ("1", "Thuế TNDN của hoạt động sản xuất kinh doanh", "E1"),
    ("2", "Thuế TNDN từ hoạt động chuyển nhượng bất động sản (E2=E3+E4)", "E2"),
    ("2.1", "Thuế TNDN từ hoạt động chuyển nhượng bất động sản", "E3"),
    (
        "2.2",
        "Thuế TNDN từ hoạt động chuyển nhượng cơ sở hạ tầng, nhà có thu tiền theo tiến độ",
        "E4",
    ),
    ("3", "Thuế TNDN phải nộp khác (nếu có)", "E5"),
    ("3.1", "Trong đó thuế TNDN từ xử lý Quỹ phát triển khoa học công nghệ", "E6"),
    ("G", "Số thuế TNDN đã tạm nộp (G=G1+G2+G3+G4+G5)", "G"),
    ("1", "Thuế TNDN đã tạm nộp của hoạt động sản xuất kinh doanh", "thueTNDN_tamnop_sxkd"),
    ("1.1", "Thuế TNDN nộp thừa kỳ trước chuyển sang kỳ này", "G1"),
    ("1.2", "Thuế TNDN đã tạm nộp trong năm", "G2"),
    ("2", "Thuế TNDN đã tạm nộp của hoạt động chuyển nhượng BĐS", "thueTNDN_tamnop_cnbds"),
    (
        "2.1",
        "Thuế TNDN nộp thừa kỳ trước chuyển sang kỳ này của hoạt động chuyển nhượng BĐS",
        "G3",
    ),
    (
        "2.2",
        "Thuế TNDN đã tạm nộp trong năm của hoạt động chuyển nhượng BĐS",
        "G4",
    ),
    (
        "2.3",
        "Thuế TNDN đã tạm nộp các kỳ trước và trong năm quyết toán của hoạt động chuyển nhượng cơ sở hạ tầng, nhà có thu tiền theo tiến độ",
        "G5",
    ),
    ("H", "Chênh lệch giữa số thuế phải nộp và số thuế đã tạm nộp", "H"),
    (
        "1",
        "Chênh lệch giữa số thuế phải nộp và số thuế đã tạm nộp trong năm của hoạt động sản xuất kinh doanh (H1=E1+E5-G2)",
        "H1",
    ),
    (
        "2",
        "Chênh lệch giữa số thuế phải nộp và số thuế đã tạm nộp trong năm của hoạt động chuyển nhượng BĐS (H2=E3-G4)",
        "H2",
    ),
    (
        "3",
        "Chênh lệch giữa số thuế phải nộp và số thuế đã tạm nộp của hoạt động chuyển nhượng cơ sở hạ tầng, nhà có thu tiền theo tiến độ (H3=E4-G5)",
        "H3",
    ),
    (
        "I",
        "Số thuế TNDN còn phải nộp đến thời hạn nộp hồ sơ khai quyết toán thuế (I=E-G=I1+I2)",
        "I",
    ),
    (
        "1",
        "Thuế TNDN còn phải nộp của hoạt động sản xuất kinh doanh",
        "I1=E1+E5-G1-G2",
    ),
    (
        "2",
        "Thuế TNDN còn phải nộp của hoạt động chuyển nhượng BĐS",
        "I2=E2-G3-G4-G5",
    ),
]


# ==============================================================================
# 3. PHỤ LỤC 03-1A/TNDN (KẾT QUẢ HOẠT ĐỘNG SXKD - SHEET 'BCHDKD')
# ==============================================================================
KQKD_156_ROWS: List[Tuple[str, str, str]] = [
    ("1", "Doanh thu bán hàng và cung cấp dịch vụ.", "[01]"),
    ("", "Trong đó: - Doanh thu bán hàng hoá, dịch vụ xuất khẩu", "[02]"),
    ("2", "Các khoản giảm trừ doanh thu ([03]=[04]+[05]+[06]+[07])", "[03]"),
    ("a", "Chiết khấu thương mại", "[04]"),
    ("b", "Giảm giá hàng bán", "[05]"),
    ("c", "Giá trị hàng bán bị trả lại", "[06]"),
    (
        "d",
        "Thuế tiêu thụ đặc biệt, thuế xuất khẩu, thuế giá trị gia tăng theo phương pháp trực tiếp phải nộp",
        "[07]",
    ),
    ("3", "Doanh thu hoạt động tài chính", "[08]"),
    (
        "4",
        "Chi phí sản xuất, kinh doanh hàng hoá, dịch vụ ([09]=[10]+[11]+[12])",
        "[09]",
    ),
    ("a", "Giá vốn hàng hóa", "[10]"),
    ("b", "Chi phí bán hàng", "[11]"),
    ("c", "Chi phí quản lý doanh nghiệp", "[12]"),
    ("5", "Chi phí tài chính", "[13]"),
    ("", "Trong đó: Chi phí lãi tiền vay dùng cho sản xuất, kinh doanh", "[14]"),
    (
        "6",
        "Lợi nhuận thuần từ hoạt động kinh doanh ([15]=[01]-[03]+[08]-[09]-[13])",
        "[15]",
    ),
    ("7", "Thu nhập khác", "[16]"),
    ("8", "Chi phí khác", "[17]"),
    ("9", "Lợi nhuận khác ([18] = [16] - [17])", "[18]"),
    (
        "10",
        "Tổng lợi nhuận kế toán trước thuế thu nhập doanh nghiệp ([19] = [15] + [18])",
        "[19]",
    ),
]

KQKD_80_ROWS: List[Tuple[str, str, str]] = [
    ("1", "Doanh thu bán hàng và cung cấp dịch vụ", "[04]"),
    ("", "Trong đó: - Doanh thu bán hàng hoá, dịch vụ xuất khẩu", "[05]"),
    ("2", "Các khoản giảm trừ doanh thu ([06]=[07]+[08]+[09])", "[06]"),
    ("a", "Chiết khấu thương mại", "[07]"),
    ("b", "Giảm giá hàng bán", "[08]"),
    ("c", "Giá trị hàng bán bị trả lại", "[09]"),
    ("3", "Doanh thu hoạt động tài chính", "[10]"),
    ("", "Trong đó: Doanh thu từ lãi tiền gửi", "[11]"),
    (
        "4",
        "Chi phí sản xuất, kinh doanh hàng hoá, dịch vụ ([12]=[13]+[14]+[15])",
        "[12]",
    ),
    ("a", "Giá vốn hàng bán", "[13]"),
    ("b", "Chi phí bán hàng", "[14]"),
    ("c", "Chi phí quản lý doanh nghiệp", "[15]"),
    ("5", "Chi phí tài chính", "[16]"),
    ("", "Trong đó: Chi phí lãi tiền vay", "[17]"),
    (
        "6",
        "Lợi nhuận thuần từ hoạt động kinh doanh([18]=[04]-[06]+[10]-[12]-[16])",
        "[18]",
    ),
    ("7", "Thu nhập khác", "[19]"),
    ("8", "Chi phí khác", "[20]"),
    ("9", "Lợi nhuận khác ([21]=[19]-[20])", "[21]"),
    (
        "10",
        "Tổng lợi nhuận kế toán trước thuế thu nhập doanh nghiệp ([22]=[18]+[21])",
        "[22]",
    ),
]
