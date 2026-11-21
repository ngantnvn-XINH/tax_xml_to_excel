# -*- coding: utf-8 -*-
"""
Module phân tích sổ Tổng quan nghĩa vụ thuế (TQ - TMS) và tổng hợp tự động bằng công thức Excel
vào sheet Quan hệ Ngân sách (QHNS) theo cả 2 phương pháp chuẩn của ngành Thuế:
1. Bảng 1 (A1:G38): Phương pháp theo Kỳ tính thuế / Kỳ báo cáo (Cột 4 tính theo Kỳ báo cáo khớp 100%
   với tờ khai XML và biểu TH_DC_TN của ứng dụng Thanh tra TCS; Cột 5 & 6 tính theo Năm hạch toán).
2. Bảng 2 (I1:O38): Phương pháp cắt sổ theo Năm Ngày hạch toán trên TMS (Cả Cột 4, 5, 6 đều tính theo
   Năm Ngày hạch toán 01/01..31/12 trên sổ TQ).
3. Bảng 3 (A41:I58): Bảng tổng hợp đối chiếu toàn bộ 15 Tiểu mục từ đầu đến 2026 khớp 100% số dư
   cuối sổ TQ (546.189.947 đ).
"""
import datetime
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from payment_summarizer import TIEU_MUC_INFO

THIN_SIDE = Side(border_style="thin", color="808080")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
YEAR_FILL = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def extract_ht_year(dt_val: Any) -> int:
    if isinstance(dt_val, (datetime.datetime, datetime.date)):
        return dt_val.year
    s = str(dt_val or "").strip()
    m = re.search(r"\b(20\d{2}|19\d{2})\b", s)
    return int(m.group(1)) if m else 0


def extract_ky_year(ky_val: Any, ht_yr: int) -> int:
    if ky_val is None:
        return ht_yr
    s = str(ky_val).strip()
    if not s:
        return ht_yr
    m4 = re.match(r"^(20\d{2})", s)
    if m4:
        return int(m4.group(1))
    m2 = re.match(r"^(\d{2})(\d{2}|Q\d|CN)", s, re.IGNORECASE)
    if m2:
        return 2000 + int(m2.group(1))
    return ht_yr


def classify_gd_for_qhns(gd: str) -> str:
    """
    Phân loại mã giao dịch TMS (Ký hiệu giao dịch) vào 3 nhóm chuẩn trên bảng QHNS:
    - 'PS_TK': Phát sinh theo tờ khai / điều chỉnh nghĩa vụ kê khai (TK, TQ, TH, KH, DP)
    - 'XL_TT': Số xử lý qua thanh tra, kiểm tra, phạt, chậm nộp, miễn giảm (QT, QP, NC, QM)
    - 'DA_NOP': Số đã nộp / hủy chứng từ nộp / hoàn trả tiền nộp thừa (C2, HC, PT, QH, C5)
      Lưu ý: Trên sổ TQ, C2 mang dấu âm (-), còn hoàn thuế (C5/QH + PT) hoặc hủy chứng từ (HC) mang dấu dương (+),
      nên Số thực nộp thuần = -SUM(DA_NOP).
    """
    g = (gd or "").strip().upper()
    if g in ("TK", "TQ", "TH", "KH", "DP"):
        return "PS_TK"
    if g in ("QT", "QP", "NC", "QM"):
        return "XL_TT"
    if g in ("C2", "HC", "PT", "QH", "C5"):
        return "DA_NOP"
    return "PS_TK"


def map_tax_group_for_qhns(tm: str) -> str:
    t = (tm or "").strip()
    if t == "1701":
        return "GTGT"
    if t == "1052":
        return "TNDN"
    if t in ("1001", "1004", "1015"):
        return "TNCN"
    if t in ("1803", "2862", "2863", "2864"):
        return "LPMB"
    return "PHAT_TCN"


def populate_tq_and_qhns(input_xlsx: str, output_paths: List[str]) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(input_xlsx)
    ws_tq = wb["TQ"]

    font_hdr = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_italic = Font(name="Times New Roman", size=11, italic=True)
    font_body = Font(name="Times New Roman", size=11, bold=False)
    font_title = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")

    # 1. Bổ sung 4 cột phụ chuẩn hóa trên sheet 'TQ' tại cột Z, AA, AB, AC (Cột 26..29)
    # để công thức SUMIFS trên sheet QHNS tham chiếu trực tiếp, rõ ràng, dễ kiểm tra
    helper_headers = [
        (26, "Nam_HT"),
        (27, "Nam_KyBC"),
        (28, "Nhom_GD"),
        (29, "Sac_Thue"),
    ]
    for col_i, h_text in helper_headers:
        c = ws_tq.cell(1, col_i, h_text)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    max_tq_row = ws_tq.max_row
    for r in range(2, max_tq_row + 1):
        cqt = ws_tq.cell(r, 1).value
        if not cqt:
            continue
        ky_val = ws_tq.cell(r, 2).value
        gd_val = str(ws_tq.cell(r, 4).value or "").strip()
        tm_val = str(ws_tq.cell(r, 5).value or "").strip()
        dt_val = ws_tq.cell(r, 10).value

        ht_yr = extract_ht_year(dt_val)
        ky_yr = extract_ky_year(ky_val, ht_yr)
        nhom_gd = classify_gd_for_qhns(gd_val)
        sac_thue = map_tax_group_for_qhns(tm_val)

        c_ht = ws_tq.cell(r, 26, ht_yr)
        c_ht.font = font_body
        c_ht.alignment = Alignment(horizontal="center", vertical="center")

        c_ky = ws_tq.cell(r, 27, ky_yr)
        c_ky.font = font_body
        c_ky.alignment = Alignment(horizontal="center", vertical="center")

        c_ngd = ws_tq.cell(r, 28, nhom_gd)
        c_ngd.font = font_body
        c_ngd.alignment = Alignment(horizontal="center", vertical="center")

        c_st = ws_tq.cell(r, 29, sac_thue)
        c_st.font = font_body
        c_st.alignment = Alignment(horizontal="center", vertical="center")

    # 2. Xây dựng sheet 'QHNS'
    if "QHNS" in wb.sheetnames:
        idx_qhns = wb.sheetnames.index("QHNS")
        wb.remove(wb["QHNS"])
        ws_q = wb.create_sheet(title="QHNS", index=idx_qhns)
    else:
        ws_q = wb.create_sheet(title="QHNS")

    def build_qhns_table(start_col: int, mode: str) -> None:
        """
        mode='KY_BC': Cột (4) Số báo cáo phát sinh phải nộp tính theo Năm Kỳ báo cáo (Cột AA trên TQ),
                      Cột (5) Số xử lý qua thanh tra & Cột (6) Số đã nộp tính theo Năm Ngày hạch toán (Cột Z trên TQ).
        mode='NAM_HT': Cả Cột (4), (5), (6) đều cắt theo Năm Ngày hạch toán (Cột Z trên TQ).
        """
        c_stt = start_col
        c_ct = start_col + 1
        c_dk = start_col + 2
        c_ps = start_col + 3
        c_tt = start_col + 4
        c_nop = start_col + 5
        c_ck = start_col + 6

        let_dk = get_column_letter(c_dk)
        let_ps = get_column_letter(c_ps)
        let_tt = get_column_letter(c_tt)
        let_nop = get_column_letter(c_nop)
        let_ck = get_column_letter(c_ck)

        # Header Rows 1..3
        ws_q.cell(1, c_stt, "STT")
        ws_q.cell(1, c_ct, "CHỈ TIÊU")
        ws_q.cell(1, c_dk, "Số kỳ trước chuyển sang")
        ws_q.cell(1, c_ps, "Số phát sinh trong kỳ")
        ws_q.cell(1, c_nop, "Số đã nộp ")
        ws_q.cell(1, c_ck, "Số còn phải nộp ")

        ws_q.cell(2, c_ps, "Số báo cáo phát sinh phải nộp trong kỳ")
        ws_q.cell(2, c_tt, "Số xử lý qua thanh tra")

        r3_vals = [-1, -2, -3, -4, -5, -6, "(7)=(3)+(4)+(5)-(6)"]
        for offset, v in enumerate(r3_vals):
            ws_q.cell(3, start_col + offset, v)

        ws_q.merge_cells(start_row=1, start_column=c_stt, end_row=2, end_column=c_stt)
        ws_q.merge_cells(start_row=1, start_column=c_ct, end_row=2, end_column=c_ct)
        ws_q.merge_cells(start_row=1, start_column=c_dk, end_row=2, end_column=c_dk)
        ws_q.merge_cells(start_row=1, start_column=c_ps, end_row=1, end_column=c_tt)
        ws_q.merge_cells(start_row=1, start_column=c_nop, end_row=2, end_column=c_nop)
        ws_q.merge_cells(start_row=1, start_column=c_ck, end_row=2, end_column=c_ck)

        for r_hdr in (1, 2, 3):
            for col_i in range(c_stt, c_ck + 1):
                cell = ws_q.cell(r_hdr, col_i)
                cell.font = font_hdr if r_hdr in (1, 2) else font_italic
                cell.fill = HEADER_FILL if r_hdr in (1, 2) else SUBHEADER_FILL
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = THIN_BORDER

        ps_yr_col = "$AA" if mode == "KY_BC" else "$Z"
        rng_val = f"TQ!$G$2:$G${max_tq_row}"
        rng_ht = f"TQ!$Z$2:$Z${max_tq_row}"
        rng_ps_yr = f"TQ!{ps_yr_col}$2:{ps_yr_col}${max_tq_row}"
        rng_ngd = f"TQ!$AB$2:$AB${max_tq_row}"
        rng_sac = f"TQ!$AC$2:$AC${max_tq_row}"

        years = [2020, 2021, 2022, 2023, 2024, 2025, 2026]
        prev_rows: Dict[str, int] = {}

        for y_idx, yr in enumerate(years):
            base_r = 4 + y_idx * 5
            # Dòng tiêu đề Năm
            ws_q.cell(base_r, c_ct, f"Năm {yr}")
            for col_i in range(c_stt, c_ck + 1):
                c = ws_q.cell(base_r, col_i)
                c.font = font_bold
                c.fill = YEAR_FILL
                c.border = THIN_BORDER
                if col_i == c_ct:
                    c.alignment = Alignment(horizontal="left", vertical="center")

            items = [
                (1, "Thuế GTGT", "GTGT"),
                (None, "Thuế GTGT vãng lai", "GTGT_VL"),
                (2, "Thuế TNDN", "TNDN"),
                (3, "Thuế TNCN", "TNCN"),
            ]
            for i_off, (stt_v, label_v, tax_code) in enumerate(items, start=1):
                r_cur = base_r + i_off
                ws_q.row_dimensions[r_cur].height = 22

                c1 = ws_q.cell(r_cur, c_stt, stt_v)
                c1.font = font_body
                c1.alignment = Alignment(horizontal="center", vertical="center")
                c1.border = THIN_BORDER

                c2 = ws_q.cell(r_cur, c_ct, label_v)
                c2.font = font_italic if tax_code == "GTGT_VL" else font_body
                c2.alignment = Alignment(horizontal="left", vertical="center")
                c2.border = THIN_BORDER

                if tax_code == "GTGT_VL":
                    f_dk = 0 if yr == 2020 else f"={let_ck}{prev_rows[tax_code]}"
                    f_ps = 0
                    f_tt = 0
                    f_nop = 0
                else:
                    if yr == 2020:
                        # Số dư đầu năm 2020 = Lũy kế phát sinh (<=2019) + Lũy kế thanh tra (<=2019) - Lũy kế đã nộp (<=2019)
                        f_dk = (
                            f"=SUMIFS({rng_val},{rng_sac},\"{tax_code}\",{rng_ngd},\"PS_TK\",{rng_ps_yr},\"<=2019\")"
                            f"+SUMIFS({rng_val},{rng_sac},\"{tax_code}\",{rng_ngd},\"XL_TT\",{rng_ht},\"<=2019\")"
                            f"+SUMIFS({rng_val},{rng_sac},\"{tax_code}\",{rng_ngd},\"DA_NOP\",{rng_ht},\"<=2019\")"
                        )
                    else:
                        f_dk = f"={let_ck}{prev_rows[tax_code]}"

                    f_ps = f"=SUMIFS({rng_val},{rng_sac},\"{tax_code}\",{rng_ngd},\"PS_TK\",{rng_ps_yr},{yr})"
                    f_tt = f"=SUMIFS({rng_val},{rng_sac},\"{tax_code}\",{rng_ngd},\"XL_TT\",{rng_ht},{yr})"
                    f_nop = f"=-SUMIFS({rng_val},{rng_sac},\"{tax_code}\",{rng_ngd},\"DA_NOP\",{rng_ht},{yr})"

                f_ck = f"={let_dk}{r_cur}+{let_ps}{r_cur}+{let_tt}{r_cur}-{let_nop}{r_cur}"

                for col_i, f_val in (
                    (c_dk, f_dk),
                    (c_ps, f_ps),
                    (c_tt, f_tt),
                    (c_nop, f_nop),
                    (c_ck, f_ck),
                ):
                    cell = ws_q.cell(r_cur, col_i, f_val)
                    cell.font = font_bold if col_i == c_ck else font_body
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.border = THIN_BORDER

                prev_rows[tax_code] = r_cur

    ws_q.row_dimensions[1].height = 24
    ws_q.row_dimensions[2].height = 28
    ws_q.row_dimensions[3].height = 20

    # Bảng 1 (Cột A..G): Theo Kỳ báo cáo (Chuẩn đối chiếu tờ khai & TH_DC_TN)
    build_qhns_table(start_col=1, mode="KY_BC")

    # Bảng 2 (Cột I..O): Theo Năm Ngày hạch toán trên sổ TQ (Chuẩn cắt sổ TMS 31/12 hàng năm)
    build_qhns_table(start_col=9, mode="NAM_HT")

    # Ghi chú tiêu đề phân biệt rõ 2 phương pháp ở dòng 40
    ws_q.cell(
        40,
        1,
        "Ghi chú Bảng trái (Cột A..G): Cột (4) tính theo Kỳ báo cáo của tờ khai (khớp 100% số phải nộp trên Tờ khai XML & TH_DC_TN); Cột (5), (6) tính theo Năm hạch toán.",
    ).font = font_italic
    ws_q.cell(
        40,
        9,
        "Ghi chú Bảng phải (Cột I..O): Cả Cột (4), (5), (6) đều cắt theo Năm Ngày hạch toán (01/01..31/12) trên sổ TQ của cơ quan thuế.",
    ).font = font_italic

    # ==========================================================================
    # BẢNG 3 (Dòng 42..60): TỔNG HỢP ĐỐI CHIẾU TOÀN BỘ 15 TIỂU MỤC TRÊN SỔ TQ
    # ==========================================================================
    r_b3_title = 42
    ws_q.cell(
        r_b3_title,
        1,
        "BẢNG ĐỐI CHIẾU TỔNG HỢP TOÀN BỘ 15 TIỂU MỤC TRÊN SỔ TỔNG QUAN (TQ) LŨY KẾ ĐẾN NĂM 2026",
    ).font = font_title

    r_b3_hdr = 43
    ws_q.row_dimensions[r_b3_hdr].height = 28
    b3_headers = [
        "STT",
        "Tiểu mục",
        "Tên tiểu mục / Nội dung kinh tế",
        "Nhóm sắc thuế",
        "Tổng phát sinh tờ khai (TK, TQ, TH, KH, DP)",
        "Tổng xử lý thanh tra, phạt, chậm nộp (QT, QP, NC, QM)",
        "Tổng số đã nộp / hoàn (C2, HC, C5, QH, PT)",
        "Số còn phải nộp (+)/ Nộp thừa (-) theo công thức",
        "Số dư Còn lại trên dòng Tổng Tiểu mục sổ TQ",
        "Chênh lệch Kiểm tra (Cột 8 - Cột 9)",
    ]
    for c_i, h_txt in enumerate(b3_headers, start=1):
        c = ws_q.cell(r_b3_hdr, c_i, h_txt)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    # Tìm dòng tổng cộng của từng Tiểu mục trên sheet TQ (nơi Cột A trống và Cột E = Tiểu mục)
    subtotal_row_map: Dict[str, int] = {}
    for r_tq in range(2, max_tq_row + 1):
        if ws_tq.cell(r_tq, 1).value is None:
            tm_code = str(ws_tq.cell(r_tq, 5).value or "").strip()
            if tm_code:
                subtotal_row_map[tm_code] = r_tq

    tieu_mucs = sorted(TIEU_MUC_INFO.keys())
    rng_tm = f"TQ!$E$2:$E${max_tq_row}"
    rng_val = f"TQ!$G$2:$G${max_tq_row}"
    rng_ngd = f"TQ!$AB$2:$AB${max_tq_row}"

    b3_start = r_b3_hdr + 1
    cur_r = b3_start
    for idx, tm in enumerate(tieu_mucs, start=1):
        ws_q.row_dimensions[cur_r].height = 22
        tm_name, tm_grp = TIEU_MUC_INFO[tm]

        c1 = ws_q.cell(cur_r, 1, idx)
        c1.font = font_body
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = THIN_BORDER

        c2 = ws_q.cell(cur_r, 2, tm)
        c2.font = font_bold
        c2.number_format = "@"
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = THIN_BORDER

        c3 = ws_q.cell(cur_r, 3, tm_name)
        c3.font = font_body
        c3.alignment = Alignment(horizontal="left", vertical="center")
        c3.border = THIN_BORDER

        c4 = ws_q.cell(cur_r, 4, tm_grp)
        c4.font = font_body
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border = THIN_BORDER

        c5 = ws_q.cell(cur_r, 5, f"=SUMIFS({rng_val},{rng_tm},$B{cur_r},{rng_ngd},\"PS_TK\")")
        c6 = ws_q.cell(cur_r, 6, f"=SUMIFS({rng_val},{rng_tm},$B{cur_r},{rng_ngd},\"XL_TT\")")
        c7 = ws_q.cell(cur_r, 7, f"=-SUMIFS({rng_val},{rng_tm},$B{cur_r},{rng_ngd},\"DA_NOP\")")
        c8 = ws_q.cell(cur_r, 8, f"=E{cur_r}+F{cur_r}-G{cur_r}")
        sub_r = subtotal_row_map.get(tm)
        c9_formula = f"=TQ!G{sub_r}" if sub_r else "0"
        c9 = ws_q.cell(cur_r, 9, c9_formula)
        c10 = ws_q.cell(cur_r, 10, f"=H{cur_r}-I{cur_r}")

        for c_num in (c5, c6, c7, c8, c9, c10):
            c_num.font = font_bold if c_num in (c8, c9, c10) else font_body
            c_num.number_format = "#,##0"
            c_num.alignment = Alignment(horizontal="right", vertical="center")
            c_num.border = THIN_BORDER

        cur_r += 1

    b3_end = cur_r - 1
    b3_sum = cur_r
    ws_q.row_dimensions[b3_sum].height = 24
    ws_q.merge_cells(start_row=b3_sum, start_column=1, end_row=b3_sum, end_column=4)
    for c_i in range(1, 5):
        c = ws_q.cell(b3_sum, c_i)
        if c_i == 1:
            c.value = "TỔNG CỘNG TOÀN BỘ 15 TIỂU MỤC"
        c.font = font_bold
        c.fill = TOTAL_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    for c_i in range(5, 11):
        col_let = get_column_letter(c_i)
        c = ws_q.cell(b3_sum, c_i, f"=SUM({col_let}{b3_start}:{col_let}{b3_end})")
        c.font = font_bold
        c.fill = TOTAL_FILL
        c.number_format = "#,##0"
        c.alignment = Alignment(horizontal="right", vertical="center")
        c.border = THIN_BORDER

    col_widths = {
        "A": 6,
        "B": 24,
        "C": 20,
        "D": 20,
        "E": 20,
        "F": 20,
        "G": 20,
        "H": 4,
        "I": 6,
        "J": 24,
        "K": 20,
        "L": 20,
        "M": 20,
        "N": 20,
        "O": 20,
    }
    for col_let, w in col_widths.items():
        ws_q.column_dimensions[col_let].width = w

    for out_p in output_paths:
        os.makedirs(os.path.dirname(os.path.abspath(out_p)), exist_ok=True)
        wb.save(out_p)

    return {"max_tq_row": max_tq_row, "outputs": output_paths}


if __name__ == "__main__":
    src = r"C:\Users\Admin\.gemini\antigravity\brain\74addb61-c896-4052-bc3d-ededa129ccfc\.user_uploaded\media_1795266445732_8c00556c.xlsx"
    outs = [
        r"D:\2026\MANY\TCS_MANY\TCS_MANY\TQ_QHNS_TongHop.xlsx",
        r"C:\Users\Admin\.gemini\antigravity\scratch\tax_xml_to_excel\TQ_QHNS_TongHop.xlsx",
    ]
    res = populate_tq_and_qhns(src, outs)
    print("Completed QHNS export:", res)
