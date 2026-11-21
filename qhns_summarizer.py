# -*- coding: utf-8 -*-
"""
Module phân tích sổ Tổng quan nghĩa vụ thuế (TQ - TMS), đối chiếu với file Số nộp (SO_NOP)
và tổng hợp tự động bằng công thức Excel SUMIFS trực tiếp theo từng mã Ký hiệu giao dịch chuẩn:
- Thuế GTGT (Tiểu mục 1701):
  + Số phát sinh (Cột 4): Ký hiệu giao dịch = "TK"
  + Số xử lý thanh tra (Cột 5): Ký hiệu giao dịch = "QT"
  + Số đã nộp (Cột 6): Ký hiệu giao dịch = "C2" (theo Năm Ngày hạch toán)
- Thuế TNDN (Tiểu mục 1052):
  + Số phát sinh (Cột 4): Ký hiệu giao dịch = "TK", "QT"
  + Số đã nộp (Cột 6): Ký hiệu giao dịch = "C2" (theo Năm Ngày hạch toán; tuyệt đối không lấy PT, QH, C5)
- Thuế TNCN (Tiểu mục 1001, 1004, 1015):
  + Số phát sinh (Cột 4): Ký hiệu giao dịch = "TK", "TQ", "TH"
  + Số đã nộp (Cột 6): Ký hiệu giao dịch = "C2" (theo Năm Ngày hạch toán)
- Số đầu kỳ năm 2020 (Cột 3):
  + Lấy các khoản phát sinh / nộp của kỳ trước năm 2020 (Kỳ báo cáo < 2020) nhưng nộp trong kỳ kiểm tra (2020-2025).
"""
import datetime
import os
import re
from typing import Any, Dict, List, Optional

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


def populate_tq_and_qhns(
    input_tq_xlsx: str,
    input_sonop_xlsx: Optional[str],
    output_paths: List[str],
) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(input_tq_xlsx)
    ws_tq = wb["TQ"]

    font_hdr = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_italic = Font(name="Times New Roman", size=11, italic=True)
    font_body = Font(name="Times New Roman", size=11, bold=False)
    font_title = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")

    # 1. Bổ sung 3 cột phụ chuẩn hóa trên sheet 'TQ' tại cột Z, AA, AB (Cột 26..28):
    # - Cột Z (26): Nam_HT (Năm của Ngày hạch toán Cột J)
    # - Cột AA (27): Nam_KyBC (Năm của Kỳ báo cáo Cột B)
    # - Cột AB (28): Sac_Thue (GTGT, TNDN, TNCN, LPMB, PHAT_TCN)
    helper_headers = [
        (26, "Nam_HT"),
        (27, "Nam_KyBC"),
        (28, "Sac_Thue"),
    ]
    for col_i, h_text in helper_headers:
        c = ws_tq.cell(1, col_i, h_text)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    # Xóa cột 29 cũ nếu có
    ws_tq.cell(1, 29, None)

    max_tq_row = ws_tq.max_row
    for r in range(2, max_tq_row + 1):
        cqt = ws_tq.cell(r, 1).value
        if not cqt:
            continue
        ky_val = ws_tq.cell(r, 2).value
        tm_val = str(ws_tq.cell(r, 5).value or "").strip()
        dt_val = ws_tq.cell(r, 10).value

        ht_yr = extract_ht_year(dt_val)
        ky_yr = extract_ky_year(ky_val, ht_yr)
        sac_thue = map_tax_group_for_qhns(tm_val)

        c_ht = ws_tq.cell(r, 26, ht_yr)
        c_ht.font = font_body
        c_ht.alignment = Alignment(horizontal="center", vertical="center")

        c_ky = ws_tq.cell(r, 27, ky_yr)
        c_ky.font = font_body
        c_ky.alignment = Alignment(horizontal="center", vertical="center")

        c_st = ws_tq.cell(r, 28, sac_thue)
        c_st.font = font_body
        c_st.alignment = Alignment(horizontal="center", vertical="center")

        ws_tq.cell(r, 29, None)

    ws_tq.column_dimensions["Z"].width = 12
    ws_tq.column_dimensions["AA"].width = 12
    ws_tq.column_dimensions["AB"].width = 14

    # 2. Nếu có file Số nộp (SO_NOP), nhúng vào sheet 'SO_NOP' để công thức đối chiếu trực tiếp
    max_sn_row = 0
    if input_sonop_xlsx and os.path.exists(input_sonop_xlsx):
        if "SO_NOP" in wb.sheetnames:
            del wb["SO_NOP"]
        wb_sn = openpyxl.load_workbook(input_sonop_xlsx, data_only=True)
        ws_sn_src = wb_sn[wb_sn.sheetnames[0]]
        ws_sn = wb.create_sheet("SO_NOP")
        max_sn_row = ws_sn_src.max_row
        for r_s in range(1, max_sn_row + 1):
            for c_s in range(1, 20):
                val = ws_sn_src.cell(r_s, c_s).value
                if r_s > 1 and c_s in (1, 2, 3, 4, 5) and val is not None:
                    val = str(val).strip()
                cell = ws_sn.cell(r_s, c_s, val)
                cell.font = font_hdr if r_s == 1 else font_body
                if r_s == 1:
                    cell.fill = HEADER_FILL
                if c_s in (7, 8, 9) and r_s > 1:
                    cell.number_format = "#,##0"
                elif c_s in (10, 11, 13, 14, 19) and r_s > 1 and val is not None:
                    cell.number_format = "dd/mm/yyyy"
            if r_s == 1:
                c_y = ws_sn.cell(1, 20, "Nam_HT")
                c_y.font = font_hdr
                c_y.fill = HEADER_FILL
                c_s_grp = ws_sn.cell(1, 21, "Sac_Thue")
                c_s_grp.font = font_hdr
                c_s_grp.fill = HEADER_FILL
            else:
                dt_v = ws_sn_src.cell(r_s, 10).value
                tm_v = str(ws_sn_src.cell(r_s, 5).value or "").strip()
                ws_sn.cell(r_s, 20, extract_ht_year(dt_v)).font = font_body
                ws_sn.cell(r_s, 21, map_tax_group_for_qhns(tm_v)).font = font_body

    # 3. Xây dựng sheet 'QHNS'
    ws_q = wb["QHNS"]
    for m_rng in list(ws_q.merged_cells.ranges):
        ws_q.unmerge_cells(str(m_rng))
    for r in range(1, ws_q.max_row + 20):
        for c in range(1, 20):
            cell = ws_q.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.border = Border()

    rng_val = f"TQ!$G$2:$G${max_tq_row}"
    rng_kh = f"TQ!$D$2:$D${max_tq_row}"
    rng_tm = f"TQ!$E$2:$E${max_tq_row}"
    rng_ht = f"TQ!$Z$2:$Z${max_tq_row}"
    rng_ky = f"TQ!$AA$2:$AA${max_tq_row}"
    rng_sac = f"TQ!$AB$2:$AB${max_tq_row}"

    def build_ps_formula(tax_code: str, yr_col_rng: str, yr_crit: Any) -> str:
        """
        Công thức Cột (4) Số báo cáo phát sinh phải nộp trong kỳ đúng theo nguyên tắc của người dùng:
        - Thuế GTGT (1701): Lấy giao dịch TK
        - Thuế TNDN (1052): Lấy giao dịch TK, QT
        - Thuế TNCN (1001, 1004, 1015): Lấy tổng giao dịch TK, TQ, TH
        """
        crit_str = f"\"{yr_crit}\"" if isinstance(yr_crit, str) else str(yr_crit)
        if tax_code == "GTGT":
            return f"=SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"TK\",{yr_col_rng},{crit_str})"
        if tax_code == "TNDN":
            return (
                f"=SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"TK\",{yr_col_rng},{crit_str})"
                f"+SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"QT\",{yr_col_rng},{crit_str})"
            )
        if tax_code == "TNCN":
            return (
                f"=SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TK\",{yr_col_rng},{crit_str})"
                f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TQ\",{yr_col_rng},{crit_str})"
                f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TH\",{yr_col_rng},{crit_str})"
            )
        return "0"

    def build_tt_formula(tax_code: str, yr_crit: Any) -> str:
        """
        Công thức Cột (5) Số xử lý qua thanh tra:
        - GTGT (1701): Lấy giao dịch QT theo Năm Ngày hạch toán
        - TNCN: Lấy giao dịch QT theo Năm Ngày hạch toán
        - TNDN: Đã gộp TK, QT ở cột (4) theo quy tắc nên cột (5) = 0
        """
        crit_str = f"\"{yr_crit}\"" if isinstance(yr_crit, str) else str(yr_crit)
        if tax_code == "GTGT":
            return f"=SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"QT\",{rng_ht},{crit_str})"
        if tax_code == "TNCN":
            return f"=SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"QT\",{rng_ht},{crit_str})"
        return "0"

    def build_nop_formula(tax_code: str, yr_crit: Any) -> str:
        """
        Công thức Cột (6) Số đã nộp:
        - Chỉ lấy giao dịch C2 (và HC nếu có) theo Năm Ngày hạch toán, đổi dấu sang dương.
        """
        crit_str = f"\"{yr_crit}\"" if isinstance(yr_crit, str) else str(yr_crit)
        if tax_code == "GTGT":
            return (
                f"=-SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"C2\",{rng_ht},{crit_str})"
                f"-SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"HC\",{rng_ht},{crit_str})"
            )
        if tax_code == "TNDN":
            return (
                f"=-SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"C2\",{rng_ht},{crit_str})"
                f"-SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"HC\",{rng_ht},{crit_str})"
            )
        if tax_code == "TNCN":
            return (
                f"=-SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"C2\",{rng_ht},{crit_str})"
                f"-SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"HC\",{rng_ht},{crit_str})"
            )
        return "0"

    def build_dk_2020_formula(tax_code: str, mode: str) -> str:
        """
        Công thức Cột (3) Số đầu kỳ năm 2020 theo đúng nguyên tắc của người dùng:
        "Số đầu kỳ là tất cả những số nộp / số phát sinh của các kỳ trước năm 2020 (Nam_KyBC <= 2019)
        nhưng lại nộp vào kỳ kiểm tra (Nam_HT từ 2020 đến 2025)".
        """
        if mode == "KY_BC":
            # Lấy tổng số C2 của các kỳ trước năm 2020 (Nam_KyBC <= 2019) nhưng nộp trong năm 2020..2025 (Nam_HT >= 2020)
            if tax_code == "GTGT":
                return f"=-SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
            if tax_code == "TNDN":
                return f"=-SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
            if tax_code == "TNCN":
                return f"=-SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
            return "0"
        else:
            # Bảng Phải (Khi Cột 4 cũng lấy theo Năm Ngày hạch toán 2020..2025):
            # Các tờ khai kỳ 1912 nộp ngày 20/01/2020 đã nằm trong Cột 4 (Nam_HT=2020),
            # nên Đầu kỳ 2020 là chênh lệch kỳ <2020 phát sinh trước 2020 nhưng nộp sau 2020 (hoặc ngược lại).
            if tax_code == "GTGT":
                # Kỳ 1901 (432.301.696) đã nộp C2 ngày 20/02/2019 nhưng TK hạch toán ngày 30/11/2020
                return (
                    f"=SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"TK\",{rng_ky},\"<=2019\",{rng_ht},\"<=2019\")"
                    f"+SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\"<=2019\")"
                )
            if tax_code == "TNDN":
                return "0"
            if tax_code == "TNCN":
                return (
                    f"=SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TK\",{rng_ht},\"<=2019\")"
                    f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TQ\",{rng_ht},\"<=2019\")"
                    f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TH\",{rng_ht},\"<=2019\")"
                    f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"C2\",{rng_ht},\"<=2019\")"
                )
            return "0"

    def build_qhns_table(start_col: int, mode: str) -> None:
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

        ps_yr_rng = rng_ky if mode == "KY_BC" else rng_ht
        years = [2020, 2021, 2022, 2023, 2024, 2025]
        prev_rows: Dict[str, int] = {}

        for y_idx, yr in enumerate(years):
            base_r = 4 + y_idx * 5
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
                        f_dk = build_dk_2020_formula(tax_code, mode)
                    else:
                        f_dk = f"={let_ck}{prev_rows[tax_code]}"

                    f_ps = build_ps_formula(tax_code, ps_yr_rng, yr)
                    f_tt = build_tt_formula(tax_code, yr)
                    f_nop = build_nop_formula(tax_code, yr)

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

    # Bảng 1 (Cột A..G): Cột (4) lấy theo Năm Kỳ báo cáo, Cột (3) đầu năm 2020 lấy số kỳ <2020 nộp trong 2020-2025, Cột (6) lấy C2 theo Năm Ngày hạch toán
    build_qhns_table(start_col=1, mode="KY_BC")

    # Bảng 2 (Cột I..O): Cả Cột (4) và Cột (6) đều lấy theo Năm Ngày hạch toán (2020..2025)
    build_qhns_table(start_col=9, mode="NAM_HT")

    ws_q.cell(
        35,
        1,
        "Ghi chú Bảng Trái (A1:G33): Cột (3) đầu năm 2020 lấy số kỳ trước 2020 nộp trong 2020-2025; Cột (4) lấy theo Năm Kỳ báo cáo (GTGT: TK | TNDN: TK, QT | TNCN: TK, TQ, TH); Cột (6) lấy C2 theo Năm Ngày hạch toán.",
    ).font = font_italic
    ws_q.cell(
        35,
        9,
        "Ghi chú Bảng Phải (I1:O33): Cả Cột (4) Phát sinh (GTGT: TK | TNDN: TK, QT | TNCN: TK, TQ, TH) và Cột (6) Số đã nộp (C2) đều lấy theo Năm Ngày hạch toán (2020-2025).",
    ).font = font_italic

    # ==========================================================================
    # BẢNG 3 (Dòng 37..47): ĐỐI CHIẾU SỐ ĐÃ NỘP (C2, HC) TRÊN FILE TQ VỚI FILE SỐ NỘP (SO_NOP)
    # ==========================================================================
    r_b3_title = 37
    ws_q.cell(
        r_b3_title,
        1,
        "BẢNG ĐỐI CHIẾU SỐ ĐÃ NỘP NGÂN SÁCH (C2, HC) THEO NĂM NGÀY HẠCH TOÁN GIỮA SỔ TỔNG QUAN (TQ) VÀ FILE SỐ NỘP (SO_NOP)",
    ).font = font_title

    r_b3_hdr = 38
    ws_q.row_dimensions[r_b3_hdr].height = 26
    b3_headers = [
        "STT",
        "Sắc thuế / Tiểu mục",
        "Nguồn dữ liệu đối chiếu",
        "Năm 2020",
        "Năm 2021",
        "Năm 2022",
        "Năm 2023",
        "Năm 2024",
        "Năm 2025",
        "Tổng cộng 2020-2025",
    ]
    for c_i, h_txt in enumerate(b3_headers, start=1):
        c = ws_q.cell(r_b3_hdr, c_i, h_txt)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    sn_val = f"SO_NOP!$G$2:$G${max_sn_row}" if max_sn_row else rng_val
    sn_kh = f"SO_NOP!$D$2:$D${max_sn_row}" if max_sn_row else rng_kh
    sn_ht = f"SO_NOP!$T$2:$T${max_sn_row}" if max_sn_row else rng_ht
    sn_sac = f"SO_NOP!$U$2:$U${max_sn_row}" if max_sn_row else rng_sac

    compare_groups = [
        (1, "Thuế GTGT (1701)", "GTGT"),
        (2, "Thuế TNDN (1052)", "TNDN"),
        (3, "Thuế TNCN (1001, 1004, 1015)", "TNCN"),
    ]
    cur_r = r_b3_hdr + 1
    for stt_g, g_label, g_code in compare_groups:
        # Dòng 1: Lấy từ TQ (C2, HC)
        r_tq_line = cur_r
        # Dòng 2: Lấy từ SO_NOP
        r_sn_line = cur_r + 1
        # Dòng 3: Chênh lệch
        r_df_line = cur_r + 2

        for r_l, src_txt in (
            (r_tq_line, "1. Lấy từ sheet TQ (Giao dịch C2)"),
            (r_sn_line, "2. Lấy từ file Số nộp (sheet SO_NOP)"),
            (r_df_line, "3. Chênh lệch (TQ - SO_NOP)"),
        ):
            ws_q.row_dimensions[r_l].height = 20
            c1 = ws_q.cell(r_l, 1, stt_g if r_l == r_tq_line else None)
            c2 = ws_q.cell(r_l, 2, g_label if r_l == r_tq_line else None)
            c3 = ws_q.cell(r_l, 3, src_txt)
            for c_i, cell in ((1, c1), (2, c2), (3, c3)):
                cell.font = font_bold if r_l == r_df_line else font_body
                if r_l == r_df_line:
                    cell.fill = SUBHEADER_FILL
                cell.alignment = Alignment(
                    horizontal="center" if c_i == 1 else "left",
                    vertical="center",
                )
                cell.border = THIN_BORDER

            for y_off, yr in enumerate([2020, 2021, 2022, 2023, 2024, 2025]):
                col_c = 4 + y_off
                col_let = get_column_letter(col_c)
                if r_l == r_tq_line:
                    f_str = f"=-SUMIFS({rng_val},{rng_sac},\"{g_code}\",{rng_kh},\"C2\",{rng_ht},{yr})"
                elif r_l == r_sn_line:
                    f_str = f"=-SUMIFS({sn_val},{sn_sac},\"{g_code}\",{sn_kh},\"C2\",{sn_ht},{yr})"
                else:
                    f_str = f"={col_let}{r_tq_line}-{col_let}{r_sn_line}"

                c_num = ws_q.cell(r_l, col_c, f_str)
                c_num.font = font_bold if r_l == r_df_line else font_body
                if r_l == r_df_line:
                    c_num.fill = SUBHEADER_FILL
                c_num.number_format = "#,##0"
                c_num.alignment = Alignment(horizontal="right", vertical="center")
                c_num.border = THIN_BORDER

            c_tot = ws_q.cell(r_l, 10, f"=SUM(D{r_l}:I{r_l})")
            c_tot.font = font_bold
            if r_l == r_df_line:
                c_tot.fill = SUBHEADER_FILL
            c_tot.number_format = "#,##0"
            c_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_tot.border = THIN_BORDER

        cur_r += 3

    col_widths = {
        "A": 6,
        "B": 26,
        "C": 28,
        "D": 20,
        "E": 20,
        "F": 20,
        "G": 20,
        "H": 18,
        "I": 18,
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

    return {"max_tq_row": max_tq_row, "max_sn_row": max_sn_row, "outputs": output_paths}


if __name__ == "__main__":
    src_tq = r"C:\Users\Admin\.gemini\antigravity\brain\74addb61-c896-4052-bc3d-ededa129ccfc\.user_uploaded\media_1795266445732_8c00556c.xlsx"
    src_sn = r"C:\Users\Admin\.gemini\antigravity\brain\74addb61-c896-4052-bc3d-ededa129ccfc\.user_uploaded\media_1795268145873_5bff95b2.xlsx"
    outs = [
        r"D:\2026\MANY\TCS_MANY\TCS_MANY\TQ_QHNS_TongHop.xlsx",
        r"C:\Users\Admin\.gemini\antigravity\scratch\tax_xml_to_excel\TQ_QHNS_TongHop.xlsx",
    ]
    res = populate_tq_and_qhns(src_tq, src_sn, outs)
    print("Completed QHNS export:", res)
