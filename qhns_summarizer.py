# -*- coding: utf-8 -*-
"""
Module phân tích sổ Tổng quan nghĩa vụ thuế (TQ - TMS) và tổng hợp tự động bằng công thức Excel
SUMIFS vào 1 biểu duy nhất trên sheet Quan hệ Ngân sách (QHNS) theo đúng nguyên tắc nghiệp vụ:
1. Cột (3) Số kỳ trước chuyển sang (Đầu năm 2020):
   - Lấy các khoản nộp (C2) của kỳ trước năm 2020 (Nam_KyBC <= 2019) nhưng nộp trong giai đoạn 2020-2025.
   - Các năm 2021..2025: Bằng Cột (7) Số còn phải nộp cuối năm liền trước chuyển sang.
2. Cột (4) Số báo cáo phát sinh phải nộp trong kỳ (Không bao giờ có số âm, tính theo Năm Kỳ báo cáo):
   - Thuế GTGT (Tiểu mục 1701): Lấy giao dịch TK.
   - Thuế TNDN (Tiểu mục 1052): Lấy giao dịch TK, QT.
   - Thuế TNCN (Tiểu mục 1001, 1004, 1015): Lấy tổng giao dịch TK, TQ, TH.
3. Cột (5) Số xử lý qua thanh tra:
   - Lấy giao dịch QT theo Năm Ngày hạch toán.
4. Cột (6) Số đã nộp:
   - Lấy giao dịch C2 (và HC nếu có) theo Năm Ngày hạch toán (khớp 100% với file Số nộp).
"""
import datetime
import os
import re
from typing import Any, Dict, List

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

THIN_SIDE = Side(border_style="thin", color="808080")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
YEAR_FILL = PatternFill(start_color="EDEDED", end_color="EDEDED", fill_type="solid")


def extract_ht_year(dt_val: Any) -> int:
    if isinstance(dt_val, (datetime.datetime, datetime.date)):
        return dt_val.year
    s = str(dt_val or "").strip()
    m = re.search(r"\b(20\d{2}|19\d{2})\b", s)
    return int(m.group(1)) if m else 0


def extract_ky_year(ky_val: Any, ht_yr: int) -> int:
    """
    Trích xuất Năm của Kỳ báo cáo TMS (ví dụ: '1912' -> 2019, '2001' -> 2020, '2012' -> 2020, '20CN' -> 2020, '26Q2' -> 2026).
    Lưu ý: '2001'..'2012' là tháng 01..12 của năm 2020 (YY=20, MM=01..12), tuyệt đối không hiểu nhầm là năm 2001..2012.
    """
    if ky_val is None:
        return ht_yr
    s = str(ky_val).strip()
    if not s:
        return ht_yr
    m2 = re.match(r"^(\d{2})(0[1-9]|1[0-2]|Q\d|CN)$", s, re.IGNORECASE)
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
    output_paths: List[str],
) -> Dict[str, Any]:
    wb = openpyxl.load_workbook(input_tq_xlsx)
    ws_tq = wb["TQ"]

    font_hdr = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_italic = Font(name="Times New Roman", size=11, italic=True)
    font_body = Font(name="Times New Roman", size=11, bold=False)

    # 1. Bổ sung 3 cột phụ chuẩn hóa trên sheet 'TQ' tại cột Z, AA, AB (Cột 26..28):
    # - Cột Z (26): Nam_HT (Năm của Ngày hạch toán Cột J)
    # - Cột AA (27): Nam_KyBC (Năm của Kỳ báo cáo Cột B: 2001..2012 -> 2020)
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

    # 2. Xây dựng 1 biểu duy nhất chuẩn xác trên sheet 'QHNS' (A1:G33)
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

    ws_q.cell(1, 1, "STT")
    ws_q.cell(1, 2, "CHỈ TIÊU")
    ws_q.cell(1, 3, "Số kỳ trước chuyển sang")
    ws_q.cell(1, 4, "Số phát sinh trong kỳ")
    ws_q.cell(1, 6, "Số đã nộp ")
    ws_q.cell(1, 7, "Số còn phải nộp ")

    ws_q.cell(2, 4, "Số báo cáo phát sinh phải nộp trong kỳ")
    ws_q.cell(2, 5, "Số xử lý qua thanh tra")

    r3_vals = [-1, -2, -3, -4, -5, -6, "(7)=(3)+(4)+(5)-(6)"]
    for idx, v in enumerate(r3_vals, start=1):
        ws_q.cell(3, idx, v)

    ws_q.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws_q.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    ws_q.merge_cells(start_row=1, start_column=3, end_row=2, end_column=3)
    ws_q.merge_cells(start_row=1, start_column=4, end_row=1, end_column=5)
    ws_q.merge_cells(start_row=1, start_column=6, end_row=2, end_column=6)
    ws_q.merge_cells(start_row=1, start_column=7, end_row=2, end_column=7)

    for r_hdr in (1, 2, 3):
        for col_i in range(1, 8):
            cell = ws_q.cell(r_hdr, col_i)
            cell.font = font_hdr if r_hdr in (1, 2) else font_italic
            cell.fill = HEADER_FILL if r_hdr in (1, 2) else SUBHEADER_FILL
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

    ws_q.row_dimensions[1].height = 24
    ws_q.row_dimensions[2].height = 28
    ws_q.row_dimensions[3].height = 20

    years = [2020, 2021, 2022, 2023, 2024, 2025]
    prev_rows: Dict[str, int] = {}

    for y_idx, yr in enumerate(years):
        base_r = 4 + y_idx * 5
        ws_q.cell(base_r, 2, f"Năm {yr}")
        for col_i in range(1, 8):
            c = ws_q.cell(base_r, col_i)
            c.font = font_bold
            c.fill = YEAR_FILL
            c.border = THIN_BORDER
            if col_i == 2:
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

            c1 = ws_q.cell(r_cur, 1, stt_v)
            c1.font = font_body
            c1.alignment = Alignment(horizontal="center", vertical="center")
            c1.border = THIN_BORDER

            c2 = ws_q.cell(r_cur, 2, label_v)
            c2.font = font_italic if tax_code == "GTGT_VL" else font_body
            c2.alignment = Alignment(horizontal="left", vertical="center")
            c2.border = THIN_BORDER

            if tax_code == "GTGT_VL":
                f_dk = 0 if yr == 2020 else f"=G{prev_rows[tax_code]}"
                f_ps = 0
                f_tt = 0
                f_nop = 0
            else:
                if yr == 2020:
                    if tax_code == "GTGT":
                        f_dk = (
                            f"=-SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
                            f"+SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"QT\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
                        )
                    elif tax_code == "TNDN":
                        f_dk = (
                            f"=-SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
                            f"+SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"QT\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
                        )
                    else:
                        f_dk = (
                            f"=-SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"C2\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
                            f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"QT\",{rng_ky},\"<=2019\",{rng_ht},\">=2020\",{rng_ht},\"<=2025\")"
                        )
                else:
                    f_dk = f"=G{prev_rows[tax_code]}"

                if tax_code == "GTGT":
                    f_ps = f"=SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"TK\",{rng_ky},{yr})"
                    f_tt = f"=SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"QT\",{rng_ky},{yr})"
                    f_nop = f"=-SUMIFS({rng_val},{rng_tm},\"1701\",{rng_kh},\"C2\",{rng_ht},{yr})"
                elif tax_code == "TNDN":
                    f_ps = (
                        f"=SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"TK\",{rng_ky},{yr})"
                        f"+SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"QT\",{rng_ky},{yr})"
                    )
                    f_tt = 0
                    f_nop = f"=-SUMIFS({rng_val},{rng_tm},\"1052\",{rng_kh},\"C2\",{rng_ht},{yr})"
                else:
                    f_ps = (
                        f"=SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TK\",{rng_ky},{yr})"
                        f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TQ\",{rng_ky},{yr})"
                        f"+SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"TH\",{rng_ky},{yr})"
                    )
                    f_tt = f"=SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"QT\",{rng_ky},{yr})"
                    f_nop = f"=-SUMIFS({rng_val},{rng_sac},\"TNCN\",{rng_kh},\"C2\",{rng_ht},{yr})"

            f_ck = f"=C{r_cur}+D{r_cur}+E{r_cur}-F{r_cur}"

            for col_i, f_val in (
                (3, f_dk),
                (4, f_ps),
                (5, f_tt),
                (6, f_nop),
                (7, f_ck),
            ):
                cell = ws_q.cell(r_cur, col_i, f_val)
                cell.font = font_bold if col_i == 7 else font_body
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = THIN_BORDER

            prev_rows[tax_code] = r_cur

    col_widths = {
        "A": 6,
        "B": 24,
        "C": 22,
        "D": 22,
        "E": 22,
        "F": 22,
        "G": 22,
    }
    for col_let, w in col_widths.items():
        ws_q.column_dimensions[col_let].width = w

    for out_p in output_paths:
        os.makedirs(os.path.dirname(os.path.abspath(out_p)), exist_ok=True)
        wb.save(out_p)

    return {"max_tq_row": max_tq_row, "outputs": output_paths}


if __name__ == "__main__":
    src_tq = r"C:\Users\Admin\.gemini\antigravity\brain\74addb61-c896-4052-bc3d-ededa129ccfc\.user_uploaded\media_1795266445732_8c00556c.xlsx"
    outs = [
        r"D:\2026\MANY\TCS_MANY\TCS_MANY\TQ_QHNS_TongHop.xlsx",
        r"C:\Users\Admin\.gemini\antigravity\scratch\tax_xml_to_excel\TQ_QHNS_TongHop.xlsx",
    ]
    res = populate_tq_and_qhns(src_tq, outs)
    print("Completed QHNS export:", res)
