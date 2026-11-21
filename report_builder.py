# -*- coding: utf-8 -*-
"""
Module tạo sheet 'BaoCao_BienBan':
BÁO CÁO TỔNG HỢP SỐ LIỆU ĐƯA VÀO BIÊN BẢN THANH TRA - KIỂM TRA THUẾ.

Tất cả các bảng trên sheet 'BaoCao_BienBan' đều được cài sẵn 100% công thức Excel liên kết trực tiếp
từ các sheet năm ('2020'..'2025'...) và sheet tính tiền chậm nộp & xử phạt ('Tinh_TienChamNop'):
  1. Thuế GTGT đầu ra ('YYYY'!C14, D14)
  2. Thuế GTGT còn được khấu trừ kỳ trước chuyển sang & đầu vào được khấu trừ trong kỳ ('YYYY'!C3+C18, D3+D18)
  3. Điều chỉnh tăng, giảm thuế GTGT còn được khấu trừ các kỳ trước ('YYYY'!C20/D20, C19/D19)
  4. Tổng hợp tăng, giảm thuế GTGT còn được khấu trừ chuyển kỳ sau, thuế GTGT phải nộp ('YYYY'!C23/C24, D23/D24)
  5. Thu nhập chịu thuế ('YYYY'!C60, D60) kèm ô ghi chú Số lỗ kỳ trước chuyển sang
  6. Thuế TNDN được hưởng ưu đãi miễn giảm ('YYYY'!C73+C74, D73+D74)
  7. Thuế TNDN phải nộp ('YYYY'!C76, D76)
  8. Thu nhập từ tiền lương, tiền công & Thuế TNCN đã khấu trừ ('YYYY'!C84+C86+C88, C90)
  9. Thuế nhà thầu nước ngoài ('YYYY'!C95, C97, C98)
  10. Tổng hợp số thuế truy thu (chi tiết từng năm), tiền chậm nộp, phạt khai sai 20%, phạt trốn thuế,
      phạt VPHC thủ tục, giảm thuế GTGT còn được khấu trừ chuyển kỳ sau & giảm lỗ chuyển kỳ sau.
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SHEET_NAME_REPORT = "BaoCao_BienBan"


def _detect_years_in_workbook(wb: openpyxl.Workbook) -> List[int]:
    years = [int(s.strip()) for s in wb.sheetnames if re.fullmatch(r"20\d{2}", s.strip())]
    years.sort()
    return years


def _locate_penalty_sheet_coords(
    wb: openpyxl.Workbook,
    penalty_sheet_name: str = "Tinh_TienChamNop",
) -> Dict[str, object]:
    """
    Quét sheet 'Tinh_TienChamNop' để dò chính xác tọa độ dòng của:
    - Bảng tổng hợp theo sắc thuế ở cuối sheet (summary_rows)
    - Dòng tổng từng sắc thuế theo từng năm (yearly_rows)
    """
    info: Dict[str, object] = {
        "exists": False,
        "summary": {},       # key -> row_idx (C=Truy thu, D=Phạt 20%, E=Trốn thuế, F=VPHC, G=Chậm nộp, H=Tổng)
        "yearly_vat": {},    # year -> row_idx (Col C = Truy thu GTGT)
        "yearly_nt_vat": {}, # year -> row_idx
        "yearly_nt_cit": {}, # year -> row_idx
    }
    if penalty_sheet_name not in wb.sheetnames:
        return info

    ws = wb[penalty_sheet_name]
    info["exists"] = True
    summary: Dict[str, int] = {}
    yearly_vat: Dict[int, int] = {}
    yearly_nt_vat: Dict[int, int] = {}
    yearly_nt_cit: Dict[int, int] = {}

    in_summary = False
    cur_year: Optional[int] = None

    for r in range(1, ws.max_row + 1):
        val_a = str(ws.cell(r, 1).value or "").strip()
        val_b = str(ws.cell(r, 2).value or "").strip()

        if "BẢNG TỔNG HỢP SỐ THUẾ TRUY THU" in val_a.upper():
            in_summary = True
            cur_year = None
            continue

        if not in_summary:
            m_yr = re.fullmatch(r"Năm\s+(20\d{2})", val_b)
            if m_yr:
                cur_year = int(m_yr.group(1))
                continue
            if cur_year is not None:
                if val_b == "Thuế GTGT" and val_a == "1":
                    yearly_vat[cur_year] = r
                elif "+ Thuế GTGT Nhà thầu" in val_b:
                    yearly_nt_vat[cur_year] = r
                elif "+ Thuế TNDN Nhà thầu" in val_b:
                    yearly_nt_cit[cur_year] = r
        else:
            if val_b == "Thuế GTGT":
                summary["vat"] = r
            elif val_b == "Thuế TNDN":
                summary["cit"] = r
            elif val_b == "Thuế TNCN":
                summary["pit"] = r
            elif "+ Thuế GTGT Nhà thầu" in val_b:
                summary["nt_vat"] = r
            elif "+ Thuế TNDN Nhà thầu" in val_b:
                summary["nt_cit"] = r
            elif val_b == "Lệ phí môn bài":
                summary["lpmb"] = r
            elif "Phạt VPHC về thủ tục thuế" in val_b:
                summary["vphc"] = r
            elif val_b == "TỔNG CỘNG":
                summary["total"] = r
                break

    info["summary"] = summary
    info["yearly_vat"] = yearly_vat
    info["yearly_nt_vat"] = yearly_nt_vat
    info["yearly_nt_cit"] = yearly_nt_cit
    return info


def add_audit_report_sheet(
    wb: openpyxl.Workbook,
    years: Optional[Sequence[int]] = None,
    sheet_name: str = SHEET_NAME_REPORT,
    penalty_sheet_name: str = "Tinh_TienChamNop",
) -> openpyxl.worksheet.worksheet.Worksheet:
    """
    Tạo mới (hoặc cập nhật) sheet 'BaoCao_BienBan' với toàn bộ 10 bảng tổng hợp cài sẵn công thức
    liên kết trực tiếp từ các sheet năm ('2020'..'2025'...) và sheet 'Tinh_TienChamNop'.
    """
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    target_years = sorted(set(years)) if years else _detect_years_in_workbook(wb)
    if not target_years:
        target_years = [2024]

    pen_info = _locate_penalty_sheet_coords(wb, penalty_sheet_name)
    pen_exists = bool(pen_info["exists"])
    pen_sum: Dict[str, int] = pen_info["summary"]  # type: ignore[assignment]
    pen_y_vat: Dict[int, int] = pen_info["yearly_vat"]  # type: ignore[assignment]
    pen_y_nt_vat: Dict[int, int] = pen_info["yearly_nt_vat"]  # type: ignore[assignment]
    pen_y_nt_cit: Dict[int, int] = pen_info["yearly_nt_cit"]  # type: ignore[assignment]

    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    # ===== STYLES =====
    FONT_NAME = "Times New Roman"
    f_title = Font(name=FONT_NAME, size=13, bold=True, color="1F4E78")
    f_subtitle = Font(name=FONT_NAME, size=10, italic=True, color="444444")
    f_sec = Font(name=FONT_NAME, size=11, bold=True, color="1F4E78")
    f_hdr = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    f_hdr_code = Font(name=FONT_NAME, size=9, italic=True, color="1F4E78")
    f_reg = Font(name=FONT_NAME, size=10)
    f_bold = Font(name=FONT_NAME, size=10, bold=True)
    f_italic = Font(name=FONT_NAME, size=10, italic=True)
    f_total = Font(name=FONT_NAME, size=10, bold=True, color="000000")
    f_grand = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_code = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_sub = PatternFill(start_color="EEF4FA", end_color="EEF4FA", fill_type="solid")
    fill_total = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_grand = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin = Side(style="thin", color="808080")
    med = Side(style="medium", color="1F4E78")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    b_hdr = Border(left=thin, right=thin, top=med, bottom=med)

    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    NUM = "#,##0"

    def yr_ref(yr: int, cell_ref: str) -> str:
        s = str(yr)
        if s in wb.sheetnames:
            return f"'{s}'!{cell_ref}"
        return "0"

    def style_range(r1: int, c1: int, r2: int, c2: int, font=None, fill=None, border=b_thin, align=None, num_fmt=None):
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                c = ws.cell(rr, cc)
                if font:
                    c.font = font
                if fill:
                    c.fill = fill
                if border:
                    c.border = border
                if align:
                    c.alignment = align
                if num_fmt:
                    c.number_format = num_fmt

    def write_section_title(row_idx: int, title: str, end_col: int = 8) -> int:
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=end_col)
        c = ws.cell(row_idx, 1, title)
        c.font = f_sec
        c.alignment = al_l
        ws.row_dimensions[row_idx].height = 22
        return row_idx + 1

    def write_std_5col_header(row_idx: int, col2_title: str = "Năm") -> int:
        ws.row_dimensions[row_idx].height = 24
        ws.row_dimensions[row_idx + 1].height = 18
        hdrs = ["STT", col2_title, "Số báo cáo của người nộp thuế", "Số kiểm tra", "Chênh lệch"]
        codes = ["(1)", "(2)", "(3)", "(4)", "(5)=(4)-(3)"]
        for idx, (h, cd) in enumerate(zip(hdrs, codes), start=1):
            c1 = ws.cell(row_idx, idx, h)
            c1.font = f_hdr
            c1.fill = fill_hdr
            c1.alignment = al_c
            c1.border = b_hdr

            c2 = ws.cell(row_idx + 1, idx, cd)
            c2.font = f_hdr_code
            c2.fill = fill_code
            c2.alignment = al_c
            c2.border = b_thin
        return row_idx + 2

    # ===== TIÊU ĐỀ CHÍNH =====
    ws.merge_cells("A1:H1")
    ws["A1"] = "BÁO CÁO TỔNG HỢP SỐ LIỆU ĐƯA VÀO BIÊN BẢN KIỂM TRA THUẾ"
    ws["A1"].font = f_title
    ws["A1"].alignment = al_c
    ws.row_dimensions[1].height = 26

    yr_span = f"{target_years[0]} – {target_years[-1]}" if len(target_years) > 1 else f"{target_years[0]}"
    ws.merge_cells("A2:H2")
    ws["A2"] = f"(Kỳ kiểm tra: Năm {yr_span} | Đơn vị tính: Đồng | Tự động liên kết công thức từ các sheet Năm và Tinh_TienChamNop)"
    ws["A2"].font = f_subtitle
    ws["A2"].alignment = al_c
    ws.row_dimensions[2].height = 18

    cur = 4

    # =========================================================================
    # BẢNG 1: THUẾ GTGT ĐẦU RA
    # =========================================================================
    cur = write_section_title(cur, "1. Thuế GTGT đầu ra:", 5)
    cur = write_std_5col_header(cur, "Năm")
    b1_start = cur
    for idx, yr in enumerate(target_years, start=1):
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        ws.cell(cur, 3, f"={yr_ref(yr, 'C14')}")
        ws.cell(cur, 4, f"={yr_ref(yr, 'D14')}")
        ws.cell(cur, 5, f"=D{cur}-C{cur}")
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1
    b1_end = cur - 1
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    for cc in (3, 4, 5):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, f"=SUM({col_l}{b1_start}:{col_l}{b1_end})")
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 5, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 2: THUẾ GTGT CÒN ĐƯỢC KHẤU TRỪ KỲ TRƯỚC & ĐẦU VÀO ĐƯỢC KHẤU TRỪ TRONG KỲ
    # =========================================================================
    cur = write_section_title(cur, "2. Thuế GTGT còn được khấu trừ kỳ trước chuyển sang và thuế GTGT đầu vào được khấu trừ trong kỳ:", 5)
    cur = write_std_5col_header(cur, "Năm")
    b2_start = cur
    for idx, yr in enumerate(target_years, start=1):
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        ws.cell(cur, 3, f"={yr_ref(yr, 'C3')}+{yr_ref(yr, 'C18')}")
        ws.cell(cur, 4, f"={yr_ref(yr, 'D3')}+{yr_ref(yr, 'D18')}")
        ws.cell(cur, 5, f"=D{cur}-C{cur}")
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1
    b2_end = cur - 1
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    for cc in (3, 4, 5):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, f"=SUM({col_l}{b2_start}:{col_l}{b2_end})")
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 5, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 3: ĐIỀU CHỈNH TĂNG, GIẢM THUẾ GTGT CÒN ĐƯỢC KHẤU TRỪ CÁC KỲ TRƯỚC
    # =========================================================================
    cur = write_section_title(cur, "3. Điều chỉnh tăng, giảm thuế GTGT còn được khấu trừ các kỳ trước (nếu có phát sinh chênh lệch qua kiểm tra):", 6)
    r_h1, r_h2, r_h3 = cur, cur + 1, cur + 2
    ws.row_dimensions[r_h1].height = 20
    ws.row_dimensions[r_h2].height = 20
    ws.row_dimensions[r_h3].height = 18
    for col_i, txt in ((1, "STT"), (2, "Năm / Chỉ tiêu"), (3, "Số báo cáo của người nộp thuế"), (4, "Số kiểm tra")):
        ws.merge_cells(start_row=r_h1, start_column=col_i, end_row=r_h2, end_column=col_i)
        ws.cell(r_h1, col_i, txt)
    ws.merge_cells(start_row=r_h1, start_column=5, end_row=r_h1, end_column=6)
    ws.cell(r_h1, 5, "Chênh lệch")
    ws.cell(r_h2, 5, "Điều chỉnh tăng")
    ws.cell(r_h2, 6, "Điều chỉnh giảm")
    style_range(r_h1, 1, r_h2, 6, font=f_hdr, fill=fill_hdr, border=b_hdr, align=al_c)
    for idx_c, cd in enumerate(["(1)", "(2)", "(3)", "(4)", "(5)=(4)-(3)", "(6)=(3)-(4)"], start=1):
        ws.cell(r_h3, idx_c, cd)
    style_range(r_h3, 1, r_h3, 6, font=f_hdr_code, fill=fill_code, align=al_c)
    cur += 3

    b3_data_rows: List[int] = []
    for idx, yr in enumerate(target_years, start=1):
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        style_range(cur, 1, cur, 6, font=f_bold, fill=fill_sub)
        cur += 1

        # a. Điều chỉnh tăng (row 20 on year sheet)
        r_a = cur
        b3_data_rows.append(r_a)
        ws.row_dimensions[r_a].height = 19
        ws.cell(r_a, 1, "a").alignment = al_c
        ws.cell(r_a, 2, "Điều chỉnh tăng").alignment = al_l
        ws.cell(r_a, 3, f"={yr_ref(yr, 'C20')}")
        ws.cell(r_a, 4, f"={yr_ref(yr, 'D20')}")
        ws.cell(r_a, 5, f"=IF(D{r_a}-C{r_a}>0,D{r_a}-C{r_a},0)")
        ws.cell(r_a, 6, f"=IF(D{r_a}-C{r_a}<0,C{r_a}-D{r_a},0)")
        style_range(r_a, 1, r_a, 2, font=f_reg)
        style_range(r_a, 3, r_a, 6, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1

        # b. Điều chỉnh giảm (row 19 on year sheet)
        r_b = cur
        b3_data_rows.append(r_b)
        ws.row_dimensions[r_b].height = 19
        ws.cell(r_b, 1, "b").alignment = al_c
        ws.cell(r_b, 2, "Điều chỉnh giảm").alignment = al_l
        ws.cell(r_b, 3, f"={yr_ref(yr, 'C19')}")
        ws.cell(r_b, 4, f"={yr_ref(yr, 'D19')}")
        ws.cell(r_b, 5, f"=IF(D{r_b}-C{r_b}>0,D{r_b}-C{r_b},0)")
        ws.cell(r_b, 6, f"=IF(D{r_b}-C{r_b}<0,C{r_b}-D{r_b},0)")
        style_range(r_b, 1, r_b, 2, font=f_reg)
        style_range(r_b, 3, r_b, 6, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1

    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    for cc in (3, 4, 5, 6):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, "=" + "+".join(f"{col_l}{rr}" for rr in b3_data_rows))
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 6, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 4: TỔNG HỢP TĂNG, GIẢM THUẾ GTGT CÒN ĐƯỢC KHẤU TRỪ & PHẢI NỘP
    # =========================================================================
    cur = write_section_title(cur, "4. Tổng hợp tăng, giảm thuế GTGT còn được khấu trừ chuyển kỳ sau, thuế GTGT phải nộp:", 8)
    r_h1, r_h2, r_h3 = cur, cur + 1, cur + 2
    ws.row_dimensions[r_h1].height = 22
    ws.row_dimensions[r_h2].height = 26
    ws.row_dimensions[r_h3].height = 18
    ws.merge_cells(start_row=r_h1, start_column=1, end_row=r_h2, end_column=1)
    ws.cell(r_h1, 1, "STT")
    ws.merge_cells(start_row=r_h1, start_column=2, end_row=r_h2, end_column=2)
    ws.cell(r_h1, 2, "Năm")
    ws.merge_cells(start_row=r_h1, start_column=3, end_row=r_h1, end_column=4)
    ws.cell(r_h1, 3, "Số báo cáo của người nộp thuế")
    ws.merge_cells(start_row=r_h1, start_column=5, end_row=r_h1, end_column=6)
    ws.cell(r_h1, 5, "Số kiểm tra")
    ws.merge_cells(start_row=r_h1, start_column=7, end_row=r_h1, end_column=8)
    ws.cell(r_h1, 7, "Chênh lệch")
    sub_b4 = [
        (3, "Thuế GTGT\nphải nộp"), (4, "Thuế GTGT còn\nđược khấu trừ"),
        (5, "Thuế GTGT\nphải nộp"), (6, "Thuế GTGT còn\nđược khấu trừ"),
        (7, "Thuế GTGT\nphải nộp"), (8, "Thuế GTGT còn\nđược khấu trừ"),
    ]
    for cc, txt in sub_b4:
        ws.cell(r_h2, cc, txt)
    style_range(r_h1, 1, r_h2, 8, font=f_hdr, fill=fill_hdr, border=b_hdr, align=al_c)
    for idx_c, cd in enumerate(["(1)", "(2)", "(3a)", "(3b)", "(4a)", "(4b)", "(5a)=(4a)-(3a)", "(5b)=(4b)-(3b)"], start=1):
        ws.cell(r_h3, idx_c, cd)
    style_range(r_h3, 1, r_h3, 8, font=f_hdr_code, fill=fill_code, align=al_c)
    cur += 3

    b4_start = cur
    b4_rows_by_year: Dict[int, int] = {}
    for idx, yr in enumerate(target_years, start=1):
        b4_rows_by_year[yr] = cur
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        ws.cell(cur, 3, f"={yr_ref(yr, 'C23')}")
        ws.cell(cur, 4, f"={yr_ref(yr, 'C24')}")
        ws.cell(cur, 5, f"={yr_ref(yr, 'D23')}")
        ws.cell(cur, 6, f"={yr_ref(yr, 'D24')}")
        ws.cell(cur, 7, f"=E{cur}-C{cur}")
        ws.cell(cur, 8, f"=F{cur}-D{cur}")
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 8, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1
    b4_end = cur - 1
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    ws.cell(cur, 3, f"=SUM(C{b4_start}:C{b4_end})")
    ws.cell(cur, 5, f"=SUM(E{b4_start}:E{b4_end})")
    ws.cell(cur, 7, f"=SUM(G{b4_start}:G{b4_end})")
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 8, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 5: THU NHẬP CHỊU THUẾ (KÈM GHI CHÚ SỐ CHUYỂN LỖ KỲ TRƯỚC)
    # =========================================================================
    r_b5_title = cur
    cur = write_section_title(cur, "5. Thu nhập chịu thuế:", 5)
    # Box Số chuyển lỗ kỳ trước bên phải (Cột G-H)
    ws.cell(r_b5_title, 7, "Theo dõi số lỗ kỳ trước chuyển sang (nếu có)").font = f_sec
    ws.merge_cells(start_row=r_b5_title, start_column=7, end_row=r_b5_title, end_column=8)

    cur = write_std_5col_header(cur, "Năm")
    ws.cell(r_b5_title + 1, 7, "Nội dung chuyển lỗ kỳ trước").font = f_hdr
    ws.cell(r_b5_title + 1, 8, "Số tiền (đồng)").font = f_hdr
    style_range(r_b5_title + 1, 7, r_b5_title + 1, 8, font=f_hdr, fill=fill_hdr, border=b_hdr, align=al_c)
    ws.cell(r_b5_title + 2, 7, f"Số lỗ trước năm {target_years[0]} chuyển sang (theo BB/TK trước)").font = f_reg
    ws.cell(r_b5_title + 2, 7).alignment = al_l
    ws.cell(r_b5_title + 2, 7).border = b_thin
    ws.cell(r_b5_title + 2, 8, 0).font = f_reg
    ws.cell(r_b5_title + 2, 8).number_format = NUM
    ws.cell(r_b5_title + 2, 8).alignment = al_r
    ws.cell(r_b5_title + 2, 8).fill = fill_input
    ws.cell(r_b5_title + 2, 8).border = b_thin

    b5_start = cur
    b5_rows_by_year: Dict[int, int] = {}
    for idx, yr in enumerate(target_years, start=1):
        b5_rows_by_year[yr] = cur
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        ws.cell(cur, 3, f"={yr_ref(yr, 'C60')}")
        ws.cell(cur, 4, f"={yr_ref(yr, 'D60')}")
        ws.cell(cur, 5, f"=D{cur}-C{cur}")
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1
    b5_end = cur - 1
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    for cc in (3, 4, 5):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, f"=SUM({col_l}{b5_start}:{col_l}{b5_end})")
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 5, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 6: THUẾ TNDN ĐƯỢC HƯỞNG ƯU ĐÃI MIỄN GIẢM
    # =========================================================================
    cur = write_section_title(cur, "6. Thuế TNDN được hưởng ưu đãi miễn giảm (nếu có phát sinh trong kỳ kiểm tra):", 5)
    cur = write_std_5col_header(cur, "Năm")
    b6_start = cur
    for idx, yr in enumerate(target_years, start=1):
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        ws.cell(cur, 3, f"={yr_ref(yr, 'C73')}+{yr_ref(yr, 'C74')}")
        ws.cell(cur, 4, f"={yr_ref(yr, 'D73')}+{yr_ref(yr, 'D74')}")
        ws.cell(cur, 5, f"=D{cur}-C{cur}")
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1
    b6_end = cur - 1
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    for cc in (3, 4, 5):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, f"=SUM({col_l}{b6_start}:{col_l}{b6_end})")
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 5, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 7: THUẾ TNDN PHẢI NỘP
    # =========================================================================
    cur = write_section_title(cur, "7. Thuế TNDN phải nộp:", 5)
    cur = write_std_5col_header(cur, "Năm")
    b7_start = cur
    b7_rows_by_year: Dict[int, int] = {}
    for idx, yr in enumerate(target_years, start=1):
        b7_rows_by_year[yr] = cur
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        ws.cell(cur, 3, f"={yr_ref(yr, 'C76')}")
        ws.cell(cur, 4, f"={yr_ref(yr, 'D76')}")
        ws.cell(cur, 5, f"=D{cur}-C{cur}")
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1
    b7_end = cur - 1
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng").alignment = al_c
    for cc in (3, 4, 5):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, f"=SUM({col_l}{b7_start}:{col_l}{b7_end})")
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 5, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 8: THU NHẬP TỪ TIỀN LƯƠNG, TIỀN CÔNG (THUẾ TNCN)
    # =========================================================================
    cur = write_section_title(cur, "8. Thu nhập từ tiền lương, tiền công (Thuế TNCN):", 5)
    cur = write_std_5col_header(cur, "Năm / Chỉ tiêu")
    b8_tax_rows: Dict[int, int] = {}
    for idx, yr in enumerate(target_years, start=1):
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, idx).alignment = al_c
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        style_range(cur, 1, cur, 5, font=f_bold, fill=fill_sub)
        cur += 1

        # a. Thu nhập tính thuế / chịu thuế (C84+C86+C88)
        r_a = cur
        ws.row_dimensions[r_a].height = 19
        ws.cell(r_a, 1, "a").alignment = al_c
        ws.cell(r_a, 2, "Thu nhập tính thuế").alignment = al_l
        ws.cell(r_a, 3, f"={yr_ref(yr, 'C84')}+{yr_ref(yr, 'C86')}+{yr_ref(yr, 'C88')}")
        ws.cell(r_a, 4, f"={yr_ref(yr, 'D84')}+{yr_ref(yr, 'D86')}+{yr_ref(yr, 'D88')}")
        ws.cell(r_a, 5, f"=D{r_a}-C{r_a}")
        style_range(r_a, 1, r_a, 2, font=f_reg)
        style_range(r_a, 3, r_a, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1

        # b. Thuế đã khấu trừ (C90)
        r_b = cur
        b8_tax_rows[yr] = r_b
        ws.row_dimensions[r_b].height = 19
        ws.cell(r_b, 1, "b").alignment = al_c
        ws.cell(r_b, 2, "Thuế đã khấu trừ").alignment = al_l
        ws.cell(r_b, 3, f"={yr_ref(yr, 'C90')}")
        ws.cell(r_b, 4, f"={yr_ref(yr, 'D90')}")
        ws.cell(r_b, 5, f"=D{r_b}-C{r_b}")
        style_range(r_b, 1, r_b, 2, font=f_reg)
        style_range(r_b, 3, r_b, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1

    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 2, "Tổng cộng (Thuế TNCN đã khấu trừ)").alignment = al_c
    for cc in (3, 4, 5):
        col_l = get_column_letter(cc)
        ws.cell(cur, cc, "=" + "+".join(f"{col_l}{b8_tax_rows[y]}" for y in target_years))
    style_range(cur, 1, cur, 2, font=f_total, fill=fill_total)
    style_range(cur, 3, cur, 5, font=f_total, fill=fill_total, align=al_r, num_fmt=NUM)
    cur += 2

    # =========================================================================
    # BẢNG 9: THUẾ NHÀ THẦU NƯỚC NGOÀI
    # =========================================================================
    cur = write_section_title(cur, "9. Thuế nhà thầu nước ngoài:", 5)
    cur = write_std_5col_header(cur, "Chỉ tiêu")
    b9_nt_vat_rows: Dict[int, int] = {}
    b9_nt_cit_rows: Dict[int, int] = {}
    for yr in target_years:
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 2, f"Năm {yr}").alignment = al_l
        style_range(cur, 1, cur, 5, font=f_bold, fill=fill_sub)
        cur += 1

        # 1. Doanh thu tính thuế NTNN
        r_dt = cur
        ws.row_dimensions[r_dt].height = 19
        ws.cell(r_dt, 1, 1).alignment = al_c
        ws.cell(r_dt, 2, "Doanh thu tính thuế của nhà thầu nước ngoài").alignment = al_l
        ws.cell(r_dt, 3, f"={yr_ref(yr, 'C95')}")
        ws.cell(r_dt, 4, f"={yr_ref(yr, 'D95')}")
        ws.cell(r_dt, 5, f"=D{r_dt}-C{r_dt}")
        style_range(r_dt, 1, r_dt, 2, font=f_reg)
        style_range(r_dt, 3, r_dt, 5, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1

        # 2. Thuế nhà thầu phải khấu trừ
        r_sum_nt = cur
        r_nt_g = cur + 1
        r_nt_t = cur + 2
        b9_nt_vat_rows[yr] = r_nt_g
        b9_nt_cit_rows[yr] = r_nt_t

        ws.row_dimensions[r_sum_nt].height = 19
        ws.cell(r_sum_nt, 1, 2).alignment = al_c
        ws.cell(r_sum_nt, 2, "Thuế nhà thầu phải khấu trừ").alignment = al_l
        ws.cell(r_sum_nt, 3, f"=C{r_nt_g}+C{r_nt_t}")
        ws.cell(r_sum_nt, 4, f"=D{r_nt_g}+D{r_nt_t}")
        ws.cell(r_sum_nt, 5, f"=D{r_sum_nt}-C{r_sum_nt}")
        style_range(r_sum_nt, 1, r_sum_nt, 2, font=f_bold)
        style_range(r_sum_nt, 3, r_sum_nt, 5, font=f_bold, align=al_r, num_fmt=NUM)

        ws.row_dimensions[r_nt_g].height = 19
        ws.cell(r_nt_g, 1, "a").alignment = al_c
        ws.cell(r_nt_g, 2, "Thuế GTGT").alignment = al_l
        ws.cell(r_nt_g, 3, f"={yr_ref(yr, 'C97')}")
        ws.cell(r_nt_g, 4, f"={yr_ref(yr, 'D97')}")
        ws.cell(r_nt_g, 5, f"=D{r_nt_g}-C{r_nt_g}")
        style_range(r_nt_g, 1, r_nt_g, 2, font=f_italic)
        style_range(r_nt_g, 3, r_nt_g, 5, font=f_reg, align=al_r, num_fmt=NUM)

        ws.row_dimensions[r_nt_t].height = 19
        ws.cell(r_nt_t, 1, "b").alignment = al_c
        ws.cell(r_nt_t, 2, "Thuế TNDN").alignment = al_l
        ws.cell(r_nt_t, 3, f"={yr_ref(yr, 'C98')}")
        ws.cell(r_nt_t, 4, f"={yr_ref(yr, 'D98')}")
        ws.cell(r_nt_t, 5, f"=D{r_nt_t}-C{r_nt_t}")
        style_range(r_nt_t, 1, r_nt_t, 2, font=f_italic)
        style_range(r_nt_t, 3, r_nt_t, 5, font=f_reg, align=al_r, num_fmt=NUM)

        cur += 3

    cur += 1

    # =========================================================================
    # BẢNG 10: TỔNG HỢP SỐ THUẾ TRUY THU, PHẠT, TIỀN CHẬM NỘP & GIẢM KHẤU TRỪ / GIẢM LỖ
    # =========================================================================
    cur = write_section_title(cur, "10. TỔNG HỢP SỐ THUẾ TRUY THU, TIỀN PHẠT, TIỀN CHẬM NỘP VÀ ĐIỀU CHỈNH KHÁC (KẾT LUẬN BIÊN BẢN):", 5)
    ws.row_dimensions[cur].height = 24
    for cc, txt in ((1, "STT"), (2, "Loại thuế / Khoản xử lý"), (3, "Số tiền (đồng)")):
        c = ws.cell(cur, cc, txt)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_hdr
    cur += 1

    def pen_sum_ref(key: str, col_letter: str) -> str:
        if pen_exists and key in pen_sum:
            return f"='{penalty_sheet_name}'!{col_letter}{pen_sum[key]}"
        return "=0"

    grand_add_rows: List[int] = []

    def write_truy_thu_group(stt: int, title: str, year_formula_fn) -> int:
        nonlocal cur
        r_parent = cur
        grand_add_rows.append(r_parent)
        ws.row_dimensions[r_parent].height = 20
        ws.cell(r_parent, 1, stt).alignment = al_c
        ws.cell(r_parent, 2, title).alignment = al_l
        cur += 1
        child_start = cur
        for yr in target_years:
            ws.row_dimensions[cur].height = 19
            ws.cell(cur, 2, f"+ Năm {yr}").alignment = al_l
            ws.cell(cur, 3, year_formula_fn(yr))
            style_range(cur, 1, cur, 2, font=f_reg)
            style_range(cur, 3, cur, 3, font=f_reg, align=al_r, num_fmt=NUM)
            cur += 1
        child_end = cur - 1
        ws.cell(r_parent, 3, f"=SUM(C{child_start}:C{child_end})")
        style_range(r_parent, 1, r_parent, 2, font=f_bold, fill=fill_sub)
        style_range(r_parent, 3, r_parent, 3, font=f_bold, fill=fill_sub, align=al_r, num_fmt=NUM)
        return r_parent

    # 1. Truy thu thuế GTGT
    write_truy_thu_group(
        1,
        "Truy thu thuế GTGT:",
        lambda y: (
            f"=IF('{penalty_sheet_name}'!C{pen_y_vat[y]}>0,'{penalty_sheet_name}'!C{pen_y_vat[y]},MAX(0,G{b4_rows_by_year[y]}))"
            if (pen_exists and y in pen_y_vat)
            else f"=MAX(0,G{b4_rows_by_year[y]})"
        ),
    )
    # 2. Truy thu thuế TNDN
    write_truy_thu_group(2, "Truy thu thuế TNDN:", lambda y: f"=MAX(0,E{b7_rows_by_year[y]})")
    # 3. Truy thu thuế TNCN
    write_truy_thu_group(3, "Truy thu thuế TNCN:", lambda y: f"=MAX(0,E{b8_tax_rows[y]})")
    # 4. Truy thu thuế GTGT nhà thầu NN
    write_truy_thu_group(
        4,
        "Truy thu thuế GTGT nhà thầu nước ngoài:",
        lambda y: (
            f"=IF('{penalty_sheet_name}'!C{pen_y_nt_vat[y]}>0,'{penalty_sheet_name}'!C{pen_y_nt_vat[y]},MAX(0,E{b9_nt_vat_rows[y]}))"
            if (pen_exists and y in pen_y_nt_vat)
            else f"=MAX(0,E{b9_nt_vat_rows[y]})"
        ),
    )
    # 5. Truy thu thuế TNDN nhà thầu NN
    write_truy_thu_group(
        5,
        "Truy thu thuế TNDN nhà thầu nước ngoài:",
        lambda y: (
            f"=IF('{penalty_sheet_name}'!C{pen_y_nt_cit[y]}>0,'{penalty_sheet_name}'!C{pen_y_nt_cit[y]},MAX(0,E{b9_nt_cit_rows[y]}))"
            if (pen_exists and y in pen_y_nt_cit)
            else f"=MAX(0,E{b9_nt_cit_rows[y]})"
        ),
    )

    # 6..17: Tiền chậm nộp, Phạt khai sai 20%, Phạt trốn thuế, Phạt VPHC thủ tục (link từ Tinh_TienChamNop)
    flat_items: List[Tuple[int, str, str]] = [
        (6, "Tiền chậm nộp do truy thu thuế GTGT", pen_sum_ref("vat", "G")),
        (7, "Tiền chậm nộp do truy thu thuế TNDN", pen_sum_ref("cit", "G")),
        (8, "Tiền chậm nộp do truy thu thuế TNCN", pen_sum_ref("pit", "G")),
        (9, "Tiền chậm nộp do truy thu thuế GTGT nhà thầu NN", pen_sum_ref("nt_vat", "G")),
        (10, "Tiền chậm nộp do truy thu thuế TNDN nhà thầu NN", pen_sum_ref("nt_cit", "G")),
        (11, "Phạt về hành vi khai sai thuế GTGT (20%)", pen_sum_ref("vat", "D")),
        (12, "Phạt về hành vi khai sai thuế TNDN (20%)", pen_sum_ref("cit", "D")),
        (13, "Phạt về hành vi khai sai thuế TNCN (20%)", pen_sum_ref("pit", "D")),
        (14, "Phạt về hành vi khai sai thuế GTGT nhà thầu NN (20%)", pen_sum_ref("nt_vat", "D")),
        (15, "Phạt về hành vi khai sai thuế TNDN nhà thầu NN (20%)", pen_sum_ref("nt_cit", "D")),
        (16, "Phạt về hành vi trốn thuế (nếu có)", pen_sum_ref("total", "E")),
        (17, "Phạt vi phạm hành chính về thủ tục thuế", pen_sum_ref("vphc", "F")),
    ]
    for stt, lbl, formula in flat_items:
        grand_add_rows.append(cur)
        ws.row_dimensions[cur].height = 19
        ws.cell(cur, 1, stt).alignment = al_c
        ws.cell(cur, 2, lbl).alignment = al_l
        ws.cell(cur, 3, formula)
        style_range(cur, 1, cur, 2, font=f_reg)
        style_range(cur, 3, cur, 3, font=f_reg, align=al_r, num_fmt=NUM)
        cur += 1

    # Tổng cộng truy thu, phạt, chậm nộp
    r_grand = cur
    ws.row_dimensions[r_grand].height = 24
    ws.cell(r_grand, 2, "Tổng cộng truy thu, phạt và tiền chậm nộp").alignment = al_c
    ws.cell(r_grand, 3, "=" + "+".join(f"C{rr}" for rr in grand_add_rows))
    style_range(r_grand, 1, r_grand, 2, font=f_grand, fill=fill_grand, border=b_hdr)
    style_range(r_grand, 3, r_grand, 3, font=f_grand, fill=fill_grand, border=b_hdr, align=al_r, num_fmt=NUM)
    cur += 1

    # 18. Giảm số thuế GTGT còn được khấu trừ chuyển kỳ sau (năm cuối kỳ kiểm tra)
    last_yr = target_years[-1]
    r_b4_last = b4_rows_by_year[last_yr]
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 1, 18).alignment = al_c
    ws.cell(cur, 2, "Giảm số thuế GTGT còn được khấu trừ chuyển kỳ sau").alignment = al_l
    ws.cell(cur, 3, f"=MAX(0,-H{r_b4_last})")
    style_range(cur, 1, cur, 2, font=f_bold)
    style_range(cur, 3, cur, 3, font=f_bold, align=al_r, num_fmt=NUM)
    cur += 1

    # 19. Giảm lỗ chuyển kỳ sau (tổng các năm có TNCT báo cáo < 0 và chênh lệch kiểm tra > 0)
    loss_terms = [
        f"IF(C{b5_rows_by_year[y]}<0,MIN(ABS(C{b5_rows_by_year[y]}),MAX(0,E{b5_rows_by_year[y]})),0)"
        for y in target_years
    ]
    ws.row_dimensions[cur].height = 20
    ws.cell(cur, 1, 19).alignment = al_c
    ws.cell(cur, 2, "Giảm lỗ chuyển kỳ sau").alignment = al_l
    ws.cell(cur, 3, "=" + "+".join(loss_terms))
    style_range(cur, 1, cur, 2, font=f_bold)
    style_range(cur, 3, cur, 3, font=f_bold, align=al_r, num_fmt=NUM)

    # ===== ĐỘ RỘNG CỘT & CÀI ĐẶT IN =====
    col_widths = {
        1: 8,   # A: STT
        2: 48,  # B: Năm / Chỉ tiêu
        3: 22,  # C: Số báo cáo / Số tiền
        4: 22,  # D: Số kiểm tra / GTGT KT
        5: 20,  # E: Chênh lệch / KT Phải nộp
        6: 22,  # F: Điều chỉnh giảm / KT Khấu trừ
        7: 26,  # G: CL Phải nộp / Ghi chú lỗ
        8: 22,  # H: CL Khấu trừ / Số lỗ kỳ trước
    }
    for cc, w in col_widths.items():
        ws.column_dimensions[get_column_letter(cc)].width = w

    ws.freeze_panes = "C4"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    return ws
