# -*- coding: utf-8 -*-
"""
Module tổng hợp số nộp ngân sách nhà nước (CTUNOP) theo từng năm và từng tiểu mục.
Xuất file Excel gồm:
1. Sheet 'TH_Nam_HachToan': Tổng hợp số nộp từng năm (theo Ngày hạch toán) x từng Tiểu mục & Nhóm sắc thuế (công thức SUMIFS)
2. Sheet 'TH_Nam_KyBaoCao': Tổng hợp số nộp từng năm (theo Kỳ báo cáo) x từng Tiểu mục (công thức SUMIFS)
3. Sheet 'Chi_Tiet_CTUNOP': Toàn bộ dữ liệu gốc kèm các cột chuẩn hóa Năm hạch toán, Năm kỳ báo cáo, Tên tiểu mục, Số tiền nộp dương.
"""
import datetime
import os
import re
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

TIEU_MUC_INFO: Dict[str, Tuple[str, str]] = {
    "1001": ("Thuế thu nhập từ tiền lương, tiền công của người lao động Việt Nam", "Thuế TNCN"),
    "1004": ("Thuế thu nhập từ đầu tư vốn của cá nhân", "Thuế TNCN"),
    "1015": ("Thuế thu nhập từ tiền lương, tiền công của người lao động nước ngoài ở Việt Nam", "Thuế TNCN"),
    "1052": ("Thuế TNDN từ hoạt động SXKD của đơn vị không hạch toán toàn ngành", "Thuế TNDN"),
    "1701": ("Thuế GTGT hàng sản xuất, kinh doanh trong nước (gồm cả dịch vụ trong lĩnh vực dầu khí)", "Thuế GTGT"),
    "1803": ("Thuế môn bài bậc 3 (áp dụng giai đoạn trước năm 2017)", "Lệ phí môn bài"),
    "2862": ("Lệ phí môn bài mức (bậc) 1", "Lệ phí môn bài"),
    "2863": ("Lệ phí môn bài mức (bậc) 2", "Lệ phí môn bài"),
    "2864": ("Lệ phí môn bài mức (bậc) 3", "Lệ phí môn bài"),
    "4254": ("Phạt vi phạm hành chính trong lĩnh vực thuế do ngành Thuế thực hiện", "Tiền phạt & Chậm nộp"),
    "4268": ("Tiền chậm nộp các khoản khác do ngành Thuế quản lý", "Tiền phạt & Chậm nộp"),
    "4272": ("Tiền chậm nộp phạt vi phạm hành chính về thuế do ngành Thuế quản lý", "Tiền phạt & Chậm nộp"),
    "4917": ("Tiền chậm nộp thuế thu nhập cá nhân", "Tiền phạt & Chậm nộp"),
    "4931": ("Tiền chậm nộp thuế giá trị gia tăng từ hàng sản xuất kinh doanh trong nước", "Tiền phạt & Chậm nộp"),
    "4944": ("Tiền chậm nộp các khoản phí, lệ phí (Lệ phí môn bài)", "Tiền phạt & Chậm nộp"),
}

TAX_GROUP_ORDER: List[str] = [
    "Thuế GTGT",
    "Thuế TNDN",
    "Thuế TNCN",
    "Lệ phí môn bài",
    "Tiền phạt & Chậm nộp",
]

THIN_SIDE = Side(border_style="thin", color="808080")
THIN_BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")


def parse_ky_bao_cao_year(ky_val: Any) -> str:
    """Trích xuất năm từ chuỗi Kỳ báo cáo của TMS (VD: '1801' -> '2018', '24CN' -> '2024', '26Q2' -> '2026')."""
    if ky_val is None:
        return "Không ghi kỳ"
    s = str(ky_val).strip()
    if not s:
        return "Không ghi kỳ"
    m4 = re.match(r"^(20\d{2})", s)
    if m4:
        return m4.group(1)
    m2 = re.match(r"^(\d{2})(\d{2}|Q\d|CN)", s, re.IGNORECASE)
    if m2:
        return f"20{m2.group(1)}"
    return "Không ghi kỳ"


def parse_year_from_date(dt_val: Any) -> int:
    if isinstance(dt_val, (datetime.datetime, datetime.date)):
        return dt_val.year
    s = str(dt_val or "").strip()
    m = re.search(r"\b(20\d{2}|19\d{2})\b", s)
    return int(m.group(1)) if m else 0


def summarize_tax_payments_to_excel(input_path: str, output_paths: List[str]) -> Dict[str, Any]:
    wb_in = openpyxl.load_workbook(input_path, data_only=True)
    ws_in = wb_in.active

    orig_headers = [ws_in.cell(1, c).value for c in range(1, ws_in.max_column + 1)]
    raw_rows: List[List[Any]] = []
    for r in range(2, ws_in.max_row + 1):
        row_vals = [ws_in.cell(r, c).value for c in range(1, len(orig_headers) + 1)]
        if any(v is not None and str(v).strip() != "" for v in row_vals):
            raw_rows.append(row_vals)

    ht_years_set = set()
    ky_years_set = set()
    tm_set = set()

    by_ht_tm: Dict[Tuple[int, str], float] = {}
    by_ky_tm: Dict[Tuple[str, str], float] = {}

    for row in raw_rows:
        ky_str = parse_ky_bao_cao_year(row[1])
        tm_str = str(row[4] or "").strip()
        so_tien_raw = float(row[6] or 0)
        so_tien_nop = -so_tien_raw
        ht_yr = parse_year_from_date(row[9])

        if ht_yr:
            ht_years_set.add(ht_yr)
        ky_years_set.add(ky_str)
        if tm_str:
            tm_set.add(tm_str)

        by_ht_tm[(ht_yr, tm_str)] = by_ht_tm.get((ht_yr, tm_str), 0.0) + so_tien_nop
        by_ky_tm[(ky_str, tm_str)] = by_ky_tm.get((ky_str, tm_str), 0.0) + so_tien_nop

    ht_years = sorted(ht_years_set)
    ky_years_num = sorted(y for y in ky_years_set if y != "Không ghi kỳ")
    ky_years = ky_years_num + (["Không ghi kỳ"] if "Không ghi kỳ" in ky_years_set else [])
    tieu_mucs = sorted(tm_set)

    wb_out = openpyxl.Workbook()
    ws_ht = wb_out.active
    ws_ht.title = "TH_Nam_HachToan"
    ws_ky = wb_out.create_sheet(title="TH_Nam_KyBaoCao")
    ws_ct = wb_out.create_sheet(title="Chi_Tiet_CTUNOP")

    font_title = Font(name="Times New Roman", size=13, bold=True, color="1F4E78")
    font_sec = Font(name="Times New Roman", size=11, bold=True, color="1F4E78")
    font_hdr = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_subhdr = Font(name="Times New Roman", size=11, bold=True, color="000000")
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_body = Font(name="Times New Roman", size=11, bold=False)

    # ==========================================================================
    # 1. GHI SHEET CHI TIẾT: 'Chi_Tiet_CTUNOP'
    # ==========================================================================
    ext_headers = orig_headers + [
        "Năm hạch toán",
        "Năm kỳ báo cáo",
        "Tên tiểu mục",
        "Số tiền đã nộp (VNĐ)",
        "Nhóm sắc thuế",
    ]
    ws_ct.append(ext_headers)
    ws_ct.row_dimensions[1].height = 28
    for c_idx in range(1, len(ext_headers) + 1):
        c = ws_ct.cell(1, c_idx)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    for idx, row in enumerate(raw_rows, start=2):
        tm_str = str(row[4] or "").strip()
        ky_yr_str = parse_ky_bao_cao_year(row[1])
        tm_name, tm_group = TIEU_MUC_INFO.get(tm_str, (f"Tiểu mục {tm_str}", "Khác"))
        ht_yr = parse_year_from_date(row[9])

        for c_off, val in enumerate(row, start=1):
            cell = ws_ct.cell(idx, c_off, val)
            cell.font = font_body
            cell.border = THIN_BORDER
            if c_off in (7, 8, 9):
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif isinstance(val, (datetime.datetime, datetime.date)):
                cell.number_format = "yyyy-mm-dd"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_off in (1, 2, 3, 4, 5, 6, 12):
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # Col T (20): Năm hạch toán
        c_ht = ws_ct.cell(idx, 20, ht_yr)
        c_ht.font = font_bold
        c_ht.alignment = Alignment(horizontal="center", vertical="center")
        c_ht.border = THIN_BORDER

        # Col U (21): Năm kỳ báo cáo
        c_ky = ws_ct.cell(idx, 21, int(ky_yr_str) if ky_yr_str.isdigit() else ky_yr_str)
        c_ky.font = font_body
        c_ky.alignment = Alignment(horizontal="center", vertical="center")
        c_ky.border = THIN_BORDER

        # Col V (22): Tên tiểu mục
        c_tmn = ws_ct.cell(idx, 22, tm_name)
        c_tmn.font = font_body
        c_tmn.border = THIN_BORDER

        # Col W (23): Số tiền đã nộp (VNĐ) = -G{idx}
        c_st = ws_ct.cell(idx, 23, f"=-G{idx}")
        c_st.font = font_bold
        c_st.number_format = "#,##0"
        c_st.alignment = Alignment(horizontal="right", vertical="center")
        c_st.border = THIN_BORDER

        # Col X (24): Nhóm sắc thuế
        c_grp = ws_ct.cell(idx, 24, tm_group)
        c_grp.font = font_body
        c_grp.alignment = Alignment(horizontal="center", vertical="center")
        c_grp.border = THIN_BORDER

    max_ct_row = len(raw_rows) + 1
    ws_ct.auto_filter.ref = f"A1:X{max_ct_row}"
    ws_ct.freeze_panes = "A2"
    for c_i in range(1, 25):
        col_let = get_column_letter(c_i)
        ws_ct.column_dimensions[col_let].width = 16
    ws_ct.column_dimensions["C"].width = 16
    ws_ct.column_dimensions["G"].width = 18
    ws_ct.column_dimensions["H"].width = 18
    ws_ct.column_dimensions["V"].width = 45
    ws_ct.column_dimensions["W"].width = 20
    ws_ct.column_dimensions["X"].width = 22

    # ==========================================================================
    # 2. GHI SHEET TỔNG HỢP 1: 'TH_Nam_HachToan' (Theo Năm Ngày hạch toán)
    # ==========================================================================
    ws_ht.cell(1, 1, "BẢNG TỔNG HỢP SỐ THUẾ ĐÃ NỘP NGÂN SÁCH NHÀ NƯỚC TỪNG NĂM THEO TỪNG TIỂU MỤC").font = font_title
    ws_ht.cell(2, 1, "1. Bảng tổng hợp theo từng Tiểu mục và Năm nộp (Năm hạch toán trên chứng từ nộp thuế)").font = font_sec

    hdr_row = 3
    ws_ht.row_dimensions[hdr_row].height = 28
    base_cols = ["STT", "Tiểu mục", "Tên tiểu mục / Nội dung kinh tế", "Nhóm sắc thuế"]
    for c_i, h_text in enumerate(base_cols, start=1):
        c = ws_ht.cell(hdr_row, c_i, h_text)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    for y_off, yr in enumerate(ht_years):
        c_i = 5 + y_off
        c = ws_ht.cell(hdr_row, c_i, yr)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    tot_col_idx = 5 + len(ht_years)
    tot_col_let = get_column_letter(tot_col_idx)
    first_yr_let = get_column_letter(5)
    last_yr_let = get_column_letter(tot_col_idx - 1)

    c_tot_hdr = ws_ht.cell(hdr_row, tot_col_idx, "Tổng cộng")
    c_tot_hdr.font = font_hdr
    c_tot_hdr.fill = HEADER_FILL
    c_tot_hdr.alignment = Alignment(horizontal="center", vertical="center")
    c_tot_hdr.border = THIN_BORDER

    start_tm_row = 4
    cur_r = start_tm_row
    for idx, tm in enumerate(tieu_mucs, start=1):
        ws_ht.row_dimensions[cur_r].height = 22
        tm_name, tm_group = TIEU_MUC_INFO.get(tm, (f"Tiểu mục {tm}", "Khác"))

        c1 = ws_ht.cell(cur_r, 1, idx)
        c1.font = font_body
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = THIN_BORDER

        c2 = ws_ht.cell(cur_r, 2, tm)
        c2.font = font_bold
        c2.number_format = "@"
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = THIN_BORDER

        c3 = ws_ht.cell(cur_r, 3, tm_name)
        c3.font = font_body
        c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c3.border = THIN_BORDER

        c4 = ws_ht.cell(cur_r, 4, tm_group)
        c4.font = font_body
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border = THIN_BORDER

        for y_off, _yr in enumerate(ht_years):
            c_i = 5 + y_off
            col_let = get_column_letter(c_i)
            formula = (
                f"=SUMIFS('Chi_Tiet_CTUNOP'!$W$2:$W${max_ct_row},"
                f"'Chi_Tiet_CTUNOP'!$E$2:$E${max_ct_row},$B{cur_r},"
                f"'Chi_Tiet_CTUNOP'!$T$2:$T${max_ct_row},{col_let}${hdr_row})"
            )
            cy = ws_ht.cell(cur_r, c_i, formula)
            cy.font = font_body
            cy.number_format = "#,##0"
            cy.alignment = Alignment(horizontal="right", vertical="center")
            cy.border = THIN_BORDER

        ct = ws_ht.cell(cur_r, tot_col_idx, f"=SUM({first_yr_let}{cur_r}:{last_yr_let}{cur_r})")
        ct.font = font_bold
        ct.number_format = "#,##0"
        ct.alignment = Alignment(horizontal="right", vertical="center")
        ct.border = THIN_BORDER
        cur_r += 1

    end_tm_row = cur_r - 1
    sum_tm_row = cur_r
    ws_ht.row_dimensions[sum_tm_row].height = 24
    ws_ht.merge_cells(start_row=sum_tm_row, start_column=1, end_row=sum_tm_row, end_column=4)
    for c_i in range(1, 5):
        c = ws_ht.cell(sum_tm_row, c_i)
        if c_i == 1:
            c.value = "TỔNG CỘNG SỐ ĐÃ NỘP"
        c.font = font_bold
        c.fill = TOTAL_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    for c_i in range(5, tot_col_idx + 1):
        col_let = get_column_letter(c_i)
        cs = ws_ht.cell(sum_tm_row, c_i, f"=SUM({col_let}{start_tm_row}:{col_let}{end_tm_row})")
        cs.font = font_bold
        cs.fill = TOTAL_FILL
        cs.number_format = "#,##0"
        cs.alignment = Alignment(horizontal="right", vertical="center")
        cs.border = THIN_BORDER

    # BẢNG 2: TỔNG HỢP THEO NHÓM SẮC THUẾ
    grp_title_row = sum_tm_row + 2
    ws_ht.cell(grp_title_row, 1, "2. Bảng tổng hợp theo Nhóm sắc thuế từng năm").font = font_sec

    grp_hdr_row = grp_title_row + 1
    ws_ht.row_dimensions[grp_hdr_row].height = 26
    grp_cols = ["STT", "Nhóm sắc thuế", "Các tiểu mục bao gồm", "Ghi chú"]
    for c_i, h_text in enumerate(grp_cols, start=1):
        c = ws_ht.cell(grp_hdr_row, c_i, h_text)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    for y_off, yr in enumerate(ht_years):
        c_i = 5 + y_off
        c = ws_ht.cell(grp_hdr_row, c_i, yr)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    c_gtot = ws_ht.cell(grp_hdr_row, tot_col_idx, "Tổng cộng")
    c_gtot.font = font_hdr
    c_gtot.fill = HEADER_FILL
    c_gtot.alignment = Alignment(horizontal="center", vertical="center")
    c_gtot.border = THIN_BORDER

    start_grp_row = grp_hdr_row + 1
    cur_r = start_grp_row
    for g_idx, g_name in enumerate(TAX_GROUP_ORDER, start=1):
        ws_ht.row_dimensions[cur_r].height = 22
        sub_tms = [tm for tm in tieu_mucs if TIEU_MUC_INFO.get(tm, ("", ""))[1] == g_name]

        c1 = ws_ht.cell(cur_r, 1, g_idx)
        c1.font = font_body
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = THIN_BORDER

        c2 = ws_ht.cell(cur_r, 2, g_name)
        c2.font = font_bold
        c2.alignment = Alignment(horizontal="left", vertical="center")
        c2.border = THIN_BORDER

        c3 = ws_ht.cell(cur_r, 3, ", ".join(sub_tms))
        c3.font = font_body
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = THIN_BORDER

        c4 = ws_ht.cell(cur_r, 4, "")
        c4.font = font_body
        c4.border = THIN_BORDER

        for y_off, _yr in enumerate(ht_years):
            c_i = 5 + y_off
            col_let = get_column_letter(c_i)
            formula = f"=SUMIF($D${start_tm_row}:$D${end_tm_row},$B{cur_r},{col_let}${start_tm_row}:{col_let}${end_tm_row})"
            cy = ws_ht.cell(cur_r, c_i, formula)
            cy.font = font_body
            cy.number_format = "#,##0"
            cy.alignment = Alignment(horizontal="right", vertical="center")
            cy.border = THIN_BORDER

        ct = ws_ht.cell(cur_r, tot_col_idx, f"=SUM({first_yr_let}{cur_r}:{last_yr_let}{cur_r})")
        ct.font = font_bold
        ct.number_format = "#,##0"
        ct.alignment = Alignment(horizontal="right", vertical="center")
        ct.border = THIN_BORDER
        cur_r += 1

    end_grp_row = cur_r - 1
    sum_grp_row = cur_r
    ws_ht.row_dimensions[sum_grp_row].height = 24
    ws_ht.merge_cells(start_row=sum_grp_row, start_column=1, end_row=sum_grp_row, end_column=4)
    for c_i in range(1, 5):
        c = ws_ht.cell(sum_grp_row, c_i)
        if c_i == 1:
            c.value = "TỔNG CỘNG TOÀN BỘ CÁC SẮC THUẾ"
        c.font = font_bold
        c.fill = TOTAL_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    for c_i in range(5, tot_col_idx + 1):
        col_let = get_column_letter(c_i)
        cs = ws_ht.cell(sum_grp_row, c_i, f"=SUM({col_let}{start_grp_row}:{col_let}{end_grp_row})")
        cs.font = font_bold
        cs.fill = TOTAL_FILL
        cs.number_format = "#,##0"
        cs.alignment = Alignment(horizontal="right", vertical="center")
        cs.border = THIN_BORDER

    # BẢNG 3: BẢNG KÊ DỌC THEO TỪNG NĂM & TIỂU MỤC (CHỈ LIỆT KÊ CÁC CẶP NĂM - TIỂU MỤC CÓ PHÁT SINH)
    v_title_row = sum_grp_row + 2
    ws_ht.cell(v_title_row, 1, "3. Bảng kê chi tiết từng Năm hạch toán theo từng Tiểu mục").font = font_sec
    v_hdr_row = v_title_row + 1
    ws_ht.row_dimensions[v_hdr_row].height = 26
    v_headers = ["STT", "Năm nộp", "Tiểu mục", "Tên tiểu mục / Nội dung kinh tế", "Nhóm sắc thuế", "Số tiền đã nộp (VNĐ)"]
    for c_i, h_text in enumerate(v_headers, start=1):
        c = ws_ht.cell(v_hdr_row, c_i, h_text)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    v_start_row = v_hdr_row + 1
    cur_r = v_start_row
    v_stt = 1
    for yr in ht_years:
        for tm in tieu_mucs:
            if by_ht_tm.get((yr, tm), 0.0) == 0:
                continue
            ws_ht.row_dimensions[cur_r].height = 21
            tm_name, tm_group = TIEU_MUC_INFO.get(tm, (f"Tiểu mục {tm}", "Khác"))

            c1 = ws_ht.cell(cur_r, 1, v_stt)
            c1.font = font_body
            c1.alignment = Alignment(horizontal="center", vertical="center")
            c1.border = THIN_BORDER

            c2 = ws_ht.cell(cur_r, 2, yr)
            c2.font = font_bold
            c2.alignment = Alignment(horizontal="center", vertical="center")
            c2.border = THIN_BORDER

            c3 = ws_ht.cell(cur_r, 3, tm)
            c3.font = font_bold
            c3.number_format = "@"
            c3.alignment = Alignment(horizontal="center", vertical="center")
            c3.border = THIN_BORDER

            c4 = ws_ht.cell(cur_r, 4, tm_name)
            c4.font = font_body
            c4.alignment = Alignment(horizontal="left", vertical="center")
            c4.border = THIN_BORDER

            c5 = ws_ht.cell(cur_r, 5, tm_group)
            c5.font = font_body
            c5.alignment = Alignment(horizontal="center", vertical="center")
            c5.border = THIN_BORDER

            formula = (
                f"=SUMIFS('Chi_Tiet_CTUNOP'!$W$2:$W${max_ct_row},"
                f"'Chi_Tiet_CTUNOP'!$T$2:$T${max_ct_row},$B{cur_r},"
                f"'Chi_Tiet_CTUNOP'!$E$2:$E${max_ct_row},$C{cur_r})"
            )
            c6 = ws_ht.cell(cur_r, 6, formula)
            c6.font = font_body
            c6.number_format = "#,##0"
            c6.alignment = Alignment(horizontal="right", vertical="center")
            c6.border = THIN_BORDER

            cur_r += 1
            v_stt += 1

    v_end_row = cur_r - 1
    v_sum_row = cur_r
    ws_ht.row_dimensions[v_sum_row].height = 24
    ws_ht.merge_cells(start_row=v_sum_row, start_column=1, end_row=v_sum_row, end_column=5)
    for c_i in range(1, 6):
        c = ws_ht.cell(v_sum_row, c_i)
        if c_i == 1:
            c.value = "TỔNG CỘNG"
        c.font = font_bold
        c.fill = TOTAL_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    c_vtot = ws_ht.cell(v_sum_row, 6, f"=SUM(F{v_start_row}:F{v_end_row})")
    c_vtot.font = font_bold
    c_vtot.fill = TOTAL_FILL
    c_vtot.number_format = "#,##0"
    c_vtot.alignment = Alignment(horizontal="right", vertical="center")
    c_vtot.border = THIN_BORDER

    ws_ht.column_dimensions["A"].width = 7
    ws_ht.column_dimensions["B"].width = 14
    ws_ht.column_dimensions["C"].width = 52
    ws_ht.column_dimensions["D"].width = 22
    for c_i in range(5, tot_col_idx + 1):
        ws_ht.column_dimensions[get_column_letter(c_i)].width = 18
    ws_ht.freeze_panes = "E4"

    # ==========================================================================
    # 3. GHI SHEET TỔNG HỢP 2: 'TH_Nam_KyBaoCao' (Theo Năm Kỳ báo cáo)
    # ==========================================================================
    ws_ky.cell(1, 1, "BẢNG TỔNG HỢP SỐ THUẾ ĐÃ NỘP THEO NĂM CỦA KỲ BÁO CÁO & TIỂU MỤC").font = font_title
    ws_ky.cell(2, 1, "Ghi chú: Phân bổ theo cột 'Kỳ báo cáo' trên chứng từ (Các chứng từ để trống Kỳ báo cáo được tách riêng ở cột 'Không ghi kỳ')").font = font_sec

    ws_ky.row_dimensions[3].height = 28
    for c_i, h_text in enumerate(base_cols, start=1):
        c = ws_ky.cell(3, c_i, h_text)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    for y_off, ky_yr in enumerate(ky_years):
        c_i = 5 + y_off
        val_hdr = int(ky_yr) if ky_yr.isdigit() else ky_yr
        c = ws_ky.cell(3, c_i, val_hdr)
        c.font = font_hdr
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = THIN_BORDER

    ky_tot_idx = 5 + len(ky_years)
    ky_first_let = get_column_letter(5)
    ky_last_let = get_column_letter(ky_tot_idx - 1)
    c_kytot = ws_ky.cell(3, ky_tot_idx, "Tổng cộng")
    c_kytot.font = font_hdr
    c_kytot.fill = HEADER_FILL
    c_kytot.alignment = Alignment(horizontal="center", vertical="center")
    c_kytot.border = THIN_BORDER

    cur_r = 4
    for idx, tm in enumerate(tieu_mucs, start=1):
        ws_ky.row_dimensions[cur_r].height = 22
        tm_name, tm_group = TIEU_MUC_INFO.get(tm, (f"Tiểu mục {tm}", "Khác"))

        c1 = ws_ky.cell(cur_r, 1, idx)
        c1.font = font_body
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = THIN_BORDER

        c2 = ws_ky.cell(cur_r, 2, tm)
        c2.font = font_bold
        c2.number_format = "@"
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = THIN_BORDER

        c3 = ws_ky.cell(cur_r, 3, tm_name)
        c3.font = font_body
        c3.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        c3.border = THIN_BORDER

        c4 = ws_ky.cell(cur_r, 4, tm_group)
        c4.font = font_body
        c4.alignment = Alignment(horizontal="center", vertical="center")
        c4.border = THIN_BORDER

        for y_off, _ky_yr in enumerate(ky_years):
            c_i = 5 + y_off
            col_let = get_column_letter(c_i)
            formula = (
                f"=SUMIFS('Chi_Tiet_CTUNOP'!$W$2:$W${max_ct_row},"
                f"'Chi_Tiet_CTUNOP'!$E$2:$E${max_ct_row},$B{cur_r},"
                f"'Chi_Tiet_CTUNOP'!$U$2:$U${max_ct_row},{col_let}$3)"
            )
            cy = ws_ky.cell(cur_r, c_i, formula)
            cy.font = font_body
            cy.number_format = "#,##0"
            cy.alignment = Alignment(horizontal="right", vertical="center")
            cy.border = THIN_BORDER

        ct = ws_ky.cell(cur_r, ky_tot_idx, f"=SUM({ky_first_let}{cur_r}:{ky_last_let}{cur_r})")
        ct.font = font_bold
        ct.number_format = "#,##0"
        ct.alignment = Alignment(horizontal="right", vertical="center")
        ct.border = THIN_BORDER
        cur_r += 1

    ky_end_row = cur_r - 1
    ky_sum_row = cur_r
    ws_ky.row_dimensions[ky_sum_row].height = 24
    ws_ky.merge_cells(start_row=ky_sum_row, start_column=1, end_row=ky_sum_row, end_column=4)
    for c_i in range(1, 5):
        c = ws_ky.cell(ky_sum_row, c_i)
        if c_i == 1:
            c.value = "TỔNG CỘNG SỐ ĐÃ NỘP"
        c.font = font_bold
        c.fill = TOTAL_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = THIN_BORDER

    for c_i in range(5, ky_tot_idx + 1):
        col_let = get_column_letter(c_i)
        cs = ws_ky.cell(ky_sum_row, c_i, f"=SUM({col_let}4:{col_let}{ky_end_row})")
        cs.font = font_bold
        cs.fill = TOTAL_FILL
        cs.number_format = "#,##0"
        cs.alignment = Alignment(horizontal="right", vertical="center")
        cs.border = THIN_BORDER

    ws_ky.column_dimensions["A"].width = 7
    ws_ky.column_dimensions["B"].width = 14
    ws_ky.column_dimensions["C"].width = 52
    ws_ky.column_dimensions["D"].width = 22
    for c_i in range(5, ky_tot_idx + 1):
        ws_ky.column_dimensions[get_column_letter(c_i)].width = 18
    ws_ky.freeze_panes = "E4"

    for out_p in output_paths:
        os.makedirs(os.path.dirname(os.path.abspath(out_p)), exist_ok=True)
        wb_out.save(out_p)

    return {
        "ht_years": ht_years,
        "ky_years": ky_years,
        "tieu_mucs": tieu_mucs,
        "by_ht_tm": by_ht_tm,
        "by_ky_tm": by_ky_tm,
        "total_rows": len(raw_rows),
    }


if __name__ == "__main__":
    src = r"C:\Users\Admin\.gemini\antigravity\brain\74addb61-c896-4052-bc3d-ededa129ccfc\.user_uploaded\media_1795265102292_afb374a5.xlsx"
    if not os.path.exists(src):
        src = r"D:\2026\MANY\TCS_MANY\TCS_MANY\0313248400-CTUNOP.XLSX"
    outs = [
        r"D:\2026\MANY\TCS_MANY\TCS_MANY\TongHop_SoNop_TheoNam_TieuMuc.xlsx",
        r"C:\Users\Admin\.gemini\antigravity\scratch\tax_xml_to_excel\TongHop_SoNop_TheoNam_TieuMuc.xlsx",
    ]
    res = summarize_tax_payments_to_excel(src, outs)
    print("Exported successfully:", outs, "Rows:", res["total_rows"])
