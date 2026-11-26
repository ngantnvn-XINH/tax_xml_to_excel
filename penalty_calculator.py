# -*- coding: utf-8 -*-
"""
Module tạo sheet 'Tinh_TienChamNop':
BẢNG CHI TIẾT SỐ THUẾ TRUY THU, XỬ PHẠT VI PHẠM HÀNH CHÍNH VÀ TIỀN CHẬM NỘP SAU THANH TRA - KIỂM TRA THUẾ.

Thiết kế chuẩn nghiệp vụ & tự động cập nhật theo quy định pháp luật qua từng thời kỳ:
1. Tỷ lệ tính tiền chậm nộp theo giai đoạn:
   - Trước 01/01/2015: 0,05%/ngày (<=90 ngày), 0,07%/ngày (>90 ngày) [Luật 21/2012/QH13]
   - Từ 01/01/2015 đến 30/06/2016: 0,05%/ngày [Luật 71/2014/QH14]
   - Từ 01/07/2016 đến nay: 0,03%/ngày [Luật 106/2016/QH13 & Khoản 2 Điều 59 Luật QLT 38/2019/QH14]
2. Hạn nộp hồ sơ khai thuế & Ngày bắt đầu tính chậm nộp ('Từ ngày' = hạn nộp + 1 ngày):
   - Thuế GTGT kê khai tháng: hạn ngày 20 tháng sau -> Từ ngày 21 tháng sau (Tháng 12 -> 21/01 năm sau)
   - Thuế GTGT kê khai quý:
     + Từ 01/07/2020 (Luật 38/2019/QH14): hạn ngày cuối tháng đầu quý sau -> 01/05, 01/08, 01/11, 01/02 năm sau
     + Trước 01/07/2020 (Luật 78/2006): hạn ngày 30 tháng đầu quý sau -> 01/05, 31/07, 31/10, 31/01 năm sau
   - Thuế TNDN quyết toán năm: hạn ngày cuối tháng thứ 3 kể từ ngày kết thúc năm dương lịch/năm tài chính
     + Năm dương lịch (kết thúc 31/12): Từ ngày 01/04 năm sau
     + Năm tài chính lệch kỳ (VD kết thúc 30/06): Tự động tính ngày 01 tháng thứ 4 sau ngày kết thúc năm tài chính
   - Thuế TNCN quyết toán năm: hạn 31/03 năm sau -> Từ ngày 01/04 năm sau
   - Lệ phí môn bài: hạn 30/01 hàng năm (NĐ 139/2016 & NĐ 22/2020) -> Từ ngày 31/01
   - Thuế Nhà thầu (GTGT & TNDN): để trống ngày để nhập theo từng lần phát sinh (+10 ngày) hoặc tháng (+20 ngày)
3. Xử phạt VPHC & Thời hiệu (Nghị định 125/2020/NĐ-CP):
   - Phạt khai sai dẫn đến thiếu thuế: 20% (Điều 16)
   - Phạt trốn thuế: 1,0 - 3,0 lần (Điều 17, mặc định 1,5 lần khi không có tình tiết tăng nặng/giảm nhẹ)
   - Phạt VPHC về thủ tục thuế: Điều 13, 14... (mục riêng, không trộn lẫn vào cột Phạt 20%)
   - Thời hiệu xử phạt: 05 năm đối với hành vi khai sai thiếu thuế/trốn thuế; 02 năm đối với thủ tục thuế (Điều 8)
     -> Quá 05 năm: tự động đưa tiền phạt 20% & phạt trốn thuế về 0, cảnh báo trên cột Thời hiệu, vẫn tính truy thu & tiền chậm nộp.
   - Thời hạn truy thu thuế: 10 năm trở về trước kể từ ngày phát hiện vi phạm (Khoản 6 Điều 8).
"""

from __future__ import annotations

import calendar
import datetime
import re
from typing import Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SHEET_NAME_PENALTY = "Tinh_TienChamNop"
RATE_CHANGE_DATE = datetime.date(2016, 6, 30)  # Hết ngày 30/06/2016 áp dụng 0,05%; từ 01/07/2016 áp dụng 0,03%

ROMAN_NUMERALS = [
    "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X",
    "XI", "XII", "XIII", "XIV", "XV", "XVI", "XVII", "XVIII", "XIX", "XX",
]


def _to_roman(idx: int) -> str:
    if 1 <= idx <= len(ROMAN_NUMERALS):
        return ROMAN_NUMERALS[idx - 1]
    return str(idx)


def compute_vat_monthly_start_date(year: int, month: int) -> datetime.date:
    """Ngày bắt đầu tính chậm nộp thuế GTGT kê khai tháng: ngày 21 tháng tiếp theo."""
    if month == 12:
        return datetime.date(year + 1, 1, 21)
    return datetime.date(year, month + 1, 21)


def compute_vat_quarterly_start_date(year: int, quarter: int) -> datetime.date:
    """
    Ngày bắt đầu tính chậm nộp thuế GTGT kê khai quý:
    - Từ kỳ Q2/2020 trở đi (hạn nộp sau 01/07/2020 theo Luật 38/2019/QH14):
      hạn nộp là ngày cuối cùng tháng đầu quý sau -> ngày bắt đầu tính chậm nộp là ngày 01 tháng thứ 2 quý sau.
    - Trước kỳ Q2/2020 (hạn nộp ngày 30 tháng đầu quý sau theo Luật 78/2006):
      Q1 -> 01/05, Q2 -> 31/07, Q3 -> 31/10, Q4 -> 31/01 năm sau.
    """
    is_law_38 = (year > 2020) or (year == 2020 and quarter >= 2)
    if is_law_38:
        if quarter == 1:
            return datetime.date(year, 5, 1)
        if quarter == 2:
            return datetime.date(year, 8, 1)
        if quarter == 3:
            return datetime.date(year, 11, 1)
        return datetime.date(year + 1, 2, 1)
    else:
        if quarter == 1:
            return datetime.date(year, 5, 1)
        if quarter == 2:
            return datetime.date(year, 7, 31)
        if quarter == 3:
            return datetime.date(year, 10, 31)
        return datetime.date(year + 1, 1, 31)


def compute_cit_start_date(
    year: int,
    fiscal_end_date: Optional[datetime.date] = None,
) -> datetime.date:
    """
    Ngày bắt đầu tính chậm nộp quyết toán thuế TNDN (Điều 44 Luật 38/2019/QH14):
    Hạn nộp là ngày cuối cùng của tháng thứ 3 kể từ ngày kết thúc năm dương lịch hoặc năm tài chính.
    -> Ngày bắt đầu tính chậm nộp là ngày 01 của tháng thứ 4 sau ngày kết thúc năm tài chính.
    """
    if fiscal_end_date is not None:
        end_y = fiscal_end_date.year
        end_m = fiscal_end_date.month
    else:
        end_y = year
        end_m = 12

    # Cộng 3 tháng để ra tháng hạn nộp, rồi +1 ngày sang ngày 01 tháng kế tiếp (tháng thứ 4)
    total_m = end_m + 4
    start_y = end_y + (total_m - 1) // 12
    start_m = ((total_m - 1) % 12) + 1
    return datetime.date(start_y, start_m, 1)


def compute_pit_start_date(year: int) -> datetime.date:
    """Ngày bắt đầu tính chậm nộp quyết toán thuế TNCN tổ chức chi trả: 01/04 năm sau."""
    return datetime.date(year + 1, 4, 1)


def compute_license_fee_start_date(year: int) -> datetime.date:
    """Ngày bắt đầu tính chậm nộp Lệ phí môn bài: 31/01 hàng năm (hạn nộp 30/01)."""
    return datetime.date(year, 1, 31)


def adjust_start_date_for_weekend(start_date: datetime.date) -> datetime.date:
    """
    Điều chỉnh ngày bắt đầu tính chậm nộp theo Điều 86 TT 80/2021/TT-BTC:
    Nếu ngày cuối cùng của thời hạn nộp thuế (start_date - 1 ngày) rơi vào Thứ 7 hoặc Chủ nhật,
    thời hạn nộp được lùi sang ngày làm việc kế tiếp (Thứ 2) -> ngày bắt đầu tính chậm nộp là Thứ 3.
    """
    deadline = start_date - datetime.timedelta(days=1)
    if deadline.weekday() == 5:  # Thứ 7 -> hạn nộp Thứ 2 -> tính chậm nộp từ Thứ 3
        return start_date + datetime.timedelta(days=2)
    if deadline.weekday() == 6:  # Chủ nhật -> hạn nộp Thứ 2 -> tính chậm nộp từ Thứ 3
        return start_date + datetime.timedelta(days=1)
    return start_date


def detect_vat_frequency_from_workbook(wb: openpyxl.Workbook) -> str:
    """
    Tự động nhận diện kỳ kê khai GTGT ('month' hoặc 'quarter') từ sheet 'So Thue GTGT Cuoi'.
    Mặc định trả về 'month' nếu không tìm thấy hoặc doanh nghiệp kê khai tháng.
    """
    ws = None
    for sname in ("So Thue GTGT Cuoi", "So Thue GTGT TH"):
        if sname in wb.sheetnames:
            ws = wb[sname]
            break
    if ws is None:
        return "month"

    quarter_hits = 0
    month_hits = 0
    for r in range(2, min(ws.max_row + 1, 250)):
        for c in range(1, 5):
            val = ws.cell(r, c).value
            if val is None:
                continue
            s = str(val).strip()
            if re.match(r"^Q[1-4]/\d{4}$", s, re.IGNORECASE) or re.match(r"^Quý\s*[1-4]", s, re.IGNORECASE):
                quarter_hits += 1
            elif re.match(r"^(0?[1-9]|1[0-2])/\d{4}$", s):
                month_hits += 1
    if quarter_hits > 0 and quarter_hits >= month_hits:
        return "quarter"
    return "month"


def detect_years_from_workbook(wb: openpyxl.Workbook) -> List[int]:
    """Tự động dò danh sách các năm thanh tra/kiểm tra từ các sheet năm ('2020'..'2026') trong workbook."""
    years: List[int] = []
    for sname in wb.sheetnames:
        s = sname.strip()
        if re.fullmatch(r"20\d{2}", s):
            years.append(int(s))
    years.sort()
    return years


def detect_fiscal_end_dates_from_workbook(
    wb: openpyxl.Workbook,
    years: Sequence[int],
) -> Dict[int, datetime.date]:
    """
    Tự động dò ngày kết thúc năm tài chính (cho doanh nghiệp lệch kỳ) từ sheet 'So Thue GTGT Cuoi (Lech ky)'
    hoặc tiêu đề cột trên sheet 'KKTNDN' (nếu có dạng 'Năm YYYY (MM/YYYY - MM/YYYY)').
    """
    result: Dict[int, datetime.date] = {}
    if "So Thue GTGT Cuoi (Lech ky)" in wb.sheetnames:
        ws = wb["So Thue GTGT Cuoi (Lech ky)"]
        for r in range(1, min(ws.max_row + 1, 200)):
            val = ws.cell(r, 1).value or ws.cell(r, 2).value
            if not val:
                continue
            m = re.search(r"Năm\s*(20\d{2}).*?(\d{2})/(20\d{2})\s*[-–]\s*(\d{2})/(20\d{2})", str(val))
            if m:
                yr = int(m.group(1))
                em = int(m.group(4))
                ey = int(m.group(5))
                last_day = calendar.monthrange(ey, em)[1]
                result[yr] = datetime.date(ey, em, last_day)
    return result


def parse_cutoff_date(date_str: Optional[str]) -> datetime.date:
    """Chuyển chuỗi ngày dd/mm/yyyy hoặc yyyy-mm-dd sang datetime.date; mặc định trả về hôm nay."""
    if not date_str or not str(date_str).strip():
        return datetime.date.today()
    s = str(date_str).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return datetime.date.today()


def add_late_payment_sheet(
    wb: openpyxl.Workbook,
    years: Optional[Sequence[int]] = None,
    cutoff_date: Optional[datetime.date] = None,
    vat_frequency: Optional[str] = None,
    fiscal_ranges: Optional[Dict[int, Tuple[datetime.date, datetime.date]]] = None,
    sheet_name: str = SHEET_NAME_PENALTY,
) -> openpyxl.worksheet.worksheet.Worksheet:
    """
    Tạo mới (hoặc ghi đè) sheet 'Tinh_TienChamNop' trong workbook `wb`.
    - Nếu `years` là None, tự động nhận diện từ các sheet năm trong `wb`.
    - Nếu `cutoff_date` là None, mặc định ngày hôm nay (người dùng chỉ cần sửa 1 ô D4 trong Excel).
    - Nếu `vat_frequency` là None, tự động nhận diện 'month' hay 'quarter' từ sheet GTGT.
    - Nếu sheet năm 'YYYY' tồn tại trong `wb`, tự động nối công thức cột Thuế truy thu cho TNDN ('YYYY'!E76)
      và TNCN ('YYYY'!E90) khi có chênh lệch dương sau kiểm tra.
    """
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    target_years = sorted(set(years)) if years else detect_years_from_workbook(wb)
    if not target_years:
        target_years = [datetime.date.today().year]

    freq = vat_frequency or detect_vat_frequency_from_workbook(wb)
    if freq not in ("month", "quarter"):
        freq = "month"

    cutoff = cutoff_date or datetime.date.today()

    fiscal_ends: Dict[int, datetime.date] = detect_fiscal_end_dates_from_workbook(wb, target_years)
    if fiscal_ranges:
        for yr, (_, end_d) in fiscal_ranges.items():
            if isinstance(end_d, datetime.date):
                fiscal_ends[yr] = end_d

    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    # ===== STYLES =====
    FONT_NAME = "Times New Roman"
    f_title = Font(name=FONT_NAME, size=13, bold=True, color="1F4E78")
    f_subtitle = Font(name=FONT_NAME, size=10, italic=True, color="333333")
    f_sec = Font(name=FONT_NAME, size=10, bold=True, color="1F4E78")
    f_hdr = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    f_hdr_sub = Font(name=FONT_NAME, size=9, bold=True, color="1F4E78")
    f_yr = Font(name=FONT_NAME, size=10, bold=True, color="000000")
    f_grp = Font(name=FONT_NAME, size=10, bold=True, color="1F4E78")
    f_reg = Font(name=FONT_NAME, size=10)
    f_bold = Font(name=FONT_NAME, size=10, bold=True)
    f_italic = Font(name=FONT_NAME, size=9, italic=True, color="444444")
    f_total = Font(name=FONT_NAME, size=10, bold=True, color="FFFFFF")
    f_note = Font(name=FONT_NAME, size=9, color="222222")

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_hdr2 = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_yr = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    fill_grp = PatternFill(start_color="EEF4FA", end_color="EEF4FA", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_ctrl = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_total = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_sum_hdr = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_law_hdr = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
    fill_law_alt = PatternFill(start_color="F5F9F2", end_color="F5F9F2", fill_type="solid")

    thin = Side(style="thin", color="808080")
    med_navy = Side(style="medium", color="1F4E78")
    border_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_hdr = Border(left=thin, right=thin, top=med_navy, bottom=med_navy)

    al_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_right = Alignment(horizontal="right", vertical="center")

    NUM_FMT = "#,##0"
    DATE_FMT = "dd/mm/yyyy"
    PCT_FMT = "0.00%"
    DEC_FMT = "0.0"

    # ===== TITLE BLOCK (Rows 1-2) =====
    ws.merge_cells("A1:P1")
    ws["A1"] = "BẢNG CHI TIẾT SỐ THUẾ TRUY THU, XỬ PHẠT VI PHẠM HÀNH CHÍNH VÀ TIỀN CHẬM NỘP"
    ws["A1"].font = f_title
    ws["A1"].alignment = al_center
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:P2")
    yr_span = f"{target_years[0]} – {target_years[-1]}" if len(target_years) > 1 else f"{target_years[0]}"
    ws["A2"] = (
        f"(Kỳ thanh tra, kiểm tra: Năm {yr_span} | Kỳ GTGT: "
        f"{'Theo Quý' if freq == 'quarter' else 'Theo Tháng'} | "
        f"Đơn vị tính: Đồng — Chỉ cần sửa ô màu vàng D4 để cập nhật toàn bảng)"
    )
    ws["A2"].font = f_subtitle
    ws["A2"].alignment = al_center
    ws.row_dimensions[2].height = 18

    # ===== CONTROL PANEL (Rows 4-7) =====
    ctrl_left = [
        (4, "Ngày tính chậm nộp / Ngày lập biên bản (Đến ngày):", cutoff, DATE_FMT, True,
         "Thay đổi duy nhất ô này để tự động cập nhật số ngày chậm nộp, tiền chậm nộp và kiểm tra thời hiệu 5 năm/10 năm cho toàn bộ các dòng."),
        (5, "Mốc chuyển đổi tỷ lệ chậm nộp (Luật 106/2016):", RATE_CHANGE_DATE, DATE_FMT, False,
         "Hết 30/06/2016: 0,05%/ngày (Luật 71/2014). Từ 01/07/2016 trở đi: 0,03%/ngày (Luật 106/2016 & Luật 38/2019)."),
        (6, "Tỷ lệ chậm nộp từ 01/07/2016 (Điều 59 Luật 38/2019):", 0.0003, PCT_FMT, False, None),
        (7, "Tỷ lệ chậm nộp 01/01/2015 – 30/06/2016 (Luật 71/2014):", 0.0005, PCT_FMT, False, None),
    ]
    ctrl_right = [
        (4, "Tỷ lệ phạt khai sai thiếu thuế (Điều 16 NĐ 125/2020):", 0.20, "0%", False, None),
        (5, "Hệ số phạt trốn thuế mặc định (Điều 17 NĐ 125/2020):", 1.5, DEC_FMT, True,
         "1,0 lần (1 tình tiết giảm nhẹ); 1,5 lần (không tăng/giảm nhẹ); 2,0 - 2,5 - 3,0 lần (1 - 2 - >=3 tình tiết tăng nặng). Có thể sửa riêng trên từng dòng tại cột G."),
        (6, "Thời hiệu xử phạt khai sai / trốn thuế (năm - Điều 8):", 5, "0", False,
         "Quá 5 năm kể từ ngày thực hiện hành vi vi phạm (Từ ngày) thì không xử phạt VPHC (cột E, H tự về 0), vẫn truy thu thuế và tính tiền chậm nộp."),
        (7, "Thời hạn truy thu thuế (năm - Khoản 6 Điều 8):", 10, "0", False,
         "Thời hạn truy thu thuế tối đa 10 năm trở về trước kể từ ngày phát hiện hành vi vi phạm."),
    ]

    for row_idx, label, val, nfmt, is_input, note in ctrl_left:
        ws.merge_cells(start_row=row_idx, start_column=1, end_row=row_idx, end_column=3)
        c_lbl = ws.cell(row_idx, 1, label)
        c_lbl.font = f_bold if is_input else f_reg
        c_lbl.alignment = al_left
        c_lbl.fill = fill_ctrl
        for cc in range(1, 4):
            ws.cell(row_idx, cc).border = border_thin
            ws.cell(row_idx, cc).fill = fill_ctrl

        c_val = ws.cell(row_idx, 4, val)
        c_val.font = f_bold
        c_val.number_format = nfmt
        c_val.alignment = al_center
        c_val.fill = fill_input if is_input else fill_ctrl
        c_val.border = border_thin
        if note:
            c_val.comment = Comment(note, "Tinh_TienChamNop")
        ws.row_dimensions[row_idx].height = 19

    for row_idx, label, val, nfmt, is_input, note in ctrl_right:
        ws.merge_cells(start_row=row_idx, start_column=6, end_row=row_idx, end_column=7)
        c_lbl = ws.cell(row_idx, 6, label)
        c_lbl.font = f_bold if is_input else f_reg
        c_lbl.alignment = al_left
        for cc in (6, 7):
            ws.cell(row_idx, cc).border = border_thin
            ws.cell(row_idx, cc).fill = fill_ctrl

        c_val = ws.cell(row_idx, 8, val)
        c_val.font = f_bold
        c_val.number_format = nfmt
        c_val.alignment = al_center
        c_val.fill = fill_input if is_input else fill_ctrl
        c_val.border = border_thin
        if note:
            c_val.comment = Comment(note, "Tinh_TienChamNop")

    # Quick legend on right side of Control Panel (Cols J-P)
    ws.merge_cells("J4:P7")
    legend_cell = ws["J4"]
    legend_cell.value = (
        "HƯỚNG DẪN NHẬP LIỆU NHANH:\n"
        "• Ô nền vàng: Ô nhập số liệu truy thu / ngày tính chậm nộp / hệ số trốn thuế.\n"
        "• Thuế TNDN & TNCN: Tự động liên kết cột Chênh lệch kiểm tra ('YYYY'!E76 & 'YYYY'!E90) nếu có trong file.\n"
        "• Quá 5 năm (từ 'Từ ngày' đến ô D4): Tiền phạt 20% & Phạt trốn thuế tự động về 0, giữ nguyên truy thu & chậm nộp."
    )
    legend_cell.font = f_note
    legend_cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for rr in range(4, 8):
        for cc in range(10, 17):
            ws.cell(rr, cc).fill = fill_grp
            ws.cell(rr, cc).border = border_thin

    # ===== TABLE HEADERS (Rows 9-12) =====
    r_h1, r_h2, r_h3, r_code = 9, 10, 11, 12
    ws.row_dimensions[8].height = 8
    ws.row_dimensions[r_h1].height = 22
    ws.row_dimensions[r_h2].height = 22
    ws.row_dimensions[r_h3].height = 30
    ws.row_dimensions[r_code].height = 18

    # Top header merges
    merges = [
        ("A9:A11", "STT"),
        ("B9:B11", "Kỳ kê khai /\nSắc thuế"),
        ("C9:C11", "Số thuế\ntruy thu\n(đồng)"),
        ("D9:N9", "CHI TIẾT TÍNH TIỀN XỬ LÝ VI PHẠM HÀNH CHÍNH VÀ TIỀN CHẬM NỘP"),
        ("D10:E10", "Phạt khai sai 20%\n(Điều 16 NĐ 125/2020)"),
        ("F10:H10", "Phạt trốn thuế 1–3 lần\n(Điều 17 NĐ 125/2020)"),
        ("I10:M10", "Tính tiền chậm nộp theo từng giai đoạn\n(Luật 71/2014, Luật 106/2016 & Luật 38/2019)"),
        ("N10:N11", "Cộng tiền xử lý\nVPHC & chậm nộp\n(đồng)"),
        ("O9:O11", "Tổng tiền thuế\ntruy thu & xử lý\n(đồng)"),
        ("P9:P11", "Tình trạng\nthời hiệu\n(5 năm / 10 năm)"),
    ]
    for rng, text in merges:
        ws.merge_cells(rng)
        top_left = rng.split(":")[0]
        ws[top_left] = text

    sub_headers = {
        "D11": "Thuế thiếu\ntính phạt 20%",
        "E11": "Tiền phạt\n20% (đồng)",
        "F11": "Thuế trốn\ntính phạt",
        "G11": "Hệ số\n(1–3 lần)",
        "H11": "Tiền phạt\ntrốn thuế (đồng)",
        "I11": "Từ ngày\n(sau hạn nộp)",
        "J11": "Đến ngày\n(=$D$4)",
        "K11": "Số ngày\n≤ 30/06/2016\n(0,05%)",
        "L11": "Số ngày\n≥ 01/07/2016\n(0,03%)",
        "M11": "Tiền chậm nộp\n(đồng)",
    }
    for coord, text in sub_headers.items():
        ws[coord] = text

    codes = [
        "(A)", "(B)", "(1)", "(2)", "(3=2×20%)", "(4)", "(5)", "(6=4×5)",
        "(7)", "(8)", "(9)", "(10)", "(11)", "(12=3+6+11)", "(13=1+12)", "(14)"
    ]
    for col_idx, code_txt in enumerate(codes, start=1):
        c = ws.cell(r_code, col_idx, code_txt)
        c.font = f_hdr_sub
        c.fill = fill_hdr2
        c.alignment = al_center
        c.border = border_thin

    for rr in (r_h1, r_h2, r_h3):
        for cc in range(1, 17):
            cell = ws.cell(rr, cc)
            cell.font = f_hdr
            cell.fill = fill_hdr
            cell.alignment = al_center
            cell.border = border_hdr

    # ===== HELPER TO WRITE A DETAIL CALCULATION ROW =====
    def _write_detail_row(
        row_idx: int,
        stt: object,
        label: str,
        tu_ngay: Optional[datetime.date],
        tax_formula_or_val: object = 0,
        is_input_tax: bool = True,
        is_group_header_style: bool = False,
        editable_date: bool = False,
    ) -> None:
        ws.row_dimensions[row_idx].height = 19
        row_font = f_grp if is_group_header_style else f_reg

        c_stt = ws.cell(row_idx, 1, stt)
        c_stt.font = f_bold if stt != "" else f_reg
        c_stt.alignment = al_center

        c_lbl = ws.cell(row_idx, 2, label)
        c_lbl.font = row_font
        c_lbl.alignment = al_left

        # Col C (1): Thuế truy thu
        c_tax = ws.cell(row_idx, 3, tax_formula_or_val)
        c_tax.font = row_font
        c_tax.number_format = NUM_FMT
        c_tax.alignment = al_right
        if is_input_tax:
            c_tax.fill = fill_input

        # Col D (2): Thuế thiếu tính phạt 20% = MAX(0, C - F)
        c_d = ws.cell(row_idx, 4, f"=MAX(0,C{row_idx}-F{row_idx})")
        c_d.font = row_font
        c_d.number_format = NUM_FMT
        c_d.alignment = al_right

        # Col E (3): Tiền phạt khai sai 20% (tự về 0 nếu quá 5 năm theo Điều 8 NĐ 125/2020)
        c_e = ws.cell(
            row_idx,
            5,
            f'=IF(D{row_idx}<=0,0,IF(AND(I{row_idx}<>"",EDATE(I{row_idx},$H$6*12)<$D$4),0,ROUND(D{row_idx}*$H$4,0)))',
        )
        c_e.font = row_font
        c_e.number_format = NUM_FMT
        c_e.alignment = al_right

        # Col F (4): Thuế trốn tính phạt (input, mặc định 0)
        c_f = ws.cell(row_idx, 6, 0)
        c_f.font = row_font
        c_f.number_format = NUM_FMT
        c_f.alignment = al_right
        c_f.fill = fill_input

        # Col G (5): Hệ số phạt trốn thuế (mặc định =$H$5 = 1.5)
        c_g = ws.cell(row_idx, 7, "=$H$5")
        c_g.font = row_font
        c_g.number_format = DEC_FMT
        c_g.alignment = al_center

        # Col H (6): Tiền phạt trốn thuế (tự về 0 nếu quá 5 năm)
        c_h = ws.cell(
            row_idx,
            8,
            f'=IF(F{row_idx}<=0,0,IF(AND(I{row_idx}<>"",EDATE(I{row_idx},$H$6*12)<$D$4),0,ROUND(F{row_idx}*G{row_idx},0)))',
        )
        c_h.font = row_font
        c_h.number_format = NUM_FMT
        c_h.alignment = al_right

        # Col I (7): Từ ngày
        c_i = ws.cell(row_idx, 9, tu_ngay if tu_ngay is not None else "")
        c_i.font = row_font
        c_i.number_format = DATE_FMT
        c_i.alignment = al_center
        if editable_date or tu_ngay is None:
            c_i.fill = fill_input

        # Col J (8): Đến ngày (link 1 ô duy nhất $D$4)
        c_j = ws.cell(row_idx, 10, "=$D$4")
        c_j.font = row_font
        c_j.number_format = DATE_FMT
        c_j.alignment = al_center

        # Col K (9): Số ngày chậm nộp đến hết 30/06/2016 (0,05%/ngày)
        c_k = ws.cell(
            row_idx,
            11,
            f'=IF(OR(C{row_idx}<=0,I{row_idx}="",I{row_idx}>$D$5),0,MAX(0,MIN(J{row_idx},$D$5)-I{row_idx}+1))',
        )
        c_k.font = row_font
        c_k.number_format = NUM_FMT
        c_k.alignment = al_right

        # Col L (10): Số ngày chậm nộp từ 01/07/2016 trở đi (0,03%/ngày)
        c_l = ws.cell(
            row_idx,
            12,
            f'=IF(OR(C{row_idx}<=0,I{row_idx}="",J{row_idx}<=$D$5),0,MAX(0,J{row_idx}-MAX(I{row_idx},$D$5+1)+1))',
        )
        c_l.font = row_font
        c_l.number_format = NUM_FMT
        c_l.alignment = al_right

        # Col M (11): Tiền chậm nộp (đồng)
        c_m = ws.cell(
            row_idx,
            13,
            f"=IF(C{row_idx}<=0,0,ROUND(C{row_idx}*(K{row_idx}*$D$7+L{row_idx}*$D$6),0))",
        )
        c_m.font = row_font
        c_m.number_format = NUM_FMT
        c_m.alignment = al_right

        # Col N (12): Cộng tiền xử lý VPHC & chậm nộp
        c_n = ws.cell(row_idx, 14, f"=E{row_idx}+H{row_idx}+M{row_idx}")
        c_n.font = f_bold if is_group_header_style else f_reg
        c_n.number_format = NUM_FMT
        c_n.alignment = al_right

        # Col O (13): Tổng thuế truy thu & xử lý
        c_o = ws.cell(row_idx, 15, f"=C{row_idx}+N{row_idx}")
        c_o.font = f_bold if is_group_header_style else f_reg
        c_o.number_format = NUM_FMT
        c_o.alignment = al_right

        # Col P (14): Tình trạng thời hiệu
        c_p = ws.cell(
            row_idx,
            16,
            f'=IF(I{row_idx}="","",IF(EDATE(I{row_idx},$H$7*12)<$D$4,"Quá hạn truy thu (>10 năm)",IF(EDATE(I{row_idx},$H$6*12)<$D$4,"Quá thời hiệu xử phạt (>5 năm)","Trong thời hiệu")))',
        )
        c_p.font = f_italic
        c_p.alignment = al_center

        for cc in range(1, 17):
            cell = ws.cell(row_idx, cc)
            cell.border = border_thin
            if is_group_header_style and cell.fill.fill_type is None:
                cell.fill = fill_grp

    # ===== WRITE YEAR BLOCKS =====
    sum_cols = [3, 4, 5, 6, 8, 13, 14, 15]
    curr_row = 13

    year_total_rows: List[int] = []
    vat_rows: List[int] = []
    cit_rows: List[int] = []
    pit_rows: List[int] = []
    nt_vat_rows: List[int] = []
    nt_cit_rows: List[int] = []
    lpmb_rows: List[int] = []

    for y_idx, yr in enumerate(target_years, start=1):
        r_yr = curr_row
        year_total_rows.append(r_yr)
        ws.row_dimensions[r_yr].height = 21
        ws.cell(r_yr, 1, _to_roman(y_idx)).alignment = al_center
        ws.cell(r_yr, 2, f"Năm {yr}").alignment = al_left
        for cc in range(1, 17):
            c = ws.cell(r_yr, cc)
            c.font = f_yr
            c.fill = fill_yr
            c.border = border_thin
            if cc in sum_cols:
                c.number_format = NUM_FMT
                c.alignment = al_right
        curr_row += 1

        # 1. Thuế GTGT
        r_vat = curr_row
        vat_rows.append(r_vat)
        ws.row_dimensions[r_vat].height = 20
        ws.cell(r_vat, 1, 1).alignment = al_center
        ws.cell(r_vat, 2, "Thuế GTGT").alignment = al_left
        for cc in range(1, 17):
            c = ws.cell(r_vat, cc)
            c.font = f_grp
            c.fill = fill_grp
            c.border = border_thin
            if cc in sum_cols:
                c.number_format = NUM_FMT
                c.alignment = al_right
        curr_row += 1

        vat_start = curr_row
        if freq == "quarter":
            for q in range(1, 5):
                _write_detail_row(
                    curr_row,
                    "",
                    f"   Quý {q}/{yr}",
                    adjust_start_date_for_weekend(compute_vat_quarterly_start_date(yr, q)),
                    tax_formula_or_val=0,
                    is_input_tax=True,
                )
                curr_row += 1
        else:
            for m in range(1, 13):
                _write_detail_row(
                    curr_row,
                    "",
                    f"   Tháng {m}",
                    adjust_start_date_for_weekend(compute_vat_monthly_start_date(yr, m)),
                    tax_formula_or_val=0,
                    is_input_tax=True,
                )
                curr_row += 1
        vat_end = curr_row - 1

        for cc in sum_cols:
            col_l = get_column_letter(cc)
            ws.cell(r_vat, cc, f"=SUM({col_l}{vat_start}:{col_l}{vat_end})")

        # 2. Thuế TNDN
        r_cit = curr_row
        cit_rows.append(r_cit)
        yr_sheet = str(yr)
        has_yr_sheet = yr_sheet in wb.sheetnames
        cit_val: object = f"=IF('{yr_sheet}'!E76>0,'{yr_sheet}'!E76,0)" if has_yr_sheet else 0
        cit_start = adjust_start_date_for_weekend(compute_cit_start_date(yr, fiscal_ends.get(yr)))
        _write_detail_row(
            r_cit,
            2,
            "Thuế TNDN" + (f" (NTC kết thúc {fiscal_ends[yr].strftime('%d/%m/%Y')})" if yr in fiscal_ends else ""),
            cit_start,
            tax_formula_or_val=cit_val,
            is_input_tax=not has_yr_sheet,
            is_group_header_style=True,
        )
        if has_yr_sheet:
            ws.cell(r_cit, 3).comment = Comment(
                f"Tự động liên kết Chênh lệch kiểm tra TNDN tại '{yr_sheet}'!E76 (chỉ lấy khi > 0). Có thể ghi đè trực tiếp nếu cần.",
                "Tinh_TienChamNop",
            )
        curr_row += 1

        # 3. Thuế TNCN
        r_pit = curr_row
        pit_rows.append(r_pit)
        pit_val: object = f"=IF('{yr_sheet}'!E90>0,'{yr_sheet}'!E90,0)" if has_yr_sheet else 0
        _write_detail_row(
            r_pit,
            3,
            "Thuế TNCN",
            adjust_start_date_for_weekend(compute_pit_start_date(yr)),
            tax_formula_or_val=pit_val,
            is_input_tax=not has_yr_sheet,
            is_group_header_style=True,
        )
        if has_yr_sheet:
            ws.cell(r_pit, 3).comment = Comment(
                f"Tự động liên kết Chênh lệch kiểm tra TNCN tại '{yr_sheet}'!E90 (chỉ lấy khi > 0). Có thể ghi đè trực tiếp nếu cần.",
                "Tinh_TienChamNop",
            )
        curr_row += 1

        # 4. Thuế Nhà thầu
        r_nt = curr_row
        ws.row_dimensions[r_nt].height = 20
        ws.cell(r_nt, 1, 4).alignment = al_center
        ws.cell(r_nt, 2, "Thuế Nhà thầu").alignment = al_left
        for cc in range(1, 17):
            c = ws.cell(r_nt, cc)
            c.font = f_grp
            c.fill = fill_grp
            c.border = border_thin
            if cc in sum_cols:
                c.number_format = NUM_FMT
                c.alignment = al_right
        curr_row += 1

        r_nt_g = curr_row
        nt_vat_rows.append(r_nt_g)
        _write_detail_row(r_nt_g, "", "   + Thuế GTGT Nhà thầu", None, 0, is_input_tax=True, editable_date=True)
        curr_row += 1

        r_nt_t = curr_row
        nt_cit_rows.append(r_nt_t)
        _write_detail_row(r_nt_t, "", "   + Thuế TNDN Nhà thầu", None, 0, is_input_tax=True, editable_date=True)
        curr_row += 1

        for cc in sum_cols:
            col_l = get_column_letter(cc)
            ws.cell(r_nt, cc, f"={col_l}{r_nt_g}+{col_l}{r_nt_t}")

        # 5. Lệ phí môn bài
        r_mb = curr_row
        lpmb_rows.append(r_mb)
        _write_detail_row(
            r_mb,
            5,
            "Lệ phí môn bài",
            adjust_start_date_for_weekend(compute_license_fee_start_date(yr)),
            tax_formula_or_val=0,
            is_input_tax=True,
            is_group_header_style=True,
        )
        curr_row += 1

        # Fill Year Subtotal row
        for cc in sum_cols:
            col_l = get_column_letter(cc)
            ws.cell(r_yr, cc, f"={col_l}{r_vat}+{col_l}{r_cit}+{col_l}{r_pit}+{col_l}{r_nt}+{col_l}{r_mb}")

    # ===== PHẠT VPHC VỀ THỦ TỤC THUẾ (MỤC RIÊNG) =====
    r_vphc_hdr = curr_row
    ws.row_dimensions[r_vphc_hdr].height = 20
    ws.cell(r_vphc_hdr, 1, _to_roman(len(target_years) + 1)).alignment = al_center
    ws.cell(r_vphc_hdr, 2, "Phạt VPHC về thủ tục thuế (NĐ 125/2020/NĐ-CP)").alignment = al_left
    for cc in range(1, 17):
        c = ws.cell(r_vphc_hdr, cc)
        c.font = f_yr
        c.fill = fill_yr
        c.border = border_thin
        if cc in sum_cols:
            c.number_format = NUM_FMT
            c.alignment = al_right
    curr_row += 1

    vphc_items = [
        ("1", "Phạt VPHC về thủ tục thuế (Điều 13, 14 NĐ 125/2020)"),
        ("2", "Phạt VPHC thiếu tờ khai / hồ sơ giao dịch liên kết (nếu có)"),
    ]
    vphc_rows: List[int] = []
    for sub_stt, sub_lbl in vphc_items:
        r_v = curr_row
        vphc_rows.append(r_v)
        ws.row_dimensions[r_v].height = 19
        ws.cell(r_v, 1, sub_stt).alignment = al_center
        ws.cell(r_v, 2, sub_lbl).alignment = al_left
        ws.cell(r_v, 3, 0).number_format = NUM_FMT
        ws.cell(r_v, 14, 0).number_format = NUM_FMT
        ws.cell(r_v, 14).fill = fill_input
        ws.cell(r_v, 14).comment = Comment(
            "Nhập số tiền phạt VPHC về thủ tục thuế vào ô này (tách riêng, không cộng lẫn vào cột Phạt 20%).",
            "Tinh_TienChamNop",
        )
        ws.cell(r_v, 15, f"=C{r_v}+N{r_v}").number_format = NUM_FMT
        ws.cell(r_v, 16, "Thời hiệu 2 năm").alignment = al_center
        for cc in range(1, 17):
            cell = ws.cell(r_v, cc)
            cell.font = f_italic if cc == 16 else f_reg
            cell.border = border_thin
            if cc in sum_cols:
                cell.alignment = al_right
        curr_row += 1

    for cc in sum_cols:
        col_l = get_column_letter(cc)
        if cc in (14, 15):
            ws.cell(r_vphc_hdr, cc, f"=SUM({col_l}{vphc_rows[0]}:{col_l}{vphc_rows[-1]})")
        else:
            ws.cell(r_vphc_hdr, cc, 0)

    # ===== TỔNG CỘNG TOÀN BẢNG =====
    r_grand = curr_row
    ws.row_dimensions[r_grand].height = 24
    ws.cell(r_grand, 2, "TỔNG CỘNG").alignment = al_center
    all_top_rows = year_total_rows + [r_vphc_hdr]
    for cc in range(1, 17):
        c = ws.cell(r_grand, cc)
        c.font = f_total
        c.fill = fill_total
        c.border = border_hdr
        if cc in sum_cols:
            col_l = get_column_letter(cc)
            c.value = "=" + "+".join(f"{col_l}{rr}" for rr in all_top_rows)
            c.number_format = NUM_FMT
            c.alignment = al_right
    curr_row += 2

    # ===== BẢNG TỔNG HỢP THEO SẮC THUẾ =====
    r_sum_title = curr_row
    ws.merge_cells(start_row=r_sum_title, start_column=1, end_row=r_sum_title, end_column=8)
    ws.cell(r_sum_title, 1, "BẢNG TỔNG HỢP SỐ THUẾ TRUY THU, TIỀN PHẠT VÀ TIỀN CHẬM NỘP THEO SẮC THUẾ").font = f_sec
    ws.cell(r_sum_title, 1).alignment = al_left
    ws.row_dimensions[r_sum_title].height = 20
    curr_row += 1

    r_sum_hdr = curr_row
    ws.row_dimensions[r_sum_hdr].height = 24
    sum_headers = [
        (1, "STT"),
        (2, "Sắc thuế / Khoản xử lý"),
        (3, "Thuế truy thu"),
        (4, "Phạt khai sai 20%"),
        (5, "Phạt trốn thuế"),
        (6, "Phạt VPHC thủ tục"),
        (7, "Tiền chậm nộp"),
        (8, "Tổng cộng"),
    ]
    for cc, txt in sum_headers:
        c = ws.cell(r_sum_hdr, cc, txt)
        c.font = f_hdr
        c.fill = fill_sum_hdr
        c.alignment = al_center
        c.border = border_hdr
    curr_row += 1

    def _sum_ref(col_letter: str, rows_list: Sequence[int]) -> str:
        if not rows_list:
            return "0"
        return "=" + "+".join(f"{col_letter}{r}" for r in rows_list)

    summary_specs = [
        ("1", "Thuế GTGT", vat_rows, False),
        ("2", "Thuế TNDN", cit_rows, False),
        ("3", "Thuế TNCN", pit_rows, False),
        ("4", "Thuế Nhà thầu (Tổng)", None, True),
        ("", "   + Thuế GTGT Nhà thầu", nt_vat_rows, False),
        ("", "   + Thuế TNDN Nhà thầu", nt_cit_rows, False),
        ("5", "Lệ phí môn bài", lpmb_rows, False),
        ("6", "Phạt VPHC về thủ tục thuế", vphc_rows, False),
    ]

    sum_row_map: Dict[str, int] = {}
    for stt, lbl, rlist, is_nt_parent in summary_specs:
        rr = curr_row
        sum_row_map[lbl] = rr
        ws.row_dimensions[rr].height = 19
        ws.cell(rr, 1, stt).alignment = al_center
        ws.cell(rr, 2, lbl).alignment = al_left

        if is_nt_parent:
            # Fill after children are known
            pass
        elif lbl == "Phạt VPHC về thủ tục thuế":
            ws.cell(rr, 3, 0)
            ws.cell(rr, 4, 0)
            ws.cell(rr, 5, 0)
            ws.cell(rr, 6, _sum_ref("N", rlist or []))
            ws.cell(rr, 7, 0)
            ws.cell(rr, 8, f"=SUM(C{rr}:G{rr})")
        else:
            ws.cell(rr, 3, _sum_ref("C", rlist or []))
            ws.cell(rr, 4, _sum_ref("E", rlist or []))
            ws.cell(rr, 5, _sum_ref("H", rlist or []))
            ws.cell(rr, 6, 0)
            ws.cell(rr, 7, _sum_ref("M", rlist or []))
            ws.cell(rr, 8, f"=SUM(C{rr}:G{rr})")

        for cc in range(1, 9):
            c = ws.cell(rr, cc)
            c.font = f_bold if is_nt_parent else f_reg
            if is_nt_parent:
                c.fill = fill_grp
            c.border = border_thin
            if cc >= 3:
                c.number_format = NUM_FMT
                c.alignment = al_right
        curr_row += 1

    # Link Nhà thầu (Tổng) = GTGT NT + TNDN NT
    r_nt_sum = sum_row_map["Thuế Nhà thầu (Tổng)"]
    r_nt_g_sum = sum_row_map["   + Thuế GTGT Nhà thầu"]
    r_nt_t_sum = sum_row_map["   + Thuế TNDN Nhà thầu"]
    for cc in range(3, 9):
        col_l = get_column_letter(cc)
        ws.cell(r_nt_sum, cc, f"={col_l}{r_nt_g_sum}+{col_l}{r_nt_t_sum}")

    # TỔNG CỘNG Bảng Tổng Hợp (chỉ cộng các dòng chính, không cộng trùng 2 dòng con của Nhà thầu)
    r_sum_total = curr_row
    ws.row_dimensions[r_sum_total].height = 22
    ws.cell(r_sum_total, 2, "TỔNG CỘNG").alignment = al_center
    main_sum_rows = [
        sum_row_map["Thuế GTGT"],
        sum_row_map["Thuế TNDN"],
        sum_row_map["Thuế TNCN"],
        sum_row_map["Thuế Nhà thầu (Tổng)"],
        sum_row_map["Lệ phí môn bài"],
        sum_row_map["Phạt VPHC về thủ tục thuế"],
    ]
    for cc in range(1, 9):
        c = ws.cell(r_sum_total, cc)
        c.font = f_total
        c.fill = fill_sum_hdr
        c.border = border_hdr
        if cc >= 3:
            col_l = get_column_letter(cc)
            c.value = "=" + "+".join(f"{col_l}{rr}" for rr in main_sum_rows)
            c.number_format = NUM_FMT
            c.alignment = al_right
    curr_row += 1

    # Dòng kiểm tra khớp tổng (= 0)
    r_chk = curr_row
    ws.row_dimensions[r_chk].height = 19
    ws.merge_cells(start_row=r_chk, start_column=1, end_row=r_chk, end_column=7)
    ws.cell(r_chk, 1, "Kiểm tra khớp Tổng bảng tổng hợp vs Biểu chi tiết (yêu cầu = 0):").font = f_italic
    ws.cell(r_chk, 1).alignment = al_right
    c_chk = ws.cell(r_chk, 8, f"=H{r_sum_total}-O{r_grand}")
    c_chk.font = f_bold
    c_chk.number_format = NUM_FMT
    c_chk.alignment = al_right
    c_chk.border = border_thin

    # ===== KHỐI KÝ TÊN (ĐẶT BÊN PHẢI BẢNG TỔNG HỢP, CỘT J-P) =====
    ws.merge_cells(start_row=r_sum_hdr, start_column=10, end_row=r_sum_hdr, end_column=16)
    ws.cell(r_sum_hdr, 10, "....., ngày ..... tháng ..... năm 20...").font = f_italic
    ws.cell(r_sum_hdr, 10).alignment = al_center

    ws.merge_cells(start_row=r_sum_hdr + 1, start_column=10, end_row=r_sum_hdr + 1, end_column=11)
    ws.cell(r_sum_hdr + 1, 10, "ĐẠI DIỆN DOANH NGHIỆP").font = f_bold
    ws.cell(r_sum_hdr + 1, 10).alignment = al_center

    ws.merge_cells(start_row=r_sum_hdr + 1, start_column=12, end_row=r_sum_hdr + 1, end_column=13)
    ws.cell(r_sum_hdr + 1, 12, "KẾ TOÁN TRƯỞNG").font = f_bold
    ws.cell(r_sum_hdr + 1, 12).alignment = al_center

    ws.merge_cells(start_row=r_sum_hdr + 1, start_column=14, end_row=r_sum_hdr + 1, end_column=16)
    ws.cell(r_sum_hdr + 1, 14, "TRƯỞNG ĐOÀN KIỂM TRA").font = f_bold
    ws.cell(r_sum_hdr + 1, 14).alignment = al_center

    ws.merge_cells(start_row=r_sum_hdr + 2, start_column=10, end_row=r_sum_hdr + 2, end_column=11)
    ws.cell(r_sum_hdr + 2, 10, "(Ký, ghi rõ họ tên, đóng dấu)").font = f_italic
    ws.cell(r_sum_hdr + 2, 10).alignment = al_center
    ws.merge_cells(start_row=r_sum_hdr + 2, start_column=12, end_row=r_sum_hdr + 2, end_column=13)
    ws.cell(r_sum_hdr + 2, 12, "(Ký, ghi rõ họ tên)").font = f_italic
    ws.cell(r_sum_hdr + 2, 12).alignment = al_center
    ws.merge_cells(start_row=r_sum_hdr + 2, start_column=14, end_row=r_sum_hdr + 2, end_column=16)
    ws.cell(r_sum_hdr + 2, 14, "(Ký, ghi rõ họ tên)").font = f_italic
    ws.cell(r_sum_hdr + 2, 14).alignment = al_center

    curr_row += 2

    # ===== BẢNG TRA CỨU CĂN CỨ PHÁP LÝ THEO TỪNG THỜI KỲ =====
    r_law_title = curr_row
    ws.merge_cells(start_row=r_law_title, start_column=1, end_row=r_law_title, end_column=16)
    ws.cell(r_law_title, 1, "PHỤ LỤC TRA CỨU: QUY ĐỊNH PHÁP LUẬT VỀ TIỀN CHẬM NỘP, XỬ PHẠT VPHC & THỜI HIỆU QUA CÁC THỜI KỲ").font = f_sec
    ws.row_dimensions[r_law_title].height = 20
    curr_row += 1

    r_law_hdr = curr_row
    ws.row_dimensions[r_law_hdr].height = 22
    law_cols = [
        (1, 2, "Nội dung / Giai đoạn"),
        (3, 5, "Mức áp dụng / Quy định"),
        (6, 10, "Căn cứ pháp lý"),
        (11, 16, "Lưu ý áp dụng trên Biểu tính"),
    ]
    for c1, c2, txt in law_cols:
        ws.merge_cells(start_row=r_law_hdr, start_column=c1, end_row=r_law_hdr, end_column=c2)
        c = ws.cell(r_law_hdr, c1, txt)
        c.font = f_hdr
        c.fill = fill_law_hdr
        c.alignment = al_center
        for cc in range(c1, c2 + 1):
            ws.cell(r_law_hdr, cc).fill = fill_law_hdr
            ws.cell(r_law_hdr, cc).border = border_hdr
    curr_row += 1

    law_rows = [
        (
            "Chậm nộp trước 01/01/2015",
            "0,05%/ngày (≤90 ngày); 0,07%/ngày (>90 ngày)",
            "Luật 21/2012/QH13 sửa đổi Luật QLT",
            "Giai đoạn trước 01/01/2015 thường đã quá thời hạn truy thu 10 năm; trường hợp đặc biệt cần tách riêng.",
        ),
        (
            "Chậm nộp 01/01/2015 – 30/06/2016",
            "0,05%/ngày (thống nhất, bỏ bậc 0,07%)",
            "Khoản 3 Điều 3 Luật 71/2014/QH14",
            "Cột K tự động tách số ngày đến hết 30/06/2016 và nhân với tỷ lệ ô D7 (0,05%).",
        ),
        (
            "Chậm nộp từ 01/07/2016 – nay",
            "0,03%/ngày trên số tiền thuế chậm nộp",
            "Luật 106/2016/QH13 & Khoản 2 Điều 59 Luật QLT 38/2019/QH14",
            "Cột L tự động tính số ngày từ 01/07/2016 đến ô D4 và nhân với tỷ lệ ô D6 (0,03%).",
        ),
        (
            "Hạn nộp & Ngày tính chậm nộp",
            "GTGT tháng: ngày 21; GTGT quý: ngày 01 tháng 2 quý sau; QT TNDN/TNCN: 01/04 (hoặc ngày 01 tháng 4 sau NTC); LPMB: 31/01",
            "Điều 44 Luật QLT 38/2019/QH14; NĐ 139/2016 & NĐ 22/2020 (LPMB)",
            "Cột I tự động điền ngày kế tiếp hạn nộp theo đúng kỳ tháng/quý và năm tài chính lệch kỳ. Thuế Nhà thầu nhập tay theo lần phát sinh (+10 ngày).",
        ),
        (
            "Phạt khai sai & Phạt trốn thuế",
            "Khai sai thiếu thuế: 20%; Trốn thuế: 1,0 – 3,0 lần (mặc định 1,5 lần)",
            "Điều 16 & Điều 17 Nghị định 125/2020/NĐ-CP",
            "Nhập số thuế trốn vào cột F sẽ tự động trừ khỏi cơ sở phạt 20% (cột D = MAX(0, C - F)) để không phạt trùng.",
        ),
        (
            "Thời hiệu xử phạt & Thời hạn truy thu",
            "Thời hiệu xử phạt: 02 năm (thủ tục), 05 năm (khai sai/trốn thuế); Truy thu: 10 năm",
            "Khoản 2 & Khoản 6 Điều 8 Nghị định 125/2020/NĐ-CP",
            "Quá 05 năm: Cột E & H tự động về 0 và Cột P cảnh báo 'Quá thời hiệu xử phạt (>5 năm)', vẫn tính truy thu & tiền chậm nộp.",
        ),
    ]
    for idx_l, (col1, col2, col3, col4) in enumerate(law_rows):
        rr = curr_row
        ws.row_dimensions[rr].height = 28
        vals_l = [(1, 2, col1, f_bold), (3, 5, col2, f_reg), (6, 10, col3, f_reg), (11, 16, col4, f_note)]
        for c1, c2, txt, fnt in vals_l:
            ws.merge_cells(start_row=rr, start_column=c1, end_row=rr, end_column=c2)
            c = ws.cell(rr, c1, txt)
            c.font = fnt
            c.alignment = al_left
            for cc in range(c1, c2 + 1):
                ws.cell(rr, cc).border = border_thin
                if idx_l % 2 == 1:
                    ws.cell(rr, cc).fill = fill_law_alt
        curr_row += 1

    # ===== COLUMN WIDTHS & PAGE SETUP =====
    col_widths = {
        1: 7,   # A: STT
        2: 32,  # B: Kỳ kê khai / Sắc thuế
        3: 17,  # C: Thuế truy thu
        4: 16,  # D: Thuế thiếu tính phạt 20%
        5: 16,  # E: Tiền phạt 20%
        6: 16,  # F: Thuế trốn tính phạt
        7: 10,  # G: Hệ số
        8: 16,  # H: Tiền phạt trốn thuế
        9: 13,  # I: Từ ngày
        10: 13, # J: Đến ngày
        11: 12, # K: Số ngày 0.05%
        12: 12, # L: Số ngày 0.03%
        13: 17, # M: Tiền chậm nộp
        14: 18, # N: Cộng xử lý VPHC & CN
        15: 19, # O: Tổng truy thu & xử lý
        16: 26, # P: Tình trạng thời hiệu
    }
    for cc, w in col_widths.items():
        ws.column_dimensions[get_column_letter(cc)].width = w

    ws.freeze_panes = "C13"
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    return ws


def integrate_late_payment_into_workbook(
    excel_path: str,
    cutoff_date: Optional[datetime.date] = None,
    years: Optional[Sequence[int]] = None,
    output_path: Optional[str] = None,
) -> str:
    """Mở file Excel có sẵn, bổ sung/cập nhật sheet 'Tinh_TienChamNop' và lưu lại."""
    wb = openpyxl.load_workbook(excel_path)
    add_late_payment_sheet(wb, years=years, cutoff_date=cutoff_date)
    dest = output_path or excel_path
    wb.save(dest)
    return dest
