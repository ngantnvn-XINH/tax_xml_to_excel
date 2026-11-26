# -*- coding: utf-8 -*-
"""
Module tạo các sheet phân tích & kiểm soát rủi ro thanh tra thuế phụ trợ độc lập cho `tax_xml_to_excel`:
(KHÔNG làm thay đổi bất kỳ cấu trúc, cột hay công thức nguyên bản nào của các sheet tờ khai & đối chiếu gốc.)

1. `KiemSoat_GTGT`:
   - Kiểm soát lỗi chuyển kỳ GTGT theo từng kỳ tháng/quý: so sánh [22] kỳ này với [43] kỳ liền trước
     -> tự đánh dấu 'X' ở cột 'Sai đầu kỳ' và tính đúng số tiền 'Chênh lệch đầu kỳ'.
   - Cảnh báo phân bổ thuế GTGT đầu vào: tính 'Tỷ lệ DT không chịu thuế [26] / Tổng DT [34]' và cảnh báo
     khi kỳ có doanh thu không chịu thuế nhưng thuế được khấu trừ [25] = thuế đầu vào [24].

2. `PhanTich_RuiRo`:
   - Bảng 1: Phân tích tỷ suất chi phí nhiều năm (%GV/DT, %CPBH/DT, %CPQLDN/DT) so với bình quân giai đoạn,
     tự tính chênh lệch tỷ lệ và ước tính biến động bất thường để khoanh vùng năm trọng tâm kiểm tra.
   - Bảng 2: Theo dõi các chỉ tiêu trọng yếu BCTC & TNCN theo năm (Phải thu KH 131, Hàng tồn kho 141,
     Phải trả NB 311, Tổng vay 320+338, Thu nhập khác 31, Quỹ lương TNCN).

3. `GDLK_30EBITDA`:
   - Bảng 1: Xác định yếu tố Giao dịch liên kết do vay vốn theo Điểm d Khoản 2 Điều 5 NĐ 132/2020/NĐ-CP
     (Vay >= 25% Vốn góp CSH [411] VÀ > 50% Nợ trung & dài hạn [330]), kèm dòng tách từng ngân hàng để nhập tay.
   - Bảng 2: Tính EBITDA & Khống chế chi phí lãi vay 30% EBITDA theo Khoản 3 Điều 16 NĐ 132/2020/NĐ-CP
     (tự động lấy Lợi nhuận thuần [30], Chi phí lãi vay [23], tạm lấy Lãi tiền gửi [21] kèm ghi chú rà soát,
     Khấu hao TSCĐ từ LCTT gián tiếp [02] hoặc biến động hao mòn lũy kế trên BCTC, và theo dõi hạn mức chuyển 5 năm).
"""

from __future__ import annotations

import re
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def _thin_border() -> Border:
    s = Side(style="thin", color="000000")
    return Border(left=s, right=s, top=s, bottom=s)


def _to_float(val: Any) -> float:
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _scan_bchdkd_coords(wb: openpyxl.Workbook) -> Tuple[Optional[str], Dict[int, Tuple[str, Dict[str, int]]]]:
    """Dò tọa độ cột từng năm và bản đồ mã chỉ tiêu -> dòng trên sheet 'BCHDKD'."""
    sheet_name = None
    for s in wb.sheetnames:
        if s.strip().lower() == "bchdkd" or "bchdkd" in s.strip().lower():
            sheet_name = s
            break
    if not sheet_name:
        return None, {}
    ws = wb[sheet_name]
    code_cols = [
        c for c in range(1, ws.max_column + 1)
        if str(ws.cell(2, c).value or "").strip() == "Mã chỉ tiêu"
    ]
    coords: Dict[int, Tuple[str, Dict[str, int]]] = {}
    for c in range(1, ws.max_column + 1):
        val_r1 = ws.cell(1, c).value
        if not val_r1:
            continue
        m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_r1))
        if not m:
            continue
        yr = int(m.group(0))
        left_cc = [cc for cc in code_cols if cc < c]
        cc = max(left_cc) if left_cc else 4
        cmap: Dict[str, int] = {}
        for r in range(4, ws.max_row + 1):
            code = str(ws.cell(r, cc).value or "").strip()
            if code and code not in cmap:
                cmap[code] = r
        coords[yr] = (get_column_letter(c), cmap)
    return sheet_name, coords


def _scan_bctc_coords(wb: openpyxl.Workbook) -> Tuple[Optional[str], Dict[int, Tuple[str, str, Dict[str, int]]]]:
    """
    Dò tọa độ cột (Số cuối năm, Số đầu năm) từng năm và bản đồ mã số -> dòng trên sheet 'BCTC' (CDKT).
    Trả về: (sheet_name, {year: (col_cuoi_letter, col_dau_letter, {code: row_idx})}).
    """
    sheet_name = None
    for s in wb.sheetnames:
        if s.strip().lower() == "bctc":
            sheet_name = s
            break
    if not sheet_name:
        return None, {}
    ws = wb[sheet_name]
    code_cols = [
        c for c in range(1, ws.max_column + 1)
        if str(ws.cell(4, c).value or "").strip().lower() in ("mã số", "ma so")
    ]
    coords: Dict[int, Tuple[str, str, Dict[str, int]]] = {}
    for c in range(1, ws.max_column + 1):
        val_r3 = ws.cell(3, c).value
        if val_r3 is None:
            continue
        m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_r3))
        if not m:
            continue
        yr = int(m.group(0))
        # Cột c là Số cuối năm, c+1 là Số đầu năm
        hdr_c = str(ws.cell(4, c).value or "").strip().lower()
        if "cuối" not in hdr_c and "cuoi" not in hdr_c:
            continue
        left_cc = [cc for cc in code_cols if cc < c]
        cc = max(left_cc) if left_cc else 2
        cmap: Dict[str, int] = {}
        for r in range(6, ws.max_row + 1):
            raw_code = ws.cell(r, cc).value
            if raw_code is None:
                continue
            code = str(raw_code).strip()
            if code.endswith(".0") and code[:-2].isdigit():
                code = code[:-2]
            if code and code not in cmap:
                cmap[code] = r
        coords[yr] = (get_column_letter(c), get_column_letter(c + 1), cmap)
    return sheet_name, coords


def _extract_lcttgt_depreciation_by_year(bctc_decls: Optional[Sequence[Any]]) -> Dict[int, float]:
    """
    Trích xuất trực tiếp chỉ tiêu Khấu hao TSCĐ và BĐSĐT (mã [02]) từ Báo cáo LCTT gián tiếp (PL_LCTTGT)
    trong danh sách tờ khai BCTC đã đọc từ XML (nếu DN nộp LCTT gián tiếp).
    """
    dep_by_yr: Dict[int, float] = {}
    if not bctc_decls:
        return dep_by_yr
    for d in bctc_decls:
        yr = getattr(d, "year", None)
        if not yr:
            continue
        appendices = getattr(d, "appendices", None) or {}
        for pl_name, rows in appendices.items():
            if "lcttgt" not in str(pl_name).lower():
                continue
            if not isinstance(rows, list):
                continue
            for row in rows:
                if not isinstance(row, dict):
                    continue
                code = str(row.get("ma_so") or row.get("maSo") or row.get("code") or "").strip()
                label = str(row.get("chi_tieu") or row.get("ten") or row.get("label") or "").lower()
                if code in ("02", "2") or "khấu hao" in label:
                    for vk in ("nam_nay", "so_cuoi_nam", "val_cur", "gia_tri", "value"):
                        if vk in row and row[vk] not in (None, "", 0, "0"):
                            val = abs(_to_float(row[vk]))
                            if val > 0:
                                dep_by_yr[yr] = val
                                break
                    if yr in dep_by_yr:
                        break
    return dep_by_yr


def write_sheet_kiem_soat_gtgt(
    wb: openpyxl.Workbook,
    existing_sheets: Set[str],
    sheet_name: str = "KiemSoat_GTGT",
) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
    """
    Sheet phụ trợ 1: Kiểm soát lỗi chuyển kỳ GTGT ([22] kỳ này vs [43] kỳ trước) & cảnh báo phân bổ
    thuế GTGT đầu vào cho doanh thu không chịu thuế ([26]).
    Đọc trực tiếp từ sheet 'So Thue GTGT Cuoi' (dòng thời gian liên tục), không sửa sheet gốc.
    """
    src_sheet = "So Thue GTGT Cuoi" if "So Thue GTGT Cuoi" in wb.sheetnames else None
    if not src_sheet:
        return None

    ws_src = wb[src_sheet]
    period_rows: List[int] = []
    for r in range(3, ws_src.max_row + 1):
        lbl_e = str(ws_src.cell(r, 5).value or "").strip().lower()
        ky_g = str(ws_src.cell(r, 7).value or "").strip()
        if not ky_g or "tổng" in lbl_e or "tong" in lbl_e or "tổng" in ky_g.lower():
            continue
        period_rows.append(r)

    if not period_rows:
        return None

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    existing_sheets.add(sheet_name)
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    b_thin = _thin_border()
    f_title = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    f_sub = Font(name="Calibri", size=9, italic=True, color="404040")
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    f_warn = Font(name="Calibri", size=9, bold=True, color="9C0006")
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_warn = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A1:Q1")
    ws["A1"] = "BẢNG KIỂM SOÁT CHUYỂN KỲ & PHÂN BỔ THUẾ GTGT ĐẦU VÀO THEO TỪNG KỲ KÊ KHAI"
    ws["A1"].font = f_title
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:Q2")
    ws["A2"] = (
        f"Nguồn liên kết tự động từ sheet '{src_sheet}'. "
        "Cột 'Sai đầu kỳ' tự đánh dấu X khi [22] kỳ này khác [43] kỳ liền trước; "
        "cột 'Cảnh báo phân bổ' tự báo khi kỳ có doanh thu không chịu thuế [26] > 0 nhưng khấu trừ 100% đầu vào ([25]=[24])."
    )
    ws["A2"].font = f_sub
    ws.row_dimensions[2].height = 18

    headers = [
        "STT",
        "Kỳ kê khai",
        "Lần BS",
        "Khấu trừ kỳ trước\nchuyển sang [22]",
        "Khấu trừ chuyển kỳ sau\ncủa kỳ liền trước [43]",
        "Sai đầu kỳ\n(X)",
        "Chênh lệch đầu kỳ\n([22] - [43] trước)",
        "Tổng DT bán ra\n[34]",
        "DT không chịu thuế\nGTGT [26]",
        "Tỷ lệ DT KCT /\nTổng DT (%)",
        "Thuế GTGT đầu vào\ntrong kỳ [24]",
        "Thuế GTGT được\nkhấu trừ [25]",
        "Cảnh báo phân bổ\nđầu vào",
        "Điều chỉnh giảm\nkỳ trước [37]",
        "Điều chỉnh tăng\nkỳ trước [38]",
        "Thuế GTGT phải nộp\ntrong kỳ [40]",
        "Thuế GTGT khấu trừ\nchuyển kỳ sau [43]",
    ]
    ws.row_dimensions[3].height = 34
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(3, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    widths = [6, 11, 8, 17, 17, 10, 17, 18, 17, 13, 17, 17, 18, 15, 15, 17, 17]
    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "D4"

    # Ánh xạ cột trên 'So Thue GTGT Cuoi' (GTGT_COLS_SPEC):
    # G(7)=Kỳ, H(8)=Lần, L(12)=[22], N(14)=[24], Q(17)=[25], R(18)=[26], AA(27)=[34],
    # AD(30)=[37], AE(31)=[38], AJ(36)=[40], AM(39)=[43]
    r_start = 4
    cur_r = r_start
    for idx, src_r in enumerate(period_rows, start=1):
        ref = f"'{src_sheet}'"
        ws.cell(cur_r, 1, idx).alignment = al_c
        ws.cell(cur_r, 2, f"={ref}!G{src_r}").alignment = al_c
        ws.cell(cur_r, 3, f"={ref}!H{src_r}").alignment = al_c
        ws.cell(cur_r, 4, f"={ref}!L{src_r}")
        if idx == 1:
            ws.cell(cur_r, 5, f"=D{cur_r}")
        else:
            ws.cell(cur_r, 5, f"=Q{cur_r - 1}")
        ws.cell(cur_r, 6, f'=IF(ABS(D{cur_r}-E{cur_r})>1,"X","")').alignment = al_c
        ws.cell(cur_r, 7, f"=D{cur_r}-E{cur_r}")
        ws.cell(cur_r, 8, f"={ref}!AA{src_r}")
        ws.cell(cur_r, 9, f"={ref}!R{src_r}")
        ws.cell(cur_r, 10, f"=IF(H{cur_r}>0,I{cur_r}/H{cur_r},0)")
        ws.cell(cur_r, 11, f"={ref}!N{src_r}")
        ws.cell(cur_r, 12, f"={ref}!Q{src_r}")
        ws.cell(cur_r, 13, f'=IF(AND(I{cur_r}>0,K{cur_r}>0,ABS(K{cur_r}-L{cur_r})<=1),"Cần KT phân bổ","")').alignment = al_c
        ws.cell(cur_r, 14, f"={ref}!AD{src_r}")
        ws.cell(cur_r, 15, f"={ref}!AE{src_r}")
        ws.cell(cur_r, 16, f"={ref}!AJ{src_r}")
        ws.cell(cur_r, 17, f"={ref}!AM{src_r}")

        # Kiểm tra nhanh giá trị thực để tô nổi bật ngay các ô bị sai đầu kỳ hoặc cần KT phân bổ
        v22 = _to_float(ws_src.cell(src_r, 12).value)
        v43_prev = v22 if idx == 1 else _to_float(ws_src.cell(period_rows[idx - 2], 39).value)
        has_dk_err = abs(v22 - v43_prev) > 1
        v26 = _to_float(ws_src.cell(src_r, 18).value)
        v24 = _to_float(ws_src.cell(src_r, 14).value)
        v25 = _to_float(ws_src.cell(src_r, 17).value)
        has_alloc_warn = (v26 > 0 and v24 > 0 and abs(v24 - v25) <= 1)

        for ci in range(1, 18):
            c = ws.cell(cur_r, ci)
            c.font = f_reg
            c.border = b_thin
            if ci in (4, 5, 7, 8, 9, 11, 12, 14, 15, 16, 17):
                c.number_format = "#,##0"
                c.alignment = al_r
            elif ci == 10:
                c.number_format = "0.00%"
                c.alignment = al_r
            if has_dk_err and ci in (6, 7):
                c.font = f_warn
                c.fill = fill_warn
            if has_alloc_warn and ci in (9, 10, 13):
                c.font = f_warn
                c.fill = fill_warn
        cur_r += 1

    r_end = cur_r - 1
    ws.cell(cur_r, 2, "TỔNG CỘNG").alignment = al_c
    ws.cell(cur_r, 6, f'=COUNTIF(F{r_start}:F{r_end},"X")&" kỳ"').alignment = al_c
    for ci in (7, 8, 9, 11, 12, 14, 15, 16):
        col_l = get_column_letter(ci)
        c = ws.cell(cur_r, ci, f"=SUM({col_l}{r_start}:{col_l}{r_end})")
        c.number_format = "#,##0"
        c.alignment = al_r
    c_pct = ws.cell(cur_r, 10, f"=IF(H{cur_r}>0,I{cur_r}/H{cur_r},0)")
    c_pct.number_format = "0.00%"
    c_pct.alignment = al_r
    ws.cell(cur_r, 13, f'=COUNTIF(M{r_start}:M{r_end},"Cần KT phân bổ")&" kỳ"').alignment = al_c
    for ci in range(1, 18):
        c = ws.cell(cur_r, ci)
        c.font = f_bold
        c.fill = fill_tot
        c.border = b_thin
    ws.auto_filter.ref = f"A3:Q{r_end}"
    return ws


def _kq_cell_ref(
    kq_sheet: Optional[str],
    kq_coords: Dict[int, Tuple[str, Dict[str, int]]],
    yr: int,
    logical_field: str,
) -> Any:
    """
    Ánh xạ chính xác mã chỉ tiêu trên sheet 'BCHDKD' (Phụ lục 03-1A/TNDN) cho cả 2 thông tư:
    - TT80/2021 (>=2021): [04]=DT bán hàng, [06]=Giảm trừ DT, [09]=Hàng bán bị trả lại,
      [10]=DT HĐTC, [11]=Lãi tiền gửi, [13]=Giá vốn, [14]=CP bán hàng, [15]=CP QLDN,
      [16]=CP tài chính, [17]=CP lãi vay, [18]=LN thuần HĐKD, [19]=Thu nhập khác, [20]=Chi phí khác.
    - TT156/2013 (<=2020): [01]=DT bán hàng, [03]=Giảm trừ DT, [06]=Hàng bán bị trả lại,
      [08]=DT HĐTC, [10]=Giá vốn, [11]=CP bán hàng, [12]=CP QLDN,
      [13]=CP tài chính, [14]=CP lãi vay, [15]=LN thuần HĐKD, [16]=Thu nhập khác, [17]=Chi phí khác.
    - Fallback mã BCTC KQKD chuẩn (01, 02, 10, 11, 21, 22, 23, 25, 26, 30, 31, 32) nếu sheet dùng mã trần.
    """
    if not kq_sheet or yr not in kq_coords:
        return 0
    col_l, cmap = kq_coords[yr]
    is_tt80 = ("[04]" in cmap and "[18]" in cmap) or ("[13]" in cmap and "[17]" in cmap and "[22]" in cmap)
    is_tt156 = ("[01]" in cmap and "[15]" in cmap and "[19]" in cmap)

    tt80_map = {
        "dt_ban_hang": ("[04]",),
        "giam_tru_dt": ("[06]",),
        "hang_ban_tra_lai": ("[09]", "[06]"),
        "dt_tai_chinh": ("[10]",),
        "lai_tien_gui": ("[11]", "[10]"),
        "gia_von": ("[13]",),
        "cp_ban_hang": ("[14]",),
        "cp_qldn": ("[15]",),
        "cp_tai_chinh": ("[16]",),
        "cp_lai_vay": ("[17]",),
        "ln_thuan": ("[18]",),
        "thu_nhap_khac": ("[19]",),
        "chi_phi_khac": ("[20]",),
    }
    tt156_map = {
        "dt_ban_hang": ("[01]",),
        "giam_tru_dt": ("[03]",),
        "hang_ban_tra_lai": ("[06]", "[03]"),
        "dt_tai_chinh": ("[08]",),
        "lai_tien_gui": ("[08]",),
        "gia_von": ("[10]",),
        "cp_ban_hang": ("[11]",),
        "cp_qldn": ("[12]",),
        "cp_tai_chinh": ("[13]",),
        "cp_lai_vay": ("[14]",),
        "ln_thuan": ("[15]",),
        "thu_nhap_khac": ("[16]",),
        "chi_phi_khac": ("[17]",),
    }
    bctc_fallback = {
        "dt_ban_hang": ("01",),
        "giam_tru_dt": ("02",),
        "hang_ban_tra_lai": ("02",),
        "dt_tai_chinh": ("21",),
        "lai_tien_gui": ("21",),
        "gia_von": ("11",),
        "cp_ban_hang": ("25",),
        "cp_qldn": ("26",),
        "cp_tai_chinh": ("22",),
        "cp_lai_vay": ("23",),
        "ln_thuan": ("30",),
        "thu_nhap_khac": ("31",),
        "chi_phi_khac": ("32",),
    }

    if is_tt80:
        cand_codes = tt80_map.get(logical_field, ())
    elif is_tt156:
        cand_codes = tt156_map.get(logical_field, ())
    else:
        cand_codes = bctc_fallback.get(logical_field, ())

    for cd in cand_codes:
        if cd in cmap:
            return f"='{kq_sheet}'!{col_l}{cmap[cd]}"
    return 0


def write_sheet_phan_tich_rui_ro(
    wb: openpyxl.Workbook,
    existing_sheets: Set[str],
    sheet_name: str = "PhanTich_RuiRo",
) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
    """
    Sheet phụ trợ 2: Phân tích rủi ro tỷ suất chi phí nhiều năm (%GV/DT, %CPBH/DT, %CPQLDN/DT)
    so với bình quân giai đoạn & theo dõi chỉ tiêu trọng yếu BCTC + TNCN.
    """
    kq_sheet, kq_coords = _scan_bchdkd_coords(wb)
    bctc_sheet, bctc_coords = _scan_bctc_coords(wb)
    years = sorted(set(kq_coords.keys()) | set(bctc_coords.keys()))
    if not years:
        return None

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    existing_sheets.add(sheet_name)
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    b_thin = _thin_border()
    f_title = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    f_sec = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A1:Q1")
    ws["A1"] = "PHÂN TÍCH RỦI RO TỶ SUẤT CHI PHÍ & CHỈ TIÊU TRỌNG YẾU QUA CÁC NĂM"
    ws["A1"].font = f_title
    ws.row_dimensions[1].height = 24

    # BẢNG 1: PHÂN TÍCH TỶ SUẤT CHI PHÍ (BCHDKD)
    ws["A3"] = "1. PHÂN TÍCH TỶ TRỌNG GIÁ VỐN, CHI PHÍ BÁN HÀNG, CHI PHÍ QLDN SO VỚI BÌNH QUÂN GIAI ĐOẠN"
    ws["A3"].font = f_sec

    hdr1 = [
        "STT",
        "Năm",
        "Doanh thu bán hàng\n[01]",
        "Giảm trừ DT /\nHàng bán bị trả lại [02]",
        "Doanh thu thuần\n[10]",
        "Giá vốn hàng bán\n[11]",
        "Chi phí tài chính\n[22]",
        "Trong đó:\nCP lãi vay [23]",
        "Chi phí bán hàng\n[25]",
        "Chi phí QLDN\n[26]",
        "%GV / DT",
        "%CPBH / DT",
        "%CPQLDN / DT",
        "Chênh lệch %GV\nso bình quân",
        "Ước tính biến động\ntăng Giá vốn",
        "Ước tính biến động\ntăng CP bán hàng",
        "Ước tính biến động\ntăng CP QLDN",
    ]
    ws.row_dimensions[4].height = 34
    for ci, h in enumerate(hdr1, start=1):
        c = ws.cell(4, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    widths = [6, 8, 17, 16, 17, 17, 16, 16, 16, 16, 11, 11, 12, 13, 17, 17, 17]
    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    r_start = 5
    r_tot = r_start + len(years)
    for idx, yr in enumerate(years, start=1):
        r = r_start + idx - 1
        ws.cell(r, 1, idx).alignment = al_c
        ws.cell(r, 2, yr).alignment = al_c
        ws.cell(r, 3, _kq_cell_ref(kq_sheet, kq_coords, yr, "dt_ban_hang"))
        ws.cell(r, 4, _kq_cell_ref(kq_sheet, kq_coords, yr, "giam_tru_dt"))
        ws.cell(r, 5, f"=C{r}-D{r}")
        ws.cell(r, 6, _kq_cell_ref(kq_sheet, kq_coords, yr, "gia_von"))
        ws.cell(r, 7, _kq_cell_ref(kq_sheet, kq_coords, yr, "cp_tai_chinh"))
        ws.cell(r, 8, _kq_cell_ref(kq_sheet, kq_coords, yr, "cp_lai_vay"))
        ws.cell(r, 9, _kq_cell_ref(kq_sheet, kq_coords, yr, "cp_ban_hang"))
        ws.cell(r, 10, _kq_cell_ref(kq_sheet, kq_coords, yr, "cp_qldn"))
        ws.cell(r, 11, f"=IF(E{r}>0,F{r}/E{r},0)")
        ws.cell(r, 12, f"=IF(E{r}>0,I{r}/E{r},0)")
        ws.cell(r, 13, f"=IF(E{r}>0,J{r}/E{r},0)")
        ws.cell(r, 14, f"=MAX(0,K{r}-$K${r_tot})")
        ws.cell(r, 15, f"=ROUND(N{r}*E{r},0)")
        ws.cell(r, 16, f"=ROUND(MAX(0,L{r}-$L${r_tot})*E{r},0)")
        ws.cell(r, 17, f"=ROUND(MAX(0,M{r}-$M${r_tot})*E{r},0)")

        for ci in range(1, 18):
            c = ws.cell(r, ci)
            c.font = f_reg
            c.border = b_thin
            if ci in (3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 17):
                c.number_format = "#,##0"
                c.alignment = al_r
            elif ci in (11, 12, 13, 14):
                c.number_format = "0.00%"
                c.alignment = al_r

    # Dòng TỔNG CỘNG / BÌNH QUÂN GIAI ĐOẠN
    ws.cell(r_tot, 2, "TỔNG / BQ").alignment = al_c
    for ci in (3, 4, 5, 6, 7, 8, 9, 10, 15, 16, 17):
        col_l = get_column_letter(ci)
        c = ws.cell(r_tot, ci, f"=SUM({col_l}{r_start}:{col_l}{r_tot - 1})")
        c.number_format = "#,##0"
        c.alignment = al_r
    ws.cell(r_tot, 11, f"=IF(E{r_tot}>0,F{r_tot}/E{r_tot},0)").number_format = "0.00%"
    ws.cell(r_tot, 12, f"=IF(E{r_tot}>0,I{r_tot}/E{r_tot},0)").number_format = "0.00%"
    ws.cell(r_tot, 13, f"=IF(E{r_tot}>0,J{r_tot}/E{r_tot},0)").number_format = "0.00%"
    ws.cell(r_tot, 14, f"=SUM(N{r_start}:N{r_tot - 1})").number_format = "0.00%"
    for ci in range(1, 18):
        c = ws.cell(r_tot, ci)
        c.font = f_bold
        c.fill = fill_tot
        c.border = b_thin
        if ci in (11, 12, 13, 14):
            c.alignment = al_r

    # BẢNG 2: CHỈ TIÊU TRỌNG YẾU BCTC & QUỸ LƯƠNG TNCN
    r_sec2 = r_tot + 2
    ws.cell(r_sec2, 1, "2. THEO DÕI CHỈ TIÊU TRỌNG YẾU BẢNG CÂN ĐỐI KẾ TOÁN & QUỸ LƯƠNG TNCN").font = f_sec
    r_h2 = r_sec2 + 1
    hdr2 = [
        "STT",
        "Năm",
        "Phải thu KH\nngắn hạn [131]",
        "Hàng tồn kho\n[141]",
        "Phải trả người bán\nngắn hạn [311]",
        "Tổng vay ngắn &\ndài hạn [320+338]",
        "Thu nhập khác\n[31]",
        "Chi phí khác\n[32]",
        "Quỹ lương chi trả\n(TNCN - Sheet năm)",
        "Vốn góp CSH\n[411]",
    ]
    ws.row_dimensions[r_h2].height = 30
    for ci, h in enumerate(hdr2, start=1):
        c = ws.cell(r_h2, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    def _bctc_ref(yr: int, code: str) -> Any:
        if bctc_sheet and yr in bctc_coords:
            col_c, _, cmap = bctc_coords[yr]
            if code in cmap:
                return f"='{bctc_sheet}'!{col_c}{cmap[code]}"
        return 0

    def _bctc_loan_ref(yr: int) -> Any:
        if bctc_sheet and yr in bctc_coords:
            col_c, _, cmap = bctc_coords[yr]
            parts = []
            for cd in ("320", "338"):
                if cd in cmap:
                    parts.append(f"'{bctc_sheet}'!{col_c}{cmap[cd]}")
            if parts:
                return "=" + "+".join(parts)
        return 0

    r2_start = r_h2 + 1
    for idx, yr in enumerate(years, start=1):
        r = r2_start + idx - 1
        ws.cell(r, 1, idx).alignment = al_c
        ws.cell(r, 2, yr).alignment = al_c
        ws.cell(r, 3, _bctc_ref(yr, "131"))
        ws.cell(r, 4, _bctc_ref(yr, "141"))
        ws.cell(r, 5, _bctc_ref(yr, "311"))
        ws.cell(r, 6, _bctc_loan_ref(yr))
        ws.cell(r, 7, _kq_cell_ref(kq_sheet, kq_coords, yr, "thu_nhap_khac"))
        ws.cell(r, 8, _kq_cell_ref(kq_sheet, kq_coords, yr, "chi_phi_khac"))
        ws.cell(r, 9, f"='{yr}'!C83" if str(yr) in wb.sheetnames else 0)
        ws.cell(r, 10, _bctc_ref(yr, "411"))
        for ci in range(1, 11):
            c = ws.cell(r, ci)
            c.font = f_reg
            c.border = b_thin
            if ci >= 3:
                c.number_format = "#,##0"
                c.alignment = al_r

    r2_tot = r2_start + len(years)
    ws.cell(r2_tot, 2, "TỔNG CỘNG").alignment = al_c
    for ci in range(3, 11):
        col_l = get_column_letter(ci)
        c = ws.cell(r2_tot, ci, f"=SUM({col_l}{r2_start}:{col_l}{r2_tot - 1})")
        c.number_format = "#,##0"
        c.alignment = al_r
    for ci in range(1, 11):
        c = ws.cell(r2_tot, ci)
        c.font = f_bold
        c.fill = fill_tot
        c.border = b_thin

    ws.freeze_panes = "C5"
    return ws


def write_sheet_gdlk_30ebitda(
    wb: openpyxl.Workbook,
    existing_sheets: Set[str],
    bctc_decls: Optional[Sequence[Any]] = None,
    sheet_name: str = "GDLK_30EBITDA",
) -> Optional[openpyxl.worksheet.worksheet.Worksheet]:
    """
    Sheet phụ trợ 3: Xác định yếu tố Giao dịch liên kết do vay vốn (Điểm d K2 Đ5 NĐ 132/2020/NĐ-CP)
    và Tính EBITDA / Khống chế chi phí lãi vay 30% EBITDA (Khoản 3 Điều 16 NĐ 132/2020/NĐ-CP).
    """
    kq_sheet, kq_coords = _scan_bchdkd_coords(wb)
    bctc_sheet, bctc_coords = _scan_bctc_coords(wb)
    years = sorted(set(kq_coords.keys()) | set(bctc_coords.keys()))
    if not years:
        return None

    lcttgt_dep = _extract_lcttgt_depreciation_by_year(bctc_decls)

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    existing_sheets.add(sheet_name)
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    b_thin = _thin_border()
    f_title = Font(name="Calibri", size=12, bold=True, color="1F4E78")
    f_sec = Font(name="Calibri", size=10, bold=True, color="1F4E78")
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    f_note = Font(name="Calibri", size=9, italic=True, color="404040")
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_hl = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")

    last_col_l = get_column_letter(max(4, 2 + len(years)))
    ws.merge_cells(f"A1:{last_col_l}1")
    ws["A1"] = "XÁC ĐỊNH YẾU TỐ GIAO DỊCH LIÊN KẾT (ĐIỂM d K2 Đ5 NĐ 132) & KHỐNG CHẾ CHI PHÍ LÃI VAY 30% EBITDA"
    ws["A1"].font = f_title
    ws.row_dimensions[1].height = 24

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 58
    for idx in range(len(years)):
        ws.column_dimensions[get_column_letter(3 + idx)].width = 18

    # =========================================================================
    # BẢNG 1: XÁC ĐỊNH YẾU TỐ GIAO DỊCH LIÊN KẾT DO VAY VỐN
    # =========================================================================
    ws["A3"] = "I. XÁC ĐỊNH YẾU TỐ GIAO DỊCH LIÊN KẾT DO VAY VỐN (Điểm d Khoản 2 Điều 5 NĐ 132/2020/NĐ-CP)"
    ws["A3"].font = f_sec

    ws.cell(4, 1, "STT").alignment = al_c
    ws.cell(4, 2, "Nội dung chỉ tiêu").alignment = al_c
    for idx, yr in enumerate(years):
        ws.cell(4, 3 + idx, f"Số dư 31/12/{yr}").alignment = al_c
    for ci in range(1, 3 + len(years)):
        c = ws.cell(4, ci)
        c.font = f_hdr
        c.fill = fill_hdr
        c.border = b_thin
    ws.row_dimensions[4].height = 26

    def _bctc_val_expr(yr: int, code: str) -> Any:
        if bctc_sheet and yr in bctc_coords:
            col_c, _, cmap = bctc_coords[yr]
            if code in cmap:
                return f"='{bctc_sheet}'!{col_c}{cmap[code]}"
        return 0

    def _bctc_total_loan_expr(yr: int) -> Any:
        if bctc_sheet and yr in bctc_coords:
            col_c, _, cmap = bctc_coords[yr]
            parts = [f"'{bctc_sheet}'!{col_c}{cmap[cd]}" for cd in ("320", "338") if cd in cmap]
            if parts:
                return "=" + "+".join(parts)
        return 0

    b1_rows = [
        ("1", "Tổng Vay và nợ thuê tài chính ([320] + [338] trên BCTC)", "loan_total", "#,##0", True, False),
        ("", "   - Ngân hàng / Tổ chức cho vay chính (nhập tách từ Thuyết minh BCTC; mặc định = Tổng vay)", "bank1", "#,##0", False, True),
        ("", "   - Ngân hàng / Tổ chức cho vay khác (nhập nếu có)", "bank2", "#,##0", False, True),
        ("2", "Vốn góp của chủ sở hữu ([411] trên BCTC)", "equity_411", "#,##0", True, False),
        ("", "   - Tham chiếu: Tổng Vốn chủ sở hữu ([410] trên BCTC)", "equity_410", "#,##0", False, False),
        ("3", "Tổng nợ trung và dài hạn ([330] trên BCTC)", "lt_debt_330", "#,##0", True, False),
        ("4", "Tỷ lệ Vay NH chính / Vốn góp CSH (Ngưỡng >= 25%)", "ratio_eq", "0.00%", True, False),
        ("5", "Tỷ lệ Vay NH chính / Nợ trung và dài hạn (Ngưỡng > 50%)", "ratio_lt", "0.00%", True, False),
        ("6", "Kết luận yếu tố GDLK (Vay >= 25% Vốn góp CSH VÀ > 50% Nợ trung & dài hạn)", "gdlk_flag", None, True, False),
    ]

    r_loan = 5
    r_bank1 = 6
    r_eq411 = 8
    r_lt330 = 10
    r_req = 11
    r_rlt = 12
    for offset, (stt_s, label, kind, nfmt, is_b, is_inp) in enumerate(b1_rows):
        r = 5 + offset
        ws.cell(r, 1, stt_s).alignment = al_c
        ws.cell(r, 2, label).alignment = al_l
        for idx, yr in enumerate(years):
            col_i = 3 + idx
            col_l = get_column_letter(col_i)
            if kind == "loan_total":
                val = _bctc_total_loan_expr(yr)
            elif kind == "bank1":
                val = f"={col_l}{r_loan}"
            elif kind == "bank2":
                val = 0
            elif kind == "equity_411":
                val = _bctc_val_expr(yr, "411")
            elif kind == "equity_410":
                val = _bctc_val_expr(yr, "410")
            elif kind == "lt_debt_330":
                val = _bctc_val_expr(yr, "330")
            elif kind == "ratio_eq":
                val = f"=IF({col_l}{r_eq411}>0,{col_l}{r_bank1}/{col_l}{r_eq411},0)"
            elif kind == "ratio_lt":
                val = f"=IF({col_l}{r_lt330}>0,{col_l}{r_bank1}/{col_l}{r_lt330},0)"
            else:
                val = f'=IF(AND({col_l}{r_req}>=0.25,{col_l}{r_rlt}>0.5),"CÓ YẾU TỐ GDLK","Không phát sinh")'
            c = ws.cell(r, col_i, val)
            if nfmt:
                c.number_format = nfmt
                c.alignment = al_r
            else:
                c.alignment = al_c

        for ci in range(1, 3 + len(years)):
            c = ws.cell(r, ci)
            c.font = f_bold if is_b else f_reg
            c.border = b_thin
            if is_inp and ci >= 3:
                c.fill = fill_input
            elif kind == "gdlk_flag":
                c.fill = fill_hl

    # =========================================================================
    # BẢNG 2: TÍNH EBITDA & KHỐNG CHẾ CHI PHÍ LÃI VAY 30% EBITDA
    # =========================================================================
    r_sec2 = 15
    ws.cell(r_sec2, 1, "II. TÍNH EBITDA & CHI PHÍ LÃI VAY KHÔNG ĐƯỢC TRỪ (KHỐNG CHẾ 30% EBITDA - K3 Đ16 NĐ 132/2020/NĐ-CP)").font = f_sec

    r_h2 = r_sec2 + 1
    ws.cell(r_h2, 1, "STT").alignment = al_c
    ws.cell(r_h2, 2, "Chỉ tiêu").alignment = al_c
    for idx, yr in enumerate(years):
        ws.cell(r_h2, 3 + idx, f"Năm {yr}").alignment = al_c
    for ci in range(1, 3 + len(years)):
        c = ws.cell(r_h2, ci)
        c.font = f_hdr
        c.fill = fill_hdr
        c.border = b_thin
    ws.row_dimensions[r_h2].height = 26

    def _kq_val_expr(yr: int, code: str) -> Any:
        if kq_sheet and yr in kq_coords:
            col_l, cmap = kq_coords[yr]
            if code in cmap:
                return f"='{kq_sheet}'!{col_l}{cmap[code]}"
        return 0

    def _dep_val_expr(yr: int) -> Tuple[Any, str]:
        if yr in lcttgt_dep:
            return round(lcttgt_dep[yr]), "Lấy trực tiếp từ LCTT gián tiếp mã [02] trong XML BCTC."
        if bctc_sheet and yr in bctc_coords:
            col_c, col_d, cmap = bctc_coords[yr]
            parts = []
            for cd in ("223", "226", "229", "232"):
                if cd in cmap:
                    rr = cmap[cd]
                    parts.append(f"ABS('{bctc_sheet}'!{col_c}{rr}-'{bctc_sheet}'!{col_d}{rr})")
            if parts:
                return "=" + "+".join(parts), (
                    "Ước tính từ biến động hao mòn lũy kế (223, 226, 229, 232) trên BCTC "
                    "(DN nộp LCTT trực tiếp không có mã [02]; lưu ý kiểm tra nếu năm có thanh lý TSCĐ)."
                )
        return 0, "Chưa có dữ liệu khấu hao trong XML BCTC - vui lòng nhập từ Thuyết minh BCTC."

    r_ln = r_h2 + 1        # (5) Lợi nhuận thuần [30]
    r_lv = r_h2 + 2        # (1) Chi phí lãi vay [23]
    r_ltg = r_h2 + 3       # (2) Lãi tiền gửi, lãi cho vay (tạm lấy [21])
    r_lv_net = r_h2 + 4    # (3) Chi phí lãi vay thuần = (1) - (2)
    r_kh = r_h2 + 5        # (4) Khấu hao TSCĐ
    r_ebitda = r_h2 + 6    # (6) EBITDA = (5) + (3) + (4)
    r_cap30 = r_h2 + 7     # (7) 30% EBITDA được trừ
    r_dis = r_h2 + 8       # (8) Chi phí lãi vay bị loại (không được trừ)
    r_carry = r_h2 + 9     # (9) Hạn mức 30% EBITDA còn dư để nhận chuyển lãi vay 5 năm

    b2_spec = [
        (r_ln, "(5)", "Lợi nhuận thuần từ hoạt động SXKD ([30] BCHDKD)", False, False),
        (r_lv, "(1)", "Chi phí lãi vay trong kỳ ([23] BCHDKD)", False, False),
        (r_ltg, "(2)", "Lãi tiền gửi, lãi cho vay (tạm lấy [21] BCHDKD - rà lại Thuyết minh BCTC nếu có cổ tức/tỷ giá)", False, True),
        (r_lv_net, "(3)", "Chi phí lãi vay thuần (= (1) - (2))", True, False),
        (r_kh, "(4)", "Chi phí khấu hao TSCĐ trong kỳ (ưu tiên LCTT gián tiếp [02], fallback biến động hao mòn BCTC)", False, True),
        (r_ebitda, "(6)", "EBITDA (= (5) + (3) + (4))", True, False),
        (r_cap30, "(7)", "Chi phí lãi vay được trừ tối đa (= 30% x EBITDA)", True, False),
        (r_dis, "(8)", "Chi phí lãi vay không được trừ trong kỳ (= MAX(0, (3) - (7)))", True, False),
        (r_carry, "(9)", "Hạn mức 30% EBITDA còn dư để nhận chuyển chi phí lãi vay từ kỳ trước (tối đa 5 năm từ 2019)", False, False),
    ]

    for rr, stt_s, lbl, is_b, is_inp in b2_spec:
        ws.cell(rr, 1, stt_s).alignment = al_c
        ws.cell(rr, 2, lbl).alignment = al_l
        for idx, yr in enumerate(years):
            col_i = 3 + idx
            col_l = get_column_letter(col_i)
            cmt_txt = None
            if rr == r_ln:
                val = _kq_cell_ref(kq_sheet, kq_coords, yr, "ln_thuan")
            elif rr == r_lv:
                val = _kq_cell_ref(kq_sheet, kq_coords, yr, "cp_lai_vay")
            elif rr == r_ltg:
                val = _kq_cell_ref(kq_sheet, kq_coords, yr, "lai_tien_gui")
                cmt_txt = (
                    "Với TT80 (>=2021): tự động lấy mã [11] Doanh thu từ lãi tiền gửi trên BCHDKD. "
                    "Với TT156 (<=2020): tạm lấy mã [08] Doanh thu HĐTC -> vui lòng rà lại Thuyết minh BCTC/TK 515."
                )
            elif rr == r_lv_net:
                val = f"={col_l}{r_lv}-{col_l}{r_ltg}"
            elif rr == r_kh:
                val, cmt_txt = _dep_val_expr(yr)
            elif rr == r_ebitda:
                val = f"={col_l}{r_ln}+{col_l}{r_lv_net}+{col_l}{r_kh}"
            elif rr == r_cap30:
                val = f"=MAX(0,ROUND(0.3*{col_l}{r_ebitda},0))"
            elif rr == r_dis:
                val = f"=MAX(0,{col_l}{r_lv_net}-{col_l}{r_cap30})"
            else:
                val = f"=MAX(0,{col_l}{r_cap30}-MAX(0,{col_l}{r_lv_net}))"
            c = ws.cell(rr, col_i, val)
            c.number_format = "#,##0"
            c.alignment = al_r
            if cmt_txt:
                c.comment = Comment(cmt_txt, "tax_xml_to_excel")

        for ci in range(1, 3 + len(years)):
            c = ws.cell(rr, ci)
            c.font = f_bold if is_b else f_reg
            c.border = b_thin
            if is_inp and ci >= 3:
                c.fill = fill_input
            elif rr in (r_ebitda, r_cap30, r_dis):
                c.fill = fill_hl

    r_note = r_carry + 2
    ws.merge_cells(f"A{r_note}:{last_col_l}{r_note}")
    ws.cell(
        r_note,
        1,
        "Lưu ý pháp lý: (1) Ô nền vàng cần đối chiếu nhanh với Thuyết minh BCTC/Sổ chi tiết. "
        "(2) Năm 2017-2018 áp dụng quy định hồi tố NĐ 68/2020/NĐ-CP; từ năm 2019 trở đi (NĐ 68/2020 & NĐ 132/2020), "
        "chi phí lãi vay không được trừ được chuyển sang kỳ tính thuế tiếp theo tối đa 05 năm liên tục khi năm sau có (30% EBITDA > Lãi vay thuần).",
    ).font = f_note
    ws.row_dimensions[r_note].height = 28
    ws.freeze_panes = "C5"
    return ws


def write_audit_analytics_sheets(
    wb: openpyxl.Workbook,
    existing_sheets: Set[str],
    bctc_decls: Optional[Sequence[Any]] = None,
    selected_sheets: Optional[Set[str]] = None,
) -> List[str]:
    """Tạo các sheet phụ trợ kiểm soát & phân tích rủi ro theo danh sách selected_sheets (không đụng chạm sheet gốc)."""
    def _want(key: str) -> bool:
        return selected_sheets is None or key in selected_sheets

    created: List[str] = []
    if _want("kiem_soat_gtgt") and write_sheet_kiem_soat_gtgt(wb, existing_sheets) is not None:
        created.append("KiemSoat_GTGT")
    if _want("phan_tich_rui_ro") and write_sheet_phan_tich_rui_ro(wb, existing_sheets) is not None:
        created.append("PhanTich_RuiRo")
    if _want("gdlk_ebitda") and write_sheet_gdlk_30ebitda(wb, existing_sheets, bctc_decls=bctc_decls) is not None:
        created.append("GDLK_30EBITDA")
    return created
