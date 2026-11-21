# -*- coding: utf-8 -*-
"""
Module xuất dữ liệu tờ khai thuế và BCTC ra Excel theo ĐÚNG MẪU CHUẨN (Phiên bản v1.2):
- Đảm bảo không thiếu dữ liệu đầu vào, không lệch dòng/cột giữa Thông tư 156 (<=2020) và Thông tư 80 (>=2021).
- Tự động tách: Mỗi năm ra 1 file Excel riêng biệt hoặc file tổng hợp.
- File TNCN hàng năm gồm 4 sheet đúng chuẩn:
    1. 'TK': Tờ khai quyết toán thuế TNCN (Mẫu 05/QTT-TNCN) - Lấy tờ khai cuối cùng
    2. '05-1': Bảng kê chi tiết cá nhân cư trú có HĐLĐ (Ánh xạ chuẩn cả TT156 và TT80)
    3. '05-2': Bảng kê cá nhân không cư trú / vãng lai (Ánh xạ chuẩn cả TT156 và TT80)
    4. '05-3': Bảng kê thông tin người phụ thuộc (Ánh xạ chuẩn cả TT156 và TT80)
- File TNDN & BCTC:
    - Sheet 'KKTNDN': Tờ khai quyết toán TNDN (Mẫu 03/TNDN - Tách chuẩn danh mục TT156 và TT80)
    - Sheet 'BCHDKD': Báo cáo Kết quả hoạt động kinh doanh (Mẫu 03-1A/TNDN - Không lẫn chỉ tiêu tờ khai chính)
    - Sheet 'BCTC': Bảng Cân đối kế toán (Mẫu B01-DN - Tôn trọng tuyệt đối Số đầu năm trong XML)
- File GTGT:
    - Sheet 'So Thue GTGT TH': Bảng kê toàn bộ tờ khai GTGT theo dòng thời gian
    - Sheet 'So Thue GTGT Cuoi': Bảng kê tờ khai GTGT cuối cùng mỗi kỳ + Tổng cộng từng năm
- Sheet 'Chenh lech thue GTGT - TNDN': Đối chiếu doanh thu GTGT và TNDN tự động dò tọa độ dòng.
- Sheet 'To_Khai_Khac': Tổng hợp đầy đủ dữ liệu các tờ khai khác (Lệ phí môn bài 01/LPMB, Nhà thầu...).
"""

from __future__ import annotations

import datetime
import os
import re
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from catalogs_dump import CDKT_ROWS, KKTNDN_ROWS, KQKD_ROWS
from split_catalogs import (
    KKTNDN_156_ROWS,
    KKTNDN_80_ROWS,
    KQKD_156_ROWS,
    KQKD_80_ROWS,
)
from tax_parser import COMMON_INDICATOR_LABELS, TaxDeclaration
from tncn_catalogs_dump import (
    TNCN_MASTER_ROWS_TT80,
    TNCN_MASTER_ROWS_TT156,
    TNCN_TT80_ROWS,
    TNCN_TT156_ROWS,
)
from tncn_templates import (
    COLS_05_1,
    COLS_05_2,
    COLS_05_3,
    MERGES_05_1,
    MERGES_05_2,
    MERGES_05_3,
)

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Calibri", size=9)
BOLD_FONT = Font(name="Calibri", size=9, bold=True)
SECTION_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
GRAY_HEADER_FILL = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
ERROR_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)


def sanitize_sheet_name(name: str, existing: Set[str]) -> str:
    clean = re.sub(r"[\\/*?:\[\]]", "_", name).strip()
    if not clean:
        clean = "Sheet"
    base = clean[:28] if len(clean) > 31 else clean
    candidate = clean[:31]
    idx = 2
    while candidate in existing:
        suffix = f"_{idx}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        idx += 1
    existing.add(candidate)
    return candidate


def write_cell_formatted(cell: openpyxl.cell.Cell, val: Any, is_mst: bool = False) -> None:
    cell.font = BODY_FONT
    cell.border = THIN_BORDER

    if is_mst:
        cell.value = str(val) if val is not None else ""
        cell.number_format = "@"
        cell.alignment = Alignment(horizontal="left", vertical="center")
        return

    if isinstance(val, (datetime.date, datetime.datetime)):
        cell.value = val
        cell.number_format = "dd/mm/yyyy"
        cell.alignment = Alignment(horizontal="center", vertical="center")
        return

    if isinstance(val, (int, float)) and not isinstance(val, bool):
        cell.value = val
        if isinstance(val, float) and not val.is_integer():
            cell.number_format = "#,##0.##"
        else:
            cell.number_format = "#,##0"
        cell.alignment = Alignment(horizontal="right", vertical="center")
        return

    cell.value = val
    cell.alignment = Alignment(horizontal="left", vertical="center")


def get_indicator_value(
    d: TaxDeclaration,
    code: str,
    source_dict: Optional[Dict[str, Any]] = None,
) -> Any:
    """
    Hàm tra cứu chính xác giá trị chỉ tiêu từ tờ khai thuế XML:
    - Tìm code gốc (VD: A1, B1, C1, ct22, [01], [04]...)
    - Tìm biến thể: ctA1, ctB1, ct_B1, [A1], ct22, 22...
    - Hỗ trợ truyền source_dict riêng (VD: d.kqkd_indicators cho phụ lục 03-1A/TNDN)
      để không bị lẫn chỉ tiêu cùng tên với tờ khai chính.
    """
    if not code:
        return None
    target_map = source_dict if source_dict is not None else d.indicators
    clean = code.replace("[", "").replace("]", "").strip()
    if "=" in clean:
        clean = clean.split("=")[0].strip()

    candidates = [
        code,
        clean,
        clean.upper(),
        clean.lower(),
        f"ct{clean}",
        f"ct{clean.upper()}",
        f"ct{clean.lower()}",
        f"ct_{clean}",
        f"ct_{clean.upper()}",
        f"ct_{clean.lower()}",
        f"[{clean}]",
        f"ms{clean}",
        f"ms_{clean}",
    ]
    if clean.isdigit():
        int_code = str(int(clean))
        candidates.extend([int_code, f"ct{int_code}", f"[{int_code}]"])
        if len(clean) == 1:
            z_code = f"0{clean}"
            candidates.extend([z_code, f"ct{z_code}", f"[{z_code}]"])

    for k in candidates:
        if k in target_map and target_map[k] is not None:
            return target_map[k]

    cand_upper = {c.upper() for c in candidates}
    for k, v in target_map.items():
        if v is not None:
            short = k.split(".")[-1]
            if short in candidates or short.upper() in cand_upper:
                return v

    return None


def _is_tt156_decl(d: TaxDeclaration) -> bool:
    """Xác định tờ khai thuộc mẫu Thông tư 156/151/92 (các năm <= 2020 hoặc mã tờ khai cũ)."""
    if d.ma_tkhai in ("01", "03", "395"):
        return True
    if d.ma_tkhai in ("842", "892", "953"):
        return False
    if d.year and d.year <= 2020:
        return True
    return False


def deduplicate_gtgt_declarations(declarations: List[TaxDeclaration]) -> List[TaxDeclaration]:
    """
    Tự động loại bỏ các tờ khai GTGT bị trùng lặp do thư mục có nhiều file của cùng 1 tờ khai
    (ví dụ: file rút gọn '...25M08_BS_1.xml' và file chứa mã giao dịch eTax '...08-2025_01-GTGT_BS1_1_...xml').
    """
    grouped: Dict[Tuple[str, str, int], List[TaxDeclaration]] = {}
    for d in declarations:
        mst = (d.mst or "").strip()
        if d.tu_ngay:
            ky_str = d.tu_ngay.strftime("%y%m")
        else:
            ky_str = (d.ky_kkhai or "").strip()
        try:
            lan = int(d.so_lan) if d.so_lan is not None and str(d.so_lan).strip() != "" else 0
        except (ValueError, TypeError):
            lan = 0
        key = (mst, ky_str, lan)
        grouped.setdefault(key, []).append(d)

    deduped: List[TaxDeclaration] = []
    for key, dlist in grouped.items():
        if len(dlist) == 1:
            deduped.append(dlist[0])
        else:
            def sort_gtgt_dups(item: TaxDeclaration):
                nlap = item.ngay_lap or datetime.date.min
                valid_inds = sum(1 for v in item.indicators.values() if v not in (None, 0, "", "0"))
                app_count = sum(len(rows) for rows in item.appendices.values())
                return (nlap, valid_inds, app_count, len(item.file_name), item.file_name)

            dlist.sort(key=sort_gtgt_dups)
            deduped.append(dlist[-1])

    return deduped


def deduplicate_declarations(declarations: List[TaxDeclaration]) -> List[TaxDeclaration]:
    """
    Loại bỏ các file XML trùng lặp của cùng 1 tờ khai (cùng loại thuế, MST, kỳ kê khai, số lần bổ sung)
    trên mọi sắc thuế (GTGT, TNDN, TNCN, BCTC, Khác).
    """
    grouped: Dict[Tuple[str, str, str, int], List[TaxDeclaration]] = {}
    for d in declarations:
        mst = (d.mst or "").strip()
        if d.tax_type == "GTGT" and d.tu_ngay:
            ky_str = d.tu_ngay.strftime("%y%m")
        else:
            ky_str = str(d.year or d.ky_kkhai or "").strip()
        try:
            lan = int(d.so_lan) if d.so_lan is not None and str(d.so_lan).strip() != "" else 0
        except (ValueError, TypeError):
            lan = 0
        key = (d.tax_type or "", mst, ky_str, lan)
        grouped.setdefault(key, []).append(d)

    deduped: List[TaxDeclaration] = []
    for key, dlist in grouped.items():
        if len(dlist) == 1:
            deduped.append(dlist[0])
        else:
            def sort_dups(item: TaxDeclaration):
                nlap = item.ngay_lap or datetime.date.min
                valid_inds = sum(1 for v in item.indicators.values() if v not in (None, 0, "", "0"))
                app_count = sum(len(rows) for rows in item.appendices.values())
                return (nlap, valid_inds, app_count, len(item.file_name), item.file_name)

            dlist.sort(key=sort_dups)
            deduped.append(dlist[-1])

    return deduped


def filter_latest_per_year(declarations: List[TaxDeclaration]) -> List[TaxDeclaration]:
    """
    Với TNDN, TNCN, BCTC, Khác: Nhóm theo (MST, Năm).
    Mỗi năm chỉ giữ lại đúng 1 tờ khai nộp cuối cùng (so_lan cao nhất, ngay_lap mới nhất).
    Nếu trùng cả số lần và ngày lập (do file trùng lặp), ưu tiên file có dữ liệu đầy đủ nhất.
    """
    grouped: Dict[Tuple[str, int], List[TaxDeclaration]] = {}
    for d in declarations:
        y = d.year or 0
        k = (d.mst or "", y)
        grouped.setdefault(k, []).append(d)

    latest_list: List[TaxDeclaration] = []
    for k, decl_list in grouped.items():
        def sort_key(item: TaxDeclaration):
            try:
                lan = int(item.so_lan) if item.so_lan is not None and str(item.so_lan).strip() != "" else 0
            except (ValueError, TypeError):
                lan = 0
            nlap = item.ngay_lap or datetime.date.min
            valid_inds = sum(1 for v in item.indicators.values() if v not in (None, 0, "", "0"))
            app_count = sum(len(rows) for rows in item.appendices.values())
            return (lan, nlap, valid_inds, app_count, len(item.file_name), item.file_name)

        decl_list.sort(key=sort_key)
        latest = decl_list[-1]
        if len(decl_list) > 1:
            merged_apps = dict(latest.appendices)
            for prev in reversed(decl_list[:-1]):
                for app_name, app_rows in prev.appendices.items():
                    if "khbs" in app_name.lower():
                        continue
                    if not merged_apps.get(app_name) and app_rows:
                        merged_apps[app_name] = app_rows
                if not latest.kqkd_indicators and prev.kqkd_indicators:
                    latest.kqkd_indicators = dict(prev.kqkd_indicators)
            latest.appendices = merged_apps
        latest_list.append(latest)

    latest_list.sort(key=lambda d: (d.mst or "", d.year or 0, d.tu_ngay or datetime.date.min))
    return latest_list


# ==============================================================================
# 1. SHEET KKTNDN: Tờ khai Quyết toán TNDN (Mẫu 03/TNDN)
# Tách riêng TT156 (các năm <= 2020, maTKhai='03') và TT80 (các năm >= 2021, maTKhai='892')
# ==============================================================================
def write_sheet_kktndn(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
) -> None:
    sname = sanitize_sheet_name("KKTNDN", existing_sheets)
    ws = wb.create_sheet(title=sname)

    sorted_decls = sorted(declarations, key=lambda d: (d.year or 0, d.tu_ngay or datetime.date.min))
    decls_156 = [d for d in sorted_decls if _is_tt156_decl(d)]
    decls_80 = [d for d in sorted_decls if not _is_tt156_decl(d)]

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 18

    internal_lookup_only_codes = {"ten_loaiMien", "thueTNDN_tamnop_sxkd", "thueTNDN_tamnop_cnbds"}

    def write_kktndn_block(c_start: int, decls: List[TaxDeclaration], catalog: List[Tuple[str, str, str]]) -> int:
        from openpyxl.comments import Comment as _Comment
        c_end = c_start + 2 + len(decls)
        for idx, d in enumerate(decls):
            c_yr = ws.cell(1, c_start + 3 + idx, d.year or d.ky_kkhai)
            c_yr.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            c_yr.alignment = Alignment(horizontal="center", vertical="center")
            c_yr.border = THIN_BORDER
            nlap_str = d.ngay_lap.strftime("%d/%m/%Y") if d.ngay_lap else ""
            c_yr.comment = _Comment(
                f"Loại: {d.loai_to_khai}\nNgày lập: {nlap_str}\nFile: {d.file_name}",
                "TaxXML",
            )

        ws.cell(2, c_start, "STT")
        ws.cell(2, c_start + 1, "Chỉ tiêu")
        ws.cell(2, c_start + 2, "Mã chỉ tiêu")
        for idx in range(len(decls)):
            ws.cell(2, c_start + 3 + idx, "Số tiền")

        ws.cell(3, c_start, -1)
        ws.cell(3, c_start + 1, -2)
        ws.cell(3, c_start + 2, -3)
        for idx in range(len(decls)):
            ws.cell(3, c_start + 3 + idx, -4)

        for col in range(c_start, c_end + 1):
            cell2 = ws.cell(2, col)
            cell2.fill = GRAY_HEADER_FILL
            cell2.font = BOLD_FONT
            cell2.border = THIN_BORDER
            cell2.alignment = Alignment(horizontal="center", vertical="center")

            cell3 = ws.cell(3, col)
            cell3.font = BODY_FONT
            cell3.border = THIN_BORDER
            cell3.alignment = Alignment(horizontal="center", vertical="center")

        cur_row = 4
        for stt, desc, code in catalog:
            if not stt and not desc and not code:
                continue
            display_code = "" if code in internal_lookup_only_codes else code
            c_stt = ws.cell(cur_row, c_start, stt)
            c_desc = ws.cell(cur_row, c_start + 1, desc)
            c_code = ws.cell(cur_row, c_start + 2, display_code)

            is_major_letter = stt in ("A", "B", "C", "D", "E", "G", "H", "I") and not any(
                display_code.endswith(digit) for digit in "0123456789"
            )
            # Kiểm tra xem dòng mục lớn (như D, E, G, H, I) có số liệu chỉ tiêu thực tế trong XML không
            has_major_value = is_major_letter and (
                stt in ("E", "G", "I")
                or any(get_indicator_value(d, code) is not None for d in decls)
            )

            if is_major_letter:
                for cell in (c_stt, c_desc, c_code):
                    cell.font = BOLD_FONT
                    cell.fill = SECTION_FILL
                    cell.border = THIN_BORDER
                c_stt.alignment = Alignment(horizontal="center", vertical="center")
                c_code.alignment = Alignment(horizontal="center", vertical="center")
                for idx, d in enumerate(decls):
                    val = get_indicator_value(d, code) if (code and stt not in ("A", "B", "C")) else None
                    cell = ws.cell(cur_row, c_start + 3 + idx)
                    if has_major_value and stt not in ("A", "B", "C"):
                        write_cell_formatted(cell, val if val is not None else 0)
                        cell.font = BOLD_FONT
                    else:
                        cell.value = ""
                    cell.fill = SECTION_FILL
                    cell.border = THIN_BORDER
            else:
                c_stt.font = BODY_FONT
                c_desc.font = BODY_FONT
                c_code.font = BODY_FONT
                c_stt.border = THIN_BORDER
                c_desc.border = THIN_BORDER
                c_code.border = THIN_BORDER
                c_stt.alignment = Alignment(horizontal="center", vertical="center")
                c_code.alignment = Alignment(horizontal="center", vertical="center")

                for idx, d in enumerate(decls):
                    val = get_indicator_value(d, code) if code else None
                    cell = ws.cell(cur_row, c_start + 3 + idx)
                    if code == "ten_loaiMien":
                        write_cell_formatted(cell, val if val is not None else "")
                    elif not code and desc == "Trong đó:":
                        write_cell_formatted(cell, 0)
                    else:
                        write_cell_formatted(cell, val if val is not None else 0)
                    cell.border = THIN_BORDER

            cur_row += 1

        ws.column_dimensions[get_column_letter(c_start)].width = 7
        ws.column_dimensions[get_column_letter(c_start + 1)].width = 54
        ws.column_dimensions[get_column_letter(c_start + 2)].width = 12
        for idx in range(len(decls)):
            ws.column_dimensions[get_column_letter(c_start + 3 + idx)].width = 16

        return c_end

    if decls_156 and decls_80:
        c156_end = write_kktndn_block(2, decls_156, KKTNDN_156_ROWS)
        c_sep = c156_end + 1
        ws.column_dimensions[get_column_letter(c_sep)].width = 6
        write_kktndn_block(c_sep + 1, decls_80, KKTNDN_80_ROWS)
    elif decls_156:
        write_kktndn_block(2, decls_156, KKTNDN_156_ROWS)
    else:
        write_kktndn_block(2, decls_80 if decls_80 else sorted_decls, KKTNDN_80_ROWS)

    ws.column_dimensions["A"].width = 3
    ws.freeze_panes = "C4"


# ==============================================================================
# 2. SHEET BCHDKD: Báo cáo kết quả hoạt động kinh doanh (Phụ lục 03-1A/TNDN)
# Tách riêng TT156 (các năm <= 2020, mã [01]..[19]) và TT80 (các năm >= 2021, mã [04]..[22])
# Đọc trực tiếp từ kqkd_indicators (03-1A/TNDN), không bị nhiễu bởi chỉ tiêu tờ khai chính.
# ==============================================================================
def write_sheet_kqkd(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
) -> None:
    sname = sanitize_sheet_name("BCHDKD", existing_sheets)
    ws = wb.create_sheet(title=sname)

    sorted_decls = sorted(declarations, key=lambda d: (d.year or 0, d.tu_ngay or datetime.date.min))
    decls_156 = [d for d in sorted_decls if _is_tt156_decl(d)]
    decls_80 = [d for d in sorted_decls if not _is_tt156_decl(d)]

    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 20

    def write_kqkd_block(
        c_start: int,
        decls: List[TaxDeclaration],
        catalog: List[Tuple[str, str, str]],
        title_text: str,
    ) -> int:
        from openpyxl.comments import Comment as _Comment
        c_end = c_start + 2 + len(decls)
        for idx, d in enumerate(decls):
            c_yr = ws.cell(1, c_start + 3 + idx, d.year or d.ky_kkhai)
            c_yr.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            c_yr.alignment = Alignment(horizontal="center", vertical="center")
            c_yr.border = THIN_BORDER
            nlap_str = d.ngay_lap.strftime("%d/%m/%Y") if d.ngay_lap else ""
            c_yr.comment = _Comment(
                f"Loại: {d.loai_to_khai}\nNgày lập: {nlap_str}\nFile: {d.file_name}",
                "TaxXML",
            )

        ws.cell(2, c_start, "STT")
        ws.cell(2, c_start + 1, "Chỉ tiêu")
        ws.cell(2, c_start + 2, "Mã chỉ tiêu")
        for idx in range(len(decls)):
            ws.cell(2, c_start + 3 + idx, "Số tiền")

        ws.cell(3, c_start, -1)
        ws.cell(3, c_start + 1, -2)
        ws.cell(3, c_start + 2, -3)
        for idx in range(len(decls)):
            ws.cell(3, c_start + 3 + idx, -4)

        for col in range(c_start, c_end + 1):
            cell2 = ws.cell(2, col)
            cell2.fill = GRAY_HEADER_FILL
            cell2.font = BOLD_FONT
            cell2.border = THIN_BORDER
            cell2.alignment = Alignment(horizontal="center", vertical="center")

            cell3 = ws.cell(3, col)
            cell3.font = BODY_FONT
            cell3.border = THIN_BORDER
            cell3.alignment = Alignment(horizontal="center", vertical="center")

        cell_t = ws.cell(4, c_start + 1, title_text)
        cell_t.font = BOLD_FONT
        for col in range(c_start, c_end + 1):
            cell4 = ws.cell(4, col)
            cell4.fill = SECTION_FILL
            cell4.border = THIN_BORDER

        cur_row = 5
        for stt, desc, code in catalog:
            c_stt = ws.cell(cur_row, c_start, stt)
            c_desc = ws.cell(cur_row, c_start + 1, desc)
            c_code = ws.cell(cur_row, c_start + 2, code)

            c_stt.font = BODY_FONT
            c_desc.font = BODY_FONT
            c_code.font = BODY_FONT
            c_stt.border = THIN_BORDER
            c_desc.border = THIN_BORDER
            c_code.border = THIN_BORDER
            c_stt.alignment = Alignment(horizontal="center", vertical="center")
            c_code.alignment = Alignment(horizontal="center", vertical="center")

            for idx, d in enumerate(decls):
                # Ưu tiên tuyệt đối kqkd_indicators từ phụ lục 03-1A/TNDN; chỉ fallback sang d.indicators khi kqkd_indicators rỗng (như unit test giả lập)
                src = d.kqkd_indicators if d.kqkd_indicators else d.indicators
                val = get_indicator_value(d, code, source_dict=src) if code else None
                cell = ws.cell(cur_row, c_start + 3 + idx)
                write_cell_formatted(cell, val if val is not None else 0)
                cell.border = THIN_BORDER

            cur_row += 1

        ws.column_dimensions[get_column_letter(c_start)].width = 7
        ws.column_dimensions[get_column_letter(c_start + 1)].width = 54
        ws.column_dimensions[get_column_letter(c_start + 2)].width = 12
        for idx in range(len(decls)):
            ws.column_dimensions[get_column_letter(c_start + 3 + idx)].width = 16

        return c_end

    if decls_156 and decls_80:
        c156_end = write_kqkd_block(
            2, decls_156, KQKD_156_ROWS, "Kết quả kinh doanh ghi nhận theo báo cáo tài chính."
        )
        c_sep = c156_end + 1
        ws.column_dimensions[get_column_letter(c_sep)].width = 6
        write_kqkd_block(
            c_sep + 1, decls_80, KQKD_80_ROWS, "Kết quả kinh doanh ghi nhận theo báo cáo tài chính:"
        )
    elif decls_156:
        write_kqkd_block(2, decls_156, KQKD_156_ROWS, "Kết quả kinh doanh ghi nhận theo báo cáo tài chính.")
    else:
        write_kqkd_block(
            2,
            decls_80 if decls_80 else sorted_decls,
            KQKD_80_ROWS,
            "Kết quả kinh doanh ghi nhận theo báo cáo tài chính:",
        )

    ws.column_dimensions["A"].width = 3
    ws.freeze_panes = "C5"


# ==============================================================================
# 3. SHEET BCTC: Bảng cân đối kế toán (Mẫu B01-DN / B01a-DNN)
# Tách riêng 2 khối: Khối các năm <= 2020 bên trái, Khối các năm >= 2021 bên phải.
# Mỗi năm gồm 2 cột 'Số cuối năm' và 'Số đầu năm' đóng khung viền rõ nét.
# Tôn trọng tuyệt đối Số đầu năm trong XML khi file XML có khai khối <SoDauNam>
# (không tự ý ghi đè các chỉ tiêu bằng 0 do phân loại lại đầu năm).
# Chỉ kế thừa Số cuối năm trước khi file XML hoàn toàn không có dữ liệu <SoDauNam>.
# ==============================================================================
def write_sheet_bctc_cdkt(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
) -> None:
    sname = sanitize_sheet_name("BCTC", existing_sheets)
    ws = wb.create_sheet(title=sname)

    sorted_decls = sorted(declarations, key=lambda d: (d.year or 0, d.tu_ngay or datetime.date.min))
    decls_156 = [d for d in sorted_decls if d.year and d.year <= 2020]
    decls_80 = [d for d in sorted_decls if not (d.year and d.year <= 2020)]

    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 24
    ws.row_dimensions[5].height = 18

    def _decl_has_any_sodaunam(d: TaxDeclaration) -> bool:
        if d.has_bctc_sodaunam:
            return True
        for k, v in d.indicators.items():
            if v not in (None, 0, "0", ""):
                kl = k.lower()
                if "khonglientuc" not in kl and (
                    kl.endswith("_dau")
                    or kl.endswith("_sodaunam")
                    or any(t in kl for t in ("sodaunam", "sodauky", "dauky", "daunam"))
                ):
                    return True
        return False

    has_dau_cache: Dict[int, bool] = {id(d): _decl_has_any_sodaunam(d) for d in sorted_decls}

    def get_bctc_val_pair(d: TaxDeclaration, prev_d: Optional[TaxDeclaration], code_str: str) -> Tuple[Any, Any]:
        # 1. Tìm số cuối năm
        v_cuoi = get_indicator_value(d, f"{code_str}_cuoi")
        if v_cuoi is None:
            v_cuoi = get_indicator_value(d, f"ms{code_str}_cuoi")
        if v_cuoi is None:
            v_cuoi = get_indicator_value(d, f"ct{code_str}_cuoi")
        if v_cuoi is None:
            v_cuoi = get_indicator_value(d, f"{code_str}_socuoinam")
        if v_cuoi is None:
            for k, v in d.indicators.items():
                kl = k.lower()
                if any(x in kl for x in ("socuoinam", "socuoi", "cuoiky", "cuoinam", "namnay")):
                    if kl.endswith(f".ct{code_str}") or kl.endswith(f".ms{code_str}") or kl.endswith(f".{code_str}"):
                        if "khonglientuc" not in kl or v_cuoi is None:
                            if v not in (None, "", 0) or v_cuoi is None:
                                v_cuoi = v
        if v_cuoi is None:
            v_cuoi = get_indicator_value(d, code_str)
        if v_cuoi is None:
            v_cuoi = get_indicator_value(d, f"ct{code_str}")

        # 2. Tìm số đầu năm
        v_dau = get_indicator_value(d, f"{code_str}_dau")
        if v_dau is None:
            v_dau = get_indicator_value(d, f"ms{code_str}_dau")
        if v_dau is None:
            v_dau = get_indicator_value(d, f"ct{code_str}_dau")
        if v_dau is None:
            v_dau = get_indicator_value(d, f"{code_str}_sodaunam")
        if v_dau is None:
            for k, v in d.indicators.items():
                kl = k.lower()
                if any(x in kl for x in ("sodaunam", "sodau", "dauky", "daunam", "namtruoc", "kytruoc")):
                    if kl.endswith(f".ct{code_str}") or kl.endswith(f".ms{code_str}") or kl.endswith(f".{code_str}"):
                        if "khonglientuc" not in kl or v_dau is None:
                            if v not in (None, "", 0) or v_dau is None:
                                v_dau = v
                                if v not in (None, "", 0):
                                    break

        # 3. Chỉ kế thừa số cuối năm trước khi file XML năm hiện tại hoàn toàn không có dữ liệu <SoDauNam>
        if not has_dau_cache.get(id(d), False) and (v_dau is None or v_dau == 0) and prev_d is not None:
            prev_cuoi = get_indicator_value(prev_d, f"{code_str}_cuoi")
            if prev_cuoi is None:
                prev_cuoi = get_indicator_value(prev_d, f"ms{code_str}_cuoi")
            if prev_cuoi is None:
                prev_cuoi = get_indicator_value(prev_d, f"ct{code_str}_cuoi")
            if prev_cuoi is None:
                prev_cuoi = get_indicator_value(prev_d, code_str)
            if prev_cuoi is None:
                prev_cuoi = get_indicator_value(prev_d, f"ct{code_str}")
            if prev_cuoi is not None and prev_cuoi != 0:
                v_dau = prev_cuoi

        return (v_cuoi if v_cuoi is not None else 0, v_dau if v_dau is not None else 0)

    def write_bctc_block(c_start: int, decls: List[TaxDeclaration]) -> int:
        c_end = c_start + 2 + len(decls) * 2

        # Hàng 3: Năm
        for idx, d in enumerate(decls):
            c1 = c_start + 3 + idx * 2
            c2 = c1 + 1
            ws.merge_cells(start_row=3, start_column=c1, end_row=3, end_column=c2)
            c_yr = ws.cell(3, c1, d.year or d.ky_kkhai)
            c_yr.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            c_yr.alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(3, c1).border = THIN_BORDER
            ws.cell(3, c2).border = THIN_BORDER

        # Hàng 4: Tiêu đề
        ws.cell(4, c_start, "TÀI SẢN ")
        ws.cell(4, c_start + 1, "Mã số")
        ws.cell(4, c_start + 2, "Thuyết minh")
        for idx in range(len(decls)):
            c1 = c_start + 3 + idx * 2
            c2 = c1 + 1
            ws.cell(4, c1, "Số cuối năm")
            ws.cell(4, c2, "Số đầu năm")

        # Hàng 5: Số thứ tự cột
        ws.cell(5, c_start, 1)
        ws.cell(5, c_start + 1, 2)
        ws.cell(5, c_start + 2, 3)
        for idx in range(len(decls)):
            c1 = c_start + 3 + idx * 2
            c2 = c1 + 1
            ws.cell(5, c1, 4)
            ws.cell(5, c2, 5)

        for col in range(c_start, c_end + 1):
            cell4 = ws.cell(4, col)
            cell4.fill = GRAY_HEADER_FILL
            cell4.font = BOLD_FONT
            cell4.border = THIN_BORDER
            cell4.alignment = Alignment(horizontal="center", vertical="center")

            cell5 = ws.cell(5, col)
            cell5.font = BODY_FONT
            cell5.border = THIN_BORDER
            cell5.alignment = Alignment(horizontal="center", vertical="center")

        cur_row = 6
        for c1_val, c2_val, c3_val in CDKT_ROWS:
            c_ts = ws.cell(cur_row, c_start, c1_val)
            c_ms = ws.cell(cur_row, c_start + 1, c2_val)
            c_tm = ws.cell(cur_row, c_start + 2, c3_val)

            is_major = (
                c1_val.startswith("A -")
                or c1_val.startswith("B -")
                or c1_val.startswith("C -")
                or c1_val.startswith("D -")
                or "TỔNG CỘNG" in c1_val
            )
            is_sub = (
                c1_val.startswith("I.")
                or c1_val.startswith("II.")
                or c1_val.startswith("III.")
                or c1_val.startswith("IV.")
                or c1_val.startswith("V.")
                or c1_val.startswith("VI.")
                or c1_val.startswith("I-")
            )

            c_ms.alignment = Alignment(horizontal="center", vertical="center")
            c_tm.alignment = Alignment(horizontal="center", vertical="center")

            for cell in (c_ts, c_ms, c_tm):
                cell.font = BOLD_FONT if (is_major or is_sub) else BODY_FONT
                cell.border = THIN_BORDER
                if is_major:
                    cell.fill = SECTION_FILL

            code_str = str(c2_val).strip()

            for idx, d in enumerate(decls):
                c_cuoi = c_start + 3 + idx * 2
                c_dau = c_cuoi + 1

                prev_d = None
                curr_year = d.year or 0
                for cand in sorted_decls:
                    if cand.year and cand.year == curr_year - 1:
                        prev_d = cand
                        break

                v_cuoi, v_dau = get_bctc_val_pair(d, prev_d, code_str)

                cell_c = ws.cell(cur_row, c_cuoi)
                cell_d = ws.cell(cur_row, c_dau)
                write_cell_formatted(cell_c, v_cuoi)
                write_cell_formatted(cell_d, v_dau)
                cell_c.border = THIN_BORDER
                cell_d.border = THIN_BORDER
                if is_major or is_sub:
                    cell_c.font = BOLD_FONT
                    cell_d.font = BOLD_FONT
                if is_major:
                    cell_c.fill = SECTION_FILL
                    cell_d.fill = SECTION_FILL

            cur_row += 1

        ws.column_dimensions[get_column_letter(c_start)].width = 54
        ws.column_dimensions[get_column_letter(c_start + 1)].width = 10
        ws.column_dimensions[get_column_letter(c_start + 2)].width = 14
        for idx in range(len(decls)):
            c1 = c_start + 3 + idx * 2
            c2 = c1 + 1
            ws.column_dimensions[get_column_letter(c1)].width = 18
            ws.column_dimensions[get_column_letter(c2)].width = 18

        return c_end

    if decls_156 and decls_80:
        c156_end = write_bctc_block(1, decls_156)
        c_sep1 = c156_end + 1
        c_sep2 = c156_end + 2
        ws.column_dimensions[get_column_letter(c_sep1)].width = 5
        ws.column_dimensions[get_column_letter(c_sep2)].width = 5
        write_bctc_block(c_sep2 + 1, decls_80)
    elif decls_156:
        write_bctc_block(1, decls_156)
    else:
        write_bctc_block(1, decls_80 if decls_80 else sorted_decls)

    ws.freeze_panes = "D6"


# ==============================================================================
# 4. CÁC HÀM XUẤT SHEET TNCN: SHEET TỔNG HỢP VÀ CÁC SHEET TỜ KHAI RIÊNG (TK, 05-1, 05-2, 05-3)
# ==============================================================================
def _compute_tncn_pl_summary(d: TaxDeclaration) -> Tuple[int, int, int, int, int, int]:
    """
    Tính trực tiếp 6 chỉ tiêu tổng hợp từ phụ lục 05-1 và 05-2 trong bộ nhớ của tờ khai d:
    (1: TNCT 05-1, 2: Thuế 05-1 sau bù trừ nộp thừa/phải nộp, 3: TNCT 05-2 cư trú 10%,
     4: Thuế 05-2 cư trú 10%, 5: TNCT 05-2 không cư trú 20%, 6: Thuế 05-2 không cư trú 20%).
    Đảm bảo sheet 'TNCN' trong file tổng hợp luôn có đầy đủ số liệu 100% kể cả khi không xuất kèm các sheet phụ lục.
    """
    is_tt156 = _is_tt156_decl(d)
    tnct_051 = 0
    thue_kt_051 = 0
    thue_thua_051 = 0
    thue_cpn_051 = 0

    tnct_052_res = 0
    thue_052_res = 0
    tnct_052_nonres = 0
    thue_052_nonres = 0

    def _to_int(v: Any) -> int:
        if v is None or v == "":
            return 0
        try:
            return int(float(v))
        except (ValueError, TypeError):
            return 0

    for pl_name, rows in d.appendices.items():
        pl_u = pl_name.upper()
        if any(t in pl_u for t in ("05_1", "05-1", "051")):
            for r in rows:
                row_156 = is_tt156 or ("caNhanNNUQQT" in r)
                if row_156:
                    tnct_051 += _to_int(_first_not_none(r, ("ct11", "tongTNCT", "11")))
                    thue_kt_051 += _to_int(_first_not_none(r, ("ct20", "thueDaKhauTru", "20")))
                    thue_thua_051 += _to_int(_first_not_none(r, ("ct23", "thueNopThua", "23")))
                    thue_cpn_051 += _to_int(_first_not_none(r, ("ct24", "thueConPhaiNop", "24")))
                else:
                    tnct_051 += _to_int(_first_not_none(r, ("ct12", "tongTNCT", "12")))
                    thue_kt_051 += _to_int(_first_not_none(r, ("ct22", "thueDaKhauTru", "22")))
                    thue_thua_051 += _to_int(_first_not_none(r, ("ct25", "thueNopThua", "25")))
                    thue_cpn_051 += _to_int(_first_not_none(r, ("ct26", "thueConPhaiNop", "26")))
        elif any(t in pl_u for t in ("05_2", "05-2", "052")):
            for r in rows:
                kct_raw = _first_not_none(r, ("ct10", "caNhanKhongCuTru", "10"))
                is_nonres = kct_raw in (1, "1", True, "true", "True", "[x]")
                tn = _to_int(_first_not_none(r, ("ct11", "tongTNCT", "11")))
                th = _to_int(_first_not_none(r, ("ct15", "thueDaKhauTru", "15")))
                if is_nonres:
                    tnct_052_nonres += tn
                    thue_052_nonres += th
                else:
                    tnct_052_res += tn
                    thue_052_res += th

    thue_051_net = thue_kt_051 - thue_thua_051 + thue_cpn_051
    return (tnct_051, thue_051_net, tnct_052_res, thue_052_res, tnct_052_nonres, thue_052_nonres)


def write_sheet_tncn_master(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    include_appendices: bool = True,
) -> None:
    """
    Xuất tờ khai quyết toán TNCN của tất cả các năm vào DUY NHẤT 1 SHEET 'TNCN' trong file tổng hợp.
    - Khối bên trái: Bảng theo TT156 (áp dụng cho các năm <= 2020, maTKhai='395')
    - Khối bên phải: Bảng theo TT80 (áp dụng cho các năm >= 2021, maTKhai='953')
    """
    sname = sanitize_sheet_name("TNCN", existing_sheets)
    ws = wb.create_sheet(title=sname)

    latest_tncn = filter_latest_per_year(declarations)
    latest_tncn = sorted(latest_tncn, key=lambda d: d.year or 0)

    decls_156 = [d for d in latest_tncn if _is_tt156_decl(d)]
    decls_80 = [d for d in latest_tncn if not _is_tt156_decl(d)]

    # 1. Khối TT156 (nếu có các năm <= 2020)
    if decls_156:
        c_156_start = 3
        c_156_end = c_156_start + 3 + len(decls_156)

        for i, d in enumerate(decls_156):
            c_yr = c_156_start + 4 + i
            c_cell = ws.cell(1, c_yr, d.year)
            c_cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            c_cell.alignment = Alignment(horizontal="center", vertical="center")
            c_cell.border = THIN_BORDER

        for r_idx, stt, desc, code, dvt in TNCN_MASTER_ROWS_TT156:
            if r_idx == 1:
                continue
            elif r_idx == 2:
                headers = ["STT", "Chỉ tiêu", "Mã chỉ tiêu", "Đơn vị tính"]
                for col_offset, h in enumerate(headers):
                    cell = ws.cell(2, c_156_start + col_offset, h)
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for i, d in enumerate(decls_156):
                    cell = ws.cell(2, c_156_start + 4 + i, "Số người/ Số tiền")
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            elif r_idx in (24, 26):
                continue
            elif r_idx == 25:
                cell = ws.cell(25, c_156_start, stt)
                cell.font = BOLD_FONT
                cell.fill = SECTION_FILL
                ws.merge_cells(start_row=25, start_column=c_156_start, end_row=25, end_column=c_156_end)
                for col_m in range(c_156_start, c_156_end + 1):
                    ws.cell(25, col_m).border = THIN_BORDER
                    ws.cell(25, col_m).fill = SECTION_FILL
            elif r_idx == 27:
                headers = ["STT", "Chỉ tiêu", "Mã chỉ tiêu", "Đơn vị tính"]
                for col_offset, h in enumerate(headers):
                    cell = ws.cell(27, c_156_start + col_offset, h)
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for i, d in enumerate(decls_156):
                    cell = ws.cell(27, c_156_start + 4 + i, "Số người/ Số tiền")
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                ws.cell(r_idx, c_156_start, stt).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(r_idx, c_156_start + 1, desc).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(r_idx, c_156_start + 2, code).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(r_idx, c_156_start + 3, dvt).alignment = Alignment(horizontal="center", vertical="center")
                for col_offset in range(4):
                    ws.cell(r_idx, c_156_start + col_offset).font = BODY_FONT
                    ws.cell(r_idx, c_156_start + col_offset).border = THIN_BORDER

                for i, d in enumerate(decls_156):
                    c_yr = c_156_start + 4 + i
                    cell = ws.cell(r_idx, c_yr)
                    val = get_indicator_value(d, code) if code else None
                    write_cell_formatted(cell, val if val is not None else 0)
                    if dvt != "VNĐ":
                        cell.number_format = "#,##0"
                    cell.border = THIN_BORDER
                    cell.font = BODY_FONT

        ws.merge_cells(start_row=3, start_column=c_156_start, end_row=4, end_column=c_156_start)

        ws.column_dimensions[get_column_letter(c_156_start)].width = 6
        ws.column_dimensions[get_column_letter(c_156_start + 1)].width = 50
        ws.column_dimensions[get_column_letter(c_156_start + 2)].width = 14
        ws.column_dimensions[get_column_letter(c_156_start + 3)].width = 12
        for i in range(len(decls_156)):
            ws.column_dimensions[get_column_letter(c_156_start + 4 + i)].width = 16

    # 2. Khối TT80 (các năm >= 2021)
    if decls_80:
        if decls_156:
            c_sep = c_156_end + 1
            ws.column_dimensions[get_column_letter(c_sep)].width = 10
            c_80_start = c_sep + 1
        else:
            c_80_start = 3

        c_80_end = c_80_start + 3 + len(decls_80)

        for i, d in enumerate(decls_80):
            c_yr = c_80_start + 4 + i
            c_cell = ws.cell(1, c_yr, d.year)
            c_cell.font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
            c_cell.alignment = Alignment(horizontal="center", vertical="center")
            c_cell.border = THIN_BORDER

        for r_idx, stt, desc, code, dvt in TNCN_MASTER_ROWS_TT80:
            if r_idx == 1:
                continue
            elif r_idx == 2:
                headers = ["STT", "Chỉ tiêu", "Mã chỉ tiêu", "Đơn vị tính"]
                for col_offset, h in enumerate(headers):
                    cell = ws.cell(2, c_80_start + col_offset, h)
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for i, d in enumerate(decls_80):
                    cell = ws.cell(2, c_80_start + 4 + i, "Số người/ Số tiền")
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            elif r_idx in (22, 24):
                continue
            elif r_idx == 23:
                cell = ws.cell(23, c_80_start, stt)
                cell.font = BOLD_FONT
                cell.fill = SECTION_FILL
                ws.merge_cells(start_row=23, start_column=c_80_start, end_row=23, end_column=c_80_end)
                for col_m in range(c_80_start, c_80_end + 1):
                    ws.cell(23, col_m).border = THIN_BORDER
                    ws.cell(23, col_m).fill = SECTION_FILL
            elif r_idx == 25:
                headers = ["STT", "Chỉ tiêu", "Mã chỉ tiêu", "Đơn vị tính"]
                for col_offset, h in enumerate(headers):
                    cell = ws.cell(25, c_80_start + col_offset, h)
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                for i, d in enumerate(decls_80):
                    cell = ws.cell(25, c_80_start + 4 + i, "Số người/ Số tiền")
                    cell.font = BOLD_FONT
                    cell.fill = GRAY_HEADER_FILL
                    cell.border = THIN_BORDER
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                ws.cell(r_idx, c_80_start, stt).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(r_idx, c_80_start + 1, desc).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(r_idx, c_80_start + 2, code).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(r_idx, c_80_start + 3, dvt).alignment = Alignment(horizontal="center", vertical="center")
                for col_offset in range(4):
                    ws.cell(r_idx, c_80_start + col_offset).font = BODY_FONT
                    ws.cell(r_idx, c_80_start + col_offset).border = THIN_BORDER

                for i, d in enumerate(decls_80):
                    c_yr = c_80_start + 4 + i
                    cell = ws.cell(r_idx, c_yr)
                    val = get_indicator_value(d, code) if code else None
                    write_cell_formatted(cell, val if val is not None else 0)
                    if dvt != "VNĐ":
                        cell.number_format = "#,##0"
                    cell.border = THIN_BORDER
                    cell.font = BODY_FONT

        ws.merge_cells(start_row=3, start_column=c_80_start, end_row=4, end_column=c_80_start)
        ws.merge_cells(start_row=27, start_column=c_80_start, end_row=28, end_column=c_80_start)

        ws.column_dimensions[get_column_letter(c_80_start)].width = 6
        ws.column_dimensions[get_column_letter(c_80_start + 1)].width = 50
        ws.column_dimensions[get_column_letter(c_80_start + 2)].width = 12
        ws.column_dimensions[get_column_letter(c_80_start + 3)].width = 14
        for i in range(len(decls_80)):
            ws.column_dimensions[get_column_letter(c_80_start + 4 + i)].width = 15

    # 3. BẢNG TỔNG HỢP ĐỐI CHIẾU TỪ SHEET '05-1' VÀ '05-2' (Dòng 38..47) TRÊN SHEET 'TNCN'
    r051_map: Dict[int, int] = {}
    r052_map: Dict[int, Tuple[int, int]] = {}
    if include_appendices:
        if len(latest_tncn) == 1:
            d0 = latest_tncn[0]
            n1 = sum(len(r) for k, r in d0.appendices.items() if any(t in k.upper() for t in ("05_1", "05-1", "051")))
            n2 = sum(len(r) for k, r in d0.appendices.items() if any(t in k.upper() for t in ("05_2", "05-2", "052")))
            if n1 > 0:
                r051_map[id(d0)] = 4 + n1
            if n2 > 0:
                r052_map[id(d0)] = (4 + n2 + 1, 4 + n2 + 2)
        else:
            cur1 = 4
            for d_item in latest_tncn:
                n1 = sum(len(r) for k, r in d_item.appendices.items() if any(t in k.upper() for t in ("05_1", "05-1", "051")))
                if n1 > 0:
                    cur1 += 1 + n1
                    r051_map[id(d_item)] = cur1
                    cur1 += 1
            cur2 = 4
            for d_item in latest_tncn:
                n2 = sum(len(r) for k, r in d_item.appendices.items() if any(t in k.upper() for t in ("05_2", "05-2", "052")))
                if n2 > 0:
                    cur2 += 1 + n2
                    r052_map[id(d_item)] = (cur2 + 1, cur2 + 2)
                    cur2 += 3

    tnr_bold = Font(name="Times New Roman", size=10, bold=True)
    tnr_reg = Font(name="Times New Roman", size=10, bold=False)
    yellow_hl = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    blocks_to_write: List[Tuple[int, List[TaxDeclaration]]] = []
    if decls_156:
        blocks_to_write.append((3, decls_156))
    if decls_80:
        b80_start = (3 + 3 + len(decls_156) + 2) if decls_156 else 3
        blocks_to_write.append((b80_start, decls_80))

    ws.row_dimensions[38].height = 25.5
    for r_num in range(41, 47):
        ws.row_dimensions[r_num].height = 26.25

    summary_rows_spec = [
        (41, 1, "Thu nhập từ tiền lương, tiền công của cá nhân cư trú tính thuế theo biểu thuế lũy tiến từng phần "),
        (42, 2, "Thuế TNCN từ tiền lương, tiền công của cá nhân cư trú tính thuế theo biểu thuế lũy tiến từng phần đã khấu trừ"),
        (43, 3, "Thu nhập từ tiền lương, tiền công của cá nhân cư trú tính thuế theo thuế suất toàn phần 10% "),
        (44, 4, "Thuế TNCN từ tiền lương, tiền công của cá nhân cư trú tính thuế theo thuế suất toàn phần 10% đã khấu trừ"),
        (45, 5, "Thu nhập từ tiền lương, tiền công của cá nhân  không cư trú"),
        (46, 6, "Thuế TNCN từ tiền lương, tiền công của cá nhân không cư trú tính theo thuế suất 20% đã khấu trừ"),
    ]

    for b_start, b_decls in blocks_to_write:
        b_end = b_start + 3 + len(b_decls)
        # Dòng 38: Tiêu đề
        ws.cell(38, b_start, "STT")
        ws.cell(38, b_start + 1, "Chỉ tiêu")
        ws.cell(38, b_start + 2, "Mã chỉ tiêu")
        ws.cell(38, b_start + 3, "Đơn vị tính")
        for i, d_item in enumerate(b_decls):
            ws.cell(38, b_start + 4 + i, f"Số báo cáo ({d_item.year or d_item.ky_kkhai})")
        for c_i in range(b_start, b_end + 1):
            c38 = ws.cell(38, c_i)
            c38.font = tnr_bold
            c38.fill = GRAY_HEADER_FILL
            c38.border = THIN_BORDER
            c38.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # Dòng 39: Mã cột
        ws.cell(39, b_start, -1)
        ws.cell(39, b_start + 1, -2)
        ws.cell(39, b_start + 2, "")
        ws.cell(39, b_start + 3, "")
        for i in range(len(b_decls)):
            ws.cell(39, b_start + 4 + i, -3)
        for c_i in range(b_start, b_end + 1):
            c39 = ws.cell(39, c_i)
            c39.font = tnr_reg
            c39.border = THIN_BORDER
            c39.alignment = Alignment(horizontal="center", vertical="center")
            if isinstance(c39.value, int):
                c39.number_format = "0_);(0)"

        # Dòng 40: Mục I
        ws.cell(40, b_start, "I").font = tnr_bold
        ws.cell(40, b_start).alignment = Alignment(horizontal="center", vertical="center")
        ws.cell(40, b_start + 1, "Thuế TNCN từ tiền lương, tiền công ").font = tnr_bold
        ws.merge_cells(start_row=40, start_column=b_start + 1, end_row=40, end_column=b_end)
        for c_i in range(b_start, b_end + 1):
            ws.cell(40, c_i).border = THIN_BORDER
            ws.cell(40, c_i).fill = SECTION_FILL

        # Dòng 41..46: 6 chỉ tiêu
        for r_num, stt_num, lbl_text in summary_rows_spec:
            c_stt = ws.cell(r_num, b_start, stt_num)
            c_stt.font = tnr_reg
            c_stt.alignment = Alignment(horizontal="center", vertical="center")
            c_stt.border = THIN_BORDER

            c_lbl = ws.cell(r_num, b_start + 1, lbl_text)
            c_lbl.font = tnr_reg
            c_lbl.alignment = Alignment(horizontal="justify", vertical="center", wrap_text=True)
            c_lbl.border = THIN_BORDER

            c_code = ws.cell(r_num, b_start + 2, "")
            c_code.font = tnr_reg
            c_code.border = THIN_BORDER

            c_dvt = ws.cell(r_num, b_start + 3, "VNĐ")
            c_dvt.font = tnr_reg
            c_dvt.alignment = Alignment(horizontal="center", vertical="center")
            c_dvt.border = THIN_BORDER

            for i, d_item in enumerate(b_decls):
                c_yr = b_start + 4 + i
                cell = ws.cell(r_num, c_yr)
                r1 = r051_map.get(id(d_item))
                r2_pair = r052_map.get(id(d_item))
                calc_vals = _compute_tncn_pl_summary(d_item)
                if r_num == 41:
                    cell.value = f"='05-1'!H{r1}" if r1 else calc_vals[0]
                elif r_num == 42:
                    cell.value = f"='05-1'!R{r1}-'05-1'!U{r1}+'05-1'!V{r1}" if r1 else calc_vals[1]
                elif r_num == 43:
                    cell.value = f"='05-2'!G{r2_pair[0]}" if r2_pair else calc_vals[2]
                elif r_num == 44:
                    cell.value = f"='05-2'!K{r2_pair[0]}" if r2_pair else calc_vals[3]
                elif r_num == 45:
                    cell.value = f"='05-2'!G{r2_pair[1]}" if r2_pair else calc_vals[4]
                elif r_num == 46:
                    cell.value = f"='05-2'!K{r2_pair[1]}" if r2_pair else calc_vals[5]
                cell.font = tnr_reg
                cell.number_format = "#,##0"
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.border = THIN_BORDER
                if r_num == 42:
                    cell.fill = yellow_hl

        # Dòng 47: Tổng hợp tổng số thuế TNCN phải khấu trừ
        ws.cell(47, b_start, "Tổng hợp tổng số thuế TNCN phải khấu trừ").font = tnr_bold
        ws.merge_cells(start_row=47, start_column=b_start, end_row=47, end_column=b_start + 3)
        for c_i in range(b_start, b_start + 4):
            ws.cell(47, c_i).border = THIN_BORDER
        for i in range(len(b_decls)):
            c_yr = b_start + 4 + i
            col_let = get_column_letter(c_yr)
            c_tot = ws.cell(47, c_yr, f"={col_let}42+{col_let}44+{col_let}46")
            c_tot.font = tnr_bold
            c_tot.number_format = "#,##0"
            c_tot.alignment = Alignment(horizontal="right", vertical="center")
            c_tot.border = THIN_BORDER
            c_tot.fill = yellow_hl

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 3
    ws.freeze_panes = "C3"


def write_sheet_tncn_tk(ws: openpyxl.worksheet.worksheet.Worksheet, d: TaxDeclaration) -> None:
    """
    Xuất sheet 'TK' của tờ khai quyết toán TNCN theo đúng mẫu thực tế:
    - Nếu thuộc Thông tư 156/92 (năm <= 2020 hoặc maTKhai='395'): Dùng bộ mã [21]..[46]
    - Nếu thuộc Thông tư 80 (năm >= 2021 hoặc maTKhai='953'): Dùng bộ mã [16]..[41]
    """
    yr = d.year or d.ky_kkhai
    is_tt156 = _is_tt156_decl(d)
    rows_catalog = TNCN_TT156_ROWS if is_tt156 else TNCN_TT80_ROWS

    ws.cell(1, 6, yr).font = Font(name="Calibri", size=11, bold=True, color="1F4E78")
    ws.cell(1, 6).alignment = Alignment(horizontal="center", vertical="center")

    headers = ["", "STT", "Chỉ tiêu", "Mã chỉ tiêu", "Đơn vị tính", "Số người / Số tiền"]
    ws.append(headers)
    ws.row_dimensions[2].height = 24

    for c in range(2, 7):
        cell = ws.cell(2, c)
        cell.fill = GRAY_HEADER_FILL
        cell.font = BOLD_FONT
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")

    cur_row = 3
    for stt, desc, code, dvt in rows_catalog:
        c_stt = ws.cell(cur_row, 2, stt)
        c_desc = ws.cell(cur_row, 3, desc)
        c_code = ws.cell(cur_row, 4, code)
        c_dvt = ws.cell(cur_row, 5, dvt)

        is_sec = stt == "II" or "NGHĨA VỤ QUYẾT TOÁN" in stt or (not code and not dvt)
        if is_sec:
            for cell in (c_stt, c_desc, c_code, c_dvt):
                cell.font = BOLD_FONT
                cell.fill = SECTION_FILL
                cell.border = THIN_BORDER
            ws.cell(cur_row, 6, "").fill = SECTION_FILL
            ws.cell(cur_row, 6).border = THIN_BORDER
            cur_row += 1
            continue

        c_stt.alignment = Alignment(horizontal="center", vertical="center")
        c_code.alignment = Alignment(horizontal="center", vertical="center")
        c_dvt.alignment = Alignment(horizontal="center", vertical="center")
        for cell in (c_stt, c_desc, c_code, c_dvt):
            cell.font = BODY_FONT
            cell.border = THIN_BORDER

        val = get_indicator_value(d, code)
        write_cell_formatted(ws.cell(cur_row, 6), val if val is not None else 0)
        cur_row += 1

    # BẢNG 2 TRÊN SHEET 'TK' (Dòng 38..47): Tổng hợp đối chiếu Thuế TNCN từ tiền lương, tiền công lấy từ sheet '05-1' và '05-2'
    n_051 = sum(
        len(r)
        for k, r in d.appendices.items()
        if any(t in k.upper() for t in ("05_1", "05-1", "051"))
    )
    n_052 = sum(
        len(r)
        for k, r in d.appendices.items()
        if any(t in k.upper() for t in ("05_2", "05-2", "052"))
    )
    r_sum_051 = 4 + n_051
    r_sum_052 = 4 + n_052
    r_res_052 = r_sum_052 + 1
    r_nonres_052 = r_sum_052 + 2

    tnr_bold = Font(name="Times New Roman", size=10, bold=True)
    tnr_reg = Font(name="Times New Roman", size=10, bold=False)
    yellow_hl = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    # Dòng 38: Tiêu đề cột
    ws.row_dimensions[38].height = 25.5
    r38_headers = ["STT", "Chỉ tiêu", "Số báo cáo của người nộp thuế", "Số kiểm tra", "Chênh lệch"]
    for idx_h, h_text in enumerate(r38_headers, start=2):
        cell = ws.cell(38, idx_h, h_text)
        cell.font = tnr_bold
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Dòng 39: Mã số cột (-1, -2, -3, -4, (5)=(4)-(3))
    r39_codes: List[Any] = [-1, -2, -3, -4, "(5)=(4)-(3)"]
    for idx_c, c_val in enumerate(r39_codes, start=2):
        cell = ws.cell(39, idx_c, c_val)
        cell.font = tnr_reg
        cell.border = THIN_BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if isinstance(c_val, int):
            cell.number_format = "0_);(0)"

    # Dòng 40: Mục I. Thuế TNCN từ tiền lương, tiền công
    c_i = ws.cell(40, 2, "I")
    c_i.font = tnr_bold
    c_i.alignment = Alignment(horizontal="center", vertical="center")
    c_i.border = THIN_BORDER

    c_i_title = ws.cell(40, 3, "Thuế TNCN từ tiền lương, tiền công ")
    c_i_title.font = tnr_bold
    c_i_title.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("C40:F40")
    for col_m in range(3, 7):
        ws.cell(40, col_m).border = THIN_BORDER

    # Dòng 41..46: 6 chỉ tiêu lấy công thức trực tiếp từ sheet '05-1' và '05-2'
    summary_items = [
        (
            41,
            1,
            "Thu nhập từ tiền lương, tiền công của cá nhân cư trú tính thuế theo biểu thuế lũy tiến từng phần ",
            f"='05-1'!H{r_sum_051}",
        ),
        (
            42,
            2,
            "Thuế TNCN từ tiền lương, tiền công của cá nhân cư trú tính thuế theo biểu thuế lũy tiến từng phần đã khấu trừ",
            f"='05-1'!R{r_sum_051}-'05-1'!U{r_sum_051}+'05-1'!V{r_sum_051}",
        ),
        (
            43,
            3,
            "Thu nhập từ tiền lương, tiền công của cá nhân cư trú tính thuế theo thuế suất toàn phần 10% ",
            f"='05-2'!G{r_res_052}",
        ),
        (
            44,
            4,
            "Thuế TNCN từ tiền lương, tiền công của cá nhân cư trú tính thuế theo thuế suất toàn phần 10% đã khấu trừ",
            f"='05-2'!K{r_res_052}",
        ),
        (
            45,
            5,
            "Thu nhập từ tiền lương, tiền công của cá nhân  không cư trú",
            f"='05-2'!G{r_nonres_052}",
        ),
        (
            46,
            6,
            "Thuế TNCN từ tiền lương, tiền công của cá nhân không cư trú tính theo thuế suất 20% đã khấu trừ",
            f"='05-2'!K{r_nonres_052}",
        ),
    ]

    for r_num, stt_num, label_text, formula_d in summary_items:
        ws.row_dimensions[r_num].height = 26.25
        c2 = ws.cell(r_num, 2, stt_num)
        c2.font = tnr_reg
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = THIN_BORDER

        c3 = ws.cell(r_num, 3, label_text)
        c3.font = tnr_reg
        c3.alignment = Alignment(horizontal="justify", vertical="center", wrap_text=True)
        c3.border = THIN_BORDER

        c4 = ws.cell(r_num, 4, formula_d)
        c4.font = tnr_reg
        c4.number_format = "#,##0"
        c4.alignment = Alignment(horizontal="right", vertical="center")
        c4.border = THIN_BORDER

        c5 = ws.cell(r_num, 5, f"=D{r_num}")
        c5.font = tnr_reg
        c5.number_format = '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)'
        c5.alignment = Alignment(horizontal="right", vertical="center")
        c5.border = THIN_BORDER

        c6 = ws.cell(r_num, 6, f"=E{r_num}-D{r_num}")
        c6.font = tnr_reg
        c6.number_format = "#,##0"
        c6.alignment = Alignment(horizontal="right", vertical="center")
        c6.border = THIN_BORDER

        if r_num == 42:
            for c_hl in (c4, c5, c6):
                c_hl.fill = yellow_hl

    # Dòng 47: Tổng hợp tổng số thuế TNCN phải khấu trừ
    c47_lbl = ws.cell(47, 2, "Tổng hợp tổng số thuế TNCN phải khấu trừ")
    c47_lbl.font = tnr_bold
    c47_lbl.alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells("B47:C47")
    ws.cell(47, 2).border = THIN_BORDER
    ws.cell(47, 3).border = THIN_BORDER

    for col_tot, col_let in ((4, "D"), (5, "E"), (6, "F")):
        c_tot = ws.cell(47, col_tot, f"={col_let}42+{col_let}44+{col_let}46")
        c_tot.font = tnr_bold
        c_tot.number_format = "#,##0"
        c_tot.alignment = Alignment(horizontal="right", vertical="center")
        c_tot.border = THIN_BORDER
        if col_tot == 4:
            c_tot.fill = yellow_hl

    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 8
    ws.column_dimensions["C"].width = 54
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 18
    ws.freeze_panes = "A3"


def _first_not_none(row_dict: Dict[str, Any], keys: Sequence[str]) -> Any:
    """Tìm giá trị đầu tiên khác None trong row_dict theo danh sách khóa ưu tiên (không làm mất giá trị 0)."""
    for k in keys:
        if not k:
            continue
        for variant in (k, k.lower(), k.upper()):
            if variant in row_dict and row_dict[variant] is not None:
                return row_dict[variant]
    return None


def write_sheet_tncn_pl(
    ws: openpyxl.worksheet.worksheet.Worksheet,
    cols_def: List[Tuple[str, str, str]],
    merges_def: List[str],
    rows: List[Dict[str, Any]],
    sheet_type: str,
    is_tt156: bool = False,
) -> None:
    """
    Xuất các phụ lục 05-1, 05-2, 05-3 theo đúng định dạng 3 dòng tiêu đề của HTKK.
    Tự động ánh xạ chuẩn xác 100% các thẻ XML của cả Thông tư 156/92 (<=2020) và Thông tư 80 (>=2021),
    không để lệch cột hoặc mất dữ liệu!
    """
    h1 = [c[0] for c in cols_def]
    h2 = [c[1] for c in cols_def]
    h3 = [c[2] for c in cols_def]

    ws.append(h1)
    ws.append(h2)
    ws.append(h3)

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 28
    ws.row_dimensions[3].height = 18

    for m in merges_def:
        try:
            ws.merge_cells(m)
        except Exception:
            pass

    for r in (1, 2, 3):
        for c in range(1, len(cols_def) + 1):
            cell = ws.cell(r, c)
            cell.fill = GRAY_HEADER_FILL
            cell.font = BOLD_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Định nghĩa danh sách khóa tra cứu chính xác cho từng cột theo Thông tư 80 và Thông tư 156/92
    if sheet_type == "05-1":
        if is_tt156:
            # Trong TT156/92 (maTKhai=395): caNhanNNUQQT là checkbox cột [11], ct11 là Tổng TNCT (cột [12]),
            # ct12 là TNCT giảm thuế KKT (cột [13]), ct13 là TNCT miễn Hiệp định (cột [14]),
            # ct14 là Số lượng NPT (cột [16]), ct15..ct24 tương ứng cột [17]..[26]!
            col_candidates: List[Tuple[str, Sequence[str]]] = [
                ("stt", ("stt",)),
                ("hoVaTen", ("ct07", "hoVaTen", "07")),
                ("mst", ("ct08", "mst", "08")),
                ("loaiGiayTo", ("ct09a_ten", "loaiGiayTo")),
                ("soCMT", ("ct09", "soCMT", "09")),
                ("caNhanUyQuyen", ("ct10", "caNhanUyQuyen", "10")),
                ("caNhanNuocNgoai", ("caNhanNNUQQT", "caNhanNuocNgoai")),
                ("tongTNCT", ("ct11", "tongTNCT", "11")),
                ("tnctTaiToChuc", ("ct12", "tnctTaiToChuc", "12")),
                ("tnctMienHiepDinh", ("ct13", "tnctMienHiepDinh", "13")),
                ("tnctMienDauKhi", ("tnctMienDauKhi",)),
                ("soLuongNPT", ("ct14", "soLuongNPT", "14")),
                ("tongGiamTruGiaCanh", ("ct15", "tongGiamTruGiaCanh", "15")),
                ("tuThien", ("ct16", "tuThien", "16")),
                ("baoHiem", ("ct17", "baoHiem", "17")),
                ("huuTri", ("ct18", "huuTri", "18")),
                ("tntt", ("ct19", "tntt", "19")),
                ("thueDaKhauTru", ("ct20", "thueDaKhauTru", "20")),
                ("thueDaKhauTruTaiToChuc", ("ct21", "thueDaKhauTruTaiToChuc", "21")),
                ("tongThuePhaiNop", ("ct22", "tongThuePhaiNop", "22")),
                ("thueNopThua", ("ct23", "thueNopThua", "23")),
                ("thueConPhaiNop", ("ct24", "thueConPhaiNop", "24")),
                ("caNhanMienDuoi50k", ("caNhanMienDuoi50k",)),
            ]
        else:
            col_candidates = [
                ("stt", ("stt",)),
                ("hoVaTen", ("ct07", "hoVaTen", "07")),
                ("mst", ("ct08", "mst", "08")),
                ("loaiGiayTo", ("ct09a_ten", "loaiGiayTo")),
                ("soCMT", ("ct09", "soCMT", "09")),
                ("caNhanUyQuyen", ("ct10", "caNhanUyQuyen", "10")),
                ("caNhanNuocNgoai", ("ct11", "caNhanNuocNgoai", "11")),
                ("tongTNCT", ("ct12", "tongTNCT", "12")),
                ("tnctTaiToChuc", ("ct13", "tnctTaiToChuc", "13")),
                ("tnctMienHiepDinh", ("ct14", "tnctMienHiepDinh", "14")),
                ("tnctMienDauKhi", ("ct15", "tnctMienDauKhi", "15")),
                ("soLuongNPT", ("ct16", "soLuongNPT", "16")),
                ("tongGiamTruGiaCanh", ("ct17", "tongGiamTruGiaCanh", "17")),
                ("tuThien", ("ct18", "tuThien", "18")),
                ("baoHiem", ("ct19", "baoHiem", "19")),
                ("huuTri", ("ct20", "huuTri", "20")),
                ("tntt", ("ct21", "tntt", "21")),
                ("thueDaKhauTru", ("ct22", "thueDaKhauTru", "22")),
                ("thueDaKhauTruTaiToChuc", ("ct23", "thueDaKhauTruTaiToChuc", "23")),
                ("tongThuePhaiNop", ("ct24", "tongThuePhaiNop", "24")),
                ("thueNopThua", ("ct25", "thueNopThua", "25")),
                ("thueConPhaiNop", ("ct26", "thueConPhaiNop", "26")),
                ("caNhanMienDuoi50k", ("ct27", "caNhanMienDuoi50k", "27")),
            ]
    elif sheet_type == "05-2":
        col_candidates = [
            ("stt", ("stt",)),
            ("hoVaTen", ("ct07", "hoVaTen", "07")),
            ("mst", ("ct08", "mst", "08")),
            ("loaiGiayTo", ("ct09a_ten", "loaiGiayTo")),
            ("soCMT", ("ct09", "soCMT", "09")),
            ("caNhanKhongCuTru", ("ct10", "caNhanKhongCuTru", "10")),
            ("tongTNCT", ("ct11", "tongTNCT", "11")),
            ("tnctBaoHiem", ("ct12", "tnctBaoHiem", "12")),
            ("tnctMienHiepDinh", ("ct13", "tnctMienHiepDinh", "13")),
            ("tnctMienDauKhi", ("ct14", "tnctMienDauKhi", "14")),
            ("thueDaKhauTru", ("ct15", "thueDaKhauTru", "15")),
            ("thueBaoHiem", ("ct16", "thueBaoHiem", "16")),
        ]
    else:  # 05-3
        if is_tt156:
            # Trong TT156/92 (maTKhai=395): ct12_ten là Quốc tịch, ct13 là CMND/Hộ chiếu,
            # ct14_ten là Quan hệ với NNT, ct15/ct16 là Quyển số/Số GKS, còn ct21 là Từ tháng, ct22 là Đến tháng!
            col_candidates = [
                ("stt", ("stt",)),
                ("hoVaTen", ("ct07", "hoVaTen", "07")),
                ("mst", ("ct08", "mst", "08")),
                ("hoVaTenNPT", ("ct09", "hoVaTenNPT", "09")),
                ("ngaySinhNPT", ("ct10", "ngaySinhNPT", "10")),
                ("mstNPT", ("ct11", "mstNPT", "11")),
                ("loaiGiayTo", ("loaiGiayTo", "ct12_ten", "ct12")),
                ("soGiayTo", ("ct13", "soGiayTo", "13")),
                ("quanHe", ("ct14_ten", "quanHe", "ct14", "14")),
                ("tuThang", ("ct21", "tuThang", "ct15", "15")),
                ("denThang", ("ct22", "denThang", "ct16", "16")),
            ]
        else:
            col_candidates = [
                ("stt", ("stt",)),
                ("hoVaTen", ("ct07", "hoVaTen", "07")),
                ("mst", ("ct08", "mst", "08")),
                ("hoVaTenNPT", ("ct09", "hoVaTenNPT", "09")),
                ("ngaySinhNPT", ("ct10", "ngaySinhNPT", "10")),
                ("mstNPT", ("ct11", "mstNPT", "11")),
                ("loaiGiayTo", ("ct12_ten", "loaiGiayTo", "ct12", "12")),
                ("soGiayTo", ("ct13", "soGiayTo", "13")),
                ("quanHe", ("ct14_ten", "quanHe", "ct14", "14")),
                ("tuThang", ("ct15", "tuThang", "15")),
                ("denThang", ("ct16", "denThang", "16")),
            ]

    checkbox_fields = {"caNhanUyQuyen", "caNhanNuocNgoai", "caNhanKhongCuTru", "caNhanMienDuoi50k"}
    center_fields = checkbox_fields | {"stt", "quanHe", "tuThang", "denThang"}

    cur_row = 4
    for idx, r in enumerate(rows, start=1):
        # Phát hiện tự động dòng thuộc cấu trúc TT156 nếu có ct21/ct22 dạng tháng ở 05-3 hoặc caNhanNNUQQT ở 05-1
        row_is_156 = is_tt156 or ("caNhanNNUQQT" in r) or (sheet_type == "05-3" and ("ct21" in r or "ct22" in r))

        for col_idx, (field_name, cand_keys) in enumerate(col_candidates, start=1):
            if field_name == "stt":
                val: Any = idx
            else:
                val = _first_not_none(r, cand_keys)

                # Xử lý thông minh cho 05-3 TT156 khi NPT là trẻ em khai theo Giấy khai sinh (ct13 trống nhưng có ct15/ct16)
                if sheet_type == "05-3" and row_is_156:
                    if field_name == "soGiayTo" and (val is None or str(val).strip() == ""):
                        q_so = _first_not_none(r, ("ct15",))
                        s_gks = _first_not_none(r, ("ct16",))
                        if q_so is not None and s_gks is not None:
                            val = f"QS: {q_so} - Số: {s_gks}"
                        elif q_so is not None:
                            val = str(q_so)
                        elif s_gks is not None:
                            val = str(s_gks)
                    elif field_name == "loaiGiayTo":
                        has_cmnd = _first_not_none(r, ("ct13",)) not in (None, "")
                        has_gks = _first_not_none(r, ("ct15", "ct16")) not in (None, "")
                        if not has_cmnd and has_gks:
                            val = "Giấy khai sinh"

            if field_name in checkbox_fields:
                val = "[x]" if val in (1, "1", True, "true", "True", "[x]") else "[]"

            is_mst_col = any(x in field_name.lower() for x in ("mst", "socmt", "sogiayto", "cccd"))
            cell = ws.cell(cur_row, col_idx)
            write_cell_formatted(cell, val, is_mst=is_mst_col)
            if field_name in center_fields:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        cur_row += 1

    if sheet_type in ("05-1", "05-2"):
        sum_row = cur_row
        ws.cell(sum_row, 1, "Tổng").font = BOLD_FONT
        ws.cell(sum_row, 1).alignment = Alignment(horizontal="center", vertical="center")
        merge_range = f"A{sum_row}:G{sum_row}" if sheet_type == "05-1" else f"A{sum_row}:F{sum_row}"
        ws.merge_cells(merge_range)

        start_col = 8 if sheet_type == "05-1" else 7
        for c in range(1, len(cols_def) + 1):
            cell = ws.cell(sum_row, c)
            cell.font = BOLD_FONT
            cell.border = THIN_BORDER
            if c >= start_col and c != (23 if sheet_type == "05-1" else 99):
                cl = get_column_letter(c)
                if len(rows) > 0:
                    cell.value = f"=SUM({cl}4:{cl}{sum_row-1})"
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    cell.value = 0
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        # Với phụ lục 05-2: Thêm 2 dòng tổng hợp tự động tách riêng Cá nhân cư trú (10%, F="[]")
        # và Cá nhân không cư trú (20%, F="[x]") ngay dưới dòng Tổng để liên kết sang bảng đối chiếu sheet TK
        if sheet_type == "05-2":
            for offset, sub_label, crit in (
                (1, "Trong đó: Cá nhân cư trú (10%)", '"[]"'),
                (2, "Trong đó: Cá nhân không cư trú (20%)", '"[x]"'),
            ):
                sub_r = sum_row + offset
                ws.cell(sub_r, 1, sub_label).font = BOLD_FONT
                ws.cell(sub_r, 1).alignment = Alignment(horizontal="right", vertical="center")
                ws.merge_cells(f"A{sub_r}:F{sub_r}")
                for c in range(1, len(cols_def) + 1):
                    cell = ws.cell(sub_r, c)
                    cell.font = BOLD_FONT
                    cell.border = THIN_BORDER
                    if c >= 7:
                        cl = get_column_letter(c)
                        if len(rows) > 0:
                            cell.value = f"=SUMIF($F$4:$F${sum_row-1},{crit},{cl}4:{cl}{sum_row-1})"
                        else:
                            cell.value = 0
                        cell.number_format = "#,##0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")

    total_cols = len(cols_def)
    for c in range(1, total_cols + 1):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 16
    ws.freeze_panes = "A4"


def write_single_year_tncn_file(
    output_filepath: str,
    decl: TaxDeclaration,
) -> None:
    """Tạo 1 file Excel riêng cho quyết toán TNCN của 1 năm với đúng 4 sheet: TK, 05-1, 05-2, 05-3."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    is_tt156 = _is_tt156_decl(decl)

    # 1. Sheet 'TK'
    ws_tk = wb.create_sheet(title="TK")
    write_sheet_tncn_tk(ws_tk, decl)

    # Lấy các bảng kê từ XML
    rows_05_1: List[Dict[str, Any]] = []
    rows_05_2: List[Dict[str, Any]] = []
    rows_05_3: List[Dict[str, Any]] = []
    for pl_name, rows in decl.appendices.items():
        pl_upper = pl_name.upper()
        if "05_1" in pl_upper or "05-1" in pl_upper or "051" in pl_upper:
            rows_05_1.extend(rows)
        elif "05_2" in pl_upper or "05-2" in pl_upper or "052" in pl_upper:
            rows_05_2.extend(rows)
        elif "05_3" in pl_upper or "05-3" in pl_upper or "053" in pl_upper:
            rows_05_3.extend(rows)

    # 2. Sheet '05-1'
    ws_05_1 = wb.create_sheet(title="05-1")
    write_sheet_tncn_pl(ws_05_1, COLS_05_1, MERGES_05_1, rows_05_1, "05-1", is_tt156=is_tt156)

    # 3. Sheet '05-2'
    ws_05_2 = wb.create_sheet(title="05-2")
    write_sheet_tncn_pl(ws_05_2, COLS_05_2, MERGES_05_2, rows_05_2, "05-2", is_tt156=is_tt156)

    # 4. Sheet '05-3'
    ws_05_3 = wb.create_sheet(title="05-3")
    write_sheet_tncn_pl(ws_05_3, COLS_05_3, MERGES_05_3, rows_05_3, "05-3", is_tt156=is_tt156)

    os.makedirs(os.path.dirname(os.path.abspath(output_filepath)), exist_ok=True)
    wb.save(output_filepath)


def write_consolidated_tncn_appendices(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
) -> None:
    """
    Xuất 3 sheet phụ lục quyết toán TNCN chuẩn mẫu ('05-1', '05-2', '05-3') vào file Excel tổng hợp:
    - Nếu chỉ có 1 năm: Xuất đúng bảng chuẩn 1 năm.
    - Nếu có nhiều năm (VD: 2019..2024): Giữ nguyên cấu trúc cột chuẩn của mẫu (23 cột cho 05-1,
      12 cột cho 05-2, 11 cột cho 05-3), phân tách từng năm bằng dòng tiêu đề NĂM và dòng
      Tổng cộng năm {NĂM} kèm dòng TỔNG CỘNG TẤT CẢ CÁC NĂM ở cuối bảng.
    """
    sorted_decls = sorted(declarations, key=lambda d: (d.year or 0, d.tu_ngay or datetime.date.min))
    if not sorted_decls:
        return

    if len(sorted_decls) == 1:
        d = sorted_decls[0]
        is_tt156 = _is_tt156_decl(d)
        r_051: List[Dict[str, Any]] = []
        r_052: List[Dict[str, Any]] = []
        r_053: List[Dict[str, Any]] = []
        for pl_name, rows in d.appendices.items():
            pl_u = pl_name.upper()
            if "05_1" in pl_u or "05-1" in pl_u or "051" in pl_u:
                r_051.extend(rows)
            elif "05_2" in pl_u or "05-2" in pl_u or "052" in pl_u:
                r_052.extend(rows)
            elif "05_3" in pl_u or "05-3" in pl_u or "053" in pl_u:
                r_053.extend(rows)
        for s_type, cols_def, merges_def, r_list in (
            ("05-1", COLS_05_1, MERGES_05_1, r_051),
            ("05-2", COLS_05_2, MERGES_05_2, r_052),
            ("05-3", COLS_05_3, MERGES_05_3, r_053),
        ):
            if r_list:
                sname = sanitize_sheet_name(s_type, existing_sheets)
                ws = wb.create_sheet(title=sname)
                write_sheet_tncn_pl(ws, cols_def, merges_def, r_list, s_type, is_tt156=is_tt156)
        return

    # Trường hợp nhiều năm: gom từng loại phụ lục 05-1, 05-2, 05-3 theo từng năm
    for sheet_type, cols_def, merges_def in (
        ("05-1", COLS_05_1, MERGES_05_1),
        ("05-2", COLS_05_2, MERGES_05_2),
        ("05-3", COLS_05_3, MERGES_05_3),
    ):
        year_blocks: List[Tuple[TaxDeclaration, List[Dict[str, Any]]]] = []
        for d in sorted_decls:
            matched_rows: List[Dict[str, Any]] = []
            for pl_name, rows in d.appendices.items():
                pl_u = pl_name.upper()
                if sheet_type == "05-1" and ("05_1" in pl_u or "05-1" in pl_u or "051" in pl_u):
                    matched_rows.extend(rows)
                elif sheet_type == "05-2" and ("05_2" in pl_u or "05-2" in pl_u or "052" in pl_u):
                    matched_rows.extend(rows)
                elif sheet_type == "05-3" and ("05_3" in pl_u or "05-3" in pl_u or "053" in pl_u):
                    matched_rows.extend(rows)
            if matched_rows:
                year_blocks.append((d, matched_rows))

        if not year_blocks:
            continue

        sname = sanitize_sheet_name(sheet_type, existing_sheets)
        ws = wb.create_sheet(title=sname)

        # Ghi 3 dòng tiêu đề chuẩn của mẫu 05-1 / 05-2 / 05-3
        ws.row_dimensions[1].height = 36
        ws.row_dimensions[2].height = 48
        ws.row_dimensions[3].height = 20
        for col_idx, (r1, r2, r3) in enumerate(cols_def, start=1):
            ws.cell(1, col_idx, r1)
            ws.cell(2, col_idx, r2)
            ws.cell(3, col_idx, r3)
        for m_range in merges_def:
            ws.merge_cells(m_range)
        for r in range(1, 4):
            for c in range(1, len(cols_def) + 1):
                cell = ws.cell(r, c)
                cell.font = BOLD_FONT
                cell.fill = GRAY_HEADER_FILL
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        checkbox_fields = {"caNhanUyQuyen", "caNhanNuocNgoai", "caNhanKhongCuTru", "caNhanMienDuoi50k"}
        center_fields = checkbox_fields | {"stt", "quanHe", "tuThang", "denThang"}

        cur_row = 4
        year_sum_rows: List[int] = []
        total_cols = len(cols_def)

        for d, rows in year_blocks:
            is_tt156 = _is_tt156_decl(d)
            yr_label = d.year or d.ky_kkhai

            # Dòng banner phân tách từng năm
            ws.row_dimensions[cur_row].height = 22
            banner_text = (
                f"NĂM {yr_label} — Mã tờ khai: {d.ma_tkhai} | Lần kê khai: {d.so_lan or '0'} | "
                f"Số dòng: {len(rows)} | File: {d.file_name}"
            )
            ws.cell(cur_row, 1, banner_text)
            ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=total_cols)
            for c_i in range(1, total_cols + 1):
                bc = ws.cell(cur_row, c_i)
                bc.font = BOLD_FONT
                bc.fill = SECTION_FILL
                bc.border = THIN_BORDER
                if c_i == 1:
                    bc.alignment = Alignment(horizontal="left", vertical="center")
            cur_row += 1

            # Xác định ánh xạ cột theo đúng Thông tư 156/92 hoặc Thông tư 80 của năm đó
            if sheet_type == "05-1":
                if is_tt156:
                    col_candidates = [
                        ("stt", ("stt",)),
                        ("hoVaTen", ("ct07", "hoVaTen", "07")),
                        ("mst", ("ct08", "mst", "08")),
                        ("loaiGiayTo", ("ct09a_ten", "loaiGiayTo")),
                        ("soCMT", ("ct09", "soCMT", "09")),
                        ("caNhanUyQuyen", ("ct10", "caNhanUyQuyen", "10")),
                        ("caNhanNuocNgoai", ("caNhanNNUQQT", "caNhanNuocNgoai")),
                        ("tongTNCT", ("ct11", "tongTNCT", "11")),
                        ("tnctTaiToChuc", ("ct12", "tnctTaiToChuc", "12")),
                        ("tnctMienHiepDinh", ("ct13", "tnctMienHiepDinh", "13")),
                        ("tnctMienDauKhi", ("tnctMienDauKhi",)),
                        ("soLuongNPT", ("ct14", "soLuongNPT", "14")),
                        ("tongGiamTruGiaCanh", ("ct15", "tongGiamTruGiaCanh", "15")),
                        ("tuThien", ("ct16", "tuThien", "16")),
                        ("baoHiem", ("ct17", "baoHiem", "17")),
                        ("huuTri", ("ct18", "huuTri", "18")),
                        ("tntt", ("ct19", "tntt", "19")),
                        ("thueDaKhauTru", ("ct20", "thueDaKhauTru", "20")),
                        ("thueDaKhauTruTaiToChuc", ("ct21", "thueDaKhauTruTaiToChuc", "21")),
                        ("tongThuePhaiNop", ("ct22", "tongThuePhaiNop", "22")),
                        ("thueNopThua", ("ct23", "thueNopThua", "23")),
                        ("thueConPhaiNop", ("ct24", "thueConPhaiNop", "24")),
                        ("caNhanMienDuoi50k", ("caNhanMienDuoi50k",)),
                    ]
                else:
                    col_candidates = [
                        ("stt", ("stt",)),
                        ("hoVaTen", ("ct07", "hoVaTen", "07")),
                        ("mst", ("ct08", "mst", "08")),
                        ("loaiGiayTo", ("ct09a_ten", "loaiGiayTo")),
                        ("soCMT", ("ct09", "soCMT", "09")),
                        ("caNhanUyQuyen", ("ct10", "caNhanUyQuyen", "10")),
                        ("caNhanNuocNgoai", ("ct11", "caNhanNuocNgoai", "11")),
                        ("tongTNCT", ("ct12", "tongTNCT", "12")),
                        ("tnctTaiToChuc", ("ct13", "tnctTaiToChuc", "13")),
                        ("tnctMienHiepDinh", ("ct14", "tnctMienHiepDinh", "14")),
                        ("tnctMienDauKhi", ("ct15", "tnctMienDauKhi", "15")),
                        ("soLuongNPT", ("ct16", "soLuongNPT", "16")),
                        ("tongGiamTruGiaCanh", ("ct17", "tongGiamTruGiaCanh", "17")),
                        ("tuThien", ("ct18", "tuThien", "18")),
                        ("baoHiem", ("ct19", "baoHiem", "19")),
                        ("huuTri", ("ct20", "huuTri", "20")),
                        ("tntt", ("ct21", "tntt", "21")),
                        ("thueDaKhauTru", ("ct22", "thueDaKhauTru", "22")),
                        ("thueDaKhauTruTaiToChuc", ("ct23", "thueDaKhauTruTaiToChuc", "23")),
                        ("tongThuePhaiNop", ("ct24", "tongThuePhaiNop", "24")),
                        ("thueNopThua", ("ct25", "thueNopThua", "25")),
                        ("thueConPhaiNop", ("ct26", "thueConPhaiNop", "26")),
                        ("caNhanMienDuoi50k", ("ct27", "caNhanMienDuoi50k", "27")),
                    ]
            elif sheet_type == "05-2":
                col_candidates = [
                    ("stt", ("stt",)),
                    ("hoVaTen", ("ct07", "hoVaTen", "07")),
                    ("mst", ("ct08", "mst", "08")),
                    ("loaiGiayTo", ("ct09a_ten", "loaiGiayTo")),
                    ("soCMT", ("ct09", "soCMT", "09")),
                    ("caNhanKhongCuTru", ("ct10", "caNhanKhongCuTru", "10")),
                    ("tongTNCT", ("ct11", "tongTNCT", "11")),
                    ("tnctBaoHiem", ("ct12", "tnctBaoHiem", "12")),
                    ("tnctMienHiepDinh", ("ct13", "tnctMienHiepDinh", "13")),
                    ("tnctMienDauKhi", ("ct14", "tnctMienDauKhi", "14")),
                    ("thueDaKhauTru", ("ct15", "thueDaKhauTru", "15")),
                    ("thueBaoHiem", ("ct16", "thueBaoHiem", "16")),
                ]
            else:  # 05-3
                if is_tt156:
                    col_candidates = [
                        ("stt", ("stt",)),
                        ("hoVaTen", ("ct07", "hoVaTen", "07")),
                        ("mst", ("ct08", "mst", "08")),
                        ("hoVaTenNPT", ("ct09", "hoVaTenNPT", "09")),
                        ("ngaySinhNPT", ("ct10", "ngaySinhNPT", "10")),
                        ("mstNPT", ("ct11", "mstNPT", "11")),
                        ("loaiGiayTo", ("loaiGiayTo", "ct12_ten", "ct12")),
                        ("soGiayTo", ("ct13", "soGiayTo", "13")),
                        ("quanHe", ("ct14_ten", "quanHe", "ct14", "14")),
                        ("tuThang", ("ct21", "tuThang", "ct15", "15")),
                        ("denThang", ("ct22", "denThang", "ct16", "16")),
                    ]
                else:
                    col_candidates = [
                        ("stt", ("stt",)),
                        ("hoVaTen", ("ct07", "hoVaTen", "07")),
                        ("mst", ("ct08", "mst", "08")),
                        ("hoVaTenNPT", ("ct09", "hoVaTenNPT", "09")),
                        ("ngaySinhNPT", ("ct10", "ngaySinhNPT", "10")),
                        ("mstNPT", ("ct11", "mstNPT", "11")),
                        ("loaiGiayTo", ("ct12_ten", "loaiGiayTo", "ct12", "12")),
                        ("soGiayTo", ("ct13", "soGiayTo", "13")),
                        ("quanHe", ("ct14_ten", "quanHe", "ct14", "14")),
                        ("tuThang", ("ct15", "tuThang", "15")),
                        ("denThang", ("ct16", "denThang", "16")),
                    ]

            yr_start = cur_row
            for idx, r in enumerate(rows, start=1):
                row_is_156 = is_tt156 or ("caNhanNNUQQT" in r) or (sheet_type == "05-3" and ("ct21" in r or "ct22" in r))
                for col_idx, (field_name, cand_keys) in enumerate(col_candidates, start=1):
                    if field_name == "stt":
                        val: Any = idx
                    else:
                        val = _first_not_none(r, cand_keys)
                        if sheet_type == "05-3" and row_is_156:
                            if field_name == "soGiayTo" and (val is None or str(val).strip() == ""):
                                q_so = _first_not_none(r, ("ct15",))
                                s_gks = _first_not_none(r, ("ct16",))
                                if q_so is not None and s_gks is not None:
                                    val = f"QS: {q_so} - Số: {s_gks}"
                                elif q_so is not None:
                                    val = str(q_so)
                                elif s_gks is not None:
                                    val = str(s_gks)
                            elif field_name == "loaiGiayTo":
                                has_cmnd = _first_not_none(r, ("ct13",)) not in (None, "")
                                has_gks = _first_not_none(r, ("ct15", "ct16")) not in (None, "")
                                if not has_cmnd and has_gks:
                                    val = "Giấy khai sinh"

                    if field_name in checkbox_fields:
                        val = "[x]" if val in (1, "1", True, "true", "True", "[x]") else "[]"

                    is_mst_col = any(x in field_name.lower() for x in ("mst", "socmt", "sogiayto", "cccd"))
                    cell = ws.cell(cur_row, col_idx)
                    write_cell_formatted(cell, val, is_mst=is_mst_col)
                    if field_name in center_fields:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                cur_row += 1

            yr_end = cur_row - 1
            if sheet_type in ("05-1", "05-2"):
                sum_row = cur_row
                year_sum_rows.append(sum_row)
                ws.row_dimensions[sum_row].height = 22
                ws.cell(sum_row, 1, f"Tổng cộng năm {yr_label}").font = BOLD_FONT
                ws.cell(sum_row, 1).alignment = Alignment(horizontal="center", vertical="center")
                merge_end_col = 7 if sheet_type == "05-1" else 6
                ws.merge_cells(start_row=sum_row, start_column=1, end_row=sum_row, end_column=merge_end_col)

                start_col = 8 if sheet_type == "05-1" else 7
                for c in range(1, total_cols + 1):
                    cell = ws.cell(sum_row, c)
                    cell.font = BOLD_FONT
                    cell.fill = SECTION_FILL
                    cell.border = THIN_BORDER
                    if c >= start_col and c != (23 if sheet_type == "05-1" else 99):
                        cl = get_column_letter(c)
                        cell.value = f"=SUM({cl}{yr_start}:{cl}{yr_end})"
                        cell.number_format = "#,##0"
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                cur_row += 1

                if sheet_type == "05-2":
                    for sub_label, crit in (
                        (f"Trong đó: Cá nhân cư trú (10%) năm {yr_label}", '"[]"'),
                        (f"Trong đó: Cá nhân không cư trú (20%) năm {yr_label}", '"[x]"'),
                    ):
                        sub_r = cur_row
                        ws.row_dimensions[sub_r].height = 20
                        ws.cell(sub_r, 1, sub_label).font = BOLD_FONT
                        ws.cell(sub_r, 1).alignment = Alignment(horizontal="right", vertical="center")
                        ws.merge_cells(start_row=sub_r, start_column=1, end_row=sub_r, end_column=6)
                        for c in range(1, total_cols + 1):
                            cell = ws.cell(sub_r, c)
                            cell.font = BOLD_FONT
                            cell.border = THIN_BORDER
                            if c >= 7:
                                cl = get_column_letter(c)
                                cell.value = f"=SUMIF($F${yr_start}:$F${yr_end},{crit},{cl}{yr_start}:{cl}{yr_end})"
                                cell.number_format = "#,##0"
                                cell.alignment = Alignment(horizontal="right", vertical="center")
                        cur_row += 1

        if sheet_type in ("05-1", "05-2") and len(year_sum_rows) > 1:
            grand_row = cur_row
            ws.row_dimensions[grand_row].height = 24
            ws.cell(grand_row, 1, "TỔNG CỘNG TẤT CẢ CÁC NĂM").font = BOLD_FONT
            ws.cell(grand_row, 1).alignment = Alignment(horizontal="center", vertical="center")
            merge_end_col = 7 if sheet_type == "05-1" else 6
            ws.merge_cells(start_row=grand_row, start_column=1, end_row=grand_row, end_column=merge_end_col)

            start_col = 8 if sheet_type == "05-1" else 7
            for c in range(1, total_cols + 1):
                cell = ws.cell(grand_row, c)
                cell.font = BOLD_FONT
                cell.fill = GRAY_HEADER_FILL
                cell.border = THIN_BORDER
                if c >= start_col and c != (23 if sheet_type == "05-1" else 99):
                    cl = get_column_letter(c)
                    sum_expr = "+".join(f"{cl}{r}" for r in year_sum_rows)
                    cell.value = f"={sum_expr}"
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        for c in range(1, total_cols + 1):
            ws.column_dimensions[get_column_letter(c)].width = 15
        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["C"].width = 16
        ws.freeze_panes = "A4"


# ==============================================================================
# 5. CÁC SHEET GTGT: 'So Thue GTGT TH' (Toàn bộ tờ khai) & 'So Thue GTGT Cuoi' (Chốt theo năm)
# ==============================================================================
GTGT_COLS_SPEC = [
    ("STT", "#", 6),
    ("Mã CQT phát sinh", "(A)", 10),
    ("Tên CQT phát sinh", "(B)", 26),
    ("Tên CQT cấp Cục", "(C)", 26),
    ("Tên NNT", "(E)", 30),
    ("Trạng thái hoạt động", "(F)", 8),
    ("Kỳ tính thuế", "(K)", 10),
    ("Lần kê khai", "(M)", 10),
    ("Ngày ghi sổ", "(N)", 12),
    ("Ngày nộp tờ khai", "(O)", 12),
    ("Tiểu mục", "(1)", 10),
    ("Thuế GTGT còn được khấu trừ kỳ trước chuyển sang (22)", "(2)", 18),
    ("Giá trị HHDV mua vào trong kỳ (23)", "(3)", 18),
    ("Thuế GTGT HHDV mua vào trong kỳ (24)", "(4)", 18),
    ("Giá trị HHDV mua vào nhập khẩu (23a)", "(5)", 18),
    ("Thuế GTGT HHDV mua vào nhập khẩu (24a)", "(6)", 18),
    ("Thuế đầu vào được khấu trừ (25)", "(7)", 18),
    ("Doanh thu HHDV bán ra không chịu thuế GTGT (26)", "(8)", 18),
    ("Doanh thu HHDV bán ra chịu thuế GTGT (27)", "(9)", 18),
    ("Thuế của HHDV bán ra chịu thuế GTGT (28)", "(10)", 18),
    ("Doanh thu HHDV thuế suất 0% (29)", "(11)", 18),
    ("Doanh thu HHDV thuế suất 5% (30)", "(12)", 18),
    ("Thuế của HHDV thuế suất 5% (31)", "(13)", 18),
    ("Doanh thu HHDV thuế suất 10% (32)", "(14)", 18),
    ("Thuế của HHDV thuế suất 10% (33)", "(15)", 18),
    ("Doanh thu HHDV không tính thuế (32a)", "(16)", 18),
    ("Tổng doanh thu (34)", "(17)", 18),
    ("Thuế đầu ra (35)", "(18)", 18),
    ("Thuế GTGT phát sinh trong kỳ (36)", "(19)", 18),
    ("Điều chỉnh giảm thuế GTGT các kỳ trước (37)", "(20)", 18),
    ("Điều chỉnh tăng thuế GTGT các kỳ trước (38)", "(21)", 18),
    ("Thuế GTGT đã nộp ở địa phương khác của HĐKD xây dựng, lắp đặt, bán hàng, bất động sản ngoại tỉnh (39)", "(22)", 18),
    ("Thuế GTGT nhận bàn giao được khấu trừ trong kỳ (39a)", "(23)", 18),
    ("Thuế GTGT phải nộp của HĐ SXKD trong kỳ (40a)", "(24)", 18),
    ("Thuế GTGT bù trừ với dự án đầu tư (40b)", "(25)", 18),
    ("Thuế GTGT còn phải nộp (40)", "(26)", 18),
    ("Thuế GTGT chưa khấu trừ hết (41)", "(27)", 18),
    ("Thuế GTGT đề nghị hoàn (42)", "(28)", 18),
    ("Thuế GTGT khấu trừ chuyển kỳ sau (43)", "(29)", 18),
]

GTGT_INDICATOR_MAP = {
    12: ("ct22", "22"),
    13: ("ct23", "23"),
    14: ("ct24", "24"),
    15: ("ct23a", "23a"),
    16: ("ct24a", "24a"),
    17: ("ct25", "25"),
    18: ("ct26", "26"),
    19: ("ct27", "27"),
    20: ("ct28", "28"),
    21: ("ct29", "29"),
    22: ("ct30", "30"),
    23: ("ct31", "31"),
    24: ("ct32", "32"),
    25: ("ct33", "33"),
    26: ("ct32a", "32a"),
    27: ("ct34", "34"),
    28: ("ct35", "35"),
    29: ("ct36", "36"),
    30: ("ct37", "37"),
    31: ("ct38", "38"),
    32: ("ct39", "39"),
    33: ("ct39a", "39a"),
    34: ("ct40a", "40a"),
    35: ("ct40b", "40b"),
    36: ("ct40", "40"),
    37: ("ct41", "41"),
    38: ("ct42", "42"),
    39: ("ct43", "43"),
}


def write_gtgt_timeline_sheet(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    sheet_name: str = "So Thue GTGT TH",
) -> None:
    """Sheet 'So Thue GTGT TH': Tất cả các tờ khai GTGT (lần đầu + các lần bổ sung theo dòng thời gian)."""
    sname = sanitize_sheet_name(sheet_name, existing_sheets)
    ws = wb.create_sheet(title=sname)

    ws.append([c[0] for c in GTGT_COLS_SPEC])
    ws.append([c[1] for c in GTGT_COLS_SPEC])

    for r in (1, 2):
        ws.row_dimensions[r].height = 26
        for c in range(1, len(GTGT_COLS_SPEC) + 1):
            cell = ws.cell(r, c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

    def gtgt_sort_key(d: TaxDeclaration):
        y = d.year or 0
        tu = d.tu_ngay or datetime.date.min
        try:
            lan = int(d.so_lan) if d.so_lan else 0
        except ValueError:
            lan = 0
        nlap = d.ngay_lap or datetime.date.min
        return (d.mst or "", y, tu, lan, nlap)

    unique_gtgt = deduplicate_gtgt_declarations(declarations)
    sorted_gtgt = sorted(unique_gtgt, key=gtgt_sort_key)

    cur_row = 3
    for idx, d in enumerate(sorted_gtgt, start=1):
        try:
            lan_num = int(d.so_lan) if d.so_lan else 0
        except ValueError:
            lan_num = 0

        ky_str = d.ky_kkhai
        if d.tu_ngay:
            ky_str = f"{d.tu_ngay.strftime('%y%m')}"

        write_cell_formatted(ws.cell(cur_row, 1), idx)
        write_cell_formatted(ws.cell(cur_row, 2), d.ma_cqt or "", is_mst=True)
        write_cell_formatted(ws.cell(cur_row, 3), d.ten_cqt)
        write_cell_formatted(ws.cell(cur_row, 4), "")
        write_cell_formatted(ws.cell(cur_row, 5), d.ten_nnt)
        write_cell_formatted(ws.cell(cur_row, 6), "00", is_mst=True)
        write_cell_formatted(ws.cell(cur_row, 7), ky_str, is_mst=True)
        write_cell_formatted(ws.cell(cur_row, 8), lan_num)
        write_cell_formatted(ws.cell(cur_row, 9), d.ngay_lap)
        write_cell_formatted(ws.cell(cur_row, 10), d.ngay_lap)
        write_cell_formatted(ws.cell(cur_row, 11), d.tieu_muc or "1701", is_mst=True)

        for col_idx, (tag1, tag2) in GTGT_INDICATOR_MAP.items():
            val = d.indicators.get(tag1)
            if val is None:
                val = d.indicators.get(tag2)
            if val is None:
                val = d.indicators.get(tag1.upper())
            write_cell_formatted(ws.cell(cur_row, col_idx), val if val is not None else 0)

        cur_row += 1

    total_rows = cur_row - 1
    total_cols = len(GTGT_COLS_SPEC)
    for c_idx, spec in enumerate(GTGT_COLS_SPEC, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = spec[2]
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}{total_rows}"
    ws.freeze_panes = "H3"


def detect_tndn_fiscal_ranges(tndn_decls: List[TaxDeclaration]) -> Dict[int, Tuple[datetime.date, datetime.date]]:
    """
    Tự động phát hiện doanh nghiệp có năm tài chính lệch năm dương lịch từ tờ khai 03/TNDN
    (VD: 01/04/2019 - 31/03/2020 cho năm tính thuế 2019).
    Trả về dict {year: (tu_ngay, den_ngay)} nếu có lệch kỳ; ngược lại trả về {}.
    """
    ranges: Dict[int, Tuple[datetime.date, datetime.date]] = {}
    has_offset = False
    for d in tndn_decls:
        if d.year and d.tu_ngay and d.den_ngay:
            ranges[d.year] = (d.tu_ngay, d.den_ngay)
            if d.tu_ngay.month != 1 or d.den_ngay.month != 12 or d.tu_ngay.year != d.den_ngay.year:
                has_offset = True
    return ranges if has_offset else {}


def write_gtgt_cuoi_sheet(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    sheet_name: str = "So Thue GTGT Cuoi",
    fiscal_ranges: Optional[Dict[int, Tuple[datetime.date, datetime.date]]] = None,
) -> None:
    """
    Sheet 'So Thue GTGT Cuoi' (hoặc 'So Thue GTGT Cuoi (Lech ky)' khi có fiscal_ranges):
    - Với mỗi kỳ kê khai thuế (tháng/quý), chỉ lấy tờ khai CUỐI CÙNG có hiệu lực (lần bổ sung cao nhất).
    - Gom nhóm các kỳ theo từng NĂM (hoặc theo NĂM TÀI CHÍNH lệch kỳ của TNDN nếu fiscal_ranges được cung cấp).
    - Cuối mỗi năm, có dòng TỔNG CỘNG NĂM {NĂM} dùng công thức =SUM(...) tính tổng số liệu cả năm.
    - Dưới cùng có dòng TỔNG CỘNG TẤT CẢ CÁC NĂM.
    """
    sname = sanitize_sheet_name(sheet_name, existing_sheets)
    ws = wb.create_sheet(title=sname)

    unique_gtgt = deduplicate_gtgt_declarations(declarations)

    grouped_period: Dict[Tuple[str, str], List[TaxDeclaration]] = {}
    for d in unique_gtgt:
        mst = (d.mst or "").strip()
        if d.tu_ngay:
            ky_str = d.tu_ngay.strftime("%y%m")
        else:
            ky_str = (d.ky_kkhai or "").strip()
        grouped_period.setdefault((mst, ky_str), []).append(d)

    latest_per_period: List[TaxDeclaration] = []
    for k, dlist in grouped_period.items():
        def sort_key_p(item: TaxDeclaration):
            try:
                lan = int(item.so_lan) if item.so_lan is not None and str(item.so_lan).strip() != "" else 0
            except (ValueError, TypeError):
                lan = 0
            nlap = item.ngay_lap or datetime.date.min
            valid_inds = sum(1 for v in item.indicators.values() if v not in (None, 0, "", "0"))
            return (lan, nlap, valid_inds, len(item.file_name), item.file_name)

        dlist.sort(key=sort_key_p)
        latest_per_period.append(dlist[-1])

    # Nếu có năm tài chính lệch kỳ, xác định tháng bắt đầu phổ biến (VD: tháng 4) để gom các kỳ nằm ngoài khoảng TNDN
    default_start_month = 1
    if fiscal_ranges:
        start_months = [r_start.month for r_start, _ in fiscal_ranges.values()]
        if start_months:
            default_start_month = max(set(start_months), key=start_months.count)

    def resolve_group_year(d: TaxDeclaration) -> int:
        if not fiscal_ranges:
            return d.year or 0
        ref_date = d.tu_ngay or d.den_ngay
        if ref_date:
            for f_yr, (f_start, f_end) in fiscal_ranges.items():
                if f_start <= ref_date <= f_end:
                    return f_yr
            # Suy luận theo tháng bắt đầu năm tài chính (VD tháng 4: 04/2019..03/2020 -> năm 2019)
            if default_start_month > 1:
                return ref_date.year if ref_date.month >= default_start_month else (ref_date.year - 1)
        return d.year or 0

    grouped_by_year: Dict[int, List[TaxDeclaration]] = {}
    for d in latest_per_period:
        yr = resolve_group_year(d)
        grouped_by_year.setdefault(yr, []).append(d)

    ws.append([c[0] for c in GTGT_COLS_SPEC])
    ws.append([c[1] for c in GTGT_COLS_SPEC])

    for r in (1, 2):
        ws.row_dimensions[r].height = 26
        for c in range(1, len(GTGT_COLS_SPEC) + 1):
            cell = ws.cell(r, c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER

    cur_row = 3
    global_idx = 1
    year_sum_rows: List[int] = []

    for yr in sorted(grouped_by_year.keys()):
        decls_in_yr = grouped_by_year[yr]
        decls_in_yr.sort(key=lambda d: d.tu_ngay or datetime.date.min)

        yr_start = cur_row
        for d in decls_in_yr:
            try:
                lan_num = int(d.so_lan) if d.so_lan is not None and str(d.so_lan).strip() != "" else 0
            except (ValueError, TypeError):
                lan_num = 0

            ky_str = d.ky_kkhai
            if d.tu_ngay:
                ky_str = f"{d.tu_ngay.strftime('%y%m')}"

            write_cell_formatted(ws.cell(cur_row, 1), global_idx)
            write_cell_formatted(ws.cell(cur_row, 2), d.ma_cqt or "", is_mst=True)
            write_cell_formatted(ws.cell(cur_row, 3), d.ten_cqt)
            write_cell_formatted(ws.cell(cur_row, 4), "")
            write_cell_formatted(ws.cell(cur_row, 5), d.ten_nnt)
            write_cell_formatted(ws.cell(cur_row, 6), "00", is_mst=True)
            write_cell_formatted(ws.cell(cur_row, 7), ky_str, is_mst=True)
            write_cell_formatted(ws.cell(cur_row, 8), lan_num)
            write_cell_formatted(ws.cell(cur_row, 9), d.ngay_lap)
            write_cell_formatted(ws.cell(cur_row, 10), d.ngay_lap)
            write_cell_formatted(ws.cell(cur_row, 11), d.tieu_muc or "1701", is_mst=True)

            for col_idx, (tag1, tag2) in GTGT_INDICATOR_MAP.items():
                val = d.indicators.get(tag1)
                if val is None:
                    val = d.indicators.get(tag2)
                if val is None:
                    val = d.indicators.get(tag1.upper())
                write_cell_formatted(ws.cell(cur_row, col_idx), val if val is not None else 0)

            cur_row += 1
            global_idx += 1

        yr_end = cur_row - 1

        # DÒNG TỔNG CỘNG NĂM
        sum_row = cur_row
        year_sum_rows.append(sum_row)
        ws.row_dimensions[sum_row].height = 24

        for col_i in range(1, len(GTGT_COLS_SPEC) + 1):
            cell = ws.cell(sum_row, col_i)
            cell.font = BOLD_FONT
            cell.fill = SECTION_FILL
            cell.border = THIN_BORDER

        if fiscal_ranges and yr in fiscal_ranges:
            f_s, f_e = fiscal_ranges[yr]
            yr_label = f"Tổng cộng năm {yr} ({f_s.strftime('%m/%Y')}-{f_e.strftime('%m/%Y')})"
        else:
            yr_label = f"Tổng cộng năm {yr}"
        ws.cell(sum_row, 5, yr_label).alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(sum_row, 7, f"Tổng {yr}").alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in range(13, len(GTGT_COLS_SPEC) + 1):
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(sum_row, col_idx)
            cell.value = f"=SUM({col_letter}{yr_start}:{col_letter}{yr_end})"
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")

        cur_row += 1

    # DÒNG TỔNG CỘNG TẤT CẢ CÁC NĂM (GRAND TOTAL)
    if len(year_sum_rows) > 1:
        grand_row = cur_row
        ws.row_dimensions[grand_row].height = 26
        for col_i in range(1, len(GTGT_COLS_SPEC) + 1):
            cell = ws.cell(grand_row, col_i)
            cell.font = BOLD_FONT
            cell.fill = GRAY_HEADER_FILL
            cell.border = THIN_BORDER

        ws.cell(grand_row, 5, "TỔNG CỘNG TẤT CẢ CÁC NĂM").alignment = Alignment(horizontal="left", vertical="center")
        ws.cell(grand_row, 7, "TỔNG CỘNG").alignment = Alignment(horizontal="center", vertical="center")

        for col_idx in range(13, len(GTGT_COLS_SPEC) + 1):
            col_letter = get_column_letter(col_idx)
            cell = ws.cell(grand_row, col_idx)
            sum_terms = "+".join([f"{col_letter}{r}" for r in year_sum_rows])
            cell.value = f"={sum_terms}"
            cell.number_format = "#,##0"
            cell.alignment = Alignment(horizontal="right", vertical="center")

        cur_row += 1

    total_rows = cur_row - 1
    total_cols = len(GTGT_COLS_SPEC)
    for c_idx, spec in enumerate(GTGT_COLS_SPEC, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = spec[2]
    ws.auto_filter.ref = f"A2:{get_column_letter(total_cols)}{total_rows}"
    ws.freeze_panes = "H3"


def write_sheet_chenh_lech_gtgt_tndn(
    wb: openpyxl.Workbook,
    existing_sheets: Set[str],
    gtgt_sheet_name: str = "So Thue GTGT Cuoi",
    kqkd_sheet_name: str = "BCHDKD",
    sheet_name: str = "Chenh lech thue GTGT - TNDN",
    fiscal_ranges: Optional[Dict[int, Tuple[datetime.date, datetime.date]]] = None,
) -> Optional[str]:
    """
    Tạo sheet bảng so sánh giữa doanh thu tính thuế GTGT hàng năm (lấy từ sheet 'So Thue GTGT Cuoi'
    hoặc 'So Thue GTGT Cuoi (Lech ky)' nếu doanh nghiệp có năm tài chính lệch năm dương lịch)
    và doanh thu tính thuế TNDN trên báo cáo hoạt động sản xuất kinh doanh ('BCHDKD').
    Tự động dò tọa độ dòng của từng chỉ tiêu trên từng khối cột của 'BCHDKD' để không bao giờ lệch dòng.
    """
    actual_gtgt_sheet = None
    for s in wb.sheetnames:
        if s.strip().lower() == gtgt_sheet_name.strip().lower():
            actual_gtgt_sheet = s
            break
    if actual_gtgt_sheet is None:
        for s in wb.sheetnames:
            if "so thue gtgt cuoi" in s.lower():
                actual_gtgt_sheet = s
                break

    actual_kqkd_sheet = None
    for s in wb.sheetnames:
        if s.strip().lower() == kqkd_sheet_name.strip().lower() or "bchdkd" in s.lower():
            actual_kqkd_sheet = s
            break

    gtgt_year_sum_rows: Dict[int, int] = {}
    if actual_gtgt_sheet and actual_gtgt_sheet in wb.sheetnames:
        ws_gt = wb[actual_gtgt_sheet]
        for r in range(1, ws_gt.max_row + 1):
            val_e = ws_gt.cell(r, 5).value
            if val_e and "Tổng cộng năm" in str(val_e):
                m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_e))
                if m:
                    gtgt_year_sum_rows[int(m.group(0))] = r

    # Dò tọa độ cột và tọa độ dòng thực tế cho từng năm trên sheet BCHDKD
    kqkd_year_cols: Dict[int, str] = {}
    kqkd_year_rows: Dict[int, Tuple[int, int, int]] = {}  # year -> (row_dt, row_giam_tru, row_tn_khac)
    if actual_kqkd_sheet and actual_kqkd_sheet in wb.sheetnames:
        ws_kq = wb[actual_kqkd_sheet]
        # Tìm các cột "Mã chỉ tiêu" trong từng khối bảng trên BCHDKD
        code_cols: List[int] = []
        for c in range(1, ws_kq.max_column + 1):
            if str(ws_kq.cell(2, c).value or "").strip() == "Mã chỉ tiêu":
                code_cols.append(c)

        for c in range(1, ws_kq.max_column + 1):
            val_r1 = ws_kq.cell(1, c).value
            if val_r1:
                m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_r1))
                if m:
                    yr_num = int(m.group(0))
                    kqkd_year_cols[yr_num] = get_column_letter(c)
                    # Tìm cột Mã chỉ tiêu gần nhất bên trái cột năm c
                    left_code_cols = [cc for cc in code_cols if cc < c]
                    cc = max(left_code_cols) if left_code_cols else 4
                    # Xác định khối cột này thuộc TT156 (dòng 5 là [01]) hay TT80 (dòng 5 là [04])
                    first_code = str(ws_kq.cell(5, cc).value or "").strip()
                    if first_code == "[01]":
                        dt_code, gt_code, tnk_code = "[01]", "[03]", "[16]"
                    else:
                        dt_code, gt_code, tnk_code = "[04]", "[06]", "[19]"
                    r_dt, r_gt, r_tnk = 5, 7, 20
                    for r_scan in range(5, ws_kq.max_row + 1):
                        c_val = str(ws_kq.cell(r_scan, cc).value or "").strip()
                        if c_val == dt_code:
                            r_dt = r_scan
                        elif c_val == gt_code:
                            r_gt = r_scan
                        elif c_val == tnk_code:
                            r_tnk = r_scan
                    kqkd_year_rows[yr_num] = (r_dt, r_gt, r_tnk)

    all_years = sorted(set(gtgt_year_sum_rows.keys()) | set(kqkd_year_cols.keys()))
    if not all_years:
        return None

    sname = sanitize_sheet_name(sheet_name, existing_sheets)
    ws = wb.create_sheet(title=sname)

    r1_vals = [None, 1, 2, 3, None, 4, 5, "6=2-3", "7=6+4", "8=1-7"]
    r2_vals = [
        "Năm",
        "DT tính thuế GTGT",
        "DT tính thuế TNDN",
        "Giảm trừ doanh thu",
        "Chênh lệch chưa bao gồm thu nhập khác",
        "Thu nhập khác",
        "Chênh lệch",
        "DT tính thuế TNDN-Giảm trừ doanh thu",
        "(DT tính thuế TNDN-Giảm trừ doanh thu)+Thu nhập khác",
        "Chênh lệch",
    ]

    ws.append(r1_vals)
    ws.append(r2_vals)

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 42

    h_font = Font(name="Times New Roman", size=11, bold=True)
    r1_font = Font(name="Times New Roman", size=11, bold=False)
    body_font = Font(name="Times New Roman", size=11, bold=False)
    bold_font = Font(name="Times New Roman", size=11, bold=True)

    for c in range(1, 11):
        c1 = ws.cell(1, c)
        c1.font = r1_font
        if c1.value is not None:
            c1.alignment = Alignment(horizontal="center", vertical="center")

        c2 = ws.cell(2, c)
        c2.font = h_font
        c2.border = THIN_BORDER
        c2.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    cur_row = 3
    for yr in all_years:
        ws.row_dimensions[cur_row].height = 22
        r_dt, r_gt, r_tnk = kqkd_year_rows.get(yr, (5, 7, 20))

        # Cột A: Năm (hiển thị kèm kỳ tài chính nếu doanh nghiệp lệch kỳ, VD: 2019 (04/2019-03/2020))
        if fiscal_ranges and yr in fiscal_ranges:
            f_s, f_e = fiscal_ranges[yr]
            col_a_val: Any = f"{yr} ({f_s.strftime('%m/%Y')}-{f_e.strftime('%m/%Y')})"
        else:
            col_a_val = yr
        cA = ws.cell(cur_row, 1, col_a_val)
        cA.font = body_font
        cA.alignment = Alignment(horizontal="center", vertical="center")
        cA.border = THIN_BORDER

        # Cột B: DT tính thuế GTGT (1)
        cB = ws.cell(cur_row, 2)
        if yr in gtgt_year_sum_rows and actual_gtgt_sheet:
            cB.value = f"='{actual_gtgt_sheet}'!AA{gtgt_year_sum_rows[yr]}"
        else:
            cB.value = 0
        cB.font = body_font
        cB.number_format = "#,##0"
        cB.alignment = Alignment(horizontal="right", vertical="center")
        cB.border = THIN_BORDER

        # Cột C: DT tính thuế TNDN (2)
        cC = ws.cell(cur_row, 3)
        if yr in kqkd_year_cols and actual_kqkd_sheet:
            cC.value = f"='{actual_kqkd_sheet}'!{kqkd_year_cols[yr]}{r_dt}"
        else:
            cC.value = 0
        cC.font = body_font
        cC.number_format = "#,##0"
        cC.alignment = Alignment(horizontal="right", vertical="center")
        cC.border = THIN_BORDER

        # Cột D: Giảm trừ doanh thu (3)
        cD = ws.cell(cur_row, 4)
        if yr in kqkd_year_cols and actual_kqkd_sheet:
            cD.value = f"='{actual_kqkd_sheet}'!{kqkd_year_cols[yr]}{r_gt}"
        else:
            cD.value = 0
        cD.font = body_font
        cD.number_format = "#,##0"
        cD.alignment = Alignment(horizontal="right", vertical="center")
        cD.border = THIN_BORDER

        # Cột E: Chênh lệch chưa bao gồm thu nhập khác = B - C
        cE = ws.cell(cur_row, 5, f"=B{cur_row}-C{cur_row}")
        cE.font = body_font
        cE.number_format = "#,##0"
        cE.alignment = Alignment(horizontal="right", vertical="center")
        cE.border = THIN_BORDER

        # Cột F: Thu nhập khác (4)
        cF = ws.cell(cur_row, 6)
        if yr in kqkd_year_cols and actual_kqkd_sheet:
            cF.value = f"='{actual_kqkd_sheet}'!{kqkd_year_cols[yr]}{r_tnk}"
        else:
            cF.value = 0
        cF.font = body_font
        cF.number_format = "#,##0"
        cF.alignment = Alignment(horizontal="right", vertical="center")
        cF.border = THIN_BORDER

        # Cột G: Chênh lệch (5) = B - C - F + D
        cG = ws.cell(cur_row, 7, f"=B{cur_row}-C{cur_row}-F{cur_row}+D{cur_row}")
        cG.font = body_font
        cG.number_format = "#,##0"
        cG.alignment = Alignment(horizontal="right", vertical="center")
        cG.border = THIN_BORDER

        # Cột H: DT tính thuế TNDN-Giảm trừ doanh thu (6=2-3) = C - D
        cH = ws.cell(cur_row, 8, f"=C{cur_row}-D{cur_row}")
        cH.font = body_font
        cH.number_format = "#,##0"
        cH.alignment = Alignment(horizontal="right", vertical="center")
        cH.border = THIN_BORDER

        # Cột I: (DT tính thuế TNDN-Giảm trừ doanh thu)+Thu nhập khác (7=6+4) = H + F
        cI = ws.cell(cur_row, 9, f"=H{cur_row}+F{cur_row}")
        cI.font = body_font
        cI.number_format = "#,##0"
        cI.alignment = Alignment(horizontal="right", vertical="center")
        cI.border = THIN_BORDER

        # Cột J: Chênh lệch (8=1-7) = B - I
        cJ = ws.cell(cur_row, 10, f"=B{cur_row}-I{cur_row}")
        cJ.font = body_font
        cJ.number_format = "#,##0"
        cJ.alignment = Alignment(horizontal="right", vertical="center")
        cJ.border = THIN_BORDER

        cur_row += 1

    # Dòng TỔNG CỘNG
    sum_row = cur_row
    ws.row_dimensions[sum_row].height = 24
    c_tot = ws.cell(sum_row, 1, "Tổng cộng")
    c_tot.font = bold_font
    c_tot.alignment = Alignment(horizontal="center", vertical="center")
    c_tot.border = THIN_BORDER

    for col_idx in range(2, 11):
        col_let = get_column_letter(col_idx)
        c_sum = ws.cell(sum_row, col_idx, f"=SUM({col_let}3:{col_let}{sum_row - 1})")
        c_sum.font = bold_font
        c_sum.number_format = "#,##0"
        c_sum.alignment = Alignment(horizontal="right", vertical="center")
        c_sum.border = THIN_BORDER

    col_widths = {
        "A": 24 if fiscal_ranges else 10,
        "B": 24,
        "C": 20,
        "D": 22,
        "E": 26,
        "F": 18,
        "G": 20,
        "H": 24,
        "I": 28,
        "J": 20,
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

    return sname


def write_yearly_verification_sheets(
    wb: openpyxl.Workbook,
    existing_sheets: Set[str],
    gtgt_sheet_name: str = "So Thue GTGT Cuoi",
    kqkd_sheet_name: str = "BCHDKD",
    kktndn_sheet_name: str = "KKTNDN",
    tncn_sheet_name: str = "TNCN",
) -> List[str]:
    """
    Tạo các sheet kiểm tra thuế theo từng năm (mỗi năm 1 sheet: '2019', '2020', ..., '2025', ...)
    chuẩn 98 dòng x 5 cột theo đúng mẫu biên bản kiểm tra thuế (Mau_kk.xlsx):
    - Bảng 1 (Dòng 1..25): Thuế GTGT (liên kết công thức trực tiếp tới sheet 'So Thue GTGT Cuoi')
    - Bảng 2 (Dòng 27..79): Thuế TNDN (liên kết công thức tới sheet 'BCHDKD' và 'KKTNDN', tự động
      nhận diện tọa độ dòng của cả mẫu TT156 và TT80)
    - Bảng 3 (Dòng 81..90): Thuế TNCN (liên kết công thức tới bảng tổng hợp 05-1/05-2 tại dòng 41..46 sheet 'TNCN')
    - Bảng 4 (Dòng 93..98): Thuế Nhà thầu
    """
    actual_gtgt_sheet = None
    actual_kqkd_sheet = None
    actual_kktndn_sheet = None
    actual_tncn_sheet = None

    for s in wb.sheetnames:
        sl = s.strip().lower()
        if sl == gtgt_sheet_name.strip().lower() or "so thue gtgt cuoi" in sl:
            actual_gtgt_sheet = s
        elif sl == kqkd_sheet_name.strip().lower() or "bchdkd" in sl:
            actual_kqkd_sheet = s
        elif sl == kktndn_sheet_name.strip().lower() or "kktndn" in sl:
            actual_kktndn_sheet = s
        elif sl == tncn_sheet_name.strip().lower() or sl == "tk":
            if actual_tncn_sheet is None or sl == tncn_sheet_name.strip().lower():
                actual_tncn_sheet = s

    # 1. Dò tọa độ dòng (r_first, r_last, r_sum) của từng năm trên sheet 'So Thue GTGT Cuoi'
    gtgt_coords: Dict[int, Tuple[int, int, int]] = {}
    if actual_gtgt_sheet and actual_gtgt_sheet in wb.sheetnames:
        ws_gt = wb[actual_gtgt_sheet]
        cur_first = 3
        for r in range(3, ws_gt.max_row + 1):
            val_e = str(ws_gt.cell(r, 5).value or "")
            if "Tổng cộng năm" in val_e:
                m = re.search(r"\b(19\d\d|20\d\d)\b", val_e)
                if m:
                    yr_num = int(m.group(0))
                    r_last = max(cur_first, r - 1)
                    gtgt_coords[yr_num] = (cur_first, r_last, r)
                cur_first = r + 1

    # 2. Dò tọa độ cột và bản đồ mã chỉ tiêu -> dòng của từng năm trên sheet 'BCHDKD'
    kqkd_coords: Dict[int, Tuple[str, Dict[str, int]]] = {}
    if actual_kqkd_sheet and actual_kqkd_sheet in wb.sheetnames:
        ws_kq = wb[actual_kqkd_sheet]
        code_cols = [
            c for c in range(1, ws_kq.max_column + 1)
            if str(ws_kq.cell(2, c).value or "").strip() == "Mã chỉ tiêu"
        ]
        for c in range(1, ws_kq.max_column + 1):
            val_r1 = ws_kq.cell(1, c).value
            if val_r1:
                m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_r1))
                if m:
                    yr_num = int(m.group(0))
                    left_cc = [cc for cc in code_cols if cc < c]
                    cc = max(left_cc) if left_cc else 4
                    cmap: Dict[str, int] = {}
                    for r_scan in range(4, ws_kq.max_row + 1):
                        c_val = str(ws_kq.cell(r_scan, cc).value or "").strip()
                        if c_val and c_val not in cmap:
                            cmap[c_val] = r_scan
                    kqkd_coords[yr_num] = (get_column_letter(c), cmap)

    # 3. Dò tọa độ cột và bản đồ mã chỉ tiêu -> dòng của từng năm trên sheet 'KKTNDN'
    kktndn_coords: Dict[int, Tuple[str, Dict[str, int]]] = {}
    if actual_kktndn_sheet and actual_kktndn_sheet in wb.sheetnames:
        ws_kk = wb[actual_kktndn_sheet]
        code_cols = [
            c for c in range(1, ws_kk.max_column + 1)
            if str(ws_kk.cell(2, c).value or "").strip() == "Mã chỉ tiêu"
        ]
        for c in range(1, ws_kk.max_column + 1):
            val_r1 = ws_kk.cell(1, c).value
            if val_r1:
                m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_r1))
                if m:
                    yr_num = int(m.group(0))
                    left_cc = [cc for cc in code_cols if cc < c]
                    cc = max(left_cc) if left_cc else 4
                    cmap = {}
                    for r_scan in range(4, ws_kk.max_row + 1):
                        c_val = str(ws_kk.cell(r_scan, cc).value or "").strip()
                        if c_val and c_val not in cmap:
                            cmap[c_val] = r_scan
                    kktndn_coords[yr_num] = (get_column_letter(c), cmap)

    # 4. Dò tọa độ cột của từng năm trên sheet 'TNCN' (hoặc 'TK')
    tncn_coords: Dict[int, str] = {}
    if actual_tncn_sheet and actual_tncn_sheet in wb.sheetnames:
        ws_tn = wb[actual_tncn_sheet]
        for c in range(5, ws_tn.max_column + 1):
            val_r1 = ws_tn.cell(1, c).value
            if val_r1:
                m = re.search(r"\b(19\d\d|20\d\d)\b", str(val_r1))
                if m:
                    tncn_coords[int(m.group(0))] = get_column_letter(c)

    all_years = sorted(
        set(gtgt_coords.keys())
        | set(kqkd_coords.keys())
        | set(kktndn_coords.keys())
        | set(tncn_coords.keys())
    )
    if not all_years:
        return []

    created_sheets: List[str] = []
    yr_sheet_names: Dict[int, str] = {}

    font_regular = Font(name="Times New Roman", size=11, bold=False)
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    font_italic = Font(name="Times New Roman", size=11, italic=True)

    bold_rows = {
        1, 4, 11, 14, 18, 22, 23, 24,
        27, 29, 48, 56, 63, 76, 81, 83, 90, 93, 95, 96,
    }
    italic_rows = {97, 98}
    header_rows = {1, 27, 81, 93}
    sub_header_rows = {2, 28, 82, 94}
    blank_rows = {26, 80, 91, 92}

    for yr in all_years:
        sname = sanitize_sheet_name(str(yr), existing_sheets)
        ws = wb.create_sheet(title=sname)
        yr_sheet_names[yr] = sname
        created_sheets.append(sname)

        def gtgt_ref(col_let: str, mode: str = "sum") -> Any:
            if yr in gtgt_coords and actual_gtgt_sheet:
                r_first, r_last, r_sum = gtgt_coords[yr]
                target_r = r_first if mode == "first" else (r_last if mode == "last" else r_sum)
                return f"='{actual_gtgt_sheet}'!{col_let}{target_r}"
            return 0

        is_kq_156 = True
        if yr in kqkd_coords:
            is_kq_156 = "[01]" in kqkd_coords[yr][1]

        def kqkd_ref(code_156: str, code_80: Optional[str] = None) -> Any:
            if yr in kqkd_coords and actual_kqkd_sheet:
                col_let, cmap = kqkd_coords[yr]
                target_code = code_156 if is_kq_156 else (code_80 or code_156)
                if target_code and target_code in cmap:
                    return f"='{actual_kqkd_sheet}'!{col_let}{cmap[target_code]}"
            return 0

        is_kk_156 = True
        if yr in kktndn_coords:
            is_kk_156 = "C9a" in kktndn_coords[yr][1]

        def kktndn_ref(codes_156: Tuple[str, ...], codes_80: Optional[Tuple[str, ...]] = None) -> Any:
            if yr in kktndn_coords and actual_kktndn_sheet:
                col_let, cmap = kktndn_coords[yr]
                target_codes = codes_156 if is_kk_156 else (codes_80 or codes_156)
                terms = [f"'{actual_kktndn_sheet}'!{col_let}{cmap[c]}" for c in target_codes if c in cmap]
                if terms:
                    return "=" + "+".join(terms)
            return 0

        def tncn_ref(r_idx: int) -> Any:
            if yr in tncn_coords and actual_tncn_sheet:
                return f"='{actual_tncn_sheet}'!{tncn_coords[yr]}{r_idx}"
            return 0

        # Kiểm tra xem số khấu trừ đầu năm có khớp với cuối năm trước để nối công thức liên năm ở cột D3
        d3_formula = "=C3"
        if (yr - 1) in yr_sheet_names and (yr - 1) in gtgt_coords and yr in gtgt_coords and actual_gtgt_sheet:
            ws_gt = wb[actual_gtgt_sheet]
            prev_last_r = gtgt_coords[yr - 1][1]
            curr_first_r = gtgt_coords[yr][0]
            v_prev_end = ws_gt.cell(prev_last_r, 39).value or 0
            v_curr_beg = ws_gt.cell(curr_first_r, 12).value or 0
            if v_prev_end == v_curr_beg:
                d3_formula = f"='{yr_sheet_names[yr - 1]}'!D24"

        c43_formula = kqkd_ref("[15]", "[18]")
        if c43_formula == 0:
            c43_formula = "=C30-C31-C36+C40-C41"

        rows_spec: Dict[int, Tuple[Any, Any, Any, Any, Any]] = {
            # BẢNG 1: THUẾ GTGT (Dòng 1..25)
            1: ("STT", "Chỉ tiêu", "Số báo cáo của người nộp thuế", "Số kiểm tra", "Chênh lệch"),
            2: (-1, -2, -3, -4, "(5)=(4)-(3)"),
            3: (1, "Thuế GTGT còn được khấu trừ kỳ trước chuyển sang", gtgt_ref("L", "first"), d3_formula, "=D3-C3"),
            4: (2, "Doanh thu hàng hóa, dịch vụ bán ra", "=C5+C6+C10", "=D5+D6+D10", "=D4-C4"),
            5: ("a", "Doanh thu hàng hóa, dịch vụ không chịu thuế GTGT", gtgt_ref("R"), "=C5", "=D5-C5"),
            6: ("b", "Doanh thu hàng hóa, dịch vụ chịu thuế GTGT", "=SUM(C7:C9)", "=SUM(D7:D9)", "=D6-C6"),
            7: (None, "Doanh thu chịu thuế suất 0%", gtgt_ref("U"), "=C7", "=D7-C7"),
            8: (None, "Doanh thu chịu thuế suất 5%", gtgt_ref("V"), "=C8", "=D8-C8"),
            9: (None, "Doanh thu chịu thuế suất 10%", gtgt_ref("X"), "=C9", "=D9-C9"),
            10: ("c", "Doanh thu không tính thuế ", gtgt_ref("Z"), "=C10", "=D10-C10"),
            11: (3, "Tổng giá trị hàng hóa, dịch vụ mua vào", "=C12+C13", "=D12+D13", "=D11-C11"),
            12: (None, "Giá trị hàng hoá dịch vụ mua vào của dự án đầu tư ", 0, "=C12", "=D12-C12"),
            13: (None, "Giá trị hàng hóa, dịch vụ phục vụ kinh doanh", gtgt_ref("M"), "=C13", "=D13-C13"),
            14: (4, "Thuế GTGT đầu ra", "=SUM(C15:C16)", "=SUM(D15:D16)", "=D14-C14"),
            15: (None, "Thuế GTGT đầu ra 5%", gtgt_ref("W"), "=C15", "=D15-C15"),
            16: (None, "Thuế GTGT đầu ra 10%", gtgt_ref("Y"), "=C16", "=D16-C16"),
            17: (5, "Thuế GTGT hàng hóa, dịch vụ mua vào", gtgt_ref("N"), "=C17", "=D17-C17"),
            18: (6, "Thuế GTGT hàng hóa, dịch vụ mua vào được khấu trừ ", gtgt_ref("Q"), "=C18", "=D18-C18"),
            19: (7, "Điều chỉnh giảm thuế GTGT còn được khấu trừ các kỳ trước", gtgt_ref("AD"), "=C19", "=D19-C19"),
            20: (8, "Điều chỉnh tăng thuế GTGT còn được khấu trừ các kỳ trước", gtgt_ref("AE"), "=C20", "=D20-C20"),
            21: (9, "Thuế GTGT đã nộp tại các địa phương khác của hoạt động xây dựng, lắp đặt, bán hàng vãng lai ngoại  tỉnh", gtgt_ref("AF"), "=C21", "=D21-C21"),
            22: (10, "Thuế GTGT phát sinh", gtgt_ref("AC"), "=D14-D18", "=D22-C22"),
            23: (11, "Thuế GTGT phải nộp trong kỳ", gtgt_ref("AH"), "=C23", "=D23-C23"),
            24: (12, "Thuế GTGT còn được khấu trừ chuyển kỳ sau", gtgt_ref("AM", "last"), "=C24-E22+E23+E3", "=D24-C24"),
            25: (13, "Thuế GTGT đã được hoàn", gtgt_ref("AL"), "=C25", "=D25-C25"),
            26: (None, None, None, None, None),
            # BẢNG 2: THUẾ TNDN (Dòng 27..79)
            27: ("STT", "Chỉ tiêu", "Số báo cáo của người nộp thuế", "Số kiểm tra", "Chênh lệch"),
            28: (-1, -2, -3, -4, "(5)=(4)-(3)"),
            29: ("A", "Xác định lợi nhuận kế toán trước thuế thu nhập doanh nghiệp", None, None, None),
            30: (1, "Doanh thu bán hàng hóa, dịch vụ ", kqkd_ref("[01]", "[04]"), "=C30", "=D30-C30"),
            31: (2, "Các khoản giảm trừ doanh thu ", "=SUM(C32:C35)", "=SUM(D32:D35)", "=D31-C31"),
            32: ("a", " Chiết khấu thương mại ", kqkd_ref("[04]", "[07]"), "=C32", "=D32-C32"),
            33: ("b", " Giảm giá hàng bán", kqkd_ref("[05]", "[08]"), "=C33", "=D33-C33"),
            34: ("c", " Giá trị hàng bán bị trả lai", kqkd_ref("[06]", "[09]"), "=C34", "=D34-C34"),
            35: ("d", "Thuế tiêu thụ đặc biệt, thuế xuất khẩu, thuế giá trị gia tăng theo phương pháp trực tiếp phải nộp", kqkd_ref("[07]", "") if is_kq_156 else 0, "=C35", "=D35-C35"),
            36: (3, "Chi phí sản xuất kinh doanh hàng hóa", "=SUM(C37:C39)", "=SUM(D37:D39)", "=D36-C36"),
            37: ("a", "Giá vốn hàng bán", kqkd_ref("[10]", "[13]"), "=C37", "=D37-C37"),
            38: ("b", " Chi phí bán hàng", kqkd_ref("[11]", "[14]"), "=C38", "=D38-C38"),
            39: ("c", " Chi phí quản lý doanh nghiệp", kqkd_ref("[12]", "[15]"), "=C39", "=D39-C39"),
            40: (4, "Doanh thu hoạt động tài chính", kqkd_ref("[08]", "[10]"), "=C40", "=D40-C40"),
            41: (5, "Chi phí hoạt động tài chính  ", kqkd_ref("[13]", "[16]"), "=C41", "=D41-C41"),
            42: (None, "Trong đó: Chi phí lãi tiền vay", kqkd_ref("[14]", "[17]"), "=C42", "=D42-C42"),
            43: (6, "Lợi nhuận thuần từ hoạt động kinh doanh", c43_formula, "=D30-D31-D36+D40-D41", "=D43-C43"),
            44: (7, "Thu nhập khác", kqkd_ref("[16]", "[19]"), "=C44", "=D44-C44"),
            45: (8, "Chi phí khác", kqkd_ref("[17]", "[20]"), "=C45", "=D45-C45"),
            46: (9, "Lợi nhuận khác", "=C44-C45", "=D44-D45", "=D46-C46"),
            47: (10, "Tổng lợi nhuận kế toán trước thuế TNDN", "=C43+C46", "=D43+D46", "=D47-C47"),
            48: ("B", "Xác định thu nhập chịu thuế theo luật thuế TNDN", None, None, None),
            49: (1, "Điều chỉnh tăng tổng lợi nhuận trước thuế thu nhập doanh nghiệp", "=SUM(C50:C55)", "=SUM(D50:D55)", "=D49-C49"),
            50: ("a", "Các khoản điều chỉnh tăng doanh thu", kktndn_ref(("B2",)), "=C50", "=D50-C50"),
            51: ("b", "Chi phí của phần doanh thu điều chỉnh giảm", kktndn_ref(("B3",)), "=C51", "=D51-C51"),
            52: ("c", "Các khoản chi không được trừ khi xác định thu nhập chịu thuế", kktndn_ref(("B4",)), "=C52", "=D52-C52"),
            53: ("d", "Thuế thu nhập đã nộp cho phần thu nhập nhận được ở nước ngoài", kktndn_ref(("B5",)), "=C53", "=D53-C53"),
            54: ("e", "Điều chỉnh tăng lợi nhuận do xác định giá thị trường đối với giao dịch liên kết", kktndn_ref(("B6",)), "=C54", "=D54-C54"),
            55: ("f", "Các khoản điều chỉnh làm tăng lợi nhuận trước thuế khác", kktndn_ref(("B7",)), "=C55", "=D55-C55"),
            56: (2, "Điều chỉnh giảm tổng lợi nhuận trước thuế thu nhập doanh nghiệp", "=SUM(C57:C59)", "=SUM(D57:D59)", "=D56-C56"),
            57: ("a", "Giảm trừ các khoản doanh thu đã tính thuế năm trước", kktndn_ref(("B9",)), "=C57", "=D57-C57"),
            58: ("b", "Chi phí của phần doanh thu điều chỉnh tăng", kktndn_ref(("B10",)), "=C58", "=D58-C58"),
            59: ("c", "Các khoản điều chỉnh làm giảm lợi nhuận trước thuế khác", kktndn_ref(("B11",), ("B11", "B12")), "=C59", "=D59-C59"),
            60: (3, "Tổng thu nhập chịu thuế (=A9+B1-B2)", "=C61+C62", "=D61+D62", "=D60-C60"),
            61: ("a", "Thu nhập chịu thuế từ hoạt động SXKD", "=C47+C49-C56", "=D47+D49-D56", "=D61-C61"),
            62: ("b", "Thu nhập chịu thuế từ hoạt động chuyển nhượng BĐS ", kktndn_ref(("B14",), ("B15",)), "=C62", "=D62-C62"),
            63: ("C", "Xác định thuế TNDN phải nộp từ hoạt động sản xuất kinh doanh", None, None, None),
            64: (1, "Thu nhập chịu thuế ", "=C47+C49-C56", "=D47+D49-D56", "=D64-C64"),
            65: (2, "Thu nhập miễn thuế", kktndn_ref(("C2",)), "=C65", "=D65-C65"),
            66: (3, "Chuyển lỗ, chi phí lãi vay và bù trừ lãi, lỗ ", "=SUM(C67:C68)", "=SUM(D67:D68)", "=D66-C66"),
            67: ("a", "Lỗ từ hoạt động sản xuất kinh doanh được chuyển trong kỳ", kktndn_ref(("C3a",)), "=C67", "=D67-C67"),
            68: ("b", "Lỗ từ hoạt động chuyển nhượng BĐS được bù trừ với lãi của hoạt động sản xuất kinh doanh", kktndn_ref(("C3b",)), "=C68", "=D68-C68"),
            69: (4, "Thu nhập tính thuế (=C1-C2-C3)", "=C64-C66-C65", "=D64-D66-D65", "=D69-C69"),
            70: (5, "Trích lập quỹ khoa học công nghệ", kktndn_ref(("C5",)), "=C70", "=D70-C70"),
            71: (6, "Thu nhập tính thuế sau khi đã trích lập quỹ khoa học công nghệ (=C4-C5)", "=C69-C70", "=D69-D70", "=D71-C71"),
            72: (7, "Thuế TNDN phải nộp theo thuế suất phổ thông ", "=IF(C71>0,ROUND(C71*0.2,0),0)", "=IF(D71>0,ROUND(D71*0.2,0),0)", "=D72-C72"),
            73: (8, "Thuế TNDN chênh lệch do áp dụng mức thuế suất ưu đãi", kktndn_ref(("C11",)), "=C73", "=D73-C73"),
            74: (9, "Thuế TNDN được miễn, giảm trong kỳ", kktndn_ref(("C12",), ("C12", "C13", "C14", "C15")), "=C74", "=D74-C74"),
            75: (10, "Số thuế TNDN đã nộp ở nước ngoài được trừ trong kỳ tính thuế", kktndn_ref(("C15",), ("C16",)), "=C75", "=D75-C75"),
            76: ("D", "Thuế TNDN phải nộp trong kỳ", "=SUM(C77:C79)", "=SUM(D77:D79)", "=D76-C76"),
            77: (1, "Thuế TNDN từ hoạt động SXKD ", "=C72-C73-C74-C75", "=D72-D73-D74-D75", "=D77-C77"),
            78: (2, "Thuế TNDN phải nộp từ hoạt động chuyển nhượng BĐS (nếu có)", kktndn_ref(("E2",)), "=C78", "=D78-C78"),
            79: (3, "Thuế TNDN phải nộp khác (nếu có)", kktndn_ref(("E3",), ("E5",)), "=C79", "=D79-C79"),
            80: (None, None, None, None, None),
            # BẢNG 3: THUẾ TNCN (Dòng 81..90)
            81: ("STT", "Chỉ tiêu", "Số báo cáo của người nộp thuế", "Số kiểm tra", "Chênh lệch"),
            82: (-1, -2, -3, -4, "(5)=(4)-(3)"),
            83: ("I", "Thuế TNCN từ tiền lương, tiền công ", None, None, None),
            84: (1, "Thu nhập từ tiền lương, tiền công của cá nhân cư trú tính thuế theo biểu thuế lũy tiến từng phần ", tncn_ref(41), "=C84", "=D84-C84"),
            85: (2, "Thuế TNCN từ tiền lương, tiền công của cá nhân cư trú tính thuế theo biểu thuế lũy tiến từng phần đã khấu trừ", tncn_ref(42), "=C85", "=D85-C85"),
            86: (3, "Thu nhập từ tiền lương, tiền công của cá nhân cư trú tính thuế theo thuế suất toàn phần 10% ", tncn_ref(43), "=C86", "=D86-C86"),
            87: (4, "Thuế TNCN từ tiền lương, tiền công của cá nhân cư trú tính thuế theo thuế suất toàn phần 10% đã khấu trừ", tncn_ref(44), "=C87", "=D87-C87"),
            88: (5, "Thu nhập từ tiền lương, tiền công của cá nhân  không cư trú", tncn_ref(45), "=C88", "=D88-C88"),
            89: (6, "Thuế TNCN từ tiền lương, tiền công của cá nhân không cư trú tính theo thuế suất 20% đã khấu trừ", tncn_ref(46), "=C89", "=D89-C89"),
            90: ("Tổng hợp tổng số thuế TNCN phải khấu trừ", None, "=C85+C87+C89", "=D85+D87+D89", "=E85+E87+E89"),
            91: (None, None, None, None, None),
            92: (None, None, None, None, None),
            # BẢNG 4: THUẾ NHÀ THẦU (Dòng 93..98)
            93: ("STT", "Chỉ tiêu", "Số báo cáo của người nộp thuế", "số kiểm tra", "Chênh lệch"),
            94: (-1, -2, -3, -4, "(5)=(4)-(3)"),
            95: (1, "Doanh thu tính thuế của nhà thầu nước ngoài ", 0, "=C95", "=D95-C95"),
            96: (2, "Thuế nhà thầu phải khấu trừ  ", "=C97+C98", "=D97+D98", "=E97+E98"),
            97: ("a", "Thuế GTGT", 0, "=C97", "=D97-C97"),
            98: ("b", "Thuế TNDN ", 0, "=C98", "=D98-C98"),
        }

        for r_idx in range(1, 99):
            vals = rows_spec[r_idx]
            if r_idx in blank_rows:
                continue

            row_font = font_bold if r_idx in bold_rows else (font_italic if r_idx in italic_rows else font_regular)
            for c_idx, val in enumerate(vals, start=1):
                cell = ws.cell(r_idx, c_idx, val)
                cell.font = row_font
                cell.border = THIN_BORDER
                if r_idx in header_rows or r_idx in sub_header_rows:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif c_idx == 1:
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif c_idx == 2:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                else:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(horizontal="right", vertical="center")

        ws.merge_cells("B83:E83")
        ws.merge_cells("A90:B90")
        ws.cell(90, 1).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws.column_dimensions["A"].width = 6
        ws.column_dimensions["B"].width = 45
        ws.column_dimensions["C"].width = 18
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 16

    return created_sheets


# ==============================================================================
# 6. SHEET TỜ KHAI KHÁC (LỆ PHÍ MÔN BÀI 01/LPMB, THUẾ NHÀ THẦU, ...)
# ==============================================================================
def write_other_declarations_sheet(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    sheet_name: str = "To_Khai_Khac",
) -> None:
    """
    Tổng hợp đầy đủ dữ liệu các tờ khai thuộc nhóm 'Khác' (ví dụ: Tờ khai Lệ phí môn bài 01/LPMB,
    Tờ khai thuế nhà thầu...) đúng bản chất dữ liệu:
    - Với Tờ khai Lệ phí môn bài (01/LPMB): Tách chuẩn từng dòng Người nộp LPMB và từng Đơn vị phụ thuộc /
      Địa điểm kinh doanh (DiaDiemKinhDoanh[1..N]) để không sót bất kỳ chi nhánh/địa điểm kinh doanh nào.
    - Với các tờ khai Khác tổng quát: Xuất đầy đủ các chỉ tiêu không trùng lặp khóa alias.
    """
    unique_others = deduplicate_declarations(declarations)
    if not unique_others:
        return

    sname = sanitize_sheet_name(sheet_name, existing_sheets)
    ws = wb.create_sheet(title=sname)

    sorted_others = sorted(
        unique_others,
        key=lambda d: (d.mst or "", d.year or 0, d.ky_kkhai or "", int(d.so_lan or 0) if str(d.so_lan or "").isdigit() else 0),
    )

    has_lpmb_structure = any(
        any("nguoinop" in k.lower() or "diadiem" in k.lower() for k in d.indicators.keys())
        for d in sorted_others
    )

    if has_lpmb_structure:
        headers = [
            "STT",
            "Tên file XML",
            "Mã tờ khai",
            "Tên tờ khai",
            "Kỳ kê khai",
            "Lần kê khai",
            "MST người nộp thuế",
            "Tên người nộp thuế",
            "Ngày lập",
            "Khai bổ sung ĐĐKD mới [09]",
            "Miễn LPMB năm đầu [10]",
            "Loại đơn vị kê khai",
            "Tên đơn vị / Địa điểm kinh doanh [02]",
            "MST / Mã địa điểm kinh doanh [04]",
            "Mã xã/phường",
            "Tên xã/phường [03]",
            "Mã tỉnh/TP",
            "Tên tỉnh/TP [03a]",
            "Vốn điều lệ / Doanh thu [05]",
            "Mức lệ phí môn bài [06]",
            "Số tiền LPMB phải nộp [07]",
            "Số tiền LPMB được miễn [08]",
            "Tổng số LPMB các ĐĐKD",
            "Tổng số LPMB phải nộp cả tờ khai",
        ]
        ws.append(headers)
        ws.row_dimensions[1].height = 28
        for c_idx in range(1, len(headers) + 1):
            cell = ws.cell(1, c_idx)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER
            ws.column_dimensions[get_column_letter(c_idx)].width = 20

        cur_row = 2
        stt_idx = 1
        text_col_indices = {3, 5, 7, 14, 15, 17}
        center_col_indices = {1, 3, 5, 6, 9, 10, 11, 15, 17}

        for d in sorted_others:
            inds = d.indicators
            ct09_raw = inds.get("ct09", "false")
            ct10_raw = inds.get("ct10", "false")
            ct09_str = "Có" if str(ct09_raw).lower() in ("true", "1", "[x]") else "Không"
            ct10_str = "Có" if str(ct10_raw).lower() in ("true", "1", "[x]") else "Không"
            tong_ddkd = inds.get("tongcong_ddkd", 0) or 0
            tong_all = inds.get("tong", inds.get("ct06", 0)) or 0
            lan_num = int(d.so_lan) if str(d.so_lan or "").isdigit() else 0

            # Thu thập danh sách các đơn vị (Người nộp LPMB + các Địa điểm kinh doanh)
            unit_prefixes: List[Tuple[str, str]] = []
            if any(k.startswith("NguoiNopLPMB.") for k in inds.keys()):
                unit_prefixes.append(("1. Người nộp lệ phí môn bài", "NguoiNopLPMB"))

            ddkd_prefixes: List[str] = []
            for k in inds.keys():
                if k.startswith("DiaDiemKD.DiaDiemKinhDoanh"):
                    p = k.split(".ct")[0].split(".ma_")[0].split(".ten_")[0]
                    if p not in ddkd_prefixes:
                        ddkd_prefixes.append(p)
            for b_idx, dp in enumerate(ddkd_prefixes, start=1):
                unit_prefixes.append((f"2.{b_idx} Đơn vị phụ thuộc / ĐĐKD", dp))

            if not unit_prefixes:
                unit_prefixes.append(("1. Người nộp lệ phí môn bài", ""))

            for u_label, u_pref in unit_prefixes:
                if u_pref:
                    u_ten = inds.get(f"{u_pref}.ct2") or d.ten_nnt
                    u_mst = inds.get(f"{u_pref}.ct4") or d.mst
                    u_ma_xa = inds.get(f"{u_pref}.ma_xa") or ""
                    u_ten_xa = inds.get(f"{u_pref}.ten_xa") or ""
                    u_ma_tinh = inds.get(f"{u_pref}.ct3a_ma") or ""
                    u_ten_tinh = inds.get(f"{u_pref}.ct3a_ten") or ""
                    u_von = inds.get(f"{u_pref}.ct5", 0) or 0
                    u_muc = inds.get(f"{u_pref}.ct6", 0) or 0
                    u_pnop = inds.get(f"{u_pref}.ct7", 0) or 0
                    u_mien = inds.get(f"{u_pref}.ct8", 0) or 0
                else:
                    u_ten = d.ten_nnt
                    u_mst = d.mst
                    u_ma_xa = ""
                    u_ten_xa = ""
                    u_ma_tinh = ""
                    u_ten_tinh = ""
                    u_von = inds.get("ct05", 0) or 0
                    u_muc = inds.get("ct06", 0) or 0
                    u_pnop = inds.get("ct07", inds.get("ct06", 0)) or 0
                    u_mien = inds.get("ct08", 0) or 0

                row_vals = [
                    stt_idx,
                    d.file_name,
                    d.ma_tkhai,
                    d.ten_tkhai,
                    d.ky_kkhai,
                    lan_num,
                    d.mst,
                    d.ten_nnt,
                    d.ngay_lap,
                    ct09_str,
                    ct10_str,
                    u_label,
                    u_ten,
                    str(u_mst) if u_mst is not None else "",
                    str(u_ma_xa) if u_ma_xa != "" else "",
                    u_ten_xa,
                    str(u_ma_tinh) if u_ma_tinh != "" else "",
                    u_ten_tinh,
                    u_von,
                    u_muc,
                    u_pnop,
                    u_mien,
                    tong_ddkd,
                    tong_all,
                ]
                for c_i, val in enumerate(row_vals, start=1):
                    cell = ws.cell(cur_row, c_i)
                    write_cell_formatted(cell, val, is_mst=(c_i in text_col_indices))
                    if c_i in center_col_indices:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                cur_row += 1
                stt_idx += 1

        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{cur_row - 1}"
        ws.freeze_panes = "F2"
        return

    # Trường hợp tờ khai Khác tổng quát
    raw_keys: List[str] = []
    seen_lower: Set[str] = set()
    for d in unique_others:
        for k in d.indicators.keys():
            kl = k.lower()
            if "." in k or k == kl:
                if kl not in seen_lower:
                    seen_lower.add(kl)
                    raw_keys.append(k)

    prefixed_leafs = {k.split(".")[-1].lower() for k in raw_keys if "." in k}
    filtered_keys = [k for k in raw_keys if ("." in k or k.lower() not in prefixed_leafs)]

    fixed_headers = [
        "STT",
        "Tên file XML",
        "Mã tờ khai",
        "Tên tờ khai",
        "Kỳ kê khai",
        "Lần kê khai",
        "MST",
        "Tên NNT",
        "Ngày lập",
    ]
    headers = fixed_headers + filtered_keys
    ws.append(headers)

    ws.row_dimensions[1].height = 26
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, c_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        ws.column_dimensions[get_column_letter(c_idx)].width = 20

    for idx, d in enumerate(sorted_others, start=1):
        r_idx = idx + 1
        lan_num = int(d.so_lan) if str(d.so_lan or "").isdigit() else 0
        fixed_vals = [
            idx,
            d.file_name,
            d.ma_tkhai,
            d.ten_tkhai,
            d.ky_kkhai,
            lan_num,
            d.mst,
            d.ten_nnt,
            d.ngay_lap,
        ]
        for c_off, v in enumerate(fixed_vals, start=1):
            write_cell_formatted(ws.cell(r_idx, c_off), v, is_mst=(c_off in (3, 5, 7)))
        for k_off, k in enumerate(filtered_keys):
            val = d.indicators.get(k)
            kl = k.lower()
            is_mst_col = any(t in kl for t in ("mst", ".ct4", "ma_xa", "ct3a_ma", "ma_cqt"))
            write_cell_formatted(ws.cell(r_idx, len(fixed_headers) + 1 + k_off), val, is_mst=is_mst_col)

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(sorted_others) + 1}"
    ws.freeze_panes = "F2"


# ==============================================================================
# 7. SHEET TỔNG HỢP DANH SÁCH FILE VÀ PHỤ LỤC KHÁC
# ==============================================================================
def write_summary_sheet(
    wb: openpyxl.Workbook,
    all_declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
) -> None:
    sname = sanitize_sheet_name("Tong_hop_File", existing_sheets)
    ws = wb.create_sheet(title=sname, index=0)

    headers = [
        "STT",
        "Tên file XML",
        "Nhóm thuế",
        "Mã tờ khai",
        "Lần",
        "MST",
        "Tên NNT",
        "Kỳ kê khai",
        "Ngày lập",
        "Trạng thái",
        "Đường dẫn",
    ]
    ws.append(headers)

    seen_keys: Set[Tuple[str, str, str, str]] = set()
    for idx, d in enumerate(all_declarations, start=1):
        row_idx = idx + 1
        if d.error:
            status = f"Lỗi: {d.error}"
        else:
            ky_str = (
                d.tu_ngay.strftime("%y%m")
                if (d.tax_type == "GTGT" and d.tu_ngay)
                else str(d.year or d.ky_kkhai or "")
            )
            lan_str = str(d.so_lan or "0").strip()
            dedup_k = (d.tax_type or "", d.mst or "", ky_str, lan_str)
            if dedup_k in seen_keys:
                status = "Thành công (Trùng tờ khai)"
            else:
                seen_keys.add(dedup_k)
                status = "Thành công"
        lan_val = int(d.so_lan) if str(d.so_lan or "").isdigit() else (d.so_lan or 0)
        vals = [
            idx,
            d.file_name,
            d.tax_type,
            d.ma_tkhai,
            lan_val,
            d.mst,
            d.ten_nnt,
            d.ky_kkhai,
            d.ngay_lap,
            status,
            d.file_path,
        ]
        for c_idx, v in enumerate(vals, start=1):
            cell = ws.cell(row_idx, c_idx)
            write_cell_formatted(cell, v, is_mst=(c_idx in (4, 6, 8)))
            if c_idx in (1, 4, 5, 8):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            if d.error:
                cell.fill = ERROR_FILL

    ws.row_dimensions[1].height = 26
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(1, c_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
        col_letter = get_column_letter(c_idx)
        ws.column_dimensions[col_letter].width = 18

    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(all_declarations) + 1}"
    ws.freeze_panes = "A2"


APPENDIX_HEADER_LABELS: Dict[str, str] = {
    "ma_giao_dich_goc": "Mã giao dịch tờ khai gốc",
    "tieu_muc_ma": "Mã tiểu mục",
    "tieu_muc_ten": "Tên tiểu mục",
    "thue_phai_nop_tang_giam": "Thuế phải nộp điều chỉnh tăng/giảm",
    "thue_hoan_phai_nop_tang_giam": "Thuế đã hoàn phải nộp trả tăng/giảm",
    "thue_khau_tru_tang_giam": "Thuế khấu trừ điều chỉnh tăng/giảm",
    "thue_de_nghi_hoan_tang_giam": "Thuế đề nghị hoàn điều chỉnh tăng/giảm",
    "ngay_tinh_cham_nop": "Ngày tính chậm nộp",
    "so_ngay_cham_nop": "Số ngày chậm nộp",
    "tien_cham_nop": "Tiền chậm nộp",
    "ly_do_giai_trinh": "Lý do / Giải trình",
    "ma_HSo": "Mã hồ sơ",
    "ten_HSo": "Tên tờ khai / Phụ lục",
    "ct3_ma": "Mã chỉ tiêu điều chỉnh",
    "ct3_ten": "Tên chỉ tiêu điều chỉnh",
    "ct04": "Số đã kê khai [04]",
    "ct05": "Số điều chỉnh [05]",
    "ct06": "Chênh lệch [06]=[05]-[04]",
    "ct7": "Thuế đã hoàn phải nộp trả [07]",
    "ct3_2_ten": "Lý do điều chỉnh chi tiết",
    "nhom_phu_luc": "Nhóm bảng kê (Mua vào / Bán ra)",
    "tenHHDVMuaVao": "Tên HHDV mua vào",
    "giaTriHHDVMuaVao": "Giá trị HHDV mua vào chưa thuế",
    "thueGTGTHHDV": "Thuế GTGT HHDV mua vào được giảm",
    "tenHHDV": "Tên HHDV bán ra",
    "giaTriHHDV": "Giá trị HHDV bán ra chưa thuế",
    "thueSuatTheoQuyDinh": "Thuế suất theo quy định (%)",
    "thueSuatSauGiam": "Thuế suất sau giảm (%)",
    "thueGTGTDuocGiam": "Thuế GTGT HHDV bán ra được giảm",
    "nhom_lo": "Nhóm chuyển lỗ",
    "namPSinhLo": "Năm phát sinh lỗ",
    "soLoPSinh": "Số lỗ phát sinh",
    "soLoDaChuyen": "Số lỗ đã chuyển các kỳ trước",
    "soLoKyNay": "Số lỗ được chuyển kỳ này",
    "soLoChuyenSau": "Số lỗ còn được chuyển kỳ sau",
    "tenLKet": "Tên bên liên kết",
    "ma_quocGia": "Mã quốc gia",
    "ten_quocGia": "Tên quốc gia",
    "mstLKet": "MST bên liên kết",
    "ma_chi_tieu": "Mã số",
    "ten_chi_tieu": "Chỉ tiêu",
    "thuyet_minh": "Thuyết minh",
    "nam_nay": "Số năm nay",
    "nam_truoc": "Số năm trước",
}


def filter_latest_per_period(declarations: List[TaxDeclaration]) -> List[TaxDeclaration]:
    """
    Lọc bản kê khai có hiệu lực cuối cùng cho mỗi kỳ tính thuế (Nhóm thuế, MST, Kỳ tính thuế)
    và tự động kế thừa các phụ lục nghiệp vụ không thay đổi từ lần kê khai trước của cùng kỳ.
    Dùng cho các sheet phụ lục nghiệp vụ (NQ43, NQ101, NQ142, 03-1A, 03-2, GDLK, KQHDSXKD, LCTTGT)
    để không bị nhân đôi số liệu giữa tờ khai lần đầu (CT0) và các tờ khai bổ sung (BS1, BS2...).
    """
    unique_decls = deduplicate_declarations(declarations)
    grouped: Dict[Tuple[str, str, str], List[TaxDeclaration]] = {}
    for d in unique_decls:
        if d.tax_type == "GTGT" and d.tu_ngay:
            ky_norm = d.tu_ngay.strftime("%y%m")
        else:
            ky_norm = str(d.year or d.ky_kkhai or "").strip()
        key = (d.tax_type or "", (d.mst or "").strip(), ky_norm)
        grouped.setdefault(key, []).append(d)

    latest_list: List[TaxDeclaration] = []
    for key, dlist in grouped.items():
        def sort_k(item: TaxDeclaration):
            try:
                lan = int(item.so_lan) if item.so_lan is not None and str(item.so_lan).strip() != "" else 0
            except (ValueError, TypeError):
                lan = 0
            nlap = item.ngay_lap or datetime.date.min
            valid_inds = sum(1 for v in item.indicators.values() if v not in (None, 0, "", "0"))
            app_count = sum(len(rows) for rows in item.appendices.values())
            return (lan, nlap, valid_inds, app_count, len(item.file_name), item.file_name)

        dlist.sort(key=sort_k)
        latest = dlist[-1]
        if len(dlist) > 1:
            merged_apps = dict(latest.appendices)
            for prev in reversed(dlist[:-1]):
                for app_name, app_rows in prev.appendices.items():
                    if "khbs" in app_name.lower():
                        continue
                    if not merged_apps.get(app_name) and app_rows:
                        merged_apps[app_name] = app_rows
            latest.appendices = merged_apps
        latest_list.append(latest)

    latest_list.sort(
        key=lambda d: (
            d.tax_type or "",
            d.mst or "",
            d.year or 0,
            d.tu_ngay or datetime.date.min,
            d.ky_kkhai or "",
        )
    )
    return latest_list


def write_appendix_sheets(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
) -> None:
    """
    Xuất các sheet phụ lục (PL_...) đã được làm sạch và chuẩn hóa đúng bản chất dữ liệu:
    - Với PL01_KHBS và PL01_1_KHBS: Giữ toàn bộ các lần khai bổ sung (đã khử trùng file) để phản ánh
      đầy đủ lịch sử điều chỉnh của từng lần bổ sung.
    - Với các phụ lục nghiệp vụ (NQ43, NQ101, NQ142, 03-1A, 03-2, GDLK, KQHDSXKD, LCTTGT): Chỉ lấy
      bản kê khai hiệu lực cuối cùng của mỗi kỳ tính thuế để không bị trùng lặp số liệu giữa CT0 và BS.
    - Bỏ qua các phụ lục 05-1, 05-2, 05-3 của TNCN vì đã được xuất thành 3 sheet chuẩn mẫu '05-1', '05-2', '05-3'.
    """
    unique_decls = sorted(
        deduplicate_declarations(declarations),
        key=lambda d: (
            d.tax_type or "",
            d.mst or "",
            d.year or 0,
            d.tu_ngay or datetime.date.min,
            int(d.so_lan) if str(d.so_lan or "").isdigit() else 0,
        ),
    )
    latest_period_decls = filter_latest_per_period(unique_decls)

    grouped: Dict[str, List[Dict[str, Any]]] = {}

    # 1. Thu thập PL01_KHBS và PL01_1_KHBS từ tất cả các tờ khai bổ sung (unique_decls)
    for d in unique_decls:
        lan_val = int(d.so_lan) if str(d.so_lan or "").isdigit() else (d.so_lan or 0)
        for pl_name, rows in d.appendices.items():
            pl_u = pl_name.upper()
            if "KHBS" not in pl_u:
                continue
            bucket = grouped.setdefault(pl_name, [])
            for r in rows:
                enriched: Dict[str, Any] = {
                    "_file": d.file_name,
                    "_mst": d.mst,
                    "_ky": d.ky_kkhai,
                    "_so_lan": lan_val,
                    "_ma_tkhai": d.ma_tkhai,
                }
                for k, v in r.items():
                    if k not in enriched:
                        enriched[k] = v
                bucket.append(enriched)

    # 2. Thu thập các phụ lục nghiệp vụ từ bản hiệu lực cuối cùng của mỗi kỳ tính thuế (latest_period_decls)
    for d in latest_period_decls:
        lan_val = int(d.so_lan) if str(d.so_lan or "").isdigit() else (d.so_lan or 0)
        for pl_name, rows in d.appendices.items():
            pl_u = pl_name.upper()
            if "KHBS" in pl_u:
                continue
            if any(t in pl_u for t in ("05_1", "05-1", "051", "05_2", "05-2", "052", "05_3", "05-3", "053")):
                continue
            bucket = grouped.setdefault(pl_name, [])
            for r in rows:
                enriched = {
                    "_file": d.file_name,
                    "_mst": d.mst,
                    "_ky": d.ky_kkhai,
                    "_so_lan": lan_val,
                    "_ma_tkhai": d.ma_tkhai,
                }
                for k, v in r.items():
                    if k not in enriched:
                        enriched[k] = v
                bucket.append(enriched)

    text_code_keys = {
        "_mst",
        "_ky",
        "_ma_tkhai",
        "mstlket",
        "ma_hso",
        "ct3_ma",
        "ma_chi_tieu",
        "tieu_muc_ma",
        "ma_giao_dich_goc",
        "ma_quocgia",
        "thuyet_minh",
    }
    plain_int_keys = {
        "_so_lan",
        "nampsinhlo",
        "so_ngay_cham_nop",
        "thuesuattheoquydinh",
        "thuesuatsaugiam",
    }

    for pl_name, rows in grouped.items():
        if not rows:
            continue
        sname = sanitize_sheet_name(f"PL_{pl_name}", existing_sheets)
        ws = wb.create_sheet(title=sname)

        fixed_cols = ["_file", "_mst", "_ky", "_so_lan", "_ma_tkhai"]
        fixed_labels = ["Tên file XML", "MST người nộp thuế", "Kỳ kê khai", "Lần kê khai", "Mã tờ khai"]

        dyn_keys: List[str] = []
        seen_lower: Set[str] = set()
        for r in rows:
            for k in r.keys():
                if k not in fixed_cols:
                    kl = k.lower()
                    if kl not in seen_lower:
                        seen_lower.add(kl)
                        dyn_keys.append(k)

        pl_u = pl_name.upper()
        kqkd_label_map: Dict[str, str] = {}
        if "03_1A" in pl_u:
            src_rows = KQKD_156_ROWS if "PL03_1A" in pl_u else KQKD_80_ROWS
            for _stt, desc, code in src_rows:
                clean_num = code.strip("[]")
                kqkd_label_map[f"ct{clean_num}"] = f"{desc} {code}"

        dyn_labels: List[str] = []
        for k in dyn_keys:
            if k in kqkd_label_map:
                dyn_labels.append(kqkd_label_map[k])
            elif k in ("ct04", "ct05", "ct06", "ct7") and "KHBS" not in pl_u:
                dyn_labels.append(k)
            else:
                dyn_labels.append(APPENDIX_HEADER_LABELS.get(k, k))
        headers = ["STT"] + fixed_labels + dyn_labels
        ws.append(headers)

        for idx, r in enumerate(rows, start=1):
            row_idx = idx + 1
            c_stt = ws.cell(row_idx, 1)
            write_cell_formatted(c_stt, idx)
            c_stt.alignment = Alignment(horizontal="center", vertical="center")

            for c_off, fk in enumerate(fixed_cols):
                col_i = 2 + c_off
                cell = ws.cell(row_idx, col_i)
                fkl = fk.lower()
                is_txt = fkl in text_code_keys
                write_cell_formatted(cell, r.get(fk), is_mst=is_txt)
                if fkl in ("_ky", "_so_lan", "_ma_tkhai"):
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            for k_off, dk in enumerate(dyn_keys):
                col_i = 2 + len(fixed_cols) + k_off
                cell = ws.cell(row_idx, col_i)
                val = _first_not_none(r, (dk,))
                dkl = dk.lower()
                is_txt = (
                    dkl in text_code_keys
                    or "mst" in dkl
                    or "socmt" in dkl
                    or "cccd" in dkl
                    or dkl.startswith("ma_")
                    or dkl.endswith("_ma")
                )
                write_cell_formatted(cell, val, is_mst=is_txt)
                if dkl in plain_int_keys:
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "0"
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif is_txt and dkl != "ma_giao_dich_goc":
                    cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.row_dimensions[1].height = 28
        for c in range(1, len(headers) + 1):
            cell = ws.cell(1, c)
            cell.fill = HEADER_FILL
            cell.font = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = THIN_BORDER
            ws.column_dimensions[get_column_letter(c)].width = 20
        ws.column_dimensions["A"].width = 7
        ws.column_dimensions["B"].width = 36

        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
        ws.freeze_panes = "G2"


# ==============================================================================
# HÀM XUẤT CHÍNH: EXPORT TO EXCEL (Hỗ trợ tách file theo năm, theo sắc thuế hoặc gộp)
# ==============================================================================
def _populate_workbook_sheets(
    wb: openpyxl.Workbook,
    all_decls: List[TaxDeclaration],
    valid_decls: List[TaxDeclaration],
    include_appendices: bool = True,
    cutoff_date: Optional[datetime.date] = None,
) -> None:
    """Điền đầy đủ các sheet dữ liệu chuẩn vào Workbook."""
    existing_sheets: Set[str] = set()

    write_summary_sheet(wb, all_decls, existing_sheets)

    tndn_items = [d for d in valid_decls if d.tax_type == "TNDN"]
    fiscal_ranges: Dict[int, Tuple[datetime.date, datetime.date]] = {}
    if tndn_items:
        latest_tndn = filter_latest_per_year(tndn_items)
        fiscal_ranges = detect_tndn_fiscal_ranges(latest_tndn)
        write_sheet_kktndn(wb, latest_tndn, existing_sheets)
        write_sheet_kqkd(wb, latest_tndn, existing_sheets)

    bctc_items = [d for d in valid_decls if d.tax_type == "BCTC"]
    if bctc_items:
        latest_bctc = filter_latest_per_year(bctc_items)
        write_sheet_bctc_cdkt(wb, latest_bctc, existing_sheets)

    tncn_items = [d for d in valid_decls if d.tax_type == "TNCN"]
    if tncn_items:
        latest_tncn = filter_latest_per_year(tncn_items)
        write_sheet_tncn_master(wb, latest_tncn, existing_sheets, include_appendices=include_appendices)
        if include_appendices:
            write_consolidated_tncn_appendices(wb, latest_tncn, existing_sheets)

    gtgt_items = [d for d in valid_decls if d.tax_type == "GTGT"]
    if gtgt_items:
        write_gtgt_timeline_sheet(wb, gtgt_items, existing_sheets, sheet_name="So Thue GTGT TH")
        write_gtgt_cuoi_sheet(wb, gtgt_items, existing_sheets, sheet_name="So Thue GTGT Cuoi")
        if fiscal_ranges:
            write_gtgt_cuoi_sheet(
                wb,
                gtgt_items,
                existing_sheets,
                sheet_name="So Thue GTGT Cuoi (Lech ky)",
                fiscal_ranges=fiscal_ranges,
            )

    if gtgt_items or tndn_items:
        write_sheet_chenh_lech_gtgt_tndn(
            wb,
            existing_sheets,
            gtgt_sheet_name="So Thue GTGT Cuoi (Lech ky)" if (gtgt_items and fiscal_ranges) else "So Thue GTGT Cuoi",
            fiscal_ranges=fiscal_ranges or None,
        )

    if gtgt_items or tndn_items or tncn_items:
        write_yearly_verification_sheets(wb, existing_sheets)
        from penalty_calculator import add_late_payment_sheet
        add_late_payment_sheet(
            wb,
            cutoff_date=cutoff_date,
            fiscal_ranges=fiscal_ranges or None,
        )

    other_items = [d for d in valid_decls if d.tax_type == "Khác"]
    if other_items:
        write_other_declarations_sheet(wb, other_items, existing_sheets, sheet_name="To_Khai_Khac")

    if include_appendices and valid_decls:
        write_appendix_sheets(wb, valid_decls, existing_sheets)


def export_to_excel(
    declarations: List[TaxDeclaration],
    output_path: str,
    include_appendices: bool = True,
    split_by_tax_type: bool = False,
    split_by_year: bool = False,
    layout: str = "vertical",
    cutoff_date: Optional[datetime.date] = None,
) -> List[str]:
    """
    Xuất danh sách TaxDeclaration ra file Excel chuẩn 100% mẫu người dùng (Phiên bản v1.2):
    - Nếu split_by_year=True: Mỗi năm ra 1 file Excel riêng biệt (VD: QTT_TNCN_2024.xlsx có 4 sheet TK, 05-1, 05-2, 05-3).
    - Nếu split_by_tax_type=True: Mỗi nhóm sắc thuế ra 1 file Excel riêng biệt.
    - Ngược lại: Xuất vào 1 file tổng hợp đầy đủ các sheet.
    """
    created_files: List[str] = []
    valid = [d for d in declarations if not d.error]
    out_dir = os.path.dirname(os.path.abspath(output_path))
    os.makedirs(out_dir, exist_ok=True)
    base_name, ext = os.path.splitext(output_path)
    if not ext:
        ext = ".xlsx"

    # CHẾ ĐỘ 1: XUẤT THEO NĂM (Mỗi năm ra 1 file Excel)
    if split_by_year:
        years_set = {d.year for d in valid if d.year is not None}
        for yr in sorted(years_set):
            yr_decls = [d for d in valid if d.year == yr]
            tncn_decls = [d for d in yr_decls if d.tax_type == "TNCN"]

            if tncn_decls:
                latest_tncn = filter_latest_per_year(tncn_decls)
                if latest_tncn:
                    tncn_file = os.path.join(out_dir, f"QTT_TNCN_{yr}{ext}")
                    write_single_year_tncn_file(tncn_file, latest_tncn[-1])
                    created_files.append(tncn_file)

            other_decls = [d for d in yr_decls if d.tax_type != "TNCN"]
            if other_decls:
                yr_wb = openpyxl.Workbook()
                yr_wb.remove(yr_wb.active)
                existing: Set[str] = set()

                tndn_items = [d for d in other_decls if d.tax_type == "TNDN"]
                if tndn_items:
                    write_sheet_kktndn(yr_wb, filter_latest_per_year(tndn_items), existing)
                    write_sheet_kqkd(yr_wb, filter_latest_per_year(tndn_items), existing)

                bctc_items = [d for d in other_decls if d.tax_type == "BCTC"]
                if bctc_items:
                    write_sheet_bctc_cdkt(yr_wb, filter_latest_per_year(bctc_items), existing)

                gtgt_items = [d for d in other_decls if d.tax_type == "GTGT"]
                if gtgt_items:
                    write_gtgt_timeline_sheet(yr_wb, gtgt_items, existing, sheet_name="So Thue GTGT TH")
                    write_gtgt_cuoi_sheet(yr_wb, gtgt_items, existing, sheet_name="So Thue GTGT Cuoi")

                if gtgt_items or tndn_items:
                    write_sheet_chenh_lech_gtgt_tndn(yr_wb, existing)
                    write_yearly_verification_sheets(yr_wb, existing)

                khac_items = [d for d in other_decls if d.tax_type == "Khác"]
                if khac_items:
                    write_other_declarations_sheet(yr_wb, khac_items, existing, sheet_name="To_Khai_Khac")

                if include_appendices:
                    write_appendix_sheets(yr_wb, other_decls, existing)

                yr_file = os.path.join(out_dir, f"ToKhai_{yr}{ext}")
                yr_wb.save(yr_file)
                created_files.append(yr_file)

        return created_files

    # CHẾ ĐỘ 2: TÁCH THEO NHÓM SẮC THUẾ
    if split_by_tax_type:
        tax_types = sorted({d.tax_type for d in valid})
        for t_type in tax_types:
            sub_decls = [d for d in valid if d.tax_type == t_type]
            if not sub_decls:
                continue
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            _populate_workbook_sheets(wb, sub_decls, sub_decls, include_appendices=include_appendices)
            sub_path = f"{base_name}_{t_type}{ext}"
            wb.save(sub_path)
            created_files.append(sub_path)
        return created_files

    # CHẾ ĐỘ 3: FILE GỘP TỔNG HỢP
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    _populate_workbook_sheets(wb, declarations, valid, include_appendices=include_appendices, cutoff_date=cutoff_date)
    wb.save(output_path)
    created_files.append(output_path)

    return created_files
