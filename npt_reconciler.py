# -*- coding: utf-8 -*-
"""
Module đối chiếu Người phụ thuộc (NPT) giữa Bảng kê 05-3/BK-QTT-TNCN và File kết xuất NPT từ hệ thống TMS.

Hỗ trợ đọc trực tiếp file Excel (.xlsx) hoặc CSV (.csv) kết xuất từ chức năng
"HIỂN THỊ DANH SÁCH ĐỘNG" trên TMS (tự động xử lý khoảng trắng không ngắt \xa0,
2 cột trùng tên 'Tháng' = Từ tháng & Đến tháng, lịch sử nhiều dòng báo tăng/báo giảm
cho cùng 1 cặp MST NNT - MST NPT, số CMND/CCCD bị mất số 0 đầu, Hộ chiếu người nước ngoài).

Kết quả tích hợp vào workbook TNCN:
1. Bổ sung các cột đối chiếu (L..S) ngay bên phải sheet '05-3' với công thức Excel động.
2. Tạo sheet 'DoiChieu_NPT' tổng hợp KPI (khớp / chưa đăng ký / vượt tháng / trùng NPT / ngừng HĐ)
   và danh sách riêng các trường hợp chênh lệch cần xử lý kèm số tiền giảm trừ loại trừ.
3. Tạo sheet 'DS_NPT_HeThong' lưu bản làm sạch dữ liệu NPT hệ thống TMS để tiện tra cứu.
"""

from __future__ import annotations

import csv
import datetime
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


SHEET_NPT_SUMMARY = "DoiChieu_NPT"
SHEET_NPT_SYSTEM = "DS_NPT_HeThong"


def _clean_text(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).replace("\xa0", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", s).strip()


def _clean_id(val: Any) -> str:
    """Chuẩn hóa MST / CMND / CCCD / Hộ chiếu (bỏ khoảng trắng, giữ số 0 đầu nếu là chuỗi)."""
    s = _clean_text(val)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.replace(" ", "").upper()


def _norm_id_no_zero(s: str) -> str:
    """Chuẩn hóa không tính số 0 đứng đầu để khớp CCCD/CMND khi TMS làm mất số 0 đầu."""
    c = _clean_id(s)
    return c.lstrip("0") or c


def parse_ym(val: Any) -> Optional[Tuple[int, int]]:
    """
    Chuyển nhiều định dạng tháng/năm về tuple (year, month):
    - TMS integer/string: '72018' -> (2018, 7), '122019' -> (2019, 12), '12021' -> (2021, 1)
    - Standard 05-3: '01/2024', '2/2023', '2024-01', '01.2024'
    - datetime/date object
    """
    if val is None:
        return None
    if isinstance(val, (datetime.date, datetime.datetime)):
        return (val.year, val.month)

    s = _clean_text(val)
    if not s:
        return None
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]

    # Dạng MM/YYYY hoặc M/YYYY hoặc MM-YYYY hoặc MM.YYYY
    m1 = re.fullmatch(r"(\d{1,2})[/.\-](\d{4})", s)
    if m1:
        mo, yr = int(m1.group(1)), int(m1.group(2))
        if 1 <= mo <= 12 and 1990 <= yr <= 2100:
            return (yr, mo)

    # Dạng YYYY-MM hoặc YYYY/MM
    m2 = re.fullmatch(r"(\d{4})[/.\-](\d{1,2})", s)
    if m2:
        yr, mo = int(m2.group(1)), int(m2.group(2))
        if 1 <= mo <= 12 and 1990 <= yr <= 2100:
            return (yr, mo)

    # Dạng TMS liền: MYYYY (5 số) hoặc MMYYYY (6 số)
    digits = re.sub(r"\D", "", s)
    if len(digits) == 5:
        mo, yr = int(digits[0]), int(digits[1:])
        if 1 <= mo <= 9 and 1990 <= yr <= 2100:
            return (yr, mo)
    elif len(digits) == 6:
        mo, yr = int(digits[:2]), int(digits[2:])
        if 1 <= mo <= 12 and 1990 <= yr <= 2100:
            return (yr, mo)

    return None


def fmt_ym(ym: Optional[Tuple[int, int]]) -> str:
    if ym is None:
        return ""
    return f"{ym[1]:02d}/{ym[0]}"


def ym_to_index(ym: Tuple[int, int]) -> int:
    return ym[0] * 12 + (ym[1] - 1)


def _parse_date_dmy(val: Any) -> datetime.date:
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = _clean_text(val)
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return datetime.date.min


def deduction_rate_for_year(year: Optional[int]) -> int:
    """Mức giảm trừ gia cảnh người phụ thuộc (đồng/tháng) theo từng giai đoạn pháp luật."""
    if year is None:
        return 4_400_000
    if year <= 2019:
        return 3_600_000
    if year <= 2025:
        return 4_400_000
    return 6_200_000


@dataclass
class TmsNptRecord:
    ngay_tao: str
    ngay_td: str
    mst_nnt: str
    ten_nnt: str
    mst_npt: str
    ten_npt: str
    trang_thai_npt: str
    ma_trang_thai: str
    ngay_sinh: str
    cmnd: str
    ho_chieu: str
    quan_he: str
    tu_thang: Optional[Tuple[int, int]]
    den_thang: Optional[Tuple[int, int]]
    mau_tk: str
    gks_so: str
    sort_date: datetime.date


def _read_raw_rows(file_path: str) -> List[List[Any]]:
    ext = os.path.splitext(file_path)[1].lower()
    if ext in (".csv", ".txt"):
        # Tự thử utf-8-sig -> utf-8 -> cp1258 / latin-1
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with open(file_path, "r", encoding=enc, newline="") as f:
                    return [row for row in csv.reader(f)]
            except UnicodeDecodeError:
                continue
        raise ValueError(f"Không thể đọc file CSV NPT: {file_path}")

    wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return rows


def parse_tms_npt_file(file_path: str) -> List[TmsNptRecord]:
    """
    Đọc file kết xuất danh sách Người phụ thuộc từ hệ thống TMS (.xlsx hoặc .csv).
    Tự động dò dòng tiêu đề chứa 'MST NNT' và 'MST NPT', xử lý 2 cột trùng tên 'Tháng'.
    """
    raw_rows = _read_raw_rows(file_path)
    if not raw_rows:
        return []

    hdr_idx: Optional[int] = None
    col_map: Dict[str, int] = {}
    thang_cols: List[int] = []

    for r_idx, row in enumerate(raw_rows[:30]):
        norm_cells = [_clean_text(c).lower() for c in row]
        if any("mst nnt" in c for c in norm_cells) and any("mst npt" in c for c in norm_cells):
            hdr_idx = r_idx
            for c_idx, h in enumerate(norm_cells):
                if h == "ngày tạo" or h == "ngày tạo" or ("ngày tạo" in h and "ngay_tao" not in col_map):
                    col_map.setdefault("ngay_tao", c_idx)
                elif h.startswith("ngày tđ") or "ngày thay đổi" in h:
                    col_map.setdefault("ngay_td", c_idx)
                elif "mst nnt" in h:
                    col_map.setdefault("mst_nnt", c_idx)
                elif "tên nnt" in h:
                    col_map.setdefault("ten_nnt", c_idx)
                elif "mst npt" in h:
                    col_map.setdefault("mst_npt", c_idx)
                elif "tên npt" in h:
                    col_map.setdefault("ten_npt", c_idx)
                elif "trạng thái npt" in h:
                    col_map.setdefault("trang_thai_npt", c_idx)
                elif "ngày sinh" in h:
                    col_map.setdefault("ngay_sinh", c_idx)
                elif h in ("cmnd", "cccd", "cmnd/cccd", "số giấy tờ"):
                    col_map.setdefault("cmnd", c_idx)
                elif "quan hệ" in h:
                    col_map.setdefault("quan_he", c_idx)
                elif h in ("tháng", "tháng", "từ tháng", "đến tháng"):
                    thang_cols.append(c_idx)
                elif "mẫu bk" in h or "mẫu tk" in h:
                    col_map.setdefault("mau_tk", c_idx)
                elif "hộ chiếu" in h:
                    col_map.setdefault("ho_chieu", c_idx)
                elif "gks/số" in h or "số gks" in h:
                    col_map.setdefault("gks_so", c_idx)
                elif h.startswith("mã trạng t"):
                    col_map.setdefault("ma_trang_thai", c_idx)
            break

    if hdr_idx is None:
        raise ValueError(
            "Không tìm thấy dòng tiêu đề chứa 'MST NNT' và 'MST NPT' trong file dữ liệu NPT hệ thống."
        )

    if len(thang_cols) >= 1:
        col_map["tu_thang"] = thang_cols[0]
    if len(thang_cols) >= 2:
        col_map["den_thang"] = thang_cols[1]

    def _get(row: Sequence[Any], key: str) -> Any:
        idx = col_map.get(key)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    records: List[TmsNptRecord] = []
    for row in raw_rows[hdr_idx + 1:]:
        mst_nnt = _clean_id(_get(row, "mst_nnt"))
        mst_npt = _clean_id(_get(row, "mst_npt"))
        ten_npt = _clean_text(_get(row, "ten_npt"))
        if not mst_nnt and not mst_npt and not ten_npt:
            continue

        ngay_tao_s = _clean_text(_get(row, "ngay_tao"))
        ngay_td_s = _clean_text(_get(row, "ngay_td"))
        d_tao = _parse_date_dmy(ngay_tao_s)
        d_td = _parse_date_dmy(ngay_td_s)
        sort_d = max(d_tao, d_td)

        records.append(
            TmsNptRecord(
                ngay_tao=ngay_tao_s,
                ngay_td=ngay_td_s,
                mst_nnt=mst_nnt,
                ten_nnt=_clean_text(_get(row, "ten_nnt")),
                mst_npt=mst_npt,
                ten_npt=ten_npt,
                trang_thai_npt=_clean_text(_get(row, "trang_thai_npt")),
                ma_trang_thai=_clean_text(_get(row, "ma_trang_thai")),
                ngay_sinh=_clean_text(_get(row, "ngay_sinh")),
                cmnd=_clean_id(_get(row, "cmnd")),
                ho_chieu=_clean_id(_get(row, "ho_chieu")),
                quan_he=_clean_text(_get(row, "quan_he")),
                tu_thang=parse_ym(_get(row, "tu_thang")),
                den_thang=parse_ym(_get(row, "den_thang")),
                mau_tk=_clean_text(_get(row, "mau_tk")),
                gks_so=_clean_text(_get(row, "gks_so")),
                sort_date=sort_d,
            )
        )

    return records


def _resolve_effective_record(recs: List[TmsNptRecord]) -> TmsNptRecord:
    """
    Hợp nhất lịch sử nhiều dòng TMS cho cùng 1 cặp (MST NNT, MST NPT):
    - Ưu tiên các dòng có 'Từ tháng' (bỏ qua các dòng 20/DK-TCT chỉ cập nhật địa chỉ không có thời gian).
    - Nếu cùng 'Từ tháng' mà có bản cập nhật mới hơn (hoặc có báo giảm 'Đến tháng'), ưu tiên bản mới nhất/có Đến tháng.
    """
    with_start = [r for r in recs if r.tu_thang is not None]
    pool = with_start if with_start else recs

    def _rank(r: TmsNptRecord) -> Tuple[datetime.date, int, int]:
        has_end = 1 if r.den_thang is not None else 0
        start_idx = ym_to_index(r.tu_thang) if r.tu_thang else -1
        return (r.sort_date, has_end, start_idx)

    pool_sorted = sorted(pool, key=_rank)
    best = pool_sorted[-1]

    # Nếu bản mới nhất có trạng thái trống nhưng bản khác có trạng thái/CMND/hộ chiếu -> bổ sung
    if not best.trang_thai_npt or not best.ma_trang_thai:
        for r in reversed(recs):
            if r.trang_thai_npt and not best.trang_thai_npt:
                best.trang_thai_npt = r.trang_thai_npt
            if r.ma_trang_thai and not best.ma_trang_thai:
                best.ma_trang_thai = r.ma_trang_thai
    return best


def _overlap_months(
    decl_from: Optional[Tuple[int, int]],
    decl_to: Optional[Tuple[int, int]],
    tms_from: Optional[Tuple[int, int]],
    tms_to: Optional[Tuple[int, int]],
) -> int:
    """Tính số tháng giao nhau hợp lệ giữa khoảng kê khai 05-3 và khoảng hiệu lực TMS."""
    if decl_from is None or decl_to is None or tms_from is None:
        return 0
    d1, d2 = ym_to_index(decl_from), ym_to_index(decl_to)
    if d2 < d1:
        return 0
    t1 = ym_to_index(tms_from)
    t2 = ym_to_index(tms_to) if tms_to is not None else ym_to_index((9999, 12))
    if t2 < t1:
        return 0
    lo = max(d1, t1)
    hi = min(d2, t2)
    return max(0, hi - lo + 1)


def _detect_05_3_year(ws_053: openpyxl.worksheet.worksheet.Worksheet, wb: openpyxl.Workbook) -> Optional[int]:
    """Dò năm quyết toán từ cột Từ tháng/Đến tháng trên 05-3 hoặc tên sheet."""
    for r in range(4, min(ws_053.max_row + 1, 200)):
        ym = parse_ym(ws_053.cell(r, 10).value) or parse_ym(ws_053.cell(r, 11).value)
        if ym is not None:
            return ym[0]
    for s in wb.sheetnames:
        if re.fullmatch(r"20\d{2}", s.strip()):
            return int(s.strip())
    return None


def _write_system_npt_sheet(
    wb: openpyxl.Workbook,
    records: List[TmsNptRecord],
    sheet_name: str = SHEET_NPT_SYSTEM,
) -> openpyxl.worksheet.worksheet.Worksheet:
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(title=sheet_name)
    ws.views.sheetView[0].showGridLines = True

    f_hdr = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    thin = Side(style="thin", color="BFBFBF")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center")

    headers = [
        "STT", "Ngày tạo", "Ngày TĐ", "MST NNT", "Tên NNT", "MST NPT", "Tên NPT",
        "Ngày sinh", "CMND/CCCD", "Hộ chiếu", "Quan hệ", "Từ tháng (TMS)", "Đến tháng (TMS)",
        "Trạng thái MST NPT", "Mã TT", "Mẫu BK/TK", "GKS/Số",
    ]
    ws.row_dimensions[1].height = 26
    for c_idx, h in enumerate(headers, start=1):
        c = ws.cell(1, c_idx, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    text_cols = {4, 6, 9, 10, 15, 17}
    for idx, rec in enumerate(records, start=1):
        r = idx + 1
        vals = [
            idx, rec.ngay_tao, rec.ngay_td, rec.mst_nnt, rec.ten_nnt, rec.mst_npt, rec.ten_npt,
            rec.ngay_sinh, rec.cmnd, rec.ho_chieu, rec.quan_he,
            fmt_ym(rec.tu_thang), fmt_ym(rec.den_thang) if rec.den_thang else "",
            rec.trang_thai_npt, rec.ma_trang_thai, rec.mau_tk, rec.gks_so,
        ]
        for c_idx, v in enumerate(vals, start=1):
            cell = ws.cell(r, c_idx, v)
            cell.font = f_reg
            cell.border = b_thin
            if c_idx in text_cols:
                cell.number_format = "@"
                cell.alignment = al_l
            elif c_idx in (1, 2, 3, 8, 11, 12, 13, 16):
                cell.alignment = al_c
            else:
                cell.alignment = al_l

    widths = [6, 12, 12, 14, 24, 14, 24, 12, 15, 13, 11, 13, 13, 34, 8, 11, 13]
    for c_idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(c_idx)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(2, len(records) + 1)}"
    return ws


def reconcile_05_3_in_workbook(
    wb: openpyxl.Workbook,
    tms_npt_path: str,
    tax_year: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Thực hiện đối chiếu bảng kê '05-3' trong workbook `wb` với file NPT kết xuất từ TMS:
    - Chép dữ liệu sạch vào sheet 'DS_NPT_HeThong'
    - Thêm 8 cột đối chiếu (L..S) ngay bên phải sheet '05-3'
    - Tạo sheet 'DoiChieu_NPT' tổng hợp KPI & liệt kê riêng các dòng chênh lệch cần xử lý
    Trả về dict thống kê {total_rows, matched, issues, total_diff_months, total_disallowed_amount}.
    """
    if "05-3" not in wb.sheetnames:
        raise ValueError("Workbook không có sheet '05-3' để đối chiếu Người phụ thuộc.")

    tms_records = parse_tms_npt_file(tms_npt_path)
    _write_system_npt_sheet(wb, tms_records, SHEET_NPT_SYSTEM)

    # Xây dựng chỉ mục tra cứu từ TMS
    by_pair_mst: Dict[Tuple[str, str], List[TmsNptRecord]] = {}
    by_pair_doc: Dict[Tuple[str, str], List[TmsNptRecord]] = {}
    by_npt_mst: Dict[str, List[TmsNptRecord]] = {}

    for rec in tms_records:
        if rec.mst_nnt and rec.mst_npt:
            by_pair_mst.setdefault((rec.mst_nnt, rec.mst_npt), []).append(rec)
        if rec.mst_nnt:
            for doc in (rec.cmnd, rec.ho_chieu):
                if doc:
                    by_pair_doc.setdefault((rec.mst_nnt, _norm_id_no_zero(doc)), []).append(rec)
        if rec.mst_npt:
            by_npt_mst.setdefault(rec.mst_npt, []).append(rec)

    resolved_pair_mst = {k: _resolve_effective_record(v) for k, v in by_pair_mst.items()}
    resolved_pair_doc = {k: _resolve_effective_record(v) for k, v in by_pair_doc.items()}

    ws_053 = wb["05-3"]
    detected_year = tax_year or _detect_05_3_year(ws_053, wb)
    default_rate = deduction_rate_for_year(detected_year)

    # Quét trước toàn bộ dòng dữ liệu trên 05-3 để phát hiện trùng NPT giữa nhiều NNT trong cùng năm
    npt_to_nnts_in_053: Dict[Tuple[Optional[int], str], Set[str]] = {}
    data_rows: List[Tuple[int, Optional[int]]] = []
    current_yr = detected_year

    for r in range(4, ws_053.max_row + 1):
        val_a = ws_053.cell(r, 1).value
        val_b = _clean_text(ws_053.cell(r, 2).value)
        val_c = _clean_id(ws_053.cell(r, 3).value)
        val_d = _clean_text(ws_053.cell(r, 4).value)

        # Bỏ qua dòng tiêu đề năm hoặc dòng tổng cộng (trường hợp sheet 05-3 gộp nhiều năm)
        if not val_c and not val_d:
            m_yr = re.search(r"NĂM\s+(20\d{2})", (str(val_a or "") + " " + val_b).upper())
            if m_yr:
                current_yr = int(m_yr.group(1))
            continue
        if "TỔNG CỘNG" in (str(val_a or "") + " " + val_b).upper():
            continue

        ym_f = parse_ym(ws_053.cell(r, 10).value)
        ym_t = parse_ym(ws_053.cell(r, 11).value)
        row_yr = (ym_f[0] if ym_f else (ym_t[0] if ym_t else current_yr))
        data_rows.append((r, row_yr))

        npt_mst = _clean_id(ws_053.cell(r, 6).value)
        npt_doc = _norm_id_no_zero(_clean_id(ws_053.cell(r, 8).value))
        npt_key = npt_mst or npt_doc
        if npt_key and val_c:
            npt_to_nnts_in_053.setdefault((row_yr, npt_key), set()).add(val_c)

    # ===== GHI TIÊU ĐỀ 8 CỘT ĐỐI CHIẾU (L..S) TRÊN SHEET '05-3' =====
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_hdr_code = Font(name="Calibri", size=9, italic=True, color="1F4E78")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_code = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_ok = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_err = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    NUM = "#,##0"

    # Merge hàng 1 cho nhóm đối chiếu
    ws_053.merge_cells("L1:S1")
    ws_053["L1"] = "ĐỐI CHIẾU DỮ LIỆU NGƯỜI PHỤ THUỘC VỚI HỆ THỐNG TMS"
    recon_headers = [
        (12, "Số tháng\n05-3", "[DC1]"),
        (13, "Từ tháng\n(TMS)", "[DC2]"),
        (14, "Đến tháng\n(TMS)", "[DC3]"),
        (15, "Trạng thái MST NPT\n(TMS)", "[DC4]"),
        (16, "Số tháng\nhợp lệ (TMS)", "[DC5]"),
        (17, "Số tháng\nchênh lệch", "[DC6]"),
        (18, "Tiền giảm trừ\nloại trừ (đồng)", "[DC7]"),
        (19, "Kết quả đối chiếu", "[DC8]"),
    ]
    for cc in range(12, 20):
        c1 = ws_053.cell(1, cc)
        c1.font = f_hdr
        c1.fill = fill_hdr
        c1.alignment = al_c
        c1.border = b_thin
    for cc, title, code in recon_headers:
        c2 = ws_053.cell(2, cc, title)
        c2.font = f_hdr
        c2.fill = fill_hdr
        c2.alignment = al_c
        c2.border = b_thin

        c3 = ws_053.cell(3, cc, code)
        c3.font = f_hdr_code
        c3.fill = fill_code
        c3.alignment = al_c
        c3.border = b_thin

    # ===== ĐỐI CHIẾU TỪNG DÒNG TRÊN '05-3' =====
    issue_rows_detail: List[Dict[str, Any]] = []
    matched_count = 0
    issue_count = 0
    total_diff_months = 0

    for r, row_yr in data_rows:
        ten_nnt = _clean_text(ws_053.cell(r, 2).value)
        mst_nnt = _clean_id(ws_053.cell(r, 3).value)
        ten_npt = _clean_text(ws_053.cell(r, 4).value)
        mst_npt = _clean_id(ws_053.cell(r, 6).value)
        so_gt = _clean_id(ws_053.cell(r, 8).value)
        quan_he = _clean_text(ws_053.cell(r, 9).value)
        tu_thang_decl = parse_ym(ws_053.cell(r, 10).value)
        den_thang_decl = parse_ym(ws_053.cell(r, 11).value)

        if tu_thang_decl and den_thang_decl and ym_to_index(den_thang_decl) >= ym_to_index(tu_thang_decl):
            decl_months = ym_to_index(den_thang_decl) - ym_to_index(tu_thang_decl) + 1
        else:
            decl_months = 0

        # Tra cứu trên hệ thống TMS
        matched_rec: Optional[TmsNptRecord] = None
        if mst_nnt and mst_npt and (mst_nnt, mst_npt) in resolved_pair_mst:
            matched_rec = resolved_pair_mst[(mst_nnt, mst_npt)]
        elif mst_nnt and so_gt and (mst_nnt, _norm_id_no_zero(so_gt)) in resolved_pair_doc:
            matched_rec = resolved_pair_doc[(mst_nnt, _norm_id_no_zero(so_gt))]

        # Kiểm tra trùng NPT giữa nhiều NNT trên 05-3
        npt_key = mst_npt or _norm_id_no_zero(so_gt)
        is_dup_053 = bool(npt_key and len(npt_to_nnts_in_053.get((row_yr, npt_key), set())) > 1)

        tms_from_str = ""
        tms_to_str = ""
        tms_status_str = ""
        valid_months = 0
        reasons: List[str] = []

        if matched_rec is None:
            # Kiểm tra xem MST NPT có đăng ký cho NNT khác trên TMS không
            if mst_npt and mst_npt in by_npt_mst:
                other_nnts = sorted({x.mst_nnt for x in by_npt_mst[mst_npt] if x.mst_nnt != mst_nnt})
                reasons.append(f"ĐĂNG KÝ CHO NNT KHÁC ({', '.join(other_nnts[:2])})")
            else:
                reasons.append("CHƯA ĐĂNG KÝ TRÊN HỆ THỐNG")
            valid_months = 0
        else:
            tms_from_str = fmt_ym(matched_rec.tu_thang)
            tms_to_str = fmt_ym(matched_rec.den_thang) if matched_rec.den_thang else "Còn hiệu lực"
            tms_status_str = (
                f"{matched_rec.ma_trang_thai} - {matched_rec.trang_thai_npt}"
                if matched_rec.ma_trang_thai and matched_rec.trang_thai_npt
                else (matched_rec.trang_thai_npt or matched_rec.ma_trang_thai)
            )

            # Kiểm tra trạng thái MST NPT bất thường
            if matched_rec.ma_trang_thai and matched_rec.ma_trang_thai != "00":
                reasons.append(f"MST NPT NGỪNG HĐ (Mã {matched_rec.ma_trang_thai})")

            if matched_rec.tu_thang is None:
                reasons.append("CẦN KIỂM TRA (TMS THIẾU TỪ THÁNG)")
                valid_months = 0
            else:
                valid_months = _overlap_months(tu_thang_decl, den_thang_decl, matched_rec.tu_thang, matched_rec.den_thang)
                if valid_months < decl_months:
                    if valid_months == 0:
                        reasons.append("NGOÀI THỜI GIAN ĐĂNG KÝ TMS")
                    else:
                        reasons.append(f"VƯỢT THỜI GIAN ĐĂNG KÝ ({decl_months - valid_months} tháng)")

        if is_dup_053:
            reasons.append("TRÙNG NPT GIỮA NHIỀU NNT TRÊN 05-3")

        diff_months = max(0, decl_months - valid_months)
        total_diff_months += diff_months

        if not reasons:
            status_text = "KHỚP"
            row_fill = fill_ok
            matched_count += 1
        else:
            status_text = "; ".join(reasons)
            row_fill = fill_err if (diff_months > 0 or is_dup_053 or matched_rec is None) else fill_warn
            issue_count += 1
            issue_rows_detail.append(
                {
                    "row_053": r,
                    "year": row_yr,
                    "ten_nnt": ten_nnt,
                    "mst_nnt": mst_nnt,
                    "ten_npt": ten_npt,
                    "mst_npt": mst_npt,
                    "so_gt": so_gt,
                    "quan_he": quan_he,
                    "decl_from": fmt_ym(tu_thang_decl),
                    "decl_to": fmt_ym(den_thang_decl),
                    "decl_months": decl_months,
                    "tms_from": tms_from_str,
                    "tms_to": tms_to_str,
                    "tms_status": tms_status_str,
                    "valid_months": valid_months,
                    "diff_months": diff_months,
                    "status": status_text,
                }
            )

        # Ghi giá trị & công thức vào L..S trên sheet '05-3'
        ws_053.cell(r, 12, decl_months).alignment = al_c
        ws_053.cell(r, 13, tms_from_str).alignment = al_c
        ws_053.cell(r, 14, tms_to_str).alignment = al_c
        ws_053.cell(r, 15, tms_status_str).alignment = al_l
        ws_053.cell(r, 16, valid_months).alignment = al_c
        ws_053.cell(r, 17, f"=MAX(0,L{r}-P{r})").alignment = al_c
        ws_053.cell(r, 18, f"=Q{r}*'{SHEET_NPT_SUMMARY}'!$C$4").alignment = al_r
        ws_053.cell(r, 18).number_format = NUM
        ws_053.cell(r, 19, status_text).alignment = al_l

        for cc in range(12, 20):
            cell = ws_053.cell(r, cc)
            cell.font = f_bold if cc == 19 else f_reg
            cell.border = b_thin
            if cc in (17, 18, 19):
                cell.fill = row_fill

    # Độ rộng các cột đối chiếu L..S
    recon_widths = {12: 11, 13: 12, 14: 13, 15: 28, 16: 12, 17: 12, 18: 18, 19: 34}
    for cc, w in recon_widths.items():
        ws_053.column_dimensions[get_column_letter(cc)].width = w

    # ===== TẠO SHEET TỔNG HỢP 'DoiChieu_NPT' =====
    if SHEET_NPT_SUMMARY in wb.sheetnames:
        del wb[SHEET_NPT_SUMMARY]
    ws_dc = wb.create_sheet(title=SHEET_NPT_SUMMARY)
    ws_dc.views.sheetView[0].showGridLines = True

    f_title = Font(name="Times New Roman", size=13, bold=True, color="1F4E78")
    f_sub = Font(name="Times New Roman", size=10, italic=True, color="333333")
    f_sec = Font(name="Times New Roman", size=11, bold=True, color="1F4E78")
    f_tb = Font(name="Times New Roman", size=10, bold=True)
    f_tr = Font(name="Times New Roman", size=10)
    fill_ctrl = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    ws_dc.merge_cells("A1:N1")
    ws_dc["A1"] = "BẢNG TỔNG HỢP ĐỐI CHIẾU NGƯỜI PHỤ THUỘC (05-3/BK-QTT-TNCN vs HỆ THỐNG TMS)"
    ws_dc["A1"].font = f_title
    ws_dc["A1"].alignment = al_c
    ws_dc.row_dimensions[1].height = 24

    yr_txt = f"Năm {detected_year}" if detected_year else "Kỳ quyết toán"
    ws_dc.merge_cells("A2:N2")
    ws_dc["A2"] = f"({yr_txt} | Đơn vị tính: Đồng | Nguồn đối chiếu: sheet '05-3' và sheet '{SHEET_NPT_SYSTEM}')"
    ws_dc["A2"].font = f_sub
    ws_dc["A2"].alignment = al_c
    ws_dc.row_dimensions[2].height = 18

    last_053_row = max([r for r, _ in data_rows], default=4)
    kpi_rows = [
        (4, "Mức giảm trừ gia cảnh NPT (đồng/tháng — sửa ô vàng để cập nhật toàn bộ):", default_rate, NUM, True),
        (5, "Tổng số dòng NPT kê khai trên bảng kê 05-3:", f"=COUNTA('05-3'!S4:S{last_053_row})", NUM, False),
        (6, "Số dòng KHỚP hoàn toàn với hệ thống TMS:", f'=COUNTIF(\'05-3\'!S4:S{last_053_row},"KHỚP")', NUM, False),
        (7, "Số dòng PHÁT HIỆN CHÊNH LỆCH / CẦN KIỂM TRA:", f'=COUNTIF(\'05-3\'!S4:S{last_053_row},"<>KHỚP")-COUNTBLANK(\'05-3\'!S4:S{last_053_row})', NUM, False),
        (8, "Tổng số tháng giảm trừ chênh lệch (không đủ điều kiện):", f"=SUM('05-3'!Q4:Q{last_053_row})", NUM, False),
        (9, "Tổng số tiền giảm trừ NPT không đủ điều kiện (đồng):", f"=SUM('05-3'!R4:R{last_053_row})", NUM, False),
    ]
    for r_k, lbl, val, nfmt, is_inp in kpi_rows:
        ws_dc.row_dimensions[r_k].height = 20
        ws_dc.merge_cells(start_row=r_k, start_column=1, end_row=r_k, end_column=2)
        ws_dc.cell(r_k, 1, lbl).font = f_tb if (is_inp or r_k == 9) else f_tr
        ws_dc.cell(r_k, 1).alignment = al_l
        ws_dc.cell(r_k, 1).fill = fill_ctrl
        ws_dc.cell(r_k, 2).fill = fill_ctrl
        ws_dc.cell(r_k, 1).border = b_thin
        ws_dc.cell(r_k, 2).border = b_thin

        c_v = ws_dc.cell(r_k, 3, val)
        c_v.font = f_tb
        c_v.number_format = nfmt
        c_v.alignment = al_r
        c_v.fill = fill_input if is_inp else (fill_err if r_k in (7, 8, 9) else fill_ok)
        c_v.border = b_thin

    # Bảng chi tiết các dòng có chênh lệch / cảnh báo
    r_tbl_title = 11
    ws_dc.merge_cells(start_row=r_tbl_title, start_column=1, end_row=r_tbl_title, end_column=14)
    ws_dc.cell(r_tbl_title, 1, "CHI TIẾT CÁC TRƯỜNG HỢP CHÊNH LỆCH / CẦN KIỂM TRA HỒ SƠ NGƯỜI PHỤ THUỘC").font = f_sec
    ws_dc.row_dimensions[r_tbl_title].height = 22

    r_hdr = 12
    ws_dc.row_dimensions[r_hdr].height = 26
    dc_headers = [
        "STT", "Dòng 05-3", "Họ tên NNT", "MST NNT", "Họ tên NPT", "MST NPT",
        "Số giấy tờ", "Từ tháng\n(05-3)", "Đến tháng\n(05-3)", "Từ tháng\n(TMS)",
        "Đến tháng\n(TMS)", "Số tháng\nchênh lệch", "Tiền giảm trừ\nloại trừ (đồng)", "Nội dung chênh lệch / Cảnh báo",
    ]
    for cc, h in enumerate(dc_headers, start=1):
        c = ws_dc.cell(r_hdr, cc, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    if not issue_rows_detail:
        ws_dc.merge_cells("A13:N13")
        ws_dc["A13"] = "Toàn bộ Người phụ thuộc trên bảng kê 05-3 đều KHỚP 100% với dữ liệu hệ thống TMS."
        ws_dc["A13"].font = f_tb
        ws_dc["A13"].alignment = al_c
        ws_dc["A13"].fill = fill_ok
        for cc in range(1, 15):
            ws_dc.cell(13, cc).border = b_thin
    else:
        for i_idx, item in enumerate(issue_rows_detail, start=1):
            rr = r_hdr + i_idx
            r_src = item["row_053"]
            ws_dc.row_dimensions[rr].height = 20
            row_vals = [
                i_idx,
                r_src,
                item["ten_nnt"],
                item["mst_nnt"],
                item["ten_npt"],
                item["mst_npt"],
                item["so_gt"],
                item["decl_from"],
                item["decl_to"],
                item["tms_from"],
                item["tms_to"],
                f"='05-3'!Q{r_src}",
                f"='05-3'!R{r_src}",
                item["status"],
            ]
            for cc, v in enumerate(row_vals, start=1):
                c = ws_dc.cell(rr, cc, v)
                c.font = f_tr
                c.border = b_thin
                if cc in (4, 6, 7):
                    c.number_format = "@"
                    c.alignment = al_l
                elif cc in (1, 2, 8, 9, 10, 11, 12):
                    c.alignment = al_c
                elif cc == 13:
                    c.number_format = NUM
                    c.alignment = al_r
                else:
                    c.alignment = al_l
                if cc in (12, 13, 14):
                    c.fill = fill_err if item["diff_months"] > 0 else fill_warn

        r_tot = r_hdr + len(issue_rows_detail) + 1
        ws_dc.row_dimensions[r_tot].height = 22
        ws_dc.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=11)
        ws_dc.cell(r_tot, 1, "TỔNG CỘNG").alignment = al_c
        ws_dc.cell(r_tot, 12, f"=SUM(L{r_hdr + 1}:L{r_tot - 1})").alignment = al_c
        ws_dc.cell(r_tot, 13, f"=SUM(M{r_hdr + 1}:M{r_tot - 1})").alignment = al_r
        ws_dc.cell(r_tot, 13).number_format = NUM
        for cc in range(1, 15):
            c = ws_dc.cell(r_tot, cc)
            c.font = f_tb
            c.fill = fill_code
            c.border = b_thin

    dc_widths = [6, 10, 24, 14, 24, 14, 15, 11, 11, 11, 12, 12, 18, 36]
    for cc, w in enumerate(dc_widths, start=1):
        ws_dc.column_dimensions[get_column_letter(cc)].width = w
    ws_dc.freeze_panes = "A13"

    return {
        "total_rows": len(data_rows),
        "matched": matched_count,
        "issues": issue_count,
        "total_diff_months": total_diff_months,
        "total_disallowed_amount": total_diff_months * default_rate,
        "rate_per_month": default_rate,
    }


def reconcile_npt_file(
    tncn_excel_path: str,
    tms_npt_path: str,
    tax_year: Optional[int] = None,
    output_path: Optional[str] = None,
) -> Dict[str, Any]:
    """Mở file Excel TNCN (QTT_TNCN_YYYY.xlsx hoặc file tổng hợp có sheet 05-3), đối chiếu NPT và lưu lại."""
    wb = openpyxl.load_workbook(tncn_excel_path)
    stats = reconcile_05_3_in_workbook(wb, tms_npt_path, tax_year=tax_year)
    dest = output_path or tncn_excel_path
    wb.save(dest)
    stats["output_path"] = dest
    return stats
