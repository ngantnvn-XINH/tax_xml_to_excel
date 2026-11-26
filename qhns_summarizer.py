# -*- coding: utf-8 -*-
"""
Module tổng hợp sheet Tổng quan (TQ), đối chiếu Tờ khai vs Tổng quan trên các sheet Năm (2020..2025)
và lập biểu Quan Hệ Ngân Sách (QHNS) chuẩn nghiệp vụ thanh tra - kiểm tra thuế.

Tích hợp trực tiếp vào ứng dụng Desktop (app.py) cho cả 2 chế độ:
1. Chạy trọn gói: Xuất XML -> Sheet năm -> Tự động ghép TQ, đối chiếu F/G & sinh QHNS.
2. Chạy bổ sung: Chọn file Excel tổng hợp đã xuất + file Sổ Tổng Quan (TQ) -> Tự động cập nhật TQ, đối chiếu & sinh QHNS.
"""
from __future__ import annotations

import calendar
import datetime
import os
import re
from collections import defaultdict
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def parse_date(val: Any) -> Optional[datetime.datetime]:
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, datetime.date):
        return datetime.datetime(val.year, val.month, val.day)
    if not val:
        return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s[:10], fmt)
        except ValueError:
            pass
    return None


def extract_ky_year(ky_val: Any, ngay_ht: Optional[datetime.datetime] = None) -> Optional[int]:
    """
    Trích xuất năm của Kỳ báo cáo:
    - '2001'..'2012' -> 2020 (YY=20, MM=01..12)
    - '1912' -> 2019, '19CN' -> 2019, '15CN' -> 2015, '16Q1' -> 2016
    - '01/2021', 'Q1/2021', '2021' -> 2021
    """
    if ky_val is not None:
        s = str(ky_val).strip().upper()
        m_short = re.match(r"^(\d{2})(0[1-9]|1[0-2]|Q\d|CN)$", s)
        if m_short:
            return 2000 + int(m_short.group(1))
        m_full = re.search(r"(20\d{2})", s)
        if m_full:
            return int(m_full.group(1))
    if ngay_ht:
        return ngay_ht.year
    return None


def _resolve_date_to_fiscal_year(
    ref_date: datetime.date,
    fiscal_start_month: int = 1,
    fiscal_ranges: Optional[Dict[int, Tuple[datetime.date, datetime.date]]] = None,
) -> int:
    """
    Quy đổi một ngày thực tế (ref_date) sang Năm tài chính:
    1. Ưu tiên đối chiếu theo khoảng (start_date, end_date) thực tế từng năm trong fiscal_ranges
       (xử lý chính xác tuyệt đối năm tài chính gộp giải thể / năm đầu <= 15 tháng, VD 04/2023-06/2024 -> 2023).
    2. Nếu không nằm trong fiscal_ranges thì fallback theo fiscal_start_month.
    """
    if fiscal_ranges:
        for fy, (f_start, f_end) in sorted(fiscal_ranges.items()):
            if f_start <= ref_date <= f_end:
                return fy
    if fiscal_start_month <= 1:
        return ref_date.year
    return ref_date.year if ref_date.month >= fiscal_start_month else ref_date.year - 1


def extract_ky_fiscal_year(
    ky_val: Any,
    ngay_ht: Optional[datetime.datetime] = None,
    fiscal_start_month: int = 1,
    fiscal_ranges: Optional[Dict[int, Tuple[datetime.date, datetime.date]]] = None,
) -> Optional[int]:
    """
    Trích xuất NĂM TÀI CHÍNH của Kỳ báo cáo khi doanh nghiệp lệch kỳ (VD fiscal_start_month=4: 04/YYYY..03/(YYYY+1))
    hoặc có năm tài chính gộp giải thể <= 15 tháng (VD 04/2023..06/2024 trong fiscal_ranges):
    - Nếu fiscal_start_month <= 1 và không có fiscal_ranges: giống hệt extract_ky_year.
    - Nếu có fiscal_ranges (VD {2023: (2023-04-01, 2024-06-30)}):
        * '2404', '2405', '2406', '15/05/2024', '240515' -> năm tài chính 2023 (đồng nhất tờ khai & tổng quan!).
    - Nếu fiscal_start_month == 4 (fallback khi ngoài fiscal_ranges):
        * '2001', '2002', '2003', '200215', '15/02/2020', '2020-02-15', '20200215' -> năm tài chính 2019
        * '2004'..'2012', '2101'..'2103', '210215', '15/02/2021', '2021-02-15'     -> năm tài chính 2020
        * '20Q1' -> 2019; '20Q2'..'20Q4', '21Q1' -> 2020
        * '20CN', '2020' -> 2020
    """
    if fiscal_start_month <= 1 and not fiscal_ranges:
        return extract_ky_year(ky_val, ngay_ht)

    def _from_ym(yr: int, mo: int, day: int = 15) -> int:
        max_d = calendar.monthrange(yr, mo)[1]
        return _resolve_date_to_fiscal_year(
            datetime.date(yr, mo, min(max(day, 1), max_d)),
            fiscal_start_month,
            fiscal_ranges,
        )

    if isinstance(ky_val, datetime.datetime):
        return _resolve_date_to_fiscal_year(ky_val.date(), fiscal_start_month, fiscal_ranges)
    if isinstance(ky_val, datetime.date):
        return _resolve_date_to_fiscal_year(ky_val, fiscal_start_month, fiscal_ranges)
    if ky_val is not None:
        if isinstance(ky_val, float) and ky_val.is_integer():
            s = str(int(ky_val))
        else:
            s = str(ky_val).strip().upper()
            if s.endswith(".0") and s[:-2].isdigit():
                s = s[:-2]
        # 1. Định dạng YYYY-MM-DD / YYYY/MM/DD / YYYY.MM.DD (bao gồm cả chuỗi '2021-02-15 00:00:00')
        m_iso = re.search(r"(20\d{2})[/.\-](0?[1-9]|1[0-2])[/.\-](\d{1,2})", s)
        if m_iso:
            return _from_ym(int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3)))
        # 2. Định dạng DD/MM/YYYY / DD-MM-YYYY / DD.MM.YYYY (kể cả D/15/02/2021)
        m_dfull = re.search(r"(\d{1,2})[/.\-](0?[1-9]|1[0-2])[/.\-](20\d{2})", s)
        if m_dfull:
            return _from_ym(int(m_dfull.group(3)), int(m_dfull.group(2)), int(m_dfull.group(1)))
        # 3. Định dạng 8 chữ số YYYYMMDD (VD: 20210215)
        m_ymd8 = re.fullmatch(r"(20\d{2})(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])", s)
        if m_ymd8:
            return _from_ym(int(m_ymd8.group(1)), int(m_ymd8.group(2)), int(m_ymd8.group(3)))
        # 4. Định dạng 6 chữ số: phân biệt YYMMDD (VD: 210215, 200215) và YYYYMM (VD: 202102)
        if re.fullmatch(r"\d{6}", s):
            y4, m_tail = int(s[:4]), int(s[4:6])
            y2, m_mid, d_tail = 2000 + int(s[:2]), int(s[2:4]), int(s[4:6])
            is_valid_yyyymm = (2010 <= y4 <= 2039) and (1 <= m_tail <= 12)
            is_valid_yymmdd = (2010 <= y2 <= 2039) and (1 <= m_mid <= 12) and (1 <= d_tail <= 31)
            if is_valid_yymmdd and not is_valid_yyyymm:
                return _from_ym(y2, m_mid, d_tail)
            if is_valid_yyyymm and not is_valid_yymmdd:
                return _from_ym(y4, m_tail)
            if is_valid_yymmdd and is_valid_yyyymm:
                # Nếu có ngay_ht, ưu tiên cách hiểu gần với năm hạch toán hơn
                if ngay_ht and abs(y2 - ngay_ht.year) < abs(y4 - ngay_ht.year):
                    return _from_ym(y2, m_mid, d_tail)
                return _from_ym(y4, m_tail)
        # 5. Định dạng 4 chữ số YYMM (VD: 2102, 2403 hoặc có hậu tố như 2403L1 / 2403PS)
        m_ym = re.match(r"^(\d{2})(0[1-9]|1[0-2])(?:\D.*)?$", s)
        if m_ym:
            return _from_ym(2000 + int(m_ym.group(1)), int(m_ym.group(2)))
        # 6. Định dạng quý YYQ1..4
        m_yq = re.search(r"\b(\d{2})Q([1-4])\b", s)
        if m_yq:
            yr, q = 2000 + int(m_yq.group(1)), int(m_yq.group(2))
            q_mo = (q - 1) * 3 + 1
            return _from_ym(yr, q_mo)
        # 7. Định dạng MM/YYYY hoặc T02/2021 hoặc có tiền tố/hậu tố (VD: 'PS 03/2024', '03/2024 (BS)')
        m_mfull = re.search(r"(?:^|[^0-9])T?(0?[1-9]|1[0-2])[/.\-](20\d{2})(?:$|[^0-9])", s)
        if m_mfull:
            return _from_ym(int(m_mfull.group(2)), int(m_mfull.group(1)))
        # 8. Định dạng YYYY/MM hoặc YYYY-MM
        m_ymfull = re.search(r"(?:^|[^0-9])(20\d{2})[/.\-](0?[1-9]|1[0-2])(?:$|[^0-9])", s)
        if m_ymfull:
            return _from_ym(int(m_ymfull.group(1)), int(m_ymfull.group(2)))
        # 9. Định dạng Q1/2021
        m_qfull = re.search(r"Q([1-4])[/.\-](20\d{2})", s)
        if m_qfull:
            q, yr = int(m_qfull.group(1)), int(m_qfull.group(2))
            q_mo = (q - 1) * 3 + 1
            return _from_ym(yr, q_mo)
        # 10. Định dạng cả năm YYCN hoặc YYYY
        m_cn = re.match(r"^(\d{2})CN$", s)
        if m_cn:
            return 2000 + int(m_cn.group(1))
        m_full = re.search(r"(20\d{2})", s)
        if m_full:
            return int(m_full.group(1))
    if ngay_ht:
        return _resolve_date_to_fiscal_year(ngay_ht.date(), fiscal_start_month, fiscal_ranges)
    return None


def detect_fiscal_ranges_from_workbook(wb: openpyxl.Workbook) -> Dict[int, Tuple[datetime.date, datetime.date]]:
    """
    Trích xuất chính xác khoảng thời gian (start_date, end_date) từng năm tài chính từ các dòng nhãn
    'Tổng cộng năm YYYY (MM/YYYY-MM/YYYY)' trên sheet 'So Thue GTGT Cuoi (Lech ky)' hoặc 'So Thue NTNN Cuoi (Lech ky)'.
    Đảm bảo đồng nhất 100% kể cả năm tài chính gộp giải thể / năm đầu <= 15 tháng (VD 04/2023-06/2024).
    """
    gtgt_lk = None
    ntnn_lk = None
    for s in wb.sheetnames:
        sl = s.strip().lower()
        if "lech ky" in sl:
            if "ntnn" in sl:
                ntnn_lk = s
            else:
                gtgt_lk = s

    ranges: Dict[int, Tuple[datetime.date, datetime.date]] = {}
    pat = re.compile(
        r"tổng\s+cộng\s+năm\s+(20\d{2})\s*\((\d{1,2})/(20\d{2})\s*[-–]\s*(\d{1,2})/(20\d{2})\)",
        re.IGNORECASE,
    )
    for sname, label_col in ((gtgt_lk, 5), (ntnn_lk, 6)):
        if not sname:
            continue
        ws = wb[sname]
        for r in range(2, ws.max_row + 1):
            val_lbl = str(ws.cell(r, label_col).value or "").strip()
            m = pat.search(val_lbl)
            if not m:
                continue
            fy = int(m.group(1))
            m1, y1 = int(m.group(2)), int(m.group(3))
            m2, y2 = int(m.group(4)), int(m.group(5))
            if 1 <= m1 <= 12 and 1 <= m2 <= 12:
                last_d = calendar.monthrange(y2, m2)[1]
                ranges.setdefault(fy, (datetime.date(y1, m1, 1), datetime.date(y2, m2, last_d)))
    return ranges


def detect_fiscal_start_month(wb: openpyxl.Workbook) -> int:
    """
    Tự động nhận diện tháng bắt đầu năm tài chính lệch kỳ từ sheet 'So Thue GTGT Cuoi (Lech ky)'
    hoặc 'So Thue NTNN Cuoi (Lech ky)':
    - Ưu tiên tuyệt đối dòng nhãn 'Tổng cộng năm YYYY (MM/YYYY-MM/YYYY)' (chính xác 100% kể cả khi
      dòng kỳ đầu tiên trong bảng là tháng 01 thuộc năm tài chính trước!).
    - Nếu không có nhãn khoảng tháng mới fallback sang đọc kỳ đầu tiên.
    """
    gtgt_lk = None
    ntnn_lk = None
    for s in wb.sheetnames:
        sl = s.strip().lower()
        if "lech ky" in sl:
            if "ntnn" in sl:
                ntnn_lk = s
            else:
                gtgt_lk = s

    # Bước 1: Quét nhãn 'Tổng cộng năm YYYY (MM/YYYY-MM/YYYY)' trên cả GTGT và NTNN lệch kỳ
    for sname, label_col in ((gtgt_lk, 5), (ntnn_lk, 6)):
        if not sname:
            continue
        ws = wb[sname]
        for r in range(2, ws.max_row + 1):
            val_lbl = str(ws.cell(r, label_col).value or "").strip()
            m_rng = re.search(r"\((\d{1,2})/20\d{2}\s*[-–]\s*\d{1,2}/20\d{2}\)", val_lbl)
            if m_rng:
                mo = int(m_rng.group(1))
                if 2 <= mo <= 12:
                    return mo

    # Bước 2: Fallback đọc kỳ đầu tiên nếu file chưa có nhãn khoảng (MM/YYYY-MM/YYYY)
    target_sheet = gtgt_lk or ntnn_lk
    if not target_sheet:
        return 1
    ws = wb[target_sheet]
    if target_sheet == ntnn_lk:
        for r in range(2, min(ws.max_row + 1, 25)):
            ky_raw = str(ws.cell(r, 6).value or "").strip().upper()
            if not ky_raw or "tổng cộng" in ky_raw.lower():
                continue
            m_d = re.search(r"(\d{1,2})[/-](\d{1,2})[/-](20\d{2})", ky_raw)
            if m_d:
                mo = int(m_d.group(2))
                return mo if 2 <= mo <= 12 else 1
            m_m = re.match(r"^T?(0?[1-9]|1[0-2])[/-](20\d{2})$", ky_raw)
            if m_m:
                mo = int(m_m.group(1))
                return mo if 2 <= mo <= 12 else 1
        return 1
    for r in range(3, min(ws.max_row + 1, 25)):
        lbl_e = str(ws.cell(r, 5).value or "").strip().lower()
        if "tổng cộng" in lbl_e:
            continue
        ky_raw = str(ws.cell(r, 7).value or "").strip().upper()
        if not ky_raw:
            continue
        m_ym = re.match(r"^(\d{2})(0[1-9]|1[0-2])$", ky_raw)
        if m_ym:
            mo = int(m_ym.group(2))
            return mo if 2 <= mo <= 12 else 1
        m_mfull = re.match(r"^T?(0?[1-9]|1[0-2])[/-](20\d{2})$", ky_raw)
        if m_mfull:
            mo = int(m_mfull.group(1))
            return mo if 2 <= mo <= 12 else 1
        m_yq = re.match(r"^(\d{2})Q([1-4])$", ky_raw)
        if m_yq:
            mo = (int(m_yq.group(2)) - 1) * 3 + 1
            return mo if 2 <= mo <= 12 else 1
        m_qfull = re.match(r"^Q([1-4])[/-](20\d{2})$", ky_raw)
        if m_qfull:
            mo = (int(m_qfull.group(1)) - 1) * 3 + 1
            return mo if 2 <= mo <= 12 else 1
    return 1


def _norm_header(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).strip().lower()
    # Chuẩn hóa Unicode tổ hợp/dựng sẵn đơn giản cho các từ khóa cột TMS
    import unicodedata
    s = unicodedata.normalize("NFC", s)
    return re.sub(r"\s+", " ", s)


def _normalize_cqt(val: Any) -> str:
    """Chuẩn hóa mã Cơ quan thuế (VD: 7901, 7901.0, '7901 - Cục Thuế...' -> '7901')."""
    if val is None:
        return ""
    if isinstance(val, (int, float)):
        if float(val).is_integer():
            return str(int(val))
        return str(val).strip()
    s = str(val).strip()
    if not s:
        return ""
    m = re.match(r"^(\d{3,6})\b", s)
    if m:
        return m.group(1)
    return s.upper()


def split_tq_rows_by_cqt(
    tq_rows: List[Dict[str, Any]],
    target_cqt: Optional[str] = "7901",
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Tách dữ liệu Sổ Tổng Quan (TQ) theo Mã Cơ quan thuế quản lý (mặc định '7901'):
    - main_rows: Các giao dịch thuộc đúng CQT quản lý (đưa vào QHNS & đối chiếu F/G).
    - other_rows: Các giao dịch thuộc CQT khác (tách riêng sang sheet TQ_CQT_Khac).
    Nếu file nguồn không có thông tin cột CQT hoặc target_cqt để trống -> giữ toàn bộ vào main_rows.
    """
    tgt = _normalize_cqt(target_cqt)
    if not tgt:
        return list(tq_rows), []
    has_cqt_info = any(bool(r.get("cqt")) for r in tq_rows)
    if not has_cqt_info:
        return list(tq_rows), []
    main_rows = [r for r in tq_rows if r.get("cqt") == tgt]
    other_rows = [r for r in tq_rows if r.get("cqt") != tgt]
    return main_rows, other_rows


def detect_tq_layout(ws_tq) -> Tuple[int, Dict[str, int]]:
    """
    Tự động nhận diện dòng tiêu đề và chỉ số cột (1-based) của sheet Sổ Tổng Quan (TQ).
    Hỗ trợ cả định dạng TMS chuẩn (Cơ quan thuế | Kỳ báo cáo | ... | Ký hiệu giao dịch | Tiểu mục | ... | Số tiền | ... | Ngày hạch toán)
    và định dạng có cột STT ở cột A.
    """
    for r in range(1, min(ws_tq.max_row + 1, 15)):
        col_map: Dict[str, int] = {}
        for c in range(1, min(ws_tq.max_column + 1, 25)):
            h = _norm_header(ws_tq.cell(r, c).value)
            if not h:
                continue
            if "cơ quan thuế" in h or "co quan thue" in h or "mã cqt" in h or "ma cqt" in h or h == "cqt":
                col_map.setdefault("cqt", c)
            elif "kỳ báo cáo" in h or "ky bao cao" in h or "kỳ thuế" == h:
                col_map.setdefault("ky", c)
            elif "ký hiệu giao dịch" in h or "ky hieu giao dich" in h or "loại giao dịch" in h or "loại gd" in h or "mã giao dịch" in h:
                col_map.setdefault("gd", c)
            elif "tiểu mục" in h or "tieu muc" in h:
                col_map.setdefault("tm", c)
            elif h in ("số tiền", "so tien") or ("số tiền" in h and "tổng" not in h and "còn" not in h):
                col_map.setdefault("so_tien", c)
            elif "ngày hạch toán" in h or "ngay hach toan" in h:
                col_map.setdefault("ngay_ht", c)
            elif "nợ/có" in h or "no/co" in h or h in ("nợ có", "no co"):
                col_map.setdefault("no_co", c)
        if all(k in col_map for k in ("ky", "gd", "tm", "so_tien", "ngay_ht")):
            return r, col_map

    # Fallback mặc định theo file TMS chuẩn
    return 1, {"cqt": 1, "ky": 2, "gd": 4, "tm": 5, "so_tien": 7, "ngay_ht": 10}


def parse_tq_rows(ws_tq) -> List[Dict[str, Any]]:
    """Đọc toàn bộ các dòng giao dịch hợp lệ từ sheet TQ (kèm Mã Cơ quan thuế)."""
    header_row, col_map = detect_tq_layout(ws_tq)
    rows: List[Dict[str, Any]] = []
    for r in range(header_row + 1, ws_tq.max_row + 1):
        raw_st = ws_tq.cell(r, col_map["so_tien"]).value
        if isinstance(raw_st, (int, float)) and not isinstance(raw_st, bool):
            so_tien = float(raw_st)
        elif isinstance(raw_st, str) and raw_st.strip():
            s_num = raw_st.strip().replace(" ", "").replace(",", "")
            try:
                so_tien = float(s_num)
            except ValueError:
                so_tien = None
        else:
            so_tien = None
        ngay_ht = parse_date(ws_tq.cell(r, col_map["ngay_ht"]).value)
        gd = str(ws_tq.cell(r, col_map["gd"]).value or "").strip().upper()
        raw_tm = ws_tq.cell(r, col_map["tm"]).value
        if isinstance(raw_tm, float) and raw_tm.is_integer():
            tm = str(int(raw_tm))
        else:
            tm = str(raw_tm or "").strip()
        raw_ky = ws_tq.cell(r, col_map["ky"]).value
        if isinstance(raw_ky, float) and raw_ky.is_integer():
            ky = str(int(raw_ky))
        else:
            ky = str(raw_ky or "").strip().upper()
            if ky.endswith(".0") and ky[:-2].isdigit():
                ky = ky[:-2]
        no_co = ""
        if "no_co" in col_map:
            no_co = str(ws_tq.cell(r, col_map["no_co"]).value or "").strip().upper()
        cqt_raw = ""
        cqt_code = ""
        if "cqt" in col_map:
            raw_v = ws_tq.cell(r, col_map["cqt"]).value
            cqt_raw = str(raw_v or "").strip()
            cqt_code = _normalize_cqt(raw_v)

        if so_tien is None or not ngay_ht or not gd or not tm:
            continue

        rows.append({
            "row": r,
            "cqt": cqt_code,
            "cqt_raw": cqt_raw,
            "ky": ky,
            "ky_year": extract_ky_year(ky, ngay_ht),
            "gd": gd,
            "tm": tm,
            "so_tien": float(so_tien),
            "ngay_ht": ngay_ht,
            "ht_year": ngay_ht.year,
            "no_co": no_co,
        })
    return rows


TIEU_MUC_NAMES: Dict[str, str] = {
    "1001": "Thuế TNCN từ tiền lương, tiền công",
    "1003": "Thuế TNCN từ hoạt động SXKD cá nhân",
    "1004": "Thuế TNCN từ đầu tư vốn",
    "1005": "Thuế TNCN từ chuyển nhượng vốn",
    "1006": "Thuế TNCN từ chuyển nhượng BĐS",
    "1007": "Thuế TNCN từ trúng thưởng",
    "1008": "Thuế TNCN từ bản quyền, nhượng quyền TM",
    "1015": "Thuế TNCN từ chuyển nhượng chứng khoán",
    "1052": "Thuế TNDN (SXKD trong nước)",
    "1053": "Thuế TNDN từ hoạt động chuyển nhượng BĐS",
    "1056": "Thuế TNDN của nhà thầu nước ngoài",
    "1701": "Thuế GTGT hàng SXKD trong nước",
    "1702": "Thuế GTGT hàng nhập khẩu",
    "1704": "Thuế GTGT của nhà thầu nước ngoài",
    "1705": "Thuế GTGT vãng lai ngoại tỉnh",
    "2862": "Lệ phí môn bài bậc 1",
    "2863": "Lệ phí môn bài bậc 2",
    "2864": "Lệ phí môn bài bậc 3",
    "4254": "Phạt vi phạm hành chính về thuế",
    "4268": "Phạt khai sai / trốn thuế",
    "4272": "Tiền chậm nộp phạt vi phạm hành chính",
    "4917": "Tiền chậm nộp thuế TNCN",
    "4918": "Tiền chậm nộp thuế TNDN",
    "4931": "Tiền chậm nộp thuế GTGT",
    "4944": "Tiền chậm nộp khoản thu khác (LPMB)",
}


def compute_tq_metrics(
    tq_rows: List[Dict[str, Any]],
    start_year: int,
    end_year: int,
    fiscal_start_month: int = 1,
    is_ntnn: bool = False,
    fiscal_ranges: Optional[Dict[int, Tuple[datetime.date, datetime.date]]] = None,
) -> Dict[str, Any]:
    """
    Tính toán toàn bộ chỉ số từ dữ liệu TQ theo đúng chuẩn nghiệp vụ thanh tra - kiểm tra thuế:
    - Hỗ trợ đầy đủ năm tài chính lệch kỳ cho GTGT (tm='1701') VÀ Thuế Nhà Thầu (is_ntnn=True: cả 1701 & 1052)
      khi fiscal_start_month > 1 (VD tháng 4: 04/YYYY..03/(YYYY+1)):
      các kỳ trước tháng bắt đầu năm tài chính (VD T1..T3 của start_year) được tính là KỲ TRƯỚC để lấy đúng số đầu kỳ!
    - Số đầu kỳ (Cột C Biên bản QHNS): Tính ròng theo từng kỳ trước (< start_year) để C2 (âm) và TK/QT/TQ/TH (dương)
      tự bù trừ đúng dấu, TUYỆT ĐỐI KHÔNG cộng trùng (+C2) + (+TK/QT).
    """
    years = list(range(start_year, end_year + 1))
    prev_year = start_year - 1
    prev_yy = str(prev_year)[-2:]

    LPMB_TMS = ("2862", "2863", "2864", "1801", "1802", "1803", "1804")
    CORE_TMS = ("1701", "1052", "1001") + LPMB_TMS

    other_tms_set = sorted({
        x["tm"] for x in tq_rows
        if x["tm"] and x["tm"] not in CORE_TMS and x["gd"] != "BT"
    })

    metrics: Dict[str, Any] = {
        "gtgt_tk_by_ky": {y: 0.0 for y in years},
        "gtgt_c2_by_ht": {y: 0.0 for y in years},
        "tndn_tk_qt_by_ky": {y: 0.0 for y in years},
        "tndn_c2_by_ht": {y: 0.0 for y in years},
        "tncn_tk_th_tq_by_ky": {y: 0.0 for y in years},
        "tncn_c2_by_ht": {y: 0.0 for y in years},
        "lpmb_ps_by_ky": {y: 0.0 for y in years},
        "lpmb_c2_by_ht": {y: 0.0 for y in years},
        "lpmb_ps_detail": defaultdict(lambda: defaultdict(float)),
        "lpmb_nop_detail": defaultdict(lambda: defaultdict(float)),
        # Tất cả các tiểu mục khác có mặt trên Sổ Tổng Quan (TQ)
        "other_tms": other_tms_set,
        "other_ps_by_ky": {tm: {y: 0.0 for y in years} for tm in other_tms_set},
        "other_nop_by_ht": {tm: {y: 0.0 for y in years} for tm in other_tms_set},
        # Đầu kỳ năm start_year
        "dk_gtgt_qt_truoc": 0.0,
        "dk_gtgt_qt_notes": [],
        "dk_gtgt_c2_prev": 0.0,
        "dk_gtgt_c2_notes": [],
        "dk_tndn_h": 0.0,
        "dk_tndn_i": 0.0,
        "dk_tncn_c2_prev": 0.0,
        "dk_tncn_c2_notes": [],
        "dk_tncn_qt_prev": 0.0,
        "dk_tncn_qt_notes": [],
        "dk_lpmb_notes": [],
        # Số dư đầu kỳ chính xác tuyệt đối tại 01/01/start_year theo phương pháp phân hoạch Sổ TQ
        "opening_exact": {
            "1701": 0.0, "1052": 0.0, "1001": 0.0, "LPMB": 0.0,
            **{tm: 0.0 for tm in other_tms_set},
        },
        # Phát hiện kỳ thiếu TK trên TQ (có C2 nhưng không có TK)
        "tncn_tk_periods": defaultdict(set),
        "tncn_c2_periods": defaultdict(set),
        "gtgt_tk_periods": defaultdict(set),
        "gtgt_c2_periods": defaultdict(set),
    }

    # Kiểm tra quy ước dấu của giao dịch C2 trong file TQ (âm như TMS chuẩn hay dương kèm cột Có)
    c2_vals = [x["so_tien"] for x in tq_rows if x["gd"] == "C2" and x["so_tien"] != 0]
    c2_is_negative = bool(c2_vals and sum(1 for v in c2_vals if v < 0) >= len(c2_vals) / 2)

    def payment_amount(item: Dict[str, Any]) -> float:
        """
        Trả về số tiền nộp (dương = nộp vào NSNN, âm = hủy/hoàn trả/điều chỉnh giảm số nộp).
        Bảo toàn tuyệt đối dấu đại số của giao dịch C2, NT, HC, C5 trên Sổ Tổng Quan (TQ)
        (tuyệt đối không dùng abs() làm mất dấu âm của chứng từ điều chỉnh/chứng từ âm BK/lệnh hoàn C5).
        """
        val = item["so_tien"]
        no_co = item.get("no_co", "")
        if no_co in ("C", "N") and not c2_is_negative and val >= 0:
            return val if no_co == "C" else -val
        return -val if c2_is_negative else val

    def ledger_signed_amount(item: Dict[str, Any]) -> float:
        """
        Quy đổi số tiền giao dịch về dấu chuẩn Sổ Tổng Quan TMS (+ tăng nợ phải nộp, - giảm nợ / đã nộp).
        """
        val = item["so_tien"]
        no_co = item.get("no_co", "")
        if no_co in ("C", "N") and not c2_is_negative and val >= 0:
            return val if no_co == "N" else -val
        return val

    payment_gds = ("C2", "HC", "NT", "C5")
    # Lưu chi tiết từng kỳ trước (< start_year) theo (tm, ky) để tính số đầu kỳ ròng chuẩn xác (KHÔNG cộng trùng C2 + TK/QT)
    prior_periods: Dict[Tuple[str, str], Dict[str, Any]] = {}

    for item in tq_rows:
        tm = item["tm"]
        gd = item["gd"]
        ky = item["ky"]
        val = item["so_tien"]
        ngay_ht_dt = item.get("ngay_ht")

        # Với GTGT (1701) hoặc Toàn bộ Thuế Nhà Thầu (is_ntnn=True: cả 1701 & 1052) khi doanh nghiệp lệch kỳ
        # hoặc có năm tài chính gộp giải thể <= 15 tháng (VD 04/2023-06/2024):
        # quy đổi ĐỒNG NHẤT cả ky_y (phát sinh) và ht_y (số nộp C2) theo cùng khoảng fiscal_ranges!
        if (fiscal_start_month > 1 or fiscal_ranges) and (tm == "1701" or is_ntnn):
            ky_y = extract_ky_fiscal_year(ky, ngay_ht_dt, fiscal_start_month, fiscal_ranges=fiscal_ranges)
            if ngay_ht_dt is not None:
                ht_y = _resolve_date_to_fiscal_year(ngay_ht_dt.date(), fiscal_start_month, fiscal_ranges)
            else:
                ht_y = item["ht_year"]
        else:
            ky_y = item["ky_year"]
            ht_y = item["ht_year"]

        # Bỏ qua bút toán bù trừ nội bộ (BT) vì mỗi cặp BT luôn tự triệt tiêu về 0
        if gd == "BT":
            continue

        # 1. Thu thập dữ liệu theo từng kỳ trước (< start_year) để tính đúng số đầu kỳ Biên bản (tránh cộng trùng C2 + TK/QT)
        if ky_y is not None and ky_y < start_year:
            if tm in ("1701", "1052", "1001"):
                p_rec = prior_periods.setdefault(
                    (tm, ky or str(ky_y)),
                    {
                        "ky_y": ky_y,
                        "accrual_before": 0.0,
                        "accrual_during": 0.0,
                        "qt_during": 0.0,
                        "pay_before": 0.0,
                        "pay_during": 0.0,
                        "qt_notes": [],
                    },
                )
                if gd in payment_gds:
                    p_amt_s = payment_amount(item)
                    if ht_y < start_year:
                        p_rec["pay_before"] += p_amt_s
                    elif start_year <= ht_y <= end_year and gd in ("C2", "HC", "NT"):
                        p_rec["pay_during"] += p_amt_s
                else:
                    l_amt = ledger_signed_amount(item)
                    if ht_y < start_year:
                        p_rec["accrual_before"] += l_amt
                    elif start_year <= ht_y <= end_year:
                        p_rec["accrual_during"] += l_amt
                        if (tm in ("1701", "1052") and gd == "QT") or (tm == "1001" and gd in ("TQ", "TH", "PP")):
                            p_rec["qt_during"] += l_amt
                            p_rec["qt_notes"].append(f"{gd} {ky_y} ({ky}) hạch toán năm {ht_y}")
            elif tm in LPMB_TMS and start_year <= ht_y <= end_year and gd in ("TK", "TQ", "QT"):
                metrics["dk_lpmb_notes"].append(f"TM {tm} ({ky}): {val:,.0f}")

        # 2. Phát sinh theo Năm Kỳ báo cáo (Bucket A)
        in_bucket_a = False
        if ky_y in years:
            if tm == "1701" and gd == "TK":
                metrics["gtgt_tk_by_ky"][ky_y] += val
                in_bucket_a = True
                if ky:
                    metrics["gtgt_tk_periods"][ky_y].add(ky)
            elif tm == "1052" and gd in ("TK", "TQ", "QT"):
                metrics["tndn_tk_qt_by_ky"][ky_y] += val
                in_bucket_a = True
            elif tm == "1001" and gd in ("TK", "TH", "TQ", "PP"):
                metrics["tncn_tk_th_tq_by_ky"][ky_y] += ledger_signed_amount(item)
                in_bucket_a = True
                if gd in ("TK", "PP") and ky:
                    metrics["tncn_tk_periods"][ky_y].add(ky)
            elif tm in LPMB_TMS and gd in ("TK", "TQ", "QT"):
                metrics["lpmb_ps_by_ky"][ky_y] += val
                metrics["lpmb_ps_detail"][ky_y][tm] += val
                in_bucket_a = True
            elif tm in metrics["other_ps_by_ky"] and gd not in payment_gds:
                metrics["other_ps_by_ky"][tm][ky_y] += ledger_signed_amount(item)
                in_bucket_a = True
        elif tm in metrics["other_ps_by_ky"] and gd not in payment_gds and ky_y is None and ht_y in years:
            metrics["other_ps_by_ky"][tm][ht_y] += ledger_signed_amount(item)
            in_bucket_a = True

        if ky_y in years:
            if tm == "1701" and gd == "C2" and ky:
                metrics["gtgt_c2_periods"][ky_y].add(ky)
            elif tm == "1001" and gd == "C2" and ky:
                metrics["tncn_c2_periods"][ky_y].add(ky)

        # 3. Số đã nộp C2 (+ HC, NT, C5 hoàn nộp thừa) theo Năm Ngày hạch toán (Bucket B)
        in_bucket_b = False
        if ht_y in years and gd in payment_gds:
            p_amt = payment_amount(item)
            if tm == "1701":
                metrics["gtgt_c2_by_ht"][ht_y] += p_amt
                in_bucket_b = True
                if gd in ("HC", "NT", "C5"):
                    metrics.setdefault("special_nop_notes", defaultdict(list))[("1701", ht_y)].append(f"{gd}: {p_amt:,.0f}")
            elif tm == "1052":
                metrics["tndn_c2_by_ht"][ht_y] += p_amt
                in_bucket_b = True
                if gd in ("HC", "NT", "C5"):
                    metrics.setdefault("special_nop_notes", defaultdict(list))[("1052", ht_y)].append(f"{gd}: {p_amt:,.0f}")
            elif tm == "1001":
                metrics["tncn_c2_by_ht"][ht_y] += p_amt
                in_bucket_b = True
                if gd in ("HC", "NT", "C5"):
                    metrics.setdefault("special_nop_notes", defaultdict(list))[("1001", ht_y)].append(f"{gd}: {p_amt:,.0f}")
            elif tm in LPMB_TMS:
                metrics["lpmb_c2_by_ht"][ht_y] += p_amt
                metrics["lpmb_nop_detail"][ht_y][tm] += p_amt
                in_bucket_b = True
                if gd in ("HC", "NT", "C5"):
                    metrics.setdefault("special_nop_notes", defaultdict(list))[("LPMB", ht_y)].append(f"{gd}({tm}): {p_amt:,.0f}")
            elif tm in metrics["other_nop_by_ht"]:
                metrics["other_nop_by_ht"][tm][ht_y] += p_amt
                in_bucket_b = True

        # 4. Bucket C: Toàn bộ giao dịch kỳ trước chuyển sang (không thuộc Bucket A & B và không vượt quá end_year)
        if tm in ("1701", "1052", "1001"):
            bucket_c_key = tm
        elif tm in LPMB_TMS:
            bucket_c_key = "LPMB"
        elif tm in metrics["opening_exact"]:
            bucket_c_key = tm
        else:
            bucket_c_key = None

        if bucket_c_key is not None and not in_bucket_a and not in_bucket_b:
            if gd in payment_gds:
                if ht_y < start_year:
                    metrics["opening_exact"][bucket_c_key] += ledger_signed_amount(item)
            else:
                if (ky_y is not None and ky_y < start_year) or (ky_y is None and ht_y < start_year) or (ky_y in years and ht_y <= end_year):
                    metrics["opening_exact"][bucket_c_key] += ledger_signed_amount(item)

    # Tổng hợp số đầu kỳ Biên bản ròng cho từng kỳ trước (< start_year):
    # Nguyên tắc: C2 (âm trên TQ) và TK/QT/TQ/TH (dương trên TQ) bù trừ đúng dấu; KHÔNG cộng trùng (+C2) + (+TK/QT)!
    for (tm, ky), rec in prior_periods.items():
        if rec["pay_during"] == 0 and rec["accrual_during"] == 0:
            continue
        tot_accrual = rec["accrual_before"] + rec["accrual_during"]
        if tot_accrual != 0:
            # Số nghĩa vụ kỳ trước chuyển sang tại đầu kỳ kiểm tra = Tổng phát sinh kỳ trước - Số đã nộp trước kỳ kiểm tra
            net_open = tot_accrual - rec["pay_before"]
            if net_open <= 0:
                continue
            paid_part = min(max(0.0, rec["pay_during"]), net_open)
            unpaid_part = max(0.0, net_open - paid_part)
        else:
            # Trường hợp Sổ TQ cũ không còn dòng TK gốc, chỉ còn chứng từ nộp C2 trong kỳ kiểm tra
            if rec["pay_during"] > 0 and rec["pay_before"] == 0:
                paid_part = rec["pay_during"]
                unpaid_part = 0.0
            else:
                continue

        if tm == "1701":
            if paid_part > 0:
                metrics["dk_gtgt_c2_prev"] += paid_part
                if ky and ky not in metrics["dk_gtgt_c2_notes"]:
                    metrics["dk_gtgt_c2_notes"].append(ky)
            if unpaid_part > 0:
                metrics["dk_gtgt_qt_truoc"] += unpaid_part
                metrics["dk_gtgt_qt_notes"].extend(rec["qt_notes"] or [f"Kỳ {ky} còn phải nộp"])
        elif tm == "1052":
            if paid_part > 0:
                metrics["dk_tndn_i"] += paid_part
            if unpaid_part > 0:
                metrics["dk_tndn_h"] += unpaid_part
        elif tm == "1001":
            if paid_part > 0:
                metrics["dk_tncn_c2_prev"] += paid_part
                if ky and ky not in metrics["dk_tncn_c2_notes"]:
                    metrics["dk_tncn_c2_notes"].append(ky)
            if unpaid_part > 0:
                metrics["dk_tncn_qt_prev"] += unpaid_part
                if ky and ky not in metrics["dk_tncn_qt_notes"]:
                    metrics["dk_tncn_qt_notes"].append(ky)

    return metrics


def _copy_tq_sheet(wb_dst: openpyxl.Workbook, ws_src, sheet_name: str = "TQ") -> None:
    """Chép toàn bộ sheet TQ (hoặc TQ_NTNN) từ file nguồn sang workbook đích."""
    if sheet_name in wb_dst.sheetnames:
        del wb_dst[sheet_name]
    ws_dst = wb_dst.create_sheet(sheet_name)

    for r in range(1, ws_src.max_row + 1):
        for c in range(1, ws_src.max_column + 1):
            src_cell = ws_src.cell(r, c)
            dst_cell = ws_dst.cell(r, c, value=src_cell.value)
            if src_cell.number_format:
                dst_cell.number_format = src_cell.number_format
            if r == 1:
                dst_cell.font = Font(name="Times New Roman", size=10, bold=True)
            else:
                dst_cell.font = Font(name="Times New Roman", size=10)

    for col_idx in range(1, ws_src.max_column + 1):
        col_letter = get_column_letter(col_idx)
        src_dim = ws_src.column_dimensions.get(col_letter)
        ws_dst.column_dimensions[col_letter].width = (src_dim.width if (src_dim and src_dim.width) else 14)


def _eval_simple_formula(
    wb_f: openpyxl.Workbook,
    wb_v: openpyxl.Workbook,
    sheet_name: str,
    cell_ref: str,
    depth: int = 0,
) -> float:
    """
    Đánh giá nhanh các công thức Excel cơ bản (=SUM(...), ='Sheet'!A1, cộng/trừ ô)
    để so sánh chính xác ngay cả khi file Excel vừa sinh bằng openpyxl chưa được mở bằng MS Excel.
    """
    if depth > 20 or sheet_name not in wb_f.sheetnames:
        return 0.0
    ws_v = wb_v[sheet_name]
    cached = ws_v[cell_ref].value
    if isinstance(cached, (int, float)):
        return float(cached)

    ws_f = wb_f[sheet_name]
    raw = ws_f[cell_ref].value
    if isinstance(raw, (int, float)):
        return float(raw)
    if not isinstance(raw, str) or not raw.startswith("="):
        return 0.0

    expr = raw[1:].strip()

    # Trường hợp SUMIF($F$571:$F$772,"[]",K571:K772)
    m_sumif = re.fullmatch(
        r'SUMIF\(\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\s*,\s*"([^"]*)"\s*,\s*\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)\)',
        expr,
        re.IGNORECASE,
    )
    if m_sumif:
        crit_col = openpyxl.utils.column_index_from_string(m_sumif.group(1).upper())
        r_start = int(m_sumif.group(2))
        r_end = int(m_sumif.group(4))
        crit_val = m_sumif.group(5)
        sum_col = openpyxl.utils.column_index_from_string(m_sumif.group(6).upper())
        tot = 0.0
        for rr in range(min(r_start, r_end), max(r_start, r_end) + 1):
            cv = ws_f.cell(rr, crit_col).value
            if str(cv if cv is not None else "") == crit_val:
                v = ws_v.cell(rr, sum_col).value
                if not isinstance(v, (int, float)):
                    v = ws_f.cell(rr, sum_col).value
                if isinstance(v, (int, float)):
                    tot += float(v)
        return tot

    # Trường hợp IF(C71>0,ROUND(C71*0.2,0),0) (dòng C72 Thuế TNDN trên sheet năm)
    m_if_round = re.fullmatch(
        r"IF\(\s*([A-Z]+\d+)\s*>\s*0\s*,\s*ROUND\(\s*([A-Z]+\d+)\s*\*\s*([0-9.]+)\s*,\s*0\s*\)\s*,\s*0\s*\)",
        expr,
        re.IGNORECASE,
    )
    if m_if_round:
        cond_val = _eval_simple_formula(wb_f, wb_v, sheet_name, m_if_round.group(1).upper(), depth + 1)
        if cond_val > 0:
            base_val = _eval_simple_formula(wb_f, wb_v, sheet_name, m_if_round.group(2).upper(), depth + 1)
            rate = float(m_if_round.group(3))
            return float(round(base_val * rate))
        return 0.0

    # Trường hợp SUM(A1:A10) hoặc SUM('Sheet'!A1:A10)
    m_sum = re.fullmatch(r"SUM\((?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?([A-Z]+\d+):([A-Z]+\d+)\)", expr, re.IGNORECASE)
    if m_sum:
        target_sheet = m_sum.group(1) or m_sum.group(2) or sheet_name
        if target_sheet not in wb_f.sheetnames:
            return 0.0
        ws_tf = wb_f[target_sheet]
        ws_tv = wb_v[target_sheet]
        c1, r1 = openpyxl.utils.coordinate_to_tuple(m_sum.group(3).upper())
        c2, r2 = openpyxl.utils.coordinate_to_tuple(m_sum.group(4).upper())
        tot = 0.0
        for rr in range(min(c1, c2), max(c1, c2) + 1):
            for cc in range(min(r1, r2), max(r1, r2) + 1):
                v = ws_tv.cell(rr, cc).value
                if isinstance(v, (int, float)):
                    tot += float(v)
                else:
                    vf = ws_tf.cell(rr, cc).value
                    if isinstance(vf, (int, float)):
                        tot += float(vf)
                    elif isinstance(vf, str) and vf.startswith("="):
                        sub_ref = f"{get_column_letter(cc)}{rr}"
                        tot += _eval_simple_formula(wb_f, wb_v, target_sheet, sub_ref, depth + 1)
        return tot

    # Tách biểu thức cộng/trừ các tham chiếu ô: VD ='05-1'!R657-'05-1'!U657+'05-1'!V657 hoặc =C85+C87+C89
    token_re = re.compile(r"([+-]?)\s*(?:(?:'([^']+)'|([A-Za-z0-9_]+))!)?([A-Z]+\d+)")
    matches = list(token_re.finditer(expr))
    if matches:
        # Đảm bảo toàn bộ biểu thức chỉ gồm các token này và khoảng trắng
        cleaned = token_re.sub("", expr).strip()
        if not cleaned:
            tot = 0.0
            for m in matches:
                sign = -1.0 if m.group(1) == "-" else 1.0
                target_sheet = m.group(2) or m.group(3) or sheet_name
                ref = m.group(4).upper()
                tot += sign * _eval_simple_formula(wb_f, wb_v, target_sheet, ref, depth + 1)
            return tot

    return 0.0


def _eval_year_sheet_values(excel_path: str, years: List[int]) -> Dict[int, Dict[str, float]]:
    """
    Đọc giá trị đã tính sẵn (cached values) hoặc tính trực tiếp công thức trên các sheet năm
    để so sánh chính xác C23 vs F23 và C90 vs F90 ngay cả khi file Excel mới xuất chưa mở bằng MS Excel.
    """
    wb_v = openpyxl.load_workbook(excel_path, data_only=True)
    wb_f = openpyxl.load_workbook(excel_path, data_only=False)
    result: Dict[int, Dict[str, float]] = {
        y: {"gtgt_c23": 0.0, "tncn_c90": 0.0, "tndn_c76": 0.0, "ntnn_c97": 0.0, "ntnn_c98": 0.0}
        for y in years
    }

    for y in years:
        sname = str(y)
        if sname not in wb_f.sheetnames:
            continue
        result[y]["gtgt_c23"] = _eval_simple_formula(wb_f, wb_v, sname, "C23")
        result[y]["tncn_c90"] = _eval_simple_formula(wb_f, wb_v, sname, "C90")
        result[y]["tndn_c76"] = _eval_simple_formula(wb_f, wb_v, sname, "C76")
        result[y]["ntnn_c97"] = _eval_simple_formula(wb_f, wb_v, sname, "C97")
        result[y]["ntnn_c98"] = _eval_simple_formula(wb_f, wb_v, sname, "C98")

    return result


def _write_other_cqt_sheet(
    wb: openpyxl.Workbook,
    other_entries: List[Tuple[str, Dict[str, Any]]],
    target_cqt: str,
) -> None:
    """
    Xuất riêng các giao dịch phát sinh / đã nộp tại Cơ quan thuế khác (không thuộc CQT quản lý chính, VD khác 7901)
    ra sheet độc lập 'TQ_CQT_Khac' gồm:
      - Bảng 1: Tổng hợp theo Nguồn (TQ / TQ_NTNN) × Mã CQT × Năm × Tiểu mục (Phát sinh, Đã nộp, Chênh lệch)
      - Bảng 2: Chi tiết toàn bộ các dòng giao dịch tại các CQT khác để đối chiếu.
    """
    sname = "TQ_CQT_Khac"
    if sname in wb.sheetnames:
        del wb[sname]
    if not other_entries:
        return

    ws = wb.create_sheet(sname)
    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    sum_fill = PatternFill("solid", fgColor="F2F2F2")
    title_font = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")
    font_hdr = Font(name="Times New Roman", size=11, bold=True)
    font_reg = Font(name="Times New Roman", size=11)
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    num_fmt = "#,##0"

    ws["A1"] = f"BẢNG 1: TỔNG HỢP SỐ PHÁT SINH & SỐ ĐÃ NỘP TẠI CÁC CƠ QUAN THUẾ KHÁC (NGOÀI CQT {target_cqt})"
    ws["A1"].font = title_font
    ws.merge_cells("A1:J1")
    ws.row_dimensions[1].height = 24

    headers_b1 = [
        "STT",
        "Nguồn",
        "Mã CQT",
        "Cơ quan thuế (gốc)",
        "Năm",
        "Tiểu mục",
        "Tên sắc thuế / Tiểu mục",
        "Phát sinh phải nộp",
        "Số đã nộp (C2/NT/HC/C5)",
        "Chênh lệch\n(Phát sinh - Đã nộp)",
    ]
    for ci, h in enumerate(headers_b1, start=1):
        c = ws.cell(2, ci, h)
        c.font = font_hdr
        c.alignment = al_c
        c.border = border_all
        c.fill = header_fill
    ws.row_dimensions[2].height = 30

    # Phát hiện quy ước dấu C2 của nhóm CQT khác
    c2_vals = [item["so_tien"] for _, item in other_entries if item["gd"] == "C2" and item["so_tien"] != 0]
    c2_is_neg = bool(c2_vals and sum(1 for v in c2_vals if v < 0) >= len(c2_vals) / 2)
    payment_gds = ("C2", "HC", "NT", "C5")

    summary: Dict[Tuple[str, str, str, int, str], Dict[str, float]] = {}
    for src_name, item in other_entries:
        if item["gd"] == "BT":
            continue
        yr = item.get("ht_year") or item.get("ky_year") or 0
        key = (src_name, item.get("cqt") or "", item.get("cqt_raw") or item.get("cqt") or "", int(yr), item.get("tm") or "")
        bucket = summary.setdefault(key, {"ps": 0.0, "nop": 0.0})
        val = item["so_tien"]
        no_co = item.get("no_co", "")
        if item["gd"] in payment_gds:
            if no_co in ("C", "N") and not c2_is_neg and val >= 0:
                pay = val if no_co == "C" else -val
            else:
                pay = -val if c2_is_neg else val
            bucket["nop"] += pay
        else:
            if no_co in ("C", "N") and not c2_is_neg and val >= 0:
                ps = val if no_co == "N" else -val
            else:
                ps = val
            bucket["ps"] += ps

    cur_r = 3
    b1_start = cur_r
    for idx, (key, agg) in enumerate(sorted(summary.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[0][3], kv[0][4])), start=1):
        src_name, cqt_code, cqt_raw, yr, tm = key
        tm_name = TIEU_MUC_NAMES.get(tm, f"Tiểu mục {tm}")
        row_vals = [
            idx,
            src_name,
            cqt_code,
            cqt_raw,
            yr,
            tm,
            tm_name,
            round(agg["ps"]),
            round(agg["nop"]),
            f"=H{cur_r}-I{cur_r}",
        ]
        for ci, v in enumerate(row_vals, start=1):
            c = ws.cell(cur_r, ci, v)
            c.font = font_bold if ci == 10 else font_reg
            c.border = border_all
            if ci in (3, 6):
                c.number_format = "@"
                c.alignment = al_c
            elif ci in (1, 2, 5):
                c.alignment = al_c
            elif ci >= 8:
                c.number_format = num_fmt
                c.alignment = al_r
            else:
                c.alignment = al_l
        cur_r += 1

    b1_end = cur_r - 1
    if b1_end >= b1_start:
        ws.cell(cur_r, 2, "TỔNG CỘNG CQT KHÁC").alignment = al_l
        ws.cell(cur_r, 8, f"=SUM(H{b1_start}:H{b1_end})")
        ws.cell(cur_r, 9, f"=SUM(I{b1_start}:I{b1_end})")
        ws.cell(cur_r, 10, f"=SUM(J{b1_start}:J{b1_end})")
        for ci in range(1, 11):
            c = ws.cell(cur_r, ci)
            c.font = font_hdr
            c.border = border_all
            c.fill = sum_fill
            if ci >= 8:
                c.number_format = num_fmt
                c.alignment = al_r
        cur_r += 2

    # Bảng 2: Chi tiết từng dòng giao dịch
    ws.cell(cur_r, 1, f"BẢNG 2: CHI TIẾT TỪNG DÒNG GIAO DỊCH TẠI CÁC CQT KHÁC {target_cqt}").font = title_font
    ws.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=10)
    ws.row_dimensions[cur_r].height = 24
    cur_r += 1

    headers_b2 = [
        "STT",
        "Nguồn",
        "Dòng trên TQ",
        "Mã CQT",
        "Cơ quan thuế (gốc)",
        "Kỳ báo cáo",
        "Ký hiệu GD",
        "Tiểu mục",
        "Số tiền",
        "Ngày hạch toán",
    ]
    for ci, h in enumerate(headers_b2, start=1):
        c = ws.cell(cur_r, ci, h)
        c.font = font_hdr
        c.alignment = al_c
        c.border = border_all
        c.fill = header_fill
    ws.row_dimensions[cur_r].height = 26
    cur_r += 1

    for idx, (src_name, item) in enumerate(other_entries, start=1):
        vals = [
            idx,
            src_name,
            item.get("row", ""),
            item.get("cqt", ""),
            item.get("cqt_raw", ""),
            item.get("ky", ""),
            item.get("gd", ""),
            item.get("tm", ""),
            round(item.get("so_tien", 0.0)),
            item.get("ngay_ht"),
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(cur_r, ci, v)
            c.font = font_reg
            c.border = border_all
            if ci in (4, 6, 8):
                c.number_format = "@"
                c.alignment = al_c
            elif ci == 9:
                c.number_format = num_fmt
                c.alignment = al_r
            elif ci == 10:
                c.number_format = "dd/mm/yyyy"
                c.alignment = al_c
            elif ci in (1, 2, 3, 7):
                c.alignment = al_c
            else:
                c.alignment = al_l
        cur_r += 1

    widths = {1: 7, 2: 12, 3: 12, 4: 26, 5: 10, 6: 12, 7: 34, 8: 20, 9: 20, 10: 20}
    for ci, w in widths.items():
        ws.column_dimensions[get_column_letter(ci)].width = w


def integrate_tq_and_qhns_into_workbook(
    excel_path: str,
    tq_excel_path: Optional[str] = None,
    so_nop_excel_path: Optional[str] = None,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    output_path: Optional[str] = None,
    cutoff_date: Optional[datetime.date] = None,
    tq_ntnn_excel_path: Optional[str] = None,
    target_cqt: Optional[str] = "7901",
    fiscal_start_month: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Hàm tích hợp chính (hỗ trợ linh hoạt: chỉ có TQ công ty, chỉ có TQ_NTNN nhà thầu, hoặc cả 2):
    1. Đọc Sổ Tổng Quan (TQ và/hoặc TQ_NTNN) và chèn vào workbook excel_path.
    2. Lọc đúng Cơ quan thuế quản lý (mặc định target_cqt='7901') để lên QHNS & đối chiếu F/G;
       tự động tách riêng các dòng giao dịch thuộc CQT khác sang sheet 'TQ_CQT_Khac'.
    3. Cập nhật cột đối chiếu F & G kèm Comment trên các sheet Năm ('2020'..'2025') cho đúng các sắc thuế có TQ.
    4. Sinh sheet QHNS & DoiChieu_QHNS chuẩn nghiệp vụ thanh tra - kiểm tra thuế.
    5. Lưu ra output_path (mặc định ghi đè/cập nhật chính excel_path).
    """
    if output_path is None:
        output_path = excel_path

    # 1. Đọc file TQ công ty (tùy chọn nếu chỉ đối chiếu TQ_NTNN)
    ws_tq_src = None
    all_tq_rows: List[Dict[str, Any]] = []
    if tq_excel_path and os.path.isfile(tq_excel_path):
        wb_tq = openpyxl.load_workbook(tq_excel_path, data_only=True)
        for cand in wb_tq.sheetnames:
            if cand.strip().upper() in ("TQ", "TONG QUAN", "TỔNG QUAN"):
                ws_tq_src = wb_tq[cand]
                break
        if ws_tq_src is None:
            ws_tq_src = wb_tq[wb_tq.sheetnames[0]]
        all_tq_rows = parse_tq_rows(ws_tq_src)

    # 1b. Đọc file TQ Nhà Thầu (tùy chọn nếu đơn vị không có Thuế nhà thầu)
    ws_ntnn_src = None
    all_ntnn_rows: List[Dict[str, Any]] = []
    if tq_ntnn_excel_path and os.path.isfile(tq_ntnn_excel_path):
        wb_ntnn = openpyxl.load_workbook(tq_ntnn_excel_path, data_only=True)
        for cand in wb_ntnn.sheetnames:
            if cand.strip().upper() in ("TQ", "TQ_NTNN", "TONG QUAN", "TỔNG QUAN", "NTNN"):
                ws_ntnn_src = wb_ntnn[cand]
                break
        if ws_ntnn_src is None:
            ws_ntnn_src = wb_ntnn[wb_ntnn.sheetnames[0]]
        all_ntnn_rows = parse_tq_rows(ws_ntnn_src)

    if not all_tq_rows and not all_ntnn_rows:
        raise ValueError("Không tìm thấy dữ liệu giao dịch hợp lệ trong Sổ Tổng Quan (TQ hoặc TQ_NTNN).")

    has_main_tq = bool(all_tq_rows)
    norm_cqt = _normalize_cqt(target_cqt)
    tq_rows, other_tq_rows = split_tq_rows_by_cqt(all_tq_rows, norm_cqt)
    other_cqt_entries: List[Tuple[str, Dict[str, Any]]] = [("TQ", r) for r in other_tq_rows]

    # 2. Mở workbook tổng hợp tờ khai
    wb = openpyxl.load_workbook(excel_path, data_only=False)

    # Xác định danh sách năm cần lên QHNS
    detected_years = sorted([int(s) for s in wb.sheetnames if re.fullmatch(r"20\d{2}", s)])
    if year_from is not None and year_to is not None and year_from <= year_to:
        years = list(range(year_from, year_to + 1))
    elif detected_years:
        years = detected_years
    else:
        src_for_yr = tq_rows or all_tq_rows or all_ntnn_rows
        ht_years = sorted({x["ht_year"] for x in src_for_yr if 2015 <= x["ht_year"] <= 2035})
        years = ht_years[-6:] if len(ht_years) > 6 else ht_years

    start_year, end_year = years[0], years[-1]
    eff_fiscal_month = fiscal_start_month if (fiscal_start_month is not None and 1 <= fiscal_start_month <= 12) else detect_fiscal_start_month(wb)
    eff_fiscal_ranges = detect_fiscal_ranges_from_workbook(wb)
    metrics = compute_tq_metrics(
        tq_rows,
        start_year,
        end_year,
        fiscal_start_month=eff_fiscal_month,
        fiscal_ranges=eff_fiscal_ranges,
    )
    decl_vals = _eval_year_sheet_values(excel_path, years)

    # Tính riêng phần TNCN phân bổ (PP, tiểu mục 1001) hạch toán tại các CQT khác để CỘNG vào số phát sinh TNCN (F90 & QHNS Cột 4)
    # vì tờ khai 05/QTT-TNCN (C90) kê khai tổng toàn công ty bao gồm cả phần phân bổ sang CQT khác,
    # trong khi số đã nộp (C2) tại CQT khác vẫn tách riêng sang sheet TQ_CQT_Khac.
    all_c2_vals = [x["so_tien"] for x in all_tq_rows if x["gd"] == "C2" and x["so_tien"] != 0]
    all_c2_neg = bool(all_c2_vals and sum(1 for v in all_c2_vals if v < 0) >= len(all_c2_vals) / 2)
    tncn_pp_other_by_ky: Dict[int, float] = {y: 0.0 for y in years}
    for item in other_tq_rows:
        if item["tm"] == "1001" and item["gd"] == "PP" and item["ky_year"] in years:
            val = item["so_tien"]
            no_co = item.get("no_co", "")
            if no_co in ("C", "N") and not all_c2_neg and val >= 0:
                signed_v = val if no_co == "N" else -val
            else:
                signed_v = val
            tncn_pp_other_by_ky[item["ky_year"]] += signed_v

    # 2b. Tính chỉ số Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) nếu có — áp dụng cùng năm tài chính lệch kỳ như GTGT
    ntnn_metrics: Optional[Dict[str, Any]] = None
    if all_ntnn_rows:
        ntnn_rows, other_ntnn_rows = split_tq_rows_by_cqt(all_ntnn_rows, norm_cqt)
        other_cqt_entries.extend([("TQ_NTNN", r) for r in other_ntnn_rows])
        ntnn_metrics = compute_tq_metrics(
            ntnn_rows,
            start_year,
            end_year,
            fiscal_start_month=eff_fiscal_month,
            is_ntnn=True,
            fiscal_ranges=eff_fiscal_ranges,
        )

    # 3. Chép sheet TQ (và TQ_NTNN nếu có) vào workbook đích
    if ws_tq_src is not None and has_main_tq:
        _copy_tq_sheet(wb, ws_tq_src, sheet_name="TQ")
    if ws_ntnn_src is not None and ntnn_metrics is not None:
        _copy_tq_sheet(wb, ws_ntnn_src, sheet_name="TQ_NTNN")

    # 4. Cập nhật cột đối chiếu F, G và Comment trên từng sheet năm
    num_fmt = "#,##0"
    font_regular = Font(name="Times New Roman", size=11)
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    alerts: List[str] = []

    # Lưu trạng thái thiếu tờ khai trên TQ theo từng năm để quyết định công thức Cột (4) & Cột (5) trên QHNS
    gtgt_missing_info: Dict[int, Optional[str]] = {}
    tndn_missing_info: Dict[int, Optional[str]] = {}
    tncn_missing_info: Dict[int, Optional[str]] = {}
    ntg_missing_info: Dict[int, Optional[str]] = {}
    ntt_missing_info: Dict[int, Optional[str]] = {}

    for y in years:
        sname = str(y)
        if sname not in wb.sheetnames:
            gtgt_missing_info[y] = None
            tndn_missing_info[y] = None
            tncn_missing_info[y] = None
            ntg_missing_info[y] = None
            ntt_missing_info[y] = None
            continue
        ws_y = wb[sname]

        # --- Dòng 23, 76, 90: Chỉ đối chiếu khi có Sổ Tổng Quan công ty (TQ) ---
        if has_main_tq:
            f23_val = round(metrics["gtgt_tk_by_ky"][y])
            c23_val = round(decl_vals[y]["gtgt_c23"])
            diff_gtgt = c23_val - f23_val

            ws_y["F23"].value = f23_val
            ws_y["F23"].number_format = num_fmt
            ws_y["F23"].font = font_bold
            ws_y["G23"].value = "=C23-F23"
            ws_y["G23"].number_format = num_fmt
            ws_y["G23"].font = font_bold

            if diff_gtgt > 1000:
                missing_p = sorted(metrics["gtgt_c2_periods"][y] - metrics["gtgt_tk_periods"][y])
                miss_str = ", ".join(missing_p) if missing_p else f"năm {y}"
                ws_y["F23"].comment = Comment(f"Phát sinh TK trên tổng quan rà soát thiếu TK {miss_str}", "QHNS")
                gtgt_missing_info[y] = miss_str
                alerts.append(f"Năm {y} - Thuế GTGT: Tờ khai ({c23_val:,.0f}) > Tổng quan ({f23_val:,.0f}), chênh lệch {diff_gtgt:,.0f} (thiếu TK {miss_str}).")
            else:
                ws_y["F23"].comment = Comment("Phát sinh TK trên tổng quan", "QHNS")
                gtgt_missing_info[y] = None

            # --- Dòng 76: Thuế TNDN phải nộp trong kỳ ---
            f76_val = round(metrics["tndn_tk_qt_by_ky"][y])
            c76_val = round(decl_vals[y]["tndn_c76"])
            diff_tndn = c76_val - f76_val

            ws_y["F76"].value = f76_val
            ws_y["F76"].number_format = num_fmt
            ws_y["F76"].font = font_bold
            ws_y["G76"].value = "=C76-F76"
            ws_y["G76"].number_format = num_fmt
            ws_y["G76"].font = font_bold

            if diff_tndn > 1000:
                ws_y["F76"].comment = Comment(f"Phát sinh TK; TQ; QT trên tổng quan (chênh lệch {diff_tndn:,.0f} so với tờ khai TNDN)", "QHNS")
                tndn_missing_info[y] = f"TNDN {y}"
                alerts.append(f"Năm {y} - Thuế TNDN: Tờ khai ({c76_val:,.0f}) > Tổng quan ({f76_val:,.0f}), chênh lệch {diff_tndn:,.0f} -> Đưa vào Cột (5) Số xử lý qua thanh tra.")
            else:
                ws_y["F76"].comment = Comment("Phát sinh TK; TQ; QT trên tổng quan", "QHNS")
                tndn_missing_info[y] = None

            # --- Dòng 90: Tổng hợp tổng số thuế TNCN phải khấu trừ ---
            f90_main_val = round(metrics["tncn_tk_th_tq_by_ky"][y])
            pp_other_val = round(tncn_pp_other_by_ky[y])
            f90_val = f90_main_val + pp_other_val
            c90_val = round(decl_vals[y]["tncn_c90"])

            # Nếu năm cuối (VD 2025) chưa có tờ khai Quyết toán 05/QTT-TNCN (C90 = 0) mà TQ có phát sinh -> lấy F90 vào C90
            if c90_val == 0 and f90_val > 0:
                ws_y["C90"].value = f90_val
                ws_y["C90"].number_format = num_fmt
                c90_val = f90_val

            diff_tncn = c90_val - f90_val
            ws_y["F90"].value = f90_val
            ws_y["F90"].number_format = num_fmt
            ws_y["F90"].font = font_bold
            ws_y["G90"].value = "=C90-F90"
            ws_y["G90"].number_format = num_fmt
            ws_y["G90"].font = font_bold

            pp_note = (
                f" (CQT {norm_cqt or 'chính'}: {f90_main_val:,.0f} + PP phân bổ CQT khác: {pp_other_val:,.0f})"
                if pp_other_val != 0
                else ""
            )
            if diff_tncn > 1000:
                missing_p = sorted(metrics["tncn_c2_periods"][y] - metrics["tncn_tk_periods"][y])
                miss_str = ", ".join(missing_p) if missing_p else f"năm {y}"
                ws_y["F90"].comment = Comment(f"Phát sinh TK; TH; TQ; PP trên tổng quan{pp_note} rà soát thiếu TK {miss_str}", "QHNS")
                tncn_missing_info[y] = miss_str
                alerts.append(f"Năm {y} - Thuế TNCN: Tờ khai QTT ({c90_val:,.0f}) > Tổng quan ({f90_val:,.0f}), chênh lệch {diff_tncn:,.0f} -> Đưa vào Cột (5) Số xử lý qua thanh tra (thiếu TK {miss_str}).")
            else:
                ws_y["F90"].comment = Comment(f"Phát sinh TK; TH; TQ; PP trên tổng quan{pp_note}", "QHNS")
                tncn_missing_info[y] = None
        else:
            gtgt_missing_info[y] = None
            tndn_missing_info[y] = None
            tncn_missing_info[y] = None

        ws_y.column_dimensions["F"].width = 18
        ws_y.column_dimensions["G"].width = 18

        # --- Dòng 97 & 98: Đối chiếu Thuế Nhà thầu (GTGT 1701 & TNDN 1052) trên TQ_NTNN nếu có ---
        if ntnn_metrics is not None:
            f97_val = round(ntnn_metrics["gtgt_tk_by_ky"][y])
            c97_val = round(decl_vals[y]["ntnn_c97"])
            diff_ntg = c97_val - f97_val
            ws_y["F97"].value = f97_val
            ws_y["F97"].number_format = num_fmt
            ws_y["F97"].font = font_bold
            ws_y["G97"].value = "=C97-F97"
            ws_y["G97"].number_format = num_fmt
            ws_y["G97"].font = font_bold
            if diff_ntg > 1000:
                ws_y["F97"].comment = Comment(
                    f"Phát sinh Thuế GTGT Nhà thầu (TM 1701) trên TQ_NTNN rà soát chênh lệch thiếu {diff_ntg:,.0f} so với tờ khai",
                    "QHNS",
                )
                ntg_missing_info[y] = f"NTNN GTGT {y}"
                alerts.append(
                    f"Năm {y} - Thuế GTGT Nhà thầu: Tờ khai ({c97_val:,.0f}) > Tổng quan NTNN ({f97_val:,.0f}), "
                    f"chênh lệch {diff_ntg:,.0f} -> Đưa vào Cột (5) Số xử lý qua thanh tra."
                )
            else:
                ws_y["F97"].comment = Comment("Phát sinh Thuế GTGT Nhà thầu (TM 1701) trên Sổ Tổng Quan Nhà thầu (TQ_NTNN)", "QHNS")
                ntg_missing_info[y] = None

            f98_val = round(ntnn_metrics["tndn_tk_qt_by_ky"][y])
            c98_val = round(decl_vals[y]["ntnn_c98"])
            diff_ntt = c98_val - f98_val
            ws_y["F98"].value = f98_val
            ws_y["F98"].number_format = num_fmt
            ws_y["F98"].font = font_bold
            ws_y["G98"].value = "=C98-F98"
            ws_y["G98"].number_format = num_fmt
            ws_y["G98"].font = font_bold
            if diff_ntt > 1000:
                ws_y["F98"].comment = Comment(
                    f"Phát sinh Thuế TNDN Nhà thầu (TM 1052) trên TQ_NTNN rà soát chênh lệch thiếu {diff_ntt:,.0f} so với tờ khai",
                    "QHNS",
                )
                ntt_missing_info[y] = f"NTNN TNDN {y}"
                alerts.append(
                    f"Năm {y} - Thuế TNDN Nhà thầu: Tờ khai ({c98_val:,.0f}) > Tổng quan NTNN ({f98_val:,.0f}), "
                    f"chênh lệch {diff_ntt:,.0f} -> Đưa vào Cột (5) Số xử lý qua thanh tra."
                )
            else:
                ws_y["F98"].comment = Comment("Phát sinh Thuế TNDN Nhà thầu (TM 1052) trên Sổ Tổng Quan Nhà thầu (TQ_NTNN)", "QHNS")
                ntt_missing_info[y] = None
        else:
            ntg_missing_info[y] = None
            ntt_missing_info[y] = None

    # 5. Xây dựng sheet QHNS chuẩn 01 biểu duy nhất
    if "QHNS" in wb.sheetnames:
        del wb["QHNS"]
    ws_q = wb.create_sheet("QHNS")

    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    year_fill = PatternFill("solid", fgColor="F2F2F2")
    font_hdr = Font(name="Times New Roman", size=11, bold=True)
    font_year = Font(name="Times New Roman", size=11, bold=True, italic=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_right = Alignment(horizontal="right", vertical="center")

    # Header 3 dòng chuẩn mẫu QHNS
    ws_q["A1"] = "STT"
    ws_q["B1"] = "CHỈ TIÊU"
    ws_q["C1"] = "Số kỳ trước chuyển sang"
    ws_q["D1"] = "Số phát sinh trong kỳ"
    ws_q["F1"] = "Số đã nộp"
    ws_q["G1"] = "Số còn phải nộp"
    ws_q["D2"] = "Số báo cáo phát sinh phải nộp trong kỳ"
    ws_q["E2"] = "Số xử lý qua thanh tra"

    ws_q.merge_cells("A1:A2")
    ws_q.merge_cells("B1:B2")
    ws_q.merge_cells("C1:C2")
    ws_q.merge_cells("D1:E1")
    ws_q.merge_cells("F1:F2")
    ws_q.merge_cells("G1:G2")

    col_nums = ["(1)", "(2)", "(3)", "(4)", "(5)", "(6)", "(7)=(3)+(4)+(5)-(6)"]
    for idx, text in enumerate(col_nums, start=1):
        ws_q.cell(3, idx, text)

    for r in (1, 2, 3):
        for c in range(1, 8):
            cell = ws_q.cell(r, c)
            cell.font = font_hdr
            cell.alignment = align_center
            cell.border = border_all
            cell.fill = header_fill
    ws_q.row_dimensions[1].height = 24
    ws_q.row_dimensions[2].height = 32
    ws_q.row_dimensions[3].height = 20

    prev_rows: Dict[str, int] = {}
    tncn_e_cells: List[str] = []
    gtgt_e_cells: List[str] = []
    tndn_e_cells: List[str] = []
    ntg_e_cells: List[str] = []
    ntt_e_cells: List[str] = []
    special_nop_notes = metrics.get("special_nop_notes", {})
    opening_bb: Dict[str, int] = {"1701": 0, "1052": 0, "1001": 0, "LPMB": 0, "NT_1701": 0, "NT_1052": 0}
    rounding_diff: Dict[str, int] = {"1701": 0, "1052": 0, "1001": 0, "LPMB": 0, "NT_1701": 0, "NT_1052": 0}
    rounding_diff_refs: Dict[str, List[str]] = {"1701": [], "1052": [], "1001": [], "LPMB": [], "NT_1701": [], "NT_1052": []}
    last_rows: Dict[str, int] = {}
    cur_r = 4

    for y_idx, y in enumerate(years):
        sname = str(y)
        has_sheet = sname in wb.sheetnames

        # Dòng tiêu đề Năm
        ws_q.cell(cur_r, 2, f"Năm {y}").font = font_year
        for c in range(1, 8):
            cell = ws_q.cell(cur_r, c)
            cell.border = border_all
            cell.fill = year_fill
        r_year = cur_r
        cur_r += 1

        # 8 dòng sắc thuế của năm
        rows_spec = [
            ("gtgt", "1", "Thuế GTGT"),
            ("gtgt_vl", "", "Thuế GTGT vãng lai"),
            ("tndn", "2", "Thuế TNDN"),
            ("tncn", "3", "Thuế TNCN"),
            ("nt_tong", "4", "Thuế nhà thầu"),
            ("nt_gtgt", "", "Trong đó:\n+ Thuế GTGT"),
            ("nt_tndn", "", "+ Thuế TNDN"),
            ("lpmb", "5", "Lệ phí môn bài"),
        ]
        this_rows: Dict[str, int] = {}
        for key, stt, label in rows_spec:
            this_rows[key] = cur_r
            ws_q.cell(cur_r, 1, stt).alignment = align_center
            ws_q.cell(cur_r, 2, label).alignment = align_left
            for c in range(1, 8):
                cell = ws_q.cell(cur_r, c)
                cell.font = font_bold if key in ("gtgt", "tndn", "tncn", "nt_tong", "lpmb") else font_regular
                cell.border = border_all
                if c >= 3:
                    cell.alignment = align_right
                    cell.number_format = num_fmt
            if key == "nt_gtgt":
                ws_q.row_dimensions[cur_r].height = 30
            cur_r += 1

        # Điền công thức & số liệu cho từng dòng (Sheet QHNS chuẩn Biên bản: KHÔNG ghi nhận số nộp thừa lũy kế kỳ trước chuyển sang)
        # 1) Thuế GTGT
        r_g = this_rows["gtgt"]
        if y_idx == 0:
            h_val = max(0, round(metrics["dk_gtgt_qt_truoc"]))
            i_val = max(0, round(metrics["dk_gtgt_c2_prev"]))
            opening_bb["1701"] = h_val + i_val
            ws_q.cell(r_g, 8, h_val).number_format = num_fmt
            ws_q.cell(r_g, 9, i_val).number_format = num_fmt
            ws_q.cell(r_g, 8).font = font_regular
            ws_q.cell(r_g, 9).font = font_regular
            if h_val != 0:
                qt_note = ", ".join(metrics["dk_gtgt_qt_notes"]) or f"QT kỳ trước nộp trong kỳ kiểm tra"
                ws_q.cell(r_g, 8).comment = Comment(f"Số nộp của {qt_note}", "QHNS")
            if i_val != 0:
                c2_note = ", ".join(metrics["dk_gtgt_c2_notes"]) or f"các kỳ trước {start_year}"
                ws_q.cell(r_g, 9).comment = Comment(f"Số nộp các kỳ trước ({c2_note}) nộp trong kỳ kiểm tra (không gồm số nộp thừa kỳ trước; xem đối chiếu tại sheet DoiChieu_QHNS)", "QHNS")
            ws_q.cell(r_g, 3, f"=H{r_g}+I{r_g}")
        else:
            ws_q.cell(r_g, 3, f"=G{prev_rows['gtgt']}")

        if has_sheet:
            if gtgt_missing_info.get(y):
                ws_q.cell(r_g, 4, f"='{sname}'!F23")
                ws_q.cell(r_g, 4).comment = Comment(
                    f"Vì bị thiếu tờ khai {gtgt_missing_info[y]} nên số phát sinh trên tổng quan bị thiếu, nên khi thể hiện số phát sinh sẽ lấy trên tổng quan",
                    "QHNS",
                )
                ws_q.cell(r_g, 5, f"='{sname}'!G23")
                ws_q.cell(r_g, 5).comment = Comment(
                    "Số chênh lệch thực tế trên tờ khai và phát sinh trên tổng quan do tờ khai bị thiếu",
                    "QHNS",
                )
                gtgt_e_cells.append(f"E{r_g}")
            else:
                ws_q.cell(r_g, 4, f"='{sname}'!C23")
        else:
            ws_q.cell(r_g, 4, round(metrics["gtgt_tk_by_ky"][y]))

        ws_q.cell(r_g, 6, round(metrics["gtgt_c2_by_ht"][y]))
        gtgt_sp_notes = special_nop_notes.get(("1701", y), [])
        if gtgt_sp_notes:
            ws_q.cell(r_g, 6).comment = Comment(
                "GD C2 + NT/HC theo ngày hạch toán trên tổng quan (bao gồm: " + "; ".join(gtgt_sp_notes) + ")",
                "QHNS",
            )
        elif y_idx == 0:
            ws_q.cell(r_g, 6).comment = Comment("GD C2 (và NT/HC nếu có) theo ngày hạch toán trên tổng quan", "QHNS")
        ws_q.cell(r_g, 7, f"=C{r_g}+D{r_g}+E{r_g}-F{r_g}")

        # 2) Thuế GTGT vãng lai
        r_vl = this_rows["gtgt_vl"]
        if y_idx > 0:
            ws_q.cell(r_vl, 3, f"=G{prev_rows['gtgt_vl']}")
        ws_q.cell(r_vl, 7, f"=C{r_vl}+D{r_vl}+E{r_vl}-F{r_vl}")

        # 3) Thuế TNDN
        r_td = this_rows["tndn"]
        if y_idx == 0:
            h_td = max(0, round(metrics["dk_tndn_h"]))
            i_td = max(0, round(metrics["dk_tndn_i"]))
            opening_bb["1052"] = h_td + i_td
            if opening_bb["1052"] > 0:
                ws_q.cell(r_td, 8, h_td).number_format = num_fmt
                ws_q.cell(r_td, 9, i_td).number_format = num_fmt
                ws_q.cell(r_td, 8).font = font_regular
                ws_q.cell(r_td, 9).font = font_regular
                if h_td != 0:
                    ws_q.cell(r_td, 8).comment = Comment("QĐ truy thu TNDN kỳ trước nộp trong kỳ kiểm tra", "QHNS")
                if i_td != 0:
                    ws_q.cell(r_td, 9).comment = Comment("Số nộp TNDN kỳ trước chuyển sang (không gồm nộp thừa kỳ trước)", "QHNS")
                ws_q.cell(r_td, 3, f"=H{r_td}+I{r_td}")
        else:
            ws_q.cell(r_td, 3, f"=G{prev_rows['tndn']}")

        tndn_ps = round(metrics["tndn_tk_qt_by_ky"][y])
        tndn_nop = round(metrics["tndn_c2_by_ht"][y])
        if has_sheet:
            if tndn_missing_info.get(y):
                ws_q.cell(r_td, 4, f"='{sname}'!F76")
                ws_q.cell(r_td, 4).comment = Comment(
                    "Số phát sinh TNDN đã ghi nhận trên tổng quan (TK; TQ; QT)",
                    "QHNS",
                )
                ws_q.cell(r_td, 5, f"='{sname}'!G76")
                ws_q.cell(r_td, 5).comment = Comment(
                    "Số chênh lệch thực tế giữa tờ khai TNDN và phát sinh trên tổng quan",
                    "QHNS",
                )
                tndn_e_cells.append(f"E{r_td}")
            else:
                ws_q.cell(r_td, 4, f"='{sname}'!C76")
        elif tndn_ps > 0:
            ws_q.cell(r_td, 4, tndn_ps)

        if tndn_nop != 0:
            ws_q.cell(r_td, 6, tndn_nop)
            tndn_sp_notes = special_nop_notes.get(("1052", y), [])
            if tndn_sp_notes:
                ws_q.cell(r_td, 6).comment = Comment(
                    "GD C2 + NT/HC theo ngày hạch toán trên tổng quan (bao gồm: " + "; ".join(tndn_sp_notes) + ")",
                    "QHNS",
                )
        ws_q.cell(r_td, 7, f"=C{r_td}+D{r_td}+E{r_td}-F{r_td}")

        # 4) Thuế TNCN
        r_tn = this_rows["tncn"]
        if y_idx == 0:
            h_tn = max(0, round(metrics["dk_tncn_c2_prev"]))
            i_tn = round(metrics["dk_tncn_qt_prev"])
            if h_tn + i_tn < 0:
                # Không ghi nhận số nộp thừa thuần kỳ trước trên Biên bản
                h_tn, i_tn = 0, 0
            opening_bb["1001"] = h_tn + i_tn
            ws_q.cell(r_tn, 8, h_tn).number_format = num_fmt
            ws_q.cell(r_tn, 9, i_tn).number_format = num_fmt
            ws_q.cell(r_tn, 8).font = font_regular
            ws_q.cell(r_tn, 9).font = font_regular
            if h_tn != 0:
                c2_tn_note = ", ".join(metrics["dk_tncn_c2_notes"]) or f"các kỳ trước {start_year}"
                ws_q.cell(r_tn, 8).comment = Comment(f"Số nộp các kỳ trước ({c2_tn_note}) nộp trong kỳ kiểm tra", "QHNS")
            if i_tn != 0:
                ws_q.cell(r_tn, 9).comment = Comment(
                    f"Số QT {start_year - 1} (TQ + TH) (không gồm số nộp thừa lũy kế kỳ trước; xem đối chiếu tại sheet DoiChieu_QHNS)",
                    "QHNS",
                )
            ws_q.cell(r_tn, 3, f"=H{r_tn}+I{r_tn}")
        else:
            ws_q.cell(r_tn, 3, f"=G{prev_rows['tncn']}")

        pp_other_y = round(tncn_pp_other_by_ky[y])
        if has_sheet:
            if tncn_missing_info.get(y):
                ws_q.cell(r_tn, 4, f"='{sname}'!F90")
                ws_q.cell(r_tn, 4).comment = Comment(
                    f"Vì bị thiếu tờ khai {tncn_missing_info[y]} nên số phát sinh trên tổng quan bị thiếu, nên khi thể hiện số phát sinh sẽ lấy trên tổng quan"
                    + (f" (đã bao gồm PP phân bổ CQT khác: {pp_other_y:,.0f})" if pp_other_y != 0 else ""),
                    "QHNS",
                )
                ws_q.cell(r_tn, 5, f"='{sname}'!G90")
                ws_q.cell(r_tn, 5).comment = Comment(
                    "Số chênh lệch thực tế trên tờ khai và phát sinh trên tổng quan do tờ khai bị thiếu",
                    "QHNS",
                )
                tncn_e_cells.append(f"E{r_tn}")
            else:
                ws_q.cell(r_tn, 4, f"='{sname}'!C90")
                if pp_other_y != 0:
                    ws_q.cell(r_tn, 4).comment = Comment(
                        f"Đã bao gồm phần TNCN phân bổ (PP) sang CQT khác: {pp_other_y:,.0f} (số đã nộp tại CQT khác tách riêng ở sheet TQ_CQT_Khac)",
                        "QHNS",
                    )
        else:
            ws_q.cell(r_tn, 4, round(metrics["tncn_tk_th_tq_by_ky"][y] + tncn_pp_other_by_ky[y]))

        ws_q.cell(r_tn, 6, round(metrics["tncn_c2_by_ht"][y]))
        tncn_sp_notes = special_nop_notes.get(("1001", y), [])
        if tncn_sp_notes:
            ws_q.cell(r_tn, 6).comment = Comment(
                "GD C2 + NT/HC theo ngày hạch toán trên tổng quan (bao gồm: " + "; ".join(tncn_sp_notes) + ")",
                "QHNS",
            )
        ws_q.cell(r_tn, 7, f"=C{r_tn}+D{r_tn}+E{r_tn}-F{r_tn}")

        # 5) Thuế nhà thầu (Tổng + GTGT + TNDN) — tự động đổ số liệu từ TQ_NTNN nếu có
        r_nt = this_rows["nt_tong"]
        r_nt_g = this_rows["nt_gtgt"]
        r_nt_t = this_rows["nt_tndn"]
        if ntnn_metrics is not None:
            ntnn_sp = ntnn_metrics.get("special_nop_notes", {})
            if y_idx == 0:
                open_ntg = max(0, round(ntnn_metrics["opening_exact"]["1701"]))
                open_ntt = max(0, round(ntnn_metrics["opening_exact"]["1052"]))
                opening_bb["NT_1701"] = open_ntg
                opening_bb["NT_1052"] = open_ntt
                ws_q.cell(r_nt_g, 3, open_ntg)
                ws_q.cell(r_nt_t, 3, open_ntt)
                ws_q.cell(r_nt, 3, f"=C{r_nt_g}+C{r_nt_t}")
            else:
                ws_q.cell(r_nt, 3, f"=G{prev_rows['nt_tong']}")
                ws_q.cell(r_nt_g, 3, f"=G{prev_rows['nt_gtgt']}")
                ws_q.cell(r_nt_t, 3, f"=G{prev_rows['nt_tndn']}")

            ntg_ps = round(ntnn_metrics["gtgt_tk_by_ky"][y])
            ntg_nop = round(ntnn_metrics["gtgt_c2_by_ht"][y])
            ntt_ps = round(ntnn_metrics["tndn_tk_qt_by_ky"][y])
            ntt_nop = round(ntnn_metrics["tndn_c2_by_ht"][y])
            c97_val_y = round(decl_vals[y]["ntnn_c97"]) if has_sheet else 0
            c98_val_y = round(decl_vals[y]["ntnn_c98"]) if has_sheet else 0

            if has_sheet and ntg_missing_info.get(y):
                ws_q.cell(r_nt_g, 4, f"='{sname}'!F97")
                ws_q.cell(r_nt_g, 4).comment = Comment(
                    "Số phát sinh Thuế GTGT Nhà thầu đã ghi nhận trên Sổ Tổng Quan Nhà thầu (TQ_NTNN)",
                    "QHNS",
                )
                ws_q.cell(r_nt_g, 5, f"='{sname}'!G97")
                ws_q.cell(r_nt_g, 5).comment = Comment(
                    "Số chênh lệch thực tế giữa tờ khai 01/NTNN và phát sinh trên TQ_NTNN",
                    "QHNS",
                )
                ntg_e_cells.append(f"E{r_nt_g}")
            elif has_sheet and c97_val_y > 0:
                ws_q.cell(r_nt_g, 4, f"='{sname}'!C97")
            else:
                ws_q.cell(r_nt_g, 4, ntg_ps)

            ws_q.cell(r_nt_g, 6, ntg_nop)
            if ntnn_sp.get(("1701", y)):
                ws_q.cell(r_nt_g, 6).comment = Comment(
                    "GD C2 + NT/HC/C5 trên TQ_NTNN (bao gồm: " + "; ".join(ntnn_sp[("1701", y)]) + ")",
                    "QHNS",
                )

            if has_sheet and ntt_missing_info.get(y):
                ws_q.cell(r_nt_t, 4, f"='{sname}'!F98")
                ws_q.cell(r_nt_t, 4).comment = Comment(
                    "Số phát sinh Thuế TNDN Nhà thầu đã ghi nhận trên Sổ Tổng Quan Nhà thầu (TQ_NTNN)",
                    "QHNS",
                )
                ws_q.cell(r_nt_t, 5, f"='{sname}'!G98")
                ws_q.cell(r_nt_t, 5).comment = Comment(
                    "Số chênh lệch thực tế giữa tờ khai 01/NTNN và phát sinh trên TQ_NTNN",
                    "QHNS",
                )
                ntt_e_cells.append(f"E{r_nt_t}")
            elif has_sheet and c98_val_y > 0:
                ws_q.cell(r_nt_t, 4, f"='{sname}'!C98")
            else:
                ws_q.cell(r_nt_t, 4, ntt_ps)

            ws_q.cell(r_nt_t, 6, ntt_nop)
            if ntnn_sp.get(("1052", y)):
                ws_q.cell(r_nt_t, 6).comment = Comment(
                    "GD C2 + NT/HC trên TQ_NTNN (bao gồm: " + "; ".join(ntnn_sp[("1052", y)]) + ")",
                    "QHNS",
                )
            ws_q.cell(r_nt, 4, f"=D{r_nt_g}+D{r_nt_t}")
            ws_q.cell(r_nt, 5, f"=E{r_nt_g}+E{r_nt_t}")
            ws_q.cell(r_nt, 6, f"=F{r_nt_g}+F{r_nt_t}")
        else:
            if y_idx > 0:
                ws_q.cell(r_nt, 3, f"=G{prev_rows['nt_tong']}")
                ws_q.cell(r_nt_g, 3, f"=G{prev_rows['nt_gtgt']}")
                ws_q.cell(r_nt_t, 3, f"=G{prev_rows['nt_tndn']}")
            if has_sheet and (round(decl_vals[y]["ntnn_c97"]) > 0 or round(decl_vals[y]["ntnn_c98"]) > 0):
                ws_q.cell(r_nt_g, 4, f"='{sname}'!C97")
                ws_q.cell(r_nt_t, 4, f"='{sname}'!C98")
                ws_q.cell(r_nt, 4, f"=D{r_nt_g}+D{r_nt_t}")
        ws_q.cell(r_nt, 7, f"=G{r_nt_g}+G{r_nt_t}")
        ws_q.cell(r_nt_g, 7, f"=C{r_nt_g}+D{r_nt_g}+E{r_nt_g}-F{r_nt_g}")
        ws_q.cell(r_nt_t, 7, f"=C{r_nt_t}+D{r_nt_t}+E{r_nt_t}-F{r_nt_t}")

        # 6) Lệ phí môn bài (TM 2862, 2863, 2864 và 1801..1804)
        r_mb = this_rows["lpmb"]
        if y_idx == 0:
            open_mb = max(0, round(metrics["opening_exact"]["LPMB"]))
            opening_bb["LPMB"] = open_mb
            if open_mb > 0:
                ws_q.cell(r_mb, 8, 0).number_format = num_fmt
                ws_q.cell(r_mb, 9, open_mb).number_format = num_fmt
                ws_q.cell(r_mb, 8).font = font_regular
                ws_q.cell(r_mb, 9).font = font_regular
                dk_mb_notes = metrics.get("dk_lpmb_notes", [])
                note_detail = "; ".join(dk_mb_notes[:6]) + ("..." if len(dk_mb_notes) > 6 else "")
                ws_q.cell(r_mb, 9).comment = Comment(
                    f"Số phải nộp LPMB kỳ trước chuyển sang ({open_mb:,.0f})" + (f": {note_detail}" if note_detail else ""),
                    "QHNS",
                )
                ws_q.cell(r_mb, 3, f"=H{r_mb}+I{r_mb}")
            else:
                ws_q.cell(r_mb, 3, 0)
        else:
            ws_q.cell(r_mb, 3, f"=G{prev_rows['lpmb']}")

        mb_ps = round(metrics["lpmb_ps_by_ky"][y])
        mb_nop = round(metrics["lpmb_c2_by_ht"][y])
        ws_q.cell(r_mb, 4, mb_ps)
        ps_detail = metrics["lpmb_ps_detail"].get(y, {})
        if ps_detail:
            ps_parts = [f"TM {tm}: {round(v):,.0f}" for tm, v in sorted(ps_detail.items()) if round(v) != 0]
            if ps_parts:
                ws_q.cell(r_mb, 4).comment = Comment("Phát sinh LPMB trên TQ: " + "; ".join(ps_parts), "QHNS")

        ws_q.cell(r_mb, 6, mb_nop)
        nop_detail = metrics["lpmb_nop_detail"].get(y, {})
        mb_sp_notes = special_nop_notes.get(("LPMB", y), [])
        if nop_detail or mb_sp_notes:
            nop_parts = [f"TM {tm}: {round(v):,.0f}" for tm, v in sorted(nop_detail.items()) if round(v) != 0]
            cmt = "Đã nộp LPMB theo ngày hạch toán trên TQ: " + "; ".join(nop_parts)
            if mb_sp_notes:
                cmt += " (bao gồm: " + "; ".join(mb_sp_notes) + ")"
            ws_q.cell(r_mb, 6).comment = Comment(cmt, "QHNS")
        ws_q.cell(r_mb, 7, f"=C{r_mb}+D{r_mb}+E{r_mb}-F{r_mb}")

        # Cộng dồn chênh lệch nhỏ (làm tròn tờ khai vs TQ) cho các năm không tách sang Cột E
        # Lưu cả danh sách tham chiếu ô G23/G76/G90 trên từng sheet năm để dùng công thức động trên DoiChieu_QHNS
        if has_sheet:
            if not gtgt_missing_info.get(y):
                rounding_diff["1701"] += round(decl_vals[y]["gtgt_c23"]) - round(metrics["gtgt_tk_by_ky"][y])
                rounding_diff_refs["1701"].append(f"'{sname}'!G23")
            if not tndn_missing_info.get(y):
                rounding_diff["1052"] += round(decl_vals[y]["tndn_c76"]) - round(metrics["tndn_tk_qt_by_ky"][y])
                rounding_diff_refs["1052"].append(f"'{sname}'!G76")
            if not tncn_missing_info.get(y):
                c90_eff = round(decl_vals[y]["tncn_c90"])
                f90_tot_eff = round(metrics["tncn_tk_th_tq_by_ky"][y] + tncn_pp_other_by_ky[y])
                if c90_eff == 0 and f90_tot_eff > 0:
                    c90_eff = f90_tot_eff
                rounding_diff["1001"] += c90_eff - f90_tot_eff
                rounding_diff_refs["1001"].append(f"'{sname}'!G90")
            if pp_other_y != 0:
                rounding_diff["1001"] += pp_other_y
                rounding_diff_refs["1001"].append(f"({pp_other_y})" if pp_other_y < 0 else str(pp_other_y))
            if ntnn_metrics is not None:
                c97_eff = round(decl_vals[y]["ntnn_c97"])
                f97_eff = round(ntnn_metrics["gtgt_tk_by_ky"][y])
                if not ntg_missing_info.get(y) and c97_eff > 0:
                    rounding_diff["NT_1701"] += c97_eff - f97_eff
                    rounding_diff_refs["NT_1701"].append(f"'{sname}'!G97")
                c98_eff = round(decl_vals[y]["ntnn_c98"])
                f98_eff = round(ntnn_metrics["tndn_tk_qt_by_ky"][y])
                if not ntt_missing_info.get(y) and c98_eff > 0:
                    rounding_diff["NT_1052"] += c98_eff - f98_eff
                    rounding_diff_refs["NT_1052"].append(f"'{sname}'!G98")

        # Nếu là năm cuối cùng và có khoản xử lý qua thanh tra (Cột E), đặt công thức đối chiếu ngược tại cột H
        if y_idx == len(years) - 1:
            last_rows = this_rows
            if tncn_e_cells:
                expr = "-".join(tncn_e_cells)
                ws_q.cell(r_tn, 8, f"=G{r_tn}-{expr}").number_format = num_fmt
                ws_q.cell(r_tn, 8).font = font_bold
                ws_q.cell(r_tn, 8).comment = Comment(
                    "Đối chiếu số dư cuối kỳ (chưa cộng số xử lý qua thanh tra Cột 5; xem đối chiếu khớp thực tế 100% tại sheet DoiChieu_QHNS)",
                    "QHNS",
                )
            if gtgt_e_cells:
                expr = "-".join(gtgt_e_cells)
                ws_q.cell(r_g, 8, f"=G{r_g}-{expr}").number_format = num_fmt
                ws_q.cell(r_g, 8).font = font_bold
                ws_q.cell(r_g, 8).comment = Comment(
                    "Đối chiếu số dư cuối kỳ (chưa cộng số xử lý qua thanh tra Cột 5; xem đối chiếu khớp thực tế 100% tại sheet DoiChieu_QHNS)",
                    "QHNS",
                )
            if tndn_e_cells:
                expr = "-".join(tndn_e_cells)
                ws_q.cell(r_td, 8, f"=G{r_td}-{expr}").number_format = num_fmt
                ws_q.cell(r_td, 8).font = font_bold
                ws_q.cell(r_td, 8).comment = Comment(
                    "Đối chiếu số dư cuối kỳ (chưa cộng số xử lý qua thanh tra Cột 5; xem đối chiếu khớp thực tế 100% tại sheet DoiChieu_QHNS)",
                    "QHNS",
                )
            if ntg_e_cells:
                expr = "-".join(ntg_e_cells)
                ws_q.cell(r_nt_g, 8, f"=G{r_nt_g}-{expr}").number_format = num_fmt
                ws_q.cell(r_nt_g, 8).font = font_bold
                ws_q.cell(r_nt_g, 8).comment = Comment(
                    "Đối chiếu số dư cuối kỳ GTGT NTNN (chưa cộng số xử lý qua thanh tra Cột 5; xem đối chiếu khớp thực tế 100% tại sheet DoiChieu_QHNS)",
                    "QHNS",
                )
            if ntt_e_cells:
                expr = "-".join(ntt_e_cells)
                ws_q.cell(r_nt_t, 8, f"=G{r_nt_t}-{expr}").number_format = num_fmt
                ws_q.cell(r_nt_t, 8).font = font_bold
                ws_q.cell(r_nt_t, 8).comment = Comment(
                    "Đối chiếu số dư cuối kỳ TNDN NTNN (chưa cộng số xử lý qua thanh tra Cột 5; xem đối chiếu khớp thực tế 100% tại sheet DoiChieu_QHNS)",
                    "QHNS",
                )

        prev_rows = this_rows

    # Dòng Tổng cộng cuối bảng QHNS
    r_tot = cur_r
    ws_q.cell(r_tot, 2, "TỔNG CỘNG").font = font_hdr
    ws_q.cell(r_tot, 4, f"=SUM(D5:D{r_tot - 1})")
    ws_q.cell(r_tot, 5, f"=SUM(E5:E{r_tot - 1})")
    ws_q.cell(r_tot, 6, f"=SUM(F5:F{r_tot - 1})")
    for c in range(1, 8):
        cell = ws_q.cell(r_tot, c)
        cell.font = font_hdr
        cell.border = border_all
        if c >= 3:
            cell.alignment = align_right
            cell.number_format = num_fmt

    # Độ rộng cột chuẩn đẹp
    col_widths = {1: 6, 2: 28, 3: 18, 4: 18, 5: 18, 6: 18, 7: 20, 8: 16, 9: 16}
    for c_idx, w in col_widths.items():
        ws_q.column_dimensions[get_column_letter(c_idx)].width = w

    # 6. Sinh sheet DoiChieu_QHNS (ĐỐI CHIẾU SỐ KHỚP THỰC TẾ VỚI SỔ TỔNG QUAN TQ — tách riêng với sheet QHNS vào Biên bản)
    if "DoiChieu_QHNS" in wb.sheetnames:
        del wb["DoiChieu_QHNS"]
    ws_dc = wb.create_sheet("DoiChieu_QHNS")

    ok_fill = PatternFill("solid", fgColor="E2EFDA")
    warn_fill = PatternFill("solid", fgColor="FCE4D6")
    title_font = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")

    # --- BẢNG 1: TỔNG HỢP ĐỐI CHIẾU CUỐI KỲ (BIÊN BẢN QHNS <-> THỰC TẾ SỔ TQ) ---
    ws_dc["A1"] = f"BẢNG 1: ĐỐI CHIẾU SỐ DƯ CUỐI KỲ NĂM {end_year} GIỮA BIÊN BẢN (QHNS) VÀ SỔ TỔNG QUAN (TQ)"
    ws_dc["A1"].font = title_font
    ws_dc.merge_cells("A1:I1")
    ws_dc.row_dimensions[1].height = 24

    dc_headers = [
        "Sắc thuế",
        "Tiểu mục",
        "Dư cuối trên QHNS\n(Số vào Biên bản)",
        "Số nộp thừa / ĐC đầu kỳ\n(Không ghi vào Biên bản)",
        "Số xử lý qua thanh tra\n(Cột 5 QHNS)",
        "Chênh lệch làm tròn\nTờ khai vs TQ",
        "Dư cuối quy đổi\nkhớp thực tế",
        "Dư cuối thực tế\ntrên Sổ TQ",
        "Chênh lệch\n(0 = KHỚP 100%)",
    ]
    for c_idx, h_text in enumerate(dc_headers, start=1):
        cell = ws_dc.cell(2, c_idx, h_text)
        cell.font = font_hdr
        cell.alignment = align_center
        cell.border = border_all
        cell.fill = header_fill
    ws_dc.row_dimensions[2].height = 34

    other_tms = metrics.get("other_tms", []) if has_main_tq else []
    ntnn_other_tms = ntnn_metrics.get("other_tms", []) if ntnn_metrics is not None else []
    main_core_count = 4 if has_main_tq else 0
    ntnn_core_count = 2 if ntnn_metrics is not None else 0

    # Xây dựng Bảng 2 trước trong bộ nhớ vị trí để Bảng 1 có thể link trực tiếp sang số dư năm cuối của các tiểu mục khác
    b1_start_row = 3
    b1_core_count = main_core_count + ntnn_core_count
    b1_total_rows = b1_core_count + len(other_tms) + len(ntnn_other_tms)
    b1_sum_row = b1_start_row + b1_total_rows
    b2_title_row = b1_sum_row + 2
    b2_header_row = b2_title_row + 1
    b2_data_start = b2_header_row + 1

    # --- BẢNG 2: CHI TIẾT QHNS KHỚP THỰC TẾ THEO TỪNG NĂM (TẤT CẢ SẮC THUẾ / TIỂU MỤC CÓ TRÊN TQ) ---
    ws_dc.cell(
        b2_title_row,
        1,
        f"BẢNG 2: CHI TIẾT QUAN HỆ NGÂN SÁCH KHỚP THỰC TẾ TẤT CẢ CÁC SẮC THUẾ TRÊN SỔ TỔNG QUAN ({start_year}–{end_year})",
    ).font = title_font
    ws_dc.merge_cells(start_row=b2_title_row, start_column=1, end_row=b2_title_row, end_column=7)
    ws_dc.row_dimensions[b2_title_row].height = 24

    b2_headers = [
        "STT",
        "CHỈ TIÊU / TIỂU MỤC",
        "Số kỳ trước chuyển sang\n(Thực tế TQ)",
        "Phát sinh phải nộp\n(Theo Sổ TQ)",
        "Điều chỉnh / Xử lý",
        "Số đã nộp / Hoàn trả\n(Theo Sổ TQ)",
        "Số còn phải nộp\n(Khớp thực tế TQ)",
    ]
    for c_idx, h_text in enumerate(b2_headers, start=1):
        cell = ws_dc.cell(b2_header_row, c_idx, h_text)
        cell.font = font_hdr
        cell.alignment = align_center
        cell.border = border_all
        cell.fill = header_fill
    ws_dc.row_dimensions[b2_header_row].height = 30

    b2_rows_spec: List[Tuple[str, str, str, Dict[int, float], Dict[int, float], float]] = []
    next_stt = 1
    if has_main_tq:
        b2_rows_spec.extend([
            ("1701", "1", "Thuế GTGT (1701)", metrics["gtgt_tk_by_ky"], metrics["gtgt_c2_by_ht"], metrics["opening_exact"]["1701"]),
            ("1052", "2", "Thuế TNDN (1052)", metrics["tndn_tk_qt_by_ky"], metrics["tndn_c2_by_ht"], metrics["opening_exact"]["1052"]),
            ("1001", "3", "Thuế TNCN (1001)", metrics["tncn_tk_th_tq_by_ky"], metrics["tncn_c2_by_ht"], metrics["opening_exact"]["1001"]),
            ("LPMB", "4", "Lệ phí môn bài (2862/2863/2864)", metrics["lpmb_ps_by_ky"], metrics["lpmb_c2_by_ht"], metrics["opening_exact"]["LPMB"]),
        ])
        next_stt = 5
    if ntnn_metrics is not None:
        b2_rows_spec.append(
            ("NT_1701", str(next_stt), "Thuế GTGT Nhà thầu (1701 - TQ_NTNN)", ntnn_metrics["gtgt_tk_by_ky"], ntnn_metrics["gtgt_c2_by_ht"], ntnn_metrics["opening_exact"]["1701"])
        )
        next_stt += 1
        b2_rows_spec.append(
            ("NT_1052", str(next_stt), "Thuế TNDN Nhà thầu (1052 - TQ_NTNN)", ntnn_metrics["tndn_tk_qt_by_ky"], ntnn_metrics["tndn_c2_by_ht"], ntnn_metrics["opening_exact"]["1052"])
        )
        next_stt += 1
    for o_tm in other_tms:
        o_name = TIEU_MUC_NAMES.get(o_tm, "Khoản thu khác")
        b2_rows_spec.append((o_tm, str(next_stt), f"{o_name} ({o_tm})", metrics["other_ps_by_ky"][o_tm], metrics["other_nop_by_ht"][o_tm], metrics["opening_exact"][o_tm]))
        next_stt += 1
    if ntnn_metrics is not None:
        for o_tm in ntnn_other_tms:
            o_name = TIEU_MUC_NAMES.get(o_tm, "Khoản thu khác NTNN")
            b2_rows_spec.append((f"NT_{o_tm}", str(next_stt), f"{o_name} ({o_tm} - TQ_NTNN)", ntnn_metrics["other_ps_by_ky"][o_tm], ntnn_metrics["other_nop_by_ht"][o_tm], ntnn_metrics["opening_exact"][o_tm]))
            next_stt += 1

    dc_r = b2_data_start
    dc_prev: Dict[str, int] = {}
    dc_last_year_rows: Dict[str, int] = {}
    for y_idx, y in enumerate(years):
        ws_dc.cell(dc_r, 2, f"Năm {y}").font = font_year
        for c in range(1, 8):
            cell = ws_dc.cell(dc_r, c)
            cell.border = border_all
            cell.fill = year_fill
        dc_r += 1

        dc_this: Dict[str, int] = {}
        for tm_key, stt, label, ps_map, nop_map, open_ex in b2_rows_spec:
            dc_this[tm_key] = dc_r
            ws_dc.cell(dc_r, 1, stt).alignment = align_center
            ws_dc.cell(dc_r, 2, label).alignment = align_left
            if y_idx == 0:
                ws_dc.cell(dc_r, 3, round(open_ex))
            else:
                ws_dc.cell(dc_r, 3, f"=G{dc_prev[tm_key]}")
            ws_dc.cell(dc_r, 4, round(ps_map[y]))
            ws_dc.cell(dc_r, 5, 0)
            ws_dc.cell(dc_r, 6, round(nop_map[y]))
            ws_dc.cell(dc_r, 7, f"=C{dc_r}+D{dc_r}+E{dc_r}-F{dc_r}")
            for c in range(1, 8):
                cell = ws_dc.cell(dc_r, c)
                cell.font = font_bold if c in (2, 7) else font_regular
                cell.border = border_all
                if c >= 3:
                    cell.alignment = align_right
                    cell.number_format = num_fmt
            dc_r += 1
        dc_prev = dc_this
        if y_idx == len(years) - 1:
            dc_last_year_rows = dc_this

    # --- ĐIỀN BẢNG 1: TỔNG HỢP ĐỐI CHIẾU CUỐI KỲ TẤT CẢ SẮC THUẾ (BIÊN BẢN QHNS <-> THỰC TẾ SỔ TQ & TQ_NTNN) ---
    dc_specs = []
    if has_main_tq:
        dc_specs.extend([
            ("Thuế GTGT", "1701", "1701", last_rows.get("gtgt"), gtgt_e_cells, metrics["gtgt_tk_by_ky"], metrics["gtgt_c2_by_ht"], metrics["opening_exact"]["1701"], True),
            ("Thuế TNDN", "1052", "1052", last_rows.get("tndn"), tndn_e_cells, metrics["tndn_tk_qt_by_ky"], metrics["tndn_c2_by_ht"], metrics["opening_exact"]["1052"], True),
            ("Thuế TNCN", "1001", "1001", last_rows.get("tncn"), tncn_e_cells, metrics["tncn_tk_th_tq_by_ky"], metrics["tncn_c2_by_ht"], metrics["opening_exact"]["1001"], True),
            ("Lệ phí môn bài", "LPMB", "2862/2863/2864", last_rows.get("lpmb"), [], metrics["lpmb_ps_by_ky"], metrics["lpmb_c2_by_ht"], metrics["opening_exact"]["LPMB"], True),
        ])
    if ntnn_metrics is not None:
        dc_specs.append(
            ("Thuế GTGT Nhà thầu (TQ_NTNN)", "NT_1701", "1701 (NTNN)", last_rows.get("nt_gtgt"), ntg_e_cells, ntnn_metrics["gtgt_tk_by_ky"], ntnn_metrics["gtgt_c2_by_ht"], ntnn_metrics["opening_exact"]["1701"], True)
        )
        dc_specs.append(
            ("Thuế TNDN Nhà thầu (TQ_NTNN)", "NT_1052", "1052 (NTNN)", last_rows.get("nt_tndn"), ntt_e_cells, ntnn_metrics["tndn_tk_qt_by_ky"], ntnn_metrics["tndn_c2_by_ht"], ntnn_metrics["opening_exact"]["1052"], True)
        )
    for o_tm in other_tms:
        o_name = TIEU_MUC_NAMES.get(o_tm, f"Tiểu mục {o_tm}")
        dc_specs.append((o_name, o_tm, o_tm, None, [], metrics["other_ps_by_ky"][o_tm], metrics["other_nop_by_ht"][o_tm], metrics["opening_exact"][o_tm], False))
    if ntnn_metrics is not None:
        for o_tm in ntnn_other_tms:
            o_name = TIEU_MUC_NAMES.get(o_tm, f"Tiểu mục {o_tm}")
            dc_specs.append((f"{o_name} (TQ_NTNN)", f"NT_{o_tm}", f"{o_tm} (NTNN)", None, [], ntnn_metrics["other_ps_by_ky"][o_tm], ntnn_metrics["other_nop_by_ht"][o_tm], ntnn_metrics["opening_exact"][o_tm], False))

    dc_r = b1_start_row
    for label, tm_key, tm_disp, r_last_q, e_cells, ps_map, nop_map, open_ex_raw, is_on_qhns in dc_specs:
        open_exact_val = round(open_ex_raw)
        tq_close_val = round(
            open_ex_raw
            + sum(ps_map.values())
            - sum(nop_map.values())
        )

        ws_dc.cell(dc_r, 1, label).alignment = align_left
        ws_dc.cell(dc_r, 2, tm_disp).alignment = align_center
        if is_on_qhns:
            open_diff_val = open_exact_val - opening_bb[tm_key]
            rdiff_val = rounding_diff[tm_key]
            ws_dc.cell(dc_r, 3, f"='QHNS'!G{r_last_q}" if r_last_q else 0)
            ws_dc.cell(dc_r, 4, open_diff_val)
            if open_diff_val != 0:
                ws_dc.cell(dc_r, 4).comment = Comment(
                    f"Đầu kỳ thực tế trên TQ ({open_exact_val:,.0f}) − Đầu kỳ ghi vào Biên bản QHNS ({opening_bb[tm_key]:,.0f}) = {open_diff_val:+,.0f}",
                    "QHNS",
                )
            if e_cells:
                e_sum_expr = "+".join([f"'QHNS'!{ref}" for ref in e_cells])
                ws_dc.cell(dc_r, 5, f"={e_sum_expr}")
            else:
                ws_dc.cell(dc_r, 5, 0)
            rdiff_refs = rounding_diff_refs.get(tm_key, [])
            if rdiff_refs:
                ws_dc.cell(dc_r, 6, "=" + "+".join(rdiff_refs))
            else:
                ws_dc.cell(dc_r, 6, rounding_diff[tm_key])
        else:
            # Sắc thuế/tiểu mục khác (Phạt, Tiền chậm nộp, TNCN khác...): đối chiếu trực tiếp từ Bảng 2
            ws_dc.cell(dc_r, 3, 0)
            b2_last_r = dc_last_year_rows.get(tm_key)
            ws_dc.cell(dc_r, 4, f"=G{b2_last_r}" if b2_last_r else tq_close_val)
            ws_dc.cell(dc_r, 4).comment = Comment(
                "Tiểu mục ngoài biểu QHNS Biên bản — liên kết trực tiếp số dư cuối kỳ từ Bảng 2 (Sổ TQ)",
                "QHNS",
            )
            ws_dc.cell(dc_r, 5, 0)
            ws_dc.cell(dc_r, 6, 0)

        ws_dc.cell(dc_r, 7, f"=C{dc_r}+D{dc_r}-E{dc_r}-F{dc_r}")
        ws_dc.cell(dc_r, 8, tq_close_val)
        ws_dc.cell(dc_r, 9, f"=G{dc_r}-H{dc_r}")

        for c in range(1, 10):
            cell = ws_dc.cell(dc_r, c)
            cell.font = font_bold if c in (1, 7, 8, 9) else font_regular
            cell.border = border_all
            if c >= 3:
                cell.alignment = align_right
                cell.number_format = num_fmt
        ws_dc.cell(dc_r, 9).fill = ok_fill
        dc_r += 1

    # Dòng TỔNG CỘNG cuối Bảng 1 (khớp tổng dòng cuối cùng toàn Sổ TQ)
    ws_dc.cell(b1_sum_row, 1, "TỔNG CỘNG TẤT CẢ SẮC THUẾ TRÊN TQ").alignment = align_left
    ws_dc.cell(b1_sum_row, 2, "Tất cả TM").alignment = align_center
    for c_col, col_let in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I")):
        ws_dc.cell(b1_sum_row, c_col, f"=SUM({col_let}{b1_start_row}:{col_let}{b1_sum_row - 1})")
    for c in range(1, 10):
        cell = ws_dc.cell(b1_sum_row, c)
        cell.font = font_hdr
        cell.border = border_all
        cell.fill = year_fill
        if c >= 3:
            cell.alignment = align_right
            cell.number_format = num_fmt
    ws_dc.cell(b1_sum_row, 9).fill = ok_fill

    dc_widths = {1: 36, 2: 34, 3: 20, 4: 22, 5: 20, 6: 20, 7: 20, 8: 20, 9: 20}
    for c_idx, w in dc_widths.items():
        ws_dc.column_dimensions[get_column_letter(c_idx)].width = w

    # Xuất riêng sheet TQ_CQT_Khac nếu Sổ Tổng Quan có giao dịch tại cơ quan thuế khác target_cqt (VD khác 7901)
    if other_cqt_entries:
        _write_other_cqt_sheet(wb, other_cqt_entries, norm_cqt or "7901")
        other_cqts_list = sorted({item.get("cqt") or "Khác" for _, item in other_cqt_entries})
        alerts.append(
            f"Đã tách riêng {len(other_cqt_entries)} dòng giao dịch thuộc CQT khác {norm_cqt or '7901'} "
            f"({', '.join(other_cqts_list)}) sang sheet 'TQ_CQT_Khac' (KHÔNG gộp vào biểu QHNS của CQT {norm_cqt or '7901'})."
        )

    # Cập nhật / bổ sung sheet biểu tính tiền chậm nộp & xử phạt VPHC và Báo cáo tổng hợp Biên bản
    from penalty_calculator import add_late_payment_sheet
    from report_builder import add_audit_report_sheet
    add_late_payment_sheet(wb, years=years, cutoff_date=cutoff_date)
    add_audit_report_sheet(wb, years=years)

    wb.save(output_path)

    return {
        "output_path": output_path,
        "years": years,
        "alerts": alerts,
        "metrics": metrics,
        "target_cqt": norm_cqt,
        "other_cqt_count": len(other_cqt_entries),
    }


def normalize_period_key(ky_val: Any, fallback_date: Optional[datetime.datetime] = None) -> Tuple[int, int, int, str]:
    """
    Chuẩn hóa Kỳ kê khai về khóa sắp xếp thống nhất (year, kind_order, num, display_label):
      - kind_order: 1=Tháng ('01/2024'), 2=Quý ('Q1/2024'), 3=Cả năm ('Năm 2024')
    Hỗ trợ các dạng: '2401', '01/2024', 'T01/2024', '24Q1', 'Q1/2024', '24CN', '2024', '15/06/2024'...
    """
    s = str(ky_val or "").strip().upper()
    # Dạng TMS ngắn: YYMM, YYQ#, YYCN
    m_tms_m = re.match(r"^(\d{2})(0[1-9]|1[0-2])$", s)
    if m_tms_m:
        yr = 2000 + int(m_tms_m.group(1))
        mo = int(m_tms_m.group(2))
        return (yr, 1, mo, f"{mo:02d}/{yr}")
    m_tms_q = re.match(r"^(\d{2})Q([1-4])$", s)
    if m_tms_q:
        yr = 2000 + int(m_tms_q.group(1))
        q = int(m_tms_q.group(2))
        return (yr, 2, q, f"Q{q}/{yr}")
    m_tms_y = re.match(r"^(\d{2})CN$", s)
    if m_tms_y:
        yr = 2000 + int(m_tms_y.group(1))
        return (yr, 3, 0, f"Năm {yr}")

    # Dạng XML/VN: Q1/2024, 01/2024, T01/2024, 15/06/2024, 2024
    m_q = re.match(r"^Q([1-4])[/-](20\d{2})$", s)
    if m_q:
        q, yr = int(m_q.group(1)), int(m_q.group(2))
        return (yr, 2, q, f"Q{q}/{yr}")
    m_m = re.match(r"^T?(0?[1-9]|1[0-2])[/-](20\d{2})$", s)
    if m_m:
        mo, yr = int(m_m.group(1)), int(m_m.group(2))
        return (yr, 1, mo, f"{mo:02d}/{yr}")
    m_d = re.match(r"^(\d{2})[/-](0?[1-9]|1[0-2])[/-](20\d{2})$", s)
    if m_d:
        mo, yr = int(m_d.group(2)), int(m_d.group(3))
        return (yr, 1, mo, f"{mo:02d}/{yr}")
    m_y = re.search(r"\b(20\d{2})\b", s)
    if m_y:
        yr = int(m_y.group(1))
        return (yr, 3, 0, f"Năm {yr}")
    if fallback_date:
        return (fallback_date.year, 3, 0, f"Năm {fallback_date.year}")
    return (0, 3, 0, s or "Khác")


def integrate_period_reconciliation_for_tax_type(
    excel_path: str,
    tax_type: str,
    tq_excel_path: Optional[str] = None,
    tq_ntnn_excel_path: Optional[str] = None,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    target_cqt: Optional[str] = "7901",
    output_path: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Đối chiếu THEO TỪNG KỲ KÊ KHAI cho file Excel tách riêng theo từng sắc thuế (GTGT, TNDN, TNCN, NTNN, Khác):
    - Đọc đúng nguồn Sổ Tổng Quan tương ứng (NTNN dùng TQ_NTNN; các sắc thuế khác dùng TQ), lọc đúng target_cqt.
    - Tạo sheet 'DoiChieu_KyKeKhai' so sánh trực tiếp từng kỳ kê khai (Tháng/Quý/Năm):
      Số thuế trên Tờ khai <-> Phát sinh trên Sổ TQ <-> Chênh lệch <-> Số đã nộp theo kỳ trên TQ <-> Còn phải nộp.
    - Chép kèm sheet TQ (hoặc TQ_NTNN) vào file để tự đối chiếu độc lập.
    """
    if output_path is None:
        output_path = excel_path

    tt = (tax_type or "").strip().upper()
    src_tq_path = tq_ntnn_excel_path if tt == "NTNN" else tq_excel_path
    if not src_tq_path or not os.path.isfile(src_tq_path):
        return {"output_path": output_path, "reconciled": False, "rows": 0}

    wb_tq = openpyxl.load_workbook(src_tq_path, data_only=True)
    ws_tq_src = None
    for cand in wb_tq.sheetnames:
        if cand.strip().upper() in ("TQ", "TQ_NTNN", "TONG QUAN", "TỔNG QUAN", "NTNN"):
            ws_tq_src = wb_tq[cand]
            break
    if ws_tq_src is None:
        ws_tq_src = wb_tq[wb_tq.sheetnames[0]]

    all_rows = parse_tq_rows(ws_tq_src)
    if not all_rows:
        return {"output_path": output_path, "reconciled": False, "rows": 0}

    norm_cqt = _normalize_cqt(target_cqt)
    main_rows, other_rows = split_tq_rows_by_cqt(all_rows, norm_cqt)

    c2_vals = [x["so_tien"] for x in all_rows if x["gd"] == "C2" and x["so_tien"] != 0]
    c2_is_neg = bool(c2_vals and sum(1 for v in c2_vals if v < 0) >= len(c2_vals) / 2)
    payment_gds = ("C2", "HC", "NT", "C5")

    def _signed_ps(it: Dict[str, Any]) -> float:
        v = it["so_tien"]
        nc = it.get("no_co", "")
        if nc in ("C", "N") and not c2_is_neg and v >= 0:
            return v if nc == "N" else -v
        return v

    def _signed_pay(it: Dict[str, Any]) -> float:
        v = it["so_tien"]
        nc = it.get("no_co", "")
        if nc in ("C", "N") and not c2_is_neg and v >= 0:
            return v if nc == "C" else -v
        return -v if c2_is_neg else v

    wb = openpyxl.load_workbook(excel_path, data_only=False)
    wb_v = openpyxl.load_workbook(excel_path, data_only=True)
    _copy_tq_sheet(wb, ws_tq_src, sheet_name="TQ_NTNN" if tt == "NTNN" else "TQ")

    # Thu thập dữ liệu Tờ khai & TQ theo (period_key, tieu_muc)
    # decl_map[(pk, tm)] = (formula_or_value, numeric_value)
    decl_map: Dict[Tuple[Tuple[int, int, int, str], str], Tuple[Any, float]] = {}
    tq_ps_map: Dict[Tuple[Tuple[int, int, int, str], str], float] = defaultdict(float)
    tq_nop_map: Dict[Tuple[Tuple[int, int, int, str], str], float] = defaultdict(float)

    def _in_year_bounds(yr: int) -> bool:
        if year_from is not None and yr and yr < year_from:
            return False
        if year_to is not None and yr and yr > year_to:
            return False
        return True

    if tt == "GTGT":
        # Đọc từng kỳ trên 'So Thue GTGT Cuoi' (Cột G(7)=Kỳ tính thuế, tự dò cột [40] Thuế GTGT còn phải nộp - mặc định AJ)
        if "So Thue GTGT Cuoi" in wb.sheetnames:
            ws_g = wb["So Thue GTGT Cuoi"]
            gtgt_tax_col = "AJ"
            for ci in range(1, ws_g.max_column + 1):
                h_txt = str(ws_g.cell(1, ci).value or "")
                if "(40)" in h_txt and "(40a)" not in h_txt and "(40b)" not in h_txt:
                    gtgt_tax_col = get_column_letter(ci)
                    break
            for r in range(3, ws_g.max_row + 1):
                lbl_e = str(ws_g.cell(r, 5).value or "").strip().lower()
                if "tổng cộng" in lbl_e:
                    continue
                ky_raw = str(ws_g.cell(r, 7).value or "").strip()
                if not ky_raw:
                    continue
                pk = normalize_period_key(ky_raw)
                if not _in_year_bounds(pk[0]):
                    continue
                val_num = _eval_simple_formula(wb, wb_v, "So Thue GTGT Cuoi", f"{gtgt_tax_col}{r}")
                decl_map[(pk, "1701")] = (f"='So Thue GTGT Cuoi'!{gtgt_tax_col}{r}", val_num)
        for it in main_rows:
            if it["tm"] != "1701" or it["gd"] == "BT":
                continue
            pk = normalize_period_key(it["ky"], it["ngay_ht"])
            if not _in_year_bounds(pk[0]):
                continue
            if it["gd"] == "TK":
                tq_ps_map[(pk, "1701")] += _signed_ps(it)
            elif it["gd"] in payment_gds:
                tq_nop_map[(pk, "1701")] += _signed_pay(it)

    elif tt == "NTNN":
        # Đọc từng kỳ trên 'So Thue NTNN Cuoi (Lech ky)' (nếu có) hoặc 'So Thue NTNN Cuoi' (Cột F=Kỳ kê khai, O=GTGT NTNN, Q=TNDN NTNN)
        ntnn_src_sheet = (
            "So Thue NTNN Cuoi (Lech ky)"
            if "So Thue NTNN Cuoi (Lech ky)" in wb.sheetnames
            else ("So Thue NTNN Cuoi" if "So Thue NTNN Cuoi" in wb.sheetnames else None)
        )
        if ntnn_src_sheet:
            ws_n = wb[ntnn_src_sheet]
            ntnn_decl_rows_g: Dict[Tuple[int, int, int, str], List[int]] = defaultdict(list)
            ntnn_decl_rows_t: Dict[Tuple[int, int, int, str], List[int]] = defaultdict(list)
            for r in range(2, ws_n.max_row + 1):
                ky_raw = str(ws_n.cell(r, 6).value or "").strip()
                if not ky_raw or "tổng cộng" in ky_raw.lower():
                    continue
                pk = normalize_period_key(ky_raw)
                if not _in_year_bounds(pk[0]):
                    continue
                ntnn_decl_rows_g[pk].append(r)
                ntnn_decl_rows_t[pk].append(r)
            for pk, rlist in ntnn_decl_rows_g.items():
                expr_g = "=" + "+".join(f"'{ntnn_src_sheet}'!O{rr}" for rr in rlist)
                val_g = sum(_eval_simple_formula(wb, wb_v, ntnn_src_sheet, f"O{rr}") for rr in rlist)
                decl_map[(pk, "1701")] = (expr_g, val_g)
                expr_t = "=" + "+".join(f"'{ntnn_src_sheet}'!Q{rr}" for rr in rlist)
                val_t = sum(_eval_simple_formula(wb, wb_v, ntnn_src_sheet, f"Q{rr}") for rr in rlist)
                decl_map[(pk, "1052")] = (expr_t, val_t)
        for it in main_rows:
            if it["tm"] not in ("1701", "1052") or it["gd"] == "BT":
                continue
            pk = normalize_period_key(it["ky"], it["ngay_ht"])
            if not _in_year_bounds(pk[0]):
                continue
            if it["gd"] in ("TK", "TQ", "QT"):
                tq_ps_map[(pk, it["tm"])] += _signed_ps(it)
            elif it["gd"] in payment_gds:
                tq_nop_map[(pk, it["tm"])] += _signed_pay(it)

    elif tt in ("TNDN", "TNCN"):
        target_tm = "1052" if tt == "TNDN" else "1001"
        target_cell = "C76" if tt == "TNDN" else "C90"
        valid_ps_gds = ("TK", "TQ", "QT") if tt == "TNDN" else ("TK", "TH", "TQ", "PP")
        # Đối chiếu theo Năm từ các sheet năm ('2020'..'2025')
        for sname in wb.sheetnames:
            if re.fullmatch(r"20\d{2}", sname):
                yr = int(sname)
                if not _in_year_bounds(yr):
                    continue
                pk = (yr, 3, 0, f"Năm {yr}")
                val_num = _eval_simple_formula(wb, wb_v, sname, target_cell)
                decl_map[(pk, target_tm)] = (f"='{sname}'!{target_cell}", val_num)
        for it in main_rows:
            if it["tm"] != target_tm or it["gd"] == "BT":
                continue
            yr = it["ky_year"] or it["ht_year"]
            if not _in_year_bounds(yr):
                continue
            pk = (yr, 3, 0, f"Năm {yr}")
            if it["gd"] in valid_ps_gds:
                tq_ps_map[(pk, target_tm)] += _signed_ps(it)
            elif it["gd"] in payment_gds:
                tq_nop_map[(pk, target_tm)] += _signed_pay(it)
        # Với TNCN: cộng thêm PP phân bổ CQT khác vào phát sinh theo năm
        if tt == "TNCN":
            for it in other_rows:
                if it["tm"] == "1001" and it["gd"] == "PP":
                    yr = it["ky_year"] or it["ht_year"]
                    if not _in_year_bounds(yr):
                        continue
                    pk = (yr, 3, 0, f"Năm {yr}")
                    tq_ps_map[(pk, "1001")] += _signed_ps(it)

    else:
        # Khác (Lệ phí môn bài: 2862/2863/2864)
        lpmb_tms = ("2862", "2863", "2864", "1801", "1802", "1803", "1804")
        for it in main_rows:
            if it["tm"] not in lpmb_tms or it["gd"] == "BT":
                continue
            pk = normalize_period_key(it["ky"], it["ngay_ht"])
            if not _in_year_bounds(pk[0]):
                continue
            if it["gd"] in ("TK", "TQ", "QT"):
                tq_ps_map[(pk, it["tm"])] += _signed_ps(it)
            elif it["gd"] in payment_gds:
                tq_nop_map[(pk, it["tm"])] += _signed_pay(it)

    all_keys = sorted(set(decl_map.keys()) | set(tq_ps_map.keys()) | set(tq_nop_map.keys()), key=lambda x: (x[0], x[1]))

    sname_dc = "DoiChieu_KyKeKhai"
    if sname_dc in wb.sheetnames:
        del wb[sname_dc]
    ws_dc = wb.create_sheet(sname_dc, index=0)

    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    sum_fill = PatternFill("solid", fgColor="D9E1F2")
    title_font = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")
    font_hdr = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
    font_reg = Font(name="Times New Roman", size=10)
    font_bold = Font(name="Times New Roman", size=10, bold=True)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    num_fmt = "#,##0"

    cqt_label = f" (CQT quản lý: {norm_cqt})" if norm_cqt else " (Tất cả CQT)"
    ws_dc["A1"] = f"ĐỐI CHIẾU TỜ KHAI THUẾ {tt} VỚI SỔ TỔNG QUAN THEO TỪNG KỲ KÊ KHAI{cqt_label}"
    ws_dc["A1"].font = title_font
    ws_dc.merge_cells("A1:J1")
    ws_dc.row_dimensions[1].height = 24

    headers = [
        "STT",
        "Kỳ kê khai",
        "Năm",
        "Tiểu mục",
        "Thuế phải nộp\ntheo Tờ khai",
        "Phát sinh trên\nSổ Tổng Quan",
        "Chênh lệch\n(Tờ khai - TQ)",
        "Số đã nộp theo kỳ\ntrên Sổ TQ",
        "Còn phải nộp theo kỳ\n(TQ - Đã nộp)",
        "Trạng thái đối chiếu",
    ]
    for ci, h in enumerate(headers, start=1):
        c = ws_dc.cell(2, ci, h)
        c.font = font_hdr
        c.fill = header_fill
        c.alignment = al_c
        c.border = border_all
    ws_dc.row_dimensions[2].height = 30

    cur_r = 3
    for idx, (pk, tm) in enumerate(all_keys, start=1):
        decl_expr = decl_map.get((pk, tm), (0, 0.0))[0]
        ps_val = round(tq_ps_map.get((pk, tm), 0.0))
        nop_val = round(tq_nop_map.get((pk, tm), 0.0))

        row_vals = [
            idx,
            pk[3],
            pk[0],
            tm,
            decl_expr,
            ps_val,
            f"=E{cur_r}-F{cur_r}",
            nop_val,
            f"=F{cur_r}-H{cur_r}",
            f'=IF(ABS(G{cur_r})>1000,"Lệch / Thiếu TK trên TQ",IF(I{cur_r}>1000,"Khớp TK - Còn nợ",IF(I{cur_r}<-1000,"Khớp TK - Nộp thừa","Khớp 100%")))',
        ]
        for ci, v in enumerate(row_vals, start=1):
            c = ws_dc.cell(cur_r, ci, v)
            c.font = font_bold if ci in (7, 9) else font_reg
            c.border = border_all
            if ci in (1, 2, 3, 4):
                c.alignment = al_c
                if ci == 4:
                    c.number_format = "@"
            elif ci == 10:
                c.alignment = al_l
            else:
                c.number_format = num_fmt
                c.alignment = al_r
        cur_r += 1

    if all_keys:
        r_end = cur_r - 1
        ws_dc.cell(cur_r, 2, "TỔNG CỘNG").alignment = al_c
        for ci, col_l in ((5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I")):
            ws_dc.cell(cur_r, ci, f"=SUM({col_l}3:{col_l}{r_end})")
        for ci in range(1, 11):
            c = ws_dc.cell(cur_r, ci)
            c.font = font_bold
            c.fill = sum_fill
            c.border = border_all
            if 5 <= ci <= 9:
                c.number_format = num_fmt
                c.alignment = al_r

    widths = {1: 6, 2: 14, 3: 8, 4: 11, 5: 19, 6: 19, 7: 18, 8: 19, 9: 20, 10: 26}
    for ci, w in widths.items():
        ws_dc.column_dimensions[get_column_letter(ci)].width = w
    ws_dc.freeze_panes = "A3"

    wb.save(output_path)
    return {"output_path": output_path, "reconciled": True, "rows": len(all_keys)}


def export_qhns_by_tax_type_period_files(
    excel_path: str,
    out_dir: str,
    tq_excel_path: Optional[str] = None,
    tq_ntnn_excel_path: Optional[str] = None,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    target_cqt: Optional[str] = "7901",
    base_prefix: str = "QHNS",
) -> List[str]:
    """
    Xuất RIÊNG phần Quan hệ ngân sách (QHNS) đối chiếu THEO TỪNG KỲ KÊ KHAI thành từng file Excel độc lập
    cho mỗi sắc thuế (chỉ xuất khi người dùng tích chọn):
      - {base_prefix}_GTGT.xlsx  (TM 1701 từ TQ, đối chiếu từng tháng/quý với 'So Thue GTGT Cuoi')
      - {base_prefix}_TNDN.xlsx  (TM 1052 từ TQ, đối chiếu theo từng kỳ báo cáo & năm với sheet năm C76)
      - {base_prefix}_TNCN.xlsx  (TM 1001 từ TQ + PP phân bổ CQT khác, đối chiếu theo từng kỳ báo cáo & năm với C90)
      - {base_prefix}_LPMB.xlsx  (TM 2862/2863/2864 từ TQ theo từng kỳ)
      - {base_prefix}_NTNN.xlsx  (TM 1701 & 1052 từ TQ_NTNN, đối chiếu từng kỳ với 'So Thue NTNN Cuoi')
    Mỗi file gồm 2 sheet:
      1. 'QHNS_TheoKy': Số dư đầu kỳ trước start_year + đối chiếu từng kỳ kê khai (Phát sinh TK, Phát sinh TQ,
         Chênh lệch, Đã nộp theo kỳ, Còn phải nộp kỳ, Dư lũy kế TQ, Trạng thái).
      2. 'ChiTiet_TQ': Toàn bộ dòng giao dịch gốc trên Sổ TQ của đúng sắc thuế đó (đã lọc đúng target_cqt).
    """
    os.makedirs(out_dir, exist_ok=True)
    norm_cqt = _normalize_cqt(target_cqt)

    def _load_tq_rows(path: Optional[str]) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        if not path or not os.path.isfile(path):
            return [], []
        wb_src = openpyxl.load_workbook(path, data_only=True)
        ws_src = None
        for cand in wb_src.sheetnames:
            if cand.strip().upper() in ("TQ", "TQ_NTNN", "TONG QUAN", "TỔNG QUAN", "NTNN"):
                ws_src = wb_src[cand]
                break
        if ws_src is None:
            ws_src = wb_src[wb_src.sheetnames[0]]
        all_r = parse_tq_rows(ws_src)
        return split_tq_rows_by_cqt(all_r, norm_cqt)

    main_tq, other_tq = _load_tq_rows(tq_excel_path)
    main_ntnn, _other_ntnn = _load_tq_rows(tq_ntnn_excel_path)

    wb_decl = openpyxl.load_workbook(excel_path, data_only=False) if (excel_path and os.path.isfile(excel_path)) else None
    wb_decl_v = openpyxl.load_workbook(excel_path, data_only=True) if (excel_path and os.path.isfile(excel_path)) else None

    detected_years = sorted([int(s) for s in wb_decl.sheetnames if re.fullmatch(r"20\d{2}", s)]) if wb_decl else []
    if year_from is not None and year_to is not None and year_from <= year_to:
        years = list(range(year_from, year_to + 1))
    elif detected_years:
        years = detected_years
    else:
        src_yr = main_tq or main_ntnn
        ht_yrs = sorted({x["ht_year"] for x in src_yr if 2015 <= x["ht_year"] <= 2035})
        years = ht_yrs[-6:] if len(ht_yrs) > 6 else ht_yrs
    if not years:
        return []
    start_year, end_year = years[0], years[-1]

    # Chuẩn bị bản đồ số kê khai theo từng kỳ từ workbook tổng hợp
    decl_gtgt_by_pk: Dict[Tuple[int, int, int, str], float] = {}
    decl_ntnn_by_pk_tm: Dict[Tuple[Tuple[int, int, int, str], str], float] = defaultdict(float)
    decl_tndn_by_yr: Dict[int, float] = {}
    decl_tncn_by_yr: Dict[int, float] = {}

    if wb_decl is not None and wb_decl_v is not None:
        if "So Thue GTGT Cuoi" in wb_decl.sheetnames:
            ws_g = wb_decl["So Thue GTGT Cuoi"]
            gtgt_col = "AJ"
            for ci in range(1, ws_g.max_column + 1):
                h_txt = str(ws_g.cell(1, ci).value or "")
                if "(40)" in h_txt and "(40a)" not in h_txt and "(40b)" not in h_txt:
                    gtgt_col = get_column_letter(ci)
                    break
            for r in range(3, ws_g.max_row + 1):
                if "tổng cộng" in str(ws_g.cell(r, 5).value or "").lower():
                    continue
                ky_raw = str(ws_g.cell(r, 7).value or "").strip()
                if not ky_raw:
                    continue
                pk = normalize_period_key(ky_raw)
                if start_year <= pk[0] <= end_year:
                    decl_gtgt_by_pk[pk] = _eval_simple_formula(wb_decl, wb_decl_v, "So Thue GTGT Cuoi", f"{gtgt_col}{r}")

        ntnn_decl_sheet = (
            "So Thue NTNN Cuoi (Lech ky)"
            if "So Thue NTNN Cuoi (Lech ky)" in wb_decl.sheetnames
            else ("So Thue NTNN Cuoi" if "So Thue NTNN Cuoi" in wb_decl.sheetnames else None)
        )
        if ntnn_decl_sheet:
            ws_n = wb_decl[ntnn_decl_sheet]
            for r in range(2, ws_n.max_row + 1):
                ky_raw = str(ws_n.cell(r, 6).value or "").strip()
                if not ky_raw or "tổng cộng" in ky_raw.lower():
                    continue
                pk = normalize_period_key(ky_raw)
                if start_year <= pk[0] <= end_year:
                    decl_ntnn_by_pk_tm[(pk, "1701")] += _eval_simple_formula(wb_decl, wb_decl_v, ntnn_decl_sheet, f"O{r}")
                    decl_ntnn_by_pk_tm[(pk, "1052")] += _eval_simple_formula(wb_decl, wb_decl_v, ntnn_decl_sheet, f"Q{r}")

        for y in years:
            sname = str(y)
            if sname in wb_decl.sheetnames:
                decl_tndn_by_yr[y] = _eval_simple_formula(wb_decl, wb_decl_v, sname, "C76")
                decl_tncn_by_yr[y] = _eval_simple_formula(wb_decl, wb_decl_v, sname, "C90")

    payment_gds = ("C2", "HC", "NT", "C5")
    lpmb_tms = ("2862", "2863", "2864", "1801", "1802", "1803", "1804")

    def _build_single_tax_qhns_file(
        code: str,
        title: str,
        rows_src: List[Dict[str, Any]],
        tms: Tuple[str, ...],
        ps_gds: Tuple[str, ...],
        decl_lookup: Dict[Tuple[Tuple[int, int, int, str], str], float],
        extra_pp_rows: Optional[List[Dict[str, Any]]] = None,
    ) -> Optional[str]:
        tax_rows = [r for r in rows_src if r["tm"] in tms and r["gd"] != "BT"]
        pp_rows = [r for r in (extra_pp_rows or []) if r["tm"] in tms and r["gd"] == "PP"]
        if not tax_rows and not pp_rows and not decl_lookup:
            return None

        c2_vals = [x["so_tien"] for x in rows_src if x["gd"] == "C2" and x["so_tien"] != 0]
        c2_neg = bool(c2_vals and sum(1 for v in c2_vals if v < 0) >= len(c2_vals) / 2)

        def _s_ps(it: Dict[str, Any]) -> float:
            v = it["so_tien"]
            nc = it.get("no_co", "")
            if nc in ("C", "N") and not c2_neg and v >= 0:
                return v if nc == "N" else -v
            return v

        def _s_pay(it: Dict[str, Any]) -> float:
            v = it["so_tien"]
            nc = it.get("no_co", "")
            if nc in ("C", "N") and not c2_neg and v >= 0:
                return v if nc == "C" else -v
            return -v if c2_neg else v

        opening_bal = 0.0
        tq_ps_by_key: Dict[Tuple[Tuple[int, int, int, str], str], float] = defaultdict(float)
        tq_nop_by_key: Dict[Tuple[Tuple[int, int, int, str], str], float] = defaultdict(float)

        for it in tax_rows:
            pk = normalize_period_key(it["ky"], it["ngay_ht"])
            ky_y = pk[0]
            ht_y = it["ht_year"]
            if ky_y < start_year:
                if it["gd"] in payment_gds:
                    if ht_y < start_year:
                        opening_bal -= _s_pay(it)
                    elif start_year <= ht_y <= end_year:
                        # Khoản nộp trong kỳ kiểm tra cho kỳ trước -> tính vào cột đã nộp theo kỳ gốc
                        tq_nop_by_key[(pk, it["tm"])] += _s_pay(it)
                else:
                    opening_bal += _s_ps(it)
            elif start_year <= ky_y <= end_year:
                if it["gd"] in ps_gds:
                    tq_ps_by_key[(pk, it["tm"])] += _s_ps(it)
                elif it["gd"] in payment_gds:
                    tq_nop_by_key[(pk, it["tm"])] += _s_pay(it)
                else:
                    tq_ps_by_key[(pk, it["tm"])] += _s_ps(it)

        for it in pp_rows:
            pk = normalize_period_key(it["ky"], it["ngay_ht"])
            if start_year <= pk[0] <= end_year:
                tq_ps_by_key[(pk, it["tm"])] += _s_ps(it)

        all_keys = sorted(
            set(decl_lookup.keys()) | set(tq_ps_by_key.keys()) | set(tq_nop_by_key.keys()),
            key=lambda x: (x[0], x[1]),
        )

        wb_out = openpyxl.Workbook()
        ws = wb_out.active
        ws.title = "QHNS_TheoKy"

        thin = Side(style="thin", color="000000")
        b_all = Border(left=thin, right=thin, top=thin, bottom=thin)
        f_title = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")
        f_hdr = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
        f_reg = Font(name="Times New Roman", size=10)
        f_bold = Font(name="Times New Roman", size=10, bold=True)
        fill_hdr = PatternFill("solid", fgColor="1F4E78")
        fill_sub = PatternFill("solid", fgColor="D9E1F2")
        fill_open = PatternFill("solid", fgColor="F2F2F2")
        al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
        al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
        al_r = Alignment(horizontal="right", vertical="center")
        num_fmt = "#,##0"

        cqt_txt = f" - CQT quản lý: {norm_cqt}" if norm_cqt else ""
        ws["A1"] = f"ĐỐI CHIẾU QUAN HỆ NGÂN SÁCH THEO KỲ KÊ KHAI - {title} ({start_year}–{end_year}{cqt_txt})"
        ws["A1"].font = f_title
        ws.merge_cells("A1:K1")
        ws.row_dimensions[1].height = 24

        headers = [
            "STT",
            "Kỳ kê khai",
            "Năm",
            "Tiểu mục",
            "Thuế phải nộp\ntheo Tờ khai",
            "Phát sinh trên\nSổ Tổng Quan",
            "Chênh lệch\n(Tờ khai - TQ)",
            "Số đã nộp theo kỳ\ntrên Sổ TQ",
            "Còn phải nộp\ntheo kỳ (TQ)",
            "Dư lũy kế\nđến kỳ (TQ)",
            "Trạng thái đối chiếu",
        ]
        for ci, h in enumerate(headers, start=1):
            c = ws.cell(2, ci, h)
            c.font = f_hdr
            c.fill = fill_hdr
            c.alignment = al_c
            c.border = b_all
        ws.row_dimensions[2].height = 30

        # Dòng 3: Số dư đầu kỳ trước start_year
        ws.cell(3, 2, f"Số dư đầu kỳ trước {start_year}").alignment = al_l
        ws.cell(3, 10, round(opening_bal))
        for ci in range(1, 12):
            c = ws.cell(3, ci)
            c.font = f_bold
            c.fill = fill_open
            c.border = b_all
            if ci in (5, 6, 7, 8, 9, 10):
                c.number_format = num_fmt
                c.alignment = al_r

        cur_r = 4
        for idx, (pk, tm) in enumerate(all_keys, start=1):
            decl_v = round(decl_lookup.get((pk, tm), 0.0))
            ps_v = round(tq_ps_by_key.get((pk, tm), 0.0))
            nop_v = round(tq_nop_by_key.get((pk, tm), 0.0))
            row_vals = [
                idx,
                pk[3],
                pk[0],
                tm,
                decl_v,
                ps_v,
                f"=E{cur_r}-F{cur_r}",
                nop_v,
                f"=F{cur_r}-H{cur_r}",
                f"=J{cur_r - 1}+F{cur_r}-H{cur_r}",
                f'=IF(ABS(G{cur_r})>1000,"Lệch / Thiếu TK trên TQ",IF(I{cur_r}>1000,"Khớp TK - Còn nợ",IF(I{cur_r}<-1000,"Khớp TK - Nộp thừa","Khớp 100%")))',
            ]
            for ci, v in enumerate(row_vals, start=1):
                c = ws.cell(cur_r, ci, v)
                c.font = f_bold if ci in (7, 9, 10) else f_reg
                c.border = b_all
                if ci in (1, 2, 3, 4):
                    c.alignment = al_c
                    if ci == 4:
                        c.number_format = "@"
                elif ci == 11:
                    c.alignment = al_l
                else:
                    c.number_format = num_fmt
                    c.alignment = al_r
            cur_r += 1

        # Dòng TỔNG CỘNG
        if all_keys:
            r_end = cur_r - 1
            ws.cell(cur_r, 2, "TỔNG CỘNG").alignment = al_c
            for ci, col_l in ((5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I")):
                ws.cell(cur_r, ci, f"=SUM({col_l}4:{col_l}{r_end})")
            ws.cell(cur_r, 10, f"=J{r_end}")
            for ci in range(1, 12):
                c = ws.cell(cur_r, ci)
                c.font = f_bold
                c.fill = fill_sub
                c.border = b_all
                if 5 <= ci <= 10:
                    c.number_format = num_fmt
                    c.alignment = al_r

        widths = {1: 6, 2: 14, 3: 8, 4: 11, 5: 18, 6: 18, 7: 18, 8: 18, 9: 18, 10: 18, 11: 24}
        for ci, w in widths.items():
            ws.column_dimensions[get_column_letter(ci)].width = w
        ws.freeze_panes = "A4"

        # Sheet 2: Chi tiết giao dịch gốc trên Sổ TQ cho đúng sắc thuế này
        ws_raw = wb_out.create_sheet("ChiTiet_TQ")
        raw_hdrs = ["STT", "Dòng TQ", "Mã CQT", "Kỳ báo cáo", "Ký hiệu GD", "Tiểu mục", "Số tiền", "Ngày hạch toán", "Ghi chú"]
        for ci, h in enumerate(raw_hdrs, start=1):
            c = ws_raw.cell(1, ci, h)
            c.font = f_hdr
            c.fill = fill_hdr
            c.alignment = al_c
            c.border = b_all
        ws_raw.row_dimensions[1].height = 24
        combined_raw = [(it, "") for it in tax_rows] + [(it, "PP phân bổ CQT khác") for it in pp_rows]
        for idx, (it, note) in enumerate(combined_raw, start=1):
            rr = idx + 1
            vals = [idx, it["row"], it["cqt"], it["ky"], it["gd"], it["tm"], round(it["so_tien"]), it["ngay_ht"], note]
            for ci, v in enumerate(vals, start=1):
                c = ws_raw.cell(rr, ci, v)
                c.font = f_reg
                c.border = b_all
                if ci in (3, 4, 6):
                    c.number_format = "@"
                    c.alignment = al_c
                elif ci == 7:
                    c.number_format = num_fmt
                    c.alignment = al_r
                elif ci == 8:
                    c.number_format = "dd/mm/yyyy"
                    c.alignment = al_c
                elif ci in (1, 2, 5):
                    c.alignment = al_c
                else:
                    c.alignment = al_l
        for ci, w in {1: 6, 2: 9, 3: 10, 4: 12, 5: 12, 6: 11, 7: 18, 8: 14, 9: 24}.items():
            ws_raw.column_dimensions[get_column_letter(ci)].width = w
        ws_raw.freeze_panes = "A2"

        out_file = os.path.join(out_dir, f"{base_prefix}_{code}.xlsx")
        wb_out.save(out_file)
        return out_file

    created_files: List[str] = []
    if main_tq:
        f_gtgt = _build_single_tax_qhns_file(
            "GTGT",
            "THUẾ GTGT (TM 1701)",
            main_tq,
            ("1701",),
            ("TK",),
            {(pk, "1701"): v for pk, v in decl_gtgt_by_pk.items()},
        )
        if f_gtgt:
            created_files.append(f_gtgt)

        f_tndn = _build_single_tax_qhns_file(
            "TNDN",
            "THUẾ TNDN (TM 1052)",
            main_tq,
            ("1052",),
            ("TK", "TQ", "QT"),
            {((y, 3, 0, f"Năm {y}"), "1052"): v for y, v in decl_tndn_by_yr.items()},
        )
        if f_tndn:
            created_files.append(f_tndn)

        f_tncn = _build_single_tax_qhns_file(
            "TNCN",
            "THUẾ TNCN (TM 1001)",
            main_tq,
            ("1001",),
            ("TK", "TH", "TQ", "PP"),
            {((y, 3, 0, f"Năm {y}"), "1001"): v for y, v in decl_tncn_by_yr.items()},
            extra_pp_rows=other_tq,
        )
        if f_tncn:
            created_files.append(f_tncn)

        f_lpmb = _build_single_tax_qhns_file(
            "LPMB",
            "LỆ PHÍ MÔN BÀI (TM 2862/2863/2864)",
            main_tq,
            lpmb_tms,
            ("TK", "TQ", "QT"),
            {},
        )
        if f_lpmb:
            created_files.append(f_lpmb)

    if main_ntnn:
        f_ntnn = _build_single_tax_qhns_file(
            "NTNN",
            "THUẾ NHÀ THẦU NƯỚC NGOÀI (TM 1701 & 1052)",
            main_ntnn,
            ("1701", "1052"),
            ("TK", "TQ", "QT"),
            dict(decl_ntnn_by_pk_tm),
        )
        if f_ntnn:
            created_files.append(f_ntnn)

    return created_files
