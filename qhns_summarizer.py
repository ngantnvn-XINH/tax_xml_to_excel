# -*- coding: utf-8 -*-
"""
Module tổng hợp sheet Tổng quan (TQ), đối chiếu Tờ khai vs Tổng quan trên các sheet Năm (2020..2025)
và lập biểu Quan Hệ Ngân Sách (QHNS) chuẩn nghiệp vụ thanh tra - kiểm tra thuế.

Tích hợp trực tiếp vào ứng dụng Desktop (app.py) cho cả 2 chế độ:
1. Chạy trọn gói: Xuất XML -> Sheet năm -> Tự động ghép TQ, đối chiếu F/G & sinh QHNS.
2. Chạy bổ sung: Chọn file Excel tổng hợp đã xuất + file Sổ Tổng Quan (TQ) -> Tự động cập nhật TQ, đối chiếu & sinh QHNS.
"""
from __future__ import annotations

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


def _norm_header(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).strip().lower()
    # Chuẩn hóa Unicode tổ hợp/dựng sẵn đơn giản cho các từ khóa cột TMS
    import unicodedata
    s = unicodedata.normalize("NFC", s)
    return re.sub(r"\s+", " ", s)


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
            if "kỳ báo cáo" in h or "ky bao cao" in h or "kỳ thuế" == h:
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
    return 1, {"ky": 2, "gd": 4, "tm": 5, "so_tien": 7, "ngay_ht": 10}


def parse_tq_rows(ws_tq) -> List[Dict[str, Any]]:
    """Đọc toàn bộ các dòng giao dịch hợp lệ từ sheet TQ."""
    header_row, col_map = detect_tq_layout(ws_tq)
    rows: List[Dict[str, Any]] = []
    for r in range(header_row + 1, ws_tq.max_row + 1):
        so_tien = ws_tq.cell(r, col_map["so_tien"]).value
        ngay_ht = parse_date(ws_tq.cell(r, col_map["ngay_ht"]).value)
        gd = str(ws_tq.cell(r, col_map["gd"]).value or "").strip().upper()
        tm = str(ws_tq.cell(r, col_map["tm"]).value or "").strip()
        ky = str(ws_tq.cell(r, col_map["ky"]).value or "").strip().upper()
        no_co = ""
        if "no_co" in col_map:
            no_co = str(ws_tq.cell(r, col_map["no_co"]).value or "").strip().upper()

        if not isinstance(so_tien, (int, float)) or not ngay_ht or not gd or not tm:
            continue

        rows.append({
            "row": r,
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


def compute_tq_metrics(tq_rows: List[Dict[str, Any]], start_year: int, end_year: int) -> Dict[str, Any]:
    """
    Tính toán toàn bộ chỉ số từ dữ liệu TQ theo đúng chuẩn nghiệp vụ thanh tra - kiểm tra thuế:
    - GTGT (1701): Phát sinh = TK theo Kỳ báo cáo; Đã nộp = C2 (+ HC nếu có) theo Năm Ngày hạch toán.
    - TNDN (1052): Phát sinh = TK + QT theo Kỳ báo cáo; Đã nộp = C2 (+ HC nếu có) theo Năm Ngày hạch toán.
    - TNCN (1001): Phát sinh = TK + TH + TQ theo Kỳ báo cáo; Đã nộp = C2 (+ HC nếu có) theo Năm Ngày hạch toán.
    - Đầu kỳ năm start_year: Các khoản Kỳ báo cáo < start_year nhưng nộp/hạch toán trong start_year..end_year.
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

    for item in tq_rows:
        tm = item["tm"]
        gd = item["gd"]
        ky = item["ky"]
        ky_y = item["ky_year"]
        ht_y = item["ht_year"]
        val = item["so_tien"]

        # Bỏ qua bút toán bù trừ nội bộ (BT) vì mỗi cặp BT luôn tự triệt tiêu về 0
        if gd == "BT":
            continue

        # 1. Ghi nhận số đầu kỳ Biên bản: TẤT CẢ số nộp (C2/HC/NT) & QT của mọi kỳ trước (< start_year) nộp/hạch toán trong kỳ kiểm tra
        if ky_y is not None and ky_y < start_year and start_year <= ht_y <= end_year:
            if tm == "1701":
                if gd == "QT":
                    metrics["dk_gtgt_qt_truoc"] += abs(val)
                    metrics["dk_gtgt_qt_notes"].append(f"QT {ky_y} ({ky}) hạch toán năm {ht_y}")
                elif gd in ("C2", "HC", "NT"):
                    metrics["dk_gtgt_c2_prev"] += payment_amount(item)
                    if ky and ky not in metrics["dk_gtgt_c2_notes"]:
                        metrics["dk_gtgt_c2_notes"].append(ky)
            elif tm == "1001":
                if gd in ("C2", "HC", "NT"):
                    metrics["dk_tncn_c2_prev"] += payment_amount(item)
                    if ky and ky not in metrics["dk_tncn_c2_notes"]:
                        metrics["dk_tncn_c2_notes"].append(ky)
                elif gd in ("TQ", "TH") and ky == f"{prev_yy}CN":
                    metrics["dk_tncn_qt_prev"] += val
                    if ky and ky not in metrics["dk_tncn_qt_notes"]:
                        metrics["dk_tncn_qt_notes"].append(ky)
            elif tm == "1052":
                if gd == "QT":
                    metrics["dk_tndn_h"] += abs(val)
                elif gd in ("C2", "HC", "NT"):
                    metrics["dk_tndn_i"] += payment_amount(item)
            elif tm in LPMB_TMS and gd in ("TK", "TQ", "QT"):
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
            elif tm == "1001" and gd in ("TK", "TH", "TQ"):
                metrics["tncn_tk_th_tq_by_ky"][ky_y] += val
                in_bucket_a = True
                if gd == "TK" and ky:
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
    result: Dict[int, Dict[str, float]] = {y: {"gtgt_c23": 0.0, "tncn_c90": 0.0, "tndn_c76": 0.0} for y in years}

    for y in years:
        sname = str(y)
        if sname not in wb_f.sheetnames:
            continue
        result[y]["gtgt_c23"] = _eval_simple_formula(wb_f, wb_v, sname, "C23")
        result[y]["tncn_c90"] = _eval_simple_formula(wb_f, wb_v, sname, "C90")
        result[y]["tndn_c76"] = _eval_simple_formula(wb_f, wb_v, sname, "C76")

    return result


def integrate_tq_and_qhns_into_workbook(
    excel_path: str,
    tq_excel_path: str,
    so_nop_excel_path: Optional[str] = None,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    output_path: Optional[str] = None,
    cutoff_date: Optional[datetime.date] = None,
    tq_ntnn_excel_path: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Hàm tích hợp chính:
    1. Đọc Sổ Tổng Quan (TQ) từ tq_excel_path và chèn vào workbook excel_path.
    2. Cập nhật cột đối chiếu F & G kèm Comment trên các sheet Năm ('2020'..'2025') cho GTGT (Dòng 23) và TNCN (Dòng 90).
    3. Sinh sheet QHNS chuẩn 01 biểu duy nhất theo đúng nghiệp vụ thanh tra - kiểm tra thuế.
    4. Lưu ra output_path (mặc định ghi đè/cập nhật chính excel_path).
    """
    if output_path is None:
        output_path = excel_path

    # 1. Đọc file TQ nguồn
    wb_tq = openpyxl.load_workbook(tq_excel_path, data_only=True)
    ws_tq_src = None
    for cand in wb_tq.sheetnames:
        if cand.strip().upper() in ("TQ", "TONG QUAN", "TỔNG QUAN"):
            ws_tq_src = wb_tq[cand]
            break
    if ws_tq_src is None:
        ws_tq_src = wb_tq[wb_tq.sheetnames[0]]

    tq_rows = parse_tq_rows(ws_tq_src)
    if not tq_rows:
        raise ValueError("Không tìm thấy dữ liệu giao dịch hợp lệ trong file Sổ Tổng Quan (TQ).")

    # 2. Mở workbook tổng hợp tờ khai
    wb = openpyxl.load_workbook(excel_path, data_only=False)

    # Xác định danh sách năm cần lên QHNS
    detected_years = sorted([int(s) for s in wb.sheetnames if re.fullmatch(r"20\d{2}", s)])
    if year_from is not None and year_to is not None and year_from <= year_to:
        years = list(range(year_from, year_to + 1))
    elif detected_years:
        years = detected_years
    else:
        ht_years = sorted({x["ht_year"] for x in tq_rows if 2015 <= x["ht_year"] <= 2035})
        years = ht_years[-6:] if len(ht_years) > 6 else ht_years

    start_year, end_year = years[0], years[-1]
    metrics = compute_tq_metrics(tq_rows, start_year, end_year)
    decl_vals = _eval_year_sheet_values(excel_path, years)

    # 2b. Đọc Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) nếu người dùng cung cấp
    ntnn_metrics: Optional[Dict[str, Any]] = None
    ws_ntnn_src = None
    if tq_ntnn_excel_path and os.path.isfile(tq_ntnn_excel_path):
        wb_ntnn = openpyxl.load_workbook(tq_ntnn_excel_path, data_only=True)
        for cand in wb_ntnn.sheetnames:
            if cand.strip().upper() in ("TQ", "TQ_NTNN", "TONG QUAN", "TỔNG QUAN", "NTNN"):
                ws_ntnn_src = wb_ntnn[cand]
                break
        if ws_ntnn_src is None:
            ws_ntnn_src = wb_ntnn[wb_ntnn.sheetnames[0]]
        ntnn_rows = parse_tq_rows(ws_ntnn_src)
        if ntnn_rows:
            ntnn_metrics = compute_tq_metrics(ntnn_rows, start_year, end_year)

    # 3. Chép sheet TQ (và TQ_NTNN nếu có) vào workbook đích
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

    for y in years:
        sname = str(y)
        if sname not in wb.sheetnames:
            gtgt_missing_info[y] = None
            tndn_missing_info[y] = None
            tncn_missing_info[y] = None
            continue
        ws_y = wb[sname]

        # --- Dòng 23: Thuế GTGT phải nộp trong kỳ ---
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
        f90_val = round(metrics["tncn_tk_th_tq_by_ky"][y])
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

        if diff_tncn > 1000:
            missing_p = sorted(metrics["tncn_c2_periods"][y] - metrics["tncn_tk_periods"][y])
            miss_str = ", ".join(missing_p) if missing_p else f"năm {y}"
            ws_y["F90"].comment = Comment(f"Phát sinh TK; TH; TQ trên tổng quan rà soát thiếu TK {miss_str}", "QHNS")
            tncn_missing_info[y] = miss_str
            alerts.append(f"Năm {y} - Thuế TNCN: Tờ khai QTT ({c90_val:,.0f}) > Tổng quan ({f90_val:,.0f}), chênh lệch {diff_tncn:,.0f} -> Đưa vào Cột (5) Số xử lý qua thanh tra (thiếu TK {miss_str}).")
        else:
            ws_y["F90"].comment = Comment("Phát sinh TK; TH; TQ trên tổng quan", "QHNS")
            tncn_missing_info[y] = None

        ws_y.column_dimensions["F"].width = 18
        ws_y.column_dimensions["G"].width = 18

        # --- Dòng 97 & 98: Đối chiếu Thuế Nhà thầu (GTGT 1701 & TNDN 1052) trên TQ_NTNN nếu có ---
        if ntnn_metrics is not None:
            f97_val = round(ntnn_metrics["gtgt_tk_by_ky"][y])
            f98_val = round(ntnn_metrics["tndn_tk_qt_by_ky"][y])
            ws_y["F97"].value = f97_val
            ws_y["F97"].number_format = num_fmt
            ws_y["F97"].font = font_bold
            ws_y["F97"].comment = Comment("Phát sinh Thuế GTGT Nhà thầu (TM 1701) trên Sổ Tổng Quan Nhà thầu (TQ_NTNN)", "QHNS")
            ws_y["G97"].value = "=C97-F97"
            ws_y["G97"].number_format = num_fmt
            ws_y["G97"].font = font_bold

            ws_y["F98"].value = f98_val
            ws_y["F98"].number_format = num_fmt
            ws_y["F98"].font = font_bold
            ws_y["F98"].comment = Comment("Phát sinh Thuế TNDN Nhà thầu (TM 1052) trên Sổ Tổng Quan Nhà thầu (TQ_NTNN)", "QHNS")
            ws_y["G98"].value = "=C98-F98"
            ws_y["G98"].number_format = num_fmt
            ws_y["G98"].font = font_bold

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

        if has_sheet:
            if tncn_missing_info.get(y):
                ws_q.cell(r_tn, 4, f"='{sname}'!F90")
                ws_q.cell(r_tn, 4).comment = Comment(
                    f"Vì bị thiếu tờ khai {tncn_missing_info[y]} nên số phát sinh trên tổng quan bị thiếu, nên khi thể hiện số phát sinh sẽ lấy trên tổng quan",
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
        else:
            ws_q.cell(r_tn, 4, round(metrics["tncn_tk_th_tq_by_ky"][y]))

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

            ws_q.cell(r_nt_g, 4, ntg_ps)
            ws_q.cell(r_nt_g, 6, ntg_nop)
            if ntnn_sp.get(("1701", y)):
                ws_q.cell(r_nt_g, 6).comment = Comment(
                    "GD C2 + NT/HC/C5 trên TQ_NTNN (bao gồm: " + "; ".join(ntnn_sp[("1701", y)]) + ")",
                    "QHNS",
                )
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
                f90_eff = round(metrics["tncn_tk_th_tq_by_ky"][y])
                if c90_eff == 0 and f90_eff > 0:
                    c90_eff = f90_eff
                rounding_diff["1001"] += c90_eff - f90_eff
                rounding_diff_refs["1001"].append(f"'{sname}'!G90")

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

    other_tms = metrics.get("other_tms", [])
    ntnn_other_tms = ntnn_metrics.get("other_tms", []) if ntnn_metrics is not None else []
    ntnn_core_count = 2 if ntnn_metrics is not None else 0

    # Xây dựng Bảng 2 trước trong bộ nhớ vị trí để Bảng 1 có thể link trực tiếp sang số dư năm cuối của các tiểu mục khác
    b1_start_row = 3
    b1_core_count = 4 + ntnn_core_count
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

    b2_rows_spec: List[Tuple[str, str, str, Dict[int, float], Dict[int, float], float]] = [
        ("1701", "1", "Thuế GTGT (1701)", metrics["gtgt_tk_by_ky"], metrics["gtgt_c2_by_ht"], metrics["opening_exact"]["1701"]),
        ("1052", "2", "Thuế TNDN (1052)", metrics["tndn_tk_qt_by_ky"], metrics["tndn_c2_by_ht"], metrics["opening_exact"]["1052"]),
        ("1001", "3", "Thuế TNCN (1001)", metrics["tncn_tk_th_tq_by_ky"], metrics["tncn_c2_by_ht"], metrics["opening_exact"]["1001"]),
        ("LPMB", "4", "Lệ phí môn bài (2862/2863/2864)", metrics["lpmb_ps_by_ky"], metrics["lpmb_c2_by_ht"], metrics["opening_exact"]["LPMB"]),
    ]
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
    dc_specs = [
        ("Thuế GTGT", "1701", "1701", last_rows.get("gtgt"), gtgt_e_cells, metrics["gtgt_tk_by_ky"], metrics["gtgt_c2_by_ht"], metrics["opening_exact"]["1701"], True),
        ("Thuế TNDN", "1052", "1052", last_rows.get("tndn"), tndn_e_cells, metrics["tndn_tk_qt_by_ky"], metrics["tndn_c2_by_ht"], metrics["opening_exact"]["1052"], True),
        ("Thuế TNCN", "1001", "1001", last_rows.get("tncn"), tncn_e_cells, metrics["tncn_tk_th_tq_by_ky"], metrics["tncn_c2_by_ht"], metrics["opening_exact"]["1001"], True),
        ("Lệ phí môn bài", "LPMB", "2862/2863/2864", last_rows.get("lpmb"), [], metrics["lpmb_ps_by_ky"], metrics["lpmb_c2_by_ht"], metrics["opening_exact"]["LPMB"], True),
    ]
    if ntnn_metrics is not None:
        dc_specs.append(
            ("Thuế GTGT Nhà thầu (TQ_NTNN)", "NT_1701", "1701 (NTNN)", last_rows.get("nt_gtgt"), [], ntnn_metrics["gtgt_tk_by_ky"], ntnn_metrics["gtgt_c2_by_ht"], ntnn_metrics["opening_exact"]["1701"], True)
        )
        dc_specs.append(
            ("Thuế TNDN Nhà thầu (TQ_NTNN)", "NT_1052", "1052 (NTNN)", last_rows.get("nt_tndn"), [], ntnn_metrics["tndn_tk_qt_by_ky"], ntnn_metrics["tndn_c2_by_ht"], ntnn_metrics["opening_exact"]["1052"], True)
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
    }
