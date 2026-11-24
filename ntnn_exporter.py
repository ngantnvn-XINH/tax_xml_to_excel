# -*- coding: utf-8 -*-
"""
Module tổng hợp Tờ khai Thuế Nhà Thầu Nước Ngoài (01/NTNN, 02/NTNN, 03/NTNN, 04/NTNN):
- Tự động phân biệt CHÍNH XÁC chuẩn mẫu theo từng thời kỳ:
    * TT103/2014 (<=2021): ct6=DT GTGT, ct7=Tỷ lệ GTGT, ct8=Thuế GTGT, ct9=DT TNDN, ct10=Tỷ lệ TNDN, ct11=Miễn giảm, ct12=Thuế TNDN, ct13=Tổng thuế
    * TT80/2021 (>=2022): ct6=Ngày thanh toán, ct7=DT GTGT, ct8=Tỷ lệ GTGT, ct9=Thuế GTGT, ct10=DT TNDN, ct11=Tỷ lệ TNDN, ct12=Miễn giảm, ct13=Thuế TNDN, ct14=Tổng thuế
    * Semantic tags (dthuTinhThueGTGT, thueGTGTPNop, dthuTinhThueTNDN, thueTNDNPNop...): đọc trực tiếp không phụ thuộc số thứ tự cột.
- Sinh 2 sheet chuẩn tương tự Thuế GTGT:
    1. 'So Thue NTNN TH': Tổng hợp toàn bộ tờ khai nhà thầu (chính thức + bổ sung) theo dòng thời gian.
    2. 'So Thue NTNN Cuoi': Chốt tờ khai hiệu lực cuối cùng theo từng kỳ/ngày phát sinh, nhóm theo từng năm
       kèm dòng 'Tổng cộng năm YYYY' (=SUM(...)) bên dưới mỗi năm để nối công thức tự động
       vào Bảng 4 (C95 Doanh thu chưa gồm GTGT, C97 GTGT NTNN, C98 TNDN NTNN) trên từng sheet năm.
"""

from __future__ import annotations

import datetime
import re
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional, Set, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from tax_parser import TaxDeclaration


def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _to_float(val: Any) -> float:
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _looks_like_date(text: Optional[str]) -> bool:
    if not text:
        return False
    s = text.strip()
    return bool(re.match(r"^(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4}|\d{2}-\d{2}-\d{4})$", s))


def _get_main_declaration_root(root: ET.Element) -> ET.Element:
    """Ưu tiên quét trong <CTieuTKhaiChinh> để tuyệt đối không đọc nhầm số liệu phụ lục KHBS."""
    for node in root.iter():
        if _strip_ns(node.tag) == "CTieuTKhaiChinh":
            return node
    return root


def _iter_non_khbs(elem: ET.Element):
    """Duyệt cây XML, bỏ qua hoàn toàn các nhánh PLuc / KHBS / ChuKyDienTu."""
    tag_l = _strip_ns(elem.tag).lower()
    if any(skip in tag_l for skip in ("khbs", "chukydientu", "pluc")):
        return
    yield elem
    for ch in elem:
        yield from _iter_non_khbs(ch)


def detect_ntnn_schema(
    root: ET.Element,
    ma_tkhai: str = "",
    mo_ta_bmau: str = "",
    ten_tkhai: str = "",
    year: Optional[int] = None,
) -> str:
    """
    Nhận diện chuẩn mẫu tờ khai 01/NTNN để tránh lệch cột giữa TT103/2014 (ct6..ct13) và TT80/2021 (ct7..ct14).
    Trả về: 'TT80' hoặc 'TT103'.
    """
    scan_root = _get_main_declaration_root(root)

    # Tín hiệu 1: Có thẻ ct14 / tong_ct14 -> CHẮC CHẮN là TT80/2021 (TT103 chỉ có tối đa 13 cột)
    leaf_tags: Set[str] = set()
    ct6_is_date = False
    ct6_num_max = 0.0
    ct7_num_max = 0.0
    tt103_math_hits = 0
    tt80_math_hits = 0

    for el in _iter_non_khbs(scan_root):
        t = _strip_ns(el.tag).lower()
        if len(el) == 0:
            leaf_tags.add(t)
            if t in ("ct6", "ct06", "ct_6") and el.text and el.text.strip():
                if _looks_like_date(el.text):
                    ct6_is_date = True
                else:
                    ct6_num_max = max(ct6_num_max, _to_float(el.text))
            elif t in ("ct7", "ct07", "ct_7") and el.text and el.text.strip():
                ct7_num_max = max(ct7_num_max, _to_float(el.text))
        else:
            child_map: Dict[str, str] = {
                _strip_ns(ch.tag).lower(): (ch.text or "").strip()
                for ch in el
                if len(ch) == 0 and ch.text and ch.text.strip()
            }
            if not child_map:
                continue
            v6 = _to_float(child_map.get("ct6") or child_map.get("ct06") or child_map.get("ct_6"))
            v7 = _to_float(child_map.get("ct7") or child_map.get("ct07") or child_map.get("ct_7"))
            v8 = _to_float(child_map.get("ct8") or child_map.get("ct08") or child_map.get("ct_8"))
            v9 = _to_float(child_map.get("ct9") or child_map.get("ct09") or child_map.get("ct_9"))
            v12 = _to_float(child_map.get("ct12") or child_map.get("ct_12"))
            v13 = _to_float(child_map.get("ct13") or child_map.get("ct_13"))
            v14 = _to_float(child_map.get("ct14") or child_map.get("ct_14"))

            # Kiểm tra toán học TT80: ct9 ≈ ct7*ct8/100 hoặc ct14 ≈ ct9+ct13
            if v14 > 0 and abs(v14 - (v9 + v13)) <= max(10.0, v14 * 0.001):
                tt80_math_hits += 2
            if v7 > 100 and 0 < v8 <= 100 and v9 > 0 and abs(v9 - (v7 * v8 / 100.0)) <= max(10.0, v9 * 0.01):
                tt80_math_hits += 1

            # Kiểm tra toán học TT103: ct8 ≈ ct6*ct7/100 hoặc ct13 ≈ ct8+ct12 (khi không có ct14)
            if v14 == 0 and v13 > 0 and (v8 > 0 or v12 > 0) and abs(v13 - (v8 + v12)) <= max(10.0, v13 * 0.001):
                tt103_math_hits += 2
            if v6 > 100 and 0 < v7 <= 100 and v8 > 0 and abs(v8 - (v6 * v7 / 100.0)) <= max(10.0, v8 * 0.01):
                tt103_math_hits += 1

    if any(t in leaf_tags for t in ("ct14", "ct_14", "tong_ct14", "ct14_tong")):
        return "TT80"
    if ct6_is_date:
        return "TT80"

    # Tín hiệu 2: Mô tả biểu mẫu / tên tờ khai trong <TKhaiThue>
    meta = f"{mo_ta_bmau} {ten_tkhai}".upper()
    if "80/2021" in meta or "TT80" in meta:
        return "TT80"
    if "103/2014" in meta or "156/2013" in meta or "TT103" in meta or "TT156" in meta:
        return "TT103"

    # Tín hiệu 3: Đối chiếu toán học tỷ lệ thuế và tổng cộng
    if tt80_math_hits > tt103_math_hits:
        return "TT80"
    if tt103_math_hits > tt80_math_hits:
        return "TT103"

    # Tín hiệu 4: Đặc trưng giá trị ct6 vs ct7 (nếu ct6 > 100 còn ct7 <= 100 thì ct7 là tỷ lệ % của TT103)
    if ct6_num_max > 100 and 0 < ct7_num_max <= 100:
        return "TT103"

    # Tín hiệu 5: Dự phòng theo năm tính thuế (TT80/2021 áp dụng từ kỳ 01/2022)
    if year is not None and year <= 2021:
        return "TT103"
    return "TT80"


def extract_ntnn_totals_from_xml_root(
    root: ET.Element,
    ma_tkhai: str = "",
    mo_ta_bmau: str = "",
    ten_tkhai: str = "",
    year: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Trích xuất chính xác các chỉ tiêu tổng cộng của tờ khai Thuế Nhà Thầu (01/NTNN, 02/NTNN, 03/NTNN, 04/NTNN)
    từ cây XML gốc, phân biệt rạch ròi chuẩn mẫu TT103/2014 và TT80/2021:
      - dt_gtgt: Doanh thu tính thuế GTGT
      - thue_gtgt: Thuế GTGT phải nộp
      - dt_tndn: Doanh thu tính thuế TNDN (Doanh thu chưa bao gồm thuế GTGT)
      - thue_tndn: Thuế TNDN phải nộp
      - tong_thue: Tổng cộng tiền thuế nhà thầu phải nộp
      - schema: 'Semantic' | 'TT80' | 'TT103'
    """
    scan_root = _get_main_declaration_root(root)
    total_map: Dict[str, Any] = {
        "dt_gtgt": 0.0,
        "thue_gtgt": 0.0,
        "dt_tndn": 0.0,
        "thue_tndn": 0.0,
        "tong_thue": 0.0,
        "schema": "",
    }

    # --- ƯU TIÊN 1: Thẻ có tên ngữ nghĩa rõ ràng (không bao giờ lệch giữa các thông tư) ---
    sem_tot_dt_g = {"tongdoanhthugtgt", "tongdthutinhthuegtgt", "tongdthugtgt", "tong_dthu_gtgt"}
    sem_tot_th_g = {"tongthuegtgt", "tongthuegtgtpnop", "tong_thue_gtgt", "tongthuegtgtphainop", "tongsothuegtgt"}
    sem_tot_dt_t = {"tongdoanhthutndn", "tongdthutinhthuetndn", "tongdthutndn", "tong_dthu_tndn"}
    sem_tot_th_t = {"tongthuetndn", "tongthuetndnpnop", "tong_thue_tndn", "tongthuetndnphainop", "tongsothuetndn"}
    sem_tot_all = {"tongthuepnop", "tongthuephainop", "tongsothuepnop", "tongcong", "tong_thue_pnop"}

    for el in _iter_non_khbs(scan_root):
        if len(el) > 0 or not el.text or not el.text.strip():
            continue
        t = _strip_ns(el.tag).lower()
        v = _to_float(el.text)
        if v == 0:
            continue
        if t in sem_tot_dt_g and total_map["dt_gtgt"] == 0:
            total_map["dt_gtgt"] = v
        elif t in sem_tot_th_g and total_map["thue_gtgt"] == 0:
            total_map["thue_gtgt"] = v
        elif t in sem_tot_dt_t and total_map["dt_tndn"] == 0:
            total_map["dt_tndn"] = v
        elif t in sem_tot_th_t and total_map["thue_tndn"] == 0:
            total_map["thue_tndn"] = v
        elif t in sem_tot_all and total_map["tong_thue"] == 0:
            total_map["tong_thue"] = v

    if any(total_map[k] != 0 for k in ("dt_gtgt", "thue_gtgt", "dt_tndn", "thue_tndn")):
        total_map["schema"] = "Semantic"
        if total_map["tong_thue"] == 0:
            total_map["tong_thue"] = total_map["thue_gtgt"] + total_map["thue_tndn"]
        return total_map

    # Kiểm tra dòng chi tiết có thẻ ngữ nghĩa rõ ràng
    sem_row_dt_g = {"dthutinhthuegtgt", "doanhthutinhthuegtgt", "doanhthugtgt"}
    sem_row_th_g = {"thuegtgtpnop", "thuegtgtphainop", "sothuegtgt", "thuegtgt"}
    sem_row_dt_t = {"dthutinhthuetndn", "doanhthutinhthuetndn", "doanhthutndn"}
    sem_row_th_t = {"thuetndnpnop", "thuetndnphainop", "sothuetndn", "thuetndn"}
    sem_row_all = {"tongthuepnop", "tongthuephainop", "tongsothuepnop"}

    for el in _iter_non_khbs(scan_root):
        children = list(el)
        if not children:
            continue
        child_vals = {
            _strip_ns(ch.tag).lower(): _to_float(ch.text)
            for ch in children
            if len(ch) == 0 and ch.text and ch.text.strip()
        }
        if not child_vals:
            continue
        if any(k in child_vals for k in sem_row_dt_g | sem_row_th_g | sem_row_dt_t | sem_row_th_t):
            for k in sem_row_dt_g:
                if k in child_vals:
                    total_map["dt_gtgt"] += child_vals[k]
                    break
            for k in sem_row_th_g:
                if k in child_vals:
                    total_map["thue_gtgt"] += child_vals[k]
                    break
            for k in sem_row_dt_t:
                if k in child_vals:
                    total_map["dt_tndn"] += child_vals[k]
                    break
            for k in sem_row_th_t:
                if k in child_vals:
                    total_map["thue_tndn"] += child_vals[k]
                    break
            for k in sem_row_all:
                if k in child_vals:
                    total_map["tong_thue"] += child_vals[k]
                    break

    if any(total_map[k] != 0 for k in ("dt_gtgt", "thue_gtgt", "dt_tndn", "thue_tndn")):
        total_map["schema"] = "Semantic"
        if total_map["tong_thue"] == 0:
            total_map["tong_thue"] = total_map["thue_gtgt"] + total_map["thue_tndn"]
        return total_map

    # --- ƯU TIÊN 2: Thẻ đánh số ctN / tong_ctN theo ĐÚNG chuẩn Thông tư (TT80 vs TT103) ---
    schema = detect_ntnn_schema(root, ma_tkhai=ma_tkhai, mo_ta_bmau=mo_ta_bmau, ten_tkhai=ten_tkhai, year=year)
    total_map["schema"] = schema

    if schema == "TT103":
        tot_dt_g = {"tong_ct6", "tong_ct06", "ct6_tong", "ct06_tong"}
        tot_th_g = {"tong_ct8", "tong_ct08", "ct8_tong", "ct08_tong"}
        tot_dt_t = {"tong_ct9", "tong_ct09", "ct9_tong", "ct09_tong"}
        tot_th_t = {"tong_ct12", "ct12_tong"}
        tot_all = {"tong_ct13", "ct13_tong"}
        row_dt_g = {"ct6", "ct06", "ct_6"}
        row_th_g = {"ct8", "ct08", "ct_8"}
        row_dt_t = {"ct9", "ct09", "ct_9"}
        row_th_t = {"ct12", "ct_12"}
        row_all = {"ct13", "ct_13"}
    else:  # TT80/2021
        tot_dt_g = {"tong_ct7", "tong_ct07", "ct7_tong", "ct07_tong"}
        tot_th_g = {"tong_ct9", "tong_ct09", "ct9_tong", "ct09_tong"}
        tot_dt_t = {"tong_ct10", "ct10_tong"}
        tot_th_t = {"tong_ct13", "ct13_tong"}
        tot_all = {"tong_ct14", "ct14_tong"}
        row_dt_g = {"ct7", "ct07", "ct_7"}
        row_th_g = {"ct9", "ct09", "ct_9"}
        row_dt_t = {"ct10", "ct_10"}
        row_th_t = {"ct13", "ct_13"}
        row_all = {"ct14", "ct_14"}

    # Đọc thẻ tổng cộng của bảng trước
    for el in _iter_non_khbs(scan_root):
        if len(el) > 0 or not el.text or not el.text.strip():
            continue
        t = _strip_ns(el.tag).lower()
        v = _to_float(el.text)
        if v == 0:
            continue
        if t in tot_dt_g and total_map["dt_gtgt"] == 0:
            total_map["dt_gtgt"] = v
        elif t in tot_th_g and total_map["thue_gtgt"] == 0:
            total_map["thue_gtgt"] = v
        elif t in tot_dt_t and total_map["dt_tndn"] == 0:
            total_map["dt_tndn"] = v
        elif t in tot_th_t and total_map["thue_tndn"] == 0:
            total_map["thue_tndn"] = v
        elif t in tot_all and total_map["tong_thue"] == 0:
            total_map["tong_thue"] = v

    # Nếu không có thẻ tong_ct* thì cộng dồn từ các dòng chi tiết theo đúng chuẩn thông tư
    if not any(total_map[k] != 0 for k in ("dt_gtgt", "thue_gtgt", "dt_tndn", "thue_tndn")):
        for el in _iter_non_khbs(scan_root):
            children = list(el)
            if not children:
                continue
            child_vals = {
                _strip_ns(ch.tag).lower(): _to_float(ch.text)
                for ch in children
                if len(ch) == 0 and ch.text and ch.text.strip()
            }
            if not child_vals:
                continue
            if any(k in child_vals for k in row_dt_g | row_th_g | row_dt_t | row_th_t):
                for k in row_dt_g:
                    if k in child_vals:
                        total_map["dt_gtgt"] += child_vals[k]
                        break
                for k in row_th_g:
                    if k in child_vals:
                        total_map["thue_gtgt"] += child_vals[k]
                        break
                for k in row_dt_t:
                    if k in child_vals:
                        total_map["dt_tndn"] += child_vals[k]
                        break
                for k in row_th_t:
                    if k in child_vals:
                        total_map["thue_tndn"] += child_vals[k]
                        break
                for k in row_all:
                    if k in child_vals:
                        total_map["tong_thue"] += child_vals[k]
                        break

    if total_map["tong_thue"] == 0:
        total_map["tong_thue"] = total_map["thue_gtgt"] + total_map["thue_tndn"]

    return total_map


def get_ntnn_metrics(d: TaxDeclaration) -> Dict[str, Any]:
    """Lấy các chỉ tiêu tài chính của 1 tờ khai NTNN (ưu tiên từ indicators đã trích xuất từ XML gốc)."""
    ind = d.indicators or {}
    dt_g = _to_float(ind.get("_ntnn_dt_gtgt"))
    th_g = _to_float(ind.get("_ntnn_thue_gtgt"))
    dt_t = _to_float(ind.get("_ntnn_dt_tndn"))
    th_t = _to_float(ind.get("_ntnn_thue_tndn"))
    tot = _to_float(ind.get("_ntnn_tong_thue"))
    schema = str(ind.get("_ntnn_schema") or "")

    if dt_g == 0 and th_g == 0 and dt_t == 0 and th_t == 0:
        if d.file_path and d.file_path.lower().endswith(".xml"):
            try:
                tree = ET.parse(d.file_path)
                m = extract_ntnn_totals_from_xml_root(
                    tree.getroot(),
                    ma_tkhai=d.ma_tkhai,
                    mo_ta_bmau=d.mo_ta_bmau,
                    ten_tkhai=d.ten_tkhai,
                    year=d.year,
                )
                dt_g = float(m["dt_gtgt"])
                th_g = float(m["thue_gtgt"])
                dt_t = float(m["dt_tndn"])
                th_t = float(m["thue_tndn"])
                tot = float(m["tong_thue"])
                schema = str(m.get("schema") or "")
            except Exception:
                pass

    # Doanh thu chưa bao gồm thuế GTGT (chuẩn lên dòng C95 biểu từng năm):
    # Ưu tiên Doanh thu tính thuế TNDN (chưa gồm GTGT); chỉ khi tờ khai không có TNDN mới lấy (DT GTGT - Thuế GTGT).
    dt_chua_gtgt = dt_t if dt_t > 0 else max(0.0, dt_g - th_g)
    if tot == 0:
        tot = th_g + th_t
    return {
        "dt_gtgt": dt_g,
        "thue_gtgt": th_g,
        "dt_tndn": dt_t,
        "thue_tndn": th_t,
        "dt_chua_gtgt": dt_chua_gtgt,
        "tong_thue": tot,
        "schema": schema or ("TT103" if (d.year or 2022) <= 2021 else "TT80"),
    }


def _sort_key_ntnn(d: TaxDeclaration) -> Tuple[int, datetime.date, str, int, datetime.date, str]:
    yr = d.year or (d.tu_ngay.year if d.tu_ngay else (d.ngay_lap.year if d.ngay_lap else 0))
    tn = d.tu_ngay or datetime.date.min
    try:
        lan = int(d.so_lan) if d.so_lan not in (None, "") else 0
    except (ValueError, TypeError):
        lan = 0
    nl = d.ngay_lap or datetime.date.min
    return (yr, tn, d.ky_kkhai or "", lan, nl, d.file_name or "")


def filter_latest_ntnn_declarations(declarations: List[TaxDeclaration]) -> List[TaxDeclaration]:
    """
    Chốt tờ khai NTNN hiệu lực cuối cùng cho từng kỳ:
    - Với kê khai theo tháng/quý: mỗi (MST, kỳ kê khai) giữ bản có so_lan lớn nhất (cùng lần giữ ngay_lap mới nhất).
    - Với kê khai theo từng lần phát sinh (có tu_ngay/den_ngay cụ thể): nhóm theo (MST, kỳ kê khai, tu_ngay, den_ngay)
      để giữ đầy đủ tất cả các lần phát sinh khác nhau trong năm và chỉ thay thế bản cũ khi có bản Bổ sung của đúng lần đó.
    """
    grouped: Dict[Tuple[str, str, Optional[datetime.date], Optional[datetime.date]], List[TaxDeclaration]] = {}
    for d in declarations:
        key = (d.mst or "", d.ky_kkhai or "", d.tu_ngay, d.den_ngay)
        grouped.setdefault(key, []).append(d)

    latest: List[TaxDeclaration] = []
    for _, dlist in grouped.items():
        dlist.sort(key=_sort_key_ntnn)
        latest.append(dlist[-1])

    latest.sort(key=_sort_key_ntnn)
    return latest


NTNN_HEADERS = [
    "STT",
    "Tên file XML",
    "MST nộp thay",
    "Tên đơn vị",
    "Mã TK",
    "Kỳ kê khai",
    "Từ ngày",
    "Đến ngày",
    "Loại TK",
    "Lần BS",
    "Ngày lập",
    "Chuẩn mẫu",
    "Doanh thu tính thuế GTGT",
    "Thuế GTGT phải nộp",
    "Doanh thu tính thuế TNDN\n(chưa gồm GTGT)",
    "Thuế TNDN phải nộp",
    "Doanh thu chưa gồm GTGT\n(Lên biểu năm)",
    "Tổng thuế NTNN\nphải nộp",
]
# Cột số liệu:
# M(13)=DT GTGT, N(14)=Thuế GTGT, O(15)=DT TNDN (chưa gồm GTGT), P(16)=Thuế TNDN,
# Q(17)=DT chưa gồm GTGT (lên C95), R(18)=Tổng thuế NTNN phải nộp (=N+P)


def _style_ntnn_header(ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 30
    for ci, h in enumerate(NTNN_HEADERS, start=1):
        c = ws.cell(1, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    widths = [6, 28, 15, 28, 10, 12, 12, 12, 11, 8, 12, 11, 18, 17, 19, 17, 19, 18]
    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "C2"


def _write_ntnn_row(ws: openpyxl.worksheet.worksheet.Worksheet, r: int, stt: int, d: TaxDeclaration) -> None:
    m = get_ntnn_metrics(d)
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center")
    al_l = Alignment(horizontal="left", vertical="center")
    al_r = Alignment(horizontal="right", vertical="center")

    lan_int = int(d.so_lan) if str(d.so_lan or "").isdigit() else 0
    loai_txt = "Bổ sung" if lan_int > 0 else "Chính thức"

    vals = [
        stt,
        d.file_name,
        d.mst,
        d.ten_nnt,
        d.ma_tkhai,
        d.ky_kkhai,
        d.tu_ngay,
        d.den_ngay,
        loai_txt,
        lan_int,
        d.ngay_lap,
        m["schema"],
        round(m["dt_gtgt"]),
        round(m["thue_gtgt"]),
        round(m["dt_tndn"]),
        round(m["thue_tndn"]),
        f"=IF(O{r}>0,O{r},MAX(0,M{r}-N{r}))",
        f"=N{r}+P{r}",
    ]
    for ci, v in enumerate(vals, start=1):
        c = ws.cell(r, ci, v)
        c.font = f_bold if ci == 18 else f_reg
        c.border = b_thin
        if ci == 3:
            c.number_format = "@"
            c.alignment = al_l
        elif ci in (7, 8, 11):
            c.number_format = "dd/mm/yyyy"
            c.alignment = al_c
        elif ci in (1, 5, 6, 9, 10, 12):
            c.alignment = al_c
        elif ci >= 13:
            c.number_format = "#,##0"
            c.alignment = al_r
        else:
            c.alignment = al_l


def write_ntnn_timeline_sheet(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    sheet_name: str = "So Thue NTNN TH",
) -> None:
    """Sheet 1: Tổng hợp TẤT CẢ tờ khai Thuế Nhà Thầu (chính thức + bổ sung) theo dòng thời gian."""
    if not declarations:
        return
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    existing_sheets.add(sheet_name)
    ws = wb.create_sheet(title=sheet_name)
    _style_ntnn_header(ws)

    sorted_decls = sorted(declarations, key=_sort_key_ntnn)
    for idx, d in enumerate(sorted_decls, start=1):
        _write_ntnn_row(ws, idx + 1, idx, d)


def write_ntnn_cuoi_sheet(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    sheet_name: str = "So Thue NTNN Cuoi",
) -> Dict[int, int]:
    """
    Sheet 2: Chốt tờ khai NTNN hiệu lực cuối cùng theo từng kỳ/lần phát sinh, nhóm theo từng năm
    và có dòng 'Tổng cộng năm YYYY' (=SUM(...)) ngay bên dưới mỗi năm.
    Trả về dict {year: row_index_of_year_total}.
    """
    if not declarations:
        return {}
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    existing_sheets.add(sheet_name)
    ws = wb.create_sheet(title=sheet_name)
    _style_ntnn_header(ws)

    latest = filter_latest_ntnn_declarations(declarations)
    by_year: Dict[int, List[TaxDeclaration]] = {}
    for d in latest:
        yr = d.year or (d.tu_ngay.year if d.tu_ngay else (d.ngay_lap.year if d.ngay_lap else 0))
        by_year.setdefault(yr, []).append(d)

    f_bold = Font(name="Calibri", size=9, bold=True)
    fill_yr = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_tot = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center")
    al_r = Alignment(horizontal="right", vertical="center")

    cur_r = 2
    stt = 0
    year_sum_rows: Dict[int, int] = {}

    for yr in sorted(by_year.keys()):
        r_start = cur_r
        for d in by_year[yr]:
            stt += 1
            _write_ntnn_row(ws, cur_r, stt, d)
            cur_r += 1
        r_end = cur_r - 1

        # Dòng Tổng cộng năm YYYY
        r_sum = cur_r
        year_sum_rows[yr] = r_sum
        ws.cell(r_sum, 6, f"Tổng cộng năm {yr}").alignment = al_c
        for ci in range(13, 19):
            col_l = get_column_letter(ci)
            c = ws.cell(r_sum, ci, f"=SUM({col_l}{r_start}:{col_l}{r_end})")
            c.number_format = "#,##0"
            c.alignment = al_r
        for ci in range(1, 19):
            c = ws.cell(r_sum, ci)
            c.font = f_bold
            c.fill = fill_yr
            c.border = b_thin
        cur_r += 1

    # Dòng TỔNG CỘNG TẤT CẢ CÁC NĂM
    if year_sum_rows:
        r_grand = cur_r
        ws.cell(r_grand, 6, "TỔNG CỘNG TẤT CẢ CÁC NĂM").alignment = al_c
        for ci in range(13, 19):
            col_l = get_column_letter(ci)
            c = ws.cell(r_grand, ci, "=" + "+".join(f"{col_l}{rr}" for rr in year_sum_rows.values()))
            c.number_format = "#,##0"
            c.alignment = al_r
        for ci in range(1, 19):
            c = ws.cell(r_grand, ci)
            c.font = f_bold
            c.fill = fill_tot
            c.border = b_thin

    return year_sum_rows
