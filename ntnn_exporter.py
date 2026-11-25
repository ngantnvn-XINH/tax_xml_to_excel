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


def extract_ntnn_contract_rows_from_xml_root(
    root: ET.Element,
    ma_tkhai: str = "",
    mo_ta_bmau: str = "",
    ten_tkhai: str = "",
    year: Optional[int] = None,
) -> List[Dict[str, Any]]:
    """
    Bóc tách CHI TIẾT TỪNG DÒNG HỢP ĐỒNG / NHÀ THẦU trên tờ khai Thuế Nhà Thầu (01/NTNN...)
    theo đúng chuẩn mẫu từng thời kỳ (TT103/2014 vs TT80/2021 vs thẻ ngữ nghĩa) để đưa lên sheet 'So Thue NTNN TH'.
    """
    scan_root = _get_main_declaration_root(root)
    schema = detect_ntnn_schema(root, ma_tkhai=ma_tkhai, mo_ta_bmau=mo_ta_bmau, ten_tkhai=ten_tkhai, year=year)

    def _pick_str(cmap: Dict[str, str], keys: Tuple[str, ...]) -> str:
        for k in keys:
            if k in cmap and cmap[k]:
                return cmap[k]
        return ""

    def _pick_num(cmap: Dict[str, str], keys: Tuple[str, ...]) -> float:
        for k in keys:
            if k in cmap and cmap[k]:
                return _to_float(cmap[k])
        return 0.0

    # Khóa nhận diện dòng chi tiết hợp đồng (phân biệt rạch ròi TT103 vs TT80)
    sem_fin_keys = {
        "dthutinhthuegtgt", "doanhthutinhthuegtgt", "doanhthugtgt",
        "thuegtgtpnop", "thuegtgtphainop", "sothuegtgt",
        "dthutinhthuetndn", "doanhthutinhthuetndn", "doanhthutndn",
        "thuetndnpnop", "thuetndnphainop", "sothuetndn",
    }
    tt80_fin_keys = {"ct7", "ct07", "ct_7", "ct9", "ct09", "ct_9", "ct10", "ct_10", "ct13", "ct_13", "ct14", "ct_14"}
    tt103_fin_keys = {"ct6", "ct06", "ct_6", "ct8", "ct08", "ct_8", "ct9", "ct09", "ct_9", "ct12", "ct_12", "ct13", "ct_13"}

    rows: List[Dict[str, Any]] = []
    for el in _iter_non_khbs(scan_root):
        children = list(el)
        if not children:
            continue
        # Chỉ lấy node dòng lá (chứa các thẻ chỉ tiêu trực tiếp, không phải node cha bao toàn tờ khai)
        cmap: Dict[str, str] = {}
        for ch in children:
            if len(ch) == 0 and ch.text and ch.text.strip():
                cmap[_strip_ns(ch.tag).lower()] = ch.text.strip()
        if not cmap:
            continue
        # Bỏ qua node dòng tổng cộng (tong_ct*)
        if any(k.startswith("tong_") or k.endswith("_tong") for k in cmap.keys()):
            continue

        is_sem = any(k in cmap for k in sem_fin_keys)
        is_num_row = any(k in cmap for k in (tt103_fin_keys if schema == "TT103" else tt80_fin_keys))
        if not is_sem and not is_num_row:
            continue

        noi_dung = _pick_str(cmap, ("noidung", "tennhathau", "tenhhdv", "ct2", "ct02", "ct_2", "ct1a", "ct1b", "ct1"))
        mst_nt = _pick_str(cmap, ("mstnhathau", "mstntnn", "mst_ntnn", "ct3", "ct03", "ct_3"))
        so_hd = _pick_str(cmap, ("sohopdong", "sohd", "so_hd", "ct4", "ct04", "ct_4"))
        ngay_hd = _pick_str(cmap, ("ngayhopdong", "ngayhd", "ngay_hd", "ct5", "ct05", "ct_5"))

        if is_sem:
            ngay_tt = _pick_str(cmap, ("ngaythanhtoan", "ngaytt", "ngay_tt", "ct6", "ct06"))
            dt_g = _pick_num(cmap, ("dthutinhthuegtgt", "doanhthutinhthuegtgt", "doanhthugtgt"))
            tl_g = _pick_num(cmap, ("tylegtgt", "tyletinhthuegtgt", "thuesuatgtgt"))
            th_g = _pick_num(cmap, ("thuegtgtpnop", "thuegtgtphainop", "sothuegtgt", "thuegtgt"))
            dt_t = _pick_num(cmap, ("dthutinhthuetndn", "doanhthutinhthuetndn", "doanhthutndn"))
            tl_t = _pick_num(cmap, ("tyletndn", "tyletinhthuetndn", "thuesuattndn"))
            mg_t = _pick_num(cmap, ("thuemiengiam", "sothuemiengiam", "miengiam"))
            th_t = _pick_num(cmap, ("thuetndnpnop", "thuetndnphainop", "sothuetndn", "thuetndn"))
            tot = _pick_num(cmap, ("tongthuepnop", "tongthuephainop", "tongsothuepnop"))
            row_schema = "Semantic"
        elif schema == "TT103":
            ngay_tt = ""
            dt_g = _pick_num(cmap, ("ct6", "ct06", "ct_6"))
            tl_g = _pick_num(cmap, ("ct7", "ct07", "ct_7"))
            th_g = _pick_num(cmap, ("ct8", "ct08", "ct_8"))
            dt_t = _pick_num(cmap, ("ct9", "ct09", "ct_9"))
            tl_t = _pick_num(cmap, ("ct10", "ct_10"))
            mg_t = _pick_num(cmap, ("ct11", "ct_11"))
            th_t = _pick_num(cmap, ("ct12", "ct_12"))
            tot = _pick_num(cmap, ("ct13", "ct_13"))
            row_schema = "TT103"
        else:  # TT80
            ngay_tt = _pick_str(cmap, ("ct6", "ct06", "ct_6", "ngaythanhtoan"))
            dt_g = _pick_num(cmap, ("ct7", "ct07", "ct_7"))
            tl_g = _pick_num(cmap, ("ct8", "ct08", "ct_8"))
            th_g = _pick_num(cmap, ("ct9", "ct09", "ct_9"))
            dt_t = _pick_num(cmap, ("ct10", "ct_10"))
            tl_t = _pick_num(cmap, ("ct11", "ct_11"))
            mg_t = _pick_num(cmap, ("ct12", "ct_12"))
            th_t = _pick_num(cmap, ("ct13", "ct_13"))
            tot = _pick_num(cmap, ("ct14", "ct_14"))
            row_schema = "TT80"

        if dt_g == 0 and th_g == 0 and dt_t == 0 and th_t == 0 and not noi_dung and not so_hd:
            continue
        if tot == 0:
            tot = th_g + th_t
        rows.append({
            "noi_dung": noi_dung,
            "mst_nt": mst_nt,
            "so_hd": so_hd,
            "ngay_hd": ngay_hd,
            "ngay_tt": ngay_tt,
            "dt_gtgt": dt_g,
            "ty_le_gtgt": tl_g,
            "thue_gtgt": th_g,
            "dt_tndn": dt_t,
            "ty_le_tndn": tl_t,
            "mien_giam": mg_t,
            "thue_tndn": th_t,
            "tong_thue": tot,
            "schema": row_schema,
        })

    return rows


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


def is_ntnn_per_occurrence(d: TaxDeclaration) -> bool:
    """
    Phân biệt tờ khai Thuế Nhà Thầu kê khai theo LẦN PHÁT SINH (THEO NGÀY) với kê khai theo THÁNG/QUÝ/NĂM:
    - Trả về True (lần phát sinh -> GIỮ & CỘNG TẤT CẢ Chính thức + Bổ sung) nếu:
        * Thẻ XML <kieuKy> thuộc ('D', 'P', 'L', 'PS')
        * Hoặc ky_kkhai chứa ngày đầy đủ (VD: '15/06/2024', 'D/15/06/2024', 'P/15/06/2024', '2024-06-15')
        * Hoặc khoảng thời gian tu_ngay..den_ngay ngắn hơn 20 ngày (VD tu_ngay == den_ngay, không phải kỳ tháng/quý/năm)
        * Hoặc chỉ có tu_ngay mà ngày bắt đầu != ngày 1 đầu tháng
    - Trả về False nếu kê khai theo tháng ('06/2024', tu_ngay=01..30), quý ('Q2/2024') hoặc năm ('2024').
    """
    kieu_ky = str((d.indicators or {}).get("_kieu_ky") or "").strip().upper()
    if kieu_ky in ("D", "P", "L", "PS"):
        return True

    ky = str(d.ky_kkhai or "").strip()
    if re.search(r"(\d{1,2}[/-]\d{1,2}[/-]\d{4}|\d{4}[/-]\d{1,2}[/-]\d{1,2})", ky):
        return True

    if d.tu_ngay and d.den_ngay:
        span_days = (d.den_ngay - d.tu_ngay).days
        if 0 <= span_days < 20:
            return True

    if d.tu_ngay and not d.den_ngay and d.tu_ngay.day != 1:
        return True

    if kieu_ky in ("M", "Q", "Y"):
        return False

    return False


def _ntnn_content_fingerprint(d: TaxDeclaration) -> Tuple[Any, ...]:
    """
    Khóa nhận diện 2 tờ khai NTNN trùng y như nhau (bất kể nằm ở thư mục MST công ty chính
    hay thư mục MST nhà thầu riêng, bất kể tên file khác nhau):
    cùng kỳ kê khai + cùng số Doanh thu GTGT, Thuế GTGT, Doanh thu TNDN, Thuế TNDN.
    """
    m = get_ntnn_metrics(d)
    ky_norm = (d.ky_kkhai or "").strip().upper() or (d.tu_ngay.strftime("%d/%m/%Y") if d.tu_ngay else "")
    return (
        ky_norm,
        round(m["dt_gtgt"]),
        round(m["thue_gtgt"]),
        round(m["dt_tndn"]),
        round(m["thue_tndn"]),
    )


def find_ntnn_duplicates(declarations: List[TaxDeclaration]) -> Dict[int, str]:
    """
    Phát hiện các tờ khai NTNN cùng kỳ và trùng y như nhau về số liệu (kể cả trùng chéo giữa
    thư mục MST công ty chính và thư mục MST nhà thầu riêng).
    Trả về dict {id(d): thông_báo_trùng} cho các bản bị trùng (bản đầu tiên được giữ là Hợp lệ).
    """
    sorted_decls = sorted(declarations, key=_sort_key_ntnn)
    seen: Dict[Tuple[Any, ...], TaxDeclaration] = {}
    dup_notes: Dict[int, str] = {}
    for d in sorted_decls:
        fp = _ntnn_content_fingerprint(d)
        # Chỉ báo trùng chéo khi có số liệu thực (tránh báo trùng các tờ khai trắng 0 đồng khác kỳ)
        if fp in seen:
            orig = seen[fp]
            mst_info = f"MST {orig.mst}" if orig.mst else ""
            dup_notes[id(d)] = f"TRÙNG TỜ KHAI với {orig.file_name} ({mst_info}) - Đã loại khỏi NTNN Cuoi"
        else:
            seen[fp] = d
    return dup_notes


def filter_latest_ntnn_declarations(declarations: List[TaxDeclaration]) -> List[TaxDeclaration]:
    """
    Chốt danh sách tờ khai NTNN đưa vào 'So Thue NTNN Cuoi' từ CẢ 2 thư mục (MST công ty chính + MST nhà thầu):
    0. Loại bỏ các tờ khai cùng kỳ trùng y như nhau về số liệu (kể cả trùng chéo giữa 2 MST).
    1. Tờ khai NTNN kê khai THEO THÁNG / QUÝ / NĂM:
       - Mỗi (MST, kỳ kê khai) chỉ lấy 01 tờ khai hiệu lực CUỐI CÙNG (so_lan lớn nhất; cùng lần giữ ngay_lap mới nhất).
    2. Tờ khai NTNN kê khai THEO LẦN PHÁT SINH (THEO NGÀY):
       - GIỮ VÀ CỘNG TẤT CẢ CẢ TỜ KHAI CHÍNH THỨC VÀ TỜ KHAI BỔ SUNG (sau khi đã loại bản trùng y như nhau ở bước 0).
    """
    dup_notes = find_ntnn_duplicates(declarations)
    periodic_grouped: Dict[Tuple[str, str], List[TaxDeclaration]] = {}
    per_occurrence_list: List[TaxDeclaration] = []

    for d in declarations:
        if id(d) in dup_notes:
            continue
        if is_ntnn_per_occurrence(d):
            per_occurrence_list.append(d)
        else:
            ky_key = (d.ky_kkhai or "").strip().upper() or (d.tu_ngay.strftime("%m/%Y") if d.tu_ngay else "")
            periodic_grouped.setdefault((d.mst or "", ky_key), []).append(d)

    result: List[TaxDeclaration] = list(per_occurrence_list)
    for _, dlist in periodic_grouped.items():
        dlist.sort(key=_sort_key_ntnn)
        result.append(dlist[-1])

    result.sort(key=_sort_key_ntnn)
    return result


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
    ky_mode_txt = "Lần PS" if is_ntnn_per_occurrence(d) else "Tháng/Kỳ"
    schema_disp = f"{m['schema']} ({ky_mode_txt})"

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
        schema_disp,
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


NTNN_DETAIL_HEADERS = [
    "STT",
    "Tên file XML",
    "MST nộp thay",
    "Kỳ kê khai",
    "Loại TK",
    "Lần BS",
    "Ngày lập",
    "Chuẩn mẫu",
    "Nội dung / Tên nhà thầu [2]",
    "MST nhà thầu [3]",
    "Số hợp đồng [4]",
    "Ngày hợp đồng [5]",
    "Ngày thanh toán [6]",
    "Doanh thu tính thuế GTGT",
    "Tỷ lệ GTGT (%)",
    "Thuế GTGT phải nộp",
    "Doanh thu tính thuế TNDN\n(chưa gồm GTGT)",
    "Tỷ lệ TNDN (%)",
    "Miễn giảm Hiệp định",
    "Thuế TNDN phải nộp",
    "Doanh thu chưa gồm GTGT",
    "Tổng thuế NTNN\nphải nộp",
    "Trạng thái / Báo trùng",
]


def get_ntnn_contract_rows(d: TaxDeclaration) -> List[Dict[str, Any]]:
    """Lấy danh sách dòng chi tiết hợp đồng của 1 tờ khai NTNN (ưu tiên từ indicators['_ntnn_rows'], dự phòng đọc trực tiếp từ XML)."""
    ind = d.indicators or {}
    cached = ind.get("_ntnn_rows")
    if isinstance(cached, list) and cached:
        return cached
    if d.file_path and d.file_path.lower().endswith(".xml"):
        try:
            tree = ET.parse(d.file_path)
            parsed = extract_ntnn_contract_rows_from_xml_root(
                tree.getroot(),
                ma_tkhai=d.ma_tkhai,
                mo_ta_bmau=d.mo_ta_bmau,
                ten_tkhai=d.ten_tkhai,
                year=d.year,
            )
            if parsed:
                return parsed
        except Exception:
            pass
    # Dự phòng: nếu tờ khai không có dòng chi tiết hợp đồng -> trả về 1 dòng tổng hợp của tờ khai để không mất dữ liệu
    m = get_ntnn_metrics(d)
    return [{
        "noi_dung": d.ten_nnt or "(Tổng hợp tờ khai)",
        "mst_nt": "",
        "so_hd": "",
        "ngay_hd": "",
        "ngay_tt": "",
        "dt_gtgt": m["dt_gtgt"],
        "ty_le_gtgt": 0.0,
        "thue_gtgt": m["thue_gtgt"],
        "dt_tndn": m["dt_tndn"],
        "ty_le_tndn": 0.0,
        "mien_giam": 0.0,
        "thue_tndn": m["thue_tndn"],
        "tong_thue": m["tong_thue"],
        "schema": m["schema"],
    }]


def write_ntnn_timeline_sheet(
    wb: openpyxl.Workbook,
    declarations: List[TaxDeclaration],
    existing_sheets: Set[str],
    sheet_name: str = "So Thue NTNN TH",
) -> None:
    """
    Sheet 1: Tổng hợp CHI TIẾT TẤT CẢ CÁC HỢP ĐỒNG trên toàn bộ tờ khai Thuế Nhà Thầu
    (chính thức + bổ sung) theo dòng thời gian để kiểm tra đối chiếu từng hợp đồng.
    """
    if not declarations:
        return
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    existing_sheets.add(sheet_name)
    ws = wb.create_sheet(title=sheet_name)

    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_dup = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    f_dup = Font(name="Calibri", size=9, bold=True, color="9C0006")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")

    ws.row_dimensions[1].height = 30
    for ci, h in enumerate(NTNN_DETAIL_HEADERS, start=1):
        c = ws.cell(1, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    widths = [6, 26, 14, 12, 11, 7, 12, 16, 32, 14, 18, 12, 13, 18, 10, 17, 19, 10, 15, 17, 19, 18, 36]
    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "E2"

    dup_notes = find_ntnn_duplicates(declarations)
    sorted_decls = sorted(declarations, key=_sort_key_ntnn)
    cur_r = 2
    stt = 0
    for d in sorted_decls:
        lan_int = int(d.so_lan) if str(d.so_lan or "").isdigit() else 0
        loai_txt = "Bổ sung" if lan_int > 0 else "Chính thức"
        ky_mode_txt = "Lần PS" if is_ntnn_per_occurrence(d) else "Tháng/Kỳ"
        dup_msg = dup_notes.get(id(d), "")
        status_txt = dup_msg if dup_msg else "Hợp lệ"
        is_dup = bool(dup_msg)
        contract_rows = get_ntnn_contract_rows(d)
        for cr in contract_rows:
            stt += 1
            schema_disp = f"{cr.get('schema') or 'TT80'} ({ky_mode_txt})"
            vals = [
                stt,
                d.file_name,
                d.mst,
                d.ky_kkhai,
                loai_txt,
                lan_int,
                d.ngay_lap,
                schema_disp,
                cr.get("noi_dung", ""),
                cr.get("mst_nt", ""),
                cr.get("so_hd", ""),
                cr.get("ngay_hd", ""),
                cr.get("ngay_tt", ""),
                round(float(cr.get("dt_gtgt", 0.0))),
                float(cr.get("ty_le_gtgt", 0.0)),
                round(float(cr.get("thue_gtgt", 0.0))),
                round(float(cr.get("dt_tndn", 0.0))),
                float(cr.get("ty_le_tndn", 0.0)),
                round(float(cr.get("mien_giam", 0.0))),
                round(float(cr.get("thue_tndn", 0.0))),
                f"=IF(Q{cur_r}>0,Q{cur_r},MAX(0,N{cur_r}-P{cur_r}))",
                f"=P{cur_r}+T{cur_r}",
                status_txt,
            ]
            for ci, v in enumerate(vals, start=1):
                c = ws.cell(cur_r, ci, v)
                if is_dup:
                    c.font = f_dup if ci in (22, 23) else f_reg
                    c.fill = fill_dup
                else:
                    c.font = f_bold if ci == 22 else f_reg
                c.border = b_thin
                if ci in (3, 10, 11):
                    c.number_format = "@"
                    c.alignment = al_l
                elif ci == 7:
                    c.number_format = "dd/mm/yyyy"
                    c.alignment = al_c
                elif ci in (1, 4, 5, 6, 8, 12, 13):
                    c.alignment = al_c
                elif ci in (15, 18):
                    c.number_format = "0.##"
                    c.alignment = al_r
                elif 14 <= ci <= 22:
                    c.number_format = "#,##0"
                    c.alignment = al_r
                else:
                    c.alignment = al_l
            cur_r += 1

    # Dòng TỔNG CỘNG chi tiết hợp đồng
    if stt > 0:
        r_end = cur_r - 1
        ws.cell(cur_r, 9, "TỔNG CỘNG").alignment = al_c
        for ci in (14, 16, 17, 19, 20, 21, 22):
            col_l = get_column_letter(ci)
            c = ws.cell(cur_r, ci, f"=SUM({col_l}2:{col_l}{r_end})")
            c.number_format = "#,##0"
            c.alignment = al_r
        if dup_notes:
            ws.cell(cur_r, 23, f"Có {len(dup_notes)} tờ khai trùng (đã tự loại khỏi So Thue NTNN Cuoi)").alignment = al_l
        for ci in range(1, 24):
            c = ws.cell(cur_r, ci)
            c.font = f_bold
            c.fill = fill_tot
            c.border = b_thin
        ws.auto_filter.ref = f"A1:W{r_end}"


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
