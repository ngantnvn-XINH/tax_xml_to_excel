# -*- coding: utf-8 -*-
"""
Module đọc và phân tích (parse) file XML tờ khai thuế Việt Nam (HTKK / eTax) - Phiên bản v1.2.
Hỗ trợ đầy đủ các tờ khai GTGT, TNCN, TNDN, BCTC, Lệ phí môn bài / Khác và các phụ lục đi kèm.
Đảm bảo không thiếu dữ liệu đầu vào, không lệch dòng/cột và không lẫn chỉ tiêu giữa tờ khai chính và phụ lục.
"""

from __future__ import annotations

import datetime
import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple


# Từ điển tên chỉ tiêu phổ biến theo chuẩn Tổng cục Thuế (TT156 & TT80)
COMMON_INDICATOR_LABELS: Dict[str, Dict[str, str]] = {
    "GTGT": {
        "ct21": "[21] Không phát sinh HĐ mua bán trong kỳ",
        "ct22": "[22] Thuế GTGT còn được khấu trừ kỳ trước chuyển sang",
        "ct23": "[23] Giá trị HHDV mua vào",
        "ct23a": "[23a] Giá trị HHDV mua vào nhập khẩu",
        "ct24": "[24] Thuế GTGT của HHDV mua vào",
        "ct24a": "[24a] Thuế GTGT của HHDV nhập khẩu",
        "ct25": "[25] Tổng số thuế GTGT được khấu trừ kỳ này",
        "ct26": "[26] HHDV bán ra không chịu thuế GTGT",
        "ct27": "[27] Tổng doanh thu HHDV bán ra chịu thuế GTGT",
        "ct28": "[28] Tổng thuế GTGT của HHDV bán ra",
        "ct29": "[29] HHDV bán ra chịu thuế suất 0%",
        "ct30": "[30] Doanh thu HHDV bán ra chịu thuế suất 5%",
        "ct31": "[31] Thuế GTGT HHDV bán ra chịu thuế suất 5%",
        "ct32": "[32] Doanh thu HHDV bán ra chịu thuế suất 10%",
        "ct32a": "[32a] HHDV bán ra không phải kê khai, tính nộp thuế",
        "ct33": "[33] Thuế GTGT HHDV bán ra chịu thuế suất 10%",
        "ct34": "[34] Tổng doanh thu HHDV bán ra",
        "ct35": "[35] Tổng thuế GTGT HHDV bán ra",
        "ct36": "[36] Thuế GTGT phát sinh trong kỳ",
        "ct37": "[37] Điều chỉnh giảm thuế GTGT còn được khấu trừ kỳ trước",
        "ct38": "[38] Điều chỉnh tăng thuế GTGT còn được khấu trừ kỳ trước",
        "ct39a": "[39a] Thuế GTGT nhận bàn giao được khấu trừ trong kỳ",
        "ct40": "[40] Thuế GTGT phải nộp trong kỳ",
        "ct40a": "[40a] Thuế GTGT phải nộp của HĐSXKD trong kỳ",
        "ct40b": "[40b] Thuế GTGT mua vào của dự án đầu tư bù trừ",
        "ct41": "[41] Thuế GTGT chưa khấu trừ hết kỳ này",
        "ct42": "[42] Tổng số thuế GTGT đề nghị hoàn",
        "ct43": "[43] Thuế GTGT còn được khấu trừ chuyển kỳ sau",
    },
    "TNCN": {
        "ct21": "[21] Tổng số người lao động",
        "ct22": "[22] Cá nhân cư trú có HĐLĐ",
        "ct23": "[23] Tổng số cá nhân đã khấu trừ thuế",
        "ct24": "[24] Cá nhân cư trú đã khấu trừ thuế",
        "ct25": "[25] Cá nhân không cư trú đã khấu trừ thuế",
        "ct26": "[26] Tổng TNCT trả cho cá nhân",
        "ct27": "[27] Tổng TNCT trả cho cá nhân cư trú",
        "ct28": "[28] Tổng TNCT trả cho cá nhân không cư trú",
        "ct29": "[29] Tổng TNCT trả cho cá nhân thuộc diện khấu trừ",
        "ct30": "[30] Tổng TNCT trả cho cá nhân cư trú thuộc diện khấu trừ",
        "ct31": "[31] Tổng TNCT trả cho cá nhân không cư trú thuộc diện khấu trừ",
        "ct32": "[32] Tổng số thuế TNCN đã khấu trừ",
        "ct33": "[33] Thuế TNCN đã khấu trừ của cá nhân cư trú",
        "ct34": "[34] Thuế TNCN đã khấu trừ của cá nhân không cư trú",
        "ct35": "[35] Tổng TNCT từ tiền phí mua bảo hiểm nhân thọ",
        "ct36": "[36] Tổng số thuế TNCN đã khấu trừ từ tiền phí bảo hiểm",
        "ct40": "[40] Tổng số thuế TNCN phải nộp",
        "ct41": "[41] Tổng số thuế TNCN đã nộp thừa",
        "ct42": "[42] Tổng số thuế TNCN còn phải nộp",
        "ct43": "[43] Tổng số thuế TNCN hoàn trả",
        "ct44": "[44] Tổng số thuế TNCN bù trừ kỳ sau",
    },
    "TNDN": {
        "A1": "[A1] Tổng lợi nhuận kế toán trước thuế TNDN",
        "B1": "[B1] Điều chỉnh tăng tổng lợi nhuận trước thuế",
        "B8": "[B8] Điều chỉnh giảm tổng lợi nhuận trước thuế",
        "B12": "[B12] Tổng thu nhập chịu thuế",
        "B13": "[B13] Thu nhập miễn thuế",
        "B14": "[B14] Lỗ kết chuyển theo quy định",
        "C1": "[C1] Thu nhập tính thuế",
        "C4": "[C4] Thuế TNDN từ HĐSXKD tính theo thuế suất phổ thông",
        "C6": "[C6] Tổng số thuế TNDN phải nộp theo thuế suất phổ thông",
        "C10": "[C10] Thuế TNDN phải nộp trong kỳ",
        "D": "[D] Tổng số thuế TNDN phải nộp",
        "E": "[E] Tổng số thuế TNDN đã tạm nộp trong năm",
        "G": "[G] Chênh lệch giữa số thuế phải nộp và số đã tạm nộp",
        "H": "[H] Số thuế TNDN còn phải nộp sau quyết toán",
        "I": "[I] Số thuế TNDN nộp thừa",
    },
}


def strip_ns(tag: str) -> str:
    """Loại bỏ XML namespace khỏi tên thẻ."""
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def parse_date(val: Optional[str]) -> Optional[datetime.date]:
    """Chuyển chuỗi ngày trong XML (YYYY-MM-DD, DD/MM/YYYY, ...) sang datetime.date."""
    if not val:
        return None
    s = val.strip()
    if not s:
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y/%m/%d", "%d-%m-%Y", "%Y-%m-%dT%H:%M:%S"):
        try:
            return datetime.datetime.strptime(s[:19], fmt).date()
        except ValueError:
            continue
    return None


def parse_value(val: Optional[str]) -> Any:
    """
    Chuẩn hóa giá trị thẻ XML:
    - Trống -> None
    - Số nguyên / số thực -> int / float (chỉ khi không bắt đầu bằng '0' như MST/SĐT trừ '0' hoặc '0.xx')
    - Boolean ('true'/'false') -> giữ nguyên
    """
    if val is None:
        return None
    s = val.strip()
    if s == "":
        return None

    # Kiểm tra số nguyên
    if re.fullmatch(r"-?\d+", s):
        # Tránh biến chuỗi có số 0 đứng đầu (như mã/số hiệu/MST/CCCD) thành số, trừ số "0"
        if len(s) > 1 and s.startswith("0") and not s.startswith("-"):
            return s
        if len(s) > 2 and s.startswith("-0"):
            return s
        return int(s)

    # Kiểm tra số thập phân
    if re.fullmatch(r"-?\d+\.\d+", s):
        try:
            f = float(s)
            return int(f) if f.is_integer() else f
        except ValueError:
            return s

    return s


def normalize_mst(val: Optional[str]) -> str:
    """Chuẩn hóa mã số thuế: giữ chuỗi sạch, bảo toàn số 0 đứng đầu."""
    if not val:
        return ""
    s = re.sub(r"[\s\u00a0]", "", str(val))
    s = re.sub(r"[–—−‐]", "-", s)
    return s


def classify_tax_type(ma_tkhai: str, ten_tkhai: str) -> str:
    """Phân loại tờ khai thành GTGT, TNCN, TNDN, BCTC, NTNN hoặc Khác (ưu tiên tách riêng Nhà thầu NTNN trước GTGT/TNDN)."""
    combined = f"{ma_tkhai} {ten_tkhai}".upper()
    if (
        "NTNN" in combined
        or "NHÀ THẦU" in combined
        or "NHA THAU" in combined
    ):
        return "NTNN"
    if (
        "BCTC" in combined
        or "BÁO CÁO TÀI CHÍNH" in combined
        or "B01" in combined
        or "B02" in combined
        or "CÂN ĐỐI KẾ TOÁN" in combined
        or "TT133" in combined
        or "TT200" in combined
        or "QD48" in combined
        or "QĐ48" in combined
    ):
        return "BCTC"
    if "GTGT" in combined or "GIÁ TRỊ GIA TĂNG" in combined:
        return "GTGT"
    if "TNCN" in combined or "THU NHẬP CÁ NHÂN" in combined or "CNKD" in combined:
        return "TNCN"
    if "TNDN" in combined or "THU NHẬP DOANH NGHIỆP" in combined:
        return "TNDN"
    return "Khác"


def flatten_element(
    elem: ET.Element,
    prefix: str = "",
    out: Optional[Dict[str, Any]] = None,
    include_aliases: bool = True,
) -> Dict[str, Any]:
    """
    Làm phẳng các thẻ con trong một nhánh XML thành dict {key: value}.
    - Khi include_aliases=True (dùng cho bảng chỉ tiêu tờ khai chính): lưu thêm alias
      chữ hoa/chữ thường để tra cứu nhanh 100%.
    - Khi include_aliases=False (dùng cho dòng bảng phụ lục): chỉ lưu đúng key chuẩn,
      không tạo cột trùng lặp chữ hoa/chữ thường.
    """
    if out is None:
        out = {}

    children = list(elem)
    if not children:
        raw_tag = strip_ns(elem.tag)
        key = prefix or raw_tag
        val = parse_value(elem.text)
        if val is not None:
            out[key] = val
            is_khong_lt = "khonglientuc" in key.lower()
            if raw_tag not in out or (not is_khong_lt and out.get(raw_tag) in (None, 0, "0", "") and val not in (0, "0", "")):
                out[raw_tag] = val
            if include_aliases:
                if raw_tag.upper() not in out:
                    out[raw_tag.upper()] = val
                if raw_tag.lower() not in out:
                    out[raw_tag.lower()] = val
        return out

    # Đếm số lần xuất hiện của mỗi tag con có dữ liệu để xử lý trường hợp trùng tên
    non_empty_children: List[ET.Element] = []
    tag_counts: Dict[str, int] = {}
    for ch in children:
        if len(ch) == 0 and (ch.text is None or ch.text.strip() == ""):
            continue
        non_empty_children.append(ch)
        t = strip_ns(ch.tag)
        tag_counts[t] = tag_counts.get(t, 0) + 1

    tag_seen: Dict[str, int] = {}
    for ch in non_empty_children:
        t = strip_ns(ch.tag)
        if tag_counts[t] > 1:
            idx = tag_seen.get(t, 0) + 1
            tag_seen[t] = idx
            next_prefix = f"{prefix}.{t}[{idx}]" if prefix else f"{t}[{idx}]"
        else:
            next_prefix = f"{prefix}.{t}" if prefix else t
        flatten_element(ch, next_prefix, out, include_aliases=include_aliases)

    return out


def find_child(elem: Optional[ET.Element], tag_name: str) -> Optional[ET.Element]:
    """Tìm thẻ con trực tiếp theo tên (bỏ qua namespace)."""
    if elem is None:
        return None
    for ch in elem:
        if strip_ns(ch.tag) == tag_name:
            return ch
    return None


def find_descendant_text(elem: Optional[ET.Element], tag_name: str) -> str:
    """Tìm text của thẻ con cháu đầu tiên có tên khớp."""
    if elem is None:
        return ""
    for node in elem.iter():
        if strip_ns(node.tag) == tag_name:
            return (node.text or "").strip()
    return ""


@dataclass
class TaxDeclaration:
    """Đại diện dữ liệu đã trích xuất từ 1 file XML tờ khai thuế."""

    file_path: str
    file_name: str
    tax_type: str  # GTGT, TNCN, TNDN, BCTC, Khác
    ma_tkhai: str = ""
    ten_tkhai: str = ""
    mo_ta_bmau: str = ""
    mst: str = ""
    ten_nnt: str = ""
    dia_chi: str = ""
    ky_kkhai: str = ""
    tu_ngay: Optional[datetime.date] = None
    den_ngay: Optional[datetime.date] = None
    ngay_lap: Optional[datetime.date] = None
    so_lan: str = ""
    ten_cqt: str = ""
    ma_cqt: str = ""
    tieu_muc: str = ""
    indicators: Dict[str, Any] = field(default_factory=dict)
    kqkd_indicators: Dict[str, Any] = field(default_factory=dict)
    appendices: Dict[str, List[Dict[str, Any]]] = field(default_factory=dict)
    has_bctc_sodaunam: bool = False
    error: Optional[str] = None

    @property
    def loai_to_khai(self) -> str:
        if self.so_lan == "0" or self.so_lan == "":
            return "Lần đầu"
        return f"Bổ sung lần {self.so_lan}"

    @property
    def year(self) -> Optional[int]:
        """
        Lấy năm tính thuế của tờ khai.
        QUAN TRỌNG: Ưu tiên tuyệt đối thẻ <kyKKhai> trước (VD: '2020', '01/2020', 'Q1/2020')
        để doanh nghiệp có năm tài chính lệch năm dương lịch (VD: 01/04/2020 - 31/03/2021, kyKKhai='2020')
        KHÔNG BAO GIỜ bị đẩy nhầm sang năm 2021.
        """
        if self.ky_kkhai:
            s_ky = str(self.ky_kkhai).strip().upper()
            # Dạng 4 chữ số năm: '2020', '01/2020', 'Q1/2020'
            m = re.search(r"(19\d\d|20\d\d)", s_ky)
            if m:
                return int(m.group(1))
            # Hỗ trợ dạng kỳ viết tắt YYMM / YYQ# / YYCN (VD: 2001..2012 -> 2020, 2301 -> 2023)
            m2 = re.match(r"^(\d{2})(0[1-9]|1[0-2]|Q\d|CN)$", s_ky)
            if m2:
                return 2000 + int(m2.group(1))
        if self.tu_ngay:
            return self.tu_ngay.year
        if self.den_ngay:
            return self.den_ngay.year
        if self.ngay_lap:
            return self.ngay_lap.year
        return None


BCTC_KQHDSXKD_NAMES: Dict[str, str] = {
    "01": "1. Doanh thu bán hàng và cung cấp dịch vụ",
    "02": "2. Các khoản giảm trừ doanh thu",
    "10": "3. Doanh thu thuần về bán hàng và cung cấp dịch vụ (10 = 01 - 02)",
    "11": "4. Giá vốn hàng bán",
    "20": "5. Lợi nhuận gộp về bán hàng và cung cấp dịch vụ (20 = 10 - 11)",
    "21": "6. Doanh thu hoạt động tài chính",
    "22": "7. Chi phí tài chính",
    "23": "  - Trong đó: Chi phí lãi vay",
    "24": "8. Phần lãi hoặc lỗ trong công ty liên doanh, liên kết",
    "25": "9. Chi phí bán hàng",
    "26": "10. Chi phí quản lý doanh nghiệp",
    "30": "11. Lợi nhuận thuần từ hoạt động kinh doanh {30 = 20 + (21 - 22) + 24 - (25 + 26)}",
    "31": "12. Thu nhập khác",
    "32": "13. Chi phí khác",
    "40": "14. Lợi nhuận khác (40 = 31 - 32)",
    "50": "15. Tổng lợi nhuận kế toán trước thuế (50 = 30 + 40)",
    "51": "16. Chi phí thuế TNDN hiện hành",
    "52": "17. Chi phí thuế TNDN hoãn lại",
    "60": "18. Lợi nhuận sau thuế thu nhập doanh nghiệp (60 = 50 - 51 - 52)",
    "70": "19. Lãi cơ bản trên cổ phiếu (*)",
    "71": "20. Lãi suy giảm trên cổ phiếu (*)",
}

BCTC_LCTTGT_NAMES: Dict[str, str] = {
    "01": "1. Lợi nhuận trước thuế",
    "02": "2. Điều chỉnh cho các khoản: Khấu hao TSCĐ và BĐSĐT",
    "03": "  - Các khoản dự phòng",
    "04": "  - Lãi, lỗ chênh lệch tỷ giá hối đoái do đánh giá lại các khoản mục tiền tệ có gốc ngoại tệ",
    "05": "  - Lãi, lỗ từ hoạt động đầu tư",
    "06": "  - Chi phí lãi vay",
    "07": "  - Các khoản điều chỉnh khác",
    "08": "3. Lợi nhuận từ hoạt động kinh doanh trước thay đổi vốn lưu động",
    "09": "  - Tăng, giảm các khoản phải thu",
    "10": "  - Tăng, giảm hàng tồn kho",
    "11": "  - Tăng, giảm các khoản phải trả (Không kể lãi vay phải trả, thuế TNDN phải nộp)",
    "12": "  - Tăng, giảm chi phí trả trước",
    "13": "  - Tăng, giảm chứng khoán kinh doanh",
    "14": "  - Tiền lãi vay đã trả",
    "15": "  - Thuế thu nhập doanh nghiệp đã nộp",
    "16": "  - Tiền thu khác từ hoạt động kinh doanh",
    "17": "  - Tiền chi khác cho hoạt động kinh doanh",
    "20": "Lưu chuyển tiền thuần từ hoạt động kinh doanh",
    "21": "1. Tiền chi để mua sắm, xây dựng TSCĐ và các tài sản dài hạn khác",
    "22": "2. Tiền thu từ thanh lý, nhượng bán TSCĐ và các tài sản dài hạn khác",
    "23": "3. Tiền chi cho vay, mua các công cụ nợ của đơn vị khác",
    "24": "4. Tiền thu hồi cho vay, bán lại các công cụ nợ của đơn vị khác",
    "25": "5. Tiền chi đầu tư góp vốn vào đơn vị khác",
    "26": "6. Tiền thu hồi đầu tư góp vốn vào đơn vị khác",
    "27": "7. Tiền thu lãi cho vay, cổ tức và lợi nhuận được chia",
    "30": "Lưu chuyển tiền thuần từ hoạt động đầu tư",
    "31": "1. Tiền thu từ phát hành cổ phiếu, nhận vốn góp của chủ sở hữu",
    "32": "2. Tiền trả lại vốn góp cho các chủ sở hữu, mua lại cổ phiếu của doanh nghiệp đã phát hành",
    "33": "3. Tiền thu từ đi vay",
    "34": "4. Tiền trả nợ gốc vay",
    "35": "5. Tiền trả nợ gốc thuê tài chính",
    "36": "6. Cổ tức, lợi nhuận đã trả cho chủ sở hữu",
    "40": "Lưu chuyển tiền thuần từ hoạt động tài chính",
    "50": "Lưu chuyển tiền thuần trong kỳ (50 = 20 + 30 + 40)",
    "60": "Tiền và tương đương tiền đầu kỳ",
    "61": "Ảnh hưởng của thay đổi tỷ giá hối đoái quy đổi ngoại tệ",
    "70": "Tiền và tương đương tiền cuối kỳ (70 = 50 + 60 + 61)",
}


def _extract_bctc_pl_table(pl: ET.Element, pl_upper: str = "") -> List[Dict[str, Any]]:
    """
    Trích xuất bảng Báo cáo KQHDSXKD (PL_KQHDSXKD) hoặc Lưu chuyển tiền tệ (PL_LCTTGT / PL_LCTTTT)
    trong bộ BCTC thành danh sách các dòng chuẩn: (ma_chi_tieu, ten_chi_tieu, thuyet_minh, nam_nay, nam_truoc).
    """
    tm_node = find_child(pl, "ThuyetMinh")
    nn_node = find_child(pl, "NamNay")
    nt_node = find_child(pl, "NamTruoc")
    if nn_node is None and nt_node is None:
        return []

    name_catalog = BCTC_KQHDSXKD_NAMES if "KQHDSXKD" in pl_upper else BCTC_LCTTGT_NAMES
    tm_map: Dict[str, Any] = {}
    nn_map: Dict[str, Any] = {}
    nt_map: Dict[str, Any] = {}
    all_codes: List[str] = []

    for src_node, target_map in ((tm_node, tm_map), (nn_node, nn_map), (nt_node, nt_map)):
        if src_node is not None:
            for ch in src_node:
                t = strip_ns(ch.tag)
                if t not in all_codes:
                    all_codes.append(t)
                target_map[t] = parse_value(ch.text)

    rows: List[Dict[str, Any]] = []
    for code_tag in all_codes:
        clean_code = code_tag[2:] if code_tag.lower().startswith("ct") else code_tag
        tm_val = tm_map.get(code_tag)
        nn_val = nn_map.get(code_tag, 0)
        nt_val = nt_map.get(code_tag, 0)
        rows.append(
            {
                "ma_chi_tieu": clean_code,
                "ten_chi_tieu": name_catalog.get(clean_code, f"Chỉ tiêu {clean_code}"),
                "thuyet_minh": tm_val if tm_val is not None else "",
                "nam_nay": nn_val if nn_val is not None else 0,
                "nam_truoc": nt_val if nt_val is not None else 0,
            }
        )
    return rows


def extract_appendices(pluc_elem: Optional[ET.Element]) -> Dict[str, List[Dict[str, Any]]]:
    """
    Trích xuất dữ liệu dạng bảng từ thẻ <PLuc>.
    Mỗi phụ lục sẽ trả về danh sách các dòng (dict) không trùng lặp khóa alias.
    Hỗ trợ đầy đủ:
    - PLuc_05_1, PLuc_05_2, PLuc_05_3 (kể cả khi bảng chỉ có 1 người lao động / 1 NPT)
    - PL_NQ43_GTGT, PL_NQ101_GTGT, PL_NQ142_GTGT (gồm cả Mua vào và Bán ra, kể cả khi chỉ có 1 dòng)
    - PL01_KHBS (chuẩn hóa 1 dòng tổng hợp trên mỗi tờ khai bổ sung cho cả TT156 và TT80)
    - PL01_1_KHBS (tách đúng từng dòng ChiTiet chỉ tiêu điều chỉnh kèm ma_HSo, ten_HSo)
    - PL_GDLK2-01 (kết hợp thông tin bên liên kết Mục I và cờ miễn kê khai Mục II, loại bỏ dòng rỗng Mục III)
    - PLuc_03_2_TNDN & PL03_2A_TNDN (tách đúng dòng năm chuyển lỗ kèm phân loại nhóm lỗ)
    - PL_KQHDSXKD & PL_LCTTGT (bóc tách chuẩn theo từng mã chỉ tiêu, tên chỉ tiêu, Năm nay / Năm trước)
    """
    result: Dict[str, List[Dict[str, Any]]] = {}
    if pluc_elem is None:
        return result

    single_row_keywords = (
        "ctiet",
        "chitiet",
        "dong",
        "row",
        "item",
        "npt",
        "bangke",
        "bke",
        "hhdv",
        "noidungnamthu",
        "diadiem",
    )

    for pl in pluc_elem:
        pl_name = strip_ns(pl.tag)
        pl_upper = pl_name.upper()
        rows: List[Dict[str, Any]] = []

        # 1. Xử lý chuyên biệt cho PL_KQHDSXKD và PL_LCTTGT / PL_LCTTTT (BCTC)
        if pl_upper in ("PL_KQHDSXKD", "PL_LCTTGT", "PL_LCTTTT"):
            bctc_rows = _extract_bctc_pl_table(pl, pl_upper)
            if bctc_rows:
                result[pl_name] = bctc_rows
                continue

        # 2. Xử lý chuyên biệt cho PL01_KHBS (Tờ khai bổ sung tổng hợp 01/KHBS)
        if pl_upper == "PL01_KHBS":
            header_node = find_child(pl, "Header")
            muc_a = find_child(pl, "Muc_A")
            muc_b = find_child(pl, "Muc_B")
            nd_dc = find_child(pl, "NDUNGKHAIBOSUNGDIEUCHINH")
            if muc_a is not None or header_node is not None:
                # Cấu trúc TT80: 1 dòng tổng hợp duy nhất cho tờ khai 01/KHBS
                muc_i = find_child(muc_a, "Muc_I") if muc_a is not None else None
                muc_1 = find_child(muc_i, "Muc_1") if muc_i is not None else None
                muc_2 = find_child(muc_i, "Muc_2") if muc_i is not None else None
                muc_3 = find_child(muc_i, "Muc_3") if muc_i is not None else None
                muc_ii = find_child(muc_a, "Muc_II") if muc_a is not None else None
                muc_iii = find_child(muc_a, "Muc_III") if muc_a is not None else None

                ngay_cn_raw = find_descendant_text(muc_3, "ct3a_1")
                if ngay_cn_raw in ("0001-01-01", "01/01/0001"):
                    ngay_cn_raw = ""
                khbs_row: Dict[str, Any] = {
                    "ma_giao_dich_goc": find_descendant_text(header_node, "maGiaoDich"),
                    "tieu_muc_ma": find_descendant_text(muc_1, "ct2_ma"),
                    "tieu_muc_ten": find_descendant_text(muc_1, "ct2_ten"),
                    "thue_phai_nop_tang_giam": parse_value(
                        find_descendant_text(muc_1, "tongCong_ct10") or find_descendant_text(muc_1, "ct3")
                    ) or 0,
                    "thue_hoan_phai_nop_tang_giam": parse_value(find_descendant_text(muc_2, "tongCong_ct11")) or 0,
                    "thue_khau_tru_tang_giam": parse_value(find_descendant_text(muc_ii, "tongCong_ct12")) or 0,
                    "thue_de_nghi_hoan_tang_giam": parse_value(find_descendant_text(muc_iii, "tongCong_ct13")) or 0,
                    "ngay_tinh_cham_nop": ngay_cn_raw,
                    "so_ngay_cham_nop": parse_value(find_descendant_text(muc_3, "ct3a_2")) or 0,
                    "tien_cham_nop": parse_value(find_descendant_text(muc_3, "ct3b")) or 0,
                    "ly_do_giai_trinh": "",
                }
                result[pl_name] = [khbs_row]
                continue
            elif nd_dc is not None:
                # Cấu trúc TT156: Tờ khai 01/KHBS kèm chi tiết điều chỉnh tăng/giảm
                phat_node = find_child(pl, "PhatChamNop")
                gt_node = find_child(pl, "NOIDUNGGIAITHICH")
                thop_node = find_child(nd_dc, "TongHopDCTangGiamPNop")
                ct40_node = find_child(thop_node, "ct40") if thop_node is not None else None
                ct43_node = find_child(thop_node, "ct43") if thop_node is not None else None

                khbs_row = {
                    "ma_giao_dich_goc": "",
                    "tieu_muc_ma": "",
                    "tieu_muc_ten": "",
                    "thue_phai_nop_tang_giam": parse_value(find_descendant_text(ct40_node, "ct6")) or 0,
                    "thue_hoan_phai_nop_tang_giam": 0,
                    "thue_khau_tru_tang_giam": parse_value(find_descendant_text(ct43_node, "ct6")) or 0,
                    "thue_de_nghi_hoan_tang_giam": 0,
                    "ngay_tinh_cham_nop": find_descendant_text(gt_node, "ngay"),
                    "so_ngay_cham_nop": parse_value(find_descendant_text(phat_node, "soNgayNopCham")) or 0,
                    "tien_cham_nop": parse_value(find_descendant_text(phat_node, "soTienPhatNopCham")) or 0,
                    "ly_do_giai_trinh": find_descendant_text(gt_node, "lyDoKhac"),
                }
                result[pl_name] = [khbs_row]

                # Trích xuất thêm các dòng chỉ tiêu điều chỉnh chi tiết của TT156 vào PL01_1_KHBS
                detail_rows: List[Dict[str, Any]] = []
                for sub_el in nd_dc.iter():
                    stag = strip_ns(sub_el.tag)
                    if stag in ("KHBSDieuChinhTang", "KHBSDieuChinhGiam", "ct40", "ct43"):
                        c_ten = find_descendant_text(sub_el, "ct2")
                        c_ma = find_descendant_text(sub_el, "ct3")
                        c4 = parse_value(find_descendant_text(sub_el, "ct4")) or 0
                        c5 = parse_value(find_descendant_text(sub_el, "ct5")) or 0
                        c6 = parse_value(find_descendant_text(sub_el, "ct6")) or 0
                        if c_ten or c_ma or c4 or c5 or c6:
                            ma_fmt = f"[{c_ma}]" if c_ma and not c_ma.startswith("[") else c_ma
                            detail_rows.append(
                                {
                                    "ma_HSo": "",
                                    "ten_HSo": "Tờ khai chính (TT156)",
                                    "ct3_ma": ma_fmt,
                                    "ct3_ten": c_ten,
                                    "ct04": c4,
                                    "ct05": c5,
                                    "ct06": c6,
                                    "ct7": 0,
                                    "ct3_2_ten": "",
                                }
                            )
                if detail_rows:
                    result.setdefault("PL01_1_KHBS", []).extend(detail_rows)
                continue

        # 3. Xử lý chuyên biệt cho PL01_1_KHBS (Bảng giải trình khai bổ sung)
        if "PL01_1_KHBS" in pl_upper:
            for node in pl.iter():
                if strip_ns(node.tag) == "BKeHSo":
                    ma_hso = find_descendant_text(node, "ma_HSo")
                    ten_hso = find_descendant_text(node, "ten_HSo")
                    ctiet_hso = find_child(node, "CTietHSo")
                    if ctiet_hso is not None:
                        for ct in ctiet_hso:
                            if strip_ns(ct.tag) == "ChiTiet":
                                r_dict = flatten_element(ct, include_aliases=False)
                                if r_dict:
                                    row_out: Dict[str, Any] = {}
                                    if ma_hso:
                                        row_out["ma_HSo"] = ma_hso
                                    if ten_hso:
                                        row_out["ten_HSo"] = ten_hso
                                    row_out.update(r_dict)
                                    rows.append(row_out)
            if rows:
                result.setdefault(pl_name, []).extend(rows)
                continue

        # 4. Xử lý chuyên biệt cho PL_GDLK2-01 (Giao dịch liên kết NĐ 132/2020/NĐ-CP)
        if "GDLK" in pl_upper:
            muc_i = find_child(pl, "Muc_I")
            muc_ii = find_child(pl, "Muc_II")
            muc_ii_info: Dict[str, Any] = {}
            if muc_ii is not None:
                for ch in muc_ii:
                    muc_ii_info[f"Muc_II_{strip_ns(ch.tag)}"] = parse_value(ch.text)

            if muc_i is not None:
                for ct in muc_i:
                    if strip_ns(ct.tag) == "CTiet":
                        r_dict = flatten_element(ct, include_aliases=False)
                        if r_dict and (r_dict.get("tenLKet") or r_dict.get("mstLKet")):
                            r_dict.update(muc_ii_info)
                            rows.append(r_dict)
            if rows:
                result[pl_name] = rows
                continue

        # 5. Xử lý chuyên biệt cho PLuc_03_2_TNDN / PL03_2A_TNDN (Chuyển lỗ)
        if "03_2" in pl_upper and "TNDN" in pl_upper:
            for parent_sec_name, group_label in (
                ("Lo_HDSXKD", "Lỗ từ HĐ SXKD"),
                ("Lo_CNBDS", "Lỗ từ chuyển nhượng BĐS"),
            ):
                sec_node = find_child(pl, parent_sec_name)
                if sec_node is not None:
                    bke = find_child(sec_node, "BangChiTietSoLo")
                    if bke is not None:
                        for ch in bke:
                            if strip_ns(ch.tag).lower().startswith("noidungnamthu"):
                                r_dict = flatten_element(ch, include_aliases=False)
                                if r_dict:
                                    # Bỏ qua các dòng Lo_CNBDS trống toàn số 0
                                    has_nonzero = any(
                                        r_dict.get(k) not in (None, 0, "0", "")
                                        for k in ("ct3", "ct4", "ct5", "ct6", "soLoPSinh", "soLoDaChuyen", "soLoKyNay", "soLoChuyenSau")
                                    )
                                    if parent_sec_name == "Lo_HDSXKD" or has_nonzero:
                                        rows.append(
                                            {
                                                "nhom_lo": group_label,
                                                "namPSinhLo": r_dict.get("ct2") or r_dict.get("namPSinhLo"),
                                                "soLoPSinh": r_dict.get("ct3", r_dict.get("soLoPSinh", 0)),
                                                "soLoDaChuyen": r_dict.get("ct4", r_dict.get("soLoDaChuyen", 0)),
                                                "soLoKyNay": r_dict.get("ct5", r_dict.get("soLoKyNay", 0)),
                                                "soLoChuyenSau": r_dict.get("ct6", r_dict.get("soLoChuyenSau", 0)),
                                            }
                                        )
            # Cấu trúc TT156 PL03_2A_TNDN: BangChiTietSoLo nằm trực tiếp dưới PL03_2A_TNDN
            if not rows:
                bke = find_child(pl, "BangChiTietSoLo")
                if bke is not None:
                    for ch in bke:
                        if strip_ns(ch.tag).lower().startswith("noidungnamthu"):
                            r_dict = flatten_element(ch, include_aliases=False)
                            if r_dict:
                                rows.append(
                                    {
                                        "nhom_lo": "Lỗ từ HĐ SXKD",
                                        "namPSinhLo": r_dict.get("namPSinhLo") or r_dict.get("ct2"),
                                        "soLoPSinh": r_dict.get("soLoPSinh", r_dict.get("ct3", 0)),
                                        "soLoDaChuyen": r_dict.get("soLoDaChuyen", r_dict.get("ct4", 0)),
                                        "soLoKyNay": r_dict.get("soLoKyNay", r_dict.get("ct5", 0)),
                                        "soLoChuyenSau": r_dict.get("soLoChuyenSau", r_dict.get("ct6", 0)),
                                    }
                                )
            if rows:
                result[pl_name] = rows
                continue

        # 4. Thuật toán đệ quy tổng quát tìm dòng lặp cho mọi phụ lục khác
        def collect_repeating_rows(node: ET.Element, group_label: str = "") -> bool:
            children = list(node)
            if not children:
                return False

            counts: Dict[str, int] = {}
            for c in children:
                if len(c) > 0:
                    t = strip_ns(c.tag)
                    counts[t] = counts.get(t, 0) + 1

            # Ưu tiên 1: Thẻ con có cấu trúc xuất hiện > 1 lần VÀ các con của nó là thẻ lá (hoặc 1 cấp dữ liệu dòng)
            multi_tags = {t for t, cnt in counts.items() if cnt > 1}
            if multi_tags:
                # Kiểm tra xem các thẻ lặp này có chứa bảng lặp sâu hơn không (ví dụ BKeHSo chứa CTietHSo/ChiTiet)
                deeper_found = False
                for c in children:
                    ct = strip_ns(c.tag)
                    if ct in multi_tags and len(c) > 0:
                        # Nếu bên trong c không có thẻ lá trực tiếp có dữ liệu mà toàn thẻ nhánh sâu
                        has_grandchildren_containers = any(len(gc) > 0 for gc in c)
                        if has_grandchildren_containers and collect_repeating_rows(c, group_label or ct):
                            deeper_found = True
                if deeper_found:
                    return True

                for c in children:
                    ct = strip_ns(c.tag)
                    if ct in multi_tags and len(c) > 0:
                        row_dict = flatten_element(c, include_aliases=False)
                        if row_dict:
                            if group_label and "nhom_phu_luc" not in row_dict:
                                row_dict = {"nhom_phu_luc": group_label, **row_dict}
                            rows.append(row_dict)
                return True

            # Ưu tiên 2: Đệ quy sâu hơn vào TẤT CẢ các thẻ con nhánh (không dừng sớm ở nhánh đầu tiên!)
            found_deeper = False
            for c in children:
                if len(c) > 0:
                    sub_label = strip_ns(c.tag) if strip_ns(c.tag) in ("HH_DV_MuaVaoTrongKy", "HH_DV_BanRaTrongKy") else group_label
                    if collect_repeating_rows(c, sub_label):
                        found_deeper = True
            if found_deeper:
                return True

            # Ưu tiên 3: Bảng chỉ có 1 dòng duy nhất (con của c đều là thẻ lá)
            found_single = False
            for c in children:
                t = strip_ns(c.tag).lower()
                if len(c) > 0 and all(len(gc) == 0 for gc in c) and any(k in t for k in single_row_keywords):
                    row_dict = flatten_element(c, include_aliases=False)
                    if row_dict:
                        if group_label and "nhom_phu_luc" not in row_dict:
                            row_dict = {"nhom_phu_luc": group_label, **row_dict}
                        rows.append(row_dict)
                        found_single = True
            if found_single:
                return True

            return False

        collect_repeating_rows(pl)

        # Nếu không có bảng lặp nhưng phụ lục có các chỉ tiêu đơn lẻ (như 03-1A/TNDN)
        if not rows:
            flat = flatten_element(pl, include_aliases=False)
            if flat:
                rows.append(flat)

        if rows:
            result[pl_name] = rows

    return result


def parse_tax_xml(file_path: str) -> TaxDeclaration:
    """Đọc và phân tích một file XML tờ khai thuế."""
    file_name = os.path.basename(file_path)
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
    except Exception as exc:
        return TaxDeclaration(
            file_path=file_path,
            file_name=file_name,
            tax_type="Lỗi",
            error=f"Không đọc được XML: {exc}",
        )

    # Tìm <HSoKhaiThue>
    hso_khai = None
    if strip_ns(root.tag) == "HSoKhaiThue":
        hso_khai = root
    else:
        for node in root.iter():
            if strip_ns(node.tag) == "HSoKhaiThue":
                hso_khai = node
                break

    if hso_khai is None:
        return TaxDeclaration(
            file_path=file_path,
            file_name=file_name,
            tax_type="Lỗi",
            error="Không tìm thấy thẻ <HSoKhaiThue> (không phải XML tờ khai thuế chuẩn).",
        )

    ttin_chung = find_child(hso_khai, "TTinChung")
    ttin_tkhai = find_child(ttin_chung, "TTinTKhaiThue")
    tkhai_thue = find_child(ttin_tkhai, "TKhaiThue")
    nnt = find_child(ttin_tkhai, "NNT")

    ma_tkhai = find_descendant_text(tkhai_thue, "maTKhai")
    ten_tkhai = find_descendant_text(tkhai_thue, "tenTKhai")
    mo_ta_bmau = find_descendant_text(tkhai_thue, "moTaBMau")
    ten_cqt = find_descendant_text(tkhai_thue, "tenCQTNoiNop")
    ma_cqt = find_descendant_text(tkhai_thue, "maCQTNoiNop")
    so_lan = (
        find_descendant_text(tkhai_thue, "soLan")
        or find_descendant_text(tkhai_thue, "soLanBoSung")
        or find_descendant_text(tkhai_thue, "soLanBSung")
        or find_descendant_text(tkhai_thue, "lanBoSung")
        or find_descendant_text(tkhai_thue, "lanBSung")
    )
    loai_tkhai_code = find_descendant_text(tkhai_thue, "loaiTKhai").strip().upper()
    if loai_tkhai_code == "B" and (not so_lan or so_lan == "0"):
        so_lan = "1"

    ky_node = find_child(tkhai_thue, "KyKKhaiThue")
    kieu_ky = find_descendant_text(ky_node, "kieuKy").strip().upper()
    ky_kkhai = find_descendant_text(ky_node, "kyKKhai")
    tu_ngay = parse_date(find_descendant_text(ky_node, "kyKKhaiTuNgay"))
    den_ngay = parse_date(find_descendant_text(ky_node, "kyKKhaiDenNgay"))
    ngay_lap = parse_date(find_descendant_text(tkhai_thue, "ngayLapTKhai"))

    mst = normalize_mst(find_descendant_text(nnt, "mst"))
    ten_nnt = find_descendant_text(nnt, "tenNNT")
    dia_chi = find_descendant_text(nnt, "dchiNNT")

    tax_type = classify_tax_type(ma_tkhai, ten_tkhai)

    # 1. Trích xuất chỉ tiêu tờ khai chính (<CTieuTKhaiChinh>)
    ctieu_node = find_child(hso_khai, "CTieuTKhaiChinh")
    indicators: Dict[str, Any] = {}
    if ctieu_node is not None:
        indicators = flatten_element(ctieu_node, include_aliases=True)
    else:
        for child in hso_khai:
            t = strip_ns(child.tag)
            if t not in ("TTinChung", "PLuc", "ChuKyDienTu"):
                child_flat = flatten_element(child, include_aliases=True)
                indicators.update(child_flat)

    tieu_muc = str(indicators.get("tieuMucHachToan") or "")
    if kieu_ky:
        indicators["_kieu_ky"] = kieu_ky

    # 2. Trích xuất phụ lục (<PLuc>) và tách riêng chỉ tiêu KQKD (03-1A/TNDN hoặc PL_KQHDSXKD)
    pluc_node = find_child(hso_khai, "PLuc")
    appendices = extract_appendices(pluc_node)
    kqkd_indicators: Dict[str, Any] = {}
    if pluc_node is not None:
        for pl in pluc_node:
            pt = strip_ns(pl.tag).upper()
            if "03_1A" in pt or "03-1A" in pt:
                kqkd_indicators.update(flatten_element(pl, include_aliases=True))

    # 3. Quét nhánh tờ khai chính (không quét nhầm vào PL01_1_KHBS) để nhận diện các dòng chỉ tiêu BCTC
    scan_root = ctieu_node if ctieu_node is not None else hso_khai
    for elem in scan_root.iter():
        children = list(elem)
        if not children:
            continue

        ms = None
        so_dau = None
        so_cuoi = None
        so_tien = None

        for ch in children:
            ch_tag = strip_ns(ch.tag).lower()
            ch_val = parse_value(ch.text)
            if ch_val is None:
                continue

            if ch_tag in ("maso", "ms", "ma", "code", "machitieu"):
                ms = str(ch_val).strip()
            elif any(k in ch_tag for k in ("sodaunam", "sodauky", "dauky", "daunam", "sodau", "namtruoc", "kytruoc")):
                so_dau = ch_val
            elif any(k in ch_tag for k in ("socuoinam", "socuoiky", "cuoiky", "cuoinam", "socuoi", "namnay", "kynay")):
                so_cuoi = ch_val
            elif ch_tag in ("sotien", "giatri", "tien", "val"):
                so_tien = ch_val

        elem_t = strip_ns(elem.tag).lower()
        if not ms:
            m_tag = re.fullmatch(r"(?:ms|ct)(\d+[a-z]?)", elem_t)
            if m_tag:
                ms = m_tag.group(1)

        if ms:
            clean_ms = ms.replace("[", "").replace("]", "").strip()
            if so_dau is not None:
                indicators[f"{clean_ms}_dau"] = so_dau
                indicators[f"ms{clean_ms}_dau"] = so_dau
                indicators[f"ct{clean_ms}_dau"] = so_dau
                indicators[f"{clean_ms}_sodaunam"] = so_dau
            if so_cuoi is not None:
                indicators[f"{clean_ms}_cuoi"] = so_cuoi
                indicators[f"ms{clean_ms}_cuoi"] = so_cuoi
                indicators[f"ct{clean_ms}_cuoi"] = so_cuoi
                indicators[f"{clean_ms}_socuoinam"] = so_cuoi
            if so_tien is not None:
                indicators[clean_ms] = so_tien
                indicators[f"ct{clean_ms}"] = so_tien
                indicators[f"ms{clean_ms}"] = so_tien
            if so_cuoi is not None and clean_ms not in indicators:
                indicators[clean_ms] = so_cuoi

    # 4. Trích xuất toàn diện số đầu năm / số cuối năm BCTC từ các khối lồng nhau (TT200, TT133, QĐ48)
    # Bảo toàn hậu tố chữ cái (411a, 411b, 421a, 421b) và ưu tiên HoatDongLienTuc!
    dau_tags = ("sodaunam", "sodauky", "dauky", "daunam", "sodau", "namtruoc", "kytruoc")
    cuoi_tags = ("socuoinam", "socuoiky", "cuoiky", "cuoinam", "socuoi", "namnay", "kynay")

    has_bctc_sodaunam = False
    for k, v in list(indicators.items()):
        if v is None:
            continue
        k_lower = k.lower()
        leaf = k_lower.split(".")[-1]
        m = re.fullmatch(r"(?:ct|ms|m)?(\d+[a-z]?)", leaf)
        if not m:
            continue
        code = m.group(1)

        is_dau = any(tag in k_lower for tag in dau_tags)
        is_cuoi = any(tag in k_lower for tag in cuoi_tags)
        is_khong_lt = "khonglientuc" in k_lower

        if is_dau:
            if not is_khong_lt and v not in (0, "0", ""):
                has_bctc_sodaunam = True
            key_dau = f"{code}_dau"
            cur = indicators.get(key_dau)
            if cur is None or (cur == 0 and v != 0 and not is_khong_lt):
                indicators[key_dau] = v
                indicators[f"ct{code}_dau"] = v
                indicators[f"ms{code}_dau"] = v
                indicators[f"{code}_sodaunam"] = v

        if is_cuoi:
            key_cuoi = f"{code}_cuoi"
            cur = indicators.get(key_cuoi)
            if cur is None or (cur == 0 and v != 0 and not is_khong_lt):
                indicators[key_cuoi] = v
                indicators[f"ct{code}_cuoi"] = v
                indicators[f"ms{code}_cuoi"] = v
                indicators[f"{code}_socuoinam"] = v
                if code not in indicators or (indicators[code] == 0 and not is_khong_lt):
                    indicators[code] = v

    if tax_type == "NTNN":
        from ntnn_exporter import extract_ntnn_contract_rows_from_xml_root, extract_ntnn_totals_from_xml_root
        ntnn_yr = None
        if ky_kkhai:
            m_yr = re.search(r"(19\d\d|20\d\d)", str(ky_kkhai))
            if m_yr:
                ntnn_yr = int(m_yr.group(1))
        if ntnn_yr is None and tu_ngay:
            ntnn_yr = tu_ngay.year
        elif ntnn_yr is None and ngay_lap:
            ntnn_yr = ngay_lap.year
        ntnn_m = extract_ntnn_totals_from_xml_root(
            root,
            ma_tkhai=ma_tkhai,
            mo_ta_bmau=mo_ta_bmau,
            ten_tkhai=ten_tkhai,
            year=ntnn_yr,
        )
        ntnn_rows = extract_ntnn_contract_rows_from_xml_root(
            root,
            ma_tkhai=ma_tkhai,
            mo_ta_bmau=mo_ta_bmau,
            ten_tkhai=ten_tkhai,
            year=ntnn_yr,
        )
        indicators["_ntnn_dt_chua_gtgt"] = ntnn_m["dt_chua_gtgt"]
        indicators["_ntnn_dt_gtgt"] = ntnn_m["dt_gtgt"]
        indicators["_ntnn_thue_gtgt"] = ntnn_m["thue_gtgt"]
        indicators["_ntnn_dt_tndn"] = ntnn_m["dt_tndn"]
        indicators["_ntnn_thue_tndn"] = ntnn_m["thue_tndn"]
        indicators["_ntnn_tong_thue"] = ntnn_m["tong_thue"]
        indicators["_ntnn_schema"] = ntnn_m["schema"]
        indicators["_ntnn_rows"] = ntnn_rows

    return TaxDeclaration(
        file_path=file_path,
        file_name=file_name,
        tax_type=tax_type,
        ma_tkhai=ma_tkhai,
        ten_tkhai=ten_tkhai,
        mo_ta_bmau=mo_ta_bmau,
        mst=mst,
        ten_nnt=ten_nnt,
        dia_chi=dia_chi,
        ky_kkhai=ky_kkhai,
        tu_ngay=tu_ngay,
        den_ngay=den_ngay,
        ngay_lap=ngay_lap,
        so_lan=so_lan,
        ten_cqt=ten_cqt,
        ma_cqt=ma_cqt,
        tieu_muc=tieu_muc,
        indicators=indicators,
        kqkd_indicators=kqkd_indicators,
        appendices=appendices,
        has_bctc_sodaunam=has_bctc_sodaunam,
    )
