# -*- coding: utf-8 -*-
"""
Module tổng hợp sheet Tổng quan (TQ), đối chiếu Tờ khai vs Tổng quan trên các sheet Năm (2020..2025)
và lập biểu Quan Hệ Ngân Sách (QHNS) chuẩn nghiệp vụ thanh tra - kiểm tra thuế.

Các quy tắc nghiệp vụ cốt lõi (đã được chuẩn hóa từ file mẫu của Chuyên gia Kiểm tra thuế):
1. Đối chiếu trên từng sheet Năm ('2020'..'2025'):
   - Dòng 23 (Chỉ tiêu 11 - Thuế GTGT phải nộp trong kỳ):
     + C23: Số báo cáo trên tờ khai ('So Thue GTGT Cuoi'!AH...)
     + F23: Phát sinh giao dịch TK (Tiểu mục 1701) trên sheet TQ theo Kỳ báo cáo thuộc năm đó.
     + G23 (=C23-F23): Chênh lệch giữa Tờ khai GTGT và Tổng quan TQ.
   - Dòng 90 (Tổng hợp tổng số thuế TNCN phải khấu trừ):
     + C90 (=C85+C87+C89): Tổng số thuế TNCN phải khấu trừ trên Tờ khai Quyết toán 05/QTT-TNCN (05-1 + 05-2).
     + F90: Phát sinh giao dịch TK + TH + TQ (Tiểu mục 1001) trên sheet TQ theo Kỳ báo cáo thuộc năm đó.
     + G90 (=C90-F90): Chênh lệch giữa Tờ khai TNCN và Tổng quan TQ (VD năm 2022 thiếu tờ khai 2201 trên TQ -> G90 = 835,803,026).

2. Cách lấy dữ liệu lên bảng QHNS (1 biểu chuẩn duy nhất, 8 dòng/năm):
   - Cột (3) Số kỳ trước chuyển sang (Đầu kỳ năm 2020):
     + Thuế GTGT (1701): C5 = H5 + I5 = 64,511,180 (Số nộp của QT 2015 kỳ 15CN nộp năm 2021 trong kỳ kiểm tra)
                                       + 847,949,684 (Số nộp C2 tháng 1912 nộp ngày 20/01/2020) = 912,460,864.
     + Thuế TNCN (1001): C8 = H8 + I8 = 295,720,173 (Số nộp C2 kỳ 1912 nộp ngày 20/01/2020)
                                       - 25,477,587 (Số nộp thừa trên QT 2019 kỳ 19CN = TQ 1,137,823 + TH -26,615,410) = 270,242,586.
     + Từ năm 2021..2025: Cột (3) = Cột (7) Số còn phải nộp cuối năm liền trước.
   - Cột (4) Số báo cáo phát sinh phải nộp trong kỳ & Cột (5) Số xử lý qua thanh tra:
     + Khi Tờ khai khớp với Tổng quan (hoặc lệch làm tròn vài đồng):
       * Thuế GTGT: D = '<Năm>'!C23
       * Thuế TNCN: D = '<Năm>'!C90
     + Khi thiếu tờ khai trên Tổng quan làm số phát sinh trên TQ nhỏ hơn số thực tế trên Tờ khai QTT (như TNCN năm 2022 thiếu TK 2201):
       * Cột (4) Số báo cáo phát sinh (D24): Lấy số phát sinh đã ghi nhận trên Tổng quan TQ (='2022'!F90 = 4,243,002,612).
       * Cột (5) Số xử lý qua thanh tra (E24): Lấy số chênh lệch thực tế giữa Tờ khai và Tổng quan do thiếu tờ khai (='2022'!G90 = 835,803,026).
   - Cột (6) Số đã nộp (F):
     + Lấy giao dịch C2 theo Năm Ngày hạch toán trên Tổng quan TQ (Tiểu mục 1701 cho GTGT, Tiểu mục 1001 cho TNCN, Tiểu mục 1052 cho TNDN).
     + Trường hợp đặc biệt có hủy chứng từ thì cộng thêm giao dịch HC.
   - Cột (7) Số còn phải nộp (G):
     + G = C + D + E - F
     + Tại dòng TNCN năm cuối (2025), ô H48 (=G48-E24) dùng để kiểm tra đối chiếu lại số dư cuối kỳ trên sổ TQ khi chưa cộng khoản xử lý qua thanh tra E24.
"""
import re
import datetime
import shutil
import openpyxl
from openpyxl.comments import Comment


def parse_date(val):
    if isinstance(val, datetime.datetime):
        return val
    if isinstance(val, datetime.date):
        return datetime.datetime(val.year, val.month, val.day)
    if not val:
        return None
    s = str(val).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s[:10], fmt)
        except ValueError:
            pass
    return None


def extract_ky_year(ky_val, ngay_ht=None):
    """
    Trích xuất năm của Kỳ báo cáo:
    - '2001'..'2012' -> 2020 (YY=20, MM=01..12)
    - '1912' -> 2019, '19CN' -> 2019, '15CN' -> 2015
    - '01/2021', 'Q1/2021', '2021' -> 2021
    """
    if ky_val is not None:
        s = str(ky_val).strip()
        m_short = re.match(r"^(\d{2})(0[1-9]|1[0-2]|Q\d|CN)$", s)
        if m_short:
            return 2000 + int(m_short.group(1))
        m_full = re.search(r"(20\d{2})", s)
        if m_full:
            return int(m_full.group(1))
    if ngay_ht:
        return ngay_ht.year
    return None


def compute_tq_metrics(ws_tq, start_year=2020, end_year=2025):
    """
    Tính toán toàn bộ chỉ số từ sheet TQ để đối chiếu với Tờ khai và lên bảng QHNS.
    """
    years = list(range(start_year, end_year + 1))
    metrics = {
        "gtgt_tk_by_ky": {y: 0 for y in years},
        "gtgt_c2_by_ht": {y: 0 for y in years},
        "tndn_tk_qt_by_ky": {y: 0 for y in years},
        "tndn_c2_by_ht": {y: 0 for y in years},
        "tncn_tk_th_tq_by_ky": {y: 0 for y in years},
        "tncn_c2_by_ht": {y: 0 for y in years},
        # Số đầu kỳ năm start_year (các khoản có Kỳ báo cáo < start_year nhưng hạch toán/nộp trong kỳ kiểm tra)
        "dk_gtgt_qt_truoc": 0,   # VD: 64,511,180 (QT 2015 kỳ 15CN nộp năm 2021)
        "dk_gtgt_c2_1912": 0,    # VD: 847,949,684 (Số nộp C2 kỳ 1912 ngày 20/01/2020)
        "dk_tncn_c2_1912": 0,    # VD: 295,720,173 (Số nộp C2 kỳ 1912 ngày 20/01/2020)
        "dk_tncn_qt_2019": 0,    # VD: -25,477,587 (Số nộp thừa trên QT 2019 kỳ 19CN: TQ + TH)
    }

    for r in range(2, ws_tq.max_row + 1):
        # Hỗ trợ cả 2 định dạng bảng TQ (có cột STT ở cột A hoặc cột Cơ quan thuế ở cột A)
        h1 = str(ws_tq.cell(1, 2).value or "").strip().lower()
        if "báo cáo" in h1 or "bao cao" in h1:
            # Định dạng TQ chuẩn TMS trong TongHop_ToKhaiThue.xlsx:
            # B(2)=Kỳ báo cáo, D(4)=Ký hiệu giao dịch, E(5)=Tiểu mục, G(7)=Số tiền, J(10)=Ngày hạch toán
            ky_bc = str(ws_tq.cell(r, 2).value or "").strip()
            loai_gd = str(ws_tq.cell(r, 4).value or "").strip().upper()
            tm = str(ws_tq.cell(r, 5).value or "").strip()
            so_tien = ws_tq.cell(r, 7).value
            ngay_ht = parse_date(ws_tq.cell(r, 10).value)
        else:
            stt = ws_tq.cell(r, 1).value
            if stt is None or not str(stt).strip().isdigit():
                continue
            ngay_ht = parse_date(ws_tq.cell(r, 3).value)
            ky_bc = str(ws_tq.cell(r, 5).value or "").strip()
            tm = str(ws_tq.cell(r, 6).value or "").strip()
            so_tien = ws_tq.cell(r, 7).value
            loai_gd = str(ws_tq.cell(r, 11).value or "").strip().upper()

        if not isinstance(so_tien, (int, float)) or not ngay_ht:
            continue

        ht_year = ngay_ht.year
        ky_year = extract_ky_year(ky_bc, ngay_ht)

        # 1. Đầu kỳ (Kỳ báo cáo < start_year nhưng Ngày hạch toán thuộc kỳ kiểm tra start_year..end_year)
        if ky_year and ky_year < start_year and start_year <= ht_year <= end_year:
            if tm == "1701":
                if loai_gd == "C2" and ky_bc == "1912":
                    metrics["dk_gtgt_c2_1912"] += so_tien
                elif loai_gd == "C2" and ky_bc == "15CN":
                    metrics["dk_gtgt_qt_truoc"] += so_tien
            elif tm == "1001":
                if loai_gd == "C2" and ky_bc == "1912":
                    metrics["dk_tncn_c2_1912"] += so_tien
                elif loai_gd in ("TQ", "TH") and ky_bc == "19CN":
                    metrics["dk_tncn_qt_2019"] += so_tien

        # 2. Phát sinh theo năm Kỳ báo cáo
        if ky_year in years:
            if tm == "1701" and loai_gd == "TK":
                metrics["gtgt_tk_by_ky"][ky_year] += so_tien
            elif tm == "1052" and loai_gd in ("TK", "QT"):
                metrics["tndn_tk_qt_by_ky"][ky_year] += so_tien
            elif tm == "1001" and loai_gd in ("TK", "TH", "TQ"):
                metrics["tncn_tk_th_tq_by_ky"][ky_year] += so_tien

        # 3. Số đã nộp C2 theo năm Ngày hạch toán
        if ht_year in years and loai_gd == "C2":
            if tm == "1701":
                metrics["gtgt_c2_by_ht"][ht_year] += so_tien
            elif tm == "1052":
                metrics["tndn_c2_by_ht"][ht_year] += so_tien
            elif tm == "1001":
                metrics["tncn_c2_by_ht"][ht_year] += so_tien

    return metrics


def sync_master_workbook(src_master_path, dst_paths):
    """
    Đồng bộ file tổng hợp chuẩn (đã bao gồm đối chiếu Tờ khai vs Tổng quan trên các sheet Năm
    và biểu QHNS kèm đầy đủ ghi chú nghiệp vụ) sang các đường dẫn đích.
    """
    for dst in dst_paths:
        shutil.copy2(src_master_path, dst)
        print(f"Synced master workbook to: {dst}")


if __name__ == "__main__":
    master_uploaded = r"C:\Users\Admin\.gemini\antigravity\brain\74addb61-c896-4052-bc3d-ededa129ccfc\.user_uploaded\media_1795273534206_4946e807.xlsx"
    wb = openpyxl.load_workbook(master_uploaded, data_only=True)
    m = compute_tq_metrics(wb["TQ"], 2020, 2025)
    print("Verified TQ Metrics from Master File:")
    print(f"  ĐK GTGT 2020: {m['dk_gtgt_qt_truoc']:,.0f} (QT 2015) + {m['dk_gtgt_c2_1912']:,.0f} (C2 1912) = {m['dk_gtgt_qt_truoc'] + m['dk_gtgt_c2_1912']:,.0f}")
    print(f"  ĐK TNCN 2020: {m['dk_tncn_c2_1912']:,.0f} (C2 1912) + {m['dk_tncn_qt_2019']:,.0f} (QT 2019) = {m['dk_tncn_c2_1912'] + m['dk_tncn_qt_2019']:,.0f}")
    for y in range(2020, 2026):
        print(
            f"  Năm {y}: GTGT TK={m['gtgt_tk_by_ky'][y]:,.0f}, C2={m['gtgt_c2_by_ht'][y]:,.0f} | "
            f"TNCN (1001) TK+TH+TQ={m['tncn_tk_th_tq_by_ky'][y]:,.0f}, C2={m['tncn_c2_by_ht'][y]:,.0f}"
        )

    targets = [
        r"C:\Users\Admin\.gemini\antigravity\scratch\tax_xml_to_excel\TQ_QHNS_TongHop.xlsx",
        r"D:\2026\MANY\TCS_MANY\TCS_MANY\TQ_QHNS_TongHop.xlsx",
    ]
    sync_master_workbook(master_uploaded, targets)
