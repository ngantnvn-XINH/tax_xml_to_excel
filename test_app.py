# -*- coding: utf-8 -*-
"""Kiểm thử tự động parse XML tờ khai GTGT, TNCN, TNDN và xuất file Excel."""

import datetime
import os
import tempfile
import unittest

import openpyxl

from excel_exporter import export_to_excel
from tax_parser import TaxDeclaration, parse_tax_xml


SAMPLE_GTGT = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>01/GTGT</maTKhai>
          <tenTKhai>Tờ khai thuế giá trị gia tăng (TT80/2021)</tenTKhai>
          <KyKKhaiThue>
            <kyKKhai>Q1/2025</kyKKhai>
            <kyKKhaiTuNgay>2025-01-01</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>2025-03-31</kyKKhaiDenNgay>
          </KyKKhaiThue>
          <ngayLapTKhai>2025-04-20</ngayLapTKhai>
          <soLan>0</soLan>
          <tenCQTNoiNop>Chi cục Thuế Quận 1</tenCQTNoiNop>
        </TKhaiThue>
        <NNT>
          <mst>0316443262</mst>
          <tenNNT>CÔNG TY TNHH ABC</tenNNT>
          <dchiNNT>TP HCM</dchiNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <ct22>15000000</ct22>
      <ct23>100000000</ct23>
      <ct24>10000000</ct24>
      <ct25>10000000</ct25>
      <DoanhThuHHDVBanRa>
        <ct32>200000000</ct32>
        <ct33>20000000</ct33>
      </DoanhThuHHDVBanRa>
      <ct43>5000000</ct43>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>
"""

SAMPLE_TNCN = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>05/QTT-TNCN</maTKhai>
          <tenTKhai>Tờ khai quyết toán thuế thu nhập cá nhân</tenTKhai>
          <KyKKhaiThue>
            <kyKKhai>2024</kyKKhai>
            <kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay>
          </KyKKhaiThue>
          <ngayLapTKhai>2025-03-25</ngayLapTKhai>
          <soLan>1</soLan>
          <tenCQTNoiNop>Cục Thuế TP Hà Nội</tenCQTNoiNop>
        </TKhaiThue>
        <NNT>
          <mst>0100109106</mst>
          <tenNNT>CÔNG TY CỔ PHẦN XYZ</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <ct21>12</ct21>
      <ct26>2400000000</ct26>
      <ct32>120000000</ct32>
    </CTieuTKhaiChinh>
    <PLuc>
      <PL_05_1BK_QTT_TNCN>
        <BKeCTietCNhan>
          <CTiet>
            <hoVaTen>Nguyễn Văn A</hoVaTen>
            <mst>0123456789</mst>
            <tongTNCT>300000000</tongTNCT>
            <thueDaKhauTru>15000000</thueDaKhauTru>
          </CTiet>
          <CTiet>
            <hoVaTen>Trần Thị B</hoVaTen>
            <mst>0987654321</mst>
            <tongTNCT>250000000</tongTNCT>
            <thueDaKhauTru>10000000</thueDaKhauTru>
          </CTiet>
        </BKeCTietCNhan>
      </PL_05_1BK_QTT_TNCN>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>
"""

SAMPLE_TNDN = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>03/TNDN</maTKhai>
          <tenTKhai>Tờ khai quyết toán thuế thu nhập doanh nghiệp</tenTKhai>
          <KyKKhaiThue>
            <kyKKhai>2024</kyKKhai>
            <kyKKhaiTuNgay>2024-01-01</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>2024-12-31</kyKKhaiDenNgay>
          </KyKKhaiThue>
          <ngayLapTKhai>2025-03-30</ngayLapTKhai>
          <soLan>0</soLan>
          <tenCQTNoiNop>Cục Thuế TP HCM</tenCQTNoiNop>
        </TKhaiThue>
        <NNT>
          <mst>0300908346-016</mst>
          <tenNNT>CHI NHÁNH CÔNG TY DEF</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <A1>500000000</A1>
      <C1>500000000</C1>
      <D>100000000</D>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>
"""


class TestTaxXmlToExcel(unittest.TestCase):
    def test_parse_and_export(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            f_gtgt = os.path.join(tmpdir, "gtgt.xml")
            f_tncn = os.path.join(tmpdir, "tncn.xml")
            f_tndn = os.path.join(tmpdir, "tndn.xml")
            out_xlsx = os.path.join(tmpdir, "ket_qua.xlsx")

            for path, content in ((f_gtgt, SAMPLE_GTGT), (f_tncn, SAMPLE_TNCN), (f_tndn, SAMPLE_TNDN)):
                with open(path, "w", encoding="utf-8") as f:
                    f.write(content)

            decls = [parse_tax_xml(f_gtgt), parse_tax_xml(f_tncn), parse_tax_xml(f_tndn)]
            self.assertEqual([d.tax_type for d in decls], ["GTGT", "TNCN", "TNDN"])
            self.assertEqual(decls[0].mst, "0316443262")
            self.assertEqual(decls[0].indicators["ct32"], 200000000)
            self.assertEqual(len(decls[1].appendices["PL_05_1BK_QTT_TNCN"]), 2)

            created = export_to_excel(
                decls, out_xlsx, include_appendices=True, split_by_tax_type=False, layout="vertical"
            )
            self.assertEqual(len(created), 1)
            self.assertTrue(os.path.exists(out_xlsx))

            wb = openpyxl.load_workbook(out_xlsx)
            self.assertIn("Tong_hop_File", wb.sheetnames)
            self.assertIn("KKTNDN", wb.sheetnames)
            self.assertIn("BCHDKD", wb.sheetnames)
            self.assertIn("TNCN", wb.sheetnames)
            self.assertIn("So Thue GTGT TH", wb.sheetnames)
            self.assertIn("So Thue GTGT Cuoi", wb.sheetnames)

            # File tổng hợp gom TNCN vào đúng 1 sheet, không có các sheet phụ lục rời
            self.assertNotIn("TK_2024", wb.sheetnames)
            self.assertNotIn("05-1_2024", wb.sheetnames)

            ws_tncn = wb["TNCN"]
            self.assertEqual(ws_tncn["G1"].value, 2024)
            self.assertEqual(ws_tncn["G8"].value, 12)  # ct21

            ws_tndn = wb["KKTNDN"]
            # Hàng 1 là Năm, hàng 2 là tiêu đề cột, hàng 5 là chỉ tiêu A1
            self.assertEqual(ws_tndn["B5"].value, "1")
            self.assertEqual(ws_tndn["D5"].value, "A1")
            self.assertEqual(ws_tndn["E5"].value, 500000000)

            # Kiểm tra sheet GTGT timeline
            ws_gtgt = wb["So Thue GTGT TH"]
            self.assertEqual(ws_gtgt["G3"].value, "2501")
            self.assertEqual(ws_gtgt["H3"].value, 0)
            self.assertEqual(ws_gtgt["L3"].value, 15000000)  # ct22
            self.assertEqual(ws_gtgt["N3"].value, 10000000)  # ct24

            # Kiểm thử thêm bố cục chiều ngang
            out_horiz = os.path.join(tmpdir, "ket_qua_ngang.xlsx")
            export_to_excel(decls, out_horiz, include_appendices=False, split_by_tax_type=False, layout="horizontal")
            wb_h = openpyxl.load_workbook(out_horiz)
            ws_h = wb_h["So Thue GTGT TH"]
            self.assertEqual(ws_h["E3"].value, "CÔNG TY TNHH ABC")


            # Kiểm thử xuất file tổng hợp gồm cả TT156 (2020) và TT80 (2024) vào chung 1 sheet TNCN
            d2020 = TaxDeclaration(file_path="", file_name="2020.xml", tax_type="TNCN", ky_kkhai="2020", indicators={"21": 70, "22": 70, "23": 42, "28": 10761646159, "43": 1195957511})
            out_master_tncn = os.path.join(tmpdir, "TongHop_TNCN_HaiThongTu.xlsx")
            export_to_excel([d2020, decls[1]], out_master_tncn, include_appendices=False, split_by_tax_type=False)
            wb_master_tncn = openpyxl.load_workbook(out_master_tncn)
            self.assertEqual(wb_master_tncn.sheetnames, ["Tong_hop_File", "TNCN", "2020", "2024", "Tinh_TienChamNop", "BaoCao_BienBan"])
            ws_master = wb_master_tncn["TNCN"]
            self.assertEqual(ws_master["G1"].value, 2020)
            self.assertEqual(ws_master["G3"].value, 70)
            self.assertEqual(ws_master["M1"].value, 2024)
            self.assertEqual(ws_master["M8"].value, 12)

            # Kiểm thử xuất riêng theo từng năm (có phụ lục 05-1, 05-2, 05-3)
            out_year_dir = os.path.join(tmpdir, "TachNam.xlsx")
            created_year_files = export_to_excel(decls, out_year_dir, split_by_year=True)
            tncn_files = [f for f in created_year_files if "QTT_TNCN" in f]
            self.assertTrue(len(tncn_files) > 0)
            wb_tncn = openpyxl.load_workbook(tncn_files[0])
            self.assertEqual(wb_tncn.sheetnames, ["TK", "05-1", "05-2", "05-3"])

    def test_gtgt_deduplication(self) -> None:
        """Kiểm thử tính năng tự động loại trừ tờ khai GTGT trùng lặp do có nhiều file cùng 1 tờ khai."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Giả lập 2 file 2508 lần 0 và 2 file 2508 lần 1 (tổng cộng 4 file)
            d1_ct0 = TaxDeclaration(
                file_path="f1.xml", file_name="0100000001_01GTGT_25M08_CT_0.xml", tax_type="GTGT",
                mst="0100000001", ky_kkhai="08/2025", so_lan="0", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 9, 17), indicators={"ct40": 374974620}
            )
            d2_ct0 = TaxDeclaration(
                file_path="f2.xml", file_name="0100000001_08-2025_01-GTGT_CT0_1_21420250009848413.xml", tax_type="GTGT",
                mst="0100000001", ky_kkhai="08/2025", so_lan="0", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 9, 17), indicators={"ct40": 374974620}
            )
            d3_bs1 = TaxDeclaration(
                file_path="f3.xml", file_name="0100000001_01GTGT_25M08_BS_1.xml", tax_type="GTGT",
                mst="0100000001", ky_kkhai="08/2025", so_lan="1", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 10, 14), indicators={"ct40": 365632924}
            )
            d4_bs1 = TaxDeclaration(
                file_path="f4.xml", file_name="0100000001_08-2025_01-GTGT_BS1_1_21420250009862365.xml", tax_type="GTGT",
                mst="0100000001", ky_kkhai="08/2025", so_lan="1", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 10, 14), indicators={"ct40": 365632924}
            )

            out_xlsx = os.path.join(tmpdir, "gtgt_dedup.xlsx")
            export_to_excel([d1_ct0, d2_ct0, d3_bs1, d4_bs1], out_xlsx)

            wb = openpyxl.load_workbook(out_xlsx)
            self.assertIn("So Thue GTGT TH", wb.sheetnames)
            self.assertIn("So Thue GTGT Cuoi", wb.sheetnames)

            # Sheet TH: gồm cả lần 0 và lần 1 (mỗi lần đúng 1 dòng)
            ws_gtgt_th = wb["So Thue GTGT TH"]
            data_rows_th = [r for r in range(3, ws_gtgt_th.max_row + 1) if ws_gtgt_th.cell(r, 7).value is not None]
            self.assertEqual(len(data_rows_th), 2)
            self.assertEqual(ws_gtgt_th.cell(3, 7).value, "2508")
            self.assertEqual(ws_gtgt_th.cell(3, 8).value, 0)
            self.assertEqual(ws_gtgt_th.cell(4, 7).value, "2508")
            self.assertEqual(ws_gtgt_th.cell(4, 8).value, 1)

            # Sheet Cuoi: kỳ 2508 chỉ lấy tờ khai cuối (lần 1) và có dòng Tổng cộng năm 2025
            ws_gtgt_cuoi = wb["So Thue GTGT Cuoi"]
            self.assertEqual(ws_gtgt_cuoi.cell(3, 7).value, "2508")
            self.assertEqual(ws_gtgt_cuoi.cell(3, 8).value, 1)  # Lần 1
            self.assertEqual(ws_gtgt_cuoi.cell(4, 5).value, "Tổng cộng năm 2025")

            # Sheet Tong_hop_File phải hiển thị thông tin trùng minh bạch
            ws_summary = wb["Tong_hop_File"]
            statuses = [ws_summary.cell(r, 10).value for r in range(2, 6)]
            self.assertEqual(statuses.count("Thành công"), 2)
            self.assertEqual(statuses.count("Thành công (Trùng tờ khai)"), 2)

    def test_bctc_sodaunam_and_borders(self) -> None:
        """Kiểm thử trích xuất đầy đủ Số cuối năm và Số đầu năm BCTC kèm đóng khung viền."""
        with tempfile.TemporaryDirectory() as tmpdir:
            dBCTC = TaxDeclaration(
                file_path="bctc2020.xml", file_name="bctc2020.xml", tax_type="BCTC",
                mst="0100000001", ky_kkhai="2020",
                indicators={
                    "CDKT_HoatDongLienTuc.SoCuoiNam.ct100": 41557960607,
                    "CDKT_HoatDongLienTuc.SoDauNam.ct100": 88318848856,
                    "CDKT_HoatDongKhongLienTuc.SoCuoiNam.ct100": 0,
                    "CDKT_HoatDongKhongLienTuc.SoDauNam.ct100": 0,
                }
            )
            out_xlsx = os.path.join(tmpdir, "bctc_test.xlsx")
            export_to_excel([dBCTC], out_xlsx)

            wb = openpyxl.load_workbook(out_xlsx)
            ws_bctc = wb["BCTC"]
            # Kiểm tra hàng 6 (Mã 100)
            val_cuoi = ws_bctc.cell(6, 4).value
            val_dau = ws_bctc.cell(6, 5).value
            self.assertEqual(val_cuoi, 41557960607)
            self.assertEqual(val_dau, 88318848856)

            # Kiểm tra đóng khung tất cả các ô có dữ liệu
            for r in range(1, ws_bctc.max_row + 1):
                for c in range(1, ws_bctc.max_column + 1):
                    cell = ws_bctc.cell(r, c)
                    if cell.value is not None:
                        b = cell.border
                        self.assertIsNotNone(b)
                        self.assertIsNotNone(b.left.style)
                        self.assertIsNotNone(b.right.style)
                        self.assertIsNotNone(b.top.style)
                        self.assertIsNotNone(b.bottom.style)

    def test_chenh_lech_gtgt_tndn(self) -> None:
        """Kiểm thử tự động tạo sheet 'Chenh lech thue GTGT - TNDN' với liên kết công thức và đóng khung."""
        with tempfile.TemporaryDirectory() as tmpdir:
            dGTGT = TaxDeclaration(
                file_path="gtgt2024.xml", file_name="gtgt2024.xml", tax_type="GTGT",
                mst="0100000001", ky_kkhai="12/2024", so_lan="0",
                tu_ngay=datetime.date(2024, 12, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 1, 20),
                indicators={"ct34": 50000000000, "ct40": 5000000000}
            )
            dTNDN = TaxDeclaration(
                file_path="tndn2024.xml", file_name="tndn2024.xml", tax_type="TNDN",
                mst="0100000001", ky_kkhai="2024", so_lan="0",
                tu_ngay=datetime.date(2024, 1, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 3, 30),
                indicators={"ct04": 48000000000, "ct06": 1000000000, "ct19": 200000000}
            )
            out_xlsx = os.path.join(tmpdir, "chenh_lech_test.xlsx")
            export_to_excel([dGTGT, dTNDN], out_xlsx)

            wb = openpyxl.load_workbook(out_xlsx, data_only=False)
            self.assertIn("Chenh lech thue GTGT - TNDN", wb.sheetnames)
            ws_cl = wb["Chenh lech thue GTGT - TNDN"]

            # Row 1: Chỉ số
            expected_r1 = [None, 1, 2, 3, None, 4, 5, "6=2-3", "7=6+4", "8=1-7"]
            self.assertEqual([ws_cl.cell(1, c).value for c in range(1, 11)], expected_r1)

            # Row 2: Tiêu đề
            self.assertEqual(ws_cl.cell(2, 1).value, "Năm")
            self.assertEqual(ws_cl.cell(2, 2).value, "DT tính thuế GTGT")
            self.assertEqual(ws_cl.cell(2, 3).value, "DT tính thuế TNDN")
            self.assertEqual(ws_cl.cell(2, 4).value, "Giảm trừ doanh thu")
            self.assertEqual(ws_cl.cell(2, 5).value, "Chênh lệch chưa bao gồm thu nhập khác")
            self.assertEqual(ws_cl.cell(2, 6).value, "Thu nhập khác")
            self.assertEqual(ws_cl.cell(2, 7).value, "Chênh lệch")
            self.assertEqual(ws_cl.cell(2, 8).value, "DT tính thuế TNDN-Giảm trừ doanh thu")
            self.assertEqual(ws_cl.cell(2, 9).value, "(DT tính thuế TNDN-Giảm trừ doanh thu)+Thu nhập khác")
            self.assertEqual(ws_cl.cell(2, 10).value, "Chênh lệch")

            # Row 3: Năm 2024
            self.assertEqual(ws_cl.cell(3, 1).value, 2024)
            self.assertIn("'So Thue GTGT Cuoi'!AA", str(ws_cl.cell(3, 2).value))
            self.assertIn("'BCHDKD'!", str(ws_cl.cell(3, 3).value))
            self.assertEqual(ws_cl.cell(3, 5).value, "=B3-C3")
            self.assertEqual(ws_cl.cell(3, 7).value, "=B3-C3-F3+D3")
            self.assertEqual(ws_cl.cell(3, 8).value, "=C3-D3")
            self.assertEqual(ws_cl.cell(3, 9).value, "=H3+F3")
            self.assertEqual(ws_cl.cell(3, 10).value, "=B3-I3")

            # Row 4: Tổng cộng
            self.assertEqual(ws_cl.cell(4, 1).value, "Tổng cộng")
            self.assertEqual(ws_cl.cell(4, 2).value, "=SUM(B3:B3)")
            self.assertEqual(ws_cl.cell(4, 10).value, "=SUM(J3:J3)")

            # Kiểm tra đóng khung toàn bộ bảng (Row 2 đến Row 4, Col 1 đến 10)
            for r in range(2, 5):
                for c in range(1, 11):
                    cell = ws_cl.cell(r, c)
                    self.assertIsNotNone(cell.border.left.style)
                    self.assertIsNotNone(cell.border.right.style)
                    self.assertIsNotNone(cell.border.top.style)
                    self.assertIsNotNone(cell.border.bottom.style)

    def test_v12_tndn_tt156_and_tt80_no_shift_and_kqkd_isolation(self) -> None:
        """Kiểm thử v1.2: Không lệch dòng KKTNDN giữa TT156/TT80, hiển thị đủ dòng tổng D/E/G/H/I và tách biệt BCHDKD."""
        xml_tt156 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>03</maTKhai><tenTKhai>Tờ khai quyết toán thuế TNDN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2020</kyKKhai><kyKKhaiTuNgay>01/01/2020</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2020</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2021-03-30</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh>
      <ctA1>-24131629971</ctA1>
      <ctB1>2254017275</ctB1>
      <ctB12>-21877612696</ctB12>
      <ctB13>-21877612696</ctB13>
      <ctB14>0</ctB14>
      <ctC1>-21877612696</ctC1>
      <ctD>0</ctD><ctE>0</ctE><ctG>0</ctG><ctH>0</ctH><ctI>0</ctI>
    </CTieuTKhaiChinh>
    <PLuc>
      <PLuc_03_1A_TNDN>
        <ct01>56037733621</ct01>
        <ct04>56037733621</ct04>
        <ct19>-24131629971</ct19>
      </PLuc_03_1A_TNDN>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>"""

        xml_tt80 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>892</maTKhai><tenTKhai>Tờ khai quyết toán thuế TNDN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2024</kyKKhai><kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2025-03-31</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh>
      <ct05>100</ct05>
      <ctA1>5009403788</ctA1>
      <ctB13>7097859895</ctB13>
      <ctD>569629519</ctD>
      <ctE>468740238</ctE>
      <ctG>569629519</ctG>
      <ctH>468740238</ctH>
      <ctI>100889281</ctI>
    </CTieuTKhaiChinh>
    <PLuc>
      <PL03_1A_TNDN>
        <ct04>52341187939</ct04>
        <ct05>0</ct05>
        <ct22>5009403788</ct22>
      </PL03_1A_TNDN>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>"""

        with tempfile.TemporaryDirectory() as tmpdir:
            f1 = os.path.join(tmpdir, "tndn2020.xml")
            f2 = os.path.join(tmpdir, "tndn2024.xml")
            out_xlsx = os.path.join(tmpdir, "tndn_v12.xlsx")
            for p, content in ((f1, xml_tt156), (f2, xml_tt80)):
                with open(p, "w", encoding="utf-8") as f:
                    f.write(content)

            decls = [parse_tax_xml(f1), parse_tax_xml(f2)]
            export_to_excel(decls, out_xlsx)
            wb = openpyxl.load_workbook(out_xlsx)

            ws_kk = wb["KKTNDN"]
            # Kiểm tra TT156: dòng 18 là B12 (3. Tổng thu nhập chịu thuế), dòng 19 là B13 (3.1)
            self.assertEqual(ws_kk.cell(18, 2).value, "3")
            self.assertEqual(ws_kk.cell(18, 4).value, "B12")
            self.assertEqual(ws_kk.cell(18, 5).value, -21877612696)
            self.assertEqual(ws_kk.cell(19, 2).value, "3.1")
            self.assertEqual(ws_kk.cell(19, 4).value, "B13")
            self.assertEqual(ws_kk.cell(19, 5).value, -21877612696)

            # Kiểm tra TT80: các dòng tổng D, E, G, H, I có giá trị (cột 9 là Mã chỉ tiêu, cột 10 là Số tiền)
            tt80_codes = {ws_kk.cell(r, 9).value: ws_kk.cell(r, 10).value for r in range(4, ws_kk.max_row + 1)}
            self.assertEqual(tt80_codes.get("D"), 569629519)
            self.assertEqual(tt80_codes.get("E"), 468740238)
            self.assertEqual(tt80_codes.get("G"), 569629519)
            self.assertEqual(tt80_codes.get("H"), 468740238)
            self.assertEqual(tt80_codes.get("I"), 100889281)

            # Kiểm tra BCHDKD: chỉ tiêu [05] của TT80 phải bằng 0 (không bị ghi đè bởi ct05=100 từ tờ khai chính)
            ws_kq = wb["BCHDKD"]
            self.assertEqual(ws_kq.cell(6, 9).value, "[05]")
            self.assertEqual(ws_kq.cell(6, 10).value, 0)

    def test_v12_tncn_appendix_inheritance_and_columns(self) -> None:
        """Kiểm thử v1.2: Kế thừa phụ lục 05-3 từ tờ khai trước khi tờ khai bổ sung chỉ gửi 05-1/05-2."""
        xml_ct0 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>953</maTKhai><tenTKhai>Tờ khai quyết toán thuế TNCN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2023</kyKKhai><kyKKhaiTuNgay>01/01/2023</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2023</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2024-03-30</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh><ct21>10</ct21><ct22>1</ct22></CTieuTKhaiChinh>
    <PLuc>
      <PLuc_05_3_BK_QTT>
        <BKeNPT><ThongTinNPT id="ID_1">
          <ct07>Dương Thủy Nguyên</ct07><ct08>8348315370</ct08><ct09>Lý Thị Loan</ct09>
          <ct10>1960-09-11</ct10><ct11>8867960913</ct11><ct12_ten>Thẻ CCCD</ct12_ten>
          <ct13>094160009275</ct13><ct14_ten>Cha/mẹ</ct14_ten>
          <ThoiGianGiamTru><ct15>02/2023</ct15><ct16>12/2023</ct16></ThoiGianGiamTru>
        </ThongTinNPT></BKeNPT>
      </PLuc_05_3_BK_QTT>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>"""

        xml_bs1 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>953</maTKhai><tenTKhai>Tờ khai quyết toán thuế TNCN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2023</kyKKhai><kyKKhaiTuNgay>01/01/2023</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2023</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2024-06-15</ngayLapTKhai><soLan>1</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh><ct21>10</ct21><ct22>1</ct22></CTieuTKhaiChinh>
    <PLuc>
      <PLuc_05_1_BK_QTT>
        <BCaoCTietCNhan><CTietCNhan id="ID_1">
          <ct07>Dương Thủy Nguyên</ct07><ct08>8348315370</ct08><ct12>100000000</ct12><ct15>1</ct15>
        </CTietCNhan></BCaoCTietCNhan>
      </PLuc_05_1_BK_QTT>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>"""

        with tempfile.TemporaryDirectory() as tmpdir:
            f0 = os.path.join(tmpdir, "tncn_ct0.xml")
            f1 = os.path.join(tmpdir, "tncn_bs1.xml")
            out_dir = os.path.join(tmpdir, "out.xlsx")
            for p, content in ((f0, xml_ct0), (f1, xml_bs1)):
                with open(p, "w", encoding="utf-8") as f:
                    f.write(content)

            created = export_to_excel([parse_tax_xml(f0), parse_tax_xml(f1)], out_dir, split_by_year=True)
            wb = openpyxl.load_workbook(created[0])
            ws_053 = wb["05-3"]
            self.assertEqual(ws_053.max_row, 4)
            row_vals = [ws_053.cell(4, c).value for c in range(1, 12)]
            self.assertEqual(
                row_vals,
                [1, "Dương Thủy Nguyên", "8348315370", "Lý Thị Loan", "1960-09-11", "8867960913", "Thẻ CCCD", "094160009275", "Cha/mẹ", "02/2023", "12/2023"],
            )

            # Kiểm tra bảng dữ liệu tổng hợp thứ 2 trên sheet TK (dòng 38..47) lấy từ sheet 05-1 và 05-2
            ws_tk = wb["TK"]
            self.assertEqual(ws_tk.cell(38, 2).value, "STT")
            self.assertEqual(ws_tk.cell(38, 3).value, "Chỉ tiêu")
            self.assertEqual(ws_tk.cell(38, 4).value, "Số báo cáo của người nộp thuế")
            self.assertEqual(ws_tk.cell(38, 5).value, "Số kiểm tra")
            self.assertEqual(ws_tk.cell(38, 6).value, "Chênh lệch")
            self.assertEqual(ws_tk.cell(40, 2).value, "I")
            self.assertEqual(ws_tk.cell(41, 4).value, "='05-1'!H5")
            self.assertEqual(ws_tk.cell(42, 4).value, "='05-1'!R5-'05-1'!U5+'05-1'!V5")
            self.assertEqual(ws_tk.cell(43, 4).value, "='05-2'!G5")
            self.assertEqual(ws_tk.cell(44, 4).value, "='05-2'!K5")
            self.assertEqual(ws_tk.cell(45, 4).value, "='05-2'!G6")
            self.assertEqual(ws_tk.cell(46, 4).value, "='05-2'!K6")
            self.assertEqual(ws_tk.cell(47, 2).value, "Tổng hợp tổng số thuế TNCN phải khấu trừ")
            self.assertEqual(ws_tk.cell(47, 4).value, "=D42+D44+D46")

    def test_v12_bctc_subcodes_and_other_declarations(self) -> None:
        """Kiểm thử v1.2: BCTC mã con 411a/411b không ghi đè mã cha 411 và xuất tờ khai Khác (01/LPMB)."""
        xml_bctc = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>998</maTKhai><tenTKhai>Bộ báo cáo tài chính</tenTKhai>
        <KyKKhaiThue><kyKKhai>2024</kyKKhai><kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2025-03-30</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh>
      <CDKT_HoatDongLienTuc>
        <SoCuoiNam><ct411>85500000000</ct411><ct411a>85500000000</ct411a><ct411b>0</ct411b></SoCuoiNam>
        <SoDauNam><ct411>85500000000</ct411><ct411a>85500000000</ct411a><ct411b>0</ct411b></SoDauNam>
      </CDKT_HoatDongLienTuc>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        xml_lpmb = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>988</maTKhai><tenTKhai>Tờ khai lệ phí môn bài (01/LPMB)</tenTKhai>
        <KyKKhaiThue><kyKKhai>2023</kyKKhai><kyKKhaiTuNgay>01/01/2023</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2023</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2023-01-15</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh><ct06>3000000</ct06></CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        with tempfile.TemporaryDirectory() as tmpdir:
            fb = os.path.join(tmpdir, "bctc.xml")
            fl = os.path.join(tmpdir, "lpmb.xml")
            out_xlsx = os.path.join(tmpdir, "bctc_lpmb.xlsx")
            for p, content in ((fb, xml_bctc), (fl, xml_lpmb)):
                with open(p, "w", encoding="utf-8") as f:
                    f.write(content)

            export_to_excel([parse_tax_xml(fb), parse_tax_xml(fl)], out_xlsx)
            wb = openpyxl.load_workbook(out_xlsx)
            self.assertIn("BCTC", wb.sheetnames)
            self.assertIn("To_Khai_Khac", wb.sheetnames)

            ws_bctc = wb["BCTC"]
            bctc_map = {str(ws_bctc.cell(r, 2).value): ws_bctc.cell(r, 4).value for r in range(6, ws_bctc.max_row + 1)}
            self.assertEqual(bctc_map.get("411"), 85500000000)
            self.assertEqual(bctc_map.get("411a"), 85500000000)
            self.assertEqual(bctc_map.get("411b"), 0)

    def test_yearly_verification_sheets_98_rows(self) -> None:
        """Kiểm thử sheet kiểm tra theo từng năm (mỗi năm 1 sheet, 98 dòng x 5 cột, công thức liên kết chuẩn)."""
        with tempfile.TemporaryDirectory() as tmpdir:
            dGTGT = TaxDeclaration(
                file_path="gtgt_2024_12.xml", file_name="gtgt_2024_12.xml",
                ma_tkhai="842", ten_tkhai="Tờ khai thuế GTGT", tax_type="GTGT",
                ky_kkhai="12/2024", so_lan="0",
                mst="0100000001", ten_nnt="CTY A",
                tu_ngay=datetime.date(2024, 12, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 1, 20),
                indicators={"ct22": 1000000, "ct32": 50000000, "ct33": 5000000, "ct43": 2000000}
            )
            dTNDN = TaxDeclaration(
                file_path="tndn_2024.xml", file_name="tndn_2024.xml",
                ma_tkhai="892", ten_tkhai="Tờ khai QTT TNDN", tax_type="TNDN",
                ky_kkhai="2024", so_lan="0",
                mst="0100000001", ten_nnt="CTY A",
                tu_ngay=datetime.date(2024, 1, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 3, 30),
                indicators={"ct04": 50000000, "ctB4": 3000000}
            )
            dTNCN = TaxDeclaration(
                file_path="tncn_2024.xml", file_name="tncn_2024.xml",
                ma_tkhai="864", ten_tkhai="Tờ khai QTT TNCN", tax_type="TNCN",
                ky_kkhai="2024", so_lan="0",
                mst="0100000001", ten_nnt="CTY A",
                tu_ngay=datetime.date(2024, 1, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 3, 30),
                indicators={"ct21": 5},
                appendices={
                    "PL05_1_BK_QTT_TNCN": [{"ct11": 100000000, "ct19": 5000000}],
                    "PL05_2_BK_QTT_TNCN": [{"ct10": False, "ct11": 20000000, "ct14": 2000000}],
                }
            )
            out_xlsx = os.path.join(tmpdir, "yearly_check.xlsx")
            export_to_excel([dGTGT, dTNDN, dTNCN], out_xlsx)
            wb = openpyxl.load_workbook(out_xlsx, data_only=False)
            self.assertIn("2024", wb.sheetnames)
            ws = wb["2024"]
            self.assertEqual(ws.max_row, 98)
            self.assertEqual(ws.max_column, 5)
            self.assertIn("'So Thue GTGT Cuoi'!L", str(ws.cell(3, 3).value))
            self.assertIn("'So Thue GTGT Cuoi'!X", str(ws.cell(9, 3).value))
            self.assertIn("'BCHDKD'!", str(ws.cell(30, 3).value))
            self.assertIn("'KKTNDN'!", str(ws.cell(52, 3).value))
            self.assertEqual(ws.cell(84, 3).value, "='TNCN'!G41")
            self.assertEqual(ws.cell(85, 3).value, "='TNCN'!G42")
            self.assertEqual(ws.cell(90, 3).value, "=C85+C87+C89")

    def test_qhns_opening_exact_c5_refund_and_bt_exclusion(self) -> None:
        """Kiểm thử tính số dư đầu kỳ tuyệt đối (Bucket C), hoàn nộp thừa C5, và loại bỏ bút toán BT."""
        from qhns_summarizer import compute_tq_metrics

        tq_rows = [
            # GTGT (1701): nộp thừa kỳ trước -49,274,310 + QT 15CN nộp năm 2021 (+64,511,180) + TK 1912 nộp 2020 (+847,949,684)
            {"ky": "18Q4", "ky_year": 2018, "gd": "C2", "tm": "1701", "so_tien": -49274310.0, "ht_year": 2019, "no_co": ""},
            {"ky": "1912", "ky_year": 2019, "gd": "TK", "tm": "1701", "so_tien": 847949684.0, "ht_year": 2020, "no_co": ""},
            {"ky": "1912", "ky_year": 2019, "gd": "C2", "tm": "1701", "so_tien": -847949684.0, "ht_year": 2020, "no_co": ""},
            {"ky": "15CN", "ky_year": 2015, "gd": "QT", "tm": "1701", "so_tien": 64511180.0, "ht_year": 2021, "no_co": ""},
            {"ky": "2001", "ky_year": 2020, "gd": "TK", "tm": "1701", "so_tien": 100000000.0, "ht_year": 2020, "no_co": ""},
            {"ky": "2001", "ky_year": 2020, "gd": "C2", "tm": "1701", "so_tien": -50000000.0, "ht_year": 2020, "no_co": ""},
            # TNDN (1052): nộp thừa 2018 (-12,081,620) và 2019 (-100,000,000), hoàn nộp thừa C5 năm 2024 (+100,000,000)
            {"ky": "1811", "ky_year": 2018, "gd": "C2", "tm": "1052", "so_tien": -12081620.0, "ht_year": 2018, "no_co": ""},
            {"ky": "1812", "ky_year": 2018, "gd": "C2", "tm": "1052", "so_tien": -100000000.0, "ht_year": 2019, "no_co": ""},
            {"ky": "2411", "ky_year": 2024, "gd": "PT", "tm": "1052", "so_tien": -100000000.0, "ht_year": 2024, "no_co": ""},
            {"ky": "2411", "ky_year": 2024, "gd": "QH", "tm": "1052", "so_tien": 100000000.0, "ht_year": 2024, "no_co": ""},
            {"ky": "2411", "ky_year": 2024, "gd": "C5", "tm": "1052", "so_tien": 100000000.0, "ht_year": 2024, "no_co": ""},
            # Bút toán BT nội bộ phải bị bỏ qua
            {"ky": "20CN", "ky_year": 2020, "gd": "BT", "tm": "1052", "so_tien": -1000000.0, "ht_year": 2021, "no_co": ""},
            {"ky": "21CN", "ky_year": 2021, "gd": "BT", "tm": "1052", "so_tien": 1000000.0, "ht_year": 2021, "no_co": ""},
        ]
        m = compute_tq_metrics(tq_rows, 2020, 2024)
        # GTGT đầu kỳ = -49,274,310 + 847,949,684 + 64,511,180 = 863,186,554
        self.assertEqual(round(m["opening_exact"]["1701"]), 863186554)
        # Dư cuối GTGT = 863,186,554 + 100,000,000 - (847,949,684 + 50,000,000) = 65,236,870 (khớp tổng TQ!)
        gtgt_close = round(m["opening_exact"]["1701"] + sum(m["gtgt_tk_by_ky"].values()) - sum(m["gtgt_c2_by_ht"].values()))
        self.assertEqual(gtgt_close, 65236870)
        # TNDN đầu kỳ = -112,081,620; năm 2024 hoàn C5 -> số nộp = -100,000,000; dư cuối = -12,081,620
        self.assertEqual(round(m["opening_exact"]["1052"]), -112081620)
        self.assertEqual(round(m["tndn_c2_by_ht"][2024]), -100000000)
        tndn_close = round(m["opening_exact"]["1052"] + sum(m["tndn_tk_qt_by_ky"].values()) - sum(m["tndn_c2_by_ht"].values()))
        self.assertEqual(tndn_close, -12081620)

    def test_qhns_lpmb_reconciliation(self) -> None:
        """Kiểm thử đối chiếu Quan hệ ngân sách cho Lệ phí môn bài (2862, 2863, 2864)."""
        from qhns_summarizer import compute_tq_metrics

        tq_rows = [
            # Kỳ trước (2019) hạch toán muộn trong 2024 và nộp năm 2024
            {"ky": "19CN", "ky_year": 2019, "gd": "TK", "tm": "2862", "so_tien": 3000000.0, "ht_year": 2024, "no_co": ""},
            {"ky": "19CN", "ky_year": 2019, "gd": "C2", "tm": "2862", "so_tien": -3000000.0, "ht_year": 2024, "no_co": ""},
            # Năm 2023: Bậc 1 (2862: 3tr) + Chi nhánh (2864: 1tr), đã nộp đủ 4tr
            {"ky": "23CN", "ky_year": 2023, "gd": "TK", "tm": "2862", "so_tien": 3000000.0, "ht_year": 2023, "no_co": ""},
            {"ky": "23CN", "ky_year": 2023, "gd": "TK", "tm": "2864", "so_tien": 1000000.0, "ht_year": 2023, "no_co": ""},
            {"ky": "23CN", "ky_year": 2023, "gd": "C2", "tm": "2862", "so_tien": -3000000.0, "ht_year": 2023, "no_co": ""},
            {"ky": "23CN", "ky_year": 2023, "gd": "C2", "tm": "2864", "so_tien": -1000000.0, "ht_year": 2023, "no_co": ""},
            # Năm 2024: Bậc 2 (2863: 2tr), nộp nhầm 1tr rồi hủy/điều chỉnh +1tr -> còn nợ 2tr
            {"ky": "24CN", "ky_year": 2024, "gd": "TK", "tm": "2863", "so_tien": 2000000.0, "ht_year": 2024, "no_co": ""},
            {"ky": "24CN", "ky_year": 2024, "gd": "C2", "tm": "2863", "so_tien": -1000000.0, "ht_year": 2024, "no_co": ""},
            {"ky": "24CN", "ky_year": 2024, "gd": "C2", "tm": "2863", "so_tien": 1000000.0, "ht_year": 2024, "no_co": ""},
        ]
        m = compute_tq_metrics(tq_rows, 2023, 2024)
        self.assertEqual(round(m["opening_exact"]["LPMB"]), 3000000)
        self.assertEqual(round(m["lpmb_ps_by_ky"][2023]), 4000000)
        self.assertEqual(round(m["lpmb_c2_by_ht"][2023]), 4000000)
        self.assertEqual(round(m["lpmb_ps_by_ky"][2024]), 2000000)
        self.assertEqual(round(m["lpmb_c2_by_ht"][2024]), 3000000)
        lpmb_close = round(m["opening_exact"]["LPMB"] + sum(m["lpmb_ps_by_ky"].values()) - sum(m["lpmb_c2_by_ht"].values()))
        self.assertEqual(lpmb_close, 2000000)

    def test_qhns_bien_ban_vs_doi_chieu_sheet(self) -> None:
        """Kiểm thử tách riêng sheet QHNS (vào Biên bản, không nhận nộp thừa kỳ trước) và sheet DoiChieu_QHNS (khớp thực tế 100% TQ)."""
        from qhns_summarizer import integrate_tq_and_qhns_into_workbook

        with tempfile.TemporaryDirectory() as tmpdir:
            wb_decl = openpyxl.Workbook()
            ws2024 = wb_decl.active
            ws2024.title = "2024"
            ws2024["C23"] = 10000000
            ws2024["C76"] = 0
            ws2024["C90"] = 5000000
            decl_path = os.path.join(tmpdir, "decl.xlsx")
            wb_decl.save(decl_path)

            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            # Nộp thừa GTGT kỳ trước (-2,000,000) và nộp thừa TNDN kỳ trước (-5,000,000)
            ws_tq.append(["7901", "2312", "1", "C2", "1701", "VND", -2000000, datetime.datetime(2023, 12, 20)])
            ws_tq.append(["7901", "23CN", "2", "C2", "1052", "VND", -5000000, datetime.datetime(2023, 12, 25)])
            # Phát sinh & nộp năm 2024
            ws_tq.append(["7901", "2401", "3", "TK", "1701", "VND", 10000000, datetime.datetime(2024, 2, 15)])
            ws_tq.append(["7901", "2401", "4", "C2", "1701", "VND", -8000000, datetime.datetime(2024, 2, 20)])
            ws_tq.append(["7901", "2401", "5", "TK", "1001", "VND", 5000000, datetime.datetime(2024, 2, 15)])
            ws_tq.append(["7901", "2401", "6", "C2", "1001", "VND", -5000000, datetime.datetime(2024, 2, 20)])
            ws_tq.append(["7901", "24CN", "7", "TK", "2862", "VND", 3000000, datetime.datetime(2024, 1, 15)])
            ws_tq.append(["7901", "24CN", "8", "C2", "2862", "VND", -3000000, datetime.datetime(2024, 1, 20)])
            # Các sắc thuế/khoản thu khác trên TQ: Phạt VPHC (4254) nộp thừa -1,000,000 & Tiền chậm nộp TNCN (4917) khớp 0
            ws_tq.append(["7901", "2405", "9", "QP", "4254", "VND", 2000000, datetime.datetime(2024, 5, 10)])
            ws_tq.append(["7901", "2405", "10", "C2", "4254", "VND", -3000000, datetime.datetime(2024, 5, 15)])
            ws_tq.append(["7901", "2406", "11", "NC", "4917", "VND", 500000, datetime.datetime(2024, 6, 10)])
            ws_tq.append(["7901", "2406", "12", "C2", "4917", "VND", -500000, datetime.datetime(2024, 6, 15)])
            tq_path = os.path.join(tmpdir, "tq.xlsx")
            wb_tq.save(tq_path)

            out_path = os.path.join(tmpdir, "out.xlsx")
            integrate_tq_and_qhns_into_workbook(decl_path, tq_path, year_from=2024, year_to=2024, output_path=out_path)

            wb_out = openpyxl.load_workbook(out_path, data_only=False)
            self.assertIn("QHNS", wb_out.sheetnames)
            self.assertIn("DoiChieu_QHNS", wb_out.sheetnames)

            ws_q = wb_out["QHNS"]
            # Trên QHNS (Biên bản): đầu kỳ GTGT (H5+I5) = 0 (không ghi nhận nộp thừa -2,000,000)
            self.assertEqual(ws_q["H5"].value, 0)
            self.assertEqual(ws_q["I5"].value, 0)
            # Trên QHNS (Biên bản): đầu kỳ TNDN không ghi nhận nộp thừa -5,000,000
            self.assertIsNone(ws_q["C7"].value)
            # Dòng 5: Lệ phí môn bài có mặt trên QHNS
            self.assertEqual(ws_q["B12"].value, "Lệ phí môn bài")
            self.assertEqual(ws_q["D12"].value, 3000000)
            self.assertEqual(ws_q["F12"].value, 3000000)

            ws_dc = wb_out["DoiChieu_QHNS"]
            # Bảng 1 Đối chiếu: GTGT (row 3) có D3 = -2,000,000 (nộp thừa kỳ trước không vào Biên bản), H3 = 0 (dư cuối thực tế TQ)
            self.assertEqual(ws_dc["D3"].value, -2000000)
            self.assertEqual(ws_dc["H3"].value, 0)
            # TNDN (row 4) có D4 = -5,000,000, F4 là công thức động ='2024'!G76, H4 = -5,000,000
            self.assertEqual(ws_dc["D4"].value, -5000000)
            self.assertEqual(ws_dc["F4"].value, "='2024'!G76")
            self.assertEqual(ws_dc["H4"].value, -5000000)
            # LPMB (row 6) có D6 = 0, H6 = 0
            self.assertEqual(ws_dc["D6"].value, 0)
            self.assertEqual(ws_dc["H6"].value, 0)
            # Các sắc thuế khác trên TQ tự động xuất hiện trên Bảng 1 DoiChieu_QHNS: 4254 (row 7) và 4917 (row 8)
            self.assertEqual(ws_dc["B7"].value, "4254")
            self.assertEqual(ws_dc["H7"].value, -1000000)
            self.assertEqual(ws_dc["B8"].value, "4917")
            self.assertEqual(ws_dc["H8"].value, 0)
            # Dòng TỔNG CỘNG cuối Bảng 1 (row 9)
            self.assertEqual(ws_dc["A9"].value, "TỔNG CỘNG TẤT CẢ SẮC THUẾ TRÊN TQ")
            # Sheet Tinh_TienChamNop cũng được tự động sinh kèm theo
            self.assertIn("Tinh_TienChamNop", wb_out.sheetnames)

    def test_v14_late_payment_and_penalty_sheet(self) -> None:
        """Kiểm thử v1.4: Biểu tính tiền chậm nộp & xử phạt VPHC theo từng thời kỳ (0,05% / 0,03%, thời hiệu 5/10 năm, lệch kỳ, quý/tháng)."""
        import datetime
        from penalty_calculator import (
            add_late_payment_sheet,
            compute_cit_start_date,
            compute_license_fee_start_date,
            compute_pit_start_date,
            compute_vat_monthly_start_date,
            compute_vat_quarterly_start_date,
        )

        # 1. Kiểm tra ngày bắt đầu tính chậm nộp theo quy định các thời kỳ
        self.assertEqual(compute_vat_monthly_start_date(2023, 5), datetime.date(2023, 6, 21))
        self.assertEqual(compute_vat_monthly_start_date(2023, 12), datetime.date(2024, 1, 21))
        # GTGT Quý trước và sau 01/07/2020 (Luật 38/2019/QH14)
        self.assertEqual(compute_vat_quarterly_start_date(2019, 2), datetime.date(2019, 7, 31))
        self.assertEqual(compute_vat_quarterly_start_date(2020, 1), datetime.date(2020, 5, 1))
        self.assertEqual(compute_vat_quarterly_start_date(2020, 2), datetime.date(2020, 8, 1))
        self.assertEqual(compute_vat_quarterly_start_date(2024, 4), datetime.date(2025, 2, 1))
        # TNDN năm dương lịch vs năm tài chính lệch kỳ (kết thúc 30/06)
        self.assertEqual(compute_cit_start_date(2024), datetime.date(2025, 4, 1))
        self.assertEqual(compute_cit_start_date(2024, datetime.date(2025, 6, 30)), datetime.date(2025, 10, 1))
        # TNCN & LPMB
        self.assertEqual(compute_pit_start_date(2024), datetime.date(2025, 4, 1))
        self.assertEqual(compute_license_fee_start_date(2024), datetime.date(2024, 1, 31))

        # 2. Tạo workbook giả lập gồm các năm 2016, 2023, 2024 (kèm sheet năm '2024' để kiểm tra tự động link E76/E90)
        wb = openpyxl.Workbook()
        ws_2024 = wb.active
        ws_2024.title = "2024"
        ws_2024["E76"] = 100000000
        ws_2024["E90"] = 25000000

        cutoff = datetime.date(2026, 9, 15)
        ws = add_late_payment_sheet(
            wb,
            years=[2016, 2023, 2024],
            cutoff_date=cutoff,
            vat_frequency="month",
            fiscal_ranges={2024: (datetime.date(2024, 7, 1), datetime.date(2025, 6, 30))},
        )

        # Kiểm tra Khối thông số điều khiển (D4..D7, H4..H7)
        self.assertEqual(ws["D4"].value, cutoff)
        self.assertEqual(ws["D5"].value, datetime.date(2016, 6, 30))
        self.assertAlmostEqual(ws["D6"].value, 0.0003)
        self.assertAlmostEqual(ws["D7"].value, 0.0005)
        self.assertAlmostEqual(ws["H4"].value, 0.20)
        self.assertAlmostEqual(ws["H5"].value, 1.5)
        self.assertEqual(ws["H6"].value, 5)
        self.assertEqual(ws["H7"].value, 10)

        # Tìm dòng TNDN năm 2024 & TNCN năm 2024 để xác nhận tự động nối '2024'!E76, '2024'!E90 và NTC lệch kỳ 01/10/2025
        found_cit_2024 = False
        found_pit_2024 = False
        found_chk_zero = False
        in_2024 = False
        for r in range(13, ws.max_row + 1):
            b_val = str(ws.cell(r, 2).value or "")
            if b_val == "Năm 2024":
                in_2024 = True
            elif b_val.startswith("Phạt VPHC về thủ tục"):
                in_2024 = False
            if in_2024 and b_val.startswith("Thuế TNDN"):
                found_cit_2024 = True
                self.assertEqual(ws.cell(r, 3).value, "=IF('2024'!E76>0,'2024'!E76,0)")
                self.assertEqual(ws.cell(r, 9).value, datetime.date(2025, 10, 1))
                self.assertEqual(ws.cell(r, 10).value, "=$D$4")
                # Công thức phạt 20% có kiểm tra thời hiệu 5 năm ($H$6*12)
                self.assertIn("EDATE(", str(ws.cell(r, 5).value))
                # Công thức tách 2 giai đoạn 0,05% (cột K) và 0,03% (cột L)
                self.assertIn("$D$5", str(ws.cell(r, 11).value))
                self.assertIn("$D$5+1", str(ws.cell(r, 12).value))
                self.assertIn("$D$7", str(ws.cell(r, 13).value))
                self.assertIn("$D$6", str(ws.cell(r, 13).value))
            if in_2024 and b_val == "Thuế TNCN":
                found_pit_2024 = True
                self.assertEqual(ws.cell(r, 3).value, "=IF('2024'!E90>0,'2024'!E90,0)")
                self.assertEqual(ws.cell(r, 9).value, datetime.date(2025, 4, 1))
            if "Kiểm tra khớp Tổng bảng tổng hợp" in str(ws.cell(r, 1).value or ""):
                found_chk_zero = True
                self.assertTrue(str(ws.cell(r, 8).value).startswith("=H"))

        self.assertTrue(found_cit_2024)
        self.assertTrue(found_pit_2024)
        self.assertTrue(found_chk_zero)

        # Kiểm tra số ngày chậm nộp 2023-06-21 đến 2026-09-15 đúng 1183 ngày như mẫu thực tế
        self.assertEqual((datetime.date(2026, 9, 15) - datetime.date(2023, 6, 21)).days + 1, 1183)

    def test_v14_audit_report_sheet_links(self) -> None:
        """Kiểm thử v1.4: Sheet BaoCao_BienBan tự động cài sẵn công thức liên kết đủ 10 bảng từ các sheet năm và Tinh_TienChamNop."""
        import datetime
        from penalty_calculator import add_late_payment_sheet
        from report_builder import add_audit_report_sheet

        wb = openpyxl.Workbook()
        ws23 = wb.active
        ws23.title = "2023"
        wb.create_sheet("2024")
        wb.create_sheet("2025")

        add_late_payment_sheet(wb, years=[2023, 2024, 2025], cutoff_date=datetime.date(2026, 9, 15))
        ws_rep = add_audit_report_sheet(wb, years=[2023, 2024, 2025])

        all_formulas = [
            str(ws_rep.cell(r, c).value or "")
            for r in range(1, ws_rep.max_row + 1)
            for c in range(1, 9)
        ]
        # Kiểm tra liên kết trực tiếp tới các sheet năm cho đủ 9 bảng chi tiết
        for expected_ref in (
            "='2023'!C14",
            "='2024'!C3+'2024'!C18",
            "='2025'!C20",
            "='2025'!C19",
            "='2025'!C23",
            "='2025'!D24",
            "='2025'!C60",
            "='2024'!C73+'2024'!C74",
            "='2025'!C76",
            "='2023'!C84+'2023'!C86+'2023'!C88",
            "='2023'!C90",
            "='2023'!C95",
            "='2023'!C97",
            "='2023'!C98",
        ):
            self.assertIn(expected_ref, all_formulas, f"Thiếu công thức liên kết {expected_ref} trên BaoCao_BienBan")

        # Kiểm tra Bảng 10 liên kết sang 'Tinh_TienChamNop' cho Tiền chậm nộp, Phạt 20%, Trốn thuế, VPHC
        pen_links = [f for f in all_formulas if "'Tinh_TienChamNop'!" in f]
        self.assertGreaterEqual(len(pen_links), 15)

        # Kiểm tra dòng 18 (Giảm GTGT khấu trừ chuyển kỳ sau) & dòng 19 (Giảm lỗ chuyển kỳ sau)
        labels = {str(ws_rep.cell(r, 2).value or ""): str(ws_rep.cell(r, 3).value or "") for r in range(1, ws_rep.max_row + 1)}
        self.assertIn("Giảm số thuế GTGT còn được khấu trừ chuyển kỳ sau", labels)
        self.assertTrue(labels["Giảm số thuế GTGT còn được khấu trừ chuyển kỳ sau"].startswith("=MAX(0,-H"))
        self.assertIn("Giảm lỗ chuyển kỳ sau", labels)
        self.assertIn("MIN(ABS(", labels["Giảm lỗ chuyển kỳ sau"])

        # Kiểm tra tuyệt đối không tô màu nền trên BaoCao_BienBan để copy thẳng vào Biên bản Word
        for r in range(1, ws_rep.max_row + 1):
            for c in range(1, 9):
                self.assertIsNone(ws_rep.cell(r, c).fill.fill_type)

    def test_v15_tq_ntnn_reconciliation_on_qhns(self) -> None:
        """Kiểm thử v1.5: Tích hợp Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) vào QHNS, DoiChieu_QHNS và dòng 97/98 sheet năm."""
        from qhns_summarizer import integrate_tq_and_qhns_into_workbook

        with tempfile.TemporaryDirectory() as tmpdir:
            wb_src = openpyxl.Workbook()
            ws24 = wb_src.active
            ws24.title = "2024"
            ws24["C23"] = 10_000_000
            ws24["C76"] = 20_000_000
            ws24["C90"] = 5_000_000
            ws24["C97"] = 12_000_000
            ws24["C98"] = 8_000_000
            src_path = os.path.join(tmpdir, "tonghop.xlsx")
            wb_src.save(src_path)

            # Sổ TQ chính của DN
            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_tq.append(["7901", "2401", "1", "TK", "1701", "VND", 10_000_000, datetime.datetime(2024, 2, 20)])
            ws_tq.append(["7901", "2401", "2", "C2", "1701", "VND", -10_000_000, datetime.datetime(2024, 2, 25)])
            ws_tq.append(["7901", "24CN", "3", "QT", "1052", "VND", 20_000_000, datetime.datetime(2025, 3, 31)])
            ws_tq.append(["7901", "24CN", "4", "C2", "1052", "VND", -20_000_000, datetime.datetime(2024, 12, 25)])
            ws_tq.append(["7901", "24CN", "5", "TK", "1001", "VND", 6_000_000, datetime.datetime(2025, 3, 31)])
            ws_tq.append(["7901", "24CN", "5b", "PP", "1001", "VND", -1_000_000, datetime.datetime(2024, 12, 20)])
            ws_tq.append(["7901", "24CN", "6", "C2", "1001", "VND", -5_000_000, datetime.datetime(2024, 12, 25)])
            tq_path = os.path.join(tmpdir, "tq_main.xlsx")
            wb_tq.save(tq_path)

            # Sổ TQ riêng của Thuế Nhà Thầu (TQ_NTNN)
            wb_nt = openpyxl.Workbook()
            ws_nt = wb_nt.active
            ws_nt.title = "TQ"
            ws_nt.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_nt.append(["7901", "2406", "1", "TK", "1701", "VND", 12_000_000, datetime.datetime(2024, 6, 15)])
            ws_nt.append(["7901", "2406", "2", "C2", "1701", "VND", -12_000_000, datetime.datetime(2024, 6, 18)])
            ws_nt.append(["7901", "2406", "3", "TK", "1052", "VND", 8_000_000, datetime.datetime(2024, 6, 15)])
            ws_nt.append(["7901", "2406", "4", "C2", "1052", "VND", -6_000_000, datetime.datetime(2024, 6, 18)])
            nt_path = os.path.join(tmpdir, "tq_ntnn.xlsx")
            wb_nt.save(nt_path)

            out_path = os.path.join(tmpdir, "out_qhns_ntnn.xlsx")
            integrate_tq_and_qhns_into_workbook(
                src_path, tq_path, year_from=2024, year_to=2024, output_path=out_path, tq_ntnn_excel_path=nt_path
            )

            wb_out = openpyxl.load_workbook(out_path, data_only=False)
            self.assertIn("TQ_NTNN", wb_out.sheetnames)
            # Kiểm tra phát sinh TNCN (1001) đã cộng cả giao dịch PP (6,000,000 TK - 1,000,000 PP = 5,000,000)
            self.assertEqual(wb_out["2024"]["F90"].value, 5_000_000)
            # Kiểm tra đối chiếu dòng 97 & 98 trên sheet 2024
            self.assertEqual(wb_out["2024"]["F97"].value, 12_000_000)
            self.assertEqual(wb_out["2024"]["F98"].value, 8_000_000)
            # Kiểm tra dòng Thuế nhà thầu trên QHNS (row 9=nt_tong, row 10=nt_gtgt, row 11=nt_tndn)
            ws_q = wb_out["QHNS"]
            self.assertEqual(ws_q["D10"].value, "='2024'!C97")
            self.assertEqual(ws_q["F10"].value, 12_000_000)
            self.assertEqual(ws_q["D11"].value, "='2024'!C98")
            self.assertEqual(ws_q["F11"].value, 6_000_000)
            # Kiểm tra trên DoiChieu_QHNS có đủ 2 dòng NTNN (1701 & 1052) với dư cuối thực tế 0 và 2,000,000
            ws_dc = wb_out["DoiChieu_QHNS"]
            dc_map = {str(ws_dc.cell(r, 2).value): ws_dc.cell(r, 8).value for r in range(3, 12)}
            self.assertEqual(dc_map.get("1701 (NTNN)"), 0)
            self.assertEqual(dc_map.get("1052 (NTNN)"), 2_000_000)

            # Kiểm tra phân loại tờ khai Nhà thầu (NTNN) tách riêng thành nhóm NTNN (không trộn sang GTGT/TNDN)
            from tax_parser import classify_tax_type
            self.assertEqual(classify_tax_type("01/NTNN", "Tờ khai thuế nhà thầu nước ngoài (GTGT, TNDN)"), "NTNN")
            self.assertEqual(classify_tax_type("03/NTNN", "Tờ khai thuế thu nhập doanh nghiệp nhà thầu nước ngoài"), "NTNN")
            self.assertEqual(classify_tax_type("02/NTNN", "Tờ khai quyết toán thuế nhà thầu nước ngoài"), "NTNN")

    def test_v16_ntnn_xml_tt103_vs_tt80_and_yearly_link(self) -> None:
        """Kiểm thử v1.6: Đọc chính xác 01/NTNN theo đúng chuẩn từng thời kỳ (TT103 ct6..ct13 vs TT80 ct7..ct14), xuất So Thue NTNN TH/Cuoi, link C95/C97/C98 và đối chiếu thiếu tờ khai trên TQ_NTNN."""
        import datetime
        from excel_exporter import export_to_excel
        from qhns_summarizer import integrate_tq_and_qhns_into_workbook
        from tax_parser import parse_tax_xml

        # 1. Tờ khai 01/NTNN năm 2021 theo chuẩn TT103/2014 (ct6=DT GTGT, ct7=Tỷ lệ 5%, ct8=Thuế GTGT, ct9=DT TNDN chưa gồm GTGT, ct10=Tỷ lệ 5%, ct12=Thuế TNDN, ct13=Tổng thuế)
        xml_tt103_2021 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu>
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>01/NTNN</maTKhai>
          <tenTKhai>Tờ khai thuế nhà thầu nước ngoài</tenTKhai>
          <moTaBMau>(Ban hành kèm theo Thông tư số 103/2014/TT-BTC)</moTaBMau>
          <soLan>0</soLan>
          <ngayLapTKhai>2021-06-15</ngayLapTKhai>
          <KyKKhaiThue>
            <kyKKhai>05/2021</kyKKhai>
            <kyKKhaiTuNgay>2021-05-01</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>2021-05-31</kyKKhaiDenNgay>
          </KyKKhaiThue>
        </TKhaiThue>
        <NNT>
          <mst>0100000099</mst>
          <tenNNT>Cong ty Mau A (Nop thay NTNN)</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <CTietTKhaiChinh>
        <ct6>105263158</ct6>
        <ct7>5</ct7>
        <ct8>5263158</ct8>
        <ct9>100000000</ct9>
        <ct10>5</ct10>
        <ct11>0</ct11>
        <ct12>5000000</ct12>
        <ct13>10263158</ct13>
      </CTietTKhaiChinh>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        # 2. Tờ khai 01/NTNN năm 2024 theo chuẩn TT80/2021 (ct6=Ngày thanh toán, ct7=DT GTGT, ct8=Tỷ lệ 5%, ct9=Thuế GTGT, ct10=DT TNDN, ct13=Thuế TNDN, ct14=Tổng thuế)
        # Bản chính thức lần 0 và bản bổ sung lần 1 cùng kỳ 06/2024 (bản BS1 phải thay thế bản lần 0 trên So Thue NTNN Cuoi)
        xml_tt80_2024_ct = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu>
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>01/NTNN</maTKhai>
          <tenTKhai>Tờ khai thuế nhà thầu nước ngoài</tenTKhai>
          <moTaBMau>(Ban hành kèm theo Thông tư số 80/2021/TT-BTC)</moTaBMau>
          <soLan>0</soLan>
          <ngayLapTKhai>2024-07-10</ngayLapTKhai>
          <KyKKhaiThue>
            <kyKKhai>06/2024</kyKKhai>
            <kyKKhaiTuNgay>2024-06-01</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>2024-06-30</kyKKhaiDenNgay>
          </KyKKhaiThue>
        </TKhaiThue>
        <NNT>
          <mst>0100000099</mst>
          <tenNNT>Cong ty Mau A (Nop thay NTNN)</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <CTietTKhaiChinh>
        <ct6>2024-06-20</ct6>
        <ct7>210526316</ct7>
        <ct8>5</ct8>
        <ct9>10526316</ct9>
        <ct10>200000000</ct10>
        <ct11>5</ct11>
        <ct12>0</ct12>
        <ct13>10000000</ct13>
        <ct14>20526316</ct14>
      </CTietTKhaiChinh>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        xml_tt80_2024_bs1 = xml_tt80_2024_ct.replace("<soLan>0</soLan>", "<soLan>1</soLan>").replace(
            "<ngayLapTKhai>2024-07-10</ngayLapTKhai>", "<ngayLapTKhai>2024-08-05</ngayLapTKhai>"
        ).replace("<ct7>210526316</ct7>", "<ct7>315789474</ct7>").replace(
            "<ct9>10526316</ct9>", "<ct9>15789474</ct9>"
        ).replace("<ct10>200000000</ct10>", "<ct10>300000000</ct10>").replace(
            "<ct13>10000000</ct13>", "<ct13>15000000</ct13>"
        ).replace("<ct14>20526316</ct14>", "<ct14>30789474</ct14>")

        with tempfile.TemporaryDirectory() as tmpdir:
            p1 = os.path.join(tmpdir, "ntnn_2021_tt103.xml")
            p2 = os.path.join(tmpdir, "ntnn_2024_06_ct.xml")
            p3 = os.path.join(tmpdir, "ntnn_2024_06_bs1.xml")
            for p, content in ((p1, xml_tt103_2021), (p2, xml_tt80_2024_ct), (p3, xml_tt80_2024_bs1)):
                with open(p, "w", encoding="utf-8") as f:
                    f.write(content)

            decls = [parse_tax_xml(p) for p in (p1, p2, p3)]
            # Kiểm tra nhận diện chính xác TT103 vs TT80 (KHÔNG lệch cột!)
            self.assertEqual(decls[0].tax_type, "NTNN")
            self.assertEqual(decls[0].indicators["_ntnn_schema"], "TT103")
            self.assertEqual(round(decls[0].indicators["_ntnn_dt_gtgt"]), 105_263_158)
            self.assertEqual(round(decls[0].indicators["_ntnn_thue_gtgt"]), 5_263_158)
            self.assertEqual(round(decls[0].indicators["_ntnn_dt_tndn"]), 100_000_000)
            self.assertEqual(round(decls[0].indicators["_ntnn_thue_tndn"]), 5_000_000)

            self.assertEqual(decls[2].indicators["_ntnn_schema"], "TT80")
            self.assertEqual(round(decls[2].indicators["_ntnn_dt_gtgt"]), 315_789_474)
            self.assertEqual(round(decls[2].indicators["_ntnn_thue_gtgt"]), 15_789_474)
            self.assertEqual(round(decls[2].indicators["_ntnn_dt_tndn"]), 300_000_000)
            self.assertEqual(round(decls[2].indicators["_ntnn_thue_tndn"]), 15_000_000)

            out_xlsx = os.path.join(tmpdir, "tonghop_ntnn.xlsx")
            export_to_excel(decls, out_xlsx)

            wb = openpyxl.load_workbook(out_xlsx, data_only=False)
            self.assertIn("So Thue NTNN TH", wb.sheetnames)
            self.assertIn("So Thue NTNN Cuoi", wb.sheetnames)
            self.assertIn("2021", wb.sheetnames)
            self.assertIn("2024", wb.sheetnames)

            # So Thue NTNN TH có đủ 3 tờ khai (row 2..4); So Thue NTNN Cuoi chỉ giữ 2 bản hiệu lực cuối (2021 + 2024 BS1) + 2 dòng Tổng cộng năm + 1 dòng Tổng cộng
            ws_th = wb["So Thue NTNN TH"]
            self.assertEqual(ws_th.max_row, 4)
            ws_cuoi = wb["So Thue NTNN Cuoi"]
            self.assertEqual(ws_cuoi.cell(3, 6).value, "Tổng cộng năm 2021")
            self.assertEqual(ws_cuoi.cell(5, 6).value, "Tổng cộng năm 2024")
            self.assertEqual(ws_cuoi.cell(4, 15).value, 300_000_000)  # Bản BS1 thay thế bản chính thức

            # Kiểm tra công thức liên kết sang sheet '2021' và '2024' (C95=DT chưa gồm GTGT, C97=GTGT NTNN, C98=TNDN NTNN)
            self.assertEqual(wb["2021"]["C95"].value, "='So Thue NTNN Cuoi'!Q3")
            self.assertEqual(wb["2021"]["C97"].value, "='So Thue NTNN Cuoi'!N3")
            self.assertEqual(wb["2021"]["C98"].value, "='So Thue NTNN Cuoi'!P3")
            self.assertEqual(wb["2024"]["C95"].value, "='So Thue NTNN Cuoi'!Q5")
            self.assertEqual(wb["2024"]["C97"].value, "='So Thue NTNN Cuoi'!N5")
            self.assertEqual(wb["2024"]["C98"].value, "='So Thue NTNN Cuoi'!P5")

            # Kiểm tra đối chiếu thiếu tờ khai trên TQ_NTNN (giả lập TQ_NTNN 2024 chỉ ghi nhận 10,000,000 GTGT < 15,789,474 tờ khai)
            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_tq.append(["7901", "24CN", "1", "TK", "2862", "VND", 3_000_000, datetime.datetime(2024, 1, 15)])
            tq_main_path = os.path.join(tmpdir, "tq_main.xlsx")
            wb_tq.save(tq_main_path)

            wb_nt = openpyxl.Workbook()
            ws_nt = wb_nt.active
            ws_nt.title = "TQ"
            ws_nt.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_nt.append(["7901", "2406", "1", "TK", "1701", "VND", 10_000_000, datetime.datetime(2024, 7, 10)])
            ws_nt.append(["7901", "2406", "2", "C2", "1701", "VND", -10_000_000, datetime.datetime(2024, 7, 12)])
            ws_nt.append(["7901", "2406", "3", "TK", "1052", "VND", 15_000_000, datetime.datetime(2024, 7, 10)])
            ws_nt.append(["7901", "2406", "4", "C2", "1052", "VND", -15_000_000, datetime.datetime(2024, 7, 12)])
            tq_nt_path = os.path.join(tmpdir, "tq_ntnn.xlsx")
            wb_nt.save(tq_nt_path)

            res = integrate_tq_and_qhns_into_workbook(
                out_xlsx, tq_main_path, year_from=2024, year_to=2024, output_path=out_xlsx, tq_ntnn_excel_path=tq_nt_path
            )
            self.assertTrue(any("Thuế GTGT Nhà thầu" in a for a in res["alerts"]))
            wb_res = openpyxl.load_workbook(out_xlsx, data_only=False)
            ws_q = wb_res["QHNS"]
            # Năm 2024: GTGT NTNN thiếu trên TQ_NTNN -> Cột D lấy F97, Cột E lấy G97; TNDN NTNN khớp 15tr -> Cột D lấy C98
            self.assertEqual(ws_q["D10"].value, "='2024'!F97")
            self.assertEqual(ws_q["E10"].value, "='2024'!G97")
            self.assertEqual(ws_q["D11"].value, "='2024'!C98")


if __name__ == "__main__":
    unittest.main()
