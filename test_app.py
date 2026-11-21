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
            self.assertEqual(wb_master_tncn.sheetnames, ["Tong_hop_File", "TNCN", "2020", "2024"])
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
                file_path="f1.xml", file_name="0313248400_01GTGT_25M08_CT_0.xml", tax_type="GTGT",
                mst="0313248400", ky_kkhai="08/2025", so_lan="0", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 9, 17), indicators={"ct40": 374974620}
            )
            d2_ct0 = TaxDeclaration(
                file_path="f2.xml", file_name="0313248400_08-2025_01-GTGT_CT0_1_21420250009848413.xml", tax_type="GTGT",
                mst="0313248400", ky_kkhai="08/2025", so_lan="0", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 9, 17), indicators={"ct40": 374974620}
            )
            d3_bs1 = TaxDeclaration(
                file_path="f3.xml", file_name="0313248400_01GTGT_25M08_BS_1.xml", tax_type="GTGT",
                mst="0313248400", ky_kkhai="08/2025", so_lan="1", tu_ngay=datetime.date(2025, 8, 1),
                ngay_lap=datetime.date(2025, 10, 14), indicators={"ct40": 365632924}
            )
            d4_bs1 = TaxDeclaration(
                file_path="f4.xml", file_name="0313248400_08-2025_01-GTGT_BS1_1_21420250009862365.xml", tax_type="GTGT",
                mst="0313248400", ky_kkhai="08/2025", so_lan="1", tu_ngay=datetime.date(2025, 8, 1),
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
                mst="0313248400", ky_kkhai="2020",
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
                mst="0313248400", ky_kkhai="12/2024", so_lan="0",
                tu_ngay=datetime.date(2024, 12, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 1, 20),
                indicators={"ct34": 50000000000, "ct40": 5000000000}
            )
            dTNDN = TaxDeclaration(
                file_path="tndn2024.xml", file_name="tndn2024.xml", tax_type="TNDN",
                mst="0313248400", ky_kkhai="2024", so_lan="0",
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
      <NNT><mst>0313248400</mst><tenNNT>CTY A</tenNNT></NNT>
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
      <NNT><mst>0313248400</mst><tenNNT>CTY A</tenNNT></NNT>
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
      <NNT><mst>0313248400</mst><tenNNT>CTY A</tenNNT></NNT>
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
      <NNT><mst>0313248400</mst><tenNNT>CTY A</tenNNT></NNT>
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
      <NNT><mst>0313248400</mst><tenNNT>CTY A</tenNNT></NNT>
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
      <NNT><mst>0313248400</mst><tenNNT>CTY A</tenNNT></NNT>
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
                mst="0313248400", ten_nnt="CTY A",
                tu_ngay=datetime.date(2024, 12, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 1, 20),
                indicators={"ct22": 1000000, "ct32": 50000000, "ct33": 5000000, "ct43": 2000000}
            )
            dTNDN = TaxDeclaration(
                file_path="tndn_2024.xml", file_name="tndn_2024.xml",
                ma_tkhai="892", ten_tkhai="Tờ khai QTT TNDN", tax_type="TNDN",
                ky_kkhai="2024", so_lan="0",
                mst="0313248400", ten_nnt="CTY A",
                tu_ngay=datetime.date(2024, 1, 1), den_ngay=datetime.date(2024, 12, 31),
                ngay_lap=datetime.date(2025, 3, 30),
                indicators={"ct04": 50000000, "ctB4": 3000000}
            )
            dTNCN = TaxDeclaration(
                file_path="tncn_2024.xml", file_name="tncn_2024.xml",
                ma_tkhai="864", ten_tkhai="Tờ khai QTT TNCN", tax_type="TNCN",
                ky_kkhai="2024", so_lan="0",
                mst="0313248400", ten_nnt="CTY A",
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


if __name__ == "__main__":
    unittest.main()
