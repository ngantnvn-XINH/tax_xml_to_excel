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
            # Kiểm tra nhận diện chính xác M13_FLAT vs M14_SHIFTED
            self.assertEqual(decls[0].tax_type, "NTNN")
            self.assertEqual(decls[0].indicators["_ntnn_schema"], "M13_FLAT")
            self.assertEqual(round(decls[0].indicators["_ntnn_dt_gtgt"]), 105_263_158)
            self.assertEqual(round(decls[0].indicators["_ntnn_thue_gtgt"]), 5_263_158)
            self.assertEqual(round(decls[0].indicators["_ntnn_dt_tndn"]), 100_000_000)
            self.assertEqual(round(decls[0].indicators["_ntnn_thue_tndn"]), 5_000_000)

            self.assertEqual(decls[2].indicators["_ntnn_schema"], "M14_SHIFTED")
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

            # So Thue NTNN TH: 3 dòng (row 2..4) + 1 dòng TỔNG CỘNG (row 5)
            self.assertEqual(wb["So Thue NTNN TH"].max_row, 5)
            ws_cuoi = wb["So Thue NTNN Cuoi"]
            self.assertEqual(ws_cuoi.cell(3, 6).value, "Tổng cộng năm 2021")
            self.assertEqual(ws_cuoi.cell(5, 6).value, "Tổng cộng năm 2024")
            self.assertEqual(ws_cuoi.cell(4, 16).value, 300_000_000)  # Bản BS1 thay thế bản chính thức (cột P = DT TNDN)

            # Kiểm tra công thức liên kết sang sheet '2021' và '2024' (C95=DT chưa gồm GTGT cột M, C97=GTGT cột O, C98=TNDN cột Q)
            self.assertEqual(wb["2021"]["C95"].value, "='So Thue NTNN Cuoi'!M3")
            self.assertEqual(wb["2021"]["C97"].value, "='So Thue NTNN Cuoi'!O3")
            self.assertEqual(wb["2021"]["C98"].value, "='So Thue NTNN Cuoi'!Q3")
            self.assertEqual(wb["2024"]["C95"].value, "='So Thue NTNN Cuoi'!M5")
            self.assertEqual(wb["2024"]["C97"].value, "='So Thue NTNN Cuoi'!O5")
            self.assertEqual(wb["2024"]["C98"].value, "='So Thue NTNN Cuoi'!Q5")

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

    def test_v16_qhns_cqt_filter_and_other_cqt_separation(self) -> None:
        """Kiểm thử v1.6: Lọc đúng Mã CQT quản lý (VD 7901 hoặc CQT tùy chọn) vào QHNS và tách riêng số nộp/phát sinh tại các CQT khác sang sheet TQ_CQT_Khac."""
        import datetime
        from qhns_summarizer import integrate_tq_and_qhns_into_workbook

        with tempfile.TemporaryDirectory() as tmpdir:
            wb_decl = openpyxl.Workbook()
            ws24 = wb_decl.active
            ws24.title = "2024"
            ws24["C23"] = 10_000_000
            ws24["C76"] = 0
            ws24["C90"] = 0
            decl_path = os.path.join(tmpdir, "decl_cqt.xlsx")
            wb_decl.save(decl_path)

            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            # Giao dịch thuộc đúng CQT quản lý 7901 (phải vào QHNS)
            ws_tq.append(["7901", "2401", "1", "TK", "1701", "VND", 10_000_000, datetime.datetime(2024, 2, 20)])
            ws_tq.append(["7901", "2401", "2", "C2", "1701", "VND", -10_000_000, datetime.datetime(2024, 2, 22)])
            # Giao dịch thuộc CQT khác (VD 0101 Hà Nội & 7401 Bình Dương -> phải tách sang TQ_CQT_Khac, KHÔNG gộp vào QHNS 7901!)
            ws_tq.append(["0101", "2403", "3", "TK", "1701", "VND", 4_000_000, datetime.datetime(2024, 4, 15)])
            ws_tq.append(["0101", "2403", "4", "C2", "1701", "VND", -4_000_000, datetime.datetime(2024, 4, 18)])
            ws_tq.append(["7401", "24CN", "5", "C2", "2864", "VND", -1_000_000, datetime.datetime(2024, 1, 25)])
            tq_path = os.path.join(tmpdir, "tq_multi_cqt.xlsx")
            wb_tq.save(tq_path)

            out_path = os.path.join(tmpdir, "out_cqt.xlsx")
            res = integrate_tq_and_qhns_into_workbook(
                decl_path, tq_path, year_from=2024, year_to=2024, output_path=out_path, target_cqt="7901"
            )

            self.assertEqual(res["target_cqt"], "7901")
            self.assertEqual(res["other_cqt_count"], 3)
            self.assertTrue(any("TQ_CQT_Khac" in a for a in res["alerts"]))

            wb_out = openpyxl.load_workbook(out_path, data_only=False)
            self.assertIn("TQ_CQT_Khac", wb_out.sheetnames)
            # F23 trên sheet 2024 & F5 (Số đã nộp GTGT) trên QHNS CHỈ lấy đúng 10,000,000 của CQT 7901 (không cộng 4,000,000 của 0101)
            self.assertEqual(wb_out["2024"]["F23"].value, 10_000_000)
            self.assertEqual(wb_out["QHNS"]["F5"].value, 10_000_000)
            # LPMB trên QHNS không cộng nhầm 1,000,000 của CQT 7401
            self.assertEqual(wb_out["QHNS"]["F12"].value, 0)

            # Kiểm tra sheet TQ_CQT_Khac tách đủ cả 0101 (4tr) và 7401 (1tr)
            ws_other = wb_out["TQ_CQT_Khac"]
            cqt_codes_b1 = [str(ws_other.cell(r, 3).value or "") for r in range(3, 8)]
            self.assertIn("0101", cqt_codes_b1)
            self.assertIn("7401", cqt_codes_b1)

    def test_v16_tncn_pp_other_cqt_included_in_accrual(self) -> None:
        """Kiểm thử v1.6: PP phân bổ Thuế TNCN (1001) hạch toán ở CQT khác (VD 0101) vẫn được cộng vào F90 & phát sinh TNCN, còn số nộp C2 ở CQT khác vẫn tách sang TQ_CQT_Khac và DoiChieu_QHNS khớp 0đ."""
        import datetime
        from qhns_summarizer import integrate_tq_and_qhns_into_workbook

        with tempfile.TemporaryDirectory() as tmpdir:
            wb_decl = openpyxl.Workbook()
            ws24 = wb_decl.active
            ws24.title = "2024"
            ws24["C23"] = 0
            ws24["C76"] = 0
            # Tờ khai QTT TNCN toàn công ty = 100tr (80tr tại 7901 + 20tr phân bổ PP sang 0101)
            ws24["C90"] = 100_000_000
            decl_path = os.path.join(tmpdir, "decl_pp.xlsx")
            wb_decl.save(decl_path)

            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_tq.append(["7901", "24CN", "1", "TK", "1001", "VND", 80_000_000, datetime.datetime(2024, 12, 20)])
            ws_tq.append(["7901", "24CN", "2", "C2", "1001", "VND", -80_000_000, datetime.datetime(2024, 12, 22)])
            # Phân bổ PP sang CQT khác 0101 (20tr) và nộp C2 tại 0101 (-20tr)
            ws_tq.append(["0101", "24CN", "3", "PP", "1001", "VND", 20_000_000, datetime.datetime(2024, 12, 20)])
            ws_tq.append(["0101", "24CN", "4", "C2", "1001", "VND", -20_000_000, datetime.datetime(2024, 12, 22)])
            tq_path = os.path.join(tmpdir, "tq_pp.xlsx")
            wb_tq.save(tq_path)

            out_path = os.path.join(tmpdir, "out_pp.xlsx")
            res = integrate_tq_and_qhns_into_workbook(
                decl_path, tq_path, year_from=2024, year_to=2024, output_path=out_path, target_cqt="7901"
            )

            wb_out = openpyxl.load_workbook(out_path, data_only=False)
            # 1. F90 trên sheet 2024 phải bằng 100tr (80tr tại 7901 + 20tr PP tại 0101) -> KHÔNG báo thiếu tờ khai!
            self.assertEqual(wb_out["2024"]["F90"].value, 100_000_000)
            self.assertFalse(any("thiếu TK" in a for a in res["alerts"]))
            # 2. Trên QHNS: D8 (Phát sinh TNCN) = ='2024'!C90 (100tr, có cả PP), F8 (Đã nộp tại 7901) = 80tr
            ws_q = wb_out["QHNS"]
            self.assertEqual(ws_q["D8"].value, "='2024'!C90")
            self.assertEqual(ws_q["F8"].value, 80_000_000)
            # 3. Trên DoiChieu_QHNS: F5 (TNCN) = ='2024'!G90+20000000 -> quy đổi khớp 0đ với dư cuối 7901 (H5 = 0)
            ws_dc = wb_out["DoiChieu_QHNS"]
            self.assertEqual(ws_dc["F5"].value, "='2024'!G90+20000000")
            self.assertEqual(ws_dc["H5"].value, 0)
            # 4. TQ_CQT_Khac vẫn tách riêng 20tr PP và 20tr C2 tại 0101
            self.assertIn("TQ_CQT_Khac", wb_out.sheetnames)

    def test_v16_qhns_single_tq_ntnn_only(self) -> None:
        """Kiểm thử v1.6: Có cái nào đối chiếu cái đó — chạy chỉ với TQ_NTNN (không cần TQ công ty) vẫn ra QHNS & DoiChieu_QHNS chuẩn."""
        import datetime
        from qhns_summarizer import integrate_tq_and_qhns_into_workbook

        with tempfile.TemporaryDirectory() as tmpdir:
            wb_decl = openpyxl.Workbook()
            ws24 = wb_decl.active
            ws24.title = "2024"
            ws24["C97"] = 5_000_000
            ws24["C98"] = 4_000_000
            decl_path = os.path.join(tmpdir, "decl_ntnn_only.xlsx")
            wb_decl.save(decl_path)

            wb_nt = openpyxl.Workbook()
            ws_nt = wb_nt.active
            ws_nt.title = "TQ"
            ws_nt.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_nt.append(["7901", "2406", "1", "TK", "1701", "VND", 5_000_000, datetime.datetime(2024, 6, 15)])
            ws_nt.append(["7901", "2406", "2", "C2", "1701", "VND", -5_000_000, datetime.datetime(2024, 6, 20)])
            ws_nt.append(["7901", "2406", "3", "TK", "1052", "VND", 4_000_000, datetime.datetime(2024, 6, 15)])
            ws_nt.append(["7901", "2406", "4", "C2", "1052", "VND", -4_000_000, datetime.datetime(2024, 6, 20)])
            nt_path = os.path.join(tmpdir, "tq_ntnn_only.xlsx")
            wb_nt.save(nt_path)

            out_path = os.path.join(tmpdir, "out_ntnn_only.xlsx")
            res = integrate_tq_and_qhns_into_workbook(
                decl_path, None, year_from=2024, year_to=2024, output_path=out_path, tq_ntnn_excel_path=nt_path
            )
            self.assertEqual(res["alerts"], [])
            wb_out = openpyxl.load_workbook(out_path, data_only=False)
            self.assertIn("TQ_NTNN", wb_out.sheetnames)
            self.assertNotIn("TQ", wb_out.sheetnames)
            self.assertIn("QHNS", wb_out.sheetnames)
            self.assertIn("DoiChieu_QHNS", wb_out.sheetnames)
            self.assertEqual(wb_out["2024"]["F97"].value, 5_000_000)
            self.assertEqual(wb_out["2024"]["F98"].value, 4_000_000)
            self.assertEqual(wb_out["QHNS"]["D10"].value, "='2024'!C97")
            self.assertEqual(wb_out["QHNS"]["F10"].value, 5_000_000)

    def test_v16_split_by_tax_type_period_reconciliation(self) -> None:
        """Kiểm thử v1.6: Tách mỗi sắc thuế ra 1 file riêng (khi người dùng tích chọn) và đối chiếu từng kỳ kê khai (DoiChieu_KyKeKhai)."""
        import datetime
        from excel_exporter import export_to_excel
        from qhns_summarizer import integrate_period_reconciliation_for_tax_type
        from tax_parser import TaxDeclaration

        with tempfile.TemporaryDirectory() as tmpdir:
            d_gtgt = TaxDeclaration(
                file_path="gtgt_01_2024.xml",
                file_name="gtgt_01_2024.xml",
                tax_type="GTGT",
                ma_tkhai="01/GTGT",
                ten_tkhai="Tờ khai thuế GTGT",
                mst="0100000088",
                ten_nnt="Cong ty Mau B",
                ky_kkhai="01/2024",
                tu_ngay=datetime.date(2024, 1, 1),
                den_ngay=datetime.date(2024, 1, 31),
                ngay_lap=datetime.date(2024, 2, 18),
                so_lan="0",
                indicators={"ct40": 15_000_000, "ct40a": 15_000_000},
            )
            d_ntnn = TaxDeclaration(
                file_path="ntnn_06_2024.xml",
                file_name="ntnn_06_2024.xml",
                tax_type="NTNN",
                ma_tkhai="01/NTNN",
                ten_tkhai="Tờ khai thuế nhà thầu nước ngoài",
                mst="0100000099",
                ten_nnt="Cong ty Mau B (NTNN)",
                ky_kkhai="06/2024",
                tu_ngay=datetime.date(2024, 6, 1),
                den_ngay=datetime.date(2024, 6, 30),
                ngay_lap=datetime.date(2024, 7, 10),
                so_lan="0",
                indicators={
                    "_ntnn_dt_gtgt": 105_263_158,
                    "_ntnn_thue_gtgt": 5_263_158,
                    "_ntnn_dt_tndn": 100_000_000,
                    "_ntnn_thue_tndn": 5_000_000,
                    "_ntnn_tong_thue": 10_263_158,
                    "_ntnn_schema": "TT80",
                },
            )

            base_out = os.path.join(tmpdir, "TongHop.xlsx")
            created = export_to_excel([d_gtgt, d_ntnn], base_out, split_by_tax_type=True)
            self.assertEqual(len(created), 2)
            gtgt_file = [p for p in created if p.endswith("_GTGT.xlsx")][0]
            ntnn_file = [p for p in created if p.endswith("_NTNN.xlsx")][0]

            # Tạo Sổ TQ cho GTGT và TQ_NTNN cho NTNN
            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_tq.append(["7901", "2401", "1", "TK", "1701", "VND", 15_000_000, datetime.datetime(2024, 2, 18)])
            ws_tq.append(["7901", "2401", "2", "C2", "1701", "VND", -15_000_000, datetime.datetime(2024, 2, 20)])
            tq_path = os.path.join(tmpdir, "tq.xlsx")
            wb_tq.save(tq_path)

            wb_nt = openpyxl.Workbook()
            ws_nt = wb_nt.active
            ws_nt.title = "TQ"
            ws_nt.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_nt.append(["7901", "2406", "1", "TK", "1701", "VND", 5_263_158, datetime.datetime(2024, 7, 10)])
            ws_nt.append(["7901", "2406", "2", "C2", "1701", "VND", -5_263_158, datetime.datetime(2024, 7, 12)])
            ws_nt.append(["7901", "2406", "3", "TK", "1052", "VND", 5_000_000, datetime.datetime(2024, 7, 10)])
            ws_nt.append(["7901", "2406", "4", "C2", "1052", "VND", -3_000_000, datetime.datetime(2024, 7, 12)])
            nt_path = os.path.join(tmpdir, "tq_ntnn.xlsx")
            wb_nt.save(nt_path)

            r_g = integrate_period_reconciliation_for_tax_type(gtgt_file, "GTGT", tq_excel_path=tq_path, target_cqt="7901")
            r_n = integrate_period_reconciliation_for_tax_type(ntnn_file, "NTNN", tq_ntnn_excel_path=nt_path, target_cqt="7901")
            self.assertTrue(r_g["reconciled"])
            self.assertTrue(r_n["reconciled"])

            wb_g_out = openpyxl.load_workbook(gtgt_file, data_only=False)
            self.assertIn("DoiChieu_KyKeKhai", wb_g_out.sheetnames)
            ws_dc_g = wb_g_out["DoiChieu_KyKeKhai"]
            self.assertEqual(ws_dc_g["B3"].value, "01/2024")
            self.assertEqual(ws_dc_g["E3"].value, "='So Thue GTGT Cuoi'!AJ3")
            self.assertEqual(ws_dc_g["F3"].value, 15_000_000)
            self.assertEqual(ws_dc_g["H3"].value, 15_000_000)

            wb_n_out = openpyxl.load_workbook(ntnn_file, data_only=False)
            self.assertIn("DoiChieu_KyKeKhai", wb_n_out.sheetnames)
            ws_dc_n = wb_n_out["DoiChieu_KyKeKhai"]
            self.assertEqual(ws_dc_n["B3"].value, "06/2024")
            self.assertEqual(ws_dc_n["D3"].value, "1052")
            self.assertEqual(ws_dc_n["F3"].value, 5_000_000)
            self.assertEqual(ws_dc_n["H3"].value, 3_000_000)
            self.assertEqual(ws_dc_n["D4"].value, "1701")
            self.assertEqual(ws_dc_n["F4"].value, 5_263_158)

    def test_v16_export_qhns_separate_files_per_tax_type(self) -> None:
        """Kiểm thử v1.6: Xuất riêng phần QHNS (đối chiếu theo kỳ kê khai) ra từng file Excel độc lập cho mỗi sắc thuế (QHNS_GTGT, QHNS_TNDN, QHNS_TNCN, QHNS_NTNN)."""
        import datetime
        from excel_exporter import export_to_excel
        from qhns_summarizer import export_qhns_by_tax_type_period_files
        from tax_parser import TaxDeclaration

        with tempfile.TemporaryDirectory() as tmpdir:
            d_gtgt = TaxDeclaration(
                file_path="g1.xml",
                file_name="g1.xml",
                tax_type="GTGT",
                ma_tkhai="01/GTGT",
                ten_tkhai="Tờ khai thuế GTGT",
                mst="0100000088",
                ten_nnt="Cong ty Mau C",
                ky_kkhai="01/2024",
                tu_ngay=datetime.date(2024, 1, 1),
                den_ngay=datetime.date(2024, 1, 31),
                ngay_lap=datetime.date(2024, 2, 18),
                so_lan="0",
                indicators={"ct40": 15_000_000, "ct40a": 15_000_000},
            )
            d_ntnn = TaxDeclaration(
                file_path="n1.xml",
                file_name="n1.xml",
                tax_type="NTNN",
                ma_tkhai="01/NTNN",
                ten_tkhai="Tờ khai thuế nhà thầu nước ngoài",
                mst="0100000099",
                ten_nnt="Cong ty Mau C (NTNN)",
                ky_kkhai="06/2024",
                tu_ngay=datetime.date(2024, 6, 1),
                den_ngay=datetime.date(2024, 6, 30),
                ngay_lap=datetime.date(2024, 7, 10),
                so_lan="0",
                indicators={
                    "_ntnn_dt_gtgt": 105_263_158,
                    "_ntnn_thue_gtgt": 5_263_158,
                    "_ntnn_dt_tndn": 100_000_000,
                    "_ntnn_thue_tndn": 5_000_000,
                    "_ntnn_tong_thue": 10_263_158,
                    "_ntnn_schema": "TT80",
                },
            )
            consolidated = os.path.join(tmpdir, "TongHop.xlsx")
            export_to_excel([d_gtgt, d_ntnn], consolidated)

            wb_tq = openpyxl.Workbook()
            ws_tq = wb_tq.active
            ws_tq.title = "TQ"
            ws_tq.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_tq.append(["7901", "2401", "1", "TK", "1701", "VND", 15_000_000, datetime.datetime(2024, 2, 18)])
            ws_tq.append(["7901", "2401", "2", "C2", "1701", "VND", -10_000_000, datetime.datetime(2024, 2, 20)])
            ws_tq.append(["7901", "24CN", "3", "TK", "1001", "VND", 8_000_000, datetime.datetime(2024, 12, 20)])
            ws_tq.append(["0101", "24CN", "4", "PP", "1001", "VND", 2_000_000, datetime.datetime(2024, 12, 20)])
            tq_path = os.path.join(tmpdir, "tq.xlsx")
            wb_tq.save(tq_path)

            wb_nt = openpyxl.Workbook()
            ws_nt = wb_nt.active
            ws_nt.title = "TQ"
            ws_nt.append(["Cơ quan thuế", "Kỳ báo cáo", "Số hệ thống", "Ký hiệu giao dịch", "Tiểu mục", "Loại tiền", "Số tiền", "Ngày hạch toán"])
            ws_nt.append(["7901", "2406", "1", "TK", "1701", "VND", 5_263_158, datetime.datetime(2024, 7, 10)])
            ws_nt.append(["7901", "2406", "2", "C2", "1701", "VND", -5_263_158, datetime.datetime(2024, 7, 12)])
            nt_path = os.path.join(tmpdir, "tq_ntnn.xlsx")
            wb_nt.save(nt_path)

            out_dir = os.path.join(tmpdir, "qhns_split")
            files = export_qhns_by_tax_type_period_files(
                consolidated, out_dir, tq_excel_path=tq_path, tq_ntnn_excel_path=nt_path, year_from=2024, year_to=2024, target_cqt="7901"
            )
            names = {os.path.basename(p) for p in files}
            self.assertIn("QHNS_GTGT.xlsx", names)
            self.assertIn("QHNS_TNCN.xlsx", names)
            self.assertIn("QHNS_NTNN.xlsx", names)

            wb_g = openpyxl.load_workbook(os.path.join(out_dir, "QHNS_GTGT.xlsx"), data_only=False)
            self.assertEqual(wb_g.sheetnames, ["QHNS_TheoKy", "ChiTiet_TQ"])
            ws_g = wb_g["QHNS_TheoKy"]
            self.assertEqual(ws_g["B4"].value, "01/2024")
            self.assertEqual(ws_g["E4"].value, 15_000_000)
            self.assertEqual(ws_g["F4"].value, 15_000_000)
            self.assertEqual(ws_g["H4"].value, 10_000_000)

            # Kiểm tra QHNS_TNCN đã cộng cả 2tr PP phân bổ ở CQT 0101 vào cột Phát sinh TQ (8tr + 2tr = 10tr)
            wb_tn = openpyxl.load_workbook(os.path.join(out_dir, "QHNS_TNCN.xlsx"), data_only=False)
            ws_tn = wb_tn["QHNS_TheoKy"]
            self.assertEqual(ws_tn["F4"].value, 10_000_000)

    def test_v16_ntnn_monthly_latest_vs_daily_sum_both(self) -> None:
        """Kiểm thử v1.6: NTNN theo tháng lấy tờ khai cuối cùng; NTNN theo lần phát sinh (theo ngày) cộng cả Chính thức và Bổ sung."""
        import datetime
        from ntnn_exporter import filter_latest_ntnn_declarations, is_ntnn_per_occurrence
        from tax_parser import TaxDeclaration

        # 1. Nộp theo THÁNG (06/2024): Chính thức + Bổ sung lần 1 -> chỉ giữ Bổ sung lần 1
        d_m0 = TaxDeclaration(
            file_path="m0.xml", file_name="m0.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ky_kkhai="06/2024",
            tu_ngay=datetime.date(2024, 6, 1), den_ngay=datetime.date(2024, 6, 30),
            ngay_lap=datetime.date(2024, 7, 10), so_lan="0",
            indicators={"_kieu_ky": "M", "_ntnn_dt_tndn": 100_000_000, "_ntnn_thue_tndn": 5_000_000},
        )
        d_m1 = TaxDeclaration(
            file_path="m1.xml", file_name="m1.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ky_kkhai="06/2024",
            tu_ngay=datetime.date(2024, 6, 1), den_ngay=datetime.date(2024, 6, 30),
            ngay_lap=datetime.date(2024, 8, 5), so_lan="1",
            indicators={"_kieu_ky": "M", "_ntnn_dt_tndn": 150_000_000, "_ntnn_thue_tndn": 7_500_000},
        )

        # 2. Nộp theo LẦN PHÁT SINH (ngày 15/06/2024): Chính thức + Bổ sung lần 1 -> GIỮ & CỘNG CẢ 2!
        d_d0 = TaxDeclaration(
            file_path="d0.xml", file_name="d0.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ky_kkhai="15/06/2024",
            tu_ngay=datetime.date(2024, 6, 15), den_ngay=datetime.date(2024, 6, 15),
            ngay_lap=datetime.date(2024, 6, 20), so_lan="0",
            indicators={"_kieu_ky": "D", "_ntnn_dt_tndn": 40_000_000, "_ntnn_thue_tndn": 2_000_000},
        )
        d_d1 = TaxDeclaration(
            file_path="d1.xml", file_name="d1.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ky_kkhai="15/06/2024",
            tu_ngay=datetime.date(2024, 6, 15), den_ngay=datetime.date(2024, 6, 15),
            ngay_lap=datetime.date(2024, 6, 25), so_lan="1",
            indicators={"_kieu_ky": "D", "_ntnn_dt_tndn": 60_000_000, "_ntnn_thue_tndn": 3_000_000},
        )

        self.assertFalse(is_ntnn_per_occurrence(d_m0))
        self.assertTrue(is_ntnn_per_occurrence(d_d0))

        # Kiểm tra cả trường hợp HTKK ghi kieuKy='P' hoặc 'L' hoặc ky_kkhai='P/15/06/2024'
        d_p = TaxDeclaration(
            file_path="dp.xml", file_name="dp.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ky_kkhai="P/15/06/2024",
            tu_ngay=datetime.date(2024, 6, 15), den_ngay=datetime.date(2024, 6, 15),
            ngay_lap=datetime.date(2024, 6, 26), so_lan="0",
            indicators={"_kieu_ky": "P", "_ntnn_dt_tndn": 20_000_000, "_ntnn_thue_tndn": 1_000_000},
        )
        self.assertTrue(is_ntnn_per_occurrence(d_p))

        kept = filter_latest_ntnn_declarations([d_m0, d_m1, d_d0, d_d1, d_p])
        kept_names = [d.file_name for d in kept]
        # Tháng: bỏ m0, chỉ giữ m1; Lần phát sinh: giữ CẢ d0, d1 và dp!
        self.assertNotIn("m0.xml", kept_names)
        self.assertIn("m1.xml", kept_names)
        self.assertIn("d0.xml", kept_names)
        self.assertIn("d1.xml", kept_names)
        self.assertIn("dp.xml", kept_names)
        self.assertEqual(len(kept), 4)

    def test_v16_ntnn_th_expands_all_contract_rows(self) -> None:
        """Kiểm thử v1.6: Sheet So Thue NTNN TH xuất CHI TIẾT TẤT CẢ CÁC HỢP ĐỒNG trên từng tờ khai 01/NTNN."""
        from excel_exporter import export_to_excel
        from tax_parser import parse_tax_xml

        xml_two_contracts = """<?xml version="1.0" encoding="UTF-8"?>
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
            <kieuKy>M</kieuKy>
            <kyKKhai>06/2024</kyKKhai>
            <kyKKhaiTuNgay>2024-06-01</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>2024-06-30</kyKKhaiDenNgay>
          </KyKKhaiThue>
        </TKhaiThue>
        <NNT>
          <mst>0100000099</mst>
          <tenNNT>Cong ty Mau D</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <CTietTKhaiChinh>
        <ct1>Nha thau X - Dich vu phan mem</ct1>
        <ct2>9900000001</ct2>
        <ct3>HD-2024-001 ngay 15/01/2024</ct3>
        <ct4>100000000</ct4>
        <ct5>2024-06-10</ct5>
        <ct6>105263158</ct6>
        <ct7>5</ct7>
        <ct8>5263158</ct8>
        <ct9>100000000</ct9>
        <ct10>5</ct10>
        <ct11>0</ct11>
        <ct12>5000000</ct12>
        <ct13>10263158</ct13>
      </CTietTKhaiChinh>
      <CTietTKhaiChinh>
        <ct1>Nha thau Y - Ban quyen</ct1>
        <ct2>9900000002</ct2>
        <ct3>HD-2024-002 ngay 01/03/2024</ct3>
        <ct4>200000000</ct4>
        <ct5>2024-06-25</ct5>
        <ct6>0</ct6>
        <ct7>0</ct7>
        <ct8>0</ct8>
        <ct9>200000000</ct9>
        <ct10>10</ct10>
        <ct11>0</ct11>
        <ct12>20000000</ct12>
        <ct13>20000000</ct13>
      </CTietTKhaiChinh>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        with tempfile.TemporaryDirectory() as tmpdir:
            p = os.path.join(tmpdir, "ntnn_2hd.xml")
            with open(p, "w", encoding="utf-8") as f:
                f.write(xml_two_contracts)
            d = parse_tax_xml(p)
            self.assertEqual(len(d.indicators.get("_ntnn_rows", [])), 2)

            out_x = os.path.join(tmpdir, "out_2hd.xlsx")
            export_to_excel([d], out_x)
            wb = openpyxl.load_workbook(out_x, data_only=False)
            ws_th = wb["So Thue NTNN TH"]
            # Header (1) + 2 dòng hợp đồng (2, 3) + 1 dòng TỔNG CỘNG (4)
            self.assertEqual(ws_th.max_row, 4)
            self.assertEqual(ws_th.cell(2, 9).value, "Nha thau X - Dich vu phan mem")
            self.assertEqual(ws_th.cell(2, 10).value, "9900000001")
            self.assertEqual(ws_th.cell(2, 11).value, "HD-2024-001 ngay 15/01/2024")
            self.assertEqual(ws_th.cell(2, 12).value, 100_000_000)
            self.assertEqual(ws_th.cell(2, 13).value, "2024-06-10")
            self.assertEqual(ws_th.cell(2, 16).value, 5_263_158)
            self.assertEqual(ws_th.cell(3, 9).value, "Nha thau Y - Ban quyen")
            self.assertEqual(ws_th.cell(3, 11).value, "HD-2024-002 ngay 01/03/2024")
            self.assertEqual(ws_th.cell(3, 12).value, 200_000_000)
            self.assertEqual(ws_th.cell(3, 18).value, 10)
            self.assertEqual(ws_th.cell(3, 20).value, 20_000_000)
            self.assertIn("TỔNG CỘNG", str(ws_th.cell(4, 9).value))
            self.assertIn('SUMIF($V$2:$V$3,"Hợp lệ"', str(ws_th.cell(4, 21).value))

    def test_v16_offset_fiscal_year_gtgt_opening_and_no_c2_tk_double_count(self) -> None:
        """Kiểm thử v1.6: Đơn vị lệch kỳ (01/04..31/03): phát sinh GTGT lệch kỳ, đầu kỳ trước 01/04, số nộp C2 theo năm dương lịch 01/01..31/12, KHÔNG cộng trùng (+C2)+(+TK/QT)."""
        import datetime
        from qhns_summarizer import compute_tq_metrics

        tq_rows = [
            # Kỳ 11/2019 (1911): đã nộp xong trong năm dương lịch 2019 -> KHÔNG vào đầu kỳ 2020
            {"ky": "1911", "ky_year": 2019, "gd": "TK", "tm": "1701", "so_tien": 60_000_000.0, "ngay_ht": datetime.datetime(2019, 12, 18), "ht_year": 2019, "no_co": ""},
            {"ky": "1911", "ky_year": 2019, "gd": "C2", "tm": "1701", "so_tien": -60_000_000.0, "ngay_ht": datetime.datetime(2019, 12, 19), "ht_year": 2019, "no_co": ""},
            # Kỳ 02/2020 (2002): thuộc kỳ trước (< 01/04/2020), nộp trong năm dương lịch 2020 (19/03/2020) -> vào đầu kỳ 50tr & đã nộp 2020 50tr
            {"ky": "2002", "ky_year": 2020, "gd": "TK", "tm": "1701", "so_tien": 50_000_000.0, "ngay_ht": datetime.datetime(2020, 3, 18), "ht_year": 2020, "no_co": ""},
            {"ky": "2002", "ky_year": 2020, "gd": "C2", "tm": "1701", "so_tien": -50_000_000.0, "ngay_ht": datetime.datetime(2020, 3, 19), "ht_year": 2020, "no_co": ""},
            # Kỳ 03/2020 (2003): thuộc kỳ trước (< 01/04/2020), TK +80tr & C2 -80tr hạch toán 20/04/2020 -> Đầu kỳ += ĐÚNG 80tr (KHÔNG phải 160tr!)
            {"ky": "2003", "ky_year": 2020, "gd": "TK", "tm": "1701", "so_tien": 80_000_000.0, "ngay_ht": datetime.datetime(2020, 4, 20), "ht_year": 2020, "no_co": ""},
            {"ky": "2003", "ky_year": 2020, "gd": "C2", "tm": "1701", "so_tien": -80_000_000.0, "ngay_ht": datetime.datetime(2020, 4, 20), "ht_year": 2020, "no_co": ""},
            # Kỳ 04/2020 (2004) & 01/2021 (2101): phát sinh thuộc năm tài chính lệch kỳ 2020 (04/2020..03/2021);
            # riêng C2 nộp tháng 02/2021 có ngày hạch toán năm dương lịch 2021 -> KHÔNG vào đã nộp 2020!
            {"ky": "2004", "ky_year": 2020, "gd": "TK", "tm": "1701", "so_tien": 30_000_000.0, "ngay_ht": datetime.datetime(2020, 5, 20), "ht_year": 2020, "no_co": ""},
            {"ky": "2004", "ky_year": 2020, "gd": "C2", "tm": "1701", "so_tien": -30_000_000.0, "ngay_ht": datetime.datetime(2020, 5, 20), "ht_year": 2020, "no_co": ""},
            {"ky": "2101", "ky_year": 2021, "gd": "TK", "tm": "1701", "so_tien": 25_000_000.0, "ngay_ht": datetime.datetime(2021, 2, 20), "ht_year": 2021, "no_co": ""},
            {"ky": "2101", "ky_year": 2021, "gd": "C2", "tm": "1701", "so_tien": -25_000_000.0, "ngay_ht": datetime.datetime(2021, 2, 20), "ht_year": 2021, "no_co": ""},
            # Kiểm tra thêm QT kỳ trước (15CN) vừa có QT +40tr vừa có C2 -40tr trong năm 2020 -> Đầu kỳ += ĐÚNG 40tr (KHÔNG cộng đôi thành 80tr!)
            {"ky": "15CN", "ky_year": 2015, "gd": "QT", "tm": "1701", "so_tien": 40_000_000.0, "ngay_ht": datetime.datetime(2020, 6, 15), "ht_year": 2020, "no_co": ""},
            {"ky": "15CN", "ky_year": 2015, "gd": "C2", "tm": "1701", "so_tien": -40_000_000.0, "ngay_ht": datetime.datetime(2020, 6, 18), "ht_year": 2020, "no_co": ""},
        ]

        m = compute_tq_metrics(tq_rows, 2020, 2020, fiscal_start_month=4)
        # Đầu kỳ GTGT (các kỳ trước 01/04/2020 nộp trong năm dương lịch 2020) = 50tr (2002) + 80tr (2003) + 40tr (15CN) = 170tr
        self.assertEqual(round(m["dk_gtgt_c2_prev"] + m["dk_gtgt_qt_truoc"]), 170_000_000)
        self.assertEqual(round(m["opening_exact"]["1701"]), 170_000_000)
        # Phát sinh GTGT năm tài chính lệch kỳ 2020 (04/2020..03/2021) = 30tr (2004) + 25tr (2101) = 55tr
        self.assertEqual(round(m["gtgt_tk_by_ky"][2020]), 55_000_000)
        # Đã nộp GTGT theo năm dương lịch 2020 (01/01/2020..31/12/2020) = 50tr (2002) + 80tr (2003) + 30tr (2004) + 40tr (15CN) = 200tr
        self.assertEqual(round(m["gtgt_c2_by_ht"][2020]), 200_000_000)
        # Dư cuối kỳ 2020 = 170tr + 55tr - 200tr = 25tr (đúng bằng kỳ 01/2021 nộp sang năm dương lịch 2021!)
        self.assertEqual(round(m["opening_exact"]["1701"] + m["gtgt_tk_by_ky"][2020] - m["gtgt_c2_by_ht"][2020]), 25_000_000)

    def test_v16_ntnn_merge_two_msts_and_flag_duplicates(self) -> None:
        """Kiểm thử v1.6: Gộp NTNN từ cả 2 thư mục (MST công ty chính + MST nhà thầu riêng), báo trùng trên So Thue NTNN TH và tự loại bản trùng khi tổng hợp trên So Thue NTNN Cuoi."""
        import datetime
        from excel_exporter import export_to_excel
        from tax_parser import TaxDeclaration

        # 1. Kỳ 03/2021 nộp qua MST công ty chính (không có ở MST nhà thầu) -> PHẢI GIỮ
        d_main_2021 = TaxDeclaration(
            file_path="main_2021_03.xml", file_name="main_2021_03.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000088", ten_nnt="Cong ty Mau E", ky_kkhai="03/2021",
            tu_ngay=datetime.date(2021, 3, 1), den_ngay=datetime.date(2021, 3, 31),
            ngay_lap=datetime.date(2021, 4, 15), so_lan="0",
            indicators={"_kieu_ky": "M", "_ntnn_dt_tndn": 200_000_000, "_ntnn_thue_tndn": 10_000_000, "_ntnn_schema": "TT103"},
        )
        # 2. Kỳ 06/2024 trùng y như nhau giữa MST công ty chính và MST nhà thầu riêng -> TH giữ cả 2 + báo trùng; Cuoi chỉ giữ 1!
        d_main_dup = TaxDeclaration(
            file_path="main_2024_06.xml", file_name="main_2024_06.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000088", ten_nnt="Cong ty Mau E", ky_kkhai="06/2024",
            tu_ngay=datetime.date(2024, 6, 1), den_ngay=datetime.date(2024, 6, 30),
            ngay_lap=datetime.date(2024, 7, 10), so_lan="0",
            indicators={"_kieu_ky": "M", "_ntnn_dt_tndn": 100_000_000, "_ntnn_thue_tndn": 5_000_000, "_ntnn_schema": "TT80"},
        )
        d_nt_dup = TaxDeclaration(
            file_path="ntnn_2024_06.xml", file_name="ntnn_2024_06.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ten_nnt="Cong ty Mau E (NTNN)", ky_kkhai="06/2024",
            tu_ngay=datetime.date(2024, 6, 1), den_ngay=datetime.date(2024, 6, 30),
            ngay_lap=datetime.date(2024, 7, 10), so_lan="0",
            indicators={"_kieu_ky": "M", "_ntnn_dt_tndn": 100_000_000, "_ntnn_thue_tndn": 5_000_000, "_ntnn_schema": "TT80"},
        )
        # 3. Kỳ 07/2024 chỉ có ở MST nhà thầu riêng -> PHẢI GIỮ
        d_nt_2024 = TaxDeclaration(
            file_path="ntnn_2024_07.xml", file_name="ntnn_2024_07.xml", tax_type="NTNN", ma_tkhai="01/NTNN",
            mst="0100000099", ten_nnt="Cong ty Mau E (NTNN)", ky_kkhai="07/2024",
            tu_ngay=datetime.date(2024, 7, 1), den_ngay=datetime.date(2024, 7, 31),
            ngay_lap=datetime.date(2024, 8, 10), so_lan="0",
            indicators={"_kieu_ky": "M", "_ntnn_dt_tndn": 160_000_000, "_ntnn_thue_tndn": 8_000_000, "_ntnn_schema": "TT80"},
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_x = os.path.join(tmpdir, "ntnn_2mst.xlsx")
            export_to_excel([d_main_2021, d_main_dup, d_nt_dup, d_nt_2024], out_x)
            wb = openpyxl.load_workbook(out_x, data_only=False)

            # So Thue NTNN TH: giữ đủ cả 4 tờ khai (row 2..5) + dòng TỔNG CỘNG (row 6), và báo trùng ở cột 22!
            ws_th = wb["So Thue NTNN TH"]
            self.assertEqual(ws_th.max_row, 6)
            statuses = [str(ws_th.cell(r, 22).value or "") for r in range(2, 6)]
            self.assertEqual(sum(1 for s in statuses if "TRÙNG TỜ KHAI" in s), 1)

            # So Thue NTNN Cuoi: loại 1 bản trùng 06/2024 -> chỉ còn 3 tờ khai (2021: 1 dòng + 1 tổng; 2024: 2 dòng + 1 tổng; + 1 grand total = 7 dòng cả header)
            ws_c = wb["So Thue NTNN Cuoi"]
            kept_kys = [ws_c.cell(r, 6).value for r in range(2, ws_c.max_row + 1)]
            self.assertEqual(kept_kys.count("03/2021"), 1)
            self.assertEqual(kept_kys.count("06/2024"), 1)
            self.assertEqual(kept_kys.count("07/2024"), 1)

    def test_v16_ntnn_official_tct_nested_xml_structure(self) -> None:
        """Kiểm thử v1.6: Đọc chính xác 100% cấu trúc XML chuẩn TCT (<BKThueNTNN>/<ThueNTNN> lồng <ThueGTGT>+<ThueTNDN>, ct8=nil, không có tong_ct4)."""
        from excel_exporter import export_to_excel
        from tax_parser import parse_tax_xml

        # Dữ liệu giả lập ẩn danh 100% theo đúng cấu trúc thẻ XML chuẩn Tổng cục Thuế (maTKhai=41)
        xml_tct_nested = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>41</maTKhai>
          <tenTKhai>Tờ khai thuế nhà thầu (01/NTNN) (Kỳ 15/06/2023)</tenTKhai>
          <moTaBMau>Ban hành kèm theo Thông tư số 103/2014/TT-BTC</moTaBMau>
          <loaiTKhai>C</loaiTKhai>
          <soLan>0</soLan>
          <KyKKhaiThue>
            <kieuKy>D</kieuKy>
            <kyKKhai>15/06/2023</kyKKhai>
            <kyKKhaiTuNgay>15/06/2023</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>15/06/2023</kyKKhaiDenNgay>
          </KyKKhaiThue>
          <ngayLapTKhai>2023-06-15</ngayLapTKhai>
        </TKhaiThue>
        <NNT>
          <mst>0100000077</mst>
          <tenNNT>Cong ty Mau F</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <BKThueNTNN>
        <ThueNTNN id="ID_1">
          <ct1>Dich vu lai tien vay (Nha thau Mau Z)</ct1>
          <ct2>0100000077</ct2>
          <ct3>HD-TEST-01</ct3>
          <ct4>500000000</ct4>
          <ct5>2023-06-10</ct5>
          <ThueGTGT>
            <ct6>0</ct6>
            <ct7>0</ct7>
            <ct8 xsi:nil="true"/>
            <ct9>0</ct9>
          </ThueGTGT>
          <ThueTNDN>
            <ct10>500000000</ct10>
            <ct11>5</ct11>
            <ct12>0</ct12>
            <ct13>25000000</ct13>
          </ThueTNDN>
          <ct14>25000000</ct14>
        </ThueNTNN>
      </BKThueNTNN>
      <tong_ct6>0</tong_ct6>
      <tong_ct9>0</tong_ct9>
      <tong_ct10>500000000</tong_ct10>
      <tong_ct12>0</tong_ct12>
      <tong_ct13>25000000</tong_ct13>
      <tong_ct14>25000000</tong_ct14>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        with tempfile.TemporaryDirectory() as tmpdir:
            p = os.path.join(tmpdir, "tct_nested.xml")
            with open(p, "w", encoding="utf-8") as f:
                f.write(xml_tct_nested)
            d = parse_tax_xml(p)
            self.assertEqual(d.tax_type, "NTNN")
            self.assertEqual(d.indicators["_ntnn_schema"], "TCT_NESTED")
            # Không có tong_ct4 -> phải tự cộng chính xác từ <ThueNTNN>/<ct4> = 500,000,000!
            self.assertEqual(round(d.indicators["_ntnn_dt_chua_gtgt"]), 500_000_000)
            self.assertEqual(round(d.indicators["_ntnn_dt_gtgt"]), 0)
            self.assertEqual(round(d.indicators["_ntnn_thue_gtgt"]), 0)
            self.assertEqual(round(d.indicators["_ntnn_dt_tndn"]), 500_000_000)
            self.assertEqual(round(d.indicators["_ntnn_thue_tndn"]), 25_000_000)
            self.assertEqual(round(d.indicators["_ntnn_tong_thue"]), 25_000_000)

            out_x = os.path.join(tmpdir, "out_nested.xlsx")
            export_to_excel([d], out_x)
            wb = openpyxl.load_workbook(out_x, data_only=False)
            ws_th = wb["So Thue NTNN TH"]
            # Chỉ đúng 1 dòng hợp đồng (row 2) + 1 dòng TỔNG CỘNG (row 3) — KHÔNG bị tách thành 3 dòng vỡ!
            self.assertEqual(ws_th.max_row, 3)
            self.assertEqual(ws_th.cell(2, 9).value, "Dich vu lai tien vay (Nha thau Mau Z)")
            self.assertEqual(ws_th.cell(2, 10).value, "0100000077")
            self.assertEqual(ws_th.cell(2, 11).value, "HD-TEST-01")
            self.assertEqual(ws_th.cell(2, 12).value, 500_000_000)   # (4) DT chưa gồm GTGT
            self.assertEqual(ws_th.cell(2, 13).value, "2023-06-10")  # (5) Ngày thanh toán
            self.assertEqual(ws_th.cell(2, 14).value, 0)             # (6) DT GTGT
            self.assertEqual(ws_th.cell(2, 16).value, 0)             # (8) Thuế GTGT
            self.assertEqual(ws_th.cell(2, 17).value, 500_000_000)   # (9) DT TNDN
            self.assertEqual(ws_th.cell(2, 18).value, 5)             # (10) Tỷ lệ TNDN (%)
            self.assertEqual(ws_th.cell(2, 20).value, 25_000_000)    # (12) Thuế TNDN phải nộp
            self.assertEqual(ws_th.cell(2, 21).value, 25_000_000)    # (13) Tổng thuế phải nộp NSNN

    def test_v16_offset_fiscal_year_ntnn_export_and_qhns(self) -> None:
        """Kiểm thử v1.6: Đơn vị lệch kỳ (01/04..31/03) tổng hợp 'So Thue NTNN Cuoi (Lech ky)', link C95/C97/C98 và QHNS TQ_NTNN lệch kỳ giống GTGT."""
        from excel_exporter import export_to_excel
        from qhns_summarizer import compute_tq_metrics, detect_fiscal_start_month
        from tax_parser import TaxDeclaration

        # TNDN lệch kỳ 01/04/2020..31/03/2021 để kích hoạt fiscal_ranges
        d_tndn = TaxDeclaration(
            file_path="tndn_2020.xml", file_name="tndn_2020.xml",
            mst="0100000088", ten_nnt="CTY LECH KY TEST",
            ma_tkhai="03/TNDN", ten_tkhai="03/TNDN", tax_type="TNDN",
            ky_kkhai="2020", tu_ngay=datetime.date(2020, 4, 1), den_ngay=datetime.date(2021, 3, 31),
            so_lan="0", ngay_lap=datetime.date(2021, 6, 15),
            indicators={"A1": 0, "E": 0},
        )
        # NTNN kỳ 02/2020 -> thuộc năm tài chính 2019 (trước 04/2020)
        d_nt_prior = TaxDeclaration(
            file_path="nt_02_2020.xml", file_name="nt_02_2020.xml",
            mst="0100000089", ten_nnt="CTY LECH KY TEST (NTNN)",
            ma_tkhai="41", ten_tkhai="01/NTNN", tax_type="NTNN",
            ky_kkhai="02/2020", tu_ngay=datetime.date(2020, 2, 1), den_ngay=datetime.date(2020, 2, 29),
            so_lan="0", ngay_lap=datetime.date(2020, 3, 10),
            indicators={"ct4": 100_000_000, "tong_ct6": 100_000_000, "tong_ct9": 5_000_000, "tong_ct10": 100_000_000, "tong_ct13": 5_000_000, "tong_ct14": 10_000_000},
        )
        # NTNN kỳ 02/2021 -> thuộc năm tài chính 2020 (04/2020..03/2021)
        d_nt_fy20 = TaxDeclaration(
            file_path="nt_02_2021.xml", file_name="nt_02_2021.xml",
            mst="0100000089", ten_nnt="CTY LECH KY TEST (NTNN)",
            ma_tkhai="41", ten_tkhai="01/NTNN", tax_type="NTNN",
            ky_kkhai="02/2021", tu_ngay=datetime.date(2021, 2, 1), den_ngay=datetime.date(2021, 2, 28),
            so_lan="0", ngay_lap=datetime.date(2021, 3, 10),
            indicators={"ct4": 200_000_000, "tong_ct6": 200_000_000, "tong_ct9": 10_000_000, "tong_ct10": 200_000_000, "tong_ct13": 10_000_000, "tong_ct14": 20_000_000},
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_x = os.path.join(tmpdir, "out_ntnn_lk.xlsx")
            export_to_excel([d_tndn, d_nt_prior, d_nt_fy20], out_x)
            wb = openpyxl.load_workbook(out_x, data_only=False)
            self.assertIn("So Thue NTNN Cuoi (Lech ky)", wb.sheetnames)
            self.assertEqual(detect_fiscal_start_month(wb), 4)
            ws_lk = wb["So Thue NTNN Cuoi (Lech ky)"]
            # Kỳ 02/2020 phải nằm dưới nhóm năm 2019; kỳ 02/2021 nằm dưới nhóm năm 2020 (04/2020-03/2021)
            col_f_vals = [str(ws_lk.cell(r, 6).value or "") for r in range(2, ws_lk.max_row + 1)]
            self.assertTrue(any("Tổng cộng năm 2019" in v for v in col_f_vals))
            self.assertTrue(any("Tổng cộng năm 2020 (04/2020-03/2021)" in v for v in col_f_vals))
            # Sheet '2020' phải link C95, C97, C98 về 'So Thue NTNN Cuoi (Lech ky)'
            self.assertIn("'So Thue NTNN Cuoi (Lech ky)'!M", str(wb["2020"]["C95"].value))
            self.assertIn("'So Thue NTNN Cuoi (Lech ky)'!O", str(wb["2020"]["C97"].value))
            self.assertIn("'So Thue NTNN Cuoi (Lech ky)'!Q", str(wb["2020"]["C98"].value))

        # Kiểm tra QHNS TQ_NTNN lệch kỳ (tháng 4):
        # - Kỳ 02/2020 (và ngày 15/03/2020) nộp sang 04/2020 -> thuộc đầu kỳ năm tài chính 2020 (KHÔNG cộng trùng C2+TK)
        # - Kỳ 02/2021 -> thuộc phát sinh năm tài chính 2020
        ntnn_tq_rows = [
            {"tm": "1701", "gd": "TK", "ky": "02/2020", "ky_year": 2020, "ht_year": 2020, "ngay_ht": datetime.datetime(2020, 3, 20), "so_tien": 5_000_000, "no_co": ""},
            {"tm": "1701", "gd": "C2", "ky": "02/2020", "ky_year": 2020, "ht_year": 2020, "ngay_ht": datetime.datetime(2020, 4, 5), "so_tien": -5_000_000, "no_co": ""},
            {"tm": "1052", "gd": "TK", "ky": "15/03/2020", "ky_year": 2020, "ht_year": 2020, "ngay_ht": datetime.datetime(2020, 3, 25), "so_tien": 8_000_000, "no_co": ""},
            {"tm": "1052", "gd": "C2", "ky": "15/03/2020", "ky_year": 2020, "ht_year": 2020, "ngay_ht": datetime.datetime(2020, 4, 10), "so_tien": -8_000_000, "no_co": ""},
            {"tm": "1701", "gd": "TK", "ky": "02/2021", "ky_year": 2021, "ht_year": 2021, "ngay_ht": datetime.datetime(2021, 3, 20), "so_tien": 10_000_000, "no_co": ""},
            {"tm": "1052", "gd": "TK", "ky": "02/2021", "ky_year": 2021, "ht_year": 2021, "ngay_ht": datetime.datetime(2021, 3, 20), "so_tien": 12_000_000, "no_co": ""},
        ]
        m_nt = compute_tq_metrics(ntnn_tq_rows, 2020, 2020, fiscal_start_month=4, is_ntnn=True)
        # Số đầu kỳ 2020: GTGT NTNN = 5tr, TNDN NTNN = 8tr (KHÔNG cộng trùng C2+TK)
        self.assertEqual(round(m_nt["opening_exact"]["1701"]), 5_000_000)
        self.assertEqual(round(m_nt["opening_exact"]["1052"]), 8_000_000)
        # Phát sinh trong năm tài chính 2020 (04/2020..03/2021): GTGT NTNN = 10tr, TNDN NTNN = 12tr
        self.assertEqual(round(m_nt["gtgt_tk_by_ky"][2020]), 10_000_000)
        self.assertEqual(round(m_nt["tndn_tk_qt_by_ky"][2020]), 12_000_000)
        # Đã nộp trong năm tài chính 2020 (từ tháng 4/2020): 5tr GTGT + 8tr TNDN (nộp kỳ trước sang)
        self.assertEqual(round(m_nt["gtgt_c2_by_ht"][2020]), 5_000_000)
        self.assertEqual(round(m_nt["tndn_c2_by_ht"][2020]), 8_000_000)

    def test_v16_ntnn_tt80_15_column_verbatim_layout(self) -> None:
        """Kiểm thử v1.6: Đọc chính xác 100% mẫu 01/NTNN theo TT80/2021 (15 cột: (1a), (1b), (2)..(14)) và xuất nguyên bản không lệch cột."""
        from excel_exporter import export_to_excel
        from tax_parser import parse_tax_xml

        # Dữ liệu giả lập ẩn danh 100% theo đúng cấu trúc TT80/2021 (maTKhai=884, 15 cột (1a)..(14))
        xml_tt80 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <HSoKhaiThue>
    <TTinChung>
      <TTinTKhaiThue>
        <TKhaiThue>
          <maTKhai>884</maTKhai>
          <tenTKhai>Tờ khai thuế nhà thầu nước ngoài (01/NTNN)</tenTKhai>
          <moTaBMau>Ban hành kèm theo Thông tư số 80/2021/TT-BTC</moTaBMau>
          <loaiTKhai>C</loaiTKhai>
          <soLan>0</soLan>
          <KyKKhaiThue>
            <kieuKy>D</kieuKy>
            <kyKKhai>07/01/2022</kyKKhai>
            <kyKKhaiTuNgay>07/01/2022</kyKKhaiTuNgay>
            <kyKKhaiDenNgay>07/01/2022</kyKKhaiDenNgay>
          </KyKKhaiThue>
          <ngayLapTKhai>2022-01-14</ngayLapTKhai>
        </TKhaiThue>
        <NNT>
          <mst>0100000099</mst>
          <tenNNT>CONG TY MAU TT80</tenNNT>
        </NNT>
      </TTinTKhaiThue>
    </TTinChung>
    <CTieuTKhaiChinh>
      <BKThueNTNN>
        <ThueNTNN id="ID_1">
          <ct1a>GLOBAL TECH TEST PTE. LTD</ct1a>
          <ct1b>Thu nhập bản quyền</ct1b>
          <ct2 xsi:nil="true"/>
          <ct3>HD-2021-999</ct3>
          <ct4>2021-11-23</ct4>
          <ct5>80000000</ct5>
          <ct6>2022-01-07</ct6>
          <ThueGTGT>
            <ct7>80000000</ct7>
            <ct8>0</ct8>
            <ct9>0</ct9>
          </ThueGTGT>
          <ThueTNDN>
            <ct10>80000000</ct10>
            <ct11>10</ct11>
            <ct12>0</ct12>
            <ct13>8000000</ct13>
          </ThueTNDN>
          <ct14>8000000</ct14>
        </ThueNTNN>
      </BKThueNTNN>
      <tong_ct5>80000000</tong_ct5>
      <tong_ct7>80000000</tong_ct7>
      <tong_ct9>0</tong_ct9>
      <tong_ct10>80000000</tong_ct10>
      <tong_ct12>0</tong_ct12>
      <tong_ct13>8000000</tong_ct13>
      <tong_ct14>8000000</tong_ct14>
    </CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>
"""
        with tempfile.TemporaryDirectory() as tmpdir:
            p = os.path.join(tmpdir, "ntnn_tt80_sample.xml")
            with open(p, "w", encoding="utf-8") as f:
                f.write(xml_tt80)
            d = parse_tax_xml(p)
            self.assertEqual(d.tax_type, "NTNN")
            self.assertEqual(d.indicators["_ntnn_dt_chua_gtgt"], 80_000_000)
            self.assertEqual(d.indicators["_ntnn_dt_gtgt"], 80_000_000)
            self.assertEqual(d.indicators["_ntnn_thue_gtgt"], 0)
            self.assertEqual(d.indicators["_ntnn_dt_tndn"], 80_000_000)
            self.assertEqual(d.indicators["_ntnn_thue_tndn"], 8_000_000)
            self.assertEqual(d.indicators["_ntnn_tong_thue"], 8_000_000)

            out_x = os.path.join(tmpdir, "out_tt80.xlsx")
            export_to_excel([d], out_x)
            wb = openpyxl.load_workbook(out_x, data_only=False)
            ws_th = wb["So Thue NTNN TH"]
            # Phải tự nhận diện và xuất đủ 15 cột nguyên bản TT80 (cột 9..23) + Trạng thái (cột 24)
            self.assertEqual(ws_th.max_column, 24)
            self.assertEqual(ws_th.cell(1, 9).value, "Tên nhà thầu (1a)")
            self.assertEqual(ws_th.cell(1, 10).value, "Tên ngành nghề (1b)")
            self.assertEqual(ws_th.cell(2, 9).value, "GLOBAL TECH TEST PTE. LTD")  # (1a)
            self.assertEqual(ws_th.cell(2, 10).value, "Thu nhập bản quyền")        # (1b)
            self.assertEqual(ws_th.cell(2, 11).value or "", "")                     # (2) MST trống (KHÔNG lấy nhầm số HĐ!)
            self.assertEqual(ws_th.cell(2, 12).value, "HD-2021-999")                # (3) Số HĐ
            self.assertEqual(ws_th.cell(2, 13).value, "2021-11-23")                 # (4) Ngày HĐ
            self.assertEqual(ws_th.cell(2, 14).value, 80_000_000)                   # (5) DT chưa gồm GTGT
            self.assertEqual(ws_th.cell(2, 15).value, "2022-01-07")                 # (6) Ngày thanh toán
            self.assertEqual(ws_th.cell(2, 16).value, 80_000_000)                   # (7) DT GTGT
            self.assertEqual(ws_th.cell(2, 17).value, 0)                            # (8) Tỷ lệ GTGT
            self.assertEqual(ws_th.cell(2, 18).value, 0)                            # (9) Thuế GTGT
            self.assertEqual(ws_th.cell(2, 19).value, 80_000_000)                   # (10) DT TNDN
            self.assertEqual(ws_th.cell(2, 20).value, 10)                           # (11) Tỷ lệ TNDN
            self.assertEqual(ws_th.cell(2, 21).value, 0)                            # (12) Miễn giảm
            self.assertEqual(ws_th.cell(2, 22).value, 8_000_000)                    # (13) Thuế TNDN
            self.assertEqual(ws_th.cell(2, 23).value, 8_000_000)                    # (14) Tổng thuế NSNN

    def test_v16_audit_analytics_sheets_and_weekend_deadline(self) -> None:
        """Kiểm thử v1.6: 3 sheet phụ trợ độc lập (KiemSoat_GTGT, PhanTich_RuiRo, GDLK_30EBITDA) + lùi hạn nộp cuối tuần."""
        from excel_exporter import export_to_excel
        from penalty_calculator import adjust_start_date_for_weekend
        from tax_parser import TaxDeclaration

        # 1. Kiểm tra lùi hạn nộp rơi vào Thứ 7 / Chủ nhật
        # 20/01/2024 là Thứ 7 -> hạn nộp Thứ 2 (22/01) -> tính chậm nộp từ Thứ 3 (23/01/2024)
        self.assertEqual(adjust_start_date_for_weekend(datetime.date(2024, 1, 21)), datetime.date(2024, 1, 23))
        # 20/08/2023 là Chủ nhật -> hạn nộp Thứ 2 (21/08) -> tính chậm nộp từ Thứ 3 (22/08/2023)
        self.assertEqual(adjust_start_date_for_weekend(datetime.date(2023, 8, 21)), datetime.date(2023, 8, 22))
        # 20/06/2023 là Thứ 3 -> giữ nguyên 21/06/2023
        self.assertEqual(adjust_start_date_for_weekend(datetime.date(2023, 6, 21)), datetime.date(2023, 6, 21))

        # 2. Tạo bộ tờ khai giả lập ẩn danh 100% (2 kỳ GTGT có sai đầu kỳ + KCT, 1 TNDN KQKD, 1 BCTC có LCTTGT)
        d_gt1 = TaxDeclaration(
            file_path="gt_01_2023.xml", file_name="gt_01_2023.xml",
            mst="0100000077", ten_nnt="CTY ANALYTICS TEST",
            ma_tkhai="01/GTGT", ten_tkhai="01/GTGT", tax_type="GTGT",
            ky_kkhai="01/2023", tu_ngay=datetime.date(2023, 1, 1), den_ngay=datetime.date(2023, 1, 31),
            so_lan="0", ngay_lap=datetime.date(2023, 2, 15),
            indicators={"ct22": 10_000_000, "ct24": 5_000_000, "ct25": 5_000_000, "ct26": 0, "ct34": 100_000_000, "ct43": 12_000_000},
        )
        # Kỳ 02/2023: ct22=11tr != ct43 kỳ trước (12tr) => Sai đầu kỳ; ct26=20tr > 0 & ct24=ct25=4tr => Cảnh báo phân bổ
        d_gt2 = TaxDeclaration(
            file_path="gt_02_2023.xml", file_name="gt_02_2023.xml",
            mst="0100000077", ten_nnt="CTY ANALYTICS TEST",
            ma_tkhai="01/GTGT", ten_tkhai="01/GTGT", tax_type="GTGT",
            ky_kkhai="02/2023", tu_ngay=datetime.date(2023, 2, 1), den_ngay=datetime.date(2023, 2, 28),
            so_lan="0", ngay_lap=datetime.date(2023, 3, 15),
            indicators={"ct22": 11_000_000, "ct24": 4_000_000, "ct25": 4_000_000, "ct26": 20_000_000, "ct34": 100_000_000, "ct43": 9_000_000},
        )
        d_tndn = TaxDeclaration(
            file_path="tndn_2023.xml", file_name="tndn_2023.xml",
            mst="0100000077", ten_nnt="CTY ANALYTICS TEST",
            ma_tkhai="03/TNDN", ten_tkhai="03/TNDN", tax_type="TNDN",
            ky_kkhai="2023", tu_ngay=datetime.date(2023, 1, 1), den_ngay=datetime.date(2023, 12, 31),
            so_lan="0", ngay_lap=datetime.date(2024, 3, 25),
            indicators={"A1": 500_000_000, "E": 100_000_000},
            appendices={
                "PL_03_1A": [
                    {"ma_so": "01", "chi_tieu": "Doanh thu", "gia_tri": 10_000_000_000},
                    {"ma_so": "10", "chi_tieu": "Doanh thu thuan", "gia_tri": 10_000_000_000},
                    {"ma_so": "11", "chi_tieu": "Gia von", "gia_tri": 8_000_000_000},
                    {"ma_so": "21", "chi_tieu": "DTTC", "gia_tri": 50_000_000},
                    {"ma_so": "22", "chi_tieu": "CPTC", "gia_tri": 600_000_000},
                    {"ma_so": "23", "chi_tieu": "CP lai vay", "gia_tri": 600_000_000},
                    {"ma_so": "25", "chi_tieu": "CPBH", "gia_tri": 500_000_000},
                    {"ma_so": "26", "chi_tieu": "CPQLDN", "gia_tri": 450_000_000},
                    {"ma_so": "30", "chi_tieu": "LN thuan", "gia_tri": 500_000_000},
                ]
            },
        )
        d_bctc = TaxDeclaration(
            file_path="bctc_2023.xml", file_name="bctc_2023.xml",
            mst="0100000077", ten_nnt="CTY ANALYTICS TEST",
            ma_tkhai="BCTC", ten_tkhai="BCTC", tax_type="BCTC",
            ky_kkhai="2023", tu_ngay=datetime.date(2023, 1, 1), den_ngay=datetime.date(2023, 12, 31),
            so_lan="0", ngay_lap=datetime.date(2024, 3, 25),
            indicators={
                "131_cuoi": 1_200_000_000,
                "311_cuoi": 900_000_000,
                "320_cuoi": 5_000_000_000,
                "330_cuoi": 4_000_000_000,
                "338_cuoi": 3_000_000_000,
                "410_cuoi": 10_000_000_000,
                "411_cuoi": 10_000_000_000,
            },
            appendices={
                "PL_LCTTGT": [{"ma_so": "02", "chi_tieu": "Khau hao TSCD", "nam_nay": 400_000_000}]
            },
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_x = os.path.join(tmpdir, "out_analytics.xlsx")
            export_to_excel([d_gt1, d_gt2, d_tndn, d_bctc], out_x)
            wb = openpyxl.load_workbook(out_x, data_only=False)
            self.assertIn("KiemSoat_GTGT", wb.sheetnames)
            self.assertIn("PhanTich_RuiRo", wb.sheetnames)
            self.assertIn("GDLK_30EBITDA", wb.sheetnames)

            # Kiểm tra công thức trên KiemSoat_GTGT (2 dòng kỳ: row 4, row 5; tổng cộng row 6)
            ws_ks = wb["KiemSoat_GTGT"]
            self.assertIn("IF(ABS(D5-E5)>1", str(ws_ks["F5"].value))
            self.assertEqual(ws_ks["G5"].value, "=D5-E5")
            self.assertIn("Cần KT phân bổ", str(ws_ks["M5"].value))

            # Kiểm tra công thức trên PhanTich_RuiRo (phải link thực sự sang 'BCHDKD' cho TT80 [04], [13], [14], [15], KHÔNG bị 0!)
            ws_rr = wb["PhanTich_RuiRo"]
            self.assertEqual(ws_rr["B5"].value, 2023)
            self.assertIn("'BCHDKD'!", str(ws_rr["C5"].value))
            self.assertIn("'BCHDKD'!", str(ws_rr["F5"].value))
            self.assertIn("'BCHDKD'!", str(ws_rr["I5"].value))
            self.assertIn("'BCHDKD'!", str(ws_rr["J5"].value))
            self.assertEqual(ws_rr["E5"].value, "=C5-D5")
            self.assertIn("F5/E5", str(ws_rr["K5"].value))

            # Kiểm tra công thức trên GDLK_30EBITDA (LCTTGT khấu hao = 400tr, LN thuần & Lãi vay link sang 'BCHDKD'!)
            ws_gd = wb["GDLK_30EBITDA"]
            self.assertIn("CÓ YẾU TỐ GDLK", str(ws_gd["C13"].value))
            self.assertIn("'BCHDKD'!", str(ws_gd["C17"].value))  # LN thuần [18]
            self.assertIn("'BCHDKD'!", str(ws_gd["C18"].value))  # CP lãi vay [17]
            self.assertEqual(ws_gd["C21"].value, 400_000_000)    # Khấu hao lấy trực tiếp từ LCTTGT [02]
            self.assertIn("0.3*C22", str(ws_gd["C23"].value))    # 30% EBITDA
            self.assertEqual(ws_gd["C26"].value, 0)              # Năm đầu chuyển kỳ trước = 0
            self.assertEqual(ws_gd["C27"].value, "=MAX(0,C24-C26)")  # Dư chuyển sang kỳ sau (5 năm)

    def test_v16_detect_fiscal_month_and_ntnn_tq_date_formats(self) -> None:
        """Kiểm thử v1.6: Nhận diện đúng tháng bắt đầu lệch kỳ ngay cả khi kỳ đầu là tháng 01 (thuộc NTC trước) & quy đổi đúng mọi định dạng kỳ ngày trên TQ_NTNN."""
        from qhns_summarizer import detect_fiscal_start_month, extract_ky_fiscal_year

        # 1. Workbook có 'So Thue GTGT Cuoi (Lech ky)' bắt đầu bằng 2001 (T1/2020 thuộc NTC 2019), nhưng nhãn tổng là (04/2020-03/2021)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "So Thue GTGT Cuoi (Lech ky)"
        ws.cell(3, 5, "CTY LECH KY")
        ws.cell(3, 7, "2001")
        ws.cell(4, 5, "Tổng cộng năm 2019")
        ws.cell(5, 5, "CTY LECH KY")
        ws.cell(5, 7, "2004")
        ws.cell(6, 5, "Tổng cộng năm 2020 (04/2020-03/2021)")
        self.assertEqual(detect_fiscal_start_month(wb), 4)

        # 2. Mọi định dạng kỳ ngày/tháng phổ biến trên Sổ Tổng Quan TMS (TQ_NTNN) khi lệch kỳ tháng 4
        self.assertEqual(extract_ky_fiscal_year("210215", datetime.datetime(2021, 2, 20), 4), 2020)
        self.assertEqual(extract_ky_fiscal_year("200215", datetime.datetime(2020, 2, 20), 4), 2019)
        self.assertEqual(extract_ky_fiscal_year("20210215", datetime.datetime(2021, 2, 20), 4), 2020)
        self.assertEqual(extract_ky_fiscal_year("2021-02-15 00:00:00", None, 4), 2020)
        self.assertEqual(extract_ky_fiscal_year(datetime.datetime(2021, 2, 15), None, 4), 2020)
        self.assertEqual(extract_ky_fiscal_year("2021/02", None, 4), 2020)
        self.assertEqual(extract_ky_fiscal_year("210415", datetime.datetime(2021, 4, 20), 4), 2021)
        # Kỳ 03/2024 (nộp/hạch toán tháng 04/2024) ở dạng float 2403.0 / hậu tố 2403L1 / 'PS 03/2024' phải về NTC 2023, KHÔNG nhảy sang 2024!
        self.assertEqual(extract_ky_fiscal_year(2403.0, datetime.datetime(2024, 4, 15), 4), 2023)
        self.assertEqual(extract_ky_fiscal_year("2403.0", datetime.datetime(2024, 4, 15), 4), 2023)
        self.assertEqual(extract_ky_fiscal_year("2403L1", datetime.datetime(2024, 4, 15), 4), 2023)
        self.assertEqual(extract_ky_fiscal_year("PS 03/2024", datetime.datetime(2024, 4, 15), 4), 2023)

        # 3. Kiểm tra detect_tndn_fiscal_ranges cho cả 2 tình huống:
        #    (a) Có năm sau (2024) chồng lấn -> tự cắt năm 2023 về 31/03/2024
        #    (b) Doanh nghiệp giải thể (không có 2024), gộp <= 15 tháng (04/2023-06/2024) -> GIỮ NGUYÊN (01/04/2023, 30/06/2024)
        from excel_exporter import detect_tndn_fiscal_ranges
        from qhns_summarizer import compute_tq_metrics, detect_fiscal_ranges_from_workbook
        from tax_parser import TaxDeclaration
        d22 = TaxDeclaration(file_path="t22.xml", file_name="t22.xml", mst="0100000011", ten_nnt="CTY X", ma_tkhai="03/TNDN", ten_tkhai="03/TNDN", tax_type="TNDN", ky_kkhai="2022", tu_ngay=datetime.date(2022, 4, 1), den_ngay=datetime.date(2023, 3, 31))
        d23 = TaxDeclaration(file_path="t23.xml", file_name="t23.xml", mst="0100000011", ten_nnt="CTY X", ma_tkhai="03/TNDN", ten_tkhai="03/TNDN", tax_type="TNDN", ky_kkhai="2023", tu_ngay=datetime.date(2023, 4, 1), den_ngay=datetime.date(2024, 6, 30))
        d24 = TaxDeclaration(file_path="t24.xml", file_name="t24.xml", mst="0100000011", ten_nnt="CTY X", ma_tkhai="03/TNDN", ten_tkhai="03/TNDN", tax_type="TNDN", ky_kkhai="2024", tu_ngay=datetime.date(2024, 4, 1), den_ngay=datetime.date(2025, 3, 31))
        fr = detect_tndn_fiscal_ranges([d22, d23, d24])
        self.assertEqual(fr[2023], (datetime.date(2023, 4, 1), datetime.date(2024, 3, 31)))
        self.assertEqual(fr[2024], (datetime.date(2024, 4, 1), datetime.date(2025, 3, 31)))
        # Giải thể gộp 15 tháng (không có tờ khai 2024): giữ nguyên 01/04/2023 -> 30/06/2024!
        fr_dissolve = detect_tndn_fiscal_ranges([d22, d23])
        self.assertEqual(fr_dissolve[2023], (datetime.date(2023, 4, 1), datetime.date(2024, 6, 30)))

        # 4. Kiểm tra phát sinh (TK) gộp theo năm giải thể 15 tháng (04/2023-06/2024) & số nộp (C2) đồng nhất theo năm dương lịch (01/01-31/12)
        ws.cell(7, 5, "Tổng cộng năm 2023 (04/2023-06/2024)")
        wb_ranges = detect_fiscal_ranges_from_workbook(wb)
        self.assertEqual(wb_ranges[2023], (datetime.date(2023, 4, 1), datetime.date(2024, 6, 30)))
        self.assertEqual(extract_ky_fiscal_year("2405", datetime.datetime(2024, 5, 20), 4, fiscal_ranges=wb_ranges), 2023)
        self.assertEqual(extract_ky_fiscal_year("240615", datetime.datetime(2024, 6, 20), 4, fiscal_ranges=wb_ranges), 2023)

        ntnn_dissolve_rows = [
            {"ky": "2310", "ky_year": 2023, "gd": "TK", "tm": "1052", "so_tien": 12_000_000.0, "ngay_ht": datetime.datetime(2023, 11, 15), "ht_year": 2023, "no_co": ""},
            {"ky": "2310", "ky_year": 2023, "gd": "C2", "tm": "1052", "so_tien": -12_000_000.0, "ngay_ht": datetime.datetime(2023, 11, 15), "ht_year": 2023, "no_co": ""},
            {"ky": "2402", "ky_year": 2024, "gd": "TK", "tm": "1052", "so_tien": 10_000_000.0, "ngay_ht": datetime.datetime(2024, 3, 15), "ht_year": 2024, "no_co": ""},
            {"ky": "2402", "ky_year": 2024, "gd": "C2", "tm": "1052", "so_tien": -10_000_000.0, "ngay_ht": datetime.datetime(2024, 3, 15), "ht_year": 2024, "no_co": ""},
            # Tháng 05/2024 (nằm trong kỳ gộp giải thể 04/2023-06/2024): phát sinh TK về năm 2023; số nộp C2 hạch toán 2024 thuộc năm dương lịch 2024!
            {"ky": "2405", "ky_year": 2024, "gd": "TK", "tm": "1052", "so_tien": 7_000_000.0, "ngay_ht": datetime.datetime(2024, 5, 25), "ht_year": 2024, "no_co": ""},
            {"ky": "2405", "ky_year": 2024, "gd": "C2", "tm": "1052", "so_tien": -7_000_000.0, "ngay_ht": datetime.datetime(2024, 5, 25), "ht_year": 2024, "no_co": ""},
        ]
        m_dis = compute_tq_metrics(ntnn_dissolve_rows, 2023, 2024, fiscal_start_month=4, is_ntnn=True, fiscal_ranges=wb_ranges)
        # Phát sinh lệch kỳ 15 tháng 2023 = 12tr + 10tr + 7tr = 29tr; phát sinh 2024 = 0
        self.assertEqual(round(m_dis["tndn_tk_qt_by_ky"][2023]), 29_000_000)
        self.assertEqual(round(m_dis["tndn_tk_qt_by_ky"][2024]), 0)
        # Số nộp NSNN đồng nhất theo năm dương lịch: 2023 = 12tr; 2024 = 10tr + 7tr = 17tr
        self.assertEqual(round(m_dis["tndn_c2_by_ht"][2023]), 12_000_000)
        self.assertEqual(round(m_dis["tndn_c2_by_ht"][2024]), 17_000_000)

    def test_v16_selected_sheets_filtering(self) -> None:
        """Kiểm thử v1.6: Tích chọn sheet tùy chọn (tờ khai gốc & biểu từng năm luôn mặc định xuất)."""
        from excel_exporter import export_to_excel
        from tax_parser import TaxDeclaration

        d_gt = TaxDeclaration(
            file_path="gt_01_2023.xml", file_name="gt_01_2023.xml",
            mst="0100000055", ten_nnt="CTY SELECT TEST",
            ma_tkhai="01/GTGT", ten_tkhai="01/GTGT", tax_type="GTGT",
            ky_kkhai="01/2023", tu_ngay=datetime.date(2023, 1, 1), den_ngay=datetime.date(2023, 1, 31),
            so_lan="0", ngay_lap=datetime.date(2023, 2, 15),
            indicators={"ct22": 0, "ct24": 1_000_000, "ct25": 1_000_000, "ct34": 100_000_000, "ct43": 0},
        )
        d_tndn = TaxDeclaration(
            file_path="tndn_2023.xml", file_name="tndn_2023.xml",
            mst="0100000055", ten_nnt="CTY SELECT TEST",
            ma_tkhai="03/TNDN", ten_tkhai="03/TNDN", tax_type="TNDN",
            ky_kkhai="2023", tu_ngay=datetime.date(2023, 1, 1), den_ngay=datetime.date(2023, 12, 31),
            so_lan="0", ngay_lap=datetime.date(2024, 3, 25),
            indicators={"A1": 200_000_000, "E": 40_000_000},
            appendices={"PL_03_1A": [{"ma_so": "01", "chi_tieu": "Doanh thu", "gia_tri": 1_000_000_000}]},
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            out_x = os.path.join(tmpdir, "out_selected.xlsx")
            # Chỉ tích chọn PhanTich_RuiRo & GDLK_30EBITDA (bỏ Tong_Hop, TH, Chenh lech, Cham nop, Bao cao, KiemSoat_GTGT)
            export_to_excel([d_gt, d_tndn], out_x, selected_sheets={"phan_tich_rui_ro", "gdlk_ebitda"})
            wb = openpyxl.load_workbook(out_x, data_only=False)
            names = set(wb.sheetnames)
            # Tờ khai gốc + Biểu từng năm (2023) PHẢI luôn có mặc định
            self.assertIn("So Thue GTGT Cuoi", names)
            self.assertIn("KKTNDN", names)
            self.assertIn("BCHDKD", names)
            self.assertIn("2023", names)
            # 2 sheet tích chọn có mặt
            self.assertIn("PhanTich_RuiRo", names)
            self.assertIn("GDLK_30EBITDA", names)
            # Các sheet bỏ tích KHÔNG xuất ra
            self.assertNotIn("Tong_Hop", names)
            self.assertNotIn("So Thue GTGT TH", names)
            self.assertNotIn("Chenh lech thue GTGT - TNDN", names)
            self.assertNotIn("Tinh_TienChamNop", names)
            self.assertNotIn("BaoCao_BienBan", names)
            self.assertNotIn("KiemSoat_GTGT", names)

    def test_v16_tncn_2022_pp_only_december(self) -> None:
        """Kiểm thử v1.6: Phát sinh TNCN năm 2022 trên TQ = TK + PP(Tháng 12) + TQ + TH (không cộng PP tháng 01..11/2022)."""
        from qhns_summarizer import compute_tq_metrics

        tq_rows = [
            # Năm 2022: TK (100tr), PP T06 (15tr - BỎ), PP T11 (20tr - BỎ), PP T12 (30tr - CỘNG), TQ (5tr - CỘNG), TH (2tr - CỘNG)
            {"ky": "2206", "ky_year": 2022, "gd": "TK", "tm": "1001", "so_tien": 100_000_000.0, "ngay_ht": datetime.datetime(2022, 7, 20), "ht_year": 2022, "no_co": ""},
            {"ky": "2206", "ky_year": 2022, "gd": "PP", "tm": "1001", "so_tien": 15_000_000.0, "ngay_ht": datetime.datetime(2022, 7, 20), "ht_year": 2022, "no_co": ""},
            {"ky": "2211", "ky_year": 2022, "gd": "PP", "tm": "1001", "so_tien": 20_000_000.0, "ngay_ht": datetime.datetime(2022, 12, 20), "ht_year": 2022, "no_co": ""},
            {"ky": "2212", "ky_year": 2022, "gd": "PP", "tm": "1001", "so_tien": 30_000_000.0, "ngay_ht": datetime.datetime(2023, 1, 20), "ht_year": 2023, "no_co": ""},
            {"ky": "22CN", "ky_year": 2022, "gd": "TQ", "tm": "1001", "so_tien": 5_000_000.0, "ngay_ht": datetime.datetime(2023, 3, 25), "ht_year": 2023, "no_co": ""},
            {"ky": "22CN", "ky_year": 2022, "gd": "TH", "tm": "1001", "so_tien": 2_000_000.0, "ngay_ht": datetime.datetime(2023, 3, 25), "ht_year": 2023, "no_co": ""},
            # Năm 2023: PP các tháng đều cộng bình thường (TK 50tr + PP T05 10tr = 60tr)
            {"ky": "2305", "ky_year": 2023, "gd": "TK", "tm": "1001", "so_tien": 50_000_000.0, "ngay_ht": datetime.datetime(2023, 6, 20), "ht_year": 2023, "no_co": ""},
            {"ky": "2305", "ky_year": 2023, "gd": "PP", "tm": "1001", "so_tien": 10_000_000.0, "ngay_ht": datetime.datetime(2023, 6, 20), "ht_year": 2023, "no_co": ""},
        ]
        m = compute_tq_metrics(tq_rows, 2022, 2023)
        # 2022 = 100tr (TK) + 30tr (PP T12) + 5tr (TQ) + 2tr (TH) = 137tr (KHÔNG cộng 15tr & 20tr của PP T06, T11!)
        self.assertEqual(round(m["tncn_tk_th_tq_by_ky"][2022]), 137_000_000)
        # 2023 = 50tr + 10tr = 60tr
        self.assertEqual(round(m["tncn_tk_th_tq_by_ky"][2023]), 60_000_000)


if __name__ == "__main__":
    unittest.main()
