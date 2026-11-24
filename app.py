# -*- coding: utf-8 -*-
"""
Ứng dụng Desktop chuyển đổi dữ liệu Tờ khai thuế XML (GTGT, TNCN, TNDN) ra file Excel.
"""

from __future__ import annotations

import datetime
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List

from excel_exporter import export_to_excel
from penalty_calculator import parse_cutoff_date
from qhns_summarizer import integrate_tq_and_qhns_into_workbook
from tax_parser import TaxDeclaration, parse_tax_xml


APP_VERSION = "v1.6"


class TaxXmlToExcelApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"Chuyển đổi Tờ khai Thuế XML & Đối chiếu QHNS (GTGT, TNCN, TNDN, BCTC, NTNN, TQ) - {APP_VERSION}")
        self.root.geometry("1140x760")
        self.root.minsize(980, 620)

        self.declarations: List[TaxDeclaration] = []
        self.file_paths_set: set[str] = set()

        self._setup_style()
        self._build_ui()

    def _setup_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1F4E78")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self) -> None:
        # Thanh tiêu đề
        top_frame = ttk.Frame(self.root, padding=(12, 10, 12, 6))
        top_frame.pack(fill=tk.X)

        ttk.Label(
            top_frame,
            text=f"CÔNG CỤ CHUYỂN ĐỔI TỜ KHAI THUẾ XML SANG EXCEL ({APP_VERSION})",
            style="Header.TLabel",
        ).pack(side=tk.LEFT)

        # Khung nút chọn file / thư mục
        btn_frame = ttk.LabelFrame(self.root, text="1. Chọn dữ liệu đầu vào (File XML HTKK / eTax)", padding=10)
        btn_frame.pack(fill=tk.X, padx=12, pady=4)

        ttk.Button(btn_frame, text="Chọn nhiều file XML...", command=self.add_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="Chọn thư mục (quét cả thư mục con)...", command=self.add_folder).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_frame, text="Xóa mục đang chọn", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_frame, text="Xóa tất cả", command=self.clear_all).pack(side=tk.LEFT)

        self.lbl_count = ttk.Label(btn_frame, text="Tổng: 0 file", font=("Segoe UI", 9, "bold"))
        self.lbl_count.pack(side=tk.RIGHT)

        # Khung chọn riêng Tờ khai Thuế Nhà Thầu XML (MST nộp thay riêng biệt)
        ntnn_xml_frame = ttk.LabelFrame(
            self.root,
            text="1b. Tờ khai Thuế Nhà Thầu XML (Chọn riêng - MST nộp thay riêng: tự động lên So Thue NTNN TH / Cuoi & link sang từng năm)",
            padding=(10, 6),
        )
        ntnn_xml_frame.pack(fill=tk.X, padx=12, pady=4)

        ttk.Button(ntnn_xml_frame, text="Chọn nhiều file XML Nhà thầu...", command=self.add_ntnn_files).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(ntnn_xml_frame, text="Chọn thư mục XML Nhà thầu...", command=self.add_ntnn_folder).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(ntnn_xml_frame, text="Xóa XML Nhà thầu", command=self.clear_ntnn_files).pack(side=tk.LEFT)
        self.lbl_ntnn_count = ttk.Label(ntnn_xml_frame, text="Nhà thầu (NTNN): 0 file", font=("Segoe UI", 9, "bold"))
        self.lbl_ntnn_count.pack(side=tk.RIGHT)

        # Khung chọn Sổ Tổng Quan (TQ) tùy chọn để đối chiếu & lên tự động biểu QHNS
        tq_frame = ttk.LabelFrame(
            self.root,
            text="1c. Sổ Tổng Quan (TQ - Tùy chọn: nếu chọn sẽ tự động đối chiếu F/G trên từng năm & xuất biểu QHNS)",
            padding=(10, 6),
        )
        tq_frame.pack(fill=tk.X, padx=12, pady=4)

        self.var_tq_path = tk.StringVar(value="")
        ttk.Label(tq_frame, text="File Sổ Tổng Quan (TQ):").pack(side=tk.LEFT, padx=(0, 6))
        self.ent_tq_path = ttk.Entry(tq_frame, textvariable=self.var_tq_path)
        self.ent_tq_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(tq_frame, text="Chọn file TQ (.xlsx)...", command=self.select_tq_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(tq_frame, text="Bỏ chọn", command=self.clear_tq_file).pack(side=tk.LEFT)

        # Khung chọn Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) tùy chọn để đối chiếu & lên tự động mục Thuế nhà thầu trên QHNS
        ntnn_frame = ttk.LabelFrame(
            self.root,
            text="1d. Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN - Tùy chọn: nếu chọn sẽ tự động đối chiếu F97/F98 & lên mục Thuế nhà thầu trên QHNS)",
            padding=(10, 6),
        )
        ntnn_frame.pack(fill=tk.X, padx=12, pady=4)

        self.var_tq_ntnn_path = tk.StringVar(value="")
        ttk.Label(ntnn_frame, text="File Sổ TQ Nhà thầu (TQ_NTNN):").pack(side=tk.LEFT, padx=(0, 6))
        self.ent_tq_ntnn_path = ttk.Entry(ntnn_frame, textvariable=self.var_tq_ntnn_path)
        self.ent_tq_ntnn_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(ntnn_frame, text="Chọn file TQ Nhà thầu (.xlsx)...", command=self.select_tq_ntnn_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(ntnn_frame, text="Bỏ chọn", command=self.clear_tq_ntnn_file).pack(side=tk.LEFT)

        # Khung bộ lọc & tùy chọn
        opt_frame = ttk.LabelFrame(self.root, text="2. Tùy chọn lọc & xuất Excel", padding=10)
        opt_frame.pack(fill=tk.X, padx=12, pady=4)

        row_opt1 = ttk.Frame(opt_frame)
        row_opt1.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(row_opt1, text="Loại tờ khai:").pack(side=tk.LEFT, padx=(0, 6))
        self.var_gtgt = tk.BooleanVar(value=True)
        self.var_tncn = tk.BooleanVar(value=True)
        self.var_tndn = tk.BooleanVar(value=True)
        self.var_bctc = tk.BooleanVar(value=True)
        self.var_ntnn = tk.BooleanVar(value=True)
        self.var_other = tk.BooleanVar(value=True)

        ttk.Checkbutton(row_opt1, text="GTGT", variable=self.var_gtgt, command=self.refresh_table).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Checkbutton(row_opt1, text="TNCN", variable=self.var_tncn, command=self.refresh_table).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Checkbutton(row_opt1, text="TNDN", variable=self.var_tndn, command=self.refresh_table).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Checkbutton(row_opt1, text="BCTC", variable=self.var_bctc, command=self.refresh_table).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Checkbutton(row_opt1, text="Nhà thầu (NTNN)", variable=self.var_ntnn, command=self.refresh_table).pack(
            side=tk.LEFT, padx=4
        )
        ttk.Checkbutton(row_opt1, text="Khác (LPMB...)", variable=self.var_other, command=self.refresh_table).pack(
            side=tk.LEFT, padx=(4, 14)
        )

        ttk.Separator(row_opt1, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        self.var_appendix = tk.BooleanVar(value=True)
        self.var_split = tk.BooleanVar(value=False)
        self.var_split_year = tk.BooleanVar(value=False)
        self.var_open_after = tk.BooleanVar(value=True)

        ttk.Checkbutton(row_opt1, text="Xuất kèm phụ lục (05-1, 05-2, 05-3...)", variable=self.var_appendix).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Checkbutton(row_opt1, text="Mỗi năm ra 1 file Excel riêng (VD: QTT_TNCN_2024.xlsx)", variable=self.var_split_year).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Checkbutton(row_opt1, text="Mở file sau khi xuất", variable=self.var_open_after).pack(
            side=tk.LEFT, padx=6
        )

        row_opt2 = ttk.Frame(opt_frame)
        row_opt2.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(row_opt2, text="Bố cục bảng:").pack(side=tk.LEFT, padx=(0, 6))
        self.var_layout = tk.StringVar(value="vertical")
        ttk.Radiobutton(
            row_opt2,
            text="Tổng hợp theo chiều dọc (STT - Tên chỉ tiêu - Mã chỉ tiêu | Các kỳ theo cột + Tổng cộng)",
            variable=self.var_layout,
            value="vertical",
        ).pack(side=tk.LEFT, padx=(0, 16))
        ttk.Radiobutton(
            row_opt2,
            text="Tổng hợp theo chiều ngang (Mỗi tờ khai 1 dòng)",
            variable=self.var_layout,
            value="horizontal",
        ).pack(side=tk.LEFT)

        # Hàng 3: Chọn lọc kỳ / năm cần xuất
        row_opt3 = ttk.Frame(opt_frame)
        row_opt3.pack(fill=tk.X)

        ttk.Label(row_opt3, text="Kỳ / Năm cần xuất:").pack(side=tk.LEFT, padx=(0, 6))
        self.var_filter_year = tk.BooleanVar(value=False)
        self.chk_filter_year = ttk.Checkbutton(
            row_opt3,
            text="Chỉ lấy từ năm:",
            variable=self.var_filter_year,
            command=self._toggle_year_filter,
        )
        self.chk_filter_year.pack(side=tk.LEFT, padx=(0, 4))

        self.var_year_from = tk.StringVar(value="2020")
        self.ent_year_from = ttk.Entry(row_opt3, textvariable=self.var_year_from, width=7, state=tk.DISABLED)
        self.ent_year_from.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Label(row_opt3, text="đến năm:").pack(side=tk.LEFT, padx=(0, 4))
        self.var_year_to = tk.StringVar(value="2025")
        self.ent_year_to = ttk.Entry(row_opt3, textvariable=self.var_year_to, width=7, state=tk.DISABLED)
        self.ent_year_to.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(row_opt3, text="(VD: 2020 - 2025)", font=("Segoe UI", 8, "italic"), foreground="#555555").pack(side=tk.LEFT)

        ttk.Separator(row_opt3, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Label(row_opt3, text="Ngày tính chậm nộp (dd/mm/yyyy):").pack(side=tk.LEFT, padx=(0, 4))
        self.var_cutoff_date = tk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        self.ent_cutoff_date = ttk.Entry(row_opt3, textvariable=self.var_cutoff_date, width=12)
        self.ent_cutoff_date.pack(side=tk.LEFT, padx=(0, 6))

        ttk.Separator(row_opt3, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Label(row_opt3, text="Mã CQT quản lý (QHNS):").pack(side=tk.LEFT, padx=(0, 4))
        self.var_target_cqt = tk.StringVar(value="7901")
        self.ent_target_cqt = ttk.Entry(row_opt3, textvariable=self.var_target_cqt, width=8)
        self.ent_target_cqt.pack(side=tk.LEFT, padx=(0, 4))
        ttk.Label(
            row_opt3,
            text="(VD: 7901; để trống = lấy tất cả CQT)",
            font=("Segoe UI", 8, "italic"),
            foreground="#555555",
        ).pack(side=tk.LEFT)

        # Bảng xem trước dữ liệu
        table_frame = ttk.Frame(self.root, padding=(12, 4, 12, 4))
        table_frame.pack(fill=tk.BOTH, expand=True)

        columns = ("file_name", "tax_type", "ma_tkhai", "mst", "ten_nnt", "ky_kkhai", "ngay_lap", "status")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", selectmode="extended")

        headers: Dict[str, tuple[str, int, str]] = {
            "file_name": ("Tên file XML", 190, "w"),
            "tax_type": ("Nhóm thuế", 80, "center"),
            "ma_tkhai": ("Mã tờ khai", 100, "center"),
            "mst": ("MST", 110, "center"),
            "ten_nnt": ("Tên người nộp thuế", 220, "w"),
            "ky_kkhai": ("Kỳ kê khai", 90, "center"),
            "ngay_lap": ("Ngày lập", 90, "center"),
            "status": ("Trạng thái", 130, "w"),
        }
        for col, (text, width, anchor) in headers.items():
            self.tree.heading(col, text=text)
            self.tree.column(col, width=width, anchor=anchor)  # type: ignore[arg-type]

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Thanh đáy: tiến trình & nút Xuất Excel
        bottom_frame = ttk.Frame(self.root, padding=(12, 8, 12, 12))
        bottom_frame.pack(fill=tk.X)

        self.progress = ttk.Progressbar(bottom_frame, mode="determinate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))

        self.lbl_status = ttk.Label(bottom_frame, text="Sẵn sàng", width=28)
        self.lbl_status.pack(side=tk.LEFT, padx=(0, 12))

        self.btn_export = ttk.Button(
            bottom_frame,
            text="XUẤT FILE TỔNG HỢP",
            style="Primary.TButton",
            command=self.export_excel,
            width=20,
        )
        self.btn_export.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_qhns = ttk.Button(
            bottom_frame,
            text="LÊN QHNS TỪ EXCEL + TQ",
            command=self.run_qhns_standalone,
            width=24,
        )
        self.btn_qhns.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_export_tncn = ttk.Button(
            bottom_frame,
            text="XUẤT TNCN RIÊNG THEO NĂM",
            command=self.export_tncn_separate,
            width=26,
        )
        self.btn_export_tncn.pack(side=tk.RIGHT)

    def select_tq_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file Excel Sổ Tổng Quan (TQ) tải từ TMS",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_tq_path.set(os.path.normpath(path))

    def clear_tq_file(self) -> None:
        self.var_tq_path.set("")

    def select_tq_ntnn_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file Excel Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) tải từ TMS",
            filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.var_tq_ntnn_path.set(os.path.normpath(path))

    def clear_tq_ntnn_file(self) -> None:
        self.var_tq_ntnn_path.set("")

    def _get_year_bounds(self) -> tuple[int | None, int | None]:
        if not self.var_filter_year.get():
            return None, None
        try:
            y_from = int(self.var_year_from.get().strip()) if self.var_year_from.get().strip() else None
        except ValueError:
            y_from = None
        try:
            y_to = int(self.var_year_to.get().strip()) if self.var_year_to.get().strip() else None
        except ValueError:
            y_to = None
        return y_from, y_to

    def _toggle_year_filter(self) -> None:
        if self.var_filter_year.get():
            self.ent_year_from.config(state=tk.NORMAL)
            self.ent_year_to.config(state=tk.NORMAL)
        else:
            self.ent_year_from.config(state=tk.DISABLED)
            self.ent_year_to.config(state=tk.DISABLED)
        self.refresh_table()

    def _matches_year(self, d: TaxDeclaration) -> bool:
        if not self.var_filter_year.get():
            return True
        y = d.year
        if y is None:
            return True
        try:
            y_from = int(self.var_year_from.get().strip()) if self.var_year_from.get().strip() else None
        except ValueError:
            y_from = None
        try:
            y_to = int(self.var_year_to.get().strip()) if self.var_year_to.get().strip() else None
        except ValueError:
            y_to = None

        if y_from is not None and y < y_from:
            return False
        if y_to is not None and y > y_to:
            return False
        return True

    def _allowed_types(self) -> set[str]:
        allowed = set()
        if self.var_gtgt.get():
            allowed.add("GTGT")
        if self.var_tncn.get():
            allowed.add("TNCN")
        if self.var_tndn.get():
            allowed.add("TNDN")
        if self.var_bctc.get():
            allowed.add("BCTC")
        if self.var_ntnn.get():
            allowed.add("NTNN")
        if self.var_other.get():
            allowed.add("Khác")
        return allowed

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn các file XML tờ khai thuế",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if paths:
            self._load_xml_files(list(paths))

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục chứa các file XML tờ khai thuế")
        if not folder:
            return
        xml_files: List[str] = []
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".xml"):
                    xml_files.append(os.path.join(root_dir, f))
        if not xml_files:
            messagebox.showinfo("Thông báo", "Không tìm thấy file .xml nào trong thư mục đã chọn.")
            return
        self._load_xml_files(xml_files)

    def add_ntnn_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn các file XML tờ khai Thuế Nhà Thầu (01/NTNN...)",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if paths:
            self._load_xml_files(list(paths), force_ntnn=True)

    def add_ntnn_folder(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục chứa các file XML tờ khai Thuế Nhà Thầu")
        if not folder:
            return
        xml_files: List[str] = []
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.lower().endswith(".xml"):
                    xml_files.append(os.path.join(root_dir, f))
        if not xml_files:
            messagebox.showinfo("Thông báo", "Không tìm thấy file .xml nào trong thư mục Nhà thầu đã chọn.")
            return
        self._load_xml_files(xml_files, force_ntnn=True)

    def clear_ntnn_files(self) -> None:
        remaining: List[TaxDeclaration] = []
        for d in self.declarations:
            if d.tax_type == "NTNN":
                self.file_paths_set.discard(d.file_path)
            else:
                remaining.append(d)
        self.declarations = remaining
        self.refresh_table()

    def _load_xml_files(self, paths: List[str], force_ntnn: bool = False) -> None:
        new_paths = [os.path.normpath(p) for p in paths if os.path.normpath(p) not in self.file_paths_set]
        if not new_paths:
            return

        self.progress["maximum"] = len(new_paths)
        self.progress["value"] = 0
        self.lbl_status.config(text=f"Đang đọc {len(new_paths)} file...")
        self.root.update_idletasks()

        def worker() -> None:
            parsed: List[TaxDeclaration] = []
            for idx, p in enumerate(new_paths, start=1):
                d = parse_tax_xml(p)
                if force_ntnn and d.error is None and d.tax_type != "NTNN":
                    d.tax_type = "NTNN"
                parsed.append(d)
                if idx % 10 == 0 or idx == len(new_paths):
                    self.root.after(0, lambda v=idx: self.progress.configure(value=v))

            def finish() -> None:
                for d in parsed:
                    self.file_paths_set.add(d.file_path)
                    self.declarations.append(d)
                self.refresh_table()
                self.lbl_status.config(text=f"Đã tải xong {len(new_paths)} file.")

            self.root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    def refresh_table(self) -> None:
        self.tree.delete(*self.tree.get_children())
        allowed = self._allowed_types()
        shown = 0
        dup_count = 0
        ntnn_total = sum(1 for d in self.declarations if d.error is None and d.tax_type == "NTNN")
        seen_keys: set[tuple[str, str, str, str]] = set()
        for idx, d in enumerate(self.declarations):
            if d.error is None and d.tax_type not in allowed:
                continue
            if d.error is None and not self._matches_year(d):
                continue
            shown += 1
            ngay_str = d.ngay_lap.strftime("%d/%m/%Y") if d.ngay_lap else ""
            if d.error:
                status = f"Lỗi: {d.error}"
            else:
                ky_str = (
                    d.tu_ngay.strftime("%y%m%d")
                    if (d.tax_type == "NTNN" and d.tu_ngay)
                    else (
                        d.tu_ngay.strftime("%y%m")
                        if (d.tax_type == "GTGT" and d.tu_ngay)
                        else str(d.year or d.ky_kkhai or "")
                    )
                )
                lan_str = str(d.so_lan or "0").strip()
                dedup_k = (d.tax_type or "", d.mst or "", ky_str, lan_str)
                if dedup_k in seen_keys:
                    status = "Hợp lệ (Trùng tờ khai)"
                    dup_count += 1
                else:
                    seen_keys.add(dedup_k)
                    status = "Hợp lệ"
            self.tree.insert(
                "",
                tk.END,
                iid=str(idx),
                values=(d.file_name, d.tax_type, d.ma_tkhai, d.mst, d.ten_nnt, d.ky_kkhai, ngay_str, status),
            )
        dup_note = f" (Trùng: {dup_count})" if dup_count > 0 else ""
        self.lbl_count.config(text=f"Hiển thị: {shown} / Tổng: {len(self.declarations)} file{dup_note}")
        self.lbl_ntnn_count.config(text=f"Nhà thầu (NTNN): {ntnn_total} file")

    def remove_selected(self) -> None:
        sel = self.tree.selection()
        if not sel:
            return
        indices_to_remove = sorted([int(i) for i in sel], reverse=True)
        for idx in indices_to_remove:
            removed = self.declarations.pop(idx)
            self.file_paths_set.discard(removed.file_path)
        self.refresh_table()

    def clear_all(self) -> None:
        self.declarations.clear()
        self.file_paths_set.clear()
        self.refresh_table()
        self.progress["value"] = 0
        self.lbl_status.config(text="Sẵn sàng")

    def export_excel(self) -> None:
        allowed = self._allowed_types()
        filtered = [
            d for d in self.declarations
            if (d.error is not None or (d.tax_type in allowed and self._matches_year(d)))
        ]
        valid_count = sum(1 for d in filtered if d.error is None)

        if valid_count == 0:
            messagebox.showwarning("Chú ý", "Không có tờ khai hợp lệ nào để xuất. Hãy thêm file XML trước.")
            return

        out_path = filedialog.asksaveasfilename(
            title="Lưu file Excel kết quả",
            defaultextension=".xlsx",
            initialfile="TongHop_ToKhaiThue.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not out_path:
            return

        include_app = self.var_appendix.get()
        split_files = self.var_split.get()
        split_years = self.var_split_year.get()
        open_after = self.var_open_after.get()
        layout_mode = self.var_layout.get()
        tq_path = self.var_tq_path.get().strip()
        tq_ntnn_path = self.var_tq_ntnn_path.get().strip()
        valid_ntnn = tq_ntnn_path if (tq_ntnn_path and os.path.isfile(tq_ntnn_path)) else None
        y_from, y_to = self._get_year_bounds()
        cutoff_dt = parse_cutoff_date(self.var_cutoff_date.get())
        target_cqt = self.var_target_cqt.get().strip() or None

        self.btn_export.config(state=tk.DISABLED)
        self.btn_export_tncn.config(state=tk.DISABLED)
        self.btn_qhns.config(state=tk.DISABLED)
        self.lbl_status.config(text="Đang xuất Excel...")
        self.root.update_idletasks()

        def worker() -> None:
            try:
                created = export_to_excel(
                    filtered,
                    out_path,
                    include_appendices=include_app,
                    split_by_tax_type=split_files,
                    split_by_year=split_years,
                    layout=layout_mode,
                    cutoff_date=cutoff_dt,
                )

                qhns_alerts: List[str] = []
                qhns_done = False
                if tq_path and os.path.isfile(tq_path) and created and not split_files and not split_years:
                    self.root.after(0, lambda: self.lbl_status.config(text="Đang đối chiếu TQ & lên QHNS..."))
                    qhns_res = integrate_tq_and_qhns_into_workbook(
                        created[0],
                        tq_path,
                        year_from=y_from,
                        year_to=y_to,
                        output_path=created[0],
                        cutoff_date=cutoff_dt,
                        tq_ntnn_excel_path=valid_ntnn,
                        target_cqt=target_cqt,
                    )
                    qhns_alerts = qhns_res.get("alerts", [])
                    qhns_done = True

                def on_done() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_export_tncn.config(state=tk.NORMAL)
                    self.btn_qhns.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Xuất Excel thành công!")
                    msg = "Đã xuất thành công:\n" + "\n".join(created)
                    if qhns_done:
                        msg += "\n\nĐã tích hợp Sổ Tổng Quan (TQ), đối chiếu cột F/G trên các sheet năm và lập biểu QHNS!"
                        if valid_ntnn:
                            msg += "\nĐã tích hợp thêm Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) vào QHNS & DoiChieu_QHNS!"
                        if qhns_alerts:
                            msg += "\n\nLƯU Ý PHÁT HIỆN CHÊNH LỆCH TỜ KHAI vs TỔNG QUAN:\n- " + "\n- ".join(qhns_alerts)
                    messagebox.showinfo("Hoàn tất", msg)
                    if open_after and created:
                        try:
                            os.startfile(created[0])
                        except Exception:
                            pass

                self.root.after(0, on_done)
            except Exception as exc:
                err_msg = str(exc)

                def on_err(msg: str = err_msg) -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_export_tncn.config(state=tk.NORMAL)
                    self.btn_qhns.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Lỗi khi xuất Excel!")
                    messagebox.showerror("Lỗi", f"Không thể xuất file Excel:\n{msg}")

                self.root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()

    def run_qhns_standalone(self) -> None:
        """Chế độ 2: Ghép Sổ Tổng Quan (TQ), đối chiếu F/G và lập biểu QHNS vào file Excel tổng hợp đã có sẵn."""
        excel_path = filedialog.askopenfilename(
            title="Bước 1/2: Chọn file Excel Tổng hợp tờ khai đã xuất (VD: TongHop_ToKhaiThue.xlsx)",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not excel_path:
            return

        tq_path = self.var_tq_path.get().strip()
        if not tq_path or not os.path.isfile(tq_path):
            tq_path = filedialog.askopenfilename(
                title="Bước 2/2: Chọn file Excel Sổ Tổng Quan (TQ) tải từ TMS",
                filetypes=[("Excel Workbook", "*.xlsx"), ("All files", "*.*")],
            )
            if not tq_path:
                return
            self.var_tq_path.set(os.path.normpath(tq_path))

        tq_ntnn_path = self.var_tq_ntnn_path.get().strip()
        valid_ntnn = tq_ntnn_path if (tq_ntnn_path and os.path.isfile(tq_ntnn_path)) else None

        y_from, y_to = self._get_year_bounds()
        open_after = self.var_open_after.get()
        cutoff_dt = parse_cutoff_date(self.var_cutoff_date.get())
        target_cqt = self.var_target_cqt.get().strip() or None

        self.btn_export.config(state=tk.DISABLED)
        self.btn_export_tncn.config(state=tk.DISABLED)
        self.btn_qhns.config(state=tk.DISABLED)
        self.lbl_status.config(text="Đang đối chiếu TQ & lên biểu QHNS...")
        self.root.update_idletasks()

        def worker() -> None:
            try:
                res = integrate_tq_and_qhns_into_workbook(
                    excel_path,
                    tq_path,
                    year_from=y_from,
                    year_to=y_to,
                    output_path=excel_path,
                    cutoff_date=cutoff_dt,
                    tq_ntnn_excel_path=valid_ntnn,
                    target_cqt=target_cqt,
                )

                def on_done() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_export_tncn.config(state=tk.NORMAL)
                    self.btn_qhns.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Lên biểu QHNS thành công!")
                    yrs = res.get("years", [])
                    yr_label = f"{yrs[0]}–{yrs[-1]}" if yrs else ""
                    msg = f"Đã tích hợp Sổ Tổng Quan (TQ), đối chiếu F/G và lập biểu QHNS ({yr_label}) vào file:\n{excel_path}"
                    if valid_ntnn:
                        msg += "\nĐã tích hợp thêm Sổ Tổng Quan Thuế Nhà Thầu (TQ_NTNN) vào QHNS & DoiChieu_QHNS!"
                    alerts = res.get("alerts", [])
                    if alerts:
                        msg += "\n\nLƯU Ý PHÁT HIỆN CHÊNH LỆCH TỜ KHAI vs TỔNG QUAN:\n- " + "\n- ".join(alerts)
                    messagebox.showinfo("Hoàn tất QHNS", msg)
                    if open_after:
                        try:
                            os.startfile(excel_path)
                        except Exception:
                            pass

                self.root.after(0, on_done)
            except Exception as exc:
                err_msg = str(exc)

                def on_err(msg: str = err_msg) -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_export_tncn.config(state=tk.NORMAL)
                    self.btn_qhns.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Lỗi khi lên biểu QHNS!")
                    messagebox.showerror("Lỗi", f"Không thể lập biểu QHNS:\n{msg}")

                self.root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()

    def export_tncn_separate(self) -> None:
        """Xuất riêng từng năm tờ khai quyết toán TNCN, mỗi năm 1 file gồm 4 sheet: TK, 05-1, 05-2, 05-3."""
        filtered = [
            d for d in self.declarations
            if (d.error is None and d.tax_type == "TNCN" and self._matches_year(d))
        ]
        if not filtered:
            messagebox.showwarning(
                "Chú ý",
                "Không tìm thấy tờ khai TNCN nào hợp lệ trong danh sách file (theo bộ lọc năm hiện tại)."
            )
            return

        out_dir = filedialog.askdirectory(title="Chọn thư mục lưu các file tờ khai quyết toán TNCN (theo từng năm)")
        if not out_dir:
            return

        open_after = self.var_open_after.get()
        self.lbl_status.config(text="Đang xuất các file TNCN riêng...")
        self.btn_export.config(state=tk.DISABLED)
        self.btn_export_tncn.config(state=tk.DISABLED)
        self.btn_qhns.config(state=tk.DISABLED)
        self.root.update_idletasks()

        def worker() -> None:
            try:
                from excel_exporter import write_single_year_tncn_file, filter_latest_per_year
                latest_tncn = filter_latest_per_year(filtered)
                created = []
                for d in latest_tncn:
                    yr = d.year or d.ky_kkhai
                    fn = os.path.join(out_dir, f"QTT_TNCN_{yr}.xlsx")
                    write_single_year_tncn_file(fn, d)
                    created.append(fn)

                def on_done() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_export_tncn.config(state=tk.NORMAL)
                    self.btn_qhns.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Xuất file TNCN thành công!")
                    msg = "Đã xuất thành công các file TNCN riêng (gồm đủ 4 sheet: TK, 05-1, 05-2, 05-3):\n" + "\n".join(created)
                    messagebox.showinfo("Hoàn tất", msg)
                    if open_after and created:
                        try:
                            os.startfile(out_dir)
                        except Exception:
                            pass

                self.root.after(0, on_done)
            except Exception as exc:
                err_msg = str(exc)

                def on_err(msg: str = err_msg) -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_export_tncn.config(state=tk.NORMAL)
                    self.btn_qhns.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Lỗi khi xuất file TNCN!")
                    messagebox.showerror("Lỗi", f"Không thể xuất file TNCN:\n{msg}")

                self.root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    TaxXmlToExcelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
