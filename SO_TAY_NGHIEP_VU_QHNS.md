# Sổ Tay Nghiệp Vụ: Đối Chiếu Tờ Khai – Sổ Tổng Quan (TQ) & Lập Biểu Quan Hệ Ngân Sách (QHNS)

> **Mục tiêu cốt lõi (Goal):** Mọi bảng **Quan Hệ Ngân Sách (`QHNS`)** khi lập phải gom vào **01 biểu chuẩn duy nhất**, số liệu chính xác tuyệt đối và **khớp 100% với số dư cuối cùng trên Sổ Tổng Quan (`TQ`) của năm tài chính**.

---

## 1. Các Nguyên Tắc Bất Di Bất Dịch

1. **Một biểu `QHNS` duy nhất:** Không trình bày nhiều phương án hay nhiều bảng phụ trên sheet `QHNS`. Các khoản cấu thành Số đầu kỳ hoặc kiểm tra Số dư cuối kỳ được đặt tại cột phụ `H`, `I` ngay ngang hàng với chỉ tiêu đó và có **Comment (Ghi chú)** giải thích rõ nguồn gốc nghiệp vụ.
2. **Cột (4) – Số báo cáo phát sinh phải nộp trong kỳ:** Không bao giờ có số âm.
3. **Phân biệt rõ theo loại giao dịch trên Sổ Tổng Quan (`TQ`):**
   - **Thuế GTGT (Tiểu mục `1701`):** Số phát sinh trên `TQ` lấy giao dịch **`TK`** theo **Kỳ báo cáo** của từng năm.
   - **Thuế TNDN (Tiểu mục `1052`):** Số phát sinh trên `TQ` lấy tổng giao dịch **`TK`, `TQ`, `QT`** theo **Kỳ báo cáo** của từng năm (lưu ý: tờ khai quyết toán TNDN năm như `20CN`, `25CN` trên TMS thường mang mã **`TQ`**).
   - **Thuế TNCN từ tiền lương, tiền công (Tiểu mục `1001`):** Số phát sinh trên `TQ` lấy tổng giao dịch **`TK`, `TH`, `TQ`** của đúng Tiểu mục **`1001`** theo **Kỳ báo cáo** của từng năm (không trộn tiểu mục `1015` hay `1004`).
   - **Số đã nộp ngân sách – Cột (6):** Lấy giao dịch **`C2`** theo **Năm Ngày hạch toán**, đồng thời ghi nhận cả các giao dịch điều chỉnh chứng từ nộp đặc biệt gồm **`HC`** (Hủy chứng từ) và **`NT`** (ví dụ `NT XU LY CTU AM BK NGAY ...` – xử lý chứng từ âm bảng kê) theo đúng năm ngày hạch toán và dấu đại số trên `TQ`. Tuyệt đối không lấy `PT`, `QH`, `C5`, `KT`.
4. **Quy tắc xử lý năm tính thuế & chuỗi `Kỳ báo cáo`:**
   - **Tờ khai XML (`03/TNDN`, `BCTC` có năm tài chính lệch năm dương lịch):** Bắt buộc lấy năm tính thuế ưu tiên tuyệt đối theo thẻ **`<kyKKhai>`** trước (tuyệt đối không lấy `den_ngay.year` trước `kyKKhai`) để năm `2020` (VD: `01/04/2020–31/03/2021`) không bị đẩy nhầm sang bảng năm `2021`.
   - **Trên Sổ Tổng Quan (`TQ`):** Các kỳ tháng của năm 2020 có định dạng `YYMM` là `"2001"` đến `"2012"`. Khi lập trình trích xuất năm của `Kỳ báo cáo`, bắt buộc kiểm tra mẫu `^(\d{2})(0[1-9]|1[0-2]|Q\d|CN)$` trước tiên để `"2001".."2012"` trả về năm **`2020`** (tránh lỗi regex `20\d{2}` nhận nhầm thành các năm `2001..2012`).

---

## 2. Quy Trình Bước 1: Đối Chiếu Giữa Tờ Khai Và Phát Sinh Trên Tổng Quan (`TQ`) Tại Các Sheet Năm (`2020`..`2025`)

Tại mỗi sheet năm (`2020`, `2021`, `2022`, `2023`, `2024`, `2025`), thiết lập 2 cột đối chiếu tại **Cột F** *(Số phát sinh trên Tổng quan `TQ`)* và **Cột G** *(Chênh lệch `= C - F`)*:

### 2.1. Dòng 23 (`Chỉ tiêu 11: Thuế GTGT phải nộp trong kỳ`)
- **`C23`:** Số báo cáo của người nộp thuế (link từ `'So Thue GTGT Cuoi'!AH...`).
- **`F23`** *(Comment: "Phát sinh TK trên tổng quan")*: Tổng giao dịch `TK` (Tiểu mục `1701`) theo Kỳ báo cáo thuộc năm đó trên `TQ`.
- **`G23` (`=C23-F23`):** Chênh lệch giữa Tờ khai và Tổng quan `TQ`.

### 2.2. Dòng 76 (`Mục D: Thuế TNDN phải nộp trong kỳ`)
- **`C76`:** Số báo cáo của người nộp thuế (`=SUM(C77:C79)`).
- **`F76`** *(Comment: "Phát sinh TK; QT trên tổng quan")*: Tổng giao dịch `TK + QT` (Tiểu mục `1052`) theo Kỳ báo cáo thuộc năm đó trên `TQ`.
- **`G76` (`=C76-F76`):** Chênh lệch giữa Tờ khai Quyết toán TNDN và Tổng quan `TQ`.

### 2.3. Dòng 90 (`Tổng hợp tổng số thuế TNCN phải khấu trừ`)
- **`C90` (`=C85+C87+C89`):** Tổng số thuế TNCN phải khấu trừ trên Tờ khai Quyết toán `05/QTT-TNCN` (tổng hợp từ `05-1/BK` lũy tiến + `05-1/BK` toàn phần 10% + `05-2/BK` không cư trú 20%).
- **`F90`** *(Comment: "Phát sinh TK; TH; TQ trên tổng quan")*: Tổng giao dịch `TK + TH + TQ` (Tiểu mục `1001`) theo Kỳ báo cáo thuộc năm đó trên `TQ`.
- **`G90` (`=C90-F90`):** Chênh lệch giữa Tờ khai Quyết toán TNCN và Phát sinh trên Tổng quan `TQ`.
  - **Bài học thực tế Năm 2022:**
    - Trên Tờ khai Quyết toán `05/QTT-TNCN` năm 2022: `C90 = 5,078,805,638`.
    - Trên Sổ Tổng Quan `TQ` năm 2022: `F90 = 4,243,002,612` (do trên `TQ` bị thiếu tờ khai tháng `01/2022` – kỳ `2201`, dù ngày `07/04/2022` doanh nghiệp vẫn nộp `C2` số tiền `797,956,214` cho kỳ `2201`).
    - Chênh lệch `G90 = C90 - F90 = 835,803,026` *(Comment tại `F90`: "Phát sinh TK; TH; TQ trên tổng quan rà soát thiếu TK 2201")*.

---

## 3. Quy Trình Bước 2: Quy Tắc Lên Bảng `QHNS` Từ Kết Quả Đối Chiếu

### 3.1. Cột (3) – Số kỳ trước chuyển sang (Đầu kỳ năm đầu tiên `2020`)
Số đầu kỳ của giai đoạn kiểm tra (`2020–2025`) bao gồm toàn bộ các khoản nộp (`C2`), quyết toán thừa/thiếu (`TQ`, `TH`) hoặc truy thu thanh kiểm tra (`QT`) của **các kỳ trước năm 2020 (`Kỳ báo cáo < 2020`) nhưng phát sinh nộp/hạch toán trong giai đoạn `2020–2025`**:

1. **Thuế GTGT (`C5 = H5 + I5 = 912,460,864`):**
   - **`H5 = 64,511,180`** *(Comment: "Số nộp của QT 2015 nộp năm 2021 trong kỳ kiểm tra" — kỳ `15CN` nộp năm 2021)*.
   - **`I5 = 847,949,684`** *(Comment: "Số nộp tháng 1912" — giao dịch `C2` kỳ `1912` nộp ngày `20/01/2020`)*.
2. **Thuế TNCN (`C8 = H8 + I8 = 270,242,586`):**
   - **`H8 = 295,720,173`** *(Comment: "Kỳ kk 1912 - Số nộp" — giao dịch `C2` kỳ `1912` nộp ngày `20/01/2020`)*.
   - **`I8 = -25,477,587`** *(Comment: "Số nộp thừa trên QT 2019" — cộng toàn bộ giao dịch `TQ = +1,137,823` và `TH = -26,615,410` của kỳ `19CN` hạch toán ngày `30/03/2020`)*.
3. **Các năm `2021..2025`:** Cột (3) `= Cột (7) Số còn phải nộp` của năm liền trước.

### 3.2. Cột (4) – Số báo cáo phát sinh phải nộp trong kỳ & Cột (5) – Số xử lý qua thanh tra
- **Khi Tờ khai khớp với Tổng quan `TQ`:**
  - Thuế GTGT: `D = '<Năm>'!C23`
  - Thuế TNCN: `D = '<Năm>'!C90`
- **Khi thiếu tờ khai trên Tổng quan `TQ` (Bài học điển hình: Thuế TNCN Năm 2022 – Dòng 24 trên `QHNS`):**
  - **`D24 = '2022'!F90` (`= 4,243,002,612`)** *(Comment: "Vì bị thiếu tờ khai 2201 nên số phát sinh trên tổng quan bị thiếu, nên khi thể hiện số phát sinh sẽ lấy trên tổng quan")*.
  - **`E24 = '2022'!G90` (`= 835,803,026`)** *(Comment: "Số chênh lệch thực tế trên tờ khai và phát sinh trên tổng quan do tờ khai bị thiếu")*.

### 3.3. Cột (6) – Số đã nộp (`F`)
- Lấy tổng giao dịch **`C2`** theo **Năm Ngày hạch toán** trên sheet `TQ` cho từng Tiểu mục (`1701`, `1052`, `1001`).

### 3.4. Cột (7) – Số còn phải nộp (`G = C + D + E - F`) & Đối chiếu khóa sổ cuối năm tài chính
- Tại cuối năm tài chính cuối cùng (`Năm 2025`):
  - **Thuế GTGT (`G45`):** `57,285,642`.
  - **Thuế TNCN (`G48`):** `940,214,600`.
  - Tại ô **`H48 = G48 - E24`** (`= 940,214,600 - 835,803,026 = 104,411,574`): Đối chiếu ngược lại số dư nghĩa vụ năm tài chính 2025 trên Sổ Tổng Quan `TQ` khi chưa cộng khoản xử lý qua thanh tra `E24` (`142,616,473 [TK 2512] - 37,204,969 [TQ 25CN] - 999,930 = 104,411,574`), đảm bảo khớp 100% đến từng đồng!


---
### Bổ sung Quy tắc Chuẩn hóa Số dư Đầu kỳ Tuyệt đối & Lệnh hoàn `C5`
1. **Luôn loại bỏ bút toán bù trừ nội bộ (`Ký hiệu giao dịch = BT`)**: `BT` luôn đi theo cặp Âm–Dương triệt tiêu nhau để tất toán công nợ trên TMS, tuyệt đối không tính vào phát sinh hay số đã nộp.
2. **Công thức Phân hoạch 3 nhóm (A/B/C) bảo toàn 100% Số dư cuối kỳ `QHNS` = `TQ`**:
   - **Nhóm A (Phát sinh trong kỳ kiểm tra `start_year..end_year`):** `TK` (GTGT `1701`), `TK/TQ/QT` (TNDN `1052`), `TK/TH/TQ` (TNCN `1001`) theo năm Kỳ báo cáo.
   - **Nhóm B (Số đã nộp / Hoàn trả trong kỳ kiểm tra `start_year..end_year`):** `C2`, `HC`, `NT`, **`C5` (Lệnh hoàn KBT thu NS — ghi âm giảm số nộp)** theo năm Ngày hạch toán, bảo toàn dấu đại số.
   - **Nhóm C (Số kỳ trước chuyển sang = Dư đầu kỳ `start_year`):** Toàn bộ giao dịch khác `BT` của tiểu mục không thuộc Nhóm A & B (và không vượt quá `end_year`), bao gồm số nộp thừa/còn nợ lũy kế từ các năm trước (`< start_year`), các tờ khai/QĐ truy thu kỳ trước hạch toán muộn (`1901 TK`, `15CN QT`), và các cặp `PT`/`QH` bù trừ hoàn.
   - Khi tách ra cột thuyết minh `H` & `I` của năm đầu tiên, luôn giữ `H + I = Tổng Nhóm C` để đảm bảo **Dư cuối kỳ `QHNS` khớp tuyệt đối 0 đồng chênh lệch với Sổ Tổng Quan (`TQ`)** cho cả 3 sắc thuế (GTGT, TNDN, TNCN).

3. **Đối chiếu Lệ phí môn bài trên biểu `QHNS` (Dòng `5 | Lệ phí môn bài`):**
   - Bao quát đầy đủ các tiểu mục LPMB theo Mục lục NSNN: **`2862`** (Bậc 1 – 3 triệu), **`2863`** (Bậc 2 – 2 triệu), **`2864`** (Bậc 3 / chi nhánh, ĐĐKD – 1 triệu) và các tiểu mục cũ `1801`–`1804`.
   - Áp dụng đồng nhất công thức phân hoạch 3 nhóm A/B/C: Phát sinh (`TK/TQ/QT` theo năm Kỳ báo cáo), Đã nộp (`C2/HC/NT/C5` theo Năm Ngày hạch toán), và Đầu kỳ (`opening_exact["LPMB"]`), kèm Comment thuyết minh chi tiết từng tiểu mục.

4. **Tách riêng 2 sheet chuẩn nghiệp vụ Đoàn kiểm tra thuế:**
   - **Sheet `QHNS` (Biểu để đưa vào Biên bản kiểm tra):** **KHÔNG** ghi nhận số nộp thừa lũy kế kỳ trước chuyển sang vào Cột (3) `Số kỳ trước chuyển sang` năm đầu tiên; chỉ đưa vào đầu kỳ **tất cả số nộp (`C2`/`HC`/`NT`) & `QT` của mọi kỳ trước không thuộc kỳ kiểm tra (`ky_year < start_year`) nộp/hạch toán trong kỳ kiểm tra** (không giới hạn riêng tháng 12 năm trước).
   - **Sheet `DoiChieu_QHNS` (Biểu đối chiếu số khớp thực tế với Sổ `TQ`):**
     - **Bảng 1:** Đối chiếu tổng hợp cuối kỳ nối thẳng công thức từ `QHNS` + `Số nộp thừa / ĐC đầu kỳ không đưa vào Biên bản` − `Số xử lý qua thanh tra (Cột 5)` − `Chênh lệch làm tròn` ↔ **Dư cuối thực tế trên `TQ`** (chênh lệch = `0` đồng).
     - **Bảng 2:** Diễn biến chi tiết từng năm khớp thực tế 100% theo `TQ` (bao gồm cả số nộp thừa kỳ trước chuyển sang).
