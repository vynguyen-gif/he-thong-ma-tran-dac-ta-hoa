# Antigravity Pair-Programming Guidelines for Chemistry Matrix System

## 1. Identity & Branding
- Giảng viên phụ trách / Biên soạn: **GV: NGUYỄN VỸ** — Tổ Hóa Học Trường THPT Bùi Thị Xuân (Mũi Né).
- Tiêu đề thanh Navbar: Biểu tượng bình tam giác + **Hóa Học** + Huy hiệu **GV: NGUYỄN VỸ** + Slogan *"Khám phá thế giới qua lăng kính Hóa học"*. Không chèn các cụm từ rườm rà như "GDPT 2018 - CV 7991" hay "Phụ trách GV...".

## 2. Automatic Quality Assurance Pipeline cho Ngân hàng Câu hỏi
Khi người dùng thêm ngân hàng câu hỏi mới, nhập câu hỏi mới hoặc cập nhật dữ liệu:

1. **Phần I (Trắc nghiệm nhiều phương án lựa chọn)**:
   - Tất cả các phương án A, B, C, D bắt buộc hiển thị **mỗi phương án trên 1 dòng riêng biệt**, thụt lề chuẩn (cả trên giao diện Web lẫn khi xuất file Word .doc).
   - Tuyệt đối không xếp 4 phương án trên 1 hàng ngang.

2. **Tự động cấu trúc hóa Bảng biểu trong Cả Nội Dung Đề & Lời giải / Hướng dẫn chấm**:
   - Khi phát hiện câu hỏi hoặc lời giải có dữ liệu dạng bảng (bảng khối lượng nguyên tử, bảng hạt cơ bản p/n/e, bảng thí nghiệm nhận biết thuốc thử - hiện tượng, bảng pin điện hóa, bảng thế điện cực chuẩn E0, bảng nhiệt độ sôi/nóng chảy, bảng cấu hình electron, bảng độ tan muối...), **tuyệt đối không để dạng văn bản phẳng nối đuôi nhau**.
   - Trong đề bài (`content`): Bảng để trống các ô cần điền dạng `...?...` hoặc ô trống cho học sinh làm bài.
   - Trong hướng dẫn giải (`solution`): Bảng điền đầy đủ đáp án nổi bật kèm các bước biến đổi, công thức tính toán khoa học.
   - Tự động chuyển đổi thành bảng HTML `<table>...</table>` chuẩn với viền kẻ rõ ràng (`border='1'`, `border-collapse: collapse`), màu nền thanh tiêu đề (`#f1f5f9`), căn lề chuẩn giữa Web và Word (.doc).

3. **Chuẩn hóa công thức hóa học và phương trình**:
   - Tất cả chỉ số nguyên tử, nhóm nguyên tử phải dùng chỉ số dưới (`H<sub>2</sub>O`, `CO<sub>2</sub>`, `C<sub>n</sub>(H<sub>2</sub>O)<sub>m</sub>`, `CH<sub>3</sub>COOH`...).
   - Các ion phải có điện tích đúng (`Na<sup>+</sup>`, `Fe<sup>2+</sup>`, `Al<sup>3+</sup>`, `SO<sub>4</sub><sup>2−</sup>`...).
   - Phương trình phản ứng phải có mũi tên chuẩn (`→`, `⇌`, `⎯⎯(t°)⎯→`, `⎯⎯(đpnc)⎯→`, `⎯⎯(xt)⎯→`) cùng trạng thái kết tủa/bay hơi (`↓`, `↑`).

4. **Chuẩn hóa số mũ lũy thừa & Ký hiệu khoa học**:
   - Tự động quy đổi các dạng viết phẳng của Word/OCR: `39,8271.10-27` hoặc `39,8271.10–27` -> `39,8271 · 10<sup>-27</sup>`.
   - Các đơn vị vật lý và hằng số nguyên tử: `10-27 kg` -> `10<sup>-27</sup> kg`, `10-19 C` -> `10<sup>-19</sup> C`, `10-31 kg` -> `10<sup>-31</sup> kg`, `109` -> `10<sup>9</sup>`, `104` -> `10<sup>4</sup>`.

5. **Sửa lỗi chính tả OCR / tài liệu quét**:
   - Tự động khắc phục các lỗi chính tả phổ biến từ tài liệu Word/PDF (ví dụ: `nững chất` -> `những chất`, `frucrose` -> `fructose`, xóa mã neo Word như `right317500`,...).

6. **Thống nhất & Gộp Chương / Chuyên đề chuẩn SGK Kết Nối Tri Thức (KNTT)**:
   - Tuyệt đối không để phân mảnh hay lẫn lộn các cách gọi như `Chủ đề 1: Cấu tạo nguyên tử`, `phần cấu tạo nguyên tử`, `Cấu Tạo Nguyên Tử`.
   - Tất cả các chủ đề/chuyên đề cùng nội dung phải được gộp thống nhất thành **Chương X: [Tên chuẩn]**:
     - **Khối 10 (7 chương)**: Chương 1: Cấu tạo nguyên tử; Chương 2: Bảng tuần hoàn các nguyên tố hóa học; Chương 3: Liên kết hóa học; Chương 4: Phản ứng oxi hóa - khử; Chương 5: Năng lượng hóa học; Chương 6: Tốc độ phản ứng hóa học; Chương 7: Nhóm Halogen (Nhóm VIIA).
     - **Khối 11 (6 chương)**: Chương 1: Cân bằng hóa học; Chương 2: Nitrogen và sulfur; Chương 3: Đại cương về hóa học hữu cơ; Chương 4: Hydrocarbon; Chương 5: Dẫn xuất halogen - Alcohol - Phenol; Chương 6: Hợp chất carbonyl - Carboxylic acid.
     - **Khối 12 (8 chương)**: Chương 1: Ester - Lipid; Chương 2: Carbohydrate; Chương 3: Hợp chất chứa nitrogen; Chương 4: Polymer; Chương 5: Pin điện và điện phân; Chương 6: Đại cương về kim loại; Chương 7: Nguyên tố nhóm IA và nhóm IIA; Chương 8: Kim loại chuyển tiếp dãy thứ nhất & Phức chất.
7. **Chuẩn hóa Ký hiệu Ô Orbital (Ô lượng tử / AO) thành các ô vuông liên kết chuẩn SGK Hóa học**:
   - Trong hóa học, orbital nguyên tử (AO) được biểu diễn bằng các **ô vuông** (không để dạng ngoặc thô `[...]`).
   - Tự động chuyển đổi các ký hiệu `[↑]`, `[↓]`, `[↑↓]`, `[↑↑]`, `[ ]` (kể cả ô trống) thành các ô vuông liền kề nhau với viền kẻ màu đen rõ nét (1.5pt solid black), kích thước vuông vắn 20pt × 20pt, căn giữa mũi tên spin electron (`↑↓`, `↑ `, `  `).
   - **Tối ưu hiển thị kép (Web & Word)**:
     + Trên giao diện Web: hiển thị ô vuông liền kề (`table.ao-boxes` hoặc `.ao-group`), căn giữa cùng dòng với tên phân lớp (`1s: [↑↓]   2s: [↑↓]   2p: [↑↓][↑ ][↑ ]`).
     + Khi xuất file Word (.doc): Sử dụng hàng bảng tích hợp gọn gàng (`table.ao-opt-row` qua hàm `formatWordOption`) chứa đồng thời nhãn phương án ($A., B., C., D.$), nhãn phân lớp và các ô orbital trên **cùng 1 hàng ngang duy nhất**.
     + **Chuẩn hóa đầu vào của `formatWordOption`**: Tự động giải mã (unescape) toàn bộ thực thể HTML (`&#8593;`, `&#8595;`, `&nbsp;`, `table.ao-boxes`) về ngoặc thô để nhận diện chính xác 100% các câu hỏi đã lưu trong bộ nhớ cache hoặc đã qua bộ lọc `autoSanitizeQuestion`.
     + **Ngăn chặn triệt để lỗi giãn 100% chiều ngang trang Word**: Tuyệt đối không khai báo `table { width: 100%; }` chung cho toàn bộ thẻ `table` trong CSS của file Word, vì Word sẽ ép tất cả bảng (kể cả bảng ô orbital) tràn ngang trang. Thay vào đó, chỉ áp dụng `width: 100%` cho các bảng lớn có định danh riêng (`table.data-table, table.header-table, table.matrix-table, table.spec-table, table.answer-table, table.chem-table`), đồng thời đặt thuộc tính `width` cụ thể trên thẻ `table.ao-opt-row`.
   - Đảm bảo tính chính xác khoa học về phân bố electron vào ô orbital:
     + Phân lớp nào thì tổng số electron trong các phương án phải đúng với phân lớp đó (ví dụ: $2p^3$ thì cả 4 phương án đều phải phân bố đúng 3 electron vào 3 ô orbital).
     + Phân biệt chính xác giữa các quy tắc: Quy tắc Hund (số electron độc thân cực đại và spin cùng chiều), Nguyên lí Pauli (tối đa 2e/orbital và ngược chiều spin ↑↓).




8. **Chuẩn hóa Ký hiệu Nguyên tử & Đồng vị ($^A_Z\text{X}$) chuẩn quốc tế & Word EQ Field**:
   - Ký hiệu nguyên tử phải có số khối $A$ ở trên bên trái, số hiệu nguyên tử $Z$ ở dưới bên trái, và ký hiệu hóa học $X$ ở bên phải ($^A_Z\text{X}$).
   - **Tự động nhận diện & sửa các chuỗi số bị dính liền do lỗi bóc tách tài liệu/OCR**:
     + Chuỗi 3-5 chữ số nối liền ký hiệu nguyên tố hóa học (như `3919K`, `6329Cu`, `6529Cu`, `3581Br`, `1656Fe`, `1123Na`, `1735Cl`...) phải tự động phân tách số khối $A$ và số hiệu nguyên tử $Z$ dựa theo bảng số hiệu nguyên tử chuẩn `ATOMIC_NUMBERS_MAP`.
     + Đảm bảo thông minh: không nhầm lẫn với các đơn vị và hằng số như nhiệt độ (`298K`, `273K`), thế điện cực (`110V`), nồng độ (`001M`), thể tích (`250C`)...
   - **Đảm bảo hiển thị hoàn hảo cả trên Web lẫn Word (.doc)**:
     + Trên Web: sử dụng cấu trúc xếp chồng `###STACK-A-Z###` chuyển thành inline-block với `<sup>A</sup>` trên và `<sub>Z</sub>` dưới trong cùng một cột thẳng đứng bên trái ký hiệu nguyên tố $X$.
     + Trong Word: tự động lồng trường Equation chuẩn Word MSO `<!--[if mso]><span style='mso-element:field-begin'></span> EQ \s\up 8(A)\s\do 3(Z) <span style='mso-element:field-separator'></span><sup>A</sup><sub>Z</sub><span style='mso-element:field-end'></span><![endif]-->` để Word render số khối $A$ nằm ngay trên đầu số hiệu nguyên tử $Z$, tuyệt đối không bị tụt xuống dòng cơ sở thành dãy số ngang như `3919K` hay `6329Cu`.

9. **Chuẩn hóa Ký hiệu & Điện tích của Ion (Cation & Anion)**:
   - Tất cả điện tích ion bắt buộc phải ở chỉ số trên (superscript): độ lớn trước, dấu sau ($2+, 3+, 2-, 3-, +, -$), ví dụ: $\text{Li}^+, \text{Na}^+, \text{K}^+, \text{Mg}^{2+}, \text{Al}^{3+}, \text{Ca}^{2+}, \text{Fe}^{2+}, \text{Fe}^{3+}, \text{Cu}^{2+}, \text{Zn}^{2+}, \text{F}^-, \text{Cl}^-, \text{Br}^-, \text{O}^{2-}, \text{S}^{2-}, \text{N}^{3-}, \text{SO}_4^{2-}, \text{CO}_3^{2-}, \text{PO}_4^{3-}, \text{NO}_3^-, \text{NH}_4^+, \text{OH}^-$.
   - **Tuyệt đối không để dấu điện tích rơi xuống dòng cơ sở** (như `Li+`, `Na+`, `Cl-`).
   - **Tuyệt đối không để nhầm điện tích ion thành chỉ số dưới của công thức phân tử** (như `N3-` bị biến thành $N_3^-$ hay $N_3-$, `O2-` bị biến thành $O_2^-$ hay $O_2-$).
   - Xử lý chính xác các cặp oxi hóa - khử có dấu gạch chéo `/`: $\text{Li}^+/\text{Li}$, $\text{Mg}^{2+}/\text{Mg}$, $\text{Al}^{3+}/\text{Al}$, $\text{Zn}^{2+}/\text{Zn}$, $\text{Fe}^{2+}/\text{Fe}$, $\text{Fe}^{3+}/\text{Fe}^{2+}$, $\text{Cu}^{2+}/\text{Cu}$, $\text{Ag}^+/\text{Ag}$.
   - Tự động chuẩn hóa qua hàm `formatChemicalIons` trong cả pipeline nạp dữ liệu `autoSanitizeQuestion` và bộ hiển thị `cleanLatexForWord`.
