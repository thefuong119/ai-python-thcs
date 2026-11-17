"""
Chương trình tự động tạo giáo án Word (.docx)
Chuyên đề 1: Cấu trúc rẽ nhánh trong lập trình thi đấu (Python)
Yêu cầu thư viện: pip install python-docx
"""

import os
import re
import sys

try:
    import docx
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    from docx.shared import Cm, Pt, RGBColor
except ImportError:
    print("LỖI: Máy tính chưa cài đặt thư viện 'python-docx'.")
    print("Vui lòng mở Terminal / Command Prompt và chạy lệnh sau:")
    print("    pip install python-docx")
    sys.exit(1)


# ==============================================================================
# CÁC HÀM HỖ TRỢ ĐỊNH DẠNG XML & GIAO DIỆN TÀI LIỆU WORD
# ==============================================================================


def set_cell_background(cell, hex_color):
    """Tô màu nền cho ô trong bảng (dùng làm khung code hoặc khung ghi chú)."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
    """Thiết lập khoảng cách lề bên trong ô (đơn vị dxa: 20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in [
        ("top", top),
        ("bottom", bottom),
        ("left", left),
        ("right", right),
    ]:
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def set_cell_borders(cell, color="CBD5E1", sz="6", left_accent=False):
    """Tạo đường viền mảnh cho khung code hoặc viền nhấn bên trái cho khung lý thuyết."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")

    if left_accent:
        # Chỉ tạo viền đậm bên trái (dạng Callout Box)
        for side in ["top", "bottom", "right"]:
            edge = OxmlElement(f"w:{side}")
            edge.set(qn("w:val"), "none")
            tcBorders.append(edge)
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "24")  # Viền trái dày 3pt
        left.set(qn("w:space"), "0")
        left.set(qn("w:color"), color)
        tcBorders.append(left)
    else:
        # Viền bao quanh 4 cạnh cho khung code
        for side in ["top", "left", "bottom", "right"]:
            edge = OxmlElement(f"w:{side}")
            edge.set(qn("w:val"), "single")
            edge.set(qn("w:sz"), sz)
            edge.set(qn("w:space"), "0")
            edge.set(qn("w:color"), color)
            tcBorders.append(edge)

    tcPr.append(tcBorders)


def add_page_number_to_run(run):
    """Chèn trường số trang tự động của Word vào footer."""
    fldChar1 = parse_xml(r'<w:fldChar %s w:fldCharType="begin"/>' % nsdecls("w"))
    instrText = parse_xml(
        r'<w:instrText %s xml:space="preserve"> PAGE </w:instrText>'
        % nsdecls("w")
    )
    fldChar2 = parse_xml(
        r'<w:fldChar %s w:fldCharType="separate"/>' % nsdecls("w")
    )
    fldChar3 = parse_xml(r'<w:fldChar %s w:fldCharType="end"/>' % nsdecls("w"))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)


def setup_document():
    """Khởi tạo tài liệu Word với khổ giấy A4, căn lề chuẩn và Header/Footer."""
    doc = docx.Document()

    # Thiết lập khổ giấy A4 và lề chuẩn giáo án (2cm)
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)

        # Header
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run(
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 1: Cấu trúc rẽ nhánh"
        )
        hrun.font.name = "Times New Roman"
        hrun.font.size = Pt(10)
        hrun.font.italic = True
        hrun.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

        # Footer (Số trang)
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun1 = fp.add_run("Trang ")
        frun1.font.name = "Times New Roman"
        frun1.font.size = Pt(10)
        frun1.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

        frun_pg = fp.add_run()
        frun_pg.font.name = "Times New Roman"
        frun_pg.font.size = Pt(10)
        frun_pg.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        add_page_number_to_run(frun_pg)

    # Định dạng font mặc định Normal
    style_normal = doc.styles["Normal"]
    style_normal.font.name = "Times New Roman"
    style_normal.font.size = Pt(12)
    style_normal.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)

    return doc


def add_rich_paragraph(
    container,
    text,
    space_before=0,
    space_after=4,
    line_spacing=1.15,
    left_indent=0,
    is_bullet=False,
):
    """
    Thêm đoạn văn hỗ trợ định dạng trộn lẫn:
    - **chữ in đậm**
    - `mã lệnh inline`
    """
    p = container.add_paragraph(
        style="List Bullet" if is_bullet else "Normal"
    )
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if left_indent > 0 and not is_bullet:
        p.paragraph_format.left_indent = Cm(left_indent)

    # Tách các cụm **bold** và `inline code`
    tokens = re.split(r"(\*\*.*?\*\*|`.*?`)", text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**"):
            run = p.add_run(token[2:-2])
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            run.bold = True
        elif token.startswith("`") and token.endswith("`"):
            run = p.add_run(token[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(10.5)
            run.font.color.rgb = RGBColor(0xB9, 0x1C, 0x1C)  # Đỏ đô nổi bật
        else:
            run = p.add_run(token)
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
    return p


def add_part_heading(doc, title):
    """Tiêu đề Phần lớn (PHẦN I, PHẦN II, PHẦN III)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True

    # Tạo đường kẻ dưới tiêu đề phần
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "4")
    bottom.set(qn("w:color"), "1E3A8A")
    pBdr.append(bottom)
    pPr.append(pBdr)

    run = p.add_run(title)
    run.font.name = "Times New Roman"
    run.font.size = Pt(14)
    run.bold = True
    run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)  # Xanh Navy đậm


def add_sub_heading(doc, title):
    """Tiêu đề mục con hoặc tên từng bài tập."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(title)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12.5)
    run.bold = True
    run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)


def add_code_block(doc, code_text, label="Code Python mẫu chuẩn:"):
    """Tạo khung hiển thị mã nguồn Python có nền xám nhạt, viền mảnh, font Consolas."""
    if label:
        lp = doc.add_paragraph()
        lp.paragraph_format.space_before = Pt(4)
        lp.paragraph_format.space_after = Pt(2)
        lp.paragraph_format.keep_with_next = True
        lrun = lp.add_run(label)
        lrun.font.name = "Times New Roman"
        lrun.font.size = Pt(11.5)
        lrun.bold = True
        lrun.italic = True
        lrun.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)

    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(17.0)

    cell = table.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_background(cell, "F8FAFC")
    set_cell_borders(cell, color="CBD5E1", sz="6", left_accent=False)
    set_cell_margins(cell, top=100, bottom=100, left=160, right=160)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.keep_together = True  # Không ngắt đôi khung code qua trang

    run = p.add_run(code_text.strip("\n"))
    run.font.name = "Consolas"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    # Tạo khoảng cách nhỏ sau bảng
    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)


def add_callout_box(doc, title, lines):
    """Tạo khung nhấn mạnh cho phần Phân tích tư duy HSG hoặc Lưu ý."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(17.0)

    cell = table.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_background(cell, "EFF6FF")
    set_cell_borders(cell, color="2563EB", left_accent=True)
    set_cell_margins(cell, top=90, bottom=90, left=160, right=140)

    p0 = cell.paragraphs[0]
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(3)
    r0 = p0.add_run(title)
    r0.font.name = "Times New Roman"
    r0.font.size = Pt(11.5)
    r0.bold = True
    r0.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)

    for idx, line in enumerate(lines):
        add_rich_paragraph(
            cell,
            line,
            space_before=0,
            space_after=2 if idx < len(lines) - 1 else 0,
            line_spacing=1.12,
            is_bullet=True,
        )

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(3)


# ==============================================================================
# NỘI DUNG CHUYÊN ĐỀ 1: CẤU TRÚC RẼ NHÁNH
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 1: CẤU TRÚC RẼ NHÁNH TRONG LẬP TRÌNH THI ĐẤU (PYTHON)"
    )
    trun.font.name = "Times New Roman"
    trun.font.size = Pt(16)
    trun.bold = True
    trun.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    sub_p = doc.add_paragraph()
    sub_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_p.paragraph_format.space_before = Pt(0)
    sub_p.paragraph_format.space_after = Pt(12)
    srun = sub_p.add_run(
        "Tài liệu bồi dưỡng Học sinh giỏi Tin học cấp Tỉnh & Quốc gia"
    )
    srun.font.name = "Times New Roman"
    srun.font.size = Pt(12)
    srun.italic = True
    srun.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # ==========================================================================
    # PHẦN I: LÝ THUYẾT NGẮN GỌN TRỌNG TÂM
    # ==========================================================================
    add_part_heading(doc, "PHẦN I. LÝ THUYẾT NGẮN GỌN TRỌNG TÂM")
    add_rich_paragraph(
        doc,
        "Trong các kỳ thi Học sinh giỏi (HSG) Tin học, bài đầu tiên (thường chiếm 25% - 35% tổng số điểm) luôn kiểm tra khả năng đọc hiểu đề, phân tích logic và xử lý các trường hợp bằng cấu trúc rẽ nhánh kết hợp kiến thức toán học.",
    )

    add_sub_heading(doc, "1. Cú pháp chuẩn và Biểu thức điều kiện")
    add_rich_paragraph(
        doc,
        "**Cấu trúc đầy đủ:** Sử dụng `if`, `elif`, `else` với thụt lề chuẩn 4 dấu cách.",
        is_bullet=True,
    )
    add_code_block(
        doc,
        """if dieu_kien_1:
    # Khối lệnh 1
elif dieu_kien_2:
    # Khối lệnh 2
else:
    # Khối lệnh mặc định""",
        label="",
    )

    add_rich_paragraph(
        doc,
        "**Biểu thức điều kiện rút gọn (Ternary Operator):** Giúp mã nguồn ngắn gọn, rõ ràng khi gán giá trị.",
        is_bullet=True,
    )
    add_code_block(
        doc,
        "res = gia_tri_dung if dieu_kien else gia_tri_sai",
        label="",
    )

    add_rich_paragraph(
        doc,
        "**So sánh chuỗi (Chained Comparisons):** Đặc trưng mạnh mẽ của Python, giúp biểu diễn khoảng giá trị tự nhiên như toán học.",
        is_bullet=True,
    )
    add_code_block(
        doc,
        """if 0 <= x <= 100 and a < b < c:
    ...""",
        label="",
    )

    add_sub_heading(
        doc, "2. Toán tử logic và Đánh giá đoản mạch (Short-circuit Evaluation)"
    )
    add_rich_paragraph(
        doc,
        "Thứ tự ưu tiên của toán tử logic: `not` → `and` → `or`.",
    )
    add_rich_paragraph(
        doc,
        "Với `A and B`: Nếu `A` sai (`False`), Python dừng ngay và trả về `False` mà không kiểm tra `B`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "Với `A or B`: Nếu `A` đúng (`True`), Python dừng ngay và trả về `True` mà không kiểm tra `B`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Ứng dụng trong thi HSG:** Luôn đặt điều kiện kiểm tra mẫu số khác 0 hoặc kiểm tra biên trước phép chia:",
        is_bullet=True,
    )
    add_code_block(
        doc,
        "if b != 0 and a % b == 0:  # An toàn tuyệt đối, không bao giờ gặp lỗi ZeroDivisionError",
        label="",
    )

    add_sub_heading(
        doc, '3. Bốn "cạm bẫy" kinh điển trong phòng thi HSG cần tránh'
    )
    add_rich_paragraph(
        doc,
        "**Bẫy sai số số thực (`float` precision):** Số nguyên trong Python có độ chính xác vô hạn, nhưng kiểu số thực `float` chỉ chuẩn xác khoảng 15 - 17 chữ số thập phân (chuẩn IEEE 754). Khi N ≤ 10^18, phép chia `/` hoặc hàm `math.sqrt()`, `math.ceil()` sẽ chuyển số sang `float` và làm mất độ chính xác ở các chữ số cuối. **Quy tắc vàng:** Quy đồng mẫu số để khử phép chia, bình phương hai vế để khử căn bậc hai, và dùng phép chia lấy nguyên `//` thay cho `/`. Nếu bắt buộc so sánh hai số thực x, y, hãy dùng sai số ε = 10^-9: `abs(x - y) < 1e-9`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy thứ tự điều kiện `elif`:** Luôn xét điều kiện **hẹp nhất (đặc biệt nhất)** trước, điều kiện **rộng nhất (tổng quát nhất)** sau. Ví dụ: Khi phân loại tam giác, phải kiểm tra tam giác **đều** trước khi kiểm tra tam giác **cân**.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy trường hợp biên (Edge cases):** Luôn kiểm tra các giá trị đặc biệt như 0, số âm, các mốc ranh giới xảy ra dấu bằng (=), hoặc trường hợp không tồn tại nghiệm.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Chuẩn nhập/xuất dữ liệu:** Đọc nhiều số nguyên trên cùng một dòng bằng cú pháp: `a, b, c = map(int, input().split())`.",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(doc, "Ví dụ 1: Phân loại tam giác theo cạnh và góc")
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho ba số nguyên dương a, b, c (1 ≤ a, b, c ≤ 10^9). Kiểm tra xem a, b, c có tạo thành một tam giác hay không. Nếu không, in ra `KHONG`. Nếu có, hãy cho biết loại tam giác theo cạnh (`DEU`, `CAN`, `THUONG`) và theo góc (`VUONG`, `NHON`, `TU`), cách nhau bởi một dấu cách.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Sắp xếp 3 cạnh tăng dần a ≤ b ≤ c. Khi đó điều kiện tồn tại tam giác chỉ còn duy nhất: a + b > c.",
            "Vì c là cạnh lớn nhất nên góc đối diện cạnh c là góc lớn nhất. Áp dụng định lý Cosin bằng số nguyên (không dùng căn bậc hai):",
            "Nếu a² + b² = c² → Tam giác vuông (`VUONG`).",
            "Nếu a² + b² > c² → Tam giác nhọn (`NHON`).",
            "Nếu a² + b² < c² → Tam giác tù (`TU`).",
        ],
    )
    add_code_block(
        doc,
        """a, b, c = sorted(map(int, input().split()))

if a + b <= c:
    print("KHONG")
else:
    # Phân loại theo cạnh (ưu tiên DEU trước CAN)
    if a == c:  # Vì a <= b <= c nên a == c tương đương a == b == c
        canh = "DEU"
    elif a == b or b == c:
        canh = "CAN"
    else:
        canh = "THUONG"

    # Phân loại theo góc (sử dụng số nguyên 100%)
    tong_binh_phuong = a * a + b * b
    c_binh_phuong = c * c
    if tong_binh_phuong == c_binh_phuong:
        goc = "VUONG"
    elif tong_binh_phuong > c_binh_phuong:
        goc = "NHON"
    else:
        goc = "TU"

    print(f"{canh} {goc}")""",
    )

    # Ví dụ 2
    add_sub_heading(doc, "Ví dụ 2: Độ dài phần giao của hai đoạn thẳng")
    add_rich_paragraph(
        doc,
        "**Đề bài:** Trên trục số Ox, cho hai đoạn thẳng [a, b] và [c, d] với -10^18 ≤ a ≤ b ≤ 10^18 và -10^18 ≤ c ≤ d ≤ 10^18. Tính độ dài phần giao nhau của hai đoạn thẳng này (nếu không giao nhau hoặc chỉ chạm nhau tại một điểm thì độ dài bằng 0).",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Thay vì chia hàng chục trường hợp vị trí tương đối của hai đoạn thẳng, nhận xét phần giao của [a, b] và [c, d] chính là đoạn [L, R] với L = max(a, c) và R = min(b, d).",
            "Nếu L < R thì độ dài phần giao là R - L; ngược lại (L ≥ R) thì độ dài bằng 0.",
        ],
    )
    add_code_block(
        doc,
        """a, b, c, d = map(int, input().split())

left = max(a, c)
right = min(b, d)

if left < right:
    print(right - left)
else:
    print(0)""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1: Tính tiền điện sinh hoạt bậc thang",
            "desc": "Giá bán lẻ điện sinh hoạt được tính theo 4 bậc: Bậc 1 (0 đến 50 kWh): 1.800 đồng/kWh; Bậc 2 (51 đến 100 kWh): 2.000 đồng/kWh; Bậc 3 (101 đến 200 kWh): 2.500 đồng/kWh; Bậc 4 (từ kWh thứ 201 trở đi): 3.000 đồng/kWh. Cho số điện tiêu thụ N (0 ≤ N ≤ 10^6). Tính tổng số tiền điện phải trả.",
            "example": "Ví dụ: Input: `120` → Output: `240000` (50×1800 + 50×2000 + 20×2500 = 240000).",
            "code": """n = int(input())

if n <= 50:
    tien = n * 1800
elif n <= 100:
    tien = 50 * 1800 + (n - 50) * 2000
elif n <= 200:
    tien = 50 * 1800 + 50 * 2000 + (n - 100) * 2500
else:
    tien = 50 * 1800 + 50 * 2000 + 100 * 2500 + (n - 200) * 3000

print(tien)""",
        },
        {
            "title": "Bài 2: Ngày kế tiếp",
            "desc": "Cho ba số nguyên dương d, m, y biểu diễn một ngày hợp lệ trong lịch Dương (1 ≤ y ≤ 10^9). Hãy tìm ngày, tháng, năm của ngày kế tiếp ngay sau ngày đã cho. (Năm nhuận là năm chia hết cho 400, hoặc chia hết cho 4 nhưng không chia hết cho 100).",
            "example": "Ví dụ: Input: `28 2 2024` → Output: `29 2 2024` | Input: `31 12 2025` → Output: `1 1 2026`.",
            "code": """d, m, y = map(int, input().split())

# Xác định số ngày tối đa của tháng m trong năm y
if m == 2:
    is_leap = (y % 400 == 0) or (y % 4 == 0 and y % 100 != 0)
    max_days = 29 if is_leap else 28
elif m in (4, 6, 9, 11):
    max_days = 30
else:
    max_days = 31

# Tăng ngày và cập nhật tháng, năm
if d < max_days:
    d += 1
else:
    d = 1
    if m < 12:
        m += 1
    else:
        m = 1
        y += 1

print(d, m, y)""",
        },
        {
            "title": "Bài 3: Diện tích phần giao của hai hình chữ nhật",
            "desc": "Trên mặt phẳng tọa độ Oxy, cho hai hình chữ nhật có các cạnh song song với các trục tọa độ. Hình chữ nhật 1 xác định bởi hai đỉnh đối diện (x1, y1) và (x2, y2). Hình chữ nhật 2 xác định bởi hai đỉnh đối diện (x3, y3) và (x4, y4). Tọa độ có trị tuyệt đối ≤ 10^9. Tính diện tích phần giao nhau của hai hình chữ nhật. (Lưu ý: Hai đỉnh đối diện chưa chắc là góc dưới-trái và trên-phải).",
            "example": "Ví dụ: Input: `0 0 4 4` và `2 2 6 6` → Output: `4`.",
            "code": """x1, y1, x2, y2 = map(int, input().split())
x3, y3, x4, y4 = map(int, input().split())

# Chuẩn hóa tọa độ về (trái, dưới) và (phải, trên)
left1, right1 = min(x1, x2), max(x1, x2)
bot1, top1 = min(y1, y2), max(y1, y2)

left2, right2 = min(x3, x4), max(x3, x4)
bot2, top2 = min(y3, y4), max(y3, y4)

# Tìm hình chữ nhật giao nhau theo từng trục Ox và Oy
x_left = max(left1, left2)
x_right = min(right1, right2)
y_bot = max(bot1, bot2)
y_top = min(top1, top2)

if x_left < x_right and y_bot < y_top:
    print((x_right - x_left) * (y_top - y_bot))
else:
    print(0)""",
        },
        {
            "title": "Bài 4: Nước đi của quân Hậu",
            "desc": "Trên bàn cờ vua 8×8 (hàng 1..8, cột 1..8), quân Hậu đứng ở ô (r1, c1) muốn di chuyển đến ô (r2, c2). Mỗi bước quân Hậu đi thẳng theo hàng ngang, cột dọc hoặc đường chéo với số ô tùy ý. Tìm số nước đi ít nhất để đến đích. (Phân tích: Trùng ô mất 0 nước; cùng hàng, cột hoặc đường chéo |r1 - r2| = |c1 - c2| mất 1 nước; mọi ô còn lại trên bàn cờ đều đi tới trong 2 nước).",
            "example": "Ví dụ: Input: `1 1 3 5` → Output: `2`.",
            "code": """r1, c1, r2, c2 = map(int, input().split())

if r1 == r2 and c1 == c2:
    print(0)
elif r1 == r2 or c1 == c2 or abs(r1 - r2) == abs(c1 - c2):
    print(1)
else:
    print(2)""",
        },
        {
            "title": "Bài 5: Lát gạch quảng trường",
            "desc": "Quảng trường hình chữ nhật kích thước N × M mét cần lát kín bằng các viên gạch vuông A × A mét (1 ≤ N, M, A ≤ 10^18). Được phép lát tràn ra ngoài biên nhưng không được cắt nhỏ gạch. Tìm số viên gạch ít nhất cần dùng. (Lưu ý bẫy số lớn: Không dùng `math.ceil(n / a)` vì sai số `float` khi N, A ≈ 10^18; hãy dùng công thức số nguyên `(X + A - 1) // A`).",
            "example": "Ví dụ: Input: `6 6 4` → Output: `4`.",
            "code": """n, m, a = map(int, input().split())

so_vien_ngang = (n + a - 1) // a
so_vien_doc = (m + a - 1) // a

print(so_vien_ngang * so_vien_doc)""",
        },
        {
            "title": "Bài 6: Tính lương nhân viên ca ngày và ca đêm",
            "desc": "Giờ làm việc ban ngày tính từ 6h đến 22h với mức lương 50.000 đồng/giờ. Giờ làm việc ban đêm (từ 22h hôm trước đến 6h sáng hôm sau) có mức lương 80.000 đồng/giờ. Nhân viên bắt đầu làm lúc S giờ và kết thúc lúc E giờ (0 ≤ S, E ≤ 24, S ≠ E, ca làm ≤ 24 tiếng). Tính tiền lương của ca làm việc. (Gợi ý: Nếu E < S, quy đổi E' = E + 24 rồi tìm phần giao với hai khung giờ ngày [6, 22] và [30, 46]).",
            "example": "Ví dụ: Input: `20 2` (2h ngày + 4h đêm) → Output: `420000`.",
            "code": """s, e = map(int, input().split())

if e < s:
    e += 24

tong_gio = e - s

# Tính giao với khung giờ ngày 1: [6, 22]
ngay_1 = max(0, min(e, 22) - max(s, 6))
# Tính giao với khung giờ ngày 2 (nếu vắt sang hôm sau): [30, 46]
ngay_2 = max(0, min(e, 46) - max(s, 30))

gio_ngay = ngay_1 + ngay_2
gio_dem = tong_gio - gio_ngay

luong = gio_ngay * 50000 + gio_dem * 80000
print(luong)""",
        },
        {
            "title": "Bài 7: Kiểm tra điểm nằm trên đoạn thẳng",
            "desc": "Trên mặt phẳng Oxy, cho ba điểm nguyên A(xA, yA), B(xB, yB) và M(xM, yM) với A ≠ B, tọa độ có trị tuyệt đối ≤ 10^9. Kiểm tra xem M có nằm trên đoạn thẳng AB hay không (kể cả hai đầu mút). In ra `YES` hoặc `NO`. (Điều kiện: Tích có hướng `(xB - xA)(yM - yA) - (yB - yA)(xM - xA) == 0` và tọa độ M nằm trong khung hình chữ nhật bao quanh AB).",
            "example": "Ví dụ: Input: `0 0`, `4 4`, `2 2` → Output: `YES`.",
            "code": """xa, ya = map(int, input().split())
xb, yb = map(int, input().split())
xm, ym = map(int, input().split())

tich_co_huong = (xb - xa) * (ym - ya) - (yb - ya) * (xm - xa)

trong_khung_x = min(xa, xb) <= xm <= max(xa, xb)
trong_khung_y = min(ya, yb) <= ym <= max(ya, yb)

if tich_co_huong == 0 and trong_khung_x and trong_khung_y:
    print("YES")
else:
    print("NO")""",
        },
        {
            "title": "Bài 8: Ghép số lớn nhất",
            "desc": "Cho ba số nguyên dương a, b, c (1 ≤ a, b, c ≤ 10^18). Hãy ghép cả ba số này liền nhau theo một thứ tự nào đó để tạo thành số nguyên lớn nhất có thể. (Gợi ý: Để quyết định xâu số X đứng trước hay Y đứng trước, ta so sánh xâu ghép `X + Y` với `Y + X`).",
            "example": "Ví dụ: Input: `12 9 85` → Output: `98512`.",
            "code": """a, b, c = input().split()

# Sắp xếp 3 phần tử bằng phép so sánh ghép xâu
if a + b < b + a:
    a, b = b, a
if a + c < c + a:
    a, c = c, a
if b + c < c + b:
    b, c = c, b

print(a + b + c)""",
        },
        {
            "title": "Bài 9: Khuyến mãi mua vở",
            "desc": "Giá bán lẻ mỗi quyển vở là P đồng, mua theo lốc K quyển giá cả lốc là Q đồng (1 ≤ P, Q ≤ 10^9; 2 ≤ K ≤ 10^9). Bạn cần mua ít nhất N quyển vở (1 ≤ N ≤ 10^9). Tìm số tiền ít nhất phải trả. (Lưu ý: Khi mua theo lốc rẻ hơn mua lẻ, phần vở lẻ còn dư `du = N % K` có thể mua lẻ tốn `du * P` hoặc mua hẳn thêm 1 lốc tốn `Q`, ta chọn phương án rẻ hơn).",
            "example": "Ví dụ: Input: `10 3 5000 12000` → Output: `41000`.",
            "code": """n, k, p, q = map(int, input().split())

if q >= k * p:
    print(n * p)
else:
    so_loc = n // k
    du = n % k
    tien_phan_du = min(du * p, q)
    print(so_loc * q + tien_phan_du)""",
        },
        {
            "title": "Bài 10: Chu kỳ đèn giao thông",
            "desc": "Trụ đèn giao thông hoạt động tuần hoàn: Xanh sáng G giây, Vàng sáng Y giây, Đỏ sáng R giây (1 ≤ G, Y, R ≤ 10^9). Tại t = 0 đèn bắt đầu chuyển sang Xanh. Hỏi tại giây thứ T (0 ≤ T ≤ 10^18), đèn đang sáng màu gì (`XANH`, `VANG` hay `DO`) và còn bao nhiêu giây nữa thì chuyển sang màu tiếp theo?",
            "example": "Ví dụ: Input: `30 5 25 32` → Output: `VANG 3`.",
            "code": """g, y, r, t = map(int, input().split())

chu_ky = g + y + r
rem = t % chu_ky

if rem < g:
    print("XANH", g - rem)
elif rem < g + y:
    print("VANG", (g + y) - rem)
else:
    print("DO", chu_ky - rem)""",
        },
    ]

    for bt in bai_tap_list:
        add_sub_heading(doc, bt["title"])
        add_rich_paragraph(doc, f"**Đề bài:** {bt['desc']}")
        add_rich_paragraph(doc, bt["example"], is_bullet=True)
        add_code_block(doc, bt["code"])

    # Lưu tệp Word ra ổ đĩa
    doc.save(output_filename)
    print(
        f"THÀNH CÔNG! Đã tạo tệp giáo án: {os.path.abspath(output_filename)}"
    )


if __name__ == "__main__":
    OUTPUT_FILE = "CAU_TRUC_RE_NHANH.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")