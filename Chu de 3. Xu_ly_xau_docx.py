import os
import re
import subprocess
import sys

try:
    import docx
except ImportError:
    print("Đang tự động cài đặt thư viện 'python-docx', vui lòng chờ trong giây lát...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "python-docx"]
        )
        import docx
    except Exception as install_err:
        print(f"LỖI: Không thể tự động cài đặt 'python-docx': {install_err}")
        print(
            f'Vui lòng mở Terminal trong IDE và chạy: "{sys.executable}" -m pip install python-docx'
        )
        sys.exit(1)

from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor


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
        for side in ["top", "bottom", "right"]:
            edge = OxmlElement(f"w:{side}")
            edge.set(qn("w:val"), "none")
            tcBorders.append(edge)
        left = OxmlElement("w:left")
        left.set(qn("w:val"), "single")
        left.set(qn("w:sz"), "24")
        left.set(qn("w:space"), "0")
        left.set(qn("w:color"), color)
        tcBorders.append(left)
    else:
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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 3: Xử lý xâu, chuỗi"
        )
        hrun.font.name = "Times New Roman"
        hrun.font.size = Pt(10)
        hrun.font.italic = True
        hrun.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

        # Footer
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
    """Thêm đoạn văn hỗ trợ định dạng trộn lẫn **bold** và `inline code`."""
    p = container.add_paragraph(
        style="List Bullet" if is_bullet else "Normal"
    )
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if left_indent > 0 and not is_bullet:
        p.paragraph_format.left_indent = Cm(left_indent)

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
            run.font.color.rgb = RGBColor(0xB9, 0x1C, 0x1C)
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
    run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)


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
    p.paragraph_format.keep_together = True

    run = p.add_run(code_text.strip("\n"))
    run.font.name = "Consolas"
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

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
# NỘI DUNG CHUYÊN ĐỀ 3: XỬ LÝ XÂU, CHUỖI
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 3: XỬ LÝ XÂU, CHUỖI TRONG LẬP TRÌNH THI ĐẤU (PYTHON)"
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
        "Trong các đề thi Học sinh giỏi (HSG) Tin học cấp Tỉnh và Quốc gia, bài toán xử lý xâu ký tự (`str`) luôn xuất hiện với tần suất rất cao (thường nằm ở câu 2 hoặc câu 3 của đề thi). Để giải quyết trọn vẹn các bộ dữ liệu lớn (|S| ≤ 10^6) trong giới hạn 1 giây, học sinh cần nắm vững cơ chế quản lý bộ nhớ của kiểu `str` trong Python và các kỹ thuật thuật toán đặc trưng trên xâu.",
    )

    add_sub_heading(
        doc, "1. Tính bất biến (Immutability) và Kỹ thuật xây dựng xâu O(N)"
    )
    add_rich_paragraph(
        doc,
        "**Đặc điểm cốt lõi:** Xâu ký tự trong Python là đối tượng **bất biến** (immutable), nghĩa là không thể thay đổi trực tiếp một ký tự tại vị trí `i` (lệnh `s[i] = 'a'` sẽ gây lỗi `TypeError`).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Cạm bẫy nối xâu bằng `+=`:** Khi viết `res += ch` trong một vòng lặp chạy N lần, Python phải cấp phát vùng nhớ mới và sao chép toàn bộ xâu cũ ở mỗi bước, khiến độ phức tạp bị đẩy lên O(N²) và gây lỗi quá thời gian (TLE) khi N ≥ 10^5.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Quy tắc vàng:** Luôn lưu các phần tử vào một danh sách (`list`) bằng `.append()`, sau đó ghép một lần duy nhất ở cuối bằng `\"\".join(res)` để đạt độ phức tạp tuyến tính O(N):",
        is_bullet=True,
    )
    add_code_block(
        doc,
        """# CÁCH VIẾT CHUẨN HSG - ĐỘ PHỨC TẠP O(N)
res = []
for ch in s:
    if ch != ' ':
        res.append(ch)
ket_qua = "".join(res)""",
        label="",
    )

    add_sub_heading(doc, "2. Kỹ thuật cắt lát (Slicing) và Đảo xâu")
    add_rich_paragraph(
        doc,
        "**Cú pháp:** `s[start : stop : step]` trích xuất xâu con từ chỉ số `start` đến sát `stop` với bước nhảy `step` (độ phức tạp O(K) với K là độ dài đoạn được cắt).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Các thao tác thường dùng:** Đảo ngược xâu `s[::-1]`; lấy K ký tự đầu `s[:k]`; lấy K ký tự cuối `s[-k:]`; kiểm tra tiền tố/hậu tố nhanh `s.startswith(prefix)`, `s.endswith(suffix)`.",
        is_bullet=True,
    )

    add_sub_heading(doc, "3. Bảng mã ASCII và Kỹ thuật đếm tần suất ký tự")
    add_rich_paragraph(
        doc,
        "**Chuyển đổi ký tự và mã số:** `ord(ch)` trả về mã Unicode/ASCII của ký tự `ch` (`ord('A') = 65`, `ord('a') = 97`, `ord('0') = 48`); `chr(code)` trả về ký tự tương ứng với mã số `code`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Ánh xạ ký tự sang chỉ số mảng:** Với xâu chỉ gồm chữ cái in thường `'a'..'z'`, quy đổi ký tự `ch` về chỉ số 0..25 bằng công thức `ord(ch) - 97` để dùng mảng đếm `cnt = [0] * 26`, tốc độ truy xuất nhanh hơn đáng kể so với dùng từ điển (`dict`).",
        is_bullet=True,
    )

    add_sub_heading(
        doc, "4. Các phương thức xử lý xâu cốt lõi và Bẫy nhập liệu"
    )
    add_rich_paragraph(
        doc,
        "**Phân loại và chuyển đổi:** `ch.isdigit()`, `ch.isalpha()`, `ch.isalnum()`, `s.lower()`, `s.upper()`, `s.capitalize()` (viết hoa chữ cái đầu tiên, các chữ sau viết thường).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy nhập liệu với `.strip()`:** Hàm `input().strip()` xóa **mọi** khoảng trắng ở cả đầu và cuối xâu. Nếu đề bài cho xâu chứa dấu cách có nghĩa ở đầu hoặc cuối (ví dụ bài toán mã hóa), phải dùng `input().rstrip(\"\\r\\n\")` để giữ nguyên dấu cách!",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(
        doc, "Ví dụ 1: Kiểm tra xâu đối xứng (Palindrome) bỏ qua dấu cách"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho một xâu ký tự S (1 ≤ |S| ≤ 10^6) chỉ gồm các chữ cái tiếng Anh và dấu cách. Hãy kiểm tra xem S có phải là xâu đối xứng hay không nếu bỏ qua toàn bộ dấu cách và không phân biệt chữ hoa, chữ thường. In ra `YES` hoặc `NO`.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Cách thông thường lọc bỏ dấu cách rồi so sánh `t == t[::-1]` tốn thêm O(N) bộ nhớ phụ.",
            "Cách tối ưu (Kỹ thuật Hai con trỏ): Đặt `l = 0` và `r = len(s) - 1`. Nếu `s[l]` là dấu cách thì tăng `l`; nếu `s[r]` là dấu cách thì giảm `r`; ngược lại so sánh `s[l].lower()` với `s[r].lower()`. Độ phức tạp thời gian O(N) và bộ nhớ phụ O(1).",
        ],
    )
    add_code_block(
        doc,
        """s = input()
l, r = 0, len(s) - 1

while l < r:
    if s[l] == " ":
        l += 1
    elif s[r] == " ":
        r -= 1
    elif s[l].lower() != s[r].lower():
        print("NO")
        break
    else:
        l += 1
        r -= 1
else:
    print("YES")""",
    )

    # Ví dụ 2
    add_sub_heading(
        doc, "Ví dụ 2: Ký tự xuất hiện nhiều nhất trong xâu"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho xâu ký tự S (1 ≤ |S| ≤ 10^6) chỉ gồm các ký tự in thường và chữ số. Hãy tìm ký tự xuất hiện nhiều lần nhất trong S và số lần xuất hiện của nó. Nếu có nhiều ký tự có cùng số lần xuất hiện lớn nhất, hãy chọn ký tự có thứ tự từ điển nhỏ nhất.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Duyệt một lần qua xâu S để đếm tần suất xuất hiện của từng ký tự bằng từ điển `dem` trong O(N).",
            "Duyệt qua các cặp `(ch, freq)` trong `dem`: cập nhật kết quả nếu `freq > max_freq` hoặc (`freq == max_freq` và `ch < best_char`). Tổng độ phức tạp là O(N).",
        ],
    )
    add_code_block(
        doc,
        """s = input().strip()
dem = {}

for ch in s:
    dem[ch] = dem.get(ch, 0) + 1

best_char = ""
max_freq = 0

for ch, freq in dem.items():
    if freq > max_freq or (freq == max_freq and ch < best_char):
        max_freq = freq
        best_char = ch

print(best_char, max_freq)""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1: Chuẩn hóa xâu họ tên",
            "desc": "Cho một xâu ký tự S (|S| ≤ 10^5) biểu diễn họ và tên một người nhưng bị gõ sai quy cách (thừa nhiều dấu cách ở đầu, cuối, giữa các từ và viết hoa/thường lộn xộn). Hãy chuẩn hóa xâu S sao cho giữa các từ chỉ có đúng một dấu cách, không có dấu cách ở đầu và cuối, chữ cái đầu mỗi từ viết hoa, các chữ cái còn lại viết thường.",
            "analysis": "Hàm `input().split()` tự động loại bỏ khoảng trắng thừa ở hai đầu và tách các từ dựa trên cụm khoảng trắng liên tiếp. Sau đó áp dụng `.capitalize()` cho từng từ rồi ghép lại bằng `\" \".join(...)`.",
            "example": 'Ví dụ: Input: `"   nGuYeN    vAn   aN  "` → Output: `"Nguyen Van An"`.',
            "code": """words = input().split()
res = [w.capitalize() for w in words]
print(" ".join(res))""",
        },
        {
            "title": "Bài 2: Kiểm tra hai xâu đảo chữ (Anagram)",
            "desc": "Hai xâu A và B được gọi là đảo chữ của nhau nếu ta có thể sắp xếp lại các ký tự của xâu A để thu được xâu B. Cho hai xâu A, B (1 ≤ |A|, |B| ≤ 10^6) chỉ gồm các chữ cái tiếng Anh in thường `'a'..'z'`. Kiểm tra xem A và B có phải là hai xâu đảo chữ của nhau hay không. In ra `YES` hoặc `NO`.",
            "analysis": "Dùng mảng đếm tần suất 26 phần tử: cộng 1 với mỗi ký tự trong A và trừ 1 với mỗi ký tự trong B. Hai xâu là đảo chữ khi toàn bộ mảng đếm bằng 0, độ phức tạp O(N).",
            "example": "Ví dụ: Input: `listen` và `silent` → Output: `YES`.",
            "code": """a = input().strip()
b = input().strip()

if len(a) != len(b):
    print("NO")
else:
    cnt = [0] * 26
    for ch in a:
        cnt[ord(ch) - 97] += 1
    for ch in b:
        cnt[ord(ch) - 97] -= 1

    print("YES" if all(x == 0 for x in cnt) else "NO")""",
        },
        {
            "title": "Bài 3: Tính tổng các số nguyên xuất hiện trong xâu",
            "desc": "Cho một xâu ký tự S (1 ≤ |S| ≤ 10^6) gồm các chữ cái và chữ số. Hãy trích xuất tất cả các số nguyên không âm xuất hiện liền nhau trong xâu và tính tổng của chúng.",
            "analysis": "Duyệt từng ký tự: nếu là chữ số thì cập nhật `so_hien_tai = so_hien_tai * 10 + int(ch)`; nếu gặp chữ cái thì cộng vào `tong` và đặt lại `so_hien_tai = 0`. Nhớ cộng nốt `so_hien_tai` sau vòng lặp.",
            "example": "Ví dụ: Input: `abc123def45g6` → Output: `174` (123 + 45 + 6 = 174).",
            "code": """s = input().strip()
tong = 0
so_hien_tai = 0

for ch in s:
    if ch.isdigit():
        so_hien_tai = so_hien_tai * 10 + int(ch)
    else:
        tong += so_hien_tai
        so_hien_tai = 0

tong += so_hien_tai
print(tong)""",
        },
        {
            "title": "Bài 4: Giải mã xâu nén (Run-Length Decoding)",
            "desc": "Một xâu ký tự ban đầu chỉ gồm các chữ cái in thường được nén bằng cách viết mỗi ký tự kèm theo số lần lặp liên tiếp của nó (ví dụ `aaabbbbbbbbbbbbc` được nén thành `a3b12c1`). Cho xâu đã nén S (1 ≤ |S| ≤ 10^5, độ dài xâu gốc ≤ 10^6). Hãy khôi phục lại xâu ban đầu.",
            "analysis": "Dùng chỉ số `i` lấy ký tự `ch = s[i]`, sau đó dùng vòng lặp `while` đọc liên tiếp các chữ số đứng sau để ghép thành số nguyên `num` rồi thêm `ch * num` vào danh sách kết quả.",
            "example": "Ví dụ: Input: `a3b12c2` → Output: `aaabbbbbbbbbbbbcc`.",
            "code": """s = input().strip()
res = []
i = 0
n = len(s)

while i < n:
    ch = s[i]
    i += 1
    num = 0
    while i < n and s[i].isdigit():
        num = num * 10 + int(s[i])
        i += 1
    res.append(ch * num)

print("".join(res))""",
        },
        {
            "title": "Bài 5: Mật mã Caesar",
            "desc": "Mật mã Caesar dịch chuyển mỗi chữ cái trong bảng chữ cái tiếng Anh tiến lên K vị trí theo vòng tròn (sau `'z'` quay về `'a'`, sau `'Z'` quay về `'A'`). Các ký tự không phải chữ cái được giữ nguyên. Cho số nguyên K (0 ≤ K ≤ 10^18) và xâu S (1 ≤ |S| ≤ 10^6). Hãy in ra xâu sau khi mã hóa.",
            "analysis": "Rút gọn khóa dịch chuyển bằng `k %= 26`, sau đó áp dụng công thức dịch vòng theo mã ASCII cho chữ cái in thường (mốc 97) và chữ cái in hoa (mốc 65).",
            "example": "Ví dụ: Input: `3` và `Hello, World! 2026` → Output: `Khoor, Zruog! 2026`.",
            "code": """k = int(input())
s = input().rstrip("\\r\\n")
k %= 26
res = []

for ch in s:
    if "a" <= ch <= "z":
        res.append(chr((ord(ch) - 97 + k) % 26 + 97))
    elif "A" <= ch <= "Z":
        res.append(chr((ord(ch) - 65 + k) % 26 + 65))
    else:
        res.append(ch)

print("".join(res))""",
        },
        {
            "title": "Bài 6: Kiểm tra xâu con không liên tiếp (Subsequence)",
            "desc": "Xâu A được gọi là xâu con của xâu B nếu ta có thể xóa đi một số ký tự của B (giữ nguyên thứ tự các ký tự còn lại) để thu được xâu A. Cho hai xâu A và B (1 ≤ |A|, |B| ≤ 10^6). Kiểm tra xem A có phải là xâu con của B hay không. In ra `YES` hoặc `NO`.",
            "analysis": "Dùng kỹ thuật Hai con trỏ tham lam: đặt `i = 0` trên xâu A, duyệt tuần tự từng ký tự `ch` của B; mỗi khi `ch == a[i]` thì tăng `i`. Cuối cùng kiểm tra `i == len(a)` trong O(|A| + |B|).",
            "example": "Ví dụ: Input: `hsg` và `hocsinhgioi` → Output: `YES`.",
            "code": """a = input().strip()
b = input().strip()

i = 0
for ch in b:
    if i < len(a) and ch == a[i]:
        i += 1

print("YES" if i == len(a) else "NO")""",
        },
        {
            "title": "Bài 7: Đếm số lượng xâu con đối xứng",
            "desc": "Cho một xâu ký tự S độ dài N (1 ≤ N ≤ 2000). Hãy đếm xem S có tất cả bao nhiêu xâu con liên tiếp là xâu đối xứng (mỗi ký tự đơn lẻ cũng được tính là một xâu đối xứng độ dài 1).",
            "analysis": "Dùng kỹ thuật Mở rộng từ tâm (Expand Around Center): tại mỗi vị trí `i`, mở rộng ra hai phía với tâm lẻ `(i, i)` và tâm chẵn `(i, i + 1)` để đạt độ phức tạp O(N²).",
            "example": 'Ví dụ: Input: `aaa` → Output: `6` (gồm `"a"`, `"a"`, `"a"`, `"aa"`, `"aa"`, `"aaa"`).',
            "code": """s = input().strip()
n = len(s)
dem = 0


def dem_tu_tam(left, right):
    cnt = 0
    while left >= 0 and right < n and s[left] == s[right]:
        cnt += 1
        left -= 1
        right += 1
    return cnt


for i in range(n):
    dem += dem_tu_tam(i, i)      # Xâu đối xứng độ dài lẻ
    dem += dem_tu_tam(i, i + 1)  # Xâu đối xứng độ dài chẵn

print(dem)""",
        },
        {
            "title": "Bài 8: Xâu con dài nhất không chứa ký tự lặp lại",
            "desc": "Cho một xâu ký tự S (1 ≤ |S| ≤ 10^6). Hãy tìm độ dài lớn nhất của một đoạn con liên tiếp trong S sao cho tất cả các ký tự trong đoạn con đó đều đôi một khác nhau.",
            "analysis": "Kỹ thuật Cửa sổ trượt (Sliding Window) kết hợp từ điển `vi_tri_cuoi` lưu vị trí xuất hiện gần nhất của từng ký tự để cập nhật biên `left` trong O(N).",
            "example": 'Ví dụ: Input: `abcabcbb` → Output: `3` (đoạn `"abc"`).',
            "code": """s = input().rstrip("\\r\\n")
vi_tri_cuoi = {}
left = 0
max_len = 0

for right, ch in enumerate(s):
    if ch in vi_tri_cuoi and vi_tri_cuoi[ch] >= left:
        left = vi_tri_cuoi[ch] + 1
    vi_tri_cuoi[ch] = right
    if right - left + 1 > max_len:
        max_len = right - left + 1

print(max_len)""",
        },
        {
            "title": "Bài 9: Xóa K chữ số để được số lớn nhất",
            "desc": "Cho một số nguyên dương có N chữ số được biểu diễn dưới dạng xâu S (1 ≤ N ≤ 10^6, chữ số đầu tiên khác `'0'`) và một số nguyên K (0 ≤ K < N). Hãy xóa đi đúng K chữ số khỏi xâu S (giữ nguyên thứ tự các chữ số còn lại) sao cho số thu được là lớn nhất có thể.",
            "analysis": "Dùng Ngăn xếp đơn điệu (Monotonic Stack): khi còn lượt xóa `k > 0` và đỉnh ngăn xếp nhỏ hơn chữ số hiện tại thì lấy đỉnh ra khỏi ngăn xếp. Nếu sau vòng lặp vẫn còn `k > 0`, cắt bỏ `k` chữ số ở cuối.",
            "example": "Ví dụ: Input: `1924` và `2` → Output: `94`.",
            "code": """s = input().strip()
k = int(input())
stack = []

for ch in s:
    while k > 0 and stack and stack[-1] < ch:
        stack.pop()
        k -= 1
    stack.append(ch)

if k > 0:
    stack = stack[:-k]

print("".join(stack))""",
        },
        {
            "title": "Bài 10: Đoạn con dài nhất chứa tối đa K ký tự phân biệt",
            "desc": "Cho xâu ký tự S (1 ≤ |S| ≤ 10^6) chỉ gồm các chữ cái tiếng Anh in thường và một số nguyên dương K (1 ≤ K ≤ 26). Hãy tìm độ dài lớn nhất của một đoạn con liên tiếp trong S sao cho trong đoạn con đó có không quá K loại ký tự khác nhau.",
            "analysis": "Kỹ thuật Cửa sổ trượt `[left, right]`: mở rộng `right` và cập nhật tần suất vào từ điển `dem`; khi `len(dem) > K` thì thu hẹp `left` cho đến khi số loại ký tự ≤ K trong O(N).",
            "example": 'Ví dụ: Input: `eceba` và `2` → Output: `3` (đoạn `"ece"`).',
            "code": """s = input().strip()
k = int(input())

dem = {}
left = 0
max_len = 0

for right, ch in enumerate(s):
    dem[ch] = dem.get(ch, 0) + 1
    while len(dem) > k:
        left_ch = s[left]
        dem[left_ch] -= 1
        if dem[left_ch] == 0:
            del dem[left_ch]
        left += 1
    if right - left + 1 > max_len:
        max_len = right - left + 1

print(max_len)""",
        },
    ]

    for bt in bai_tap_list:
        add_sub_heading(doc, bt["title"])
        add_rich_paragraph(doc, f"**Đề bài:** {bt['desc']}")
        add_rich_paragraph(
            doc, f"**Phân tích:** {bt['analysis']}", is_bullet=True
        )
        add_rich_paragraph(doc, bt["example"], is_bullet=True)
        add_code_block(doc, bt["code"])

    doc.save(output_filename)
    print(
        f"THÀNH CÔNG! Đã tạo tệp giáo án: {os.path.abspath(output_filename)}"
    )


if __name__ == "__main__":
    OUTPUT_FILE = "XU_LY_XAU.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")