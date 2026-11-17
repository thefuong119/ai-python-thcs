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
        print(f"Vui lòng mở Terminal trong IDE và chạy: \"{sys.executable}\" -m pip install python-docx")
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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 2: Vòng lặp for"
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
# NỘI DUNG CHUYÊN ĐỀ 2: VÒNG LẶP FOR
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 2: VÒNG LẶP FOR TRONG LẬP TRÌNH THI ĐẤU (PYTHON)"
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
        "Trong lập trình thi đấu bằng Python, vòng lặp `for` là công cụ chủ lực để duyệt dãy số, xử lý xâu ký tự và thực hiện các phép toán số học. Tuy nhiên, do đặc thù thông dịch của Python, việc nắm vững kỹ thuật tối ưu vòng lặp là yếu tố quyết định để chương trình không bị quá thời gian chạy (TLE - Time Limit Exceeded).",
    )

    add_sub_heading(doc, "1. Cú pháp chuẩn và Các công cụ duyệt dữ liệu")
    add_rich_paragraph(
        doc,
        "**Hàm `range(start, stop, step)`:** Sinh ra dãy số nguyên từ `start` đến sát `stop` (không bao gồm `stop`) với bước nhảy `step`. Ví dụ: `range(n)` duyệt từ 0 đến n - 1; `range(a, b + 1)` duyệt từ a đến b; `range(b, a - 1, -1)` duyệt ngược từ b về a.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Duyệt trực tiếp phần tử và chỉ số:** Dùng `for x in arr:` khi chỉ cần giá trị, hoặc `for idx, val in enumerate(arr):` khi cần cả vị trí và giá trị.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Cấu trúc `for ... else`:** Khối `else` chỉ thực thi khi vòng lặp `for` kết thúc bình thường (không bị ngắt giữa chừng bởi lệnh `break`). Rất hữu ích khi kiểm tra tính nguyên tố hoặc tìm kiếm phần tử mà không cần biến đánh dấu:",
        is_bullet=True,
    )
    add_code_block(
        doc,
        """for i in range(2, n):
    if n % i == 0:
        print("HOP SO")
        break
else:
    print("SO NGUYEN TO")""",
        label="",
    )

    add_sub_heading(
        doc, "2. Ngưỡng tốc độ của Python và Tư duy tối ưu độ phức tạp"
    )
    add_rich_paragraph(
        doc,
        "Với giới hạn thời gian tiêu chuẩn 1 giây của máy chấm thi HSG, Python thực hiện được khoảng 10^7 phép lặp đơn giản. Nhìn vào giới hạn N của đề bài, học sinh cần định hình ngay thuật toán:",
    )
    add_rich_paragraph(
        doc,
        "Nếu N ≤ 10^3: Có thể dùng 2 vòng `for` lồng nhau O(N²).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "Nếu N ≤ 10^6: Chỉ được dùng 1 vòng `for` tuyến tính O(N).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "Nếu N ≤ 10^12 hoặc 10^14: Bắt buộc tối ưu vòng lặp xuống O(√N) (duyệt đến `math.isqrt(N)`).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "Nếu N ≤ 10^18: Không được dùng vòng lặp O(N) hay O(√N), phải tìm công thức toán học O(1) hoặc O(log N).",
        is_bullet=True,
    )

    add_sub_heading(
        doc, '3. Bốn "cạm bẫy" kinh điển với vòng lặp `for` cần tránh'
    )
    add_rich_paragraph(
        doc,
        "**Bẫy cận trên của `range()` (Off-by-one error):** Muốn duyệt đến N, luôn viết `range(1, N + 1)`. Muốn duyệt đến √N, luôn dùng căn nguyên `math.isqrt(N)` và viết `range(1, math.isqrt(N) + 1)` thay vì `int(N ** 0.5)` để tránh sai số số thực.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy nối xâu bằng toán tử `+=` trong vòng lặp:** Xâu ký tự (`str`) trong Python là bất biến. Lệnh `s += char` lặp lại 10^5 lần sẽ tạo ra 10^5 bản sao mới, gây độ phức tạp O(N²) dẫn đến TLE. **Giải pháp chuẩn:** Thêm từng phần vào danh sách `res.append(char)` rồi ghép một lần ở cuối bằng `\"\".join(res)` để đạt O(N).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy thay đổi danh sách khi đang duyệt:** Không bao giờ dùng `arr.remove(x)` hoặc `del arr[i]` bên trong vòng lặp `for x in arr:` vì sẽ làm trượt chỉ số và bỏ sót phần tử.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Đọc nhanh dữ liệu lớn (N ≥ 10^5):** Khi đề bài cho nhiều dòng dữ liệu lớn, nên dùng `sys.stdin.readline` thay cho `input()` thông thường.",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(
        doc, "Ví dụ 1: Đếm số ước và Tính tổng các ước của số nguyên dương N"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho số nguyên dương N (1 ≤ N ≤ 10^12). Hãy tính số lượng ước nguyên dương và tổng tất cả các ước nguyên dương của N.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Với N ≤ 10^12, duyệt `for i in range(1, N + 1)` mất 10^12 bước và chắc chắn bị TLE.",
            "Các ước của N luôn đi theo từng cặp (i, N // i) với 1 ≤ i ≤ √N.",
            "Trường hợp biên: Nếu N là số chính phương thì tại i = √N, ta có i = N // i. Khi đó chỉ được cộng 1 lần thay vì 2 lần!",
        ],
    )
    add_code_block(
        doc,
        """import math

n = int(input())
dem_uoc = 0
tong_uoc = 0
can = math.isqrt(n)

for i in range(1, can + 1):
    if n % i == 0:
        j = n // i
        if i == j:
            dem_uoc += 1
            tong_uoc += i
        else:
            dem_uoc += 2
            tong_uoc += i + j

print(dem_uoc, tong_uoc)""",
    )

    # Ví dụ 2
    add_sub_heading(
        doc, "Ví dụ 2: Kiểm tra số nguyên tố tối ưu với `for ... else`"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho số nguyên N (-10^12 ≤ N ≤ 10^12). Kiểm tra xem N có phải là số nguyên tố hay không. In ra `YES` hoặc `NO`.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Loại nhanh các trường hợp biên: N < 2 không phải số nguyên tố; N = 2 là số nguyên tố chẵn duy nhất; mọi số chẵn lớn hơn 2 đều là hợp số.",
            "Với N lẻ (N ≥ 3), chỉ cần kiểm tra các ước lẻ i ∈ {3, 5, 7, ..., ⌊√N⌋} bằng bước nhảy `step = 2`.",
            "Kết hợp cấu trúc `for ... else` giúp mã nguồn gọn gàng, không cần biến đánh dấu.",
        ],
    )
    add_code_block(
        doc,
        """import math

n = int(input())

if n < 2:
    print("NO")
elif n == 2:
    print("YES")
elif n % 2 == 0:
    print("NO")
else:
    for i in range(3, math.isqrt(n) + 1, 2):
        if n % i == 0:
            print("NO")
            break
    else:
        print("YES")""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1: Tổng các bội số có điều kiện trong đoạn [L, R]",
            "desc": "Cho bốn số nguyên dương L, R, A, B (1 ≤ L ≤ R ≤ 10^6; 1 ≤ A, B ≤ 10^6). Hãy tính tổng tất cả các số nguyên x ∈ [L, R] sao cho x chia hết cho A nhưng không chia hết cho B.",
            "analysis": "Tìm bội nhỏ nhất của A ≥ L bằng công thức nguyên `start = ((L + A - 1) // A) * A`, sau đó dùng bước nhảy A trong `range(start, R + 1, A)` để duyệt trực tiếp các bội của A.",
            "example": "Ví dụ: Input: `1 15 3 2` → Output: `27` (3 + 9 + 15 = 27).",
            "code": """L, R, A, B = map(int, input().split())

start = ((L + A - 1) // A) * A
tong = 0

for x in range(start, R + 1, A):
    if x % B != 0:
        tong += x

print(tong)""",
        },
        {
            "title": "Bài 2: Phân loại số hoàn hảo, số phong phú và số thiếu hụt",
            "desc": "Gọi S(N) là tổng các ước thực sự (các ước nguyên dương nhỏ hơn chính nó) của N (1 ≤ N ≤ 10^12). Nếu S(N) = N in ra `HOAN HAO`; nếu S(N) > N in ra `PHONG PHU`; nếu S(N) < N in ra `THIEU HUT`.",
            "analysis": "Trường hợp biên N = 1 có S(1) = 0 < 1 (`THIEU HUT`). Với N > 1, khởi tạo `tong_uoc = 1` và duyệt các cặp ước (i, N // i) từ 2 đến `math.isqrt(N)` trong O(√N).",
            "example": "Ví dụ: Input: `28` (1 + 2 + 4 + 7 + 14 = 28) → Output: `HOAN HAO`.",
            "code": """import math

n = int(input())

if n == 1:
    print("THIEU HUT")
else:
    tong_uoc = 1  # 1 luôn là ước thực sự của mọi n > 1
    can = math.isqrt(n)
    for i in range(2, can + 1):
        if n % i == 0:
            j = n // i
            if i == j:
                tong_uoc += i
            else:
                tong_uoc += i + j

    if tong_uoc == n:
        print("HOAN HAO")
    elif tong_uoc > n:
        print("PHONG PHU")
    else:
        print("THIEU HUT")""",
        },
        {
            "title": "Bài 3: Ước nguyên tố lớn nhất",
            "desc": "Cho số nguyên dương N (2 ≤ N ≤ 10^12). Hãy tìm ước số nguyên tố lớn nhất của N.",
            "analysis": "Chia triệt để N cho 2, sau đó duyệt các ước lẻ i từ 3 đến ⌊√N⌋ bằng vòng lặp `for` và chia triệt để N cho i. Sau vòng lặp, nếu phần còn lại N > 1 thì đó chính là ước nguyên tố lớn nhất.",
            "example": "Ví dụ: Input: `13195` (5 × 7 × 13 × 29) → Output: `29`.",
            "code": """import math

n = int(input())
max_prime = -1

while n % 2 == 0:
    max_prime = 2
    n //= 2

for i in range(3, math.isqrt(n) + 1, 2):
    while n % i == 0:
        max_prime = i
        n //= i

if n > 1:
    max_prime = n

print(max_prime)""",
        },
        {
            "title": "Bài 4: Biểu diễn số thành tổng hai số chính phương",
            "desc": "Cho số nguyên dương N (1 ≤ N ≤ 10^12). Hãy đếm xem có bao nhiêu cặp số nguyên dương (a, b) thỏa mãn 1 ≤ a ≤ b và a² + b² = N.",
            "analysis": "Vì a ≤ b nên 2a² ≤ N ⇒ a ≤ ⌊√(N / 2)⌋. Dùng 1 vòng `for` cho a và kiểm tra b² = N - a² có là số chính phương bằng `math.isqrt()`.",
            "example": "Ví dụ: Input: `50` (1² + 7² và 5² + 5²) → Output: `2`.",
            "code": """import math

n = int(input())
dem = 0
gioi_han = math.isqrt(n // 2)

for a in range(1, gioi_han + 1):
    b_binh_phuong = n - a * a
    b = math.isqrt(b_binh_phuong)
    if b >= a and b * b == b_binh_phuong:
        dem += 1

print(dem)""",
        },
        {
            "title": "Bài 5: Số Fibonacci thứ N chia dư",
            "desc": "Dãy số Fibonacci được định nghĩa bởi F_1 = 1, F_2 = 1 và F_n = F_{n-1} + F_{n-2} với mọi n ≥ 3. Cho số nguyên dương N (1 ≤ N ≤ 10^6). Hãy tính phần dư của F_N khi chia cho 10^9 + 7.",
            "analysis": "Chỉ dùng 2 biến `f1, f2` cập nhật luân phiên trong vòng lặp `for` và chia dư ngay sau mỗi phép cộng để tối ưu bộ nhớ O(1) và tốc độ.",
            "example": "Ví dụ: Input: `10` → Output: `55`.",
            "code": """MOD = 10**9 + 7
n = int(input())

if n <= 2:
    print(1)
else:
    f1, f2 = 1, 1
    for _ in range(3, n + 1):
        f1, f2 = f2, (f1 + f2) % MOD
    print(f2)""",
        },
        {
            "title": "Bài 6: Giá trị lớn thứ hai phân biệt trong dãy",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (2 ≤ N ≤ 10^6; |a_i| ≤ 10^18). Hãy tìm giá trị lớn thứ hai (nghiêm ngặt nhỏ hơn giá trị lớn nhất) trong dãy. Nếu tất cả phần tử bằng nhau thì in ra `KHONG CO`.",
            "analysis": "Duyệt tuyến tính O(N) với hai biến `max1`, `max2` khởi tạo bằng `None` để không bị phụ thuộc vào giới hạn giá trị âm của dữ liệu.",
            "example": "Ví dụ: Input: `5` và dãy `7 3 9 9 5` → Output: `7`.",
            "code": """n = int(input())
a = list(map(int, input().split()))

max1 = None
max2 = None

for x in a:
    if max1 is None or x > max1:
        if max1 is not None and x != max1:
            max2 = max1
        max1 = x
    elif x != max1 and (max2 is None or x > max2):
        max2 = x

if max2 is None:
    print("KHONG CO")
else:
    print(max2)""",
        },
        {
            "title": "Bài 7: Dãy con tăng liên tiếp dài nhất",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^9). Hãy tìm độ dài của đoạn con gồm các phần tử đứng liền nhau sao cho phần tử đứng sau lớn hơn hẳn phần tử đứng trước (a_i < a_{i+1}) có độ dài lớn nhất.",
            "analysis": "Duyệt i từ 1 đến N - 1. Nếu a[i] > a[i - 1] thì tăng `cur_len` và cập nhật `max_len`; ngược lại đặt lại `cur_len = 1`.",
            "example": "Ví dụ: Input: `7` và dãy `1 2 5 3 4 8 9` → Output: `4` (đoạn 3 4 8 9).",
            "code": """n = int(input())
a = list(map(int, input().split()))

max_len = 1
cur_len = 1

for i in range(1, n):
    if a[i] > a[i - 1]:
        cur_len += 1
        if cur_len > max_len:
            max_len = cur_len
    else:
        cur_len = 1

print(max_len)""",
        },
        {
            "title": "Bài 8: Nén xâu ký tự (Run-Length Encoding)",
            "desc": "Cho một xâu ký tự S chỉ gồm các chữ cái tiếng Anh in thường, độ dài không quá 10^6. Hãy nén xâu S bằng cách thay thế mỗi nhóm các ký tự giống nhau đứng liên tiếp bằng ký tự đó kèm theo số lần lặp liên tiếp của nó.",
            "analysis": "Dùng danh sách `ket_qua` lưu các cụm đã nén và ghép một lần ở cuối bằng `\"\".join(ket_qua)` để đạt độ phức tạp O(N), tránh lỗi TLE do nối xâu.",
            "example": "Ví dụ: Input: `aaabbcdddd` → Output: `a3b2c1d4`.",
            "code": """s = input().strip()
ket_qua = []
dem = 1

for i in range(1, len(s)):
    if s[i] == s[i - 1]:
        dem += 1
    else:
        ket_qua.append(f"{s[i - 1]}{dem}")
        dem = 1

# Ghi nhận nhóm ký tự cuối cùng
ket_qua.append(f"{s[-1]}{dem}")

print("".join(ket_qua))""",
        },
        {
            "title": "Bài 9: Đếm bộ ba số nguyên dương có tổng bằng N",
            "desc": "Cho số nguyên dương N (1 ≤ N ≤ 10^7). Hãy đếm xem có bao nhiêu bộ ba số nguyên dương (a, b, c) thỏa mãn 1 ≤ a ≤ b ≤ c và a + b + c = N.",
            "analysis": "Cố định a bằng 1 vòng `for` từ 1 đến ⌊N / 3⌋; khi đó a ≤ b ≤ ⌊(N - a) / 2⌋ nên số cách chọn b ứng với mỗi a tính trực tiếp trong O(1).",
            "example": "Ví dụ: Input: `6` → Output: `3` (gồm (1, 1, 4), (1, 2, 3), (2, 2, 2)).",
            "code": """n = int(input())
dem = 0

for a in range(1, n // 3 + 1):
    b_max = (n - a) // 2
    if b_max >= a:
        dem += b_max - a + 1

print(dem)""",
        },
        {
            "title": "Bài 10: Đếm số tam giác nguyên có chu vi N",
            "desc": "Cho số nguyên dương N (3 ≤ N ≤ 5000). Hãy đếm xem có bao nhiêu tam giác không suy biến có độ dài ba cạnh là các số nguyên dương (a, b, c) thỏa mãn a ≤ b ≤ c và chu vi a + b + c = N.",
            "analysis": "Từ bất đẳng thức tam giác a + b > c = N - a - b, thu hẹp cận dưới của b là `b_min = max(a, n // 2 - a + 1)` và cận trên `b_max = (n - a) // 2`.",
            "example": "Ví dụ: Input: `7` → Output: `2` (gồm (1, 3, 3) và (2, 2, 3)).",
            "code": """n = int(input())
dem = 0

for a in range(1, n // 3 + 1):
    b_min = max(a, n // 2 - a + 1)
    b_max = (n - a) // 2
    for b in range(b_min, b_max + 1):
        c = n - a - b
        if a + b > c:
            dem += 1

print(dem)""",
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
    OUTPUT_FILE = "VONG_LAP_FOR.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")