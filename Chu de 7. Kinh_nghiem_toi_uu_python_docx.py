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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 7: Kinh nghiệm tối ưu mã nguồn"
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
# NỘI DUNG CHUYÊN ĐỀ 7: KINH NGHIỆM THỰC CHIẾN TỐI ƯU MÃ NGUỒN PYTHON
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 7: KINH NGHIỆM THỰC CHIẾN TỐI ƯU MÃ NGUỒN PYTHON TRONG THI HSG"
    )
    trun.font.name = "Times New Roman"
    trun.font.size = Pt(15)
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
        "Python có ưu thế vượt trội về tốc độ viết mã, xử lý số nguyên lớn vô hạn và thư viện chuẩn phong phú. Tuy nhiên, do là ngôn ngữ thông dịch, nếu thí sinh viết mã theo thói quen thông thường thì rất dễ gặp lỗi quá thời gian (TLE - Time Limit Exceeded) hoặc tràn bộ nhớ ngăn xếp (RTE - Runtime Error) khi dữ liệu đạt ngưỡng N ≥ 10^5.",
    )

    add_sub_heading(doc, "1. Kỹ thuật Đọc/Ghi dữ liệu lớn (Fast I/O)")
    add_rich_paragraph(
        doc,
        "**Đọc toàn bộ dữ liệu một lần (Nhanh nhất cho mảng, ma trận, đồ thị):** Dùng `sys.stdin.buffer.read().split()` (đọc trực tiếp luồng byte, bỏ qua bước giải mã UTF-8) hoặc `sys.stdin.read().split()`, kết hợp `list(map(int, ...))`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Đọc từng dòng nhanh:** Khi cần giữ nguyên dấu cách trên từng dòng, khai báo `input = sys.stdin.readline` và dùng `input().rstrip(\"\\r\\n\")`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Ghi kết quả siêu tốc (Fast Output):** Không gọi `print()` trong vòng lặp chạy 10^5 lần. Hãy lưu từng dòng kết quả vào danh sách `out` rồi xuất một lần duy nhất bằng `sys.stdout.write(\"\\n\".join(out) + \"\\n\")`.",
        is_bullet=True,
    )

    add_sub_heading(
        doc, "2. Xử lý giới hạn đệ quy (`sys.setrecursionlimit`) & Khử đệ quy"
    )
    add_rich_paragraph(
        doc,
        "**Tăng giới hạn đệ quy:** Giới hạn mặc định của Python chỉ là **1000** tầng. Với bài toán duyệt cây/đồ thị lớn (N ≤ 2·10^5), bắt buộc thêm `sys.setrecursionlimit(300000)` ở đầu tệp.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Khử đệ quy bằng `stack` (Chuẩn mực cao nhất):** Ngay cả khi đã gọi `setrecursionlimit`, đệ quy sâu vẫn tiêu tốn nhiều bộ nhớ ngăn xếp C và chậm hơn vòng lặp. Với cây/đồ thị lớn, dùng `stack` lưu thứ tự duyệt `order` rồi duyệt ngược `reversed(order)` giúp chạy nhanh gấp đôi và an toàn 100%.",
        is_bullet=True,
    )

    add_sub_heading(
        doc, "3. Tư duy tối ưu độ phức tạp & Tận dụng lõi C của Python"
    )
    add_rich_paragraph(
        doc,
        "**Loại bỏ phép toán O(N) trong vòng lặp lớn:** Không dùng `list.insert(0, x)` hoặc `list.pop(0)` (thay bằng `collections.deque` trong O(1)); không dùng `x in list` (thay bằng `set`/`dict` trong O(1)); không nối xâu `s += ch` (thay bằng `list.append` + `\"\".join`).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Đóng gói toàn bộ thuật toán vào trong hàm `def solve():`:** Trong máy ảo Python, việc truy xuất biến cục bộ bên trong hàm (mã lệnh `LOAD_FAST`) nhanh hơn từ **20% đến 30%** so với truy xuất biến toàn cục ở ngoài cùng tệp (`LOAD_GLOBAL`).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Tận dụng `bytearray` và các hàm lõi C:** Dùng `bytearray(n)` thay cho `[False] * n` để tiết kiệm 8 lần RAM; ưu tiên dùng `sum()`, `min()`, `max()`, `math.isqrt()`, `math.gcd()`, `bisect`, `heapq`.",
        is_bullet=True,
    )

    add_sub_heading(doc, '4. Bốn "cạm bẫy" hiệu năng thường gặp trên máy chấm')
    add_rich_paragraph(
        doc,
        "**Viết vòng lặp lớn ở ngoài hàm (Global scope):** Làm chậm tốc độ truy cập biến trong mọi vòng lặp lồng nhau. Luôn đặt mã giải trong `def solve():`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Dùng `try ... except` kiểm tra điều kiện trong vòng lặp triệu lần:** Cơ chế bắt ngoại lệ của Python rất chậm khi lỗi xảy ra liên tục. Luôn dùng `if` kiểm tra biên trước.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Cấp phát lại mảng lớn nhiều lần trong vòng lặp:** Khi cần đặt lại mảng đánh dấu nhiều lần, dùng `visited[:] = bytearray(n + 1)` thay vì tạo đối tượng mới liên tục.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Dùng `sys.stdin.read()` cho bài toán tương tác (Interactive):** Bài tương tác yêu cầu xuất ra và đọc phản hồi từng bước nên bắt buộc dùng `input()` và `print(..., flush=True)`.",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(
        doc, "Ví dụ 1 (Fast I/O & Phạm vi cục bộ): Xử lý 10^6 truy vấn tổng đoạn con"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^9) và Q truy vấn (1 ≤ Q ≤ 10^6). Mỗi truy vấn gồm hai số nguyên L, R (1 ≤ L ≤ R ≤ N). Hãy in ra tổng các phần tử từ vị trí L đến vị trí R.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy tối ưu:",
        [
            "Đóng gói toàn bộ xử lý vào hàm `solve()`, đọc luồng byte bằng `sys.stdin.buffer.read().split()`.",
            "Xây dựng mảng tổng tiền tố `pref` trong O(N), trả lời từng truy vấn trong O(1) và xuất toàn bộ 10^6 dòng kết quả một lần bằng `sys.stdout.write(\"\\n\".join(out))`.",
        ],
    )
    add_code_block(
        doc,
        """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    q = int(data[1])

    pref = [0] * (n + 1)
    cur = 0
    for i in range(1, n + 1):
        cur += int(data[i + 1])
        pref[i] = cur

    out = []
    idx = n + 2
    for _ in range(q):
        L = int(data[idx])
        R = int(data[idx + 1])
        idx += 2
        out.append(str(pref[R] - pref[L - 1]))

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
    )

    # Ví dụ 2
    add_sub_heading(
        doc, "Ví dụ 2 (Xử lý giới hạn đệ quy trên cây lớn): Tính kích thước các cây con"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho một cây gồm N đỉnh đánh số từ 1 đến N (1 ≤ N ≤ 2·10^5), gốc của cây là đỉnh 1. Với mỗi đỉnh u (1 ≤ u ≤ N), hãy tính số lượng đỉnh nằm trong cây con gốc u (tính cả chính đỉnh u).",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy tối ưu:",
        [
            "Thay vì đệ quy dễ gây tràn ngăn xếp khi cây suy biến thành đường thẳng 2·10^5 đỉnh, ta dùng `stack` lưu thứ tự duyệt `order` và mảng `parent`.",
            "Duyệt ngược `reversed(order)` từ lá lên gốc để cộng dồn `sub_size[parent[u]] += sub_size[u]`. Tốc độ cực nhanh và an toàn 100%.",
        ],
    )
    add_code_block(
        doc,
        """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    adj = [[] for _ in range(n + 1)]
    idx = 1
    for _ in range(n - 1):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        adj[u].append(v)
        adj[v].append(u)

    parent = [0] * (n + 1)
    order = []
    stack = [1]
    parent[1] = -1

    while stack:
        u = stack.pop()
        order.append(u)
        for v in adj[u]:
            if v != parent[u]:
                parent[v] = u
                stack.append(v)

    sub_size = [1] * (n + 1)
    for u in reversed(order):
        p = parent[u]
        if p != -1:
            sub_size[p] += sub_size[u]

    sys.stdout.write(" ".join(map(str, sub_size[1:])) + "\\n")


if __name__ == "__main__":
    solve()""",
    )

    # Ví dụ 3
    add_sub_heading(
        doc, "Ví dụ 3 (Tối ưu phép toán trong vòng lặp): Hàng đợi biến đổi hai đầu"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Ban đầu có một dãy rỗng. Thực hiện Q thao tác (1 ≤ Q ≤ 5·10^5): `1 x` (chèn x vào đầu dãy), `2 x` (chèn x vào cuối dãy), `3` (xóa đầu dãy), `4` (xóa cuối dãy). Sau Q thao tác, in ra dãy còn lại từ trái sang phải (nếu rỗng in `EMPTY`).",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy tối ưu:",
        [
            "Tránh tuyệt đối `list.insert(0, x)` và `list.pop(0)` (O(Q) mỗi lần → O(Q²) tổng thể).",
            "Dùng `collections.deque` với `appendleft`, `append`, `popleft`, `pop` đạt O(1) cho mọi thao tác.",
        ],
    )
    add_code_block(
        doc,
        """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    q = int(data[0])
    dq = deque()
    idx = 1

    for _ in range(q):
        op = data[idx]
        idx += 1
        if op == b"1":
            dq.appendleft(data[idx].decode())
            idx += 1
        elif op == b"2":
            dq.append(data[idx].decode())
            idx += 1
        elif op == b"3":
            if dq:
                dq.popleft()
        else:
            if dq:
                dq.pop()

    if dq:
        sys.stdout.write(" ".join(dq) + "\\n")
    else:
        sys.stdout.write("EMPTY\\n")


if __name__ == "__main__":
    solve()""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1 (Fast I/O & Xử lý EOF): Tính tổng luồng số nguyên lớn không biết trước số dòng",
            "desc": "Đầu vào gồm một số lượng không xác định các dòng, mỗi dòng chứa một hoặc nhiều số nguyên cách nhau bởi khoảng trắng (tổng cộng có thể lên tới 2·10^6 số nguyên, mỗi số có trị tuyệt đối ≤ 10^9). Hãy tính tổng tất cả các số nguyên có trong tệp đầu vào.",
            "analysis": "Dùng `sys.stdin.buffer.read().split()` để đọc toàn bộ luồng byte đến hết tệp (EOF), sau đó dùng hàm lõi C `sum(map(int, data))` để cộng toàn bộ mảng trong chưa đầy 0.2 giây.",
            "example": "Ví dụ: Input gồm 2 dòng `10 20` và `30 -5 15` → Output: `70`.",
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    ans = sum(map(int, data))
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 2 (Fast I/O & `bisect`): Xử lý 5·10^5 truy vấn tìm kiếm nhị phân",
            "desc": "Cho dãy gồm N số nguyên không giảm a_1 ≤ a_2 ≤ ... ≤ a_N (1 ≤ N ≤ 5·10^5; |a_i| ≤ 10^9) và Q truy vấn (1 ≤ Q ≤ 5·10^5). Mỗi truy vấn cho một số nguyên X, hãy cho biết có bao nhiêu phần tử trong dãy có giá trị đúng bằng X.",
            "analysis": "Đọc dữ liệu bằng `sys.stdin.buffer.read().split()`, dùng `bisect_right(a, x) - bisect_left(a, x)` để trả lời mỗi truy vấn trong O(log N) và xuất kết quả bằng `\"\\n\".join(out)`.",
            "example": "Ví dụ: Input: `6 3`, dãy `1 2 2 2 5 8` và truy vấn `2`, `5`, `4` → Output: `3`, `1`, `0`.",
            "code": """from bisect import bisect_left, bisect_right
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    q = int(data[1])
    a = list(map(int, data[2 : 2 + n]))

    out = []
    idx = 2 + n
    for _ in range(q):
        x = int(data[idx])
        idx += 1
        cnt = bisect_right(a, x) - bisect_left(a, x)
        out.append(str(cnt))

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 3 (Tối ưu `deque` thay `list.pop(0)`): Trò chơi loại bỏ thẻ bài",
            "desc": "Trên bàn có một chồng gồm N thẻ bài đánh số từ 1 đến N theo thứ tự từ trên xuống dưới (1 ≤ N ≤ 10^6). Chừng nào còn ít nhất 2 thẻ bài, lặp lại quy tắc: vứt bỏ thẻ trên cùng, sau đó chuyển thẻ trên cùng tiếp theo xuống dưới cùng. Tìm thẻ bài cuối cùng còn lại.",
            "analysis": "Dùng `collections.deque` với `dq.popleft()` và `dq.append(dq.popleft())` để đạt độ phức tạp O(N) thay vì O(N²) khi dùng `list.pop(0)`.",
            "example": "Ví dụ: Input: `7` → Output: `6`.",
            "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    dq = deque(range(1, n + 1))

    while len(dq) > 1:
        dq.popleft()
        dq.append(dq.popleft())

    sys.stdout.write(f"{dq[0]}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 4 (Tối ưu `deque` + `set`): Hàng đợi nhận thông báo không trùng lặp",
            "desc": "Một màn hình hiển thị tối đa K thông báo gần nhất (1 ≤ K ≤ 2·10^5). Có N tin nhắn lần lượt gửi tới từ các mã người dùng id_1, ..., id_N (1 ≤ N ≤ 5·10^5; 1 ≤ id_i ≤ 10^9): nếu id_i đã có trên màn hình thì bỏ qua; nếu chưa có thì chèn id_i lên đầu màn hình (nếu đủ K thông báo thì đẩy thông báo cũ nhất ở cuối ra). In ra số lượng và danh sách id trên màn hình từ trên xuống dưới.",
            "analysis": "Kết hợp `deque` và tập hợp `on_screen = set()` để kiểm tra sự tồn tại của `msg_id` trong O(1), tránh phép tìm kiếm O(K) trên `deque`.",
            "example": "Ví dụ: Input: `7 3` và dãy `1 2 3 2 1 4 5` → Output: dòng 1 là `3`, dòng 2 là `5 4 1`.",
            "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    k = int(data[1])

    dq = deque()
    on_screen = set()

    for i in range(2, 2 + n):
        msg_id = int(data[i])
        if msg_id not in on_screen:
            if len(dq) == k:
                removed = dq.pop()
                on_screen.remove(removed)
            dq.appendleft(msg_id)
            on_screen.add(msg_id)

    sys.stdout.write(f"{len(dq)}\\n")
    sys.stdout.write(" ".join(map(str, dq)) + "\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 5 (Xử lý đệ quy sâu trên Cây): Tính đường kính của cây",
            "desc": "Cho một cây gồm N đỉnh (1 ≤ N ≤ 2·10^5) và N - 1 cạnh vô hướng. Đường kính của cây là khoảng cách (số cạnh) lớn nhất giữa hai đỉnh bất kỳ trên cây. Hãy tìm đường kính của cây.",
            "analysis": "Dùng thuật toán 2 lần duyệt DFS (từ đỉnh 1 tìm đỉnh A xa nhất, từ A tìm đỉnh B xa nhất) kết hợp `sys.setrecursionlimit(300000)` để tránh lỗi tràn ngăn xếp khi cây suy biến thành đường thẳng.",
            "example": "Ví dụ: Input: `5` và các cạnh `(1, 2), (1, 3), (3, 4), (3, 5)` → Output: `3`.",
            "code": """import sys

sys.setrecursionlimit(300000)


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    adj = [[] for _ in range(n + 1)]
    idx = 1
    for _ in range(n - 1):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        adj[u].append(v)
        adj[v].append(u)

    def get_farthest(start):
        best_node = start
        max_dist = -1

        def dfs(u, p, d):
            nonlocal best_node, max_dist
            if d > max_dist:
                max_dist = d
                best_node = u
            for v in adj[u]:
                if v != p:
                    dfs(v, u, d + 1)

        dfs(start, 0, 0)
        return best_node, max_dist

    node_a, _ = get_farthest(1)
    _, diameter = get_farthest(node_a)
    sys.stdout.write(f"{diameter}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 6 (Quy hoạch động trên Cây - Tree DP): Tập độc lập cực đại có trọng số",
            "desc": "Một công ty có cấu trúc nhân sự dạng cây gồm N nhân viên (1 ≤ N ≤ 2·10^5), giám đốc là đỉnh 1. Nhân viên thứ i có chỉ số cống hiến w_i (1 ≤ w_i ≤ 10^9). Không được chọn đồng thời cả quản lý trực tiếp và nhân viên cấp dưới trực tiếp vào hội đồng. Tìm tổng chỉ số cống hiến lớn nhất.",
            "analysis": "Dùng DFS khử đệ quy duyệt ngược từ lá lên gốc: `dp1[u]` là tổng lớn nhất khi chọn u (`w[u] + sum(dp0[v])`), `dp0[u]` là tổng lớn nhất khi không chọn u (`sum(max(dp0[v], dp1[v]))`).",
            "example": "Ví dụ: Input: `4`, trọng số `10 20 30 40`, cạnh `(1, 2), (1, 3), (3, 4)` → Output: `60`.",
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    w = [0] + list(map(int, data[1 : 1 + n]))

    adj = [[] for _ in range(n + 1)]
    idx = 1 + n
    for _ in range(n - 1):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        adj[u].append(v)
        adj[v].append(u)

    parent = [0] * (n + 1)
    order = []
    stack = [1]
    parent[1] = -1

    while stack:
        u = stack.pop()
        order.append(u)
        for v in adj[u]:
            if v != parent[u]:
                parent[v] = u
                stack.append(v)

    dp0 = [0] * (n + 1)
    dp1 = w[:]

    for u in reversed(order):
        p = parent[u]
        if p != -1:
            dp1[p] += dp0[u]
            dp0[p] += dp1[u] if dp1[u] > dp0[u] else dp0[u]

    ans = dp1[1] if dp1[1] > dp0[1] else dp0[1]
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 7 (Euler Tour trên Cây + Fast I/O): Kiểm tra quan hệ tổ tiên trong O(1)",
            "desc": "Cho một cây gồm N đỉnh (1 ≤ N ≤ 2·10^5), gốc là đỉnh 1, và Q truy vấn (1 ≤ Q ≤ 5·10^5). Mỗi truy vấn gồm hai đỉnh u, v. Kiểm tra xem u có phải là tổ tiên của v hay không (mỗi đỉnh cũng là tổ tiên của chính nó). In ra `YES` hoặc `NO`.",
            "analysis": "Ghi nhận thời điểm vào `tin[u]` và thoát `tout[u]` bằng DFS với `sys.setrecursionlimit(300000)`. Đỉnh u là tổ tiên của v khi và chỉ khi `tin[u] <= tin[v] <= tout[u]` trong O(1).",
            "example": "Ví dụ: Input: `5 3`, cạnh `(1, 2), (1, 3), (3, 4), (3, 5)`, truy vấn `(1, 4), (2, 5), (3, 3)` → Output: `YES`, `NO`, `YES`.",
            "code": """import sys

sys.setrecursionlimit(300000)


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    q = int(data[1])

    adj = [[] for _ in range(n + 1)]
    idx = 2
    for _ in range(n - 1):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        adj[u].append(v)
        adj[v].append(u)

    tin = [0] * (n + 1)
    tout = [0] * (n + 1)
    timer = 0

    def dfs(u, p):
        nonlocal timer
        timer += 1
        tin[u] = timer
        for v in adj[u]:
            if v != p:
                dfs(v, u)
        tout[u] = timer

    dfs(1, 0)

    out = []
    for _ in range(q):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        if tin[u] <= tin[v] <= tout[u]:
            out.append("YES")
        else:
            out.append("NO")

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 8 (Tối ưu bộ nhớ & tốc độ với `bytearray`): Đếm số nguyên tố dạng 4k + 1",
            "desc": "Cho số nguyên dương N (1 ≤ N ≤ 10^7). Hãy đếm xem có bao nhiêu số nguyên tố p ≤ N thỏa mãn p ≡ 1 (mod 4).",
            "analysis": "Dùng `bytearray` và phép gán lát cắt để sàng nguyên tố đến 10^7 trong 0.12 giây, sau đó dùng `sum(is_prime[5 : n + 1 : 4])` để đếm siêu tốc bằng C-loop.",
            "example": "Ví dụ: Input: `20` → Output: `3` (gồm 5, 13, 17).",
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    if n < 5:
        sys.stdout.write("0\\n")
        return

    is_prime = bytearray(b"\\x01") * (n + 1)
    is_prime[0] = is_prime[1] = 0
    for i in range(2, int(n**0.5) + 1):
        if is_prime[i]:
            is_prime[i * i : n + 1 : i] = bytearray((n - i * i) // i + 1)

    ans = sum(is_prime[5 : n + 1 : 4])
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 9 (Tư duy giảm độ phức tạp từ O(N²) xuống O(N log N)): Tổng khoảng cách Manhattan mọi cặp điểm",
            "desc": "Trên mặt phẳng tọa độ Oxy, cho N điểm có tọa độ nguyên (x_i, y_i) (2 ≤ N ≤ 5·10^5; |x_i|, |y_i| ≤ 10^9). Hãy tính tổng khoảng cách Manhattan |x_i - x_j| + |y_i - y_j| giữa tất cả các cặp điểm 1 ≤ i < j ≤ N.",
            "analysis": "Tách độc lập theo hai trục Ox và Oy. Sau khi sắp xếp mảng tọa độ tăng dần, phần tử tại chỉ số i đóng góp đúng `(2 * i - n + 1) * val` vào tổng khoảng cách trong O(N log N).",
            "example": "Ví dụ: Input: `3` và các điểm `(0, 0), (1, 2), (3, 1)` → Output: `10`.",
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    xs = []
    ys = []
    idx = 1
    for _ in range(n):
        xs.append(int(data[idx]))
        ys.append(int(data[idx + 1]))
        idx += 2

    def sum_1d(arr):
        arr.sort()
        total = 0
        for i, val in enumerate(arr):
            total += (2 * i - n + 1) * val
        return total

    ans = sum_1d(xs) + sum_1d(ys)
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "Bài 10 (Tổng hợp thực chiến): Đảo ngược xâu và Thêm ký tự hai đầu tốc độ O(1)",
            "desc": "Cho một xâu ban đầu S (|S| ≤ 10^5) và Q thao tác (1 ≤ Q ≤ 5·10^5): `1` (đảo ngược toàn bộ xâu hiện tại), `2 1 c` (thêm ký tự c vào đầu xâu), `2 2 c` (thêm ký tự c vào cuối xâu). In ra xâu kết quả cuối cùng.",
            "analysis": "Dùng `collections.deque` kết hợp biến cờ `rev = False` để lật trạng thái trong O(1) thay vì sao chép đảo xâu O(N) mỗi lần; chỉ đảo thực sự tối đa 1 lần duy nhất ở cuối chương trình.",
            "example": "Ví dụ: Input: `abc`, `4`, `1`, `2 1 x`, `2 2 y`, `1` → Output: `ycbax`.",
            "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    s = data[0].decode()
    q = int(data[1])

    dq = deque(s)
    rev = False
    idx = 2

    for _ in range(q):
        op = data[idx]
        idx += 1
        if op == b"1":
            rev = not rev
        else:
            pos = data[idx]
            ch = data[idx + 1].decode()
            idx += 2
            if (pos == b"1" and not rev) or (pos == b"2" and rev):
                dq.appendleft(ch)
            else:
                dq.append(ch)

    if rev:
        dq.reverse()

    sys.stdout.write("".join(dq) + "\\n")


if __name__ == "__main__":
    solve()""",
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
    OUTPUT_FILE = "KINH_NGHIEM_TOI_UU_PYTHON.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")