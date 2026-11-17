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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 6: Lý thuyết số & Lý thuyết đồ thị"
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
# NỘI DUNG CHUYÊN ĐỀ 6: LÝ THUYẾT SỐ & LÝ THUYẾT ĐỒ THỊ
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 6: LÝ THUYẾT SỐ, TOÁN HỌC TÍNH TOÁN & LÝ THUYẾT ĐỒ THỊ (PYTHON)"
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
        "Trong cấu trúc đề thi Học sinh giỏi (HSG) Tin học cấp Tỉnh và Quốc gia, Lý thuyết số & Toán học tính toán cùng Lý thuyết đồ thị là hai mảng kiến thức xương sống, thường xuất hiện ở các bài toán phân loại thí sinh (chiếm 40% - 60% tổng số điểm đề thi). Việc kết hợp tư duy toán học chặt chẽ với các thư viện tối ưu của Python (`math`, `collections`, `heapq`) giúp thí sinh đạt tốc độ thực thi cao trên các bộ dữ liệu lớn.",
    )

    add_sub_heading(doc, "1. Lý thuyết số & Toán học tính toán")
    add_rich_paragraph(
        doc,
        "**Ước chung lớn nhất & Bội chung nhỏ nhất:** Sử dụng trực tiếp hàm tích hợp sẵn viết bằng C: `math.gcd(a, b)` và `math.lcm(a, b)` với độ phức tạp O(log(min(a, b))).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Sàng nguyên tố Eratosthenes tối ưu `bytearray`:** Tìm tất cả các số nguyên tố ≤ N (N ≤ 10^7) trong O(N log log N) bằng phép gán lát cắt (`is_prime[i*i : n+1 : i] = bytearray(...)`), chạy dưới 0.15 giây và chỉ tốn 10 MB RAM.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Sàng ước nguyên tố nhỏ nhất (SPF - Smallest Prime Factor):** Lưu `spf[x]` là ước nguyên tố nhỏ nhất của x, cho phép phân tích thừa số nguyên tố của mọi số X ≤ 10^6 chỉ trong **O(log X)**.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Lũy thừa nhị phân & Nghịch đảo mô-đun:** Hàm có sẵn `pow(a, b, mod)` tính a^b mod M trong **O(log b)**; khi M là số nguyên tố và gcd(a, M) = 1, nghịch đảo mô-đun a^(-1) mod M được tính trực tiếp bằng `pow(a, -1, mod)`.",
        is_bullet=True,
    )

    add_sub_heading(
        doc, "2. Lý thuyết đồ thị (Graph Theory) & Các thuật toán cốt lõi"
    )
    add_rich_paragraph(
        doc,
        "**Biểu diễn đồ thị thưa (N ≤ 10^5, M ≤ 2·10^5):** Sử dụng Danh sách kề `adj = [[] for _ in range(n + 1)]`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Duyệt theo chiều rộng (BFS):** Cài đặt bằng `collections.deque`. Tìm đường đi ngắn nhất trên đồ thị không trọng số, kiểm tra tính liên thông và đồ thị hai phía trong **O(V + E)**.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Duyệt theo chiều sâu (DFS) & Sắp xếp Tô-pô (Thuật toán Kahn):** Dùng mảng bán bậc vào `in_deg` và hàng đợi các đỉnh có `in_deg == 0` để sắp xếp thứ tự công việc và phát hiện chu trình có hướng trong **O(V + E)**.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Đường đi ngắn nhất có trọng số không âm (Dijkstra):** Kết hợp danh sách kề và hàng đợi ưu tiên `heapq` (Min-Heap) với điều kiện lọc `if d > dist[u]: continue`, đạt độ phức tạp **O((V + E) log V)**.",
        is_bullet=True,
    )

    add_sub_heading(doc, '3. Bốn "cạm bẫy" kinh điển trong Python cần tránh')
    add_rich_paragraph(
        doc,
        "**Bẫy tràn ngăn xếp đệ quy (`RecursionError`) khi chạy DFS trên đồ thị lớn (V ≥ 1000):** Python mặc định giới hạn 1000 tầng đệ quy. Luôn ưu tiên cài đặt DFS khử đệ quy bằng `stack` (hoặc BFS) hoặc khai báo `sys.setrecursionlimit(300000)`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy chia lấy dư trực tiếp trong tổ hợp:** Tuyệt đối không viết `(fact[n] // (fact[k] * fact[n - k])) % MOD` sau khi đã lấy dư `fact[i] % MOD`. Phải nhân với nghịch đảo mô-đun `pow(fact[k], -1, MOD)`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy đồ thị không liên thông hoặc có đa cạnh:** Luôn duyệt `for u in range(1, n + 1):` để không bỏ sót các thành phần liên thông tách rời.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy nhập xuất chậm với dữ liệu đồ thị lớn (M ≥ 10^5):** Bắt buộc dùng `sys.stdin.read().split()` để đọc toàn bộ danh sách cạnh trong O(M).",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(
        doc, "Ví dụ 1 (Lý thuyết số): Đếm số nguyên tố trong đoạn [L, R] cho Q truy vấn"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho Q truy vấn (1 ≤ Q ≤ 10^5), mỗi truy vấn gồm hai số nguyên L, R (1 ≤ L ≤ R ≤ 10^6). Với mỗi truy vấn, hãy cho biết có bao nhiêu số nguyên tố thuộc đoạn [L, R].",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Dùng Sàng Eratosthenes tối ưu `bytearray` đánh dấu toàn bộ số nguyên tố đến 10^6 trong O(MAX log log MAX).",
            "Xây dựng mảng cộng dồn `pref[i]` lưu số lượng số nguyên tố từ 1 đến i để trả lời mỗi truy vấn trong O(1) bằng `pref[R] - pref[L - 1]`.",
        ],
    )
    add_code_block(
        doc,
        """import sys

MAX = 1000000
is_prime = bytearray(b"\\x01") * (MAX + 1)
is_prime[0] = is_prime[1] = 0
for i in range(2, int(MAX**0.5) + 1):
    if is_prime[i]:
        is_prime[i * i : MAX + 1 : i] = bytearray((MAX - i * i) // i + 1)

pref = [0] * (MAX + 1)
cnt = 0
for i in range(1, MAX + 1):
    if is_prime[i]:
        cnt += 1
    pref[i] = cnt

data = sys.stdin.read().split()
q = int(data[0])
out = []
idx = 1
for _ in range(q):
    L = int(data[idx])
    R = int(data[idx + 1])
    idx += 2
    out.append(str(pref[R] - pref[L - 1]))

print("\\n".join(out))""",
    )

    # Ví dụ 2
    add_sub_heading(
        doc, "Ví dụ 2 (Đồ thị - BFS trên lưới): Thoát khỏi mê cung ngắn nhất"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho một mê cung kích thước M × N (1 ≤ M, N ≤ 1000) gồm các ký tự: `.` (ô trống), `#` (tường chắn), `S` (xuất phát) và `E` (thoát hiểm). Mỗi bước được đi sang 4 ô chung cạnh. Tìm số bước ít nhất từ `S` đến `E` (in `-1` nếu không đến được).",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Mỗi ô trên lưới là một đỉnh của đồ thị không trọng số. Sử dụng thuật toán BFS với `collections.deque` và mảng `dist` kích thước M × N khởi tạo bằng `-1` để vừa lưu khoảng cách vừa đánh dấu đã thăm trong O(M · N).",
        ],
    )
    add_code_block(
        doc,
        """from collections import deque
import sys

data = sys.stdin.read().split()
m = int(data[0])
n = int(data[1])
grid = data[2 : 2 + m]

dist = [[-1] * n for _ in range(m)]
dq = deque()
end_r, end_c = -1, -1

for r in range(m):
    for c in range(n):
        if grid[r][c] == "S":
            dist[r][c] = 0
            dq.append((r, c))
        elif grid[r][c] == "E":
            end_r, end_c = r, c

dirs = [(-1, 0), (1, 0), (0, -1), (0, 1)]

while dq:
    r, c = dq.popleft()
    if r == end_r and c == end_c:
        break
    d = dist[r][c] + 1
    for dr, dc in dirs:
        nr, nc = r + dr, c + dc
        if 0 <= nr < m and 0 <= nc < n and grid[nr][nc] != "#" and dist[nr][nc] == -1:
            dist[nr][nc] = d
            dq.append((nr, nc))

print(dist[end_r][end_c])""",
    )

    # Ví dụ 3
    add_sub_heading(
        doc, "Ví dụ 3 (Đồ thị - Dijkstra): Đường đi ngắn nhất có trọng số không âm"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho đơn đồ thị có hướng gồm N đỉnh và M cung (1 ≤ N ≤ 10^5; 1 ≤ M ≤ 2·10^5). Cung thứ i đi từ u_i đến v_i có trọng số không âm w_i (0 ≤ w_i ≤ 10^9). Tìm độ dài đường đi ngắn nhất từ đỉnh 1 đến đỉnh N (in `-1` nếu không có đường đi).",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Cài đặt thuật toán Dijkstra bằng danh sách kề `adj` và hàng đợi ưu tiên `heapq`. Luôn kiểm tra `if d > dist[u]: continue` để loại bỏ các trạng thái cũ trong heap, đạt độ phức tạp O((N + M) log N).",
        ],
    )
    add_code_block(
        doc,
        """import heapq
import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])

adj = [[] for _ in range(n + 1)]
idx = 2
for _ in range(m):
    u = int(data[idx])
    v = int(data[idx + 1])
    w = int(data[idx + 2])
    idx += 3
    adj[u].append((v, w))

INF = 10**18
dist = [INF] * (n + 1)
dist[1] = 0
pq = [(0, 1)]

while pq:
    d, u = heapq.heappop(pq)
    if d > dist[u]:
        continue
    if u == n:
        break
    for v, w in adj[u]:
        nd = d + w
        if nd < dist[v]:
            dist[v] = nd
            heapq.heappush(pq, (nd, v))

print(dist[n] if dist[n] < INF else -1)""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1 (Số học - Sàng SPF): Phân tích thừa số nguyên tố nhanh cho Q truy vấn",
            "desc": "Cho Q truy vấn (1 ≤ Q ≤ 10^5), mỗi truy vấn cho một số nguyên X (2 ≤ X ≤ 10^6). Hãy phân tích X ra thừa số nguyên tố theo thứ tự tăng dần dưới dạng `p1^e1 p2^e2 ...`.",
            "analysis": "Dùng sàng Eratosthenes tạo mảng `spf[x]` lưu ước nguyên tố nhỏ nhất của mọi số x ≤ 10^6, sau đó phân tích mỗi số X trong O(log X).",
            "example": "Ví dụ: Input: `2` và các số `60`, `97` → Output: dòng 1 là `2^2 3^1 5^1`, dòng 2 là `97^1`.",
            "code": """import sys

MAX = 1000000
spf = list(range(MAX + 1))
for i in range(2, int(MAX**0.5) + 1):
    if spf[i] == i:
        for j in range(i * i, MAX + 1, i):
            if spf[j] == j:
                spf[j] = i

data = sys.stdin.read().split()
q = int(data[0])
out = []

for i in range(1, q + 1):
    x = int(data[i])
    factors = []
    while x > 1:
        p = spf[x]
        cnt = 0
        while x % p == 0:
            x //= p
            cnt += 1
        factors.append(f"{p}^{cnt}")
    out.append(" ".join(factors))

print("\\n".join(out))""",
        },
        {
            "title": "Bài 2 (Số học - Tổ hợp chia dư): Tính C(N, K) mod (10^9 + 7) cho nhiều truy vấn",
            "desc": "Cho Q truy vấn (1 ≤ Q ≤ 10^5), mỗi truy vấn gồm hai số nguyên N, K (0 ≤ K ≤ N ≤ 10^6). Hãy tính giá trị tổ hợp chập K của N chia lấy dư cho 10^9 + 7.",
            "analysis": "Tiền xử lý mảng giai thừa `fact` và nghịch đảo giai thừa `inv_fact` từ 0 đến 10^6 bằng `pow(fact[MAX], -1, MOD)`, trả lời mỗi truy vấn trong O(1).",
            "example": "Ví dụ: Input: `2`, `5 2`, `10 3` → Output: `10` và `120`.",
            "code": """import sys

MOD = 10**9 + 7
MAX = 1000000

fact = [1] * (MAX + 1)
for i in range(1, MAX + 1):
    fact[i] = fact[i - 1] * i % MOD

inv_fact = [1] * (MAX + 1)
inv_fact[MAX] = pow(fact[MAX], -1, MOD)
for i in range(MAX - 1, -1, -1):
    inv_fact[i] = inv_fact[i + 1] * (i + 1) % MOD

data = sys.stdin.read().split()
q = int(data[0])
out = []
idx = 1

for _ in range(q):
    n = int(data[idx])
    k = int(data[idx + 1])
    idx += 2
    ans = fact[n] * inv_fact[k] % MOD * inv_fact[n - k] % MOD
    out.append(str(ans))

print("\\n".join(out))""",
        },
        {
            "title": "Bài 3 (Số học - Phi hàm Euler): Đếm số nguyên tố cùng nhau với N",
            "desc": "Cho số nguyên dương N (1 ≤ N ≤ 10^12). Hãy đếm xem có bao nhiêu số nguyên dương x (1 ≤ x ≤ N) thỏa mãn gcd(x, N) = 1.",
            "analysis": "Áp dụng công thức Phi hàm Euler: φ(N) = N · Π (1 - 1/p) với p là các ước nguyên tố phân biệt của N, duyệt p ≤ √N trong O(√N).",
            "example": "Ví dụ: Input: `12` → Output: `4` (các số 1, 5, 7, 11).",
            "code": """import math

n = int(input())
res = n
temp = n

for i in range(2, math.isqrt(temp) + 1):
    if temp % i == 0:
        while temp % i == 0:
            temp //= i
        res -= res // i

if temp > 1:
    res -= res // temp

print(res)""",
        },
        {
            "title": "Bài 4 (Số học - Nguyên lý bù trừ & LCM): Đếm bội chung trong đoạn [1, N]",
            "desc": "Cho ba số nguyên dương N, A, B (1 ≤ N, A, B ≤ 10^18). Hãy đếm xem có bao nhiêu số nguyên dương x ≤ N sao cho x chia hết cho A hoặc x chia hết cho B.",
            "analysis": "Áp dụng nguyên lý bù trừ: `(N // A) + (N // B) - (N // math.lcm(A, B))` trong O(log(min(A, B))).",
            "example": "Ví dụ: Input: `15 3 5` → Output: `7` (gồm 3, 5, 6, 9, 10, 12, 15).",
            "code": """import math

n, a, b = map(int, input().split())
boi_chung = math.lcm(a, b)
ans = (n // a) + (n // b) - (n // boi_chung)
print(ans)""",
        },
        {
            "title": "Bài 5 (Toán học tính toán - Lũy thừa ma trận): Số Fibonacci thứ N với N ≤ 10^18",
            "desc": "Dãy Fibonacci được định nghĩa bởi F_1 = 1, F_2 = 1 và F_n = F_{n-1} + F_{n-2} (n ≥ 3). Cho số nguyên dương N (1 ≤ N ≤ 10^18). Hãy tính phần dư của F_N khi chia cho 10^9 + 7.",
            "analysis": "Sử dụng phép lũy thừa nhị phân trên ma trận 2×2 `[[1, 1], [1, 0]]^N` để tính kết quả trong O(log N).",
            "example": "Ví dụ: Input: `50` → Output: `586268941`.",
            "code": """MOD = 10**9 + 7
n = int(input())


def mat_mul(A, B):
    return [
        [
            (A[0][0] * B[0][0] + A[0][1] * B[1][0]) % MOD,
            (A[0][0] * B[0][1] + A[0][1] * B[1][1]) % MOD,
        ],
        [
            (A[1][0] * B[0][0] + A[1][1] * B[1][0]) % MOD,
            (A[1][0] * B[0][1] + A[1][1] * B[1][1]) % MOD,
        ],
    ]


def mat_pow(M, p):
    res = [[1, 0], [0, 1]]  # Ma trận đơn vị
    while p > 0:
        if p & 1:
            res = mat_mul(res, M)
        M = mat_mul(M, M)
        p >>= 1
    return res


ans_mat = mat_pow([[1, 1], [1, 0]], n)
print(ans_mat[0][1])""",
        },
        {
            "title": "Bài 6 (Đồ thị - Liên thông): Đếm thành phần liên thông và Kích thước vùng lớn nhất",
            "desc": "Cho đơn đồ thị vô hướng gồm N đỉnh và M cạnh (1 ≤ N ≤ 10^5; 0 ≤ M ≤ 2·10^5). Hãy tìm số lượng thành phần liên thông của đồ thị và số lượng đỉnh của thành phần liên thông lớn nhất.",
            "analysis": "Duyệt qua các đỉnh 1..N; mỗi khi gặp đỉnh chưa thăm, dùng BFS loang sang toàn bộ thành phần liên thông và đếm kích thước trong O(N + M).",
            "example": "Ví dụ: Input: `6 4` và các cạnh `(1, 2), (2, 3), (4, 5)` → Output: `3 3`.",
            "code": """from collections import deque
import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])

adj = [[] for _ in range(n + 1)]
idx = 2
for _ in range(m):
    u = int(data[idx])
    v = int(data[idx + 1])
    idx += 2
    adj[u].append(v)
    adj[v].append(u)

visited = bytearray(n + 1)
num_components = 0
max_size = 0

for i in range(1, n + 1):
    if not visited[i]:
        num_components += 1
        visited[i] = 1
        dq = deque([i])
        size = 0
        while dq:
            u = dq.popleft()
            size += 1
            for v in adj[u]:
                if not visited[v]:
                    visited[v] = 1
                    dq.append(v)
        if size > max_size:
            max_size = size

print(num_components, max_size)""",
        },
        {
            "title": "Bài 7 (Đồ thị - BFS): Kiểm tra đồ thị hai phía (Bipartite Graph)",
            "desc": "Cho đồ thị vô hướng gồm N đỉnh và M cạnh (1 ≤ N ≤ 10^5; 0 ≤ M ≤ 2·10^5). Hãy kiểm tra xem đồ thị có phải là đồ thị hai phía hay không (có thể tô 2 màu sao cho hai đỉnh kề nhau luôn khác màu). In ra `YES` hoặc `NO`.",
            "analysis": "Dùng BFS tô xen kẽ 2 màu `0` và `1` trên từng thành phần liên thông; nếu gặp cạnh nối hai đỉnh đã tô cùng màu thì trả về `NO` trong O(N + M).",
            "example": "Ví dụ: Input: `4 4` và chu trình 4 đỉnh `(1, 2), (2, 3), (3, 4), (4, 1)` → Output: `YES`.",
            "code": """from collections import deque
import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])

adj = [[] for _ in range(n + 1)]
idx = 2
for _ in range(m):
    u = int(data[idx])
    v = int(data[idx + 1])
    idx += 2
    adj[u].append(v)
    adj[v].append(u)

color = [-1] * (n + 1)
is_bipartite = True

for i in range(1, n + 1):
    if color[i] == -1 and is_bipartite:
        color[i] = 0
        dq = deque([i])
        while dq and is_bipartite:
            u = dq.popleft()
            next_c = 1 - color[u]
            for v in adj[u]:
                if color[v] == -1:
                    color[v] = next_c
                    dq.append(v)
                elif color[v] == color[u]:
                    is_bipartite = False
                    break

print("YES" if is_bipartite else "NO")""",
        },
        {
            "title": "Bài 8 (Đồ thị - Sắp xếp Tô-pô): Phát hiện chu trình trên đồ thị có hướng",
            "desc": "Cho đồ thị có hướng gồm N đỉnh và M cung (1 ≤ N ≤ 10^5; 1 ≤ M ≤ 2·10^5). Hãy kiểm tra xem đồ thị có tồn tại chu trình có hướng nào hay không. In ra `YES` nếu có chu trình, ngược lại in ra `NO`.",
            "analysis": "Áp dụng thuật toán Kahn (khử đệ quy bằng `deque` dựa trên mảng bán bậc vào `in_deg`): nếu số đỉnh lấy ra được nhỏ hơn N thì đồ thị có chu trình. Độ phức tạp O(N + M).",
            "example": "Ví dụ: Input: `3 3` và các cung `(1, 2), (2, 3), (3, 1)` → Output: `YES`.",
            "code": """from collections import deque
import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])

adj = [[] for _ in range(n + 1)]
in_deg = [0] * (n + 1)
idx = 2
for _ in range(m):
    u = int(data[idx])
    v = int(data[idx + 1])
    idx += 2
    adj[u].append(v)
    in_deg[v] += 1

dq = deque(u for u in range(1, n + 1) if in_deg[u] == 0)
visited_count = 0

while dq:
    u = dq.popleft()
    visited_count += 1
    for v in adj[u]:
        in_deg[v] -= 1
        if in_deg[v] == 0:
            dq.append(v)

print("YES" if visited_count < n else "NO")""",
        },
        {
            "title": "Bài 9 (Đồ thị - Sắp xếp Tô-pô + `heapq`): Thứ tự thực hiện công việc nhỏ nhất",
            "desc": "Có N công việc đánh số từ 1 đến N (1 ≤ N ≤ 10^5) và M điều kiện (u, v) (0 ≤ M ≤ 2·10^5) yêu cầu công việc u phải làm xong trước công việc v. Hãy tìm thứ tự thực hiện toàn bộ N công việc có thứ tự từ điển nhỏ nhất (in `-1` nếu không tồn tại).",
            "analysis": "Kết hợp thuật toán Kahn với hàng đợi ưu tiên `heapq` (Min-Heap) thay cho `deque` để luôn chọn đỉnh có chỉ số nhỏ nhất trong số các đỉnh có `in_deg == 0` trong O((N + M) log N).",
            "example": "Ví dụ: Input: `4 2` và các điều kiện `(3, 1), (4, 2)` → Output: `3 1 4 2`.",
            "code": """import heapq
import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])

adj = [[] for _ in range(n + 1)]
in_deg = [0] * (n + 1)
idx = 2
for _ in range(m):
    u = int(data[idx])
    v = int(data[idx + 1])
    idx += 2
    adj[u].append(v)
    in_deg[v] += 1

pq = [u for u in range(1, n + 1) if in_deg[u] == 0]
heapq.heapify(pq)
order = []

while pq:
    u = heapq.heappop(pq)
    order.append(str(u))
    for v in adj[u]:
        in_deg[v] -= 1
        if in_deg[v] == 0:
            heapq.heappush(pq, v)

if len(order) < n:
    print(-1)
else:
    print(" ".join(order))""",
        },
        {
            "title": "Bài 10 (Đồ thị - Dijkstra mở rộng trạng thái): Đường đi ngắn nhất với 1 phiếu giảm giá",
            "desc": "Cho đồ thị có hướng gồm N đỉnh và M cung (2 ≤ N ≤ 10^5; 1 ≤ M ≤ 2·10^5), cung (u, v, w) có chi phí w (0 ≤ w ≤ 10^9). Bạn có đúng 1 phiếu giảm giá cho phép giảm chi phí một cung bất kỳ từ w xuống ⌊w / 2⌋. Tìm chi phí nhỏ nhất từ đỉnh 1 đến đỉnh N.",
            "analysis": "Dijkstra phân tầng (Layered Graph): tách mỗi đỉnh u thành hai trạng thái `dist0[u]` (chưa dùng phiếu) và `dist1[u]` (đã dùng phiếu) trong O((N + M) log N).",
            "example": "Ví dụ: Input: `3 3`, các cung `(1, 2, 3), (2, 3, 1), (1, 3, 7)` → Output: `2`.",
            "code": """import heapq
import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])

adj = [[] for _ in range(n + 1)]
idx = 2
for _ in range(m):
    u = int(data[idx])
    v = int(data[idx + 1])
    w = int(data[idx + 2])
    idx += 3
    adj[u].append((v, w))

INF = 10**18
dist0 = [INF] * (n + 1)
dist1 = [INF] * (n + 1)
dist0[1] = 0

pq = [(0, 1, 0)]

while pq:
    d, u, used = heapq.heappop(pq)
    if used == 0:
        if d > dist0[u]:
            continue
        for v, w in adj[u]:
            if d + w < dist0[v]:
                dist0[v] = d + w
                heapq.heappush(pq, (d + w, v, 0))
            half_w = d + (w // 2)
            if half_w < dist1[v]:
                dist1[v] = half_w
                heapq.heappush(pq, (half_w, v, 1))
    else:
        if d > dist1[u]:
            continue
        for v, w in adj[u]:
            if d + w < dist1[v]:
                dist1[v] = d + w
                heapq.heappush(pq, (d + w, v, 1))

print(min(dist0[n], dist1[n]))""",
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
    OUTPUT_FILE = "LY_THUYET_SO_VA_DO_THI.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")