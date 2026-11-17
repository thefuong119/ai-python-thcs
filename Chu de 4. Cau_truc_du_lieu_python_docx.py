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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 4: Cấu trúc dữ liệu & Thư viện chuẩn"
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
# NỘI DUNG CHUYÊN ĐỀ 4: CẤU TRÚC DỮ LIỆU & THƯ VIỆN CHUẨN
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 4: CẤU TRÚC DỮ LIỆU (LIST, TUPLE, SET, DICT) & THƯ VIỆN COLLECTIONS, ITERTOOLS"
    )
    trun.font.name = "Times New Roman"
    trun.font.size = Pt(15.5)
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
        "Trong các kỳ thi Học sinh giỏi (HSG) Tin học cấp Tỉnh và Quốc gia, việc lựa chọn đúng cấu trúc dữ liệu quyết định trực tiếp đến độ phức tạp thuật toán. Với giới hạn N ≤ 10^6, việc giảm thao tác tìm kiếm hoặc cập nhật từ O(N) xuống O(1) nhờ bảng băm (Hash Table) và hàng đợi hai đầu là chìa khóa để đạt điểm tuyệt đối.",
    )

    add_sub_heading(
        doc, "1. So sánh độ phức tạp của 4 cấu trúc dữ liệu cốt lõi"
    )
    add_rich_paragraph(
        doc,
        "**`list` (Danh sách động):** Truy cập theo chỉ số `a[i]`, thêm/xóa cuối `a.append(x)`, `a.pop()` mất **O(1)**. Tuy nhiên, tìm kiếm `x in a`, đếm `a.count(x)`, thêm/xóa đầu `a.pop(0)`, `a.insert(0, x)` mất **O(N)** → Dễ gây TLE nếu đặt trong vòng lặp.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**`tuple` (Bộ giá trị bất biến):** Chiếm ít bộ nhớ hơn `list` và có thể băm được (*hashable*), do đó dùng làm khóa (`key`) của `dict` hoặc phần tử của `set` (ví dụ: lưu tọa độ `(x, y)`, cặp trạng thái `(u, v)`).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**`set` (Tập hợp) và `dict` (Từ điển):** Cài đặt dựa trên **Bảng băm (Hash Table)**. Các thao tác thêm, xóa, kiểm tra tồn tại `x in s`, truy xuất `d[key]` đạt trung bình **O(1)**. Tập hợp `set` hỗ trợ các phép toán siêu tốc: Hợp (`A | B`), Giao (`A & B`), Hiệu (`A - B`).",
        is_bullet=True,
    )

    add_sub_heading(doc, "2. Khai thác hiệu quả module `collections`")
    add_rich_paragraph(
        doc,
        "**`Counter`:** Lớp con của `dict` chuyên dùng để đếm tần suất xuất hiện trong **O(N)** (`cnt = Counter(arr)`), hỗ trợ `.most_common(k)` lấy top k phần tử phổ biến nhất.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**`defaultdict`:** Tự động khởi tạo giá trị mặc định khi truy cập khóa chưa tồn tại, tránh lỗi `KeyError`: `defaultdict(int)` mặc định `0` (dùng để đếm); `defaultdict(list)` mặc định `[]` (dùng biểu diễn danh sách kề đồ thị hoặc gom nhóm).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**`deque` (Hàng đợi hai đầu):** Khắc phục triệt để điểm yếu O(N) của `list.pop(0)`. Cả 4 thao tác `append`, `appendleft`, `pop`, `popleft` đều đạt **O(1)**, bắt buộc dùng khi cài đặt thuật toán BFS hoặc Cửa sổ trượt đơn điệu.",
        is_bullet=True,
    )

    add_sub_heading(doc, "3. Khai thác hiệu quả module `itertools`")
    add_rich_paragraph(
        doc,
        "**`permutations(iterable, r)`:** Sinh tất cả các **chỉnh hợp không lặp** chập r (nếu bỏ qua r sẽ sinh toàn bộ N! **hoán vị**). Chỉ dùng vét cạn hoán vị khi N ≤ 9 hoặc 10 (10! = 3.628.800).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**`combinations(iterable, r)`:** Sinh tất cả các **tổ hợp** chập r của N phần tử theo thứ tự từ điển (số lượng C(N, r)). Phù hợp để giải quyết trọn vẹn các subtask vét cạn tập con kích thước nhỏ (C(N, r) ≤ 10^6).",
        is_bullet=True,
    )

    add_sub_heading(doc, '4. Bốn "cạm bẫy" kinh điển cần tránh')
    add_rich_paragraph(
        doc,
        "**Dùng `x in list` thay vì `x in set`:** Kiểm tra phần tử thuộc danh sách dài N bên trong vòng lặp N lần đẩy độ phức tạp lên O(N²). Luôn chuyển danh sách tra cứu sang `set` trước khi lặp.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Dùng `list.pop(0)` làm hàng đợi BFS:** Với đồ thị 10^5 đỉnh, `list.pop(0)` mất O(V²) gây TLE. Phải dùng `collections.deque` và `popleft()`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Đưa `list` vào `set` hoặc làm `key` của `dict`:** Gây lỗi `TypeError: unhashable type: 'list'`. Phải chuyển sang `tuple(arr)`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy đụng độ băm (Anti-hash test):** Trên một số hệ thống chấm thi, với số nguyên lớn làm khóa của `dict`/`set`, có thể trộn thêm một hằng số ngẫu nhiên bằng phép XOR (`x ^ MASK`) để chống đụng độ băm.",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(doc, "Ví dụ 1: Đếm số cặp phần tử có tổng bằng K")
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (2 ≤ N ≤ 10^6; |a_i| ≤ 10^12) và một số nguyên K (|K| ≤ 10^12). Hãy đếm xem có bao nhiêu cặp chỉ số (i, j) thỏa mãn 1 ≤ i < j ≤ N và a_i + a_j = K.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Dùng từ điển `dem` lưu số lần xuất hiện của các phần tử đã duyệt qua ở phía trước chỉ số hiện tại.",
            "Khi xét phần tử `x`, số lượng phần tử đứng trước nó có giá trị bằng `K - x` chính là `dem.get(K - x, 0)`. Cộng giá trị này vào đáp án rồi mới tăng `dem[x]` lên 1.",
            "Cách làm này xử lý chuẩn xác cả trường hợp x = K - x mà không bị đếm trùng, đạt độ phức tạp O(N).",
        ],
    )
    add_code_block(
        doc,
        """n, k = map(int, input().split())
a = list(map(int, input().split()))

dem = {}
ans = 0

for x in a:
    ans += dem.get(k - x, 0)
    dem[x] = dem.get(x, 0) + 1

print(ans)""",
    )

    # Ví dụ 2
    add_sub_heading(
        doc, "Ví dụ 2: Gom nhóm các từ đảo chữ (Group Anagrams)"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cho N xâu ký tự (1 ≤ N ≤ 10^5, tổng độ dài các xâu ≤ 10^6) chỉ gồm các chữ cái tiếng Anh in thường `'a'..'z'`. Hãy gom các xâu là đảo chữ của nhau vào cùng một nhóm. In ra số lượng nhóm và liệt kê từng nhóm trên một dòng.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Hai xâu là đảo chữ khi và chỉ khi vector tần suất 26 chữ cái của chúng giống hệt nhau.",
            "Chuyển danh sách tần suất 26 phần tử thành `tuple(cnt)` (mất O(|S|), nhanh hơn sắp xếp xâu O(|S| log |S|)) để làm khóa cho `defaultdict(list)`.",
        ],
    )
    add_code_block(
        doc,
        """from collections import defaultdict

n = int(input())
nhom = defaultdict(list)

for _ in range(n):
    s = input().strip()
    cnt = [0] * 26
    for ch in s:
        cnt[ord(ch) - 97] += 1
    nhom[tuple(cnt)].append(s)

print(len(nhom))
for ds in nhom.values():
    print(" ".join(ds))""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1: Đếm giá trị phân biệt và Phần tử xuất hiện nhiều nhất",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^18). Dòng đầu tiên in ra số lượng giá trị phân biệt trong dãy. Dòng thứ hai in ra giá trị xuất hiện nhiều lần nhất cùng số lần xuất hiện của nó (nếu có nhiều giá trị cùng tần suất lớn nhất thì chọn giá trị nhỏ nhất).",
            "analysis": "Dùng `Counter(a)` để đếm tần suất trong O(N). Số lượng giá trị phân biệt là `len(dem)`. Dùng hàm `min` với khóa `(-dem[x], x)` để ưu tiên tần suất lớn nhất trước, giá trị nhỏ nhất sau.",
            "example": "Ví dụ: Input: `5` và dãy `4 2 4 2 5` → Output: dòng 1 là `3`, dòng 2 là `2 2`.",
            "code": """from collections import Counter

n = int(input())
a = list(map(int, input().split()))

dem = Counter(a)
so_phan_biet = len(dem)
best_val = min(dem.keys(), key=lambda x: (-dem[x], x))

print(so_phan_biet)
print(best_val, dem[best_val])""",
        },
        {
            "title": "Bài 2: Tìm phần tử đa số (Majority Element)",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^18). Một phần tử được gọi là phần tử đa số nếu nó xuất hiện nhiều hơn ⌊N / 2⌋ lần trong dãy. Hãy tìm phần tử đa số của dãy; nếu không tồn tại in ra `KHONG CO`.",
            "analysis": "Dùng `Counter(a)` đếm tần suất trong O(N), sau đó duyệt qua từ điển kiểm tra điều kiện `freq > n // 2`.",
            "example": "Ví dụ: Input: `7` và dãy `3 1 3 3 2 3 4` → Output: `3` (xuất hiện 4 > 3 lần).",
            "code": """from collections import Counter

n = int(input())
a = list(map(int, input().split()))

dem = Counter(a)
nguong = n // 2

for val, freq in dem.items():
    if freq > nguong:
        print(val)
        break
else:
    print("KHONG CO")""",
        },
        {
            "title": "Bài 3: Phép toán giao và hiệu trên hai tập dữ liệu lớn",
            "desc": "Cho hai dãy số nguyên A gồm N phần tử và B gồm M phần tử (1 ≤ N, M ≤ 10^6; trị tuyệt đối ≤ 10^9). Dòng 1 in ra theo thứ tự tăng dần các giá trị phân biệt xuất hiện ở cả hai dãy A và B (nếu không có in `-1`). Dòng 2 in ra theo thứ tự tăng dần các giá trị phân biệt chỉ có ở dãy A mà không có ở dãy B (nếu không có in `-1`).",
            "analysis": "Chuyển hai dãy thành `set_a` và `set_b`, sau đó dùng trực tiếp toán tử giao `&` và hiệu `-` trên tập hợp trong O(N + M).",
            "example": "Ví dụ: Input: `5 4`, dãy `1 2 2 3 5` và `2 5 5 8` → Output: dòng 1 là `2 5`, dòng 2 là `1 3`.",
            "code": """n, m = map(int, input().split())
set_a = set(map(int, input().split()))
set_b = set(map(int, input().split()))

giao = sorted(set_a & set_b)
hieu = sorted(set_a - set_b)

print(*(giao if giao else [-1]))
print(*(hieu if hieu else [-1]))""",
        },
        {
            "title": "Bài 4: Đếm số cặp phần tử có chênh lệch bằng K",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (2 ≤ N ≤ 10^6; |a_i| ≤ 10^12) và một số nguyên không âm K (0 ≤ K ≤ 10^12). Hãy đếm xem có bao nhiêu cặp chỉ số (i, j) thỏa mãn 1 ≤ i < j ≤ N và |a_i - a_j| = K.",
            "analysis": "Dùng `Counter(a)` đếm tần suất trong O(N). Với trường hợp biên K = 0, mỗi giá trị xuất hiện c lần tạo ra c*(c - 1)//2 cặp; với K > 0, mỗi giá trị phân biệt v đóng góp `dem[v] * dem[v + K]` cặp.",
            "example": "Ví dụ: Input: `5 2` và dãy `1 5 3 4 2` → Output: `3` (các cặp (1, 3), (3, 5), (2, 4)).",
            "code": """from collections import Counter

n, k = map(int, input().split())
a = list(map(int, input().split()))

dem = Counter(a)
ans = 0

if k == 0:
    for c in dem.values():
        ans += c * (c - 1) // 2
else:
    for v, c in dem.items():
        if (v + k) in dem:
            ans += c * dem[v + k]

print(ans)""",
        },
        {
            "title": "Bài 5: Đoạn con liên tiếp dài nhất có tổng bằng K",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^9) và một số nguyên K (|K| ≤ 10^15). Hãy tìm độ dài lớn nhất của một đoạn con gồm các phần tử đứng liên tiếp có tổng đúng bằng K (nếu không tồn tại in ra `0`).",
            "analysis": "Gọi P_i là tổng tiền tố đến i (P_0 = 0). Đoạn (j + 1 .. i) có tổng bằng K khi P_j = P_i - K. Dùng từ điển `first_pos` chỉ lưu vị trí xuất hiện đầu tiên của mỗi tổng tiền tố để đạt độ dài lớn nhất trong O(N).",
            "example": "Ví dụ: Input: `6 15` và dãy `10 5 2 7 1 9` → Output: `4` (đoạn `5 2 7 1`).",
            "code": """n, k = map(int, input().split())
a = list(map(int, input().split()))

first_pos = {0: -1}
prefix_sum = 0
max_len = 0

for i, x in enumerate(a):
    prefix_sum += x
    if (prefix_sum - k) in first_pos:
        length = i - first_pos[prefix_sum - k]
        if length > max_len:
            max_len = length
    if prefix_sum not in first_pos:
        first_pos[prefix_sum] = i

print(max_len)""",
        },
        {
            "title": "Bài 6: Đếm số đoạn con liên tiếp có tổng chia hết cho K",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^9) và một số nguyên dương K (1 ≤ K ≤ 10^9). Hãy đếm xem có tất cả bao nhiêu đoạn con gồm các phần tử đứng liên tiếp có tổng chia hết cho K.",
            "analysis": "Đoạn con (j + 1 .. i) có tổng chia hết cho K khi và chỉ khi P_i mod K == P_j mod K. Dùng `defaultdict(int)` để đếm số lượng tổng tiền tố đã gặp có cùng số dư khi chia cho K trong O(N).",
            "example": "Ví dụ: Input: `6 5` và dãy `4 5 0 -2 -3 1` → Output: `7`.",
            "code": """from collections import defaultdict

n, k = map(int, input().split())
a = list(map(int, input().split()))

dem_du = defaultdict(int)
dem_du[0] = 1
prefix_sum = 0
ans = 0

for x in a:
    prefix_sum += x
    rem = prefix_sum % k
    ans += dem_du[rem]
    dem_du[rem] += 1

print(ans)""",
        },
        {
            "title": "Bài 7: Đếm số cặp điểm có trung điểm tọa độ nguyên",
            "desc": "Trên mặt phẳng tọa độ Oxy, cho N điểm phân biệt có tọa độ nguyên (x_i, y_i) (2 ≤ N ≤ 10^6; |x_i|, |y_i| ≤ 10^9). Hãy đếm xem có bao nhiêu cặp điểm (i, j) (1 ≤ i < j ≤ N) sao cho trung điểm của đoạn thẳng nối hai điểm đó cũng có tọa độ nguyên.",
            "analysis": "Trung điểm có tọa độ nguyên khi x_i, x_j cùng tính chẵn lẻ và y_i, y_j cùng tính chẵn lẻ. Ánh xạ mỗi điểm về `tuple` `(x % 2, y % 2)` và đếm bằng `Counter`; mỗi nhóm kích thước c đóng góp c*(c - 1)//2 cặp.",
            "example": "Ví dụ: Input: `4` và các điểm `(1, 2), (3, 4), (2, 1), (5, 6)` → Output: `3`.",
            "code": """from collections import Counter

n = int(input())
dem = Counter()

for _ in range(n):
    x, y = map(int, input().split())
    dem[(x % 2, y % 2)] += 1

ans = sum(c * (c - 1) // 2 for c in dem.values())
print(ans)""",
        },
        {
            "title": "Bài 8: Giá trị lớn nhất trên mọi cửa sổ trượt độ dài K",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ K ≤ N ≤ 10^6; |a_i| ≤ 10^9). Với mỗi đoạn con liên tiếp độ dài K từ trái sang phải, hãy tìm giá trị lớn nhất trong đoạn con đó.",
            "analysis": "Dùng `collections.deque` lưu chỉ số các phần tử theo thứ tự giá trị giảm dần. Loại bỏ chỉ số hết hạn ở đầu (`dq.popleft()`) và các phần tử nhỏ hơn hoặc bằng a[i] ở cuối (`dq.pop()`) để đạt độ phức tạp O(N).",
            "example": "Ví dụ: Input: `8 3` và dãy `1 3 -1 -3 5 3 6 7` → Output: `3 3 5 5 6 7`.",
            "code": """from collections import deque

n, k = map(int, input().split())
a = list(map(int, input().split()))

dq = deque()
res = []

for i in range(n):
    if dq and dq[0] <= i - k:
        dq.popleft()
    while dq and a[dq[-1]] <= a[i]:
        dq.pop()
    dq.append(i)
    if i >= k - 1:
        res.append(str(a[dq[0]]))

print(" ".join(res))""",
        },
        {
            "title": "Bài 9: Tuyển chọn đội hình tối ưu không xung khắc",
            "desc": "Một câu lạc bộ có N thành viên (1 ≤ N ≤ 18), thành viên thứ i có chỉ số năng lực w_i (1 ≤ w_i ≤ 10^6). Có M cặp thành viên (u, v) xung khắc nên không thể cùng được chọn vào đội. Hãy chọn ra đúng K thành viên (1 ≤ K ≤ min(N, 5)) đôi một không xung khắc sao cho tổng chỉ số năng lực là lớn nhất (nếu không chọn được in `-1`).",
            "analysis": "Lưu các cặp xung khắc dạng `tuple` chuẩn hóa `(min, max)` vào `set` để tra cứu O(1), kết hợp `itertools.combinations` duyệt toàn bộ C(N, K) đội hình.",
            "example": "Ví dụ: Input: `4 2 2`, năng lực `10 20 30 40`, xung khắc `(3, 4)` và `(2, 4)` → Output: `50`.",
            "code": """from itertools import combinations

n, m, k = map(int, input().split())
w = list(map(int, input().split()))

xung_khac = set()
for _ in range(m):
    u, v = map(int, input().split())
    u -= 1
    v -= 1
    if u > v:
        u, v = v, u
    xung_khac.add((u, v))

max_diem = -1
for nhom in combinations(range(n), k):
    hop_le = True
    for u, v in combinations(nhom, 2):
        if (u, v) in xung_khac:
            hop_le = False
            break
    if hop_le:
        tong = sum(w[i] for i in nhom)
        if tong > max_diem:
            max_diem = tong

print(max_diem)""",
        },
        {
            "title": "Bài 10: Hành trình giao hàng ngắn nhất (TSP quy mô nhỏ)",
            "desc": "Một nhân viên giao hàng xuất phát từ kho tại địa điểm 1, cần đi giao hàng qua tất cả N - 1 địa điểm còn lại (từ 2 đến N, mỗi nơi đúng 1 lần) rồi quay trở về kho 1 (2 ≤ N ≤ 10). Cho ma trận chi phí C kích thước N × N (1 ≤ C_{i, j} ≤ 10^6). Hãy tìm tổng chi phí nhỏ nhất cho toàn bộ hành trình.",
            "analysis": "Cố định xuất phát tại đỉnh 0 (đỉnh 1), dùng `itertools.permutations(range(1, n))` để duyệt toàn bộ (N - 1)! hoán vị thứ tự ghé thăm các đỉnh còn lại.",
            "example": "Ví dụ: Input: N = 4 và ma trận đối xứng chi phí 10, 15, 20, 35, 25, 30 → Output: `80`.",
            "code": """from itertools import permutations

n = int(input())
c = [list(map(int, input().split())) for _ in range(n)]

min_cost = float("inf")

# Cố định xuất phát tại đỉnh 0 (đỉnh 1), hoán vị N - 1 đỉnh còn lại
for lo_trinh in permutations(range(1, n)):
    cost = c[0][lo_trinh[0]]
    for i in range(len(lo_trinh) - 1):
        cost += c[lo_trinh[i]][lo_trinh[i + 1]]
    cost += c[lo_trinh[-1]][0]
    if cost < min_cost:
        min_cost = cost

print(min_cost)""",
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
    OUTPUT_FILE = "CAU_TRUC_DU_LIEU_PYTHON.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")