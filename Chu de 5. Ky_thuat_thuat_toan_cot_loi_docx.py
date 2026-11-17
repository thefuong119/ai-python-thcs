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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 5: Các kỹ thuật thuật toán cốt lõi"
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
# NỘI DUNG CHUYÊN ĐỀ 5: CÁC KỸ THUẬT THUẬT TOÁN CỐT LÕI
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 5: CÁC KỸ THUẬT THUẬT TOÁN CỐT LÕI (SẮP XẾP & TÌM KIẾM NHỊ PHÂN, THAM LAM, QUY HOẠCH ĐỘNG)"
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
        "Trong các đề thi Học sinh giỏi (HSG) Tin học cấp Tỉnh và Quốc gia, ba kỹ thuật Tìm kiếm – Sắp xếp nâng cao, Thuật toán Tham lam (Greedy) và Quy hoạch động (Dynamic Programming - DP) là những công cụ quyết định để chinh phục các bài toán phân loại thí sinh (chiếm 50% - 65% tổng số điểm đề thi).",
    )

    add_sub_heading(doc, "1. Kỹ thuật Sắp xếp nâng cao & Tìm kiếm nhị phân")
    add_rich_paragraph(
        doc,
        "**Sắp xếp tùy biến (Custom Sorting) trong O(N log N):** Python sử dụng thuật toán Timsort cực kỳ tối ưu. Có thể sắp xếp theo nhiều tiêu chí bằng `tuple`: `arr.sort(key=lambda x: (x[1], -x[0]))` hoặc định nghĩa hàm so sánh phức tạp qua `functools.cmp_to_key(cmp_func)`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Tìm kiếm nhị phân trên mảng đã sắp xếp (`bisect`):** `bisect_left(a, x, lo, hi)` trả về vị trí đầu tiên có giá trị ≥ x; `bisect_right(a, x, lo, hi)` trả về vị trí đầu tiên có giá trị > x trong O(log N).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Tìm kiếm nhị phân trên tập kết quả (Binary Search on Answer):** Áp dụng khi đề bài yêu cầu *'Tìm giá trị lớn nhất của cái nhỏ nhất'* hoặc *'nhỏ nhất của cái lớn nhất'*, và hàm kiểm tra `check(mid)` có tính đơn điệu (`False...False, True...True`). Độ phức tạp O(T_check · log(MAX_ANS)).",
        is_bullet=True,
    )

    add_sub_heading(doc, "2. Thuật toán Tham lam (Greedy)")
    add_rich_paragraph(
        doc,
        "**Tư tưởng cốt lõi:** Tại mỗi bước, chọn phương án tối ưu nhất ở thời điểm hiện tại (tối ưu cục bộ) để hướng tới tối ưu toàn cục.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Các mô hình tham lam kinh điển:** Lập lịch / Chọn hoạt động (sắp xếp theo thời điểm kết thúc tăng dần); Ghép cặp tối ưu bằng hai con trỏ; Tham lam động kết hợp Hàng đợi ưu tiên (`heapq`) để luôn trích xuất phần tử nhỏ nhất/lớn nhất trong O(log N).",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Nguyên tắc kiểm chứng:** Luôn thử tìm phản ví dụ nhỏ trước khi chốt Tham lam (ví dụ: đổi tiền với bộ mệnh giá tùy ý {1, 3, 4} đổi 6 đồng thì tham lam chọn 4+1+1 mất 3 tờ, trong khi tối ưu DP là 3+3 chỉ mất 2 tờ).",
        is_bullet=True,
    )

    add_sub_heading(
        doc, "3. Kỹ thuật Quy hoạch động (Dynamic Programming - DP)"
    )
    add_rich_paragraph(
        doc,
        "**Điều kiện áp dụng:** Bài toán có tính chất *bài toán con gối nhau* và *cấu trúc con tối ưu*.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Quy trình 4 bước thiết kế lời giải DP chuẩn HSG:** (1) Định nghĩa rõ ý nghĩa trạng thái `dp[...]`; (2) Thiết lập công thức truy hồi; (3) Khởi tạo bài toán cơ sở (Base cases); (4) Xác định thứ tự tính và tối ưu không gian bộ nhớ từ 2D xuống 1D nếu có thể.",
        is_bullet=True,
    )

    add_sub_heading(doc, '4. Bốn "cạm bẫy" kinh điển trong Python cần tránh')
    add_rich_paragraph(
        doc,
        "**Bẫy khởi tạo mảng 2 chiều:** Viết `dp = [[0] * m] * n` khiến tất cả các hàng cùng trỏ vào một vùng nhớ duy nhất. **Bắt buộc viết:** `dp = [[0] * m for _ in range(n)]`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy đệ quy sâu (`RecursionError`):** Python giới hạn mặc định 1000 tầng đệ quy. Với DP có N ≥ 1000, luôn ưu tiên cài đặt **DP Bottom-up (dùng vòng lặp `for`)** vì không bị tràn ngăn xếp và chạy nhanh gấp 3 - 5 lần.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy lặp vô tận trong Tìm kiếm nhị phân:** Nên dùng cấu trúc chuẩn `while low <= high:` kết hợp biến lưu nghiệm `ans = mid` và cập nhật `low = mid + 1` hoặc `high = mid - 1`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bẫy nhập xuất dữ liệu lớn (N ≥ 10^5):** Sử dụng `sys.stdin.read().split()` để đọc toàn bộ dữ liệu đầu vào siêu tốc trong O(N).",
        is_bullet=True,
    )

    # ==========================================================================
    # PHẦN II: VÍ DỤ MINH HỌA
    # ==========================================================================
    add_part_heading(doc, "PHẦN II. VÍ DỤ MINH HỌA")

    # Ví dụ 1
    add_sub_heading(
        doc, "Ví dụ 1 (Tìm kiếm nhị phân trên kết quả): Chặt cây lấy gỗ"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Cần khai thác ít nhất M mét gỗ (1 ≤ M ≤ 2·10^9) từ N cây (1 ≤ N ≤ 10^6), cây thứ i cao h_i (0 ≤ h_i ≤ 10^9). Máy cưa đặt ở độ cao H sẽ cắt bỏ phần ngọn cao hơn H của các cây có h_i > H. Hãy tìm độ cao H lớn nhất để tổng lượng gỗ thu được ít nhất là M mét.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Độ cao H càng lớn thì lượng gỗ thu được càng ít (tính đơn điệu).",
            "Tìm kiếm nhị phân H trong đoạn [0, max(h_i)]. Với mỗi `mid`, tính tổng gỗ thu được trong O(N); nếu tổng ≥ M thì lưu `ans = mid` và thử nâng `low = mid + 1`, ngược lại hạ `high = mid - 1`.",
        ],
    )
    add_code_block(
        doc,
        """import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])
h = list(map(int, data[2 : 2 + n]))

low, high = 0, max(h)
ans = 0

while low <= high:
    mid = (low + high) // 2
    wood = 0
    for x in h:
        if x > mid:
            wood += x - mid
    if wood >= m:
        ans = mid
        low = mid + 1
    else:
        high = mid - 1

print(ans)""",
    )

    # Ví dụ 2
    add_sub_heading(
        doc, "Ví dụ 2 (Thuật toán Tham lam): Bố trí phòng họp tối đa"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Có N cuộc họp (1 ≤ N ≤ 10^6), cuộc họp thứ i bắt đầu tại s_i và kết thúc tại e_i (0 ≤ s_i < e_i ≤ 10^9). Một phòng họp chỉ phục vụ một cuộc họp tại một thời điểm (s_j ≥ e_i). Hãy chọn ra nhiều cuộc họp nhất có thể.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Tiêu chí tham lam chuẩn: Luôn ưu tiên chọn cuộc họp kết thúc sớm nhất (`key=lambda x: x[1]`) để giải phóng phòng họp sớm nhất.",
            "Sau khi sắp xếp theo e_i tăng dần, duyệt lần lượt và chọn mọi cuộc họp có s_i ≥ last_end trong O(N log N).",
        ],
    )
    add_code_block(
        doc,
        """import sys

data = sys.stdin.read().split()
n = int(data[0])
meetings = []
idx = 1
for _ in range(n):
    s = int(data[idx])
    e = int(data[idx + 1])
    meetings.append((s, e))
    idx += 2

meetings.sort(key=lambda x: x[1])

count = 0
last_end = -1

for s, e in meetings:
    if s >= last_end:
        count += 1
        last_end = e

print(count)""",
    )

    # Ví dụ 3
    add_sub_heading(
        doc, "Ví dụ 3 (Quy hoạch động): Bài toán Cái túi 0/1 tối ưu bộ nhớ"
    )
    add_rich_paragraph(
        doc,
        "**Đề bài:** Có N đồ vật (1 ≤ N ≤ 1000), vật thứ i có trọng lượng w_i (1 ≤ w_i ≤ 10^4) và giá trị v_i (1 ≤ v_i ≤ 10^9). Một chiếc ba lô có sức chứa tối đa W (1 ≤ W ≤ 10^4). Hãy chọn tập đồ vật (mỗi vật chọn tối đa 1 lần) sao cho tổng trọng lượng ≤ W và tổng giá trị lớn nhất.",
    )
    add_callout_box(
        doc,
        "Phân tích tư duy HSG:",
        [
            "Gọi `dp[j]` là tổng giá trị lớn nhất đạt được với sức chứa j.",
            "Để đảm bảo mỗi đồ vật chỉ được chọn tối đa 1 lần, ta duyệt trọng lượng j từ lớn về nhỏ (`range(W, w - 1, -1)`): `dp[j] = max(dp[j], dp[j - w] + v)`. Bộ nhớ tối ưu chỉ O(W).",
        ],
    )
    add_code_block(
        doc,
        """import sys

data = sys.stdin.read().split()
n = int(data[0])
W = int(data[1])

dp = [0] * (W + 1)
idx = 2

for _ in range(n):
    w = int(data[idx])
    v = int(data[idx + 1])
    idx += 2
    for j in range(W, w - 1, -1):
        if dp[j - w] + v > dp[j]:
            dp[j] = dp[j - w] + v

print(dp[W])""",
    )

    # ==========================================================================
    # PHẦN III: 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)
    # ==========================================================================
    add_part_heading(doc, "PHẦN III. 10 BÀI TẬP VẬN DỤNG (CÓ CODE MẪU CHUẨN)")

    bai_tap_list = [
        {
            "title": "Bài 1 (Sắp xếp nâng cao): Ghép mảng số thành số lớn nhất",
            "desc": "Cho N số nguyên không âm a_1, a_2, ..., a_N (1 ≤ N ≤ 10^5; 0 ≤ a_i ≤ 10^9). Hãy sắp xếp và ghép tất cả N số này liền nhau để tạo thành một số nguyên lớn nhất có thể.",
            "analysis": "Dùng `functools.cmp_to_key` so sánh hai xâu số X và Y theo quy tắc đặt X trước Y nếu `X + Y > Y + X`. Lưu ý biên: nếu phần tử đầu sau khi sắp xếp là `'0'` thì in ra `'0'`.",
            "example": "Ví dụ: Input: `5` và dãy `3 30 34 5 9` → Output: `9534330`.",
            "code": """from functools import cmp_to_key
import sys

data = sys.stdin.read().split()
n = int(data[0])
nums = data[1 : 1 + n]


def compare(x, y):
    if x + y > y + x:
        return -1
    if x + y < y + x:
        return 1
    return 0


nums.sort(key=cmp_to_key(compare))

if nums[0] == "0":
    print("0")
else:
    print("".join(nums))""",
        },
        {
            "title": "Bài 2 (Tìm kiếm nhị phân `bisect`): Đếm cặp có tổng trong đoạn [L, R]",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (2 ≤ N ≤ 2·10^5; |a_i| ≤ 10^9) và hai số nguyên L, R (-2·10^9 ≤ L ≤ R ≤ 2·10^9). Hãy đếm xem có bao nhiêu cặp chỉ số (i, j) (1 ≤ i < j ≤ N) thỏa mãn L ≤ a_i + a_j ≤ R.",
            "analysis": "Sắp xếp dãy tăng dần. Với mỗi phần tử a_i, dùng `bisect_left` và `bisect_right` trên đoạn chỉ số `[i + 1, n)` để đếm số phần tử a_j nằm trong đoạn `[L - a_i, R - a_i]` trong O(log N).",
            "example": "Ví dụ: Input: `4 4 7` và dãy `5 1 2 3` → Output: `4`.",
            "code": """from bisect import bisect_left, bisect_right
import sys

data = sys.stdin.read().split()
n = int(data[0])
L = int(data[1])
R = int(data[2])
a = sorted(map(int, data[3 : 3 + n]))

ans = 0
for i in range(n - 1):
    left_idx = bisect_left(a, L - a[i], i + 1, n)
    right_idx = bisect_right(a, R - a[i], i + 1, n)
    ans += right_idx - left_idx

print(ans)""",
        },
        {
            "title": "Bài 3 (Tìm kiếm nhị phân trên kết quả): Chia mảng thành tối đa K đoạn",
            "desc": "Cho dãy gồm N số nguyên dương a_1, a_2, ..., a_N (1 ≤ K ≤ N ≤ 10^5; 1 ≤ a_i ≤ 10^9). Hãy chia dãy thành không quá K đoạn con liên tiếp sao cho tổng lớn nhất trong các đoạn con là nhỏ nhất có thể.",
            "analysis": "Tìm kiếm nhị phân ngưỡng tổng `limit` trong đoạn `[max(a), sum(a)]`. Hàm `check(limit)` duyệt tham lam từ trái sang phải đếm số đoạn cần thiết; nếu ≤ K thì giảm `high = mid - 1`.",
            "example": "Ví dụ: Input: `5 3` và dãy `2 4 7 3 5` → Output: `8`.",
            "code": """import sys

data = sys.stdin.read().split()
n = int(data[0])
k = int(data[1])
a = list(map(int, data[2 : 2 + n]))


def check(limit):
    segments = 1
    cur_sum = 0
    for x in a:
        if cur_sum + x > limit:
            segments += 1
            cur_sum = x
            if segments > k:
                return False
        else:
            cur_sum += x
    return True


low, high = max(a), sum(a)
ans = high

while low <= high:
    mid = (low + high) // 2
    if check(mid):
        ans = mid
        high = mid - 1
    else:
        low = mid + 1

print(ans)""",
        },
        {
            "title": "Bài 4 (Tìm kiếm nhị phân trên kết quả): Bố trí trạm phát sóng",
            "desc": "Dọc theo một tuyến đường thẳng có N vị trí khả thi có tọa độ nguyên x_1, x_2, ..., x_N (2 ≤ K ≤ N ≤ 10^5; 0 ≤ x_i ≤ 10^9). Cần lắp đặt K trạm phát sóng tại K vị trí khác nhau sao cho khoảng cách nhỏ nhất giữa hai trạm bất kỳ là lớn nhất có thể.",
            "analysis": "Sắp xếp tọa độ tăng dần. Tìm kiếm nhị phân khoảng cách tối thiểu `dist` trong đoạn `[1, x[-1] - x[0]]` kết hợp kiểm tra tham lam từ trái sang phải.",
            "example": "Ví dụ: Input: `5 3` và tọa độ `1 2 8 4 9` → Output: `3`.",
            "code": """import sys

data = sys.stdin.read().split()
n = int(data[0])
k = int(data[1])
x = sorted(map(int, data[2 : 2 + n]))


def can_place(dist):
    count = 1
    last_pos = x[0]
    for i in range(1, n):
        if x[i] - last_pos >= dist:
            count += 1
            last_pos = x[i]
            if count >= k:
                return True
    return False


low, high = 1, x[-1] - x[0]
ans = 0

while low <= high:
    mid = (low + high) // 2
    if can_place(mid):
        ans = mid
        low = mid + 1
    else:
        high = mid - 1

print(ans)""",
        },
        {
            "title": "Bài 5 (Tham lam + Sắp xếp): Phủ kín đoạn thẳng [0, M]",
            "desc": "Cho một số nguyên dương M (1 ≤ M ≤ 10^9) và N đoạn thẳng (1 ≤ N ≤ 10^5), đoạn thẳng thứ i bao phủ khoảng [l_i, r_i] (|l_i|, |r_i| ≤ 10^9). Hãy chọn ra ít đoạn thẳng nhất sao cho hợp của chúng phủ kín toàn bộ đoạn [0, M]. Nếu không thể phủ kín, in ra `-1`.",
            "analysis": "Sắp xếp các đoạn thẳng theo đầu mút trái tăng dần. Tại mỗi bước, trong các đoạn có `l_i ≤ cur_end`, chọn đoạn vươn xa nhất sang phải (`max_r`) trong O(N log N).",
            "example": "Ví dụ: Input: `4 10` và các đoạn `[-1, 4], [2, 7], [5, 10], [0, 3]` → Output: `3`.",
            "code": """import sys

data = sys.stdin.read().split()
n = int(data[0])
m = int(data[1])
segs = []
idx = 2
for _ in range(n):
    l = int(data[idx])
    r = int(data[idx + 1])
    idx += 2
    if r > 0 and l < m:
        segs.append((l, r))

segs.sort()

cur_end = 0
count = 0
i = 0
total = len(segs)

while cur_end < m:
    max_r = cur_end
    while i < total and segs[i][0] <= cur_end:
        if segs[i][1] > max_r:
            max_r = segs[i][1]
        i += 1
    if max_r == cur_end:
        print(-1)
        break
    cur_end = max_r
    count += 1
else:
    print(count)""",
        },
        {
            "title": "Bài 6 (Tham lam + `heapq`): Nối dây với chi phí nhỏ nhất",
            "desc": "Có N sợi dây thừng (1 ≤ N ≤ 10^5), sợi thứ i có độ dài a_i (1 ≤ a_i ≤ 10^9). Mỗi lần nối hai sợi dây có độ dài x và y thành một sợi dài x + y tốn chi phí là x + y. Hãy tìm tổng chi phí nhỏ nhất để nối tất cả N sợi dây thành một sợi duy nhất.",
            "analysis": "Áp dụng thuật toán Huffman: dùng hàng đợi ưu tiên `heapq` để luôn lấy ra hai sợi dây ngắn nhất hiện tại, cộng vào tổng chi phí rồi đẩy sợi dây mới trở lại heap trong O(N log N).",
            "example": "Ví dụ: Input: `4` và dãy `4 3 2 6` → Output: `29`.",
            "code": """import heapq
import sys

data = sys.stdin.read().split()
n = int(data[0])
a = list(map(int, data[1 : 1 + n]))

heapq.heapify(a)
total_cost = 0

while len(a) > 1:
    x = heapq.heappop(a)
    y = heapq.heappop(a)
    cost = x + y
    total_cost += cost
    heapq.heappush(a, cost)

print(total_cost)""",
        },
        {
            "title": "Bài 7 (Quy hoạch động 1 chiều): Đổi tiền ít tờ nhất",
            "desc": "Một hệ thống tiền tệ có N loại mệnh giá tiền (1 ≤ N ≤ 100), mệnh giá thứ i là c_i (1 ≤ c_i ≤ 10^5), số lượng mỗi loại không hạn chế. Cần đổi một số tiền đúng bằng S (1 ≤ S ≤ 10^5). Hãy tìm số lượng tờ tiền ít nhất cần dùng (nếu không thể đổi được đúng số tiền S, in ra `-1`).",
            "analysis": "Gọi `dp[j]` là số tờ tiền ít nhất đổi được số tiền j. Vì mỗi mệnh giá được dùng nhiều lần, duyệt j từ nhỏ đến lớn (`range(c, s + 1)`) và cập nhật `dp[j] = min(dp[j], dp[j - c] + 1)`.",
            "example": "Ví dụ: Input: `3 11` và mệnh giá `1 2 5` → Output: `3` (5 + 5 + 1 = 11).",
            "code": """import sys

data = sys.stdin.read().split()
n = int(data[0])
s = int(data[1])
coins = list(map(int, data[2 : 2 + n]))

INF = 10**9
dp = [INF] * (s + 1)
dp[0] = 0

for c in coins:
    for j in range(c, s + 1):
        if dp[j - c] + 1 < dp[j]:
            dp[j] = dp[j - c] + 1

print(dp[s] if dp[s] < INF else -1)""",
        },
        {
            "title": "Bài 8 (Quy hoạch động + `bisect`): Dãy con tăng nghiêm ngặt dài nhất",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (1 ≤ N ≤ 2·10^5; |a_i| ≤ 10^9). Hãy tìm độ dài của dãy con (không nhất thiết liên tiếp) tăng nghiêm ngặt dài nhất (a_{i_1} < a_{i_2} < ... < a_{i_k}).",
            "analysis": "Duy trì mảng `tails` lưu phần tử kết thúc nhỏ nhất của dãy con tăng theo từng độ dài, kết hợp tìm kiếm nhị phân `bisect_left(tails, x)` để đạt độ phức tạp O(N log N).",
            "example": "Ví dụ: Input: `6` và dãy `10 9 2 5 3 7` → Output: `3` (dãy con `2, 3, 7`).",
            "code": """from bisect import bisect_left
import sys

data = sys.stdin.read().split()
n = int(data[0])
a = list(map(int, data[1 : 1 + n]))

tails = []
for x in a:
    pos = bisect_left(tails, x)
    if pos == len(tails):
        tails.append(x)
    else:
        tails[pos] = x

print(len(tails))""",
        },
        {
            "title": "Bài 9 (Quy hoạch động 2 chiều trên xâu): Xâu con chung dài nhất (LCS)",
            "desc": "Cho hai xâu ký tự A và B (1 ≤ |A|, |B| ≤ 3000) chỉ gồm các chữ cái tiếng Anh in hoa. Hãy tìm độ dài của xâu con chung (không nhất thiết liên tiếp) dài nhất xuất hiện trong cả A và B.",
            "analysis": "Nén mảng quy hoạch động 2 chiều thành mảng 1 chiều `dp` kích thước `|B| + 1` và dùng biến `prev_diag` lưu giá trị chéo trên trái `dp[i - 1][j - 1]` để tiết kiệm bộ nhớ và tăng tốc độ chạy.",
            "example": 'Ví dụ: Input: `ABCBDAB` và `BDCABA` → Output: `4` (xâu `"BCBA"`).',
            "code": """a = input().strip()
b = input().strip()

m = len(b)
dp = [0] * (m + 1)

for ch_a in a:
    prev_diag = 0
    for j, ch_b in enumerate(b, 1):
        temp = dp[j]
        if ch_a == ch_b:
            dp[j] = prev_diag + 1
        elif dp[j - 1] > dp[j]:
            dp[j] = dp[j - 1]
        prev_diag = temp

print(dp[m])""",
        },
        {
            "title": "Bài 10 (Quy hoạch động trên lưới ô vuông): Đường đi tổng điểm lớn nhất",
            "desc": "Cho một bảng số gồm M hàng và N cột (1 ≤ M, N ≤ 1000), ô ở hàng i cột j chứa một số nguyên A_{i, j} (|A_{i, j}| ≤ 10^9). Xuất phát từ ô (1, 1), mỗi bước chỉ được phép đi sang phải sang ô (i, j + 1) hoặc đi xuống dưới sang ô (i + 1, j) để đến đích là ô (M, N). Hãy tìm tổng điểm lớn nhất thu được trên đường đi.",
            "analysis": "Gọi `dp[i][j]` là tổng điểm lớn nhất đi tới ô `(i, j)`. Công thức truy hồi: `dp[i][j] = max(dp[i - 1][j], dp[i][j - 1]) + grid[i][j]`, độ phức tạp O(M · N).",
            "example": "Ví dụ: Input: bảng 3×3 `[[1,3,1],[1,5,1],[4,2,1]]` → Output: `12`.",
            "code": """import sys

data = sys.stdin.read().split()
m = int(data[0])
n = int(data[1])

grid = []
idx = 2
for _ in range(m):
    grid.append(list(map(int, data[idx : idx + n])))
    idx += n

dp = [[0] * n for _ in range(m)]
dp[0][0] = grid[0][0]

for j in range(1, n):
    dp[0][j] = dp[0][j - 1] + grid[0][j]

for i in range(1, m):
    dp[i][0] = dp[i - 1][0] + grid[i][0]

for i in range(1, m):
    row_grid = grid[i]
    row_dp = dp[i]
    prev_dp = dp[i - 1]
    for j in range(1, n):
        best = prev_dp[j] if prev_dp[j] > row_dp[j - 1] else row_dp[j - 1]
        row_dp[j] = best + row_grid[j]

print(dp[m - 1][n - 1])""",
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
    OUTPUT_FILE = "KY_THUAT_THUAT_TOAN_COT_LOI.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")