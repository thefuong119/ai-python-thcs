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
            "Bồi dưỡng HSG Tin học (Python) | Chuyên đề 8: Bộ đề thi thử cấp Tỉnh tổng hợp"
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
    """Tiêu đề mục con hoặc tên từng bài thi."""
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
    """Tạo khung nhấn mạnh cho phần Hướng dẫn chấm, Thang điểm & Bộ test biên."""
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
# NỘI DUNG CHUYÊN ĐỀ 8: BỘ ĐỀ THI THỬ HSG CẤP TỈNH TỔNG HỢP
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # --- TIÊU ĐỀ CHÍNH CỦA GIÁO ÁN ---
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "CHUYÊN ĐỀ 8: BỘ ĐỀ THI THỬ HỌC SINH GIỎI CẤP TỈNH TỔNG HỢP (PYTHON)"
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
        "Kèm Hướng dẫn chấm, Thang điểm từng Subtask, Code mẫu chuẩn & Bộ test biên"
    )
    srun.font.name = "Times New Roman"
    srun.font.size = Pt(12)
    srun.italic = True
    srun.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # ==========================================================================
    # PHẦN I: CẤU TRÚC ĐỀ THI & CHIẾN THUẬT PHÒNG THI HSG CẤP TỈNH
    # ==========================================================================
    add_part_heading(
        doc, "PHẦN I. CẤU TRÚC ĐỀ THI & CHIẾN THUẬT PHÒNG THI HSG CẤP TỈNH"
    )
    add_rich_paragraph(
        doc,
        "Một đề thi Học sinh giỏi Tin học cấp Tỉnh tiêu chuẩn thường gồm 4 bài toán có độ khó phân tầng rõ rệt qua các Subtask (tổng 20.0 điểm, thời gian làm bài 150 phút). Để tối đa hóa điểm số bằng ngôn ngữ Python, thí sinh cần tuân thủ kỷ luật phân bổ thời gian:",
    )
    add_rich_paragraph(
        doc,
        "**Bài 1 (6.0 điểm – Mục tiêu 20-25 phút):** Số học cơ bản, Xử lý xâu, Cấu trúc rẽ nhánh & Vòng lặp tối ưu. Bắt buộc ăn trọn 6.0/6.0 điểm và kiểm tra kỹ trường hợp biên số 0, số âm, số chính phương.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bài 2 (5.0 điểm – Mục tiêu 30-35 phút):** Mảng đếm (`dict`/`Counter`), Tổng tiền tố, Hai con trỏ, Sweep-line. Đạt trọn 5.0/5.0 điểm với `sys.stdin.buffer.read().split()`.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bài 3 (5.0 điểm – Mục tiêu 40-45 phút):** Tìm kiếm nhị phân trên kết quả, Tham lam nâng cao (`heapq`). Nếu chưa nghĩ ra thuật toán O(N log N), hãy cài đặt ngay thuật toán trâu ăn 40% điểm Subtask 1 trước khi tối ưu.",
        is_bullet=True,
    )
    add_rich_paragraph(
        doc,
        "**Bài 4 (4.0 điểm – Mục tiêu 45-50 phút):** Quy hoạch động (DP), Lý thuyết đồ thị (BFS, 0-1 BFS, Dijkstra phân tầng). Ăn chắc 2.0 - 3.0 điểm Subtask nhỏ rồi nâng cấp lên thuật toán tối ưu.",
        is_bullet=True,
    )

    add_callout_box(
        doc,
        "Quy trình 5 bước kiểm tra bài trước khi nộp lên hệ thống chấm:",
        [
            "Kiểm tra đúng tên tệp nguồn/dữ liệu theo quy định và tuyệt đối không in thừa chữ giải thích.",
            "Chạy kiểm thử với bộ test biên tự tạo: N = 1, N = 2, số 0, số âm, các phần tử bằng nhau toàn bộ, đồ thị không liên thông.",
            "Chạy Stress Test với dữ liệu lớn nhất (N = 2·10^5 hoặc 10^6) để đảm bảo thời gian chạy dưới 0.6 giây.",
            "Đóng gói toàn bộ mã xử lý trong `def solve():` và sử dụng `sys.stdin.buffer.read().split()`.",
            "Lưu dự phòng bản code ăn điểm Subtask nhỏ trước khi sửa sang thuật toán tối ưu.",
        ],
    )

    # ==========================================================================
    # PHẦN II: ĐỀ THI THỬ SỐ 01 (CHUẨN CẤU TRÚC THI HSG CẤP TỈNH)
    # ==========================================================================
    add_part_heading(
        doc, "PHẦN II. ĐỀ THI THỬ SỐ 01 (THỜI GIAN: 150 PHÚT – TỔNG 20.0 ĐIỂM)"
    )

    de_1 = [
        {
            "title": "BÀI 1 (6.0 điểm): SỐ ĐẶC BIỆT TRÊN ĐOẠN (SPECIAL.PY)",
            "desc": "Một số nguyên dương X được gọi là số đặc biệt nếu X = p², trong đó p là một số nguyên tố lẻ. Cho hai số nguyên dương L và R (1 ≤ L ≤ R ≤ 10^12). Hãy đếm xem trong đoạn [L, R] có bao nhiêu số đặc biệt.",
            "io": "Input: Một dòng duy nhất chứa hai số nguyên dương L, R. | Output: Số lượng số đặc biệt trong đoạn [L, R]. | Ví dụ: Input `4 30` → Output `2` (gồm 9 = 3² và 25 = 5², loại 4 = 2²).",
            "guide": [
                "Subtask 1 (3.0 điểm): 1 ≤ L ≤ R ≤ 10^6. Duyệt p từ ⌈√L⌉ đến ⌊√R⌋ và kiểm tra nguyên tố lẻ trong O(√p).",
                "Subtask 2 (3.0 điểm): 1 ≤ L ≤ R ≤ 10^12. Dùng `math.isqrt()` tìm cận dưới `low_p = max(3, math.isqrt(L - 1) + 1)` và cận trên `high_p = math.isqrt(R)`, kết hợp Sàng Eratosthenes tối ưu `bytearray` đến 10^6.",
                "Bộ test biên: (1) `1 8` → `0` (loại 4 = 2²); (2) `9 9` → `1` (trùng biên chính phương lẻ); (3) `999998000000 1000000000000` → `0` (ngưỡng 10^12).",
            ],
            "code": """import math
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    L = int(data[0])
    R = int(data[1])

    low_p = max(3, math.isqrt(L - 1) + 1)
    high_p = math.isqrt(R)

    if low_p > high_p:
        sys.stdout.write("0\\n")
        return

    is_prime = bytearray(b"\\x01") * (high_p + 1)
    is_prime[0] = is_prime[1] = 0
    for i in range(2, math.isqrt(high_p) + 1):
        if is_prime[i]:
            is_prime[i * i : high_p + 1 : i] = bytearray(
                (high_p - i * i) // i + 1
            )

    ans = sum(is_prime[low_p : high_p + 1])
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "BÀI 2 (5.0 điểm): MẬT MÃ TẦN SUẤT (FREQSTR.PY)",
            "desc": "Cho một xâu ký tự S độ dài N chỉ gồm các chữ cái tiếng Anh in thường ('a'..'z') và một số nguyên dương K (1 ≤ K ≤ N ≤ 10^6). Một xâu con liên tiếp của S được gọi là cân bằng cấp K nếu trong xâu con đó, không có bất kỳ chữ cái nào xuất hiện quá K lần. Hãy tìm độ dài lớn nhất của một xâu con liên tiếp cân bằng cấp K trong S.",
            "io": "Input: Dòng 1 chứa N, K; Dòng 2 chứa xâu S độ dài N. | Output: Độ dài lớn nhất tìm được. | Ví dụ: Input `8 2` và `aaabbbcc` → Output `4` (đoạn `aabb` hoặc `bbcc`).",
            "guide": [
                "Subtask 1 (2.5 điểm): 1 ≤ N ≤ 2000. Duyệt hai vòng lặp kiểm tra mọi đoạn con [l, r] trong O(N²).",
                "Subtask 2 (2.5 điểm): 1 ≤ N ≤ 10^6. Sử dụng kỹ thuật Cửa sổ trượt `[left, right]` kết hợp mảng đếm ASCII 128 phần tử `cnt` trong O(N).",
                "Bộ test biên: (1) `5 1` và `abcde` → `5`; (2) `6 1` và `aaaaaa` → `1`; (3) `1 1` và `z` → `1`.",
            ],
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    k = int(data[1])
    s = data[2]

    cnt = [0] * 128
    left = 0
    max_len = 0

    for right in range(n):
        ch = s[right]
        cnt[ch] += 1
        while cnt[ch] > k:
            cnt[s[left]] -= 1
            left += 1
        cur_len = right - left + 1
        if cur_len > max_len:
            max_len = cur_len

    sys.stdout.write(f"{max_len}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "BÀI 3 (5.0 điểm): PHÂN PHỐI KHO HÀNG (WAREHOUSE.PY)",
            "desc": "Có N kiện hàng xếp thành một hàng dọc, kiện thứ i có khối lượng a_i (1 ≤ a_i ≤ 10^9). Điều động K xe tải giống nhau (1 ≤ K ≤ N ≤ 2·10^5) để chở hết N kiện hàng sao cho mỗi xe chở một dãy kiện hàng đứng liên tiếp nhau. Hãy tìm tải trọng tối thiểu C của mỗi xe tải để hoàn thành công việc.",
            "io": "Input: Dòng 1 chứa N, K; Dòng 2 chứa N số nguyên a_1..a_N. | Output: Tải trọng tối thiểu C. | Ví dụ: Input `5 3` và `2 4 7 3 5` → Output `8` (chia thành `[2,4]`, `[7]`, `[3,5]`).",
            "guide": [
                "Subtask 1 (2.0 điểm): 1 ≤ K ≤ N ≤ 100. Quy hoạch động O(K · N²).",
                "Subtask 2 (3.0 điểm): 1 ≤ K ≤ N ≤ 2·10^5. Tìm kiếm nhị phân kết quả C trong đoạn `[max(a), sum(a)]` kết hợp hàm kiểm tra tham lam `can_ship(mid)` trong O(N log(sum(a))).",
                "Bộ test biên: (1) K = N (`5 5`, `3 1 9 2 4`) → `9` (bằng max(a)); (2) K = 1 (`4 1`, `10 20 30 40`) → `100` (bằng sum(a)); (3) 5 phần tử đều bằng 10^9 với K = 2 → `3000000000`.",
            ],
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    k = int(data[1])
    a = list(map(int, data[2 : 2 + n]))

    def can_ship(limit):
        trucks = 1
        cur_load = 0
        for w in a:
            if cur_load + w > limit:
                trucks += 1
                cur_load = w
                if trucks > k:
                    return False
            else:
                cur_load += w
        return True

    low = max(a)
    high = sum(a)
    ans = high

    while low <= high:
        mid = (low + high) // 2
        if can_ship(mid):
            ans = mid
            high = mid - 1
        else:
            low = mid + 1

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "BÀI 4 (4.0 điểm): HÀNH TRÌNH GIAO THƯƠNG TIẾT KIỆM (TRADE.PY)",
            "desc": "Cho đồ thị có hướng gồm N thành phố (2 ≤ N ≤ 10^5) và M tuyến đường (1 ≤ M ≤ 2·10^5), tuyến thứ i đi từ u_i đến v_i có cước phí w_i (1 ≤ w_i ≤ 10^9). Xuất phát từ thành phố 1 đến thành phố N, bạn được quyền chọn tối đa một tuyến đường trên hành trình để miễn phí hoàn toàn (w_i = 0). Tìm tổng cước phí nhỏ nhất (in `-1` nếu không có đường đi).",
            "io": "Input: Dòng 1 chứa N, M; M dòng tiếp theo mỗi dòng chứa u_i, v_i, w_i. | Output: Tổng cước phí nhỏ nhất hoặc `-1`. | Ví dụ: Input `4 4` và các cung `(1,2,5), (2,4,10), (1,3,2), (3,4,12)` → Output `2`.",
            "guide": [
                "Subtask 1 (2.0 điểm): N ≤ 1000, M ≤ 2000. Chạy Dijkstra xuôi từ 1 và Dijkstra ngược từ N, sau đó thử miễn phí từng cạnh.",
                "Subtask 2 (2.0 điểm): N ≤ 10^5, M ≤ 2·10^5. Dijkstra phân tầng với hai trạng thái `dist0[u]` (chưa dùng thẻ miễn phí) và `dist1[u]` (đã dùng thẻ miễn phí) trong O((N + M) log N).",
                "Bộ test biên: (1) Không có đường đi từ 1 đến N → `-1`; (2) Đường đi chỉ gồm 1 cung duy nhất `(1, N, 10^9)` → `0`; (3) Có đa cung giữa 2 đỉnh.",
            ],
            "code": """import heapq
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

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
                if d < dist1[v]:
                    dist1[v] = d
                    heapq.heappush(pq, (d, v, 1))
        else:
            if d > dist1[u]:
                continue
            for v, w in adj[u]:
                if d + w < dist1[v]:
                    dist1[v] = d + w
                    heapq.heappush(pq, (d + w, v, 1))

    ans = dist0[n] if dist0[n] < dist1[n] else dist1[n]
    sys.stdout.write(f"{ans if ans < INF else -1}\\n")


if __name__ == "__main__":
    solve()""",
        },
    ]

    for b in de_1:
        add_sub_heading(doc, b["title"])
        add_rich_paragraph(doc, f"**Đề bài:** {b['desc']}")
        add_rich_paragraph(doc, f"**Quy cách & Ví dụ:** {b['io']}", is_bullet=True)
        add_callout_box(doc, "Hướng dẫn chấm, Thang điểm & Bộ test biên:", b["guide"])
        add_code_block(doc, b["code"])

    # ==========================================================================
    # PHẦN III: ĐỀ THI THỬ SỐ 02 (BỘ ĐỀ TỰ LUYỆN NÂNG CAO)
    # ==========================================================================
    add_part_heading(
        doc, "PHẦN III. ĐỀ THI THỬ SỐ 02 (BỘ ĐỀ TỰ LUYỆN NÂNG CAO – 20.0 ĐIỂM)"
    )

    de_2 = [
        {
            "title": "BÀI 1 (6.0 điểm): CẶP SỐ ĐỒNG DƯ (MODPAIR.PY)",
            "desc": "Cho dãy gồm N số nguyên a_1, a_2, ..., a_N (2 ≤ N ≤ 10^6; |a_i| ≤ 10^12) và một số nguyên dương K (1 ≤ K ≤ 10^6). Hãy đếm xem có tất cả bao nhiêu cặp chỉ số (i, j) (1 ≤ i < j ≤ N) sao cho hiệu (a_i - a_j) chia hết cho K.",
            "io": "Input: Dòng 1 chứa N, K; Dòng 2 chứa N số nguyên a_1..a_N. | Output: Số lượng cặp tìm được.",
            "guide": [
                "Subtask 1 (3.0 điểm): N ≤ 2000. Duyệt 2 vòng lặp O(N²).",
                "Subtask 2 (3.0 điểm): N ≤ 10^6. Nhận xét (a_i - a_j) mod K == 0 ⇔ a_i mod K == a_j mod K. Dùng mảng đếm `cnt = [0] * K` trong O(N).",
                "Bộ test biên: (1) Số âm `[-5, 5, 0]` với K = 5 → `3`; (2) K = 1, N = 5 → `10` (mọi cặp đều chia hết cho 1).",
            ],
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    k = int(data[1])

    cnt = [0] * k
    ans = 0
    for i in range(2, 2 + n):
        rem = int(data[i]) % k
        ans += cnt[rem]
        cnt[rem] += 1

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "BÀI 2 (5.0 điểm): TRẠM THU PHÍ CAO TỐC (TOLL.PY)",
            "desc": "Ghi nhận dữ liệu của N xe lưu thông trên cao tốc (1 ≤ N ≤ 2·10^5): xe thứ i đi vào lúc s_i và rời khỏi cao tốc lúc e_i (0 ≤ s_i < e_i ≤ 10^9). Tìm số lượng xe lớn nhất có mặt đồng thời trên tuyến cao tốc tại cùng một thời điểm (nếu xe A ra đúng lúc xe B vào tại cùng thời điểm t thì không tính là đồng thời).",
            "io": "Input: Dòng 1 chứa N; N dòng tiếp theo mỗi dòng chứa s_i, e_i. | Output: Số lượng xe đồng thời lớn nhất.",
            "guide": [
                "Subtask 1 (2.0 điểm): N ≤ 2000. Kiểm tra tại từng mốc thời điểm vào s_i.",
                "Subtask 2 (3.0 điểm): N ≤ 2·10^5. Kỹ thuật Quét sự kiện (Sweep-line): tạo cặp `(s_i, 1)` và `(e_i, -1)`, sắp xếp tăng dần (sự kiện `-1` tự đứng trước `1` khi cùng thời điểm) trong O(N log N).",
                "Bộ test biên: (1) Các xe nối đuôi `[1, 5], [5, 10], [10, 15]` → `1`; (2) Tất cả xe trùng giờ `[2, 8], [2, 8], [2, 8]` → `3`.",
            ],
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    events = []
    idx = 1
    for _ in range(n):
        s = int(data[idx])
        e = int(data[idx + 1])
        idx += 2
        events.append((s, 1))
        events.append((e, -1))

    events.sort()

    cur_cars = 0
    max_cars = 0
    for _, delta in events:
        cur_cars += delta
        if cur_cars > max_cars:
            max_cars = cur_cars

    sys.stdout.write(f"{max_cars}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "BÀI 3 (5.0 điểm): THU HOẠCH NÔNG SẢN (HARVEST.PY)",
            "desc": "Trên một cánh đồng gồm N thửa ruộng đánh số từ 1 đến N (2 ≤ N ≤ 2·10^5), thửa thứ i mang lại lợi nhuận ròng là a_i (|a_i| ≤ 10^9, có thể âm). Hãy chọn ra hai dãy thửa ruộng liên tiếp không giao nhau (mỗi dãy có ít nhất 1 thửa ruộng) sao cho tổng lợi nhuận của cả hai dãy là lớn nhất.",
            "io": "Input: Dòng 1 chứa N; Dòng 2 chứa N số nguyên a_1..a_N. | Output: Tổng lợi nhuận lớn nhất.",
            "guide": [
                "Subtask 1 (2.0 điểm): N ≤ 300. | Subtask 2 (1.5 điểm): N ≤ 5000. | Subtask 3 (1.5 điểm): N ≤ 2·10^5.",
                "Thuật toán tối ưu O(N): Dùng thuật toán Kadane hai chiều tính `pref[i]` (tổng đoạn con lớn nhất trong `[0..i]`) và `suff[i]` (tổng đoạn con lớn nhất trong `[i..N-1]`), sau đó lấy `max(pref[i] + suff[i + 1])`.",
                "Bộ test biên: (1) Tất cả phần tử âm `[-5, -2, -8, -1]` → `-3` (chọn `[-2]` và `[-1]`); (2) N = 2 với `[10, -20]` → `-10`.",
            ],
            "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    n = int(data[0])
    a = list(map(int, data[1 : 1 + n]))

    pref = [0] * n
    cur_max = a[0]
    best = a[0]
    pref[0] = best
    for i in range(1, n):
        cur_max = a[i] if a[i] > cur_max + a[i] else cur_max + a[i]
        if cur_max > best:
            best = cur_max
        pref[i] = best

    suff = [0] * n
    cur_max = a[n - 1]
    best = a[n - 1]
    suff[n - 1] = best
    for i in range(n - 2, -1, -1):
        cur_max = a[i] if a[i] > cur_max + a[i] else cur_max + a[i]
        if cur_max > best:
            best = cur_max
        suff[i] = best

    ans = pref[0] + suff[1]
    for i in range(1, n - 1):
        val = pref[i] + suff[i + 1]
        if val > ans:
            ans = val

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
        },
        {
            "title": "BÀI 4 (4.0 điểm): XÂY DỰNG ĐƯỜNG HẦM (TUNNEL.PY)",
            "desc": "Cho bản đồ địa hình lưới M × N (1 ≤ M, N ≤ 1000) gồm '.' (đồng bằng, đi qua mất 0 chi phí đào hầm) và '#' (núi đá, đi qua mất 1 chi phí đào hầm). Xuất phát từ ô (1, 1) đến ô (M, N) (tính cả chi phí tại ô đầu và ô cuối nếu là '#'), mỗi bước đi sang 4 ô chung cạnh. Tìm tổng chi phí đào hầm ít nhất.",
            "io": "Input: Dòng 1 chứa M, N; M dòng tiếp theo là lưới M × N. | Output: Chi phí đào hầm ít nhất.",
            "guide": [
                "Subtask 1 (2.0 điểm): M, N ≤ 50. | Subtask 2 (2.0 điểm): M, N ≤ 1000.",
                "Thuật toán tối ưu O(M · N): Dùng 0-1 BFS với `collections.deque`: ô '.' (chi phí 0) đẩy vào đầu hàng đợi `dq.appendleft()`, ô '#' (chi phí 1) đẩy vào cuối hàng đợi `dq.append()`.",
                "Bộ test biên: (1) Ô (1, 1) hoặc (M, N) là '#'; (2) Lưới 1×1 chứa '#' → `1`.",
            ],
            "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return

    m = int(data[0])
    n = int(data[1])
    grid = data[2 : 2 + m]

    INF = 10**9
    dist = [[INF] * n for _ in range(m)]
    start_cost = 1 if grid[0][0] == 35 else 0  # 35 là mã ASCII của '#'
    dist[0][0] = start_cost

    dq = deque([(0, 0)])
    dirs = [(-1, 0), (1, 0), (0, -1), (0, 1)]

    while dq:
        r, c = dq.popleft()
        d = dist[r][c]
        if r == m - 1 and c == n - 1:
            break
        for dr, dc in dirs:
            nr, nc = r + dr, c + dc
            if 0 <= nr < m and 0 <= nc < n:
                w = 1 if grid[nr][nc] == 35 else 0
                if d + w < dist[nr][nc]:
                    dist[nr][nc] = d + w
                    if w == 0:
                        dq.appendleft((nr, nc))
                    else:
                        dq.append((nr, nc))

    sys.stdout.write(f"{dist[m - 1][n - 1]}\\n")


if __name__ == "__main__":
    solve()""",
        },
    ]

    for b in de_2:
        add_sub_heading(doc, b["title"])
        add_rich_paragraph(doc, f"**Đề bài:** {b['desc']}")
        add_rich_paragraph(doc, f"**Quy cách & Ví dụ:** {b['io']}", is_bullet=True)
        add_callout_box(doc, "Hướng dẫn chấm, Thang điểm & Bộ test biên:", b["guide"])
        add_code_block(doc, b["code"])

    doc.save(output_filename)
    print(
        f"THÀNH CÔNG! Đã tạo tệp giáo án: {os.path.abspath(output_filename)}"
    )


if __name__ == "__main__":
    OUTPUT_FILE = "DE_THI_THU_HSG_TIN_HOC.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")