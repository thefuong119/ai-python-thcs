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
    """Tô màu nền cho ô trong bảng."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_margins(cell, top=100, bottom=100, left=160, right=160):
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
    """Tạo đường viền mảnh cho khung code hoặc viền nhấn bên trái cho khung phân tích."""
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
            "Bồi dưỡng HSG Tin học (Python) | Tuyển tập 40 bài tập trọng tâm cấp Tỉnh & Quốc gia"
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
    """Tiêu đề Mảng kiến thức lớn."""
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
    """Tiêu đề từng bài tập."""
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
    set_cell_margins(cell, top=90, bottom=90, left=150, right=150)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12
    p.paragraph_format.keep_together = True

    run = p.add_run(code_text.strip("\n"))
    run.font.name = "Consolas"
    run.font.size = Pt(9.8)
    run.font.color.rgb = RGBColor(0x0F, 0x17, 0x2A)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(4)


def add_analysis_box(doc, analysis_text):
    """Tạo khung nhấn mạnh cho phần Phân tích thuật toán & Độ phức tạp."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(17.0)

    cell = table.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_background(cell, "EFF6FF")
    set_cell_borders(cell, color="2563EB", left_accent=True)
    set_cell_margins(cell, top=80, bottom=80, left=150, right=130)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12

    r_lbl = p.add_run("Phân tích thuật toán & Độ phức tạp: ")
    r_lbl.font.name = "Times New Roman"
    r_lbl.font.size = Pt(11.5)
    r_lbl.bold = True
    r_lbl.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)

    tokens = re.split(r"(\*\*.*?\*\*|`.*?`)", analysis_text)
    for token in tokens:
        if not token:
            continue
        if token.startswith("**") and token.endswith("**"):
            run = p.add_run(token[2:-2])
            run.font.name = "Times New Roman"
            run.font.size = Pt(11.5)
            run.bold = True
        elif token.startswith("`") and token.endswith("`"):
            run = p.add_run(token[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0xB9, 0x1C, 0x1C)
        else:
            run = p.add_run(token)
            run.font.name = "Times New Roman"
            run.font.size = Pt(11.5)

    sp = doc.add_paragraph()
    sp.paragraph_format.space_before = Pt(0)
    sp.paragraph_format.space_after = Pt(3)


# ==============================================================================
# DỮ LIỆU 40 BÀI TẬP TRỌNG TÂM (8 MẢNG KIẾN THỨC × 5 BÀI)
# ==============================================================================

SECTIONS_DATA = [
    (
        "MẢNG 1: SỐ HỌC & TOÁN HỌC TÍNH TOÁN (BÀI 1 – BÀI 5)",
        [
            {
                "title": "Bài 1: Số giàu ước nguyên tố trên đoạn (Sàng biến thể + Mảng cộng dồn)",
                "desc": "Cho Q truy vấn (1 ≤ Q ≤ 10^5), mỗi truy vấn gồm ba số nguyên L, R, K (1 ≤ L ≤ R ≤ 10^6; 1 ≤ K ≤ 7). Hãy đếm xem trong đoạn [L, R] có bao nhiêu số nguyên dương có đúng K ước số nguyên tố phân biệt.",
                "example": "Ví dụ: Input: `2`, `1 20 2`, `10 30 3` → Output: `7` và `1` (số 30 = 2 × 3 × 5).",
                "analysis": "Dùng sàng Eratosthenes tạo mảng `omega[x]` đếm số ước nguyên tố phân biệt của x (x ≤ 10^6), sau đó xây dựng 7 mảng cộng dồn `pref[k][i]` để trả lời mỗi truy vấn trong O(1).",
                "code": """import sys


def solve():
    MAX = 1000000
    omega = bytearray(MAX + 1)
    for i in range(2, MAX + 1):
        if omega[i] == 0:
            for j in range(i, MAX + 1, i):
                omega[j] += 1

    pref = [[0] * (MAX + 1) for _ in range(8)]
    for i in range(2, MAX + 1):
        w = omega[i]
        for k in range(1, 8):
            pref[k][i] = pref[k][i - 1] + (1 if w == k else 0)

    data = sys.stdin.buffer.read().split()
    if not data:
        return
    q = int(data[0])
    out = []
    idx = 1
    for _ in range(q):
        L = int(data[idx])
        R = int(data[idx + 1])
        K = int(data[idx + 2])
        idx += 3
        out.append(str(pref[K][R] - pref[K][L - 1]))

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 2: Tổng các ước lẻ lớn nhất từ 1 đến N (Quy nạp toán học O(log N))",
                "desc": "Ký hiệu f(x) là ước số lẻ lớn nhất của số nguyên dương x (ví dụ f(12) = 3, f(7) = 7). Cho số nguyên dương N (1 ≤ N ≤ 10^12). Hãy tính tổng S(N) = f(1) + f(2) + ... + f(N).",
                "example": "Ví dụ: Input: `7` → Output: `21` (1 + 1 + 3 + 1 + 5 + 3 + 7 = 21).",
                "analysis": "Có k = ⌊(N + 1) / 2⌋ số lẻ trong [1, N] với tổng bằng k², còn các số chẵn 2y có f(2y) = f(y) nên tổng bằng S(⌊N / 2⌋). Truy hồi S(N) = k² + S(⌊N / 2⌋) trong O(log N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    ans = 0
    while n > 0:
        k = (n + 1) // 2
        ans += k * k
        n //= 2
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 3: Chia kẹo Euler có điều kiện (Tổ hợp & Nghịch đảo mô-đun)",
                "desc": "Có N chiếc kẹo giống hệt nhau cần chia hết cho K em bé sao cho mỗi em nhận được ít nhất M chiếc kẹo (1 ≤ N, K ≤ 10^6; 0 ≤ M ≤ 10^6). Đếm số cách chia khác nhau, lấy phần dư khi chia cho 10^9 + 7.",
                "example": "Ví dụ: Input: `7 3 1` → Output: `15` (tổ hợp chập 2 của 6).",
                "analysis": "Phát trước cho mỗi em M chiếc kẹo, còn lại R = N - K·M chiếc. Nếu R < 0 in `0`, ngược lại số cách chia chính là tổ hợp C(R + K - 1, K - 1) mod (10^9 + 7).",
                "code": """import sys


def solve():
    MOD = 10**9 + 7
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    k = int(data[1])
    m = int(data[2])

    rem = n - k * m
    if rem < 0:
        sys.stdout.write("0\\n")
        return

    top = rem + k - 1
    bot = k - 1

    fact = [1] * (top + 1)
    for i in range(1, top + 1):
        fact[i] = fact[i - 1] * i % MOD

    ans = fact[top] * pow(fact[bot], -1, MOD) % MOD * pow(fact[top - bot], -1, MOD) % MOD
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 4: Chữ số tận cùng khác 0 đầu tiên của N! (N ≤ 10^12)",
                "desc": "Cho số nguyên dương N (1 ≤ N ≤ 10^12). Hãy tìm chữ số tận cùng khác 0 đầu tiên tính từ phải sang trái của N!.",
                "example": "Ví dụ: Input: `10` → Output: `8` (10! = 3628800).",
                "analysis": "Áp dụng công thức đệ quy số học đồng dư cơ số 10 dựa vào chữ số hàng chục `tens = (x // 10) % 10`: nhân hệ số 4 (nếu tens lẻ) hoặc 6 (nếu tens chẵn) với D(⌊x / 5⌋) và D(x mod 10) trong O(log_5 N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])

    base = [1, 1, 2, 6, 4, 2, 2, 4, 2, 8]

    def last_nonzero(x):
        if x < 10:
            return base[x]
        tens = (x // 10) % 10
        coeff = 4 if (tens & 1) else 6
        return (coeff * last_nonzero(x // 5) * base[x % 10]) % 10

    sys.stdout.write(f"{last_nonzero(n)}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 5: Xâu nhị phân không có hai bit 1 kề nhau (Lũy thừa ma trận O(log N))",
                "desc": "Cho số nguyên dương N (1 ≤ N ≤ 10^18). Hãy đếm số lượng xâu nhị phân độ dài N sao cho không có hai ký tự '1' nào đứng cạnh nhau. In ra kết quả chia dư cho 10^9 + 7.",
                "example": "Ví dụ: Input: `3` → Output: `5` (gồm 000, 001, 010, 100, 101).",
                "analysis": "Số xâu hợp lệ độ dài N thỏa mãn truy hồi Fibonacci f(1) = 2, f(2) = 3, f(n) = f(n - 1) + f(n - 2) = F_{n + 2}. Dùng lũy thừa nhị phân trên ma trận 2×2 tính trong O(log N).",
                "code": """import sys


def solve():
    MOD = 10**9 + 7
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])

    def mul(A, B):
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

    res = [[1, 0], [0, 1]]
    base = [[1, 1], [1, 0]]
    p = n + 2
    while p > 0:
        if p & 1:
            res = mul(res, base)
        base = mul(base, base)
        p >>= 1

    sys.stdout.write(f"{res[0][1]}\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 2: XỬ LÝ XÂU & THUẬT TOÁN TRÊN CHUỖI (BÀI 6 – BÀI 10)",
        [
            {
                "title": "Bài 6: Xâu con đối xứng dài nhất (Kỹ thuật Mở rộng từ tâm)",
                "desc": "Cho xâu ký tự S độ dài N (1 ≤ N ≤ 5000). Hãy tìm độ dài của xâu con liên tiếp đối xứng dài nhất trong S.",
                "example": "Ví dụ: Input: `babad` → Output: `3` (`bab` hoặc `aba`).",
                "analysis": "Tại mỗi vị trí i ∈ [0, N - 1], mở rộng dần ra hai phía với tâm lẻ (i, i) và tâm chẵn (i, i + 1) để cập nhật độ dài lớn nhất trong O(N²) và bộ nhớ O(1).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    s = data[0]
    n = len(s)
    best = 1

    for i in range(n):
        l, r = i, i
        while l >= 0 and r < n and s[l] == s[r]:
            if r - l + 1 > best:
                best = r - l + 1
            l -= 1
            r += 1
        l, r = i, i + 1
        while l >= 0 and r < n and s[l] == s[r]:
            if r - l + 1 > best:
                best = r - l + 1
            l -= 1
            r += 1

    sys.stdout.write(f"{best}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 7: Xâu con ngắn nhất chứa đủ ký tự của xâu mẫu (Cửa sổ trượt)",
                "desc": "Cho hai xâu S và T (1 ≤ |T| ≤ |S| ≤ 10^6) gồm các chữ cái tiếng Anh. Tìm độ dài nhỏ nhất của một xâu con liên tiếp trong S chứa đầy đủ mọi ký tự của T (tính cả số lần lặp). In `-1` nếu không tồn tại.",
                "example": "Ví dụ: Input: `ADOBECODEBANC` và `ABC` → Output: `4` (`BANC`).",
                "analysis": "Dùng mảng tần suất `need[128]` và biến đếm `missing = len(T)`. Mở rộng biên phải `right`, khi `missing == 0` thì liên tục thu hẹp biên trái `left` trong O(|S| + |T|).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if len(data) < 2:
        return
    s, t = data[0], data[1]
    n, m = len(s), len(t)
    if m > n:
        sys.stdout.write("-1\\n")
        return

    need = [0] * 128
    for ch in t:
        need[ch] += 1

    missing = m
    left = 0
    best_len = n + 1

    for right, ch in enumerate(s):
        if need[ch] > 0:
            missing -= 1
        need[ch] -= 1

        while missing == 0:
            cur_len = right - left + 1
            if cur_len < best_len:
                best_len = cur_len
            l_ch = s[left]
            need[l_ch] += 1
            if need[l_ch] > 0:
                missing += 1
            left += 1

    sys.stdout.write(f"{best_len if best_len <= n else -1}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 8: Khử các cụm K ký tự giống nhau liên tiếp (Ngăn xếp nén)",
                "desc": "Cho xâu S (1 ≤ |S| ≤ 10^6) và số nguyên K (2 ≤ K ≤ 10^6). Liên tục xóa các cụm gồm K ký tự giống nhau đứng cạnh nhau cho đến khi không còn cụm nào như vậy. In ra xâu cuối cùng (nếu rỗng in `EMPTY`).",
                "example": "Ví dụ: Input: `deeedbbcccbdaa` và `3` → Output: `aa`.",
                "analysis": "Lưu các cặp `[char, count]` vào ngăn xếp `stack`. Khi gặp ký tự trùng đỉnh stack thì tăng `count`, nếu `count == K` thì `stack.pop()` trong O(N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if len(data) < 2:
        return
    s = data[0].decode()
    k = int(data[1])

    stack = []
    for ch in s:
        if stack and stack[-1][0] == ch:
            stack[-1][1] += 1
            if stack[-1][1] == k:
                stack.pop()
        else:
            stack.append([ch, 1])

    if not stack:
        sys.stdout.write("EMPTY\\n")
    else:
        sys.stdout.write("".join(ch * cnt for ch, cnt in stack) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 9: Đếm số lần xuất hiện của xâu mẫu P trong văn bản T (Thuật toán Z-Function)",
                "desc": "Cho hai xâu T và P (1 ≤ |P| ≤ |T| ≤ 10^6). Hãy đếm số vị trí xuất hiện của xâu P bên trong xâu T (cho phép các lần xuất hiện chồng lấn lên nhau).",
                "example": "Ví dụ: Input: `aaaaa` và `aa` → Output: `4`.",
                "analysis": "Ghép `S = P + b'#' + T` và tính mảng `Z` trong O(|P| + |T|). Đếm số vị trí `i > len(P)` thỏa mãn `Z[i] == len(P)`.",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if len(data) < 2:
        return
    t, p = data[0], data[1]
    m = len(p)
    s = p + b"#" + t
    n = len(s)

    z = [0] * n
    l = r = 0
    for i in range(1, n):
        if i <= r:
            z[i] = min(r - i + 1, z[i - l])
        while i + z[i] < n and s[z[i]] == s[i + z[i]]:
            z[i] += 1
        if i + z[i] - 1 > r:
            l = i
            r = i + z[i] - 1

    ans = sum(1 for i in range(m + 1, n) if z[i] == m)
    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 10: Truy vấn so sánh hai xâu con trong O(1) (Polynomial Rolling Hash)",
                "desc": "Cho xâu S độ dài N (1 ≤ N ≤ 2·10^5) và Q truy vấn (1 ≤ Q ≤ 2·10^5). Mỗi truy vấn gồm 4 số l1, r1, l2, r2. Kiểm tra xem hai xâu con S[l1..r1] và S[l2..r2] có giống hệt nhau không. In `YES` hoặc `NO`.",
                "example": "Ví dụ: Input: `abacaba`, `2`, `1 3 5 7`, `1 2 4 5` → Output: `YES` và `NO`.",
                "analysis": "Sử dụng băm đa thức kép (Double Rolling Hash) với cơ số 311 và hai mô-đun nguyên tố 10^9 + 7, 10^9 + 9 để so sánh hai đoạn con bất kỳ trong O(1).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    s = data[0]
    n = len(s)
    q = int(data[1])

    BASE = 311
    MOD1 = 10**9 + 7
    MOD2 = 10**9 + 9

    p1 = [1] * (n + 1)
    p2 = [1] * (n + 1)
    h1 = [0] * (n + 1)
    h2 = [0] * (n + 1)

    for i in range(1, n + 1):
        ch = s[i - 1]
        p1[i] = (p1[i - 1] * BASE) % MOD1
        p2[i] = (p2[i - 1] * BASE) % MOD2
        h1[i] = (h1[i - 1] * BASE + ch) % MOD1
        h2[i] = (h2[i - 1] * BASE + ch) % MOD2

    out = []
    idx = 2
    for _ in range(q):
        l1 = int(data[idx])
        r1 = int(data[idx + 1])
        l2 = int(data[idx + 2])
        r2 = int(data[idx + 3])
        idx += 4
        len1 = r1 - l1 + 1
        len2 = r2 - l2 + 1
        if len1 != len2:
            out.append("NO")
            continue
        val1_a = (h1[r1] - h1[l1 - 1] * p1[len1]) % MOD1
        val1_b = (h1[r2] - h1[l2 - 1] * p1[len2]) % MOD1
        val2_a = (h2[r1] - h2[l1 - 1] * p2[len1]) % MOD2
        val2_b = (h2[r2] - h2[l2 - 1] * p2[len2]) % MOD2
        out.append("YES" if val1_a == val1_b and val2_a == val2_b else "NO")

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 3: MẢNG CỘNG DỒN, MẢNG HIỆU & HAI CON TRỎ (BÀI 11 – BÀI 15)",
        [
            {
                "title": "Bài 11: Tăng đoạn [L, R] hàng loạt (Mảng hiệu 1 chiều - Difference Array)",
                "desc": "Cho mảng A gồm N số 0 (1 ≤ N ≤ 10^6) và Q thao tác (1 ≤ Q ≤ 10^6). Mỗi thao tác gồm L, R, V (1 ≤ L ≤ R ≤ N; |V| ≤ 10^9): cộng thêm V vào tất cả phần tử từ L đến R. In ra mảng A sau Q thao tác.",
                "example": "Ví dụ: Input: `5 2`, `1 3 5`, `2 5 -2` → Output: `5 3 3 -2 -2`.",
                "analysis": "Dùng mảng hiệu `diff` kích thước N + 2: mỗi truy vấn cập nhật `diff[L] += V` và `diff[R + 1] -= V` trong O(1), cuối cùng cộng dồn một lần trong O(N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    q = int(data[1])

    diff = [0] * (n + 2)
    idx = 2
    for _ in range(q):
        L = int(data[idx])
        R = int(data[idx + 1])
        V = int(data[idx + 2])
        idx += 3
        diff[L] += V
        diff[R + 1] -= V

    res = []
    cur = 0
    for i in range(1, n + 1):
        cur += diff[i]
        res.append(str(cur))

    sys.stdout.write(" ".join(res) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 12: Truy vấn tổng hình chữ nhật con (Mảng cộng dồn 2 chiều)",
                "desc": "Cho ma trận A kích thước M × N (1 ≤ M, N ≤ 1000; |A_{i,j}| ≤ 10^9) và Q truy vấn (1 ≤ Q ≤ 10^5), mỗi truy vấn gồm (x1, y1, x2, y2). Tính tổng các phần tử trong hình chữ nhật con (x1, y1) → (x2, y2).",
                "example": "Ví dụ: Input: `2 3 1`, `1 2 3`, `4 5 6`, `1 2 2 3` → Output: `16`.",
                "analysis": "Xây dựng bảng tổng tiền tố 2D `P[i][j]` trong O(M · N), trả lời mỗi truy vấn trong O(1) bằng nguyên lý bù trừ 4 góc.",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    m = int(data[0])
    n = int(data[1])
    q = int(data[2])

    P = [[0] * (n + 1) for _ in range(m + 1)]
    idx = 3
    for i in range(1, m + 1):
        row_p = P[i]
        prev_p = P[i - 1]
        for j in range(1, n + 1):
            val = int(data[idx])
            idx += 1
            row_p[j] = prev_p[j] + row_p[j - 1] - prev_p[j - 1] + val

    out = []
    for _ in range(q):
        x1 = int(data[idx])
        y1 = int(data[idx + 1])
        x2 = int(data[idx + 2])
        y2 = int(data[idx + 3])
        idx += 4
        ans = P[x2][y2] - P[x1 - 1][y2] - P[x2][y1 - 1] + P[x1 - 1][y1 - 1]
        out.append(str(ans))

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 13: Phủ hình chữ nhật trên bảng lưới (Mảng hiệu 2 chiều)",
                "desc": "Cho bảng lưới M × N (1 ≤ M, N ≤ 1000) ban đầu toàn số 0 và Q thao tác (1 ≤ Q ≤ 10^5), mỗi thao tác tăng tất cả các ô trong hình chữ nhật con (x1, y1) → (x2, y2) lên 1 đơn vị. Đếm số ô có giá trị lẻ sau Q thao tác.",
                "example": "Ví dụ: Input: `3 3 2`, `1 1 2 2`, `2 2 3 3` → Output: `6`.",
                "analysis": "Dùng mảng hiệu 2D cập nhật 4 góc trong O(1) mỗi truy vấn, sau đó cộng dồn 2D trong O(M · N) để khôi phục giá trị từng ô.",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    m = int(data[0])
    n = int(data[1])
    q = int(data[2])

    D = [[0] * (n + 2) for _ in range(m + 2)]
    idx = 3
    for _ in range(q):
        x1 = int(data[idx])
        y1 = int(data[idx + 1])
        x2 = int(data[idx + 2])
        y2 = int(data[idx + 3])
        idx += 4
        D[x1][y1] += 1
        D[x2 + 1][y1] -= 1
        D[x1][y2 + 1] -= 1
        D[x2 + 1][y2 + 1] += 1

    odd_count = 0
    for i in range(1, m + 1):
        row_d = D[i]
        prev_d = D[i - 1]
        for j in range(1, n + 1):
            row_d[j] += prev_d[j] + row_d[j - 1] - prev_d[j - 1]
            if row_d[j] & 1:
                odd_count += 1

    sys.stdout.write(f"{odd_count}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 14: Đoạn con liên tiếp ngắn nhất có tổng ≥ S (Hai con trỏ)",
                "desc": "Cho dãy gồm N số nguyên dương a_1, ..., a_N (1 ≤ N ≤ 10^6; 1 ≤ a_i ≤ 10^9) và số nguyên dương S (1 ≤ S ≤ 10^15). Tìm độ dài nhỏ nhất của đoạn con liên tiếp có tổng ≥ S (in `0` nếu không tồn tại).",
                "example": "Ví dụ: Input: `6 7` và dãy `2 3 1 2 4 3` → Output: `2` (đoạn `[4, 3]`).",
                "analysis": "Do các phần tử đều dương nên tổng đoạn tăng đơn điệu khi mở rộng biên phải. Dùng Hai con trỏ `[left, right]` chạy tuyến tính trong O(N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    s = int(data[1])
    a = list(map(int, data[2 : 2 + n]))

    left = 0
    cur_sum = 0
    min_len = n + 1

    for right, x in enumerate(a):
        cur_sum += x
        while cur_sum >= s:
            length = right - left + 1
            if length < min_len:
                min_len = length
            cur_sum -= a[left]
            left += 1

    sys.stdout.write(f"{min_len if min_len <= n else 0}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 15: Đếm số tam giác lập được từ dãy số (Sắp xếp + Hai con trỏ O(N²))",
                "desc": "Cho N đoạn thẳng có độ dài nguyên dương a_1, ..., a_N (3 ≤ N ≤ 3000; 1 ≤ a_i ≤ 10^9). Hãy đếm xem có bao nhiêu bộ ba chỉ số (i, j, k) (1 ≤ i < j < k ≤ N) ghép thành một tam giác không suy biến.",
                "example": "Ví dụ: Input: `4` và dãy `4 2 3 4` → Output: `3`.",
                "analysis": "Sắp xếp mảng tăng dần. Cố định cạnh lớn nhất tại chỉ số `k`, dùng hai con trỏ `l = 0, r = k - 1`: nếu `a[l] + a[r] > a[k]` thì cộng thêm `r - l` và giảm `r` trong O(N²).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    a = sorted(map(int, data[1 : 1 + n]))

    ans = 0
    for k in range(n - 1, 1, -1):
        l, r = 0, k - 1
        target = a[k]
        while l < r:
            if a[l] + a[r] > target:
                ans += r - l
                r -= 1
            else:
                l += 1

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 4: CẤU TRÚC DỮ LIỆU (DICT, SET, DEQUE, HEAPQ, DSU) (BÀI 16 – BÀI 20)",
        [
            {
                "title": "Bài 16: Đoạn con cân bằng chẵn lẻ (Bảng băm + Quy đổi trọng số)",
                "desc": "Cho dãy gồm N số nguyên a_1, ..., a_N (1 ≤ N ≤ 10^6; |a_i| ≤ 10^9). Đếm số đoạn con liên tiếp chứa số lượng số chẵn bằng đúng số lượng số lẻ.",
                "example": "Ví dụ: Input: `4` và dãy `1 2 3 4` → Output: `4`.",
                "analysis": "Quy đổi số chẵn thành `+1` và số lẻ thành `-1`, bài toán quy về đếm số đoạn con có tổng bằng `0` bằng từ điển lưu tần suất tổng tiền tố trong O(N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])

    cnt = {0: 1}
    pref = 0
    ans = 0
    for i in range(1, n + 1):
        x = int(data[i])
        pref += -1 if (x & 1) else 1
        c = cnt.get(pref, 0)
        ans += c
        cnt[pref] = c + 1

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 17: Diện tích hình chữ nhật lớn nhất trong biểu đồ cột (Ngăn xếp đơn điệu)",
                "desc": "Cho biểu đồ gồm N cột liền nhau chiều rộng 1 và chiều cao h_1, ..., h_N (1 ≤ N ≤ 10^6; 0 ≤ h_i ≤ 10^9). Tìm diện tích hình chữ nhật lớn nhất nằm trọn bên trong biểu đồ.",
                "example": "Ví dụ: Input: `6` và dãy `2 1 5 6 2 3` → Output: `10`.",
                "analysis": "Dùng ngăn xếp đơn điệu tăng lưu chỉ số các cột kết hợp lính canh `0` ở cuối mảng để tính diện tích hình chữ nhật cực đại ứng với mỗi cột làm chiều cao trong O(N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    h = list(map(int, data[1 : 1 + n]))
    h.append(0)

    stack = []
    max_area = 0
    for i, height in enumerate(h):
        while stack and h[stack[-1]] > height:
            top_h = h[stack.pop()]
            width = i if not stack else (i - stack[-1] - 1)
            area = top_h * width
            if area > max_area:
                max_area = area
        stack.append(i)

    sys.stdout.write(f"{max_area}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 18: Đoạn con ngắn nhất có tổng ≥ K trên mảng có số âm (Monotonic Deque)",
                "desc": "Cho dãy gồm N số nguyên a_1, ..., a_N (1 ≤ N ≤ 5·10^5; |a_i| ≤ 10^9, có thể âm) và số nguyên dương K (1 ≤ K ≤ 10^15). Tìm độ dài nhỏ nhất của đoạn con liên tiếp có tổng ≥ K (in `-1` nếu không tồn tại).",
                "example": "Ví dụ: Input: `3 3` và dãy `2 -1 2` → Output: `3`.",
                "analysis": "Tính mảng tổng tiền tố `P[0..N]` và duy trì `deque` lưu các chỉ số có giá trị `P` tăng dần nghiêm ngặt để tìm đầu mút trái gần nhất thỏa mãn `P[i] - P[j] ≥ K` trong O(N).",
                "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    k = int(data[1])

    P = [0] * (n + 1)
    cur = 0
    for i in range(1, n + 1):
        cur += int(data[i + 1])
        P[i] = cur

    dq = deque()
    ans = n + 1
    for i, val in enumerate(P):
        while dq and val - P[dq[0]] >= k:
            length = i - dq.popleft()
            if length < ans:
                ans = length
        while dq and P[dq[-1]] >= val:
            dq.pop()
        dq.append(i)

    sys.stdout.write(f"{ans if ans <= n else -1}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 19: Tối đa số lượng dự án hoàn thành đúng hạn (Moore-Hodgson + `heapq`)",
                "desc": "Có N dự án (1 ≤ N ≤ 2·10^5), dự án thứ i cần thời gian t_i và phải xong trước hạn chót d_i (1 ≤ t_i, d_i ≤ 10^9). Tìm số lượng dự án tối đa hoàn thành đúng hạn.",
                "example": "Ví dụ: Input: `4`, `(100, 200), (200, 1300), (1000, 1250), (2000, 3200)` → Output: `3`.",
                "analysis": "Sắp xếp các dự án theo hạn chót d_i tăng dần, dùng Max-Heap (`heapq` với giá trị âm) để khi tổng thời gian vượt quá d_i thì loại bỏ dự án tốn nhiều thời gian nhất đã chọn trong O(N log N).",
                "code": """import heapq
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    jobs = []
    idx = 1
    for _ in range(n):
        t = int(data[idx])
        d = int(data[idx + 1])
        idx += 2
        if t <= d:
            jobs.append((d, t))

    jobs.sort()
    pq = []
    cur_time = 0
    for d, t in jobs:
        cur_time += t
        heapq.heappush(pq, -t)
        if cur_time > d:
            cur_time += heapq.heappop(pq)

    sys.stdout.write(f"{len(pq)}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 20: Kết nối mạng động (Hệ các tập hợp rời nhau - DSU / Union-Find)",
                "desc": "Có N máy tính rời nhau (1 ≤ N ≤ 2·10^5) và M lần nối dây cáp giữa hai máy (u_i, v_i) (1 ≤ M ≤ 2·10^5). Sau mỗi lần nối dây, in ra số cụm mạng liên thông hiện tại và kích thước cụm lớn nhất.",
                "example": "Ví dụ: Input: `5 3`, `(1, 2), (1, 3), (4, 5)` → Output: `4 2`, `3 3`, `2 3`.",
                "analysis": "Cài đặt DSU với tối ưu nén đường đi (Path Compression) và gộp theo kích thước (Union by Size) để xử lý mỗi truy vấn trong gần như O(1).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    m = int(data[1])

    parent = list(range(n + 1))
    sz = [1] * (n + 1)
    components = n
    max_sz = 1

    def find(u):
        root = u
        while parent[root] != root:
            root = parent[root]
        while parent[u] != root:
            nxt = parent[u]
            parent[u] = root
            u = nxt
        return root

    out = []
    idx = 2
    for _ in range(m):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        ru, rv = find(u), find(v)
        if ru != rv:
            if sz[ru] < sz[rv]:
                ru, rv = rv, ru
            parent[rv] = ru
            sz[ru] += sz[rv]
            components -= 1
            if sz[ru] > max_sz:
                max_sz = sz[ru]
        out.append(f"{components} {max_sz}")

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 5: TÌM KIẾM NHỊ PHÂN & THUẬT TOÁN THAM LAM (BÀI 21 – BÀI 25)",
        [
            {
                "title": "Bài 21: Cắt dây cáp mạng (Tìm kiếm nhị phân trên kết quả)",
                "desc": "Có N cuộn dây cáp có độ dài nguyên dương L_1, ..., L_N (1 ≤ N ≤ 10^5; 1 ≤ L_i ≤ 10^9). Cần cắt ra ít nhất K đoạn dây bằng nhau có độ dài nguyên dương X (1 ≤ K ≤ 10^9). Tìm X lớn nhất (in `0` nếu không thể cắt).",
                "example": "Ví dụ: Input: `4 11` và độ dài `802 743 457 539` → Output: `200`.",
                "analysis": "Tìm kiếm nhị phân độ dài X trong [1, max(L)] kết hợp kiểm tra tổng số đoạn `sum(x // mid for x in L) >= K` trong O(N log(max L)).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    k = int(data[1])
    L = list(map(int, data[2 : 2 + n]))

    low, high = 1, max(L)
    ans = 0
    while low <= high:
        mid = (low + high) // 2
        pieces = sum(x // mid for x in L)
        if pieces >= k:
            ans = mid
            low = mid + 1
        else:
            high = mid - 1

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 22: Dây chuyền nướng bánh tự động (Tìm kiếm nhị phân trên thời gian)",
                "desc": "Có N lò nướng hoạt động song song (1 ≤ N ≤ 2·10^5), lò thứ i cần t_i giây để nướng xong 1 chiếc bánh (1 ≤ t_i ≤ 10^9). Tìm thời gian tối thiểu để nướng được ít nhất M chiếc bánh (1 ≤ M ≤ 10^9).",
                "example": "Ví dụ: Input: `3 7` và thời gian `3 2 5` → Output: `8`.",
                "analysis": "Trong T giây, tổng số bánh nướng được là `sum(T // t_i)`. Tìm kiếm nhị phân T trong đoạn `[1, min(t) * M]` trong O(N log(min(t) · M)).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    m = int(data[1])
    t = list(map(int, data[2 : 2 + n]))

    low = 1
    high = min(t) * m
    ans = high

    while low <= high:
        mid = (low + high) // 2
        baked = 0
        for ti in t:
            baked += mid // ti
            if baked >= m:
                break
        if baked >= m:
            ans = mid
            high = mid - 1
        else:
            low = mid + 1

    sys.stdout.write(f"{ans}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 23: Số phòng họp tối thiểu cần thiết (Quét sự kiện - Sweep-line)",
                "desc": "Có N cuộc họp (1 ≤ N ≤ 2·10^5), cuộc họp thứ i diễn ra trong khoảng [s_i, e_i) (0 ≤ s_i < e_i ≤ 10^9). Tìm số phòng họp tối thiểu cần có để phục vụ toàn bộ các cuộc họp.",
                "example": "Ví dụ: Input: `3`, `(0, 30), (5, 10), (15, 20)` → Output: `2`.",
                "analysis": "Tạo các cặp sự kiện `(s_i, 1)` và `(e_i, -1)`, sắp xếp tăng dần và tìm giá trị cực đại của tổng tiền tố trong O(N log N).",
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
    cur = 0
    best = 0
    for _, delta in events:
        cur += delta
        if cur > best:
            best = cur

    sys.stdout.write(f"{best}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 24: Lập lịch giảm thiểu tổng phạt trễ hạn (Quy tắc Smith - Tích chéo)",
                "desc": "Có N đơn hàng (1 ≤ N ≤ 2·10^5), đơn hàng thứ i mất t_i phút gia công và chịu mức phạt w_i đồng cho mỗi phút tính từ thời điểm 0 đến khi hoàn thành (1 ≤ t_i, w_i ≤ 10^6). Tìm tổng tiền phạt nhỏ nhất.",
                "example": "Ví dụ: Input: `2`, `(3, 10), (1, 5)` → Output: `45`.",
                "analysis": "Áp dụng quy tắc Smith: sắp xếp các đơn hàng sao cho `t_i * w_j < t_j * w_i` bằng `functools.cmp_to_key` (số nguyên 100%, không dùng phép chia số thực).",
                "code": """from functools import cmp_to_key
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    orders = []
    idx = 1
    for _ in range(n):
        t = int(data[idx])
        w = int(data[idx + 1])
        idx += 2
        orders.append((t, w))

    def cmp_order(a, b):
        left = a[0] * b[1]
        right = b[0] * a[1]
        if left < right:
            return -1
        if left > right:
            return 1
        return 0

    orders.sort(key=cmp_to_key(cmp_order))

    cur_t = 0
    total_penalty = 0
    for t, w in orders:
        cur_t += t
        total_penalty += cur_t * w

    sys.stdout.write(f"{total_penalty}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 25: Thuyền cứu hộ vượt sông (Tham lam + Hai con trỏ)",
                "desc": "Có N người cần qua sông (1 ≤ N ≤ 2·10^5), người thứ i nặng w_i (1 ≤ w_i ≤ W ≤ 10^9). Mỗi thuyền chở tối đa 2 người với tổng cân nặng ≤ W. Tìm số chiếc thuyền ít nhất cần dùng.",
                "example": "Ví dụ: Input: `4 100` và cân nặng `70 50 80 50` → Output: `3`.",
                "analysis": "Sắp xếp cân nặng tăng dần, dùng hai con trỏ `l = 0, r = N - 1`: ghép người nhẹ nhất với người nặng nhất nếu tổng ≤ W, ngược lại người nặng nhất đi riêng trong O(N log N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    limit = int(data[1])
    w = sorted(map(int, data[2 : 2 + n]))

    l, r = 0, n - 1
    boats = 0
    while l <= r:
        if w[l] + w[r] <= limit:
            l += 1
        r -= 1
        boats += 1

    sys.stdout.write(f"{boats}\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 6: QUY HOẠCH ĐỘNG CƠ BẢN ĐẾN NÂNG CAO (BÀI 26 – BÀI 32)",
        [
            {
                "title": "Bài 26: Đếm số cách leo cầu thang có bậc hỏng (DP 1D + Cửa sổ trượt O(N))",
                "desc": "Cầu thang có N bậc (1 ≤ K ≤ N ≤ 10^6), trong đó có M bậc bị hỏng. Xuất phát từ bậc 0, mỗi bước nhảy lên 1, 2, ..., K bậc. Đếm số cách đi tới bậc N chia dư cho 10^9 + 7.",
                "example": "Ví dụ: Input: `5 2 1` và bậc hỏng `3` → Output: `2`.",
                "analysis": "Duy trì tổng trượt `window_sum` của K phần tử `dp` liền trước để tính mỗi trạng thái `dp[i]` trong O(1), tổng độ phức tạp O(N).",
                "code": """import sys


def solve():
    MOD = 10**9 + 7
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    k = int(data[1])
    m = int(data[2])

    broken = bytearray(n + 1)
    for i in range(3, 3 + m):
        broken[int(data[i])] = 1

    dp = [0] * (n + 1)
    dp[0] = 1
    window_sum = 1

    for i in range(1, n + 1):
        if i - k - 1 >= 0:
            window_sum = (window_sum - dp[i - k - 1]) % MOD
        if not broken[i]:
            dp[i] = window_sum
            window_sum = (window_sum + dp[i]) % MOD

    sys.stdout.write(f"{dp[n]}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 27: Chia kẹo thành hai phần có chênh lệch nhỏ nhất (DP Subset Sum)",
                "desc": "Cho N gói kẹo (1 ≤ N ≤ 200), gói thứ i có a_i viên kẹo (1 ≤ a_i ≤ 1000). Hãy chia N gói kẹo thành 2 phần sao cho chênh lệch tổng số kẹo giữa hai phần là nhỏ nhất.",
                "example": "Ví dụ: Input: `4` và dãy `3 1 4 2` → Output: `0`.",
                "analysis": "Dùng mảng đánh dấu `dp[0 .. S // 2]` kiểu Cái túi 0/1 để tìm tổng lớn nhất j ≤ S // 2 tạo được; chênh lệch nhỏ nhất là S - 2j.",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    a = list(map(int, data[1 : 1 + n]))

    total = sum(a)
    half = total // 2
    dp = bytearray(half + 1)
    dp[0] = 1

    for x in a:
        for j in range(half, x - 1, -1):
            if dp[j - x]:
                dp[j] = 1

    for j in range(half, -1, -1):
        if dp[j]:
            sys.stdout.write(f"{total - 2 * j}\\n")
            return


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 28: Bài toán Cái túi vô hạn (Unbounded Knapsack)",
                "desc": "Có N loại vật phẩm (1 ≤ N ≤ 1000), số lượng mỗi loại không giới hạn, trọng lượng w_i (1 ≤ w_i ≤ 10^4) và giá trị v_i (1 ≤ v_i ≤ 10^9). Ba lô sức chứa W (1 ≤ W ≤ 10^4). Tìm tổng giá trị lớn nhất.",
                "example": "Ví dụ: Input: `3 10`, `(3, 4), (4, 5), (5, 7)` → Output: `14`.",
                "analysis": "Duyệt trọng lượng `j` từ nhỏ đến lớn (`range(w, W + 1)`) để cho phép mỗi loại vật phẩm được chọn nhiều lần trong O(N · W).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    W = int(data[1])

    dp = [0] * (W + 1)
    idx = 2
    for _ in range(n):
        w = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        for j in range(w, W + 1):
            val = dp[j - w] + v
            if val > dp[j]:
                dp[j] = val

    sys.stdout.write(f"{dp[W]}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 29: Khoảng cách biến đổi xâu (Edit Distance - Levenshtein DP 1D)",
                "desc": "Cho hai xâu ký tự A và B (1 ≤ |A|, |B| ≤ 2000). Mỗi bước được phép thêm, xóa hoặc thay thế 1 ký tự trên xâu A. Tìm số bước ít nhất để biến xâu A thành xâu B.",
                "example": "Ví dụ: Input: `horse` và `ros` → Output: `3`.",
                "analysis": "Quy hoạch động nén mảng 1 chiều kết hợp biến `prev_diag` lưu trạng thái chéo trên trái để đạt độ phức tạp O(|A| · |B|) và bộ nhớ O(|B|).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if len(data) < 2:
        return
    a, b = data[0], data[1]
    m = len(b)
    dp = list(range(m + 1))

    for i, ch_a in enumerate(a, 1):
        prev_diag = dp[0]
        dp[0] = i
        for j, ch_b in enumerate(b, 1):
            temp = dp[j]
            if ch_a == ch_b:
                dp[j] = prev_diag
            else:
                best = prev_diag
                if dp[j] < best:
                    best = dp[j]
                if dp[j - 1] < best:
                    best = dp[j - 1]
                dp[j] = best + 1
            prev_diag = temp

    sys.stdout.write(f"{dp[m]}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 30: Đếm số lượng dãy con tăng nghiêm ngặt dài nhất (O(N²))",
                "desc": "Cho dãy gồm N số nguyên a_1, ..., a_N (1 ≤ N ≤ 2000; |a_i| ≤ 10^9). Tìm độ dài của dãy con tăng nghiêm ngặt dài nhất và số lượng dãy con tăng đạt được độ dài cực đại đó (chia dư 10^9 + 7).",
                "example": "Ví dụ: Input: `5` và dãy `1 3 5 4 7` → Output: `4 2`.",
                "analysis": "Duy trì song song mảng độ dài `length[i]` và mảng số cách `count[i]` kết thúc tại vị trí i trong O(N²).",
                "code": """import sys


def solve():
    MOD = 10**9 + 7
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    a = list(map(int, data[1 : 1 + n]))

    length = [1] * n
    count = [1] * n
    max_len = 1

    for i in range(n):
        ai = a[i]
        li = length[i]
        ci = count[i]
        for j in range(i):
            if a[j] < ai:
                if length[j] + 1 > li:
                    li = length[j] + 1
                    ci = count[j]
                elif length[j] + 1 == li:
                    ci = (ci + count[j]) % MOD
        length[i] = li
        count[i] = ci
        if li > max_len:
            max_len = li

    total_ways = sum(count[i] for i in range(n) if length[i] == max_len) % MOD
    sys.stdout.write(f"{max_len} {total_ways}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 31: Ghép các đống sỏi kề nhau (Quy hoạch động trên khoảng - Interval DP)",
                "desc": "Có N đống sỏi xếp thành hàng ngang (1 ≤ N ≤ 300), đống thứ i có a_i viên sỏi (1 ≤ a_i ≤ 10^6). Mỗi bước chỉ được gộp hai đống kề nhau với chi phí bằng tổng số sỏi của hai đống đó. Tìm tổng chi phí nhỏ nhất.",
                "example": "Ví dụ: Input: `4` và dãy `40 20 30 10` → Output: `190`.",
                "analysis": "Quy hoạch động trên đoạn [i, j] theo độ dài tăng dần: `dp[i][j] = min(dp[i][k] + dp[k + 1][j]) + sum(i, j)` trong O(N³).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    a = list(map(int, data[1 : 1 + n]))

    pref = [0] * (n + 1)
    for i in range(n):
        pref[i + 1] = pref[i] + a[i]

    INF = 10**18
    dp = [[0] * n for _ in range(n)]

    for length in range(2, n + 1):
        for i in range(n - length + 1):
            j = i + length - 1
            best = INF
            dp_i = dp[i]
            for k in range(i, j):
                val = dp_i[k] + dp[k + 1][j]
                if val < best:
                    best = val
            dp_i[j] = best + (pref[j + 1] - pref[i])

    sys.stdout.write(f"{dp[0][n - 1]}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 32: Phân công công việc tối ưu (Quy hoạch động trạng thái Bitmask)",
                "desc": "Có N nhân viên và N công việc (1 ≤ N ≤ 18). Chi phí để nhân viên i làm công việc j là C_{i, j} (1 ≤ C_{i, j} ≤ 10^6). Phân công mỗi nhân viên làm đúng một việc khác nhau sao cho tổng chi phí nhỏ nhất.",
                "example": "Ví dụ: Input: `3` và ma trận `[[9, 2, 7], [6, 4, 3], [5, 8, 1]]` → Output: `9`.",
                "analysis": "Dùng mặt nạ bit `mask` biểu diễn tập công việc đã giao; chỉ số nhân viên tiếp theo chính là `mask.bit_count()`. Độ phức tạp O(N · 2^N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    c = []
    idx = 1
    for _ in range(n):
        c.append(list(map(int, data[idx : idx + n])))
        idx += n

    size = 1 << n
    INF = 10**18
    dp = [INF] * size
    dp[0] = 0

    for mask in range(size):
        cur_val = dp[mask]
        if cur_val == INF:
            continue
        person = mask.bit_count()
        if person >= n:
            continue
        row = c[person]
        for job in range(n):
            if not (mask & (1 << job)):
                nmask = mask | (1 << job)
                nval = cur_val + row[job]
                if nval < dp[nmask]:
                    dp[nmask] = nval

    sys.stdout.write(f"{dp[size - 1]}\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 7: LÝ THUYẾT ĐỒ THỊ & CÂY (BÀI 33 – BÀI 38)",
        [
            {
                "title": "Bài 33: Thoát khỏi đám cháy rừng (BFS đa nguồn - Multi-source BFS)",
                "desc": "Khu rừng lưới M × N (1 ≤ M, N ≤ 1000) gồm `.` (trống), `#` (tường), `F` (điểm phát lửa, có thể nhiều điểm) và `J` (vị trí của bạn). Mỗi phút bạn và lửa cùng lan sang 4 ô kề cạnh. Tìm số phút ít nhất để bước ra khỏi biên lưới (in `IMPOSSIBLE` nếu không thể thoát).",
                "example": "Ví dụ: Input: `4 4`, `####`, `#JF#`, `#..#`, `#..#` → Output: `3`.",
                "analysis": "Chạy BFS đa nguồn từ tất cả ô `F` trước để tính mảng `fire_time`, sau đó chạy BFS từ `J` với điều kiện thời điểm đến ô mới `nd < fire_time[nr][nc]` trong O(M · N).",
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
    fire_time = [[INF] * n for _ in range(m)]
    dq_fire = deque()
    jr = jc = -1

    for r in range(m):
        for c in range(n):
            ch = grid[r][c]
            if ch == 70:  # 'F'
                fire_time[r][c] = 0
                dq_fire.append((r, c))
            elif ch == 74:  # 'J'
                jr, jc = r, c

    dirs = [(-1, 0), (1, 0), (0, -1), (0, 1)]
    while dq_fire:
        r, c = dq_fire.popleft()
        nt = fire_time[r][c] + 1
        for dr, dc in dirs:
            nr, nc = r + dr, c + dc
            if 0 <= nr < m and 0 <= nc < n and grid[nr][nc] != 35 and fire_time[nr][nc] == INF:
                fire_time[nr][nc] = nt
                dq_fire.append((nr, nc))

    dist = [[-1] * n for _ in range(m)]
    dist[jr][jc] = 0
    dq = deque([(jr, jc)])

    while dq:
        r, c = dq.popleft()
        d = dist[r][c]
        if r == 0 or r == m - 1 or c == 0 or c == n - 1:
            sys.stdout.write(f"{d + 1}\\n")
            return
        nd = d + 1
        for dr, dc in dirs:
            nr, nc = r + dr, c + dc
            if (
                0 <= nr < m
                and 0 <= nc < n
                and grid[nr][nc] != 35
                and dist[nr][nc] == -1
                and nd < fire_time[nr][nc]
            ):
                dist[nr][nc] = nd
                dq.append((nr, nc))

    sys.stdout.write("IMPOSSIBLE\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 34: Đường đi ít lần đổi hướng nhất trên lưới (0-1 BFS mở rộng trạng thái)",
                "desc": "Cho lưới M × N (1 ≤ M, N ≤ 500) gồm `.` (trống) và `#` (tường). Xuất phát từ ô (1, 1) đi tới ô (M, N) theo 4 hướng (bước đầu tiên được tự do chọn hướng). Tìm số lần rẽ hướng ít nhất để đến đích (in `-1` nếu không đến được).",
                "example": "Ví dụ: Input: `3 3`, `...`, `##.`, `...` → Output: `2`.",
                "analysis": "Mở rộng trạng thái `(r, c, dir)` và dùng 0-1 BFS với `deque`: đi thẳng cùng hướng có trọng số 0 (`appendleft`), đổi hướng có trọng số 1 (`append`) trong O(4 · M · N).",
                "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    m = int(data[0])
    n = int(data[1])
    grid = data[2 : 2 + m]

    if grid[0][0] == 35 or grid[m - 1][n - 1] == 35:
        sys.stdout.write("-1\\n")
        return
    if m == 1 and n == 1:
        sys.stdout.write("0\\n")
        return

    INF = 10**9
    total_states = m * n * 4
    dist = [INF] * total_states
    dq = deque()

    for d in range(4):
        dist[d] = 0
        dq.append(d)

    dirs = [(-1, 0), (1, 0), (0, -1), (0, 1)]
    target_cell = (m - 1) * n + (n - 1)

    while dq:
        st = dq.popleft()
        cur_d = dist[st]
        cell, d = divmod(st, 4)
        if cell == target_cell:
            sys.stdout.write(f"{cur_d}\\n")
            return
        r, c = divmod(cell, n)
        for nd, (dr, dc) in enumerate(dirs):
            nr, nc = r + dr, c + dc
            if 0 <= nr < m and 0 <= nc < n and grid[nr][nc] != 35:
                w = 0 if nd == d else 1
                nst = (nr * n + nc) * 4 + nd
                if cur_d + w < dist[nst]:
                    dist[nst] = cur_d + w
                    if w == 0:
                        dq.appendleft(nst)
                    else:
                        dq.append(nst)

    sys.stdout.write("-1\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 35: Đường đi ngắn nhất và Truy vết lộ trình (Dijkstra + Parent Array)",
                "desc": "Cho đồ thị vô hướng có trọng số dương gồm N đỉnh và M cạnh (1 ≤ N ≤ 10^5; 1 ≤ M ≤ 2·10^5; 1 ≤ w_i ≤ 10^9). Tìm đường đi ngắn nhất từ đỉnh 1 đến đỉnh N: dòng 1 in độ dài ngắn nhất, dòng 2 in lộ trình các đỉnh đi qua (in `-1` nếu không có đường đi).",
                "example": "Ví dụ: Input: `5 6`, `(1,2,2), (2,5,5), (2,3,4), (1,4,1), (4,3,3), (3,5,1)` → Output: dòng 1 là `5`, dòng 2 là `1 4 3 5`.",
                "analysis": "Cài đặt Dijkstra kết hợp mảng `parent[v] = u` để ghi nhận đỉnh liền trước khi thư giãn cạnh, sau đó truy vết ngược từ N về 1 trong O((N + M) log N).",
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
        adj[v].append((u, w))

    INF = 10**18
    dist = [INF] * (n + 1)
    parent = [0] * (n + 1)
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
                parent[v] = u
                heapq.heappush(pq, (nd, v))

    if dist[n] == INF:
        sys.stdout.write("-1\\n")
        return

    path = []
    cur = n
    while cur != 0:
        path.append(str(cur))
        cur = parent[cur]
    path.reverse()

    sys.stdout.write(f"{dist[n]}\\n")
    sys.stdout.write(" ".join(path) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 36: Cây khung nhỏ nhất (Thuật toán Kruskal + DSU)",
                "desc": "Cho đơn đồ thị vô hướng gồm N đỉnh và M cạnh (1 ≤ N ≤ 10^5; 1 ≤ M ≤ 2·10^5), cạnh (u_i, v_i) có chi phí w_i (1 ≤ w_i ≤ 10^9). Tìm tổng chi phí nhỏ nhất để nối thông toàn bộ N đỉnh (in `IMPOSSIBLE` nếu không liên thông).",
                "example": "Ví dụ: Input: `5 6`, `(1,2,3), (2,3,5), (2,4,2), (3,4,8), (5,1,7), (5,4,4)` → Output: `14`.",
                "analysis": "Sắp xếp danh sách cạnh theo chi phí w_i tăng dần và dùng DSU chọn dần N - 1 cạnh không tạo chu trình trong O(M log M).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    m = int(data[1])

    edges = []
    idx = 2
    for _ in range(m):
        u = int(data[idx])
        v = int(data[idx + 1])
        w = int(data[idx + 2])
        idx += 3
        edges.append((w, u, v))

    edges.sort()
    parent = list(range(n + 1))

    def find(u):
        root = u
        while parent[root] != root:
            root = parent[root]
        while parent[u] != root:
            nxt = parent[u]
            parent[u] = root
            u = nxt
        return root

    total_cost = 0
    edges_used = 0
    for w, u, v in edges:
        ru, rv = find(u), find(v)
        if ru != rv:
            parent[rv] = ru
            total_cost += w
            edges_used += 1
            if edges_used == n - 1:
                break

    if edges_used == n - 1:
        sys.stdout.write(f"{total_cost}\\n")
    else:
        sys.stdout.write("IMPOSSIBLE\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 37: Đường đi dài nhất trên đồ thị có hướng không chu trình (Topo Sort + DAG DP)",
                "desc": "Cho đồ thị có hướng không chu trình (DAG) gồm N đỉnh và M cung (1 ≤ N ≤ 10^5; 1 ≤ M ≤ 2·10^5). Hãy tìm độ dài (tính theo số cạnh) của đường đi dài nhất trên đồ thị này.",
                "example": "Ví dụ: Input: `4 5`, `(1,2), (1,3), (3,2), (2,4), (3,4)` → Output: `3`.",
                "analysis": "Kết hợp sắp xếp Tô-pô Kahn (`deque` + `in_deg`) và quy hoạch động trên DAG `dp[v] = max(dp[v], dp[u] + 1)` trong O(N + M).",
                "code": """from collections import deque
import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
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
    dp = [0] * (n + 1)

    while dq:
        u = dq.popleft()
        du = dp[u] + 1
        for v in adj[u]:
            if du > dp[v]:
                dp[v] = du
            in_deg[v] -= 1
            if in_deg[v] == 0:
                dq.append(v)

    sys.stdout.write(f"{max(dp)}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 38: Khoảng cách giữa hai đỉnh bất kỳ trên Cây (LCA bằng Nhảy nhị phân)",
                "desc": "Cho một cây gồm N đỉnh (1 ≤ N ≤ 10^5) và Q truy vấn (1 ≤ Q ≤ 10^5). Mỗi truy vấn gồm hai đỉnh u, v. Tính khoảng cách (số cạnh) giữa hai đỉnh u và v trên cây.",
                "example": "Ví dụ: Input: `5 3`, cạnh `(1,2), (1,3), (3,4), (3,5)`, truy vấn `(2, 4), (4, 5), (1, 5)` → Output: `3`, `2`, `2`.",
                "analysis": "Xây dựng bảng nhảy nhị phân `up[k][u]` và độ sâu `depth[u]` bằng BFS từ gốc 1 trong O(N log N). Trả lời mỗi truy vấn trong O(log N) qua `depth[u] + depth[v] - 2 * depth[lca(u, v)]`.",
                "code": """from collections import deque
import sys


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

    LOG = 18
    up = [[0] * (n + 1) for _ in range(LOG)]
    depth = [-1] * (n + 1)
    depth[1] = 0
    dq = deque([1])

    while dq:
        u = dq.popleft()
        for v in adj[u]:
            if depth[v] == -1:
                depth[v] = depth[u] + 1
                up[0][v] = u
                dq.append(v)

    for k in range(1, LOG):
        prev_up = up[k - 1]
        cur_up = up[k]
        for u in range(1, n + 1):
            cur_up[u] = prev_up[prev_up[u]]

    def lca(u, v):
        if depth[u] < depth[v]:
            u, v = v, u
        diff = depth[u] - depth[v]
        for k in range(LOG):
            if (diff >> k) & 1:
                u = up[k][u]
        if u == v:
            return u
        for k in range(LOG - 1, -1, -1):
            if up[k][u] != up[k][v]:
                u = up[k][u]
                v = up[k][v]
        return up[0][u]

    out = []
    for _ in range(q):
        u = int(data[idx])
        v = int(data[idx + 1])
        idx += 2
        w = lca(u, v)
        out.append(str(depth[u] + depth[v] - 2 * depth[w]))

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
    (
        "MẢNG 8: CẤU TRÚC DỮ LIỆU NÂNG CAO & KỸ THUẬT QUỐC GIA (BÀI 39 – BÀI 40)",
        [
            {
                "title": "Bài 39: Đếm số cặp nghịch thế trong mảng lớn (Rời rạc hóa + Fenwick Tree / BIT)",
                "desc": "Cho dãy gồm N số nguyên a_1, ..., a_N (1 ≤ N ≤ 2·10^5; |a_i| ≤ 10^9). Hãy đếm số cặp chỉ số (i, j) thỏa mãn 1 ≤ i < j ≤ N và a_i > a_j.",
                "example": "Ví dụ: Input: `5` và dãy `2 4 1 3 5` → Output: `3` (các cặp `(2, 1), (4, 1), (4, 3)`).",
                "analysis": "Rời rạc hóa giá trị các phần tử về miền [1, M] (M ≤ N) và dùng Cây chỉ số nhị phân (Fenwick Tree / BIT) đếm số phần tử lớn hơn đứng trước mỗi vị trí trong O(N log N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    a = list(map(int, data[1 : 1 + n]))

    vals = sorted(set(a))
    comp = {v: i + 1 for i, v in enumerate(vals)}
    m = len(vals)

    bit = [0] * (m + 1)
    inv_count = 0

    for inserted, x in enumerate(a):
        idx = comp[x]
        s = 0
        q_idx = idx
        while q_idx > 0:
            s += bit[q_idx]
            q_idx -= q_idx & -q_idx
        inv_count += inserted - s

        u_idx = idx
        while u_idx <= m:
            bit[u_idx] += 1
            u_idx += u_idx & -u_idx

    sys.stdout.write(f"{inv_count}\\n")


if __name__ == "__main__":
    solve()""",
            },
            {
                "title": "Bài 40: Cập nhật điểm và Truy vấn giá trị nhỏ nhất trên đoạn (Iterative Segment Tree)",
                "desc": "Cho dãy gồm N số nguyên a_1, ..., a_N (1 ≤ N ≤ 2·10^5; |a_i| ≤ 10^9) và Q truy vấn (1 ≤ Q ≤ 2·10^5): `1 k u` (gán a_k = u) hoặc `2 L R` (tìm giá trị nhỏ nhất trên đoạn [L, R]).",
                "example": "Ví dụ: Input: `5 3`, `3 2 4 5 1`, `2 1 4`, `1 2 8`, `2 1 4` → Output: `2` và `3`.",
                "analysis": "Cài đặt Cây phân đoạn dạng lặp (Iterative Segment Tree) kích thước 2N để loại bỏ hoàn toàn chi phí đệ quy và xử lý mỗi truy vấn trong O(log N).",
                "code": """import sys


def solve():
    data = sys.stdin.buffer.read().split()
    if not data:
        return
    n = int(data[0])
    q = int(data[1])
    a = list(map(int, data[2 : 2 + n]))

    INF = 10**18
    seg = [INF] * (2 * n)
    seg[n : 2 * n] = a
    for i in range(n - 1, 0, -1):
        left_val = seg[i << 1]
        right_val = seg[i << 1 | 1]
        seg[i] = left_val if left_val < right_val else right_val

    out = []
    idx = 2 + n
    for _ in range(q):
        op = int(data[idx])
        if op == 1:
            pos = int(data[idx + 1]) - 1 + n
            val = int(data[idx + 2])
            idx += 3
            seg[pos] = val
            while pos > 1:
                pos >>= 1
                l_v = seg[pos << 1]
                r_v = seg[pos << 1 | 1]
                seg[pos] = l_v if l_v < r_v else r_v
        else:
            l = int(data[idx + 1]) - 1 + n
            r = int(data[idx + 2]) + n
            idx += 3
            res = INF
            while l < r:
                if l & 1:
                    if seg[l] < res:
                        res = seg[l]
                    l += 1
                if r & 1:
                    r -= 1
                    if seg[r] < res:
                        res = seg[r]
                l >>= 1
                r >>= 1
            out.append(str(res))

    sys.stdout.write("\\n".join(out) + "\\n")


if __name__ == "__main__":
    solve()""",
            },
        ],
    ),
]


# ==============================================================================
# HÀM KHỞI TẠO & XUẤT TỆP WORD HOÀN CHỈNH
# ==============================================================================


def build_document(output_filename):
    doc = setup_document()

    # Tiêu đề chính
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    trun = title_p.add_run(
        "TUYỂN TẬP 40 BÀI TẬP TRỌNG TÂM ÔN THI HSG CẤP TỈNH & QUỐC GIA (PYTHON)"
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
        "Đầy đủ 8 mảng kiến thức cốt lõi – Phân tích thuật toán & Code mẫu chuẩn Fast I/O"
    )
    srun.font.name = "Times New Roman"
    srun.font.size = Pt(12)
    srun.italic = True
    srun.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    for section_title, exercises in SECTIONS_DATA:
        add_part_heading(doc, section_title)
        for ex in exercises:
            add_sub_heading(doc, ex["title"])
            add_rich_paragraph(doc, f"**Đề bài:** {ex['desc']}")
            add_rich_paragraph(doc, ex["example"], is_bullet=True)
            add_analysis_box(doc, ex["analysis"])
            add_code_block(doc, ex["code"])

    doc.save(output_filename)
    print(
        f"THÀNH CÔNG! Đã tạo tệp giáo án: {os.path.abspath(output_filename)}"
    )


if __name__ == "__main__":
    OUTPUT_FILE = "40_BAI_TAP_HSG_TINH_QUOC_GIA.DOCX"
    try:
        build_document(OUTPUT_FILE)
    except PermissionError:
        print(
            f"LỖI: Không thể ghi đè lên tệp '{OUTPUT_FILE}' vì tệp đang được mở trong Word."
        )
        print("Vui lòng đóng tệp Word đó lại và chạy lại chương trình.")
    except Exception as err:
        print(f"Đã xảy ra lỗi không mong muốn: {err}")