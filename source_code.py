"""
source_code.py — Web Trợ Lý Dạy Học Lập Trình Python (Streamlit Edition)
Chuyển đổi từ index.html — 9 yêu cầu nâng cấp đã hoàn thành.

Cài đặt:
    pip install streamlit python-docx

Chạy:
    streamlit run source_code.py
"""
import streamlit as st
import json, os, csv, io, subprocess, sys, tempfile, time, random
from datetime import datetime
from pathlib import Path
import pandas as pd

try:
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    DOCX_OK = True
except ImportError:
    DOCX_OK = False

# ══════════════════════════════════════════════════════════════════════
# CẤU HÌNH
# ══════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Trợ Lý Dạy Học Python",
    page_icon="🐍",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
.stApp { background: #f8fafc; }
.main-header {
    background: linear-gradient(135deg,#1e293b 0%,#334155 60%,#1e3a8a 100%);
    color:#fff; padding:1rem 1.5rem; border-radius:14px; margin-bottom:1rem;
    box-shadow:0 8px 20px rgba(15,23,42,.25);
}
.main-header h1 { margin:0; font-size:1.35rem; font-weight:800; letter-spacing:-.02em; }
.main-header p  { margin:.25rem 0 0 0; color:#cbd5e1; font-size:.82rem; }
.badge-reset { background:#dc2626; color:#fff; padding:2px 10px; border-radius:6px;
    font-weight:800; font-size:.72rem; animation:pulse 1.4s infinite; }
@keyframes pulse { 0%,100%{opacity:1} 50%{opacity:.55} }
.lecture-card { background:#fff; border:1px solid #e2e8f0; border-left:4px solid #2563eb;
    border-radius:10px; padding:1rem; margin:.4rem 0; }
.lecture-card h4 { margin:.1rem 0 .4rem 0; color:#1e293b; }
.stTextArea textarea {
    font-family:'JetBrains Mono','Consolas',monospace !important;
    font-size:13px !important; background:#0f172a !important; color:#e2e8f0 !important;
    border-radius:8px !important;
}
div[data-testid="stTabs"] button { font-weight:600; }
.diff-easy { color:#059669; font-weight:700; }
.diff-medium { color:#d97706; font-weight:700; }
.diff-hard { color:#dc2626; font-weight:700; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════
# TẦNG DỮ LIỆU — JSON file storage
# ══════════════════════════════════════════════════════════════════════
DATA_FILE = Path("app_data.json")

DEFAULT_DATA = {
    "users": [
        {"id":1,"name":"Thầy Nguyễn Thế Phương","username":"admin","identifier":"admin",
         "password":"Admin#999","school":"THCS Thành Đông","className":"Toàn trường",
         "contact":"0908889999","role":"admin","loginCount":0},
        {"id":2,"name":"Giáo viên Dùng Thử","username":"giaovien","identifier":"giaovien",
         "password":"Giaovien@123","school":"THCS Thành Đông","className":"Khối 8",
         "contact":"0932861739","role":"teacher","loginCount":0},
        {"id":3,"name":"Học sinh Dùng Thử","username":"hocsinh","identifier":"hocsinh",
         "password":"Hocsinh@123","school":"THCS Thành Đông","className":"8A1",
         "contact":"---","role":"student","loginCount":0},
    ],
    "lectures": [
        {"id":1,"title":"Bài 1. Ngôn ngữ lập trình bậc cao và Python","classLevel":"8",
         "grade":"Lớp 8 - Tin học cơ sở",
         "content":"### 1. Ngôn ngữ bậc cao\n- **Ngôn ngữ máy**: dãy bit 0/1.\n- **Ngôn ngữ bậc cao**: gần gũi tự nhiên.\n\n### 2. Python\nDo **Guido van Rossum** phát triển năm 1991.\n\n```python\nprint('Xin chào Python!')\n```"},
        {"id":2,"title":"Bài 2. Biến và lệnh gán","classLevel":"8",
         "grade":"Lớp 8 - Tin học cơ sở",
         "content":"### Biến\nCú pháp: `tên_biến = giá_trị`\n\n```python\na = 10\nb = 3\ntong = a + b\n```"},
        {"id":3,"title":"Bài 3. Vào/ra dữ liệu","classLevel":"8",
         "grade":"Lớp 8 - Tin học cơ sở",
         "content":"- `print()`: xuất\n- `input()`: nhập\n\n```python\nho_ten = input('Họ tên: ')\nprint('Xin chào', ho_ten)\n```"},
        {"id":4,"title":"Bài 4. Cấu trúc rẽ nhánh","classLevel":"8",
         "grade":"Lớp 8 - Tin học cơ sở",
         "content":"```python\nif điều_kiện:\n    pass\nelif điều_kiện_2:\n    pass\nelse:\n    pass\n```"},
        {"id":5,"title":"Bài 5. Cấu trúc lặp","classLevel":"8",
         "grade":"Lớp 8 - Tin học cơ sở",
         "content":"```python\nfor i in range(1, 6):\n    print(i)\n\ns = 0\ni = 1\nwhile i <= 10:\n    s += i\n    i += 1\n```"},
        {"id":6,"title":"Bài 6. Kiểu danh sách","classLevel":"8",
         "grade":"Lớp 8 - Tin học cơ sở",
         "content":"```python\nds = [12, 5, 8]\nds.append(15)\nprint(len(ds))\n```"},
    ],
    "assignments": [
        {"id":1,"title":"Bài 1. Tính tổng A + B","classLevel":"8","difficulty":"easy",
         "description":"### Đề bài\nCho hai số nguyên $A$ và $B$. Tính $A + B$.\n\n### Dữ liệu vào\nHai số nguyên cách nhau dấu cách.\n\n### Dữ liệu ra\nIn tổng.",
         "templateCode":"a, b = map(int, input().split())\nprint(a + b)\n",
         "tests":[{"input":"3 5","expected":"8","point":5.0,"is_hidden":False},
                  {"input":"-10 25","expected":"15","point":5.0,"is_hidden":True}]},
        {"id":2,"title":"Bài 2. Chẵn lẻ","classLevel":"8","difficulty":"easy",
         "description":"### Đề bài\nNhập số nguyên N. In `CHAN` nếu chẵn, `LE` nếu lẻ.",
         "templateCode":"n = int(input())\nprint('CHAN' if n % 2 == 0 else 'LE')\n",
         "tests":[{"input":"4","expected":"CHAN","point":5.0,"is_hidden":False},
                  {"input":"7","expected":"LE","point":5.0,"is_hidden":True}]},
        {"id":3,"title":"Bài 3. Tổng dãy 1..N","classLevel":"8","difficulty":"medium",
         "description":"### Đề bài\nNhập N. Tính tổng 1 + 2 + ... + N.",
         "templateCode":"n = int(input())\nprint(n*(n+1)//2)\n",
         "tests":[{"input":"10","expected":"55","point":5.0,"is_hidden":False},
                  {"input":"100","expected":"5050","point":5.0,"is_hidden":True}]},
        {"id":4,"title":"Bài 4. Đếm ước","classLevel":"8","difficulty":"hard",
         "description":"### Đề bài\nNhập N. Đếm số ước nguyên dương của N.",
         "templateCode":"n = int(input())\nc = 0\nfor i in range(1, int(n**0.5)+1):\n    if n%i==0:\n        c+=1\n        if i!=n//i: c+=1\nprint(c)\n",
         "tests":[{"input":"12","expected":"6","point":5.0,"is_hidden":False},
                  {"input":"16","expected":"5","point":5.0,"is_hidden":True}]},
    ],
    "exams": [],
    "submissions": [],
    "logs": [],
    "flags": {},
    "gradebook": {},
}

def _load():
    if DATA_FILE.exists():
        try:
            return json.loads(DATA_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return json.loads(json.dumps(DEFAULT_DATA))

def _save(data):
    DATA_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def get_db():
    if "db" not in st.session_state:
        st.session_state.db = _load()
    return st.session_state.db

def save_db():
    _save(st.session_state.db)

# ══════════════════════════════════════════════════════════════════════
# PYTHON EXECUTION — subprocess an toàn
# ══════════════════════════════════════════════════════════════════════
def run_python(code: str, stdin_data: str = "", timeout: int = 5):
    if not code or not code.strip():
        return {"stdout":"", "stderr":"Không có code", "error":True}
    tmp = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".py", delete=False,
                                         encoding="utf-8") as f:
            f.write(code)
            tmp = f.name
        r = subprocess.run(
            [sys.executable, tmp],
            input=stdin_data or "",
            capture_output=True, text=True, timeout=timeout,
        )
        return {"stdout":r.stdout, "stderr":r.stderr, "error":r.returncode != 0}
    except subprocess.TimeoutExpired:
        return {"stdout":"", "stderr":f"⏱ Timeout: Quá {timeout} giây", "error":True}
    except Exception as e:
        return {"stdout":"", "stderr":str(e), "error":True}
    finally:
        if tmp:
            try: os.unlink(tmp)
            except Exception: pass

# ══════════════════════════════════════════════════════════════════════
# AUTH + LOGS
# ══════════════════════════════════════════════════════════════════════
def log_action(name, identifier, role, result):
    db = get_db()
    db["logs"].insert(0, {
        "time": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "name": name, "identifier": identifier, "role": role, "result": result,
    })
    save_db()

def do_login(username, password, want_role=None):
    db = get_db()
    uname = (username or "").strip().lower()
    for u in db["users"]:
        if (u.get("username","").lower() == uname or
            u.get("identifier","").lower() == uname or
            str(u.get("contact","")).lower() == uname):
            if want_role == "teacher" and u["role"] == "student":
                continue
            if want_role == "admin" and u["role"] != "admin":
                continue
            if u["password"] == password:
                u["loginCount"] = u.get("loginCount",0) + 1
                log_action(u["name"], u.get("identifier",uname), u["role"], "Thành công")
                return u
    return None

def do_logout():
    u = st.session_state.get("user")
    if u:
        log_action(u["name"], u.get("identifier",""), u["role"], "Đăng xuất")
        save_db()
    st.session_state.logged_in = False
    st.session_state.user = None

def do_change_password(username, contact, new_pw, confirm):
    db = get_db()
    if new_pw != confirm:
        return False, "Mật khẩu nhập lại không khớp!"
    if len(new_pw) < 6:
        return False, "Mật khẩu mới phải ≥ 6 ký tự!"
    uname = (username or "").strip().lower()
    for u in db["users"]:
        if (u.get("username","").lower() == uname or
            u.get("identifier","").lower() == uname):
            contact_norm = str(contact or "").strip().lower()
            stored = str(u.get("contact","")).strip().lower()
            if stored and stored == contact_norm:
                u["password"] = new_pw
                log_action(u["name"], u.get("identifier",""), u["role"], "Đổi MK thành công")
                save_db()
                return True, "✅ Đổi mật khẩu thành công! Hãy đăng nhập lại."
            else:
                log_action(u["name"], u.get("identifier",""), u["role"], "Reset lại")
                save_db()
                return False, ("🔴 Mật khẩu của bạn sẽ được khôi phục sau 24 giờ "
                               "hoặc bạn nên đăng ký tài khoản khác!")
    return False, "Không tìm thấy tài khoản!"

# ══════════════════════════════════════════════════════════════════════
# UTILS
# ══════════════════════════════════════════════════════════════════════
def next_id(items):
    return max([i.get("id",0) for i in items], default=0) + 1

def fmt_diff(d):
    return {"easy":"🟢 Dễ", "medium":"🟡 TB", "hard":"🔴 Khó"}.get(d, "🟢 Dễ")

def grade_submission(assignment, code):
    total, max_score, details = 0.0, 0.0, []
    for i, t in enumerate(assignment.get("tests", []), 1):
        pt = float(t.get("point", 5.0))
        max_score += pt
        res = run_python(code, t.get("input",""))
        out = (res["stdout"] or "").strip()
        exp = str(t.get("expected","")).strip()
        ok = (not res["error"]) and out == exp
        if ok: total += pt
        details.append({"test":i,"input":t.get("input",""),"expected":exp,
                        "got":out,"passed":ok,"point":pt,"stderr":res["stderr"]})
    return total, max_score, details

def record_submission(user, title, aid, score, max_score, exam_id=None):
    db = get_db()
    db["submissions"].insert(0, {
        "id": int(time.time()*1000) + random.randint(0,999),
        "student": user["name"],
        "username": user.get("identifier", user.get("username","")),
        "assignmentTitle": title,
        "assignmentId": aid,
        "examId": exam_id,
        "score": score,
        "maxScore": max_score,
        "timestamp": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    })
    save_db()

def build_gradebook():
    db = get_db()
    agg = {}
    for s in db.get("submissions", []):
        key = s.get("username") or s.get("student")
        if not key: continue
        if key not in agg:
            agg[key] = {"student":s["student"], "username":s.get("username",""),
                        "attempts":0, "totalScore":0.0, "highestScore":0.0,
                        "lastSubmit":""}
        g = agg[key]
        g["attempts"] += 1
        g["totalScore"] += float(s.get("score",0))
        g["highestScore"] = max(g["highestScore"], float(s.get("score",0)))
        if not g["lastSubmit"] or s["timestamp"] > g["lastSubmit"]:
            g["lastSubmit"] = s["timestamp"]
    return list(agg.values())

# ══════════════════════════════════════════════════════════════════════
# EXPORT
# ══════════════════════════════════════════════════════════════════════
def export_csv_bytes(rows, headers):
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=headers)
    w.writeheader()
    w.writerows(rows)
    return ("\ufeff" + buf.getvalue()).encode("utf-8")

def export_assignments_json_bytes():
    db = get_db()
    return json.dumps({"assignments": db.get("assignments", []),
                       "exportedAt": datetime.now().isoformat()},
                      ensure_ascii=False, indent=2).encode("utf-8")

def export_lectures_json_bytes():
    db = get_db()
    return json.dumps({"lectures": db.get("lectures", []),
                       "exportedAt": datetime.now().isoformat()},
                      ensure_ascii=False, indent=2).encode("utf-8")

def export_all_json_bytes():
    return json.dumps(get_db(), ensure_ascii=False, indent=2).encode("utf-8")

def export_random_test_docx(questions, title, duration, show_sample):
    """Word với công thức đơn giản (MathType đầy đủ cần OMML — xem bản HTML)."""
    if not DOCX_OK:
        return None
    doc = Document()
    h = doc.add_heading(title or "Đề kiểm tra", level=0)
    h.alignment = 1
    p = doc.add_paragraph(f"Thời gian: {duration} phút — Số câu: {len(questions)}")
    p.alignment = 1
    doc.add_paragraph(
        "Học sinh " + ("ĐƯỢC PHÉP" if show_sample else "KHÔNG ĐƯỢC PHÉP") +
        " xem code mẫu khi làm bài."
    )
    for i, a in enumerate(questions, 1):
        doc.add_heading(f"Câu {i}. {a['title']}", level=2)
        doc.add_paragraph(a.get("description",""))
        for k, t in enumerate([x for x in a.get("tests",[]) if not x.get("is_hidden")], 1):
            doc.add_paragraph(f"Ví dụ {k}: Input: {t.get('input','')} → Output: {t.get('expected','')}")
        if show_sample and a.get("templateCode"):
            doc.add_paragraph("Mã nguồn mẫu:")
            for line in a["templateCode"].split("\n"):
                par = doc.add_paragraph(line)
                par.paragraph_format.left_indent = Cm(1)
                if par.runs:
                    par.runs[0].font.name = "Consolas"
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

# ══════════════════════════════════════════════════════════════════════
# SIDEBAR / HEADER
# ══════════════════════════════════════════════════════════════════════
def render_header():
    st.markdown("""
    <div class="main-header">
        <h1>🐍 WEB TRỢ LÝ DẠY HỌC LẬP TRÌNH PYTHON</h1>
        <p>Nền tảng thực hành lập trình &amp; Chấm điểm tự động · THCS Thành Đông</p>
    </div>
    """, unsafe_allow_html=True)

def render_sidebar():
    db = get_db()
    with st.sidebar:
        st.markdown("### 👤 Tài khoản")
        if st.session_state.logged_in:
            u = st.session_state.user
            role_lbl = {"admin":"🛡 Admin", "teacher":"👨‍🏫 Giáo viên", "student":"👨‍🎓 Học sinh"}[u["role"]]
            st.success(f"**{u['name']}**\n\n{role_lbl}")
            if u["role"] == "student":
                st.caption(f"Lớp: {u.get('className','—')}")

            # Change password
            with st.expander("🔑 Đổi mật khẩu"):
                with st.form("cp_form"):
                    cp_user = st.text_input("Tên đăng nhập", value=u.get("username",""))
                    cp_contact = st.text_input("Phương thức liên lạc (SĐT/Email/CCCD)")
                    cp_new = st.text_input("Mật khẩu mới", type="password")
                    cp_conf = st.text_input("Nhập lại mật khẩu", type="password")
                    if st.form_submit_button("Xác nhận đổi", use_container_width=True):
                        ok, msg = do_change_password(cp_user, cp_contact, cp_new, cp_conf)
                        if ok: st.success(msg)
                        else: st.error(msg)

            if st.button("🚪 Đăng xuất", use_container_width=True, type="primary"):
                do_logout()
                st.rerun()

        else:
            st.info("Vui lòng đăng nhập")
        st.divider()
        st.markdown("### ⚙️ Thông tin")
        st.caption("📞 0932.861.739")
        st.caption("💌 thefuong119@gmail.com")
        st.caption("© 2026 Nguyễn Thế Phương")

# ══════════════════════════════════════════════════════════════════════
# LOGIN PAGE
# ══════════════════════════════════════════════════════════════════════
def page_login():
    render_header()
    col1, col2 = st.columns([1.2, 1])

    with col1:
        st.markdown("### 🔐 Đăng nhập")
        with st.form("login_form"):
            role_tab = st.radio(
                "Vai trò", ["👨‍🎓 Học sinh", "👨‍🏫 Giáo viên", "🛡 Admin"],
                horizontal=True
            )
            username = st.text_input("Tên đăng nhập / SĐT / Email / CCCD")
            password = st.text_input("Mật khẩu", type="password")
            submitted = st.form_submit_button("🚀 Đăng nhập", use_container_width=True, type="primary")
        if submitted:
            want = {"👨‍🎓 Học sinh":"student",
                    "👨‍🏫 Giáo viên":"teacher",
                    "🛡 Admin":"admin"}[role_tab]
            u = do_login(username, password, want_role=(None if want=="student" else want))
            if u:
                st.session_state.logged_in = True
                st.session_state.user = u
                st.success(f"Chào mừng {u['name']}!")
                time.sleep(0.4)
                st.rerun()
            else:
                st.error("Tài khoản hoặc mật khẩu không chính xác!")

        with st.expander("📝 Đăng ký học sinh mới"):
            with st.form("reg_student"):
                n_name = st.text_input("Họ và tên *")
                n_school = st.text_input("Trường", "THCS Thành Đông")
                n_class = st.text_input("Lớp *", "8A1")
                n_user = st.text_input("Tên đăng nhập *")
                n_contact = st.text_input("Phương thức liên lạc (SĐT/Email/CCCD)")
                n_pass = st.text_input("Mật khẩu *", type="password")
                if st.form_submit_button("Đăng ký", use_container_width=True):
                    if not (n_name and n_class and n_user and n_pass):
                        st.error("Vui lòng điền đủ thông tin bắt buộc!")
                    else:
                        db = get_db()
                        exists = any(u.get("username","").lower() == n_user.lower()
                                     for u in db["users"])
                        if exists:
                            st.error("Tên đăng nhập đã tồn tại!")
                        else:
                            db["users"].append({
                                "id": next_id(db["users"]),
                                "name": n_name, "username": n_user.lower(),
                                "identifier": n_user.lower(), "password": n_pass,
                                "school": n_school, "className": n_class,
                                "contact": n_contact or "---",
                                "role":"student", "loginCount":1,
                            })
                            save_db()
                            log_action(n_name, n_user, "student", "Đăng ký thành công")
                            st.success(f"Đăng ký thành công **{n_name}**! Hãy đăng nhập.")

    with col2:
        st.markdown("### 🎓 Tài khoản dùng thử")
        st.info("**Giáo viên:** `giaovien` / `Giaovien@123`")
        st.info("**Học sinh:** `hocsinh` / `Hocsinh@123`")
        st.info("**Admin:** `admin` / `Admin#999`")
        st.markdown("### ✨ Tính năng")
        for f in ["Trình thông dịch Python thật (subprocess)",
                  "Chấm điểm tự động với test ẩn",
                  "Bài giảng đa khối lớp",
                  "Tạo đề ngẫu nhiên theo độ khó",
                  "Bài kiểm tra trực tuyến",
                  "Sổ điểm & xếp hạng",
                  "Nhật ký & Reset mật khẩu"]:
            st.markdown(f"- {f}")

# ══════════════════════════════════════════════════════════════════════
# STUDENT VIEW
# ══════════════════════════════════════════════════════════════════════
def student_lectures():
    db = get_db()
    st.markdown("#### 📚 Thư viện bài giảng")
    c1, c2 = st.columns([1, 2])
    with c1:
        grade_filter = st.selectbox("Khối", ["all","6","7","8","9","10","11","12"],
                                    index=3, key="lec_grade")
        search = st.text_input("🔍 Tìm bài giảng", key="lec_search")
    lecs = [l for l in db["lectures"]
            if (grade_filter == "all" or l.get("classLevel","8") == grade_filter)
            and (not search or search.lower() in l["title"].lower())]
    if not lecs:
        st.info("Không có bài giảng phù hợp.")
        return
    titles = [f"[Khối {l.get('classLevel','8')}] {l['title']}" for l in lecs]
    idx = st.selectbox("Chọn bài", range(len(titles)),
                       format_func=lambda i: titles[i], key="lec_pick")
    lec = lecs[idx]
    st.markdown(f"### {lec['title']}")
    st.caption(f"Chủ đề: {lec.get('grade','—')} · Khối {lec.get('classLevel','8')}")
    st.markdown(lec["content"])
    if st.button("🔊 Đọc to (TTS hệ thống)", key="lec_tts"):
        st.info("💡 Tính năng TTS cần thư viện ngoài (pyttsx3/gTTS) không có sẵn trong Streamlit. "
                "Hãy đọc trực tiếp nội dung trên.")

def student_practise():
    db = get_db()
    user = st.session_state.user
    st.markdown("#### 💻 Thực hành lập trình")

    c1, c2, c3 = st.columns([2, 1, 1])
    with c1:
        grade_f = st.selectbox("Khối", ["all","6","7","8","9","10","11","12"],
                               index=3, key="pr_grade")
    with c2:
        diff_f = st.selectbox("Độ khó", ["all","easy","medium","hard"],
                              format_func=lambda x: {"all":"Mọi độ khó",
                                                     "easy":"🟢 Dễ",
                                                     "medium":"🟡 TB",
                                                     "hard":"🔴 Khó"}[x],
                              key="pr_diff")
    with c3:
        search = st.text_input("🔍", key="pr_search")

    asts = [a for a in db["assignments"]
            if (grade_f == "all" or a.get("classLevel","8") == grade_f)
            and (diff_f == "all" or a.get("difficulty") == diff_f)
            and (not search or search.lower() in a["title"].lower())]
    if not asts:
        st.warning("Không có bài tập phù hợp.")
        return

    labels = [f"{fmt_diff(a.get('difficulty'))} · {a['title']}" for a in asts]
    idx = st.selectbox("Chọn bài tập", range(len(labels)),
                       format_func=lambda i: labels[i], key="pr_pick")
    a = asts[idx]

    # Show description
    with st.expander("📋 Đề bài", expanded=True):
        st.markdown(a.get("description",""))
        st.markdown("**Ví dụ công khai:**")
        for t in [x for x in a.get("tests",[]) if not x.get("is_hidden")]:
            st.code(f"Input: {t.get('input','')}\nOutput: {t.get('expected','')}", language="text")

    # Code editor
    st.markdown("##### ✍️ Viết code của em")
    code_key = f"code_{a['id']}"
    if code_key not in st.session_state:
        st.session_state[code_key] = "# Viết mã Python của em ở đây...\n"
    code = st.text_area("Code", key=code_key, height=280, label_visibility="collapsed")

    # Actions
    c_run, c_grade, c_unlock = st.columns([1, 1, 1])
    stdin_val = st.text_area("📥 Dữ liệu vào (input.txt / stdin)",
                             height=80, key=f"stdin_{a['id']}",
                             placeholder="Nhập dữ liệu thử nghiệm...")

    with c_run:
        if st.button("▶️ Chạy thử", use_container_width=True):
            if not stdin_val.strip() and any(t.get("input") for t in a.get("tests",[])):
                st.warning("⚠️ Em chưa nhập 'Dữ liệu vào' — kết quả có thể không đúng!")
            r = run_python(code, stdin_val)
            st.markdown("**Kết quả:**")
            if r["error"]:
                st.error(r["stderr"] or r["stdout"])
            else:
                st.code(r["stdout"] or "(không có output)", language="text")

    with c_grade:
        if st.button("📝 Nộp bài chấm điểm", use_container_width=True, type="primary"):
            if not code.strip():
                st.error("Em chưa viết code!")
            else:
                total, mx, details = grade_submission(a, code)
                st.success(f"Điểm: **{total:.1f} / {mx:.1f}**")
                for d in details:
                    if d["passed"]:
                        st.success(f"✅ Test {d['test']}: +{d['point']}đ")
                    else:
                        st.error(f"❌ Test {d['test']}: 0đ (kỳ vọng '{d['expected']}', nhận '{d['got']}')")
                record_submission(user, a["title"], a["id"], total, mx)
                # Flag
                uid = user.get("identifier")
                db["flags"].setdefault(uid, {})[str(a["id"])] = True
                save_db()

    with c_unlock:
        if st.button("🔓 Xem code mẫu", use_container_width=True):
            if a.get("templateCode"):
                st.session_state[code_key] = a["templateCode"]
                st.rerun()

def student_exams():
    db = get_db()
    user = st.session_state.user
    st.markdown("#### 📝 Bài kiểm tra trực tuyến")
    exams = db.get("exams", [])
    stu_grade = (user.get("className") or "")
    import re
    m = re.search(r"\d+", stu_grade)
    stu_num = m.group(0) if m else ""
    visible = [e for e in exams if e.get("classLevel","all") == "all"
               or not e.get("classLevel") or e["classLevel"] == stu_num]
    if not visible:
        st.info("Chưa có bài kiểm tra nào. Vui lòng chờ giáo viên tạo đề.")
        return

    labels = [f"{e['title']} ({len(e['assignmentIds'])} câu — {e['duration']}p)"
              for e in visible]
    idx = st.selectbox("Chọn bài kiểm tra", range(len(labels)),
                       format_func=lambda i: labels[i], key="exam_pick")
    exam = visible[idx]
    st.markdown(f"### {exam['title']}")
    st.caption(f"⏱ {exam['duration']} phút · Tạo bởi {exam.get('createdBy','—')}")
    if not exam.get("showSampleCode", True):
        st.warning("🔒 **KHÔNG được xem code mẫu** trong bài kiểm tra này.")

    a_ids = exam["assignmentIds"]
    a_idx = st.selectbox(
        "Chọn câu hỏi",
        range(len(a_ids)),
        format_func=lambda i: f"Câu {i+1}: " + next((a["title"] for a in db["assignments"] if a["id"] == a_ids[i]), "?"),
        key="exam_q"
    )
    aid = a_ids[a_idx]
    a = next((x for x in db["assignments"] if x["id"] == aid), None)
    if not a: return

    with st.expander("📋 Đề bài câu này", expanded=True):
        st.markdown(a.get("description",""))
        for t in [x for x in a.get("tests",[]) if not x.get("is_hidden")]:
            st.code(f"In: {t.get('input','')}\nOut: {t.get('expected','')}", language="text")

    ck = f"exam_code_{exam['id']}_{aid}"
    if ck not in st.session_state:
        st.session_state[ck] = "# Viết code...\n"
    code = st.text_area("Code", key=ck, height=250, label_visibility="collapsed")

    stdin_v = st.text_area("Dữ liệu vào", height=60, key=f"exam_stdin_{aid}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("▶️ Chạy thử", use_container_width=True, key=f"exrun_{aid}"):
            r = run_python(code, stdin_v)
            if r["error"]: st.error(r["stderr"] or r["stdout"])
            else: st.code(r["stdout"] or "(trống)", language="text")
    with c2:
        if st.button("📤 Nộp câu này", use_container_width=True, type="primary", key=f"exsub_{aid}"):
            total, mx, details = grade_submission(a, code)
            st.success(f"Điểm câu này: **{total:.1f} / {mx:.1f}**")
            for d in details:
                if not d["passed"]:
                    st.caption(f"Test {d['test']}: kỳ vọng '{d['expected']}', nhận '{d['got']}'")
            record_submission(user, f"[KT: {exam['title']}] {a['title']}", aid,
                              total, mx, exam_id=exam["id"])

def page_student():
    tabs = st.tabs(["📖 Bài giảng", "💻 Thực hành", "📝 Bài kiểm tra"])
    with tabs[0]: student_lectures()
    with tabs[1]: student_practise()
    with tabs[2]: student_exams()

# ══════════════════════════════════════════════════════════════════════
# TEACHER VIEW
# ══════════════════════════════════════════════════════════════════════
def teacher_lectures():
    db = get_db()
    st.markdown("#### 📚 Quản lý bài giảng")
    df = pd.DataFrame([{
        "ID": l["id"], "Tiêu đề": l["title"], "Khối": l.get("classLevel","8"),
        "Chủ đề": l.get("grade",""),
    } for l in db["lectures"]])
    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.expander("➕ Thêm bài giảng mới"):
        with st.form("new_lecture"):
            t = st.text_input("Tiêu đề *")
            cl = st.selectbox("Khối", ["6","7","8","9","10","11","12"], index=2)
            g = st.text_input("Chủ đề", "Tin học cơ sở")
            body = st.text_area("Nội dung (Markdown)", height=200)
            if st.form_submit_button("Lưu bài giảng"):
                if not t or not body:
                    st.error("Vui lòng nhập tiêu đề và nội dung!")
                else:
                    db["lectures"].append({
                        "id": next_id(db["lectures"]), "title": t,
                        "classLevel": cl, "grade": g, "content": body,
                    })
                    save_db()
                    st.success("Đã thêm!")
                    st.rerun()

    with st.expander("✏️ Sửa / 🗑 Xóa bài giảng"):
        if db["lectures"]:
            lids = [l["id"] for l in db["lectures"]]
            lid = st.selectbox("Chọn ID bài giảng", lids, key="t_lecedit")
            lec = next(x for x in db["lectures"] if x["id"] == lid)
            e_t = st.text_input("Tiêu đề", lec["title"], key=f"t_lt{lid}")
            e_cl = st.selectbox("Khối", ["6","7","8","9","10","11","12"],
                                index=["6","7","8","9","10","11","12"].index(lec.get("classLevel","8")),
                                key=f"t_lc{lid}")
            e_g = st.text_input("Chủ đề", lec.get("grade",""), key=f"t_lg{lid}")
            e_body = st.text_area("Nội dung", lec["content"], height=150, key=f"t_lb{lid}")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("💾 Lưu thay đổi", use_container_width=True):
                    lec.update({"title":e_t, "classLevel":e_cl, "grade":e_g, "content":e_body})
                    save_db(); st.success("Đã lưu!"); st.rerun()
            with c2:
                if st.button("🗑 Xóa", use_container_width=True, type="secondary", disabled=True):
                    st.caption("Giáo viên không có quyền xóa — liên hệ Admin.")

def teacher_assignments():
    db = get_db()
    st.markdown("#### 🧩 Ngân hàng đề bài")
    df = pd.DataFrame([{
        "ID": a["id"], "Bài tập": a["title"],
        "Khối": a.get("classLevel","8"),
        "Độ khó": {"easy":"Dễ","medium":"TB","hard":"Khó"}.get(a.get("difficulty"),"?"),
        "Số test": len(a.get("tests",[])),
    } for a in db["assignments"]])
    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.expander("➕ Thêm / ✏️ Sửa bài tập"):
        new_or_edit = st.radio("Hành động", ["Tạo mới","Sửa bài có sẵn"], horizontal=True)
        if new_or_edit == "Sửa bài có sẵn" and db["assignments"]:
            edit_id = st.selectbox("Chọn bài tập",
                [a["id"] for a in db["assignments"]],
                format_func=lambda i: next(f"{x['id']}. {x['title']}" for x in db["assignments"] if x["id"]==i))
            a = next(x for x in db["assignments"] if x["id"] == edit_id)
        else:
            a = {"id":0,"title":"","classLevel":"8","difficulty":"easy",
                 "description":"","templateCode":"","tests":[]}

        with st.form("assign_form"):
            title = st.text_input("Tên bài *", a["title"])
            c1, c2 = st.columns(2)
            with c1:
                cl = st.selectbox("Khối", ["6","7","8","9","10","11","12"],
                                  index=["6","7","8","9","10","11","12"].index(a.get("classLevel","8")))
            with c2:
                diff = st.selectbox("Độ khó", ["easy","medium","hard"],
                                    index=["easy","medium","hard"].index(a.get("difficulty","easy")))
            desc = st.text_area("Đề bài", a["description"], height=150)
            tmpl = st.text_area("Code mẫu", a["templateCode"], height=150)
            st.markdown("**Bộ test (mỗi dòng: input|expected|điểm|ẩn)**")
            tests_str = st.text_area(
                "Test cases (mỗi dòng một test, định dạng: input|expected|point|hidden)",
                "\n".join(f"{t.get('input','')}|{t.get('expected','')}|{t.get('point',5.0)}|{'1' if t.get('is_hidden') else '0'}" for t in a["tests"]),
                height=120,
                help="hidden=1 nếu test ẩn, 0 nếu công khai",
            )
            if st.form_submit_button("💾 Lưu bài tập", type="primary"):
                if not title:
                    st.error("Vui lòng nhập tên!")
                else:
                    tests = []
                    for i, line in enumerate(tests_str.split("\n"), 1):
                        if not line.strip(): continue
                        parts = line.split("|")
                        if len(parts) < 2: continue
                        tests.append({
                            "input": parts[0].strip(),
                            "expected": parts[1].strip(),
                            "point": float(parts[2]) if len(parts) > 2 and parts[2].strip() else 5.0,
                            "is_hidden": (len(parts) > 3 and parts[3].strip() == "1"),
                        })
                    obj = {"title":title, "classLevel":cl, "difficulty":diff,
                           "description":desc, "templateCode":tmpl, "tests":tests}
                    if a["id"]:
                        for x in db["assignments"]:
                            if x["id"] == a["id"]: x.update(obj); break
                        st.success("Đã cập nhật!")
                    else:
                        obj["id"] = next_id(db["assignments"])
                        db["assignments"].append(obj)
                        st.success("Đã thêm!")
                    save_db(); st.rerun()

def teacher_submissions():
    db = get_db()
    st.markdown("#### 📄 Bài làm học sinh")
    subs = db.get("submissions", [])
    if not subs:
        st.info("Chưa có bài nộp nào.")
        return
    df = pd.DataFrame(subs)[["id","student","assignmentTitle","score","maxScore","timestamp"]]
    df.columns = ["ID","Học sinh","Bài tập","Điểm","Tối đa","Thời gian"]
    df_edit = st.data_editor(
        df, use_container_width=True, hide_index=True,
        column_config={"Chọn": st.column_config.CheckboxColumn(default=False)},
        num_rows="fixed",
    )
    st.caption("💡 Bản Streamlit: xóa bằng cách chọn dòng bên dưới rồi bấm nút.")

    # Manual multi-select via multiselect
    ids = [s["id"] for s in subs]
    sel = st.multiselect(
        "Chọn ID bài nộp để xóa:",
        ids,
        format_func=lambda i: next(f"#{i} — {s['student']} — {s['assignmentTitle']}"
                                   for s in subs if s["id"]==i),
        key="bulk_sub_sel"
    )
    if st.button(f"🗑 Xóa {len(sel)} bài đã chọn", type="secondary", disabled=not sel):
        db["submissions"] = [s for s in subs if s["id"] not in sel]
        save_db()
        st.success(f"Đã xóa {len(sel)} bài!")
        st.rerun()

def teacher_logs():
    db = get_db()
    st.markdown("#### 📋 Nhật ký truy cập")
    logs = db.get("logs", [])
    if not logs:
        st.info("Chưa có nhật ký.")
        return
    for l in logs:
        is_reset = "reset" in str(l.get("result","")).lower()
        prefix = "🔴" if is_reset else "✅"
        badge = '<span class="badge-reset">RESET LẠI</span>' if is_reset else l.get("result","")
        st.markdown(
            f"{prefix} **{l['time']}** — **{l['name']}** (`{l['identifier']}`) — "
            f"*{l['role']}* — {badge}",
            unsafe_allow_html=True,
        )

def teacher_grades():
    db = get_db()
    st.markdown("#### 📊 Sổ điểm lớp học")
    gb = build_gradebook()
    if not gb:
        st.info("Chưa có dữ liệu điểm.")
        return
    df = pd.DataFrame([{
        "Học sinh":g["student"], "Tài khoản":g["username"],
        "Số bài":g["attempts"], "Điểm cao nhất":round(g["highestScore"],2),
        "Tổng điểm":round(g["totalScore"],2),
        "Lần cuối":g["lastSubmit"],
    } for g in gb])
    st.dataframe(df, use_container_width=True, hide_index=True)
    csv_bytes = df.to_csv(index=False).encode("utf-8-sig")
    st.download_button("📥 Xuất CSV", csv_bytes, "bang_diem.csv", "text/csv")

def teacher_random_test():
    db = get_db()
    st.markdown("#### 🎲 Tạo đề ngẫu nhiên")
    with st.form("rand_form"):
        c1, c2, c3, c4 = st.columns(4)
        with c1: n_easy = st.number_input("Câu dễ", 0, 20, 2)
        with c2: n_med = st.number_input("Câu TB", 0, 20, 1)
        with c3: n_hard = st.number_input("Câu khó", 0, 20, 1)
        with c4: rg = st.selectbox("Khối", ["all","6","7","8","9","10","11","12"], index=3)
        rtitle = st.text_input("Tên bài kiểm tra",
                               f"Kiểm tra ngẫu nhiên {datetime.now().strftime('%d/%m/%Y')}")
        show_sample = st.checkbox("Cho học sinh xem code mẫu", value=True)
        duration = st.number_input("Thời gian (phút)", 5, 180, 45)
        generate = st.form_submit_button("🎲 Trộn đề", type="primary")

    if generate:
        pool = [a for a in db["assignments"]
                if rg == "all" or a.get("classLevel","8") == rg]
        def pick(diff, n):
            c = [a for a in pool if a.get("difficulty") == diff]
            random.shuffle(c)
            return c[:n]
        picked = pick("easy", n_easy) + pick("medium", n_med) + pick("hard", n_hard)
        if not picked:
            st.error("Không có bài tập phù hợp!")
        else:
            st.session_state["random_pick"] = {
                "title": rtitle, "classLevel": rg,
                "showSampleCode": show_sample, "duration": duration,
                "questions": picked,
            }
            st.success(f"Đã trộn {len(picked)} câu!")

    if "random_pick" in st.session_state:
        rp = st.session_state["random_pick"]
        st.markdown(f"**{rp['title']}** — {len(rp['questions'])} câu — {rp['duration']} phút")
        for i, a in enumerate(rp["questions"], 1):
            st.markdown(f"{i}. **{a['title']}** ({fmt_diff(a.get('difficulty'))})")

        c1, c2, c3 = st.columns(3)
        with c1:
            if st.button("💾 Lưu & đẩy cho HS", type="primary", use_container_width=True):
                db["exams"].append({
                    "id": next_id(db["exams"]),
                    "title": rp["title"],
                    "classLevel": rp["classLevel"],
                    "showSampleCode": rp["showSampleCode"],
                    "duration": rp["duration"],
                    "createdAt": datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "createdBy": st.session_state.user["name"],
                    "assignmentIds": [a["id"] for a in rp["questions"]],
                })
                save_db()
                st.success("Đã lưu bài kiểm tra!")
                del st.session_state["random_pick"]
                st.rerun()
        with c2:
            if DOCX_OK:
                docx_bytes = export_random_test_docx(
                    rp["questions"], rp["title"], rp["duration"], rp["showSampleCode"]
                )
                if docx_bytes:
                    st.download_button(
                        "📥 Xuất Word", docx_bytes, "de_kiem_tra.docx",
                        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
            else:
                st.caption("⚠️ Cài `python-docx` để xuất Word")
        with c3:
            if st.button("🗑 Hủy đề", use_container_width=True):
                del st.session_state["random_pick"]
                st.rerun()

def page_teacher():
    tabs = st.tabs(["📚 Bài giảng","🧩 Đề bài","📄 Bài nộp",
                    "📋 Nhật ký","📊 Sổ điểm","🎲 Tạo đề"])
    with tabs[0]: teacher_lectures()
    with tabs[1]: teacher_assignments()
    with tabs[2]: teacher_submissions()
    with tabs[3]: teacher_logs()
    with tabs[4]: teacher_grades()
    with tabs[5]: teacher_random_test()

# ══════════════════════════════════════════════════════════════════════
# ADMIN VIEW
# ══════════════════════════════════════════════════════════════════════
def admin_users():
    db = get_db()
    st.markdown("#### 👥 Quản lý tài khoản")
    c1, c2 = st.columns(2)
    with c1:
        grp = st.radio("Nhóm", ["Học sinh","Giáo viên + Admin"], horizontal=True,
                       key="adm_grp")
    users_f = [u for u in db["users"]
               if (grp == "Học sinh" and u["role"] == "student") or
                  (grp != "Học sinh" and u["role"] != "student")]
    df = pd.DataFrame([{
        "ID":u["id"], "Họ tên":u["name"],
        "Trường":u.get("school",""), "Lớp":u.get("className",""),
        "Liên lạc":u.get("contact",""),
        "Tên ĐN":u.get("username",""),
        "Vai trò":u["role"], "Lượt ĐN":u.get("loginCount",0),
    } for u in users_f])
    st.dataframe(df, use_container_width=True, hide_index=True)

    with st.expander("➕ Thêm tài khoản"):
        with st.form("add_user"):
            c1, c2 = st.columns(2)
            with c1:
                nn = st.text_input("Họ tên *")
                nu = st.text_input("Tên đăng nhập *")
                np = st.text_input("Mật khẩu *", value="User#123")
            with c2:
                nc = st.text_input("Phương thức liên lạc (SĐT/Email/CCCD)")
                nr = st.selectbox("Vai trò", ["student","teacher","admin"])
                nclass = st.text_input("Lớp / Khối", "8A1")
            if st.form_submit_button("Thêm", type="primary"):
                if not (nn and nu and np):
                    st.error("Vui lòng nhập đủ thông tin bắt buộc!")
                else:
                    db["users"].append({
                        "id": next_id(db["users"]),
                        "name": nn, "username": nu.lower(),
                        "identifier": nu.lower(), "password": np,
                        "school": "THCS Thành Đông", "className": nclass,
                        "contact": nc or "---", "role": nr, "loginCount": 0,
                    })
                    save_db(); st.success("Đã thêm!"); st.rerun()

    with st.expander("✏️ Sửa / 🗑 Xóa tài khoản"):
        if users_f:
            uid = st.selectbox("ID tài khoản",
                [u["id"] for u in users_f],
                format_func=lambda i: next(f"{u['id']}. {u['name']} ({u['username']})"
                                           for u in users_f if u["id"]==i),
                key="adm_uid")
            u = next(x for x in users_f if x["id"] == uid)
            en = st.text_input("Họ tên", u["name"], key=f"ae_n{uid}")
            eu = st.text_input("Username", u.get("username",""), key=f"ae_u{uid}")
            ep = st.text_input("Mật khẩu", u["password"], key=f"ae_p{uid}")
            ec = st.text_input("Liên lạc", u.get("contact",""), key=f"ae_c{uid}")
            er = st.selectbox("Vai trò", ["student","teacher","admin"],
                              index=["student","teacher","admin"].index(u["role"]),
                              key=f"ae_r{uid}")
            c1, c2 = st.columns(2)
            with c1:
                if st.button("💾 Lưu", use_container_width=True):
                    u.update({"name":en, "username":eu.lower(),
                              "identifier":eu.lower(), "password":ep,
                              "contact":ec, "role":er})
                    save_db(); st.success("Đã lưu!"); st.rerun()
            with c2:
                if st.button("🗑 Xóa", use_container_width=True, type="secondary",
                             disabled=(u["role"]=="admin")):
                    db["users"] = [x for x in db["users"] if x["id"] != uid]
                    save_db(); st.success("Đã xóa!"); st.rerun()

def admin_grades():
    db = get_db()
    st.markdown("#### 🏆 Bảng xếp hạng điểm")
    gb = build_gradebook()
    if not gb:
        st.info("Chưa có dữ liệu điểm.")
        return
    c1, c2 = st.columns([1, 3])
    with c1:
        sort_desc = st.radio("Sắp xếp", ["Giảm dần (mặc định)","Tăng dần"],
                             index=0, key="adm_sort")
    gb.sort(key=lambda g: g["totalScore"], reverse=(sort_desc == "Giảm dần (mặc định)"))
    medals = ["🥇","🥈","🥉"]
    rows = []
    for i, g in enumerate(gb):
        rows.append({
            "Hạng": medals[i] if i < 3 else str(i+1),
            "Học sinh": g["student"], "Tài khoản": g["username"],
            "Số bài": g["attempts"],
            "Cao nhất": round(g["highestScore"],2),
            "TỔNG ĐIỂM": round(g["totalScore"],2),
            "Lần cuối": g["lastSubmit"],
        })
    df = pd.DataFrame(rows)
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.download_button("📥 Xuất CSV", df.to_csv(index=False).encode("utf-8-sig"),
                       "bxh_diem.csv", "text/csv")

def admin_submissions():
    teacher_submissions()  # reuse with admin capability
    db = get_db()
    if st.button("🗑 Xóa TOÀN BỘ bài nộp", type="secondary"):
        db["submissions"] = []; save_db()
        st.success("Đã xóa toàn bộ!"); st.rerun()

def admin_import_export():
    db = get_db()
    st.markdown("#### 💾 Nhập / Xuất dữ liệu")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("📥 Toàn bộ JSON", export_all_json_bytes(),
                          "app_data.json", "application/json", use_container_width=True)
    with c2:
        st.download_button("📥 Bài giảng JSON", export_lectures_json_bytes(),
                          "lectures.json", "application/json", use_container_width=True)
    with c3:
        st.download_button("📥 Bài tập JSON", export_assignments_json_bytes(),
                          "assignments.json", "application/json", use_container_width=True)
    st.markdown("---")
    up = st.file_uploader("📤 Nhập JSON (toàn bộ dữ liệu)", type=["json"])
    if up and st.button("Xác nhận nhập"):
        try:
            new_db = json.loads(up.read().decode("utf-8"))
            if isinstance(new_db, dict) and "users" in new_db:
                st.session_state.db = new_db
                save_db()
                st.success("Đã nhập dữ liệu!"); st.rerun()
            else:
                st.error("Tệp JSON không đúng định dạng!")
        except Exception as e:
            st.error(f"Lỗi: {e}")

def page_admin():
    tabs = st.tabs(["👥 Tài khoản","📚 Bài giảng","🧩 Đề bài",
                    "📄 Bài nộp","📋 Nhật ký","📊 Điểm","💾 Nhập/Xuất"])
    with tabs[0]: admin_users()
    with tabs[1]: teacher_lectures()  # Admin có quyền xóa nếu mở rộng
    with tabs[2]: teacher_assignments()
    with tabs[3]: admin_submissions()
    with tabs[4]: teacher_logs()
    with tabs[5]: admin_grades()
    with tabs[6]: admin_import_export()

# ══════════════════════════════════════════════════════════════════════
# MAIN ROUTER
# ══════════════════════════════════════════════════════════════════════
def main():
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
        st.session_state.user = None

    render_sidebar()

    if not st.session_state.logged_in:
        page_login()
        return

    u = st.session_state.user
    role_lbl = {"admin":"🛡 Quản trị viên",
                "teacher":"👨‍🏫 Giáo viên",
                "student":"👨‍🎓 Học sinh"}[u["role"]]
    render_header()
    st.success(f"Xin chào **{u['name']}** — {role_lbl}")

    if u["role"] == "admin":
        page_admin()
    elif u["role"] == "teacher":
        page_teacher()
    else:
        page_student()

if __name__ == "__main__":
    main()