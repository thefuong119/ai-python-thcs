# -*- coding: utf-8 -*-
"""
HỆ THỐNG TRỢ LÝ DẠY HỌC LẬP TRÌNH PYTHON - STREAMLIT APP
Tác giả nguyên bản: Thầy Nguyễn Thế Phương (THCS Thành Đông, Tân Quới, Vĩnh Long)
Chuyển đổi sang Python Streamlit
"""

import streamlit as st
import sys
import io
import time
import datetime
import pandas as pd
import json

# Cấu hình trang Streamlit
st.set_page_config(
    page_title="Trợ Lý Dạy Học Lập Trình Python",
    page_icon="🐍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------------------
# KHỞI TẠO DỮ LIỆU BAN ĐẦU TRONG SESSION STATE
# ----------------------------------------------------
DEFAULT_LECTURES = [
    {"id": 1, "title": "Bài 1. Ngôn ngữ lập trình bậc cao và python", "classLevel": "8", "grade": "Lớp 8 - Tin học cơ sở", "content": "### 1. Ngôn ngữ lập trình bậc cao là gì?\n- **Ngôn ngữ máy**: Là ngôn ngữ duy nhất máy tính hiểu trực tiếp, cấu thành từ các dãy số nhị phân $0$ và $1$.\n- **Ngôn ngữ bậc cao**: Là ngôn ngữ có cú pháp gần gũi với ngôn ngữ tự nhiên và biểu thức toán học.\n\n### 2. Giới thiệu về ngôn ngữ lập trình Python\n- Python do nhà khoa học máy tính người Hà Lan **Guido van Rossum** phát triển và phát hành lần đầu năm 1991.\n- **Ưu điểm vượt trội**: Cú pháp ngắn gọn, dễ học, mã nguồn mở, cộng đồng hỗ trợ lớn.\n\n### 3. Chương trình đầu tiên với Python\n```python\nprint('Xin chào thế giới lập trình Python!')\n```"},
    {"id": 2, "title": "Bài 2. Biến và lệnh gán và biểu thức số học", "classLevel": "8", "grade": "Lớp 8 - Tin học cơ sở", "content": "### 1. Biến\n- Biến là đại lượng dùng để lưu trữ giá trị trong bộ nhớ.\n- **Quy tắc đặt tên biến**: Bắt đầu bằng chữ cái hoặc `_`; không chứa khoảng trắng; không trùng từ khóa.\n\n### 2. Lệnh gán và biểu thức số học\n- Cú pháp: `tên_biến = giá_trị`\n- Phép toán: `+`, `-`, `*`, `/`, `//`, `%`, `**`\n\n```python\na = 10\nb = 3\ntong = a + b\n```"},
    {"id": 3, "title": "Bài 3. Các lệnh vào ra đơn giản và các kiểu dữ liệu", "classLevel": "8", "grade": "Lớp 8 - Tin học cơ sở", "content": "### 1. Lệnh vào / ra\n- `print()`: Xuất dữ liệu\n- `input()`: Nhập dữ liệu từ bàn phím\n\n### 2. Kiểu dữ liệu cơ bản\n- `int`: Số nguyên, `float`: Số thực, `str`: Xâu, `bool`: Logic\n\n```python\nho_ten = input('Nhập họ tên: ')\nprint('Xin chào', ho_ten)\n```"},
    {"id": 4, "title": "Bài 4. Cấu trúc rẽ nhánh", "classLevel": "8", "grade": "Lớp 8 - Tin học cơ sở", "content": "### 1. Ý nghĩa\nDùng để điều khiển luồng dựa trên điều kiện logic.\n\n### 2. Cú pháp\n```python\nif điều_kiện_1:\n    pass\nelif điều_kiện_2:\n    pass\nelse:\n    pass\n```"},
    {"id": 5, "title": "Bài 5. Cấu trúc lặp", "classLevel": "8", "grade": "Lớp 8 - Tin học cơ sở", "content": "### 1. Vòng lặp `for`\n```python\nfor i in range(1, 6):\n    print(i)\n```\n\n### 2. Vòng lặp `while`\n```python\ns = 0\ni = 1\nwhile i <= 10:\n    s += i\n    i += 1\n```"}
]

DEFAULT_ASSIGNMENTS = [
    {
        "id": 1,
        "title": "Bài 1. Tính tổng hai số nguyên A và B",
        "classLevel": "8",
        "difficulty": "easy",
        "timeLimit": 2.0,
        "description": "### Đề bài\nCho hai số nguyên $A$ và $B$. Hãy tính tổng $A + B$.\n\n### Dữ liệu vào\nHai số nguyên $A, B$ cách nhau bởi dấu cách.\n\n### Dữ liệu ra\nIn ra tổng $A + B$.",
        "templateCode": "a, b = map(int, input().split())\nprint(a + b)\n",
        "tests": [
            {"id": 1, "input": "3 5", "expected": "8", "point": 5.0, "is_hidden": False},
            {"id": 2, "input": "-10 25", "expected": "15", "point": 5.0, "is_hidden": True}
        ]
    },
    {
        "id": 2,
        "title": "Bài 2. Kiểm tra tính chẵn lẻ",
        "classLevel": "8",
        "difficulty": "easy",
        "timeLimit": 2.0,
        "description": "### Đề bài\nNhập số nguyên $N$. In `CHAN` nếu chẵn, `LE` nếu lẻ.\n\n### Dữ liệu vào\nMột số nguyên $N$.\n\n### Dữ liệu ra\nIn ra `CHAN` hoặc `LE`.",
        "templateCode": "n = int(input())\nif n % 2 == 0:\n    print('CHAN')\nelse:\n    print('LE')\n",
        "tests": [
            {"id": 1, "input": "4", "expected": "CHAN", "point": 5.0, "is_hidden": False},
            {"id": 2, "input": "7", "expected": "LE", "point": 5.0, "is_hidden": True}
        ]
    }
]

DEFAULT_USERS = [
    {"id": 1, "name": "Thầy Nguyễn Thế Phương", "username": "086081006094", "password": "Admin#999", "school": "THCS Thành Đông", "className": "Toàn trường", "phone": "0908889999", "role": "admin", "loginCount": 45},
    {"id": 2, "name": "Cô Trần Thị Trúc Mai", "username": "comai", "password": "GiaoVien#123", "school": "THCS - THPT Mỹ Thuận", "className": "11A1", "phone": "0912345678", "role": "teacher", "loginCount": 34},
    {"id": 3, "name": "Giáo viên Dùng Thử", "username": "giaovien", "password": "Giaovien@123", "school": "THCS Thành Đông", "className": "Khối 8", "phone": "0932861739", "role": "teacher", "loginCount": 10},
    {"id": 4, "name": "Học sinh Dùng Thử", "username": "hocsinh", "password": "Hocsinh@123", "school": "THCS Thành Đông", "className": "8A1", "phone": "---", "role": "student", "loginCount": 15}
]

if "lectures" not in st.session_state:
    st.session_state.lectures = DEFAULT_LECTURES

if "assignments" not in st.session_state:
    st.session_state.assignments = DEFAULT_ASSIGNMENTS

if "users" not in st.session_state:
    st.session_state.users = DEFAULT_USERS

if "submissions" not in st.session_state:
    st.session_state.submissions = []

if "logs" not in st.session_state:
    st.session_state.logs = [
        {"time": datetime.datetime.now().strftime("%d/%m/%Y, %H:%M:%S"), "name": "Hệ thống", "identifier": "system", "role": "admin", "result": "Khởi động"}
    ]

if "current_user" not in st.session_state:
    st.session_state.current_user = None

if "ai_chat" not in st.session_state:
    st.session_state.ai_chat = []

# ----------------------------------------------------
# TIỆN ÍCH THỰC THI CODE & CHẤM ĐIỂM
# ----------------------------------------------------
def execute_python_code(code: str, stdin_data: str = ""):
    """Thực thi mã nguồn Python với dữ liệu stdin được cung cấp, trả về stdout, stderr."""
    old_stdin = sys.stdin
    old_stdout = sys.stdout
    old_stderr = sys.stderr

    sys.stdin = io.StringIO(stdin_data)
    sys.stdout = io.StringIO()
    sys.stderr = io.StringIO()

    error_msg = ""
    start_time = time.time()
    try:
        scope = {}
        exec(code, scope)
    except Exception as e:
        error_msg = str(e)
    finally:
        stdout_res = sys.stdout.getvalue()
        stderr_res = sys.stderr.getvalue()
        sys.stdin = old_stdin
        sys.stdout = old_stdout
        sys.stderr = old_stderr

    elapsed = time.time() - start_time
    final_err = stderr_res if stderr_res else error_msg
    return {
        "stdout": stdout_res,
        "stderr": final_err,
        "is_error": bool(final_err.strip()),
        "elapsed": elapsed
    }

def record_log(name, identifier, role, result):
    now = datetime.datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
    st.session_state.logs.insert(0, {
        "time": now, "name": name, "identifier": identifier, "role": role, "result": result
    })

# ----------------------------------------------------
# GIAO DIỆN CHÍNH & SIDEBAR
# ----------------------------------------------------
st.markdown("""
<div style="background-color: #0f172a; padding: 12px 20px; border-radius: 12px; margin-bottom: 20px; text-align: center; border: 1px solid #1e293b;">
    <h3 style="color: #f59e0b; margin: 0; font-size: 1.15rem;">🏆 Giáo viên Sáng tạo Ứng Dụng AI Trong Giáo Dục Việt Nam 2026</h3>
    <p style="color: #94a3b8; margin: 5px 0 0 0; font-size: 0.85rem;">Trường THCS Thành Đông, xã Tân Quới, tỉnh Vĩnh Long | Thầy Nguyễn Thế Phương</p>
</div>
""", unsafe_allow_html=True)

# SIDEBAR: ĐĂNG NHẬP & TRẠNG THÁI
with st.sidebar:
    st.title("💻 Bảng Điều Khiển")
    
    if st.session_state.current_user is None:
        st.subheader("🔐 Đăng Nhập Hệ Thống")
        login_type = st.radio("Đối tượng đăng nhập:", ["Học Sinh", "Giáo Viên / Admin"], horizontal=True)
        
        with st.form("login_form"):
            username = st.text_input("Tên đăng nhập / SĐT / Mã định danh")
            password = st.text_input("Mật khẩu", type="password")
            submitted = st.form_submit_button("Đăng Nhập", use_container_width=True)
            
            if submitted:
                found_user = None
                for u in st.session_state.users:
                    if u["username"].lower() == username.strip().lower() and u["password"] == password.strip():
                        found_user = u
                        break
                
                if found_user:
                    st.session_state.current_user = found_user
                    found_user["loginCount"] = found_user.get("loginCount", 0) + 1
                    record_log(found_user["name"], found_user["username"], found_user["role"], "Thành công")
                    st.success(f"Xin chào {found_user['name']}!")
                    st.rerun()
                else:
                    record_log(username, username, "guest", "Thất bại")
                    st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")
        
        st.info("💡 **Tài khoản dùng thử:**\n- Giáo viên: `giaovien` / `Giaovien@123`\n- Học sinh: `hocsinh` / `Hocsinh@123`")
    else:
        user = st.session_state.current_user
        st.success(f"👤 **{user['name']}**")
        st.caption(f"Vai trò: **{user['role'].upper()}** | Trường: {user.get('school', 'THCS Thành Đông')}")
        
        if st.button("🚪 Đăng Xuất", use_container_width=True):
            record_log(user["name"], user["username"], user["role"], "Đăng xuất")
            st.session_state.current_user = None
            st.rerun()

# ----------------------------------------------------
# KHÔNG GIAN LÀM VIỆC DỰA TRÊN QUYỀN
# ----------------------------------------------------
if st.session_state.current_user is None:
    # Màn hình giới thiệu khi chưa đăng nhập
    st.markdown("""
    ### 👋 Chào Mừng Đến Với Web Trợ Lý Dạy Học Lập Trình Python!
    Hệ thống hỗ trợ học sinh thực hành lập trình, tự động chấm điểm và tích hợp gia sư AI định hướng.
    
    Vui lòng đăng nhập từ thanh bên trái để vào không gian học tập và thực hành.
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        #### 🌟 Tính năng nổi bật:
        - 📚 **Học lý thuyết tương tác:** Bài giảng chuẩn hóa theo chương trình.
        - ⚡ **Thực thi Python trực tiếp:** Chạy thử và debug mã nguồn tức thì.
        - 🤖 **Gia sư AI:** Hỗ trợ phương pháp Socratic, gợi ý thuật toán và chỉ điểm lỗi sai.
        - 📊 **Chấm điểm tự động:** Đánh giá độ chính xác qua nhiều bộ test case.
        """)
    with col2:
        st.markdown("""
        #### 📞 Thông tin liên hệ tác giả:
        - **Tác giả:** Thầy Nguyễn Thế Phương
        - **Đơn vị:** Trường THCS Thành Đông, xã Tân Quới, tỉnh Vĩnh Long
        - **Điện thoại:** 0932.861.739
        - **Email:** thefuong119@gmail.com
        """)

else:
    user = st.session_state.current_user
    user_role = user["role"]
    
    # Menu điều hướng các tab theo vai trò
    if user_role == "student":
        tab1, tab2, tab3 = st.tabs(["📚 Bài Giảng & Lý Thuyết", "💻 Thực Hành Lập Trình", "🤖 Gia Sư AI"])
        
        # 1. BÀI GIẢNG HỌC SINH
        with tab1:
            st.subheader("Thư Viện Bài Giảng Python Chuẩn")
            grade_filter = st.selectbox("Lọc theo khối lớp:", ["Tất cả", "Khối 6", "Khối 7", "Khối 8", "Khối 9"], index=3)
            
            selected_lectures = st.session_state.lectures
            if grade_filter != "Tất cả":
                g_num = grade_filter.replace("Khối ", "")
                selected_lectures = [l for l in selected_lectures if str(l.get("classLevel", "8")) == g_num]
            
            if selected_lectures:
                lecture_titles = [f"{l['title']} (Khối {l.get('classLevel', '8')})" for l in selected_lectures]
                sel_idx = st.selectbox("Chọn bài giảng để đọc:", range(len(lecture_titles)), format_func=lambda i: lecture_titles[i])
                curr_lec = selected_lectures[sel_idx]
                
                st.markdown("---")
                st.markdown(f"### {curr_lec['title']}")
                st.caption(f"Chủ đề: {curr_lec.get('grade', 'Tin học cơ sở')} | Khối {curr_lec.get('classLevel', '8')}")
                st.markdown(curr_lec["content"])
            else:
                st.warning("Chưa có bài giảng nào trong khối lớp này.")

        # 2. THỰC HÀNH LẬP TRÌNH HỌC SINH
        with tab2:
            st.subheader("Không Gian Thực Hành & Nộp Bài Chấm Điểm")
            
            # Chọn bài tập
            assign_options = [f"#{a['id']}. {a['title']} ({a.get('difficulty', 'easy').upper()})" for a in st.session_state.assignments]
            selected_assign_idx = st.selectbox("Chọn bài tập thực hành:", range(len(assign_options)), format_func=lambda i: assign_options[i])
            curr_assign = st.session_state.assignments[selected_assign_idx]
            
            col_left, col_right = st.columns([1, 1.2])
            
            with col_left:
                st.markdown(f"#### 📝 {curr_assign['title']}")
                st.markdown(curr_assign["description"])
                
                st.markdown("##### 📌 Test mẫu công khai:")
                public_tests = [t for t in curr_assign["tests"] if not t.get("is_hidden", False)]
                for pt in public_tests:
                    st.code(f"Input : {pt['input']}\nOutput: {pt['expected']}", language="text")
                
                if st.button("💡 Xem Code Mẫu / Gợi Ý"):
                    st.code(curr_assign.get("templateCode", "# Chưa có mẫu"), language="python")

            with col_right:
                st.markdown("#### 💻 Trình Soạn Thảo & Chạy Thử")
                code_input = st.text_area(
                    "Mã Python của bạn:",
                    value=curr_assign.get("templateCode", "# Viết mã tại đây...\n"),
                    height=250,
                    key="student_code_area"
                )
                
                custom_in = st.text_input("Dữ liệu đầu vào thử nghiệm (stdin):", value="3 5")
                
                btn_run, btn_submit = st.columns(2)
                with btn_run:
                    if st.button("▶️ Chạy Thử Nghiệm", use_container_width=True):
                        res = execute_python_code(code_input, custom_in)
                        if res["is_error"]:
                            st.error(f"❌ Lỗi thực thi:\n{res['stderr']}")
                        else:
                            st.success(f"✅ Kết quả (thời gian: {res['elapsed']:.3f}s):")
                            st.code(res["stdout"], language="text")
                
                with btn_submit:
                    if st.button("📤 Nộp Bài Chấm Điểm", use_container_width=True, type="primary"):
                        total_pts = 0
                        max_pts = 0
                        results_summary = []
                        
                        for idx, test in enumerate(curr_assign["tests"]):
                            p = test.get("point", 5.0)
                            max_pts += p
                            exec_res = execute_python_code(code_input, test["input"])
                            
                            is_passed = (not exec_res["is_error"]) and (exec_res["stdout"].strip() == str(test["expected"]).strip())
                            if is_passed:
                                total_pts += p
                                results_summary.append(f"✅ Test #{idx+1}: ĐẠT (+{p}đ)")
                            else:
                                results_summary.append(f"❌ Test #{idx+1}: SAI (0đ)")
                        
                        st.subheader(f"🎯 Điểm số: {total_pts:.1f} / {max_pts:.1f}")
                        for r in results_summary:
                            st.write(r)
                            
                        # Lưu lịch sử nộp bài
                        st.session_state.submissions.insert(0, {
                            "student": user["name"],
                            "username": user["username"],
                            "assignmentTitle": curr_assign["title"],
                            "score": total_pts,
                            "timestamp": datetime.datetime.now().strftime("%d/%m/%Y, %H:%M:%S")
                        })
                        st.toast(f"Đã ghi nhận bài nộp: {total_pts} điểm!", icon="🎉")

        # 3. GIA SƯ AI
        with tab3:
            st.subheader("🤖 Trợ Lý Gia Sư Trí Tuệ Nhân Tạo")
            mode = st.radio("Chế độ hướng dẫn:", ["1. Socratic (Gợi mở câu hỏi)", "2. Hướng tiếp cận thuật toán", "3. Phân tích lỗi code"], horizontal=True)
            
            user_question = st.chat_input("Nhập câu hỏi hoặc vướng mắc của em...")
            if user_question:
                st.session_state.ai_chat.append({"role": "user", "content": user_question})
                
                # Logic phản hồi trợ lý
                if "Socratic" in mode:
                    reply = "Thầy muốn hỏi em trước: Trong bài toán này, biến nào là đầu vào và công thức toán học cần áp dụng là gì? Hãy chia nhỏ bài toán nhé!"
                elif "thuật toán" in mode:
                    reply = "💡 Hướng tiếp cận: Em hãy dùng `input().split()` để đọc dữ liệu trên cùng một dòng, sau đó ép kiểu `int()` và thực hiện phép toán tương ứng."
                else:
                    reply = "🔍 Phân tích: Hãy kiểm tra kỹ kiểu dữ liệu khi đọc từ `input()` (mặc định là string). Nếu cần tính toán số học, em phải bọc bằng `int()` hoặc `float()`."
                
                st.session_state.ai_chat.append({"role": "assistant", "content": reply})

            for msg in st.session_state.ai_chat:
                with st.chat_message(msg["role"]):
                    st.write(msg["content"])

    elif user_role in ["teacher", "admin"]:
        tabs = st.tabs(["📚 Quản Lý Bài Giảng", "📝 Ngân Hàng Đề Bài", "📊 Bài Nộp Học Sinh", "📜 Nhật Ký Hệ Thống", "👥 Quản Trị Người Dùng"] if user_role == "admin" else ["📚 Quản Lý Bài Giảng", "📝 Ngân Hàng Đề Bài", "📊 Bài Nộp Học Sinh", "📜 Nhật Ký Hệ Thống"])
        
        # 1. QUẢN LÝ BÀI GIẢNG
        with tabs[0]:
            st.subheader("Thư Viện Bài Giảng")
            
            with st.expander("➕ Thêm Bài Giảng Mới", expanded=False):
                with st.form("add_lecture_form"):
                    new_l_title = st.text_input("Tiêu đề bài giảng:", placeholder="VD: Bài 6. Kiểu dữ liệu danh sách")
                    col_k, col_cd = st.columns(2)
                    with col_k:
                        new_l_class = st.selectbox("Khối lớp:", ["6", "7", "8", "9", "10", "11", "12"], index=2)
                    with col_cd:
                        new_l_grade = st.text_input("Chủ đề:", value="Tin học cơ sở")
                    new_l_body = st.text_area("Nội dung bài giảng (Markdown & LaTeX):", height=200)
                    if st.form_submit_button("Lưu Bài Giảng"):
                        if new_l_title and new_l_body:
                            new_id = len(st.session_state.lectures) + 1
                            st.session_state.lectures.append({
                                "id": new_id, "title": new_l_title, "classLevel": new_l_class, "grade": new_l_grade, "content": new_l_body
                            })
                            st.success("Đã thêm bài giảng thành công!")
                            st.rerun()
                        else:
                            st.error("Vui lòng điền đủ tiêu đề và nội dung.")
            
            # Danh sách bài giảng hiện tại
            for l in st.session_state.lectures:
                with st.expander(f"📖 #{l['id']} - {l['title']} (Khối {l.get('classLevel', '8')})"):
                    st.markdown(l["content"])
                    if user_role == "admin":
                        if st.button("🗑️ Xóa bài giảng", key=f"del_lec_{l['id']}"):
                            st.session_state.lectures = [x for x in st.session_state.lectures if x["id"] != l["id"]]
                            st.success("Đã xóa bài giảng!")
                            st.rerun()

        # 2. NGÂN HÀNG ĐỀ BÀI
        with tabs[1]:
            st.subheader("Ngân Hàng Bài Tập & Test Cases")
            
            with st.expander("➕ Soạn Thảo Đề Bài Mới", expanded=False):
                with st.form("create_assign_form"):
                    a_title = st.text_input("Tên bài tập:", placeholder="VD: Bài 3. Tính diện tích tam giác")
                    c1, c2 = st.columns(2)
                    with c1:
                        a_level = st.selectbox("Khối:", ["6", "7", "8", "9"], index=2)
                    with c2:
                        a_diff = st.selectbox("Độ khó:", ["easy", "medium", "hard"])
                    a_desc = st.text_area("Yêu cầu đề bài (Markdown & Math):", height=150)
                    a_sol = st.text_area("Mã nguồn mẫu Python:", height=100)
                    
                    st.write("**Bộ test mẫu:**")
                    t1_in = st.text_input("Test 1 Input:", value="5")
                    t1_out = st.text_input("Test 1 Expected Output:", value="25")
                    
                    if st.form_submit_button("Lưu Bài Tập Mới"):
                        if a_title:
                            new_aid = len(st.session_state.assignments) + 1
                            st.session_state.assignments.append({
                                "id": new_aid, "title": a_title, "classLevel": a_level, "difficulty": a_diff,
                                "description": a_desc, "templateCode": a_sol,
                                "tests": [{"id": 1, "input": t1_in, "expected": t1_out, "point": 10.0, "is_hidden": False}]
                            })
                            st.success("Đã tạo bài tập mới thành công!")
                            st.rerun()

            # Hiển thị danh sách đề bài
            df_assign = pd.DataFrame([
                {"ID": a["id"], "Tên Bài": a["title"], "Khối": a["classLevel"], "Độ Khó": a["difficulty"], "Số Test": len(a.get("tests", []))}
                for a in st.session_state.assignments
            ])
            st.dataframe(df_assign, use_container_width=True)

        # 3. BÀI NỘP HỌC SINH
        with tabs[2]:
            st.subheader("Kết Quả Nộp Bài Học Sinh")
            if st.session_state.submissions:
                df_sub = pd.DataFrame(st.session_state.submissions)
                st.dataframe(df_sub, use_container_width=True)
                
                # Xuất CSV
                csv_data = df_sub.to_csv(index=False).encode('utf-8-sig')
                st.download_button("📥 Xuất Tệp CSV (Excel)", data=csv_data, file_name="ket_qua_nop_bai.csv", mime="text/csv")
            else:
                st.info("Chưa có lượt nộp bài nào.")

        # 4. NHẬT KÝ HỆ THỐNG
        with tabs[3]:
            st.subheader("Nhật Ký Truy Cập")
            if st.session_state.logs:
                df_logs = pd.DataFrame(st.session_state.logs)
                st.dataframe(df_logs, use_container_width=True)
            else:
                st.info("Nhật ký trống.")

        # 5. QUẢN TRỊ NGƯỜI DÙNG (CHỈ ADMIN)
        if user_role == "admin" and len(tabs) > 4:
            with tabs[4]:
                st.subheader("Quản Trị Danh Sách Người Dùng")
                df_users = pd.DataFrame(st.session_state.users)[["id", "name", "username", "role", "school", "className", "loginCount"]]
                st.dataframe(df_users, use_container_width=True)
                
                with st.expander("➕ Tạo Tài Khoản Mới"):
                    with st.form("new_user_form"):
                        nu_name = st.text_input("Họ và tên:")
                        nu_uname = st.text_input("Tên đăng nhập:")
                        nu_pass = st.text_input("Mật khẩu:", value="123456")
                        nu_role = st.selectbox("Vai trò:", ["student", "teacher", "admin"])
                        nu_class = st.text_input("Lớp / Khối:", value="Khối 8")
                        if st.form_submit_button("Thêm Tài Khoản"):
                            if nu_name and nu_uname:
                                new_uid = len(st.session_state.users) + 1
                                st.session_state.users.append({
                                    "id": new_uid, "name": nu_name, "username": nu_uname, "password": nu_pass,
                                    "role": nu_role, "school": "THCS Thành Đông", "className": nu_class, "phone": "---", "loginCount": 0
                                })
                                st.success("Đã thêm người dùng thành công!")
                                st.rerun()

# ----------------------------------------------------
# CHÂN TRANG
# ----------------------------------------------------
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #64748b; font-size: 0.85rem;">
    Bản quyền © 2026 Giáo viên: Nguyễn Thế Phương, Trường THCS Thành Đông, xã Tân Quới, tỉnh Vĩnh Long.
</div>
""", unsafe_allow_html=True)
