"""
WEB TRỢ LÝ DẠY HỌC LẬP TRÌNH PYTHON
Tác giả: Thầy Nguyễn Thế Phương - Trường THCS Thành Đông, Tân Quới, Vĩnh Long
Năm học: 2025 - 2027
Mô tả: Nền tảng học trực tuyến, thực hành lập trình Python & Chấm điểm tự động với trợ lý AI.
"""

import streamlit as st
import io
import sys
import json
import time
import pandas as pd
from datetime import datetime

# ==============================================================================
# 1. CẤU HÌNH TRANG VÀ GIAO DIỆN CHÍNH
# ==============================================================================
st.set_page_config(
    page_title="WEB TRỢ LÝ DẠY HỌC LẬP TRÌNH PYTHON",
    page_icon="🐍",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS cho giao diện hiện đại
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        padding: 1.5rem;
        border-radius: 1rem;
        color: white;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .banner-award {
        background-color: #0f172a;
        color: #fbbf24;
        padding: 0.5rem;
        text-align: center;
        font-weight: bold;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
        font-size: 0.9rem;
    }
    .author-card {
        background: linear-gradient(180deg, #ffffff 0%, #f8fafc 100%);
        border: 1px solid #e2e8f0;
        border-radius: 1rem;
        padding: 1.5rem;
        margin-top: 1rem;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }
    .stButton>button {
        border-radius: 0.75rem;
        font-weight: 600;
    }
    .lecture-card {
        background-color: #f8fafc;
        border-left: 4px solid #3b82f6;
        padding: 1rem;
        border-radius: 0.5rem;
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. KHỞI TẠO STATE & DỮ LIỆU BAN ĐẦU
# ==============================================================================
if "user" not in st.session_state:
    st.session_state.user = {"is_logged_in": False, "username": "", "name": "", "role": "guest", "class_name": ""}

if "lectures" not in st.session_state:
    st.session_state.lectures = [
        {
            "id": 1,
            "title": "Bài 1. Ngôn ngữ lập trình bậc cao và python",
            "class_level": "8",
            "grade": "Lớp 8 - Tin học cơ sở",
            "content": """### 1. Ngôn ngữ lập trình bậc cao là gì?
- **Ngôn ngữ máy**: Là ngôn ngữ duy nhất máy tính hiểu trực tiếp, cấu thành từ các dãy số nhị phân $0$ và $1$.
- **Ngôn ngữ bậc cao**: Là ngôn ngữ có cú pháp gần gũi với ngôn ngữ tự nhiên (tiếng Anh) và biểu thức toán học, giúp con người dễ đọc, dễ viết.

### 2. Giới thiệu về ngôn ngữ lập trình Python
- Python do nhà khoa học máy tính **Guido van Rossum** phát triển năm 1991.
- **Ưu điểm**: Cú pháp ngắn gọn, dễ học, mã nguồn mở, cộng đồng hỗ trợ lớn.

### 3. Chương trình đầu tiên
```python
print('Xin chào thế giới lập trình Python!')
```"""
        },
        {
            "id": 2,
            "title": "Bài 2. Biến và lệnh gán và biểu thức số học",
            "class_level": "8",
            "grade": "Lớp 8 - Tin học cơ sở",
            "content": """### 1. Biến (Variable)
- Biến dùng để lưu trữ giá trị trong bộ nhớ.
- **Quy tắc**: Bắt đầu bằng chữ cái hoặc `_`, không chứa khoảng trắng, không trùng từ khóa.

### 2. Phép toán số học
- Cộng `+`, Trừ `-`, Nhân `*`, Chia `/`
- Chia lấy nguyên `//`, Chia lấy dư `%`, Lũy thừa `**`

```python
a = 10
b = 3
print("Chia nguyên:", a // b) # Kết quả: 3
print("Chia dư:", a % b)     # Kết quả: 1
```"""
        },
        {
            "id": 3,
            "title": "Bài 3. Các lệnh vào ra đơn giản và các kiểu dữ liệu",
            "class_level": "8",
            "grade": "Lớp 8 - Tin học cơ sở",
            "content": """### 1. Lệnh Vào / Ra
- `print()`: Xuất dữ liệu ra màn hình.
- `input()`: Nhập dữ liệu dạng chuỗi (`str`). Ép kiểu số nguyên dùng `int(input())`.

### 2. Các kiểu dữ liệu cơ bản
- `int`: Số nguyên
- `float`: Số thực
- `str`: Chuỗi ký tự
- `bool`: Logic (`True`/`False`)"""
        }
    ]

if "assignments" not in st.session_state:
    st.session_state.assignments = [
        {
            "id": 1,
            "title": "Tính tổng hai số nguyên A và B",
            "class_level": "8",
            "difficulty": "Dễ",
            "description": "Cho hai số nguyên $A$ và $B$. Hãy viết chương trình tính và in ra tổng $A + B$.\n\n**Input:** Nhập 2 số cách nhau bởi khoảng trắng.\n**Output:** In ra kết quả tổng.",
            "template_code": "# Nhập 2 số từ dữ liệu vào\na, b = map(int, input().split())\nprint(a + b)",
            "tests": [
                {"input": "3 5", "expected": "8", "point": 5.0, "is_hidden": False},
                {"input": "-10 25", "expected": "15", "point": 5.0, "is_hidden": True}
            ]
        },
        {
            "id": 2,
            "title": "Kiểm tra tính chẵn lẻ",
            "class_level": "8",
            "difficulty": "Dễ",
            "description": "Nhập vào một số nguyên $N$. In ra `CHAN` nếu $N$ là số chẵn, ngược lại in ra `LE`.",
            "template_code": "n = int(input())\nif n % 2 == 0:\n    print('CHAN')\nelse:\n    print('LE')",
            "tests": [
                {"input": "4", "expected": "CHAN", "point": 5.0, "is_hidden": False},
                {"input": "7", "expected": "LE", "point": 5.0, "is_hidden": True}
            ]
        }
    ]

if "registered_users" not in st.session_state:
    st.session_state.registered_users = [
        {"id": 1, "name": "Thầy Nguyễn Thế Phương", "username": "admin", "password": "Admin#999", "role": "admin", "school": "THCS Thành Đông", "className": "Toàn trường", "phone": "0932861739", "loginCount": 50},
        {"id": 2, "name": "Cô Trần Thị Trúc Mai", "username": "comai", "password": "GiaoVien#123", "role": "teacher", "school": "THCS Thành Đông", "className": "Khối 8", "phone": "0912345678", "loginCount": 20},
        {"id": 3, "name": "Trần Ngọc Hân", "username": "han", "password": "HocSinh#123", "role": "student", "school": "THCS Thành Đông", "className": "8A1", "phone": "---", "loginCount": 5}
    ]

if "submissions" not in st.session_state:
    st.session_state.submissions = []

if "logs" not in st.session_state:
    st.session_state.logs = []

# ==============================================================================
# 3. HÀM XỬ LÝ LOGIC TRUNG TÂM
# ==============================================================================
def log_access(name, username, role, status):
    st.session_state.logs.insert(0, {
        "time": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "name": name,
        "username": username,
        "role": role,
        "status": status
    })

def execute_python(code, input_data=""):
    """Thực thi mã Python trong Sandbox an toàn"""
    old_stdout = sys.stdout
    old_stdin = sys.stdin
    redirected_output = io.StringIO()
    sys.stdout = redirected_output
    sys.stdin = io.StringIO(input_data)
    
    error_msg = ""
    try:
        exec(code, {})
    except Exception as e:
        error_msg = str(e)
    finally:
        sys.stdout = old_stdout
        sys.stdin = old_stdin
        
    output = redirected_output.getvalue()
    return output, error_msg

# ==============================================================================
# 4. GIAO DIỆN HEADER & BANNER
# ==============================================================================
st.markdown("""
<div class="banner-award">
    🏆 Giáo viên Sáng tạo Ứng Dụng AI Trong Giáo Dục Việt Nam 2026 🌟
</div>
<div class="main-header">
    <h1 style="margin:0; font-size: 2rem;">WEB TRỢ LÝ DẠY HỌC LẬP TRÌNH PYTHON</h1>
    <p style="margin:5px 0 0 0; opacity: 0.9;">Nền tảng thực hành lập trình & Chấm điểm tự động</p>
</div>
""", unsafe_allow_html=True)

# ==============================================================================
# 5. SIDEBAR: XÁC THỰC TÀI KHOẢN & ĐIỀU HƯỚNG VAI TRÒ
# ==============================================================================
with st.sidebar:
    st.image("https://placehold.co/150x150/3b82f6/white?text=AI+Python", width=120)
    st.title("📌 Danh Mục Hệ Thống")
    
    if not st.session_state.user["is_logged_in"]:
        st.subheader("🔑 Đăng Nhập / Đăng Ký")
        auth_mode = st.radio("Thao tác:", ["Đăng nhập", "Đăng ký Học Sinh"])
        role_login = st.selectbox("Vai trò đăng nhập:", ["Học Sinh", "Giáo Viên / Admin"])
        
        if auth_mode == "Đăng nhập":
            username_input = st.text_input("Tên đăng nhập / Số điện thoại:")
            password_input = st.text_input("Mật khẩu:", type="password")
            
            if st.button("🚀 XÁC THỰC ĐĂNG NHẬP", use_container_width=True):
                user_found = None
                for u in st.session_state.registered_users:
                    if (u["username"] == username_input or u["phone"] == username_input) and u["password"] == password_input:
                        user_found = u
                        break
                
                if user_found:
                    st.session_state.user = {
                        "is_logged_in": True,
                        "username": user_found["username"],
                        "name": user_found["name"],
                        "role": user_found["role"],
                        "class_name": user_found.get("className", "")
                    }
                    user_found["loginCount"] = user_found.get("loginCount", 0) + 1
                    log_access(user_found["name"], user_found["username"], user_found["role"], "Thành công")
                    st.success(f"Chào mừng {user_found['name']}!")
                    st.rerun()
                else:
                    log_access(username_input, username_input, role_login, "Thất bại")
                    st.error("Tên đăng nhập hoặc mật khẩu không chính xác!")
        else:
            st.write("**Đăng ký tài khoản Học Sinh mới:**")
            reg_name = st.text_input("Họ và tên:")
            reg_class = st.text_input("Lớp:", value="8A1")
            reg_user = st.text_input("Tên đăng nhập mới:")
            reg_pass = st.text_input("Mật khẩu mới:", type="password")
            
            if st.button("📝 ĐĂNG KÝ NGAY", use_container_width=True):
                if reg_name and reg_user and reg_pass:
                    st.session_state.registered_users.append({
                        "id": len(st.session_state.registered_users) + 1,
                        "name": reg_name,
                        "username": reg_user,
                        "password": reg_pass,
                        "role": "student",
                        "school": "THCS Thành Đông",
                        "className": reg_class,
                        "phone": "---",
                        "loginCount": 0
                    })
                    st.success("Đăng ký thành công! Vui lòng chọn tab Đăng nhập.")
                else:
                    st.error("Vui lòng nhập đầy đủ thông tin!")
    else:
        st.success(f"👤 **{st.session_state.user['name']}**")
        st.info(f"🔰 Vai trò: **{st.session_state.user['role'].upper()}** | Lớp: **{st.session_state.user['class_name']}**")
        
        if st.button("🚪 Đăng xuất", use_container_width=True):
            log_access(st.session_state.user["name"], st.session_state.user["username"], st.session_state.user["role"], "Đăng xuất")
            st.session_state.user = {"is_logged_in": False, "username": "", "name": "", "role": "guest", "class_name": ""}
            st.rerun()

# ==============================================================================
# 6. KHÔNG GIAN HIỂN THỊ CHÍNH (THEO VAI TRÒ)
# ==============================================================================

if not st.session_state.user["is_logged_in"]:
    # MÀN HÌNH GIỚI THIỆU / CHƯA ĐĂNG NHẬP
    st.warning("⚠️ **Hệ thống yêu cầu xác thực danh tính! Vui lòng đăng nhập tài khoản ở thanh điều hướng bên trái.**")
    
    st.markdown("""
    <div class="author-card">
        <h3>📘 Giới thiệu Ứng dụng & Tác giả</h3>
        <p>Ứng dụng do <strong>thầy Nguyễn Thế Phương</strong>, giáo viên Trường THCS Thành Đông, xã Tân Quới, tỉnh Vĩnh Long xây dựng để giúp học sinh ôn thi học sinh giỏi Tin học và luyện tập lập trình Python.</p>
        <ul>
            <li><strong>Tính năng:</strong> Bài giảng theo chuẩn kiến thức, Ngân hàng đề bài phong phú, Chấm điểm tự động.</li>
            <li><strong>Liên hệ:</strong> 📞 0932.861.739 | ✉️ thefuong119@gmail.com</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

else:
    role = st.session_state.user["role"]

    # --------------------------------------------------------------------------
    # A. PHÂN HỆ HỌC SINH (STUDENT)
    # --------------------------------------------------------------------------
    if role == "student":
        tab1, tab2 = st.tabs(["📚 Bài Giảng & Lý Thuyết", "💻 Thực Hành Lập Trình & Chấm Điểm"])
        
        with tab1:
            st.subheader(" Thư Viện Bài Giảng Python Chuẩn")
            selected_lecture_id = st.selectbox(
                "Chọn bài giảng:",
                options=[l["id"] for l in st.session_state.lectures],
                format_func=lambda x: next(l["title"] for l in st.session_state.lectures if l["id"] == x)
            )
            lec = next(l for l in st.session_state.lectures if l["id"] == selected_lecture_id)
            st.markdown(f"### {lec['title']}")
            st.caption(f"Chủ đề: {lec['grade']}")
            st.markdown(lec["content"])
            
        with tab2:
            st.subheader("⚡ Khung Thực Hành Lập Trình Python")
            col_desc, col_code = st.columns([1, 1])
            
            with col_desc:
                selected_assign_id = st.selectbox(
                    "Chọn bài tập thực hành:",
                    options=[a["id"] for a in st.session_state.assignments],
                    format_func=lambda x: next(a["title"] for a in st.session_state.assignments if a["id"] == x)
                )
                ass = next(a for a in st.session_state.assignments if a["id"] == selected_assign_id)
                st.markdown(f"### {ass['title']}")
                st.write(f"**Độ khó:** {ass['difficulty']}")
                st.markdown(ass["description"])
                
                st.write("**Ví dụ mẫu (Public Test):**")
                for t in ass["tests"]:
                    if not t["is_hidden"]:
                        st.code(f"Input: {t['input']}\nOutput: {t['expected']}", language="text")

            with col_code:
                user_code = st.text_area("Viết mã Python của bạn tại đây:", value=ass["template_code"], height=250)
                custom_in = st.text_input("Dữ liệu đầu vào thử nghiệm (stdin):", value="3 5")
                
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("▶️ CHẠY THỬ", use_container_width=True):
                        out, err = execute_python(user_code, custom_in)
                        st.write("**Kết quả xuất (stdout):**")
                        if err:
                            st.error(err)
                        else:
                            st.code(out, language="text")
                
                with col_btn2:
                    if st.button("🚀 NỘP BÀI CHẤM ĐIỂM", type="primary", use_container_width=True):
                        total_score = 0.0
                        max_score = 0.0
                        st.write("---")
                        st.write("**Kết quả chấm chi tiết:**")
                        for idx, t in enumerate(ass["tests"]):
                            max_score += t["point"]
                            out, err = execute_python(user_code, t["input"])
                            passed = (not err) and (out.strip() == t["expected"].strip())
                            if passed:
                                total_score += t["point"]
                                st.success(f"Test #{idx+1}: CHÍNH XÁC (+{t['point']}đ)")
                            else:
                                st.error(f"Test #{idx+1}: SAI KẾT QUẢ (0đ)")
                        
                        st.metric("TỔNG ĐIỂM", f"{total_score} / {max_score}")
                        st.session_state.submissions.append({
                            "student": st.session_state.user["name"],
                            "title": ass["title"],
                            "score": total_score,
                            "time": datetime.now().strftime("%d/%m/%Y %H:%M")
                        })

    # --------------------------------------------------------------------------
    # B. PHÂN HỆ GIÁO VIÊN & ADMIN
    # --------------------------------------------------------------------------
    elif role in ["teacher", "admin"]:
        st.subheader("👨‍🏫 Bảng Điều Khiển Quản Lý Dạy Học")
        
        menu_options = ["📚 Bài Giảng", "📝 Ngân Hàng Đề Bài", "📊 Sổ Điểm"]
        if role == "admin":
            menu_options.extend(["👥 Quản Lý Tài Khoản", "📜 Nhật Ký Truy Cập"])
            
        tabs = st.tabs(menu_options)
        
        # Tab 1: Quản lý bài giảng
        with tabs[0]:
            st.write("### Danh sách bài giảng lý thuyết")
            df_lec = pd.DataFrame(st.session_state.lectures)[["id", "title", "class_level", "grade"]]
            st.dataframe(df_lec, use_container_width=True)
            
            st.write("---")
            st.write("**Thêm bài giảng mới:**")
            new_title = st.text_input("Tiêu đề bài giảng:")
            new_level = st.selectbox("Khối lớp:", ["6", "7", "8", "9", "10", "11", "12"], index=2)
            new_content = st.text_area("Nội dung bài giảng (Markdown):", height=150)
            if st.button("➕ Thêm Bài Giảng", type="primary"):
                if new_title and new_content:
                    new_id = max([l["id"] for l in st.session_state.lectures], default=0) + 1
                    st.session_state.lectures.append({
                        "id": new_id,
                        "title": new_title,
                        "class_level": new_level,
                        "grade": f"Lớp {new_level} - Tin học",
                        "content": new_content
                    })
                    st.success("Đã lưu bài giảng mới!")
                    st.rerun()

        # Tab 2: Ngân hàng đề bài
        with tabs[1]:
            st.write("### Danh sách bài tập chấm điểm")
            df_ass = pd.DataFrame(st.session_state.assignments)[["id", "title", "class_level", "difficulty"]]
            st.dataframe(df_ass, use_container_width=True)
            
            st.write("---")
            st.write("**Soạn bài tập mới:**")
            ass_title = st.text_input("Tên bài tập:")
            ass_diff = st.selectbox("Độ khó:", ["Dễ", "Trung bình", "Nâng cao"])
            ass_desc = st.text_area("Mô tả đề bài:")
            ass_sol = st.text_area("Code mẫu (Python):", value="a, b = map(int, input().split())\nprint(a+b)")
            
            if st.button("➕ Lưu Đề Bài Mới", type="primary"):
                if ass_title and ass_desc:
                    new_id = max([a["id"] for a in st.session_state.assignments], default=0) + 1
                    st.session_state.assignments.append({
                        "id": new_id,
                        "title": ass_title,
                        "class_level": "8",
                        "difficulty": ass_diff,
                        "description": ass_desc,
                        "template_code": ass_sol,
                        "tests": [{"input": "1 2", "expected": "3", "point": 10.0, "is_hidden": False}]
                    })
                    st.success("Đã thêm bài tập mới vào ngân hàng!")
                    st.rerun()

        # Tab 3: Sổ điểm
        with tabs[2]:
            st.write("### Kết quả nộp bài của học sinh")
            if st.session_state.submissions:
                st.dataframe(pd.DataFrame(st.session_state.submissions), use_container_width=True)
            else:
                st.info("Chưa có lượt nộp bài nào.")

        # Tab Admin
        if role == "admin":
            with tabs[3]:
                st.write("### Quản lý danh sách tài khoản")
                st.dataframe(pd.DataFrame(st.session_state.registered_users)[["id", "name", "username", "role", "className"]], use_container_width=True)
                
            with tabs[4]:
                st.write("### Lịch sử truy cập hệ thống")
                st.dataframe(pd.DataFrame(st.session_state.logs), use_container_width=True)

# Footer
st.markdown("---")
st.caption("Bản quyền © 2026 Thầy Nguyễn Thế Phương - THCS Thành Đông, Tân Quới, Vĩnh Long.")