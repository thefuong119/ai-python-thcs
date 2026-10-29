# ai-python-thcs
Trợ lý AI hỗ trợ lập trình Python cho học sinh Trung học cơ sở
CÁCH CHẠY GITHUB TRÊN PC
--------------------------
KHỞI TẠO GIT TẠI THƯ MỤC GỐC
Tạo một kho lưu trữ Git (Git repository) trong thư mục dự án
    git init
Bước 1. Kiểm tra trạng thái Git
    git status
Sau đó chạy tiếp:
    git branch
Kiểm tra kết quả:
- Nếu hiện No commits yet: bạn chưa tạo commit đầu tiên.
- Nếu nhánh hiện tại là master thay vì main: tên nhánh chưa khớp.
- Nếu chưa có nhánh nào được hiển thị: có thể bạn chưa khởi tạo Git hoặc chưa có commit.
Bước 2. Cách xử lý phổ biến nhất
Nếu bạn chưa tạo commit đầu tiên, hãy chạy lần lượt các lệnh sau.
1. Thêm các tệp vào vùng chuẩn bị commit
    git add .
2. Tạo commit đầu tiên
    git commit -m "Initial commit"
Nếu Git yêu cầu cấu hình tên và email, hãy sử dụng tên và email bạn đăng ký với Git:
    git config --global user.name "Nguyễn Thế Phương"
    git config --global user.email "thefuong119@gmail.com"
Thay hai giá trị mẫu bằng thông tin của bạn, sau đó chạy lại lệnh git commit.
3. Đổi tên nhánh hiện tại thành main
    git branch -M main
4. Kiểm tra địa chỉ kho lưu trữ GitHub
    git remote -v
Địa chỉ cần trỏ đến kho lưu trữ của bạn:
    git remote set-url origin https://github.com/thefuong119/ai-python-thcs.git
Nếu Git báo chưa có remote origin, hãy thêm bằng:
    git remote add origin https://github.com/thefuong119/ai-python-thcs.git
5. Đẩy mã nguồn lên GitHub
    git push -u origin main
Bước 3. Nếu vẫn gặp lỗi
Nếu bạn đã có commit nhưng vẫn không đẩy được, hãy chạy:
    git log --oneline -5
    git branch -avv
    git remote -v
