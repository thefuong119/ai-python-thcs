# TÀI LIỆU KỸ THUẬT XỬ LÝ GIỌNG ĐỌC CHO ROBOT

## 1. Tổng quan kiến trúc hệ thống giọng đọc

Thay vì phụ thuộc vào `window.speechSynthesis` (vốn phụ thuộc hoàn toàn vào gói giọng cục bộ cài trên máy tính hay điện thoại của người dùng, thường bị thiếu giọng tiếng Việt tự nhiên trên nhiều phiên bản Windows/Android cũ), hệ thống triển khai một **Audio Streaming Engine** tùy biến cao cấp thông qua đối tượng toàn cục `RobotTTS`.

### Sơ đồ luồng xử lý âm thanh (Audio Pipeline)

```text
[Nút bấm / Kéo thả Robot]
          |
          v
1. speakRobotText(text)
          |
          v
2. RobotTTS.cleanMarkdownForSpeech(text)
   ├── Chuyển đổi công thức Toán học KaTeX/MathType ($...$) thành lời đọc tự nhiên
   └── Làm sạch Markdown (loại bỏ dấu *, #, code fences, link...)
          |
          v
3. RobotTTS.splitTextIntoBilingualPlaylist(cleanText)
   ├── Nhận diện từ vựng: Tiếng Việt (vi) vs Từ khóa lập trình Tiếng Anh (en)
   └── Phân đoạn chuỗi thành các token âm thanh tối ưu (< 140 ký tự)
          |
          v
4. RobotTTS.playNext(sessionId)
   ├── Gọi dịch vụ luồng âm thanh phát thanh (TTS CDN Endpoints)
   ├── Tải trước (Preload) đoạn âm thanh kế tiếp trong nền
   └── Kích hoạt hiệu ứng visualizer (sóng âm) trên Robot AI
```

## 2. Chi tiết các hàm điều khiển & tương tác

### 2.1. Hàm điều khiển trung gian: `speakRobotText(text)`

- **Vị trí:** Nằm ở tầng điều phối ứng dụng.
- **Chức năng:** Tiếp nhận văn bản thô từ bất kỳ phân hệ nào (bài giảng, kết quả chấm điểm, phản hồi gia sư AI, giới thiệu tác giả) và chuyển giao cho `RobotTTS.speak(text)`.
- **Cú pháp:**

```javascript
function speakRobotText(text) {
    RobotTTS.speak(text);
}
```

### 2.2. Hàm đọc lời giới thiệu tác giả: `speakAuthorIntro()`

- **Sự kiện kích hoạt:** Khi người dùng nhấn nút “Nghe giới thiệu” trên thẻ thông tin tác giả Thầy Nguyễn Thế Phương.
- **Cơ chế hoạt động:**
  1. Kiểm tra và tự động đóng modal đăng nhập (nếu đang mở) để không chắn tầm nhìn người dùng.
  2. Chuẩn bị đoạn văn bản giới thiệu đầy đủ bao gồm thông tin trường THCS Thành Đông, mục tiêu năm học 2026–2027 và thông tin liên hệ.
  3. Gọi `speakRobotText()` và gửi thông báo Toast cho người dùng.

### 2.3. Hàm đọc tại vị trí Robot: `readTextAtRobotPosition(isFromDrag)`

- **Sự kiện kích hoạt:** Khi người dùng kéo thả Robot và thả chuột/ngón tay lên một đoạn văn bản bất kỳ trên màn hình, hoặc khi nhấp trực tiếp vào Robot.
- **Cơ chế hoạt động:**
  1. Sử dụng tọa độ trung tâm của Robot trên màn hình `(x, y)`.
  2. Sử dụng `document.elementsFromPoint(x, y)` để quét xuyên qua các lớp DOM và phát hiện khối văn bản nằm bên dưới (thẻ `<p>`, `<h1>...<h6>`, `<code>`, `<textarea>`).
  3. Lấy nội dung văn bản tìm thấy và chuyển đến `speakRobotText()`.

## 3. Bộ máy phát âm thanh `RobotTTS`

Đối tượng `RobotTTS` là hạt nhân âm thanh của ứng dụng với các module chuyên trách.

### 3.1. Phiên âm công thức Toán học KaTeX/MathType: `cleanMarkdownForSpeech(text)`

Học sinh khi học lập trình thường gặp các công thức phức tạp. Bộ máy tự động chuyển đổi ký hiệu toán học sang lời đọc tiếng Việt chuẩn xác:

- `$x^2$` → “x bình phương”
- `\frac{a}{b}` → “phân số a trên b”
- `\sqrt{x}` → “căn bậc hai của x”
- `\le` hoặc `<=` → “nhỏ hơn hoặc bằng”
- `\ge` hoặc `>=` → “lớn hơn hoặc bằng”
- `\times` hoặc `*` → “nhân”
- `\div` hoặc `/` → “chia”

### 3.2. Nhận diện song ngữ Anh – Việt: `splitTextIntoBilingualPlaylist(text, maxLength)`

Trong một bài giảng lập trình Python, câu văn vừa có tiếng Việt vừa có các thuật ngữ và từ khóa kỹ thuật như `print`, `def`, `while`, `variable`, `syntax`.

- Nếu đọc toàn bộ bằng giọng tiếng Việt, các từ tiếng Anh sẽ bị đọc méo mó và sai ngữ âm.
- Hàm `isEnglishWord(word)` sử dụng bảng từ điển kỹ thuật lập trình và biểu thức chính quy (chứa các chữ cái `f`, `j`, `w`, `z`, chữ hoa viết tắt công nghệ như `AI`, `IDE`, `LMS`) để phân loại.
- Bộ máy chia câu thành các phân đoạn: đoạn tiếng Việt gán mã `tl=vi`, đoạn từ vựng tiếng Anh gán mã `tl=en`.

### 3.3. Cơ chế Preload và nối chuỗi âm thanh (Zero-Latency Audio Chaining)

Khi phát đoạn thứ *i*, hàm `playNext()` ngay lập tức tạo một đối tượng `new Audio()` ở chế độ nền để tải trước (preload) dữ liệu âm thanh của đoạn *i + 1*. Nhờ vậy, khi đoạn *i* kết thúc sự kiện `onended`, đoạn kế tiếp được phát ngay tức thì mà không có độ trễ giật lag.

## 4. Bảng ánh xạ nút bấm và các hàm gọi trong giao diện

| Chức năng | Thao tác trên giao diện | Hàm được gọi | Dữ liệu phát âm |
|---|---|---|---|
| Nghe giới thiệu tác giả | Bấm nút “Nghe Giới Thiệu” ở Banner tác giả | `speakAuthorIntro()` | Tiểu sử và thông tin dự án thầy Nguyễn Thế Phương |
| Học sinh nghe bài giảng | Bấm nút “AI Đọc Bài Giảng” | `speakStudentLecture(id)` | Toàn bộ tiêu đề và nội dung bài giảng hiện tại |
| Giáo viên nghe bài giảng | Bấm nút “Nghe AI” trong bảng quản lý | `previewTeacherLectureById(id)` | Nội dung bài giảng được chọn |
| Giáo viên nghe thử soạn thảo | Bấm nút “Nghe Thử AI Đọc” trong form soạn thảo | `previewCurrentTeacherLecture()` | Nội dung nhập trong ô tiêu đề và nội dung soạn |
| Nghe câu trả lời từ AI | Đặt câu hỏi trong hộp thoại Gia Sư AI | `speakRobotText(reply)` | Lời giải thích và định hướng Socratic của AI |
| Kéo thả Robot đọc chữ | Kéo biểu tượng Robot thả lên đoạn văn bất kỳ | `readTextAtRobotPosition(true)` | Nội dung thẻ văn bản nằm ngay bên dưới Robot |
| Bật/Tắt nhanh giọng đọc | Bấm chuột hoặc chạm vào Robot | `toggleRobotSpeechDirectly()` | Bật đọc văn bản hiện tại hoặc dừng phát ngay |

## 5. Ưu điểm kỹ thuật vượt trội

1. **Tính tương thích:** Hoạt động trên Google Chrome, Cốc Cốc, Safari, Edge, Firefox và các hệ điều hành Windows, macOS, iOS, Android.
2. **Âm điệu tự nhiên:** Giọng đọc trong trẻo, có ngữ điệu tự nhiên của tiếng Việt chuẩn, không bị giật tiếng như các bộ tổng hợp mặc định của hệ điều hành.
3. **Mở khóa tương tác âm thanh tự động:** Tích hợp cơ chế **Unlock Audio Context** ngay từ tương tác nhấp chuột đầu tiên của người dùng, giải quyết hạn chế “Autoplay Policy” trên các trình duyệt hiện đại.

---

*Được chuyển đổi từ tài liệu PDF đính kèm sang Markdown, giữ nguyên nội dung và cấu trúc chính của tài liệu nguồn.*
