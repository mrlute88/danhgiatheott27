# TT27 Assessor AI 🎓
> **Trợ lý AI Hỗ trợ Giáo viên Viết Nhận xét Học sinh Tiểu học theo Thông tư 27/2020/TT-BGDĐT**

---

## 📌 Giới thiệu Tổng quan

**TT27 Assessor AI** là hệ thống trợ lý sư phạm thông minh ứng dụng kiến trúc **Đa tác tử (Multi-Agent System)**, giúp giáo viên tiểu học tại Việt Nam tiết kiệm tới 90% thời gian đánh giá học sinh cuối kỳ và giữa kỳ:
- ✅ **Chuẩn hóa Thông tư 27/2020/TT-BGDĐT**: Tự động quy đổi điểm số và phân loại 3 mức đánh giá định kỳ: **Mức T (Hoàn thành tốt)**, **Mức H (Hoàn thành)**, **Mức C (Chưa hoàn thành)**.
- ✅ **Cá nhân hóa sâu sắc**: Lồng ghép hài hòa năng lực cốt lõi từng môn học (Toán, Tiếng Việt, Tiếng Anh, TN&XH, Khoa học, Sử Địa, Tin học...) và ghi chú riêng của giáo viên.
- ✅ **Bảo toàn 100% file Excel**: Giữ nguyên toàn bộ cấu trúc bảng điểm, màu sắc, font chữ, đường viền (border) và công thức tính toán ban đầu của trường lớp.
- ✅ **Hoạt động linh hoạt (Online & Offline)**: Hỗ trợ kết nối Google Gemini API (gemini-3.6-flash) hoặc tự động chạy ở chế độ **Ngân hàng Sư phạm Offline thông minh** khi không có internet/API key.

---

## 🤖 Kiến trúc Đa Tác tử (Multi-Agent Architecture)

```
       [ Bảng điểm Excel / Dữ liệu Học sinh ]
                         │
                         ▼
┌────────────────────────────────────────────────────────┐
│  🤖 Agent 1: Data Classifier (agents/classifier.py)    │
│  - Chuẩn hóa mức đạt T / H / C theo TT27               │
│  - Ánh xạ ma trận năng lực chuyên biệt theo môn học    │
│  - Xây dựng chỉ dẫn định hướng sư phạm cá nhân hóa     │
└────────────────────────┬───────────────────────────────┘
                         │
                         ▼
┌────────────────────────────────────────────────────────┐
│  🧠 Agent 2: TT27 Commentator (agents/commentator.py)  │
│  - Khởi tạo nhận xét bằng Google Gemini AI             │
│  - Cơ chế tự động Fallback sang Ngân hàng Sư phạm      │
│  - Exponential backoff xử lý giới hạn tốc độ gọi API   │
└────────────────────────┬───────────────────────────────┘
                         │
                         ▼
┌────────────────────────────────────────────────────────┐
│  🛡️ Agent 3: Output Validator (agents/validator.py)   │
│  - Khử tiền tố AI, chuẩn hóa độ dài (15-35 từ)         │
│  - Chống lặp câu liên tiếp qua cửa sổ bộ đệm đa dạng   │
│  - Chống tiêm nhiễm công thức Excel (=, +, -, @)       │
└────────────────────────┬───────────────────────────────┘
                         │
                         ▼
[ Bảng duyệt trực tiếp & Tải về File Excel Hoàn chỉnh (.xlsx) ]
```

---

## 🚀 Hướng dẫn Cài đặt & Chạy Ứng dụng

### 1. Cài đặt thư viện cần thiết
```bash
pip install -r requirements.txt
```

### 2. Khởi chạy ứng dụng Streamlit
```bash
streamlit run app.py
```
Sau khi chạy lệnh, trình duyệt web sẽ tự động mở tại địa chỉ: `http://localhost:8501`.

---

## 💡 Hướng dẫn Sử dụng Nhanh

1. **Trải nghiệm ngay trong 1 click:**
   - Tại Tab 1, bấm nút **"Nạp dữ liệu mẫu 15 học sinh để thử ngay"** hoặc bấm **"Tải file Excel mẫu chuẩn TT27"**.
2. **Nhập bảng điểm lớp bạn:**
   - Kéo thả file Excel (`.xlsx`, `.xls`) vào khung tải lên.
   - Hệ thống tự động nhận diện dòng tiêu đề trường học và ánh xạ các cột: *Họ và tên*, *Điểm số/Mức đạt*, *Ghi chú giáo viên*.
3. **Cấu hình & Tạo nhận xét:**
   - Chọn Môn học, Khối lớp, Thời điểm đánh giá (Giữa HK1, Cuối HK1, Giữa HK2, Cuối HK2).
   - Chọn Phong cách lời phê (Chuẩn mực, Ấm áp, Cụ thể hóa kỹ năng, Ngắn gọn).
   - Bấm **"🚀 BẮT ĐẦU TẠO NHẬN XÉT"**.
4. **Duyệt & Tải file:**
   - Xem thống kê tỷ lệ Mức T, H, C.
   - Chỉnh sửa trực tiếp bất kỳ lời phê nào trên bảng tương tác.
   - Bấm **"📥 TẢI VỀ FILE EXCEL HOÀN CHỈNH (.XLSX)"**.

---

## 📄 Tính Năng Mới: Chuyển Đổi Ảnh / PDF sang Word (.docx) (Vision AI)

Hệ thống tích hợp công nghệ Thị giác máy tính đa phương thức (**Multimodal Vision**) từ Google Gemini:
- 🔍 **OCR Đa định dạng**: Nhận diện chữ viết tiếng Việt có dấu từ ảnh chụp, ảnh scan (`.png`, `.jpg`, `.jpeg`, `.webp`, `.bmp`) và tài liệu PDF (`.pdf`).
- 📊 **Tái lập Bảng Word thực tế**: Tự động chuyển đổi các bảng số liệu, bảng điểm, ma trận thành bảng Word (`docx.Table`) có màu nền header, đường viền thanh lịch, căn lề chuẩn.
- 📝 **Chế độ chuyên biệt cho Giáo dục**:
  - *Đề thi & Phiếu bài tập*: Nhận diện rõ câu hỏi trắc nghiệm A-B-C-D, bài tự luận, biểu thức toán học.
  - *Giáo án & Kế hoạch bài dạy*: Nhận diện cấu trúc bảng 2 cột Hoạt động GV & Hoạt động HS.
  - *Công văn & Quyết định hành chính*: Định dạng chuẩn thể thức theo Nghị định 30/2020/NĐ-CP.
  - *Trích xuất Bảng số liệu*: Tối ưu quét bảng danh sách, bảng điểm.
- ⚙️ **Tùy chỉnh định dạng Word**: Font chữ (Times New Roman, Calibri, Arial), cỡ chữ (12, 13, 14pt), lề chuẩn văn bản (Trái 2.5cm, các lề còn lại 2.0cm), đánh số trang tự động ở chân trang.
- ✏️ **Biên tập & Chỉnh sửa trực tiếp**: Xem trước văn bản, sửa đổi nội dung trực tiếp trên giao diện và tái tạo file Word mới tức thì.
- 🧪 **Trải nghiệm 1-Click**: Có sẵn nút nạp Đề kiểm tra mẫu dạng Ảnh và dạng PDF để thử nghiệm ngay lập tức.

