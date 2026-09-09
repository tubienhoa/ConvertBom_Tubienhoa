# BOM Converter — Kỹ thuật ➜ ERP (Stovax)

Công cụ web (Streamlit) giúp chuyển đổi **BOM kỹ thuật** (dạng cây: Cụm > Chi tiết > Vật tư,
có kích thước/khối lượng/vật liệu) sang **BOM ERP** (dạng bảng phẳng theo cột
`STT, QTSX, Mã sản phẩm, Loại, Tên sản phẩm, SL sản phẩm, ĐVT sp, Mã cũ, Mã NVL, Tên NVL,
ĐVT (NVL), Số lượng, ĐVT Kho, Số lượng kho, Số lượng setup, % Tỉ lệ hao hụt, Công đoạn`).

## ⚠️ Giới hạn cần biết

BOM kỹ thuật chỉ có thông tin **thiết kế**. BOM ERP cần thêm thông tin **sản xuất**
(quy trình QTSX, công đoạn, có khi 1 chi tiết tách nhiều bước gia công) — thông tin
này **không có sẵn** trong file kỹ thuật. Công cụ sẽ:

- Tự dựng đúng cấu trúc phân cấp (FG > SG > PT), mã hoá theo quy ước `Loại - KKK - Mã`.
- Tự lấy vật liệu / khối lượng / số lượng từ file kỹ thuật.
- **Đoán** công đoạn theo từ khoá trong mô tả (có thể chỉnh trong thanh bên của app).
- Để trống hoặc giữ tên gốc những chỗ không đủ dữ liệu để suy luận chắc chắn
  (ví dụ tên tiếng Anh của một số chi tiết trung gian).

➡️ Kết quả xuất ra là **bản nháp**, bảng có thể sửa trực tiếp trong app trước khi tải về,
kỹ sư/QC vẫn nên rà soát lại trước khi nạp vào ERP thật.

## Cấu trúc project

```
bom_converter/
├── app.py            # Giao diện Streamlit
├── converter.py       # Logic đọc & chuyển đổi (dùng lại được ngoài Streamlit)
├── requirements.txt
└── README.md
```

## Chạy thử ở máy local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Mở trình duyệt tại địa chỉ Streamlit in ra (thường là http://localhost:8501),
tải lên file BOM kỹ thuật (.xlsx, sheet tên "Form") và bấm "Chuyển đổi".

## Đưa lên GitHub & deploy Streamlit Community Cloud

Vì bạn đã có tài khoản GitHub và Streamlit (đã liên kết với GitHub), làm theo các bước sau:

1. **Tạo repo mới trên GitHub** (ví dụ tên `bom-converter`), để trống (không tạo README/license
   để tránh xung đột khi push).

2. **Push code lên repo đó** (chạy trên máy bạn, trong thư mục vừa tải về từ đây):
   ```bash
   cd bom_converter
   git init
   git add .
   git commit -m "Initial commit: BOM converter tool"
   git branch -M main
   git remote add origin https://github.com/<username>/bom-converter.git
   git push -u origin main
   ```

3. **Deploy trên Streamlit Community Cloud**:
   - Vào https://share.streamlit.io/ (đăng nhập bằng tài khoản đã liên kết GitHub).
   - Bấm **"New app"**.
   - Chọn repo `bom-converter`, branch `main`, file chính là `app.py`.
   - Bấm **Deploy**. Sau khoảng 1–2 phút, bạn sẽ có một URL công khai (hoặc riêng tư nếu
     đặt private) để chia sẻ cho đồng nghiệp dùng ngay trên trình duyệt, không cần cài gì.

4. Mỗi khi bạn `git push` cập nhật code, Streamlit Cloud sẽ tự động build lại app.

## Tuỳ biến thêm

- Muốn đổi shortcode khách hàng mặc định, tiền tố mã vẽ, hay bảng từ khoá đoán công đoạn:
  sửa trực tiếp trong `converter.py` (biến `DEFAULT_RULES`) hoặc chỉnh ngay trong thanh bên
  của app lúc chạy (không cần sửa code).
- Nếu cấu trúc file BOM kỹ thuật của bạn có sheet tên khác "Form", hoặc số dòng tiêu đề
  khác so với mẫu, có thể cần chỉnh hàm `load_technical_bom()` trong `converter.py`
  (phần dò dòng "Khách hàng:", "Dự án:", và dòng tiêu đề bảng "MỤC").
"# ConvertBom_Tubienhoa" 
