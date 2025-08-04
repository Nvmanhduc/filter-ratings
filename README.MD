- Truy cập bảng xếp hạng cờ nhanh trên [chess.com](https://www.chess.com/vi/ratings)
- Lấy danh sách người chơi hàng đầu (top 50)
- Ghi dữ liệu vào Google Sheet (sheet `Tops`)
- Theo dõi các ô nhập điểm trong sheet (ô `B1`, `C1`, `D1`)
- Lọc và xuất danh sách người chơi có điểm số **lớn hơn hoặc bằng** giá trị được nhập
- Ghi kết quả lọc vào sheet `Result`

1. Tạo `credentials.json`

- Vào [Google Cloud Console](https://console.cloud.google.com/)
- Tạo service account → cấp quyền `Editor` hoặc `Sheets API + Drive API`
- Tạo khóa JSON → tải về → đặt tên là `credentials.json`

2. Run `python filter_ratings.py` 