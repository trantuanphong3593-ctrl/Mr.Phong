# Giọng AI cho PRACTICE ENGLISH BY TTP

Thư mục này chứa file âm thanh đọc sẵn (mp3, mono) cho từng từ/câu trong ngân hàng câu hỏi.

- `manifest.json`: danh sách giọng và mã băm (hash) của các câu đã có file.
- `<giọng>/<mã>.mp3`: bản đọc thường; `<giọng>/s/<mã>.mp3`: bản đọc chậm.
- Giọng: `us_f` (Mỹ · Nữ), `us_m` (Mỹ · Nam), `uk_f` (Anh · Nữ), `uk_m` (Anh · Nam).
- Trang web tự dùng file ở đây khi học sinh chọn "Giọng AI tự nhiên"; câu nào chưa có file thì đọc bằng giọng có sẵn của thiết bị.

Mô hình tạo giọng: Kokoro v1.0 (giấy phép Apache-2.0), chạy bằng kokoro-onnx.
Giọng gốc: af_heart, am_michael, bf_emma, bm_george. Đây là giọng tổng hợp bằng AI, không phải giọng người thật.
