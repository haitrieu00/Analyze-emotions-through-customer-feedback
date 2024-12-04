# Hướng dẫn sử dụng Vietnamese Sentiment Analysis

## Yêu cầu

Trước khi sử dụng, vui lòng cài đặt các gói cần thiết bằng cách chạy lệnh sau:
pip install openpyxl


## Chuẩn bị mô hình

1. Chạy notebook Vietnamese Sentiment Analysis.ipynb để huấn luyện mô hình.
2. Sau khi hoàn thành, hai tập tin sau sẽ được lưu trữ:
   - models.h5
   - word.model
*Nếu đã có sẵn 2 file models.h5, word.model thì có thể chạy luôn file Using Vietnamese Sentiment Analysis.py mà không cần train lại


## Cấu hình đường dẫn

Trước khi chạy file Using Vietnamese Sentiment Analysis.py, bạn cần thay đổi ba đường dẫn sau đây để phản ánh thư mục trên máy của bạn:

1. 'D:\3. Study SIU\BigData\31012202538_11012202551\Code\models.h5'
2. 'D:\3. Study SIU\BigData\31012202538_11012202551\Code\word.model'
3. 'D:\3. Study SIU\BigData\31012202538_11012202551\Misc\text.png'


## Sử dụng giao diện

- Nhập văn bản vào ô văn bản và nhấn "Đánh giá" để hiển thị kết quả cảm xúc.
- Khi sử dụng nút "Upload file", chọn file Test.xlsx. Mô hình sẽ gán nhãn cho các mẫu trong file và hiển thị một cửa sổ để chọn vị trí để lưu file đã được dự đoán.
