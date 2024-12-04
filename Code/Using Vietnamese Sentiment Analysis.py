# Import thư viện
import re
import underthesea
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog
from PIL import Image, ImageTk
from keras.models import load_model
import gensim.models.keyedvectors as word2vec


# Load Model đã train
model_sentiment = load_model('models.h5')
# Load word embedding model
model_embedding = word2vec.KeyedVectors.load('word.model')


# Tiền xử lý dữ liệu
def preprocess_text(text):
    # Chuẩn hóa về chữ thường
    text = text.lower()
    # Thay thế các URL bằng nhãn 'link_spam'
    text = re.sub(r'http\S+', 'link_spam', text)
    # Tách từ
    words = underthesea.word_tokenize(text)
    # Loại bỏ dấu câu và các ký tự đặc biệt
    words = [re.sub(r'[^\w\s]', '', word) for word in words]
    words = [re.sub(r'(\w)\1+', r'\1', word) for word in words]
    # Chuẩn hóa các từ viết tắt cơ bản
    abbreviation_dict = {
        'k': 'không',
        'ko': 'không',
        'k0': 'không',
        'khong': 'không',
        'khôg': 'không',
        'bt': 'bình thường'
    }
    words = [abbreviation_dict.get(word, word) for word in words]
    # Loại bỏ số và các từ chỉ có 1 ký tự
    words = [word for word in words if not (word.isdigit() or len(word) == 1)]
    # Ghép lại các từ thành câu
    cleaned_text = ' '.join(words)
    
    return cleaned_text


# Nhúng các từ trong một bình luận thành ma trận nhúng
word_labels = []
max_seq = 200
embedding_size = 128

for word in model_embedding.index_to_key:
    word_labels.append(word)
    
def comment_embedding(comment):
    matrix = np.zeros((max_seq, embedding_size))
    words = comment.split()
    
    try:
        lencmt = len(words)
        
        if lencmt == 0:
            sentiment_label.config(text="Không đoán được")
            print("Không đoán được")
            return "Không đoán được"
        
        for i in range(max_seq):
            indexword = i % lencmt
            if (max_seq - i < lencmt):
                break
            if(words[indexword] in word_labels):
                matrix[i] = model_embedding[words[indexword]]
        matrix = np.array(matrix)
        return matrix
    except ZeroDivisionError:
        print("Lỗi: Không thể chia cho 0")
        return "Không đoán được" 


# Hàm xử lý sự kiện khi nút "Upload File" được nhấn
def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])  # Mở hộp thoại chọn file Excel
    if file_path:
        try:
            df = pd.read_excel(file_path)  # Đọc dữ liệu từ file Excel
            
            # Thêm cột "Sentiment" vào DataFrame và đánh giá cảm xúc cho từng dòng
            df['Sentiment'] = df['Comment'].apply(lambda comment: evaluate_comment_sentiment(comment))
            
            # Lưu kết quả vào file Excel mới
            output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
            if output_file_path:
                df.to_excel(output_file_path, index=False)
                messagebox.showinfo("Thông báo", f"Đã lưu kết quả vào file: {output_file_path}")
            else:
                messagebox.showwarning("Cảnh báo", "Vui lòng chọn nơi lưu file.")
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra: {str(e)}")


# Đánh giá cảm xúc của file Excel
def evaluate_comment_sentiment(comment):
    text = preprocess_text(comment)
    maxtrix_embedding = np.expand_dims(comment_embedding(text), axis=0)
    maxtrix_embedding = np.expand_dims(maxtrix_embedding, axis=3)
    result = model_sentiment.predict(maxtrix_embedding)
    result = np.argmax(result)
    return "Tích cực" if result == 1 else "Tiêu cực" if result == 2 else "Bình thường"


# Xử lý sự kiện sừ dụng text trực tiếp
def evaluate_sentiment():
    text = text_entry.get("1.0","end-1c") # Lấy văn bản từ input
    if text.strip() == "":
        messagebox.showwarning("Warning", "Vui lòng nhập văn bản!")
        return
    
    text = preprocess_text(text)
    maxtrix_embedding = np.expand_dims(comment_embedding(text), axis=0)
    maxtrix_embedding = np.expand_dims(maxtrix_embedding, axis=3)

    result = model_sentiment.predict(maxtrix_embedding)
    result = np.argmax(result)
    
    if result == 1:
        sentiment_label.config(text="Tích cực")
    elif result == 2:
        sentiment_label.config(text="Tiêu cực")
    else:
        sentiment_label.config(text="Bình thường")


# Tạo giao diện đồ họa
app = tk.Tk()
app.title("Phân tích cảm xúc")
app.geometry("400x300")

text_label = tk.Label(app, text="Nhập văn bản:")
text_label.pack()

text_entry = tk.Text(app, height=5, width=48)
text_entry.pack()

evaluate_button = tk.Button(app, text="Đánh giá", command=evaluate_sentiment)
evaluate_button.pack()

sentiment_label = tk.Label(app, text="", font=("Helvetica", 14), fg="purple")
sentiment_label.pack()

# Thêm nút "Upload File" vào giao diện đồ họa
upload_button = tk.Button(app, text="Upload File", command=upload_file)
upload_button.pack()

# Load hình ảnh
image = Image.open('text.png')

# Chuyển đổi hình ảnh thành đối tượng hình ảnh của tkinter
photo = ImageTk.PhotoImage(image)

# Tạo widget Label để chứa hình ảnh
image_label = tk.Label(app, image=photo)
image_label.pack()

# Xác định hành động khi cửa sổ đóng lại
def on_closing():
    app.destroy()

# Gọi hàm on_closing() khi cửa sổ đóng lại
app.protocol("WM_DELETE_WINDOW", on_closing) 

app.mainloop()
