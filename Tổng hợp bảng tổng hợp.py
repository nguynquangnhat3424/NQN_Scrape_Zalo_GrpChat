import pandas as pd
from openpyxl import load_workbook
import re

# Đọc file Excel và bỏ cột đầu, cột cuối
file_path = r"C:\Users\Welcome\Documents\VDSC\Task\Task16 - EDA data zalo\Summary_Zalo.xlsx"
df = pd.read_excel(file_path)
df = df.iloc[:, 1:-1]  # Bỏ cột đầu và cột cuối

# Các label cần tìm (dùng chữ thường để so sánh)
labels = ['tích cực', 'tiêu cực', 'trung lập']

# Hàm để kiểm tra và tách label từ chuỗi, áp dụng lower để không phân biệt hoa thường
def extract_label(text):
    if isinstance(text, str):  # Kiểm tra xem giá trị có phải là chuỗi không
        text_lower = text.lower()  # Chuyển toàn bộ chuỗi về chữ thường
        for label in labels:
            if label in text_lower:
                return label.capitalize()  # Trả về nhãn với chữ cái đầu in hoa
    return None  # Trường hợp giá trị không phải chuỗi hoặc không tìm thấy nhãn nào

# Áp dụng hàm để tách label cho từng ô trong bảng
df_labels = df.applymap(extract_label)

# Thực hiện group by bằng cách đếm số lần xuất hiện của mỗi label
grouped_df = df_labels.apply(lambda col: col.value_counts()).fillna(0).astype(int)

# Sắp xếp các dòng theo thứ tự "Tích cực", "Tiêu cực", "Trung lập"
ordered_labels = ['Tích cực', 'Tiêu cực', 'Trung lập']
grouped_df = grouped_df.reindex(ordered_labels)

# Đọc file Excel và chỉ lấy cột cuối cùng
file_path = r"C:\Users\Welcome\Documents\VDSC\Task\Task16 - EDA data zalo\Summary_Zalo.xlsx"
df = pd.read_excel(file_path)
df_last_column = df.iloc[:, -1]  # Lấy cột cuối cùng

# Hàm để trích xuất các mã cổ phiếu từ chuỗi "Tích cực: HPG, VHM. Tiêu cực: DIG, PDR."
def extract_stocks(text):
    stock_data = {'Tích cực': [], 'Tiêu cực': []}
    
    if isinstance(text, str):  # Kiểm tra nếu là chuỗi
        text = text.lower()  # Chuyển chuỗi về chữ thường để không phân biệt hoa thường
        
        # Tìm các phần liên quan đến "Tích cực" và "Tiêu cực"
        positive_match = re.search(r'tích cực:([^.]*)', text)
        negative_match = re.search(r'tiêu cực:([^.]*)', text)
        
        # Nếu tìm thấy phần "Tích cực", tách mã cổ phiếu
        if positive_match:
            stocks_positive = re.findall(r'\b[a-z]{3}\b', positive_match.group(1))
            stock_data['Tích cực'] = [stock.upper() for stock in stocks_positive]
        
        # Nếu tìm thấy phần "Tiêu cực", tách mã cổ phiếu
        if negative_match:
            stocks_negative = re.findall(r'\b[a-z]{3}\b', negative_match.group(1))
            stock_data['Tiêu cực'] = [stock.upper() for stock in stocks_negative]
    
    return stock_data

# Tạo bảng để lưu trữ số lần xuất hiện các mã cổ phiếu theo nhãn
stock_count = {'Tích cực': {}, 'Tiêu cực': {}}

# Duyệt qua từng dòng và trích xuất mã cổ phiếu
for entry in df_last_column:
    stock_data = extract_stocks(entry)
    
    # Cập nhật số lần xuất hiện cho "Tích cực"
    for stock in stock_data['Tích cực']:
        if stock in stock_count['Tích cực']:
            stock_count['Tích cực'][stock] += 1
        else:
            stock_count['Tích cực'][stock] = 1
    
    # Cập nhật số lần xuất hiện cho "Tiêu cực"
    for stock in stock_data['Tiêu cực']:
        if stock in stock_count['Tiêu cực']:
            stock_count['Tiêu cực'][stock] += 1
        else:
            stock_count['Tiêu cực'][stock] = 1

# Tạo bảng tổng hợp với các mã cổ phiếu xuất hiện từ 2 lần trở lên
summary_data = {'Label': ['Tích cực', 'Tiêu cực']}

# Duyệt qua từng nhãn và thêm các mã cổ phiếu xuất hiện từ 2 lần trở lên vào bảng
for label in ['Tích cực', 'Tiêu cực']:
    for stock, count in stock_count[label].items():
        if count >= 2:
            if stock not in summary_data:
                summary_data[stock] = [0, 0]  # Khởi tạo với 2 dòng cho "Tích cực" và "Tiêu cực"
            summary_data[stock][0 if label == 'Tích cực' else 1] = count

# Chuyển dữ liệu thành DataFrame
summary_df = pd.DataFrame(summary_data)

# In kết quả
print(summary_df)


# Ghi cả 2 bảng tổng hợp vào file Excel gốc trong một sheet mới
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:  # Mở file gốc để thêm sheet
    grouped_df.to_excel(writer, sheet_name='Summary_Labels', startrow=0, index=True)
    summary_df.to_excel(writer, sheet_name='Summary_MACP', startrow = 0, index=False)

print("Đã thêm hai bảng tổng hợp vào sheet mới trong file Excel.")
