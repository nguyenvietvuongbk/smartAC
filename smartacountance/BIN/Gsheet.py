import gspread
import os  # <--- THIẾU DÒNG NÀY ĐÂY
from oauth2client.service_account import ServiceAccountCredentials

# 1. Xác định thư mục chứa file Python hiện tại
current_dir = os.path.dirname(os.path.abspath(__file__))
# 2. Xây dựng đường dẫn đến file nằm trong thư mục 'cloundkey'
# Cấu trúc: current_dir -> cloundkey -> smartacountance-a0f266ed648d.json
key_path = os.path.join(current_dir, 'cloundkey', 'smartacountance-a0f266ed648d.json')
# Cấu hình xác thực
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(key_path, scope)
client = gspread.authorize(creds)

# Mở Google Sheet theo tên
sheet = client.open('smartAC').sheet1

# Ghi dữ liệu vào ô A1
sheet.update_cell(1, 1, 'Hello World!')

# Đọc toàn bộ dữ liệu
data = sheet.get_all_records()
print(data)