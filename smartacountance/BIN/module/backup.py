import pandas as pd
import json
import os

def excel_to_json(excel_path, json_path):
    # Đọc tất cả các sheet
    excel_file = pd.read_excel(excel_path, sheet_name=None, engine='openpyxl')
    
    data_dict = {}
    for sheet_name, df in excel_file.items():
        # Xử lý các giá trị NaN/NaT thành None để JSON không bị lỗi
        df = df.where(pd.notnull(df), None)
        # Chuyển DataFrame thành list các dict (mỗi hàng là 1 object)
        data_dict[sheet_name] = df.to_dict(orient='records')
    
    # Ghi ra file JSON
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data_dict, f, ensure_ascii=False, indent=4)
    print(f"Đã chuyển đổi thành công sang: {json_path}")

# Sử dụng:
# excel_to_json('Data/dulieuduan.xlsx', 'Data/dulieuduan.json')

def json_to_excel(json_path, excel_path):
    with open(json_path, 'r', encoding='utf-8') as f:
        data_dict = json.load(f)
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        for sheet_name, data in data_dict.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Đã xuất file Excel thành công tại: {excel_path}")

# Sử dụng:
# json_to_excel('Data/dulieuduan.json', 'Data/dulieuduan_backup.xlsx')