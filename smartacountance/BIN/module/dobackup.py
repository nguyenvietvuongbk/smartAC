import os
import shutil
import re

def backup_utils_file(dir_, preword):
    # Lấy thư mục hiện tại
    current_dir = dir_#os.getcwd()
    #parent_dir = os.path.dirname(current_dir)
    # Regex để tìm file có chữ 'utils' và số ở cuối (trước phần mở rộng)
    # Ví dụ: utils_1.py, my_utils_10.txt
    # (\d+) sẽ bắt con số ở cuối
    pattern_str = fr"{preword}.*?(\d+)\.[^.]+$"

    # Bước 2: Compile chuỗi đó
    pattern = re.compile(pattern_str)
    
    max_num = -1
    target_file = None
    
    # Duyệt qua các file trong thư mục
    for filename in os.listdir(current_dir):
        
        match = pattern.search(filename)
        if match:
            print(f"file in {current_dir}: {filename}")
            # Lấy số từ tên file
            num = int(match.group(1))
            
            # Cập nhật số lớn nhất và file tương ứng
            if num > max_num:
                max_num = num
                target_file = filename
    
    # Kiểm tra nếu tìm thấy file
    if target_file:
        # Tính số mới
        new_num = max_num + 1
         
        # Lấy phần mở rộng của file gốc (ví dụ: .py, .txt)
        file_extension = os.path.splitext(target_file)[1]
        target_file = os.path.join(dir_, preword + file_extension)
        # Tạo tên file mới
        new_filename = os.path.join(dir_, f"{preword}_back_V{new_num}{file_extension}")
        
        # Thực hiện copy
        shutil.copy2(target_file, new_filename)
        print(f"Đã copy thành công: '{target_file}' -> '{new_filename}'")
    else:
        print(f"Không tìm thấy file nào có định dạng '{preword}' kèm số ở cuối trong thư mục này.")

if __name__ == "__main__":
    current_dir_ = os.getcwd()
    parent_dir_ = os.path.dirname(current_dir_)
    print(f"parent_dir_: {parent_dir_}")
    backup_utils_file(current_dir_,"utils")
    backup_utils_file(current_dir_,"license_utils")
    backup_utils_file(current_dir_,"globalconfig")
    backup_utils_file(parent_dir_,"smartAC")