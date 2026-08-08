import subprocess
import os
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.hazmat.primitives import serialization

import tkinter as tk
from tkinter import filedialog, messagebox
import base64 # Đảm bảo đã import thư viện này
import sys
import winreg
from datetime import datetime, timedelta
from pathlib import Path
REG_PATH = r"Software\smartAC" # Tên ứng dụng của bạn trong Registry
VALUE_NAME = "FirstInstallDate"

def get_project_root():
    # 1. Xác định vị trí file thực thi hoặc file script hiện tại
    if getattr(sys, 'frozen', False):
        # Đang chạy .exe: Root chính là thư mục chứa .exe
        return Path(sys.executable).parent
    else:
        # Đang chạy .py: Bắt đầu từ vị trí file hiện tại
        current_path = Path(__file__).resolve().parent.parent.parent          
        return current_path

def get_path(relative_path):
    """Hàm tiện ích gom chung cho mọi loại đường dẫn"""
    print(f"root: {get_project_root()}")
    base_dir = get_project_root()
    
    path_at_root = os.path.join(base_dir, relative_path)
    path_at_internal = os.path.join(base_dir, "_internal", relative_path)
    if os.path.exists(path_at_root):
        return path_at_root
    # Nếu không, kiểm tra xem nó có bị đẩy vào _internal không
    if os.path.exists(path_at_internal):
        return path_at_internal    
        
    return path_at_root
    
def get_trial_status():
    try:
        # Mở hoặc tạo Key Registry
        key = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_ALL_ACCESS)
        
        # 1. Kiểm tra xem đã từng bị "đánh dấu quá hạn" chưa
        try:
            expired_flag, _ = winreg.QueryValueEx(key, "TrialExpired")
            if expired_flag == "1":
                return False, 0 # Đã từng quá hạn, cấm vĩnh viễn
        except FileNotFoundError:
            pass # Chưa có cờ này, tiếp tục kiểm tra ngày tháng

        # 2. Kiểm tra ngày cài đặt
        try:
            date_str, _ = winreg.QueryValueEx(key, "FirstInstallDate")
            start_date = datetime.strptime(date_str, "%Y-%m-%d")
        except FileNotFoundError:
            start_date = datetime.now()
            winreg.SetValueEx(key, VALUE_NAME, 0, winreg.REG_SZ, start_date.strftime("%Y-%m-%d"))
        
        # 3. Tính toán chênh lệch ngày
        delta = datetime.now() - start_date
        
        # PHÁT HIỆN GIAN LẬN:
        # - delta.days < 0: Người dùng chỉnh đồng hồ lùi lại so với lúc cài
        # - delta.days > 30: Đã quá hạn 30 ngày
        if delta.days < 0 or delta.days > 30:
            # Ghi dấu "vĩnh viễn" vào Registry
            winreg.SetValueEx(key, "TrialExpired", 0, winreg.REG_SZ, "1")
            return False, 0
            
        remaining_days = 30 - delta.days
        return True, remaining_days # Còn hạn

    except Exception:
        # Nếu có lỗi (quyền truy cập, registry lỗi...), an toàn nhất là trả về False (không cho dùng)
        return False, 0
        

def get_hwid():
    """Lấy Hardware ID từ CPU và Mainboard"""
    try:
        cpu = subprocess.check_output("wmic cpu get processorid", shell=True).decode().split('\n')[1].strip()
        mb = subprocess.check_output("wmic baseboard get serialnumber", shell=True).decode().split('\n')[1].strip()
        return f"{cpu}-{mb}"
    except:
        return "UNKNOWN_HWID"

def verify_and_get_expiry(license_path, public_key_path):
    try:
        # 1. Load Public Key
        with open(public_key_path, "rb") as key_file:
            public_key = serialization.load_pem_public_key(key_file.read())

        # 2. Đọc file license
        with open(license_path, "rb") as f:
            content = f.read()

        # 3. Tách dữ liệu
        parts = content.split(b"|SIGN|") 
        if len(parts) != 2:
            print("File license không đúng định dạng!")
            return None
            
        data = parts[0] # Đây là b"HWID|Expiry"
        signature_b64 = parts[1] # Đây là chuỗi Base64 (ví dụ: b"SGVsbG8...")

        # 4. QUAN TRỌNG: Giải mã Base64 về dạng raw bytes
        signature = base64.b64decode(signature_b64)

        # 5. Xác thực (Sử dụng signature đã được giải mã)
        public_key.verify(
            signature,
            data,
            padding.PSS(mgf=padding.MGF1(hashes.SHA256()), salt_length=padding.PSS.MAX_LENGTH),
            hashes.SHA256()
        )
        
        # 6. Nếu verify thành công, trả về ngày hết hạn
        decoded_data = data.decode().split("|")
        return decoded_data[1] 

    except Exception as e:
        print(f"Lỗi xác thực chi tiết: {e} _ {public_key_path}")
        return None
def check_license_and_launch(root, status_label):
    license_path = "Data/license.lic"
    public_key_path = get_path(os.path.join("public key","public_key.pem"))
    
    # 1. Kiểm tra License File
    if os.path.exists(license_path):
        expiry = verify_and_get_expiry(license_path, public_key_path)
        if expiry:
            status_label.config(text=f"Bản quyền: Đã kích hoạt (Hết hạn: {expiry})", fg="green")
            return True
    
    # 2. Kiểm tra Trial
    is_trial, days_left = get_trial_status()
    if is_trial:
        status_label.config(text=f"Trạng thái: Dùng thử ({days_left} ngày còn lại)", fg="#0066cc") # Màu xanh dương
        return True
    
    # 3. Không hợp lệ
    status_label.config(text="Trạng thái: Chưa kích hoạt - Yêu cầu bản quyền", fg="red")
    return show_activation_dialog(root, public_key_path)

def show_activation_dialog(parent, public_key_path):
    dialog = tk.Toplevel(parent)
    dialog.title("Kích hoạt bản quyền")
    dialog.geometry("400x250")

    hwid = get_hwid()
    
    tk.Label(dialog, text="Product ID (Hardware Fingerprint):").pack(pady=10)
    entry_hwid = tk.Entry(dialog, width=50)
    entry_hwid.insert(0, hwid)
    entry_hwid.config(state='readonly')
    entry_hwid.pack()

    lbl_expiry = tk.Label(dialog, text="Ngày hết hạn: Chưa kích hoạt", fg="red")
    lbl_expiry.pack(pady=20)

    def load_license():
        file_path = filedialog.askopenfilename(filetypes=[("License files", "*.lic")])
        if file_path:
            expiry = verify_and_get_expiry(file_path, public_key_path)
            if expiry:
                # Copy file vào thư mục Data để lưu lại
                if not os.path.exists("Data"): os.makedirs("Data")
                import shutil
                shutil.copy(file_path, "Data/license.lic")
                
                lbl_expiry.config(text=f"Kích hoạt thành công! Hết hạn: {expiry}", fg="green")
                messagebox.showinfo("Thông báo", "Đã kích hoạt thành công!")
            else:
                messagebox.showerror("Lỗi", "License không hợp lệ cho thiết bị này!")

    btn_load = tk.Button(dialog, text="Tải License File", command=load_license)
    btn_load.pack()        