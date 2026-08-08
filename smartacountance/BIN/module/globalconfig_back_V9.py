import os
import sys
from pathlib import Path
database_url = "localhost:5432"
system_status = "READY"
curent_team_name = ""
so_hop_dong_selected= ""
file_path = ""
tree_hd = None
tree_todoi = None
root = None
status_label = None
dlg_addhopdong = None
global_container = None
data = "Data"
ketoan_nhancongthuengoai = "nhan_cong_thue_ngoai"
file_thong_tin_du_an = "thong_tin_du_an"
ketoan_phathu = "ke_toan_phai_thu"

folder_nhancongthuengoai = ""
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
def khoitaothongso():
    folder_nhancongthuengoai_ = get_path(os.path.join("Data", ketoan_nhancongthuengoai))
    if not os.path.exists(folder_nhancongthuengoai_):
        os.makedirs(folder_nhancongthuengoai_, exist_ok=True)
    folder_nhancongthuengoai = folder_nhancongthuengoai_    
    print (f"folder_nhancongthuengoai_: {folder_nhancongthuengoai_}")
    return folder_nhancongthuengoai_
    