def nhap_thong_tin_du_an(parent):
    # 1. Cấu hình đường dẫn lưu file JSON
    folder_path = get_path(f"Data/{globalconfig.ketoan_nhancongthuengoai}")
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    
    file_path = get_path(os.path.join(folder_path, f"{globalconfig.file_thong_tin_du_an}.JSON"))

    # 2. Tạo cửa sổ Modal
    dialog = tk.Toplevel(parent)
    dialog.title("Khai báo thông tin dự án")
    dialog.geometry("650x550")
    dialog.grab_set()
    dialog.transient(parent)

    # Danh sách các trường thông tin theo bảng
    field_names = [
        "Tên công ty", 
        "Tên công trình", 
        "Giám đốc", 
        "Chỉ huy trưởng", 
        "Kế toán trưởng", 
        "Người lập biểu", 
        "Địa điểm", 
        "Năm thực hiện",
        "Thu nhập năm miễn thuế"
    ]
    
    entries = {}

    # 3. Tạo khung giao diện nhập liệu
    main_frame = ttk.Frame(dialog, padding="20")
    main_frame.pack(fill="both", expand=True)

    ttk.Label(main_frame, text="THÔNG TIN CHUNG DỰ ÁN", font=("Arial", 14, "bold")).grid(row=0, column=0, columnspan=2, pady=(0, 20))

    # Đọc dữ liệu cũ từ file JSON nếu có (dùng .get() để chống lỗi thiếu trường dữ liệu)
    old_data = {}
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                old_data = json.load(f)
        except Exception:
            old_data = {}

    # Vòng lặp tạo Label và Entry theo dạng 1 cột dọc trực quan, dễ điền
    for idx, field in enumerate(field_names):
        row = idx + 1
        
        ttk.Label(main_frame, text=field, font=("Arial", 10)).grid(row=row, column=0, sticky="w", pady=6, padx=5)
        entry = ttk.Entry(main_frame, width=45, font=("Arial", 10))
        
        # Điền dữ liệu cũ an toàn, nếu thiếu trường sẽ không báo lỗi mà để trống
        val = old_data.get(field, "")
        if val:
            entry.insert(0, str(val))
            
        entry.grid(row=row, column=1, pady=6, padx=5, sticky="ew")
        entries[field] = entry

    # Cấu hình co giãn cột trong khung
    main_frame.columnconfigure(1, weight=1)

    # 4. Hàm xử lý lưu thông tin vào JSON
    def save_project_info():
        # Lấy dữ liệu và loại bỏ khoảng trắng thừa
        new_data = {field: entries[field].get().strip() for field in field_names}
        
        # Kiểm tra không được để trống các trường bắt buộc
        if not new_data["Tên công ty"] or not new_data["Tên công trình"]:
            messagebox.showerror("Lỗi", "Tên công ty và Tên công trình không được để trống!", parent=dialog)
            return

        # Kiểm tra năm thực hiện phải là số (nếu có nhập)
        nam_thuc_hien = new_data["Năm thực hiện"]
        if nam_thuc_hien and not nam_thuc_hien.isdigit():
            messagebox.showerror("Lỗi định dạng", "Trường 'Năm thực hiện' phải là định dạng số (Ví dụ: 2026)!", parent=dialog)
            entries["Năm thực hiện"].focus()
            return

        # Ghi dữ liệu vào file JSON
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, ensure_ascii=False, indent=4)
        except Exception as ex:
            messagebox.showerror("Lỗi", f"Không thể ghi file dữ liệu: {ex}", parent=dialog)
            return

        messagebox.showinfo("Thành công", "Đã lưu thông tin dự án thành công!", parent=dialog)
        dialog.destroy()

    # Nút Lưu thông tin
    btn_save = ttk.Button(main_frame, text="Lưu thông tin", command=save_project_info)
    btn_save.grid(row=len(field_names) + 1, column=0, columnspan=2, pady=25)

    parent.wait_window(dialog)  
    
    
    
def add_new_team(parent):
    # 1. Cấu hình đường dẫn
    folder_path = f"Data/{globalconfig.ketoan_nhancongthuengoai}"
    file_path = get_path(os.path.join(folder_path, "danh_sach_to_doi.JSON"))
    
    if not os.path.exists(get_path(folder_path)):
        os.makedirs(get_path(folder_path))

    # 2. Xử lý logic lưu dữ liệu
    def save_team():
        new_name = entry_team.get().strip()
        if not new_name: return

        # Đọc dữ liệu cũ (đọc thành dict thay vì list)
        data = {} 
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            except:
                data = {}

        # Kiểm tra trùng key
        if new_name in data:
            messagebox.showerror("Lỗi", f"Tổ đội '{new_name}' đã tồn tại!")
            return

        # Khởi tạo node con (workmen) cho team này
        data[new_name] = {"workmen": []}
        
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        
        dialog.destroy()
        #cập nhật tree_todoi        
        lam_moi_tree_todoi()
    # 3. Tạo Dialog Modal
    dialog = tk.Toplevel(parent) # Gán parent để làm modal
    dialog.title("Khai báo tên tổ đội")
    dialog.geometry("350x200")
    
    # Biến cửa sổ thành Modal
    dialog.transient(parent) 
    dialog.grab_set() 

    # Khung chứa nội dung (Padding cho thoáng)
    frame = tk.Frame(dialog, padx=20, pady=20)
    frame.pack(fill="both", expand=True)

    # Title
    tk.Label(frame, text="Khai báo tên tổ đội", font=("Arial", 14, "bold")).pack(pady=(0, 20))

    # Row Tên tổ đội
    row_frame = tk.Frame(frame)
    row_frame.pack(fill="x", pady=5)
    
    tk.Label(row_frame, text="Tên tổ đội", font=("Arial", 10)).pack(side="left", padx=(0, 10))
    entry_team = tk.Entry(row_frame, width=25, font=("Arial", 10))
    entry_team.pack(side="left", expand=True, fill="x")
    entry_team.focus()

    # Nút Tạo mới
    btn_create = tk.Button(frame, text="Tạo mới tổ đội", command=save_team, 
                           bg="#5b9bd5", fg="white", font=("Arial", 10, "bold"), 
                           height=2, padx=10)
    btn_create.pack(pady=20)

    # Chờ dialog đóng
    parent.wait_window(dialog)



def quan_ly_danh_muc_cong_viec(team_name):
    nctnpath = globalconfig.folder_nhancongthuengoai
    folder_path = os.path.join(nctnpath, team_name)
    file_path = os.path.join(folder_path, f"khai báo danh mục công việc_{team_name}.json")
    
    # Đảm bảo file tồn tại
    if not os.path.exists(folder_path): os.makedirs(folder_path)
    if not os.path.exists(file_path):
        with open(file_path, "w", encoding="utf-8") as f: json.dump([], f)

    dialog = tk.Toplevel()
    dialog.title(f"Khai báo danh mục công việc: {team_name}")
    
    # Bảng hiển thị
    columns = ("cv", "dv", "dg")
    tree = ttk.Treeview(dialog, columns=columns, show="headings")
    tree.heading("cv", text="Tên công việc")
    tree.heading("dv", text="Đơn vị tính")
    tree.heading("dg", text="Đơn giá")
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    def load_data():
        for i in tree.get_children(): tree.delete(i)
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
            for item in data:
                tree.insert("", "end", values=(item["cong_viec"], item["don_vi"], item["don_gia"]))
    def them_cv():
        # Gọi hàm tạo ở bước trước hoặc popup tại đây
        # 1. Tạo cửa sổ popup nhập liệu
        popup = tk.Toplevel(dialog)
        popup.title("Thêm công việc mới")
        popup.geometry("300x200")

        tk.Label(popup, text="Tên công việc:").pack(pady=5)
        entry_cv = tk.Entry(popup, width=30)
        entry_cv.pack()

        tk.Label(popup, text="Đơn vị tính:").pack(pady=5)
        entry_dv = tk.Entry(popup, width=30)
        entry_dv.pack()

        tk.Label(popup, text="Đơn giá:").pack(pady=5)
        entry_dg = tk.Entry(popup, width=30)
        entry_dg.pack()
        def luu_moi():
            ten = entry_cv.get()
            dv = entry_dv.get()
            dg = entry_dg.get()
            if not (ten and dv and dg):
               messagebox.showwarning("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
               return

            # 2. Đọc dữ liệu hiện có từ file
            with open(file_path, "r", encoding="utf-8") as f:
                 danh_sach = json.load(f)
            # --- KIỂM TRA TRÙNG LẶP ---
            # Chuyển tên nhập vào thành chữ thường để so sánh
            ten_normalized = ten.lower()
            # Kiểm tra nếu tên công việc đã tồn tại trong danh sách
            da_ton_tai = any(item["cong_viec"].strip().lower() == ten_normalized for item in danh_sach)            
            if da_ton_tai:
                messagebox.showerror("Lỗi", f"Công việc '{ten}' đã tồn tại trong danh sách!")
                return
            # ---------------------------
            # 3. Thêm mới
            danh_sach.append({
                "cong_viec": ten,
                "don_vi": dv,
                "don_gia": float(dg),
            })

            # 4. Ghi lại vào file
            with open(file_path, "w", encoding="utf-8") as f:
                 json.dump(danh_sach, f, ensure_ascii=False, indent=4)

            popup.destroy()
            load_data() # Làm mới lại bảng chính
            messagebox.showinfo("Thành công", "Đã thêm công việc!")

        tk.Button(popup, text="Lưu", command=luu_moi).pack(pady=10)        
    def xoa_cv():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Thông báo", "Vui lòng chọn một công việc để xóa!")
            return
            
        # Lấy thông tin dòng đang chọn
        item_values = tree.item(selected[0])['values']
        ten_cv_can_xoa = item_values[0]
        
        # 1. Hộp thoại cảnh báo xác nhận xóa
        if not messagebox.askyesno("Xác nhận", f"Bạn có chắc chắn muốn xóa công việc: '{ten_cv_can_xoa}' không?"):
            return
            
        # 2. Xóa trong file JSON
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                danh_sach = json.load(f)
            
            # Lọc lại danh sách: giữ lại những cái KHÔNG trùng với tên vừa chọn
            # Lưu ý: Nếu tên công việc không duy nhất, cách này sẽ xóa tất cả các dòng có cùng tên đó.
            # Nếu cần chính xác, bạn nên dùng chỉ số index của Treeview để lọc.
            danh_sach_moi = [item for item in danh_sach if item["cong_viec"] != ten_cv_can_xoa]
            
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(danh_sach_moi, f, ensure_ascii=False, indent=4)
                
            # 3. Làm mới danh sách và thông báo
            load_data()
            messagebox.showinfo("Thành công", f"Đã xóa: {ten_cv_can_xoa}")
            
        except Exception as e:
            messagebox.showerror("Lỗi", f"Có lỗi xảy ra khi xóa dữ liệu: {e}")
    def sua_cv():
        selected = tree.selection()
        if not selected:
            messagebox.showwarning("Thông báo", "Vui lòng chọn một công việc để sửa!")
            return
        
        # Lấy giá trị hiện tại của dòng đang chọn
        item_values = tree.item(selected[0])['values']
        ten_cu, dv_cu, dg_cu = item_values[0], item_values[1], item_values[2]
        
        # Tạo popup sửa
        popup = tk.Toplevel(dialog)
        popup.title("Sửa công việc")
        popup.geometry("300x250")

        tk.Label(popup, text="Tên công việc:").pack(pady=5)
        entry_cv = tk.Entry(popup, width=30)
        entry_cv.insert(0, ten_cu) # Điền sẵn dữ liệu cũ
        entry_cv.pack()

        tk.Label(popup, text="Đơn vị tính:").pack(pady=5)
        entry_dv = tk.Entry(popup, width=30)
        entry_dv.insert(0, dv_cu)
        entry_dv.pack()

        tk.Label(popup, text="Đơn giá:").pack(pady=5)
        entry_dg = tk.Entry(popup, width=30)
        entry_dg.insert(0, dg_cu)
        entry_dg.pack()

        def luu_sua():
            ten_moi = entry_cv.get().strip()
            dv_moi = entry_dv.get().strip()
            dg_moi = entry_dg.get().strip()

            if not (ten_moi and dv_moi and dg_moi):
                messagebox.showwarning("Lỗi", "Vui lòng nhập đầy đủ thông tin!")
                return

            # Đọc file
            with open(file_path, "r", encoding="utf-8") as f:
                danh_sach = json.load(f)
            
            # Kiểm tra trùng tên (nếu tên mới khác tên cũ)
            if ten_moi.lower() != ten_cu.lower():
                if any(item["cong_viec"].strip().lower() == ten_moi.lower() for item in danh_sach):
                    messagebox.showerror("Lỗi", "Tên công việc đã tồn tại!")
                    return

            # Cập nhật dữ liệu
            for item in danh_sach:
                if item["cong_viec"] == ten_cu:
                    item["cong_viec"] = ten_moi
                    item["don_vi"] = dv_moi
                    item["don_gia"] = dg_moi
                    break
            
            # Ghi lại file
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(danh_sach, f, ensure_ascii=False, indent=4)
            
            popup.destroy()
            load_data()
            messagebox.showinfo("Thành công", "Đã cập nhật công việc!")

        tk.Button(popup, text="Lưu thay đổi", command=luu_sua).pack(pady=10)
    load_data()
    # Nút bấm điều khiển
    btn_frame = tk.Frame(dialog)
    btn_frame.pack(fill="x", pady=5)
    tk.Button(btn_frame, text="Thêm", command=them_cv).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Sửa", command=sua_cv).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Xóa", command=xoa_cv).pack(side="left", padx=5)    