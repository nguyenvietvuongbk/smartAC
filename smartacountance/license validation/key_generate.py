import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import base64
from cryptography.hazmat.primitives.asymmetric import rsa, padding
from cryptography.hazmat.primitives import serialization, hashes

# --- Các hàm xử lý logic ---
def generate_keys():
    directory = os.path.dirname(os.path.abspath(__file__))
    try:
        private_key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
        public_key = private_key.public_key()

        with open(os.path.join(directory, "private_key.pem"), "wb") as f:
            f.write(private_key.private_bytes(
                encoding=serialization.Encoding.PEM, format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            ))

        with open(os.path.join(directory, "public_key.pem"), "wb") as f:
            f.write(public_key.public_bytes(
                encoding=serialization.Encoding.PEM, format=serialization.PublicFormat.SubjectPublicKeyInfo
            ))
        messagebox.showinfo("Thành công", "Đã tạo cặp khóa tại thư mục hiện tại!")
    except Exception as e:
        messagebox.showerror("Lỗi", str(e))

def create_license_file():
    hwid = entry_hwid.get()
    expiry = entry_expiry.get()
    
    if not hwid or not expiry:
        messagebox.showwarning("Thiếu thông tin", "Vui lòng nhập đầy đủ HWID và ngày hết hạn.")
        return

    if not os.path.exists("private_key.pem"):
        messagebox.showerror("Lỗi", "Không tìm thấy file 'private_key.pem' trong thư mục này!")
        return

    try:
        # 1. Đọc Private Key
        with open("private_key.pem", "rb") as key_file:
            private_key = serialization.load_pem_private_key(key_file.read(), password=None)

        # 2. Tạo nội dung ký
        message = f"{hwid}|{expiry}".encode()
        
        # 3. Ký số
        signature = private_key.sign(
            message,
            padding.PSS(mgf=padding.MGF1(hashes.SHA256()), salt_length=padding.PSS.MAX_LENGTH),
            hashes.SHA256()
        )
        
        # 4. Gói dữ liệu: DATA|SIGN|BASE64_SIGNATURE
        encoded_sig = base64.b64encode(signature).decode('utf-8')
        final_payload = f"{hwid}|{expiry}|SIGN|{encoded_sig}"

        # 5. Lưu file
        file_path = filedialog.asksaveasfilename(defaultextension=".lic", filetypes=[("License files", "*.lic")])
        if file_path:
            with open(file_path, "w") as f:
                f.write(final_payload)
            messagebox.showinfo("Thành công", f"Đã tạo file license tại:\n{file_path}")

    except Exception as e:
        messagebox.showerror("Lỗi kỹ thuật", str(e))

# --- Thiết lập giao diện ---
root = tk.Tk()
root.title("System License Manager")
root.geometry("400x300")

notebook = ttk.Notebook(root)
notebook.pack(pady=10, expand=True)

# Tab 1: Tạo khóa (Chỉ dùng 1 lần)
tab1 = ttk.Frame(notebook)
notebook.add(tab1, text="Tạo Cặp Khóa")
tk.Button(tab1, text="Tạo Private/Public Key", command=generate_keys, height=3).pack(pady=50)

# Tab 2: Tạo License (Dùng cấp cho khách)
tab2 = ttk.Frame(notebook)
notebook.add(tab2, text="Tạo License .lic")

tk.Label(tab2, text="Hardware ID (Product ID):").pack(pady=5)
entry_hwid = tk.Entry(tab2, width=40)
entry_hwid.pack()

tk.Label(tab2, text="Ngày hết hạn (dd/mm/yyyy):").pack(pady=5)
entry_expiry = tk.Entry(tab2, width=40)
entry_expiry.insert(0, "31/12/2026")
entry_expiry.pack()

tk.Button(tab2, text="Tạo file .lic", command=create_license_file, bg="#2196F3", fg="white", height=2).pack(pady=20)

root.mainloop()