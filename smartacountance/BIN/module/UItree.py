UI tree
root
   
   ->smartAC---
   ->utils --- addcontract (dialog = tk.Toplevel())
                tk.Button(dialog, text="Lưu Hợp đồng", command=save_contract, bg="green", fg="white").pack(pady=10)
                tk.Button(dialog, text="Xem danh sách hợp đồng", 
                  command=lambda: hien_thi_danh_sach_hop_dong(team_name)).pack(pady=5)
                def on_closing():
                    hien_thi_danh_sach_hop_dong(team_name)  
                dialog.protocol("WM_DELETE_WINDOW", on_closing)  
   ->utils ---- hien_thi_danh_sach_hop_dong () (list_win = tk.Toplevel())
                tree_hd = ttk.Treeview(list_win, columns=("col1", "col2", "col3", "col4"), show="tree headings")
                button_frame = tk.Frame(list_win)
                button_frame.pack(pady=10)
                tk.Button(
                    button_frame, 
                    text="Trở lại danh sách tổ đội", 
                    width=25, 
                    height=2, 
                    bg='#ADD8E6',             # Màu xanh sáng (LightBlue)
                    activebackground='#87CEFA', # Màu khi nhấn vào (LightSkyBlue)
                    command=open_root
                ).pack(side=tk.LEFT, padx=10) # padx=10 giúp 2 nút cách nhau rõ ràng hơn

                tk.Button(
                    button_frame, 
                    text="Mở thư mục in", 
                    width=25, 
                    height=2,
                    bg='#ADD8E6',             # Màu xanh sáng (LightBlue)
                    activebackground='#87CEFA', # Màu khi nhấn vào (LightSkyBlue)        
                    command=open_output
                ).pack(side=tk.LEFT, padx=10)
    ->utils --- mo_dialog_hop_dong(team_name, data=data, file_path=file_path)    #mở cửa sổ sửa hợp đồng -> tạo hợp đồng, sửa hợp đồng
    ->utils --- load_names() # tải danh sách nhân công thuê ngoài
    
    #kiến trúc dữ liệu
    #Data/nhan_cong_thue_ngoai/
    #Data/nhan_cong_thue_ngoai/->tạo mới tổ đội -> danh_sach_to_doi.JSON -> menu sửa, xóa , thêm hợp đồng -> 
                                  -> thêm workman -> danh_sach_to_doi.JSON   -> menu sửa, xóa                       
                                  -> thêm hợp đồng -> 
    #Data/ke_toan_phai_thu/
    #luồng UI-Dữ liệu cũ
    sidebar: btn_cha_1 
                      -> "Tải danh mục nhân công", lambda: load_names() ->create new tree: treeview fill column and data from excel file 
                                                                        ->tree.bind("<Button-3>", lambda event: show_custom_menu(event, globalconfig.root))
                                                                        ->tree.bind("<<TreeviewSelect>>", cap_nhat_team_name_selected)
                                                                        -> from tree -> update globalconfig.tree_todoi
    ##########################################################################################################################################################
                     
    kiến trúc UI-dữ liệu, định nghĩa luồng UI-dữ liệu
    root.title("Smart AC Management")
    tải danh sách tổ đội sửa lại hàm load_names()
    khởi tạo root: gọi hàm load_names() để khởi tạo tree_todoi
    
    menu: menubar-> file_menu
                      -> command= lambda: back_up_system()
                      -> command= lambda: restore_system()
                      -> command= lambda: open_restore_system()
    sidebar: btn_cha_1
                      -> Thêm thông tin dự án -> call function nhap_thong_tin_du_an(parent) 
                      -> "Tải danh mục nhân công", lambda: load_names()
                                                   tree.bind("<Button-3>", lambda event: show_custom_menu(event, globalconfig.root))                                                         
                                                       -> lambda: mo_dialog_hop_dong(team_name)) : khi tạo hợp đồng cho phép khối lượng = 0 nhưng tổng giá trị hợp đồng > 0 
                                                       
                                                   tree.bind("<<TreeviewSelect>>", cap_nhat_team_name_selected)
                                                   
                                                   
                                                   
                      -> "Tạo mới tổ đội", lambda: add_new_team(globalconfig.root) -> hiển thị lại danh sách tổ độidựa trên globalconfig.tree_todoi
                                                   đã lưu tổ đội vào file 
                                                   
                                                   
                      -> "Khai báo danh mục công việc", lambda: open_danh_muc_congviec()
                      -> "Soạn thảo danh sách nhân công", lambda: open_excel_file()
    Hiển thị danh sách hợp đồng
                hien_thi_danh_sach_hop_dong() 
                            tree_hd.bind("<Button-3>", lambda event: show_custom_menu_danh_sach_hop_dong(event, globalconfig.root))
                                                                     
                            tree_hd.bind("<Button-3>", handle_right_click)
                                                                     
                                                                     
                            
                            button_1 = ctk.CTkButton (command=utilis.open_root)
                            button_2 = ctk.CTkButton(command=utilis.open_output)