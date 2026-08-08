import sys
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout, QTreeWidget, 
                             QTreeWidgetItem, QTableView, QSplitter, QToolBar, 
                             QDockWidget, QLabel)
#from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QHBoxLayout, 
#                             QTreeWidget, QTreeWidgetItem, QTableView, QSplitter, 
#                             QToolBar, QDockWidget, QLabel)                             
from PyQt6.QtCore import Qt

from PyQt6.QtWidgets import QTreeWidget, QTreeWidgetItem
from PyQt6.QtGui import QFont
from PyQt6.QtCore import Qt, QAbstractTableModel

class NhanSuModel(QAbstractTableModel):
    def __init__(self, data=None):
        super().__init__()
        # Data là danh sách các dòng (mỗi dòng là một list hoặc dict)
        self._data = data or []
        self._headers = ["STT", "Họ và tên", "Ngày sinh", "Giới tính", "MST", 
                         "CMT/CCCD", "Ngày cấp", "Nơi cấp", "Số TK", 
                         "Tên ngân hàng", "Tên chủ TK", "Địa chỉ", "Nơi làm việc", 
                         "Công việc", "Tổ đội"]

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._headers)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole:
            return self._data[index.row()][index.column()]
        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role == Qt.ItemDataRole.DisplayRole and orientation == Qt.Orientation.Horizontal:
            return self._headers[section]
        return None

    # --- PHẦN QUAN TRỌNG ĐỂ SOẠN THẢO (EDITABLE) ---
    def flags(self, index):
        # Cho phép các ô có thể chỉnh sửa được
        return Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsEditable

    def setData(self, index, value, role=Qt.ItemDataRole.EditRole):
        if role == Qt.ItemDataRole.EditRole:
            self._data[index.row()][index.column()] = value
            self.dataChanged.emit(index, index) # Thông báo cho bảng cập nhật lại giao diện
            return True
        return False
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Phần mềm Nghiệp vụ - Mô phỏng")
        self.resize(1200, 800)

        # 1. Main Layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # --- SỬ DỤNG QSPLITTER ĐỂ CHO PHÉP CO GIÃN ---
        splitter = QSplitter(Qt.Orientation.Horizontal)
        main_layout.addWidget(splitter)

        # 2. Sidebar (Sidebar được thêm vào Splitter)
        self.sidebar = QTreeWidget()
        self.sidebar.setHeaderHidden(True)
        splitter.addWidget(self.sidebar) 

        # Tạo các mục menu
        item_danhmuc = QTreeWidgetItem(self.sidebar, ["Danh mục"])
        child1 = QTreeWidgetItem(item_danhmuc, ["Khách hàng"])
        child2 = QTreeWidgetItem(item_danhmuc, ["Nhà cung cấp"])
        child3 = QTreeWidgetItem(item_danhmuc, ["Vật tư hàng hóa"])

        item_nghiepvu = QTreeWidgetItem(self.sidebar, ["Nghiệp vụ"])
        child4 = QTreeWidgetItem(item_nghiepvu, ["Mua hàng"])
        child5 = QTreeWidgetItem(item_nghiepvu, ["Bán hàng"])

        self.sidebar.expandAll()



        # 3. Main View (Bảng dữ liệu)
        self.table = QTableView()
        splitter.addWidget(self.table) # <--- QUAN TRỌNG: Phải thêm Table vào splitter
        
        # Thiết lập tỉ lệ co giãn ban đầu (Sidebar rộng 200, Table rộng 1000)
        splitter.setSizes([200, 1000])
        
        #main_layout.addWidget(self.table)
        self.table.setStyleSheet("""
            QTableView {
                background-color: #E0F7FA;
                alternate-background-color: #C8E6C9; /* Màu dòng xen kẽ (nếu bạn bật) */
                gridline-color: #81D4FA;             /* Màu đường kẻ lưới */
                selection-background-color: #0288D1; /* Màu khi chọn dòng */
            }
            QHeaderView::section {
                background-color: #0288D1;           /* Màu nền tiêu đề */
                color: white;                        /* Màu chữ tiêu đề */
                font-weight: bold;
                padding: 4px;
                border: 1px solid #01579B;
            }
        """)
        self.table.setAlternatingRowColors(True)
        # --- ĐÂY LÀ ĐOẠN BẠN THÊM VÀO ---
        # 1. Chuẩn bị dữ liệu mẫu
        data_goc = [
            ["1", "Dương Công Văn", "20/07/1986", "Nam", "01908601438", "01908601438", "05/08/2022", "Cục CS QLHC", "", "", "", "Xóm soi 1", "Công trình", "Nhân công", "Tổ mài nền"],
            ["2", "Nguyễn Hồng Thái", "05/07/1966", "Nam", "019066011776", "019066011776", "29/10/2025", "Bộ Công An", "", "", "", "Xóm 2", "Công trình", "Nhân công", "Tổ mài nền"],
        ]
        
        # 2. Khởi tạo Model và gán vào Table
        self.model = NhanSuModel(data_goc)
        self.table.setModel(self.model)
        
        # 3. Tối ưu hiển thị cột
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.verticalHeader().setDefaultSectionSize(30)
        # ---------------------------------
        # 4. Toolbar (Menu trên cùng)
        toolbar = QToolBar("Main Toolbar")
        self.addToolBar(toolbar)
        toolbar.addAction("Thêm")
        toolbar.addAction("Xem")
        toolbar.addAction("Nạp")

        # 5. Dock Widget (Panel Nhắc nhở bên phải)
        self.dock = QDockWidget("Nhắc nhở", self)
        self.dock.setWidget(QLabel("Nội dung nhắc nhở tại đây..."))
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.dock)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())