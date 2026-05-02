"""
برنامه تبدیل Word به PDF با PyQt5 (بدون نیاز به Tkinter)
"""

import sys
import os
import threading
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

# بررسی docx2pdf
try:
    from docx2pdf import convert
    CONVERT_AVAILABLE = True
except ImportError:
    CONVERT_AVAILABLE = False
    print("⚠ Warning: docx2pdf not installed. Install with: pip install docx2pdf")

class WordToPDFConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.files_list = []
        self.initUI()
        
        if not CONVERT_AVAILABLE:
            QMessageBox.warning(self, "هشدار", 
                "کتابخانه docx2pdf نصب نیست!\n\n"
                "برای نصب، دستور زیر را در ترمینال اجرا کنید:\n"
                "pip install docx2pdf")
    
    def initUI(self):
        self.setWindowTitle('تبدیل Word به PDF')
        self.setGeometry(100, 100, 700, 600)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px;
                font-size: 12px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton#convertBtn {
                background-color: #FF9800;
                font-size: 14px;
                font-weight: bold;
                padding: 12px;
            }
            QPushButton#convertBtn:hover {
                background-color: #e68900;
            }
            QPushButton#removeBtn {
                background-color: #f44336;
            }
            QPushButton#removeBtn:hover {
                background-color: #da190b;
            }
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 5px;
                padding: 5px;
            }
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 5px;
            }
        """)
        
        # ویجت مرکزی
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # عنوان
        title = QLabel("📄 تبدیل همزمان Word به PDF")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        layout.addWidget(title)
        
        # دکمه انتخاب فایل
        self.select_btn = QPushButton("📁 افزودن فایل‌های Word")
        self.select_btn.clicked.connect(self.select_files)
        layout.addWidget(self.select_btn)
        
        # لیبل لیست فایل‌ها
        list_label = QLabel("📋 لیست فایل‌های انتخاب شده:")
        list_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(list_label)
        
        # لیست فایل‌ها
        self.files_listbox = QListWidget()
        self.files_listbox.setSelectionMode(QAbstractItemView.ExtendedSelection)
        layout.addWidget(self.files_listbox)
        
        # فریم دکمه‌های مدیریت
        btn_frame = QFrame()
        btn_layout = QHBoxLayout(btn_frame)
        
        self.remove_btn = QPushButton("🗑 حذف انتخاب شده")
        self.remove_btn.setObjectName("removeBtn")
        self.remove_btn.clicked.connect(self.remove_selected)
        btn_layout.addWidget(self.remove_btn)
        
        self.clear_btn = QPushButton("🗑 پاک کردن همه")
        self.clear_btn.clicked.connect(self.clear_all)
        btn_layout.addWidget(self.clear_btn)
        
        layout.addWidget(btn_frame)
        
        # انتخاب پوشه خروجی
        output_label = QLabel("📂 پوشه خروجی:")
        output_label.setFont(QFont("Arial", 10, QFont.Bold))
        layout.addWidget(output_label)
        
        output_frame = QFrame()
        output_layout = QHBoxLayout(output_frame)
        output_layout.setContentsMargins(0, 0, 0, 0)
        
        self.output_path = QLineEdit()
        default_output = os.path.expanduser("~/Desktop/PDF_Output")
        os.makedirs(default_output, exist_ok=True)
        self.output_path.setText(default_output)
        self.output_path.setReadOnly(True)
        output_layout.addWidget(self.output_path)
        
        self.output_btn = QPushButton("📂 تغییر")
        self.output_btn.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_btn)
        
        layout.addWidget(output_frame)
        
        # دکمه تبدیل
        self.convert_btn = QPushButton("🔄 شروع تبدیل به PDF")
        self.convert_btn.setObjectName("convertBtn")
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)
        
        # نوار پیشرفت
        self.progress = QProgressBar()
        layout.addWidget(self.progress)
        
        # وضعیت
        self.status_label = QLabel("✅ آماده برای شروع تبدیل")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #666; padding: 5px;")
        layout.addWidget(self.status_label)
        
        # اطلاعات
        info_label = QLabel("نکته: برای تبدیل فایل‌ها، نیاز به Microsoft Word نصب شده دارید")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet("color: #999; font-size: 10px; padding: 10px;")
        layout.addWidget(info_label)
    
    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "انتخاب فایل‌های Word",
            "",
            "Word Files (*.docx *.doc);;All Files (*.*)"
        )
        
        for file in files:
            if file not in self.files_list:
                self.files_list.append(file)
                self.files_listbox.addItem(f"📄 {os.path.basename(file)}")
        
        self.update_status(f"✅ {len(files)} فایل جدید اضافه شد (مجموع: {len(self.files_list)})")
    
    def remove_selected(self):
        selected_items = self.files_listbox.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "توجه", "لطفاً فایل‌هایی را که می‌خواهید حذف کنید انتخاب نمایید")
            return
        
        for item in selected_items:
            row = self.files_listbox.row(item)
            del self.files_list[row]
            self.files_listbox.takeItem(row)
        
        self.update_status(f"🗑 {len(selected_items)} فایل حذف شد (باقیمانده: {len(self.files_list)})")
    
    def clear_all(self):
        if self.files_list:
            reply = QMessageBox.question(self, "تأیید", 
                                         "آیا از پاک کردن تمام فایل‌های لیست اطمینان دارید؟",
                                         QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.files_list.clear()
                self.files_listbox.clear()
                self.update_status("🗑 لیست فایل‌ها پاک شد")
    
    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "انتخاب پوشه خروجی PDF",
            self.output_path.text()
        )
        if folder:
            self.output_path.setText(folder)
            self.update_status(f"📂 پوشه خروجی تغییر یافت: {folder}")
    
    def update_status(self, message, color="green"):
        self.status_label.setText(message)
        if color == "green":
            self.status_label.setStyleSheet("color: green; padding: 5px;")
        elif color == "red":
            self.status_label.setStyleSheet("color: red; padding: 5px;")
        elif color == "blue":
            self.status_label.setStyleSheet("color: blue; padding: 5px;")
        QApplication.processEvents()
    
    def convert_files(self):
        if not self.files_list:
            QMessageBox.warning(self, "هشدار", "لطفاً حداقل یک فایل Word انتخاب کنید!")
            return
        
        if not CONVERT_AVAILABLE:
            QMessageBox.critical(self, "خطا", "کتابخانه docx2pdf نصب نیست!")
            return
        
        output_dir = self.output_path.text()
        os.makedirs(output_dir, exist_ok=True)
        
        total_files = len(self.files_list)
        success_count = 0
        error_count = 0
        
        self.progress.setMaximum(total_files)
        self.progress.setValue(0)
        
        for i, word_file in enumerate(self.files_list, 1):
            file_name = os.path.basename(word_file)
            self.update_status(f"🔄 در حال تبدیل {file_name}... ({i}/{total_files})", "blue")
            
            try:
                convert(word_file, output_dir)
                success_count += 1
                self.update_status(f"✅ {file_name} با موفقیت تبدیل شد", "green")
                
            except Exception as e:
                error_count += 1
                self.update_status(f"❌ خطا در تبدیل {file_name}: {str(e)[:50]}", "red")
            
            self.progress.setValue(i)
            QApplication.processEvents()
        
        # نمایش نتیجه
        result_msg = f"✅ موفق: {success_count} از {total_files} فایل"
        if error_count > 0:
            result_msg += f"\n❌ ناموفق: {error_count} فایل"
        
        QMessageBox.information(self, "پایان تبدیل", result_msg)
        
        if success_count == total_files:
            self.files_list.clear()
            self.files_listbox.clear()
        
        self.progress.setValue(0)
        self.update_status("✅ آماده برای تبدیل مجدد", "green")
    
    def start_conversion(self):
        # اجرا در ترد جداگانه
        thread = threading.Thread(target=self.convert_files)
        thread.daemon = True
        thread.start()

def main():
    app = QApplication(sys.argv)
    window = WordToPDFConverter()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()