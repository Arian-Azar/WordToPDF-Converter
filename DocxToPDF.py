"""
Professional Word to PDF Batch Converter with PyQt5
Converts multiple Word documents to PDF simultaneously
"""

import signal
signal.signal(signal.SIGINT, signal.SIG_DFL)

import sys
import os
import threading
from pathlib import Path
from typing import List, Optional
from datetime import datetime

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

# Try to import docx2pdf with fallback
try:
    from docx2pdf import convert
    CONVERT_AVAILABLE = True
except ImportError:
    CONVERT_AVAILABLE = False
    print("⚠️ Warning: docx2pdf not installed. Run: pip install docx2pdf")

class ConversionWorker(QThread):
    """
    Worker thread for handling PDF conversion without blocking the UI
    """
    progress_updated = pyqtSignal(int, int)  # current, total
    file_status = pyqtSignal(str, str)  # filename, status
    conversion_finished = pyqtSignal(int, int)  # success_count, error_count
    
    def __init__(self, files: List[str], output_dir: str):
        super().__init__()
        self.files = files
        self.output_dir = output_dir
        self.is_cancelled = False
        
    def cancel(self):
        """Cancel the ongoing conversion"""
        self.is_cancelled = True
    
    def run(self):
        """Main conversion logic running in background thread"""
        total_files = len(self.files)
        success_count = 0
        error_count = 0
        
        for i, word_file in enumerate(self.files, 1):
            if self.is_cancelled:
                self.file_status.emit("Conversion cancelled by user", "warning")
                break
            
            file_name = os.path.basename(word_file)
            self.progress_updated.emit(i, total_files)
            
            try:
                # Perform the conversion
                convert(word_file, self.output_dir)
                success_count += 1
                self.file_status.emit(f"✓ {file_name} - Converted successfully", "success")
                
            except Exception as e:
                error_count += 1
                error_msg = str(e)[:100]
                self.file_status.emit(f"✗ {file_name} - Error: {error_msg}", "error")
            
            # Small delay to prevent UI freezing
            self.msleep(10)
        
        self.conversion_finished.emit(success_count, error_count)


class WordToPDFConverter(QMainWindow):
    """
    Main application window for Word to PDF batch converter
    """
    
    def __init__(self):
        super().__init__()
        self.files_list: List[str] = []
        self.conversion_worker: Optional[ConversionWorker] = None
        self.init_ui()
        self.check_dependencies()
        
    def check_dependencies(self):
        """Check if required dependencies are installed"""
        if not CONVERT_AVAILABLE:
            QMessageBox.warning(
                self, 
                "Dependency Missing",
                "The 'docx2pdf' library is not installed!\n\n"
                "Please install it using:\n"
                "pip install docx2pdf\n\n"
                "Then restart the application."
            )
    
    def init_ui(self):
        """Initialize the user interface"""
        self.setWindowTitle('Professional Word to PDF Batch Converter')
        self.setGeometry(100, 100, 800, 700)
        self.setMinimumSize(600, 500)
        
        # Set application icon (optional - replace with your icon file)
        # self.setWindowIcon(QIcon('icon.png'))
        
        # Modern stylesheet
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
            }
            QPushButton {
                background-color: #0078d4;
                color: white;
                border: none;
                padding: 10px 20px;
                font-size: 13px;
                font-weight: 500;
                border-radius: 6px;
                font-family: 'Segoe UI', Arial;
            }
            QPushButton:hover {
                background-color: #106ebe;
            }
            QPushButton:pressed {
                background-color: #005a9e;
            }
            QPushButton#convertBtn {
                background-color: #107c10;
                font-size: 14px;
                font-weight: bold;
                padding: 12px 24px;
            }
            QPushButton#convertBtn:hover {
                background-color: #0b5e0b;
            }
            QPushButton#removeBtn {
                background-color: #d13438;
            }
            QPushButton#removeBtn:hover {
                background-color: #b12e32;
            }
            QPushButton#cancelBtn {
                background-color: #6c757d;
            }
            QPushButton#cancelBtn:hover {
                background-color: #5a6268;
            }
            QListWidget {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                padding: 8px;
                background-color: white;
                font-family: 'Segoe UI', Arial;
                font-size: 12px;
                outline: none;
            }
            QListWidget::item {
                padding: 8px;
                border-bottom: 1px solid #f0f0f0;
            }
            QListWidget::item:selected {
                background-color: #e3f2fd;
                color: #0078d4;
            }
            QProgressBar {
                border: 2px solid #dee2e6;
                border-radius: 8px;
                text-align: center;
                font-weight: bold;
                height: 25px;
                background-color: white;
            }
            QProgressBar::chunk {
                background-color: #107c10;
                border-radius: 6px;
            }
            QLineEdit {
                border: 2px solid #dee2e6;
                border-radius: 6px;
                padding: 8px;
                font-family: 'Segoe UI', Arial;
                font-size: 12px;
                background-color: white;
            }
            QLineEdit:focus {
                border-color: #0078d4;
            }
            QLabel {
                font-family: 'Segoe UI', Arial;
                color: #212529;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
        """)
        
        # Central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(25, 25, 25, 25)
        
        # Header section
        self.create_header_section(main_layout)
        
        # File selection section
        self.create_file_selection_section(main_layout)
        
        # Output configuration section
        self.create_output_section(main_layout)
        
        # Progress and status section
        self.create_progress_section(main_layout)
        
        # Statistics footer
        self.update_statistics()
    
    def create_header_section(self, parent_layout: QVBoxLayout):
        """Create the header/title section"""
        header_widget = QWidget()
        header_layout = QHBoxLayout(header_widget)
        
        title_label = QLabel("📄 Word to PDF Batch Converter")
        title_font = QFont("Segoe UI", 20, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #0078d4;")
        header_layout.addWidget(title_label)
        
        # Version badge
        version_label = QLabel("v2.0")
        version_label.setStyleSheet("""
            background-color: #0078d4;
            color: white;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
        """)
        header_layout.addWidget(version_label)
        header_layout.addStretch()
        
        parent_layout.addWidget(header_widget)
        
        # Description
        desc_label = QLabel("Convert multiple Microsoft Word documents to PDF format simultaneously")
        desc_label.setStyleSheet("color: #6c757d; font-size: 12px; margin-bottom: 10px;")
        parent_layout.addWidget(desc_label)
    
    def create_file_selection_section(self, parent_layout: QVBoxLayout):
        """Create file selection and management section"""
        # Selection buttons group
        btn_group = QGroupBox("File Management")
        btn_layout = QVBoxLayout(btn_group)
        
        # Top buttons
        top_btn_layout = QHBoxLayout()
        self.select_btn = QPushButton("📁 Add Word Files")
        self.select_btn.clicked.connect(self.select_files)
        self.select_btn.setMinimumHeight(40)
        top_btn_layout.addWidget(self.select_btn)
        
        top_btn_layout.addStretch()
        
        # Info label
        self.file_count_label = QLabel("No files selected")
        self.file_count_label.setStyleSheet("color: #6c757d; font-weight: normal;")
        top_btn_layout.addWidget(self.file_count_label)
        
        btn_layout.addLayout(top_btn_layout)
        
        # Files list
        self.files_listbox = QListWidget()
        self.files_listbox.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.files_listbox.setAlternatingRowColors(True)
        self.files_listbox.setMinimumHeight(200)
        btn_layout.addWidget(self.files_listbox)
        
        # Management buttons
        management_layout = QHBoxLayout()
        self.remove_btn = QPushButton("🗑 Remove Selected")
        self.remove_btn.setObjectName("removeBtn")
        self.remove_btn.clicked.connect(self.remove_selected)
        management_layout.addWidget(self.remove_btn)
        
        self.clear_btn = QPushButton("🗑 Clear All")
        self.clear_btn.clicked.connect(self.clear_all)
        management_layout.addWidget(self.clear_btn)
        
        management_layout.addStretch()
        btn_layout.addLayout(management_layout)
        
        parent_layout.addWidget(btn_group)
    
    def create_output_section(self, parent_layout: QVBoxLayout):
        """Create output configuration section"""
        output_group = QGroupBox("Output Configuration")
        output_layout = QVBoxLayout(output_group)
        
        # Output path selection
        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel("Output Directory:"))
        
        self.output_path = QLineEdit()
        default_output = os.path.expanduser(f"~/Desktop/PDF_Output_{datetime.now().strftime('%Y%m%d')}")
        os.makedirs(default_output, exist_ok=True)
        self.output_path.setText(default_output)
        path_layout.addWidget(self.output_path)
        
        self.output_btn = QPushButton("📂 Browse")
        self.output_btn.clicked.connect(self.select_output_folder)
        path_layout.addWidget(self.output_btn)
        
        output_layout.addLayout(path_layout)
        
        # Open output folder checkbox
        self.open_folder_checkbox = QCheckBox("Open output folder after conversion")
        self.open_folder_checkbox.setChecked(True)
        output_layout.addWidget(self.open_folder_checkbox)
        
        parent_layout.addWidget(output_group)
    
    def create_progress_section(self, parent_layout: QVBoxLayout):
        """Create progress monitoring section"""
        progress_group = QGroupBox("Conversion Progress")
        progress_layout = QVBoxLayout(progress_group)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        progress_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("✓ Ready to convert documents")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("""
            color: #28a745;
            padding: 10px;
            font-weight: 500;
            background-color: #f8f9fa;
            border-radius: 6px;
        """)
        progress_layout.addWidget(self.status_label)
        
        # Convert button
        self.convert_btn = QPushButton("🚀 Start Conversion")
        self.convert_btn.setObjectName("convertBtn")
        self.convert_btn.clicked.connect(self.start_conversion)
        self.convert_btn.setMinimumHeight(45)
        progress_layout.addWidget(self.convert_btn)
        
        parent_layout.addWidget(progress_group)
        
        # Information footer
        info_label = QLabel(
            "ℹ️ Note: Microsoft Word must be installed on your system for conversion to work properly"
        )
        info_label.setWordWrap(True)
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet("color: #6c757d; font-size: 10px; padding: 10px;")
        parent_layout.addWidget(info_label)
    
    def update_statistics(self):
        """Update file statistics display"""
        count = len(self.files_list)
        if count == 0:
            self.file_count_label.setText("No files selected")
        elif count == 1:
            self.file_count_label.setText("1 file selected")
        else:
            self.file_count_label.setText(f"{count} files selected")
    
    def select_files(self):
        """Open file dialog to select Word documents"""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select Word Documents to Convert",
            "",
            "Word Documents (*.docx *.doc);;All Files (*.*)"
        )
        
        added_count = 0
        for file in files:
            if file not in self.files_list:
                self.files_list.append(file)
                self.files_listbox.addItem(f"📄 {os.path.basename(file)}")
                added_count += 1
        
        if added_count > 0:
            self.update_status(f"Added {added_count} new file(s)", "success")
            self.update_statistics()
        elif files:
            self.update_status("Selected files already in list", "info")
    
    def remove_selected(self):
        """Remove selected files from the list"""
        selected_items = self.files_listbox.selectedItems()
        if not selected_items:
            QMessageBox.information(self, "No Selection", "Please select files to remove")
            return
        
        reply = QMessageBox.question(
            self, 
            "Confirm Removal",
            f"Remove {len(selected_items)} selected file(s)?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            for item in selected_items:
                row = self.files_listbox.row(item)
                del self.files_list[row]
                self.files_listbox.takeItem(row)
            
            self.update_status(f"Removed {len(selected_items)} file(s)", "success")
            self.update_statistics()
    
    def clear_all(self):
        """Clear all files from the list"""
        if not self.files_list:
            return
            
        reply = QMessageBox.question(
            self, 
            "Confirm Clear All",
            "Are you sure you want to remove all files from the list?",
            QMessageBox.Yes | QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.files_list.clear()
            self.files_listbox.clear()
            self.update_status("Cleared all files from list", "success")
            self.update_statistics()
    
    def select_output_folder(self):
        """Open dialog to select output directory"""
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Directory for PDF Files",
            self.output_path.text()
        )
        
        if folder:
            self.output_path.setText(folder)
            self.update_status(f"Output directory changed to: {folder}", "success")
    
    def update_status(self, message: str, status_type: str = "info"):
        """Update status label with different color schemes"""
        colors = {
            "success": "#28a745",
            "error": "#dc3545",
            "info": "#17a2b8",
            "warning": "#ffc107"
        }
        
        icons = {
            "success": "✓",
            "error": "✗",
            "info": "ℹ️",
            "warning": "⚠️"
        }
        
        color = colors.get(status_type, "#6c757d")
        icon = icons.get(status_type, "")
        
        self.status_label.setText(f"{icon} {message}")
        self.status_label.setStyleSheet(f"""
            color: {color};
            padding: 10px;
            font-weight: 500;
            background-color: #f8f9fa;
            border-radius: 6px;
        """)
        QApplication.processEvents()
    
    def on_file_status(self, filename: str, status: str):
        """Handle file conversion status updates"""
        if "success" in status.lower():
            self.update_status(filename, "success")
        elif "error" in status.lower():
            self.update_status(filename, "error")
        else:
            self.update_status(filename, "info")
    
    def on_conversion_finished(self, success_count: int, error_count: int):
        """Handle conversion completion"""
        total = success_count + error_count
        
        # Enable buttons
        self.convert_btn.setEnabled(True)
        self.convert_btn.setText("🚀 Start Conversion")
        self.select_btn.setEnabled(True)
        self.output_btn.setEnabled(True)
        self.remove_btn.setEnabled(True)
        self.clear_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        
        # Show result message
        result_msg = f"Conversion Complete!\n\n✅ Successfully converted: {success_count}/{total} files"
        if error_count > 0:
            result_msg += f"\n❌ Failed: {error_count} files"
        
        QMessageBox.information(self, "Conversion Results", result_msg)
        
        # Clear successful files if all converted
        if success_count == total and total > 0:
            self.files_list.clear()
            self.files_listbox.clear()
            self.update_statistics()
            self.update_status("All files converted successfully", "success")
        else:
            self.update_status(f"Conversion finished: {success_count} succeeded, {error_count} failed", 
                             "info" if error_count == 0 else "warning")
        
        # Open output folder if requested
        if self.open_folder_checkbox.isChecked() and success_count > 0:
            output_dir = self.output_path.text()
            if os.path.exists(output_dir):
                os.startfile(output_dir) if sys.platform == "win32" else None
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        # Validation checks
        if not self.files_list:
            QMessageBox.warning(self, "No Files", "Please select at least one Word document to convert")
            return
        
        if not CONVERT_AVAILABLE:
            QMessageBox.critical(
                self, 
                "Missing Dependency",
                "docx2pdf library is not installed.\n\n"
                "Please run: pip install docx2pdf\n"
                "Then restart the application."
            )
            return
        
        # Verify output directory
        output_dir = self.output_path.text()
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Cannot create output directory:\n{str(e)}")
            return
        
        # Disable buttons during conversion
        self.convert_btn.setEnabled(False)
        self.convert_btn.setText("⏳ Converting...")
        self.select_btn.setEnabled(False)
        self.output_btn.setEnabled(False)
        self.remove_btn.setEnabled(False)
        self.clear_btn.setEnabled(False)
        
        # Setup progress bar
        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(len(self.files_list))
        self.progress_bar.setValue(0)
        
        # Create and start worker thread
        self.conversion_worker = ConversionWorker(self.files_list, output_dir)
        self.conversion_worker.progress_updated.connect(self.progress_bar.setValue)
        self.conversion_worker.file_status.connect(self.on_file_status)
        self.conversion_worker.conversion_finished.connect(self.on_conversion_finished)
        self.conversion_worker.start()
        
        self.update_status("Conversion in progress...", "info")


def main():
    """Main application entry point"""
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Modern cross-platform style
    
    # Set application metadata
    app.setApplicationName("Word to PDF Converter")
    app.setApplicationVersion("2.0")
    app.setOrganizationName("PDFConverter")
    
    # Create and show main window
    window = WordToPDFConverter()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()