import sys
import pandas as pd
import numpy as np
import random
import traceback
import os
import re
from datetime import datetime

from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QInputDialog, QLineEdit, QProgressBar, QDateEdit, QLabel
from PyQt6.QtCore import Qt, QObject, QThread, pyqtSignal as Signal, QRect, QDate
from PyQt6.QtGui import QPainter, QColor, QPen, QIcon

try:
    from searchui import Ui_MainWindow
except ImportError:
    app = QApplication(sys.argv)
    QMessageBox.critical(None, "خطای حیاتی", "فایل searchui.py یافت نشد. لطفاً ابتدا فایل .ui را با دستور pyuic6 به .py تبدیل کنید.")
    sys.exit()

# =============================================================================
# class for hard works (filters and process)
# =============================================================================
class ExcelWorker(QObject):
    finished = Signal(str)
    error = Signal(str)
    progress_update = Signal(int)
    status_update = Signal(str)

    def __init__(self, files_to_read_info, filters, output_path):
        super().__init__()
        self.files_to_read_info = files_to_read_info
        self.filters = filters
        self.output_path = output_path

    def run(self):
        try:
            self.status_update.emit("شروع پردازش...")
            self.progress_update.emit(5)
            
            filtered_dfs = []
            total_files = len(self.files_to_read_info)
            file_count = 0

            for key, file_path in self.files_to_read_info.items():
                self.status_update.emit(f"پردازش پلتفرم: {key}...")
                
                df = pd.read_excel(file_path)
                df.replace(r'^\s*$', np.nan, regex=True, inplace=True)

                # ======== craeting filters ========
                
                if 'country' in df.columns and df['country'].notna().any() and self.filters['country'] not in ["همه کشورها", "----------"]:
                    country_filter = self.filters['country'].strip()
                    if country_filter == "فقط ایرانی":
                        df = df[df['country'].astype(str).str.contains("ایران", na=False, case=False)]
                    elif country_filter == "فقط خارجی":
                        df = df[~df['country'].astype(str).str.contains("ایران", na=False, case=False) & df['country'].notna()]
                    else:
                        pattern = r'(^|,|\s)' + re.escape(country_filter) + r'($|,|\s)'
                        df = df[df['country'].astype(str).str.contains(pattern, na=False, case=False, regex=True)]

                if 'age' in df.columns and df['age'].notna().any() and self.filters['age'] != "ترکیبی":
                    age_number_list = [s for s in self.filters['age'].split() if s.isdigit()]
                    if age_number_list:
                        pattern = r'\b' + age_number_list[0] + r'\b'
                        df = df[df['age'].astype(str).str.contains(pattern, na=False, regex=True)]

                if 'genre' in df.columns and df['genre'].notna().any() and self.filters['genre'] != "انواع ژانر":
                    genre_filter = self.filters['genre'].strip()
                    pattern = r'(^|,|\s)' + re.escape(genre_filter) + r'($|,|\s)'
                    df = df[df['genre'].astype(str).str.contains(pattern, na=False, case=False, regex=True)]

                if 'type' in df.columns and df['type'].notna().any():
                    if self.filters['film'] and not self.filters['series']: df = df[df['type'].astype(str).str.contains("فیلم", na=False, case=False)]
                    elif not self.filters['film'] and self.filters['series']: df = df[df['type'].astype(str).str.contains("سریال", na=False, case=False)]
                
                # filter publish
                if 'publish' in df.columns and df['publish'].notna().any():
                    #filter publish
                    df['publish_date'] = pd.to_datetime(df['publish'], errors='coerce')
                    df.dropna(subset=['publish_date'], inplace=True)
                    
                    start_date = self.filters['start_date']
                    end_date = self.filters['end_date']
                    
                    # creating date filter
                    df = df[(df['publish_date'].dt.date >= start_date) & (df['publish_date'].dt.date <= end_date)]
                    df = df.drop(columns=['publish_date'])

                if not df.empty:
                    filtered_dfs.append(df)
                
                file_count += 1
                self.progress_update.emit(int(5 + (file_count / total_files) * 75))

            if not filtered_dfs:
                self.error.emit("هیچ داده‌ای با فیلترهای انتخابی شما یافت نشد.")
                return

            self.status_update.emit("ترکیب داده‌های نهایی...")
            final_df = pd.concat(filtered_dfs, ignore_index=True).drop_duplicates()
            self.progress_update.emit(85)
            
            display_count = self.filters['display_count']
            if not final_df.empty and display_count > 0:
                 if len(final_df) > display_count:
                     final_df = final_df.sample(n=display_count)
            
            if final_df.empty:
                self.error.emit("هیچ داده‌ای پس از فیلتر کردن باقی نماند.")
                return

            self.status_update.emit("ذخیره فایل خروجی...")
            self.progress_update.emit(95)
            final_df.to_excel(self.output_path, index=False)
            self.progress_update.emit(100)
            self.finished.emit(self.output_path)

        except Exception as e:
            self.error.emit(f"یک خطای غیرمنتظره رخ داد: {e}\n{traceback.format_exc()}")


# =============================================================================
# main class for codes
# =============================================================================
class SearchWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.old_pos = None

        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        def resource_path(relative_path):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)

        self.ui.toolButton_5.setIcon(QIcon(resource_path("icon/close1.png")))
        self.ui.toolButton_4.setIcon(QIcon(resource_path("icon/minus1.png")))

        self.setStyleSheet(
            """QMainWindow {
                background: none;
        }
        """)

        # *** changing ui ***
        self.setup_date_filters()

        self.ui.progressBar = QProgressBar(self.ui.frame)
        self.ui.progressBar.setGeometry(QRect(260, 495, 181, 15))
        self.ui.progressBar.setStyleSheet("""
            QProgressBar { border: 1px solid #555; border-radius: 7px; text-align: center; background-color: #333; color: white; }
            QProgressBar::chunk { background-color: #4CAF50; border-radius: 7px; }
        """)
        self.ui.progressBar.setValue(0)
        self.ui.progressBar.hide()

        def resource_path(relative_path):
            try: base_path = sys._MEIPASS
            except Exception: base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)

        self.file_map = {
            self.ui.filimo:  {"key": "filimo",  "path": resource_path("فیلیمو.xlsx")},
            self.ui.filmnet: {"key": "filmnet", "path": resource_path("Filment_update.xlsx")},
            self.ui.gapfilm: {"key": "gapfilm", "path": resource_path("gapfilm.xlsx")},
            self.ui.opera:   {"key": "opera",   "path": resource_path("upera.xlsx")},
            self.ui.namava:  {"key": "namava",  "path": resource_path("Namava_14030728.xlsx")}
        }
        
        self.ui.createfile.clicked.connect(self.start_processing)
        
    def setup_date_filters(self):
        # hidding old widget
        self.ui.frame_2.hide()
        self.ui.label_3.hide()

        # creating date label
        self.date_label = QLabel("تاریخ انتشار", self.ui.filterframe)
        self.date_label.setGeometry(QRect(470, 10, 71, 16))
        font = self.ui.label_3.font() 
        self.date_label.setFont(font)
        self.date_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.date_label.show()
        
        # creating label از تاریخ
        self.from_date_label = QLabel("از:", self.ui.filterframe)
        self.from_date_label.setGeometry(QRect(540, 40, 31, 22))
        self.from_date_label.show()

        # creating label for start date
        self.startDateEdit = QDateEdit(self.ui.filterframe)
        self.startDateEdit.setGeometry(QRect(410, 40, 121, 22))
        self.startDateEdit.setCalendarPopup(True)
        self.startDateEdit.setDate(QDate(2000, 1, 1))
        self.startDateEdit.setDisplayFormat("yyyy-MM-dd")
        self.startDateEdit.show()

        # creating label تا تاریخ
        self.to_date_label = QLabel("تا:", self.ui.filterframe)
        self.to_date_label.setGeometry(QRect(540, 70, 31, 22))
        self.to_date_label.show()

        # creating end date
        self.endDateEdit = QDateEdit(self.ui.filterframe)
        self.endDateEdit.setGeometry(QRect(410, 70, 121, 22))
        self.endDateEdit.setCalendarPopup(True)
        self.endDateEdit.setDate(QDate(2030, 1, 1))
        self.endDateEdit.setDisplayFormat("yyyy-MM-dd")
        self.endDateEdit.show()

    def start_processing(self):
        selected_files_info = {}
        for checkbox, info in self.file_map.items():
            if checkbox.isChecked(): selected_files_info[info['key']] = info['path']

        if not selected_files_info:
            QMessageBox.warning(self, "خطا", "لطفاً حداقل یک پلتفرم را انتخاب کنید.")
            return

        filters = {
            'country': self.ui.country.currentText(), 'age': self.ui.age.currentText(),
            'genre': self.ui.zhaner.currentText(), 'film': self.ui.film.isChecked(),
            'series': self.ui.series.isChecked(),
            'start_date': self.startDateEdit.date().toPyDate(),
            'end_date': self.endDateEdit.date().toPyDate(),
            'display_count': self.ui.spinBox_5.value(),
            'jdatetime_module': None 
        }

        output_path, _ = QFileDialog.getSaveFileName(self, "ذخیره فایل خروجی", "", "Excel Files (*.xlsx)")
        if not output_path: return

        self.ui.createfile.setEnabled(False)
        self.ui.createfile.setText("...")
        self.ui.progressBar.setValue(0)
        self.ui.progressBar.show()

        self.thread = QThread()
        self.worker = ExcelWorker(selected_files_info, filters, output_path)
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_processing_finished)
        self.worker.error.connect(self.on_processing_error)
        self.worker.progress_update.connect(self.ui.progressBar.setValue)
        self.worker.status_update.connect(self.statusBar().showMessage)

        self.thread.start()

    def on_processing_finished(self, output_path):
        self.statusBar().showMessage(f"پردازش با موفقیت تمام شد.", 10000)
        QMessageBox.information(self, "موفقیت", f"فایل خروجی با موفقیت ایجاد شد:\n{output_path}")
        self.cleanup_thread()

    def on_processing_error(self, error_message):
        QMessageBox.critical(self, "خطا در پردازش", error_message)
        self.cleanup_thread()

    def cleanup_thread(self):
        self.ui.progressBar.hide()
        if hasattr(self, 'thread') and self.thread.isRunning():
            self.thread.quit()
            self.thread.wait()
        
        self.ui.createfile.setEnabled(True)
        self.ui.createfile.setText("ساخت فایل")

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton: self.old_pos = event.globalPosition().toPoint()

    def mouseMoveEvent(self, event):
        if self.old_pos and event.buttons() == Qt.MouseButton.LeftButton:
            delta = event.globalPosition().toPoint() - self.old_pos
            self.move(self.pos() + delta)
            self.old_pos = event.globalPosition().toPoint()
    
    def mouseReleaseEvent(self, event): self.old_pos = None

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)  # smooth edge

        rect = self.rect()

        background_color = QColor(28, 28, 29)  # background color
        painter.setBrush(background_color)
        painter.setPen(Qt.PenStyle.NoPen)
        # draw rounded rectangle background
        painter.drawRoundedRect(rect, 15, 15)

        border_color = QColor(255, 255, 255, 50)  # border color
        painter.setPen(QPen(border_color, 2))
        painter.setBrush(Qt.GlobalColor.transparent)
        painter.drawRoundedRect(rect, 15, 15)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SearchWindow()
    window.show()
    sys.exit(app.exec())
    sys.exit()
