import sys
import imaplib
import email
import os
import ssl
from email.header import decode_header
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QLineEdit, QPushButton, QTextEdit, QProgressBar, 
                             QFileDialog, QMessageBox, QGroupBox, QDateEdit, QRadioButton)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QDate
from PyQt5.QtGui import QFont

class EmailWorker(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal(int, int, str)
    error = pyqtSignal(str)

    def __init__(self, email_address, password, save_folder, start_date, download_sent):
        super().__init__()
        self.email_address = email_address
        self.password = password
        self.save_folder = save_folder
        self.start_date = start_date
        self.download_sent = download_sent

    def run(self):
        try:
            folder_type = "отправленных" if self.download_sent else "входящих"
            self.progress.emit(f"Подключаемся к серверу Mail.ru для загрузки {folder_type} писем...")
            
            context = ssl.create_default_context()
            mail = imaplib.IMAP4_SSL("imap.mail.ru", ssl_context=context)
            mail.login(self.email_address, self.password)
            
            if self.download_sent:
                folder_names = ['Sent', 'Отправленные', '&BB4EQgQ,BEAEMAQyBDsENQQ9BD0ESwQ1-', 'Sent Items']
            else:
                folder_names = ['inbox', 'INBOX']
            
            selected_folder = None
            for folder in folder_names:
                try:
                    status, response = mail.select(folder)
                    if status == 'OK':
                        selected_folder = folder
                        self.progress.emit(f"Выбрана папка: {folder}")
                        break
                except:
                    continue
            
            if not selected_folder:
                self.error.emit("Не удалось найти папку!")
                return
            
            imap_date = self.start_date.strftime('%d-%b-%Y')
            self.progress.emit(f"Ищем письма начиная с: {imap_date}")
            
            status, messages = mail.search(None, f'(SINCE "{imap_date}")')
            
            if status != "OK":
                self.error.emit("Ошибка при поиске писем")
                return
            
            email_ids = messages[0].split()
            self.progress.emit(f"Найдено {folder_type} писем: {len(email_ids)}")
            
            if not email_ids:
                self.progress.emit("Не найдено писем за указанный период")
                self.finished.emit(0, 0, self.save_folder)
                return
            
            saved_count = 0
            skipped_count = 0
            
            for i, email_id in enumerate(email_ids):
                try:
                    status, msg_data = mail.fetch(email_id, "(RFC822)")
                    
                    if status != "OK":
                        continue
                    
                    msg = email.message_from_bytes(msg_data[0][1])
                    
                    # Проверяем дату письма
                    email_date_str = msg.get("Date")
                    email_date = self.parse_email_date(email_date_str) if email_date_str else None
                    
                    if email_date and not self.is_date_after_start(email_date, self.start_date):
                        skipped_count += 1
                        continue
                    
                    # Получаем тему письма
                    subject = "Без_темы"
                    if msg["Subject"]:
                        try:
                            subject_header, encoding = decode_header(msg["Subject"])[0]
                            if isinstance(subject_header, bytes):
                                subject = subject_header.decode(encoding if encoding else "utf-8")
                            else:
                                subject = subject_header
                        except:
                            subject = "Ошибка_декодирования_темы"
                    
                    # Создаем безопасное имя файла
                    safe_subject = "".join(c if c.isalnum() or c in " -_." else "_" for c in subject)
                    safe_subject = safe_subject[:50]
                    
                    # Формируем имя файла
                    prefix = "SENT" if self.download_sent else "INBOX"
                    date_for_filename = email_date.strftime("%Y%m%d") if email_date else "nodate"
                    filename = f"{prefix}_{saved_count+1:04d}_{date_for_filename}_{safe_subject}.eml"
                    filepath = os.path.join(self.save_folder, filename)
                    
                    # Сохраняем письмо
                    with open(filepath, "wb") as f:
                        f.write(msg_data[0][1])
                    
                    self.progress.emit(f"Сохранено: {filename}")
                    saved_count += 1
                    
                except Exception as e:
                    self.progress.emit(f"Ошибка при обработке письма: {str(e)}")
                    continue
            
            mail.close()
            mail.logout()
            self.finished.emit(saved_count, len(email_ids), self.save_folder)
            
        except Exception as e:
            self.error.emit(f"Ошибка подключения: {str(e)}")

    def parse_email_date(self, date_str):
        try:
            return parsedate_to_datetime(date_str)
        except:
            return None

    def is_date_after_start(self, email_date, start_date):
        if email_date is None:
            return False
        
        if email_date.tzinfo is None:
            email_date = email_date.replace(tzinfo=timezone.utc)
        if start_date.tzinfo is None:
            start_date = start_date.replace(tzinfo=timezone.utc)
        
        return email_date >= start_date

class EmailBackupApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mail.ru Backup Tool")
        self.setGeometry(100, 100, 800, 600)
        self.setup_ui()
        
        # Установка красивого стиля для всего приложения
        self.setStyleSheet("""
            QMainWindow {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #41B3A3, stop: 1 #e9ecef);
            }
            
            QGroupBox {
                font-weight: bold;
                border: 2px solid #dee2e6;
                border-radius: 8px;
                margin-top: 1ex;
                padding-top: 10px;
                background-color: white;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #495057;
            }
            
            QLineEdit, QDateEdit {
                padding: 8px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                background-color: white;
                font-size: 12px;
            }
            
            QTextEdit {
                background-color: white;
                border: 1px solid #ced4da;
                border-radius: 4px;
                font-family: 'Courier New';
                font-size: 11px;
            }
            
            QProgressBar {
                border: 1px solid #ced4da;
                border-radius: 4px;
                text-align: center;
                background-color: white;
            }
            
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
            
            QRadioButton {
                font-weight: normal;
                padding: 4px;
            }
            
            QRadioButton::indicator {
                width: 16px;
                height: 16px;
            }
            
            QRadioButton::indicator:checked {
                background-color: #4CAF50;
                border: 2px solid #45a049;
                border-radius: 8px;
            }
        """)
        
    def setup_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Заголовок
        title = QLabel("Mail Saver")
        title.setFont(QFont("Courier new", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #2c3e50; margin-bottom: 20px;")
        layout.addWidget(title)
        
        # Поля ввода
        group = QGroupBox("Настройки")
        group_layout = QVBoxLayout()
        group_layout.setSpacing(12)
        
        # Email
        email_layout = QHBoxLayout()
        email_label = QLabel("Email:")
        email_label.setFixedWidth(150)
        email_layout.addWidget(email_label)
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("ваш@mail.ru")
        email_layout.addWidget(self.email_input)
        group_layout.addLayout(email_layout)
        
        # Пароль
        password_layout = QHBoxLayout()
        password_label = QLabel("Пароль приложения:")
        password_label.setFixedWidth(150)
        password_layout.addWidget(password_label)
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Пароль из настроек безопасности")
        password_layout.addWidget(self.password_input)
        group_layout.addLayout(password_layout)
        
        # Дата начала
        date_layout = QHBoxLayout()
        date_label = QLabel("Начиная с даты:")
        date_label.setFixedWidth(150)
        date_layout.addWidget(date_label)
        self.date_input = QDateEdit()
        self.date_input.setDate(QDate(2022, 10, 21))
        self.date_input.setCalendarPopup(True)
        date_layout.addWidget(self.date_input)
        group_layout.addLayout(date_layout)
        
        # Тип писем (Radio buttons)
        email_type_group = QGroupBox("Тип писем для сохранения")
        email_type_layout = QVBoxLayout()
        
        self.inbox_radio = QRadioButton("Только входящие письма")
        self.sent_radio = QRadioButton("Только отправленные письма")
        self.sent_radio.setChecked(True)
        
        email_type_layout.addWidget(self.inbox_radio)
        email_type_layout.addWidget(self.sent_radio)
        email_type_group.setLayout(email_type_layout)
        group_layout.addWidget(email_type_group)
        
        # Папка сохранения
        folder_layout = QHBoxLayout()
        folder_label = QLabel("Папка сохранения:")
        folder_label.setFixedWidth(150)
        folder_layout.addWidget(folder_label)
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("Выберите папку для сохранения")
        folder_layout.addWidget(self.folder_input)
        
        self.browse_btn = QPushButton("Обзор...")
        self.browse_btn.clicked.connect(self.browse_folder)
        self.browse_btn.setStyleSheet("""
            QPushButton {
                background-color: #41B3A3;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #41B3A3;
            }
        """)
        folder_layout.addWidget(self.browse_btn)
        group_layout.addLayout(folder_layout)
        
        group.setLayout(group_layout)
        layout.addWidget(group)
        
        # Кнопки
        button_layout = QHBoxLayout()
        button_layout.setSpacing(15)
        
        self.start_btn = QPushButton("Начать сохранение")
        self.start_btn.clicked.connect(self.start_backup)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #baacc7;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #c691f8;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        self.stop_btn = QPushButton("Остановить")
        self.stop_btn.clicked.connect(self.stop_backup)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("""
            QPushButton {
                background-color: #f44336;
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 6px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #da190b;
            }
            QPushButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)
        
        button_layout.addWidget(self.start_btn)
        button_layout.addWidget(self.stop_btn)
        layout.addLayout(button_layout)
        
        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        # Лог
        log_label = QLabel("Лог выполнения:")
        log_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        layout.addWidget(log_label)
        
        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        self.log_output.setMaximumHeight(200)
        layout.addWidget(self.log_output)
        
        self.worker = None
        
    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку для сохранения")
        if folder:
            self.folder_input.setText(folder)
            
    def log(self, message):
        self.log_output.append(message)
        
    def start_backup(self):
        email = self.email_input.text().strip()
        password = self.password_input.text().strip()
        folder = self.folder_input.text().strip()
        
        if not email or not password or not folder:
            QMessageBox.warning(self, "Ошибка", "Заполните все поля!")
            return
            
        if not os.path.exists(folder):
            try:
                os.makedirs(folder)
            except:
                QMessageBox.warning(self, "Ошибка", "Не удалось создать папку!")
                return
        
        start_date = self.date_input.date().toPyDate()
        start_datetime = datetime(start_date.year, start_date.month, start_date.day, tzinfo=timezone.utc)
        
        # Определяем тип писем из radio buttons
        download_sent = self.sent_radio.isChecked()
        email_type = "отправленных" if download_sent else "входящих"
        
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.log_output.clear()
        
        self.log("=" * 60)
        self.log("🚀 ЗАПУСК СОХРАНЕНИЯ ПИСЕМ")
        self.log("=" * 60)
        self.log(f"📧 Email: {email}")
        self.log(f"📁 Папка: {folder}")
        self.log(f"📅 Дата начала: {start_date.strftime('%d.%m.%Y')}")
        self.log(f"📦 Тип: {email_type} письма")
        self.log("=" * 60)
        
        self.worker = EmailWorker(email, password, folder, start_datetime, download_sent)
        self.worker.progress.connect(self.log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()
        
    def stop_backup(self):
        if self.worker and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait()
            self.log("⏹️ Остановлено пользователем")
            self.reset_ui()
            
    def on_finished(self, saved_count, total_count, folder):
        self.log("=" * 60)
        self.log("✅ ОБРАБОТКА ЗАВЕРШЕНА")
        self.log(f"📊 Найдено писем: {total_count}")
        self.log(f"💾 Сохранено писем: {saved_count}")
        self.log(f"📂 Папка: {folder}")
        self.log("=" * 60)
        
        QMessageBox.information(self, "Готово", 
                               f"Сохранено писем: {saved_count}\nИз них: {total_count}\n\nПапка: {folder}")
        self.reset_ui()
        
    def on_error(self, error_message):
        self.log(f"❌ ОШИБКА: {error_message}")
        QMessageBox.critical(self, "Ошибка", error_message)
        self.reset_ui()
        
    def reset_ui(self):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.progress_bar.setVisible(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Установка стиля для всего приложения
    app.setStyle("Fusion")
    
    window = EmailBackupApp()
    window.show()
    sys.exit(app.exec_())
