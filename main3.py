import os
import smtplib
import threading
import pandas as pd
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from PySide6 import QtWidgets
from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QTextEdit,
    QVBoxLayout, QHBoxLayout, QFileDialog, QMessageBox, QListWidget,
    QFormLayout, QProgressBar
)
from PySide6.QtCore import Qt, Signal, QObject

recipient_list = []
attachment_paths = []
uploaded_files = []

class WorkerSignals(QObject):
    result = Signal(int, int)

class EmailApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Professional Email Automation App by Gauri Salunke")
        self.setMinimumSize(800, 600)
        self.signals = WorkerSignals()
        self.signals.result.connect(self.show_result)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        form_layout = QFormLayout()

        self.email_input = QLineEdit()
        form_layout.addRow("Your Gmail Email:", self.email_input)

        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        form_layout.addRow("App Password:", self.password_input)

        self.subject_input = QLineEdit()
        form_layout.addRow("Subject:", self.subject_input)

        self.message_input = QTextEdit()
        form_layout.addRow("Message:", self.message_input)

        layout.addLayout(form_layout)

        file_btn_layout = QHBoxLayout()

        upload_btn = QPushButton("Upload Excel/CSV")
        upload_btn.clicked.connect(self.browse_files)
        file_btn_layout.addWidget(upload_btn)

        attach_btn = QPushButton("Add Attachments")
        attach_btn.clicked.connect(self.browse_attachments)
        file_btn_layout.addWidget(attach_btn)

        layout.addLayout(file_btn_layout)

        self.uploaded_list_widget = QListWidget()
        layout.addWidget(QLabel("Uploaded Files:"))
        layout.addWidget(self.uploaded_list_widget)

        self.attachments_list_widget = QListWidget()
        layout.addWidget(QLabel("Attachments:"))
        layout.addWidget(self.attachments_list_widget)

        self.status_label = QLabel("")
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.status_label)
        layout.addWidget(self.progress_bar)

        self.send_btn = QPushButton("Send Emails")
        self.send_btn.setStyleSheet("font-weight: bold; background-color: #2196F3; color: white; height: 40px;")
        self.send_btn.clicked.connect(self.send_bulk_emails)
        layout.addWidget(self.send_btn)

        self.setLayout(layout)

    def browse_files(self):
        global recipient_list
        files, _ = QFileDialog.getOpenFileNames(self, "Select Excel/CSV files", "", "Excel files (*.xlsx);;CSV files (*.csv)")
        for file in files:
            extracted = self.extract_emails_and_names(file)
            if extracted:
                uploaded_files.append(os.path.basename(file))
                recipient_list.extend(extracted)
                self.uploaded_list_widget.addItem(os.path.basename(file))

    def extract_emails_and_names(self, file_path):
        try:
            if file_path.endswith('.xlsx'):
                data = pd.read_excel(file_path, engine='openpyxl')
            else:
                data = pd.read_csv(file_path)

            name_column = [col for col in data.columns if "name" in col.lower()]
            email_column = [col for col in data.columns if "email" in col.lower()]

            if not name_column or not email_column:
                raise ValueError("Required columns 'name' and/or 'email' not found.")

            names = data[name_column[0]].dropna().tolist()
            emails = data[email_column[0]].dropna().tolist()
            return list(zip(names, emails))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract names and emails:\n{e}")
            return []

    def browse_attachments(self):
        global attachment_paths
        files, _ = QFileDialog.getOpenFileNames(self, "Select Attachments")
        if files:
            attachment_paths.clear()
            attachment_paths.extend(files)
            self.attachments_list_widget.clear()
            for file in attachment_paths:
                self.attachments_list_widget.addItem(os.path.basename(file))

    def send_bulk_emails(self):
        sender_email = self.email_input.text().strip()
        sender_password = self.password_input.text().strip()
        subject = self.subject_input.text().strip()
        message = self.message_input.toPlainText().strip()

        if not sender_email or not sender_password or not subject or not message:
            QMessageBox.warning(self, "Warning", "Please fill in all fields.")
            return

        if not recipient_list:
            QMessageBox.warning(self, "Warning", "No recipients to send emails.")
            return

        self.send_btn.setEnabled(False)
        self.status_label.setText("Sending emails... Please wait.")
        self.progress_bar.setMaximum(len(recipient_list))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)

        threading.Thread(target=self.email_task, args=(sender_email, sender_password, subject, message), daemon=True).start()

    def email_task(self, sender_email, sender_password, subject, message):
        success_count = 0
        failure_count = 0

        for index, (recipient_name, recipient_email) in enumerate(recipient_list):
            result = self.send_email(sender_email, sender_password, recipient_name, recipient_email, subject, message)
            if result:
                success_count += 1
            else:
                failure_count += 1

            self.progress_bar.setValue(index + 1)

        self.signals.result.emit(success_count, failure_count)

    def show_result(self, success_count, failure_count):
        self.send_btn.setEnabled(True)
        self.status_label.setText("Email sending completed.")
        self.progress_bar.setVisible(False)
        QMessageBox.information(self, "Email Status", f"Emails sent: {success_count}\nFailed: {failure_count}")

    def send_email(self, sender_email, sender_password, recipient_name, recipient_email, subject, message):
        try:
            server = smtplib.SMTP('smtp.gmail.com', 587)
            server.starttls()
            server.login(sender_email, sender_password)

            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient_email
            msg['Subject'] = subject

            personalized_message = f"Hello {recipient_name},\n\n{message}"
            msg.attach(MIMEText(personalized_message, 'plain'))

            for file_path in attachment_paths:
                try:
                    with open(file_path, "rb") as f:
                        part = MIMEApplication(f.read(), Name=os.path.basename(file_path))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                        msg.attach(part)
                except Exception as e:
                    print(f"Could not attach {file_path}: {e}")

            server.send_message(msg)
            server.quit()
            return True
        except Exception as e:
            print(f"Failed to send to {recipient_email}: {e}")
            return False

if __name__ == "__main__":
    app = QApplication([])
    window = EmailApp()
    window.show()
    app.exec()
