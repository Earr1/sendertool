import sys, re, csv, smtplib, time
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QTextEdit, QPushButton, QFileDialog, QVBoxLayout,
    QHBoxLayout, QMessageBox, QListWidget, QFormLayout
)
from PyQt5.QtCore import Qt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication


class EmailMarketingApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Email Marketing Tool")
        self.setGeometry(100, 100, 800, 700)
        self.leads = []
        self.attachment_path = ""
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # Form
        form_layout = QFormLayout()
        self.smtp_input = QLineEdit("smtp.gmail.com")
        self.port_input = QLineEdit("587")
        self.user_input = QLineEdit()
        self.pass_input = QLineEdit()
        self.pass_input.setEchoMode(QLineEdit.Password)
        self.from_name = QLineEdit()
        self.from_email = QLineEdit()
        self.subject_input = QLineEdit()
        self.delay_input = QLineEdit("2")

        form_layout.addRow("SMTP Server:", self.smtp_input)
        form_layout.addRow("Port:", self.port_input)
        form_layout.addRow("Username:", self.user_input)
        form_layout.addRow("Password:", self.pass_input)
        form_layout.addRow("From Name:", self.from_name)
        form_layout.addRow("From Email:", self.from_email)
        form_layout.addRow("Subject:", self.subject_input)
        form_layout.addRow("Delay (sec):", self.delay_input)

        layout.addLayout(form_layout)

        # HTML body
        self.html_body = QTextEdit()
        self.html_body.setPlaceholderText("Write your HTML email here. Use {{FirstName}} etc.")
        layout.addWidget(QLabel("HTML Body:"))
        layout.addWidget(self.html_body)

        # Attachment
        attach_layout = QHBoxLayout()
        self.attach_input = QLineEdit()
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browseAttachment)
        attach_layout.addWidget(QLabel("Attachment:"))
        attach_layout.addWidget(self.attach_input)
        attach_layout.addWidget(browse_btn)
        layout.addLayout(attach_layout)

        # Buttons
        btn_layout = QHBoxLayout()
        load_btn = QPushButton("Load Leads (CSV)")
        load_btn.clicked.connect(self.loadLeads)

        preview_btn = QPushButton("Preview")
        preview_btn.clicked.connect(self.previewEmail)

        send_btn = QPushButton("Send Emails")
        send_btn.clicked.connect(self.sendEmails)

        btn_layout.addWidget(load_btn)
        btn_layout.addWidget(preview_btn)
        btn_layout.addWidget(send_btn)
        layout.addLayout(btn_layout)

        # Status box
        self.status_box = QListWidget()
        layout.addWidget(QLabel("Status:"))
        layout.addWidget(self.status_box)

        self.setLayout(layout)

    def browseAttachment(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Attachment")
        if file:
            self.attachment_path = file
            self.attach_input.setText(file)

    def loadLeads(self):
        file, _ = QFileDialog.getOpenFileName(self, "Load CSV", "", "CSV Files (*.csv)")
        if file:
            with open(file, newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                email_col = next((col for col in reader.fieldnames if "email" in col.lower()), None)
                if not email_col:
                    QMessageBox.warning(self, "Error", "No 'email' column found.")
                    return
                self.leads = [row for row in reader if re.match(r"[^@]+@[^@]+\.[^@]+", row.get(email_col, ""))]
                self.status_box.addItem(f"Loaded {len(self.leads)} valid leads.")

    def previewEmail(self):
        if not self.leads:
            QMessageBox.warning(self, "Error", "No leads loaded.")
            return
        sample = self.leads[0]
        body = self.html_body.toPlainText()
        preview = self.replacePlaceholders(body, sample)
        QMessageBox.information(self, "Email Preview", preview)

    def replacePlaceholders(self, text, data):
        for key, value in data.items():
            text = re.sub(r"{{\s*" + re.escape(key) + r"\s*}}", value, text)
        return text

    def sendEmails(self):
        if not self.leads:
            QMessageBox.warning(self, "Error", "No leads loaded.")
            return

        smtp_server = self.smtp_input.text()
        port = int(self.port_input.text())
        user = self.user_input.text()
        password = self.pass_input.text()
        from_name = self.from_name.text()
        from_email = self.from_email.text()
        subject = self.subject_input.text()
        delay = int(self.delay_input.text())
        body_template = self.html_body.toPlainText()

        try:
            server = smtplib.SMTP(smtp_server, port)
            server.starttls()
            server.login(user, password)
        except Exception as e:
            QMessageBox.critical(self, "SMTP Error", str(e))
            return

        for lead in self.leads:
            to_email = lead.get("email")
            msg = MIMEMultipart()
            msg['From'] = f"{from_name.strip()} <{from_email.strip()}>".replace('\n', '').replace('\r', '')
            msg['To'] = to_email
            msg['Subject'] = subject

            body = self.replacePlaceholders(body_template, lead)
            msg.attach(MIMEText(body, 'html'))

            if self.attachment_path:
                with open(self.attachment_path, "rb") as f:
                    part = MIMEApplication(f.read(), Name=self.attachment_path.split("/")[-1])
                    part['Content-Disposition'] = f'attachment; filename="{self.attachment_path.split("/")[-1]}"'
                    msg.attach(part)

            try:
                server.sendmail(from_email, to_email, msg.as_string())
                self.status_box.addItem(f"✓ Sent to {to_email}")
            except Exception as e:
                self.status_box.addItem(f"✗ Failed to {to_email}: {e}")
            QApplication.processEvents()
            time.sleep(delay)

        server.quit()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = EmailMarketingApp()
    win.show()
    sys.exit(app.exec_())
