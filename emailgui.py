import sys
import json
import os
from datetime import datetime
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QTableWidget, QTableWidgetItem, QPushButton, 
                            QLabel, QLineEdit, QComboBox, QFileDialog, QMessageBox,
                            QCompleter, QHeaderView)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import QShortcut
from PyQt5.QtGui import QKeySequence




import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime
import os
from email.utils import formatdate

from themes import Themes
import time

class ResearchAreas:
    def __init__(self):
        with open('research_areas.json', 'r') as f:
            self.AREAS = json.load(f)
        
    @classmethod
    def get_all_areas(cls):
        areas = []
        with open('research_areas.json', 'r') as f:
            AREAS = json.load(f)
            for total_genre in AREAS.keys():
                subgenreslist = AREAS[total_genre]
                areas.extend(subgenreslist)
        return areas



class SubstringCompleter(QCompleter):
    def __init__(self, items, parent=None):
        super().__init__(items, parent)
        self.setFilterMode(Qt.MatchFlag.MatchContains)
        self.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)

class PreventScrollComboBox(QComboBox):
    def wheelEvent(self, event):
        # Prevent the default wheel event from changing the selection
        event.ignore()

class CustomTableWidget(QTableWidget):
    def __init__(self, headers, is_primary=False, parent=None):
        super().__init__(parent)
        self.is_primary = is_primary
        self.setColumnCount(len(headers))
        self.setHorizontalHeaderLabels(headers)
        self.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.followup_table = None  # Will be set later for primary table
       
    def add_row(self):
        row_position = self.rowCount()
        self.insertRow(row_position)
       
        # Add ComboBox for Working Area with wheel event prevention and substring completer
        area_combo = PreventScrollComboBox()
        all_areas = ResearchAreas.get_all_areas()
        area_combo.addItems(all_areas)
        area_combo.setEditable(True)
        
        # Use custom substring completer
        completer = SubstringCompleter(all_areas)
        area_combo.setCompleter(completer)
        
        self.setCellWidget(row_position, 2, area_combo)
       
        # Add ComboBox for Response Status
        status_combo = QComboBox()
        status_combo.addItems(["", "Yes", "No"])
        self.setCellWidget(row_position, 4, status_combo)
       
        # Add Remove button
        remove_btn = QPushButton("Remove")
        remove_btn.clicked.connect(lambda: self.removeRow(self.indexAt(remove_btn.pos()).row()))
        self.setCellWidget(row_position, 5, remove_btn)
       
        # Add Move to Follow-up button only for primary table
        if self.is_primary:
            followup_btn = QPushButton("Move to Follow-up")
            followup_btn.clicked.connect(lambda: self.move_to_followup(self.indexAt(followup_btn.pos()).row()))
            self.setCellWidget(row_position, 6, followup_btn)

    def move_to_followup(self, row):
        if not self.followup_table:
            return
       
        # Add new row to follow-up table
        self.followup_table.add_row()
        target_row = self.followup_table.rowCount() - 1
       
        # Copy data from primary to follow-up
        for col in range(5):  # Only copy until the "Remove" button column
            if isinstance(self.cellWidget(row, col), QComboBox):
                # Copy combo box selection
                source_combo = self.cellWidget(row, col)
                target_combo = self.followup_table.cellWidget(target_row, col)
                target_combo.setCurrentText(source_combo.currentText())
            else:
                # Copy regular cell content
                item = self.item(row, col)
                if item:
                    self.followup_table.setItem(target_row, col,
                                              QTableWidgetItem(item.text()))
       
        # Remove row from primary table
        self.removeRow(row)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.current_theme = "light"
        self.font_size = 12


        self.setWindowTitle("Professor Email Manager")
        self.setGeometry(100, 100, 1200, 800)
        
        # Create main widget and layout
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        
        # Create credentials section
        cred_layout = QHBoxLayout()
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("Your Email")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Your Password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.cv_path = QLineEdit()
        self.cv_path.setPlaceholderText("Path to CV")
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_cv)
        
        cred_layout.addWidget(QLabel("Email:"))
        cred_layout.addWidget(self.email_input)
        cred_layout.addWidget(QLabel("Password:"))
        cred_layout.addWidget(self.password_input)
        cred_layout.addWidget(QLabel("CV:"))
        cred_layout.addWidget(self.cv_path)
        cred_layout.addWidget(browse_btn)
        layout.addLayout(cred_layout)

        self.AREAS = {}
        with open('research_areas.json', 'r') as f:
            self.AREAS = json.load(f)
        
        # Create primary table
        self.primary_table = CustomTableWidget(
            ["Family Name", "Email", "Working Area", "Last Email Date", 
             "Response Status", "Remove", "Move to Follow-up"],
            is_primary=True
        )
        layout.addWidget(QLabel("Primary Table"))
        layout.addWidget(self.primary_table)
        
        # Create control buttons
        btn_layout = QHBoxLayout()
        add_row_btn = QPushButton("Add Row")
        add_row_btn.clicked.connect(self.primary_table.add_row)
        send_btn = QPushButton("Send Emails")
        send_btn.clicked.connect(self.send_emails)
        delte_duplicate_btn = QPushButton("Delete Duplicates")
        delte_duplicate_btn.clicked.connect(self.delete_duplicates)
        moveall_btn = QPushButton("Move All to Follow-up")
        moveall_btn.clicked.connect(self.move_all_to_followup)
        btn_layout.addWidget(add_row_btn)
        btn_layout.addWidget(send_btn)
        btn_layout.addWidget(delte_duplicate_btn)
        btn_layout.addWidget(moveall_btn)

        layout.addLayout(btn_layout)

        # Create follow-up table
        self.followup_table = CustomTableWidget(
            ["Family Name", "Email", "Working Area", "Last Email Date", 
             "Response Status", "Remove"]
        )
        layout.addWidget(QLabel("Follow-up Table"))
        layout.addWidget(self.followup_table)
        
        # Link the follow-up table to the primary table
        self.primary_table.followup_table = self.followup_table
        
        # create push buuton for  follow-up table
        followup_btn_layout = QHBoxLayout()
        add_followup_row_btn = QPushButton("Add follow-up Row")
        add_followup_row_btn.clicked.connect(self.followup_table.add_row)
        send_followup_btn = QPushButton("Send Follow-up Emails")
        send_followup_btn.clicked.connect(self.send_followup_emails)
        delte_duplicate_btn = QPushButton("Delete Duplicates")
        delte_duplicate_btn.clicked.connect(self.delete_followup_duplicates)

        followup_btn_layout.addWidget(add_followup_row_btn)
        followup_btn_layout.addWidget(send_followup_btn)
        followup_btn_layout.addWidget(delte_duplicate_btn)






        layout.addLayout(followup_btn_layout)






        # Load saved data
        self.load_data()
        
        # Set up auto-save
        self.setup_autosave()

        # Apply theme
        self.apply_theme()

        # Set the setting if I press the ctrl + + or ctrl + - it will change the font size
        self.shortcut_plus = QShortcut(QKeySequence("Ctrl+w"), self)
        self.shortcut_minus = QShortcut(QKeySequence("Ctrl+s"), self)

        self.shortcut_plus.activated.connect(self.increase_font_size)
        self.shortcut_minus.activated.connect(self.decrease_font_size)

    def increase_font_size(self):
        self.font_size += 0.5
        self.apply_theme()

    def decrease_font_size(self):
        self.font_size -= 0.5
        self.apply_theme()


    def move_all_to_followup(self):
        # Move all rows from primary table to follow-up table if they the last email date is not empty
        for row in range(self.primary_table.rowCount() - 1, -1, -1):
            try:
                # Get the last email date from the current row
                last_email_date = self.primary_table.item(row, 3).text()
            except AttributeError:
                # Skip rows without an email
                continue
            
            # Check if the last email date is not empty
            if last_email_date is not None or last_email_date != "":
                # push the move to follow-up button
                move_btn = self.primary_table.cellWidget(row, 6)
                move_btn.click()
            else:
                print(f"Error: {last_email_date} is empty for {self.primary_table.item(row, 0).text()}")
            
                
                



    
    def browse_cv(self):
        file_name, _ = QFileDialog.getOpenFileName(
            self, "Select CV", "", "PDF Files (*.pdf)"
        )
        if file_name:
            self.cv_path.setText(file_name)
            self.save_data()

    def delete_duplicates(self):
        # Create a set to track unique emails
        emails = set()

        # first iterate through the followup table
        for i in range(self.followup_table.rowCount() - 1, -1, -1):
            try:
                # Get the email from the current row
                email = self.followup_table.item(i, 1).text()
            except AttributeError:
                # Skip rows without an email
                continue
            
            # add the email to the set
            emails.add(email)
        
        # Iterate through the primary table in reverse order to safely remove rows
        for i in range(self.primary_table.rowCount() - 1, -1, -1):
            try:
                # Get the email from the current row
                email = self.primary_table.item(i, 1).text()
            except AttributeError:
                # Skip rows without an email
                continue
            
            # Check if email is a duplicate
            if email in emails:
                # Remove the duplicate row
                self.primary_table.removeRow(i)
            else:
                # Add unique email to the set
                emails.add(email)
    

    
    def delete_followup_duplicates(self):
        # Create a set to track unique emails
        emails = set()
        
        # Iterate through the primary table in reverse order to safely remove rows
        for i in range(self.followup_table.rowCount() - 1, -1, -1):
            try:
                # Get the email from the current row
                email = self.followup_table.item(i, 1).text()
            except AttributeError:
                # Skip rows without an email
                continue
            
            # Check if email is a duplicate
            if email in emails:
                # Remove the duplicate row
                self.followup_table.removeRow(i)
            else:
                # Add unique email to the set
                emails.add(email)
                
            


    def send_email_with_template(self,smtp_server, smtp_port, gmail_user, gmail_password, to_email, subject, template_path, cv_path, professor_name, total_genre, subgenre):
        """Send an email with a template file."""
        try:
            template_content = f'''Dear Professor {professor_name},

I hope this message finds you well. My name is Elina Adibi, and I am a senior at CSU Monterey Bay majoring in Computer Science with a 3.9 GPA. My academic experiences, including teaching assistantship and project-based learning, have provided me with a strong foundation in {total_genre}.

As a Teaching Assistant Coordinator and Data Structures TA, I have led lab sessions, mentored students, and conducted technical evaluations. My project work spans full-stack development and mobile applications, such as creating a Christmas Wishlist app using Spring Boot and React and developing a live music event discovery app with React Native. These experiences have honed my technical and collaborative skills, preparing me for research and problem-solving challenges in {total_genre}.

I am particularly interested in your research on {subgenre}. I find this area exciting due to its potential to drive innovation and address complex challenges in {subgenre}, and I would be honored to contribute to your lab’s ongoing work.

I am exploring graduate opportunities for Fall 2025 and would be thrilled to discuss potential openings in your lab, including funded positions. My resume is attached for your reference, and I would welcome the opportunity to learn more about your research and how I could contribute.

Thank you for your time and consideration.

Best regards,
Elina Adibi
fadibi@csumb.edu'''

            # Personalize the template (optional)
            # Example: template_content = template_content.replace("{name}", "Professor X")

            # Create email message
            msg = MIMEMultipart()
            msg['From'] = gmail_user
            msg['To'] = to_email
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = subject

            # Attach the email body (template content)
            msg.attach(MIMEText(template_content, 'plain'))

            # Attach the PDF file
            if not os.path.exists(cv_path):
                raise FileNotFoundError(f"Attachment file {cv_path} not found.")
            with open(cv_path, 'rb') as attachment_file:
                attachment = MIMEApplication(attachment_file.read(), _subtype="pdf")
                attachment.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{os.path.basename(cv_path)}"'
                )
                msg.attach(attachment)


            # Set up the SMTP server
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()  # Secure the connection
                server.login(gmail_user, gmail_password)  # Log in to your email account
                server.sendmail(gmail_user, to_email, msg.as_string())  # Send the email

            print(f"Email sent successfully to {to_email}!")

        except Exception as e:
            print(f"Failed to send email: {e}")
    
    def send_emails(self):
        # Implementation of email sending functionality
        try:
            # Basic email validation
            if not self.email_input.text() or not self.password_input.text():
                QMessageBox.warning(self, "Error", "Please enter email credentials")
                return
            
            if not self.cv_path.text() or not os.path.exists(self.cv_path.text()):
                QMessageBox.warning(self, "Error", "Please select a valid CV file")
                return
            
            # Email sending logic would go here
            # For security reasons, actual implementation would need proper
            # email server configuration and error handling
                # Configuration
            for row in range(self.primary_table.rowCount()):
                time.sleep(4)
                SMTP_SERVER = 'smtp.gmail.com'
                SMTP_PORT = 587
                GMAIL_USER = self.email_input.text()  # Replace with your Gmail address
                GMAIL_PASSWORD = self.password_input.text()    # Replace with your Gmail app password
                TO_EMAIL = self.primary_table.item(row, 1).text()  # Replace with recipient's email address
                TEMPLATE_PATH = 'email_template.txt'  # Path to your email template file
                CV_PATH = self.cv_path.text()
                professor_name = self.primary_table.item(row, 0).text()
                subgenre = self.primary_table.cellWidget(row, 2).currentText()

                # find the total genre
                total_genre = ""
                for main_area, subareas in self.AREAS.items():
                    if subgenre in subareas:
                        total_genre = main_area
                        break
                
                if total_genre == "":
                    # print(f"Error:{subgenre} not found in ResearchAreas for {professor_name}")
                    # continue
                    total_genre = "computer science"

                SUBJECT = "Prospective Graduate Student for Fall 2025"

                try:
                    self.send_email_with_template(SMTP_SERVER, SMTP_PORT, GMAIL_USER, GMAIL_PASSWORD, TO_EMAIL, SUBJECT, TEMPLATE_PATH, CV_PATH, professor_name, total_genre, subgenre)
                
                    # set the last email date
                    self.primary_table.setItem(row, 3, QTableWidgetItem(str(datetime.now().date())))
                except Exception as e:
                    print(f"Failed to send email to {TO_EMAIL}: {e}")
                        
            QMessageBox.information(self, "Success", "Emails sent successfully")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send emails: {str(e)}")
    
    def save_data(self):
        # Implementation of data saving functionality
        data = {
            "email": self.email_input.text(),
            "password": self.password_input.text(),
            "cv_path": self.cv_path.text(),
            "primary_table": self.get_table_data(self.primary_table),
            "followup_table": self.get_table_data(self.followup_table)
        }
        
        try:
            with open("./app_data.json", "w") as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save data: {str(e)}")
    
    def load_data(self):
        try:
            if os.path.exists("./app_data.json"):
                with open("./app_data.json", "r") as f:
                    data = json.load(f)
                
                self.email_input.setText(data.get("email", ""))
                self.cv_path.setText(data.get("cv_path", ""))
                self.password_input.setText(data.get("password", ""))
                
                # Load table data
                self.load_table_data(self.primary_table, data.get("primary_table", []))
                self.load_table_data(self.followup_table, data.get("followup_table", []))
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to load data: {str(e)}")
    
    def get_table_data(self, table):
        data = []
        for row in range(table.rowCount()):
            row_data = []
            for col in range(table.columnCount()):
                if isinstance(table.cellWidget(row, col), QComboBox):
                    row_data.append(table.cellWidget(row, col).currentText())
                elif isinstance(table.cellWidget(row, col), QPushButton):
                    continue
                else:
                    item = table.item(row, col)
                    row_data.append(item.text() if item else "")
            data.append(row_data)
        return data
    
    def load_table_data(self, table, data):
        for row_data in data:
            table.add_row()
            row = table.rowCount() - 1
            for col, value in enumerate(row_data):
                if isinstance(table.cellWidget(row, col), QComboBox):
                    table.cellWidget(row, col).setCurrentText(value)
                elif not isinstance(table.cellWidget(row, col), QPushButton):
                    table.setItem(row, col, QTableWidgetItem(value))
    
    def setup_autosave(self):
        # Set up timer for autosave
        from PyQt5.QtCore import QTimer
        self.autosave_timer = QTimer()
        self.autosave_timer.timeout.connect(self.save_data)
        self.autosave_timer.start(1800000)  # Autosave every 3 minutes

    def closeEvent(self, event):
        QMessageBox.information(self, "Info", "Saving data before closing")
        self.save_data()  # Explicitly save data when closing
        event.accept()

    def apply_theme(self):
        style = Themes.get_dynamic_style(self.current_theme, self.font_size)
        self.setStyleSheet(style)

    def send_followup_emails(self):
        # Implementation of email sending functionality
        try:
            # Basic email validation
            if not self.email_input.text() or not self.password_input.text():
                QMessageBox.warning(self, "Error", "Please enter email credentials")
                return
            
            if not self.cv_path.text() or not os.path.exists(self.cv_path.text()):
                QMessageBox.warning(self, "Error", "Please select a valid CV file")
                return
            
            # Email sending logic would go here
            # For security reasons, actual implementation would need proper
            # email server configuration and error handling
                # Configuration
            for row in range(self.followup_table.rowCount()):
                time.sleep(4)
                SMTP_SERVER = 'smtp.gmail.com'
                SMTP_PORT = 587
                GMAIL_USER = self.email_input.text()  # Replace with your Gmail address
                GMAIL_PASSWORD = self.password_input.text()    # Replace with your Gmail app password
                TO_EMAIL = self.followup_table.item(row, 1).text()  # Replace with recipient's email address
                TEMPLATE_PATH = 'email_template.txt'  # Path to your email template file
                CV_PATH = self.cv_path.text()
                professor_name = self.followup_table.item(row, 0).text()
                subgenre = self.followup_table.cellWidget(row, 2).currentText()

                # find the total genre
                total_genre = ""
                for main_area, subareas in ResearchAreas.AREAS.items():
                    if subgenre in subareas:
                        total_genre = main_area
                        break
                
                if total_genre == "":
                    print(f"Error: {subgenre} not found in ResearchAreas for {professor_name}")
                    continue

                SUBJECT = "Prospective Graduate Student for Fall 2025 (Follow-up)"

                try:
                    self.send_email_with_template(SMTP_SERVER, SMTP_PORT, GMAIL_USER, GMAIL_PASSWORD, TO_EMAIL, SUBJECT, TEMPLATE_PATH, CV_PATH, professor_name, total_genre, subgenre)

                except Exception as e:
                    print(f"Failed to send email to {TO_EMAIL}: {e}")
                        
            QMessageBox.information(self, "Success", "follow up emails sent successfully")
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to send emails: {str(e)}")
    

    

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())