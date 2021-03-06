from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from email.mime.text import MIMEText

import openpyxl
import sys
import os
import pickle
import base64
from time import sleep

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from core.parser import ContentParser

DEBUG = True


class Invitation():
    # Initialzie invitation structure form a excel row
    def __init__(self, row):
        LANG, MAIL, NAME, SENDER, FIELD, ONE_SEN, DATE, DESC, DONE, ETC = range(10)
        self.lang = row[LANG].value
        self.mail = row[MAIL].value
        self.name = row[NAME].value
        self.sender = row[SENDER].value
        self.field = row[FIELD].value
        self.one_sen = row[ONE_SEN].value
        self.date = row[DATE].value
        self.desc = row[DESC].value
        self.done = row[DONE].value == 'O'
        self.etc = row[ETC].value

    def is_eng(self):
        return self.lang == '영'

    def get_summary(self):
        return QStandardItem(self.name), QStandardItem(self.mail)

    def test(self):
        return "This is a test message"

    def batchim(self, name):
        last_char = list(name).pop()
        chk = (ord(last_char) - 44032) % 28
        if chk:
            return 1
        else:
            return 0
    
    def use_yi(self, chk):
        if chk:
            return '이'
        else:
            return ''

    def use_leul(self, chk):
        if chk:
            return '을'
        else:
            return '를'
    
    def send_invi_msg(self, service, user_id='me'):
        if self.is_eng():
            template = os.path.dirname(os.path.realpath(__file__)) + '/data/eng.json'
        else:
            template = os.path.dirname(os.path.realpath(__file__)) + '/data/kor.json'
        
        val = {
            'name': self.name,
            'sender': self.sender,
            'field': self.field,
            'date': self.date,
            'one_sen': self.one_sen,
            'leul': self.use_leul(self.batchim(self.name)),
            'yi': self.use_yi(self.batchim(self.sender))
        }
        parser = ContentParser(template = template, values = val)

        subject = parser.get_title()
        print("To: {:30}\nTitle: {:40}\n".format(str(self), parser.get_title()))
        # build msg
        msg_txt = parser.get_content()
        message = MIMEText(msg_txt, _charset = 'utf-8')
        message['subject'] = subject
        message['from'] = user_id
        message['to'] = self.mail

        raw = base64.urlsafe_b64encode(message.as_bytes())
        raw = raw.decode()
        msg_body = {'raw': raw}
        # send message
        if not DEBUG:
            sleep(2)
            print(msg_txt)
            message = (service.users().messages().send(userId=user_id, body=msg_body).execute())
        else:
            print("[!] DEBUG Mode, mails are not sent!")

    def __str__(self):
        return "{:30} {}".format(self.name, self.mail)


class MainUI(QWidget):
    def __init__(self):
        super().__init__()

        # MainUI Data Structures
        self.creds = None
        self.service = None
        self.user_profile = None
        self.user_email = ""
        self.is_logged_in = False

        # Mail Data Structures
        self.invitations = []
        self.mails = QStandardItemModel()
        self.mails.setHorizontalHeaderItem(0, QStandardItem('Name'))
        self.mails.setHorizontalHeaderItem(1, QStandardItem('Email Address'))

        # Box UI Structrue For Mail List
        box = QHBoxLayout()
        self.setLayout(box)

        # Grid UI Structure For Buttons and Labels
        grid = QGridLayout()
        box.addLayout(grid)

        self.login_btn = QPushButton('Login', self)
        self.login_btn.resize(self.login_btn.sizeHint())
        if not self.is_logged_in:
            self.login_btn.released.connect(self.login_gmail)
        else:
            self.login_btn.released.connect(self.reset_gmail)
        grid.addWidget(self.login_btn, 0, 0)

        self.login_label = QLabel('Login to Gmail Server', self)
        grid.addWidget(self.login_label, 0, 1)

        upload_btn = QPushButton('Upload', self)
        upload_btn.resize(upload_btn.sizeHint())
        upload_btn.released.connect(self.file_upload)
        grid.addWidget(upload_btn, 1, 0)

        file_label = QLabel('Contact Excel File(.xlsx)', self)
        grid.addWidget(file_label, 1, 1)

        check_btn = QPushButton('Check', self)
        check_btn.resize(check_btn.sizeHint())
        check_btn.released.connect(self.list_mails)
        grid.addWidget(check_btn, 2, 0)

        check_label = QLabel('Check if the mails are formed well', self)
        grid.addWidget(check_label, 2, 1)

        self.check_table = QTableView(self)
        self.check_table.doubleClicked.connect(self.show_email)
        box.addWidget(self.check_table)

        send_btn = QPushButton('Send', self)
        send_btn.resize(send_btn.sizeHint())
        send_btn.released.connect(self.ask_send)
        grid.addWidget(send_btn, 3, 0)

        send_label = QLabel('Send invitation mails', self)
        grid.addWidget(send_label, 3, 1)

    def login_gmail(self):
        def reset_gmail():
            os.remove('token.pickle')
            self.login_btn.setText('Login')
            self.login_label.setText('Login to Gmail Server')
            self.is_logged_in = False

        # Log Out
        if self.is_logged_in:
            reset_gmail()
            return

        SCOPES = [
            'https://www.googleapis.com/auth/gmail.send',
            'https://www.googleapis.com/auth/gmail.metadata',
        ]
        self.creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                credential = os.path.dirname(os.path.realpath(__file__)) + \
                    '/data/credentials.json'
                try:
                    flow = InstalledAppFlow.from_client_secrets_file(
                            credential, SCOPES)
                except FileNotFoundError as err:
                    err_box = QMessageBox.question(self, 'Error', "credentials.json not found",
                        QMessageBox.Yes, QMessageBox.Yes)
                    self.login_label.setText('Failed to log in to Gmail')
                    print("[!] Credentials Not Found at /data")
                    return
                self.creds = flow.run_local_server(port=8000)
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('gmail', 'v1', credentials=self.creds)
        self.user_file = self.service.users().getProfile(userId='me').execute()
        self.user_email = self.user_file['emailAddress']

        self.login_label.setText(self.user_email)
        self.login_btn.setText("Reset")
        self.is_logged_in = True
        print("[*] Logged in to Gmail: {}".format(self.user_email))

    def file_upload(self):
        filename = QFileDialog.getOpenFileName(self, 'Open file', './')

        if filename[0]:
            if not self.is_valid_xlsx(filename[0]):
                return
            contact_excel = openpyxl.load_workbook(filename=filename[0])
            contact_sheet = contact_excel['Sheet1']
            self.parse_excel_sheet(contact_sheet)
            print("[*] File Uploaded")

    # Need Improvement
    def parse_excel_sheet(self, sheet, header=True):
        for i, row in enumerate(sheet.iter_rows()):
            if header:
                if i == 0:
                    continue
            # Too many nones... ignore them!
            is_valid_row = True
            for cell in row[:7]:
                if cell.value == None:
                    is_valid_row = False
                    break
            if is_valid_row and not Invitation(row).done:
                self.invitations.append(Invitation(row))
            

    def is_valid_xlsx(self, filename):
        if not filename.endswith('.xlsx'):
            print("[!] Invalid File Extension")
            return False
        print("[*] Valid File Extension")
        return True

    def list_mails(self):
        for invi in self.invitations:
            print(invi)
            invi_item = invi.get_summary()
            self.mails.appendRow(invi_item)
        self.check_table.setModel(self.mails)

    def ask_send(self):
        ask = QMessageBox.question(self,
                'Message',
                "Are you sure to send mails?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if DEBUG:
            print("[*] DEBUG mode is on")
        if ask == QMessageBox.Yes:
            for i, invi in enumerate(self.invitations, 1):
                self.send_mails(invi, self.service)
                ex.statusBar().showMessage(str(i)+'/'+str(len(self.invitations)))

    def send_mails(self, invi, service, user_id='me'):
        invi.send_invi_msg(service)

    def show_email(self):
        for idx in self.check_table.selectionModel().selectedIndexes():
            print(idx.row(), idx.column())

class MainApp(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        '''
        UI Configuration
        0. GMail Crudential Validation
        1. Excel File Uploading
        2. Check the content of each mail
        3. Send mails
        '''

        central_wg = MainUI()
        self.setCentralWidget(central_wg)
        # StatusBar Configration
        self.statusBar().showMessage('Status')

        # Windows Configuration
        self.setWindowTitle("ICISTS Mail Management")
        self.resize(800, 600)
        self.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MainApp()
    sys.exit(app.exec_())
