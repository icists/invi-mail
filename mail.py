from PyQt5.QtWidgets import *
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import QCoreApplication
from email.mime.text import MIMEText
import openpyxl
import sys, os
import pickle

from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

class Invitation():
    # Initialzie invitation structure form a excel row
    def __init__(self, row):
        LANG, MAIL, NAME, SENDER, FIELD, ONE_SEN, DATE, DESC, DONE, ETC = \
                tuple(number for number in range(10))
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

    def __str__(self):
        return "{}, {}".format(self.name, self.mail)


class MainUI(QWidget):
    def __init__(self):
        super().__init__()

        # MainUI Data Structures
        self.creds = None
        self.invitations = []

        # Grid UI Structure
        grid = QGridLayout()
        self.setLayout(grid)

        login_btn = QPushButton('Login', self)
        login_btn.resize(login_btn.sizeHint())
        login_btn.released.connect(self.login_gmail)
        grid.addWidget(login_btn, 0, 0)

        login_label = QLabel('Login to Gmail Server', self)
        grid.addWidget(login_label, 0, 1)

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

        send_btn = QPushButton('Send', self)
        send_btn.resize(send_btn.sizeHint())
        send_btn.released.connect(self.send_mails)
        grid.addWidget(send_btn, 3, 0)

        send_label = QLabel('Send invitation mails', self)
        grid.addWidget(send_label, 3, 1)

    def login_gmail(self):
        SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        self.creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                credential = os.path.dirname(os.path.realpath(__file__)) + '/credentials.json'
                flow = InstalledAppFlow.from_client_secrets_file(
                        credential, SCOPES)
                self.creds = flow.run_local_server(port=8000)
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

    def file_upload(self):
        filename = QFileDialog.getOpenFileName(self, 'Open file', './')

        if filename[0]:
            if not self.is_valid_xlsx(filename[0]):
                return
            contact_excel = openpyxl.load_workbook(filename=filename[0])
            contact_sheet = contact_excel['Sheet1']

    def parse_excel_sheet(sheet, header=True):
        self.inviations = []
        for i, row in enumerate(sheet.iter_rows()):
            if header == True:
                if i == 0:
                    continue
            self.invitations.append(Inviation(row))

    def is_valid_xlsx(self, filename):
        if not filename.endswith('.xlsx'):
            return False
        return True

    def list_mails(self):
        return

    def send_mails(self, invi, service, user_id='me'):
        if invi.is_eng():
            template = os.path.dirname(os.path.realpath(__file__)) + '/data/eng.json'
        else:
            template = os.path.dirname(os.path.realpath(__file__)) + '/data/kor.json'
        val = {
            'name': invi.name
            'sender': invi.sender
            'field': invi.get_field
            'date': invi.date
            'one_sen': invi.one_sen
        }


        return
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

        # Action Configuration
        exit_icon = 'exit.png'
        exitAction = QAction(QIcon(exit_icon), 'Exit', self)
        exitAction.setShortcut('Ctrl+Q')
        exitAction.setStatusTip('Exit Application')
        exitAction.triggered.connect(qApp.quit)

        # MenuBar Configuration


        # Windows Configuration
        icon_img = ""
        self.setWindowTitle("ICISTS Mail Management")
        # self.setWindowIcon(QIcon(icon_img))
        self.resize(500, 300)
        self.show()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    ex = MainApp()
    sys.exit(app.exec_())
