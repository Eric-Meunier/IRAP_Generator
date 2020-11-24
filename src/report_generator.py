import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow, Flow
from google.auth.transport.requests import Request
import os
import pickle
import sys
from PyQt5 import (QtGui, QtCore, uic)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QComboBox, QShortcut, QFileDialog)


# Modify the paths for when the script is being run in a frozen state (i.e. as an EXE)
if getattr(sys, 'frozen', False):
    application_path = sys.executable
    generator_ui_file = 'qt\\report_generator.ui'
    icons_path = 'qt\\icons'
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
    generator_ui_file = os.path.join(application_path, 'qt\\report_generator.ui')
    icons_path = os.path.join(application_path, "qt\\icons")

# Load Qt ui file into a class
generator_ui, _ = uic.loadUiType(generator_ui_file)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# here enter the id of your google sheet
SAMPLE_SPREADSHEET_ID_input = open('..//sheet.id', 'r').read()
SAMPLE_RANGE_NAME = 'A3:Q368'  # 1 year of rows


def get_sheet_df():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '..//credentials.json', SCOPES)  # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                      range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    if not values_input:
        print('No data found.')

    df = pd.DataFrame(values_input[1:], columns=values_input[0])
    return df


class ReportGenerator(QMainWindow, generator_ui):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.df_to_table()

    def df_to_table(self):

        def fill_row(row):
            pass

        df = get_sheet_df()
        self.table.setRowCount(len(df))
        df.apply(fill_row, axis=1)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    lc = ReportGenerator()
    lc.show()

    app.exec_()
