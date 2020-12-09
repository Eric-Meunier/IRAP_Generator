import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import sys
import calendar
import re
import xlwings as xw
import numpy as np
from mailmerge import MailMerge
from datetime import datetime
from PyQt5 import (QtGui, QtCore, uic)
from PyQt5.QtWidgets import (QMainWindow, QApplication, QTableWidgetItem, QInputDialog, QErrorMessage, QFileDialog,
                             QLineEdit)

# Modify the paths for when the script is being run in a frozen state (i.e. as an EXE)
if getattr(sys, 'frozen', False):
    application_path = sys.executable
    generator_ui_file = 'qt\\report_generator.ui'
    icons_path = 'qt\\icons'
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
    generator_ui_file = os.path.join(application_path, 'qt\\report_generator.ui')
    icons_path = os.path.join(application_path, "qt\\icons")

print(f"Application path: {application_path}")
# Load Qt ui file into a class
generator_ui, _ = uic.loadUiType(generator_ui_file)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# here enter the id of your google sheet
# SAMPLE_SPREADSHEET_ID_input = open('..//sheet.id', 'r').read()
SAMPLE_RANGE_NAME = 'A3:Q368'  # 1 year of rows


def get_sheet_df(sheet_id):
    # global values_input, service
    print("Retrieving sheet data")
    creds = None
    if os.path.exists(os.path.join(application_path, 'token.pickle')):
        with open(os.path.join(application_path, 'token.pickle'), 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '..//credentials.json', SCOPES)  # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open(os.path.join(application_path, 'token.pickle'), 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=sheet_id,
                                      range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    if not values_input:
        print('No data found.')

    df = pd.DataFrame(values_input[1:], columns=values_input[0])
    # Format the date column
    df.Date = df.Date.map(lambda x: datetime.strptime(x, r'%a, %b %d %Y'))
    return df


class ReportGenerator(QMainWindow, generator_ui):

    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle("IRAP Report Generator")
        self.setWindowIcon(QtGui.QIcon(os.path.join(icons_path, 'icon.png')))
        self.err_msg = QErrorMessage()
        self.table.setColumnWidth(0, 200)

        self.df = pd.DataFrame()
        self.employee = open(os.path.join(application_path, "..//employee.txt"), "r").read()
        self.sheet_id = open(os.path.join(application_path, '..//sheet.id'), 'r').read()
        self.total_hours = None

        # Add every year from 2020 to the current year as drop downs
        for year in range(2020, datetime.today().year + 1):
            self.year_cbox.addItem(str(year))
        self.year_cbox.setCurrentText('2020')
        self.month_cbox.setCurrentText('December')

        # Signals
        self.year_cbox.currentIndexChanged.connect(self.update_table)
        self.month_cbox.currentIndexChanged.connect(self.update_table)
        self.update_table_btn.clicked.connect(self.update_data)
        self.update_table_btn.clicked.connect(self.update_table)
        self.actionEmployee_Name.triggered.connect(self.update_employee_name)
        self.actionGoogle_Sheets_ID.triggered.connect(self.update_sheets_id)
        self.generate_files_btn.clicked.connect(self.generate_files)

        self.update_data()
        self.update_table()

    def update_employee_name(self):
        name, ok_pressed = QInputDialog.getText(self, 'Employee Name', 'Enter Employee Name:',
                                                QLineEdit.Normal,
                                                self.employee)

        if ok_pressed:
            id_file = open(os.path.join(application_path, "..//employee.txt"), "w+")
            id_file.write(name)
            print(f"New employee name: {open(os.path.join(application_path, '..//employee.txt'), 'r').read()}")
            id_file.close()

            self.employee = name
            self.statusBar().showMessage(f"Employee name changed to {name}.", 1000)

    def update_sheets_id(self):
        sheet_id, ok_pressed = QInputDialog.getText(self, 'Google Sheets ID', 'Enter Google Sheets ID:',
                                                    QLineEdit.Normal,
                                                    self.sheet_id)

        if ok_pressed:
            id_file = open(os.path.join(application_path, "..//sheet.id"), "w+")
            id_file.write(sheet_id)
            print(f"New sheet ID: {open(os.path.join(application_path, '..//sheet.id'), 'r').read()}")
            id_file.close()

            self.sheet_id = sheet_id
            self.statusBar().showMessage(f"Sheet ID changed to {sheet_id}.", 1000)

    def update_data(self):
        print("Updating sheet data")
        try:
            self.df = get_sheet_df(self.sheet_id)
        except Exception:
            self.err_msg.showMessage(f"Error retrieving Google Sheets data. "
                                     f"Please make sure the Google Sheets ID is correct.")
            self.df = pd.DataFrame()

    def format_df(self, df, month, year):
        # Remove entries with no comments
        df = df[df['Comments'].astype(bool)]
        # Filter df to only include IRAP
        df = df[df['Comments'].str.contains('IRAP')].reset_index(drop=True)
        # Only keep relevant columns
        df = df.loc[:, ['Date', ' Statutory Holiday', 'Comments']]
        # Filter df to only include selected month and year
        df = df[df.Date.map(lambda x: x.month == month and x.year == year)]
        return df

    def update_table(self):
        print("Updating table")
        self.table.clearContents()
        if self.df.empty:
            print(f"DataFrame is empty")
            return

        # self.table.setRowCount(0)

        month = self.month_cbox.currentIndex() + 1
        # month = 11
        year = int(self.year_cbox.currentText())

        # Add each day of the month as a row
        dates = []
        num_days = calendar.monthrange(year, month)[1]
        self.table.setRowCount(num_days)
        for day in range(1, num_days + 1):
            date = datetime(year, month, day)
            dates.append(date)
            date_str = date.strftime(r'%B %d, %Y (%A)')
            item = QTableWidgetItem(date_str)
            item.setFlags(item.flags() ^ QtCore.Qt.ItemIsEditable)
            self.table.setItem(day - 1, 0, item)

        df = self.format_df(self.df, month, year)
        if not df.empty:
            # Add the description and IRAP hours
            for ind, row in df.iterrows():
                # Split by periods into lists.
                comments = row.Comments.split('\n')

                for comment in comments:
                    research_comment = re.search(r'research:(.*)[({[](.*)[)\]}]\.', comment, flags=re.I)
                    if research_comment:
                        if 'irap' in research_comment.group(0).lower():
                            comment = research_comment.group(1).strip()
                            comment = re.sub(r'\(irap\) ', '', comment, flags=re.I)  # Remove (IRAP)
                            hours = research_comment.group(2).strip()

                            hours_item = QTableWidgetItem(hours)
                            hours_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
                            self.table.setItem(row.Date.day - 1, 3, QTableWidgetItem(f"{comment}."))
                            self.table.setItem(row.Date.day - 1, 1, hours_item)
                        # continue

        self.table.resizeRowsToContents()
        self.statusBar().showMessage(f"Table updated.", 1000)

    def generate_files(self):

        def table_to_dataframe():

            def get_hours(row):
                # Write SAT or SUN if the day is a weekend
                if row.Date.weekday() == 5:
                    hours = 'SAT'
                elif row.Date.weekday() == 6:
                    hours = 'SUN'
                elif row.Holiday is True:
                    print(row.Holiday)
                    hours = 'Holiday'
                else:
                    hours = row.Hours

                return hours

            number_of_rows = self.table.rowCount()
            number_of_columns = self.table.columnCount()

            table_df = pd.DataFrame(
                columns=['Date', 'Hours', 'Task_number', 'Description'],  # Fill columnets
                index=range(number_of_rows)  # Fill rows
            )

            for i in range(number_of_rows):
                for j in range(number_of_columns):
                    item = self.table.item(i, j)
                    if item:
                        table_df.iloc[i, j] = item.text()
                    else:
                        table_df.iloc[i, j] = None

            table_df.Date = table_df.Date.map(lambda x: datetime.strptime(x, r'%B %d, %Y (%A)'))

            # Merge the holidays column
            data_df = self.df[self.df.Date.map(lambda x: x.month == self.month_cbox.currentIndex() + 1 and
                                                         x.year == int(self.year_cbox.currentText()))]

            # Merge the data frames
            df = table_df.merge(data_df.loc[:, ['Date', ' Statutory Holiday']], how='outer', on='Date')
            df.rename(columns={' Statutory Holiday': 'Holiday'}, inplace=True)
            df.Holiday.replace(np.nan, False, inplace=True)
            df.Holiday = df.Holiday.astype(bool)

            self.total_hours = df.Hours.astype(float).sum()
            df.Hours = df.apply(get_hours, axis=1)
            return df

        def save_timesheet():
            """Create and save the Excel time sheet"""
            print("Creating Time Sheet")
            cells_dict = {'1': 'D18', '2': 'D19', '3': 'D20', '4': 'D21', '5': 'D22', '6': 'D23', '7': 'D24',
                          '8': 'D25', '9': 'G18', '10': 'G19', '11': 'G20', '12': 'G21', '13': 'G22', '14': 'G23',
                          '15': 'G24', '16': 'G25', '17': 'J18', '18': 'J19', '19': 'J20', '20': 'J21', '21': 'J22',
                          '22': 'J23', '23': 'J24', '24': 'J25', '25': 'M18', '26': 'M19', '27': 'M20', '28': 'M21',
                          '29': 'M22', '30': 'M23', '31': 'M24'}

            excel_app = xw.App(visible=False)
            excel_file = excel_app.books.open(os.path.join(application_path, r'../timesheet_template.xlsx'))
            sheet = excel_file.sheets('Sheet1')

            # Fill the hours
            for row in data.itertuples():
                cell = sheet.range(cells_dict[str(row.Date.day)])
                # data_row = self.df.loc[self.df.Date == row.Date]

                cell.value = row.Hours

            # Add the hyphen for days shorter than 31
            if len(data) < 31:
                missing_days = 31 - len(data)
                for i in range(missing_days):
                    cell = sheet.range(cells_dict[str(len(data) + (i + 1))])
                    cell.value = '-'

            # Add the employee name
            sheet.range('E10').value = self.employee
            # Add the month and year
            sheet.range('K10').value = month
            sheet.range('N10').value = year

            # Add the total number of hours worked
            # Filter the month
            month_filt = self.df.Date[self.df.Date.map(lambda x: x.month == self.month_cbox.currentIndex() + 1)]
            # Get weekdays
            weekday_filt = month_filt[~month_filt.map(lambda x: x.weekday() in [5, 6])]
            num_weekdays = len(weekday_filt)
            sheet.range('I28').value = num_weekdays * 7.5

            # excel_file.save(f"{folder}\\{month} {year} IRAP Time Sheet.xlsx")
            excel_file.save(f"{month} {year} IRAP Time Sheet.xlsx")
            excel_file.close()
            print(f"Time Sheet save successful.")

            os.startfile(f"{month} {year} IRAP Time Sheet.xlsx")

        def save_worklog():
            """Create and save the worklog"""

            def row_to_dict(row):
                d = {
                    'Date': str(row.Date.day),
                    'Hours': str(row.Hours),
                    'Description': row.Description,
                    'Task': ''
                }
                return d

            template = os.path.join(application_path, r'../worklog_template.docx')

            document = MailMerge(template)

            # Fill the header
            document.merge(
                Name=self.employee,
                Year=year,
                Month=month,
                Total_hours=str(self.total_hours),
            )

            # Fill the table
            data.Hours = data.Hours.replace(np.nan, 0)
            table_dict = data.replace(np.nan, '').apply(row_to_dict, axis=1)
            document.merge_rows('Date', table_dict)

            document.write(f"{month} {year} IRAP Worklog.docx")
            os.startfile(f"{month} {year} IRAP Worklog.docx")
            print(f"Worklog save successful.")

        print(f"Generating files")

        month = self.month_cbox.currentText()
        year = self.year_cbox.currentText()

        folder = QFileDialog.getExistingDirectory(self, "Selected Output Folder")
        if not folder:
            return

        global data
        data = table_to_dataframe()

        try:
            save_timesheet()
        except Exception as e:
            self.err_msg.showMessage(f"Error occurred creating the time sheet: {e}.")
            return

        try:
            save_worklog()
        except Exception as e:
            self.err_msg.showMessage(f"Error occurred creating the work log: {e}.")
            return

        self.statusBar().showMessage(f"Save complete. Files saved to {folder}.", 2000)


if __name__ == '__main__':
    app = QApplication(sys.argv)

    lc = ReportGenerator()
    lc.show()

    # lc.generate_files()

    app.exec_()
