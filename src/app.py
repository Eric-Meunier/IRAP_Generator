import streamlit as st
import pandas as pd
import numpy as np
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


def get_sheet_df(sheet_id):
    # global values_input, service
    print("Retrieving sheet data")

    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    sample_range_name = 'A3:Q368'  # 1 year of rows

    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '..//credentials.json', scopes)  # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds, cache_discovery=False)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=sheet_id,
                                      range=sample_range_name).execute()
    values_input = result_input.get('values', [])

    if not values_input:
        print('No data found.')

    df = pd.DataFrame(values_input[1:], columns=values_input[0])
    # Format the date column
    df.Date = df.Date.map(lambda x: datetime.strptime(x, r'%a, %b %d %Y'))
    return df


def format_df(df, month, year):
    # # Remove entries with no comments
    # df = df[df['Comments'].astype(bool)]

    # Filter df to only include IRAP
    # df = df[df['Comments'].str.contains('IRAP')].reset_index(drop=True)
    # Only keep relevant columns
    df = df.loc[:, ['Date', ' Statutory Holiday', 'Comments']].copy()
    # Filter df to only include selected month and year
    global month_index
    month_index = months.index(month)
    df = df[df.Date.map(lambda x: x.month == month_index + 1 and x.year == int(year))]
    df.set_index('Date', inplace=True)
    return df


def update_table():
    try:
        sheet_df = get_sheet_df(sheet_id)
    except Exception as e:

        st.write(f"Error retrieving Google Sheets data. "
                 f"Please make sure the Google Sheets ID is correct.")
        st.write(f"{e}")
    else:
        global df
        df = format_df(sheet_df, month, year)
        st.write(df)


def generate_files():

        def save_timesheet():
            """Create and save the Excel time sheet"""
            print("Creating Time Sheet")
            cells_dict = {'1': 'D18', '2': 'D19', '3': 'D20', '4': 'D21', '5': 'D22', '6': 'D23', '7': 'D24',
                          '8': 'D25', '9': 'G18', '10': 'G19', '11': 'G20', '12': 'G21', '13': 'G22', '14': 'G23',
                          '15': 'G24', '16': 'G25', '17': 'J18', '18': 'J19', '19': 'J20', '20': 'J21', '21': 'J22',
                          '22': 'J23', '23': 'J24', '24': 'J25', '25': 'M18', '26': 'M19', '27': 'M20', '28': 'M21',
                          '29': 'M22', '30': 'M23', '31': 'M24'}

            excel_app = xw.App(visible=False)
            excel_file = excel_app.books.open(r'../timesheet_template.xlsx')
            sheet = excel_file.sheets('Sheet1')

            # Fill the hours
            for row in df.itertuples():
                cell = sheet.range(cells_dict[str(row.Date.day)])
                # data_row = self.df.loc[self.df.Date == row.Date]

                cell.value = row.Hours

            # Add the hyphen for days shorter than 31
            if len(df) < 31:
                missing_days = 31 - len(df)
                for i in range(missing_days):
                    cell = sheet.range(cells_dict[str(len(df) + (i + 1))])
                    cell.value = '-'

            # Add the employee name
            sheet.range('E10').value = name
            # Add the month and year
            sheet.range('K10').value = month
            sheet.range('N10').value = year

            # Add the total number of hours worked
            # Filter the month
            month_filt = df.Date[df.Date.map(lambda x: x.month == month_index + 1)]
            # Get weekdays
            weekday_filt = month_filt[~month_filt.map(lambda x: x.weekday() in [5, 6])]
            num_weekdays = len(weekday_filt)
            sheet.range('I28').value = num_weekdays * 7.5

            # excel_file.save(f"{folder}\\{month} {year} IRAP Time Sheet.xlsx")
            excel_file.save(f"{month} {year} IRAP Time Sheet.xlsx")
            excel_file.close()
            print(f"Time Sheet save successful.")

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

            template = r'../worklog_template.docx'

            document = MailMerge(template)

            # Fill the header
            document.merge(
                Name=name,
                Year=year,
                Month=month,
                Total_hours=str(total_hours),
            )

            # Fill the table
            df.Hours = df.Hours.replace(np.nan, 0)
            table_dict = df.replace(np.nan, '').apply(row_to_dict, axis=1)
            document.merge_rows('Date', table_dict)

            document.write(f"{month} {year} IRAP Worklog.docx")
            document.close()
            print(f"Worklog save successful.")

        print(f"Generating files")

        try:
            save_timesheet()
        except Exception as e:
            st.write(f"Error occurred creating the time sheet: {e}.")
            return

        try:
            save_worklog()
        except Exception as e:
            st.write(f"Error occurred creating the work log: {e}.")
            return

        st.write(f"Save complete.")


# Add the title
title = st.title(f"IRAP Time Sheet & Worklog Generator")
# df = get_sheet_df()

# Text inputs for timesheet ID and employee name
sheet_id = st.sidebar.text_input("Timesheet ID", '183TvCEIn3R9rsqCuseDtcGVtUVIPxO8a_fCg0iHlAhY')
name = st.sidebar.text_input("Name", 'Eric Meunier')

# Add dropdown options for year
years = []
for year in range(2020, datetime.today().year + 1):
    years.append(str(year))
year = st.sidebar.selectbox('Year', years)

# Add dropdown options for month
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November',
          'December']
last_month = datetime.today().month - 1
print(F"Last month: {last_month}")
if last_month == 0:
    last_month = 12
month = st.sidebar.selectbox('Month', months, index=last_month - 1)

# Button that creates/updates the table
if st.button('Update'):
    update_table()

if st.button('Generate Files'):
    generate_files()

update_table()

if __name__ == '__main__':
    sheet_id = '183TvCEIn3R9rsqCuseDtcGVtUVIPxO8a_fCg0iHlAhY'
    year = 2020
    month = 'November'
    sheet_df = get_sheet_df(sheet_id)



