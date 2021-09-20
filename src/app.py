import streamlit as st
import pandas as pd
import numpy as np
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
import re
import base64
import xlwings as xw
from mailmerge import MailMerge
from datetime import datetime
import shutil
import pythoncom


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


def update_data():

    def format_df(df):
        """
        Remove irrelevant columns, filter only IRAP comments, and add a IRAP hours column
        :param df: DataFrame of timesheet
        :return: DataFrame
        """
        def get_irap_info(row):
            """"Filter a data frame to only include IRAP information"""
            irap_row = {'Date': row.Date,
                        'Hours': '',
                        'Holiday': row.Holiday,
                        'Comments': ''}

            # Split by new line into lists.
            comments = row.Comments.split('\n')

            for comment in comments:
                research_comment = re.search(r'research:(.*)[({[](.*)[)\]}]\.', comment, flags=re.I)
                if research_comment:
                    if 'irap' in research_comment.group(0).lower():
                        comment = research_comment.group(1).strip()
                        comment = re.sub(r'\(irap\)', '', comment, flags=re.I)  # Remove (IRAP)
                        hours = research_comment.group(2).strip()

                        irap_row['Comments'] = f"{comment.strip()}."
                        irap_row['Hours'] = hours

            return irap_row

        def get_hours(row):
            """
            Replaces weekends and holidays with strings
            :param row: DataFrame row
            :return: value to be used in the generated files for the Hours column
            """
            # Write SAT or SUN if the day is a weekend
            if row.Date.weekday() == 5:
                hours = 'SAT'
            elif row.Date.weekday() == 6:
                hours = 'SUN'
            elif row.Holiday is True:
                hours = 'Holiday'
            else:
                hours = row.Hours

            return hours

        # Only keep relevant columns
        df = df.loc[:, ['Date', ' Statutory Holiday', 'Comments']].copy()

        # Make the holidays a bool
        df.rename(columns={' Statutory Holiday': 'Holiday'}, inplace=True)
        df.Holiday = df.Holiday.astype(bool)

        # Filter df to only include selected month and year
        global month_index
        month_index = months.index(month)
        df = df[df.Date.map(lambda x: x.month == month_index + 1 and x.year == int(year))]
        if df.empty:
            return df, 0
        df.insert(1, 'Hours', '')
        df.fillna('', inplace=True)

        # Add the description and IRAP hours
        df = pd.DataFrame(df.apply(get_irap_info, axis=1).to_numpy())
        total_hours = df.Hours.replace('', 0).astype(float).sum()
        df.Hours = df.apply(get_hours, axis=1)

        return df, total_hours

    try:
        sheet_df = get_sheet_df(sheet_id)
    except Exception as e:
        st.write(f"Error retrieving Google Sheets data. "
                 f"Please make sure the Google Sheets ID is correct.")
        st.write(f"{e}")
    else:
        irap_df, total_hours = format_df(sheet_df)
        return irap_df, total_hours


def draw_table(data):
    if data.empty:
        st.write(f"No data found.")
    else:
        # Create the table
        table_df = data.loc[:, ['Date', 'Hours', 'Comments']].copy()
        table_df.set_index('Date').sort_index(ascending=True)
        st.table(table_df)


def generate_files():

    def get_binary_file_downloader(bin_file, file_label='File'):
        with open(bin_file, 'rb') as f:
            data = f.read()
        bin_str = base64.b64encode(data).decode()
        href = f'<a href="data:application/octet-stream;base64,{bin_str}" download="{os.path.basename(bin_file)}">Download {file_label}</a>'
        return href

    def save_timesheet():
        """Create and save the Excel time sheet"""
        print("Creating Time Sheet")
        cells_dict = {'1': 'D18', '2': 'D19', '3': 'D20', '4': 'D21', '5': 'D22', '6': 'D23', '7': 'D24',
                      '8': 'D25', '9': 'G18', '10': 'G19', '11': 'G20', '12': 'G21', '13': 'G22', '14': 'G23',
                      '15': 'G24', '16': 'G25', '17': 'J18', '18': 'J19', '19': 'J20', '20': 'J21', '21': 'J22',
                      '22': 'J23', '23': 'J24', '24': 'J25', '25': 'M18', '26': 'M19', '27': 'M20', '28': 'M21',
                      '29': 'M22', '30': 'M23', '31': 'M24'}

        # Create a copy of the template and place it in the same dir as app.py,
        #  or else there are "Member not found" errors
        file_name = f"{name} {month} {year} IRAP Time Sheet.xlsx"
        shutil.copy(r'../templates/timesheet_template.xlsx', file_name)
        
        # Create the excel object
        pythoncom.CoInitialize()  # Required to avoid "CoInitialize has not been called" error.
        excel_app = xw.App(visible=False)  # This prevents the file from opening
        excel_file = excel_app.books.open(file_name)
        sheet = excel_file.sheets('Sheet1')

        # Fill the hours
        for row in irap_df.itertuples():
            cell = sheet.range(cells_dict[str(row.Date.day)])
            # data_row = df.loc[df.Date == row.Date]
            cell.value = row.Hours

        # Add the hyphen for days shorter than 31
        if len(irap_df) < 31:
            missing_days = 31 - len(irap_df)
            for i in range(missing_days):
                cell = sheet.range(cells_dict[str(len(irap_df) + (i + 1))])
                cell.value = '-'

        # Add the employee name
        sheet.range('E10').value = name
        # Add the month and year
        sheet.range('K10').value = month
        sheet.range('N10').value = year

        # Add the total number of hours worked
        # Filter the month
        month_filt = irap_df.Date[irap_df.Date.map(lambda x: x.month == month_index + 1)]
        # Get weekdays
        weekday_filt = month_filt[~month_filt.map(lambda x: x.weekday() in [5, 6])]
        num_weekdays = len(weekday_filt)
        sheet.range('I28').value = num_weekdays * 7.5

        excel_file.save()
        excel_file.close()
        print(f"Time Sheet save successful.")
        st.markdown(get_binary_file_downloader(file_name, file_name), unsafe_allow_html=True)

    def save_worklog():
        """Create and save the worklog"""

        def row_to_dict(row):
            d = {
                'Date': str(row.Date.day),
                'Hours': str(row.Hours),
                'Description': row.Comments,
                'Task': ''
            }
            return d

        template = r'../templates/worklog_template.docx'

        document = MailMerge(template)

        # Fill the header
        document.merge(
            Name=name,
            Year=year,
            Month=month,
            Total_hours=str(total_hours),
        )

        # Fill the table
        # print(f"IRAP df hours: {irap_df.Hours}")
        irap_df.Hours = irap_df.Hours.replace(np.nan, 0)
        table_dict = irap_df.replace(np.nan, '').apply(row_to_dict, axis=1)
        document.merge_rows('Date', table_dict)

        file_name = f"{name} {month} {year} IRAP Worklog.docx"
        document.write(file_name)
        st.markdown(get_binary_file_downloader(file_name, file_name), unsafe_allow_html=True)
        document.close()
        print(f"Worklog save successful.")

    print(f"Generating files")
    irap_df, total_hours = update_data()

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


# Add the title
title = st.title(f"IRAP Time Sheet & Worklog Generator")

# Text inputs for timesheet ID and employee name
name = st.sidebar.selectbox("Name", ['Eric Meunier', 'Dave Campbell'], index=0)
default_sheet_id = '183TvCEIn3R9rsqCuseDtcGVtUVIPxO8a_fCg0iHlAhY' if \
    name == 'Eric Meunier' else '1WBizXmsAGYNiLdVf-znRxX2JH3Z2DijETQd7ceaYVo0'
sheet_id = st.sidebar.text_input("Timesheet ID", default_sheet_id)

# Add dropdown options for year
years = []
for year in range(2020, datetime.today().year + 1):
    years.append(str(year))
year = st.sidebar.selectbox('Year', years)

# Add dropdown options for month
months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
          'November', 'December']
last_month = datetime.today().month - 1
if last_month == 0:
    last_month = 12
month = st.sidebar.selectbox('Month', months, index=last_month - 1)

# Button that creates/updates the table
if st.button('Update Data'):
    irap_df, total_hours = update_data()
    draw_table(irap_df)
    st.write(f"Total IRAP hours: {total_hours}")

if st.button('Generate Files'):
    generate_files()

# update_data()
# generate_files()


# if __name__ == '__main__':
#     fg = FileGenerator()
#
#     sheet_id = '183TvCEIn3R9rsqCuseDtcGVtUVIPxO8a_fCg0iHlAhY'
#     year = 2020
#     month = 'November'
#     sheet_df = get_sheet_df(sheet_id)



