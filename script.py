import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle
import numpy as np
import math

# The ID and range of a sample spreadsheet.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
MY_GOOGLE_SHEET_ID = '1DCZ5lmPCBybMiWXyG7tscTj0YJlmwOZQU1eT_9NlZmg'
SAMPLE_READ_RANGE = 'A1:H27'

#Write range is different since the first two rows are headers
SAMPLE_WRITE_RANGE = 'A3:H27'

#function to read data from a google spreadsheet with id MY_GOOGLE_SHEET_ID
#code written according to google's api (https://developers.google.com/sheets/api/quickstart/python)
def read_spreadsheet():
    global values_input, service
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=MY_GOOGLE_SHEET_ID,
                                range=SAMPLE_READ_RANGE).execute()
    values_input = result_input.get('values', [])

    if not values_input and not values_expansion:
        print('No data found.')

#function to calculate the minimum score necessary to get a passing grade (>=5)
#rounds up to the nearest integer
def calc_final_score(mean_score):
    return math.ceil(10 - mean_score)

#function to grade student based on their mean test score
def grade_student (student):
    absence = int(student[2])

    if (absence / total_classes >= 0.25):
        student[6] = "Reprovado por Falta"
        student[7] = 0
    else:
        mean_score = sum([int(x) for x in student[3:6]]) / 30
        if (mean_score >= 7):
            student[6] = "Aprovado"
            student[7] = 0
        elif (mean_score >= 5):
            student[6] = "Exame final"
            student[7] = calc_final_score(mean_score)
        else:
            student[6] = "Reprovado"
            student[7] = 0
    return student

def export_data_to_sheets(df):
    response_date = service.spreadsheets().values().update(
        spreadsheetId=MY_GOOGLE_SHEET_ID,
        valueInputOption='RAW',
        range=SAMPLE_WRITE_RANGE,
        body=dict(
            majorDimension='ROWS',
            values=df.T.reset_index().T.values.tolist())
        ).execute()
    print('Sheet successfully Updated')

def main():
    read_spreadsheet()

    #fill missing values in all cells with NaN
    #necessary to create pandas dataframe
    for values in values_input[3:]:
        if (len(values) == 6):
            values.extend([np.nan, np.nan])

    #create dataframe
    df=pd.DataFrame(values_input[3:], columns=values_input[2])

    #print(df)

    #extract the number of total classes in a semester
    global total_classes
    total_classes = int("".join(filter(str.isdigit, values_input[1][0])))

    #grade every student
    df = df.apply(grade_student, axis=1)

    #print(df)

    export_data_to_sheets(df)

main()
