import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle
import numpy as np
import math

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
MY_GOOGLE_SHEET_ID = '1DCZ5lmPCBybMiWXyG7tscTj0YJlmwOZQU1eT_9NlZmg'
MY_SHEET_CELLS_RANGE = 'A1:H27'

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID_input = MY_GOOGLE_SHEET_ID
SAMPLE_RANGE_NAME = MY_SHEET_CELLS_RANGE

def main():
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
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    if not values_input and not values_expansion:
        print('No data found.')

#function to calculate the minimum score necessary to get a passing grade (>=5)
#rounds up to the nearest integer
def calc_final_score(mean_score):
    return math.ceil(10 - mean_score)

#function to grade student based on their mean test score
#returns a new pandas dataframe
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

main()

#fill missing values in all cells with NaN
#necessary to create pandas dataframe
for values in values_input[3:]:
    values.extend([np.nan, np.nan])

#create dataframe
df=pd.DataFrame(values_input[3:], columns=values_input[2])

#extract the number of total classes in a semester
global total_classes
total_classes = int("".join(filter(str.isdigit, values_input[1][0])))

#grade every student
df = df.apply(grade_student, axis=1)

print(df)
