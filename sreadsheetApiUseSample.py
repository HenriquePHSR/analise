# DONT RUN THIS SCRIPT
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# it maybe required give share permissions for app's service email 
# OAuthCreds should be adquired from google developers console

def Login():
    # Login into google for api's use authorization and return a api backend instance
    creds = None
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets'] # permissions
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                # OAuthCreds should be adquired from google developers console
                'OAuthCreds.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
        
    # Call the Sheets API
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()

    return sheet

def get(SHEET_API_INSTANCE, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME):
    # get values from a sreadsheet
    result = SHEET_API_INSTANCE.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    # else:
    #     print('Data:')
    #     for row in values:
    #         print(row)
    return result

def create(SHEET_API_INSTANCE, TITLE):
    # create a sreadsheet and returns its id
    spreadsheet = {
        'properties': {
            'title': TITLE
        }
    }

    spreadsheet = SHEET_API_INSTANCE.create(body=spreadsheet,
                                        fields='spreadsheetId').execute()
    spreadsheetID = spreadsheet.get('spreadsheetId')
    print('Spreadsheet ID: {0}'.format(spreadsheetID))
    return spreadsheetID

def get_batchget(SHEET_API_INSTANCE, spreadsheet_id, range_names):
    result = SHEET_API_INSTANCE.values().batchGet(
        spreadsheetId=spreadsheet_id, ranges=range_names).execute()
    ranges = result.get('valueRanges', [])
    print('{0} ranges retrieved.'.format(len(ranges)))

    return result

def clear(SHEET_API_INSTANCE, spreadsheet_id, range):
    # The A1 notation of the values to clear.
    clear_values_request_body = {
    }

    result = SHEET_API_INSTANCE.values().clear(spreadsheetId=spreadsheet_id, range=range, body=clear_values_request_body).execute()
    return result

def update(SHEET_API_INSTANCE, spreadsheet_id, value_input_option, range_name, body):
    # update cells in a continuos range and return cells changed
    result = SHEET_API_INSTANCE.values().update(
        spreadsheetId=spreadsheet_id,
        valueInputOption=value_input_option, range=range_name, body=body).execute()
    op_info = result.get('updatedCells')
    print('{0} cells updated.'.format(op_info))
    return result

def update_batchupdate(SHEET_API_INSTANCE, spreadsheet_id, body):
    result = SHEET_API_INSTANCE.values().batchUpdate(
        spreadsheetId=spreadsheet_id, body=body).execute()
    print('{0} cells updated.'.format(result.get('totalUpdatedCells')))
    return result

sheet_api_instance = Login()
# ==============================================
# The ID and range of a sample spreadsheet.
# A1 notation
# Sheet1
# Sheet1!A2:X2
SAMPLE_SPREADSHEET_ID = '1D9Wdt-ZjHUBk6lTdwjYSBnmqn2kpIMMNKQK0aiEaRLg'
SAMPLE_RANGE_NAME = 'Sheet1!A2:X4'
values = get(sheet_api_instance, SAMPLE_SPREADSHEET_ID, SAMPLE_RANGE_NAME)
df = pd.DataFrame(values.get('values', []))
df
# ==============================================
TITLE = 'automatedCreateSpreadsheet2'
spreadsheetID = create(sheet_api_instance, TITLE)
# ==============================================
range_names = [
    'Sheet1!A2:X4',
    'Sheet1!A10:X11'
    ]

values_range = get_batchget(sheet_api_instance, SAMPLE_SPREADSHEET_ID, range_names)
# ==============================================
values = [
    [
        444444, 444444
    ],
    [
        555555, 555555
    ]
]
range_name = 'Página1'
body = {
    'range' : range_name,
    'values': values
}
SAMPLE_SPREADSHEET_ID = '1D9Wdt-ZjHUBk6lTdwjYSBnmqn2kpIMMNKQK0aiEaRLg'
value_input_option = 'RAW'
update_info = update(sheet_api_instance, SAMPLE_SPREADSHEET_ID, value_input_option, range_name, body)
# ==============================================
values = [
    [
        12345, 12345 # cells values
    ],
    # Additional rows
]
values2 = [
    [
        6789, 6789
    ],
]
data = [
    { # a range object
        'range': 'Sheet1!B2:C2',
        'values': values
    }, 
    {
        'range': 'Página1',
        'values': values2
    }
    # Additional ranges to update
]
body = {
    'valueInputOption': 'RAW',
    'data': data
}
result = update_batchupdate(sheet_api_instance, SAMPLE_SPREADSHEET_ID, body)
# ==============================================
values = [
    [
        
    ]
]
range_name = 'Página1'
body = {
    'range' : range_name,
    'values': values
}
SAMPLE_SPREADSHEET_ID = '1D9Wdt-ZjHUBk6lTdwjYSBnmqn2kpIMMNKQK0aiEaRLg'
value_input_option = 'RAW'
update_info = update(sheet_api_instance, SAMPLE_SPREADSHEET_ID, value_input_option, range_name, body)
# ==============================================
