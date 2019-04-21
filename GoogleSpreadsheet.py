import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


class GoogleSpreadsheet:
  def __init__(self, spreadsheet_id, range_name, scopes):
    self.scopes = scopes
    self.spreadsheet_id = spreadsheet_id
    self.range_name = range_name

    # self.service = self.authorize()

    # # If modifying these scopes, delete the file token.pickle.
    # SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

    # # The ID and range of a sample spreadsheet.
    # SAMPLE_SPREADSHEET_ID = '1thT4M0AFcBucct1PLk7vL29Fhrxz1YnxOAq-06d64V8'
    # SAMPLE_RANGE_NAME = 'April 19'

  def authorize(self):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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
            'credentials.json', self.scopes)
        creds = flow.run_local_server()
      # Save the credentials for the next run
      with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

    self.service = build('sheets', 'v4', credentials=creds)

    return True

  def loadSpreadsheet(self):
    return self.service.spreadsheets()

  def loadValues(self):
    sheet = self.loadSpreadsheet()
    result = sheet.values().get(spreadsheetId=self.spreadsheet_id, range=self.range_name).execute()
    values = result.get('values', [])

    return values
