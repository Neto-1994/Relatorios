import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.errors import HttpError

def validar():
    try:
        # Parte de log da API Google Sheets
        # If modifying these scopes, delete the file token.json.
        SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("E:/Relatorios/token.json"):
            creds = Credentials.from_authorized_user_file(
                "E:/Relatorios/token.json", SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    "E:/Relatorios/CredenciaisJair.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("E:/Relatorios/token.json", "w") as token:
                token.write(creds.to_json())
        return creds

    except HttpError as err:
        print(err)
