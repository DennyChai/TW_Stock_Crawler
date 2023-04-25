from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
from googleapiclient.discovery import build
import os


class GoogleSheet(object):
    def __init__(self, send_msg) -> None:
        self.send_msg = send_msg

    def load_credentials(self):
        service = None
        try:
            cred_path = r"D:\SideProject\Stock\projects\cred.json"
            token_path = r"D:\SideProject\Stock\projects\token.json"
            scopes = ["https://www.googleapis.com/auth/spreadsheets"]
            credentials = None
            if os.path.exists(token_path):
                credentials = Credentials.from_authorized_user_file(token_path, scopes)
            if not credentials or not credentials.valid:
                if credentials and credentials.expired and credentials.refresh_token:
                    credentials.refresh(Request())
                else:
                    flow = InstalledAppFlow.from_client_secrets_file(cred_path, scopes)
                    credentials = flow.run_local_server(port=0)
                with open(token_path, "w") as token:
                    token.write(credentials.to_json())
            service = build("sheets", "v4", credentials=credentials)
        except:
            self.send_msg()
        return service

    def get_sheet(self):
        sheet = None
        try:
            service = self.load_credentials()
            if service is not None:
                sheet = service.spreadsheets()
        except:
            self.send_msg()
        return sheet
