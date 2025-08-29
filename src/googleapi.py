import os
from typing import List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from helpers import setup_logger

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

TOKEN_CACHE_FILE = "token.json"
CREDENTIALS_FILE = "google-credentials.json"

logger = setup_logger("google")

__all__ = [
    "get_google_sheet",
    "get_sheet_titles",
    "get_title",
]


def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(TOKEN_CACHE_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_CACHE_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_CACHE_FILE, "w") as token:
            token.write(creds.to_json())
    return creds


DOCUMENT_CACHE = {}


def get_google_sheet(sheet_id: str, range_name: str):
    if (sheet_id, range_name) not in DOCUMENT_CACHE:
        logger.info(
            f"Getting values from sheet '{get_title(sheet_id)}' range '{range_name}'"
        )
        try:
            creds = get_credentials()
            service = build("sheets", "v4", credentials=creds)
            sheet = service.spreadsheets()
            result = (
                sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
            )
            DOCUMENT_CACHE[(sheet_id, range_name)] = result.get("values", [])
        except HttpError as err:
            print(err)
            return []

    return DOCUMENT_CACHE[(sheet_id, range_name)]


METADATA_CACHE = {}


def get_metadata(sheet_id: str) -> dict:
    if sheet_id not in METADATA_CACHE:
        try:
            creds = get_credentials()
            service = build("sheets", "v4", credentials=creds)
            sheet_metadata = (
                service.spreadsheets().get(spreadsheetId=sheet_id).execute()
            )
            METADATA_CACHE[sheet_id] = sheet_metadata
        except HttpError as err:
            print(err)
            return {}
    return METADATA_CACHE[sheet_id]


def get_title(sheet_id: str) -> str:
    metadata = get_metadata(sheet_id)
    return metadata.get("properties", {}).get("title", "Untitled")


def get_sheet_titles(sheet_id: str) -> List[str]:
    metadata = get_metadata(sheet_id)
    sheets = metadata.get("sheets", "")
    titles = [sheet.get("properties", {}).get("title", "Sheet1") for sheet in sheets]
    return titles
