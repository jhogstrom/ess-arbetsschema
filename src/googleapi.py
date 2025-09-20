"""Google Sheets API utilities for reading spreadsheet data.

This module provides a simplified interface for accessing Google Sheets data
using the Google Sheets API v4. It handles authentication, caching, and
provides convenient methods for retrieving sheet data and metadata.

The module requires Google API credentials in 'google-credentials.json' and
will create a 'token.json' file for caching authentication tokens.

Example:
    >>> from googleapi import get_google_sheet, get_sheet_titles
    >>> sheet_id = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms"
    >>> titles = get_sheet_titles(sheet_id)
    >>> data = get_google_sheet(sheet_id, f"{titles[0]}!A1:Z100")
"""

import os
from typing import List

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from helpers import setup_logger

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/gmail.modify",
]

TOKEN_CACHE_FILE = "token.json"
CREDENTIALS_FILE = "google-credentials.json"

logger = setup_logger("google")

__all__ = [
    "get_google_sheet",
    "get_sheet_titles",
    "get_title",
]


def get_credentials():
    """Retrieve and manage Google API credentials for accessing Google Sheets.

    This function handles the OAuth2 flow for Google Sheets API access. It loads
    existing credentials from a token file, refreshes them if expired, or initiates
    a new authorization flow if needed.

    Returns:
        google.oauth2.credentials.Credentials: Valid Google API credentials

    Raises:
        FileNotFoundError: If google-credentials.json is missing
        google.auth.exceptions.RefreshError: If token refresh fails

    Example:
        >>> creds = get_credentials()
        >>> service = build("sheets", "v4", credentials=creds)
    """
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
    """Retrieve data from a specific range in a Google Sheet.

    This function fetches data from a Google Sheet using the Sheets API and caches
    the results to avoid repeated API calls for the same data.

    Args:
        sheet_id: The unique identifier of the Google Sheet
        range_name: The A1 notation range to retrieve (e.g., "Sheet1!A1:C10")

    Returns:
        List[List[str]]: A 2D list representing the sheet data, where each inner
        list is a row and each element is a cell value as a string

    Raises:
        googleapiclient.errors.HttpError: If the API request fails

    Example:
        >>> data = get_google_sheet("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms", "Class Data!A2:E")
        >>> print(data[0])  # First row of data
        ['Alexandra', 'Female', '4. Senior', 'CA', 'English']
    """
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
    """Retrieve metadata for a Google Sheet including properties and sheet information.

    This function fetches comprehensive metadata about a Google Sheet, including
    the document title, sheet names, and other properties. Results are cached
    to minimize API calls.

    Args:
        sheet_id: The unique identifier of the Google Sheet

    Returns:
        dict: Dictionary containing sheet metadata with keys like 'properties',
        'sheets', etc. Returns empty dict if the request fails.

    Raises:
        googleapiclient.errors.HttpError: If the API request fails

    Example:
        >>> metadata = get_metadata("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
        >>> print(metadata["properties"]["title"])
        "Sample Spreadsheet"
    """
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
    """Get the title of a Google Sheet document.

    This function retrieves the main title/name of the Google Sheet document
    from its metadata.

    Args:
        sheet_id: The unique identifier of the Google Sheet

    Returns:
        str: The title of the Google Sheet document, or "Untitled" if no title
        is found or if the request fails

    Example:
        >>> title = get_title("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
        >>> print(title)
        "Sample Spreadsheet"
    """
    metadata = get_metadata(sheet_id)
    return metadata.get("properties", {}).get("title", "Untitled")


def get_sheet_titles(sheet_id: str) -> List[str]:
    """Get the names of all individual sheets within a Google Sheet document.

    This function retrieves the names of all worksheets/tabs within a Google
    Sheet document from its metadata.

    Args:
        sheet_id: The unique identifier of the Google Sheet

    Returns:
        List[str]: List of sheet names within the document. Returns ["Sheet1"]
        as default if no sheets are found or if the request fails

    Example:
        >>> sheets = get_sheet_titles("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
        >>> print(sheets)
        ["Class Data", "Form Responses 1", "Summary"]
    """
    metadata = get_metadata(sheet_id)
    sheets = metadata.get("sheets", "")
    titles = [sheet.get("properties", {}).get("title", "Sheet1") for sheet in sheets]
    return titles
