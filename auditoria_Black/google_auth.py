from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2 import service_account


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

BASE_DIR = Path(__file__).resolve().parent
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

WRITER_CREDENTIALS_FILE = BASE_DIR / "writer_credentials.json"

WRITER_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]


def get_writer_service():
    """
    Serviço exclusivo para escrita.

    Usa a Service Account, não o OAuth pessoal.
    """

    credentials = service_account.Credentials.from_service_account_file(
        WRITER_CREDENTIALS_FILE,
        scopes=WRITER_SCOPES,
    )

    return build(
        "sheets",
        "v4",
        credentials=credentials,
    )



def get_credentials():
    creds = None

    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(
            TOKEN_FILE,
            SCOPES,
        )

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())

        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE,
                SCOPES,
            )

            creds = flow.run_local_server(port=0)

        TOKEN_FILE.write_text(
            creds.to_json(),
            encoding="utf-8",
        )

    return creds


def get_sheets_service():
    creds = get_credentials()

    return build(
        "sheets",
        "v4",
        credentials=creds,
    )


def get_drive_service():
    creds = get_credentials()

    return build(
        "drive",
        "v3",
        credentials=creds,
    )