import os
import pickle

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


SCOPES = [
    "https://www.googleapis.com/auth/gmail.modify",
    "https://www.googleapis.com/auth/gmail.send"
]

TOKEN_FILE = "token.pkl"
CREDENTIALS_FILE = "credentials.json"


def main():
    creds = None

    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "rb") as token:
            creds = pickle.load(token)

    if creds and creds.valid:
        print("Token Gmail valid.")
        return

    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        if not os.path.exists(CREDENTIALS_FILE):
            raise FileNotFoundError("Nu exista credentials.json in folderul principal.")

        flow = InstalledAppFlow.from_client_secrets_file(
            CREDENTIALS_FILE,
            SCOPES
        )

        creds = flow.run_local_server(
            port=0
        )

    with open(TOKEN_FILE, "wb") as token:
        pickle.dump(creds, token)

    print("Token Gmail creat cu permisiuni pentru citire, modificare si trimitere email.")


if __name__ == "__main__":
    main()