import streamlit as st
import json
import io
import requests
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from google.auth.transport.requests import AuthorizedSession

# === SCOPES and Credential Loading ===
SCOPES = ['https://www.googleapis.com/auth/drive']

creds_dict = json.loads(st.secrets["google"]["service_account"])
credentials = service_account.Credentials.from_service_account_info(
    creds_dict, scopes=SCOPES
)

# Build Google Drive API client
drive_service = build('drive', 'v3', credentials=credentials)

# === List Files in a Google Drive Folder ===
def list_files_in_folder(folder_id, mime_types=None):
    query = f"'{folder_id}' in parents and trashed = false"
    if mime_types:
        mime_query = " or ".join([f"mimeType='{mt}'" for mt in mime_types])
        query += f" and ({mime_query})"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    return results.get('files', [])

# === Download File from Google Drive (SSL-safe) ===
def download_file(file_id):
    authed_session = AuthorizedSession(credentials)
    download_url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"

    response = authed_session.get(download_url)
    if response.status_code != 200:
        raise Exception(f"Failed to download file. HTTP {response.status_code}: {response.text}")

    return io.BytesIO(response.content)

# === Upload File to Google Drive ===
def upload_file(buffer, filename, folder_id):
    file_metadata = {
        'name': filename,
        'parents': [folder_id]
    }
    media = MediaIoBaseUpload(buffer, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    return file.get('id')
