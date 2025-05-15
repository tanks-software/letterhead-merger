import streamlit as st
import json
import io
import requests
from google.oauth2 import service_account
from google.auth.transport.requests import AuthorizedSession
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Load credentials from Streamlit secrets
SCOPES = ['https://www.googleapis.com/auth/drive']
creds_dict = json.loads(st.secrets["google"]["service_account"])
credentials = service_account.Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
authed_session = AuthorizedSession(credentials)

# === SAFELY list files using requests instead of .execute() ===
def list_files_in_folder(folder_id, mime_types=None):
    mime_filter = ""
    if mime_types:
        mime_query = " or ".join([f"mimeType='{mt}'" for mt in mime_types])
        mime_filter = f" and ({mime_query})"

    query = f"'{folder_id}' in parents and trashed = false{mime_filter}"
    url = f"https://www.googleapis.com/drive/v3/files?q={requests.utils.quote(query)}&fields=files(id,name)&pageSize=1000"
    
    response = authed_session.get(url)
    if response.status_code != 200:
        raise Exception(f"Failed to list files. HTTP {response.status_code}: {response.text}")
    return response.json().get("files", [])

# === SAFELY download file content ===
def download_file(file_id):
    url = f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media"
    response = authed_session.get(url)
    if response.status_code != 200:
        raise Exception(f"Failed to download file. HTTP {response.status_code}: {response.text}")
    return io.BytesIO(response.content)

# === SAFELY upload a file ===
def upload_file(buffer, filename, folder_id):
    metadata = {
        "name": filename,
        "parents": [folder_id]
    }
    drive_service = build('drive', 'v3', credentials=credentials)
    media = MediaIoBaseUpload(buffer, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    uploaded = drive_service.files().create(body=metadata, media_body=media, fields='id').execute()
    return uploaded.get("id")
