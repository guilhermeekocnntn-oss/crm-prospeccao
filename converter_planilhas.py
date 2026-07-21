import os
import io
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")
SCOPES = ["https://www.googleapis.com/auth/drive"]

creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

print("\n🔄 Buscando TODOS os arquivos Excel (.xlsx) no Google Drive (sem limite de páginas)...\n")

files = []
page_token = None

# Laço para percorrer TODAS as páginas de resultados do Drive
while True:
    results = drive_service.files().list(
        q="mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed=false",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
        pageSize=1000,
        fields="nextPageToken, files(id, name, parents)",
        pageToken=page_token
    ).execute()
    
    files.extend(results.get('files', []))
    page_token = results.get('nextPageToken', None)
    if not page_token:
        break

print(f"📌 Total real de arquivos .xlsx encontrados no Drive: {len(files)}\n")

for index, f in enumerate(files, 1):
    file_id = f['id']
    file_name = f['name']
    parents = f.get('parents', [])

    new_name = file_name.replace(".xlsx", "").replace(".XLSX", "")
    print(f"[{index}/{len(files)}] Convertendo: '{file_name}' ➔ '{new_name}'...")

    try:
        # Download do conteúdo do Excel
        request = drive_service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()

        fh.seek(0)

        # Upload convertendo para formato nativo do Google Sheets
        file_metadata = {
            'name': new_name,
            'mimeType': 'application/vnd.google-apps.spreadsheet',
            'parents': parents
        }
        media = MediaIoBaseUpload(
            fh, 
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
            resumable=True
        )

        new_file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields='id',
            supportsAllDrives=True
        ).execute()

        print(f"   ✅ Convertido!")
    except Exception as e:
        print(f"   ❌ Erro ao converter '{file_name}': {e}")

print("\n🎉 Conversão de TODOS os arquivos finalizada com sucesso!\n")