import os
import json
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")
SCOPES = ["https://www.googleapis.com/auth/drive"]

creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

print("\n🔎 Buscando TODOS os arquivos e pastas que o robô consegue enxergar...")

results = drive_service.files().list(
    supportsAllDrives=True,
    includeItemsFromAllDrives=True,
    pageSize=50,
    fields="files(id, name, mimeType)"
).execute()

arquivos = results.get('files', [])

print(f"\n📌 Total de itens encontrados: {len(arquivos)}\n")
for item in arquivos:
    print(f" 📄 Nome: {item['name']}")
    print(f"    └─ Tipo (mimeType): {item['mimeType']}\n")