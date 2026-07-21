import os
import sys
import io
import json
import time
import unicodedata
import gspread

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
import openpyxl
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_PATH = os.path.join(BASE_DIR, "credentials.json")

# Suporte ao sistema de arquivos do Vercel (/tmp)
if os.environ.get("VERCEL"):
    CACHE_FILE = "/tmp/leads_cache.json"
else:
    CACHE_FILE = os.path.join(BASE_DIR, "leads_cache.json")

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

ESTADOS_MAP = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapá", "AM": "Amazonas",
    "BA": "Bahia", "CE": "Ceará", "DF": "Distrito Federal", "ES": "Espírito Santo",
    "GO": "Goiás", "MA": "Maranhão", "MT": "Mato Grosso", "MS": "Mato Grosso do Sul",
    "MG": "Minas Gerais", "PA": "Pará", "PB": "Paraíba", "PR": "Paraná",
    "PE": "Pernambuco", "PI": "Piauí", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte",
    "RS": "Rio Grande do Sul", "RO": "Rondônia", "RR": "Roraima", "SC": "Santa Catarina",
    "SP": "São Paulo", "SE": "Sergipe", "TO": "Tocantins"
}

folder_cache = {}

def normalizar(texto):
    if not texto:
        return ""
    nfkd = unicodedata.normalize('NFD', str(texto))
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower().strip()

def criar_servicos():
    """Cria instâncias isoladas usando credentials.json local ou a variável GOOGLE_CREDENTIALS_JSON no Vercel"""
    creds_json = os.environ.get("GOOGLE_CREDENTIALS_JSON") or os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
    
    if creds_json:
        try:
            creds_info = json.loads(creds_json)
            creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
        except Exception as e:
            print(f"Erro ao decodificar GOOGLE_CREDENTIALS_JSON: {e}")
            if os.path.exists(CREDENTIALS_PATH):
                creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
            else:
                return None, None
    elif os.path.exists(CREDENTIALS_PATH):
        creds = Credentials.from_service_account_file(CREDENTIALS_PATH, scopes=SCOPES)
    else:
        print("⚠️ Arquivo credentials.json não encontrado e GOOGLE_CREDENTIALS_JSON não configurada.")
        return None, None

    try:
        gclient = gspread.authorize(creds)
        drive_service = build('drive', 'v3', credentials=creds)
        return gclient, drive_service
    except Exception as e:
        print(f"Erro ao inicializar serviços do Google: {e}")
        return None, None

def obter_info_pasta(drive_service, folder_id):
    if not folder_id:
        return "", []
    if folder_id in folder_cache:
        return folder_cache[folder_id]['name'], folder_cache[folder_id]['parents']
    
    try:
        res = drive_service.files().get(
            fileId=folder_id, 
            fields='name, parents', 
            supportsAllDrives=True
        ).execute()
        name = res.get('name', '')
        parents = res.get('parents', [])
        folder_cache[folder_id] = {'name': name, 'parents': parents}
        return name, parents
    except Exception:
        return "", []

def identificar_estado_da_planilha(drive_service, parents):
    if not parents:
        return "OUTROS", "Geral", "Sem Pasta"
    
    curr_id = parents[0]
    visited = set()
    
    while curr_id and curr_id not in visited:
        visited.add(curr_id)
        folder_name, parent_parents = obter_info_pasta(drive_service, curr_id)
        if not folder_name:
            break
            
        norm_folder = normalizar(folder_name)
        
        for sigla, nome in ESTADOS_MAP.items():
            norm_nome = normalizar(nome)
            norm_sigla = normalizar(sigla)
            
            if norm_nome in norm_folder or norm_folder in norm_nome or norm_sigla == norm_folder:
                return sigla, nome, folder_name
                
        curr_id = parent_parents[0] if parent_parents else None
        
    return "OUTROS", "Geral", "Geral"

def extrair_valor(row, chaves_possiveis):
    for key in row.keys():
        norm_key = normalizar(key)
        for chave in chaves_possiveis:
            if normalizar(chave) in norm_key:
                val = row[key]
                return str(val).strip() if val is not None else ""
    return ""

def processar_item_planilha(item):
    # Conexão dedicada para esta thread
    gclient, drive_service = criar_servicos()
    
    sheet_id = item['id']
    nome_planilha = item['name']
    mime_type = item.get('mimeType', '')
    parents = item.get('parents', [])
    
    sigla_uf, nome_uf, pasta_origem = identificar_estado_da_planilha(drive_service, parents)
    leads_item = []

    max_tentativas = 3
    for tentativa in range(1, max_tentativas + 1):
        try:
            if mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
                request = drive_service.files().get_media(fileId=sheet_id)
                fh = io.BytesIO()
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk()
                fh.seek(0)
                
                wb = openpyxl.load_workbook(fh, data_only=True)
                for sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    rows = list(ws.iter_rows(values_only=True))
                    if not rows or len(rows) < 2:
                        continue
                    headers = [str(h).strip() if h is not None else f"col_{i}" for i, h in enumerate(rows[0])]
                    
                    for r_idx, row_values in enumerate(rows[1:], start=2):
                        row = {headers[i]: row_values[i] if i < len(row_values) and row_values[i] is not None else "" for i in range(len(headers))}
                        empresa = extrair_valor(row, ["Nome da Empresa", "Empresa", "Razao Social", "Nome"])
                        if empresa:
                            leads_item.append({
                                "sheet_id": sheet_id,
                                "linha_id": r_idx,
                                "uf": sigla_uf,
                                "uf_nome": nome_uf,
                                "pasta": pasta_origem,
                                "macroregiao": nome_planilha,
                                "aba": sheet_name,
                                "cidade": sheet_name.split("_")[0] if "_" in sheet_name else sheet_name,
                                "empresa": empresa,
                                "tipo": extrair_valor(row, ["Tipo", "Categoria"]),
                                "bairro": extrair_valor(row, ["Bairro", "Regiao", "Bairro / Região"]),
                                "telefone": extrair_valor(row, ["Telefone", "WhatsApp", "Contato", "Telefone / WhatsApp"]),
                                "decisor": extrair_valor(row, ["Decisor", "Nome do Decisor", "Socio"]),
                                "instagram_site": extrair_valor(row, ["Instagram", "Site", "Redes"]),
                                "marca_propria": extrair_valor(row, ["Marca Propria", "Possui Marca"]),
                                "potencial": extrair_valor(row, ["Potencial", "Potencial do Cliente"]) or "Médio",
                                "status": extrair_valor(row, ["Status", "Status do Contato"]) or "A Ligar (Novo)",
                                "data_ultimo": extrair_valor(row, ["Data Ultimo", "Ultimo Contato"]),
                                "data_retorno": extrair_valor(row, ["Data de Retorno", "Retorno"]),
                                "resumo": extrair_valor(row, ["Resumo", "Objeção", "Observacao", "Notas"])
                            })
            else:
                sheet = gclient.open_by_key(sheet_id)
                for worksheet in sheet.worksheets():
                    aba_title = worksheet.title
                    try:
                        registros = worksheet.get_all_records()
                    except Exception:
                        continue
                    for index, row in enumerate(registros, start=2):
                        empresa = extrair_valor(row, ["Nome da Empresa", "Empresa", "Razao Social", "Nome"])
                        if empresa:
                            leads_item.append({
                                "sheet_id": sheet_id,
                                "linha_id": index,
                                "uf": sigla_uf,
                                "uf_nome": nome_uf,
                                "pasta": pasta_origem,
                                "macroregiao": nome_planilha,
                                "aba": aba_title,
                                "cidade": aba_title.split("_")[0] if "_" in aba_title else aba_title,
                                "empresa": empresa,
                                "tipo": extrair_valor(row, ["Tipo", "Categoria"]),
                                "bairro": extrair_valor(row, ["Bairro", "Regiao", "Bairro / Região"]),
                                "telefone": extrair_valor(row, ["Telefone", "WhatsApp", "Contato", "Telefone / WhatsApp"]),
                                "decisor": extrair_valor(row, ["Decisor", "Nome do Decisor", "Socio"]),
                                "instagram_site": extrair_valor(row, ["Instagram", "Site", "Redes"]),
                                "marca_propria": extrair_valor(row, ["Marca Propria", "Possui Marca"]),
                                "potencial": extrair_valor(row, ["Potencial", "Potencial do Cliente"]) or "Médio",
                                "status": extrair_valor(row, ["Status", "Status do Contato"]) or "A Ligar (Novo)",
                                "data_ultimo": extrair_valor(row, ["Data Ultimo", "Ultimo Contato"]),
                                "data_retorno": extrair_valor(row, ["Data de Retorno", "Retorno"]),
                                "resumo": extrair_valor(row, ["Resumo", "Objeção", "Observacao", "Notas"])
                            })
            break # Leitura concluída com sucesso

        except Exception as e:
            if tentativa == max_tentativas:
                print(f" ⚠️ Falha definitiva ao processar '{nome_planilha}': {e}")
            else:
                time.sleep(1) # Pausa rápida antes de tentar de novo
                
    return leads_item

def obter_todos_leads_eko(force_refresh=False):
    if not force_refresh and os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                todos_leads = json.load(f)
                print(f"⚡ CACHE LOCAL CARREGADO INSTANTANEAMENTE: {len(todos_leads)} leads.")
                return todos_leads
        except Exception:
            pass

    gclient, drive_service = criar_servicos()
    if not drive_service:
        print("⚠️ Conexão com o Google Drive indisponível (credenciais ausentes).")
        return []
    
    print("\n" + "🚀 " + "="*55)
    print(" Buscando e Sincronizando Planilhas do Drive (Conexões Estáveis)...")
    
    planilhas = []
    page_token = None
    query = "(mimeType='application/vnd.google-apps.spreadsheet' or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and trashed=false"
    
    try:
        while True:
            results = drive_service.files().list(
                q=query,
                fields="nextPageToken, files(id, name, mimeType, parents)",
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
                pageSize=1000,
                pageToken=page_token
            ).execute()
            
            planilhas.extend(results.get('files', []))
            page_token = results.get('nextPageToken', None)
            if not page_token:
                break
    except Exception as e:
        print(f" ❌ ERRO ao listar arquivos do Drive: {e}")
        return []
    
    print(f" 📌 Total de planilhas encontradas no Drive: {len(planilhas)}")
    print(" ⚡ Processando com 4 conexões simultâneas e re-tentativas automáticas...")
    
    todos_leads = []
    
    # 4 Workers para garantia de estabilidade total na rede
    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(processar_item_planilha, item) for item in planilhas]
        for future in as_completed(futures):
            resultado = future.result()
            todos_leads.extend(resultado)

    try:
        with open(CACHE_FILE, 'w', encoding='utf-8') as f:
            json.dump(todos_leads, f, ensure_ascii=False, indent=2)
        print(f" 💾 Cache salvo com sucesso em '{CACHE_FILE}'!")
    except Exception as e:
        print(f" ⚠️ Erro ao salvar cache: {e}")

    print(f"\n ✅ TOTAL GERAL DE LEADS CARREGADOS: {len(todos_leads)}")
    print("="*55 + "\n")
    return todos_leads

def atualizar_lead_tempo_real(sheet_id, nome_aba, linha, coluna_nome, novo_valor):
    gclient, drive_service = criar_servicos()
    if not drive_service:
        return False
    
    colunas = {
        "empresa": 1, "tipo": 2, "bairro": 3, "telefone": 4,
        "decisor": 5, "instagram_site": 6, "marca_propria": 7,
        "potencial": 8, "status": 9, "data_ultimo": 10,
        "data_retorno": 11, "resumo": 12
    }
    
    col_idx = colunas.get(coluna_nome)
    if not col_idx:
        return False

    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, 'r', encoding='utf-8') as f:
                leads_cache = json.load(f)
            for lead in leads_cache:
                if str(lead.get('sheet_id')) == str(sheet_id) and str(lead.get('aba')) == str(nome_aba) and str(lead.get('linha_id')) == str(linha):
                    lead[coluna_nome] = novo_valor
            with open(CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(leads_cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Erro ao atualizar cache local: {e}")

    try:
        res = drive_service.files().get(fileId=sheet_id, fields='mimeType', supportsAllDrives=True).execute()
        mime_type = res.get('mimeType', '')

        if mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            request = drive_service.files().get_media(fileId=sheet_id)
            fh = io.BytesIO()
            downloader = MediaIoBaseDownload(fh, request)
            done = False
            while not done:
                _, done = downloader.next_chunk()
            fh.seek(0)

            wb = openpyxl.load_workbook(fh)
            if nome_aba in wb.sheetnames:
                ws = wb[nome_aba]
                ws.cell(row=int(linha), column=int(col_idx), value=str(novo_valor))
                
                output = io.BytesIO()
                wb.save(output)
                output.seek(0)
                
                media = MediaIoBaseUpload(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
                drive_service.files().update(fileId=sheet_id, media_body=media, supportsAllDrives=True).execute()
                return True
        else:
            sheet = gclient.open_by_key(sheet_id)
            worksheet = sheet.worksheet(nome_aba)
            worksheet.update_cell(int(linha), col_idx, str(novo_valor))
            return True
    except Exception as e:
        print(f"❌ Erro ao atualizar lead no Drive: {e}")
        return False