"""Funções auxiliares para ler/escrever arquivos no Google Drive via conta de serviço,
substituindo o acesso direto ao sistema de arquivos que o pipeline usa no Mac (Google
Drive Desktop montado em ~/Library/CloudStorage/...).
"""
import io
import json
import os

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

SCOPES = ['https://www.googleapis.com/auth/drive']

FOLDER_MIME = 'application/vnd.google-apps.folder'


def get_service():
    raw = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if not raw:
        raise RuntimeError('GOOGLE_SERVICE_ACCOUNT_JSON não está definida.')
    info = json.loads(raw)
    creds = service_account.Credentials.from_service_account_info(info, scopes=SCOPES)
    return build('drive', 'v3', credentials=creds, cache_discovery=False)


def find_child(service, parent_id, name, mime_type=None):
    """Acha um item (pasta ou arquivo) pelo nome exato dentro de uma pasta. None se não achar."""
    safe_name = name.replace("'", "\\'")
    q = f"'{parent_id}' in parents and name = '{safe_name}' and trashed = false"
    if mime_type:
        q += f" and mimeType = '{mime_type}'"
    resp = service.files().list(q=q, fields='files(id, name, mimeType, modifiedTime)',
                                 pageSize=5, supportsAllDrives=True,
                                 includeItemsFromAllDrives=True).execute()
    arquivos = resp.get('files', [])
    return arquivos[0] if arquivos else None


def find_root_folder(service, name):
    """Acha uma pasta pelo nome em qualquer lugar acessível pela conta de serviço
    (usado só pra achar a pasta raiz "Arrancada de Vendas", compartilhada com a conta
    de serviço — não tem um 'parent' conhecido de antemão).

    Usa "contains" em vez de igualdade exata porque o nome real da pasta no Drive tem
    um espaço sobrando no final ("Arrancada de Vendas ") — já mordeu o pipeline local
    mais de uma vez (ver memória do projeto), então aqui já nasce tolerante a isso.
    """
    safe_name = name.strip().replace("'", "\\'")
    q = f"name contains '{safe_name}' and mimeType = '{FOLDER_MIME}' and trashed = false"
    resp = service.files().list(q=q, fields='files(id, name)', pageSize=10,
                                 supportsAllDrives=True, includeItemsFromAllDrives=True).execute()
    arquivos = resp.get('files', [])
    if not arquivos:
        raise RuntimeError(f"Pasta '{name}' não encontrada — confirme que foi compartilhada com a conta de serviço.")
    # prioriza um match cujo nome (sem espaços nas pontas) seja idêntico ao pedido
    for f in arquivos:
        if f['name'].strip() == name.strip():
            return f['id']
    return arquivos[0]['id']


def listar_arquivos(service, folder_id):
    """Lista todos os arquivos (não-pastas) dentro de uma pasta."""
    arquivos = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{folder_id}' in parents and mimeType != '{FOLDER_MIME}' and trashed = false",
            fields='nextPageToken, files(id, name, mimeType, modifiedTime, size)',
            pageSize=200, pageToken=page_token,
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        arquivos.extend(resp.get('files', []))
        page_token = resp.get('nextPageToken')
        if not page_token:
            break
    return arquivos


def baixar_arquivo(service, file_id, mime_type, destino):
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    if mime_type.startswith('application/vnd.google-apps'):
        # Planilha nativa do Google Sheets (não deveria acontecer aqui, mas por segurança
        # exporta como xlsx em vez de falhar).
        request = service.files().export_media(
            fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    else:
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)
    fh = io.FileIO(destino, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    fh.close()
    return destino


def baixar_pasta(service, folder_id, destino_dir):
    """Baixa todos os arquivos de uma pasta do Drive pra uma pasta local, preservando nomes."""
    os.makedirs(destino_dir, exist_ok=True)
    baixados = []
    for item in listar_arquivos(service, folder_id):
        destino = os.path.join(destino_dir, item['name'])
        baixar_arquivo(service, item['id'], item['mimeType'], destino)
        baixados.append(destino)
    return baixados


def achar_ou_criar_pasta(service, parent_id, nome):
    existente = find_child(service, parent_id, nome, FOLDER_MIME)
    if existente:
        return existente['id']
    metadata = {'name': nome, 'mimeType': FOLDER_MIME, 'parents': [parent_id]}
    pasta = service.files().create(body=metadata, supportsAllDrives=True, fields='id').execute()
    return pasta['id']


def enviar_arquivo(service, folder_id, nome, caminho_local, mime_type='application/octet-stream'):
    """Cria ou atualiza (se já existir com o mesmo nome) um arquivo numa pasta do Drive."""
    existente = find_child(service, folder_id, nome)
    media = MediaFileUpload(caminho_local, mimetype=mime_type, resumable=True)
    if existente:
        arquivo = service.files().update(fileId=existente['id'], media_body=media,
                                          supportsAllDrives=True).execute()
    else:
        metadata = {'name': nome, 'parents': [folder_id]}
        arquivo = service.files().create(body=metadata, media_body=media,
                                          supportsAllDrives=True, fields='id').execute()
    return arquivo
