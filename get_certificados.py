"""
Baixa certificados de subpastas do Google Drive para a pasta local 'certificados'.
Substitui arquivos locais se a versão do Drive for mais recente.

Setup necessário (apenas uma vez):
  1. Acesse https://console.cloud.google.com/
  2. Crie um projeto e ative a API "Google Drive API"
  3. Em "APIs e Serviços > Credenciais", crie uma credencial OAuth 2.0 (tipo: Desktop App)
  4. Baixe o JSON e salve como credentials.json na mesma pasta deste script
  5. pip install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client
"""

import io
from pathlib import Path
from datetime import datetime, timezone

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
FOLDER_ID = "1KjdEHy1PmQ1YeQLK3HZldtUoSWAzSLeq"
LOCAL_DIR = Path(__file__).parent / "certificados"
CREDENTIALS_FILE = Path(__file__).parent / "credentials.json"
TOKEN_FILE = Path(__file__).parent / "token.json"


def autenticar():
    creds = None
    if TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not CREDENTIALS_FILE.exists():
                raise FileNotFoundError(
                    f"Arquivo '{CREDENTIALS_FILE}' não encontrado.\n"
                    "Siga as instruções no topo deste arquivo para configurar as credenciais."
                )
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        TOKEN_FILE.write_text(creds.to_json())
    return creds


def listar_arquivos(service, pasta_id):
    """Retorna todos os arquivos (não pastas) dentro de uma pasta do Drive."""
    arquivos = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{pasta_id}' in parents and mimeType != 'application/vnd.google-apps.folder' and trashed = false",
            fields="nextPageToken, files(id, name, modifiedTime, mimeType)",
            pageToken=page_token,
        ).execute()
        arquivos.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return arquivos


def listar_subpastas(service, pasta_id):
    """Retorna todas as subpastas dentro de uma pasta do Drive."""
    subpastas = []
    page_token = None
    while True:
        resp = service.files().list(
            q=f"'{pasta_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false",
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
        ).execute()
        subpastas.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return subpastas


def baixar_arquivo(service, file_id, destino: Path):
    request = service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    destino.write_bytes(buffer.getvalue())


def deve_substituir(arquivo_drive: dict, caminho_local: Path) -> bool:
    """Retorna True se o arquivo do Drive é mais recente que o local."""
    if not caminho_local.exists():
        return True
    modified_drive = datetime.fromisoformat(
        arquivo_drive["modifiedTime"].replace("Z", "+00:00")
    )
    modified_local = datetime.fromtimestamp(
        caminho_local.stat().st_mtime, tz=timezone.utc
    )
    return modified_drive > modified_local


def main():
    LOCAL_DIR.mkdir(exist_ok=True)
    creds = autenticar()
    service = build("drive", "v3", credentials=creds)

    print(f"Buscando subpastas em '{FOLDER_ID}'...")
    subpastas = listar_subpastas(service, FOLDER_ID)

    # Também coleta arquivos diretamente na raiz da pasta
    todas_as_fontes = [{"id": FOLDER_ID, "name": "(raiz)"}] + subpastas

    total_baixados = 0
    total_ignorados = 0

    for pasta in todas_as_fontes:
        arquivos = listar_arquivos(service, pasta["id"])
        if not arquivos:
            continue
        for arq in arquivos:
            destino = LOCAL_DIR / arq["name"]
            if deve_substituir(arq, destino):
                print(f"  Baixando: {arq['name']} (pasta: {pasta['name']})")
                baixar_arquivo(service, arq["id"], destino)
                total_baixados += 1
            else:
                print(f"  Ignorando (já atualizado): {arq['name']}")
                total_ignorados += 1

    print(f"\nConcluído. {total_baixados} arquivo(s) baixado(s), {total_ignorados} já estavam atualizados.")


if __name__ == "__main__":
    main()
