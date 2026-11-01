"""
drive_io.py — Acceso a la carpeta de Google Drive del corrector
================================================================
La web (app.py) usa una «cuenta de servicio» de Google para leer y escribir
en UNA carpeta de Drive que el usuario ha compartido con ella como Editor.

Estructura esperada dentro de esa carpeta compartida:

  Corrector Cambridge/            ← la carpeta compartida (su ID va en los secrets)
  ├── estudiantes/                ← los Excel (.xlsx) descargados de Google Forms
  ├── enunciados/                 ← enunciados (.docx, .md, .txt o Documento de Google)
  └── redacciones/
      ├── B1/  b1_rubrica.md, b1_essay.md, b1_report.md ...
      ├── B2/  b2_rubrica.md, b2_essay.md ...
      ├── C1/  ...
      └── C2/  ...

Los archivos de redacciones/ pueden ser .md, .txt, .docx o Documentos de Google:
lo que cuenta es el nombre (p. ej. «c1_essay» o «c1_essay.md»).
"""

import io
from pathlib import PurePath

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]

MIME_FOLDER = "application/vnd.google-apps.folder"
MIME_XLSX   = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME_DOCX   = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
MIME_GDOC   = "application/vnd.google-apps.document"

TEXT_EXTENSIONS = {".md", ".txt", ".docx", ""}   # "" = Documento de Google (sin extensión)


def _stem(name: str) -> str:
    """'C1_Essay.md' → 'c1_essay'. Los Documentos de Google no tienen extensión."""
    p = PurePath(name)
    return (p.stem if p.suffix.lower() in TEXT_EXTENSIONS else name).strip().lower()


class DriveClient:

    def __init__(self, service_account_info: dict, root_folder_id: str):
        creds = service_account.Credentials.from_service_account_info(
            dict(service_account_info), scopes=SCOPES
        )
        self.service = build("drive", "v3", credentials=creds, cache_discovery=False)
        self.root_folder_id = root_folder_id
        self.service_account_email = service_account_info.get("client_email", "")

    # ── Listados ─────────────────────────────────────────────────

    def list_children(self, folder_id: str) -> list:
        """Archivos y carpetas (no borrados) dentro de folder_id."""
        items, page_token = [], None
        while True:
            resp = self.service.files().list(
                q=f"'{folder_id}' in parents and trashed = false",
                fields="nextPageToken, files(id, name, mimeType, modifiedTime)",
                pageSize=200,
                orderBy="name",
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute(num_retries=3)
            items.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                return items

    def find_subfolder(self, parent_id: str, name: str):
        """ID de la subcarpeta llamada `name` (sin distinguir mayúsculas), o None."""
        for f in self.list_children(parent_id):
            if f["mimeType"] == MIME_FOLDER and f["name"].strip().lower() == name.lower():
                return f["id"]
        return None

    def check_root(self):
        """Comprueba que la carpeta compartida es accesible. Devuelve su nombre."""
        meta = self.service.files().get(
            fileId=self.root_folder_id, fields="id, name, mimeType",
            supportsAllDrives=True,
        ).execute(num_retries=3)
        if meta["mimeType"] != MIME_FOLDER:
            raise ValueError("El DRIVE_FOLDER_ID no corresponde a una carpeta.")
        return meta["name"]

    def list_excels(self) -> list:
        """Los .xlsx de estudiantes/."""
        folder = self.find_subfolder(self.root_folder_id, "estudiantes")
        if not folder:
            return []
        return [f for f in self.list_children(folder)
                if f["mimeType"] == MIME_XLSX or f["name"].lower().endswith(".xlsx")]

    def list_enunciados(self) -> list:
        """Los .docx / .md / .txt / Documentos de Google de enunciados/."""
        folder = self.find_subfolder(self.root_folder_id, "enunciados")
        if not folder:
            return []
        return [f for f in self.list_children(folder) if self._is_text_file(f)]

    def get_level_files(self, level: str) -> dict:
        """
        Archivos de referencia de un nivel (carpeta redacciones/<NIVEL>/).
        Devuelve {"rubric": archivo o None, "types": {"essay": archivo, ...}}.
        """
        result = {"rubric": None, "types": {}}
        redacciones = self.find_subfolder(self.root_folder_id, "redacciones")
        if not redacciones:
            return result
        level_folder = self.find_subfolder(redacciones, level)
        if not level_folder:
            return result

        prefix = f"{level.lower()}_"
        for f in self.list_children(level_folder):
            if not self._is_text_file(f):
                continue
            stem = _stem(f["name"])
            if not stem.startswith(prefix):
                continue
            kind = stem[len(prefix):]
            if kind == "rubrica":
                result["rubric"] = f
            elif kind:
                result["types"][kind] = f
        return result

    @staticmethod
    def _is_text_file(f: dict) -> bool:
        if f["mimeType"] == MIME_GDOC:
            return True
        if f["mimeType"] == MIME_FOLDER:
            return False
        return PurePath(f["name"]).suffix.lower() in {".md", ".txt", ".docx"}

    # ── Lectura ──────────────────────────────────────────────────

    def download_bytes(self, file_id: str) -> bytes:
        request = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
        buffer = io.BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk(num_retries=3)
        return buffer.getvalue()

    def read_text(self, f: dict) -> str:
        """Texto de un .md/.txt/.docx o Documento de Google."""
        if f["mimeType"] == MIME_GDOC:
            data = self.service.files().export(
                fileId=f["id"], mimeType="text/plain"
            ).execute(num_retries=3)
            return data.decode("utf-8-sig").strip()

        data = self.download_bytes(f["id"])
        if f["name"].lower().endswith(".docx") or f["mimeType"] == MIME_DOCX:
            from docx import Document
            doc = Document(io.BytesIO(data))
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip())
        return data.decode("utf-8-sig")

    # ── Escritura ────────────────────────────────────────────────

    def overwrite_xlsx(self, file_id: str, data: bytes):
        """
        Sustituye el contenido del .xlsx original (mismo archivo, mismo enlace).
        Drive guarda la versión anterior en el historial de versiones.
        """
        media = MediaIoBaseUpload(io.BytesIO(data), mimetype=MIME_XLSX, resumable=False)
        self.service.files().update(
            fileId=file_id, media_body=media, supportsAllDrives=True,
        ).execute(num_retries=3)

    def keep_current_version(self, file_id: str) -> bool:
        """
        Marca la versión actual del archivo como «Conservar para siempre» en el
        historial de Drive, para que el original no se borre aunque se suban
        muchas versiones nuevas (Drive solo guarda ~100 o 30 días). Es la copia
        de seguridad. Devuelve False si Drive no lo permite.
        """
        try:
            revs = self.service.revisions().list(
                fileId=file_id, fields="revisions(id)", pageSize=1000,
            ).execute(num_retries=3).get("revisions", [])
            if not revs:
                return False
            self.service.revisions().update(
                fileId=file_id, revisionId=revs[-1]["id"], body={"keepForever": True},
            ).execute(num_retries=3)
            return True
        except Exception:
            return False
