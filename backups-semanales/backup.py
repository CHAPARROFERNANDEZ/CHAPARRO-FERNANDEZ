"""
backup.py — Servicio de backup semanal, independiente de la app principal.

Se ejecuta con el cron de Railway (ver railway service config, no en este archivo) una vez a la
semana. Hace dos cosas, cada una en su propia carpeta de Google Drive, para tener una copia
FUERA de Railway (si Railway tuviera un problema grave, esto no se ve afectado):

1. Excel: descarga el inversiones.xlsx actual de Drive y lo vuelve a subir, con fecha en el
   nombre, a la carpeta "Backups semanales - Excel". Es una copia congelada, distinta del
   archivo "vivo" que edita la app — protege contra un guardado erróneo o una fila borrada
   sin querer en el archivo principal.

2. Postgres: hace un pg_dump completo (esquema + datos) de la base de datos y lo sube, también
   con fecha, a la carpeta "Backups semanales - SQL (Postgres)". Es un volcado SQL estándar,
   restaurable en cualquier Postgres (Railway, otro proveedor, tu propio servidor) con
   `psql < archivo.sql` — no depende de Railway para nada una vez descargado.

Cada semana se sube una copia nueva y se borran las más antiguas, manteniendo las últimas
RETENCION_SEMANAS (por defecto 12, ~3 meses de histórico).

Si una de las dos partes falla, la otra se intenta igualmente (no se bloquean entre sí) y el
fallo queda bien visible en los logs de Railway.
"""
import base64
import json
import os
import subprocess
import sys
from datetime import datetime, timezone
from io import BytesIO

GDRIVE_FILE_ID = "1CImiIbg7kSLrYNpWgzPHEBCmI3KRVlBX"  # inversiones.xlsx, el archivo "vivo"
CARPETA_BACKUP_EXCEL = "1yMwRBEckejztQbES8-5xPcvPsOW0RSzk"   # "Backups semanales - Excel"
CARPETA_BACKUP_SQL = "13a-ALj9vSDqtjsJ1vixwcLgcbbZ5c-bc"      # "Backups semanales - SQL (Postgres)"
RETENCION_SEMANAS = 12


def log(msg: str) -> None:
    print(f"[backup] {msg}", file=sys.stderr, flush=True)


def _servicio_drive():
    gcp_json_b64 = os.environ.get("GCP_SA_JSON_B64", "")
    if not gcp_json_b64:
        raise RuntimeError("No hay GCP_SA_JSON_B64 en el entorno — no se puede hablar con Drive.")
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    info = json.loads(base64.b64decode(gcp_json_b64))
    credenciales = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=credenciales)


def _podar_carpeta(servicio, carpeta_id: str, mantener: int) -> None:
    """Borra los archivos más antiguos de una carpeta de backups, dejando solo los últimos
    'mantener'. Nunca lanza excepción hacia arriba — un fallo aquí no debe tirar todo el backup
    (peor caso: se acumulan copias de más, no se pierde nada)."""
    try:
        archivos = servicio.files().list(
            q=f"'{carpeta_id}' in parents and trashed=false",
            fields="files(id, name, createdTime)",
            orderBy="createdTime desc",
            pageSize=100,
        ).execute().get("files", [])
        for viejo in archivos[mantener:]:
            servicio.files().delete(fileId=viejo["id"]).execute()
            log(f"Podado: {viejo['name']} (fuera de las últimas {mantener} copias).")
    except Exception as e:
        log(f"AVISO: no se pudo podar la carpeta {carpeta_id}: {e}")


def backup_excel(servicio, fecha: str) -> bool:
    log("--- Backup del Excel ---")
    try:
        from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

        request = servicio.files().export_media(
            fileId=GDRIVE_FILE_ID,
            mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        buffer = BytesIO()
        downloader = MediaIoBaseDownload(buffer, request)
        listo = False
        while not listo:
            _status, listo = downloader.next_chunk(num_retries=2)
        contenido = buffer.getvalue()
        log(f"Excel descargado ({len(contenido) / 1024:.0f} KB).")

        nombre = f"inversiones_backup_{fecha}.xlsx"
        media = MediaIoBaseUpload(
            BytesIO(contenido),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            resumable=False,
        )
        servicio.files().create(
            body={"name": nombre, "parents": [CARPETA_BACKUP_EXCEL]},
            media_body=media, fields="id",
        ).execute()
        log(f"Subido a Drive como '{nombre}'.")
        _podar_carpeta(servicio, CARPETA_BACKUP_EXCEL, RETENCION_SEMANAS)
        return True
    except Exception as e:
        log(f"ERROR en el backup del Excel: {e}")
        return False


def backup_postgres(servicio, fecha: str) -> bool:
    log("--- Backup de Postgres ---")
    database_url = os.environ.get("DATABASE_URL", "")
    if not database_url:
        log("ERROR: no hay DATABASE_URL en el entorno, se omite el backup de Postgres.")
        return False
    try:
        ruta_dump = "/tmp/backup.sql"
        resultado = subprocess.run(
            ["pg_dump", database_url, "--no-owner", "--no-privileges", "-f", ruta_dump],
            capture_output=True, text=True, timeout=600,
        )
        if resultado.returncode != 0:
            log(f"ERROR: pg_dump devolvió código {resultado.returncode}: {resultado.stderr[:2000]}")
            return False

        tamano_kb = os.path.getsize(ruta_dump) / 1024
        log(f"pg_dump completado ({tamano_kb:.0f} KB).")

        from googleapiclient.http import MediaFileUpload

        nombre = f"postgres_backup_{fecha}.sql"
        media = MediaFileUpload(ruta_dump, mimetype="application/sql", resumable=False)
        servicio.files().create(
            body={"name": nombre, "parents": [CARPETA_BACKUP_SQL]},
            media_body=media, fields="id",
        ).execute()
        log(f"Subido a Drive como '{nombre}'.")
        _podar_carpeta(servicio, CARPETA_BACKUP_SQL, RETENCION_SEMANAS)
        return True
    except subprocess.TimeoutExpired:
        log("ERROR: pg_dump tardó más de 10 minutos, se abortó.")
        return False
    except Exception as e:
        log(f"ERROR en el backup de Postgres: {e}")
        return False


def main() -> int:
    fecha = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    log(f"=== Backup semanal — {fecha} ===")

    try:
        servicio = _servicio_drive()
    except Exception as e:
        log(f"ERROR FATAL: no se pudo conectar con Drive, no se puede hacer ningún backup: {e}")
        return 1

    ok_excel = backup_excel(servicio, fecha)
    ok_sql = backup_postgres(servicio, fecha)

    if ok_excel and ok_sql:
        log("=== Backup semanal completado sin errores ===")
        return 0
    log(f"=== Backup semanal terminado CON AVISOS — Excel: {'OK' if ok_excel else 'FALLÓ'}, Postgres: {'OK' if ok_sql else 'FALLÓ'} ===")
    return 1


if __name__ == "__main__":
    sys.exit(main())
