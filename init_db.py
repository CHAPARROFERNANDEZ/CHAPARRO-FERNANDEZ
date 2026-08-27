"""
init_db.py — Se ejecuta automáticamente ANTES de arrancar la app en cada despliegue de Railway
(configurado como preDeployCommand: `python3 init_db.py`).

Qué hace:
1. Crea (si no existen) las tablas de Postgres que reflejan las hojas del Excel del fondo.
2. Descarga la última versión del Excel desde Google Drive.
3. Sincroniza los datos hacia Postgres (sobreescribiendo, no acumulando).

Fase actual: "convivencia". El Excel de Drive sigue siendo la ÚNICA fuente de verdad que usa
app.py para leer y escribir — este script NO SE LEE todavía desde la aplicación. Su único
objetivo es mantener una copia siempre actualizada en Postgres, para que el código se pueda ir
migrando hoja por hoja sin ningún riesgo (si algo sale mal aquí, la app sigue funcionando igual
que hasta ahora, porque no depende de esto para nada).

Si esta sincronización falla por cualquier motivo, el script termina con código de salida 0
igualmente (no bloquea el arranque de la app) — solo deja constancia en los logs de deploy.
"""
import os
import sys
import json
import base64
from datetime import datetime, timezone

import requests
import pandas as pd
from sqlalchemy import create_engine, text

GDRIVE_FILE_ID = "1CImiIbg7kSLrYNpWgzPHEBCmI3KRVlBX"

# Hojas que se migran automáticamente "tal cual" (estructura inferida de sus columnas actuales).
# USUARIOS se trata aparte, con esquema explícito, por ser la más sensible.
HOJAS_AUTOMATICAS = [
    "INVERSIONES", "REINVERSIONES", "CONTROL_NOTAS", "CALENDARIO_NOTAS",
    "CALENDARIO_CALLS", "DEUDA_JORDI", "TRANSFERENCIAS_JORDI",
    "MOVIMIENTOS_MOTOCLICK", "REPARTO_DIVIDENDOS", "BORRADORES_NOTAS",
    "BORRADORES_INVERSIONES", "AUDITORIA_NOTAS", "LOG_IA_USO",
]


def log(msg: str):
    print(f"[init_db] {msg}", file=sys.stderr, flush=True)


def obtener_database_url() -> str:
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        raise RuntimeError("No hay DATABASE_URL en el entorno — ¿está Postgres conectado a este servicio?")
    # SQLAlchemy 2.x quiere el driver explícito para psycopg2.
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif url.startswith("postgresql://") and "+psycopg2" not in url:
        url = url.replace("postgresql://", "postgresql+psycopg2://", 1)
    return url


def descargar_excel() -> dict:
    """Descarga el Excel de Drive y devuelve {nombre_hoja: DataFrame}.

    Intenta primero la API autenticada (service account), igual que app.py — evita el lag de
    caché del enlace público de exportación. Si no hay credenciales o falla, cae al enlace
    público como respaldo.
    """
    from io import BytesIO

    gcp_json_b64 = os.environ.get("GCP_SA_JSON_B64", "")
    if gcp_json_b64:
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
            from googleapiclient.http import MediaIoBaseDownload

            info = json.loads(base64.b64decode(gcp_json_b64))
            credenciales = service_account.Credentials.from_service_account_info(
                info, scopes=["https://www.googleapis.com/auth/drive.readonly"],
            )
            servicio = build("drive", "v3", credentials=credenciales)
            request = servicio.files().export_media(
                fileId=GDRIVE_FILE_ID,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            buffer = BytesIO()
            downloader = MediaIoBaseDownload(buffer, request)
            done = False
            while not done:
                _status, done = downloader.next_chunk(num_retries=1)
            buffer.seek(0)
            return pd.read_excel(buffer, sheet_name=None)
        except Exception as e:
            log(f"Descarga autenticada falló ({e}), probando enlace público de respaldo.")

    url = f"https://docs.google.com/spreadsheets/d/{GDRIVE_FILE_ID}/export?format=xlsx"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return pd.read_excel(BytesIO(r.content), sheet_name=None)


def normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip().lower().replace(" ", "_").replace("ñ", "n")
        for c in df.columns
    ]
    # Columnas "unnamed: N" de Excel (celdas vacías arrastradas) no aportan nada — fuera.
    df = df.loc[:, [c for c in df.columns if not c.startswith("unnamed")]]
    return df


def crear_tabla_usuarios(engine):
    with engine.begin() as conn:
        conn.execute(text("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id SERIAL PRIMARY KEY,
                usuario TEXT NOT NULL,
                tipo_usuario TEXT NOT NULL,
                password TEXT NOT NULL DEFAULT '',
                debe_cambiar_password BOOLEAN NOT NULL DEFAULT FALSE,
                totp_email TEXT NOT NULL DEFAULT '',
                totp_activo BOOLEAN NOT NULL DEFAULT FALSE,
                actualizado_en TIMESTAMPTZ NOT NULL DEFAULT now(),
                UNIQUE (usuario, tipo_usuario)
            );
        """))


def sincronizar_usuarios(engine, df_usuarios: pd.DataFrame):
    if df_usuarios is None or df_usuarios.empty:
        log("USUARIOS: hoja vacía en el Excel, no hay nada que sincronizar todavía.")
        return
    df = normalizar_columnas(df_usuarios)
    columnas_esperadas = {"usuario", "tipo_usuario", "password", "debe_cambiar_password", "totp_secret", "totp_activo"}
    if not columnas_esperadas.issubset(df.columns):
        log(f"USUARIOS: faltan columnas esperadas ({columnas_esperadas - set(df.columns)}), se omite esta sincronización.")
        return
    with engine.begin() as conn:
        for _, fila in df.iterrows():
            usuario = str(fila.get("usuario", "")).strip()
            tipo = str(fila.get("tipo_usuario", "")).strip().lower()
            if not usuario or not tipo:
                continue
            conn.execute(text("""
                INSERT INTO usuarios (usuario, tipo_usuario, password, debe_cambiar_password, totp_email, totp_activo, actualizado_en)
                VALUES (:usuario, :tipo, :password, :debe_cambiar, :totp_email, :totp_activo, now())
                ON CONFLICT (usuario, tipo_usuario) DO UPDATE SET
                    password = EXCLUDED.password,
                    debe_cambiar_password = EXCLUDED.debe_cambiar_password,
                    totp_email = EXCLUDED.totp_email,
                    totp_activo = EXCLUDED.totp_activo,
                    actualizado_en = now();
            """), {
                "usuario": usuario,
                "tipo": tipo,
                "password": str(fila.get("password", "") or ""),
                "debe_cambiar": str(fila.get("debe_cambiar_password", "NO")).strip().upper() == "SI",
                "totp_email": str(fila.get("totp_secret", "") or ""),
                "totp_activo": str(fila.get("totp_activo", "NO")).strip().upper() == "SI",
            })
    log(f"USUARIOS: {len(df)} fila(s) sincronizada(s).")


def _parsear_fecha_celda(valor):
    """Parsea UNA celda de fecha admitiendo tanto un Timestamp/datetime ya parseado por
    pandas (caso normal) como texto en distintos formatos (DD/MM/AAAA, AAAA-MM-DD, con o sin
    hora) — que es lo que ocurre cuando alguien escribe la fecha a mano en el Sheet y Google
    Sheets la guarda como texto en vez de como fecha real, mezclada con otras celdas de la
    misma columna que sí son fechas genuinas.

    BUG que corrige esta función (encontrado 27/08/2026): cuando una columna de fecha mezcla
    fechas reales (dtype datetime) con texto (p.ej. "2026-08-24" sin hora, tecleado a mano),
    pandas marca la columna ENTERA como dtype 'object' en vez de 'datetime64'. El código antiguo
    comprobaba is_datetime64_any_dtype(columna) para decidir si normalizarla, así que esa
    comprobación fallaba para la columna completa y NINGUNA fecha se normalizaba de forma
    explícita — las celdas de texto llegaban a Postgres tal cual (sin hora), y luego
    pd.to_datetime() en postgres_reader.py no siempre las reconocía, dejándolas como NaT
    (nulas). Como el capital activo exige fecha_inversion no nula, esas filas desaparecían en
    silencio del dashboard aunque el resto de sus datos estuviera perfecto (caso real: OP132 y
    OP138, 95.000€ desaparecidos del capital activo). Aquí parseamos celda a celda, sin
    depender del dtype de la columna, así da igual cómo se haya escrito la fecha origen."""
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, (pd.Timestamp, datetime)):
        return pd.Timestamp(valor)
    texto = str(valor).strip()
    if not texto:
        return pd.NaT
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return pd.to_datetime(texto, format=fmt)
        except (ValueError, TypeError):
            continue
    return pd.to_datetime(texto, errors="coerce", dayfirst=True)


def sincronizar_hoja_automatica(engine, nombre_hoja: str, df: pd.DataFrame) -> int:
    if df is None or df.empty:
        log(f"{nombre_hoja}: hoja vacía, se omite (la tabla se crea igualmente si no existía).")
        return 0
    tabla = f"excel_{nombre_hoja.lower()}"
    df = normalizar_columnas(df)
    # Fechas y JSON como texto plano por ahora — sin fricción de tipos mientras estamos en fase
    # de espejo/lectura. Se afinarán tipos y claves cuando migremos esta hoja a código de verdad.
    #
    # IMPORTANTE: cualquier columna cuyo nombre contenga "fecha" se parsea celda a celda con
    # _parsear_fecha_celda ANTES de mirar el dtype. No nos fiamos de is_datetime64_any_dtype
    # sobre la columna completa: una sola celda de texto mezclada con fechas reales hace que
    # pandas marque la columna entera como 'object' y el chequeo de dtype se salte esa columna
    # por completo (ver docstring de _parsear_fecha_celda para el caso real que esto causó).
    columnas_fecha = [c for c in df.columns if "fecha" in c]
    for col in columnas_fecha:
        celda_estaba_rellena = df[col].notna() & (df[col].astype(str).str.strip() != "")
        df[col] = df[col].apply(_parsear_fecha_celda)
        # Si una celda que SÍ tenía contenido en origen se quedó en NaT tras el parseo, es un
        # dato real perdido — lo avisamos en el log en vez de dejarlo pasar en silencio, para
        # detectarlo en el propio deploy en vez de descubrirlo semanas después comparando el
        # dashboard contra el Excel a mano (como pasó con OP132 y OP138 el 27/08/2026).
        perdidas = celda_estaba_rellena & df[col].isna()
        if perdidas.any():
            n_perdidas = int(perdidas.sum())
            log(f"AVISO {nombre_hoja}.{col}: {n_perdidas} fecha(s) no se pudieron interpretar y quedan vacías. Revisa esas filas en el Sheet (formato de fecha no reconocido).")
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")
    df.to_sql(tabla, engine, if_exists="replace", index=False, method="multi", chunksize=500)
    log(f"{nombre_hoja} → tabla '{tabla}': {len(df)} fila(s) sincronizada(s).")
    return len(df)


def main() -> dict:
    """Ejecuta la sincronización completa Drive → Postgres y devuelve un resumen:
    {"ok": bool, "detalle": [str, ...], "error": str | None}.
    Se sigue pudiendo ejecutar como script (preDeployCommand) y también se puede importar y
    llamar en caliente desde la app (botón 'Traer cambios de Drive ahora'), sin duplicar
    lógica."""
    detalle = []
    try:
        database_url = obtener_database_url()
    except Exception as e:
        msg = f"AVISO: {e} — se omite la sincronización con Postgres, la app sigue arrancando con Excel/Drive como siempre."
        log(msg)
        return {"ok": False, "detalle": [], "error": msg}

    try:
        engine = create_engine(database_url, pool_pre_ping=True)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        log("Conexión a Postgres OK.")
    except Exception as e:
        msg = f"AVISO: no se pudo conectar a Postgres ({e}) — se omite la sincronización, la app sigue arrancando igual."
        log(msg)
        return {"ok": False, "detalle": [], "error": msg}

    try:
        hojas = descargar_excel()
        log(f"Excel descargado de Drive: {len(hojas)} hoja(s) encontradas.")
    except Exception as e:
        msg = f"AVISO: no se pudo descargar el Excel de Drive ({e}) — se omite la sincronización."
        log(msg)
        return {"ok": False, "detalle": [], "error": msg}

    try:
        crear_tabla_usuarios(engine)
        sincronizar_usuarios(engine, hojas.get("USUARIOS"))
        n_usuarios = len(hojas.get("USUARIOS")) if hojas.get("USUARIOS") is not None else 0
        detalle.append(f"USUARIOS: {n_usuarios} fila(s)")
    except Exception as e:
        log(f"ERROR sincronizando USUARIOS: {e}")
        detalle.append(f"USUARIOS: ERROR ({e})")

    for nombre_hoja in HOJAS_AUTOMATICAS:
        try:
            n = sincronizar_hoja_automatica(engine, nombre_hoja, hojas.get(nombre_hoja))
            detalle.append(f"{nombre_hoja}: {n} fila(s)")
        except Exception as e:
            log(f"ERROR sincronizando {nombre_hoja}: {e}")
            detalle.append(f"{nombre_hoja}: ERROR ({e})")

    log("Sincronización con Postgres completada.")
    return {"ok": True, "detalle": detalle, "error": None}


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Red de seguridad final: bajo NINGUNA circunstancia este script debe hacer fallar el
        # preDeployCommand y bloquear el arranque de la app. Si algo no contemplado revienta acá,
        # se loguea y se sigue — la sincronización con Postgres simplemente no se actualizó esta
        # vez, pero Drive sigue siendo la fuente de verdad y la app arranca igual.
        log(f"ERROR INESPERADO no controlado: {e}")
    sys.exit(0)
