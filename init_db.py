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
    """Descarga el Excel público de Drive y devuelve {nombre_hoja: DataFrame}."""
    url = f"https://docs.google.com/spreadsheets/d/{GDRIVE_FILE_ID}/export?format=xlsx"
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    from io import BytesIO
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


def sincronizar_hoja_automatica(engine, nombre_hoja: str, df: pd.DataFrame):
    if df is None or df.empty:
        log(f"{nombre_hoja}: hoja vacía, se omite (la tabla se crea igualmente si no existía).")
        return
    tabla = f"excel_{nombre_hoja.lower()}"
    df = normalizar_columnas(df)
    # Fechas y JSON como texto plano por ahora — sin fricción de tipos mientras estamos en fase
    # de espejo/lectura. Se afinarán tipos y claves cuando migremos esta hoja a código de verdad.
    for col in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime("%Y-%m-%d %H:%M:%S")
    df.to_sql(tabla, engine, if_exists="replace", index=False, method="multi", chunksize=500)
    log(f"{nombre_hoja} → tabla '{tabla}': {len(df)} fila(s) sincronizada(s).")


def main():
    try:
        database_url = obtener_database_url()
    except Exception as e:
        log(f"AVISO: {e} — se omite la sincronización con Postgres, la app sigue arrancando con Excel/Drive como siempre.")
        return

    try:
        engine = create_engine(database_url, pool_pre_ping=True)
        with engine.connect() as conn:
            conn.execute(text("SELECT 1"))
        log("Conexión a Postgres OK.")
    except Exception as e:
        log(f"AVISO: no se pudo conectar a Postgres ({e}) — se omite la sincronización, la app sigue arrancando igual.")
        return

    try:
        hojas = descargar_excel()
        log(f"Excel descargado de Drive: {len(hojas)} hoja(s) encontradas.")
    except Exception as e:
        log(f"AVISO: no se pudo descargar el Excel de Drive ({e}) — se omite la sincronización.")
        return

    try:
        crear_tabla_usuarios(engine)
        sincronizar_usuarios(engine, hojas.get("USUARIOS"))
    except Exception as e:
        log(f"ERROR sincronizando USUARIOS: {e}")

    for nombre_hoja in HOJAS_AUTOMATICAS:
        try:
            sincronizar_hoja_automatica(engine, nombre_hoja, hojas.get(nombre_hoja))
        except Exception as e:
            log(f"ERROR sincronizando {nombre_hoja}: {e}")

    log("Sincronización con Postgres completada.")


if __name__ == "__main__":
    main()
