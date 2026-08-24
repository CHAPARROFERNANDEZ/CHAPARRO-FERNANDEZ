"""
postgres_writer.py — Capa de ESCRITURA en paralelo hacia Postgres, "doble escritura".

Fase: "convivencia con escritura en paralelo". A diferencia de init_db.py (que sincroniza una
vez por deploy), este módulo se llama EN VIVO cada vez que la app guarda algo en Drive, para que
Postgres vaya quedando al día casi en tiempo real — sin esperar al próximo despliegue.

Regla de oro, la misma que ya rige en toda la app para las escrituras a Drive: esto NUNCA debe
romper ni bloquear el guardado real (que sigue siendo el de Drive/Excel, la única fuente de
verdad mientras dure esta fase). Cualquier fallo aquí se traga en silencio y se loguea a stderr
— la operación del usuario ya se guardó bien en Drive independientemente de lo que pase aquí.

Se puede desactivar por completo sin tocar código, poniendo la variable de entorno
POSTGRES_DUAL_WRITE=off en Railway. Por defecto está activada (best-effort) si hay DATABASE_URL.

Usa exactamente la misma convención de nombres de tabla que init_db.py (excel_<hoja_en_minuscula>
para las hojas normales, y la tabla dedicada `usuarios` para USUARIOS), para que
postgres_reader.py pueda leerlas sin ningún cambio.
"""
import os
import sys

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

_ENGINE: "Engine | None" = None

# Debe coincidir con HOJAS_AUTOMATICAS de init_db.py — mismas hojas, mismo criterio.
HOJAS_AUTOMATICAS = [
    "INVERSIONES", "REINVERSIONES", "CONTROL_NOTAS", "CALENDARIO_NOTAS",
    "CALENDARIO_CALLS", "DEUDA_JORDI", "TRANSFERENCIAS_JORDI",
    "MOVIMIENTOS_MOTOCLICK", "REPARTO_DIVIDENDOS", "BORRADORES_NOTAS",
    "BORRADORES_INVERSIONES", "AUDITORIA_NOTAS", "LOG_IA_USO",
]


def _log(msg: str) -> None:
    print(f"[postgres_writer] {msg}", file=sys.stderr, flush=True)


def _dual_write_activo() -> bool:
    if os.environ.get("POSTGRES_DUAL_WRITE", "on").strip().lower() in ("off", "false", "0", "no"):
        return False
    return bool(os.environ.get("DATABASE_URL", ""))


def _get_engine() -> Engine:
    global _ENGINE
    if _ENGINE is not None:
        return _ENGINE
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        raise RuntimeError("No hay DATABASE_URL en el entorno.")
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif url.startswith("postgresql://") and "+psycopg2" not in url:
        url = url.replace("postgresql://", "postgresql+psycopg2://", 1)
    _ENGINE = create_engine(url, pool_pre_ping=True, pool_size=3, max_overflow=2)
    return _ENGINE


def _normalizar_columnas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip().lower().replace(" ", "_").replace("ñ", "n")
        for c in df.columns
    ]
    df = df.loc[:, [c for c in df.columns if not c.startswith("unnamed")]]
    return df


def escribir_hojas_postgres(hojas: dict) -> None:
    """Best-effort: replica en Postgres cada hoja del dict 'hojas' que ya se acaba de guardar
    en Drive. Nunca lanza excepción hacia arriba — quien llama no necesita ni comprobar el
    resultado. Si Postgres está caído, o DATABASE_URL no está configurada, o falla cualquier
    hoja concreta, simplemente se omite y queda constancia en los logs de Railway."""
    if not _dual_write_activo():
        return
    try:
        engine = _get_engine()
    except Exception as e:
        _log(f"AVISO: no se pudo preparar la conexión a Postgres, se omite esta escritura en paralelo: {e}")
        return

    hojas_ok = []
    for nombre_hoja in HOJAS_AUTOMATICAS:
        df = hojas.get(nombre_hoja)
        if df is None or df.empty:
            continue
        try:
            tabla = f"excel_{nombre_hoja.lower()}"
            df_norm = _normalizar_columnas(df)
            for col in df_norm.columns:
                if pd.api.types.is_datetime64_any_dtype(df_norm[col]):
                    df_norm[col] = df_norm[col].dt.strftime("%Y-%m-%d %H:%M:%S")
            df_norm.to_sql(tabla, engine, if_exists="replace", index=False, method="multi", chunksize=500)
            hojas_ok.append(f"{nombre_hoja}({len(df_norm)})")
        except Exception as e:
            _log(f"AVISO: fallo replicando '{nombre_hoja}' a Postgres (Drive ya quedó guardado bien, no se pierde nada): {e}")

    if hojas_ok:
        _log(f"OK — escritura en paralelo confirmada: {', '.join(hojas_ok)}")


def sincronizar_usuarios_postgres(df_usuarios: pd.DataFrame) -> None:
    """Best-effort: mismo upsert que init_db.sincronizar_usuarios(), pero llamado en vivo cada
    vez que se guarda la hoja USUARIOS (cambio de contraseña, alta de 2FA, etc.)."""
    if not _dual_write_activo():
        return
    if df_usuarios is None or df_usuarios.empty:
        return
    try:
        engine = _get_engine()
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
        df = _normalizar_columnas(df_usuarios)
        columnas_esperadas = {"usuario", "tipo_usuario", "password", "debe_cambiar_password", "totp_secret", "totp_activo"}
        if not columnas_esperadas.issubset(df.columns):
            _log(f"AVISO: USUARIOS no tiene las columnas esperadas ({columnas_esperadas - set(df.columns)}), se omite esta escritura en paralelo.")
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
    except Exception as e:
        _log(f"AVISO: fallo replicando USUARIOS a Postgres (Drive ya quedó guardado bien, no se pierde nada): {e}")
