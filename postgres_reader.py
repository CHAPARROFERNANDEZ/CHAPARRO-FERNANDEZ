"""
postgres_reader.py — Capa de LECTURA desde Postgres, alternativa a leer el Excel de Drive.

Fase: "espejo activable". Este módulo nunca se importa ni se conecta a nada salvo que la
variable de entorno DATA_SOURCE valga "postgres" (ver app.py). Con DATA_SOURCE sin definir o en
"drive" (default), este archivo entero puede no existir y la app funciona exactamente igual que
antes de esta migración.

Las tablas que lee (excel_<hoja_en_minuscula>) las crea y llena init_db.py, que corre antes de
cada deploy sincronizando desde el Excel de Drive. Este archivo NO escribe nada — solo lee.

Si cualquier lectura de aquí falla (Postgres caído, tabla vacía, columna que cambió, etc.),
lanza la excepción hacia arriba; quien llama (app.py) es responsable de capturarla y caer de
vuelta a Drive. Esa es la red de seguridad real: nunca queda la app sin datos.
"""
import os

import pandas as pd
from sqlalchemy import create_engine, text
from sqlalchemy.engine import Engine

_ENGINE: "Engine | None" = None


def _get_engine() -> Engine:
    global _ENGINE
    if _ENGINE is not None:
        return _ENGINE
    url = os.environ.get("DATABASE_URL", "")
    if not url:
        raise RuntimeError("DATA_SOURCE=postgres pero no hay DATABASE_URL en el entorno.")
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif url.startswith("postgresql://") and "+psycopg2" not in url:
        url = url.replace("postgresql://", "postgresql+psycopg2://", 1)
    _ENGINE = create_engine(url, pool_pre_ping=True, pool_size=3, max_overflow=2)
    return _ENGINE


def _leer_tabla(nombre_hoja: str) -> pd.DataFrame:
    tabla = f"excel_{nombre_hoja.lower()}"
    engine = _get_engine()
    with engine.connect() as conn:
        existe = conn.execute(text(
            "SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = :t)"
        ), {"t": tabla}).scalar()
        if not existe:
            raise RuntimeError(f"La tabla '{tabla}' todavía no existe en Postgres (¿corrió init_db.py alguna vez?).")
        df = pd.read_sql(text(f'SELECT * FROM "{tabla}"'), conn)
    if "index" in df.columns:
        df = df.drop(columns=["index"])
    return df


def leer_hoja_excel_postgres(nombre_hoja: str) -> pd.DataFrame:
    """Equivalente a leer_hoja_excel() de app.py pero desde Postgres.
    Las columnas ya salen en minúsculas porque init_db.py las normaliza al sincronizar."""
    return _leer_tabla(nombre_hoja)


def cargar_excel_completo_postgres():
    """Equivalente a cargar_excel_completo() de app.py pero desde Postgres.
    Devuelve (inv, cal, control) con el MISMO tipado/normalización que la versión de Drive,
    para que el resto del código de app.py no note ninguna diferencia."""
    inv = _leer_tabla("INVERSIONES")
    cal = _leer_tabla("CALENDARIO_NOTAS")
    try:
        control = _leer_tabla("CONTROL_NOTAS")
    except Exception:
        control = pd.DataFrame()

    if "unnamed:_6" in inv.columns and "cuenta_cobro" not in inv.columns:
        inv = inv.rename(columns={"unnamed:_6": "cuenta_cobro"})

    for col in ["id_inversion", "inversor", "tipo_inversion", "subtipo_inversion", "nombre_activo",
                "metodo_calculo", "activo_generador_interes", "tipo_operacion", "capital_nuevo_real",
                "cuenta_cobro", "motivo"]:
        if col in inv.columns:
            inv[col] = inv[col].fillna("").astype(str).str.strip()

    for col in ["fecha_inversion", "fecha_final_inversion"]:
        if col in inv.columns:
            inv[col] = pd.to_datetime(inv[col], errors="coerce")

    for col in ["capital_invertido", "interes_inversor_anual", "interes_nota_anual"]:
        if col in inv.columns:
            inv[col] = pd.to_numeric(inv[col], errors="coerce").fillna(0)
        else:
            inv[col] = 0

    if "periodicidad_meses" in inv.columns:
        inv["periodicidad_meses"] = pd.to_numeric(inv["periodicidad_meses"], errors="coerce").fillna(1).astype(int)
    else:
        inv["periodicidad_meses"] = 1

    if "nota" in cal.columns:
        cal["nota"] = pd.to_numeric(cal["nota"], errors="coerce").astype("Int64")
    if "tipo_evento" in cal.columns:
        cal["tipo_evento"] = cal["tipo_evento"].fillna("").astype(str).str.strip().str.upper()
    if "fecha" in cal.columns:
        cal["fecha"] = pd.to_datetime(cal["fecha"], errors="coerce").dt.normalize()

    if not control.empty:
        if "nota" in control.columns:
            control["nota"] = pd.to_numeric(control["nota"], errors="coerce").astype("Int64")
        if "ticker" in control.columns:
            control["ticker"] = control["ticker"].fillna("").astype(str).str.strip().str.upper()
        for col in ["precio_compra", "barrera_cupon", "contingency", "barrera_capital"]:
            if col in control.columns:
                control[col] = pd.to_numeric(control[col], errors="coerce")

    return inv, cal, control


def leer_pdf_nota_postgres(numero_nota: int) -> "bytes | None":
    """Lee el PDF de una nota guardado en la tabla pdfs_notas, si existe. A diferencia del
    resto de este módulo, nunca lanza excepción hacia arriba: devuelve None ante cualquier
    problema (tabla inexistente, fila inexistente, Postgres caído), porque quien la llama
    (leer_pdf_nota_guardado en app.py) ya tiene su propia cadena de respaldo (volumen → Drive)
    y no necesita tratar esto como un fallo grave."""
    try:
        engine = _get_engine()
        with engine.connect() as conn:
            existe = conn.execute(text(
                "SELECT EXISTS (SELECT 1 FROM information_schema.tables WHERE table_name = 'pdfs_notas')"
            )).scalar()
            if not existe:
                return None
            fila = conn.execute(text(
                "SELECT pdf_data FROM pdfs_notas WHERE numero_nota = :n"
            ), {"n": int(numero_nota)}).fetchone()
            return bytes(fila[0]) if fila else None
    except Exception:
        return None
