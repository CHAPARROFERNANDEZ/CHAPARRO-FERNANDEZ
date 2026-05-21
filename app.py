import calendar
import re
import zipfile
from datetime import datetime
from io import BytesIO
from typing import Optional

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
except Exception:
    px = None

try:
    import yfinance as yf
except Exception:
    yf = None

import requests

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(
    page_title="Sistema Fondo",
    layout="wide",
    initial_sidebar_state="expanded",
)

ARCHIVO = "inversiones.xlsx"
GDRIVE_FILE_ID = "1CImiIbg7kSLrYNpWgzPHEBCmI3KRVlBX"
HOJA_INVERSIONES = "INVERSIONES"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"
HOJA_CONTROL = "CONTROL_NOTAS"

TASA_ANUAL_FUTBOL = 0.15
TASA_ANUAL_MOTOCLICK = 0.25
TASA_ANUAL_PARAGUAY = 0.15


# =========================
# ESTILO PROFESIONAL
# =========================
def aplicar_estilo_profesional():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
        html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
        .stApp {
            background:
                radial-gradient(circle at top left, rgba(191, 154, 95, 0.18), transparent 30%),
                linear-gradient(135deg, #071425 0%, #0e2338 42%, #f6f3ee 42%, #f6f3ee 100%);
            background-attachment: fixed;
        }
        .block-container {
            max-width: 1240px;
            padding-top: 2.2rem;
            padding-bottom: 3rem;
            background: rgba(255, 255, 255, 0.94);
            border: 1px solid rgba(255, 255, 255, 0.55);
            border-radius: 28px;
            box-shadow: 0 24px 70px rgba(4, 20, 37, 0.18);
            margin-top: 1.2rem;
            margin-bottom: 1.2rem;
        }
        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg, #071425 0%, #102a43 100%);
            border-right: 1px solid rgba(191, 154, 95, 0.32);
            transform: none !important;
            visibility: visible !important;
        }
        section[data-testid="stSidebar"] label,
        section[data-testid="stSidebar"] p,
        section[data-testid="stSidebar"] span,
        section[data-testid="stSidebar"] div { color: #f7f1e8 !important; }
        section[data-testid="stSidebar"] [data-baseweb="select"] div,
        section[data-testid="stSidebar"] input { color: #071425 !important; }
        h1, h2, h3 { color: #102033; letter-spacing: -0.03em; }
        div[data-testid="stMetric"] {
            background: linear-gradient(145deg, #ffffff 0%, #f5f0e8 100%);
            border: 1px solid rgba(191, 154, 95, 0.24);
            border-radius: 20px;
            padding: 18px 20px;
            box-shadow: 0 10px 28px rgba(15, 35, 55, 0.08);
        }
        div[data-testid="stMetricValue"] { color: #0e2338; font-weight: 800; }
        .stButton > button, .stDownloadButton > button, button[kind="primary"] {
            border-radius: 999px !important;
            background: linear-gradient(135deg, #0e2338 0%, #173b5c 60%, #bf9a5f 100%) !important;
            color: white !important;
            border: 0 !important;
            font-weight: 700 !important;
            padding: 0.55rem 1.2rem !important;
            box-shadow: 0 10px 24px rgba(14, 35, 56, 0.20);
        }
        div[data-testid="stDataFrame"] {
            border-radius: 18px;
            overflow: hidden;
            border: 1px solid rgba(14, 35, 56, 0.10);
            box-shadow: 0 10px 28px rgba(15, 35, 55, 0.06);
        }
        .brand-hero {
            display: flex; align-items: center; justify-content: space-between; gap: 24px;
            padding: 26px 30px; margin-bottom: 26px; border-radius: 26px;
            background: linear-gradient(135deg, rgba(14,35,56,0.97) 0%, rgba(20,53,82,0.94) 58%, rgba(191,154,95,0.92) 100%);
            box-shadow: 0 18px 45px rgba(7, 20, 37, 0.22); color: white;
        }
        .brand-left { display: flex; align-items: center; gap: 18px; }
        .brand-logo {
            width: 76px; height: 76px; border-radius: 20px; background: rgba(255,255,255,0.94); color: #0e2338;
            display: flex; align-items: center; justify-content: center; font-size: 34px; font-weight: 800;
            letter-spacing: -0.10em; border: 1px solid rgba(191,154,95,0.45);
            box-shadow: inset 0 0 0 2px rgba(191,154,95,0.12), 0 10px 24px rgba(0,0,0,0.16);
        }
        .brand-title { font-size: 30px; line-height: 1.05; font-weight: 800; letter-spacing: -0.04em; }
        .brand-subtitle { margin-top: 8px; color: rgba(255,255,255,0.78); font-size: 14px; font-weight: 500; }
        .brand-tag {
            padding: 9px 14px; border-radius: 999px; background: rgba(255,255,255,0.13);
            border: 1px solid rgba(255,255,255,0.25); font-size: 13px; font-weight: 700; color: #fff; white-space: nowrap;
        }
        .login-card {
            max-width: 460px; margin: 6vh auto 0 auto; padding: 34px 34px 30px 34px; border-radius: 30px;
            background: rgba(255,255,255,0.94); border: 1px solid rgba(191,154,95,0.30);
            box-shadow: 0 26px 80px rgba(4, 20, 37, 0.32); text-align: center;
        }
        .login-logo {
            width: 96px; height: 96px; border-radius: 26px; margin: 0 auto 18px auto;
            background: linear-gradient(145deg, #ffffff, #f3eadc); color: #0e2338;
            display: flex; align-items: center; justify-content: center; font-size: 42px; font-weight: 800;
            letter-spacing: -0.10em; border: 1px solid rgba(191,154,95,0.45);
        }
        .login-title { font-size: 26px; font-weight: 800; color: #0e2338; letter-spacing: -0.03em; margin-bottom: 6px; }
        .login-subtitle { font-size: 14px; color: #667085; margin-bottom: 20px; }
        #MainMenu, footer {visibility: hidden;}
        header {visibility: visible;}
        </style>
        """,
        unsafe_allow_html=True,
    )


def mostrar_hero(usuario=None):
    tag = f"Sesión: {usuario}" if usuario else "Private Wealth Dashboard"
    st.markdown(
        f"""
        <div class="brand-hero">
            <div class="brand-left">
                <div class="brand-logo">CF</div>
                <div>
                    <div class="brand-title">Chaparro Fernández Wealth</div>
                    <div class="brand-subtitle">Sistema privado de control, inversiones, notas, alertas y extractos</div>
                </div>
            </div>
            <div class="brand-tag">{tag}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


aplicar_estilo_profesional()


# =========================
# LOGIN
# =========================
USUARIOS = {"Yuri": "1234", "Jordi": "12345", "Alan": "123456"}

if "autenticado" not in st.session_state:
    st.session_state.autenticado = False
if "usuario" not in st.session_state:
    st.session_state.usuario = None

if not st.session_state.autenticado:
    st.markdown(
        """
        <div class="login-card">
            <div class="login-logo">CF</div>
            <div class="login-title">Chaparro Fernández Wealth</div>
            <div class="login-subtitle">Acceso privado al sistema financiero interno</div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("login_form"):
        usuario = st.selectbox("Usuario", list(USUARIOS.keys()))
        password = st.text_input("Contraseña", type="password")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        if USUARIOS.get(usuario) == password:
            st.session_state.autenticado = True
            st.session_state.usuario = usuario
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos")
    st.stop()

st.sidebar.markdown(f"**Usuario conectado:** {st.session_state.usuario}")
st.sidebar.caption("Si el menú se oculta, recarga la página: ahora se abrirá automáticamente.")
if st.sidebar.button("Cerrar sesión"):
    st.session_state.autenticado = False
    st.session_state.usuario = None
    st.rerun()


# =========================
# UTILIDADES Y CARGA
# =========================
def fmt(x):
    try:
        return f"${float(x):,.2f}"
    except Exception:
        return "$0.00"


def fmt_pct(x):
    try:
        if pd.isna(x):
            return "0.00%"
        return f"{float(x) * 100:.2f}%"
    except Exception:
        return "0.00%"


def nombre_mes_es(mes: int) -> str:
    meses = {1: "enero", 2: "febrero", 3: "marzo", 4: "abril", 5: "mayo", 6: "junio", 7: "julio", 8: "agosto", 9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"}
    return meses.get(int(mes), str(mes))


def ultimo_dia_mes(anio: int, mes: int) -> int:
    return calendar.monthrange(int(anio), int(mes))[1]


def limpiar_texto(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().lower()


def es_chaparro_fernandez_row(row) -> bool:
    """Detecta SOLO las inversiones internas de la sociedad Chaparro Fernández.

    IMPORTANTE:
    No se debe marcar como Chaparro Fernández a personas que se apellidan
    Chaparro o Fernández, como JORDI CHAPARRO, YURI FERNANDEZ, EVA CHAPARRO
    o PAOLA CHAPARRO. Por eso la detección se hace únicamente sobre la columna
    inversor y exige el nombre completo de la sociedad.
    """
    inversor = limpiar_texto(row.get("inversor", ""))
    inversor = (
        inversor.replace("á", "a")
        .replace("é", "e")
        .replace("í", "i")
        .replace("ó", "o")
        .replace("ú", "u")
        .replace("ñ", "n")
        .replace("-", " ")
    )
    inversor = " ".join(inversor.split())

    nombres_sociedad = {
        "chaparro fernandez",
        "chaparro fernandez sl",
        "chaparro fernandez s.l.",
        "chaparro fernandez sociedad",
        "chaparro fernandez wealth",
    }
    return inversor in nombres_sociedad


def aplicar_filtro_chaparro_fernandez(df: pd.DataFrame, incluir_chaparro: bool) -> pd.DataFrame:
    """Incluye o excluye las inversiones internas de Chaparro Fernández.

    Si incluir_chaparro=True, no toca el dataframe.
    Si incluir_chaparro=False, elimina filas detectadas como Chaparro Fernández.
    Además añade una columna auxiliar es_chaparro_fernandez para poder auditarlo.
    """
    if df is None or df.empty:
        return df
    out = df.copy()
    out["es_chaparro_fernandez"] = out.apply(es_chaparro_fernandez_row, axis=1)
    if incluir_chaparro:
        return out
    return out[~out["es_chaparro_fernandez"]].copy()


def descargar_excel_desde_drive():
    """Descarga el Excel desde Google Sheets y lo guarda localmente."""
    try:
        url = f"https://docs.google.com/spreadsheets/d/{GDRIVE_FILE_ID}/export?format=xlsx"
        response = requests.get(url, timeout=30)
        if response.status_code == 200:
            with open(ARCHIVO, "wb") as f:
                f.write(response.content)
            return True
        else:
            st.warning(f"No se pudo descargar el Excel desde Google Drive (status {response.status_code}).")
            return False
    except Exception as e:
        st.warning(f"No se pudo descargar el Excel desde Google Drive: {e}")
        return False


@st.cache_data(show_spinner=False)
def cargar_excel_completo():
    inv = pd.read_excel(ARCHIVO, sheet_name=HOJA_INVERSIONES)
    cal = pd.read_excel(ARCHIVO, sheet_name=HOJA_CALENDARIO)
    try:
        control = pd.read_excel(ARCHIVO, sheet_name=HOJA_CONTROL)
    except Exception:
        control = pd.DataFrame()

    inv.columns = [str(c).strip().lower() for c in inv.columns]
    cal.columns = [str(c).strip().lower() for c in cal.columns]
    control.columns = [str(c).strip().lower() for c in control.columns]

    if "unnamed: 6" in inv.columns and "cuenta_cobro" not in inv.columns:
        inv = inv.rename(columns={"unnamed: 6": "cuenta_cobro"})

    for col in ["id_inversion", "inversor", "tipo_inversion", "subtipo_inversion", "nombre_activo", "metodo_calculo", "activo_generador_interes", "tipo_operacion", "capital_nuevo_real", "cuenta_cobro", "motivo"]:
        if col in inv.columns:
            inv[col] = inv[col].fillna("").astype(str).str.strip()

    for col in ["fecha_inversion", "fecha_final_inversion"]:
        if col in inv.columns:
            inv[col] = pd.to_datetime(inv[col], errors="coerce", dayfirst=True)

    for col in ["capital_invertido", "interes_inversor_anual", "interes_nota_anual"]:
        if col in inv.columns:
            inv[col] = pd.to_numeric(inv[col], errors="coerce").fillna(0)
        else:
            inv[col] = 0

    if "nota" in cal.columns:
        cal["nota"] = pd.to_numeric(cal["nota"], errors="coerce").astype("Int64")
    if "tipo_evento" in cal.columns:
        cal["tipo_evento"] = cal["tipo_evento"].fillna("").astype(str).str.strip().str.upper()
    if "fecha" in cal.columns:
        cal["fecha"] = pd.to_datetime(cal["fecha"], errors="coerce", dayfirst=True).dt.normalize()

    if not control.empty:
        if "nota" in control.columns:
            control["nota"] = pd.to_numeric(control["nota"], errors="coerce").astype("Int64")
        if "ticker" in control.columns:
            control["ticker"] = control["ticker"].fillna("").astype(str).str.strip().str.upper()
        for col in ["precio_compra", "barrera_cupon", "contingency", "barrera_capital"]:
            if col in control.columns:
                control[col] = pd.to_numeric(control[col], errors="coerce")
    return inv, cal, control


def leer_hoja_excel(nombre_hoja: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(ARCHIVO, sheet_name=nombre_hoja)
        df.columns = [str(c).strip().lower() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def preparar_tabla_monetaria(df: pd.DataFrame, columnas_monetarias) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if "fecha" in col:
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d/%m/%Y")
    for col in columnas_monetarias:
        if col in out.columns:
            out[col] = out[col].map(fmt)
    return out


def mostrar_metricas(titulo, valores):
    st.subheader(titulo)
    cols = st.columns(len(valores))
    for col, (label, value) in zip(cols, valores):
        col.metric(label, value)


# =========================
# CÁLCULOS ACTIVOS
# =========================
def filtrar_activo(df_base: pd.DataFrame, activo: str) -> pd.DataFrame:
    activo_l = activo.lower()
    subtipo = df_base.get("subtipo_inversion", pd.Series(index=df_base.index, dtype=str)).astype(str).str.lower()
    nombre = df_base.get("nombre_activo", pd.Series(index=df_base.index, dtype=str)).astype(str).str.lower()
    if activo_l == "futbol":
        return df_base[subtipo.isin(["futbol", "fútbol"]) | nombre.isin(["futbol", "fútbol"])].copy()
    return df_base[subtipo.eq(activo_l) | nombre.eq(activo_l)].copy()


def dias_activos_en_mes(fecha_inicio, fecha_fin, anio: int, mes: int) -> int:
    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))
    if pd.isna(fecha_inicio) or fecha_inicio > fin_mes:
        return 0
    if pd.notna(fecha_fin) and fecha_fin < inicio_mes:
        return 0
    inicio_real = max(fecha_inicio, inicio_mes)
    fin_real = fin_mes if pd.isna(fecha_fin) else min(fecha_fin, fin_mes)
    if inicio_real > fin_real:
        return 0
    return (fin_real - inicio_real).days + 1


def detalle_activo_mes(df_base: pd.DataFrame, activo: str, tasa_anual: float, anio: int, mes: int) -> pd.DataFrame:
    df_activo = filtrar_activo(df_base, activo)
    dias_mes = ultimo_dia_mes(anio, mes)
    filas = []
    for _, fila in df_activo.iterrows():
        dias = dias_activos_en_mes(fila.get("fecha_inversion"), fila.get("fecha_final_inversion"), anio, mes)
        if dias == 0:
            continue
        proporcion = dias / dias_mes
        capital = float(fila.get("capital_invertido", 0))
        ingreso_bruto = capital * tasa_anual / 12 * proporcion
        pago_inversor = capital * float(fila.get("interes_inversor_anual", 0)) / 12 * proporcion
        filas.append({
            "id_inversion": fila.get("id_inversion", ""), "inversor": fila.get("inversor", ""),
            "capital_invertido": capital, "fecha_inversion": fila.get("fecha_inversion"),
            "fecha_final_inversion": fila.get("fecha_final_inversion"), "dias_activos": dias,
            "dias_mes": dias_mes, "ingreso_bruto": ingreso_bruto,
            "pago_inversor_mes": pago_inversor, "beneficio_empresa_mes": ingreso_bruto - pago_inversor,
        })
    return pd.DataFrame(filas)


def capital_activo_en_fecha(df_base: pd.DataFrame, fecha_consulta, activo: Optional[str] = None, solo_real: bool = False) -> float:
    fecha_consulta = pd.Timestamp(fecha_consulta).normalize()
    trabajo = df_base.copy()
    if activo:
        trabajo = filtrar_activo(trabajo, activo)
    if solo_real and "capital_nuevo_real" in trabajo.columns:
        trabajo = trabajo[trabajo["capital_nuevo_real"].astype(str).str.lower() == "si"].copy()
    filtrado = trabajo[(trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= fecha_consulta) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha_consulta))]
    return float(filtrado["capital_invertido"].sum()) if not filtrado.empty else 0.0


def total_pagado_activo_desde_inicio(df_base: pd.DataFrame, activo: str, tasa_anual: float) -> float:
    df_activo = filtrar_activo(df_base, activo)
    if df_activo.empty:
        return 0.0
    fecha_min = df_activo["fecha_inversion"].dropna().min()
    if pd.isna(fecha_min):
        return 0.0
    hoy = pd.Timestamp.today().normalize()
    total = 0.0
    anio, mes = fecha_min.year, fecha_min.month
    while (anio < hoy.year) or (anio == hoy.year and mes <= hoy.month):
        detalle = detalle_activo_mes(df_base, activo, tasa_anual, anio, mes)
        if not detalle.empty:
            total += detalle["pago_inversor_mes"].sum()
        mes += 1
        if mes == 13:
            mes = 1
            anio += 1
    return float(total)


def total_ingresado_activo_desde_inicio(df_base: pd.DataFrame, activo: str, tasa_anual: float) -> float:
    df_activo = filtrar_activo(df_base, activo)
    if df_activo.empty:
        return 0.0
    fecha_min = df_activo["fecha_inversion"].dropna().min()
    if pd.isna(fecha_min):
        return 0.0
    hoy = pd.Timestamp.today().normalize()
    total = 0.0
    anio, mes = fecha_min.year, fecha_min.month
    while (anio < hoy.year) or (anio == hoy.year and mes <= hoy.month):
        detalle = detalle_activo_mes(df_base, activo, tasa_anual, anio, mes)
        if not detalle.empty:
            total += detalle["ingreso_bruto"].sum()
        mes += 1
        if mes == 13:
            mes = 1
            anio += 1
    return float(total)


# =========================
# NOTAS
# =========================
def normalizar_cuenta(valor):
    texto = str(valor).strip().lower()
    if texto in ["jordi", "cuenta jordi"]:
        return "JORDI"
    if texto in ["compañia", "compania", "empresa", "sociedad"]:
        return "COMPAÑÍA"
    return "SIN CLASIFICAR"


def extraer_numero_nota(nombre_activo):
    if pd.isna(nombre_activo):
        return pd.NA
    match = re.search(r"NOTA[_\s]?(\d+)", str(nombre_activo).strip().upper())
    return int(match.group(1)) if match else pd.NA


def filtrar_notas(df_base: pd.DataFrame) -> pd.DataFrame:
    trabajo = df_base.copy()
    if "tipo_inversion" in trabajo.columns:
        trabajo = trabajo[trabajo["tipo_inversion"].astype(str).str.lower() == "nota"].copy()
    if "nombre_activo" not in trabajo.columns:
        trabajo["nombre_activo"] = ""
    trabajo["nota_num"] = trabajo["nombre_activo"].apply(extraer_numero_nota)
    trabajo["nota_num"] = pd.to_numeric(trabajo["nota_num"], errors="coerce").astype("Int64")
    if "activo_generador_interes" in trabajo.columns:
        trabajo = trabajo[trabajo["activo_generador_interes"].astype(str).str.upper() == "SI"].copy()
    if "cuenta_cobro" not in trabajo.columns:
        trabajo["cuenta_cobro"] = "SIN CLASIFICAR"
    trabajo["cuenta_cobro"] = trabajo["cuenta_cobro"].apply(normalizar_cuenta)
    return trabajo


def inversiones_activas_para_nota(df_base: pd.DataFrame, nota: int, fecha_pago) -> pd.DataFrame:
    fecha_pago = pd.Timestamp(fecha_pago).normalize()
    trabajo = filtrar_notas(df_base)
    return trabajo[(trabajo["nota_num"] == nota) & (trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= fecha_pago) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha_pago))].copy()


def pagos_notas_mes(df_cal: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    if df_cal.empty:
        return pd.DataFrame()
    return df_cal[(df_cal["tipo_evento"] == "PAGO") & (df_cal["fecha"].notna()) & (df_cal["fecha"].dt.year == anio) & (df_cal["fecha"].dt.month == mes)].copy().sort_values(["fecha", "nota"])


def pagos_notas_hasta_hoy(df_cal: pd.DataFrame) -> pd.DataFrame:
    hoy = pd.Timestamp.today().normalize()
    return df_cal[(df_cal["tipo_evento"] == "PAGO") & (df_cal["fecha"].notna()) & (df_cal["fecha"] <= hoy)].copy().sort_values(["fecha", "nota"])


def obtener_observacion_previa_nota(df_cal: pd.DataFrame, nota: int, fecha_pago):
    if df_cal is None or df_cal.empty:
        return None
    fecha_pago = pd.Timestamp(fecha_pago).normalize()
    obs = df_cal[(df_cal["nota"] == nota) & (df_cal["tipo_evento"] == "OBSERVACION") & (df_cal["fecha"].notna()) & (df_cal["fecha"] <= fecha_pago)].copy().sort_values("fecha")
    return None if obs.empty else obs.iloc[-1]["fecha"]


def normalizar_barrera(valor):
    if pd.isna(valor):
        return None
    try:
        valor = float(valor)
        return valor / 100 if valor > 1 else valor
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=3600)
def obtener_cierre_ticker_fecha(ticker: str, fecha_objetivo):
    if yf is None:
        return None
    try:
        fecha_objetivo = pd.Timestamp(fecha_objetivo).normalize()
        inicio = fecha_objetivo - pd.Timedelta(days=10)
        fin = fecha_objetivo + pd.Timedelta(days=2)
        data = yf.download(str(ticker).strip().upper(), start=inicio.strftime("%Y-%m-%d"), end=fin.strftime("%Y-%m-%d"), progress=False, auto_adjust=False)
        if data.empty:
            return None
        if isinstance(data.columns, pd.MultiIndex):
            if "Close" in data.columns.get_level_values(0):
                cierres = data["Close"].iloc[:, 0].dropna()
            else:
                return None
        else:
            if "Close" not in data.columns:
                return None
            cierres = data["Close"].dropna()
        cierres.index = pd.to_datetime(cierres.index).normalize()
        cierres = cierres[cierres.index <= fecha_objetivo]
        if cierres.empty:
            return None
        return float(cierres.iloc[-1])
    except Exception:
        return None


def columna_barrera_control(df_control: pd.DataFrame, preferida="contingency"):
    for col in [preferida, "contingency", "barrera_cupon", "barrera_capital"]:
        if col in df_control.columns:
            return col
    return None


def evaluar_nota_en_fecha(df_control: pd.DataFrame, nota: int, fecha_obs, preferida="contingency") -> tuple[str, pd.DataFrame]:
    """
    Evalúa una nota en una fecha de observación.

    Lógica definitiva:
    - Si la fecha de observación es futura: PENDIENTE y cuenta como cobro previsto.
    - Si ya llegó la observación y no hay dato de precio: SIN DATO, pero NO se trata como negativa.
    - Solo es NEGATIVA si ya llegó la observación y existe precio real por debajo de la barrera.
    """
    hoy = pd.Timestamp.today().normalize()

    if fecha_obs is None or pd.isna(fecha_obs):
        return "SIN_OBSERVACION", pd.DataFrame()

    fecha_obs = pd.Timestamp(fecha_obs).normalize()

    if df_control is None or df_control.empty:
        return "SIN_CONTROL", pd.DataFrame()

    sub = df_control[df_control.get("nota") == nota].copy()
    if sub.empty:
        return "SIN_CONTROL", pd.DataFrame()

    barrera_col = columna_barrera_control(sub, preferida=preferida)
    if barrera_col is None or "precio_compra" not in sub.columns or "ticker" not in sub.columns:
        return "SIN_COLUMNAS", pd.DataFrame()

    filas = []

    # Si la observación aún no ha llegado, no se descarga precio histórico.
    # Se marca como pendiente y se mantiene el cobro previsto.
    if fecha_obs > hoy:
        for _, row in sub.iterrows():
            ticker = row.get("ticker", "")
            compra = pd.to_numeric(row.get("precio_compra"), errors="coerce")
            barrera_pct = normalizar_barrera(row.get(barrera_col))
            precio_barrera = float(compra) * float(barrera_pct) if pd.notna(compra) and barrera_pct is not None else None
            filas.append({
                "ticker": ticker,
                "precio_compra": float(compra) if pd.notna(compra) else None,
                "barrera_%": barrera_pct,
                "precio_barrera": precio_barrera,
                "cierre_usado": None,
                "estado": "PENDIENTE",
            })
        return "PENDIENTE", pd.DataFrame(filas)

    hay_negativa_real = False
    hay_sin_dato = False

    for _, row in sub.iterrows():
        ticker = row.get("ticker", "")
        compra = pd.to_numeric(row.get("precio_compra"), errors="coerce")
        barrera_pct = normalizar_barrera(row.get(barrera_col))

        if pd.isna(compra) or barrera_pct is None:
            filas.append({
                "ticker": ticker,
                "precio_compra": float(compra) if pd.notna(compra) else None,
                "barrera_%": barrera_pct,
                "precio_barrera": None,
                "cierre_usado": None,
                "estado": "FALTAN DATOS",
            })
            hay_sin_dato = True
            continue

        precio_barrera = float(compra) * float(barrera_pct)
        cierre = obtener_cierre_ticker_fecha(ticker, fecha_obs)

        if cierre is None:
            filas.append({
                "ticker": ticker,
                "precio_compra": float(compra),
                "barrera_%": barrera_pct,
                "precio_barrera": precio_barrera,
                "cierre_usado": None,
                "estado": "SIN DATO",
            })
            hay_sin_dato = True
            continue

        estado = "OK" if cierre >= precio_barrera else "NO OK"
        if estado == "NO OK":
            hay_negativa_real = True

        filas.append({
            "ticker": ticker,
            "precio_compra": float(compra),
            "barrera_%": barrera_pct,
            "precio_barrera": precio_barrera,
            "cierre_usado": cierre,
            "estado": estado,
        })

    detalle = pd.DataFrame(filas)

    if hay_negativa_real:
        return "NEGATIVA", detalle
    if hay_sin_dato:
        return "SIN DATO", detalle
    return "POSITIVA", detalle


def resumen_detalle_observacion(detalle_obs: pd.DataFrame) -> str:
    if detalle_obs is None or detalle_obs.empty:
        return ""
    partes = []
    for _, row in detalle_obs.iterrows():
        ticker = row.get("ticker", "")
        estado = row.get("estado", "")
        cierre = row.get("cierre_usado", None)
        barrera = row.get("precio_barrera", None)
        if pd.notna(cierre) and pd.notna(barrera):
            partes.append(f"{ticker}: {estado} cierre {float(cierre):.2f} / barrera {float(barrera):.2f}")
        else:
            partes.append(f"{ticker}: {estado}")
    return " | ".join(partes)


def preparar_detalle_notas(df_inv: pd.DataFrame, df_pagos: pd.DataFrame, df_cal: pd.DataFrame | None = None, df_control: pd.DataFrame | None = None) -> pd.DataFrame:
    filas = []
    cache_observaciones = {}

    for _, evento in df_pagos.iterrows():
        nota = evento.get("nota")
        fecha_pago = evento.get("fecha")

        if pd.isna(nota) or pd.isna(fecha_pago):
            continue

        nota_int = int(nota)
        fecha_pago = pd.Timestamp(fecha_pago).normalize()
        fecha_obs = obtener_observacion_previa_nota(df_cal, nota_int, fecha_pago) if df_cal is not None else None

        resultado_obs = "NO_EVALUADA"
        detalle_obs = pd.DataFrame()

        # Por defecto el cobro cuenta como previsto.
        # Solo se elimina si la observación es NEGATIVA real.
        ingreso_habilitado = True

        if fecha_obs is not None and df_control is not None and not df_control.empty:
            clave = (nota_int, pd.Timestamp(fecha_obs).normalize())
            if clave not in cache_observaciones:
                cache_observaciones[clave] = evaluar_nota_en_fecha(df_control, nota_int, fecha_obs, preferida="contingency")
            resultado_obs, detalle_obs = cache_observaciones[clave]

            if resultado_obs == "NEGATIVA":
                ingreso_habilitado = False

        activas = inversiones_activas_para_nota(df_inv, nota_int, fecha_pago)

        for _, fila in activas.iterrows():
            capital = float(fila.get("capital_invertido", 0))
            cobro_teorico = capital * float(fila.get("interes_nota_anual", 0)) / 12
            cobro_compania = cobro_teorico if ingreso_habilitado else 0.0

            # Tratamiento especial Chaparro Fernández:
            # - Si se excluye Chaparro, la fila ya no llega aquí porque se filtra antes.
            # - Si se incluye Chaparro, se considera una operación interna: el pago al
            #   inversionista debe ser igual al cobro de la nota y el beneficio debe ser 0.
            es_chaparro = bool(fila.get("es_chaparro_fernandez", False)) or es_chaparro_fernandez_row(fila)
            if es_chaparro:
                pago_inversor = cobro_compania
                beneficio_empresa = 0.0
                tratamiento_chaparro = "INTERNO: pago = cobro nota"
            else:
                pago_inversor = capital * float(fila.get("interes_inversor_anual", 0)) / 12
                beneficio_empresa = cobro_compania - pago_inversor
                tratamiento_chaparro = "NO"

            filas.append({
                "fecha_pago": fecha_pago,
                "nota": nota_int,
                "fecha_observacion_usada": fecha_obs,
                "resultado_observacion": resultado_obs,
                "detalle_observacion": resumen_detalle_observacion(detalle_obs),
                "ingreso_habilitado": "SI" if ingreso_habilitado else "NO",
                "id_inversion": fila.get("id_inversion", ""),
                "inversor": fila.get("inversor", ""),
                "cuenta_cobro": fila.get("cuenta_cobro", "SIN CLASIFICAR"),
                "es_chaparro_fernandez": es_chaparro,
                "tratamiento_chaparro": tratamiento_chaparro,
                "capital_invertido": capital,
                "interes_nota_anual": fila.get("interes_nota_anual", 0),
                "interes_inversor_anual": fila.get("interes_inversor_anual", 0),
                "cobro_teorico_compania": cobro_teorico,
                "cobro_compania": cobro_compania,
                "pago_inversor": pago_inversor,
                "beneficio_empresa": beneficio_empresa,
            })

    return pd.DataFrame(filas)


def resumen_notas_mes(df_inv: pd.DataFrame, df_cal: pd.DataFrame, df_control: pd.DataFrame, anio: int, mes: int):
    pagos = pagos_notas_mes(df_cal, anio, mes)
    detalle = preparar_detalle_notas(df_inv, pagos, df_cal=df_cal, df_control=df_control)
    if detalle.empty:
        return 0.0, 0.0, 0.0, detalle, pagos
    return float(detalle["cobro_compania"].sum()), float(detalle["pago_inversor"].sum()), float(detalle["beneficio_empresa"].sum()), detalle, pagos


def resumen_por_cuenta_cobro(detalle: pd.DataFrame) -> pd.DataFrame:
    if detalle.empty:
        return pd.DataFrame(columns=["cuenta_cobro", "cobro_compania"])
    return detalle.groupby("cuenta_cobro", as_index=False)["cobro_compania"].sum().sort_values("cobro_compania", ascending=False)


def resumen_capital_por_inversor_notas(df_inv: pd.DataFrame, solo_activo: bool = False) -> pd.DataFrame:
    trabajo = filtrar_notas(df_inv)
    hoy = pd.Timestamp.today().normalize()
    if solo_activo:
        trabajo = trabajo[(trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= hoy) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= hoy))]
    if trabajo.empty:
        return pd.DataFrame(columns=["inversor", "capital"])
    return trabajo.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)


def proximo_evento_nota(df_cal: pd.DataFrame, nota: int, tipo: str):
    hoy = pd.Timestamp.today().normalize()
    eventos = df_cal[(df_cal["tipo_evento"] == tipo) & (df_cal["nota"] == nota) & (df_cal["fecha"].notna()) & (df_cal["fecha"] >= hoy)].sort_values("fecha")
    return None if eventos.empty else eventos.iloc[0]["fecha"]


# =========================
# GLOBAL Y DASHBOARD
# =========================
def detectar_activo(row):
    tipo = limpiar_texto(row.get("tipo_inversion", ""))
    subtipo = limpiar_texto(row.get("subtipo_inversion", ""))
    nombre = limpiar_texto(row.get("nombre_activo", ""))
    if tipo == "nota" or nombre.startswith("nota"):
        return "notas"
    for activo in ["paraguay", "motoclick", "futbol", "fútbol"]:
        if activo in subtipo or activo in nombre:
            return "futbol" if activo == "fútbol" else activo
    return "otros"


def excluir_call(df: pd.DataFrame) -> pd.DataFrame:
    if "motivo" not in df.columns:
        return df.copy()
    return df[df["motivo"].apply(limpiar_texto) != "call"].copy()


def inversiones_activas_global(df_inv: pd.DataFrame, fecha=None) -> pd.DataFrame:
    if fecha is None:
        fecha = pd.Timestamp.today().normalize()
    fecha = pd.Timestamp(fecha).normalize()
    trabajo = excluir_call(df_inv)
    return trabajo[(trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= fecha) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha))].copy()


def tarjeta_kpi(titulo, valor, subtitulo="", estado="normal"):
    colores = {"normal": ("#ffffff", "#0e2338"), "positivo": ("#edf7ed", "#166534"), "riesgo": ("#fff4e5", "#b45309"), "negativo": ("#fee2e2", "#991b1b")}
    fondo, color = colores.get(estado, colores["normal"])
    st.markdown(
        f"""
        <div style="background:{fondo};padding:22px 24px;border-radius:22px;border:1px solid rgba(191,154,95,0.24);box-shadow:0 12px 32px rgba(15,35,55,0.08);min-height:130px;">
            <div style="font-size:13px;color:#667085;font-weight:700;text-transform:uppercase;">{titulo}</div>
            <div style="font-size:32px;color:{color};font-weight:850;margin-top:8px;">{valor}</div>
            <div style="font-size:13px;color:#667085;margin-top:8px;">{subtitulo}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def validar_base_datos(df_inv, df_cal, df_control):
    resultados = []
    def add(nombre, cantidad, gravedad):
        resultados.append({"Validación": nombre, "Incidencias": int(cantidad), "Estado": gravedad if cantidad > 0 else "OK"})
    add("Inversiones sin fecha de inversión", df_inv["fecha_inversion"].isna().sum() if "fecha_inversion" in df_inv.columns else len(df_inv), "ALTA")
    add("Inversiones sin capital invertido", (df_inv["capital_invertido"].fillna(0) <= 0).sum() if "capital_invertido" in df_inv.columns else len(df_inv), "ALTA")
    add("Inversiones sin inversor", (df_inv["inversor"].fillna("").astype(str).str.strip() == "").sum() if "inversor" in df_inv.columns else len(df_inv), "MEDIA")
    notas_filtradas = filtrar_notas(df_inv) if not df_inv.empty else pd.DataFrame()
    add("Notas sin número detectado", notas_filtradas["nota_num"].isna().sum() if not notas_filtradas.empty and "nota_num" in notas_filtradas.columns else 0, "ALTA")
    add("Eventos de calendario sin fecha", df_cal["fecha"].isna().sum() if "fecha" in df_cal.columns else 0, "MEDIA")
    add("Control de notas sin ticker", (df_control["ticker"].fillna("").astype(str).str.strip() == "").sum() if "ticker" in df_control.columns else 0, "ALTA")
    add("Control de notas sin precio de compra", df_control["precio_compra"].isna().sum() if "precio_compra" in df_control.columns else 0, "ALTA")
    return pd.DataFrame(resultados)


def detectar_alertas_financieras(df_inv, df_cal, df_control):
    """
    Alertas mejoradas para notas:
    - Evento próximo: observaciones y pagos cercanos.
    - Observación negativa real: solo si ya pasó y el precio está bajo barrera.
    - Observación sin dato: se avisa, pero no bloquea el cobro previsto.
    - Pago bloqueado: pago futuro o histórico cuya observación previa fue negativa real.
    - Pago previsto: pago futuro con observación pendiente o sin dato.
    """
    hoy = pd.Timestamp.today().normalize()
    alertas = []

    def add(tipo, detalle, fecha, prioridad, nota=""):
        alertas.append({
            "Tipo": tipo,
            "Nota": nota,
            "Detalle": detalle,
            "Fecha": pd.Timestamp(fecha).strftime("%d/%m/%Y") if pd.notna(fecha) else "",
            "Prioridad": prioridad,
        })

    if df_cal is not None and not df_cal.empty and "fecha" in df_cal.columns:
        eventos_7 = df_cal[
            (df_cal["fecha"].notna())
            & (df_cal["fecha"] >= hoy)
            & (df_cal["fecha"] <= hoy + pd.Timedelta(days=7))
        ].copy().sort_values(["fecha", "tipo_evento", "nota"])

        for _, row in eventos_7.iterrows():
            add(
                "Evento próximo",
                f"{row.get('tipo_evento', '')} de NOTA {row.get('nota', '')}",
                row.get("fecha"),
                "MEDIA",
                row.get("nota", ""),
            )

        # Observaciones ya vencidas o de hoy: revisar si fueron negativas reales o sin dato.
        observaciones_vencidas = df_cal[
            (df_cal["tipo_evento"] == "OBSERVACION")
            & (df_cal["fecha"].notna())
            & (df_cal["fecha"] <= hoy)
        ].copy().sort_values("fecha")

        for _, row in observaciones_vencidas.iterrows():
            nota = row.get("nota")
            fecha_obs = row.get("fecha")
            if pd.isna(nota):
                continue
            nota_int = int(nota)
            resultado, detalle = evaluar_nota_en_fecha(df_control, nota_int, fecha_obs, preferida="contingency")
            detalle_txt = resumen_detalle_observacion(detalle)

            if resultado == "NEGATIVA":
                add(
                    "Nota negativa real",
                    f"NOTA {nota_int}: observación negativa. No debe contarse el cobro. {detalle_txt}",
                    fecha_obs,
                    "ALTA",
                    nota_int,
                )
            elif resultado == "SIN DATO":
                add(
                    "Revisar dato faltante",
                    f"NOTA {nota_int}: la observación ya pasó, pero falta precio. Se mantiene como cobro previsto hasta revisar. {detalle_txt}",
                    fecha_obs,
                    "MEDIA",
                    nota_int,
                )
            elif resultado in ["SIN_CONTROL", "SIN_COLUMNAS", "SIN_OBSERVACION"]:
                add(
                    "Revisar configuración",
                    f"NOTA {nota_int}: no se ha podido evaluar correctamente ({resultado}).",
                    fecha_obs,
                    "ALTA",
                    nota_int,
                )

        # Observaciones futuras a 30 días: seguimiento.
        observaciones_futuras = df_cal[
            (df_cal["tipo_evento"] == "OBSERVACION")
            & (df_cal["fecha"].notna())
            & (df_cal["fecha"] > hoy)
            & (df_cal["fecha"] <= hoy + pd.Timedelta(days=30))
        ].copy().sort_values("fecha")

        for _, row in observaciones_futuras.iterrows():
            add(
                "Observación pendiente",
                f"NOTA {row.get('nota', '')}: observación futura. Se cuenta como cobro previsto hasta que llegue la fecha.",
                row.get("fecha"),
                "BAJA",
                row.get("nota", ""),
            )

        # Pagos próximos: indicar si están habilitados, previstos o bloqueados.
        pagos_30 = df_cal[
            (df_cal["tipo_evento"] == "PAGO")
            & (df_cal["fecha"].notna())
            & (df_cal["fecha"] >= hoy)
            & (df_cal["fecha"] <= hoy + pd.Timedelta(days=30))
        ].copy().sort_values("fecha")

        for _, row in pagos_30.iterrows():
            nota = row.get("nota")
            fecha_pago = row.get("fecha")
            if pd.isna(nota):
                continue
            nota_int = int(nota)
            fecha_obs = obtener_observacion_previa_nota(df_cal, nota_int, fecha_pago)
            resultado, detalle = evaluar_nota_en_fecha(df_control, nota_int, fecha_obs, preferida="contingency") if fecha_obs is not None else ("SIN_OBSERVACION", pd.DataFrame())
            detalle_txt = resumen_detalle_observacion(detalle)

            if resultado == "NEGATIVA":
                add(
                    "Pago bloqueado",
                    f"NOTA {nota_int}: pago próximo bloqueado por observación negativa. {detalle_txt}",
                    fecha_pago,
                    "ALTA",
                    nota_int,
                )
            elif resultado in ["PENDIENTE", "SIN DATO", "NO_EVALUADA"]:
                add(
                    "Pago previsto",
                    f"NOTA {nota_int}: pago próximo contado como previsto. Estado observación: {resultado}. {detalle_txt}",
                    fecha_pago,
                    "MEDIA",
                    nota_int,
                )
            elif resultado == "POSITIVA":
                add(
                    "Pago habilitado",
                    f"NOTA {nota_int}: pago próximo habilitado por observación positiva. {detalle_txt}",
                    fecha_pago,
                    "BAJA",
                    nota_int,
                )

    validaciones = validar_base_datos(df_inv, df_cal, df_control)
    errores_altos = validaciones[(validaciones["Incidencias"] > 0) & (validaciones["Estado"] == "ALTA")]
    for _, row in errores_altos.iterrows():
        add(
            "Validación crítica",
            f"{row['Validación']}: {row['Incidencias']} incidencias",
            hoy,
            "ALTA",
            "",
        )

    if not alertas:
        return pd.DataFrame(columns=["Tipo", "Nota", "Detalle", "Fecha", "Prioridad"])

    orden = {"ALTA": 0, "MEDIA": 1, "BAJA": 2}
    out = pd.DataFrame(alertas)
    out["orden_prioridad"] = out["Prioridad"].map(orden).fillna(9)
    out["fecha_orden"] = pd.to_datetime(out["Fecha"], errors="coerce", dayfirst=True)
    out = out.sort_values(["orden_prioridad", "fecha_orden", "Tipo"]).drop(columns=["orden_prioridad", "fecha_orden"])
    return out


def calcular_rentabilidad_inversiones_mes(df_inv, df_cal, df_control, anio: int, mes: int) -> pd.DataFrame:
    """
    Construye una tabla homogénea de rentabilidad mensual por inversión.
    - rentabilidad_beneficio_mes: beneficio empresa / capital.
    - rentabilidad_beneficio_anualizada: rentabilidad mensual x 12.
    - rentabilidad_pagada_inversor_mes: pago inversor / capital.
    - rentabilidad_pagada_inversor_anualizada: rentabilidad mensual pagada x 12.
    """
    filas = []

    # Notas estructuradas
    _, _, _, detalle_notas, _ = resumen_notas_mes(df_inv, df_cal, df_control, anio, mes)
    if detalle_notas is not None and not detalle_notas.empty:
        for _, row in detalle_notas.iterrows():
            capital = float(row.get("capital_invertido", 0) or 0)
            cobro = float(row.get("cobro_compania", 0) or 0)
            pago = float(row.get("pago_inversor", 0) or 0)
            beneficio = float(row.get("beneficio_empresa", 0) or 0)
            filas.append({
                "activo": "notas",
                "nombre_activo": f"NOTA {row.get('nota', '')}",
                "id_inversion": row.get("id_inversion", ""),
                "inversor": row.get("inversor", ""),
                "capital": capital,
                "cobro_compania_mes": cobro,
                "pago_inversor_mes": pago,
                "beneficio_empresa_mes": beneficio,
                "resultado_observacion": row.get("resultado_observacion", ""),
                "rentabilidad_beneficio_mes": beneficio / capital if capital else 0,
                "rentabilidad_beneficio_anualizada": (beneficio / capital * 12) if capital else 0,
                "rentabilidad_pagada_inversor_mes": pago / capital if capital else 0,
                "rentabilidad_pagada_inversor_anualizada": (pago / capital * 12) if capital else 0,
            })

    # Activos con rentabilidad fija / operativa
    for activo, tasa in [("paraguay", TASA_ANUAL_PARAGUAY), ("motoclick", TASA_ANUAL_MOTOCLICK), ("futbol", TASA_ANUAL_FUTBOL)]:
        det = detalle_activo_mes(df_inv, activo, tasa, anio, mes)
        if det is None or det.empty:
            continue
        for _, row in det.iterrows():
            capital = float(row.get("capital_invertido", 0) or 0)
            cobro = float(row.get("ingreso_bruto", 0) or 0)
            pago = float(row.get("pago_inversor_mes", 0) or 0)
            beneficio = float(row.get("beneficio_empresa_mes", 0) or 0)
            filas.append({
                "activo": activo,
                "nombre_activo": activo,
                "id_inversion": row.get("id_inversion", ""),
                "inversor": row.get("inversor", ""),
                "capital": capital,
                "cobro_compania_mes": cobro,
                "pago_inversor_mes": pago,
                "beneficio_empresa_mes": beneficio,
                "resultado_observacion": "NO APLICA",
                "rentabilidad_beneficio_mes": beneficio / capital if capital else 0,
                "rentabilidad_beneficio_anualizada": (beneficio / capital * 12) if capital else 0,
                "rentabilidad_pagada_inversor_mes": pago / capital if capital else 0,
                "rentabilidad_pagada_inversor_anualizada": (pago / capital * 12) if capital else 0,
            })

    return pd.DataFrame(filas)


def preparar_tabla_rentabilidad(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in ["capital", "cobro_compania_mes", "pago_inversor_mes", "beneficio_empresa_mes"]:
        if col in out.columns:
            out[col] = out[col].map(fmt)
    for col in [
        "rentabilidad_beneficio_mes",
        "rentabilidad_beneficio_anualizada",
        "rentabilidad_pagada_inversor_mes",
        "rentabilidad_pagada_inversor_anualizada",
    ]:
        if col in out.columns:
            out[col] = out[col].map(fmt_pct)
    return out


def clasificar_alerta_variacion(variacion):
    """Clasifica la alerta de una nota según su variación porcentual."""
    if pd.isna(variacion):
        return "SIN DATO"
    try:
        variacion = float(variacion)
    except Exception:
        return "SIN DATO"
    if variacion <= -35:
        return "ROJO"
    if variacion <= -25:
        return "AMARILLO"
    return "OK"


@st.cache_data(show_spinner=False, ttl=1800)
def construir_resumen_actual_notas_alertas(df_control: pd.DataFrame) -> pd.DataFrame:
    """Calcula precios actuales, variación y alerta por ticker/nota para la sección Notas y el dashboard."""
    if yf is None or df_control is None or df_control.empty:
        return pd.DataFrame()

    control = df_control.copy()
    barrera_col = next((c for c in ["contingency", "barrera_capital", "barrera_cupon"] if c in control.columns), None)
    faltan = [c for c in ["nota", "ticker", "precio_compra"] if c not in control.columns]
    if faltan or barrera_col is None:
        return pd.DataFrame()

    control["nota"] = pd.to_numeric(control["nota"], errors="coerce")
    control["ticker"] = control["ticker"].astype(str).str.strip().str.upper()
    control["precio_compra"] = pd.to_numeric(control["precio_compra"], errors="coerce")
    control[barrera_col] = pd.to_numeric(control[barrera_col], errors="coerce").apply(lambda x: x / 100 if pd.notna(x) and x > 1 else x)
    control = control.dropna(subset=["nota", "ticker", "precio_compra", barrera_col]).copy()

    filas = []
    for _, row in control.iterrows():
        ticker = row["ticker"]
        precio_actual = None
        try:
            hist = yf.Ticker(ticker).history(period="5d")
            if hist is not None and not hist.empty:
                cierre = hist["Close"].dropna()
                if not cierre.empty:
                    precio_actual = float(cierre.iloc[-1])
        except Exception:
            precio_actual = None

        precio_compra = float(row["precio_compra"])
        barrera = float(row[barrera_col])
        precio_contingencia = precio_compra * barrera
        variacion = None if precio_actual is None else ((precio_actual - precio_compra) / precio_compra) * 100
        alerta_variacion = clasificar_alerta_variacion(variacion)
        estado_barrera = "SIN DATO" if precio_actual is None else ("OK" if precio_actual >= precio_contingencia else "RIESGO")

        filas.append({
            "nota": int(row["nota"]),
            "ticker": ticker,
            "precio_compra": precio_compra,
            "precio_actual": precio_actual,
            "variacion_%": variacion,
            "precio_contingencia": precio_contingencia,
            "estado_barrera": estado_barrera,
            "alerta_variacion": alerta_variacion,
        })

    return pd.DataFrame(filas)


def resumen_alertas_por_nota(resumen_notas_actual: pd.DataFrame) -> pd.DataFrame:
    """Resume alertas amarillas/rojas por nota, tomando la peor variación de cada nota."""
    if resumen_notas_actual is None or resumen_notas_actual.empty:
        return pd.DataFrame(columns=["nota", "alerta", "peor_variacion_%", "tickers"])
    alertas = resumen_notas_actual[resumen_notas_actual["alerta_variacion"].isin(["AMARILLO", "ROJO"])].copy()
    if alertas.empty:
        return pd.DataFrame(columns=["nota", "alerta", "peor_variacion_%", "tickers"])

    orden_alerta = {"ROJO": 0, "AMARILLO": 1}
    filas = []
    for nota, grupo in alertas.groupby("nota"):
        grupo = grupo.copy()
        grupo["orden_alerta"] = grupo["alerta_variacion"].map(orden_alerta).fillna(9)
        peor = grupo.sort_values(["orden_alerta", "variacion_%"], ascending=[True, True]).iloc[0]
        filas.append({
            "nota": int(nota),
            "alerta": peor["alerta_variacion"],
            "peor_variacion_%": peor.get("variacion_%", None),
            "tickers": ", ".join(grupo["ticker"].astype(str).unique()),
        })
    out = pd.DataFrame(filas)
    out["orden"] = out["alerta"].map(orden_alerta).fillna(9)
    return out.sort_values(["orden", "peor_variacion_%"]).drop(columns=["orden"])


def colorear_filas_alerta_notas(row):
    alerta = row.get("alerta_variacion", "")
    if alerta == "ROJO":
        return ["background-color: #fee2e2; color: #7f1d1d; font-weight: 700"] * len(row)
    if alerta == "AMARILLO":
        return ["background-color: #fef3c7; color: #78350f; font-weight: 700"] * len(row)
    return [""] * len(row)


def inicio_semana_lunes(fecha):
    fecha = pd.Timestamp(fecha).normalize()
    return fecha - pd.Timedelta(days=fecha.weekday())


def resumen_cobros_semanales_mes_notas(df_inv: pd.DataFrame, df_cal: pd.DataFrame, df_control: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    """Agrupa los cobros de notas por semanas naturales lunes-domingo dentro de un mes."""
    pagos_mes = pagos_notas_mes(df_cal, anio, mes)
    detalle = preparar_detalle_notas(df_inv, pagos_mes, df_cal=df_cal, df_control=df_control)
    if detalle is None or detalle.empty:
        return pd.DataFrame(columns=["semana", "nota", "fecha_pago", "cobro_compania"])

    trabajo = detalle.copy()
    trabajo["fecha_pago"] = pd.to_datetime(trabajo["fecha_pago"], errors="coerce").dt.normalize()
    trabajo = trabajo[trabajo["fecha_pago"].notna()].copy()
    trabajo["inicio_semana"] = trabajo["fecha_pago"].apply(inicio_semana_lunes)
    trabajo["fin_semana"] = trabajo["inicio_semana"] + pd.Timedelta(days=6)
    trabajo["semana"] = trabajo.apply(
        lambda r: f"Semana del {r['inicio_semana'].day} - {r['fin_semana'].day} de {nombre_mes_es(int(r['inicio_semana'].month))}",
        axis=1,
    )

    resumen = trabajo.groupby(["inicio_semana", "fin_semana", "semana", "nota", "fecha_pago"], as_index=False)["cobro_compania"].sum()
    resumen = resumen.sort_values(["inicio_semana", "fecha_pago", "nota"])
    return resumen[["semana", "nota", "fecha_pago", "cobro_compania"]]


def mostrar_cobros_semanales_dashboard(df_inv: pd.DataFrame, df_cal: pd.DataFrame, df_control: pd.DataFrame, anio: int, mes: int):
    st.markdown("### Cobros semanales del mes")
    st.caption("Cada semana muestra el total previsto y un desplegable con el desglose por nota.")

    tabla_semanal = resumen_cobros_semanales_mes_notas(df_inv, df_cal, df_control, anio, mes)
    if tabla_semanal.empty:
        st.info("No hay cobros de notas previstos para ese mes.")
        return

    resumen_semana = (
        tabla_semanal
        .groupby("semana", as_index=False)["cobro_compania"]
        .sum()
        .rename(columns={"cobro_compania": "total_semana"})
    )

    st.dataframe(preparar_tabla_monetaria(resumen_semana, ["total_semana"]), use_container_width=True)

    for _, fila_semana in resumen_semana.iterrows():
        semana = fila_semana["semana"]
        total = float(fila_semana["total_semana"] or 0)
        detalle = tabla_semanal[tabla_semanal["semana"] == semana].copy()
        detalle = (
            detalle.groupby(["nota", "fecha_pago"], as_index=False)["cobro_compania"]
            .sum()
            .sort_values(["fecha_pago", "nota"])
        )
        detalle["nota"] = detalle["nota"].apply(lambda x: f"NOTA {int(x)}" if pd.notna(x) else "NOTA")
        with st.expander(f"{semana} · Total {fmt(total)}", expanded=False):
            st.dataframe(preparar_tabla_monetaria(detalle, ["cobro_compania"]), use_container_width=True)



def obtener_resumen_dashboard(df_inv, df_cal, df_control, anio: int | None = None, mes: int | None = None, vista_activo: str = "General", incluir_chaparro: bool = True):
    hoy_real = pd.Timestamp.today().normalize()
    if anio is None:
        anio = hoy_real.year
    if mes is None:
        mes = hoy_real.month
    fecha_analisis = pd.Timestamp(int(anio), int(mes), ultimo_dia_mes(int(anio), int(mes))).normalize()
    df_inv = aplicar_filtro_chaparro_fernandez(df_inv, incluir_chaparro)
    activas = inversiones_activas_global(df_inv, fecha_analisis)
    if not activas.empty:
        activas["activo"] = activas.apply(detectar_activo, axis=1)
    capital_total = activas["capital_invertido"].sum() if not activas.empty else 0
    c_notas, p_notas, b_notas, detalle_notas, _ = resumen_notas_mes(df_inv, df_cal, df_control, int(anio), int(mes))
    detalles_fijos = []
    for activo, tasa in [("paraguay", TASA_ANUAL_PARAGUAY), ("motoclick", TASA_ANUAL_MOTOCLICK), ("futbol", TASA_ANUAL_FUTBOL)]:
        det = detalle_activo_mes(df_inv, activo, tasa, int(anio), int(mes))
        if not det.empty:
            det["activo"] = activo
            detalles_fijos.append(det)
    d_fijos = pd.concat(detalles_fijos, ignore_index=True) if detalles_fijos else pd.DataFrame()
    cobro_fijos = d_fijos["ingreso_bruto"].sum() if not d_fijos.empty else 0
    pago_fijos = d_fijos["pago_inversor_mes"].sum() if not d_fijos.empty else 0
    beneficio_fijos = d_fijos["beneficio_empresa_mes"].sum() if not d_fijos.empty else 0

    cobro_total_mes = c_notas + cobro_fijos
    pago_total_mes = p_notas + pago_fijos
    beneficio_total_mes = b_notas + beneficio_fijos

    rentabilidad_beneficio_mes = beneficio_total_mes / capital_total if capital_total else 0
    rentabilidad_beneficio_anualizada = rentabilidad_beneficio_mes * 12
    rentabilidad_pagada_inversor_mes = pago_total_mes / capital_total if capital_total else 0
    rentabilidad_pagada_inversor_anualizada = rentabilidad_pagada_inversor_mes * 12

    rentabilidad_inversiones = calcular_rentabilidad_inversiones_mes(df_inv, df_cal, df_control, int(anio), int(mes))

    if not rentabilidad_inversiones.empty:
        rentabilidad_por_activo = rentabilidad_inversiones.groupby("activo", as_index=False).agg(
            capital=("capital", "sum"),
            cobro_compania_mes=("cobro_compania_mes", "sum"),
            pago_inversor_mes=("pago_inversor_mes", "sum"),
            beneficio_empresa_mes=("beneficio_empresa_mes", "sum"),
        )
        rentabilidad_por_activo["rentabilidad_beneficio_mes"] = rentabilidad_por_activo.apply(lambda r: r["beneficio_empresa_mes"] / r["capital"] if r["capital"] else 0, axis=1)
        rentabilidad_por_activo["rentabilidad_beneficio_anualizada"] = rentabilidad_por_activo["rentabilidad_beneficio_mes"] * 12
        rentabilidad_por_activo["rentabilidad_pagada_inversor_mes"] = rentabilidad_por_activo.apply(lambda r: r["pago_inversor_mes"] / r["capital"] if r["capital"] else 0, axis=1)
        rentabilidad_por_activo["rentabilidad_pagada_inversor_anualizada"] = rentabilidad_por_activo["rentabilidad_pagada_inversor_mes"] * 12
    else:
        rentabilidad_por_activo = pd.DataFrame()

    # Si el dashboard se filtra por activo, recalculamos los KPIs sobre ese bloque concreto.
    mapa_vista_activo = {
        "Notas": "notas",
        "Fútbol": "futbol",
        "MotoClick": "motoclick",
        "Paraguay": "paraguay",
    }
    activo_filtrado = mapa_vista_activo.get(str(vista_activo), None)
    if activo_filtrado:
        activas = activas[activas["activo"] == activo_filtrado].copy() if not activas.empty and "activo" in activas.columns else pd.DataFrame()
        capital_total = activas["capital_invertido"].sum() if not activas.empty else 0

        rentabilidad_inversiones = rentabilidad_inversiones[rentabilidad_inversiones["activo"] == activo_filtrado].copy() if not rentabilidad_inversiones.empty and "activo" in rentabilidad_inversiones.columns else pd.DataFrame()
        cobro_total_mes = float(rentabilidad_inversiones["cobro_compania_mes"].sum()) if not rentabilidad_inversiones.empty else 0.0
        pago_total_mes = float(rentabilidad_inversiones["pago_inversor_mes"].sum()) if not rentabilidad_inversiones.empty else 0.0
        beneficio_total_mes = float(rentabilidad_inversiones["beneficio_empresa_mes"].sum()) if not rentabilidad_inversiones.empty else 0.0

        rentabilidad_beneficio_mes = beneficio_total_mes / capital_total if capital_total else 0
        rentabilidad_beneficio_anualizada = rentabilidad_beneficio_mes * 12
        rentabilidad_pagada_inversor_mes = pago_total_mes / capital_total if capital_total else 0
        rentabilidad_pagada_inversor_anualizada = rentabilidad_pagada_inversor_mes * 12

        if not rentabilidad_inversiones.empty:
            rentabilidad_por_activo = rentabilidad_inversiones.groupby("activo", as_index=False).agg(
                capital=("capital", "sum"),
                cobro_compania_mes=("cobro_compania_mes", "sum"),
                pago_inversor_mes=("pago_inversor_mes", "sum"),
                beneficio_empresa_mes=("beneficio_empresa_mes", "sum"),
            )
            rentabilidad_por_activo["rentabilidad_beneficio_mes"] = rentabilidad_por_activo.apply(lambda r: r["beneficio_empresa_mes"] / r["capital"] if r["capital"] else 0, axis=1)
            rentabilidad_por_activo["rentabilidad_beneficio_anualizada"] = rentabilidad_por_activo["rentabilidad_beneficio_mes"] * 12
            rentabilidad_por_activo["rentabilidad_pagada_inversor_mes"] = rentabilidad_por_activo.apply(lambda r: r["pago_inversor_mes"] / r["capital"] if r["capital"] else 0, axis=1)
            rentabilidad_por_activo["rentabilidad_pagada_inversor_anualizada"] = rentabilidad_por_activo["rentabilidad_pagada_inversor_mes"] * 12
        else:
            rentabilidad_por_activo = pd.DataFrame()

    eventos_futuros = df_cal[(df_cal["fecha"].notna()) & (df_cal["fecha"] >= fecha_analisis)].copy().sort_values("fecha") if not df_cal.empty else pd.DataFrame()
    return {
        "activas": activas,
        "capital_total": capital_total,
        "cobro_total_mes": cobro_total_mes,
        "pago_total_mes": pago_total_mes,
        "beneficio_total_mes": beneficio_total_mes,
        "rentabilidad_beneficio_mes": rentabilidad_beneficio_mes,
        "rentabilidad_beneficio_anualizada": rentabilidad_beneficio_anualizada,
        "rentabilidad_pagada_inversor_mes": rentabilidad_pagada_inversor_mes,
        "rentabilidad_pagada_inversor_anualizada": rentabilidad_pagada_inversor_anualizada,
        "rentabilidad_inversiones": rentabilidad_inversiones,
        "rentabilidad_por_activo": rentabilidad_por_activo,
        "eventos_futuros": eventos_futuros,
        "detalle_notas": detalle_notas,
        "detalle_fijos": d_fijos,
    }


def grafico_capital_por_activo(activas):
    if px is None:
        st.warning("Falta plotly. Añade plotly a requirements.txt.")
        return
    if activas.empty:
        st.info("No hay inversiones activas para graficar.")
        return
    resumen = activas.groupby("activo", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
    fig = px.pie(resumen, names="activo", values="capital", hole=0.45, title="Distribución del capital activo por activo")
    fig.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", size=13), title_font=dict(size=20))
    st.plotly_chart(fig, use_container_width=True)


def grafico_capital_por_inversor(activas):
    if px is None:
        st.warning("Falta plotly. Añade plotly a requirements.txt.")
        return
    if activas.empty:
        st.info("No hay inversiones activas para graficar.")
        return
    resumen = activas.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False).head(10)
    fig = px.bar(resumen, x="inversor", y="capital", title="Top inversores por capital activo", text_auto=".2s")
    fig.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", size=13), title_font=dict(size=20), xaxis_title="Inversor", yaxis_title="Capital activo")
    st.plotly_chart(fig, use_container_width=True)


def grafico_beneficio_mensual(df_inv_calculo, df_cal, df_control):
    if px is None:
        st.warning("Falta plotly. Añade plotly a requirements.txt.")
        return
    hoy = pd.Timestamp.today().normalize()
    filas = []
    for i in range(11, -1, -1):
        fecha = hoy - pd.DateOffset(months=i)
        anio, mes = fecha.year, fecha.month
        _, _, b_notas, _, _ = resumen_notas_mes(df_inv, df_cal, df_control, anio, mes)
        detalles_fijos = []
        for activo, tasa in [("paraguay", TASA_ANUAL_PARAGUAY), ("motoclick", TASA_ANUAL_MOTOCLICK), ("futbol", TASA_ANUAL_FUTBOL)]:
            det = detalle_activo_mes(df_inv, activo, tasa, anio, mes)
            if not det.empty:
                detalles_fijos.append(det)
        d_fijos = pd.concat(detalles_fijos, ignore_index=True) if detalles_fijos else pd.DataFrame()
        b_fijos = d_fijos["beneficio_empresa_mes"].sum() if not d_fijos.empty else 0
        filas.append({"mes": f"{mes:02d}/{anio}", "beneficio": b_notas + b_fijos})
    data = pd.DataFrame(filas)
    fig = px.line(data, x="mes", y="beneficio", markers=True, title="Evolución del beneficio mensual estimado")
    fig.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(family="Inter", size=13), title_font=dict(size=20), xaxis_title="Mes", yaxis_title="Beneficio")
    st.plotly_chart(fig, use_container_width=True)





def etiqueta_tipo_interes(valor) -> str:
    """Convierte un interés decimal o porcentual en una etiqueta limpia."""
    try:
        v = float(valor)
        pct = v * 100 if abs(v) <= 1 else v
        if abs(pct - round(pct)) < 1e-9:
            return f"{int(round(pct))}%"
        return f"{pct:.2f}%".replace(".00%", "%")
    except Exception:
        return "SIN TIPO"


def construir_desglose_notas_por_tipo_inversor(df_inv: pd.DataFrame, detalle_notas: pd.DataFrame, fecha_analisis) -> pd.DataFrame:
    """Resume NOTAS por tipo pagado al inversor: capital activo, cobros, pagos y beneficio.

    Capital invertido: capital activo vivo al cierre del periodo seleccionado.
    Cobros/pagos/beneficio: importes del mes seleccionado construidos desde CALENDARIO_NOTAS.
    """
    fecha_analisis = pd.Timestamp(fecha_analisis).normalize()

    notas = filtrar_notas(df_inv)
    if notas is None or notas.empty:
        return pd.DataFrame(columns=[
            "tipo_inversor", "num_inversiones", "capital_invertido", "cobro_compania_mes",
            "pago_inversor_mes", "beneficio_empresa_mes", "rentabilidad_bruta_mes",
            "rentabilidad_bruta_anualizada", "coste_inversor_mes", "coste_inversor_anualizado",
            "margen_beneficio_mes", "margen_beneficio_anualizado",
        ])

    notas_activas = notas[
        (notas["fecha_inversion"].notna())
        & (notas["fecha_inversion"] <= fecha_analisis)
        & (notas["fecha_final_inversion"].isna() | (notas["fecha_final_inversion"] >= fecha_analisis))
    ].copy()

    if not notas_activas.empty:
        notas_activas["interes_inversor_anual"] = pd.to_numeric(notas_activas["interes_inversor_anual"], errors="coerce").fillna(0)
        capital_por_tipo = notas_activas.groupby("interes_inversor_anual", as_index=False).agg(
            num_inversiones=("id_inversion", "count"),
            capital_invertido=("capital_invertido", "sum"),
        )
    else:
        capital_por_tipo = pd.DataFrame(columns=["interes_inversor_anual", "num_inversiones", "capital_invertido"])

    if detalle_notas is not None and not detalle_notas.empty:
        det = detalle_notas.copy()
        det["interes_inversor_anual"] = pd.to_numeric(det["interes_inversor_anual"], errors="coerce").fillna(0)
        flujo_por_tipo = det.groupby("interes_inversor_anual", as_index=False).agg(
            cobro_compania_mes=("cobro_compania", "sum"),
            pago_inversor_mes=("pago_inversor", "sum"),
            beneficio_empresa_mes=("beneficio_empresa", "sum"),
        )
    else:
        flujo_por_tipo = pd.DataFrame(columns=["interes_inversor_anual", "cobro_compania_mes", "pago_inversor_mes", "beneficio_empresa_mes"])

    resumen = capital_por_tipo.merge(flujo_por_tipo, on="interes_inversor_anual", how="outer")
    for col in ["num_inversiones", "capital_invertido", "cobro_compania_mes", "pago_inversor_mes", "beneficio_empresa_mes"]:
        if col in resumen.columns:
            resumen[col] = pd.to_numeric(resumen[col], errors="coerce").fillna(0)

    if resumen.empty:
        return resumen

    resumen["tipo_inversor"] = resumen["interes_inversor_anual"].apply(etiqueta_tipo_interes)
    resumen["rentabilidad_bruta_mes"] = resumen.apply(lambda r: r["cobro_compania_mes"] / r["capital_invertido"] if r["capital_invertido"] else 0, axis=1)
    resumen["rentabilidad_bruta_anualizada"] = resumen["rentabilidad_bruta_mes"] * 12
    resumen["coste_inversor_mes"] = resumen.apply(lambda r: r["pago_inversor_mes"] / r["capital_invertido"] if r["capital_invertido"] else 0, axis=1)
    resumen["coste_inversor_anualizado"] = resumen["coste_inversor_mes"] * 12
    resumen["margen_beneficio_mes"] = resumen.apply(lambda r: r["beneficio_empresa_mes"] / r["capital_invertido"] if r["capital_invertido"] else 0, axis=1)
    resumen["margen_beneficio_anualizado"] = resumen["margen_beneficio_mes"] * 12

    total = {
        "interes_inversor_anual": 999,
        "tipo_inversor": "TOTAL",
        "num_inversiones": resumen["num_inversiones"].sum(),
        "capital_invertido": resumen["capital_invertido"].sum(),
        "cobro_compania_mes": resumen["cobro_compania_mes"].sum(),
        "pago_inversor_mes": resumen["pago_inversor_mes"].sum(),
        "beneficio_empresa_mes": resumen["beneficio_empresa_mes"].sum(),
    }
    total["rentabilidad_bruta_mes"] = total["cobro_compania_mes"] / total["capital_invertido"] if total["capital_invertido"] else 0
    total["rentabilidad_bruta_anualizada"] = total["rentabilidad_bruta_mes"] * 12
    total["coste_inversor_mes"] = total["pago_inversor_mes"] / total["capital_invertido"] if total["capital_invertido"] else 0
    total["coste_inversor_anualizado"] = total["coste_inversor_mes"] * 12
    total["margen_beneficio_mes"] = total["beneficio_empresa_mes"] / total["capital_invertido"] if total["capital_invertido"] else 0
    total["margen_beneficio_anualizado"] = total["margen_beneficio_mes"] * 12

    resumen = pd.concat([resumen, pd.DataFrame([total])], ignore_index=True)
    resumen = resumen.sort_values("interes_inversor_anual").reset_index(drop=True)
    columnas = [
        "tipo_inversor", "num_inversiones", "capital_invertido", "cobro_compania_mes",
        "pago_inversor_mes", "beneficio_empresa_mes", "rentabilidad_bruta_mes",
        "rentabilidad_bruta_anualizada", "coste_inversor_mes", "coste_inversor_anualizado",
        "margen_beneficio_mes", "margen_beneficio_anualizado",
    ]
    return resumen[columnas]


def preparar_tabla_tipo_inversor_notas(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in ["capital_invertido", "cobro_compania_mes", "pago_inversor_mes", "beneficio_empresa_mes"]:
        if col in out.columns:
            out[col] = out[col].map(fmt)
    for col in [
        "rentabilidad_bruta_mes", "rentabilidad_bruta_anualizada",
        "coste_inversor_mes", "coste_inversor_anualizado",
        "margen_beneficio_mes", "margen_beneficio_anualizado",
    ]:
        if col in out.columns:
            out[col] = out[col].map(fmt_pct)
    return out


def mostrar_desglose_notas_por_tipo_inversor(df_inv: pd.DataFrame, detalle_notas: pd.DataFrame, fecha_analisis, anio: int, mes: int):
    """Pinta en el dashboard el desglose de notas por 7,5%, 10%, 15%, etc."""
    st.markdown("### Notas por tipo pagado al inversor")
    st.caption("Capital activo de notas y resultado mensual separado por el interés pactado con cada inversor.")
    tabla = construir_desglose_notas_por_tipo_inversor(df_inv, detalle_notas, fecha_analisis)
    if tabla is None or tabla.empty:
        st.info("No hay notas activas o pagos de notas para mostrar en este periodo.")
        return

    columnas = [
        "tipo_inversor", "num_inversiones", "capital_invertido", "cobro_compania_mes",
        "pago_inversor_mes", "beneficio_empresa_mes", "rentabilidad_bruta_anualizada",
        "coste_inversor_anualizado", "margen_beneficio_anualizado",
    ]
    st.dataframe(preparar_tabla_tipo_inversor_notas(tabla[columnas]), use_container_width=True)

    excel_bytes = BytesIO()
    with pd.ExcelWriter(excel_bytes, engine="openpyxl") as writer:
        tabla.to_excel(writer, index=False, sheet_name="NOTAS_TIPO_INVERSOR")
    st.download_button(
        "Descargar desglose de notas por tipo inversor",
        data=excel_bytes.getvalue(),
        file_name=f"notas_por_tipo_inversor_{anio}_{mes:02d}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def mostrar_rentabilidad_por_activo_dashboard(tabla_activo: pd.DataFrame):
    """Muestra tarjetas y tabla de rentabilidad por activo directamente en el dashboard."""
    st.markdown("### Rentabilidad por activo")
    st.caption("Beneficio y coste de inversores por cada tipo de activo en el mes actual.")

    if tabla_activo is None or tabla_activo.empty:
        st.info("No hay datos de rentabilidad por activo para este mes.")
        return

    tabla = tabla_activo.copy().sort_values("beneficio_empresa_mes", ascending=False)

    cols = st.columns(min(4, len(tabla)))
    for i, (_, row) in enumerate(tabla.iterrows()):
        activo = str(row.get("activo", "Activo")).upper()
        capital = float(row.get("capital", 0) or 0)
        beneficio = float(row.get("beneficio_empresa_mes", 0) or 0)
        rent_mes = float(row.get("rentabilidad_beneficio_mes", 0) or 0)
        rent_anual = float(row.get("rentabilidad_beneficio_anualizada", 0) or 0)
        pagado_mes = float(row.get("rentabilidad_pagada_inversor_mes", 0) or 0)
        pagado_anual = float(row.get("rentabilidad_pagada_inversor_anualizada", 0) or 0)

        subtitulo = (
            f"Capital {fmt(capital)} · Beneficio {fmt(beneficio)}<br>"
            f"Pago inversores {fmt_pct(pagado_mes)} mes / {fmt_pct(pagado_anual)} anual"
        )
        estado = "positivo" if beneficio >= 0 else "negativo"
        with cols[i % len(cols)]:
            tarjeta_kpi(
                f"{activo} · rent. beneficio",
                f"{fmt_pct(rent_mes)} / {fmt_pct(rent_anual)} anual",
                subtitulo,
                estado,
            )

    with st.expander("Ver tabla completa de rentabilidad por activo", expanded=True):
        columnas = [
            "activo",
            "capital",
            "cobro_compania_mes",
            "pago_inversor_mes",
            "beneficio_empresa_mes",
            "rentabilidad_beneficio_mes",
            "rentabilidad_beneficio_anualizada",
            "rentabilidad_pagada_inversor_mes",
            "rentabilidad_pagada_inversor_anualizada",
        ]
        columnas = [c for c in columnas if c in tabla.columns]
        st.dataframe(preparar_tabla_rentabilidad(tabla[columnas]), use_container_width=True)


# =========================
# HISTÓRICO Y PROYECCIONES
# =========================
def primer_dia_mes_fecha(fecha):
    fecha = pd.Timestamp(fecha).normalize()
    return pd.Timestamp(fecha.year, fecha.month, 1)


def etiqueta_mes(fecha):
    fecha = pd.Timestamp(fecha).normalize()
    return f"{fecha.year}-{fecha.month:02d}"


def rango_meses(fecha_inicio, fecha_fin):
    inicio = primer_dia_mes_fecha(fecha_inicio)
    fin = primer_dia_mes_fecha(fecha_fin)
    meses = []
    actual = inicio
    while actual <= fin:
        meses.append(actual)
        actual = actual + pd.DateOffset(months=1)
    return meses


def fecha_minima_sistema(df_inv: pd.DataFrame, df_cal: pd.DataFrame):
    fechas = []
    if df_inv is not None and not df_inv.empty and "fecha_inversion" in df_inv.columns:
        serie = pd.to_datetime(df_inv["fecha_inversion"], errors="coerce").dropna()
        if not serie.empty:
            fechas.append(serie.min())
    if df_cal is not None and not df_cal.empty and "fecha" in df_cal.columns:
        serie = pd.to_datetime(df_cal["fecha"], errors="coerce").dropna()
        if not serie.empty:
            fechas.append(serie.min())
    if fechas:
        return min(fechas).normalize()
    return pd.Timestamp.today().normalize()


def construir_movimientos_historico_proyeccion(df_inv: pd.DataFrame, df_cal: pd.DataFrame, df_control: pd.DataFrame, fecha_inicio, fecha_fin, incluir_chaparro: bool = True) -> pd.DataFrame:
    """Construye movimientos mensuales de cobros, pagos y beneficio desde inicio y hacia futuro.

    Reglas:
    - NOTAS: usa CALENDARIO_NOTAS. Cada PAGO genera cobro compañía, pago inversor y beneficio por inversión.
    - Paraguay, MotoClick y Fútbol: devenga mes a mes según fecha_inversion / fecha_final_inversion.
    - Histórico/proyección se clasifica según si el mes es anterior o posterior al mes actual.
    """
    df_inv = aplicar_filtro_chaparro_fernandez(df_inv, incluir_chaparro)
    filas = []
    hoy = pd.Timestamp.today().normalize()

    for fecha_mes in rango_meses(fecha_inicio, fecha_fin):
        anio = int(fecha_mes.year)
        mes = int(fecha_mes.month)
        fin_mes = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes)).normalize()
        tipo_dato = "HISTÓRICO" if fin_mes <= hoy else "PROYECCIÓN"
        mes_label = etiqueta_mes(fecha_mes)

        # 1) Notas: se calculan únicamente cuando hay evento PAGO en calendario.
        _, _, _, detalle_notas, _ = resumen_notas_mes(df_inv, df_cal, df_control, anio, mes)
        if detalle_notas is not None and not detalle_notas.empty:
            for _, row in detalle_notas.iterrows():
                nota = row.get("nota", "")
                filas.append({
                    "mes_fecha": fecha_mes,
                    "mes": mes_label,
                    "tipo_dato": tipo_dato,
                    "activo": "notas",
                    "nombre_activo": f"NOTA {nota}",
                    "nota": nota,
                    "id_inversion": row.get("id_inversion", ""),
                    "inversor": row.get("inversor", ""),
                    "capital_base": float(row.get("capital_invertido", 0) or 0),
                    "cobrado_compania": float(row.get("cobro_compania", 0) or 0),
                    "pagado_inversores": float(row.get("pago_inversor", 0) or 0),
                    "beneficio_empresa": float(row.get("beneficio_empresa", 0) or 0),
                    "resultado_observacion": row.get("resultado_observacion", ""),
                })

        # 2) Activos con ingreso fijo o operativo.
        for activo, tasa in [("paraguay", TASA_ANUAL_PARAGUAY), ("motoclick", TASA_ANUAL_MOTOCLICK), ("futbol", TASA_ANUAL_FUTBOL)]:
            det = detalle_activo_mes(df_inv, activo, tasa, anio, mes)
            if det is None or det.empty:
                continue
            for _, row in det.iterrows():
                cobro = float(row.get("ingreso_bruto", 0) or 0)
                pago = float(row.get("pago_inversor_mes", 0) or 0)
                filas.append({
                    "mes_fecha": fecha_mes,
                    "mes": mes_label,
                    "tipo_dato": tipo_dato,
                    "activo": activo,
                    "nombre_activo": activo,
                    "nota": "",
                    "id_inversion": row.get("id_inversion", ""),
                    "inversor": row.get("inversor", ""),
                    "capital_base": float(row.get("capital_invertido", 0) or 0),
                    "cobrado_compania": cobro,
                    "pagado_inversores": pago,
                    "beneficio_empresa": cobro - pago,
                    "resultado_observacion": "NO APLICA",
                })

    if not filas:
        return pd.DataFrame(columns=[
            "mes_fecha", "mes", "tipo_dato", "activo", "nombre_activo", "nota", "id_inversion", "inversor",
            "capital_base", "cobrado_compania", "pagado_inversores", "beneficio_empresa", "resultado_observacion"
        ])

    out = pd.DataFrame(filas)
    out = out.sort_values(["mes_fecha", "activo", "nombre_activo", "inversor", "id_inversion"]).reset_index(drop=True)
    return out


def resumir_movimientos_por_mes(movimientos: pd.DataFrame) -> pd.DataFrame:
    if movimientos is None or movimientos.empty:
        return pd.DataFrame(columns=["mes", "tipo_dato", "cobrado_compania", "pagado_inversores", "beneficio_empresa"])
    resumen = movimientos.groupby(["mes_fecha", "mes", "tipo_dato"], as_index=False).agg(
        cobrado_compania=("cobrado_compania", "sum"),
        pagado_inversores=("pagado_inversores", "sum"),
        beneficio_empresa=("beneficio_empresa", "sum"),
    )
    return resumen.sort_values("mes_fecha")


def resumir_movimientos_por_mes_activo(movimientos: pd.DataFrame) -> pd.DataFrame:
    if movimientos is None or movimientos.empty:
        return pd.DataFrame(columns=["mes", "tipo_dato", "activo", "cobrado_compania", "pagado_inversores", "beneficio_empresa"])
    resumen = movimientos.groupby(["mes_fecha", "mes", "tipo_dato", "activo"], as_index=False).agg(
        cobrado_compania=("cobrado_compania", "sum"),
        pagado_inversores=("pagado_inversores", "sum"),
        beneficio_empresa=("beneficio_empresa", "sum"),
    )
    return resumen.sort_values(["mes_fecha", "activo"])


def resumir_movimientos_por_inversion(movimientos: pd.DataFrame) -> pd.DataFrame:
    if movimientos is None or movimientos.empty:
        return pd.DataFrame(columns=["activo", "nombre_activo", "id_inversion", "inversor", "cobrado_compania", "pagado_inversores", "beneficio_empresa"])
    resumen = movimientos.groupby(["activo", "nombre_activo", "id_inversion", "inversor"], as_index=False).agg(
        capital_base=("capital_base", "max"),
        cobrado_compania=("cobrado_compania", "sum"),
        pagado_inversores=("pagado_inversores", "sum"),
        beneficio_empresa=("beneficio_empresa", "sum"),
    )
    return resumen.sort_values("beneficio_empresa", ascending=False)


def preparar_tabla_movimientos(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in ["mes_fecha"]:
        if col in out.columns:
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d/%m/%Y")
    for col in ["capital_base", "cobrado_compania", "pagado_inversores", "beneficio_empresa", "cobro_simulado", "pago_simulado", "beneficio_simulado", "cobrado_total_con_simulacion", "pagado_total_con_simulacion", "beneficio_total_con_simulacion"]:
        if col in out.columns:
            out[col] = out[col].map(fmt)
    return out


def construir_simulacion_capital_extra(fecha_inicio, meses_duracion: int, capital_extra: float, tasa_cobro_anual: float, tasa_pago_anual: float, fecha_fin_rango) -> pd.DataFrame:
    if capital_extra <= 0 or meses_duracion <= 0:
        return pd.DataFrame()
    fecha_inicio = primer_dia_mes_fecha(fecha_inicio)
    fecha_fin_rango = primer_dia_mes_fecha(fecha_fin_rango)
    filas = []
    for i in range(int(meses_duracion)):
        fecha_mes = fecha_inicio + pd.DateOffset(months=i)
        if fecha_mes > fecha_fin_rango:
            break
        cobro = float(capital_extra) * float(tasa_cobro_anual) / 12
        pago = float(capital_extra) * float(tasa_pago_anual) / 12
        filas.append({
            "mes_fecha": fecha_mes,
            "mes": etiqueta_mes(fecha_mes),
            "capital_simulado": float(capital_extra),
            "cobro_simulado": cobro,
            "pago_simulado": pago,
            "beneficio_simulado": cobro - pago,
        })
    return pd.DataFrame(filas)


def exportar_historico_proyeccion_excel(resumen_mes, resumen_activo, resumen_inversion, detalle, simulacion=None) -> bytes:
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen_mes.to_excel(writer, index=False, sheet_name="RESUMEN_MENSUAL")
        resumen_activo.to_excel(writer, index=False, sheet_name="MES_ACTIVO")
        resumen_inversion.to_excel(writer, index=False, sheet_name="POR_INVERSION")
        detalle.to_excel(writer, index=False, sheet_name="DETALLE")
        if simulacion is not None and not simulacion.empty:
            simulacion.to_excel(writer, index=False, sheet_name="SIMULACION")
    return salida.getvalue()


def seccion_historico_y_proyecciones():
    df_inv, df_cal, df_control = cargar_excel_completo()
    st.markdown("## Histórico y proyecciones")
    st.caption("Control mensual de cuánto se cobra, cuánto se paga y qué beneficio queda desde el inicio, con proyección futura y simulador de nuevas inversiones.")

    hoy = pd.Timestamp.today().normalize()
    fecha_inicio_default = fecha_minima_sistema(df_inv, df_cal)
    fecha_fin_default = hoy + pd.DateOffset(months=12)

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1.2])
    fecha_inicio = pd.Timestamp(c1.date_input("Desde", value=fecha_inicio_default.date(), key="hist_proj_desde")).normalize()
    fecha_fin = pd.Timestamp(c2.date_input("Hasta", value=fecha_fin_default.date(), key="hist_proj_hasta")).normalize()
    activo_filtro = c3.selectbox("Activo", ["Todos", "notas", "paraguay", "motoclick", "futbol"], key="hist_proj_activo")
    incluir_chaparro = c4.checkbox(
        "Incluir Chaparro Fernández",
        value=False,
        key="hist_proj_incluir_chaparro",
        help="Si está desactivado, Chaparro Fernández queda fuera de capital, cobros, pagos y beneficio. Si está activado, se incluye como interno: pago inversor = cobro de la nota y beneficio = 0.",
    )

    if fecha_fin < fecha_inicio:
        st.error("La fecha final no puede ser anterior a la fecha inicial.")
        return

    df_inv_marcado = aplicar_filtro_chaparro_fernandez(df_inv, True)
    inversiones_chaparro = df_inv_marcado[df_inv_marcado.get("es_chaparro_fernandez", False) == True].copy() if not df_inv_marcado.empty else pd.DataFrame()
    if not inversiones_chaparro.empty:
        with st.expander("Ver inversiones detectadas como Chaparro Fernández", expanded=False):
            columnas_auditoria = [c for c in ["id_inversion", "inversor", "tipo_inversion", "subtipo_inversion", "nombre_activo", "capital_invertido", "interes_inversor_anual", "fecha_inversion", "fecha_final_inversion"] if c in inversiones_chaparro.columns]
            st.dataframe(preparar_tabla_monetaria(inversiones_chaparro[columnas_auditoria], ["capital_invertido"]), use_container_width=True)

    with st.spinner("Calculando histórico y proyecciones..."):
        movimientos = construir_movimientos_historico_proyeccion(
            df_inv,
            df_cal,
            df_control,
            fecha_inicio,
            fecha_fin,
            incluir_chaparro=incluir_chaparro,
        )

    if activo_filtro != "Todos" and not movimientos.empty:
        movimientos = movimientos[movimientos["activo"] == activo_filtro].copy()

    resumen_mes = resumir_movimientos_por_mes(movimientos)
    resumen_activo = resumir_movimientos_por_mes_activo(movimientos)
    resumen_inversion = resumir_movimientos_por_inversion(movimientos)

    historico = movimientos[movimientos["tipo_dato"] == "HISTÓRICO"].copy() if not movimientos.empty else pd.DataFrame()
    futuro = movimientos[movimientos["tipo_dato"] == "PROYECCIÓN"].copy() if not movimientos.empty else pd.DataFrame()

    total_cobrado_hist = float(historico["cobrado_compania"].sum()) if not historico.empty else 0.0
    total_pagado_hist = float(historico["pagado_inversores"].sum()) if not historico.empty else 0.0
    total_beneficio_hist = float(historico["beneficio_empresa"].sum()) if not historico.empty else 0.0

    total_cobrado_fut = float(futuro["cobrado_compania"].sum()) if not futuro.empty else 0.0
    total_pagado_fut = float(futuro["pagado_inversores"].sum()) if not futuro.empty else 0.0
    total_beneficio_fut = float(futuro["beneficio_empresa"].sum()) if not futuro.empty else 0.0

    st.markdown("### Totales")
    k1, k2, k3 = st.columns(3)
    k1.metric("Cobrado histórico", fmt(total_cobrado_hist))
    k2.metric("Pagado histórico", fmt(total_pagado_hist))
    k3.metric("Beneficio histórico", fmt(total_beneficio_hist))

    k4, k5, k6 = st.columns(3)
    k4.metric("Cobro proyectado", fmt(total_cobrado_fut))
    k5.metric("Pago proyectado", fmt(total_pagado_fut))
    k6.metric("Beneficio proyectado", fmt(total_beneficio_fut))

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Mes a mes",
        "Mes + activo",
        "Por inversión",
        "Simulador",
        "Detalle completo",
    ])

    with tab1:
        st.caption("Aquí ves, mes a mes, cuánto se cobra, cuánto se paga y qué beneficio queda.")
        if resumen_mes.empty:
            st.info("No hay movimientos para el rango seleccionado.")
        else:
            st.dataframe(preparar_tabla_movimientos(resumen_mes), use_container_width=True)
            if px is not None:
                fig = px.bar(
                    resumen_mes,
                    x="mes",
                    y=["cobrado_compania", "pagado_inversores", "beneficio_empresa"],
                    title="Cobrado, pagado y beneficio por mes",
                    barmode="group",
                )
                fig.update_layout(height=420, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig, use_container_width=True)

    with tab2:
        st.caption("El mismo cálculo mensual, pero separado por activo.")
        if resumen_activo.empty:
            st.info("No hay movimientos por activo para el rango seleccionado.")
        else:
            st.dataframe(preparar_tabla_movimientos(resumen_activo), use_container_width=True)

    with tab3:
        st.caption("Total cobrado, pagado y beneficio acumulado por cada inversión dentro del rango seleccionado.")
        if resumen_inversion.empty:
            st.info("No hay movimientos por inversión para el rango seleccionado.")
        else:
            st.dataframe(preparar_tabla_movimientos(resumen_inversion), use_container_width=True)

    with tab4:
        st.caption("Simula una nueva entrada de capital. Ejemplo: 200.000 al 25% de cobro y pagando 10% al inversor.")
        s1, s2, s3, s4 = st.columns(4)
        capital_extra = float(s1.number_input("Capital nuevo", min_value=0.0, value=200000.0, step=10000.0, key="sim_capital_nuevo"))
        tasa_cobro_pct = float(s2.number_input("% cobro anual compañía", min_value=0.0, value=25.0, step=0.5, key="sim_cobro_pct"))
        tasa_pago_pct = float(s3.number_input("% pago anual inversor", min_value=0.0, value=10.0, step=0.5, key="sim_pago_pct"))
        meses_duracion = int(s4.number_input("Meses de duración", min_value=1, max_value=120, value=12, step=1, key="sim_meses"))

        s5, s6 = st.columns(2)
        fecha_inicio_sim = pd.Timestamp(s5.date_input("Inicio simulación", value=hoy.date(), key="sim_fecha_inicio")).normalize()
        nombre_sim = s6.text_input("Nombre escenario", value="Nueva inversión simulada", key="sim_nombre")

        simulacion = construir_simulacion_capital_extra(
            fecha_inicio=fecha_inicio_sim,
            meses_duracion=meses_duracion,
            capital_extra=capital_extra,
            tasa_cobro_anual=tasa_cobro_pct / 100,
            tasa_pago_anual=tasa_pago_pct / 100,
            fecha_fin_rango=fecha_fin,
        )

        if simulacion.empty:
            st.info("La simulación no genera movimientos dentro del rango seleccionado.")
        else:
            total_sim_cobro = float(simulacion["cobro_simulado"].sum())
            total_sim_pago = float(simulacion["pago_simulado"].sum())
            total_sim_beneficio = float(simulacion["beneficio_simulado"].sum())
            mensual_cobro = float(capital_extra * (tasa_cobro_pct / 100) / 12)
            mensual_pago = float(capital_extra * (tasa_pago_pct / 100) / 12)
            mensual_beneficio = mensual_cobro - mensual_pago

            p1, p2, p3 = st.columns(3)
            p1.metric("Cobro mensual simulado", fmt(mensual_cobro))
            p2.metric("Pago mensual simulado", fmt(mensual_pago))
            p3.metric("Beneficio mensual simulado", fmt(mensual_beneficio))

            p4, p5, p6 = st.columns(3)
            p4.metric("Cobro total simulado", fmt(total_sim_cobro))
            p5.metric("Pago total simulado", fmt(total_sim_pago))
            p6.metric("Beneficio total simulado", fmt(total_sim_beneficio))

            base_mes = resumen_mes[["mes_fecha", "mes", "tipo_dato", "cobrado_compania", "pagado_inversores", "beneficio_empresa"]].copy() if not resumen_mes.empty else pd.DataFrame()
            if base_mes.empty:
                base_mes = pd.DataFrame(columns=["mes_fecha", "mes", "tipo_dato", "cobrado_compania", "pagado_inversores", "beneficio_empresa"])
            combinado = base_mes.merge(simulacion, on=["mes_fecha", "mes"], how="outer")
            combinado["tipo_dato"] = combinado["tipo_dato"].fillna("PROYECCIÓN")
            for col in ["cobrado_compania", "pagado_inversores", "beneficio_empresa", "cobro_simulado", "pago_simulado", "beneficio_simulado"]:
                if col in combinado.columns:
                    combinado[col] = pd.to_numeric(combinado[col], errors="coerce").fillna(0)
            combinado["cobrado_total_con_simulacion"] = combinado["cobrado_compania"] + combinado["cobro_simulado"]
            combinado["pagado_total_con_simulacion"] = combinado["pagado_inversores"] + combinado["pago_simulado"]
            combinado["beneficio_total_con_simulacion"] = combinado["beneficio_empresa"] + combinado["beneficio_simulado"]
            combinado = combinado.sort_values("mes_fecha")

            st.markdown(f"#### Resultado combinado: {nombre_sim}")
            columnas_sim = [
                "mes", "tipo_dato", "cobrado_compania", "pagado_inversores", "beneficio_empresa",
                "cobro_simulado", "pago_simulado", "beneficio_simulado",
                "cobrado_total_con_simulacion", "pagado_total_con_simulacion", "beneficio_total_con_simulacion",
            ]
            st.dataframe(preparar_tabla_movimientos(combinado[columnas_sim]), use_container_width=True)

    with tab5:
        st.caption("Detalle línea a línea usado para construir todos los cálculos.")
        if movimientos.empty:
            st.info("No hay detalle para el rango seleccionado.")
        else:
            columnas = [
                "mes", "tipo_dato", "activo", "nombre_activo", "nota", "id_inversion", "inversor", "capital_base",
                "cobrado_compania", "pagado_inversores", "beneficio_empresa", "resultado_observacion"
            ]
            columnas = [c for c in columnas if c in movimientos.columns]
            st.dataframe(preparar_tabla_movimientos(movimientos[columnas]), use_container_width=True)

    st.markdown("### Exportar")
    excel_bytes = exportar_historico_proyeccion_excel(resumen_mes, resumen_activo, resumen_inversion, movimientos)
    st.download_button(
        "Descargar histórico y proyecciones en Excel",
        data=excel_bytes,
        file_name=f"historico_proyecciones_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

def dashboard_financiero():
    df_inv, df_cal, df_control = cargar_excel_completo()
    st.markdown("## Dashboard financiero")
    st.caption("Panel ejecutivo de capital activo, cobros, pagos, beneficio y rentabilidades.")

    hoy = pd.Timestamp.today().normalize()
    col_activo, col_periodo_1, col_periodo_2, col_chaparro = st.columns([1.4, 1, 1, 1.2])
    vista_dashboard = col_activo.selectbox(
        "Dashboard",
        ["General", "Notas", "Fútbol", "MotoClick", "Paraguay"],
        key="dashboard_vista_activo",
    )
    incluir_chaparro = col_chaparro.checkbox(
        "Incluir Chaparro Fernández",
        value=False,
        key="dashboard_incluir_chaparro",
        help="Si está desactivado, Chaparro Fernández queda fuera de capital, cobros, pagos y rentabilidad. Si está activado, se incluye como interno: pago inversor = cobro de la nota y beneficio = 0.",
    )
    anio_dashboard = int(col_periodo_1.number_input(
        "Año del dashboard",
        min_value=2020,
        max_value=2100,
        value=hoy.year,
        key="dashboard_anio_general",
    ))
    mes_dashboard = int(col_periodo_2.number_input(
        "Mes del dashboard",
        min_value=1,
        max_value=12,
        value=hoy.month,
        key="dashboard_mes_general",
    ))
    st.caption(
        f"Vista seleccionada: {vista_dashboard} · Periodo: {nombre_mes_es(mes_dashboard)} {anio_dashboard} · "
        f"Chaparro Fernández: {'incluido' if incluir_chaparro else 'excluido'}"
    )

    df_inv_marcado = aplicar_filtro_chaparro_fernandez(df_inv, True)
    inversiones_chaparro = df_inv_marcado[df_inv_marcado.get("es_chaparro_fernandez", False) == True].copy() if not df_inv_marcado.empty else pd.DataFrame()
    if not inversiones_chaparro.empty:
        with st.expander("Ver inversiones detectadas como Chaparro Fernández", expanded=False):
            columnas_auditoria = [c for c in ["id_inversion", "inversor", "tipo_inversion", "subtipo_inversion", "nombre_activo", "capital_invertido", "interes_inversor_anual", "fecha_inversion", "fecha_final_inversion"] if c in inversiones_chaparro.columns]
            st.dataframe(preparar_tabla_monetaria(inversiones_chaparro[columnas_auditoria], ["capital_invertido"]), use_container_width=True)

    resumen_notas_actual = construir_resumen_actual_notas_alertas(df_control)
    alertas_notas = resumen_alertas_por_nota(resumen_notas_actual)
    if not alertas_notas.empty:
        rojas = int((alertas_notas["alerta"] == "ROJO").sum())
        amarillas = int((alertas_notas["alerta"] == "AMARILLO").sum())
        if rojas > 0:
            st.error(f"Alertas de notas: {rojas} en rojo y {amarillas} en amarillo por variación negativa.")
        else:
            st.warning(f"Alertas de notas: {amarillas} en amarillo por variación negativa.")
        with st.expander("Ver alertas de notas por variación", expanded=False):
            tabla_alertas = alertas_notas.copy()
            tabla_alertas["peor_variacion_%"] = tabla_alertas["peor_variacion_%"].apply(lambda x: f"{float(x):.2f}%" if pd.notna(x) else "Sin dato")
            st.dataframe(tabla_alertas, use_container_width=True)

    resumen = obtener_resumen_dashboard(
        df_inv,
        df_cal,
        df_control,
        anio_dashboard,
        mes_dashboard,
        vista_dashboard,
        incluir_chaparro=incluir_chaparro,
    )
    df_inv_calculo = aplicar_filtro_chaparro_fernandez(df_inv, incluir_chaparro)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        tarjeta_kpi("Capital activo total", fmt(resumen["capital_total"]), "Capital actualmente vivo", "normal")
    with c2:
        tarjeta_kpi("Cobro estimado mes", fmt(resumen["cobro_total_mes"]), "Ingresos brutos esperados", "positivo")
    with c3:
        tarjeta_kpi("Pago inversores mes", fmt(resumen["pago_total_mes"]), "Obligaciones estimadas", "riesgo")
    with c4:
        estado = "positivo" if resumen["beneficio_total_mes"] >= 0 else "negativo"
        tarjeta_kpi("Beneficio estimado mes", fmt(resumen["beneficio_total_mes"]), "Margen neto estimado", estado)

    st.markdown("### Rentabilidad del mes")
    r1, r2, r3, r4 = st.columns(4)
    with r1:
        tarjeta_kpi("Rent. beneficio mensual", fmt_pct(resumen["rentabilidad_beneficio_mes"]), "Beneficio / capital activo", "positivo" if resumen["rentabilidad_beneficio_mes"] >= 0 else "negativo")
    with r2:
        tarjeta_kpi("Rent. beneficio anualizada", fmt_pct(resumen["rentabilidad_beneficio_anualizada"]), "Mensual x 12", "positivo" if resumen["rentabilidad_beneficio_anualizada"] >= 0 else "negativo")
    with r3:
        tarjeta_kpi("% pagado inversores mes", fmt_pct(resumen["rentabilidad_pagada_inversor_mes"]), "Pago inversores / capital", "riesgo")
    with r4:
        tarjeta_kpi("% pagado inversores anual", fmt_pct(resumen["rentabilidad_pagada_inversor_anualizada"]), "Coste anualizado del capital", "riesgo")

    if vista_dashboard in ["General", "Notas"]:
        fecha_analisis_notas = pd.Timestamp(anio_dashboard, mes_dashboard, ultimo_dia_mes(anio_dashboard, mes_dashboard)).normalize()
        mostrar_desglose_notas_por_tipo_inversor(
            df_inv_calculo,
            resumen.get("detalle_notas", pd.DataFrame()),
            fecha_analisis_notas,
            anio_dashboard,
            mes_dashboard,
        )
        mostrar_cobros_semanales_dashboard(df_inv_calculo, df_cal, df_control, anio_dashboard, mes_dashboard)

    mostrar_rentabilidad_por_activo_dashboard(resumen.get("rentabilidad_por_activo", pd.DataFrame()))

    st.markdown("---")

    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "Capital por activo",
        "Capital por inversor",
        "Beneficio mensual",
        "Rentabilidad por activo",
        "Rentabilidad por inversión",
    ])
    with tab1:
        grafico_capital_por_activo(resumen["activas"])
    with tab2:
        grafico_capital_por_inversor(resumen["activas"])
    with tab3:
        grafico_beneficio_mensual(df_inv_calculo, df_cal, df_control)
    with tab4:
        st.caption("Resumen del periodo seleccionado por tipo de activo. La rentabilidad anualizada es la rentabilidad mensual multiplicada por 12.")
        tabla_activo = resumen.get("rentabilidad_por_activo", pd.DataFrame())
        if tabla_activo is None or tabla_activo.empty:
            st.info("No hay datos de rentabilidad por activo para este mes.")
        else:
            st.dataframe(preparar_tabla_rentabilidad(tabla_activo), use_container_width=True)
    with tab5:
        st.caption("Detalle inversión por inversión: cuánto genera, cuánto se paga al inversor y qué rentabilidad de beneficio deja.")
        tabla_inv = resumen.get("rentabilidad_inversiones", pd.DataFrame())
        if tabla_inv is None or tabla_inv.empty:
            st.info("No hay datos de rentabilidad por inversión para este mes.")
        else:
            columnas = [
                "activo", "nombre_activo", "id_inversion", "inversor", "capital",
                "cobro_compania_mes", "pago_inversor_mes", "beneficio_empresa_mes",
                "rentabilidad_beneficio_mes", "rentabilidad_beneficio_anualizada",
                "rentabilidad_pagada_inversor_mes", "rentabilidad_pagada_inversor_anualizada",
                "resultado_observacion",
            ]
            columnas = [c for c in columnas if c in tabla_inv.columns]
            st.dataframe(preparar_tabla_rentabilidad(tabla_inv[columnas]), use_container_width=True)


def centro_control_inversiones():
    df_inv, _, _ = cargar_excel_completo()
    st.markdown("## Centro de control de inversiones")
    st.caption("Consulta profesional por filtros.")
    c1, c2, c3 = st.columns(3)
    activo = c1.selectbox("Activo", ["Todos", "notas", "paraguay", "motoclick", "futbol", "otros"])
    inversores = ["Todos"] + sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
    inversor = c2.selectbox("Inversor", inversores)
    fecha = pd.Timestamp(c3.date_input("Fecha de análisis", value=pd.Timestamp.today().date())).normalize()
    activas = inversiones_activas_global(df_inv, fecha)
    if not activas.empty:
        activas["activo"] = activas.apply(detectar_activo, axis=1)
    if activo != "Todos":
        activas = activas[activas["activo"] == activo]
    if inversor != "Todos":
        activas = activas[activas["inversor"].astype(str).str.lower() == inversor.lower()]
    capital = activas["capital_invertido"].sum() if not activas.empty else 0
    num_inversiones = len(activas)
    ticket_medio = capital / num_inversiones if num_inversiones else 0
    c1, c2, c3 = st.columns(3)
    c1.metric("Capital activo", fmt(capital))
    c2.metric("Inversiones activas", num_inversiones)
    c3.metric("Ticket medio", fmt(ticket_medio))
    if activas.empty:
        st.info("No hay inversiones activas con estos filtros.")
        return
    tab1, tab2, tab3 = st.tabs(["Detalle", "Por activo", "Por inversor"])
    with tab1:
        st.dataframe(preparar_tabla_monetaria(activas, ["capital_invertido", "interes_inversor_anual", "interes_nota_anual"]), use_container_width=True)
    with tab2:
        resumen_activo = activas.groupby("activo", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
        st.dataframe(preparar_tabla_monetaria(resumen_activo, ["capital"]), use_container_width=True)
    with tab3:
        resumen_inv = activas.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
        st.dataframe(preparar_tabla_monetaria(resumen_inv, ["capital"]), use_container_width=True)


# =========================
# SECCIONES ORIGINALES MEJORADAS
# =========================
def seccion_activo(nombre_visible: str, activo_key: str, tasa_anual: float, incluir_ingresado_desde_inicio: bool = False):
    df_inv, _, _ = cargar_excel_completo()
    st.header(f"📌 Consultas {nombre_visible}")
    opciones = [
        f"¿Cuánto ingresará {nombre_visible} en un mes?", "¿Cuánto cobrará cada inversor ese mes?", "¿Cuánto cobrará un inversor concreto ese mes?",
        "¿Cuál será el beneficio de la empresa ese mes?", "¿Cuál es el total pagado a inversores desde el inicio?",
        f"¿Cuánto capital hay actualmente activo en {nombre_visible} hoy?", f"¿Cuánto capital había activo en {nombre_visible} en un mes concreto?",
    ]
    if incluir_ingresado_desde_inicio:
        opciones += ["¿Cuánto ha ingresado la compañía desde el inicio?", "¿Cuál es el beneficio total acumulado desde el inicio?"]
    consulta = st.selectbox("Elige una pregunta", opciones)
    necesita_mes = consulta in opciones[:4] or consulta == f"¿Cuánto capital había activo en {nombre_visible} en un mes concreto?"
    anio = mes = None
    if necesita_mes:
        c1, c2 = st.columns(2)
        anio = int(c1.number_input("Año", 2020, 2100, pd.Timestamp.today().year, key=f"{activo_key}_anio"))
        mes = int(c2.number_input("Mes", 1, 12, pd.Timestamp.today().month, key=f"{activo_key}_mes"))
    nombre_inversor = None
    if consulta == "¿Cuánto cobrará un inversor concreto ese mes?":
        inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        nombre_inversor = st.selectbox("Inversor", inversores) if inversores else st.text_input("Inversor")
    if st.button("Calcular", key=f"calc_{activo_key}_{consulta}"):
        if consulta == f"¿Cuánto ingresará {nombre_visible} en un mes?":
            detalle = detalle_activo_mes(df_inv, activo_key, tasa_anual, anio, mes)
            mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Ingreso bruto", fmt(detalle["ingreso_bruto"].sum() if not detalle.empty else 0))])
            if not detalle.empty:
                st.dataframe(preparar_tabla_monetaria(detalle, ["capital_invertido", "ingreso_bruto", "pago_inversor_mes", "beneficio_empresa_mes"]), use_container_width=True)
        elif consulta == "¿Cuánto cobrará cada inversor ese mes?":
            detalle = detalle_activo_mes(df_inv, activo_key, tasa_anual, anio, mes)
            if detalle.empty:
                st.info("No hay cobros de inversores para ese mes.")
            else:
                resumen = detalle.groupby("inversor", as_index=False)["pago_inversor_mes"].sum().rename(columns={"pago_inversor_mes": "cobro_mes"}).sort_values("cobro_mes", ascending=False)
                st.dataframe(preparar_tabla_monetaria(resumen, ["cobro_mes"]), use_container_width=True)
        elif consulta == "¿Cuánto cobrará un inversor concreto ese mes?":
            detalle = detalle_activo_mes(df_inv, activo_key, tasa_anual, anio, mes)
            filtrado = detalle[detalle["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()] if not detalle.empty else pd.DataFrame()
            mostrar_metricas("Resultado", [(f"Cobro de {nombre_inversor}", fmt(filtrado["pago_inversor_mes"].sum() if not filtrado.empty else 0))])
        elif consulta == "¿Cuál será el beneficio de la empresa ese mes?":
            detalle = detalle_activo_mes(df_inv, activo_key, tasa_anual, anio, mes)
            mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Beneficio empresa", fmt(detalle["beneficio_empresa_mes"].sum() if not detalle.empty else 0))])
            if not detalle.empty:
                st.dataframe(preparar_tabla_monetaria(detalle, ["capital_invertido", "ingreso_bruto", "pago_inversor_mes", "beneficio_empresa_mes"]), use_container_width=True)
        elif consulta == "¿Cuál es el total pagado a inversores desde el inicio?":
            mostrar_metricas("Resultado", [("Total pagado", fmt(total_pagado_activo_desde_inicio(df_inv, activo_key, tasa_anual)))])
        elif consulta == f"¿Cuánto capital hay actualmente activo en {nombre_visible} hoy?":
            bruto = capital_activo_en_fecha(df_inv, pd.Timestamp.today(), activo_key, False)
            real = capital_activo_en_fecha(df_inv, pd.Timestamp.today(), activo_key, True)
            mostrar_metricas("Resultado", [("Capital activo", fmt(bruto)), ("Capital activo real", fmt(real))])
        elif consulta == f"¿Cuánto capital había activo en {nombre_visible} en un mes concreto?":
            fecha = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))
            bruto = capital_activo_en_fecha(df_inv, fecha, activo_key, False)
            real = capital_activo_en_fecha(df_inv, fecha, activo_key, True)
            mostrar_metricas(f"Cierre {nombre_mes_es(mes)} {anio}", [("Capital activo", fmt(bruto)), ("Capital activo real", fmt(real))])
        elif consulta == "¿Cuánto ha ingresado la compañía desde el inicio?":
            mostrar_metricas("Resultado", [("Total ingresado", fmt(total_ingresado_activo_desde_inicio(df_inv, activo_key, tasa_anual)))])
        elif consulta == "¿Cuál es el beneficio total acumulado desde el inicio?":
            ingreso = total_ingresado_activo_desde_inicio(df_inv, activo_key, tasa_anual)
            pagado = total_pagado_activo_desde_inicio(df_inv, activo_key, tasa_anual)
            mostrar_metricas("Resultado", [("Beneficio acumulado", fmt(ingreso - pagado))])


def seccion_notas():
    df_inv, df_cal, df_control = cargar_excel_completo()
    st.header("🧾 Consultas Notas")
    consulta = st.selectbox("Elige una pregunta", [
        "¿Cuánto cobrará la compañía en un mes de notas?", "¿Cuánto se pagará a inversores en un mes de notas?", "¿Cuál será el beneficio de la empresa en un mes de notas?",
        "¿Cuánto cobrará cada inversor ese mes?", "¿Cuánto cobrará un inversor concreto ese mes?", "¿Cuánto ha cobrado la compañía desde el inicio?",
        "¿Cuánto se ha pagado a inversores desde el inicio?", "¿Cuál es el beneficio total desde el inicio?", "¿Cuál es el próximo pago de una nota?",
        "¿Cuál es la próxima observación de una nota?", "¿Cuánto capital hay invertido en total?", "¿Cuánto capital hay actualmente activo?",
        "¿Cuánto capital tiene un inversor?", "¿Cuánto capital activo tiene un inversor?", "Ver ranking de capital por inversor", "Ver ranking de capital activo",
    ])
    consultas_mes = ["¿Cuánto cobrará la compañía en un mes de notas?", "¿Cuánto se pagará a inversores en un mes de notas?", "¿Cuál será el beneficio de la empresa en un mes de notas?", "¿Cuánto cobrará cada inversor ese mes?", "¿Cuánto cobrará un inversor concreto ese mes?"]
    anio = mes = None
    if consulta in consultas_mes:
        c1, c2 = st.columns(2)
        anio = int(c1.number_input("Año", 2020, 2100, pd.Timestamp.today().year, key="notas_anio"))
        mes = int(c2.number_input("Mes", 1, 12, pd.Timestamp.today().month, key="notas_mes"))
    inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
    nombre_inversor = None
    if consulta in ["¿Cuánto cobrará un inversor concreto ese mes?", "¿Cuánto capital tiene un inversor?", "¿Cuánto capital activo tiene un inversor?"]:
        nombre_inversor = st.selectbox("Inversor", inversores) if inversores else st.text_input("Inversor")
    nota = None
    if consulta in ["¿Cuál es el próximo pago de una nota?", "¿Cuál es la próxima observación de una nota?"]:
        notas_disponibles = sorted([int(x) for x in df_cal.get("nota", pd.Series(dtype="Int64")).dropna().unique()])
        nota = st.selectbox("Número de nota", notas_disponibles) if notas_disponibles else int(st.number_input("Número de nota", min_value=1, value=1))
    if st.button("Calcular", key=f"calc_notas_{consulta}"):
        if consulta in consultas_mes:
            total_cobrado, total_pagado, total_beneficio, detalle, pagos = resumen_notas_mes(df_inv, df_cal, df_control, anio, mes)
            if consulta == "¿Cuánto cobrará la compañía en un mes de notas?":
                mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Cobra compañía", fmt(total_cobrado))])
                resumen_cuentas = resumen_por_cuenta_cobro(detalle)
                if not resumen_cuentas.empty:
                    st.dataframe(preparar_tabla_monetaria(resumen_cuentas, ["cobro_compania"]), use_container_width=True)
            elif consulta == "¿Cuánto se pagará a inversores en un mes de notas?":
                mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Pago inversores", fmt(total_pagado))])
            elif consulta == "¿Cuál será el beneficio de la empresa en un mes de notas?":
                mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Beneficio empresa", fmt(total_beneficio))])
            elif consulta == "¿Cuánto cobrará cada inversor ese mes?":
                resumen = detalle.groupby("inversor", as_index=False)["pago_inversor"].sum().rename(columns={"pago_inversor": "cobro_mes"}).sort_values("cobro_mes", ascending=False) if not detalle.empty else pd.DataFrame()
                st.dataframe(preparar_tabla_monetaria(resumen, ["cobro_mes"]), use_container_width=True) if not resumen.empty else st.info("No hay cobros de inversores para ese mes.")
            elif consulta == "¿Cuánto cobrará un inversor concreto ese mes?":
                filtrado = detalle[detalle["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()] if not detalle.empty else pd.DataFrame()
                mostrar_metricas("Resultado", [(f"Cobro de {nombre_inversor}", fmt(filtrado["pago_inversor"].sum() if not filtrado.empty else 0))])
            if not pagos.empty:
                with st.expander("Ver pagos detectados"):
                    st.dataframe(preparar_tabla_monetaria(pagos, []), use_container_width=True)
            if not detalle.empty:
                with st.expander("Ver detalle por nota e inversión"):
                    st.dataframe(preparar_tabla_monetaria(detalle, ["capital_invertido", "cobro_compania", "pago_inversor", "beneficio_empresa"]), use_container_width=True)
        elif consulta == "¿Cuánto ha cobrado la compañía desde el inicio?":
            detalle = preparar_detalle_notas(df_inv, pagos_notas_hasta_hoy(df_cal), df_cal=df_cal, df_control=df_control)
            mostrar_metricas("Resultado", [("Total cobrado compañía", fmt(detalle["cobro_compania"].sum() if not detalle.empty else 0))])
        elif consulta == "¿Cuánto se ha pagado a inversores desde el inicio?":
            detalle = preparar_detalle_notas(df_inv, pagos_notas_hasta_hoy(df_cal), df_cal=df_cal, df_control=df_control)
            mostrar_metricas("Resultado", [("Total pagado inversores", fmt(detalle["pago_inversor"].sum() if not detalle.empty else 0))])
        elif consulta == "¿Cuál es el beneficio total desde el inicio?":
            detalle = preparar_detalle_notas(df_inv, pagos_notas_hasta_hoy(df_cal), df_cal=df_cal, df_control=df_control)
            mostrar_metricas("Resultado", [("Beneficio total", fmt(detalle["beneficio_empresa"].sum() if not detalle.empty else 0))])
        elif consulta == "¿Cuál es el próximo pago de una nota?":
            fecha = proximo_evento_nota(df_cal, int(nota), "PAGO")
            st.success(f"El próximo pago de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}") if fecha is not None else st.info("No hay pagos futuros para esa nota.")
        elif consulta == "¿Cuál es la próxima observación de una nota?":
            fecha = proximo_evento_nota(df_cal, int(nota), "OBSERVACION")
            st.success(f"La próxima observación de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}") if fecha is not None else st.info("No hay observaciones futuras para esa nota.")
        elif consulta == "¿Cuánto capital hay invertido en total?":
            mostrar_metricas("Resultado", [("Capital total invertido", fmt(filtrar_notas(df_inv)["capital_invertido"].sum()))])
        elif consulta == "¿Cuánto capital hay actualmente activo?":
            trabajo = filtrar_notas(df_inv); hoy = pd.Timestamp.today().normalize()
            activas = trabajo[(trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= hoy) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= hoy))]
            mostrar_metricas("Resultado", [("Capital activo hoy", fmt(activas["capital_invertido"].sum() if not activas.empty else 0))])
        elif consulta == "¿Cuánto capital tiene un inversor?":
            trabajo = filtrar_notas(df_inv); filtrado = trabajo[trabajo["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()]
            mostrar_metricas("Resultado", [(f"Capital total de {nombre_inversor}", fmt(filtrado["capital_invertido"].sum() if not filtrado.empty else 0))])
        elif consulta == "¿Cuánto capital activo tiene un inversor?":
            trabajo = filtrar_notas(df_inv); hoy = pd.Timestamp.today().normalize()
            filtrado = trabajo[(trabajo["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()) & (trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= hoy) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= hoy))]
            mostrar_metricas("Resultado", [(f"Capital activo de {nombre_inversor}", fmt(filtrado["capital_invertido"].sum() if not filtrado.empty else 0))])
        elif consulta == "Ver ranking de capital por inversor":
            st.dataframe(preparar_tabla_monetaria(resumen_capital_por_inversor_notas(df_inv, False), ["capital"]), use_container_width=True)
        elif consulta == "Ver ranking de capital activo":
            st.dataframe(preparar_tabla_monetaria(resumen_capital_por_inversor_notas(df_inv, True), ["capital"]), use_container_width=True)


def seccion_notas_archivo():
    _, _, df_control = cargar_excel_completo()
    st.header("🧾 Notas")
    st.caption("Resumen de precios actuales, variación, barrera de contingencia y alertas por nota.")

    if yf is None:
        st.error("Falta yfinance. Añade yfinance a requirements.txt.")
        return
    if df_control is None or df_control.empty:
        st.warning("La hoja CONTROL_NOTAS está vacía o no existe.")
        return

    faltan = [c for c in ["nota", "ticker", "precio_compra"] if c not in df_control.columns]
    barrera_col = next((c for c in ["contingency", "barrera_capital", "barrera_cupon"] if c in df_control.columns), None)
    if faltan:
        st.error(f"En CONTROL_NOTAS faltan columnas: {', '.join(faltan)}")
        return
    if barrera_col is None:
        st.error("En CONTROL_NOTAS falta una columna de barrera: CONTINGENCY, BARRERA_CAPITAL o BARRERA_CUPON.")
        return

    if st.button("Actualizar precios actuales"):
        st.cache_data.clear()
        st.rerun()

    with st.spinner("Descargando precios actuales..."):
        resumen = construir_resumen_actual_notas_alertas(df_control)

    if resumen.empty:
        st.warning("No se pudo generar el resumen.")
        return

    alertas_resumen = resumen_alertas_por_nota(resumen)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Notas analizadas", resumen["nota"].nunique())
    c2.metric("Tickers", len(resumen))
    c3.metric("Notas en amarillo", int((alertas_resumen["alerta"] == "AMARILLO").sum()) if not alertas_resumen.empty else 0)
    c4.metric("Notas en rojo", int((alertas_resumen["alerta"] == "ROJO").sum()) if not alertas_resumen.empty else 0)

    tabla = resumen.copy()
    tabla["variacion_%"] = pd.to_numeric(tabla["variacion_%"], errors="coerce")
    columnas_dinero = ["precio_compra", "precio_actual", "precio_contingencia"]
    tabla_mostrar = preparar_tabla_monetaria(tabla, columnas_dinero)
    if "variacion_%" in tabla_mostrar.columns:
        tabla_mostrar["variacion_%"] = tabla["variacion_%"].apply(lambda x: f"{float(x):.2f}%" if pd.notna(x) else "Sin dato")

    st.dataframe(tabla_mostrar.style.apply(colorear_filas_alerta_notas, axis=1), use_container_width=True)

    st.markdown("### Alertas por variación")
    st.caption("Amarillo: variación igual o inferior a -25%. Rojo: variación igual o inferior a -35%.")
    if alertas_resumen.empty:
        st.success("No hay notas en amarillo ni en rojo por variación.")
    else:
        rojas = int((alertas_resumen["alerta"] == "ROJO").sum())
        amarillas = int((alertas_resumen["alerta"] == "AMARILLO").sum())
        if rojas > 0:
            st.error(f"Hay {rojas} notas en rojo y {amarillas} notas en amarillo.")
        else:
            st.warning(f"Hay {amarillas} notas en amarillo.")
        alertas_mostrar = alertas_resumen.copy()
        alertas_mostrar["peor_variacion_%"] = alertas_mostrar["peor_variacion_%"].apply(lambda x: f"{float(x):.2f}%" if pd.notna(x) else "Sin dato")
        st.dataframe(alertas_mostrar, use_container_width=True)


def seccion_alertas_notas():
    df_inv, df_cal, df_control = cargar_excel_completo()
    st.header("🚨 Alertas Notas")
    fecha = pd.Timestamp(st.date_input("Fecha de consulta", value=pd.Timestamp.today().date())).normalize()
    if df_cal.empty:
        st.warning("No existe la hoja CALENDARIO_NOTAS o está vacía.")
        return
    eventos = df_cal[df_cal["fecha"] == fecha].copy()
    st.subheader(f"Eventos del {fecha.strftime('%d/%m/%Y')}")
    st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True) if not eventos.empty else st.info("No hay observaciones ni pagos para esta fecha.")
    observaciones = eventos[eventos["tipo_evento"] == "OBSERVACION"].copy() if not eventos.empty else pd.DataFrame()
    pagos = eventos[eventos["tipo_evento"] == "PAGO"].copy() if not eventos.empty else pd.DataFrame()
    if not observaciones.empty:
        st.subheader("Evaluación de observaciones")
        for _, row in observaciones.iterrows():
            nota = int(row["nota"])
            resultado, detalle = evaluar_nota_en_fecha(df_control, nota, fecha, preferida="contingency")
            (st.success if resultado == "POSITIVA" else st.error if resultado == "NEGATIVA" else st.warning)(f"NOTA {nota}: {resultado}")
            if not detalle.empty:
                st.dataframe(preparar_tabla_monetaria(detalle, ["precio_compra", "precio_barrera", "cierre_usado"]), use_container_width=True)
    if not pagos.empty:
        st.subheader("Pagos del día")
        for _, row in pagos.iterrows():
            nota = int(row["nota"])
            previas = df_cal[(df_cal["nota"] == nota) & (df_cal["tipo_evento"] == "OBSERVACION") & (df_cal["fecha"] < fecha)].sort_values("fecha")
            if previas.empty:
                st.warning(f"NOTA {nota}: pago hoy, pero no he encontrado observación previa.")
                continue
            fecha_obs = previas.iloc[-1]["fecha"]
            resultado, detalle = evaluar_nota_en_fecha(df_control, nota, fecha_obs, preferida="contingency")
            (st.success if resultado == "POSITIVA" else st.error if resultado == "NEGATIVA" else st.warning)(f"NOTA {nota}: pago hoy. Observación previa {pd.Timestamp(fecha_obs).strftime('%d/%m/%Y')}: {resultado}")
            if not detalle.empty:
                with st.expander(f"Detalle NOTA {nota}"):
                    st.dataframe(preparar_tabla_monetaria(detalle, ["precio_compra", "precio_barrera", "cierre_usado"]), use_container_width=True)


def seccion_alertas_semana():
    _, df_cal, _ = cargar_excel_completo()
    st.header("📆 Alertas Semana")
    fecha_inicio = pd.Timestamp(st.date_input("Fecha de inicio", value=pd.Timestamp.today().date())).normalize()
    fecha_fin = fecha_inicio + pd.Timedelta(days=6)
    st.caption(f"Del {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}")
    eventos = df_cal[(df_cal["fecha"].notna()) & (df_cal["fecha"] >= fecha_inicio) & (df_cal["fecha"] <= fecha_fin)].copy().sort_values(["fecha", "tipo_evento", "nota"])
    if eventos.empty:
        st.info("No hay observaciones ni pagos esta semana.")
        return
    c1, c2 = st.columns(2)
    c1.metric("Observaciones", len(eventos[eventos["tipo_evento"] == "OBSERVACION"]))
    c2.metric("Pagos", len(eventos[eventos["tipo_evento"] == "PAGO"]))
    st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True)


def eventos_calendario_mes(df_cal: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    inicio = pd.Timestamp(anio, mes, 1)
    fin = inicio + pd.offsets.MonthEnd(0)
    eventos = df_cal[(df_cal["fecha"].notna()) & (df_cal["fecha"] >= inicio) & (df_cal["fecha"] <= fin)].copy()
    if eventos.empty:
        return eventos
    eventos["semana_mes"] = ((eventos["fecha"].dt.day - 1) // 7) + 1
    return eventos.sort_values(["fecha", "nota", "tipo_evento"])


def seccion_calendario_notas():
    _, df_cal, _ = cargar_excel_completo()
    st.header("🗓️ Calendario Notas")
    consulta = st.selectbox("Consulta", ["Esta semana", "Mes completo", "Semana concreta de un mes", "Exportar calendario de un mes"])
    if consulta == "Esta semana":
        hoy = pd.Timestamp.today().normalize()
        inicio = hoy - pd.Timedelta(days=hoy.weekday())
        fin = inicio + pd.Timedelta(days=6)
        eventos = df_cal[(df_cal["fecha"].notna()) & (df_cal["fecha"] >= inicio) & (df_cal["fecha"] <= fin)].copy().sort_values(["fecha", "nota", "tipo_evento"])
        st.caption(f"Del {inicio.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}")
        st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True) if not eventos.empty else st.info("No hay eventos esta semana.")
        return
    c1, c2 = st.columns(2)
    anio = int(c1.number_input("Año", 2020, 2100, pd.Timestamp.today().year, key=f"cal_{consulta}_anio"))
    mes = int(c2.number_input("Mes", 1, 12, pd.Timestamp.today().month, key=f"cal_{consulta}_mes"))
    eventos = eventos_calendario_mes(df_cal, anio, mes)
    if consulta == "Mes completo":
        st.subheader(f"Calendario de {nombre_mes_es(mes)} {anio}")
        st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True) if not eventos.empty else st.info("No hay eventos ese mes.")
    elif consulta == "Semana concreta de un mes":
        semana = int(st.number_input("Semana del mes", min_value=1, max_value=5, value=1))
        filtrado = eventos[eventos["semana_mes"] == semana].copy() if not eventos.empty else pd.DataFrame()
        st.dataframe(preparar_tabla_monetaria(filtrado, []), use_container_width=True) if not filtrado.empty else st.info("No hay eventos en esa semana.")
    else:
        if eventos.empty:
            st.info("No hay eventos para exportar en ese mes.")
        else:
            salida = BytesIO()
            exportar = eventos.copy(); exportar["fecha"] = exportar["fecha"].dt.strftime("%d/%m/%Y")
            with pd.ExcelWriter(salida, engine="openpyxl") as writer:
                exportar.to_excel(writer, index=False, sheet_name="CALENDARIO")
            st.download_button("Descargar Excel", data=salida.getvalue(), file_name=f"calendario_notas_{mes}_{anio}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")




def preparar_calendario_integrado_notas(df_inv: pd.DataFrame, df_cal: pd.DataFrame, df_control: pd.DataFrame, df_calls: pd.DataFrame | None = None, fecha_inicio=None, fecha_fin=None) -> pd.DataFrame:
    """Calendario único de notas: observaciones, pagos cobrables y calls.

    Reglas de pago:
    - Si la observación asociada es NEGATIVA real, el pago NO se muestra.
    - Si está PENDIENTE, POSITIVA, SIN DATO, SIN CONTROL o NO EVALUADA, se muestra como pago previsto/habilitado.
    """
    filas = []

    if fecha_inicio is not None:
        fecha_inicio = pd.Timestamp(fecha_inicio).normalize()
    if fecha_fin is not None:
        fecha_fin = pd.Timestamp(fecha_fin).normalize()

    def dentro_rango(fecha):
        if pd.isna(fecha):
            return False
        fecha = pd.Timestamp(fecha).normalize()
        if fecha_inicio is not None and fecha < fecha_inicio:
            return False
        if fecha_fin is not None and fecha > fecha_fin:
            return False
        return True

    # 1) Observaciones
    if df_cal is not None and not df_cal.empty:
        observaciones = df_cal[(df_cal["tipo_evento"] == "OBSERVACION") & (df_cal["fecha"].notna())].copy()
        for _, row in observaciones.iterrows():
            fecha = row.get("fecha")
            if not dentro_rango(fecha):
                continue
            nota = row.get("nota")
            if pd.isna(nota):
                continue
            nota_int = int(nota)
            resultado, detalle = evaluar_nota_en_fecha(df_control, nota_int, fecha, preferida="contingency") if df_control is not None else ("NO_EVALUADA", pd.DataFrame())
            filas.append({
                "fecha": pd.Timestamp(fecha).normalize(),
                "tipo_evento": "OBSERVACION",
                "nota": nota_int,
                "estado": resultado,
                "monto_cobro": 0.0,
                "detalle": resumen_detalle_observacion(detalle),
            })

        # 2) Pagos: solo se muestran si el cobro_compania total de la nota es > 0
        pagos = df_cal[(df_cal["tipo_evento"] == "PAGO") & (df_cal["fecha"].notna())].copy()
        for _, row in pagos.iterrows():
            fecha = row.get("fecha")
            if not dentro_rango(fecha):
                continue
            nota = row.get("nota")
            if pd.isna(nota):
                continue
            nota_int = int(nota)
            pago_df = pd.DataFrame([row])
            detalle_pago = preparar_detalle_notas(df_inv, pago_df, df_cal=df_cal, df_control=df_control)
            monto = float(detalle_pago["cobro_compania"].sum()) if not detalle_pago.empty and "cobro_compania" in detalle_pago.columns else 0.0
            fecha_obs = obtener_observacion_previa_nota(df_cal, nota_int, fecha)
            resultado_obs = "NO_EVALUADA"
            detalle_obs = pd.DataFrame()
            if fecha_obs is not None and df_control is not None and not df_control.empty:
                resultado_obs, detalle_obs = evaluar_nota_en_fecha(df_control, nota_int, fecha_obs, preferida="contingency")

            # Si la observación fue negativa real o el monto queda a 0, no ponemos el pago en calendario.
            if resultado_obs == "NEGATIVA" or monto <= 0:
                continue

            filas.append({
                "fecha": pd.Timestamp(fecha).normalize(),
                "tipo_evento": "PAGO",
                "nota": nota_int,
                "estado": resultado_obs,
                "monto_cobro": monto,
                "detalle": f"Cobro previsto/habilitado. Observación usada: {pd.Timestamp(fecha_obs).strftime('%d/%m/%Y') if fecha_obs is not None else 'sin observación'}",
            })

    # 3) Calls
    if df_calls is not None and not df_calls.empty:
        calls = df_calls.copy()
        # Normalización flexible por si la hoja tiene fecha_call o fecha.
        if "fecha_call" in calls.columns:
            calls["fecha_call"] = pd.to_datetime(calls["fecha_call"], errors="coerce", dayfirst=True).dt.normalize()
            col_fecha_call = "fecha_call"
        elif "fecha" in calls.columns:
            calls["fecha"] = pd.to_datetime(calls["fecha"], errors="coerce", dayfirst=True).dt.normalize()
            col_fecha_call = "fecha"
        else:
            col_fecha_call = None

        if col_fecha_call is not None:
            if "nota" in calls.columns:
                calls["nota"] = pd.to_numeric(calls["nota"], errors="coerce").astype("Int64")
            for _, row in calls.iterrows():
                fecha = row.get(col_fecha_call)
                if not dentro_rango(fecha):
                    continue
                nota = row.get("nota", pd.NA)
                filas.append({
                    "fecha": pd.Timestamp(fecha).normalize(),
                    "tipo_evento": "CALL",
                    "nota": int(nota) if pd.notna(nota) else "",
                    "estado": str(row.get("estado", "CALL POSIBLE")).upper() if "estado" in calls.columns else "CALL POSIBLE",
                    "monto_cobro": 0.0,
                    "detalle": "Fecha de posible call / cancelación anticipada",
                })

    calendario = pd.DataFrame(filas)
    if calendario.empty:
        return calendario
    orden_tipo = {"OBSERVACION": 1, "CALL": 2, "PAGO": 3}
    calendario["orden_tipo"] = calendario["tipo_evento"].map(orden_tipo).fillna(9)
    calendario = calendario.sort_values(["fecha", "orden_tipo", "nota"]).drop(columns=["orden_tipo"])
    return calendario


def preparar_tabla_calendario_integrado(calendario: pd.DataFrame) -> pd.DataFrame:
    if calendario is None or calendario.empty:
        return calendario
    out = calendario.copy()
    out["fecha"] = pd.to_datetime(out["fecha"], errors="coerce").dt.strftime("%d/%m/%Y")
    if "monto_cobro" in out.columns:
        out["monto_cobro"] = out["monto_cobro"].map(fmt)
    return out


def panel_alertas_y_calendario():
    df_inv, df_cal, df_control = cargar_excel_completo()
    df_calls = leer_hoja_excel("CALENDARIO_CALLS")

    st.markdown("## Alertas y calendario")
    st.caption("Calendario único con observaciones, pagos cobrables y calls. Los pagos con observación negativa real no aparecen como cobro.")

    hoy = pd.Timestamp.today().normalize()

    c1, c2, c3 = st.columns(3)
    vista = c1.selectbox("Vista", ["Próximos 30 días", "Este mes", "Mes concreto", "Rango personalizado"])

    if vista == "Próximos 30 días":
        fecha_inicio = hoy
        fecha_fin = hoy + pd.Timedelta(days=30)
    elif vista == "Este mes":
        fecha_inicio = pd.Timestamp(hoy.year, hoy.month, 1)
        fecha_fin = fecha_inicio + pd.offsets.MonthEnd(0)
    elif vista == "Mes concreto":
        anio = int(c2.number_input("Año", 2020, 2100, hoy.year, key="cal_unico_anio"))
        mes = int(c3.number_input("Mes", 1, 12, hoy.month, key="cal_unico_mes"))
        fecha_inicio = pd.Timestamp(anio, mes, 1)
        fecha_fin = fecha_inicio + pd.offsets.MonthEnd(0)
    else:
        fecha_inicio = pd.Timestamp(c2.date_input("Desde", value=hoy.date(), key="cal_unico_desde")).normalize()
        fecha_fin = pd.Timestamp(c3.date_input("Hasta", value=(hoy + pd.Timedelta(days=30)).date(), key="cal_unico_hasta")).normalize()

    calendario = preparar_calendario_integrado_notas(
        df_inv=df_inv,
        df_cal=df_cal,
        df_control=df_control,
        df_calls=df_calls,
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    alertas = detectar_alertas_financieras(df_inv, df_cal, df_control)

    pagos = calendario[calendario["tipo_evento"] == "PAGO"] if not calendario.empty else pd.DataFrame()
    observaciones = calendario[calendario["tipo_evento"] == "OBSERVACION"] if not calendario.empty else pd.DataFrame()
    calls = calendario[calendario["tipo_evento"] == "CALL"] if not calendario.empty else pd.DataFrame()
    monto_total = float(pagos["monto_cobro"].sum()) if not pagos.empty and "monto_cobro" in pagos.columns else 0.0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Observaciones", len(observaciones))
    m2.metric("Pagos cobrables", len(pagos))
    m3.metric("Monto a cobrar", fmt(monto_total))
    m4.metric("Calls", len(calls))

    if not alertas.empty:
        criticas = alertas[alertas["Prioridad"] == "ALTA"] if "Prioridad" in alertas.columns else pd.DataFrame()
        if not criticas.empty:
            st.error(f"Hay {len(criticas)} alertas críticas que requieren revisión.")
        else:
            st.warning(f"Hay {len(alertas)} alertas de seguimiento.")
        with st.expander("Ver alertas del sistema", expanded=False):
            st.dataframe(alertas, use_container_width=True)
    else:
        st.success("No hay alertas activas.")

    st.markdown("### Calendario único")
    if calendario.empty:
        st.info("No hay observaciones, pagos cobrables ni calls en el periodo seleccionado.")
    else:
        filtro_tipo = st.multiselect(
            "Filtrar eventos",
            ["OBSERVACION", "PAGO", "CALL"],
            default=["OBSERVACION", "PAGO", "CALL"],
        )
        tabla = calendario[calendario["tipo_evento"].isin(filtro_tipo)].copy() if filtro_tipo else calendario.copy()
        st.dataframe(preparar_tabla_calendario_integrado(tabla), use_container_width=True)

        salida = BytesIO()
        exportar = tabla.copy()
        exportar["fecha"] = pd.to_datetime(exportar["fecha"], errors="coerce").dt.strftime("%d/%m/%Y")
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            exportar.to_excel(writer, index=False, sheet_name="CALENDARIO_UNICO")
        st.download_button(
            "Descargar calendario único en Excel",
            data=salida.getvalue(),
            file_name=f"calendario_unico_notas_{fecha_inicio.strftime('%Y%m%d')}_{fecha_fin.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if px is not None and not calendario.empty:
        st.markdown("### Resumen visual")
        resumen_eventos = calendario.groupby("tipo_evento", as_index=False).size()
        fig = px.bar(resumen_eventos, x="tipo_evento", y="size", title="Eventos por tipo")
        fig.update_layout(height=330, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", xaxis_title="Tipo", yaxis_title="Cantidad")
        st.plotly_chart(fig, use_container_width=True)

def panel_calidad_datos():
    df_inv, df_cal, df_control = cargar_excel_completo()
    st.markdown("## Calidad de datos")
    validaciones = validar_base_datos(df_inv, df_cal, df_control)
    total_inc = validaciones["Incidencias"].sum()
    criticas = validaciones[(validaciones["Incidencias"] > 0) & (validaciones["Estado"] == "ALTA")]
    c1, c2, c3 = st.columns(3)
    c1.metric("Incidencias totales", int(total_inc))
    c2.metric("Incidencias críticas", len(criticas))
    c3.metric("Validaciones OK", len(validaciones[validaciones["Incidencias"] == 0]))
    if len(criticas) > 0:
        st.error("Hay incidencias críticas que pueden afectar a los cálculos.")
    elif total_inc > 0:
        st.warning("Hay incidencias menores a revisar.")
    else:
        st.success("Base de datos validada correctamente.")
    st.dataframe(validaciones, use_container_width=True)


def seccion_sistema_fondo():
    df_inv, df_cal, df_control = cargar_excel_completo()
    df_calls = leer_hoja_excel("CALENDARIO_CALLS")
    if not df_calls.empty:
        if "fecha_call" in df_calls.columns:
            df_calls["fecha_call"] = pd.to_datetime(df_calls["fecha_call"], errors="coerce", dayfirst=True).dt.normalize()
        if "nota" in df_calls.columns:
            df_calls["nota"] = pd.to_numeric(df_calls["nota"], errors="coerce").astype("Int64")
    st.header("🏦 Sistema Fondo")
    consulta = st.selectbox("Consulta", ["Panel global", "Capital activo total", "Capital activo por activo", "Capital activo por inversor", "Capital activo de un inversor concreto", "Resumen mensual global", "Validaciones", "Calls de esta semana", "Calls de este mes", "Próximos calls", "Calls vencidos", "Capital desglosado por inversor"])
    if consulta == "Panel global":
        dashboard_financiero()
    elif consulta == "Capital activo total":
        activas = inversiones_activas_global(df_inv)
        mostrar_metricas("Resultado", [("Capital activo total", fmt(activas["capital_invertido"].sum() if not activas.empty else 0))])
    elif consulta == "Capital activo por activo":
        activas = inversiones_activas_global(df_inv)
        if activas.empty:
            st.info("No hay inversiones activas.")
        else:
            activas["activo"] = activas.apply(detectar_activo, axis=1)
            resumen = activas.groupby("activo", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
            st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True)
    elif consulta == "Capital activo por inversor":
        activas = inversiones_activas_global(df_inv)
        resumen = activas.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False) if not activas.empty else pd.DataFrame()
        st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True) if not resumen.empty else st.info("No hay inversiones activas.")
    elif consulta in ["Capital activo de un inversor concreto", "Capital desglosado por inversor"]:
        inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        nombre = st.selectbox("Inversor", inversores) if inversores else st.text_input("Inversor")
        if consulta == "Capital desglosado por inversor":
            c1, c2 = st.columns(2)
            anio = int(c1.number_input("Año", 2020, 2100, pd.Timestamp.today().year))
            mes = int(c2.number_input("Mes", 1, 12, pd.Timestamp.today().month))
            fecha = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))
        else:
            fecha = pd.Timestamp.today().normalize()
        activas = inversiones_activas_global(df_inv, fecha=fecha)
        filtrado = activas[activas["inversor"].astype(str).str.lower() == str(nombre).lower()].copy()
        mostrar_metricas("Resultado", [("Capital activo", fmt(filtrado["capital_invertido"].sum() if not filtrado.empty else 0))])
        if not filtrado.empty:
            filtrado["activo"] = filtrado.apply(detectar_activo, axis=1)
            resumen = filtrado.groupby(["activo", "nombre_activo"], as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"})
            st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True)
    elif consulta == "Resumen mensual global":
        c1, c2 = st.columns(2)
        anio = int(c1.number_input("Año", 2020, 2100, pd.Timestamp.today().year))
        mes = int(c2.number_input("Mes", 1, 12, pd.Timestamp.today().month))
        c_notas, p_notas, b_notas, d_notas, _ = resumen_notas_mes(df_inv, df_cal, df_control, anio, mes)
        detalles = []
        for activo, tasa in [("paraguay", TASA_ANUAL_PARAGUAY), ("motoclick", TASA_ANUAL_MOTOCLICK), ("futbol", TASA_ANUAL_FUTBOL)]:
            det = detalle_activo_mes(df_inv, activo, tasa, anio, mes)
            if not det.empty:
                det["activo"] = activo; detalles.append(det)
        d_fijos = pd.concat(detalles, ignore_index=True) if detalles else pd.DataFrame()
        c_fijos = d_fijos["ingreso_bruto"].sum() if not d_fijos.empty else 0
        p_fijos = d_fijos["pago_inversor_mes"].sum() if not d_fijos.empty else 0
        b_fijos = d_fijos["beneficio_empresa_mes"].sum() if not d_fijos.empty else 0
        mostrar_metricas(f"Resumen global {nombre_mes_es(mes)} {anio}", [("Cobro compañía", fmt(c_notas + c_fijos)), ("Pago inversores", fmt(p_notas + p_fijos)), ("Beneficio", fmt(b_notas + b_fijos))])
        if not d_notas.empty:
            with st.expander("Detalle notas"):
                st.dataframe(preparar_tabla_monetaria(d_notas, ["capital_invertido", "cobro_compania", "pago_inversor", "beneficio_empresa"]), use_container_width=True)
        if not d_fijos.empty:
            with st.expander("Detalle activos fijos"):
                st.dataframe(preparar_tabla_monetaria(d_fijos, ["capital_invertido", "ingreso_bruto", "pago_inversor_mes", "beneficio_empresa_mes"]), use_container_width=True)
    elif consulta == "Validaciones":
        st.dataframe(validar_base_datos(df_inv, df_cal, df_control), use_container_width=True)
    else:
        if df_calls.empty or "fecha_call" not in df_calls.columns:
            st.warning("No existe la hoja CALENDARIO_CALLS o no tiene la columna fecha_call.")
            return
        hoy = pd.Timestamp.today().normalize()
        if consulta == "Calls de esta semana":
            inicio = hoy - pd.Timedelta(days=hoy.weekday()); fin = inicio + pd.Timedelta(days=6)
            res = df_calls[(df_calls["fecha_call"] >= inicio) & (df_calls["fecha_call"] <= fin)].copy()
        elif consulta == "Calls de este mes":
            res = df_calls[(df_calls["fecha_call"].dt.year == hoy.year) & (df_calls["fecha_call"].dt.month == hoy.month)].copy()
        elif consulta == "Próximos calls":
            res = df_calls[df_calls["fecha_call"] >= hoy].copy().sort_values("fecha_call").head(20)
        else:
            res = df_calls[df_calls["fecha_call"] < hoy].copy()
            if "estado" in res.columns:
                res = res[~res["estado"].apply(limpiar_texto).isin(["hecho", "realizado", "ejecutado", "call ejecutado"])]
        st.dataframe(preparar_tabla_monetaria(res, []), use_container_width=True) if not res.empty else st.info("No hay calls para esta consulta.")


# =========================
# EXTRACTOS
# =========================
def formatear_extracto_excel_bytes(contenido_raw: bytes, inversor: str, fecha_corte: datetime, detalle_df: "pd.DataFrame | None" = None) -> bytes:
    """
    Genera el extracto Excel profesional con:
    - Hoja PORTADA con resumen ejecutivo
    - Hoja DETALLE con filas de operación, cierres mensuales, anuales y cierre final
    - Hoja RESUMEN_MENSUAL con tabla de totales por mes
    - Todo en formato dólar, diseño premium
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.drawing.image import Image as XLImage
    import calendar as cal_mod

    # ── Paleta de colores ──────────────────────────────────────────────
    C_AZUL_OSC   = "0D2137"   # cabeceras principales
    C_AZUL_MED   = "1A3F5C"   # cabeceras secundarias
    C_AZUL_CLARO = "D6E9F8"   # fondo filas par detalle
    C_VERDE_OSC  = "1E4620"   # cierre anual texto
    C_VERDE      = "D9EAD3"   # cierre anual fondo
    C_NARANJA_OSC= "7F3F00"   # cierre mensual texto
    C_NARANJA    = "FCE5CD"   # cierre mensual fondo
    C_DORADO_OSC = "4A3000"   # cierre final texto
    C_DORADO     = "FFF2CC"   # cierre final fondo
    C_BLANCO     = "FFFFFF"
    C_GRIS_CLARO = "F7F9FC"
    C_GRIS_MED   = "D9D9D9"

    fmt_usd = '"$"#,##0.00'

    borde_fino  = Side(style="thin",   color=C_GRIS_MED)
    borde_medio = Side(style="medium", color=C_AZUL_MED)
    borde_std   = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)
    borde_top   = Border(left=borde_fino, right=borde_fino, top=borde_medio, bottom=borde_fino)

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def font(bold=False, size=11, color="000000", italic=False):
        return Font(name="Calibri", bold=bold, size=size, color=color, italic=italic)

    def aln(h="left", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    def set_cell(ws, row, col, value, bold=False, size=11, fcolor=None, tcolor="000000",
                 align_h="left", fmt=None, italic=False, border=None):
        c = ws.cell(row=row, column=col, value=value)
        c.font = font(bold=bold, size=size, color=tcolor, italic=italic)
        c.alignment = aln(h=align_h)
        if fcolor:
            c.fill = fill(fcolor)
        if fmt:
            c.number_format = fmt
        c.border = border if border else borde_std
        return c

    # ── Leer datos originales ─────────────────────────────────────────
    bio = BytesIO(contenido_raw)
    from openpyxl import load_workbook as lw
    wb_orig = lw(bio)

    # Leer RESUMEN
    ws_res_orig = wb_orig["RESUMEN"] if "RESUMEN" in wb_orig.sheetnames else None
    capital_total = 0.0
    total_intereses = 0.0
    if ws_res_orig:
        for row in ws_res_orig.iter_rows(min_row=2, values_only=True):
            if row and len(row) >= 4 and row[2] is not None:
                try:
                    capital_total = float(row[2])
                    total_intereses = float(row[3])
                except Exception:
                    pass

    # Leer DETALLE
    ws_det_orig = wb_orig["DETALLE"] if "DETALLE" in wb_orig.sheetnames else None
    det_rows = []
    det_cols = []
    if ws_det_orig:
        headers = [c.value for c in next(ws_det_orig.iter_rows(min_row=1, max_row=1))]
        det_cols = headers
        for row in ws_det_orig.iter_rows(min_row=2, values_only=True):
            det_rows.append(dict(zip(headers, row)))

    # Leer TOTALES_MES
    ws_tot_orig = wb_orig["TOTALES_MES"] if "TOTALES_MES" in wb_orig.sheetnames else None
    tot_rows = []
    if ws_tot_orig:
        for row in ws_tot_orig.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                tot_rows.append({"mes": row[0], "total_mes": row[1] if len(row) > 1 else 0})

    # Ordenar totales_mes cronológicamente
    def mes_sort_key(m):
        try:
            partes = str(m["mes"]).split("/")
            return (int(partes[1]), int(partes[0]))
        except Exception:
            return (9999, 99)
    tot_rows.sort(key=mes_sort_key)

    # ── Crear workbook nuevo ──────────────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)

    # ═══════════════════════════════════════════════════════════════
    # HOJA 1: PORTADA
    # ═══════════════════════════════════════════════════════════════
    ws_p = wb.create_sheet("PORTADA")
    ws_p.sheet_view.showGridLines = False
    ws_p.sheet_view.showRowColHeaders = False

    for col in range(1, 6):
        ws_p.column_dimensions[get_column_letter(col)].width = 22
    for r in range(1, 40):
        ws_p.row_dimensions[r].height = 30

    # Banda superior
    for r in range(1, 5):
        for c in range(1, 6):
            ws_p.cell(row=r, column=c).fill = fill(C_AZUL_OSC)

    ws_p.merge_cells("A1:E4")
    t = ws_p["A1"]
    t.value = "EXTRACTO DE INVERSIÓN"
    t.font = Font(name="Calibri", size=28, bold=True, color=C_BLANCO)
    t.alignment = aln(h="center", v="center")

    # Línea decorativa
    for c in range(1, 6):
        ws_p.cell(row=5, column=c).fill = fill("2E86C1")
    ws_p.row_dimensions[5].height = 6

    # Datos inversor
    ws_p.row_dimensions[7].height = 30
    ws_p.merge_cells("A7:E7")
    inv_cell = ws_p["A7"]
    inv_cell.value = inversor.upper()
    inv_cell.font = Font(name="Calibri", size=22, bold=True, color=C_AZUL_OSC)
    inv_cell.alignment = aln(h="center", v="center")

    ws_p.row_dimensions[8].height = 18
    ws_p.merge_cells("A8:E8")
    fecha_cell = ws_p["A8"]
    fecha_cell.value = f"Fecha de corte: {fecha_corte.strftime('%d/%m/%Y')}"
    fecha_cell.font = Font(name="Calibri", size=13, italic=True, color="555555")
    fecha_cell.alignment = aln(h="center")

    # Separador
    ws_p.row_dimensions[10].height = 4
    for c in range(1, 6):
        ws_p.cell(row=10, column=c).fill = fill(C_AZUL_CLARO)

    # Tarjetas resumen
    def tarjeta(ws, fila, col, titulo, valor, fmt_val, color_fondo, color_titulo):
        ws.row_dimensions[fila].height = 28
        ws.row_dimensions[fila+1].height = 36
        ws.row_dimensions[fila+2].height = 10
        ws.merge_cells(start_row=fila, start_column=col, end_row=fila, end_column=col+1)
        c1 = ws.cell(row=fila, column=col, value=titulo)
        c1.font = Font(name="Calibri", size=10, bold=True, color=color_titulo)
        c1.fill = fill(color_fondo)
        c1.alignment = aln(h="center", v="center")
        c1.border = borde_std
        ws.merge_cells(start_row=fila+1, start_column=col, end_row=fila+1, end_column=col+1)
        c2 = ws.cell(row=fila+1, column=col, value=valor)
        c2.font = Font(name="Calibri", size=18, bold=True, color=C_AZUL_OSC)
        c2.fill = fill(C_BLANCO)
        c2.alignment = aln(h="center", v="center")
        c2.number_format = fmt_val
        c2.border = borde_std

    tarjeta(ws_p, 12, 1, "CAPITAL INVERTIDO ACTIVO", capital_total, fmt_usd, C_AZUL_CLARO, C_AZUL_MED)
    tarjeta(ws_p, 12, 3, "TOTAL INTERESES ACUMULADOS", total_intereses, fmt_usd, C_VERDE, C_VERDE_OSC)
    tarjeta(ws_p, 17, 2, "TOTAL ACUMULADO (Capital + Intereses)", capital_total + total_intereses, fmt_usd, C_DORADO, C_DORADO_OSC)

    # Nota pie
    ws_p.row_dimensions[28].height = 20
    ws_p.merge_cells("A28:E28")
    n = ws_p["A28"]
    n.value = "Documento confidencial — Chaparro Fernández Wealth Management"
    n.font = Font(name="Calibri", size=9, italic=True, color="999999")
    n.alignment = aln(h="center")
    for c in range(1, 6):
        ws_p.cell(row=29, column=c).fill = fill(C_AZUL_OSC)
    ws_p.row_dimensions[29].height = 6

    # ═══════════════════════════════════════════════════════════════
    # HOJA 2: DETALLE con cierres mensuales, anuales y final
    # ═══════════════════════════════════════════════════════════════
    ws_d = wb.create_sheet("DETALLE")
    ws_d.sheet_view.showGridLines = False
    ws_d.freeze_panes = "A5"

    # Anchos de columna
    col_widths = [14, 16, 16, 18, 18, 12, 14, 16, 14, 12, 16]
    col_names  = ["ID", "Tipo inversión", "Subtipo", "Activo", "Mes",
                  "Fecha inversión", "Capital ($)", "Días devengados", "Días mes", "Interés mes ($)"]
    for i, w in enumerate(col_widths[:len(col_names)], 1):
        ws_d.column_dimensions[get_column_letter(i)].width = w
    # Ocultar columnas A, B, C, D, H, I, K (1, 2, 3, 4, 8, 9, 11)
    # La columna 11 (fecha_fin_op) se usa internamente para calcular capital activo al cierre
    for col_oculta in [1, 2, 3, 4, 8, 9, 11]:
        ws_d.column_dimensions[get_column_letter(col_oculta)].hidden = True

    # Título
    ws_d.merge_cells(f"A1:{get_column_letter(len(col_names))}1")
    t = ws_d["A1"]
    t.value = f"DETALLE DEL EXTRACTO — {inversor.upper()}"
    t.font = Font(name="Calibri", size=16, bold=True, color=C_BLANCO)
    t.fill = fill(C_AZUL_OSC)
    t.alignment = aln(h="center", v="center")
    ws_d.row_dimensions[1].height = 42

    ws_d.merge_cells(f"A2:{get_column_letter(len(col_names))}2")
    sub = ws_d["A2"]
    sub.value = f"Fecha de corte: {fecha_corte.strftime('%d/%m/%Y')}   |   Inversor: {inversor}"
    sub.font = Font(name="Calibri", size=10, italic=True, color="444444")
    sub.alignment = aln(h="center")
    ws_d.row_dimensions[2].height = 24

    ws_d.row_dimensions[3].height = 8
    for c in range(1, len(col_names)+1):
        ws_d.cell(row=3, column=c).fill = fill("2E86C1")

    # Cabecera
    for ci, nombre in enumerate(col_names, 1):
        c = ws_d.cell(row=4, column=ci, value=nombre)
        c.font = Font(name="Calibri", size=10, bold=True, color=C_BLANCO)
        c.fill = fill(C_AZUL_MED)
        c.alignment = aln(h="center")
        c.border = borde_std
    ws_d.row_dimensions[4].height = 28

    # Mapeo columnas detalle
    col_map = {
        "id_inversion": 1, "tipo_inversion": 2, "subtipo_inversion": 3,
        "nombre_activo": 4, "mes": 5, "fecha_inversion": 6,
        "capital_invertido": 7, "dias_devengados": 8, "dias_mes": 9, "interes_mes": 10,
        "fecha_fin_op": 11
    }

    def mes_key_from_str(mes_str):
        try:
            p = str(mes_str).split("/")
            return (int(p[1]), int(p[0]))
        except Exception:
            return (9999, 99)

    # Agrupar filas por mes
    from collections import defaultdict
    meses_orden = []
    filas_por_mes = defaultdict(list)
    for row in det_rows:
        mk = mes_key_from_str(row.get("mes", ""))
        if mk not in filas_por_mes:
            meses_orden.append(mk)
        filas_por_mes[mk].append(row)
    meses_orden = sorted(set(meses_orden))

    fila_excel = 5
    intereses_acum_anio = 0.0
    capital_anio = 0.0
    anio_actual = None
    intereses_acum_total = 0.0

    for mk in meses_orden:
        anio_mk, mes_mk = mk
        rows_mes = filas_por_mes[mk]

        # Cambio de año: insertar cierre anual del año anterior
        if anio_actual is not None and anio_mk != anio_actual:
            # CIERRE ANUAL
            ws_d.row_dimensions[fila_excel].height = 24
            ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=6)
            c = ws_d.cell(row=fila_excel, column=1, value=f"CIERRE {anio_actual}")
            c.font = Font(name="Calibri", size=11, bold=True, color=C_VERDE_OSC)
            c.fill = fill(C_VERDE)
            c.alignment = aln(h="center")
            c.border = borde_top

            cap_c = ws_d.cell(row=fila_excel, column=7, value=capital_anio)
            cap_c.font = Font(name="Calibri", size=11, bold=True, color=C_VERDE_OSC)
            cap_c.fill = fill(C_VERDE)
            cap_c.number_format = fmt_usd
            cap_c.alignment = aln(h="right")
            cap_c.border = borde_top

            for col_v in [8, 9]:
                cx = ws_d.cell(row=fila_excel, column=col_v, value="")
                cx.fill = fill(C_VERDE)
                cx.border = borde_top

            int_c = ws_d.cell(row=fila_excel, column=10, value=intereses_acum_anio)
            int_c.font = Font(name="Calibri", size=11, bold=True, color=C_VERDE_OSC)
            int_c.fill = fill(C_VERDE)
            int_c.number_format = fmt_usd
            int_c.alignment = aln(h="right")
            int_c.border = borde_top

            fila_excel += 1
            # Fila acumulado anual
            ws_d.row_dimensions[fila_excel].height = 20
            ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=8)
            ca = ws_d.cell(row=fila_excel, column=1, value=f"   Capital + Intereses acumulados {anio_actual}")
            ca.font = Font(name="Calibri", size=10, italic=True, color=C_VERDE_OSC)
            ca.fill = fill(C_VERDE)
            ca.alignment = aln(h="right")
            ca.border = borde_std
            for col_v in [9]:
                cx = ws_d.cell(row=fila_excel, column=col_v, value="")
                cx.fill = fill(C_VERDE)
                cx.border = borde_std
            total_anio_c = ws_d.cell(row=fila_excel, column=10, value=capital_anio + intereses_acum_anio)
            total_anio_c.font = Font(name="Calibri", size=10, bold=True, italic=True, color=C_VERDE_OSC)
            total_anio_c.fill = fill(C_VERDE)
            total_anio_c.number_format = fmt_usd
            total_anio_c.alignment = aln(h="right")
            total_anio_c.border = borde_std
            fila_excel += 1
            intereses_acum_anio = 0.0

        anio_actual = anio_mk

        # Filas de detalle del mes
        intereses_mes = 0.0
        capital_mes = 0.0
        par = (fila_excel % 2 == 0)
        for ri, row in enumerate(rows_mes):
            fondo = C_AZUL_CLARO if (fila_excel % 2 == 0) else C_GRIS_CLARO
            ws_d.row_dimensions[fila_excel].height = 26
            for col_key, col_idx in col_map.items():
                val = row.get(col_key, "")
                c = ws_d.cell(row=fila_excel, column=col_idx, value=val)
                c.font = Font(name="Calibri", size=10, color="222222")
                c.fill = fill(fondo)
                c.border = borde_std
                c.alignment = aln(h="right" if col_idx in [7, 10] else "center" if col_idx in [8, 9] else "left")
                if col_idx == 7:
                    c.number_format = fmt_usd
                if col_idx == 10:
                    c.number_format = fmt_usd
            v_cap = row.get("capital_invertido", 0) or 0
            v_int = row.get("interes_mes", 0) or 0
            try:
                intereses_mes += float(v_int)
                if float(v_cap) > capital_mes:
                    capital_mes = float(v_cap)
            except Exception:
                pass
            fila_excel += 1

        # Capital activo al cierre del mes = solo operaciones vivas el último día del mes
        # Una operación está viva si su fecha_fin_op >= último día del mes
        import calendar as _cal
        ultimo_dia = _cal.monthrange(anio_mk, mes_mk)[1]
        fin_mes_dt = datetime(anio_mk, mes_mk, ultimo_dia)
        capital_mes_real = 0.0
        for r in rows_mes:
            fecha_fin_op_str = str(r.get("fecha_fin_op", "") or "")
            capital_row = float(r.get("capital_invertido", 0) or 0)
            if not fecha_fin_op_str or fecha_fin_op_str in ("", "None", "nan"):
                # Sin fecha fin → siempre activa
                capital_mes_real += capital_row
            else:
                try:
                    ffo = datetime.strptime(fecha_fin_op_str, "%d/%m/%Y")
                    if ffo >= fin_mes_dt:
                        capital_mes_real += capital_row
                    # Si ffo < fin_mes_dt → cancelada dentro del mes → no suma al capital
                except Exception:
                    capital_mes_real += capital_row
        intereses_mes_real = sum(float(r.get("interes_mes", 0) or 0) for r in rows_mes)

        # CIERRE MENSUAL
        ws_d.row_dimensions[fila_excel].height = 28
        mes_label = f"{mes_mk:02d}/{anio_mk}"
        ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=6)
        cm = ws_d.cell(row=fila_excel, column=1, value=f"CIERRE {mes_label}")
        cm.font = Font(name="Calibri", size=10, bold=True, color=C_NARANJA_OSC)
        cm.fill = fill(C_NARANJA)
        cm.alignment = aln(h="center")
        cm.border = borde_top

        cap_cm = ws_d.cell(row=fila_excel, column=7, value=capital_mes_real)
        cap_cm.font = Font(name="Calibri", size=10, bold=True, color=C_NARANJA_OSC)
        cap_cm.fill = fill(C_NARANJA)
        cap_cm.number_format = fmt_usd
        cap_cm.alignment = aln(h="right")
        cap_cm.border = borde_top

        for col_v in [8, 9]:
            cx = ws_d.cell(row=fila_excel, column=col_v, value="")
            cx.fill = fill(C_NARANJA)
            cx.border = borde_top

        int_cm = ws_d.cell(row=fila_excel, column=10, value=intereses_mes_real)
        int_cm.font = Font(name="Calibri", size=10, bold=True, color=C_NARANJA_OSC)
        int_cm.fill = fill(C_NARANJA)
        int_cm.number_format = fmt_usd
        int_cm.alignment = aln(h="right")
        int_cm.border = borde_top

        fila_excel += 1
        intereses_acum_anio += intereses_mes_real
        intereses_acum_total += intereses_mes_real
        capital_anio = capital_mes_real

    # CIERRE ANUAL del último año
    if anio_actual is not None:
        ws_d.row_dimensions[fila_excel].height = 30
        ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=6)
        c = ws_d.cell(row=fila_excel, column=1, value=f"CIERRE {anio_actual}")
        c.font = Font(name="Calibri", size=11, bold=True, color=C_VERDE_OSC)
        c.fill = fill(C_VERDE)
        c.alignment = aln(h="center")
        c.border = borde_top
        cap_c = ws_d.cell(row=fila_excel, column=7, value=capital_anio)
        cap_c.font = Font(name="Calibri", size=11, bold=True, color=C_VERDE_OSC)
        cap_c.fill = fill(C_VERDE)
        cap_c.number_format = fmt_usd
        cap_c.alignment = aln(h="right")
        cap_c.border = borde_top
        for col_v in [8, 9]:
            cx = ws_d.cell(row=fila_excel, column=col_v, value="")
            cx.fill = fill(C_VERDE)
            cx.border = borde_top
        int_c = ws_d.cell(row=fila_excel, column=10, value=intereses_acum_anio)
        int_c.font = Font(name="Calibri", size=11, bold=True, color=C_VERDE_OSC)
        int_c.fill = fill(C_VERDE)
        int_c.number_format = fmt_usd
        int_c.alignment = aln(h="right")
        int_c.border = borde_top
        fila_excel += 1

        ws_d.row_dimensions[fila_excel].height = 20
        ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=8)
        ca = ws_d.cell(row=fila_excel, column=1, value=f"   Capital + Intereses acumulados {anio_actual}")
        ca.font = Font(name="Calibri", size=10, italic=True, color=C_VERDE_OSC)
        ca.fill = fill(C_VERDE)
        ca.alignment = aln(h="right")
        ca.border = borde_std
        cx = ws_d.cell(row=fila_excel, column=9, value="")
        cx.fill = fill(C_VERDE)
        cx.border = borde_std
        total_anio_c = ws_d.cell(row=fila_excel, column=10, value=capital_anio + intereses_acum_anio)
        total_anio_c.font = Font(name="Calibri", size=10, bold=True, italic=True, color=C_VERDE_OSC)
        total_anio_c.fill = fill(C_VERDE)
        total_anio_c.number_format = fmt_usd
        total_anio_c.alignment = aln(h="right")
        total_anio_c.border = borde_std
        fila_excel += 1

    # CIERRE FINAL
    ws_d.row_dimensions[fila_excel].height = 36
    ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=6)
    cf = ws_d.cell(row=fila_excel, column=1, value=f"CIERRE FINAL  {fecha_corte.strftime('%d/%m/%Y')}")
    cf.font = Font(name="Calibri", size=13, bold=True, color=C_DORADO_OSC)
    cf.fill = fill(C_DORADO)
    cf.alignment = aln(h="center")
    cf.border = borde_top
    cap_cf = ws_d.cell(row=fila_excel, column=7, value=capital_total)
    cap_cf.font = Font(name="Calibri", size=13, bold=True, color=C_DORADO_OSC)
    cap_cf.fill = fill(C_DORADO)
    cap_cf.number_format = fmt_usd
    cap_cf.alignment = aln(h="right")
    cap_cf.border = borde_top
    for col_v in [8, 9]:
        cx = ws_d.cell(row=fila_excel, column=col_v, value="")
        cx.fill = fill(C_DORADO)
        cx.border = borde_top
    int_cf = ws_d.cell(row=fila_excel, column=10, value=total_intereses)
    int_cf.font = Font(name="Calibri", size=13, bold=True, color=C_DORADO_OSC)
    int_cf.fill = fill(C_DORADO)
    int_cf.number_format = fmt_usd
    int_cf.alignment = aln(h="right")
    int_cf.border = borde_top
    fila_excel += 1

    ws_d.row_dimensions[fila_excel].height = 30
    ws_d.merge_cells(start_row=fila_excel, start_column=1, end_row=fila_excel, end_column=8)
    cfa = ws_d.cell(row=fila_excel, column=1, value="   TOTAL ACUMULADO  (Capital + Intereses)")
    cfa.font = Font(name="Calibri", size=12, bold=True, color=C_DORADO_OSC)
    cfa.fill = fill(C_DORADO)
    cfa.alignment = aln(h="right")
    cfa.border = borde_std
    cx = ws_d.cell(row=fila_excel, column=9, value="")
    cx.fill = fill(C_DORADO)
    cx.border = borde_std
    total_cf = ws_d.cell(row=fila_excel, column=10, value=capital_total + total_intereses)
    total_cf.font = Font(name="Calibri", size=12, bold=True, color=C_DORADO_OSC)
    total_cf.fill = fill(C_DORADO)
    total_cf.number_format = fmt_usd
    total_cf.alignment = aln(h="right")
    total_cf.border = borde_std

    # ═══════════════════════════════════════════════════════════════
    # HOJA 3: RESUMEN MENSUAL
    # ═══════════════════════════════════════════════════════════════
    ws_m = wb.create_sheet("RESUMEN MENSUAL")
    ws_m.sheet_view.showGridLines = False
    ws_m.column_dimensions["A"].width = 16
    ws_m.column_dimensions["B"].width = 22
    ws_m.column_dimensions["C"].width = 22

    ws_m.merge_cells("A1:C1")
    tm = ws_m["A1"]
    tm.value = "RESUMEN DE INTERESES POR MES"
    tm.font = Font(name="Calibri", size=16, bold=True, color=C_BLANCO)
    tm.fill = fill(C_AZUL_OSC)
    tm.alignment = aln(h="center", v="center")
    ws_m.row_dimensions[1].height = 42

    ws_m.merge_cells("A2:C2")
    ws_m["A2"].fill = fill("2E86C1")
    ws_m.row_dimensions[2].height = 5

    for ci, hdr in enumerate(["MES", "INTERESES ($)", "ACUMULADO ($)"], 1):
        c = ws_m.cell(row=3, column=ci, value=hdr)
        c.font = Font(name="Calibri", size=11, bold=True, color=C_BLANCO)
        c.fill = fill(C_AZUL_MED)
        c.alignment = aln(h="center")
        c.border = borde_std
    ws_m.row_dimensions[3].height = 28

    acum = 0.0
    for ri, tr in enumerate(tot_rows, 4):
        fondo = C_AZUL_CLARO if ri % 2 == 0 else C_GRIS_CLARO
        ws_m.row_dimensions[ri].height = 26
        c1 = ws_m.cell(row=ri, column=1, value=tr["mes"])
        c1.font = font(size=10)
        c1.fill = fill(fondo)
        c1.alignment = aln(h="center")
        c1.border = borde_std

        val = float(tr["total_mes"] or 0)
        acum += val
        c2 = ws_m.cell(row=ri, column=2, value=val)
        c2.font = font(size=10)
        c2.fill = fill(fondo)
        c2.number_format = fmt_usd
        c2.alignment = aln(h="right")
        c2.border = borde_std

        c3 = ws_m.cell(row=ri, column=3, value=acum)
        c3.font = font(size=10)
        c3.fill = fill(fondo)
        c3.number_format = fmt_usd
        c3.alignment = aln(h="right")
        c3.border = borde_std

    # Fila total
    fila_tot = 4 + len(tot_rows)
    ws_m.row_dimensions[fila_tot].height = 30
    ct = ws_m.cell(row=fila_tot, column=1, value="TOTAL")
    ct.font = Font(name="Calibri", size=11, bold=True, color=C_DORADO_OSC)
    ct.fill = fill(C_DORADO)
    ct.alignment = aln(h="center")
    ct.border = borde_top
    ct2 = ws_m.cell(row=fila_tot, column=2, value=acum)
    ct2.font = Font(name="Calibri", size=11, bold=True, color=C_DORADO_OSC)
    ct2.fill = fill(C_DORADO)
    ct2.number_format = fmt_usd
    ct2.alignment = aln(h="right")
    ct2.border = borde_top
    ct3 = ws_m.cell(row=fila_tot, column=3, value=acum)
    ct3.font = Font(name="Calibri", size=11, bold=True, color=C_DORADO_OSC)
    ct3.fill = fill(C_DORADO)
    ct3.number_format = fmt_usd
    ct3.alignment = aln(h="right")
    ct3.border = borde_top

    # Orden de hojas
    wb.move_sheet("PORTADA", offset=0)

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def generar_extractos(df_inv: pd.DataFrame, modo: str, inversor_elegido: str | None, anio: int, mes: int):
    """Genera extractos para inversores.

    REGLA DEFINITIVA PARA EXTRACTOS:
    - SOLO se tienen en cuenta las filas cuya columna tipo_operacion sea exactamente NUEVA.
    - NO se tienen en cuenta reinversiones, canceladas, vacías ni cualquier otro valor.
    - Las reinversiones no modifican el extracto del inversor: el inversor cobra según su operación matriz NUEVA.
    """
    df = df_inv.copy()

    # Normalizamos columnas de texto necesarias.
    for col in [
        "inversor",
        "tipo_inversion",
        "subtipo_inversion",
        "nombre_activo",
        "tipo_operacion",
        "capital_nuevo_real",
        "motivo",
        "id_inversion",
    ]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    # ==========================================
    # FILTRO PRINCIPAL DE EXTRACTOS
    # ==========================================
    # Según la regla definida: para extractos SOLO cuenta columna O / tipo_operacion = NUEVA.
    # Todo lo demás queda fuera: reinversion, cancelada, call, vacío, etc.
    if "tipo_operacion" not in df.columns:
        st.error("Falta la columna tipo_operacion en la hoja INVERSIONES. Para generar extractos debe existir y contener 'NUEVA'.")
        return []

    df["tipo_operacion_normalizada"] = df["tipo_operacion"].astype(str).str.strip().str.upper()

    # NUEVA: incluir, calcular hasta fecha de corte (ignorar fecha_final)
    # CANCELADA: incluir, calcular hasta fecha_final_inversion
    # REINVERSION y cualquier otro: excluir
    df = df[df["tipo_operacion_normalizada"].isin(["NUEVA", "CANCELADA"])].copy()

    if df.empty:
        return []

    if modo == "Un inversor" and inversor_elegido:
        df = df[df["inversor"].str.upper() == inversor_elegido.upper()].copy()

    if df.empty:
        return []

    fecha_corte = datetime(anio, mes, ultimo_dia_mes(anio, mes))

    filas = []
    for _, row in df.iterrows():
        fecha_inicio = row.get("fecha_inversion")
        if pd.isna(fecha_inicio):
            continue

        fecha_inicio_dt = pd.Timestamp(fecha_inicio).to_pydatetime()
        tipo_op = str(row.get("tipo_operacion_normalizada", "")).strip().upper()
        fecha_final_excel = row.get("fecha_final_inversion")

        if tipo_op == "CANCELADA":
            if pd.isna(fecha_final_excel):
                continue
            fecha_fin = min(pd.Timestamp(fecha_final_excel).to_pydatetime(), fecha_corte)
        else:
            # NUEVA: siempre hasta fecha de corte
            fecha_fin = fecha_corte

        if fecha_inicio_dt > fecha_fin:
            continue

        actual = datetime(fecha_inicio_dt.year, fecha_inicio_dt.month, 1)
        fin_mes = datetime(fecha_fin.year, fecha_fin.month, 1)

        while actual <= fin_mes:
            dias_mes = calendar.monthrange(actual.year, actual.month)[1]
            inicio_mes = datetime(actual.year, actual.month, 1)
            fin_mes_real = datetime(actual.year, actual.month, dias_mes)
            inicio_calc = max(fecha_inicio_dt, inicio_mes)
            fin_calc = min(fecha_fin, fin_mes_real)

            if inicio_calc <= fin_calc:
                dias = (fin_calc - inicio_calc).days + 1
                capital = float(row.get("capital_invertido", 0))
                interes = float(row.get("interes_inversor_anual", 0))
                interes_mes = round((capital * interes / 12) * dias / dias_mes, 2)
                mes_fecha = datetime(actual.year, actual.month, 1)
                filas.append({
                    "mes_fecha": mes_fecha,
                    "fecha_inversion_orden": fecha_inicio_dt,
                    "inversor": row.get("inversor", ""),
                    "id_inversion": row.get("id_inversion", ""),
                    "tipo_inversion": row.get("tipo_inversion", ""),
                    "subtipo_inversion": row.get("subtipo_inversion", ""),
                    "nombre_activo": row.get("nombre_activo", ""),
                    "mes": f"{actual.month:02d}/{actual.year}",
                    "fecha_inversion": pd.Timestamp(fecha_inicio).strftime("%d/%m/%Y"),
                    "capital_invertido": capital,
                    "dias_devengados": dias,
                    "dias_mes": dias_mes,
                    "interes_mes": interes_mes,
                    "fecha_fin_op": fecha_fin.strftime("%d/%m/%Y"),
                })

            actual = datetime(actual.year + 1, 1, 1) if actual.month == 12 else datetime(actual.year, actual.month + 1, 1)

    resultado = pd.DataFrame(filas)
    if resultado.empty:
        return []

    resultado = resultado.sort_values(
        ["inversor", "mes_fecha", "fecha_inversion_orden", "id_inversion", "nombre_activo"],
        ascending=[True, True, True, True, True],
        kind="mergesort",
    ).reset_index(drop=True)

    archivos = []
    for inversor, grupo in resultado.groupby("inversor", sort=True):
        detalle = grupo.copy().sort_values(
            ["mes_fecha", "fecha_inversion_orden", "id_inversion", "nombre_activo"],
            ascending=[True, True, True, True],
            kind="mergesort",
        )

        totales_mes = (
            detalle.groupby(["mes_fecha", "mes"], as_index=False)["interes_mes"]
            .sum()
            .sort_values("mes_fecha")
            .rename(columns={"interes_mes": "total_mes"})
        )
        totales_mes = totales_mes[["mes", "total_mes"]]

        # Capital total activo: NUEVA + REINVERSION activas a fecha de corte
        # Las reinversiones son capital real del inversor aunque no generen intereses propios en el extracto
        base_inversor = df_inv[df_inv["inversor"].astype(str).str.upper() == str(inversor).upper()].copy()
        base_inversor["tipo_op_norm"] = base_inversor["tipo_operacion"].astype(str).str.strip().str.upper()
        base_inversor["fecha_inversion"] = pd.to_datetime(base_inversor["fecha_inversion"], errors="coerce", dayfirst=True)
        base_inversor["fecha_final_inversion"] = pd.to_datetime(base_inversor["fecha_final_inversion"], errors="coerce", dayfirst=True)
        base_inversor["capital_invertido"] = pd.to_numeric(base_inversor["capital_invertido"], errors="coerce").fillna(0)
        fecha_corte_ts = pd.Timestamp(fecha_corte).normalize()
        activas_corte = base_inversor[
            (base_inversor["tipo_op_norm"].isin(["NUEVA", "REINVERSION"]))
            & (base_inversor["fecha_inversion"].notna())
            & (base_inversor["fecha_inversion"] <= fecha_corte_ts)
            & (
                base_inversor["fecha_final_inversion"].isna()
                | (base_inversor["fecha_final_inversion"] >= fecha_corte_ts)
            )
        ].copy()
        capital_total = float(activas_corte["capital_invertido"].sum()) if not activas_corte.empty else 0.0

        resumen = pd.DataFrame([{
            "inversor": inversor,
            "fecha_corte": fecha_corte.strftime("%d/%m/%Y"),
            "capital_total": round(capital_total, 2),
            "total_intereses_acumulados": round(detalle["interes_mes"].sum(), 2),
        }])

        detalle_exportar = detalle.drop(columns=["mes_fecha", "fecha_inversion_orden"], errors="ignore")
        # fecha_fin_op se mantiene en el Excel (col 11) para calcular capital activo al cierre mensual

        salida = BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
            totales_mes.to_excel(writer, sheet_name="TOTALES_MES", index=False)
            detalle_exportar.to_excel(writer, sheet_name="DETALLE", index=False)
        nombre_archivo = f"extracto_{str(inversor).upper().replace(' ', '_')}_{fecha_corte.strftime('%d%m%Y')}.xlsx"
        archivos.append((nombre_archivo, formatear_extracto_excel_bytes(salida.getvalue(), str(inversor), fecha_corte)))
    return archivos

def seccion_extractos():
    df_inv, _, _ = cargar_excel_completo()
    st.header("📤 Extractos")
    modo = st.radio("¿Qué quieres generar?", ["Todos", "Un inversor"], horizontal=True)
    inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
    inversor = st.selectbox("Inversor", inversores) if modo == "Un inversor" and inversores else None
    c1, c2 = st.columns(2)
    anio = int(c1.number_input("Año de corte", 2020, 2100, pd.Timestamp.today().year))
    mes = int(c2.number_input("Mes de corte", 1, 12, pd.Timestamp.today().month))
    if st.button("Generar extractos"):
        archivos = generar_extractos(df_inv, modo, inversor, anio, mes)
        if not archivos:
            st.warning("No se han generado extractos. Revisa el Excel o la fecha seleccionada.")
        elif len(archivos) == 1:
            nombre, contenido = archivos[0]
            st.success(f"Extracto generado: {nombre}")
            st.download_button("Descargar extracto", contenido, file_name=nombre, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for nombre, contenido in archivos:
                    zf.writestr(nombre, contenido)
            st.success(f"Se han generado {len(archivos)} extractos.")
            st.download_button("Descargar todos en ZIP", zip_buffer.getvalue(), file_name=f"extractos_{mes}_{anio}.zip", mime="application/zip")



# =========================
# GESTIÓN DE EXCEL DESDE LA APP
# =========================
def leer_todas_las_hojas_excel() -> dict:
    """Lee todas las hojas del archivo Excel para poder conservarlas al guardar."""
    try:
        hojas = pd.read_excel(ARCHIVO, sheet_name=None)
        return {str(nombre): df for nombre, df in hojas.items()}
    except Exception:
        return {}


def excel_hojas_a_bytes(hojas: dict) -> bytes:
    """Convierte un diccionario de hojas en un Excel descargable."""
    salida = BytesIO()
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        for nombre_hoja, df in hojas.items():
            nombre_limpio = str(nombre_hoja)[:31] if str(nombre_hoja).strip() else "Hoja"
            if df is None:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=nombre_limpio, index=False)
    return salida.getvalue()


def guardar_excel_completo_desde_hojas(hojas: dict):
    """Guarda todas las hojas en inversiones.xlsx y limpia la caché."""
    contenido = excel_hojas_a_bytes(hojas)
    with open(ARCHIVO, "wb") as f:
        f.write(contenido)
    st.cache_data.clear()


def aplicar_formula_simple(df: pd.DataFrame, operacion: str, columna_a: str, columna_b: str | None, nueva_columna: str) -> pd.DataFrame:
    """Aplica cálculos tipo Excel básicos sobre columnas numéricas."""
    out = df.copy()
    if not nueva_columna or not str(nueva_columna).strip():
        nueva_columna = "columna_calculada"
    nueva_columna = str(nueva_columna).strip()

    a = pd.to_numeric(out[columna_a], errors="coerce") if columna_a in out.columns else pd.Series(0, index=out.index)
    b = pd.to_numeric(out[columna_b], errors="coerce") if columna_b and columna_b in out.columns else pd.Series(0, index=out.index)

    if operacion == "Sumar A + B":
        out[nueva_columna] = a.fillna(0) + b.fillna(0)
    elif operacion == "Restar A - B":
        out[nueva_columna] = a.fillna(0) - b.fillna(0)
    elif operacion == "Multiplicar A x B":
        out[nueva_columna] = a.fillna(0) * b.fillna(0)
    elif operacion == "Dividir A / B":
        out[nueva_columna] = a / b.replace(0, pd.NA)
    elif operacion == "Porcentaje A sobre B":
        out[nueva_columna] = (a / b.replace(0, pd.NA)) * 100
    elif operacion == "Interés mensual: capital x interés / 12":
        if "capital_invertido" in out.columns and "interes_inversor_anual" in out.columns:
            capital = pd.to_numeric(out["capital_invertido"], errors="coerce").fillna(0)
            interes = pd.to_numeric(out["interes_inversor_anual"], errors="coerce").fillna(0)
            out[nueva_columna] = capital * interes / 12
        else:
            st.warning("Para esta fórmula necesitas las columnas capital_invertido e interes_inversor_anual.")
    elif operacion == "Interés nota mensual: capital x interés nota / 12":
        if "capital_invertido" in out.columns and "interes_nota_anual" in out.columns:
            capital = pd.to_numeric(out["capital_invertido"], errors="coerce").fillna(0)
            interes = pd.to_numeric(out["interes_nota_anual"], errors="coerce").fillna(0)
            out[nueva_columna] = capital * interes / 12
        else:
            st.warning("Para esta fórmula necesitas las columnas capital_invertido e interes_nota_anual.")
    return out


def mostrar_sumatorias_excel(df: pd.DataFrame):
    """Muestra sumatorias rápidas de columnas numéricas como apoyo tipo Excel."""
    if df is None or df.empty:
        return
    numericas = []
    for col in df.columns:
        serie = pd.to_numeric(df[col], errors="coerce")
        if serie.notna().sum() > 0:
            numericas.append(col)
    if not numericas:
        st.info("No hay columnas numéricas para sumar.")
        return
    seleccion = st.multiselect("Columnas para calcular sumatorias", numericas, default=numericas[: min(4, len(numericas))])
    if seleccion:
        cols = st.columns(min(4, len(seleccion)))
        for i, col in enumerate(seleccion):
            total = pd.to_numeric(df[col], errors="coerce").sum()
            cols[i % len(cols)].metric(f"Suma {col}", fmt(total))


def seccion_gestion_excel():
    st.markdown("## Gestión de Excel")
    st.caption("Sube, edita, calcula, guarda y descarga la base de datos directamente desde la app.")

    tab_subir, tab_editar, tab_descargar = st.tabs(["Subir Excel", "Editar y calcular", "Descargar copia"])

    with tab_subir:
        st.subheader("Recargar Excel desde Google Drive")
        st.info(
            "El Excel se lee desde Google Drive. "
            "Si has actualizado el archivo en Drive, pulsa el botón para recargar."
        )
        if st.button("🔄 Recargar Excel desde Google Drive", type="primary"):
            st.cache_data.clear()
            for k in list(st.session_state.keys()):
                if str(k).startswith("excel_editor_"):
                    del st.session_state[k]
            ok = descargar_excel_desde_drive()
            if ok:
                st.success("Excel recargado correctamente desde Google Drive.")
            st.rerun()

        st.markdown("---")
        st.subheader("O sube un Excel manualmente")
        st.caption("Si prefieres subir el archivo directamente, recuerda actualizarlo también en Google Drive para que sea permanente.")
        archivo_subido = st.file_uploader("Sube el archivo actualizado", type=["xlsx"])
        if archivo_subido is not None:
            nombre = archivo_subido.name.strip()
            if nombre != ARCHIVO:
                st.warning(f"El archivo se llama '{nombre}'. El sistema trabaja con '{ARCHIVO}'. Se guardará igualmente como {ARCHIVO}.")
            if st.button("Reemplazar Excel actual", type="primary"):
                with open(ARCHIVO, "wb") as f:
                    f.write(archivo_subido.read())
                st.cache_data.clear()
                for k in list(st.session_state.keys()):
                    if str(k).startswith("excel_editor_"):
                        del st.session_state[k]
                st.success("Excel actualizado. Recuerda actualizarlo también en Google Drive para que sea permanente.")
                st.rerun()

    with tab_editar:
        hojas = leer_todas_las_hojas_excel()
        if not hojas:
            st.error("No se ha podido leer el Excel actual.")
            return

        hoja = st.selectbox("Selecciona la hoja que quieres editar", list(hojas.keys()))
        editor_key = f"excel_editor_{hoja}"

        c1, c2 = st.columns([1, 1])
        if editor_key not in st.session_state:
            st.session_state[editor_key] = hojas[hoja].copy()
        if c1.button("Recargar hoja desde el Excel"):
            st.session_state[editor_key] = hojas[hoja].copy()
            st.rerun()
        if c2.button("Limpiar caché de datos"):
            st.cache_data.clear()
            st.success("Caché limpiada.")

        st.info("Puedes editar celdas, añadir filas nuevas y después guardar los cambios en el Excel.")
        df_editado = st.data_editor(
            st.session_state[editor_key],
            use_container_width=True,
            num_rows="dynamic",
            key=f"data_editor_{hoja}",
        )
        st.session_state[editor_key] = df_editado

        with st.expander("Sumatorias rápidas", expanded=True):
            mostrar_sumatorias_excel(df_editado)

        with st.expander("Añadir columna calculada tipo fórmula", expanded=False):
            columnas = list(df_editado.columns)
            columnas_numericas = [c for c in columnas if pd.to_numeric(df_editado[c], errors="coerce").notna().sum() > 0]
            if not columnas_numericas:
                st.info("No hay columnas numéricas disponibles para crear fórmulas.")
            else:
                operacion = st.selectbox(
                    "Fórmula",
                    [
                        "Sumar A + B",
                        "Restar A - B",
                        "Multiplicar A x B",
                        "Dividir A / B",
                        "Porcentaje A sobre B",
                        "Interés mensual: capital x interés / 12",
                        "Interés nota mensual: capital x interés nota / 12",
                    ],
                )
                c1, c2, c3 = st.columns(3)
                columna_a = c1.selectbox("Columna A", columnas_numericas)
                columna_b = c2.selectbox("Columna B", columnas_numericas) if operacion not in ["Interés mensual: capital x interés / 12", "Interés nota mensual: capital x interés nota / 12"] else None
                nueva_columna = c3.text_input("Nombre nueva columna", value="columna_calculada")
                if st.button("Aplicar fórmula a la hoja"):
                    st.session_state[editor_key] = aplicar_formula_simple(df_editado, operacion, columna_a, columna_b, nueva_columna)
                    st.success("Fórmula aplicada. Revisa la nueva columna en la tabla.")
                    st.rerun()

        c1, c2 = st.columns(2)
        if c1.button("Guardar cambios en inversiones.xlsx", type="primary"):
            hojas_actualizadas = leer_todas_las_hojas_excel()
            hojas_actualizadas[hoja] = st.session_state[editor_key].copy()
            guardar_excel_completo_desde_hojas(hojas_actualizadas)
            st.success("Cambios guardados en el Excel de la app.")
            st.rerun()

        hojas_para_descargar = leer_todas_las_hojas_excel()
        hojas_para_descargar[hoja] = st.session_state[editor_key].copy()
        c2.download_button(
            "Descargar Excel con estos cambios",
            data=excel_hojas_a_bytes(hojas_para_descargar),
            file_name=ARCHIVO,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with tab_descargar:
        st.subheader("Descargar copia actual")
        hojas = leer_todas_las_hojas_excel()
        if hojas:
            st.download_button(
                "Descargar inversiones.xlsx",
                data=excel_hojas_a_bytes(hojas),
                file_name=ARCHIVO,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.error("No se ha podido preparar la descarga.")

# =========================
# APP FINAL
# =========================
mostrar_hero(st.session_state.usuario)

try:
    df_inv, df_cal, df_control = cargar_excel_completo()
except Exception as e:
    st.error("No se ha podido cargar inversiones.xlsx. Revisa que el archivo esté subido a GitHub y que las hojas existan.")
    with st.expander("Ver detalle técnico"):
        st.exception(e)
    st.stop()

menu = st.sidebar.selectbox(
    "Menú principal",
    [
        "Dashboard financiero", "Histórico y proyecciones", "Centro de control", "Consultas Fútbol", "Consultas Notas", "Consultas Paraguay", "Consultas MotoClick",
        "Notas estructuradas", "Alertas y calendario", "Sistema Fondo", "Extractos", "Gestión de Excel", "Calidad de datos", "Base de datos",
    ],
)

if menu == "Dashboard financiero":
    dashboard_financiero()
elif menu == "Histórico y proyecciones":
    seccion_historico_y_proyecciones()
elif menu == "Centro de control":
    centro_control_inversiones()
elif menu == "Consultas Fútbol":
    seccion_activo("Fútbol", "futbol", TASA_ANUAL_FUTBOL)
elif menu == "Consultas Notas":
    seccion_notas()
elif menu == "Consultas Paraguay":
    seccion_activo("Paraguay", "paraguay", TASA_ANUAL_PARAGUAY, incluir_ingresado_desde_inicio=True)
elif menu == "Consultas MotoClick":
    seccion_activo("MotoClick", "motoclick", TASA_ANUAL_MOTOCLICK)
elif menu == "Notas estructuradas":
    seccion_notas_archivo()
elif menu == "Alertas y calendario":
    panel_alertas_y_calendario()
elif menu == "Sistema Fondo":
    seccion_sistema_fondo()
elif menu == "Extractos":
    seccion_extractos()
elif menu == "Gestión de Excel":
    seccion_gestion_excel()
elif menu == "Calidad de datos":
    panel_calidad_datos()
elif menu == "Base de datos":
    st.markdown("## Base de datos")
    hojas = {"INVERSIONES": df_inv, "CALENDARIO_NOTAS": df_cal, "CONTROL_NOTAS": df_control}
    hoja = st.selectbox("Selecciona hoja", list(hojas.keys()))
    st.dataframe(hojas[hoja], use_container_width=True)
