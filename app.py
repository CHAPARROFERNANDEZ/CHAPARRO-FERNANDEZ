import os
import calendar
from datetime import timedelta

import pandas as pd
import streamlit as st

try:
    import yfinance as yf
except Exception:
    yf = None

ARCHIVO = "inversiones.xlsx"
HOJA_INVERSIONES = "INVERSIONES"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"
HOJA_CONTROL = "CONTROL_NOTAS"
HOJA_CALLS = "CALENDARIO_CALLS"
HOJA_RESULTADOS = "RESULTADOS_OBSERVACION"

TASAS_ACTIVOS = {
    "paraguay": 0.15,
    "motoclick": 0.25,
    "futbol": 0.15,
}

st.set_page_config(page_title="Sistema Fondo", layout="wide")

# =========================
# UTILIDADES
# =========================
def fmt(x):
    try:
        return f"{float(x):,.2f} €"
    except Exception:
        return "0,00 €"


def normalizar_columnas(df, upper=False):
    if upper:
        df.columns = [str(c).strip().upper().replace(" ", "_") for c in df.columns]
    else:
        df.columns = [str(c).strip().lower() for c in df.columns]
    return df


def nombre_mes_es(mes):
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    return meses.get(int(mes), str(mes))


def ultimo_dia_mes(anio, mes):
    return calendar.monthrange(int(anio), int(mes))[1]


def hoy():
    return pd.Timestamp.today().normalize()


def limpiar_texto(x):
    if pd.isna(x):
        return ""
    return str(x).strip().lower()


@st.cache_data(ttl=60)
def leer_excel(nombre_archivo):
    if not os.path.exists(nombre_archivo):
        raise FileNotFoundError(f"No encuentro el archivo {nombre_archivo}. Súbelo a GitHub en la misma carpeta que app.py")
    return pd.ExcelFile(nombre_archivo)


@st.cache_data(ttl=60)
def cargar_hoja(nombre_archivo, hoja, upper=False):
    xls = leer_excel(nombre_archivo)
    if hoja not in xls.sheet_names:
        raise ValueError(f"No existe la hoja {hoja} dentro de {nombre_archivo}. Hojas disponibles: {', '.join(xls.sheet_names)}")
    df = pd.read_excel(nombre_archivo, sheet_name=hoja)
    return normalizar_columnas(df, upper=upper)


def cargar_inversiones():
    df = cargar_hoja(ARCHIVO, HOJA_INVERSIONES).copy()

    if "unnamed: 6" in df.columns and "cuenta_cobro" not in df.columns:
        df = df.rename(columns={"unnamed: 6": "cuenta_cobro"})

    for col in [
        "id_inversion", "inversor", "tipo_inversion", "subtipo_inversion",
        "nombre_activo", "metodo_calculo", "activo_generador_interes",
        "tipo_operacion", "capital_nuevo_real", "cuenta_cobro", "motivo"
    ]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    for col in ["fecha_inversion", "fecha_final_inversion"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce", dayfirst=True)

    for col in ["capital_invertido", "interes_inversor_anual", "interes_nota_anual"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        else:
            df[col] = 0

    return df


def cargar_calendario():
    df = cargar_hoja(ARCHIVO, HOJA_CALENDARIO).copy()
    df["nota"] = pd.to_numeric(df.get("nota"), errors="coerce").astype("Int64")
    df["tipo_evento"] = df.get("tipo_evento", "").astype(str).str.strip().str.upper()
    df["fecha"] = pd.to_datetime(df.get("fecha"), errors="coerce", dayfirst=True).dt.normalize()
    return df.dropna(subset=["nota", "tipo_evento", "fecha"]).copy()


def cargar_control():
    df = cargar_hoja(ARCHIVO, HOJA_CONTROL).copy()
    df["nota"] = pd.to_numeric(df.get("nota"), errors="coerce").astype("Int64")
    df["ticker"] = df.get("ticker", "").astype(str).str.strip().str.upper()
    if "precio_compra" in df.columns:
        df["precio_compra"] = pd.to_numeric(df["precio_compra"], errors="coerce")
    if "barrera_cupon" in df.columns:
        df["barrera_cupon"] = pd.to_numeric(df["barrera_cupon"], errors="coerce")
    if "contingency" in df.columns:
        df["contingency"] = pd.to_numeric(df["contingency"], errors="coerce")
    return df


def cargar_calls():
    try:
        df = cargar_hoja(ARCHIVO, HOJA_CALLS).copy()
    except Exception:
        return pd.DataFrame()
    df["nota"] = pd.to_numeric(df.get("nota"), errors="coerce").astype("Int64")
    df["fecha_call"] = pd.to_datetime(df.get("fecha_call"), errors="coerce", dayfirst=True).dt.normalize()
    if "estado" in df.columns:
        df["estado"] = df["estado"].fillna("").astype(str).str.strip()
    if "observaciones" in df.columns:
        df["observaciones"] = df["observaciones"].fillna("").astype(str).str.strip()
    return df


def detectar_activo(row):
    tipo = limpiar_texto(row.get("tipo_inversion", ""))
    subtipo = limpiar_texto(row.get("subtipo_inversion", ""))
    nombre = limpiar_texto(row.get("nombre_activo", ""))
    if tipo == "nota" or nombre.startswith("nota"):
        return "notas"
    for activo in TASAS_ACTIVOS:
        if activo in subtipo or activo in nombre:
            return activo
    return "otros"


def inversiones_activas(df, fecha=None):
    if fecha is None:
        fecha = hoy()
    fecha = pd.Timestamp(fecha).normalize()
    trabajo = df.copy()
    if "motivo" in trabajo.columns:
        trabajo = trabajo[trabajo["motivo"].apply(limpiar_texto) != "call"].copy()
    return trabajo[
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= fecha) &
        (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha))
    ].copy()


def extraer_numero_nota(nombre_activo):
    import re
    if pd.isna(nombre_activo):
        return pd.NA
    m = re.search(r"NOTA[_\s]?(\d+)", str(nombre_activo).strip().upper())
    return int(m.group(1)) if m else pd.NA


def filtrar_notas(df):
    trabajo = df.copy()
    if "tipo_inversion" in trabajo.columns:
        trabajo = trabajo[trabajo["tipo_inversion"].apply(limpiar_texto) == "nota"].copy()
    trabajo["nota_num"] = trabajo["nombre_activo"].apply(extraer_numero_nota)
    trabajo["nota_num"] = pd.to_numeric(trabajo["nota_num"], errors="coerce").astype("Int64")
    if "activo_generador_interes" in trabajo.columns:
        trabajo = trabajo[trabajo["activo_generador_interes"].apply(limpiar_texto) == "si"].copy()
    return trabajo


def dias_activos_en_mes(fecha_inicio, fecha_fin, anio, mes):
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


def detalle_activo_fijo_mes(df, nombre_activo, anio, mes):
    tasa = TASAS_ACTIVOS[nombre_activo]
    trabajo = df.copy()
    trabajo = trabajo[
        trabajo.get("subtipo_inversion", "").astype(str).str.lower().str.contains(nombre_activo, na=False) |
        trabajo.get("nombre_activo", "").astype(str).str.lower().str.contains(nombre_activo, na=False)
    ].copy()
    dias_mes = ultimo_dia_mes(anio, mes)
    resultados = []
    for _, fila in trabajo.iterrows():
        dias = dias_activos_en_mes(fila["fecha_inversion"], fila["fecha_final_inversion"], anio, mes)
        if dias == 0:
            continue
        proporcion = dias / dias_mes
        capital = fila["capital_invertido"]
        cobro_empresa = capital * tasa / 12 * proporcion
        pago_inversor = capital * fila["interes_inversor_anual"] / 12 * proporcion
        resultados.append({
            "activo": nombre_activo,
            "id_inversion": fila.get("id_inversion", ""),
            "inversor": fila.get("inversor", ""),
            "capital": capital,
            "dias_activos": dias,
            "cobro_empresa": cobro_empresa,
            "pago_inversor": pago_inversor,
            "beneficio": cobro_empresa - pago_inversor,
        })
    return pd.DataFrame(resultados)


def inversiones_activas_nota(df, nota, fecha_pago):
    trabajo = filtrar_notas(df)
    fecha_pago = pd.Timestamp(fecha_pago).normalize()
    return trabajo[
        (trabajo["nota_num"] == nota) &
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= fecha_pago) &
        (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha_pago))
    ].copy()


def resumen_notas_mes(df_inv, df_cal, anio, mes):
    pagos = df_cal[
        (df_cal["tipo_evento"] == "PAGO") &
        (df_cal["fecha"].dt.year == anio) &
        (df_cal["fecha"].dt.month == mes)
    ].copy()
    resultados = []
    for _, evento in pagos.iterrows():
        nota = evento["nota"]
        fecha_pago = evento["fecha"]
        activas = inversiones_activas_nota(df_inv, nota, fecha_pago)
        for _, fila in activas.iterrows():
            capital = fila["capital_invertido"]
            cobro_empresa = capital * fila["interes_nota_anual"] / 12
            pago_inversor = capital * fila["interes_inversor_anual"] / 12
            resultados.append({
                "fecha_pago": fecha_pago,
                "nota": int(nota),
                "id_inversion": fila.get("id_inversion", ""),
                "inversor": fila.get("inversor", ""),
                "capital": capital,
                "cobro_empresa": cobro_empresa,
                "pago_inversor": pago_inversor,
                "beneficio": cobro_empresa - pago_inversor,
            })
    detalle = pd.DataFrame(resultados)
    if detalle.empty:
        return 0, 0, 0, detalle
    return detalle["cobro_empresa"].sum(), detalle["pago_inversor"].sum(), detalle["beneficio"].sum(), detalle


def preparar_df_mostrar(df):
    if df is None or df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if "fecha" in col:
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d/%m/%Y")
    for col in ["capital", "capital_invertido", "cobro_empresa", "pago_inversor", "beneficio", "cobro_mes"]:
        if col in out.columns:
            out[col] = out[col].apply(fmt)
    return out


def obtener_precio_actual(ticker):
    if yf is None:
        return None
    try:
        data = yf.Ticker(ticker)
        hist = data.history(period="5d")
        if hist.empty:
            return None
        return round(float(hist["Close"].dropna().iloc[-1]), 2)
    except Exception:
        return None


# =========================
# INTERFAZ
# =========================
st.title("📊 Sistema Fondo")
st.caption("Aplicación conectada al Excel inversiones.xlsx")

try:
    xls = leer_excel(ARCHIVO)
except Exception as e:
    st.error(str(e))
    st.stop()

with st.sidebar:
    st.header("Menú")
    seccion = st.selectbox(
        "Selecciona una sección",
        [
            "Inicio",
            "Panel global",
            "Alertas notas",
            "Alertas semana",
            "Calendario notas",
            "Consultas notas",
            "Consultas Paraguay",
            "Consultas Motoclick",
            "Consultas Fútbol",
            "Sistema fondo",
            "Ver hojas del Excel",
        ]
    )

try:
    df_inv = cargar_inversiones()
    df_cal = cargar_calendario()
except Exception as e:
    st.error(f"Error cargando datos principales: {e}")
    st.stop()

if seccion == "Inicio":
    st.success("Sistema cargado correctamente.")
    st.write("Hojas detectadas en el Excel:", xls.sheet_names)
    st.write("Usa el menú lateral para consultar el sistema.")

elif seccion == "Panel global":
    st.subheader("Panel global del fondo")
    activas = inversiones_activas(df_inv)
    activas["activo"] = activas.apply(detectar_activo, axis=1) if not activas.empty else []

    c1, c2, c3 = st.columns(3)
    c1.metric("Capital activo total", fmt(activas["capital_invertido"].sum() if not activas.empty else 0))
    c2.metric("Número de inversiones activas", len(activas))
    c3.metric("Fecha cálculo", hoy().strftime("%d/%m/%Y"))

    st.subheader("Capital activo por activo")
    if activas.empty:
        st.info("No hay inversiones activas.")
    else:
        resumen = activas.groupby("activo", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"})
        st.dataframe(preparar_df_mostrar(resumen), use_container_width=True)

    st.subheader("Próximos pagos")
    pagos = df_cal[(df_cal["tipo_evento"] == "PAGO") & (df_cal["fecha"] >= hoy())].sort_values("fecha").head(10)
    st.dataframe(preparar_df_mostrar(pagos), use_container_width=True)

elif seccion == "Alertas notas":
    st.subheader("Alertas de notas")
    try:
        control = cargar_control()
        if "ticker" not in control.columns:
            st.error("La hoja CONTROL_NOTAS no tiene columna ticker.")
        else:
            control = control.copy()
            precio_compra_col = "precio_compra" if "precio_compra" in control.columns else None
            barrera_col = "contingency" if "contingency" in control.columns else ("barrera_cupon" if "barrera_cupon" in control.columns else None)
            if not precio_compra_col or not barrera_col:
                st.error("CONTROL_NOTAS necesita precio_compra y contingency/barrera_cupon.")
            else:
                if st.button("Actualizar precios y calcular alertas"):
                    control["precio_actual"] = control["ticker"].apply(obtener_precio_actual)
                    control["barrera"] = control[precio_compra_col] * control[barrera_col].apply(lambda x: x / 100 if pd.notna(x) and x > 1 else x)
                    control["estado"] = control.apply(lambda r: "OK" if pd.notna(r["precio_actual"]) and r["precio_actual"] >= r["barrera"] else "RIESGO", axis=1)
                    st.dataframe(control, use_container_width=True)
                    riesgo = control[control["estado"] == "RIESGO"]
                    if riesgo.empty:
                        st.success("Ninguna nota en riesgo.")
                    else:
                        st.error("Hay notas en riesgo.")
                        st.dataframe(riesgo, use_container_width=True)
                else:
                    st.info("Pulsa el botón para actualizar precios con Yahoo Finance.")
    except Exception as e:
        st.error(f"Error en Alertas notas: {e}")

elif seccion == "Alertas semana":
    st.subheader("Novedades de la semana")
    fecha_base = st.date_input("Fecha inicio", value=hoy().date())
    inicio = pd.Timestamp(fecha_base).normalize()
    fin = inicio + timedelta(days=6)
    eventos = df_cal[(df_cal["fecha"] >= inicio) & (df_cal["fecha"] <= fin)].sort_values(["fecha", "tipo_evento", "nota"])
    st.write(f"Del {inicio.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}")
    if eventos.empty:
        st.info("No hay observaciones ni pagos en este periodo.")
    else:
        st.dataframe(preparar_df_mostrar(eventos), use_container_width=True)

elif seccion == "Calendario notas":
    st.subheader("Calendario de notas")
    tipo = st.selectbox("Tipo", ["Todos", "PAGO", "OBSERVACION"])
    anio = st.number_input("Año", min_value=2020, max_value=2100, value=hoy().year, step=1)
    mes = st.number_input("Mes", min_value=1, max_value=12, value=hoy().month, step=1)
    inicio = pd.Timestamp(int(anio), int(mes), 1)
    fin = pd.Timestamp(int(anio), int(mes), ultimo_dia_mes(int(anio), int(mes)))
    eventos = df_cal[(df_cal["fecha"] >= inicio) & (df_cal["fecha"] <= fin)].copy()
    if tipo != "Todos":
        eventos = eventos[eventos["tipo_evento"] == tipo]
    eventos = eventos.sort_values(["fecha", "nota", "tipo_evento"])
    st.dataframe(preparar_df_mostrar(eventos), use_container_width=True)

elif seccion == "Consultas notas":
    st.subheader("Consultas de notas")
    anio = st.number_input("Año", min_value=2020, max_value=2100, value=hoy().year, step=1, key="notas_anio")
    mes = st.number_input("Mes", min_value=1, max_value=12, value=hoy().month, step=1, key="notas_mes")
    c, p, b, detalle = resumen_notas_mes(df_inv, df_cal, int(anio), int(mes))
    c1, c2, c3 = st.columns(3)
    c1.metric("Cobro compañía", fmt(c))
    c2.metric("Pago inversores", fmt(p))
    c3.metric("Beneficio", fmt(b))
    st.dataframe(preparar_df_mostrar(detalle), use_container_width=True)

elif seccion in ["Consultas Paraguay", "Consultas Motoclick", "Consultas Fútbol"]:
    activo = {"Consultas Paraguay": "paraguay", "Consultas Motoclick": "motoclick", "Consultas Fútbol": "futbol"}[seccion]
    st.subheader(seccion)
    anio = st.number_input("Año", min_value=2020, max_value=2100, value=hoy().year, step=1, key=f"{activo}_anio")
    mes = st.number_input("Mes", min_value=1, max_value=12, value=hoy().month, step=1, key=f"{activo}_mes")
    detalle = detalle_activo_fijo_mes(df_inv, activo, int(anio), int(mes))
    c = detalle["cobro_empresa"].sum() if not detalle.empty else 0
    p = detalle["pago_inversor"].sum() if not detalle.empty else 0
    b = detalle["beneficio"].sum() if not detalle.empty else 0
    c1, c2, c3 = st.columns(3)
    c1.metric("Cobro compañía", fmt(c))
    c2.metric("Pago inversores", fmt(p))
    c3.metric("Beneficio", fmt(b))
    st.dataframe(preparar_df_mostrar(detalle), use_container_width=True)

elif seccion == "Sistema fondo":
    st.subheader("Sistema fondo")
    consulta = st.selectbox("Consulta", ["Capital activo por inversor", "Capital activo por activo", "Calls próximos", "Calls vencidos"])
    if consulta == "Capital activo por inversor":
        activas = inversiones_activas(df_inv)
        resumen = activas.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
        st.dataframe(preparar_df_mostrar(resumen), use_container_width=True)
    elif consulta == "Capital activo por activo":
        activas = inversiones_activas(df_inv)
        activas["activo"] = activas.apply(detectar_activo, axis=1) if not activas.empty else []
        resumen = activas.groupby("activo", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
        st.dataframe(preparar_df_mostrar(resumen), use_container_width=True)
    elif consulta == "Calls próximos":
        calls = cargar_calls()
        if calls.empty:
            st.info("No hay hoja CALENDARIO_CALLS o no hay calls.")
        else:
            resultado = calls[(calls["fecha_call"].notna()) & (calls["fecha_call"] >= hoy())].sort_values(["fecha_call", "nota"])
            st.dataframe(preparar_df_mostrar(resultado), use_container_width=True)
    elif consulta == "Calls vencidos":
        calls = cargar_calls()
        if calls.empty:
            st.info("No hay hoja CALENDARIO_CALLS o no hay calls.")
        else:
            resultado = calls[(calls["fecha_call"].notna()) & (calls["fecha_call"] < hoy())].copy()
            if "estado" in resultado.columns:
                resultado = resultado[~resultado["estado"].apply(limpiar_texto).isin(["hecho", "realizado", "ejecutado", "call ejecutado"])]
            st.dataframe(preparar_df_mostrar(resultado), use_container_width=True)

elif seccion == "Ver hojas del Excel":
    st.subheader("Hojas del Excel")
    hoja = st.selectbox("Elige hoja", xls.sheet_names)
    df = pd.read_excel(ARCHIVO, sheet_name=hoja)
    st.dataframe(df, use_container_width=True)
