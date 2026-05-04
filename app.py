import calendar
import re
from datetime import datetime
from typing import Optional, Tuple

import pandas as pd
import streamlit as st

try:
    import yfinance as yf
except Exception:
    yf = None

st.set_page_config(page_title="Sistema Fondo", layout="wide")

ARCHIVO = "inversiones.xlsx"
HOJA_INVERSIONES = "INVERSIONES"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"
HOJA_CONTROL = "CONTROL_NOTAS"

TASA_ANUAL_FUTBOL = 0.15
TASA_ANUAL_MOTOCLICK = 0.25
TASA_ANUAL_PARAGUAY = 0.15


def fmt(x):
    try:
        return f"{float(x):,.2f} €"
    except Exception:
        return "0.00 €"


def nombre_mes_es(mes: int) -> str:
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
    }
    return meses.get(int(mes), str(mes))


def ultimo_dia_mes(anio: int, mes: int) -> int:
    return calendar.monthrange(int(anio), int(mes))[1]


def limpiar_texto(x) -> str:
    if pd.isna(x):
        return ""
    return str(x).strip().lower()


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

    for col in [
        "id_inversion", "inversor", "tipo_inversion", "subtipo_inversion",
        "nombre_activo", "metodo_calculo", "activo_generador_interes",
        "tipo_operacion", "capital_nuevo_real", "cuenta_cobro",
    ]:
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
        for col in ["precio_compra", "barrera_cupon", "contingency"]:
            if col in control.columns:
                control[col] = pd.to_numeric(control[col], errors="coerce")

    return inv, cal, control


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
            "id_inversion": fila.get("id_inversion", ""),
            "inversor": fila.get("inversor", ""),
            "capital_invertido": capital,
            "fecha_inversion": fila.get("fecha_inversion"),
            "fecha_final_inversion": fila.get("fecha_final_inversion"),
            "dias_activos": dias,
            "dias_mes": dias_mes,
            "ingreso_bruto": ingreso_bruto,
            "pago_inversor_mes": pago_inversor,
            "beneficio_empresa_mes": ingreso_bruto - pago_inversor,
        })
    return pd.DataFrame(filas)


def capital_activo_en_fecha(df_base: pd.DataFrame, fecha_consulta, activo: Optional[str] = None, solo_real: bool = False) -> float:
    fecha_consulta = pd.Timestamp(fecha_consulta).normalize()
    trabajo = df_base.copy()
    if activo:
        trabajo = filtrar_activo(trabajo, activo)
    if solo_real and "capital_nuevo_real" in trabajo.columns:
        trabajo = trabajo[trabajo["capital_nuevo_real"].astype(str).str.lower() == "si"].copy()

    filtrado = trabajo[
        (trabajo["fecha_inversion"].notna())
        & (trabajo["fecha_inversion"] <= fecha_consulta)
        & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha_consulta))
    ]
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
    return trabajo[
        (trabajo["nota_num"] == nota)
        & (trabajo["fecha_inversion"].notna())
        & (trabajo["fecha_inversion"] <= fecha_pago)
        & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= fecha_pago))
    ].copy()


def pagos_notas_mes(df_cal: pd.DataFrame, anio: int, mes: int) -> pd.DataFrame:
    if df_cal.empty:
        return pd.DataFrame()
    return df_cal[
        (df_cal["tipo_evento"] == "PAGO")
        & (df_cal["fecha"].notna())
        & (df_cal["fecha"].dt.year == anio)
        & (df_cal["fecha"].dt.month == mes)
    ].copy().sort_values(["fecha", "nota"])


def pagos_notas_hasta_hoy(df_cal: pd.DataFrame) -> pd.DataFrame:
    hoy = pd.Timestamp.today().normalize()
    return df_cal[
        (df_cal["tipo_evento"] == "PAGO")
        & (df_cal["fecha"].notna())
        & (df_cal["fecha"] <= hoy)
    ].copy().sort_values(["fecha", "nota"])


def preparar_detalle_notas(df_inv: pd.DataFrame, df_pagos: pd.DataFrame) -> pd.DataFrame:
    filas = []
    for _, evento in df_pagos.iterrows():
        nota = evento.get("nota")
        fecha_pago = evento.get("fecha")
        if pd.isna(nota) or pd.isna(fecha_pago):
            continue
        activas = inversiones_activas_para_nota(df_inv, int(nota), fecha_pago)
        for _, fila in activas.iterrows():
            capital = float(fila.get("capital_invertido", 0))
            cobro_compania = capital * float(fila.get("interes_nota_anual", 0)) / 12
            pago_inversor = capital * float(fila.get("interes_inversor_anual", 0)) / 12
            filas.append({
                "fecha_pago": fecha_pago,
                "nota": int(nota),
                "id_inversion": fila.get("id_inversion", ""),
                "inversor": fila.get("inversor", ""),
                "cuenta_cobro": fila.get("cuenta_cobro", "SIN CLASIFICAR"),
                "capital_invertido": capital,
                "interes_nota_anual": fila.get("interes_nota_anual", 0),
                "interes_inversor_anual": fila.get("interes_inversor_anual", 0),
                "cobro_compania": cobro_compania,
                "pago_inversor": pago_inversor,
                "beneficio_empresa": cobro_compania - pago_inversor,
            })
    return pd.DataFrame(filas)


def resumen_notas_mes(df_inv: pd.DataFrame, df_cal: pd.DataFrame, anio: int, mes: int):
    pagos = pagos_notas_mes(df_cal, anio, mes)
    detalle = preparar_detalle_notas(df_inv, pagos)
    if detalle.empty:
        return 0.0, 0.0, 0.0, detalle, pagos
    return (
        float(detalle["cobro_compania"].sum()),
        float(detalle["pago_inversor"].sum()),
        float(detalle["beneficio_empresa"].sum()),
        detalle,
        pagos,
    )


def resumen_por_cuenta_cobro(detalle: pd.DataFrame) -> pd.DataFrame:
    if detalle.empty:
        return pd.DataFrame(columns=["cuenta_cobro", "cobro_compania"])
    return detalle.groupby("cuenta_cobro", as_index=False)["cobro_compania"].sum().sort_values("cobro_compania", ascending=False)


def resumen_capital_por_inversor_notas(df_inv: pd.DataFrame, solo_activo: bool = False) -> pd.DataFrame:
    trabajo = filtrar_notas(df_inv)
    hoy = pd.Timestamp.today().normalize()
    if solo_activo:
        trabajo = trabajo[
            (trabajo["fecha_inversion"].notna())
            & (trabajo["fecha_inversion"] <= hoy)
            & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= hoy))
        ]
    if trabajo.empty:
        return pd.DataFrame(columns=["inversor", "capital"])
    return trabajo.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)


def proximo_evento_nota(df_cal: pd.DataFrame, nota: int, tipo: str):
    hoy = pd.Timestamp.today().normalize()
    eventos = df_cal[
        (df_cal["tipo_evento"] == tipo)
        & (df_cal["nota"] == nota)
        & (df_cal["fecha"].notna())
        & (df_cal["fecha"] >= hoy)
    ].sort_values("fecha")
    return None if eventos.empty else eventos.iloc[0]["fecha"]


def mostrar_metricas(titulo, valores):
    st.subheader(titulo)
    cols = st.columns(len(valores))
    for col, (label, value) in zip(cols, valores):
        col.metric(label, value)


def preparar_tabla_monetaria(df: pd.DataFrame, columnas_monetarias) -> pd.DataFrame:
    if df.empty:
        return df
    out = df.copy()
    for col in out.columns:
        if "fecha" in col:
            out[col] = pd.to_datetime(out[col], errors="coerce").dt.strftime("%d/%m/%Y")
    for col in columnas_monetarias:
        if col in out.columns:
            out[col] = out[col].map(fmt)
    return out


def seccion_activo(nombre_visible: str, activo_key: str, tasa_anual: float, incluir_ingresado_desde_inicio: bool = False):
    df_inv, _, _ = cargar_excel_completo()
    st.header(f"📌 Consultas {nombre_visible}")

    consulta = st.selectbox(
        "Elige una pregunta",
        [
            f"¿Cuánto ingresará {nombre_visible} en un mes?",
            "¿Cuánto cobrará cada inversor ese mes?",
            "¿Cuánto cobrará un inversor concreto ese mes?",
            "¿Cuál será el beneficio de la empresa ese mes?",
            "¿Cuál es el total pagado a inversores desde el inicio?",
            f"¿Cuánto capital hay actualmente activo en {nombre_visible} hoy?",
            f"¿Cuánto capital había activo en {nombre_visible} en un mes concreto?",
        ] + (["¿Cuánto ha ingresado la compañía desde el inicio?", "¿Cuál es el beneficio total acumulado desde el inicio?"] if incluir_ingresado_desde_inicio else []),
    )

    necesita_mes = consulta in [
        f"¿Cuánto ingresará {nombre_visible} en un mes?",
        "¿Cuánto cobrará cada inversor ese mes?",
        "¿Cuánto cobrará un inversor concreto ese mes?",
        "¿Cuál será el beneficio de la empresa ese mes?",
        f"¿Cuánto capital había activo en {nombre_visible} en un mes concreto?",
    ]

    anio = mes = None
    if necesita_mes:
        c1, c2 = st.columns(2)
        anio = int(c1.number_input("Año", min_value=2020, max_value=2100, value=pd.Timestamp.today().year, key=f"{activo_key}_anio"))
        mes = int(c2.number_input("Mes", min_value=1, max_value=12, value=pd.Timestamp.today().month, key=f"{activo_key}_mes"))

    nombre_inversor = None
    if consulta == "¿Cuánto cobrará un inversor concreto ese mes?":
        inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        nombre_inversor = st.selectbox("Inversor", inversores) if inversores else st.text_input("Inversor")

    if st.button("Calcular", key=f"calc_{activo_key}_{consulta}"):
        if consulta == f"¿Cuánto ingresará {nombre_visible} en un mes?":
            detalle = detalle_activo_mes(df_inv, activo_key, tasa_anual, anio, mes)
            total = detalle["ingreso_bruto"].sum() if not detalle.empty else 0
            mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Ingreso bruto", fmt(total))])
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
            if detalle.empty:
                total = 0
            else:
                filtrado = detalle[detalle["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()]
                total = filtrado["pago_inversor_mes"].sum() if not filtrado.empty else 0
            mostrar_metricas("Resultado", [(f"Cobro de {nombre_inversor}", fmt(total))])

        elif consulta == "¿Cuál será el beneficio de la empresa ese mes?":
            detalle = detalle_activo_mes(df_inv, activo_key, tasa_anual, anio, mes)
            total = detalle["beneficio_empresa_mes"].sum() if not detalle.empty else 0
            mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Beneficio empresa", fmt(total))])
            if not detalle.empty:
                st.dataframe(preparar_tabla_monetaria(detalle, ["capital_invertido", "ingreso_bruto", "pago_inversor_mes", "beneficio_empresa_mes"]), use_container_width=True)

        elif consulta == "¿Cuál es el total pagado a inversores desde el inicio?":
            total = total_pagado_activo_desde_inicio(df_inv, activo_key, tasa_anual)
            mostrar_metricas("Resultado", [("Total pagado", fmt(total))])

        elif consulta == f"¿Cuánto capital hay actualmente activo en {nombre_visible} hoy?":
            bruto = capital_activo_en_fecha(df_inv, pd.Timestamp.today(), activo_key, solo_real=False)
            real = capital_activo_en_fecha(df_inv, pd.Timestamp.today(), activo_key, solo_real=True)
            mostrar_metricas("Resultado", [("Capital activo", fmt(bruto)), ("Capital activo real", fmt(real))])

        elif consulta == f"¿Cuánto capital había activo en {nombre_visible} en un mes concreto?":
            fecha = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))
            bruto = capital_activo_en_fecha(df_inv, fecha, activo_key, solo_real=False)
            real = capital_activo_en_fecha(df_inv, fecha, activo_key, solo_real=True)
            mostrar_metricas(f"Cierre {nombre_mes_es(mes)} {anio}", [("Capital activo", fmt(bruto)), ("Capital activo real", fmt(real))])

        elif consulta == "¿Cuánto ha ingresado la compañía desde el inicio?":
            total = total_ingresado_activo_desde_inicio(df_inv, activo_key, tasa_anual)
            mostrar_metricas("Resultado", [("Total ingresado", fmt(total))])

        elif consulta == "¿Cuál es el beneficio total acumulado desde el inicio?":
            ingreso = total_ingresado_activo_desde_inicio(df_inv, activo_key, tasa_anual)
            pagado = total_pagado_activo_desde_inicio(df_inv, activo_key, tasa_anual)
            mostrar_metricas("Resultado", [("Beneficio acumulado", fmt(ingreso - pagado))])


def seccion_notas():
    df_inv, df_cal, _ = cargar_excel_completo()
    st.header("🧾 Consultas Notas")

    consulta = st.selectbox(
        "Elige una pregunta",
        [
            "¿Cuánto cobrará la compañía en un mes de notas?",
            "¿Cuánto se pagará a inversores en un mes de notas?",
            "¿Cuál será el beneficio de la empresa en un mes de notas?",
            "¿Cuánto cobrará cada inversor ese mes?",
            "¿Cuánto cobrará un inversor concreto ese mes?",
            "¿Cuánto ha cobrado la compañía desde el inicio?",
            "¿Cuánto se ha pagado a inversores desde el inicio?",
            "¿Cuál es el beneficio total desde el inicio?",
            "¿Cuál es el próximo pago de una nota?",
            "¿Cuál es la próxima observación de una nota?",
            "¿Cuánto capital hay invertido en total?",
            "¿Cuánto capital hay actualmente activo?",
            "¿Cuánto capital tiene un inversor?",
            "¿Cuánto capital activo tiene un inversor?",
            "Ver ranking de capital por inversor",
            "Ver ranking de capital activo",
        ],
    )

    consultas_mes = [
        "¿Cuánto cobrará la compañía en un mes de notas?",
        "¿Cuánto se pagará a inversores en un mes de notas?",
        "¿Cuál será el beneficio de la empresa en un mes de notas?",
        "¿Cuánto cobrará cada inversor ese mes?",
        "¿Cuánto cobrará un inversor concreto ese mes?",
    ]

    anio = mes = None
    if consulta in consultas_mes:
        c1, c2 = st.columns(2)
        anio = int(c1.number_input("Año", min_value=2020, max_value=2100, value=pd.Timestamp.today().year, key="notas_anio"))
        mes = int(c2.number_input("Mes", min_value=1, max_value=12, value=pd.Timestamp.today().month, key="notas_mes"))

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
            total_cobrado, total_pagado, total_beneficio, detalle, pagos = resumen_notas_mes(df_inv, df_cal, anio, mes)
            if consulta == "¿Cuánto cobrará la compañía en un mes de notas?":
                mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Cobra compañía", fmt(total_cobrado))])
                resumen_cuentas = resumen_por_cuenta_cobro(detalle)
                if not resumen_cuentas.empty:
                    st.write("Separado por cuenta de cobro")
                    st.dataframe(preparar_tabla_monetaria(resumen_cuentas, ["cobro_compania"]), use_container_width=True)
            elif consulta == "¿Cuánto se pagará a inversores en un mes de notas?":
                mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Pago inversores", fmt(total_pagado))])
            elif consulta == "¿Cuál será el beneficio de la empresa en un mes de notas?":
                mostrar_metricas(f"Resultado {nombre_mes_es(mes)} {anio}", [("Beneficio empresa", fmt(total_beneficio))])
            elif consulta == "¿Cuánto cobrará cada inversor ese mes?":
                if detalle.empty:
                    st.info("No hay cobros de inversores para ese mes.")
                else:
                    resumen = detalle.groupby("inversor", as_index=False)["pago_inversor"].sum().rename(columns={"pago_inversor": "cobro_mes"}).sort_values("cobro_mes", ascending=False)
                    st.dataframe(preparar_tabla_monetaria(resumen, ["cobro_mes"]), use_container_width=True)
            elif consulta == "¿Cuánto cobrará un inversor concreto ese mes?":
                if detalle.empty:
                    total = 0
                else:
                    filtrado = detalle[detalle["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()]
                    total = filtrado["pago_inversor"].sum() if not filtrado.empty else 0
                mostrar_metricas("Resultado", [(f"Cobro de {nombre_inversor}", fmt(total))])

            if not pagos.empty:
                with st.expander("Ver pagos detectados"):
                    st.dataframe(preparar_tabla_monetaria(pagos, []), use_container_width=True)
            if not detalle.empty:
                with st.expander("Ver detalle por nota e inversión"):
                    st.dataframe(preparar_tabla_monetaria(detalle, ["capital_invertido", "cobro_compania", "pago_inversor", "beneficio_empresa"]), use_container_width=True)

        elif consulta == "¿Cuánto ha cobrado la compañía desde el inicio?":
            detalle = preparar_detalle_notas(df_inv, pagos_notas_hasta_hoy(df_cal))
            total = detalle["cobro_compania"].sum() if not detalle.empty else 0
            mostrar_metricas("Resultado", [("Total cobrado compañía", fmt(total))])
            resumen = resumen_por_cuenta_cobro(detalle)
            if not resumen.empty:
                st.dataframe(preparar_tabla_monetaria(resumen, ["cobro_compania"]), use_container_width=True)

        elif consulta == "¿Cuánto se ha pagado a inversores desde el inicio?":
            detalle = preparar_detalle_notas(df_inv, pagos_notas_hasta_hoy(df_cal))
            total = detalle["pago_inversor"].sum() if not detalle.empty else 0
            mostrar_metricas("Resultado", [("Total pagado inversores", fmt(total))])

        elif consulta == "¿Cuál es el beneficio total desde el inicio?":
            detalle = preparar_detalle_notas(df_inv, pagos_notas_hasta_hoy(df_cal))
            total = detalle["beneficio_empresa"].sum() if not detalle.empty else 0
            mostrar_metricas("Resultado", [("Beneficio total", fmt(total))])

        elif consulta == "¿Cuál es el próximo pago de una nota?":
            fecha = proximo_evento_nota(df_cal, int(nota), "PAGO")
            st.success(f"El próximo pago de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}") if fecha is not None else st.info("No hay pagos futuros para esa nota.")

        elif consulta == "¿Cuál es la próxima observación de una nota?":
            fecha = proximo_evento_nota(df_cal, int(nota), "OBSERVACION")
            st.success(f"La próxima observación de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}") if fecha is not None else st.info("No hay observaciones futuras para esa nota.")

        elif consulta == "¿Cuánto capital hay invertido en total?":
            total = filtrar_notas(df_inv)["capital_invertido"].sum()
            mostrar_metricas("Resultado", [("Capital total invertido", fmt(total))])

        elif consulta == "¿Cuánto capital hay actualmente activo?":
            trabajo = filtrar_notas(df_inv)
            hoy = pd.Timestamp.today().normalize()
            activas = trabajo[(trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= hoy) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= hoy))]
            mostrar_metricas("Resultado", [("Capital activo hoy", fmt(activas["capital_invertido"].sum() if not activas.empty else 0))])

        elif consulta == "¿Cuánto capital tiene un inversor?":
            trabajo = filtrar_notas(df_inv)
            filtrado = trabajo[trabajo["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()]
            mostrar_metricas("Resultado", [(f"Capital total de {nombre_inversor}", fmt(filtrado["capital_invertido"].sum() if not filtrado.empty else 0))])

        elif consulta == "¿Cuánto capital activo tiene un inversor?":
            trabajo = filtrar_notas(df_inv)
            hoy = pd.Timestamp.today().normalize()
            filtrado = trabajo[(trabajo["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()) & (trabajo["fecha_inversion"].notna()) & (trabajo["fecha_inversion"] <= hoy) & (trabajo["fecha_final_inversion"].isna() | (trabajo["fecha_final_inversion"] >= hoy))]
            mostrar_metricas("Resultado", [(f"Capital activo de {nombre_inversor}", fmt(filtrado["capital_invertido"].sum() if not filtrado.empty else 0))])

        elif consulta == "Ver ranking de capital por inversor":
            resumen = resumen_capital_por_inversor_notas(df_inv, solo_activo=False)
            st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True)

        elif consulta == "Ver ranking de capital activo":
            resumen = resumen_capital_por_inversor_notas(df_inv, solo_activo=True)
            st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True)


st.title("📊 Sistema Fondo")
st.caption("Aplicación conectada al Excel inversiones.xlsx")

try:
    df_inv, df_cal, df_control = cargar_excel_completo()
except Exception as e:
    st.error("No se ha podido cargar inversiones.xlsx. Revisa que el archivo esté subido a GitHub y que las hojas existan.")
    st.exception(e)
    st.stop()

menu = st.sidebar.selectbox(
    "Selecciona una sección",
    ["Inicio", "Ver Excel", "Consultas Fútbol", "Consultas Notas", "Consultas Paraguay", "Consultas MotoClick"],
)

if menu == "Inicio":
    st.subheader("Panel inicial")
    c1, c2, c3 = st.columns(3)
    c1.metric("Filas inversiones", len(df_inv))
    c2.metric("Eventos calendario", len(df_cal))
    c3.metric("Filas control notas", len(df_control))
    st.success("Sistema cargado correctamente.")

elif menu == "Ver Excel":
    hojas = {"INVERSIONES": df_inv, "CALENDARIO_NOTAS": df_cal, "CONTROL_NOTAS": df_control}
    hoja = st.selectbox("Selecciona hoja", list(hojas.keys()))
    st.dataframe(hojas[hoja], use_container_width=True)

elif menu == "Consultas Fútbol":
    seccion_activo("Fútbol", "futbol", TASA_ANUAL_FUTBOL)

elif menu == "Consultas Notas":
    seccion_notas()

elif menu == "Consultas Paraguay":
    seccion_activo("Paraguay", "paraguay", TASA_ANUAL_PARAGUAY, incluir_ingresado_desde_inicio=True)

elif menu == "Consultas MotoClick":
    seccion_activo("MotoClick", "motoclick", TASA_ANUAL_MOTOCLICK)

