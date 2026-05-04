mport calendar
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
            if fecha is not None:
                st.success(f"El próximo pago de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}")
            else:
                st.info("No hay pagos futuros para esa nota.")

        elif consulta == "¿Cuál es la próxima observación de una nota?":
            fecha = proximo_evento_nota(df_cal, int(nota), "OBSERVACION")
            if fecha is not None:
                st.success(f"La próxima observación de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}")
            else:
                st.info("No hay observaciones futuras para esa nota.")

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



# =========================
# SECCIONES EXTRA: ALERTAS, CALENDARIO, SISTEMA GLOBAL Y EXTRACTOS
# =========================
from io import BytesIO
import zipfile
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


def leer_hoja_excel(nombre_hoja: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(ARCHIVO, sheet_name=nombre_hoja)
        df.columns = [str(c).strip().lower() for c in df.columns]
        return df
    except Exception:
        return pd.DataFrame()


def normalizar_barrera(valor):
    if pd.isna(valor):
        return None
    try:
        valor = float(valor)
        return valor / 100 if valor > 1 else valor
    except Exception:
        return None


@st.cache_data(show_spinner=False, ttl=3600)
def obtener_ultimo_precio(ticker: str):
    if yf is None:
        return None
    try:
        hist = yf.Ticker(str(ticker).strip().upper()).history(period="5d")
        if hist.empty or "Close" not in hist.columns:
            return None
        return float(hist["Close"].dropna().iloc[-1])
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
        data = yf.download(
            str(ticker).strip().upper(),
            start=inicio.strftime("%Y-%m-%d"),
            end=fin.strftime("%Y-%m-%d"),
            progress=False,
            auto_adjust=False,
        )
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
    if df_control.empty:
        return "SIN_CONTROL", pd.DataFrame()
    sub = df_control[df_control.get("nota") == nota].copy()
    if sub.empty:
        return "SIN_CONTROL", pd.DataFrame()
    barrera_col = columna_barrera_control(sub, preferida=preferida)
    if barrera_col is None or "precio_compra" not in sub.columns or "ticker" not in sub.columns:
        return "SIN_COLUMNAS", pd.DataFrame()
    filas = []
    todo_ok = True
    for _, row in sub.iterrows():
        ticker = row.get("ticker", "")
        compra = pd.to_numeric(row.get("precio_compra"), errors="coerce")
        barrera_pct = normalizar_barrera(row.get(barrera_col))
        if pd.isna(compra) or barrera_pct is None:
            filas.append({"ticker": ticker, "estado": "FALTAN DATOS"})
            todo_ok = False
            continue
        precio_barrera = float(compra) * float(barrera_pct)
        cierre = obtener_cierre_ticker_fecha(ticker, fecha_obs)
        if cierre is None:
            filas.append({"ticker": ticker, "precio_compra": compra, "barrera": precio_barrera, "cierre": None, "estado": "SIN DATO"})
            todo_ok = False
            continue
        estado = "OK" if cierre >= precio_barrera else "NO OK"
        if estado != "OK":
            todo_ok = False
        filas.append({
            "ticker": ticker,
            "precio_compra": float(compra),
            "barrera_%": barrera_pct,
            "precio_barrera": precio_barrera,
            "cierre_usado": cierre,
            "estado": estado,
        })
    return ("POSITIVA" if todo_ok else "NEGATIVA"), pd.DataFrame(filas)



def seccion_notas_archivo():
    """Equivalente web de notas.py: resumen de CONTROL_NOTAS con precios actuales."""
    _, _, df_control = cargar_excel_completo()
    st.header("🧾 Notas")
    st.caption("Resumen tipo notas.py: precio actual, variación, barrera de contingencia y alertas por nota.")

    if yf is None:
        st.error("Falta yfinance. Añade yfinance a requirements.txt.")
        return None

    control = df_control.copy()
    if control.empty:
        st.warning("La hoja CONTROL_NOTAS está vacía o no existe.")
        return None

    # Compatibilidad con nombres usados en tus archivos: CONTINGENCY, BARRERA_CAPITAL o BARRERA_CUPON.
    barrera_col = None
    for candidato in ["contingency", "barrera_capital", "barrera_cupon"]:
        if candidato in control.columns:
            barrera_col = candidato
            break

    columnas_minimas = ["nota", "ticker", "precio_compra"]
    faltan = [c for c in columnas_minimas if c not in control.columns]
    if faltan:
        st.error(f"En CONTROL_NOTAS faltan columnas: {', '.join(faltan)}")
        return None
    if barrera_col is None:
        st.error("En CONTROL_NOTAS falta una columna de barrera: CONTINGENCY, BARRERA_CAPITAL o BARRERA_CUPON.")
        return None

    control["nota"] = pd.to_numeric(control["nota"], errors="coerce")
    control["ticker"] = control["ticker"].astype(str).str.strip().str.upper()
    control["precio_compra"] = pd.to_numeric(control["precio_compra"], errors="coerce")
    control[barrera_col] = pd.to_numeric(control[barrera_col], errors="coerce")
    control[barrera_col] = control[barrera_col].apply(lambda x: x / 100 if pd.notna(x) and x > 1 else x)
    control = control.dropna(subset=["nota", "ticker", "precio_compra", barrera_col]).copy()

    if control.empty:
        st.warning("No hay filas válidas en CONTROL_NOTAS.")
        return None

    if st.button("Actualizar precios actuales"):
        st.cache_data.clear()

    filas = []
    with st.spinner("Descargando precios actuales..."):
        for _, row in control.iterrows():
            ticker = row["ticker"]
            precio_actual = None
            try:
                hist = yf.Ticker(ticker).history(period="5d")
                if not hist.empty:
                    precio_actual = float(hist["Close"].dropna().iloc[-1])
            except Exception:
                precio_actual = None

            precio_compra = float(row["precio_compra"])
            barrera = float(row[barrera_col])
            precio_contingencia = precio_compra * barrera
            variacion = None if precio_actual is None else ((precio_actual - precio_compra) / precio_compra) * 100
            estado = "SIN DATO" if precio_actual is None else ("OK" if precio_actual >= precio_contingencia else "RIESGO")

            filas.append({
                "nota": int(row["nota"]),
                "ticker": ticker,
                "precio_compra": precio_compra,
                "precio_actual": precio_actual,
                "variacion_%": variacion,
                "precio_contingencia": precio_contingencia,
                "estado": estado,
            })

    resumen = pd.DataFrame(filas)
    if resumen.empty:
        st.warning("No se pudo generar el resumen.")
        return None

    notas_riesgo = resumen[resumen["estado"].eq("RIESGO")]["nota"].nunique()
    c1, c2, c3 = st.columns(3)
    c1.metric("Notas analizadas", resumen["nota"].nunique())
    c2.metric("Tickers", len(resumen))
    c3.metric("Notas en riesgo", int(notas_riesgo))

    st.subheader("Resumen por ticker")
    st.dataframe(preparar_tabla_monetaria(resumen, ["precio_compra", "precio_actual", "precio_contingencia"]), use_container_width=True)

    alertas = []
    for nota, grupo in resumen.groupby("nota"):
        riesgo = grupo[grupo["estado"].eq("RIESGO")]
        if not riesgo.empty:
            alertas.append({"nota": int(nota), "tickers_en_riesgo": ", ".join(riesgo["ticker"].astype(str))})
    if alertas:
        st.error("Hay notas en riesgo.")
        st.dataframe(pd.DataFrame(alertas), use_container_width=True)
    else:
        st.success("Ninguna nota en riesgo.")

    peor = resumen.dropna(subset=["variacion_%"]).sort_values("variacion_%").head(1)
    mejor = resumen.dropna(subset=["variacion_%"]).sort_values("variacion_%", ascending=False).head(1)
    if not peor.empty and not mejor.empty:
        c1, c2, c3 = st.columns(3)
        c1.metric("Peor ticker", f"{peor.iloc[0]['ticker']} ({peor.iloc[0]['variacion_%']:.2f}%)")
        c2.metric("Mejor ticker", f"{mejor.iloc[0]['ticker']} ({mejor.iloc[0]['variacion_%']:.2f}%)")
        c3.metric("Variación media", f"{resumen['variacion_%'].mean():.2f}%")

    return None

def seccion_alertas_notas():
    df_inv, df_cal, df_control = cargar_excel_completo()
    st.header("🚨 Alertas Notas")
    fecha = st.date_input("Fecha de consulta", value=pd.Timestamp.today().date())
    fecha = pd.Timestamp(fecha).normalize()

    if df_cal.empty:
        st.warning("No existe la hoja CALENDARIO_NOTAS o está vacía.")
        return

    eventos = df_cal[df_cal["fecha"] == fecha].copy()
    st.subheader(f"Eventos del {fecha.strftime('%d/%m/%Y')}")

    if eventos.empty:
        st.info("No hay observaciones ni pagos para esta fecha.")
    else:
        st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True)

    observaciones = eventos[eventos["tipo_evento"] == "OBSERVACION"].copy() if not eventos.empty else pd.DataFrame()
    pagos = eventos[eventos["tipo_evento"] == "PAGO"].copy() if not eventos.empty else pd.DataFrame()

    if not observaciones.empty:
        st.subheader("Evaluación de observaciones")
        for _, row in observaciones.iterrows():
            nota = int(row["nota"])
            resultado, detalle = evaluar_nota_en_fecha(df_control, nota, fecha, preferida="contingency")
            if resultado == "POSITIVA":
                st.success(f"NOTA {nota}: observación POSITIVA")
            elif resultado == "NEGATIVA":
                st.error(f"NOTA {nota}: observación NEGATIVA")
            else:
                st.warning(f"NOTA {nota}: {resultado}")
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
            if resultado == "POSITIVA":
                st.success(f"NOTA {nota}: pago hoy. La observación previa del {pd.Timestamp(fecha_obs).strftime('%d/%m/%Y')} fue positiva.")
            elif resultado == "NEGATIVA":
                st.error(f"NOTA {nota}: pago hoy. La observación previa del {pd.Timestamp(fecha_obs).strftime('%d/%m/%Y')} fue negativa.")
            else:
                st.warning(f"NOTA {nota}: pago hoy. Resultado observación previa: {resultado}")
            if not detalle.empty:
                with st.expander(f"Detalle NOTA {nota}"):
                    st.dataframe(preparar_tabla_monetaria(detalle, ["precio_compra", "precio_barrera", "cierre_usado"]), use_container_width=True)


def seccion_alertas_semana():
    _, df_cal, _ = cargar_excel_completo()
    st.header("📆 Alertas Semana")
    fecha_inicio = st.date_input("Fecha de inicio", value=pd.Timestamp.today().date())
    fecha_inicio = pd.Timestamp(fecha_inicio).normalize()
    fecha_fin = fecha_inicio + pd.Timedelta(days=6)
    st.caption(f"Del {fecha_inicio.strftime('%d/%m/%Y')} al {fecha_fin.strftime('%d/%m/%Y')}")

    eventos = df_cal[(df_cal["fecha"].notna()) & (df_cal["fecha"] >= fecha_inicio) & (df_cal["fecha"] <= fecha_fin)].copy().sort_values(["fecha", "tipo_evento", "nota"])
    if eventos.empty:
        st.info("No hay observaciones ni pagos esta semana.")
        return
    observaciones = eventos[eventos["tipo_evento"] == "OBSERVACION"]
    pagos = eventos[eventos["tipo_evento"] == "PAGO"]
    c1, c2 = st.columns(2)
    c1.metric("Observaciones", len(observaciones))
    c2.metric("Pagos", len(pagos))
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
        if not eventos.empty:
            st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True)
        else:
            st.info("No hay eventos esta semana.")

    else:
        c1, c2 = st.columns(2)
        anio = int(c1.number_input("Año", min_value=2020, max_value=2100, value=pd.Timestamp.today().year, key=f"cal_{consulta}_anio"))
        mes = int(c2.number_input("Mes", min_value=1, max_value=12, value=pd.Timestamp.today().month, key=f"cal_{consulta}_mes"))
        eventos = eventos_calendario_mes(df_cal, anio, mes)

        if consulta == "Mes completo":
            st.subheader(f"Calendario de {nombre_mes_es(mes)} {anio}")
            if not eventos.empty:
                st.dataframe(preparar_tabla_monetaria(eventos, []), use_container_width=True)
            else:
                st.info("No hay eventos ese mes.")

        elif consulta == "Semana concreta de un mes":
            semana = int(st.number_input("Semana del mes", min_value=1, max_value=5, value=1))
            filtrado = eventos[eventos["semana_mes"] == semana].copy() if not eventos.empty else pd.DataFrame()
            if not filtrado.empty:
                st.dataframe(preparar_tabla_monetaria(filtrado, []), use_container_width=True)
            else:
                st.info("No hay eventos en esa semana.")

        elif consulta == "Exportar calendario de un mes":
            if eventos.empty:
                st.info("No hay eventos para exportar en ese mes.")
            else:
                salida = BytesIO()
                exportar = eventos.copy()
                exportar["fecha"] = exportar["fecha"].dt.strftime("%d/%m/%Y")
                with pd.ExcelWriter(salida, engine="openpyxl") as writer:
                    exportar.to_excel(writer, index=False, sheet_name="CALENDARIO")
                st.download_button(
                    "Descargar Excel",
                    data=salida.getvalue(),
                    file_name=f"calendario_notas_{mes}_{anio}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    return None


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


def seccion_sistema_fondo():
    df_inv, df_cal, _ = cargar_excel_completo()
    df_calls = leer_hoja_excel("CALENDARIO_CALLS")
    if not df_calls.empty:
        if "fecha_call" in df_calls.columns:
            df_calls["fecha_call"] = pd.to_datetime(df_calls["fecha_call"], errors="coerce", dayfirst=True).dt.normalize()
        if "nota" in df_calls.columns:
            df_calls["nota"] = pd.to_numeric(df_calls["nota"], errors="coerce").astype("Int64")
    st.header("🏦 Sistema Fondo")
    consulta = st.selectbox("Consulta", [
        "Panel global", "Capital activo total", "Capital activo por activo", "Capital activo por inversor",
        "Capital activo de un inversor concreto", "Resumen mensual global", "Validaciones", "Calls de esta semana",
        "Calls de este mes", "Próximos calls", "Calls vencidos", "Capital desglosado por inversor"
    ])

    if consulta == "Panel global":
        activas = inversiones_activas_global(df_inv)
        activas["activo"] = activas.apply(detectar_activo, axis=1) if not activas.empty else []
        c_global, p_global, b_global, _detalle_dummy, _ = resumen_notas_mes(df_inv, df_cal, pd.Timestamp.today().year, pd.Timestamp.today().month)
        c1, c2, c3 = st.columns(3)
        c1.metric("Capital activo total", fmt(activas["capital_invertido"].sum() if not activas.empty else 0))
        c2.metric("Cobro notas mes actual", fmt(c_global))
        c3.metric("Beneficio notas mes actual", fmt(b_global))
        if not activas.empty:
            resumen = activas.groupby("activo", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
            st.subheader("Capital activo por activo")
            st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True)
        proximos = df_cal[(df_cal["fecha"].notna()) & (df_cal["fecha"] >= pd.Timestamp.today().normalize())].sort_values("fecha").head(10)
        st.subheader("Próximos eventos")
        if not proximos.empty:
            st.dataframe(preparar_tabla_monetaria(proximos, []), use_container_width=True)
        else:
            st.info("No hay próximos eventos.")

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
        if activas.empty:
            st.info("No hay inversiones activas.")
        else:
            resumen = activas.groupby("inversor", as_index=False)["capital_invertido"].sum().rename(columns={"capital_invertido": "capital"}).sort_values("capital", ascending=False)
            st.dataframe(preparar_tabla_monetaria(resumen, ["capital"]), use_container_width=True)

    elif consulta in ["Capital activo de un inversor concreto", "Capital desglosado por inversor"]:
        inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
        nombre = st.selectbox("Inversor", inversores) if inversores else st.text_input("Inversor")
        if consulta == "Capital desglosado por inversor":
            c1, c2 = st.columns(2)
            anio = int(c1.number_input("Año", min_value=2020, max_value=2100, value=pd.Timestamp.today().year))
            mes = int(c2.number_input("Mes", min_value=1, max_value=12, value=pd.Timestamp.today().month))
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
        anio = int(c1.number_input("Año", min_value=2020, max_value=2100, value=pd.Timestamp.today().year))
        mes = int(c2.number_input("Mes", min_value=1, max_value=12, value=pd.Timestamp.today().month))
        c_notas, p_notas, b_notas, d_notas, _ = resumen_notas_mes(df_inv, df_cal, anio, mes)
        detalles = []
        for activo, tasa in [("paraguay", TASA_ANUAL_PARAGUAY), ("motoclick", TASA_ANUAL_MOTOCLICK), ("futbol", TASA_ANUAL_FUTBOL)]:
            det = detalle_activo_mes(df_inv, activo, tasa, anio, mes)
            if not det.empty:
                det["activo"] = activo
                detalles.append(det)
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
        resultados = []
        if "motivo" in df_inv.columns:
            call_sin_fecha = df_inv[(df_inv["motivo"].apply(limpiar_texto) == "call") & (df_inv["fecha_final_inversion"].isna())]
            resultados.append({"validacion": "Inversiones con motivo CALL y sin fecha final", "cantidad": len(call_sin_fecha)})
        notas = filtrar_notas(df_inv)
        resultados.append({"validacion": "Inversiones de notas sin número detectado", "cantidad": int(notas[notas["nota_num"].isna()].shape[0])})
        resultados.append({"validacion": "Inversiones sin fecha de inversión", "cantidad": int(df_inv[df_inv["fecha_inversion"].isna()].shape[0])})
        if not df_calls.empty and "fecha_call" in df_calls.columns:
            resultados.append({"validacion": "Calls sin fecha válida", "cantidad": int(df_calls[df_calls["fecha_call"].isna()].shape[0])})
        st.dataframe(pd.DataFrame(resultados), use_container_width=True)

    else:
        if df_calls.empty or "fecha_call" not in df_calls.columns:
            st.warning("No existe la hoja CALENDARIO_CALLS o no tiene la columna fecha_call.")
            return
        hoy = pd.Timestamp.today().normalize()
        if consulta == "Calls de esta semana":
            inicio = hoy - pd.Timedelta(days=hoy.weekday())
            fin = inicio + pd.Timedelta(days=6)
            res = df_calls[(df_calls["fecha_call"] >= inicio) & (df_calls["fecha_call"] <= fin)].copy()
        elif consulta == "Calls de este mes":
            res = df_calls[(df_calls["fecha_call"].dt.year == hoy.year) & (df_calls["fecha_call"].dt.month == hoy.month)].copy()
        elif consulta == "Próximos calls":
            res = df_calls[df_calls["fecha_call"] >= hoy].copy().sort_values("fecha_call").head(20)
        else:
            res = df_calls[df_calls["fecha_call"] < hoy].copy()
            if "estado" in res.columns:
                res = res[~res["estado"].apply(limpiar_texto).isin(["hecho", "realizado", "ejecutado", "call ejecutado"])]
        if not res.empty:
            st.dataframe(preparar_tabla_monetaria(res, []), use_container_width=True)
        else:
            st.info("No hay calls para esta consulta.")



def formatear_extracto_excel_bytes(contenido: bytes, inversor: str, fecha_corte: datetime) -> bytes:
    """Aplica el formato bonito del generador de terminal, pero en memoria para descarga web."""
    bio = BytesIO(contenido)
    wb = load_workbook(bio)

    azul = "1F4E78"
    azul_claro = "D9EAF7"
    verde = "E2F0D9"
    blanco = "FFFFFF"
    borde_fino = Side(style="thin", color="D9D9D9")
    borde = Border(left=borde_fino, right=borde_fino, top=borde_fino, bottom=borde_fino)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
        for row in ws.iter_rows():
            for cell in row:
                cell.font = Font(name="Calibri", size=11)
                cell.alignment = Alignment(vertical="center")
                cell.border = borde
        for col in range(1, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
        for row_num in range(1, ws.max_row + 1):
            ws.row_dimensions[row_num].height = 22

    if "RESUMEN" in wb.sheetnames:
        ws = wb["RESUMEN"]
        ws.insert_rows(1, 5)
        ws["A1"] = "EXTRACTO DE INVERSIÓN"
        ws["A1"].font = Font(name="Calibri", size=20, bold=True, color=blanco)
        ws["A1"].fill = PatternFill("solid", fgColor=azul)
        ws["A1"].alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(4, ws.max_column))
        ws["A3"] = "Inversor"
        ws["B3"] = inversor
        ws["A4"] = "Fecha de corte"
        ws["B4"] = fecha_corte.strftime("%d/%m/%Y")
        for cell in ["A3", "A4"]:
            ws[cell].font = Font(bold=True)
            ws[cell].fill = PatternFill("solid", fgColor=azul_claro)
        header_row = 6
        for cell in ws[header_row]:
            cell.font = Font(bold=True, color=blanco)
            cell.fill = PatternFill("solid", fgColor=azul)
            cell.alignment = Alignment(horizontal="center")
        for row_num in range(header_row + 1, ws.max_row + 1):
            for col_num in range(1, ws.max_column + 1):
                ws.cell(row_num, col_num).fill = PatternFill("solid", fgColor=verde)
                ws.cell(row_num, col_num).alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 20
        ws.column_dimensions["C"].width = 22
        ws.column_dimensions["D"].width = 26

    if "TOTALES_MES" in wb.sheetnames:
        ws = wb["TOTALES_MES"]
        ws.insert_rows(1, 3)
        ws["A1"] = "RESUMEN MENSUAL"
        ws["A1"].font = Font(size=18, bold=True, color=blanco)
        ws["A1"].fill = PatternFill("solid", fgColor=azul)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, ws.max_column))
        ws["A1"].alignment = Alignment(horizontal="center")
        for cell in ws[4]:
            cell.font = Font(bold=True, color=blanco)
            cell.fill = PatternFill("solid", fgColor=azul)
            cell.alignment = Alignment(horizontal="center")
        for row_num in range(5, ws.max_row + 1):
            if ws.max_column >= 2:
                ws.cell(row_num, 2).number_format = '#,##0.00 €'
            for col_num in range(1, ws.max_column + 1):
                ws.cell(row_num, col_num).alignment = Alignment(horizontal="center")
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 18

    if "DETALLE" in wb.sheetnames:
        ws = wb["DETALLE"]
        ws.insert_rows(1, 3)
        ws["A1"] = "DETALLE DEL EXTRACTO"
        ws["A1"].font = Font(size=18, bold=True, color=blanco)
        ws["A1"].fill = PatternFill("solid", fgColor=azul)
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
        ws["A1"].alignment = Alignment(horizontal="center")
        for cell in ws[4]:
            cell.font = Font(bold=True, color=blanco)
            cell.fill = PatternFill("solid", fgColor=azul)
            cell.alignment = Alignment(horizontal="center")
        for row_num in range(5, ws.max_row + 1):
            fill = azul_claro if row_num % 2 == 0 else blanco
            for col_num in range(1, ws.max_column + 1):
                ws.cell(row_num, col_num).fill = PatternFill("solid", fgColor=fill)
                ws.cell(row_num, col_num).alignment = Alignment(horizontal="center")
            for col_num in [8, 11]:
                if ws.max_column >= col_num:
                    ws.cell(row_num, col_num).number_format = '#,##0.00 €'
        anchos = {"A": 24, "B": 16, "C": 18, "D": 18, "E": 20, "F": 14, "G": 18, "H": 18, "I": 16, "J": 12, "K": 18}
        for col, ancho in anchos.items():
            ws.column_dimensions[col].width = ancho

    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def generar_extractos(df_inv: pd.DataFrame, modo: str, inversor_elegido: str | None, anio: int, mes: int):
    df = df_inv.copy()
    for col in ["inversor", "tipo_inversion", "subtipo_inversion", "nombre_activo", "tipo_operacion", "capital_nuevo_real"]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()
    if "capital_nuevo_real" in df.columns:
        df = df[df["capital_nuevo_real"].str.upper().isin(["SI", ""])].copy()
    if modo == "Un inversor" and inversor_elegido:
        df = df[df["inversor"].str.upper() == inversor_elegido.upper()].copy()
    fecha_corte = datetime(anio, mes, ultimo_dia_mes(anio, mes))
    filas = []
    for _, row in df.iterrows():
        fecha_inicio = row.get("fecha_inversion")
        if pd.isna(fecha_inicio):
            continue
        fecha_final_excel = row.get("fecha_final_inversion")
        fecha_fin = fecha_corte if pd.isna(fecha_final_excel) else min(pd.Timestamp(fecha_final_excel).to_pydatetime(), fecha_corte)
        if pd.Timestamp(fecha_inicio).to_pydatetime() > fecha_fin:
            continue
        actual = datetime(fecha_inicio.year, fecha_inicio.month, 1)
        fin_mes = datetime(fecha_fin.year, fecha_fin.month, 1)
        while actual <= fin_mes:
            dias_mes = calendar.monthrange(actual.year, actual.month)[1]
            inicio_mes = datetime(actual.year, actual.month, 1)
            fin_mes_real = datetime(actual.year, actual.month, dias_mes)
            inicio_calc = max(pd.Timestamp(fecha_inicio).to_pydatetime(), inicio_mes)
            fin_calc = min(fecha_fin, fin_mes_real)
            if inicio_calc <= fin_calc:
                dias = (fin_calc - inicio_calc).days + 1
                capital = float(row.get("capital_invertido", 0))
                interes = float(row.get("interes_inversor_anual", 0))
                interes_mes = round((capital * interes / 12) * dias / dias_mes, 2)
                filas.append({
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
                })
            actual = datetime(actual.year + 1, 1, 1) if actual.month == 12 else datetime(actual.year, actual.month + 1, 1)
    resultado = pd.DataFrame(filas)
    if resultado.empty:
        return []
    archivos = []
    for inversor, grupo in resultado.groupby("inversor"):
        detalle = grupo.copy()
        totales_mes = detalle.groupby("mes", as_index=False)["interes_mes"].sum().rename(columns={"interes_mes": "total_mes"})
        capital_total = detalle.groupby("id_inversion")["capital_invertido"].first().sum()
        resumen = pd.DataFrame([{"inversor": inversor, "fecha_corte": fecha_corte.strftime("%d/%m/%Y"), "capital_total": round(capital_total, 2), "total_intereses_acumulados": round(detalle["interes_mes"].sum(), 2)}])
        salida = BytesIO()
        with pd.ExcelWriter(salida, engine="openpyxl") as writer:
            resumen.to_excel(writer, sheet_name="RESUMEN", index=False)
            totales_mes.to_excel(writer, sheet_name="TOTALES_MES", index=False)
            detalle.to_excel(writer, sheet_name="DETALLE", index=False)
        nombre_archivo = f"extracto_{str(inversor).upper().replace(' ', '_')}_{fecha_corte.strftime('%d%m%Y')}.xlsx"
        contenido_formateado = formatear_extracto_excel_bytes(salida.getvalue(), str(inversor), fecha_corte)
        archivos.append((nombre_archivo, contenido_formateado))
    return archivos


def seccion_extractos():
    df_inv, _, _ = cargar_excel_completo()
    st.header("📤 Extractos")
    modo = st.radio("¿Qué quieres generar?", ["Todos", "Un inversor"], horizontal=True)
    inversores = sorted([x for x in df_inv.get("inversor", pd.Series(dtype=str)).dropna().astype(str).unique() if x.strip()])
    inversor = None
    if modo == "Un inversor":
        inversor = st.selectbox("Inversor", inversores) if inversores else st.text_input("Inversor")
    c1, c2 = st.columns(2)
    anio = int(c1.number_input("Año de corte", min_value=2020, max_value=2100, value=pd.Timestamp.today().year))
    mes = int(c2.number_input("Mes de corte", min_value=1, max_value=12, value=pd.Timestamp.today().month))
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
    ["Inicio", "Ver Excel", "Consultas Fútbol", "Consultas Notas", "Consultas Paraguay", "Consultas MotoClick", "Notas", "Alertas Notas", "Alertas Semana", "Calendario Notas", "Sistema Fondo", "Extractos"],
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

elif menu == "Notas":
    seccion_notas_archivo()

elif menu == "Alertas Notas":
    seccion_alertas_notas()

elif menu == "Alertas Semana":
    seccion_alertas_semana()

elif menu == "Calendario Notas":
    seccion_calendario_notas()

elif menu == "Sistema Fondo":
    seccion_sistema_fondo()

elif menu == "Extractos":
    seccion_extractos()

