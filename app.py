import streamlit as st
import pandas as pd
import calendar

st.set_page_config(page_title="Sistema Fondo", layout="wide")

ARCHIVO = "inversiones.xlsx"
HOJA_INVERSIONES = "INVERSIONES"
TASA_ANUAL_FUTBOL = 0.15


@st.cache_data
def cargar_excel():
    data = {}
    xls = pd.ExcelFile(ARCHIVO)
    for hoja in xls.sheet_names:
        data[hoja] = pd.read_excel(ARCHIVO, sheet_name=hoja)
    return data


@st.cache_data
def cargar_inversiones():
    df = pd.read_excel(ARCHIVO, sheet_name=HOJA_INVERSIONES)
    df.columns = [str(c).strip().lower() for c in df.columns]

    for col in [
        "inversor",
        "tipo_inversion",
        "subtipo_inversion",
        "nombre_activo",
        "metodo_calculo",
        "activo_generador_interes",
        "tipo_operacion",
        "capital_nuevo_real",
    ]:
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str).str.strip()

    if "fecha_inversion" in df.columns:
        df["fecha_inversion"] = pd.to_datetime(df["fecha_inversion"], errors="coerce", dayfirst=True)

    if "fecha_final_inversion" in df.columns:
        df["fecha_final_inversion"] = pd.to_datetime(df["fecha_final_inversion"], errors="coerce", dayfirst=True)

    if "capital_invertido" in df.columns:
        df["capital_invertido"] = pd.to_numeric(df["capital_invertido"], errors="coerce").fillna(0)

    if "interes_inversor_anual" in df.columns:
        df["interes_inversor_anual"] = pd.to_numeric(df["interes_inversor_anual"], errors="coerce").fillna(0)

    return df


def ultimo_dia_mes(anio, mes):
    return calendar.monthrange(int(anio), int(mes))[1]


def nombre_mes_es(mes):
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    return meses.get(int(mes), str(mes))


def fmt(x):
    try:
        return f"{float(x):,.2f} €"
    except Exception:
        return "0.00 €"


def filtrar_futbol(df_base):
    if "subtipo_inversion" not in df_base.columns or "nombre_activo" not in df_base.columns:
        return pd.DataFrame()

    subtipo = df_base["subtipo_inversion"].astype(str).str.lower()
    nombre = df_base["nombre_activo"].astype(str).str.lower()

    return df_base[
        subtipo.eq("futbol") |
        subtipo.eq("fútbol") |
        nombre.eq("futbol") |
        nombre.eq("fútbol")
    ].copy()


def calcular_dias_activos_en_mes(fecha_inicio, fecha_fin, anio, mes):
    inicio_mes = pd.Timestamp(int(anio), int(mes), 1)
    fin_mes = pd.Timestamp(int(anio), int(mes), ultimo_dia_mes(anio, mes))

    if pd.isna(fecha_inicio):
        return 0

    if fecha_inicio > fin_mes:
        return 0

    if pd.notna(fecha_fin) and fecha_fin < inicio_mes:
        return 0

    inicio_real = max(fecha_inicio, inicio_mes)
    fin_real = fin_mes if pd.isna(fecha_fin) else min(fecha_fin, fin_mes)

    if inicio_real > fin_real:
        return 0

    return (fin_real - inicio_real).days + 1


def capital_activo_en_fecha(df_base, fecha_consulta, solo_futbol=False, solo_real=False):
    fecha_consulta = pd.Timestamp(fecha_consulta).normalize()
    trabajo = df_base.copy()

    if solo_futbol:
        trabajo = filtrar_futbol(trabajo)

    if solo_real and "capital_nuevo_real" in trabajo.columns:
        trabajo = trabajo[trabajo["capital_nuevo_real"].astype(str).str.lower() == "si"].copy()

    filtrado = trabajo[
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= fecha_consulta) &
        (
            trabajo["fecha_final_inversion"].isna() |
            (trabajo["fecha_final_inversion"] >= fecha_consulta)
        )
    ].copy()

    return filtrado["capital_invertido"].sum()


def preparar_detalle_futbol_mes(df_base, anio, mes):
    df_ft = filtrar_futbol(df_base).copy()
    dias_mes = ultimo_dia_mes(anio, mes)

    resultados = []

    for _, fila in df_ft.iterrows():
        dias_activos = calcular_dias_activos_en_mes(
            fila.get("fecha_inversion"),
            fila.get("fecha_final_inversion"),
            anio,
            mes
        )

        if dias_activos == 0:
            continue

        proporcion = dias_activos / dias_mes
        capital = fila.get("capital_invertido", 0)

        ingreso_bruto = capital * TASA_ANUAL_FUTBOL / 12 * proporcion
        pago_inversor = capital * fila.get("interes_inversor_anual", 0) / 12 * proporcion
        beneficio_empresa = ingreso_bruto - pago_inversor

        resultados.append({
            "id_inversion": fila.get("id_inversion", ""),
            "inversor": fila.get("inversor", ""),
            "capital_invertido": capital,
            "fecha_inversion": fila.get("fecha_inversion"),
            "fecha_final_inversion": fila.get("fecha_final_inversion"),
            "dias_activos": dias_activos,
            "dias_mes": dias_mes,
            "ingreso_bruto_futbol": ingreso_bruto,
            "pago_inversor_mes": pago_inversor,
            "beneficio_empresa_mes": beneficio_empresa
        })

    return pd.DataFrame(resultados)


def ingresos_futbol_mes(df_base, anio, mes):
    detalle = preparar_detalle_futbol_mes(df_base, anio, mes)
    total = detalle["ingreso_bruto_futbol"].sum() if not detalle.empty else 0
    return total, detalle


def cobro_por_inversor_mes(df_base, anio, mes):
    detalle = preparar_detalle_futbol_mes(df_base, anio, mes)

    if detalle.empty:
        return pd.DataFrame(columns=["inversor", "cobro_mes"])

    resumen = (
        detalle.groupby("inversor", as_index=False)["pago_inversor_mes"]
        .sum()
        .rename(columns={"pago_inversor_mes": "cobro_mes"})
        .sort_values("cobro_mes", ascending=False)
    )

    return resumen


def cobro_inversor_concreto_mes(df_base, anio, mes, nombre_inversor):
    resumen = cobro_por_inversor_mes(df_base, anio, mes)

    filtrado = resumen[
        resumen["inversor"].astype(str).str.lower() == str(nombre_inversor).strip().lower()
    ]

    if filtrado.empty:
        return 0

    return filtrado["cobro_mes"].iloc[0]


def beneficio_empresa_mes(df_base, anio, mes):
    detalle = preparar_detalle_futbol_mes(df_base, anio, mes)
    total = detalle["beneficio_empresa_mes"].sum() if not detalle.empty else 0
    return total, detalle


def total_pagado_inversores_desde_inicio(df_base):
    df_ft = filtrar_futbol(df_base).copy()

    if df_ft.empty:
        return 0

    fecha_min = df_ft["fecha_inversion"].dropna().min()

    if pd.isna(fecha_min):
        return 0

    hoy = pd.Timestamp.today().normalize()
    total = 0.0

    anio = fecha_min.year
    mes = fecha_min.month

    while (anio < hoy.year) or (anio == hoy.year and mes <= hoy.month):
        detalle_mes = preparar_detalle_futbol_mes(df_base, anio, mes)

        if not detalle_mes.empty:
            total += detalle_mes["pago_inversor_mes"].sum()

        if mes == 12:
            mes = 1
            anio += 1
        else:
            mes += 1

    return total


def capital_activo_futbol_hoy(df_base, solo_real=False):
    hoy = pd.Timestamp.today().normalize()
    return capital_activo_en_fecha(df_base, hoy, solo_futbol=True, solo_real=solo_real)


def capital_activo_futbol_en_mes(df_base, anio, mes, solo_real=False):
    fecha_consulta = pd.Timestamp(int(anio), int(mes), ultimo_dia_mes(anio, mes))
    return capital_activo_en_fecha(df_base, fecha_consulta, solo_futbol=True, solo_real=solo_real)


def preparar_tabla_monedas(df, columnas):
    mostrar = df.copy()
    for col in columnas:
        if col in mostrar.columns:
            mostrar[col] = mostrar[col].apply(fmt)

    for col in ["fecha_inversion", "fecha_final_inversion"]:
        if col in mostrar.columns:
            mostrar[col] = pd.to_datetime(mostrar[col], errors="coerce").dt.strftime("%d/%m/%Y")

    return mostrar


st.title("📊 Sistema Fondo")
st.write("Aplicación conectada al Excel `inversiones.xlsx`")

try:
    data = cargar_excel()
    df_inv = cargar_inversiones()
except Exception as e:
    st.error(f"No se ha podido cargar el Excel: {e}")
    st.stop()

menu = st.sidebar.selectbox(
    "Selecciona una sección",
    [
        "Inicio",
        "Ver Excel",
        "Consultas Fútbol",
    ]
)

if menu == "Inicio":
    st.subheader("Inicio")
    st.success("Sistema conectado correctamente.")
    st.write("Desde el menú lateral puedes consultar el Excel y las consultas de Fútbol.")

elif menu == "Ver Excel":
    st.subheader("Ver Excel")
    hoja = st.selectbox("Selecciona hoja", list(data.keys()))
    st.dataframe(data[hoja], use_container_width=True)

elif menu == "Consultas Fútbol":
    st.subheader("⚽ Consultas Fútbol")

    consulta = st.selectbox(
        "Elige una pregunta",
        [
            "1. ¿Cuánto ingresará Fútbol en un mes?",
            "2. ¿Cuánto cobrará cada inversor ese mes?",
            "3. ¿Cuánto cobrará un inversor concreto ese mes?",
            "4. ¿Cuál será el beneficio de la empresa ese mes?",
            "5. ¿Cuál es el total pagado a inversores desde el inicio?",
            "6. ¿Cuánto capital hay actualmente activo en Fútbol hoy?",
            "7. ¿Cuánto capital había activo en Fútbol en un mes concreto?",
        ]
    )

    necesita_fecha = consulta.startswith(("1.", "2.", "3.", "4.", "7."))
    necesita_inversor = consulta.startswith("3.")

    col1, col2 = st.columns(2)

    if necesita_fecha:
        with col1:
            anio = st.number_input("Año", min_value=2020, max_value=2100, value=pd.Timestamp.today().year, step=1)
        with col2:
            mes = st.number_input("Mes", min_value=1, max_value=12, value=pd.Timestamp.today().month, step=1)
    else:
        anio = None
        mes = None

    if necesita_inversor:
        inversores = sorted([x for x in df_inv["inversor"].dropna().astype(str).unique()]) if "inversor" in df_inv.columns else []
        if inversores:
            nombre = st.selectbox("Selecciona inversor", inversores)
        else:
            nombre = st.text_input("Nombre del inversor")
    else:
        nombre = None

    if st.button("Calcular"):
        if consulta.startswith("1."):
            total, detalle = ingresos_futbol_mes(df_inv, anio, mes)
            st.success(f"Ingresos de Fútbol en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

            if not detalle.empty:
                mostrar = preparar_tabla_monedas(
                    detalle,
                    ["capital_invertido", "ingreso_bruto_futbol", "pago_inversor_mes", "beneficio_empresa_mes"]
                )
                st.dataframe(mostrar, use_container_width=True)
            else:
                st.info("No hay datos de Fútbol activos para ese mes.")

        elif consulta.startswith("2."):
            resumen = cobro_por_inversor_mes(df_inv, anio, mes)

            if resumen.empty:
                st.info("No hay cobros de inversores para ese mes.")
            else:
                resumen["cobro_mes"] = resumen["cobro_mes"].apply(fmt)
                st.dataframe(resumen, use_container_width=True)

        elif consulta.startswith("3."):
            total = cobro_inversor_concreto_mes(df_inv, anio, mes, nombre)
            st.success(f"{nombre} cobrará en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

        elif consulta.startswith("4."):
            total, detalle = beneficio_empresa_mes(df_inv, anio, mes)
            st.success(f"Beneficio de la empresa en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

            if not detalle.empty:
                mostrar = preparar_tabla_monedas(
                    detalle,
                    ["capital_invertido", "ingreso_bruto_futbol", "pago_inversor_mes", "beneficio_empresa_mes"]
                )
                st.dataframe(mostrar, use_container_width=True)
            else:
                st.info("No hay datos de Fútbol activos para ese mes.")

        elif consulta.startswith("5."):
            total = total_pagado_inversores_desde_inicio(df_inv)
            st.success(f"Total pagado a inversores desde el inicio: {fmt(total)}")

        elif consulta.startswith("6."):
            bruto = capital_activo_futbol_hoy(df_inv, solo_real=False)
            real = capital_activo_futbol_hoy(df_inv, solo_real=True)

            col_a, col_b = st.columns(2)
            col_a.metric("Capital activo en Fútbol", fmt(bruto))
            col_b.metric("Capital activo real", fmt(real))

        elif consulta.startswith("7."):
            bruto = capital_activo_futbol_en_mes(df_inv, anio, mes, solo_real=False)
            real = capital_activo_futbol_en_mes(df_inv, anio, mes, solo_real=True)

            st.success(f"Capital activo en Fútbol al cierre de {nombre_mes_es(mes)} {anio}")
            col_a, col_b = st.columns(2)
            col_a.metric("Capital activo bruto", fmt(bruto))
            col_b.metric("Capital activo real", fmt(real))
