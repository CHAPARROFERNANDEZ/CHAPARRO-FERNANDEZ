import pandas as pd
import re
import calendar
import unicodedata

ARCHIVO = "inversiones.xlsx"
HOJA_INVERSIONES = "INVERSIONES"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"
HOJA_CALLS = "CALENDARIO_CALLS"

TASAS_ACTIVOS = {
    "paraguay": 0.15,
    "motoclick": 0.25,
    "futbol": 0.15,
}

def fmt(x):
    return f"{float(x):,.2f} €"

def limpiar_texto(x):
    if pd.isna(x):
        return ""
    x = str(x).strip().lower()
    x = unicodedata.normalize("NFKD", x).encode("ascii", "ignore").decode("ascii")
    return x

def hoy():
    return pd.Timestamp.today().normalize()

def ultimo_dia_mes(anio, mes):
    return calendar.monthrange(anio, mes)[1]

def nombre_mes_es(mes):
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    return meses.get(mes, str(mes))

def pedir_anio_mes():
    anio = int(input("Año: ").strip())
    mes = int(input("Mes (1-12): ").strip())
    return anio, mes

def cargar_datos():
    df_inv = pd.read_excel(ARCHIVO, sheet_name=HOJA_INVERSIONES)
    df_cal = pd.read_excel(ARCHIVO, sheet_name=HOJA_CALENDARIO)
    df_calls = pd.read_excel(ARCHIVO, sheet_name=HOJA_CALLS)

    df_inv.columns = [c.strip().lower() for c in df_inv.columns]
    df_cal.columns = [c.strip().lower() for c in df_cal.columns]
    df_calls.columns = [c.strip().lower() for c in df_calls.columns]

    for col in df_inv.columns:
        if df_inv[col].dtype == "object":
            df_inv[col] = df_inv[col].astype(str).str.strip()

    df_inv["fecha_inversion"] = pd.to_datetime(df_inv["fecha_inversion"], errors="coerce", dayfirst=True)
    df_inv["fecha_final_inversion"] = pd.to_datetime(df_inv["fecha_final_inversion"], errors="coerce", dayfirst=True)
    df_inv["capital_invertido"] = pd.to_numeric(df_inv["capital_invertido"], errors="coerce").fillna(0)
    df_inv["interes_inversor_anual"] = pd.to_numeric(df_inv["interes_inversor_anual"], errors="coerce").fillna(0)

    if "interes_nota_anual" in df_inv.columns:
        df_inv["interes_nota_anual"] = pd.to_numeric(df_inv["interes_nota_anual"], errors="coerce").fillna(0)
    else:
        df_inv["interes_nota_anual"] = 0

    df_cal["nota"] = pd.to_numeric(df_cal["nota"], errors="coerce").astype("Int64")
    df_cal["tipo_evento"] = df_cal["tipo_evento"].astype(str).str.strip().str.upper()
    df_cal["fecha"] = pd.to_datetime(df_cal["fecha"], errors="coerce", dayfirst=True)

    df_calls["nota"] = pd.to_numeric(df_calls["nota"], errors="coerce").astype("Int64")
    df_calls["fecha_call"] = pd.to_datetime(df_calls["fecha_call"], errors="coerce", dayfirst=True)

    if "estado" in df_calls.columns:
        df_calls["estado"] = df_calls["estado"].fillna("").astype(str).str.strip()

    if "observaciones" in df_calls.columns:
        df_calls["observaciones"] = df_calls["observaciones"].fillna("").astype(str).str.strip()

    return df_inv, df_cal, df_calls

df_inv, df_cal, df_calls = cargar_datos()

def extraer_numero_nota(nombre_activo):
    if pd.isna(nombre_activo):
        return pd.NA
    texto = str(nombre_activo).strip().upper()
    match = re.search(r"NOTA[_\s]?(\d+)", texto)
    return int(match.group(1)) if match else pd.NA

def filtrar_notas(df):
    trabajo = df.copy()

    if "tipo_inversion" in trabajo.columns:
        trabajo = trabajo[trabajo["tipo_inversion"].apply(limpiar_texto) == "nota"].copy()

    trabajo["nota_num"] = trabajo["nombre_activo"].apply(extraer_numero_nota)
    trabajo["nota_num"] = pd.to_numeric(trabajo["nota_num"], errors="coerce").astype("Int64")

    if "activo_generador_interes" in trabajo.columns:
        trabajo = trabajo[trabajo["activo_generador_interes"].apply(limpiar_texto) == "si"].copy()

    return trabajo

def excluir_call(df):
    if "motivo" not in df.columns:
        return df
    return df[df["motivo"].apply(limpiar_texto) != "call"].copy()

def inversiones_activas(df, fecha=None):
    if fecha is None:
        fecha = hoy()
    fecha = pd.Timestamp(fecha).normalize()

    trabajo = df.copy()
    trabajo = excluir_call(trabajo)

    return trabajo[
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= fecha) &
        (
            trabajo["fecha_final_inversion"].isna() |
            (trabajo["fecha_final_inversion"] >= fecha)
        )
    ].copy()

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

def capital_activo_total():
    return inversiones_activas(df_inv)["capital_invertido"].sum()

def capital_activo_por_activo():
    activas = inversiones_activas(df_inv)
    if activas.empty:
        return pd.DataFrame(columns=["activo", "capital"])

    activas["activo"] = activas.apply(detectar_activo, axis=1)

    return (
        activas.groupby("activo", as_index=False)["capital_invertido"]
        .sum()
        .rename(columns={"capital_invertido": "capital"})
        .sort_values("capital", ascending=False)
    )

def capital_por_inversor(activo=False):
    trabajo = inversiones_activas(df_inv) if activo else df_inv.copy()
    trabajo = excluir_call(trabajo)

    return (
        trabajo.groupby("inversor", as_index=False)["capital_invertido"]
        .sum()
        .rename(columns={"capital_invertido": "capital"})
        .sort_values("capital", ascending=False)
    )

def capital_inversor_concreto(nombre, activo=True):
    trabajo = inversiones_activas(df_inv) if activo else excluir_call(df_inv.copy())
    filtrado = trabajo[trabajo["inversor"].apply(limpiar_texto) == limpiar_texto(nombre)]
    return filtrado["capital_invertido"].sum()

def capital_desglosado_por_inversor_mes(anio, mes, nombre_inversor):
    fecha_corte = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))

    trabajo = inversiones_activas(df_inv, fecha=fecha_corte).copy()

    trabajo = trabajo[
        trabajo["inversor"].apply(limpiar_texto) == limpiar_texto(nombre_inversor)
    ].copy()

    if trabajo.empty:
        print(f"\nNo hay capital activo para {nombre_inversor} en {nombre_mes_es(mes)} {anio}.")
        return

    trabajo["activo"] = trabajo.apply(detectar_activo, axis=1)
    trabajo["nota_num"] = trabajo["nombre_activo"].apply(extraer_numero_nota)

    trabajo["inversion"] = trabajo.apply(
        lambda row: f"NOTA {int(row['nota_num'])}"
        if row["activo"] == "notas" and pd.notna(row["nota_num"])
        else str(row["nombre_activo"]),
        axis=1
    )

    resumen = (
        trabajo.groupby("inversion", as_index=False)["capital_invertido"]
        .sum()
        .rename(columns={"capital_invertido": "capital"})
        .sort_values("inversion")
    )

    total = resumen["capital"].sum()
    resumen["capital"] = resumen["capital"].map(fmt)

    print("\n" + "=" * 80)
    print(f"CAPITAL DESGLOSADO DE {nombre_inversor.upper()}")
    print(f"Fecha de corte: {ultimo_dia_mes(anio, mes):02d}/{mes:02d}/{anio}")
    print("=" * 80)
    print(resumen.to_string(index=False))
    print("-" * 80)
    print(f"TOTAL INVERTIDO: {fmt(total)}")

def pagos_notas_mes(anio, mes):
    pagos = df_cal[
        (df_cal["tipo_evento"] == "PAGO") &
        (df_cal["fecha"].notna()) &
        (df_cal["fecha"].dt.year == anio) &
        (df_cal["fecha"].dt.month == mes)
    ].copy()

    print(f"\nPAGOS DETECTADOS EN {nombre_mes_es(mes).upper()} {anio}:")
    if pagos.empty:
        print("No hay pagos.")
    else:
        print(pagos[["nota", "tipo_evento", "fecha"]].sort_values(["fecha", "nota"]).to_string(index=False))

    return pagos.sort_values(["fecha", "nota"])

def inversiones_activas_nota(nota, fecha_pago):
    trabajo = filtrar_notas(df_inv)
    trabajo = excluir_call(trabajo)
    fecha_pago = pd.Timestamp(fecha_pago).normalize()

    return trabajo[
        (trabajo["nota_num"] == nota) &
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= fecha_pago) &
        (
            trabajo["fecha_final_inversion"].isna() |
            (trabajo["fecha_final_inversion"] >= fecha_pago)
        )
    ].copy()

def detalle_notas_por_pagos(pagos):
    resultados = []

    for _, evento in pagos.iterrows():
        nota = evento["nota"]
        fecha_pago = evento["fecha"]
        activas = inversiones_activas_nota(nota, fecha_pago)

        for _, fila in activas.iterrows():
            capital = fila["capital_invertido"]
            cobro_empresa = capital * fila["interes_nota_anual"] / 12
            pago_inversor = capital * fila["interes_inversor_anual"] / 12

            resultados.append({
                "activo": "notas",
                "fecha": fecha_pago,
                "nota": int(nota),
                "id_inversion": fila.get("id_inversion", ""),
                "inversor": fila.get("inversor", ""),
                "capital": capital,
                "cobro_empresa": cobro_empresa,
                "pago_inversor": pago_inversor,
                "beneficio": cobro_empresa - pago_inversor,
            })

    return pd.DataFrame(resultados)

def resumen_notas_mes(anio, mes):
    pagos = pagos_notas_mes(anio, mes)
    detalle = detalle_notas_por_pagos(pagos)

    if detalle.empty:
        return 0, 0, 0, detalle

    return (
        detalle["cobro_empresa"].sum(),
        detalle["pago_inversor"].sum(),
        detalle["beneficio"].sum(),
        detalle
    )

def pagos_notas_hasta_hoy():
    pagos = df_cal[
        (df_cal["tipo_evento"] == "PAGO") &
        (df_cal["fecha"].notna()) &
        (df_cal["fecha"] <= hoy())
    ].copy()

    return pagos.sort_values(["fecha", "nota"])

def acumulado_notas():
    detalle = detalle_notas_por_pagos(pagos_notas_hasta_hoy())

    if detalle.empty:
        return 0, 0, 0

    return (
        detalle["cobro_empresa"].sum(),
        detalle["pago_inversor"].sum(),
        detalle["beneficio"].sum(),
    )

def filtrar_activo_fijo(nombre_activo):
    clave = limpiar_texto(nombre_activo)

    trabajo = df_inv.copy()
    trabajo = trabajo[
        trabajo["subtipo_inversion"].apply(limpiar_texto).str.contains(clave, na=False) |
        trabajo["nombre_activo"].apply(limpiar_texto).str.contains(clave, na=False)
    ].copy()

    return excluir_call(trabajo)

def dias_activos_en_mes(fecha_inicio, fecha_fin, anio, mes, fecha_corte=None):
    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))

    if fecha_corte is not None:
        fecha_corte = pd.Timestamp(fecha_corte).normalize()
        if fecha_corte.year == anio and fecha_corte.month == mes:
            fin_mes = min(fin_mes, fecha_corte)

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

def detalle_activo_fijo_mes(nombre_activo, anio, mes, fecha_corte=None):
    tasa = TASAS_ACTIVOS[nombre_activo]
    trabajo = filtrar_activo_fijo(nombre_activo)
    dias_mes = ultimo_dia_mes(anio, mes)

    if fecha_corte is not None:
        fecha_corte = pd.Timestamp(fecha_corte).normalize()
        if fecha_corte.year == anio and fecha_corte.month == mes:
            dias_mes = fecha_corte.day

    resultados = []

    for _, fila in trabajo.iterrows():
        dias = dias_activos_en_mes(
            fila["fecha_inversion"],
            fila["fecha_final_inversion"],
            anio,
            mes,
            fecha_corte=fecha_corte
        )

        if dias == 0:
            continue

        proporcion = dias / dias_mes
        capital = fila["capital_invertido"]

        cobro_empresa = capital * tasa / 12 * proporcion
        pago_inversor = capital * fila["interes_inversor_anual"] / 12 * proporcion

        resultados.append({
            "activo": nombre_activo,
            "fecha": pd.Timestamp(anio, mes, min(dias_mes, ultimo_dia_mes(anio, mes))),
            "id_inversion": fila.get("id_inversion", ""),
            "inversor": fila.get("inversor", ""),
            "capital": capital,
            "dias_activos": dias,
            "cobro_empresa": cobro_empresa,
            "pago_inversor": pago_inversor,
            "beneficio": cobro_empresa - pago_inversor,
        })

    return pd.DataFrame(resultados)

def resumen_activos_fijos_mes(anio, mes):
    detalles = []

    for activo in TASAS_ACTIVOS:
        detalles.append(detalle_activo_fijo_mes(activo, anio, mes))

    detalles = [d for d in detalles if not d.empty]

    if not detalles:
        return 0, 0, 0, pd.DataFrame()

    detalle = pd.concat(detalles, ignore_index=True)

    return (
        detalle["cobro_empresa"].sum(),
        detalle["pago_inversor"].sum(),
        detalle["beneficio"].sum(),
        detalle
    )

def resumen_global_mes(anio, mes):
    c_notas, p_notas, b_notas, d_notas = resumen_notas_mes(anio, mes)
    c_fijos, p_fijos, b_fijos, d_fijos = resumen_activos_fijos_mes(anio, mes)

    detalles = []
    if not d_notas.empty:
        detalles.append(d_notas)
    if not d_fijos.empty:
        detalles.append(d_fijos)

    detalle = pd.concat(detalles, ignore_index=True) if detalles else pd.DataFrame()

    return c_notas + c_fijos, p_notas + p_fijos, b_notas + b_fijos, detalle

def acumulado_activos_fijos():
    activos = df_inv.copy()
    activos = excluir_call(activos)
    fecha_min = activos["fecha_inversion"].dropna().min()

    if pd.isna(fecha_min):
        return 0, 0, 0

    total_cobro = 0
    total_pago = 0
    total_beneficio = 0

    fecha_actual = hoy()
    anio = fecha_min.year
    mes = fecha_min.month

    while (anio < fecha_actual.year) or (anio == fecha_actual.year and mes <= fecha_actual.month):
        fecha_corte = fecha_actual if (anio == fecha_actual.year and mes == fecha_actual.month) else None

        for activo in TASAS_ACTIVOS:
            detalle = detalle_activo_fijo_mes(activo, anio, mes, fecha_corte=fecha_corte)
            if not detalle.empty:
                total_cobro += detalle["cobro_empresa"].sum()
                total_pago += detalle["pago_inversor"].sum()
                total_beneficio += detalle["beneficio"].sum()

        if mes == 12:
            mes = 1
            anio += 1
        else:
            mes += 1

    return total_cobro, total_pago, total_beneficio

def acumulado_global():
    c_notas, p_notas, b_notas = acumulado_notas()
    c_fijos, p_fijos, b_fijos = acumulado_activos_fijos()

    return c_notas + c_fijos, p_notas + p_fijos, b_notas + b_fijos

def proximos_eventos(tipo="PAGO", limite=10):
    eventos = df_cal[
        (df_cal["tipo_evento"] == tipo.upper()) &
        (df_cal["fecha"].notna()) &
        (df_cal["fecha"] >= hoy())
    ].copy()

    return eventos.sort_values("fecha").head(limite)

def mostrar_calls(df, titulo):
    print("\n" + "=" * 80)
    print(titulo)
    print("=" * 80)

    if df.empty:
        print("No hay calls para esta consulta.")
        return

    mostrar = df.copy()
    mostrar["fecha_call"] = mostrar["fecha_call"].dt.strftime("%d/%m/%Y")

    columnas = ["nota", "fecha_call", "estado", "observaciones"]
    columnas = [c for c in columnas if c in mostrar.columns]

    print(mostrar[columnas].sort_values(["fecha_call", "nota"]).to_string(index=False))

def calls_esta_semana():
    fecha_hoy = hoy()
    inicio_semana = fecha_hoy - pd.Timedelta(days=fecha_hoy.weekday())
    fin_semana = inicio_semana + pd.Timedelta(days=6)

    resultado = df_calls[
        (df_calls["fecha_call"].notna()) &
        (df_calls["fecha_call"] >= inicio_semana) &
        (df_calls["fecha_call"] <= fin_semana)
    ].copy()

    mostrar_calls(resultado, "CALLS DE ESTA SEMANA")

def calls_este_mes():
    fecha_hoy = hoy()

    resultado = df_calls[
        (df_calls["fecha_call"].notna()) &
        (df_calls["fecha_call"].dt.year == fecha_hoy.year) &
        (df_calls["fecha_call"].dt.month == fecha_hoy.month)
    ].copy()

    mostrar_calls(resultado, "CALLS DE ESTE MES")

def proximos_calls(limite=20):
    resultado = df_calls[
        (df_calls["fecha_call"].notna()) &
        (df_calls["fecha_call"] >= hoy())
    ].copy()

    resultado = resultado.sort_values(["fecha_call", "nota"]).head(limite)

    mostrar_calls(resultado, "PRÓXIMOS CALLS")

def calls_vencidos():
    resultado = df_calls[
        (df_calls["fecha_call"].notna()) &
        (df_calls["fecha_call"] < hoy())
    ].copy()

    if "estado" in resultado.columns:
        resultado = resultado[
            ~resultado["estado"].apply(limpiar_texto).isin(["hecho", "realizado", "ejecutado", "call ejecutado"])
        ].copy()

    resultado = resultado.sort_values(["fecha_call", "nota"])

    mostrar_calls(resultado, "CALLS VENCIDOS")

def panel_global():
    print("\n" + "=" * 80)
    print("PANEL GLOBAL DEL FONDO")
    print("=" * 80)
    print(f"Fecha de cálculo: {hoy().strftime('%d/%m/%Y')}")
    print(f"Capital activo total: {fmt(capital_activo_total())}")

    print("\nCAPITAL ACTIVO POR ACTIVO:")
    cap_activos = capital_activo_por_activo()
    if not cap_activos.empty:
        cap_activos["capital"] = cap_activos["capital"].map(fmt)
        print(cap_activos.to_string(index=False))

    c, p, b = acumulado_global()
    print("\nACUMULADO DESDE EL INICIO:")
    print(f"Cobrado compañía:  {fmt(c)}")
    print(f"Pagado inversores: {fmt(p)}")
    print(f"Beneficio empresa: {fmt(b)}")

    print("\nPRÓXIMOS PAGOS:")
    pagos = proximos_eventos("PAGO", 5)
    if pagos.empty:
        print("No hay pagos próximos.")
    else:
        pagos = pagos[["nota", "tipo_evento", "fecha"]].copy()
        pagos["fecha"] = pagos["fecha"].dt.strftime("%d/%m/%Y")
        print(pagos.to_string(index=False))

    print("\nPRÓXIMOS CALLS:")
    calls = df_calls[
        (df_calls["fecha_call"].notna()) &
        (df_calls["fecha_call"] >= hoy())
    ].copy().sort_values(["fecha_call", "nota"]).head(5)

    if calls.empty:
        print("No hay calls próximos.")
    else:
        calls["fecha_call"] = calls["fecha_call"].dt.strftime("%d/%m/%Y")
        columnas = ["nota", "fecha_call", "estado", "observaciones"]
        columnas = [c for c in columnas if c in calls.columns]
        print(calls[columnas].to_string(index=False))

def mostrar_detalle(detalle):
    if detalle.empty:
        print("No hay detalle.")
        return

    mostrar = detalle.copy()

    for col in ["fecha"]:
        if col in mostrar.columns:
            mostrar[col] = pd.to_datetime(mostrar[col]).dt.strftime("%d/%m/%Y")

    for col in ["capital", "cobro_empresa", "pago_inversor", "beneficio"]:
        if col in mostrar.columns:
            mostrar[col] = mostrar[col].map(fmt)

    print(mostrar.to_string(index=False))

def validaciones():
    print("\nVALIDACIONES DEL SISTEMA")
    print("-" * 80)

    if "motivo" in df_inv.columns:
        call_sin_fecha = df_inv[
            (df_inv["motivo"].apply(limpiar_texto) == "call") &
            (df_inv["fecha_final_inversion"].isna())
        ].copy()

        print(f"Inversiones con motivo CALL y sin fecha final: {len(call_sin_fecha)}")
        if not call_sin_fecha.empty:
            cols = ["id_inversion", "inversor", "nombre_activo", "capital_invertido", "motivo"]
            cols = [c for c in cols if c in call_sin_fecha.columns]
            print(call_sin_fecha[cols].to_string(index=False))

    notas = filtrar_notas(df_inv)
    sin_num = notas[notas["nota_num"].notna() == False]
    print(f"Inversiones de notas sin número detectado: {len(sin_num)}")

    sin_fecha = df_inv[df_inv["fecha_inversion"].isna()]
    print(f"Inversiones sin fecha de inversión: {len(sin_fecha)}")

    calls_sin_fecha = df_calls[df_calls["fecha_call"].isna()]
    print(f"Calls sin fecha válida: {len(calls_sin_fecha)}")

def menu():
    print("\n" + "=" * 80)
    print("SISTEMA FONDO")
    print("=" * 80)
    print("1. Panel global")
    print("2. Capital activo total")
    print("3. Capital activo por activo")
    print("4. Capital activo por inversor")
    print("5. Capital activo de un inversor concreto")
    print("6. Resumen mensual GLOBAL")
    print("7. Resumen mensual NOTAS")
    print("8. Resumen mensual PARAGUAY")
    print("9. Resumen mensual MOTOCLICK")
    print("10. Resumen mensual FÚTBOL")
    print("11. Acumulado global desde inicio")
    print("12. Próximos pagos de notas")
    print("13. Próximas observaciones de notas")
    print("14. Validaciones")
    print("15. Calls de esta semana")
    print("16. Calls de este mes")
    print("17. Próximos calls")
    print("18. Calls vencidos")
    print("19. Capital desglosado por inversor")
    print("20. Salir")
    print("=" * 80)

def main():
    while True:
        menu()
        op = input("Elige una opción: ").strip()

        if op == "1":
            panel_global()

        elif op == "2":
            print(f"\nCapital activo total: {fmt(capital_activo_total())}")

        elif op == "3":
            resumen = capital_activo_por_activo()
            resumen["capital"] = resumen["capital"].map(fmt)
            print("\nCAPITAL ACTIVO POR ACTIVO:\n")
            print(resumen.to_string(index=False))

        elif op == "4":
            resumen = capital_por_inversor(activo=True)
            resumen["capital"] = resumen["capital"].map(fmt)
            print("\nCAPITAL ACTIVO POR INVERSOR:\n")
            print(resumen.to_string(index=False))

        elif op == "5":
            nombre = input("Nombre del inversor: ").strip()
            print(f"\nCapital activo de {nombre}: {fmt(capital_inversor_concreto(nombre, activo=True))}")

        elif op == "6":
            anio, mes = pedir_anio_mes()
            c, p, b, detalle = resumen_global_mes(anio, mes)
            print(f"\nRESUMEN GLOBAL {nombre_mes_es(mes).upper()} {anio}")
            print(f"Cobro compañía:  {fmt(c)}")
            print(f"Pago inversores: {fmt(p)}")
            print(f"Beneficio:       {fmt(b)}")
            if input("¿Ver detalle? (s/n): ").lower() == "s":
                mostrar_detalle(detalle)

        elif op == "7":
            anio, mes = pedir_anio_mes()
            c, p, b, detalle = resumen_notas_mes(anio, mes)
            print(f"\nRESUMEN NOTAS {nombre_mes_es(mes).upper()} {anio}")
            print(f"Cobro compañía:  {fmt(c)}")
            print(f"Pago inversores: {fmt(p)}")
            print(f"Beneficio:       {fmt(b)}")
            if input("¿Ver detalle? (s/n): ").lower() == "s":
                mostrar_detalle(detalle)

        elif op == "8":
            anio, mes = pedir_anio_mes()
            detalle = detalle_activo_fijo_mes("paraguay", anio, mes)
            print(f"\nPARAGUAY {nombre_mes_es(mes).upper()} {anio}")
            print(f"Cobro compañía:  {fmt(detalle['cobro_empresa'].sum() if not detalle.empty else 0)}")
            print(f"Pago inversores: {fmt(detalle['pago_inversor'].sum() if not detalle.empty else 0)}")
            print(f"Beneficio:       {fmt(detalle['beneficio'].sum() if not detalle.empty else 0)}")
            if input("¿Ver detalle? (s/n): ").lower() == "s":
                mostrar_detalle(detalle)

        elif op == "9":
            anio, mes = pedir_anio_mes()
            detalle = detalle_activo_fijo_mes("motoclick", anio, mes)
            print(f"\nMOTOCLICK {nombre_mes_es(mes).upper()} {anio}")
            print(f"Cobro compañía:  {fmt(detalle['cobro_empresa'].sum() if not detalle.empty else 0)}")
            print(f"Pago inversores: {fmt(detalle['pago_inversor'].sum() if not detalle.empty else 0)}")
            print(f"Beneficio:       {fmt(detalle['beneficio'].sum() if not detalle.empty else 0)}")
            if input("¿Ver detalle? (s/n): ").lower() == "s":
                mostrar_detalle(detalle)

        elif op == "10":
            anio, mes = pedir_anio_mes()
            detalle = detalle_activo_fijo_mes("futbol", anio, mes)
            print(f"\nFÚTBOL {nombre_mes_es(mes).upper()} {anio}")
            print(f"Cobro compañía:  {fmt(detalle['cobro_empresa'].sum() if not detalle.empty else 0)}")
            print(f"Pago inversores: {fmt(detalle['pago_inversor'].sum() if not detalle.empty else 0)}")
            print(f"Beneficio:       {fmt(detalle['beneficio'].sum() if not detalle.empty else 0)}")
            if input("¿Ver detalle? (s/n): ").lower() == "s":
                mostrar_detalle(detalle)

        elif op == "11":
            c, p, b = acumulado_global()
            print(f"\nACUMULADO GLOBAL HASTA {hoy().strftime('%d/%m/%Y')}")
            print(f"Cobro compañía:  {fmt(c)}")
            print(f"Pago inversores: {fmt(p)}")
            print(f"Beneficio:       {fmt(b)}")

        elif op == "12":
            pagos = proximos_eventos("PAGO", 20)
            pagos["fecha"] = pagos["fecha"].dt.strftime("%d/%m/%Y")
            print("\nPRÓXIMOS PAGOS:\n")
            print(pagos.to_string(index=False))

        elif op == "13":
            obs = proximos_eventos("OBSERVACION", 20)
            obs["fecha"] = obs["fecha"].dt.strftime("%d/%m/%Y")
            print("\nPRÓXIMAS OBSERVACIONES:\n")
            print(obs.to_string(index=False))

        elif op == "14":
            validaciones()

        elif op == "15":
            calls_esta_semana()

        elif op == "16":
            calls_este_mes()

        elif op == "17":
            proximos_calls(20)

        elif op == "18":
            calls_vencidos()

        elif op == "19":
            anio, mes = pedir_anio_mes()
            nombre = input("Nombre del inversor: ").strip()
            capital_desglosado_por_inversor_mes(anio, mes, nombre)

        elif op == "20":
            print("Saliendo...")
            break

        else:
            print("Opción no válida.")

if __name__ == "__main__":
    main()