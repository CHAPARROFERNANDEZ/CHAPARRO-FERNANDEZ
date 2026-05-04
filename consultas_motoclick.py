import pandas as pd
import calendar

# =========================================================
# CONFIG
# =========================================================
ARCHIVO = "inversiones.xlsx"
HOJA = "INVERSIONES"
TASA_ANUAL_MOTOCLICK = 0.25   # 25% anual

# =========================================================
# CARGA Y LIMPIEZA
# =========================================================
df = pd.read_excel(ARCHIVO, sheet_name=HOJA)
df.columns = [c.strip().lower() for c in df.columns]

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
        df[col] = df[col].astype(str).str.strip()

df["fecha_inversion"] = pd.to_datetime(df["fecha_inversion"], errors="coerce")
df["fecha_final_inversion"] = pd.to_datetime(df["fecha_final_inversion"], errors="coerce")
df["capital_invertido"] = pd.to_numeric(df["capital_invertido"], errors="coerce").fillna(0)
df["interes_inversor_anual"] = pd.to_numeric(df["interes_inversor_anual"], errors="coerce").fillna(0)

# =========================================================
# FUNCIONES AUXILIARES
# =========================================================
def ultimo_dia_mes(anio, mes):
    return calendar.monthrange(anio, mes)[1]

def nombre_mes_es(mes):
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    return meses.get(mes, str(mes))

def fmt(x):
    return f"{x:,.2f} €"

def filtrar_motoclick(df_base):
    return df_base[
        df_base["subtipo_inversion"].astype(str).str.lower().eq("motoclick")
        | df_base["nombre_activo"].astype(str).str.lower().eq("motoclick")
    ].copy()

def calcular_dias_activos_en_mes(fecha_inicio, fecha_fin, anio, mes):
    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))

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

def capital_activo_en_fecha(df_base, fecha_consulta, solo_motoclick=False, solo_real=False):
    fecha_consulta = pd.Timestamp(fecha_consulta).normalize()
    trabajo = df_base.copy()

    if solo_motoclick:
        trabajo = filtrar_motoclick(trabajo)

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

# =========================================================
# CÁLCULO MENSUAL MOTOCLICK
# =========================================================
def preparar_detalle_motoclick_mes(df_base, anio, mes):
    df_mc = filtrar_motoclick(df_base).copy()
    dias_mes = ultimo_dia_mes(anio, mes)

    resultados = []

    for _, fila in df_mc.iterrows():
        dias_activos = calcular_dias_activos_en_mes(
            fila["fecha_inversion"],
            fila["fecha_final_inversion"],
            anio,
            mes
        )

        if dias_activos == 0:
            continue

        proporcion = dias_activos / dias_mes
        capital = fila["capital_invertido"]

        # Ingreso mensual de MotoClick
        ingreso_bruto = capital * TASA_ANUAL_MOTOCLICK / 12 * proporcion

        # Pago mensual al inversor
        # interes_inversor_anual ya viene en decimal: 0.10 = 10%
        pago_inversor = capital * fila["interes_inversor_anual"] / 12 * proporcion

        beneficio_empresa = ingreso_bruto - pago_inversor

        resultados.append({
            "id_inversion": fila["id_inversion"] if "id_inversion" in fila else "",
            "inversor": fila["inversor"],
            "capital_invertido": capital,
            "fecha_inversion": fila["fecha_inversion"],
            "fecha_final_inversion": fila["fecha_final_inversion"],
            "dias_activos": dias_activos,
            "dias_mes": dias_mes,
            "ingreso_bruto_motoclick": ingreso_bruto,
            "pago_inversor_mes": pago_inversor,
            "beneficio_empresa_mes": beneficio_empresa
        })

    return pd.DataFrame(resultados)

# =========================================================
# CONSULTAS
# =========================================================
def ingresos_motoclick_mes(df_base, anio, mes):
    detalle = preparar_detalle_motoclick_mes(df_base, anio, mes)
    total = detalle["ingreso_bruto_motoclick"].sum() if not detalle.empty else 0
    return total, detalle

def cobro_por_inversor_mes(df_base, anio, mes):
    detalle = preparar_detalle_motoclick_mes(df_base, anio, mes)
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
    filtrado = resumen[resumen["inversor"].str.lower() == nombre_inversor.strip().lower()]
    if filtrado.empty:
        return 0
    return filtrado["cobro_mes"].iloc[0]

def beneficio_empresa_mes(df_base, anio, mes):
    detalle = preparar_detalle_motoclick_mes(df_base, anio, mes)
    total = detalle["beneficio_empresa_mes"].sum() if not detalle.empty else 0
    return total, detalle

def total_pagado_inversores_desde_inicio(df_base):
    df_mc = filtrar_motoclick(df_base).copy()
    if df_mc.empty:
        return 0

    fecha_min = df_mc["fecha_inversion"].dropna().min()
    if pd.isna(fecha_min):
        return 0

    hoy = pd.Timestamp.today().normalize()
    total = 0.0

    anio = fecha_min.year
    mes = fecha_min.month

    while (anio < hoy.year) or (anio == hoy.year and mes <= hoy.month):
        detalle_mes = preparar_detalle_motoclick_mes(df_base, anio, mes)
        if not detalle_mes.empty:
            total += detalle_mes["pago_inversor_mes"].sum()

        if mes == 12:
            mes = 1
            anio += 1
        else:
            mes += 1

    return total

def capital_activo_motoclick_hoy(df_base, solo_real=False):
    hoy = pd.Timestamp.today().normalize()
    return capital_activo_en_fecha(
        df_base,
        hoy,
        solo_motoclick=True,
        solo_real=solo_real
    )

def capital_activo_motoclick_en_mes(df_base, anio, mes, solo_real=False):
    fecha_consulta = pd.Timestamp(anio, mes, ultimo_dia_mes(anio, mes))
    return capital_activo_en_fecha(
        df_base,
        fecha_consulta,
        solo_motoclick=True,
        solo_real=solo_real
    )

# =========================================================
# INPUTS
# =========================================================
def pedir_anio_mes():
    while True:
        try:
            anio = int(input("Año: ").strip())
            mes = int(input("Mes (1-12): ").strip())
            if 1 <= mes <= 12:
                return anio, mes
            print("El mes debe estar entre 1 y 12.")
        except ValueError:
            print("Introduce números válidos.")

# =========================================================
# MENÚ
# =========================================================
def mostrar_menu():
    print("\n" + "=" * 72)
    print("CONSULTAS MOTOCLICK / INVERSORES")
    print("=" * 72)
    print("1. ¿Cuánto ingresará MotoClick en un mes?")
    print("2. ¿Cuánto cobrará cada inversor ese mes?")
    print("3. ¿Cuánto cobrará un inversor concreto ese mes?")
    print("4. ¿Cuál será el beneficio de la empresa ese mes?")
    print("5. ¿Cuál es el total pagado a inversores desde el inicio?")
    print("6. ¿Cuánto capital hay actualmente activo en MotoClick hoy?")
    print("7. ¿Cuánto capital había activo en MotoClick en un mes concreto?")
    print("8. Salir")
    print("=" * 72)

def main():
    while True:
        mostrar_menu()
        opcion = input("Elige una opción: ").strip()

        if opcion == "1":
            anio, mes = pedir_anio_mes()
            total, detalle = ingresos_motoclick_mes(df, anio, mes)
            print(f"\nIngresos de MotoClick en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

            if not detalle.empty:
                ver = input("¿Quieres ver el detalle por inversión? (s/n): ").strip().lower()
                if ver == "s":
                    mostrar = detalle[[
                        "id_inversion",
                        "inversor",
                        "capital_invertido",
                        "dias_activos",
                        "ingreso_bruto_motoclick"
                    ]].copy()
                    mostrar["capital_invertido"] = mostrar["capital_invertido"].map(fmt)
                    mostrar["ingreso_bruto_motoclick"] = mostrar["ingreso_bruto_motoclick"].map(fmt)
                    print("\n" + mostrar.to_string(index=False))

        elif opcion == "2":
            anio, mes = pedir_anio_mes()
            resumen = cobro_por_inversor_mes(df, anio, mes)

            if resumen.empty:
                print("\nNo hay cobros de inversores para ese mes.")
            else:
                resumen["cobro_mes"] = resumen["cobro_mes"].map(fmt)
                print(f"\nCobro por inversor en {nombre_mes_es(mes)} {anio}:\n")
                print(resumen.to_string(index=False))

        elif opcion == "3":
            anio, mes = pedir_anio_mes()
            nombre = input("Nombre del inversor: ").strip()
            total = cobro_inversor_concreto_mes(df, anio, mes, nombre)
            print(f"\n{nombre} cobrará en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

        elif opcion == "4":
            anio, mes = pedir_anio_mes()
            total, detalle = beneficio_empresa_mes(df, anio, mes)
            print(f"\nBeneficio de la empresa en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

            if not detalle.empty:
                ver = input("¿Quieres ver el detalle por inversión? (s/n): ").strip().lower()
                if ver == "s":
                    mostrar = detalle[[
                        "id_inversion",
                        "inversor",
                        "ingreso_bruto_motoclick",
                        "pago_inversor_mes",
                        "beneficio_empresa_mes"
                    ]].copy()
                    mostrar["ingreso_bruto_motoclick"] = mostrar["ingreso_bruto_motoclick"].map(fmt)
                    mostrar["pago_inversor_mes"] = mostrar["pago_inversor_mes"].map(fmt)
                    mostrar["beneficio_empresa_mes"] = mostrar["beneficio_empresa_mes"].map(fmt)
                    print("\n" + mostrar.to_string(index=False))

        elif opcion == "5":
            total = total_pagado_inversores_desde_inicio(df)
            print(f"\nTotal pagado a inversores desde el inicio: {fmt(total)}")

        elif opcion == "6":
            bruto = capital_activo_motoclick_hoy(df, solo_real=False)
            real = capital_activo_motoclick_hoy(df, solo_real=True)

            print(f"\nCapital actualmente activo en MotoClick: {fmt(bruto)}")
            print(f"Capital actualmente activo en MotoClick (solo capital_nuevo_real = si): {fmt(real)}")

        elif opcion == "7":
            anio, mes = pedir_anio_mes()
            bruto = capital_activo_motoclick_en_mes(df, anio, mes, solo_real=False)
            real = capital_activo_motoclick_en_mes(df, anio, mes, solo_real=True)

            print(f"\nCapital activo en MotoClick al cierre de {nombre_mes_es(mes)} {anio}: {fmt(bruto)}")
            print(f"Capital activo en MotoClick al cierre de {nombre_mes_es(mes)} {anio} (solo capital_nuevo_real = si): {fmt(real)}")

        elif opcion == "8":
            print("\nSaliendo del programa.")
            break

        else:
            print("\nOpción no válida.")

if __name__ == "__main__":
    main()