import pandas as pd
import calendar
import re
import yfinance as yf
from datetime import timedelta

ARCHIVO = "inversiones.xlsx"
HOJA_INVERSIONES = "INVERSIONES"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"
HOJA_CONTROL = "CONTROL_NOTAS"

df_inv = pd.read_excel(ARCHIVO, sheet_name=HOJA_INVERSIONES)
df_cal = pd.read_excel(ARCHIVO, sheet_name=HOJA_CALENDARIO)
df_control = pd.read_excel(ARCHIVO, sheet_name=HOJA_CONTROL)

df_inv.columns = [c.strip().lower() for c in df_inv.columns]
df_cal.columns = [c.strip().lower() for c in df_cal.columns]
df_control.columns = [c.strip().lower() for c in df_control.columns]

if "unnamed: 6" in df_inv.columns and "cuenta_cobro" not in df_inv.columns:
    df_inv = df_inv.rename(columns={"unnamed: 6": "cuenta_cobro"})

for col in [
    "id_inversion",
    "inversor",
    "tipo_inversion",
    "subtipo_inversion",
    "nombre_activo",
    "metodo_calculo",
    "activo_generador_interes",
    "tipo_operacion",
    "capital_nuevo_real",
    "cuenta_cobro",
]:
    if col in df_inv.columns:
        df_inv[col] = df_inv[col].astype(str).str.strip()

df_inv["fecha_inversion"] = pd.to_datetime(df_inv["fecha_inversion"], errors="coerce", dayfirst=True)
df_inv["fecha_final_inversion"] = pd.to_datetime(df_inv["fecha_final_inversion"], errors="coerce", dayfirst=True)
df_inv["capital_invertido"] = pd.to_numeric(df_inv["capital_invertido"], errors="coerce").fillna(0)
df_inv["interes_nota_anual"] = pd.to_numeric(df_inv["interes_nota_anual"], errors="coerce").fillna(0)
df_inv["interes_inversor_anual"] = pd.to_numeric(df_inv["interes_inversor_anual"], errors="coerce").fillna(0)

df_cal["nota"] = pd.to_numeric(df_cal["nota"], errors="coerce").astype("Int64")
df_cal["tipo_evento"] = df_cal["tipo_evento"].astype(str).str.strip().str.upper()
df_cal["fecha"] = pd.to_datetime(df_cal["fecha"], errors="coerce", dayfirst=True)

df_control["nota"] = pd.to_numeric(df_control["nota"], errors="coerce").astype("Int64")
df_control["ticker"] = df_control["ticker"].astype(str).str.strip().str.upper()
df_control["precio_compra"] = pd.to_numeric(df_control["precio_compra"], errors="coerce")
df_control["barrera_cupon"] = pd.to_numeric(df_control["barrera_cupon"], errors="coerce")


def nombre_mes_es(mes):
    meses = {
        1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
        5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
        9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre"
    }
    return meses.get(mes, str(mes))


def fmt(x):
    return f"{x:,.2f} €"


def fecha_hoy_texto():
    return pd.Timestamp.today().normalize().strftime("%d/%m/%Y")


def normalizar_cuenta(valor):
    texto = str(valor).strip().lower()

    if texto in ["jordi", "cuenta jordi"]:
        return "JORDI"

    if texto in ["compañia", "compania", "empresa", "sociedad"]:
        return "COMPAÑÍA"

    return "SIN CLASIFICAR"


def normalizar_barrera(valor):
    if pd.isna(valor):
        return None

    valor = float(valor)

    if valor > 1:
        return valor / 100

    return valor


def extraer_numero_nota(nombre_activo):
    if pd.isna(nombre_activo):
        return pd.NA

    texto = str(nombre_activo).strip().upper()
    match = re.search(r"NOTA[_\s]?(\d+)", texto)

    if match:
        return int(match.group(1))

    return pd.NA


def filtrar_notas(df_base):
    trabajo = df_base.copy()

    if "tipo_inversion" in trabajo.columns:
        trabajo = trabajo[
            trabajo["tipo_inversion"].astype(str).str.lower() == "nota"
        ].copy()

    trabajo["nota_num"] = trabajo["nombre_activo"].apply(extraer_numero_nota)
    trabajo["nota_num"] = pd.to_numeric(trabajo["nota_num"], errors="coerce").astype("Int64")

    if "activo_generador_interes" in trabajo.columns:
        trabajo = trabajo[
            trabajo["activo_generador_interes"].astype(str).str.upper() == "SI"
        ].copy()

    if "cuenta_cobro" not in trabajo.columns:
        trabajo["cuenta_cobro"] = "SIN CLASIFICAR"

    trabajo["cuenta_cobro"] = trabajo["cuenta_cobro"].apply(normalizar_cuenta)

    return trabajo


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


def inversiones_activas_en_fecha_para_nota(df_base, nota, fecha_pago):
    fecha_pago = pd.Timestamp(fecha_pago).normalize()
    trabajo = filtrar_notas(df_base).copy()

    filtrado = trabajo[
        (trabajo["nota_num"] == nota) &
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= fecha_pago) &
        (
            trabajo["fecha_final_inversion"].isna() |
            (trabajo["fecha_final_inversion"] >= fecha_pago)
        )
    ].copy()

    return filtrado


def pagos_notas_hasta_fecha(df_cal_base, fecha_corte):
    fecha_corte = pd.Timestamp(fecha_corte).normalize()

    pagos = df_cal_base[
        (df_cal_base["tipo_evento"] == "PAGO") &
        (df_cal_base["fecha"].notna()) &
        (df_cal_base["fecha"] <= fecha_corte)
    ].copy()

    return pagos.sort_values(["fecha", "nota"])


def pagos_notas_mes(df_cal_base, anio, mes):
    pagos = df_cal_base[
        (df_cal_base["tipo_evento"] == "PAGO") &
        (df_cal_base["fecha"].notna()) &
        (df_cal_base["fecha"].dt.year == anio) &
        (df_cal_base["fecha"].dt.month == mes)
    ].copy()

    print(f"\nPAGOS DETECTADOS EN {mes}/{anio}:")
    if pagos.empty:
        print("No hay pagos.")
    else:
        print(
            pagos[["nota", "tipo_evento", "fecha"]]
            .sort_values(["fecha", "nota"])
            .to_string(index=False)
        )

    return pagos.sort_values(["fecha", "nota"])


def obtener_precio_cierre(ticker, fecha_observacion):
    fecha_observacion = pd.Timestamp(fecha_observacion).normalize()

    inicio = fecha_observacion - timedelta(days=10)
    fin = fecha_observacion + timedelta(days=1)

    try:
        data = yf.download(
            ticker,
            start=inicio.strftime("%Y-%m-%d"),
            end=fin.strftime("%Y-%m-%d"),
            progress=False,
            auto_adjust=False
        )
    except Exception:
        return None, None

    if data.empty:
        return None, None

    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)

    data = data.reset_index()

    if "Date" not in data.columns:
        return None, None

    data["Date"] = pd.to_datetime(data["Date"], errors="coerce").dt.normalize()
    data = data[data["Date"] <= fecha_observacion]

    if data.empty:
        return None, None

    ultima = data.iloc[-1]

    precio = ultima["Close"]
    fecha_usada = ultima["Date"]

    if isinstance(precio, pd.Series):
        precio = precio.iloc[0]

    if isinstance(fecha_usada, pd.Series):
        fecha_usada = fecha_usada.iloc[0]

    return float(precio), pd.Timestamp(fecha_usada)


def calcular_resultado_observacion_automatico(nota, fecha_observacion):
    hoy = pd.Timestamp.today().normalize()
    fecha_observacion = pd.Timestamp(fecha_observacion).normalize()

    if fecha_observacion > hoy:
        return (
            "PENDIENTE_FUTURO",
            f"La observación del {fecha_observacion.strftime('%d/%m/%Y')} aún no ha llegado. Se cuenta como cobrada prevista."
        )

    activos = df_control[df_control["nota"] == nota].copy()

    if activos.empty:
        return "SIN_CONTROL", "No existe esta nota en CONTROL_NOTAS"

    detalles = []
    positiva = True

    for _, fila in activos.iterrows():
        ticker = fila["ticker"]
        precio_compra = fila["precio_compra"]
        barrera_cupon = normalizar_barrera(fila["barrera_cupon"])

        if pd.isna(precio_compra) or barrera_cupon is None:
            positiva = False
            detalles.append(f"{ticker}: falta precio de compra o barrera de cupón")
            continue

        precio_barrera = precio_compra * barrera_cupon
        precio_cierre, fecha_usada = obtener_precio_cierre(ticker, fecha_observacion)

        if precio_cierre is None:
            positiva = False
            detalles.append(f"{ticker}: no se pudo obtener precio de cierre")
            continue

        ratio = precio_cierre / precio_compra

        if precio_cierre >= precio_barrera:
            estado = "OK"
        else:
            estado = "NO OK"
            positiva = False

        detalles.append(
            f"{ticker}: cierre={precio_cierre:.2f} | compra={precio_compra:.2f} | "
            f"barrera={precio_barrera:.2f} | ratio={ratio:.2%} | "
            f"fecha cierre usada={fecha_usada.strftime('%d/%m/%Y')} -> {estado}"
        )

    resultado = "POSITIVA" if positiva else "NEGATIVA"

    return resultado, " ; ".join(detalles)


def resultado_observacion_para_pago(df_cal_base, nota, fecha_pago):
    fecha_pago = pd.Timestamp(fecha_pago).normalize()

    observaciones = df_cal_base[
        (df_cal_base["tipo_evento"] == "OBSERVACION") &
        (df_cal_base["nota"] == nota) &
        (df_cal_base["fecha"].notna()) &
        (df_cal_base["fecha"] <= fecha_pago)
    ].copy().sort_values("fecha")

    if observaciones.empty:
        return "SIN_OBSERVACION", None, "No se ha encontrado observación previa al pago"

    fecha_obs = pd.Timestamp(observaciones.iloc[-1]["fecha"]).normalize()

    resultado, detalle = calcular_resultado_observacion_automatico(nota, fecha_obs)

    return resultado, fecha_obs, detalle


def preparar_detalle_pagos(df_inversiones, df_calendario_pagos):
    resultados = []

    for _, evento in df_calendario_pagos.iterrows():
        nota = evento["nota"]
        fecha_pago = evento["fecha"]

        estado_obs, fecha_obs, detalle_obs = resultado_observacion_para_pago(
            df_cal,
            nota,
            fecha_pago
        )

        activas = inversiones_activas_en_fecha_para_nota(df_inversiones, nota, fecha_pago)

        if activas.empty:
            continue

        for _, fila in activas.iterrows():
            capital = fila["capital_invertido"]

            cobro_teorico_compania = capital * fila["interes_nota_anual"] / 12
            pago_inversor = capital * fila["interes_inversor_anual"] / 12

            if estado_obs in ["POSITIVA", "PENDIENTE_FUTURO"]:
                cobro_compania = cobro_teorico_compania
            else:
                cobro_compania = 0

            beneficio = cobro_compania - pago_inversor

            resultados.append({
                "fecha_pago": fecha_pago,
                "fecha_observacion": fecha_obs,
                "nota": int(nota) if pd.notna(nota) else None,
                "resultado_observacion": estado_obs,
                "detalle_observacion": detalle_obs,
                "id_inversion": fila["id_inversion"] if "id_inversion" in fila else "",
                "inversor": fila["inversor"] if "inversor" in fila else "",
                "cuenta_cobro": fila["cuenta_cobro"] if "cuenta_cobro" in fila else "SIN CLASIFICAR",
                "capital_invertido": capital,
                "interes_nota_anual": fila["interes_nota_anual"],
                "interes_inversor_anual": fila["interes_inversor_anual"],
                "cobro_teorico_compania": cobro_teorico_compania,
                "cobro_compania": cobro_compania,
                "pago_inversor": pago_inversor,
                "beneficio_empresa": beneficio
            })

    return pd.DataFrame(resultados)


def resumen_por_cuenta_cobro(detalle):
    if detalle.empty:
        return pd.DataFrame(columns=["cuenta_cobro", "cobro_compania"])

    resumen = (
        detalle.groupby("cuenta_cobro", as_index=False)["cobro_compania"]
        .sum()
        .sort_values("cobro_compania", ascending=False)
    )

    return resumen


def resumen_notas_mes(df_inv_base, df_cal_base, anio, mes):
    pagos_mes = pagos_notas_mes(df_cal_base, anio, mes)
    detalle = preparar_detalle_pagos(df_inv_base, pagos_mes)

    total_cobrado = detalle["cobro_compania"].sum() if not detalle.empty else 0
    total_pagado = detalle["pago_inversor"].sum() if not detalle.empty else 0
    total_beneficio = detalle["beneficio_empresa"].sum() if not detalle.empty else 0

    return total_cobrado, total_pagado, total_beneficio, detalle


def cobro_por_inversor_mes(df_inv_base, df_cal_base, anio, mes):
    _, _, _, detalle = resumen_notas_mes(df_inv_base, df_cal_base, anio, mes)

    if detalle.empty:
        return pd.DataFrame(columns=["inversor", "cobro_mes"])

    resumen = (
        detalle.groupby("inversor", as_index=False)["pago_inversor"]
        .sum()
        .rename(columns={"pago_inversor": "cobro_mes"})
        .sort_values("cobro_mes", ascending=False)
    )
    return resumen


def cobro_inversor_concreto_mes(df_inv_base, df_cal_base, anio, mes, nombre_inversor):
    resumen = cobro_por_inversor_mes(df_inv_base, df_cal_base, anio, mes)

    filtrado = resumen[
        resumen["inversor"].astype(str).str.lower() == nombre_inversor.strip().lower()
    ]

    if filtrado.empty:
        return 0

    return filtrado["cobro_mes"].iloc[0]


def total_cobrado_compania_desde_inicio(df_inv_base, df_cal_base):
    hoy = pd.Timestamp.today().normalize()
    pagos = pagos_notas_hasta_fecha(df_cal_base, hoy)
    detalle = preparar_detalle_pagos(df_inv_base, pagos)

    if detalle.empty:
        return 0, pd.DataFrame(columns=["cuenta_cobro", "cobro_compania"])

    return detalle["cobro_compania"].sum(), resumen_por_cuenta_cobro(detalle)


def total_pagado_inversores_desde_inicio(df_inv_base, df_cal_base):
    hoy = pd.Timestamp.today().normalize()
    pagos = pagos_notas_hasta_fecha(df_cal_base, hoy)
    detalle = preparar_detalle_pagos(df_inv_base, pagos)

    if detalle.empty:
        return 0

    return detalle["pago_inversor"].sum()


def beneficio_total_desde_inicio(df_inv_base, df_cal_base):
    hoy = pd.Timestamp.today().normalize()
    pagos = pagos_notas_hasta_fecha(df_cal_base, hoy)
    detalle = preparar_detalle_pagos(df_inv_base, pagos)

    if detalle.empty:
        return 0

    return detalle["beneficio_empresa"].sum()


def proximo_pago_nota(df_cal_base, nota):
    hoy = pd.Timestamp.today().normalize()

    pagos = df_cal_base[
        (df_cal_base["tipo_evento"] == "PAGO") &
        (df_cal_base["nota"] == nota) &
        (df_cal_base["fecha"] >= hoy)
    ].sort_values("fecha")

    if pagos.empty:
        return None

    return pagos.iloc[0]["fecha"]


def proxima_observacion_nota(df_cal_base, nota):
    hoy = pd.Timestamp.today().normalize()

    obs = df_cal_base[
        (df_cal_base["tipo_evento"] == "OBSERVACION") &
        (df_cal_base["nota"] == nota) &
        (df_cal_base["fecha"] >= hoy)
    ].sort_values("fecha")

    if obs.empty:
        return None

    return obs.iloc[0]["fecha"]


def capital_total_invertido(df_inv_base):
    trabajo = filtrar_notas(df_inv_base).copy()
    return trabajo["capital_invertido"].sum() if not trabajo.empty else 0


def capital_activo_hoy(df_inv_base):
    hoy = pd.Timestamp.today().normalize()
    trabajo = filtrar_notas(df_inv_base).copy()

    activas = trabajo[
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= hoy) &
        (
            trabajo["fecha_final_inversion"].isna() |
            (trabajo["fecha_final_inversion"] >= hoy)
        )
    ]

    return activas["capital_invertido"].sum() if not activas.empty else 0


def capital_por_inversor(df_inv_base, nombre):
    trabajo = filtrar_notas(df_inv_base).copy()

    filtrado = trabajo[
        trabajo["inversor"].astype(str).str.lower() == nombre.strip().lower()
    ]

    return filtrado["capital_invertido"].sum() if not filtrado.empty else 0


def capital_activo_por_inversor(df_inv_base, nombre):
    hoy = pd.Timestamp.today().normalize()
    trabajo = filtrar_notas(df_inv_base).copy()

    filtrado = trabajo[
        (trabajo["inversor"].astype(str).str.lower() == nombre.strip().lower()) &
        (trabajo["fecha_inversion"].notna()) &
        (trabajo["fecha_inversion"] <= hoy) &
        (
            trabajo["fecha_final_inversion"].isna() |
            (trabajo["fecha_final_inversion"] >= hoy)
        )
    ]

    return filtrado["capital_invertido"].sum() if not filtrado.empty else 0


def resumen_capital_por_inversor(df_inv_base, solo_activo=False):
    trabajo = filtrar_notas(df_inv_base).copy()
    hoy = pd.Timestamp.today().normalize()

    if solo_activo:
        trabajo = trabajo[
            (trabajo["fecha_inversion"].notna()) &
            (trabajo["fecha_inversion"] <= hoy) &
            (
                trabajo["fecha_final_inversion"].isna() |
                (trabajo["fecha_final_inversion"] >= hoy)
            )
        ]

    if trabajo.empty:
        return pd.DataFrame(columns=["inversor", "capital"])

    resumen = (
        trabajo.groupby("inversor", as_index=False)["capital_invertido"]
        .sum()
        .rename(columns={"capital_invertido": "capital"})
        .sort_values("capital", ascending=False)
    )

    return resumen


def mostrar_menu():
    print("\n" + "=" * 78)
    print("CONSULTAS DE NOTAS")
    print("=" * 78)
    print("1. ¿Cuánto cobrará la compañía en un mes de notas?")
    print("2. ¿Cuánto se pagará a inversores en un mes de notas?")
    print("3. ¿Cuál será el beneficio de la empresa en un mes de notas?")
    print("4. ¿Cuánto cobrará cada inversor ese mes?")
    print("5. ¿Cuánto cobrará un inversor concreto ese mes?")
    print("6. ¿Cuánto ha cobrado la compañía desde el inicio?")
    print("7. ¿Cuánto se ha pagado a inversores desde el inicio?")
    print("8. ¿Cuál es el beneficio total desde el inicio?")
    print("9. ¿Cuál es el próximo pago de una nota?")
    print("10. ¿Cuál es la próxima observación de una nota?")
    print("11. ¿Cuánto capital hay invertido en total?")
    print("12. ¿Cuánto capital hay actualmente activo?")
    print("13. ¿Cuánto capital tiene un inversor?")
    print("14. ¿Cuánto capital activo tiene un inversor?")
    print("15. Ver ranking de capital por inversor")
    print("16. Ver ranking de capital activo")
    print("17. Salir")
    print("=" * 78)


def main():
    while True:
        mostrar_menu()
        opcion = input("Elige una opción: ").strip()

        if opcion == "1":
            anio, mes = pedir_anio_mes()
            total_cobrado, _, _, detalle = resumen_notas_mes(df_inv, df_cal, anio, mes)
            print(f"\nLa compañía cobrará en {nombre_mes_es(mes)} {anio}: {fmt(total_cobrado)}")

            resumen_cuentas = resumen_por_cuenta_cobro(detalle)

            if not resumen_cuentas.empty:
                resumen_cuentas["cobro_compania"] = resumen_cuentas["cobro_compania"].map(fmt)
                print("\nSeparado por cuenta de cobro:\n")
                print(resumen_cuentas.to_string(index=False))

            if not detalle.empty:
                ver = input("¿Quieres ver el detalle por nota e inversión? (s/n): ").strip().lower()
                if ver == "s":
                    mostrar = detalle[[
                        "fecha_pago",
                        "fecha_observacion",
                        "nota",
                        "resultado_observacion",
                        "id_inversion",
                        "inversor",
                        "cuenta_cobro",
                        "capital_invertido",
                        "cobro_teorico_compania",
                        "cobro_compania"
                    ]].copy()

                    mostrar["fecha_pago"] = pd.to_datetime(mostrar["fecha_pago"]).dt.strftime("%d/%m/%Y")
                    mostrar["fecha_observacion"] = pd.to_datetime(mostrar["fecha_observacion"]).dt.strftime("%d/%m/%Y")
                    mostrar["capital_invertido"] = mostrar["capital_invertido"].map(fmt)
                    mostrar["cobro_teorico_compania"] = mostrar["cobro_teorico_compania"].map(fmt)
                    mostrar["cobro_compania"] = mostrar["cobro_compania"].map(fmt)

                    print("\n" + mostrar.to_string(index=False))

        elif opcion == "2":
            anio, mes = pedir_anio_mes()
            _, total_pagado, _, detalle = resumen_notas_mes(df_inv, df_cal, anio, mes)
            print(f"\nSe pagará a inversores en {nombre_mes_es(mes)} {anio}: {fmt(total_pagado)}")

            if not detalle.empty:
                ver = input("¿Quieres ver el detalle por nota e inversión? (s/n): ").strip().lower()
                if ver == "s":
                    mostrar = detalle[[
                        "fecha_pago",
                        "fecha_observacion",
                        "nota",
                        "resultado_observacion",
                        "id_inversion",
                        "inversor",
                        "cuenta_cobro",
                        "capital_invertido",
                        "pago_inversor"
                    ]].copy()

                    mostrar["fecha_pago"] = pd.to_datetime(mostrar["fecha_pago"]).dt.strftime("%d/%m/%Y")
                    mostrar["fecha_observacion"] = pd.to_datetime(mostrar["fecha_observacion"]).dt.strftime("%d/%m/%Y")
                    mostrar["capital_invertido"] = mostrar["capital_invertido"].map(fmt)
                    mostrar["pago_inversor"] = mostrar["pago_inversor"].map(fmt)

                    print("\n" + mostrar.to_string(index=False))

        elif opcion == "3":
            anio, mes = pedir_anio_mes()
            _, _, total_beneficio, detalle = resumen_notas_mes(df_inv, df_cal, anio, mes)
            print(f"\nBeneficio de la empresa en {nombre_mes_es(mes)} {anio}: {fmt(total_beneficio)}")

            if not detalle.empty:
                ver = input("¿Quieres ver el detalle por nota e inversión? (s/n): ").strip().lower()
                if ver == "s":
                    mostrar = detalle[[
                        "fecha_pago",
                        "fecha_observacion",
                        "nota",
                        "resultado_observacion",
                        "id_inversion",
                        "inversor",
                        "cuenta_cobro",
                        "cobro_teorico_compania",
                        "cobro_compania",
                        "pago_inversor",
                        "beneficio_empresa"
                    ]].copy()

                    mostrar["fecha_pago"] = pd.to_datetime(mostrar["fecha_pago"]).dt.strftime("%d/%m/%Y")
                    mostrar["fecha_observacion"] = pd.to_datetime(mostrar["fecha_observacion"]).dt.strftime("%d/%m/%Y")
                    mostrar["cobro_teorico_compania"] = mostrar["cobro_teorico_compania"].map(fmt)
                    mostrar["cobro_compania"] = mostrar["cobro_compania"].map(fmt)
                    mostrar["pago_inversor"] = mostrar["pago_inversor"].map(fmt)
                    mostrar["beneficio_empresa"] = mostrar["beneficio_empresa"].map(fmt)

                    print("\n" + mostrar.to_string(index=False))

        elif opcion == "4":
            anio, mes = pedir_anio_mes()
            resumen = cobro_por_inversor_mes(df_inv, df_cal, anio, mes)

            if resumen.empty:
                print("\nNo hay cobros de inversores para ese mes.")
            else:
                resumen["cobro_mes"] = resumen["cobro_mes"].map(fmt)
                print(f"\nCobro por inversor en {nombre_mes_es(mes)} {anio}:\n")
                print(resumen.to_string(index=False))

        elif opcion == "5":
            anio, mes = pedir_anio_mes()
            nombre = input("Nombre del inversor: ").strip()
            total = cobro_inversor_concreto_mes(df_inv, df_cal, anio, mes, nombre)
            print(f"\n{nombre} cobrará en {nombre_mes_es(mes)} {anio}: {fmt(total)}")

        elif opcion == "6":
            total, resumen_cuentas = total_cobrado_compania_desde_inicio(df_inv, df_cal)
            print(f"\nCálculo realizado hasta: {fecha_hoy_texto()}")
            print(f"Total cobrado por la compañía desde el inicio: {fmt(total)}")

            if not resumen_cuentas.empty:
                resumen_cuentas["cobro_compania"] = resumen_cuentas["cobro_compania"].map(fmt)
                print("\nSeparado por cuenta de cobro:\n")
                print(resumen_cuentas.to_string(index=False))

        elif opcion == "7":
            total = total_pagado_inversores_desde_inicio(df_inv, df_cal)
            print(f"\nCálculo realizado hasta: {fecha_hoy_texto()}")
            print(f"Total pagado a inversores desde el inicio: {fmt(total)}")

        elif opcion == "8":
            total = beneficio_total_desde_inicio(df_inv, df_cal)
            print(f"\nCálculo realizado hasta: {fecha_hoy_texto()}")
            print(f"Beneficio total desde el inicio: {fmt(total)}")

        elif opcion == "9":
            try:
                nota = int(input("Número de nota: ").strip())
                fecha = proximo_pago_nota(df_cal, nota)
                if fecha is None:
                    print("\nNo hay pagos futuros para esa nota.")
                else:
                    print(f"\nEl próximo pago de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}")
            except ValueError:
                print("\nNúmero de nota no válido.")

        elif opcion == "10":
            try:
                nota = int(input("Número de nota: ").strip())
                fecha = proxima_observacion_nota(df_cal, nota)
                if fecha is None:
                    print("\nNo hay observaciones futuras para esa nota.")
                else:
                    print(f"\nLa próxima observación de la nota {nota} es el {pd.Timestamp(fecha).strftime('%d/%m/%Y')}")
            except ValueError:
                print("\nNúmero de nota no válido.")

        elif opcion == "11":
            total = capital_total_invertido(df_inv)
            print(f"\nCapital total invertido: {fmt(total)}")

        elif opcion == "12":
            total = capital_activo_hoy(df_inv)
            print(f"\nCapital activo a día de hoy: {fmt(total)}")

        elif opcion == "13":
            nombre = input("Nombre del inversor: ").strip()
            total = capital_por_inversor(df_inv, nombre)
            print(f"\nCapital total de {nombre}: {fmt(total)}")

        elif opcion == "14":
            nombre = input("Nombre del inversor: ").strip()
            total = capital_activo_por_inversor(df_inv, nombre)
            print(f"\nCapital activo de {nombre}: {fmt(total)}")

        elif opcion == "15":
            resumen = resumen_capital_por_inversor(df_inv)

            if resumen.empty:
                print("\nNo hay datos.")
            else:
                resumen["capital"] = resumen["capital"].map(fmt)
                print("\nRANKING CAPITAL TOTAL:\n")
                print(resumen.to_string(index=False))

        elif opcion == "16":
            resumen = resumen_capital_por_inversor(df_inv, solo_activo=True)

            if resumen.empty:
                print("\nNo hay datos.")
            else:
                resumen["capital"] = resumen["capital"].map(fmt)
                print("\nRANKING CAPITAL ACTIVO:\n")
                print(resumen.to_string(index=False))

        elif opcion == "17":
            print("\nSaliendo...")
            break

        else:
            print("\nOpción no válida.")


if __name__ == "__main__":
    main()