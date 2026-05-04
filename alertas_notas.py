import os
from datetime import timedelta
import pandas as pd
import yfinance as yf

# =========================================================
# CONFIGURACIÓN
# =========================================================
ARCHIVO_EXCEL = "inversiones.xlsx"
HOJA_CONTROL = "CONTROL_NOTAS"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"
HOJA_RESULTADOS = "RESULTADOS_OBSERVACION"

# Si quieres probar una fecha concreta, ponla aquí:
# FECHA_MANUAL = "2027-01-04"
FECHA_MANUAL = None

# True = imprime detalles técnicos
DEBUG = True


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================
def obtener_fecha_hoy():
    if FECHA_MANUAL:
        return pd.to_datetime(FECHA_MANUAL).normalize()
    return pd.Timestamp.today().normalize()


def normalizar_columnas(df):
    df.columns = [
        str(col).strip().upper().replace(" ", "_")
        for col in df.columns
    ]
    return df


def convertir_numero(valor):
    """
    Convierte números escritos con coma o punto:
    410,04 -> 410.04
    0,4 -> 0.4
    40 -> 40.0
    """
    if pd.isna(valor):
        return None

    texto = str(valor).strip()
    if texto == "":
        return None

    texto = texto.replace("%", "").replace("€", "").replace("$", "")
    texto = texto.replace(".", "") if texto.count(",") == 1 and texto.count(".") > 1 else texto
    texto = texto.replace(",", ".")

    try:
        return float(texto)
    except ValueError:
        return None


def convertir_contingency(valor):
    """
    Acepta:
    0.4  -> 0.4
    40   -> 0.4
    40%  -> 0.4
    0,4  -> 0.4
    """
    num = convertir_numero(valor)
    if num is None:
        return None

    if num > 1:
        num = num / 100

    return num


def cargar_datos():
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(f"No existe el archivo {ARCHIVO_EXCEL}")

    control = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_CONTROL)
    calendario = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_CALENDARIO)

    control = normalizar_columnas(control)
    calendario = normalizar_columnas(calendario)

    if DEBUG:
        print("COLUMNAS CONTROL:", control.columns.tolist())
        print("COLUMNAS CALENDARIO:", calendario.columns.tolist())

    # Normalizar CONTROL_NOTAS
    if "NOTA" not in control.columns:
        raise ValueError("En CONTROL_NOTAS falta la columna NOTA")
    if "TICKER" not in control.columns:
        raise ValueError("En CONTROL_NOTAS falta la columna TICKER")
    if "PRECIO_COMPRA" not in control.columns:
        raise ValueError("En CONTROL_NOTAS falta la columna PRECIO_COMPRA")
    if "CONTINGENCY" not in control.columns:
        raise ValueError("En CONTROL_NOTAS falta la columna CONTINGENCY")

    control["NOTA"] = pd.to_numeric(control["NOTA"], errors="coerce")
    control["TICKER"] = control["TICKER"].astype(str).str.strip().str.upper()
    control["PRECIO_COMPRA"] = control["PRECIO_COMPRA"].apply(convertir_numero)
    control["CONTINGENCY"] = control["CONTINGENCY"].apply(convertir_contingency)

    control = control.dropna(subset=["NOTA", "TICKER", "PRECIO_COMPRA", "CONTINGENCY"]).copy()

    # Normalizar CALENDARIO_NOTAS
    if "NOTA" not in calendario.columns:
        raise ValueError("En CALENDARIO_NOTAS falta la columna NOTA")
    if "TIPO_EVENTO" not in calendario.columns:
        raise ValueError("En CALENDARIO_NOTAS falta la columna TIPO_EVENTO")
    if "FECHA" not in calendario.columns:
        raise ValueError("En CALENDARIO_NOTAS falta la columna FECHA")

    calendario["NOTA"] = pd.to_numeric(calendario["NOTA"], errors="coerce")
    calendario["TIPO_EVENTO"] = calendario["TIPO_EVENTO"].astype(str).str.strip().str.upper()
    calendario["FECHA"] = pd.to_datetime(
        calendario["FECHA"],
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    calendario = calendario.dropna(subset=["NOTA", "TIPO_EVENTO", "FECHA"]).copy()

    return control, calendario


def cargar_resultados_existentes():
    try:
        resultados = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_RESULTADOS)
        resultados = normalizar_columnas(resultados)

        resultados["NOTA"] = pd.to_numeric(resultados["NOTA"], errors="coerce")
        resultados["FECHA_OBSERVACION"] = pd.to_datetime(
            resultados["FECHA_OBSERVACION"],
            dayfirst=True,
            errors="coerce"
        ).dt.normalize()
        resultados["RESULTADO"] = resultados["RESULTADO"].astype(str).str.strip().str.upper()
        resultados["DETALLE"] = resultados["DETALLE"].fillna("").astype(str)

        return resultados

    except Exception:
        return pd.DataFrame(columns=["NOTA", "FECHA_OBSERVACION", "RESULTADO", "DETALLE"])


def guardar_resultados(resultados_df):
    with pd.ExcelWriter(
        ARCHIVO_EXCEL,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        resultados_df.to_excel(writer, sheet_name=HOJA_RESULTADOS, index=False)


def obtener_cierre_ticker_en_fecha(ticker, fecha_objetivo):
    """
    Devuelve el último cierre disponible hasta esa fecha.
    Esto sirve si la fecha exacta cae en festivo/no trading.
    """
    inicio = fecha_objetivo - timedelta(days=10)
    fin = fecha_objetivo + timedelta(days=2)

    try:
        data = yf.download(
            ticker,
            start=inicio.strftime("%Y-%m-%d"),
            end=fin.strftime("%Y-%m-%d"),
            progress=False,
            auto_adjust=False
        )

        if data.empty:
            return None

        if isinstance(data.columns, pd.MultiIndex):
            if ("Close", ticker) in data.columns:
                cierres = data[("Close", ticker)].dropna()
            elif "Close" in data.columns.get_level_values(0):
                cierres = data["Close"].iloc[:, 0].dropna()
            else:
                return None
        else:
            if "Close" not in data.columns:
                return None
            cierres = data["Close"].dropna()

        if cierres.empty:
            return None

        cierres.index = pd.to_datetime(cierres.index).normalize()
        cierres_validos = cierres[cierres.index <= fecha_objetivo]

        if cierres_validos.empty:
            return None

        return float(cierres_validos.iloc[-1])

    except Exception as e:
        if DEBUG:
            print(f"Error descargando {ticker}: {e}")
        return None


def evaluar_observacion(nota, fecha_obs, control_df):
    """
    POSITIVA si todos los tickers de la nota cierran >= precio_compra * contingency
    """
    subset = control_df[control_df["NOTA"] == nota].copy()

    if subset.empty:
        return "NEGATIVA", "No hay tickers configurados para esta nota"

    detalles = []
    todo_ok = True

    for _, row in subset.iterrows():
        ticker = row["TICKER"]
        precio_compra = float(row["PRECIO_COMPRA"])
        contingency = float(row["CONTINGENCY"])
        barrera = precio_compra * contingency

        cierre = obtener_cierre_ticker_en_fecha(ticker, fecha_obs)

        if DEBUG:
            print("\n--- DEBUG OBSERVACIÓN ---")
            print("Nota:", int(nota))
            print("Ticker:", ticker)
            print("Precio compra:", precio_compra)
            print("Contingency:", contingency)
            print("Barrera:", barrera)
            print("Cierre usado:", cierre)
            print("Fecha observación:", fecha_obs.strftime("%d/%m/%Y"))

        if cierre is None:
            todo_ok = False
            detalles.append(f"{ticker}: SIN_DATO")
            continue

        cumple = cierre >= barrera
        estado = "OK" if cumple else "NO"

        if not cumple:
            todo_ok = False

        detalles.append(
            f"{ticker}: cierre={cierre:.2f} | barrera={barrera:.2f} -> {estado}"
        )

    resultado = "POSITIVA" if todo_ok else "NEGATIVA"
    detalle_texto = " ; ".join(detalles)

    return resultado, detalle_texto


def actualizar_resultado_observacion(resultados_df, nota, fecha_obs, resultado, detalle):
    fecha_obs = pd.to_datetime(fecha_obs).normalize()

    mask = (
        (resultados_df["NOTA"] == nota) &
        (pd.to_datetime(resultados_df["FECHA_OBSERVACION"], errors="coerce").dt.normalize() == fecha_obs)
    )

    if mask.any():
        resultados_df.loc[mask, "RESULTADO"] = resultado
        resultados_df.loc[mask, "DETALLE"] = detalle
    else:
        nueva = pd.DataFrame([{
            "NOTA": nota,
            "FECHA_OBSERVACION": fecha_obs,
            "RESULTADO": resultado,
            "DETALLE": detalle
        }])
        resultados_df = pd.concat([resultados_df, nueva], ignore_index=True)

    return resultados_df


def buscar_ultima_observacion_antes_de_pago(nota, fecha_pago, calendario_df):
    sub = calendario_df[
        (calendario_df["NOTA"] == nota) &
        (calendario_df["TIPO_EVENTO"] == "OBSERVACION") &
        (calendario_df["FECHA"] < fecha_pago)
    ].sort_values("FECHA")

    if sub.empty:
        return None

    return pd.to_datetime(sub.iloc[-1]["FECHA"]).normalize()


def imprimir_alerta_observacion(nota, fecha_obs, resultado, detalle):
    print("=" * 90)
    print(f"📅 HOY HAY OBSERVACIÓN | NOTA {int(nota)} | {fecha_obs.strftime('%d/%m/%Y')}")
    print(f"Resultado: {resultado}")
    print(f"Detalle: {detalle}")

    if resultado == "POSITIVA":
        print("✅ La observación ha sido positiva. Se pagará el cupón en la fecha de pago asociada.")
    else:
        print("❌ La observación ha sido negativa. No se pagará el cupón en la fecha de pago asociada.")


def imprimir_alerta_pago(nota, fecha_pago, fecha_obs, resultado_prev):
    print("=" * 90)
    print(f"💰 HOY HAY PAGO | NOTA {int(nota)} | {fecha_pago.strftime('%d/%m/%Y')}")

    if fecha_obs is None:
        print("⚠️ No se encontró una observación previa asociada.")
        return

    print(f"Observación asociada: {fecha_obs.strftime('%d/%m/%Y')}")

    if resultado_prev == "POSITIVA":
        print("✅ La observación previa fue positiva. HOY COBRARÁS.")
    elif resultado_prev == "NEGATIVA":
        print("❌ La observación previa fue negativa. HOY NO COBRARÁS.")
    else:
        print("⚠️ No hay resultado guardado para la observación previa.")


# =========================================================
# PROGRAMA PRINCIPAL
# =========================================================
def main():
    hoy = obtener_fecha_hoy()

    control_df, calendario_df = cargar_datos()
    resultados_df = cargar_resultados_existentes()

    eventos_hoy = calendario_df[calendario_df["FECHA"] == hoy].copy()

    if eventos_hoy.empty:
        print(f"No hay observaciones ni pagos para hoy ({hoy.strftime('%d/%m/%Y')}).")
        return

    # -----------------------------------------------------
    # 1) OBSERVACIONES DE HOY
    # -----------------------------------------------------
    observaciones_hoy = eventos_hoy[eventos_hoy["TIPO_EVENTO"] == "OBSERVACION"].copy()

    for _, row in observaciones_hoy.iterrows():
        nota = row["NOTA"]
        fecha_obs = row["FECHA"]

        resultado, detalle = evaluar_observacion(nota, fecha_obs, control_df)

        resultados_df = actualizar_resultado_observacion(
            resultados_df,
            nota,
            fecha_obs,
            resultado,
            detalle
        )

        imprimir_alerta_observacion(nota, fecha_obs, resultado, detalle)

    if not resultados_df.empty:
        resultados_df = resultados_df.sort_values(["NOTA", "FECHA_OBSERVACION"]).reset_index(drop=True)
        guardar_resultados(resultados_df)

    # -----------------------------------------------------
    # 2) PAGOS DE HOY
    # -----------------------------------------------------
    pagos_hoy = eventos_hoy[eventos_hoy["TIPO_EVENTO"] == "PAGO"].copy()

    for _, row in pagos_hoy.iterrows():
        nota = row["NOTA"]
        fecha_pago = row["FECHA"]

        fecha_obs = buscar_ultima_observacion_antes_de_pago(nota, fecha_pago, calendario_df)

        if fecha_obs is None:
            imprimir_alerta_pago(nota, fecha_pago, None, None)
            continue

        mask = (
            (resultados_df["NOTA"] == nota) &
            (pd.to_datetime(resultados_df["FECHA_OBSERVACION"], errors="coerce").dt.normalize() == fecha_obs)
        )

        if mask.any():
            resultado_prev = str(resultados_df.loc[mask, "RESULTADO"].iloc[0]).strip().upper()
        else:
            resultado_prev = None

        imprimir_alerta_pago(nota, fecha_pago, fecha_obs, resultado_prev)


if __name__ == "__main__":
    main()