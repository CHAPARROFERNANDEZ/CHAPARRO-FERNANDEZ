import pandas as pd
import yfinance as yf

ARCHIVO = "inversiones.xlsx"
HOJA = "CONTROL_NOTAS"


def obtener_precio_actual(ticker):
    try:
        data = yf.Ticker(ticker)
        hist = data.history(period="5d")

        if hist.empty:
            return None

        precio = hist["Close"].dropna().iloc[-1]
        return round(float(precio), 2)

    except:
        return None


def sistema_notas_pro():
    df = pd.read_excel(ARCHIVO, sheet_name=HOJA)
    df.columns = df.columns.str.strip().str.upper()

    # =========================
    # PRECIOS
    # =========================
    df["PRECIO_ACTUAL"] = df["TICKER"].apply(obtener_precio_actual)

    # =========================
    # VARIACIÓN
    # =========================
    df["VARIACION"] = (
        (df["PRECIO_ACTUAL"] - df["PRECIO_COMPRA"])
        / df["PRECIO_COMPRA"]
    ) * 100

    df["VARIACION"] = df["VARIACION"].round(2)

    # =========================
    # PRECIO CONTINGENCIA
    # =========================
    df["PRECIO_CONTINGENCIA"] = (
        df["PRECIO_COMPRA"] * df["BARRERA_CAPITAL"]
    ).round(2)

    # =========================
    # CONTINGENCIA
    # =========================
    df["CONTINGENCIA"] = df.apply(
        lambda x: "OK" if x["PRECIO_ACTUAL"] >= x["PRECIO_CONTINGENCIA"]
        else "RIESGO",
        axis=1
    )

    # =========================
    # PRINT BONITO
    # =========================
    print("\n" + "="*100)
    print("📊 RESUMEN NOTAS".center(100))
    print("="*100 + "\n")

    nota_actual = None

    for _, row in df.iterrows():

        if row["NOTA"] != nota_actual:
            nota_actual = row["NOTA"]
            print(f"\n🔹 NOTA {int(nota_actual)}")
            print("-"*100)

        # Variación colores
        var = row["VARIACION"]
        if var >= 0:
            var_str = f"\033[92m{var:.2f}%\033[0m"
        else:
            var_str = f"\033[91m{var:.2f}%\033[0m"

        # Contingencia colores
        estado = row["CONTINGENCIA"]
        if estado == "OK":
            estado_str = "\033[92mOK\033[0m"
        else:
            estado_str = "\033[91mRIESGO\033[0m"

        print(
            f"{row['TICKER']:<6} | "
            f"Compra: {row['PRECIO_COMPRA']:>8.2f} | "
            f"Actual: {row['PRECIO_ACTUAL']:>8.2f} | "
            f"Contingencia: {row['PRECIO_CONTINGENCIA']:>8.2f} | "
            f"Var: {var_str:>10} | "
            f"{estado_str}"
        )

    print("\n" + "="*100)

    # =========================
    # 🚨 ALERTAS
    # =========================
    print("\n🚨 ALERTAS IMPORTANTES\n")

    alertas = []

    for nota, grupo in df.groupby("NOTA"):
        en_riesgo = grupo[grupo["CONTINGENCIA"] == "RIESGO"]

        if not en_riesgo.empty:
            tickers_riesgo = ", ".join(en_riesgo["TICKER"])
            alertas.append(f"⚠️ NOTA {nota} en riesgo por: {tickers_riesgo}")

    if alertas:
        for a in alertas:
            print(a)
    else:
        print("✅ Ninguna nota en riesgo")

    # =========================
    # 📉 ANÁLISIS GLOBAL
    # =========================
    print("\n📊 ANÁLISIS GLOBAL\n")

    peor = df.loc[df["VARIACION"].idxmin()]
    mejor = df.loc[df["VARIACION"].idxmax()]

    print(f"📉 Peor ticker: {peor['TICKER']} ({peor['VARIACION']}%)")
    print(f"📈 Mejor ticker: {mejor['TICKER']} (+{mejor['VARIACION']}%)")

    media = df["VARIACION"].mean()
    print(f"\n🏦 Rentabilidad media del sistema: {round(media,2)}%")

    print("\n" + "="*100)


# Ejecutar
sistema_notas_pro()