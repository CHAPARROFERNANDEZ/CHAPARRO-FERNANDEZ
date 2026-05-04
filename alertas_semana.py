import os
from datetime import timedelta
import pandas as pd

# =========================================================
# CONFIGURACIÓN
# =========================================================
ARCHIVO_EXCEL = "inversiones.xlsx"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"

# Si quieres probar una fecha concreta, ponla aquí:
# FECHA_MANUAL = "2026-04-08"
FECHA_MANUAL = None


# =========================================================
# FUNCIONES AUXILIARES
# =========================================================
def obtener_fecha_base():
    if FECHA_MANUAL:
        return pd.to_datetime(FECHA_MANUAL).normalize()
    return pd.Timestamp.today().normalize()


def normalizar_columnas(df):
    df.columns = [
        str(col).strip().upper().replace(" ", "_")
        for col in df.columns
    ]
    return df


def cargar_calendario():
    if not os.path.exists(ARCHIVO_EXCEL):
        raise FileNotFoundError(f"No existe el archivo {ARCHIVO_EXCEL}")

    calendario = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_CALENDARIO)
    calendario = normalizar_columnas(calendario)

    columnas_necesarias = ["NOTA", "TIPO_EVENTO", "FECHA"]
    for col in columnas_necesarias:
        if col not in calendario.columns:
            raise ValueError(f"En la hoja {HOJA_CALENDARIO} falta la columna {col}")

    calendario["NOTA"] = pd.to_numeric(calendario["NOTA"], errors="coerce")
    calendario["TIPO_EVENTO"] = calendario["TIPO_EVENTO"].astype(str).str.strip().str.upper()
    calendario["FECHA"] = pd.to_datetime(
        calendario["FECHA"],
        dayfirst=True,
        errors="coerce"
    ).dt.normalize()

    calendario = calendario.dropna(subset=["NOTA", "TIPO_EVENTO", "FECHA"]).copy()
    calendario = calendario[calendario["TIPO_EVENTO"].isin(["OBSERVACION", "PAGO"])].copy()

    return calendario


def mostrar_eventos_semana(calendario_df, fecha_base):
    fecha_inicio = pd.to_datetime(fecha_base).normalize()
    fecha_fin = fecha_inicio + timedelta(days=6)

    eventos_semana = calendario_df[
        (calendario_df["FECHA"] >= fecha_inicio) &
        (calendario_df["FECHA"] <= fecha_fin)
    ].copy()

    print("=" * 90)
    print(f"📆 NOVEDADES DE LA SEMANA | DEL {fecha_inicio.strftime('%d/%m/%Y')} AL {fecha_fin.strftime('%d/%m/%Y')}")
    print("=" * 90)

    if eventos_semana.empty:
        print("No hay observaciones ni pagos esta semana.")
        return

    eventos_semana = eventos_semana.sort_values(["FECHA", "TIPO_EVENTO", "NOTA"]).reset_index(drop=True)

    observaciones = eventos_semana[eventos_semana["TIPO_EVENTO"] == "OBSERVACION"].copy()
    pagos = eventos_semana[eventos_semana["TIPO_EVENTO"] == "PAGO"].copy()

    if not observaciones.empty:
        print("\n📅 OBSERVACIONES DE ESTA SEMANA:")
        fecha_actual = None
        for _, row in observaciones.iterrows():
            fecha_evento = row["FECHA"]
            nota = int(row["NOTA"])

            if fecha_actual is None or fecha_evento != fecha_actual:
                fecha_actual = fecha_evento
                print(f"\n{fecha_evento.strftime('%d/%m/%Y')}")
            print(f"   - NOTA {nota}")

    if not pagos.empty:
        print("\n💰 COBROS DE ESTA SEMANA:")
        fecha_actual = None
        for _, row in pagos.iterrows():
            fecha_evento = row["FECHA"]
            nota = int(row["NOTA"])

            if fecha_actual is None or fecha_evento != fecha_actual:
                fecha_actual = fecha_evento
                print(f"\n{fecha_evento.strftime('%d/%m/%Y')}")
            print(f"   - NOTA {nota}")

    print("\nℹ️ Este resumen solo muestra lo que ocurrirá esta semana.")
    print("   No evalúa si se cobrará o no, porque eso depende del cierre del día de observación.")


# =========================================================
# PROGRAMA PRINCIPAL
# =========================================================
def main():
    fecha_base = obtener_fecha_base()
    calendario_df = cargar_calendario()
    mostrar_eventos_semana(calendario_df, fecha_base)


if __name__ == "__main__":
    main()