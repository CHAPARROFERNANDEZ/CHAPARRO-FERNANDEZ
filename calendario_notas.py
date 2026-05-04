import pandas as pd
from datetime import datetime, timedelta

ARCHIVO_EXCEL = "inversiones.xlsx"
HOJA_CALENDARIO = "CALENDARIO_NOTAS"


def cargar_calendario():
    df = pd.read_excel(ARCHIVO_EXCEL, sheet_name=HOJA_CALENDARIO)

    df.columns = [str(c).strip().upper() for c in df.columns]

    df["FECHA"] = pd.to_datetime(df["FECHA"], errors="coerce")
    df = df.dropna(subset=["FECHA"])

    df["NOTA"] = df["NOTA"].astype(int)
    df["TIPO_EVENTO"] = df["TIPO_EVENTO"].astype(str).str.upper().str.strip()

    return df


def mostrar_eventos(df):
    if df.empty:
        print("\nNo hay observaciones ni pagos en este periodo.")
        return

    df = df.sort_values(["FECHA", "NOTA", "TIPO_EVENTO"])

    print("\nEVENTOS ENCONTRADOS:")
    print("-" * 70)

    for _, fila in df.iterrows():
        fecha = fila["FECHA"].strftime("%d/%m/%Y")
        nota = fila["NOTA"]
        tipo = fila["TIPO_EVENTO"]

        print(f"{fecha} | NOTA {nota} | {tipo}")

    print("-" * 70)
    print(f"Total eventos: {len(df)}")


def calendario_esta_semana(df):
    hoy = datetime.today()

    inicio_semana = hoy - timedelta(days=hoy.weekday())
    fin_semana = inicio_semana + timedelta(days=6)

    eventos = df[
        (df["FECHA"] >= inicio_semana.replace(hour=0, minute=0, second=0)) &
        (df["FECHA"] <= fin_semana.replace(hour=23, minute=59, second=59))
    ]

    print(f"\nCALENDARIO DE ESTA SEMANA")
    print(f"Del {inicio_semana.strftime('%d/%m/%Y')} al {fin_semana.strftime('%d/%m/%Y')}")

    mostrar_eventos(eventos)


def calendario_mes_completo(df):
    anio = int(input("Introduce el año: "))
    mes = int(input("Introduce el mes (1-12): "))

    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = inicio_mes + pd.offsets.MonthEnd(0)

    eventos_mes = df[
        (df["FECHA"] >= inicio_mes) &
        (df["FECHA"] <= fin_mes)
    ].copy()

    if eventos_mes.empty:
        print("\nNo hay eventos en ese mes.")
        return

    eventos_mes["SEMANA_MES"] = ((eventos_mes["FECHA"].dt.day - 1) // 7) + 1

    print(f"\nCALENDARIO COMPLETO DE {mes}/{anio}")

    for semana in sorted(eventos_mes["SEMANA_MES"].unique()):
        eventos_semana = eventos_mes[eventos_mes["SEMANA_MES"] == semana]

        print(f"\nSEMANA {semana}")
        mostrar_eventos(eventos_semana)


def calendario_mes_semana(df):
    anio = int(input("Introduce el año: "))
    mes = int(input("Introduce el mes (1-12): "))
    semana = int(input("Introduce la semana del mes (1, 2, 3, 4 o 5): "))

    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = inicio_mes + pd.offsets.MonthEnd(0)

    eventos_mes = df[
        (df["FECHA"] >= inicio_mes) &
        (df["FECHA"] <= fin_mes)
    ].copy()

    eventos_mes["SEMANA_MES"] = ((eventos_mes["FECHA"].dt.day - 1) // 7) + 1

    eventos_semana = eventos_mes[eventos_mes["SEMANA_MES"] == semana]

    print(f"\nCALENDARIO DE {mes}/{anio} - SEMANA {semana}")

    mostrar_eventos(eventos_semana)


def exportar_mes_a_excel(df):
    anio = int(input("Introduce el año: "))
    mes = int(input("Introduce el mes (1-12): "))

    inicio_mes = pd.Timestamp(anio, mes, 1)
    fin_mes = inicio_mes + pd.offsets.MonthEnd(0)

    eventos_mes = df[
        (df["FECHA"] >= inicio_mes) &
        (df["FECHA"] <= fin_mes)
    ].copy()

    if eventos_mes.empty:
        print("\nNo hay eventos para exportar en ese mes.")
        return

    eventos_mes["SEMANA_MES"] = ((eventos_mes["FECHA"].dt.day - 1) // 7) + 1
    eventos_mes["FECHA"] = eventos_mes["FECHA"].dt.strftime("%d/%m/%Y")

    eventos_mes = eventos_mes[
        ["SEMANA_MES", "NOTA", "TIPO_EVENTO", "FECHA"]
    ].sort_values(["SEMANA_MES", "FECHA", "NOTA"])

    nombre_archivo = f"calendario_notas_{mes}_{anio}.xlsx"
    eventos_mes.to_excel(nombre_archivo, index=False)

    print(f"\nArchivo exportado correctamente: {nombre_archivo}")


def menu():
    df = cargar_calendario()

    while True:
        print("\n==============================")
        print("CONSULTAS CALENDARIO NOTAS")
        print("==============================")
        print("1. Ver observaciones y pagos de esta semana")
        print("2. Ver un mes completo dividido por semanas")
        print("3. Ver una semana concreta de un mes")
        print("4. Exportar calendario de un mes a Excel")
        print("0. Salir")

        opcion = input("\nElige una opción: ").strip()

        if opcion == "1":
            calendario_esta_semana(df)

        elif opcion == "2":
            calendario_mes_completo(df)

        elif opcion == "3":
            calendario_mes_semana(df)

        elif opcion == "4":
            exportar_mes_a_excel(df)

        elif opcion == "0":
            print("\nSaliendo del programa.")
            break

        else:
            print("\nOpción no válida.")


if __name__ == "__main__":
    menu()