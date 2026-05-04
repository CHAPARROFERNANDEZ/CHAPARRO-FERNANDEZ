import streamlit as st
import pandas as pd

st.set_page_config(page_title="Sistema Fondo", layout="wide")

@st.cache_data
def cargar_datos():
    return pd.read_excel("inversiones.xlsx", sheet_name="INVERSIONES")

df = cargar_datos()

st.title("📊 Sistema Fondo")

menu = st.sidebar.selectbox(
    "Menú",
    ["Inicio", "Consultas Fútbol"]
)

if menu == "Inicio":
    st.write("Sistema funcionando correctamente")

elif menu == "Consultas Fútbol":

    st.subheader("Consulta ingresos fútbol")

    año = st.number_input("Año", value=2025)
    mes = st.number_input("Mes (1-12)", min_value=1, max_value=12, value=1)

    if st.button("Calcular"):

        df_futbol = df[
            df["subtipo_inversion"].str.lower().str.contains("futbol", na=False)
        ]

        total = df_futbol["capital_invertido"].sum()

        st.success(f"Capital total en fútbol: {total:,.2f} €")
        st.dataframe(df_futbol)
