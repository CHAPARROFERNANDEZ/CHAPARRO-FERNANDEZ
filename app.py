import streamlit as st
import pandas as pd

st.set_page_config(page_title="Sistema Fondo", layout="wide")

# =========================
# CARGA EXCEL (CORREGIDO)
# =========================
@st.cache_data
def leer_excel():
    archivo = "inversiones.xlsx"
    xls = pd.ExcelFile(archivo)

    data = {}
    for hoja in xls.sheet_names:
        data[hoja] = pd.read_excel(xls, sheet_name=hoja)

    return data


data = leer_excel()

# =========================
# INTERFAZ
# =========================
st.title("📊 Sistema Fondo")
st.write("Aplicación conectada al Excel inversiones.xlsx")

menu = st.sidebar.selectbox(
    "Selecciona una sección",
    [
        "Inicio",
        "Ver Excel",
    ]
)

# =========================
# FUNCIONES
# =========================
if menu == "Inicio":
    st.subheader("Bienvenido")
    st.write("Sistema conectado correctamente.")

elif menu == "Ver Excel":
    hoja = st.selectbox("Selecciona hoja", list(data.keys()))
    st.dataframe(data[hoja])
