import streamlit as st

st.set_page_config(page_title="Sistema Fondo", layout="wide")

st.title("Sistema Fondo")

menu = st.sidebar.selectbox(
    "Selecciona una sección",
    [
        "Inicio",
        "Alertas notas",
        "Alertas semana",
        "Calendario notas",
        "Consultas fútbol",
        "Consultas motoclick",
        "Consultas notas",
        "Consultas Paraguay",
        "Extractos",
        "Notas",
        "Sistema fondo",
    ]
)

if menu == "Inicio":
    st.write("Bienvenido al sistema.")

elif menu == "Alertas notas":
    import alertas_notas
    alertas_notas.main()

elif menu == "Alertas semana":
    import alertas_semana
    alertas_semana.main()

elif menu == "Calendario notas":
    import calendario_notas
    calendario_notas.main()

elif menu == "Consultas fútbol":
    import consultas_futbol
    consultas_futbol.main()

elif menu == "Consultas motoclick":
    import consultas_motoclick
    consultas_motoclick.main()

elif menu == "Consultas notas":
    import consultas_notas
    consultas_notas.main()

elif menu == "Consultas Paraguay":
    import consultas_paraguay
    consultas_paraguay.main()

elif menu == "Extractos":
    import extractos
    extractos.main()

elif menu == "Notas":
    import notas
    notas.main()

elif menu == "Sistema fondo":
    import sistema_fondo
    sistema_fondo.main()