import streamlit as st

st.title("Mi programa compartido")

nombre = st.text_input("Nombre")
numero = st.number_input("Número", min_value=0.0)

if st.button("Calcular"):
    resultado = numero * 2
    st.write(f"Hola {nombre}, el resultado es {resultado}")