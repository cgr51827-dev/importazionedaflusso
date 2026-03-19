import streamlit as st

st.title("Compilatore BCC")

st.write("App online in preparazione 🚀")

uploaded_file = st.file_uploader("Carica file flusso")

if uploaded_file:
    st.success("File caricato correttamente!")
