import streamlit as st
import pandas as pd
from utils import processar_ficheiro
# import outros helpers no futuro conforme precisares

def run_acp():
    st.subheader("ACP 🚗")

    st.subheader("Upload do Ficheiro Exportado do Gestow")
    uploaded_file = st.file_uploader("Ficheiro de serviços da ACP (Gestow)", type=["xlsx"], key="acp_gestow")

    if uploaded_file:
        df = processar_ficheiro(uploaded_file)

        if df is not None:
            st.success("Ficheiro carregado com sucesso!")
            st.write("Pré-visualização dos dados:")
            st.dataframe(df, height=600)
