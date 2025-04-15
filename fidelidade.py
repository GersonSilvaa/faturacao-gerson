import streamlit as st
import pandas as pd
from utils import processar_ficheiro

def run_fidelidade():
    st.title("Gestão de Faturação - FIDELIDADE 🛡️")

    st.subheader("Upload do Ficheiro de Comparação")
    uploaded_file = st.file_uploader("Ficheiro de comparação (FIDELIDADE)", type=["xlsx"], key="fidelidade_comparacao")

    if uploaded_file:
        df = processar_ficheiro(uploaded_file)
        if df is not None:
            st.success("Ficheiro da Fidelidade carregado com sucesso!")
            st.write("Pré-visualização dos dados:")
            st.dataframe(df.head())
