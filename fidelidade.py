import streamlit as st
import pandas as pd
from utils import processar_ficheiro

def run_fidelidade():
    st.title("Gestão de Faturação - FIDELIDADE 🛡️")

    st.subheader("Upload do Ficheiro Exportado do Gestow")
    uploaded_file = st.file_uploader("Ficheiro de serviços da Fidelidade (Gestow)", type=["xlsx"], key="fidelidade_gestow")

    if uploaded_file:
        df = processar_ficheiro(
            uploaded_file,
            colunas_obrigatorias=[
                "Matrícula",
                "Marca",
                "Modelo",
                "Categoria de Veículo",
                "KMS a Faturar no Serviço",
                "Valor a Faturar S/IVA"
            ]
        )

        if df is not None:
            st.success("Ficheiro carregado com sucesso!")
            st.write("Pré-visualização dos dados importados:")
            st.dataframe(df.head())
