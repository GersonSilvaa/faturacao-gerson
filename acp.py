import streamlit as st
import pandas as pd
from utils import processar_ficheiro
from export_helpers import exportar_acp_corrigido


def run_acp():
    st.subheader("ACP 🚗")

    st.subheader("Upload do Ficheiro do ACP (com Matrícula e Interv.)")
    acp_file = st.file_uploader("Ficheiro ACP", type=["xlsx"], key="acp_file")

    st.subheader("Upload do Ficheiro do Gestow (com Matrícula e Processo da Companhia)")
    gestow_file = st.file_uploader("Ficheiro Gestow", type=["xlsx"], key="gestow_file_acp")

    acp_df = None
    gestow_df = None

    if acp_file:
        acp_df = processar_ficheiro(acp_file, colunas_obrigatorias=["Matrícula", "Interv."])
        if acp_df is not None:
            st.success("Ficheiro do ACP carregado com sucesso!")
            st.dataframe(acp_df.head(), height=250)

    if gestow_file:
        gestow_df = processar_ficheiro(gestow_file, colunas_obrigatorias=["Matrícula", "Processo da Companhia"])
        if gestow_df is not None:
            st.success("Ficheiro do Gestow carregado com sucesso!")
            st.dataframe(gestow_df.head(), height=250)

    if acp_df is not None and gestow_df is not None:
        if st.button("Exportar Ficheiro Corrigido ACP"):
            output, filename = exportar_acp_corrigido(acp_df, gestow_df)
            st.download_button(
                label="Descarregar CSV Corrigido",
                data=output,
                file_name=filename,
                mime="text/csv"
            )
