import streamlit as st
import pandas as pd
import math
import io
from datetime import datetime
from utils import processar_ficheiro
from export_helpers import exportar_listas, exportar_divergencias, exportar_cruzamento_weboffice


def run_ipa():
    st.title("Gestão de Faturação - IPA 🚛")

    st.subheader("Upload do Ficheiro de Comparação")
    uploaded_file = st.file_uploader("Escolhe o ficheiro Excel de comparação", type=["xlsx"], key="comparacao")

    st.subheader("Upload do Ficheiro de Referência (com colunas Matrícula + Marca/Modelo/Categoria + KMS + Valor a Faturar S/IVA)")
    referencia_file = st.file_uploader("Escolhe o ficheiro de referência", type=["xlsx"], key="referencia")
    
    st.subheader("Upload do Ficheiro WebOffice (Portal IPA)")
    weboffice_file = st.file_uploader("Ficheiro WebOffice (com Dossier e Total)", type=["xlsx"], key="weboffice")

    referencia_df = None
    if referencia_file:
        referencia_df = processar_ficheiro(
            referencia_file,
            colunas_obrigatorias=[
                "Matrícula",
                "Marca",
                "Modelo",
                "Categoria de Veículo",
                "KMS a Faturar no Serviço",
                "Valor a Faturar S/IVA"
            ]
        )
        if referencia_df is not None:
            st.write("Pré-visualização do ficheiro de referência:")
            st.dataframe(referencia_df.head())
            st.subheader("Upload do Ficheiro WebOffice (Portal IPA)")
            weboffice_file = st.file_uploader("Ficheiro WebOffice (com Dossier e Total)", type=["xlsx"], key="weboffice")
            
            if weboffice_file is not None:
                weboffice_df = processar_ficheiro(
                weboffice_file,
                colunas_obrigatorias=["Dossier", "Total"]
            )

            if weboffice_df is not None:
                st.write("Pré-visualização WebOffice:")
                st.dataframe(weboffice_df, height=300)
                if st.button("Exportar Cruzamento WebOffice vs Gestow"):
                    output, filename = exportar_cruzamento_weboffice(weboffice_df, referencia_df)
                    st.download_button(
                        "Descarregar Excel Cruzado",
                        data=output,
                        file_name=filename,
                     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            

    if uploaded_file:
        df = processar_ficheiro(uploaded_file)

        if df is not None:
            st.write("Pré-visualização dos dados de comparação:")
            st.dataframe(df.head())

            prontos_faturar = df[df['Diferenca'] == 0]
            total_processos = len(prontos_faturar)

            if total_processos > 0:
                st.success(f"Total de processos prontos a faturar: {total_processos}")

                sugestoes = []
                for i in range(1, min(5, total_processos + 1)):
                    tamanho_lista = math.ceil(total_processos / i)
                    sugestoes.append(f"{i} lista(s) de aprox. {tamanho_lista} processos cada")

                st.write("Sugestões de divisão:")
                for sugestao in sugestoes:
                    st.write("- ", sugestao)

                num_listas = st.number_input(
                    "Escolhe o número de listas que queres gerar:",
                    min_value=1,
                    max_value=total_processos,
                    value=1,
                    step=1
                )

                if st.button("Exportar Listas para Faturação"):
                    output, filename = exportar_listas(prontos_faturar, int(num_listas))
                    st.download_button(
                        "Descarregar Excel das Listas",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            else:
                st.warning("Não existem processos prontos a faturar neste ficheiro.")

            if referencia_df is not None:
                if st.button("Exportar Divergências para Análise"):
                    output, filename = exportar_divergencias(df, referencia_df)
                    st.download_button(
                        "Descarregar Excel das Divergências",
                        data=output,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            else:
                st.warning("Carrega primeiro o ficheiro de referência para exportar divergências!")

    else:
        st.info("Aguardo o carregamento do ficheiro Excel de comparação.")


