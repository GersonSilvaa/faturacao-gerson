# fidelidade.py
import streamlit as st
from export_helpers import exportar_fidelidade_excel

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
                "Valor a Faturar S/IVA",
                "Descrição da Avaria no Local",
                "Serviço de Desempanagem no Local",
                "Contacto Técnico",
                "Observações",
                "Tipo de Avaria"
            ]
        )

        if df is not None:
            st.success("Ficheiro carregado com sucesso!")
            st.write("Pré-visualização dos dados importados:")
            st.dataframe(df.head())

            if st.button("Exportar Excel com Destaques"):
                output, filename = exportar_fidelidade_excel(df)
                st.download_button(
                    "Descarregar Excel Formatado",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
