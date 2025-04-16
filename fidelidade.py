import streamlit as st
import pandas as pd
from utils import processar_ficheiro, FIDELIDADE_COLUNAS_EXTRA, contem_texto_suspeito
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
            # Verificar se há palavras suspeitas
            df["Tem_Alerta"] = df.apply(lambda row: contem_texto_suspeito(row, FIDELIDADE_COLUNAS_EXTRA), axis=1)
            num_alertas = df["Tem_Alerta"].sum()
            st.warning(f"Foram encontrados {num_alertas} serviço(s) com possíveis alertas de atenção.")

            # Filtro para mostrar só serviços com alerta
            mostrar_so_alertas = st.checkbox("Mostrar apenas serviços com alerta")
            df_filtrado = df[df["Tem_Alerta"]] if mostrar_so_alertas else df

            # Pré-visualização
            st.write("Pré-visualização dos dados importados:")
            st.dataframe(df_filtrado)

            # Exportação geral
            if st.button("Exportar Excel com Destaques (Todos os Serviços)"):
                output, filename = exportar_fidelidade_excel(df)
                st.download_button(
                    "Descarregar Excel Formatado",
                    data=output,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            # Exportação só de alertas
            if num_alertas > 0:
                if st.button("Exportar Apenas Serviços com Alerta"):
                    df_alertas = df[df["Tem_Alerta"]]
                    output_alertas, filename_alertas = exportar_fidelidade_excel(df_alertas)
                    st.download_button(
                        "Descarregar Excel Só com Alertas",
                        data=output_alertas,
                        file_name="Alertas_Fidelidade_Formatado.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
