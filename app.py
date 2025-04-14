import streamlit as st
import pandas as pd
from datetime import datetime
import math
import io
import json

# ----- Gestor de Utilizadores -----
utilizadores = st.secrets["utilizadores"]

# ----- Funções auxiliares -----
def verificar_login():
    user = st.text_input("Utilizador")
    password = st.text_input("Password", type="password")
    if st.button("Entrar"):
        if user in utilizadores and utilizadores[user] == password:
            st.session_state['login'] = True
            st.success("Login feito com sucesso!")
        else:
            st.error("Credenciais inválidas!")


def processar_ficheiro(uploaded_file, colunas_obrigatorias=None):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip()  # Limpa espaços em branco nos nomes das colunas
        st.success("Ficheiro carregado com sucesso!")

        # Verificação de colunas obrigatórias
        if colunas_obrigatorias:
            colunas_em_falta = [col for col in colunas_obrigatorias if col not in df.columns]
            if colunas_em_falta:
                st.error(f"Atenção! O ficheiro está a faltar as colunas: {', '.join(colunas_em_falta)}")
                return None

        return df
    else:
        st.warning("Por favor, carrega o ficheiro Excel.")
        return None


def exportar_listas(prontos_faturar, num_listas):
    prontos_faturar = prontos_faturar.sort_values(by='match_key')
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    workbook = writer.book
    worksheet = workbook.add_worksheet('Listas')
    writer.sheets['Listas'] = worksheet

    total_geral = 0
    current_col = 0
    listas = [prontos_faturar.iloc[i::num_listas] for i in range(num_listas)]

    for idx, lista in enumerate(listas):
        row = 0
        worksheet.write(row, current_col, f'Lista {idx + 1}')
        row += 1
        worksheet.write(row, current_col, 'match_key')
        worksheet.write(row, current_col + 1, 'Total')
        row += 1

        subtotal = 0
        for _, processo in lista.iterrows():
            worksheet.write(row, current_col, processo['match_key'])
            worksheet.write(row, current_col + 1, processo['Total'])
            subtotal += processo['Total']
            row += 1

        worksheet.write(row, current_col, 'Subtotal:')
        worksheet.write(row, current_col + 1, subtotal)
        total_geral += subtotal
        current_col += 3

    worksheet.write(row + 2, 0, 'Total Geral:')
    worksheet.write(row + 2, 1, total_geral)

    writer.close()
    output.seek(0)

    data_atual = datetime.now().strftime("%Y-%m-%d")
    return output, f"Listagem_Para_Faturar_{data_atual}.xlsx"


def exportar_divergencias(df, referencia):
    divergencias = df[df['Diferenca'] != 0].copy()

    def verificar_agravamento(data):
        if pd.isna(data):
            return "Não"
        if data.weekday() >= 5 or data.hour < 7 or data.hour >= 20:
            return "Sim"
        return "Não"

    divergencias['Data_Requisicao'] = pd.to_datetime(divergencias['Data_Requisicao'], errors='coerce')
    divergencias['Agravamento'] = divergencias['Data_Requisicao'].apply(verificar_agravamento)

    # Cruzar com a referência para buscar Marca, Modelo e Categoria
    if referencia is not None:
        referencia = referencia.rename(columns={
            'Marca': 'Marca',
            'Modelo': 'Modelo',
            'Categoria de Veículo': 'Categoria de Veículo'
        })
        divergencias = divergencias.merge(
            referencia[['Matrícula', 'Marca', 'Modelo', 'Categoria de Veículo']],
            how='left',
            left_on='Matrícula',
            right_on='Matrícula'
        )

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    divergencias.to_excel(writer, index=False, sheet_name='Divergencias')

    workbook = writer.book
    worksheet = writer.sheets['Divergencias']

    format_agravado = workbook.add_format({'bg_color': '#FFFF00'})  # Amarelo
    format_faltante = workbook.add_format({'bg_color': '#FFFF00'})  # Amarelo para células em falta

    headers = list(divergencias.columns)
    col_agravamento_idx = headers.index("Agravamento")

    # Destacar agravamento
    for row_num, agravamento in enumerate(divergencias['Agravamento'], start=1):
        if agravamento == "Sim":
            worksheet.write(row_num, col_agravamento_idx, agravamento, format_agravado)

    # Destacar células em falta nas novas colunas
    for row_num, row in divergencias.iterrows():
        if pd.isna(row.get('Marca')):
            worksheet.write(row_num + 1, headers.index('Marca'), '', format_faltante)
        if pd.isna(row.get('Modelo')):
            worksheet.write(row_num + 1, headers.index('Modelo'), '', format_faltante)
        if pd.isna(row.get('Categoria de Veículo')):
            worksheet.write(row_num + 1, headers.index('Categoria de Veículo'), '', format_faltante)

    writer.close()
    output.seek(0)

    data_atual = datetime.now().strftime("%Y-%m-%d")
    return output, f"Listagem_Divergencias_{data_atual}.xlsx"


# ----- Início da aplicação -----
st.title("Gestão de Faturação - IPA 🚛")

if 'login' not in st.session_state:
    st.session_state['login'] = False

if not st.session_state['login']:
    st.subheader("Login de Acesso")
    verificar_login()
else:
    st.subheader("Upload do Ficheiro de Comparação")
    uploaded_file = st.file_uploader("Escolhe o ficheiro Excel de comparação", type=["xlsx"])

    st.subheader("Upload do Ficheiro de Referência (Matrículas + Marca/Modelo/Categoria)")
    referencia_file = st.file_uploader("Escolhe o ficheiro de referência", type=["xlsx"])

    referencia_df = None
    if referencia_file:
        referencia_df = processar_ficheiro(
            referencia_file,
            colunas_obrigatorias=["Matrícula", "Marca", "Modelo", "Categoria de Veículo"]
        )
        if referencia_df is not None:
            st.write("Pré-visualização do ficheiro de referência:")
            st.dataframe(referencia_df.head())

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

                num_listas = st.number_input("Escolhe o número de listas que queres gerar:", min_value=1, max_value=total_processos, value=1, step=1)

                if st.button("Exportar Listas para Faturação"):
                    output, filename = exportar_listas(prontos_faturar, int(num_listas))
                    st.download_button("Descarregar Excel das Listas", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            else:
                st.warning("Não existem processos prontos a faturar neste ficheiro.")

            if referencia_df is not None:
                if st.button("Exportar Divergências para Análise"):
                    output, filename = exportar_divergencias(df, referencia_df)
                    st.download_button("Descarregar Excel das Divergências", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("Carrega primeiro o ficheiro de referência para exportar divergências!")

    else:
        st.info("Aguardo o carregamento do ficheiro Excel de comparação.")
