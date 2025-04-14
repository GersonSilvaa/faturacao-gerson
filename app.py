import streamlit as st
import pandas as pd
from datetime import datetime
import math
import io

# ----- Gestor de Utilizadores -----
utilizadores = {
    "gerson": "gerson123",
    "filipe": "filipe123",
    "catarina": "catarina123",
    "andre": "andre123"
}

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


def processar_ficheiro(uploaded_file):
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.success("Ficheiro carregado com sucesso!")
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


def exportar_divergencias(df):
    divergencias = df[df['Diferenca'] != 0].copy()

    def verificar_agravamento(data):
        if pd.isna(data):
            return "Não"
        if data.weekday() >= 5 or data.hour < 7 or data.hour >= 20:
            return "Sim"
        return "Não"

    divergencias['Data_IPA'] = pd.to_datetime(divergencias['Data_IPA'], errors='coerce')
    divergencias['Agravamento'] = divergencias['Data_IPA'].apply(verificar_agravamento)
    total_agravados = divergencias['Agravamento'].value_counts().get("Sim", 0)

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    divergencias.to_excel(writer, index=False, sheet_name='Divergencias')

    workbook = writer.book
    worksheet = writer.sheets['Divergencias']

    format_agravado = workbook.add_format({'bg_color': '#FFC7CE'})
    for row_num, agravamento in enumerate(divergencias['Agravamento'], start=1):
        if agravamento == "Sim":
            worksheet.set_row(row_num, cell_format=format_agravado)

    last_row = len(divergencias) + 2
    worksheet.write(f'A{last_row}', 'Total de processos com agravamento:')
    worksheet.write(f'B{last_row}', total_agravados)

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
    uploaded_file = st.file_uploader("Escolhe o ficheiro Excel", type=["xlsx"])

    if uploaded_file:
        df = processar_ficheiro(uploaded_file)

        if df is not None:
            st.write("Pré-visualização dos dados:")
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

            if st.button("Exportar Divergências para Análise"):
                output, filename = exportar_divergencias(df)
                st.download_button("Descarregar Excel das Divergências", data=output, file_name=filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    else:
        st.info("Aguardo o carregamento do ficheiro Excel.")
