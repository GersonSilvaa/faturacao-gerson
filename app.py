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
    login_button = st.button("Entrar")

    if login_button:
        if user in utilizadores and utilizadores[user] == password:
            st.session_state['login'] = True
            st.rerun()  # força o reload com o login feito
        else:
            st.error("Credenciais inválidas!")


def processar_ficheiro(uploaded_file, colunas_obrigatorias=None):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()  # Limpa espaços em branco nos nomes das colunas
            st.success("Ficheiro carregado com sucesso!")

            if colunas_obrigatorias:
                colunas_em_falta = [col for col in colunas_obrigatorias if col not in df.columns]
                if colunas_em_falta:
                    st.error(f"Atenção! O ficheiro está a faltar as colunas: {', '.join(colunas_em_falta)}")
                    return None

            return df
        except Exception:
            st.error("Erro ao ler o ficheiro. Verifica se é um ficheiro Excel válido (.xlsx).")
            return None
    else:
        st.warning("Por favor, carrega o ficheiro Excel.")
        return None


# ----- Funções de Cálculo -----

def calcular_valor_categoria(categoria, kms, agravamento):
    """
    Calcula o valor a faturar para Furgão ou Rodado Duplo,
    de acordo com as regras definidas:
      - Furgão: 30€ (fixo) + 0.40€/km acima dos 20km, +25% se agravamento
      - Rodado Duplo: 42€ (fixo) + 0.58€/km acima dos 20km, +25% se agravamento
    """
    if pd.isna(kms):
        return 0.0

    valor = 0.0
    if categoria == 'Furgão':
        if kms <= 20:
            valor = 30
        else:
            valor = 30 + (0.40 * (kms - 20))
    elif categoria == 'Rodado Duplo':
        if kms <= 20:
            valor = 42
        else:
            valor = 42 + (0.58 * (kms - 20))
    else:
        valor = 0

    if agravamento == 'Sim':
        valor *= 1.25

    return valor


def calcular_upgrade(row):
    """
    Se 'Categoria de Veículo' for 'Ligeiro', testa Furgão.
    Se for 'Furgão', testa Rodado Duplo.
    Caso contrário, fica sem upgrade.
    Retorna (valor_potencial, diferenca, sugestao).
    """
    cat_atual = row.get('Categoria de Veículo')
    kms = row.get('KMS a Faturar no Serviço')
    agravamento = row.get('Agravamento', 'Não')
    valor_base = row.get('Valor a Faturar S/IVA', 0.0)

    if pd.isna(valor_base):
        valor_base = 0.0

    # Ligeiro -> Furgão
    if cat_atual == 'Ligeiro':
        valor_pot = calcular_valor_categoria('Furgão', kms, agravamento)
        dif = valor_pot - valor_base
        if dif > 0:
            return (valor_pot,
                    dif,
                    f"Analisar upgrade p/ Furgão (+{dif:.2f}€)")
        else:
            return (valor_pot, dif, "")
    # Furgão -> Rodado Duplo
    elif cat_atual == 'Furgão':
        valor_pot = calcular_valor_categoria('Rodado Duplo', kms, agravamento)
        dif = valor_pot - valor_base
        if dif > 0:
            return (valor_pot,
                    dif,
                    f"Analisar upgrade p/ Rodado Duplo (+{dif:.2f}€)")
        else:
            return (valor_pot, dif, "")
    else:
        # Se já for Rodado Duplo ou outra coisa, não sugerimos upgrade
        return (0.0, 0.0, "")

# ----- Exportações -----

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
    """Exporta os processos com diferenca != 0, adicionando:
       - Marca, Modelo, Categoria de Veículo, KMS a Faturar no Serviço,
         Valor a Faturar S/IVA (vindas do ficheiro de referência)
       - Coluna Agravamento
       - Colunas Valor Potencial, Diferença Upgrade, Sugestão Upgrade
    """
    divergencias = df[df['Diferenca'] != 0].copy()

    # Determinar Agravamento
    def verificar_agravamento(data):
        if pd.isna(data):
            return "Não"
        if data.weekday() >= 5 or data.hour < 7 or data.hour >= 20:
            return "Sim"
        return "Não"

    divergencias['Data_Requisicao'] = pd.to_datetime(divergencias['Data_Requisicao'], errors='coerce')
    divergencias['Agravamento'] = divergencias['Data_Requisicao'].apply(verificar_agravamento)

    # Merge com o ficheiro de referência
    if referencia is not None:
        # Ajustar se o teu ficheiro de referência tiver nomes diferentes
        # (Aqui estamos a assumir que existem estas colunas exatas)
        colunas_ref = [
            'Matrícula',
            'Marca',
            'Modelo',
            'Categoria de Veículo',
            'KMS a Faturar no Serviço',
            'Valor a Faturar S/IVA'
        ]
        divergencias = divergencias.merge(
            referencia[colunas_ref],
            how='left',
            left_on='Matricula',   # no DF base: 'Matricula'
            right_on='Matrícula'   # no ref: 'Matrícula'
        )

        # Calcular Valor Potencial / Diferença / Sugestão
        divergencias[['Valor Potencial', 'Diferença Upgrade', 'Sugestão Upgrade']] = divergencias.apply(
            lambda row: pd.Series(calcular_upgrade(row)),
            axis=1
        )

    # Construir Excel
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    divergencias.to_excel(writer, index=False, sheet_name='Divergencias')

    workbook = writer.book
    worksheet = writer.sheets['Divergencias']

    format_agravado = workbook.add_format({'bg_color': '#FFFF00'})  # Amarelo
    format_faltante = workbook.add_format({'bg_color': '#FFFF00'})  # Amarelo

    headers = list(divergencias.columns)

    # Destacar Agravamento
    if "Agravamento" in headers:
        col_agravamento_idx = headers.index("Agravamento")
        for row_num, valor_agrav in enumerate(divergencias['Agravamento'], start=1):
            if valor_agrav == "Sim":
                worksheet.write(row_num, col_agravamento_idx, valor_agrav, format_agravado)

    # Destacar células em falta
    for row_num, row_data in divergencias.iterrows():
        if 'Marca' in headers and pd.isna(row_data.get('Marca')):
            worksheet.write(row_num + 1, headers.index('Marca'), '', format_faltante)
        if 'Modelo' in headers and pd.isna(row_data.get('Modelo')):
            worksheet.write(row_num + 1, headers.index('Modelo'), '', format_faltante)
        if 'Categoria de Veículo' in headers and pd.isna(row_data.get('Categoria de Veículo')):
            worksheet.write(row_num + 1, headers.index('Categoria de Veículo'), '', format_faltante)
        if 'KMS a Faturar no Serviço' in headers and pd.isna(row_data.get('KMS a Faturar no Serviço')):
            worksheet.write(row_num + 1, headers.index('KMS a Faturar no Serviço'), '', format_faltante)
        if 'Valor a Faturar S/IVA' in headers and pd.isna(row_data.get('Valor a Faturar S/IVA')):
            worksheet.write(row_num + 1, headers.index('Valor a Faturar S/IVA'), '', format_faltante)

    writer.close()
    output.seek(0)
    data_atual = datetime.now().strftime("%Y-%m-%d")
    return output, f"Listagem_Divergencias_{data_atual}.xlsx"


# ----- Início da aplicação -----
st.title("Gestão de Faturação - 🚛")

if 'login' not in st.session_state:
    st.session_state['login'] = False

if not st.session_state['login']:
    st.subheader("Login de Acesso")
    verificar_login()
else:
    # Upload do ficheiro de comparação
    st.subheader("Upload do Ficheiro de Comparação")
    uploaded_file = st.file_uploader("Escolhe o ficheiro Excel de comparação", type=["xlsx"])

    # Upload do ficheiro de referência
    st.subheader("Upload do Ficheiro de Referência (com colunas Matrícula + Marca/Modelo/Categoria + KMS + Valor a Faturar S/IVA)")
    referencia_file = st.file_uploader("Escolhe o ficheiro de referência", type=["xlsx"])

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

            # Só exportamos divergências se já tivermos um df de referência
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
