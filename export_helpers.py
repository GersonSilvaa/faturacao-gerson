# export_helpers.py
import pandas as pd
import io
from datetime import datetime
from utils import PALAVRAS_SUSPEITAS, FIDELIDADE_COLUNAS_EXTRA, contem_texto_suspeito, calcular_upgrade

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

    if referencia is not None:
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
            left_on='Matricula',
            right_on='Matrícula'
        )

        divergencias[['Valor Potencial', 'Diferença Upgrade', 'Sugestão Upgrade']] = divergencias.apply(
            lambda row: pd.Series(calcular_upgrade(row)),
            axis=1
        )

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    divergencias.to_excel(writer, index=False, sheet_name='Divergencias')

    workbook = writer.book
    worksheet = writer.sheets['Divergencias']

    format_agravado = workbook.add_format({'bg_color': '#FFFF00'})
    format_faltante = workbook.add_format({'bg_color': '#FFFF00'})

    headers = list(divergencias.columns)

    if "Agravamento" in headers:
        col_agravamento_idx = headers.index("Agravamento")
        for row_num, valor_agrav in enumerate(divergencias['Agravamento'], start=1):
            if valor_agrav == "Sim":
                worksheet.write(row_num, col_agravamento_idx, valor_agrav, format_agravado)

    for row_num, row_data in divergencias.iterrows():
        for col in ['Marca', 'Modelo', 'Categoria de Veículo', 'KMS a Faturar no Serviço', 'Valor a Faturar S/IVA']:
            if col in headers and pd.isna(row_data.get(col)):
                worksheet.write(row_num + 1, headers.index(col), '', format_faltante)

    writer.close()
    output.seek(0)
    data_atual = datetime.now().strftime("%Y-%m-%d")
    return output, f"Listagem_Divergencias_{data_atual}.xlsx"


def exportar_fidelidade_excel(df):
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Fidelidade')

    workbook = writer.book
    worksheet = writer.sheets['Fidelidade']

    amarelo = workbook.add_format({'bg_color': '#FFFF00'})
    colunas = list(df.columns)

    for row_idx, row in df.iterrows():
        marcar = False
        for col_name in FIDELIDADE_COLUNAS_EXTRA:
            if col_name in colunas:
                valor = str(row.get(col_name, "")).lower()
                if any(p in valor for p in PALAVRAS_SUSPEITAS):
                    col_idx = colunas.index(col_name)
                    worksheet.write(row_idx + 1, col_idx, row[col_name], amarelo)
                    marcar = True

        if marcar and 'Matrícula' in colunas:
            col_idx = colunas.index('Matrícula')
            worksheet.write(row_idx + 1, col_idx, row['Matrícula'], amarelo)

    writer.close()
    output.seek(0)
    return output, "Analise_Fidelidade_Formatado.xlsx"

# IPA
def exportar_cruzamento_weboffice(weboffice_df, referencia_df):
    weboffice_df.columns = weboffice_df.columns.str.strip()
    referencia_df.columns = referencia_df.columns.str.strip()

    # Normalizar Matrículas
    weboffice_df["matricula_normalizada"] = weboffice_df["Apol/Mat"].astype(str).str.replace("-", "").str.upper().str.strip()
    referencia_df["matricula_normalizada"] = referencia_df["Matrícula"].astype(str).str.replace("-", "").str.upper().str.strip()

    # Agrupar no Gestow por matrícula e somar valores
    gestow_agrupado = (
        referencia_df
        .groupby("matricula_normalizada", as_index=False)
        .agg({"Valor a Faturar S/IVA": "sum"})
    )
    gestow_agrupado["Total_Gestow"] = gestow_agrupado["Valor a Faturar S/IVA"] * 1.23

    # Merge com WebOffice
    merged = weboffice_df.merge(
        gestow_agrupado[["matricula_normalizada", "Total_Gestow"]],
        on="matricula_normalizada",
        how="left"
    )

    # Limpeza de valores e diferença
    merged["Total"] = (
        merged["Total"]
        .astype(str)
        .str.replace("€", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )
    merged["Total_IPA"] = pd.to_numeric(merged["Total"], errors="coerce")
    merged["Total_Gestow"] = pd.to_numeric(merged["Total_Gestow"], errors="coerce")
    merged["Diferença €"] = merged["Total_IPA"] - merged["Total_Gestow"]

    # Comentário (Obs.)
    def classificar_diferenca(diff):
        if pd.isna(diff):
            return "Sem correspondência"
        elif abs(diff) < 0.01:
            return "Igual"
        elif diff > 0:
            return "A ganhar"
        else:
            return ""

    merged["Obs."] = merged["Diferença €"].apply(classificar_diferenca)

    # Construir dataframe final
    final = merged[[
        "Apol/Mat",
        "Dossier",
        "Total_IPA",
        "Diferença €",
        "Total_Gestow",
        "Obs."
    ]].rename(columns={
        "Apol/Mat": "Matricula",
        "Dossier": "Dossier_IPA"
    }).sort_values(by="Dossier_IPA")

    # Exportar para Excel com cores
    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    final.to_excel(writer, index=False, sheet_name="Cruzamento WebOffice")

    workbook = writer.book
    worksheet = writer.sheets["Cruzamento WebOffice"]

    # Estilos para cabeçalhos
    header_format_ipa = workbook.add_format({
        'bold': True,
        'bg_color': '#DDEBF7',
        'border': 1
    })

    header_format_gestow = workbook.add_format({
        'bold': True,
        'bg_color': '#E2F0D9',
        'border': 1
    })

    for col_num, column in enumerate(final.columns):
        if column in ["Matricula", "Dossier_IPA", "Total_IPA"]:
            worksheet.write(0, col_num, column, header_format_ipa)
        elif column in ["Total_Gestow", "Obs.", "Diferença €"]:
            worksheet.write(0, col_num, column, header_format_gestow)

    writer.close()
    output.seek(0)
    return output, "Cruzamento_WebOffice_vs_Gestow.xlsx"

# ACP
def exportar_acp_corrigido(acp_df, gestow_df):
    acp_df = acp_df.copy()
    gestow_df = gestow_df.copy()

    # Normalizar Matrículas
    acp_df["matricula_normalizada"] = acp_df["Matrícula"].astype(str).str.replace("-", "").str.upper().str.strip()
    gestow_df["matricula_normalizada"] = gestow_df["Matrícula"].astype(str).str.replace("-", "").str.upper().str.strip()

    # Criar dicionário: matrícula → processo
    mapa_processos = gestow_df.set_index("matricula_normalizada")["Processo da Companhia"].to_dict()

    # Substituir a coluna "Interv." por valores correspondentes
    acp_df["Interv."] = acp_df["matricula_normalizada"].map(mapa_processos).fillna("NAO ENCONTRADO")

    # Remover coluna auxiliar
    acp_df.drop(columns=["matricula_normalizada"], inplace=True)

    # Corrigir número do processo para remover ".0"
    if "Interv." in acp_df.columns:
        acp_df["Interv."] = acp_df["Interv."].apply(lambda x: str(int(x)) if pd.notna(x) and str(x).replace('.', '', 1).isdigit() else x)

    # Exportar
    output = io.BytesIO()
    conteudo_csv = acp_df.to_csv(index=False, sep=";", encoding="utf-8-sig")
    output.write(conteudo_csv.encode("utf-8-sig"))
    output.seek(0)

    return output, "ACP_Corrigido.csv"
