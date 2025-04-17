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


def exportar_cruzamento_weboffice(weboffice_df, referencia_df):
    weboffice_df.columns = weboffice_df.columns.str.strip()
    referencia_df.columns = referencia_df.columns.str.strip()

    weboffice_df["matricula_normalizada"] = weboffice_df["Apol/Mat"].astype(str).str.replace("-", "").str.upper().str.strip()
    referencia_df["matricula_normalizada"] = referencia_df["Matrícula"].str.replace("-", "").str.upper().str.strip()

    referencia_df["Valor Gestow c/IVA"] = referencia_df["Valor a Faturar S/IVA"] * 1.23

    merged = weboffice_df.merge(
        referencia_df[["matricula_normalizada", "Valor Gestow c/IVA"]],
        on="matricula_normalizada",
        how="left"
    )

    merged["Total"] = (
    merged["Total"]
    .astype(str)
    .str.replace("€", "", regex=False)
    .str.replace(",", ".", regex=False)
    .str.strip()
    )

    merged["Total"] = pd.to_numeric(merged["Total"], errors="coerce")
    merged["Valor Gestow c/IVA"] = pd.to_numeric(merged["Valor Gestow c/IVA"], errors="coerce")
    merged["Diferença €"] = merged["Total"] - merged["Valor Gestow c/IVA"]

    def classificar_diferenca(diff):
        if pd.isna(diff):
            return "Sem correspondência"
        elif abs(diff) < 0.01:
            return "Igual"
        elif diff > 0:
            return "A ganhar"
        else:
            return "A perder"

    merged["Comentário"] = merged["Diferença €"].apply(classificar_diferenca)

    final = merged[[
        "Apol/Mat",
        "Dossier",
        "Total",
        "Diferença €",
        "Valor Gestow c/IVA",
        "Comentário"
    ]].sort_values(by="Dossier")

    output = io.BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    final.to_excel(writer, index=False, sheet_name="Cruzamento WebOffice")
    writer.close()
    output.seek(0)

    return output, "Cruzamento_WebOffice_vs_Gestow.xlsx"
