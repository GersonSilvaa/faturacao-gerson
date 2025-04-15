# utils.py
import streamlit as st
import pandas as pd

def verificar_login():
    user = st.text_input("Utilizador")
    password = st.text_input("Password", type="password")
    if st.button("Entrar"):
        if user in st.secrets["utilizadores"] and st.secrets["utilizadores"][user] == password:
            st.session_state['login'] = True
            st.rerun()
        else:
            st.error("Credenciais inválidas!")

def processar_ficheiro(uploaded_file, colunas_obrigatorias=None):
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df.columns = df.columns.str.strip()
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

def calcular_valor_categoria(categoria, kms, agravamento, tabela='IPA'):
    if pd.isna(kms):
        return 0.0

    valor = 0.0

    if tabela == 'IPA':
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
    elif tabela == 'FIDELIDADE':
        # Aqui poderás definir regras específicas da Fidelidade no futuro
        pass

    if agravamento == 'Sim':
        valor *= 1.25

    return valor

def calcular_upgrade(row):
    cat_atual = row.get('Categoria de Veículo')
    kms = row.get('KMS a Faturar no Serviço')
    agravamento = row.get('Agravamento', 'Não')
    valor_base = row.get('Valor a Faturar S/IVA', 0.0)

    if pd.isna(valor_base):
        valor_base = 0.0

    if cat_atual == 'Ligeiro':
        valor_pot = calcular_valor_categoria('Furgão', kms, agravamento)
        dif = valor_pot - valor_base
        if dif > 0:
            return (valor_pot, dif, f"Analisar upgrade p/ Furgão (+{dif:.2f}€)")
        else:
            return (valor_pot, dif, "")
    elif cat_atual == 'Furgão':
        valor_pot = calcular_valor_categoria('Rodado Duplo', kms, agravamento)
        dif = valor_pot - valor_base
        if dif > 0:
            return (valor_pot, dif, f"Analisar upgrade p/ Rodado Duplo (+{dif:.2f}€)")
        else:
            return (valor_pot, dif, "")
    else:
        return (0.0, 0.0, "")

# Estas colunas adicionais da Fidelidade poderão ser analisadas por novas funções se necessário
FIDELIDADE_COLUNAS_EXTRA = [
    "Descrição da Avaria no Local",
    "Serviço de Desempanagem no Local",
    "Contacto Técnico",
    "Observações"
]

def obter_info_extra(row):
    return {
        "Avaria": row.get("Descrição da Avaria no Local", ""),
        "Desempanagem": row.get("Serviço de Desempanagem no Local", ""),
        "Contacto": row.get("Contacto Técnico", ""),
        "Observacoes": row.get("Observações", "")
    }
