import streamlit as st
from PIL import Image
from utils import verificar_login
from ipa import run_ipa
from fidelidade import run_fidelidade
from acp import run_acp

# ----- Layout da página -----
st.set_page_config(
page_title="Gestão de Faturação",
layout="wide"
)

# ----- Login e Navegação -----
if 'login' not in st.session_state:
    st.session_state['login'] = False

if not st.session_state['login']:
    st.subheader("Login de Acesso")
    verificar_login()
else: # ----- Sidebar com logo e menu -----
    try:
        logo = Image.open("assets/logo.png")  # Altera o caminho se estiver noutra pasta
        st.sidebar.image(logo, use_container_width=True)
    except:
        st.sidebar.write("🧾 Logo não encontrado.")

st.sidebar.title("Menu")
st.sidebar.markdown(f"👤 Utilizador: **{st.session_state.get('utilizador', 'Desconhecido')}**")
companhia = st.sidebar.selectbox("Companhia:", ["IPA", "FIDELIDADE", "ACP"], key="companhia")

# ----- Título da app -----
st.title("Gestão de Faturação")

# ----- Chamar o módulo da companhia selecionada -----
if companhia == "IPA":
    run_ipa()
elif companhia == "FIDELIDADE":
    run_fidelidade()
elif companhia == "ACP":
    run_acp()
