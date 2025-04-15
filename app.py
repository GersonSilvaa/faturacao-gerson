import streamlit as st
from utils import verificar_login
from ipa import run_ipa
from fidelidade import run_fidelidade  # mesmo que esteja vazio por enquanto

# ----- Login e menu inicial -----
if 'login' not in st.session_state:
    st.session_state['login'] = False

if not st.session_state['login']:
    st.subheader("Login de Acesso")
    verificar_login()
else:
    st.title("Gestão de Faturação 💼")
    st.subheader("Escolhe a companhia com que vais trabalhar")
    companhia = st.selectbox("Companhia:", ["IPA", "FIDELIDADE"])

    if companhia == "IPA":
        run_ipa()
    elif companhia == "FIDELIDADE":
        run_fidelidade()
