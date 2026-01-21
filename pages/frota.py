import streamlit as st
from utils import json_utils

# Lê o arquivo config e define o arquivo de linhas
config = json_utils.ler_json("config.json")
arq_frota = config["matrizes"]["frota"]

def carregar():
    return json_utils.ler_json(arq_frota)

st.header("🚌 FROTA", anchor=False)
st.markdown("**EDIÇÃO DE MATRIZ**")
st.divider()

# CSS para limitar a largura do data_editor 
st.markdown(""" <style> div[data-testid="stFullScreenFrame"] { max-width: 400px; } </style> """, unsafe_allow_html=True)

# Mostrar mensagem se acabou de salvar 
if st.session_state.get("salvo"): 
    st.success("Configurações salvas!") 
    st.session_state["salvo"] = False

editado = st.data_editor(carregar(), num_rows="dynamic")
editado.sort(key=lambda x: x["Prefixo"])

if st.button("💾 Salvar"):
    json_utils.salvar_json(editado, arq_frota)
    st.session_state["salvo"] = True 
    st.rerun()

