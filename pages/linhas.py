import streamlit as st
import json
from utils import json_utils

# Lê o arquivo config e define o arquivo de linhas
config = json_utils.ler_json("config.json")
arq_linhas = config["matrizes"]["linhas"]

def carregar():
    return json_utils.ler_json(arq_linhas)

st.header("🚏 LINHAS", anchor=False)
st.markdown("**EDIÇÃO DE MATRIZ**")
st.divider()

# CSS para limitar a largura do data_editor 
st.markdown(""" <style> div[data-testid="stFullScreenFrame"] { max-width: 900px; } </style> """, unsafe_allow_html=True)

# Mostrar mensagem se acabou de salvar 
if st.session_state.get("salvo"): 
    st.success("Configurações salvas!") 
    st.session_state["salvo"] = False

editado = st.data_editor(
    carregar(), 
    column_config={ 
        "Cod_Bil": st.column_config.TextColumn(width="small"), 
        "Cod_Met": st.column_config.TextColumn(width="small"), 
        "Linha": st.column_config.TextColumn(width="large"), 
        "Raiz": st.column_config.TextColumn(width="small"), 
        "Modal": st.column_config.TextColumn(width="small"), 
        },
    num_rows="dynamic"
)
editado.sort(key=lambda x: x["Cod_Met"])

if st.button("💾 Salvar"):
    json_utils.salvar_json(editado, arq_linhas)
    st.session_state["salvo"] = True 
    st.rerun()
