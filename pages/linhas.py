import streamlit as st
from utils import json_utils

# Lê o arquivo config e define o arquivo de linhas
config = json_utils.ler_json("config.json")
arq_linhas = config["matrizes"]["linhas"]
arq_linhas_raiz = config["matrizes"]["linhas_raiz"]

def carregar_linhas():
    return json_utils.ler_json(arq_linhas)

def carregar_linhas_raiz():
    return json_utils.ler_json(arq_linhas_raiz)

st.header("🚏 LINHAS", anchor=False)
st.markdown("**EDIÇÃO DE MATRIZ**")
st.divider()

# CSS para limitar a largura do data_editor 
st.markdown(""" <style> div[data-testid="stFullScreenFrame"] { max-width: 900px; } </style> """, unsafe_allow_html=True)

# Mostrar mensagem se acabou de salvar 
if st.session_state.get("salvo"): 
    st.success("Configurações salvas!") 
    st.session_state["salvo"] = False


# Abas
tab1, tab2 = st.tabs(["Linhas", "Linhas Raiz"])

with tab1:
    linhas = st.data_editor(
        carregar_linhas(), 
        column_config={ 
            "Cod_Bil": st.column_config.TextColumn(width="small"), 
            "Cod_Met": st.column_config.TextColumn(width="small"), 
            "Nome_Met": st.column_config.TextColumn(width="large"), 
            "Cod_Raiz": st.column_config.TextColumn(width="small"), 
            "Modal": st.column_config.TextColumn(width="small"), 
            },
        num_rows="dynamic"
    )
    linhas.sort(key=lambda x: x["Cod_Met"])

    if st.button("💾 Salvar", key="linhas"):
        json_utils.salvar_json(linhas, arq_linhas)
        st.session_state["salvo"] = True 
        st.rerun()

with tab2:
    linhas_raiz = st.data_editor(
        carregar_linhas_raiz(), 
        column_config={ 
            "Cod_Raiz": st.column_config.TextColumn(width="small"), 
            "Nome_Raiz": st.column_config.TextColumn(width="large"), 
            "Modal": st.column_config.TextColumn(width="small"), 
            },
        num_rows="dynamic"
    )
    linhas_raiz.sort(key=lambda x: x["Cod_Raiz"])

    if st.button("💾 Salvar", key="raiz"):
        json_utils.salvar_json(linhas_raiz, arq_linhas_raiz)
        st.session_state["salvo"] = True 
        st.rerun()
