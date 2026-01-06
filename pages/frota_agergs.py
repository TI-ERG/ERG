import streamlit as st
import json
from utils import json_utils

# Caminho arquivos
#ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
#CONFIG_PATH = os.path.join(ROOT_DIR, "config.json")

# Lê o arquivo config e define o arquivo de linhas
config = json_utils.ler_json("config.json")
arq_frota = config["agergs"]["frota"]

def carregar():
    return json_utils.ler_json(arq_frota)

st.subheader("Edição de tabela - Frota")
st.markdown("Indicadores AGERGS")
st.divider()

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

