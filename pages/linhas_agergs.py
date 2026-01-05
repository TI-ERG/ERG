import streamlit as st
import json
import os

# Caminho arquivos
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT_DIR, "config.json")

# Lê o arquivo config e define o arquivo de linhas
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)
    LINHAS_PATH = os.path.join(ROOT_DIR, config["agergs"]["linhas"])

def carregar():
    with open(LINHAS_PATH, "r", encoding="utf-8") as linhas:
        return json.load(linhas)

def salvar(df):
    with open(LINHAS_PATH, "w", encoding="utf-8") as linhas:
        json.dump(df, linhas, indent=4, ensure_ascii=False)

st.subheader("Edição de tabela - Linhas")
st.markdown("Indicadores AGERGS")
st.divider()

editado = st.data_editor(carregar(), num_rows="dynamic")

if st.button("💾 Salvar"):
    salvar(editado)
    st.success("Configurações salvas!")

