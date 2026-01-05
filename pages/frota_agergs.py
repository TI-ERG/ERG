import streamlit as st
import json
import os

# Caminho arquivos
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT_DIR, "config.json")

# Lê o arquivo config e define o arquivo de linhas
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)
    FROTA_PATH = os.path.join(ROOT_DIR, config["agergs"]["frota"])

def carregar():
    with open(FROTA_PATH, "r", encoding="utf-8") as frota:
        return json.load(frota)

def salvar(df):
    with open(FROTA_PATH, "w", encoding="utf-8") as frota:
        json.dump(df, frota, indent=4, ensure_ascii=False)

st.subheader("Edição de tabela - Frota")
st.markdown("Indicadores AGERGS")
st.divider()

editado = st.data_editor(carregar(), num_rows="dynamic")

if st.button("💾 Salvar"):
    salvar(editado)
    st.success("Configurações salvas!")

