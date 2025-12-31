import streamlit as st
import json
import os

ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(ROOT_DIR, "linhas.json")

def carregar():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def salvar(dados):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=4, ensure_ascii=False)

st.subheader("Edição de tabela - Linhas")
st.markdown("Indicadores AGERGS")
st.divider()

dados = carregar()

editado = st.data_editor(dados, num_rows="dynamic")

if st.button("💾 Salvar"):
    salvar(editado)
    st.success("Configurações salvas!")

