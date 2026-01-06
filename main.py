import streamlit as st
import subprocess

# Configuração da página
#st.set_page_config(layout="wide")

def get_git_version():
    try:
        version = subprocess.check_output(["git", "describe", "--tags"]).decode().strip()
        return version
    except:
        return "versão desconhecida"
    
def pagina_inicial():
    st.title("ERG Tools 🪛")
    st.subheader("-- Sistema interno de funções --")
    st.write("Você pode navegar pelas seções ao lado.")


versao = get_git_version()
st.sidebar.write(f"Versão do sistema: {versao}")

pages = {
    "Exportação de Arquivos": [
        st.Page("pages/bod.py", title="Boletim Oferta e Demanada", icon="📄"),
        st.Page("pages/pdo.py", title="Dados Operacionais", icon="📄"),
        st.Page("pages/agergs.py", title="Indicadores AGERGS", icon="📄")
    ]
}

pages_dados = {
    "Dados AGERGS": [
        st.Page("pages/frota_agergs.py", title="Frota", icon="🚌"),
        st.Page("pages/linhas_agergs.py", title="Linhas", icon="🚏")
    ]
}

navegacao = { 
    "": [st.Page(pagina_inicial, title="ERG Tools", icon="🪛")], 
    **pages, 
    **pages_dados }


pg = st.navigation(navegacao)
pg.run()

if st.session_state.get("page") == "frota": 
    st.switch_page("pages/frota.py")
elif st.session_state.get("page") == "linhas":
    st.switch_page("pages/linhas.py")




