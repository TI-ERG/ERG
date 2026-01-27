import streamlit as st
import subprocess

def get_git_version():
    try:
        version = subprocess.check_output(["git", "describe", "--tags"]).decode().strip()
        return version
    except:
        return "versão desconhecida"
    
def pagina_inicial():
    st.title("🪛 ERG Tools")
    st.header("[**Sistema Interno de Funções**]")
    st.write("Você pode navegar pelas seções ao lado.")

st.set_page_config(layout="wide")
st.logo("images/guaiba-logo.svg", size="small")
st.sidebar.write(f"Versão do sistema: {get_git_version()}")

pages = {
    "Exportação de Arquivos": [
        st.Page("pages/bod.py", title="[BOD] Boletim Oferta e Demanda", icon="📄"),
        st.Page("pages/pdo.py", title="[PDO] Dados Operacionais", icon="📄"),
        st.Page("pages/agergs.py", title="Indicadores AGERGS", icon="📄")
    ]
}

pages_dados = {
    "Matrizes de Dados": [
        st.Page("pages/frota.py", title="Frota", icon="🚌"),
        st.Page("pages/linhas.py", title="Linhas", icon="🚏"),
        st.Page("pages/teste.py", title="Teste", icon="🚏")
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




