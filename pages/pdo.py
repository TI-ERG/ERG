import traceback
from io import BytesIO
from datetime import date
from calendar import monthrange
import json
import streamlit as st
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from utils import json_utils
from utils import files_utils

# Configuração da página
st.set_page_config(layout="wide")

# Layout
with st.container():
    col1, col2, col3, col4 = st.columns([2, 2, 2, 1], vertical_alignment='top')

    with col1:
        # Upload do arquivo de dados de passageiros
        st.subheader("Dados de passageiros", help="Transnet > Módulos > Tráfego/Arrecadação > Consultas/Relatórios > Controle Operacional/Tráfego > Desempenho Diário das Linhas", anchor=False)
        up_passageiros = st.file_uploader("Arquivo Relatório Desempenho Diário das Linhas.csv", type='csv', key=1)
        
    with col2:
        # Upload do arquivo dos dados das viagens
        st.subheader("Dados de viagens", help="Transnet > Módulos > Tráfego/Arrecadação > Consultas/Relatórios > Controle Operacional/Tráfego > Controle Operacional Detalhado Por Linha", anchor=False)
        up_viagens = st.file_uploader("Arquivo Relatório Controle Operacional Detalhado por Linha.csv", type='csv', key=2)

    with col3:
        # Upload da planilha para conferência das viagens
        st.subheader("Planilha para conferência", help="Planilha enviada pelo Paulo", anchor=False)
        up_conferencia = st.file_uploader("Selecione um arquivo .XLSX", type='xlsx', key=3)

with st.container():       
    col1, col2, col3 = st.columns([3, 3, 3], vertical_alignment='top')
    with col1:
        # Feriados
        st.subheader("Feriados", anchor=False)

        df_feriado = pd.DataFrame([{"data": None, "escala": None}])

        # Editor de tabela
        df_feriado_editado = st.data_editor(
            df_feriado,
            num_rows="dynamic",
            column_config={
                "data": st.column_config.DateColumn("Data do feriado", format="DD/MM/YYYY"),
                "escala": st.column_config.SelectboxColumn("Escala", options=["Sábado", "Domingo"])
            }
        )

        # Converte depois do editor
        df_feriado_editado["data"] = pd.to_datetime(df_feriado_editado["data"], errors="coerce").dt.date


botao = st.sidebar.button("Iniciar", type="primary")

st.divider()

if botao:
    try:
        # Remove os botões
        st.session_state.pop("mostrar_downloads_pdo", None)

        # Verificações de seleção dos arquivos
        if up_passageiros is None:
            st.warning("Arquivo Relatório Desempenho Diário das Linhas não foi selecionado!", icon=":material/error_outline:")
            st.stop()

        if up_viagens is None:
            st.warning("Arquivo Relatório Controle Operacional Detalhado por Linha!", icon=":material/error_outline:")        
            st.stop()

        if up_conferencia is None:
            st.warning("Planilha para conferência não foi selecionada!", icon=":material/error_outline:")
            st.stop()

        with st.status("Processando...", expanded=False) as status:
            st.write("Lendo arquivos...")
            # Lê arquivo de configuração
            config = json_utils.ler_json("config.json")
            # Lê matriz de linhas
            df_linhas = pd.DataFrame(json_utils.ler_json(config["matrizes"]["linhas"]))
            # Lê arquivo detalhado por linha
            df_det = files_utils.ler_detalhado_linha(up_viagens)
            # Lê arquivo desempenho diário das linhas

            st.write("Tratando os dados do controle operacional detalhado por linha...")
            # Dropa colunas desnecessárias
            columns_to_drop = ['#', 'Orig', 'Dest', 'Dif', 'Parado', 'Prev', 'Real2', 'Dif2', 'CVg', 'Veiculo', 'Docmto', 'Motorista', 'Cobrador', 'EmPe', 'Sent.1', 'Km_h', 'Meta', 'CVg2', 'TipoViagem']
            df_det = df_det.drop(columns=columns_to_drop)
            # Merge com arquivo de linhas para ter a modalidade/serviço
            df_det = df_det.merge(df_linhas[["Cod_Met", "Modal"]], left_on="Codigo", right_on="Cod_Met", how="left")
            df_det = df_det.drop(columns=["Cod_Met"])
            # Exclui viagens que NÃO TEM passageiros (possíveis erros de digitação)
            df_det_filtrado = df_det[~(df_det["Passag"].isna() & (df_det["Observacao"].str.strip() != "Furo de Viagem"))]

            st.write("Tratando os dados do desempenho diário das linhas...")


            status.update(label="Processo terminado!", state="complete", expanded=False)
            st.session_state["mostrar_downloads_pdo"] = True
            st.success("Arquivos gerados com sucesso!")

    except Exception as e:  
        status.update(label="Erro durante o processamento!", state="error")  
        st.error(f"🐞 Erro: {traceback.format_exc()}")

# ✳️ Downloads ✳️
if st.session_state.get("mostrar_downloads_pdo", False):       
    # Botões
    col1, col2, col3 = st.columns([1,1,5], vertical_alignment='top')
    with col1:
        st.download_button(
            label="📥 Baixar PDO-ERG", 
            data="conteúdo do arquivo", 
            file_name="relatorio.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
        )
    with col2:     
        st.download_button(
            label="📥 Baixar PDO-TM5", 
            data="conteúdo do arquivo", 
            file_name="relatorio.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
        )