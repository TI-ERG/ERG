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

# Lê arquivo de configuração
config = json_utils.ler_json("config.json")

# Layout
with st.container():
    col1, col2, col3, col4 = st.columns([2, 2, 2, 1], vertical_alignment='top')

    with col1:
        # Upload do arquivo de dados de passageiros
        st.subheader("Dados de passageiros")
        up_passageiros = st.file_uploader("Selecione um arquivo .CSV", type='csv', key=1)
        
    with col2:
        # Upload do arquivo dos dados das viagens
        st.subheader("Dados de viagens")
        up_viagens = st.file_uploader("Selecione um arquivo .CSV", type='csv', key=2)

    with col3:
        # Upload da planilha para conferência das viagens
        st.subheader("Planilha para conferência")
        up_conferencia = st.file_uploader("Selecione um arquivo .XLSX", type='xlsx', key=3)

with st.container():       
    col1, col2, col3 = st.columns([3, 3, 3], vertical_alignment='top')
    with col1:
        # Feriados
        feriados = {} 
        st.subheader("Informe os feriados") 

        datas_feriados = st.date_input( "Datas dos feriados:", value=[], format="DD/MM/YYYY" )

        if isinstance(datas_feriados, date): datas_feriados = [datas_feriados]

        if datas_feriados: 
            st.subheader("Defina a escala de cada feriado")

        for data in datas_feriados: 
            escala = st.selectbox( f"Escala do feriado em {data.strftime('%d/%m/%Y')}:", ["Sábado", "Domingo"], key=f"escala_{data}" ) 
            feriados[data] = escala


        if feriados: 
            for data, escala in feriados.items(): 
                st.write(f"📅 {data.strftime('%d/%m/%Y')} → {escala}")

botao = st.sidebar.button("Iniciar", type="primary")

st.divider()

if botao:
    try:
        # Verificações de seleção dos arquivos
        if up_passageiros is None:
            st.warning("Arquivo de dados dos passageiros não foi selecionado!", icon=":material/error_outline:")
            st.stop()

        if up_viagens is None:
            st.warning("Arquivo de dados das viagens não foi selecionado!", icon=":material/error_outline:")        
            st.stop()

        if up_conferencia is None:
            st.warning("Planilha para conferência não foi selecionada!", icon=":material/error_outline:")
            st.stop()

        df = files_utils.ler_detalhado_linha(up_viagens)
        
        # Dropa colunas desnecessárias
        columns_to_drop = ['Parado', 'Prev', 'Real2', 'Dif2', 'CVg', 'Docmto', 'Motorista', 'Cobrador','Km_h', 'Meta', 'CVg2', 'TipoViagem']
        df = df.drop(columns=columns_to_drop)
        
        st.write(df)


    except Exception as e:    
        st.error(f"🐞 Erro: {traceback.format_exc()}")
       
