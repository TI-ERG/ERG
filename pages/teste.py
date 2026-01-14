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


up_viagens = st.file_uploader("Arquivo Relatório Controle Operacional Detalhado por Linha.csv", type='csv', key=2)

botao = st.button("Iniciar", type="primary")
if botao:
    try:
        if up_viagens is None:
            st.warning("Arquivo Relatório Controle Operacional Detalhado por Linha!", icon=":material/error_outline:")        
            st.stop()

        df = files_utils.ler_detalhado_linha(up_viagens)
        
        # Dropa colunas desnecessárias
        #columns_to_drop = ['Parado', 'Prev', 'Real2', 'Dif2', 'CVg', 'Docmto', 'Motorista', 'Cobrador','Km_h', 'Meta', 'CVg2', 'TipoViagem']
        #df = df.drop(columns=columns_to_drop)
        
        st.write(df)


    except Exception as e:    
        st.error(f"🐞 Erro: {traceback.format_exc()}")