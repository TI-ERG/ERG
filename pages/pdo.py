import traceback
from io import BytesIO
from datetime import date
from calendar import monthrange
import json
import streamlit as st
import pandas as pd
import numpy as np
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from utils import json_utils
from utils import files_utils
from utils import date_utils

def criar_abas_por_semana(wb, data):
    if "Modelo" not in wb.sheetnames:
        raise ValueError("A aba 'Modelo' não existe no workbook.")

    aba_modelo = wb["Modelo"]
    # Insiro informações padrões 
    aba_modelo["A2"] = "Nome da Empresa: Expresso Rio Guaíba"
    aba_modelo["D2"] = "Códgo da Empresa: GU99"
    aba_modelo["G2"] = f"Mês de referência: {pd.to_datetime(data).month_name(locale="pt_BR")}/{data.year}"

    total_semanas = date_utils.semanas_no_mes(data)

    for i in range(1, total_semanas + 1):
        nome_aba = date_utils.semana_extenso_numero(i)

        if nome_aba in wb.sheetnames:
            del wb[nome_aba]

        nova_aba = wb.copy_worksheet(aba_modelo)
        nova_aba.title = nome_aba

    del wb["Modelo"]

    return wb



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
                "data": st.column_config.DateColumn("Data", format="DD/MM/YYYY"),
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
        st.session_state.pop("pdo", None)

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

        with st.status("Lendo arquivos...", expanded=False) as status:
            st.write("📄 Processando os arquivos...")
            config = json_utils.ler_json("config.json") # Arquivo de configuração
            df_linhas = pd.DataFrame(json_utils.ler_json(config["matrizes"]["linhas"])) # Matriz de linhas
            df_det1 = files_utils.ler_detalhado_linha(up_viagens) # Arquivo detalhado por linha
            # Arquivo desempenho diário das linhas
            #‼️‼️‼️‼️‼️‼️‼️‼️


            status.update(label="Analisando dados...", state="running", expanded=False)
            st.write("🧠 Processando - Controle Operacional Detalhado por Linha...")
            # Dropa colunas desnecessárias
            columns_to_drop = ['#', 'Orig', 'Dest', 'Dif', 'Parado', 'Prev', 'Real2', 'Dif2', 'CVg', 'Veiculo', 'Docmto', 'Motorista', 'Cobrador', 'EmPe', 'Sent.1', 'Km_h', 'Meta', 'CVg2', 'TipoViagem']
            df_det1 = df_det1.drop(columns=columns_to_drop)
            # Merge com arquivo de linhas para ter a modalidade/serviço
            df_det1 = df_det1.merge(df_linhas[["Cod_Met", "Modal"]], left_on="Codigo", right_on="Cod_Met", how="left")
            df_det1 = df_det1.drop(columns=["Cod_Met"])
            df_det1["Observacao"] = df_det1["Observacao"].astype(str).str.strip()
            # Exclui viagens que NÃO TEM passageiros (possíveis erros de digitação)
            df_det2 = df_det1[~(df_det1["Passag"].isna() & (df_det1["Observacao"] != "Furo de Viagem"))]
            # Preenche o horário das previstas com as realizadas nas viagens extras (vou usar esta coluna para os horários)
            df_det2.loc[:, "THor"] = df_det2["THor"].fillna(df_det2["Real"])
            
            df_det2["Dia"] = pd.to_datetime(df_det2["Dia"], dayfirst=True, errors="coerce").dt.date
            df_det2 = df_det2.sort_values(["Sent", "Codigo", "Dia", "THor"])
            
            st.write("🧠 Processando - Desempenho Diário das Linhas...")
            #‼️‼️‼️‼️‼️‼️‼️‼️

            st.write("🧠 Processando os dados para conferência...")
            df_conf1 = df_det2[["Codigo", "Sent", "Dia", "Observacao"]]
            df_conf1["Dia Semana"] = df_conf1["Dia"].map(lambda d: {0:"U",1:"U",2:"U",3:"U",4:"U",5:"S",6:"D"}[d.weekday()]) # Crio Dia Semana
            # Atualizo com o feriado
            # 1. Faz o merge com base na data
            df_conf_merged = df_conf1.merge(df_feriado_editado, left_on='Dia', right_on='data', how='left')
            # 2. Atualiza 'Dia Semana' com base na escala
            df_conf_merged.loc[df_conf_merged['escala'] == 'Sábado', 'Dia Semana'] = 'S'
            df_conf_merged.loc[df_conf_merged['escala'] == 'Domingo', 'Dia Semana'] = 'D'
            # 3. Remove colunas extras
            df_conf1 = df_conf_merged
            df_conf1 = df_conf1.drop(columns=["Dia", "data", "escala"])
            # 4. Agrupa
            df_conf2 = (
                df_conf1
                    .assign(
                        Tipo=lambda x: x['Observacao'].map({
                            'OK': 'ERG',
                            'Viagem Extra': 'EXT',
                            'Furo de Viagem': 'FURO'
                        }),
                        Chave=lambda x: x['Tipo'] + '_' + x['Dia Semana'] + x['Sent'].astype(str)
                    )
                    .groupby(['Codigo', 'Chave'])
                    .size()
                    .unstack(fill_value=0)
                    .reindex(columns=['ERG_U1', 'EXT_U1', 'FURO_U1', 'ERG_U2', 'EXT_U2', 'FURO_U2',
                                      'ERG_S1', 'EXT_S1', 'FURO_S1', 'ERG_S2', 'EXT_S2', 'FURO_S2',
                                      'ERG_D1', 'EXT_D1', 'FURO_D1', 'ERG_D2', 'EXT_D2', 'FURO_D2'],
                                        fill_value=0)
                    .reset_index()
            )

            # Lendo Planilha modelo ModeloPDO.xlsx
            status.update(label="Preenchendo a planilha modelo...", state="running", expanded=False)
            wb = load_workbook(config['pdo']['modelo_pdo'])

            # 1️⃣ CONFERÊNCIA
            st.write("✒️ Editando a planilha - Aba de conferência...")
            ws_conf = wb["Conferência"]
            linha_excel = 4
            for row in df_conf2.itertuples(index=False): 
                ws_conf[f"A{linha_excel}"] = row.Codigo 
                ws_conf[f"B{linha_excel}"] = row.ERG_U1
                ws_conf[f"E{linha_excel}"] = row.EXT_U1
                ws_conf[f"F{linha_excel}"] = row.FURO_U1
                ws_conf[f"H{linha_excel}"] = row.ERG_U2
                ws_conf[f"K{linha_excel}"] = row.EXT_U2
                ws_conf[f"L{linha_excel}"] = row.FURO_U2
                ws_conf[f"N{linha_excel}"] = row.ERG_S1
                ws_conf[f"Q{linha_excel}"] = row.EXT_S1
                ws_conf[f"R{linha_excel}"] = row.FURO_S1
                ws_conf[f"T{linha_excel}"] = row.ERG_S2
                ws_conf[f"W{linha_excel}"] = row.EXT_S2
                ws_conf[f"X{linha_excel}"] = row.FURO_S2
                ws_conf[f"Z{linha_excel}"] = row.ERG_D1
                ws_conf[f"AC{linha_excel}"] = row.EXT_D1
                ws_conf[f"AD{linha_excel}"] = row.FURO_D1
                ws_conf[f"AF{linha_excel}"] = row.ERG_D2
                ws_conf[f"AI{linha_excel}"] = row.EXT_D2
                ws_conf[f"AJ{linha_excel}"] = row.FURO_D2
                linha_excel += 1

                        
            # 2️⃣ SEMANAS
            st.write("✒️ Editando a planilha - Abas semanais...")
            # Crio as abas das semanas
            wb = criar_abas_por_semana(wb, df_det2.loc[0, "Dia"])
            # Informo os dias e estilizo os dias de feriados
            posicoes_dias = {
                0: ("E5", "E9"),    # segunda
                1: ("K5", "K9"),    # terça
                2: ("Q5", "Q9"),    # quarta
                3: ("W5", "W9"),    # quinta
                4: ("AC5", "AC9"),  # sexta
                5: ("AI5", "AI9"),  # sábado
                6: ("AO5", "AO9"),  # domingo
            }
            
            fill_feriado = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

            df_dias = df_det2[["Dia"]].drop_duplicates().sort_values("Dia") # Dias únicos no mês
            feriados_dict = {
                row["data"]: row["escala"]
                for _, row in df_feriado_editado.iterrows()
            }
            abas_semanas = wb.sheetnames[2:]  # ignora as duas primeiras abas

            for _, row in df_dias.iterrows():
                data = row["Dia"] 
                dia_semana = data.weekday() # 0=segunda ... 6=domingo
                semana = date_utils.semana_do_mes(data) # 1 a 6
                texto = f"{date_utils.dia_da_semana(data)} Dia: {data.day}"
                cel1, cel2 = posicoes_dias[dia_semana]

                ws_temp = wb[abas_semanas[semana - 1]]
                # Estiliza feriado
                if data in feriados_dict:
                    texto = f"{texto} (Escala de {feriados_dict[data]})"
                    ws_temp[cel1].fill = fill_feriado
                    ws_temp[cel2].fill = fill_feriado

                ws_temp[cel1] = texto
                ws_temp[cel2] = texto


            #‼️‼️‼️‼️‼️‼️‼️‼️


            # 3️⃣ TOTAL GERAL
            st.write("✒️ Editando a planilha - Aba de totais...")

            # Salvar em memória
            status.update(label="Salvando...", state="running", expanded=False)
            st.write("💾 Salvando a planilha...")
            buffer_pdo = BytesIO()
            wb.save(buffer_pdo)
            buffer_pdo.seek(0)
            st.session_state["buffer_pdo"] = buffer_pdo # Arquivo
            st.session_state["pdo"] = f"{df_det2.loc[0, "Dia"].strftime("%m.%Y")}" # Condição para os botões

            status.update(label="Processo terminado!", state="complete", expanded=False)
            st.success("Arquivos gerados com sucesso!")

    except Exception as e:  
        status.update(label="Erro durante o processamento!", state="error")  
        st.error(f"🐞 Erro: {traceback.format_exc()}")
    finally:
        df_linhas = None
        df_feriado_editado = None
        df_feriado = None
        df_det1 = None
        df_det2 = None
        df_conf1 = None
        df_conf2 = None
        df_conf_merged = None
        buffer_pdo = None
        wb = None

# ✳️ Downloads ✳️
if st.session_state.get("pdo", False):  
    col1, col2, col3 = st.columns([1,1,5], vertical_alignment='top')
    with col1:
        st.download_button(
            label="📥 Baixar PDO-ERG", 
            data=st.session_state["buffer_pdo"], 
            file_name=f"GUAIBA [{st.session_state["pdo"]}].xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
        )
    with col2:     
        st.download_button(
            label="📥 Baixar PDO-TM5", 
            data="conteúdo do arquivo", 
            file_name=f"GUAIBA-TM5 [{st.session_state["pdo"]}].xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" 
        )