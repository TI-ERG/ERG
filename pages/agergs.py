import traceback
from datetime import date
import calendar
from utils import json_utils
from utils import format_utils
from utils import files_utils
from utils import error_utils
import xml.etree.ElementTree as ET
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from copy import copy

# region FUNÇÕES
def atualizar_dados(df):
    def safe_div(a, b):
        return a / b if (b not in (0, None) and b != 0) else 0

    df["VG Interrompidas"] = df["Quebras"] + df["Acidentes"]
    df["Indice VG OF Realizadas"] = df.apply(
        lambda x: safe_div(x["VG Realizadas"], x["VG Previstas"]),
        axis=1
    )
    df["Indice Pontualidade"] = df.apply(
        lambda x: safe_div(x["VG Realizadas"] - x["Atrasos"], x["VG Realizadas"]),
        axis=1
    )
    df["Desemp. VG Interromp"] = df.apply(
        lambda x: safe_div(x["VG Realizadas"] - x["VG Interrompidas"], x["VG Realizadas"]),
        axis=1
    )
    df["Indice Quebra"] = df.apply(
        lambda x: safe_div(x["Quebras"], x["VG Realizadas"]) * 1000,
        axis=1
    )
    df["Indice Desv. Itinerário"] = df.apply(
        lambda x: safe_div(x["Desv. Itinerário"], x["VG Realizadas"]) * 1000,
        axis=1
    )
    df["Indice Acidentes"] = df.apply(
        lambda x: safe_div(x["Acidentes"], x["VG Realizadas"]) * 1000,
        axis=1
    )

    return df

def ler_detalhado():
    df = files_utils.ler_detalhado_linha(up_detalhado) # Lê o arquivo
    # Verifica se o período e o arquivo conferem
    periodo = pd.to_datetime(df.iloc[0]["Dia"], dayfirst=True, errors="coerce").date()
    if periodo.month != mes or periodo.year != ano:
        raise error_utils.ErroDePeriodo(periodo)
    
    st.write(f"{periodo.month} {mes}")

    # Dropa colunas desnecessárias
    columns_to_drop = ['Sent', 'Parado', 'Prev', 'Real2', 'Dif2', 'CVg', 'Docmto', 'Motorista', 'Cobrador','Km_h', 'Meta', 'CVg2', 'TipoViagem']
    df = df.drop(columns=columns_to_drop)

    # Convert 'Veiculo', 'Passag', 'Oferta' and 'Dif' to numeric, coercing errors to NaN
    df['Veiculo'] = pd.to_numeric(df['Veiculo'], errors='coerce')
    df['Passag'] = pd.to_numeric(df['Passag'], errors='coerce')
    df['Oferta'] = pd.to_numeric(df['Oferta'], errors='coerce')
    df['Dif'] = pd.to_numeric(df['Dif'], errors='coerce')

    df['%Lotacao'] = (df['Passag'] / df['Oferta'] * 100).fillna(0.0) # Calcula nova coluna '%Lotacao'
    df['%Lotacao'] = df['%Lotacao'].round(2) # Formata para 2 decimais
    df = df[~df['Observacao'].str.contains("Furo", na=False)].reset_index(drop=True) # Dropa linhas com furo de viagens

    return df

def ler_frota():
    # Ler o arquivo matriz da frota
    config = json_utils.ler_json('config.json')
    frota = json_utils.ler_json(config["matrizes"]["frota"])

    df = pd.DataFrame(frota)
    df['Aquisição'] = pd.to_datetime(df['Aquisição'], format="%d/%m/%Y")
    df['Prefixo'] = pd.to_numeric(df['Prefixo'], errors='coerce')
    df["Idade"] = (competencia - df["Aquisição"]).dt.days / 365

    return df

def ler_linhas():
    # Ler o arquivo matriz das linhas
    config = json_utils.ler_json('config.json')
    linhas = json_utils.ler_json(config["matrizes"]["linhas"])

    df = pd.DataFrame(linhas)
    return df

def gerar_resumo():
    st.write("📄 Processando os arquivos...")
    df_frota = ler_frota()
    df_linhas = ler_linhas()
    df = ler_detalhado()

    st.write("🧠 Processando as informações...")
    df = pd.merge(
        df,
        df_frota[['Prefixo', 'Idade']],
        left_on='Veiculo',
        right_on='Prefixo',
        how='left'
    )

    df = df.drop(columns=['Prefixo'])

    cols = df.columns.tolist()
    veiculo_idx = cols.index('Veiculo')
    cols.remove('Idade')
    cols.insert(veiculo_idx + 1, 'Idade')
    df = df[cols]

    df_aggregated = df.groupby('Codigo').agg(
        Linha=('Linha', 'first'),
        Total_THor=('THor', 'count'),
        Total_Real=('Real', 'count'),
        Atrasos=('Dif', lambda x: (x > 5).sum()),
        ate80=('%Lotacao', lambda x: (x <= 80.0).sum()),
        de80a100=('%Lotacao', lambda x: ((x > 80.0) & (x <= 100.00)).sum()),
        maior100=('%Lotacao', lambda x: (x > 100.00).sum()),
        Idade=('Idade', 'sum')
    ).reset_index()

    df_linhas = df_linhas.rename(columns={'Cod_Met': 'Codigo'})

    df_merged_with_linha_raiz = pd.merge(
        df_aggregated,
        df_linhas[['Codigo', 'Raiz', 'Linha', 'Modal']],
        on='Codigo',
        how='left'
    )

    df_agregado_sum_cols = df_merged_with_linha_raiz.groupby('Raiz').agg({
        'Total_THor': 'sum',
        'Total_Real': 'sum',
        'Atrasos': 'sum',
        'ate80': 'sum',
        "de80a100": 'sum',
        "maior100": 'sum',
        'Idade': 'sum'
    }).reset_index()

    df_linha_name = df_merged_with_linha_raiz.groupby('Raiz')['Linha_y'].first().reset_index()
    df_linha_name = df_linha_name.rename(columns={'Linha_y': 'Linha'})

    # Extract 'Modal' for each Raiz group
    df_modal = df_merged_with_linha_raiz.groupby('Raiz')['Modal'].first().reset_index()

    # Merge all three aggregated parts
    df_base = pd.merge(df_linha_name, df_agregado_sum_cols, on='Raiz', how='left')
    df_base = pd.merge(df_base, df_modal, on='Raiz', how='left') # Merge Modal
    df_base = df_base.rename(columns={'Raiz': 'Codigo'})

    # Calculate 'Idade_media', handling division by zero
    df_base['Idade Media'] = df_base.apply(
        lambda row: row['Idade'] / row['Total_Real'] if row['Total_Real'] != 0 else 0,
        axis=1
    )

    # Sorting by 'Modal' with custom order and then by 'Codigo'
    modal_order = ['IO', 'CM', 'SD']
    df_base['Modal'] = pd.Categorical(df_base['Modal'], categories=modal_order, ordered=True)
    df_base = df_base.sort_values(by=['Modal', 'Codigo']).reset_index(drop=True)

    # Apply the conditional update: if 'Modal' is 'SD' and 'maior100' > 0
    condition = (df_base['Modal'] == 'SD') & (df_base['maior100'] > 0)
    df_base.loc[condition, 'de80a100'] += df_base.loc[condition, 'maior100']
    df_base.loc[condition, 'maior100'] = 0

    # Renomeando colunas
    df_base = df_base.rename(columns={"Total_Real": "VG Realizadas"})
    # Criando novas
    df_base["VG Previstas"] = df_base.get("VG Previstas", 0)
    df_base["VG Interrompidas"] = df_base.get("VG Interrompidas", 0)
    df_base["Quebras"] = df_base.get("Quebras", 0)
    df_base["Acidentes"] = df_base.get("Acidentes", 0)
    df_base["Desv. Itinerário"] = df_base.get("Desv. Itinerário", 0)
    df_base["Lotação até 80%"] = ((df_base["ate80"] / df_base["VG Realizadas"]) * 100).round(0)
    df_base["Lotação 80a100%"] = ((df_base["de80a100"] / df_base["VG Realizadas"]) * 100).round(0)
    df_base["Lotação >100%"] = ((df_base["maior100"] / df_base["VG Realizadas"]) * 100).round(0)
    # Excluindo
    df_base = df_base.drop(columns=['Idade'])
    df_base = df_base.drop(columns=['Total_THor'])
    df_base = df_base.drop(columns=['ate80'])
    df_base = df_base.drop(columns=['de80a100'])
    df_base = df_base.drop(columns=['maior100'])

    df_base = atualizar_dados(df_base)

    # Ordenando
    ordem_final = [
        "Linha",
        "Codigo",
        "Modal",
        "Indice VG OF Realizadas",
        "VG Realizadas",
        "VG Previstas",
        "Indice Pontualidade",
        "Atrasos",
        "Desemp. VG Interromp",
        "VG Interrompidas",
        "Idade Media",
        "Indice Quebra",
        "Indice Desv. Itinerário",
        "Lotação até 80%",
        "Lotação 80a100%",
        "Lotação >100%",
        "Indice Acidentes",
        "Quebras",
        "Acidentes",
        "Desv. Itinerário"
    ]

    df_base = df_base[ordem_final]

    return format_utils.arredondar_decimais(df_base, COLUNAS_2_DECIMAIS)

def gerar_xml(df):
    # Formato algumas colunas para duas casas decimais. 301, 304, 306, 308, 309, 310, 314
    cols = ["Indice VG OF Realizadas", "Indice Pontualidade", "Desemp. VG Interromp", "Idade Media", "Indice Quebra", "Indice Desv. Itinerário", "Indice Acidentes"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")
    df[cols] = df[cols].map(lambda x: f"{x:.2f}")
    # Formato algumas colunas para zero casas decimais. 311, 312, 313
    cols = ["Lotação até 80%", "Lotação 80a100%", "Lotação >100%"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")
    df[cols] = df[cols].map(lambda x: f"{x:.0f}")

    root = ET.Element("carga_dados")

    # Cabeçalho fixo
    ET.SubElement(root, "cnpj").text = "90348517000169"
    ET.SubElement(root, "razao_social").text = "Expresso Rio Guaiba Ltda"
    ET.SubElement(root, "mes_ano").text = f"{ano}-{mes}-01"

    # Para cada linha do DataFrame
    for _, row in df.iterrows():

        carga_linha = ET.SubElement(root, "carga_linha")

        ET.SubElement(carga_linha, "cod_linha").text = str(row["Codigo"])
        ET.SubElement(carga_linha, "nome_linha").text = str(row["Linha"])
        ET.SubElement(carga_linha, "modalidade").text = str(row["Modal"])

        # Indicador 301
        ind301 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind301, "id_indicador").text = "301"
        ET.SubElement(ind301, "cod_precisao").text = "A"
        ET.SubElement(ind301, "val_indicador").text = str(row["Indice VG OF Realizadas"])

        # Indicador 302
        ind302 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind302, "id_indicador").text = "302"
        ET.SubElement(ind302, "cod_precisao").text = "A"
        ET.SubElement(ind302, "val_indicador").text = str(row["VG Realizadas"])

        # Indicador 303
        ind303 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind303, "id_indicador").text = "303"
        ET.SubElement(ind303, "cod_precisao").text = "A"
        ET.SubElement(ind303, "val_indicador").text = str(row["VG Previstas"])

        # Indicador 304
        ind304 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind304, "id_indicador").text = "304"
        ET.SubElement(ind304, "cod_precisao").text = "B"
        ET.SubElement(ind304, "val_indicador").text = str(row["Indice Pontualidade"])

        # Indicador 305
        ind305 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind305, "id_indicador").text = "305"
        ET.SubElement(ind305, "cod_precisao").text = "B"
        ET.SubElement(ind305, "val_indicador").text = str(row["Atrasos"])

        # Indicador 306
        ind306 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind306, "id_indicador").text = "306"
        ET.SubElement(ind306, "cod_precisao").text = "A"
        ET.SubElement(ind306, "val_indicador").text = str(row["Desemp. VG Interromp"])

        # Indicador 307
        ind307 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind307, "id_indicador").text = "307"
        ET.SubElement(ind307, "cod_precisao").text = "A"
        ET.SubElement(ind307, "val_indicador").text = str(row["VG Interrompidas"])

        # Indicador 308
        ind308 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind308, "id_indicador").text = "308"
        ET.SubElement(ind308, "cod_precisao").text = "A"
        ET.SubElement(ind308, "val_indicador").text = str(row["Idade Media"])

        # Indicador 309
        ind309 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind309, "id_indicador").text = "309"
        ET.SubElement(ind309, "cod_precisao").text = "A"
        ET.SubElement(ind309, "val_indicador").text = str(row["Indice Quebra"])

        # Indicador 310
        ind310 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind310, "id_indicador").text = "310"
        ET.SubElement(ind310, "cod_precisao").text = "A"
        ET.SubElement(ind310, "val_indicador").text = str(row["Indice Desv. Itinerário"])

        # Indicador 311
        ind311 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind311, "id_indicador").text = "311"
        ET.SubElement(ind311, "cod_precisao").text = "A"
        ET.SubElement(ind311, "val_indicador").text = str(row["Lotação até 80%"])

        # Indicador 312
        ind312 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind312, "id_indicador").text = "312"
        ET.SubElement(ind312, "cod_precisao").text = "A"
        ET.SubElement(ind312, "val_indicador").text = str(row["Lotação 80a100%"])

        # Indicador 313
        ind313 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind313, "id_indicador").text = "313"
        ET.SubElement(ind313, "cod_precisao").text = "A"
        ET.SubElement(ind313, "val_indicador").text = str(row["Lotação >100%"])

        # Indicador 314
        ind314 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind314, "id_indicador").text = "314"
        ET.SubElement(ind314, "cod_precisao").text = "A"
        ET.SubElement(ind314, "val_indicador").text = str(row["Indice Acidentes"])

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
# endregion

# region CONSTANTES
COLUNAS_2_DECIMAIS = [
    "Indice VG OF Realizadas",
    "Indice Pontualidade",
    "Desemp. VG Interromp",
    "Idade Media",
    "Indice Quebra",
    "Indice Desv. Itinerário",
    "Indice Acidentes"
]

# Colunas para mostrar help 
COLUNAS_EDITAVEIS = [
    "VG Realizadas",
    "VG Previstas",
    "Quebras",
    "Acidentes",
    "Desv. Itinerário"
]
# endregion

st.set_page_config(layout="wide")
st.subheader("Indicadores AGERGS")
st.divider()

# CSS para limitar apenas o formulário
st.markdown(
    """
    <style>
    /* Limita largura de TODOS os forms da página */
    div[data-testid="stForm"] {
        max-width: 750px;
        margin-left: auto;
        margin-right: auto;
        padding-top: 20px;
        padding-bottom: 40px;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# Container do formulário
with st.container():
    with st.form("form_inputs"):
        st.subheader("Período de competência")

        col_mes, col_ano, col_space = st.columns([2, 1, 4])
        with col_mes:
            meses = {
                "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
                "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
            }
            mes_nome = st.selectbox("Mês", list(meses.keys()))

        with col_ano:
            ano = st.number_input("Ano", min_value=2000, max_value=2100, value=2025)

        mes = meses[mes_nome]
        ultimo_dia = calendar.monthrange(int(ano), mes)[1]
        competencia = pd.Timestamp(int(ano), mes, ultimo_dia)

        st.subheader("Arquivo Transnet")
        up_detalhado = st.file_uploader(
            "Relatório Controle Operacional Detalhado por Linha",
            type="csv",
            key="upload_detalhado"
        )

        gerar = st.form_submit_button("Gerar resumo", type="primary")

# ✳️ Processamento no submit
if gerar:
    if up_detalhado is None:
        st.error("❌ Você precisa selecionar o arquivo CSV antes de continuar.")
        st.stop()

    try:
        with st.status("Processando...", expanded=False) as status:
            df_base = gerar_resumo()
            st.write("💾 Salvando...")
            
            st.session_state.df = df_base.copy() # salva o df inicial
            st.session_state["mostrar_resumo"] = True # ativa o modo resumo

            status.update(label="Processo terminado!", state="complete", expanded=False)
            st.success("Arquivo gerado com sucesso!")

    except Exception:
        status.update(label="Erro durante o processamento!", state="error")
        st.error(f"🐞 Erro inesperado:\n\n```\n{traceback.format_exc()}\n```")
    
    finally:
        df = None
        df_base = None
        df_editado = None
        

# Criar configuração de colunas com help
if "df" in st.session_state:
    column_config = {}
    for col in st.session_state.df.columns:
        if col in COLUNAS_EDITAVEIS:
            column_config[col] = st.column_config.NumberColumn(
                col,
                help="Coluna editável",
            )

# ✳️ Mostrar data editor
if st.session_state.get("mostrar_resumo", False):

    df_editado = st.data_editor(
        st.session_state.df,
        key="editor_resumo",
        disabled=[col for col in st.session_state.df.columns if col not in COLUNAS_EDITAVEIS],
        hide_index=True,
        height="content",
        column_config=column_config
    )

    with st.container():
        col3, col4, col5 = st.columns([1.2, 1.2, 8])

        with col3:
            # Botão — só ativa o gatilho
            if st.button("🔄 Atualizar dados"):
                st.session_state.df = format_utils.arredondar_decimais(df_editado, COLUNAS_2_DECIMAIS)
                st.session_state["recalcular"] = True
                st.rerun()

        with col4:
            # Botão — Exporta XML
            st.download_button(
                label="📥 Baixar XML",
                data=gerar_xml(st.session_state.df.copy()),
                file_name="9034851700169-ITM-" + str(ano) + str(mes) + ".xml",
                mime="application/xml"
            )

    # ✳️ REPROCESSAMENTO — acontece em um rerun separado
    if st.session_state.get("recalcular", False):
        
        df = atualizar_dados(st.session_state.df.copy()) # Atualizar os dados
        st.session_state.df = format_utils.arredondar_decimais(df, COLUNAS_2_DECIMAIS)
        st.session_state["recalcular"] = False
        st.rerun()
    




