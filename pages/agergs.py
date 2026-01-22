import traceback
import calendar
from datetime import datetime
from utils import format_utils
from utils import files_utils
import xml.etree.ElementTree as ET
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from copy import copy

# region FUNÇÕES
def column_help():
    column_help = {} # Criar configuração de colunas com help
    for col in st.session_state.df.columns:
        if col in COLUNAS["Coluna"]:
            idx = COLUNAS["Coluna"].index(col)
            column_help[col] = st.column_config.NumberColumn(
                col,
                help=str(COLUNAS["Descrição"][idx]),
            )
    return column_help

def atualizar_dados(df):
    def safe_div(a, b):
        return a / b if (b not in (0, None) and b != 0) else 0

    df["301"] = df.apply( # Índice de viagens oficiais realizadas
        lambda x: safe_div(x["302"], x["303"]), # Viagens realizadas * Viagens previstas
        axis=1
    )
    df["304"] = df.apply( # Índice pontualidade
        lambda x: safe_div(x["302"] - x["305"], x["302"]), # (Viagens realizadas - Atrasos) * Viagens realizadas
        axis=1
    )
    df["306"] = df.apply( # Desempenho de viagens interrompidas
        lambda x: safe_div(x["302"] - x["307"], x["302"]), # (Viagens realizadas - Viagens interrompidas) * Viagens realizadas
        axis=1
    )
    df["307"] = df["Quebras"] + df["Acidentes"] # Viagens interrompidas
    df["309"] = df.apply( # Índice Quebra
        lambda x: safe_div(x["Quebras"], x["302"]) * 1000, # Quebras * Viagens realizadas * 1000
        axis=1
    )
    df["310"] = df.apply( # Índice desvio de itinerário
        lambda x: safe_div(x["Desv. Itinerário"], x["302"]) * 1000, # Desvio itinerário * Viagens realizadas * 1000
        axis=1
    )
    df["314"] = df.apply( # Índice acidentes
        lambda x: safe_div(x["Acidentes"], x["302"]) * 1000, # Acidentes * Viagens realizadas * 1000
        axis=1
    )

    return df

def ler_detalhado():
    df = files_utils.ler_detalhado_linha(up_detalhado) # Lê o arquivo
    # Verifica se o período e o arquivo conferem
    periodo = pd.to_datetime(df.iloc[0]["Dia"], dayfirst=True, errors="coerce").date()
    if periodo.month != mes or periodo.year != ano:
        status.update(label="Erro durante o processamento!", state="error", expanded=True)
        st.warning(f"⚠️ O arquivo não confere com o período informado! Arquivo: {periodo.month}/{periodo.year}")
        st.stop()
    
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

def gerar_resumo():
    st.write("📄 Processando os arquivos...")
    df_frota = files_utils.ler_frota(competencia)
    df_linhas = files_utils.ler_linhas()
    df_raiz = files_utils.ler_linhas_raiz()
    df_previstas = files_utils.ler_viagens_previstas(up_previstas)
    df = ler_detalhado()

    st.write("🧠 Processando as informações...")
    # Merge com a frota
    df = pd.merge(
        df,
        df_frota[['Prefixo', 'Idade']],
        left_on='Veiculo',
        right_on='Prefixo',
        how='left'
    )
    # Organiza colunas
    df = df.drop(columns=['Prefixo'])
    cols = df.columns.tolist()
    veiculo_idx = cols.index('Veiculo')
    cols.remove('Idade')
    cols.insert(veiculo_idx + 1, 'Idade')
    df = df[cols]
    # Agrega totais
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
    # Totaliza previstas
    cols_somar = ["U1", "S1", "D1", "U2", "S2", "D2"]
    df_previstas["TPrevMet"] = df_previstas[cols_somar].sum(axis=1)
    df_previstas = df_previstas[["Codigo", "TPrevMet"]]
    df_aggregated = df_aggregated.merge(df_previstas[["Codigo", "TPrevMet"]], on="Codigo", how="left")
    df_aggregated["TPrevMet"] = df_aggregated["TPrevMet"].fillna(0).astype(int)

    # Adiciona coluna nome_raiz
    df_linhas = df_linhas.merge(
        df_raiz[['Cod_Raiz', 'Modal', 'Nome_Raiz']],
        on=['Cod_Raiz', 'Modal'],
        how="left"
    )
    df_linhas = df_linhas.rename(columns={'Cod_Met': 'Codigo'})
    # Adiciona coluna cod_raiz
    df_merged_raiz = pd.merge(
        df_aggregated,
        df_linhas[['Codigo', 'Cod_Raiz', "Modal"]],
        on='Codigo',
        how='left'
    ).drop(columns=["Codigo"])
    # Agrupa por raíz e soma
    df_agregado_sum_cols = df_merged_raiz.groupby(["Cod_Raiz", "Modal"]).agg({
        'Total_THor': 'sum',
        'Total_Real': 'sum',
        'TPrevMet': 'sum',
        'Atrasos': 'sum',
        'ate80': 'sum',
        "de80a100": 'sum',
        "maior100": 'sum',
        'Idade': 'sum'
    }).reset_index()
    # Retira duplicados
    df_linhas = df_linhas.drop_duplicates(subset=['Cod_Raiz', 'Modal'])
    # Adiciona Nome_Raiz
    df_base = df_agregado_sum_cols.merge(
        df_linhas[['Cod_Raiz', 'Modal', 'Nome_Raiz']],
        on=['Cod_Raiz', 'Modal'],
        how='left'
    )

    df_base = df_base.rename(columns={"Cod_Raiz": "Codigo", "Nome_Raiz": "Linha"})
    
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
    df_base = df_base.rename(columns={"Indice VG OF Realizadas" : "301", "Total_Real": "302", "TPrevMet": "303", "Indice Pontualidade" : "304",
                                      "Atrasos" : "305", "Desemp. VG Interromp" : "306", "VG Interrompidas" : "307", "Idade Media" : "308",
                                      "Indice Quebra" : "309", "Indice Desv. Itinerário" : "310", "Indice Acidentes" : "314"})

    # Criando novas
    df_base["307"] = df_base.get("307", 0) # Viagens interrompidas
    df_base["Quebras"] = df_base.get("Quebras", 0)
    df_base["Acidentes"] = df_base.get("Acidentes", 0)
    df_base["Desv. Itinerário"] = df_base.get("Desv. Itinerário", 0)
    df_base["311"] = ((df_base["ate80"] / df_base["302"]) * 100).round(0)
    df_base["312"] = ((df_base["de80a100"] / df_base["302"]) * 100).round(0)
    df_base["313"] = ((df_base["maior100"] / df_base["302"]) * 100).round(0)
    # Excluindo
    df_base = df_base.drop(columns=['Idade'])
    df_base = df_base.drop(columns=['Total_THor'])
    df_base = df_base.drop(columns=['ate80'])
    df_base = df_base.drop(columns=['de80a100'])
    df_base = df_base.drop(columns=['maior100'])

    df_base = atualizar_dados(df_base)

    # Atualiza os valores de lotações [311, 312, 313] das alimentadoras para 999
    df_base.loc[df_base['Modal'] == 'IO', ['311', '312', '313']] = 999

    # Ordenando
    ordem_final = [
        "Linha",
        "Codigo",
        "Modal",
        "301",
        "302",
        "303",
        "304",
        "305",
        "306",
        "307",
        "308",
        "309",
        "310",
        "311",
        "312",
        "313",
        "314",
        "Quebras",
        "Acidentes",
        "Desv. Itinerário"
    ]

    df_base = df_base[ordem_final]

    return format_utils.arredondar_decimais(df_base, COLUNAS_2_DECIMAIS)

def gerar_xml(df):
    # Formato algumas colunas para duas casas decimais. 301, 304, 306, 308, 309, 310, 314
    cols = ["301", "304", "306", "308", "309", "310", "314"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")
    df[cols] = df[cols].map(lambda x: f"{x:.2f}")
    # Formato algumas colunas para zero casas decimais. 311, 312, 313
    cols = ["311", "312", "313"]
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

        # Indicador 301 - Indice de Viagens Oficiais Realizadas
        ind301 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind301, "id_indicador").text = "301"
        ET.SubElement(ind301, "cod_precisao").text = "A"
        ET.SubElement(ind301, "val_indicador").text = str(row["301"])

        # Indicador 302 - Viagens Realizadas
        ind302 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind302, "id_indicador").text = "302"
        ET.SubElement(ind302, "cod_precisao").text = "A"
        ET.SubElement(ind302, "val_indicador").text = str(row["302"])

        # Indicador 303 - Viagens Previstas
        ind303 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind303, "id_indicador").text = "303"
        ET.SubElement(ind303, "cod_precisao").text = "A"
        ET.SubElement(ind303, "val_indicador").text = str(row["303"])

        # Indicador 304 - Indice de Pontualidade
        ind304 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind304, "id_indicador").text = "304"
        ET.SubElement(ind304, "cod_precisao").text = "B"
        ET.SubElement(ind304, "val_indicador").text = str(row["304"])

        # Indicador 305 - Atrasos
        ind305 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind305, "id_indicador").text = "305"
        ET.SubElement(ind305, "cod_precisao").text = "B"
        ET.SubElement(ind305, "val_indicador").text = str(row["305"])

        # Indicador 306 - Desempenho Viagens Interrompidas
        ind306 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind306, "id_indicador").text = "306"
        ET.SubElement(ind306, "cod_precisao").text = "A"
        ET.SubElement(ind306, "val_indicador").text = str(row["306"])

        # Indicador 307 - Viagens Interrompidas
        ind307 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind307, "id_indicador").text = "307"
        ET.SubElement(ind307, "cod_precisao").text = "A"
        ET.SubElement(ind307, "val_indicador").text = str(row["307"])

        # Indicador 308 - Idade Media
        ind308 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind308, "id_indicador").text = "308"
        ET.SubElement(ind308, "cod_precisao").text = "A"
        ET.SubElement(ind308, "val_indicador").text = str(row["308"])

        # Indicador 309 - Indice Quebra
        ind309 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind309, "id_indicador").text = "309"
        ET.SubElement(ind309, "cod_precisao").text = "A"
        ET.SubElement(ind309, "val_indicador").text = str(row["309"])

        # Indicador 310 - Indice Desvio Itinerário
        ind310 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind310, "id_indicador").text = "310"
        ET.SubElement(ind310, "cod_precisao").text = "A"
        ET.SubElement(ind310, "val_indicador").text = str(row["310"])

        # Indicador 311 - Lotação até 80%
        ind311 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind311, "id_indicador").text = "311"
        ET.SubElement(ind311, "cod_precisao").text = "A"
        ET.SubElement(ind311, "val_indicador").text = str(row["311"])

        # Indicador 312 - Lotação 80a100%
        ind312 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind312, "id_indicador").text = "312"
        ET.SubElement(ind312, "cod_precisao").text = "A"
        ET.SubElement(ind312, "val_indicador").text = str(row["312"])

        # Indicador 313 - Lotação >100%
        ind313 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind313, "id_indicador").text = "313"
        ET.SubElement(ind313, "cod_precisao").text = "A"
        ET.SubElement(ind313, "val_indicador").text = str(row["313"])

        # Indicador 314 - Indice Acidentes
        ind314 = ET.SubElement(carga_linha, "carga_indicador")
        ET.SubElement(ind314, "id_indicador").text = "314"
        ET.SubElement(ind314, "cod_precisao").text = "A"
        ET.SubElement(ind314, "val_indicador").text = str(row["314"])

    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
# endregion

# region CONSTANTES
COLUNAS_2_DECIMAIS = [
    "301",
    "304",
    "306",
    "308",
    "309",
    "310",
    "314"
]

COLUNAS_EDITAVEIS = [
    "302",
    "303",
    "Quebras",
    "Acidentes",
    "Desv. Itinerário"
]
# Colunas para mostrar help 
COLUNAS = {
    "Coluna": ["301", "302", "303", "304", "305", "306", "307", "308", "309", "310", "311", "312", "313", "314"],
    "Descrição": ["Indice de Viagens Oficiais Realizadas", "Viagens Realizadas", "Viagens Previstas", "Indice de Pontualidade", "Atrasos",
             "Desempenho Viagens Interrompidas", "Viagens Interrompidas", "Idade Media", "Indice Quebra", "Indice Desvio de Itinerário", 
             "Lotação até 80%", "Lotação de 80 a 100%", "Lotação maior 100%", "Indice Acidentes"
        ],
}
# endregion

st.set_page_config(layout="wide")
st.header("📄 Indicadores AGERGS", anchor=False)
st.divider()

# Inputs
c1, c2, c3 = st.columns([1.8, 2, 2], gap="medium")
with c1:
    st.subheader("Período de competência")
    col_mes, col_ano = st.columns([3, 2])
    with col_mes:
        meses = {
            "Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6,
            "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12
        }
        mes_nome = st.selectbox("Mês", list(meses.keys()))
    with col_ano:
        ano = st.number_input("Ano", min_value=2000, max_value=2100, value=datetime.today().year)

with c2:
    st.subheader("Dados de viagens", help="Transnet > Módulos > Tráfego/Arrecadação > Consultas/Relatórios > Controle Operacional/Tráfego > Controle Operacional Detalhado Por Linha", anchor=False)
    up_detalhado = st.file_uploader("Relatório Controle Operacional Detalhado por Linha", type="csv", key="upload_detalhado")

with c3:
    st.subheader("Viagens previstas", help="Planilha de viagens previstas", anchor=False)
    up_previstas = st.file_uploader("Selecione um arquivo .XLSX", type='xlsx', key='upload_previstas')

mes = meses[mes_nome]
ultimo_dia = calendar.monthrange(int(ano), mes)[1]
competencia = pd.Timestamp(int(ano), mes, ultimo_dia)

gerar = st.sidebar.button("Gerar resumo", type="primary")

# ✳️ Processamento no submit
if gerar:
    # Remove os botões
    st.session_state.pop("agergs", None)

    if up_detalhado is None:
        st.error("⚠️ Você precisa selecionar o arquivo do relatório controle operacional detalhado por linha antes de continuar.")
        st.stop()
    
    if up_previstas is None:
        st.error("⚠️ Você precisa selecionar a planilha de viagens previstas antes de continuar.")
        st.stop()

    try:
        with st.status("Processando...", expanded=False) as status:
            df_base = gerar_resumo()
            st.write("💾 Salvando...")
            
            st.session_state.df = df_base.copy() # salva o df inicial
            st.session_state["agergs"] = True # ativa o modo resumo

            status.update(label="Processo terminado!", state="complete", expanded=False)
            st.success("Arquivo gerado com sucesso!")

    except Exception:
        status.update(label="Erro durante o processamento!", state="error")
        st.error(f"🐞 Erro inesperado:\n\n```\n{traceback.format_exc()}\n```")
    
    finally:
        df = None
        df_base = None
        df_editado = None


# ✳️ Mostrar data editor
if st.session_state.get("agergs", False):
    with st.expander("ℹ️ Glossário de colunas"):
        st.table(COLUNAS, border="horizontal")

    df_editado = st.data_editor(
        st.session_state.df,
        key="editor_resumo",
        disabled=[col for col in st.session_state.df.columns if col not in COLUNAS_EDITAVEIS],
        hide_index=True,
        height="content",
        column_config=column_help()
    )

    with st.container():
        col3, col4, col5 = st.columns([1.2, 1.2, 8], gap="small")
        with col3:
            # Botão — só ativa o gatilho
            if st.button("🔄 Atualizar tabela"):
                st.session_state.df = format_utils.arredondar_decimais(df_editado, COLUNAS_2_DECIMAIS)
                st.session_state["recalcular_agergs"] = True
                st.rerun()

        with col4:
            # Botão — Exporta XML
            st.download_button(
                label="📥 Download XML",
                data=gerar_xml(st.session_state.df.copy()),
                file_name="9034851700169-ITM-" + str(ano) + str(mes) + ".xml",
                mime="application/xml"
            )

    # ✳️ REPROCESSAMENTO — acontece em um rerun separado
    if st.session_state.get("recalcular_agergs", False):
        df = atualizar_dados(st.session_state.df.copy()) # Atualizar os dados
        st.session_state.df = format_utils.arredondar_decimais(df, COLUNAS_2_DECIMAIS)
        st.session_state["recalcular_agergs"] = False
        st.rerun()
    




