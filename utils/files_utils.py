import io
import re
import pandas as pd
from utils import json_utils

# Ler o arquivo matriz da frota
def ler_frota(data):
    config = json_utils.ler_json('config.json')
    frota = json_utils.ler_json(config["matrizes"]["frota"])

    df = pd.DataFrame(frota)
    df['Aquisição'] = pd.to_datetime(df['Aquisição'], format="%d/%m/%Y")
    df['Prefixo'] = pd.to_numeric(df['Prefixo'], errors='coerce')
    df["Idade"] = (data - df["Aquisição"]).dt.days / 365

    return df

# Ler o arquivo matriz das linhas
def ler_linhas():
    config = json_utils.ler_json('config.json')
    linhas = json_utils.ler_json(config["matrizes"]["linhas"])

    df = pd.DataFrame(linhas)
    return df

# Ler o arquivo matriz das linhas raiz
def ler_linhas_raiz():
    config = json_utils.ler_json('config.json')
    raiz = json_utils.ler_json(config["matrizes"]["linhas_raiz"])

    df = pd.DataFrame(raiz)
    return df

# Ler o arquivo do relatório detalhado por linha
def ler_detalhado_linha(file):
    # Ler o arquivo do Transnet
    linhas_limpas = []
    codigo_linha = None
    nome_linha = None
    sentido = 1
    data_dia = None
    dentro_tabela = False
    # Cabeçalho
    header = "Codigo;Linha;Dia;Sent;#;THor;Real;Orig;Dest;Dif;Parado;Prev;Real2;Dif2;Km_h;CVg;Veiculo;Docmto;Motorista;Cobrador;EmPe;Sent;Oferta;Meta;Passag;CVg2;TipoViagem;Observacao"
    linhas_limpas.append(header)

    for linha_original in file.getvalue().decode("latin-1").splitlines():
        linha_original = linha_original.strip()

        # Detectar início de um novo dia
        if linha_original.startswith("Dia:"):
            dentro_tabela = False

            # Extrair data
            m = re.search(r"Dia:\s*(\d{2}/\d{2}/\d{4})", linha_original)
            if m:
                data_dia = m.group(1)

            # Extrair nome da linha e código da linha
            if "Linha:" in linha_original and "|" in linha_original:
                # Extracts the full line identifier, e.g., "Linha: A131B ALIMENTADORA COLINA"
                full_line_info_str = linha_original.split("|", 1)[1].split("Sentido")[0].strip()
                
                # Remove "Linha: " prefix and strip any leading/trailing whitespace
                line_content_no_prefix = full_line_info_str.replace("Linha:", "", 1).strip()

                # Split the content by the first space to get code and name
                parts = line_content_no_prefix.split(' ', 1)
                codigo_linha = parts[0]
                nome_linha = parts[1] if len(parts) > 1 else ""

                # Sentido
                sentido = 2 if linha_original.split("Sentido:", 1)[1].strip() == "Volta" else 1
            continue

        # Detectar cabeçalho real (this indicates the start of a data block)
        if linha_original.startswith(";#;"):
            dentro_tabela = True
            continue

        # Processar linhas da tabela
        if dentro_tabela and linha_original:
            linha_cleaned = linha_original.replace('"', '')
            linha_cleaned = linha_cleaned.replace(',', '.')

            # Adicionar Codigo, Linha (nome), e Dia no início e depois a linha processada
            # lstrip(';') is used to remove the leading semicolon from `linha_cleaned` to avoid an extra empty column.
            linhas_limpas.append(f"{codigo_linha};{nome_linha};{data_dia};{sentido};{linha_cleaned.lstrip(';')}")

    # Criar um objeto StringIO para ler a lista de linhas como se fosse um arquivo
    data_io = io.StringIO('\n'.join(linhas_limpas))

    # Ler no pandas diretamente do StringIO
    # index_col=False is added to prevent pandas from using the first column as an index.
    df = pd.read_csv(data_io, sep=";", index_col=False)

    df = df.dropna(subset=['Observacao']).reset_index(drop=True)

    return df 

# Ler a planilha de viagens previstas
def ler_viagens_previstas(file):
    df = pd.read_excel(file, skiprows=6)
    df.rename(columns={'Unnamed: 0': 'Codigo', 'TOTAL': 'MET_U1', 'TOTAL.1': 'MET_S1', 'TOTAL.2': 'MET_D1', 'TOTAL.3': 'MET_U2', 'TOTAL.4': 'MET_S2', 'TOTAL.5': 'MET_D2'}, inplace=True)
    df = df.dropna(subset=['Codigo', 'DIAS'])
    df = df.drop(columns=['ÚTEIS', 'DIAS', 'SAB', 'DIAS.1', 'DOM', 'DIAS.2', 'Unnamed: 10', 'ÚTEIS.1', 'DIAS.3', 'SAB.1', 'DIAS.4', 'DOM.1', 'DIAS.5', 'Unnamed: 20'])
    df = df[~df["Codigo"].str.contains("TOTAL", case=False, na=False)]
    df = df.sort_values("Codigo")

    return df

#def ler_desempenho_diario_linha(file):