# VERSION_FINAL_PRODUCAO_SEGURANCA_DASHBOARD_V6_DUCKDB
# App Streamlit para dados de seguranca publica
# Fonte principal: base oficial MJSP/SINESP em XLSX
# Fonte experimental: API comunitaria rayonnunes/api_seguranca_publica
# V6: usa cache em Parquet + DuckDB para filtrar antes de materializar a base no Streamlit

import io
import json
import os
import re
import tempfile
import unicodedata
from datetime import datetime

import duckdb
import pandas as pd
import requests
import streamlit as st
import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# -----------------------------------------------------------------------------
# CONFIGURACAO DA PAGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Inteligencia Territorial | Seguranca Publica",
    page_icon="🛡️",
    layout="wide",
)

st.markdown(
    """
    <style>
    .header-seguranca {
        background-color: #1c2d42;
        padding: 20px;
        color: white;
        border-radius: 5px;
        text-align: center;
        margin-bottom: 20px;
    }
    .stButton>button {
        background-color: #1c2d42;
        color: white;
        width: 100%;
    }
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1c2d42;
        text-align: center;
        margin-bottom: 15px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        min-height: 132px;
    }
    .small-note {
        color: #666;
        font-size: 13px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# FONTES DE DADOS
# -----------------------------------------------------------------------------
URL_MJSP_MUNICIPIOS = (
    "https://dados.mj.gov.br/dataset/210b9ae2-21fc-4986-89c6-2006eb4db247/"
    "resource/03af7ce2-174e-4ebd-b085-384503cfb40f/download/"
    "indicadoressegurancapublicamunic.xlsx"
)

URL_MJSP_UF = (
    "https://dados.mj.gov.br/dataset/210b9ae2-21fc-4986-89c6-2006eb4db247/"
    "resource/feeae05e-faba-406c-8a4a-512aec91a9d1/download/"
    "indicadoressegurancapublicauf.xlsx"
)

API_RAYONNUNES = "http://ec2-54-174-4-15.compute-1.amazonaws.com/api"

UFS = [
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS",
    "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC",
    "SP", "SE", "TO",
]

UF_NOMES = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapa", "AM": "Amazonas",
    "BA": "Bahia", "CE": "Ceara", "DF": "Distrito Federal", "ES": "Espirito Santo",
    "GO": "Goias", "MA": "Maranhao", "MT": "Mato Grosso", "MS": "Mato Grosso do Sul",
    "MG": "Minas Gerais", "PA": "Para", "PB": "Paraiba", "PR": "Parana",
    "PE": "Pernambuco", "PI": "Piaui", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte",
    "RS": "Rio Grande do Sul", "RO": "Rondonia", "RR": "Roraima", "SC": "Santa Catarina",
    "SP": "Sao Paulo", "SE": "Sergipe", "TO": "Tocantins",
}

MAPA_CRIMES_API = {
    "Todos os crimes": None,
    "1 - Estupro": "1",
    "2 - Furto de veiculo": "2",
    "3 - Homicidio doloso": "3",
    "4 - Lesao corporal seguida de morte": "4",
    "5 - Roubo a instituicao financeira": "5",
    "6 - Roubo de carga": "6",
    "7 - Roubo de veiculo": "7",
    "8 - Roubo seguido de morte (latrocinio)": "8",
    "9 - Tentativa de homicidio": "9",
}

REVERSE_CRIMES_API = {
    v: k.split(" - ", 1)[1] for k, v in MAPA_CRIMES_API.items() if v is not None
}

MESES_DISPLAY = {
    "Todos os meses": None,
    "01 - Janeiro": 1,
    "02 - Fevereiro": 2,
    "03 - Marco": 3,
    "04 - Abril": 4,
    "05 - Maio": 5,
    "06 - Junho": 6,
    "07 - Julho": 7,
    "08 - Agosto": 8,
    "09 - Setembro": 9,
    "10 - Outubro": 10,
    "11 - Novembro": 11,
    "12 - Dezembro": 12,
}

MES_API = {
    1: "jan", 2: "fev", 3: "mar", 4: "abr", 5: "mai", 6: "jun",
    7: "jul", 8: "ago", 9: "set", 10: "out", 11: "nov", 12: "dez",
}

MES_LABEL = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
}

MES_NOME_LONGO = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
}

MES_NORMALIZADO = {
    "1": 1, "01": 1, "JAN": 1, "JANEIRO": 1,
    "2": 2, "02": 2, "FEV": 2, "FEVEREIRO": 2,
    "3": 3, "03": 3, "MAR": 3, "MARCO": 3, "MARÇO": 3,
    "4": 4, "04": 4, "ABR": 4, "ABRIL": 4,
    "5": 5, "05": 5, "MAI": 5, "MAIO": 5,
    "6": 6, "06": 6, "JUN": 6, "JUNHO": 6,
    "7": 7, "07": 7, "JUL": 7, "JULHO": 7,
    "8": 8, "08": 8, "AGO": 8, "AGOSTO": 8,
    "9": 9, "09": 9, "SET": 9, "SETEMBRO": 9,
    "10": 10, "OUT": 10, "OUTUBRO": 10,
    "11": 11, "NOV": 11, "NOVEMBRO": 11,
    "12": 12, "DEZ": 12, "DEZEMBRO": 12,
}

COLUNAS_CANONICAS = {
    "UF": [
        "UF", "SIGLA_UF", "SIGLA_DA_UF", "UNIDADE_DA_FEDERACAO", "UNIDADE_FEDERACAO",
        "ESTADO", "ESTADO_SIGLA",
    ],
    "MUNICIPIO": [
        "MUNICIPIO", "MUNICIPIO_IBGE", "NOME_MUNICIPIO", "NOME_DO_MUNICIPIO",
        "MUNICIPIO_NOME", "CIDADE",
    ],
    "CODIGO_MUNICIPIO": [
        "CODIGO_MUNICIPIO", "COD_MUNICIPIO", "CODIGO_IBGE", "COD_IBGE", "IBGE",
        "CODMUN", "ID_MUNICIPIO",
    ],
    "REGIAO": ["REGIAO", "REGIAO_BRASIL"],
    "ANO": ["ANO", "ANO_REFERENCIA", "ANO_DE_REFERENCIA", "PERIODO_ANO"],
    "MES": ["MES", "MÊS", "MES_REFERENCIA", "MES_DE_REFERENCIA", "PERIODO_MES"],
    "TIPO_CRIME": [
        "CRIME", "TIPO_CRIME", "TIPO_DE_CRIME", "NATUREZA", "INDICADOR",
        "INDICADOR_CRIMINAL", "TIPIFICACAO", "TIPIFICACAO_DO_DELITO",
    ],
    "OCORRENCIAS": [
        "OCORRENCIAS", "OCORRENCIA", "TOTAL_OCORRENCIAS", "QTDE_OCORRENCIAS",
        "QTD_OCORRENCIAS", "QUANTIDADE_OCORRENCIAS", "REGISTROS", "TOTAL", "VALOR",
    ],
    "VITIMAS": [
        "VITIMAS", "VITIMA", "TOTAL_VITIMAS", "QTDE_VITIMAS", "QTD_VITIMAS",
        "QUANTIDADE_VITIMAS",
    ],
    "VITIMAS_MUNICIPIO": [
        "VITIMAS_MUNICIPIO", "VITIMAS_NO_MUNICIPIO", "TOTAL_VITIMAS_MUNICIPIO",
    ],
}

COLUNAS_MESES_WIDE = {
    "JAN": 1, "JANEIRO": 1,
    "FEV": 2, "FEVEREIRO": 2,
    "MAR": 3, "MARCO": 3, "MARÇO": 3,
    "ABR": 4, "ABRIL": 4,
    "MAI": 5, "MAIO": 5,
    "JUN": 6, "JUNHO": 6,
    "JUL": 7, "JULHO": 7,
    "AGO": 8, "AGOSTO": 8,
    "SET": 9, "SETEMBRO": 9,
    "OUT": 10, "OUTUBRO": 10,
    "NOV": 11, "NOVEMBRO": 11,
    "DEZ": 12, "DEZEMBRO": 12,
}

# -----------------------------------------------------------------------------
# FUNCOES AUXILIARES
# -----------------------------------------------------------------------------
def remover_acentos(texto):
    if pd.isna(texto):
        return ""
    return unicodedata.normalize("NFKD", str(texto)).encode("ASCII", "ignore").decode("ASCII")


def normalizar_coluna(nome):
    texto = remover_acentos(nome).upper().strip()
    texto = re.sub(r"[^A-Z0-9]+", "_", texto)
    texto = re.sub(r"_+", "_", texto).strip("_")
    return texto


def normalizar_valor(texto):
    texto = remover_acentos(texto).upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def chave_comparacao(texto):
    texto = normalizar_valor(texto)
    texto = re.sub(r"[^A-Z0-9]", "", texto)
    return texto


def normalizar_municipio_api(nome_municipio):
    if not nome_municipio:
        return ""
    nome_alterado = remover_acentos(nome_municipio).lower().strip()
    nome_alterado = re.sub(r"\s+", "", nome_alterado)
    return nome_alterado


def converter_numero(valor):
    if pd.isna(valor):
        return 0
    texto = str(valor).strip()
    if texto in ["", "-", "--", "nan", "None"]:
        return 0
    texto = texto.replace("\u00a0", "")
    texto = re.sub(r"[^0-9,.-]", "", texto)

    # Padrao brasileiro: 1.234,56
    if "," in texto and "." in texto:
        texto = texto.replace(".", "").replace(",", ".")
    elif "," in texto:
        texto = texto.replace(",", ".")

    try:
        return float(texto)
    except Exception:
        return 0


def formatar_inteiro(valor):
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")
    except Exception:
        return "0"


def parse_data_mista(texto_original):
    """Converte datas sem gerar warning de dayfirst em formato ISO.

    O aviso informado no log ocorre quando pandas recebe uma string no formato
    YYYY-MM-DD HH:MM:SS e, ao mesmo tempo, dayfirst=True. Aqui detectamos o
    formato antes e passamos format explicito.
    """
    if pd.isna(texto_original):
        return pd.NaT
    s = str(texto_original).strip()
    if not s:
        return pd.NaT

    if re.fullmatch(r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}", s):
        return pd.to_datetime(s, errors="coerce", format="%Y-%m-%d %H:%M:%S")
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", s):
        return pd.to_datetime(s, errors="coerce", format="%Y-%m-%d")
    if re.fullmatch(r"\d{2}/\d{2}/\d{4} \d{2}:\d{2}:\d{2}", s):
        return pd.to_datetime(s, errors="coerce", format="%d/%m/%Y %H:%M:%S")
    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", s):
        return pd.to_datetime(s, errors="coerce", format="%d/%m/%Y")
    if re.fullmatch(r"\d{2}/\d{4}", s):
        return pd.to_datetime("01/" + s, errors="coerce", format="%d/%m/%Y")

    return pd.to_datetime(s, errors="coerce", dayfirst=True)


def mes_para_ordem(valor):
    if pd.isna(valor):
        return None
    texto_original = str(valor).strip()
    if texto_original == "":
        return None

    dt = parse_data_mista(texto_original)
    if pd.notna(dt):
        return int(dt.month)

    texto = normalizar_valor(texto_original)
    texto_sem_ponto = texto.replace(".", "")
    if texto_sem_ponto in MES_NORMALIZADO:
        return MES_NORMALIZADO[texto_sem_ponto]

    match = re.search(r"\b(0?[1-9]|1[0-2])\b", texto_sem_ponto)
    if match:
        return int(match.group(1))

    return None


def ano_para_texto(valor):
    """Extrai ano em 4 digitos de colunas como ANO, MES_ANO, Jan/2025 ou 01/2025."""
    if pd.isna(valor):
        return None

    texto_original = str(valor).strip()
    if texto_original == "":
        return None

    dt = parse_data_mista(texto_original)
    if pd.notna(dt):
        return str(int(dt.year))

    texto = normalizar_valor(texto_original)
    match4 = re.search(r"(19\d{2}|20\d{2})", texto)
    if match4:
        return match4.group(1)

    match2 = re.search(r"(?:^|[/_\-\s])(\d{2})(?:$|\D)", texto)
    if match2:
        yy = int(match2.group(1))
        return str(2000 + yy if yy <= 49 else 1900 + yy)

    return None


def titulo_localidade(uf, municipio):
    if municipio:
        return f"{municipio} - {uf}"
    return f"Todos os municipios - {uf}"


@st.cache_data(ttl=86400, show_spinner=False)
def buscar_municipios_ibge(uf_sigla):
    url = f"https://servicodados.ibge.gov.br/api/v1/localidades/estados/{uf_sigla}/municipios"
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20).json()
        nomes = [m.get("nome", "") for m in res if m.get("nome")]
        return sorted(nomes)
    except Exception:
        return []


def identificar_e_renomear_colunas(df):
    df = df.copy()
    df.columns = [normalizar_coluna(c) for c in df.columns]

    # Se a planilha vier em formato largo, com uma coluna para cada mes,
    # transformamos para formato longo antes de aplicar o restante do tratamento.
    colunas_mensais = [c for c in df.columns if c in COLUNAS_MESES_WIDE]
    tem_coluna_mes = any(c in df.columns for c in COLUNAS_CANONICAS["MES"])
    tem_coluna_ocorrencias = any(c in df.columns for c in COLUNAS_CANONICAS["OCORRENCIAS"])

    if colunas_mensais and not tem_coluna_mes and not tem_coluna_ocorrencias:
        id_vars = [c for c in df.columns if c not in colunas_mensais]
        df = df.melt(
            id_vars=id_vars,
            value_vars=colunas_mensais,
            var_name="MES",
            value_name="OCORRENCIAS",
        )
        df["MES"] = df["MES"].map(lambda c: MES_NOME_LONGO.get(COLUNAS_MESES_WIDE.get(c), c))

    renomear = {}
    colunas_existentes = set(df.columns)

    for canonica, candidatas in COLUNAS_CANONICAS.items():
        if canonica in colunas_existentes:
            continue
        for candidata in candidatas:
            candidata_norm = normalizar_coluna(candidata)
            if candidata_norm in colunas_existentes:
                renomear[candidata_norm] = canonica
                break

    if renomear:
        df = df.rename(columns=renomear)

    return df


def padronizar_base_seguranca(df, fonte):
    if df is None or df.empty:
        return pd.DataFrame()

    df = identificar_e_renomear_colunas(df)

    # Campos basicos em texto.
    for col in ["UF", "MUNICIPIO", "TIPO_CRIME", "MES", "ANO", "REGIAO"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    if "UF" in df.columns:
        df["UF"] = df["UF"].map(lambda x: normalizar_valor(x)[:2] if len(normalizar_valor(x)) == 2 else normalizar_valor(x))
        # Corrige casos em que a coluna veio com nome por extenso em vez de sigla.
        nome_para_sigla = {normalizar_valor(v): k for k, v in UF_NOMES.items()}
        df["UF"] = df["UF"].map(lambda x: nome_para_sigla.get(x, x))

    if "MUNICIPIO" in df.columns:
        df["MUNICIPIO"] = df["MUNICIPIO"].map(lambda x: str(x).strip().title())

    # Algumas planilhas oficiais usam MES_ANO em vez de MES/ANO separados.
    if "MES" not in df.columns and "MES_ANO" in df.columns:
        df["MES"] = df["MES_ANO"]

    if "TIPO_CRIME" not in df.columns and "CRIME" in df.columns:
        df["TIPO_CRIME"] = df["CRIME"]

    if "CRIME" in df.columns and fonte == "api_rayonnunes":
        df["TIPO_CRIME"] = df["CRIME"].astype(str).map(REVERSE_CRIMES_API).fillna(df.get("TIPO_CRIME", "Outros"))

    if "TIPO_CRIME" not in df.columns and "ABA_ORIGEM" in df.columns:
        # Quando cada aba representa um indicador/crime, usa-se o nome da aba.
        df["TIPO_CRIME"] = df["ABA_ORIGEM"]

    if "TIPO_CRIME" in df.columns:
        df["TIPO_CRIME"] = df["TIPO_CRIME"].astype(str).str.strip()
        df.loc[df["TIPO_CRIME"].isin(["", "nan", "None"]), "TIPO_CRIME"] = "Nao informado"
        if "ABA_ORIGEM" in df.columns:
            mascara_na = df["TIPO_CRIME"].map(lambda x: chave_comparacao(x) in ["", "NAOINFORMADO", "NAN", "NONE"])
            df.loc[mascara_na, "TIPO_CRIME"] = df.loc[mascara_na, "ABA_ORIGEM"].astype(str).str.strip()
    else:
        df["TIPO_CRIME"] = "Nao informado"

    if "ANO" in df.columns:
        ano_extraido = df["ANO"].map(ano_para_texto)
        if "MES_ANO" in df.columns:
            ano_extraido = ano_extraido.fillna(df["MES_ANO"].map(ano_para_texto))
        df["ANO"] = ano_extraido.fillna("Nao informado")
    elif "MES_ANO" in df.columns:
        df["ANO"] = df["MES_ANO"].map(ano_para_texto).fillna("Nao informado")
    else:
        df["ANO"] = "Nao informado"

    if "MES" in df.columns:
        df["MES_ORDEM"] = df["MES"].map(mes_para_ordem)
        if "MES_ANO" in df.columns:
            df["MES_ORDEM"] = df["MES_ORDEM"].fillna(df["MES_ANO"].map(mes_para_ordem))
        df["MES_NOME"] = df["MES_ORDEM"].map(MES_LABEL).fillna(df["MES"].astype(str))
    elif "MES_ANO" in df.columns:
        df["MES_ORDEM"] = df["MES_ANO"].map(mes_para_ordem)
        df["MES_NOME"] = df["MES_ORDEM"].map(MES_LABEL).fillna(df["MES_ANO"].astype(str))
    else:
        df["MES_ORDEM"] = None
        df["MES_NOME"] = "Nao informado"

    # Numericos.
    if "OCORRENCIAS" in df.columns:
        df["OCORRENCIAS"] = df["OCORRENCIAS"].map(converter_numero)
    else:
        df["OCORRENCIAS"] = 0

    if "VITIMAS" in df.columns:
        df["VITIMAS"] = df["VITIMAS"].map(converter_numero)
    else:
        df["VITIMAS"] = 0

    if "VITIMAS_MUNICIPIO" in df.columns:
        df["VITIMAS_MUNICIPIO"] = df["VITIMAS_MUNICIPIO"].map(converter_numero)
    else:
        df["VITIMAS_MUNICIPIO"] = 0

    # A base oficial municipal do MJSP/SINESP e documentada como base por
    # municipio com unidade principal em vitimas. Portanto, nao se deve
    # interpretar ausencia de OCORRENCIAS como zero ocorrencias criminais.
    # Para o painel municipal, a metrica principal sera VITIMAS.
    if fonte == "oficial_mjsp_municipios":
        try:
            if df["VITIMAS_MUNICIPIO"].sum() == 0 and df["VITIMAS"].sum() > 0:
                df["VITIMAS_MUNICIPIO"] = df["VITIMAS"]
        except Exception:
            pass

    df["FONTE_PROCESSADA"] = fonte
    df = df.loc[:, ~df.columns.duplicated()].copy()
    return df


def detectar_linha_cabecalho_excel(conteudo_bytes, aba):
    """Tenta localizar o cabecalho real quando a planilha vem com linhas de titulo/notas."""
    try:
        preview = pd.read_excel(
            io.BytesIO(conteudo_bytes),
            sheet_name=aba,
            header=None,
            nrows=15,
            dtype=str,
            engine="openpyxl",
        )
    except Exception:
        return 0

    palavras_chave = [
        "UF", "MUNICIPIO", "CODIGO", "REGIAO", "ANO", "MES",
        "CRIME", "INDICADOR", "OCORRENCIA", "VITIMA",
    ]

    melhor_linha = 0
    melhor_pontos = -1
    for idx, row in preview.iterrows():
        valores = [normalizar_coluna(v) for v in row.dropna().tolist()]
        if not valores:
            continue
        pontos = sum(
            1 for palavra in palavras_chave
            if any(palavra in valor for valor in valores)
        )
        if pontos > melhor_pontos:
            melhor_pontos = pontos
            melhor_linha = int(idx)

    return melhor_linha if melhor_pontos >= 2 else 0


def ler_excel_oficial_multiplas_abas(conteudo_bytes):
    """Le todas as abas do XLSX oficial e concatena em uma unica base.

    Isso evita o bug de carregar apenas a primeira aba da planilha, que pode fazer
    a amostra mostrar outra UF/recorte enquanto o usuario selecionou BH, RJ, SP etc.
    """
    xls = pd.ExcelFile(io.BytesIO(conteudo_bytes), engine="openpyxl")
    frames = []
    metadados_abas = []

    for aba in xls.sheet_names:
        try:
            header_row = detectar_linha_cabecalho_excel(conteudo_bytes, aba)
            df_aba = pd.read_excel(
                io.BytesIO(conteudo_bytes),
                sheet_name=aba,
                header=header_row,
                dtype=str,
                engine="openpyxl",
            )
            df_aba = df_aba.dropna(how="all")
            df_aba = df_aba.loc[:, ~df_aba.columns.astype(str).str.startswith("Unnamed")]
            if df_aba.empty:
                metadados_abas.append({"aba": aba, "linhas": 0, "status": "vazia"})
                continue
            df_aba["ABA_ORIGEM"] = aba
            frames.append(df_aba)
            metadados_abas.append({"aba": aba, "linhas": len(df_aba), "cabecalho_linha": header_row})
        except Exception as e:
            metadados_abas.append({"aba": aba, "linhas": 0, "erro": str(e)})

    if not frames:
        return pd.DataFrame(), metadados_abas

    return pd.concat(frames, ignore_index=True, sort=False), metadados_abas


def adicionar_colunas_filtro(df):
    """Cria colunas auxiliares para filtro SQL no DuckDB.

    Essas colunas evitam aplicar normalizacao pesada em pandas a cada consulta.
    O filtro fica barato: DuckDB le o Parquet e materializa somente o recorte.
    """
    df = df.copy()
    if "UF" in df.columns:
        df["UF_FILTRO"] = df["UF"].astype(str).str.upper().str.strip()
    else:
        df["UF_FILTRO"] = ""

    if "MUNICIPIO" in df.columns:
        df["MUNICIPIO_FILTRO"] = df["MUNICIPIO"].map(chave_comparacao)
    else:
        df["MUNICIPIO_FILTRO"] = ""

    if "TIPO_CRIME" in df.columns:
        df["TIPO_CRIME_FILTRO"] = df["TIPO_CRIME"].map(chave_comparacao)
    else:
        df["TIPO_CRIME_FILTRO"] = ""

    if "ANO" in df.columns:
        df["ANO_FILTRO"] = df["ANO"].astype(str).str.strip()
    else:
        df["ANO_FILTRO"] = ""

    if "MES_ORDEM" in df.columns:
        df["MES_ORDEM_NUM"] = pd.to_numeric(df["MES_ORDEM"], errors="coerce")
    else:
        df["MES_ORDEM_NUM"] = pd.NA

    return df


def caminho_parquet_cache(tipo_base):
    pasta = os.path.join(tempfile.gettempdir(), "sinesp_mjsp_cache_duckdb")
    os.makedirs(pasta, exist_ok=True)
    return os.path.join(pasta, f"mjsp_sinesp_{tipo_base}.parquet")


@st.cache_data(ttl=86400, show_spinner=False)
def preparar_base_oficial_parquet(tipo_base):
    """Baixa a base oficial, padroniza uma vez e grava em Parquet.

    A operacao pesada continua existindo na primeira carga do dia, porque o MJSP
    publica XLSX e XLSX nao permite filtro remoto por UF/municipio. Depois disso,
    cada consulta usa DuckDB diretamente sobre Parquet, sem materializar a base inteira.
    """
    url = URL_MJSP_MUNICIPIOS if tipo_base == "municipios" else URL_MJSP_UF
    parquet_path = caminho_parquet_cache(tipo_base)

    try:
        resp = requests.get(
            url,
            timeout=180,
            verify=False,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        if resp.status_code != 200:
            return {
                "ok": False,
                "erro": f"HTTP {resp.status_code}",
                "url": url,
                "texto": resp.text[:1500],
                "parquet_path": parquet_path,
            }

        df_raw, meta_abas = ler_excel_oficial_multiplas_abas(resp.content)
        if df_raw.empty:
            return {
                "ok": False,
                "erro": "XLSX carregado, mas nenhuma aba util foi identificada.",
                "url": url,
                "abas": meta_abas,
                "texto": "",
                "parquet_path": parquet_path,
            }

        df = padronizar_base_seguranca(df_raw, f"oficial_mjsp_{tipo_base}")
        df = adicionar_colunas_filtro(df)
        df.to_parquet(parquet_path, index=False)

        colunas = list(df.columns)
        return {
            "ok": True,
            "url": url,
            "linhas_brutas": int(len(df_raw)),
            "linhas": int(len(df)),
            "abas": meta_abas,
            "colunas": colunas,
            "parquet_path": parquet_path,
            "motor_filtro": "DuckDB sobre Parquet local em cache",
        }

    except Exception as e:
        return {
            "ok": False,
            "erro": str(e),
            "url": url,
            "texto": "",
            "parquet_path": parquet_path,
        }


def construir_where_duckdb(uf, municipio, crime, ano, mes_num):
    clauses = []
    params = []

    if uf:
        clauses.append("UF_FILTRO = ?")
        params.append(uf.upper())

    if municipio:
        chave_mun = chave_comparacao(municipio)
        clauses.append("(MUNICIPIO_FILTRO = ? OR MUNICIPIO_FILTRO LIKE ?)")
        params.extend([chave_mun, f"%{chave_mun}%"])

    if crime and crime != "Todos os indicadores":
        chave_crime = chave_comparacao(crime)
        clauses.append("TIPO_CRIME_FILTRO = ?")
        params.append(chave_crime)

    if ano and ano != "Todos os anos":
        clauses.append("ANO_FILTRO = ?")
        params.append(str(ano))

    if mes_num:
        clauses.append("CAST(MES_ORDEM_NUM AS INTEGER) = ?")
        params.append(int(mes_num))

    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    return where, params


def query_base_oficial_duckdb(parquet_path, uf, municipio, crime, ano, mes_num, limit=None):
    if not parquet_path or not os.path.exists(parquet_path):
        return pd.DataFrame()

    path_sql = parquet_path.replace("'", "''")
    where, params = construir_where_duckdb(uf, municipio, crime, ano, mes_num)
    limite_sql = f" LIMIT {int(limit)}" if limit else ""
    query = f"SELECT * FROM read_parquet('{path_sql}'){where}{limite_sql}"

    con = duckdb.connect(database=":memory:")
    try:
        return con.execute(query, params).fetchdf()
    finally:
        con.close()


def contar_duckdb(parquet_path, where="", params=None):
    if params is None:
        params = []
    if not parquet_path or not os.path.exists(parquet_path):
        return 0
    path_sql = parquet_path.replace("'", "''")
    con = duckdb.connect(database=":memory:")
    try:
        return int(con.execute(f"SELECT COUNT(*) FROM read_parquet('{path_sql}'){where}", params).fetchone()[0])
    finally:
        con.close()


def diagnosticar_etapas_filtro_duckdb(parquet_path, uf, municipio, crime, ano, mes_num):
    etapas = []
    etapas.append({"etapa": "Base completa carregada em Parquet", "linhas": contar_duckdb(parquet_path)})

    filtros = []
    params = []

    if uf:
        filtros.append("UF_FILTRO = ?")
        params.append(uf.upper())
        etapas.append({"etapa": f"Apos filtro UF = {uf}", "linhas": contar_duckdb(parquet_path, " WHERE " + " AND ".join(filtros), params)})

    if municipio:
        chave_mun = chave_comparacao(municipio)
        filtros.append("(MUNICIPIO_FILTRO = ? OR MUNICIPIO_FILTRO LIKE ?)")
        params.extend([chave_mun, f"%{chave_mun}%"])
        etapas.append({"etapa": f"Apos filtro municipio = {municipio}", "linhas": contar_duckdb(parquet_path, " WHERE " + " AND ".join(filtros), params)})

    if crime and crime != "Todos os indicadores":
        filtros.append("TIPO_CRIME_FILTRO = ?")
        params.append(chave_comparacao(crime))
        etapas.append({"etapa": f"Apos filtro indicador = {crime}", "linhas": contar_duckdb(parquet_path, " WHERE " + " AND ".join(filtros), params)})

    if ano and ano != "Todos os anos":
        filtros.append("ANO_FILTRO = ?")
        params.append(str(ano))
        etapas.append({"etapa": f"Apos filtro ano = {ano}", "linhas": contar_duckdb(parquet_path, " WHERE " + " AND ".join(filtros), params)})

    if mes_num:
        filtros.append("CAST(MES_ORDEM_NUM AS INTEGER) = ?")
        params.append(int(mes_num))
        etapas.append({"etapa": f"Apos filtro mes = {mes_num}", "linhas": contar_duckdb(parquet_path, " WHERE " + " AND ".join(filtros), params)})

    amostra = query_base_oficial_duckdb(parquet_path, uf, municipio, crime, ano, mes_num, limit=30)

    # contexto UF/municipio para listar anos e indicadores existentes sem carregar tudo
    contexto_where, contexto_params = construir_where_duckdb(uf, municipio, "Todos os indicadores", "Todos os anos", None)
    path_sql = parquet_path.replace("'", "''")
    con = duckdb.connect(database=":memory:")
    try:
        anos = con.execute(
            f"SELECT DISTINCT ANO_FILTRO FROM read_parquet('{path_sql}'){contexto_where} WHERE ANO_FILTRO IS NOT NULL" if not contexto_where else f"SELECT DISTINCT ANO_FILTRO FROM read_parquet('{path_sql}'){contexto_where}",
            contexto_params,
        ).fetchdf()
        crimes = con.execute(
            f"SELECT DISTINCT TIPO_CRIME FROM read_parquet('{path_sql}'){contexto_where}",
            contexto_params,
        ).fetchdf()
    except Exception:
        anos = pd.DataFrame()
        crimes = pd.DataFrame()
    finally:
        con.close()

    return {
        "etapas": etapas,
        "amostra_contexto": amostra,
        "anos_disponiveis_no_contexto": sorted([str(x) for x in anos.get("ANO_FILTRO", pd.Series(dtype=str)).dropna().tolist()], reverse=True)[:30],
        "indicadores_disponiveis_no_contexto": sorted([str(x) for x in crimes.get("TIPO_CRIME", pd.Series(dtype=str)).dropna().tolist()])[:50],
    }


def opcoes_anos_padrao():
    ano_atual = datetime.now().year
    return ["Todos os anos"] + [str(a) for a in range(ano_atual, 2014, -1)]


def opcoes_crimes_padrao(usa_oficial_municipios, usa_oficial_uf):
    if usa_oficial_municipios:
        return ["Todos os indicadores", "Homicidio doloso"]
    if usa_oficial_uf:
        return [
            "Todos os indicadores",
            "Estupro",
            "Furto de veiculo",
            "Homicidio doloso",
            "Lesao corporal seguida de morte",
            "Roubo a instituicao financeira",
            "Roubo de carga",
            "Roubo de veiculo",
            "Roubo seguido de morte (latrocinio)",
            "Tentativa de homicidio",
        ]
    return ["Todos os indicadores"]



@st.cache_data(ttl=3600, show_spinner=False)
def consultar_api_rayonnunes(uf, municipio, crime_id, ano, mes_num):
    """
    Consulta a API comunitaria rayonnunes com estrategia defensiva.

    Motivo: essa API antiga frequentemente da timeout quando per_page=1000,
    principalmente em municipios grandes ou consultas sem filtro de crime/mes.
    A funcao tenta paginas menores automaticamente e registra diagnostico.
    """
    params_base = {
        "uf": uf.lower(),
    }

    if municipio:
        params_base["municipio"] = normalizar_municipio_api(municipio)
    if crime_id:
        params_base["crime"] = crime_id
    if ano and ano != "Todos os anos":
        params_base["ano"] = str(ano)
    if mes_num:
        params_base["mes"] = MES_API.get(int(mes_num))

    frames = []
    chamadas = []
    erros = []

    # A documentacao da API aceita ate 1000 por pagina, mas 1000 costuma
    # estourar timeout. Comecamos menor para aumentar a chance de resposta.
    # Se ainda falhar, reduzimos novamente.
    per_page_tentativas = [200, 100, 50]
    timeout_por_chamada = 75
    max_paginas = 200
    houve_resposta_valida = False

    for per_page in per_page_tentativas:
        frames_tentativa = []
        erros_tentativa = []
        chamadas_tentativa = []

        for pagina in range(1, max_paginas + 1):
            params = dict(params_base)
            params["per_page"] = per_page
            params["page"] = pagina

            try:
                resp = requests.get(
                    API_RAYONNUNES,
                    params=params,
                    timeout=timeout_por_chamada,
                    headers={"User-Agent": "Mozilla/5.0"},
                )
                chamadas_tentativa.append(
                    {
                        "pagina": pagina,
                        "per_page": per_page,
                        "url": resp.url,
                        "status": resp.status_code,
                    }
                )

                if resp.status_code != 200:
                    erros_tentativa.append(
                        f"per_page={per_page}, pagina {pagina}: HTTP {resp.status_code} - {resp.text[:500]}"
                    )
                    break

                try:
                    payload = resp.json()
                except Exception:
                    erros_tentativa.append(
                        f"per_page={per_page}, pagina {pagina}: resposta nao veio em JSON - {resp.text[:500]}"
                    )
                    break

                dados = payload.get("data", []) if isinstance(payload, dict) else []
                houve_resposta_valida = True

                if not dados:
                    # Fim normal da paginacao.
                    break

                frames_tentativa.append(pd.DataFrame(dados))

                # Se vier menos do que o limite, acabou.
                if len(dados) < per_page:
                    break

            except requests.exceptions.Timeout:
                erros_tentativa.append(
                    f"per_page={per_page}, pagina {pagina}: tempo limite excedido."
                )
                break
            except requests.exceptions.RequestException as e:
                erros_tentativa.append(
                    f"per_page={per_page}, pagina {pagina}: falha de conexao - {e}"
                )
                break
            except Exception as e:
                erros_tentativa.append(
                    f"per_page={per_page}, pagina {pagina}: erro inesperado - {e}"
                )
                break

        chamadas.extend(chamadas_tentativa)
        erros.extend(erros_tentativa)

        if frames_tentativa:
            frames = frames_tentativa
            erros.append(f"Consulta concluida usando per_page={per_page}.")
            break

        # Se a API respondeu sem dados, nao adianta tentar outro per_page.
        # Diferente de timeout: resposta valida sem registros.
        if houve_resposta_valida and not frames_tentativa and not erros_tentativa:
            break

    if frames:
        df = pd.concat(frames, ignore_index=True)
        df = padronizar_base_seguranca(df, "api_rayonnunes")
    else:
        df = pd.DataFrame()

    meta = {
        "chamadas": chamadas,
        "erros": erros,
        "params": params_base,
        "observacao": (
            "A API comunitaria pode ficar indisponivel ou exceder tempo de resposta. "
            "Para producao, prefira a fonte oficial MJSP/SINESP em XLSX."
        ),
    }
    return df, meta

def filtrar_base(df, uf, municipio, crime, ano, mes_num):
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "UF" in out.columns and uf:
        out = out[out["UF"].astype(str).str.upper() == uf.upper()].copy()

    if municipio and "MUNICIPIO" in out.columns:
        chave_mun = chave_comparacao(municipio)
        chaves_base = out["MUNICIPIO"].map(chave_comparacao)
        mascara_mun = chaves_base == chave_mun
        # Fallback para casos como "Belo Horizonte/MG" ou "Belo Horizonte - MG".
        if not mascara_mun.any():
            mascara_mun = chaves_base.str.contains(chave_mun, na=False)
        out = out[mascara_mun].copy()

    if crime and crime != "Todos os indicadores" and "TIPO_CRIME" in out.columns:
        chave_crime = chave_comparacao(crime)
        out = out[out["TIPO_CRIME"].map(chave_comparacao) == chave_crime].copy()

    if ano and ano != "Todos os anos" and "ANO" in out.columns:
        out = out[out["ANO"].astype(str) == str(ano)].copy()

    if mes_num and "MES_ORDEM" in out.columns:
        out = out[pd.to_numeric(out["MES_ORDEM"], errors="coerce") == int(mes_num)].copy()

    return out


def obter_opcoes_anos(df):
    ano_atual = datetime.now().year
    fallback = [str(a) for a in range(ano_atual, 2014, -1)]
    if df is None or df.empty or "ANO" not in df.columns:
        return ["Todos os anos"] + fallback
    anos = sorted(
        {str(a) for a in df["ANO"].dropna().astype(str).tolist() if re.fullmatch(r"\d{4}", str(a))},
        reverse=True,
    )
    return ["Todos os anos"] + (anos if anos else fallback)


def obter_opcoes_crimes(df):
    if df is None or df.empty or "TIPO_CRIME" not in df.columns:
        return ["Todos os indicadores"]
    crimes = sorted(
        {
            str(x).strip()
            for x in df["TIPO_CRIME"].dropna().astype(str).tolist()
            if str(x).strip() and str(x).strip().lower() not in ["nan", "none"]
        }
    )
    return ["Todos os indicadores"] + crimes


def obter_opcoes_municipios(df, uf):
    if df is not None and not df.empty and "MUNICIPIO" in df.columns and "UF" in df.columns:
        temp = df[df["UF"].astype(str).str.upper() == uf.upper()].copy()
        municipios = sorted(
            {
                str(x).strip().title()
                for x in temp["MUNICIPIO"].dropna().astype(str).tolist()
                if str(x).strip() and str(x).strip().lower() not in ["nan", "none"]
            }
        )
        if municipios:
            return municipios
    return buscar_municipios_ibge(uf)


def diagnosticar_etapas_filtro(df, uf, municipio, crime, ano, mes_num):
    """Retorna contagens apos cada filtro para explicar por que uma consulta ficou vazia."""
    etapas = []
    amostra_contexto = pd.DataFrame()

    if df is None or df.empty:
        return {"etapas": [{"etapa": "Base completa", "linhas": 0}], "amostra_contexto": amostra_contexto}

    atual = df.copy()
    etapas.append({"etapa": "Base completa carregada", "linhas": len(atual)})

    if "UF" in atual.columns and uf:
        atual = atual[atual["UF"].astype(str).str.upper() == uf.upper()].copy()
        etapas.append({"etapa": f"Apos filtro UF = {uf}", "linhas": len(atual)})
        amostra_contexto = atual.head(30)

    if municipio and "MUNICIPIO" in atual.columns:
        chave_mun = chave_comparacao(municipio)
        chaves_base = atual["MUNICIPIO"].map(chave_comparacao)
        mascara_mun = chaves_base == chave_mun
        if not mascara_mun.any():
            mascara_mun = chaves_base.str.contains(chave_mun, na=False)
        atual = atual[mascara_mun].copy()
        etapas.append({"etapa": f"Apos filtro municipio = {municipio}", "linhas": len(atual)})
        amostra_contexto = atual.head(30)

    if crime and crime != "Todos os indicadores" and "TIPO_CRIME" in atual.columns:
        chave_crime = chave_comparacao(crime)
        atual = atual[atual["TIPO_CRIME"].map(chave_comparacao) == chave_crime].copy()
        etapas.append({"etapa": f"Apos filtro indicador = {crime}", "linhas": len(atual)})
        amostra_contexto = atual.head(30)

    if ano and ano != "Todos os anos" and "ANO" in atual.columns:
        atual = atual[atual["ANO"].astype(str) == str(ano)].copy()
        etapas.append({"etapa": f"Apos filtro ano = {ano}", "linhas": len(atual)})
        amostra_contexto = atual.head(30)

    if mes_num and "MES_ORDEM" in atual.columns:
        atual = atual[pd.to_numeric(atual["MES_ORDEM"], errors="coerce") == int(mes_num)].copy()
        etapas.append({"etapa": f"Apos filtro mes = {mes_num}", "linhas": len(atual)})
        amostra_contexto = atual.head(30)

    resumo = {"etapas": etapas, "amostra_contexto": amostra_contexto}

    # Opcoes disponiveis no contexto filtrado por UF/municipio, quando possivel.
    contexto = df.copy()
    if "UF" in contexto.columns and uf:
        contexto = contexto[contexto["UF"].astype(str).str.upper() == uf.upper()].copy()
    if municipio and "MUNICIPIO" in contexto.columns:
        chave_mun = chave_comparacao(municipio)
        chaves_base = contexto["MUNICIPIO"].map(chave_comparacao)
        mascara_mun = chaves_base == chave_mun
        if not mascara_mun.any():
            mascara_mun = chaves_base.str.contains(chave_mun, na=False)
        contexto = contexto[mascara_mun].copy()

    resumo["anos_disponiveis_no_contexto"] = sorted(
        {str(x) for x in contexto.get("ANO", pd.Series(dtype=str)).dropna().astype(str).tolist()},
        reverse=True,
    )[:30]
    resumo["indicadores_disponiveis_no_contexto"] = sorted(
        {str(x) for x in contexto.get("TIPO_CRIME", pd.Series(dtype=str)).dropna().astype(str).tolist()}
    )[:50]

    return resumo


def mostrar_alerta_colunas(df, fonte_label):
    colunas_necessarias = ["UF", "ANO", "MES_ORDEM", "TIPO_CRIME", "OCORRENCIAS"]
    faltantes = [c for c in colunas_necessarias if c not in df.columns]
    if faltantes:
        st.warning(
            f"A base {fonte_label} foi carregada, mas faltam colunas padronizadas: {', '.join(faltantes)}. "
            "Os filtros/graficos podem ficar limitados. Veja a aba de diagnostico."
        )


def obter_metrica_principal(df, usa_base_municipal=False):
    """Define a metrica principal respeitando a metodologia da fonte.

    - Base oficial municipal: usar VITIMAS como metrica principal.
    - Base oficial UF e API comunitaria: usar OCORRENCIAS quando existir.
    """
    if df is None or df.empty:
        return "OCORRENCIAS", "Ocorrencias"

    def coluna_tem_soma(coluna):
        if coluna not in df.columns:
            return False
        return pd.to_numeric(df[coluna], errors="coerce").fillna(0).sum() > 0

    if usa_base_municipal:
        if coluna_tem_soma("VITIMAS"):
            return "VITIMAS", "Vitimas"
        if coluna_tem_soma("VITIMAS_MUNICIPIO"):
            return "VITIMAS_MUNICIPIO", "Vitimas municipio"
        return "VITIMAS", "Vitimas"

    if coluna_tem_soma("OCORRENCIAS"):
        return "OCORRENCIAS", "Ocorrencias"
    if coluna_tem_soma("VITIMAS"):
        return "VITIMAS", "Vitimas"
    if coluna_tem_soma("VITIMAS_MUNICIPIO"):
        return "VITIMAS_MUNICIPIO", "Vitimas municipio"
    return "OCORRENCIAS", "Ocorrencias"


def soma_segura(df, coluna):
    if df is None or df.empty or coluna not in df.columns:
        return 0
    return pd.to_numeric(df[coluna], errors="coerce").fillna(0).sum()


def plotar_serie_mensal(df, metrica_coluna="OCORRENCIAS", metrica_label="Ocorrencias"):
    if df.empty or "MES_ORDEM" not in df.columns or metrica_coluna not in df.columns:
        st.info(f"Nao ha dados mensais suficientes para montar a serie temporal de {metrica_label.lower()}.")
        return
    base = df.copy()
    base["MES_ORDEM"] = pd.to_numeric(base["MES_ORDEM"], errors="coerce")
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    base = base.dropna(subset=["MES_ORDEM"])
    if base.empty:
        st.info(f"Nao ha dados mensais validos para montar a serie temporal de {metrica_label.lower()}.")
        return
    serie = (
        base.groupby("MES_ORDEM", dropna=True)[metrica_coluna]
        .sum()
        .reset_index()
        .sort_values("MES_ORDEM")
    )
    serie["MES"] = serie["MES_ORDEM"].astype(int).map(MES_LABEL)
    serie = serie.set_index("MES")[metrica_coluna]
    if serie.empty or serie.sum() == 0:
        st.info(f"Nao ha valores de {metrica_label.lower()} para a serie temporal.")
        return
    st.line_chart(serie)


def plotar_crimes(df, metrica_coluna="OCORRENCIAS", metrica_label="Ocorrencias"):
    if df.empty or "TIPO_CRIME" not in df.columns or metrica_coluna not in df.columns:
        st.info("Nao ha dados suficientes para a distribuicao por indicador/tipo de crime.")
        return
    base = df.copy()
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    serie = base.groupby("TIPO_CRIME")[metrica_coluna].sum().sort_values(ascending=False).head(15)
    if serie.empty or serie.sum() == 0:
        st.info(f"Nao ha valores de {metrica_label.lower()} para a distribuicao por tipo de crime.")
        return
    st.bar_chart(serie)


def plotar_top_municipios(df, metrica_coluna="OCORRENCIAS", metrica_label="Ocorrencias"):
    if df.empty or "MUNICIPIO" not in df.columns or metrica_coluna not in df.columns:
        st.info("Selecione a base municipal para ver o ranking por municipio.")
        return
    base = df.copy()
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    serie = base.groupby("MUNICIPIO")[metrica_coluna].sum().sort_values(ascending=False).head(15)
    if serie.empty or serie.sum() == 0:
        st.info(f"Nao ha valores de {metrica_label.lower()} para montar o ranking municipal.")
        return
    st.bar_chart(serie)


# -----------------------------------------------------------------------------
# INTERFACE
# -----------------------------------------------------------------------------
st.sidebar.title("🛡️ Filtros de Seguranca")

fonte_dados = st.sidebar.radio(
    "Fonte dos dados:",
    [
        "Oficial MJSP/SINESP - Municipios XLSX",
        "Oficial MJSP/SINESP - UF XLSX",
        "API comunitaria rayonnunes - experimental",
    ],
    index=0,
)

usa_oficial_municipios = fonte_dados.startswith("Oficial") and "Municipios" in fonte_dados
usa_oficial_uf = fonte_dados.startswith("Oficial") and "UF" in fonte_dados
usa_api = fonte_dados.startswith("API")

uf_sel = st.sidebar.selectbox("Selecione o Estado:", sorted(UFS), index=sorted(UFS).index("MG"))

municipio_sel = "Todos os municipios"
municipio_param = None
if usa_oficial_municipios or usa_api:
    # Municipio vem do IBGE, sem carregar a base oficial inteira no sidebar.
    municipios = buscar_municipios_ibge(uf_sel)
    municipio_sel = st.sidebar.selectbox(
        "Selecione o Municipio:",
        ["Todos os municipios"] + municipios,
    )
    municipio_param = None if municipio_sel == "Todos os municipios" else municipio_sel
else:
    st.sidebar.info("A base por UF nao possui recorte municipal.")

mes_nome = st.sidebar.selectbox("Mes:", list(MESES_DISPLAY.keys()), index=0)
mes_param = MESES_DISPLAY[mes_nome]

if usa_api:
    crime_nome = st.sidebar.selectbox("Classificacao do crime:", list(MAPA_CRIMES_API.keys()))
    crime_param = MAPA_CRIMES_API[crime_nome]
    ano_sel = st.sidebar.selectbox("Ano de Referencia:", opcoes_anos_padrao(), index=1)
else:
    crime_opcoes = opcoes_crimes_padrao(usa_oficial_municipios, usa_oficial_uf)
    if usa_oficial_municipios:
        st.sidebar.caption(
            "Na base oficial municipal, a unidade principal e vitimas. "
            "O dicionario municipal indica foco em homicidio doloso/vitimas."
        )
    crime_nome = st.sidebar.selectbox("Indicador / Tipo de Crime:", crime_opcoes)
    crime_param = crime_nome
    ano_sel = st.sidebar.selectbox("Ano de Referencia:", opcoes_anos_padrao(), index=1)

with st.sidebar.form("form_seguranca"):
    submit_btn = st.form_submit_button("🔍 Consultar Indicadores")

# -----------------------------------------------------------------------------
# PROCESSAMENTO E SAIDA
# -----------------------------------------------------------------------------
if not submit_btn:
    st.info("💡 Escolha a localidade, fonte e indicador na barra lateral e clique em Consultar Indicadores.")
    st.markdown(
        """
        **Recomendacao de uso:** deixe a fonte oficial MJSP/SINESP como padrao para producao.  
        Nesta versao, o app nao carrega a base oficial inteira no sidebar: ele prepara um Parquet em cache e usa DuckDB para filtrar antes de exibir.
        """
    )
    st.stop()

st.markdown(
    '<div class="header-seguranca"><h1>Painel de Ocorrencias Criminais - SINESP/MJSP</h1>'
    '<p>Monitoramento territorial de indicadores de seguranca publica</p></div>',
    unsafe_allow_html=True,
)

meta_api = {}
meta_oficial = {}
parquet_path = ""
base_tipo = "municipios" if usa_oficial_municipios else "uf"

if usa_api:
    with st.spinner("Consultando API comunitaria rayonnunes..."):
        df_consulta, meta_api = consultar_api_rayonnunes(
            uf=uf_sel,
            municipio=municipio_param,
            crime_id=crime_param,
            ano=ano_sel,
            mes_num=mes_param,
        )
else:
    with st.spinner("Preparando cache oficial MJSP/SINESP e filtrando com DuckDB..."):
        meta_oficial = preparar_base_oficial_parquet(base_tipo)
        parquet_path = meta_oficial.get("parquet_path", "")

        if not meta_oficial.get("ok"):
            df_consulta = pd.DataFrame()
        else:
            df_consulta = query_base_oficial_duckdb(
                parquet_path,
                uf=uf_sel,
                municipio=municipio_param,
                crime=crime_param,
                ano=ano_sel,
                mes_num=mes_param,
            )

if df_consulta.empty:
    st.error(
        "🛑 Nao foram encontrados registros para a combinacao selecionada. "
        "Isso pode indicar ausencia de dados, indicador diferente na fonte, ano/mes indisponivel ou falha da fonte experimental."
    )
    with st.expander("Diagnostico tecnico", expanded=True):
        st.write("**Fonte selecionada:**", fonte_dados)
        st.write("**UF:**", uf_sel)
        st.write("**Municipio:**", municipio_param or "Todos")
        st.write("**Indicador/crime:**", crime_nome)
        st.write("**Ano:**", ano_sel)
        st.write("**Mes:**", mes_nome)
        if usa_api:
            st.json(meta_api)
        else:
            st.json(meta_oficial)
            if meta_oficial.get("ok"):
                diag = diagnosticar_etapas_filtro_duckdb(
                    parquet_path,
                    uf=uf_sel,
                    municipio=municipio_param,
                    crime=crime_param,
                    ano=ano_sel,
                    mes_num=mes_param,
                )
                st.write("Contagem por etapa do filtro:")
                st.dataframe(pd.DataFrame(diag["etapas"]), width="stretch")
                st.write("Anos disponiveis para o contexto UF/municipio selecionado:")
                st.write(diag.get("anos_disponiveis_no_contexto", []))
                st.write("Indicadores disponiveis para o contexto UF/municipio selecionado:")
                st.write(diag.get("indicadores_disponiveis_no_contexto", []))
                st.write("Colunas padronizadas carregadas:")
                st.code("\n".join(meta_oficial.get("colunas", [])))
                st.write("Amostra contextual filtrada pelo DuckDB:")
                amostra = diag.get("amostra_contexto", pd.DataFrame())
                if amostra.empty:
                    st.info("Nao ha amostra para o filtro selecionado.")
                else:
                    st.dataframe(amostra, width="stretch")
    st.stop()

# Cards
localidade = titulo_localidade(uf_sel, municipio_param)
metrica_coluna, metrica_label = obter_metrica_principal(
    df_consulta,
    usa_base_municipal=usa_oficial_municipios,
)

valor_principal = soma_segura(df_consulta, metrica_coluna)
total_ocorrencias = soma_segura(df_consulta, "OCORRENCIAS")
total_vitimas = soma_segura(df_consulta, "VITIMAS")
total_vitimas_mun = soma_segura(df_consulta, "VITIMAS_MUNICIPIO")
registros = len(df_consulta)

if usa_oficial_municipios:
    subtitulo_principal = "Unidade principal da base municipal"
    titulo_card_secundario = "Ocorrencias"
    valor_card_secundario = total_ocorrencias
    nota_secundaria = "Pode vir zerado na base municipal"
else:
    subtitulo_principal = "Metrica principal do filtro"
    titulo_card_secundario = "Vitimas"
    valor_card_secundario = total_vitimas
    nota_secundaria = "Quando informado pela fonte"

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="metric-card"><h4>🚨 {metrica_label}</h4>'
        f'<h2 style="color:#d9534f; margin:0;">{formatar_inteiro(valor_principal)}</h2>'
        f'<p>{subtitulo_principal}</p></div>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f'<div class="metric-card"><h4>👥 {titulo_card_secundario}</h4>'
        f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(valor_card_secundario)}</h2>'
        f'<p>{nota_secundaria}</p></div>',
        unsafe_allow_html=True,
    )
with c3:
    valor_controle = total_vitimas_mun if total_vitimas_mun > 0 else registros
    label_controle = "Vitimas municipio" if total_vitimas_mun > 0 else "Linhas retornadas"
    st.markdown(
        f'<div class="metric-card"><h4>📋 {label_controle}</h4>'
        f'<h2 style="color:#f0ad4e; margin:0;">{formatar_inteiro(valor_controle)}</h2>'
        '<p>Controle da consulta</p></div>',
        unsafe_allow_html=True,
    )
with c4:
    st.markdown(
        f'<div class="metric-card"><h4>📍 Localidade</h4>'
        f'<h3 style="color:#1c2d42; margin:0;">{localidade}</h3>'
        f'<p>{ano_sel} | {mes_nome}</p></div>',
        unsafe_allow_html=True,
    )

if usa_oficial_municipios:
    st.info(
        "Leitura metodologica: na base oficial municipal do MJSP/SINESP, a metrica principal e vitimas. "
        "Por isso, os graficos e o card principal usam VITIMAS, nao OCORRENCIAS."
    )

if not usa_api:
    st.caption("Motor de filtro: DuckDB sobre Parquet local em cache. A primeira carga do XLSX ainda pode demorar, mas as consultas seguintes ficam mais leves.")

st.caption(
    "Fonte: "
    + ("MJSP/SINESP - dados nacionais de seguranca publica" if not usa_api else "API comunitaria rayonnunes/api_seguranca_publica - experimental")
)

st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Painel Estatistico",
    "📋 Dados Tratados",
    "⚙️ Dados Brutos/Diagnostico",
    "📥 Exportacao",
])

with tab1:
    col_a, col_b = st.columns(2)
    with col_a:
        st.write(f"**Evolucao mensal - {ano_sel} ({metrica_label})**")
        plotar_serie_mensal(df_consulta, metrica_coluna, metrica_label)
    with col_b:
        st.write(f"**Distribuicao por indicador/tipo de crime ({metrica_label})**")
        plotar_crimes(df_consulta, metrica_coluna, metrica_label)

    st.write(f"**Ranking municipal - top 15 ({metrica_label})**")
    plotar_top_municipios(df_consulta, metrica_coluna, metrica_label)

with tab2:
    colunas_preferenciais = [
        "ANO", "MES_NOME", "UF", "MUNICIPIO", "TIPO_CRIME",
        "OCORRENCIAS", "VITIMAS", "VITIMAS_MUNICIPIO", "FONTE_PROCESSADA",
    ]
    colunas_existentes = [c for c in colunas_preferenciais if c in df_consulta.columns]
    outras_colunas = [c for c in df_consulta.columns if c not in colunas_existentes]
    df_grid = df_consulta[colunas_existentes + outras_colunas].copy()
    st.dataframe(df_grid, width="stretch")

with tab3:
    st.write("**Fonte selecionada:**", fonte_dados)
    if usa_api:
        st.write("**Metadados da API:**")
        st.json(meta_api)
    else:
        st.write("**Metadados da base oficial/cache:**")
        st.json(meta_oficial)
        st.write("**Diagnostico das etapas do filtro:**")
        diag_ok = diagnosticar_etapas_filtro_duckdb(
            parquet_path,
            uf=uf_sel,
            municipio=municipio_param,
            crime=crime_param,
            ano=ano_sel,
            mes_num=mes_param,
        )
        st.dataframe(pd.DataFrame(diag_ok["etapas"]), width="stretch")
    st.write("**Colunas finais:**")
    st.code("\n".join(df_consulta.columns.astype(str).tolist()))
    st.write("**Amostra da consulta já filtrada:**")
    st.dataframe(df_consulta.head(50), width="stretch")

with tab4:
    nome_arquivo = f"seguranca_{uf_sel}_{municipio_param or 'todos'}_{ano_sel}_{mes_param or 'todos'}.csv"
    nome_arquivo = normalizar_coluna(nome_arquivo).lower().replace("_csv", ".csv")
    csv = df_consulta.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
    st.download_button(
        "📥 Baixar CSV tratado",
        data=csv,
        file_name=nome_arquivo,
        mime="text/csv",
    )

    st.markdown(
        """
        **Observacao metodologica:** os dados oficiais do MJSP/SINESP dependem da alimentacao,
        validacao e consolidacao feita pelos estados e pelo Distrito Federal. Na base municipal,
        a unidade principal e vitimas. Valores podem mudar quando a base oficial e atualizada.
        """
    )
