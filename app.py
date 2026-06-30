# VERSION_FINAL_PRODUCAO_SEGURANCA_DASHBOARD_V9_OFICIAL_DUCKDB
# App Streamlit para dados oficiais MJSP/SINESP de seguranca publica.
# Mudancas da V8:
# - Remove API comunitaria e fonte experimental.
# - Remove seletor confuso de fonte; usa apenas fontes oficiais por nivel de analise.
# - Nivel Municipio: base oficial municipal, unidade principal = Vitimas, indicador = Homicidio doloso.
# - Nivel UF: base oficial por Unidade da Federacao, com varios tipos de crime.
# - DuckDB + Parquet em cache versionado para evitar parquet antigo com TIPO_CRIME = MG.
# - Ranking municipal aparece somente quando a consulta esta em Todos os municipios.
# - Distribuicao por tipo de crime aparece somente quando ha mais de um indicador no recorte.

import io
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
    .small-note { color: #666; font-size: 13px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# FONTES OFICIAIS MJSP/SINESP
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

CACHE_SCHEMA_VERSION = "v9_oficial_20260630"

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

MES_LABEL = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
}

MES_NOME_LONGO = {
    1: "Janeiro", 2: "Fevereiro", 3: "Marco", 4: "Abril", 5: "Maio", 6: "Junho",
    7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro",
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
        "CODIGO_MUNICIPIO", "COD_MUNICIPIO", "CODIGO_IBGE", "COD_IBGE", "COD_IBGE",
        "COD_MUN", "CODMUN", "ID_MUNICIPIO",
    ],
    "REGIAO": ["REGIAO", "REGIAO_BRASIL"],
    "ANO": ["ANO", "ANO_REFERENCIA", "ANO_DE_REFERENCIA", "PERIODO_ANO"],
    "MES": ["MES", "MES_REFERENCIA", "MES_DE_REFERENCIA", "PERIODO_MES"],
    "MES_ANO": ["MES_ANO", "MESANO", "MES_REFERENCIA_ANO", "PERIODO"],
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
    "SEXO_VITIMA": ["SEXO_DA_VITIMA", "SEXO_VITIMA"],
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
# FUNCOES DE NORMALIZACAO
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


def converter_numero(valor):
    if pd.isna(valor):
        return 0
    texto = str(valor).strip()
    if texto in ["", "-", "--", "nan", "None"]:
        return 0
    texto = texto.replace("\u00a0", "")
    texto = re.sub(r"[^0-9,.-]", "", texto)
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


def parse_data_mista(valor):
    """Converte datas sem gerar warning de dayfirst para datas ISO."""
    if pd.isna(valor):
        return pd.NaT
    s = str(valor).strip()
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

    return pd.to_datetime(s, errors="coerce", dayfirst=False)


def ano_para_texto(valor):
    if pd.isna(valor):
        return None
    s = str(valor).strip()
    if not s:
        return None
    dt = parse_data_mista(s)
    if pd.notna(dt):
        return str(int(dt.year))
    match4 = re.search(r"(19\d{2}|20\d{2})", normalizar_valor(s))
    if match4:
        return match4.group(1)
    return None


def mes_para_ordem(valor):
    if pd.isna(valor):
        return None
    s = str(valor).strip()
    if not s:
        return None
    dt = parse_data_mista(s)
    if pd.notna(dt):
        return int(dt.month)
    chave = normalizar_valor(s).replace(".", "")
    if chave in MES_NORMALIZADO:
        return MES_NORMALIZADO[chave]
    match = re.search(r"\b(0?[1-9]|1[0-2])\b", chave)
    if match:
        return int(match.group(1))
    return None


def titulo_localidade(nivel, uf, municipio):
    if nivel == "Municipio":
        if municipio:
            return f"{municipio} - {uf}"
        return f"Municipios de {uf} (homicidio doloso)"
    return f"Estado {uf} - agregado estadual"

# -----------------------------------------------------------------------------
# LEITURA E PADRONIZACAO DAS BASES OFICIAIS
# -----------------------------------------------------------------------------
def identificar_e_renomear_colunas(df):
    df = df.copy()
    df.columns = [normalizar_coluna(c) for c in df.columns]

    colunas_mensais = [c for c in df.columns if c in COLUNAS_MESES_WIDE]
    tem_coluna_mes = any(c in df.columns for c in COLUNAS_CANONICAS["MES"])
    tem_coluna_mes_ano = any(c in df.columns for c in COLUNAS_CANONICAS["MES_ANO"])
    tem_coluna_ocorrencias = any(c in df.columns for c in COLUNAS_CANONICAS["OCORRENCIAS"])

    if colunas_mensais and not tem_coluna_mes and not tem_coluna_mes_ano and not tem_coluna_ocorrencias:
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


def padronizar_base_seguranca(df, tipo_base):
    if df is None or df.empty:
        return pd.DataFrame()

    df = identificar_e_renomear_colunas(df)

    for col in ["UF", "MUNICIPIO", "TIPO_CRIME", "MES", "MES_ANO", "ANO", "REGIAO", "UNIDADE_MEDIDA"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # UF pode vir como sigla ou nome por extenso.
    nome_para_sigla = {chave_comparacao(v): k for k, v in UF_NOMES.items()}
    if "UF" in df.columns:
        def ajustar_uf(x):
            norm = normalizar_valor(x)
            if len(norm) == 2 and norm in UFS:
                return norm
            return nome_para_sigla.get(chave_comparacao(x), norm[:2] if len(norm) == 2 else norm)
        df["UF"] = df["UF"].map(ajustar_uf)
    elif "ABA_ORIGEM" in df.columns:
        df["UF"] = df["ABA_ORIGEM"].map(lambda x: normalizar_valor(x) if normalizar_valor(x) in UFS else "")

    if "MUNICIPIO" in df.columns:
        df["MUNICIPIO"] = df["MUNICIPIO"].map(lambda x: str(x).strip().title())

    # Municipio oficial: as abas sao UFs, nao indicadores. O dicionario municipal
    # documenta Homicidio doloso como indicador e Vitimas como unidade principal.
    if tipo_base == "municipios":
        df["TIPO_CRIME"] = "Homicidio doloso"
        df["UNIDADE_MEDIDA"] = "Vitimas"
        if "OCORRENCIAS" not in df.columns:
            df["OCORRENCIAS"] = 0
        if "VITIMAS_MUNICIPIO" not in df.columns:
            df["VITIMAS_MUNICIPIO"] = df["VITIMAS"] if "VITIMAS" in df.columns else 0

    # UF oficial: as abas representam unidades de medida, nao UF nem crime.
    if tipo_base == "uf":
        if "UNIDADE_MEDIDA" not in df.columns and "ABA_ORIGEM" in df.columns:
            df["UNIDADE_MEDIDA"] = df["ABA_ORIGEM"].map(
                lambda x: "Vitimas" if "VIT" in normalizar_valor(x) else "Ocorrencias"
            )
        if "TIPO_CRIME" not in df.columns:
            df["TIPO_CRIME"] = "Nao informado"
        if "MUNICIPIO" not in df.columns:
            df["MUNICIPIO"] = ""
        if "CODIGO_MUNICIPIO" not in df.columns:
            df["CODIGO_MUNICIPIO"] = ""
        if "VITIMAS_MUNICIPIO" not in df.columns:
            df["VITIMAS_MUNICIPIO"] = 0

    # ANO/MES podem vir separados ou em MES_ANO.
    if "MES" not in df.columns and "MES_ANO" in df.columns:
        df["MES"] = df["MES_ANO"]

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
        mes_extraido = df["MES"].map(mes_para_ordem)
        if "MES_ANO" in df.columns:
            mes_extraido = mes_extraido.fillna(df["MES_ANO"].map(mes_para_ordem))
        df["MES_ORDEM"] = mes_extraido
        df["MES_NOME"] = df["MES_ORDEM"].map(MES_LABEL).fillna(df["MES"].astype(str))
    elif "MES_ANO" in df.columns:
        df["MES_ORDEM"] = df["MES_ANO"].map(mes_para_ordem)
        df["MES_NOME"] = df["MES_ORDEM"].map(MES_LABEL).fillna(df["MES_ANO"].astype(str))
    else:
        df["MES_ORDEM"] = None
        df["MES_NOME"] = "Nao informado"

    for col in ["OCORRENCIAS", "VITIMAS", "VITIMAS_MUNICIPIO"]:
        if col in df.columns:
            df[col] = df[col].map(converter_numero)
        else:
            df[col] = 0

    if "TIPO_CRIME" in df.columns:
        df["TIPO_CRIME"] = df["TIPO_CRIME"].astype(str).str.strip()
        df.loc[df["TIPO_CRIME"].isin(["", "nan", "None"]), "TIPO_CRIME"] = "Nao informado"
    else:
        df["TIPO_CRIME"] = "Nao informado"

    df["FONTE_PROCESSADA"] = f"oficial_mjsp_{tipo_base}"
    df = df.loc[:, ~df.columns.duplicated()].copy()
    return df


def detectar_linha_cabecalho_excel(conteudo_bytes, aba):
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

    palavras_chave = ["UF", "MUNICIPIO", "COD", "REGIAO", "ANO", "MES", "CRIME", "OCORRENCIA", "VITIMA"]
    melhor_linha = 0
    melhor_pontos = -1
    for idx, row in preview.iterrows():
        valores = [normalizar_coluna(v) for v in row.dropna().tolist()]
        pontos = sum(1 for palavra in palavras_chave if any(palavra in valor for valor in valores))
        if pontos > melhor_pontos:
            melhor_pontos = pontos
            melhor_linha = int(idx)
    return melhor_linha if melhor_pontos >= 2 else 0


def ler_excel_oficial_multiplas_abas(conteudo_bytes):
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
            metadados_abas.append({"aba": aba, "linhas": int(len(df_aba)), "cabecalho_linha": header_row})
        except Exception as e:
            metadados_abas.append({"aba": aba, "linhas": 0, "erro": str(e)})

    if not frames:
        return pd.DataFrame(), metadados_abas
    return pd.concat(frames, ignore_index=True, sort=False), metadados_abas


def adicionar_colunas_filtro(df):
    df = df.copy()
    df["UF_FILTRO"] = df.get("UF", pd.Series([""] * len(df))).astype(str).str.upper().str.strip()
    df["MUNICIPIO_FILTRO"] = df.get("MUNICIPIO", pd.Series([""] * len(df))).map(chave_comparacao)
    df["TIPO_CRIME_FILTRO"] = df.get("TIPO_CRIME", pd.Series([""] * len(df))).map(chave_comparacao)
    df["ANO_FILTRO"] = df.get("ANO", pd.Series([""] * len(df))).astype(str).str.strip()
    df["UNIDADE_MEDIDA_FILTRO"] = df.get("UNIDADE_MEDIDA", pd.Series([""] * len(df))).map(chave_comparacao)
    df["MES_ORDEM_NUM"] = pd.to_numeric(df.get("MES_ORDEM", pd.Series([pd.NA] * len(df))), errors="coerce")
    return df


def caminho_parquet_cache(tipo_base):
    pasta = os.path.join(tempfile.gettempdir(), "sinesp_mjsp_cache_duckdb")
    os.makedirs(pasta, exist_ok=True)
    return os.path.join(pasta, f"mjsp_sinesp_{tipo_base}_{CACHE_SCHEMA_VERSION}.parquet")


@st.cache_data(ttl=86400, show_spinner=False)
def preparar_base_oficial_parquet(tipo_base):
    url = URL_MJSP_MUNICIPIOS if tipo_base == "municipios" else URL_MJSP_UF
    parquet_path = caminho_parquet_cache(tipo_base)

    try:
        # Fallback local ajuda em desenvolvimento; no Streamlit Cloud a URL oficial sera usada.
        nome_local = "indicadoressegurancapublicamunic.xlsx" if tipo_base == "municipios" else "indicadoressegurancapublicauf.xlsx"
        if os.path.exists(nome_local):
            with open(nome_local, "rb") as f:
                conteudo = f.read()
            origem = f"arquivo local: {nome_local}"
        else:
            resp = requests.get(url, timeout=180, verify=False, headers={"User-Agent": "Mozilla/5.0"})
            if resp.status_code != 200:
                return {
                    "ok": False,
                    "erro": f"HTTP {resp.status_code}",
                    "url": url,
                    "texto": resp.text[:1500],
                    "parquet_path": parquet_path,
                }
            conteudo = resp.content
            origem = url

        df_raw, meta_abas = ler_excel_oficial_multiplas_abas(conteudo)
        if df_raw.empty:
            return {
                "ok": False,
                "erro": "XLSX carregado, mas nenhuma aba util foi identificada.",
                "url": url,
                "abas": meta_abas,
                "parquet_path": parquet_path,
            }

        df = padronizar_base_seguranca(df_raw, tipo_base)
        df = adicionar_colunas_filtro(df)
        df.to_parquet(parquet_path, index=False)

        return {
            "ok": True,
            "url": url,
            "origem_usada": origem,
            "tipo_base": tipo_base,
            "cache_schema_version": CACHE_SCHEMA_VERSION,
            "linhas_brutas": int(len(df_raw)),
            "linhas": int(len(df)),
            "abas": meta_abas,
            "colunas": list(df.columns),
            "parquet_path": parquet_path,
            "motor_filtro": "DuckDB sobre Parquet local em cache versionado",
        }
    except Exception as e:
        return {"ok": False, "erro": str(e), "url": url, "parquet_path": parquet_path}

# -----------------------------------------------------------------------------
# DUCKDB
# -----------------------------------------------------------------------------
def duckdb_df(parquet_path, query_suffix="", params=None):
    if params is None:
        params = []
    if not parquet_path or not os.path.exists(parquet_path):
        return pd.DataFrame()
    path_sql = parquet_path.replace("'", "''")
    con = duckdb.connect(database=":memory:")
    try:
        return con.execute(f"SELECT * FROM read_parquet('{path_sql}') {query_suffix}", params).fetchdf()
    finally:
        con.close()


def duckdb_scalar(parquet_path, sql_expr, where="", params=None):
    if params is None:
        params = []
    if not parquet_path or not os.path.exists(parquet_path):
        return None
    path_sql = parquet_path.replace("'", "''")
    con = duckdb.connect(database=":memory:")
    try:
        return con.execute(f"SELECT {sql_expr} FROM read_parquet('{path_sql}') {where}", params).fetchone()[0]
    finally:
        con.close()


def get_distinct_options(parquet_path, coluna, where="", params=None):
    if params is None:
        params = []
    if not parquet_path or not os.path.exists(parquet_path):
        return []
    path_sql = parquet_path.replace("'", "''")
    con = duckdb.connect(database=":memory:")
    try:
        df = con.execute(
            f"SELECT DISTINCT {coluna} AS valor FROM read_parquet('{path_sql}') {where} "
            "WHERE valor IS NOT NULL" if not where else
            f"SELECT DISTINCT {coluna} AS valor FROM read_parquet('{path_sql}') {where} AND {coluna} IS NOT NULL",
            params,
        ).fetchdf()
    except Exception:
        try:
            df = con.execute(f"SELECT DISTINCT {coluna} AS valor FROM read_parquet('{path_sql}') {where}", params).fetchdf()
        except Exception:
            df = pd.DataFrame()
    finally:
        con.close()
    valores = [str(x).strip() for x in df.get("valor", pd.Series(dtype=str)).dropna().tolist()]
    return sorted([v for v in valores if v and v.lower() not in ["nan", "none", "nao informado"]])


def construir_where(uf, municipio, crime, ano, mes_num, unidade_medida):
    clauses = []
    params = []
    if uf:
        clauses.append("UF_FILTRO = ?")
        params.append(uf.upper())
    if municipio:
        chave = chave_comparacao(municipio)
        clauses.append("(MUNICIPIO_FILTRO = ? OR MUNICIPIO_FILTRO LIKE ?)")
        params.extend([chave, f"%{chave}%"])
    if crime and crime != "Todos os indicadores":
        clauses.append("TIPO_CRIME_FILTRO = ?")
        params.append(chave_comparacao(crime))
    if ano and ano != "Todos os anos":
        clauses.append("ANO_FILTRO = ?")
        params.append(str(ano))
    if mes_num:
        clauses.append("CAST(MES_ORDEM_NUM AS INTEGER) = ?")
        params.append(int(mes_num))
    if unidade_medida:
        clauses.append("UNIDADE_MEDIDA_FILTRO = ?")
        params.append(chave_comparacao(unidade_medida))
    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    return where, params


def query_dados(parquet_path, uf, municipio, crime, ano, mes_num, unidade_medida, limit=None):
    where, params = construir_where(uf, municipio, crime, ano, mes_num, unidade_medida)
    limite = f" LIMIT {int(limit)}" if limit else ""
    return duckdb_df(parquet_path, where + limite, params)


def diagnosticar(parquet_path, uf, municipio, crime, ano, mes_num, unidade_medida):
    etapas = []
    total = duckdb_scalar(parquet_path, "COUNT(*)") or 0
    etapas.append({"etapa": "Base completa em cache", "linhas": int(total)})

    filtros = []
    params = []
    def contar(etapa):
        where = " WHERE " + " AND ".join(filtros) if filtros else ""
        linhas = duckdb_scalar(parquet_path, "COUNT(*)", where, params) or 0
        etapas.append({"etapa": etapa, "linhas": int(linhas)})

    if uf:
        filtros.append("UF_FILTRO = ?")
        params.append(uf.upper())
        contar(f"Apos UF = {uf}")
    if municipio:
        chave = chave_comparacao(municipio)
        filtros.append("(MUNICIPIO_FILTRO = ? OR MUNICIPIO_FILTRO LIKE ?)")
        params.extend([chave, f"%{chave}%"])
        contar(f"Apos municipio = {municipio}")
    if crime and crime != "Todos os indicadores":
        filtros.append("TIPO_CRIME_FILTRO = ?")
        params.append(chave_comparacao(crime))
        contar(f"Apos indicador = {crime}")
    if ano and ano != "Todos os anos":
        filtros.append("ANO_FILTRO = ?")
        params.append(str(ano))
        contar(f"Apos ano = {ano}")
    if mes_num:
        filtros.append("CAST(MES_ORDEM_NUM AS INTEGER) = ?")
        params.append(int(mes_num))
        contar(f"Apos mes = {mes_num}")
    if unidade_medida:
        filtros.append("UNIDADE_MEDIDA_FILTRO = ?")
        params.append(chave_comparacao(unidade_medida))
        contar(f"Apos unidade = {unidade_medida}")

    where, p = construir_where(uf, municipio, "Todos os indicadores", "Todos os anos", None, unidade_medida)
    amostra = query_dados(parquet_path, uf, municipio, crime, ano, mes_num, unidade_medida, limit=30)
    anos = get_distinct_options(parquet_path, "ANO_FILTRO", where, p)
    crimes = get_distinct_options(parquet_path, "TIPO_CRIME", where, p)
    return {"etapas": etapas, "amostra": amostra, "anos": sorted(anos, reverse=True), "indicadores": crimes}

# -----------------------------------------------------------------------------
# GRAFICOS
# -----------------------------------------------------------------------------
def soma_segura(df, coluna):
    if df is None or df.empty or coluna not in df.columns:
        return 0
    return pd.to_numeric(df[coluna], errors="coerce").fillna(0).sum()


def plotar_serie_mensal(df, metrica_coluna, metrica_label):
    if df.empty or "MES_ORDEM" not in df.columns or metrica_coluna not in df.columns:
        st.info(f"Nao ha dados mensais suficientes para {metrica_label.lower()}.")
        return
    base = df.copy()
    base["MES_ORDEM"] = pd.to_numeric(base["MES_ORDEM"], errors="coerce")
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    base = base.dropna(subset=["MES_ORDEM"])
    if base.empty or base[metrica_coluna].sum() == 0:
        st.info(f"Nao ha valores de {metrica_label.lower()} para a serie mensal.")
        return
    serie = base.groupby("MES_ORDEM")[metrica_coluna].sum().reset_index().sort_values("MES_ORDEM")
    serie["Mes"] = serie["MES_ORDEM"].astype(int).map(lambda m: f"{int(m):02d} - {MES_LABEL.get(int(m), str(int(m)))}")
    # Usamos indice numerico/ordenado para evitar que o eixo fique alfabetico (Abr, Ago, Dez...).
    serie_plot = serie.set_index("MES_ORDEM")[[metrica_coluna]].rename(columns={metrica_coluna: metrica_label})
    st.line_chart(serie_plot)
    with st.expander("Ver serie mensal em tabela", expanded=False):
        st.dataframe(serie[["Mes", metrica_coluna]].rename(columns={metrica_coluna: metrica_label}), width="stretch")


def plotar_crimes(df, metrica_coluna, metrica_label, nivel):
    if df.empty or "TIPO_CRIME" not in df.columns or metrica_coluna not in df.columns:
        st.info("Nao ha dados suficientes para a distribuicao por indicador/tipo de crime.")
        return
    base = df.copy()
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    serie = base.groupby("TIPO_CRIME")[metrica_coluna].sum().sort_values(ascending=False)
    serie = serie[serie > 0]
    if len(serie) <= 1:
        if nivel == "Municipio":
            st.info("A base oficial municipal do MJSP/SINESP possui apenas Homicidio doloso como indicador; por isso nao ha distribuicao por varios tipos de crime nesse nivel.")
        else:
            st.info("O recorte selecionado possui apenas um indicador com valor positivo.")
        return
    st.bar_chart(serie.head(15))


def plotar_top_municipios(df, metrica_coluna, metrica_label, municipio_param):
    if municipio_param:
        st.info("Ranking municipal fica oculto quando um municipio especifico esta selecionado. Use 'Todos os municipios' para comparar o top 15 da UF.")
        return
    if df.empty or "MUNICIPIO" not in df.columns or metrica_coluna not in df.columns:
        st.info("Selecione o nivel Municipio e 'Todos os municipios' para ver o ranking municipal.")
        return
    base = df.copy()
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    serie = base.groupby("MUNICIPIO")[metrica_coluna].sum().sort_values(ascending=False)
    serie = serie[serie > 0]
    if serie.empty:
        st.info(f"Nao ha valores de {metrica_label.lower()} para montar o ranking municipal.")
        return
    st.bar_chart(serie.head(15))

# -----------------------------------------------------------------------------
# SIDEBAR: PREPARA BASE SELECIONADA E MONTA FILTROS LIMITADOS AO QUE EXISTE
# -----------------------------------------------------------------------------
st.sidebar.title("🛡️ Filtros de Seguranca")

nivel_label = st.sidebar.radio(
    "Nivel de analise:",
    [
        "Municipio - homicidio doloso/vitimas",
        "UF - varios tipos de crime",
    ],
    index=0,
    help=(
        "Municipio usa a base oficial municipal, que disponibiliza homicidio doloso em vitimas. "
        "UF usa a base oficial agregada por Unidade da Federacao, com varios tipos de crime."
    ),
)
nivel_analise = "Municipio" if nivel_label.startswith("Municipio") else "UF"
tipo_base = "municipios" if nivel_analise == "Municipio" else "uf"

with st.spinner("Preparando base oficial MJSP/SINESP em cache DuckDB..."):
    meta = preparar_base_oficial_parquet(tipo_base)

if not meta.get("ok"):
    st.error("Nao foi possivel preparar a base oficial MJSP/SINESP.")
    st.json(meta)
    st.stop()

parquet_path = meta.get("parquet_path", "")

uf_sel = st.sidebar.selectbox("Selecione o Estado:", sorted(UFS), index=sorted(UFS).index("MG"))

municipio_param = None
municipio_sel = "Todos os municipios"
if nivel_analise == "Municipio":
    municipios_uf = get_distinct_options(parquet_path, "MUNICIPIO", " WHERE UF_FILTRO = ?", [uf_sel])
    municipios_uf = sorted({str(m).strip().title() for m in municipios_uf if str(m).strip()})
    municipio_sel = st.sidebar.selectbox("Selecione o Municipio:", ["Todos os municipios"] + municipios_uf)
    municipio_param = None if municipio_sel == "Todos os municipios" else municipio_sel
else:
    st.sidebar.info("A base por UF nao possui recorte municipal.")

# Unidade/metrica
if nivel_analise == "Municipio":
    unidade_medida = "Vitimas"
    metrica_coluna = "VITIMAS"
    metrica_label = "Vitimas"
    st.sidebar.caption("No nivel municipal, a base oficial disponivel e Homicidio doloso / Vitimas.")
else:
    metrica_label = st.sidebar.selectbox("Metrica:", ["Ocorrencias", "Vitimas"], index=0)
    unidade_medida = metrica_label
    metrica_coluna = "OCORRENCIAS" if metrica_label == "Ocorrencias" else "VITIMAS"

# Contexto para opcoes: UF + municipio + unidade.
contexto_where, contexto_params = construir_where(
    uf=uf_sel,
    municipio=municipio_param,
    crime="Todos os indicadores",
    ano="Todos os anos",
    mes_num=None,
    unidade_medida=unidade_medida,
)

indicadores = get_distinct_options(parquet_path, "TIPO_CRIME", contexto_where, contexto_params)
if nivel_analise == "Municipio":
    indicadores = ["Homicidio doloso"]
    indicadores_opcoes = indicadores
else:
    indicadores_opcoes = ["Todos os indicadores"] + indicadores
crime_sel = st.sidebar.selectbox("Indicador / Tipo de Crime:", indicadores_opcoes)

anos_disponiveis = get_distinct_options(parquet_path, "ANO_FILTRO", contexto_where, contexto_params)
anos_disponiveis = sorted([a for a in anos_disponiveis if re.fullmatch(r"\d{4}", str(a))], reverse=True)
ano_opcoes = ["Todos os anos"] + anos_disponiveis
ano_sel = st.sidebar.selectbox("Ano de Referencia:", ano_opcoes, index=1 if len(ano_opcoes) > 1 else 0)

mes_nome = st.sidebar.selectbox("Mes:", list(MESES_DISPLAY.keys()), index=0)
mes_param = MESES_DISPLAY[mes_nome]

with st.sidebar.form("form_seguranca"):
    submit_btn = st.form_submit_button("🔍 Consultar Indicadores")

# -----------------------------------------------------------------------------
# SAIDA
# -----------------------------------------------------------------------------
if not submit_btn:
    st.info("💡 Escolha o nivel, localidade, indicador e periodo na barra lateral e clique em Consultar Indicadores.")
    st.markdown(
        """
        **Como usar:**
        
        - **Municipio:** usa a base oficial municipal do MJSP/SINESP. O dicionario municipal indica unidade principal em **vitimas** e indicador **homicidio doloso**.
        - **UF:** usa a base oficial por Unidade da Federacao, com mais tipos de crime e metricas de ocorrencias/vitimas. Esta base e agregada por estado, nao por municipio.
        - O app usa **DuckDB sobre Parquet em cache versionado**, evitando reaproveitar arquivos antigos com estrutura errada.
        """
    )
    st.stop()

st.markdown(
    '<div class="header-seguranca"><h1>Painel de Ocorrencias Criminais - SINESP/MJSP</h1>'
    '<p>Monitoramento territorial de indicadores oficiais de seguranca publica</p></div>',
    unsafe_allow_html=True,
)

df_consulta = query_dados(
    parquet_path=parquet_path,
    uf=uf_sel,
    municipio=municipio_param,
    crime=crime_sel,
    ano=ano_sel,
    mes_num=mes_param,
    unidade_medida=unidade_medida,
)

if df_consulta.empty:
    st.error("🛑 Nao foram encontrados registros para a combinacao selecionada.")
    diag = diagnosticar(parquet_path, uf_sel, municipio_param, crime_sel, ano_sel, mes_param, unidade_medida)
    with st.expander("Diagnostico tecnico", expanded=True):
        st.write("**Nivel:**", nivel_analise)
        st.write("**UF:**", uf_sel)
        st.write("**Municipio:**", municipio_param or "Todos")
        st.write("**Indicador:**", crime_sel)
        st.write("**Ano:**", ano_sel)
        st.write("**Mes:**", mes_nome)
        st.write("**Unidade/metrica:**", unidade_medida)
        st.dataframe(pd.DataFrame(diag["etapas"]), width="stretch")
        st.write("Anos disponiveis no contexto:", diag.get("anos", []))
        st.write("Indicadores disponiveis no contexto:", diag.get("indicadores", []))
        st.write("Metadados da base/cache:")
        st.json(meta)
        st.write("Amostra contextual:")
        if diag["amostra"].empty:
            st.info("Sem amostra para esse recorte.")
        else:
            st.dataframe(diag["amostra"], width="stretch")
    st.stop()

# Cards
localidade = titulo_localidade(nivel_analise, uf_sel, municipio_param)
valor_principal = soma_segura(df_consulta, metrica_coluna)
total_ocorrencias = soma_segura(df_consulta, "OCORRENCIAS")
total_vitimas = soma_segura(df_consulta, "VITIMAS")
registros = len(df_consulta)

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="metric-card"><h4>🚨 {metrica_label}</h4>'
        f'<h2 style="color:#d9534f; margin:0;">{formatar_inteiro(valor_principal)}</h2>'
        f'<p>Metrica principal do recorte</p></div>',
        unsafe_allow_html=True,
    )
with c2:
    valor_sec = total_ocorrencias if metrica_coluna != "OCORRENCIAS" else total_vitimas
    label_sec = "Ocorrencias" if metrica_coluna != "OCORRENCIAS" else "Vitimas"
    st.markdown(
        f'<div class="metric-card"><h4>📋 {label_sec}</h4>'
        f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(valor_sec)}</h2>'
        '<p>Quando informado pela fonte</p></div>',
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f'<div class="metric-card"><h4>🧾 Linhas retornadas</h4>'
        f'<h2 style="color:#f0ad4e; margin:0;">{formatar_inteiro(registros)}</h2>'
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

if nivel_analise == "Municipio":
    st.info("Leitura metodologica: no nivel municipal, a fonte oficial disponivel no XLSX do MJSP/SINESP corresponde a Homicidio doloso e unidade principal em Vitimas. Para consultar mais tipos de crime, use o nivel UF.")
else:
    st.info("Leitura metodologica: este e um agregado estadual. A base por UF traz varios tipos de crime, mas nao permite desagregar esses crimes por municipio.")

st.caption("Fonte: MJSP/SINESP - Dados Nacionais de Seguranca Publica. Motor: DuckDB sobre Parquet local em cache versionado.")
st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Painel Estatistico",
    "📋 Dados Tratados",
    "⚙️ Diagnostico",
    "📥 Exportacao",
])

with tab1:
    col_a, col_b = st.columns(2)
    with col_a:
        st.write(f"**Evolucao mensal ({metrica_label})**")
        plotar_serie_mensal(df_consulta, metrica_coluna, metrica_label)
    with col_b:
        st.write(f"**Distribuicao por indicador/tipo de crime ({metrica_label})**")
        plotar_crimes(df_consulta, metrica_coluna, metrica_label, nivel_analise)

    if nivel_analise == "Municipio":
        st.write(f"**Ranking municipal - top 15 ({metrica_label})**")
        plotar_top_municipios(df_consulta, metrica_coluna, metrica_label, municipio_param)
    else:
        st.info("Ranking municipal nao se aplica ao nivel UF, pois essa base nao possui recorte por municipio.")

with tab2:
    colunas_preferenciais = [
        "ANO", "MES_NOME", "UF", "MUNICIPIO", "TIPO_CRIME", "UNIDADE_MEDIDA",
        "OCORRENCIAS", "VITIMAS", "VITIMAS_MUNICIPIO", "SEXO_VITIMA", "FONTE_PROCESSADA",
    ]
    colunas_existentes = [c for c in colunas_preferenciais if c in df_consulta.columns]
    outras = [c for c in df_consulta.columns if c not in colunas_existentes]
    st.dataframe(df_consulta[colunas_existentes + outras], width="stretch")

with tab3:
    diag_ok = diagnosticar(parquet_path, uf_sel, municipio_param, crime_sel, ano_sel, mes_param, unidade_medida)
    st.write("**Contagem por etapa do filtro:**")
    st.dataframe(pd.DataFrame(diag_ok["etapas"]), width="stretch")
    st.write("**Anos disponiveis no contexto:**", diag_ok.get("anos", []))
    st.write("**Indicadores disponiveis no contexto:**", diag_ok.get("indicadores", []))
    st.write("**Metadados da base/cache:**")
    st.json(meta)
    st.write("**Colunas finais:**")
    st.code("\n".join(df_consulta.columns.astype(str).tolist()))
    st.write("**Amostra da consulta filtrada:**")
    st.dataframe(df_consulta.head(50), width="stretch")

with tab4:
    nome_base = f"seguranca_{nivel_analise}_{uf_sel}_{municipio_param or 'todos'}_{ano_sel}_{mes_param or 'todos'}_{metrica_label}.csv"
    nome_base = normalizar_coluna(nome_base).lower().replace("_csv", ".csv")
    csv = df_consulta.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
    st.download_button(
        "📥 Baixar CSV tratado",
        data=csv,
        file_name=nome_base,
        mime="text/csv",
    )
    st.markdown(
        """
        **Observacao metodologica:** os dados oficiais do MJSP/SINESP dependem da alimentacao,
        validacao e consolidacao pelas Unidades da Federacao. No nivel municipal, a base oficial
        carregada pelo recurso de municipios apresenta homicidio doloso em vitimas. No nivel UF,
        ha maior variedade de tipos de crime e metricas.
        """
    )
