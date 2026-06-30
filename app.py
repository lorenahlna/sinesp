# VERSION_FINAL_PRODUCAO_SEGURANCA_DASHBOARD_V11_SEXO_DUCKDB
# Fonte: bases oficiais MJSP/SINESP em XLSX, com DuckDB + Parquet em cache.
# V11 corrige:
# - Remove API comunitaria e fonte experimental.
# - Separa corretamente UF/Ocorrencias e UF/Vitimas, sem misturar linhas das duas abas.
# - Acrescenta painel estatistico de perfil das vitimas por sexo, com cards, percentuais e tabela.
# - Quando a metrica principal e Ocorrencias, consulta separadamente a aba Vitimas para o perfil por sexo.
# - Trata SEXO_DA_VITIMA como estratificacao da aba Vitimas, nao como vitimas individuais.
# - Municipio volta a consultar sem filtro de unidade que podia zerar registros.
# - Municipio = Homicidio doloso / Vitimas, conforme dicionario oficial municipal.
# - UF = varios tipos de crime; metrica selecionada filtra a aba correta.
# - Anos, meses, indicadores e municipios limitados ao que existe no cache.

import io
import os
import re
import tempfile
import unicodedata

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
    .stButton>button { background-color: #1c2d42; color: white; width: 100%; }
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

CACHE_SCHEMA_VERSION = "v11_sexo_20260630"

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
        "CODIGO_MUNICIPIO", "COD_MUNICIPIO", "CODIGO_IBGE", "COD_IBGE",
        "COD_MUN", "CODMUN", "ID_MUNICIPIO",
    ],
    "REGIAO": ["REGIAO", "REGIAO_BRASIL"],
    "ANO": ["ANO", "ANO_REFERENCIA", "ANO_DE_REFERENCIA", "PERIODO_ANO"],
    "MES": ["MES", "MES_REFERENCIA", "MES_DE_REFERENCIA", "PERIODO_MES"],
    "MES_ANO": ["MES_ANO", "MESANO", "MES_REFERENCIA_ANO", "PERIODO", "REFERENCIA"],
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
        "QUANTIDADE_VITIMAS", "VALOR",
    ],
    "VITIMAS_MUNICIPIO": [
        "VITIMAS_MUNICIPIO", "VITIMAS_NO_MUNICIPIO", "TOTAL_VITIMAS_MUNICIPIO",
    ],
    "SEXO_DA_VITIMA": ["SEXO_DA_VITIMA", "SEXO_VITIMA", "SEXO", "GENERO", "SEXO_DAS_VITIMAS"],
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
# NORMALIZACAO
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
    if texto in ["", "-", "--", "nan", "None", "NoneType"]:
        return 0
    texto = texto.replace("\u00a0", "")
    texto = re.sub(r"[^0-9,.-]", "", texto)
    if texto in ["", ".", ",", "-"]:
        return 0
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
    if nivel == "municipios":
        if municipio:
            return f"{municipio} - {uf}"
        return f"Municipios de {uf}"
    return f"Estado {uf} - agregado estadual"

# -----------------------------------------------------------------------------
# PADRONIZACAO DAS BASES
# -----------------------------------------------------------------------------
def identificar_e_renomear_colunas(df):
    df = df.copy()
    df.columns = [normalizar_coluna(c) for c in df.columns]

    renomear = {}
    colunas_existentes = set(df.columns)
    for canonica, candidatas in COLUNAS_CANONICAS.items():
        if canonica in colunas_existentes:
            continue
        for candidata in candidatas:
            candidata_norm = normalizar_coluna(candidata)
            if candidata_norm in colunas_existentes:
                # Evita mapear VALOR simultaneamente para ocorrencias e vitimas.
                if candidata_norm == "VALOR" and ("OCORRENCIAS" in renomear.values() or "VITIMAS" in renomear.values()):
                    continue
                renomear[candidata_norm] = canonica
                break
    if renomear:
        df = df.rename(columns=renomear)
    return df


def detectar_colunas_mensais(df):
    mensais = []
    for c in df.columns:
        if normalizar_coluna(c) in COLUNAS_MESES_WIDE:
            mensais.append(c)
    return mensais


def transformar_wide_para_long(df):
    df = df.copy()
    colunas_mensais = detectar_colunas_mensais(df)
    if not colunas_mensais:
        return df

    id_vars = [c for c in df.columns if c not in colunas_mensais]
    df_long = df.melt(
        id_vars=id_vars,
        value_vars=colunas_mensais,
        var_name="MES",
        value_name="VALOR_MENSAL",
    )
    df_long["MES_ORDEM"] = df_long["MES"].map(lambda c: COLUNAS_MESES_WIDE.get(normalizar_coluna(c)))
    df_long["MES_NOME"] = df_long["MES_ORDEM"].map(MES_LABEL)
    return df_long


def padronizar_sexo(valor):
    s = normalizar_valor(valor)
    if not s or s in ["NAN", "NONE", "NAO INFORMADO", "NI", "IGNORADO"]:
        return "Sexo NI"
    if "FEM" in s:
        return "Feminino"
    if "MASC" in s:
        return "Masculino"
    if "NAO" in s or "IGN" in s or "NI" in s:
        return "Sexo NI"
    return str(valor).strip().title()


def inferir_unidade_da_aba(aba):
    s = normalizar_valor(aba)
    if "VIT" in s:
        return "Vitimas"
    if "OCOR" in s:
        return "Ocorrencias"
    return ""


def ajustar_uf_coluna(df):
    nome_para_sigla = {chave_comparacao(v): k for k, v in UF_NOMES.items()}

    def ajustar_uf(x):
        norm = normalizar_valor(x)
        if norm in UFS:
            return norm
        chave = chave_comparacao(x)
        if chave in nome_para_sigla:
            return nome_para_sigla[chave]
        if len(norm) == 2:
            return norm
        return ""

    if "UF" in df.columns:
        df["UF"] = df["UF"].map(ajustar_uf)
    else:
        df["UF"] = ""

    # No arquivo municipal, normalmente cada aba e uma UF.
    if "ABA_ORIGEM" in df.columns:
        aba_uf = df["ABA_ORIGEM"].map(lambda x: normalizar_valor(x) if normalizar_valor(x) in UFS else "")
        df["UF"] = df["UF"].where(df["UF"].astype(str).str.len() > 0, aba_uf)
    return df


def padronizar_base_seguranca(df_raw, tipo_base):
    if df_raw is None or df_raw.empty:
        return pd.DataFrame()

    df = identificar_e_renomear_colunas(df_raw)
    df = transformar_wide_para_long(df)
    df = identificar_e_renomear_colunas(df)

    df = ajustar_uf_coluna(df)

    if "MUNICIPIO" in df.columns:
        df["MUNICIPIO"] = df["MUNICIPIO"].astype(str).str.strip().str.title()
    else:
        df["MUNICIPIO"] = ""

    if "CODIGO_MUNICIPIO" not in df.columns:
        df["CODIGO_MUNICIPIO"] = ""

    # Periodo
    if "MES_ANO" not in df.columns:
        df["MES_ANO"] = ""
    if "MES" not in df.columns:
        df["MES"] = df["MES_ANO"]

    if "ANO" in df.columns:
        anos = df["ANO"].map(ano_para_texto)
        anos = anos.fillna(df["MES_ANO"].map(ano_para_texto))
        df["ANO"] = anos.fillna("Nao informado")
    else:
        df["ANO"] = df["MES_ANO"].map(ano_para_texto).fillna("Nao informado")

    if "MES_ORDEM" not in df.columns:
        meses = df["MES"].map(mes_para_ordem)
        meses = meses.fillna(df["MES_ANO"].map(mes_para_ordem))
        df["MES_ORDEM"] = meses
    else:
        df["MES_ORDEM"] = pd.to_numeric(df["MES_ORDEM"], errors="coerce")

    df["MES_NOME"] = df["MES_ORDEM"].map(lambda m: MES_LABEL.get(int(m), "Nao informado") if pd.notna(m) else "Nao informado")

    # Indicador e unidade
    if tipo_base == "municipios":
        # O dicionario municipal do MJSP/SINESP documenta Homicidio doloso / Vitimas.
        df["TIPO_CRIME"] = "Homicidio doloso"
        df["UNIDADE_MEDIDA"] = "Vitimas"
    else:
        if "TIPO_CRIME" not in df.columns:
            df["TIPO_CRIME"] = "Nao informado"
        df["TIPO_CRIME"] = df["TIPO_CRIME"].astype(str).str.strip().replace({"": "Nao informado", "nan": "Nao informado"})
        if "UNIDADE_MEDIDA" not in df.columns:
            if "ABA_ORIGEM" in df.columns:
                df["UNIDADE_MEDIDA"] = df["ABA_ORIGEM"].map(inferir_unidade_da_aba)
            else:
                df["UNIDADE_MEDIDA"] = ""
        df["UNIDADE_MEDIDA"] = df["UNIDADE_MEDIDA"].replace({"": "Nao informado"})

    # Sexo da vitima e uma dimensao de agregacao da aba Vitimas.
    if "SEXO_DA_VITIMA" in df.columns:
        df["SEXO_DA_VITIMA"] = df["SEXO_DA_VITIMA"].map(padronizar_sexo)
    else:
        df["SEXO_DA_VITIMA"] = ""

    # Campo de valor: apos melt, VALOR_MENSAL e o valor do mes.
    if "VALOR_MENSAL" in df.columns:
        valor = df["VALOR_MENSAL"].map(converter_numero)
    elif "OCORRENCIAS" in df.columns:
        valor = df["OCORRENCIAS"].map(converter_numero)
    elif "VITIMAS" in df.columns:
        valor = df["VITIMAS"].map(converter_numero)
    else:
        valor = pd.Series([0] * len(df), index=df.index)

    df["OCORRENCIAS"] = 0.0
    df["VITIMAS"] = 0.0
    df["VITIMAS_MUNICIPIO"] = 0.0

    if tipo_base == "municipios":
        df["VITIMAS"] = valor
        df["VITIMAS_MUNICIPIO"] = valor
    else:
        unidade_chave = df["UNIDADE_MEDIDA"].map(chave_comparacao)
        mask_oc = unidade_chave.eq("OCORRENCIAS")
        mask_vit = unidade_chave.eq("VITIMAS")
        df.loc[mask_oc, "OCORRENCIAS"] = valor[mask_oc]
        df.loc[mask_vit, "VITIMAS"] = valor[mask_vit]
        # Se a fonte vier sem aba/unidade, preserva como ocorrencia por seguranca.
        mask_sem = ~(mask_oc | mask_vit)
        df.loc[mask_sem, "OCORRENCIAS"] = valor[mask_sem]

    df["FONTE_PROCESSADA"] = f"oficial_mjsp_{tipo_base}"
    df = df.loc[:, ~df.columns.duplicated()].copy()
    return df

# -----------------------------------------------------------------------------
# EXCEL OFICIAL
# -----------------------------------------------------------------------------
def detectar_linha_cabecalho_excel(conteudo_bytes, aba):
    try:
        preview = pd.read_excel(
            io.BytesIO(conteudo_bytes),
            sheet_name=aba,
            header=None,
            nrows=20,
            dtype=str,
            engine="openpyxl",
        )
    except Exception:
        return 0

    palavras_chave = ["UF", "MUNICIPIO", "COD", "REGIAO", "ANO", "MES", "CRIME", "OCORRENCIA", "VITIMA", "SEXO"]
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
    df["SEXO_DA_VITIMA_FILTRO"] = df.get("SEXO_DA_VITIMA", pd.Series([""] * len(df))).map(chave_comparacao)
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
    where_final = where
    if where_final:
        where_final += f" AND {coluna} IS NOT NULL"
    else:
        where_final = f" WHERE {coluna} IS NOT NULL"
    con = duckdb.connect(database=":memory:")
    try:
        df = con.execute(
            f"SELECT DISTINCT {coluna} AS valor FROM read_parquet('{path_sql}') {where_final} ORDER BY valor",
            params,
        ).fetchdf()
    except Exception:
        df = pd.DataFrame()
    finally:
        con.close()
    valores = [str(x).strip() for x in df.get("valor", pd.Series(dtype=str)).dropna().tolist()]
    return sorted([v for v in valores if v and v.lower() not in ["nan", "none", "nao informado", "não informado"]])


def construir_where(uf=None, municipio=None, crime=None, ano=None, mes_num=None, unidade_medida=None, sexo=None, filtrar_unidade=True):
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
    if filtrar_unidade and unidade_medida:
        clauses.append("UNIDADE_MEDIDA_FILTRO = ?")
        params.append(chave_comparacao(unidade_medida))
    if sexo and sexo != "Todos os sexos":
        clauses.append("SEXO_DA_VITIMA_FILTRO = ?")
        params.append(chave_comparacao(sexo))
    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    return where, params


def query_dados(parquet_path, uf=None, municipio=None, crime=None, ano=None, mes_num=None, unidade_medida=None, sexo=None, limit=None, filtrar_unidade=True):
    where, params = construir_where(uf, municipio, crime, ano, mes_num, unidade_medida, sexo, filtrar_unidade)
    limite = f" LIMIT {int(limit)}" if limit else ""
    return duckdb_df(parquet_path, where + limite, params)


def diagnosticar(parquet_path, uf=None, municipio=None, crime=None, ano=None, mes_num=None, unidade_medida=None, sexo=None, filtrar_unidade=True):
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
    if filtrar_unidade and unidade_medida:
        filtros.append("UNIDADE_MEDIDA_FILTRO = ?")
        params.append(chave_comparacao(unidade_medida))
        contar(f"Apos unidade = {unidade_medida}")
    if sexo and sexo != "Todos os sexos":
        filtros.append("SEXO_DA_VITIMA_FILTRO = ?")
        params.append(chave_comparacao(sexo))
        contar(f"Apos sexo = {sexo}")

    amostra = query_dados(parquet_path, uf, municipio, crime, ano, mes_num, unidade_medida, sexo, limit=50, filtrar_unidade=filtrar_unidade)
    where_contexto, params_contexto = construir_where(uf, municipio, "Todos os indicadores", "Todos os anos", None, unidade_medida, None, filtrar_unidade)
    anos = get_distinct_options(parquet_path, "ANO_FILTRO", where_contexto, params_contexto)
    crimes = get_distinct_options(parquet_path, "TIPO_CRIME", where_contexto, params_contexto)
    sexos = get_distinct_options(parquet_path, "SEXO_DA_VITIMA", where_contexto, params_contexto)
    return {"etapas": etapas, "amostra": amostra, "anos": sorted(anos, reverse=True), "indicadores": crimes, "sexos": sexos}

# -----------------------------------------------------------------------------
# GRAFICOS E METRICAS
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
    serie = base.groupby("MES_ORDEM", dropna=True)[metrica_coluna].sum().reset_index().sort_values("MES_ORDEM")
    serie["Mes"] = serie["MES_ORDEM"].astype(int).map(lambda m: f"{int(m):02d} - {MES_LABEL.get(int(m), str(int(m)))}")
    serie_plot = serie.set_index("Mes")[[metrica_coluna]].rename(columns={metrica_coluna: metrica_label})
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
        if nivel == "municipios":
            st.info("A base oficial municipal disponibiliza Homicidio doloso como indicador. Nao ha multiplos tipos de crime nesse nivel.")
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


def preparar_resumo_sexo(df):
    """Resume a aba de vitimas por sexo, mantendo Sexo NI separado."""
    if df is None or df.empty or "SEXO_DA_VITIMA" not in df.columns or "VITIMAS" not in df.columns:
        return pd.DataFrame()

    base = df.copy()
    base["SEXO_DA_VITIMA"] = base["SEXO_DA_VITIMA"].astype(str).str.strip()
    base = base[(base["SEXO_DA_VITIMA"] != "") & (base["SEXO_DA_VITIMA"].str.lower() != "nan")]
    if base.empty:
        return pd.DataFrame()

    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    resumo = base.groupby("SEXO_DA_VITIMA", dropna=False)["VITIMAS"].sum().reset_index()
    resumo = resumo.rename(columns={"SEXO_DA_VITIMA": "Sexo da vitima", "VITIMAS": "Vitimas"})
    resumo = resumo[resumo["Vitimas"] > 0].copy()
    if resumo.empty:
        return pd.DataFrame()

    ordem = {"Masculino": 1, "Feminino": 2, "Sexo NI": 3}
    resumo["_ordem"] = resumo["Sexo da vitima"].map(ordem).fillna(99)
    resumo = resumo.sort_values(["_ordem", "Vitimas"], ascending=[True, False]).drop(columns="_ordem")
    total = resumo["Vitimas"].sum()
    resumo["Participacao (%)"] = (resumo["Vitimas"] / total * 100).round(2) if total > 0 else 0
    return resumo


def valor_sexo(resumo, sexo):
    if resumo is None or resumo.empty:
        return 0
    linha = resumo[resumo["Sexo da vitima"].astype(str).str.upper() == sexo.upper()]
    if linha.empty:
        return 0
    return linha["Vitimas"].sum()


def exibir_painel_sexo(df_vitimas, contexto_texto=""):
    resumo = preparar_resumo_sexo(df_vitimas)
    if resumo.empty:
        st.info(
            "A fonte carregada nao traz estratificacao por sexo para este recorte. "
            "No arquivo oficial por UF, essa informacao aparece na aba de vitimas; "
            "no arquivo municipal, a base disponivel nao apresenta sexo da vitima."
        )
        return

    total_vitimas = resumo["Vitimas"].sum()
    masc = valor_sexo(resumo, "Masculino")
    fem = valor_sexo(resumo, "Feminino")
    ni = valor_sexo(resumo, "Sexo NI")
    pct_ni = (ni / total_vitimas * 100) if total_vitimas else 0

    st.write("**Perfil das vitimas por sexo**")
    if contexto_texto:
        st.caption(contexto_texto)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(
            f'<div class="metric-card"><h4>👥 Total de vitimas</h4>'
            f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(total_vitimas)}</h2>'
            '<p>Aba Vitimas</p></div>',
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f'<div class="metric-card"><h4>♂ Masculino</h4>'
            f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(masc)}</h2>'
            f'<p>{(masc / total_vitimas * 100 if total_vitimas else 0):.1f}% das vitimas</p></div>',
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f'<div class="metric-card"><h4>♀ Feminino</h4>'
            f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(fem)}</h2>'
            f'<p>{(fem / total_vitimas * 100 if total_vitimas else 0):.1f}% das vitimas</p></div>',
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f'<div class="metric-card"><h4>Sexo NI</h4>'
            f'<h2 style="color:#f0ad4e; margin:0;">{formatar_inteiro(ni)}</h2>'
            f'<p>{pct_ni:.1f}% sem informacao</p></div>',
            unsafe_allow_html=True,
        )

    chart = resumo.set_index("Sexo da vitima")[["Vitimas"]]
    st.bar_chart(chart)

    with st.expander("Ver tabela de vitimas por sexo", expanded=False):
        tabela = resumo.copy()
        tabela["Vitimas"] = tabela["Vitimas"].round(0).astype(int)
        st.dataframe(tabela, width="stretch")

    st.caption(
        "Sexo NI significa sexo nao informado/nao identificado na fonte. "
        "Ele deve permanecer como categoria propria; nao deve ser redistribuido entre masculino e feminino."
    )


def plotar_sexo(df, metrica_coluna, metrica_label):
    # Mantido por compatibilidade, mas o painel novo usa exibir_painel_sexo().
    if df.empty or "SEXO_DA_VITIMA" not in df.columns:
        return
    base = df.copy()
    base = base[base["SEXO_DA_VITIMA"].astype(str).str.strip() != ""]
    if base.empty:
        return
    base[metrica_coluna] = pd.to_numeric(base[metrica_coluna], errors="coerce").fillna(0)
    serie = base.groupby("SEXO_DA_VITIMA")[metrica_coluna].sum().sort_values(ascending=False)
    serie = serie[serie > 0]
    if len(serie) > 1:
        st.write(f"**Distribuicao por sexo da vitima ({metrica_label})**")
        st.bar_chart(serie)

# -----------------------------------------------------------------------------
# SIDEBAR
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
        "Municipio usa a base oficial municipal, com homicidio doloso em vitimas. "
        "UF usa a base oficial agregada por Unidade da Federacao, com varios tipos de crime."
    ),
)
tipo_base = "municipios" if nivel_label.startswith("Municipio") else "uf"

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
if tipo_base == "municipios":
    municipios_uf = get_distinct_options(parquet_path, "MUNICIPIO", " WHERE UF_FILTRO = ?", [uf_sel])
    municipios_uf = sorted({str(m).strip().title() for m in municipios_uf if str(m).strip()})
    municipio_sel = st.sidebar.selectbox("Selecione o Municipio:", ["Todos os municipios"] + municipios_uf)
    municipio_param = None if municipio_sel == "Todos os municipios" else municipio_sel
else:
    st.sidebar.info("A base por UF e agregada por estado e nao possui recorte municipal.")

# Metrica/unidade
if tipo_base == "municipios":
    metrica_label = "Vitimas"
    metrica_coluna = "VITIMAS"
    unidade_medida = "Vitimas"
    filtrar_unidade = False  # Evita zerar municipio se a fonte vier sem campo de unidade.
    st.sidebar.caption("No nivel municipal, a fonte oficial carrega Homicidio doloso / Vitimas.")
else:
    metrica_label = st.sidebar.selectbox("Metrica:", ["Ocorrencias", "Vitimas"], index=0)
    metrica_coluna = "OCORRENCIAS" if metrica_label == "Ocorrencias" else "VITIMAS"
    unidade_medida = metrica_label
    filtrar_unidade = True

contexto_where, contexto_params = construir_where(
    uf=uf_sel,
    municipio=municipio_param,
    crime="Todos os indicadores",
    ano="Todos os anos",
    mes_num=None,
    unidade_medida=unidade_medida,
    sexo=None,
    filtrar_unidade=filtrar_unidade,
)

if tipo_base == "municipios":
    crime_sel = "Homicidio doloso"
    st.sidebar.selectbox("Indicador / Tipo de Crime:", ["Homicidio doloso"], index=0, disabled=True)
else:
    indicadores = get_distinct_options(parquet_path, "TIPO_CRIME", contexto_where, contexto_params)
    indicadores_opcoes = ["Todos os indicadores"] + indicadores
    crime_sel = st.sidebar.selectbox("Indicador / Tipo de Crime:", indicadores_opcoes)

# Sexo so faz sentido para vitimas, quando existir na fonte.
sexo_sel = "Todos os sexos"
if metrica_coluna == "VITIMAS":
    sexos_disponiveis = get_distinct_options(parquet_path, "SEXO_DA_VITIMA", contexto_where, contexto_params)
    sexos_disponiveis = [s for s in sexos_disponiveis if s.strip()]
    if sexos_disponiveis:
        sexo_sel = st.sidebar.selectbox("Sexo da vitima:", ["Todos os sexos"] + sexos_disponiveis)

anos_disponiveis = get_distinct_options(parquet_path, "ANO_FILTRO", contexto_where, contexto_params)
anos_disponiveis = sorted([a for a in anos_disponiveis if re.fullmatch(r"\d{4}", str(a))], reverse=True)
ano_opcoes = ["Todos os anos"] + anos_disponiveis
ano_sel = st.sidebar.selectbox("Ano de Referencia:", ano_opcoes, index=1 if len(ano_opcoes) > 1 else 0)

mes_nome = st.sidebar.selectbox("Mes:", list(MESES_DISPLAY.keys()), index=0)
mes_param = MESES_DISPLAY[mes_nome]

with st.sidebar.form("form_seguranca"):
    submit_btn = st.form_submit_button("🔍 Consultar Indicadores")

# -----------------------------------------------------------------------------
# TELA INICIAL
# -----------------------------------------------------------------------------
if not submit_btn:
    st.info("💡 Escolha nivel, localidade, indicador e periodo na barra lateral e clique em Consultar Indicadores.")
    st.markdown(
        """
        **Como usar:**
        
        - **Municipio:** base oficial municipal; indicador disponivel = **Homicidio doloso**; metrica = **vitimas**.
        - **UF:** base oficial agregada por Unidade da Federacao; traz varios tipos de crime e permite alternar **ocorrencias** e **vitimas**.
        - **Sexo da vitima:** quando aparece, e uma estratificacao agregada da aba de vitimas. Nao representa uma linha por pessoa.
        - O app usa **DuckDB sobre Parquet em cache versionado** para acelerar consultas depois da primeira carga.
        """
    )
    st.stop()

# -----------------------------------------------------------------------------
# CONSULTA
# -----------------------------------------------------------------------------
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
    sexo=sexo_sel,
    filtrar_unidade=filtrar_unidade,
)

if df_consulta.empty:
    st.error("🛑 Nao foram encontrados registros para a combinacao selecionada.")
    diag = diagnosticar(parquet_path, uf_sel, municipio_param, crime_sel, ano_sel, mes_param, unidade_medida, sexo_sel, filtrar_unidade)
    with st.expander("Diagnostico tecnico", expanded=True):
        st.write("**Nivel:**", tipo_base)
        st.write("**UF:**", uf_sel)
        st.write("**Municipio:**", municipio_param or "Todos")
        st.write("**Indicador:**", crime_sel)
        st.write("**Ano:**", ano_sel)
        st.write("**Mes:**", mes_nome)
        st.write("**Metrica:**", metrica_label)
        st.write("**Sexo:**", sexo_sel)
        st.dataframe(pd.DataFrame(diag["etapas"]), width="stretch")
        st.write("Anos disponiveis no contexto:", diag.get("anos", []))
        st.write("Indicadores disponiveis no contexto:", diag.get("indicadores", []))
        st.write("Sexos disponiveis no contexto:", diag.get("sexos", []))
        st.write("Metadados da base/cache:")
        st.json(meta)
        st.write("Amostra contextual:")
        if diag["amostra"].empty:
            st.info("Sem amostra para esse recorte.")
        else:
            st.dataframe(diag["amostra"], width="stretch")
    st.stop()

# Base auxiliar para perfil de vitimas por sexo.
# Importante: mesmo quando a metrica principal e Ocorrencias, o perfil por sexo
# deve vir da aba/unidade Vitimas, sem misturar linhas de ocorrencias com vitimas.
df_vitimas_sexo = pd.DataFrame()
if tipo_base == "uf":
    df_vitimas_sexo = query_dados(
        parquet_path=parquet_path,
        uf=uf_sel,
        municipio=None,
        crime=crime_sel,
        ano=ano_sel,
        mes_num=mes_param,
        unidade_medida="Vitimas",
        sexo="Todos os sexos",
        filtrar_unidade=True,
    )
elif tipo_base == "municipios":
    # A base municipal oficial nao traz sexo da vitima; mantemos para diagnostico.
    df_vitimas_sexo = df_consulta.copy()

# Cards
localidade = titulo_localidade(tipo_base, uf_sel, municipio_param)
valor_principal = soma_segura(df_consulta, metrica_coluna)
registros = len(df_consulta)

# Para cards secundarios, consulta a outra metrica de forma separada e sem misturar linhas.
if tipo_base == "uf":
    outra_label = "Vitimas" if metrica_coluna == "OCORRENCIAS" else "Ocorrencias"
    outra_col = "VITIMAS" if outra_label == "Vitimas" else "OCORRENCIAS"
    df_outra = query_dados(
        parquet_path=parquet_path,
        uf=uf_sel,
        municipio=None,
        crime=crime_sel,
        ano=ano_sel,
        mes_num=mes_param,
        unidade_medida=outra_label,
        sexo="Todos os sexos",
        filtrar_unidade=True,
    )
    valor_secundario = soma_segura(df_outra, outra_col)
else:
    outra_label = "Ocorrencias"
    valor_secundario = 0

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="metric-card"><h4>🚨 {metrica_label}</h4>'
        f'<h2 style="color:#d9534f; margin:0;">{formatar_inteiro(valor_principal)}</h2>'
        '<p>Metrica principal do recorte</p></div>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f'<div class="metric-card"><h4>📋 {outra_label}</h4>'
        f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(valor_secundario)}</h2>'
        '<p>Consulta separada da outra aba/unidade</p></div>',
        unsafe_allow_html=True,
    )
with c3:
    st.markdown(
        f'<div class="metric-card"><h4>🧾 Linhas retornadas</h4>'
        f'<h2 style="color:#f0ad4e; margin:0;">{formatar_inteiro(registros)}</h2>'
        '<p>Controle da consulta filtrada</p></div>',
        unsafe_allow_html=True,
    )
with c4:
    sexo_txt = f" | {sexo_sel}" if sexo_sel != "Todos os sexos" else ""
    st.markdown(
        f'<div class="metric-card"><h4>📍 Localidade</h4>'
        f'<h3 style="color:#1c2d42; margin:0;">{localidade}</h3>'
        f'<p>{ano_sel} | {mes_nome}{sexo_txt}</p></div>',
        unsafe_allow_html=True,
    )

if tipo_base == "municipios":
    st.info("Leitura metodologica: no nivel municipal, a base oficial disponivel corresponde a Homicidio doloso e unidade principal em Vitimas. Para mais tipos de crime, use o nivel UF; para crimes por municipio, seria necessario integrar bases estaduais especificas.")
else:
    st.info("Leitura metodologica: este e um agregado estadual. A base por UF traz varios tipos de crime. Quando a metrica e Vitimas, o campo sexo da vitima e uma estratificacao agregada da propria fonte.")

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
        plotar_crimes(df_consulta, metrica_coluna, metrica_label, tipo_base)

    st.markdown("---")
    if tipo_base == "uf":
        contexto_sexo = (
            "Painel calculado exclusivamente com a aba/unidade Vitimas. "
            "Ele nao deve ser somado nem comparado linha a linha com a aba de Ocorrencias."
        )
        exibir_painel_sexo(df_vitimas_sexo, contexto_texto=contexto_sexo)
    else:
        st.info(
            "A base oficial municipal disponivel neste recurso nao traz estratificacao por sexo da vitima. "
            "Para sexo da vitima, use o nivel UF, que consulta a aba de Vitimas do arquivo estadual."
        )

    st.markdown("---")
    if tipo_base == "municipios":
        st.write(f"**Ranking municipal - top 15 ({metrica_label})**")
        plotar_top_municipios(df_consulta, metrica_coluna, metrica_label, municipio_param)
    else:
        st.info("Ranking municipal nao se aplica ao nivel UF, pois essa base nao possui recorte por municipio.")

with tab2:
    colunas_preferenciais = [
        "ANO", "MES_NOME", "UF", "MUNICIPIO", "TIPO_CRIME", "UNIDADE_MEDIDA",
        "OCORRENCIAS", "VITIMAS", "VITIMAS_MUNICIPIO", "SEXO_DA_VITIMA", "FONTE_PROCESSADA",
    ]
    colunas_existentes = [c for c in colunas_preferenciais if c in df_consulta.columns]
    outras = [c for c in df_consulta.columns if c not in colunas_existentes]
    st.dataframe(df_consulta[colunas_existentes + outras], width="stretch")

with tab3:
    diag_ok = diagnosticar(parquet_path, uf_sel, municipio_param, crime_sel, ano_sel, mes_param, unidade_medida, sexo_sel, filtrar_unidade)
    st.write("**Contagem por etapa do filtro:**")
    st.dataframe(pd.DataFrame(diag_ok["etapas"]), width="stretch")
    st.write("**Anos disponiveis no contexto:**", diag_ok.get("anos", []))
    st.write("**Indicadores disponiveis no contexto:**", diag_ok.get("indicadores", []))
    st.write("**Sexos disponiveis no contexto:**", diag_ok.get("sexos", []))
    st.write("**Metadados da base/cache:**")
    st.json(meta)
    st.write("**Colunas finais:**")
    st.code("\n".join(df_consulta.columns.astype(str).tolist()))
    st.write("**Amostra da consulta filtrada:**")
    st.dataframe(df_consulta.head(50), width="stretch")

with tab4:
    nome_base = f"seguranca_{tipo_base}_{uf_sel}_{municipio_param or 'todos'}_{ano_sel}_{mes_param or 'todos'}_{metrica_label}_{sexo_sel}.csv"
    nome_base = normalizar_coluna(nome_base).lower().replace("_csv", ".csv")
    csv = df_consulta.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
    st.download_button(
        "📥 Baixar CSV tratado",
        data=csv,
        file_name=nome_base,
        mime="text/csv",
    )

    resumo_sexo_export = preparar_resumo_sexo(df_vitimas_sexo)
    if not resumo_sexo_export.empty:
        csv_sexo = resumo_sexo_export.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
        nome_sexo = f"seguranca_resumo_sexo_{uf_sel}_{ano_sel}_{mes_param or 'todos'}_{crime_sel}.csv"
        nome_sexo = normalizar_coluna(nome_sexo).lower().replace("_csv", ".csv")
        st.download_button(
            "📥 Baixar resumo de vitimas por sexo",
            data=csv_sexo,
            file_name=nome_sexo,
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
