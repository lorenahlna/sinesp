# VERSION_FINAL_SEGURANCA_SINESP_OFICIAL_DUCKDB_V16_ROBUSTO
# Fonte unica: MJSP/SINESP - Dados Nacionais de Seguranca Publica
# Regras metodologicas:
# 1) MUNICIPIO: base oficial municipal = Homicidio doloso / Vitimas.
# 2) UF: base oficial estadual = Ocorrencias por varios crimes + Vitimas por sexo/tipo de crime.
# 3) A analise de sexo sempre usa a estrutura estadual de VITIMAS, independentemente da metrica principal selecionada.

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
    .metric-card {
        background-color: #f8f9fa;
        padding: 15px;
        border-radius: 10px;
        border-left: 5px solid #1c2d42;
        text-align: center;
        margin-bottom: 15px;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
        min-height: 145px;
    }
    .metric-card h4 { margin-top: 0; color: #0b2239; font-size: 0.95rem; }
    .metric-card h2 { margin: 0; font-size: 1.55rem; }
    .method-box {
        background-color: #eef3f8;
        border-left: 5px solid #1c2d42;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
    .warning-box {
        background-color: #fff8e8;
        border-left: 5px solid #f0ad4e;
        padding: 12px;
        border-radius: 6px;
        margin-bottom: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

CACHE_SCHEMA_VERSION = "v16_robusto_20260630"

URL_MJSP_MUNICIPIOS = (
    "https://dados.mj.gov.br/dataset/210b9ae2-21fc-4986-89c6-2006eb4db247/"
    "resource/03af7ce2-174e-4ebd-b085-384503cfb40f/download/indicadoressegurancapublicamunic.xlsx"
)
URL_MJSP_UF = (
    "https://dados.mj.gov.br/dataset/210b9ae2-21fc-4986-89c6-2006eb4db247/"
    "resource/feeae05e-faba-406c-8a4a-512aec91a9d1/download/indicadoressegurancapublicauf.xlsx"
)

UFS = [
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA", "MT", "MS", "MG", "PA", "PB",
    "PR", "PE", "PI", "RJ", "RN", "RS", "RO", "RR", "SC", "SP", "SE", "TO",
]

UF_NOMES = {
    "AC": "Acre", "AL": "Alagoas", "AP": "Amapa", "AM": "Amazonas", "BA": "Bahia", "CE": "Ceara",
    "DF": "Distrito Federal", "ES": "Espirito Santo", "GO": "Goias", "MA": "Maranhao", "MT": "Mato Grosso",
    "MS": "Mato Grosso do Sul", "MG": "Minas Gerais", "PA": "Para", "PB": "Paraiba", "PR": "Parana",
    "PE": "Pernambuco", "PI": "Piaui", "RJ": "Rio de Janeiro", "RN": "Rio Grande do Norte",
    "RS": "Rio Grande do Sul", "RO": "Rondonia", "RR": "Roraima", "SC": "Santa Catarina",
    "SP": "Sao Paulo", "SE": "Sergipe", "TO": "Tocantins",
}
NOME_UF_NORMALIZADO = {v.upper(): k for k, v in UF_NOMES.items()}

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
MES_ORDEM = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}
MES_LABEL = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}

# -----------------------------------------------------------------------------
# NORMALIZACAO
# -----------------------------------------------------------------------------
def remover_acentos(texto):
    return unicodedata.normalize("NFKD", str(texto)).encode("ASCII", "ignore").decode("ASCII")


def normalizar_texto(texto):
    if pd.isna(texto):
        return ""
    texto = remover_acentos(str(texto)).upper().strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def chave_filtro(texto):
    return re.sub(r"[^A-Z0-9]+", "", normalizar_texto(texto))


def normalizar_coluna(coluna):
    col = normalizar_texto(coluna).replace(".", "")
    col = re.sub(r"[^A-Z0-9]+", "_", col)
    return col.strip("_")


def converter_numero(valor):
    if pd.isna(valor):
        return 0.0
    s = str(valor).strip()
    if not s:
        return 0.0
    if re.search(r"\d+[,]\d+$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    return float(pd.to_numeric(s, errors="coerce") or 0)


def parse_data_mes_ano(valor):
    if pd.isna(valor):
        return pd.NaT
    s = str(valor).strip()
    if not s:
        return pd.NaT
    formatos = [
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y",
        "%m/%Y",
        "%Y/%m",
    ]
    for fmt in formatos:
        try:
            return pd.to_datetime(s, format=fmt, errors="raise")
        except Exception:
            pass
    return pd.to_datetime(s, errors="coerce", dayfirst=False)


def uf_para_sigla(valor):
    if pd.isna(valor):
        return ""
    v = normalizar_texto(valor)
    if v in UFS:
        return v
    return NOME_UF_NORMALIZADO.get(v, "")


def padronizar_municipio(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().title()


def padronizar_sexo(valor):
    s = normalizar_texto(valor)
    if not s or s in ["NAN", "NONE", "NULL"]:
        return "Nao informado"
    if "FEM" in s:
        return "Feminino"
    if "MASC" in s:
        return "Masculino"
    if "SEXO NI" in s or s == "NI" or "NAO INFORM" in s or "NAO IDENT" in s or "IGN" in s or "INDETER" in s:
        return "Sexo NI"
    return str(valor).strip().title()


def localizar_coluna(df, candidatos):
    mapa = {normalizar_coluna(c): c for c in df.columns}
    for cand in candidatos:
        c = normalizar_coluna(cand)
        if c in mapa:
            return mapa[c]
    return None

# -----------------------------------------------------------------------------
# CACHE / DOWNLOAD
# -----------------------------------------------------------------------------
def pasta_cache():
    path = os.path.join(tempfile.gettempdir(), "sinesp_mjsp_v16_cache")
    os.makedirs(path, exist_ok=True)
    return path


def caminho_parquet(tipo):
    return os.path.join(pasta_cache(), f"sinesp_{tipo}_{CACHE_SCHEMA_VERSION}.parquet")


def baixar_xlsx(url, nome_local):
    if os.path.exists(nome_local):
        with open(nome_local, "rb") as f:
            return f.read(), f"arquivo local: {nome_local}"
    resp = requests.get(url, timeout=240, verify=False, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.content, url


def adicionar_colunas_filtro(df):
    df = df.copy()
    colunas_base = [
        "CODIGO_MUNICIPIO", "MUNICIPIO", "UF", "REGIAO", "ANO", "MES_ORDEM", "MES_NOME", "MES_ANO",
        "TIPO_CRIME", "UNIDADE_MEDIDA", "SEXO_DA_VITIMA", "OCORRENCIAS", "VITIMAS", "NIVEL_BASE",
        "METRICA_FONTE", "ABA_ORIGEM", "FONTE_PROCESSADA",
    ]
    for c in colunas_base:
        if c not in df.columns:
            df[c] = ""

    df["UF"] = df["UF"].map(uf_para_sigla)
    df["UF_FILTRO"] = df["UF"].astype(str).str.upper().str.strip()
    df["MUNICIPIO_FILTRO"] = df["MUNICIPIO"].map(chave_filtro)
    df["TIPO_CRIME_FILTRO"] = df["TIPO_CRIME"].map(chave_filtro)
    df["UNIDADE_MEDIDA_FILTRO"] = df["UNIDADE_MEDIDA"].map(chave_filtro)
    df["SEXO_DA_VITIMA_FILTRO"] = df["SEXO_DA_VITIMA"].map(chave_filtro)
    df["METRICA_FONTE_FILTRO"] = df["METRICA_FONTE"].map(chave_filtro)
    df["ANO_FILTRO"] = df["ANO"].astype(str).str.extract(r"(\d{4})", expand=False).fillna("")
    df["MES_ORDEM_NUM"] = pd.to_numeric(df["MES_ORDEM"], errors="coerce")
    df["OCORRENCIAS"] = pd.to_numeric(df["OCORRENCIAS"], errors="coerce").fillna(0).astype(float)
    df["VITIMAS"] = pd.to_numeric(df["VITIMAS"], errors="coerce").fillna(0).astype(float)

    colunas = colunas_base + [
        "UF_FILTRO", "MUNICIPIO_FILTRO", "TIPO_CRIME_FILTRO", "UNIDADE_MEDIDA_FILTRO",
        "SEXO_DA_VITIMA_FILTRO", "METRICA_FONTE_FILTRO", "ANO_FILTRO", "MES_ORDEM_NUM",
    ]
    return df[colunas]


def preparar_base_municipal(conteudo):
    xls = pd.ExcelFile(io.BytesIO(conteudo), engine="openpyxl")
    frames = []
    for aba in xls.sheet_names:
        df_raw = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, dtype=str, engine="openpyxl")
        df_raw = df_raw.dropna(how="all")
        if df_raw.empty:
            continue
        df_raw.columns = [normalizar_coluna(c) for c in df_raw.columns]

        col_mun = localizar_coluna(df_raw, ["MUNICIPIO", "MUNICÍPIO", "NOME_MUNICIPIO"])
        col_uf = localizar_coluna(df_raw, ["UF", "SIGLA_UF", "ESTADO"])
        col_data = localizar_coluna(df_raw, ["MES_ANO", "MES/ANO", "MESANO", "PERIODO"])
        col_vit = localizar_coluna(df_raw, ["VITIMAS", "VÍTIMAS", "VALOR"])
        col_cod = localizar_coluna(df_raw, ["COD_IBGE", "CODIGO_IBGE", "CODIGO_MUNICIPIO", "COD MUNICIPIO"])
        col_reg = localizar_coluna(df_raw, ["REGIAO", "REGIÃO"])

        if not col_mun or not col_vit:
            continue
        if not col_uf and normalizar_texto(aba) not in UFS:
            continue

        df = pd.DataFrame()
        df["MUNICIPIO"] = df_raw[col_mun].map(padronizar_municipio)
        df["UF"] = df_raw[col_uf].map(uf_para_sigla) if col_uf else normalizar_texto(aba)
        df["CODIGO_MUNICIPIO"] = df_raw[col_cod].astype(str).str.strip() if col_cod else ""
        df["REGIAO"] = df_raw[col_reg].astype(str).str.strip() if col_reg else ""
        df["MES_ANO"] = df_raw[col_data].astype(str).str.strip() if col_data else ""
        df["VITIMAS"] = df_raw[col_vit].map(converter_numero).fillna(0).astype(float)

        if col_data:
            dt = df["MES_ANO"].map(parse_data_mes_ano)
            df["ANO"] = dt.dt.year.astype("Int64").astype(str).replace("<NA>", "")
            df["MES_ORDEM"] = dt.dt.month.astype("Int64")
        else:
            col_ano = localizar_coluna(df_raw, ["ANO"])
            col_mes = localizar_coluna(df_raw, ["MES"])
            df["ANO"] = df_raw[col_ano].astype(str).str.extract(r"(\d{4})", expand=False).fillna("") if col_ano else ""
            df["MES_ORDEM"] = df_raw[col_mes].map(lambda x: MES_ORDEM.get(normalizar_texto(x), pd.NA)) if col_mes else pd.NA

        df["MES_NOME"] = df["MES_ORDEM"].map(lambda x: MES_LABEL.get(int(x), "Nao informado") if pd.notna(x) else "Nao informado")
        df["TIPO_CRIME"] = "Homicidio doloso"
        df["UNIDADE_MEDIDA"] = "Vitimas"
        df["SEXO_DA_VITIMA"] = "Nao informado na base municipal"
        df["OCORRENCIAS"] = 0.0
        df["NIVEL_BASE"] = "municipio"
        df["METRICA_FONTE"] = "VITIMAS"
        df["ABA_ORIGEM"] = "MUNICIPIOS"
        df["FONTE_PROCESSADA"] = "MJSP/SINESP - Municipios XLSX"
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    return adicionar_colunas_filtro(pd.concat(frames, ignore_index=True, sort=False))


def inferir_tipo_aba(nome_aba, df_raw):
    aba_norm = normalizar_texto(nome_aba)
    cols_norm = {normalizar_coluna(c) for c in df_raw.columns}
    if "VIT" in aba_norm or "VITIMAS" in cols_norm or "SEXO_DA_VITIMA" in cols_norm or "SEXO_VITIMA" in cols_norm:
        return "VITIMAS"
    if "OCOR" in aba_norm or "OCORRENCIAS" in cols_norm:
        return "OCORRENCIAS"
    return ""


def preparar_base_uf(conteudo):
    xls = pd.ExcelFile(io.BytesIO(conteudo), engine="openpyxl")
    frames = []
    for aba in xls.sheet_names:
        df_raw = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, dtype=str, engine="openpyxl")
        df_raw = df_raw.dropna(how="all")
        if df_raw.empty:
            continue
        df_raw.columns = [normalizar_coluna(c) for c in df_raw.columns]
        tipo_aba = inferir_tipo_aba(aba, df_raw)
        if tipo_aba not in ["OCORRENCIAS", "VITIMAS"]:
            continue

        col_uf = localizar_coluna(df_raw, ["UF", "SIGLA_UF", "ESTADO"])
        col_crime = localizar_coluna(df_raw, ["TIPO_CRIME", "TIPO_DE_CRIME", "CRIME", "INDICADOR"])
        col_ano = localizar_coluna(df_raw, ["ANO"])
        col_mes = localizar_coluna(df_raw, ["MES"])
        col_sexo = localizar_coluna(df_raw, ["SEXO_DA_VITIMA", "SEXO_VITIMA", "SEXO"])
        if tipo_aba == "OCORRENCIAS":
            col_valor = localizar_coluna(df_raw, ["OCORRENCIAS", "OCORRÊNCIAS", "VALOR"])
            metrica = "OCORRENCIAS"
            unidade = "Ocorrencias"
        else:
            col_valor = localizar_coluna(df_raw, ["VITIMAS", "VÍTIMAS", "VALOR"])
            metrica = "VITIMAS"
            unidade = "Vitimas"

        if not all([col_uf, col_crime, col_ano, col_mes, col_valor]):
            continue

        df = pd.DataFrame()
        df["UF"] = df_raw[col_uf].map(uf_para_sigla)
        df["TIPO_CRIME"] = df_raw[col_crime].astype(str).str.strip().replace({"": "Nao informado"})
        df["ANO"] = df_raw[col_ano].astype(str).str.extract(r"(\d{4})", expand=False).fillna("")
        df["MES_ORDEM"] = df_raw[col_mes].map(lambda x: MES_ORDEM.get(normalizar_texto(x), pd.NA))
        df["MES_NOME"] = df["MES_ORDEM"].map(lambda x: MES_LABEL.get(int(x), "Nao informado") if pd.notna(x) else "Nao informado")
        df["MES_ANO"] = ""
        df["MUNICIPIO"] = ""
        df["CODIGO_MUNICIPIO"] = ""
        df["REGIAO"] = ""
        valores = df_raw[col_valor].map(converter_numero).fillna(0).astype(float)
        if metrica == "OCORRENCIAS":
            df["OCORRENCIAS"] = valores
            df["VITIMAS"] = 0.0
            df["SEXO_DA_VITIMA"] = "Nao se aplica"
        else:
            df["OCORRENCIAS"] = 0.0
            df["VITIMAS"] = valores
            df["SEXO_DA_VITIMA"] = df_raw[col_sexo].map(padronizar_sexo) if col_sexo else "Nao informado"
        df["UNIDADE_MEDIDA"] = unidade
        df["NIVEL_BASE"] = "uf"
        df["METRICA_FONTE"] = metrica
        df["ABA_ORIGEM"] = metrica
        df["FONTE_PROCESSADA"] = "MJSP/SINESP - UF XLSX"
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    return adicionar_colunas_filtro(pd.concat(frames, ignore_index=True, sort=False))


@st.cache_data(ttl=86400, show_spinner=False)
def preparar_parquet(tipo):
    parquet_path = caminho_parquet(tipo)
    try:
        if os.path.exists(parquet_path):
            return {"ok": True, "parquet_path": parquet_path, "cache": True, "tipo": tipo, "schema": CACHE_SCHEMA_VERSION}
        if tipo == "municipios":
            conteudo, origem = baixar_xlsx(URL_MJSP_MUNICIPIOS, "indicadoressegurancapublicamunic.xlsx")
            df = preparar_base_municipal(conteudo)
        else:
            conteudo, origem = baixar_xlsx(URL_MJSP_UF, "indicadoressegurancapublicauf.xlsx")
            df = preparar_base_uf(conteudo)
        if df.empty:
            return {"ok": False, "erro": "A base foi lida, mas nenhum registro valido foi identificado.", "tipo": tipo}
        df.to_parquet(parquet_path, index=False)
        return {
            "ok": True,
            "parquet_path": parquet_path,
            "cache": False,
            "tipo": tipo,
            "schema": CACHE_SCHEMA_VERSION,
            "origem": origem,
            "linhas": int(len(df)),
            "colunas": df.columns.tolist(),
            "anos": sorted(df["ANO_FILTRO"].dropna().unique().tolist()),
            "metricas": sorted(df["METRICA_FONTE"].dropna().unique().tolist()),
        }
    except Exception as e:
        return {"ok": False, "erro": str(e), "tipo": tipo, "schema": CACHE_SCHEMA_VERSION}

# -----------------------------------------------------------------------------
# DUCKDB
# -----------------------------------------------------------------------------
def build_where(tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, metrica=None, sexo=None):
    clauses = []
    params = []
    if uf:
        clauses.append("UF_FILTRO = ?")
        params.append(uf)
    if tipo == "municipios" and municipio and municipio != "Todos os municipios":
        clauses.append("MUNICIPIO_FILTRO = ?")
        params.append(chave_filtro(municipio))
    if tipo != "municipios" and crime and crime not in ["Todos os indicadores", "Todos os tipos de crime"]:
        clauses.append("TIPO_CRIME_FILTRO = ?")
        params.append(chave_filtro(crime))
    if ano and ano != "Todos os anos":
        clauses.append("ANO_FILTRO = ?")
        params.append(str(ano))
    if mes is not None:
        clauses.append("MES_ORDEM_NUM = ?")
        params.append(int(mes))
    if tipo != "municipios" and metrica:
        clauses.append("METRICA_FONTE_FILTRO = ?")
        params.append(chave_filtro(metrica))
    if tipo != "municipios" and sexo and sexo != "Todos os sexos":
        clauses.append("SEXO_DA_VITIMA_FILTRO = ?")
        params.append(chave_filtro(sexo))
    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    return where, params


def query_df(path, tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, metrica=None, sexo=None, only_positive=False, limit=None):
    where, params = build_where(tipo, uf, municipio, crime, ano, mes, metrica, sexo)
    extra = ""
    if only_positive:
        if metrica == "VITIMAS" or tipo == "municipios":
            extra = " AND VITIMAS > 0" if where else " WHERE VITIMAS > 0"
        elif metrica == "OCORRENCIAS":
            extra = " AND OCORRENCIAS > 0" if where else " WHERE OCORRENCIAS > 0"
    lim = f" LIMIT {int(limit)}" if limit else ""
    sql = f"SELECT * FROM read_parquet(?) {where}{extra}{lim}"
    return duckdb.execute(sql, [path] + params).df()


def distinct_values(path, coluna, tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, metrica=None, sexo=None):
    where, params = build_where(tipo, uf, municipio, crime, ano, mes, metrica, sexo)
    sql = f"SELECT DISTINCT {coluna} AS valor FROM read_parquet(?) {where} ORDER BY valor"
    try:
        df = duckdb.execute(sql, [path] + params).df()
        return [str(v) for v in df["valor"].dropna().tolist() if str(v).strip() and str(v) != "<NA>"]
    except Exception:
        return []


def count_rows(path, tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, metrica=None, sexo=None):
    where, params = build_where(tipo, uf, municipio, crime, ano, mes, metrica, sexo)
    sql = f"SELECT COUNT(*) AS n FROM read_parquet(?) {where}"
    return int(duckdb.execute(sql, [path] + params).df()["n"].iloc[0])

# -----------------------------------------------------------------------------
# VISUALIZACAO
# -----------------------------------------------------------------------------
def fmt_int(valor):
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")
    except Exception:
        return "0"


def fmt_pct(valor):
    try:
        return f"{float(valor):.1f}%".replace(".", ",")
    except Exception:
        return "0,0%"


def metric_card(titulo, valor, subtitulo, cor="#1c2d42"):
    st.markdown(
        f"""
        <div class="metric-card">
            <h4>{titulo}</h4>
            <h2 style="color:{cor};">{valor}</h2>
            <p>{subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def serie_mensal(df, metrica):
    if df.empty:
        return pd.Series(dtype=float)
    base = df.copy()
    base[metrica] = pd.to_numeric(base[metrica], errors="coerce").fillna(0)
    grp = base.groupby(["MES_ORDEM_NUM", "MES_NOME"], dropna=False)[metrica].sum().reset_index()
    grp = grp.sort_values("MES_ORDEM_NUM")
    return pd.Series(grp[metrica].values, index=grp["MES_NOME"].values)


def serie_por_coluna(df, coluna, metrica, top=15):
    if df.empty or coluna not in df.columns:
        return pd.Series(dtype=float)
    base = df.copy()
    base[metrica] = pd.to_numeric(base[metrica], errors="coerce").fillna(0)
    s = base.groupby(coluna)[metrica].sum().sort_values(ascending=False)
    return s[s > 0].head(top)


def resumo_sexo(df_vitimas):
    if df_vitimas.empty or "SEXO_DA_VITIMA" not in df_vitimas.columns:
        return pd.DataFrame(columns=["SEXO_DA_VITIMA", "VITIMAS", "PARTICIPACAO_%"])
    base = df_vitimas.copy()
    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    base = base[base["VITIMAS"] > 0]
    base = base[~base["SEXO_DA_VITIMA"].isin(["", "Nao se aplica", "Nao informado na base municipal"])]
    if base.empty:
        return pd.DataFrame(columns=["SEXO_DA_VITIMA", "VITIMAS", "PARTICIPACAO_%"])
    tab = base.groupby("SEXO_DA_VITIMA", as_index=False)["VITIMAS"].sum().sort_values("VITIMAS", ascending=False)
    total = tab["VITIMAS"].sum()
    tab["PARTICIPACAO_%"] = (tab["VITIMAS"] / total * 100).round(2) if total else 0
    return tab


def resumo_vitimas_por_crime(df_vitimas, top=15):
    if df_vitimas.empty:
        return pd.DataFrame(columns=["TIPO_CRIME", "VITIMAS", "PARTICIPACAO_%"])
    base = df_vitimas.copy()
    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    base = base[base["VITIMAS"] > 0]
    if base.empty:
        return pd.DataFrame(columns=["TIPO_CRIME", "VITIMAS", "PARTICIPACAO_%"])
    tab = base.groupby("TIPO_CRIME", as_index=False)["VITIMAS"].sum().sort_values("VITIMAS", ascending=False).head(top)
    total = tab["VITIMAS"].sum()
    tab["PARTICIPACAO_%"] = (tab["VITIMAS"] / total * 100).round(2) if total else 0
    return tab


def tabela_crime_sexo(df_vitimas, top_crimes=12):
    if df_vitimas.empty:
        return pd.DataFrame()
    base = df_vitimas.copy()
    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    base = base[base["VITIMAS"] > 0]
    base = base[~base["SEXO_DA_VITIMA"].isin(["", "Nao se aplica", "Nao informado na base municipal"])]
    if base.empty:
        return pd.DataFrame()
    crimes_top = base.groupby("TIPO_CRIME")["VITIMAS"].sum().sort_values(ascending=False).head(top_crimes).index.tolist()
    base = base[base["TIPO_CRIME"].isin(crimes_top)]
    piv = pd.pivot_table(
        base,
        index="TIPO_CRIME",
        columns="SEXO_DA_VITIMA",
        values="VITIMAS",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    sexo_cols = [c for c in piv.columns if c != "TIPO_CRIME"]
    piv["TOTAL_VITIMAS"] = piv[sexo_cols].sum(axis=1) if sexo_cols else 0
    return piv.sort_values("TOTAL_VITIMAS", ascending=False)

# -----------------------------------------------------------------------------
# SIDEBAR
# -----------------------------------------------------------------------------
st.sidebar.title("🛡️ Filtros de Seguranca")

nivel = st.sidebar.radio(
    "Nivel de analise:",
    ["Municipio - homicidio doloso/vitimas", "UF - varios tipos de crime"],
    index=0,
)
tipo = "municipios" if nivel.startswith("Municipio") else "uf"

with st.spinner("Preparando base oficial MJSP/SINESP em cache DuckDB..."):
    meta = preparar_parquet(tipo)

if not meta.get("ok"):
    st.error("Nao foi possivel preparar a base oficial MJSP/SINESP.")
    st.json(meta)
    st.stop()

path = meta["parquet_path"]
uf_sel = st.sidebar.selectbox("Estado:", sorted(UFS), index=sorted(UFS).index("MG"))
municipio_sel = "Todos os municipios"
crime_sel = "Todos os indicadores"
sexo_sel = "Todos os sexos"

if tipo == "municipios":
    municipios = distinct_values(path, "MUNICIPIO", tipo, uf=uf_sel)
    municipios = sorted(set([m for m in municipios if str(m).strip()]))
    municipio_sel = st.sidebar.selectbox("Municipio:", ["Todos os municipios"] + municipios)
    metrica_label = "Vitimas"
    metrica_coluna = "VITIMAS"
    metrica_fonte = "VITIMAS"
    crime_sel = "Homicidio doloso"
    st.sidebar.markdown("**Indicador fixo:** Homicidio doloso")
    st.sidebar.markdown("**Unidade:** Vitimas")
    st.sidebar.info("A base municipal oficial nao traz todos os tipos de crime. Ela representa homicidio doloso medido em vitimas.")
else:
    metrica_label = st.sidebar.selectbox("Metrica principal:", ["Ocorrencias", "Vitimas"], index=0)
    metrica_coluna = "OCORRENCIAS" if metrica_label == "Ocorrencias" else "VITIMAS"
    metrica_fonte = "OCORRENCIAS" if metrica_label == "Ocorrencias" else "VITIMAS"
    crimes = distinct_values(path, "TIPO_CRIME", tipo, uf=uf_sel, metrica=metrica_fonte)
    crimes = sorted(set([c for c in crimes if str(c).strip()]))
    crime_sel = st.sidebar.selectbox("Indicador/tipo de crime:", ["Todos os indicadores"] + crimes)
    if metrica_fonte == "VITIMAS":
        sexos = distinct_values(path, "SEXO_DA_VITIMA", tipo, uf=uf_sel, crime=crime_sel, metrica="VITIMAS")
        sexos = sorted(set([s for s in sexos if str(s).strip() and s != "Nao se aplica"]))
        if sexos:
            sexo_sel = st.sidebar.selectbox("Sexo da vitima:", ["Todos os sexos"] + sexos)
    st.sidebar.info("A base UF possui ocorrencias por tipo de crime e vitimas por sexo/tipo de crime.")

anos = distinct_values(
    path,
    "ANO_FILTRO",
    tipo,
    uf=uf_sel,
    municipio=municipio_sel if tipo == "municipios" else None,
    crime=crime_sel if tipo != "municipios" else None,
    metrica=metrica_fonte if tipo != "municipios" else None,
)
anos = sorted([a for a in anos if re.fullmatch(r"\d{4}", str(a))], reverse=True)
ano_sel = st.sidebar.selectbox("Ano:", ["Todos os anos"] + anos, index=1 if anos else 0)

mes_nome = st.sidebar.selectbox("Mes:", list(MESES_DISPLAY.keys()), index=0)
mes_sel = MESES_DISPLAY[mes_nome]

with st.sidebar.form("form_seguranca"):
    consultar = st.form_submit_button("🔍 Consultar Indicadores")

# -----------------------------------------------------------------------------
# TELA INICIAL
# -----------------------------------------------------------------------------
if not consultar:
    st.info("Escolha os filtros na barra lateral e clique em Consultar Indicadores.")
    st.markdown(
        """
        <div class="method-box">
        <b>Leitura metodologica:</b><br>
        <b>Municipio:</b> somente homicidio doloso, medido em vitimas. Nao ha sexo da vitima nem demais crimes no XLSX municipal.<br>
        <b>UF:</b> varios tipos de crime em ocorrencias; vitimas por sexo e tipo de crime na estrutura estadual de vitimas.<br>
        <b>Sexo NI:</b> sexo nao informado/nao identificado, mantido como categoria propria.
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.stop()

# -----------------------------------------------------------------------------
# CONSULTA
# -----------------------------------------------------------------------------
st.markdown(
    '<div class="header-seguranca"><h1>Painel de Ocorrencias Criminais - SINESP/MJSP</h1>'
    '<p>Indicadores oficiais de seguranca publica com filtro territorial</p></div>',
    unsafe_allow_html=True,
)

if tipo == "municipios":
    df_main = query_df(path, tipo, uf=uf_sel, municipio=municipio_sel, ano=ano_sel, mes=mes_sel, only_positive=False)
    df_vitimas = pd.DataFrame()
else:
    df_main = query_df(path, tipo, uf=uf_sel, crime=crime_sel, ano=ano_sel, mes=mes_sel, metrica=metrica_fonte, sexo=sexo_sel, only_positive=False)
    # Analise complementar de vitimas: nunca depende de ABA_ORIGEM acentuada; usa METRICA_FONTE = VITIMAS.
    df_vitimas = query_df(path, tipo, uf=uf_sel, crime=crime_sel, ano=ano_sel, mes=mes_sel, metrica="VITIMAS", only_positive=False)

if df_main.empty:
    st.warning("Nao foram encontrados registros para a combinacao selecionada.")
    st.json({
        "nivel": tipo,
        "uf": uf_sel,
        "municipio": municipio_sel if tipo == "municipios" else None,
        "crime": crime_sel,
        "ano": ano_sel,
        "mes": mes_nome,
        "metrica_fonte": metrica_fonte if tipo != "municipios" else "VITIMAS",
        "linhas_uf": count_rows(path, tipo, uf=uf_sel),
        "linhas_uf_municipio": count_rows(path, tipo, uf=uf_sel, municipio=municipio_sel) if tipo == "municipios" else None,
        "metricas_disponiveis": distinct_values(path, "METRICA_FONTE", tipo, uf=uf_sel),
        "abas_disponiveis": distinct_values(path, "ABA_ORIGEM", tipo, uf=uf_sel),
        "anos_disponiveis_no_contexto": anos,
        "cache": meta,
    })
    st.stop()

# Totais principais
total_principal = pd.to_numeric(df_main[metrica_coluna], errors="coerce").fillna(0).sum()
if tipo == "uf" and df_vitimas.empty:
    # Fallback defensivo: se por algum motivo o filtro de metrica falhar, busca linhas com VITIMAS > 0 no contexto.
    df_vitimas = query_df(path, tipo, uf=uf_sel, crime=crime_sel, ano=ano_sel, mes=mes_sel, metrica=None, only_positive=True)

total_vitimas_complementar = pd.to_numeric(df_vitimas["VITIMAS"], errors="coerce").fillna(0).sum() if not df_vitimas.empty else 0
linhas = len(df_main)
localidade = f"{municipio_sel} - {uf_sel}" if tipo == "municipios" and municipio_sel != "Todos os municipios" else (f"Municipios de {uf_sel}" if tipo == "municipios" else f"Estado {uf_sel}")
periodo = f"{ano_sel} | {mes_nome}"

sexo_tab = resumo_sexo(df_vitimas) if tipo == "uf" else pd.DataFrame()
vitimas_crime_tab = resumo_vitimas_por_crime(df_vitimas, top=15) if tipo == "uf" else pd.DataFrame()
crime_sexo_tab = tabela_crime_sexo(df_vitimas, top_crimes=12) if tipo == "uf" else pd.DataFrame()
sexo_dict = dict(zip(sexo_tab.get("SEXO_DA_VITIMA", []), sexo_tab.get("VITIMAS", []))) if not sexo_tab.empty else {}
crime_top = vitimas_crime_tab.iloc[0]["TIPO_CRIME"] if not vitimas_crime_tab.empty else "Nao informado"
crime_top_v = vitimas_crime_tab.iloc[0]["VITIMAS"] if not vitimas_crime_tab.empty else 0

# Cards principais
if tipo == "municipios":
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("👥 Vitimas de homicidio doloso", fmt_int(total_principal), "Metrica municipal oficial", "#d9534f")
    with c2:
        metric_card("📌 Indicador", "Homicidio doloso", "Indicador fixo na base municipal", "#0b2239")
    with c3:
        metric_card("📋 Linhas", fmt_int(linhas), "Registros retornados", "#f0ad4e")
    with c4:
        metric_card("📍 Localidade", localidade, periodo, "#1c2d42")
    st.markdown(
        """
        <div class="warning-box">
        <b>Base municipal:</b> somente Homicidio doloso, medido em Vitimas. A base municipal oficial nao possui todos os tipos de crime nem sexo da vitima.
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card(f"🚨 {metrica_label}", fmt_int(total_principal), "Metrica principal selecionada", "#d9534f")
    with c2:
        metric_card("👥 Vitimas", fmt_int(total_vitimas_complementar), "Estrutura estadual de vitimas", "#0b2239")
    with c3:
        metric_card("📋 Linhas", fmt_int(linhas), "Registros da metrica principal", "#f0ad4e")
    with c4:
        metric_card("📍 Localidade", localidade, periodo, "#1c2d42")

st.caption("Motor de filtro: DuckDB sobre Parquet local em cache. Fonte: MJSP/SINESP - Dados Nacionais de Seguranca Publica.")
st.markdown("---")

# Cards especificos de vitimas por sexo para UF
if tipo == "uf":
    st.subheader("Analise das vitimas por sexo e tipo de crime")
    if df_vitimas.empty or (sexo_tab.empty and vitimas_crime_tab.empty):
        st.warning("A estrutura de Vitimas nao retornou dados para os filtros de UF/ano/mes/crime selecionados.")
    else:
        total_v = total_vitimas_complementar
        masc = float(sexo_dict.get("Masculino", 0))
        fem = float(sexo_dict.get("Feminino", 0))
        ni = float(sexo_dict.get("Sexo NI", 0))
        s1, s2, s3, s4, s5 = st.columns(5)
        with s1:
            metric_card("👥 Total de vitimas", fmt_int(total_v), "Total na estrutura Vitimas", "#0b2239")
        with s2:
            metric_card("♂ Masculino", fmt_int(masc), f"{fmt_pct(masc / total_v * 100 if total_v else 0)} do total", "#1c2d42")
        with s3:
            metric_card("♀ Feminino", fmt_int(fem), f"{fmt_pct(fem / total_v * 100 if total_v else 0)} do total", "#d9534f")
        with s4:
            metric_card("NI", fmt_int(ni), "Sexo nao informado", "#f0ad4e")
        with s5:
            metric_card("🔎 Crime com mais vitimas", crime_top, f"{fmt_int(crime_top_v)} vitimas", "#1c2d42")

aba_painel, aba_dados, aba_diag, aba_export = st.tabs(["📊 Painel Estatistico", "📋 Dados Tratados", "⚙️ Diagnostico", "📥 Exportacao"])

with aba_painel:
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Evolucao mensal ({metrica_label})**")
        s = serie_mensal(df_main, metrica_coluna)
        if s.empty:
            st.info("Sem serie mensal para os filtros selecionados.")
        else:
            st.line_chart(s)
    with col2:
        if tipo == "municipios":
            if municipio_sel == "Todos os municipios":
                st.write("**Ranking municipal - homicidio doloso (Vitimas)**")
                rank = serie_por_coluna(df_main, "MUNICIPIO", "VITIMAS", top=15)
                if rank.empty:
                    st.info("Sem ranking para os filtros selecionados.")
                else:
                    st.bar_chart(rank)
            else:
                st.info("No municipio especifico, nao ha distribuicao por tipo de crime: a base municipal e fixa em homicidio doloso/vitimas.")
        else:
            st.write(f"**Distribuicao por indicador/tipo de crime ({metrica_label})**")
            crime_series = serie_por_coluna(df_main, "TIPO_CRIME", metrica_coluna, top=15)
            if crime_series.empty:
                st.info("Sem distribuicao por crime para os filtros selecionados.")
            else:
                st.bar_chart(crime_series)

    if tipo == "uf":
        st.markdown("---")
        csexo1, csexo2 = st.columns(2)
        with csexo1:
            st.write("**Vitimas por sexo**")
            if sexo_tab.empty:
                st.info("Sem recorte por sexo para estes filtros.")
            else:
                st.bar_chart(sexo_tab.set_index("SEXO_DA_VITIMA")["VITIMAS"])
                st.dataframe(sexo_tab, width="stretch", hide_index=True)
        with csexo2:
            st.write("**Vitimas por tipo de crime**")
            if vitimas_crime_tab.empty:
                st.info("Sem vitimas por tipo de crime para estes filtros.")
            else:
                st.bar_chart(vitimas_crime_tab.set_index("TIPO_CRIME")["VITIMAS"])
                st.dataframe(vitimas_crime_tab, width="stretch", hide_index=True)

        with st.expander("Tabela detalhada: tipo de crime x sexo da vitima", expanded=True):
            if crime_sexo_tab.empty:
                st.info("Sem tabela cruzada para estes filtros.")
            else:
                st.dataframe(crime_sexo_tab, width="stretch", hide_index=True)

    if tipo == "municipios":
        st.markdown("---")
        st.info("A base municipal oficial nao inclui sexo da vitima nem os demais tipos de crime. Para sexo e tipos de crime, use o nivel UF.")

with aba_dados:
    cols = [
        "ANO", "MES_NOME", "UF", "MUNICIPIO", "TIPO_CRIME", "UNIDADE_MEDIDA", "SEXO_DA_VITIMA",
        "OCORRENCIAS", "VITIMAS", "CODIGO_MUNICIPIO", "METRICA_FONTE", "ABA_ORIGEM", "FONTE_PROCESSADA",
    ]
    cols = [c for c in cols if c in df_main.columns]
    st.write("**Dados da metrica principal**")
    st.dataframe(df_main[cols], width="stretch", hide_index=True)

    if tipo == "uf" and not df_vitimas.empty:
        st.write("**Dados complementares - estrutura de Vitimas**")
        cols_v = [c for c in cols if c in df_vitimas.columns]
        st.dataframe(df_vitimas[cols_v], width="stretch", hide_index=True)

with aba_diag:
    st.write("**Diagnostico da consulta**")
    st.json({
        "nivel": tipo,
        "uf": uf_sel,
        "municipio": municipio_sel if tipo == "municipios" else None,
        "crime": crime_sel,
        "ano": ano_sel,
        "mes": mes_nome,
        "metrica_principal": metrica_label,
        "metrica_fonte": metrica_fonte if tipo != "municipios" else "VITIMAS",
        "linhas_main": int(len(df_main)),
        "linhas_vitimas_complementar": int(len(df_vitimas)) if tipo == "uf" else None,
        "total_principal": float(total_principal),
        "total_vitimas_complementar": float(total_vitimas_complementar),
        "metricas_disponiveis_uf": distinct_values(path, "METRICA_FONTE", tipo, uf=uf_sel),
        "abas_disponiveis": distinct_values(path, "ABA_ORIGEM", tipo, uf=uf_sel),
        "sexo_disponivel": distinct_values(path, "SEXO_DA_VITIMA", tipo, uf=uf_sel, metrica="VITIMAS") if tipo == "uf" else [],
        "crimes_vitimas_disponiveis": distinct_values(path, "TIPO_CRIME", tipo, uf=uf_sel, metrica="VITIMAS") if tipo == "uf" else [],
        "observacao_municipal": "Municipio e fixo em Homicidio doloso/Vitimas; nao ha filtro de tipo de crime, ocorrencias ou sexo." if tipo == "municipios" else None,
        "cache": meta,
    })
    st.write("**Amostra filtrada da metrica principal**")
    st.dataframe(df_main.head(80), width="stretch", hide_index=True)
    if tipo == "uf" and not df_vitimas.empty:
        st.write("**Amostra filtrada de vitimas**")
        st.dataframe(df_vitimas.head(80), width="stretch", hide_index=True)

with aba_export:
    st.download_button(
        "📥 Baixar dados filtrados CSV",
        data=df_main.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
        file_name=f"sinesp_{tipo}_{uf_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
        mime="text/csv",
    )
    if tipo == "uf" and not sexo_tab.empty:
        st.download_button(
            "📥 Baixar resumo por sexo CSV",
            data=sexo_tab.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
            file_name=f"sinesp_resumo_sexo_{uf_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )
    if tipo == "uf" and not vitimas_crime_tab.empty:
        st.download_button(
            "📥 Baixar vitimas por tipo de crime CSV",
            data=vitimas_crime_tab.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
            file_name=f"sinesp_vitimas_crime_{uf_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )
    if tipo == "uf" and not crime_sexo_tab.empty:
        st.download_button(
            "📥 Baixar tipo de crime x sexo CSV",
            data=crime_sexo_tab.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
            file_name=f"sinesp_crime_sexo_{uf_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )
