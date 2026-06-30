# VERSION_FINAL_SEGURANCA_SINESP_OFICIAL_DUCKDB_V14_SEXO_TIPO_CRIME
# Fonte: MJSP/SINESP - Dados Nacionais de Seguranca Publica
# Estrutura metodologica:
# - Municipio: base oficial municipal, indicador Homicidio doloso, unidade Vitimas.
# - UF: base oficial por UF, aba Ocorrencias com varios crimes e aba Vitimas com perfil por sexo para Homicidio doloso.

import os
import io
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
# CONFIGURACAO GERAL
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
        min-height: 145px;
    }
    .metric-card h4 { margin-top: 0; color: #0b2239; }
    .metric-card h2 { margin: 0; }
    .small-note { color: #666; font-size: 0.88rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

CACHE_SCHEMA_VERSION = "v14_sexo_tipo_crime_20260630"

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
NOME_PARA_UF = {v: k for k, v in UF_NOMES.items()}

MES_ORDEM = {
    "JANEIRO": 1, "FEVEREIRO": 2, "MARCO": 3, "MARÇO": 3, "ABRIL": 4, "MAIO": 5, "JUNHO": 6,
    "JULHO": 7, "AGOSTO": 8, "SETEMBRO": 9, "OUTUBRO": 10, "NOVEMBRO": 11, "DEZEMBRO": 12,
    "JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6, "JUL": 7, "AGO": 8,
    "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12,
}
MES_LABEL = {
    1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez",
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
    texto = normalizar_texto(texto)
    texto = re.sub(r"[^A-Z0-9]+", "", texto)
    return texto


def normalizar_coluna(col):
    col = normalizar_texto(col)
    col = col.replace(".", "")
    col = re.sub(r"[^A-Z0-9]+", "_", col)
    return col.strip("_")


def converter_numero(valor):
    if pd.isna(valor):
        return 0.0
    s = str(valor).strip()
    if s == "":
        return 0.0
    s = s.replace(".", "").replace(",", ".") if re.search(r"\d+,\d+$", s) else s.replace(",", "")
    return pd.to_numeric(s, errors="coerce") if s else 0.0


def parse_mes_ano(valor):
    if pd.isna(valor):
        return pd.NaT
    s = str(valor).strip()
    if not s:
        return pd.NaT
    # XLSX oficial municipal vem como YYYY-MM-DD HH:MM:SS.
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


def padronizar_nome_municipio(valor):
    if pd.isna(valor):
        return ""
    txt = str(valor).strip()
    if not txt:
        return ""
    # Title preserva legibilidade. O filtro usa chave sem acento, entao nao depende da grafia exata.
    return txt.title()


def uf_para_sigla(valor):
    if pd.isna(valor):
        return ""
    val = normalizar_texto(valor)
    if val in UFS:
        return val
    return NOME_PARA_UF.get(val.title(), NOME_PARA_UF.get(val, ""))


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

# -----------------------------------------------------------------------------
# LEITURA / PREPARO DAS BASES OFICIAIS
# -----------------------------------------------------------------------------
def pasta_cache():
    p = os.path.join(tempfile.gettempdir(), "sinesp_mjsp_duckdb_cache")
    os.makedirs(p, exist_ok=True)
    return p


def caminho_parquet(tipo):
    return os.path.join(pasta_cache(), f"sinesp_{tipo}_{CACHE_SCHEMA_VERSION}.parquet")


def baixar_xlsx(url, nome_local):
    # Em desenvolvimento local, usa o arquivo se estiver ao lado do app.
    if os.path.exists(nome_local):
        with open(nome_local, "rb") as f:
            return f.read(), f"arquivo local: {nome_local}"
    resp = requests.get(url, timeout=180, verify=False, headers={"User-Agent": "Mozilla/5.0"})
    resp.raise_for_status()
    return resp.content, url


def preparar_municipios(conteudo):
    xls = pd.ExcelFile(io.BytesIO(conteudo), engine="openpyxl")
    frames = []
    for aba in xls.sheet_names:
        # Cada aba e uma UF.
        if normalizar_texto(aba) not in UFS:
            continue
        df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, dtype=str, engine="openpyxl")
        df = df.dropna(how="all")
        df.columns = [normalizar_coluna(c) for c in df.columns]
        ren = {
            "COD_IBGE": "CODIGO_MUNICIPIO",
            "CODIGO_IBGE": "CODIGO_MUNICIPIO",
            "MUNICIPIO": "MUNICIPIO",
            "SIGLA_UF": "UF",
            "UF": "UF",
            "REGIAO": "REGIAO",
            "MES_ANO": "MES_ANO",
            "VITIMAS": "VITIMAS",
        }
        df = df.rename(columns={c: ren[c] for c in df.columns if c in ren})
        obrigatorias = ["MUNICIPIO", "UF", "MES_ANO", "VITIMAS"]
        if not set(obrigatorias).issubset(df.columns):
            continue
        if "CODIGO_MUNICIPIO" not in df.columns:
            df["CODIGO_MUNICIPIO"] = ""
        if "REGIAO" not in df.columns:
            df["REGIAO"] = ""
        df["ABA_ORIGEM"] = aba
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    df = pd.concat(frames, ignore_index=True, sort=False)
    df["UF"] = df["UF"].map(uf_para_sigla)
    df["MUNICIPIO"] = df["MUNICIPIO"].map(padronizar_nome_municipio)
    dt = df["MES_ANO"].map(parse_mes_ano)
    df["ANO"] = dt.dt.year.astype("Int64").astype(str).replace("<NA>", "Nao informado")
    df["MES_ORDEM"] = dt.dt.month.astype("Int64")
    df["MES_NOME"] = df["MES_ORDEM"].map(lambda x: MES_LABEL.get(int(x), "Nao informado") if pd.notna(x) else "Nao informado")
    df["TIPO_CRIME"] = "Homicidio doloso"
    df["UNIDADE_MEDIDA"] = "Vitimas"
    df["OCORRENCIAS"] = 0.0
    df["VITIMAS"] = df["VITIMAS"].map(converter_numero).fillna(0).astype(float)
    df["SEXO_DA_VITIMA"] = "Nao informado na base municipal"
    df["NIVEL_BASE"] = "municipio"
    df["FONTE_PROCESSADA"] = "MJSP/SINESP - Municipios XLSX"

    return adicionar_filtros(df)


def preparar_uf(conteudo):
    xls = pd.ExcelFile(io.BytesIO(conteudo), engine="openpyxl")
    frames = []
    for aba in xls.sheet_names:
        aba_norm = normalizar_texto(aba)
        if "OCOR" in aba_norm:
            unidade = "Ocorrencias"
            valor_col = "OCORRENCIAS"
        elif "VIT" in aba_norm:
            unidade = "Vitimas"
            valor_col = "VITIMAS"
        else:
            continue

        df = pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, dtype=str, engine="openpyxl")
        df = df.dropna(how="all")
        df.columns = [normalizar_coluna(c) for c in df.columns]
        ren = {
            "UF": "UF",
            "TIPO_CRIME": "TIPO_CRIME",
            "TIPO_DE_CRIME": "TIPO_CRIME",
            "CRIME": "TIPO_CRIME",
            "ANO": "ANO",
            "MES": "MES",
            "OCORRENCIAS": "VALOR",
            "VITIMAS": "VALOR",
            "SEXO_DA_VITIMA": "SEXO_DA_VITIMA",
            "SEXO_VITIMA": "SEXO_DA_VITIMA",
            "SEXO": "SEXO_DA_VITIMA",
        }
        df = df.rename(columns={c: ren[c] for c in df.columns if c in ren})
        obrigatorias = ["UF", "TIPO_CRIME", "ANO", "MES", "VALOR"]
        if not set(obrigatorias).issubset(df.columns):
            continue

        df["UF"] = df["UF"].map(uf_para_sigla)
        df["MUNICIPIO"] = ""
        df["CODIGO_MUNICIPIO"] = ""
        df["REGIAO"] = ""
        df["ANO"] = df["ANO"].astype(str).str.extract(r"(\d{4})", expand=False).fillna("Nao informado")
        df["MES_ORDEM"] = df["MES"].map(lambda x: MES_ORDEM.get(normalizar_texto(x), pd.NA))
        df["MES_NOME"] = df["MES_ORDEM"].map(lambda x: MES_LABEL.get(int(x), "Nao informado") if pd.notna(x) else "Nao informado")
        df["TIPO_CRIME"] = df["TIPO_CRIME"].astype(str).str.strip().replace({"": "Nao informado"})
        df["UNIDADE_MEDIDA"] = unidade
        df["OCORRENCIAS"] = 0.0
        df["VITIMAS"] = 0.0
        valores = df["VALOR"].map(converter_numero).fillna(0).astype(float)
        df[valor_col] = valores
        if "SEXO_DA_VITIMA" in df.columns:
            df["SEXO_DA_VITIMA"] = df["SEXO_DA_VITIMA"].map(padronizar_sexo)
        else:
            df["SEXO_DA_VITIMA"] = ""
        df["ABA_ORIGEM"] = unidade
        df["NIVEL_BASE"] = "uf"
        df["FONTE_PROCESSADA"] = "MJSP/SINESP - UF XLSX"
        frames.append(df)

    if not frames:
        return pd.DataFrame()
    df = pd.concat(frames, ignore_index=True, sort=False)
    return adicionar_filtros(df)


def adicionar_filtros(df):
    df = df.copy()
    df["UF_FILTRO"] = df["UF"].astype(str).str.upper().str.strip()
    df["MUNICIPIO_FILTRO"] = df["MUNICIPIO"].map(chave_filtro)
    df["TIPO_CRIME_FILTRO"] = df["TIPO_CRIME"].map(chave_filtro)
    df["UNIDADE_MEDIDA_FILTRO"] = df["UNIDADE_MEDIDA"].map(chave_filtro)
    df["SEXO_DA_VITIMA_FILTRO"] = df["SEXO_DA_VITIMA"].map(chave_filtro)
    df["ANO_FILTRO"] = df["ANO"].astype(str).str.strip()
    df["MES_ORDEM_NUM"] = pd.to_numeric(df["MES_ORDEM"], errors="coerce")
    colunas = [
        "CODIGO_MUNICIPIO", "MUNICIPIO", "UF", "REGIAO", "ANO", "MES_ORDEM", "MES_NOME", "MES_ANO",
        "TIPO_CRIME", "UNIDADE_MEDIDA", "SEXO_DA_VITIMA", "OCORRENCIAS", "VITIMAS", "NIVEL_BASE", "ABA_ORIGEM",
        "FONTE_PROCESSADA", "UF_FILTRO", "MUNICIPIO_FILTRO", "TIPO_CRIME_FILTRO", "UNIDADE_MEDIDA_FILTRO",
        "SEXO_DA_VITIMA_FILTRO", "ANO_FILTRO", "MES_ORDEM_NUM",
    ]
    for c in colunas:
        if c not in df.columns:
            df[c] = ""
    return df[colunas]


@st.cache_data(ttl=86400, show_spinner=False)
def preparar_parquet(tipo):
    path = caminho_parquet(tipo)
    try:
        if os.path.exists(path):
            return {"ok": True, "parquet_path": path, "cache": True, "tipo": tipo}
        if tipo == "municipios":
            conteudo, origem = baixar_xlsx(URL_MJSP_MUNICIPIOS, "indicadoressegurancapublicamunic.xlsx")
            df = preparar_municipios(conteudo)
        else:
            conteudo, origem = baixar_xlsx(URL_MJSP_UF, "indicadoressegurancapublicauf.xlsx")
            df = preparar_uf(conteudo)
        if df.empty:
            return {"ok": False, "erro": "A base foi lida, mas nenhum registro valido foi identificado.", "tipo": tipo}
        df.to_parquet(path, index=False)
        return {
            "ok": True,
            "parquet_path": path,
            "cache": False,
            "tipo": tipo,
            "origem": origem,
            "linhas": int(len(df)),
            "anos": sorted([a for a in df["ANO_FILTRO"].dropna().unique().tolist() if re.fullmatch(r"\d{4}", str(a))]),
            "colunas": df.columns.tolist(),
        }
    except Exception as e:
        return {"ok": False, "erro": str(e), "tipo": tipo}

# -----------------------------------------------------------------------------
# DUCKDB HELPERS
# -----------------------------------------------------------------------------
def sql_val(v):
    return str(v).replace("'", "''")


def build_where(tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, unidade=None, sexo=None):
    clauses = []
    params = []
    if uf:
        clauses.append("UF_FILTRO = ?")
        params.append(uf)
    if tipo == "municipios" and municipio and municipio != "Todos os municipios":
        clauses.append("MUNICIPIO_FILTRO = ?")
        params.append(chave_filtro(municipio))
    # Na base MUNICIPAL oficial, o indicador e fixo por metodologia: Homicidio doloso / Vitimas.
    # Portanto NAO aplicamos filtro por TIPO_CRIME aqui.
    # Isso evita o bug em que "Homicidio doloso" nao casava exatamente com o valor padronizado
    # e a consulta so retornava quando o usuario escolhia "Todos os indicadores".
    if tipo != "municipios" and crime and crime not in ["Todos os indicadores", "Todos os tipos de crime"]:
        clauses.append("TIPO_CRIME_FILTRO = ?")
        params.append(chave_filtro(crime))
    if ano and ano != "Todos os anos":
        clauses.append("ANO_FILTRO = ?")
        params.append(str(ano))
    if mes is not None:
        clauses.append("MES_ORDEM_NUM = ?")
        params.append(int(mes))
    # A base municipal ja e homogenea em Vitimas. O filtro por unidade tambem fica
    # desabilitado para municipio para evitar zeragem por divergencia de grafia.
    if tipo != "municipios" and unidade:
        clauses.append("UNIDADE_MEDIDA_FILTRO = ?")
        params.append(chave_filtro(unidade))
    if sexo and sexo != "Todos os sexos":
        clauses.append("SEXO_DA_VITIMA_FILTRO = ?")
        params.append(chave_filtro(sexo))
    where = " WHERE " + " AND ".join(clauses) if clauses else ""
    return where, params


def query_df(path, tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, unidade=None, sexo=None, limit=None):
    where, params = build_where(tipo, uf, municipio, crime, ano, mes, unidade, sexo)
    lim = f" LIMIT {int(limit)}" if limit else ""
    sql = f"SELECT * FROM read_parquet(?) {where}{lim}"
    return duckdb.execute(sql, [path] + params).df()


def distinct_values(path, coluna, tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, unidade=None):
    where, params = build_where(tipo, uf, municipio, crime, ano, mes, unidade)
    sql = f"SELECT DISTINCT {coluna} AS valor FROM read_parquet(?) {where} ORDER BY valor"
    try:
        df = duckdb.execute(sql, [path] + params).df()
        return [str(v) for v in df["valor"].dropna().tolist() if str(v).strip() and str(v) != "<NA>"]
    except Exception:
        return []


def count_rows(path, tipo, uf=None, municipio=None, crime=None, ano=None, mes=None, unidade=None, sexo=None):
    where, params = build_where(tipo, uf, municipio, crime, ano, mes, unidade, sexo)
    sql = f"SELECT COUNT(*) AS n FROM read_parquet(?) {where}"
    return int(duckdb.execute(sql, [path] + params).df()["n"].iloc[0])

# -----------------------------------------------------------------------------
# VISUALIZACAO
# -----------------------------------------------------------------------------
def fmt_int(v):
    try:
        return f"{int(round(float(v))):,}".replace(",", ".")
    except Exception:
        return "0"


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


def tabela_sexo(df_vitimas):
    if df_vitimas.empty or "SEXO_DA_VITIMA" not in df_vitimas.columns:
        return pd.DataFrame()
    base = df_vitimas.copy()
    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    base = base[base["SEXO_DA_VITIMA"].astype(str).str.strip().ne("")]
    if base.empty:
        return pd.DataFrame()
    tab = base.groupby("SEXO_DA_VITIMA", as_index=False)["VITIMAS"].sum()
    tab = tab[tab["VITIMAS"] > 0].sort_values("VITIMAS", ascending=False)
    total = tab["VITIMAS"].sum()
    tab["PARTICIPACAO_%"] = (tab["VITIMAS"] / total * 100).round(2) if total else 0
    return tab


def tabela_vitimas_por_crime(df_vitimas, top=15):
    if df_vitimas.empty or "TIPO_CRIME" not in df_vitimas.columns:
        return pd.DataFrame()
    base = df_vitimas.copy()
    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    tab = base.groupby("TIPO_CRIME", as_index=False)["VITIMAS"].sum()
    tab = tab[tab["VITIMAS"] > 0].sort_values("VITIMAS", ascending=False).head(top)
    total = tab["VITIMAS"].sum()
    tab["PARTICIPACAO_%"] = (tab["VITIMAS"] / total * 100).round(2) if total else 0
    return tab


def tabela_vitimas_crime_sexo(df_vitimas, top_crimes=12):
    if df_vitimas.empty or not {"TIPO_CRIME", "SEXO_DA_VITIMA", "VITIMAS"}.issubset(df_vitimas.columns):
        return pd.DataFrame()
    base = df_vitimas.copy()
    base["VITIMAS"] = pd.to_numeric(base["VITIMAS"], errors="coerce").fillna(0)
    base = base[base["VITIMAS"] > 0]
    if base.empty:
        return pd.DataFrame()
    top = (
        base.groupby("TIPO_CRIME")["VITIMAS"]
        .sum()
        .sort_values(ascending=False)
        .head(top_crimes)
        .index
        .tolist()
    )
    base = base[base["TIPO_CRIME"].isin(top)]
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
if tipo == "municipios":
    municipios = distinct_values(path, "MUNICIPIO", tipo, uf=uf_sel)
    municipios = sorted(set([m for m in municipios if m.strip()]))
    municipio_sel = st.sidebar.selectbox("Municipio:", ["Todos os municipios"] + municipios)
    st.sidebar.caption("Base municipal oficial: Homicidio doloso em vitimas. Nao ha sexo da vitima nesta base municipal.")
else:
    st.sidebar.info("A base por UF e agregada por estado. Ela traz ocorrencias por tipo de crime e, na aba Vitimas, permite analise por sexo da vitima e tipo de crime.")

if tipo == "municipios":
    metrica_label = "Vitimas"
    metrica_coluna = "VITIMAS"
    unidade = "Vitimas"
    crime_sel = "Homicidio doloso"
    st.sidebar.markdown("**Indicador:** Homicidio doloso")
    st.sidebar.caption("No recorte municipal oficial, o indicador e fixo. O app nao aplica filtro textual de crime nesta base; filtra apenas UF, municipio, ano e mes.")
    sexo_sel = "Todos os sexos"
else:
    metrica_label = st.sidebar.selectbox("Metrica principal:", ["Ocorrencias", "Vitimas"], index=0)
    metrica_coluna = "OCORRENCIAS" if metrica_label == "Ocorrencias" else "VITIMAS"
    unidade = metrica_label
    crimes = distinct_values(path, "TIPO_CRIME", tipo, uf=uf_sel, unidade=unidade)
    crime_sel = st.sidebar.selectbox("Indicador/tipo de crime:", ["Todos os indicadores"] + crimes)
    sexo_sel = "Todos os sexos"
    if metrica_label == "Vitimas":
        sexos = distinct_values(path, "SEXO_DA_VITIMA", tipo, uf=uf_sel, crime=crime_sel, unidade="Vitimas")
        sexos = [s for s in sexos if s.strip()]
        if sexos:
            sexo_sel = st.sidebar.selectbox("Sexo da vitima:", ["Todos os sexos"] + sexos)

anos = distinct_values(path, "ANO_FILTRO", tipo, uf=uf_sel, municipio=municipio_sel if tipo == "municipios" else None, crime=crime_sel, unidade=unidade)
anos = sorted([a for a in anos if re.fullmatch(r"\d{4}", str(a))], reverse=True)
ano_sel = st.sidebar.selectbox("Ano:", ["Todos os anos"] + anos, index=1 if anos else 0)

mes_nome = st.sidebar.selectbox("Mes:", list(MESES_DISPLAY.keys()), index=0)
mes_sel = MESES_DISPLAY[mes_nome]

with st.sidebar.form("form_seguranca"):
    consultar = st.form_submit_button("🔍 Consultar Indicadores")

# -----------------------------------------------------------------------------
# INICIO
# -----------------------------------------------------------------------------
if not consultar:
    st.info("Escolha os filtros na barra lateral e clique em Consultar Indicadores.")
    st.markdown(
        """
        **Leitura correta da fonte:**
        - **Municipio:** base oficial municipal, com **homicidio doloso** medido em **vitimas**.
        - **UF:** base oficial estadual, com varios crimes em **ocorrencias** e, na aba de vitimas, analise por **sexo da vitima** e **tipo de crime**.
        - **Sexo NI:** sexo nao informado/nao identificado; deve permanecer como categoria propria.
        """
    )
    st.stop()

# -----------------------------------------------------------------------------
# CONSULTA PRINCIPAL
# -----------------------------------------------------------------------------
st.markdown(
    '<div class="header-seguranca"><h1>Painel de Ocorrencias Criminais - SINESP/MJSP</h1>'
    '<p>Indicadores oficiais de seguranca publica com filtro territorial</p></div>',
    unsafe_allow_html=True,
)

if tipo == "municipios":
    # Base municipal: indicador e unidade sao fixos por metodologia.
    # Consulta sem filtro de crime/unidade para evitar divergencia de grafia/codificacao.
    df_main = query_df(path, tipo, uf=uf_sel, municipio=municipio_sel, ano=ano_sel, mes=mes_sel)
    df_vitimas = df_main.copy()
else:
    df_main = query_df(path, tipo, uf=uf_sel, crime=crime_sel, ano=ano_sel, mes=mes_sel, unidade=unidade, sexo=sexo_sel)
    # Perfil de vitimas calculado separadamente para nao misturar aba Ocorrencias e aba Vitimas.
    df_vitimas = query_df(path, tipo, uf=uf_sel, crime=crime_sel, ano=ano_sel, mes=mes_sel, unidade="Vitimas")

if df_main.empty:
    st.warning("Nao foram encontrados registros para a combinacao selecionada.")
    st.caption("Use a aba de diagnostico abaixo para verificar disponibilidade de UF, municipio, ano e unidade de medida.")
    diag = {
        "nivel": tipo,
        "uf": uf_sel,
        "municipio": municipio_sel if tipo == "municipios" else None,
        "crime": crime_sel,
        "observacao_municipal": "No nivel municipal, Homicidio doloso/Vitimas e indicador fixo; o filtro por crime nao e aplicado." if tipo == "municipios" else None,
        "ano": ano_sel,
        "mes": mes_nome,
        "unidade": unidade,
        "linhas_uf": count_rows(path, tipo, uf=uf_sel),
        "linhas_uf_municipio": count_rows(path, tipo, uf=uf_sel, municipio=municipio_sel) if tipo == "municipios" else None,
        "anos_disponiveis_no_contexto": anos,
        "cache": meta,
    }
    st.json(diag)
    st.stop()

# Totais
total_principal = pd.to_numeric(df_main[metrica_coluna], errors="coerce").fillna(0).sum()
total_ocorrencias = pd.to_numeric(df_main.get("OCORRENCIAS", pd.Series([0]*len(df_main))), errors="coerce").fillna(0).sum()
total_vitimas_contexto = pd.to_numeric(df_vitimas.get("VITIMAS", pd.Series([0]*len(df_vitimas))), errors="coerce").fillna(0).sum()
linhas = len(df_main)
localidade = f"{municipio_sel} - {uf_sel}" if tipo == "municipios" and municipio_sel != "Todos os municipios" else (f"Municipios de {uf_sel}" if tipo == "municipios" else f"Estado {uf_sel}")
periodo = f"{ano_sel} | {mes_nome}"

c1, c2, c3, c4 = st.columns(4)
with c1:
    metric_card(f"🚨 {metrica_label}", fmt_int(total_principal), "Metrica principal do filtro", "#d9534f")
with c2:
    if tipo == "uf" and metrica_label == "Ocorrencias":
        metric_card("👥 Vitimas", fmt_int(total_vitimas_contexto), "Perfil calculado pela aba Vitimas", "#0b2239")
    elif tipo == "uf" and metrica_label == "Vitimas":
        metric_card("👥 Vitimas", fmt_int(total_vitimas_contexto), "Total de vitimas do filtro", "#0b2239")
    else:
        metric_card("📌 Indicador", "Homicidio doloso", "Base municipal oficial", "#0b2239")
with c3:
    metric_card("📋 Linhas", fmt_int(linhas), "Registros retornados", "#f0ad4e")
with c4:
    metric_card("📍 Localidade", localidade, periodo, "#1c2d42")

st.caption("Motor de filtro: DuckDB sobre Parquet local em cache. A primeira carga do XLSX pode demorar, mas as consultas seguintes ficam mais leves.")
st.caption("Fonte: MJSP/SINESP - Dados Nacionais de Seguranca Publica.")
st.markdown("---")

# Analises complementares de vitimas para UF.
sexo_tab = tabela_sexo(df_vitimas) if tipo == "uf" else pd.DataFrame()
vitimas_crime_tab = tabela_vitimas_por_crime(df_vitimas, top=15) if tipo == "uf" else pd.DataFrame()
crime_sexo_tab = tabela_vitimas_crime_sexo(df_vitimas, top_crimes=12) if tipo == "uf" else pd.DataFrame()

if tipo == "uf":
    st.subheader("Analise complementar das vitimas")
    if sexo_tab.empty and vitimas_crime_tab.empty:
        st.warning(
            "A aba de Vitimas nao retornou recorte complementar para estes filtros. "
            "Verifique se o ano, mes e indicador possuem vitimas informadas na fonte."
        )
    else:
        sdict = dict(zip(sexo_tab.get("SEXO_DA_VITIMA", []), sexo_tab.get("VITIMAS", []))) if not sexo_tab.empty else {}
        crime_top = vitimas_crime_tab.iloc[0]["TIPO_CRIME"] if not vitimas_crime_tab.empty else "Nao informado"
        crime_top_v = vitimas_crime_tab.iloc[0]["VITIMAS"] if not vitimas_crime_tab.empty else 0
        s1, s2, s3, s4, s5 = st.columns(5)
        with s1:
            metric_card("👥 Total vitimas", fmt_int(total_vitimas_contexto), "Aba Vitimas", "#0b2239")
        with s2:
            metric_card("♂ Masculino", fmt_int(sdict.get("Masculino", 0)), "Vitimas masculinas", "#1c2d42")
        with s3:
            metric_card("♀ Feminino", fmt_int(sdict.get("Feminino", 0)), "Vitimas femininas", "#d9534f")
        with s4:
            metric_card("NI", fmt_int(sdict.get("Sexo NI", 0)), "Sexo nao informado", "#f0ad4e")
        with s5:
            metric_card("🔎 Crime com mais vitimas", crime_top, f"{fmt_int(crime_top_v)} vitimas", "#1c2d42")

# Abas
aba_painel, aba_dados, aba_diag, aba_export = st.tabs(["📊 Painel Estatistico", "📋 Dados Tratados", "⚙️ Diagnostico", "📥 Exportacao"])

with aba_painel:
    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Evolucao mensal ({metrica_label})**")
        s_mes = serie_mensal(df_main, metrica_coluna)
        if s_mes.empty:
            st.info("Sem serie mensal para os filtros selecionados.")
        else:
            st.line_chart(s_mes)
    with col2:
        if tipo == "uf":
            st.write(f"**Distribuicao por indicador/tipo de crime ({metrica_label})**")
            s_crime = serie_por_coluna(df_main, "TIPO_CRIME", metrica_coluna, top=15)
            if s_crime.empty:
                st.info("Sem distribuicao por crime para os filtros selecionados.")
            else:
                st.bar_chart(s_crime)
        else:
            if municipio_sel == "Todos os municipios":
                st.write("**Ranking municipal - homicidio doloso (Vitimas)**")
                s_rank = serie_por_coluna(df_main, "MUNICIPIO", "VITIMAS", top=15)
                if s_rank.empty:
                    st.info("Sem ranking municipal para os filtros selecionados.")
                else:
                    st.bar_chart(s_rank)
            else:
                st.info("Ranking municipal fica oculto quando um municipio especifico esta selecionado.")

    if tipo == "uf":
        st.markdown("---")
        st.write("### Analise das vitimas")
        csexo1, csexo2 = st.columns(2)
        with csexo1:
            st.write("**Vitimas por sexo**")
            if sexo_tab.empty:
                st.info("Sem recorte por sexo para os filtros selecionados.")
            else:
                st.bar_chart(sexo_tab.set_index("SEXO_DA_VITIMA")["VITIMAS"])
        with csexo2:
            st.write("**Vitimas por tipo de crime**")
            if vitimas_crime_tab.empty:
                st.info("Sem vitimas por tipo de crime para os filtros selecionados.")
            else:
                st.bar_chart(vitimas_crime_tab.set_index("TIPO_CRIME")["VITIMAS"])

        with st.expander("Tabela detalhada: vitimas por sexo e tipo de crime", expanded=True):
            if crime_sexo_tab.empty:
                st.info("A fonte nao retornou tabela cruzada de sexo por tipo de crime para este filtro.")
            else:
                st.dataframe(crime_sexo_tab, width="stretch", hide_index=True)

        with st.expander("Resumo de vitimas por sexo", expanded=False):
            if sexo_tab.empty:
                st.info("Sem resumo por sexo para os filtros selecionados.")
            else:
                st.dataframe(sexo_tab, width="stretch", hide_index=True)
    elif tipo == "municipios":
        st.info("A base municipal oficial nao traz sexo da vitima. O perfil por sexo esta disponivel apenas na base agregada por UF.")

with aba_dados:
    cols = [
        "ANO", "MES_NOME", "UF", "MUNICIPIO", "TIPO_CRIME", "UNIDADE_MEDIDA", "SEXO_DA_VITIMA",
        "OCORRENCIAS", "VITIMAS", "CODIGO_MUNICIPIO", "FONTE_PROCESSADA",
    ]
    cols = [c for c in cols if c in df_main.columns]
    st.write("**Dados da metrica principal selecionada**")
    st.dataframe(df_main[cols], width="stretch", hide_index=True)
    if tipo == "uf" and not df_vitimas.empty:
        st.write("**Dados complementares da aba Vitimas usados para sexo e tipo de crime**")
        cols_v = [c for c in cols if c in df_vitimas.columns]
        st.dataframe(df_vitimas[cols_v], width="stretch", hide_index=True)

with aba_diag:
    st.write("**Resumo da consulta**")
    st.json({
        "nivel": tipo,
        "uf": uf_sel,
        "municipio": municipio_sel if tipo == "municipios" else None,
        "crime": crime_sel,
        "observacao_municipal": "No nivel municipal, Homicidio doloso/Vitimas e indicador fixo; o filtro por crime nao e aplicado." if tipo == "municipios" else None,
        "ano": ano_sel,
        "mes": mes_nome,
        "metrica": metrica_label,
        "linhas_retornadas": linhas,
        "total_principal": float(total_principal),
        "total_vitimas_contexto": float(total_vitimas_contexto),
        "linhas_aba_vitimas": int(len(df_vitimas)) if tipo == "uf" else None,
        "sexo_vitima_disponivel": (not sexo_tab.empty) if tipo == "uf" else None,
        "vitimas_por_tipo_crime_disponivel": (not vitimas_crime_tab.empty) if tipo == "uf" else None,
        "cache": meta,
    })
    st.write("**Amostra do parquet filtrado**")
    st.dataframe(df_main.head(50), width="stretch", hide_index=True)

with aba_export:
    csv = df_main.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
    st.download_button(
        "📥 Baixar dados filtrados CSV",
        data=csv,
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
            "📥 Baixar tabela crime x sexo CSV",
            data=crime_sexo_tab.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
            file_name=f"sinesp_crime_sexo_{uf_sel}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
            mime="text/csv",
        )
