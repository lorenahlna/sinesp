# VERSION_SEGURANCA_SINESP_OFICIAL_V17_FINAL
# App Streamlit para Dados Nacionais de Seguranca Publica - MJSP/SINESP
# Regras metodologicas implementadas:
# - MUNICIPIO: fonte oficial municipal, indicador unico = Homicidio doloso, unidade = Vitimas.
# - UF: fonte oficial estadual, multiplos tipos de crime, abas separadas para Ocorrencias e Vitimas.
# - Sexo da vitima: exibido somente no nivel UF, a partir da aba Vitimas.

import os
import re
import tempfile
import unicodedata
from datetime import datetime

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

APP_VERSION = "V17_FINAL_OFICIAL_20260630"

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

MES_LABEL = {1: "Jan", 2: "Fev", 3: "Mar", 4: "Abr", 5: "Mai", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Set", 10: "Out", 11: "Nov", 12: "Dez"}
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
        min-height: 140px;
    }
    .metric-card h4 { margin-top: 0; color: #0b2239; font-size: 1.0rem; }
    .metric-card h2 { margin: 0; font-size: 1.65rem; }
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

# -----------------------------------------------------------------------------
# Utilitarios
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
    if not s or s.upper() in ["NAN", "NONE", "NULL"]:
        return 0.0
    if re.search(r"\d+[,]\d+$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return float(s)
    except Exception:
        return 0.0


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


def mes_para_numero(valor):
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)) and not pd.isna(valor):
        n = int(valor)
        return n if 1 <= n <= 12 else None
    s = normalizar_texto(valor)
    return MES_ORDEM.get(s)


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


def padronizar_crime(valor):
    if pd.isna(valor) or str(valor).strip() == "":
        return "Nao informado"
    return str(valor).strip().title()


def padronizar_sexo(valor):
    s = normalizar_texto(valor)
    if not s or s in ["NAN", "NONE", "NULL"]:
        return "Sexo NI"
    if "FEM" in s:
        return "Feminino"
    if "MASC" in s:
        return "Masculino"
    if "NI" in s or "NAO INFORM" in s or "NAO IDENT" in s or "IGN" in s or "INDETER" in s:
        return "Sexo NI"
    return str(valor).strip().title()


def localizar_coluna(df, candidatos):
    mapa = {normalizar_coluna(c): c for c in df.columns}
    for cand in candidatos:
        cand_norm = normalizar_coluna(cand)
        if cand_norm in mapa:
            return mapa[cand_norm]
    return None


def fmt_int(valor):
    try:
        return f"{int(round(float(valor))):,}".replace(",", ".")
    except Exception:
        return "0"


def card(titulo, valor, subtitulo="", cor="#1c2d42"):
    st.markdown(
        f"""
        <div class="metric-card" style="border-left-color:{cor};">
            <h4>{titulo}</h4>
            <h2 style="color:{cor};">{valor}</h2>
            <p>{subtitulo}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def grafico_barra_serie(serie, titulo=None):
    if titulo:
        st.write(f"**{titulo}**")
    if serie is None or len(serie) == 0 or float(pd.to_numeric(serie, errors="coerce").fillna(0).sum()) == 0:
        st.info("Sem dados suficientes para gerar este grafico.")
    else:
        st.bar_chart(serie)


def grafico_linha_mensal(df, coluna_valor, titulo):
    st.write(f"**{titulo}**")
    if df.empty or coluna_valor not in df.columns:
        st.info("Sem dados para evolucao mensal.")
        return
    base = df.copy()
    base["MES_ORDEM"] = pd.to_numeric(base["MES_ORDEM"], errors="coerce")
    base = base.dropna(subset=["MES_ORDEM"])
    if base.empty:
        st.info("Sem mes valido para evolucao mensal.")
        return
    mensal = base.groupby("MES_ORDEM", as_index=False)[coluna_valor].sum()
    mensal["MES_LABEL"] = mensal["MES_ORDEM"].astype(int).map(MES_LABEL)
    mensal = mensal.sort_values("MES_ORDEM").set_index("MES_LABEL")[coluna_valor]
    if mensal.empty or mensal.sum() == 0:
        st.info("Sem valores para evolucao mensal.")
    else:
        st.line_chart(mensal)


# -----------------------------------------------------------------------------
# Download/cache
# -----------------------------------------------------------------------------
def cache_dir():
    path = os.path.join(tempfile.gettempdir(), "sinesp_mjsp_v17_final")
    os.makedirs(path, exist_ok=True)
    return path


@st.cache_data(ttl=86400, show_spinner=False)
def baixar_arquivo_cacheado(url, nome_arquivo):
    caminho = os.path.join(cache_dir(), nome_arquivo)
    if os.path.exists(caminho) and os.path.getsize(caminho) > 1000:
        return caminho
    resp = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=120, verify=False)
    resp.raise_for_status()
    with open(caminho, "wb") as f:
        f.write(resp.content)
    return caminho


@st.cache_data(ttl=86400, show_spinner=False)
def carregar_municipal_uf(uf):
    caminho = baixar_arquivo_cacheado(URL_MJSP_MUNICIPIOS, f"municipios_{APP_VERSION}.xlsx")
    try:
        df = pd.read_excel(caminho, sheet_name=uf, dtype=object, engine="openpyxl")
    except Exception as e:
        return pd.DataFrame(), {"ok": False, "erro": str(e), "arquivo": caminho, "aba": uf}

    df.columns = [normalizar_coluna(c) for c in df.columns]

    col_cod = localizar_coluna(df, ["COD_IBGE", "CODIGO_MUNICIPIO", "Cód_IBGE"])
    col_mun = localizar_coluna(df, ["MUNICIPIO", "Município"])
    col_uf = localizar_coluna(df, ["SIGLA_UF", "UF"])
    col_reg = localizar_coluna(df, ["REGIAO", "Região"])
    col_mesano = localizar_coluna(df, ["MES_ANO", "MÊS/ANO", "Mês/Ano"])
    col_vit = localizar_coluna(df, ["VITIMAS", "Vítimas"])

    if not col_mun or not col_vit or not col_mesano:
        return pd.DataFrame(), {"ok": False, "erro": "Colunas obrigatorias nao localizadas", "colunas": df.columns.tolist()}

    dt = df[col_mesano].apply(parse_data_mes_ano)
    saida = pd.DataFrame({
        "CODIGO_MUNICIPIO": df[col_cod].astype(str) if col_cod else "",
        "MUNICIPIO": df[col_mun].apply(padronizar_municipio),
        "MUNICIPIO_CHAVE": df[col_mun].apply(chave_filtro),
        "UF": df[col_uf].apply(uf_para_sigla) if col_uf else uf,
        "REGIAO": df[col_reg].astype(str) if col_reg else "",
        "MES_ANO": dt,
        "ANO": dt.dt.year,
        "MES_ORDEM": dt.dt.month,
        "MES_NOME": dt.dt.month.map(MES_LABEL),
        "TIPO_CRIME": "Homicidio doloso",
        "METRICA_FONTE": "VITIMAS",
        "VITIMAS": df[col_vit].apply(converter_numero),
        "OCORRENCIAS": 0.0,
        "FONTE_PROCESSADA": "MJSP/SINESP - Municipios XLSX",
    })
    saida = saida.dropna(subset=["ANO", "MES_ORDEM"])
    saida["ANO"] = saida["ANO"].astype(int)
    saida["MES_ORDEM"] = saida["MES_ORDEM"].astype(int)
    return saida, {"ok": True, "linhas": int(len(saida)), "arquivo": caminho, "aba": uf}


@st.cache_data(ttl=86400, show_spinner=False)
def carregar_uf_base():
    caminho = baixar_arquivo_cacheado(URL_MJSP_UF, f"uf_{APP_VERSION}.xlsx")
    frames = []
    diagnostico = {"ok": True, "arquivo": caminho, "abas": []}

    try:
        xls = pd.ExcelFile(caminho, engine="openpyxl")
    except Exception as e:
        return pd.DataFrame(), {"ok": False, "erro": str(e), "arquivo": caminho}

    for aba in xls.sheet_names:
        try:
            df = pd.read_excel(caminho, sheet_name=aba, dtype=object, engine="openpyxl")
        except Exception as e:
            diagnostico["abas"].append({"aba": aba, "erro": str(e)})
            continue

        df.columns = [normalizar_coluna(c) for c in df.columns]
        col_uf = localizar_coluna(df, ["UF"])
        col_crime = localizar_coluna(df, ["TIPO_CRIME", "Tipo Crime"])
        col_ano = localizar_coluna(df, ["ANO"])
        col_mes = localizar_coluna(df, ["MES", "Mês"])
        col_occ = localizar_coluna(df, ["OCORRENCIAS", "Ocorrências"])
        col_vit = localizar_coluna(df, ["VITIMAS", "Vítimas"])
        col_sexo = localizar_coluna(df, ["SEXO_DA_VITIMA", "Sexo da Vítima", "SEXO"])

        if not col_uf or not col_crime or not col_ano or not col_mes:
            diagnostico["abas"].append({"aba": aba, "erro": "colunas base ausentes", "colunas": df.columns.tolist()})
            continue

        metrica = "VITIMAS" if col_vit else "OCORRENCIAS"
        valor_col = col_vit if col_vit else col_occ
        if not valor_col:
            diagnostico["abas"].append({"aba": aba, "erro": "sem coluna de valor"})
            continue

        temp = pd.DataFrame({
            "UF": df[col_uf].apply(uf_para_sigla),
            "UF_NOME": df[col_uf].astype(str),
            "TIPO_CRIME": df[col_crime].apply(padronizar_crime),
            "TIPO_CRIME_CHAVE": df[col_crime].apply(chave_filtro),
            "ANO": pd.to_numeric(df[col_ano], errors="coerce"),
            "MES_ORDEM": df[col_mes].apply(mes_para_numero),
            "MES_NOME": "",
            "METRICA_FONTE": metrica,
            "OCORRENCIAS": df[col_occ].apply(converter_numero) if col_occ else 0.0,
            "VITIMAS": df[col_vit].apply(converter_numero) if col_vit else 0.0,
            "SEXO_DA_VITIMA": df[col_sexo].apply(padronizar_sexo) if col_sexo else "Nao se aplica",
            "ABA_ORIGEM": normalizar_texto(aba),
            "FONTE_PROCESSADA": "MJSP/SINESP - UF XLSX",
        })
        temp = temp.dropna(subset=["ANO", "MES_ORDEM"])
        temp["ANO"] = temp["ANO"].astype(int)
        temp["MES_ORDEM"] = temp["MES_ORDEM"].astype(int)
        temp["MES_NOME"] = temp["MES_ORDEM"].map(MES_LABEL)
        frames.append(temp)
        diagnostico["abas"].append({"aba": aba, "linhas": int(len(temp)), "metrica": metrica})

    if not frames:
        return pd.DataFrame(), diagnostico
    final = pd.concat(frames, ignore_index=True)
    return final, diagnostico


@st.cache_data(ttl=86400, show_spinner=False)
def lista_municipios_uf(uf):
    df, diag = carregar_municipal_uf(uf)
    if df.empty:
        return ["Todos os municipios"]
    nomes = sorted(df["MUNICIPIO"].dropna().unique().tolist())
    return ["Todos os municipios"] + nomes


@st.cache_data(ttl=86400, show_spinner=False)
def anos_municipal_uf(uf):
    df, diag = carregar_municipal_uf(uf)
    if df.empty:
        return ["Todos os anos"]
    anos = sorted([int(x) for x in df["ANO"].dropna().unique()], reverse=True)
    return ["Todos os anos"] + [str(a) for a in anos]


@st.cache_data(ttl=86400, show_spinner=False)
def anos_uf():
    df, diag = carregar_uf_base()
    if df.empty:
        return ["Todos os anos"]
    anos = sorted([int(x) for x in df["ANO"].dropna().unique()], reverse=True)
    return ["Todos os anos"] + [str(a) for a in anos]


@st.cache_data(ttl=86400, show_spinner=False)
def crimes_uf():
    df, diag = carregar_uf_base()
    if df.empty:
        return ["Todos os indicadores"]
    crimes = sorted(df["TIPO_CRIME"].dropna().unique().tolist())
    return ["Todos os indicadores"] + crimes


# -----------------------------------------------------------------------------
# Interface
# -----------------------------------------------------------------------------
st.sidebar.title("🛡️ Filtros de Seguranca")
st.sidebar.caption(f"Versao: {APP_VERSION}")

nivel = st.sidebar.radio(
    "Nivel de analise:",
    ["Municipio - homicidio doloso/vitimas", "UF - varios tipos de crime"],
)
uf_sel = st.sidebar.selectbox("Selecione o Estado:", sorted(UFS), index=sorted(UFS).index("MG"))

if nivel.startswith("Municipio"):
    municipios = lista_municipios_uf(uf_sel)
    municipio_sel = st.sidebar.selectbox("Selecione o Municipio:", municipios)
    st.sidebar.markdown(
        """
        <div class="warning-box">
        <b>Indicador municipal fixo:</b><br>
        Homicidio doloso / Vitimas.<br><br>
        A fonte nacional municipal do MJSP/SINESP nao disponibiliza, nesse XLSX, todos os tipos de crime por municipio.
        </div>
        """,
        unsafe_allow_html=True,
    )
    indicador_sel = "Homicidio doloso"
    metrica_principal = "VITIMAS"
    anos_opcoes = anos_municipal_uf(uf_sel)
else:
    municipio_sel = None
    metrica_principal = st.sidebar.radio("Metrica principal:", ["OCORRENCIAS", "VITIMAS"], format_func=lambda x: "Ocorrencias" if x == "OCORRENCIAS" else "Vitimas")
    indicador_sel = st.sidebar.selectbox("Indicador / Tipo de Crime:", crimes_uf())
    anos_opcoes = anos_uf()

ano_sel = st.sidebar.selectbox("Ano de Referencia:", anos_opcoes)
mes_label = st.sidebar.selectbox("Mes:", list(MESES_DISPLAY.keys()))
mes_sel = MESES_DISPLAY[mes_label]

with st.sidebar.form("form_consulta"):
    submit = st.form_submit_button("🔍 Consultar Indicadores")

st.markdown(
    f"""
    <div class="header-seguranca">
        <h1>Painel de Ocorrencias Criminais - SINESP/MJSP</h1>
        <p>Monitoramento territorial de indicadores de seguranca publica</p>
        <small>{APP_VERSION}</small>
    </div>
    """,
    unsafe_allow_html=True,
)

if not submit:
    st.info("Escolha os filtros na barra lateral e clique em Consultar Indicadores.")
    st.stop()

# -----------------------------------------------------------------------------
# Processamento - Municipio
# -----------------------------------------------------------------------------
if nivel.startswith("Municipio"):
    with st.spinner("Carregando base oficial municipal do MJSP/SINESP..."):
        df_base, diag = carregar_municipal_uf(uf_sel)

    df = df_base.copy()
    etapas = []
    etapas.append({"etapa": "Base municipal da UF", "linhas": len(df)})

    if municipio_sel and municipio_sel != "Todos os municipios":
        chave = chave_filtro(municipio_sel)
        df = df[df["MUNICIPIO_CHAVE"] == chave].copy()
        etapas.append({"etapa": f"Filtro municipio = {municipio_sel}", "linhas": len(df)})
    else:
        etapas.append({"etapa": "Todos os municipios da UF", "linhas": len(df)})

    if ano_sel != "Todos os anos":
        df = df[df["ANO"] == int(ano_sel)].copy()
        etapas.append({"etapa": f"Filtro ano = {ano_sel}", "linhas": len(df)})
    if mes_sel is not None:
        df = df[df["MES_ORDEM"] == int(mes_sel)].copy()
        etapas.append({"etapa": f"Filtro mes = {mes_label}", "linhas": len(df)})

    localidade = f"{municipio_sel} - {uf_sel}" if municipio_sel != "Todos os municipios" else f"Todos os municipios - {uf_sel}"
    periodo = f"{ano_sel} | {mes_label}"

    total_vitimas = df["VITIMAS"].sum() if not df.empty else 0
    municipios_unicos = df["MUNICIPIO"].nunique() if not df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        card("🚨 Vitimas de homicidio doloso", fmt_int(total_vitimas), "Unidade principal da base municipal", "#d9534f")
    with c2:
        card("📍 Municipios no recorte", fmt_int(municipios_unicos), "Municipios com linhas retornadas", "#1c2d42")
    with c3:
        card("📋 Linhas retornadas", fmt_int(len(df)), "Controle da consulta", "#f0ad4e")
    with c4:
        card("📌 Localidade", localidade, periodo, "#1c2d42")

    st.markdown(
        """
        <div class="method-box">
        <b>Leitura metodologica municipal:</b> a base oficial municipal do MJSP/SINESP usada neste app apresenta o indicador
        <b>Homicidio doloso</b> com unidade de medida <b>Vitimas</b>. Por isso, no nivel municipal nao ha filtro de tipo de crime,
        nao ha ocorrencias e nao ha sexo da vitima.
        </div>
        """,
        unsafe_allow_html=True,
    )

    if df.empty:
        st.error("Nao foram encontrados registros para a combinacao selecionada.")
        with st.expander("Diagnostico da consulta"):
            st.json(diag)
            st.dataframe(pd.DataFrame(etapas), width="stretch")
        st.stop()

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Painel Estatistico", "📋 Dados Tratados", "⚙️ Diagnostico", "📥 Exportacao"])

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            grafico_linha_mensal(df, "VITIMAS", f"Evolucao mensal - {ano_sel} (Vitimas de homicidio doloso)")
        with c2:
            st.write("**Indicador disponivel na base municipal**")
            st.info("Homicidio doloso / Vitimas. A fonte municipal nao traz distribuicao por outros tipos de crime.")

        if municipio_sel == "Todos os municipios":
            st.markdown("### Ranking municipal - vitimas de homicidio doloso")
            ranking = df.groupby("MUNICIPIO")["VITIMAS"].sum().sort_values(ascending=False).head(20)
            grafico_barra_serie(ranking, None)
        else:
            st.info("Ranking municipal oculto porque um municipio especifico esta selecionado.")

    with tab2:
        cols = ["ANO", "MES_NOME", "UF", "MUNICIPIO", "TIPO_CRIME", "VITIMAS", "FONTE_PROCESSADA"]
        st.dataframe(df[cols].sort_values(["ANO", "MES_ORDEM", "MUNICIPIO"]), width="stretch")

    with tab3:
        st.json(diag)
        st.dataframe(pd.DataFrame(etapas), width="stretch")
        st.write("Amostra filtrada")
        st.dataframe(df.head(30), width="stretch")

    with tab4:
        csv = df.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig")
        st.download_button("Baixar CSV municipal filtrado", csv, f"sinesp_municipal_{uf_sel}_{ano_sel}.csv", "text/csv")

# -----------------------------------------------------------------------------
# Processamento - UF
# -----------------------------------------------------------------------------
else:
    with st.spinner("Carregando base oficial estadual do MJSP/SINESP..."):
        df_base, diag = carregar_uf_base()

    df_uf = df_base[df_base["UF"] == uf_sel].copy()
    etapas = [{"etapa": "Base UF selecionada", "linhas": len(df_uf)}]

    if ano_sel != "Todos os anos":
        df_uf = df_uf[df_uf["ANO"] == int(ano_sel)].copy()
        etapas.append({"etapa": f"Filtro ano = {ano_sel}", "linhas": len(df_uf)})
    if mes_sel is not None:
        df_uf = df_uf[df_uf["MES_ORDEM"] == int(mes_sel)].copy()
        etapas.append({"etapa": f"Filtro mes = {mes_label}", "linhas": len(df_uf)})
    if indicador_sel != "Todos os indicadores":
        chave_crime = chave_filtro(indicador_sel)
        df_uf = df_uf[df_uf["TIPO_CRIME_CHAVE"] == chave_crime].copy()
        etapas.append({"etapa": f"Filtro tipo de crime = {indicador_sel}", "linhas": len(df_uf)})

    df_metric = df_uf[df_uf["METRICA_FONTE"] == metrica_principal].copy()
    df_occ = df_uf[df_uf["METRICA_FONTE"] == "OCORRENCIAS"].copy()
    df_vit = df_uf[df_uf["METRICA_FONTE"] == "VITIMAS"].copy()

    total_occ = df_occ["OCORRENCIAS"].sum() if not df_occ.empty else 0
    total_vit = df_vit["VITIMAS"].sum() if not df_vit.empty else 0
    sexo = df_vit.groupby("SEXO_DA_VITIMA")["VITIMAS"].sum() if not df_vit.empty else pd.Series(dtype=float)
    vit_masc = sexo.get("Masculino", 0)
    vit_fem = sexo.get("Feminino", 0)
    vit_ni = sexo.get("Sexo NI", 0)

    localidade = f"Estado {uf_sel} - agregado estadual"
    periodo = f"{ano_sel} | {mes_label}"

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        card("🚨 Ocorrencias", fmt_int(total_occ), "Aba Ocorrencias", "#d9534f")
    with c2:
        card("👥 Vitimas", fmt_int(total_vit), "Aba Vitimas", "#1c2d42")
    with c3:
        card("♂ Masculino", fmt_int(vit_masc), "Vitimas por sexo", "#1c2d42")
    with c4:
        card("♀ Feminino", fmt_int(vit_fem), "Vitimas por sexo", "#7a3db8")
    with c5:
        card("? Sexo NI", fmt_int(vit_ni), "Nao informado/identificado", "#f0ad4e")

    st.markdown(
        f"""
        <div class="method-box">
        <b>Leitura metodologica estadual:</b> a base UF possui uma estrutura de <b>Ocorrencias</b> e outra de <b>Vitimas</b>.
        A analise por sexo utiliza exclusivamente a estrutura de <b>Vitimas</b>. Localidade: <b>{localidade}</b>. Periodo: <b>{periodo}</b>.
        </div>
        """,
        unsafe_allow_html=True,
    )

    if df_uf.empty:
        st.error("Nao foram encontrados registros para a combinacao selecionada.")
        with st.expander("Diagnostico da consulta"):
            st.json(diag)
            st.dataframe(pd.DataFrame(etapas), width="stretch")
        st.stop()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Painel Estatistico", "👥 Vitimas por sexo", "📋 Dados Tratados", "⚙️ Diagnostico", "📥 Exportacao"])

    with tab1:
        c1, c2 = st.columns(2)
        valor = "OCORRENCIAS" if metrica_principal == "OCORRENCIAS" else "VITIMAS"
        label = "Ocorrencias" if metrica_principal == "OCORRENCIAS" else "Vitimas"
        with c1:
            grafico_linha_mensal(df_metric, valor, f"Evolucao mensal - {ano_sel} ({label})")
        with c2:
            serie_crime = df_metric.groupby("TIPO_CRIME")[valor].sum().sort_values(ascending=False)
            grafico_barra_serie(serie_crime, f"Distribuicao por indicador/tipo de crime ({label})")

        st.markdown("### Analise complementar de vitimas")
        c3, c4 = st.columns(2)
        with c3:
            serie_sexo = df_vit.groupby("SEXO_DA_VITIMA")["VITIMAS"].sum().sort_values(ascending=False)
            grafico_barra_serie(serie_sexo, "Vitimas por sexo")
        with c4:
            serie_vit_crime = df_vit.groupby("TIPO_CRIME")["VITIMAS"].sum().sort_values(ascending=False)
            grafico_barra_serie(serie_vit_crime, "Vitimas por tipo de crime")

    with tab2:
        if df_vit.empty:
            st.warning("A estrutura de vitimas nao retornou dados para este filtro.")
        else:
            resumo_sexo = df_vit.groupby("SEXO_DA_VITIMA", as_index=False)["VITIMAS"].sum().sort_values("VITIMAS", ascending=False)
            total = resumo_sexo["VITIMAS"].sum()
            resumo_sexo["PARTICIPACAO_%"] = (resumo_sexo["VITIMAS"] / total * 100).round(2) if total else 0
            st.subheader("Resumo por sexo da vitima")
            st.dataframe(resumo_sexo, width="stretch")

            st.subheader("Tipo de crime x sexo da vitima")
            tabela = pd.pivot_table(
                df_vit,
                index="TIPO_CRIME",
                columns="SEXO_DA_VITIMA",
                values="VITIMAS",
                aggfunc="sum",
                fill_value=0,
            )
            tabela["TOTAL"] = tabela.sum(axis=1)
            st.dataframe(tabela.sort_values("TOTAL", ascending=False), width="stretch")

    with tab3:
        st.subheader("Dados usados na metrica principal")
        st.dataframe(df_metric.sort_values(["ANO", "MES_ORDEM", "TIPO_CRIME"]), width="stretch")
        with st.expander("Ocorrencias filtradas"):
            st.dataframe(df_occ.sort_values(["ANO", "MES_ORDEM", "TIPO_CRIME"]), width="stretch")
        with st.expander("Vitimas filtradas"):
            st.dataframe(df_vit.sort_values(["ANO", "MES_ORDEM", "TIPO_CRIME", "SEXO_DA_VITIMA"]), width="stretch")

    with tab4:
        st.json(diag)
        st.dataframe(pd.DataFrame(etapas), width="stretch")
        st.write("Metricas disponiveis no recorte")
        st.write(sorted(df_uf["METRICA_FONTE"].dropna().unique().tolist()))
        st.write("Sexos disponiveis na estrutura de vitimas")
        st.write(sorted(df_vit["SEXO_DA_VITIMA"].dropna().unique().tolist()) if not df_vit.empty else [])
        st.write("Amostra da estrutura de vitimas")
        st.dataframe(df_vit.head(30), width="stretch")

    with tab5:
        st.download_button(
            "Baixar CSV - metrica principal",
            df_metric.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
            f"sinesp_uf_{uf_sel}_{metrica_principal}_{ano_sel}.csv",
            "text/csv",
        )
        st.download_button(
            "Baixar CSV - vitimas por sexo/tipo de crime",
            df_vit.to_csv(index=False, sep=";", decimal=",", encoding="utf-8-sig"),
            f"sinesp_uf_{uf_sel}_vitimas_{ano_sel}.csv",
            "text/csv",
        )
