# VERSION_FINAL_PRODUCAO_SEGURANCA_DASHBOARD_V2
# App Streamlit para dados de seguranca publica
# Fonte principal: base oficial MJSP/SINESP em XLSX
# Fonte experimental: API comunitaria rayonnunes/api_seguranca_publica

import io
import re
import unicodedata
from datetime import datetime

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


def mes_para_ordem(valor):
    if pd.isna(valor):
        return None
    texto_original = str(valor).strip()
    if texto_original == "":
        return None

    # Trata datas completas, caso alguma planilha venha com referencia temporal.
    dt = pd.to_datetime(texto_original, errors="coerce", dayfirst=True)
    if pd.notna(dt):
        return int(dt.month)

    texto = normalizar_valor(texto_original)
    texto_sem_ponto = texto.replace(".", "")
    if texto_sem_ponto in MES_NORMALIZADO:
        return MES_NORMALIZADO[texto_sem_ponto]

    # Casos como "01 - Janeiro" ou "1/Janeiro".
    match = re.search(r"\b(0?[1-9]|1[0-2])\b", texto_sem_ponto)
    if match:
        return int(match.group(1))

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

    if "TIPO_CRIME" not in df.columns and "CRIME" in df.columns:
        df["TIPO_CRIME"] = df["CRIME"]

    if "CRIME" in df.columns and fonte == "api_rayonnunes":
        df["TIPO_CRIME"] = df["CRIME"].astype(str).map(REVERSE_CRIMES_API).fillna(df.get("TIPO_CRIME", "Outros"))

    if "TIPO_CRIME" in df.columns:
        df["TIPO_CRIME"] = df["TIPO_CRIME"].astype(str).str.strip()
        df.loc[df["TIPO_CRIME"].isin(["", "nan", "None"]), "TIPO_CRIME"] = "Nao informado"
    else:
        df["TIPO_CRIME"] = "Nao informado"

    if "ANO" in df.columns:
        df["ANO"] = df["ANO"].astype(str).str.extract(r"(\d{4})", expand=False).fillna(df["ANO"].astype(str))
    else:
        df["ANO"] = "Nao informado"

    if "MES" in df.columns:
        df["MES_ORDEM"] = df["MES"].map(mes_para_ordem)
        df["MES_NOME"] = df["MES_ORDEM"].map(MES_LABEL).fillna(df["MES"].astype(str))
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

    df["FONTE_PROCESSADA"] = fonte
    df = df.loc[:, ~df.columns.duplicated()].copy()
    return df


@st.cache_data(ttl=86400, show_spinner=False)
def carregar_base_oficial(tipo_base):
    url = URL_MJSP_MUNICIPIOS if tipo_base == "municipios" else URL_MJSP_UF
    try:
        resp = requests.get(
            url,
            timeout=180,
            verify=False,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        if resp.status_code != 200:
            return pd.DataFrame(), {
                "ok": False,
                "erro": f"HTTP {resp.status_code}",
                "url": url,
                "texto": resp.text[:1500],
            }

        conteudo = io.BytesIO(resp.content)
        df = pd.read_excel(conteudo, dtype=str, engine="openpyxl")
        df = padronizar_base_seguranca(df, f"oficial_mjsp_{tipo_base}")
        return df, {"ok": True, "url": url, "linhas": len(df), "colunas": list(df.columns)}

    except Exception as e:
        return pd.DataFrame(), {
            "ok": False,
            "erro": str(e),
            "url": url,
            "texto": "",
        }


@st.cache_data(ttl=3600, show_spinner=False)
def consultar_api_rayonnunes(uf, municipio, crime_id, ano, mes_num):
    params_base = {
        "uf": uf.lower(),
        "per_page": 1000,
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
    max_paginas = 100

    for pagina in range(1, max_paginas + 1):
        params = dict(params_base)
        params["page"] = pagina

        try:
            resp = requests.get(
                API_RAYONNUNES,
                params=params,
                timeout=45,
                headers={"User-Agent": "Mozilla/5.0"},
            )
            chamadas.append({"pagina": pagina, "url": resp.url, "status": resp.status_code})

            if resp.status_code != 200:
                erros.append(f"Pagina {pagina}: HTTP {resp.status_code} - {resp.text[:500]}")
                break

            try:
                payload = resp.json()
            except Exception:
                erros.append(f"Pagina {pagina}: resposta nao veio em JSON - {resp.text[:500]}")
                break

            dados = payload.get("data", []) if isinstance(payload, dict) else []
            if not dados:
                break

            frames.append(pd.DataFrame(dados))

            # A API limita a 1000 itens por pagina. Se vier menos, acabou.
            if len(dados) < 1000:
                break

        except requests.exceptions.Timeout:
            erros.append(f"Pagina {pagina}: tempo limite excedido.")
            break
        except requests.exceptions.RequestException as e:
            erros.append(f"Pagina {pagina}: falha de conexao - {e}")
            break
        except Exception as e:
            erros.append(f"Pagina {pagina}: erro inesperado - {e}")
            break

    if frames:
        df = pd.concat(frames, ignore_index=True)
        df = padronizar_base_seguranca(df, "api_rayonnunes")
    else:
        df = pd.DataFrame()

    meta = {"chamadas": chamadas, "erros": erros, "params": params_base}
    return df, meta


def filtrar_base(df, uf, municipio, crime, ano, mes_num):
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "UF" in out.columns and uf:
        out = out[out["UF"].astype(str).str.upper() == uf.upper()].copy()

    if municipio and "MUNICIPIO" in out.columns:
        chave_mun = chave_comparacao(municipio)
        out = out[out["MUNICIPIO"].map(chave_comparacao) == chave_mun].copy()

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


def mostrar_alerta_colunas(df, fonte_label):
    colunas_necessarias = ["UF", "ANO", "MES_ORDEM", "TIPO_CRIME", "OCORRENCIAS"]
    faltantes = [c for c in colunas_necessarias if c not in df.columns]
    if faltantes:
        st.warning(
            f"A base {fonte_label} foi carregada, mas faltam colunas padronizadas: {', '.join(faltantes)}. "
            "Os filtros/graficos podem ficar limitados. Veja a aba de diagnostico."
        )


def plotar_serie_mensal(df):
    if df.empty or "MES_ORDEM" not in df.columns:
        st.info("Nao ha dados mensais suficientes para montar a serie temporal.")
        return
    base = df.copy()
    base["MES_ORDEM"] = pd.to_numeric(base["MES_ORDEM"], errors="coerce")
    base = base.dropna(subset=["MES_ORDEM"])
    if base.empty:
        st.info("Nao ha dados mensais validos para montar a serie temporal.")
        return
    serie = (
        base.groupby("MES_ORDEM", dropna=True)["OCORRENCIAS"]
        .sum()
        .reset_index()
        .sort_values("MES_ORDEM")
    )
    serie["MES"] = serie["MES_ORDEM"].astype(int).map(MES_LABEL)
    serie = serie.set_index("MES")["OCORRENCIAS"]
    st.line_chart(serie)


def plotar_crimes(df):
    if df.empty or "TIPO_CRIME" not in df.columns:
        st.info("Nao ha dados suficientes para a distribuicao por tipo de crime.")
        return
    serie = df.groupby("TIPO_CRIME")["OCORRENCIAS"].sum().sort_values(ascending=False).head(15)
    if serie.empty or serie.sum() == 0:
        st.info("Nao ha ocorrencias para a distribuicao por tipo de crime.")
        return
    st.bar_chart(serie)


def plotar_top_municipios(df):
    if df.empty or "MUNICIPIO" not in df.columns:
        st.info("Selecione a base municipal para ver o ranking por municipio.")
        return
    serie = df.groupby("MUNICIPIO")["OCORRENCIAS"].sum().sort_values(ascending=False).head(15)
    if serie.empty or serie.sum() == 0:
        st.info("Nao ha ocorrencias para montar o ranking municipal.")
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

base_oficial = pd.DataFrame()
meta_oficial = {}

if usa_oficial_municipios:
    with st.sidebar.status("Carregando base oficial municipal...", expanded=False) as status:
        base_oficial, meta_oficial = carregar_base_oficial("municipios")
        if meta_oficial.get("ok"):
            status.update(label="Base municipal carregada", state="complete")
        else:
            status.update(label="Falha ao carregar base municipal", state="error")
elif usa_oficial_uf:
    with st.sidebar.status("Carregando base oficial UF...", expanded=False) as status:
        base_oficial, meta_oficial = carregar_base_oficial("uf")
        if meta_oficial.get("ok"):
            status.update(label="Base UF carregada", state="complete")
        else:
            status.update(label="Falha ao carregar base UF", state="error")

uf_sel = st.sidebar.selectbox("Selecione o Estado:", sorted(UFS), index=sorted(UFS).index("MG"))

municipio_sel = "Todos os municipios"
municipio_param = None
if usa_oficial_municipios or usa_api:
    municipios = obter_opcoes_municipios(base_oficial if usa_oficial_municipios else pd.DataFrame(), uf_sel)
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
    ano_opcoes = ["Todos os anos"] + [str(a) for a in range(datetime.now().year, 2014, -1)]
    ano_sel = st.sidebar.selectbox("Ano de Referencia:", ano_opcoes, index=1)
else:
    if not base_oficial.empty:
        mostrar_alerta_colunas(base_oficial, fonte_dados)
    crime_opcoes = obter_opcoes_crimes(base_oficial)
    crime_nome = st.sidebar.selectbox("Indicador / Tipo de Crime:", crime_opcoes)
    crime_param = crime_nome
    ano_sel = st.sidebar.selectbox("Ano de Referencia:", obter_opcoes_anos(base_oficial), index=0)

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
        A API comunitaria e util para testes, mas pode estar fora do ar, defasada ou limitada.
        """
    )
    st.stop()

st.markdown(
    '<div class="header-seguranca"><h1>Painel de Ocorrencias Criminais - SINESP/MJSP</h1>'
    '<p>Monitoramento territorial de indicadores de seguranca publica</p></div>',
    unsafe_allow_html=True,
)

meta_api = {}

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
    if base_oficial.empty:
        st.error("Nao foi possivel carregar a base oficial selecionada.")
        with st.expander("Diagnostico da carga oficial"):
            st.json(meta_oficial)
        st.stop()

    with st.spinner("Filtrando base oficial MJSP/SINESP..."):
        df_consulta = filtrar_base(
            base_oficial,
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
    with st.expander("Diagnostico tecnico"):
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
            if not base_oficial.empty:
                st.write("Colunas padronizadas carregadas:")
                st.code("\n".join(base_oficial.columns.astype(str).tolist()))
                st.write("Amostra da base carregada:")
                st.dataframe(base_oficial.head(20), width="stretch")
    st.stop()

# Cards
localidade = titulo_localidade(uf_sel, municipio_param)
total_ocorrencias = df_consulta["OCORRENCIAS"].sum() if "OCORRENCIAS" in df_consulta.columns else 0
total_vitimas = df_consulta["VITIMAS"].sum() if "VITIMAS" in df_consulta.columns else 0
total_vitimas_mun = df_consulta["VITIMAS_MUNICIPIO"].sum() if "VITIMAS_MUNICIPIO" in df_consulta.columns else 0
registros = len(df_consulta)

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown(
        f'<div class="metric-card"><h4>🚨 Ocorrencias</h4>'
        f'<h2 style="color:#d9534f; margin:0;">{formatar_inteiro(total_ocorrencias)}</h2>'
        '<p>Registros no filtro</p></div>',
        unsafe_allow_html=True,
    )
with c2:
    st.markdown(
        f'<div class="metric-card"><h4>👥 Vitimas</h4>'
        f'<h2 style="color:#1c2d42; margin:0;">{formatar_inteiro(total_vitimas)}</h2>'
        '<p>Quando informado pela fonte</p></div>',
        unsafe_allow_html=True,
    )
with c3:
    valor_vitima_mun = total_vitimas_mun if total_vitimas_mun > 0 else registros
    label_vitima_mun = "Vitimas municipio" if total_vitimas_mun > 0 else "Linhas retornadas"
    st.markdown(
        f'<div class="metric-card"><h4>📋 {label_vitima_mun}</h4>'
        f'<h2 style="color:#f0ad4e; margin:0;">{formatar_inteiro(valor_vitima_mun)}</h2>'
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
        st.write(f"**Evolucao mensal - {ano_sel}**")
        plotar_serie_mensal(df_consulta)
    with col_b:
        st.write("**Distribuicao por indicador/tipo de crime**")
        plotar_crimes(df_consulta)

    st.write("**Ranking municipal - top 15**")
    plotar_top_municipios(df_consulta)

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
        st.write("**Metadados da base oficial:**")
        st.json(meta_oficial)
    st.write("**Colunas finais:**")
    st.code("\n".join(df_consulta.columns.astype(str).tolist()))
    st.write("**Amostra da consulta:**")
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
        validacao e consolidacao feita pelos estados e pelo Distrito Federal. Valores podem mudar
        quando a base oficial e atualizada.
        """
    )
