# VERSION_FINAL_PRODUCAO_SEGURANCA_DASHBOARD_V1
import streamlit as st
import pandas as pd
import requests
import urllib3
import unicodedata
import re

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- CONFIGURAÇÃO DE DESIGN DA PÁGINA ---
st.set_page_config(
    page_title="Inteligência Territorial | Segurança Pública",
    page_icon="🛡️",
    layout="wide"
)

st.markdown("""
    <style>
    .header-seguranca { background-color: #1c2d42; padding: 20px; color: white; border-radius: 5px; text-align: center; margin-bottom: 20px; }
    .stButton>button { background-color: #1c2d42; color: white; width: 100%; }
    .metric-card { background-color: #f8f9fa; padding: 15px; border-radius: 10px; border-left: 5px solid #1c2d42; text-align: center; margin-bottom: 15px; box-shadow: 2px 2px 5px rgba(0,0,0,0.05);}
    </style>
""", unsafe_allow_html=True)

# --- DICIONÁRIOS OFICIAIS DA API SINESP ---
MAPA_CRIMES = {
    "Todos os Crimes": None,
    "1 - Estupro": "1",
    "2 - Furto de veículo": "2",
    "3 - Homicídio doloso": "3",
    "4 - Lesão corporal seguida de morte": "4",
    "5 - Roubo a instituição financeira": "5",
    "6 - Roubo de carga": "6",
    "7 - Roubo de veículo": "7",
    "8 - Roubo seguido de morte (latrocínio)": "8",
    "9 - Tentativa de homicídio": "9"
}

REVERSE_CRIMES = {v: k.split(" - ")[1] for k, v in MAPA_CRIMES.items() if v is not None}

UFS = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS","MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC","SP","SE","TO"]

MESES_API = {"Jan": 1, "Fev": 2, "Mar": 3, "Abr": 4, "Mai": 5, "Jun": 6, "Jul": 7, "Ago": 8, "Set": 9, "Out": 10, "Nov": 11, "Dez": 12}

# --- FUNÇÕES DE TRATAMENTO DE TEXTO EXIGIDAS PELA API ---
def normalizar_municipio_api(nome_municipio):
    """ Remove espaços, acentos e joga para minúsculo conforme manual da API SINESP """
    if not nome_municipio: return ""
    nome_alterado = unicodedata.normalize('NFKD', nome_municipio).encode('ASCII', 'ignore').decode('ASCII')
    nome_alterado = re.sub(r'\s+', '', nome_alterado)
    return nome_alterado.lower()

@st.cache_data
def buscar_municipios_ibge(uf_sigla):
    url = f"https://servicodados.ibge.gov.br/api/v1/localidades/estados/{uf_sigla}/municipios"
    try:
        res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10).json()
        return sorted([m['nome'] for m in res])
    except:
        return []

# --- MOTOR DE CONSUMO DA API REAL (SEM SIMULAÇÃO) ---
def consultar_api_sinesp(uf, municipio, crime_id, ano_texto):
    URL_BASE = "http://ec2-54-174-4-15.compute-1.amazonaws.com/api"
    
    params = {
        "uf": uf.lower(),
        "per_page": 1000, 
        "page": 1
    }
    if municipio:
        params["municipio"] = normalizar_municipio_api(municipio)
    if crime_id:
        params["crime"] = crime_id
    if ano_texto:
        params["ano"] = str(ano_texto)

    try:
        response = requests.get(URL_BASE, params=params, timeout=10)
        if response.status_code == 200:
            json_data = response.json()
            if "data" in json_data and len(json_data["data"]) > 0:
                return pd.DataFrame(json_data["data"])
    except Exception:
        pass
    
    return pd.DataFrame()  # Retorna um DataFrame vazio se a API falhar ou não trouxer nada

# --- INTERFACE GRÁFICA ---
st.sidebar.title("🛡️ Filtros de Segurança")

uf_sel = st.sidebar.selectbox("Selecione o Estado:", sorted(UFS), index=UFS.index("CE"))
lista_municipios = buscar_municipios_ibge(uf_sel)
municipio_sel = st.sidebar.selectbox("Selecione o Município:", ["Todos os Municípios"] + lista_municipios)
municipio_param = None if municipio_sel == "Todos os Municípios" else municipio_sel

crime_nome = st.sidebar.selectbox("Classificação do Crime:", list(MAPA_CRIMES.keys()))
crime_param = MAPA_CRIMES[crime_nome]

ano_sel = st.sidebar.selectbox("Ano de Referência:", ["2024", "2023", "2022", "2021", "2020"])

with st.sidebar.form("form_seguranca"):
    submit_btn = st.form_submit_button("🔍 Consultar Indicadores")

# --- PROCESSAMENTO DOS DADOS ---
if submit_btn:
    st.markdown('<div class="header-seguranca"><h1>Painel de Ocorrências Criminais (SINESP)</h1><p>Monitoramento Analítico Geográfico e de Criminalidade</p></div>', unsafe_allow_html=True)
    
    with st.spinner("Conectando à base nacional de segurança pública..."):
        df_bruto = consultar_api_sinesp(uf_sel, municipio_param, crime_param, ano_sel)
        
    if not df_bruto.empty:
        # --- LIMPEZA E TRATAMENTO DA PLANILHA ---
        df_tratado = df_bruto.copy()
        df_tratado.columns = [str(c).upper() for c in df_tratado.columns]
        
        df_tratado["OCORRÊNCIAS"] = pd.to_numeric(df_tratado["OCORRENCIAS"], errors='coerce').fillna(0).astype(int)
        df_tratado["VÍTIMAS MUNICÍPIO"] = pd.to_numeric(df_tratado["VITIMAS_MUNICIPIO"], errors='coerce').fillna(0).astype(int)
        df_tratado["TIPO_CRIME"] = df_tratado["CRIME"].map(REVERSE_CRIMES).fillna("Outros")
        df_tratado["MÊS_NOME"] = df_tratado["MES"].str.capitalize()
        df_tratado["MÊS_ORDEM"] = df_tratado["MÊS_NOME"].map(MESES_API).fillna(13)
        df_tratado["MUNICIPIO"] = df_tratado["MUNICIPIO"].str.title()
        df_tratado["UF"] = df_tratado["UF"].str.upper()
        
        # --- CARDS DE MÉTRICAS ANALÍTICAS ---
        total_casos = df_tratado["OCORRÊNCIAS"].sum()
        
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(f'<div class="metric-card"><h4>🚨 TOTAL DE OCORRÊNCIAS</h4><h2 style="color:#d9534f; margin:0;">{total_casos:,}</h2><p>Registradas no período selecionado</p></div>', unsafe_allow_html=True)
        with c2:
            st.markdown(f'<div class="metric-card"><h4>📍 LOCALIDADE MONITORADA</h4><h2 style="color:#1c2d42; margin:0;">{municipio_sel} - {uf_sel}</h2><p>Filtro Geográfico Aplicado</p></div>', unsafe_allow_html=True)
        with c3:
            st.markdown(f'<div class="metric-card"><h4>📅 ANO CONTEXTO</h4><h2 style="color:#f0ad4e; margin:0;">{ano_sel}</h2><p>Origem: Base Oficial SINESP</p></div>', unsafe_allow_html=True)
            
        st.markdown("---")
        
        # --- ABAS DE VISUALIZAÇÃO ---
        tab1, tab2, tab3 = st.tabs(["📊 Painel Estatístico", "📋 Dados Estruturados", "⚙️ JSON Original da API"])
        
        with tab1:
            col_graph1, col_col_graph2 = st.columns(2)
            
            with col_graph1:
                st.write(f"**Evolução Mensal de Casos — {ano_sel}**")
                df_linha = df_tratado.groupby(["MÊS_ORDEM", "MÊS_NOME"])["OCORRÊNCIAS"].sum().reset_index().sort_values("MÊS_ORDEM")
                df_linha = df_linha.set_index("MÊS_NOME")["OCORRÊNCIAS"]
                st.line_chart(df_linha, color="#d9534f")
                
            with col_col_graph2:
                st.write("**Distribuição por Tipificação de Delito (Top Crimes)**")
                df_barras = df_tratado.groupby("TIPO_CRIME")["OCORRÊNCIAS"].sum().sort_values(ascending=False)
                st.bar_chart(df_barras, color="#1c2d42")
        
        with tab2:
            cols_exibicao = ["ANO", "MÊS_NOME", "UF", "MUNICIPIO", "TIPO_CRIME", "OCORRÊNCIAS"]
            df_final_grid = df_tratado[[c for c in cols_exibicao if c in df_tratado.columns]]
            st.dataframe(df_final_grid, use_container_width=True)
            st.download_button("📥 Exportar Dados Oficiais (CSV)", df_final_grid.to_csv(index=False, sep=';'), f"sinesp_{uf_sel}_{ano_sel}.csv", "text/csv")
            
        with tab3:
            st.info("Abaixo está a estrutura de objetos retornada diretamente da API SINESP:")
            st.json(df_bruto.head(10).to_dict(orient="records"))
    else:
        st.error("🛑 **Falha na Requisição ou Sem Registros:** Não foi possível obter dados da API SINESP. O servidor do Ministério da Justiça pode estar temporariamente indisponível ou não existem ocorrências registradas para essa combinação exata de filtros.")
else:
    st.info("💡 Escolha a localidade e o crime na barra lateral e clique em **Consultar Indicadores** para renderizar os dados reais do SINESP.")
