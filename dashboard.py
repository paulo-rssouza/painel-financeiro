import streamlit as st
import pandas as pd
import gspread
import subprocess
import numpy as np
from datetime import datetime
from google.oauth2.service_account import Credentials
import plotly.graph_objects as go
import re
import time

st.set_page_config(page_title="Controle Financeiro Pro", layout="wide")

# ===============================================
# CONEXÃO COM O GOOGLE SHEETS (VIA SECRETS)
# ===============================================
@st.cache_resource
def get_ws():
    credentials_dict = dict(st.secrets["gcp_service_account"])
    scope = ["https://www.googleapis.com/auth/spreadsheets"]
    cred = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(cred)
    # ATENÇÃO: Substitua abaixo pelo nome exato da sua planilha no Google Drive
    return client.open("Despesas Cartões").sheet1

# Tenta carregar a aba da planilha
try:
    ws = get_ws()
except Exception as e:
    st.error(f"Erro ao conectar com a planilha: {e}")
    st.stop()

# ===============================================
# CSS GLOBAL
# ===============================================
st.markdown("""
<style>
    div[data-testid="stDataFrame"] td,
    div[data-testid="stDataFrame"] th {
        white-space: nowrap !important;
        padding-left: 14px !important;
        padding-right: 14px !important;
    }
</style>
""", unsafe_allow_html=True)

# ===============================================
# FUNÇÕES DE APOIO
# ===============================================
def parse_float(val):
    try:
        if pd.isna(val) or val == "": return 0.0
        if isinstance(val, (int, float)): return float(val)
        v = str(val).replace("R$", "").replace(" ", "").strip()
        
        if "." in v and "," in v:
            v = v.replace(".", "").replace(",", ".")
        elif "," in v:
            v = v.replace(",", ".")
        elif "." in v:
            parts = v.split(".")
            if len(parts[-1]) == 3:
                v = v.replace(".", "")
                
        return float(v)
    except: return 0.0

# ===============================================
# CORPO PRINCIPAL DO PAINEL
# ===============================================
st.title("💰 Controle Financeiro Pro")
st.write("Painel conectado com sucesso ao Google Sheets!")

# Exemplo básico para testar se os dados da planilha estão carregando
try:
    dados = ws.get_all_records()
    if dados:
        df = pd.DataFrame(dados)
        st.success(f"Planilha carregada com sucesso! Total de registros: {len(df)}")
        st.dataframe(df.head())
    else:
        st.warning("A planilha está vazia ou os dados não puderam ser lidos no formato de tabela.")
except Exception as e:
    st.error(f"Erro ao ler os dados da planilha: {e}")
