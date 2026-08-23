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

# ==========================================
# CSS GLOBAL (DESKTOP E MOBILE RESPONSIVO)
# ==========================================
st.markdown("""
    <style>
    div[data-testid="stDataFrame"] td, 
    div[data-testid="stDataFrame"] th {
        white-space: nowrap !important;
        padding-left: 14px !important;
        padding-right: 14px !important;
    }

    @media (max-width: 768px) {
        div[data-testid="stDataFrame"] {
            overflow-x: auto !important;
        }
        .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            padding-top: 2rem !important;
        }
        h1 { font-size: 1.8rem !important; }
        h2 { font-size: 1.4rem !important; }
        h3 { font-size: 1.2rem !important; }
    }
    </style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXÃO E FUNÇÕES DE APOIO
# ==========================================
@st.cache_resource
def get_ws():
    credentials_dict = dict(st.secrets["gcp_service_account"])
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    cred = Credentials.from_service_account_info(credentials_dict, scopes=scope)
    client = gspread.authorize(cred)
    return client.open_by_url("https://docs.google.com/spreadsheets/d/1DosfnqIt8ioBxfXMrpEBT8Md7apgsdhvWg0YswUA-IA/edit")

ws = get_ws()

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

def fmt_br(valor):
    if pd.isna(valor): return ""
    return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def formatar_tabela_br(x):
    try:
        if pd.isna(x): return ""
        return f"R$ {float(x):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return x

def ler_aba(nome):
    try:
        dados = ws.worksheet(nome).get_all_values()
        if len(dados) > 1: return pd.DataFrame(dados[1:], columns=dados[0])
        elif len(dados) == 1: return pd.DataFrame(columns=dados[0])
        return pd.DataFrame()
    except: return pd.DataFrame()

def salvar_tabela_google(df_editado, nome_aba):
    try:
        planilhas = [sheet.title for sheet in ws.worksheets()]
        if nome_aba not in planilhas:
            aba = ws.add_worksheet(title=nome_aba, rows=100, cols=20)
        else:
            aba = ws.worksheet(nome_aba)
            
        aba.clear()
        df_editado = df_editado.fillna("")
        dados = [df_editado.columns.tolist()] + df_editado.values.tolist()
        aba.update(dados, value_input_option="USER_ENTERED")
        return True
    except Exception as e:
        st.error(f"Erro ao salvar a aba {nome_aba}: {e}")
        return False

def mapear_banco(nome_cartao):
    n = str(nome_cartao).upper()
    if "SANTANDER" in n: return "🔴 " + nome_cartao
    if "ITAÚ" in n: return "🟧 " + nome_cartao
    if "NUBANK" in n: return "🟪 " + nome_cartao
    if "MERCADO" in n: return "🟦 " + nome_cartao
    return "🏦 " + nome_cartao

def obter_descricao(row):
    for col in row.index:
        if "desc" in str(col).lower():
            return str(row[col])
    return ""

def estilizar_tendencia(df_pivot):
    estilos = pd.DataFrame('', index=df_pivot.index, columns=df_pivot.columns)
    for i in range(1, len(df_pivot.columns)):
        col_atual = df_pivot.columns[i]
        col_ant = df_pivot.columns[i-1]

        for idx in df_pivot.index:
            val_atual = df_pivot.loc[idx, col_atual]
            val_ant = df_pivot.loc[idx, col_ant]

            if val_atual == 0 and val_ant == 0: continue
            elif val_atual > val_ant: estilos.loc[idx, col_atual] = 'color: #dc3545;'
            elif val_atual < val_ant: estilos.loc[idx, col_atual] = 'color: #28a745;'
    return estilos

def destacar_mes(col):
    if isinstance(col.name, tuple) and len(col.name) > 1:
        if str(col.name[1]).startswith("📍"): return ['background-color: #fffacd'] * len(col)
    elif isinstance(col.name, str):
        if col.name.startswith("📍"): return ['background-color: #fffacd'] * len(col)
    return [''] * len(col)

def classificar_tipo_compra(row):
    cartao = str(row.get("Cartão", "")).upper()
    if "DEBITO" in cartao or "DÉBITO" in cartao:
        return "À Vista / Mês"
        
    p = str(row.get("Parcela", "")).strip()
    if not p or p.lower() == "nan" or p == "1/1":
        return "À Vista / Mês"
    
    if "/" in p:
        try:
            atual, total = p.split("/")
            if int(total.strip()) > 1:
                return "Parcelado"
        except:
            pass
    return "À Vista / Mês"

def normalizar_chave_base(c, d):
    c_norm = re.sub(r'[^a-z0-9]', '', str(c).lower())
    d_norm = re.sub(r'[^a-z0-9]', '', str(d).lower())
    return c_norm + "|" + d_norm

def normalizar_chave(c, d, v):
    c_norm = re.sub(r'[^a-z0-9]', '', str(c).lower())
    d_norm = re.sub(r'[^a-z0-9]', '', str(d).lower())
    v_norm = f"{parse_float(v):.2f}" 
    return c_norm + "|" + d_norm + "|" + v_norm

meses_pt = {1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"}

def sort_data_br(d):
    try: return datetime.strptime(str(d).strip(), "%d/%m/%Y")
    except: return datetime(1900, 1, 1)

# ==========================================
# CARREGAMENTO DE DADOS (CACHE: 10 MINUTOS)
# ==========================================
@st.cache_data(ttl=600)
def carregar_dados():
    df_cartoes = ler_aba("Despesas Cartões")
    df_fixas = ler_aba("Dashboard Fixas")
    
    if not df_cartoes.empty:
        df_cartoes["Origem_Aba"] = "Cartoes"
    if not df_fixas.empty:
        df_fixas["Origem_Aba"] = "Fixas"
        
    df_rec_total = ler_aba("Avaliar Recorrentes")
    
    colunas_padrao_rec = ["Cartão", "Descrição", "Valor (R$)", "Vencimentos Encontrados", "Status", "Motivo do Alerta"]
    if df_rec_total.empty and len(df_rec_total.columns) == 0:
        df_rec_total = pd.DataFrame(columns=colunas_padrao_rec)
    else:
        for col in colunas_padrao_rec:
            if col not in df_rec_total.columns:
                df_rec_total[col] = ""
    
    if not df_rec_total.empty and "Status" in df_rec_total.columns and "Cartão" in df_rec_total.columns and "Descrição" in df_rec_total.columns:
        df_rec_temp = df_rec_total.copy()
        df_rec_temp["Chave"] = df_rec_temp.apply(lambda r: normalizar_chave(r.get("Cartão", ""), r.get("Descrição", ""), r.get("Valor (R$)", "")), axis=1)
        df_rec_temp["_peso"] = df_rec_temp["Status"].apply(lambda s: 1 if "aprovado" in str(s).lower() or "ignorar" in str(s).lower() else 2)
        df_rec_temp = df_rec_temp.sort_values(by=["_peso", "Cartão", "Descrição"])
        df_rec_display = df_rec_temp.drop_duplicates(subset=["Chave"], keep="first").drop(columns=["_peso", "Chave"])
    else:
        df_rec_display = df_rec_total.copy()
    
    df_reclass = ler_aba("Reclassificacao")
    if df_reclass.empty or "Palavras-chave (separadas por vírgula)" not in df_reclass.columns:
        df_reclass = pd.DataFrame([
            {"Palavras-chave (separadas por vírgula)": "pix amanda", "Categoria": "🔄 Repasse"},
            {"Palavras-chave (separadas por vírgula)": "iof, juros, anuidade", "Categoria": "🏦 Taxas Financeiras"},
            {"Palavras-chave (separadas por vírgula)": "ifd, rappi, restaurante, fast food, bar, padaria, osnir, beco sao paulo, beco são paulo", "Categoria": "🍔 Alimentação"},
            {"Palavras-chave (separadas por vírgula)": "mambo, swift, casa do norte, sacolão, sacolao, hortifruti, açougue, acougue, supermercado, carrefour, extra", "Categoria": "🛒 Mercado e Feira"},
            {"Palavras-chave (separadas por vírgula)": "uber, 99, taxi, hoteis.com, metrô, metro, azul, vfs brasil", "Categoria": "🚗 Transporte"},
            {"Palavras-chave (separadas por vírgula)": "condominio, condomínio", "Categoria": "🏠 Moradia"},
            {"Palavras-chave (separadas por vírgula)": "caixa, financiamento, emprestimo reforma ap, empréstimo reforma ap", "Categoria": "🏦 Dívidas e Empréstimos"},
            {"Palavras-chave (separadas por vírgula)": "leroy, telhanorte, madeiramadeira, vmt máquinas, vmt maquinas, obra prima, casadivaferram", "Categoria": "🧱 Casa e Reforma"},
            {"Palavras-chave (separadas por vírgula)": "drogaria, extrafarma, farmacia, pague menos, sephora, beleza na web, nivea, unicpharma", "Categoria": "💊 Saúde e Beleza"},
            {"Palavras-chave (separadas por vírgula)": "prime, globoplay, apple, duolingo, youtube, getpeech, linkedin, premiere, smiles, netflix, spotify, disney, hbo", "Categoria": "📺 Assinaturas"},
            {"Palavras-chave (separadas por vírgula)": "shopee, mercado livre, mercadolivre, amazon, casas bahia, fast shop, magalu, aliexpress, shein, c&a, polishop, temu", "Categoria": "🛍️ Compras e Varejo"},
            {"Palavras-chave (separadas por vírgula)": "vivo, conta vivo, enel, claro", "Categoria": "📄 Serviços e Contas"}
        ])
    
    df_inativos = ler_aba("Cartoes_Inativos")
    if df_inativos.empty or "Cartão" not in df_inativos.columns:
        df_inativos = pd.DataFrame(columns=["Cartão"])
        
    for df in [df_cartoes, df_fixas]:
        if not df.empty:
            df["Valor (R$)"] = df["Valor (R$)"].apply(parse_float)
            df["Data_Dt"] = pd.to_datetime(df["Vencimento"], format="%d/%m/%Y", errors="coerce")
            df["Mes_Ano"] = df["Data_Dt"].dt.strftime("%m/%Y")
            df["Sort_Data"] = df["Data_Dt"].dt.strftime("%Y-%m")
            df["Dia"] = df["Data_Dt"].dt.day

    df_cartoes["Tipo_Despesa"] = "💳 Cartão"
    
    def classificar_origem(cartao):
        c = str(cartao).upper()
        if "DEBITO" in c or "DÉBITO" in c or "DINHEIRO" in c or "PIX" in c: return "Fixa"
        return "💳 Cartão"
        
    if "Cartão" in df_fixas.columns: df_fixas["Tipo_Despesa"] = df_fixas["Cartão"].apply(classificar_origem)
    else: df_fixas["Tipo_Despesa"] = "Fixa"
    
    df_despesas = pd.concat([df_cartoes, df_fixas], ignore_index=True)
    if not df_despesas.empty and "Sort_Data" in df_despesas.columns:
        df_despesas = df_despesas.sort_values("Sort_Data")
        
    return df_despesas, ler_aba("Receitas_Extra"), ler_aba("Receitas_Historico"), ler_aba("Fixas"), df_rec_display, df_rec_total, df_reclass, df_inativos

@st.cache_data(ttl=600)
def carregar_despesas_brutas():
    return ler_aba("Despesas")

def processar_tabela_financeira(df_raw):
    if df_raw is None or df_raw.empty: return pd.DataFrame()
    df_s = df_raw.copy()
    col_valor = "Valor (R$)" if "Valor (R$)" in df_s.columns else "Valor" if "Valor" in df_s.columns else None
    df_s["Valor Numérico"] = df_s[col_valor].apply(parse_float) if col_valor else 0.0

    def to_sort_date(ma):
        try: return pd.to_datetime(ma, format="%m/%Y").strftime("%Y-%m")
        except: return "1900-01"
        
    col_ma = "Mês_Ano" if "Mês_Ano" in df_s.columns else "Mes_Ano" if "Mes_Ano" in df_s.columns else None
    df_s["Sort_Data"] = df_s[col_ma].apply(to_sort_date) if col_ma else "1900-01"
    return df_s.sort_values("Sort_Data")

@st.cache_data
def preparar_base(df_raw, df_reclass_raw):
    if df_raw is None or df_raw.empty: return pd.DataFrame()
    
    df_proc = df_raw.copy()
    df_proc["Tipo_Compra"] = df_proc.apply(classificar_tipo_compra, axis=1)
    df_proc["Cartão_Icon"] = df_proc["Cartão"].apply(mapear_banco)
    
    regras = []
    if not df_reclass_raw.empty and "Palavras-chave (separadas por vírgula)" in df_reclass_raw.columns:
        for _, row in df_reclass_raw.iterrows():
            palavras = str(row.get("Palavras-chave (separadas por vírgula)", "")).split(",")
            cat = str(row.get("Categoria", "")).strip()
            lista_palavras = [p.strip().lower() for p in palavras if p.strip()]
            regras.append((lista_palavras, cat))
            
    def fast_categorize(desc):
        d = str(desc).lower()
        for lista_p, categoria in regras:
            for p in lista_p:
                if p in d:
                    return categoria
        return "🏷️ Outros / Não Classificado"

    df_proc["Categoria"] = df_proc["Descrição"].apply(fast_categorize)
    df_proc["Ano"] = df_proc["Data_Dt"].dt.year.astype(str)
    df_proc["Mes_Nome"] = df_proc["Data_Dt"].dt.month.map(meses_pt)
    df_proc["Grupo"] = df_proc["Dia"].apply(lambda x: "Dia 10" if x <= 15 else "Dia 25")
    
    return df_proc

@st.cache_data
def gerar_tabela_visao_geral(df_base, df_salarios, df_extras, ano_atual, mes_atual_nome):
    periodos = df_base[["Sort_Data", "Ano", "Mes_Nome"]].drop_duplicates().sort_values("Sort_Data")
    colunas_multi = []

    for _, row in periodos.iterrows():
        ano, mes = row["Ano"], row["Mes_Nome"]
        mes_rotulo = f"📍 {mes} (Atual)" if (ano == str(ano_atual) and mes == mes_atual_nome) else mes
        colunas_multi.extend([(ano, mes_rotulo, "Dia 10"), (ano, mes_rotulo, "Dia 25"), (ano, mes_rotulo, "Total Mensal")])

    dados_despesas, dados_receita, dados_saldo = {}, {}, {}
    tem_extra_dict = {}

    for _, row in periodos.iterrows():
        ano, mes, sort_s = row["Ano"], row["Mes_Nome"], row["Sort_Data"]
        mes_rotulo = f"📍 {mes} (Atual)" if (ano == str(ano_atual) and mes == mes_atual_nome) else mes
        
        d10 = df_base[(df_base["Sort_Data"] == sort_s) & (df_base["Grupo"] == "Dia 10")]["Valor (R$)"].sum()
        d25 = df_base[(df_base["Sort_Data"] == sort_s) & (df_base["Grupo"] == "Dia 25")]["Valor (R$)"].sum()
        tot_desp = d10 + d25
        
        receita_mes_atual = 0.0
        receita_d10 = 0.0
        receita_d25 = 0.0
        
        if not df_salarios.empty and "Sort_Data" in df_salarios.columns:
            historico_valido = df_salarios[df_salarios["Sort_Data"] <= sort_s].copy()
            if not historico_valido.empty and "Origem da Renda" in historico_valido.columns:
                if "Grupo" not in historico_valido.columns: historico_valido["Grupo"] = ""
                historico_valido["Grupo"] = historico_valido["Grupo"].fillna("").astype(str)
                ultimas_rendas = historico_valido.groupby(["Origem da Renda", "Grupo"]).last().reset_index()
                receita_d10 += ultimas_rendas[ultimas_rendas["Grupo"].str.strip() == "Dia 10"]["Valor Numérico"].sum()
                receita_d25 += ultimas_rendas[ultimas_rendas["Grupo"].str.strip() == "Dia 25"]["Valor Numérico"].sum()
                receita_mes_atual += ultimas_rendas["Valor Numérico"].sum()
        
        possui_extra = False
        if not df_extras.empty and "Sort_Data" in df_extras.columns:
            extras_filtrados = df_extras[df_extras["Sort_Data"] == sort_s]
            if not extras_filtrados.empty:
                for _, row_ext in extras_filtrados.iterrows():
                    v_ext = row_ext.get("Valor Numérico", 0.0)
                    desc_ext = obter_descricao(row_ext).lower()
                    if "pix amanda" in desc_ext: continue
                        
                    receita_mes_atual += v_ext
                    possui_extra = True
                    grp = str(row_ext.get("Grupo", "")).strip()
                    if grp == "Dia 10": receita_d10 += v_ext
                    elif grp == "Dia 25": receita_d25 += v_ext

        saldo_mes = receita_mes_atual - tot_desp
        col_chave = (ano, mes_rotulo, "Total Mensal")
        tem_extra_dict[col_chave] = possui_extra

        dados_despesas[(ano, mes_rotulo, "Dia 10")], dados_despesas[(ano, mes_rotulo, "Dia 25")], dados_despesas[(ano, mes_rotulo, "Total Mensal")] = d10, d25, tot_desp
        dados_receita[(ano, mes_rotulo, "Dia 10")] = receita_d10 if receita_d10 > 0 else ""
        dados_receita[(ano, mes_rotulo, "Dia 25")] = receita_d25 if receita_d25 > 0 else ""
        dados_receita[(ano, mes_rotulo, "Total Mensal")] = receita_mes_atual
        dados_saldo[(ano, mes_rotulo, "Dia 10")], dados_saldo[(ano, mes_rotulo, "Dia 25")], dados_saldo[(ano, mes_rotulo, "Total Mensal")] = "", "", saldo_mes

    multi_idx = pd.MultiIndex.from_tuples(colunas_multi, names=["Ano", "Mês", "Vencimento"])
    tabela_geral = pd.DataFrame([dados_despesas, dados_receita, dados_saldo], index=["Despesas Totais", "Receita Total (Fixa + Extras)", "Saldo Real (Sobra / Falta)"])
    tabela_geral = tabela_geral.reindex(columns=multi_idx)

    tabela_visual = tabela_geral.copy().astype(object)
    for col in tabela_visual.columns:
        val = tabela_visual.loc["Receita Total (Fixa + Extras)", col]
        if isinstance(val, (int, float)):
            formatado = fmt_br(val)
            if col in tem_extra_dict and tem_extra_dict[col]: formatado += " 💰"
            tabela_visual.loc["Receita Total (Fixa + Extras)", col] = formatado
            
        for idx_linha in ["Despesas Totais", "Saldo Real (Sobra / Falta)"]:
            v_cel = tabela_visual.loc[idx_linha, col]
            if isinstance(v_cel, (int, float)):
                tabela_visual.loc[idx_linha, col] = fmt_br(v_cel)

    return tabela_visual

# ==========================================
# INTERFACE PRINCIPAL
# ==========================================
st.title("💸 Painel Financeiro")

df_raw_geral, df_extras_raw, df_salarios_raw, df_fixas_raw, df_rec_display, df_rec_total, df_reclass_raw, df_inativos_raw = carregar_dados()

if "editor_extras" in st.session_state:
    val_ext = st.session_state["editor_extras"]
    if isinstance(val_ext, pd.DataFrame): df_extras_raw = val_ext

if "editor_salarios" in st.session_state:
    val_sal = st.session_state["editor_salarios"]
    if isinstance(val_sal, pd.DataFrame): df_salarios_raw = val_sal

df_reclass_atuais = df_reclass_raw.copy()
if "editor_reclass" in st.session_state:
    val_rec = st.session_state["editor_reclass"]
    if isinstance(val_rec, dict):
        edits = val_rec.get("edited_rows", {})
        for idx, row in edits.items():
            for col, val in row.items():
                df_reclass_atuais.at[int(idx), col] = val
    else:
        df_reclass_atuais = val_rec

if not df_raw_geral.empty:
    
    df_base = preparar_base(df_raw_geral, df_reclass_atuais)

    ordem_meses = df_base[["Mes_Ano", "Sort_Data"]].drop_duplicates().sort_values("Sort_Data")["Mes_Ano"].tolist()
    mes_atual_str = datetime.now().strftime("%m/%Y")
    
    if mes_atual_str in ordem_meses:
        idx_mes_padrao = ordem_meses.index(mes_atual_str)
    else:
        idx_mes_padrao = len(ordem_meses) - 1 if ordem_meses else 0

    mes_padrao_inicial = [mes_atual_str] if mes_atual_str in ordem_meses else (ordem_meses[-1:] if ordem_meses else [])

    if "shared_cat" not in st.session_state: st.session_state["shared_cat"] = "Todas as Categorias"
    if "shared_tipo" not in st.session_state: st.session_state["shared_tipo"] = "Todos"
    if "shared_meses" not in st.session_state: st.session_state["shared_meses"] = []

    def update_cat_from_g4(): st.session_state["shared_cat"] = st.session_state["g4_cat"]
    def update_tipo_from_g4(): st.session_state["shared_tipo"] = st.session_state["g4_tipo"]

    col_btn, _ = st.columns([1, 4])
    with col_btn:
        if st.button("🔄 Atualizar Painel Financeiro", type="primary"):
            with st.spinner("Processando..."):
                try:
                    subprocess.run(["python", "faturas.py"], check=True)
                except:
                    pass
                carregar_dados.clear()
                carregar_despesas_brutas.clear()
                preparar_base.clear()
                gerar_tabela_visao_geral.clear()
            st.rerun()

    # ==========================================
    # MENU UNIFICADO EM ABAS
    # ==========================================
    with st.expander("⚙️ Painel de Cadastros e Ajustes (Clique para expandir)", expanded=False):
        tab_desp, tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Despesas (Lançamentos)", "🌐 Recorrentes", "🛠️ Fixas", "💼 Salários & Rendas", "💸 Extras", "🔀 Dicionário de Categorias"])
        
        with tab_desp:
            st.markdown("### 📝 Gerenciar Aba 'Despesas' (Lançamento em Massa)")
            
            df_desp_bruto = carregar_despesas_brutas()
            
            st.markdown(f"#### 📡 Radar de Faturas (Mês Atual: {mes_atual_str})")
            
            todos_cartoes = []
            if not df_desp_bruto.empty and "Cartão" in df_desp_bruto.columns:
                todos_cartoes = sorted([c for c in df_desp_bruto["Cartão"].dropna().astype(str).str.strip().unique() if c])
            
            inativos_atuais = df_inativos_raw["Cartão"].dropna().astype(str).str.strip().tolist() if not df_inativos_raw.empty else []
            cartoes_ativos = [c for c in todos_cartoes if c not in inativos_atuais]
            
            cartoes_enviados = []
            if not df_desp_bruto.empty and "Vencimento" in df_desp_bruto.columns:
                mask_mes_atual = df_desp_bruto["Vencimento"].astype(str).str.endswith(mes_atual_str)
                cartoes_enviados = df_desp_bruto[mask_mes_atual]["Cartão"].astype(str).str.strip().unique().tolist()
                
            pendentes = [c for c in cartoes_ativos if c not in cartoes_enviados]
            enviados = [c for c in cartoes_ativos if c in cartoes_enviados]
            
            with st.container(border=True):
                col_pend, col_env = st.columns(2)
                with col_pend:
                    if pendentes:
                        st.error("**⚠️ Pendentes de Lançamento neste mês:**\n\n" + "\n".join([f"- {c}" for c in pendentes]))
                    else:
                        st.success("**✅ Todas as faturas ativas deste mês foram lançadas!**")
                        
                with col_env:
                    if enviados:
                        st.success("**✅ Faturas já lançadas:**\n\n" + "\n".join([f"- {c}" for c in enviados]))
                    else:
                        st.warning("**⚠️ Nenhuma fatura foi lançada para este mês ainda.**")
            
            with st.expander("⚙️ Ocultar Cartões Cancelados / Inativos"):
                st.markdown("Selecione os cartões do seu histórico que você não usa mais. Eles pararão de ser cobrados pelo Radar.")
                valid_defaults = [c for c in inativos_atuais if c in todos_cartoes]
                selecao_inativos = st.multiselect("Cartões Ignorados:", options=todos_cartoes, default=valid_defaults)
                
                if st.button("💾 Salvar Lista de Inativos"):
                    df_novos_inativos = pd.DataFrame({"Cartão": selecao_inativos})
                    if salvar_tabela_google(df_novos_inativos, "Cartoes_Inativos"):
                        st.success("Lista de cartões inativos atualizada com sucesso!")
                        carregar_dados.clear()
                        st.rerun()

            st.divider()
            
            st.markdown("💡 *Filtre a tabela abaixo para encontrar e editar um gasto específico ou limpe os filtros e role até o final para colar dados em massa. **Esta tabela altera apenas a aba 'Despesas' no seu Google Sheets e é totalmente independente das outras funções.***")
            
            if not df_desp_bruto.empty:
                col_f_desp1, col_f_desp2, col_f_desp3 = st.columns(3)
                
                with col_f_desp1:
                    lista_cartoes_bruto = ["Todos"] + sorted(df_desp_bruto["Cartão"].dropna().astype(str).unique().tolist())
                    f_desp_cartao = st.selectbox("💳 Cartão:", lista_cartoes_bruto, key="f_desp_cartao")
                
                with col_f_desp2:
                    df_vencimentos = df_desp_bruto if f_desp_cartao == "Todos" else df_desp_bruto[df_desp_bruto["Cartão"] == f_desp_cartao]
                    vencimentos_unicos = df_vencimentos["Vencimento"].dropna().astype(str).unique().tolist()
                    lista_mes_bruto = ["Todos"] + sorted(vencimentos_unicos, key=sort_data_br)
                    f_desp_mes = st.selectbox("📅 Vencimento:", lista_mes_bruto, key="f_desp_mes")
                
                with col_f_desp3:
                    f_desp_desc = st.text_input("🔍 Descrição contém (vazio = todas):", key="f_desp_desc")
                    
                mask_desp = pd.Series(True, index=df_desp_bruto.index)
                if f_desp_cartao != "Todos":
                    mask_desp &= (df_desp_bruto["Cartão"] == f_desp_cartao)
                if f_desp_mes != "Todos":
                    mask_desp &= (df_desp_bruto["Vencimento"] == f_desp_mes)
                if f_desp_desc:
                    mask_desp &= (df_desp_bruto["Descrição"].str.contains(f_desp_desc, case=False, na=False))
                    
                df_display_desp = df_desp_bruto[mask_desp].copy()
                original_indices_desp = df_display_desp.index.tolist()
                
                total_filtrado_desp = 0.0
                if not df_display_desp.empty and "Valor (R$)" in df_display_desp.columns:
                    total_filtrado_desp = df_display_desp["Valor (R$)"].apply(parse_float).sum()
                
                st.metric(label="💰 Soma Total (Filtro Atual)", value=fmt_br(total_filtrado_desp))
                
                st.data_editor(
                    df_display_desp, 
                    num_rows="dynamic", 
                    width='stretch', 
                    hide_index=True, 
                    key="editor_despesas_brutas"
                )
                
                changes = st.session_state.get("editor_despesas_brutas", {})
                num_added = len(changes.get("added_rows", []))
                num_edited = len(changes.get("edited_rows", {}))
                num_deleted = len(changes.get("deleted_rows", []))

                confirm_del = True
                if num_deleted > 5:
                    st.warning(f"🚨 ALERTA CRÍTICO: Você marcou {num_deleted} linhas para exclusão. Isso apagará o histórico em massa!")
                    confirm_del = st.checkbox("Confirmo a exclusão destas linhas", key="confirm_del_bruto")

                if st.button("💾 Salvar Lançamentos e Edições na aba 'Despesas'", type="primary"):
                    if num_deleted > 5 and not confirm_del:
                        st.error("🛑 Ação bloqueada: Por segurança, marque a caixa de confirmação para autorizar exclusões em massa.")
                        st.stop()
                        
                    df_final_desp = df_desp_bruto.copy()
                    
                    for row_pos, col_changes in changes.get("edited_rows", {}).items():
                        row_pos_int = int(row_pos)
                        if row_pos_int < len(original_indices_desp):
                            orig_idx = original_indices_desp[row_pos_int]
                            for col_name, new_val in col_changes.items():
                                df_final_desp.at[orig_idx, col_name] = new_val
                                
                    deleted_positions = changes.get("deleted_rows", [])
                    if deleted_positions:
                        indices_to_drop = [original_indices_desp[int(pos)] for pos in deleted_positions if int(pos) < len(original_indices_desp)]
                        df_final_desp = df_final_desp.drop(index=indices_to_drop)
                        
                    added_rows = changes.get("added_rows", [])
                    if added_rows:
                        df_new_desp = pd.DataFrame(added_rows)
                        for c in df_final_desp.columns:
                            if c not in df_new_desp.columns:
                                df_new_desp[c] = ""
                        df_new_desp = df_new_desp[df_final_desp.columns]
                        df_final_desp = pd.concat([df_final_desp, df_new_desp], ignore_index=True)
                        
                    date_errors = []
                    date_regex = re.compile(r'^\d{2}/\d{2}/\d{4}$')
                    for col in ["Data", "Vencimento"]:
                        if col in df_final_desp.columns:
                            invalid_mask = df_final_desp[col].astype(str).str.strip().apply(lambda x: x != "" and not bool(date_regex.match(x)))
                            if invalid_mask.any():
                                date_errors.extend(df_final_desp[invalid_mask][col].unique().tolist())
                    
                    if date_errors:
                        st.error(f"🛑 Erro de Formatação: Encontramos datas fora do padrão DD/MM/AAAA (ex: {', '.join(date_errors[:3])}). Por favor, corrija a tabela e salve novamente.")
                        st.stop()
                        
                    existing_cards = [str(c).strip() for c in df_desp_bruto["Cartão"].dropna().unique() if str(c).strip()]
                    current_cards = [str(c).strip() for c in df_final_desp["Cartão"].dropna().unique() if str(c).strip()]
                    new_cards = [c for c in current_cards if c not in existing_cards and c != ""]
                    
                    dupe_warn = False
                    if num_added > 0:
                        df_new_only = pd.DataFrame(added_rows)
                        if all(c in df_new_only.columns for c in ["Vencimento", "Cartão", "Descrição"]):
                            for _, r in df_new_only.iterrows():
                                match = df_desp_bruto[
                                    (df_desp_bruto["Vencimento"] == r.get("Vencimento", "")) &
                                    (df_desp_bruto["Cartão"] == r.get("Cartão", "")) &
                                    (df_desp_bruto["Descrição"] == r.get("Descrição", ""))
                                ]
                                if not match.empty:
                                    dupe_warn = True
                                    break
                                    
                    with st.spinner("Sincronizando banco de dados..."):
                        if salvar_tabela_google(df_final_desp, "Despesas"):
                            st.success(f"✅ Lançamentos processados e aba 'Despesas' atualizada com sucesso! ({num_added} incluídos, {num_edited} editados, {num_deleted} excluídos).")
                            
                            if new_cards:
                                st.info(f"⚠️ Atenção: Identificamos cartões que não existiam no histórico: {', '.join(new_cards)}.")
                            if dupe_warn:
                                st.warning("⚠️ Nota: Pelo menos um dos lançamentos adicionados parece já existir na tabela (Mesma Data, Cartão e Descrição).")
                            
                            carregar_despesas_brutas.clear()
                            time.sleep(2.0)
                            st.rerun()
            else:
                st.warning("A aba 'Despesas' parece estar vazia no Google Sheets.")

        with tab1:
            st.markdown("### Descoberta Automática de Recorrentes")
            st.markdown("💡 *Todas as suas avaliações vindas do motor estão aqui. Alertas e pendências aparecem no topo para fácil auditoria.*")
            
            df_rec_display_float = df_rec_display.copy()
            
            if not df_rec_display_float.empty:
                if "Valor (R$)" in df_rec_display_float.columns:
                    df_rec_display_float["Valor (R$)"] = df_rec_display_float["Valor (R$)"].apply(parse_float)
                
                df_rec_display_float["Motivo do Alerta"] = ""
                df_rec_display_float["Mês da Avaliação"] = ""
                df_rec_display_float["Novo Valor"] = ""
                
                if not df_base.empty:
                    df_temp_base = df_base[(df_base["Origem_Aba"] == "Cartoes") & (df_base["Tipo_Compra"] == "À Vista / Mês")].copy()
                    
                    if "Data" in df_temp_base.columns:
                        df_temp_base = df_temp_base[df_temp_base["Data"].astype(str).str.strip() != ""]
                    
                    if not df_temp_base.empty:
                        df_temp_base["Cartao_Norm"] = df_temp_base["Cartão"].apply(lambda c: re.sub(r'[^a-z0-9]', '', str(c).lower()))
                        df_temp_base["Desc_Norm"] = df_temp_base["Descrição"].apply(lambda d: re.sub(r'[^a-z0-9]', '', str(d).lower()))
                        df_temp_base["Valor_Norm"] = df_temp_base["Valor (R$)"].apply(lambda v: f"{parse_float(v):.2f}")
                        
                        max_mes_por_cartao = {}
                        for c_norm in df_temp_base["Cartao_Norm"].unique():
                            df_c_vista = df_temp_base[df_temp_base["Cartao_Norm"] == c_norm]
                            if not df_c_vista.empty:
                                max_mes_por_cartao[c_norm] = df_c_vista["Sort_Data"].max()
                        
                        chaves_existentes = set()
                        if not df_rec_display_float.empty:
                            chaves_existentes = set(df_rec_display_float.apply(lambda r: normalizar_chave(r.get("Cartão", ""), r.get("Descrição", ""), r.get("Valor (R$)")), axis=1))

                        novas_recorrentes = []
                        for c_norm, max_mes in max_mes_por_cartao.items():
                            if not max_mes: continue
                            try:
                                y, m = map(int, max_mes.split('-'))
                                meses_esperados = [max_mes]
                                for _ in range(2):
                                    m -= 1
                                    if m == 0: m = 12; y -= 1
                                    meses_esperados.append(f"{y}-{m:02d}")
                                
                                df_cartao_meses = df_temp_base[(df_temp_base["Cartao_Norm"] == c_norm) & (df_temp_base["Sort_Data"].isin(meses_esperados))]
                                agrupado = df_cartao_meses.groupby(["Desc_Norm", "Valor_Norm"])["Sort_Data"].nunique().reset_index()
                                candidatas = agrupado[agrupado["Sort_Data"] == len(meses_esperados)]

                                for _, cand in candidatas.iterrows():
                                    d_norm_cand = cand["Desc_Norm"]
                                    v_norm_cand = cand["Valor_Norm"]
                                    
                                    orig_rows = df_cartao_meses[(df_cartao_meses["Desc_Norm"] == d_norm_cand) & (df_cartao_meses["Valor_Norm"] == v_norm_cand)]
                                    if orig_rows.empty: continue
                                    
                                    orig_row = orig_rows.iloc[0]
                                    k_full = normalizar_chave(orig_row["Cartão"], orig_row["Descrição"], orig_row["Valor (R$)"])

                                    if k_full not in chaves_existentes:
                                        chaves_existentes.add(k_full) 
                                        v_str = ", ".join(sorted([pd.to_datetime(m + "-01").strftime("%m/%Y") for m in meses_esperados], reverse=True))
                                        novas_recorrentes.append({
                                            "Cartão": orig_row["Cartão"], "Descrição": orig_row["Descrição"], "Valor (R$)": parse_float(orig_row["Valor (R$)"]),
                                            "Vencimentos Encontrados": v_str, "Status": "⏳ Pendente",
                                            "Motivo do Alerta": "🌟 Nova Assinatura", "Mês da Avaliação": "", "Novo Valor": "-"
                                        })
                            except: pass

                        if novas_recorrentes:
                            df_novas = pd.DataFrame(novas_recorrentes)
                            df_rec_display_float = pd.concat([df_novas, df_rec_display_float], ignore_index=True)

                        df_rec_display_float["Cartao_Norm"] = df_rec_display_float["Cartão"].apply(lambda c: re.sub(r'[^a-z0-9]', '', str(c).lower()))
                        df_rec_display_float["Desc_Norm"] = df_rec_display_float["Descrição"].apply(lambda d: re.sub(r'[^a-z0-9]', '', str(d).lower()))

                        for (c_norm, d_norm), group_regras in df_rec_display_float.groupby(["Cartao_Norm", "Desc_Norm"]):
                            mask_audit = (~group_regras["Status"].str.lower().str.contains("ignorar")) & (~group_regras["Motivo do Alerta"].str.contains("🌟", na=False))
                            regras_audit = group_regras[mask_audit]
                            
                            if regras_audit.empty: continue
                            
                            ultimo_sort_cartao = max_mes_por_cartao.get(c_norm)
                            if not ultimo_sort_cartao: continue
                            
                            try: mes_ano_str = pd.to_datetime(ultimo_sort_cartao + "-01").strftime("%m/%Y")
                            except: mes_ano_str = ultimo_sort_cartao
                            
                            df_fatura_loja = df_temp_base[
                                (df_temp_base["Cartao_Norm"] == c_norm) & 
                                (df_temp_base["Desc_Norm"] == d_norm) & 
                                (df_temp_base["Sort_Data"] == ultimo_sort_cartao)
                            ]
                            
                            valores_disponiveis = [parse_float(v) for v in df_fatura_loja["Valor (R$)"].tolist()]
                            regras_sem_match = []
                            
                            for idx, row in regras_audit.iterrows():
                                val_regra = parse_float(row["Valor (R$)"])
                                matched = False
                                for i, v_fat in enumerate(valores_disponiveis):
                                    if abs(v_fat - val_regra) <= 0.02:
                                        valores_disponiveis.pop(i)
                                        matched = True
                                        break
                                
                                if matched:
                                    df_rec_display_float.at[idx, "Motivo do Alerta"] = ""
                                    df_rec_display_float.at[idx, "Mês da Avaliação"] = ""
                                    df_rec_display_float.at[idx, "Novo Valor"] = ""
                                else:
                                    regras_sem_match.append(idx)
                                    
                            for idx in regras_sem_match:
                                df_rec_display_float.at[idx, "Mês da Avaliação"] = mes_ano_str
                                if len(valores_disponiveis) > 0:
                                    df_rec_display_float.at[idx, "Motivo do Alerta"] = "⚠️ Mudou de Valor"
                                    valores_unicos = list(set([f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") for v in valores_disponiveis]))
                                    df_rec_display_float.at[idx, "Novo Valor"] = " / ".join(valores_unicos)
                                else:
                                    df_rec_display_float.at[idx, "Motivo do Alerta"] = "🛑 Sumiu / Cancelada"
                                    df_rec_display_float.at[idx, "Novo Valor"] = "-"

                        df_rec_display_float = df_rec_display_float.drop(columns=["Cartao_Norm", "Desc_Norm"])
            
            df_display_final = df_rec_display_float.copy()
            if not df_display_final.empty:
                if (df_display_final["Motivo do Alerta"].str.strip() != "").any():
                    st.warning("⚠️ **ALERTA DE RECORRÊNCIA:** Algumas despesas aprovadas **mudaram de valor ou sumiram da última fatura do cartão**. Elas estão destacadas no topo da lista.")
                
                def calc_peso_ordem(r):
                    m = str(r.get("Motivo do Alerta", "")).strip()
                    s = str(r.get("Status", "")).lower()
                    if m != "": return 0  
                    if "pendente" in s: return 1 
                    return 2
                
                df_display_final["_peso_ordem"] = df_display_final.apply(calc_peso_ordem, axis=1)
                df_display_final = df_display_final.sort_values(by=["_peso_ordem", "Cartão", "Descrição"]).drop(columns=["_peso_ordem"])
                
                col_order = ["Cartão", "Descrição", "Valor (R$)", "Novo Valor", "Mês da Avaliação", "Motivo do Alerta", "Status", "Vencimentos Encontrados"]
                for col in col_order:
                    if col not in df_display_final.columns:
                        df_display_final[col] = ""
                disp_cols = [c for c in col_order if c in df_display_final.columns]
                df_display_final = df_display_final[disp_cols]

            df_display_final = df_display_final.reset_index(drop=True)

            df_editado_rec = st.data_editor(
                df_display_final, num_rows="dynamic", width='stretch', hide_index=True, key="editor_recorrentes",
                column_config={
                    "Status": st.column_config.SelectboxColumn("Decisão", options=["⏳ Pendente", "✅ Aprovado", "❌ Ignorar"], required=True), 
                    "Motivo do Alerta": st.column_config.TextColumn("Motivo do Alerta", disabled=True),
                    "Mês da Avaliação": st.column_config.TextColumn("Mês da Avaliação", disabled=True),
                    "Novo Valor": st.column_config.TextColumn("Novo Valor", disabled=True),
                    "Vencimentos Encontrados": st.column_config.TextColumn("Vencimentos Anteriores", disabled=True),
                    "Descrição": st.column_config.TextColumn("Descrição", width="large"),
                    "Valor (R$)": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f")
                }
            )
            
            if st.button("💾 Salvar TODAS as Decisões"):
                df_rec_total["Chave_Temp"] = df_rec_total.apply(lambda r: normalizar_chave(r.get("Cartão", ""), r.get("Descrição", ""), r.get("Valor (R$)", "")), axis=1)
                
                if not df_editado_rec.empty:
                    for _, row_ed in df_editado_rec.iterrows():
                        chave_ed = normalizar_chave(row_ed.get("Cartão", ""), row_ed.get("Descrição", ""), row_ed.get("Valor (R$)", ""))
                        status_ed = row_ed.get("Status")
                        motivo_ed = row_ed.get("Motivo do Alerta", "")
                        
                        if status_ed in ["✅ Aprovado", "❌ Ignorar"]:
                            motivo_ed = ""
                            
                        mask = df_rec_total["Chave_Temp"] == chave_ed
                        
                        if not df_rec_total[mask].empty:
                            df_rec_total.loc[mask, "Status"] = status_ed
                            df_rec_total.loc[mask, "Motivo do Alerta"] = motivo_ed
                        else:
                            novo_dict = row_ed.to_dict()
                            novo_dict["Valor (R$)"] = str(novo_dict.get("Valor (R$)", ""))
                            novo_dict["Status"] = status_ed
                            novo_dict["Motivo do Alerta"] = motivo_ed
                            
                            valid_cols = ["Cartão", "Descrição", "Valor (R$)", "Vencimentos Encontrados", "Status", "Motivo do Alerta"]
                            novo_dict = {k: v for k, v in novo_dict.items() if k in valid_cols}
                            
                            novo_registro = pd.DataFrame([novo_dict])
                            df_rec_total = pd.concat([df_rec_total, novo_registro], ignore_index=True)
                
                if "Chave_Temp" in df_rec_total.columns:
                    df_rec_total = df_rec_total.drop(columns=["Chave_Temp"])
                    
                if salvar_tabela_google(df_rec_total, "Avaliar Recorrentes"):
                    st.success("Salvo com sucesso!")
                    carregar_dados.clear()
                    st.rerun()

        with tab2:
            st.markdown("### Gerenciar Despesas Fixas & Provisões")
            df_editado_fixas = st.data_editor(
                df_fixas_raw, num_rows="dynamic", width='stretch', hide_index=True, key="editor_fixas",
                column_config={"Descrição": st.column_config.TextColumn("Descrição", width="large")}
            )
            if st.button("💾 Salvar Alterações nas Fixas/Provisões"):
                if salvar_tabela_google(df_editado_fixas, "Fixas"): st.success("Salvo!"); carregar_dados.clear(); st.rerun()

        with tab3:
            st.markdown("### Histórico de Rendas Fixas (Salários, Aluguéis, etc)")
            df_editado_salarios = st.data_editor(
                df_salarios_raw, num_rows="dynamic", width='stretch', hide_index=True, key="editor_salarios",
                column_config={"Grupo": st.column_config.TextColumn("Grupo (Ex: Dia 10 ou Dia 25)")}
            )
            if st.button("💾 Salvar Histórico de Rendas"):
                if salvar_tabela_google(df_editado_salarios, "Receitas_Historico"): st.success("Atualizado!"); carregar_dados.clear(); st.rerun()

        with tab4:
            st.markdown("### Histórico de Dinheiro Extra (Bônus, Vendas, etc)")
            st.markdown("💡 *Para registrar transferência da Amanda, escreva **'Pix Amanda'** na Descrição.*")
            df_editado_extras = st.data_editor(
                df_extras_raw, num_rows="dynamic", width='stretch', hide_index=True, key="editor_extras",
                column_config={"Grupo": st.column_config.TextColumn("Grupo (Ex: Dia 10 ou Dia 25)"), "Sort_Data": st.column_config.TextColumn("Filtro (Automático)", disabled=True), "Descrição": st.column_config.TextColumn("Descrição", width="large")}
            )
            if st.button("💾 Salvar Histórico de Extras"):
                col_ma = "Mes_Ano" if "Mes_Ano" in df_extras_raw.columns else "Mês_Ano" if "Mês_Ano" in df_extras_raw.columns else None
                if col_ma: df_extras_raw["Sort_Data"] = df_extras_raw[col_ma].apply(lambda ma: pd.to_datetime(ma, format="%m/%Y").strftime("%Y-%m") if pd.notna(ma) else "")
                if salvar_tabela_google(df_extras_raw, "Receitas_Extra"): st.success("Salvo!"); carregar_dados.clear(); st.rerun()
                
        with tab5:
            st.markdown("### 🔀 Controle de Categorias (Crie, Edite ou Apague)")
            st.markdown("Aqui estão **todas** as regras do painel. Quer que uma palavra vá para outra categoria? Basta adicionar a palavra (separada por vírgula) na linha da categoria desejada. Quer criar uma nova? Preencha uma nova linha em branco no final da tabela!")
            
            df_editado_reclass = st.data_editor(
                df_reclass_raw, num_rows="dynamic", width='stretch', hide_index=True, key="editor_reclass",
                column_config={
                    "Palavras-chave (separadas por vírgula)": st.column_config.TextColumn("Palavras-chave (separadas por vírgula)", width="large", required=True),
                    "Categoria": st.column_config.TextColumn("Nome da Categoria", width="medium", required=True)
                }
            )
            if st.button("💾 Salvar Dicionário de Categorias"):
                if salvar_tabela_google(df_editado_reclass, "Reclassificacao"): 
                    st.success("Regras salvas e aplicadas a todo o histórico!")
                    carregar_dados.clear()
                    preparar_base.clear()
                    st.rerun()

    df_salarios = processar_tabela_financeira(df_editado_salarios)
    df_extras = processar_tabela_financeira(df_editado_extras)
    
    ano_atual = datetime.now().year
    mes_atual_nome = meses_pt[datetime.now().month]

    # ==========================================
    # 1. VISÃO GERAL (Cached Table)
    # ==========================================
    st.subheader("🗓️ 1. Visão Geral: Vencimentos, Receitas e Saldo")
    
    tabela_visual_geral = gerar_tabela_visao_geral(df_base, df_salarios, df_extras, ano_atual, mes_atual_nome)
    
    def colorir_negativos_texto(val):
        if isinstance(val, str) and "R$ -" in val: return 'color: #dc3545; font-weight: bold;'
        return ''

    if not tabela_visual_geral.empty:
        cols_totais = [col for col in tabela_visual_geral.columns if col[2] == 'Total Mensal']
        st.dataframe(
            tabela_visual_geral.style.apply(destacar_mes, axis=0).set_properties(subset=cols_totais, **{'background-color': '#e8f4f8', 'font-weight': 'bold'}).map(colorir_negativos_texto), 
            width='stretch'
        )

    st.divider()

    # ==========================================
    # 2. CONFERÊNCIA
    # ==========================================
    st.subheader("🔎 2. Conferência: Totais Mensais por Cartão")
    st.markdown("💡 *Dica: Os valores ficam na cor **Vermelha** se a fatura aumentou em relação ao mês anterior, e **Verde** se diminuiu.*")
    
    filtro_tipo = st.radio("Selecione o tipo de conta para as tabelas de conferência e detalhamento:", ["Todos", "💳 Cartão", "Fixa"], horizontal=True)
    
    df_filtrado = df_base.copy()
    if filtro_tipo != "Todos": df_filtrado = df_filtrado[df_filtrado["Tipo_Despesa"] == filtro_tipo]

    tabela_conf = df_filtrado.pivot_table(index="Cartão_Icon", columns="Mes_Ano", values="Valor (R$)", aggfunc="sum").fillna(0)
    tabela_conf = tabela_conf.reindex(columns=[m for m in ordem_meses if m in tabela_conf.columns]).fillna(0)
    
    novas_colunas = []
    for col in tabela_conf.columns:
        if col == mes_atual_str: novas_colunas.append(f"📍 {col} (Atual)")
        else: novas_colunas.append(col)
    tabela_conf.columns = novas_colunas
    
    st.dataframe(
        tabela_conf.style.apply(estilizar_tendencia, axis=None).apply(destacar_mes, axis=0).format(fmt_br), 
        width='stretch'
    )
    st.divider()

    # ==========================================
    # 3. ACERTO MENSAL (DETALHADO E ESTILIZADO)
    # ==========================================
    st.subheader("🤝 3. Acerto Mensal (Caixas Individuais)")
    col_mes_acerto, _ = st.columns([1, 3])
    with col_mes_acerto:
        mes_acerto = st.selectbox("Selecione o mês para o acerto:", ordem_meses, index=idx_mes_padrao, key="select_acerto")

    df_acerto = df_base[df_base["Mes_Ano"] == mes_acerto]
    sort_s_acerto = df_acerto["Sort_Data"].iloc[0] if not df_acerto.empty else "1900-01"

    desp_amanda = df_acerto[df_acerto["Cartão"].str.contains("Amanda", case=False, na=False)]["Valor (R$)"].sum()
    desp_paulo = df_acerto["Valor (R$)"].sum() - desp_amanda

    renda_amanda, renda_paulo = 0.0, 0.0
    if not df_salarios.empty and "Sort_Data" in df_salarios.columns:
        hist_s = df_salarios[df_salarios["Sort_Data"] <= sort_s_acerto].copy()
        if not hist_s.empty and "Origem da Renda" in hist_s.columns:
            if "Grupo" not in hist_s.columns: hist_s["Grupo"] = ""
            ultimas_r = hist_s.groupby(["Origem da Renda", "Grupo"]).last().reset_index()
            for _, r in ultimas_r.iterrows():
                if "amanda" in str(r["Origem da Renda"]).lower(): renda_amanda += r["Valor Numérico"]
                else: renda_paulo += r["Valor Numérico"]

    pix_amanda, extra_paulo, extra_amanda = 0.0, 0.0, 0.0
    if not df_extras.empty and "Sort_Data" in df_extras.columns:
        ext_acerto = df_extras[df_extras["Sort_Data"] == sort_s_acerto]
        if not ext_acerto.empty:
            for _, r_ext in ext_acerto.iterrows():
                desc_ext = obter_descricao(r_ext).lower()
                val_ext = r_ext.get("Valor Numérico", 0.0)
                if "pix amanda" in desc_ext: pix_amanda += val_ext
                elif "amanda" in desc_ext and "paulo" not in desc_ext: extra_amanda += val_ext
                else: extra_paulo += val_ext

    saldo_parcial_paulo = (renda_paulo + pix_amanda) - desp_paulo
    saldo_final_paulo = saldo_parcial_paulo + extra_paulo

    saldo_parcial_amanda = renda_amanda - pix_amanda - desp_amanda
    saldo_final_amanda = saldo_parcial_amanda + extra_amanda

    colP, colA = st.columns(2)
    with colP:
        with st.container(border=True):
            st.markdown("### 👦 Caixa Paulo")
            st.markdown(f"""
                <div style="line-height: 1.8;">
                    <strong>(+) Rendas:</strong> {fmt_br(renda_paulo)}<br>
                    <strong>(+) Pix Recebido da Amanda:</strong> {fmt_br(pix_amanda)}<br>
                    <strong>(-) Contas Pagas:</strong> {fmt_br(desp_paulo)}
                </div>
            """, unsafe_allow_html=True)
            st.divider()
            st.markdown(f"""
                <div style="line-height: 1.8;">
                    <strong>= Saldo Parcial (Faltou/Sobrou):</strong> {fmt_br(saldo_parcial_paulo)}<br>
                    <strong>(+) Dinheiro Extra (PLR, etc):</strong> {fmt_br(extra_paulo)}
                </div>
            """, unsafe_allow_html=True)
            st.divider()
            cor_saldo_p = "🔴" if saldo_final_paulo < 0 else "🟢"
            st.markdown(f"#### Sobrou no Final: {fmt_br(saldo_final_paulo)} {cor_saldo_p}")
            
    with colA:
        with st.container(border=True):
            st.markdown("### 👩 Caixa Amanda")
            st.markdown(f"""
                <div style="line-height: 1.8;">
                    <strong>(+) Rendas:</strong> {fmt_br(renda_amanda)}<br>
                    <strong>(-) Pix Feito para o Paulo:</strong> {fmt_br(pix_amanda)}<br>
                    <strong>(-) Contas Pagas:</strong> {fmt_br(desp_amanda)}
                </div>
            """, unsafe_allow_html=True)
            st.divider()
            st.markdown(f"""
                <div style="line-height: 1.8;">
                    <strong>= Saldo Parcial (Faltou/Sobrou):</strong> {fmt_br(saldo_parcial_amanda)}<br>
                    <strong>(+) Dinheiro Extra (Bônus, etc):</strong> {fmt_br(extra_amanda)}
                </div>
            """, unsafe_allow_html=True)
            st.divider()
            cor_saldo_a = "🔴" if saldo_final_amanda < 0 else "🟢"
            st.markdown(f"#### Sobrou no Final: {fmt_br(saldo_final_amanda)} {cor_saldo_a}")

    st.divider()

    # ==========================================
    # 4 & 5. EVOLUÇÃO, RAIO-X E DETALHAMENTO (INTEGRADOS)
    # ==========================================
    col_tit_s4, col_rst_s4 = st.columns([5, 1])
    with col_tit_s4:
        st.subheader("📊 4. Evolução e Raio-X de Gastos")
    with col_rst_s4:
        st.write("")
        if st.button("🔄 Restaurar Filtros", key="btn_reset_geral", help="Restaurar todos os filtros da análise"):
            for k in ["g4_cat", "shared_meses", "g4_tipo", "filtro_cartao", "shared_cat", "shared_tipo"]:
                if k in st.session_state: del st.session_state[k]
            st.rerun()
    
    col_f1, col_f2, col_f3 = st.columns([1.5, 1.5, 1.2])
    
    todas_categorias_disponiveis = sorted(df_filtrado["Categoria"].astype(str).unique().tolist())
    opcoes_cat_seletor = ["Todas as Categorias"] + todas_categorias_disponiveis
    tipos_op = ["Todos", "À Vista / Mês", "Parcelado"]

    with col_f1:
        idx_g4 = opcoes_cat_seletor.index(st.session_state["shared_cat"]) if st.session_state["shared_cat"] in opcoes_cat_seletor else 0
        st.selectbox("Categoria:", options=opcoes_cat_seletor, index=idx_g4, key="g4_cat", on_change=update_cat_from_g4)

    with col_f2:
        meses_selecionados = st.multiselect("Mês(es) (Vazio = Todos):", options=ordem_meses, key="shared_meses")

    with col_f3:
        idx_t4 = tipos_op.index(st.session_state["shared_tipo"]) if st.session_state["shared_tipo"] in tipos_op else 0
        st.selectbox("Tipo:", options=tipos_op, index=idx_t4, key="g4_tipo", on_change=update_tipo_from_g4)

    if not meses_selecionados:
        df_meses_filtrados = df_filtrado.copy()
        label_meses_str = "Todos os Meses"
    else:
        df_meses_filtrados = df_filtrado[df_filtrado["Mes_Ano"].isin(meses_selecionados)]
        label_meses_str = ", ".join(meses_selecionados)

    df_analise = df_meses_filtrados.copy()
    if st.session_state["shared_cat"] != "Todas as Categorias":
        df_analise = df_analise[df_analise["Categoria"] == st.session_state["shared_cat"]]
    if st.session_state["shared_tipo"] != "Todos":
        df_analise = df_analise[df_analise["Tipo_Compra"] == st.session_state["shared_tipo"]]

    df_chart = df_filtrado.copy()
    if st.session_state["shared_cat"] != "Todas as Categorias":
        df_chart = df_chart[df_chart["Categoria"] == st.session_state["shared_cat"]]
    if st.session_state["shared_tipo"] != "Todos":
        df_chart = df_chart[df_chart["Tipo_Compra"] == st.session_state["shared_tipo"]]

    df_evolucao_mes = df_chart.groupby(["Sort_Data", "Mes_Ano"])["Valor (R$)"].sum().reset_index()
    df_evolucao_mes = df_evolucao_mes.sort_values("Sort_Data")

    col_tabela, col_grafico = st.columns([1, 1.8])
    
    with col_tabela:
        st.markdown(f"**🏆 Ranking Consolidado ({label_meses_str})**")
        
        df_ranking = df_meses_filtrados.copy()
            
        if not df_ranking.empty:
            if st.session_state["shared_tipo"] != "Todos":
                df_ranking = df_ranking[df_ranking["Tipo_Compra"] == st.session_state["shared_tipo"]]

            df_ranking_pivot = df_ranking.pivot_table(
                index="Categoria", 
                columns="Tipo_Compra", 
                values="Valor (R$)", 
                aggfunc="sum"
            ).fillna(0.0)
            
            for col_t in ["À Vista / Mês", "Parcelado"]:
                if col_t not in df_ranking_pivot.columns:
                    df_ranking_pivot[col_t] = 0.0
                    
            df_ranking_pivot["Total"] = df_ranking_pivot["À Vista / Mês"] + df_ranking_pivot["Parcelado"]
            df_ranking_pivot = df_ranking_pivot.sort_values("Total", ascending=False).reset_index()
            
            def destacar_categoria_ranking(row):
                if st.session_state["shared_cat"] != "Todas as Categorias" and row["Categoria"] == st.session_state["shared_cat"]:
                    return ['background-color: #fff3cd; font-weight: bold; color: #856404;'] * len(row)
                return [''] * len(row)

            altura_tabela = len(df_ranking_pivot) * 36 + 43

            st.dataframe(
                df_ranking_pivot[["Categoria", "À Vista / Mês", "Parcelado", "Total"]].style.apply(destacar_categoria_ranking, axis=1).format({
                    "À Vista / Mês": formatar_tabela_br, 
                    "Parcelado": formatar_tabela_br, 
                    "Total": formatar_tabela_br
                }), 
                hide_index=True, 
                width='stretch',
                height=altura_tabela
            )
        else:
            st.info(f"Nenhum dado para o período selecionado.")
            
    with col_grafico:
        titulo_grafico = f"📈 Evolução ({st.session_state['shared_cat']})"
        st.markdown(f"**{titulo_grafico}** - (🔴 Maior que o mês anterior)")
        
        if not df_evolucao_mes.empty:
            meses_x = df_evolucao_mes["Mes_Ano"].tolist()
            valores_y = df_evolucao_mes["Valor (R$)"].tolist()
            
            cores_barras = []
            for i in range(len(valores_y)):
                if i == 0:
                    cores_barras.append('#1f77b4')
                else:
                    if valores_y[i] > valores_y[i-1]:
                        cores_barras.append('#dc3545')
                    else:
                        cores_barras.append('#1f77b4')
            
            rotulos = [fmt_br(v) for v in valores_y]
            
            fig = go.Figure(data=[
                go.Bar(
                    x=meses_x, 
                    y=valores_y, 
                    text=rotulos,
                    textposition='outside',
                    textfont=dict(size=12, color='black'),
                    cliponaxis=False,
                    marker_color=cores_barras,
                    hovertemplate='<b>Mês</b>: %{x}<br><b>Valor</b>: R$ %{y:,.2f}<extra></extra>'
                )
            ])
            
            max_y = max(valores_y) if valores_y else 0
            
            fig.update_layout(
                height=450, 
                margin=dict(l=20, r=20, t=20, b=20),
                yaxis=dict(showgrid=True, zeroline=True, range=[0, max_y * 1.25])
            )
            
            if len(meses_x) > 12:
                try:
                    idx_atual = meses_x.index(mes_atual_str)
                except ValueError:
                    idx_atual = len(meses_x) - 1
                
                end_idx = idx_atual + 3  
                start_idx = end_idx - 12
                
                if start_idx < 0:
                    start_idx = 0
                    end_idx = min(12, len(meses_x))
                if end_idx > len(meses_x):
                    end_idx = len(meses_x)
                    start_idx = max(0, len(meses_x) - 12)

                fig.update_xaxes(
                    tickangle=-45, 
                    type='category', 
                    range=[start_idx - 0.5, end_idx - 0.5]
                )
                fig.update_layout(dragmode='pan')
                st.plotly_chart(fig, width='stretch')
                st.caption("👈 *A visão foca no período atual. **Clique no gráfico e arraste para os lados** para ver o passado e as projeções futuras completas.*")
            else:
                fig.update_xaxes(tickangle=-45, type='category')
                st.plotly_chart(fig, width='stretch')
        else:
            st.info(f"Nenhum dado registrado para a categoria selecionada.")

    # ==========================================
    # DETALHAMENTO POR DESPESA (Lupa Final)
    # ==========================================
    st.divider()
    st.subheader("📑 Detalhamento das Despesas Filtradas")
    
    col_cartao, col_tot, _ = st.columns([1.5, 1.5, 2])

    with col_cartao:
        lista_cartoes = ["Todos"] + sorted(df_analise["Cartão"].unique().tolist())
        cartao_sel_s5 = st.selectbox("Filtrar por Cartão / Débito:", options=lista_cartoes, key="filtro_cartao")
    
    df_detalhe = df_analise.copy()
    if cartao_sel_s5 != "Todos": 
        df_detalhe = df_detalhe[df_detalhe["Cartão"] == cartao_sel_s5]
    
    total_detalhe = df_detalhe["Valor (R$)"].sum() if not df_detalhe.empty else 0.0
    with col_tot:
        st.metric(label="Total Detalhado", value=f"R$ {total_detalhe:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    colunas_exibicao = ["Data", "Vencimento", "Mes_Ano", "Cartão_Icon", "Categoria", "Tipo_Compra", "Descrição", "Parcela", "Valor (R$)"]
    cols_finais = [c for c in colunas_exibicao if c in df_detalhe.columns]
    
    def destacar_categoria_detalhe(row):
        if st.session_state["shared_cat"] != "Todas as Categorias" and row["Categoria"] == st.session_state["shared_cat"]:
            return ['background-color: #fff3cd; font-weight: 500;'] * len(row)
        return [''] * len(row)

    if not df_detalhe.empty:
        st.dataframe(
            df_detalhe.sort_values(["Sort_Data", "Vencimento"])[cols_finais].style.apply(destacar_categoria_detalhe, axis=1).format({"Valor (R$)": formatar_tabela_br}), 
            width='stretch', 
            hide_index=True
        )
    else:
        st.info("Nenhum dado encontrado para os filtros selecionados.")

else:
    st.info("Nenhum dado encontrado.")
