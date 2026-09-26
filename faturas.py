import pandas as pd
import gspread
import re
import streamlit as st
from google.oauth2.service_account import Credentials

# --- CONEXÃO ÚNICA ---
cred = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=["https://www.googleapis.com/auth/spreadsheets"])
ws = gspread.authorize(cred).open_by_url("https://docs.google.com/spreadsheets/d/1DosfnqIt8ioBxfXMrpEBT8Md7apgsdhvWg0YswUA-IA/edit")

def tratar_valor(valor):
    if isinstance(valor, (int, float)): return float(valor)
    v = str(valor).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".").strip()
    try: return float(v)
    except: return 0.0

def limpar_desc(texto):
    t = str(texto).upper()
    t = re.sub(r'\d+\s*(/|DE)\s*\d+', '', t)
    t = re.sub(r'[^A-Z0-9]', '', t)
    return t[:15]

def salvar_aba(df, nome_aba):
    aba = ws.worksheet(nome_aba)
    aba.clear()
    if not df.empty:
        df_aux = df.copy().fillna("")
        if "Valor (R$)" in df_aux.columns:
            df_aux["Valor (R$)"] = df_aux["Valor (R$)"].apply(tratar_valor)
            df_aux["Valor (R$)"] = df_aux["Valor (R$)"].apply(lambda x: f"{x:.2f}".replace(".", ","))
        aba.update([df_aux.columns.tolist()] + df_aux.values.tolist(), value_input_option="USER_ENTERED")

# ==========================================
# 1. PROCESSAMENTO DE DESPESAS E RECORRENTES
# ==========================================
dados_despesas = ws.worksheet("Despesas").get_all_values()
df_full = pd.DataFrame(dados_despesas[1:], columns=dados_despesas[0]) if len(dados_despesas) > 1 else pd.DataFrame()

if df_full.empty:
    print("Nenhuma despesa encontrada na aba Despesas.")
    exit()

df_full["Valor (R$)"] = df_full["Valor (R$)"].apply(tratar_valor)
df_full["Vencimento_dt"] = pd.to_datetime(df_full["Vencimento"], format="%d/%m/%Y", errors="coerce")
df_valid = df_full.dropna(subset=["Vencimento_dt"]).copy()

try: 
    aba_rec = ws.worksheet("Avaliar Recorrentes")
except: 
    aba_rec = ws.add_worksheet(title="Avaliar Recorrentes", rows="100", cols="6")
    
dados_rec = aba_rec.get_all_values()
colunas_rec = ["Cartão", "Descrição", "Valor (R$)", "Vencimentos Encontrados", "Status", "Motivo do Alerta"]

if len(dados_rec) > 1:
    df_rec = pd.DataFrame(dados_rec[1:], columns=dados_rec[0])
else:
    df_rec = pd.DataFrame(columns=colunas_rec)
    
for col in colunas_rec:
    if col not in df_rec.columns:
        df_rec[col] = ""

# Base estrita: Só compras à vista de cartão para descobrir assinaturas
df_sem_parcela = df_valid[~df_valid["Parcela"].str.contains(r'\d+/\d+', na=False)].copy()
df_sem_parcela["Desc_Limpa"] = df_sem_parcela["Descrição"].apply(limpar_desc)
df_sem_parcela["Valor_Str"] = df_sem_parcela["Valor (R$)"].apply(lambda x: f"{x:.2f}")

df_sem_parcela["Chave_Base"] = df_sem_parcela["Cartão"] + "_" + df_sem_parcela["Desc_Limpa"]
df_sem_parcela["Chave_Completa"] = df_sem_parcela["Chave_Base"] + "_" + df_sem_parcela["Valor_Str"]
df_sem_parcela["Mes_Ano"] = df_sem_parcela["Vencimento_dt"].dt.strftime("%Y-%m")

# =========================================================================================
# AUDITORIA NA RAIZ (Confere quem está Aprovado contra a última fatura)
# =========================================================================================
if not df_rec.empty:
    df_rec["Valor_Str"] = df_rec["Valor (R$)"].apply(lambda x: f"{tratar_valor(x):.2f}")
    df_rec["Desc_Limpa"] = df_rec["Descrição"].apply(limpar_desc)
    df_rec["Chave_Base"] = df_rec["Cartão"] + "_" + df_rec["Desc_Limpa"]
    df_rec["Chave_Completa"] = df_rec["Chave_Base"] + "_" + df_rec["Valor_Str"]
    
    for idx, row in df_rec.iterrows():
        if row["Status"] == "✅ Aprovado":
            c_norm = row["Cartão"]
            chave_comp = row["Chave_Completa"]
            chave_base = row["Chave_Base"]
            
            df_cartao = df_sem_parcela[df_sem_parcela["Cartão"] == c_norm]
            if df_cartao.empty: continue
            
            ultimo_mes = df_cartao["Mes_Ano"].max()
            df_ultimo_mes = df_cartao[df_cartao["Mes_Ano"] == ultimo_mes]
            
            if chave_comp not in df_ultimo_mes["Chave_Completa"].tolist():
                if chave_base in df_ultimo_mes["Chave_Base"].tolist():
                    df_rec.at[idx, "Status"] = "⏳ Pendente"
                    df_rec.at[idx, "Motivo do Alerta"] = f"⚠️ Mudou de Valor na fatura {ultimo_mes}"
                else:
                    df_rec.at[idx, "Status"] = "⏳ Pendente"
                    df_rec.at[idx, "Motivo do Alerta"] = f"🛑 Sumiu / Cancelada na fatura {ultimo_mes}"
            else:
                df_rec.at[idx, "Motivo do Alerta"] = ""
                
    existentes_completas = df_rec["Chave_Completa"].tolist()
    aprovadas_base = df_rec[df_rec["Status"] == "✅ Aprovado"]["Chave_Base"].tolist()
else:
    existentes_completas = []
    aprovadas_base = []

# =========================================================================================
# [NOVA MELHORIA - REGRA DE 2 MESES (Atual e Anterior)] E O PULO DO GATO
# =========================================================================================
chaves_recorrentes = set()
novos = []

for nome_cartao, df_c in df_sem_parcela.groupby("Cartão"):
    meses_unicos = sorted(df_c["Mes_Ano"].unique().tolist(), reverse=True)
    if not meses_unicos: continue
    
    # 1. Regra dos 2 Meses Consecutivos (Mês Atual + Mês Anterior)
    if len(meses_unicos) >= 2:
        chaves_m0 = set(df_c[df_c["Mes_Ano"] == meses_unicos[0]]["Chave_Completa"])
        chaves_m1 = set(df_c[df_c["Mes_Ano"] == meses_unicos[1]]["Chave_Completa"])
        chaves_recorrentes.update(chaves_m0.intersection(chaves_m1))
        
    # 2. Pulo do Gato (Sugestão Imediata de Novo Valor)
    df_ultimo_mes = df_c[df_c["Mes_Ano"] == meses_unicos[0]]
    for _, row_ult in df_ultimo_mes.iterrows():
        if row_ult["Chave_Base"] in aprovadas_base:
            if row_ult["Chave_Completa"] not in existentes_completas:
                existentes_completas.append(row_ult["Chave_Completa"])
                novos.append({
                    "Cartão": row_ult["Cartão"], 
                    "Descrição": row_ult["Descrição"], 
                    "Valor (R$)": row_ult["Valor (R$)"], 
                    "Vencimentos Encontrados": row_ult["Vencimento"], 
                    "Status": "⏳ Pendente",
                    "Motivo do Alerta": "⚡ Novo Valor Detectado"
                })

# Junta as novas encontradas pelos 2 meses
for chave in chaves_recorrentes:
    if chave not in existentes_completas:
        existentes_completas.append(chave)
        ocorrencias = df_sem_parcela[df_sem_parcela["Chave_Completa"] == chave].sort_values("Vencimento_dt", ascending=False)
        novos.append({
            "Cartão": ocorrencias.iloc[0]["Cartão"], 
            "Descrição": ocorrencias.iloc[0]["Descrição"], 
            "Valor (R$)": ocorrencias.iloc[0]["Valor (R$)"], 
            "Vencimentos Encontrados": ", ".join(ocorrencias["Vencimento"].unique().tolist()[:5]), 
            "Status": "⏳ Pendente",
            "Motivo do Alerta": ""
        })

# Salva a tabela limpa
if novos:
    df_rec = pd.concat([df_rec.drop(columns=["Chave_Base", "Chave_Completa", "Valor_Str", "Desc_Limpa"], errors="ignore"), pd.DataFrame(novos)], ignore_index=True)

if not df_rec.empty:
    colunas_finais = ["Cartão", "Descrição", "Valor (R$)", "Vencimentos Encontrados", "Status", "Motivo do Alerta"]
    for col in colunas_finais:
        if col not in df_rec.columns: df_rec[col] = ""
    df_save_rec = df_rec[colunas_finais].copy()
    
    df_save_rec["_peso"] = df_save_rec["Status"].apply(lambda s: 0 if "Pendente" in str(s) else 1)
    df_save_rec = df_save_rec.sort_values(by=["_peso", "Cartão", "Descrição"]).drop(columns=["_peso"])
    salvar_aba(df_save_rec, "Avaliar Recorrentes")

# =========================================================================================
# PROJEÇÃO DAS DESPESAS APROVADAS PARA O DASHBOARD LER (12 Meses)
# =========================================================================================
lista_proj = []
df_rec_aprovados = df_rec[df_rec["Status"] == "✅ Aprovado"].copy() if not df_rec.empty else pd.DataFrame()

for nome_cartao, df_cartao in df_valid.groupby("Cartão"):
    ultima_data = df_cartao["Vencimento_dt"].max()
    df_recente = df_cartao[df_cartao["Vencimento_dt"] == ultima_data]
    mask_parc = df_recente["Parcela"].str.contains(r'\d+/\d+', na=False)
    
    for _, linha in df_recente[mask_parc].iterrows():
        p = str(linha["Parcela"]).split('/')
        if int(p[0]) < int(p[1]):
            for i in range(1, (int(p[1]) - int(p[0]) + 1)):
                nova = linha.copy()
                nova["Vencimento"] = (linha["Vencimento_dt"] + pd.DateOffset(months=i)).strftime("%d/%m/%Y")
                nova["Parcela"] = f"{int(p[0])+i:02d}/{p[1]}"
                lista_proj.append(nova.to_dict())
                
    if not df_rec_aprovados.empty:
        aprovados_cartao = df_rec_aprovados[df_rec_aprovados["Cartão"] == nome_cartao]
        for _, linha_rec in aprovados_cartao.iterrows():
            for i in range(1, 13):
                lista_proj.append({
                    "Cartão": nome_cartao, 
                    "Vencimento": (ultima_data + pd.DateOffset(months=i)).strftime("%d/%m/%Y"), 
                    "Descrição": linha_rec["Descrição"], 
                    "Parcela": "-", 
                    "Valor (R$)": tratar_valor(linha_rec["Valor (R$)"])
                })

df_final = pd.concat([df_full, pd.DataFrame(lista_proj)], ignore_index=True).drop(columns=["Vencimento_dt"], errors="ignore")
salvar_aba(df_final, "Despesas Cartões")

# ==========================================
# 2. PROCESSAMENTO DAS FIXAS (Débito e Parceladas em Cartão)
# ==========================================
try:
    aba_fixas = ws.worksheet("Fixas")
    dados_fixas = aba_fixas.get_all_values()
    df_fixas_full = pd.DataFrame(dados_fixas[1:], columns=dados_fixas[0]) if len(dados_fixas) > 1 else pd.DataFrame(columns=["Cartão", "Vencimento", "Descrição", "Parcela", "Valor (R$)"])
    
    if not df_fixas_full.empty and "Vencimento" in df_fixas_full.columns:
        df_fixas_full["Valor (R$)"] = df_fixas_full["Valor (R$)"].apply(tratar_valor)
        df_fixas_full["Vencimento_dt"] = pd.to_datetime(df_fixas_full["Vencimento"], format="%d/%m/%Y", errors="coerce")
        
        lista_proj_fixas = []
        df_fixas_valid = df_fixas_full.dropna(subset=["Vencimento_dt"]).copy()
        
        if not df_fixas_valid.empty:
            for _, linha in df_fixas_valid.iterrows():
                parcela_str = str(linha.get("Parcela", ""))
                if pd.notna(linha.get("Parcela")) and re.search(r'\d+/\d+', parcela_str):
                    try:
                        p = parcela_str.split('/')
                        atual, total = int(p[0]), int(p[1])
                        for i in range(1, (total - atual + 1)):
                            nova = linha.copy()
                            nova["Vencimento_dt"] = nova["Vencimento_dt"] + pd.DateOffset(months=i)
                            nova["Vencimento"] = nova["Vencimento_dt"].strftime("%d/%m/%Y")
                            nova["Parcela"] = f"{atual+i:02d}/{total:02d}"
                            lista_proj_fixas.append(nova.to_dict())
                    except: pass
                else:
                    for i in range(1, 13):
                        nova = linha.copy()
                        nova["Vencimento_dt"] = nova["Vencimento_dt"] + pd.DateOffset(months=i)
                        nova["Vencimento"] = nova["Vencimento_dt"].strftime("%d/%m/%Y")
                        lista_proj_fixas.append(nova.to_dict())

        df_fixas_final = pd.concat([df_fixas_full, pd.DataFrame(lista_proj_fixas)], ignore_index=True) if lista_proj_fixas else df_fixas_full.copy()
        df_fixas_final = df_fixas_final.drop(columns=["Vencimento_dt"], errors="ignore")
        salvar_aba(df_fixas_final, "Dashboard Fixas")
except Exception as e:
    print(f"Erro ao processar fixas: {e}")
