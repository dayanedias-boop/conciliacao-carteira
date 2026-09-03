"""
Carteira de Direitos Creditórios — SCD
"""
import io, re, logging, base64, json, datetime, requests
from datetime import datetime
from pathlib import Path
import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.graph_objects as go

st.set_page_config(page_title="Carteira DC — SCD", page_icon="🏦", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.hero {
    background: linear-gradient(135deg, #1F3864 0%, #2F5496 100%);
    border-radius: 14px; padding: 28px 40px; color: white; margin-bottom: 20px;
}
.hero h1 { font-size: 24px; font-weight: 700; margin: 0 0 4px 0; }
.hero p  { font-size: 13px; opacity: 0.8; margin: 0; }
.stat-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 14px; margin: 20px 0; }
.stat { border-radius: 12px; padding: 18px 20px; border: 1px solid #E8EDF5; }
.stat .num { font-size: 26px; font-weight: 700; line-height: 1; }
.stat .lbl { font-size: 11px; font-weight: 500; margin-top: 5px; opacity: 0.75; }
.stat.verde { background:#F0FAF4; border-color:#C3E6CE; }
.stat.verde .num,.stat.verde .lbl { color:#1A7F3C; }
.stat.verm  { background:#FFF5F5; border-color:#FFC9C9; }
.stat.verm .num,.stat.verm .lbl   { color:#C92A2A; }
.stat.laran { background:#FFF8F0; border-color:#FFCC80; }
.stat.laran .num,.stat.laran .lbl { color:#E65100; }
.stat.cinza { background:#F7F9FC; border-color:#D0D9E8; }
.stat.cinza .num,.stat.cinza .lbl { color:#364F6B; }
.financeiro { display:grid; grid-template-columns:1fr 1fr 1fr; gap:14px; margin:14px 0; }
.fin-item { background:white; border-radius:10px; padding:14px 18px; border:1px solid #E8EDF5; }
.fin-item .fin-lbl { font-size:11px; color:#8898AA; font-weight:500; margin-bottom:5px; }
.fin-item .fin-val { font-size:18px; font-weight:700; color:#1F3864; }
.fin-item .fin-val.pos { color:#1A7F3C; }
.fin-item .fin-val.neg { color:#C92A2A; }
.hist-row { display:flex; align-items:center; gap:12px; padding:12px 16px;
            background:white; border-radius:10px; border:1px solid #E8EDF5; margin-bottom:8px; }
.hist-data { font-size:11px; color:#8898AA; min-width:90px; }
.hist-label { font-size:13px; font-weight:600; color:#1F3864; flex:1; }
.hist-badge { font-size:11px; padding:3px 10px; border-radius:20px;
              background:#EEF3FB; color:#2F5496; font-weight:600; }
.hist-val { font-size:11px; color:#2E4189; font-weight:600;
            background:#EEF3FB; padding:2px 8px; border-radius:12px; }
.main-header {
    background: linear-gradient(135deg, #1F2D5A 0%, #2E4189 100%);
    color: white; padding: 1.2rem 1.8rem; border-radius: 10px; margin-bottom: 1.2rem;
}
.main-header h1 { color: white; margin: 0; font-size: 1.3rem; }
.main-header p  { color: #D6DCF0; margin: 0.2rem 0 0; font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)


st.markdown("""
<div class="hero">
  <h1>🏦 Carteira de Direitos Creditórios — SCD</h1>
  <p>BCB 352 / Resolução BCB nº 4.966/2021</p>
</div>
""", unsafe_allow_html=True)

# ── FUNÇÕES CONCILIAÇÃO ─────────────────────────────────────────────
# ── GitHub ──────────────────────────────────────────────────────────────────
def github_headers():
    token = st.secrets.get("GITHUB_TOKEN", "")
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

def github_repo():
    return st.secrets.get("GITHUB_REPO", "")

def listar_historico():
    url = f"https://api.github.com/repos/{github_repo()}/contents/historico"
    r = requests.get(url, headers=github_headers(), timeout=10)
    if r.status_code != 200:
        return []
    return sorted(
        [{"nome":a["name"],"url_download":a["download_url"],"url_html":a["html_url"],"sha":a["sha"]}
         for a in r.json() if a["name"].endswith(".xlsx")],
        key=lambda x: x["nome"], reverse=True
    )

def salvar_no_github(nome_arquivo, excel_bytes, resumo_json):
    repo = github_repo(); hdrs = github_headers()
    conteudo_b64 = base64.b64encode(excel_bytes).decode()
    url_xlsx = f"https://api.github.com/repos/{repo}/contents/historico/{nome_arquivo}"
    r_check = requests.get(url_xlsx, headers=hdrs, timeout=10)
    payload_xlsx = {"message": f"Conciliação: {nome_arquivo}", "content": conteudo_b64}
    if r_check.status_code == 200:
        payload_xlsx["sha"] = r_check.json()["sha"]
    r = requests.put(url_xlsx, headers=hdrs, json=payload_xlsx, timeout=30)
    if r.status_code not in (200, 201):
        return False, f"Erro ao salvar Excel: {r.status_code}"
    idx_url = f"https://api.github.com/repos/{repo}/contents/historico/indice.json"
    r_idx = requests.get(idx_url, headers=hdrs, timeout=10)
    registros = []; idx_sha = None
    if r_idx.status_code == 200:
        try:
            registros = json.loads(base64.b64decode(r_idx.json()["content"]).decode())
            idx_sha = r_idx.json()["sha"]
        except Exception:
            registros = []
    registros.insert(0, resumo_json)
    registros = registros[:100]
    payload_idx = {"message": f"Índice atualizado: {nome_arquivo}",
                   "content": base64.b64encode(json.dumps(registros, ensure_ascii=False, indent=2).encode()).decode()}
    if idx_sha:
        payload_idx["sha"] = idx_sha
    r2 = requests.put(idx_url, headers=hdrs, json=payload_idx, timeout=15)
    if r2.status_code not in (200, 201):
        return False, f"Erro ao atualizar índice: {r2.status_code}"
    return True, "ok"

def carregar_indice():
    url = f"https://api.github.com/repos/{github_repo()}/contents/historico/indice.json"
    r = requests.get(url, headers=github_headers(), timeout=10)
    if r.status_code != 200:
        return []
    try:
        return json.loads(base64.b64decode(r.json()["content"]).decode())
    except Exception:
        return []

def baixar_excel_github(url_download):
    r = requests.get(url_download, timeout=30)
    return r.content if r.status_code == 200 else None

# ── Conciliação ─────────────────────────────────────────────────────────────
COLUNAS_CHAVE = ["id_cobranca","valor_cobranca","valor_pago",
                 "valor_excedente_pago","status_pagamento","contas_a_receber"]

C_AZE="1F3864"; C_AZM="2F5496"; C_BRA="FFFFFF"; C_CZC="F2F2F2"
C_VER="E2EFDA"; C_VRM="FFDCE1"; C_LAR="FCE4D6"; C_AMA="FFF2CC"
FMT_R="#,##0.00"; FMT_RD="#,##0.00;[RED]-#,##0.00"; FMT_N="#,##0"
COR_MOV={"🟢 NOVO":C_VER,"🔴 SAIU DA CARTEIRA":C_VRM,"🟡 ALTERADO":C_LAR,"⚪ SEM MOVIMENTAÇÃO":C_CZC}

def fill(c):    return PatternFill("solid", fgColor=c)
def borda():
    s=Side(style="thin",color="BFBFBF")
    return Border(left=s,right=s,top=s,bottom=s)
def alin(h="center",v="center",wrap=False,indent=0):
    return Alignment(horizontal=h,vertical=v,wrap_text=wrap,indent=indent)
def hdr(cell,val,bg=C_AZM,size=10):
    cell.value=val; cell.font=Font(name="Arial",bold=True,color="FFFFFF",size=size)
    cell.fill=fill(bg); cell.alignment=alin(wrap=True); cell.border=borda()
def cel(cell,val=None,bg=C_BRA,fmt=None,align="center"):
    if val is not None: cell.value=val
    cell.font=Font(name="Arial",size=10)
    cell.fill=fill(bg); cell.alignment=alin(align); cell.border=borda()
    if fmt: cell.number_format=fmt

def detectar_aba(file_bytes):
    wb=openpyxl.load_workbook(io.BytesIO(file_bytes),read_only=True,data_only=True)
    for nome in wb.sheetnames:
        ws=wb[nome]
        try:
            headers=[c.value for c in next(ws.iter_rows(min_row=1,max_row=1))]
            if "id_cobranca" in headers:
                wb.close(); return nome
        except StopIteration: continue
    wb.close()
    raise ValueError("Nenhuma aba com 'id_cobranca' encontrada.")

def ler_base(file_bytes, label):
    aba=detectar_aba(file_bytes)
    df=pd.read_excel(io.BytesIO(file_bytes),sheet_name=aba,dtype={"id_cobranca":str},engine="openpyxl")
    ausentes=[c for c in COLUNAS_CHAVE if c not in df.columns]
    if ausentes: raise ValueError(f"Colunas ausentes em '{label}': {ausentes}")
    for col in ["valor_cobranca","valor_pago","valor_excedente_pago","contas_a_receber"]:
        df[col]=pd.to_numeric(df[col],errors="coerce").fillna(0)
    df["status_pagamento"]=df["status_pagamento"].fillna("").astype(str)
    return df, aba

def conciliar(df_ant, df_atu):
    ant=df_ant[COLUNAS_CHAVE].copy(); atu=df_atu[COLUNAS_CHAVE].copy()
    merged=pd.merge(ant,atu,on="id_cobranca",how="outer",suffixes=("_ant","_atu"))
    for sfx in ("_ant","_atu"):
        for col in ["valor_cobranca","valor_pago","valor_excedente_pago","contas_a_receber"]:
            merged[f"{col}{sfx}"]=merged[f"{col}{sfx}"].fillna(0)
        merged[f"status_pagamento{sfx}"]=merged[f"status_pagamento{sfx}"].fillna("")
    def classificar(row):
        tem_ant=row["valor_cobranca_ant"]!=0 or row["status_pagamento_ant"]!=""
        tem_atu=row["valor_cobranca_atu"]!=0 or row["status_pagamento_atu"]!=""
        if not tem_ant and tem_atu:  return "🟢 NOVO"
        if tem_ant and not tem_atu:  return "🔴 SAIU DA CARTEIRA"
        if any([row["status_pagamento_ant"]!=row["status_pagamento_atu"],
                abs(row["valor_pago_ant"]-row["valor_pago_atu"])>0.001,
                abs(row["valor_excedente_pago_ant"]-row["valor_excedente_pago_atu"])>0.001,
                abs(row["contas_a_receber_ant"]-row["contas_a_receber_atu"])>0.001]):
            return "🟡 ALTERADO"
        return "⚪ SEM MOVIMENTAÇÃO"
    merged["tipo_movimentacao"]=merged.apply(classificar,axis=1)
    merged["var_valor_pago"]      =merged["valor_pago_atu"]          -merged["valor_pago_ant"]
    merged["var_excedente_pago"]  =merged["valor_excedente_pago_atu"]-merged["valor_excedente_pago_ant"]
    merged["var_contas_a_receber"]=merged["contas_a_receber_atu"]    -merged["contas_a_receber_ant"]
    def campos_alt(row):
        if row["tipo_movimentacao"]!="🟡 ALTERADO": return "—"
        p=[]
        if row["status_pagamento_ant"]!=row["status_pagamento_atu"]: p.append("status_pagamento")
        if abs(row["valor_pago_ant"]-row["valor_pago_atu"])>0.001: p.append("valor_pago")
        if abs(row["valor_excedente_pago_ant"]-row["valor_excedente_pago_atu"])>0.001: p.append("valor_excedente_pago")
        if abs(row["contas_a_receber_ant"]-row["contas_a_receber_atu"])>0.001: p.append("contas_a_receber")
        return " | ".join(p)
    merged["campos_alterados"]=merged.apply(campos_alt,axis=1)
    return merged

def gravar_aba_excel(wb,nome,df_d,titulo,tab_color,colunas):
    ws=wb.create_sheet(nome); ws.sheet_view.showGridLines=False
    ws.sheet_properties.tabColor=tab_color; ws.freeze_panes="A4"
    lc=get_column_letter(len(colunas))
    ws.row_dimensions[1].height=44; ws.merge_cells(f"A1:{lc}1")
    c=ws["A1"]; c.value=titulo
    c.font=Font(name="Arial",bold=True,color="FFFFFF",size=13)
    c.fill=fill(C_AZE); c.alignment=alin(wrap=True)
    ws.row_dimensions[2].height=20; ws.merge_cells(f"A2:{lc}2")
    c=ws["A2"]; c.value=f"Total de registros: {len(df_d):,}"
    c.font=Font(name="Arial",italic=True,color="595959",size=9)
    c.fill=fill(C_AMA); c.alignment=alin("left",indent=1); c.border=borda()
    ws.row_dimensions[3].height=42
    for ci,(_,hd,w,_,_) in enumerate(colunas,1):
        ws.column_dimensions[get_column_letter(ci)].width=w; hdr(ws.cell(3,ci),hd)
    for ri,(_,row) in enumerate(df_d.iterrows(),4):
        ws.row_dimensions[ri].height=18
        tipo=row.get("tipo_movimentacao","")
        bg=COR_MOV.get(tipo,C_BRA if ri%2==0 else C_CZC)
        for ci,(col_df,_,_,fmt,align) in enumerate(colunas,1):
            c2=ws.cell(ri,ci); val=row.get(col_df)
            if pd.isna(val) if not isinstance(val,str) else val=="nan": val=None
            cel(c2,val,bg=bg,fmt=fmt,align=align)

def gerar_excel_bytes_conc(conc,df_ant,df_atu,label_ant,label_atu):
    wb=openpyxl.Workbook()
    ws_i=wb.active; ws_i.title="📋 Instruções"
    ws_i.sheet_view.showGridLines=False; ws_i.sheet_properties.tabColor=C_AZE
    for col,w in [("A",3),("B",36),("C",54),("D",3)]: ws_i.column_dimensions[col].width=w
    ws_i.row_dimensions[1].height=48; ws_i.merge_cells("B1:C1")
    c=ws_i["B1"]; c.value="CONCILIAÇÃO — CARTEIRA DE DIREITOS CREDITÓRIOS"
    c.font=Font(name="Arial",bold=True,color="FFFFFF",size=14)
    c.fill=fill(C_AZE); c.alignment=alin(wrap=True)
    def s(row,txt):
        ws_i.row_dimensions[row].height=26; ws_i.merge_cells(f"B{row}:C{row}")
        c=ws_i[f"B{row}"]; c.value=txt
        c.font=Font(name="Arial",bold=True,color="FFFFFF",size=10)
        c.fill=fill(C_AZM); c.alignment=alin("left",indent=1); c.border=borda()
    def it(row,lbl,desc,bg=C_BRA):
        ws_i.row_dimensions[row].height=22
        cl=ws_i[f"B{row}"]; cl.value=lbl
        cl.font=Font(name="Arial",bold=True,color=C_AZE,size=10)
        cl.fill=fill(bg); cl.alignment=alin("left",indent=1); cl.border=borda()
        cd=ws_i[f"C{row}"]; cd.value=desc
        cd.font=Font(name="Arial",size=10,color="404040")
        cd.fill=fill(bg); cd.alignment=alin("left",indent=1,wrap=True); cd.border=borda()
    r=3
    s(r,"BASES UTILIZADAS"); r+=1
    it(r,"Mês Anterior",label_ant); r+=1
    it(r,"Mês Atual",label_atu); r+=1; r+=1
    s(r,"LEGENDA"); r+=1
    it(r,"🟢 NOVO","Presente somente no Mês Atual.",C_VER); r+=1
    it(r,"🔴 SAIU","Presente somente no Mês Anterior.",C_VRM); r+=1
    it(r,"🟡 ALTERADO","Presente em ambas com diferença nos valores.",C_LAR); r+=1
    it(r,"⚪ SEM MOVIMENTAÇÃO","Idêntico nas duas bases.",C_CZC)
    ws_r=wb.create_sheet("📊 Resumo")
    ws_r.sheet_view.showGridLines=False; ws_r.sheet_properties.tabColor=C_AZE
    for col,w in [("A",3),("B",44),("C",26),("D",3)]: ws_r.column_dimensions[col].width=w
    ws_r.row_dimensions[1].height=48; ws_r.merge_cells("B1:C1")
    c=ws_r["B1"]; c.value=f"RESUMO — CONCILIAÇÃO\n{label_ant}  ×  {label_atu}"
    c.font=Font(name="Arial",bold=True,color="FFFFFF",size=14)
    c.fill=fill(C_AZE); c.alignment=alin(wrap=True)
    def bloco(rs,titulo,itens):
        ws_r.row_dimensions[rs].height=28; ws_r.merge_cells(f"B{rs}:C{rs}")
        c=ws_r[f"B{rs}"]; c.value=titulo
        c.font=Font(name="Arial",bold=True,color="FFFFFF",size=11)
        c.fill=fill(C_AZM); c.alignment=alin(); c.border=borda()
        rr=rs+1
        for lbl,val,fmt,bg in itens:
            ws_r.row_dimensions[rr].height=22
            cl=ws_r[f"B{rr}"]; cl.value=lbl
            cl.font=Font(name="Arial",bold=True,color=C_AZE,size=10)
            cl.fill=fill(bg); cl.alignment=alin("left",indent=1); cl.border=borda()
            cv=ws_r[f"C{rr}"]; cv.value=val; cv.number_format=fmt
            cv.font=Font(name="Arial",bold=True,size=11)
            cv.fill=fill(bg); cv.alignment=alin("right"); cv.border=borda()
            rr+=1
        return rr+1
    m_nov=conc["tipo_movimentacao"]=="🟢 NOVO"
    m_sai=conc["tipo_movimentacao"]=="🔴 SAIU DA CARTEIRA"
    m_alt=conc["tipo_movimentacao"]=="🟡 ALTERADO"
    def soma(mask,col): return conc.loc[mask,col].fillna(0).sum()
    rr=3
    rr=bloco(rr,"POSIÇÃO GERAL",[
        (f"Cobranças — {label_ant}",len(df_ant),FMT_N,C_CZC),
        (f"Cobranças — {label_atu}",len(df_atu),FMT_N,C_CZC),
        (f"CAR — {label_ant} (R$)",df_ant["contas_a_receber"].sum(),FMT_R,C_CZC),
        (f"CAR — {label_atu} (R$)",df_atu["contas_a_receber"].sum(),FMT_R,C_CZC),
        ("Variação CAR (R$)",df_atu["contas_a_receber"].sum()-df_ant["contas_a_receber"].sum(),FMT_RD,C_AMA),
        (f"valor_pago — {label_ant} (R$)",df_ant["valor_pago"].sum(),FMT_R,C_CZC),
        (f"valor_pago — {label_atu} (R$)",df_atu["valor_pago"].sum(),FMT_R,C_CZC),
    ])
    rr=bloco(rr,"MOVIMENTAÇÕES — QUANTIDADE",[
        ("🟢 NOVO",m_nov.sum(),FMT_N,C_VER),
        ("🔴 SAIU DA CARTEIRA",m_sai.sum(),FMT_N,C_VRM),
        ("🟡 ALTERADO",m_alt.sum(),FMT_N,C_LAR),
        ("⚪ SEM MOVIMENTAÇÃO",(conc["tipo_movimentacao"]=="⚪ SEM MOVIMENTAÇÃO").sum(),FMT_N,C_CZC),
        ("TOTAL",len(conc),FMT_N,C_AMA),
    ])
    rr=bloco(rr,"MOVIMENTAÇÕES — CAR (R$)",[
        ("🟢 CAR Novos (R$)",soma(m_nov,"contas_a_receber_atu"),FMT_R,C_VER),
        ("🔴 CAR Saiu (R$)",soma(m_sai,"contas_a_receber_ant"),FMT_R,C_VRM),
        ("🟡 CAR Alterados Ant. (R$)",soma(m_alt,"contas_a_receber_ant"),FMT_R,C_LAR),
        ("🟡 CAR Alterados Atu. (R$)",soma(m_alt,"contas_a_receber_atu"),FMT_R,C_LAR),
        ("🟡 Variação CAR Alterados (R$)",soma(m_alt,"var_contas_a_receber"),FMT_RD,C_LAR),
    ])
    COLS_G=[("id_cobranca","id_cobrança",18,"@","center"),
            ("tipo_movimentacao","Tipo de\nMovimentação",26,"@","center"),
            ("campos_alterados","Campos Alterados",34,"@","left"),
            ("status_pagamento_ant",f"status_pag.\n({label_ant})",22,"@","center"),
            ("status_pagamento_atu",f"status_pag.\n({label_atu})",22,"@","center"),
            ("valor_cobranca_ant","valor_cobrança\nAnt. (R$)",18,FMT_R,"right"),
            ("valor_cobranca_atu","valor_cobrança\nAtu. (R$)",18,FMT_R,"right"),
            ("valor_pago_ant","valor_pago\nAnt. (R$)",18,FMT_R,"right"),
            ("valor_pago_atu","valor_pago\nAtu. (R$)",18,FMT_R,"right"),
            ("var_valor_pago","Var. valor_pago\n(R$)",18,FMT_RD,"right"),
            ("valor_excedente_pago_ant","excedente\nAnt. (R$)",20,FMT_R,"right"),
            ("valor_excedente_pago_atu","excedente\nAtu. (R$)",20,FMT_R,"right"),
            ("var_excedente_pago","Var. excedente\n(R$)",18,FMT_RD,"right"),
            ("contas_a_receber_ant","CAR\nAnt. (R$)",20,FMT_R,"right"),
            ("contas_a_receber_atu","CAR\nAtu. (R$)",20,FMT_R,"right"),
            ("var_contas_a_receber","Var. CAR (R$)",18,FMT_RD,"right")]
    conc_mov=conc[conc["tipo_movimentacao"]!="⚪ SEM MOVIMENTAÇÃO"].sort_values("tipo_movimentacao")
    gravar_aba_excel(wb,"🔄 Conciliação Geral",conc_mov,f"CONCILIAÇÃO GERAL — {label_ant} × {label_atu}","833C00",COLS_G)
    conc_alt=conc[conc["tipo_movimentacao"]=="🟡 ALTERADO"].sort_values("campos_alterados")
    COLS_A=[c for c in COLS_G if c[0]!="tipo_movimentacao"]
    gravar_aba_excel(wb,"🟡 Alterados",conc_alt,f"ALTERADOS — {label_ant} × {label_atu}","833C00",COLS_A)
    conc_nov=conc[conc["tipo_movimentacao"]=="🟢 NOVO"]
    COLS_N=[("id_cobranca","id_cobrança",18,"@","center"),("valor_cobranca_atu","valor_cobrança (R$)",18,FMT_R,"right"),
            ("status_pagamento_atu","status_pagamento",22,"@","center"),("valor_pago_atu","valor_pago (R$)",18,FMT_R,"right"),
            ("valor_excedente_pago_atu","excedente_pago (R$)",20,FMT_R,"right"),("contas_a_receber_atu","contas_a_receber (R$)",22,FMT_R,"right")]
    gravar_aba_excel(wb,"🟢 Novos",conc_nov,f"NOVOS — {label_atu}","375623",COLS_N)
    conc_sai=conc[conc["tipo_movimentacao"]=="🔴 SAIU DA CARTEIRA"]
    COLS_S=[("id_cobranca","id_cobrança",18,"@","center"),("valor_cobranca_ant","valor_cobrança (R$)",18,FMT_R,"right"),
            ("status_pagamento_ant","status_pagamento",22,"@","center"),("valor_pago_ant","valor_pago (R$)",18,FMT_R,"right"),
            ("valor_excedente_pago_ant","excedente_pago (R$)",20,FMT_R,"right"),("contas_a_receber_ant","contas_a_receber (R$)",22,FMT_R,"right")]
    gravar_aba_excel(wb,"🔴 Saíram",conc_sai,f"SAÍRAM — {label_ant}","9C0006",COLS_S)
    ordem=["📋 Instruções","📊 Resumo","🔄 Conciliação Geral","🟡 Alterados","🟢 Novos","🔴 Saíram"]
    wb._sheets.sort(key=lambda s: ordem.index(s.title) if s.title in ordem else 99)
    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

# ── Interface ────────────────────────────────────────────────────────────────


# ── FUNÇÕES PDD C5 ──────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════
# TABELAS NORMATIVAS — NÃO ALTERAR SEM FUNDAMENTAÇÃO NORMATIVA
# ══════════════════════════════════════════════════════════════════════

ANEXO2 = {
    0:  0.500, 1:  0.534, 2:  0.568, 3:  0.602,
    4:  0.636, 5:  0.670, 6:  0.704, 7:  0.738,
    8:  0.772, 9:  0.806, 10: 0.840, 11: 0.874,
    12: 0.908, 13: 0.942, 14: 0.976, 15: 1.000,
}
PCT_50_7        = 0.034
PCT_E2_31_60    = 0.15
PCT_E2_61_90    = 0.38
POCI_THRESHOLD  = 90
DIAS_MES        = 30.44
PRAZO_E2_DIAS   = 450
PRAZO_E3_DIAS   = 42 * 30   # 1260

DESAGIO_FAIXA = [
    (0,0,0.038),(1,14,0.038),(15,30,0.150),(31,60,0.260),(61,90,0.280),
    (91,120,0.300),(121,150,0.320),(151,180,0.340),(181,210,0.360),
    (211,240,0.380),(241,270,0.400),(271,300,0.420),(301,330,0.440),
    (331,360,0.460),(361,390,0.480),(391,420,0.500),(421,450,0.520),
    (451,480,0.540),(481,510,0.560),(511,540,0.580),(541,570,0.600),
    (571,600,0.620),(601,630,0.640),(631,660,0.660),(661,690,0.680),
    (691,720,0.700),(721,99999,0.720),
]
PECLD_FAIXA = [
    (0,90,0.000),(91,120,0.534),(121,150,0.568),(151,180,0.602),
    (181,210,0.636),(211,240,0.670),(241,270,0.704),(271,300,0.738),
    (301,330,0.772),(331,360,0.806),(361,390,0.840),(391,420,0.874),
    (421,450,0.908),(451,480,0.942),(481,510,0.976),(511,99999,1.000),
]


# ══════════════════════════════════════════════════════════════════════
# FUNÇÕES DE CÁLCULO
# ══════════════════════════════════════════════════════════════════════

def lookup(aging, table):
    for lo, hi, val in table:
        if lo <= aging <= hi:
            return val
    return table[-1][2]

def meses_inad(aging):
    return 0 if aging <= 90 else min(int((aging - 90) / DIAS_MES), 15)

def tje_na_compra(aging_compra):
    d = lookup(aging_compra, DESAGIO_FAIXA)
    k = lookup(aging_compra, PECLD_FAIXA)
    if aging_compra <= POCI_THRESHOLD:
        return round((1 / (1 - d)) ** (365 / PRAZO_E2_DIAS) - 1, 4)
    ratio = (1 - k) / (1 - d)
    return 0.0 if ratio <= 1.0 else round(ratio ** (365 / PRAZO_E3_DIAS) - 1, 4)

def ajuste_centavo(series, target):
    diff = round(target - series.sum(), 2)
    if diff == 0:
        return series
    n   = int(round(abs(diff) / 0.01))
    sig = 1 if diff > 0 else -1
    s   = series.copy()
    s.iloc[:n] = (s.iloc[:n] + sig * 0.01).round(2)
    return s

def get_grupo(row):
    if row['is_poci'] and row['aging_dias'] > 630:
        return 'C1_POCI'
    if row['is_poci']:
        return f"C4_POCI_{meses_inad(row['aging_dias'])}m"
    if row['aging_dias'] <= 60:
        return 'E2_31_60'
    if row['aging_dias'] <= 90:
        return 'E2_61_90'
    return f"E3_{meses_inad(row['aging_dias'])}m"


# ══════════════════════════════════════════════════════════════════════
# LEITURA DE ARQUIVOS (aceita buffer do Streamlit)
# ══════════════════════════════════════════════════════════════════════

def ler_car(buf, sheet=None):
    df = pd.read_excel(buf, sheet_name=sheet or 0)
    df.columns = [c.lower().strip() for c in df.columns]
    renomear = {
        'id': 'id_cobranca', 'id cobrança': 'id_cobranca',
        'saldo': 'contas_a_receber', 'car': 'contas_a_receber',
        'face': 'valor_cobranca', 'ead': 'valor_cobranca',
        'aging': 'aging_dias', 'dias_atraso': 'aging_dias',
        'vencimento': 'data_vencimento',
    }
    df = df.rename(columns={k: v for k, v in renomear.items() if k in df.columns})
    for col in ['contas_a_receber', 'valor_cobranca', 'valor_pago', 'aging_dias']:
        df[col] = pd.to_numeric(df.get(col, 0), errors='coerce').fillna(0)
    df['id_cobranca'] = pd.to_numeric(df.get('id_cobranca'), errors='coerce')
    df['data_vencimento'] = pd.to_datetime(df.get('data_vencimento'), errors='coerce')
    df = df[df['id_cobranca'].notna() & (df['contas_a_receber'] > 0)]
    df = df.drop_duplicates('id_cobranca').copy().reset_index(drop=True)
    df['id_cobranca'] = df['id_cobranca'].astype(int)
    return df

def ler_cessao(buf, sheet, tipo, data_cessao_override=None):
    df = pd.read_excel(buf, sheet_name=sheet or 0)
    if tipo == "padrao":
        df = df.rename(columns={
            'ID COBRANÇA': 'id', 'VALOR FACE BOLETO': 'face',
            'PREÇO DE AQUISIÇÃO': 'custo_aq', 'DESAGIO': 'desagio',
            'CESSÃO': 'data_cessao', 'DATA DE VENCIMENTO BOLETO': 'data_venc_orig',
        })
        df['id'] = pd.to_numeric(df.get('id'), errors='coerce')
        df = df[df['id'].notna() & df.get('data_cessao', pd.Series()).notna()].copy()
        df['data_cessao']    = pd.to_datetime(df['data_cessao'], errors='coerce')
        df['data_venc_orig'] = pd.to_datetime(df.get('data_venc_orig'), errors='coerce')
        df['aging_compra']   = (df['data_cessao'] - df['data_venc_orig']).dt.days.clip(lower=0).fillna(0).astype(int)
    else:  # nova
        df = df.rename(columns={
            'ID': 'id', 'VALOR FACE DIREITO ECONÔMICO': 'face',
            'PREÇO AQUISIÇÃO': 'custo_aq', 'DESÁGIO': 'desagio',
            'DIAS EM ATRASO': 'aging_compra',
        })
        df['id'] = pd.to_numeric(df.get('id'), errors='coerce')
        if data_cessao_override:
            df['data_cessao'] = pd.to_datetime(data_cessao_override, dayfirst=True)
        df['aging_compra'] = pd.to_numeric(df.get('aging_compra', 0), errors='coerce').fillna(0).astype(int)
    for col in ['face', 'custo_aq', 'desagio']:
        df[col] = pd.to_numeric(df.get(col, 0), errors='coerce').fillna(0)
    df = df[df['id'].notna()].drop_duplicates('id').copy()
    df['id']      = df['id'].astype(int)
    df['is_poci'] = df['aging_compra'] > POCI_THRESHOLD
    df['tje']     = df['aging_compra'].apply(tje_na_compra)
    return df[['id', 'face', 'custo_aq', 'desagio', 'aging_compra', 'is_poci', 'tje']]


# ══════════════════════════════════════════════════════════════════════
# MOTOR PDD
# ══════════════════════════════════════════════════════════════════════

def calcular_pdd(df_car, df_cessoes):
    df = df_car.merge(
        df_cessoes[['id', 'face', 'custo_aq', 'desagio', 'aging_compra', 'is_poci', 'tje']],
        left_on='id_cobranca', right_on='id', how='left'
    )
    df['face']          = df['face'].fillna(df['valor_cobranca'])
    df['custo_aq']      = df['custo_aq'].fillna(0)
    df['desagio']       = df['desagio'].fillna(0)
    df['aging_compra']  = df['aging_compra'].fillna(0).astype(int)
    df['is_poci']       = df['is_poci'].fillna(False)
    df['tje']           = df['tje'].fillna(0.0)
    df['liq_contabil']  = (df['face'] - df['desagio']).round(2)
    df['liq_contabil']  = df['liq_contabil'].where(df['is_poci'], other=None)
    df['grupo']         = df.apply(get_grupo, axis=1)
    df['v40'] = 0.0; df['v50'] = 0.0; df['v60'] = 0.0

    for grp in df['grupo'].unique():
        m     = df['grupo'] == grp
        m_idx = df.index[m]
        car_g = df.loc[m, 'contas_a_receber'].sum()
        if not car_g:
            continue

        if grp == 'E2_31_60':
            t = round(car_g * PCT_E2_31_60, 2)
            df.loc[m_idx, 'v60'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*t).round(2), t).values

        elif grp == 'E2_61_90':
            t = round(car_g * PCT_E2_61_90, 2)
            df.loc[m_idx, 'v60'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*t).round(2), t).values

        elif grp == 'C1_POCI':
            pdd_c1 = df.loc[m, 'liq_contabil'].sum()
            df.loc[m_idx, 'v40'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*pdd_c1).round(2), pdd_c1).values

        elif grp.startswith('C4_POCI_'):
            mi      = int(re.search(r'_(\d+)m', grp).group(1))
            liq_g   = df.loc[m, 'liq_contabil'].sum()
            t40     = round(car_g * ANEXO2.get(mi, 1.0), 2)
            t50     = round(car_g * PCT_50_7, 2)
            if t40 + t50 > liq_g:
                t50 = max(0, round(liq_g - t40, 2))
            df.loc[m_idx, 'v40'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*t40).round(2), t40).values
            df.loc[m_idx, 'v50'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*t50).round(2), t50).values

        elif grp.startswith('E3_'):
            mi  = int(re.search(r'_(\d+)m', grp).group(1))
            t40 = round(car_g * ANEXO2.get(mi, 1.0), 2)
            t50 = round(car_g * PCT_50_7, 2)
            df.loc[m_idx, 'v40'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*t40).round(2), t40).values
            df.loc[m_idx, 'v50'] = ajuste_centavo(
                (df.loc[m,'contas_a_receber']/car_g*t50).round(2), t50).values

    df['vtot'] = (df['v40'] + df['v50'] + df['v60']).round(2)
    return df


# ══════════════════════════════════════════════════════════════════════
# GERADOR EXCEL (em memória → download)
# ══════════════════════════════════════════════════════════════════════

DB="1F2D5A"; MB="2E4189"; LB="D6DCF0"; AC="3A5CC5"
WH="FFFFFF"; GR="1E7145"; RE="C00000"; AM="FFF2CC"
TBG="F0F3FB"; CZ="F5F5F5"; E2="EBF3FB"; E3="FCE4D6"; PO="D9E2F3"; GBGR="E2EFDA"
fmt_r = '_("R$"* #,##0.00_);_("R$"* (#,##0.00);_("R$"* "-"??_);_(@_)'
fmt_i = '#,##0'

def tb():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def sc(ws, r, c, v, bold=False, fc="000000", bg=None, ha="center",
       wrap=False, fmt=None, it=False, sz=10):
    try:
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name="Calibri", bold=bold, color=fc, size=sz, italic=it)
        if bg: cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
        cell.border = tb()
        if fmt: cell.number_format = fmt
    except AttributeError: pass

def mg(ws, r1, c1, r2, c2, v, bold=False, fc=WH, bg=DB,
       ha="center", wrap=True, sz=10, it=False):
    ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
    cell = ws.cell(row=r1, column=c1, value=v)
    cell.font = Font(name="Calibri", bold=bold, color=fc, size=sz, italic=it)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    cell.border = tb()

def gerar_excel_bytes_pdd(df, data_base):
    wb   = Workbook()
    dstr = data_base.strftime('%d/%m/%Y')
    S40  = round(df.v40.sum(), 2); S50 = round(df.v50.sum(), 2)
    S60  = round(df.v60.sum(), 2); STOT = round(df.vtot.sum(), 2)
    TCAR = round(df.contas_a_receber.sum(), 2)
    TEAD = round(df.valor_cobranca.sum(), 2)
    TPAG = round(df.valor_pago.sum(), 2)
    TQTD = len(df)

    # ── Aba 1: Premissas ─────────────────────────────────────────
    ws0 = wb.active; ws0.title = "Premissas e Metodologia"
    for col, w in zip('ABCDEFGHIJ',[30,14,14,14,12,12,12,12,12,30]):
        ws0.column_dimensions[col].width = w
    mg(ws0,1,1,1,10,f"PREMISSAS E METODOLOGIA — PDD C5 SCD  |  BCB 352 / Res. 4.966/2021  |  {dstr}",bold=True,sz=12,bg=DB)
    mg(ws0,3,1,3,10,"1. CLASSIFICAÇÃO DE INSTRUMENTOS",bold=True,sz=10,bg=MB)
    for i,h in enumerate(["Grupo","Critério","% PDD","Conta COSIF","Teto?","Tipo PEC","","","",""],1):
        sc(ws0,4,i,h,bold=True,fc=WH,bg=AC,ha="center",sz=9,wrap=True)
    for r_off,(grp,crit,pct,conta,teto,tipo,bg) in enumerate([
        ("E2 ≤90d","Adimplente/até 90d vencido","15%/38% CAR","60-0","Não","Perda Esperada",CZ),
        ("E3 >90d","Inadimplido aging>90d","Anexo II +3,4%","40-4+50-7","Sim","Perda Incorrida",E3),
        ("C1 POCI >630d","Adq. >90d na compra, aging>630d","100% liq.contábil","40-4","Não","Art.15 Res.4966",PO),
        ("C4 POCI 91-630d","Adq. >90d na compra, aging 91-630d","Anexo II +3,4%","40-4+50-7","Sim","Art.14 Res.4966",PO),
    ],5):
        for ci,v in enumerate([grp,crit,pct,conta,teto,tipo,"","","",""],1):
            sc(ws0,r_off,ci,v,bg=bg,ha="left" if ci<=2 else "center",sz=9,bold=(ci==1))
    mg(ws0,10,1,10,10,"2. TABELA E2 — Perda Esperada (60-0)",bold=True,sz=10,bg=MB)
    for i,h in enumerate(["Faixa","% CAR","Conta","","","","","","",""],1):
        sc(ws0,11,i,h,bold=True,fc=WH,bg=AC,ha="center",sz=9)
    for r_off,(fl,pct) in enumerate([("31-60 dias","15,0%"),("61-90 dias","38,0%")],12):
        sc(ws0,r_off,1,fl,bg=E2,ha="left",sz=9)
        sc(ws0,r_off,2,pct,bg=E2,ha="center",bold=True,fc=MB,sz=11)
        sc(ws0,r_off,3,"1.8.1.50.10.60-0",bg=E2,sz=8,it=True,ha="left")
        for c in range(4,11): sc(ws0,r_off,c,"",bg=E2)
    mg(ws0,15,1,15,10,"3. TABELA ANEXO II C5 — Estágio 3 (% exatos +3,4 p.p./mês)",bold=True,sz=10,bg=MB)
    for i,h in enumerate(["Meses Inad.","Aging mín.","Aging máx.","% 40-4","% 50-7","% Total","Conta 40-4","Conta 50-7","Teto POCI",""],1):
        sc(ws0,16,i,h,bold=True,fc=WH,bg=AC,ha="center",sz=9,wrap=True)
    e3_tab=[(0,91,121,0.500),(1,122,151,0.534),(2,152,182,0.568),(3,183,212,0.602),
            (4,213,243,0.636),(5,244,273,0.670),(6,274,304,0.704),(7,305,334,0.738),
            (8,335,365,0.772),(9,366,395,0.806),(10,396,426,0.840),(11,427,456,0.874),
            (12,457,487,0.908),(13,488,517,0.942),(14,518,548,0.976),(15,549,9999,1.000)]
    for r_off,(mi,amin,amax,p40) in enumerate(e3_tab,17):
        bg=PO if mi>=15 else (E3 if r_off%2==0 else "FDE9D9")
        sc(ws0,r_off,1,f"{'≥' if mi>=15 else ''}{mi}m",bold=(mi==15),ha="center",bg=bg,sz=9)
        sc(ws0,r_off,2,f"{amin}d",ha="center",bg=bg,sz=9)
        sc(ws0,r_off,3,f"{'∞' if amax==9999 else str(amax)+'d'}",ha="center",bg=bg,sz=9)
        sc(ws0,r_off,4,p40,ha="center",bg=bg,fmt="0.0%",bold=True,
           fc=RE if p40>=0.9 else "1F4E79",sz=11)
        sc(ws0,r_off,5,0.034,ha="center",bg=bg,fmt="0.0%",sz=9,fc="595959")
        sc(ws0,r_off,6,round(p40+0.034,3),ha="center",bg=bg,fmt="0.0%",bold=True,sz=10)
        sc(ws0,r_off,7,"1.8.1.50.10.40-4",bg=bg,sz=8,it=True)
        sc(ws0,r_off,8,"1.8.1.50.10.50-7",bg=bg,sz=8,it=True)
        sc(ws0,r_off,9,"Liq.contábil",ha="center",bg=bg,sz=9,fc=RE)
        sc(ws0,r_off,10,"",bg=bg)
        ws0.row_dimensions[r_off].height=20
    tje_r=34
    mg(ws0,tje_r,1,tje_r,10,"4. TJE — PREMISSAS CADOC 3040 (Taxa de Juros Efetiva % a.a.)",bold=True,sz=10,bg=MB)
    for r_off,(desc,val,obs) in enumerate([
        ("Prazo recebimento adimpl. vencidos (dias)",PRAZO_E2_DIAS,"← E2: 15m×30d"),
        ("Prazo recuperação POCI (meses)",42,"← E3/POCI: 42m×30d=1260d"),
    ],tje_r+1):
        bg=LB if r_off==tje_r+1 else CZ
        ws0.merge_cells(start_row=r_off,start_column=1,end_row=r_off,end_column=2)
        c=ws0.cell(row=r_off,column=1,value=desc)
        c.font=Font(name="Calibri",size=9,color=MB); c.fill=PatternFill("solid",fgColor=bg)
        c.alignment=Alignment(horizontal="left",vertical="center"); c.border=tb()
        sc(ws0,r_off,3,val,bold=True,fc=RE,bg=bg,ha="center",fmt="0",sz=11)
        ws0.merge_cells(start_row=r_off,start_column=4,end_row=r_off,end_column=10)
        c2=ws0.cell(row=r_off,column=4,value=obs)
        c2.font=Font(name="Calibri",size=8,italic=True,color="595959")
        c2.fill=PatternFill("solid",fgColor=bg)
        c2.alignment=Alignment(horizontal="left",vertical="center"); c2.border=tb()
        ws0.row_dimensions[r_off].height=22
    mg(ws0,tje_r+4,1,tje_r+4,10,
       f"FÓRMULA TJE (fixa na data de aquisição):\n"
       f"  E2 (aging_compra ≤90d): TJE=(1/(1-deságio%))^(365/{PRAZO_E2_DIAS})−1\n"
       f"  E3/POCI: ratio=(1-PECLD%)/(1-deságio%); TJE=ratio^(365/{PRAZO_E3_DIAS})−1 se ratio>1, senão 0\n"
       f"  Resultados: faixa4 31-60d→27.7% | faixa5 61-90d→30.5% | E3/POCI→0.0%",
       fc="000000",bg=AM,ha="left",sz=9,bold=False)
    ws0.row_dimensions[tje_r+4].height=54

    # ── Aba 2: COSIF ─────────────────────────────────────────────
    ws1=wb.create_sheet("PDD C5 — COSIF Completo")
    for col,w in zip('ABCDEFGHIJKLMNO',[5,64,8,18,18,18,10,7,14,18,18,18,16,14,44]):
        ws1.column_dimensions[col].width=w
    mg(ws1,1,1,1,15,f"PDD CARTEIRA C5 — COSIF  |  Posição {dstr}  |  BCB 352 / Res. 4.966/2021",bold=True,sz=12,bg=DB)
    mg(ws1,2,1,2,15,
       f"Qtd: {TQTD:,} cob. | CAR: R${TCAR:,.2f} | EAD: R${TEAD:,.2f} | PDD: R${STOT:,.2f} ({STOT/TEAD*100:.1f}% EAD)",
       it=True,sz=9,fc=LB,bg=MB)
    for i,h in enumerate(["#","Cessão / Faixa","Qtd","EAD Face","CAR (saldo)","CAR Base PDD",
                           "Estágio","Meses","% / Base","40-4","50-7","60-0","TOTAL PDD","Pago Parc.","Fundamento"],1):
        sc(ws1,4,i,h,bold=True,fc=WH,bg=AC,ha="center",wrap=True)
    ws1.row_dimensions[4].height=40

    grp_sum=df.groupby('grupo').agg(
        qtd=('id_cobranca','count'),ead=('valor_cobranca','sum'),
        car=('contas_a_receber','sum'),liq=('liq_contabil','sum'),pago=('valor_pago','sum'),
        v40=('v40','sum'),v50=('v50','sum'),v60=('v60','sum'),vtot=('vtot','sum'),
        aging_min=('aging_dias','min'),aging_max=('aging_dias','max'),
    ).round(2)

    def ordem(g):
        if g=='C1_POCI':          return 0
        if g.startswith('C4_'): return 1+int(re.search(r'_(\d+)m',g).group(1))
        if g=='E2_31_60':         return 100
        if g=='E2_61_90':         return 101
        if g.startswith('E3_'): return 200+int(re.search(r'_(\d+)m',g).group(1))
        return 999

    def mi_g(g):
        m=re.search(r'_(\d+)m',g); return int(m.group(1)) if m else 0

    GDESC={
        'C1_POCI':'Cessão 1 — POCI orig. (aging >630d)',
        'E2_31_60':'Cessões — 31 a 60 dias (E2)',
        'E2_61_90':'Cessões — 61 a 90 dias (E2)',
    }
    for g in grp_sum.index:
        if g.startswith('C4_POCI_'):
            mi=mi_g(g); GDESC[g]=f'C4 POCI — {mi}-{mi+1}m inad.'
        elif g.startswith('E3_'):
            mi=mi_g(g); GDESC[g]=f'Cessões — {mi}-{mi+1}m inad. (E3)'

    BG_G  = lambda g: PO if 'POCI' in g else (E2 if 'E2' in g else E3)
    EST_G = lambda g: 'E3 POCI' if 'POCI' in g else ('E2' if 'E2' in g else 'E3')

    def PCT_G(g):
        if g=='C1_POCI': return '100% liq.'
        if 'POCI' in g: return f"{ANEXO2.get(mi_g(g),0)*100:.1f}%+3,4%≤liq."
        if g=='E2_31_60': return '15%'
        if g=='E2_61_90': return '38%'
        return f"{ANEXO2.get(mi_g(g),0)*100:.1f}%+3,4%"

    def FUND_G(g):
        if g=='C1_POCI': return 'Art.15 Res.4.966/2021 — 100% liq.contábil'
        if 'POCI' in g: return f"Anexo II {ANEXO2.get(mi_g(g),0)*100:.1f}%+50-7:3,4% | teto=liq.contábil"
        if 'E2' in g: return '60-0 Perda Esperada | BCB352'
        return f"40-4:{ANEXO2.get(mi_g(g),0)*100:.1f}%+50-7:3,4% | Anexo II C5"

    for r_off,grp in enumerate(sorted(grp_sum.index,key=ordem),5):
        row=grp_sum.loc[grp]; bg=BG_G(grp); est=EST_G(grp)
        fc_e="595959" if "POCI" in est else ("2E75B6" if est=="E2" else "C55A11")
        base=float(row.liq) if 'POCI' in grp else float(row.car)
        mi_s="—" if "E2" in grp else ("≥21m" if grp=="C1_POCI" else (f"{mi_g(grp)}-{mi_g(grp)+1}m" if mi_g(grp) else "—"))
        sc(ws1,r_off,1,r_off-4,ha="center",bg=bg)
        sc(ws1,r_off,2,GDESC.get(grp,grp),bg=bg,ha="left",wrap=True)
        sc(ws1,r_off,3,int(row.qtd),ha="right",bg=bg,fmt=fmt_i)
        sc(ws1,r_off,4,float(row.ead),ha="right",bg=bg,fmt=fmt_r)
        sc(ws1,r_off,5,float(row.car),ha="right",bg=bg,fmt=fmt_r,bold=True)
        sc(ws1,r_off,6,base,ha="right",bg=bg,fmt=fmt_r,bold=True,fc=MB)
        sc(ws1,r_off,7,est,ha="center",bg=bg,bold=True,fc=fc_e)
        sc(ws1,r_off,8,mi_s,ha="center",bg=bg,sz=9)
        sc(ws1,r_off,9,PCT_G(grp),ha="center",bg=bg,sz=9,it=True)
        sc(ws1,r_off,10,float(row.v40) if row.v40>0 else "-",ha="right",bg=bg,fmt=fmt_r if row.v40>0 else "@")
        sc(ws1,r_off,11,float(row.v50) if row.v50>0 else "-",ha="right",bg=bg,fmt=fmt_r if row.v50>0 else "@")
        sc(ws1,r_off,12,float(row.v60) if row.v60>0 else "-",ha="right",bg=bg,fmt=fmt_r if row.v60>0 else "@")
        sc(ws1,r_off,13,float(row.vtot),ha="right",bg=bg,bold=True,fmt=fmt_r)
        sc(ws1,r_off,14,float(row.pago) if row.pago>0 else "-",ha="right",bg=bg,fmt=fmt_r if row.pago>0 else "@")
        sc(ws1,r_off,15,FUND_G(grp),it=True,sz=9,bg=bg,wrap=True)
        ws1.row_dimensions[r_off].height=34

    lr=5+len(grp_sum)
    sc(ws1,lr,1,"",bg=TBG)
    mg(ws1,lr,2,lr,2,"TOTAL GERAL",bold=True,fc="000000",bg=TBG,ha="left")
    sc(ws1,lr,3,TQTD,bold=True,ha="right",bg=TBG,fmt=fmt_i)
    sc(ws1,lr,4,TEAD,bold=True,ha="right",bg=TBG,fmt=fmt_r)
    sc(ws1,lr,5,TCAR,bold=True,ha="right",bg=TBG,fmt=fmt_r,fc=GR)
    base_tot=round(sum(float(grp_sum.loc[g,'liq']) if 'POCI' in g else float(grp_sum.loc[g,'car']) for g in grp_sum.index),2)
    sc(ws1,lr,6,base_tot,bold=True,ha="right",bg=TBG,fmt=fmt_r,fc=MB)
    for c in [7,8,9]: sc(ws1,lr,c,"",bg=TBG)
    sc(ws1,lr,10,S40,bold=True,ha="right",bg=TBG,fmt=fmt_r)
    sc(ws1,lr,11,S50,bold=True,ha="right",bg=TBG,fmt=fmt_r)
    sc(ws1,lr,12,S60,bold=True,ha="right",bg=TBG,fmt=fmt_r)
    sc(ws1,lr,13,STOT,bold=True,ha="right",bg=TBG,fmt=fmt_r,fc=RE)
    sc(ws1,lr,14,TPAG,bold=True,ha="right",bg=TBG,fmt=fmt_r)
    sc(ws1,lr,15,"",bg=TBG)

    lc_row=lr+2
    mg(ws1,lc_row,1,lc_row,15,"LANÇAMENTOS CONTÁBEIS — COSIF",bold=True,sz=10,bg=MB)
    for i,h in enumerate(["D/C","Conta COSIF","","","","Descrição","","","","","Valor (R$)","","","","Comp."],1):
        sc(ws1,lc_row+1,i,h,bold=True,fc=WH,bg=AC,ha="center",wrap=True)
    for r_off,(dc,conta,desc,val,bg) in enumerate([
        ("D","3.d.1.x.x — Despesa PEC/PDD","Total PDD",STOT,WH),
        ("C","1.8.1.50.10.40-4 — (-) Perda Incorrida","C1+C4POCI+E3",S40,LB),
        ("C","1.8.1.50.10.50-7 — (-) Prov. Adicional 3,4%","POCI+E3",S50,WH),
        ("C","1.8.1.50.10.60-0 — (-) Perda Esperada","31-60d + 61-90d",S60,LB),
    ],lc_row+2):
        sc(ws1,r_off,1,dc,ha="center",bg=bg,bold=True)
        ws1.merge_cells(start_row=r_off,start_column=2,end_row=r_off,end_column=5)
        c=ws1.cell(row=r_off,column=2,value=conta); c.font=Font(name="Calibri",size=9)
        c.fill=PatternFill("solid",fgColor=bg); c.alignment=Alignment(horizontal="left",vertical="center"); c.border=tb()
        ws1.merge_cells(start_row=r_off,start_column=6,end_row=r_off,end_column=10)
        c2=ws1.cell(row=r_off,column=6,value=desc); c2.font=Font(name="Calibri",size=9)
        c2.fill=PatternFill("solid",fgColor=bg); c2.alignment=Alignment(horizontal="left",vertical="center"); c2.border=tb()
        sc(ws1,r_off,11,val,bold=True,ha="right",bg=bg,fmt=fmt_r)
        for c_ in [12,13,14,15]: sc(ws1,r_off,c_,"",bg=bg)
        ws1.row_dimensions[r_off].height=28
    ok=lc_row+6
    ws1.merge_cells(f'A{ok}:J{ok}')
    cell=ws1.cell(row=ok,column=1,value=f"  ✅  D(Despesa) = C(40-4)+C(50-7)+C(60-0) = R$ {STOT:,.2f}  |  % PDD/EAD = {STOT/TEAD*100:.1f}%")
    cell.font=Font(name="Calibri",bold=True,color=GR); cell.fill=PatternFill("solid",fgColor=GBGR)
    cell.alignment=Alignment(horizontal="left",vertical="center"); cell.border=tb()
    sc(ws1,ok,11,STOT,bold=True,ha="right",bg=GBGR,fc=GR,fmt=fmt_r)
    for c_ in [12,13,14,15]: sc(ws1,ok,c_,"",bg=GBGR)

    # ── Aba 3: Composição ────────────────────────────────────────
    ws2=wb.create_sheet("Composição por Cobrança")
    ws2.freeze_panes="A4"
    for col,w in zip('ABCDEFGHIJKLMNOPQR',[14,12,14,14,14,12,14,14,14,12,12,12,12,14,22,14,14,14]):
        ws2.column_dimensions[col].width=w
    mg(ws2,1,1,1,18,f"COMPOSIÇÃO PDD POR COBRANÇA  |  {dstr}  |  BCB 352 / Res. 4.966/2021",bold=True,sz=11)
    mg(ws2,2,1,2,18,f"Total: {TQTD:,} cob. | CAR: R${TCAR:,.2f} | PDD: R${STOT:,.2f} | Coluna Q: TJE fixo (data de aquisição)",it=True,sz=9,fc=LB,bg=MB)
    for i,h in enumerate(["ID Cobrança","Data Venc.","Cessão/Grupo","Faixa / Estágio","Conta CNPJ","Aging (dias)",
                           "EAD Face","Pago Parcial","CAR (saldo)","Liq.Contábil","40-4","50-7","60-0",
                           "TOTAL PDD","Base Cálculo","Custo Aq.","Aging Compra","TJE % a.a."],1):
        sc(ws2,3,i,h,bold=True,fc=WH,bg=AC,ha="center",wrap=True,sz=9)
    ws2.row_dimensions[3].height=40

    row_w=4
    for _,rd in df.iterrows():
        r=row_w; grp=str(rd.get('grupo',''))
        bg=(PO if r%2==0 else "E8EEF8") if 'POCI' in grp else \
           (E3 if r%2==0 else "FDE9D9") if 'E3' in grp else \
           (E2 if r%2==0 else "F0F7FF")
        def cv(col,v,fmt_=None,bold_=False,fc_="000000"):
            try:
                c=ws2.cell(row=r,column=col,value=v)
                c.font=Font(name="Calibri",size=8,bold=bold_,color=fc_)
                c.fill=PatternFill("solid",fgColor=bg)
                c.alignment=Alignment(horizontal="right" if fmt_ and fmt_!="@" else "center",vertical="center")
                if fmt_ and fmt_!="@": c.number_format=fmt_
            except: pass
        mid=rd.get('id_cobranca')
        cv(1, int(mid) if pd.notna(mid) else "","0")
        dv=rd.get('data_vencimento')
        cv(2, dv.date() if pd.notna(dv) else "","@")
        cv(3,grp,"@"); cv(4,grp,"@"); cv(5,"","@")
        cv(6,int(rd.get('aging_dias',0)),fmt_i)
        cv(7,float(rd.get('valor_cobranca',0)),fmt_r)
        vp=float(rd.get('valor_pago',0)); cv(8,vp if vp>0 else "-",fmt_r if vp>0 else "@")
        cv(9,float(rd.get('contas_a_receber',0)),fmt_r,bold_=True)
        liq=rd.get('liq_contabil'); cv(10,float(liq) if pd.notna(liq) else "-",fmt_r if pd.notna(liq) else "@")
        cv(11,float(rd.get('v40',0)),fmt_r); cv(12,float(rd.get('v50',0)),fmt_r); cv(13,float(rd.get('v60',0)),fmt_r)
        cv(14,float(rd.get('vtot',0)),fmt_r,bold_=True)
        cv(15,GDESC.get(grp,grp),"@")
        cust=rd.get('custo_aq'); cv(16,float(cust) if pd.notna(cust) and cust>0 else "-",fmt_r if pd.notna(cust) and cust>0 else "@")
        cv(17,int(rd.get('aging_compra',0)),fmt_i)
        tje=float(rd.get('tje',0))
        try:
            c18=ws2.cell(row=r,column=18,value=tje)
            c18.font=Font(name="Calibri",size=8,bold=(tje>0),color="1E7145" if tje>0 else "595959")
            c18.fill=PatternFill("solid",fgColor=bg)
            c18.alignment=Alignment(horizontal="right",vertical="center")
            c18.border=tb(); c18.number_format='0.0%'
        except: pass
        row_w+=1

    buf=io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


# ── INTERFACE: ABAS ─────────────────────────────────────────────────
tab1, tab2 = st.tabs(["🔄  Conciliação de Carteira", "📊  PDD Carteira C5"])

with tab1:

    with st.sidebar:
        st.markdown("## 📁 Histórico de Conciliações")
        tem_secrets = "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets
        if not tem_secrets:
            st.warning("Configure GITHUB_TOKEN e GITHUB_REPO nos Secrets para ativar o histórico.")
        else:
            if st.button("🔄 Atualizar histórico"):
                st.cache_data.clear()
            with st.spinner("Carregando..."):
                registros = carregar_indice()
            if not registros:
                st.info("Nenhuma conciliação salva ainda.")
            else:
                st.markdown(f"**{len(registros)} conciliação(ões) salva(s)**")
                st.markdown("---")
                for reg in registros:
                    st.markdown(f"""
                    <div class="hist-row">
                        <div class="hist-data">{reg.get('data','—')}</div>
                        <div class="hist-label">{reg.get('label_ant','—')} × {reg.get('label_atu','—')}</div>
                        <div class="hist-badge">{reg.get('total',0):,} registros</div>
                    </div>""", unsafe_allow_html=True)
                    hist_lista = listar_historico()
                    match = next((h for h in hist_lista if h["nome"] == reg.get("arquivo")), None)
                    if match:
                        dados = baixar_excel_github(match["url_download"])
                        if dados:
                            st.download_button(label="⬇️ Baixar", data=dados,
                                file_name=reg.get("arquivo","conciliacao.xlsx"),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{reg.get('arquivo')}")
                    st.markdown("---")

    st.markdown("""
    <div class="hero">
        <h1>🔄 Conciliação de Carteira</h1>
        <p>Direitos Creditórios — cruzamento automático entre dois meses</p>
    </div>""", unsafe_allow_html=True)

    for key in ["resultado","conc","df_ant","df_atu","label_ant","label_atu","nome_saida"]:
        if key not in st.session_state:
            st.session_state[key] = None

    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="card"><h3>📥 Mês Anterior</h3><p>Base de comparação</p></div>', unsafe_allow_html=True)
        f_ant = st.file_uploader("", type=["xlsx"], key="ant", label_visibility="collapsed")
        if f_ant: st.success(f"✓ {f_ant.name}")
    with col2:
        st.markdown('<div class="card"><h3>📥 Mês Atual</h3><p>Nova posição da carteira</p></div>', unsafe_allow_html=True)
        f_atu = st.file_uploader("", type=["xlsx"], key="atu", label_visibility="collapsed")
        if f_atu: st.success(f"✓ {f_atu.name}")

    st.markdown("<br>", unsafe_allow_html=True)
    col_btn, _ = st.columns([1, 2])
    with col_btn:
        executar = st.button("🔄  Executar Conciliação", disabled=not (f_ant and f_atu))

    if executar and f_ant and f_atu:
        with st.spinner("Lendo arquivos e cruzando bases..."):
            try:
                bytes_ant=f_ant.read(); bytes_atu=f_atu.read()
                df_ant,_=ler_base(bytes_ant,f_ant.name); df_atu,_=ler_base(bytes_atu,f_atu.name)
                label_ant=Path(f_ant.name).stem; label_atu=Path(f_atu.name).stem
                conc=conciliar(df_ant,df_atu)
                excel_bytes=gerar_excel_bytes_conc(conc,df_ant,df_atu,label_ant,label_atu)
                nome_saida=f"Conciliacao_{label_atu}_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                st.session_state.resultado=excel_bytes; st.session_state.conc=conc
                st.session_state.df_ant=df_ant; st.session_state.df_atu=df_atu
                st.session_state.label_ant=label_ant; st.session_state.label_atu=label_atu
                st.session_state.nome_saida=nome_saida
                if tem_secrets:
                    with st.spinner("Salvando no GitHub..."):
                        resumo={"arquivo":nome_saida,"data":datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "label_ant":label_ant,"label_atu":label_atu,"total":len(conc),
                                "novos":int((conc["tipo_movimentacao"]=="🟢 NOVO").sum()),
                                "saiu":int((conc["tipo_movimentacao"]=="🔴 SAIU DA CARTEIRA").sum()),
                                "alterados":int((conc["tipo_movimentacao"]=="🟡 ALTERADO").sum()),
                                "sem_mov":int((conc["tipo_movimentacao"]=="⚪ SEM MOVIMENTAÇÃO").sum()),
                                "car_ant":round(float(df_ant["contas_a_receber"].sum()),2),
                                "car_atu":round(float(df_atu["contas_a_receber"].sum()),2)}
                        ok,msg=salvar_no_github(nome_saida,excel_bytes,resumo)
                        if ok: st.success("✅ Excel salvo no GitHub!")
                        else: st.warning(f"⚠️ Não foi possível salvar no GitHub: {msg}")
            except Exception as e:
                st.error(f"❌ Erro: {e}")

    if st.session_state.resultado:
        conc=st.session_state.conc; df_ant=st.session_state.df_ant; df_atu=st.session_state.df_atu
        label_ant=st.session_state.label_ant; label_atu=st.session_state.label_atu
        n_nov=(conc["tipo_movimentacao"]=="🟢 NOVO").sum()
        n_sai=(conc["tipo_movimentacao"]=="🔴 SAIU DA CARTEIRA").sum()
        n_alt=(conc["tipo_movimentacao"]=="🟡 ALTERADO").sum()
        n_sem=(conc["tipo_movimentacao"]=="⚪ SEM MOVIMENTAÇÃO").sum()
        st.markdown("---"); st.markdown("### ✅ Conciliação concluída")
        st.markdown(f"""
        <div class="stat-grid">
            <div class="stat verde"><div class="num">{n_nov:,}</div><div class="lbl">🟢 Novos</div></div>
            <div class="stat verm"><div class="num">{n_sai:,}</div><div class="lbl">🔴 Saíram da Carteira</div></div>
            <div class="stat laran"><div class="num">{n_alt:,}</div><div class="lbl">🟡 Alterados</div></div>
            <div class="stat cinza"><div class="num">{n_sem:,}</div><div class="lbl">⚪ Sem Movimentação</div></div>
        </div>""", unsafe_allow_html=True)
        car_ant=df_ant["contas_a_receber"].sum(); car_atu=df_atu["contas_a_receber"].sum()
        var_car=car_atu-car_ant; sinal="pos" if var_car>=0 else "neg"
        st.markdown(f"""
        <div class="financeiro">
            <div class="fin-item"><div class="fin-lbl">contas_a_receber — {label_ant}</div><div class="fin-val">R$ {car_ant:,.2f}</div></div>
            <div class="fin-item"><div class="fin-lbl">contas_a_receber — {label_atu}</div><div class="fin-val">R$ {car_atu:,.2f}</div></div>
            <div class="fin-item"><div class="fin-lbl">Variação contas_a_receber</div><div class="fin-val {sinal}">R$ {var_car:+,.2f}</div></div>
        </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        col_dl,_=st.columns([1,2])
        with col_dl:
            st.download_button(label="⬇️  Baixar Excel com resultado",data=st.session_state.resultado,
                file_name=st.session_state.nome_saida,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.markdown("---"); st.markdown("### 📊 Distribuição por Status — Mês Atual")
        df_status=df_atu.copy()
        df_status["status_pagamento"]=df_status["status_pagamento"].replace("","(sem status — em aberto)")
        contagem=df_status["status_pagamento"].value_counts().reset_index()
        contagem.columns=["Status","Quantidade"]; total_st=contagem["Quantidade"].sum()
        contagem["%"]=(contagem["Quantidade"]/total_st*100).round(1).astype(str)+"%"
        CORES_STATUS={"liquidado":"#1A7F3C","pagamento_parcial":"#E65100","original_liquidado":"#1565C0","(sem status — em aberto)":"#6B7A99"}
        contagem["cor"]=contagem["Status"].map(lambda x: CORES_STATUS.get(x,"#9E9E9E"))
        col_graf,col_tab=st.columns([3,2])
        with col_graf:
            fig=go.Figure(go.Bar(x=contagem["Status"],y=contagem["Quantidade"],marker_color=contagem["cor"].tolist(),
                text=contagem["Quantidade"].apply(lambda v:f"{v:,}"),textposition="outside",textfont=dict(size=13,family="Inter, sans-serif")))
            fig.update_layout(plot_bgcolor="white",paper_bgcolor="white",margin=dict(t=20,b=40,l=20,r=20),height=360,
                xaxis=dict(tickfont=dict(size=12,family="Inter"),showgrid=False,linecolor="#E8EDF5"),
                yaxis=dict(tickfont=dict(size=11,family="Inter"),gridcolor="#F0F4FA",zeroline=False),bargap=0.35)
            st.plotly_chart(fig,use_container_width=True)
        with col_tab:
            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(contagem[["Status","Quantidade","%"]],use_container_width=True,hide_index=True,height=300)
        st.markdown("---"); st.markdown("### 🔍 Explorar Cobranças Alteradas")
        conc_alt=conc[conc["tipo_movimentacao"]=="🟡 ALTERADO"].copy()
        if conc_alt.empty:
            st.info("Nenhuma cobrança alterada nesta conciliação.")
        else:
            todas_opcoes=sorted(set(campo.strip() for val in conc_alt["campos_alterados"]
                for campo in val.split("|") if campo.strip() and campo.strip()!="—"))
            col_f1,col_f2,col_f3=st.columns([2,2,1])
            with col_f1:
                campos_sel=st.multiselect("Filtrar por campo alterado",options=todas_opcoes,default=[],placeholder="Todos os campos...")
            with col_f2:
                status_opcoes=sorted(conc_alt["status_pagamento_atu"].unique().tolist())
                status_sel=st.multiselect("Filtrar por status atual",options=status_opcoes,default=[],placeholder="Todos os status...")
            with col_f3:
                st.markdown("<br>", unsafe_allow_html=True)
                var_neg=st.checkbox("Só variação negativa em CAR",value=False)
            df_filtrado=conc_alt.copy()
            if campos_sel: df_filtrado=df_filtrado[df_filtrado["campos_alterados"].apply(lambda x: any(c in x for c in campos_sel))]
            if status_sel: df_filtrado=df_filtrado[df_filtrado["status_pagamento_atu"].isin(status_sel)]
            if var_neg: df_filtrado=df_filtrado[df_filtrado["var_contas_a_receber"]<0]
            st.markdown(f"<p style='font-size:13px;color:#6B7A99;margin:8px 0 12px;'>Exibindo <b>{len(df_filtrado):,}</b> de <b>{len(conc_alt):,}</b> cobranças alteradas</p>",unsafe_allow_html=True)
            COLS_EX=["id_cobranca","campos_alterados","status_pagamento_ant","status_pagamento_atu",
                     "valor_pago_ant","valor_pago_atu","var_valor_pago","valor_excedente_pago_ant",
                     "valor_excedente_pago_atu","var_excedente_pago","contas_a_receber_ant","contas_a_receber_atu","var_contas_a_receber"]
            RENAME={"id_cobranca":"ID Cobrança","campos_alterados":"Campos Alterados",
                    "status_pagamento_ant":f"Status ({label_ant})","status_pagamento_atu":f"Status ({label_atu})",
                    "valor_pago_ant":"Valor Pago Ant.","valor_pago_atu":"Valor Pago Atu.","var_valor_pago":"Var. Valor Pago",
                    "valor_excedente_pago_ant":"Excedente Ant.","valor_excedente_pago_atu":"Excedente Atu.","var_excedente_pago":"Var. Excedente",
                    "contas_a_receber_ant":"CAR Ant.","contas_a_receber_atu":"CAR Atu.","var_contas_a_receber":"Var. CAR"}
            df_exibir=df_filtrado[COLS_EX].rename(columns=RENAME)
            str_cols=["ID Cobrança","Campos Alterados",f"Status ({label_ant})",f"Status ({label_atu})"]
            df_fmt=df_exibir.copy()
            for c in df_fmt.columns:
                if c not in str_cols: df_fmt[c]=df_fmt[c].apply(lambda v: f"R$ {v:,.2f}" if pd.notna(v) else "")
            st.dataframe(df_fmt,use_container_width=True,hide_index=True,height=420)

    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown("<p style='text-align:center;font-size:12px;color:#A0AEC0;'>Conciliação de Carteira — Direitos Creditórios &nbsp;|&nbsp; Chave: <code>id_cobranca</code></p>",unsafe_allow_html=True)

with tab2:
    # ══════════════════════════════════════════════════════════════════════
    # GITHUB — HISTÓRICO PDD
    # ══════════════════════════════════════════════════════════════════════

    import base64, json, requests as _requests

    def _gh_headers():
        token = st.secrets.get("GITHUB_TOKEN", "")
        return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

    def _gh_repo():
        return st.secrets.get("GITHUB_REPO", "")

    def _tem_secrets():
        return "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets

    def salvar_pdd_github(nome_arquivo, excel_bytes, resumo_json):
        """Salva Excel em historico_pdd/ e atualiza indice_pdd.json."""
        repo = _gh_repo(); hdrs = _gh_headers()
        conteudo_b64 = base64.b64encode(excel_bytes).decode()

        # 1. Salvar Excel
        url_xlsx = f"https://api.github.com/repos/{repo}/contents/historico_pdd/{nome_arquivo}"
        r_check  = _requests.get(url_xlsx, headers=hdrs, timeout=10)
        payload  = {"message": f"PDD: {nome_arquivo}", "content": conteudo_b64}
        if r_check.status_code == 200:
            payload["sha"] = r_check.json()["sha"]
        r = _requests.put(url_xlsx, headers=hdrs, json=payload, timeout=60)
        if r.status_code not in (200, 201):
            return False, f"Erro ao salvar Excel: {r.status_code} — {r.text[:200]}"

        # 2. Atualizar índice JSON
        idx_url = f"https://api.github.com/repos/{repo}/contents/historico_pdd/indice_pdd.json"
        r_idx   = _requests.get(idx_url, headers=hdrs, timeout=10)
        registros = []; idx_sha = None
        if r_idx.status_code == 200:
            try:
                registros = json.loads(base64.b64decode(r_idx.json()["content"]).decode())
                idx_sha   = r_idx.json()["sha"]
            except Exception:
                registros = []
        registros.insert(0, resumo_json)
        registros = registros[:50]
        payload_idx = {
            "message": f"Índice PDD atualizado: {nome_arquivo}",
            "content": base64.b64encode(
                json.dumps(registros, ensure_ascii=False, indent=2).encode()
            ).decode(),
        }
        if idx_sha:
            payload_idx["sha"] = idx_sha
        r2 = _requests.put(idx_url, headers=hdrs, json=payload_idx, timeout=15)
        if r2.status_code not in (200, 201):
            return False, f"Erro ao atualizar índice: {r2.status_code}"
        return True, "ok"

    def carregar_indice_pdd():
        url = f"https://api.github.com/repos/{_gh_repo()}/contents/historico_pdd/indice_pdd.json"
        r   = _requests.get(url, headers=_gh_headers(), timeout=10)
        if r.status_code != 200:
            return []
        try:
            return json.loads(base64.b64decode(r.json()["content"]).decode())
        except Exception:
            return []

    def listar_historico_pdd():
        url = f"https://api.github.com/repos/{_gh_repo()}/contents/historico_pdd"
        r   = _requests.get(url, headers=_gh_headers(), timeout=10)
        if r.status_code != 200:
            return []
        return sorted(
            [{"nome": a["name"], "url_download": a["download_url"], "sha": a["sha"]}
             for a in r.json() if a["name"].endswith(".xlsx")],
            key=lambda x: x["nome"], reverse=True,
        )

    def baixar_excel_pdd(url_download):
        r = _requests.get(url_download, timeout=60)
        return r.content if r.status_code == 200 else None

    # ══════════════════════════════════════════════════════════════════════
    # INTERFACE STREAMLIT
    # ══════════════════════════════════════════════════════════════════════

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .main-header {
        background: linear-gradient(135deg, #1F2D5A 0%, #2E4189 100%);
        color: white; padding: 1.5rem 2rem; border-radius: 12px;
        margin-bottom: 1.5rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 1.5rem; }
    .main-header p  { color: #D6DCF0; margin: 0.3rem 0 0; font-size: 0.85rem; }
    .hist-row {
        display: flex; align-items: center; gap: 10px;
        padding: 10px 14px; background: white; border-radius: 8px;
        border: 1px solid #E8EDF5; margin-bottom: 6px;
    }
    .hist-data  { font-size: 11px; color: #8898AA; min-width: 75px; }
    .hist-label { font-size: 13px; font-weight: 700; color: #1F2D5A; flex: 1; }
    .hist-val   { font-size: 11px; color: #2E4189; font-weight: 600;
                  background: #EEF3FB; padding: 2px 8px; border-radius: 12px; }
    </style>
    """, unsafe_allow_html=True)

    # ── Sidebar: histórico ────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## 📁 Histórico PDD C5")
        tem_gh = _tem_secrets()

        if not tem_gh:
            st.warning("Configure GITHUB_TOKEN e GITHUB_REPO nos Secrets para ativar o histórico.")
        else:
            if st.button("🔄 Atualizar", key="refresh_pdd"):
                st.cache_data.clear()
            with st.spinner("Carregando histórico..."):
                registros = carregar_indice_pdd()
                hist_lista = listar_historico_pdd()

            if not registros:
                st.info("Nenhum cálculo salvo ainda.")
            else:
                st.markdown(f"**{len(registros)} cálculo(s) salvo(s)**")
                st.markdown("---")
                for reg in registros:
                    car  = reg.get('car', 0)
                    pdd  = reg.get('pdd', 0)
                    perc = reg.get('pct_ead', 0)
                    st.markdown(f"""
                    <div class="hist-row">
                        <div class="hist-data">{reg.get('data_base','—')}</div>
                        <div class="hist-label">{reg.get('data_base','—')}</div>
                        <div class="hist-val">{perc:.1f}% EAD</div>
                    </div>
                    <div style="font-size:11px;color:#6B7A99;padding:0 14px 6px;">
                        CAR: R$ {car:,.0f} &nbsp;|&nbsp; PDD: R$ {pdd:,.0f}
                    </div>
                    """, unsafe_allow_html=True)
                    match = next((h for h in hist_lista if h["nome"] == reg.get("arquivo")), None)
                    if match:
                        dados = baixar_excel_pdd(match["url_download"])
                        if dados:
                            st.download_button(
                                label="⬇️ Baixar Excel",
                                data=dados,
                                file_name=reg.get("arquivo", "pdd.xlsx"),
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_pdd_{reg.get('arquivo')}",
                            )
                    st.markdown("---")

    # ── Cabeçalho ─────────────────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
      <h1>📊 PDD Carteira C5 — SCD</h1>
      <p>BCB 352 / Resolução BCB nº 4.966/2021 &nbsp;|&nbsp; Cálculo Mensal Automatizado</p>
    </div>
    """, unsafe_allow_html=True)

    # ── Data-base ──────────────────────────────────────────────────────────
    data_base = st.date_input(
        "📅 Data-base do fechamento",
        value=datetime.today().replace(day=1),
        help="Último dia do mês de referência"
    )

    st.divider()

    # ── Upload CAR ────────────────────────────────────────────────────────
    st.subheader("1️⃣  Base CAR — Contas a Receber")
    col1, col2 = st.columns([2, 1])
    with col1:
        car_file = st.file_uploader(
            "Arquivo CAR do mês (.xlsx)",
            type=["xlsx", "xls"],
            help="Colunas: id_cobranca, valor_cobranca, contas_a_receber, aging_dias"
        )
    with col2:
        car_sheet = st.text_input("Nome da aba (em branco = primeira)", placeholder="ex: DADOS 31-07")

    if car_file:
        xl = pd.ExcelFile(car_file)
        if not car_sheet:
            car_sheet = xl.sheet_names[0]
        st.caption(f"✅ Arquivo carregado | Aba: **{car_sheet}** | Abas: {xl.sheet_names}")

    st.divider()

    # ── Cessões ───────────────────────────────────────────────────────────
    st.subheader("2️⃣  Planilhas de Cessão (acumuladas)")
    st.info(
        "Inclua **todas** as cessões desde o início da carteira.  \n"
        "**Tipo padrão** = `ID COBRANÇA / VALOR FACE BOLETO / CESSÃO (data)`  \n"
        "**Tipo nova** = `ID / VALOR FACE DIREITO ECONÔMICO / DIAS EM ATRASO`",
        icon="ℹ️"
    )

    if "cessoes" not in st.session_state:
        st.session_state["cessoes"] = []

    def add_cessao():
        st.session_state["cessoes"].append({"file": None, "sheet": "", "tipo": "padrao", "data_cessao": ""})

    def remove_cessao(i):
        st.session_state["cessoes"].pop(i)

    st.button("➕ Adicionar cessão", on_click=add_cessao)

    for i, cess in enumerate(st.session_state["cessoes"]):
        with st.container(border=True):
            cols = st.columns([3, 1, 1, 2, 0.5])
            with cols[0]:
                f = st.file_uploader(f"Cessão #{i+1}", type=["xlsx","xls"], key=f"cess_file_{i}")
                if f: st.session_state["cessoes"][i]["file"] = f
            with cols[1]:
                t = st.selectbox("Tipo", ["padrao","nova"], key=f"cess_tipo_{i}")
                st.session_state["cessoes"][i]["tipo"] = t
            with cols[2]:
                sh = st.text_input("Aba", key=f"cess_sheet_{i}", placeholder="1ª aba")
                st.session_state["cessoes"][i]["sheet"] = sh or None
            with cols[3]:
                if t == "nova":
                    dc = st.text_input("Data cessão (DD/MM/AAAA)", key=f"cess_data_{i}", placeholder="ex: 16/07/2026")
                    st.session_state["cessoes"][i]["data_cessao"] = dc
                else:
                    st.caption("Data lida da coluna CESSÃO")
            with cols[4]:
                st.button("🗑️", key=f"del_{i}", on_click=remove_cessao, args=(i,))

    st.divider()

    # ── Botão calcular ─────────────────────────────────────────────────────
    calcular = st.button("🚀  Calcular PDD", type="primary", use_container_width=True)

    if calcular:
        erros = []
        if not car_file:       erros.append("Faça o upload do arquivo CAR.")
        if not st.session_state["cessoes"]: erros.append("Adicione ao menos uma cessão.")
        for i, c in enumerate(st.session_state["cessoes"]):
            if not c.get("file"): erros.append(f"Cessão #{i+1}: nenhum arquivo carregado.")

        if erros:
            for e in erros: st.error(e)
        else:
            with st.spinner("Calculando PDD..."):
                try:
                    # Ler CAR
                    car_file.seek(0)
                    df_car = ler_car(car_file, car_sheet or None)

                    # Ler cessões
                    cessoes_dfs = []
                    for i, c in enumerate(st.session_state["cessoes"]):
                        c["file"].seek(0)
                        df_c = ler_cessao(
                            buf=c["file"], sheet=c.get("sheet"),
                            tipo=c.get("tipo","padrao"),
                            data_cessao_override=c.get("data_cessao") or None,
                        )
                        cessoes_dfs.append(df_c)

                    df_cessoes = pd.concat(cessoes_dfs, ignore_index=True).drop_duplicates('id').copy()

                    # Calcular
                    data_base_dt = datetime(data_base.year, data_base.month, data_base.day)
                    df_res = calcular_pdd(df_car, df_cessoes)

                    S40   = round(df_res.v40.sum(), 2)
                    S50   = round(df_res.v50.sum(), 2)
                    S60   = round(df_res.v60.sum(), 2)
                    STOT  = round(df_res.vtot.sum(), 2)
                    TCAR  = round(df_res.contas_a_receber.sum(), 2)
                    TEAD  = round(df_res.valor_cobranca.sum(), 2)
                    PCT   = round(STOT / TEAD * 100, 1) if TEAD else 0

                    # Gerar Excel
                    excel_bytes = gerar_excel_bytes_pdd(df_res, data_base_dt)
                    data_str    = data_base_dt.strftime('%d%m%Y')
                    nome_saida  = f"PDD_C5_COSIF_{data_str}_BCB352.xlsx"

                    # Salvar no session_state
                    st.session_state["pdd_resultado"]   = excel_bytes
                    st.session_state["pdd_df"]          = df_res
                    st.session_state["pdd_totais"]      = {"S40":S40,"S50":S50,"S60":S60,"STOT":STOT,"TCAR":TCAR,"TEAD":TEAD,"PCT":PCT}
                    st.session_state["pdd_nome_saida"]  = nome_saida
                    st.session_state["pdd_data_base"]   = data_base_dt

                    # Salvar no GitHub
                    if tem_gh:
                        with st.spinner("Salvando histórico no GitHub..."):
                            resumo = {
                                "arquivo":   nome_saida,
                                "data_base": data_base_dt.strftime("%d/%m/%Y"),
                                "data":      datetime.now().strftime("%d/%m/%Y %H:%M"),
                                "qtd":       len(df_res),
                                "car":       TCAR,
                                "ead":       TEAD,
                                "pdd":       STOT,
                                "pct_ead":   PCT,
                                "v40":       S40,
                                "v50":       S50,
                                "v60":       S60,
                            }
                            ok, msg = salvar_pdd_github(nome_saida, excel_bytes, resumo)
                            if ok:
                                st.success("✅ Excel salvo no histórico (GitHub)!")
                            else:
                                st.warning(f"⚠️ Não foi possível salvar no GitHub: {msg}")

                except Exception as e:
                    st.error(f"❌ Erro durante o cálculo: {e}")
                    st.exception(e)

    # ── Resultado ─────────────────────────────────────────────────────────
    if st.session_state.get("pdd_resultado"):
        t    = st.session_state["pdd_totais"]
        df   = st.session_state["pdd_df"]

        st.markdown("---")
        st.subheader("📋 Resultado")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Cobranças", f"{len(df):,}")
        m2.metric("CAR Total", f"R$ {t['TCAR']:,.0f}")
        m3.metric("EAD (Face)", f"R$ {t['TEAD']:,.0f}")
        m4.metric("PDD Total", f"R$ {t['STOT']:,.0f}", delta=f"{t['PCT']:.1f}% EAD")

        c1, c2, c3 = st.columns(3)
        c1.metric("40-4 Perda Incorrida",      f"R$ {t['S40']:,.2f}")
        c2.metric("50-7 Prov. Adicional (3,4%)", f"R$ {t['S50']:,.2f}")
        c3.metric("60-0 Perda Esperada",        f"R$ {t['S60']:,.2f}")

        st.markdown("**Resumo por faixa:**")
        grp_view = df.groupby('grupo').agg(
            Qtd=('id_cobranca','count'),
            CAR=('contas_a_receber','sum'),
            PDD=('vtot','sum'),
        ).round(2)
        grp_view['% PDD/CAR'] = (grp_view['PDD'] / grp_view['CAR'] * 100).round(1)
        st.dataframe(
            grp_view.style.format({'CAR':'R$ {:,.2f}','PDD':'R$ {:,.2f}','% PDD/CAR':'{:.1f}%'}),
            use_container_width=True,
        )

        st.markdown("---")
        st.download_button(
            label="📥  Baixar Excel (Premissas + COSIF + Composição)",
            data=st.session_state["pdd_resultado"],
            file_name=st.session_state["pdd_nome_saida"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
        st.caption(
            f"Arquivo: **{st.session_state['pdd_nome_saida']}** | "
            "3 abas: Premissas e Metodologia | PDD C5 — COSIF Completo | Composição por Cobrança"
        )

    # ── Rodapé ────────────────────────────────────────────────────────────
    st.markdown("---")
    with st.expander("ℹ️ Premissas e regras aplicadas"):
        st.markdown(f"""
    **Normas:** BCB 352 (Circular 3.547/2011) / Resolução BCB nº 4.966/2021

    | Estágio | Critério | Base PDD | Conta | % |
    |---|---|---|---|---|
    | E2 — 31-60d | Aging ≤ 60d | CAR × 15% | 60-0 | 15% |
    | E2 — 61-90d | Aging ≤ 90d | CAR × 38% | 60-0 | 38% |
    | E3 — >90d | Inadimplido | CAR × Anexo II + 3,4% | 40-4 + 50-7 | 50% → 100% |
    | C1 POCI | Aging >630d | 100% liq.contábil | 40-4 | 100% |
    | C4 POCI | Aging 91-630d | CAR × Anexo II ≤ liq. | 40-4 + 50-7 | 50% → 100% |

    **TJE (CADOC 3040):** calculada **uma única vez na data de aquisição** — permanece fixa.
    - E2: `TJE = (1/(1-deságio%))^(365/450) - 1`
    - E3/POCI: `ratio = (1-PECLD%)/(1-deságio%); TJE = ratio^(365/1260) - 1`
        """)
