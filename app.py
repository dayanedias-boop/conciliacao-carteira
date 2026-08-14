"""
Página inicial — Menu de navegação
"""
import streamlit as st

st.set_page_config(
    page_title="Carteira DC — SCD",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.hero {
    background: linear-gradient(135deg, #1F3864 0%, #2F5496 100%);
    border-radius: 16px; padding: 48px 56px; color: white;
    margin-bottom: 32px; text-align: center;
}
.hero h1 { font-size: 32px; font-weight: 700; margin: 0 0 10px 0; }
.hero p  { font-size: 15px; opacity: 0.8; margin: 0; }
.card-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; max-width: 820px; margin: 0 auto; }
.card {
    background: white; border-radius: 14px; padding: 28px 32px;
    border: 1px solid #E8EDF5; cursor: pointer;
    box-shadow: 0 2px 8px rgba(31,56,100,0.08);
    transition: box-shadow 0.2s;
}
.card:hover { box-shadow: 0 6px 20px rgba(31,56,100,0.15); }
.card .icon { font-size: 36px; margin-bottom: 12px; }
.card h2 { font-size: 17px; font-weight: 700; color: #1F3864; margin: 0 0 8px 0; }
.card p  { font-size: 13px; color: #6B7A99; margin: 0; line-height: 1.5; }
.badge {
    display: inline-block; font-size: 11px; font-weight: 600;
    padding: 3px 10px; border-radius: 20px; margin-top: 12px;
}
.badge-blue  { background: #EEF3FB; color: #2F5496; }
.badge-green { background: #E2EFDA; color: #1A7F3C; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>🏦 Carteira de Direitos Creditórios — SCD</h1>
  <p>Selecione o módulo desejado no menu lateral ou clique nos cards abaixo</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="card-grid">

  <div class="card">
    <div class="icon">🔄</div>
    <h2>Conciliação de Carteira</h2>
    <p>Cruzamento automático entre dois meses — identifica cobranças novas,
       liquidadas, alteradas e sem movimentação.</p>
    <span class="badge badge-blue">Chave: id_cobranca</span>
  </div>

  <div class="card">
    <div class="icon">📊</div>
    <h2>PDD Carteira C5</h2>
    <p>Cálculo mensal de PDD e TJE conforme BCB 352 e Resolução BCB nº 4.966/2021.
       Gera o arquivo COSIF completo com Premissas + Composição por Cobrança.</p>
    <span class="badge badge-green">BCB 352 / Res. 4.966/2021</span>
  </div>

</div>
""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)
st.info("👈 Use o menu lateral para navegar entre os módulos.", icon="ℹ️")

st.markdown("---")
st.markdown(
    "<p style='text-align:center;font-size:12px;color:#A0AEC0;'>"
    "Carteira DC — SCD &nbsp;|&nbsp; Conciliação + PDD C5 &nbsp;|&nbsp; BCB 352 / Res. 4.966/2021"
    "</p>",
    unsafe_allow_html=True,
)
