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
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>🏦 Carteira de Direitos Creditórios — SCD</h1>
  <p>Selecione o módulo desejado no menu lateral</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    with st.container(border=True):
        st.markdown("### 🔄 Conciliação de Carteira")
        st.markdown(
            "Cruzamento automático entre dois meses — identifica cobranças "
            "novas, liquidadas, alteradas e sem movimentação."
        )
        st.caption("Chave: `id_cobranca`")
        if st.button("Abrir Conciliação →", key="btn_conc", use_container_width=True):
            st.switch_page("pages/1_Conciliacao.py")

with col2:
    with st.container(border=True):
        st.markdown("### 📊 PDD Carteira C5")
        st.markdown(
            "Cálculo mensal de PDD e TJE conforme BCB 352 e "
            "Resolução BCB nº 4.966/2021. Gera o arquivo COSIF completo."
        )
        st.caption("BCB 352 / Res. 4.966/2021")
        if st.button("Abrir PDD C5 →", key="btn_pdd", use_container_width=True):
            st.switch_page("pages/2_PDD_C5.py")

st.markdown("---")
st.markdown(
    "<p style='text-align:center;font-size:12px;color:#A0AEC0;'>"
    "Carteira DC — SCD &nbsp;|&nbsp; Conciliação + PDD C5 &nbsp;|&nbsp; BCB 352 / Res. 4.966/2021"
    "</p>",
    unsafe_allow_html=True,
)
