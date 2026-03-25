import streamlit as st

# ──────────────────────────────────────────────
# CONFIGURAÇÃO DA PÁGINA
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="Análise de Dados — Zanattex",
    page_icon="🧵",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────
# CSS CUSTOMIZADO
# ──────────────────────────────────────────────
st.markdown("""
<style>
    footer { visibility: hidden; }
    .stApp { background-color: #0E1117; }

    /* Esconde o sidebar nessa página */
    section[data-testid="stSidebar"] { display: none; }
    [data-testid="collapsedControl"]  { display: none; }

    /* ── Cabeçalho ── */
    .hero-wrapper {
        text-align: center;
        padding: 52px 20px 36px 20px;
    }
    .hero-badge {
        display: inline-block;
        background: rgba(78,205,196,0.12);
        border: 1px solid rgba(78,205,196,0.35);
        color: #4ECDC4;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 2px;
        text-transform: uppercase;
        padding: 5px 16px;
        border-radius: 20px;
        margin-bottom: 18px;
    }
    .hero-title {
        color: #FFFFFF;
        font-size: 3rem;
        font-weight: 800;
        letter-spacing: -0.5px;
        margin: 0 0 10px 0;
        line-height: 1.1;
    }
    .hero-title span {
        background: linear-gradient(90deg, #4ECDC4, #45B7D1);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    .hero-subtitle {
        color: #8899AA;
        font-size: 1.1rem;
        margin: 0 auto;
        max-width: 520px;
        line-height: 1.6;
    }

    /* ── Divisor ── */
    .section-divider {
        border: none;
        border-top: 1px solid rgba(255,255,255,0.08);
        margin: 8px 0 36px 0;
    }

    /* ── Cards ── */
    .dash-card {
        background: linear-gradient(145deg, #13131A 0%, #1C1C26 60%, #22222E 100%);
        border: 1px solid rgba(255,255,255,0.09);
        border-radius: 18px;
        padding: 28px 26px 22px 26px;
        height: 100%;
        transition: border-color 0.3s, box-shadow 0.3s;
        position: relative;
        overflow: hidden;
    }
    .dash-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        border-radius: 18px 18px 0 0;
    }
    .card-teal::before   { background: linear-gradient(90deg, #4ECDC4, #45B7D1); }
    .card-orange::before { background: linear-gradient(90deg, #FFA726, #FF6B6B); }
    .card-blue::before   { background: linear-gradient(90deg, #45B7D1, #4ECDC4); }
    .card-purple::before { background: linear-gradient(90deg, #AB47BC, #7E57C2); }

    .dash-card:hover {
        border-color: rgba(78,205,196,0.35);
        box-shadow: 0 8px 32px rgba(78,205,196,0.1);
    }

    .card-icon {
        font-size: 2.4rem;
        margin-bottom: 12px;
        display: block;
    }
    .card-tag {
        display: inline-block;
        font-size: 0.66rem;
        font-weight: 700;
        letter-spacing: 1.5px;
        text-transform: uppercase;
        padding: 3px 10px;
        border-radius: 10px;
        margin-bottom: 10px;
    }
    .tag-teal   { background: rgba(78,205,196,0.14);  color: #4ECDC4; }
    .tag-orange { background: rgba(255,167,38,0.14);  color: #FFA726; }
    .tag-blue   { background: rgba(69,183,209,0.14);  color: #45B7D1; }
    .tag-purple { background: rgba(171,71,188,0.14);  color: #AB47BC; }

    .card-title {
        color: #F0F4F8;
        font-size: 1.2rem;
        font-weight: 700;
        margin: 0 0 8px 0;
        line-height: 1.3;
    }
    .card-desc {
        color: #7A8898;
        font-size: 0.88rem;
        line-height: 1.6;
        margin: 0 0 20px 0;
    }
    .card-meta {
        display: flex;
        gap: 14px;
        flex-wrap: wrap;
        margin-bottom: 18px;
    }
    .card-meta-item {
        display: flex;
        align-items: center;
        gap: 5px;
        color: #556070;
        font-size: 0.76rem;
        font-weight: 500;
    }
    .card-meta-item .dot {
        width: 6px; height: 6px;
        border-radius: 50%;
        background: #4ECDC4;
        flex-shrink: 0;
    }
    .dot-orange { background: #FFA726 !important; }
    .dot-blue   { background: #45B7D1 !important; }
    .dot-purple { background: #AB47BC !important; }

    /* ── Botão do card ── */
    .stButton > button {
        background: linear-gradient(135deg, #1C1C22, #28282E) !important;
        border: 1px solid rgba(255,255,255,0.12) !important;
        color: #FFFFFF !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-size: 0.88rem !important;
        padding: 10px 20px !important;
        width: 100% !important;
        transition: all 0.3s ease !important;
        letter-spacing: 0.3px !important;
    }
    .stButton > button:hover {
        border-color: #4ECDC4 !important;
        box-shadow: 0 0 18px rgba(78,205,196,0.25) !important;
        color: #4ECDC4 !important;
    }

    /* ── Rodapé ── */
    .page-footer {
        text-align: center;
        color: #3A4456;
        font-size: 0.78rem;
        padding: 36px 0 20px 0;
        letter-spacing: 0.3px;
    }
    .page-footer span {
        color: #4ECDC4;
    }

    /* ── Stats bar ── */
    .stats-bar {
        display: flex;
        justify-content: center;
        gap: 48px;
        padding: 16px 0 40px 0;
        flex-wrap: wrap;
    }
    .stat-item {
        text-align: center;
    }
    .stat-number {
        color: #FFFFFF;
        font-size: 1.8rem;
        font-weight: 800;
        line-height: 1;
        margin-bottom: 4px;
    }
    .stat-label {
        color: #4A5568;
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
</style>
""", unsafe_allow_html=True)

# ──────────────────────────────────────────────
# HERO / CABEÇALHO
# ──────────────────────────────────────────────
st.markdown("""
<div class="hero-wrapper">
    <div class="hero-badge">🧵 Zanattex · Plataforma de Analíse</div>
    <h1 class="hero-title">Análise de Dados <span>Zanattex</span></h1>
    <p class="hero-subtitle">
        Selecione um painel abaixo para acompanhar métricas de produção,
        corte e fechamento de containers em tempo real.
    </p>
</div>
""", unsafe_allow_html=True)

# Stats bar
st.markdown(""""""
<div class="stats-bar">
    <div class="stat-item">
        <div class="stat-number">4</div>
        <div class="stat-label">Painéis</div>
    </div>
    <div class="stat-item">
        <div class="stat-number">7</div>
        <div class="stat-label">Empresas</div>
<hr class="section-divider"/>
"""""",""" unsafe_allow_html=True)

# ──────────────────────────────────────────────
# CARDS DOS DASHBOARDS
# ──────────────────────────────────────────────
col1, col2 = st.columns(2, gap="large")

# ── Card 1 — Dashboard Produção Empresas ──
with col1:
    st.markdown("""
    <div class="dash-card card-teal">
        <span class="card-icon">🏭</span>
        <span class="card-tag tag-teal">Produção Geral</span>
        <p class="card-title">Dashboard de Produção — Empresas</p>
        <p class="card-desc">
            Visão consolidada de todas as empresas do grupo. Acompanhe
            produção diária por facção, evolução mensal, atingimento de metas
            e ranking por produto com treemap interativo.
        </p>
        <div class="card-meta">
            <span class="card-meta-item"><span class="dot"></span>Burdays · Camesa · Niazitex</span>
            <span class="card-meta-item"><span class="dot"></span>Cortex · Sultan · Decor</span>
            <span class="card-meta-item"><span class="dot"></span>Google Sheets · 10 min cache</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("▶  Acessar Produção — Empresas", key="btn_app"):
        st.switch_page("app.py")

# ── Card 2 — Dashboard Controle de Corte ──
with col2:
    st.markdown("""
    <div class="dash-card card-orange">
        <span class="card-icon">✂️</span>
        <span class="card-tag tag-orange">Corte</span>
        <p class="card-title">Dashboard Controle de Corte</p>
        <p class="card-desc">
            Monitoramento da produção no setor de corte por estação de trabalho
            (Máquina, Mesa 1, Mesa 2). Análise de metas diárias, evolução por
            operador e eficiência do setor.
        </p>
        <div class="card-meta">
            <span class="card-meta-item"><span class="dot dot-orange"></span>Máquina · Mesa 1 · Mesa 2</span>
            <span class="card-meta-item"><span class="dot dot-orange"></span>Meta diária configurável</span>
            <span class="card-meta-item"><span class="dot dot-orange"></span>Google Sheets · CSV</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("▶  Acessar Controle de Corte", key="btn_dashboard"):
        st.switch_page("dashboard.py")

st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)

col3, col4 = st.columns(2, gap="large")

# ── Card 3 — Dashboard Produção CAMESA ──
with col3:
    st.markdown("""
    <div class="dash-card card-blue">
        <span class="card-icon">📊</span>
        <span class="card-tag tag-blue">Análise Detalhada</span>
        <p class="card-title">Dashboard Produção — CAMESA</p>
        <p class="card-desc">
            Análise aprofundada da produção CAMESA com KPIs de atingimento,
            saldo vs meta, gráficos diários e breakdown por facção e produto.
            Inclui filtros por período e exportação de dados.
        </p>
        <div class="card-meta">
            <span class="card-meta-item"><span class="dot dot-blue"></span>Métricas de atingimento</span>
            <span class="card-meta-item"><span class="dot dot-blue"></span>Visão por facção e produto</span>
            <span class="card-meta-item"><span class="dot dot-blue"></span>Exportação CSV</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("▶  Acessar Análise CAMESA", key="btn_analiseprod"):
        st.switch_page("AnaliseProd.py")

# ── Card 4 — Programa de Fechamento de Containers ──
with col4:
    st.markdown("""
    <div class="dash-card card-purple">
        <span class="card-icon">📦</span>
        <span class="card-tag tag-purple">Logística</span>
        <p class="card-title">Fechamento de Containers</p>
        <p class="card-desc">
            Interface para preenchimento e geração de dashboards HTML de
            fechamento de containers. Registro de dados logísticos com
            exportação automática de relatório interativo.
        </p>
        <div class="card-meta">
            <span class="card-meta-item"><span class="dot dot-purple"></span>Geração de relatório HTML</span>
            <span class="card-meta-item"><span class="dot dot-purple"></span>Entrada de dados via formulário</span>
            <span class="card-meta-item"><span class="dot dot-purple"></span>Exportação automática</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
    if st.button("▶  Acessar Fechamento de Containers", key="btn_fechamento"):
        st.switch_page("programa_fechamento.py")

# ──────────────────────────────────────────────
# RODAPÉ
# ──────────────────────────────────────────────
st.markdown("""
<hr class="section-divider" style="margin-top:44px"/>
<div class="page-footer">
    <span>Zanattex</span> · Plataforma de Análise de Dados · Todos os direitos reservados
</div>
""", unsafe_allow_html=True)