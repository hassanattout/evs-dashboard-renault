import streamlit as st
import pandas as pd
from utils import apply_global_style, load_data

# Configuration de la page
st.set_page_config(
    page_title="EVS Dashboard | Ponts Roulants",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Application du style global (CSS)
apply_global_style()

# --- CHARGEMENT DES DONNÉES ---
df = load_data()

# --- CALCULS DES KPIs ---
equipements = len(df)
pays = df["pays"].nunique() if "pays" in df.columns else 0
sites = df["site"].nunique() if "site" in df.columns else 0

# --- EXTRACTION DYNAMIQUE DE L'HORIZON BUDGÉTAIRE ---
# On cherche les colonnes qui commencent par une année (ex: 2026 OPEX)
years_list = [
    int(str(col)[:4])
    for col in df.columns
    if str(col)[:4].isdigit()
    and 2025 <= int(str(col)[:4]) <= 2100
]

# Définition de la variable max_year (utilisée dans le texte et les KPIs)
max_year = max(years_list) if years_list else 2030

# --- STYLE CSS PERSONNALISÉ ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@500;700&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'IBM Plex Sans', sans-serif;
    }

    [data-testid="stSidebar"] {
        background: #0b0d13;
        border-right: 1px solid #242635;
    }

    [data-testid="stSidebar"] * {
        color: #d7ddf2 !important;
    }

    .block-container {
        padding-top: 3rem;
        padding-left: 3rem;
        padding-right: 3rem;
        max-width: none;
        width: 100%;
    }

    h1 {
        font-family: 'IBM Plex Mono', monospace !important;
        color: #e8ecff !important;
        font-size: 2.6rem !important;
        font-weight: 700 !important;
        letter-spacing: -0.04em;
    }

    h2 {
        color: #e8ecff !important;
        font-size: 1.35rem !important;
        font-weight: 700 !important;
    }

    .hero-box {
        background: linear-gradient(135deg, #141824 0%, #0f121a 100%);
        border: 1px solid #2a2e3f;
        border-radius: 22px;
        padding: 34px 38px;
        margin-top: 26px;
        margin-bottom: 28px;
        box-shadow: 0 18px 50px rgba(0,0,0,0.30);
    }

    .hero-subtitle {
        color: #aeb7cf;
        font-size: 1.05rem;
        line-height: 1.7;
        max-width: 980px;
    }

    .kpi-card {
        background: linear-gradient(180deg, #171b27 0%, #10131c 100%);
        border: 1px solid #31384d;
        border-radius: 18px;
        padding: 28px 26px;
        box-shadow: 0 18px 40px rgba(0,0,0,0.28);
        height: 100%;
    }

    .kpi-value {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 2.35rem;
        font-weight: 700;
        color: #9cc7ff;
        line-height: 1;
    }

    .kpi-label {
        font-size: 0.78rem;
        color: #9aa3b8;
        text-transform: uppercase;
        letter-spacing: 0.08em;
        margin-top: 10px;
    }

    .objective-box {
        background: #101723;
        border-left: 4px solid #89b4fa;
        border-radius: 14px;
        padding: 18px 22px;
        color: #cdd6f4;
        margin-top: 28px;
        margin-bottom: 26px;
    }
</style>
""", unsafe_allow_html=True)

# --- CONTENU DE LA PAGE ---
st.markdown("# 🏗️ EVS Dashboard")
st.markdown("### Ponts roulants | Digitalisation de l'Évaluation Spéciale")

st.markdown(f"""
<div class="hero-box">
    <div class="hero-subtitle">
        Outil de pilotage du parc de ponts roulants, combinant vision globale,
        priorisation EVS / CAPEX, fiche équipement et calcul de fatigue des mécanismes.
        <br><br>
        L’horizon <b>2025–{max_year}</b> correspond à une période de planification budgétaire,
        avec une projection annuelle des besoins CAPEX et OPEX.
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="objective-box">
    <b>Objectif :</b> identifier les équipements à traiter en priorité, sécuriser la planification EVS / CAPEX
    et transformer les données du parc en outil d'aide à la décision industrielle.
</div>
""", unsafe_allow_html=True)

st.divider()

st.markdown("## Synthèse exécutive")

# Affichage des KPIs sur deux lignes
col1, col2 = st.columns(2)

with col1:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-value">{equipements}</div>
        <div class="kpi-label">Équipements analysés</div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-value">{pays}</div>
        <div class="kpi-label">Pays couverts</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

col3, col4 = st.columns(2)

with col3:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-value">{sites}</div>
        <div class="kpi-label">Sites industriels</div>
    </div>
    """, unsafe_allow_html=True)

with col4:
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-value">2025–{max_year}</div>
        <div class="kpi-label">Horizon budgétaire</div>
    </div>
    """, unsafe_allow_html=True)