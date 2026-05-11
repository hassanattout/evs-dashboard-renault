import json
from pathlib import Path
import pandas as pd
import streamlit as st
import io
import requests

# --- Constants & Configuration ---
COLORS = {
    "bg": "#0d0f16",
    "card": "#141824",
    "border": "#2a2e3f",
    "text": "#e8ecff",
    "muted": "#9aa3b8",
    "blue": "#89b4fa",
    "green": "#a6e3a1",
    "yellow": "#f9e2af",
    "orange": "#fab387",
    "red": "#f38ba8",
    "mauve": "#cba6f7",
    "teal": "#94e2d5",
    "sky": "#89dceb",
    "grey": "#6c7086",
}

PLOTLY_LAYOUT = dict(
    plot_bgcolor="#141824",
    paper_bgcolor="#0d0f16",
    font=dict(color="#e8ecff", family="IBM Plex Sans, Arial, sans-serif", size=13),
    margin=dict(t=55, b=45, l=55, r=30),
)

# Configuration SharePoint
EXCEL_PATH = "https://grouperenault-my.sharepoint.com/:x:/g/personal/olivier_charron_renault_com/EV-ifyR8QNZNvKWJqfRFaFQB8OasV-O_x_I0lKEn9Zit-w?download=1"
JSON_FALLBACK_PATH = Path(__file__).parent / "data" / "ponts_clean.json"

@st.cache_data(ttl=3600)
def process_data(file_content, is_excel=True):
    """
    Cette fonction s'occupe uniquement du nettoyage (le calcul lourd).
    Elle est séparée pour pouvoir être mise en cache.
    """
    if is_excel:
        df = pd.read_excel(file_content, sheet_name="Ponts", header=9)
    else:
        # Cas du JSON fallback
        df = pd.read_json(file_content)

    # --- LOGIQUE DE NETTOYAGE ---
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    if "Pont" in df.columns:
        df = df[df["Pont"].notna()]

    rename_map = {
        "PAYS": "pays", "Site": "site", "Pont": "pont",
        "MES": "annee_mes", "Age": "age",
        "Evaluation Spéciale O/N": "evs_statut", "EVS Année": "evs_annee",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    if "pays" in df.columns:
        df["pays"] = df["pays"].astype(str).str.strip().str.upper().str.replace(r"^\d+\s*-\s*", "", regex=True)
    if "site" in df.columns:
        df["site"] = df["site"].astype(str).str.strip().str.upper()

    for col in ["annee_mes", "age", "evs_annee"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "evs_statut" in df.columns:
        df["evs_statut"] = df["evs_statut"].astype(str).str.strip().replace({
            "O": "Obligatoire", "Oui": "Obligatoire", "N": "Non requis", "Non": "Non requis"
        })

    budget_cols = [c for c in df.columns if any(k in str(c).upper() for k in ["OPEX", "RGE/RGM", "ACHAT NEUF"])]
    if budget_cols:
        df["budget_total"] = df[budget_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1)
    else:
        df["budget_total"] = 0
    
    evs_cols = [c for c in df.columns if "montant" in str(c).lower()]
    df["evs_montant"] = pd.to_numeric(df[evs_cols[0]], errors="coerce").fillna(0) if evs_cols else 0

    return df

def load_data():
    """
    Fonction principale appelée par vos pages. 
    Les widgets sont ICI, en dehors du cache.
    """
    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 Source de données")
    uploaded_file = st.sidebar.file_uploader("Mettre à jour via l'Excel", type="xlsx")

    # 1. Priorité à l'upload manuel
    if uploaded_file is not None:
        return process_data(uploaded_file, is_excel=True)

    # 2. Tentative SharePoint
    try:
        response = requests.get(EXCEL_PATH, timeout=5)
        response.raise_for_status()
        st.sidebar.success("🌐 SharePoint connecté")
        return process_data(io.BytesIO(response.content), is_excel=True)
    except:
        pass

    # 3. Fallback JSON
    if JSON_FALLBACK_PATH.exists():
        st.sidebar.info("💡 Données de secours")
        return process_data(JSON_FALLBACK_PATH, is_excel=False)
    
    st.error("Aucune donnée disponible.")
    return pd.DataFrame()

    # --- LOGIQUE DE NETTOYAGE (Identique à votre original) ---
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")
    if "Pont" in df.columns:
        df = df[df["Pont"].notna()]

    rename_map = {
        "PAYS": "pays", "Site": "site", "Pont": "pont",
        "MES": "annee_mes", "Age": "age",
        "Evaluation Spéciale O/N": "evs_statut", "EVS Année": "evs_annee",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    if "pays" in df.columns:
        df["pays"] = df["pays"].astype(str).str.strip().str.upper().str.replace(r"^\d+\s*-\s*", "", regex=True)
    if "site" in df.columns:
        df["site"] = df["site"].astype(str).str.strip().str.upper()

    for col in ["annee_mes", "age", "evs_annee"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "evs_statut" in df.columns:
        df["evs_statut"] = df["evs_statut"].astype(str).str.strip().replace({
            "O": "Obligatoire", "Oui": "Obligatoire", "N": "Non requis", "Non": "Non requis"
        })

    budget_cols = [c for c in df.columns if any(k in str(c).upper() for k in ["OPEX", "RGE/RGM", "ACHAT NEUF"])]
    df["budget_total"] = df[budget_cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum(axis=1) if budget_cols else 0
    
    evs_cols = [c for c in df.columns if "montant" in str(c).lower()]
    df["evs_montant"] = pd.to_numeric(df[evs_cols[0]], errors="coerce").fillna(0) if evs_cols else 0

    return df

# --- GARDER LES AUTRES FONCTIONS (calc_ifm, get_recommendation, etc.) ---
def apply_global_style():
    custom_css = f"""
    <style>
    .stApp {{ background-color: {COLORS['bg']}; color: {COLORS['text']}; }}
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

def format_euro(value):
    if pd.isna(value) or value == 0: return "—"
    return f"{value:,.0f} €".replace(",", " ")

def apply_global_style():
    custom_css = f"""
    <style>
    .stApp {{
        background-color: {COLORS['bg']};
        color: {COLORS['text']};
    }}
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

FEM_TIME_CLASSES = {
    "T0 - T ≤ 200 h": 200, "T1 - 200 < T ≤ 400 h": 400, "T2 - 400 < T ≤ 800 h": 800,
    "T3 - 800 < T ≤ 1 600 h": 1600, "T4 - 1 600 < T ≤ 3 200 h": 3200, "T5 - 3 200 < T ≤ 6 300 h": 6300,
    "T6 - 6 300 < T ≤ 12 500 h": 12500, "T7 - 12 500 < T ≤ 25 000 h": 25000, "T8 - 25 000 < T ≤ 50 000 h": 50000,
    "T9 - T > 50 000 h": 100000,
}

LOAD_SPECTRUM = {
    "L1 - km ≤ 0,125": 0.125, "L2 - 0,125 < km ≤ 0,250": 0.25,
    "L3 - 0,250 < km ≤ 0,500": 0.50, "L4 - 0,500 < km ≤ 1,000": 1.00,
}

def get_mechanism_group(time_class_label, spectrum_label):
    try:
        t_code = time_class_label.split(" - ")[0].strip()
        l_code = spectrum_label.split(" - ")[0].strip()
        t_index = int(t_code.replace("T", ""))
        return MECHANISM_GROUP_MATRIX[l_code][t_index]
    except (KeyError, ValueError, IndexError):
        return "N/A"

def calc_hresid(R, pi_c, pi_r, Kmf):
    if Kmf == 0: return None
    return (R * pi_c - pi_r) / Kmf

def calc_ifm(Hr, Kmr, Hc, Kmc):
    pi_r = Hr * Kmr
    pi_c = Hc * Kmc
    ifm = pi_r / pi_c if pi_c > 0 else None
    return ifm, pi_r, pi_c

def get_recommendation(ifm, hresid, hu):
    if ifm is None: return "Données insuffisantes."
    annees = hresid / hu if hresid is not None and hu > 0 else None
    if ifm >= 1.0: msg = "Analyse immédiate requise."
    elif ifm >= 0.75: msg = "EVS à planifier prioritairement."
    elif ifm >= 0.50: msg = "Surveillance renforcée recommandée."
    else: msg = "Situation normale."
    if annees is not None: msg += f" Durée résiduelle estimée : {annees:.1f} ans."
    return msg

def get_status(ifm):
    if ifm is None: return "Inconnu", "badge-normal", COLORS["grey"]
    if ifm < 0.50: return "Normal", "badge-normal", COLORS["green"]
    elif ifm < 0.75: return "Surveillance", "badge-watch", COLORS["yellow"]
    elif ifm < 1.0: return "Critique", "badge-critical", COLORS["orange"]
    return "Urgent", "badge-urgent", COLORS["red"]