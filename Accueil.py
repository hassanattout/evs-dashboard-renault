import streamlit as st

from utils import (
    apply_global_style,
    upload_excel_center,
    has_uploaded_excel,
    clear_uploaded_excel,
    load_data,
)

st.set_page_config(
    page_title="EVS Dashboard | Ponts Roulants",
    page_icon="🏗️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

apply_global_style()

st.markdown("# 🏗️ EVS Dashboard")
st.markdown("### Ponts roulants | Digitalisation de l'Évaluation Spéciale")

if not has_uploaded_excel():
    st.markdown(
        """
        <div class="hero-box">
        <div class="hero-subtitle">
        Importez le fichier Excel du parc pour démarrer. 
        Les pages du dashboard apparaîtront ensuite avec les données chargées.
        </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("## Import du fichier Excel")
    upload_excel_center()
    st.stop()

df = load_data()

st.success("Excel chargé avec succès. Vous pouvez maintenant utiliser les pages du menu à gauche.")

st.markdown("## Synthèse rapide")

c1, c2, c3, c4 = st.columns(4)

with c1:
    st.metric("Équipements", len(df))

with c2:
    st.metric("Pays", df["pays"].nunique() if "pays" in df.columns else 0)

with c3:
    st.metric("Sites", df["site"].nunique() if "site" in df.columns else 0)

with c4:
    budget = df["budget_total"].sum() if "budget_total" in df.columns else 0
    st.metric("Budget total", f"{budget / 1e6:.1f} M€")

st.divider()

if st.button("Changer de fichier Excel"):
    clear_uploaded_excel()
    st.rerun()