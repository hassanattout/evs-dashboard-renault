import io
import sys
from pathlib import Path

import pandas as pd
import streamlit as st

sys.path.append(str(Path(__file__).parent.parent))

from utils import load_data, apply_global_style, require_uploaded_excel

apply_global_style()
require_uploaded_excel()

# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────
def get_years(df):
    years = sorted({
        int(str(col).strip()[:4])
        for col in df.columns
        if str(col).strip()[:4].isdigit()
        and 2025 <= int(str(col).strip()[:4]) <= 2100
    })

    return years if years else [2025]


def numeric_sum(df, cols):
    if not cols:
        return 0

    return (
        df[cols]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0)
        .sum()
        .sum()
    )


def cols_for(df, year, keyword):
    return [
        c for c in df.columns
        if str(year) in str(c)
        and keyword.lower() in str(c).lower()
    ]


# ─────────────────────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────────────────────
def to_excel(df, years):
    output = io.BytesIO()

    df_export = df.copy()

    # Budget by year
    budget_year_rows = []

    for y in years:
        opex_cols = cols_for(df_export, y, "OPEX")
        rge_cols = cols_for(df_export, y, "RGE/RGM")
        neuf_cols = cols_for(df_export, y, "ACHAT NEUF")

        total = (
            numeric_sum(df_export, opex_cols)
            + numeric_sum(df_export, rge_cols)
            + numeric_sum(df_export, neuf_cols)
        )

        budget_year_rows.append({
            "Année": y,
            "Budget total (€)": total,
        })

    budget_year_df = pd.DataFrame(budget_year_rows)

    # Synthèse
    synthese_df = pd.DataFrame({
        "Indicateur": [
            "Nombre d'équipements",
            "Nombre de pays",
            "Nombre de sites",
            "Âge moyen",
            "EVS obligatoires",
            f"Budget total {min(years)}-{max(years)} (€)",
        ],
        "Valeur": [
            len(df_export),
            df_export["pays"].nunique() if "pays" in df_export.columns else 0,
            df_export["site"].nunique() if "site" in df_export.columns else 0,
            round(df_export["age"].mean(), 1) if "age" in df_export.columns else 0,
            (df_export["evs_statut"] == "Obligatoire").sum()
            if "evs_statut" in df_export.columns else 0,
            df_export["budget_total"].sum()
            if "budget_total" in df_export.columns else 0,
        ],
    })

    # Résumé par site
    if {"pays", "site", "pont"}.issubset(df_export.columns):

        site_summary = (
            df_export.groupby(["pays", "site"])
            .agg(
                nb_equipements=("pont", "count"),
                age_moyen=("age", "mean"),
                evs_obligatoires=(
                    "evs_statut",
                    lambda x: (x == "Obligatoire").sum()
                ),
                budget_total=("budget_total", "sum"),
            )
            .reset_index()
        )

        site_summary["age_moyen"] = site_summary["age_moyen"].round(1)

        site_summary = site_summary.sort_values(
            "budget_total",
            ascending=False
        )

    else:
        site_summary = pd.DataFrame()

    # EVS obligatoires
    if "evs_statut" in df_export.columns:

        evs_oblig = df_export[
            df_export["evs_statut"] == "Obligatoire"
        ].copy()

        if not evs_oblig.empty:

            evs_cols = [
                "pays",
                "site",
                "pont",
                "age",
                "evs_annee",
                "evs_montant",
                "budget_total",
            ]

            evs_cols = [
                c for c in evs_cols
                if c in evs_oblig.columns
            ]

            evs_oblig = evs_oblig[evs_cols]

            if "evs_annee" in evs_oblig.columns:
                evs_oblig = evs_oblig.sort_values(
                    "evs_annee",
                    ascending=True,
                    na_position="last",
                )

        else:
            evs_oblig = pd.DataFrame()

    # Top critiques
    top_critical = df_export.copy()

    if "evs_statut" in top_critical.columns:

        top_critical["score_criticite"] = 0

        top_critical.loc[
            top_critical["evs_statut"] == "Obligatoire",
            "score_criticite"
        ] += 100

        top_critical.loc[
            top_critical["evs_statut"] == "A désinvestir",
            "score_criticite"
        ] += 80

        top_critical.loc[
            top_critical["evs_statut"] == "Surveillance",
            "score_criticite"
        ] += 50

        if "age" in top_critical.columns:
            top_critical["score_criticite"] += (
                top_critical["age"].fillna(0)
            )

        if "budget_total" in top_critical.columns:

            max_budget = top_critical["budget_total"].max()

            if max_budget and max_budget > 0:

                top_critical["score_criticite"] += (
                    top_critical["budget_total"].fillna(0)
                    / max_budget
                    * 20
                )

        top_cols = [
            "pays",
            "site",
            "pont",
            "age",
            "evs_statut",
            "evs_annee",
            "evs_montant",
            "budget_total",
            "score_criticite",
        ]

        top_cols = [
            c for c in top_cols
            if c in top_critical.columns
        ]

        top_critical = (
            top_critical[top_cols]
            .sort_values(
                "score_criticite",
                ascending=False
            )
            .head(10)
        )

    else:
        top_critical = pd.DataFrame()

    # Excel writer
    with pd.ExcelWriter(output, engine="openpyxl") as writer:

        synthese_df.to_excel(
            writer,
            sheet_name="Synthèse",
            index=False
        )

        site_summary.to_excel(
            writer,
            sheet_name="Par site",
            index=False
        )

        evs_oblig.to_excel(
            writer,
            sheet_name="EVS obligatoires",
            index=False
        )

        budget_year_df.to_excel(
            writer,
            sheet_name=f"Budget {min(years)}-{max(years)}",
            index=False
        )

        top_critical.to_excel(
            writer,
            sheet_name="Top 10 critiques",
            index=False
        )

        df_export.to_excel(
            writer,
            sheet_name="Données brutes",
            index=False
        )

    output.seek(0)

    return output


# ─────────────────────────────────────────────────────────────
# Streamlit App
# ─────────────────────────────────────────────────────────────
st.markdown("# 🔍 Tableau de priorisation")

if st.button("🔄 Actualiser les données", key="refresh_priorisation"):
    st.cache_data.clear()
    st.rerun()

df = load_data()

years = get_years(df)

st.caption(
    f"{len(df)} équipements — filtrable par pays, site, statut EVS, budget"
)

st.divider()

# Debug colonnes
with st.expander("Colonnes détectées", expanded=False):
    st.write(df.columns.tolist())


# ─────────────────────────────────────────────────────────────
# Filters
# ─────────────────────────────────────────────────────────────
with st.expander("🔧 Filtres", expanded=True):

    fc1, fc2, fc3, fc4 = st.columns(4)

    pays_opts = ["Tous"] + sorted(
        df["pays"].dropna().unique().tolist()
    )

    sel_pays = fc1.selectbox(
        "Pays",
        pays_opts
    )

    if sel_pays != "Tous":

        site_opts = ["Tous"] + sorted(
            df[df["pays"] == sel_pays]["site"]
            .dropna()
            .unique()
            .tolist()
        )

    else:

        site_opts = ["Tous"] + sorted(
            df["site"].dropna().unique().tolist()
        )

    sel_site = fc2.selectbox(
        "Site",
        site_opts
    )

    evs_opts = ["Tous"] + sorted(
        df["evs_statut"].dropna().unique().tolist()
    )

    sel_evs = fc3.selectbox(
        "Statut EVS",
        evs_opts
    )

    hors_scope = fc4.checkbox(
        "Exclure Hors Scope",
        value=True
    )


# ─────────────────────────────────────────────────────────────
# Apply filters
# ─────────────────────────────────────────────────────────────
filtered = df.copy()

if sel_pays != "Tous":
    filtered = filtered[
        filtered["pays"] == sel_pays
    ]

if sel_site != "Tous":
    filtered = filtered[
        filtered["site"] == sel_site
    ]

if sel_evs != "Tous":
    filtered = filtered[
        filtered["evs_statut"] == sel_evs
    ]

if hors_scope and "hors_scope" in filtered.columns:
    filtered = filtered[
        ~filtered["hors_scope"].fillna(False)
    ]

st.caption(f"**{len(filtered)}** équipements affichés")

st.divider()


# ─────────────────────────────────────────────────────────────
# Display table
# ─────────────────────────────────────────────────────────────
display_cols = {
    "pays": "Pays",
    "site": "Site",
    "pont": "Pont",
    "annee_mes": "MES",
    "age": "Âge",
    "travaux": "Travaux",
    "evs_statut": "Statut EVS",
    "evs_annee": "Année EVS",
    "evs_montant": "Montant EVS (€)",
    "budget_total": "Budget total planifié (€)",
}

for y in years:

    year_cols = [
        c for c in filtered.columns
        if str(y) in str(c)
        and (
            "OPEX" in str(c).upper()
            or "RGE/RGM" in str(c).upper()
            or "ACHAT NEUF" in str(c).upper()
        )
    ]

    if year_cols:

        filtered[f"budget_{y}"] = (
            filtered[year_cols]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
            .sum(axis=1)
        )

        display_cols[f"budget_{y}"] = f"{y} (€)"

available_cols = {
    k: v
    for k, v in display_cols.items()
    if k in filtered.columns
}

df_display = filtered[
    list(available_cols.keys())
].copy()

df_display.columns = list(available_cols.values())


# ─────────────────────────────────────────────────────────────
# Sorting
# ─────────────────────────────────────────────────────────────
sort_options = [
    c for c in [
        "Budget total planifié (€)",
        "Âge",
        "Année EVS",
        "Pays",
        "Site",
    ]
    if c in df_display.columns
]

sort_by = st.selectbox(
    "Trier par",
    sort_options
)

asc = st.checkbox(
    "Ordre croissant",
    value=False
)

if sort_by in df_display.columns:

    df_display = df_display.sort_values(
        sort_by,
        ascending=asc,
        na_position="last",
    )


# ─────────────────────────────────────────────────────────────
# Formatting
# ─────────────────────────────────────────────────────────────
money_cols = [
    c for c in df_display.columns
    if "€" in c
]

for c in money_cols:
    df_display[c] = (
        pd.to_numeric(
            df_display[c],
            errors="coerce"
        )
        .round(0)
    )

for c in ["Âge", "MES", "Année EVS"]:

    if c in df_display.columns:

        df_display[c] = (
            pd.to_numeric(
                df_display[c],
                errors="coerce"
            )
            .round(0)
        )


def color_evs(val):

    colors = {
        "Obligatoire":
        "background-color:#3a0a0a; color:#f38ba8",

        "Surveillance":
        "background-color:#3a3010; color:#f9e2af",

        "Non requis":
        "background-color:#0d0d14; color:#6c7086",

        "Non concerné":
        "background-color:#0d1a2a; color:#89dceb",

        "A désinvestir":
        "background-color:#2a1a0a; color:#fab387",

        "Non renseigné":
        "background-color:#1a1a1a; color:#585b70",
    }

    return colors.get(val, "")


def color_budget(val):

    if pd.isna(val) or val == 0:
        return "color:#313244"

    return "color:#cdd6f4"


format_dict = {}

for c in money_cols:

    format_dict[c] = (
        lambda x:
        f"{x:,.0f}".replace(",", " ")
        if pd.notna(x) and x != 0
        else "—"
    )

for c in ["Âge", "MES", "Année EVS"]:

    if c in df_display.columns:

        format_dict[c] = (
            lambda x:
            f"{x:.0f}"
            if pd.notna(x)
            else "—"
        )

styled = (
    df_display.style
    .map(
        color_evs,
        subset=["Statut EVS"]
        if "Statut EVS" in df_display.columns
        else []
    )
    .map(
        color_budget,
        subset=money_cols
    )
    .format(
        format_dict,
        na_rep="—"
    )
)

st.dataframe(
    styled,
    use_container_width=True,
    hide_index=True,
    height=520,
)


# ─────────────────────────────────────────────────────────────
# Summary by site
# ─────────────────────────────────────────────────────────────
st.divider()

st.markdown("## Récapitulatif par site")

summary = (
    filtered.groupby(["pays", "site"])
    .agg(
        Équipements=("pont", "count"),
        Âge_moyen=("age", "mean"),
        EVS_obligatoires=(
            "evs_statut",
            lambda x: (x == "Obligatoire").sum()
        ),
        Budget_total=("budget_total", "sum"),
    )
    .reset_index()
)

summary.columns = [
    "Pays",
    "Site",
    "Équipements",
    "Âge moyen",
    "EVS Obligatoires",
    "Budget total (€)",
]

summary["Âge moyen"] = (
    summary["Âge moyen"].round(1)
)

summary["Budget total (€)"] = (
    pd.to_numeric(
        summary["Budget total (€)"],
        errors="coerce"
    )
    .round(0)
)

summary = summary.sort_values(
    "Budget total (€)",
    ascending=False
)

st.dataframe(
    summary,
    use_container_width=True,
    hide_index=True
)


# ─────────────────────────────────────────────────────────────
# CSV Export
# ─────────────────────────────────────────────────────────────
st.divider()

csv = (
    df_display.to_csv(
        index=False,
        sep=";",
        decimal=","
    )
    .encode("utf-8")
)

st.download_button(
    "⬇️ Exporter en CSV",
    data=csv,
    file_name="ponts_priorisation.csv",
    mime="text/csv"
)


# ─────────────────────────────────────────────────────────────
# Excel Export
# ─────────────────────────────────────────────────────────────
st.divider()

st.markdown("## ⬇️ Export par filtre")

excel_filtered = to_excel(
    filtered,
    years
)

st.download_button(
    label="📥 Télécharger la sélection (Excel)",
    data=excel_filtered,
    file_name="export_filtré.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)