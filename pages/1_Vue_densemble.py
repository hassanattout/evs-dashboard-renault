import sys
from pathlib import Path
import io

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

sys.path.append(str(Path(__file__).parent.parent))

from utils import COLORS, PLOTLY_LAYOUT, load_data, require_uploaded_excel
from utils import apply_global_style

apply_global_style()
require_uploaded_excel()

def format_fr_number(x):
    if pd.isna(x):
        return ""
    return f"{int(round(x)):,}".replace(",", " ")


def numeric_sum(df, cols):
    if not cols:
        return 0
    return df[cols].apply(pd.to_numeric, errors="coerce").fillna(0).sum().sum()


def get_years(df):
    years = sorted({
        int(str(col).strip()[:4])
        for col in df.columns
        if str(col).strip()[:4].isdigit()
        and 2025 <= int(str(col).strip()[:4]) <= 2100
    })
    return years if years else [2025]


def cols_for(df, year, keyword):
    return [
        c for c in df.columns
        if str(year) in str(c)
        and keyword.lower() in str(c).lower()
    ]


if st.button("🔄 Actualiser les données"):
    st.cache_data.clear()
    st.rerun()

df = load_data()
years = get_years(df)

st.markdown("# 📊 Vue d'ensemble du parc")
st.caption(f"Synthèse globale du parc et vision budgétaire CAPEX / OPEX sur l'horizon {min(years)}–{max(years)}")
st.divider()

total_budget = df["budget_total"].sum() if "budget_total" in df.columns else 0
age_moyen = df["age"].dropna().mean() if "age" in df.columns else 0
nb_sites = df["site"].dropna().nunique() if "site" in df.columns else 0
nb_pays = df["pays"].dropna().nunique() if "pays" in df.columns else 0

c1, c2, c3, c4, c5 = st.columns(5)

kpis = [
    (c1, f"{len(df)}", "Équipements"),
    (c2, f"{nb_pays}", "Pays"),
    (c3, f"{nb_sites}", "Sites"),
    (c4, f"{age_moyen:.0f} ans", "Âge moyen"),
    (c5, f"{total_budget / 1e6:.1f} M€", "Budget planifié"),
]

for col, val, label in kpis:
    with col:
        st.markdown(
            f"""
            <div class="kpi-card">
            <div class="kpi-value">{val}</div>
            <div class="kpi-label">{label}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.divider()

col_l, col_r = st.columns(2)

with col_l:
    st.markdown("## Équipements par pays")

    by_pays = (
        df.groupby("pays")
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=True)
    )

    fig = go.Figure(
        go.Bar(
            x=by_pays["count"],
            y=by_pays["pays"],
            orientation="h",
            marker=dict(color=COLORS["blue"], opacity=0.9),
            text=by_pays["count"],
            textposition="outside",
            cliponaxis=False,
        )
    )

    fig.update_layout(
        **PLOTLY_LAYOUT,
        height=380,
        xaxis_title="Nombre d'équipements",
        yaxis_title="",
    )

    st.plotly_chart(fig, use_container_width=True)

with col_r:
    st.markdown("## Statut EVS")

    evs_counts = df["evs_statut"].fillna("Non renseigné").value_counts().reset_index()
    evs_counts.columns = ["statut", "count"]

    color_map = {
        "Non requis": COLORS["grey"],
        "Obligatoire": COLORS["red"],
        "Non concerné": COLORS["sky"],
        "Non renseigné": COLORS["yellow"],
        "A désinvestir": COLORS["orange"],
    }

    fig2 = go.Figure(
        go.Pie(
            labels=evs_counts["statut"],
            values=evs_counts["count"],
            hole=0.58,
            marker=dict(
                colors=[color_map.get(s, COLORS["blue"]) for s in evs_counts["statut"]],
                line=dict(color="#0d0f16", width=2),
            ),
            textinfo="label+value",
        )
    )

    fig2.update_layout(
        **PLOTLY_LAYOUT,
        height=380,
        showlegend=False,
    )

    st.plotly_chart(fig2, use_container_width=True)

st.divider()

st.markdown(f"## Plan CAPEX / OPEX {min(years)}–{max(years)}")
st.caption("Horizon de planification budgétaire. Les montants regroupent OPEX, CAPEX RGE/RGM et CAPEX achat neuf.")

opex_by_year = []
capex_rge_by_year = []
capex_neuf_by_year = []

for y in years:
    opex_cols = cols_for(df, y, "OPEX")
    rge_cols = cols_for(df, y, "RGE/RGM")
    neuf_cols = cols_for(df, y, "ACHAT NEUF")

    opex_by_year.append(numeric_sum(df, opex_cols))
    capex_rge_by_year.append(numeric_sum(df, rge_cols))
    capex_neuf_by_year.append(numeric_sum(df, neuf_cols))

fig3 = go.Figure()

fig3.add_trace(go.Bar(name="OPEX", x=years, y=opex_by_year, marker_color=COLORS["blue"]))
fig3.add_trace(go.Bar(name="CAPEX RGE/RGM", x=years, y=capex_rge_by_year, marker_color=COLORS["mauve"]))
fig3.add_trace(go.Bar(name="CAPEX Achat neuf", x=years, y=capex_neuf_by_year, marker_color=COLORS["red"]))

fig3.update_layout(
    **PLOTLY_LAYOUT,
    barmode="stack",
    height=420,
    yaxis_title="Montant (€)",
    xaxis_title="Année",
)

fig3.update_xaxes(tickvals=years, ticktext=[str(y) for y in years])
fig3.update_yaxes(tickformat=",.0f")

st.plotly_chart(fig3, use_container_width=True)

st.divider()

st.markdown("## Carte thermique budgétaire par site")
st.caption("Matrice site × année : les montants élevés apparaissent en rouge foncé.")

heat_metric = st.selectbox(
    "Indicateur affiché",
    ["Budget total", "Interventions planifiées"],
)

heat_rows = []

for _, row in df.iterrows():
    for y in years:
        year_cols = [
            c for c in df.columns
            if str(y) in str(c)
            and (
                "OPEX" in str(c).upper()
                or "RGE/RGM" in str(c).upper()
                or "ACHAT NEUF" in str(c).upper()
            )
        ]

        budget_y = 0
        for c in year_cols:
            val = pd.to_numeric(row.get(c, 0), errors="coerce")
            if pd.notna(val):
                budget_y += val

        has_budget = budget_y > 0
        is_evs_oblig = row.get("evs_statut") == "Obligatoire"
        evs_due_this_year = is_evs_oblig and row.get("evs_annee") == y

        intervention_planned = has_budget or evs_due_this_year
        intervention_count = len([
	    c for c in year_cols
	    if pd.to_numeric(row.get(c,0), errors="coerce") > 0
	])

        if intervention_planned:
            heat_rows.append(
                {
                    "Site": row["site"],
                    "Année": str(y),
                    "Budget": budget_y,
                    "Intervention planifiées": int(has_budget or evs_due_this_year),
                }
            )

heat_df = pd.DataFrame(heat_rows)

if not heat_df.empty:
    value_col = {
        "Budget total": "Budget",
        "Interventions planifiées": "Intervention planifiées",
    }[heat_metric]

    value_label = {
        "Budget total": "Budget",
        "Interventions planifiées": "Interventions planifiées",
    }[heat_metric]

    pivot = (
        heat_df.groupby(["Site", "Année"])[value_col]
        .sum()
        .unstack()
        .fillna(0)
    )

    text_values = pivot.map(lambda x: format_fr_number(x) if x > 0 else "")

    fig_heat = go.Figure(
        go.Heatmap(
            z=pivot.values,
            x=pivot.columns,
            y=pivot.index,
            text=text_values.values,
            texttemplate="%{text}",
            colorscale="YlOrRd",
            zmin=0,
            zmax=pivot.values.max() if pivot.values.max() > 0 else 1,
            hovertemplate=(
                "Site: %{y}<br>"
                "Année: %{x}<br>"
                f"{value_label}: %{{z:,.0f}}"
                "<extra></extra>"
            ),
            colorbar=dict(title=value_label),
        )
    )

    fig_heat.update_layout(
        **PLOTLY_LAYOUT,
        height=max(600, 38 * len(pivot.index)),
        xaxis_title="Année",
        yaxis_title="Site",
    )

    fig_heat.update_xaxes(
        side="top",
        tickmode="array",
        tickvals=list(pivot.columns),
        ticktext=[str(y) for y in pivot.columns],
    )

    fig_heat.update_yaxes(automargin=True)

    st.plotly_chart(fig_heat, use_container_width=True)

else:
    st.info("Aucune donnée disponible pour générer la carte thermique.")

st.divider()

col_age, col_evs = st.columns(2)

with col_age:
    st.markdown("## Distribution des âges")

    fig_age = go.Figure(
        go.Histogram(
            x=df["age"].dropna(),
            nbinsx=20,
            marker=dict(
                color=COLORS["teal"],
                opacity=0.85,
                line=dict(color="#0d0f16", width=1),
            ),
        )
    )

    fig_age.add_vline(
        x=25,
        line_dash="dash",
        line_color=COLORS["yellow"],
        annotation_text="25 ans",
        annotation_position="top",
    )

    fig_age.add_vline(
        x=50,
        line_dash="dash",
        line_color=COLORS["red"],
        annotation_text="50 ans",
        annotation_position="top",
    )

    fig_age.update_layout(
        **PLOTLY_LAYOUT,
        height=340,
        xaxis_title="Âge (années)",
        yaxis_title="Nombre d'équipements",
        bargap=0.05,
    )

    st.plotly_chart(fig_age, use_container_width=True)

with col_evs:
    st.markdown("## Échéances EVS")

    evs_df = df[df["evs_annee"].notna()].copy()
    evs_df["evs_annee"] = pd.to_numeric(evs_df["evs_annee"], errors="coerce")
    evs_df = evs_df.dropna(subset=["evs_annee"])
    evs_df["evs_annee"] = evs_df["evs_annee"].astype(int)

    evs_by_year = (
        evs_df.groupby("evs_annee")
        .size()
        .reset_index(name="count")
        .sort_values("evs_annee")
    )

    evs_by_year = evs_by_year[
        (evs_by_year["evs_annee"] >= min(years))
        & (evs_by_year["evs_annee"] <= max(years))
    ]

    if evs_by_year.empty:
        st.info("Aucune échéance EVS sur l'horizon sélectionné.")
    else:
        bar_colors = [
            COLORS["red"] if y <= 2026
            else COLORS["orange"] if y <= 2030
            else COLORS["blue"]
            for y in evs_by_year["evs_annee"]
        ]

        fig_evs = go.Figure(
            go.Bar(
                x=evs_by_year["evs_annee"].astype(str),
                y=evs_by_year["count"],
                marker=dict(color=bar_colors),
                text=evs_by_year["count"],
                textposition="outside",
                cliponaxis=False,
            )
        )

        evs_layout = PLOTLY_LAYOUT.copy()
        evs_layout["margin"] = dict(t=80, b=50, l=50, r=20)

        fig_evs.update_layout(
            **evs_layout,
            height=420,
            xaxis_title="Année EVS",
            yaxis_title="Nombre d'équipements",
            bargap=0.20,
            uniformtext_minsize=12,
            uniformtext_mode="show",
        )

        fig_evs.update_traces(
            textfont_size=16,
            cliponaxis=False,
        )

        fig_evs.update_xaxes(
            type="category",
            tickangle=0,
        )

        fig_evs.update_yaxes(
            rangemode="tozero",
            dtick=1,
            range=[0, max(evs_by_year["count"]) + 2],
        )

        st.plotly_chart(fig_evs, use_container_width=True)

st.divider()

st.markdown("## ⚠️ Équipements EVS obligatoires")

required_cols = ["pays", "site", "pont", "age", "evs_annee", "evs_montant", "budget_total"]
available_cols = [c for c in required_cols if c in df.columns]

evs_oblig = df[df["evs_statut"] == "Obligatoire"][available_cols].copy()

if "evs_annee" in evs_oblig.columns:
    evs_oblig = evs_oblig.sort_values("evs_annee")

rename_cols = {
    "pays": "Pays",
    "site": "Site",
    "pont": "Pont",
    "age": "Âge",
    "evs_annee": "Année EVS",
    "evs_montant": "Montant EVS (€)",
    "budget_total": "Budget total (€)",
}

evs_oblig = evs_oblig.rename(columns=rename_cols)

st.dataframe(evs_oblig, use_container_width=True, hide_index=True)

st.divider()

st.markdown("## ⬇️ Export global Excel")


def to_excel(df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Données brutes", index=False)

        summary = (
            df.groupby(["pays", "site"])
            .agg(
                nb_equipements=("pont", "count"),
                age_moyen=("age", "mean"),
                evs_obligatoires=("evs_statut", lambda x: (x == "Obligatoire").sum()),
                budget_total=("budget_total", "sum"),
            )
            .reset_index()
        )

        summary.to_excel(writer, sheet_name="Résumé par site", index=False)

    output.seek(0)
    return output


excel_file = to_excel(df)

st.download_button(
    label="📥 Télécharger tout le parc (Excel)",
    data=excel_file,
    file_name="parc_global.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)