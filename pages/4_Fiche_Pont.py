from io import BytesIO
import io
import sys
from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

sys.path.append(str(Path(__file__).parent.parent))

from utils import load_data, PLOTLY_LAYOUT, COLORS, apply_global_style, format_euro

apply_global_style()

st.markdown("# 📋 Fiche Pont")
st.caption("Détail complet d'un équipement")
st.divider()

df = load_data()

if st.button("🔄 Actualiser les données", key="refresh_fiche"):
    st.cache_data.clear()
    st.rerun()

df = load_data()

def safe_num(row, col):
    val = row[col] if col in row.index else None
    if val is None:
        return 0.0
    try:
        v = float(val)
        return 0.0 if pd.isna(v) else v
    except (ValueError, TypeError):
        return 0.0


def make_fiche_pdf(r, df_budget=None):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40,
    )

    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph(f"Fiche Pont - {r.get('pont', 'Equipement')}", styles["Title"]))
    story.append(Spacer(1, 14))

    info_data = [
        ["Pays", str(r.get("pays", "—"))],
        ["Site", str(r.get("site", "—"))],
        ["Pont", str(r.get("pont", "—"))],
        ["Année MES", f"{r.get('annee_mes', '—'):.0f}" if pd.notna(r.get("annee_mes")) else "—"],
        ["Âge", f"{r.get('age', '—'):.0f} ans" if pd.notna(r.get("age")) else "—"],
        ["Statut EVS", str(r.get("evs_statut", "—"))],
        ["Année EVS", f"{r.get('evs_annee', '—'):.0f}" if pd.notna(r.get("evs_annee")) else "—"],
        ["Budget 2025-2030", format_euro(safe_num(r, "budget_total"))],
    ]

    table = Table(info_data, colWidths=[150, 330])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eeeeee")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("PADDING", (0, 0), (-1, -1), 8),
    ]))

    story.append(table)
    story.append(Spacer(1, 18))
    story.append(Paragraph("Caractéristiques techniques", styles["Heading2"]))

    details = [
        ["Accessoire", r.get("accessoire")],
        ["Travaux planifiés", r.get("travaux")],
        ["Prix neuf estimé", format_euro(safe_num(r, "prix_neuf"))],
        ["Montant EVS", format_euro(safe_num(r, "evs_montant"))],
        ["Observations", r.get("observations")],
        ["Hors scope", "Oui" if r.get("hors_scope") else "Non"],
        ["À désinvestir", "Oui" if r.get("a_desinvestir") else "Non"],
    ]

    details = [
        [k, str(v)]
        for k, v in details
        if v and str(v).lower() not in ("none", "nan", "false", "")
    ]

    if details:
        details_table = Table(details, colWidths=[150, 330])
        details_table.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("PADDING", (0, 0), (-1, -1), 8),
        ]))
        story.append(details_table)
    else:
        story.append(Paragraph("Aucune information complémentaire disponible.", styles["Normal"]))

    if df_budget is not None and df_budget["Total (€)"].notna().any():
        story.append(Spacer(1, 18))
        story.append(Paragraph("Budget prévisionnel détaillé", styles["Heading2"]))

        budget_data = [["Année", "OPEX (€)", "CAPEX RGE (€)", "CAPEX Neuf (€)", "Total (€)"]]

        for _, row in df_budget.iterrows():
            budget_data.append([
                str(row["Année"]),
                format_euro(row["OPEX (€)"]),
                format_euro(row["CAPEX RGE (€)"]),
                format_euro(row["CAPEX Neuf (€)"]),
                format_euro(row["Total (€)"]),
            ])

        budget_table = Table(budget_data, colWidths=[60, 95, 110, 110, 105])
        budget_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eeeeee")),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("PADDING", (0, 0), (-1, -1), 7),
        ]))
        story.append(budget_table)

    doc.build(story)
    buffer.seek(0)
    return buffer


def fiche_to_excel(r, df_budget):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        info_df = pd.DataFrame({
            "Champ": [
                "Pays",
                "Site",
                "Pont",
                "Année MES",
                "Âge",
                "Statut EVS",
                "Année EVS",
                "Budget total 2025-2030",
            ],
            "Valeur": [
                r.get("pays"),
                r.get("site"),
                r.get("pont"),
                r.get("annee_mes"),
                r.get("age"),
                r.get("evs_statut"),
                r.get("evs_annee"),
                safe_num(r, "budget_total"),
            ],
        })

        details_df = pd.DataFrame({
            "Champ": [
                "Accessoire",
                "Travaux planifiés",
                "Prix neuf estimé",
                "Montant EVS",
                "Observations",
                "Hors scope",
                "À désinvestir",
            ],
            "Valeur": [
                r.get("accessoire"),
                r.get("travaux"),
                safe_num(r, "prix_neuf"),
                safe_num(r, "evs_montant"),
                r.get("observations"),
                "Oui" if r.get("hors_scope") else "Non",
                "Oui" if r.get("a_desinvestir") else "Non",
            ],
        })

        info_df.to_excel(writer, sheet_name="Infos générales", index=False)
        details_df.to_excel(writer, sheet_name="Caractéristiques", index=False)

        if df_budget is not None:
            df_budget.to_excel(writer, sheet_name="Budget 2025-2030", index=False)

    output.seek(0)
    return output


col_s1, col_s2, col_s3 = st.columns(3)

pays_opts = sorted(df["pays"].dropna().unique().tolist())
sel_pays = col_s1.selectbox("Pays", pays_opts)

site_opts = sorted(df[df["pays"] == sel_pays]["site"].dropna().unique().tolist())
sel_site = col_s2.selectbox("Site", site_opts)

pont_opts = sorted(
    df[(df["pays"] == sel_pays) & (df["site"] == sel_site)]["pont"].dropna().unique().tolist()
)
sel_pont = col_s3.selectbox("Pont", pont_opts)

st.divider()

mask = (df["pays"] == sel_pays) & (df["site"] == sel_site) & (df["pont"] == sel_pont)
row = df[mask]

if row.empty:
    st.warning("Équipement non trouvé.")
    st.stop()

r = row.iloc[0]

st.markdown(f"## {r['pont']}")
st.markdown(f"**{r['pays']}** — {r['site']}")

budget_total = safe_num(r, "budget_total")

badge_map = {
    "Obligatoire": "badge-urgent",
    "Surveillance": "badge-watch",
    "Non requis": "badge-normal",
    "Non concerné": "badge-normal",
    "A désinvestir": "badge-critical",
    "Non renseigné": "badge-normal",
}

c1, c2, c3, c4, c5 = st.columns(5)

kpis = [
    (c1, f"{r['annee_mes']:.0f}" if pd.notna(r["annee_mes"]) else "—", "Année MES"),
    (c2, f"{r['age']:.0f} ans" if pd.notna(r["age"]) else "—", "Âge"),
    (c3, str(r["evs_statut"]), "Statut EVS"),
    (c4, f"{r['evs_annee']:.0f}" if pd.notna(r["evs_annee"]) else "—", "Année EVS"),
    (c5, format_euro(budget_total) if budget_total > 0 else "—", "Budget 2025-30"),
]

for col, val, label in kpis:
    with col:
        if label == "Statut EVS":
            badge = badge_map.get(val, "badge-normal")
            st.markdown(f"""
            <div class="kpi-card">
                <div style="margin-top:4px"><span class="{badge}">{val}</span></div>
                <div class="kpi-label">{label}</div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-value">{val}</div>
                <div class="kpi-label">{label}</div>
            </div>
            """, unsafe_allow_html=True)

st.divider()

col_d1, col_d2 = st.columns([1, 1])

with col_d1:
    st.markdown("### Caractéristiques techniques")

    details = {
        "Accessoire": r.get("accessoire"),
        "Travaux planifiés": r.get("travaux"),
        "Prix neuf estimé": format_euro(safe_num(r, "prix_neuf")) if safe_num(r, "prix_neuf") > 0 else None,
        "Montant EVS": format_euro(safe_num(r, "evs_montant")) if safe_num(r, "evs_montant") > 0 else None,
        "Observations": r.get("observations"),
        "Hors scope": "Oui" if r.get("hors_scope") else None,
        "À désinvestir": "Oui" if r.get("a_desinvestir") else None,
    }

    any_shown = False

    for k, v in details.items():
        if v and str(v).lower() not in ("none", "nan", "false", ""):
            st.markdown(f"**{k}** : {v}")
            any_shown = True

    if not any_shown:
        st.info("Aucune information complémentaire disponible.")

with col_d2:
    st.markdown("### Budget prévisionnel détaillé")

    years = list(range(2025, 2031))
    budget_rows = []

    for y in years:
        opex = safe_num(r, f"opex_{y}")
        rge = safe_num(r, f"capex_rge_{y}")
        neuf = safe_num(r, f"capex_neuf_{y}")
        total_y = opex + rge + neuf

        budget_rows.append({
            "Année": y,
            "OPEX (€)": opex if opex > 0 else None,
            "CAPEX RGE (€)": rge if rge > 0 else None,
            "CAPEX Neuf (€)": neuf if neuf > 0 else None,
            "Total (€)": total_y if total_y > 0 else None,
        })

    df_budget = pd.DataFrame(budget_rows)
    has_budget = df_budget["Total (€)"].notna().any()

    pdf_buffer = make_fiche_pdf(r, df_budget)
    excel_fiche = fiche_to_excel(r, df_budget)

    st.download_button(
        label="📄 Télécharger la fiche PDF",
        data=pdf_buffer,
        file_name=f"fiche_pont_{str(r['pont']).replace(' ', '_')}.pdf",
        mime="application/pdf",
    )

    st.download_button(
        label="📊 Télécharger la fiche Excel",
        data=excel_fiche,
        file_name=f"fiche_pont_{str(r['pont']).replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if has_budget:
        money_cols = ["OPEX (€)", "CAPEX RGE (€)", "CAPEX Neuf (€)", "Total (€)"]
        fmt = {c: (lambda x: f"{x:,.0f}" if pd.notna(x) else "—") for c in money_cols}

        st.dataframe(df_budget.style.format(fmt), use_container_width=True, hide_index=True)

        grand_total = df_budget["Total (€)"].sum()
        st.markdown(f"**Total 2025-2030 : `{grand_total:,.0f} €`**")
    else:
        st.info("Aucune dépense prévue 2025-2030 pour cet équipement.")

if has_budget:
    st.divider()
    st.markdown("### Visualisation budgétaire")

    df_b = pd.DataFrame(budget_rows)

    fig = go.Figure()

    fig.add_trace(go.Bar(
        name="OPEX",
        x=df_b["Année"],
        y=df_b["OPEX (€)"].fillna(0),
        marker_color=COLORS["blue"],
    ))

    fig.add_trace(go.Bar(
        name="CAPEX RGE/RGM",
        x=df_b["Année"],
        y=df_b["CAPEX RGE (€)"].fillna(0),
        marker_color=COLORS["mauve"],
    ))

    fig.add_trace(go.Bar(
        name="CAPEX Achat neuf",
        x=df_b["Année"],
        y=df_b["CAPEX Neuf (€)"].fillna(0),
        marker_color=COLORS["red"],
    ))

    fig.update_layout(**PLOTLY_LAYOUT, barmode="stack", height=300)

    fig.update_layout(
        yaxis_title="Montant (€)",
        xaxis_title="Année",
        xaxis=dict(tickvals=years, ticktext=[str(y) for y in years], gridcolor="#24283a"),
        yaxis=dict(gridcolor="#24283a"),
    )

    st.plotly_chart(fig, use_container_width=True)

st.divider()
st.markdown("### Autres équipements sur le même site")

peers = df[(df["site"] == sel_site) & (df["pont"] != sel_pont)][
    ["pont", "age", "evs_statut", "evs_annee", "budget_total"]
].copy()

peers.columns = ["Pont", "Âge", "Statut EVS", "Année EVS", "Budget total (€)"]
peers = peers.sort_values("Âge", ascending=False)

st.dataframe(peers, use_container_width=True, hide_index=True)
