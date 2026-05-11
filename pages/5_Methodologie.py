import streamlit as st
from utils import load_data, PLOTLY_LAYOUT, COLORS, apply_global_style, require_uploaded_excel

apply_global_style()
require_uploaded_excel()

st.markdown("# ℹ️ Méthodologie")
st.caption("Fondements de calcul, hypothèses et limites de l'outil EVS")
st.divider()

tab1, tab2, tab3, tab4 = st.tabs([
    "📐 Méthode de calcul",
    "📊 Lecture des résultats",
    "📚 Références",
    "⚠️ Limites & apport"
])

with tab1:
    st.markdown("## 1. Indice de Fatigue du Mécanisme")

    st.markdown("""
    L'indice de fatigue du mécanisme, noté **IFm**, permet d'estimer le niveau de consommation
    de la durée de vie théorique d'un mécanisme de pont roulant.

    Il compare l'endommagement réellement accumulé par le mécanisme à sa capacité théorique
    d'endommagement issue des données de conception.
    """)

    st.latex(r"IFm = \frac{\pi_r}{\pi_c}")
    st.latex(r"\pi_r = H_r \times K_{mr}")
    st.latex(r"\pi_c = H_c \times K_{mc}")

    st.markdown("""
    | Symbole | Définition | Unité |
    |---|---|---|
    | **IFm** | Indice de fatigue du mécanisme | Sans dimension |
    | **πr** | Endommagement réel cumulé | Heures équivalentes |
    | **πc** | Capacité d'endommagement de conception | Heures équivalentes |
    | **Hr** | Heures de fonctionnement réelles | Heures |
    | **Kmr** | Spectre de charge réel estimé | Sans dimension |
    | **Hc** | Heures de conception du mécanisme | Heures |
    | **Kmc** | Spectre de charge de conception | Sans dimension |
    """)

    st.divider()

    st.markdown("## 2. Durée de Vie Résiduelle")

    st.markdown("""
    La durée de vie résiduelle, notée **Hresid**, estime le nombre d'heures de fonctionnement
    encore disponibles avant atteinte de la limite théorique d'endommagement.
    """)

    st.latex(r"H_{resid} = \frac{R \cdot \pi_c - \pi_r}{K_{mf}}")

    st.markdown("""
    | Symbole | Définition | Unité |
    |---|---|---|
    | **Hresid** | Heures de fonctionnement résiduelles estimées | Heures |
    | **R** | Rapport d'endommagement admissible | Sans dimension |
    | **Kmf** | Spectre de charge futur estimé | Sans dimension |
    """)

    st.markdown("### Conversion en années résiduelles")

    st.latex(r"\text{Années résiduelles} = \frac{H_{resid}}{H_u}")

    st.markdown("""
    Où **Hu** représente le nombre d'heures d'utilisation annuelle réelle du pont.
    """)

    st.divider()

    st.markdown("## 3. Principe physique")

    st.markdown("""
    Le calcul repose sur une logique d'endommagement cumulatif. Chaque utilisation du mécanisme
    consomme une fraction de sa durée de vie théorique. Lorsque l'indice **IFm** atteint 1,
    la limite théorique de conception est considérée comme atteinte.
    """)

    st.latex(r"D = \sum_i \frac{n_i}{N_i} \leq 1")


with tab2:
    st.markdown("## Grille d'interprétation")

    st.markdown("""
    | IFm | Zone | Statut | Décision recommandée |
    |---|---|---|---|
    | **0 à 0,50** | Verte | Normal | Surveillance périodique standard |
    | **0,50 à 0,75** | Jaune | Surveillance | Inspection renforcée et suivi rapproché |
    | **0,75 à 1,00** | Orange | Critique | EVS à planifier prioritairement |
    | **> 1,00** | Rouge | Urgent | Analyse immédiate, arrêt ou remplacement selon expertise |
    """)

    st.info(
        "Ces seuils servent à structurer l'aide à la décision. Ils ne remplacent pas "
        "l'expertise d'un organisme compétent ou d'un inspecteur habilité."
    )

    st.divider()

    st.markdown("## Exemple d'application")

    st.markdown("### Mécanisme de levage")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        **Données d'entrée**

        | Variable | Valeur |
        |---|---|
        | Hr | 4 278 h |
        | Kmr | 0,32 |
        | Hc | 12 500 h |
        | Kmc | 0,25 |
        | R | 1,0 |
        | Kmf | 0,32 |
        """)

    with col2:
        st.markdown("""
        **Résultats**

        | Indicateur | Calcul | Résultat |
        |---|---|---|
        | πr | 4 278 × 0,32 | 1 369 |
        | πc | 12 500 × 0,25 | 3 125 |
        | IFm | 1 369 / 3 125 | 0,44 |
        | Hresid | (3 125 - 1 369) / 0,32 | 5 488 h |
        """)

    st.markdown("### Mécanisme de direction")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        **Données d'entrée**

        | Variable | Valeur |
        |---|---|
        | Hr | 1 455 h |
        | Kmr | 0,83 |
        | Hc | 12 500 h |
        | Kmc | 0,25 |
        | R | 1,0 |
        | Kmf | 0,83 |
        """)

    with col2:
        st.markdown("""
        **Résultats**

        | Indicateur | Calcul | Résultat |
        |---|---|---|
        | πr | 1 455 × 0,83 | 1 208 |
        | πc | 12 500 × 0,25 | 3 125 |
        | IFm | 1 208 / 3 125 | 0,39 |
        | Hresid | (3 125 - 1 208) / 0,83 | 2 310 h |
        """)

    st.markdown("### Mécanisme de translation")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        **Données d'entrée**

        | Variable | Valeur |
        |---|---|
        | Hr | 2 137 h |
        | Kmr | 0,41 |
        | Hc | 12 500 h |
        | Kmc | 0,25 |
        | R | 1,0 |
        | Kmf | 0,41 |
        """)

    with col2:
        st.markdown("""
        **Résultats**

        | Indicateur | Calcul | Résultat |
        |---|---|---|
        | πr | 2 137 × 0,41 | 876 |
        | πc | 12 500 × 0,25 | 3 125 |
        | IFm | 876 / 3 125 | 0,28 |
        | Hresid | (3 125 - 876) / 0,41 | 5 485 h |
        """)


with tab3:
    st.markdown("## Références utilisées")

    st.markdown("""
    | Référence | Rôle dans la méthode |
    |---|---|
    | **FEM 1.001** | Base de classification et de calcul des appareils de levage |
    | **ISO 4301-1** | Classification des appareils de levage |
    | **EN 13001-1** | Principes généraux de conception des appareils de levage |
    | **EN 13001-2** | Prise en compte des charges et sollicitations |
    | **EN 15011** | Exigences applicables aux ponts roulants et portiques |
    """)

    st.divider()

    st.markdown("## Spectre de charge")

    st.markdown("""
    Le facteur de spectre de charge traduit la sévérité d'utilisation du mécanisme.
    Plus le facteur est élevé, plus les sollicitations sont sévères.
    """)

    st.markdown("""
    | Classe | Facteur indicatif | Description |
    |---|---:|---|
    | L1 | 0,125 | Charges faibles ou occasionnelles |
    | L2 | 0,25 | Charges modérées |
    | L3 | 0,50 | Charges lourdes fréquentes |
    | L4 | 1,00 | Charge nominale très fréquente |
    """)

    st.divider()

    st.markdown("## Groupes de mécanismes")

    st.markdown("""
    | Groupe | Heures de conception indicatives |
    |---|---:|
    | M1 | 3 200 h |
    | M2 | 6 300 h |
    | M3 | 12 500 h |
    | M4 | 25 000 h |
    | M5 | 50 000 h |
    | M6 | 100 000 h |
    | M7 | 200 000 h |
    | M8 | 400 000 h |
    """)


with tab4:
    st.markdown("## Limites méthodologiques")

    st.markdown("""
    ### 1. Estimation du spectre réel

    En l'absence de mesures instrumentées, le facteur **Kmr** repose sur une estimation
    issue de l'historique, de l'usage constaté et de l'expertise métier.

    ### 2. Hypothèse d'endommagement linéaire

    La méthode suppose une accumulation linéaire de l'endommagement. Elle ne modélise pas
    finement les effets de séquence, de surcharge exceptionnelle ou de propagation réelle
    de fissures.

    ### 3. Données constructeur parfois incomplètes

    Pour les équipements anciens, les données de conception peuvent être absentes,
    approximatives ou reconstruites par analogie.

    ### 4. Outil d'aide à la décision

    Le dashboard ne remplace pas une expertise réglementaire. Il sert à structurer l'analyse,
    prioriser les équipements et préparer les décisions EVS / CAPEX.
    """)

    st.divider()

    st.markdown("## Hypothèses retenues dans l'outil")

    st.markdown("""
    | Hypothèse | Valeur ou principe |
    |---|---|
    | Endommagement | Linéaire et cumulatif |
    | Kmr | Constant sur la période passée |
    | Kmf | Constant sur la période future |
    | R | 1,0 par défaut |
    | Hc par défaut | 12 500 h lorsque la donnée constructeur est absente |
    | Âge | Calculé à partir de l'année de mise en service |
    """)

    st.divider()

    st.markdown("## Apport du mémoire")

    st.success("""
    L'apport principal est la digitalisation d'une méthode technique existante afin de transformer
    des données dispersées en outil d'aide à la décision industrielle.

    L'outil permet de standardiser les calculs, centraliser les données du parc, prioriser les actions
    EVS / CAPEX et faciliter la préparation des dossiers d'analyse.
    """)

    st.caption("Mémoire M2 | Digitalisation de l'EVS des ponts roulants | FEM 1.001 / EN 13001")


