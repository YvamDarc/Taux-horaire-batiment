import streamlit as st
import pandas as pd

st.set_page_config(page_title="Rapprochement CA / Achats / Heures (Bâtiment)", layout="wide")

def fmt_eur(x: float) -> str:
    try:
        return f"{x:,.0f} €".replace(",", " ")
    except Exception:
        return ""

def fmt_pct(x: float) -> str:
    try:
        return f"{x*100:.1f} %"
    except Exception:
        return ""

def safe_div(a: float, b: float) -> float:
    return a / b if b not in (0, None) else 0.0

st.title("Rapprochement CA / Achats / Heures — Bâtiment")
st.caption("Objectif : comparer le CA réel vs un CA théorique (heures + refacturation achats).")

with st.sidebar:
    st.header("Paramètres de facturation")
    taux_horaire = st.number_input("Taux horaire de facturation (€/h)", min_value=0.0, value=55.0, step=1.0)
    coef_refact_achats = st.number_input(
        "Coefficient de refacturation des achats (ex: 1.15)",
        min_value=0.0, value=1.15, step=0.01, format="%.2f"
    )
    st.divider()
    st.header("Aide")
    st.write(
        "- **Coef de production** : transforme les heures payées en heures facturables.\n"
        "  - ex : 140h payées × 0,75 = 105h facturables\n"
        "- **Coef refact achats** : marge brute théorique sur achats.\n"
        "  - ex : achats 100k × 1,15 = 115k de CA marchandises théorique"
    )

tab1, tab2 = st.tabs(["Saisie & Calculs", "Lecture & Interprétation"])

with tab1:
    st.subheader("1) Ventes / Achats (N et N-1)")
    colN, colN1 = st.columns(2)

    with colN:
        st.markdown("### Année N")
        ca_n = st.number_input("Chiffre d'affaires N", min_value=0.0, value=300000.0, step=1000.0)
        achats_n = st.number_input("Achats (marchandises / sous-traitance / matériaux) N", min_value=0.0, value=150000.0, step=1000.0)

    with colN1:
        st.markdown("### Année N-1 (optionnel)")
        use_n1 = st.checkbox("Renseigner N-1", value=False)
        if use_n1:
            ca_n1 = st.number_input("Chiffre d'affaires N-1", min_value=0.0, value=280000.0, step=1000.0)
            achats_n1 = st.number_input("Achats N-1", min_value=0.0, value=140000.0, step=1000.0)
        else:
            ca_n1 = None
            achats_n1 = None

    # Marge / taux marge
    marge_n = ca_n - achats_n
    tx_marge_n = safe_div(marge_n, ca_n)

    if use_n1:
        marge_n1 = ca_n1 - achats_n1
        tx_marge_n1 = safe_div(marge_n1, ca_n1)
    else:
        marge_n1 = None
        tx_marge_n1 = None

    st.divider()
    st.subheader("2) Heures par personne")

    if "hours_df" not in st.session_state:
        st.session_state["hours_df"] = pd.DataFrame(
            [
                {"Personne": "Ouvrier 1", "Heures": 140.0, "Coef_production": 0.75},
                {"Personne": "Ouvrier 2", "Heures": 140.0, "Coef_production": 0.70},
            ]
        )

    st.session_state["hours_df"] = st.data_editor(
        st.session_state["hours_df"],
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Personne": st.column_config.TextColumn(required=True),
            "Heures": st.column_config.NumberColumn(min_value=0.0, step=1.0),
            "Coef_production": st.column_config.NumberColumn(min_value=0.0, max_value=2.0, step=0.01, format="%.2f"),
        },
    )

    df = st.session_state["hours_df"].copy()
    # Nettoyage minimal
    for c in ["Heures", "Coef_production"]:
        df[c] = pd.to_numeric(df.get(c), errors="coerce").fillna(0.0)

    df["Heures_facturables"] = df["Heures"] * df["Coef_production"]
    total_heures = float(df["Heures"].sum())
    total_heures_fact = float(df["Heures_facturables"].sum())

    st.caption("Les **heures facturables** = Heures × Coef de production.")
    st.dataframe(df, use_container_width=True)

    st.divider()
    st.subheader("3) CA théorique & comparaison")

    # Théorie
    ca_theo_achats_n = achats_n * coef_refact_achats
    ca_theo_heures_n = total_heures_fact * taux_horaire
    ca_theo_total_n = ca_theo_achats_n + ca_theo_heures_n

    # Écarts vs CA réel N
    ecart_n = ca_n - ca_theo_total_n
    ecart_pct_n = safe_div(ecart_n, ca_n) if ca_n else 0.0

    # N-1 si présent
    if use_n1:
        ca_theo_achats_n1 = achats_n1 * coef_refact_achats
        ca_theo_heures_n1 = total_heures_fact * taux_horaire  # même heures par défaut (on pourra dupliquer plus tard)
        ca_theo_total_n1 = ca_theo_achats_n1 + ca_theo_heures_n1
        ecart_n1 = ca_n1 - ca_theo_total_n1
        ecart_pct_n1 = safe_div(ecart_n1, ca_n1) if ca_n1 else 0.0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Marge N", fmt_eur(marge_n), fmt_pct(tx_marge_n))
    k2.metric("Heures (total)", f"{total_heures:.1f} h", f"{total_heures_fact:.1f} h fact.")
    k3.metric("CA théorique N", fmt_eur(ca_theo_total_n), f"Achats: {fmt_eur(ca_theo_achats_n)} | Heures: {fmt_eur(ca_theo_heures_n)}")
    k4.metric("Écart N (CA réel - théorique)", fmt_eur(ecart_n), fmt_pct(ecart_pct_n))

    if use_n1:
        st.divider()
        k1b, k2b, k3b, k4b = st.columns(4)
        k1b.metric("Marge N-1", fmt_eur(marge_n1), fmt_pct(tx_marge_n1))
        k2b.metric("Heures (hypothèse N-1)", f"{total_heures:.1f} h", "mêmes heures/coeff par défaut")
        k3b.metric("CA théorique N-1", fmt_eur(ca_theo_total_n1), f"Achats: {fmt_eur(ca_theo_achats_n1)} | Heures: {fmt_eur(ca_theo_heures_n1)}")
        k4b.metric("Écart N-1 (CA réel - théorique)", fmt_eur(ecart_n1), fmt_pct(ecart_pct_n1))

    st.divider()
    st.subheader("Synthèse visuelle")
    comp = pd.DataFrame(
        [
            {"Année": "N", "CA réel": ca_n, "CA théorique": ca_theo_total_n},
        ]
    )
    if use_n1:
        comp = pd.concat(
            [
                pd.DataFrame([{"Année": "N-1", "CA réel": ca_n1, "CA théorique": ca_theo_total_n1}]),
                comp,
            ],
            ignore_index=True,
        )

    c1, c2 = st.columns(2)
    with c1:
        st.bar_chart(comp.set_index("Année")[["CA réel", "CA théorique"]])
    with c2:
        parts = pd.DataFrame(
            [
                {"Composant": "CA théorique achats", "Montant": ca_theo_achats_n},
                {"Composant": "CA théorique heures", "Montant": ca_theo_heures_n},
            ]
        ).set_index("Composant")
        st.bar_chart(parts)

with tab2:
    st.subheader("Comment lire le résultat")
    st.markdown(
        """
- Si **CA réel < CA théorique** : soit le **taux horaire**/la **refacturation achats** est sous-estimée *dans la réalité*,
  soit les **heures facturables** sont surestimées (coef production trop élevé), soit une partie du CA n’est pas “dans” le chantier (ou inversement).
- Si **CA réel > CA théorique** : soit tu factures plus que le modèle (bonus, travaux non saisis, coef refact plus fort, plus d’heures réellement facturées, etc.).
- Le but n’est pas la perfection, mais un **détecteur rapide** : “est-ce qu’on se paie correctement ?”.
"""
    )
    st.info("Astuce : commence avec 2-3 personnes et des coefficients prudents (ex 0,60 à 0,80), puis ajuste.")
