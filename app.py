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

def normalize_hours_df(df: pd.DataFrame) -> pd.DataFrame:
    """Garantit les colonnes et types attendus."""
    if df is None or df.empty:
        df = pd.DataFrame([{"Personne": "", "Heures": 0.0, "Coef_production": 0.0}])

    for col in ["Personne", "Heures", "Coef_production"]:
        if col not in df.columns:
            df[col] = "" if col == "Personne" else 0.0

    out = df.copy()
    out["Heures"] = pd.to_numeric(out["Heures"], errors="coerce").fillna(0.0)
    out["Coef_production"] = pd.to_numeric(out["Coef_production"], errors="coerce").fillna(0.0)
    out["Heures_facturables"] = out["Heures"] * out["Coef_production"]
    return out[["Personne", "Heures", "Coef_production", "Heures_facturables"]]

def compute_year(ca: float, achats: float, df_hours: pd.DataFrame, taux_horaire: float, coef_refact_achats: float):
    marge = ca - achats
    tx_marge = safe_div(marge, ca) if ca else 0.0

    dfh = normalize_hours_df(df_hours)
    total_heures = float(dfh["Heures"].sum())
    total_heures_fact = float(dfh["Heures_facturables"].sum())

    ca_theo_achats = achats * coef_refact_achats
    ca_theo_heures = total_heures_fact * taux_horaire
    ca_theo_total = ca_theo_achats + ca_theo_heures

    ecart = ca - ca_theo_total
    ecart_pct = safe_div(ecart, ca) if ca else 0.0

    return {
        "marge": marge,
        "tx_marge": tx_marge,
        "dfh": dfh,
        "total_heures": total_heures,
        "total_heures_fact": total_heures_fact,
        "ca_theo_achats": ca_theo_achats,
        "ca_theo_heures": ca_theo_heures,
        "ca_theo_total": ca_theo_total,
        "ecart": ecart,
        "ecart_pct": ecart_pct,
    }

st.title("Rapprochement CA / Achats / Heures — Bâtiment")
st.caption("Comparer le CA réel à un CA théorique (refact achats + facturation heures), en N et N-1.")

# -------------------------
# Sidebar paramètres
# -------------------------
with st.sidebar:
    st.header("Paramètres — Année N")
    taux_horaire_n = st.number_input("Taux horaire N (€/h)", min_value=0.0, value=55.0, step=1.0)
    coef_refact_achats_n = st.number_input(
        "Coef refact achats N (ex: 1.15)", min_value=0.0, value=1.15, step=0.01, format="%.2f"
    )

    st.divider()

    st.header("Paramètres — Année N-1 (optionnel)")
    use_n1 = st.checkbox("Activer N-1", value=False)

    if use_n1:
        taux_horaire_n1 = st.number_input("Taux horaire N-1 (€/h)", min_value=0.0, value=52.0, step=1.0)
        coef_refact_achats_n1 = st.number_input(
            "Coef refact achats N-1 (ex: 1.15)", min_value=0.0, value=1.12, step=0.01, format="%.2f"
        )
        copy_n_to_n1 = st.button("Copier les paramètres N → N-1")
        # Astuce: on copie à l'affichage via session_state
        if copy_n_to_n1:
            st.session_state["taux_horaire_n1_override"] = taux_horaire_n
            st.session_state["coef_refact_achats_n1_override"] = coef_refact_achats_n
    else:
        taux_horaire_n1 = None
        coef_refact_achats_n1 = None

    # Appliquer override si bouton "copier"
    if use_n1:
        if "taux_horaire_n1_override" in st.session_state:
            taux_horaire_n1 = float(st.session_state.pop("taux_horaire_n1_override"))
        if "coef_refact_achats_n1_override" in st.session_state:
            coef_refact_achats_n1 = float(st.session_state.pop("coef_refact_achats_n1_override"))

    st.divider()
    st.write(
        "Rappels :\n"
        "- **Coef de production** : heures payées → heures facturables\n"
        "- **Coef refact achats** : transforme les achats en CA marchandises théorique"
    )

# -------------------------
# Saisie CA / Achats
# -------------------------
st.subheader("1) Ventes / Achats")
cN, cN1 = st.columns(2)

with cN:
    st.markdown("### Année N")
    ca_n = st.number_input("Chiffre d'affaires N", min_value=0.0, value=300000.0, step=1000.0)
    achats_n = st.number_input("Achats N", min_value=0.0, value=150000.0, step=1000.0)

with cN1:
    st.markdown("### Année N-1")
    if use_n1:
        ca_n1 = st.number_input("Chiffre d'affaires N-1", min_value=0.0, value=280000.0, step=1000.0)
        achats_n1 = st.number_input("Achats N-1", min_value=0.0, value=140000.0, step=1000.0)
    else:
        st.info("N-1 désactivé. Active-le dans la barre latérale si besoin.")
        ca_n1 = None
        achats_n1 = None

st.divider()

# -------------------------
# Tableaux heures N / N-1
# -------------------------
st.subheader("2) Heures par personne")

if "hours_df_n" not in st.session_state:
    st.session_state["hours_df_n"] = pd.DataFrame(
        [
            {"Personne": "Ouvrier 1", "Heures": 140.0, "Coef_production": 0.75},
            {"Personne": "Ouvrier 2", "Heures": 140.0, "Coef_production": 0.70},
        ]
    )

if "hours_df_n1" not in st.session_state:
    st.session_state["hours_df_n1"] = pd.DataFrame(
        [
            {"Personne": "Ouvrier 1", "Heures": 140.0, "Coef_production": 0.70},
            {"Personne": "Ouvrier 2", "Heures": 140.0, "Coef_production": 0.68},
        ]
    )

colH1, colH2 = st.columns(2)

with colH1:
    st.markdown("### Ouvriers — N")
    st.session_state["hours_df_n"] = st.data_editor(
        st.session_state["hours_df_n"],
        num_rows="dynamic",
        use_container_width=True,
        key="editor_hours_n",
        column_config={
            "Personne": st.column_config.TextColumn(required=True),
            "Heures": st.column_config.NumberColumn(min_value=0.0, step=1.0),
            "Coef_production": st.column_config.NumberColumn(min_value=0.0, max_value=2.0, step=0.01, format="%.2f"),
        },
    )

with colH2:
    st.markdown("### Ouvriers — N-1")
    if use_n1:
        copy_hours_n_to_n1 = st.button("Copier le tableau N → N-1")
        if copy_hours_n_to_n1:
            st.session_state["hours_df_n1"] = st.session_state["hours_df_n"].copy()

        st.session_state["hours_df_n1"] = st.data_editor(
            st.session_state["hours_df_n1"],
            num_rows="dynamic",
            use_container_width=True,
            key="editor_hours_n1",
            column_config={
                "Personne": st.column_config.TextColumn(required=True),
                "Heures": st.column_config.NumberColumn(min_value=0.0, step=1.0),
                "Coef_production": st.column_config.NumberColumn(min_value=0.0, max_value=2.0, step=0.01, format="%.2f"),
            },
        )
    else:
        st.info("N-1 désactivé : tableau N-1 masqué.")

st.caption("Heures facturables = Heures × Coef de production.")

st.divider()

# -------------------------
# Calculs N et N-1
# -------------------------
res_n = compute_year(
    ca=ca_n,
    achats=achats_n,
    df_hours=st.session_state["hours_df_n"],
    taux_horaire=taux_horaire_n,
    coef_refact_achats=coef_refact_achats_n,
)

if use_n1:
    res_n1 = compute_year(
        ca=ca_n1,
        achats=achats_n1,
        df_hours=st.session_state["hours_df_n1"],
        taux_horaire=taux_horaire_n1,
        coef_refact_achats=coef_refact_achats_n1,
    )
else:
    res_n1 = None

st.subheader("3) Résultats & comparaison")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Marge N", fmt_eur(res_n["marge"]), fmt_pct(res_n["tx_marge"]))
k2.metric("Heures N", f'{res_n["total_heures"]:.1f} h', f'{res_n["total_heures_fact"]:.1f} h fact.')
k3.metric("CA théorique N", fmt_eur(res_n["ca_theo_total"]),
          f'Achats: {fmt_eur(res_n["ca_theo_achats"])} | Heures: {fmt_eur(res_n["ca_theo_heures"])}')
k4.metric("Écart N (réel - théorique)", fmt_eur(res_n["ecart"]), fmt_pct(res_n["ecart_pct"]))

if use_n1 and res_n1 is not None:
    st.divider()
    k1b, k2b, k3b, k4b = st.columns(4)
    k1b.metric("Marge N-1", fmt_eur(res_n1["marge"]), fmt_pct(res_n1["tx_marge"]))
    k2b.metric("Heures N-1", f'{res_n1["total_heures"]:.1f} h', f'{res_n1["total_heures_fact"]:.1f} h fact.')
    k3b.metric("CA théorique N-1", fmt_eur(res_n1["ca_theo_total"]),
              f'Achats: {fmt_eur(res_n1["ca_theo_achats"])} | Heures: {fmt_eur(res_n1["ca_theo_heures"])}')
    k4b.metric("Écart N-1 (réel - théorique)", fmt_eur(res_n1["ecart"]), fmt_pct(res_n1["ecart_pct"]))

st.divider()
st.subheader("Synthèse graphique")

comp_rows = [{"Année": "N", "CA réel": ca_n, "CA théorique": res_n["ca_theo_total"]}]
if use_n1 and res_n1 is not None:
    comp_rows.insert(0, {"Année": "N-1", "CA réel": ca_n1, "CA théorique": res_n1["ca_theo_total"]})

comp = pd.DataFrame(comp_rows).set_index("Année")
c1, c2 = st.columns(2)

with c1:
    st.bar_chart(comp[["CA réel", "CA théorique"]])

with c2:
    parts_n = pd.DataFrame(
        [
            {"Composant": "CA théorique achats (N)", "Montant": res_n["ca_theo_achats"]},
            {"Composant": "CA théorique heures (N)", "Montant": res_n["ca_theo_heures"]},
        ]
    ).set_index("Composant")
    st.bar_chart(parts_n)

if use_n1 and res_n1 is not None:
    st.caption("Le graphique de droite montre la décomposition du CA théorique de N (tu peux dupliquer pour N-1 si tu veux).")

with st.expander("Détail calculs (tableaux heures avec heures facturables)"):
    st.markdown("#### N")
    st.dataframe(res_n["dfh"], use_container_width=True)
    if use_n1 and res_n1 is not None:
        st.markdown("#### N-1")
        st.dataframe(res_n1["dfh"], use_container_width=True)
