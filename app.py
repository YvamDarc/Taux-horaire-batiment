import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(
    page_title="Rapprochement CA / Achats / Heures – Bâtiment",
    layout="wide"
)

# =========================
# Utils
# =========================
def fmt_eur(x):
    try:
        return f"{x:,.0f} €".replace(",", " ")
    except Exception:
        return ""

def fmt_pct(x):
    try:
        return f"{x*100:.1f} %"
    except Exception:
        return ""

def safe_div(a, b):
    return a / b if b not in (0, None) else 0.0

def normalize_hours_df(df):
    if df is None or df.empty:
        df = pd.DataFrame([{"Personne": "", "Heures": 0.0, "Coef_production": 0.0}])

    out = df.copy()
    out["Heures"] = pd.to_numeric(out["Heures"], errors="coerce").fillna(0.0)
    out["Coef_production"] = pd.to_numeric(out["Coef_production"], errors="coerce").fillna(0.0)
    out["Heures_facturables"] = out["Heures"] * out["Coef_production"]
    return out

def compute_year(ca, achats, df_hours, taux_horaire, coef_refact):
    marge = ca - achats
    tx_marge = safe_div(marge, ca)

    dfh = normalize_hours_df(df_hours)
    heures = dfh["Heures"].sum()
    heures_fact = dfh["Heures_facturables"].sum()

    ca_theo_achats = achats * coef_refact
    ca_theo_heures = heures_fact * taux_horaire
    ca_theo_total = ca_theo_achats + ca_theo_heures

    ecart = ca - ca_theo_total
    ecart_pct = safe_div(ecart, ca)

    return {
        "marge": marge,
        "tx_marge": tx_marge,
        "dfh": dfh,
        "heures": heures,
        "heures_fact": heures_fact,
        "ca_theo_achats": ca_theo_achats,
        "ca_theo_heures": ca_theo_heures,
        "ca_theo_total": ca_theo_total,
        "ecart": ecart,
        "ecart_pct": ecart_pct,
    }

# =========================
# UI
# =========================
st.title("Rapprochement CA / Achats / Heures — Bâtiment")
st.caption("Comparer ce qui est facturé à ce qui devrait l’être, en N et N-1.")

# =========================
# Sidebar – Paramètres
# =========================
with st.sidebar:
    st.header("Paramètres – Année N")
    taux_horaire_n = st.number_input("Taux horaire N (€/h)", 0.0, 200.0, 55.0, 1.0)
    coef_refact_n = st.number_input("Coef refact achats N", 0.0, 3.0, 1.15, 0.01)

    st.divider()

    use_n1 = st.checkbox("Activer l’année N-1")

    if use_n1:
        st.header("Paramètres – Année N-1")
        taux_horaire_n1 = st.number_input("Taux horaire N-1 (€/h)", 0.0, 200.0, 52.0, 1.0)
        coef_refact_n1 = st.number_input("Coef refact achats N-1", 0.0, 3.0, 1.12, 0.01)

# =========================
# CA / Achats
# =========================
st.subheader("1️⃣ Chiffre d’affaires et achats")

c1, c2 = st.columns(2)

with c1:
    st.markdown("### Année N")
    ca_n = st.number_input("CA N", 0.0, value=300000.0, step=1000.0)
    achats_n = st.number_input("Achats N", 0.0, value=150000.0, step=1000.0)

with c2:
    st.markdown("### Année N-1")
    if use_n1:
        ca_n1 = st.number_input("CA N-1", 0.0, value=280000.0, step=1000.0)
        achats_n1 = st.number_input("Achats N-1", 0.0, value=140000.0, step=1000.0)
    else:
        ca_n1 = achats_n1 = None
        st.info("N-1 désactivé")

# =========================
# Heures
# =========================
st.subheader("2️⃣ Heures par personne")

if "hours_n" not in st.session_state:
    st.session_state["hours_n"] = pd.DataFrame([
        {"Personne": "Ouvrier 1", "Heures": 140, "Coef_production": 0.75},
        {"Personne": "Ouvrier 2", "Heures": 140, "Coef_production": 0.70},
    ])

if "hours_n1" not in st.session_state:
    st.session_state["hours_n1"] = pd.DataFrame([
        {"Personne": "Ouvrier 1", "Heures": 140, "Coef_production": 0.70},
        {"Personne": "Ouvrier 2", "Heures": 140, "Coef_production": 0.68},
    ])

h1, h2 = st.columns(2)

with h1:
    st.markdown("### Ouvriers – N")
    st.session_state["hours_n"] = st.data_editor(
        st.session_state["hours_n"],
        num_rows="dynamic",
        use_container_width=True,
    )

with h2:
    st.markdown("### Ouvriers – N-1")
    if use_n1:
        if st.button("Copier N → N-1"):
            st.session_state["hours_n1"] = st.session_state["hours_n"].copy()
        st.session_state["hours_n1"] = st.data_editor(
            st.session_state["hours_n1"],
            num_rows="dynamic",
            use_container_width=True,
        )
    else:
        st.info("N-1 désactivé")

# =========================
# Calculs
# =========================
res_n = compute_year(ca_n, achats_n, st.session_state["hours_n"], taux_horaire_n, coef_refact_n)
res_n1 = None

if use_n1:
    res_n1 = compute_year(ca_n1, achats_n1, st.session_state["hours_n1"], taux_horaire_n1, coef_refact_n1)

# =========================
# KPI
# =========================
st.subheader("3️⃣ Résultats")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Marge N", fmt_eur(res_n["marge"]), fmt_pct(res_n["tx_marge"]))
k2.metric("Heures facturables N", f"{res_n['heures_fact']:.1f} h")
k3.metric("CA théorique N", fmt_eur(res_n["ca_theo_total"]))
k4.metric("Écart N", fmt_eur(res_n["ecart"]), fmt_pct(res_n["ecart_pct"]))

if res_n1:
    st.divider()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Marge N-1", fmt_eur(res_n1["marge"]), fmt_pct(res_n1["tx_marge"]))
    k2.metric("Heures facturables N-1", f"{res_n1['heures_fact']:.1f} h")
    k3.metric("CA théorique N-1", fmt_eur(res_n1["ca_theo_total"]))
    k4.metric("Écart N-1", fmt_eur(res_n1["ecart"]), fmt_pct(res_n1["ecart_pct"]))

# =========================
# GRAPHIQUES
# =========================
st.subheader("4️⃣ Analyse graphique")

# --- CA réel vs théorique
rows = [
    {"Année": "N", "Type": "CA réel", "Montant": ca_n},
    {"Année": "N", "Type": "CA théorique", "Montant": res_n["ca_theo_total"]},
]
if res_n1:
    rows += [
        {"Année": "N-1", "Type": "CA réel", "Montant": ca_n1},
        {"Année": "N-1", "Type": "CA théorique", "Montant": res_n1["ca_theo_total"]},
    ]

df_compare = pd.DataFrame(rows)

chart_ca = (
    alt.Chart(df_compare)
    .mark_bar()
    .encode(
        x=alt.X("Année:N"),
        xOffset="Type:N",
        y=alt.Y("Montant:Q", title="€"),
        color="Type:N",
        tooltip=["Type", alt.Tooltip("Montant:Q", format=",.0f")],
    )
)

# --- Écart
df_gap = pd.DataFrame([
    {"Année": "N", "Écart": res_n["ecart"]},
] + (
    [{"Année": "N-1", "Écart": res_n1["ecart"]}] if res_n1 else []
))

chart_gap = (
    alt.Chart(df_gap)
    .mark_bar()
    .encode(
        x="Année:N",
        y="Écart:Q",
        color=alt.condition(
            alt.datum["Écart"] >= 0,
            alt.value("#2e7d32"),
            alt.value("#c62828"),
        ),
        tooltip=[alt.Tooltip("Écart:Q", format=",.0f")],
    )
)

# --- Composition CA théorique
rows_comp = [
    {"Année": "N", "Composant": "Achats / revente", "Montant": res_n["ca_theo_achats"]},
    {"Année": "N", "Composant": "Heures", "Montant": res_n["ca_theo_heures"]},
]
if res_n1:
    rows_comp += [
        {"Année": "N-1", "Composant": "Achats / revente", "Montant": res_n1["ca_theo_achats"]},
        {"Année": "N-1", "Composant": "Heures", "Montant": res_n1["ca_theo_heures"]},
    ]

df_comp = pd.DataFrame(rows_comp)

chart_comp = (
    alt.Chart(df_comp)
    .mark_bar()
    .encode(
        x="Année:N",
        y="sum(Montant):Q",
        color="Composant:N",
        tooltip=["Composant", alt.Tooltip("Montant:Q", format=",.0f")],
    )
)

c1, c2 = st.columns(2)
with c1:
    st.markdown("### CA réel vs CA théorique")
    st.altair_chart(chart_ca, use_container_width=True)

with c2:
    st.markdown("### Écart (réel – théorique)")
    st.altair_chart(chart_gap, use_container_width=True)

st.markdown("### Composition du CA théorique")
st.altair_chart(chart_comp, use_container_width=True)
