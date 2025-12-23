import io
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
        return f"{float(x):,.0f} €".replace(",", " ")
    except Exception:
        return ""

def fmt_pct(x):
    try:
        return f"{float(x)*100:.1f} %"
    except Exception:
        return ""

def safe_div(a, b):
    return a / b if b not in (0, None) else 0.0

def normalize_hours_df(df: pd.DataFrame) -> pd.DataFrame:
    """Normalise la table des heures (types + colonne heures facturables)."""
    if df is None or df.empty:
        df = pd.DataFrame([{"Personne": "", "Heures": 0.0, "Coef_production": 0.0}])

    out = df.copy()
    # Garantir colonnes
    for col in ["Personne", "Heures", "Coef_production"]:
        if col not in out.columns:
            out[col] = "" if col == "Personne" else 0.0

    out["Personne"] = out["Personne"].astype(str).fillna("")
    out["Heures"] = pd.to_numeric(out["Heures"], errors="coerce").fillna(0.0)
    out["Coef_production"] = pd.to_numeric(out["Coef_production"], errors="coerce").fillna(0.0)

    out["Heures_facturables"] = out["Heures"] * out["Coef_production"]
    return out[["Personne", "Heures", "Coef_production", "Heures_facturables"]]

def compute_year(ca, achats, df_hours, taux_horaire, coef_refact):
    ca = float(ca or 0.0)
    achats = float(achats or 0.0)
    taux_horaire = float(taux_horaire or 0.0)
    coef_refact = float(coef_refact or 0.0)

    marge = ca - achats
    tx_marge = safe_div(marge, ca) if ca else 0.0

    dfh = normalize_hours_df(df_hours)
    heures = float(dfh["Heures"].sum())
    heures_fact = float(dfh["Heures_facturables"].sum())

    ca_theo_achats = achats * coef_refact
    ca_theo_heures = heures_fact * taux_horaire
    ca_theo_total = ca_theo_achats + ca_theo_heures

    ecart = ca - ca_theo_total
    ecart_pct = safe_div(ecart, ca) if ca else 0.0

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
# Excel (trame + import)
# =========================
REQUIRED_COLS = ["Personne", "Heures", "Coef_production"]

def make_template_excel_bytes() -> bytes:
    """Génère une trame Excel (2 onglets N et N-1) en mémoire."""
    df_n = pd.DataFrame(
        [
            {"Personne": "Ouvrier 1", "Heures": 140, "Coef_production": 0.75},
            {"Personne": "Ouvrier 2", "Heures": 152, "Coef_production": 0.70},
        ],
        columns=REQUIRED_COLS
    )
    df_n1 = pd.DataFrame(
        [
            {"Personne": "Ouvrier 1", "Heures": 138, "Coef_production": 0.72},
            {"Personne": "Ouvrier 2", "Heures": 150, "Coef_production": 0.68},
        ],
        columns=REQUIRED_COLS
    )

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_n.to_excel(writer, index=False, sheet_name="N")
        df_n1.to_excel(writer, index=False, sheet_name="N-1")
    return buf.getvalue()

def read_hours_sheet(xls: pd.ExcelFile, sheet_name: str) -> pd.DataFrame | None:
    if sheet_name not in xls.sheet_names:
        return None
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    missing = set(REQUIRED_COLS) - set(df.columns)
    if missing:
        raise ValueError(f"Onglet '{sheet_name}' : colonnes manquantes: {', '.join(sorted(missing))}")
    df = df[REQUIRED_COLS].copy()
    return df

def load_hours_from_excel(uploaded_file):
    """Attend un xlsx avec onglet N (obligatoire) + onglet N-1 (optionnel)."""
    xls = pd.ExcelFile(uploaded_file)
    df_n = read_hours_sheet(xls, "N")
    if df_n is None:
        raise ValueError("L'onglet 'N' est obligatoire.")
    df_n1 = read_hours_sheet(xls, "N-1")  # optionnel
    return df_n, df_n1

# =========================
# Default session state
# =========================
if "hours_n" not in st.session_state:
    st.session_state["hours_n"] = pd.DataFrame(
        [
            {"Personne": "Ouvrier 1", "Heures": 140, "Coef_production": 0.75},
            {"Personne": "Ouvrier 2", "Heures": 140, "Coef_production": 0.70},
        ],
        columns=REQUIRED_COLS
    )

if "hours_n1" not in st.session_state:
    st.session_state["hours_n1"] = pd.DataFrame(
        [
            {"Personne": "Ouvrier 1", "Heures": 140, "Coef_production": 0.70},
            {"Personne": "Ouvrier 2", "Heures": 140, "Coef_production": 0.68},
        ],
        columns=REQUIRED_COLS
    )

# =========================
# UI
# =========================
st.title("Rapprochement CA / Achats / Heures — Bâtiment")
st.caption("Comparer le CA réel à un CA théorique (refact achats + facturation heures), en N et N-1.")

# =========================
# Sidebar – Paramètres
# =========================
with st.sidebar:
    st.header("Paramètres – Année N")
    taux_horaire_n = st.number_input("Taux horaire N (€/h)", min_value=0.0, value=55.0, step=1.0)
    coef_refact_n = st.number_input("Coef refact achats N", min_value=0.0, value=1.15, step=0.01, format="%.2f")

    st.divider()
    use_n1 = st.checkbox("Activer l’année N-1", value=False)

    if use_n1:
        st.header("Paramètres – Année N-1")
        taux_horaire_n1 = st.number_input("Taux horaire N-1 (€/h)", min_value=0.0, value=52.0, step=1.0)
        coef_refact_n1 = st.number_input("Coef refact achats N-1", min_value=0.0, value=1.12, step=0.01, format="%.2f")
    else:
        taux_horaire_n1 = None
        coef_refact_n1 = None

# =========================
# 0) Trame Excel + Import
# =========================
st.subheader("0️⃣ Trame Excel (téléchargement) & Import")

left, right = st.columns([1, 1])

with left:
    st.markdown("### Télécharger une trame Excel")
    template_bytes = make_template_excel_bytes()
    st.download_button(
        label="📄 Télécharger la trame (.xlsx)",
        data=template_bytes,
        file_name="trame_heures_batiment.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("Trame avec 2 onglets : **N** et **N-1** (colonnes : Personne, Heures, Coef_production).")

with right:
    st.markdown("### Importer une trame Excel")
    uploaded = st.file_uploader(
        "Importer un fichier .xlsx (onglet 'N' obligatoire, 'N-1' optionnel)",
        type=["xlsx"],
        accept_multiple_files=False
    )

    col_btn1, col_btn2 = st.columns([1, 1])
    with col_btn1:
        if uploaded is not None and st.button("🔄 Charger depuis l'Excel", use_container_width=True):
            try:
                df_n, df_n1 = load_hours_from_excel(uploaded)
                st.session_state["hours_n"] = df_n.copy()
                if df_n1 is not None:
                    st.session_state["hours_n1"] = df_n1.copy()
                st.success("Import OK ✅ Les tableaux ont été rechargés. Tu peux ensuite modifier manuellement.")
            except Exception as e:
                st.error(f"Import impossible : {e}")

    with col_btn2:
        with st.expander("Aperçu trame (exemple)"):
            st.write("**Onglet N** / **Onglet N-1** : mêmes colonnes")
            st.dataframe(
                pd.DataFrame(
                    [
                        {"Personne": "Ouvrier 1", "Heures": 140, "Coef_production": 0.75},
                        {"Personne": "Ouvrier 2", "Heures": 152, "Coef_production": 0.70},
                    ],
                    columns=REQUIRED_COLS
                ),
                use_container_width=True
            )

st.divider()

# =========================
# 1) CA / Achats
# =========================
st.subheader("1️⃣ Chiffre d’affaires et achats")

c1, c2 = st.columns(2)

with c1:
    st.markdown("### Année N")
    ca_n = st.number_input("CA N", min_value=0.0, value=300000.0, step=1000.0)
    achats_n = st.number_input("Achats N", min_value=0.0, value=150000.0, step=1000.0)

with c2:
    st.markdown("### Année N-1")
    if use_n1:
        ca_n1 = st.number_input("CA N-1", min_value=0.0, value=280000.0, step=1000.0)
        achats_n1 = st.number_input("Achats N-1", min_value=0.0, value=140000.0, step=1000.0)
    else:
        ca_n1 = None
        achats_n1 = None
        st.info("N-1 désactivé")

st.divider()

# =========================
# 2) Heures – Edition manuelle (toujours possible)
# =========================
st.subheader("2️⃣ Heures par personne (modifiable manuellement)")

h1, h2 = st.columns(2)

with h1:
    st.markdown("### Ouvriers – N")
    st.session_state["hours_n"] = st.data_editor(
        st.session_state["hours_n"],
        num_rows="dynamic",
        use_container_width=True,
        key="editor_hours_n",
        column_config={
            "Personne": st.column_config.TextColumn(required=True),
            "Heures": st.column_config.NumberColumn(min_value=0.0, step=1.0),
            "Coef_production": st.column_config.NumberColumn(min_value=0.0, max_value=2.0, step=0.01, format="%.2f"),
        },
    )

with h2:
    st.markdown("### Ouvriers – N-1")
    if use_n1:
        if st.button("Copier le tableau N → N-1", use_container_width=True):
            st.session_state["hours_n1"] = st.session_state["hours_n"].copy()

        st.session_state["hours_n1"] = st.data_editor(
            st.session_state["hours_n1"],
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
        st.info("N-1 désactivé")

st.caption("Heures facturables = Heures × Coef de production.")
st.divider()

# =========================
# 3) Calculs
# =========================
res_n = compute_year(ca_n, achats_n, st.session_state["hours_n"], taux_horaire_n, coef_refact_n)
res_n1 = None
if use_n1 and ca_n1 is not None and achats_n1 is not None:
    res_n1 = compute_year(ca_n1, achats_n1, st.session_state["hours_n1"], taux_horaire_n1, coef_refact_n1)

# =========================
# KPI
# =========================
st.subheader("3️⃣ Résultats")

k1, k2, k3, k4 = st.columns(4)
k1.metric("Marge N", fmt_eur(res_n["marge"]), fmt_pct(res_n["tx_marge"]))
k2.metric("Heures facturables N", f"{res_n['heures_fact']:.1f} h")
k3.metric("CA théorique N", fmt_eur(res_n["ca_theo_total"]))
k4.metric("Écart N (réel − théorique)", fmt_eur(res_n["ecart"]), fmt_pct(res_n["ecart_pct"]))

if res_n1 is not None:
    st.divider()
    k1b, k2b, k3b, k4b = st.columns(4)
    k1b.metric("Marge N-1", fmt_eur(res_n1["marge"]), fmt_pct(res_n1["tx_marge"]))
    k2b.metric("Heures facturables N-1", f"{res_n1['heures_fact']:.1f} h")
    k3b.metric("CA théorique N-1", fmt_eur(res_n1["ca_theo_total"]))
    k4b.metric("Écart N-1 (réel − théorique)", fmt_eur(res_n1["ecart"]), fmt_pct(res_n1["ecart_pct"]))

st.divider()

# =========================
# 4) Graphiques (Altair)
# =========================
st.subheader("4️⃣ Analyse graphique")

# --- CA réel vs CA théorique (groupé)
rows = [
    {"Année": "N", "Type": "CA réel", "Montant": float(ca_n)},
    {"Année": "N", "Type": "CA théorique", "Montant": float(res_n["ca_theo_total"])},
]
if res_n1 is not None:
    rows += [
        {"Année": "N-1", "Type": "CA réel", "Montant": float(ca_n1)},
        {"Année": "N-1", "Type": "CA théorique", "Montant": float(res_n1["ca_theo_total"])},
    ]
df_compare = pd.DataFrame(rows)

chart_ca = (
    alt.Chart(df_compare)
    .mark_bar()
    .encode(
        x=alt.X("Année:N", title="Année"),
        xOffset=alt.XOffset("Type:N"),
        y=alt.Y("Montant:Q", title="Montant (€)"),
        color=alt.Color("Type:N", legend=alt.Legend(title="")),
        tooltip=[alt.Tooltip("Année:N"), alt.Tooltip("Type:N"), alt.Tooltip("Montant:Q", format=",.0f")],
    )
)

labels_ca = (
    alt.Chart(df_compare)
    .mark_text(dy=-8)
    .encode(
        x=alt.X("Année:N"),
        xOffset=alt.XOffset("Type:N"),
        y=alt.Y("Montant:Q"),
        detail="Type:N",
        text=alt.Text("Montant:Q", format=",.0f"),
    )
)

# --- Écart (réel - théorique)
gap_rows = [{"Année": "N", "Écart": float(res_n["ecart"])}]
if res_n1 is not None:
    gap_rows.append({"Année": "N-1", "Écart": float(res_n1["ecart"])})
df_gap = pd.DataFrame(gap_rows)

gap_bar = (
    alt.Chart(df_gap)
    .mark_bar()
    .encode(
        x=alt.X("Année:N", title="Année"),
        y=alt.Y("Écart:Q", title="Écart (€)"),
        color=alt.condition(
            alt.datum["Écart"] >= 0,
            alt.value("#2e7d32"),  # vert
            alt.value("#c62828"),  # rouge
        ),
        tooltip=[alt.Tooltip("Année:N"), alt.Tooltip("Écart:Q", format=",.0f")],
    )
)

gap_zero = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule().encode(y="y:Q")

gap_labels = (
    alt.Chart(df_gap)
    .mark_text(dy=-8)
    .encode(
        x="Année:N",
        y="Écart:Q",
        text=alt.Text("Écart:Q", format=",.0f"),
    )
)

# --- Composition du CA théorique (empilé)
comp_rows = [
    {"Année": "N", "Composant": "Achats / revente", "Montant": float(res_n["ca_theo_achats"])},
    {"Année": "N", "Composant": "Heures", "Montant": float(res_n["ca_theo_heures"])},
]
if res_n1 is not None:
    comp_rows += [
        {"Année": "N-1", "Composant": "Achats / revente", "Montant": float(res_n1["ca_theo_achats"])},
        {"Année": "N-1", "Composant": "Heures", "Montant": float(res_n1["ca_theo_heures"])},
    ]
df_comp = pd.DataFrame(comp_rows)

chart_comp = (
    alt.Chart(df_comp)
    .mark_bar()
    .encode(
        x=alt.X("Année:N", title="Année"),
        y=alt.Y("sum(Montant):Q", title="CA théorique (€)"),
        color=alt.Color("Composant:N", legend=alt.Legend(title="")),
        tooltip=[alt.Tooltip("Année:N"), alt.Tooltip("Composant:N"), alt.Tooltip("Montant:Q", format=",.0f")],
    )
)

labels_comp = (
    alt.Chart(df_comp)
    .mark_text(color="white")
    .encode(
        x="Année:N",
        y=alt.Y("Montant:Q", stack="zero"),
        detail="Composant:N",
        text=alt.Text("Montant:Q", format=",.0f"),
    )
)

colA, colB = st.columns(2)
with colA:
    st.markdown("### CA réel vs CA théorique")
    st.altair_chart((chart_ca + labels_ca).properties(height=340), use_container_width=True)

with colB:
    st.markdown("### Écart (réel − théorique)")
    st.altair_chart((gap_bar + gap_zero + gap_labels).properties(height=340), use_container_width=True)

st.markdown("### Composition du CA théorique")
st.altair_chart((chart_comp + labels_comp).properties(height=360), use_container_width=True)

# =========================
# Détails (tables calculées)
# =========================
with st.expander("Détails des heures (avec heures facturables)"):
    st.markdown("#### N")
    st.dataframe(res_n["dfh"], use_container_width=True)
    if res_n1 is not None:
        st.markdown("#### N-1")
        st.dataframe(res_n1["dfh"], use_container_width=True)
