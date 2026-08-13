"""
Trainingsanalyse: ATL / CTL / TSB (Streamlit-App)
===================================================

Interaktive Streamlit-App zur Trainingssteuerung auf Basis von TSS
(Training Stress Score):

  - TSS-Werte per Excel-Import UND/ODER manuelle Eingabe erfassen
  - ATL  (Acute Training Load / "Fatigue")   - Zeitkonstante frei wählbar
  - CTL  (Chronic Training Load / "Fitness") - Zeitkonstante frei wählbar
  - TSB  (Training Stress Balance / "Form")  = CTL - ATL
  - TSB-Bereiche (z.B. "Übertrainingsrisiko", "optimale Form") frei
    definierbar und im Diagramm als farbige Zonen dargestellt

Benötigte Pakete (requirements.txt):
    streamlit
    pandas
    matplotlib
    openpyxl

Start lokal:  streamlit run Trainingsanalyse.py
"""

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


st.set_page_config(page_title="Trainingsanalyse: ATL / CTL / TSB", layout="wide")
st.title("Trainingsanalyse: ATL / CTL / TSB")
st.caption(
    "TSS-Werte importieren oder manuell erfassen, ATL/CTL/TSB berechnen und "
    "die Form (TSB) mit frei definierbaren Bereichen visualisieren."
)


# =====================================================================
# Session State initialisieren
# =====================================================================

if "manuelle_eintraege" not in st.session_state:
    st.session_state.manuelle_eintraege = pd.DataFrame(
        {"Datum": pd.Series(dtype="datetime64[ns]"), "TSS": pd.Series(dtype="float")}
    )

if "tsb_bereiche" not in st.session_state:
    # "von" des ersten und "bis" des letzten Bereichs werden automatisch auf
    # -unendlich / +unendlich erweitert (siehe effektive_bereiche()) - hier
    # reichen die inneren Grenzwerte.
    st.session_state.tsb_bereiche = pd.DataFrame(
        [
            {"von": -60.0, "bis": -30.0, "label": "Hohes Übertrainingsrisiko", "farbe": "#d03b3b"},
            {"von": -30.0, "bis": -10.0, "label": "Ermüdung / Formaufbau", "farbe": "#ec835a"},
            {"von": -10.0, "bis": 5.0, "label": "Optimale Form", "farbe": "#0ca30c"},
            {"von": 5.0, "bis": 25.0, "label": "Frische / Taper", "farbe": "#2a78d6"},
            {"von": 25.0, "bis": 60.0, "label": "Formverlust (zu viel Ruhe)", "farbe": "#898781"},
        ]
    )


# =====================================================================
# Hilfsfunktionen
# =====================================================================

def effektive_bereiche(bereiche_df):
    """Sortiert die Bereiche nach 'von' und erweitert den ersten/letzten
    Bereich automatisch auf -unendlich / +unendlich, damit man am Rand
    keine willkürliche Zahl eintragen muss."""
    b = bereiche_df.dropna(subset=["von", "bis", "label"]).copy()
    b = b.sort_values("von").reset_index(drop=True)
    if len(b) == 0:
        return b
    b.loc[0, "von_eff"] = -np.inf
    b.loc[len(b) - 1, "bis_eff"] = np.inf
    b["von_eff"] = b.get("von_eff", b["von"])
    b["bis_eff"] = b.get("bis_eff", b["bis"])
    b["von_eff"] = b["von_eff"].fillna(b["von"])
    b["bis_eff"] = b["bis_eff"].fillna(b["bis"])
    return b


def tsb_bereich_zuordnen(tsb_wert, bereiche_eff):
    for _, b in bereiche_eff.iterrows():
        if b["von_eff"] <= tsb_wert < b["bis_eff"]:
            return b["label"]
    return "unbekannt"


def tagesreihe_aufbauen(df_tss):
    """Nimmt ein DataFrame mit Spalten Datum/TSS, summiert Mehrfacheinträge
    pro Tag und füllt fehlende Tage mit TSS=0 (lückenlose Tagesreihe -
    wichtig, damit ATL/CTL bei Trainingspausen korrekt abklingen)."""
    df = df_tss.dropna(subset=["Datum"]).copy()
    df["Datum"] = pd.to_datetime(df["Datum"])
    df["TSS"] = pd.to_numeric(df["TSS"], errors="coerce").fillna(0)
    if len(df) == 0:
        return None
    s = df.groupby("Datum")["TSS"].sum().sort_index()
    alle_tage = pd.date_range(s.index.min(), s.index.max(), freq="D")
    s = s.reindex(alle_tage, fill_value=0)
    s.index.name = "Datum"
    return s.to_frame("TSS")


def berechne_atl_ctl_tsb(df, atl_tage, ctl_tage, ctl_start=0.0, atl_start=0.0):
    """Standard-EWMA-Formel:
    Wert_heute = Wert_gestern + (TSS_heute - Wert_gestern) / Zeitkonstante
    TSB_t = CTL_(t-1) - ATL_(t-1)  (Form zu Beginn des Tages, Standard-
    Konvention wie z.B. bei TrainingPeaks)."""
    df = df.copy()
    n = len(df)
    ctl = np.zeros(n)
    atl = np.zeros(n)
    tss = df["TSS"].to_numpy()

    ctl_prev, atl_prev = ctl_start, atl_start
    for i in range(n):
        ctl[i] = ctl_prev + (tss[i] - ctl_prev) / ctl_tage
        atl[i] = atl_prev + (tss[i] - atl_prev) / atl_tage
        ctl_prev, atl_prev = ctl[i], atl[i]

    df["CTL"] = ctl
    df["ATL"] = atl
    ctl_gestern = np.concatenate(([ctl_start], ctl[:-1]))
    atl_gestern = np.concatenate(([atl_start], atl[:-1]))
    df["TSB"] = ctl_gestern - atl_gestern
    return df


def plot_tsb(df, bereiche_eff, atl_tage, ctl_tage):
    farbe_tss, farbe_ctl, farbe_atl, farbe_tsb = "#c3c2b7", "#2a78d6", "#eb6834", "#0b0b0b"

    fig, (ax_last, ax_tsb) = plt.subplots(
        2, 1, figsize=(13, 8), sharex=True, gridspec_kw={"height_ratios": [1, 1.3]}
    )
    fig.patch.set_facecolor("#fcfcfb")

    ax_last.bar(df.index, df["TSS"], color=farbe_tss, width=1.0, label="TSS (täglich)", zorder=1)
    ax_last.plot(df.index, df["CTL"], color=farbe_ctl, linewidth=2, label=f"CTL (Fitness, {ctl_tage} Tage)", zorder=3)
    ax_last.plot(df.index, df["ATL"], color=farbe_atl, linewidth=2, label=f"ATL (Fatigue, {atl_tage} Tage)", zorder=3)
    ax_last.set_ylabel("TSS / CTL / ATL")
    ax_last.set_facecolor("#fcfcfb")
    ax_last.legend(loc="upper left", frameon=False)
    ax_last.spines[["top", "right"]].set_visible(False)
    ax_last.spines[["left", "bottom"]].set_color("#c3c2b7")
    ax_last.grid(axis="y", color="#e1e0d9", linewidth=0.8)
    ax_last.set_title("Trainingsbelastung: TSS, CTL, ATL", loc="left", color="#0b0b0b", fontsize=12, pad=10)

    if len(bereiche_eff) > 0:
        for _, b in bereiche_eff.iterrows():
            untere = b["von_eff"] if np.isfinite(b["von_eff"]) else df["TSB"].min() - 10
            obere = b["bis_eff"] if np.isfinite(b["bis_eff"]) else df["TSB"].max() + 10
            ax_tsb.axhspan(untere, obere, color=b["farbe"], alpha=0.18, zorder=0, label=b["label"])

    ax_tsb.plot(df.index, df["TSB"], color=farbe_tsb, linewidth=2, label="TSB (Form)", zorder=3)
    ax_tsb.axhline(0, color="#898781", linewidth=1, linestyle="--", zorder=2)
    ax_tsb.set_ylabel("TSB")
    ax_tsb.set_facecolor("#fcfcfb")
    ax_tsb.spines[["top", "right"]].set_visible(False)
    ax_tsb.spines[["left", "bottom"]].set_color("#c3c2b7")
    ax_tsb.grid(axis="y", color="#e1e0d9", linewidth=0.8)
    ax_tsb.set_title("Training Stress Balance (Form) mit definierten Bereichen", loc="left", color="#0b0b0b", fontsize=12, pad=10)

    handles, labels = ax_tsb.get_legend_handles_labels()
    ax_tsb.legend(handles, labels, loc="upper left", frameon=False, fontsize=9, ncol=1)

    ax_tsb.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax_tsb.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%Y"))
    fig.autofmt_xdate()
    fig.tight_layout()
    return fig


# =====================================================================
# Sidebar: Einstellungen
# =====================================================================

with st.sidebar:
    st.header("Einstellungen")

    atl_tage = st.number_input("ATL-Zeitkonstante (Tage)", min_value=1, max_value=200, value=7, step=1)
    ctl_tage = st.number_input("CTL-Zeitkonstante (Tage)", min_value=1, max_value=200, value=42, step=1)

    with st.expander("Erweitert: Startwerte"):
        ctl_start = st.number_input("CTL-Startwert", value=0.0, step=1.0)
        atl_start = st.number_input("ATL-Startwert", value=0.0, step=1.0)

    st.divider()
    st.subheader("TSB-Bereiche")
    st.caption(
        "Der unterste 'von'-Wert und der oberste 'bis'-Wert reichen "
        "automatisch bis -∞ / +∞. Zeilen über das '+' unten hinzufügen "
        "oder löschen."
    )
    st.session_state.tsb_bereiche = st.data_editor(
        st.session_state.tsb_bereiche,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "von": st.column_config.NumberColumn("von", help="untere Grenze (inklusive)"),
            "bis": st.column_config.NumberColumn("bis", help="obere Grenze (exklusive)"),
            "label": st.column_config.TextColumn("Bezeichnung"),
            "farbe": st.column_config.TextColumn("Farbe (Hex)", help="z.B. #2a78d6"),
        },
        key="tsb_bereiche_editor",
    )


# =====================================================================
# Datenerfassung: Excel-Import + manuelle Eingabe
# =====================================================================

col_import, col_manuell = st.columns(2)

with col_import:
    st.subheader("1. Excel-Import")
    hochgeladene_datei = st.file_uploader("Excel-Datei mit TSS-Werten", type=["xlsx", "xls"])

    df_import = pd.DataFrame({"Datum": pd.Series(dtype="datetime64[ns]"), "TSS": pd.Series(dtype="float")})

    if hochgeladene_datei is not None:
        try:
            excel_datei = pd.ExcelFile(hochgeladene_datei)
            sheet = st.selectbox("Tabellenblatt", excel_datei.sheet_names)
            df_roh = excel_datei.parse(sheet)

            spalten = list(df_roh.columns)
            datum_default = next((c for c in spalten if "datum" in str(c).lower() or "date" in str(c).lower()), spalten[0])
            tss_default = next((c for c in spalten if "tss" in str(c).lower()), spalten[-1])

            c1, c2 = st.columns(2)
            with c1:
                spalte_datum = st.selectbox("Spalte mit Datum", spalten, index=spalten.index(datum_default))
            with c2:
                spalte_tss = st.selectbox("Spalte mit TSS", spalten, index=spalten.index(tss_default))

            df_import = df_roh[[spalte_datum, spalte_tss]].rename(columns={spalte_datum: "Datum", spalte_tss: "TSS"})
            st.success(f"{len(df_import)} Zeilen aus '{hochgeladene_datei.name}' eingelesen.")
        except Exception as e:
            st.error(f"Datei konnte nicht gelesen werden: {e}")

with col_manuell:
    st.subheader("2. Manuelle Eingabe")
    st.caption("Einzelne TSS-Werte direkt eintragen, bearbeiten oder löschen (Zeilen über '+' unten hinzufügen).")
    st.session_state.manuelle_eintraege = st.data_editor(
        st.session_state.manuelle_eintraege,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Datum": st.column_config.DateColumn("Datum", format="DD.MM.YYYY"),
            "TSS": st.column_config.NumberColumn("TSS", min_value=0.0, step=1.0),
        },
        key="manuelle_eintraege_editor",
    )


# =====================================================================
# Daten kombinieren, berechnen, anzeigen
# =====================================================================

df_kombiniert = pd.concat([df_import, st.session_state.manuelle_eintraege], ignore_index=True)
df_kombiniert = df_kombiniert.dropna(subset=["Datum"])

st.divider()

if len(df_kombiniert) == 0:
    st.info("Noch keine Daten vorhanden. Excel-Datei importieren oder oben rechts manuell TSS-Werte eintragen.")
else:
    tagesreihe = tagesreihe_aufbauen(df_kombiniert)

    if tagesreihe is None or len(tagesreihe) == 0:
        st.warning("Aus den eingegebenen Daten konnte keine gültige Tagesreihe gebildet werden.")
    else:
        ergebnis = berechne_atl_ctl_tsb(tagesreihe, atl_tage, ctl_tage, ctl_start, atl_start)
        bereiche_eff = effektive_bereiche(st.session_state.tsb_bereiche)

        if len(bereiche_eff) > 0:
            ergebnis["TSB_Bereich"] = ergebnis["TSB"].apply(lambda v: tsb_bereich_zuordnen(v, bereiche_eff))
        else:
            ergebnis["TSB_Bereich"] = "unbekannt"

        st.subheader("Diagramm")
        fig = plot_tsb(ergebnis, bereiche_eff, atl_tage, ctl_tage)
        st.pyplot(fig, width="stretch")

        letzter_tag = ergebnis.iloc[-1]
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("CTL (Fitness)", f"{letzter_tag['CTL']:.1f}")
        m2.metric("ATL (Fatigue)", f"{letzter_tag['ATL']:.1f}")
        m3.metric("TSB (Form)", f"{letzter_tag['TSB']:.1f}")
        m4.metric("Bereich", letzter_tag["TSB_Bereich"])

        st.subheader("Ergebnistabelle")
        st.dataframe(ergebnis.round(2), width="stretch")

        excel_puffer = pd.ExcelWriter("ergebnis_export.xlsx", engine="openpyxl")
        ergebnis.round(2).to_excel(excel_puffer)
        excel_puffer.close()
        with open("ergebnis_export.xlsx", "rb") as f:
            st.download_button(
                "Ergebnistabelle als Excel herunterladen",
                data=f,
                file_name="tss_atl_ctl_tsb_ergebnis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )







