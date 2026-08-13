"""
Trainingsanalyse: ATL / CTL / TSB (Streamlit-App)
===================================================

Interaktive Streamlit-App zur Trainingssteuerung auf Basis von TSS
(Training Stress Score):

  - TSS-Werte (inkl. optionaler Trainingsdauer) per Excel-Import UND/ODER
    manuelle Eingabe erfassen
  - ATL  (Acute Training Load / "Fatigue")   - Zeitkonstante frei wählbar
  - CTL  (Chronic Training Load / "Fitness") - Zeitkonstante frei wählbar
  - TSB  (Training Stress Balance / "Form")  = CTL - ATL
  - TSB-Bereiche (z.B. "Übertrainingsrisiko", "optimale Form") frei
    definierbar und im Diagramm als farbige Zonen dargestellt
  - Diagramm und Auswertung frei auf einen Datumsbereich eingrenzbar

Benötigte Pakete (requirements.txt):
    streamlit
    pandas
    matplotlib
    openpyxl

Start lokal:  streamlit run Trainingsanalyse.py
"""

import datetime as dt

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
        {
            "Datum": pd.Series(dtype="datetime64[ns]"),
            "Dauer": pd.Series(dtype="object"),  # Trainingsdauer als datetime.time (hh:mm)
            "TSS": pd.Series(dtype="float"),
        }
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


def dauer_zu_minuten(wert):
    """Wandelt einen Dauer-Wert (datetime.time, timedelta, 'hh:mm'-String
    oder Dezimalstunden als Zahl) in Minuten um. Leere/ungültige Werte -> 0."""
    if wert is None:
        return 0.0
    try:
        if pd.isna(wert):
            return 0.0
    except (TypeError, ValueError):
        pass
    if isinstance(wert, dt.timedelta):
        return wert.total_seconds() / 60.0
    if isinstance(wert, dt.time):
        return wert.hour * 60 + wert.minute + wert.second / 60.0
    if isinstance(wert, pd.Timestamp):
        return wert.hour * 60 + wert.minute + wert.second / 60.0
    if isinstance(wert, str):
        s = wert.strip()
        if ":" in s:
            teile = s.split(":")
            try:
                h = int(teile[0])
                m = int(teile[1]) if len(teile) > 1 else 0
                return h * 60 + m
            except ValueError:
                return 0.0
        try:
            return float(s.replace(",", ".")) * 60.0  # Dezimalstunden, z.B. "1,5" -> 90 Min
        except ValueError:
            return 0.0
    if isinstance(wert, (int, float)):
        return float(wert) * 60.0  # Dezimalstunden, z.B. 1.5 -> 90 Min
    return 0.0


def minuten_zu_hhmm(minuten):
    """Formatiert eine Minutenanzahl als 'hh:mm'-String (auch > 24h möglich)."""
    minuten = int(round(minuten))
    return f"{minuten // 60:02d}:{minuten % 60:02d}"


def dauer_anzeige(wert):
    """Formatiert einen einzelnen Dauer-Wert für die Anzeige als 'hh:mm'
    (leer, falls keine Trainingszeit angegeben wurde)."""
    try:
        if wert is None or pd.isna(wert):
            return ""
    except (TypeError, ValueError):
        pass
    return minuten_zu_hhmm(dauer_zu_minuten(wert))


def hex_zu_rgba(hex_farbe, alpha=0.28):
    """Wandelt eine Hex-Farbe (#rrggbb) in einen rgba(...)-CSS-String um."""
    h = str(hex_farbe).lstrip("#")
    if len(h) != 6:
        return f"rgba(200,200,200,{alpha})"
    try:
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    except ValueError:
        return f"rgba(200,200,200,{alpha})"
    return f"rgba({r},{g},{b},{alpha})"


def zeilen_nach_tsb_bereich_faerben(tabelle, bereiche_eff):
    """Gibt einen pandas Styler zurück, der jede Zeile passend zu ihrem
    TSB-Bereich einfärbt (gleiche Farben wie im Diagramm)."""
    farb_map = dict(zip(bereiche_eff["label"], bereiche_eff["farbe"])) if len(bereiche_eff) > 0 else {}

    def zeile_stylen(zeile):
        farbe = farb_map.get(zeile.get("TSB_Bereich"))
        if farbe:
            return [f"background-color: {hex_zu_rgba(farbe)}"] * len(zeile)
        return [""] * len(zeile)

    return tabelle.style.apply(zeile_stylen, axis=1)


def tagesreihe_aufbauen(df_tss):
    """Nimmt ein DataFrame mit Spalten Datum/Dauer/TSS, summiert Mehrfach-
    einträge pro Tag und füllt fehlende Tage mit TSS=0/Dauer=0 (lückenlose
    Tagesreihe - wichtig, damit ATL/CTL bei Trainingspausen korrekt
    abklingen)."""
    df = df_tss.dropna(subset=["Datum"]).copy()
    df["Datum"] = pd.to_datetime(df["Datum"])
    df["TSS"] = pd.to_numeric(df["TSS"], errors="coerce").fillna(0)
    df["Dauer_min"] = df["Dauer"].apply(dauer_zu_minuten) if "Dauer" in df.columns else 0.0
    if len(df) == 0:
        return None
    tag_gruppe = df.groupby("Datum").agg(TSS=("TSS", "sum"), Dauer_min=("Dauer_min", "sum")).sort_index()
    alle_tage = pd.date_range(tag_gruppe.index.min(), tag_gruppe.index.max(), freq="D")
    tag_gruppe = tag_gruppe.reindex(alle_tage, fill_value=0)
    tag_gruppe.index.name = "Datum"
    return tag_gruppe


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

    df_import = pd.DataFrame(
        {
            "Datum": pd.Series(dtype="datetime64[ns]"),
            "Dauer": pd.Series(dtype="object"),
            "TSS": pd.Series(dtype="float"),
        }
    )

    if hochgeladene_datei is not None:
        try:
            excel_datei = pd.ExcelFile(hochgeladene_datei)
            sheet = st.selectbox("Tabellenblatt", excel_datei.sheet_names)
            df_roh = excel_datei.parse(sheet)

            spalten = list(df_roh.columns)
            datum_default = next((c for c in spalten if "datum" in str(c).lower() or "date" in str(c).lower()), spalten[0])
            tss_default = next((c for c in spalten if "tss" in str(c).lower()), spalten[-1])
            dauer_default = next((c for c in spalten if "dauer" in str(c).lower() or "duration" in str(c).lower() or "zeit" in str(c).lower()), None)

            c1, c2, c3 = st.columns(3)
            with c1:
                spalte_datum = st.selectbox("Spalte mit Datum", spalten, index=spalten.index(datum_default))
            with c2:
                spalte_tss = st.selectbox("Spalte mit TSS", spalten, index=spalten.index(tss_default))
            with c3:
                dauer_optionen = ["(keine)"] + spalten
                dauer_index = dauer_optionen.index(dauer_default) if dauer_default else 0
                spalte_dauer = st.selectbox("Spalte mit Trainingszeit (optional, hh:mm)", dauer_optionen, index=dauer_index)

            df_import = df_roh[[spalte_datum, spalte_tss]].rename(columns={spalte_datum: "Datum", spalte_tss: "TSS"})
            if spalte_dauer != "(keine)":
                df_import["Dauer"] = df_roh[spalte_dauer]
            else:
                df_import["Dauer"] = None
            st.success(f"{len(df_import)} Zeilen aus '{hochgeladene_datei.name}' eingelesen.")
        except Exception as e:
            st.error(f"Datei konnte nicht gelesen werden: {e}")

with col_manuell:
    st.subheader("2. Manuelle Eingabe")

    with st.form("neue_einheit_formular", clear_on_submit=True):
        st.caption("Neue Trainingseinheit erfassen (mit Datum, Trainingszeit und TSS):")
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            neues_datum = st.date_input("Datum", value=dt.date.today(), format="DD.MM.YYYY")
        with fc2:
            neue_dauer = st.time_input("Trainingszeit (hh:mm)", value=dt.time(1, 0), step=60)
        with fc3:
            neuer_tss = st.number_input("TSS", min_value=0.0, step=1.0, value=0.0)
        abgeschickt = st.form_submit_button("Einheit hinzufügen", width="stretch")

    if abgeschickt:
        neue_dauer_hhmm = f"{neue_dauer.hour:02d}:{neue_dauer.minute:02d}"
        neue_zeile = pd.DataFrame(
            [{"Datum": pd.Timestamp(neues_datum), "Dauer": neue_dauer_hhmm, "TSS": neuer_tss}]
        )
        st.session_state.manuelle_eintraege = pd.concat(
            [st.session_state.manuelle_eintraege, neue_zeile], ignore_index=True
        )
        st.success(
            f"Einheit am {neues_datum.strftime('%d.%m.%Y')} "
            f"({neue_dauer_hhmm} h) hinzugefügt."
        )

    st.caption("Vorhandene manuelle Einträge bearbeiten oder löschen (Zeilen über '+'/Papierkorb unten):")
    st.session_state.manuelle_eintraege = st.data_editor(
        st.session_state.manuelle_eintraege,
        num_rows="dynamic",
        width="stretch",
        column_config={
            "Datum": st.column_config.DateColumn("Datum", format="DD.MM.YYYY"),
            "Dauer": st.column_config.TextColumn(
                "Trainingsdauer (hh:mm)", help="Format hh:mm, z.B. 1:30 für 1 Stunde 30 Minuten"
            ),
            "TSS": st.column_config.NumberColumn("TSS", min_value=0.0, step=1.0),
        },
        column_order=["Datum", "Dauer", "TSS"],
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
        # ATL/CTL/TSB werden IMMER über die komplette Historie berechnet -
        # sonst würde ein später gewählter Anzeige-Zeitraum die kumulierten
        # Werte verfälschen (ATL/CTL bräuchten sonst wieder bei 0 anfangen).
        ergebnis_gesamt = berechne_atl_ctl_tsb(tagesreihe, atl_tage, ctl_tage, ctl_start, atl_start)
        bereiche_eff = effektive_bereiche(st.session_state.tsb_bereiche)

        if len(bereiche_eff) > 0:
            ergebnis_gesamt["TSB_Bereich"] = ergebnis_gesamt["TSB"].apply(lambda v: tsb_bereich_zuordnen(v, bereiche_eff))
        else:
            ergebnis_gesamt["TSB_Bereich"] = "unbekannt"

        # ---- Zeitraum-Auswahl (gilt für Diagramm + Auswertung) ----
        gesamt_min = ergebnis_gesamt.index.min().date()
        gesamt_max = ergebnis_gesamt.index.max().date()

        st.subheader("Zeitraum")
        zeitraum = st.date_input(
            "Anzeigezeitraum (Diagramm & Auswertung)",
            value=(gesamt_min, gesamt_max),
            min_value=gesamt_min,
            max_value=gesamt_max,
            format="DD.MM.YYYY",
        )
        if isinstance(zeitraum, tuple) and len(zeitraum) == 2:
            start_datum, end_datum = zeitraum
        else:
            # Solange der Nutzer erst ein Datum ausgewählt hat, den vollen
            # Bereich als Fallback anzeigen.
            start_datum, end_datum = gesamt_min, gesamt_max

        ergebnis = ergebnis_gesamt.loc[
            (ergebnis_gesamt.index.date >= start_datum) & (ergebnis_gesamt.index.date <= end_datum)
        ]

        if len(ergebnis) == 0:
            st.warning("Für den gewählten Zeitraum liegen keine Daten vor.")
        else:
            st.subheader("Diagramm")
            fig = plot_tsb(ergebnis, bereiche_eff, atl_tage, ctl_tage)
            st.pyplot(fig, width="stretch")

            letzter_tag = ergebnis.iloc[-1]
            st.caption(
                f"Stand: **{letzter_tag.name.strftime('%d.%m.%Y')}** "
                f"(letzter Tag im oben gewählten Anzeigezeitraum). "
                "CTL/ATL sind gleitende Mittelwerte über die jeweilige "
                "Zeitkonstante (z.B. 42 bzw. 7 Tage) bis zu diesem Tag, "
                "TSB ist die Differenz aus dem CTL/ATL-Stand des Vortages."
            )
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("CTL (Fitness)", f"{letzter_tag['CTL']:.1f}")
            m2.metric("ATL (Fatigue)", f"{letzter_tag['ATL']:.1f}")
            m3.metric("TSB (Form)", f"{letzter_tag['TSB']:.1f}")
            m4.metric("Bereich", letzter_tag["TSB_Bereich"])

            rohdaten = df_kombiniert[
                (pd.to_datetime(df_kombiniert["Datum"]).dt.date >= start_datum)
                & (pd.to_datetime(df_kombiniert["Datum"]).dt.date <= end_datum)
            ].copy()
            rohdaten = rohdaten.sort_values(["Datum", "Dauer"])[["Datum", "Dauer", "TSS"]]
            rohdaten_anzeige = rohdaten.copy()
            rohdaten_anzeige["Dauer"] = rohdaten_anzeige["Dauer"].apply(dauer_anzeige)
            rohdaten_anzeige = rohdaten_anzeige.rename(columns={"Dauer": "Trainingszeit (hh:mm)"})
            with st.expander("Erfasste Trainingseinheiten", expanded=False):
                st.dataframe(rohdaten_anzeige, width="stretch", hide_index=True)

            st.subheader("Tagesauswertung (ATL / CTL / TSB)")

            # Trainingszeit als hh:mm-Spalte direkt nach Datum einfügen
            ergebnis_anzeige = ergebnis.round(2).copy()
            ergebnis_anzeige.insert(0, "Trainingszeit (hh:mm)", ergebnis_anzeige["Dauer_min"].apply(minuten_zu_hhmm))
            ergebnis_anzeige = ergebnis_anzeige.drop(columns=["Dauer_min"])

            # Fixierte Summenzeile oberhalb der (scrollbaren) Tabelle - eine
            # echte "frozen row" innerhalb einer einzelnen Tabelle bietet
            # Streamlit nicht an, daher als eigene, nicht scrollende Zeile
            # direkt über der Tabelle dargestellt.
            summe_dauer = minuten_zu_hhmm(ergebnis["Dauer_min"].sum())
            summe_tss = round(ergebnis["TSS"].sum(), 1)
            summenzeile = pd.DataFrame(
                [{
                    "Datum": "Summe",
                    "Trainingszeit (hh:mm)": summe_dauer,
                    "TSS": summe_tss,
                    "CTL": "–",
                    "ATL": "–",
                    "TSB": "–",
                    "TSB_Bereich": "–",
                }]
            )
            st.dataframe(summenzeile, width="stretch", hide_index=True)
            st.dataframe(
                zeilen_nach_tsb_bereich_faerben(ergebnis_anzeige, bereiche_eff),
                width="stretch",
            )
            st.caption("Die Zeilenfarbe entspricht dem TSB-Bereich des jeweiligen Tages (gleiche Farben wie die Zonen im TSB-Diagramm oben).")

            excel_puffer = pd.ExcelWriter("ergebnis_export.xlsx", engine="openpyxl")
            rohdaten_anzeige.to_excel(excel_puffer, sheet_name="Trainingseinheiten", index=False)
            ergebnis_anzeige.to_excel(excel_puffer, sheet_name="Tagesauswertung")
            excel_puffer.close()
            with open("ergebnis_export.xlsx", "rb") as f:
                st.download_button(
                    "Auswertung als Excel herunterladen",
                    data=f,
                    file_name="tss_atl_ctl_tsb_ergebnis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )





