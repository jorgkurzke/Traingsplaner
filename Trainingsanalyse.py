"""
TSS / ATL / CTL / TSB - Analyse und Visualisierung
====================================================

Liest tägliche TSS-Werte (Training Stress Score) aus einer Excel-Datei ein
und berechnet daraus:

  - ATL  (Acute Training Load / akute Belastung, "Fatigue")
  - CTL  (Chronic Training Load / chronische Belastung, "Fitness")
  - TSB  (Training Stress Balance, "Form")   TSB = CTL - ATL

ATL und CTL werden als exponentiell gleitender Mittelwert (EWMA) berechnet:

    Wert_heute = Wert_gestern + (TSS_heute - Wert_gestern) / Zeitkonstante

Die Zeitkonstanten sind frei wählbar (Standard: ATL = 7 Tage, CTL = 42 Tage,
klassische Coggan-Werte) - siehe Konfigurationsblock unten.

Die TSB-Bereiche (z.B. "Übertrainingsrisiko", "optimale Form", "Frische")
sind ebenfalls frei definierbar und werden im Diagramm als farbige Zonen
im Hintergrund dargestellt.

Benötigte Pakete: pandas, matplotlib, openpyxl
    pip install pandas matplotlib openpyxl --break-system-packages
"""

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates


# =====================================================================
# KONFIGURATION - hier alles anpassen
# =====================================================================

# ---- Eingabedatei ----
EXCEL_DATEI = "training_tss.xlsx"   # Pfad zur Excel-Datei
EXCEL_SHEET = 0                     # Blattname oder Index (0 = erstes Blatt)
SPALTE_DATUM = "Datum"              # Name der Spalte mit dem Datum
SPALTE_TSS = "TSS"                  # Name der Spalte mit dem TSS-Wert

# ---- Zeitkonstanten für ATL/CTL (frei wählbar, NICHT fix 7/42) ----
ATL_TAGE = 7          # Zeitkonstante ATL (z.B. 5-10 Tage üblich)
CTL_TAGE = 42          # Zeitkonstante CTL (z.B. 28-56 Tage üblich)

# ---- Startwerte am ersten Tag der Zeitreihe ----
CTL_START = 0.0
ATL_START = 0.0

# ---- TSB-Bereiche (frei definierbar) ----
# Liste von Dictionaries: von, bis, Label, Farbe (Hex).
# "von" ist inklusive, "bis" exklusiv. -inf/inf sind erlaubt (np.inf).
TSB_BEREICHE = [
    {"von": -np.inf, "bis": -30, "label": "Hohes Übertrainingsrisiko", "farbe": "#d03b3b"},
    {"von": -30,      "bis": -10, "label": "Ermüdung / Formaufbau",     "farbe": "#ec835a"},
    {"von": -10,      "bis": 5,   "label": "Optimale Form",             "farbe": "#0ca30c"},
    {"von": 5,        "bis": 25,  "label": "Frische / Taper",           "farbe": "#2a78d6"},
    {"von": 25,       "bis": np.inf, "label": "Formverlust (zu viel Ruhe)", "farbe": "#898781"},
]

# ---- Ausgabe ----
OUTPUT_EXCEL = "tss_atl_ctl_tsb_ergebnis.xlsx"
OUTPUT_PLOT = "tsb_verlauf.png"


# =====================================================================
# 1. Daten einlesen
# =====================================================================

def daten_einlesen(pfad, sheet, spalte_datum, spalte_tss):
    """
    Liest Datum + TSS aus Excel ein und liefert eine tägliche Zeitreihe
    zurück (lückenlos, fehlende Tage = TSS 0). Mehrere Einträge am selben
    Tag (z.B. zwei Einheiten) werden aufsummiert.
    """
    df = pd.read_excel(pfad, sheet_name=sheet)

    if spalte_datum not in df.columns or spalte_tss not in df.columns:
        raise ValueError(
            f"Erwartete Spalten '{spalte_datum}' und '{spalte_tss}' nicht "
            f"gefunden. Vorhandene Spalten: {list(df.columns)}"
        )

    df = df[[spalte_datum, spalte_tss]].copy()
    df[spalte_datum] = pd.to_datetime(df[spalte_datum])
    df[spalte_tss] = pd.to_numeric(df[spalte_tss], errors="coerce").fillna(0)

    # Mehrere Einheiten am selben Tag aufsummieren
    df = df.groupby(spalte_datum, as_index=True)[spalte_tss].sum()
    df = df.sort_index()

    # Lückenlose Tagesreihe erzeugen (fehlende Tage = 0 TSS), das ist
    # wichtig, damit ATL/CTL korrekt abklingen, wenn nicht trainiert wird.
    alle_tage = pd.date_range(df.index.min(), df.index.max(), freq="D")
    df = df.reindex(alle_tage, fill_value=0)
    df.index.name = "Datum"
    df.name = "TSS"

    return df.to_frame()


# =====================================================================
# 2. ATL, CTL, TSB berechnen
# =====================================================================

def berechne_atl_ctl_tsb(df, atl_tage=ATL_TAGE, ctl_tage=CTL_TAGE,
                          ctl_start=CTL_START, atl_start=ATL_START):
    """
    Berechnet ATL, CTL und TSB auf Basis der täglichen TSS-Werte.

    TSB des Tages t wird aus den Werten des VORTAGS berechnet (Standard-
    Konvention, z.B. TrainingPeaks): TSB_t = CTL_(t-1) - ATL_(t-1).
    Das bildet die Form ab, mit der man in den Tag t hineingeht, also vor
    der heutigen Einheit.
    """
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


# =====================================================================
# 3. TSB-Bereich je Tag bestimmen (optional, z.B. für Tabellenexport)
# =====================================================================

def tsb_bereich_zuordnen(tsb_wert, bereiche):
    for b in bereiche:
        if b["von"] <= tsb_wert < b["bis"]:
            return b["label"]
    return "unbekannt"


# =====================================================================
# 4. Diagramm: TSS/CTL/ATL oben, TSB mit farbigen Zonen unten
# =====================================================================

def plot_tsb(df, bereiche, output_pfad=OUTPUT_PLOT,
             atl_tage=ATL_TAGE, ctl_tage=CTL_TAGE):
    farbe_tss = "#c3c2b7"   # gedeckter Grauton für die TSS-Balken
    farbe_ctl = "#2a78d6"   # blau
    farbe_atl = "#eb6834"   # orange
    farbe_tsb = "#0b0b0b"   # fast schwarz

    fig, (ax_last, ax_tsb) = plt.subplots(
        2, 1, figsize=(13, 8), sharex=True,
        gridspec_kw={"height_ratios": [1, 1.3]},
    )
    fig.patch.set_facecolor("#fcfcfb")

    # ---- oberes Panel: TSS-Balken + CTL/ATL-Linien ----
    ax_last.bar(df.index, df["TSS"], color=farbe_tss, width=1.0,
                label="TSS (täglich)", zorder=1)
    ax_last.plot(df.index, df["CTL"], color=farbe_ctl, linewidth=2,
                 label=f"CTL (Fitness, {ctl_tage} Tage)", zorder=3)
    ax_last.plot(df.index, df["ATL"], color=farbe_atl, linewidth=2,
                 label=f"ATL (Fatigue, {atl_tage} Tage)", zorder=3)
    ax_last.set_ylabel("TSS / CTL / ATL")
    ax_last.set_facecolor("#fcfcfb")
    ax_last.legend(loc="upper left", frameon=False)
    ax_last.spines[["top", "right"]].set_visible(False)
    ax_last.spines[["left", "bottom"]].set_color("#c3c2b7")
    ax_last.grid(axis="y", color="#e1e0d9", linewidth=0.8)
    ax_last.set_title("Trainingsbelastung: TSS, CTL, ATL", loc="left",
                       color="#0b0b0b", fontsize=12, pad=10)

    # ---- unteres Panel: TSB mit farbigen Bereichen ----
    x_min, x_max = df.index.min(), df.index.max()
    for b in bereiche:
        untere = b["von"] if np.isfinite(b["von"]) else df["TSB"].min() - 10
        obere = b["bis"] if np.isfinite(b["bis"]) else df["TSB"].max() + 10
        ax_tsb.axhspan(untere, obere, color=b["farbe"], alpha=0.18,
                        zorder=0, label=b["label"])

    ax_tsb.plot(df.index, df["TSB"], color=farbe_tsb, linewidth=2,
                label="TSB (Form)", zorder=3)
    ax_tsb.axhline(0, color="#898781", linewidth=1, linestyle="--", zorder=2)

    ax_tsb.set_ylabel("TSB")
    ax_tsb.set_facecolor("#fcfcfb")
    ax_tsb.spines[["top", "right"]].set_visible(False)
    ax_tsb.spines[["left", "bottom"]].set_color("#c3c2b7")
    ax_tsb.grid(axis="y", color="#e1e0d9", linewidth=0.8)
    ax_tsb.set_title("Training Stress Balance (Form) mit definierten Bereichen",
                      loc="left", color="#0b0b0b", fontsize=12, pad=10)

    # Legende der TSB-Bereiche unterhalb der Linie, ohne die Linien-Legende
    # zu verdrängen
    handles, labels = ax_tsb.get_legend_handles_labels()
    ax_tsb.legend(handles, labels, loc="upper left", frameon=False,
                  fontsize=9, ncol=1)

    ax_tsb.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax_tsb.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%Y"))
    fig.autofmt_xdate()

    fig.tight_layout()
    fig.savefig(output_pfad, dpi=150, facecolor=fig.get_facecolor())
    plt.close(fig)
    print(f"Diagramm gespeichert: {output_pfad}")


# =====================================================================
# main
# =====================================================================

def main():
    df = daten_einlesen(EXCEL_DATEI, EXCEL_SHEET, SPALTE_DATUM, SPALTE_TSS)
    df = berechne_atl_ctl_tsb(df, atl_tage=ATL_TAGE, ctl_tage=CTL_TAGE,
                               ctl_start=CTL_START, atl_start=ATL_START)
    df["TSB_Bereich"] = df["TSB"].apply(lambda v: tsb_bereich_zuordnen(v, TSB_BEREICHE))

    df.round(2).to_excel(OUTPUT_EXCEL)
    print(f"Ergebnistabelle gespeichert: {OUTPUT_EXCEL}")

    plot_tsb(df, TSB_BEREICHE, OUTPUT_PLOT, atl_tage=ATL_TAGE, ctl_tage=CTL_TAGE)


if __name__ == "__main__":
    main()







