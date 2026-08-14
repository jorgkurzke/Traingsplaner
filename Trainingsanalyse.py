"""
Trainingsanalyse: ATL / CTL / TSB (Streamlit-App)
===================================================

Interaktive Streamlit-App zur Trainingssteuerung auf Basis von TSS
(Training Stress Score):

  - TSS-Werte (inkl. optionaler Trainingsdauer) per Excel-Import UND/ODER
    manuelle Eingabe erfassen; importierte Daten werden auf Wunsch
    dauerhaft in OneDrive gespeichert, ein erneutes Hochladen ist danach
    nicht mehr nötig (Einrichtung siehe ONEDRIVE_EINRICHTUNG.md)
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
    msal
    requests

Start lokal:  streamlit run Trainingsanalyse.py
"""

import calendar
import datetime as dt
import io
import json
import os
import time

import numpy as np
import pandas as pd
import requests
import streamlit as st
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

try:
    import msal
except ImportError:
    msal = None

# Datei, in der importierte Excel-Daten als Rückfall-Option gespeichert
# werden, FALLS OneDrive (noch) nicht eingerichtet ist (siehe unten). Liegt
# im Arbeitsverzeichnis der App und übersteht nur Reruns/Schlafmodus, aber
# kein Neu-Deployment.
GESPEICHERTE_IMPORT_DATEI = "gespeicherte_importe.csv"

# Pfad/Dateiname der Import-Daten im OneDrive des Nutzers (im Wurzel-
# verzeichnis, damit kein Ordner vorab angelegt werden muss).
ONEDRIVE_DATEIPFAD = "/me/drive/root:/gespeicherte_importe_trainingsplaner.csv"

# Einstellungen (Zeitkonstanten, Startwerte, TSB-Bereiche) - werden
# zusammen mit den Trainingsdaten gespeichert, ebenfalls im OneDrive-
# Wurzelverzeichnis bzw. lokal als Rückfall.
ONEDRIVE_EINSTELLUNGEN_PFAD = "/me/drive/root:/einstellungen_trainingsplaner.json"
GESPEICHERTE_EINSTELLUNGEN_DATEI = "gespeicherte_einstellungen.json"

GRAPH_SCOPES = ["Files.ReadWrite"]

STANDARD_TSB_BEREICHE = [
    {"von": -60.0, "bis": -30.0, "label": "Hohes Übertrainingsrisiko", "farbe": "#d03b3b"},
    {"von": -30.0, "bis": -10.0, "label": "Ermüdung / Formaufbau", "farbe": "#ec835a"},
    {"von": -10.0, "bis": 5.0, "label": "Optimale Form", "farbe": "#0ca30c"},
    {"von": 5.0, "bis": 25.0, "label": "Frische / Taper", "farbe": "#2a78d6"},
    {"von": 25.0, "bis": 60.0, "label": "Formverlust (zu viel Ruhe)", "farbe": "#898781"},
]

# Bewertungsbereiche für den systolischen Blutdruck (SYS), fest vorgegeben
# (nicht über die Oberfläche konfigurierbar). Grenzen so gewählt, dass sie
# nahtlos aneinander anschließen (keine Lücke zwischen den Bereichen):
# <=130 grün, 131-139 gelb, 140-159 orange, >159 (also auch >160) rot.
SYS_BEREICHE = [
    {"von": -np.inf, "bis": 130.0, "label": "SYS ≤ 130", "farbe": "#0ca30c"},
    {"von": 130.0, "bis": 139.0, "label": "SYS 131 - 139", "farbe": "#e8c93a"},
    {"von": 139.0, "bis": 159.0, "label": "SYS 140 - 159", "farbe": "#eb6834"},
    {"von": 159.0, "bis": np.inf, "label": "SYS > 159", "farbe": "#d03b3b"},
]


st.set_page_config(page_title="Trainingsanalyse: ATL / CTL / TSB", layout="wide")
st.title("Trainingsanalyse: ATL / CTL / TSB")
st.caption(
    "TSS-Werte importieren oder manuell erfassen, ATL/CTL/TSB berechnen und "
    "die Form (TSB) mit frei definierbaren Bereichen visualisieren."
)


# =====================================================================
# Persistenz für importierte Daten: OneDrive (Microsoft Graph API), mit
# automatischem Rückfall auf eine lokale Datei, falls OneDrive (noch)
# nicht in den Streamlit-Secrets eingerichtet ist. Einrichtung siehe
# ONEDRIVE_EINRICHTUNG.md.
# =====================================================================

def leere_eintraege():
    return pd.DataFrame(
        {
            "Datum": pd.Series(dtype="datetime64[ns]"),
            "Dauer": pd.Series(dtype="object"),
            "TSS": pd.Series(dtype="float"),
            "Training": pd.Series(dtype="object"),
            "Kg": pd.Series(dtype="float"),
            "SYS": pd.Series(dtype="float"),
            "DIA": pd.Series(dtype="float"),
            "Puls": pd.Series(dtype="float"),
        }
    )


def onedrive_konfiguriert():
    """Prüft, ob client_id + refresh_token in den Streamlit-Secrets unter
    [onedrive] hinterlegt sind."""
    if msal is None:
        return False
    try:
        return "onedrive" in st.secrets and {"client_id", "refresh_token"} <= set(st.secrets["onedrive"].keys())
    except Exception:
        return False


def _onedrive_access_token():
    """Holt ein aktuelles Graph-Access-Token über den in den Secrets
    hinterlegten Refresh Token. Wird pro Sitzung zwischengespeichert, damit
    nicht bei jedem Rerun ein neuer Netzwerk-Aufruf nötig ist."""
    jetzt = time.time()
    cache = st.session_state.get("_onedrive_token_cache")
    if cache and cache["ablauf"] > jetzt + 60:
        return cache["access_token"]

    client_id = st.secrets["onedrive"]["client_id"]
    refresh_token = st.secrets["onedrive"]["refresh_token"]
    authority = st.secrets["onedrive"].get("authority", "https://login.microsoftonline.com/consumers")

    app = msal.PublicClientApplication(client_id, authority=authority)
    ergebnis = app.acquire_token_by_refresh_token(refresh_token, scopes=GRAPH_SCOPES)
    if "access_token" not in ergebnis:
        raise RuntimeError(ergebnis.get("error_description", str(ergebnis)))

    st.session_state["_onedrive_token_cache"] = {
        "access_token": ergebnis["access_token"],
        "ablauf": jetzt + ergebnis.get("expires_in", 3600),
    }
    return ergebnis["access_token"]


def _onedrive_datei_lesen(pfad):
    token = _onedrive_access_token()
    antwort = requests.get(
        f"https://graph.microsoft.com/v1.0{pfad}:/content",
        headers={"Authorization": f"Bearer {token}"},
        timeout=20,
    )
    if antwort.status_code == 404:
        return None  # Datei existiert noch nicht - erster Aufruf
    antwort.raise_for_status()
    return antwort.content


def _onedrive_datei_schreiben(pfad, inhalt_bytes, content_type="application/octet-stream"):
    token = _onedrive_access_token()
    antwort = requests.put(
        f"https://graph.microsoft.com/v1.0{pfad}:/content",
        headers={"Authorization": f"Bearer {token}", "Content-Type": content_type},
        data=inhalt_bytes,
        timeout=20,
    )
    antwort.raise_for_status()


def _eintraege_normalisieren(df):
    """Stellt sicher, dass eine geladene Tabelle immer die Spalten
    Datum/Dauer/TSS/Training/Kg/SYS/DIA/Puls enthält (schützt vor Abstürzen
    bei älteren oder von Hand bearbeiteten gespeicherten Dateien, denen z.B.
    die Dauer-Spalte oder die später hinzugekommenen Spalten fehlen)."""
    df = df.copy()
    if "Dauer" not in df.columns:
        df["Dauer"] = None
    if "TSS" not in df.columns:
        df["TSS"] = 0.0
    if "Training" not in df.columns:
        df["Training"] = None
    for _spalte in ["Kg", "SYS", "DIA", "Puls"]:
        if _spalte not in df.columns:
            df[_spalte] = np.nan
    df["Datum"] = pd.to_datetime(df["Datum"])
    return df[["Datum", "Dauer", "TSS", "Training", "Kg", "SYS", "DIA", "Puls"]]


def importierte_daten_laden():
    """Lädt zuvor gespeicherte Import-Daten - aus OneDrive, falls
    eingerichtet, sonst als Rückfall aus einer lokalen Datei."""
    if onedrive_konfiguriert():
        try:
            inhalt = _onedrive_datei_lesen(ONEDRIVE_DATEIPFAD)
            if inhalt is None:
                return leere_eintraege()
            return _eintraege_normalisieren(pd.read_csv(io.BytesIO(inhalt)))
        except Exception as e:
            st.warning(f"Gespeicherte Daten konnten nicht aus OneDrive geladen werden ({e}).")
            return leere_eintraege()

    if os.path.exists(GESPEICHERTE_IMPORT_DATEI):
        try:
            return _eintraege_normalisieren(pd.read_csv(GESPEICHERTE_IMPORT_DATEI))
        except Exception:
            pass
    return leere_eintraege()


def importierte_daten_speichern(df):
    """Schreibt die aktuellen Import-Daten dauerhaft weg - nach OneDrive,
    falls eingerichtet, sonst als Rückfall in eine lokale Datei (übersteht
    dann nur Reruns/Schlafmodus, kein Neu-Deployment)."""
    if onedrive_konfiguriert():
        try:
            puffer = io.StringIO()
            df.to_csv(puffer, index=False)
            _onedrive_datei_schreiben(ONEDRIVE_DATEIPFAD, puffer.getvalue().encode("utf-8"), "text/csv")
            return True
        except Exception as e:
            st.error(
                f"Speichern in OneDrive fehlgeschlagen ({e}). Die Daten wurden "
                "stattdessen nur lokal gespeichert (übersteht kein Neu-Deployment)."
            )
    df.to_csv(GESPEICHERTE_IMPORT_DATEI, index=False)
    return False


def einstellungen_laden():
    """Lädt gespeicherte Einstellungen (Zeitkonstanten, Startwerte,
    TSB-Bereiche) - aus OneDrive, falls eingerichtet, sonst als Rückfall
    aus einer lokalen Datei. Fehlt beides, gelten die Standardwerte."""
    standard = {
        "atl_tage": 7,
        "ctl_tage": 42,
        "ctl_start": 0.0,
        "atl_start": 0.0,
        "tsb_bereiche": STANDARD_TSB_BEREICHE,
    }
    geladen = None

    if onedrive_konfiguriert():
        try:
            inhalt = _onedrive_datei_lesen(ONEDRIVE_EINSTELLUNGEN_PFAD)
            if inhalt is not None:
                geladen = json.loads(inhalt.decode("utf-8"))
        except Exception as e:
            st.warning(f"Gespeicherte Einstellungen konnten nicht aus OneDrive geladen werden ({e}).")

    if geladen is None and os.path.exists(GESPEICHERTE_EINSTELLUNGEN_DATEI):
        try:
            with open(GESPEICHERTE_EINSTELLUNGEN_DATEI, "r", encoding="utf-8") as f:
                geladen = json.load(f)
        except Exception:
            geladen = None

    if geladen:
        standard.update(geladen)
    return standard


def einstellungen_speichern(einstellungen):
    """Schreibt die aktuellen Einstellungen dauerhaft weg - nach OneDrive,
    falls eingerichtet, sonst als Rückfall in eine lokale Datei."""
    inhalt_bytes = json.dumps(einstellungen, ensure_ascii=False, indent=2).encode("utf-8")

    if onedrive_konfiguriert():
        try:
            _onedrive_datei_schreiben(ONEDRIVE_EINSTELLUNGEN_PFAD, inhalt_bytes, "application/json")
            return True
        except Exception as e:
            st.error(
                f"Speichern der Einstellungen in OneDrive fehlgeschlagen ({e}). "
                "Sie wurden stattdessen nur lokal gespeichert (übersteht kein Neu-Deployment)."
            )

    with open(GESPEICHERTE_EINSTELLUNGEN_DATEI, "w", encoding="utf-8") as f:
        f.write(inhalt_bytes.decode("utf-8"))
    return False


# =====================================================================
# Session State initialisieren
# =====================================================================

if "importierte_eintraege" not in st.session_state:
    st.session_state.importierte_eintraege = importierte_daten_laden()

if "manuelle_eintraege" not in st.session_state:
    st.session_state.manuelle_eintraege = pd.DataFrame(
        {
            "Datum": pd.Series(dtype="datetime64[ns]"),
            "Dauer": pd.Series(dtype="object"),  # Trainingsdauer als datetime.time (hh:mm)
            "TSS": pd.Series(dtype="float"),
            "Training": pd.Series(dtype="object"),
            "Kg": pd.Series(dtype="float"),
            "SYS": pd.Series(dtype="float"),
            "DIA": pd.Series(dtype="float"),
            "Puls": pd.Series(dtype="float"),
        }
    )

if "einstellungen" not in st.session_state:
    st.session_state.einstellungen = einstellungen_laden()

if "tsb_bereiche" not in st.session_state:
    # "von" des ersten und "bis" des letzten Bereichs werden automatisch auf
    # -unendlich / +unendlich erweitert (siehe effektive_bereiche()) - hier
    # reichen die inneren Grenzwerte.
    st.session_state.tsb_bereiche = pd.DataFrame(st.session_state.einstellungen["tsb_bereiche"])


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


def monate_subtrahieren(datum, anzahl_monate):
    """Zieht von einem Datum eine Anzahl Kalendermonate ab (Tag wird ggf.
    auf den letzten gültigen Tag des Zielmonats begrenzt, z.B. 31.03. minus
    1 Monat -> 28./29.02.)."""
    monat_index = datum.month - 1 - anzahl_monate
    jahr = datum.year + monat_index // 12
    monat = monat_index % 12 + 1
    letzter_tag_im_monat = calendar.monthrange(jahr, monat)[1]
    tag = min(datum.day, letzter_tag_im_monat)
    return dt.date(jahr, monat, tag)


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


def _excel_zahlenformate_setzen(worksheet, formate):
    """Setzt für die per Spaltenname (in der Kopfzeile) angegebenen Spalten
    ein Excel-Zahlenformat. Die Werte müssen dafür bereits als echte
    Zahlen (float) geschrieben worden sein - Excel zeigt sie dann als
    Zahl an und übernimmt automatisch das Dezimaltrennzeichen der
    Excel-Spracheinstellung (i.d.R. Komma bei deutschem Excel), statt
    fest einprogrammierter Zeichen."""
    for zelle in worksheet[1]:
        format_code = formate.get(zelle.value)
        if not format_code:
            continue
        buchstabe = zelle.column_letter
        for zeile in range(2, worksheet.max_row + 1):
            worksheet[f"{buchstabe}{zeile}"].number_format = format_code


def zahl_anzeige(wert, nachkommastellen=1):
    """Formatiert einen optionalen Zahlenwert (z.B. Kg/SYS/DIA/Puls) mit
    fester Anzahl Nachkommastellen, leer statt '0.0' falls kein Wert
    vorhanden ist."""
    try:
        if wert is None or pd.isna(wert):
            return ""
    except (TypeError, ValueError):
        pass
    return f"{float(wert):.{nachkommastellen}f}"


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
        farbe = farb_map.get(zeile.get("TSB_Bereich", zeile.get("TSB-Bereich")))
        if farbe:
            return [f"background-color: {hex_zu_rgba(farbe)}"] * len(zeile)
        return [""] * len(zeile)

    return tabelle.style.apply(zeile_stylen, axis=1)


def _training_zusammenfassen(werte):
    """Fasst mehrere Trainingsbezeichnungen desselben Tages zusammen (z.B.
    zwei Einheiten am selben Tag) - eindeutige, nicht-leere Werte, mit
    ' + ' verbunden."""
    eindeutig = [str(w).strip() for w in werte if pd.notna(w) and str(w).strip() != ""]
    eindeutig = list(dict.fromkeys(eindeutig))  # Reihenfolge erhalten, Duplikate entfernen
    return " + ".join(eindeutig)


def tagesreihe_aufbauen(df_tss):
    """Nimmt ein DataFrame mit Spalten Datum/Dauer/TSS/Training/Kg/SYS/DIA/
    Puls, summiert bzw. mittelt Mehrfacheinträge pro Tag und füllt fehlende
    Tage mit TSS=0/Dauer=0 (lückenlose Tagesreihe - wichtig, damit ATL/CTL
    bei Trainingspausen korrekt abklingen). Kg/SYS/DIA/Puls werden bei
    fehlenden Tagen NICHT mit 0 aufgefüllt (bleiben leer/NaN), da es sich
    um Messwerte und keine kumulierbaren Belastungsgrößen handelt."""
    df = df_tss.dropna(subset=["Datum"]).copy()
    df["Datum"] = pd.to_datetime(df["Datum"])
    df["TSS"] = pd.to_numeric(df["TSS"], errors="coerce").fillna(0)
    df["Dauer_min"] = df["Dauer"].apply(dauer_zu_minuten) if "Dauer" in df.columns else 0.0
    if "Training" not in df.columns:
        df["Training"] = None
    for _spalte in ["Kg", "SYS", "DIA", "Puls"]:
        if _spalte not in df.columns:
            df[_spalte] = np.nan
        else:
            df[_spalte] = pd.to_numeric(df[_spalte], errors="coerce")
    if len(df) == 0:
        return None
    tag_gruppe = df.groupby("Datum").agg(
        TSS=("TSS", "sum"),
        Dauer_min=("Dauer_min", "sum"),
        Training=("Training", _training_zusammenfassen),
        Kg=("Kg", "mean"),
        SYS=("SYS", "mean"),
        DIA=("DIA", "mean"),
        Puls=("Puls", "mean"),
    ).sort_index()
    alle_tage = pd.date_range(tag_gruppe.index.min(), tag_gruppe.index.max(), freq="D")
    tag_gruppe = tag_gruppe.reindex(alle_tage)
    tag_gruppe["TSS"] = tag_gruppe["TSS"].fillna(0)
    tag_gruppe["Dauer_min"] = tag_gruppe["Dauer_min"].fillna(0)
    tag_gruppe["Training"] = tag_gruppe["Training"].fillna("")
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


def plot_kg_sys(df, sys_bereiche):
    """Balkengrafik Kg (oben) + Balkengrafik SYS mit farbig hinterlegten
    Bewertungsbereichen (unten). Tage ohne Messwert bleiben in beiden
    Grafiken einfach leer (keine 0-Werte)."""
    farbe_kg = "#2a78d6"
    farbe_sys = "#8a8a8a"

    fig, (ax_kg, ax_sys) = plt.subplots(
        2, 1, figsize=(13, 7), sharex=True, gridspec_kw={"height_ratios": [1, 1.3]}
    )
    fig.patch.set_facecolor("#fcfcfb")

    ax_kg.bar(df.index, df["Kg"], color=farbe_kg, width=1.0, label="Kg", zorder=3)
    ax_kg.set_ylabel("Kg")
    ax_kg.set_facecolor("#fcfcfb")
    ax_kg.legend(loc="upper left", frameon=False)
    ax_kg.spines[["top", "right"]].set_visible(False)
    ax_kg.spines[["left", "bottom"]].set_color("#c3c2b7")
    ax_kg.grid(axis="y", color="#e1e0d9", linewidth=0.8)
    if df["Kg"].notna().any():
        kg_min, kg_max = df["Kg"].min(), df["Kg"].max()
        spanne = max(kg_max - kg_min, 1.0)
        ax_kg.set_ylim(kg_min - spanne * 0.2, kg_max + spanne * 0.2)
    ax_kg.set_title("Körpergewicht (Kg)", loc="left", color="#0b0b0b", fontsize=12, pad=10)

    sys_min = df["SYS"].min() if df["SYS"].notna().any() else 90.0
    sys_max = df["SYS"].max() if df["SYS"].notna().any() else 180.0
    for b in sys_bereiche:
        untere = b["von"] if np.isfinite(b["von"]) else sys_min - 10
        obere = b["bis"] if np.isfinite(b["bis"]) else sys_max + 10
        ax_sys.axhspan(untere, obere, color=b["farbe"], alpha=0.22, zorder=0, label=b["label"])

    ax_sys.bar(df.index, df["SYS"], color=farbe_sys, width=1.0, label="SYS", zorder=3)
    ax_sys.set_ylabel("SYS")
    ax_sys.set_facecolor("#fcfcfb")
    ax_sys.spines[["top", "right"]].set_visible(False)
    ax_sys.spines[["left", "bottom"]].set_color("#c3c2b7")
    ax_sys.grid(axis="y", color="#e1e0d9", linewidth=0.8)
    ax_sys.set_ylim(sys_min - 10, sys_max + 10)
    ax_sys.set_title("Blutdruck systolisch (SYS) mit Bewertungsbereichen", loc="left",
                      color="#0b0b0b", fontsize=12, pad=10)

    handles, labels = ax_sys.get_legend_handles_labels()
    ax_sys.legend(handles, labels, loc="upper left", frameon=False, fontsize=9, ncol=1)

    ax_sys.xaxis.set_major_locator(mdates.AutoDateLocator())
    ax_sys.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%Y"))
    fig.autofmt_xdate()
    fig.tight_layout()
    return fig


# =====================================================================
# Sidebar: Einstellungen
# =====================================================================

with st.sidebar:
    st.header("Einstellungen")

    _e = st.session_state.einstellungen

    atl_tage = st.number_input(
        "ATL-Zeitkonstante (Tage)", min_value=1, max_value=200, value=int(_e["atl_tage"]), step=1
    )
    ctl_tage = st.number_input(
        "CTL-Zeitkonstante (Tage)", min_value=1, max_value=200, value=int(_e["ctl_tage"]), step=1
    )

    with st.expander("Erweitert: Startwerte"):
        ctl_start = st.number_input("CTL-Startwert", value=float(_e["ctl_start"]), step=1.0)
        atl_start = st.number_input("ATL-Startwert", value=float(_e["atl_start"]), step=1.0)

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

    if onedrive_konfiguriert():
        st.caption("Werden zusammen mit den Trainingsdaten in OneDrive gespeichert.")
    else:
        st.caption(
            "OneDrive ist noch nicht eingerichtet - Einstellungen werden bis "
            "dahin nur lokal zwischengespeichert (siehe ONEDRIVE_EINRICHTUNG.md)."
        )

    if st.button("Einstellungen speichern", width="stretch"):
        neue_einstellungen = {
            "atl_tage": atl_tage,
            "ctl_tage": ctl_tage,
            "ctl_start": ctl_start,
            "atl_start": atl_start,
            "tsb_bereiche": st.session_state.tsb_bereiche.to_dict("records"),
        }
        st.session_state.einstellungen = neue_einstellungen
        in_onedrive = einstellungen_speichern(neue_einstellungen)
        ziel = "OneDrive" if in_onedrive else "der App (lokal, ohne OneDrive-Einrichtung)"
        st.success(f"Einstellungen in {ziel} gespeichert.")


# =====================================================================
# Datenerfassung: Excel-Import + manuelle Eingabe
# =====================================================================

col_import, col_manuell = st.columns(2)

with col_import:
    st.subheader("1. Excel-Import")

    if onedrive_konfiguriert():
        st.caption("OneDrive-Speicherung aktiv.")
    else:
        st.caption(
            "OneDrive ist noch nicht eingerichtet (siehe ONEDRIVE_EINRICHTUNG.md) - "
            "importierte Daten werden bis dahin nur lokal auf der App-Festplatte "
            "zwischengespeichert (übersteht kein Neu-Deployment)."
        )

    if len(st.session_state.importierte_eintraege) > 0:
        st.caption(
            f"Aktuell {len(st.session_state.importierte_eintraege)} importierte Zeilen "
            "dauerhaft gespeichert - dafür muss keine Excel-Datei mehr hochgeladen werden."
        )

    hochgeladene_datei = st.file_uploader("Excel-Datei mit TSS-Werten", type=["xlsx", "xls"])

    df_import = leere_eintraege()

    if hochgeladene_datei is not None:
        try:
            excel_datei = pd.ExcelFile(hochgeladene_datei)
            sheet = st.selectbox("Tabellenblatt", excel_datei.sheet_names)
            df_roh = excel_datei.parse(sheet)

            spalten = list(df_roh.columns)
            datum_default = next((c for c in spalten if "datum" in str(c).lower() or "date" in str(c).lower()), spalten[0])
            tss_default = next((c for c in spalten if "tss" in str(c).lower()), spalten[-1])
            dauer_default = next((c for c in spalten if "dauer" in str(c).lower() or "duration" in str(c).lower() or "zeit" in str(c).lower()), None)
            training_default = next((c for c in spalten if "training" in str(c).lower() or "sport" in str(c).lower() or "übung" in str(c).lower()), None)
            kg_default = next((c for c in spalten if str(c).strip().lower() in ("kg", "gewicht", "weight")), None)
            sys_default = next((c for c in spalten if "sys" in str(c).lower()), None)
            dia_default = next((c for c in spalten if "dia" in str(c).lower()), None)
            puls_default = next((c for c in spalten if "puls" in str(c).lower() or "pulse" in str(c).lower() or "hr" in str(c).lower()), None)

            c1, c2, c3 = st.columns(3)
            with c1:
                spalte_datum = st.selectbox("Spalte mit Datum", spalten, index=spalten.index(datum_default))
            with c2:
                spalte_tss = st.selectbox("Spalte mit TSS", spalten, index=spalten.index(tss_default))
            with c3:
                dauer_optionen = ["(keine)"] + spalten
                dauer_index = dauer_optionen.index(dauer_default) if dauer_default else 0
                spalte_dauer = st.selectbox("Spalte mit Trainingszeit (optional, hh:mm)", dauer_optionen, index=dauer_index)

            st.caption("Optional: Spalten für Training, Kg, SYS, DIA, Puls zuordnen (falls in der Excel-Datei vorhanden):")
            d1, d2, d3, d4, d5 = st.columns(5)
            optionen = ["(keine)"] + spalten
            with d1:
                spalte_training = st.selectbox(
                    "Spalte mit Training", optionen, index=optionen.index(training_default) if training_default else 0
                )
            with d2:
                spalte_kg = st.selectbox("Spalte mit Kg", optionen, index=optionen.index(kg_default) if kg_default else 0)
            with d3:
                spalte_sys = st.selectbox("Spalte mit SYS", optionen, index=optionen.index(sys_default) if sys_default else 0)
            with d4:
                spalte_dia = st.selectbox("Spalte mit DIA", optionen, index=optionen.index(dia_default) if dia_default else 0)
            with d5:
                spalte_puls = st.selectbox("Spalte mit Puls", optionen, index=optionen.index(puls_default) if puls_default else 0)

            df_import = df_roh[[spalte_datum, spalte_tss]].rename(columns={spalte_datum: "Datum", spalte_tss: "TSS"})
            if spalte_dauer != "(keine)":
                df_import["Dauer"] = df_roh[spalte_dauer]
            else:
                df_import["Dauer"] = None
            for _ziel, _quelle in [
                ("Training", spalte_training),
                ("Kg", spalte_kg),
                ("SYS", spalte_sys),
                ("DIA", spalte_dia),
                ("Puls", spalte_puls),
            ]:
                df_import[_ziel] = df_roh[_quelle] if _quelle != "(keine)" else None
            st.success(f"{len(df_import)} Zeilen aus '{hochgeladene_datei.name}' eingelesen.")

            st.caption(
                "Achtung: Beim Speichern werden ALLE bisher gespeicherten "
                "importierten Trainingsdaten gelöscht und vollständig durch "
                "die Zeilen aus dieser Datei ersetzt (kein Zusammenführen "
                "mit vorherigen Importen)."
            )
            if st.button("Importierte Daten dauerhaft in der App speichern", width="stretch"):
                # Erst alle bisher gespeicherten Import-Daten löschen, dann
                # die neu importierten Daten als alleinigen, vollständigen
                # Datenbestand speichern. Das verhindert doppelte/mehrfach
                # auftauchende Tage, die durch ein Zusammenführen mit dem
                # alten Bestand entstehen könnten.
                kombiniert = (
                    df_import.dropna(subset=["Datum"])
                    .assign(Datum=lambda d: pd.to_datetime(d["Datum"]))
                    .sort_values("Datum")
                    .reset_index(drop=True)
                )

                st.session_state.importierte_eintraege = kombiniert
                in_onedrive_gespeichert = importierte_daten_speichern(kombiniert)
                ziel = "OneDrive" if in_onedrive_gespeichert else "der App (lokal, ohne OneDrive-Einrichtung)"
                st.success(
                    f"Bisherige gespeicherte Import-Daten gelöscht und durch die neuen Daten "
                    f"ersetzt. In {ziel} gespeichert. Insgesamt {len(kombiniert)} importierte "
                    "Zeilen liegen jetzt dauerhaft vor - die Excel-Datei muss dafür nicht "
                    "erneut hochgeladen werden."
                )
        except Exception as e:
            st.error(f"Datei konnte nicht gelesen werden: {e}")

    if len(st.session_state.importierte_eintraege) > 0:
        with st.expander(f"Gespeicherte Import-Daten verwalten ({len(st.session_state.importierte_eintraege)} Zeilen)"):
            importierte_anzeige = st.session_state.importierte_eintraege.copy()
            importierte_anzeige["Datum"] = pd.to_datetime(importierte_anzeige["Datum"]).dt.strftime("%d.%m.%Y")
            importierte_anzeige["Dauer"] = importierte_anzeige["Dauer"].apply(dauer_anzeige)
            importierte_anzeige["TSS"] = importierte_anzeige["TSS"].apply(lambda v: f"{v:.1f}")
            for _spalte in ["Kg", "SYS", "DIA", "Puls"]:
                importierte_anzeige[_spalte] = importierte_anzeige[_spalte].apply(zahl_anzeige)
            importierte_anzeige["Training"] = importierte_anzeige["Training"].fillna("")
            importierte_anzeige = importierte_anzeige.rename(columns={"Dauer": "Trainingszeit (hh:mm)"})
            st.dataframe(importierte_anzeige, width="stretch", hide_index=True)
            if st.button("Gespeicherte Import-Daten löschen", width="stretch"):
                st.session_state.importierte_eintraege = leere_eintraege()
                importierte_daten_speichern(leere_eintraege())  # überschreibt OneDrive/lokale Datei mit leerer Tabelle
                if os.path.exists(GESPEICHERTE_IMPORT_DATEI):
                    os.remove(GESPEICHERTE_IMPORT_DATEI)
                st.success("Gespeicherte Import-Daten gelöscht.")
                st.rerun()

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

        st.caption("Optional: Training, Körpergewicht und Blutdruck/Puls:")
        fc4, fc5, fc6, fc7, fc8 = st.columns(5)
        with fc4:
            neues_training = st.text_input("Training", value="")
        with fc5:
            neues_kg = st.number_input("Kg", min_value=0.0, step=0.1, value=0.0, format="%.1f")
        with fc6:
            neuer_sys = st.number_input("SYS", min_value=0.0, step=1.0, value=0.0)
        with fc7:
            neuer_dia = st.number_input("DIA", min_value=0.0, step=1.0, value=0.0)
        with fc8:
            neuer_puls = st.number_input("Puls", min_value=0.0, step=1.0, value=0.0)

        abgeschickt = st.form_submit_button("Einheit hinzufügen", width="stretch")

    if abgeschickt:
        neue_dauer_hhmm = f"{neue_dauer.hour:02d}:{neue_dauer.minute:02d}"
        neue_zeile = pd.DataFrame(
            [{
                "Datum": pd.Timestamp(neues_datum),
                "Dauer": neue_dauer_hhmm,
                "TSS": neuer_tss,
                "Training": neues_training.strip() if neues_training else None,
                "Kg": neues_kg if neues_kg > 0 else np.nan,
                "SYS": neuer_sys if neuer_sys > 0 else np.nan,
                "DIA": neuer_dia if neuer_dia > 0 else np.nan,
                "Puls": neuer_puls if neuer_puls > 0 else np.nan,
            }]
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
            "Training": st.column_config.TextColumn("Training"),
            "Kg": st.column_config.NumberColumn("Kg", step=0.1, format="%.1f"),
            "SYS": st.column_config.NumberColumn("SYS", step=1.0),
            "DIA": st.column_config.NumberColumn("DIA", step=1.0),
            "Puls": st.column_config.NumberColumn("Puls", step=1.0),
        },
        column_order=["Datum", "Dauer", "TSS", "Training", "Kg", "SYS", "DIA", "Puls"],
        key="manuelle_eintraege_editor",
    )


# =====================================================================
# Daten kombinieren, berechnen, anzeigen
# =====================================================================

df_kombiniert = pd.concat(
    [st.session_state.importierte_eintraege, st.session_state.manuelle_eintraege], ignore_index=True
)
df_kombiniert = df_kombiniert.dropna(subset=["Datum"])

st.divider()

if len(df_kombiniert) == 0:
    st.info(
        "Noch keine Daten vorhanden. Excel-Datei importieren und mit dem Button "
        "speichern, oder oben rechts manuell TSS-Werte eintragen."
    )
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
        heute = dt.date.today()

        st.subheader("Zeitraum")
        ZEITRAUM_OPTIONEN = [
            "Aktuelles Jahr",
            "Letzte 6 Monate",
            "Letzte 3 Monate",
            "Letzte 6 Wochen",
            "Gesamter Zeitraum",
            "Freie Eingabe",
        ]
        zeitraum_wahl = st.selectbox(
            "Anzeigezeitraum (Diagramm & Auswertung)", ZEITRAUM_OPTIONEN, index=4
        )

        if zeitraum_wahl == "Aktuelles Jahr":
            start_datum, end_datum = dt.date(heute.year, 1, 1), heute
        elif zeitraum_wahl == "Letzte 6 Monate":
            start_datum, end_datum = monate_subtrahieren(heute, 6), heute
        elif zeitraum_wahl == "Letzte 3 Monate":
            start_datum, end_datum = monate_subtrahieren(heute, 3), heute
        elif zeitraum_wahl == "Letzte 6 Wochen":
            start_datum, end_datum = heute - dt.timedelta(weeks=6), heute
        elif zeitraum_wahl == "Gesamter Zeitraum":
            start_datum, end_datum = gesamt_min, gesamt_max
        else:  # Freie Eingabe
            zeitraum = st.date_input(
                "Von / bis",
                value=(gesamt_min, gesamt_max),
                min_value=gesamt_min,
                max_value=gesamt_max,
                format="DD.MM.YYYY",
            )
            if isinstance(zeitraum, tuple) and len(zeitraum) == 2:
                start_datum, end_datum = zeitraum
            else:
                # Solange der Nutzer erst ein Datum ausgewählt hat, den
                # vollen Bereich als Fallback anzeigen.
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

            if ergebnis["Kg"].notna().any() or ergebnis["SYS"].notna().any():
                st.subheader("Körpergewicht (Kg) & Blutdruck systolisch (SYS)")
                fig_kg_sys = plot_kg_sys(ergebnis, SYS_BEREICHE)
                st.pyplot(fig_kg_sys, width="stretch")

            rohdaten = df_kombiniert[
                (pd.to_datetime(df_kombiniert["Datum"]).dt.date >= start_datum)
                & (pd.to_datetime(df_kombiniert["Datum"]).dt.date <= end_datum)
            ].copy()
            rohdaten = rohdaten.sort_values(["Datum", "Dauer"])[
                ["Datum", "Training", "Dauer", "TSS", "Kg", "SYS", "DIA", "Puls"]
            ]
            rohdaten_anzeige = rohdaten.copy()
            rohdaten_anzeige["Dauer"] = rohdaten_anzeige["Dauer"].apply(dauer_anzeige)
            rohdaten_anzeige["TSS"] = rohdaten_anzeige["TSS"].apply(lambda v: f"{v:.1f}")
            for _spalte in ["Kg", "SYS", "DIA", "Puls"]:
                rohdaten_anzeige[_spalte] = rohdaten_anzeige[_spalte].apply(zahl_anzeige)
            rohdaten_anzeige["Training"] = rohdaten_anzeige["Training"].fillna("")
            rohdaten_anzeige["Datum"] = pd.to_datetime(rohdaten_anzeige["Datum"]).dt.strftime("%d.%m.%Y")
            rohdaten_anzeige = rohdaten_anzeige.rename(columns={"Dauer": "Trainingszeit (hh:mm)"})
            with st.expander("Erfasste Trainingseinheiten", expanded=False):
                st.dataframe(rohdaten_anzeige, width="stretch", hide_index=True)

            st.subheader("Tagesauswertung (ATL / CTL / TSB)")

            # Spalten in der vom Nutzer festgelegten Reihenfolge aufbauen:
            # Datum, Training, Trainingszeit, TSS, ATL, CTL, TSB,
            # TSB-Bereich, Kg, SYS, DIA, Puls. TSS/ATL/CTL/TSB jeweils mit
            # genau einer Dezimalen als Text formatiert (garantiert
            # einheitliche Darstellung), Datum als tt.mm.jjjj.
            ergebnis_anzeige = ergebnis.copy()
            ergebnis_anzeige["Trainingszeit (hh:mm)"] = ergebnis_anzeige["Dauer_min"].apply(minuten_zu_hhmm)
            ergebnis_anzeige["Training"] = ergebnis_anzeige["Training"].fillna("")
            for _spalte in ["TSS", "ATL", "CTL", "TSB"]:
                ergebnis_anzeige[_spalte] = ergebnis_anzeige[_spalte].apply(lambda v: f"{v:.1f}")
            for _spalte in ["Kg", "SYS", "DIA", "Puls"]:
                ergebnis_anzeige[_spalte] = ergebnis_anzeige[_spalte].apply(zahl_anzeige)
            ergebnis_anzeige = ergebnis_anzeige.rename(columns={"TSB_Bereich": "TSB-Bereich"})
            ergebnis_anzeige.index = ergebnis_anzeige.index.strftime("%d.%m.%Y")
            ergebnis_anzeige.index.name = "Datum"
            ergebnis_anzeige = ergebnis_anzeige[
                ["Training", "Trainingszeit (hh:mm)", "TSS", "ATL", "CTL", "TSB", "TSB-Bereich",
                 "Kg", "SYS", "DIA", "Puls"]
            ]

            # Fixierte Summenzeile oberhalb der (scrollbaren) Tabelle - eine
            # echte "frozen row" innerhalb einer einzelnen Tabelle bietet
            # Streamlit nicht an, daher als eigene, nicht scrollende Zeile
            # direkt über der Tabelle dargestellt. Kg/SYS/DIA/Puls werden
            # hier als Mittelwerte (nicht Summen) ausgewiesen.
            summe_dauer = minuten_zu_hhmm(ergebnis["Dauer_min"].sum())
            summe_tss = f"{ergebnis['TSS'].sum():.1f}"
            summenzeile = pd.DataFrame(
                [{
                    "Datum": "Summe",
                    "Training": "–",
                    "Trainingszeit (hh:mm)": summe_dauer,
                    "TSS": summe_tss,
                    "ATL": "–",
                    "CTL": "–",
                    "TSB": "–",
                    "TSB-Bereich": "–",
                    "Kg": zahl_anzeige(ergebnis["Kg"].mean()),
                    "SYS": zahl_anzeige(ergebnis["SYS"].mean()),
                    "DIA": zahl_anzeige(ergebnis["DIA"].mean()),
                    "Puls": zahl_anzeige(ergebnis["Puls"].mean()),
                }]
            )
            st.dataframe(summenzeile, width="stretch", hide_index=True)
            st.dataframe(
                zeilen_nach_tsb_bereich_faerben(ergebnis_anzeige, bereiche_eff),
                width="stretch",
            )
            st.caption("Die Zeilenfarbe entspricht dem TSB-Bereich des jeweiligen Tages (gleiche Farben wie die Zonen im TSB-Diagramm oben).")

            # Für den Excel-Export bewusst die ROHEN (numerischen) Daten
            # verwenden statt der Text-formatierten Anzeige-Tabellen, damit
            # TSS/ATL/CTL/TSB/Kg/SYS/DIA/Puls als echte Zahlen exportiert
            # werden (sortier-/rechenbar in Excel) und Trainingszeit als
            # echter Excel-Zeitwert. Das Dezimaltrennzeichen zeigt Excel
            # dann automatisch passend zur eigenen Spracheinstellung an
            # (bei deutschem Excel Komma) - fest einprogrammierte Zeichen
            # sind dafür nicht nötig bzw. nicht zuverlässig.
            rohdaten_export = rohdaten.copy()
            rohdaten_export["Dauer"] = rohdaten_export["Dauer"].apply(dauer_zu_minuten) / (24 * 60)
            rohdaten_export["TSS"] = pd.to_numeric(rohdaten_export["TSS"], errors="coerce")
            for _spalte in ["Kg", "SYS", "DIA", "Puls"]:
                rohdaten_export[_spalte] = pd.to_numeric(rohdaten_export[_spalte], errors="coerce")
            rohdaten_export["Training"] = rohdaten_export["Training"].fillna("")
            rohdaten_export["Datum"] = pd.to_datetime(rohdaten_export["Datum"]).dt.strftime("%d.%m.%Y")
            rohdaten_export = rohdaten_export.rename(columns={"Dauer": "Trainingszeit (hh:mm)"})

            ergebnis_export = ergebnis.copy()
            ergebnis_export["Trainingszeit (hh:mm)"] = ergebnis_export["Dauer_min"] / (24 * 60)
            ergebnis_export = ergebnis_export.rename(columns={"TSB_Bereich": "TSB-Bereich"})
            ergebnis_export.index = ergebnis_export.index.strftime("%d.%m.%Y")
            ergebnis_export.index.name = "Datum"
            ergebnis_export = ergebnis_export[
                ["Training", "Trainingszeit (hh:mm)", "TSS", "ATL", "CTL", "TSB", "TSB-Bereich",
                 "Kg", "SYS", "DIA", "Puls"]
            ]

            excel_puffer = pd.ExcelWriter("ergebnis_export.xlsx", engine="openpyxl")
            rohdaten_export.to_excel(excel_puffer, sheet_name="Trainingseinheiten", index=False)
            ergebnis_export.to_excel(excel_puffer, sheet_name="Tagesauswertung")

            _excel_zeit_zahl_formate = {
                "Trainingszeit (hh:mm)": "[hh]:mm",
                "TSS": "0.0", "ATL": "0.0", "CTL": "0.0", "TSB": "0.0",
                "Kg": "0.0", "SYS": "0.0", "DIA": "0.0", "Puls": "0.0",
            }
            _excel_zahlenformate_setzen(excel_puffer.sheets["Trainingseinheiten"], _excel_zeit_zahl_formate)
            _excel_zahlenformate_setzen(excel_puffer.sheets["Tagesauswertung"], _excel_zeit_zahl_formate)

            excel_puffer.close()
            with open("ergebnis_export.xlsx", "rb") as f:
                st.download_button(
                    "Auswertung als Excel herunterladen",
                    data=f,
                    file_name="tss_atl_ctl_tsb_ergebnis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
