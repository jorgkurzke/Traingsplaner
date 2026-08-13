import streamlit as st
import gpxpy
import gpxpy.gpx
import pandas as pd
from datetime import datetime, timedelta
import math
import io

# ------------------------------------------------------------
# Hilfsfunktionen
# ------------------------------------------------------------

def haversine(lat1, lon1, lat2, lon2):
    R = 6371000
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi = math.radians(lat2 - lat1)
    dlambda = math.radians(lon2 - lon1)

    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlambda/2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
    return R * c


def gpx_to_df(gpx_file):
    gpx = gpxpy.parse(gpx_file)
    points = []

    for track in gpx.tracks:
        for segment in track.segments:
            for p in segment.points:
                points.append({
                    "time": p.time.replace(tzinfo=None),
                    "lat": p.latitude,
                    "lon": p.longitude,
                    "ele": p.elevation
                })

    df = pd.DataFrame(points)
    df = df.sort_values("time").reset_index(drop=True)

    # Distanz berechnen
    dists = [0.0]
    for i in range(1, len(df)):
        d = haversine(df.loc[i-1, "lat"], df.loc[i-1, "lon"],
                      df.loc[i, "lat"], df.loc[i, "lon"])
        dists.append(d)

    df["dist_m"] = dists
    df["dist_km_cum"] = df["dist_m"].cumsum() / 1000.0
    return df


def compute_stand_times(df):
    """Berechnet Standzeiten (Stopps) zwischen Punkten im GPX."""
    stand_times = [0]
    for i in range(1, len(df)):
        dt = (df.loc[i, "time"] - df.loc[i-1, "time"]).total_seconds()
        d = df.loc[i, "dist_m"]
        stand_times.append(dt if d < 1 else 0)
    df["stand_seconds"] = stand_times
    return df


def compute_gradient(df):
    """Berechnet Steigung in % zwischen Punkten."""
    gradients = [0]
    for i in range(1, len(df)):
        ele_diff = df.loc[i, "ele"] - df.loc[i-1, "ele"]
        dist = df.loc[i, "dist_m"]
        gradients.append((ele_diff / dist) * 100 if dist > 0 else 0)
    df["gradient"] = gradients
    return df


def interpolate_time_at_distance(df, target_km):
    if target_km <= df["dist_km_cum"].iloc[0]:
        return df["time"].iloc[0]
    if target_km >= df["dist_km_cum"].iloc[-1]:
        return df["time"].iloc[-1]

    before = df[df["dist_km_cum"] <= target_km].iloc[-1]
    after = df[df["dist_km_cum"] >= target_km].iloc[0]

    if before["dist_km_cum"] == after["dist_km_cum"]:
        return before["time"]

    ratio = ((target_km - before["dist_km_cum"]) /
             (after["dist_km_cum"] - before["dist_km_cum"]))
    dt = after["time"] - before["time"]
    return before["time"] + ratio * dt


def stand_time_between(df, km_start, km_end):
    segment = df[(df["dist_km_cum"] >= km_start) & (df["dist_km_cum"] < km_end)]
    return segment["stand_seconds"].sum()

def fmt_speed(v):
    if v is None:
        return ""
    return f"{v:.1f}".replace(".", ",")

def speed_by_gradient(df, km_start, km_end):
    """Berechnet Geschwindigkeiten nach Steigungskategorien ohne Pausen/Standzeiten."""
    segment = df[(df["dist_km_cum"] >= km_start) & (df["dist_km_cum"] < km_end)]

    def calc_speed(mask):
        seg = segment[mask]
        if len(seg) < 2:
            return None  # keine Daten → Zelle leer

        speeds = []
        weights = []

        for i in range(1, len(seg)):
            dt = (seg.iloc[i]["time"] - seg.iloc[i-1]["time"]).total_seconds()
            dist = seg.iloc[i]["dist_m"]

            # Pausen ignorieren
            if dt <= 0 or dist <= 0:
                continue

            v = (dist / 1000.0) / (dt / 3600.0)  # km/h
            speeds.append(v)
            weights.append(dist)

        if not speeds:
            return None

        # distanzgewichteter Mittelwert
        return sum(s * w for s, w in zip(speeds, weights)) / sum(weights)

    speed_down = calc_speed(segment["gradient"] < -6)
    speed_light_down = calc_speed((segment["gradient"] < 0) & (segment["gradient"] >= -6))
    speed_flat = calc_speed((segment["gradient"] >= 0) & (segment["gradient"] <= 2))
    speed_light_up = calc_speed((segment["gradient"] > 2) & (segment["gradient"] <= 6))
    speed_medium_up = calc_speed((segment["gradient"] > 6) & (segment["gradient"] <= 8))
    speed_steep_up = calc_speed((segment["gradient"] > 8) & (segment["gradient"] <= 10))
    speed_very_steep_up = calc_speed(segment["gradient"] > 10)

    return (
        speed_down,
        speed_light_down,
        speed_flat,
        speed_light_up,
        speed_medium_up,
        speed_steep_up,
        speed_very_steep_up
    )




def format_hhmm(hours_float):
    total_minutes = int(hours_float * 60)
    hh = total_minutes // 60
    mm = total_minutes % 60
    return f"{hh:02d}:{mm:02d}"


# ------------------------------------------------------------
# Streamlit App
# ------------------------------------------------------------

st.title("GPS-Track Analyse")

uploaded_file = st.file_uploader("GPX-Datei hochladen", type=["gpx"], key="gpx_upload")

if uploaded_file is not None:
    df = gpx_to_df(uploaded_file)
    df = compute_stand_times(df)
    df = compute_gradient(df)

    st.success(f"Track geladen: {len(df)} Punkte, {df['dist_km_cum'].iloc[-1]:.1f} km")

    default_start = df["time"].iloc[0]
    start_time = st.time_input("Startzeit", default_start.time())
    start_date = st.date_input("Startdatum", default_start.date())
    start_datetime = datetime.combine(start_date, start_time)

    max_dist = df["dist_km_cum"].iloc[-1]

    st.subheader("Kontrollpunkte aus Excel laden")

    excel_file = st.file_uploader(
        "Excel-Datei mit Kontrollpunkten (Spalten: km, name)",
        type=["xlsx", "xls"],
        key="excel_controls"
    )

    controls = []

    if excel_file is not None:
        try:
            df_controls = pd.read_excel(excel_file, engine="openpyxl")
            if "km" not in df_controls.columns or "name" not in df_controls.columns:
                st.error("Excel muss die Spalten 'km' und 'name' enthalten.")
            else:
                for _, row in df_controls.iterrows():
                    controls.append({"km": float(row["km"]), "name": str(row["name"])})
                st.success(f"{len(controls)} Kontrollpunkte aus Excel geladen.")
        except Exception as e:
            st.error(f"Fehler beim Lesen der Excel-Datei: {e}")

    if len(controls) == 0:
        st.info("Keine Excel-Datei geladen – Kontrollpunkte manuell eingeben.")
        num_points = st.number_input("Anzahl Kontrollpunkte", min_value=1, max_value=30, value=3)

        for i in range(num_points):
            col1, col2 = st.columns(2)
            with col1:
                dist = st.number_input(
                    f"Distanz Punkt {i+1} (km)",
                    min_value=0.0,
                    max_value=float(max_dist),
                    value=min(float(max_dist), (i+1)*50.0),
                    key=f"dist_{i}"
                )
            with col2:
                name = st.text_input(
                    f"Name Punkt {i+1}",
                    value=f"Punkt {i+1}",
                    key=f"name_{i}"
                )
            controls.append({"km": dist, "name": name})

    if st.button("Berechnen", key="calc_button"):
        controls = sorted(controls, key=lambda x: x["km"])

        results = []
        last_km = 0.0
        current_start = start_datetime

        for cp in controls:
            cp_km = cp["km"]

            gps_time_at_cp = interpolate_time_at_distance(df, cp_km)
            gps_time_at_last = interpolate_time_at_distance(df, last_km)

            segment_duration = gps_time_at_cp - gps_time_at_last
            segment_seconds = segment_duration.total_seconds()
            segment_hours = segment_seconds / 3600.0
            segment_dist = cp_km - last_km
            
            pause_seconds = stand_time_between(df, last_km, cp_km)
            pause_td = timedelta(seconds=pause_seconds)
            
            # Netto-Fahrzeit (ohne Pausen/Standzeiten)
            netto_seconds = max(segment_seconds - pause_seconds, 0)
            netto_hours = netto_seconds / 3600.0
            
            # Geschwindigkeiten
            speed_brutto = segment_dist / (segment_hours) if segment_hours > 0 else 0
            speed_netto = segment_dist / (netto_hours) if netto_hours > 0 else 0
            
            # Ankunft = letzte Abfahrt + Netto-Fahrzeit
            arrival_time = current_start + timedelta(seconds=netto_seconds)
            
            # Abfahrt = Ankunft + Pause
            departure_time = arrival_time + pause_td


            (
                speed_down,
                speed_light_down,
                speed_flat,
                speed_light_up,
                speed_medium_up,
                speed_steep_up,
                speed_very_steep_up
            ) = speed_by_gradient(df, last_km, cp_km)

            base_offset = gps_time_at_cp - df["time"].iloc[0]
            #arrival_time = current_start + base_offset
            #departure_time = arrival_time + pause_td

            def fmt_speed(v):
                if v is None:
                    return ""
                return f"{v:.1f}".replace(".", ",")
            
            results.append({
                "Name": cp["name"],
                "km": cp_km,
                "Ankunft": arrival_time,
                "Pause_min": round(pause_seconds / 60.0, 1),
                "Segment_h": format_hhmm(segment_hours),
                "Netto_kmh": f"{speed_netto:.1f}".replace(".", ","),
                "Brutto_kmh": f"{speed_brutto:.1f}".replace(".", ","),
                "Speed_down": fmt_speed(speed_down),
                "Speed_light_down": fmt_speed(speed_light_down),
                "Speed_flat": fmt_speed(speed_flat),
                "Speed_light_up": fmt_speed(speed_light_up),
                "Speed_medium_up": fmt_speed(speed_medium_up),
                "Speed_steep_up": fmt_speed(speed_steep_up),
                "Speed_very_steep_up": fmt_speed(speed_very_steep_up),
                "Abfahrt": departure_time
            })


            last_km = cp_km
            current_start = departure_time

        res_df = pd.DataFrame(results)

        st.subheader("Ergebnisse")
        st.dataframe(res_df)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            res_df.to_excel(writer, index=False)

        st.download_button(
            label="Ergebnisse als Excel herunterladen",
            data=output.getvalue(),
            file_name="brevet_ergebnisse.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

else:
    st.info("Bitte eine GPX-Datei hochladen.")







