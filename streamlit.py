from __future__ import annotations

from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
import altair as alt


st.set_page_config(
    page_title="Orac Energie Analyse",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)


APP_DIR = Path(__file__).resolve().parent
DEFAULT_FILES = [
    APP_DIR / "energie_orac_2024.xlsx",
    APP_DIR / "energie_orac_2025.xlsx",
    APP_DIR / "data" / "energiebeheer_2024.xlsx",
    APP_DIR / "data" / "combined_electricity_consumption_2024.xlsx",
    APP_DIR / "data" / "combined_electricity_consumption_2025.xlsx",
]

EV_CONSUMPTION_FILES = [
    APP_DIR / "data" / "Consumption of Orac 01-05-2025 tot 01-05-2026.xlsx",
    APP_DIR / "data" / "Consumption of Orac.xlsx",
]

EV_SESSION_FILES = [
    APP_DIR / "data" / "Sessions (1).xlsx",
    APP_DIR / "data" / "Sessions.xlsx",
]

COLUMN_ALIASES = {
    "datetime": ["datetime", "date", "timestamp", "tijdstip", "datum", "DateTime"],
    "afname_kwh": ["afname_kwh", "import_kwh", "verbruik_kwh", "consumption_kwh", "afname"],
    "productie_kwh": ["productie_kwh", "production_kwh", "opwek_kwh", "generation_kwh", "productie"],
    "injectie_kwh": ["injectie_kwh", "export_kwh", "teruglevering_kwh", "injection_kwh", "injectie"],
}


def detect_file() -> Path | None:
    for candidate in DEFAULT_FILES:
        if candidate.exists():
            return candidate
    return None


def resolve_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    lower_map = {str(column).strip().lower(): column for column in df.columns}
    for alias in aliases:
        column = lower_map.get(alias.lower())
        if column is not None:
            return column
    return None


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return buffer.getvalue()


def energy_scale(unit: str) -> tuple[float, str]:
    if unit == "MWh":
        return 1000.0, "MWh"
    if unit == "GWh":
        return 1_000_000.0, "GWh"
    return 1.0, "kWh"


def format_energy(value: float, unit: str, decimals: int = 2) -> str:
    factor, label = energy_scale(unit)
    return f"{value / factor:.{decimals}f} {label}"


def auto_energy_unit(value: float) -> str:
    if abs(value) >= 1_000_000:
        return "GWh"
    if abs(value) >= 1_000:
        return "MWh"
    return "kWh"


def format_dutch_datetime(value: pd.Timestamp | None, include_year: bool = False) -> str:
    if value is None or pd.isna(value):
        return "n.v.t."

    month_names = {
        1: "januari",
        2: "februari",
        3: "maart",
        4: "april",
        5: "mei",
        6: "juni",
        7: "juli",
        8: "augustus",
        9: "september",
        10: "oktober",
        11: "november",
        12: "december",
    }
    date_part = f"{value.day} {month_names[value.month]}"
    if include_year:
        date_part = f"{date_part} {value.year}"
    return f"{date_part} om {value:%H:%M}"


def classify_peak_user_type(sessions: pd.DataFrame, peak_time: pd.Timestamp | None) -> str:
    if peak_time is None or sessions.empty:
        return "Onbekend"

    if "start_datetime" not in sessions.columns or "user_type" not in sessions.columns:
        return "Onbekend"

    session_frame = sessions.copy()
    session_frame["session_end"] = session_frame["end_datetime"]
    missing_end_mask = session_frame["session_end"].isna() & session_frame["start_datetime"].notna() & session_frame["duration_hours"].notna()
    if missing_end_mask.any():
        session_frame.loc[missing_end_mask, "session_end"] = (
            session_frame.loc[missing_end_mask, "start_datetime"]
            + pd.to_timedelta(session_frame.loc[missing_end_mask, "duration_hours"], unit="h")
        )

    peak_end = peak_time + pd.Timedelta(hours=1)
    overlaps = session_frame[
        session_frame["start_datetime"].notna()
        & session_frame["session_end"].notna()
        & (session_frame["start_datetime"] < peak_end)
        & (session_frame["session_end"] > peak_time)
    ].copy()

    if overlaps.empty:
        return "Onbekend"

    if "session_kwh" in overlaps.columns and overlaps["session_kwh"].notna().any():
        weighted = overlaps.groupby("user_type", as_index=False)["session_kwh"].sum()
        winner = weighted.sort_values("session_kwh", ascending=False).iloc[0]
    else:
        weighted = overlaps.groupby("user_type", as_index=False).size().rename(columns={"size": "count"})
        winner = weighted.sort_values("count", ascending=False).iloc[0]

    user_type = str(winner["user_type"]).strip().lower()
    if user_type == "werknemer":
        return "Werknemer"
    if user_type == "visitor":
        return "Visitor"
    return str(winner["user_type"]).strip() or "Onbekend"


def convert_energy_frame(frame: pd.DataFrame, unit: str, columns: list[str]) -> pd.DataFrame:
    factor, _ = energy_scale(unit)
    converted = frame.copy()
    for column in columns:
        if column in converted.columns:
            converted[column] = converted[column] / factor
    return converted


def detect_first_existing(paths: list[Path]) -> Path | None:
    for candidate in paths:
        if candidate.exists():
            return candidate
    return None


def find_column_by_keywords(df: pd.DataFrame, keywords: list[str]) -> str | None:
    lower_map = {str(column).strip().lower(): column for column in df.columns}
    for column_lower, original_column in lower_map.items():
        if any(keyword in column_lower for keyword in keywords):
            return original_column
    return None


def parse_duration_hours(value: object) -> float:
    if pd.isna(value):
        return 0.0
    if isinstance(value, pd.Timedelta):
        return value.total_seconds() / 3600.0

    text = str(value).strip()
    if not text:
        return 0.0

    if ":" in text:
        parts = text.split(":")
        try:
            hours = int(parts[0])
            minutes = int(parts[1]) if len(parts) > 1 else 0
            seconds = int(parts[2]) if len(parts) > 2 else 0
            return hours + minutes / 60.0 + seconds / 3600.0
        except ValueError:
            return 0.0

    try:
        return float(text)
    except ValueError:
        return 0.0


def load_ev_consumption_data(source: str | bytes | Path) -> pd.DataFrame:
    df = load_data(source)
    if df.empty:
        return df

    datetime_col = find_column_by_keywords(df, ["datetime", "timestamp", "datum", "date", "time", "tijd"])
    if datetime_col is None:
        datetime_col = df.columns[0]

    df = df.copy()
    df[datetime_col] = pd.to_datetime(df[datetime_col], errors="coerce")

    for column in df.columns:
        if column == datetime_col:
            continue
        cleaned = df[column].astype(str).str.replace(",", ".", regex=False)
        df[column] = pd.to_numeric(cleaned, errors="coerce")

    df = df.dropna(subset=[datetime_col]).sort_values(datetime_col).reset_index(drop=True)
    # Determine sampling interval in hours (median diff) to convert kW -> kWh when needed
    if len(df) >= 2:
        diffs = df[datetime_col].diff().dt.total_seconds().dropna() / 3600.0
        interval_hours = float(diffs.median()) if not diffs.empty else 1.0
        if interval_hours <= 0:
            interval_hours = 1.0
    else:
        interval_hours = 1.0

    numeric_columns = df.select_dtypes(include="number").columns.tolist()
    # Separate columns that look like kWh vs kW
    kwh_cols = [c for c in numeric_columns if "kwh" in c.lower()]
    kw_cols = [c for c in numeric_columns if "kw" in c.lower() and c not in kwh_cols]

    total_series = pd.Series(0.0, index=df.index)
    if kwh_cols:
        total_series = total_series.add(df[kwh_cols].sum(axis=1), fill_value=0)
    if kw_cols:
        # Convert kW to kWh by multiplying with the sampling interval
        total_series = total_series.add(df[kw_cols].sum(axis=1) * interval_hours, fill_value=0)
    # Fallback: if no kw/kwh-like columns found, sum all numeric columns (existing behaviour)
    if not kwh_cols and not kw_cols and numeric_columns:
        total_series = df[numeric_columns].sum(axis=1)

    df["total_kwh"] = total_series
    df = df.rename(columns={datetime_col: "datetime"})
    df["date"] = df["datetime"].dt.date
    df["hour"] = df["datetime"].dt.hour
    df["weekday"] = df["datetime"].dt.day_name()
    df["month"] = df["datetime"].dt.to_period("M").astype(str)
    return df


def load_ev_sessions_data(source: str | bytes | Path) -> pd.DataFrame:
    df = load_data(source)
    if df.empty:
        return df

    df = df.copy()
    start_col = find_column_by_keywords(df, ["from", "start", "begin"])
    end_col = find_column_by_keywords(df, ["to", "end"])
    duration_col = find_column_by_keywords(df, ["duration"])
    kwh_col = find_column_by_keywords(df, ["kwh", "energy"])
    token_col = find_column_by_keywords(df, ["token"])
    station_col = find_column_by_keywords(df, ["charging station", "station"])
    connector_col = find_column_by_keywords(df, ["connector"])

    if start_col is not None:
        df[start_col] = pd.to_datetime(df[start_col], errors="coerce")
        df = df.rename(columns={start_col: "start_datetime"})
    else:
        df["start_datetime"] = pd.NaT

    if end_col is not None:
        df[end_col] = pd.to_datetime(df[end_col], errors="coerce")
        df = df.rename(columns={end_col: "end_datetime"})
    else:
        df["end_datetime"] = pd.NaT

    if duration_col is not None:
        df[duration_col] = df[duration_col].apply(parse_duration_hours)
        df = df.rename(columns={duration_col: "duration_hours"})
    else:
        df["duration_hours"] = 0.0

    if kwh_col is not None:
        cleaned = df[kwh_col].astype(str).str.replace(",", ".", regex=False)
        df[kwh_col] = pd.to_numeric(cleaned, errors="coerce")
        df = df.rename(columns={kwh_col: "session_kwh"})
    else:
        df["session_kwh"] = 0.0

    if token_col is not None:
        df = df.rename(columns={token_col: "token"})
        df["user_type"] = df["token"].apply(lambda value: "werknemer" if pd.notna(value) else "visitor")
    else:
        df["user_type"] = "onbekend"

    if station_col is not None:
        df = df.rename(columns={station_col: "station_name"})
    else:
        df["station_name"] = "Onbekend"

    if connector_col is not None:
        df = df.rename(columns={connector_col: "connector"})
    elif "connector" not in df.columns:
        df["connector"] = df["station_name"].astype(str).str.extract(r"(\d+)$", expand=False)

    if df["duration_hours"].eq(0).any() and df["start_datetime"].notna().any() and df["end_datetime"].notna().any():
        mask = df["duration_hours"].eq(0)
        df.loc[mask, "duration_hours"] = (df.loc[mask, "end_datetime"] - df.loc[mask, "start_datetime"]).dt.total_seconds() / 3600.0

    if df["end_datetime"].isna().all():
        df["end_datetime"] = df["start_datetime"] + pd.to_timedelta(df["duration_hours"], unit="h")

    base_datetime = df["start_datetime"].combine_first(df["end_datetime"])
    df["date"] = base_datetime.dt.date
    df["hour"] = base_datetime.dt.hour
    df["weekday"] = base_datetime.dt.day_name()
    df["month"] = base_datetime.dt.to_period("M").astype(str)
    df = df.sort_values("start_datetime").reset_index(drop=True)
    return df


@st.cache_data(show_spinner=False)
def load_data(source: str | bytes | Path) -> pd.DataFrame:
    if isinstance(source, bytes):
        df = pd.read_excel(BytesIO(source))
    else:
        df = pd.read_excel(source)
    return df.copy()


def standardize_data(raw_df: pd.DataFrame) -> pd.DataFrame:
    df = raw_df.copy()

    datetime_col = resolve_column(df, COLUMN_ALIASES["datetime"])
    import_col = resolve_column(df, COLUMN_ALIASES["afname_kwh"])
    production_col = resolve_column(df, COLUMN_ALIASES["productie_kwh"])
    export_col = resolve_column(df, COLUMN_ALIASES["injectie_kwh"])

    missing = [
        label
        for label, column in [
            ("datetime", datetime_col),
            ("afname_kwh", import_col),
            ("productie_kwh", production_col),
            ("injectie_kwh", export_col),
        ]
        if column is None
    ]
    if missing:
        raise ValueError(
            "Ontbrekende kolommen: " + ", ".join(missing) + ". Controleer de Excel-bron."
        )

    df = df.rename(
        columns={
            datetime_col: "datetime",
            import_col: "afname_kwh",
            production_col: "productie_kwh",
            export_col: "injectie_kwh",
        }
    )

    df["datetime"] = pd.to_datetime(df["datetime"], errors="coerce")
    numeric_columns = ["afname_kwh", "productie_kwh", "injectie_kwh"]
    for column in numeric_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0)

    df = df.dropna(subset=["datetime"]).sort_values("datetime").reset_index(drop=True)
    df["productie_kwh"] = df["productie_kwh"].clip(lower=0)
    df["totaalverbruik"] = df["afname_kwh"] + df["productie_kwh"] - df["injectie_kwh"]
    df["netto_verbruik"] = df["afname_kwh"] - df["productie_kwh"]
    df["dag"] = df["datetime"].dt.normalize()
    df["datum"] = df["datetime"].dt.date
    df["maand"] = df["datetime"].dt.to_period("M").astype(str)
    df["uur"] = df["datetime"].dt.hour
    df["dagtype"] = df["datetime"].dt.dayofweek.map(lambda x: "Werkdag" if x < 5 else "Weekend")
    return df


def metric_card(label: str, value: str, delta: str | None = None) -> None:
    delta_html = f"<div class='metric-delta'>{delta}</div>" if delta else ""
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def build_occupancy_heatmap(
    sessions: pd.DataFrame,
    value_field: str,
) -> pd.DataFrame:
    if sessions.empty or "start_datetime" not in sessions.columns:
        return pd.DataFrame()

    heatmap = sessions.copy()
    iso_calendar = heatmap["start_datetime"].dt.isocalendar()
    heatmap["iso_year"] = iso_calendar["year"].astype(int)
    heatmap["iso_week"] = iso_calendar["week"].astype(int)
    heatmap["week_label"] = (
        heatmap["iso_year"].astype(str) + "-W" + heatmap["iso_week"].astype(str).str.zfill(2)
    )
    heatmap["hour_label"] = heatmap["start_datetime"].dt.hour.map(lambda hour: f"{int(hour):02d}u")
    weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    weekday_labels = {
        "Monday": "Ma",
        "Tuesday": "Di",
        "Wednesday": "Wo",
        "Thursday": "Do",
        "Friday": "Vr",
        "Saturday": "Za",
        "Sunday": "Zo",
    }
    heatmap["weekday_label"] = pd.Categorical(
        heatmap["start_datetime"].dt.day_name().map(weekday_labels),
        categories=[weekday_labels[day] for day in weekday_order],
        ordered=True,
    )

    if value_field == "hour":
        grouped = heatmap.groupby(["week_label", "hour_label"], as_index=False).size()
        grouped = grouped.rename(columns={"size": "sessies"})
        grouped["week_sort"] = grouped["week_label"]
        grouped["heatmap_x"] = grouped["hour_label"]
    else:
        grouped = heatmap.groupby(["week_label", "weekday_label"], as_index=False).size()
        grouped = grouped.rename(columns={"size": "sessies"})
        grouped["week_sort"] = grouped["week_label"]
        grouped["heatmap_x"] = grouped["weekday_label"]

    return grouped


def render_occupancy_heatmap(data: pd.DataFrame, x_title: str) -> None:
    if data.empty:
        st.info("Geen data beschikbaar voor de heatmap.")
        return

    if x_title == "Dag van de week":
        x_order = ["Ma", "Di", "Wo", "Do", "Vr", "Za", "Zo"]
    else:
        x_order = [f"{hour:02d}u" for hour in range(24)]

    chart = {
        "mark": {"type": "rect"},
        "encoding": {
            "x": {
                "field": "heatmap_x",
                "type": "ordinal",
                "sort": x_order,
                "title": x_title,
            },
            "y": {
                "field": "week_sort",
                "type": "ordinal",
                "sort": sorted(data["week_sort"].dropna().astype(str).unique().tolist()),
                "title": "Week van het jaar",
            },
            "color": {
                "field": "sessies",
                "type": "quantitative",
                "title": "Sessies",
                "scale": {"scheme": "blues"},
            },
            "tooltip": [
                {"field": "week_sort", "type": "ordinal", "title": "Week"},
                {"field": "heatmap_x", "type": "ordinal", "title": x_title},
                {"field": "sessies", "type": "quantitative", "title": "Sessies"},
            ],
        },
        "height": 420,
    }
    st.vega_lite_chart(data, chart, use_container_width=True)


st.markdown(
    """
    <style>
        .metric-card {
            padding: 1rem 1.1rem;
            border-radius: 18px;
            border: 1px solid rgba(255,255,255,0.08);
            background: linear-gradient(180deg, rgba(20, 24, 31, 0.88), rgba(12, 15, 20, 0.96));
            box-shadow: 0 12px 30px rgba(0,0,0,0.14);
            min-height: 110px;
        }
        .metric-label {
            font-size: 0.82rem;
            text-transform: uppercase;
            letter-spacing: 0.08em;
            color: rgba(255,255,255,0.68);
            margin-bottom: 0.35rem;
        }
        .metric-value {
            font-size: 1.7rem;
            font-weight: 700;
            color: #ffffff;
            line-height: 1.1;
        }
        .metric-delta {
            margin-top: 0.4rem;
            color: #9bb7ff;
            font-size: 0.88rem;
        }
        .block-container {
            padding-top: 1.2rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)

def render_ev_tab() -> None:
    st.title("EV analyse")
    st.caption("Interactieve analyse van EV-verbruik en laadsessies met datumselectie, filters en downloads.")

    default_consumption = detect_first_existing(EV_CONSUMPTION_FILES)
    default_sessions = detect_first_existing(EV_SESSION_FILES)

    source_col1, source_col2 = st.columns(2)
    with source_col1:
        uploaded_consumption = st.file_uploader(
            "Upload EV verbruiksbestand",
            type=["xlsx", "xls"],
            key="ev_consumption_upload",
        )
    with source_col2:
        uploaded_sessions = st.file_uploader(
            "Upload EV sessiebestand",
            type=["xlsx", "xls"],
            key="ev_sessions_upload",
        )

    consumption_source = uploaded_consumption.getvalue() if uploaded_consumption is not None else default_consumption
    sessions_source = uploaded_sessions.getvalue() if uploaded_sessions is not None else default_sessions

    if consumption_source is None:
        st.warning("Geen EV-verbruiksbestand gevonden. Upload een Excel-bestand om verder te gaan.")
        return

    ev_consumption = load_ev_consumption_data(consumption_source)
    ev_sessions = load_ev_sessions_data(sessions_source) if sessions_source is not None else pd.DataFrame()

    if ev_consumption.empty:
        st.warning("Het EV-verbruiksbestand bevat geen bruikbare data.")
        return

    min_date = ev_consumption["date"].min()
    max_date = ev_consumption["date"].max()

    with st.expander("EV filters", expanded=True):
        filter_col1, filter_col2, filter_col3 = st.columns([1.2, 1.2, 1])
        with filter_col1:
            selected_range = st.date_input(
                "Periode",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date,
                key="ev_date_range",
            )
            if isinstance(selected_range, tuple) and len(selected_range) == 2:
                start_date, end_date = selected_range
            else:
                start_date = end_date = selected_range

        user_types = ["visitor", "werknemer"] if not ev_sessions.empty and "user_type" in ev_sessions.columns else []
        with filter_col2:
            selected_user_types = st.multiselect(
                "User type",
                options=user_types,
                default=user_types,
                key="ev_user_types",
            ) if user_types else []

        station_options = []
        if not ev_sessions.empty and "station_name" in ev_sessions.columns:
            station_options = sorted(ev_sessions["station_name"].dropna().astype(str).unique().tolist())
        selected_stations = st.multiselect(
            "Laadpaal",
            options=station_options,
            default=station_options[: min(5, len(station_options))],
            key="ev_station_filter",
        ) if station_options else []

    selected_days = (end_date - start_date).days + 1
    if selected_days <= 31:
        resolution = "uurdata"
    else:
        resolution = "dagdata"

    filtered_consumption = ev_consumption[(ev_consumption["date"] >= start_date) & (ev_consumption["date"] <= end_date)].copy()
    filtered_sessions = ev_sessions.copy()
    if not filtered_sessions.empty and "date" in filtered_sessions.columns:
        filtered_sessions = filtered_sessions[(filtered_sessions["date"] >= start_date) & (filtered_sessions["date"] <= end_date)].copy()
    if selected_user_types and not filtered_sessions.empty and "user_type" in filtered_sessions.columns:
        filtered_sessions = filtered_sessions[filtered_sessions["user_type"].isin(selected_user_types)].copy()
    if selected_stations and not filtered_sessions.empty and "station_name" in filtered_sessions.columns:
        filtered_sessions = filtered_sessions[filtered_sessions["station_name"].astype(str).isin(selected_stations)].copy()

    if filtered_consumption.empty:
        st.warning("Geen EV-data binnen deze selectie.")
        return

    total_kwh = float(filtered_consumption["total_kwh"].sum())
    avg_daily_kwh = float(filtered_consumption.groupby("date")["total_kwh"].sum().mean()) if not filtered_consumption.empty else 0.0
    peak_idx = filtered_consumption["total_kwh"].idxmax()
    peak_value = float(filtered_consumption.loc[peak_idx, "total_kwh"]) if pd.notna(peak_idx) else 0.0
    peak_time = filtered_consumption.loc[peak_idx, "datetime"] if pd.notna(peak_idx) else None
    peak_user_type = classify_peak_user_type(filtered_sessions, peak_time)

    session_count = len(filtered_sessions) if not filtered_sessions.empty else 0
    session_kwh = float(filtered_sessions["session_kwh"].sum()) if not filtered_sessions.empty and "session_kwh" in filtered_sessions.columns else 0.0
    avg_session_kwh = float(filtered_sessions["session_kwh"].mean()) if session_count and "session_kwh" in filtered_sessions.columns else 0.0
    avg_duration = float(filtered_sessions["duration_hours"].mean()) if session_count and "duration_hours" in filtered_sessions.columns else 0.0

    st.subheader("Samenvatting")
    ev_metric_cols = st.columns(4)
    with ev_metric_cols[0]:
        metric_card("EV verbruik", f"{total_kwh:,.0f} kWh", f"Resolutie: {resolution}")
    with ev_metric_cols[1]:
        metric_card("Gem. per dag", f"{avg_daily_kwh:,.1f} kWh")
    with ev_metric_cols[2]:
        peak_label = format_dutch_datetime(peak_time, include_year=True)
        if peak_user_type != "Onbekend":
            peak_label = f"{peak_label} · {peak_user_type}"
        metric_card("Piekmoment", f"{peak_value:,.1f} kWh", peak_label)
    with ev_metric_cols[3]:
        metric_card("Sessies", f"{session_count:,}", f"{session_kwh:,.0f} kWh totaal")

    ev_overview_tab, ev_sessions_tab, ev_occupancy_tab, ev_downloads_tab = st.tabs(
        ["Overzicht", "Sessies", "Bezetting", "Downloads"]
    )

    with ev_overview_tab:
        left, right = st.columns([1.35, 1])
        with left:
            st.markdown("### EV verbruik over de geselecteerde periode")
            if selected_days <= 31:
                overview_series = filtered_consumption[["datetime", "total_kwh"]].copy()
            else:
                overview_series = filtered_consumption.set_index("datetime")[["total_kwh"]].resample("D").sum().reset_index()
            st.caption(f"{resolution} getoond voor {selected_days} dag(en). Tot 1 maand = uurdata, daarboven = dagdata.")
            st.line_chart(overview_series, x="datetime", y="total_kwh", height=360)

        with right:
            st.markdown("### Verbruik per dag")
            daily = filtered_consumption.groupby("date", as_index=False)["total_kwh"].sum()
            st.bar_chart(daily, x="date", y="total_kwh", height=360)

        st.markdown("### Gemiddeld uurprofiel")
        hourly_profile = filtered_consumption.groupby("hour", as_index=False)["total_kwh"].mean()
        st.line_chart(hourly_profile, x="hour", y="total_kwh", height=300)

    with ev_sessions_tab:
        st.markdown("### Sessies")
        if filtered_sessions.empty:
            st.info("Geen sessiedata beschikbaar voor de huidige filters.")
        else:
            session_cols = st.columns(4)
            with session_cols[0]:
                metric_card("Sessies", f"{len(filtered_sessions):,}")
            with session_cols[1]:
                metric_card("Totaal kWh", f"{filtered_sessions['session_kwh'].sum():,.0f} kWh")
            with session_cols[2]:
                metric_card("Gem. kWh/sessie", f"{filtered_sessions['session_kwh'].mean():.2f} kWh")
            with session_cols[3]:
                metric_card("Gem. duur", f"{filtered_sessions['duration_hours'].mean():.2f} uur")

            st.markdown("#### Filterde sessietabel")
            display_cols = [column for column in ["start_datetime", "end_datetime", "station_name", "connector", "user_type", "session_kwh", "duration_hours"] if column in filtered_sessions.columns]
            st.dataframe(filtered_sessions[display_cols], use_container_width=True, hide_index=True)

    with ev_occupancy_tab:
        st.markdown("### Bezetting en pieken")
        if filtered_sessions.empty:
            st.info("Geen sessiedata beschikbaar voor bezettingsanalyse.")
        else:
            by_hour = filtered_sessions.groupby(filtered_sessions["start_datetime"].dt.hour).size().reindex(range(24), fill_value=0)
            by_day = filtered_sessions.groupby(filtered_sessions["start_datetime"].dt.day_name()).size()
            weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            by_day = by_day.reindex(weekday_order, fill_value=0)

            occ_col1, occ_col2 = st.columns(2)
            with occ_col1:
                by_hour_df = by_hour.reset_index()
                by_hour_df.columns = ["hour", "sessies"]
                st.bar_chart(by_hour_df, x="hour", y="sessies", height=320)
            with occ_col2:
                by_day_df = by_day.reset_index()
                by_day_df.columns = ["weekday", "sessies"]
                st.bar_chart(by_day_df, x="weekday", y="sessies", height=320)

            st.markdown("#### Interactieve heatmap")
            heatmap_col1, heatmap_col2 = st.tabs(["Week x uur", "Week x dag van de week"])

            with heatmap_col1:
                st.caption("Links de ISO-week, onderaan het uur van de dag. Hover voor details per cel.")
                heatmap_hour = build_occupancy_heatmap(filtered_sessions, "hour")
                render_occupancy_heatmap(heatmap_hour, "Uur van de dag")

            with heatmap_col2:
                st.caption("Links de ISO-week, onderaan de dag van de week. Hover voor details per cel.")
                heatmap_weekday = build_occupancy_heatmap(filtered_sessions, "weekday")
                render_occupancy_heatmap(heatmap_weekday, "Dag van de week")

    with ev_downloads_tab:
        st.markdown("### Downloads")
        download_cols = st.columns(2)
        with download_cols[0]:
            st.download_button(
                "Download EV verbruik als Excel",
                data=to_excel_bytes(filtered_consumption),
                file_name="ev_verbruik_selectie.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with download_cols[1]:
            if not filtered_sessions.empty:
                st.download_button(
                    "Download sessies als Excel",
                    data=to_excel_bytes(filtered_sessions),
                    file_name="ev_sessies_selectie.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.info("Geen sessiedata om te downloaden.")


def render_energy_tab() -> None:
    st.title("Orac Energie Analyse")
    st.caption("Interactieve analyse van afname, productie, injectie en totaalverbruik.")

    with st.sidebar:
        st.header("Bron & filters")

        unit = st.selectbox("Weergave-eenheid", ["kWh", "MWh", "GWh"], index=0, key="energy_unit")

        default_file = detect_file()
        uploaded = st.file_uploader("Upload een Excel-bestand", type=["xlsx", "xls"], key="energy_upload")

        source_label = "Upload"
        if uploaded is not None:
            raw_df = load_data(uploaded.getvalue())
        elif default_file is not None:
            source_label = default_file.name
            raw_df = load_data(default_file)
        else:
            raw_df = None
            st.info("Geen standaard Excel-bestand gevonden. Upload een bestand om te starten.")

        st.caption(f"Geselecteerde bron: {source_label}")

        st.divider()
        ev_default_file = detect_first_existing(EV_CONSUMPTION_FILES)
        ev_uploaded = st.file_uploader("EV verbruiksbestand (optioneel)", type=["xlsx", "xls"], key="energy_ev_upload")
        if ev_uploaded is not None:
            ev_consumption_source = ev_uploaded.getvalue()
        else:
            ev_consumption_source = ev_default_file

    if raw_df is None:
        st.stop()

    try:
        df = standardize_data(raw_df)
    except ValueError as exc:
        st.error(str(exc))
        st.stop()

    ev_consumption_df = load_ev_consumption_data(ev_consumption_source) if ev_consumption_source is not None else pd.DataFrame()

    min_date = df["datum"].min()
    max_date = df["datum"].max()
    available_dates = sorted(pd.Series(df["datum"].unique()).dropna().tolist())

    with st.sidebar:
        st.subheader("Datumselectie")
        selected_range = st.date_input(
            "Periode",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date,
            key="energy_date_range",
        )

        if isinstance(selected_range, tuple) and len(selected_range) == 2:
            start_date, end_date = selected_range
        else:
            start_date = end_date = selected_range

        exact_date = st.selectbox(
            "Exacte datum voor detailweergave",
            options=available_dates,
            index=min(len(available_dates) - 1, max(0, len(available_dates) - 1)),
            format_func=lambda value: value.strftime("%d-%m-%Y"),
            key="energy_exact_date",
        )

        dagtype_filter = st.multiselect(
            "Dagtype",
            ["Werkdag", "Weekend"],
            default=["Werkdag", "Weekend"],
            key="energy_daytype_filter",
        )

    filtered = df[
        (df["datum"] >= start_date)
        & (df["datum"] <= end_date)
        & (df["dagtype"].isin(dagtype_filter))
    ].copy()

    day_df = df[df["datum"] == exact_date].copy()

    if filtered.empty:
        st.warning("Geen data binnen deze selectie.")
        st.stop()

    totals = filtered[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].sum()
    self_consumption = (filtered["productie_kwh"] - filtered["injectie_kwh"]).clip(lower=0).sum()
    self_sufficiency = (self_consumption / totals["afname_kwh"] * 100) if totals["afname_kwh"] else 0

    summary_unit = auto_energy_unit(float(totals.max()))
    summary_factor, summary_label = energy_scale(summary_unit)
    display_totals = totals / summary_factor

    st.subheader("Samenvatting")
    metric_cols = st.columns(4)
    with metric_cols[0]:
        metric_card("Import", f"{display_totals['afname_kwh']:.2f} {summary_label}", f"{len(filtered):,} rijen")
    with metric_cols[1]:
        metric_card("Export", f"{display_totals['injectie_kwh']:.2f} {summary_label}")
    with metric_cols[2]:
        metric_card("Productie", f"{display_totals['productie_kwh']:.2f} {summary_label}")
    with metric_cols[3]:
        metric_card("Totaalverbruik", f"{display_totals['totaalverbruik']:.2f} {summary_label}", f"Zelfvoorziening: {self_sufficiency:.1f}%")

    tab_overview, tab_ev_comparison, tab_day, tab_tables, tab_downloads = st.tabs(
        ["Overzicht", "Laadpark & Verbruik", "Dagdetail", "Tabellen", "Downloads"]
    )

    with tab_overview:
        left, right = st.columns([1.35, 1])

        with left:
            st.markdown("### Verloop over de geselecteerde periode")
            selected_days = (end_date - start_date).days + 1
            if selected_days <= 7:
                time_series = filtered[["datetime", "afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].copy()
                resolution_label = "Kwartierdata"
            elif selected_days <= 31:
                time_series = (
                    filtered.set_index("datetime")[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]]
                    .resample("H")
                    .sum()
                    .reset_index()
                )
                resolution_label = "Uurdata"
            else:
                time_series = (
                    filtered.set_index("datetime")[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]]
                    .resample("D")
                    .sum()
                    .reset_index()
                )
                resolution_label = "Dagdata"
            time_series = convert_energy_frame(time_series, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
            st.caption(f"{resolution_label} getoond voor {selected_days} dag(en). Tot 7 dagen = kwartierdata, tot 1 maand = uurdata, daarboven = dagdata.")
            # Maak een long-form dataframe voor interactieve weergave
            series_cols = ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]
            try:
                long = time_series.melt(id_vars="datetime", value_vars=series_cols, var_name="series", value_name="value")
                label_map = {
                    "afname_kwh": "Afname",
                    "productie_kwh": "Productie",
                    "injectie_kwh": "Injectie",
                    "totaalverbruik": "Totaalverbruik",
                }
                long["series_label"] = long["series"].map(label_map)

                selector = alt.selection_multi(fields=["series_label"], bind="legend")

                chart = (
                    alt.Chart(long)
                    .mark_line()
                    .encode(
                        x=alt.X("datetime:T", title="Datum"),
                        y=alt.Y("value:Q", title=f"Energiemeting"),
                        color=alt.Color("series_label:N", title="Serie"),
                        opacity=alt.condition(selector, alt.value(1), alt.value(0.15)),
                        tooltip=[
                            alt.Tooltip("datetime:T", title="Datum"),
                            alt.Tooltip("series_label:N", title="Serie"),
                            alt.Tooltip("value:Q", title=summary_label),
                        ],
                    )
                    .add_selection(selector)
                    .properties(height=380)
                    .interactive()
                )
                st.altair_chart(chart, use_container_width=True)
            except Exception:
                st.warning("Interactiviteit niet beschikbaar (Altair/Fallback). Standaardgrafiek wordt gebruikt.")
                st.line_chart(time_series, x="datetime", y=series_cols, height=380)

        with right:
            st.markdown("### Maandelijkse totalen")
            monthly = filtered.groupby("maand", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].sum()
            monthly = convert_energy_frame(monthly, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
            st.bar_chart(monthly, x="maand", y=["afname_kwh", "productie_kwh", "injectie_kwh"], height=380)

        st.markdown("### Gemiddeld profiel per uur")
        hourly_profile = filtered.groupby("uur", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].mean()
        hourly_profile = convert_energy_frame(hourly_profile, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
        st.line_chart(hourly_profile, x="uur", y=["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"], height=320)
    with tab_ev_comparison:
        st.markdown("### Laadpark & Verbruik")
        if ev_consumption_df.empty:
            st.info("Geen EV-verbruiksbestand gevonden. Upload een bestand om de vergelijking te zien.")
        else:
            ev_filtered = ev_consumption_df[(ev_consumption_df["date"] >= start_date) & (ev_consumption_df["date"] <= end_date)].copy()
            if ev_filtered.empty:
                st.info("Geen EV-data beschikbaar voor de geselecteerde periode.")
            else:
                # Bereken hourly gemiddelden voor beide (normalize column names)
                building_hourly = filtered.groupby("uur", as_index=False)[["totaalverbruik"]].mean()
                building_hourly = building_hourly.rename(columns={"uur": "hour", "totaalverbruik": "building_kwh"})
                
                ev_hourly = ev_filtered.groupby("hour", as_index=False)[["total_kwh"]].mean()
                ev_hourly = ev_hourly.rename(columns={"total_kwh": "ev_kwh"})
                
                # Merge de data
                combined_hourly = pd.merge(building_hourly, ev_hourly, on="hour", how="outer").fillna(0)
                combined_hourly = convert_energy_frame(combined_hourly, unit, ["building_kwh", "ev_kwh"])
                
                st.line_chart(combined_hourly, x="hour", y=["building_kwh", "ev_kwh"], height=400)
                
                st.markdown("#### Gemiddelden per uur")
                st.dataframe(combined_hourly, use_container_width=True, hide_index=True)
    with tab_day:
        st.markdown(f"### Detail voor {exact_date.strftime('%d-%m-%Y')}")
        if day_df.empty:
            st.info("Geen data beschikbaar voor deze datum.")
        else:
            day_daily = day_df.set_index("datetime")[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].resample("D").sum().reset_index()
            day_daily = convert_energy_frame(day_daily, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
            st.bar_chart(day_daily, x="datetime", y=["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"], height=320)

            st.markdown("#### Uurverbruik op deze datum")
            day_hourly = day_df.groupby("uur", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].sum()
            day_hourly = convert_energy_frame(day_hourly, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
            st.dataframe(day_hourly, use_container_width=True, hide_index=True)

            st.markdown("#### Ruwe rijen voor deze datum")
            day_raw = convert_energy_frame(day_df[["datetime", "afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]], unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
            st.dataframe(day_raw, use_container_width=True, hide_index=True)

    with tab_tables:
        st.markdown("### Interactieve tabellen")
        table_mode = st.radio(
            "Tabeltype",
            ["Ruwe data", "Dagtotalen", "Maandtotalen"],
            horizontal=True,
        )

        if table_mode == "Ruwe data":
            table_df = convert_energy_frame(filtered[["datetime", "afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik", "dagtype"]].copy(), unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
        elif table_mode == "Dagtotalen":
            table_df = (
                filtered.groupby("datum", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]]
                .sum()
                .sort_values("datum")
            )
            table_df = convert_energy_frame(table_df, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
        else:
            table_df = (
                filtered.groupby("maand", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]]
                .sum()
                .sort_values("maand")
            )
            table_df = convert_energy_frame(table_df, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])

        st.dataframe(table_df, use_container_width=True, hide_index=True)

        st.markdown("### Samenvatting per dagtype")
        daytype_summary = filtered.groupby("dagtype", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]].sum()
        daytype_summary = convert_energy_frame(daytype_summary, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
        st.dataframe(daytype_summary, use_container_width=True, hide_index=True)

    with tab_downloads:
        st.markdown("### Exporteer geselecteerde data")
        export_cols = ["datetime", "afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik", "netto_verbruik", "dagtype", "maand"]
        export_df = filtered[export_cols].copy()
        export_df = convert_energy_frame(export_df, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik", "netto_verbruik"])
        daily_export = (
            filtered.groupby("datum", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]]
            .sum()
            .sort_values("datum")
        )
        daily_export = convert_energy_frame(daily_export, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])
        monthly_export = (
            filtered.groupby("maand", as_index=False)[["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"]]
            .sum()
            .sort_values("maand")
        )
        monthly_export = convert_energy_frame(monthly_export, unit, ["afname_kwh", "productie_kwh", "injectie_kwh", "totaalverbruik"])

        download_cols = st.columns(3)
        with download_cols[0]:
            st.download_button(
                "Download ruwe selectie als Excel",
                data=to_excel_bytes(export_df),
                file_name="orac_selectie.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with download_cols[1]:
            st.download_button(
                "Download dagtotalen als Excel",
                data=to_excel_bytes(daily_export),
                file_name="orac_dagtotalen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        with download_cols[2]:
            st.download_button(
                "Download maandtotalen als Excel",
                data=to_excel_bytes(monthly_export),
                file_name="orac_maandtotalen.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        st.markdown("### Snelle export als CSV")
        st.download_button(
            "Download ruwe selectie als CSV",
            data=export_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="orac_selectie.csv",
            mime="text/csv",
        )
with st.sidebar:
    app_section = st.radio(
        "Menu",
        ["EV", "Energie analyse"],
        index=0,
    )

if app_section == "EV":
    render_ev_tab()
else:
    render_energy_tab()
