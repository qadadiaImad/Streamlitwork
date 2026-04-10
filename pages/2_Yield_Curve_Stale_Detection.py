"""
Page 2 — Yield Curve Stale Detection
======================================
Load daily yield curve CSVs, build y(date, horizon) surface,
detect stale data points, and visualize on 3D surface + 2D slices.

Expected CSV format per file (named dataJJMMYYYY.csv or any name):
    horizon, rate
    1M,      0.045
    3M,      0.047
    6M,      0.051
    1Y,      0.053
    2Y,      0.055
    ...

Or a single consolidated CSV/Excel:
    date,    1M,    3M,    6M,    1Y,    2Y, ...
    01012024, 0.045, 0.047, 0.051, 0.053, 0.055
    02012024, 0.045, 0.047, 0.051, 0.054, 0.056
    ...
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import re
import os
from pathlib import Path
from datetime import datetime, timedelta
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))
from modules.shared import inject_css, render_header, render_stat_card
from modules.formula_engine import FormulaEngine, apply_formula_to_surface

inject_css()
render_header("📈", "Yield Curve — Stale Detection", "Load curves · Build surface y(date, horizon) · Detect & highlight stale zones")


# ============================================================
# HORIZON PARSING
# ============================================================
HORIZON_ORDER = {
    'ON': 1/365, '1W': 7/365, '2W': 14/365, '1M': 1/12, '2M': 2/12, '3M': 3/12,
    '4M': 4/12, '5M': 5/12, '6M': 6/12, '9M': 9/12,
    '1Y': 1, '2Y': 2, '3Y': 3, '4Y': 4, '5Y': 5,
    '6Y': 6, '7Y': 7, '8Y': 8, '9Y': 9, '10Y': 10,
    '12Y': 12, '15Y': 15, '20Y': 20, '25Y': 25, '30Y': 30,
    '40Y': 40, '50Y': 50,
}

def parse_horizon(h: str) -> float:
    """Convert horizon string like '3M', '10Y' to numeric year fraction."""
    h = str(h).strip().upper()
    if h in HORIZON_ORDER:
        return HORIZON_ORDER[h]
    m = re.match(r'^(\d+(?:\.\d+)?)\s*([DWMY])$', h)
    if m:
        val, unit = float(m.group(1)), m.group(2)
        if unit == 'D': return val / 365
        if unit == 'W': return val * 7 / 365
        if unit == 'M': return val / 12
        if unit == 'Y': return val
    # Try pure numeric (already in years)
    try:
        return float(h)
    except ValueError:
        return None


def parse_date_from_filename(filename: str):
    """Extract date from filename like dataJJMMYYYY.csv or data_JJMMYYYY.csv."""
    patterns = [
        (r'(\d{8})', None),  # raw 8 digits
    ]
    base = Path(filename).stem
    m = re.search(r'(\d{8})', base)
    if m:
        digits = m.group(1)
        # Try JJMMYYYY
        try:
            return datetime.strptime(digits, '%d%m%Y')
        except ValueError:
            pass
        # Try YYYYMMDD
        try:
            return datetime.strptime(digits, '%Y%m%d')
        except ValueError:
            pass
        # Try DDMMYYYY
        try:
            return datetime.strptime(digits, '%d%m%Y')
        except ValueError:
            pass
    return None


def parse_date_flexible(val):
    """Parse a date from various formats."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return None
    if isinstance(val, (datetime, pd.Timestamp)):
        return pd.Timestamp(val)
    # Excel stores dates as serial numbers (float/int); handle before string conversion
    if isinstance(val, (int, float, np.integer, np.floating)):
        try:
            return pd.Timestamp('1899-12-30') + pd.Timedelta(days=int(val))
        except Exception:
            return None
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'nat', 'none', 'n/a', ''):
        return None
    for fmt in ('%d%m%Y', '%d/%m/%Y', '%Y-%m-%d', '%Y%m%d', '%m/%d/%Y',
                '%d-%m-%Y', '%d.%m.%Y', '%Y/%m/%d', '%d %b %Y', '%d %B %Y'):
        try:
            return pd.Timestamp(datetime.strptime(s, fmt))
        except ValueError:
            continue
    try:
        return pd.Timestamp(pd.to_datetime(s, dayfirst=True))
    except Exception:
        return None


# ============================================================
# STALE DETECTION ENGINE
# ============================================================
def detect_stale(surface_df: pd.DataFrame, tolerance: float = 1e-8,
                 lookback: int = 1) -> pd.DataFrame:
    """
    Detect stale data points on the yield surface.

    A point y(date_i, horizon_j) is STALE if:
        |y(date_i, h_j) - y(date_{i-lookback}, h_j)| <= tolerance

    Returns a boolean DataFrame of same shape: True = stale.
    """
    stale = pd.DataFrame(False, index=surface_df.index, columns=surface_df.columns)

    for shift in range(1, lookback + 1):
        shifted = surface_df.shift(shift)
        diff = (surface_df - shifted).abs()
        stale = stale | (diff <= tolerance)

    # First `lookback` rows can't be compared — mark as not stale
    stale.iloc[:lookback] = False
    return stale


def detect_stale_horizon(surface_df: pd.DataFrame, tolerance: float = 1e-8) -> pd.DataFrame:
    """
    Detect stale along the horizon axis:
    y(date_i, h_j) == y(date_i, h_{j-1})  (flat across tenors on same date).
    """
    stale_h = pd.DataFrame(False, index=surface_df.index, columns=surface_df.columns)
    for j in range(1, len(surface_df.columns)):
        diff = (surface_df.iloc[:, j] - surface_df.iloc[:, j-1]).abs()
        stale_h.iloc[:, j] = diff <= tolerance
    return stale_h


def compute_stale_streaks(surface_df: pd.DataFrame, tolerance: float = 1e-8) -> pd.DataFrame:
    """
    For each (date, horizon), compute consecutive days the value has been unchanged.
    0 = fresh, 1 = same as yesterday, 5 = unchanged for 5 days, etc.
    """
    streaks = pd.DataFrame(0, index=surface_df.index, columns=surface_df.columns)
    for i in range(1, len(surface_df)):
        for col in surface_df.columns:
            if abs(surface_df.iloc[i][col] - surface_df.iloc[i-1][col]) <= tolerance:
                streaks.iloc[i][col] = streaks.iloc[i-1][col] + 1
            else:
                streaks.iloc[i][col] = 0
    return streaks


# ============================================================
# SIDEBAR — DATA INPUT
# ============================================================
with st.sidebar:
    st.markdown("### 📁 Load Yield Curves")

    input_mode = st.radio("Input mode", [
        "📄 Single consolidated file",
        "📂 Multiple daily files"
    ], label_visibility="collapsed")

    if "Single" in input_mode:
        uploaded = st.file_uploader(
            "Consolidated CSV/Excel (dates as rows, horizons as columns)",
            type=['csv', 'txt', 'xlsx', 'xlsb', 'xlsm'],
            key="yc_single"
        )
    else:
        uploaded = st.file_uploader(
            "Daily curve files (dataJJMMYYYY.csv)",
            type=['csv', 'txt'],
            accept_multiple_files=True,
            key="yc_multi"
        )

    st.markdown("---")
    st.markdown("### ⚙️ Stale Detection")
    tolerance = st.number_input("Tolerance (abs diff ≤ tol → stale)",
                                value=0.0, min_value=0.0, step=0.0001, format="%.6f")
    lookback = st.slider("Lookback days", 1, 10, 1,
                         help="Number of prior days to compare against")
    streak_threshold = st.slider("Streak alert threshold (days)", 1, 30, 3,
                                 help="Highlight if unchanged for ≥ N consecutive days")

    st.markdown("---")
    st.markdown("### 📅 Date Filter")
    filter_dates = st.checkbox("Filter date range", value=False)


# ============================================================
# DATA LOADING
# ============================================================
surface_df = None

if "Single" in input_mode and uploaded is not None:
    filename = getattr(uploaded, 'name', 'file.csv')
    ext = Path(filename).suffix.lower()
    try:
        if ext in ('.csv', '.txt'):
            raw = pd.read_csv(uploaded)
        elif ext == '.xlsb':
            raw = pd.read_excel(uploaded, engine='pyxlsb')
        else:
            raw = pd.read_excel(uploaded, engine='openpyxl')

        # Expect first column = date, rest = horizons
        date_col = raw.columns[0]
        raw[date_col] = raw[date_col].apply(parse_date_flexible)
        raw = raw.dropna(subset=[date_col])
        raw = raw.set_index(date_col).sort_index()

        # Parse horizon columns to numeric and sort
        horizon_map = {}
        for col in raw.columns:
            hval = parse_horizon(str(col))
            if hval is not None:
                horizon_map[col] = hval
        raw = raw[[c for c in raw.columns if c in horizon_map]]
        raw = raw.rename(columns=horizon_map)
        raw = raw.sort_index(axis=1)
        raw = raw.apply(pd.to_numeric, errors='coerce')
        surface_df = raw

    except Exception as e:
        st.error(f"Failed to load: {e}")

elif "Multiple" in input_mode and uploaded:
    curves = {}
    for f in uploaded:
        fname = getattr(f, 'name', 'unknown.csv')
        dt = parse_date_from_filename(fname)
        if dt is None:
            st.warning(f"Cannot parse date from '{fname}', skipping.")
            continue
        try:
            df = pd.read_csv(f)
            # Expect columns: horizon, rate (or similar 2-col format)
            if len(df.columns) >= 2:
                df.columns = ['horizon', 'rate'] + list(df.columns[2:])
                row = {}
                for _, r in df.iterrows():
                    hval = parse_horizon(str(r['horizon']))
                    if hval is not None:
                        row[hval] = float(r['rate'])
                curves[pd.Timestamp(dt)] = row
        except Exception as e:
            st.warning(f"Error reading {fname}: {e}")

    if curves:
        surface_df = pd.DataFrame(curves).T.sort_index()
        surface_df = surface_df.sort_index(axis=1)


# ============================================================
# MAIN DISPLAY
# ============================================================
if surface_df is not None and not surface_df.empty:
    surface_df = surface_df.dropna(how='all')

    # Date filter
    if filter_dates:
        min_d, max_d = surface_df.index.min().date(), surface_df.index.max().date()
        d1, d2 = st.slider("Date range", min_value=min_d, max_value=max_d, value=(min_d, max_d))
        surface_df = surface_df.loc[str(d1):str(d2)]

    # ── Stats ──────────────────────────────────────────────
    stale_mask = detect_stale(surface_df, tolerance=tolerance, lookback=lookback)
    stale_horizon = detect_stale_horizon(surface_df, tolerance=tolerance)
    stale_combined = stale_mask | stale_horizon
    streaks = compute_stale_streaks(surface_df, tolerance=tolerance)

    total_points = surface_df.size
    stale_date_count = int(stale_mask.values.sum())
    stale_hor_count = int(stale_horizon.values.sum())
    stale_total = int(stale_combined.values.sum())
    streak_alerts = int((streaks.values >= streak_threshold).sum())

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: render_stat_card(f"{len(surface_df)}", "Dates")
    with c2: render_stat_card(f"{len(surface_df.columns)}", "Horizons")
    with c3: render_stat_card(f"{total_points:,}", "Total Points")
    with c4:
        pct = stale_total / total_points * 100 if total_points else 0
        badge = "badge-stale" if pct > 10 else "badge-warn" if pct > 2 else "badge-ok"
        st.markdown(f"""<div class="stat-card">
            <div class="value">{pct:.1f}%</div>
            <div class="label">Stale <span class="{badge}">{stale_total}</span></div>
        </div>""", unsafe_allow_html=True)
    with c5:
        st.markdown(f"""<div class="stat-card">
            <div class="value">{streak_alerts}</div>
            <div class="label">Streak ≥ {streak_threshold}d</div>
        </div>""", unsafe_allow_html=True)

    # ── Horizon label mapping for display ──────────────────
    def horizon_label(h):
        for name, val in HORIZON_ORDER.items():
            if abs(val - h) < 1e-6:
                return name
        if h < 1:
            return f"{h*12:.0f}M"
        return f"{h:.0f}Y"

    h_labels = [horizon_label(h) for h in surface_df.columns]
    date_labels = [d.strftime('%d/%m/%Y') for d in surface_df.index]

    # ── 1) 3D Surface with stale overlay ────────────────────
    st.markdown('<div class="section-header">🌐 3D Yield Surface — y(date, horizon)</div>', unsafe_allow_html=True)

    Z = surface_df.values
    # Build color overlay: green = fresh, red = stale
    stale_intensity = np.zeros_like(Z, dtype=float)
    stale_intensity[stale_combined.values] = 1.0
    # Use streak length for intensity
    streak_vals = streaks.values.astype(float)
    max_streak = max(streak_vals.max(), 1)
    stale_intensity = np.where(stale_combined.values, streak_vals / max_streak, 0.0)

    # Custom colorscale: 0=blue (fresh) → 1=red (stale)
    surface_colors = np.zeros((*Z.shape, 4))
    for i in range(Z.shape[0]):
        for j in range(Z.shape[1]):
            s = stale_intensity[i, j]
            if s > 0:
                # Red with intensity
                surface_colors[i, j] = [0.95, 0.2 + 0.1*(1-s), 0.2, 0.6 + 0.4*s]
            else:
                surface_colors[i, j] = [0.23, 0.51, 0.96, 0.7]  # Blue fresh

    # Build customdata for 3D hover — pre-allocate object array to avoid
    # numpy dtype-inference issues when mixing strings and numerics via dstack
    _cd = np.empty((len(date_labels), len(h_labels), 3), dtype=object)
    _cd[:, :, 0] = [[date_labels[i]] * len(h_labels) for i in range(len(date_labels))]
    _cd[:, :, 1] = [h_labels] * len(date_labels)
    _cd[:, :, 2] = streaks.values

    fig_3d = go.Figure()

    # Fresh surface
    fig_3d.add_trace(go.Surface(
        z=Z,
        x=list(range(len(surface_df.columns))),
        y=list(range(len(surface_df.index))),
        surfacecolor=stale_intensity,
        colorscale=[
            [0.0, '#3b82f6'],   # Fresh — blue
            [0.3, '#60a5fa'],
            [0.5, '#fbbf24'],   # Warning — amber
            [0.7, '#f97316'],   # Orange
            [1.0, '#ef4444'],   # Stale — red
        ],
        colorbar=dict(title="Stale<br>streak", tickvals=[0, 0.5, 1],
                      ticktext=["Fresh", f"{max_streak//2}d", f"{int(max_streak)}d"]),
        hovertemplate=(
            "Date: %{customdata[0]}<br>"
            "Horizon: %{customdata[1]}<br>"
            "Rate: %{z:.6f}<br>"
            "Streak: %{customdata[2]}d"
            "<extra></extra>"
        ),
        customdata=_cd,
    ))

    fig_3d.update_layout(
        scene=dict(
            xaxis=dict(title="Horizon", tickvals=list(range(len(h_labels))),
                       ticktext=h_labels, gridcolor='#1e293b'),
            yaxis=dict(title="Date", tickvals=list(range(0, len(date_labels), max(1, len(date_labels)//8))),
                       ticktext=[date_labels[i] for i in range(0, len(date_labels), max(1, len(date_labels)//8))],
                       gridcolor='#1e293b'),
            zaxis=dict(title="Rate", gridcolor='#1e293b'),
            bgcolor='#0f172a',
        ),
        paper_bgcolor='rgba(0,0,0,0)',
        font=dict(family='DM Sans', color='#94a3b8'),
        margin=dict(l=0, r=0, t=30, b=0),
        height=600,
    )
    st.plotly_chart(fig_3d, use_container_width=True)

    # ── 2) Heatmap — Stale streaks ─────────────────────────
    st.markdown('<div class="section-header">🔥 Stale Streak Heatmap — y(date, horizon)</div>', unsafe_allow_html=True)

    fig_heat = go.Figure(go.Heatmap(
        z=streaks.values,
        x=h_labels,
        y=date_labels,
        colorscale=[
            [0.0, '#0f172a'],
            [0.01, '#1e40af'],
            [0.2, '#3b82f6'],
            [0.4, '#fbbf24'],
            [0.6, '#f97316'],
            [0.8, '#ef4444'],
            [1.0, '#991b1b'],
        ],
        colorbar=dict(title="Days<br>stale"),
        hovertemplate="Date: %{y}<br>Horizon: %{x}<br>Streak: %{z}d<extra></extra>",
    ))
    fig_heat.update_layout(
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#0f172a',
        font=dict(family='DM Sans', color='#94a3b8'),
        xaxis=dict(title="Horizon", gridcolor='#1e293b'),
        yaxis=dict(title="Date", gridcolor='#1e293b', autorange='reversed'),
        margin=dict(l=80, r=20, t=30, b=60),
        height=500,
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    # ── 3) 2D Slice explorer ────────────────────────────────
    st.markdown('<div class="section-header">🔬 2D Curve Slicer</div>', unsafe_allow_html=True)

    slice_mode = st.radio("Slice by", ["Fixed date (curve across horizons)", "Fixed horizon (rate across dates)"],
                          horizontal=True)

    if "Fixed date" in slice_mode:
        sel_date = st.select_slider("Date", options=surface_df.index,
                                     format_func=lambda d: d.strftime('%d/%m/%Y'))
        curve = surface_df.loc[sel_date]
        stale_row = stale_combined.loc[sel_date]
        streak_row = streaks.loc[sel_date]

        fig_slice = go.Figure()
        # Fresh points
        fresh_idx = [i for i, s in enumerate(stale_row) if not s]
        stale_idx = [i for i, s in enumerate(stale_row) if s]

        fig_slice.add_trace(go.Scatter(
            x=[h_labels[i] for i in fresh_idx],
            y=[curve.iloc[i] for i in fresh_idx],
            mode='markers+lines', name='Fresh',
            marker=dict(color='#3b82f6', size=8),
            line=dict(color='#3b82f6', width=2),
        ))
        if stale_idx:
            fig_slice.add_trace(go.Scatter(
                x=[h_labels[i] for i in stale_idx],
                y=[curve.iloc[i] for i in stale_idx],
                mode='markers', name='STALE',
                marker=dict(color='#ef4444', size=14, symbol='x',
                            line=dict(width=2, color='#fca5a5')),
                text=[f"Streak: {int(streak_row.iloc[i])}d" for i in stale_idx],
            ))

        fig_slice.update_layout(
            title=f"Curve as of {sel_date.strftime('%d/%m/%Y')}",
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#0f172a',
            font=dict(family='DM Sans', color='#94a3b8'),
            xaxis=dict(title="Horizon", gridcolor='#1e293b'),
            yaxis=dict(title="Rate", gridcolor='#1e293b'),
            margin=dict(l=60, r=20, t=50, b=60),
            height=400,
        )
        st.plotly_chart(fig_slice, use_container_width=True)

    else:
        sel_hor_idx = st.selectbox("Horizon", range(len(h_labels)), format_func=lambda i: h_labels[i])
        col_name = surface_df.columns[sel_hor_idx]
        ts = surface_df[col_name]
        stale_col = stale_combined.iloc[:, sel_hor_idx]
        streak_col = streaks.iloc[:, sel_hor_idx]

        fig_ts = go.Figure()
        # Background fill for stale zones
        stale_zones = []
        in_zone = False
        for i, s in enumerate(stale_col):
            if s and not in_zone:
                start = i
                in_zone = True
            elif not s and in_zone:
                stale_zones.append((start, i - 1))
                in_zone = False
        if in_zone:
            stale_zones.append((start, len(stale_col) - 1))

        for s, e in stale_zones:
            fig_ts.add_vrect(
                x0=surface_df.index[s], x1=surface_df.index[e],
                fillcolor='rgba(239,68,68,0.15)', line_width=0,
                annotation_text="STALE" if (e - s) >= 2 else "",
                annotation_font=dict(color='#fca5a5', size=10),
            )

        fig_ts.add_trace(go.Scatter(
            x=surface_df.index, y=ts,
            mode='lines+markers', name=h_labels[sel_hor_idx],
            line=dict(color='#3b82f6', width=2),
            marker=dict(size=4, color='#3b82f6'),
        ))

        fig_ts.update_layout(
            title=f"{h_labels[sel_hor_idx]} rate over time",
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='#0f172a',
            font=dict(family='DM Sans', color='#94a3b8'),
            xaxis=dict(title="Date", gridcolor='#1e293b'),
            yaxis=dict(title="Rate", gridcolor='#1e293b'),
            margin=dict(l=60, r=20, t=50, b=60),
            height=400,
        )
        st.plotly_chart(fig_ts, use_container_width=True)

    # ── 4) Stale report table ───────────────────────────────
    st.markdown('<div class="section-header">📋 Stale Data Report</div>', unsafe_allow_html=True)

    # Build flat table of stale points
    stale_records = []
    for i, date in enumerate(surface_df.index):
        for j, hor in enumerate(surface_df.columns):
            if stale_combined.iloc[i, j]:
                stale_records.append({
                    'Date': date.strftime('%d/%m/%Y'),
                    'Horizon': h_labels[j],
                    'Rate': surface_df.iloc[i, j],
                    'Streak (days)': int(streaks.iloc[i, j]),
                    'Stale on date': bool(stale_mask.iloc[i, j]),
                    'Stale on horizon': bool(stale_horizon.iloc[i, j]),
                })

    if stale_records:
        stale_report = pd.DataFrame(stale_records)
        # Filter by streak threshold
        severe = stale_report[stale_report['Streak (days)'] >= streak_threshold]

        tab1, tab2 = st.tabs([f"⚠️ Severe (streak ≥ {streak_threshold}d) — {len(severe)}",
                              f"All stale — {len(stale_report)}"])
        with tab1:
            if not severe.empty:
                st.dataframe(severe.sort_values('Streak (days)', ascending=False),
                            use_container_width=True, height=350)
            else:
                st.success(f"No points with streak ≥ {streak_threshold} days.")
        with tab2:
            st.dataframe(stale_report, use_container_width=True, height=350)

        # Download
        st.download_button("📥 Download stale report (CSV)",
                           stale_report.to_csv(index=False).encode(),
                           "stale_report.csv", "text/csv")
    else:
        st.success("✅ No stale data detected with current settings.")

    # ── 5) Raw surface ──────────────────────────────────────
    st.markdown('<div class="section-header">📄 Raw Surface Data</div>', unsafe_allow_html=True)
    with st.expander("Show raw y(date, horizon) matrix"):
        display_df = surface_df.copy()
        display_df.index = [d.strftime('%d/%m/%Y') for d in display_df.index]
        display_df.columns = h_labels
        st.dataframe(display_df, use_container_width=True, height=350)

    # ── 6) Formula Calculator (Excel + Custom) ───────────────
    st.markdown('<div class="section-header">🧮 Formula Calculator</div>', unsafe_allow_html=True)
    
    with st.expander("Create calculated columns with Excel or custom formulas", expanded=False):
        # Initialize session state for formulas if not exists
        if 'calculated_columns' not in st.session_state:
            st.session_state.calculated_columns = {}
        
        # Formula input
        col1, col2 = st.columns([3, 1])
        with col1:
            formula = st.text_input(
                "Formula",
                placeholder="e.g., SPREAD(0, 8) for 10Y-1M spread",
                help="Use Excel functions (SUM, AVERAGE) or custom yield curve functions (SPREAD, BUTTERFLY)"
            )
        with col2:
            col_name = st.text_input("Column name", placeholder="e.g., 10Y-1M_Spread")
        
        # Quick reference
        with st.expander("📖 Available Functions"):
            tab1, tab2 = st.tabs(["Excel Functions", "Yield Curve Functions"])
            with tab1:
                st.markdown("""
                **Excel-compatible:**
                - `SUM(0:11)` - Sum across all tenors for each date
                - `AVERAGE(0, 2, 4)` - Average of specific tenors
                - `MIN(0:5)` / `MAX(0:5)` - Min/max across range
                - `STD(0:11)` - Standard deviation
                - `0 + 1` / `8 - 0` - Arithmetic operations on columns
                """)
            with tab2:
                st.markdown("""
                **Custom Yield Curve:**
                - `SPREAD(short, long)` - e.g., `SPREAD(0, 8)` = 10Y - 1M
                - `BUTTERFLY(s, m, l)` - 2×mid - short - long
                - `CARRY(tenor, funding)` - Carry return
                - `ROLLDOWN(tenor, days)` - Expected rolldown
                - `DELTA_Y(tenor, lookback)` - Rate change over N days
                - `ZSCORE(tenor, window)` - Rolling Z-score
                - `NORMALIZE(tenor)` - Min-max to [0,1]
                - `CHANGE(tenor, periods)` / `PCT_CHANGE(tenor, periods)`
                """)
                st.caption(f"**Horizon indices:** {', '.join([f'{i}={h}' for i, h in enumerate(h_labels)])}")
        
        # Apply formula
        c1, c2 = st.columns([1, 3])
        with c1:
            if st.button("▶ Calculate", use_container_width=True, disabled=not formula):
                if formula and col_name:
                    engine = FormulaEngine(surface_df, h_labels)
                    result = engine.evaluate(formula)
                    
                    if isinstance(result, pd.Series):
                        st.session_state.calculated_columns[col_name] = result
                        st.success(f"✅ Created column '{col_name}'")
                    else:
                        st.error(f"❌ {result}")
                elif formula:
                    st.warning("Please enter a column name")
        
        with c2:
            if st.session_state.calculated_columns:
                if st.button("🗑️ Clear All Calculated", use_container_width=False):
                    st.session_state.calculated_columns = {}
                    st.rerun()
        
        # Show calculated columns
        if st.session_state.calculated_columns:
            st.markdown("**Calculated Columns:**")
            calc_df = pd.DataFrame(st.session_state.calculated_columns)
            calc_df.index = surface_df.index
            st.dataframe(calc_df, use_container_width=True, height=200)
            
            # Download calculated data
            output_cols = list(surface_df.columns) + list(st.session_state.calculated_columns.keys())
            combined_df = surface_df.copy()
            for name, series in st.session_state.calculated_columns.items():
                combined_df[name] = series
            
            st.download_button(
                "📥 Download with calculated columns",
                combined_df.to_csv().encode(),
                "surface_with_calculated.csv",
                use_container_width=True
            )

else:
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#64748b;">
        <div style="font-size:4rem; margin-bottom:1rem;">📈</div>
        <h2 style="color:#94a3b8;">No yield curve data loaded</h2>
        <p style="max-width:500px; margin:0 auto; font-size:0.9rem;">
            Upload daily curve CSVs or a consolidated file from the sidebar.<br><br>
            <b>Consolidated format:</b> first column = date (JJMMYYYY), other columns = horizons (1M, 3M, 1Y, …)<br><br>
            <b>Multiple files:</b> one file per date named <code>dataJJMMYYYY.csv</code>, each with columns <code>horizon, rate</code>
        </p>
    </div>
    """, unsafe_allow_html=True)
