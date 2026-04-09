"""
Shared utilities and styling for the IPV EQD Dashboard platform.
"""
import streamlit as st
import pandas as pd
import tempfile
from pathlib import Path

# --- Try importing xlwings (Windows + Excel only) ---
try:
    import xlwings as xw
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False


def inject_css():
    """Inject shared CSS theme across all pages."""
    st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=JetBrains+Mono:wght@400;500&display=swap');
        .stApp { font-family: 'DM Sans', sans-serif; }

        .dashboard-header {
            background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
            border: 1px solid #334155; border-radius: 12px;
            padding: 1.5rem 2rem; margin-bottom: 1.5rem;
            display: flex; align-items: center; gap: 1rem;
        }
        .dashboard-header h1 { color: #f1f5f9; font-size: 1.6rem; font-weight: 700; margin: 0; letter-spacing: -0.02em; }
        .dashboard-header .subtitle { color: #64748b; font-size: 0.85rem; margin: 0; }
        .dashboard-header .icon {
            font-size: 2rem;
            background: linear-gradient(135deg, #3b82f6, #8b5cf6);
            -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        }

        .stat-card {
            background: #1e293b; border: 1px solid #334155; border-radius: 10px;
            padding: 1rem 1.2rem; text-align: center;
        }
        .stat-card .value { font-family: 'JetBrains Mono', monospace; font-size: 1.4rem; font-weight: 700; color: #e2e8f0; }
        .stat-card .label { font-size: 0.75rem; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; margin-top: 0.25rem; }

        .macro-card {
            background: #1e293b; border: 1px solid #334155; border-radius: 8px;
            padding: 0.8rem 1rem; margin-bottom: 0.5rem;
            display: flex; align-items: center; justify-content: space-between;
        }
        .macro-name { font-family: 'JetBrains Mono', monospace; font-size: 0.85rem; color: #a5b4fc; }
        .macro-type { font-size: 0.7rem; color: #64748b; background: #0f172a; padding: 2px 8px; border-radius: 4px; }

        .badge-ok { background: #065f46; color: #6ee7b7; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 500; }
        .badge-warn { background: #713f12; color: #fcd34d; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 500; }
        .badge-info { background: #1e3a5f; color: #7dd3fc; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 500; }
        .badge-stale { background: #7f1d1d; color: #fca5a5; padding: 4px 12px; border-radius: 20px; font-size: 0.75rem; font-weight: 500; }

        .section-header {
            color: #94a3b8; font-size: 0.8rem; font-weight: 600;
            text-transform: uppercase; letter-spacing: 0.08em;
            margin: 1.5rem 0 0.75rem 0; padding-bottom: 0.4rem;
            border-bottom: 1px solid #1e293b;
        }

        section[data-testid="stSidebar"] { background: #0f172a; }
        section[data-testid="stSidebar"] .stMarkdown { color: #cbd5e1; }
        header[data-testid="stHeader"] { background: transparent; }
        .stDataFrame { border-radius: 8px; overflow: hidden; }
    </style>
    """, unsafe_allow_html=True)


def get_file_type(filename: str) -> str:
    ext = Path(filename).suffix.lower()
    if ext in ('.xlsx', '.xlsm', '.xls', '.xlsb'):
        return 'excel'
    elif ext in ('.csv', '.txt', '.tsv'):
        return 'csv'
    return 'unknown'


def get_engine_for_ext(file_ext: str) -> str:
    return 'pyxlsb' if file_ext == '.xlsb' else 'openpyxl'


def load_excel_file(uploaded_file) -> dict[str, pd.DataFrame]:
    filename = getattr(uploaded_file, 'name', 'file.xlsx')
    file_ext = Path(filename).suffix.lower()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=file_ext)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    st.session_state['temp_file_path'] = tmp.name
    st.session_state['original_filename'] = filename

    engine = get_engine_for_ext(file_ext)
    xls = pd.ExcelFile(tmp.name, engine=engine)
    sheets = {}
    for sheet_name in xls.sheet_names:
        sheets[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
    return sheets


def load_csv_file(uploaded_file, separator=',') -> dict[str, pd.DataFrame]:
    filename = getattr(uploaded_file, 'name', 'file.csv')
    df = pd.read_csv(uploaded_file, sep=separator)
    st.session_state['original_filename'] = filename
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=Path(filename).suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    st.session_state['temp_file_path'] = tmp.name
    return {'Sheet1': df}


def render_header(icon: str, title: str, subtitle: str):
    st.markdown(f"""
    <div class="dashboard-header">
        <div class="icon">{icon}</div>
        <div>
            <h1>{title}</h1>
            <p class="subtitle">{subtitle}</p>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_stat_card(value, label):
    st.markdown(f"""
    <div class="stat-card">
        <div class="value">{value}</div>
        <div class="label">{label}</div>
    </div>""", unsafe_allow_html=True)
