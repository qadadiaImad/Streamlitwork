"""
IPV EQD Dashboard — Main Hub
=============================
Multi-module platform for Excel/CSV tools.
Streamlit multi-page app: this is the landing page.
"""
import streamlit as st
import sys
from pathlib import Path

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from modules.shared import inject_css, render_header

st.set_page_config(
    page_title="IPV EQD Dashboard",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
)

inject_css()

render_header("🏦", "IPV EQD Dashboard", "Excel · CSV · Yield Curves · Stale Detection")

st.markdown("""
<div style="
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 1.5rem;
    margin-top: 1rem;
">
    <a href="/Excel_Tools" target="_self" style="text-decoration: none;">
        <div style="
            background: #1e293b; border: 1px solid #334155; border-radius: 12px;
            padding: 2rem; cursor: pointer; transition: border-color 0.2s;
        ">
            <div style="font-size: 2.5rem; margin-bottom: 0.75rem;">📊</div>
            <h3 style="color: #e2e8f0; margin: 0 0 0.5rem 0; font-size: 1.1rem;">Excel Tools</h3>
            <p style="color: #64748b; font-size: 0.85rem; margin: 0;">
                Open, visualize, edit Excel & CSV files.<br>
                Detect and execute VBA macros.
            </p>
        </div>
    </a>
    <a href="/Yield_Curve_Stale_Detection" target="_self" style="text-decoration: none;">
        <div style="
            background: #1e293b; border: 1px solid #334155; border-radius: 12px;
            padding: 2rem; cursor: pointer; transition: border-color 0.2s;
        ">
            <div style="font-size: 2.5rem; margin-bottom: 0.75rem;">📈</div>
            <h3 style="color: #e2e8f0; margin: 0 0 0.5rem 0; font-size: 1.1rem;">Yield Curve — Stale Detection</h3>
            <p style="color: #64748b; font-size: 0.85rem; margin: 0;">
                Load daily curve CSVs, build y(date, horizon) surface.<br>
                Detect & highlight stale data on 3D graph.
            </p>
        </div>
    </a>
</div>
""", unsafe_allow_html=True)

st.markdown("---")
st.caption("Use the sidebar **← pages** to navigate between modules.")
