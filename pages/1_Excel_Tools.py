"""
Page 1 — Excel Tools
Open, visualize, edit Excel/CSV files. Detect & run VBA macros.
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import tempfile
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent.parent))
from modules.shared import (
    inject_css, render_header, render_stat_card,
    get_file_type, get_engine_for_ext,
    load_excel_file, load_csv_file, XLWINGS_AVAILABLE,
)

inject_css()
render_header("📊", "Excel Tools", "Open · Visualize · Edit · Execute VBA")

# ── Macro helpers ──────────────────────────────────────────
def detect_macros_openpyxl(file_path):
    from openpyxl import load_workbook
    macros = []
    try:
        wb = load_workbook(file_path, keep_vba=True)
        if wb.vba_archive is not None:
            for name in wb.vba_archive.namelist():
                if name.startswith('VBA/') and not name.endswith('/'):
                    mod = name.replace('VBA/', '').replace('.bin', '')
                    skip = ('_VBA_PROJECT','dir','__SRP_0','__SRP_1','__SRP_2','__SRP_3','PROJECTwm','PROJECT')
                    if mod not in skip:
                        macros.append(mod)
        wb.close()
    except Exception as e:
        st.warning(f"openpyxl VBA inspect failed: {e}")
    return macros

def detect_macros_xlwings(file_path):
    if not XLWINGS_AVAILABLE:
        return []
    import xlwings as xw
    macros = []
    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        for component in wb.api.VBProject.VBComponents:
            code_module = component.CodeModule
            if code_module.CountOfLines > 0:
                code_text = code_module.Lines(1, code_module.CountOfLines)
                for line in code_text.split('\n'):
                    ls = line.strip()
                    if ls.startswith(('Sub ','Public Sub ')):
                        proc = ls.split('(')[0].replace('Sub ','').replace('Public ','').strip()
                        macros.append({'name': proc, 'type': 'Sub', 'module': component.Name})
                    elif ls.startswith(('Function ','Public Function ')):
                        proc = ls.split('(')[0].replace('Function ','').replace('Public ','').strip()
                        macros.append({'name': proc, 'type': 'Function', 'module': component.Name})
        wb.close(); app.quit()
    except Exception as e:
        st.warning(f"xlwings macro detection failed: {e}")
    return macros

def execute_macro_xlwings(file_path, macro_name):
    if not XLWINGS_AVAILABLE:
        return "❌ xlwings not available"
    import xlwings as xw
    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        result = wb.macro(macro_name)()
        wb.save(); wb.close(); app.quit()
        return f"✅ '{macro_name}' executed. Return: {result}"
    except Exception as e:
        return f"❌ Failed: {e}"

def save_sheets_to_excel(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False)
    return output.getvalue()

# ── Sidebar ────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📁 Open File")
    uploaded_file = st.file_uploader(
        "Drop your file here",
        type=['xlsx','xlsm','xlsb','xls','csv','txt','tsv'],
        key="excel_uploader"
    )
    separator = st.text_input("CSV/TXT separator", value=",", max_chars=5)

# ── Main ───────────────────────────────────────────────────
if uploaded_file is not None:
    filename = getattr(uploaded_file, 'name', 'file.xlsx')
    file_type = get_file_type(filename)
    file_ext = Path(filename).suffix.lower()

    if 'xl_sheets' not in st.session_state or st.session_state.get('xl_loaded') != filename:
        with st.spinner("Loading..."):
            try:
                if file_type == 'excel':
                    st.session_state['xl_sheets'] = load_excel_file(uploaded_file)
                elif file_type == 'csv':
                    st.session_state['xl_sheets'] = load_csv_file(uploaded_file, separator)
                else:
                    st.error("Unsupported format."); st.stop()
                st.session_state['xl_loaded'] = filename
            except Exception as e:
                st.error(f"Load failed: {e}"); st.stop()

    sheets = st.session_state['xl_sheets']
    sheet_names = list(sheets.keys())
    is_macro_file = file_ext in ('.xlsm', '.xlsb')

    # Stats
    total_rows = sum(len(df) for df in sheets.values())
    total_cols = max(len(df.columns) for df in sheets.values())
    c1, c2, c3, c4 = st.columns(4)
    with c1: render_stat_card(len(sheet_names), "Sheets")
    with c2: render_stat_card(f"{total_rows:,}", "Total Rows")
    with c3: render_stat_card(total_cols, "Max Columns")
    with c4:
        badge = '<span class="badge-ok">VBA</span>' if is_macro_file else '<span class="badge-info">Data</span>'
        render_stat_card(f'{file_ext} {badge}', "Format")

    # Sheet selector
    st.markdown('<div class="section-header">📄 Sheet Navigation</div>', unsafe_allow_html=True)
    selected_sheet = st.selectbox("Sheet", sheet_names, label_visibility="collapsed") if len(sheet_names) > 1 else sheet_names[0]
    df = sheets[selected_sheet].copy()

    # Editor
    st.markdown('<div class="section-header">✏️ Data Editor</div>', unsafe_allow_html=True)
    with st.expander("🔍 Column filter", expanded=False):
        sel_cols = st.multiselect("Columns", df.columns.tolist(), default=df.columns.tolist())
    edited_df = st.data_editor(df[sel_cols] if sel_cols else df, use_container_width=True, num_rows="dynamic", height=400)
    sheets[selected_sheet] = edited_df
    st.session_state['xl_sheets'] = sheets

    # Stats
    numeric_cols = edited_df.select_dtypes(include='number').columns.tolist()
    if numeric_cols:
        st.markdown('<div class="section-header">📈 Statistics</div>', unsafe_allow_html=True)
        with st.expander("Summary", expanded=True):
            st.dataframe(edited_df[numeric_cols].describe().round(4), use_container_width=True)

    # Viz
    if numeric_cols:
        st.markdown('<div class="section-header">📊 Visualization</div>', unsafe_allow_html=True)
        v1, v2 = st.columns(2)
        with v1:
            chart_type = st.selectbox("Chart", ["Histogram","Box Plot","Scatter","Line","Bar","Correlation Heatmap"])
        with v2:
            if chart_type in ("Histogram","Box Plot"):
                viz_col = st.selectbox("Column", numeric_cols, key="vc")
            elif chart_type == "Scatter":
                sc = st.multiselect("X, Y", numeric_cols, default=numeric_cols[:2], max_selections=2, key="sc")
            elif chart_type in ("Line","Bar"):
                vy = st.multiselect("Y", numeric_cols, default=[numeric_cols[0]], key="vy")

        colors = {'p':'#3b82f6','s':'#8b5cf6','bg':'#0f172a','g':'#1e293b','t':'#94a3b8'}
        lo = dict(paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor=colors['bg'],
                  font=dict(family='DM Sans', color=colors['t']),
                  xaxis=dict(gridcolor=colors['g']), yaxis=dict(gridcolor=colors['g']),
                  margin=dict(l=40,r=20,t=40,b=40))
        fig = None
        if chart_type == "Histogram":
            fig = px.histogram(edited_df, x=viz_col, nbins=30, color_discrete_sequence=[colors['p']])
        elif chart_type == "Box Plot":
            fig = px.box(edited_df, y=viz_col, color_discrete_sequence=[colors['s']])
        elif chart_type == "Scatter" and len(sc) == 2:
            fig = px.scatter(edited_df, x=sc[0], y=sc[1], color_discrete_sequence=[colors['p']], opacity=0.7)
        elif chart_type == "Line" and vy:
            fig = go.Figure()
            for i, c in enumerate(vy):
                fig.add_trace(go.Scatter(y=edited_df[c], mode='lines', name=c))
        elif chart_type == "Correlation Heatmap":
            fig = px.imshow(edited_df[numeric_cols].corr(), text_auto='.2f', color_continuous_scale='RdBu_r', zmin=-1, zmax=1)
        if fig:
            fig.update_layout(**lo)
            st.plotly_chart(fig, use_container_width=True)

    # VBA
    st.markdown('<div class="section-header">⚡ VBA Macros</div>', unsafe_allow_html=True)
    if is_macro_file:
        temp_path = st.session_state.get('temp_file_path', '')
        if file_ext == '.xlsm':
            mods = detect_macros_openpyxl(temp_path)
            if mods:
                for m in mods:
                    st.markdown(f'<div class="macro-card"><span class="macro-name">{m}</span><span class="macro-type">Module</span></div>', unsafe_allow_html=True)
        elif file_ext == '.xlsb':
            st.markdown('<span class="badge-info">.xlsb — VBA detection via xlwings only</span>', unsafe_allow_html=True)

        if XLWINGS_AVAILABLE:
            xlm = detect_macros_xlwings(temp_path)
            if xlm:
                for m in xlm:
                    st.markdown(f'<div class="macro-card"><span class="macro-name">{m["module"]}.{m["name"]}</span><span class="macro-type">{m["type"]}</span></div>', unsafe_allow_html=True)
                subs = [f'{m["module"]}.{m["name"]}' for m in xlm if m['type'] == 'Sub']
                if subs:
                    sel = st.selectbox("Macro to run", subs)
                    if st.button("▶ Run", type="primary"):
                        with st.spinner("Executing..."):
                            res = execute_macro_xlwings(temp_path, sel.split('.')[-1])
                            st.info(res)
                            try:
                                eng = get_engine_for_ext(file_ext)
                                xls = pd.ExcelFile(temp_path, engine=eng)
                                for sn in xls.sheet_names:
                                    st.session_state['xl_sheets'][sn] = pd.read_excel(xls, sheet_name=sn)
                                st.rerun()
                            except: pass

        manual = st.text_input("Manual macro name", placeholder="Module1.MyMacro")
        if st.button("▶ Execute", key="man") and manual and XLWINGS_AVAILABLE:
            st.info(execute_macro_xlwings(temp_path, manual))
    else:
        st.caption("Upload `.xlsm` or `.xlsb` for VBA features.")

    # Export
    st.markdown('<div class="section-header">💾 Export</div>', unsafe_allow_html=True)
    e1, e2, e3 = st.columns(3)
    with e1:
        st.download_button("📥 Excel", save_sheets_to_excel(sheets),
                           Path(filename).stem + '_modified.xlsx',
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with e2:
        st.download_button("📥 CSV", edited_df.to_csv(index=False).encode(), f"{selected_sheet}.csv", "text/csv")
    with e3:
        st.download_button("📥 JSON", edited_df.to_json(orient='records', indent=2).encode(), f"{selected_sheet}.json")

else:
    st.markdown("""
    <div style="text-align:center; padding:4rem 2rem; color:#64748b;">
        <div style="font-size:4rem; margin-bottom:1rem;">📂</div>
        <h2 style="color:#94a3b8;">No file loaded</h2>
        <p>Upload an Excel or CSV file from the sidebar.</p>
    </div>""", unsafe_allow_html=True)
