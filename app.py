"""
Excel Transfer App
==================

Supports two workflows:
1. Report -> Cost Control (Weekly Transfer)
2. Cost Control -> PPC (Cumulative Transfer)
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from excel_io import load_excel
from excel_layout import detect_week_blocks
from actions.report_to_cost import ReportToCostAction
from actions.cost_to_ppc import CostToPPCAction

st.set_page_config(page_title="Excel Transfer Tool", layout="wide")

def render_action_1():
    """Report -> Cost Control"""
    st.header("Action: Report -> Cost Control")
    st.markdown("Transfer **Planned**, **Actual**, and **Timesheet** data based on Weekly Blocks.")

    # 1. Uploads
    with st.sidebar:
        st.subheader("1. Upload Files")
        src_file = st.file_uploader("Upload REPORT (Source)", type=["xlsx", "xlsm"], key="a1_src")
        tgt_file = st.file_uploader("Upload COST CONTROL (Target)", type=["xlsx", "xlsm"], key="a1_tgt")

    if not src_file or not tgt_file:
        st.info("Upload both files to start.")
        return

    # Load Workbooks
    if 'a1_wb_src' not in st.session_state or st.session_state.get('a1_src_name') != src_file.name:
        st.session_state.a1_wb_src = load_excel(src_file)
        st.session_state.a1_src_name = src_file.name
        st.toast("Source loaded")
    
    if 'a1_wb_tgt' not in st.session_state or st.session_state.get('a1_tgt_name') != tgt_file.name:
        st.session_state.a1_wb_tgt = load_excel(tgt_file)
        st.session_state.a1_tgt_name = tgt_file.name
        st.toast("Target loaded")

    wb_src = st.session_state.a1_wb_src
    wb_tgt = st.session_state.a1_wb_tgt
    
    if not wb_src or not wb_tgt:
        return

    # 2. Config
    with st.sidebar:
        st.subheader("2. Settings")
        src_sheet = st.selectbox("Source Sheet", wb_src.sheetnames, key="a1_ss")
        tgt_sheet = st.selectbox("Target Sheet", wb_tgt.sheetnames, key="a1_ts")

        ws_src = wb_src[src_sheet]
        ws_tgt = wb_tgt[tgt_sheet]

        # Detect Weeks
        try:
            src_weeks = detect_week_blocks(ws_src)
            tgt_weeks = detect_week_blocks(ws_tgt)
        except Exception as e:
            st.error(f"Error parsing layout: {e}")
            return

        if not src_weeks:
            st.warning("No week blocks found in Source.")
            return

        week_options = [w.label for w in src_weeks]
        selected_weeks = st.multiselect("Select Week(s)", week_options)

        overwrite_formulas = st.checkbox("Overwrite Formulas", value=False, key="a1_of")
        write_first = st.checkbox("Write to FIRST match only (Default: All)", value=False, key="a1_wf")
        
        btn_analyze = st.button("Analyze", key="a1_btn_an")
        btn_transfer = st.button("Transfer Selected", type="primary", key="a1_btn_tr")

    # Main Logic
    if selected_weeks:
        # Match Blocks
        week_pairs = []
        tgt_week_map = {w.label: w for w in tgt_weeks} # Exact match map
        
        missing = []
        for sl in selected_weeks:
            # Find source block object
            s_blk = next(w for w in src_weeks if w.label == sl)
            # Find target block object (loose matching logic?)
            # The original logic had extract_week_number fallback.
            # Let's simple check strict match first, then number match logic?
            # Re-implement simple matching here or use helper?
            # Using map is strict. Let's do a quick loop for loose match if not exact.
            
            t_blk = tgt_week_map.get(sl)
            if not t_blk:
                # Try number matching?
                import re
                s_num = re.search(r'\d+', sl)
                if s_num:
                    s_n = int(s_num.group(0))
                    for t_candidate in tgt_weeks:
                        t_num = re.search(r'\d+', t_candidate.label)
                        if t_num and int(t_num.group(0)) == s_n:
                            t_blk = t_candidate
                            break
            
            if s_blk and t_blk:
                week_pairs.append((s_blk, t_blk))
            else:
                missing.append(sl)
        
        if missing:
            st.error(f"Targets not found for: {', '.join(missing)}")

        if week_pairs:
            action = ReportToCostAction(ws_src, ws_tgt, week_pairs, overwrite_formulas, write_first)
            
            if btn_analyze or btn_transfer:
                with st.spinner("Analyzing..."):
                    diffs = action.analyze()
                
                # Stats
                n_write = sum(1 for d in diffs if d.action == "Write")
                st.metric("Changes Detected", n_write, delta=f"{len(diffs)-n_write} Skipped")
                
                # Table
                if diffs:
                    data = [{
                        "Week": d.week_label,
                        "Section": d.key[0],
                        "Material": d.key[1],
                        "Action": d.action,
                        "Reason": d.reason,
                        "Src Plan": d.src_planned,
                        "Tgt Plan": d.tgt_planned
                    } for d in diffs]
                    st.dataframe(pd.DataFrame(data), use_container_width=True)
                else:
                    st.info("No matches found.")

                if btn_transfer:
                    with st.spinner("Transferring..."):
                        tgt_map_combined = {blk.label: blk for blk in tgt_weeks}
                        # We also need to map matched labels if loose match was used?
                        # The action 'execute' uses diff.week_label.
                        # diff.week_label comes from src_wk.label.
                        # We need to ensure logic maps src_label -> tgt_block correctly during execute.
                        # Actually 'execute' needs 'tgt_week_bfs'.
                        # Let's update 'tgt_week_bfs' with the pairings we found.
                        
                        # Build specific map for this batch
                        batch_map = {}
                        for s_blk, t_blk in week_pairs:
                            batch_map[s_blk.label] = t_blk
                        
                        count = action.execute(diffs, batch_map)
                        
                        # Download
                        out = BytesIO()
                        st.session_state.a1_wb_tgt.save(out)
                        out.seek(0)
                        st.success(f"Transferred {count} records!")
                        st.download_button("Download Result", out, "cost_control_updated.xlsx")

def render_action_2():
    """Cost Control -> PPC"""
    st.header("Action: Cost Control -> PPC")
    st.markdown("Transfer **DONE** quantities (Material matches only).")

    with st.sidebar:
        st.subheader("1. Upload Files")
        src_file = st.file_uploader("Upload COST CONTROL (Source)", type=["xlsx", "xlsm"], key="a2_src")
        tgt_file = st.file_uploader("Upload PPC (Target)", type=["xlsx", "xlsm"], key="a2_tgt")
    
    if not src_file or not tgt_file:
        st.info("Upload files.")
        return

    # Load
    if 'a2_wb_src' not in st.session_state or st.session_state.get('a2_src_name') != src_file.name:
        # DATA ONLY TRUE for source to get calculated values
        st.session_state.a2_wb_src = load_excel(src_file, data_only=True)
        st.session_state.a2_src_name = src_file.name
        st.toast("Source loaded (Values only)")

    if 'a2_wb_tgt' not in st.session_state or st.session_state.get('a2_tgt_name') != tgt_file.name:
        st.session_state.a2_wb_tgt = load_excel(tgt_file, data_only=False)
        st.session_state.a2_tgt_name = tgt_file.name
        st.toast("Target loaded")

    wb_src = st.session_state.a2_wb_src
    wb_tgt = st.session_state.a2_wb_tgt

    if not wb_src or not wb_tgt:
        return

    with st.sidebar:
        st.subheader("2. Config")
        # Source Sheet
        src_sheet_name = st.selectbox("Source Sheet", wb_src.sheetnames, key="a2_ss")
        ws_src = wb_src[src_sheet_name]

        # Target Sheets (Multi)
        tgt_sheets_sel = st.multiselect("Target Sheets", wb_tgt.sheetnames, default=wb_tgt.sheetnames[:1], key="a2_ts")
        target_sheets = [wb_tgt[n] for n in tgt_sheets_sel]

        use_section = st.checkbox("Match by Section too? (Optional)", value=False, help="Requires PPC to have Section headers.", key="a2_use_sec")
        overwrite_formulas = st.checkbox("Overwrite Formulas", value=False, key="a2_of")
        
        btn_analyze = st.button("Analyze", key="a2_btn_an")
        btn_transfer = st.button("Transfer", type="primary", key="a2_btn_tr")

    if target_sheets and (btn_analyze or btn_transfer):
        action = CostToPPCAction(ws_src, target_sheets, use_section, overwrite_formulas)
        
        with st.spinner("Analyzing..."):
            diffs = action.analyze()
        
        st.metric("Potential Updates", len([d for d in diffs if d.action=="Write"]))
        
        if diffs:
            df = pd.DataFrame([{
                "Sheet": d.sheet_name,
                "Key": d.key,
                "Action": d.action,
                "Source Val": d.src_val,
                "Target Val (Curr)": d.tgt_val
            } for d in diffs])
            st.dataframe(df, use_container_width=True)
        
        if btn_transfer:
            with st.spinner("Writing..."):
                cnt = action.execute(diffs)
                
                out = BytesIO()
                st.session_state.a2_wb_tgt.save(out)
                out.seek(0)
                st.success(f"Updated {cnt} cells.")
                st.download_button("Download PPC", out, "ppc_updated.xlsx")

def main():
    st.sidebar.title("Configuration")
    mode = st.sidebar.radio("Select Action workflow:", ["Report -> Cost Control", "Cost Control -> PPC"])
    
    st.sidebar.markdown("---")
    
    if mode == "Report -> Cost Control":
        render_action_1()
    else:
        render_action_2()

if __name__ == "__main__":
    main()
