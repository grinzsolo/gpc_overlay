import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import os

st.set_page_config(page_title="GPC Multi-File Overlay Dashboard", layout="wide")

# Custom CSS for a clean, professional, and modern laboratory dashboard style
st.markdown("""
    <style>
        .main { background-color: #f8f9fa; }
        .block-container { padding-top: 1.5rem; padding-bottom: 1.5rem; }
        h1 { color: #0f172a; font-weight: 700; font-size: 2.2rem; margin-bottom: 0.2rem; }
        h3 { color: #334155; font-weight: 600; margin-top: 1rem; }
        div.stButton > button:first-child {
            background-color: #2563eb; color: white; border-radius: 6px; border: none;
            padding: 0.5rem 2rem; font-weight: 500;
        }
        .stDataFrame { border: 1px solid #e2e8f0; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 GPC Multi-File Overlay Dashboard")
st.write("Upload GPC files to overlay molecular profiles with advanced chart-embedded Excel exporting tools.")
st.markdown("---")

# Initialize Session States to prevent data from disappearing on download
if "data_mmd_list" not in st.session_state:
    st.session_state.data_mmd_list = []
if "results_list" not in st.session_state:
    st.session_state.results_list = []
if "max_scb_value" not in st.session_state:
    st.session_state.max_scb_value = 0.0
if "global_min_logm" not in st.session_state:
    st.session_state.global_min_logm = 0.0
if "global_max_logm" not in st.session_state:
    st.session_state.global_max_logm = 0.0

with st.form(key="gpc_upload_form"):
    uploaded_files = st.file_uploader(
        "Select GPC Files (.xls / .xlsx)", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    submit_button = st.form_submit_button(label="🚀 Process and Overlay Data")

# Process file calculations upon submit
if submit_button and uploaded_files:
    if len(uploaded_files) > 5:
        st.error("⚠️ Maximum 5 files allowed. Please remove excess files and submit again.")
    else:
        # Clear previous session records
        st.session_state.data_mmd_list = []
        st.session_state.results_list = []
        st.session_state.max_scb_value = 0.0
        
        all_min_logm = []
        all_max_logm = []
        
        for file in uploaded_files:
            # Clean filename by removing extension (.xls / .xlsx)
            file_name_clean = os.path.splitext(file.name)[0]
            
            try:
                excel_file = pd.ExcelFile(file)
                
                # Read Sheets safely
                df_mmd = pd.read_excel(file, sheet_name="Data MMD" if "Data MMD" in excel_file.sheet_names else 0)
                df_res = pd.read_excel(file, sheet_name="Results" if "Results" in excel_file.sheet_names else 1, header=None)
                
                df_mmd.columns = [str(c).strip() for c in df_mmd.columns]
                
                st.session_state.data_mmd_list.append({
                    "file_name": file_name_clean,
                    "df": df_mmd
                })
                
                # Trace dynamic Min/Max of LogM for customized layout widths
                for col in df_mmd.columns:
                    if 'logm' in col.lower():
                        numeric_logm = pd.to_numeric(df_mmd[col], errors='coerce').dropna()
                        if not numeric_logm.empty:
                            all_min_logm.append(numeric_logm.min())
                            all_max_logm.append(numeric_logm.max())
                
                # Trace highest SCB values across the files
                cols_lower_temp = [c.lower() for c in df_mmd.columns]
                col_scb_idx_temp = next((idx for idx, c in enumerate(cols_lower_temp) if 'scb' in c or '1000tc' in c), None)
                if col_scb_idx_temp is not None:
                    current_max = pd.to_numeric(df_mmd.iloc[:, col_scb_idx_temp], errors='coerce').max()
                    if pd.notna(current_max) and current_max > st.session_state.max_scb_value:
                        st.session_state.max_scb_value = current_max
                
                res_dict = {"Sample Name": file_name_clean}
                target_metrics = [
                    "Mw", "Mn", "Mw / Mn", "Mz", "Mz1", "Mv", "Mp", "IV", 
                    "Amount of material", "Bulk CH3 / 1000TC", "Bulk SCB / 1000TC", "Bulk Comonomer"
                ]
                
                for metric in target_metrics:
                    val = None
                    for row_idx, row in df_res.iterrows():
                        for col_idx, cell_value in enumerate(row):
                            if str(cell_value).strip() == metric:
                                if col_idx + 1 < len(row):
                                    val = row.iloc[col_idx + 1]
                                    if pd.isna(val) and col_idx + 2 < len(row):
                                        val = row.iloc[col_idx + 2]
                                break
                        if val is not None:
                            break
                    
                    try:
                        if val is not None and not isinstance(val, str):
                            val_float = float(val)
                            # Remove decimals completely for standard Mw, Mn, Mz, Mz1, Mv, Mp fields
                            if metric in ["Mw", "Mn", "Mz", "Mz1", "Mv", "Mp"]:
                                res_dict[metric] = int(round(val_float))
                            else:
                                res_dict[metric] = val_float
                        else:
                            res_dict[metric] = val
                    except:
                        res_dict[metric] = val
                        
                st.session_state.results_list.append(res_dict)
                
            except Exception as e:
                st.error(f"Error reading file {file.name}: {e}")
                
        if all_min_logm and all_max_logm:
            st.session_state.global_min_logm = min(all_min_logm) - 2.0
            st.session_state.global_max_logm = max(all_max_logm) + 2.0

# Render Report Blocks if session data holds true
if st.session_state.results_list:
    df_summary = pd.DataFrame(st.session_state.results_list)
    df_summary.set_index("Sample Name", inplace=True)
    df_summary_transposed = df_summary.T
    
    units = {
        "Mw": "g/mol", "Mn": "g/mol", "Mw / Mn": "", "Mz": "g/mol", 
        "Mz1": "g/mol", "Mv": "g/mol", "Mp": "g/mol", "IV": "dL/g", 
        "Amount of material": "%", "Bulk CH3 / 1000TC": "", 
        "Bulk SCB / 1000TC": "", "Bulk Comonomer": ""
    }
    df_summary_transposed.insert(0, "unit", df_summary_transposed.index.map(units))
    df_summary_transposed.index.name = "GPC-IR"

    # Compile flattened Horizontal dataframes
    horizontal_raw_list = []
    for item in st.session_state.data_mmd_list:
        temp_df = item["df"].copy()
        temp_df.columns = [f"{item['file_name']}_{col}" for col in temp_df.columns]
        horizontal_raw_list.append(temp_df)
    master_raw_horizontal = pd.concat(horizontal_raw_list, axis=1)

    # --- Header Action Block ---
    col_title, col_download = st.columns([3, 1])
    with col_title:
        st.subheader("📋 GPC-IR Summary Report")
    with col_download:
        # Generate Excel Output with embedded charts directly mapped to coordinates
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_summary_transposed.to_excel(writer, sheet_name='Summary_Report', index=True)
            master_raw_horizontal.to_excel(writer, sheet_name='Raw_Data_MMD', index=False)
            
            # Excel Native Chart Component Setup via xlsxwriter
            workbook  = writer.book
            worksheet_summary = writer.sheets['Summary_Report']
            worksheet_raw = writer.sheets['Raw_Data_MMD']
            
            # Create a Scatter chart type with lines connecting markers natively
            chart_mwd = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
            chart_scb = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
            
            # Loop column matrices to feed data blocks into Excel Chart series
            col_offset = 0
            for idx, item in enumerate(st.session_state.data_mmd_list):
                df_len = len(item["df"])
                
                # Sheet 2 columns tracking: [LogM, MMD, LogM, Cumulative, LogM, SCB...]
                # Plot MWD profile series (Left Y Axis)
                chart_mwd.add_series({
                    'name':       f"{item['file_name']} MWD",
                    'categories': ['Raw_Data_MMD', 1, col_offset, df_len, col_offset], # Preceding LogM Column
                    'values':     ['Raw_Data_MMD', 1, col_offset + 1, df_len, col_offset + 1], # MMD Target Column
                    'line':       {'width': 2.2},
                })
                
                # Plot SCB profile series (Right Secondary Y Axis)
                chart_scb.add_series({
                    'name':       f"{item['file_name']} SCB",
                    'categories': ['Raw_Data_MMD', 1, col_offset + 4, df_len, col_offset + 4], # Preceding LogM Column
                    'values':     ['Raw_Data_MMD', 1, col_offset + 5, df_len, col_offset + 5], # SCB Target Column
                    'y2_axis':    True,
                    'line':       {'width': 1.8, 'dash_type': 'dash_dot'},
                })
                col_offset += len(item["df"].columns)
            
            # Combine profiles into a dual chart stack inside the workbook
            chart_mwd.combine(chart_scb)
            chart_mwd.set_title({'name': 'GPC MWD & SCB Overlay Profile'})
            chart_mwd.set_x_axis({
                'name': 'Log M',
                'min': st.session_state.global_min_logm,
                'max': st.session_state.global_max_logm
            })
            chart_mwd.set_y_axis({'name': 'MMD (Molecular Weight Distribution)'})
            chart_mwd.set_y2_axis({
                'name': 'SCB / 1000TC',
                'min': 0,
                'max': 5.0 if st.session_state.max_scb_value == 0.0 else st.session_state.max_scb_value * 5.0
            })
            chart_mwd.set_size({'width': 850, 'height': 500})
            
            # Insert native chart artifact into Summary sheet location
            worksheet_summary.insert_chart('B18', chart_mwd)
        
        st.download_button(
            label="📥 Download Excel Output",
            data=buffer.getvalue(),
            file_name="GPC_Overlay_Comprehensive_Report.xlsx",
            mime="application/vnd.ms-excel",
            use_container_width=True
        )

    # Dynamic styling and width mapping rules
    col_table, col_spacer = st.columns([4, 1])
    with col_table:
        # Formatting rules to dynamically keep integers clean without floating decimal periods
        formatted_df = df_summary_transposed.style.format(
            formatter=lambda x: f"{int(x)}" if isinstance(x, (int, float)) and x.is_integer() else (f"{x:.2f}" if isinstance(x, (int, float)) else f"{x}"),
            na_rep="-"
        )
        st.dataframe(
            formatted_df,
            use_container_width=False,
            column_config={
                "GPC-IR": st.column_config.Column("GPC-IR", width=220, required=True),
                "unit": st.column_config.Column("unit", width=90)
            }
        )
    st.markdown("---")

# --- Section 2: Dual Y-Axis Clean Overlay Plot ---
if st.session_state.data_mmd_list:
    st.subheader("📈 MWD & SCB Overlay Profile")
    
    fig = go.Figure()
    colors = px.colors.qualitative.Plotly
    
    for i, data_item in enumerate(st.session_state.data_mmd_list):
        f_name = data_item["file_name"]
        df = data_item["df"]
        color = colors[i % len(colors)]
        cols_lower = [c.lower() for c in df.columns]
        
        col_mmd_idx = next((idx for idx, c in enumerate(cols_lower) if 'mmd' in c), None)
        col_scb_idx = next((idx for idx, c in enumerate(cols_lower) if 'scb' in c or '1000tc' in c), None)
        
        if col_mmd_idx is not None and col_mmd_idx > 0:
            col_mwd_logm_idx = col_mmd_idx - 1
            fig.add_trace(go.Scatter(
                x=df.iloc[:, col_mwd_logm_idx], y=df.iloc[:, col_mmd_idx],
                mode='lines', name=f"{f_name} (MWD)",
                line=dict(color=color, width=2.5), yaxis='y1'
            ))
            
        if col_scb_idx is not None and col_scb_idx > 0:
            col_scb_logm_idx = col_scb_idx - 1
            fig.add_trace(go.Scatter(
                x=df.iloc[:, col_scb_logm_idx], y=df.iloc[:, col_scb_idx],
                mode='lines', name=f"{f_name} (SCB)",
                line=dict(color=color, width=2, dash='dashdot'), yaxis='y2'
            ))
    
    scb_upper_limit = 5.0 if st.session_state.max_scb_value == 0.0 else st.session_state.max_scb_value * 5.0
    
    fig.update_layout(
        xaxis=dict(
            title="Log M", showgrid=True, gridcolor='#e2e8f0',
            zeroline=True, zerolinecolor='#cbd5e1',
            range=[st.session_state.global_min_logm, st.session_state.global_max_logm] # Dynamic range margins applied
        ),
        yaxis=dict(title="MMD (Molecular Weight Distribution)", showgrid=True, gridcolor='#e2e8f0', side="left"),
        yaxis2=dict(title="SCB / 1000TC", showgrid=False, anchor="x", overlaying="y", side="right", range=[0, scb_upper_limit]),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        plot_bgcolor='white', paper_bgcolor='white', height=650, margin=dict(l=60, r=60, t=30, b=60)
    )
    
    st.plotly_chart(fig, use_container_width=True)