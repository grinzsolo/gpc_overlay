import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import os
import math

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

# Initialize Session States to prevent data from disappearing on download action
if "data_mmd_list" not in st.session_state:
    st.session_state.data_mmd_list = []
if "results_list" not in st.session_state:
    st.session_state.results_list = []
if "global_min_logm" not in st.session_state:
    st.session_state.global_min_logm = 0
if "global_max_logm" not in st.session_state:
    st.session_state.global_max_logm = 0

with st.form(key="gpc_upload_form"):
    uploaded_files = st.file_uploader(
        "Select GPC Files (.xls / .xlsx)", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    submit_button = st.form_submit_button(label="🚀 Process and Overlay Data")

# Process file calculations upon submit trigger
if submit_button and uploaded_files:
    if len(uploaded_files) > 5:
        st.error("⚠️ Maximum 5 files allowed. Please remove excess files and submit again.")
    else:
        st.session_state.data_mmd_list = []
        st.session_state.results_list = []
        
        all_min_logm = []
        all_max_logm = []
        
        for file in uploaded_files:
            file_name_clean = os.path.splitext(file.name)[0]
            
            try:
                excel_file = pd.ExcelFile(file)
                df_mmd = pd.read_excel(file, sheet_name="Data MMD" if "Data MMD" in excel_file.sheet_names else 0)
                df_res = pd.read_excel(file, sheet_name="Results" if "Results" in excel_file.sheet_names else 1, header=None)
                
                df_mmd.columns = [str(c).strip() for c in df_mmd.columns]
                
                st.session_state.data_mmd_list.append({
                    "file_name": file_name_clean,
                    "df": df_mmd
                })
                
                # Global X axis range tracking based on the primary LogM column (Column 0)
                if len(df_mmd.columns) > 0:
                    numeric_logm = pd.to_numeric(df_mmd.iloc[:, 0], errors='coerce').dropna()
                    if not numeric_logm.empty:
                        all_min_logm.append(numeric_logm.min())
                        all_max_logm.append(numeric_logm.max())
                
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
            st.session_state.global_min_logm = int(math.floor(min(all_min_logm) - 2.0))
            st.session_state.global_max_logm = int(math.ceil(max(all_max_logm) + 2.0))

# Render active UI elements from sessions
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

    # --- Header Action Block ---
    col_title, col_download = st.columns([3, 1])
    with col_title:
        st.subheader("📋 GPC-IR Summary Report")
    with col_download:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_summary_transposed.to_excel(writer, sheet_name='Summary_Report', index=True)
            
            workbook  = writer.book
            worksheet_summary = writer.sheets['Summary_Report']
            worksheet_raw = workbook.add_worksheet('Raw_Data_MMD')
            
            # --- Dark Color Theme Styles Configuration ---
            dark_header_format = workbook.add_format({
                'bg_color': '#1E3A8A', 
                'font_color': '#FFFFFF',
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1,
                'border_color': '#94A3B8'
            })
            
            soft_stripe_formats = [
                workbook.add_format({'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#EFF6FF', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#F0FDF4', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#FEF2F2', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#FFFBEB', 'border': 1, 'border_color': '#E2E8F0'})
            ]
            
            # --- Populate Raw MMD Sheet with Multi-Level Dark Headers ---
            current_col_idx = 0
            for file_idx, item in enumerate(st.session_state.data_mmd_list):
                df_item = item["df"]
                num_cols = len(df_item.columns)
                sample_name = item["file_name"]
                
                worksheet_raw.merge_range(
                    0, current_col_idx, 0, current_col_idx + num_cols - 1, 
                    sample_name, dark_header_format
                )
                
                for sub_col_idx, col_name in enumerate(df_item.columns):
                    worksheet_raw.write(1, current_col_idx + sub_col_idx, col_name, dark_header_format)
                
                active_cell_format = soft_stripe_formats[file_idx % len(soft_stripe_formats)]
                for r_idx in range(len(df_item)):
                    for c_idx in range(num_cols):
                        cell_val = df_item.iloc[r_idx, c_idx]
                        if pd.notna(cell_val):
                            worksheet_raw.write_number(r_idx + 2, current_col_idx + c_idx, float(cell_val), active_cell_format)
                        else:
                            worksheet_raw.write(r_idx + 2, current_col_idx + c_idx, "", active_cell_format)
                            
                current_col_idx += num_cols

            # --- Chart Native Overlay Integration ---
            chart_mwd = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
            chart_scb = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
            
            col_offset = 0
            for idx, item in enumerate(st.session_state.data_mmd_list):
                df_len = len(item["df"])
                
                # MWD Profile Series: Column 0 (LogM) vs Column 1 (MMD)
                chart_mwd.add_series({
                    'name':       f"{item['file_name']} (MWD)",
                    'categories': ['Raw_Data_MMD', 2, col_offset, df_len + 1, col_offset],
                    'values':     ['Raw_Data_MMD', 2, col_offset + 1, df_len + 1, col_offset + 1],
                    'line':       {'width': 2.2},
                })
                
                # SCB Profile Series: Column 4 (LogM.2) vs Column 5 (SCB / 1000TC)
                chart_scb.add_series({
                    'name':       f"{item['file_name']} (SCB / 1000TC)",
                    'categories': ['Raw_Data_MMD', 2, col_offset + 4, df_len + 1, col_offset + 4],
                    'values':     ['Raw_Data_MMD', 2, col_offset + 5, df_len + 1, col_offset + 5],
                    'y2_axis':    True, 
                    'line':       {'width': 1.8, 'dash_type': 'dash_dot'},
                })
                col_offset += len(item["df"].columns)
            
            # Master configuration for Native Excel Dual Y-Axis Chart Rendering
            chart_mwd.set_title({'name': 'GPC MWD & SCB Overlay Profile'})
            chart_mwd.set_x_axis({
                'name': 'Log M',
                'min': st.session_state.global_min_logm,
                'max': st.session_state.global_max_logm,
                'major_unit': 1
            })
            
            # Force primary Excel axis to start strictly at 0
            chart_mwd.set_y_axis({
                'name': 'MMD (Molecular Weight Distribution)',
                'min': 0
            })
            
            # Force secondary Excel axis to start strictly at 0 and top at 40
            chart_scb.set_y2_axis({
                'name': 'SCB / 1000TC',
                'min': 0,
                'max': 40,
                'visible': True
            })
            
            chart_mwd.combine(chart_scb)
            chart_mwd.set_size({'width': 850, 'height': 500})
            worksheet_summary.insert_chart('B18', chart_mwd)
        
        st.download_button(
            label="📥 Download Excel Output",
            data=buffer.getvalue(),
            file_name="GPC_Overlay_Comprehensive_Report.xlsx",
            mime="application/vnd.ms-excel",
            use_container_width=True
        )

    # --- Compressed Table Layout Frame ---
    dynamic_ratio = min(max(len(st.session_state.results_list) * 1, 2), 4)
    col_table, col_spacer = st.columns([dynamic_ratio, 5 - dynamic_ratio])
    with col_table:
        formatted_df = df_summary_transposed.style.format(
            formatter=lambda x: f"{int(x)}" if isinstance(x, (int, float)) and x.is_integer() else (f"{x:.2f}" if isinstance(x, (int, float)) else f"{x}"),
            na_rep="-"
        )
        st.dataframe(
            formatted_df,
            use_container_width=False, 
            column_config={
                "GPC-IR": st.column_config.Column("GPC-IR", width=190, required=True),
                "unit": st.column_config.Column("unit", width=75)
            }
        )
    st.markdown("---")

# --- Section 2: Dual Y-Axis Clean Overlay Plot ---
if st.session_state.data_mmd_list:
    st.subheader("📈 MWD & SCB Overlay Profile")
    
    fig = go.Figure()
    colors = px.colors.qualitative.Plotly
    
    # First Loop: Add MWD profile traces for ALL samples first to group them at the top of the legend
    for i, data_item in enumerate(st.session_state.data_mmd_list):
        f_name = data_item["file_name"]
        df = data_item["df"]
        color = colors[i % len(colors)]
        cols_lower = [c.lower() for c in df.columns]
        col_mmd_idx = next((idx for idx, c in enumerate(cols_lower) if 'mmd' in c), None)
        
        if col_mmd_idx is not None and col_mmd_idx > 0:
            fig.add_trace(go.Scatter(
                x=df.iloc[:, col_mmd_idx - 1], y=df.iloc[:, col_mmd_idx],
                mode='lines', name=f"{f_name} (MWD)",
                line=dict(color=color, width=2.5), yaxis='y1'
            ))
            
    # Second Loop: Add SCB / 1000TC profile traces next to push them below MWD in the legend
    for i, data_item in enumerate(st.session_state.data_mmd_list):
        f_name = data_item["file_name"]
        df = data_item["df"]
        color = colors[i % len(colors)]
        cols_lower = [c.lower() for c in df.columns]
        col_scb_idx = next((idx for idx, c in enumerate(cols_lower) if 'scb' in c or '1000tc' in c), None)
        
        if col_scb_idx is not None and col_scb_idx > 0:
            fig.add_trace(go.Scatter(
                x=df.iloc[:, col_scb_idx - 1], y=df.iloc[:, col_scb_idx],
                mode='lines', name=f"{f_name} (SCB / 1000TC)",  
                line=dict(color=color, width=2, dash='dashdot'), yaxis='y2'
            ))
    
    fig.update_layout(
        xaxis=dict(
            title="Log M", showgrid=True, gridcolor='#e2e8f0',
            zeroline=True, zerolinecolor='#cbd5e1',
            range=[st.session_state.global_min_logm, st.session_state.global_max_logm],
            dtick=1
        ),
        # Fix baseline zero alignment perfectly
        yaxis=dict(
            title="MMD (Molecular Weight Distribution)", 
            showgrid=True, 
            gridcolor='#e2e8f0', 
            side="left",
            rangemode="tozero"  # Forces primary axis to strictly anchor 0 line
        ),
        yaxis2=dict(
            title="SCB / 1000TC", 
            showgrid=False, 
            anchor="x", 
            overlaying="y", 
            side="right", 
            range=[0, 40],
            rangemode="tozero"  # Forces secondary axis to strictly anchor 0 line at the exact horizontal plane
        ),
        hovermode="x unified",
        legend=dict(
            orientation="v",       
            yanchor="middle",      
            y=0.5,                 
            xanchor="left",        
            x=1.05                 
        ),
        plot_bgcolor='white', paper_bgcolor='white', height=650, margin=dict(l=60, r=60, t=30, b=60)
    )
    
    st.plotly_chart(fig, use_container_width=True)
            
elif not uploaded_files:
    st.info("💡 Please upload GPC Excel files inside the box above and click 'Process and Overlay Data'.")