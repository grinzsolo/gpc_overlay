import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io

st.set_page_config(page_title="GPC Multi-File Overlay Dashboard", layout="wide")

# Custom CSS for a clean, professional, and modern laboratory dashboard style
st.markdown("""
    <style>
        .main { background-color: #f8f9fa; }
        .block-container { padding-top: 1.5rem; padding-bottom: 1.5rem; }
        h1 { color: #0f172a; font-weight: 700; font-size: 2.2rem; margin-bottom: 0.2rem; }
        h3 { color: #334155; font-weight: 600; margin-top: 1rem; }
        
        /* Style for the custom process button inside the form */
        div.stButton > button:first-child {
            background-color: #2563eb; color: white; border-radius: 6px; border: none;
            padding: 0.5rem 2rem; font-weight: 500;
        }
        .stDataFrame { border: 1px solid #e2e8f0; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

st.title("📊 GPC Multi-File Overlay Dashboard")
st.write("Upload GPC files and click Submit to overlay molecular profiles using strict sequential column matching.")
st.markdown("---")

# Wrap the file uploader and submit action inside a clean Streamlit Form
with st.form(key="gpc_upload_form"):
    uploaded_files = st.file_uploader(
        "Select GPC Files (.xls / .xlsx)", 
        type=["xlsx", "xls"], 
        accept_multiple_files=True
    )
    
    # Form submission trigger button
    submit_button = st.form_submit_button(label="🚀 Process and Overlay Data")

# Execution logic starts only when the user explicitly clicks the Submit Button
if submit_button and uploaded_files:
    if len(uploaded_files) > 5:
        st.error("⚠️ Maximum 5 files allowed. Please remove excess files and submit again.")
    else:
        data_mmd_list = []
        results_list = []
        max_scb_value = 0.0 # Variable to track global maximum SCB for dynamic axis scaling
        
        # Loop through each uploaded file to extract information
        for file in uploaded_files:
            file_name = file.name
            
            try:
                excel_file = pd.ExcelFile(file)
                
                # 1. Read Sheet: "Data MMD"
                if "Data MMD" in excel_file.sheet_names:
                    df_mmd = pd.read_excel(file, sheet_name="Data MMD")
                else:
                    df_mmd = pd.read_excel(file, sheet_name=0)
                    
                # 2. Read Sheet: "Results"
                if "Results" in excel_file.sheet_names:
                    df_res = pd.read_excel(file, sheet_name="Results", header=None)
                else:
                    df_res = pd.read_excel(file, sheet_name=1, header=None)
                
                # Standardize column names by dropping empty spaces
                df_mmd.columns = [str(c).strip() for c in df_mmd.columns]
                
                data_mmd_list.append({
                    "file_name": file_name,
                    "df": df_mmd
                })
                
                # Dynamic tracking of the highest SCB value across all uploaded datasets
                cols_lower_temp = [c.lower() for c in df_mmd.columns]
                col_scb_idx_temp = next((idx for idx, c in enumerate(cols_lower_temp) if 'scb' in c or '1000tc' in c), None)
                if col_scb_idx_temp is not None:
                    current_max = pd.to_numeric(df_mmd.iloc[:, col_scb_idx_temp], errors='coerce').max()
                    if pd.notna(current_max) and current_max > max_scb_value:
                        max_scb_value = current_max
                
                # Target metrics list to extract from the "Results" sheet
                res_dict = {"Sample Name": file_name}
                target_metrics = [
                    "Mw", "Mn", "Mw / Mn", "Mz", "Mz1", "Mv", "Mp", "IV", 
                    "Amount of material", "Bulk CH3 / 1000TC", "Bulk SCB / 1000TC", "Bulk Comonomer"
                ]
                
                # Dynamic cell search matching for target metrics and values
                for metric in target_metrics:
                    val = None
                    for row_idx, row in df_res.iterrows():
                        for col_idx, cell_value in enumerate(row):
                            if str(cell_value).strip() == metric:
                                if col_idx + 1 < len(row):
                                    val = row.iloc[col_idx + 1]
                                    # Fallback if the next column is an empty unit separator cell
                                    if pd.isna(val) and col_idx + 2 < len(row):
                                        val = row.iloc[col_idx + 2]
                                break
                        if val is not None:
                            break
                    
                    try:
                        if val is not None and not isinstance(val, str):
                            res_dict[metric] = float(val)
                        else:
                            res_dict[metric] = val
                    except:
                        res_dict[metric] = val
                        
                results_list.append(res_dict)
                
            except Exception as e:
                st.error(f"Error reading file {file_name}: {e}")

        # --- Data Compilation for Tables & Excel Export ---
        if results_list:
            df_summary = pd.DataFrame(results_list)
            df_summary.set_index("Sample Name", inplace=True)
            df_summary_transposed = df_summary.T
            
            # Map standard GPC units
            units = {
                "Mw": "g/mol", "Mn": "g/mol", "Mw / Mn": "", "Mz": "g/mol", 
                "Mz1": "g/mol", "Mv": "g/mol", "Mp": "g/mol", "IV": "dL/g", 
                "Amount of material": "%", "Bulk CH3 / 1000TC": "", 
                "Bulk SCB / 1000TC": "", "Bulk Comonomer": ""
            }
            df_summary_transposed.insert(0, "unit", df_summary_transposed.index.map(units))
            df_summary_transposed.index.name = "GPC-IR"

            # --- Fix NotImplementedError: Flatten MultiIndex for Horizontal Export ---
            horizontal_raw_list = []
            for item in data_mmd_list:
                temp_df = item["df"].copy()
                # Create flat names like "gpc1.xls_LogM" instead of a MultiIndex layout
                temp_df.columns = [f"{item['file_name']}_{col}" for col in temp_df.columns]
                horizontal_raw_list.append(temp_df)
            
            # Join all DataFrames side-by-side cleanly along columns
            master_raw_horizontal = pd.concat(horizontal_raw_list, axis=1)

            # --- Layout Grid: Action Headers ---
            col_title, col_download = st.columns([3, 1])
            with col_title:
                st.subheader("📋 GPC-IR Summary Report")
            with col_download:
                # Multi-Sheet Excel Generation
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_summary_transposed.to_excel(writer, sheet_name='Summary_Report', index=True)
                    # Writing flat columns with index=False works perfectly now
                    master_raw_horizontal.to_excel(writer, sheet_name='Raw_Data_MMD', index=False)
                
                st.download_button(
                    label="📥 Download Excel Output",
                    data=buffer.getvalue(),
                    file_name="GPC_Overlay_Comprehensive_Report.xlsx",
                    mime="application/vnd.ms-excel",
                    use_container_width=True
                )

            # --- Layout Optimization for Table Width ---
            col_table, col_spacer = st.columns([4, 1])
            with col_table:
                st.dataframe(
                    df_summary_transposed.style.format(precision=2, na_rep="-"),
                    use_container_width=False,
                    column_config={
                        "GPC-IR": st.column_config.Column(
                            "GPC-IR",
                            width=220,
                            required=True
                        ),
                        "unit": st.column_config.Column(
                            "unit",
                            width=90
                        )
                    }
                )
            
            st.markdown("---")

        # --- Section 2: Dual Y-Axis Clean Overlay Plot ---
        if data_mmd_list:
            st.subheader("📈 MWD & SCB Overlay Profile")
            
            fig = go.Figure()
            colors = px.colors.qualitative.Plotly
            
            for i, data_item in enumerate(data_mmd_list):
                f_name = data_item["file_name"]
                df = data_item["df"]
                color = colors[i % len(colors)]
                
                # Match columns case-insensitively
                cols_lower = [c.lower() for c in df.columns]
                
                # Find target column indices
                col_mmd_idx = next((idx for idx, c in enumerate(cols_lower) if 'mmd' in c), None)
                col_scb_idx = next((idx for idx, c in enumerate(cols_lower) if 'scb' in c or '1000tc' in c), None)
                
                # 1. Plot MWD profile (Left Y-axis) using Sequential Column Matching (Index - 1)
                if col_mmd_idx is not None and col_mmd_idx > 0:
                    col_mwd_logm_idx = col_mmd_idx - 1
                    
                    fig.add_trace(go.Scatter(
                        x=df.iloc[:, col_mwd_logm_idx],
                        y=df.iloc[:, col_mmd_idx],
                        mode='lines',
                        name=f"{f_name} (MWD)",
                        line=dict(color=color, width=2.5),
                        yaxis='y1'
                    ))
                    
                # 2. Plot SCB profile (Right Y-axis) using Sequential Column Matching (Index - 1)
                if col_scb_idx is not None and col_scb_idx > 0:
                    col_scb_logm_idx = col_scb_idx - 1
                    
                    fig.add_trace(go.Scatter(
                        x=df.iloc[:, col_scb_logm_idx],
                        y=df.iloc[:, col_scb_idx],
                        mode='lines',
                        name=f"{f_name} (SCB)",
                        line=dict(color=color, width=2, dash='dashdot'),
                        yaxis='y2'
                    ))
            
            # --- Dynamic Y2 Axis Range Setting ---
            scb_upper_limit = 5.0 if max_scb_value == 0.0 else max_scb_value * 5.0
            
            # Figure layout configuration for standard clean scientific layout
            fig.update_layout(
                xaxis=dict(
                    title="Log M", 
                    showgrid=True, 
                    gridcolor='#e2e8f0',
                    zeroline=True,
                    zerolinecolor='#cbd5e1'
                ),
                yaxis=dict(
                    title="MMD (Molecular Weight Distribution)",
                    showgrid=True,
                    gridcolor='#e2e8f0',
                    side="left"
                ),
                yaxis2=dict(
                    title="SCB / 1000TC",
                    showgrid=False,
                    anchor="x",
                    overlaying="y",
                    side="right",
                    range=[0, scb_upper_limit]
                ),
                hovermode="x unified",
                legend=dict(
                    orientation="h", 
                    yanchor="bottom", 
                    y=1.02, 
                    xanchor="center", 
                    x=0.5
                ),
                plot_bgcolor='white',
                paper_bgcolor='white',
                height=650,
                margin=dict(l=60, r=60, t=30, b=60)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
elif not uploaded_files:
    st.info("💡 Please upload GPC Excel files inside the box above and click 'Process and Overlay Data'.")