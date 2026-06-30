import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px

st.set_page_config(page_title="GPC Multi-File Overlay Dashboard", layout="wide")

st.title("📊 GPC Multi-File Overlay Dashboard")
st.write("Upload GPC files (Up to 5 files) to overlay MWD, SCB plots and compare the summary results table.")

# File uploader section (Restricted to maximum 5 files)
uploaded_files = st.file_uploader(
    "Upload GPC Files (.xls / .xlsx)", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

if uploaded_files:
    if len(uploaded_files) > 5:
        st.error("⚠️ Maximum 5 files allowed. Please remove excess files.")
    else:
        data_mmd_list = []
        results_list = []
        
        # Loop through each uploaded file
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
                
                # Clean up column whitespace names
                df_mmd.columns = [str(c).strip() for c in df_mmd.columns]
                
                data_mmd_list.append({
                    "file_name": file_name,
                    "df": df_mmd
                })
                
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

        # --- Section 1: GPC Summary Report Table ---
        if results_list:
            st.subheader("📋 GPC-IR Summary Report")
            
            df_summary = pd.DataFrame(results_list)
            df_summary.set_index("Sample Name", inplace=True)
            df_summary_transposed = df_summary.T
            
            # Map standard GPC units based on the image format reference (image_029ebd.png)
            units = {
                "Mw": "g/mol", "Mn": "g/mol", "Mw / Mn": "", "Mz": "g/mol", 
                "Mz1": "g/mol", "Mv": "g/mol", "Mp": "g/mol", "IV": "dL/g", 
                "Amount of material": "%", "Bulk CH3 / 1000TC": "", 
                "Bulk SCB / 1000TC": "", "Bulk Comonomer": ""
            }
            df_summary_transposed.insert(0, "unit", df_summary_transposed.index.map(units))
            df_summary_transposed.index.name = "GPC-IR"
            
            st.dataframe(df_summary_transposed, use_container_width=True)

        # --- Section 2: Dual Y-Axis Overlay Plot ---
        if data_mmd_list:
            st.subheader("📈 MWD & SCB Overlay Profile")
            
            fig = go.Figure()
            colors = px.colors.qualitative.Plotly 
            
            for i, data_item in enumerate(data_mmd_list):
                f_name = data_item["file_name"]
                df = data_item["df"]
                color = colors[i % len(colors)]
                
                # Case-insensitive column search to handle any dynamic naming formats
                col_logm = [c for c in df.columns if 'logm' in c.lower()]
                col_mmd = [c for c in df.columns if 'mmd' in c.lower()]
                col_scb = [c for c in df.columns if 'scb' in c.lower() or '1000tc' in c.lower()]
                
                if col_logm and col_mmd:
                    # 1. Plot MWD profile (Left Y-axis - Solid line)
                    fig.add_trace(go.Scatter(
                        x=df[col_logm[0]],
                        y=df[col_mmd[0]],
                        mode='lines',
                        name=f"{f_name}",
                        line=dict(color=color, width=2.5),
                        yaxis='y1'
                    ))
                    
                if col_logm and col_scb:
                    # 2. Plot SCB profile (Right Y-axis - Dash-dot line)
                    fig.add_trace(go.Scatter(
                        x=df[col_logm[0]],
                        y=df[col_scb[0]],
                        mode='lines',
                        name=f"{f_name} SCB/1000TC",
                        line=dict(color=color, width=2, dash='dashdot'),
                        yaxis='y2'
                    ))
            
            # Figure layout configuration for strict standard clean grid layout
            fig.update_layout(
                xaxis=dict(
                    title="LogM", 
                    showgrid=True, 
                    gridcolor='lightgray',
                    zeroline=True,
                    zerolinecolor='gray'
                ),
                yaxis=dict(
                    title="MMD",
                    showgrid=True,
                    gridcolor='lightgray',
                    side="left"
                ),
                yaxis2=dict(
                    title="SCB / 1000TC",
                    showgrid=False,
                    anchor="x",
                    overlaying="y",
                    side="right"
                ),
                hovermode="x unified",
                legend=dict(
                    orientation="v", 
                    yanchor="top", 
                    y=0.5, 
                    xanchor="left", 
                    x=1.08
                ),
                plot_bgcolor='white',
                height=600,
                margin=dict(l=50, r=50, t=20, b=50)
            )
            
            st.plotly_chart(fig, use_container_width=True)
else:
    st.info("💡 Please upload GPC Excel files to generate the dashboard profile.")