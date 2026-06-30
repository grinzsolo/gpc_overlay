import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import os
import math

st.set_page_config(page_title="GPC Multi-File Overlay Dashboard", layout="wide")

# Modern and Professional UI Customization
st.markdown("""
    <style>        
        /* Base page background */
        .main { background-color: #f8fafc; }
        .block-container { padding-top: 2rem; padding-bottom: 2rem; }
        
        /* Modern Typography */
        h1 { color: #0f172a; font-weight: 800; font-size: 2.4rem; margin-bottom: 0.5rem; }
        h2 { color: #1e293b; font-weight: 700; margin-top: 1.5rem; }
        h3 { color: #334155; font-weight: 600; margin-top: 1rem; }
        p { color: #64748b; font-size: 1rem; }
        
        /* Form & Component Styling */
        div[data-testid="stForm"] {
            background-color: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            padding: 1.5rem;
            box-shadow: 0 1px 3px 0 rgb(0 0 0 / 0.05);
        }
        
        /* Custom Button */
        div.stButton > button:first-child {
            background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%);
            color: white;
            border-radius: 8px;
            border: none;
            padding: 0.6rem 2.5rem;
            font-weight: 600;
            transition: all 0.2s ease;
            box-shadow: 0 4px 6px -1px rgb(37 99 235 / 0.2);
        }
        div.stButton > button:first-child:hover {
            background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%);
            box-shadow: 0 6px 12px -1px rgb(37 99 235 / 0.3);
        }
        
        /* Interactive Metric Cards */
        .metric-container {
            background-color: white;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            padding: 1rem;
            text-align: center;
            box-shadow: 0 1px 2px 0 rgb(0 0 0 / 0.05);
        }
        .metric-value { font-size: 1.6rem; font-weight: 700; color: #1e3a8a; }
        .metric-label { font-size: 0.85rem; color: #64748b; font-weight: 500; }
        
        /* Dataframe border smoothness (used by other st.dataframe instances, if any) */
        .stDataFrame { border: 1px solid #e2e8f0; border-radius: 8px; background-color: white; }
    </style>
""", unsafe_allow_html=True)

# Main Dashboard Header Block
st.title("GPC Multi-File Overlay Dashboard")
st.write("Upload Multiple GPC files to automatically compile molecular weight profiles and export publication-ready Excel reports.")
st.markdown("---")

# Initialize Session States
if "data_mmd_list" not in st.session_state:
    st.session_state.data_mmd_list = []
if "results_list" not in st.session_state:
    st.session_state.results_list = []
if "global_min_logm" not in st.session_state:
    st.session_state.global_min_logm = 0
if "global_max_logm" not in st.session_state:
    st.session_state.global_max_logm = 0

# --- Sidebar Controls for Uploading ---
with st.sidebar:
    st.header("⚙️ Control Panel")
    st.write("Configure and upload raw analytical documents.")
    
    with st.form(key="gpc_upload_form"):
        uploaded_files = st.file_uploader(
            "Select GPC Files (.xls / .xlsx)", 
            type=["xlsx", "xls"], 
            accept_multiple_files=True
        )
        submit_button = st.form_submit_button(label="🚀 Process & Overlay Data")

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

    # --- Analytics Highlight Summary Cards ---
    st.subheader("💡 Key Sample Highlights")
    m_col1, m_col2, m_col3, m_col4 = st.columns(4)
    with m_col1:
        avg_mw = df_summary["Mw"].mean() if "Mw" in df_summary.columns else 0
        st.markdown(f'<div class="metric-container"><div class="metric-value">{int(round(avg_mw)):,}</div><div class="metric-label">Average Mw (g/mol)</div></div>', unsafe_allow_html=True)
    with m_col2:
        avg_mn = df_summary["Mn"].mean() if "Mn" in df_summary.columns else 0
        st.markdown(f'<div class="metric-container"><div class="metric-value">{int(round(avg_mn)):,}</div><div class="metric-label">Average Mn (g/mol)</div></div>', unsafe_allow_html=True)
    with m_col3:
        avg_pdi = df_summary["Mw / Mn"].mean() if "Mw / Mn" in df_summary.columns else 0
        st.markdown(f'<div class="metric-container"><div class="metric-value">{avg_pdi:.2f}</div><div class="metric-label">Average PDI (Mw/Mn)</div></div>', unsafe_allow_html=True)
    with m_col4:
        total_samples = len(df_summary)
        st.markdown(f'<div class="metric-container"><div class="metric-value">{total_samples}</div><div class="metric-label">Total Samples Analysed</div></div>', unsafe_allow_html=True)
    
    st.write("")

    # --- Header Action Block ---
    col_title, col_download = st.columns([3, 1])
    with col_title:
        st.subheader("📋 Compiled GPC-IR Summary Report")
    with col_download:
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_summary_transposed.to_excel(writer, sheet_name='Summary_Report', index=True)
            
            workbook  = writer.book
            worksheet_summary = writer.sheets['Summary_Report']
            worksheet_raw = workbook.add_worksheet('Raw_Data_MMD')
            
            # Formats definitions for Excel
            summary_header_format = workbook.add_format({
                'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'bold': True,
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#94A3B8'
            })
            summary_index_format = workbook.add_format({
                'bg_color': '#F1F5F9', 'font_color': '#0F172A', 'bold': True,
                'align': 'left', 'valign': 'vcenter', 'border': 1, 'border_color': '#CBD5E1'
            })
            summary_data_odd = workbook.add_format({
                'bg_color': '#FFFFFF', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0'
            })
            summary_data_even = workbook.add_format({
                'bg_color': '#F8FAFC', 'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#E2E8F0'
            })
            
            dark_header_format = workbook.add_format({
                'bg_color': '#1E3A8A', 'font_color': '#FFFFFF', 'bold': True,
                'align': 'center', 'valign': 'vcenter', 'border': 1, 'border_color': '#94A3B8'
            })
            soft_stripe_formats = [
                workbook.add_format({'bg_color': '#F8FAFC', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#EFF6FF', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#F0FDF4', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#FEF2F2', 'border': 1, 'border_color': '#E2E8F0'}),
                workbook.add_format({'bg_color': '#FFFBEB', 'border': 1, 'border_color': '#E2E8F0'})
            ]

            # Populate Summary sheet
            worksheet_summary.write(0, 0, "GPC-IR", summary_header_format)
            for col_num, col_name in enumerate(df_summary_transposed.columns):
                worksheet_summary.write(0, col_num + 1, col_name, summary_header_format)

            for row_num, (index_val, row_data) in enumerate(df_summary_transposed.iterrows()):
                fmt = summary_data_even if row_num % 2 == 0 else summary_data_odd
                worksheet_summary.write(row_num + 1, 0, index_val, summary_index_format)
                
                for col_num, cell_val in enumerate(row_data):
                    if index_val in ["Mw", "Mn", "Mz", "Mz1", "Mv", "Mp"] and pd.notna(cell_val) and col_num > 0:
                        worksheet_summary.write_number(row_num + 1, col_num + 1, int(cell_val), fmt)
                    elif pd.notna(cell_val) and isinstance(cell_val, (int, float)):
                        worksheet_summary.write_number(row_num + 1, col_num + 1, float(cell_val), fmt)
                    else:
                        val_to_write = "" if pd.isna(cell_val) else str(cell_val)
                        worksheet_summary.write(row_num + 1, col_num + 1, val_to_write, fmt)

            # Excel Column Width Setting
            worksheet_summary.set_column(0, 0, 24)
            worksheet_summary.set_column(1, 1, 12)
            for col_num, col_name in enumerate(df_summary_transposed.columns):
                if col_name != "unit":
                    max_header_len = len(str(col_name)) + 6
                    worksheet_summary.set_column(col_num + 1, col_num + 1, max(max_header_len, 22))

            # Populate Raw MMD sheet
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

            current_col_idx = 0
            for item in st.session_state.data_mmd_list:
                df_item = item["df"]
                for sub_col_idx, col_name in enumerate(df_item.columns):
                    max_len = max(df_item.iloc[:, sub_col_idx].astype(str).str.len().max(), len(str(col_name))) + 4
                    worksheet_raw.set_column(current_col_idx + sub_col_idx, current_col_idx + sub_col_idx, max(max_len, 14))
                current_col_idx += len(df_item.columns)

            # Native Excel Chart integration
            chart_master = workbook.add_chart({'type': 'scatter', 'subtype': 'straight'})
            col_offset = 0
            for idx, item in enumerate(st.session_state.data_mmd_list):
                df_len = len(item["df"])
                chart_master.add_series({
                    'name':       f"{item['file_name']} (MWD)",
                    'categories': ['Raw_Data_MMD', 2, col_offset, df_len + 1, col_offset],
                    'values':     ['Raw_Data_MMD', 2, col_offset + 1, df_len + 1, col_offset + 1],
                    'line':       {'width': 2.2},
                })
                chart_master.add_series({
                    'name':       f"{item['file_name']} (SCB / 1000TC)",
                    'categories': ['Raw_Data_MMD', 2, col_offset + 4, df_len + 1, col_offset + 4],
                    'values':     ['Raw_Data_MMD', 2, col_offset + 5, df_len + 1, col_offset + 5],
                    'y2_axis':     1,  
                    'line':       {'width': 1.8, 'dash_type': 'dash_dot'},
                })
                col_offset += len(item["df"].columns)
            
            chart_master.set_title({'name': 'GPC MWD & SCB/1000TC Overlay Profile'})
            chart_master.set_x_axis({'name': 'Log M', 'min': st.session_state.global_min_logm, 'max': st.session_state.global_max_logm, 'major_unit': 1})
            chart_master.set_y_axis({'name': 'MMD (Molecular Weight Distribution)', 'min': 0})
            chart_master.set_y2_axis({'name': 'SCB / 1000TC', 'min': 0, 'max': 40, 'visible': True})
            chart_master.set_size({'width': 900, 'height': 520})
            worksheet_summary.insert_chart('B18', chart_master)
        
        st.download_button(
            label="📥 Download Comprehensive Excel Report",
            data=buffer.getvalue(),
            file_name="GPC_Overlay_Comprehensive_Report.xlsx",
            mime="application/vnd.ms-excel",
            use_container_width=True
        )

    # --- 🛠️ ตั้งค่าตาราง Web App: 
    #     - unit column แคบพอดีตัวอักษร
    #     - sample column กว้างพอดีกับตัวเลขที่ยาวที่สุดในคอลัมน์นั้น
    #     - ชื่อ sample ที่ยาว wrap ลงบรรทัดใหม่ได้จริง
    #     - ไม่มี row ว่าง (null) ปนอยู่
    #
    #     หมายเหตุ: st.dataframe เรนเดอร์ตารางด้วย canvas (glide-data-grid)
    #     ดังนั้น CSS ที่เขียนไว้ก่อนหน้านี้จะไม่มีผลกับ header/cell จริง ๆ
    #     (ข้อความเลยถูกตัด/จมแทนที่จะ wrap) จึงเปลี่ยนมาใช้ HTML table ปกติ
    #     เพื่อให้ควบคุมความกว้างคอลัมน์และการขึ้นบรรทัดใหม่ได้แม่นยำ 100%

    import html as _html

    def format_cell_for_width(val, index_val):
        """ช่วยประเมินความยาวของค่าตามรูปแบบที่จะแสดงผลจริง เพื่อคำนวณความกว้างคอลัมน์"""
        if pd.isna(val):
            return "-"
        if isinstance(val, (int, float)):
            if float(val).is_integer():
                return f"{int(val):,}"
            return f"{val:.2f}"
        return str(val)

    # ตัดแถวที่เป็น null ทั้งแถวออก (ไม่นับคอลัมน์ unit) ก่อนแสดงผล
    data_only_cols = [c for c in df_summary_transposed.columns if c != "unit"]
    df_display = df_summary_transposed[df_summary_transposed[data_only_cols].notna().any(axis=1)]

    # คำนวณความกว้าง unit column ให้พอดีกับข้อความ unit ที่ยาวที่สุด
    unit_values = df_display["unit"].fillna("").astype(str)
    max_unit_len = max([len(v) for v in unit_values] + [len("unit")])
    unit_col_width = max(55, min(90, max_unit_len * 8 + 24))

    # คำนวณความกว้าง GPC-IR (index) column ให้พอดีกับชื่อ metric ที่ยาวที่สุด
    max_index_len = max([len(str(v)) for v in df_display.index] + [len("GPC-IR")])
    index_col_width = max(120, min(220, max_index_len * 8 + 28))

    # คำนวณความกว้างแต่ละ sample column ให้พอดีตัวเลขที่ยาวที่สุด แต่ยอมให้ชื่อ sample wrap ได้
    sample_col_widths = {}
    for sample_col in data_only_cols:
        formatted_vals = [
            format_cell_for_width(val, idx)
            for idx, val in zip(df_display.index, df_display[sample_col])
        ]
        max_val_len = max([len(v) for v in formatted_vals] + [1])
        data_width = max_val_len * 9 + 28

        # ความกว้างตามคำที่ยาวที่สุดในชื่อ sample (ส่วนที่เหลือจะ wrap ลงบรรทัดใหม่)
        header_words = str(sample_col).replace("_", " ").replace("-", " ").split(" ")
        max_word_len = max([len(w) for w in header_words] + [1])
        header_width = max_word_len * 7.5 + 28

        sample_col_widths[sample_col] = int(max(data_width, header_width, 100))

    # --- สร้างตารางด้วย HTML จริง (ไม่ใช้ canvas) เพื่อให้ wrap ข้อความและกำหนดความกว้างได้แม่นยำ ---
    colgroup_html = f'<col style="width:{index_col_width}px;"><col style="width:{unit_col_width}px;">'
    header_cells_html = '<th class="gpc-th gpc-th-index">GPC-IR</th><th class="gpc-th gpc-th-unit">unit</th>'
    for sample_col in data_only_cols:
        w = sample_col_widths[sample_col]
        colgroup_html += f'<col style="width:{w}px;">'
        header_cells_html += f'<th class="gpc-th gpc-th-sample">{_html.escape(str(sample_col))}</th>'

    body_rows_html = ""
    for row_i, (index_val, row_data) in enumerate(df_display.iterrows()):
        row_class = "gpc-row-even" if row_i % 2 == 0 else "gpc-row-odd"
        unit_val = row_data["unit"]
        unit_display = _html.escape(str(unit_val)) if pd.notna(unit_val) and str(unit_val) != "" else ""
        cells = f'<td class="gpc-td gpc-td-index">{_html.escape(str(index_val))}</td>'
        cells += f'<td class="gpc-td gpc-td-unit">{unit_display}</td>'
        for sample_col in data_only_cols:
            val = row_data[sample_col]
            val_display = format_cell_for_width(val, index_val) if pd.notna(val) else "-"
            cells += f'<td class="gpc-td gpc-td-data">{_html.escape(val_display)}</td>'
        body_rows_html += f'<tr class="{row_class}">{cells}</tr>'

    gpc_table_html = f"""
    <style>
        .gpc-table-wrapper {{
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            overflow: hidden;
            background-color: white;
        }}
        table.gpc-table {{
            border-collapse: collapse;
            table-layout: fixed;
            font-family: "Source Sans Pro", sans-serif;
            font-size: 0.85rem;
        }}
        table.gpc-table th.gpc-th {{
            background-color: #1E3A8A;
            color: #ffffff;
            font-weight: 700;
            text-align: left;
            padding: 8px 10px;
            border: 1px solid #94A3B8;
            white-space: normal;
            word-wrap: break-word;
            line-height: 1.25;
            vertical-align: middle;
        }}
        table.gpc-table th.gpc-th-sample {{
            text-align: center;
        }}
        table.gpc-table td.gpc-td {{
            padding: 7px 10px;
            border: 1px solid #e2e8f0;
            white-space: normal;
            word-wrap: break-word;
        }}
        table.gpc-table td.gpc-td-index {{
            color: #334155;
            font-weight: 600;
            text-align: left;
        }}
        table.gpc-table td.gpc-td-unit {{
            color: #64748b;
            text-align: left;
        }}
        table.gpc-table td.gpc-td-data {{
            text-align: right;
            color: #0f172a;
            font-variant-numeric: tabular-nums;
        }}
        table.gpc-table tr.gpc-row-even {{ background-color: #ffffff; }}
        table.gpc-table tr.gpc-row-odd {{ background-color: #f8fafc; }}
    </style>
    <div class="gpc-table-wrapper">
        <table class="gpc-table">
            <colgroup>{colgroup_html}</colgroup>
            <thead><tr>{header_cells_html}</tr></thead>
            <tbody>{body_rows_html}</tbody>
        </table>
    </div>
    """

    st.markdown(gpc_table_html, unsafe_allow_html=True)
    st.markdown("---")

# --- Section 2: Dual Y-Axis Clean Overlay Plot ---
if st.session_state.data_mmd_list:
    st.subheader("📈 Interactive Distribution Profile (MWD & SCB Overlay)")
    
    fig = go.Figure()
    colors = px.colors.qualitative.Safe
    
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
            
    for i, data_item in enumerate(st.session_state.data_mmd_list):
        f_name = data_item["file_name"]
        df = data_item["df"]
        color = colors[i % len(colors)]
        cols_lower = [c.lower() for c in df.columns]
        col_scb_idx = next((idx for idx, c in enumerate(cols_lower) if 'scb' in c or '1000tc' in c), None)
        
        if col_scb_idx is not None and col_scb_idx > 0:
            fig.add_trace(go.Scatter(
                x=df.iloc[:, col_scb_idx - 1], y=df.iloc[:, col_scb_idx],
                mode='lines', name=f"{f_name} (SCB)",  
                line=dict(color=color, width=2, dash='dashdot'), yaxis='y2'
            ))
    
    fig.update_layout(
        xaxis=dict(
            title="Log M", showgrid=True, gridcolor='#e2e8f0',
            zeroline=True, zerolinecolor='#cbd5e1',
            showline=True, linewidth=1, linecolor='#cbd5e1', mirror=True,
            range=[st.session_state.global_min_logm, st.session_state.global_max_logm],
            dtick=1
        ),
        yaxis=dict(
            title="MMD (Molecular Weight Distribution)", 
            showgrid=True, gridcolor='#e2e8f0', side="left",
            showline=True, linewidth=1, linecolor='#cbd5e1', mirror=True,
            rangemode="tozero"
        ),
        yaxis2=dict(
            title="SCB / 1000TC", 
            showgrid=False, anchor="x", overlaying="y", side="right", 
            showline=True, linewidth=1, linecolor='#cbd5e1', mirror=True,
            range=[0, 40],
            rangemode="tozero"
        ),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=-0.25, xanchor="center", x=0.5),
        plot_bgcolor='white', paper_bgcolor='white', height=600, margin=dict(l=60, r=60, t=20, b=80)
    )
    
    st.plotly_chart(fig, use_container_width=True)
            
elif not uploaded_files:
    st.info("💡 Please upload GPC files in the sidebar and click 'Process & Overlay Data' to view the dashboard.")
