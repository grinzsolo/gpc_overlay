import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import os
import math

# Configuration
st.set_page_config(page_title="GPC Laboratory Dashboard", layout="wide", page_icon="📊")

# --- Custom CSS for Laboratory Look ---
st.markdown("""
    <style>
        .main { background-color: #f1f5f9; }
        .stMetric { background-color: #ffffff; padding: 15px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        h1, h2, h3 { color: #1e293b; }
        .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; }
        div.stDeployButton { display: none; } /* Hide streamlit watermark */
    </style>
""", unsafe_allow_html=True)

# Header Section
st.title("📊 GPC Analysis Dashboard")
st.markdown("Advanced molecular profile overlay and automated reporting tool for lab workflows.")
st.markdown("---")

# ... [คงส่วนการ Initialize Session State และ Logic การอ่านไฟล์ไว้เหมือนเดิม] ...

# Logic ส่วนการแสดงผล
if st.session_state.results_list:
    # 1. Dashboard Metrics Summary (Highlights)
    st.subheader("💡 Key Sample Insights")
    cols = st.columns(min(len(st.session_state.results_list), 4))
    for i, res in enumerate(st.session_state.results_list[:4]):
        cols[i].metric(label=res['Sample Name'], value=f"{res.get('Mw', 0):,}", delta="Mw (g/mol)")

    # 2. Results Table with better formatting
    st.subheader("📋 Summary Report")
    df_summary = pd.DataFrame(st.session_state.results_list).set_index("Sample Name").T
    
    # Render interactive table
    st.dataframe(df_summary.style.highlight_max(axis=1), use_container_width=True)

    # 3. Action Bar
    col1, col2 = st.columns([1, 4])
    with col1:
        # ย้าย Download button มาไว้ในตำแหน่งที่เด่นชัด
        st.download_button(
            label="📥 Export Report",
            data=buffer.getvalue(),
            file_name="GPC_Analysis_Report.xlsx",
            mime="application/vnd.ms-excel",
        )

    st.markdown("---")
    
    # 4. Chart Section
    st.subheader("📈 MWD & SCB/1000TC Analysis")
    # [ใส่โค้ด Plotly ของคุณที่นี่... ปรับ height=500 ให้สมดุล]
    st.plotly_chart(fig, use_container_width=True)

else:
    # Empty State with better call to action
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.info("### Start by uploading your GPC files")
        st.write("Upload `.xlsx` files using the sidebar or form above to generate your overlay analysis.")