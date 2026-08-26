import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io

# ReportLab imports for PDF Generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ---------------------------------------------------------
# GLOBAL CONFIGURATION
# Set your preferred currency symbol here (e.g., "FCFA", "$", "€", "MAD")
# ---------------------------------------------------------
CURRENCY_SYMBOL = "FCFA"

st.set_page_config(page_title="Professional Sales & Comparison Dashboard", layout="wide")
st.title("📊 Multi-Period Commercial Performance & Variance Dashboard")
st.markdown("---")

# 1. Multi-File Uploader
uploaded_files = st.file_uploader(
    "Upload three or more Excel files to perform 3-month comparative analysis", 
    type=["xlsx", "xls"], 
    accept_multiple_files=True
)

def format_currency(val):
    if CURRENCY_SYMBOL in ["$", "€", "£"]:
        return f"{CURRENCY_SYMBOL}{val:,.2f}"
    return f"{val:,.0f} {CURRENCY_SYMBOL}"

if uploaded_files:
    all_dfs = []
    
    for uploaded_file in uploaded_files:
        try:
            df_temp = pd.read_excel(uploaded_file)
            df_temp.columns = [col.strip() for col in df_temp.columns]
            
            # Dynamic column detection
            date_col = [col for col in df_temp.columns if 'date' in col.lower()]
            comm_col = [col for col in df_temp.columns if 'commercial' in col.lower() or 'agent' in col.lower()]
            net_col = [col for col in df_temp.columns if 'net' in col.lower() or 'payer' in col.lower() or 'total' in col.lower()]
            
            if date_col and comm_col and net_col:
                df_temp = df_temp.rename(columns={comm_col[0]: 'Commercial', net_col[0]: 'Net_a_Payer', date_col[0]: 'Date'})
                df_temp['Date'] = pd.to_datetime(df_temp['Date'])
                df_temp['Net_a_Payer'] = pd.to_numeric(df_temp['Net_a_Payer'], errors='coerce').fillna(0)
                df_temp['Source_File'] = uploaded_file.name
                all_dfs.append(df_temp)
            else:
                st.warning(f"⚠️ Skipped '{uploaded_file.name}': Missing required columns ('Date', 'Commercial', 'Net à Payer').")
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")

    if all_dfs:
        # Merge all uploaded excel sheets
        df = pd.concat(all_dfs, ignore_index=True)
        df['YearMonth'] = df['Date'].dt.to_period('M')
        
        # Determine Periods dynamically
        unique_periods = sorted(df['YearMonth'].unique())
        
        # --- Section 1: Executive Overview KPIs ---
        st.subheader("📌 Executive Performance Summary")
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        total_revenue = df['Net_a_Payer'].sum()
        total_proformas = len(df)
        total_commercials = df['Commercial'].nunique()
        
        kpi1.metric("Total Overall Revenue", format_currency(total_revenue))
        kpi2.metric("Total Proformas Issued", f"{total_proformas:,}")
        kpi3.metric("Active Commercials", f"{total_commercials}")
        
        # Monthly Growth Metric
        monthly_totals = df.groupby('YearMonth')['Net_a_Payer'].sum().sort_index(ascending=False).reset_index()
        
        if len(monthly_totals) >= 2:
            latest_month_val = monthly_totals.iloc[0]['Net_a_Payer']
            prev_month_val = monthly_totals.iloc[1]['Net_a_Payer']
            growth = ((latest_month_val - prev_month_val) / prev_month_val) * 100 if prev_month_val != 0 else 0
            
            kpi4.metric(
                label=f"Latest Month Trend ({monthly_totals.iloc[0]['YearMonth']})", 
                value=f"{growth:+.1f}%", 
                delta=format_currency(latest_month_val - prev_month_val)
            )
        else:
            kpi4.metric("Latest Month Trend", "N/A", delta="Requires data from 2+ months")

        st.markdown("---")

        # --- Section 2: 3-Month Commercial Progression Analysis ---
        st.subheader("👩‍💼 3-Month Revenue Progression Analysis")
        
        if len(unique_periods) >= 3:
            periods_3 = unique_periods[-3:]
            p1, p2, p3 = periods_3[0], periods_3[1], periods_3[2]
            
            st.info(f"💡 3-Month Progression evaluated across periods: **{p1}** ➔ **{p2}** ➔ **{p3}**")
            
            # Pivot table to get revenue for each month per commercial
            pivot_rev = df.pivot_table(index='Commercial', columns='YearMonth', values='Net_a_Payer', aggfunc='sum', fill_value=0)
            
            # 3-MONTH CALCULATION LOGIC: Evaluates total change across ALL 3 MONTHS (p3 vs p1)
            def eval_3month_progression(row):
                m1 = row.get(p1, 0)
                m2 = row.get(p2, 0)
                m3 = row.get(p3, 0)
                
                total_3m_growth = m3 - m1
                
                if total_3m_growth > 0 or (m3 > m2 and m3 >= m1):
                    return "(+) Advancing"
                else:
                    return "(-) Declining"

            pivot_rev['Status'] = pivot_rev.apply(eval_3month_progression, axis=1)
            
            # Build Revenue Table (No 'Latest Difference' column)
            display_rev_df = pd.DataFrame({'Status': pivot_rev['Status']})
            display_rev_df[f"Rev. {p3} (Latest)"] = pivot_rev[p3].apply(format_currency)
            display_rev_df[f"Rev. {p2}"] = pivot_rev[p2].apply(format_currency)
            display_rev_df[f"Rev. {p1}"] = pivot_rev[p1].apply(format_currency)
            
            display_rev_df = display_rev_df.reset_index().sort_values(by=f"Rev. {p3} (Latest)", ascending=False)
            
            st.dataframe(
                display_rev_df[['Commercial', 'Status', f"Rev. {p3} (Latest)", f"Rev. {p2}", f"Rev. {p1}"]], 
                use_container_width=True
            )
            
            # --- Section 3: NEW TABLE - Total Proformas Count per Month ---
            st.markdown("---")
            st.subheader("📑 Total Number of Proformas Issued per Commercial (3-Month Breakdown)")
            
            pivot_prof = df.pivot_table(index='Commercial', columns='YearMonth', values='Net_a_Payer', aggfunc='count', fill_value=0)
            
            display_prof_df = pd.DataFrame(index=pivot_prof.index)
            display_prof_df[f"Proformas {p3} (Latest)"] = pivot_prof[p3]
            display_prof_df[f"Proformas {p2}"] = pivot_prof[p2]
            display_prof_df[f"Proformas {p1}"] = pivot_prof[p1]
            display_prof_df["Total 3-Month Proformas"] = pivot_prof[p3] + pivot_prof[p2] + pivot_prof[p1]
            
            display_prof_df = display_prof_df.reset_index().sort_values(by="Total 3-Month Proformas", ascending=False)
            
            st.dataframe(display_prof_df, use_container_width=True)
            
            # Side-by-Side 3-Month Bar Chart
            st.markdown("---")
            fig_comp = go.Figure()
            for p in periods_3:
                fig_comp.add_trace(go.Bar(
                    x=pivot_rev.index, 
                    y=pivot_rev[p], 
                    name=str(p)
                ))
            fig_comp.update_layout(
                barmode='group', 
                title="3-Month Side-by-Side Revenue Comparison", 
                xaxis_title="Commercial", 
                yaxis_title=f"Net à Payer ({CURRENCY_SYMBOL})"
            )
            st.plotly_chart(fig_comp, use_container_width=True)
            
        else:
            st.warning(f"⚠️ Please upload data spanning across at least 3 distinct months. Currently detected: {len(unique_periods)} month(s).")

        st.markdown("---")

        # --- Section 4: Monthly Evolution & Variance ---
        st.subheader("📅 Combined Monthly Revenue Evolution")
        
        monthly_df = df.groupby('YearMonth').agg(
            Total_Revenue=('Net_a_Payer', 'sum'),
            Total_Proformas=('Net_a_Payer', 'count')
        ).reset_index().sort_values('YearMonth')
        
        monthly_df['Prev_Revenue'] = monthly_df['Total_Revenue'].shift(1).fillna(0)
        monthly_df['Variance_Val'] = monthly_df['Total_Revenue'] - monthly_df['Prev_Revenue']
        monthly_df['Growth_%'] = ((monthly_df['Total_Revenue'] - monthly_df['Prev_Revenue']) / monthly_df['Prev_Revenue']) * 100
        monthly_df['Growth_%'] = monthly_df['Growth_%'].fillna(0)
        
        col_m1, col_m2 = st.columns([1, 1])
        
        with col_m1:
            display_monthly = monthly_df.copy()
            display_monthly['YearMonth'] = display_monthly['YearMonth'].astype(str)
            display_monthly['Total Revenue'] = display_monthly['Total_Revenue'].apply(format_currency)
            display_monthly['Variance'] = display_monthly['Variance_Val'].apply(format_currency)
            display_monthly['Growth Rate'] = display_monthly['Growth_%'].apply(lambda x: f"{x:+.1f}%" if x != 0 else "0.0%")
            
            st.dataframe(
                display_monthly[['YearMonth', 'Total Revenue', 'Total_Proformas', 'Variance', 'Growth Rate']],
                use_container_width=True
            )
            
        with col_m2:
            monthly_chart = monthly_df.copy()
            monthly_chart['YearMonth'] = monthly_chart['YearMonth'].astype(str)
            fig_line = px.line(
                monthly_chart, 
                x='YearMonth', 
                y='Total_Revenue', 
                markers=True, 
                title="Monthly Revenue Trend",
                labels={'Total_Revenue': 'Revenue', 'YearMonth': 'Month'}
            )
            st.plotly_chart(fig_line, use_container_width=True)

        st.markdown("---")

        # --- Section 5: PDF Export Logic ---
        def generate_comparison_pdf(overall_rev, overall_prof, rev_table, prof_table, monthly_table):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1E3A8A'), spaceAfter=15)
            heading_style = ParagraphStyle('HeadStyle', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0F172A'), spaceBefore=12, spaceAfter=8)
            
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10)
            cell_header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.whitesmoke)

            story.append(Paragraph("Commercial 3-Month Performance & Proforma Report", title_style))
            story.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
            story.append(Spacer(1, 10))
            
            # Executive Summary
            exec_data = [
                [Paragraph("<b>Metric</b>", cell_style), Paragraph("<b>Value</b>", cell_style)],
                ["Total Consolidated Revenue", format_currency(overall_rev)],
                ["Total Proformas Generated", f"{overall_prof:,}"]
            ]
            t_exec = Table(exec_data, colWidths=[200, 350])
            t_exec.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F1F5F9')),
                ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ]))
            story.append(Paragraph("Executive Summary", heading_style))
            story.append(t_exec)
            story.append(Spacer(1, 12))
            
            # 3-Month Commercial Revenue Progression
            if not rev_table.empty:
                story.append(Paragraph("3-Month Revenue Progression Analysis", heading_style))
                cols = list(rev_table.columns)
                c_data = [[Paragraph(f"<b>{col}</b>", cell_header_style) for col in cols]]
                
                for _, r in rev_table.iterrows():
                    row_cells = []
                    for col in cols:
                        val = str(r[col])
                        row_cells.append(Paragraph(val, cell_style))
                    c_data.append(row_cells)
                
                t_comm = Table(c_data, colWidths=[150, 85, 110, 110, 110])
                t_comm.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]))
                story.append(t_comm)
                story.append(Spacer(1, 12))

            # 3-Month Proforma Count Breakdown Table
            if not prof_table.empty:
                story.append(Paragraph("3-Month Proforma Count Breakdown", heading_style))
                p_cols = list(prof_table.columns)
                p_data = [[Paragraph(f"<b>{col}</b>", cell_header_style) for col in p_cols]]
                
                for _, r in prof_table.iterrows():
                    p_row_cells = []
                    for col in p_cols:
                        p_row_cells.append(Paragraph(str(r[col]), cell_style))
                    p_data.append(p_row_cells)
                
                t_prof = Table(p_data, colWidths=[150, 105, 105, 105, 110])
                t_prof.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#0F172A')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]))
                story.append(t_prof)
                story.append(Spacer(1, 12))
            
            doc.build(story)
            buffer.seek(0)
            return buffer

        # PDF Download Trigger
        rev_pass = display_rev_df if len(unique_periods) >= 3 else pd.DataFrame()
        prof_pass = display_prof_df if len(unique_periods) >= 3 else pd.DataFrame()
        pdf_file = generate_comparison_pdf(total_revenue, total_proformas, rev_pass, prof_pass, monthly_df)
        
        st.download_button(
            label="📥 Export Professional 3-Month Report as PDF",
            data=pdf_file,
            file_name=f"3Month_Commercial_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )

else:
    st.info("👋 Upload 3 Excel sales files above to calculate 3-month performance and proformas.")
