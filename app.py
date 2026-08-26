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
st.title("📊 Multi-Period Commercial Performance Dashboard")
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

def format_signed_currency(val):
    sign = "+" if val > 0 else ""
    if CURRENCY_SYMBOL in ["$", "€", "£"]:
        return f"{sign}{CURRENCY_SYMBOL}{val:,.2f}"
    return f"{sign}{val:,.0f} {CURRENCY_SYMBOL}"

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
        df = pd.concat(all_dfs, ignore_index=True)
        df['YearMonth'] = df['Date'].dt.to_period('M')
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

        if len(unique_periods) >= 3:
            periods_3 = unique_periods[-3:]
            p1, p2, p3 = periods_3[0], periods_3[1], periods_3[2]
            
            # --- Section 2: Revenue Table ---
            st.subheader("👩‍💼 1. Monthly Revenue breakdown per Commercial")
            st.info(f"💡 Evaluated across periods: **{p1}** ➔ **{p2}** ➔ **{p3}**")
            
            pivot_rev = df.pivot_table(index='Commercial', columns='YearMonth', values='Net_a_Payer', aggfunc='sum', fill_value=0)
            
            display_rev_df = pd.DataFrame(index=pivot_rev.index)
            display_rev_df[f"Rev. {p3} (Latest)"] = pivot_rev[p3].apply(format_currency)
            display_rev_df[f"Rev. {p2}"] = pivot_rev[p2].apply(format_currency)
            display_rev_df[f"Rev. {p1}"] = pivot_rev[p1].apply(format_currency)
            display_rev_df["Total 3-Month Revenue"] = (pivot_rev[p3] + pivot_rev[p2] + pivot_rev[p1]).apply(format_currency)
            
            st.dataframe(display_rev_df.reset_index(), use_container_width=True)

            # --- Section 3: Proforma Count Table ---
            st.markdown("---")
            st.subheader("📑 2. Monthly Proforma Count Breakdown per Commercial")
            
            pivot_prof = df.pivot_table(index='Commercial', columns='YearMonth', values='Net_a_Payer', aggfunc='count', fill_value=0)
            
            display_prof_df = pd.DataFrame(index=pivot_prof.index)
            display_prof_df[f"Proformas {p3} (Latest)"] = pivot_prof[p3]
            display_prof_df[f"Proformas {p2}"] = pivot_prof[p2]
            display_prof_df[f"Proformas {p1}"] = pivot_prof[p1]
            display_prof_df["Total 3-Month Proformas"] = pivot_prof[p3] + pivot_prof[p2] + pivot_prof[p1]
            
            st.dataframe(display_prof_df.reset_index(), use_container_width=True)

            # --- Section 4: CLEAR TABLE 3 - EXECUTIVE CONCLUSION MATRIX ---
            st.markdown("---")
            st.subheader("🏆 3. Commercial Progression Verdict (Comparison Table)")
            st.write("This clear summary evaluates each commercial's overall 3-month progression by comparing **Month 3 vs. Month 1** across both **Revenue** and **Proforma Volume**.")

            conclusion_rows = []
            
            for comm in pivot_rev.index:
                rev_m1 = pivot_rev.loc[comm, p1]
                rev_m3 = pivot_rev.loc[comm, p3]
                rev_diff = rev_m3 - rev_m1
                
                prof_m1 = pivot_prof.loc[comm, p1]
                prof_m3 = pivot_prof.loc[comm, p3]
                prof_diff = prof_m3 - prof_m1
                
                # Clear Categorization Logic
                if rev_diff > 0 and prof_diff > 0:
                    status = "🟢 Strong Advance"
                    explanation = "Both Revenue and Proformas increased"
                elif rev_diff > 0 and prof_diff <= 0:
                    status = "🔵 Revenue Advance"
                    explanation = "Revenue grew despite fewer proformas"
                elif rev_diff <= 0 and prof_diff > 0:
                    status = "🟡 Volume Advance"
                    explanation = "More proformas issued, but lower total revenue"
                else:
                    status = "🔴 Declining"
                    explanation = "Both Revenue and Proforma count dropped"
                    
                conclusion_rows.append({
                    'Commercial': comm,
                    'Verdict Status': status,
                    'Revenue Diff (M3 vs M1)': format_signed_currency(rev_diff),
                    'Proforma Diff (M3 vs M1)': f"{prof_diff:+} proformas",
                    'Summary Explanation': explanation,
                    'M1 Revenue': format_currency(rev_m1),
                    'M3 Revenue': format_currency(rev_m3),
                    'M1 Proformas': prof_m1,
                    'M3 Proformas': prof_m3,
                    'raw_rev_diff': rev_diff
                })
                
            conclusion_df = pd.DataFrame(conclusion_rows).sort_values(by='raw_rev_diff', ascending=False)
            
            # Quick Stats Callout Cards above Table 3
            advancing_count = len(conclusion_df[conclusion_df['Verdict Status'].str.contains('Advance')])
            declining_count = len(conclusion_df[conclusion_df['Verdict Status'].str.contains('Declining')])
            
            c_col1, c_col2, c_col3 = st.columns(3)
            c_col1.metric("Total Commercials Analyzed", len(conclusion_df))
            c_col2.metric("Commercials Advancing 🟢", advancing_count)
            c_col3.metric("Commercials Declining 🔴", declining_count)
            
            # Simplified columns display for Table 3
            table3_display = conclusion_df[[
                'Commercial', 
                'Verdict Status', 
                'Revenue Diff (M3 vs M1)', 
                'Proforma Diff (M3 vs M1)', 
                'Summary Explanation',
                'M1 Revenue', 
                'M3 Revenue'
            ]]
            
            st.dataframe(table3_display, use_container_width=True)

            # Chart Visualization
            fig_comp = go.Figure()
            for p in periods_3:
                fig_comp.add_trace(go.Bar(
                    x=pivot_rev.index, 
                    y=pivot_rev[p], 
                    name=str(p)
                ))
            fig_comp.update_layout(
                barmode='group', 
                title="3-Month Side-by-Side Revenue Comparison Chart", 
                xaxis_title="Commercial", 
                yaxis_title=f"Net à Payer ({CURRENCY_SYMBOL})"
            )
            st.plotly_chart(fig_comp, use_container_width=True)

        else:
            st.warning(f"⚠️ Please upload data spanning across at least 3 distinct months. Currently detected: {len(unique_periods)} month(s).")

        st.markdown("---")

        # --- Section 5: Monthly Evolution ---
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

        # --- Section 6: PDF Export Logic ---
        def generate_comparison_pdf(overall_rev, overall_prof, conc_df):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor('#1E3A8A'), spaceAfter=15)
            heading_style = ParagraphStyle('HeadStyle', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0F172A'), spaceBefore=12, spaceAfter=8)
            
            cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=8, leading=10)
            cell_header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.whitesmoke)

            story.append(Paragraph("Commercial Executive Progression Report", title_style))
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
            
            # Final Conclusion Table in PDF
            if not conc_df.empty:
                story.append(Paragraph("Commercial Progression Verdict Matrix", heading_style))
                cols = ['Commercial', 'Verdict Status', 'Revenue Diff (M3 vs M1)', 'Proforma Diff (M3 vs M1)']
                c_data = [[Paragraph(f"<b>{col}</b>", cell_header_style) for col in cols]]
                
                for _, r in conc_df.iterrows():
                    clean_status = str(r['Verdict Status']).replace("🟢 ", "").replace("🔵 ", "").replace("🟡 ", "").replace("🔴 ", "")
                    c_data.append([
                        Paragraph(str(r['Commercial']), cell_style),
                        Paragraph(clean_status, cell_style),
                        Paragraph(str(r['Revenue Diff (M3 vs M1)']), cell_style),
                        Paragraph(str(r['Proforma Diff (M3 vs M1)']), cell_style)
                    ])
                
                t_conc = Table(c_data, colWidths=[150, 140, 140, 140])
                t_conc.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
                    ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
                    ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
                    ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ]))
                story.append(t_conc)
            
            doc.build(story)
            buffer.seek(0)
            return buffer

        # PDF Download Trigger
        conc_pass = conclusion_df if len(unique_periods) >= 3 else pd.DataFrame()
        pdf_file = generate_comparison_pdf(total_revenue, total_proformas, conc_pass)
        
        st.download_button(
            label="📥 Export Executive Conclusion Report as PDF",
            data=pdf_file,
            file_name=f"Executive_Conclusion_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )

else:
    st.info("👋 Upload 3 Excel sales files above to generate the analysis.")
