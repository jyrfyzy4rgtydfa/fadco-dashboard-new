import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io

# Import ReportLab elements for PDF generation
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# Set up page config
st.set_page_config(page_title="Sales Performance Dashboard", layout="wide")
st.title("📊 Fadco Analytics Dashboard")
st.markdown("---")

# 1. Multi-File Uploader
uploaded_files = st.file_uploader("Upload one or more Excel Sales Files", type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    
    for uploaded_file in uploaded_files:
        try:
            df_temp = pd.read_excel(uploaded_file)
            df_temp.columns = [col.strip() for col in df_temp.columns]
            
            # Detect required columns dynamically
            date_col = [col for col in df_temp.columns if 'date' in col.lower()]
            comm_col = [col for col in df_temp.columns if 'commercial' in col.lower() or 'agent' in col.lower()]
            net_col = [col for col in df_temp.columns if 'net' in col.lower() or 'payer' in col.lower() or 'total' in col.lower()]
            
            if date_col and comm_col and net_col:
                df_temp = df_temp.rename(columns={comm_col[0]: 'Commercial', net_col[0]: 'Net_a_Payer', date_col[0]: 'Date'})
                df_temp['Date'] = pd.to_datetime(df_temp['Date'])
                df_temp['Net_a_Payer'] = pd.to_numeric(df_temp['Net_a_Payer'], errors='coerce').fillna(0)
                all_dfs.append(df_temp)
            else:
                st.warning(f"⚠️ Skipped '{uploaded_file.name}': Missing 'Date', 'Commercial', or 'Net à Payer' columns.")
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")

    if all_dfs:
        # Combine all uploaded excel sheets into one single database
        df = pd.concat(all_dfs, ignore_index=True)
        
        # --- Data Calculations ---
        # 1 & 2. Metrics per Commercial
        commercial_stats = df.groupby('Commercial').agg(
            Total_Proformas=('Net_a_Payer', 'count'),
            Total_Net_a_Payer=('Net_a_Payer', 'sum')
        ).reset_index().sort_values(by='Total_Net_a_Payer', ascending=False)

        # 3. Monthly Comparisons
        df['YearMonth'] = df['Date'].dt.to_period('M')
        monthly_sales = df.groupby('YearMonth')['Net_a_Payer'].sum().sort_index(ascending=False).reset_index()
        
        # --- KPI Overview Layout ---
        st.subheader("📈 Combined Quick Overview")
        kpi1, kpi2, kpi3 = st.columns(3)
        
        total_revenue = df['Net_a_Payer'].sum()
        total_proformas_count = df.shape[0]
        
        kpi1.metric(label="Total Combined Revenue", value=f"{total_revenue:,.2f} FCFA")
        kpi2.metric(label="Total Combined Proformas", value=f"{total_proformas_count:,}")
        
        growth_text = "N/A (Need 2+ months)"
        if len(monthly_sales) >= 2:
            last_month_val = monthly_sales.iloc[0]['Net_a_Payer']
            prev_month_val = monthly_sales.iloc[1]['Net_a_Payer']
            last_month_name = str(monthly_sales.iloc[0]['YearMonth'])
            prev_month_name = str(monthly_sales.iloc[1]['YearMonth'])
            
            growth_rate = ((last_month_val - prev_month_val) / prev_month_val) * 100 if prev_month_val != 0 else 0
            growth_text = f"{growth_rate:+.1f}% ({last_month_name} vs {prev_month_name})"
            
            if last_month_val > prev_month_val:
                kpi3.metric(label=f"Growth Trend", value=f"+{growth_rate:.1f}%", delta=f"More than {prev_month_name}")
            else:
                kpi3.metric(label=f"Growth Trend", value=f"{growth_rate:.1f}%", delta=f"Less than {prev_month_name}", delta_color="inverse")
        else:
            kpi3.metric(label="Growth Trend", value="N/A", delta="Insufficient timeline")

        # --- PDF Export Logic (Requirement 2) ---
        def generate_pdf(stats_df, total_rev, total_prof, growth_str):
            buffer = io.BytesIO()
            doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40)
            story = []
            
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=22, textColor=colors.HexColor('#1E3A8A'), spaceAfter=20)
            subtitle_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontSize=10, textColor=colors.gray, spaceAfter=20)
            heading_style = ParagraphStyle('HeadStyle', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor('#0F172A'), spaceBefore=15, spaceAfter=10)
            
            # Title
            story.append(Paragraph("Commercial Performance Report", title_style))
            story.append(Paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_style))
            story.append(Spacer(1, 10))
            
            # Summary Metrics Table
            summary_data = [
                [Paragraph("<b>Metric</b>", styles['Normal']), Paragraph("<b>Value</b>", styles['Normal'])],
                ["Total Combined Revenue", f"{total_rev:,.2f} €"],
                ["Total Combined Proformas", f"{total_prof:,}"],
                ["Month-over-Month Growth", growth_str]
            ]
            t_summary = Table(summary_data, colWidths=[200, 250])
            t_summary.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (1,0), colors.HexColor('#F1F5F9')),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ]))
            story.append(Paragraph("Executive Summary", heading_style))
            story.append(t_summary)
            story.append(Spacer(1, 20))
            
            # Commercial Breakdown Table
            story.append(Paragraph("Performance Breakdown by Commercial", heading_style))
            table_data = [[Paragraph("<b>Commercial Name</b>", styles['Normal']), 
                           Paragraph("<b>Total Proformas</b>", styles['Normal']), 
                           Paragraph("<b>Total Net à Payer</b>", styles['Normal'])]]
            
            for _, row in stats_df.iterrows():
                table_data.append([
                    str(row['Commercial']),
                    f"{row['Total_Proformas']:,}",
                    f"{row['Total_Net_a_Payer']:,.2f} FCFA"
                ])
                
            t_breakdown = Table(table_data, colWidths=[200, 110, 140])
            t_breakdown.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (2,0), colors.HexColor('#1E3A8A')),
                ('TEXTCOLOR', (0,0), (2,0), colors.whitesmoke),
                ('ALIGN', (1,0), (-1,-1), 'CENTER'),
                ('BOTTOMPADDING', (0,0), (-1,-1), 6),
                ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
                ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ]))
            story.append(t_breakdown)
            
            doc.build(story)
            buffer.seek(0)
            return buffer

            # Create download trigger
        pdf_data = generate_pdf(commercial_stats, total_revenue, total_proformas_count, growth_text)
        st.download_button(
            label="📥 Export Report as PDF",
            data=pdf_data,
            file_name=f"Sales_Report_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf"
        )
        
        st.markdown("---")

        # --- Dashboard Visuals ---
        col_left, col_right = st.columns(2)
        
        with col_left:
            st.subheader("👩‍💼 Performance Breakdown Table")
            st.dataframe(
                commercial_stats.style.format({'Total_Net_a_Payer': '{:,.2f} FCFA', 'Total_Proformas': '{:,}'}),
                use_container_width=True
            )
            
        with col_right:
            st.subheader("📊 Net à Payer by Commercial Chart")
            fig = px.bar(commercial_stats, x='Commercial', y='Total_Net_a_Payer', text_auto='.2s', labels={'Total_Net_a_Payer': 'Net à Payer'})
            st.plotly_chart(fig, use_container_width=True)
            
        st.markdown("---")
        st.subheader("📅 Combined Monthly Evolution")
        monthly_sales_str = monthly_sales.copy()
        monthly_sales_str['YearMonth'] = monthly_sales_str['YearMonth'].astype(str)
        fig_trend = px.line(monthly_sales_str.sort_values('YearMonth'), x='YearMonth', y='Net_a_Payer', markers=True)
        st.plotly_chart(fig_trend, use_container_width=True)
else:
    st.info("👋 You can select and drop multiple files together here to calculate everything all at once.")