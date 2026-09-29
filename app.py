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
# GLOBAL CONFIGURATION & LANGUAGE DICTIONARY
# ---------------------------------------------------------
CURRENCY_SYMBOL = "FCFA"

st.set_page_config(page_title="Commercial Performance Dashboard", layout="wide")

# Sidebar - Language Selection
st.sidebar.title("🌐 Language / Langue")
lang = st.sidebar.radio("Select Interface Language:", ["Français 🇫🇷", "English 🇬🇧"])
is_fr = lang == "Français 🇫🇷"

# Translation Dictionary
t = {
    "title": "📊 Tableau de Bord de Performance Commerciale Multi-Périodes" if is_fr else "📊 Multi-Period Commercial Performance Dashboard",
    "uploader": "Téléchargez un ou plusieurs fichiers Excel pour effectuer l'analyse" if is_fr else "Upload one or more Excel files to perform analysis",
    "kpi_header": "📌 Résumé Exécutif des Performances" if is_fr else "📌 Executive Performance Summary",
    "kpi_rev": "Chiffre d'Affaires Total" if is_fr else "Total Overall Revenue",
    "kpi_prof": "Nombre Total de Proformas" if is_fr else "Total Proformas Issued",
    "kpi_comm": "Commercials Actifs" if is_fr else "Active Commercials",
    "kpi_trend": "Tendance Dernier Mois" if is_fr else "Latest Month Trend",
    "req_2m": "Nécessite au moins 2 mois" if is_fr else "Requires data from 2+ months",
    
    # Table 1
    "t1_header": "👩‍💼 1. Répartition du Chiffre d'Affaires Mensuel par Commercial" if is_fr else "👩‍💼 1. Monthly Revenue Breakdown per Commercial",
    "eval_period": "💡 Évalué sur les périodes :" if is_fr else "💡 Evaluated across periods:",
    "rev_col": "CA" if is_fr else "Rev.",
    "latest": "Dernier" if is_fr else "Latest",
    "t1_total": "CA Total sur la Période" if is_fr else "Total Revenue Across Periods",
    "btn_t1_pdf": "📄 Imprimer / Télécharger Tableau 1 (CA) en PDF" if is_fr else "📄 Print / Download Table 1 (Revenue) as PDF",

    # Table 2
    "t2_header": "📑 2. Répartition du Nombre de Proformas Mensuel par Commercial" if is_fr else "📑 2. Monthly Proforma Count Breakdown per Commercial",
    "prof_col": "Proformas" if is_fr else "Proformas",
    "t2_total": "Total Proformas sur la Période" if is_fr else "Total Proformas Across Periods",
    "btn_t2_pdf": "📄 Imprimer / Télécharger Tableau 2 (Proformas) en PDF" if is_fr else "📄 Print / Download Table 2 (Proformas) as PDF",

    # Table 3
    "t3_header": "🏆 3. Verdict de Progression Commerciale (Tableau Comparatif)" if is_fr else "🏆 3. Commercial Progression Verdict (Comparison Table)",
    "col_comm": "Commercial" if is_fr else "Commercial",
    "col_status": "Statut du Verdict" if is_fr else "Verdict Status",
    "col_rev_diff": "Écart CA (Période Récente vs Ancienne)" if is_fr else "Revenue Diff (Recent vs Earlier)",
    "col_prof_diff": "Écart Proforma (Période Récente vs Ancienne)" if is_fr else "Proforma Diff (Recent vs Earlier)",
    "col_exp": "Explication du Résumé" if is_fr else "Summary Explanation",
    
    # Verdict Badges & Descriptions
    "st_strong": "🟢 Forte Avancée" if is_fr else "🟢 Strong Advance",
    "exp_strong": "Le CA et le nombre de proformas ont augmenté" if is_fr else "Both Revenue and Proformas increased",
    "st_rev": "🔵 Avancée du CA" if is_fr else "🔵 Revenue Advance",
    "exp_rev": "Le CA a augmenté malgré moins de proformas" if is_fr else "Revenue grew despite fewer proformas",
    "st_vol": "🟡 Avancée du Volume" if is_fr else "🟡 Volume Advance",
    "exp_vol": "Plus de proformas émises, mais CA total inférieur" if is_fr else "More proformas issued, but lower total revenue",
    "st_dec": "🔴 En Déclin" if is_fr else "🔴 Declining",
    "exp_dec": "Le CA et le nombre de proformas ont diminué" if is_fr else "Both Revenue and Proforma count dropped",
    
    # Export Options
    "export_t3_title": "📥 Sélectionnez le Type d'Exportation pour le Tableau 3 :" if is_fr else "📥 Select Export Type for Table 3:",
    "btn_t3_admin": "👑 Exporter PDF ADMIN (Toutes les Infos + Chiffre d'Affaires)" if is_fr else "👑 Export ADMIN PDF (Full Info + Revenue)",
    "btn_t3_user": "👤 Exporter PDF UTILISATEUR (Sans Chiffre d'Affaires)" if is_fr else "👤 Export USER PDF (Without Revenue)",
    
    # Section 5 & Warnings
    "chart_title": "Graphique Comparatif des Revenus par Période" if is_fr else "Side-by-Side Revenue Comparison Chart",
    "monthly_header": "📅 Évolution Mensuelle Consolidée du Chiffre d'Affaires" if is_fr else "📅 Combined Monthly Revenue Evolution",
    "warn_single_period": "ℹ️ 1 seule période détectée : Téléchargez au moins 2 fichiers/mois pour voir l'analyse comparative de progression." if is_fr else "ℹ️ Single period detected: Upload 2+ files/months to enable comparison verdict.",
    "upload_prompt": "👋 Veuillez télécharger au moins 1 fichier Excel ci-dessus pour générer l'analyse." if is_fr else "👋 Upload 1 or more Excel sales files above to generate the analysis."
}

st.title(t["title"])
st.markdown("---")

# File Uploader
uploaded_files = st.file_uploader(t["uploader"], type=["xlsx", "xls"], accept_multiple_files=True)

def format_currency(val):
    if CURRENCY_SYMBOL in ["$", "€", "£"]:
        return f"{CURRENCY_SYMBOL}{val:,.2f}"
    return f"{val:,.0f} {CURRENCY_SYMBOL}"

def format_signed_currency(val):
    sign = "+" if val > 0 else ""
    if CURRENCY_SYMBOL in ["$", "€", "£"]:
        return f"{sign}{CURRENCY_SYMBOL}{val:,.2f}"
    return f"{sign}{val:,.0f} {CURRENCY_SYMBOL}"

# Generic Function to Generate PDF for any DataFrame Table
def generate_table_pdf(title_text, df_to_print):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=30)
    story = []
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontSize=15, textColor=colors.HexColor('#1E3A8A'), spaceAfter=12)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontSize=7, leading=9)
    cell_header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'], fontSize=7, leading=9, textColor=colors.whitesmoke)

    story.append(Paragraph(title_text, title_style))
    story.append(Paragraph(f"{'Généré le' if is_fr else 'Generated on'}: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles['Normal']))
    story.append(Spacer(1, 10))
    
    if not df_to_print.empty:
        cols = list(df_to_print.columns)
        table_data = [[Paragraph(f"<b>{col}</b>", cell_header_style) for col in cols]]
        
        for _, r in df_to_print.iterrows():
            row_cells = []
            for col in cols:
                val = str(r[col])
                val = val.replace("🟢 ", "").replace("🔵 ", "").replace("🟡 ", "").replace("🔴 ", "")
                row_cells.append(Paragraph(val, cell_style))
            table_data.append(row_cells)
        
        num_cols = len(cols)
        col_width = 570 / num_cols
        
        t_pdf = Table(table_data, colWidths=[col_width] * num_cols)
        t_pdf.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor('#F8FAFC')]),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('TOPPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_pdf)
    
    doc.build(story)
    buffer.seek(0)
    return buffer


if uploaded_files:
    all_dfs = []
    
    for uploaded_file in uploaded_files:
        try:
            df_temp = pd.read_excel(uploaded_file)
            df_temp.columns = [col.strip() for col in df_temp.columns]
            
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
                st.warning(f"⚠️ Skipped '{uploaded_file.name}': Missing required columns.")
        except Exception as e:
            st.error(f"Error reading {uploaded_file.name}: {e}")

    if all_dfs:
        df = pd.concat(all_dfs, ignore_index=True)
        df['YearMonth'] = df['Date'].dt.to_period('M')
        unique_periods = sorted(df['YearMonth'].unique())
        
        # Section 1: KPIs Overview
        st.subheader(t["kpi_header"])
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        total_revenue = df['Net_a_Payer'].sum()
        total_proformas = len(df)
        total_commercials = df['Commercial'].nunique()
        
        kpi1.metric(t["kpi_rev"], format_currency(total_revenue))
        kpi2.metric(t["kpi_prof"], f"{total_proformas:,}")
        kpi3.metric(t["kpi_comm"], f"{total_commercials}")
        
        monthly_totals = df.groupby('YearMonth')['Net_a_Payer'].sum().sort_index(ascending=False).reset_index()
        
        if len(monthly_totals) >= 2:
            latest_month_val = monthly_totals.iloc[0]['Net_a_Payer']
            prev_month_val = monthly_totals.iloc[1]['Net_a_Payer']
            growth = ((latest_month_val - prev_month_val) / prev_month_val) * 100 if prev_month_val != 0 else 0
            kpi4.metric(
                label=f"{t['kpi_trend']} ({monthly_totals.iloc[0]['YearMonth']})", 
                value=f"{growth:+.1f}%", 
                delta=format_currency(latest_month_val - prev_month_val)
            )
        else:
            kpi4.metric(t["kpi_trend"], "N/A", delta=t["req_2m"])

        st.markdown("---")

        # ---------------------------------------------------------
        # TABLE 1: REVENUE BREAKDOWN (WORKS FOR ANY NUMBER OF PERIODS)
        # ---------------------------------------------------------
        st.subheader(t["t1_header"])
        st.info(f"{t['eval_period']} **{', '.join([str(p) for p in unique_periods])}**")
        
        pivot_rev = df.pivot_table(index='Commercial', columns='YearMonth', values='Net_a_Payer', aggfunc='sum', fill_value=0)
        
        display_rev_df = pd.DataFrame(index=pivot_rev.index)
        for p in reversed(unique_periods):
            display_rev_df[f"{t['rev_col']} {p}"] = pivot_rev[p].apply(format_currency)
            
        display_rev_df[t["t1_total"]] = pivot_rev.sum(axis=1).apply(format_currency)
        display_rev_df = display_rev_df.reset_index().rename(columns={'Commercial': t['col_comm']})
        
        st.dataframe(display_rev_df, use_container_width=True)
        
        pdf_table1 = generate_table_pdf(t["t1_header"], display_rev_df)
        st.download_button(
            label=t["btn_t1_pdf"],
            data=pdf_table1,
            file_name=f"Table1_Revenue_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            key="btn_pdf_table1"
        )

        # ---------------------------------------------------------
        # TABLE 2: PROFORMA BREAKDOWN (WORKS FOR ANY NUMBER OF PERIODS)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader(t["t2_header"])
        
        pivot_prof = df.pivot_table(index='Commercial', columns='YearMonth', values='Net_a_Payer', aggfunc='count', fill_value=0)
        
        display_prof_df = pd.DataFrame(index=pivot_prof.index)
        for p in reversed(unique_periods):
            display_prof_df[f"{t['prof_col']} {p}"] = pivot_prof[p]
            
        display_prof_df[t["t2_total"]] = pivot_prof.sum(axis=1)
        display_prof_df = display_prof_df.reset_index().rename(columns={'Commercial': t['col_comm']})
        
        st.dataframe(display_prof_df, use_container_width=True)
        
        pdf_table2 = generate_table_pdf(t["t2_header"], display_prof_df)
        st.download_button(
            label=t["btn_t2_pdf"],
            data=pdf_table2,
            file_name=f"Table2_Proformas_{datetime.now().strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
            key="btn_pdf_table2"
        )

        # ---------------------------------------------------------
        # TABLE 3: EXECUTIVE CONCLUSION MATRIX (IF 2+ PERIODS)
        # ---------------------------------------------------------
        st.markdown("---")
        st.subheader(t["t3_header"])
        
        if len(unique_periods) >= 2:
            p_first = unique_periods[0]
            p_latest = unique_periods[-1]
            
            conclusion_rows = []
            
            for comm in pivot_rev.index:
                rev_first = pivot_rev.loc[comm, p_first]
                rev_latest = pivot_rev.loc[comm, p_latest]
                rev_diff = rev_latest - rev_first
                
                prof_first = pivot_prof.loc[comm, p_first]
                prof_latest = pivot_prof.loc[comm, p_latest]
                prof_diff = prof_latest - prof_first
                
                if rev_diff > 0 and prof_diff > 0:
                    status = t["st_strong"]
                    explanation = t["exp_strong"]
                elif rev_diff > 0 and prof_diff <= 0:
                    status = t["st_rev"]
                    explanation = t["exp_rev"]
                elif rev_diff <= 0 and prof_diff > 0:
                    status = t["st_vol"]
                    explanation = t["exp_vol"]
                else:
                    status = t["st_dec"]
                    explanation = t["exp_dec"]
                    
                conclusion_rows.append({
                    t['col_comm']: comm,
                    t['col_status']: status,
                    f"{t['col_rev_diff']} ({p_latest} vs {p_first})": format_signed_currency(rev_diff),
                    f"{t['col_prof_diff']} ({p_latest} vs {p_first})": f"{prof_diff:+} proformas",
                    t['col_exp']: explanation,
                    'raw_rev_diff': rev_diff
                })
                
            conclusion_df = pd.DataFrame(conclusion_rows).sort_values(by='raw_rev_diff', ascending=False)
            
            col_rev_diff_name = f"{t['col_rev_diff']} ({p_latest} vs {p_first})"
            col_prof_diff_name = f"{t['col_prof_diff']} ({p_latest} vs {p_first})"
            
            table3_display = conclusion_df[[
                t['col_comm'], 
                t['col_status'], 
                col_rev_diff_name, 
                col_prof_diff_name, 
                t['col_exp']
            ]]
            
            st.dataframe(table3_display, use_container_width=True)
            
            # Dual PDF Export Section
            st.write(t["export_t3_title"])
            exp_col1, exp_col2 = st.columns(2)
            
            with exp_col1:
                table3_admin_df = table3_display.copy()
                admin_pdf_title = f"{t['t3_header']} (ADMIN)"
                pdf_table3_admin = generate_table_pdf(admin_pdf_title, table3_admin_df)
                
                st.download_button(
                    label=t["btn_t3_admin"],
                    data=pdf_table3_admin,
                    file_name=f"Table3_ADMIN_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    key="btn_pdf_table3_admin"
                )
                
            with exp_col2:
                table3_user_df = table3_display[[t['col_comm'], t['col_status'], col_prof_diff_name, t['col_exp']]].copy()
                user_pdf_title = f"{t['t3_header']} (UTILISATEUR / STAFF)" if is_fr else f"{t['t3_header']} (USER / STAFF)"
                pdf_table3_user = generate_table_pdf(user_pdf_title, table3_user_df)
                
                st.download_button(
                    label=t["btn_t3_user"],
                    data=pdf_table3_user,
                    file_name=f"Table3_USER_{datetime.now().strftime('%Y%m%d')}.pdf",
                    mime="application/pdf",
                    key="btn_pdf_table3_user"
                )
        else:
            st.info(t["warn_single_period"])

        # Chart Visualization
        st.markdown("---")
        fig_comp = go.Figure()
        for p in unique_periods:
            fig_comp.add_trace(go.Bar(
                x=pivot_rev.index, 
                y=pivot_rev[p], 
                name=str(p)
            ))
        fig_comp.update_layout(
            barmode='group', 
            title=t["chart_title"], 
            xaxis_title=t["col_comm"], 
            yaxis_title=f"Net à Payer ({CURRENCY_SYMBOL})"
        )
        st.plotly_chart(fig_comp, use_container_width=True)

        st.markdown("---")

        # Section 5: Monthly Revenue Evolution
        st.subheader(t["monthly_header"])
        
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
                title=t["monthly_header"],
                labels={'Total_Revenue': 'Revenue', 'YearMonth': 'Month'}
            )
            st.plotly_chart(fig_line, use_container_width=True)

else:
    st.info(t["upload_prompt"])
