import streamlit as st
import pandas as pd
import plotly.express as px
import base64
from io import BytesIO
from datetime import timedelta, datetime, time
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# Streamlit app
def main():
    # Set page config as the first Streamlit command
    st.set_page_config(page_title="eSanjeevani Consultation Analysis", layout="wide")

    # Custom CSS for professional styling
    st.markdown("""
        <style>
        .main {
            background-color: #f8fafc;
            padding: 15px;
            font-family: 'Arial', sans-serif;
        }
        .stApp {
            max-width: 1400px;
            margin: 0 auto;
        }
        h1 {
            color: #1e40af;
            font-size: 26px;
            margin-bottom: 10px;
        }
        h2 {
            color: #1e40af;
            font-size: 18px;
            margin-top: 8px;
            margin-bottom: 5px;
        }
        .stButton>button {
            background-color: #1e40af;
            color: white;
            border-radius: 6px;
            padding: 6px 12px;
            font-size: 14px;
            margin-right: 8px;
        }
        .stButton>button:hover {
            background-color: #2563eb;
        }
        .stDataFrame {
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .stExpander {
            background-color: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            margin-bottom: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .sidebar .sidebar-content {
            background-color: #ffffff;
            border-right: 1px solid #e5e7eb;
            padding: 10px;
        }
        .stMetric {
            background-color: #ffffff;
            border: 1px solid #e5e7eb;
            border-radius: 6px;
            padding: 8px;
            margin-bottom: 8px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            font-size: 14px;
        }
        .stMetric label {
            font-size: 14px;
            color: #1e40af;
        }
        .stMetric div {
            font-size: 16px;
            font-weight: bold;
        }
        </style>
    """, unsafe_allow_html=True)

    # Initialize session state for date inputs
    if 'start_date' not in st.session_state:
        st.session_state.start_date = datetime.today().date()  # 04:24 PM IST, Thursday, July 17, 2025
    if 'end_date' not in st.session_state:
        st.session_state.end_date = datetime.today().date()  # 04:24 PM IST, Thursday, July 17, 2025

    # Main container for all content
    with st.container():
        # Header section
        st.title("eSanjeevani Teleconsultation Report")
        st.markdown("Upload an Excel file to analyze consultation completion scores, time taken, and missing fields.", unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Choose Excel File", type=["xlsx"], help="Upload an Excel file (up to 1000 MB).", key="uploader")

    # Sidebar for filters
    with st.sidebar:
        st.header("Filters")
        st.subheader("Date Range")
        start_date = st.date_input("Start Date", value=st.session_state.start_date, key="start_date_unique")
        end_date = st.date_input("End Date", value=st.session_state.end_date, key="end_date_unique")
        
        # Update session state with current selections
        st.session_state.start_date = start_date
        st.session_state.end_date = end_date
        
        st.subheader("Patient Search")
        patient_id = st.text_input("Patient ID", "", key="patient_id")
        patient_name = st.text_input("Patient Name", "", key="patient_name")

    if uploaded_file is not None:
        try:
            # Load Excel file
            df = pd.read_excel(uploaded_file)
            
            # Ensure ConsultationCreatedDate is in datetime format
            df['ConsultationCreatedDate'] = pd.to_datetime(df['ConsultationCreatedDate'], errors='coerce')
            
            # Update session state with data-driven date defaults if available
            if not df['ConsultationCreatedDate'].isna().all():
                min_date = df['ConsultationCreatedDate'].min().date()
                max_date = df['ConsultationCreatedDate'].max().date()
                st.session_state.start_date = min_date
                st.session_state.end_date = max_date
            
            # Filter data by date range
            filtered_df = filter_by_date_range(df, st.session_state.start_date, st.session_state.end_date)
            
            # Further filter by patient ID or name
            filtered_df = filter_by_patient_search(filtered_df, patient_id, patient_name)
            
            # Generate report
            report_df = generate_consultation_report(filtered_df)
            
            if report_df is not None and not report_df.empty:
                # Calculate summary statistics
                total_patients = len(report_df)
                avg_score = report_df['CompletionScore (%)'].mean()
                max_score = report_df['CompletionScore (%)'].max()
                min_score = report_df['CompletionScore (%)'].min()
                avg_missing_score = report_df['MissingFieldScore'].mean()
                avg_time, valid_count, error_messages = calculate_average_consultation_time(filtered_df)
                
                # Calculate percentages for pie chart
                high_score_count = len(report_df[report_df['CompletionScore (%)'] >= 75])
                low_score_count = len(report_df[report_df['CompletionScore (%)'] < 50])
                other_score_count = total_patients - high_score_count - low_score_count
                high_score_percent = (high_score_count / total_patients * 100) if total_patients > 0 else 0
                low_score_percent = (low_score_count / total_patients * 100) if total_patients > 0 else 0
                other_score_percent = (other_score_count / total_patients * 100) if total_patients > 0 else 0

                # Dashboard section
                with st.container():
                    st.subheader("Dashboard")
                    col1, col2 = st.columns([1, 1])
                    with col1:
                        pie_data = pd.DataFrame({
                            'Category': ['Score >= 75%', 'Score < 50%', 'Score 50-75%'],
                            'Percentage': [high_score_percent, low_score_percent, other_score_percent]
                        })
                        fig = px.pie(pie_data, values='Percentage', names='Category', title='Score Distribution')
                        fig.update_layout(margin=dict(t=30, b=10, l=10, r=10))
                        st.plotly_chart(fig, use_container_width=True)
                    with col2:
                        st.markdown("**Summary Statistics**")
                        col3, col4 = st.columns(2)
                        with col3:
                            st.metric("Total Patients", total_patients)
                            st.metric("Avg Completion Score", f"{avg_score:.2f}%")
                            st.metric("Avg MissingFieldScore", f"{avg_missing_score:.2f}%")
                            st.metric("Max Completion Score", f"{max_score:.2f}%")
                        with col4:
                            st.metric("Min Completion Score", f"{min_score:.2f}%")
                            st.metric("Avg Consultation Time", avg_time)
                            st.metric("Score >= 75%", f"{high_score_percent:.2f}%")
                            st.metric("Score < 50%", f"{low_score_percent:.2f}%")

                # Patient Consultation Report
                with st.expander("Patient Consultation Report", expanded=True):
                    st.dataframe(report_df, use_container_width=True, height=300)

                # Download buttons
                with st.container():
                    st.subheader("Download Reports")
                    col5, col6, col7 = st.columns(3)
                    with col5:
                        csv = convert_df_to_csv(report_df)
                        st.download_button(
                            label="CSV",
                            data=csv,
                            file_name="consultation_report.csv",
                            mime="text/csv",
                        )
                    with col6:
                        excel_data = generate_excel_report(report_df, total_patients, avg_score, max_score, min_score, avg_time, high_score_percent, low_score_percent, avg_missing_score)
                        st.download_button(
                            label="Excel",
                            data=excel_data,
                            file_name="consultation_report.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )
                    with col7:
                        pdf_data = generate_pdf_report(report_df, total_patients, avg_score, max_score, min_score, avg_time, high_score_percent, low_score_percent, avg_missing_score)
                        st.download_button(
                            label="PDF",
                            data=pdf_data,
                            file_name="consultation_report.pdf",
                            mime="application/pdf",
                        )

                # Notes section
                with st.expander("Notes", expanded=False):
                    st.markdown(f"**Time Processing Info:** Processed {valid_count} valid time entries.")
                    if error_messages:
                        st.markdown("**Warnings:**\n- " + "\n- ".join(error_messages[:5]) + ("\n- ...and more" if len(error_messages) > 5 else ""))
            else:
                st.error("No data to display. The file may be empty, incorrectly formatted, or no records match the filters.")
        
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
    else:
        st.info("Please upload an Excel file to proceed.")

# Function to parse HH_MM_SS column to seconds for average calculation
def parse_hh_mm_ss(time_input):
    try:
        if isinstance(time_input, time):
            time_str = time_input.strftime('%H:%M:%S')
        else:
            time_str = str(time_input)
            if pd.isna(time_str) or time_str == '':
                return None, "Empty or NaN"
        
        parts = time_str.split(':')
        if len(parts) == 2:
            minutes, seconds = map(int, parts)
            hours = 0
        elif len(parts) == 3:
            hours, minutes, seconds = map(int, parts)
        else:
            return None, f"Invalid format: {time_str}"
        
        if hours < 0 or minutes < 0 or seconds < 0 or minutes >= 60 or seconds >= 60:
            return None, f"Invalid time values: {time_str}"
        
        return hours * 3600 + minutes * 60 + seconds, None
    except Exception as e:
        return None, f"Error parsing {time_str}: {str(e)}"

# Function to get time taken from HH_MM_SS column
def get_time_taken(row):
    try:
        time_input = row.get('HH_MM_SS', 'Unknown')
        if pd.isna(time_input) or time_input == '':
            return "Unknown"
        if isinstance(time_input, time):
            return time_input.strftime('%H:%M:%S')
        return str(time_input)
    except:
        return "Unknown"

# Function to calculate consultation completion score, filled fields, and missing fields
def calculate_completion_score(row):
    score = 0
    total_fields = 8
    filled_fields = []
    missing_fields = []

    # Fields based on assumed image data
    fields_to_check = [
        'PatientName',
        'Age',
        'GenderDisplay',
        'ConsultationCreatedDate',
        'ConsultationStatus',
        'Symptoms_',
        'Provisional Diagnosis',
        'Advice'
    ]

    for field in fields_to_check:
        value = row.get(field)
        if pd.isna(value):
            missing_fields.append(field)
        elif field == 'Symptoms_' and isinstance(value, str) and '"Alias"' in value:
            # Handle malformed Symptoms_ data (e.g., [{Alias":"Common Cold")
            score += 1
            filled_fields.append(field)
        elif value != '':
            score += 1
            filled_fields.append(field)
        else:
            missing_fields.append(field)

    # Calculate completion percentage
    completion_percentage = (score / total_fields) * 100 if total_fields > 0 else 0
    # Calculate missing fields percentage
    missing_percentage = ((total_fields - score) / total_fields) * 100 if total_fields > 0 else 0
    return round(completion_percentage, 2), filled_fields, missing_fields, round(missing_percentage, 2), fields_to_check

# Calculate average consultation time from HH_MM_SS
def calculate_average_consultation_time(df):
    total_seconds = 0
    valid_count = 0
    error_messages = []
    for _, row in df.iterrows():
        time_seconds, error = parse_hh_mm_ss(row.get('HH_MM_SS', ''))
        if time_seconds is not None:
            total_seconds += time_seconds
            valid_count += 1
        else:
            if error:
                error_messages.append(f"Row {row.name + 2}: {error}")
    if valid_count > 0:
        avg_seconds = total_seconds / valid_count
        minutes = int(avg_seconds // 60)
        seconds = int(avg_seconds % 60)
        result = f"{minutes:02d}:{seconds:02d}"
    else:
        result = "00:00"
        error_messages.append("No valid time entries found in HH_MM_SS column.")
    
    return result, valid_count, error_messages

# Filter DataFrame by date range
def filter_by_date_range(df, start_date, end_date):
    try:
        df['ConsultationCreatedDate'] = pd.to_datetime(df['ConsultationCreatedDate'], errors='coerce')
        mask = (df['ConsultationCreatedDate'].dt.date >= start_date) & (df['ConsultationCreatedDate'].dt.date <= end_date)
        return df[mask]
    except Exception as e:
        st.error(f"Error filtering by date: {e}")
        return df

# Filter DataFrame by PatientId or PatientName
def filter_by_patient_search(df, patient_id, patient_name):
    try:
        filtered_df = df
        if patient_id:
            filtered_df = filtered_df[filtered_df['PatientId'].astype(str).str.contains(patient_id, case=False, na=False)]
        if patient_name:
            filtered_df = filtered_df[filtered_df['PatientName'].astype(str).str.contains(patient_name, case=False, na=False)]
        return filtered_df
    except Exception as e:
        st.error(f"Error filtering by patient search: {e}")
        return df

# Process the data and generate report
def generate_consultation_report(df):
    if df is None or df.empty:
        return None

    report = []
    for index, row in df.iterrows():
        patient_id = row.get('PatientId', 'Unknown')
        consultation_id = row.get('ConsultationId', 'Unknown')
        completion_score, filled_fields, missing_fields, missing_percentage, _ = calculate_completion_score(row)
        time_taken = get_time_taken(row)
        
        patient_report = {
            'PatientId': patient_id,
            'ConsultationId': consultation_id,
            'CompletionScore (%)': completion_score,
            'CompletionField': ', '.join(filled_fields),
            'MissingFieldScore': missing_percentage,
            'MissingFields': ', '.join(missing_fields),
            'TimeTaken (MM:SS)': time_taken,
            'Status': row.get('ConsultationStatus', 'Unknown'),
            'Symptoms': row.get('Symptoms_', ''),
            'Diagnosis': row.get('Provisional Diagnosis', ''),
            'Advice': row.get('Advice', '')
        }
        report.append(patient_report)

    report_df = pd.DataFrame(report)
    column_order = [
        'PatientId', 'ConsultationId', 'CompletionScore (%)', 'CompletionField',
        'MissingFieldScore', 'MissingFields', 'TimeTaken (MM:SS)', 'Status',
        'Symptoms', 'Diagnosis', 'Advice'
    ]
    report_df = report_df[column_order]
    return report_df

# Generate Excel report with four sheets
def generate_excel_report(report_df, total_patients, avg_score, max_score, min_score, avg_time, high_score_percent, low_score_percent, avg_missing_score):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        report_df.to_excel(writer, sheet_name='Patient Report', index=False)
        high_score_df = report_df[report_df['CompletionScore (%)'] >= 75]
        high_score_df.to_excel(writer, sheet_name='Score >= 75%', index=False)
        mid_score_df = report_df[(report_df['CompletionScore (%)'] >= 50) & (report_df['CompletionScore (%)'] < 75)]
        mid_score_df.to_excel(writer, sheet_name='Score 50-75%', index=False)
        dashboard_data = {
            'Category': ['Score >= 75%', 'Score < 50%', 'Score 50-75%'],
            'Percentage': [high_score_percent, low_score_percent, (100 - high_score_percent - low_score_percent)]
        }
        dashboard_df = pd.DataFrame(dashboard_data)
        dashboard_df.to_excel(writer, sheet_name='Dashboard', index=False)
    
    return output.getvalue()

# Generate PDF report using reportlab
def generate_pdf_report(report_df, total_patients, avg_score, max_score, min_score, avg_time, high_score_percent, low_score_percent, avg_missing_score):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph("eSanjeevani Teleconsultation Report", styles['Title']))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Summary Statistics", styles['Heading2']))
    summary_data = [
        ['Metric', 'Value'],
        ['Total Patients', str(total_patients)],
        ['Average Completion Score', f"{avg_score:.2f}%"],
        ['Maximum Completion Score', f"{max_score:.2f}%"],
        ['Minimum Completion Score', f"{min_score:.2f}%"],
        ['Average MissingFieldScore', f"{avg_missing_score:.2f}%"],
        ['Average Consultation Time (MM:SS)', avg_time],
        ['Patients with Score >= 75%', f"{high_score_percent:.2f}%"],
        ['Patients with Score < 50%', f"{low_score_percent:.2f}%"]
    ]
    summary_table = Table(summary_data)
    summary_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ]))
    elements.append(summary_table)
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Patient Consultation Details", styles['Heading2']))
    table_data = [list(report_df.columns)]
    for _, row in report_df.iterrows():
        row_data = [
            str(row['PatientId'])[:20],
            str(row['ConsultationId'])[:20],
            str(row['CompletionScore (%)']),
            str(row['CompletionField']),
            str(row['MissingFieldScore']),
            str(row['MissingFields']),
            str(row['TimeTaken (MM:SS)']),
            str(row['Status'])[:20],
            str(row['Symptoms']),
            str(row['Diagnosis']),
            str(row['Advice'])
        ]
        table_data.append(row_data)
    
    table = Table(table_data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    elements.append(table)
    doc.build(elements)
    return output.getvalue()

# Convert DataFrame to CSV for download
def convert_df_to_csv(df):
    output = BytesIO()
    df.to_csv(output, index=False)
    return output.getvalue()

if __name__ == "__main__":
    main()