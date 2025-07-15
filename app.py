import streamlit as st
import pandas as pd
import base64
from io import BytesIO
from datetime import timedelta
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

# Function to parse HH_MM_SS column to seconds for average calculation
def parse_hh_mm_ss(time_str):
    try:
        if pd.isna(time_str) or time_str == '':
            return None
        # Expecting format MM:SS or HH:MM:SS
        parts = time_str.split(':')
        if len(parts) == 2:
            minutes, seconds = map(int, parts)
            hours = 0
        elif len(parts) == 3:
            hours, minutes, seconds = map(int, parts)
        else:
            return None
        return hours * 3600 + minutes * 60 + seconds
    except:
        return None

# Function to get time taken from HH_MM_SS column
def get_time_taken(row):
    try:
        time_str = row.get('HH_MM_SS', 'Unknown')
        if pd.isna(time_str) or time_str == '':
            return "Unknown"
        return time_str
    except:
        return "Unknown"

# Function to calculate consultation completion score and missing fields
def calculate_completion_score(row):
    score = 0
    total_fields = 0
    missing_fields = []

    # Define fields to check for completeness
    fields_to_check = [
        'PatientName', 'Age', 'GenderDisplay', 'ABHANumber', 'IsFollowUp',
        'SentByLocationName', 'SentByName', 'SentToLocationName', 'SentToName',
        'SentToSpecialityDisplay', 'ConsultationCreatedDate', 'ConsultationStatus',
        'StartDate', 'CloseDate', 'Snomed FamilyHistory', 'Snomed MedicalHistory',
        'Snomed PersonalHistory', 'Additional MedicalHistory', 'Snomed Allergy',
        'AdditionalAllergy', 'Snomed Active Medicine', 'Diagnostics',
        'AdditionalDiagnostics', 'Query', 'Additional Medicine',
        'Provisional Diagnosis', 'Additional Diagnosis', 'Snomed Medicine',
        'Advice', 'Symptoms_', 'DifferentialDiagnosis_'
    ]

    for field in fields_to_check:
        total_fields += 1
        # Check if the field exists and is not empty or NaN
        if field in row and pd.notna(row[field]) and row[field] != '':
            score += 1
        else:
            missing_fields.append(field)

    # Calculate percentage
    completion_percentage = (score / total_fields) * 100 if total_fields > 0 else 0
    return round(completion_percentage, 2), missing_fields

# Calculate average consultation time from HH_MM_SS
def calculate_average_consultation_time(df):
    total_seconds = 0
    valid_count = 0
    for _, row in df.iterrows():
        time_seconds = parse_hh_mm_ss(row.get('HH_MM_SS', ''))
        if time_seconds is not None:
            total_seconds += time_seconds
            valid_count += 1
    if valid_count > 0:
        avg_seconds = total_seconds / valid_count
        minutes = int(avg_seconds // 60)
        seconds = int(avg_seconds % 60)
        return f"{minutes:02d}:{seconds:02d}"
    return "Unknown"

# Process the data and generate report
def generate_consultation_report(df):
    if df is None or df.empty:
        return None

    # Initialize report list
    report = []

    # Process each patient
    for index, row in df.iterrows():
        patient_id = row.get('PatientId', 'Unknown')
        consultation_id = row.get('ConsultationId', 'Unknown')
        completion_score, missing_fields = calculate_completion_score(row)
        time_taken = get_time_taken(row)
        
        patient_report = {
            'PatientId': patient_id,
            'ConsultationId': consultation_id,
            'CompletionScore (%)': completion_score,
            'MissingFields': ', '.join(missing_fields) if missing_fields else 'None',
            'TimeTaken (MM:SS)': time_taken,
            'Status': row.get('ConsultationStatus', 'Unknown'),
            'Symptoms': row.get('Symptoms_', 'No symptoms recorded'),
            'Diagnosis': row.get('Provisional Diagnosis', 'No diagnosis recorded'),
            'Advice': row.get('Advice', 'No advice recorded')
        }
        report.append(patient_report)

    # Convert report to DataFrame
    report_df = pd.DataFrame(report)
    return report_df

# Generate Excel report with dashboard sheet
def generate_excel_report(report_df, total_patients, avg_score, max_score, min_score, avg_time):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write patient report to first sheet
        report_df.to_excel(writer, sheet_name='Patient Report', index=False)
        
        # Create dashboard sheet
        dashboard_data = {
            'Metric': [
                'Total Patients',
                'Average Completion Score',
                'Maximum Completion Score',
                'Minimum Completion Score',
                'Average Consultation Time (MM:SS)'
            ],
            'Value': [
                total_patients,
                f"{avg_score:.2f}%",
                f"{max_score:.2f}%",
                f"{min_score:.2f}%",
                avg_time
            ]
        }
        dashboard_df = pd.DataFrame(dashboard_data)
        dashboard_df.to_excel(writer, sheet_name='Dashboard', index=False)
    
    return output.getvalue()

# Generate PDF report using reportlab
def generate_pdf_report(report_df, total_patients, avg_score, max_score, min_score, avg_time):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    # Add title
    elements.append(Paragraph("eSanjeevani Teleconsultation Report", styles['Title']))
    elements.append(Spacer(1, 12))

    # Add summary statistics
    elements.append(Paragraph("Summary Statistics", styles['Heading2']))
    summary_data = [
        ['Metric', 'Value'],
        ['Total Patients', str(total_patients)],
        ['Average Completion Score', f"{avg_score:.2f}%"],
        ['Maximum Completion Score', f"{max_score:.2f}%"],
        ['Minimum Completion Score', f"{min_score:.2f}%"],
        ['Average Consultation Time (MM:SS)', avg_time]
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

    # Add patient consultation details
    elements.append(Paragraph("Patient Consultation Details", styles['Heading2']))
    # Prepare table data
    table_data = [list(report_df.columns)]
    for _, row in report_df.iterrows():
        # Truncate long fields to prevent overflow
        row_data = [
            str(row['PatientId'])[:20],
            str(row['ConsultationId'])[:20],
            str(row['CompletionScore (%)']),
            str(row['MissingFields'])[:50] + ('...' if len(str(row['MissingFields'])) > 50 else ''),
            str(row['TimeTaken (MM:SS)']),
            str(row['Status'])[:20],
            str(row['Symptoms'])[:50] + ('...' if len(str(row['Symptoms'])) > 50 else ''),
            str(row['Diagnosis'])[:50] + ('...' if len(str(row['Diagnosis'])) > 50 else ''),
            str(row['Advice'])[:50] + ('...' if len(str(row['Advice'])) > 50 else '')
        ]
        table_data.append(row_data)
    
    # Create table
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

    # Build PDF
    doc.build(elements)
    return output.getvalue()

# Convert DataFrame to CSV for download
def convert_df_to_csv(df):
    output = BytesIO()
    df.to_csv(output, index=False)
    return output.getvalue()

# Streamlit app
def main():
    st.set_page_config(page_title="eSanjeevani Consultation Analysis", layout="wide")
    st.title("eSanjeevani Teleconsultation Completion Report")
    st.write("Upload an Excel file containing consultation details to generate a completion score report for each patient, including time taken and missing fields.")

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

    if uploaded_file is not None:
        try:
            # Load Excel file
            df = pd.read_excel(uploaded_file)
            
            # Generate report
            report_df = generate_consultation_report(df)
            
            if report_df is not None and not report_df.empty:
                # Calculate summary statistics
                total_patients = len(report_df)
                avg_score = report_df['CompletionScore (%)'].mean()
                max_score = report_df['CompletionScore (%)'].max()
                min_score = report_df['CompletionScore (%)'].min()
                avg_time = calculate_average_consultation_time(df)
                
                # Display summary statistics
                st.subheader("Report Summary")
                st.write(f"**Total Patients:** {total_patients}")
                st.write(f"**Average Completion Score:** {avg_score:.2f}%")
                st.write(f"**Maximum Completion Score:** {max_score:.2f}%")
                st.write(f"**Minimum Completion Score:** {min_score:.2f}%")
                st.write(f"**Average Consultation Time (MM:SS):** {avg_time}")

                # Display the report
                st.subheader("Patient Consultation Report")
                st.dataframe(report_df, use_container_width=True)

                # Generate and provide download buttons
                # CSV Download
                csv = convert_df_to_csv(report_df)
                st.download_button(
                    label="Download Report as CSV",
                    data=csv,
                    file_name="consultation_completion_report.csv",
                    mime="text/csv",
                )

                # Excel Download
                excel_data = generate_excel_report(report_df, total_patients, avg_score, max_score, min_score, avg_time)
                st.download_button(
                    label="Download Report as Excel",
                    data=excel_data,
                    file_name="consultation_completion_report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                # PDF Download
                pdf_data = generate_pdf_report(report_df, total_patients, avg_score, max_score, min_score, avg_time)
                st.download_button(
                    label="Download Report as PDF",
                    data=pdf_data,
                    file_name="consultation_completion_report.pdf",
                    mime="application/pdf",
                )

            else:
                st.error("No data to display. The file may be empty or incorrectly formatted.")
        
        except Exception as e:
            st.error(f"Error processing the file: {e}")
    else:
        st.info("Please upload an Excel file to proceed.")

if __name__ == "__main__":
    main()