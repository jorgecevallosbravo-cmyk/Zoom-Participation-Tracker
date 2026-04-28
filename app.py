"""
Zoom Attendance Tracker - Streamlit Web App
A user-friendly tool to generate attendance reports from Zoom transcripts.

Author: Jorge Bienvenido Cevallos Bravo
"""

import streamlit as st
import pandas as pd
import re
from datetime import datetime
from collections import defaultdict
import unicodedata
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.enums import TA_CENTER, TA_LEFT
import tempfile
import os

# ============================================================================
# CONFIGURATION
# ============================================================================

TEACHER_NAME = "JORGE BIENVENIDO CEVALLOS BRAVO"
PARTICIPATION_THRESHOLD = 5  # Words needed to be marked present

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def normalize_name(name):
    """Normalize a name by removing accents, converting to uppercase, and removing extra spaces."""
    if not name:
        return ""
    
    # Remove accents
    nfkd_form = unicodedata.normalize('NFKD', name)
    name_no_accents = ''.join([c for c in nfkd_form if not unicodedata.combining(c)])
    
    # Convert to uppercase and remove extra spaces
    normalized = ' '.join(name_no_accents.upper().split())
    
    return normalized


def fuzzy_match_name(transcript_name, student_name):
    """Check if a transcript name matches a student name using fuzzy matching."""
    transcript_normalized = normalize_name(transcript_name)
    student_normalized = normalize_name(student_name)
    
    # Split into words
    transcript_words = transcript_normalized.split()
    
    # Check if any two consecutive words from transcript appear in student name
    if len(transcript_words) >= 2:
        for i in range(len(transcript_words) - 1):
            bigram = f"{transcript_words[i]} {transcript_words[i+1]}"
            if bigram in student_normalized:
                return True
    
    # Also check if single word matches (for very short names)
    if len(transcript_words) == 1:
        return transcript_words[0] in student_normalized
    
    return False


def count_words(text):
    """Count the number of words in a text string."""
    words = re.findall(r'\b\w+\b', text)
    return len(words)


def parse_transcript(transcript_content):
    """Parse the Zoom transcript and count words per speaker."""
    word_counts = defaultdict(int)
    
    lines = transcript_content.decode('utf-8').split('\n')
    
    current_speaker = None
    
    for line in lines:
        line = line.strip()
        
        # Check if this is a speaker line (format: [Speaker Name] timestamp)
        speaker_match = re.match(r'\[(.*?)\]\s+\d{2}:\d{2}:\d{2}', line)
        
        if speaker_match:
            current_speaker = speaker_match.group(1)
        elif current_speaker and line:
            # This is speech content
            # Skip if it's the teacher
            if not fuzzy_match_name(current_speaker, TEACHER_NAME):
                word_counts[current_speaker] += count_words(line)
    
    return dict(word_counts)


def match_students_to_transcript(student_df, word_counts, student_name_column):
    """Match students from the Excel list to their word counts in the transcript."""
    results = []
    
    for _, student_row in student_df.iterrows():
        student_name = student_row[student_name_column]
        matched_count = 0
        
        # Try to match this student with transcript entries
        for transcript_speaker, count in word_counts.items():
            if fuzzy_match_name(transcript_speaker, student_name):
                matched_count += count
        
        # Determine status based on threshold
        if matched_count >= PARTICIPATION_THRESHOLD:
            status = "present"
        else:
            status = "absent"
        
        results.append((student_name, matched_count, status))
    
    return results


def create_pdf_report(attendance_data, output_path, date_str, course_code):
    """Create a professional PDF attendance report."""
    doc = SimpleDocTemplate(
        output_path,
        pagesize=letter,
        rightMargin=0.75*inch,
        leftMargin=0.75*inch,
        topMargin=1*inch,
        bottomMargin=0.75*inch
    )
    
    # Define custom styles
    styles = getSampleStyleSheet()
    
    # Title style - Blue
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1e3a8a'),
        spaceAfter=12,
        alignment=TA_CENTER,
        fontName='Helvetica-Bold'
    )
    
    # Header info style - Grey
    header_style = ParagraphStyle(
        'HeaderInfo',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#4b5563'),
        spaceAfter=6,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )
    
    # Build the document
    story = []
    
    # Title
    title = Paragraph("ATTENDANCE REPORT", title_style)
    story.append(title)
    story.append(Spacer(1, 0.2*inch))
    
    # Header information
    header_info = [
        f"<b>Course:</b> {course_code}",
        f"<b>Date:</b> {date_str}",
        f"<b>Instructor:</b> {TEACHER_NAME}"
    ]
    
    for info in header_info:
        story.append(Paragraph(info, header_style))
    
    story.append(Spacer(1, 0.3*inch))
    
    # Summary section
    present_count = sum(1 for _, _, status in attendance_data if status == "present")
    total_students = len(attendance_data)
    absent_count = total_students - present_count
    
    summary_style = ParagraphStyle(
        'Summary',
        parent=styles['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#1f2937'),
        spaceAfter=4,
        alignment=TA_LEFT,
        fontName='Helvetica'
    )
    
    summary_data = [
        f"<b>Total Students:</b> {total_students}",
        f"<b>Present:</b> {present_count}",
        f"<b>Absent:</b> {absent_count}"
    ]
    
    for summary in summary_data:
        story.append(Paragraph(summary, summary_style))
    
    story.append(Spacer(1, 0.3*inch))
    
    # Table styles
    normal_style = ParagraphStyle(
        'NormalText',
        parent=styles['Normal'],
        fontSize=9,
        leading=11,
        fontName='Helvetica',
        alignment=TA_LEFT
    )
    
    header_text_style = ParagraphStyle(
        'HeaderText',
        parent=styles['Normal'],
        fontSize=10,
        leading=12,
        fontName='Helvetica-Bold',
        textColor=colors.whitesmoke,
        alignment=TA_CENTER
    )
    
    center_style = ParagraphStyle(
        'CenterText',
        parent=styles['Normal'],
        fontSize=14,
        alignment=TA_CENTER,
        fontName='Helvetica'
    )
    
    table_data = [
        [
            Paragraph('<b>APELLIDOS Y NOMBRES</b>', header_text_style),
            Paragraph('<b>PARTICIPATION<br/>&amp;<br/>ATTENDANCE</b>', header_text_style),
            Paragraph('<b>PARTICIPATION DETAILS</b>', header_text_style)
        ]
    ]
    
    for student_name, word_count, status in sorted(attendance_data, key=lambda x: x[0]):
        # Status symbol with color
        if status == "present":
            status_symbol = '<font color="green">✓</font>'
        else:
            status_symbol = '<font color="red">✗</font>'
        
        # Participation message
        if word_count == 0:
            participation_msg = "The student had no participation and has been marked as not present."
        else:
            participation_msg = f"The student participated with a total of {word_count} words."
        
        # Wrap text in Paragraph objects
        name_para = Paragraph(student_name, normal_style)
        status_para = Paragraph(status_symbol, center_style)
        details_para = Paragraph(participation_msg, normal_style)
        
        table_data.append([
            name_para,
            status_para,
            details_para
        ])
    
    # Create table
    table = Table(table_data, colWidths=[2.2*inch, 1.8*inch, 3.4*inch])
    
    # Table style
    table.setStyle(TableStyle([
        # Header row
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a8a')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('TOPPADDING', (0, 0), (-1, 0), 10),
        
        # Data rows
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('ALIGN', (0, 1), (0, -1), 'LEFT'),
        ('ALIGN', (1, 1), (1, -1), 'CENTER'),
        ('ALIGN', (2, 1), (2, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        
        # Alternating row colors
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f3f4f6')]),
        
        # Grid
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#9ca3af')),
        
        # Padding
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
    ]))
    
    story.append(table)
    
    # Footer
    story.append(Spacer(1, 0.3*inch))
    footer_style = ParagraphStyle(
        'Footer',
        parent=styles['Normal'],
        fontSize=8,
        textColor=colors.HexColor('#6b7280'),
        alignment=TA_CENTER,
        fontName='Helvetica-Oblique'
    )
    footer = Paragraph(f"Report generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", footer_style)
    story.append(footer)
    
    # Build PDF
    doc.build(story)


# ============================================================================
# STREAMLIT APP
# ============================================================================

def main():
    st.set_page_config(
        page_title="Zoom Attendance Tracker",
        page_icon="📊",
        layout="centered"
    )
    
    # Header
    st.title("📊 Zoom Attendance Tracker")
    st.markdown("Generate professional attendance reports from Zoom transcripts")
    st.markdown("---")
    
    # Instructions
    with st.expander("📖 How to use this tool"):
        st.markdown("""
        1. **Upload your Zoom transcript** (`.txt` file)
        2. **Upload your student list** (`.xlsx` file with student names)
        3. **Click 'Generate Report'** to create the PDF
        4. **Download** your attendance report
        
        **Note:** The Excel file name will be used as the course code in the report.
        """)
    
    # File uploaders
    st.subheader("Step 1: Upload Files")
    
    col1, col2 = st.columns(2)
    
    with col1:
        transcript_file = st.file_uploader(
            "📄 Zoom Transcript (.txt)",
            type=['txt'],
            help="Upload the Zoom meeting transcript file"
        )
    
    with col2:
        student_file = st.file_uploader(
            "📋 Student List (.xlsx)",
            type=['xlsx'],
            help="Upload the Excel file with student names"
        )
    
    # Process button
    if transcript_file and student_file:
        st.markdown("---")
        st.subheader("Step 2: Generate Report")
        
        if st.button("🚀 Generate Attendance Report", type="primary", use_container_width=True):
            try:
                with st.spinner("Processing attendance data..."):
                    # Extract course code from Excel filename
                    course_code = os.path.splitext(student_file.name)[0]
                    
                    # Parse transcript
                    st.info("📝 Analyzing Zoom transcript...")
                    word_counts = parse_transcript(transcript_file.read())
                    transcript_file.seek(0)  # Reset file pointer
                    
                    # Load student list
                    st.info("👥 Loading student list...")
                    student_df = None
                    student_name_column = None
                    
                    # Try different header row positions
                    for header_row in [2, 1, 0]:
                        try:
                            temp_df = pd.read_excel(student_file, header=header_row)
                            temp_df = temp_df.dropna(how='all')
                            
                            # Find ALUMNO column
                            for col in temp_df.columns:
                                col_str = str(col).strip().upper()
                                if 'ALUMNO' in col_str:
                                    student_df = temp_df
                                    student_name_column = col
                                    break
                            
                            if student_name_column:
                                break
                        except:
                            continue
                    
                    if student_df is None or student_name_column is None:
                        st.error("❌ Could not find student name column in Excel file. Please ensure it contains 'ALUMNO' in the header.")
                        return
                    
                    # Match students to transcript
                    st.info("🔍 Matching students to participation data...")
                    attendance_data = match_students_to_transcript(student_df, word_counts, student_name_column)
                    
                    # Count statistics
                    present_count = sum(1 for _, _, status in attendance_data if status == "present")
                    absent_count = len(attendance_data) - present_count
                    
                    # Display summary
                    st.success("✅ Processing complete!")
                    
                    col1, col2, col3 = st.columns(3)
                    col1.metric("Total Students", len(attendance_data))
                    col2.metric("Present", present_count, delta=None)
                    col3.metric("Absent", absent_count, delta=None)
                    
                    # Generate PDF
                    st.info("📄 Generating PDF report...")
                    
                    # Create temporary file for PDF
                    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                        pdf_path = tmp_file.name
                        date_str = datetime.now().strftime("%B %d, %Y")
                        create_pdf_report(attendance_data, pdf_path, date_str, course_code)
                    
                    # Read PDF for download
                    with open(pdf_path, 'rb') as f:
                        pdf_data = f.read()
                    
                    # Clean up temp file
                    os.unlink(pdf_path)
                    
                    # Download button
                    st.markdown("---")
                    st.subheader("Step 3: Download Report")
                    
                    filename = f"Attendance_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                    
                    st.download_button(
                        label="⬇️ Download PDF Report",
                        data=pdf_data,
                        file_name=filename,
                        mime="application/pdf",
                        use_container_width=True,
                        type="primary"
                    )
                    
                    st.success(f"🎉 Report generated successfully!")
                    
            except Exception as e:
                st.error(f"❌ An error occurred: {str(e)}")
                st.error("Please check your files and try again.")
    
    else:
        st.info("👆 Please upload both files to continue")
    
    # Footer
    st.markdown("---")
    st.markdown(
        "<p style='text-align: center; color: #666; font-size: 0.9em;'>"
        "Zoom Attendance Tracker | Built for educators"
        "</p>",
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()
